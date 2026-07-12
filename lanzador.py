import sys
import os
import json
import time
import hashlib
import shutil
import subprocess
import tempfile
import zipfile
import requests
from PySide6.QtCore import Qt, QThread, Signal
from PySide6.QtWidgets import QApplication, QDialog, QVBoxLayout, QLabel, QProgressBar, QMessageBox

UPDATE_JSON_URL = "https://gist.githubusercontent.com/jhoelproo/d05cfbcf776796f0aafe0bbbf8947eba/raw/"
MAIN_APP_NAME = "CALCULOS_QT.exe"
UPDATER_NAME = "APLICAR_ACTUALIZACION.exe"
CONFIG_FILE = "version_config.json"
LOG_FILE = "lanzador_log.txt"
DEFAULT_VERSION = "2.1.0"

# ---> FUNCIÓN MAESTRA PARA RUTAS DE PYINSTALLER <---
def get_real_dir():
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def get_bundle_dir():
    """Carpeta de recursos de PyInstaller (``_internal`` en modo onedir)."""
    return getattr(sys, "_MEIPASS", get_real_dir())


def write_launcher_log(message: str):
    """Guarda diagnósticos sin detener el inicio con ventanas de OK."""
    try:
        log_path = os.path.join(get_real_dir(), LOG_FILE)
        with open(log_path, "a", encoding="utf-8") as f:
            f.write(f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] {message}\n")
    except Exception:
        pass


def version_tuple(version: str):
    """Convierte '1.1.7' en (1, 1, 7) para comparar correctamente versiones."""
    try:
        parts = tuple(int(part) for part in str(version).strip().split("."))
        return parts + (0,) * max(0, 3 - len(parts))
    except Exception:
        return (0, 0, 0)


def is_remote_newer(remote_version: str, local_version: str) -> bool:
    return version_tuple(remote_version) > version_tuple(local_version)


def get_local_version():
    config_paths = [
        os.path.join(get_real_dir(), CONFIG_FILE),
        os.path.join(get_bundle_dir(), CONFIG_FILE),
    ]
    for config_path in dict.fromkeys(config_paths):
        if not os.path.exists(config_path):
            continue
        try:
            with open(config_path, 'r', encoding="utf-8") as f:
                return json.load(f).get("version", DEFAULT_VERSION)
        except Exception as e:
            write_launcher_log(f"No se pudo leer version_config.json: {e}")
    return DEFAULT_VERSION


def save_local_version(version):
    config_path = os.path.join(get_real_dir(), CONFIG_FILE)
    with open(config_path, 'w', encoding="utf-8") as f:
        json.dump({"version": version}, f)


class UpdateChecker(QThread):
    update_found = Signal(str, str, str)
    no_update = Signal(str)
    error_found = Signal(str)

    def run(self):
        try:
            # Anti-caché: obliga a GitHub/Gist a devolver la versión más reciente.
            no_cache_url = f"{UPDATE_JSON_URL}?t={int(time.time())}"
            response = requests.get(no_cache_url, timeout=10)

            if response.status_code == 200:
                try:
                    data = response.json()
                    remote_version = data.get("version")
                    download_url = data.get("onedir_url") or data.get("url")
                    expected_sha256 = str(
                        data.get("onedir_sha256") or data.get("sha256") or ""
                    ).strip().lower()
                    local_version = get_local_version()

                    if remote_version and download_url and is_remote_newer(remote_version, local_version):
                        self.update_found.emit(remote_version, download_url, expected_sha256)
                    else:
                        self.no_update.emit(
                            f"Versión Local leída: {local_version} | "
                            f"Versión Nube leída: {remote_version} | Sin actualización pendiente."
                        )
                except json.JSONDecodeError:
                    self.error_found.emit(
                        "El texto en Gist no es un JSON válido. "
                        f"Texto recibido: {response.text[:150]}"
                    )
            else:
                self.error_found.emit(f"Fallo de conexión con Gist. Código: {response.status_code}")
        except Exception as e:
            self.error_found.emit(f"Error general de red: {str(e)}")


class DownloadWorker(QThread):
    finished_download = Signal(str)
    progress_update = Signal(int, int)
    error_download = Signal(str)

    def __init__(self, url, expected_sha256=""):
        super().__init__()
        self.url = url
        self.expected_sha256 = expected_sha256

    def run(self):
        new_zip_path = ""
        try:
            response = requests.get(self.url, stream=True, timeout=60)
            response.raise_for_status()
            total_size = int(response.headers.get('content-length', 0))

            update_dir = tempfile.mkdtemp(prefix="hospital_update_")
            new_zip_path = os.path.join(update_dir, "HOSPITAL-update.zip")

            downloaded_size = 0
            digest = hashlib.sha256()
            with open(new_zip_path, "wb") as f:
                for chunk in response.iter_content(32768):
                    if chunk:
                        f.write(chunk)
                        digest.update(chunk)
                        downloaded_size += len(chunk)
                        self.progress_update.emit(downloaded_size, total_size)

            if total_size > 0 and downloaded_size != total_size:
                raise IOError(
                    f"Descarga incompleta: {downloaded_size} de {total_size} bytes."
                )

            with open(new_zip_path, "rb") as executable_file:
                if executable_file.read(4) != b"PK\x03\x04":
                    raise IOError("El archivo descargado no es un ejecutable válido de Windows.")

            if self.expected_sha256 and digest.hexdigest().lower() != self.expected_sha256:
                raise IOError("La firma SHA-256 del paquete no coincide.")
            if not zipfile.is_zipfile(new_zip_path):
                raise IOError("El archivo descargado no es un ZIP valido.")
            self.finished_download.emit(new_zip_path)
        except Exception as e:
            if new_zip_path and os.path.exists(new_zip_path):
                try:
                    shutil.rmtree(os.path.dirname(new_zip_path), ignore_errors=True)
                except Exception:
                    pass
            self.error_download.emit(str(e))
            self.finished_download.emit("")


class LauncherDialog(QDialog):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Iniciando Sistema...")
        self.setFixedSize(450, 150)
        self.setWindowFlags(Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint)
        self.setStyleSheet("background-color: #ffffff; border: 2px solid #1565c0; border-radius: 8px;")

        lay = QVBoxLayout(self)
        self.lbl_status = QLabel("Comprobando actualización...")
        self.lbl_status.setAlignment(Qt.AlignCenter)
        self.lbl_status.setStyleSheet("font-size: 12pt; font-weight: bold; color: #1565c0; border: none;")
        lay.addWidget(self.lbl_status)

        self.progress = QProgressBar()
        self.progress.setStyleSheet("""
            QProgressBar { border: 1px solid #d1d9e6; border-radius: 4px; text-align: center; font-weight: bold; }
            QProgressBar::chunk { background-color: #2e7d32; border-radius: 4px; }
        """)
        self.progress.hide()
        lay.addWidget(self.progress)

        self.new_version_string = ""

        self.checker = UpdateChecker()
        self.checker.update_found.connect(self.start_download)
        self.checker.no_update.connect(self.handle_no_update)
        self.checker.error_found.connect(self.handle_update_check_error)
        self.checker.start()

    def handle_no_update(self, msg):
        # Ya NO muestra QMessageBox ni requiere presionar OK.
        write_launcher_log(msg)
        self.lbl_status.setText("Sin actualización pendiente. Iniciando sistema...")
        QApplication.processEvents()
        self.launch_main_app()

    def handle_update_check_error(self, err):
        # Si falla la verificación, abre la app local sin bloquear al usuario.
        write_launcher_log(f"Error verificando actualización: {err}")
        self.lbl_status.setText("No se pudo verificar actualización. Iniciando sistema...")
        QApplication.processEvents()
        self.launch_main_app()

    def start_download(self, version, url, expected_sha256):
        self.new_version_string = version
        self.lbl_status.setText(f"Descargando actualización {version}...")
        self.progress.show()

        self.downloader = DownloadWorker(url, expected_sha256)
        self.downloader.progress_update.connect(self.update_progress)
        self.downloader.finished_download.connect(self.apply_update)
        self.downloader.error_download.connect(self.handle_download_error)
        self.downloader.start()

    def handle_download_error(self, err):
        # No bloquea con OK. Deja registro y continúa con la versión local.
        write_launcher_log(f"Error de descarga: {err}")
        self.lbl_status.setText("No se pudo descargar actualización. Iniciando versión local...")
        QApplication.processEvents()

    def update_progress(self, downloaded, total):
        if total > 0:
            self.progress.setMaximum(total)
            self.progress.setValue(downloaded)
            mb_down = downloaded / (1024 * 1024)
            mb_total = total / (1024 * 1024)
            self.progress.setFormat(f"{mb_down:.1f} MB / {mb_total:.1f} MB")
        else:
            self.progress.setMaximum(0)
            self.progress.setFormat(f"{(downloaded / (1024 * 1024)):.1f} MB")

    def _apply_legacy_update(self, new_exe_path):
        if not new_exe_path or not os.path.exists(new_exe_path):
            self.launch_main_app()
            return

        self.lbl_status.setText("Instalando actualización...")
        QApplication.processEvents()
        time.sleep(1.0)

        current_dir = get_real_dir()
        main_app_path = os.path.join(current_dir, MAIN_APP_NAME)
        old_app_path = main_app_path + f".{int(time.time())}.old"
        backup_created = False

        try:
            if os.path.exists(main_app_path):
                os.replace(main_app_path, old_app_path)
                backup_created = True
            os.replace(new_exe_path, main_app_path)
            try:
                save_local_version(self.new_version_string)
            except Exception as config_error:
                write_launcher_log(
                    f"La actualización se instaló, pero no se guardó su versión: {config_error}"
                )
            write_launcher_log(f"Actualización instalada: {self.new_version_string}")
            self.lbl_status.setText(f"Actualización {self.new_version_string} instalada. Iniciando...")
            QApplication.processEvents()
        except Exception as e:
            # Si falla el segundo reemplazo, restaura la versión funcional.
            if backup_created and not os.path.exists(main_app_path) and os.path.exists(old_app_path):
                try:
                    os.replace(old_app_path, main_app_path)
                    write_launcher_log("Se restauró la versión anterior tras fallar la actualización.")
                except Exception as restore_error:
                    write_launcher_log(f"No se pudo restaurar la versión anterior: {restore_error}")
            write_launcher_log(f"No se pudo reemplazar el ejecutable: {e}")
            self.lbl_status.setText("No se pudo instalar actualización. Iniciando versión local...")
            QApplication.processEvents()

        self.launch_main_app()

    def apply_update(self, new_zip_path):
        if not new_zip_path or not os.path.exists(new_zip_path):
            self.launch_main_app()
            return

        self.lbl_status.setText("Preparando actualizacion segura...")
        QApplication.processEvents()
        current_dir = os.path.abspath(get_real_dir())
        update_root = os.path.dirname(new_zip_path)
        extract_dir = os.path.join(update_root, "extraido")

        try:
            with zipfile.ZipFile(new_zip_path) as archive:
                extract_abs = os.path.abspath(extract_dir)
                for member in archive.infolist():
                    target = os.path.abspath(os.path.join(extract_abs, member.filename))
                    if os.path.commonpath([extract_abs, target]) != extract_abs:
                        raise IOError("El ZIP contiene una ruta no permitida.")
                archive.extractall(extract_abs)

            candidates = [os.path.join(extract_dir, "HOSPITAL"), extract_dir]
            payload_dir = next(
                (
                    path
                    for path in candidates
                    if os.path.isfile(os.path.join(path, MAIN_APP_NAME))
                    and os.path.isfile(os.path.join(path, "INICIAR_SISTEMA.exe"))
                    and os.path.isdir(os.path.join(path, "_internal"))
                ),
                "",
            )
            if not payload_dir:
                raise IOError("El ZIP no contiene una distribucion HOSPITAL onedir valida.")

            installed_updater = os.path.join(current_dir, UPDATER_NAME)
            if not os.path.isfile(installed_updater):
                raise FileNotFoundError(f"No se encontro {UPDATER_NAME}.")
            temporary_updater = os.path.join(update_root, UPDATER_NAME)
            shutil.copy2(installed_updater, temporary_updater)

            manifest_path = os.path.join(update_root, "update_manifest.json")
            with open(manifest_path, "w", encoding="utf-8") as manifest_file:
                json.dump(
                    {
                        "install_dir": current_dir,
                        "payload_dir": os.path.abspath(payload_dir),
                        "version": self.new_version_string,
                    },
                    manifest_file,
                )

            creationflags = 0
            if sys.platform.startswith("win"):
                creationflags = subprocess.CREATE_NEW_PROCESS_GROUP | subprocess.DETACHED_PROCESS
            subprocess.Popen(
                [temporary_updater, "--manifest", manifest_path, "--wait-pid", str(os.getpid())],
                cwd=update_root,
                close_fds=True,
                creationflags=creationflags,
            )
            write_launcher_log(f"Actualizacion {self.new_version_string} entregada al instalador externo.")
            self.lbl_status.setText("Cerrando para completar la actualizacion...")
            QApplication.processEvents()
            QApplication.instance().quit()
        except Exception as exc:
            write_launcher_log(f"No se pudo preparar la actualizacion onedir: {exc}")
            self.lbl_status.setText("No se pudo instalar. Iniciando version local...")
            QApplication.processEvents()
            self.launch_main_app()

    def launch_main_app(self):
        current_dir = get_real_dir()
        main_app_path = os.path.join(current_dir, MAIN_APP_NAME)

        if os.path.exists(main_app_path):
            try:
                if sys.platform.startswith("win"):
                    os.startfile(main_app_path)
                else:
                    import subprocess
                    subprocess.Popen([main_app_path], cwd=current_dir)
            except Exception as e:
                QMessageBox.critical(self, "Error Fatal", f"No se pudo iniciar el programa:\n{e}")
        else:
            QMessageBox.critical(
                self,
                "Archivo no encontrado",
                f"No se encontró el sistema principal.\n\nAsegúrate de tener '{MAIN_APP_NAME}' en la carpeta."
            )
        sys.exit(0)


if __name__ == "__main__":
    if sys.argv[1:] == ["--self-test"]:
        raise SystemExit(0)
    app = QApplication(sys.argv)
    launcher = LauncherDialog()
    launcher.show()
    sys.exit(app.exec())
