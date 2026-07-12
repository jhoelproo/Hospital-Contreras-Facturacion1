# -*- mode: python ; coding: utf-8 -*-

import json
import os
from importlib.util import find_spec
from pathlib import Path

from PyInstaller.utils.hooks import collect_data_files, collect_submodules


ROOT = Path(SPECPATH).resolve()
ASSETS = ROOT / "assets"
PDF_ENGINE = ROOT / "pdf_engine"
REPORT_ENGINE = ROOT / "report_engine"

REQUIRED_FILES = [
    ROOT / "CALCULOS_QT.py",
    ROOT / "lanzador.py",
    ROOT / "updater.py",
    ROOT / "migrador_onedir.py",
    ROOT / "config_local.py",
    ASSETS / "logo.jpg",
    ASSETS / "favicon.ico",
    PDF_ENGINE / "__init__.py",
    PDF_ENGINE / "renderer.py",
    PDF_ENGINE / "template.html",
    PDF_ENGINE / "styles.css",
    REPORT_ENGINE / "__init__.py",
    REPORT_ENGINE / "data_service.py",
    REPORT_ENGINE / "excel_exporter.py",
    REPORT_ENGINE / "html_renderer.py",
    REPORT_ENGINE / "report_template.html",
    REPORT_ENGINE / "report_styles.css",
    ROOT / "version_config.json",
]
missing_files = [str(path) for path in REQUIRED_FILES if not path.is_file()]
if missing_files:
    raise FileNotFoundError(
        "Faltan archivos requeridos para compilar:\n" + "\n".join(missing_files)
    )


def find_playwright_browser():
    """Localiza el Chromium Headless Shell de la versión instalada."""
    playwright_spec = find_spec("playwright")
    if playwright_spec is None or not playwright_spec.submodule_search_locations:
        raise RuntimeError("Playwright no está instalado en el Python usado para compilar.")

    playwright_dir = Path(next(iter(playwright_spec.submodule_search_locations))).resolve()
    manifest_path = playwright_dir / "driver" / "package" / "browsers.json"
    manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    browser_info = next(
        browser
        for browser in manifest["browsers"]
        if browser["name"] == "chromium-headless-shell"
    )
    revision = str(browser_info["revision"])
    folder_names = (
        f"chromium_headless_shell-{revision}",
        f"chromium-headless-shell-{revision}",
    )

    browser_roots = []
    configured_root = os.environ.get("PLAYWRIGHT_BROWSERS_PATH")
    if configured_root == "0":
        browser_roots.append(playwright_dir / "driver" / "package" / ".local-browsers")
    elif configured_root:
        browser_roots.append(Path(configured_root).expanduser())

    local_app_data = os.environ.get("LOCALAPPDATA")
    if local_app_data:
        browser_roots.append(Path(local_app_data) / "ms-playwright")
    browser_roots.append(playwright_dir / "driver" / "package" / ".local-browsers")

    browser_dir = next(
        (
            root / folder_name
            for root in browser_roots
            for folder_name in folder_names
            if (root / folder_name).is_dir()
        ),
        None,
    )
    if browser_dir is None:
        raise FileNotFoundError(
            "No se encontró Chromium Headless Shell para Playwright "
            f"(revisión {revision}). Instálalo con el mismo Python del build: "
            "python -m playwright install chromium"
        )
    return browser_dir


PLAYWRIGHT_BROWSER = find_playwright_browser()
PLAYWRIGHT_DATAS = collect_data_files("playwright")
PLAYWRIGHT_HIDDEN_IMPORTS = collect_submodules("playwright")
OPENPYXL_DATAS = collect_data_files("openpyxl")
OPENPYXL_HIDDEN_IMPORTS = collect_submodules("openpyxl")

main_datas = [
    (str(ASSETS / "logo.jpg"), "assets"),
    (str(ASSETS / "favicon.ico"), "assets"),
    (str(PDF_ENGINE / "template.html"), "pdf_engine"),
    (str(PDF_ENGINE / "styles.css"), "pdf_engine"),
    (str(REPORT_ENGINE / "report_template.html"), "report_engine"),
    (str(REPORT_ENGINE / "report_styles.css"), "report_engine"),
    (str(ROOT / "version_config.json"), "."),
    (
        str(PLAYWRIGHT_BROWSER),
        f"playwright-browsers/{PLAYWRIGHT_BROWSER.name}",
    ),
] + PLAYWRIGHT_DATAS + OPENPYXL_DATAS

main_hidden_imports = sorted(
    set(
        [
            "pdf_engine",
            "pdf_engine.renderer",
            "psycopg2",
            "psycopg2.extras",
            "psycopg2.pool",
            "docx",
            "dotenv",
            "playwright.sync_api",
            "openpyxl",
        ]
        + PLAYWRIGHT_HIDDEN_IMPORTS
        + OPENPYXL_HIDDEN_IMPORTS
    )
)


main_analysis = Analysis(
    [str(ROOT / "CALCULOS_QT.py")],
    pathex=[str(ROOT)],
    binaries=[],
    datas=main_datas,
    hiddenimports=main_hidden_imports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
main_pyz = PYZ(main_analysis.pure)
main_exe = EXE(
    main_pyz,
    main_analysis.scripts,
    [],
    exclude_binaries=True,
    name="CALCULOS_QT",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=[str(ASSETS / "favicon.ico")],
)


launcher_analysis = Analysis(
    [str(ROOT / "lanzador.py")],
    pathex=[str(ROOT)],
    binaries=[],
    datas=[],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
launcher_pyz = PYZ(launcher_analysis.pure)
launcher_exe = EXE(
    launcher_pyz,
    launcher_analysis.scripts,
    [],
    exclude_binaries=True,
    name="INICIAR_SISTEMA",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=[str(ASSETS / "favicon.ico")],
)


# Una sola distribución onedir: ambos ejecutables quedan juntos y comparten
# todas sus dependencias y recursos dentro de la carpeta _internal.
migrator_analysis = Analysis(
    [str(ROOT / "migrador_onedir.py")],
    pathex=[str(ROOT)],
    binaries=[],
    datas=[],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
migrator_pyz = PYZ(migrator_analysis.pure)
migrator_exe = EXE(
    migrator_pyz,
    migrator_analysis.scripts,
    migrator_analysis.binaries,
    migrator_analysis.datas,
    [],
    name="MIGRAR_A_ONEDIR",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=[str(ASSETS / "favicon.ico")],
)


updater_analysis = Analysis(
    [str(ROOT / "updater.py")],
    pathex=[str(ROOT)],
    binaries=[],
    datas=[],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
updater_pyz = PYZ(updater_analysis.pure)
updater_exe = EXE(
    updater_pyz,
    updater_analysis.scripts,
    updater_analysis.binaries,
    updater_analysis.datas,
    [],
    name="APLICAR_ACTUALIZACION",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=[str(ASSETS / "favicon.ico")],
)


distribution = COLLECT(
    launcher_exe,
    main_exe,
    updater_exe,
    launcher_analysis.binaries,
    launcher_analysis.datas,
    main_analysis.binaries,
    main_analysis.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name="HOSPITAL",
)
