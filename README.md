# Hospital Contreras - Facturacion

Aplicacion de escritorio para facturacion medica, construida con Python,
PySide6 y PostgreSQL.

## Desarrollo local

1. Crea y activa un entorno virtual.
2. Instala las dependencias con `pip install -r requirements.txt`.
3. Copia `config_local.example.py` como `config_local.py` y configura
   `DATABASE_URL`. Este archivo es privado y Git lo ignora.
4. Ejecuta `python CALCULOS_QT.py`.

## Compilacion para Windows

La aplicacion se distribuye en modo PyInstaller `onedir`:

```powershell
pyinstaller --noconfirm --clean build_app.spec
```

El resultado completo se genera en `dist/HOSPITAL`. Para una release se debe
comprimir esa carpeta completa; no se debe publicar solo `CALCULOS_QT.exe`.

## Actualizaciones

`INICIAR_SISTEMA.exe` consulta el manifiesto remoto, verifica la version y el
SHA-256 del ZIP, y delega la sustitucion completa de `onedir` a
`APLICAR_ACTUALIZACION.exe`. El instalador conserva recibos, reportes y
respaldos locales, y restaura la version anterior si la instalacion falla.
