# IAPAY · Atención al Comercio (SoundBox Q191)

App de escritorio Windows (.exe) ligada a `IAPAY_Atencion_Comercio.xlsx`
(se crea sola junto al .exe; ruta alternativa: `config.json` -> {"excel_path": "C:\\ruta\\archivo.xlsx"}).

## Obtener el .exe sin instalar nada
1. Cree un repositorio privado en GitHub y suba el contenido de esta carpeta (incluida `.github`).
2. Pestaña **Actions** -> "Compilar EXE (Windows)" -> **Run workflow**.
3. Al terminar, descargue el artefacto `IAPAY_Atencion_Comercio_exe` (zip con el .exe).

## Alternativa
En cualquier PC con Python: doble clic en `build_exe.bat` -> `dist\IAPAY_Atencion_Comercio.exe`.

Copie el .exe a la PC de trabajo y ábralo. Mantenga el Excel cerrado mientras guarda desde la app.
