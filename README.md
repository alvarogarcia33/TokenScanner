# Token Scanner Local

Scanner local para Netsbo con dos herramientas:
- `Token Scanner` para holders por token.
- `Wallet Report` para rastrear entradas, salidas y self-transfers de una wallet.

## Requisitos
- Python 3.11 o 3.12
- Internet para consultar el RPC de Netsbo

## Instalación
```powershell
cd token_scanner_local
python -m venv .venv
.\.venv\Scripts\activate
pip install -r requirements.txt
python app.py
```

Abrí en el navegador:

http://127.0.0.1:5000

## Publicarlo en internet
La forma más simple de dejarlo online es usar Render con un disco persistente.

Qué necesitás:
- una cuenta en Render
- este repo en GitHub
- un servicio `Web Service` con plan pago, porque el scanner necesita disco persistente para SQLite y reportes

Este repo ya quedó preparado para eso con:
- [render.yaml](./render.yaml)
- [requirements.txt](./requirements.txt)
- [.python-version](./.python-version)

Pasos:
1. Entrá a Render y conectá tu cuenta de GitHub.
2. Creá un nuevo `Blueprint` o `Web Service` usando este repo.
3. Si usás `Blueprint`, Render va a leer `render.yaml` y crear el servicio casi solo.
4. Si lo hacés manual, usá estos valores:
   - `Build Command`: `pip install -r requirements.txt`
   - `Start Command`: `gunicorn -w 1 -k gthread --threads 8 -b 0.0.0.0:$PORT app:app`
   - `Health Check Path`: `/healthz`
   - disco persistente en `/var/data`
   - variable `TOKEN_SCANNER_DATA_DIR=/var/data`
   - Python `3.13.5`
5. Esperá a que Render termine el deploy.
6. Primero probalo con la URL pública de Render, algo como `https://token-scanner.onrender.com`.
7. Después agregá tu dominio en Render y copiá los registros DNS que te pida tu proveedor de dominio.

Importante:
- El servicio debe quedar con `1` sola instancia. Hoy los jobs activos viven en memoria del proceso.
- Sin disco persistente, se perderían la base SQLite y los reportes en cada redeploy o reinicio.
- Si más adelante querés escalar a varias instancias, conviene mover jobs y estado a una base/cola externa.

## Operacion en Render
Estado actual esperado del sitio publicado:
- dominio principal: `https://imperium-tokenscanner.lat`
- subdominio de respaldo: `https://imperium-tokenscanner.onrender.com`
- plan del servicio: `Starter`
- una sola instancia

Checklist inicial:
1. Abrir `https://imperium-tokenscanner.lat` y confirmar que carga la UI.
2. En Render, verificar `Build Command`: `pip install -r requirements.txt`.
3. En Render, verificar `Start Command`: `gunicorn -w 1 -k gthread --threads 8 -b 0.0.0.0:$PORT app:app`.
4. Verificar `Health Check Path`: `/healthz`.
5. Verificar variable de entorno: `TOKEN_SCANNER_DATA_DIR=/var/data`.
6. Verificar disco persistente montado en `/var/data`.
7. Verificar `Auto-Deploy` en `On Commit`.
8. Verificar dominios custom activos: `imperium-tokenscanner.lat` y `www.imperium-tokenscanner.lat`.

Checklist de mantenimiento:
1. Cuando hagas cambios en el proyecto, subilos a `main` y Render redeploya solo.
2. Después de cada deploy, mirar logs de arranque y confirmar que no haya errores de importación o arranque de `gunicorn`.
3. Probar un escaneo corto de `Token Scanner`.
4. Probar un `Wallet Report` corto y descargar el `.xlsx`.
5. Si tocás DNS o dominio, mantener estos registros:
   - `@` -> `A` -> `216.24.57.1`
   - `www` -> `CNAME` -> `imperium-tokenscanner.onrender.com`
6. No agregar registros `AAAA` para el dominio si Render no los pidió.
7. Si alguna vez se expone públicamente el `Deploy Hook`, regenerarlo desde Render.
8. Los `Wallet Report` guardados en disco se limpian automáticamente con retención y, si hace falta, el sistema elimina reportes viejos para recuperar espacio.

Antes de borrar, recrear o mover el servicio:
1. Hacer backup de `scanner_data.db`.
2. Hacer backup de la carpeta `wallet_reports`.
3. Confirmar que el nuevo servicio también use `/var/data`.

## Qué hace ahora
- Construye un índice local persistente por token usando eventos `Transfer`.
- Reutiliza ese índice en escaneos futuros y solo sincroniza bloques nuevos.
- Guarda el estado local en SQLite para no perder el progreso entre reinicios.
- Revalida una ventana reciente de bloques antes de continuar para tolerar reorgs superficiales.
- Puede ejecutar una verificación estricta final con `balanceOf()` sobre los resultados filtrados.
- Permite cancelar jobs largos y exportar los resultados completos a XLSX o CSV.
- Genera reportes de wallet con todos los movimientos `IN`, `OUT` y `SELF` en un rango de días.
- Guarda los reportes de wallet en disco para poder descargarlos sin depender solo de la memoria del proceso.
- Está preparado para ejecutarse normal con Python o empaquetado como `.exe`.

## Cómo funciona
1. La primera vez que escaneás un token, la app sincroniza su historial de `Transfer` y crea un índice local.
2. En escaneos posteriores, la app lee ese índice y solo agrega bloques nuevos.
3. Antes de seguir, la app puede resincronizar una ventana reciente para evitar inconsistencias por reorgs superficiales.
4. El filtro por balance se hace sobre la base local, evitando llamar `balanceOf()` wallet por wallet.
5. Si activás la verificación estricta, la salida final se confirma otra vez on-chain con `balanceOf()`.

Esto acelera muchísimo los usos repetidos sin perder consistencia del historial ya procesado.

## Notas
- La primera sincronización de un token puede tardar bastante; después los escaneos son mucho más rápidos.
- El índice se construye desde eventos `Transfer` y se mantiene en SQLite.
- La app reconsulta una ventana reciente de bloques para mantener la consistencia del índice.
- Si un token define `startBlock` en `tokens.json`, el índice empieza desde ahí para acelerar el proceso.
- La exportación XLSX y CSV descarga todos los resultados del job.
- El `Wallet Report` revisa eventos `Transfer` donde la wallet aparece como origen o destino y arma un Excel con hoja de resumen y hoja de movimientos.
- Para no perder precisión, el Excel del `Wallet Report` guarda `Amount` como texto exacto y también incluye `Raw Value`.
- En ejecución normal, la base `scanner_data.db` queda junto a `app.py`.
- En modo `.exe`, la base y los reportes se guardan por defecto en `%LOCALAPPDATA%\TokenScannerLocal`.

## Variables útiles
Si querés cambiar el RPC:

```powershell
$env:NETSBO_RPC_URL="https://rpc1.netsbo.io"
python app.py
```

También podés ajustar el comportamiento del scanner:

```powershell
$env:TOKEN_SCANNER_BATCH_SIZE="5000"
$env:TOKEN_SCANNER_MIN_BATCH_SIZE="100"
$env:TOKEN_SCANNER_CONFIRMATION_BLOCKS="6"
$env:TOKEN_SCANNER_REORG_LOOKBACK_BLOCKS="25"
$env:TOKEN_SCANNER_STRICT_RESULT_VALIDATION="0"
$env:TOKEN_SCANNER_JOB_RETENTION_SECONDS="3600"
$env:TOKEN_SCANNER_DATA_DIR="C:\mis_datos\token_scanner"
$env:TOKEN_SCANNER_WALLET_REPORT_BATCH_SIZE="3000"
$env:TOKEN_SCANNER_WALLET_REPORT_MIN_BATCH_SIZE="100"
$env:TOKEN_SCANNER_WALLET_REPORT_MAX_DAYS="3650"
$env:TOKEN_SCANNER_WALLET_REPORT_RETENTION_SECONDS="604800"
python app.py
```

## Construir el `.exe`
Hay un script listo para compilar la app con PyInstaller:

```powershell
.\build_exe.ps1
```

Eso genera un ejecutable en `dist\TokenScannerLocal.exe`.

Para generar la release final con el nombre principal:

```powershell
.\build_release.ps1
```

Eso deja una sola release principal en:

`release\TokenScannerLocal.exe`

## Construir el instalador
Hay un script para compilar un instalador Windows con Inno Setup:

```powershell
.\build_installer.ps1
```

El instalador resultante queda en:

`release\TokenScannerLocal-Setup.exe`

El script busca `ISCC.exe` en instalación normal o por usuario. Si no lo encuentra, instalá Inno Setup o pasá la ruta manualmente:

```powershell
.\build_installer.ps1 -IsccPath "C:\ruta\ISCC.exe"
```

El instalador está configurado para instalar por usuario en:

`%LOCALAPPDATA%\Programs\TokenScannerLocal`

Nota:
- El build quedó probado con Python 3.14, pero `web3`/`pydantic` emiten warnings en ese entorno.
- Para una distribución más conservadora y estable, conviene generar el `.exe` final con Python 3.12.
- En esta máquina también quedó probado un build funcional con Python 3.13.

Si querés instalar primero las dependencias de build:

```powershell
.\.venv\Scripts\python -m pip install -r requirements-build.txt
```
