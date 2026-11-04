# Daily Exception — Python

Programa de consola para obtener los datos de Power BI, ejecutar el SQL de Oracle y crear el reporte final. No usa Power Automate ni ejecuta Office Scripts: reproduce el script suministrado en Python.

## Estado real

La transformación se prueba con los Excel suministrados. Los conectores en vivo están implementados, pero NO se han probado contra tus cuentas. Para ejecutar todo sin archivos locales faltan la conexión Oracle y el registro/acceso al modelo de Power BI, además de la consulta DAX real del visual. El título del visual no permite deducir esa consulta. El enlace de una aplicación Power BI tampoco contiene el dataset_id.

## Instalación (Windows, Python 3.11 o superior)

Descomprimir el ZIP, abrir una terminal en esta carpeta y ejecutar:

```powershell
py -m venv .venv
    .venv\Scripts\python.exe -m pip install -r requirements.txt
Copy-Item config.example.json config.json
```

## Prueba con los archivos que ya tienes

```powershell
        .venv\Scripts\python.exe main.py --raw "C:\Users\E1568382\OneDrive - Emerson\proyectos\Daily_Exception_Python\bases\RAW DATA DAILY EXCEPTION (14).xlsx" --enhance "C:\Users\E1568382\OneDrive - Emerson\proyectos\Daily_Exception_Python\bases\ENHANCE.xlsx" --output-dir "C:\Users\E1568382\OneDrive - Emerson\TF Purchasing Analytics - Documents\General Reports\2. Inventory Management\7. Enhanced Daily Exception Messages" --log-dir "C:\Users\E1568382\OneDrive - Emerson\proyectos\Daily_Exception_Python\tests\logs"
```

Esta prueba no necesita credenciales ni conexión a Power BI/Oracle. No vuelve a usar el reporte terminado como plantilla: reconstruye las 92 columnas originales y las diez nuevas.

Cada ejecución crea una subcarpeta única con:

- `Revision Enhanced Daily Exception Messages AAAA-MM-DD.xlsx`: reporte final.
- `RAW DATA DAILY EXCEPTION.xlsx`: datos de entrada de esa ejecución.
- `ENHANCE.xlsx`: resultado SQL de esa ejecución.
- `resultado.json`: conteos, coincidencias y avisos.

La fecha del nombre corresponde a America/Mexico_City. No sobrescribe ejecuciones previas.

## Reglas del cruce y diferencias del ejemplo

- Por defecto busca solamente por ITEM, como describiste para VLOOKUP. Es insensible a mayúsculas; conserva espacios y ceros iniciales. Devuelve la primera coincidencia sin volver a sumar.
- El ejemplo RAW tiene MMC y TFC, pero el SQL únicamente devuelve MMC (organización 2035). Buscar solo por ITEM puede traer datos MMC a una fila TFC. El programa cuenta y avisa estas coincidencias; no cambia tu regla en silencio.
- Para buscar por organización e ítem usa `--lookup org_item`. Si conservas el SQL actual, las filas TFC no tendrán coincidencia.
- `TOTAL_QTY_LAST_24H` va a **LAST FLOOR TRANS. QTY** y `NEWEST_TRANSACTION_DATE` a **LAST FLOOR TRANS. DATE**. En el archivo terminado proporcionado esos dos contenidos están invertidos. Este programa sigue los nombres y tu descripción, no esa inversión.
- Conserva el signo SQL de la cantidad (por ejemplo -630). El ejemplo terminado tiene cantidades positivas; no se aplica ABS al dato fuente sin una regla explícita. La fórmula de cobertura sí usa ABS, tal como tu script.
- Sin coincidencia deja QTY/DATE vacíos. Si quieres el error de un VLOOKUP sin IFERROR, agrega `--missing na`.
- Excluye la fila vacía y la nota `Applied filters` cuando no tienen ITEM; no las trata como artículos. Conserva todas las filas que sí tienen ITEM, incluso ítems repetidos.
- No calcula días transcurridos, ESTIMATED FLOOR DOH ni ESTIMATED FLOOR QOH: tu script los deja vacíos. Mantiene también `REMAINING WEEKS OF COVERAGE (DELTA) = DAYS / 5`, aunque el nombre sugiera otra operación.
- Traduce las cuatro fórmulas del script a referencias por celda. Incluye valores calculados iniciales y solicita recálculo al abrir Excel. Conserva columnas ocultas y bordes; no crea una tabla porque el script termina convirtiéndola a rango.
- REVISION COMMENT no tiene relleno especial en el script actual; se mantiene así.

## Configurar Oracle

Editar la sección `oracle` de `config.json`:

| Campo | Uso |
|---|---|
| mode | `thin` o `thick`, según tu instalación/base |
| host / port | Servidor y puerto |
| connection_type | `service_name` o `sid` |
| service_name / sid | Completar el correspondiente |
| user | Usuario; también puede introducirse al ejecutar |
| dsn | Alternativa: descriptor/TNS configurado; si se completa se usa en lugar de host/port |
| client_lib_dir | Carpeta Oracle Instant Client si usas Thick |
| config_dir | Carpeta opcional de configuración de Oracle en Thick |

La contraseña se pide en terminal sin mostrarse. Para una ejecución gestionada también se leen `ORACLE_USER` y `ORACLE_PASSWORD` del entorno; no se guardan en el JSON ni en el ZIP. La conexión necesita acceso de red/VPN a Oracle y permisos SELECT sobre las tablas. Se inicia una transacción de solo lectura y no se realizan cambios en la base.

`transactions.sql` es la consulta incluida en ENHANCE, sin comandos SET ni punto y coma final. Mantiene organización 2035, action=1, source=13, signo y ventana inclusiva de 24 horas relativa a la última transacción de cada ítem. No significa últimas 24 horas desde hoy.

## Configurar Power BI

Se utiliza la API Execute Queries con autenticación interactiva de Microsoft mediante MSAL. No automatiza clics del navegador. Debe configurarse una aplicación de cliente público en Microsoft Entra, con redirect URI de escritorio `http://localhost` y permiso delegado de Power BI `Dataset.Read.All`; puede requerir aprobación de TI.

En `config.json`, completar:

- `powerbi.tenant_id`: directorio de tu empresa.
- `powerbi.client_id`: identificador de esa aplicación Entra.
- `powerbi.dataset_id`: identificador del modelo semántico que alimenta el reporte.
- Crear `query.dax` junto a main.py con la consulta real del visual, las 92 columnas y sus filtros. No se incluye una consulta inventada.

Filtros observados en la nota del RAW suministrado:

- ORG is MMC or TFC.
- PLAN_NAME is TODAY_UNC.
- COMPANY is not PWS_USA_OU.

Esos filtros deben incorporarse a la consulta DAX junto con el orden deseado. La consulta API no hereda los filtros del enlace ni de una sesión del navegador. Comparar los conteos y artículos con una exportación manual del mismo momento antes de usar el resultado en producción.

El usuario necesita Read + Build sobre el modelo y la empresa debe habilitar Dataset Execute Queries REST API. Se requiere una sola tabla de resultados por consulta. La API tiene límites de 100.000 filas, un millón de valores y 15 MB por consulta. Con 92 columnas el límite por valores es aproximadamente 10.869 filas, y el de tamaño puede alcanzarse antes. El programa rechaza errores reportados y resultados que alcancen los límites de filas/valores. No pagina automáticamente: si se excede el volumen, hay que dividir la consulta con criterios completos y sin solapamientos. No usar TOPN para esconder filas.

`raw_headers.json` conserva los nombres y el orden exactos del RAW de ejemplo. Se reconocen encabezados DAX `Tabla[ITEM]` o `[ITEM]`. Si un nombre es distinto o ambiguo, añadir en `column_mapping` una relación explícita, por ejemplo `"ITEM": "Tabla real[Nombre real]"`. No renombrar ítems ni convertirlos a números.

Para validar conectores por separado:

```powershell
# Archivo Power BI local + consulta Oracle en vivo:
.venv\Scripts\python.exe main.py --raw "C:\Reportes\RAW DATA DAILY EXCEPTION (14).xlsx"

# Power BI en vivo + resultado SQL local:
.venv\Scripts\python.exe main.py --enhance "C:\Reportes\ENHANCE.xlsx"

# Todo en vivo, una vez completada la configuración:
.venv\Scripts\python.exe main.py
```

El modo en vivo abre el inicio de sesión Microsoft en cada proceso; no se ha implementado ejecución desatendida ni persistencia del token. Se respeta MFA y la política de la empresa.

## Documentación oficial

- https://learn.microsoft.com/en-us/rest/api/power-bi/datasets/execute-queries
- https://learn.microsoft.com/en-us/entra/msal/python/getting-started/acquiring-tokens
- https://python-oracledb.readthedocs.io/en/stable/user_guide/connection_handling.html
- https://xlsxwriter.readthedocs.io/
