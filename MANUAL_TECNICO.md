# Manual Técnico — SOW Manager

Documentación para quien mantenga o extienda el código: qué hace cada parte, cómo fluyen los datos
y cómo intervenir sin romper nada.

---

## 1. Stack tecnológico y por qué se eligió

| Tecnología | Uso | Por qué |
|---|---|---|
| **Python + Tkinter** | GUI de escritorio (`src/ui.py`) | Viene incluido con Python, sin dependencias de UI externas, y empaqueta fácil con PyInstaller en un `.exe` que corre en la máquina de la usuaria sin instalar nada. |
| **pywin32** (`win32com.client`) | Automatización de Outlook (`src/outlook_reader.py`) | Controla el Outlook de escritorio ya instalado y logueado vía COM. Se prefirió sobre Microsoft Graph API porque no requiere registrar una app en Azure AD ni gestionar OAuth — solo que Outlook esté abierto. |
| **openpyxl** | Lectura/escritura de `.xlsm` en memoria | Entiende estilos, fórmulas, tablas de Excel e hipervínculos. **Importante:** su método `.save()` completo nunca se usa directamente contra el archivo real — ver sección 4, porque reescribe partes del OOXML que openpyxl no modela (Power Query, pivots, slicers) y las corrompe. |
| **python-docx** + `xml.etree.ElementTree` | Lectura de los `.docx` de SOW (`src/word_extractor.py`) | python-docx para navegar párrafos/tablas cómodamente; acceso directo al XML (`doc.element.body`) para leer *Content Controls* (elementos `<w:sdt>`) con tags específicos como `StartDateSOW`/`EndDateSOW`, que la API de alto nivel de python-docx no expone. |
| **BeautifulSoup + lxml** | Parseo de la tabla HTML del cuerpo del correo | Se usa directamente en `src/main.py` (`parsear_tabla_sin_encabezados`) para una tabla sin encabezados con formato fijo (SOW / Empresa / Personas). |
| **PyInstaller** | Empaquetado a `.exe` | La usuaria final no tiene Python instalado; se distribuye un ejecutable autocontenido. |

No hay versión de Python fijada explícitamente en el proyecto. Se recomienda una versión reciente de Python 3 (3.11+).

---

## 2. Estructura de carpetas y archivos

```
sow_automation/
├── run.py                      Punto de entrada. Ajusta sys.path (funciona tanto
│                                como script como bundle de PyInstoller) y lanza
│                                la ventana principal (src.ui.AppSOWs).
├── src/
│   ├── main.py                 Orquestador del flujo completo (procesar_correos()).
│   ├── outlook_reader.py       Conexión a Outlook vía COM, lectura de correos y adjuntos.
│   ├── excel_writer.py         Motor de lectura/escritura del Excel maestro — el módulo
│   │                            más importante, ver sección 4.
│   ├── word_extractor.py       Extracción de fechas/valor y análisis de 6 puntos de los .docx.
│   ├── archivador.py           Guarda los adjuntos en disco, organizados por Persona/Año.
│   └── ui.py                   Interfaz Tkinter (tabs Pendientes/Novedades, modal de detalle).
├── requirements.txt            Dependencias directas del proyecto. Ver notas en sección 6.
├── build.bat / "SOW Manager.spec"          Build de la app principal (PyInstaller --onedir --windowed).
├── .gitignore
└── README.md                    Portada corta del repo, enlaza a este documento y a MANUAL_DE_USO.md.
```

`build/` y `dist/` no están en el repo (quedan excluidas por `.gitignore`): las genera PyInstaller
cada vez que se compila. `build/` es una carpeta de trabajo intermedia del compilador (no se usa para
nada fuera de la compilación); `dist/SOW Manager/` es la carpeta con el ejecutable final que se le
entrega a la usuaria. Ambas se recrean solas al correr `build.bat` — no hay que tocarlas a mano ni
versionarlas.

---

## 3. Flujo de datos

```
Outlook (carpeta "SOWs")
        │  outlook_reader.conectar_outlook()
        │  fuerza sync real (SyncObjects + SendAndReceive + espera 15-35s)
        ▼
main.procesar_correos()
        │
        ├─ 1) listar_correos()               → lista de correos con cuerpo HTML/adjuntos
        │
        ├─ 2) parsear_tabla_sin_encabezados() → por correo, extrae filas [SOW | Empresa | Personas]
        │      construir_indice_global()      → combina TODOS los correos del batch en un índice
        │                                        { "NEO-SERV-2026-0737": ["PERSONA A", "PERSONA B"] }
        │      (esto permite que el PDF firmado llegue en un correo/respuesta distinto al que
        │       tenía la tabla con los nombres, y aun así se pueda vincular por número de SOW)
        │
        ├─ 3) por cada adjunto PDF/Word:
        │      extraer_numero_sow_de_archivo() → saca el SOW del NOMBRE del archivo (regex)
        │      word_extractor.extraer_datos_word()  → SOLO si es .docx: fecha_inicio, fecha_fin,
        │                                              valor_cop (SDT StartDateSOW/EndDateSOW,
        │                                              o fallback a tablas del documento)
        │      archivador.archivar_pdf_para_varias_personas()
        │             → guarda el archivo en [ruta_raiz]/[Persona]/[Año]/ para CADA persona
        │               indicada en el índice global para ese SOW (un SOW puede ser compartido)
        │
        └─ 4) (solo si escribir_excel=True — la UI no lo usa así, ver más abajo)
               excel_writer.escribir_sows_en_excel() / escribir_actualizaciones_excel()
```

**En la UI (`ui.py`), el flujo es de dos pasos**, no uno solo:

1. `on_sincronizar()` → llama a `procesar_correos(..., escribir_excel=False)`: solo lee y arma
   la estructura de "novedades" en memoria — **todavía no escribe nada en el Excel**. Para decidir
   qué mostrar como nuevo, compara contra `excel_writer.leer_indice_excel()`, que devuelve el estado
   actual de cada `(SOW, persona)` ya registrado en el Excel (`"cerrado"` si la columna R tiene
   valor, `"pendiente"` si está vacía).
2. El usuario revisa y presiona "Guardar" → recién ahí se llama a `escribir_sows_en_excel()` /
   `escribir_actualizaciones_excel()`, que hacen el guardado real (sección 4).

Este diseño de dos pasos es deliberado: separa "qué encontramos" de "qué se confirma guardar",
y evita escribir en el Excel por cada sincronización aunque el usuario no revise nada.

**Nota:** `main.py` también expone un punto de entrada directo (`if __name__ == "__main__":`) que
llama a `procesar_correos(..., escribir_excel=True)` — útil para correr todo el flujo por consola
sin la UI (pruebas, automatización por línea de comandos).

---

## 4. El motor de escritura OOXML (`excel_writer.py`) — la parte más compleja

### 4.1 El problema que resuelve

`Listado SOW's.xlsm` tiene Power Query, tablas dinámicas, slicers y macros VBA. Un `.xlsm` es en
realidad un **archivo ZIP** con varios XML adentro (formato OOXML): `xl/worksheets/sheetN.xml` para
cada hoja, `xl/sharedStrings.xml` para el texto, `xl/styles.xml` para estilos, `xl/queryTables/` y
`xl/connections.xml` para Power Query, `xl/vbaProject.bin` para las macros, etc.

`openpyxl` sabe **leer** casi todo eso, pero cuando hace `wb.save(archivo)` **reescribe el ZIP
completo desde su propio modelo interno**. Cualquier parte del OOXML que openpyxl no modela
completamente (Power Query, cachés de tablas dinámicas, slicers) se pierde o se corrompe en esa
reescritura. Controlar Excel directamente (`win32com`/`xlwings`) tampoco es viable para este caso:
sin archivo abierto interactivamente, Excel termina lanzando el diálogo de "Guardar como" o falla
con errores de RPC, porque esa automatización está pensada para uso interactivo, no para escritura
desatendida en segundo plano.

### 4.2 La solución: reemplazo quirúrgico de archivos dentro del ZIP

La idea central (`_guardar_zip_quirurgico`, en `excel_writer.py`): **dejar que openpyxl haga todo su
trabajo normal en memoria**, pero **nunca dejar que su ZIP de salida reemplace el archivo real**.
En cambio, se construye un ZIP nuevo copiando el original **archivo por archivo**, y solo se
reemplazan los pocos archivos internos que realmente cambiaron.

Paso a paso de lo que hace la función:

1. **Modificaciones normales con openpyxl, todo en memoria.** El resto del módulo (`_escribir_fila_ox`,
   `_copiar_estilo_fila_ox`, etc.) trabaja sobre el objeto `Workbook` de openpyxl exactamente como
   cualquier script de openpyxl: escribe celdas, copia estilos de una fila plantilla, ajusta fórmulas.
   Nada de esto toca el disco todavía.

2. **Extender el rango de la Tabla de Excel.** Si la hoja `SOWS` está definida como una "Tabla" de
   Excel (`ws.tables`), su rango (`tbl.ref`, ej. `A7:T230`) está guardado explícitamente en el XML de
   la tabla. Si no se extiende ese rango para cubrir las filas nuevas, Excel no las reconoce como
   parte de la tabla (se rompen filtros, formato de tabla, referencias estructuradas). El código
   recalcula la última fila real con datos y reescribe `tbl.ref` antes de guardar.

3. **`wb_openpyxl.save(buf)` a un `BytesIO` en memoria** (nunca al archivo real). Esto genera un ZIP
   completo "a la manera de openpyxl", con todas sus partes — incluidas las que NO nos interesan
   tocar (Power Query, etc., que openpyxl reescribe con una versión empobrecida o las pierde).

4. **Ubicar el path real de la hoja dentro del ZIP** (`_encontrar_ruta_hoja`): en OOXML el nombre de
   hoja (`SOWS`) no corresponde directamente a un nombre de archivo fijo como `sheet1.xml` — hay que
   resolverlo leyendo `xl/workbook.xml` (para el `r:id` de la hoja por nombre) y luego
   `xl/_rels/workbook.xml.rels` (para el path real asociado a ese `r:id`). Se hace así en vez de
   asumir `sheetN.xml` porque el número puede no coincidir con el orden de las hojas visibles.

5. **Post-procesar el XML de la hoja para extender el formato condicional.** Las reglas de formato
   condicional (colores automáticos de la columna de alertas, por ejemplo) tienen un rango fijo
   (`sqref="$E$8:$E$230"`). Si no se extiende, las filas nuevas no heredan el color automático. Se
   hace con una regex sobre el XML ya generado por openpyxl, extendiendo cualquier rango que termine
   antes de la última fila real a `...:$COL$<última_fila>`.

6. **Definir la lista de archivos a reemplazar** (`REEMPLAZAR`): únicamente
   - el XML de la hoja `SOWS` (`xl/worksheets/sheetN.xml`)
   - sus relaciones — hipervínculos — (`xl/worksheets/_rels/sheetN.xml.rels`)
   - `xl/sharedStrings.xml` (tabla de strings compartidos)
   - `xl/calcChain.xml` (orden de cálculo de fórmulas)
   - `xl/styles.xml` (estilos — se puede reemplazar completo porque openpyxl lo cargó del **mismo**
     archivo original antes de modificarlo, así que su salida es una re-serialización equivalente
     más los estilos nuevos agregados por el script, no un estilo distinto)
   - cualquier `xl/tables/table*.xml` (definición de Tablas de Excel)

7. **Armar el ZIP final copiando el original entero**, reemplazando solo esos archivos:
   - Para cada archivo del ZIP **original**: si su nombre está en `REEMPLAZAR` y existe en el ZIP
     nuevo de openpyxl, se escribe la versión **nueva**; si no, se copia el **byte original tal
     cual**, sin pasar por openpyxl en absoluto.
   - Esto es lo que preserva intactos `xl/queryTables/`, `xl/connections.xml` (Power Query),
     `xl/vbaProject.bin` (macros), definiciones de tablas dinámicas y slicers: openpyxl nunca llega
     a tocarlos porque **directamente no se usa su versión de esos archivos**.
   - También se agregan archivos nuevos que openpyxl haya generado y no existieran antes (por
     ejemplo, un `sheetN.xml.rels` nuevo si la hoja no tenía hipervínculos previamente).

8. **Escritura atómica.** El ZIP nuevo se arma en un archivo temporal (`ruta_excel + ".writing"`) y
   solo al final se hace `os.replace(temp_path, ruta_abs)`, que en Windows reemplaza el archivo de
   forma atómica. Si algo falla en el camino, se borra el temporal y se relanza la excepción — el
   archivo original nunca queda a medio escribir.

### 4.3 Por qué esto funciona y las alternativas fallaban

- No se abre Excel ni se usa COM para guardar (`win32com`/`xlwings`) → no hay diálogos de "Guardar
  como" ni errores de RPC, porque no hay ninguna instancia de Excel involucrada en absoluto.
- openpyxl solo se usa como "motor de edición en memoria": su capacidad de escribir celdas, copiar
  estilos y manejar fórmulas se aprovecha, pero su serialización final **solo se acepta para las
  partes que él mismo entiende y que necesitábamos cambiar** (la hoja de datos, strings, estilos,
  fórmulas, definición de tabla). Todo lo que openpyxl no modela bien (Power Query, pivots, slicers,
  VBA) nunca pasa por su serializador — se copia byte a byte del archivo original.
- **Consecuencia práctica para quien mantenga esto:** si en el futuro hace falta que el script
  modifique algo que vive en una de esas partes "no tocadas" (por ejemplo, refrescar una tabla
  dinámica), hay que agregar explícitamente ese archivo a `REEMPLAZAR` y asegurarse de que la
  versión de openpyxl de esa parte sea correcta — si no, seguirá copiándose la versión vieja del
  original.

### 4.4 Funciones clave del módulo

| Función | Qué hace |
|---|---|
| `_guardar_zip_quirurgico(ruta_excel, wb_openpyxl, nombre_hoja)` | Todo lo descrito arriba — el corazón del módulo. |
| `_encontrar_ruta_hoja(zf, nombre_hoja)` | Resuelve el path interno del XML de una hoja por su nombre visible. |
| `_encontrar_ultima_fila_ox(ws)` | Recorre desde la fila 8 hasta encontrar 15 filas vacías seguidas, para ubicar dónde termina la data real. |
| `_encontrar_fila_template_ox(ws, desde_fila)` | Busca hacia atrás la última fila que tenga fórmulas en columnas E y M, para usarla como plantilla de estilo/fórmulas de la fila nueva. |
| `_escribir_fila_ox(ws, nueva_fila, datos, fila_template)` | Escribe los valores de negocio y copia/desplaza las fórmulas de las columnas E, K, L, M, T (ver tabla de columnas abajo) desde la fila plantilla. Maneja tanto fórmulas de texto normales como `ArrayFormula` (BUSCARX/XLOOKUP dinámico, que openpyxl representa distinto a una fórmula de texto simple). |
| `escribir_sows_en_excel(lista_sows, ruta_excel)` | API pública: agrega filas nuevas (Words sin firmar). Evita duplicados comparando `(SOW, nombre)` contra las filas ya existentes. |
| `escribir_actualizaciones_excel(lista_actualizaciones, ruta_excel)` | API pública: marca como "Firmado" una fila ya existente y le agrega el hipervínculo al PDF firmado. |
| `leer_sows_del_excel(ruta_excel)` | Lee los SOWs pendientes (columna R vacía) para poblar la pestaña Pendientes. |
| `leer_indice_excel(ruta_excel)` | Devuelve `{ (sow, nombre): "cerrado"/"pendiente" }` para TODAS las filas — se usa para decidir qué mostrar como "novedad" al sincronizar. |

### 4.5 Columnas del Excel (hoja `SOWS`, encabezados en fila 7, datos desde fila 8)

| Col | Letra | Constante | Origen del valor |
|---|---|---|---|
| 1 | A | `COL_PAIS` | Escrito por el script (inferido de la empresa, o `"COLOMBIA"` por defecto). |
| 2 | B | `COL_PROYECTO` | Escrito por el script (`PROYECTO_NEORIS` por defecto). |
| 3 | C | `COL_NOMBRE` | Escrito por el script (nombre de la persona, en mayúsculas). |
| 4 | D | `COL_EMPLOYEE_ID` | Escrito por el script (buscado en la hoja "Query" del mismo libro si existe). |
| 5 | E | `COL_TALENTO_ACTIVO` | **Fórmula** — se copia/desplaza desde la fila plantilla, no la escribe el script. |
| 6 | F | `COL_ESTADO_SOW` | Se deja vacío por el script. |
| 7 | G | `COL_SOW` | Escrito por el script (número de SOW). |
| 8 | H | `COL_ANIO` | Escrito por el script (extraído del número de SOW). |
| 9 | I | `COL_FECHA_INICIO` | Escrito por el script (extraído del Word). |
| 10 | J | `COL_FECHA_FIN` | Escrito por el script (extraído del Word). |
| 11 | K | `COL_ALERTA` | **Fórmula** copiada/desplazada desde la plantilla. |
| 12 | L | `COL_PERIODO` | **Fórmula** copiada/desplazada desde la plantilla. |
| 13 | M | `COL_DIAS` | **Fórmula** copiada/desplazada desde la plantilla (días restantes). |
| 14 | N | `COL_COMENTARIOS` | Se deja vacío. |
| 15 | O | `COL_COMENTARIOS_2` | Se deja vacío. |
| 16 | P | `COL_FIRMADO` | Escrito por el script (`"Si"` si ya llegó el PDF firmado). |
| 17 | Q | `COL_CARPETA` | Escrito por el script — hipervínculo a la carpeta de la persona. |
| 18 | R | `COL_RUTA_ARCHIVO` | Escrito por el script — hipervínculo al PDF firmado. **Esta es la columna que define si un SOW está "cerrado" o "pendiente"** en `leer_indice_excel`. |
| 19 | S | `COL_VALOR_COP` | Escrito por el script (extraído del Word). |
| 20 | T | `COL_VALOR_USD` | **Fórmula** copiada/desplazada desde la plantilla. |

---

## 5. Cómo hacer modificaciones comunes

**Agregar una columna nueva al Excel:** agregar la constante `COL_*` correspondiente en
`excel_writer.py`, escribirla en `_escribir_fila_ox`, y si es una columna con fórmula (como E/K/L/M/T),
agregarla a la lista de letras que se copian/desplazan desde la fila plantilla en esa misma función.
No hace falta tocar `_guardar_zip_quirurgico`: mientras la columna esté dentro de la misma hoja `SOWS`,
ya está cubierta por el reemplazo del XML de la hoja.

**Agregar un punto de validación al análisis del Word:** agregar una función `_validar_*` en
`word_extractor.py` siguiendo el patrón de las existentes (recibe texto/doc, devuelve
`{"ok": bool, "detalle": ...}`), llamarla desde `analizar_sow_completo()`, y agregar su bloque
correspondiente en el modal de `ui.py` (`_abrir_modal`, sección de bloques `b1`, `b_cm`, `b_ts`, etc.).

**Cambiar las rutas del Excel / carpeta de archivado:** ⚠️ las rutas están **duplicadas en tres
lugares** — `RUTA_EXCEL_PRUEBA` en `excel_writer.py`, `RUTA_RAIZ_DEFAULT` en `archivador.py`, y
`RUTA_SOWS`/`RUTA_EXCEL` (con su propia constante `_BASE`) en `ui.py`. Si el proyecto cambia de
ubicación de carpeta o de proyecto de SharePoint, hay que actualizar los tres. No hay un
`config.py`/`.yaml` centralizado en el proyecto — si se necesita configuración externa (por ejemplo,
para no hardcodear rutas), es un punto pendiente a resolver.

**Cambiar el formato de la tabla del correo (nuevo remitente/formato):** el parser vive en
`main.py::parsear_tabla_sin_encabezados`. Asume una tabla HTML sin encabezados con columnas fijas
`[SOW, Empresa, Personas]` y que la primera columna matchea el patrón `NEO[A-Z]*-SERV-\d{4}-\d+`. Un
formato de tabla distinto (con encabezados, por ejemplo) requiere escribir un parser nuevo para ese
caso.

**Recompilar el ejecutable:** con el entorno virtual activado, `build.bat`. Usa PyInstaller con el
`.spec` del repo; el resultado queda en `dist/SOW Manager/`.

---

## 6. Dependencias y cómo configurar el entorno desde cero

### 6.1 Instalación

```bat
python -m venv venv
venv\Scripts\activate
pip install -r requirements.txt
```

### 6.2 Notas sobre `requirements.txt`

Lista únicamente las dependencias que el código en `src/` realmente importa (directas o de
propósito claro): `beautifulsoup4`/`lxml` para parsear el HTML del correo, `openpyxl` para el
Excel, `python-docx` para los Word, `pywin32` para Outlook, y el resto son dependencias de esos
paquetes. No incluye librerías de análisis de datos ni de lectura de PDF porque nada en `src/` las
usa — los PDFs adjuntos solo se archivan en disco, nunca se leen.

### 6.3 Requisitos del entorno de ejecución

- Windows (la dependencia de `pywin32`/COM con Outlook es exclusiva de Windows).
- Microsoft Outlook de escritorio instalado, con una cuenta configurada y con sesión iniciada.
- Acceso (vía OneDrive sincronizado) a la carpeta de SharePoint del proyecto, en la ruta esperada
  por las constantes de la sección 5.

### 6.4 Compilar el ejecutable

```bat
build.bat   REM → dist\SOW Manager\SOW Manager.exe
```

### 6.5 Tests

No hay suite de tests automatizados en el proyecto todavía.
