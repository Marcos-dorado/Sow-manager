"""
excel_writer.py

ESTRATEGIA DEFINITIVA — sin win32com para guardar:
Un .xlsm es un archivo ZIP. Modificamos los datos con openpyxl en memoria,
luego construimos el ZIP de salida copiando el original archivo por archivo,
reemplazando SOLO los que openpyxl necesita cambiar:
  - La hoja SOWS (xl/worksheets/sheetN.xml)
  - Sus relaciones (hipervínculos) (xl/worksheets/_rels/sheetN.xml.rels)
  - La tabla de strings (xl/sharedStrings.xml)
  - La cadena de cálculo (xl/calcChain.xml)

TODO LO DEMÁS viene del original sin tocar:
  - Power Query (xl/queryTables/, xl/connections.xml)
  - VBA (xl/vbaProject.bin)
  - Slicers, tablas dinámicas, etc.

Sin Excel abierto → sin diálogos posibles.
"""

import os
import re
import io
import zipfile
import shutil
from copy import copy
from datetime import datetime
from xml.etree import ElementTree as ET

import warnings
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")
warnings.filterwarnings("ignore", message=".*Slicer.*")

# RUTA_EXCEL_PRUEBA = os.path.join(
#     os.path.expanduser("~"), "Desktop", "Listado SOW's.xlsm"
# )

RUTA_EXCEL_PRUEBA = os.path.join(
    os.path.expanduser("~"), "EPAM",
    "92101343 - CHUBB COLOMBIA - 1 DEV FT - Documents",
    "0. Documentacion Contractual",
    "Listado SOW's.xlsm")

NOMBRE_HOJA       = "SOWS"
FILA_INICIO_DATOS = 8
PROYECTO_NEORIS   = "92101343 CHUBB Colombia - 1 Dev FT"

COL_PAIS           = 1
COL_PROYECTO       = 2
COL_NOMBRE         = 3
COL_EMPLOYEE_ID    = 4
COL_TALENTO_ACTIVO = 5
COL_ESTADO_SOW     = 6
COL_SOW            = 7
COL_ANIO           = 8
COL_FECHA_INICIO   = 9
COL_FECHA_FIN      = 10
COL_ALERTA         = 11
COL_PERIODO        = 12
COL_DIAS           = 13
COL_COMENTARIOS    = 14
COL_COMENTARIOS_2  = 15
COL_FIRMADO        = 16
COL_CARPETA        = 17
COL_RUTA_ARCHIVO   = 18
COL_VALOR_COP      = 19
COL_VALOR_USD      = 20


# ── Helpers ────────────────────────────────────────────────────────────────

def extraer_pais_de_empresa(empresa: str) -> str:
    empresa = (empresa or "").lower()
    mapa = {
        "colombia": "COLOMBIA", "mexico": "MEXICO", "méxico": "MEXICO",
        "argentina": "ARGENTINA", "peru": "PERU", "perú": "PERU",
        "ecuador": "ECUADOR", "chile": "CHILE",
        "usa": "USA", "estados unidos": "USA",
    }
    for clave, pais in mapa.items():
        if clave in empresa:
            return pais
    return "COLOMBIA"


def extraer_anio_de_sow(numero_sow) -> int:
    if not numero_sow:
        return datetime.now().year
    m = re.search(r"(20\d{2})", str(numero_sow))
    return int(m.group(1)) if m else datetime.now().year


def _limpiar_valor_cop(valor_cop):
    if not valor_cop:
        return None
    if isinstance(valor_cop, (int, float)):
        return float(valor_cop)
    try:
        return float(str(valor_cop).replace("$", "").replace(",", "").strip())
    except Exception:
        return None


# ── ZIP quirúrgico ─────────────────────────────────────────────────────────

def _encontrar_ruta_hoja(zf: zipfile.ZipFile, nombre_hoja: str) -> str:
    """
    Lee workbook.xml y workbook.xml.rels para encontrar
    el path dentro del ZIP del XML de la hoja buscada.
    """
    NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    NS_R    = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    NS_PKG  = "http://schemas.openxmlformats.org/package/2006/relationships"

    wb_root = ET.fromstring(zf.read("xl/workbook.xml"))
    r_id = None
    for sheet in wb_root.iter(f"{{{NS_MAIN}}}sheet"):
        if sheet.get("name") == nombre_hoja:
            r_id = sheet.get(f"{{{NS_R}}}id")
            break
    if not r_id:
        raise ValueError(f"Hoja '{nombre_hoja}' no encontrada en workbook.xml")

    rels_root = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
    for rel in rels_root.iter(f"{{{NS_PKG}}}Relationship"):
        if rel.get("Id") == r_id:
            target = rel.get("Target")
            return target if target.startswith("xl/") else f"xl/{target}"

    raise ValueError(f"Relationship '{r_id}' no encontrada en workbook.xml.rels")


def _guardar_zip_quirurgico(ruta_excel: str, wb_openpyxl, nombre_hoja: str):
    """
    Guarda el workbook modificado reemplazando SOLO los archivos necesarios
    en el ZIP original. Power Query, VBA y todo lo demás queda intacto.
    """
    ruta_abs = os.path.abspath(ruta_excel)

# Extender la tabla Excel para incluir las filas nuevas
    import re as _re
    ws_ext = wb_openpyxl[nombre_hoja]
    _ultima_real = FILA_INICIO_DATOS - 1
    for _r in range(FILA_INICIO_DATOS, ws_ext.max_row + 1):
        if ws_ext.cell(row=_r, column=COL_SOW).value or \
           ws_ext.cell(row=_r, column=COL_NOMBRE).value:
            _ultima_real = _r

    for tbl in ws_ext.tables.values():
        partes = tbl.ref.split(":")
        if len(partes) == 2:
            col_ini  = _re.match(r"[A-Z]+", partes[0]).group()
            fila_ini = _re.search(r"\d+", partes[0]).group()
            col_fin  = _re.match(r"[A-Z]+", partes[1]).group()
            tbl.ref  = f"{col_ini}{fila_ini}:{col_fin}{_ultima_real}"

    # Guardar el output de openpyxl en memoria
    buf = io.BytesIO()
    wb_openpyxl.save(buf)
    buf.seek(0)

    # Determinar qué archivos reemplazar
    with zipfile.ZipFile(ruta_abs, "r") as orig_zip:
        sheet_path = _encontrar_ruta_hoja(orig_zip, nombre_hoja)

    sheet_dir  = sheet_path.rsplit("/", 1)[0]       # xl/worksheets
    sheet_file = sheet_path.rsplit("/", 1)[1]        # sheetN.xml
    sheet_rels = f"{sheet_dir}/_rels/{sheet_file}.rels"

    # Extender rangos de formato condicional automáticamente.
    # Convierte $E$8:$E$230 → $E$8:$E$9999 para que las filas nuevas
    # tengan los mismos colores automáticos que las anteriores.
    # Se ejecuta en cada guardado — sin tocar nada en Excel manualmente.
    try:
        with zipfile.ZipFile(buf, 'r') as _zf_cf:
            _xml_hoja = _zf_cf.read(sheet_path).decode('utf-8')
        buf.seek(0)

        _xml_hoja = _re.sub(
            r'(sqref="[^"]*?:\$?[A-Z]+\$?)(\d+)(")',
            lambda m: (f'{m.group(1)}{_ultima_real}{m.group(3)}'
                       if int(m.group(2)) < _ultima_real else m.group(0)),
            _xml_hoja
        )

        _buf_cf = io.BytesIO()
        with zipfile.ZipFile(buf, 'r') as _zf_src:
            with zipfile.ZipFile(_buf_cf, 'w', zipfile.ZIP_DEFLATED) as _zf_dst:
                for _item in _zf_src.namelist():
                    _zf_dst.writestr(
                        _item,
                        _xml_hoja.encode('utf-8') if _item == sheet_path
                        else _zf_src.read(_item)
                    )
        _buf_cf.seek(0)
        buf = _buf_cf
    except Exception:
        buf.seek(0)

    REEMPLAZAR = {
        sheet_path,
        sheet_rels,
        "xl/sharedStrings.xml",
        "xl/calcChain.xml",
        "xl/styles.xml",
    }

    # Agregar archivos de tablas Excel (xl/tables/table1.xml, etc.)
    with zipfile.ZipFile(buf, "r") as _zf:
        for name in _zf.namelist():
            if _re.match(r"xl/tables/table\d+\.xml$", name):
                REEMPLAZAR.add(name)
    buf.seek(0)   # ← resetear después de leer

    temp_path = ruta_abs + ".writing"
    try:
        with zipfile.ZipFile(ruta_abs, "r")  as orig_zip, \
             zipfile.ZipFile(buf,       "r")  as new_zip,  \
             zipfile.ZipFile(temp_path, "w",
                             compression=zipfile.ZIP_DEFLATED) as out_zip:

            orig_names = set(orig_zip.namelist())
            new_names  = set(new_zip.namelist())

            # Todos los archivos del original
            for info in orig_zip.infolist():
                name = info.filename
                if name in REEMPLAZAR and name in new_names:
                    # Reemplazar con la versión de openpyxl (datos actualizados)
                    out_zip.writestr(info, new_zip.read(name))
                else:
                    # Conservar del original (Power Query, VBA, slicers…)
                    out_zip.writestr(info, orig_zip.read(name))

            # Archivos nuevos de openpyxl que no existían (ej: sheet_rels nuevo)
            for name in new_names:
                if name not in orig_names and name in REEMPLAZAR:
                    out_zip.writestr(name, new_zip.read(name))

        os.replace(temp_path, ruta_abs)

    except Exception:
        if os.path.exists(temp_path):
            try:
                os.remove(temp_path)
            except Exception:
                pass
        raise


# ── Operaciones sobre worksheet (openpyxl) ────────────────────────────────

def _encontrar_ultima_fila_ox(ws) -> int:
    ultima = FILA_INICIO_DATOS - 1
    vacias = 0
    for fila in range(FILA_INICIO_DATOS, FILA_INICIO_DATOS + 5000):
        val_pais   = ws.cell(row=fila, column=COL_PAIS).value
        val_nombre = ws.cell(row=fila, column=COL_NOMBRE).value
        if val_pais or val_nombre:
            ultima = fila
            vacias = 0
        else:
            vacias += 1
            if vacias >= 15:
                break
    return ultima


def _sow_ya_existe_ox(ws, sow: str, nombre: str, ultima_fila: int):
    sow_up    = sow.strip().upper()
    nombre_up = nombre.strip().upper()
    for fila in range(FILA_INICIO_DATOS, ultima_fila + 1):
        if (str(ws.cell(row=fila, column=COL_SOW).value    or "").strip().upper() == sow_up
                and str(ws.cell(row=fila, column=COL_NOMBRE).value or "").strip().upper() == nombre_up):
            return True, fila
    return False, None


def _buscar_employee_id_ox(ws, nombre_persona: str) -> str:
    """Busca employee ID en la hoja Query (si existe)."""
    nombre_buscar = str(nombre_persona or "").strip().upper()
    try:
        wb = ws.parent
        if "Query" not in wb.sheetnames:
            return ""
        ws_q = wb["Query"]
        for fila in range(2, ws_q.max_row + 1):
            nombre_q = str(ws_q.cell(row=fila, column=2).value or "").strip().upper()
            if not nombre_q:
                continue
            if nombre_q == nombre_buscar or nombre_q.replace(" ", "") == nombre_buscar.replace(" ", ""):
                emp_id = ws_q.cell(row=fila, column=3).value
                if emp_id:
                    return str(int(emp_id)) if isinstance(emp_id, float) else str(emp_id).strip()
    except Exception:
        pass
    return ""


def _copiar_estilo_fila_ox(ws, fila_origen: int, fila_destino: int):
    from openpyxl.styles import Border, Side
    thin = Side(style="thin", color="000000")
    borde_fino = Border(left=thin, right=thin, top=thin, bottom=thin)

    for col in range(1,24):
        try:
            c_orig = ws.cell(row=fila_origen,  column=col)
            c_dest = ws.cell(row=fila_destino, column=col)
            c_dest.number_format = c_orig.number_format
            if c_orig.has_style:
                c_dest.font      = copy(c_orig.font)
                # K (11) y L (12) las colorea la macro GenerarAlertasSOW — no copiar fill
                if col not in (11, 12):
                    c_dest.fill  = copy(c_orig.fill)
                c_dest.alignment = copy(c_orig.alignment)
            c_dest.border = borde_fino  # borde explícito siempre
        except Exception:
            pass

    try:
        h = ws.row_dimensions[fila_origen].height
        if h:
            ws.row_dimensions[fila_destino].height = h
    except Exception:
        pass

def _encontrar_fila_template_ox(ws, desde_fila: int) -> int:
    """Busca hacia atrás la última fila con fórmulas en E y M para usar como template."""
    for fila in range(desde_fila, FILA_INICIO_DATOS - 1, -1):
        val_e = ws[f"E{fila}"].value
        val_m = ws[f"M{fila}"].value
        tiene_e = (val_e is not None and (
            (type(val_e).__name__ == "ArrayFormula" and hasattr(val_e, "text")) or
            (isinstance(val_e, str) and val_e.startswith("="))
        ))
        tiene_m = (val_m is not None and (
            (type(val_m).__name__ == "ArrayFormula" and hasattr(val_m, "text")) or
            (isinstance(val_m, str) and val_m.startswith("="))
        ))
        if tiene_e and tiene_m:
            return fila
    return FILA_INICIO_DATOS


def _escribir_fila_ox(ws, nueva_fila: int, datos: dict, fila_template: int):
    from openpyxl.styles import Font
    from openpyxl.utils import get_column_letter
    from datetime import datetime

    def _es_array_formula(val):
        return type(val).__name__ == "ArrayFormula" and hasattr(val, "text")

    _copiar_estilo_fila_ox(ws, fila_template, nueva_fila)
    ws.row_dimensions[nueva_fila].height = ws.row_dimensions[fila_template].height

    pais    = (datos.get("pais") or extraer_pais_de_empresa(datos.get("empresa", ""))).upper()
    firmado = datos.get("firmado", False)
    anio    = extraer_anio_de_sow(datos.get("sow"))

    ws.cell(row=nueva_fila, column=COL_PAIS).value     = pais
    ws.cell(row=nueva_fila, column=COL_PROYECTO).value = datos.get("proyecto") or PROYECTO_NEORIS
    ws.cell(row=nueva_fila, column=COL_NOMBRE).value   = str(datos.get("nombre") or "").upper()

    _emp_id = (datos.get("employee_id") or
               _buscar_employee_id_ox(ws, str(datos.get("nombre") or "")))
    try:
        ws.cell(row=nueva_fila, column=COL_EMPLOYEE_ID).value = int(_emp_id) if _emp_id else ""
    except (ValueError, TypeError):
        ws.cell(row=nueva_fila, column=COL_EMPLOYEE_ID).value = _emp_id or ""

    ws.cell(row=nueva_fila, column=COL_ESTADO_SOW).value = ""
    ws.cell(row=nueva_fila, column=COL_SOW).value        = datos.get("sow") or ""
    ws.cell(row=nueva_fila, column=COL_ANIO).value       = anio

    fecha_inicio = datos.get("fecha_inicio")
    if fecha_inicio:
        c               = ws.cell(row=nueva_fila, column=COL_FECHA_INICIO)
        c.value         = fecha_inicio
        c.number_format = "DD/MM/YYYY"

    fecha_fin = datos.get("fecha_fin")
    if fecha_fin:
        c               = ws.cell(row=nueva_fila, column=COL_FECHA_FIN)
        c.value         = fecha_fin
        c.number_format = "DD/MM/YYYY"

    # ── Fórmulas E, K, L, M, T — copiar del template ajustando fila ──
    # BUSCARX en Excel 365 se guarda como ArrayFormula en openpyxl,
    # no como string — hay que extraer el texto con .text
    delta = nueva_fila - fila_template
    for col_letra in ("E", "K", "L", "M", "T"):
            c_orig = ws[f"{col_letra}{fila_template}"]
            val    = c_orig.value

            if val is None:
                continue

            if _es_array_formula(val):
                formula_str   = val.text
                formula_nueva = re.sub(
                    r"(?<!\$)([A-Z]+)(\d+)",
                    lambda m: f"{m.group(1)}{int(m.group(2)) + delta}",
                    formula_str,
                )
                # Usar type(val) recrea ArrayFormula sin import explícito
                # — preserva el formato {=...} original y evita el warning
                try:
                    _AF_cls  = type(val)
                    cell_ref = f"{col_letra}{nueva_fila}:{col_letra}{nueva_fila}"
                    ws[f"{col_letra}{nueva_fila}"] = _AF_cls(cell_ref, formula_nueva)
                except Exception:
                    ws[f"{col_letra}{nueva_fila}"] = formula_nueva

            elif isinstance(val, str) and val.startswith("="):
                formula_nueva = re.sub(
                    r"(?<!\$)([A-Z]+)(\d+)",
                    lambda m: f"{m.group(1)}{int(m.group(2)) + delta}",
                    val,
                )
                ws[f"{col_letra}{nueva_fila}"] = formula_nueva

    # ── Color columna M según días restantes ──────────────────────────

    ws.cell(row=nueva_fila, column=COL_COMENTARIOS).value   = ""
    ws.cell(row=nueva_fila, column=COL_COMENTARIOS_2).value = ""
    ws.cell(row=nueva_fila, column=COL_FIRMADO).value       = "Si" if firmado else ""

    # Carpeta (col Q) — hipervínculo
    nombre_persona = str(datos.get("nombre") or "").title()
    ruta_carpeta   = datos.get("ruta_carpeta") or ""
    celda_q        = ws.cell(row=nueva_fila, column=COL_CARPETA)
    if ruta_carpeta:
        celda_q.value     = nombre_persona
        celda_q.hyperlink = ruta_carpeta
        celda_q.font      = Font(color="0000FF", underline="single")
    else:
        celda_q.value = nombre_persona

    # Ruta archivo (col R) — hipervínculo si firmado
    ruta       = datos.get("ruta_archivo") or ""
    sow_numero = str(datos.get("sow") or "")
    celda_r    = ws.cell(row=nueva_fila, column=COL_RUTA_ARCHIVO)
    if ruta and firmado:
        celda_r.value     = sow_numero
        celda_r.hyperlink = ruta
        celda_r.font      = Font(color="0000FF", underline="single")
    else:
        celda_r.value = ""

    # Valor COP (col S)
    valor_cop = _limpiar_valor_cop(datos.get("valor_cop"))
    celda_s   = ws.cell(row=nueva_fila, column=COL_VALOR_COP)
    if valor_cop is not None:
        celda_s.value         = valor_cop
        celda_s.number_format = "$#,##0.00"
    else:
        celda_s.value = ""

    return nueva_fila


# ── API pública ────────────────────────────────────────────────────────────

def escribir_sows_en_excel(lista_sows: list, ruta_excel: str = None) -> list:
    from openpyxl import load_workbook

    if ruta_excel is None:
        ruta_excel = RUTA_EXCEL_PRUEBA
    if not os.path.exists(ruta_excel):
        raise FileNotFoundError(f"No se encontró el Excel: {ruta_excel}")

    wb = load_workbook(ruta_excel, keep_vba=True)
    ws = wb[NOMBRE_HOJA]

    ultima_fila = _encontrar_ultima_fila_ox(ws)
    resultados  = []

    for datos in lista_sows:
        nombre = str(datos.get("nombre") or "").strip()
        sow    = str(datos.get("sow")    or "").strip()

        if not nombre or not sow:
            resultados.append({"nombre": nombre, "sow": sow,
                "accion": "error", "mensaje": "Nombre o SOW vacío", "fila": None})
            continue

        existe, fila_existente = _sow_ya_existe_ox(ws, sow, nombre, ultima_fila)
        if existe:
            resultados.append({"nombre": nombre, "sow": sow,
                "accion": "ya_existe",
                "mensaje": f"Ya existe en fila {fila_existente}",
                "fila": fila_existente})
            continue

        nueva_fila     = ultima_fila + 1
        fila_template  = _encontrar_fila_template_ox(ws, ultima_fila)
        _escribir_fila_ox(ws, nueva_fila, datos, fila_template)
        ultima_fila = nueva_fila
        resultados.append({"nombre": nombre, "sow": sow,
            "accion": "agregado",
            "mensaje": f"Agregado en fila {nueva_fila}",
            "fila": nueva_fila})

    _guardar_zip_quirurgico(ruta_excel, wb, NOMBRE_HOJA)
    wb.close()
    return resultados


def escribir_actualizaciones_excel(lista_actualizaciones: list, ruta_excel: str = None) -> list:
    from openpyxl import load_workbook
    from openpyxl.styles import Font

    if ruta_excel is None:
        ruta_excel = RUTA_EXCEL_PRUEBA
    if not os.path.exists(ruta_excel):
        raise FileNotFoundError(f"No se encontró el Excel: {ruta_excel}")

    wb = load_workbook(ruta_excel, keep_vba=True)
    ws = wb[NOMBRE_HOJA]

    ultima_fila = _encontrar_ultima_fila_ox(ws)
    resultados  = []

    for datos in lista_actualizaciones:
        sow    = str(datos.get("sow")    or "").strip()
        nombre = str(datos.get("nombre") or "").strip()

        if not sow or not nombre:
            resultados.append({"nombre": nombre, "sow": sow,
                "accion": "error", "mensaje": "Nombre o SOW vacío", "fila": None})
            continue

        encontrado = False
        for fila in range(FILA_INICIO_DATOS, ultima_fila + 1):
            sow_c    = str(ws.cell(row=fila, column=COL_SOW).value    or "").strip().upper()
            nombre_c = str(ws.cell(row=fila, column=COL_NOMBRE).value or "").strip().upper()

            if sow_c == sow.upper() and nombre_c == nombre.upper():
                ws.cell(row=fila, column=COL_FIRMADO).value = "Si"
                ruta    = datos.get("ruta_archivo") or ""
                celda_r = ws.cell(row=fila, column=COL_RUTA_ARCHIVO)
                if ruta:
                    celda_r.value     = sow.upper()
                    celda_r.hyperlink = ruta
                    celda_r.font      = Font(color="0000FF", underline="single")
                resultados.append({"nombre": nombre, "sow": sow,
                    "accion": "firmado",
                    "mensaje": f"Firmado en fila {fila}",
                    "fila": fila})
                encontrado = True
                break

        if not encontrado:
            resultados.append({"nombre": nombre, "sow": sow,
                "accion": "no_encontrado",
                "mensaje": "No se encontró en el Excel",
                "fila": None})

    _guardar_zip_quirurgico(ruta_excel, wb, NOMBRE_HOJA)
    wb.close()
    return resultados


def actualizar_fila_existente(numero_sow: str, nombre_persona: str,
                              datos: dict, ruta_excel: str = None) -> tuple:
    return escribir_actualizaciones_excel(
        [{"sow": numero_sow, "nombre": nombre_persona, **datos}],
        ruta_excel=ruta_excel,
    )[0], None


def leer_sows_del_excel(ruta_excel: str = None) -> list:
    from openpyxl import load_workbook

    if ruta_excel is None:
        ruta_excel = RUTA_EXCEL_PRUEBA
    if not os.path.exists(ruta_excel):
        raise FileNotFoundError(f"No se encontró el Excel: {ruta_excel}")

    wb = load_workbook(ruta_excel, keep_vba=True, data_only=True)

    if NOMBRE_HOJA not in wb.sheetnames:
        raise ValueError(f"No se encontró la hoja '{NOMBRE_HOJA}'.")

    ws    = wb[NOMBRE_HOJA]
    filas = []

    for fila in range(FILA_INICIO_DATOS, ws.max_row + 1):
        sow    = str(ws.cell(row=fila, column=COL_SOW).value    or "").strip()
        nombre = str(ws.cell(row=fila, column=COL_NOMBRE).value or "").strip()

        if not sow:
            continue

        # celda_r            = ws.cell(row=fila, column=COL_RUTA_ARCHIVO)
        # tiene_hipervinculo = bool(celda_r.hyperlink and
        #                           str(celda_r.hyperlink.target or "").strip())
        # tiene_valor        = bool(celda_r.value and str(celda_r.value).strip())
        # if tiene_hipervinculo or tiene_valor:
        #     continue
        celda_r     = ws.cell(row=fila, column=COL_RUTA_ARCHIVO)
        # read_only=True no expone .hyperlink, pero cuando hay hipervínculo
        # siempre hay valor (SOW number como texto). Chequear valor es suficiente.
        tiene_valor = bool(celda_r.value and str(celda_r.value).strip())
        if tiene_valor:
            continue

        fecha_inicio = ws.cell(row=fila, column=COL_FECHA_INICIO).value
        fecha_fin    = ws.cell(row=fila, column=COL_FECHA_FIN).value

        valor_cop = ws.cell(row=fila, column=COL_VALOR_COP).value
        try:
            valor_cop = float(valor_cop) if valor_cop else None
        except (ValueError, TypeError):
            valor_cop = None

        celda_q      = ws.cell(row=fila, column=COL_CARPETA)
        ruta_carpeta = ""
        if celda_q.hyperlink:
            ruta_carpeta = str(celda_q.hyperlink.target or "").strip()
        elif celda_q.value:
            ruta_carpeta = str(celda_q.value).strip()

        filas.append({
            "sow":          sow,
            "nombre":       nombre,
            "empresa":      str(ws.cell(row=fila, column=COL_PROYECTO).value or "").strip(),
            "fecha_inicio": fecha_inicio,
            "fecha_fin":    fecha_fin,
            "valor_cop":    valor_cop,
            "firmado":      False,
            "ruta_archivo": "",
            "ruta_carpeta": ruta_carpeta,
            "estado":       "pendiente",
            "tipo":         "pendiente",
            "remitente":    "Gema Vale",
            "fecha_email":  None,
        })

    wb.close()
    return filas

def leer_indice_excel(ruta_excel: str = None) -> dict:
    """
    Devuelve TODAS las filas del Excel (cerradas y pendientes):
      { (sow_upper, nombre_norm): "cerrado" | "pendiente" }
    "cerrado"   = columna R tiene valor (ya firmado)
    "pendiente" = columna R vacía
    """
    if ruta_excel is None:
        ruta_excel = RUTA_EXCEL_PRUEBA
    if not os.path.exists(ruta_excel):
        return {}

    from openpyxl import load_workbook
    wb = load_workbook(ruta_excel, keep_vba=True, data_only=True)

    if NOMBRE_HOJA not in wb.sheetnames:
        wb.close()
        return {}

    ws     = wb[NOMBRE_HOJA]
    indice = {}

    for fila in range(FILA_INICIO_DATOS, ws.max_row + 1):
        sow = str(ws.cell(row=fila, column=COL_SOW).value or "").strip().upper()
        if not sow:
            continue
        nombre = str(ws.cell(row=fila, column=COL_NOMBRE).value or "").strip().upper()
        nombre = " ".join(nombre.split())  # normalizar espacios

        celda_r     = ws.cell(row=fila, column=COL_RUTA_ARCHIVO)
        tiene_valor = bool(celda_r.value and str(celda_r.value).strip())
        indice[(sow, nombre)] = "cerrado" if tiene_valor else "pendiente"

    wb.close()
    return indice