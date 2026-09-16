"""
Flujo principal de procesamiento de SOWs.
Conecta todos los módulos:
  1. outlook_reader              → lee correos de Outlook
  2. parsear_tabla_sin_encabezados (en este mismo archivo) → extrae tabla del correo (sin encabezados)
  3. word_extractor              → extrae fechas y valor del Word
  4. archivador                  → guarda archivos organizados por persona/año
  5. excel_writer                → escribe filas en el Excel maestro
"""

import os
import re
from datetime import datetime

from src.outlook_reader import (
    conectar_outlook,
    obtener_carpeta_sows,
    listar_correos,
    obtener_adjuntos_pdf,
    marcar_como_procesado,
)
from src.archivador import (
    archivar_pdf_para_varias_personas,
    extraer_numero_sow_de_archivo,
    extraer_anio_de_sow,
    RUTA_RAIZ_DEFAULT,
)
from src.excel_writer import (
    escribir_sows_en_excel,
    escribir_actualizaciones_excel,   # ← firma nueva, no necesita wb
    RUTA_EXCEL_PRUEBA,
)


def parsear_tabla_sin_encabezados(html):
    """
    Parsea la tabla del correo de Gema Vale que NO tiene encabezados.
    Estructura: SOW | Empresa | Personas (separadas por salto de línea).
    """
    from bs4 import BeautifulSoup

    if not html:
        return []

    resultados = []
    soup       = BeautifulSoup(html, "lxml")
    tablas     = soup.find_all("table")

    for tabla in tablas:
        filas = tabla.find_all("tr")
        for fila in filas:
            celdas = fila.find_all(["td", "th"])
            if len(celdas) < 2:
                continue

            sow = celdas[0].get_text(separator=" ", strip=True).strip()
            if not sow or not re.search(
                r"NEO[A-Z]*-SERV-\d{4}-\d+", sow, re.IGNORECASE
            ):
                continue

            empresa = ""
            if len(celdas) >= 2:
                empresa = celdas[1].get_text(separator=" ", strip=True)

            personas = []
            if len(celdas) >= 3:
                texto_personas = celdas[2].get_text(separator="\n", strip=True)
                for linea in texto_personas.split("\n"):
                    linea = linea.strip()
                    if linea and linea.lower() not in ["nan", "none", ""]:
                        personas.append(linea.upper())

            if sow and personas:
                resultados.append({
                    "sow":      sow.upper(),
                    "empresa":  empresa,
                    "personas": personas,
                })

    return resultados


def construir_indice_global(correos):
    """
    Construye índice global: { "NEO-SERV-2026-0737": ["PERSONA A", ...] }
    combinando tablas de TODOS los correos.
    """
    indice          = {}
    empresa_por_sow = {}

    for correo in correos:
        filas = parsear_tabla_sin_encabezados(correo["cuerpo_html"])
        for fila in filas:
            sow     = fila["sow"]
            empresa = fila.get("empresa", "")
            empresa_por_sow[sow] = empresa
            indice.setdefault(sow, [])
            for persona in fila.get("personas", []):
                if persona and persona not in indice[sow]:
                    indice[sow].append(persona)

    return indice, empresa_por_sow


def contar_correos_nuevos():
    """Cuenta correos en la carpeta SOWs sin categoría."""
    try:
        outlook, carpeta = conectar_outlook()
        if carpeta is None:
            carpeta = obtener_carpeta_sows(outlook)
        mensajes = carpeta.Items
        mensajes.Sort("[ReceivedTime]", True)
        count = 0
        for msg in mensajes:
            if not getattr(msg, "Categories", ""):
                count += 1
        return count
    except Exception:
        return 0


def procesar_correos(
    ruta_excel=None,
    ruta_sows=None,
    nombre_carpeta_outlook="SOWs",
    solo_no_procesados=True,
    marcar_procesados=False,
    escribir_excel=False,
    callback_progreso=None,
):
    """
    Función principal que ejecuta el flujo completo.

    Args:
        ruta_excel:              ruta al Excel maestro (.xlsm)
        ruta_sows:               ruta raíz donde guardar archivos
        nombre_carpeta_outlook:  subcarpeta en Outlook
        solo_no_procesados:      omitir correos ya marcados
        marcar_procesados:       marcar correos como procesados en Outlook
        escribir_excel:          True  = escribir en Excel (punto de entrada directo)
                                 False = solo leer (modo UI)
        callback_progreso:       función(msg) para reportar progreso a la UI
    """
    def log(msg):
        if callback_progreso:
            callback_progreso(msg)
        else:
            print(msg)

    if ruta_excel is None:
        ruta_excel = RUTA_EXCEL_PRUEBA
    if ruta_sows is None:
        ruta_sows = RUTA_RAIZ_DEFAULT

    resultado_global = {
        "correos":          [],
        "total_correos":    0,
        "total_pdfs":       0,
        "total_guardados":  0,
        "total_excel":      0,
        "total_duplicados": 0,
        "total_sin_match":  0,
        "total_errores":    0,
        "errores_detalle":  [],
    }

    # ── PASO 1: Conectar a Outlook ─────────────────────────────
    log("Conectando a Outlook...")
    try:
        ns, carpeta = conectar_outlook(nombre_carpeta_outlook)
        if carpeta is None:
            carpeta = obtener_carpeta_sows(ns, nombre_carpeta_outlook)
        correos = listar_correos(carpeta, solo_no_procesados=solo_no_procesados)
        log(f"Correos encontrados: {len(correos)}")
    except Exception as e:
        log(f"Error conectando a Outlook: {e}")
        resultado_global["errores_detalle"].append(str(e))
        return resultado_global

    if not correos:
        log("No hay correos para procesar.")
        return resultado_global

    resultado_global["total_correos"] = len(correos)

    # ── PASO 2: Construir índice global SOW -> Personas ────────
    log("Construyendo indice de SOWs y personas...")
    indice_global, empresa_por_sow = construir_indice_global(correos)
    log(f"SOWs en indice: {len(indice_global)}")
    for sow, personas in list(indice_global.items())[:8]:
        personas_str = ", ".join(personas) if isinstance(personas, list) else str(personas)
        log(f"  {sow} -> {personas_str}")
    if len(indice_global) > 8:
        log(f"  ... y {len(indice_global) - 8} mas")

    # ── PASO 3: Procesar archivos adjuntos de cada correo ──────
    log("\nProcesando archivos adjuntos...")

    for i, correo in enumerate(correos, 1):
        log(f"\n[{i}/{len(correos)}] {correo['asunto'][:70]}")

        resultado_correo = {
            "asunto":           correo["asunto"],
            "remitente":        correo["remitente"],
            "remitente_nombre": correo["remitente_nombre"],
            "fecha":            correo["fecha"],
            "sows_tabla":       [],
            "archivos":         [],
            "estado":           "ok",
        }

        filas_tabla = parsear_tabla_sin_encabezados(correo["cuerpo_html"])
        resultado_correo["sows_tabla"] = filas_tabla

        if not filas_tabla:
            log("  Sin tabla de SOWs — puede ser respuesta con PDF")

        if correo["num_adjuntos"] == 0:
            resultado_global["correos"].append(resultado_correo)
            continue

        adjuntos = obtener_adjuntos_pdf(correo)

        for adj in adjuntos:
            nombre_archivo = adj["nombre"]
            contenido      = adj["contenido"]

            if not nombre_archivo.lower().endswith((".pdf", ".docx", ".doc")):
                continue

            resultado_global["total_pdfs"] += 1

            es_pdf  = nombre_archivo.lower().endswith(".pdf")
            es_word = nombre_archivo.lower().endswith((".docx", ".doc"))

            sow_archivo = extraer_numero_sow_de_archivo(nombre_archivo).upper()
            anio        = extraer_anio_de_sow(sow_archivo)
            personas    = indice_global.get(sow_archivo, [])
            empresa     = empresa_por_sow.get(sow_archivo, "")

            # ── PASO 3A: Extraer fechas y valor del Word ───────
            datos_word = {"fecha_inicio": None, "fecha_fin": None, "valor_cop": None}
            if es_word:
                try:
                    from src.word_extractor import extraer_datos_word
                    datos_word = extraer_datos_word(contenido, sow_archivo)
                    if datos_word.get("fecha_inicio"):
                        fi_str = datos_word["fecha_inicio"].strftime("%d/%m/%Y")
                        ff_str = (datos_word["fecha_fin"].strftime("%d/%m/%Y")
                                  if datos_word.get("fecha_fin") else "?")
                        log(f"  Fechas {sow_archivo}: {fi_str} -> {ff_str}")
                    if datos_word.get("valor_cop"):
                        log(f"  Valor {sow_archivo}: COP ${datos_word['valor_cop']:,.2f}")
                except Exception as e:
                    log(f"  Advertencia leyendo Word {nombre_archivo}: {e}")

            # ── PASO 3B: Construir resultado_archivo ───────────
            resultado_archivo = {
                "nombre":       nombre_archivo,
                "sow":          sow_archivo,
                "personas":     personas,
                "empresa":      empresa,
                "guardados":    [],
                "estado":       "ok",
                "es_pdf":       es_pdf,
                "es_word":      es_word,
                "firmado":      es_pdf,
                "fecha_inicio": datos_word.get("fecha_inicio"),
                "fecha_fin":    datos_word.get("fecha_fin"),
                "valor_cop":    datos_word.get("valor_cop"),
                "contenido":    contenido if es_word else None,
            }

            if not personas:
                log(f"  Sin match: {nombre_archivo} -> {sow_archivo}")
                resultado_archivo["estado"] = "sin_match"
                resultado_global["total_sin_match"] += 1
                resultado_correo["estado"] = "warning"
                resultado_correo["archivos"].append(resultado_archivo)
                continue

            # ── PASO 3C: Archivar en disco ─────────────────────
            resultados_archivado = archivar_pdf_para_varias_personas(
                contenido, nombre_archivo, personas, anio, ruta_sows
            )

            for r in resultados_archivado:
                if r["accion"] == "guardado":
                    resultado_global["total_guardados"] += 1
                    log(f"  Guardado ({'PDF' if es_pdf else 'Word'}): "
                        f"{nombre_archivo} -> {r['persona']}")
                elif r["accion"] == "ya_existia":
                    log(f"  Ya existia: {nombre_archivo} -> {r['persona']}")
                elif r["accion"] == "error":
                    resultado_global["total_errores"] += 1
                    resultado_correo["estado"] = "error"
                    resultado_global["errores_detalle"].append(
                        f"{nombre_archivo}: {r['error']}")
                resultado_archivo["guardados"].append(r)

            resultado_correo["archivos"].append(resultado_archivo)

        # ── PASO 4: Escribir en Excel ──────────────────────────
        # Solo cuando escribir_excel=True (punto de entrada directo).
        # La UI siempre usa escribir_excel=False y escribe ella misma.
        if escribir_excel and ruta_excel and os.path.exists(ruta_excel):

            if filas_tabla:
                # ── Correo CON tabla ───────────────────────────
                log("  Escribiendo en Excel...")
                sows_para_excel  = []
                sows_para_update = []

                for fila in filas_tabla:
                    sow          = fila["sow"]
                    empresa      = fila.get("empresa", "")
                    personas_sow = fila.get("personas", [])

                    fecha_inicio = None
                    fecha_fin    = None
                    valor_cop    = None
                    ruta_archivo = ""
                    firmado      = False

                    for arch in resultado_correo["archivos"]:
                        if arch.get("sow") == sow:
                            fecha_inicio = arch.get("fecha_inicio") or fecha_inicio
                            fecha_fin    = arch.get("fecha_fin")    or fecha_fin
                            valor_cop    = arch.get("valor_cop")    or valor_cop
                            if arch.get("es_pdf") and arch.get("guardados"):
                                ruta_archivo = arch["guardados"][0].get("ruta_guardada", "")
                                firmado = True
                            break

                    for persona in personas_sow:
                        if not (isinstance(persona, str) and persona.strip()):
                            continue
                        entrada = {
                            "nombre":       persona.strip(),
                            "employee_id":  "",
                            "sow":          sow,
                            "empresa":      empresa,
                            "fecha_inicio": fecha_inicio,
                            "fecha_fin":    fecha_fin,
                            "ruta_archivo": ruta_archivo,
                            "ruta_carpeta": os.path.join(
                                ruta_sows, persona.title(),
                                str(extraer_anio_de_sow(sow))),
                            "valor_cop":    valor_cop,
                            "firmado":      firmado,
                        }
                        if firmado:
                            sows_para_update.append(entrada)
                        else:
                            sows_para_excel.append(entrada)

                # Words nuevos
                if sows_para_excel:
                    try:
                        res = escribir_sows_en_excel(sows_para_excel, ruta_excel=ruta_excel)
                        for r in res:
                            if r["accion"] == "agregado":
                                resultado_global["total_excel"] += 1
                                log(f"  Excel nuevo: {r['nombre']} -> fila {r['fila']}")
                            elif r["accion"] == "ya_existe":
                                resultado_global["total_duplicados"] += 1
                                log(f"  Excel duplicado: {r['nombre']}")
                            elif r["accion"] == "error":
                                log(f"  Excel error: {r['nombre']} -> {r['mensaje']}")
                    except Exception as e:
                        log(f"  Error escribiendo en Excel: {e}")
                        resultado_global["errores_detalle"].append(f"Excel: {e}")

                # PDFs firmados
                if sows_para_update:
                    try:
                        res_upd      = escribir_actualizaciones_excel(
                            sows_para_update, ruta_excel=ruta_excel)
                        no_encontrados = []
                        for idx, r in enumerate(res_upd):
                            if r["accion"] == "firmado":
                                resultado_global["total_excel"] += 1
                                log(f"  Excel actualizado (PDF): {r['nombre']} -> fila {r['fila']}")
                            elif r["accion"] == "no_encontrado":
                                no_encontrados.append(sows_para_update[idx])
                        if no_encontrados:
                            res2 = escribir_sows_en_excel(no_encontrados, ruta_excel=ruta_excel)
                            for r2 in res2:
                                if r2["accion"] == "agregado":
                                    resultado_global["total_excel"] += 1
                                    log(f"  Excel nuevo (PDF): {r2['nombre']} -> fila {r2['fila']}")
                    except Exception as e:
                        log(f"  Error actualizando Excel con PDF: {e}")
                        resultado_global["errores_detalle"].append(f"Excel PDF: {e}")

            else:
                # ── Correo SIN tabla: solo PDFs firmados ──────
                pdfs_sin_tabla = [
                    arch for arch in resultado_correo["archivos"]
                    if arch.get("es_pdf") and arch.get("guardados")
                ]

                if pdfs_sin_tabla:
                    log("  Correo sin tabla — actualizando PDFs firmados...")
                    entradas_update = []
                    for arch in pdfs_sin_tabla:
                        sow      = arch.get("sow", "")
                        personas = arch.get("personas", [])
                        ruta_pdf = arch["guardados"][0].get("ruta_guardada", "")
                        if not personas:
                            log(f"  Sin personas en indice para {sow} — omitido")
                            continue
                        for persona in personas:
                            entradas_update.append({
                                "nombre":       persona,
                                "sow":          sow,
                                "ruta_archivo": ruta_pdf,
                                "firmado":      True,
                            })

                    if entradas_update:
                        try:
                            res_upd = escribir_actualizaciones_excel(
                                entradas_update, ruta_excel=ruta_excel)
                            for r in res_upd:
                                if r["accion"] == "firmado":
                                    resultado_global["total_excel"] += 1
                                    log(f"  Firmado: {r['nombre']} -> fila {r['fila']}")
                                elif r["accion"] == "no_encontrado":
                                    log(f"  No encontrado: {r['nombre']} | {r['sow']}"
                                        f" — puede no haberse enviado aun")
                        except Exception as e:
                            log(f"  Error actualizando PDF sin tabla: {e}")
                            resultado_global["errores_detalle"].append(f"Excel PDF sin tabla: {e}")
                else:
                    log("  Sin tabla y sin PDFs — correo omitido")

        # ── Marcar correo como procesado ───────────────────────
        if marcar_procesados:
            try:
                marcar_como_procesado(correo)
            except Exception as e:
                log(f"  No se pudo marcar el correo: {e}")

        resultado_global["correos"].append(resultado_correo)

    # ── RESUMEN FINAL ──────────────────────────────────────────
    log("\n" + "=" * 60)
    log("RESUMEN")
    log("=" * 60)
    log(f"Correos procesados:    {resultado_global['total_correos']}")
    log(f"Archivos procesados:   {resultado_global['total_pdfs']}")
    log(f"Archivos guardados:    {resultado_global['total_guardados']}")
    log(f"Filas en Excel:        {resultado_global['total_excel']}")
    log(f"Duplicados omitidos:   {resultado_global['total_duplicados']}")
    log(f"Sin match en indice:   {resultado_global['total_sin_match']}")
    log(f"Errores:               {resultado_global['total_errores']}")

    return resultado_global


# ── Punto de entrada directo ───────────────────────────────────
if __name__ == "__main__":
    procesar_correos(
        ruta_excel=RUTA_EXCEL_PRUEBA,
        ruta_sows=RUTA_RAIZ_DEFAULT,
        solo_no_procesados=False,
        marcar_procesados=False,
        escribir_excel=True,
    )