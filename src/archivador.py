"""
Módulo para archivar PDFs de SOWs en una estructura organizada:
  [RAIZ]/
    [Nombre Persona]/
      [Año]/
        [archivo.pdf]

Si un mismo PDF corresponde a varias personas (SOW compartido),
se guarda una copia en la carpeta de cada persona.
"""

import os
import re


# Ruta raíz donde se guardan los SOWs procesados.
# Cuando pasemos a OneDrive, solo cambiamos esta línea.


# RUTA_RAIZ_DEFAULT = os.path.join(
#     os.path.expanduser("~"),
#     "Desktop",
#     "SOWs Procesados",
# )

RUTA_RAIZ_DEFAULT = os.path.join(
    os.path.expanduser("~"), "EPAM",
    "92101343 - CHUBB COLOMBIA - 1 DEV FT - Documents",
    "0. Documentacion Contractual",
    "Contratos - SOW")

def limpiar_nombre_archivo(nombre):
    """
    Limpia el nombre del archivo:
      - Quita '.docx.pdf' o '.doc.pdf' redundantes → deja solo '.pdf'.
      - Mantiene todo lo demás (sufijos importantes como _CL, _CR01, etc).
    """
    nombre_lower = nombre.lower()
    if nombre_lower.endswith(".docx.pdf"):
        return nombre[:-9] + ".pdf"
    if nombre_lower.endswith(".doc.pdf"):
        return nombre[:-8] + ".pdf"
    return nombre


def limpiar_nombre_carpeta(nombre):
    """
    Limpia un nombre para que sea válido como carpeta en Windows.
    Quita caracteres no permitidos: \\ / : * ? " < > |
    """
    if not nombre:
        return "Sin Nombre"
    nombre = str(nombre).strip()
    # Caracteres prohibidos en nombres de carpeta de Windows
    prohibidos = r'[\\/:*?"<>|]'
    nombre = re.sub(prohibidos, "", nombre)
    # Reducir espacios múltiples
    nombre = re.sub(r"\s+", " ", nombre)
    return nombre.strip() or "Sin Nombre"


def extraer_anio_de_sow(numero_sow):
    """
    Saca el año del número de SOW.
    Ej: 'NEO-SERV-2026-0182' → 2026
        'NEOCOL-SERV-2025-0715' → 2025
    Si no se encuentra un año válido, devuelve None.
    """
    if not numero_sow:
        return None
    match = re.search(r"(20\d{2})", numero_sow)
    if match:
        return int(match.group(1))
    return None


def construir_ruta_destino(ruta_raiz, nombre_persona, anio):
    """Construye la ruta: ruta_raiz/[Persona]/[Año]/"""
    persona_limpia = limpiar_nombre_carpeta(nombre_persona)
    anio_str = str(anio) if anio else "Sin Año"
    ruta = os.path.join(ruta_raiz, persona_limpia, anio_str)
    os.makedirs(ruta, exist_ok=True)
    return ruta


def archivar_pdf(
    contenido_bytes,
    nombre_archivo,
    nombre_persona,
    anio,
    ruta_raiz=None,
):
    """
    Guarda un PDF en la estructura organizada.

    Args:
      contenido_bytes: bytes del PDF.
      nombre_archivo: nombre original del archivo.
      nombre_persona: nombre del consultor.
      anio: año del SOW.
      ruta_raiz: carpeta raíz donde guardar (opcional).

    Returns:
      dict con: ruta_guardada, accion ('guardado' o 'ya_existia'), error.
    """
    if ruta_raiz is None:
        ruta_raiz = RUTA_RAIZ_DEFAULT

    nombre_limpio = limpiar_nombre_archivo(nombre_archivo)
    carpeta_destino = construir_ruta_destino(ruta_raiz, nombre_persona, anio)
    ruta_final = os.path.join(carpeta_destino, nombre_limpio)

    # Si el archivo ya existe, no lo sobrescribimos
    if os.path.exists(ruta_final):
        return {
            "ruta_guardada": ruta_final,
            "accion": "ya_existia",
            "error": None,
        }

    try:
        with open(ruta_final, "wb") as f:
            f.write(contenido_bytes)
        return {
            "ruta_guardada": ruta_final,
            "accion": "guardado",
            "error": None,
        }
    except Exception as e:
        return {
            "ruta_guardada": None,
            "accion": "error",
            "error": str(e),
        }


def archivar_pdf_para_varias_personas(
    contenido_bytes,
    nombre_archivo,
    lista_personas,
    anio,
    ruta_raiz=None,
):
    """
    Guarda el mismo PDF en la carpeta de cada persona de la lista.
    Útil cuando un SOW corresponde a varios consultores.

    Returns:
      Lista de resultados (uno por persona).
    """
    resultados = []
    for persona in lista_personas:
        resultado = archivar_pdf(
            contenido_bytes,
            nombre_archivo,
            persona,
            anio,
            ruta_raiz,
        )
        resultado["persona"] = persona
        resultados.append(resultado)
    return resultados


def extraer_numero_sow_de_archivo(nombre_archivo):
    """
    Extrae SOLO el número de SOW base del nombre del archivo,
    sin sufijos como _CL, _Chubb Matrix Data, etc.

    Ej:
      'NEO-SERV-2026-0490-CR01.docx.pdf'                    → 'NEO-SERV-2026-0490-CR01'
      'NEO-SERV-2026-0490-CR01_CL.docx.pdf'                 → 'NEO-SERV-2026-0490-CR01'
      'NEO-SERV-2026-0490-CR01_Chubb Matrix Data.docx.pdf'  → 'NEO-SERV-2026-0490-CR01'
      'NEOCOL-SERV-2026-0715.docx.pdf'                      → 'NEOCOL-SERV-2026-0715'
    """
    # Patrón: NEO[país opcional]-SERV-YYYY-NNNN[-CRxx opcional]
    # Capturamos solo hasta el final del SOW base, antes de cualquier _ o espacio.
    match = re.search(
        r"(NEO[A-Z]*-SERV-\d{4}-\d{4}(?:-CR\d+)?)",
        nombre_archivo,
        re.IGNORECASE,
    )
    if match:
        return match.group(1).upper()
    # Fallback: nombre sin extensiones
    return os.path.splitext(os.path.splitext(nombre_archivo)[0])[0]


# Prueba si lo ejecutas directamente
if __name__ == "__main__":
    print("=" * 60)
    print("PRUEBA DEL ARCHIVADOR")
    print("=" * 60)
    print(f"Ruta raíz: {RUTA_RAIZ_DEFAULT}\n")

    # Caso 1: limpiar nombres
    pruebas_nombres = [
        "NEO-SERV-2026-0490-CR01.docx.pdf",
        "NEO-SERV-2026-0490-CR01_CL.docx.pdf",
        "NEO-SERV-2026-0490-CR01_Chubb Matrix Data.docx.pdf",
        "NEO-SERV-2026-0182.pdf",
    ]
    print("Limpieza de nombres:")
    for p in pruebas_nombres:
        print(f"  '{p}' → '{limpiar_nombre_archivo(p)}'")

    print()

    # Caso 2: extraer año
    pruebas_sow = [
        "NEO-SERV-2026-0182",
        "NEOCOL-SERV-2025-0715",
        "NEO-SERV-2024-0001-CR02",
    ]
    print("Extracción de año:")
    for s in pruebas_sow:
        print(f"  '{s}' → {extraer_anio_de_sow(s)}")

    print()

    # Caso 3: extraer SOW del nombre del archivo
    print("Extracción de SOW del nombre:")
    for p in pruebas_nombres:
        print(f"  '{p}' → '{extraer_numero_sow_de_archivo(p)}'")