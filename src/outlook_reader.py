"""
Módulo para leer correos de Outlook usando COM (win32com).
Lee la subcarpeta 'SOWs' dentro de la Bandeja de Entrada.
Si la carpeta no existe, la crea automáticamente.
"""

import time
import pythoncom
import win32com.client
import os
import tempfile


def _obtener_o_crear_carpeta(inbox, namespace, nombre_carpeta):
    """
    Busca la subcarpeta dentro de inbox.
    Si no existe, la crea automáticamente en la Bandeja de Entrada.
    Devuelve siempre (entry_id, store_id) — nunca None.
    """
    # Buscar (case-insensitive)
    for c in inbox.Folders:
        if c.Name.lower() == nombre_carpeta.lower():
            return c.EntryID, c.StoreID

    # No existe → crearla
    print(f"[SOW Manager] Carpeta '{nombre_carpeta}' no encontrada. Creándola automáticamente...")
    nueva = inbox.Folders.Add(nombre_carpeta)
    print(f"[SOW Manager] Carpeta '{nombre_carpeta}' creada en Bandeja de Entrada.")
    return nueva.EntryID, nueva.StoreID


def conectar_outlook(nombre_carpeta="SOWs"):
    """
    Conecta a Outlook, fuerza sincronización y devuelve la carpeta SOWs
    obtenida por EntryID (no por navegación de Folders).

    Si la carpeta no existe la crea automáticamente dentro de
    la Bandeja de Entrada, así el usuario no tiene que hacerlo a mano.

    Devuelve: (namespace, carpeta)
    """
    outlook   = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")

    entry_id = None
    store_id = None

    try:
        inbox    = namespace.GetDefaultFolder(6)   # 6 = Bandeja de Entrada
        entry_id, store_id = _obtener_o_crear_carpeta(inbox, namespace, nombre_carpeta)
    except Exception as e:
        print(f"[SOW Manager] Error al buscar/crear carpeta: {e}")

    def _carpeta_fresca():
        """Re-obtiene la carpeta del store, sin caché de navegación."""
        if entry_id and store_id:
            try:
                return namespace.GetFolderFromID(entry_id, store_id)
            except Exception:
                return None
        return None

    try:
        # ── Estrategia 1: SyncObjects (Exchange corporativo, sin recordatorios) ──
        # Fuerza sync explícito en cada grupo Send/Receive configurado en Outlook.
        # Funciona aunque la máquina sea nueva y no tenga eventos COM pendientes.
        sync_objects = namespace.SyncObjects
        for i in range(sync_objects.Count):
            try:
                sync_objects.Item(i + 1).Start()
            except Exception:
                pass

        # ── Estrategia 2: SendAndReceive + PumpWaitingMessages (fallback) ────────
        # Cubre casos donde SyncObjects no dispara (ej. cuentas IMAP/POP).
        namespace.SendAndReceive(False)

        MIN_CICLOS = 30    # 15 segundos mínimo garantizado
        MAX_CICLOS = 70    # 35 segundos máximo absoluto
        conteo_anterior = -1
        estables        = 0

        for ciclo in range(MAX_CICLOS):
            pythoncom.PumpWaitingMessages()
            time.sleep(0.5)

            cf = _carpeta_fresca()
            if cf is not None:
                try:
                    conteo_actual = cf.Items.Count
                except Exception:
                    conteo_actual = conteo_anterior
            else:
                conteo_actual = conteo_anterior

            if conteo_actual != conteo_anterior:
                estables        = 0
                conteo_anterior = conteo_actual
            else:
                estables += 1

            if ciclo >= MIN_CICLOS and estables >= 8:
                break

    except Exception:
        pass

    # Devolver SIEMPRE una carpeta fresca recién obtenida del store
    carpeta_final = _carpeta_fresca()
    return namespace, carpeta_final


def obtener_carpeta_sows(namespace, nombre_carpeta="SOWs"):
    """
    Busca la subcarpeta SOWs dentro del Inbox.
    Fallback por si conectar_outlook no pudo devolver la carpeta.
    Si no existe, la crea automáticamente.
    Usa GetFolderFromID para evitar caché de navegación.
    """
    inbox    = namespace.GetDefaultFolder(6)
    entry_id, store_id = _obtener_o_crear_carpeta(inbox, namespace, nombre_carpeta)

    try:
        return namespace.GetFolderFromID(entry_id, store_id)
    except Exception:
        # Fallback: buscar por nombre directamente
        for carpeta in inbox.Folders:
            if carpeta.Name.lower() == nombre_carpeta.lower():
                return carpeta
        raise Exception(f"No se pudo obtener la carpeta '{nombre_carpeta}' tras crearla.")


def obtener_email_remitente(item):
    """Devuelve el email del remitente en formato normal (no Exchange)."""
    try:
        if item.SenderEmailType == "EX":
            sender = item.Sender
            if sender:
                exchange_user = sender.GetExchangeUser()
                if exchange_user:
                    return exchange_user.PrimarySmtpAddress
            return item.SenderName
        else:
            return item.SenderEmailAddress
    except Exception:
        return item.SenderName or "(desconocido)"


def listar_correos(carpeta, solo_no_procesados=True):
    """
    Lista correos de la carpeta.
    El doble Sort invalida el cache en memoria de Outlook y
    fuerza una re-lectura fresca del OST.
    """
    correos = []
    items   = carpeta.Items
    items.Sort("[ReceivedTime]", True)
    items.Sort("[ReceivedTime]", True)   # doble Sort: invalida cache interno

    for item in items:
        if item.Class != 43:
            continue

        correo = {
            "asunto":           item.Subject or "",
            "remitente":        obtener_email_remitente(item),
            "remitente_nombre": item.SenderName or "",
            "fecha":            item.ReceivedTime.strftime("%Y-%m-%d %H:%M:%S"),
            "cuerpo_html":      item.HTMLBody or "",
            "cuerpo_texto":     item.Body or "",
            "num_adjuntos":     item.Attachments.Count,
            "entry_id":         item.EntryID,
            "item":             item,
        }
        correos.append(correo)

    return correos


def obtener_adjuntos_pdf(correo):
    """Devuelve los adjuntos PDF/Word de un correo como bytes en memoria."""
    adjuntos = []
    item     = correo["item"]

    with tempfile.TemporaryDirectory(prefix="sow_") as carpeta_temp:
        for att in item.Attachments:
            nombre = att.FileName.lower()
            if nombre.endswith((".pdf", ".doc", ".docx")):
                ruta_temp = os.path.join(carpeta_temp, att.FileName)
                try:
                    att.SaveAsFile(ruta_temp)
                    with open(ruta_temp, "rb") as f:
                        contenido = f.read()
                    adjuntos.append({
                        "nombre":    att.FileName,
                        "contenido": contenido,
                    })
                except Exception as e:
                    print(f"   ⚠️ No se pudo leer {att.FileName}: {e}")

    return adjuntos


def guardar_adjuntos(correo, carpeta_destino):
    """[Legacy] Guarda adjuntos en disco. Mantenida por compatibilidad."""
    carpeta_destino = os.path.abspath(carpeta_destino)
    os.makedirs(carpeta_destino, exist_ok=True)
    rutas = []
    item  = correo["item"]
    for att in item.Attachments:
        nombre = att.FileName.lower()
        if nombre.endswith((".pdf", ".doc", ".docx")):
            ruta = os.path.join(carpeta_destino, att.FileName)
            try:
                att.SaveAsFile(ruta)
                rutas.append(ruta)
            except Exception as e:
                print(f"   ⚠️ No se pudo guardar {att.FileName}: {e}")
    return rutas


def marcar_como_procesado(correo):
    """Asigna la categoría 'Procesado SOW' al correo."""
    item = correo["item"]
    cats = item.Categories or ""
    if "Procesado SOW" not in cats:
        item.Categories = (cats + ";" if cats else "") + "Procesado SOW"
        item.Save()


if __name__ == "__main__":
    print("Conectando a Outlook...")
    ns, carpeta = conectar_outlook()
    if carpeta is None:
        carpeta = obtener_carpeta_sows(ns)
    print(f"Carpeta encontrada: {carpeta.Name}\n")

    correos = listar_correos(carpeta, solo_no_procesados=False)
    print(f"Total de correos: {len(correos)}\n")

    for i, c in enumerate(correos, 1):
        print(f"{i}. {c['fecha']}")
        print(f"   De: {c['remitente_nombre']} <{c['remitente']}>")
        print(f"   Asunto: {c['asunto']}")
        print(f"   Adjuntos: {c['num_adjuntos']}")
        print()