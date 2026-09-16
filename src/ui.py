"""
SOW Manager — Interfaz gráfica.
Layout: Header + Tabs (Pendientes / Novedades) + Modal de detalles con análisis Word.
"""

import tkinter as tk
from tkinter import ttk, messagebox
import os
import threading
from datetime import datetime
from src.excel_writer import (
    escribir_sows_en_excel,
    escribir_actualizaciones_excel,
)

FONT      = "Segoe UI"
FONT_MONO = "Consolas"

C = {
    "bg":            "#F7F8FA",
    "surface":       "#FFFFFF",
    "surface_2":     "#F0F2F5",
    "header_bg":     "#0F1923",
    "header_line":   "#1E88E5",
    "btn_primary":   "#1E88E5",
    "btn_primary_h": "#1565C0",
    "btn_success":   "#16A34A",
    "btn_success_h": "#15803D",
    "btn_neutral":   "#475569",
    "btn_neutral_h": "#334155",
    "accent":        "#1E88E5",
    "text":          "#1A1A2E",
    "text_2":        "#6B7280",
    "text_inv":      "#FFFFFF",
    "border":        "#E5E7EB",
    "tab_active":    "#1E88E5",
    "tab_inactive":  "#F0F2F5",
    "row_pendiente": "#FFFBEB",
    "row_nuevo":     "#EFF6FF",
    "row_firmado":   "#F0FDF4",
    "row_err":       "#FEF2F2",
    "dot_pendiente": "#F59E0B",
    "dot_nuevo":     "#1E88E5",
    "dot_firmado":   "#22C55E",
    "dot_err":       "#EF4444",
}

# RUTA_SOWS  = os.path.join(os.path.expanduser("~"), "Desktop", "SOWs Procesados")
# RUTA_EXCEL = os.path.join(os.path.expanduser("~"), "Desktop", "Listado SOW's.xlsm")

_BASE      = os.path.join(os.path.expanduser("~"), "EPAM",
                 "92101343 - CHUBB COLOMBIA - 1 DEV FT - Documents",
                 "0. Documentacion Contractual")
RUTA_SOWS  = os.path.join(_BASE, "Contratos - SOW")
RUTA_EXCEL = os.path.join(_BASE, "Listado SOW's.xlsm")


class AppSOWs:

    def __init__(self, root):
        self.root        = root
        self._pendientes = []
        self._novedades  = []
        self._tab_activo = "pendientes"
        self._setup_styles()
        self._build()
        self.root.after(300, self._cargar_pendientes)

    # ── Estilos ────────────────────────────────────────────────

    def _setup_styles(self):
        self.root.title("SOW Manager")
        self.root.configure(bg=C["bg"])
        self.root.update_idletasks()

        screen_w = self.root.winfo_screenwidth()
        screen_h = self.root.winfo_screenheight()

        # Usar 92% del ancho y 88% del alto disponible
        # Funciona en cualquier escala sin distorsionar el diseño
        w = int(screen_w * 0.92)
        h = int(screen_h * 0.88)

        # Respetar máximos para pantallas grandes
        w = min(w, 1060)
        h = min(h, 680)

        x = (screen_w - w) // 2
        y = max(0, (screen_h - h) // 2)
        self.root.geometry(f"{w}x{h}+{x}+{y}")
        self.root.minsize(700, 420)

        s = ttk.Style()
        s.theme_use("clam")
        s.configure("SOW.Treeview",
            background=C["surface"], foreground=C["text"],
            rowheight=42, fieldbackground=C["surface"],
            borderwidth=0, font=(FONT, 10))
        s.configure("SOW.Treeview.Heading",
            background=C["surface_2"], foreground=C["text_2"],
            font=(FONT, 9, "bold"), padding=(12, 10),
            relief="flat", borderwidth=0)
        s.map("SOW.Treeview",
            background=[("selected", "#DBEAFE")],
            foreground=[("selected", C["text"])])
        s.configure("Thin.Vertical.TScrollbar",
            troughcolor=C["bg"], background=C["border"],
            borderwidth=0, arrowsize=0, width=6)
    # ── Build ──────────────────────────────────────────────────

    def _build(self):
        self._build_header()
        self._build_tabs()
        self._build_footer()     # ← AHORA va antes del contenido
        self._build_content()    # ← el expandible siempre al final

    def _build_header(self):
        hdr = tk.Frame(self.root, bg=C["header_bg"], height=72)
        hdr.pack(fill="x")
        hdr.pack_propagate(False)
        tk.Frame(self.root, bg=C["header_line"], height=2).pack(fill="x")

        left = tk.Frame(hdr, bg=C["header_bg"])
        left.pack(side="left", padx=24)
        tk.Label(left, text="SOW Manager",
            font=(FONT, 15, "bold"),
            bg=C["header_bg"], fg=C["text_inv"]
        ).pack(anchor="w", pady=(14, 0))
        tk.Label(left, text="Automatización de Statements of Work · NEORIS",
            font=(FONT, 8),
            bg=C["header_bg"], fg="#64748B"
        ).pack(anchor="w")

        right = tk.Frame(hdr, bg=C["header_bg"])
        right.pack(side="right", padx=24)

        self.lbl_fecha = tk.Label(right,
            text="Cargando...", font=(FONT, 8),
            bg=C["header_bg"], fg="#64748B")
        self.lbl_fecha.pack(side="right", padx=(16, 0))

        for txt, cmd in [("📊", self.on_abrir_excel), ("📂", self.on_abrir_carpeta)]:
            b = tk.Button(right, text=txt, font=(FONT, 12),
                bg=C["header_bg"], fg=C["text_inv"],
                activebackground="#1E2A38", activeforeground=C["text_inv"],
                relief="flat", padx=10, pady=6, cursor="hand2", bd=0,
                command=cmd)
            b.pack(side="right", padx=(4, 0), pady=18)

        self.btn_sync = tk.Button(right,
            text="🔄  Sincronizar", font=(FONT, 10, "bold"),
            bg=C["btn_primary"], fg=C["text_inv"],
            activebackground=C["btn_primary_h"], activeforeground=C["text_inv"],
            relief="flat", padx=20, pady=8, cursor="hand2", bd=0,
            command=self.on_sincronizar)
        self.btn_sync.pack(side="right", pady=18, padx=(0, 8))
        self._bind_hover(self.btn_sync, C["btn_primary_h"], C["btn_primary"])

    def _build_tabs(self):
        self.tabs_frame = tk.Frame(self.root, bg=C["bg"])
        self.tabs_frame.pack(fill="x", padx=20, pady=(12, 0))

        self._tab_btns = {}
        for key, label in [("pendientes", "🟡  Pendientes"),
                           ("novedades",  "🔵  Novedades")]:
            btn = tk.Button(self.tabs_frame, text=label,
                font=(FONT, 10, "bold"), relief="flat", bd=0,
                padx=20, pady=10, cursor="hand2",
                command=lambda k=key: self._cambiar_tab(k))
            btn.pack(side="left", padx=(0, 4))
            self._tab_btns[key] = btn

        self._actualizar_tab_visual()

    def _build_content(self):
        self.content_frame = tk.Frame(self.root, bg=C["bg"])
        self.content_frame.pack(fill="both", expand=True, padx=20, pady=(6, 0))

        # ── Vista Pendientes ──────────────────────────────────
        self.frame_pendientes = tk.Frame(self.content_frame, bg=C["bg"])

        tb_pend = tk.Frame(self.frame_pendientes, bg=C["bg"], height=36)
        tb_pend.pack(fill="x", pady=(0, 6))
        tb_pend.pack_propagate(False)

        self.lbl_count_pend = tk.Label(tb_pend,
            text="0 SOWs pendientes por firmar",
            font=(FONT, 9), bg=C["bg"], fg=C["text_2"])
        self.lbl_count_pend.pack(side="left", pady=8)

        self.tree_pend, self._wrap_pend = self._build_tree(
            self.frame_pendientes,
            cols=[("estado","Estado",110,"center"),("sow","SOW",230,"w"),
                  ("persona","Persona",220,"w"),("remitente","Remitente",160,"w"),
                  ("fecha","Fecha Email",120,"center")])
        self.tree_pend.tag_configure("pendiente", background=C["row_pendiente"])
        self.tree_pend.tag_configure("err",       background=C["row_err"])
        # Mejora 1: modal de pendientes ahora muestra análisis del Word archivado
        self.tree_pend.bind("<Double-Button-1>",
            lambda e: self._abrir_modal(
                self.tree_pend, self._pendientes, mostrar_analisis=True))

        self._empty_pend = self._build_empty(
            self._wrap_pend, "✅", "Todo al día",
            "No hay SOWs pendientes por firmar")

        # ── Vista Novedades ───────────────────────────────────
        self.frame_novedades = tk.Frame(self.content_frame, bg=C["bg"])

        tb_nov = tk.Frame(self.frame_novedades, bg=C["bg"], height=36)
        tb_nov.pack(fill="x", pady=(0, 6))
        tb_nov.pack_propagate(False)

        self.lbl_count_nov = tk.Label(tb_nov,
            text="Presiona Sincronizar para buscar novedades",
            font=(FONT, 9), bg=C["bg"], fg=C["text_2"])
        self.lbl_count_nov.pack(side="left", pady=8)

        self.btn_guardar = tk.Button(tb_nov,
            text="📥  Guardar en Excel", font=(FONT, 9, "bold"),
            bg=C["btn_success"], fg=C["text_inv"],
            activebackground=C["btn_success_h"], activeforeground=C["text_inv"],
            relief="flat", padx=14, pady=5, cursor="hand2", bd=0,
            state="disabled", command=self.on_guardar_excel)
        self.btn_guardar.pack(side="right", pady=4)
        self._bind_hover(self.btn_guardar, C["btn_success_h"], C["btn_success"])

        self.btn_guardar_sel = tk.Button(tb_nov,
            text="☑  Guardar selección", font=(FONT, 9, "bold"),
            bg=C["btn_neutral"], fg=C["text_inv"],
            activebackground=C["btn_neutral_h"], activeforeground=C["text_inv"],
            relief="flat", padx=14, pady=5, cursor="hand2", bd=0,
            state="disabled", command=self.on_guardar_seleccion)
        self.btn_guardar_sel.pack(side="right", pady=4, padx=(0, 6))
        self._bind_hover(self.btn_guardar_sel, C["btn_neutral_h"], C["btn_neutral"])

        self.tree_nov, self._wrap_nov = self._build_tree(
            self.frame_novedades,
            cols=[("tipo","Tipo",110,"center"),("sow","SOW",230,"w"),
                  ("persona","Persona",220,"w"),("remitente","Remitente",160,"w"),
                  ("fecha","Fecha Email",120,"center")])
        self.tree_nov.tag_configure("nuevo",   background=C["row_nuevo"])
        self.tree_nov.tag_configure("firmado", background=C["row_firmado"])
        self.tree_nov.tag_configure("err",     background=C["row_err"])
        self.tree_nov.bind("<Double-Button-1>",
            lambda e: self._abrir_modal(
                self.tree_nov, self._novedades, mostrar_analisis=True))
        self.tree_nov.bind("<<TreeviewSelect>>", self._on_nov_seleccion_cambio)

        self._empty_nov = self._build_empty(
            self._wrap_nov, "📭", "Sin novedades",
            "Presiona  Sincronizar  para buscar correos nuevos")

        self.frame_pendientes.pack(fill="both", expand=True)

    def _build_footer(self):
        footer = tk.Frame(self.root, bg=C["bg"], height=28)
        footer.pack(side="bottom", fill="x", padx=20, pady=(4, 8))   # ← side="bottom"
        footer.pack_propagate(False)

        tk.Label(footer, text=f"Excel: {RUTA_EXCEL}",
            font=(FONT_MONO, 8), bg=C["bg"], fg=C["text_2"],
            anchor="w").pack(side="left")

        self.lbl_status = tk.Label(footer, text="",
            font=(FONT, 8), bg=C["bg"], fg=C["accent"], anchor="e")
        self.lbl_status.pack(side="right")

    # ── Helpers UI ─────────────────────────────────────────────

    def _build_tree(self, parent, cols):
        wrapper = tk.Frame(parent, bg=C["surface"],
            highlightbackground=C["border"], highlightthickness=1)
        wrapper.pack(fill="both", expand=True)

        tree = ttk.Treeview(wrapper,
            columns=[c[0] for c in cols], show="headings",
            style="SOW.Treeview", selectmode="extended")

        for col_id, txt, w, anchor in cols:
            tree.heading(col_id, text=txt.upper(),
                anchor="w" if anchor == "w" else "center")
            tree.column(col_id, width=w,
                anchor=anchor, minwidth=60, stretch=True)

        sb = ttk.Scrollbar(wrapper, orient="vertical",
            command=tree.yview, style="Thin.Vertical.TScrollbar")
        tree.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y", padx=(0, 2), pady=4)
        tree.pack(side="left", fill="both", expand=True)

        return tree, wrapper

    def _build_empty(self, parent, icon, title, subtitle):
        frame = tk.Frame(parent, bg=C["surface"])
        tk.Label(frame, text=icon,
            font=("Segoe UI Emoji", 32),
            bg=C["surface"], fg=C["border"]).pack()
        tk.Label(frame, text=title,
            font=(FONT, 13, "bold"),
            bg=C["surface"], fg=C["text_2"]).pack(pady=(4, 2))
        tk.Label(frame, text=subtitle,
            font=(FONT, 9),
            bg=C["surface"], fg=C["text_2"]).pack()
        return frame

    def _bind_hover(self, widget, color_in, color_out):
        widget.bind("<Enter>", lambda e: widget.config(bg=color_in))
        widget.bind("<Leave>", lambda e: widget.config(bg=color_out))

    def _actualizar_tab_visual(self):
        for key, btn in self._tab_btns.items():
            if key == self._tab_activo:
                btn.config(bg=C["tab_active"], fg=C["text_inv"],
                    activebackground=C["btn_primary_h"],
                    activeforeground=C["text_inv"])
            else:
                btn.config(bg=C["tab_inactive"], fg=C["text_2"],
                    activebackground=C["surface_2"],
                    activeforeground=C["text"])

    def _cambiar_tab(self, key):
        self._tab_activo = key
        self._actualizar_tab_visual()
        self.frame_pendientes.pack_forget()
        self.frame_novedades.pack_forget()
        if key == "pendientes":
            self.frame_pendientes.pack(fill="both", expand=True)
        else:
            self.frame_novedades.pack(fill="both", expand=True)

    def _fmt_fecha(self, f):
        if isinstance(f, datetime):
            return f.strftime("%d/%m/%Y")
        if f:
            return str(f)[:10]
        return "—"

    def _poblar_tree(self, tree, filas, empty_frame, wrapper):
        for item in tree.get_children():
            tree.delete(item)

        if not filas:
            empty_frame.place(relx=0.5, rely=0.5, anchor="center")
            return

        empty_frame.place_forget()

        for f in filas:
            tipo = f.get("tipo", f.get("estado", "pendiente"))
            etiqueta = {
                "pendiente": "🟡  Pendiente",
                "nuevo":     "🔵  Word nuevo",
                "firmado":   "✅  PDF firmado",
                "err":       "🔴  Error",
            }.get(tipo, "🟡  Pendiente")

            tree.insert("", "end",
                values=(etiqueta, f.get("sow",""),
                        f.get("nombre","").title(),
                        f.get("remitente","Gema Vale"),
                        self._fmt_fecha(f.get("fecha_email",""))),
                tags=(tipo,))

    # ── Cargar pendientes desde Excel ─────────────────────────

    def _cargar_pendientes(self):
        if not os.path.exists(RUTA_EXCEL):
            self.lbl_fecha.config(text="Excel no encontrado")
            self.lbl_status.config(text="⚠ Excel no encontrado")
            return
        try:
            self.lbl_status.config(text="Leyendo Excel...")
            self.root.update()

            from src.excel_writer import leer_sows_del_excel
            self._pendientes = leer_sows_del_excel(ruta_excel=RUTA_EXCEL)

            self._poblar_tree(self.tree_pend, self._pendientes,
                              self._empty_pend, self._wrap_pend)

            total = len(self._pendientes)
            self.lbl_count_pend.config(
                text=f"{total} SOW{'s' if total != 1 else ''} "
                     f"pendiente{'s' if total != 1 else ''} por firmar")

            ahora = datetime.now().strftime("%d/%m/%Y  %H:%M")
            self.lbl_fecha.config(text=f"Actualizado · {ahora}")
            self.lbl_status.config(text="")

        except Exception as e:
            import traceback
            print(traceback.format_exc())
            self.lbl_status.config(text=f"Error al leer Excel: {e}")

    # ── Sincronizar ───────────────────────────────────────────

    def on_sincronizar(self):
        self.btn_sync.config(state="disabled", text="🔄  Sincronizando...")
        self.lbl_status.config(text="Conectando a Outlook...")
        self.root.update()
        self.root.after(120, self._ejecutar_sync)

    def _ejecutar_sync(self):
        try:
            from src.main import procesar_correos
            from src.excel_writer import leer_indice_excel

            def log(msg):
                short = msg[:80] + "..." if len(msg) > 80 else msg
                self.lbl_status.config(text=short)
                self.root.update()

            resultado = procesar_correos(
                ruta_excel         = RUTA_EXCEL,
                solo_no_procesados = False,
                marcar_procesados  = False,
                escribir_excel     = False,
                callback_progreso  = log,
            )

            try:
                indice_excel = leer_indice_excel(ruta_excel=RUTA_EXCEL)
            except Exception:
                indice_excel = {}

            def _norm(n):
                return " ".join(str(n or "").upper().split())

            novedades = []
            vistos    = set()

            for correo in resultado.get("correos", []):
                remitente   = correo.get("remitente_nombre", "Gema Vale")
                fecha_email = correo.get("fecha", "")
                filas_tabla = correo.get("sows_tabla", [])
                archivos    = correo.get("archivos", [])

                arch_por_sow = {}
                for arch in archivos:
                    sow_arch = arch.get("sow", "").upper()
                    if sow_arch not in arch_por_sow:
                        arch_por_sow[sow_arch] = arch

                if filas_tabla:
                    for fila in filas_tabla:
                        sow_num  = fila.get("sow", "").upper()
                        personas = fila.get("personas", [])
                        empresa  = fila.get("empresa", "")
                        if isinstance(personas, str):
                            personas = [personas]

                        arch           = arch_por_sow.get(sow_num, {})
                        es_pdf         = arch.get("es_pdf", False)
                        contenido_word = arch.get("contenido")
                        fecha_inicio   = arch.get("fecha_inicio")
                        fecha_fin      = arch.get("fecha_fin")
                        valor_cop      = arch.get("valor_cop")
                        firmado        = es_pdf and bool(arch)

                        ruta_archivo = ""
                        ruta_carpeta = ""
                        for g in arch.get("guardados", []):
                            rg = g.get("ruta_guardada", "")
                            if rg:
                                ruta_carpeta = os.path.dirname(rg)
                                if es_pdf:
                                    ruta_archivo = rg
                                break

                        for persona in personas:
                            if not persona:
                                continue
                            key          = (sow_num, _norm(persona))
                            if key in vistos:
                                continue
                            estado_excel = indice_excel.get(key)

                            if firmado:
                                if estado_excel == "cerrado":
                                    continue
                                tipo = "firmado"
                            else:
                                if estado_excel in ("pendiente", "cerrado"):
                                    continue
                                tipo = "nuevo"

                            vistos.add(key)
                            novedades.append({
                                "tipo": tipo, "estado": tipo,
                                "sow": sow_num, "nombre": persona,
                                "empresa": empresa,
                                "remitente": remitente,
                                "fecha_email": fecha_email,
                                "fecha_inicio": fecha_inicio,
                                "fecha_fin": fecha_fin,
                                "valor_cop": valor_cop,
                                "ruta_archivo": ruta_archivo,
                                "ruta_carpeta": ruta_carpeta,
                                "firmado": firmado,
                                "contenido_word": contenido_word,
                            })

                else:
                    for arch in archivos:
                        if not arch.get("es_pdf"):
                            continue
                        sow_num  = arch.get("sow", "").upper()
                        personas = arch.get("personas", [])
                        if isinstance(personas, str):
                            personas = [personas]
                        if not personas:
                            continue

                        ruta_archivo = ""
                        ruta_carpeta = ""
                        for g in arch.get("guardados", []):
                            rg = g.get("ruta_guardada", "")
                            if rg:
                                ruta_archivo = rg
                                ruta_carpeta = os.path.dirname(rg)
                                break

                        for persona in personas:
                            if not persona:
                                continue
                            key = (sow_num, _norm(persona))
                            if key in vistos:
                                continue
                            if indice_excel.get(key) == "cerrado":
                                continue
                            vistos.add(key)
                            novedades.append({
                                "tipo": "firmado", "estado": "firmado",
                                "sow": sow_num, "nombre": persona,
                                "empresa": arch.get("empresa",""),
                                "remitente": remitente,
                                "fecha_email": fecha_email,
                                "fecha_inicio": arch.get("fecha_inicio"),
                                "fecha_fin": arch.get("fecha_fin"),
                                "valor_cop": arch.get("valor_cop"),
                                "ruta_archivo": ruta_archivo,
                                "ruta_carpeta": ruta_carpeta,
                                "firmado": True,
                                "contenido_word": None,
                            })

            self._novedades = novedades
            self._poblar_tree(self.tree_nov, self._novedades,
                              self._empty_nov, self._wrap_nov)

            total_nov = len(novedades)
            nuevos    = sum(1 for n in novedades if n["tipo"] == "nuevo")
            firmados  = sum(1 for n in novedades if n["tipo"] == "firmado")

            self.lbl_count_nov.config(
                text=(f"{total_nov} novedad"
                      f"{'es' if total_nov != 1 else ''}  ·  "
                      f"{nuevos} Word nuevo"
                      f"{'s' if nuevos != 1 else ''}  ·  "
                      f"{firmados} PDF firmado"
                      f"{'s' if firmados != 1 else ''}"))

            self.btn_guardar.config(state="normal" if total_nov > 0 else "disabled")
            self.btn_guardar_sel.config(state="disabled")

            # ── Mejora 4: releer Pendientes desde Excel en cada sync ──
            self._cargar_pendientes()

            ahora = datetime.now().strftime("%d/%m/%Y  %H:%M")
            self.lbl_fecha.config(text=f"Sincronizado · {ahora}")
            self.lbl_status.config(text="")
            self.btn_sync.config(state="normal", text="🔄  Sincronizar")
            self._cambiar_tab("novedades")

        except Exception as e:
            import traceback
            print(traceback.format_exc())
            self.btn_sync.config(state="normal", text="🔄  Sincronizar")
            self.lbl_status.config(text="Error al sincronizar")
            messagebox.showerror("Error",
                f"Error durante la sincronización:\n\n{str(e)}")

    # ── Guardar en Excel ──────────────────────────────────────

    def _on_nov_seleccion_cambio(self, event=None):
        tiene = len(self.tree_nov.selection()) > 0
        self.btn_guardar_sel.config(state="normal" if tiene else "disabled")

    def on_guardar_excel(self):
        self._ejecutar_guardar(self._novedades)

    def on_guardar_seleccion(self):
        sel = self.tree_nov.selection()
        if not sel:
            return
        indices       = [self.tree_nov.index(item) for item in sel]
        seleccionadas = [self._novedades[i] for i in indices
                         if i < len(self._novedades)]
        if seleccionadas:
            self._ejecutar_guardar(seleccionadas)

    def _ejecutar_guardar(self, novedades_lista: list):
        if not novedades_lista:
            messagebox.showinfo("Sin datos", "No hay novedades para guardar.")
            return
        if not os.path.exists(RUTA_EXCEL):
            messagebox.showerror("Excel no encontrado",
                f"No se encontró:\n{RUTA_EXCEL}")
            return
        try:
            from src.excel_writer import (
                escribir_sows_en_excel,
                escribir_actualizaciones_excel,
            )

            sows_nuevos = []
            sows_update = []

            for nov in novedades_lista:
                nombre = nov.get("nombre", "").strip()
                if not nombre or "(sin persona)" in nombre or \
                        "(no identificada)" in nombre:
                    continue
                entrada = {
                    "nombre":       nombre,
                    "employee_id":  "",
                    "sow":          nov.get("sow", ""),
                    "empresa":      nov.get("empresa", ""),
                    "fecha_inicio": nov.get("fecha_inicio"),
                    "fecha_fin":    nov.get("fecha_fin"),
                    "valor_cop":    nov.get("valor_cop"),
                    "ruta_archivo": nov.get("ruta_archivo", ""),
                    "ruta_carpeta": nov.get("ruta_carpeta", ""),
                    "firmado":      nov.get("firmado", False),
                }
                if nov.get("firmado"):
                    sows_update.append(entrada)
                else:
                    sows_nuevos.append(entrada)

            if not sows_nuevos and not sows_update:
                messagebox.showwarning("Sin datos válidos",
                    "No hay SOWs con personas identificadas.")
                return

            # ── Mejora 2: advertir SOWs sin fechas ni valor ───────────
            incompletos = [
                n for n in novedades_lista
                if not n.get("firmado", False)
                   and not n.get("fecha_inicio")
                   and not n.get("valor_cop")
            ]
            if incompletos:
                lista_txt = "\n".join(
                    f"  • {n['sow']}  —  {n.get('nombre','').title()}"
                    for n in incompletos[:5]
                )
                if len(incompletos) > 5:
                    lista_txt += f"\n  ... y {len(incompletos) - 5} más"
                continuar = messagebox.askyesno(
                    "SOWs sin datos completos",
                    f"Los siguientes SOWs no tienen fechas ni valor\n"
                    f"(probablemente no llegó el Word adjunto):\n\n"
                    f"{lista_txt}\n\n"
                    f"¿Guardarlos igual con datos incompletos?"
                )
                if not continuar:
                    return

            self.btn_guardar.config(state="disabled", text="Guardando...")
            self.btn_guardar_sel.config(state="disabled")
            self.lbl_status.config(text="Escribiendo en Excel...")
            self.root.update()

            agregados       = 0
            existentes      = 0
            firmados_count  = 0
            errores_ex      = 0
            words_guardados = []
            pdfs_guardados  = set()
            guardados_keys  = set()

            if sows_nuevos:
                res        = escribir_sows_en_excel(sows_nuevos, ruta_excel=RUTA_EXCEL)
                agregados  += sum(1 for r in res if r["accion"] == "agregado")
                existentes += sum(1 for r in res if r["accion"] == "ya_existe")
                errores_ex += sum(1 for r in res if r["accion"] == "error")
                agregados_set = {(r["sow"].upper(), r["nombre"].upper())
                                 for r in res if r["accion"] == "agregado"}
                for nov in novedades_lista:
                    if not nov.get("firmado", False):
                        k = (nov.get("sow","").upper(), nov.get("nombre","").upper())
                        if k in agregados_set:
                            words_guardados.append(nov)
                            guardados_keys.add(k)
                guardados_keys |= {(r["sow"].upper(), r["nombre"].upper())
                                   for r in res if r["accion"] == "ya_existe"}

            if sows_update:
                res_upd         = escribir_actualizaciones_excel(
                    sows_update, ruta_excel=RUTA_EXCEL)
                firmados_count += sum(1 for r in res_upd if r["accion"] == "firmado")
                for r in res_upd:
                    if r["accion"] == "firmado":
                        k = (r["sow"].upper(), r["nombre"].upper())
                        pdfs_guardados.add(k)
                        guardados_keys.add(k)
                no_encontrados = [
                    sows_update[i] for i, r in enumerate(res_upd)
                    if r["accion"] == "no_encontrado"
                ]
                if no_encontrados:
                    res2       = escribir_sows_en_excel(no_encontrados, ruta_excel=RUTA_EXCEL)
                    agregados += sum(1 for r in res2 if r["accion"] == "agregado")
                    for r in res2:
                        if r["accion"] == "agregado":
                            guardados_keys.add((r["sow"].upper(), r["nombre"].upper()))

            self.btn_guardar.config(
                state="normal" if self._novedades else "disabled",
                text="📥  Guardar en Excel")
            self.lbl_status.config(text="")

            messagebox.showinfo("✅ Listo",
                f"Guardado correctamente\n\n"
                f"Filas nuevas:     {agregados}\n"
                f"Ya existían:      {existentes}\n"
                f"PDFs firmados:    {firmados_count}\n"
                f"Con errores:      {errores_ex}")

            # ── Actualizar Pendientes en memoria ──────────────────────
            for nov in words_guardados:
                self._pendientes.append({
                    "sow":          nov.get("sow", ""),
                    "nombre":       nov.get("nombre", ""),
                    "empresa":      nov.get("empresa", ""),
                    "fecha_inicio": nov.get("fecha_inicio"),
                    "fecha_fin":    nov.get("fecha_fin"),
                    "valor_cop":    nov.get("valor_cop"),
                    "firmado":      False,
                    "ruta_archivo": "",
                    "ruta_carpeta": nov.get("ruta_carpeta", ""),
                    "estado":       "pendiente",
                    "tipo":         "pendiente",
                    "remitente":    nov.get("remitente", "Gema Vale"),
                    "fecha_email":  nov.get("fecha_email"),
                })

            if pdfs_guardados:
                self._pendientes = [
                    p for p in self._pendientes
                    if (p.get("sow","").upper(), p.get("nombre","").upper())
                       not in pdfs_guardados
                ]

            self._poblar_tree(self.tree_pend, self._pendientes,
                              self._empty_pend, self._wrap_pend)
            total_p = len(self._pendientes)
            self.lbl_count_pend.config(
                text=f"{total_p} SOW{'s' if total_p != 1 else ''} "
                     f"pendiente{'s' if total_p != 1 else ''} por firmar")
            self.lbl_fecha.config(
                text=f"Actualizado · {datetime.now().strftime('%d/%m/%Y  %H:%M')}")

            self._novedades = [
                n for n in self._novedades
                if (n.get("sow","").upper(), n.get("nombre","").upper())
                   not in guardados_keys
            ]
            self._poblar_tree(self.tree_nov, self._novedades,
                              self._empty_nov, self._wrap_nov)

            total_restante = len(self._novedades)
            if total_restante == 0:
                self.lbl_count_nov.config(
                    text="Presiona Sincronizar para buscar novedades")
                self.btn_guardar.config(state="disabled")
            else:
                nv = sum(1 for n in self._novedades if n["tipo"] == "nuevo")
                fi = sum(1 for n in self._novedades if n["tipo"] == "firmado")
                self.lbl_count_nov.config(
                    text=f"{total_restante} restantes · {nv} Word · {fi} PDF")

            self._cambiar_tab("pendientes")

        except Exception as e:
            import traceback
            print(traceback.format_exc())
            self.btn_guardar.config(state="normal", text="📥  Guardar en Excel")
            self.btn_guardar_sel.config(state="disabled")
            self.lbl_status.config(text="")
            messagebox.showerror("Error",
                f"Error al escribir en Excel:\n\n{str(e)}")

    # ── Modal detalles ────────────────────────────────────────

    def _abrir_modal(self, tree, fuente, mostrar_analisis=True):
        sel = tree.selection()
        if not sel:
            return
        idx = tree.index(sel[0])
        if idx >= len(fuente):
            return
        dato = fuente[idx]

        modal = tk.Toplevel(self.root)
        modal.title("Detalles del SOW")
        modal.configure(bg=C["surface"])
        modal.resizable(True, True)
        modal.grab_set()

        mw, mh = 580, 700
        px = self.root.winfo_x() + (self.root.winfo_width()  - mw) // 2
        py = self.root.winfo_y() + (self.root.winfo_height() - mh) // 2
        modal.geometry(f"{mw}x{mh}+{px}+{py}")

        hdr = tk.Frame(modal, bg=C["header_bg"], height=64)
        hdr.pack(fill="x")
        hdr.pack_propagate(False)

        tk.Label(hdr, text=dato.get("sow", "Sin SOW"),
            font=(FONT, 13, "bold"),
            bg=C["header_bg"], fg=C["text_inv"]
        ).pack(side="left", padx=24, pady=(16, 0))

        tipo = dato.get("tipo", dato.get("estado", "pendiente"))
        col_tipo = {"pendiente": C["dot_pendiente"], "nuevo": C["dot_nuevo"],
                    "firmado": C["dot_firmado"], "err": C["dot_err"]}.get(tipo, C["dot_pendiente"])
        lbl_tipo = {"pendiente": "⏳ Pendiente firma", "nuevo": "🔵 Word nuevo",
                    "firmado": "✅ PDF firmado", "err": "❌ Error"}.get(tipo, "⏳ Pendiente")

        tk.Label(hdr, text=lbl_tipo, font=(FONT, 9, "bold"),
            bg=C["header_bg"], fg=col_tipo
        ).pack(side="right", padx=24, pady=(16, 0))

        tk.Frame(modal, bg=C["header_line"], height=2).pack(fill="x")

        outer  = tk.Frame(modal, bg=C["surface"])
        outer.pack(fill="both", expand=True)
        canvas = tk.Canvas(outer, bg=C["surface"], highlightthickness=0)
        sb     = ttk.Scrollbar(outer, orient="vertical",
                                command=canvas.yview,
                                style="Thin.Vertical.TScrollbar")
        body   = tk.Frame(canvas, bg=C["surface"])

        body.bind("<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=body, anchor="nw")
        canvas.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)
        canvas.bind_all("<MouseWheel>",
            lambda e: canvas.yview_scroll(int(-1*(e.delta/120)), "units"))

        def _bloque(titulo, color_borde=C["accent"]):
            outer_f = tk.Frame(body, bg=C["surface"])
            outer_f.pack(fill="x", padx=20, pady=(10, 0))
            tk.Frame(outer_f, bg=color_borde, width=3).pack(side="left", fill="y")
            inner = tk.Frame(outer_f, bg=C["surface_2"])
            inner.pack(side="left", fill="x", expand=True)
            tk.Label(inner, text=titulo, font=(FONT, 8, "bold"),
                bg=C["surface_2"], fg=C["text_2"],
                padx=12, pady=6).pack(anchor="w")
            return inner

        def _fila(parent, label, valor, clickable=False, es_carpeta=False):
            row = tk.Frame(parent, bg=C["surface_2"])
            row.pack(fill="x", padx=12, pady=2)
            tk.Label(row, text=label, font=(FONT, 8),
                bg=C["surface_2"], fg=C["text_2"],
                width=13, anchor="w").pack(side="left")
            if clickable and valor and os.path.exists(str(valor)):
                icono = "📂 Abrir" if es_carpeta else "📄 Abrir"
                lnk   = tk.Label(row, text=icono, font=(FONT, 8, "bold"),
                    bg=C["surface_2"], fg=C["accent"], cursor="hand2")
                lnk.pack(side="left")
                lnk.bind("<Button-1>", lambda e, v=valor: os.startfile(v))
            else:
                tk.Label(row,
                    text=str(valor) if valor else "—",
                    font=(FONT, 8, "bold"),
                    bg=C["surface_2"], fg=C["text"],
                    anchor="w").pack(side="left")

        def _resultado(parent, texto):
            if "✅" in texto:
                color, bg = C["dot_firmado"], "#F0FDF4"
            elif "⚠️" in texto:
                color, bg = "#D97706", "#FFFBEB"
            else:
                color, bg = C["dot_err"], "#FEF2F2"
            f = tk.Frame(parent, bg=bg)
            f.pack(fill="x", padx=12, pady=2)
            tk.Label(f, text=texto, font=(FONT, 8), bg=bg, fg=color,
                anchor="w", padx=8, pady=4,
                wraplength=460, justify="left").pack(anchor="w")

        def _separador():
            tk.Frame(body, bg=C["border"], height=1).pack(
                fill="x", padx=20, pady=4)

        b1 = _bloque("📄  INFORMACIÓN BÁSICA", C["accent"])
        _fila(b1, "Persona",     dato.get("nombre", "").title())
        _fila(b1, "Empresa",     dato.get("empresa", ""))
        _fila(b1, "Remitente",   dato.get("remitente", "Gema Vale"))
        _fila(b1, "Fecha email", self._fmt_fecha(dato.get("fecha_email")))
        _fila(b1, "Archivo",     dato.get("ruta_archivo", ""),
            clickable=True, es_carpeta=False)
        _fila(b1, "Carpeta",     dato.get("ruta_carpeta", ""),
            clickable=True, es_carpeta=True)

        if mostrar_analisis:
            sow_num        = dato.get("sow", "")
            contenido_word = dato.get("contenido_word")

            # ── Mejora 1: buscar Word en carpeta archivada (para Pendientes) ──
            if not contenido_word:
                ruta_carpeta = dato.get("ruta_carpeta", "")
                if ruta_carpeta and os.path.exists(ruta_carpeta) and sow_num:
                    sow_upper = sow_num.upper()
                    try:
                        for archivo in os.listdir(ruta_carpeta):
                            if (archivo.lower().endswith((".docx", ".doc"))
                                    and sow_upper in archivo.upper()):
                                ruta_word = os.path.join(ruta_carpeta, archivo)
                                with open(ruta_word, "rb") as f:
                                    contenido_word = f.read()
                                break
                    except Exception:
                        pass

            analisis = None
            if contenido_word:
                try:
                    from src.word_extractor import analizar_sow_completo
                    analisis = analizar_sow_completo(
                        contenido_word,
                        numero_sow   = sow_num,
                        fecha_inicio = dato.get("fecha_inicio"),
                        fecha_fin    = dato.get("fecha_fin"),
                    )
                except Exception as ex:
                    print(f"Error analizando Word: {ex}")

            _separador()

            if analisis and not analisis.get("error"):
                global_ok = analisis.get("validacion_global", False)
                bg_ban    = "#F0FDF4" if global_ok else "#FEF2F2"
                col_ban   = C["dot_firmado"] if global_ok else C["dot_err"]
                txt_ban   = ("✅  Documento válido" if global_ok
                             else "⚠️  Hay puntos que requieren atención")
                banner = tk.Frame(body, bg=bg_ban)
                banner.pack(fill="x", padx=20, pady=(6, 2))
                tk.Label(banner, text=txt_ban,
                    font=(FONT, 10, "bold"), bg=bg_ban, fg=col_ban,
                    padx=16, pady=8).pack(anchor="w")

                _separador()

                b_cm = _bloque("1️⃣  CONTRATO MARCO", "#6366F1")
                for linea in analisis["contrato_marco"]["detalle"]:
                    _resultado(b_cm, linea)

                b_ts = _bloque("2️⃣  TIPO DE CONTRATO", "#6366F1")
                _resultado(b_ts, analisis["tipo_servicio"]["detalle"])

                b_cg = _bloque("3️⃣  CARGOS CONTRATADOS", "#0891B2")
                p3   = analisis["cargos"]
                if p3:
                    for cargo in p3:
                        f = tk.Frame(b_cg, bg=C["surface_2"])
                        f.pack(fill="x", padx=12, pady=2)
                        tk.Label(f,
                            text=f"  •  {cargo['cantidad']} × {cargo['cargo']}",
                            font=(FONT, 9, "bold"),
                            bg=C["surface_2"], fg=C["text"],
                            anchor="w", pady=3).pack(anchor="w")
                else:
                    _resultado(b_cg, "⚠️  No se detectaron cargos")

                p4      = analisis["vigencia"]
                alerta  = p4.get("alerta", "")
                col_vig = (C["dot_firmado"] if alerta == "vigente"
                           else "#D97706"   if alerta == "por_vencer"
                           else C["dot_err"])
                b_vig = _bloque("4️⃣  VIGENCIA DEL SERVICIO", col_vig)
                _resultado(b_vig, p4["detalle"])
                fi = p4.get("fecha_inicio")
                ff = p4.get("fecha_fin")
                if fi or ff:
                    f2 = tk.Frame(b_vig, bg=C["surface_2"])
                    f2.pack(fill="x", padx=12, pady=2)
                    tk.Label(f2,
                        text=(f"Inicio: {self._fmt_fecha(fi)}"
                              f"    →    Fin: {self._fmt_fecha(ff)}"),
                        font=(FONT, 8), bg=C["surface_2"], fg=C["text_2"],
                        pady=2).pack(anchor="w")

                b_dt = _bloque("5️⃣  SERVICIOS  (DT Original)", "#0891B2")
                p5   = analisis["tabla_dt"]
                if p5:
                    for item_dt in p5:
                        f = tk.Frame(b_dt, bg=C["surface_2"])
                        f.pack(fill="x", padx=12, pady=2)
                        tk.Label(f,
                            text=f"  •  {item_dt['servicio']}",
                            font=(FONT, 9, "bold"),
                            bg=C["surface_2"], fg=C["text"],
                            anchor="w", pady=3,
                            wraplength=460).pack(anchor="w")
                else:
                    _resultado(b_dt, "⚠️  No se encontró la tabla DT Original")

                b_imp = _bloque("6️⃣  CLÁUSULA DE IMPUESTOS", "#6366F1")
                _resultado(b_imp, analisis["impuestos"]["detalle"])

            else:
                b_na = _bloque("🔍  ANÁLISIS DEL DOCUMENTO", C["text_2"])
                f    = tk.Frame(b_na, bg=C["surface_2"])
                f.pack(fill="x", padx=12, pady=8)
                if not contenido_word:
                    msg = ("⚠️  No se encontró el archivo Word en la carpeta archivada.\n"
                           "Verifica que el documento esté en:\n"
                           f"{dato.get('ruta_carpeta', '—')}")
                else:
                    msg = (f"❌  Error al analizar.\n"
                           f"{analisis.get('error','') if analisis else ''}")
                tk.Label(f, text=msg, font=(FONT, 9),
                    bg=C["surface_2"], fg=C["text_2"],
                    justify="left", wraplength=460).pack(anchor="w", padx=8, pady=4)

        tk.Frame(body, bg=C["border"], height=1).pack(
            fill="x", padx=20, pady=(12, 0))
        tk.Button(body, text="  Cerrar  ",
            font=(FONT, 9, "bold"),
            bg=C["btn_primary"], fg=C["text_inv"],
            activebackground=C["btn_primary_h"],
            activeforeground=C["text_inv"],
            relief="flat", padx=20, pady=8,
            cursor="hand2", bd=0,
            command=lambda: [
                canvas.unbind_all("<MouseWheel>"),
                modal.destroy()
            ]).pack(pady=16)

    # ── Handlers secundarios ──────────────────────────────────

    def on_abrir_carpeta(self):
        os.makedirs(RUTA_SOWS, exist_ok=True)
        os.startfile(RUTA_SOWS)

    def on_abrir_excel(self):
        if os.path.exists(RUTA_EXCEL):
            os.startfile(RUTA_EXCEL)
        else:
            messagebox.showinfo("No encontrado",
                f"No se encontró:\n{RUTA_EXCEL}")


if __name__ == "__main__":
    root = tk.Tk()
    app  = AppSOWs(root)
    root.mainloop()