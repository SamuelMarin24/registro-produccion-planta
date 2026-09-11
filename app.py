"""
Registro de Producción en Planta — interfaz de escritorio para el operario.

Cada PC de máquina corre esta app. El operario cronometra el evento, digita la
OP (que autocompleta referencia y cliente) y guarda: el registro se escribe como
un archivo JSON en una carpeta compartida, que luego consolida consolidador.py.

Diseño: header con reloj y turno, cards limpias, cronómetro grande.
"""

import customtkinter as ctk
import tkinter as tk
from tkinter import messagebox, ttk
from PIL import Image
import os
from datetime import datetime
import threading

from data_manager import DataManager
from config import Config, NOMBRE_APP, TEXTO_LOGO

# ── Paleta ────────────────────────────────────────────────────────────────────
VERDE        = "#1a5c2a"
VERDE_OSC    = "#0f3d1a"
VERDE_MED    = "#22753a"
VERDE_CLARO  = "#e8f5ec"
VERDE_BORDE  = "#7ed89a"
BLANCO       = "#ffffff"
GRIS_BG      = "#f0f0f0"
GRIS_CARD    = "#ffffff"
GRIS_BORDE   = "#d0d0d0"
GRIS_TEXTO   = "#1a1a1a"
GRIS_SUB     = "#555555"
ROJO         = "#c0392b"
AMARILLO     = "#d29922"
VERDE_EXITO  = "#1a7a3c"

C = {
    "bg":        GRIS_BG,
    "card":      GRIS_CARD,
    "border":    GRIS_BORDE,
    "text":      GRIS_TEXTO,
    "subtext":   GRIS_SUB,
    "accent":    VERDE,
    "accent_med":VERDE_MED,
    "success":   VERDE_EXITO,
    "danger":    ROJO,
    "warning":   AMARILLO,
    "header_bg": VERDE,
    "readonly":  "#f4f4f4",
}

FONT_LABEL  = ("Segoe UI", 11)
FONT_INPUT  = ("Segoe UI", 12)
FONT_SMALL  = ("Segoe UI", 9)
FONT_BTN    = ("Segoe UI", 12, "bold")
FONT_TIMER  = ("Courier New", 38, "bold")
FONT_TITLE  = ("Segoe UI", 10, "bold")


# ── Searchable Combobox ───────────────────────────────────────────────────────

class SearchableCombobox(ctk.CTkFrame):
    def __init__(self, master, values=None, placeholder="Buscar...", width=300, **kwargs):
        super().__init__(master, fg_color="transparent", **kwargs)
        self.all_values = values or []
        self.selected   = tk.StringVar()
        self._popup     = None
        self._listbox   = None

        self.entry = ctk.CTkEntry(
            self, textvariable=self.selected,
            placeholder_text=placeholder,
            width=width, height=40, font=FONT_INPUT,
            fg_color=C["card"], border_color=C["border"],
            text_color=C["text"]
        )
        self.entry.pack(fill="x")
        self.entry.bind("<KeyRelease>", self._on_key)
        self.entry.bind("<FocusIn>",    self._open)
        self.entry.bind("<FocusOut>",   self._sched_close)
        self.entry.bind("<Down>",       self._focus_list)

    def _filtered(self):
        q = self.selected.get().lower()
        return [v for v in self.all_values if q in v.lower()] if q else self.all_values

    def _on_key(self, e=None):
        if e and e.keysym in ("Return", "Escape"):
            if self._popup: self._popup.destroy(); self._popup = None
            return
        self._open()

    def _open(self, e=None):
        items = self._filtered()
        if not items: return
        if self._popup and self._popup.winfo_exists():
            self._update_list(); return
        x = self.entry.winfo_rootx()
        y = self.entry.winfo_rooty() + self.entry.winfo_height() + 2
        w = self.entry.winfo_width()
        h = min(220, len(items) * 28 + 4)
        self._popup = tk.Toplevel(self)
        self._popup.wm_overrideredirect(True)
        self._popup.geometry(f"{w}x{h}+{x}+{y}")
        self._popup.configure(bg=BLANCO)
        self._popup.attributes("-topmost", True)
        self._listbox = tk.Listbox(
            self._popup, bg=BLANCO, fg=C["text"],
            selectbackground=VERDE, selectforeground=BLANCO,
            font=FONT_INPUT, relief="flat", bd=0,
            highlightthickness=1, highlightcolor=VERDE,
            highlightbackground=C["border"], activestyle="none"
        )
        self._listbox.pack(fill="both", expand=True, padx=1, pady=1)
        self._listbox.bind("<<ListboxSelect>>", self._on_select)
        self._listbox.bind("<FocusOut>", self._sched_close)
        self._update_list()

    def _update_list(self):
        if not self._popup or not self._popup.winfo_exists(): return
        items = self._filtered()
        self._listbox.delete(0, tk.END)
        for item in items: self._listbox.insert(tk.END, item)
        h = min(220, len(items) * 28 + 4)
        x = self.entry.winfo_rootx()
        y = self.entry.winfo_rooty() + self.entry.winfo_height() + 2
        self._popup.geometry(f"{self.entry.winfo_width()}x{h}+{x}+{y}")

    def _on_select(self, e=None):
        if not self._listbox.curselection(): return
        self.selected.set(self._listbox.get(self._listbox.curselection()[0]))
        if self._popup: self._popup.destroy(); self._popup = None

    def _sched_close(self, e=None): self.after(150, self._maybe_close)

    def _maybe_close(self):
        if self._popup and self._popup.winfo_exists():
            try:
                fw = self._popup.focus_get()
                if fw not in (self._listbox, self.entry._entry):
                    self._popup.destroy(); self._popup = None
            except: pass

    def _focus_list(self, e=None):
        if self._popup and self._popup.winfo_exists():
            self._listbox.focus_set()
            if self._listbox.size() > 0: self._listbox.selection_set(0)

    def get(self): return self.selected.get()
    def set(self, v): self.selected.set(v)
    def configure_values(self, values): self.all_values = values


# ── OP Search Widget ──────────────────────────────────────────────────────────

class OPSearchWidget(ctk.CTkFrame):
    """Campo de búsqueda de OP con autocomplete y llenado automático de campos derivados."""

    def __init__(self, master, data_mgr, on_op_selected=None, width=300, **kwargs):
        super().__init__(master, fg_color="transparent", **kwargs)
        self.data_mgr       = data_mgr
        self.on_op_selected = on_op_selected
        self.all_ops        = []
        self._popup         = None
        self._listbox       = None
        self.op_valida      = False
        self._var           = tk.StringVar()

        self.entry = ctk.CTkEntry(
            self, textvariable=self._var,
            placeholder_text="Escribe para buscar OP...",
            width=width, height=40, font=FONT_INPUT,
            fg_color=C["card"], border_color=C["border"],
            text_color=C["text"]
        )
        self.entry.pack(fill="x")

        self.lbl_status = ctk.CTkLabel(
            self, text="", font=FONT_SMALL,
            text_color=C["danger"], fg_color="transparent"
        )
        self.lbl_status.pack(anchor="w", pady=(2, 0))

        self.entry.bind("<KeyRelease>", self._on_key)
        self.entry.bind("<FocusIn>",    self._open)
        self.entry.bind("<FocusOut>",   self._sched_close)
        self.entry.bind("<Down>",       self._focus_list)

        self._refresh_ops()

    def _refresh_ops(self):
        self.all_ops = self.data_mgr.get_lista_ops()

    def _filtered(self):
        q = self._var.get().strip().upper()
        return [op for op in self.all_ops if q in op.upper()] if q else self.all_ops[:50]

    def _on_key(self, e=None):
        if e and e.keysym in ("Return", "Escape"):
            if self._popup: self._popup.destroy(); self._popup = None
            return
        self.op_valida = False
        self._open()

    def _open(self, e=None):
        items = self._filtered()
        if not items: return
        if self._popup and self._popup.winfo_exists():
            self._update_list(); return
        x = self.entry.winfo_rootx()
        y = self.entry.winfo_rooty() + self.entry.winfo_height() + 2
        w = self.entry.winfo_width()
        h = min(220, len(items) * 28 + 4)
        self._popup = tk.Toplevel(self)
        self._popup.wm_overrideredirect(True)
        self._popup.geometry(f"{w}x{h}+{x}+{y}")
        self._popup.configure(bg=BLANCO)
        self._popup.attributes("-topmost", True)
        self._listbox = tk.Listbox(
            self._popup, bg=BLANCO, fg=C["text"],
            selectbackground=VERDE, selectforeground=BLANCO,
            font=FONT_INPUT, relief="flat", bd=0,
            highlightthickness=1, highlightcolor=VERDE,
            highlightbackground=C["border"], activestyle="none"
        )
        self._listbox.pack(fill="both", expand=True, padx=1, pady=1)
        self._listbox.bind("<<ListboxSelect>>", self._on_select)
        self._listbox.bind("<FocusOut>", self._sched_close)
        self._update_list()

    def _update_list(self):
        if not self._popup or not self._popup.winfo_exists(): return
        items = self._filtered()
        self._listbox.delete(0, tk.END)
        for item in items: self._listbox.insert(tk.END, item)
        h = min(220, len(items) * 28 + 4)
        x = self.entry.winfo_rootx()
        y = self.entry.winfo_rooty() + self.entry.winfo_height() + 2
        self._popup.geometry(f"{self.entry.winfo_width()}x{h}+{x}+{y}")

    def _on_select(self, e=None):
        if not self._listbox.curselection(): return
        op = self._listbox.get(self._listbox.curselection()[0])
        self._var.set(op)
        if self._popup: self._popup.destroy(); self._popup = None
        self._aplicar_op(op)

    def _aplicar_op(self, op):
        data = self.data_mgr.buscar_op(op)
        if data:
            self.op_valida = True
            self.entry.configure(border_color=VERDE_BORDE)
            self.lbl_status.configure(text="⚡ OP encontrada", text_color=C["success"])
            if self.on_op_selected:
                self.on_op_selected(op, data)
        else:
            self.op_valida = False
            self.entry.configure(border_color=C["danger"])
            self.lbl_status.configure(text="⚠ OP no encontrada en la base de datos", text_color=C["danger"])

    def _sched_close(self, e=None): self.after(150, self._maybe_close)

    def _maybe_close(self):
        if self._popup and self._popup.winfo_exists():
            try:
                fw = self._popup.focus_get()
                if fw not in (self._listbox, self.entry._entry):
                    self._popup.destroy(); self._popup = None
            except: pass
        # Validar al perder foco
        val = self._var.get().strip()
        if val and not self.op_valida:
            self.entry.configure(border_color=C["danger"])
            self.lbl_status.configure(text="⚠ OP no encontrada en la base de datos", text_color=C["danger"])

    def _focus_list(self, e=None):
        if self._popup and self._popup.winfo_exists():
            self._listbox.focus_set()
            if self._listbox.size() > 0: self._listbox.selection_set(0)

    def get(self): return self._var.get().strip()

    def set(self, v):
        self._var.set(v)
        self.op_valida = False
        self.entry.configure(border_color=C["border"])
        self.lbl_status.configure(text="")

    def reset_error(self):
        self.entry.configure(border_color=C["border"])
        self.lbl_status.configure(text="")


# ── App principal ─────────────────────────────────────────────────────────────

class App(ctk.CTk):

    def __init__(self):
        super().__init__()
        self.config_mgr  = Config()
        self.data_mgr    = DataManager(self.config_mgr)
        ctk.set_appearance_mode("light")

        self.title(NOMBRE_APP)
        self.geometry("1060x820")
        self.minsize(820, 560)
        self.configure(fg_color=C["bg"])

        self.timer_start    = None
        self.timer_running  = False
        self._timer_id      = None
        self.campos_widgets = {}
        self.maquina_info   = {}

        self._build_ui()
        self._load_machine_context()

    # ── Construcción UI ───────────────────────────────────────────────────────

    def _build_ui(self):
        self._build_header()
        self._build_footer()   # footer primero: reserva su espacio y nunca queda oculto
        self._build_body()

    def _build_header(self):
        self.header = tk.Frame(self, bg=VERDE, height=72)
        self.header.pack(fill="x")
        self.header.pack_propagate(False)

        # Logo opcional: si existe logo_header.png se usa, si no, texto
        logo_path = os.path.join(os.path.dirname(__file__), "logo_header.png")
        try:
            img = Image.open(logo_path).convert("RGBA")
            self._logo_img = ctk.CTkImage(light_image=img, dark_image=img,
                                           size=(img.width, img.height))
            tk.Label(self.header, image=self._logo_img,
                     bg=VERDE, bd=0).pack(side="left", padx=20, pady=10)
        except (FileNotFoundError, OSError):
            tk.Label(self.header, text=TEXTO_LOGO,
                     font=("Segoe UI", 18, "bold"),
                     fg=BLANCO, bg=VERDE).pack(side="left", padx=20)
        self.lbl_maquina = tk.Label(self.header, text="",
                                     font=("Segoe UI", 11, "bold"),
                                     fg=VERDE_BORDE, bg=VERDE)
        self.lbl_maquina.pack(side="left", padx=4)

        right = tk.Frame(self.header, bg=VERDE)
        right.pack(side="right", padx=20)

        self.lbl_turno = tk.Label(right, text="",
                                   font=("Segoe UI", 10, "bold"),
                                   fg=VERDE_BORDE, bg=VERDE)
        self.lbl_turno.pack(anchor="e")

        self.lbl_clock = tk.Label(right, text="",
                                   font=("Courier New", 12, "bold"),
                                   fg=BLANCO, bg=VERDE)
        self.lbl_clock.pack(anchor="e")
        self._tick_clock()

    def _build_body(self):
        self.body = ctk.CTkFrame(self, fg_color=C["bg"])
        self.body.pack(fill="both", expand=True, padx=14, pady=(12, 0))
        self.body.columnconfigure(0, weight=3)
        self.body.columnconfigure(1, weight=2)
        self.body.rowconfigure(0, weight=1)

        self.left = ctk.CTkScrollableFrame(
            self.body, fg_color=C["bg"],
            scrollbar_button_color=C["border"],
            scrollbar_button_hover_color=C["accent"]
        )
        self.left.grid(row=0, column=0, sticky="nsew", padx=(0, 10))

        self.right_panel = ctk.CTkFrame(
            self.body, fg_color=C["card"],
            corner_radius=10, border_width=1, border_color=C["border"]
        )
        self.right_panel.grid(row=0, column=1, sticky="nsew")
        self._build_history()

    def _build_footer(self):
        self.footer = tk.Frame(self, bg=BLANCO, height=78,
                                relief="flat", bd=0)
        self.footer.pack(fill="x", side="bottom")
        self.footer.pack_propagate(False)

        # Separador verde
        tk.Frame(self.footer, bg=VERDE, height=3).pack(fill="x", side="top")

        inner = tk.Frame(self.footer, bg=BLANCO)
        inner.pack(fill="both", expand=True, padx=18)

        self.lbl_status = tk.Label(inner, text="",
                                    font=("Segoe UI", 10),
                                    fg=C["subtext"], bg=BLANCO, anchor="w")
        self.lbl_status.pack(side="left", fill="x", expand=True, pady=14)

        self.btn_limpiar = ctk.CTkButton(
            inner, text="🗑  LIMPIAR",
            font=("Segoe UI", 13, "bold"),
            fg_color=ROJO, hover_color="#9b2335",
            text_color=BLANCO, border_width=0,
            height=46, width=150, corner_radius=8,
            command=self._limpiar
        )
        self.btn_limpiar.pack(side="right", pady=12, padx=(10, 0))

        self.btn_guardar = ctk.CTkButton(
            inner, text="💾  GUARDAR REGISTRO",
            font=("Segoe UI", 13, "bold"),
            fg_color=VERDE, hover_color=VERDE_OSC,
            text_color=BLANCO, height=46, width=240, corner_radius=8,
            command=self._guardar
        )
        self.btn_guardar.pack(side="right", pady=12)

    def _build_history(self):
        hdr = ctk.CTkFrame(self.right_panel, fg_color="transparent")
        hdr.pack(fill="x", padx=12, pady=(12, 6))
        ctk.CTkLabel(hdr, text="ÚLTIMOS REGISTROS",
                     font=FONT_TITLE,
                     text_color=C["subtext"]).pack(side="left")
        self.badge = ctk.CTkLabel(hdr, text="0",
                                   font=("Segoe UI", 9, "bold"),
                                   text_color=BLANCO, fg_color=VERDE,
                                   corner_radius=10, width=24, height=18)
        self.badge.pack(side="left", padx=6)

        style = ttk.Style()
        style.theme_use("default")
        style.configure("T.Treeview",
            background=BLANCO, foreground=C["text"],
            fieldbackground=BLANCO, borderwidth=0,
            font=("Segoe UI", 10), rowheight=30)
        style.configure("T.Treeview.Heading",
            background=VERDE_CLARO, foreground=VERDE_OSC,
            font=("Segoe UI", 9, "bold"), borderwidth=0, relief="flat")
        style.map("T.Treeview",
            background=[("selected", VERDE_CLARO)],
            foreground=[("selected", VERDE_OSC)])

        frm = ctk.CTkFrame(self.right_panel, fg_color=BLANCO, corner_radius=8)
        frm.pack(fill="both", expand=True, padx=10, pady=(0, 10))

        cols = ("hora", "op", "evento", "min", "maq")
        self.tree = ttk.Treeview(frm, columns=cols, show="headings", style="T.Treeview")
        for col, lbl, w in [("hora","Hora",60), ("op","OP",70),
                              ("evento","Evento",80), ("min","Min",50), ("maq","Máquina",80)]:
            self.tree.heading(col, text=lbl)
            self.tree.column(col, width=w, anchor="center")
        sb = ttk.Scrollbar(frm, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

    # ── Sección card ──────────────────────────────────────────────────────────

    def _section(self, parent, title, icon=""):
        outer = ctk.CTkFrame(parent, fg_color=C["card"],
                              corner_radius=10, border_width=1,
                              border_color=C["border"])
        outer.pack(fill="x", pady=(0, 10))

        # Barra verde lateral
        tk.Frame(outer, bg=VERDE, width=4).pack(side="left", fill="y")

        wrap = ctk.CTkFrame(outer, fg_color="transparent")
        wrap.pack(side="left", fill="both", expand=True, padx=14, pady=12)

        ctk.CTkLabel(wrap, text=f"{icon}  {title}" if icon else title,
                     font=FONT_TITLE,
                     text_color=C["subtext"]).pack(anchor="w", pady=(0, 10))

        inner = ctk.CTkFrame(wrap, fg_color="transparent")
        inner.pack(fill="both", expand=True)
        return inner

    # ── Formulario dinámico ───────────────────────────────────────────────────

    def _load_machine_context(self):
        cfg = self.config_mgr.get_machine_config()
        self.maquina_info = self.data_mgr.get_maquina_info(cfg.get("maquina", ""))
        nombre = self.maquina_info.get("MAQUINA", "?")
        area   = self.maquina_info.get("AREA", "?")
        self.lbl_maquina.configure(text=f"  {nombre}  ·  {area}")
        self._build_form()
        self._load_history_data()

    def _build_form(self):
        for w in self.left.winfo_children():
            w.destroy()
        self.campos_widgets = {}

        grupo    = self.maquina_info.get("grupo", "general")
        campos   = self.data_mgr.get_campos(grupo)
        maestras = self.data_mgr.maestras

        # ── Cronómetro ────────────────────────────────────────────────────────
        timer_sec = self._section(self.left, "CRONÓMETRO", "⏱")

        self.lbl_timer = ctk.CTkLabel(timer_sec, text="00:00:00",
                                       font=FONT_TIMER,
                                       text_color=VERDE)
        self.lbl_timer.pack(pady=(0, 8))

        # Fechas y horas
        dates_row = ctk.CTkFrame(timer_sec, fg_color="transparent")
        dates_row.pack(fill="x", pady=(0, 10))
        self.var_fecha_ini = tk.StringVar(); self.var_hora_ini  = tk.StringVar()
        self.var_fecha_fin = tk.StringVar(); self.var_hora_fin  = tk.StringVar()
        self.var_minutos   = tk.StringVar()

        for i, (lbl, var) in enumerate([
            ("Fecha Ini", self.var_fecha_ini), ("Hora Ini", self.var_hora_ini),
            ("Fecha Fin", self.var_fecha_fin), ("Hora Fin", self.var_hora_fin),
            ("Minutos",   self.var_minutos)
        ]):
            c = ctk.CTkFrame(dates_row, fg_color=GRIS_BG, corner_radius=6,
                              border_width=1, border_color=C["border"])
            c.grid(row=0, column=i, padx=3, sticky="ew")
            dates_row.columnconfigure(i, weight=1)
            ctk.CTkLabel(c, text=lbl, font=("Segoe UI", 8),
                         text_color=C["subtext"]).pack(pady=(4, 0))
            ctk.CTkLabel(c, textvariable=var,
                         font=("Courier New", 10, "bold"),
                         text_color=VERDE).pack(pady=(0, 4))

        # Botones cronómetro — grandes y claros
        btns = ctk.CTkFrame(timer_sec, fg_color="transparent")
        btns.pack(pady=(4, 0))

        self.btn_start = ctk.CTkButton(
            btns, text="▶  INICIAR",
            font=("Segoe UI", 13, "bold"),
            fg_color=VERDE, hover_color=VERDE_OSC,
            text_color=BLANCO, width=160, height=44,
            corner_radius=8, command=self._timer_start
        )
        self.btn_start.grid(row=0, column=0, padx=8)

        self.btn_stop = ctk.CTkButton(
            btns, text="⏹  DETENER",
            font=("Segoe UI", 13, "bold"),
            fg_color=ROJO, hover_color="#9b2335",
            text_color=BLANCO, width=160, height=44,
            corner_radius=8, state="disabled",
            command=self._timer_stop
        )
        self.btn_stop.grid(row=0, column=1, padx=8)

        # Hint
        self.lbl_form_hint = ctk.CTkLabel(
            timer_sec,
            text="▶  Inicia el cronómetro para habilitar los campos",
            font=("Segoe UI", 11, "bold"),
            text_color=AMARILLO,
            fg_color=GRIS_BG,
            corner_radius=6
        )
        self.lbl_form_hint.pack(fill="x", pady=(10, 0), ipady=8)

        # ── Datos del registro ────────────────────────────────────────────────
        form_sec = self._section(self.left, "DATOS DEL REGISTRO", "📋")
        form_sec.columnconfigure(1, weight=1)

        for i, campo in enumerate(campos):
            ctk.CTkLabel(form_sec, text=campo["label"], font=FONT_LABEL,
                         text_color=C["subtext"], anchor="w", width=160
                         ).grid(row=i, column=0, sticky="w", pady=6)

            tipo = campo["tipo"]; name = campo["name"]; widget = None

            if tipo == "op_search":
                widget = OPSearchWidget(
                    form_sec,
                    data_mgr=self.data_mgr,
                    on_op_selected=self._on_op_selected,
                    width=380
                )
                widget.grid(row=i, column=1, sticky="ew", pady=6)

            elif tipo == "readonly":
                var = tk.StringVar()
                widget = ctk.CTkEntry(
                    form_sec, textvariable=var, font=FONT_INPUT, height=40,
                    fg_color=C["readonly"], border_color=C["border"],
                    text_color=C["subtext"], state="disabled"
                )
                widget.grid(row=i, column=1, sticky="ew", pady=6)
                widget._var = var

            elif tipo == "combo":
                ops = []
                if name == "operario":
                    ops = [f"{r['ID']} - {r['OPERARIO']}" for r in maestras.get("operarios", [])]
                elif name == "evento":
                    ops = [f"{r['COD EVENTO']} - {r['EVENTO']}" for r in maestras.get("eventos", [])]
                widget = SearchableCombobox(form_sec, values=ops,
                                            placeholder=campo["label"], width=380)
                widget.grid(row=i, column=1, sticky="ew", pady=6)
                if name == "evento":
                    widget.selected.trace_add("write", lambda *a: self._on_evento_changed())

            elif tipo == "entry":
                var = tk.StringVar()
                widget = ctk.CTkEntry(
                    form_sec, textvariable=var, font=FONT_INPUT, height=40,
                    fg_color=C["card"], border_color=C["border"],
                    text_color=C["text"]
                )
                widget.grid(row=i, column=1, sticky="ew", pady=6)
                widget._var = var
                # Campos dependientes de TIRAJE arrancan deshabilitados con 0
                if name in ("cantidad", "cavidad", "planchas"):
                    var.set("0")
                    widget.configure(state="disabled", fg_color=C["readonly"])

            elif tipo == "text":
                widget = ctk.CTkTextbox(
                    form_sec, font=FONT_INPUT, height=70,
                    fg_color=C["card"], border_color=C["border"],
                    border_width=2, text_color=C["text"]
                )
                widget.grid(row=i, column=1, sticky="ew", pady=6)

            if widget:
                self.campos_widgets[name] = (widget, tipo)
                if tipo not in ("op_search", "readonly"):
                    self._set_field_enabled(widget, tipo, False)

        # Deshabilitar OP search también al inicio
        if "op" in self.campos_widgets:
            w, t = self.campos_widgets["op"]
            w.entry.configure(state="disabled")

    def _on_evento_changed(self):
        """Habilita/deshabilita cantidad, cavidad y planchas según el evento seleccionado."""
        if "evento" not in self.campos_widgets: return

        evento_widget, _ = self.campos_widgets["evento"]
        val = evento_widget.get().strip().upper()
        es_tiraje = val.startswith("P02")

        campos_tiraje = ["cantidad", "cavidad", "planchas"]
        for name in campos_tiraje:
            if name not in self.campos_widgets: continue
            widget, _ = self.campos_widgets[name]
            if es_tiraje:
                widget.configure(state="normal", fg_color=C["card"],
                                  border_color=C["border"])
                if hasattr(widget, "_var") and widget._var.get() == "0":
                    widget._var.set("")
            else:
                if hasattr(widget, "_var"):
                    widget._var.set("0")
                widget.configure(state="disabled", fg_color=C["readonly"],
                                  border_color=C["border"])

    def _on_op_selected(self, op, data):
        """Callback cuando se selecciona una OP válida — autocompleta campos."""
        mapping = {
            "tipo_trabajo":   data.get("tipo_trabajo", ""),
            "cod_referencia": data.get("cod_referencia", ""),
            "referencia":     data.get("referencia", ""),
            "nit_cliente":    data.get("nit_cliente", ""),
            "cliente":        data.get("cliente", ""),
        }
        for name, val in mapping.items():
            if name in self.campos_widgets:
                widget, tipo = self.campos_widgets[name]
                if tipo == "readonly" and hasattr(widget, "_var"):
                    widget._var.set(val)

    # ── Habilitar/Deshabilitar campos ─────────────────────────────────────────

    def _set_field_enabled(self, widget, tipo, enabled: bool):
        state = "normal" if enabled else "disabled"
        try:
            if tipo == "combo":
                widget.entry.configure(state=state)
            elif tipo in ("entry", "text"):
                widget.configure(state=state)
        except: pass

    def _enable_fields(self):
        for name, (widget, tipo) in self.campos_widgets.items():
            if tipo == "op_search":
                widget.entry.configure(state="normal")
            elif tipo != "readonly":
                self._set_field_enabled(widget, tipo, True)
        try:
            self.lbl_form_hint.configure(
                text="✓  Cronómetro iniciado — completa los campos",
                text_color=C["success"]
            )
        except: pass

    def _disable_fields(self):
        for name, (widget, tipo) in self.campos_widgets.items():
            if tipo == "op_search":
                widget.entry.configure(state="disabled")
            elif tipo != "readonly":
                self._set_field_enabled(widget, tipo, False)
        try:
            self.lbl_form_hint.configure(
                text="▶  Inicia el cronómetro para habilitar los campos",
                text_color=AMARILLO
            )
        except: pass

    # ── Cronómetro ────────────────────────────────────────────────────────────

    def _timer_start(self):
        self.timer_start   = datetime.now()
        self.timer_running = True
        self.var_fecha_ini.set(self.timer_start.strftime("%Y-%m-%d"))
        self.var_hora_ini.set(self.timer_start.strftime("%H:%M:%S"))
        self.var_fecha_fin.set(""); self.var_hora_fin.set(""); self.var_minutos.set("")
        self.btn_start.configure(state="disabled", fg_color=C["border"], text_color=C["subtext"])
        self.btn_stop.configure(state="normal", fg_color=ROJO, text_color=BLANCO)
        self._enable_fields()
        self._tick_timer()

    def _timer_stop(self):
        if not self.timer_running: return
        fin = datetime.now(); self.timer_running = False
        if self._timer_id: self.after_cancel(self._timer_id)
        delta = fin - self.timer_start
        mins  = delta.total_seconds() / 60
        self.var_fecha_fin.set(fin.strftime("%Y-%m-%d"))
        self.var_hora_fin.set(fin.strftime("%H:%M:%S"))
        self.var_minutos.set(f"{mins:.2f}")
        self.lbl_timer.configure(text=self._fmt(delta), text_color=C["success"])
        self.btn_start.configure(state="normal", fg_color=VERDE, text_color=BLANCO)
        self.btn_stop.configure(state="disabled")

    def _tick_timer(self):
        if self.timer_running:
            self.lbl_timer.configure(text=self._fmt(datetime.now() - self.timer_start))
            self._timer_id = self.after(500, self._tick_timer)

    @staticmethod
    def _fmt(delta):
        t = int(delta.total_seconds()); h, r = divmod(t, 3600); m, s = divmod(r, 60)
        return f"{h:02d}:{m:02d}:{s:02d}"

    def _tick_clock(self):
        now = datetime.now()
        self.lbl_clock.configure(text=now.strftime("%d/%m/%Y  %H:%M:%S"))
        h = now.hour
        turno = "Turno Mañana" if 6 <= h < 14 else "Turno Tarde" if 14 <= h < 22 else "Turno Noche"
        self.lbl_turno.configure(text=turno)
        self.after(1000, self._tick_clock)

    # ── Guardar ───────────────────────────────────────────────────────────────

    CAMPOS_ENTERO   = {"cantidad", "cavidad", "planchas"}
    CAMPOS_NUMERICO = {"calibre", "ancho", "corte"}

    def _guardar(self):
        if not self.var_fecha_ini.get():
            self._show_error("⏱  Debes iniciar el cronómetro antes de guardar.")
            return
        if self.timer_running:
            self._show_error("⏹  Debes detener el cronómetro antes de guardar.")
            return

        # Validar OP
        if "op" in self.campos_widgets:
            op_widget, _ = self.campos_widgets["op"]
            if not op_widget.op_valida:
                self._show_error("La OP ingresada no existe en la base de datos.")
                return

        faltantes = self._validar_campos()
        if faltantes:
            return

        registro = {
            "FECHA_INICIAL": self.var_fecha_ini.get(),
            "HORA_INICIAL":  self.var_hora_ini.get(),
            "FECHA_FINAL":   self.var_fecha_fin.get(),
            "HORA_FINAL":    self.var_hora_fin.get(),
            "MINUTOS":       self.var_minutos.get(),
            "MAQUINA":       self.maquina_info.get("COD MAQUINA", ""),
        }

        grupo  = self.maquina_info.get("grupo", "")
        campos = self.data_mgr.get_campos(grupo)

        for campo in campos:
            name = campo["name"]
            if name not in self.campos_widgets: continue
            widget, tipo = self.campos_widgets[name]
            if tipo == "op_search":         val = widget.get()
            elif tipo == "readonly":        val = widget._var.get() if hasattr(widget, "_var") else ""
            elif tipo == "combo":           val = widget.get()
            elif tipo == "entry":           val = widget._var.get()
            elif tipo == "text":            val = widget.get("1.0", "end-1c")
            else:                           val = ""
            registro[campo["col_excel"]] = val

        self.btn_guardar.configure(state="disabled", text="Guardando...")
        threading.Thread(target=self._save_thread, args=(registro,), daemon=True).start()

    def _validar_campos(self):
        grupo  = self.maquina_info.get("grupo", "")
        campos = self.data_mgr.get_campos(grupo)
        errores = []

        for name, (widget, tipo) in self.campos_widgets.items():
            if tipo not in ("readonly",):
                self._set_field_error(widget, tipo, False)

        # Determinar si el evento es TIRAJE
        es_tiraje = False
        if "evento" in self.campos_widgets:
            ev_widget, _ = self.campos_widgets["evento"]
            es_tiraje = ev_widget.get().strip().upper().startswith("P02")

        for campo in campos:
            name = campo["name"]
            if name in ("observacion", "tipo_trabajo", "cod_referencia",
                        "referencia", "nit_cliente", "cliente"): continue
            # Campos dependientes de TIRAJE: solo validar si es tiraje
            if name in ("cantidad", "cavidad", "planchas") and not es_tiraje: continue
            if name not in self.campos_widgets: continue
            widget, tipo = self.campos_widgets[name]

            if tipo == "op_search":    val = widget.get()
            elif tipo == "combo":      val = widget.get()
            elif tipo == "entry":      val = widget._var.get() if hasattr(widget, "_var") else ""
            else:                      continue

            val = val.strip()
            if not val:
                errores.append(f"{campo['label']}: campo requerido")
                self._set_field_error(widget, tipo, True)
                continue

            if name in self.CAMPOS_ENTERO:
                try:
                    n = int(val)
                    if n <= 0: raise ValueError
                except ValueError:
                    errores.append(f"{campo['label']}: debe ser un número entero positivo")
                    self._set_field_error(widget, tipo, True)
                    continue

            if name in self.CAMPOS_NUMERICO:
                try:
                    float(val)
                except ValueError:
                    errores.append(f"{campo['label']}: debe ser un valor numérico")
                    self._set_field_error(widget, tipo, True)
                    continue

        if errores:
            self._show_error(errores[0])
        return errores

    def _set_field_error(self, widget, tipo, error: bool):
        color = C["danger"] if error else C["border"]
        try:
            if tipo == "combo":
                widget.entry.configure(border_color=color)
            elif tipo == "entry":
                widget.configure(border_color=color)
            elif tipo == "op_search":
                widget.entry.configure(border_color=color)
        except: pass

    def _show_error(self, msg):
        self.lbl_status.configure(text=f"⚠  {msg}", fg=C["danger"])

    def _save_thread(self, registro):
        try:
            self.data_mgr.guardar_registro(registro)
            self.after(0, self._on_save_ok, registro)
        except ConnectionError as e:
            self.after(0, self._on_save_err, str(e))
        except Exception as e:
            self.after(0, self._on_save_err, str(e))

    def _on_save_ok(self, registro):
        self.btn_guardar.configure(state="normal", text="💾  GUARDAR REGISTRO")
        evento  = registro.get("EVENTO", "")
        minutos = registro.get("MINUTOS", "")
        self.tree.insert("", 0, values=(
            datetime.now().strftime("%H:%M"),
            registro.get("OP", ""),
            evento[:10] if evento else "",
            f"{float(minutos):.1f}" if minutos else "",
            self.maquina_info.get("MAQUINA", "")[:10]
        ))
        self.badge.configure(text=str(len(self.tree.get_children())))
        self._limpiar()
        self.lbl_status.configure(
            text=f"✓  Guardado correctamente a las {datetime.now().strftime('%H:%M:%S')}",
            fg=C["success"]
        )

    def _on_save_err(self, err):
        self.btn_guardar.configure(state="normal", text="💾  GUARDAR REGISTRO")
        messagebox.showerror("Error al guardar", f"No se pudo guardar:\n\n{err}")

    def _limpiar(self):
        for name, (widget, tipo) in self.campos_widgets.items():
            if tipo == "op_search":
                widget.set("")
                widget.reset_error()
            elif tipo == "readonly":
                if hasattr(widget, "_var"): widget._var.set("")
            elif tipo == "combo":
                widget.set("")
                self._set_field_error(widget, tipo, False)
            elif tipo == "entry":
                if name in ("cantidad", "cavidad", "planchas"):
                    widget._var.set("0")
                    widget.configure(state="disabled", fg_color=C["readonly"],
                                      border_color=C["border"])
                else:
                    widget._var.set("")
                self._set_field_error(widget, tipo, False)
            elif tipo == "text":
                widget.delete("1.0", tk.END)

        self.timer_running = False
        if self._timer_id: self.after_cancel(self._timer_id)
        self.lbl_timer.configure(text="00:00:00", text_color=VERDE)
        self.timer_start = None
        for v in [self.var_fecha_ini, self.var_hora_ini,
                  self.var_fecha_fin, self.var_hora_fin, self.var_minutos]:
            v.set("")
        self.btn_start.configure(state="normal", fg_color=VERDE, text_color=BLANCO)
        self.btn_stop.configure(state="disabled")
        self._disable_fields()

    def _load_history_data(self):
        registros = self.data_mgr.get_history(limit=50)
        for r in registros:
            self.tree.insert("", tk.END, values=(
                str(r.get("Hora Ini", ""))[:5], r.get("OP", ""),
                str(r.get("Evento", ""))[:10], str(r.get("Minutos", ""))[:5],
                str(r.get("Maquina", ""))
            ))
        self.badge.configure(text=str(len(registros)))


# ── Ventana configuración ─────────────────────────────────────────────────────

class ConfigWindow(ctk.CTk):

    def __init__(self, config_mgr):
        super().__init__()
        self.config_mgr = config_mgr
        ctk.set_appearance_mode("light")

        self.title(f"Configuración inicial · {NOMBRE_APP}")
        self.geometry("520x480")
        self.configure(fg_color=BLANCO)
        self.resizable(False, False)

        # Header verde
        hdr = tk.Frame(self, bg=VERDE, height=80)
        hdr.pack(fill="x")
        hdr.pack_propagate(False)
        logo_path = os.path.join(os.path.dirname(__file__), "logo_header.png")
        try:
            img = Image.open(logo_path).convert("RGBA")
            w2 = int(img.width * 1.2); h2 = int(img.height * 1.2)
            img2 = img.resize((w2, h2), Image.LANCZOS)
            self._logo = ctk.CTkImage(light_image=img2, dark_image=img2, size=(w2, h2))
            tk.Label(hdr, image=self._logo, bg=VERDE, bd=0).pack(pady=12)
        except (FileNotFoundError, OSError):
            tk.Label(hdr, text=TEXTO_LOGO, font=("Segoe UI", 22, "bold"),
                     fg=BLANCO, bg=VERDE).pack(pady=20)

        body = ctk.CTkFrame(self, fg_color=BLANCO)
        body.pack(fill="both", expand=True, padx=30, pady=20)

        ctk.CTkLabel(body, text="Configuración de máquina",
                     font=("Segoe UI", 16, "bold"),
                     text_color=GRIS_TEXTO).pack(pady=(0, 4))
        ctk.CTkLabel(body, text="Selecciona la máquina asignada a este PC",
                     font=FONT_SMALL, text_color=C["subtext"]).pack(pady=(0, 20))

        dm       = DataManager(config_mgr)
        maquinas = [f"{r['COD MAQUINA']} - {r['MAQUINA']}" for r in dm.maestras.get("maquinas", [])]

        ctk.CTkLabel(body, text="Máquina", font=FONT_LABEL,
                     text_color=C["subtext"]).pack(anchor="w")
        self.combo_maq = SearchableCombobox(body, values=maquinas,
                                             placeholder="Buscar máquina...", width=420)
        self.combo_maq.pack(fill="x", pady=(4, 20))

        ctk.CTkButton(
            body, text="GUARDAR Y CONTINUAR", font=FONT_BTN,
            fg_color=VERDE, hover_color=VERDE_OSC, text_color=BLANCO,
            width=260, height=44, corner_radius=8, command=self._guardar
        ).pack(pady=(10, 0))

    def _guardar(self):
        maq = self.combo_maq.get()
        if not maq:
            messagebox.showwarning("Requerido", "Debes seleccionar una máquina.")
            return
        cod = maq.split(" - ")[0].strip()
        self.config_mgr.save_machine_config(cod)
        self.destroy()


# ── Entry point ───────────────────────────────────────────────────────────────

def main():
    config_mgr = Config()
    if not config_mgr.is_configured():
        # Verificar red antes de mostrar configuración
        try:
            DataManager(config_mgr)
        except ConnectionError as e:
            root = tk.Tk(); root.withdraw()
            messagebox.showerror("Sin conexión a la red", str(e))
            root.destroy(); return
        cw = ConfigWindow(config_mgr)
        cw.mainloop()
        if not config_mgr.is_configured():
            return
    try:
        App().mainloop()
    except ConnectionError as e:
        messagebox.showerror("Sin conexión a la red", str(e))

if __name__ == "__main__":
    main()
