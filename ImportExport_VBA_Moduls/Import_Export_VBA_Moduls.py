"""
Import_Export_VBA_Moduls.py
Exporta E importa modulos VBA de/hacia archivos .xlsm
Requiere: Python 3.9+, pywin32, tkinter
"""

import datetime
import os
import threading
import tkinter as tk
from tkinter import filedialog, font as tkfont, messagebox, scrolledtext, ttk

import pythoncom
import win32api
import win32com.client
import win32con
import win32gui
import win32process
import winreg

# ── Registro Windows ──────────────────────────────────────────────────────────
_REG_KEY = r"Software\ImportExportVBAModuls"

def _reg_read(name: str) -> str:
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, _REG_KEY) as k:
            v, _ = winreg.QueryValueEx(k, name)
            return v
    except OSError:
        return ""

def _reg_write(name: str, value: str) -> None:
    try:
        with winreg.CreateKey(winreg.HKEY_CURRENT_USER, _REG_KEY) as k:
            winreg.SetValueEx(k, name, 0, winreg.REG_SZ, value)
    except OSError:
        pass

# ── Centro de confianza Excel ─────────────────────────────────────────────────
def _vba_access_enabled() -> bool:
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Office") as base:
            i = 0
            while True:
                try:
                    ver = winreg.EnumKey(base, i); i += 1
                    try: float(ver)
                    except ValueError: continue
                    try:
                        with winreg.OpenKey(winreg.HKEY_CURRENT_USER,
                                rf"Software\Microsoft\Office\{ver}\Excel\Security") as sec:
                            val, _ = winreg.QueryValueEx(sec, "AccessVBOM")
                            if val == 1: return True
                    except OSError: pass
                except OSError: break
    except OSError: pass
    return False

# ── Constantes VBA ────────────────────────────────────────────────────────────
_COMP_EXT   = {1: ".bas", 2: ".cls", 3: ".frm", 100: ".cls"}
_COMP_TAG   = {1: "Módulo", 2: "Clase", 3: "Formulario", 100: "Documento"}
_TYPE_ORDER = {1: 0, 2: 1, 3: 2, 100: 3}

# ── COM helpers ───────────────────────────────────────────────────────────────
def _xl_open_ro(path: str):
    xl = win32com.client.DispatchEx("Excel.Application")
    xl.Visible = False; xl.DisplayAlerts = False
    xl.ScreenUpdating = False; xl.EnableEvents = False
    wb = xl.Workbooks.Open(os.path.abspath(path), ReadOnly=True, UpdateLinks=False)
    return xl, wb

def _xl_open_rw(path: str):
    xl = win32com.client.DispatchEx("Excel.Application")
    xl.Visible = False; xl.DisplayAlerts = False
    xl.ScreenUpdating = False; xl.EnableEvents = False
    wb = xl.Workbooks.Open(os.path.abspath(path), ReadOnly=False, UpdateLinks=False)
    return xl, wb

def _xl_close(xl, wb) -> None:
    try:
        if wb: wb.Close(False)
    except Exception: pass
    try:
        if xl: xl.Quit()
    except Exception: pass

# ── win32 — rellena dialogo de contrasena de Excel ───────────────────────────
def _fill_excel_vba_dialog(xl_pid: int, password: str, timeout: float = 10.0) -> bool:
    import time
    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        found_hwnd = None
        def _enum(hwnd, _):
            nonlocal found_hwnd
            if found_hwnd or not win32gui.IsWindowVisible(hwnd): return True
            try: _, pid = win32process.GetWindowThreadProcessId(hwnd)
            except Exception: return True
            if pid != xl_pid: return True
            title = win32gui.GetWindowText(hwnd).lower()
            cls   = win32gui.GetClassName(hwnd)
            if any(kw in title for kw in ("contrasena", "password", "kennwort",
                                           "mot de passe", "wachtwoord")):
                found_hwnd = hwnd
            elif cls == "#32770" and win32gui.FindWindowEx(hwnd, None, "Edit", None):
                found_hwnd = hwnd
            return True
        win32gui.EnumWindows(_enum, None)
        if found_hwnd:
            edit = win32gui.FindWindowEx(found_hwnd, None, "Edit", None)
            if edit:
                win32gui.SendMessage(edit, win32con.WM_SETTEXT, 0, password)
                import time as _t; _t.sleep(0.05)
                win32api.keybd_event(win32con.VK_RETURN, 0, 0, 0)
                win32api.keybd_event(win32con.VK_RETURN, 0, win32con.KEYEVENTF_KEYUP, 0)
                return True
        time.sleep(0.12)
    return False

def _unlock_vbproject(xl, xl_pid: int, password: str) -> None:
    xl.Visible = True
    try: xl.WindowState = -4140
    except Exception: pass
    try: xl.VBE.MainWindow.Visible = True
    except Exception: pass
    filled = threading.Event()
    def _f():
        if _fill_excel_vba_dialog(xl_pid, password): filled.set()
    threading.Thread(target=_f, daemon=True).start()
    try:
        _ = xl.ActiveWorkbook.VBProject.VBComponents
        filled.wait(timeout=12.0)
    except Exception as e:
        xl.Visible = False
        s = str(e)
        if any(c in s for c in ("-2146827284", "-2147352567")):
            raise RuntimeError(
                "Excel no pudo desbloquear el proyecto automaticamente.\n\n"
                "Alternativa: abre el archivo en Excel, desprotege el proyecto VBA\n"
                "en Herramientas > Propiedades > Proteccion, guarda y reintenta."
            ) from None
        raise
    xl.Visible = False

# ── Extraer codigo VBA de un archivo exportado ────────────────────────────────
def _extract_code(filepath: str) -> str:
    try:
        with open(filepath, "r", encoding="utf-8", errors="ignore") as f:
            lines = f.readlines()
    except OSError:
        return ""
    last_attr = -1
    for i, ln in enumerate(lines):
        if ln.startswith("Attribute "):
            last_attr = i
    return "".join(lines[last_attr + 1:] if last_attr >= 0 else lines).rstrip()

# ── Firma personal ────────────────────────────────────────────────────────────
_FIRMA_B64 = "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"

# ── Sistema de fuentes escalable ─────────────────────────────────────────────
BASE       = 14
FONT_UI    = ("Verdana", BASE)
FONT_BOLD  = ("Verdana", BASE, "bold")
FONT_TITLE = ("Verdana", BASE + 1, "bold")
FONT_ENTRY = ("Verdana", BASE - 1)
FONT_LOG   = ("Courier New", BASE)

# ── Colores (WCAG 2.1) ────────────────────────────────────────────────────────
BG_APP      = "#F2F3F4"
BG_ARCHIVO  = "#2874A6"   # azul     4.98:1
BG_MODULOS  = "#7D3C98"   # violeta  7.14:1
BG_DESTINO  = "#1E8449"   # verde    4.69:1
BG_IMPORTAR = "#117A65"   # teal     5.25:1
BG_LOG_HDR  = "#5D6D7E"   # gris     5.20:1

C_SELEC    = "#D5D8DC"
C_TODOS    = "#AED6F1"
C_NINGUNO  = "#F5B7B1"
C_EXPORTAR = "#A9DFBF"
C_IMPORTAR = "#A2D9CE"
C_ACCION   = "#F0B27A"


# ── Dialogo de contrasena ─────────────────────────────────────────────────────
class _PasswordDialog(tk.Toplevel):
    def __init__(self, parent, filename: str, bg_hdr: str = None) -> None:
        super().__init__(parent)
        bg_hdr = bg_hdr or BG_ARCHIVO
        self.title("Proyecto VBA protegido")
        self.configure(bg=BG_APP)
        self.resizable(False, False)
        self.grab_set()
        self.password: str | None = None
        self.update_idletasks()
        pw = parent.winfo_rootx() + parent.winfo_width() // 2
        ph = parent.winfo_rooty() + parent.winfo_height() // 2
        self.geometry(f"+{pw - 220}+{ph - 110}")
        tk.Frame(self, bg=bg_hdr, height=4).pack(fill="x")
        tk.Label(self, text=f"  \U0001f512  {filename}",
                 bg=bg_hdr, fg="white", font=FONT_BOLD,
                 padx=14, pady=8, anchor="w").pack(fill="x")
        body = tk.Frame(self, bg=BG_APP, padx=20, pady=14)
        body.pack(fill="both", expand=True)
        body.columnconfigure(1, weight=1)
        tk.Label(body, text="El proyecto VBA esta protegido.",
                 bg=BG_APP, font=FONT_UI, anchor="w").grid(
            row=0, column=0, columnspan=2, sticky="w", pady=(0, 10))
        tk.Label(body, text="Contrasena:", bg=BG_APP, font=FONT_UI).grid(
            row=1, column=0, sticky="w", padx=(0, 10), pady=4)
        self._entry = ttk.Entry(body, show="●", font=FONT_UI, width=26)
        self._entry.grid(row=1, column=1, sticky="ew", pady=4)
        self._entry.focus_set()
        self._entry.bind("<Return>", lambda _: self._ok())
        self._entry.bind("<Escape>", lambda _: self._cancel())
        self._pw_vis = False
        def _toggle():
            self._pw_vis = not self._pw_vis
            self._entry.configure(show="" if self._pw_vis else "●")
            _eye.configure(text="\U0001f648" if self._pw_vis else "\U0001f441")
        _eye = tk.Button(body, text="\U0001f441", command=_toggle,
                         bg=BG_APP, activebackground=BG_APP,
                         relief="flat", bd=0, cursor="hand2",
                         font=("Segoe UI Emoji", BASE + 2))
        _eye.grid(row=1, column=2, padx=(4, 0), pady=4)
        ttk.Separator(self).pack(fill="x", padx=14, pady=(4, 0))
        frm = tk.Frame(self, bg=BG_APP, pady=8)
        frm.pack(fill="x", padx=14)
        ttk.Button(frm, text="Cancelar", command=self._cancel).pack(side="right", padx=(6, 0))
        ttk.Button(frm, text="Aceptar", style="Export.TButton", command=self._ok).pack(side="right")
        self.protocol("WM_DELETE_WINDOW", self._cancel)
        self.wait_window()

    def _ok(self) -> None:
        self.password = self._entry.get() or None; self.destroy()
    def _cancel(self) -> None:
        self.destroy()


# ── Tooltip ───────────────────────────────────────────────────────────────────
class _Tooltip:
    def __init__(self, widget: tk.Widget, text: str) -> None:
        self._w = widget; self._text = text; self._win = None
        widget.bind("<Enter>", self._show)
        widget.bind("<Leave>", self._hide)
        widget.bind("<ButtonPress>", self._hide)

    def _show(self, _=None) -> None:
        if self._win: return
        x = self._w.winfo_rootx() + self._w.winfo_width() // 2
        y = self._w.winfo_rooty() + self._w.winfo_height() + 4
        T = "#f0f0f1"
        self._win = tk.Toplevel(self._w)
        self._win.wm_overrideredirect(True)
        self._win.wm_geometry(f"+{x}+{y}")
        self._win.configure(bg=T)
        try: self._win.wm_attributes("-transparentcolor", T)
        except Exception: pass
        tk.Label(self._win, text=self._text, bg=T, fg="#555555",
                 relief="flat", bd=0,
                 font=("Verdana", BASE - 2, "italic"), padx=3, pady=1).pack()

    def _hide(self, _=None) -> None:
        if self._win: self._win.destroy(); self._win = None


# ── Frame con scroll ──────────────────────────────────────────────────────────
class _ScrollFrame(ttk.Frame):
    def __init__(self, parent, **kw):
        super().__init__(parent, **kw)
        self._c = tk.Canvas(self, borderwidth=0, highlightthickness=0, bg=BG_APP)
        vsb = ttk.Scrollbar(self, orient="vertical", command=self._c.yview)
        self._c.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")
        self._c.pack(side="left", fill="both", expand=True)
        self.inner = tk.Frame(self._c, bg=BG_APP)
        self._wid = self._c.create_window((0, 0), window=self.inner, anchor="nw")
        self.inner.bind("<Configure>",
            lambda e: self._c.configure(scrollregion=self._c.bbox("all")))
        self._c.bind("<Configure>", lambda e: self._c.itemconfig(self._wid, width=e.width))
        self._c.bind("<MouseWheel>", self._w)
        self.inner.bind("<MouseWheel>", self._w)

    def _w(self, e): self._c.yview_scroll(-1 if e.delta > 0 else 1, "units")
    def bind_child_wheel(self, w): w.bind("<MouseWheel>", self._w)


# ── Aplicacion ────────────────────────────────────────────────────────────────
class App(tk.Tk):

    W, H = 1400, 820

    def __init__(self):
        super().__init__()
        self.title("Import / Export VBA Moduls")
        self.resizable(True, True)
        self.minsize(900, 640)
        self.configure(bg=BG_APP)

        # Estado EXPORTAR
        self._exp_xlsm_var = tk.StringVar()
        self._exp_dest_var  = tk.StringVar()
        self._exp_modules: list[tuple[str, int, str]] = []  # (name, typ, ext)
        self._exp_checked: dict[str, bool] = {}
        self._exp_sort_keys: list[tuple[str, bool]] = []
        self._exp_anchor: str | None = None
        self._exp_anchor_state: bool = True
        self._exp_password: str | None = None

        # Estado IMPORTAR
        self._imp_xlsm_var   = tk.StringVar()
        self._imp_folder_var = tk.StringVar()
        self._imp_rows: list[tuple[str, str, str, float, bool]] = []  # (fname, ext, tag, mtime, warn_no_frx)
        self._imp_checked: dict[str, bool] = {}
        self._imp_sort_keys: list[tuple[str, bool]] = []
        self._imp_anchor: str | None = None
        self._imp_anchor_state: bool = True
        self._imp_password: str | None = None
        self._imp_backup_var = tk.BooleanVar(value=True)

        self._firma_img = tk.PhotoImage(data=_FIRMA_B64)
        self._build_styles()
        self._build_ui()
        self._center(self.W, self.H)

        if not _vba_access_enabled():
            self._show_warn_banner()

        last_exp = _reg_read("LastXlsmExport")
        if last_exp and os.path.exists(last_exp):
            self._exp_xlsm_var.set(last_exp)
            self._lbl_exp_title.config(text=os.path.basename(last_exp))
            self._propose_dest(last_exp)
            self._load_exp_modules_bg(last_exp)

        last_imp = _reg_read("LastXlsmImport")
        if last_imp and os.path.exists(last_imp):
            self._imp_xlsm_var.set(last_imp)
            self._lbl_imp_title.config(text=os.path.basename(last_imp))

        last_fld = _reg_read("LastImportFolder")
        if last_fld and os.path.isdir(last_fld):
            self._imp_folder_var.set(last_fld)
            self._lbl_fld_title.config(text=os.path.basename(last_fld))
            self._populate_imp_files(last_fld)

        # Restaurar última pestaña activa
        last_tab = _reg_read("LastActiveTab")
        if last_tab in ("exp", "imp"):
            self._switch_tab(last_tab)

    def _center(self, w, h):
        self.update_idletasks()
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        self.geometry(f"{w}x{h}+{(sw-w)//2}+{max(0,(sh-h)//2)}")

    # ── Estilos ───────────────────────────────────────────────────────────────
    def _build_styles(self):
        s = ttk.Style(self)
        s.theme_use("clam")
        s.configure(".",              font=FONT_UI, background=BG_APP)
        s.configure("TFrame",         background=BG_APP)
        s.configure("TCheckbutton",   background=BG_APP, font=FONT_UI)
        s.map("TCheckbutton",         background=[("active", BG_APP)])
        s.configure("TEntry",         font=FONT_UI, fieldbackground="white")
        s.configure("TScrollbar",     background=BG_APP)
        s.configure("Sel.TButton",     font=FONT_UI,   background=C_SELEC)
        s.map("Sel.TButton",           background=[("active", "#BFC9CA")])
        s.configure("Todos.TButton",   font=FONT_UI,   background=C_TODOS)
        s.map("Todos.TButton",         background=[("active", "#85C1E9")])
        s.configure("Ninguno.TButton", font=FONT_UI,   background=C_NINGUNO)
        s.map("Ninguno.TButton",       background=[("active", "#F1948A")])
        s.configure("Export.TButton",  font=FONT_BOLD, background=C_EXPORTAR)
        s.map("Export.TButton",        background=[("active", "#7DCEA0")])
        s.configure("Import.TButton",  font=FONT_BOLD, background=C_IMPORTAR)
        s.map("Import.TButton",        background=[("active", "#76D7C4")])
        s.configure("Accion.TButton",  font=FONT_UI,   background=C_ACCION)
        s.map("Accion.TButton",        background=[("active", "#E59866")])
        s.configure("Treeview",         font=FONT_UI, rowheight=30,
                    background="white", fieldbackground="white", borderwidth=0)
        s.configure("Treeview.Heading", font=FONT_BOLD, background="#D5D8DC",
                    relief="flat", padding=(4, 4))
        s.map("Treeview",
              background=[("selected", "#AED6F1")],
              foreground=[("selected", "#1A1A1A")])

    # ── Bloque coloreado (parent explicito) ───────────────────────────────────
    def _bloque(self, parent, titulo: str, bg_titulo: str,
                expand: bool = False) -> tuple[tk.Frame, tk.Frame]:
        outer = tk.Frame(parent, bg=BG_APP, bd=1, relief="solid",
                         highlightbackground="#AAAAAA", highlightthickness=1)
        outer.pack(fill="both" if expand else "x", expand=expand, pady=(0, 6))
        tb = tk.Frame(outer, bg=bg_titulo)
        tb.pack(fill="x")
        tk.Label(tb, text=f"  {titulo}",
                 bg=bg_titulo, fg="white", font=FONT_TITLE, anchor="w", pady=2
                 ).grid(row=0, column=0, sticky="w")
        c = tk.Frame(outer, bg=BG_APP, padx=10, pady=6)
        c.pack(fill="both", expand=True)
        c.columnconfigure(1, weight=1)
        return tb, c

    # ── UI ────────────────────────────────────────────────────────────────────
    def _build_ui(self):

        # ── Banner advertencia VBA (oculto por defecto) ───────────────────────
        self._frm_warn = tk.Frame(self, bg="#FFF3CD", bd=1, relief="solid")
        tk.Label(self._frm_warn,
                 text="  El acceso al modelo de objetos VBA no esta activado en Excel.",
                 bg="#FFF3CD", fg="#7d5800", font=FONT_BOLD,
                 padx=10, pady=6).pack(side="left", fill="x", expand=True)
        tk.Button(self._frm_warn, text="Como activarlo  >",
                  bg="#FFEEBA", fg="#3d2b00", activebackground="#FFD966",
                  relief="flat", bd=1, padx=8, pady=3, font=FONT_UI, cursor="hand2",
                  command=self._show_vba_instructions
                  ).pack(side="right", padx=(0, 8), pady=4)

        # ── Barra de pestañas estilizada ──────────────────────────────────────
        self._tab_bar = tk.Frame(self, bg=BG_APP)
        self._tab_bar.pack(fill="x", padx=10, pady=(10, 0))

        self._btn_tab_exp = tk.Button(
            self._tab_bar,
            text="   EXPORTAR MÓDULOS   ",
            command=lambda: self._switch_tab("exp"),
            font=FONT_BOLD, fg="white", bg=BG_ARCHIVO,
            activebackground=BG_ARCHIVO, activeforeground="white",
            relief="flat", bd=0, padx=18, pady=10, cursor="arrow")
        self._btn_tab_exp.pack(side="left", padx=(0, 3))

        self._btn_tab_imp = tk.Button(
            self._tab_bar,
            text="   IMPORTAR MÓDULOS   ",
            command=lambda: self._switch_tab("imp"),
            font=FONT_UI, fg="#546E7A", bg="#B8C1CC",
            activebackground=BG_IMPORTAR, activeforeground="white",
            relief="flat", bd=0, padx=18, pady=10, cursor="hand2")
        self._btn_tab_imp.pack(side="left")

        # Línea indicadora de pestaña activa (3 px, cambia de color)
        self._tab_indicator = tk.Frame(self, height=3, bg=BG_ARCHIVO)
        self._tab_indicator.pack(fill="x", padx=10)

        # ── Área de contenido (contenedor fijo) ───────────────────────────────
        self._content_area = tk.Frame(self, bg=BG_APP)
        self._content_area.pack(fill="both", expand=True, padx=10, pady=(4, 0))

        # ── Frame EXPORTAR ────────────────────────────────────────────────────
        self._frm_exp = tk.Frame(self._content_area, bg=BG_APP)

        tb_a, c_a = self._bloque(self._frm_exp, "ARCHIVO .xlsm", BG_ARCHIVO)
        tb_a.columnconfigure(1, weight=1)
        self._lbl_exp_title = tk.Label(
            tb_a, text="", bg=BG_ARCHIVO, fg="white", font=FONT_UI, anchor="w")
        self._lbl_exp_title.grid(row=0, column=1, sticky="ew", padx=(6, 8))
        _fl = tk.Label(tb_a, image=self._firma_img, bg=BG_ARCHIVO, bd=0, cursor="hand2")
        _fl.grid(row=0, column=2, padx=(0, 8), pady=2)
        _Tooltip(_fl, "Dugarry")
        tk.Label(c_a, text="Archivo.Xlsm:", bg=BG_APP, font=FONT_UI, anchor="w").grid(
            row=0, column=0, sticky="w", padx=(0, 8), pady=3)
        tk.Entry(c_a, textvariable=self._exp_xlsm_var, state="readonly",
                 font=FONT_ENTRY, bg="white", readonlybackground="#ECECEC",
                 fg="#222222", relief="sunken", bd=1).grid(
            row=0, column=1, sticky="ew", pady=3, ipady=4)
        ttk.Button(c_a, text="Seleccionar", style="Sel.TButton",
                   command=self._pick_exp_file).grid(
            row=0, column=2, padx=(6, 0), pady=3, ipadx=4)

        tb_m, c_m = self._bloque(self._frm_exp, "MÓDULOS VBA", BG_MODULOS, expand=True)
        tb_m.columnconfigure(0, weight=1)
        self._exp_count_lbl = tk.Label(tb_m, text="", bg=BG_MODULOS, fg="white", font=FONT_UI)
        self._exp_count_lbl.grid(row=0, column=1, padx=(0, 4))
        ttk.Button(tb_m, text="Todos",   style="Todos.TButton",  command=self._exp_sel_all
                   ).grid(row=0, column=2, padx=(0, 3), pady=3, ipadx=3)
        ttk.Button(tb_m, text="Ninguno", style="Ninguno.TButton", command=self._exp_sel_none
                   ).grid(row=0, column=3, padx=(0, 6), pady=3, ipadx=3)
        exp_tv_frame = tk.Frame(c_m, bg=BG_APP)
        exp_tv_frame.pack(fill="both", expand=True)
        exp_tv_frame.columnconfigure(0, weight=1)
        exp_tv_frame.rowconfigure(0, weight=1)

        self._exp_tv_cols = ["sel", "nombre", "clase"]
        self._exp_tv = ttk.Treeview(
            exp_tv_frame,
            columns=tuple(self._exp_tv_cols),
            show="headings",
            selectmode="browse",
        )
        self._exp_tv.heading("sel",    text="☑",      anchor="center")
        self._exp_tv.heading("nombre", text="Nombre",  anchor="w")
        self._exp_tv.heading("clase",  text="Clase",   anchor="center")
        self._exp_tv.column("sel",    width=40,  minwidth=40,  stretch=False, anchor="center")
        self._exp_tv.column("nombre", width=300, minwidth=120, stretch=True,  anchor="w")
        self._exp_tv.column("clase",  width=110, minwidth=80,  stretch=False, anchor="center")

        vsb_exp = ttk.Scrollbar(exp_tv_frame, orient="vertical", command=self._exp_tv.yview)
        self._exp_tv.configure(yscrollcommand=vsb_exp.set)
        self._exp_tv.grid(row=0, column=0, sticky="nsew")
        vsb_exp.grid(row=0, column=1, sticky="ns")

        self._exp_tv.bind("<ButtonRelease-1>", self._exp_on_click)
        self._exp_tv.bind("<space>", self._exp_on_space)
        self._exp_tv.tag_configure("row_a", background="white")
        self._exp_tv.tag_configure("row_b", background="#EBEBEB")

        tb_d, c_d = self._bloque(self._frm_exp, "CARPETA DE DESTINO", BG_DESTINO)
        tb_d.columnconfigure(1, weight=1)
        self._lbl_dest_title = tk.Label(
            tb_d, text="", bg=BG_DESTINO, fg="white", font=FONT_UI, anchor="w")
        self._lbl_dest_title.grid(row=0, column=1, sticky="ew", padx=(6, 8))
        tk.Label(c_d, text="Carpeta:", bg=BG_APP, font=FONT_UI, anchor="w").grid(
            row=0, column=0, sticky="w", padx=(0, 8), pady=3)
        tk.Entry(c_d, textvariable=self._exp_dest_var,
                 font=FONT_ENTRY, bg="white", fg="#222222",
                 relief="sunken", bd=1).grid(
            row=0, column=1, sticky="ew", pady=3, ipady=4)
        ttk.Button(c_d, text="Examinar", style="Sel.TButton",
                   command=self._pick_exp_dest).grid(
            row=0, column=2, padx=(6, 0), pady=3, ipadx=4)

        # ── Frame IMPORTAR ────────────────────────────────────────────────────
        self._frm_imp = tk.Frame(self._content_area, bg=BG_APP)

        tb_ia, c_ia = self._bloque(self._frm_imp, "XLSM DESTINO", BG_IMPORTAR)
        tb_ia.columnconfigure(1, weight=1)
        self._lbl_imp_title = tk.Label(
            tb_ia, text="", bg=BG_IMPORTAR, fg="white", font=FONT_UI, anchor="w")
        self._lbl_imp_title.grid(row=0, column=1, sticky="ew", padx=(6, 8))
        tk.Label(c_ia, text="Archivo.Xlsm:", bg=BG_APP, font=FONT_UI, anchor="w").grid(
            row=0, column=0, sticky="w", padx=(0, 8), pady=3)
        tk.Entry(c_ia, textvariable=self._imp_xlsm_var, state="readonly",
                 font=FONT_ENTRY, bg="white", readonlybackground="#ECECEC",
                 fg="#222222", relief="sunken", bd=1).grid(
            row=0, column=1, sticky="ew", pady=3, ipady=4)
        ttk.Button(c_ia, text="Seleccionar", style="Sel.TButton",
                   command=self._pick_imp_file).grid(
            row=0, column=2, padx=(6, 0), pady=3, ipadx=4)

        tb_if, c_if = self._bloque(self._frm_imp, "CARPETA FUENTE", BG_IMPORTAR)
        tb_if.columnconfigure(1, weight=1)
        self._lbl_fld_title = tk.Label(
            tb_if, text="", bg=BG_IMPORTAR, fg="white", font=FONT_UI, anchor="w")
        self._lbl_fld_title.grid(row=0, column=1, sticky="ew", padx=(6, 8))
        tk.Label(c_if, text="Carpeta:", bg=BG_APP, font=FONT_UI, anchor="w").grid(
            row=0, column=0, sticky="w", padx=(0, 8), pady=3)
        tk.Entry(c_if, textvariable=self._imp_folder_var, state="readonly",
                 font=FONT_ENTRY, bg="white", readonlybackground="#ECECEC",
                 fg="#222222", relief="sunken", bd=1).grid(
            row=0, column=1, sticky="ew", pady=3, ipady=4)
        ttk.Button(c_if, text="Examinar", style="Sel.TButton",
                   command=self._pick_imp_folder).grid(
            row=0, column=2, padx=(6, 0), pady=3, ipadx=4)

        tb_im, c_im = self._bloque(self._frm_imp, "MÓDULOS A IMPORTAR", BG_IMPORTAR, expand=True)
        tb_im.columnconfigure(0, weight=1)
        self._imp_count_lbl = tk.Label(tb_im, text="", bg=BG_IMPORTAR, fg="white", font=FONT_UI)
        self._imp_count_lbl.grid(row=0, column=1, padx=(0, 4))
        ttk.Button(tb_im, text="Todos",   style="Todos.TButton",  command=self._imp_sel_all
                   ).grid(row=0, column=2, padx=(0, 3), pady=3, ipadx=3)
        ttk.Button(tb_im, text="Ninguno", style="Ninguno.TButton", command=self._imp_sel_none
                   ).grid(row=0, column=3, padx=(0, 6), pady=3, ipadx=3)
        tv_frame = tk.Frame(c_im, bg=BG_APP)
        tv_frame.pack(fill="both", expand=True)
        tv_frame.columnconfigure(0, weight=1)
        tv_frame.rowconfigure(0, weight=1)

        self._imp_tv_cols = ["sel", "nombre", "clase", "fecha"]
        self._imp_tv = ttk.Treeview(
            tv_frame,
            columns=tuple(self._imp_tv_cols),
            show="headings",
            selectmode="browse",
        )
        self._imp_tv.heading("sel",    text="☑",         anchor="center")
        self._imp_tv.heading("nombre", text="Nombre",     anchor="w")
        self._imp_tv.heading("clase",  text="Clase",      anchor="center")
        self._imp_tv.heading("fecha",  text="Modificado", anchor="center")
        self._imp_tv.column("sel",    width=40,  minwidth=40,  stretch=False, anchor="center")
        self._imp_tv.column("nombre", width=280, minwidth=120, stretch=True,  anchor="w")
        self._imp_tv.column("clase",  width=110, minwidth=80,  stretch=False, anchor="center")
        self._imp_tv.column("fecha",  width=165, minwidth=120, stretch=False, anchor="center")

        vsb_imp = ttk.Scrollbar(tv_frame, orient="vertical", command=self._imp_tv.yview)
        self._imp_tv.configure(yscrollcommand=vsb_imp.set)
        self._imp_tv.grid(row=0, column=0, sticky="nsew")
        vsb_imp.grid(row=0, column=1, sticky="ns")

        self._imp_tv.bind("<ButtonRelease-1>", self._imp_on_click)
        self._imp_tv.bind("<space>", self._imp_on_space)
        self._imp_tv.tag_configure("row_a", background="white")
        self._imp_tv.tag_configure("row_b", background="#EBEBEB")

        tk.Checkbutton(c_im,
                       text="  Copia de seguridad antes de importar",
                       variable=self._imp_backup_var,
                       bg=BG_APP, font=FONT_UI, anchor="w",
                       activebackground=BG_APP, pady=4
                       ).pack(fill="x", pady=(4, 0))

        # ── REGISTRO (siempre visible, ancho completo) ────────────────────────
        reg_outer = tk.Frame(self, bg=BG_APP, bd=1, relief="solid",
                             highlightbackground="#AAAAAA", highlightthickness=1)
        reg_outer.pack(fill="x", padx=10, pady=(0, 6))
        tb_log = tk.Frame(reg_outer, bg=BG_LOG_HDR)
        tb_log.pack(fill="x")
        tb_log.columnconfigure(0, weight=1)
        tk.Label(tb_log, text="  REGISTRO",
                 bg=BG_LOG_HDR, fg="white", font=FONT_TITLE, anchor="w", pady=2
                 ).grid(row=0, column=0, sticky="w")
        ttk.Button(tb_log, text="Limpiar", style="Accion.TButton",
                   command=self._log_clear).grid(
            row=0, column=1, padx=(0, 6), pady=3, ipadx=5)
        self._btn_exp = ttk.Button(tb_log, text="Exportar Módulos",
                                   style="Export.TButton", command=self._export)
        self._btn_exp.grid(row=0, column=2, padx=(0, 6), pady=3, ipadx=5)
        self._btn_imp = ttk.Button(tb_log, text="Importar Módulos",
                                   style="Import.TButton", command=self._import)
        self._btn_imp.grid(row=0, column=3, padx=(0, 8), pady=3, ipadx=5)

        c_log = tk.Frame(reg_outer, bg=BG_APP, padx=6, pady=4)
        c_log.pack(fill="x")
        c_log.columnconfigure(0, weight=1)
        self._log = scrolledtext.ScrolledText(
            c_log, height=10, state="disabled", font=FONT_LOG, wrap="word",
            spacing1=5, spacing3=5,
            bg="#1E1E1E", fg="#D4D4D4", insertbackground="white",
            relief="flat", bd=0)
        self._log.grid(row=0, column=0, sticky="ew")
        self._log.tag_configure("ok",    foreground="#4EC94E")
        self._log.tag_configure("error", foreground="#FF6B6B")
        self._log.tag_configure("info",  foreground="#85C1E9")
        self._log.tag_configure("warn",  foreground="#F0B27A")

        tk.Frame(self, bg=BG_APP, height=6).pack()

        # Activar pestaña inicial
        self._switch_tab("exp")

    # ── Cambio de pestaña ─────────────────────────────────────────────────────
    def _switch_tab(self, mode: str) -> None:
        """Muestra el frame de la pestaña seleccionada y actualiza el estilo."""
        self._active_tab = mode
        _reg_write("LastActiveTab", mode)

        # Ocultar ambos, luego mostrar el activo
        self._frm_exp.pack_forget()
        self._frm_imp.pack_forget()

        if mode == "exp":
            self._frm_exp.pack(fill="both", expand=True)
            # Pestaña EXPORTAR — activa
            self._btn_tab_exp.configure(
                bg=BG_ARCHIVO, fg="white", font=FONT_BOLD, cursor="arrow",
                pady=10)
            # Pestaña IMPORTAR — inactiva
            self._btn_tab_imp.configure(
                bg="#B8C1CC", fg="#546E7A", font=FONT_UI, cursor="hand2",
                pady=8)
            self._tab_indicator.configure(bg=BG_ARCHIVO)
        else:
            self._frm_imp.pack(fill="both", expand=True)
            # Pestaña EXPORTAR — inactiva
            self._btn_tab_exp.configure(
                bg="#B8C1CC", fg="#546E7A", font=FONT_UI, cursor="hand2",
                pady=8)
            # Pestaña IMPORTAR — activa
            self._btn_tab_imp.configure(
                bg=BG_IMPORTAR, fg="white", font=FONT_BOLD, cursor="arrow",
                pady=10)

    # ── Banner y ayuda VBA ────────────────────────────────────────────────────
    def _show_warn_banner(self):
        self._frm_warn.pack(fill="x", padx=10, pady=(6, 0),
                            before=self._tab_bar)

    def _show_vba_instructions(self):
        win = tk.Toplevel(self)
        win.title("Activar acceso VBA"); win.configure(bg=BG_APP)
        win.resizable(False, False); win.grab_set()
        win.update_idletasks()
        pw = self.winfo_rootx() + self.winfo_width() // 2
        ph = self.winfo_rooty() + self.winfo_height() // 2
        win.geometry(f"+{pw-260}+{ph-200}")
        tk.Frame(win, bg="#856404", height=4).pack(fill="x")
        hdr = tk.Frame(win, bg="#FFF3CD"); hdr.pack(fill="x")
        tk.Label(hdr, text="  Acceso al modelo VBA desactivado",
                 bg="#FFF3CD", fg="#7d5800", font=FONT_BOLD,
                 padx=14, pady=10).pack(side="left")
        body = tk.Frame(win, bg=BG_APP, padx=16, pady=12); body.pack(fill="both", expand=True)
        tk.Label(body,
                 text=("Para usar esta herramienta activa el acceso VBA en Excel:\n\n"
                       "  1.  Archivo -> Opciones -> Centro de confianza\n"
                       "  2.  Configuracion del Centro de confianza...\n"
                       "  3.  Panel: Configuracion de macros\n"
                       "  4.  Activar: Confiar en el acceso al modelo de objetos VBA\n"
                       "  5.  Aceptar y reiniciar la herramienta"),
                 justify="left", bg=BG_APP, font=FONT_UI).pack(anchor="w")
        ttk.Separator(win).pack(fill="x", padx=14)
        ttk.Button(win, text="  Entendido  ", command=win.destroy).pack(
            side="right", padx=14, pady=10)

    # ── Log ───────────────────────────────────────────────────────────────────
    def _log_write(self, text: str, tag: str = "") -> None:
        self._log.configure(state="normal")
        self._log.insert("end", text, tag)
        self._log.see("end")
        self._log.configure(state="disabled")

    def _log_clear(self) -> None:
        self._log.configure(state="normal")
        self._log.delete("1.0", "end")
        self._log.configure(state="disabled")

    # ════════════════════════════════════════════════════════════════════════
    #  EXPORTAR
    # ════════════════════════════════════════════════════════════════════════

    def _pick_exp_file(self) -> None:
        init = (os.path.dirname(self._exp_xlsm_var.get())
                if self._exp_xlsm_var.get() else os.path.expanduser("~"))
        path = filedialog.askopenfilename(
            parent=self, title="Seleccionar .xlsm para exportar", initialdir=init,
            filetypes=[("Excel con macros", "*.xlsm"), ("Todos", "*.*")])
        if not path: return
        self._exp_xlsm_var.set(path)
        self._lbl_exp_title.config(text=os.path.basename(path))
        self._exp_password = None
        _reg_write("LastXlsmExport", path)
        self._propose_dest(path)
        self._load_exp_modules_bg(path)

    def _propose_dest(self, xlsm_path: str) -> None:
        folder = os.path.dirname(xlsm_path)
        stem   = os.path.splitext(os.path.basename(xlsm_path))[0]
        dest   = os.path.join(folder, stem + "_VBA-Moduls")
        self._exp_dest_var.set(dest)
        self._lbl_dest_title.config(text=os.path.basename(dest))

    def _pick_exp_dest(self) -> None:
        init = (os.path.dirname(self._exp_xlsm_var.get())
                if self._exp_xlsm_var.get() else os.path.expanduser("~"))
        folder = filedialog.askdirectory(parent=self, title="Carpeta de destino", initialdir=init)
        if folder:
            self._exp_dest_var.set(folder)
            self._lbl_dest_title.config(text=os.path.basename(folder) or folder)

    def _exp_sel_all(self) -> None:
        for name, *_ in self._exp_modules:
            self._exp_checked[name] = True
            try: self._exp_tv.set(name, "sel", "☑")
            except Exception: pass
        self._exp_update_count()

    def _exp_sel_none(self) -> None:
        for name, *_ in self._exp_modules:
            self._exp_checked[name] = False
            try: self._exp_tv.set(name, "sel", "☐")
            except Exception: pass
        self._exp_update_count()

    def _exp_update_count(self) -> None:
        tot = len(self._exp_modules)
        sel = sum(1 for name, *_ in self._exp_modules if self._exp_checked.get(name, True))
        self._exp_count_lbl.config(text=f"{sel} / {tot}" if tot else "")

    def _load_exp_modules_bg(self, path: str) -> None:
        self._log_write(f"\nLeyendo módulos de:  {os.path.basename(path)}\n", "info")
        self._btn_exp.config(state="disabled")
        self._clear_exp_checks()
        threading.Thread(target=self._load_exp_thread, args=(path,), daemon=True).start()

    def _load_exp_thread(self, path: str) -> None:
        pythoncom.CoInitialize()
        xl = wb = None
        try:
            xl, wb = _xl_open_ro(path)
            if wb.VBProject.Protection == 1:
                pwd_r: dict = {}; evt = threading.Event()
                def _ask():
                    dlg = _PasswordDialog(self, os.path.basename(path), BG_ARCHIVO)
                    pwd_r["pw"] = dlg.password; evt.set()
                self.after(0, _ask); evt.wait()
                pw = pwd_r.get("pw")
                if not pw:
                    self.after(0, lambda: (self._btn_exp.config(state="normal"),
                                           self._log_write("\n  Cancelado.\n", "info")))
                    return
                self._exp_password = pw
                _, xl_pid = win32process.GetWindowThreadProcessId(xl.Hwnd)
                _unlock_vbproject(xl, xl_pid, pw)
            mods = sorted(
                [(c.Name, c.Type, _COMP_EXT[c.Type])
                 for c in wb.VBProject.VBComponents if c.Type in _COMP_EXT],
                key=lambda m: (_TYPE_ORDER[m[1]], m[0].lower()))
            self.after(0, self._populate_exp, mods)
        except Exception as exc:
            self.after(0, self._on_exp_load_error, str(exc))
        finally:
            _xl_close(xl, wb); pythoncom.CoUninitialize()

    def _populate_exp(self, mods: list) -> None:
        self._clear_exp_checks()
        for name, typ, ext in mods:
            self._exp_modules.append((name, typ, ext))
            self._exp_checked[name] = True
        self._exp_refresh_treeview()
        self._exp_autosize_columns()
        self._exp_update_count()
        self._log_write(f"  -> {len(mods)} módulo(s) encontrado(s)\n", "ok")
        self._btn_exp.config(state="normal")

    def _on_exp_load_error(self, msg: str) -> None:
        self._btn_exp.config(state="normal")
        if not _vba_access_enabled(): self._show_warn_banner()
        self._log_write(f"\n  ERROR: {msg}\n", "error")

    def _clear_exp_checks(self) -> None:
        for iid in self._exp_tv.get_children():
            self._exp_tv.delete(iid)
        self._exp_modules = []
        self._exp_checked = {}
        self._exp_sort_keys = []
        self._exp_anchor = None
        self._exp_anchor_state = True
        self._exp_count_lbl.config(text="")
        self._exp_update_headers()

    def _exp_on_click(self, event) -> None:
        region = self._exp_tv.identify_region(event.x, event.y)
        col    = self._exp_tv.identify_column(event.x)
        try:
            col_idx = int(col[1:]) - 1
        except (ValueError, IndexError):
            return
        col_id = (self._exp_tv_cols[col_idx]
                  if 0 <= col_idx < len(self._exp_tv_cols) else None)
        if not col_id:
            return
        if region == "heading" and col_id != "sel":
            shift = bool(event.state & 0x0001)
            self._exp_sort_by(col_id, shift)
        elif region == "cell":
            iid = self._exp_tv.identify_row(event.y)
            if iid:
                shift = bool(event.state & 0x0001)
                if shift and self._exp_anchor:
                    self._exp_range_check(iid)
                else:
                    self._exp_toggle_check(iid)
                    self._exp_anchor       = iid
                    self._exp_anchor_state = self._exp_checked.get(iid, True)

    def _exp_toggle_check(self, iid: str) -> None:
        if not iid or iid not in self._exp_checked:
            return
        self._exp_checked[iid] = not self._exp_checked[iid]
        sym = "☑" if self._exp_checked[iid] else "☐"
        try:
            self._exp_tv.set(iid, "sel", sym)
        except Exception:
            pass
        self._exp_update_count()

    def _exp_on_space(self, event) -> None:
        iid = self._exp_tv.focus()
        if iid:
            self._exp_toggle_check(iid)
            self._exp_anchor       = iid
            self._exp_anchor_state = self._exp_checked.get(iid, True)

    def _exp_range_check(self, end_iid: str) -> None:
        children = self._exp_tv.get_children()
        try:
            i1 = children.index(self._exp_anchor)
            i2 = children.index(end_iid)
        except ValueError:
            self._exp_toggle_check(end_iid)
            return
        start, stop = min(i1, i2), max(i1, i2) + 1
        for iid in children[start:stop]:
            if iid in self._exp_checked:
                self._exp_checked[iid] = self._exp_anchor_state
                sym = "☑" if self._exp_anchor_state else "☐"
                try: self._exp_tv.set(iid, "sel", sym)
                except Exception: pass
        self._exp_update_count()

    def _exp_sort_by(self, col_id: str, shift: bool) -> None:
        if shift:
            idx = next((i for i, (c, _) in enumerate(self._exp_sort_keys)
                        if c == col_id), -1)
            if idx >= 0:
                _, rev = self._exp_sort_keys[idx]
                self._exp_sort_keys[idx] = (col_id, not rev)
            else:
                self._exp_sort_keys.append((col_id, False))
        else:
            idx = next((i for i, (c, _) in enumerate(self._exp_sort_keys)
                        if c == col_id), -1)
            if idx == 0 and len(self._exp_sort_keys) == 1:
                _, rev = self._exp_sort_keys[0]
                self._exp_sort_keys[0] = (col_id, not rev)
            else:
                self._exp_sort_keys = [(col_id, False)]
        self._exp_refresh_treeview()
        self._exp_update_headers()

    def _exp_refresh_treeview(self) -> None:
        rows = list(self._exp_modules)
        for col_id, rev in reversed(self._exp_sort_keys):
            if col_id == "nombre":
                rows.sort(key=lambda r: r[0].lower(), reverse=rev)
            elif col_id == "clase":
                rows.sort(key=lambda r: _COMP_TAG.get(r[1], "").lower(), reverse=rev)
        for iid in self._exp_tv.get_children():
            self._exp_tv.delete(iid)
        for i, (name, typ, ext) in enumerate(rows):
            checked = self._exp_checked.get(name, True)
            sym     = "☑" if checked else "☐"
            row_tag = "row_a" if i % 2 == 0 else "row_b"
            self._exp_tv.insert("", "end", iid=name,
                                values=(sym, name + ext, _COMP_TAG.get(typ, "?")),
                                tags=(row_tag,))

    def _exp_update_headers(self) -> None:
        col_labels = {"nombre": "Nombre", "clase": "Clase"}
        for col_id, base in col_labels.items():
            idx = next((i for i, (c, _) in enumerate(self._exp_sort_keys)
                        if c == col_id), -1)
            if idx >= 0:
                _, rev = self._exp_sort_keys[idx]
                arrow = " ↓" if rev else " ↑"
                pri   = str(idx + 1) if len(self._exp_sort_keys) > 1 else ""
                label = base + arrow + pri
            else:
                label = base
            self._exp_tv.heading(col_id, text=label)

    def _exp_autosize_columns(self) -> None:
        f_ui   = tkfont.Font(font=FONT_UI)
        f_bold = tkfont.Font(font=FONT_BOLD)
        pad    = 24
        heading_text = self._exp_tv.heading("clase")["text"]
        max_w = f_bold.measure(heading_text) + pad
        for iid in self._exp_tv.get_children():
            w = f_ui.measure(str(self._exp_tv.set(iid, "clase"))) + pad
            if w > max_w:
                max_w = w
        self._exp_tv.column("clase", width=max_w, minwidth=max_w)

    def _export(self) -> None:
        path = self._exp_xlsm_var.get(); dest = self._exp_dest_var.get().strip()
        if not path or not os.path.exists(path):
            messagebox.showerror("Error", "Selecciona un .xlsm valido.", parent=self); return
        selected = [(n, t, e) for n, t, e in self._exp_modules if self._exp_checked.get(n, True)]
        if not selected:
            messagebox.showwarning("Sin seleccion", "Selecciona al menos un modulo.", parent=self); return
        if not dest:
            messagebox.showerror("Error", "Especifica una carpeta.", parent=self); return
        if os.path.exists(dest):
            files = [f for f in os.listdir(dest) if os.path.isfile(os.path.join(dest, f))]
            if files:
                resp = messagebox.askyesnocancel(
                    "Carpeta no vacia",
                    f"La carpeta contiene {len(files)} archivo(s).\n\n"
                    "Si = Borrar y exportar\nNo = Sobreescribir\nCancelar = Salir",
                    parent=self)
                if resp is None: return
                if resp:
                    for f in files:
                        try: os.remove(os.path.join(dest, f))
                        except OSError: pass
        else:
            try: os.makedirs(dest, exist_ok=True)
            except OSError as exc:
                messagebox.showerror("Error", f"No se pudo crear la carpeta:\n{exc}", parent=self); return
        self._btn_exp.config(state="disabled")
        self._log_write(f"\nExportando {len(selected)} modulo(s) a:\n  {dest}\n", "info")
        sel_names = {n for n, _, _ in selected}
        threading.Thread(target=self._export_thread,
                         args=(path, dest, sel_names), daemon=True).start()

    def _export_thread(self, path: str, dest: str, names: set) -> None:
        pythoncom.CoInitialize()
        xl = wb = None
        try:
            xl, wb = _xl_open_ro(path)
            if wb.VBProject.Protection == 1 and self._exp_password:
                _, xl_pid = win32process.GetWindowThreadProcessId(xl.Hwnd)
                _unlock_vbproject(xl, xl_pid, self._exp_password)
            count = 0
            for comp in wb.VBProject.VBComponents:
                if comp.Name in names and comp.Type in _COMP_EXT:
                    out = os.path.join(dest, comp.Name + _COMP_EXT[comp.Type])
                    comp.Export(out)
                    self.after(0, self._log_write, f"  -> {comp.Name}{_COMP_EXT[comp.Type]}\n", "ok")
                    count += 1
            self.after(0, self._export_done, count, dest)
        except Exception as exc:
            self.after(0, lambda m=str(exc): (
                self._btn_exp.config(state="normal"),
                self._log_write(f"\n  ERROR: {m}\n", "error")))
        finally:
            _xl_close(xl, wb); pythoncom.CoUninitialize()

    def _export_done(self, count: int, dest: str) -> None:
        self._btn_exp.config(state="normal")
        self._log_write(f"\n  Exportacion completada: {count} modulo(s)\n", "ok")
        messagebox.showinfo("Exportacion completada",
                            f"Se exportaron {count} modulo(s) en:\n\n{dest}", parent=self)

    # ════════════════════════════════════════════════════════════════════════
    #  IMPORTAR
    # ════════════════════════════════════════════════════════════════════════

    def _pick_imp_file(self) -> None:
        init = (os.path.dirname(self._imp_xlsm_var.get())
                if self._imp_xlsm_var.get() else os.path.expanduser("~"))
        path = filedialog.askopenfilename(
            parent=self, title="Seleccionar .xlsm destino", initialdir=init,
            filetypes=[("Excel con macros", "*.xlsm"), ("Todos", "*.*")])
        if not path: return
        self._imp_xlsm_var.set(path)
        self._lbl_imp_title.config(text=os.path.basename(path))
        self._imp_password = None
        _reg_write("LastXlsmImport", path)

    def _pick_imp_folder(self) -> None:
        init = (self._imp_folder_var.get() if self._imp_folder_var.get()
                else os.path.expanduser("~"))
        folder = filedialog.askdirectory(
            parent=self, title="Carpeta con módulos .bas/.cls/.frm", initialdir=init)
        if not folder: return
        self._imp_folder_var.set(folder)
        self._lbl_fld_title.config(text=os.path.basename(folder) or folder)
        _reg_write("LastImportFolder", folder)
        self._populate_imp_files(folder)

    def _populate_imp_files(self, folder: str) -> None:
        self._clear_imp_checks()
        exts = {".bas", ".cls", ".frm"}
        EXT_TAG = {".bas": "Módulo", ".cls": "Clase/Doc", ".frm": "Formulario"}
        try:
            fnames = [f for f in os.listdir(folder)
                      if os.path.splitext(f)[1].lower() in exts
                      and os.path.isfile(os.path.join(folder, f))]
        except OSError:
            fnames = []
        for fname in fnames:
            ext  = os.path.splitext(fname)[1].lower()
            tag  = EXT_TAG.get(ext, "?")
            path = os.path.join(folder, fname)
            try:
                mtime = os.path.getmtime(path)
            except OSError:
                mtime = 0.0
            warn = False
            if ext == ".frm":
                frx = os.path.join(folder, os.path.splitext(fname)[0] + ".frx")
                warn = not os.path.exists(frx)
            self._imp_rows.append((fname, ext, tag, mtime, warn))
            self._imp_checked[fname] = True
        self._imp_refresh_treeview()
        self._imp_autosize_columns()
        self._imp_update_count()
        msg = (f"\n  {len(fnames)} archivo(s) en: {os.path.basename(folder)}\n"
               if fnames else f"\n  No hay .bas/.cls/.frm en: {folder}\n")
        self._log_write(msg, "info" if fnames else "warn")

    def _imp_sel_all(self) -> None:
        for fname, *_ in self._imp_rows:
            self._imp_checked[fname] = True
            try: self._imp_tv.set(fname, "sel", "☑")
            except Exception: pass
        self._imp_update_count()

    def _imp_sel_none(self) -> None:
        for fname, *_ in self._imp_rows:
            self._imp_checked[fname] = False
            try: self._imp_tv.set(fname, "sel", "☐")
            except Exception: pass
        self._imp_update_count()

    def _imp_update_count(self) -> None:
        tot = len(self._imp_rows)
        sel = sum(1 for fname, *_ in self._imp_rows if self._imp_checked.get(fname, True))
        self._imp_count_lbl.config(text=f"{sel} / {tot}" if tot else "")

    def _clear_imp_checks(self) -> None:
        for iid in self._imp_tv.get_children():
            self._imp_tv.delete(iid)
        self._imp_rows = []
        self._imp_checked = {}
        self._imp_sort_keys = []
        self._imp_anchor = None
        self._imp_anchor_state = True
        self._imp_count_lbl.config(text="")
        self._imp_update_headers()

    def _imp_on_click(self, event) -> None:
        region = self._imp_tv.identify_region(event.x, event.y)
        col    = self._imp_tv.identify_column(event.x)
        try:
            col_idx = int(col[1:]) - 1
        except (ValueError, IndexError):
            return
        col_id = (self._imp_tv_cols[col_idx]
                  if 0 <= col_idx < len(self._imp_tv_cols) else None)
        if not col_id:
            return
        if region == "heading" and col_id != "sel":
            shift = bool(event.state & 0x0001)
            self._imp_sort_by(col_id, shift)
        elif region == "cell":
            iid = self._imp_tv.identify_row(event.y)
            if iid:
                shift = bool(event.state & 0x0001)
                if shift and self._imp_anchor:
                    self._imp_range_check(iid)
                else:
                    self._imp_toggle_check(iid)
                    self._imp_anchor       = iid
                    self._imp_anchor_state = self._imp_checked.get(iid, True)

    def _imp_toggle_check(self, iid: str) -> None:
        if not iid or iid not in self._imp_checked:
            return
        self._imp_checked[iid] = not self._imp_checked[iid]
        sym = "☑" if self._imp_checked[iid] else "☐"
        try:
            self._imp_tv.set(iid, "sel", sym)
        except Exception:
            pass
        self._imp_update_count()

    def _imp_on_space(self, event) -> None:
        iid = self._imp_tv.focus()
        if iid:
            self._imp_toggle_check(iid)
            self._imp_anchor       = iid
            self._imp_anchor_state = self._imp_checked.get(iid, True)

    def _imp_range_check(self, end_iid: str) -> None:
        children = self._imp_tv.get_children()
        try:
            i1 = children.index(self._imp_anchor)
            i2 = children.index(end_iid)
        except ValueError:
            self._imp_toggle_check(end_iid)
            return
        start, stop = min(i1, i2), max(i1, i2) + 1
        for iid in children[start:stop]:
            if iid in self._imp_checked:
                self._imp_checked[iid] = self._imp_anchor_state
                sym = "☑" if self._imp_anchor_state else "☐"
                try: self._imp_tv.set(iid, "sel", sym)
                except Exception: pass
        self._imp_update_count()

    def _imp_autosize_columns(self) -> None:
        f_ui   = tkfont.Font(font=FONT_UI)
        f_bold = tkfont.Font(font=FONT_BOLD)
        pad    = 24
        for col_id in ("clase", "fecha"):
            heading_text = self._imp_tv.heading(col_id)["text"]
            max_w = f_bold.measure(heading_text) + pad
            for iid in self._imp_tv.get_children():
                w = f_ui.measure(str(self._imp_tv.set(iid, col_id))) + pad
                if w > max_w:
                    max_w = w
            self._imp_tv.column(col_id, width=max_w, minwidth=max_w)

    def _imp_sort_by(self, col_id: str, shift: bool) -> None:
        if shift:
            idx = next((i for i, (c, _) in enumerate(self._imp_sort_keys)
                        if c == col_id), -1)
            if idx >= 0:
                _, rev = self._imp_sort_keys[idx]
                self._imp_sort_keys[idx] = (col_id, not rev)
            else:
                self._imp_sort_keys.append((col_id, False))
        else:
            idx = next((i for i, (c, _) in enumerate(self._imp_sort_keys)
                        if c == col_id), -1)
            if idx == 0 and len(self._imp_sort_keys) == 1:
                _, rev = self._imp_sort_keys[0]
                self._imp_sort_keys[0] = (col_id, not rev)
            else:
                self._imp_sort_keys = [(col_id, False)]
        self._imp_refresh_treeview()
        self._imp_update_headers()

    def _imp_refresh_treeview(self) -> None:
        rows = list(self._imp_rows)
        for col_id, rev in reversed(self._imp_sort_keys):
            if col_id == "nombre":
                rows.sort(key=lambda r: r[0].lower(), reverse=rev)
            elif col_id == "clase":
                rows.sort(key=lambda r: r[2].lower(), reverse=rev)
            elif col_id == "fecha":
                rows.sort(key=lambda r: r[3], reverse=rev)
        for iid in self._imp_tv.get_children():
            self._imp_tv.delete(iid)
        for i, (fname, ext, tag, mtime, warn) in enumerate(rows):
            checked   = self._imp_checked.get(fname, True)
            sym       = "☑" if checked else "☐"
            name_str  = fname + (" ⚠" if warn else "")
            mtime_str = (datetime.datetime.fromtimestamp(mtime).strftime("%d/%m/%Y  %H:%M")
                         if mtime else "—")
            row_tag   = "row_a" if i % 2 == 0 else "row_b"
            self._imp_tv.insert("", "end", iid=fname,
                                values=(sym, name_str, tag, mtime_str),
                                tags=(row_tag,))

    def _imp_update_headers(self) -> None:
        col_labels = {"nombre": "Nombre", "clase": "Clase", "fecha": "Modificado"}
        for col_id, base in col_labels.items():
            idx = next((i for i, (c, _) in enumerate(self._imp_sort_keys)
                        if c == col_id), -1)
            if idx >= 0:
                _, rev = self._imp_sort_keys[idx]
                arrow = " ↓" if rev else " ↑"
                pri   = str(idx + 1) if len(self._imp_sort_keys) > 1 else ""
                label = base + arrow + pri
            else:
                label = base
            self._imp_tv.heading(col_id, text=label)

    def _import(self) -> None:
        target = self._imp_xlsm_var.get()
        folder = self._imp_folder_var.get()
        if not target or not os.path.exists(target):
            messagebox.showerror("Error", "Selecciona el .xlsm destino.", parent=self); return
        selected = [(fname, ext) for fname, ext, tag, mtime, warn in self._imp_rows
                    if self._imp_checked.get(fname, True)]
        if not selected:
            messagebox.showwarning("Sin seleccion", "Selecciona al menos un modulo.", parent=self); return
        do_backup = self._imp_backup_var.get()
        resp = messagebox.askyesno(
            "Confirmar importacion",
            f"Se importaran {len(selected)} modulo(s) en:\n  {os.path.basename(target)}\n\n"
            + ("Se creara copia de seguridad antes.\n\n" if do_backup else "")
            + "Esta operacion modifica el archivo. Continuar?",
            parent=self)
        if not resp: return
        self._btn_imp.config(state="disabled")
        self._btn_exp.config(state="disabled")
        self._log_write(f"\nImportando {len(selected)} modulo(s) en:\n  {target}\n", "info")
        threading.Thread(target=self._import_thread,
                         args=(target, folder, selected, do_backup), daemon=True).start()

    def _import_thread(self, target: str, src_folder: str,
                       files: list, do_backup: bool) -> None:
        pythoncom.CoInitialize()
        xl = wb = None
        try:
            xl, wb = _xl_open_rw(target)

            if wb.VBProject.Protection == 1:
                if not self._imp_password:
                    pwd_r: dict = {}; evt = threading.Event()
                    def _ask():
                        dlg = _PasswordDialog(self, os.path.basename(target), BG_IMPORTAR)
                        pwd_r["pw"] = dlg.password; evt.set()
                    self.after(0, _ask); evt.wait()
                    self._imp_password = pwd_r.get("pw")
                if not self._imp_password:
                    self.after(0, self._import_cancelled); return
                _, xl_pid = win32process.GetWindowThreadProcessId(xl.Hwnd)
                _unlock_vbproject(xl, xl_pid, self._imp_password)

            vbp = wb.VBProject

            # Copia de seguridad
            if do_backup:
                ts  = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
                bk  = os.path.join(os.path.dirname(target),
                                   os.path.splitext(os.path.basename(target))[0]
                                   + f"_BACKUP_{ts}")
                os.makedirs(bk, exist_ok=True)
                bk_n = 0
                for comp in vbp.VBComponents:
                    ext = _COMP_EXT.get(comp.Type, "")
                    if ext:
                        try: comp.Export(os.path.join(bk, comp.Name + ext)); bk_n += 1
                        except Exception: pass
                self.after(0, self._log_write,
                           f"  Backup: {bk_n} modulos -> {os.path.basename(bk)}\n", "info")

            # Importar modulos
            count = 0
            for fname, ext in files:
                src_path = os.path.join(src_folder, fname)
                mod_name = os.path.splitext(fname)[0]
                if not os.path.exists(src_path):
                    self.after(0, self._log_write, f"  ! {fname}: no encontrado\n", "warn")
                    continue

                existing = None
                try: existing = vbp.VBComponents(mod_name)
                except Exception: pass

                if existing is not None and existing.Type == 100:
                    # Modulo Documento: vaciar + pegar codigo
                    try:
                        code = _extract_code(src_path)
                        cm = existing.CodeModule
                        n  = cm.CountOfLines
                        if n > 0: cm.DeleteLines(1, n)
                        if code.strip(): cm.InsertLines(1, code)
                        self.after(0, self._log_write,
                                   f"  -> {fname}  [Documento]\n", "ok")
                        count += 1
                    except Exception as e:
                        self.after(0, self._log_write, f"  ! {fname}: {e}\n", "error")
                else:
                    # Modulo estandar: eliminar si existe + importar
                    if existing:
                        try: vbp.VBComponents.Remove(existing)
                        except Exception: pass
                    try:
                        vbp.VBComponents.Import(src_path)
                        self.after(0, self._log_write, f"  -> {fname}\n", "ok")
                        count += 1
                    except Exception as e:
                        self.after(0, self._log_write, f"  ! {fname}: {e}\n", "error")

            wb.Save()
            self.after(0, self._import_done, count, target)

        except Exception as exc:
            self.after(0, self._on_import_error, str(exc))
        finally:
            if wb:
                try: wb.Close(True)
                except: pass
            if xl:
                try: xl.Quit()
                except: pass
            pythoncom.CoUninitialize()

    def _import_cancelled(self) -> None:
        self._btn_imp.config(state="normal")
        self._btn_exp.config(state="normal")
        self._log_write("\n  Importacion cancelada.\n", "info")

    def _import_done(self, count: int, target: str) -> None:
        self._btn_imp.config(state="normal")
        self._btn_exp.config(state="normal")
        self._log_write(
            f"\n  Importacion completada: {count} modulo(s) en {os.path.basename(target)}\n", "ok")
        messagebox.showinfo("Importacion completada",
                            f"Se importaron {count} modulo(s) en:\n\n{target}", parent=self)

    def _on_import_error(self, msg: str) -> None:
        self._btn_imp.config(state="normal")
        self._btn_exp.config(state="normal")
        self._log_write(f"\n  ERROR importacion: {msg}\n", "error")
        messagebox.showerror("Error de importacion", msg, parent=self)


if __name__ == "__main__":
    App().mainloop()
