"""
Export_VBA_Moduls.py
Exporta modulos VBA de archivos .xlsm  (.bas / .cls / .frm)
Requiere: Python 3.9+, pywin32, tkinter (stdlib)
"""

import os
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk

import pythoncom
import win32api
import win32com.client
import win32con
import win32gui
import win32process
import winreg

_REG_KEY   = r"Software\ExportVBAModuls"
_REG_VALUE = "LastXlsmFile"

def _reg_read() -> str:
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, _REG_KEY) as k:
            v, _ = winreg.QueryValueEx(k, _REG_VALUE)
            return v
    except OSError:
        return ""

def _reg_write(path: str) -> None:
    try:
        with winreg.CreateKey(winreg.HKEY_CURRENT_USER, _REG_KEY) as k:
            winreg.SetValueEx(k, _REG_VALUE, 0, winreg.REG_SZ, path)
    except OSError:
        pass

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

_COMP_EXT = {1: ".bas", 2: ".cls", 3: ".frm", 100: ".cls"}
_COMP_TAG = {1: "Modulo", 2: "Clase", 3: "Formulario", 100: "Documento"}

def _xl_open(path: str):
    xl = win32com.client.DispatchEx("Excel.Application")
    xl.Visible = False; xl.DisplayAlerts = False
    xl.ScreenUpdating = False; xl.EnableEvents = False
    wb = xl.Workbooks.Open(os.path.abspath(path), ReadOnly=True, UpdateLinks=False)
    return xl, wb

def _xl_close(xl, wb) -> None:
    try:
        if wb: wb.Close(False)
    except Exception: pass
    try:
        if xl: xl.Quit()
    except Exception: pass

_FIRMA_B64 = "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"

# Sistema de fuentes escalable — equivalente a rem/em en CSS.
# Cambiar BASE escala todo el programa proporcionalmente.
BASE = 14
FONT_UI    = ("Verdana", BASE)
FONT_BOLD  = ("Verdana", BASE, "bold")
FONT_TITLE = ("Verdana", BASE + 1, "bold")
FONT_ENTRY = ("Verdana", BASE - 1)       # recuadros de ruta (1pt menos que UI)
FONT_LOG   = ("Courier New", BASE)

BG_APP     = "#F2F3F4"
BG_ARCHIVO = "#2874A6"
BG_MODULOS = "#7D3C98"
BG_DESTINO = "#1E8449"
BG_LOG_HDR = "#5D6D7E"

C_SELEC    = "#D5D8DC"
C_TODOS    = "#AED6F1"
C_NINGUNO  = "#F5B7B1"
C_EXPORTAR = "#A9DFBF"
C_ACCION   = "#F0B27A"


# ── Tooltip ─────────────────────────────────────────────────────────────────
# ── Automatización win32 para el diálogo de contraseña VBA de Excel ──────────
def _fill_excel_vba_dialog(xl_pid: int, password: str, timeout: float = 10.0) -> bool:
    """
    Espera a que Excel muestre su diálogo de contraseña VBA y lo rellena
    automáticamente con la contraseña proporcionada.
    Funciona con Excel en cualquier idioma (busca Edit control en la ventana).
    """
    import time
    deadline = time.monotonic() + timeout

    while time.monotonic() < deadline:
        found_hwnd = None

        def _enum(hwnd, _):
            nonlocal found_hwnd
            if found_hwnd or not win32gui.IsWindowVisible(hwnd):
                return True
            try:
                _, pid = win32process.GetWindowThreadProcessId(hwnd)
            except Exception:
                return True
            if pid != xl_pid:
                return True
            title = win32gui.GetWindowText(hwnd).lower()
            cls   = win32gui.GetClassName(hwnd)
            # Detectar por título (multiidioma) o por clase de diálogo estándar con Edit
            if any(kw in title for kw in ("contraseña", "password", "mot de passe",
                                           "kennwort", "wachtwoord")):
                found_hwnd = hwnd
            elif cls == "#32770":          # diálogo Windows estándar
                if win32gui.FindWindowEx(hwnd, None, "Edit", None):
                    found_hwnd = hwnd      # tiene campo de texto → probable contraseña
            return True

        win32gui.EnumWindows(_enum, None)

        if found_hwnd:
            edit = win32gui.FindWindowEx(found_hwnd, None, "Edit", None)
            if edit:
                win32gui.SendMessage(edit, win32con.WM_SETTEXT, 0, password)
                time.sleep(0.05)
                # Pulsar Enter para confirmar
                win32api.keybd_event(win32con.VK_RETURN, 0, 0, 0)
                win32api.keybd_event(win32con.VK_RETURN, 0, win32con.KEYEVENTF_KEYUP, 0)
                return True

        time.sleep(0.12)

    return False


# ── Diálogo de contraseña (tkinter modal) ─────────────────────────────────────
class _PasswordDialog(tk.Toplevel):
    """Solicita la contraseña de un proyecto VBA protegido."""

    def __init__(self, parent: tk.Tk, filename: str) -> None:
        super().__init__(parent)
        self.title("Proyecto VBA protegido")
        self.configure(bg=BG_APP)
        self.resizable(False, False)
        self.grab_set()
        self.password: str | None = None

        self.update_idletasks()
        pw = parent.winfo_rootx() + parent.winfo_width()  // 2
        ph = parent.winfo_rooty() + parent.winfo_height() // 2
        self.geometry(f"+{pw - 220}+{ph - 110}")

        # Cabecera azul (mismo estilo que bloque ARCHIVO)
        tk.Frame(self, bg=BG_ARCHIVO, height=4).pack(fill="x")
        tk.Label(self, text=f"  🔒  {filename}",
                 bg=BG_ARCHIVO, fg="white", font=FONT_BOLD,
                 padx=14, pady=8, anchor="w").pack(fill="x")

        # Cuerpo
        body = tk.Frame(self, bg=BG_APP, padx=20, pady=14)
        body.pack(fill="both", expand=True)
        body.columnconfigure(1, weight=1)

        tk.Label(body, text="El proyecto VBA está protegido.",
                 bg=BG_APP, font=FONT_UI, anchor="w").grid(
            row=0, column=0, columnspan=2, sticky="w", pady=(0, 10))

        tk.Label(body, text="Contraseña:", bg=BG_APP, font=FONT_UI).grid(
            row=1, column=0, sticky="w", padx=(0, 10), pady=4)

        self._entry = ttk.Entry(body, show="●", font=FONT_UI, width=26)
        self._entry.grid(row=1, column=1, sticky="ew", pady=4)
        self._entry.focus_set()
        self._entry.bind("<Return>", lambda _: self._ok())
        self._entry.bind("<Escape>", lambda _: self._cancel())

        # Botón ojo — mostrar / ocultar contraseña
        self._pw_visible = False
        def _toggle_pw():
            self._pw_visible = not self._pw_visible
            self._entry.configure(show="" if self._pw_visible else "●")
            _btn_eye.configure(text="🙈" if self._pw_visible else "👁")
        _btn_eye = tk.Button(body, text="👁", command=_toggle_pw,
                             bg=BG_APP, activebackground=BG_APP,
                             relief="flat", bd=0, cursor="hand2",
                             font=("Segoe UI Emoji", BASE + 2))
        _btn_eye.grid(row=1, column=2, padx=(4, 0), pady=4)

        ttk.Separator(self).pack(fill="x", padx=14, pady=(4, 0))
        frm = tk.Frame(self, bg=BG_APP, pady=8)
        frm.pack(fill="x", padx=14)
        ttk.Button(frm, text="Cancelar", command=self._cancel).pack(side="right", padx=(6, 0))
        ttk.Button(frm, text="Aceptar", style="Export.TButton",
                   command=self._ok).pack(side="right")

        self.protocol("WM_DELETE_WINDOW", self._cancel)
        self.wait_window()

    def _ok(self) -> None:
        self.password = self._entry.get() or None
        self.destroy()

    def _cancel(self) -> None:
        self.destroy()


class _Tooltip:
    """Etiqueta flotante que aparece al situar el cursor sobre un widget."""
    def __init__(self, widget: tk.Widget, text: str) -> None:
        self._w    = widget
        self._text = text
        self._win: tk.Toplevel | None = None
        widget.bind("<Enter>",       self._show)
        widget.bind("<Leave>",       self._hide)
        widget.bind("<ButtonPress>", self._hide)

    def _show(self, _=None) -> None:
        if self._win:
            return
        x = self._w.winfo_rootx() + self._w.winfo_width() // 2
        y = self._w.winfo_rooty() + self._w.winfo_height() + 4
        _TRANSP = "#f0f0f1"                    # color clave para transparencia
        self._win = tk.Toplevel(self._w)
        self._win.wm_overrideredirect(True)
        self._win.wm_geometry(f"+{x}+{y}")
        self._win.configure(bg=_TRANSP)
        try:
            self._win.wm_attributes("-transparentcolor", _TRANSP)  # fondo invisible
        except Exception:
            pass                               # fallback: fondo neutro
        tk.Label(self._win, text=self._text,
                 bg=_TRANSP, fg="#555555",
                 relief="flat", bd=0,
                 font=("Verdana", BASE - 2, "italic"), padx=3, pady=1).pack()

    def _hide(self, _=None) -> None:
        if self._win:
            self._win.destroy()
            self._win = None


class _ScrollFrame(ttk.Frame):
    def __init__(self, parent, **kw):
        super().__init__(parent, **kw)
        self._canvas = tk.Canvas(self, borderwidth=0, highlightthickness=0, bg=BG_APP)
        vsb = ttk.Scrollbar(self, orient="vertical", command=self._canvas.yview)
        self._canvas.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")
        self._canvas.pack(side="left", fill="both", expand=True)
        self.inner = tk.Frame(self._canvas, bg=BG_APP)
        self._win_id = self._canvas.create_window((0, 0), window=self.inner, anchor="nw")
        self.inner.bind("<Configure>",
            lambda e: self._canvas.configure(scrollregion=self._canvas.bbox("all")))
        self._canvas.bind("<Configure>",
            lambda e: self._canvas.itemconfig(self._win_id, width=e.width))
        self._canvas.bind("<MouseWheel>", self._on_wheel)
        self.inner.bind("<MouseWheel>", self._on_wheel)

    def _on_wheel(self, event):
        self._canvas.yview_scroll(-1 if event.delta > 0 else 1, "units")

    def bind_child_wheel(self, widget):
        widget.bind("<MouseWheel>", self._on_wheel)


class App(tk.Tk):

    W, H = 1400, 800

    def __init__(self):
        super().__init__()
        self.title("Export VBA Moduls")
        self.resizable(True, True)
        self.minsize(700, 560)
        self.configure(bg=BG_APP)

        self._xlsm_var = tk.StringVar()
        self._dest_var  = tk.StringVar()
        self._modules: list[tuple[str, int, str, tk.BooleanVar]] = []
        self._vba_password: str | None = None   # contraseña del proyecto VBA protegido

        self._firma_img = tk.PhotoImage(data=_FIRMA_B64)

        self._build_styles()
        self._build_ui()
        self._center(self.W, self.H)

        if not _vba_access_enabled():
            self._show_warn_banner()

        last = _reg_read()
        if last and os.path.exists(last):
            self._xlsm_var.set(last)
            self._lbl_xlsm_title.config(text=os.path.basename(last))
            self._propose_dest(last)
            self._load_modules_bg(last)

    def _center(self, w: int, h: int) -> None:
        self.update_idletasks()
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        self.geometry(f"{w}x{h}+{(sw - w) // 2}+{max(0, (sh - h) // 2)}")

    def _build_styles(self) -> None:
        s = ttk.Style(self)
        s.theme_use("clam")
        s.configure(".",             font=FONT_UI,  background=BG_APP)
        s.configure("TFrame",        background=BG_APP)
        s.configure("TCheckbutton",  background=BG_APP, font=FONT_UI)
        s.map("TCheckbutton",        background=[("active", BG_APP)])
        s.configure("TEntry",        font=FONT_UI, fieldbackground="white")
        s.configure("TScrollbar",    background=BG_APP)
        s.configure("Sel.TButton",     font=FONT_UI,   background=C_SELEC)
        s.map("Sel.TButton",           background=[("active", "#BFC9CA")])
        s.configure("Todos.TButton",   font=FONT_UI,   background=C_TODOS)
        s.map("Todos.TButton",         background=[("active", "#85C1E9")])
        s.configure("Ninguno.TButton", font=FONT_UI,   background=C_NINGUNO)
        s.map("Ninguno.TButton",       background=[("active", "#F1948A")])
        s.configure("Export.TButton",  font=FONT_BOLD, background=C_EXPORTAR)
        s.map("Export.TButton",        background=[("active", "#7DCEA0")])
        s.configure("Accion.TButton",  font=FONT_UI,   background=C_ACCION)
        s.map("Accion.TButton",        background=[("active", "#E59866")])

    def _bloque(self, titulo: str, bg_titulo: str,
                expand: bool = False) -> tuple[tk.Frame, tk.Frame]:
        outer = tk.Frame(self, bg=BG_APP, bd=1, relief="solid",
                         highlightbackground="#AAAAAA", highlightthickness=1)
        outer.pack(fill="both" if expand else "x",
                   expand=expand, padx=10, pady=(6, 0))
        title_bar = tk.Frame(outer, bg=bg_titulo)
        title_bar.pack(fill="x")
        tk.Label(title_bar, text=f"  {titulo}",
                 bg=bg_titulo, fg="white", font=FONT_TITLE, anchor="w", pady=2
                 ).grid(row=0, column=0, sticky="w")
        content = tk.Frame(outer, bg=BG_APP, padx=12, pady=8)
        content.pack(fill="both", expand=True)
        content.columnconfigure(1, weight=1)
        return title_bar, content

    def _build_ui(self) -> None:

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

        # ARCHIVO
        tb_arc, c_arc = self._bloque("ARCHIVO .xlsm", BG_ARCHIVO)
        tb_arc.columnconfigure(1, weight=1)
        self._lbl_xlsm_title = tk.Label(
            tb_arc, text="", bg=BG_ARCHIVO, fg="white", font=FONT_UI, anchor="w")  # blanco: 4.98:1 WCAG ✓
        self._lbl_xlsm_title.grid(row=0, column=1, sticky="ew", padx=(6, 8))
        _firma_lbl = tk.Label(tb_arc, image=self._firma_img, bg=BG_ARCHIVO, bd=0,
                              cursor="hand2")
        _firma_lbl.grid(row=0, column=2, padx=(0, 8), pady=2)
        _Tooltip(_firma_lbl, "Dugarry")

        tk.Label(c_arc, text="Archivo.Xlsm:", bg=BG_APP, font=FONT_UI, anchor="w").grid(
            row=0, column=0, sticky="w", padx=(0, 10), pady=4)
        tk.Entry(c_arc, textvariable=self._xlsm_var, state="readonly",
                 font=FONT_ENTRY, bg="white", readonlybackground="#ECECEC",
                 fg="#222222", relief="sunken", bd=1).grid(
            row=0, column=1, sticky="ew", pady=4, ipady=4)
        ttk.Button(c_arc, text="Seleccionar", style="Sel.TButton",
                   command=self._pick_file).grid(
            row=0, column=2, padx=(8, 0), pady=4, ipadx=4)

        # MODULOS
        tb_mod, c_mod = self._bloque("MODULOS VBA", BG_MODULOS, expand=True)
        tb_mod.columnconfigure(0, weight=1)
        self._lbl_count = tk.Label(tb_mod, text="", bg=BG_MODULOS, fg="white", font=FONT_UI)  # blanco: 7.14:1 WCAG ✓
        self._lbl_count.grid(row=0, column=1, padx=(0, 6))
        ttk.Button(tb_mod, text="Todos",   style="Todos.TButton",  command=self._sel_all
                   ).grid(row=0, column=2, padx=(0, 4), pady=4, ipadx=4)
        ttk.Button(tb_mod, text="Ninguno", style="Ninguno.TButton", command=self._sel_none
                   ).grid(row=0, column=3, padx=(0, 8), pady=4, ipadx=4)
        self._scroll = _ScrollFrame(c_mod)
        self._scroll.pack(fill="both", expand=True)

        # DESTINO
        tb_dest, c_dest = self._bloque("CARPETA DE DESTINO", BG_DESTINO)
        tb_dest.columnconfigure(1, weight=1)
        self._lbl_dest_title = tk.Label(
            tb_dest, text="", bg=BG_DESTINO, fg="white", font=FONT_UI, anchor="w")  # blanco: 4.69:1 WCAG ✓
        self._lbl_dest_title.grid(row=0, column=1, sticky="ew", padx=(6, 8))

        tk.Label(c_dest, text="Carpeta:", bg=BG_APP, font=FONT_UI, anchor="w").grid(
            row=0, column=0, sticky="w", padx=(0, 10), pady=4)
        tk.Entry(c_dest, textvariable=self._dest_var,
                 font=FONT_ENTRY, bg="white", fg="#222222",
                 relief="sunken", bd=1).grid(
            row=0, column=1, sticky="ew", pady=4, ipady=4)
        ttk.Button(c_dest, text="Examinar", style="Sel.TButton",
                   command=self._pick_dest).grid(
            row=0, column=2, padx=(8, 0), pady=4, ipadx=4)

        # REGISTRO
        tb_log, c_log = self._bloque("REGISTRO", BG_LOG_HDR)
        tb_log.columnconfigure(0, weight=1)
        c_log.rowconfigure(0, weight=1)
        ttk.Button(tb_log, text="Limpiar", style="Accion.TButton",
                   command=self._log_clear).grid(
            row=0, column=1, padx=(0, 8), pady=4, ipadx=6)
        self._btn_export = ttk.Button(
            tb_log, text="Exportar Módulos",
            style="Export.TButton", command=self._export)
        self._btn_export.grid(row=0, column=2, padx=(0, 8), pady=4, ipadx=6)
        self._log = scrolledtext.ScrolledText(
            c_log, height=20, state="disabled", font=FONT_LOG, wrap="word",
            spacing1=5, spacing3=5,          # interlineado 1.5x (WCAG 1.4.8)
            bg="#1E1E1E", fg="#D4D4D4", insertbackground="white", relief="flat", bd=0)
        self._log.grid(row=0, column=0, columnspan=3, sticky="nsew")
        self._log.tag_configure("ok",    foreground="#4EC94E")
        self._log.tag_configure("error", foreground="#FF6B6B")
        self._log.tag_configure("info",  foreground="#85C1E9")

        tk.Frame(self, bg=BG_APP, height=8).pack()

    def _show_warn_banner(self) -> None:
        self._frm_warn.pack(fill="x", padx=10, pady=(8, 0),
                            before=self.winfo_children()[0])

    def _show_vba_instructions(self) -> None:
        win = tk.Toplevel(self)
        win.title("Activar acceso al modelo VBA")
        win.configure(bg=BG_APP)
        win.resizable(False, False)
        win.grab_set()
        win.update_idletasks()
        pw = self.winfo_rootx() + self.winfo_width()  // 2
        ph = self.winfo_rooty() + self.winfo_height() // 2
        win.geometry(f"+{pw - 260}+{ph - 230}")
        tk.Frame(win, bg="#856404", height=4).pack(fill="x")
        hdr = tk.Frame(win, bg="#FFF3CD")
        hdr.pack(fill="x")
        tk.Label(hdr, text="  Acceso al modelo de objetos VBA desactivado",
                 bg="#FFF3CD", fg="#7d5800", font=FONT_BOLD,
                 padx=14, pady=10).pack(side="left")
        body = tk.Frame(win, bg=BG_APP, padx=16, pady=12)
        body.pack(fill="both", expand=True)
        txt = (
            "Para que esta herramienta pueda leer y exportar los modulos VBA,\n"
            "debes activar el acceso en el Centro de confianza de Excel:\n\n"
            "  1.  Abre Microsoft Excel\n\n"
            "  2.  Ve a:  Archivo -> Opciones -> Centro de confianza\n\n"
            "  3.  Haz clic en:\n"
            "       Configuracion del Centro de confianza...\n\n"
            "  4.  Panel izquierdo:  Configuracion de macros\n\n"
            "  5.  Activa la casilla:\n"
            "       Confiar en el acceso al modelo de objetos\n"
            "       de proyectos de VBA\n\n"
            "  6.  Acepta con Aceptar en todas las ventanas\n\n"
            "  7.  Cierra y vuelve a abrir  Export VBA Moduls"
        )
        tk.Label(body, text=txt, justify="left", bg=BG_APP, font=FONT_UI).pack(anchor="w")
        ttk.Separator(win).pack(fill="x", padx=14)
        ttk.Button(win, text="  Entendido  ", command=win.destroy).pack(
            side="right", padx=14, pady=10)

    def _pick_file(self) -> None:
        init = (os.path.dirname(self._xlsm_var.get())
                if self._xlsm_var.get() else os.path.expanduser("~"))
        path = filedialog.askopenfilename(
            parent=self, title="Seleccionar archivo .xlsm", initialdir=init,
            filetypes=[("Excel con macros", "*.xlsm"), ("Todos los archivos", "*.*")])
        if not path: return
        self._xlsm_var.set(path)
        self._lbl_xlsm_title.config(text=os.path.basename(path))
        self._vba_password = None        # al cambiar de archivo, limpiar contraseña anterior
        _reg_write(path)
        self._propose_dest(path)
        self._load_modules_bg(path)

    def _propose_dest(self, xlsm_path: str) -> None:
        folder = os.path.dirname(xlsm_path)
        stem   = os.path.splitext(os.path.basename(xlsm_path))[0]
        dest   = os.path.join(folder, stem + "_VBA-Moduls")
        self._dest_var.set(dest)
        self._lbl_dest_title.config(text=os.path.basename(dest))

    def _pick_dest(self) -> None:
        init = (os.path.dirname(self._xlsm_var.get())
                if self._xlsm_var.get() else os.path.expanduser("~"))
        folder = filedialog.askdirectory(parent=self, title="Carpeta de destino", initialdir=init)
        if folder:
            self._dest_var.set(folder)
            self._lbl_dest_title.config(text=os.path.basename(folder) or folder)

    def _sel_all(self) -> None:
        for *_, v in self._modules: v.set(True)
        self._update_count()

    def _sel_none(self) -> None:
        for *_, v in self._modules: v.set(False)
        self._update_count()

    def _update_count(self) -> None:
        sel = sum(1 for *_, v in self._modules if v.get())
        tot = len(self._modules)
        self._lbl_count.config(text=f"{sel} / {tot}" if tot else "")

    def _log_write(self, text: str, tag: str = "") -> None:
        self._log.configure(state="normal")
        self._log.insert("end", text, tag)
        self._log.see("end")
        self._log.configure(state="disabled")

    def _log_clear(self) -> None:
        self._log.configure(state="normal")
        self._log.delete("1.0", "end")
        self._log.configure(state="disabled")

    def _load_modules_bg(self, path: str) -> None:
        self._log_write(f"\nLeyendo modulos de:  {os.path.basename(path)}\n", "info")
        self._btn_export.config(state="disabled")
        self._clear_checks()
        threading.Thread(target=self._load_thread, args=(path,), daemon=True).start()

    def _load_thread(self, path: str) -> None:
        pythoncom.CoInitialize()
        xl = wb = None
        try:
            xl, wb = _xl_open(path)

            # ── Proyecto VBA protegido con contraseña ─────────────────────────
            if wb.VBProject.Protection == 1:
                # 1. Pedir contraseña al usuario desde el hilo UI
                pwd_result: dict = {}
                pwd_event  = threading.Event()
                fname = os.path.basename(path)

                def _ask():
                    dlg = _PasswordDialog(self, fname)
                    pwd_result["pw"] = dlg.password
                    pwd_event.set()

                self.after(0, _ask)
                pwd_event.wait()
                password = pwd_result.get("pw")

                if not password:
                    self.after(0, self._on_load_cancelled)
                    return

                self._vba_password = password   # guardar para reutilizar en exportación

                # 2. Obtener PID mientras Excel aún está oculto (Hwnd funciona en oculto)
                _, xl_pid = win32process.GetWindowThreadProcessId(xl.Hwnd)

                # 3. Hacer Excel visible y minimizado al taskbar
                #    (xl.Left/Top no funcionan con Visible=False → error corregido)
                xl.Visible = True
                try:
                    xl.WindowState = -4140          # xlMinimized
                except Exception:
                    pass
                # Abrir VBE: en algunas versiones de Excel es necesario
                # para que el diálogo de contraseña aparezca vía COM
                try:
                    xl.VBE.MainWindow.Visible = True
                except Exception:
                    pass

                # 4. Hilo que detecta y rellena el diálogo nativo de Excel
                filled = threading.Event()
                def _filler():
                    if _fill_excel_vba_dialog(xl_pid, password):
                        filled.set()
                threading.Thread(target=_filler, daemon=True).start()

                # 5. Acceder a VBComponents dispara el diálogo de contraseña.
                #    En Excel 365 puede lanzar error sin mostrar diálogo → capturar
                try:
                    _ = wb.VBProject.VBComponents
                    filled.wait(timeout=12.0)
                except Exception as prot_exc:
                    prot_str = str(prot_exc)
                    xl.Visible = False
                    if any(c in prot_str for c in ("-2146827284", "-2147352567")):
                        raise RuntimeError(
                            "Excel no mostró el diálogo de contraseña automáticamente.\n\n"
                            "Solución alternativa:\n"
                            "  1. Abre el archivo en Excel\n"
                            "  2. En el Editor VBA (Alt+F11) introduce la contraseña\n"
                            "     del proyecto (clic derecho → Propiedades)\n"
                            "  3. Desprotege el proyecto y guarda\n"
                            "  4. Vuelve a usar esta herramienta"
                        ) from None
                    raise
                xl.Visible = False

            # ── Leer y ordenar módulos ────────────────────────────────────────
            _TYPE_ORDER = {1: 0, 2: 1, 3: 2, 100: 3}
            mods = sorted(
                [(c.Name, c.Type, _COMP_EXT[c.Type])
                 for c in wb.VBProject.VBComponents if c.Type in _COMP_EXT],
                key=lambda m: (_TYPE_ORDER[m[1]], m[0].lower())
            )
            self.after(0, self._populate, mods)
        except Exception as exc:
            self.after(0, self._on_load_error, str(exc))
        finally:
            _xl_close(xl, wb); pythoncom.CoUninitialize()

    def _populate(self, mods: list) -> None:
        self._clear_checks()
        for i, (name, typ, ext) in enumerate(mods):
            var = tk.BooleanVar(value=True)
            cb  = tk.Checkbutton(
                self._scroll.inner,
                text=f"  {name}{ext}   [{_COMP_TAG.get(typ, '?')}]",
                variable=var, command=self._update_count,
                bg=BG_APP, font=FONT_UI, anchor="w", activebackground=BG_APP,
                pady=5)            # interlineado 1.5x sobre BASE=14 (WCAG 1.4.8)
            cb.grid(row=i, column=0, sticky="w", padx=8)
            self._scroll.bind_child_wheel(cb)
            self._modules.append((name, typ, ext, var))
        self._update_count()
        self._log_write(f"  -> {len(mods)} modulo(s) encontrado(s)\n", "ok")
        self._btn_export.config(state="normal")

    def _on_load_cancelled(self) -> None:
        """El usuario canceló el diálogo de contraseña."""
        self._btn_export.config(state="normal")
        self._log_write("\n  Operación cancelada — proyecto VBA protegido.\n", "info")

    def _on_load_error(self, msg: str) -> None:
        self._btn_export.config(state="normal")
        if not _vba_access_enabled():
            self._show_warn_banner()
            self._log_write("\n  ERROR: Acceso al modelo VBA desactivado.\n"
                            "  Pulsa 'Como activarlo' en el banner superior.\n", "error")
            self._show_vba_instructions()
        else:
            self._log_write(f"\n  ERROR: {msg}\n", "error")

    def _clear_checks(self) -> None:
        for w in self._scroll.inner.winfo_children(): w.destroy()
        self._modules = []; self._lbl_count.config(text="")

    def _export(self) -> None:
        path = self._xlsm_var.get(); dest = self._dest_var.get().strip()
        if not path or not os.path.exists(path):
            messagebox.showerror("Error", "Selecciona un archivo .xlsm valido.", parent=self); return
        selected = [(n, t, e) for n, t, e, v in self._modules if v.get()]
        if not selected:
            messagebox.showwarning("Sin seleccion", "Selecciona al menos un modulo.", parent=self); return
        if not dest:
            messagebox.showerror("Error", "Especifica una carpeta de destino.", parent=self); return
        if os.path.exists(dest):
            archivos = [f for f in os.listdir(dest) if os.path.isfile(os.path.join(dest, f))]
            if archivos:
                resp = messagebox.askyesnocancel(
                    "Carpeta no vacia",
                    f"La carpeta ya existe y contiene {len(archivos)} archivo(s):\n\n{dest}\n\n"
                    "Si = Borrar y exportar\nNo = Exportar (sobreescribir)\nCancelar = Salir",
                    parent=self)
                if resp is None: return
                if resp:
                    for f in archivos:
                        try: os.remove(os.path.join(dest, f))
                        except OSError: pass
        else:
            try: os.makedirs(dest, exist_ok=True)
            except OSError as exc:
                messagebox.showerror("Error", f"No se pudo crear la carpeta:\n{exc}", parent=self); return
        self._btn_export.config(state="disabled")
        self._log_write(f"\nExportando {len(selected)} modulo(s) a:\n  {dest}\n", "info")
        sel_names = {n for n, _, _ in selected}
        threading.Thread(target=self._export_thread, args=(path, dest, sel_names), daemon=True).start()

    def _export_thread(self, path: str, dest: str, names: set) -> None:
        pythoncom.CoInitialize()
        xl = wb = None
        try:
            xl, wb = _xl_open(path)

            # ── Proyecto protegido: reutilizar contraseña del paso de carga ──
            if wb.VBProject.Protection == 1:
                password = self._vba_password
                if not password:
                    self.after(0, self._on_export_error,
                               "El proyecto VBA está protegido. Vuelve a cargar el archivo para introducir la contraseña.")
                    return
                _, xl_pid = win32process.GetWindowThreadProcessId(xl.Hwnd)
                xl.Visible = True
                try: xl.WindowState = -4140
                except Exception: pass
                try: xl.VBE.MainWindow.Visible = True
                except Exception: pass
                filled = threading.Event()
                def _filler2():
                    if _fill_excel_vba_dialog(xl_pid, password): filled.set()
                threading.Thread(target=_filler2, daemon=True).start()
                try:
                    _ = wb.VBProject.VBComponents
                    filled.wait(timeout=12.0)
                except Exception:
                    xl.Visible = False
                    raise
                xl.Visible = False

            count = 0
            for comp in wb.VBProject.VBComponents:
                if comp.Name in names and comp.Type in _COMP_EXT:
                    out = os.path.join(dest, comp.Name + _COMP_EXT[comp.Type])
                    comp.Export(out)
                    self.after(0, self._log_write,
                               f"  ✔  {comp.Name}{_COMP_EXT[comp.Type]}\n", "ok")
                    count += 1
            self.after(0, self._export_done, count, dest)
        except Exception as exc:
            self.after(0, self._on_export_error, str(exc))
        finally:
            _xl_close(xl, wb); pythoncom.CoUninitialize()

    def _export_done(self, count: int, dest: str) -> None:
        self._btn_export.config(state="normal")
        self._log_write(f"\n  Exportacion completada: {count} modulo(s) guardado(s).\n", "ok")
        messagebox.showinfo("Exportacion completada",
                            f"Se exportaron  {count}  modulo(s) en:\n\n{dest}", parent=self)

    def _on_export_error(self, msg: str) -> None:
        self._btn_export.config(state="normal")
        self._log_write(f"\n  ERROR: {msg}\n", "error")
        messagebox.showerror("Error de exportacion", msg, parent=self)


if __name__ == "__main__":
    App().mainloop()
