"""
Rename_Files.py  v2
Renombra (o copia y renombra) ficheros basándose en una tabla Excel de dos columnas.
Requiere: Python 3.9+, openpyxl, tkinter (stdlib)
"""

import os
import shutil
import threading
import tkinter as tk
import tkinter.font as tkfont
from tkinter import filedialog, messagebox, scrolledtext, ttk
import winreg

try:
    import openpyxl
except ImportError:
    openpyxl = None

_REG_KEY     = r"Software\RenameFiles"
_REG_XLSX    = "LastExcelFile"
_REG_CARPETA = "LastCarpetaExcels"


def _reg_read(key: str) -> str:
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, _REG_KEY) as k:
            v, _ = winreg.QueryValueEx(k, key)
            return v
    except OSError:
        return ""


def _reg_write(key: str, value: str) -> None:
    try:
        with winreg.CreateKey(winreg.HKEY_CURRENT_USER, _REG_KEY) as k:
            winreg.SetValueEx(k, key, 0, winreg.REG_SZ, value)
    except OSError:
        pass


_FIRMA_B64 = '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'

BASE       = 14
FONT_UI    = ("Verdana", BASE)
FONT_BOLD  = ("Verdana", BASE, "bold")
FONT_TITLE = ("Verdana", BASE + 1, "bold")
FONT_ENTRY = ("Verdana", BASE - 1)
FONT_LOG   = ("Courier New", BASE)

BG_APP     = "#F2F3F4"
BG_ARCHIVO = "#2874A6"
BG_CARPETA = "#1A5276"
BG_TABLA   = "#7D3C98"
BG_DESTINO = "#1E8449"
BG_LOG_HDR = "#5D6D7E"

C_SELEC    = "#D5D8DC"
C_TODOS    = "#AED6F1"
C_NINGUNO  = "#F5B7B1"
C_EXPORTAR = "#A9DFBF"
C_ACCION   = "#F0B27A"

_HINT_DEST    = "(opcional — si vacío, renombra en el origen)"
_HINT_CARPETA = "(carpeta donde están los ficheros a renombrar)"

_COL_IDS     = ("mail", "fichero", "newname", "resultado")
_COL_FIXED   = ("New Name", "Resultado")


def _mail_clean(mail: str) -> str:
    """Quita los últimos 5 caracteres ("ua.es") del mail."""
    return mail.removesuffix("ua.es")


class _Tooltip:
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
        _TRANSP = "#f0f0f1"
        self._win = tk.Toplevel(self._w)
        self._win.wm_overrideredirect(True)
        self._win.wm_geometry(f"+{x}+{y}")
        self._win.configure(bg=_TRANSP)
        try:
            self._win.wm_attributes("-transparentcolor", _TRANSP)
        except Exception:
            pass
        tk.Label(self._win, text=self._text,
                 bg=_TRANSP, fg="#555555",
                 relief="flat", bd=0,
                 font=("Verdana", BASE - 2, "italic"), padx=3, pady=1).pack()

    def _hide(self, _=None) -> None:
        if self._win:
            self._win.destroy()
            self._win = None


class App(tk.Tk):

    H = 860

    def __init__(self):
        super().__init__()
        self.title("Rename Files")
        self.resizable(True, True)
        self.minsize(700, 600)
        self.configure(bg=BG_APP)

        self._xlsx_var    = tk.StringVar()
        self._carpeta_var = tk.StringVar()
        self._dest_var    = tk.StringVar()
        self._col_headers: tuple[str, str] = ("Mail", "Fichero")

        self._firma_img = tk.PhotoImage(data=_FIRMA_B64)

        self._build_styles()
        self._build_ui()
        self.update_idletasks()
        self.state("zoomed")

        self._dest_var.trace_add("write",    self._on_dest_change)
        self._carpeta_var.trace_add("write", self._on_carpeta_change)

        if openpyxl is None:
            self._log_write(
                "AVISO: openpyxl no está instalado.\n"
                "Instala con:  pip install openpyxl\n", "warn")

        last_xlsx    = _reg_read(_REG_XLSX)
        last_carpeta = _reg_read(_REG_CARPETA)
        if last_carpeta and os.path.isdir(last_carpeta):
            self._carpeta_var.set(last_carpeta)
        if last_xlsx and os.path.exists(last_xlsx):
            self._xlsx_var.set(last_xlsx)
            self._lbl_xlsx_title.config(text=os.path.basename(last_xlsx))
            self._load_excel_bg(last_xlsx)

    def _center(self, h: int) -> None:
        self.update_idletasks()
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        self.geometry(f"{sw}x{h}+0+{max(0, (sh - h) // 2)}")

    def _build_styles(self) -> None:
        s = ttk.Style(self)
        s.theme_use("clam")
        s.configure(".",              font=FONT_UI,  background=BG_APP)
        s.configure("TFrame",         background=BG_APP)
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
        s.configure("Accion.TButton",  font=FONT_UI,   background=C_ACCION)
        s.map("Accion.TButton",        background=[("active", "#E59866")])
        _row_h = int(BASE * 2.4)

        # Separador: imagen 1×1 px del color de la cabecera, anclada al borde sur de cada fila
        self._sep_px = tk.PhotoImage(width=1, height=1)
        self._sep_px.put(BG_TABLA, to=(0, 0, 1, 1))
        s.element_create("RowSep", "image", self._sep_px, sticky="ew", border=0)
        try:
            layout = list(s.layout("Treeview.Item"))
            layout.append(("RowSep", {"sticky": "sew"}))
            s.layout("Treeview.Item", layout)
        except Exception:
            pass

        s.configure("Treeview",
                     background="white", fieldbackground="white",
                     foreground="#222222", rowheight=_row_h + 1,
                     font=FONT_UI)
        s.configure("Treeview.Heading",
                     font=FONT_BOLD, background="#DDDDDD", relief="flat")
        s.map("Treeview",
              background=[("selected", "#2E86C1")],
              foreground=[("selected", "white")])

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

        # ── Archivo Excel ────────────────────────────────────────────────────
        tb_arc, c_arc = self._bloque("Archivo Excel", BG_ARCHIVO)
        tb_arc.columnconfigure(1, weight=1)
        self._lbl_xlsx_title = tk.Label(
            tb_arc, text="", bg=BG_ARCHIVO, fg="white", font=FONT_UI, anchor="w")
        self._lbl_xlsx_title.grid(row=0, column=1, sticky="ew", padx=(6, 8))
        _firma_lbl = tk.Label(tb_arc, image=self._firma_img,
                              bg=BG_ARCHIVO, bd=0, cursor="hand2")
        _firma_lbl.grid(row=0, column=2, padx=(0, 8), pady=2)
        _Tooltip(_firma_lbl, "Dugarry")

        tk.Label(c_arc, text="Archivo Excel:", bg=BG_APP, font=FONT_UI, anchor="w").grid(
            row=0, column=0, sticky="w", padx=(0, 10), pady=4)
        tk.Entry(c_arc, textvariable=self._xlsx_var, state="readonly",
                 font=FONT_ENTRY, bg="white", readonlybackground="#ECECEC",
                 fg="#222222", relief="sunken", bd=1).grid(
            row=0, column=1, sticky="ew", pady=4, ipady=4)
        ttk.Button(c_arc, text="Seleccionar", style="Sel.TButton",
                   command=self._pick_file).grid(
            row=0, column=2, padx=(8, 0), pady=4, ipadx=4)

        # ── Carpeta Excels ───────────────────────────────────────────────────
        tb_carp, c_carp = self._bloque("Carpeta Excels", BG_CARPETA)
        tb_carp.columnconfigure(1, weight=1)
        self._lbl_carpeta_title = tk.Label(
            tb_carp, text=_HINT_CARPETA,
            bg=BG_CARPETA, fg="white", font=FONT_UI, anchor="w")
        self._lbl_carpeta_title.grid(row=0, column=1, sticky="ew", padx=(6, 8))

        tk.Label(c_carp, text="Carpeta:", bg=BG_APP, font=FONT_UI, anchor="w").grid(
            row=0, column=0, sticky="w", padx=(0, 10), pady=4)
        tk.Entry(c_carp, textvariable=self._carpeta_var,
                 font=FONT_ENTRY, bg="white", fg="#222222",
                 relief="sunken", bd=1).grid(
            row=0, column=1, sticky="ew", pady=4, ipady=4)
        ttk.Button(c_carp, text="Seleccionar", style="Sel.TButton",
                   command=self._pick_carpeta).grid(
            row=0, column=2, padx=(8, 0), pady=4, ipadx=4)

        # ── Tabla Correos/Ficheros ───────────────────────────────────────────
        tb_tab, c_tab = self._bloque("Tabla Correos/Ficheros", BG_TABLA, expand=True)
        tb_tab.columnconfigure(0, weight=1)
        self._lbl_count = tk.Label(
            tb_tab, text="", bg=BG_TABLA, fg="white", font=FONT_UI)
        self._lbl_count.grid(row=0, column=1, padx=(0, 6))
        ttk.Button(tb_tab, text="Todos",   style="Todos.TButton",
                   command=self._sel_all).grid(
            row=0, column=2, padx=(0, 4), pady=4, ipadx=4)
        ttk.Button(tb_tab, text="Ninguno", style="Ninguno.TButton",
                   command=self._sel_none).grid(
            row=0, column=3, padx=(0, 8), pady=4, ipadx=4)

        c_tab.rowconfigure(0, weight=1)
        tree_frame = tk.Frame(c_tab, bg=BG_APP)
        tree_frame.grid(row=0, column=0, columnspan=3, sticky="nsew")
        tree_frame.rowconfigure(0, weight=1)
        tree_frame.columnconfigure(0, weight=1)

        self._tree = ttk.Treeview(
            tree_frame, columns=_COL_IDS,
            show="headings", selectmode="extended")

        for col_id in _COL_IDS:
            anc = "center" if col_id == "resultado" else "w"
            self._tree.heading(col_id, text=col_id.capitalize(),
                               anchor=anc,
                               command=lambda c=col_id: self._sort_col(c))
            self._tree.column(col_id, width=180, minwidth=60, stretch=True, anchor=anc)

        vsb = ttk.Scrollbar(tree_frame, orient="vertical",   command=self._tree.yview)
        hsb = ttk.Scrollbar(tree_frame, orient="horizontal", command=self._tree.xview)
        self._tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        self._tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self._tree.bind("<<TreeviewSelect>>", self._update_count)
        self._tree.bind("<MouseWheel>",
                        lambda e: self._tree.yview_scroll(-1 if e.delta > 0 else 1, "units"))

        self._sort_reverse: dict[str, bool] = {c: False for c in _COL_IDS}

        # ── Carpeta de Destino ───────────────────────────────────────────────
        tb_dest, c_dest = self._bloque("Carpeta de Destino", BG_DESTINO)
        tb_dest.columnconfigure(1, weight=1)
        self._lbl_dest_title = tk.Label(
            tb_dest, text=_HINT_DEST,
            bg=BG_DESTINO, fg="white", font=FONT_UI, anchor="w")
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
        ttk.Button(c_dest, text="Limpiar", style="Accion.TButton",
                   command=lambda: self._dest_var.set("")).grid(
            row=0, column=3, padx=(6, 0), pady=4, ipadx=4)

        # ── Registro ─────────────────────────────────────────────────────────
        tb_log, c_log = self._bloque("Registro", BG_LOG_HDR)
        tb_log.columnconfigure(0, weight=1)
        c_log.rowconfigure(0, weight=1)
        ttk.Button(tb_log, text="Limpiar", style="Accion.TButton",
                   command=self._log_clear).grid(
            row=0, column=1, padx=(0, 8), pady=4, ipadx=6)
        self._btn_rename = ttk.Button(
            tb_log, text="Renombrar",
            style="Export.TButton", command=self._do_rename)
        self._btn_rename.grid(row=0, column=2, padx=(0, 8), pady=4, ipadx=6)

        self._log = scrolledtext.ScrolledText(
            c_log, height=8, state="disabled", font=FONT_LOG, wrap="word",
            spacing1=5, spacing3=5,
            bg="#1E1E1E", fg="#D4D4D4", insertbackground="white", relief="flat", bd=0)
        self._log.grid(row=0, column=0, columnspan=3, sticky="nsew")
        self._log.tag_configure("ok",    foreground="#4EC94E")
        self._log.tag_configure("error", foreground="#FF6B6B")
        self._log.tag_configure("info",  foreground="#85C1E9")
        self._log.tag_configure("warn",  foreground="#F0B27A")

        tk.Frame(self, bg=BG_APP, height=8).pack()

    # ── Callbacks de selección ────────────────────────────────────────────────

    def _pick_file(self) -> None:
        init = (os.path.dirname(self._xlsx_var.get())
                if self._xlsx_var.get() else os.path.expanduser("~"))
        path = filedialog.askopenfilename(
            parent=self, title="Seleccionar archivo Excel", initialdir=init,
            filetypes=[("Excel", "*.xlsx *.xlsm *.xls"), ("Todos los archivos", "*.*")])
        if not path:
            return
        self._xlsx_var.set(path)
        self._lbl_xlsx_title.config(text=os.path.basename(path))
        _reg_write(_REG_XLSX, path)
        if not self._carpeta_var.get():
            self._carpeta_var.set(os.path.dirname(path))
        self._load_excel_bg(path)

    def _pick_carpeta(self) -> None:
        init = (self._carpeta_var.get()
                if self._carpeta_var.get() and os.path.isdir(self._carpeta_var.get())
                else (os.path.dirname(self._xlsx_var.get())
                      if self._xlsx_var.get() else os.path.expanduser("~")))
        folder = filedialog.askdirectory(
            parent=self, title="Carpeta con los ficheros a renombrar", initialdir=init)
        if folder:
            self._carpeta_var.set(folder)
            _reg_write(_REG_CARPETA, folder)

    def _pick_dest(self) -> None:
        init = (self._carpeta_var.get()
                if self._carpeta_var.get() and os.path.isdir(self._carpeta_var.get())
                else os.path.expanduser("~"))
        folder = filedialog.askdirectory(
            parent=self, title="Carpeta de destino", initialdir=init)
        if folder:
            self._dest_var.set(folder)

    def _on_dest_change(self, *_) -> None:
        dest = self._dest_var.get().strip()
        if dest:
            self._lbl_dest_title.config(text=os.path.basename(dest) or dest)
            self._btn_rename.config(text="Copiar y Renombrar")
        else:
            self._lbl_dest_title.config(text=_HINT_DEST)
            self._btn_rename.config(text="Renombrar")

    def _on_carpeta_change(self, *_) -> None:
        val = self._carpeta_var.get().strip()
        self._lbl_carpeta_title.config(
            text=(os.path.basename(val) or val) if val else _HINT_CARPETA)

    # ── Selección de filas ────────────────────────────────────────────────────

    def _sel_all(self) -> None:
        self._tree.selection_set(self._tree.get_children())

    def _sel_none(self) -> None:
        self._tree.selection_remove(self._tree.get_children())

    def _update_count(self, _=None) -> None:
        sel = len(self._tree.selection())
        tot = len(self._tree.get_children())
        self._lbl_count.config(text=f"{sel} / {tot}" if tot else "")

    # ── Ordenación de columnas ────────────────────────────────────────────────

    def _sort_col(self, col: str) -> None:
        items = [(self._tree.set(iid, col), iid) for iid in self._tree.get_children()]
        rev   = self._sort_reverse.get(col, False)
        items.sort(key=lambda t: t[0].lower(), reverse=rev)
        for idx, (_, iid) in enumerate(items):
            self._tree.move(iid, "", idx)
            self._tree.item(iid, tags=("even" if idx % 2 == 0 else "odd",))
        self._sort_reverse[col] = not rev
        arrow = " ▲" if not rev else " ▼"
        for c in _COL_IDS:
            raw = self._tree.heading(c, "text").rstrip(" ▲▼")
            self._tree.heading(c, text=raw + (arrow if c == col else ""))

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

    # ── Carga del Excel ───────────────────────────────────────────────────────

    def _load_excel_bg(self, path: str) -> None:
        self._log_write(f"\nLeyendo tabla de:  {os.path.basename(path)}\n", "info")
        self._btn_rename.config(state="disabled")
        self._clear_tree()
        threading.Thread(target=self._load_thread, args=(path,), daemon=True).start()

    def _load_thread(self, path: str) -> None:
        try:
            if openpyxl is None:
                raise ImportError("openpyxl no está instalado.\nInstala con:  pip install openpyxl")
            wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
            ws = wb.active
            rows = list(ws.iter_rows(values_only=True))
            wb.close()
            if not rows or len(rows[0]) < 2:
                raise ValueError("La tabla debe tener al menos dos columnas.")
            h0 = str(rows[0][0] or "Mail").strip()
            h1 = str(rows[0][1] or "Fichero").strip()
            data = []
            for r in rows[1:]:
                if not r[0] and not r[1]:
                    continue
                mail    = str(r[0] or "").strip()
                fichero = str(r[1] or "").strip()
                clean   = _mail_clean(mail)
                repl    = f"({clean})" if clean else "()"
                newname = fichero.replace("Indicadores", repl)
                data.append((mail, fichero, newname, "-.-"))
            self.after(0, self._populate_tree, (h0, h1), data)
        except Exception as exc:
            self.after(0, self._on_load_error, str(exc))

    def _populate_tree(self, headers: tuple, data: list) -> None:
        self._clear_tree()
        h0, h1 = headers
        self._col_headers = headers
        labels = (h0, h1) + _COL_FIXED
        for col_id, label in zip(_COL_IDS, labels):
            self._tree.heading(col_id, text=label,
                               command=lambda c=col_id: self._sort_col(c))

        self._sort_reverse = {c: False for c in _COL_IDS}
        for i, row in enumerate(data):
            tag = "even" if i % 2 == 0 else "odd"
            self._tree.insert("", "end", iid=str(i), values=row, tags=(tag,))
        self._tree.tag_configure("odd",  background="#FFFFFF")
        self._tree.tag_configure("even", background="#F4F6F7")

        self._tree.selection_set(self._tree.get_children())
        self._update_count()
        self.after(120, self._autosize_columns, labels)
        self._log_write(f"  -> {len(data)} fila(s) cargada(s)\n", "ok")
        self._btn_rename.config(state="normal")

    def _autosize_columns(self, labels: tuple) -> None:
        self.update_idletasks()
        font_n = tkfont.Font(family=FONT_UI[0],   size=FONT_UI[1])
        font_b = tkfont.Font(family=FONT_BOLD[0], size=FONT_BOLD[1], weight="bold")

        # Col 4 (Resultado): máximo entre el título y los valores posibles, sin stretch
        resultado_vals = ["-.-", "Rename", "CopyRename", "Falló"]
        col4_w = max(font_b.measure(labels[3]),
                     *[font_n.measure(v) for v in resultado_vals]) + 15
        self._tree.column("resultado", width=col4_w, minwidth=col4_w,
                          stretch=False, anchor="center")

        # Col 1 (mail): máximo de todos los valores de la columna + 15px
        col1_w = font_b.measure(labels[0])
        for iid in self._tree.get_children():
            w = font_n.measure(str(self._tree.set(iid, "mail")))
            if w > col1_w:
                col1_w = w
        col1_w += 15
        self._tree.column("mail", width=col1_w, minwidth=col1_w,
                          stretch=False, anchor="w")

        # Cols 2 y 3: reparten el espacio restante del treeview a partes iguales
        tree_w = self._tree.winfo_width()
        if tree_w <= 1:
            tree_w = self.winfo_screenwidth() - 30
        vsb_w     = 18
        remaining = max(200, tree_w - col1_w - col4_w - vsb_w)
        col23_w   = remaining // 2
        self._tree.column("fichero", width=col23_w, minwidth=80,
                          stretch=True, anchor="w")
        self._tree.column("newname", width=col23_w, minwidth=80,
                          stretch=True, anchor="w")

    def _on_load_error(self, msg: str) -> None:
        self._btn_rename.config(state="normal")
        self._log_write(f"\n  ERROR: {msg}\n", "error")
        messagebox.showerror("Error al leer Excel", msg, parent=self)

    def _clear_tree(self) -> None:
        for iid in self._tree.get_children():
            self._tree.delete(iid)
        self._lbl_count.config(text="")

    # ── Acción de renombrar / copiar y renombrar ──────────────────────────────

    def _do_rename(self) -> None:
        xlsx_path = self._xlsx_var.get()
        carpeta   = self._carpeta_var.get().strip()
        dest      = self._dest_var.get().strip()
        copy_mode = bool(dest)

        if not xlsx_path or not os.path.exists(xlsx_path):
            messagebox.showerror("Error", "Selecciona un archivo Excel válido.", parent=self)
            return
        if not carpeta or not os.path.isdir(carpeta):
            messagebox.showerror(
                "Error",
                "Selecciona la carpeta donde están los ficheros a renombrar (Carpeta Excels).",
                parent=self)
            return

        selected = self._tree.selection()
        if not selected:
            messagebox.showwarning("Sin selección",
                                   "Selecciona al menos una fila de la tabla.", parent=self)
            return

        if copy_mode and not os.path.exists(dest):
            try:
                os.makedirs(dest, exist_ok=True)
            except OSError as exc:
                messagebox.showerror("Error",
                                     f"No se pudo crear la carpeta de destino:\n{exc}",
                                     parent=self)
                return

        action_lbl = "Copiando y renombrando" if copy_mode else "Renombrando"
        self._log_write(f"\n{action_lbl} {len(selected)} fichero(s)...\n", "info")
        self._btn_rename.config(state="disabled")

        threading.Thread(
            target=self._rename_thread,
            args=(selected, carpeta, dest, copy_mode),
            daemon=True).start()

    @staticmethod
    def _add_correo_sheet(file_path: str, mail: str) -> None:
        """Abre el Excel, crea/reemplaza la hoja 'Correo' con el mail en A1, veryHidden."""
        ext      = os.path.splitext(file_path)[1].lower()
        keep_vba = ext in (".xlsm", ".xlam", ".xltm")
        wb = openpyxl.load_workbook(file_path, keep_vba=keep_vba)
        if "Correo" in wb.sheetnames:
            del wb["Correo"]
        ws = wb.create_sheet("Correo")
        ws["A1"] = mail
        ws.sheet_state = "veryHidden"
        wb.save(file_path)
        wb.close()

    def _rename_thread(self, selected: tuple, carpeta: str,
                       dest: str, copy_mode: bool) -> None:
        for iid in self._tree.get_children():
            self.after(0, self._tree.set, iid, "resultado", "-.-")

        ok = 0; errors = 0
        for iid in selected:
            vals     = self._tree.item(iid, "values")
            mail     = str(vals[0]).strip()  # col 1: mail
            fichero  = str(vals[1]).strip()  # col 2: nombre original
            new_name = str(vals[2]).strip()  # col 3: New Name
            if not fichero or not new_name:
                self.after(0, self._log_write, "  ⚠  Fila vacía ignorada\n", "warn")
                continue
            orig_path = os.path.join(carpeta, fichero)
            try:
                if not os.path.exists(orig_path):
                    raise FileNotFoundError(f"No encontrado: {fichero}")
                if copy_mode:
                    new_path = os.path.join(dest, new_name)
                    shutil.copy2(orig_path, new_path)
                    self._add_correo_sheet(new_path, mail)
                    result = "CopyRename"
                else:
                    self._add_correo_sheet(orig_path, mail)
                    new_path = os.path.join(carpeta, new_name)
                    os.rename(orig_path, new_path)
                    result = "Rename"
                self.after(0, self._tree.set, iid, "resultado", result)
                self.after(0, self._log_write,
                           f"  ✔  {fichero}  →  {new_name}\n", "ok")
                ok += 1
            except Exception as exc:
                self.after(0, self._tree.set, iid, "resultado", "Falló")
                self.after(0, self._log_write,
                           f"  ✖  {fichero}: {exc}\n", "error")
                errors += 1
        self.after(0, self._rename_done, ok, errors, copy_mode)

    def _rename_done(self, ok: int, errors: int, copy_mode: bool) -> None:
        self._btn_rename.config(state="normal")
        action = "copiado(s) y renombrado(s)" if copy_mode else "renombrado(s)"
        suffix = f", {errors} error(es)" if errors else ""
        self._log_write(
            f"\n  Completado: {ok} fichero(s) {action}{suffix}.\n",
            "warn" if errors else "ok")
        if errors:
            messagebox.showwarning(
                "Completado con errores",
                f"{ok} fichero(s) {action}.\n{errors} error(es) — ver Registro.",
                parent=self)
        else:
            messagebox.showinfo(
                "Completado",
                f"{ok} fichero(s) {action} correctamente.",
                parent=self)


if __name__ == "__main__":
    App().mainloop()
