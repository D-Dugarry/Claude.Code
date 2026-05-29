"""
Rename_Files.py  v3
Renombra (o copia y renombra) ficheros de una carpeta aplicando
una sustitución de texto en sus nombres.
Requiere: Python 3.9+, tkinter (stdlib)
"""

import os
import re
import shutil
import sys
import threading
import tkinter as tk
import tkinter.font as tkfont
from tkinter import filedialog, messagebox, scrolledtext, ttk
import winreg

_REG_KEY     = r"Software\RenameFiles"
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
BG_CARPETA = "#1A5276"
BG_SUSTIT  = "#2874A6"
BG_TABLA   = "#7D3C98"
BG_DESTINO = "#1E8449"
BG_LOG_HDR = "#5D6D7E"

C_SELEC    = "#D5D8DC"
C_TODOS    = "#AED6F1"
C_NINGUNO  = "#F5B7B1"
C_EXPORTAR = "#A9DFBF"
C_ACCION   = "#F0B27A"

_HINT_DEST    = "(opcional — si vacío, renombra en el origen)"
_HINT_CARPETA = "(carpeta con los ficheros a renombrar)"

_COL_IDS   = ("nombre", "newname", "resultado")
_COL_NAMES = ("Nombre Actual", "New Name", "Resultado")

_FLUJO_MD = """\
# Rename Files — Flujo del Programa

## Visión general

Herramienta de escritorio (Python/tkinter, compilada a `.exe`) que renombra o copia-y-renombra ficheros de una carpeta aplicando una sustitución de texto simple en sus nombres.

---

## Pantalla principal — secciones

```
┌─────────────────────────────────────────────────────────┐
│  Carpeta de Ficheros   [ruta/carpeta/origen]  Seleccionar│
├─────────────────────────────────────────────────────────┤
│  Sustitución en nombres de fichero                       │
│    Texto a sustituir:  [_____________________]           │
│    Texto nuevo:        [_____________________]           │
├─────────────────────────────────────────────────────────┤
│  Ficheros en Carpeta          N / Total  Todos  Ninguno  │
│  ┌───────────────┬────────────────────┬──────────┐       │
│  │ Nombre Actual │ New Name           │ Resultado│       │
│  │ fichero_A.xlsx│ fichero_nuevo.xlsx │  -.-     │       │
│  │ fichero_B.xlsx│ fichero_nuevo.xlsx │  -.-     │       │
│  └───────────────┴────────────────────┴──────────┘       │
├─────────────────────────────────────────────────────────┤
│  Carpeta de Destino    [ruta/destino]  Examinar  Limpiar │
├─────────────────────────────────────────────────────────┤
│  Registro              Limpiar   [Renombrar / Copiar...] │
│  > log de operaciones                                    │
└─────────────────────────────────────────────────────────┘
```

---

## Flujo paso a paso

### 1 · Seleccionar Carpeta de Ficheros

- El usuario pulsa **Seleccionar** o escribe la ruta directamente en el campo.
- La ruta se guarda en el registro de Windows (`HKCU\\Software\\RenameFiles\\LastCarpetaExcels`) para recordarla en la próxima sesión.
- Cambiar la ruta dispara `_on_params_change` → espera 300 ms (debounce) → llama a `_load_files`.

### 2 · Escribir la sustitución

- **Texto a sustituir**: cadena que se buscará en cada nombre de fichero.
- **Texto nuevo**: cadena por la que se reemplazará.
- Cada pulsación de tecla dispara `_on_params_change` → debounce 300 ms → `_load_files`.
- Si "Texto a sustituir" está vacío, `New Name` = `Nombre Actual` (sin cambio).

### 3 · Carga y visualización de ficheros (`_load_files`)

```
_load_files()
    │
    ├─ Lee os.listdir(carpeta) → solo ficheros (no subdirectorios)
    ├─ Ordena alfabéticamente
    ├─ Para cada fichero:
    │       new_name = nombre.replace(sustituir, nuevo)
    │       fila = (nombre, new_name, "-.-")
    │
    └─ _populate_tree(filas)
            │
            ├─ Inserta filas en el Treeview con tags even/odd
            ├─ Selecciona todas las filas por defecto
            ├─ _on_sel_change()  → aplica colores de selección
            └─ after(120ms) → _autosize_columns()
```

### 4 · Ajuste de columnas (`_autosize_columns`)

| Columna        | Anchura                                         | Stretch |
|----------------|-------------------------------------------------|---------|
| `Resultado`    | Fija: máx. entre título y valores posibles +15px | No      |
| `Nombre Actual`| Fija: máx. de todos los nombres de fichero +15px | No      |
| `New Name`     | Todo el espacio restante de la ventana           | Sí      |

### 5 · Selección de filas

- **Todos** / **Ninguno**: selecciona o deselecciona todas las filas.
- **Clic individual**: selecciona una fila (deselecciona el resto).
- **Shift + clic**: selección de rango.
- **Ctrl + clic**: añade/quita una fila a la selección.
- Cualquier cambio de selección llama a `_on_sel_change`:
  - Actualiza el tag de cada fila (even / odd / even_sel / odd_sel).
  - Actualiza el contador `N / Total` en la cabecera.

### 6 · Carpeta de Destino (opcional)

| Estado del campo | Comportamiento          | Texto del botón      |
|------------------|-------------------------|----------------------|
| Vacío            | Renombra en la carpeta origen | `Renombrar`    |
| Con ruta         | Copia los ficheros con el nuevo nombre a esa carpeta | `Copiar y Renombrar` |

### 7 · Operación: Renombrar / Copiar y Renombrar (`_do_rename`)

**Validaciones previas:**
- Carpeta de ficheros existe.
- Al menos una fila seleccionada.
- Si hay destino y no existe, se crea con `os.makedirs`.

**Hilo de trabajo (`_rename_thread`)** — se ejecuta en un hilo secundario para no bloquear la UI:

```
Para cada fila seleccionada:
    │
    ├─ Construye orig_path = carpeta / Nombre Actual
    ├─ Verifica que el fichero existe
    │
    ├─ [Modo Renombrar]
    │       os.rename(orig_path, carpeta / New Name)
    │       → Resultado = "Rename"
    │
    └─ [Modo Copiar y Renombrar]
            shutil.copy2(orig_path, destino / New Name)
            → Resultado = "CopyRename"

    Si cualquier paso falla:
        → Resultado = "Falló"
        → Mensaje en el Registro en rojo
```

**Actualización de la columna Resultado** (en el hilo principal vía `after(0, ...)`):

| Valor        | Significado                              |
|--------------|------------------------------------------|
| `-.-`        | Fila no procesada (no seleccionada o reset) |
| `Rename`     | Renombrado con éxito en origen           |
| `CopyRename` | Copiado y renombrado en destino          |
| `Falló`      | Error — ver detalle en el Registro       |

---

## Colores del Treeview

| Tag         | Fondo     | Texto    | Cuándo                        |
|-------------|-----------|----------|-------------------------------|
| `even`      | `#FFFFFF` | `#222222`| Fila par, no seleccionada     |
| `odd`       | `#E8EDF2` | `#222222`| Fila impar, no seleccionada   |
| `even_sel`  | `#5B9BD5` | blanco   | Fila par, seleccionada        |
| `odd_sel`   | `#2E75B6` | blanco   | Fila impar, seleccionada      |

> Los tags anulan el color de selección nativo del tema clam.

---

## Persistencia (Registro de Windows)

| Clave                                          | Valor guardado                   |
|------------------------------------------------|----------------------------------|
| `HKCU\\Software\\RenameFiles\\LastCarpetaExcels` | Última carpeta de ficheros usada |

---

## Archivos del proyecto

| Fichero             | Descripción                          |
|---------------------|--------------------------------------|
| `Rename_Files.py`   | Código fuente principal              |
| `Rename_Files.exe`  | Ejecutable compilado (PyInstaller)   |
| `build.bat`         | Script de recompilación              |
| `requirements.txt`  | Dependencias Python (`pyinstaller`)  |
| `paleta_colores.py` | Utilidad auxiliar para elegir colores|
| `Flujo_APP.md`      | Documentación del flujo              |
"""

_INLINE_RE = re.compile(r'\*\*(.+?)\*\*|`(.+?)`')


def _insert_inline(txt: "tk.Text", text: str, base_tag: str) -> None:
    """Inserta texto aplicando bold (**...**) e inline-code (`...`) como tags."""
    pos = 0
    for m in _INLINE_RE.finditer(text):
        if m.start() > pos:
            txt.insert("end", text[pos:m.start()], base_tag)
        if m.group(1) is not None:
            txt.insert("end", m.group(1), (base_tag, "md_bold"))
        else:
            txt.insert("end", m.group(2), (base_tag, "md_icode"))
        pos = m.end()
    if pos < len(text):
        txt.insert("end", text[pos:], base_tag)


def _base_dir() -> str:
    """Directorio del exe (compilado) o del script (desarrollo)."""
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


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
        _BG = BG_CARPETA          # mismo color que la cabecera → transparente
        self._win = tk.Toplevel(self._w)
        self._win.wm_overrideredirect(True)
        self._win.configure(bg=_BG)
        try:
            self._win.wm_attributes("-transparentcolor", _BG)
        except Exception:
            pass
        tk.Label(self._win, text=self._text, bg=_BG, fg="#C8A96E",
                 relief="flat", bd=0,
                 font=("Mistral", BASE + 4), padx=6, pady=3).pack()
        self._win.update_idletasks()
        tip_w = self._win.winfo_reqwidth()
        tip_h = self._win.winfo_reqheight()
        x = self._w.winfo_rootx() - tip_w - 6
        y = self._w.winfo_rooty() + (self._w.winfo_height() - tip_h) // 2
        self._win.wm_geometry(f"+{x}+{y}")

    def _hide(self, _=None) -> None:
        if self._win:
            self._win.destroy()
            self._win = None


class App(tk.Tk):

    H = 820

    def __init__(self):
        super().__init__()
        self.title("Rename Files")
        self.resizable(True, True)
        self.minsize(700, 540)
        self.configure(bg=BG_APP)

        self._carpeta_var  = tk.StringVar()
        self._sustituir_var = tk.StringVar()
        self._nuevo_var    = tk.StringVar()
        self._dest_var     = tk.StringVar()
        self._after_id     = None   # para debounce del refresco

        self._firma_img = tk.PhotoImage(data=_FIRMA_B64)

        self._build_styles()
        self._build_ui()
        self.update_idletasks()
        self.state("zoomed")

        self._carpeta_var.trace_add("write",   self._on_params_change)
        self._sustituir_var.trace_add("write", self._on_params_change)
        self._nuevo_var.trace_add("write",     self._on_params_change)
        self._dest_var.trace_add("write",      self._on_dest_change)

        last = _reg_read(_REG_CARPETA)
        if last and os.path.isdir(last):
            self._carpeta_var.set(last)

    # ── Estilos ───────────────────────────────────────────────────────────────

    def _build_styles(self) -> None:
        s = ttk.Style(self)
        s.theme_use("clam")
        self.tk.eval(
            "ttk::style theme settings clam "
            "{ ttk::style map Treeview -background {} }"
        )
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
        s.map("Treeview", foreground=[("selected", "white")])

    # ── UI ────────────────────────────────────────────────────────────────────

    def _bloque(self, titulo: str, bg_titulo: str,
                expand: bool = False) -> tuple[tk.Frame, tk.Frame]:
        outer = tk.Frame(self, bg=BG_APP, bd=1, relief="solid",
                         highlightbackground="#AAAAAA", highlightthickness=1)
        outer.pack(fill="both" if expand else "x",
                   expand=expand, padx=10, pady=(6, 0))
        tb = tk.Frame(outer, bg=bg_titulo)
        tb.pack(fill="x")
        tk.Label(tb, text=f"  {titulo}",
                 bg=bg_titulo, fg="white", font=FONT_TITLE, anchor="w", pady=2
                 ).grid(row=0, column=0, sticky="w")
        c = tk.Frame(outer, bg=BG_APP, padx=12, pady=8)
        c.pack(fill="both", expand=True)
        c.columnconfigure(1, weight=1)
        return tb, c

    def _build_ui(self) -> None:

        # ── Carpeta Excels ───────────────────────────────────────────────────
        tb_carp, c_carp = self._bloque("Carpeta de Ficheros", BG_CARPETA)
        tb_carp.columnconfigure(1, weight=1)
        self._lbl_carpeta_title = tk.Label(
            tb_carp, text=_HINT_CARPETA,
            bg=BG_CARPETA, fg="white", font=FONT_UI, anchor="w")
        self._lbl_carpeta_title.grid(row=0, column=1, sticky="ew", padx=(6, 8))
        _firma_lbl = tk.Label(tb_carp, image=self._firma_img,
                              bg=BG_CARPETA, bd=0, cursor="hand2")
        _firma_lbl.grid(row=0, column=2, padx=(0, 8), pady=2)
        _firma_lbl.bind("<Button-1>", lambda _: self._show_flujo())
        _Tooltip(_firma_lbl, "Dugarry'26")

        tk.Label(c_carp, text="Carpeta:", bg=BG_APP, font=FONT_UI, anchor="w").grid(
            row=0, column=0, sticky="w", padx=(0, 10), pady=4)
        tk.Entry(c_carp, textvariable=self._carpeta_var,
                 font=FONT_ENTRY, bg="white", fg="#222222",
                 relief="sunken", bd=1).grid(
            row=0, column=1, sticky="ew", pady=4, ipady=4)
        ttk.Button(c_carp, text="Seleccionar", style="Sel.TButton",
                   command=self._pick_carpeta).grid(
            row=0, column=2, padx=(8, 0), pady=4, ipadx=4)

        # ── Sustitución ──────────────────────────────────────────────────────
        tb_sus, c_sus = self._bloque("Sustitución en nombres de fichero", BG_SUSTIT)
        tb_sus.columnconfigure(1, weight=1)
        self._lbl_sus_preview = tk.Label(
            tb_sus, text="", bg=BG_SUSTIT, fg="white", font=FONT_UI, anchor="w")
        self._lbl_sus_preview.grid(row=0, column=1, sticky="ew", padx=(6, 8))

        tk.Label(c_sus, text="Texto a sustituir:", bg=BG_APP,
                 font=FONT_UI, anchor="w").grid(
            row=0, column=0, sticky="w", padx=(0, 10), pady=(4, 2))
        tk.Entry(c_sus, textvariable=self._sustituir_var,
                 font=FONT_ENTRY, bg="white", fg="#222222",
                 relief="sunken", bd=1).grid(
            row=0, column=1, sticky="ew", pady=(4, 2), ipady=4)

        tk.Label(c_sus, text="Texto nuevo:", bg=BG_APP,
                 font=FONT_UI, anchor="w").grid(
            row=1, column=0, sticky="w", padx=(0, 10), pady=(2, 4))
        tk.Entry(c_sus, textvariable=self._nuevo_var,
                 font=FONT_ENTRY, bg="white", fg="#222222",
                 relief="sunken", bd=1).grid(
            row=1, column=1, sticky="ew", pady=(2, 4), ipady=4)

        # ── Ficheros en Carpeta ──────────────────────────────────────────────
        tb_tab, c_tab = self._bloque("Ficheros en Carpeta", BG_TABLA, expand=True)
        self._tab_bloque = tb_tab.master   # referencia al bloque exterior
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
        for col_id, col_name in zip(_COL_IDS, _COL_NAMES):
            anc = "center" if col_id == "resultado" else "w"
            self._tree.heading(col_id, text=col_name, anchor=anc,
                               command=lambda c=col_id: self._sort_col(c))
            self._tree.column(col_id, width=200, minwidth=60,
                              stretch=True, anchor=anc)

        vsb = ttk.Scrollbar(tree_frame, orient="vertical",   command=self._tree.yview)
        hsb = ttk.Scrollbar(tree_frame, orient="horizontal", command=self._tree.xview)
        self._tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        self._tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        self._tree.tag_configure("even",     background="#FFFFFF", foreground="#222222")
        self._tree.tag_configure("odd",      background="#E8EDF2", foreground="#222222")
        self._tree.tag_configure("even_sel", background="#5B9BD5", foreground="white")
        self._tree.tag_configure("odd_sel",  background="#2E75B6", foreground="white")

        self._tree.bind("<<TreeviewSelect>>", self._on_sel_change)
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
            c_log, height=6, state="disabled", font=FONT_LOG, wrap="word",
            spacing1=5, spacing3=5,
            bg="#1E1E1E", fg="#D4D4D4", insertbackground="white",
            relief="flat", bd=0)
        self._log.grid(row=0, column=0, columnspan=3, sticky="nsew")
        self._log.tag_configure("ok",    foreground="#4EC94E")
        self._log.tag_configure("error", foreground="#FF6B6B")
        self._log.tag_configure("info",  foreground="#85C1E9")
        self._log.tag_configure("warn",  foreground="#F0B27A")

        tk.Frame(self, bg=BG_APP, height=8).pack()

    # ── Carpeta / destino ─────────────────────────────────────────────────────

    def _show_flujo(self) -> None:
        if hasattr(self, "_flujo_panel") and self._flujo_panel.winfo_ismapped():
            self._hide_flujo()
            return

        # Crear el panel integrado la primera vez
        if not hasattr(self, "_flujo_panel"):
            pan = tk.Frame(self._tab_bloque, bg=BG_APP)
            self._flujo_panel = pan

            hdr = tk.Frame(pan, bg=BG_TABLA)
            hdr.pack(fill="x")
            tk.Label(hdr, text="  Flujo de la Aplicación",
                     bg=BG_TABLA, fg="white", font=FONT_TITLE,
                     anchor="w", pady=4).pack(side="left", fill="x", expand=True)
            tk.Button(hdr, text="  ✕  ", command=self._hide_flujo,
                      bg=BG_TABLA, fg="white",
                      activebackground="#C0392B", activeforeground="white",
                      relief="flat", bd=0, font=FONT_BOLD, cursor="hand2",
                      ).pack(side="right", padx=4)

            self._flujo_txt = scrolledtext.ScrolledText(
                pan, font=FONT_UI, wrap="word",
                bg=BG_APP, fg="#222222", relief="flat", bd=0,
                spacing1=1, spacing3=1, padx=14, pady=6)
            self._flujo_txt.pack(fill="both", expand=True)
            self._render_md(self._flujo_txt, _FLUJO_MD)

        self._flujo_panel.place(x=0, y=0, relwidth=1, relheight=1)
        self._flujo_panel.lift()

    def _hide_flujo(self) -> None:
        if hasattr(self, "_flujo_panel"):
            self._flujo_panel.place_forget()

    def _render_md(self, txt: scrolledtext.ScrolledText, content: str) -> None:
        """Renderiza markdown básico con colores en un widget Text."""
        txt.tag_configure("h1",      font=("Verdana", BASE + 4, "bold"),
                          foreground=BG_TABLA,   spacing1=12, spacing3=6)
        txt.tag_configure("h2",      font=("Verdana", BASE + 2, "bold"),
                          foreground=BG_CARPETA, spacing1=10, spacing3=4)
        txt.tag_configure("h3",      font=("Verdana", BASE, "bold"),
                          foreground=BG_DESTINO, spacing1=8,  spacing3=2)
        txt.tag_configure("code",    font=("Courier New", BASE - 1),
                          background="#E8EDF2", foreground="#2C3E50",
                          lmargin1=20, lmargin2=20, spacing1=1, spacing3=1)
        txt.tag_configure("table",   font=("Courier New", BASE - 1),
                          background="#F4F6F7", foreground="#2C3E50",
                          lmargin1=10, lmargin2=10)
        txt.tag_configure("tabsep",  font=("Courier New", BASE - 1),
                          background="#D5D8DC", foreground="#7F8C8D",
                          lmargin1=10, lmargin2=10)
        txt.tag_configure("bullet",  font=FONT_UI, foreground="#2C3E50",
                          lmargin1=24, lmargin2=40)
        txt.tag_configure("rule",    font=("Verdana", 4),
                          foreground=BG_TABLA, spacing1=6, spacing3=6)
        txt.tag_configure("quote",   font=("Verdana", BASE, "italic"),
                          foreground="#5D6D7E", lmargin1=30, lmargin2=30)
        txt.tag_configure("normal",  font=FONT_UI, foreground="#222222")
        txt.tag_configure("md_bold", font=FONT_BOLD)
        txt.tag_configure("md_icode",font=("Courier New", BASE - 1),
                          foreground="#C0392B")

        txt.configure(state="normal")
        txt.delete("1.0", "end")

        in_code = False
        for line in content.split("\n"):
            stripped = line.strip()

            if stripped == "```":
                in_code = not in_code
                if not in_code:
                    txt.insert("end", "\n")
                continue

            if in_code:
                txt.insert("end", line + "\n", "code")
                continue

            if stripped in ("---", "***", "___"):
                txt.insert("end", "─" * 72 + "\n", "rule")
                continue

            if line.startswith("### "):
                _insert_inline(txt, line[4:] + "\n", "h3"); continue
            if line.startswith("## "):
                _insert_inline(txt, line[3:] + "\n", "h2"); continue
            if line.startswith("# "):
                _insert_inline(txt, line[2:] + "\n", "h1"); continue

            if line.startswith("|"):
                tag = "tabsep" if all(c in "-|: " for c in line) else "table"
                txt.insert("end", line + "\n", tag)
                continue

            if line.startswith("> "):
                _insert_inline(txt, line[2:] + "\n", "quote"); continue

            if re.match(r"^( {0,4})-\s", line):
                _insert_inline(txt, "  •  " + line.lstrip("- ").lstrip() + "\n",
                               "bullet"); continue

            if stripped:
                _insert_inline(txt, line + "\n", "normal")
            else:
                txt.insert("end", "\n", "normal")

        txt.configure(state="disabled")

    def _pick_carpeta(self) -> None:
        init = (self._carpeta_var.get()
                if self._carpeta_var.get() and os.path.isdir(self._carpeta_var.get())
                else os.path.expanduser("~"))
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

    # ── Parámetros → recarga con debounce ─────────────────────────────────────

    def _on_params_change(self, *_) -> None:
        val = self._carpeta_var.get().strip()
        self._lbl_carpeta_title.config(
            text=(os.path.basename(val) or val) if val else _HINT_CARPETA)
        if self._after_id is not None:
            self.after_cancel(self._after_id)
        self._after_id = self.after(300, self._load_files)

    # ── Selección de filas ────────────────────────────────────────────────────

    def _sel_all(self) -> None:
        self._tree.selection_set(self._tree.get_children())

    def _sel_none(self) -> None:
        self._tree.selection_remove(self._tree.get_children())

    def _on_sel_change(self, _=None) -> None:
        selection = set(self._tree.selection())
        for iid in self._tree.get_children():
            even = self._tree.index(iid) % 2 == 0
            if iid in selection:
                tag = "even_sel" if even else "odd_sel"
            else:
                tag = "even" if even else "odd"
            self._tree.item(iid, tags=(tag,))
        sel = len(selection)
        tot = len(self._tree.get_children())
        self._lbl_count.config(text=f"{sel} / {tot}" if tot else "")

    def _update_count(self) -> None:
        sel = len(self._tree.selection())
        tot = len(self._tree.get_children())
        self._lbl_count.config(text=f"{sel} / {tot}" if tot else "")

    # ── Ordenación ────────────────────────────────────────────────────────────

    def _sort_col(self, col: str) -> None:
        items = [(self._tree.set(iid, col), iid) for iid in self._tree.get_children()]
        rev   = self._sort_reverse.get(col, False)
        items.sort(key=lambda t: t[0].lower(), reverse=rev)
        for idx, (_, iid) in enumerate(items):
            self._tree.move(iid, "", idx)
        self._sort_reverse[col] = not rev
        arrow = " ▲" if not rev else " ▼"
        for c in _COL_IDS:
            raw = self._tree.heading(c, "text").rstrip(" ▲▼")
            self._tree.heading(c, text=raw + (arrow if c == col else ""))
        self._on_sel_change()

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

    # ── Carga de ficheros ─────────────────────────────────────────────────────

    def _load_files(self) -> None:
        self._after_id = None
        carpeta   = self._carpeta_var.get().strip()
        sustituir = self._sustituir_var.get()
        nuevo     = self._nuevo_var.get()

        self._clear_tree()
        if not carpeta or not os.path.isdir(carpeta):
            return

        try:
            files = sorted(
                f for f in os.listdir(carpeta)
                if os.path.isfile(os.path.join(carpeta, f)))
        except PermissionError as exc:
            self._log_write(f"\n  ERROR: {exc}\n", "error")
            return

        rows = []
        for f in files:
            new_name = f.replace(sustituir, nuevo) if sustituir else f
            rows.append((f, new_name, "-.-"))

        self._populate_tree(rows)
        self._log_write(
            f"\n{len(rows)} fichero(s) en  {carpeta}\n", "info")

    def _populate_tree(self, rows: list) -> None:
        self._clear_tree()
        self._sort_reverse = {c: False for c in _COL_IDS}
        for col_id, col_name in zip(_COL_IDS, _COL_NAMES):
            self._tree.heading(col_id, text=col_name,
                               command=lambda c=col_id: self._sort_col(c))

        for i, row in enumerate(rows):
            tag = "even" if i % 2 == 0 else "odd"
            self._tree.insert("", "end", iid=str(i), values=row, tags=(tag,))

        self._tree.selection_set(self._tree.get_children())
        self._on_sel_change()
        self.after(120, self._autosize_columns)

    def _autosize_columns(self) -> None:
        self.update_idletasks()
        font_n = tkfont.Font(family=FONT_UI[0],   size=FONT_UI[1])
        font_b = tkfont.Font(family=FONT_BOLD[0], size=FONT_BOLD[1], weight="bold")

        # Resultado: fijo al máximo de sus valores posibles
        res_vals = ["-.-", "Rename", "CopyRename", "Falló"]
        col_res  = max(font_b.measure("Resultado"),
                       *[font_n.measure(v) for v in res_vals]) + 15
        self._tree.column("resultado", width=col_res, minwidth=col_res,
                          stretch=False, anchor="center")

        # Nombre Actual: ajuste al contenido más ancho
        col_nom = font_b.measure("Nombre Actual")
        for iid in self._tree.get_children():
            w = font_n.measure(str(self._tree.set(iid, "nombre")))
            if w > col_nom:
                col_nom = w
        col_nom += 15
        self._tree.column("nombre", width=col_nom, minwidth=80,
                          stretch=False, anchor="w")

        # New Name: ocupa el espacio restante
        tree_w    = self._tree.winfo_width()
        if tree_w <= 1:
            tree_w = self.winfo_screenwidth() - 30
        remaining = max(150, tree_w - col_nom - col_res - 18)
        self._tree.column("newname", width=remaining, minwidth=80,
                          stretch=True, anchor="w")

    def _clear_tree(self) -> None:
        for iid in self._tree.get_children():
            self._tree.delete(iid)
        self._lbl_count.config(text="")

    # ── Renombrar / Copiar y Renombrar ────────────────────────────────────────

    def _do_rename(self) -> None:
        carpeta = self._carpeta_var.get().strip()
        dest    = self._dest_var.get().strip()

        if not carpeta or not os.path.isdir(carpeta):
            messagebox.showerror(
                "Error", "Selecciona la carpeta con los ficheros.", parent=self)
            return

        selected = self._tree.selection()
        if not selected:
            messagebox.showwarning(
                "Sin selección", "Selecciona al menos un fichero.", parent=self)
            return

        copy_mode = bool(dest)
        if copy_mode and not os.path.exists(dest):
            try:
                os.makedirs(dest, exist_ok=True)
            except OSError as exc:
                messagebox.showerror(
                    "Error", f"No se pudo crear la carpeta de destino:\n{exc}",
                    parent=self)
                return

        action = "Copiando y renombrando" if copy_mode else "Renombrando"
        self._log_write(f"\n{action} {len(selected)} fichero(s)...\n", "info")
        self._btn_rename.config(state="disabled")

        threading.Thread(
            target=self._rename_thread,
            args=(selected, carpeta, dest, copy_mode),
            daemon=True).start()

    def _rename_thread(self, selected: tuple, carpeta: str,
                       dest: str, copy_mode: bool) -> None:
        for iid in self._tree.get_children():
            self.after(0, self._tree.set, iid, "resultado", "-.-")

        ok = 0; errors = 0
        for iid in selected:
            vals     = self._tree.item(iid, "values")
            nombre   = str(vals[0]).strip()
            new_name = str(vals[1]).strip()
            if not nombre or not new_name:
                self.after(0, self._log_write, "  ⚠  Fila vacía ignorada\n", "warn")
                continue
            orig_path = os.path.join(carpeta, nombre)
            try:
                if not os.path.exists(orig_path):
                    raise FileNotFoundError(f"No encontrado: {nombre}")
                if copy_mode:
                    new_path = os.path.join(dest, new_name)
                    shutil.copy2(orig_path, new_path)
                    result = "CopyRename"
                else:
                    new_path = os.path.join(carpeta, new_name)
                    os.rename(orig_path, new_path)
                    result = "Rename"
                self.after(0, self._tree.set, iid, "resultado", result)
                self.after(0, self._log_write,
                           f"  ✔  {nombre}  →  {new_name}\n", "ok")
                ok += 1
            except Exception as exc:
                self.after(0, self._tree.set, iid, "resultado", "Falló")
                self.after(0, self._log_write,
                           f"  ✖  {nombre}: {exc}\n", "error")
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
