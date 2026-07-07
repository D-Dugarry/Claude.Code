"""
data_table.py — Tabla de datos reutilizable para tkinter (Windows), sobre
ttk.Treeview. Extraída del skill tkinter-app-design y generalizada.

Encapsula el comportamiento de tabla ya definido en aquel proyecto:
  - Treeview con scrollbars (vertical y horizontal) + rueda del ratón.
  - Franjas alternas (row_a / row_b) re-pintadas también al ordenar.
  - Color de fila seleccionada por TAGS de item (variante "<tag>_sel"),
    NUNCA por style.map("Treeview", ...selected...). Esto es la clave: si se
    mapea el estado "selected" en el estilo, ese mapeo gana siempre al tag de
    la fila y todas las filas seleccionadas se pintan igual (bug real).
  - Autoajuste de columnas al ancho real del contenido (cabecera en negrita
    vs. valores), repartiendo el hueco sobrante como aire extra entre columnas.
  - Ordenación al pulsar la cabecera (numérica si puede, si no alfabética).
  - Selección Todos / Ninguno y contador "seleccionadas/total" vía callback.

Uso mínimo:
    from data_table import install_treeview_style, DataTable
    install_treeview_style(root)                       # una vez por app
    t = DataTable(parent, ("nombre", "tam"), ("Nombre", "Tamaño"),
                  on_select=lambda sel, tot: print(f"{sel}/{tot}"))
    t.pack(fill="both", expand=True)
    t.load([("alfa", "10 KB"), ("beta", "5 KB")])

Tabla por ESTADO (colores por fila en vez de franjas):
    t = DataTable(parent, ids, names, stripe=False)
    t.set_tag("ok",    "#D5F5E3", "#82E0AA")
    t.set_tag("error", "#FADBD8", "#F1948A")
    t.load(rows, tags=["ok", "error", ...])
"""

# Última actualización: 2026-07-07 19:25

import tkinter as tk
import tkinter.font as tkfont
from tkinter import ttk

# ── Valores por defecto (paleta heredada de tkinter-app-design) ──────────────
BASE      = 14
FONT_UI   = ("Verdana", BASE)
FONT_BOLD = ("Verdana", BASE, "bold")
ROW_H     = int(BASE * 2.4) + 1

STRIPE_A     = "white"      # fila par
STRIPE_B     = "#EBEBEB"    # fila impar
STRIPE_A_SEL = "#AED6F1"    # fila par seleccionada
STRIPE_B_SEL = "#85C1E9"    # fila impar seleccionada
SEL_FG       = "#1A1A1A"    # texto de fila seleccionada
SEP_COLOR    = "#D5D8DC"    # línea separadora inferior de cada fila
HEAD_BG      = "#DDDDDD"    # fondo de la cabecera de columnas
GRID_BG      = "#F2F3F4"    # fondo del frame contenedor

_keepalive: list = []        # mantiene vivas las PhotoImage del separador


def install_treeview_style(root: tk.Misc, *,
                           font_ui=FONT_UI, font_bold=FONT_BOLD,
                           row_height=ROW_H, sep_color=SEP_COLOR,
                           head_bg=HEAD_BG, fg="#222222") -> ttk.Style:
    """Configura el estilo ttk "Treeview" para toda la aplicación. Llamar UNA
    vez tras crear la ventana raíz y antes de instanciar DataTable.

    Lo importante (y lo que suele olvidarse): limpia el mapa nativo del tema
    "clam" para el estado 'selected' y deja s.map("Treeview", background=[])
    para que el color de la fila seleccionada lo gobiernen los tags de item.
    """
    s = ttk.Style(root)
    try:
        if s.theme_use() != "clam":
            s.theme_use("clam")
    except tk.TclError:
        pass

    # Limpia el mapa nativo de "clam" para 'selected' (a nivel de tema).
    root.tk.eval("ttk::style theme settings clam "
                 "{ ttk::style map Treeview -background {} }")

    # Separador inferior de 1px por fila (elemento de imagen en el layout).
    try:
        px = tk.PhotoImage(master=root, width=1, height=1)
        px.put(sep_color, to=(0, 0, 1, 1))
        _keepalive.append(px)
        s.element_create("DT.RowSep", "image", px, sticky="ew", border=0)
        layout = list(s.layout("Treeview.Item"))
        layout.append(("DT.RowSep", {"sticky": "sew"}))
        s.layout("Treeview.Item", layout)
    except tk.TclError:
        pass   # ya instalado (segunda llamada): no pasa nada

    s.configure("Treeview", background="white", fieldbackground="white",
                foreground=fg, rowheight=row_height, font=font_ui)
    s.configure("Treeview.Heading", font=font_bold,
                background=head_bg, relief="flat")
    # Nunca reintroducir aquí un color para 'selected': ganaría a los tags.
    s.map("Treeview", background=[], foreground=[])
    return s


class DataTable(tk.Frame):
    """Frame con un ttk.Treeview listo para usar. Ver docstring del módulo."""

    def __init__(self, parent: tk.Widget,
                 col_ids: tuple, col_names: tuple, *,
                 stripe: bool = True, sortable: bool = True,
                 selectmode: str = "extended",
                 on_select=None, anchor: str = "w", bg: str = GRID_BG,
                 stripe_a=STRIPE_A, stripe_b=STRIPE_B,
                 stripe_a_sel=STRIPE_A_SEL, stripe_b_sel=STRIPE_B_SEL,
                 sel_fg=SEL_FG, font_ui=FONT_UI, font_bold=FONT_BOLD,
                 pad: int = 15):
        super().__init__(parent, bg=bg)
        self.col_ids   = tuple(col_ids)
        self.col_names = tuple(col_names)
        self._stripe   = stripe
        self._on_select = on_select
        self._font_ui   = font_ui
        self._font_bold = font_bold
        self._pad       = pad
        self._fixed: dict[str, int] = {}             # col_id -> ancho fijo
        self._sticky: dict[str, str] = {}            # iid -> tag fijo (no franja)
        self._sort_rev = {c: False for c in self.col_ids}

        self.rowconfigure(0, weight=1)
        self.columnconfigure(0, weight=1)

        self.tree = ttk.Treeview(self, columns=self.col_ids,
                                 show="headings", selectmode=selectmode)
        for cid, cname in zip(self.col_ids, self.col_names):
            cmd = (lambda c=cid: self._sort_by(c)) if sortable else ""
            self.tree.heading(cid, text=cname, anchor=anchor, command=cmd)
            self.tree.column(cid, width=160, minwidth=50,
                             stretch=True, anchor=anchor)

        vsb = ttk.Scrollbar(self, orient="vertical",   command=self.tree.yview)
        hsb = ttk.Scrollbar(self, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        self.tree.bind("<MouseWheel>", lambda e: self.tree.yview_scroll(
            -1 if e.delta > 0 else 1, "units"))
        self.tree.bind("<<TreeviewSelect>>", lambda e: self._on_sel_change())

        # Franjas alternas + sus variantes "_sel". set_tag() añade estados.
        self.tree.tag_configure("row_a",     background=stripe_a)
        self.tree.tag_configure("row_b",     background=stripe_b)
        self.tree.tag_configure("row_a_sel", background=stripe_a_sel, foreground=sel_fg)
        self.tree.tag_configure("row_b_sel", background=stripe_b_sel, foreground=sel_fg)

    # ── Tags de estado (tablas no a franjas) ─────────────────────────────────
    def set_tag(self, name: str, background: str,
                selected_background: str, foreground: str = SEL_FG) -> None:
        """Registra un tag de fila por estado (p. ej. 'ok'/'warn'/'error') junto
        con su variante de selección. Úsalo con load(rows, tags=[...])."""
        self.tree.tag_configure(name, background=background)
        self.tree.tag_configure(f"{name}_sel",
                                background=selected_background, foreground=foreground)

    # ── Carga de datos ───────────────────────────────────────────────────────
    def load(self, rows, tags=None) -> list:
        """Vacía y rellena la tabla. rows: iterable de tuplas (alineadas con
        col_ids). tags: lista paralela con el tag base de cada fila; si es None
        y stripe=True, se asignan franjas alternas por posición. Los tags que
        no son de franja (row_a/row_b) se consideran "sticky": sobreviven a la
        ordenación por cabecera (ver _restripe). Devuelve la lista de iids
        insertados, en orden de carga."""
        self.tree.delete(*self.tree.get_children())
        self._sticky.clear()
        iids = []
        for i, values in enumerate(rows):
            if tags is not None:
                tag = tags[i]
            elif self._stripe:
                tag = "row_a" if i % 2 == 0 else "row_b"
            else:
                tag = ""
            iid = self.tree.insert("", "end", values=values,
                                   tags=(tag,) if tag else ())
            if tag and tag not in ("row_a", "row_b"):
                self._sticky[iid] = tag
            iids.append(iid)
        self._on_sel_change()
        self.after(120, self.autosize)
        return iids

    def scroll_to(self, iid: str, center: bool = True) -> None:
        """Desplaza la vista para dejar visible `iid`; con center=True lo sitúa
        aproximadamente en el centro vertical de la tabla."""
        self.tree.see(iid)
        if not center:
            return
        hijos = self.tree.get_children()
        total = len(hijos)
        if total == 0 or iid not in hijos:
            return
        self.update_idletasks()
        try:
            visibles = max(1, self.tree.winfo_height() //
                           max(1, self.tree.bbox(hijos[0])[3] if self.tree.bbox(hijos[0]) else 20))
        except Exception:
            visibles = 10
        pos = hijos.index(iid)
        top = max(0, pos - visibles // 2)
        self.tree.yview_moveto(top / total)

    def sort(self, col_id: str, reverse: bool = False) -> None:
        """Ordena por una columna de forma determinista (no alterna sentido
        como el clic en la cabecera). Reasigna franjas y conserva sticky."""
        # _sort_by ordena con el sentido actual y luego lo togglea; dejamos el
        # sentido actual = reverse para que ordene como se pide.
        self._sort_rev[col_id] = reverse
        self._sort_by(col_id)

    def set_values(self, iid: str, values) -> None:
        """Sustituye los valores de una fila ya existente (sin recrearla)."""
        self.tree.item(iid, values=values)

    def unmark(self, iid: str) -> None:
        """Quita el tag sticky de una fila: vuelve a ser una franja normal en
        el siguiente restripe (p. ej. tras aprobar una fila de revisión)."""
        self._sticky.pop(iid, None)

    def set_anchor(self, col_id: str, anchor: str) -> None:
        """Alinea una columna concreta distinto del resto de la tabla (p. ej.
        'e' para columnas de importe). Afecta a cabecera y a los valores."""
        self.tree.heading(col_id, anchor=anchor)
        self.tree.column(col_id, anchor=anchor)

    def set_fixed_width(self, col_id: str, width: int) -> None:
        """Fija un ancho MÍNIMO para una columna (fechas, tamaños…): si la
        cabecera o el contenido medido necesitan más, prevalece la medida
        real (nunca trunca). Aplica en la siguiente autosize()."""
        self._fixed[col_id] = width

    # ── Selección ────────────────────────────────────────────────────────────
    def select_all(self) -> None:
        self.tree.selection_set(self.tree.get_children())
        self._on_sel_change()

    def select_none(self) -> None:
        self.tree.selection_remove(self.tree.get_children())
        self._on_sel_change()

    def selection(self) -> list:
        """Valores (tuplas) de las filas seleccionadas, en orden visual."""
        return [self.tree.item(iid, "values") for iid in self.tree.selection()]

    def count(self) -> tuple[int, int]:
        return len(self.tree.selection()), len(self.tree.get_children())

    def _on_sel_change(self) -> None:
        """Repinta cada fila con su tag base o su variante '<tag>_sel' según
        esté seleccionada. Vale para franjas o para estados; si falta la
        variante '_sel' de un tag, esa fila simplemente no cambia de color."""
        selection = set(self.tree.selection())
        for iid in self.tree.get_children():
            current = self.tree.item(iid, "tags")
            base = current[0] if current else ""
            if base.endswith("_sel"):
                base = base[:-4]
            tag = f"{base}_sel" if (iid in selection and base) else base
            self.tree.item(iid, tags=(tag,) if tag else ())
        if self._on_select:
            self._on_select(len(selection), len(self.tree.get_children()))

    # ── Autoajuste de columnas ───────────────────────────────────────────────
    def autosize(self) -> None:
        """Ajusta cada columna al ancho real de su contenido (cabecera en
        negrita vs. filas), sin repartir hueco sobrante como aire extra."""
        self.update_idletasks()
        font_n = tkfont.Font(family=self._font_ui[0],   size=self._font_ui[1])
        font_b = tkfont.Font(family=self._font_bold[0], size=self._font_bold[1],
                             weight="bold")
        PAD = self._pad

        for cid, cname in zip(self.col_ids, self.col_names):
            width = font_b.measure(cname)
            for iid in self.tree.get_children():
                w = font_n.measure(str(self.tree.set(iid, cid)))
                if w > width:
                    width = w
            width += PAD
            if cid in self._fixed:
                width = max(width, self._fixed[cid])
            self.tree.column(cid, width=int(width),
                             minwidth=min(width, 80), stretch=False)

    def set_measure_fonts(self, font_ui, font_bold) -> None:
        """Actualiza las fuentes que autosize() usa para MEDIR el ancho de las
        columnas. El tamaño RENDERIZADO de la tabla lo fija el estilo ttk
        'Treeview' a nivel de app; esto solo mantiene la medición en sintonía
        con ese tamaño para que el autoajuste de columnas siga cuadrando."""
        self._font_ui = font_ui
        self._font_bold = font_bold

    def fit_font_size(self, smin: int, smax: int, avail_px: int) -> int:
        """Mayor tamaño de fuente del rango [smin, smax] con el que TODAS las
        columnas (medidas a ese tamaño) caben en `avail_px` píxeles de ancho,
        es decir, la tabla entra entera sin scroll horizontal. Si ni siquiera
        `smin` cabe, devuelve `smin`. Solo calcula: NO aplica nada.

        Optimización: el argmax de ancho por columna (la celda más ancha) es
        estable al cambiar de tamaño, así que se localiza una sola vez —con una
        única pasada por las filas y midiendo solo las cadenas más largas por
        nº de caracteres— y luego el barrido de tamaños mide únicamente esa
        celda representativa por columna (cols × tamaños medidas, no filas)."""
        smin, smax = int(smin), int(smax)
        if smax < smin:
            smin, smax = smax, smin
        children = self.tree.get_children()
        if not children:
            return smax
        self.update_idletasks()
        fam   = self._font_ui[0]
        fam_b = self._font_bold[0]
        uniq = {cid: set() for cid in self.col_ids}
        for iid in children:
            for cid, v in zip(self.col_ids, self.tree.item(iid, "values")):
                uniq[cid].add(str(v))
        ref = tkfont.Font(family=fam, size=smax)
        widest: dict[str, str] = {}
        for cid in self.col_ids:
            bs, bw = "", 0
            for v in sorted(uniq[cid], key=len, reverse=True)[:8]:
                w = ref.measure(v)
                if w > bw:
                    bw, bs = w, v
            widest[cid] = bs
        for size in range(smax, smin - 1, -1):
            fn = tkfont.Font(family=fam,   size=size)
            fb = tkfont.Font(family=fam_b, size=size, weight="bold")
            total = 0
            for cid, cname in zip(self.col_ids, self.col_names):
                total += max(fb.measure(cname), fn.measure(widest[cid])) + self._pad
            if total <= avail_px:
                return size
        return smin

    # ── Ordenación por cabecera ──────────────────────────────────────────────
    def _sort_by(self, col: str) -> None:
        rows = [(self.tree.set(iid, col), iid) for iid in self.tree.get_children()]
        rev = self._sort_rev[col]
        rows.sort(key=lambda t: self._sort_key(t[0]), reverse=rev)
        for pos, (_, iid) in enumerate(rows):
            self.tree.move(iid, "", pos)
        self._sort_rev[col] = not rev
        if self._stripe:
            self._restripe()
        self._on_sel_change()

    def _restripe(self) -> None:
        """Reasigna row_a/row_b según el orden visual actual (tras ordenar),
        conservando los tags sticky (p. ej. filas marcadas para revisión), que
        no deben perder su color al reordenar."""
        for pos, iid in enumerate(self.tree.get_children()):
            sticky = self._sticky.get(iid)
            tag = sticky if sticky else ("row_a" if pos % 2 == 0 else "row_b")
            self.tree.item(iid, tags=(tag,))

    @staticmethod
    def _sort_key(v):
        """Clave de ordenación: numérica si el valor parsea a float, si no
        alfabética en minúsculas. Los numéricos van antes que los textos."""
        s = str(v).strip()
        try:
            return (0, float(s.replace(",", ".")))
        except ValueError:
            return (1, s.lower())
