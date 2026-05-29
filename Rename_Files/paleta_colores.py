"""
paleta_colores.py
Muestra combinaciones de colores para el Treeview de Rename_Files.
Cierra la ventana cuando hayas elegido y dile a Claude el número de paleta.
"""
import tkinter as tk
from tkinter import ttk

PALETAS = [
    {
        "nombre": "1 · Azul clásico suave",
        "even":     "#FFFFFF",
        "odd":      "#E8EDF2",
        "even_sel": "#5B9BD5",
        "odd_sel":  "#2E75B6",
    },
    {
        "nombre": "2 · Azul pastel",
        "even":     "#FFFFFF",
        "odd":      "#EBF5FB",
        "even_sel": "#85C1E9",
        "odd_sel":  "#5499C7",
    },
    {
        "nombre": "3 · Azul acero",
        "even":     "#FFFFFF",
        "odd":      "#ECF4F9",
        "even_sel": "#6BAED6",
        "odd_sel":  "#3182BD",
    },
    {
        "nombre": "4 · Verde-azul (teal)",
        "even":     "#FFFFFF",
        "odd":      "#E8F6F3",
        "even_sel": "#76D7C4",
        "odd_sel":  "#1ABC9C",
    },
    {
        "nombre": "5 · Gris-azul profesional",
        "even":     "#FFFFFF",
        "odd":      "#EAECEE",
        "even_sel": "#7FB3D3",
        "odd_sel":  "#5D8AA8",
    },
    {
        "nombre": "6 · Índigo suave",
        "even":     "#FFFFFF",
        "odd":      "#EEF0F8",
        "even_sel": "#9B8EC4",
        "odd_sel":  "#6A5ACD",
    },
    {
        "nombre": "7 · Naranja cálido",
        "even":     "#FFFFFF",
        "odd":      "#FDF5E6",
        "even_sel": "#F0A500",
        "odd_sel":  "#D4880A",
    },
    {
        "nombre": "8 · Originales (actuales)",
        "even":     "#F4F6F7",
        "odd":      "#FFFFFF",
        "even_sel": "#2E86C1",
        "odd_sel":  "#1A5276",
    },
]

FILAS = [
    ("dab@ua.es",   "Indicadores DP Análisis Geog. 2027 ok.xlsx",   "(dab@) DP Análisis Geog. 2027 ok.xlsx",   "-.-"),
    ("daea@ua.es",  "Indicadores DP Bioquímica 2027 ok.xlsx",        "(daea@) DP Bioquímica 2027 ok.xlsx",        "-.-"),
    ("dagr@ua.es",  "Indicadores DP Biotecnología 2027 ok.xlsx",     "(dagr@) DP Biotecnología 2027 ok.xlsx",     "-.-"),
    ("dbt@ua.es",   "Indicadores DP Ciencias Ambientales 2027.xlsx", "(dbt@) DP Ciencias Ambientales 2027.xlsx",  "-.-"),
    ("dcarn@ua.es", "Indicadores DP Ciencias Hist.-Jur. 2027.xlsx",  "(dcarn@) DP Ciencias Hist.-Jur. 2027.xlsx", "-.-"),
    ("dctma@ua.es", "Indicadores DP Ciencias de la Tierra 2027.xlsx","(dctma@) DP Ciencias de la Tierra 2027.xlsx","-.-"),
]
SEL_IIDS = {"0", "1", "2"}   # primeras 3 filas "seleccionadas"

BASE = 11
FONT_UI   = ("Verdana", BASE)
FONT_BOLD = ("Verdana", BASE, "bold")
FONT_H    = ("Verdana", BASE + 1, "bold")

root = tk.Tk()
root.title("Paleta de colores — elige un número y díselo a Claude")
root.configure(bg="#F2F3F4")
sw = root.winfo_screenwidth()
sh = root.winfo_screenheight()
root.geometry(f"{min(sw, 1400)}x{min(sh-60, 900)}+0+0")

# ── Forzar fondo de selección vacío en el tema clam ─────────────────────────
s = ttk.Style(root)
s.theme_use("clam")
root.tk.eval("ttk::style theme settings clam { ttk::style map Treeview -background {} }")
s.configure("Treeview", rowheight=22, font=FONT_UI)
s.configure("Treeview.Heading", font=FONT_BOLD, background="#DDDDDD", relief="flat")
s.map("Treeview", foreground=[("selected", "white")])

titulo = tk.Label(root,
    text="Filas SIN seleccionar: 4ª, 5ª, 6ª    |    Filas SELECCIONADAS: 1ª, 2ª, 3ª",
    bg="#F2F3F4", font=("Verdana", 12, "italic"), pady=8)
titulo.pack()

canvas = tk.Canvas(root, bg="#F2F3F4", highlightthickness=0)
vsb    = ttk.Scrollbar(root, orient="vertical",   command=canvas.yview)
hsb    = ttk.Scrollbar(root, orient="horizontal", command=canvas.xview)
canvas.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
hsb.pack(side="bottom", fill="x")
vsb.pack(side="right",  fill="y")
canvas.pack(fill="both", expand=True, padx=8, pady=4)

frame = tk.Frame(canvas, bg="#F2F3F4")
canvas.create_window((0, 0), window=frame, anchor="nw")
frame.bind("<Configure>", lambda e: canvas.configure(
    scrollregion=canvas.bbox("all")))

cols = ("mail", "fichero", "newname", "resultado")

for p in PALETAS:
    lbl = tk.Label(frame, text=p["nombre"], bg="#F2F3F4",
                   font=FONT_H, anchor="w", pady=4)
    lbl.pack(fill="x", padx=8)

    tree = ttk.Treeview(frame, columns=cols, show="headings",
                        height=len(FILAS), selectmode="none")
    for col, w in zip(cols, (160, 340, 340, 70)):
        tree.heading(col, text=col.capitalize())
        tree.column(col, width=w, minwidth=60)

    tree.tag_configure("even",     background=p["even"],     foreground="#222222")
    tree.tag_configure("odd",      background=p["odd"],      foreground="#222222")
    tree.tag_configure("even_sel", background=p["even_sel"], foreground="white")
    tree.tag_configure("odd_sel",  background=p["odd_sel"],  foreground="white")

    for i, fila in enumerate(FILAS):
        even = i % 2 == 0
        if str(i) in SEL_IIDS:
            tag = "even_sel" if even else "odd_sel"
        else:
            tag = "even" if even else "odd"
        tree.insert("", "end", iid=str(i), values=fila, tags=(tag,))

    # Leyenda de colores
    info = (f"  Sin sel: par {p['even']}  impar {p['odd']}    "
            f"Con sel: par {p['even_sel']}  impar {p['odd_sel']}")
    tk.Label(frame, text=info, bg="#F2F3F4",
             font=("Courier New", 9), fg="#555555", anchor="w").pack(
        fill="x", padx=8)

    tree.pack(fill="x", padx=8, pady=(0, 12))

root.mainloop()
