"""
font_sampler.py
Muestra 15 fuentes manuscritas/caligráficas disponibles en el sistema.
Dile a Claude el número de la que prefieras.
"""
import tkinter as tk
from tkinter import font as tkfont

CANDIDATAS = [
    "Segoe Script", "Segoe Print", "Lucida Handwriting",
    "Brush Script MT", "Freestyle Script", "Mistral",
    "French Script MT", "Monotype Corsiva", "Vladimir Script",
    "Edwardian Script ITC", "Kunstler Script", "Palace Script MT",
    "Vivaldi", "Gabriola", "Harlow Solid Italic",
    "Rage Italic", "Gigi", "Comic Sans MS",
    "Script MT Bold", "Calligraph421 BT",
]

root = tk.Tk()
root.title("Muestra de fuentes — dile a Claude el número")
root.configure(bg="#1A5276")
root.resizable(False, False)

familias = set(tkfont.families())
disponibles = [f for f in CANDIDATAS if f in familias][:15]

if not disponibles:
    tk.Label(root, text="No se encontraron fuentes manuscritas instaladas.",
             bg="#1A5276", fg="white", font=("Verdana", 12), padx=20, pady=20).pack()
else:
    tk.Label(root, text="  Elige un número y díselo a Claude  ",
             bg="#1A5276", fg="#FFFDE7", font=("Verdana", 11, "italic"),
             pady=10).pack(fill="x")

    for i, nombre in enumerate(disponibles, 1):
        fila = tk.Frame(root, bg="#1A5276")
        fila.pack(fill="x", padx=20, pady=4)

        tk.Label(fila, text=f"{i:2}.", bg="#1A5276", fg="#A8C8E0",
                 font=("Courier New", 13, "bold"), width=4, anchor="e").pack(side="left")

        tk.Label(fila, text="Dugarry", bg="#1A5276", fg="#C8A96E",
                 font=(nombre, 22), anchor="w").pack(side="left", padx=(8, 0))

        tk.Label(fila, text=f"  [{nombre}]", bg="#1A5276", fg="#5A7A9A",
                 font=("Verdana", 8), anchor="w").pack(side="left")

    tk.Frame(root, bg="#1A5276", height=12).pack()

sw = root.winfo_screenwidth()
sh = root.winfo_screenheight()
root.update_idletasks()
w = root.winfo_reqwidth()
h = root.winfo_reqheight()
root.geometry(f"{w}x{h}+{(sw-w)//2}+{(sh-h)//2}")

root.mainloop()
