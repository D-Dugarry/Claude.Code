#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Convierte ficheros .md en .html autocontenidos para leer en el navegador.

    python md2html.py docs/Informe.md          # un fichero
    python md2html.py docs/A.md docs/B.md      # varios
    python md2html.py docs                     # todos los .md de una carpeta

Cada .html se escribe junto a su .md (mismo nombre, extension .html) y es un
ARTEFACTO GENERADO: no lo edites a mano, edita el .md y vuelve a ejecutar esto.
Sin dependencias externas (nada de pip install) a proposito: debe funcionar en
cualquier maquina con Python, sin instalar nada. Soporta el subconjunto de
Markdown habitual en informes tecnicos: encabezados, tablas, listas, citas,
reglas, codigo (inline y bloque), negrita/cursiva/tachado y enlaces.
"""
# Last Rev. 2026-08-11 13:05

import html
import re
import sys
from pathlib import Path

CSS = """
:root{
  --bg:#f7f7f5; --panel:#fff; --ink:#23211e; --soft:#6b6660; --line:#e2ded7;
  --accent:#a8563f; --accent-soft:#f2e7e2; --code-bg:#f0eeea; --quote:#fbf7ee;
  --quote-line:#d8a54a; --shadow:0 1px 3px rgba(0,0,0,.06),0 8px 24px rgba(0,0,0,.05);
}
@media (prefers-color-scheme:dark){
  :root{
    --bg:#16151a; --panel:#1e1d23; --ink:#e6e2dc; --soft:#9d968c; --line:#33313a;
    --accent:#e29070; --accent-soft:#3a2a26; --code-bg:#26252c; --quote:#2a271f;
    --quote-line:#b8873a; --shadow:0 1px 3px rgba(0,0,0,.3),0 8px 24px rgba(0,0,0,.25);
  }
}
*{box-sizing:border-box}
body{margin:0;background:var(--bg);color:var(--ink);
  font:16px/1.65 -apple-system,"Segoe UI",Roboto,Helvetica,Arial,sans-serif;
  -webkit-font-smoothing:antialiased}
main{max-width:60rem;margin:0 auto;padding:3rem 1.25rem 6rem}
article{background:var(--panel);border:1px solid var(--line);border-radius:14px;
  box-shadow:var(--shadow);padding:2.5rem clamp(1.1rem,4vw,3rem)}
h1{font-size:clamp(1.6rem,4vw,2.15rem);line-height:1.2;margin:0 0 1.2rem;
  letter-spacing:-.02em}
h2{font-size:1.4rem;margin:2.8rem 0 1rem;padding-bottom:.4rem;
  border-bottom:2px solid var(--accent-soft);letter-spacing:-.01em}
h3{font-size:1.12rem;margin:2rem 0 .7rem;color:var(--accent)}
h4{font-size:1rem;margin:1.5rem 0 .5rem;color:var(--soft);
  text-transform:uppercase;letter-spacing:.06em}
p,li{margin:.7rem 0}
ul,ol{padding-left:1.4rem}
li>ul,li>ol{margin:.3rem 0}
a{color:var(--accent);text-decoration:none;border-bottom:1px solid transparent}
a:hover{border-bottom-color:var(--accent)}
strong{font-weight:650}
del{color:var(--soft)}
hr{border:0;border-top:1px solid var(--line);margin:2.5rem 0}
code{background:var(--code-bg);padding:.12em .38em;border-radius:5px;
  font:.875em/1.5 "Cascadia Code",Consolas,"SF Mono",monospace;
  word-break:break-word}
pre{background:var(--code-bg);border:1px solid var(--line);border-radius:10px;
  padding:1rem 1.1rem;overflow-x:auto;margin:1.2rem 0}
pre code{background:none;padding:0;font-size:.85em;line-height:1.55}
blockquote{margin:1.4rem 0;padding:.9rem 1.2rem;background:var(--quote);
  border-left:4px solid var(--quote-line);border-radius:0 8px 8px 0}
blockquote p:first-child{margin-top:0}
blockquote p:last-child{margin-bottom:0}
.tw{overflow-x:auto;margin:1.4rem 0;border:1px solid var(--line);border-radius:10px}
table{border-collapse:collapse;width:100%;font-size:.925rem}
th,td{text-align:left;padding:.6rem .85rem;border-bottom:1px solid var(--line);
  vertical-align:top}
th{background:var(--accent-soft);font-weight:650;white-space:nowrap}
tbody tr:last-child td{border-bottom:0}
tbody tr:nth-child(even){background:color-mix(in srgb,var(--line) 22%,transparent)}
footer{max-width:60rem;margin:0 auto;padding:0 1.25rem 3rem;
  color:var(--soft);font-size:.82rem;text-align:center}
"""

PAGE = """<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>{title}</title>
<style>{css}</style>
</head>
<body>
<main><article>
{body}
</article></main>
<footer>Generado desde <code>{src}</code> con <code>md2html.py</code> &middot;
No editar este HTML a mano.</footer>
</body>
</html>
"""


def inline(text: str) -> str:
    """Marcas de nivel de linea. El codigo inline se aparta antes de escapar."""
    trozos: list[str] = []

    def guardar(m):
        trozos.append(m.group(1))
        return f"\x00{len(trozos) - 1}\x00"

    text = re.sub(r"`([^`]+)`", guardar, text)
    text = html.escape(text, quote=False)
    text = re.sub(r"\[([^\]]+)\]\(([^)]+)\)", r'<a href="\2">\1</a>', text)
    text = re.sub(r"\*\*([^*]+)\*\*", r"<strong>\1</strong>", text)
    text = re.sub(r"~~([^~]+)~~", r"<del>\1</del>", text)
    text = re.sub(r"(?<![\w*])\*([^*\n]+)\*(?![\w*])", r"<em>\1</em>", text)
    return re.sub(
        r"\x00(\d+)\x00",
        lambda m: "<code>" + html.escape(trozos[int(m.group(1))], quote=False) + "</code>",
        text,
    )


def fila_tabla(linea: str) -> list[str]:
    return [c.strip() for c in linea.strip().strip("|").split("|")]


def convertir(md: str) -> str:
    lineas = md.replace("\r\n", "\n").split("\n")
    out: list[str] = []
    i, n = 0, len(lineas)
    pila: list[str] = []          # listas abiertas: "ul" / "ol"

    def cerrar_listas(hasta: int = 0):
        while len(pila) > hasta:
            out.append(f"</{pila.pop()}>")

    while i < n:
        ln = lineas[i]
        crudo = ln.rstrip()
        txt = crudo.strip()

        if not txt:                                        # linea en blanco
            cerrar_listas()
            i += 1
            continue

        if txt.startswith("```"):                          # bloque de codigo
            cerrar_listas()
            i += 1
            buf = []
            while i < n and not lineas[i].strip().startswith("```"):
                buf.append(lineas[i])
                i += 1
            i += 1
            out.append("<pre><code>" + html.escape("\n".join(buf), quote=False) + "</code></pre>")
            continue

        if re.fullmatch(r"(-{3,}|\*{3,}|_{3,})", txt):     # regla horizontal
            cerrar_listas()
            out.append("<hr>")
            i += 1
            continue

        m = re.match(r"(#{1,6})\s+(.*)", txt)              # encabezado
        if m:
            cerrar_listas()
            niv = len(m.group(1))
            out.append(f"<h{niv}>{inline(m.group(2))}</h{niv}>")
            i += 1
            continue

        if txt.startswith(">"):                            # cita (parrafos sueltos)
            cerrar_listas()
            buf = []
            while i < n and lineas[i].strip().startswith(">"):
                buf.append(re.sub(r"^\s*>\s?", "", lineas[i]))
                i += 1
            parr = [p.strip() for p in "\n".join(buf).split("\n\n") if p.strip()]
            out.append("<blockquote>" +
                       "".join(f"<p>{inline(' '.join(p.split()))}</p>" for p in parr) +
                       "</blockquote>")
            continue

        if txt.startswith("|") and i + 1 < n and re.match(       # tabla
                r"^\s*\|[\s:|-]+\|\s*$", lineas[i + 1]):
            cerrar_listas()
            cab = fila_tabla(txt)
            i += 2
            cuerpo = []
            while i < n and lineas[i].strip().startswith("|"):
                cuerpo.append(fila_tabla(lineas[i].strip()))
                i += 1
            th = "".join(f"<th>{inline(c)}</th>" for c in cab)
            trs = "".join(
                "<tr>" + "".join(f"<td>{inline(c)}</td>" for c in fila) + "</tr>"
                for fila in cuerpo)
            out.append(f'<div class="tw"><table><thead><tr>{th}</tr></thead>'
                       f"<tbody>{trs}</tbody></table></div>")
            continue

        m = re.match(r"^(\s*)([-*+]|\d+[.)])\s+(.*)", crudo)     # elemento de lista
        if m:
            sangria, marca, cuerpo = m.group(1), m.group(2), m.group(3)
            tipo = "ul" if marca in "-*+" else "ol"
            nivel = len(sangria) // 2 + 1
            while len(pila) > nivel:
                out.append(f"</{pila.pop()}>")
            if len(pila) < nivel:
                out.append(f"<{tipo}>")
                pila.append(tipo)
            elif pila and pila[-1] != tipo:
                out.append(f"</{pila.pop()}>")
                out.append(f"<{tipo}>")
                pila.append(tipo)
            # continuaciones sangradas del mismo elemento
            i += 1
            while (i < n and lineas[i].strip()
                   and not re.match(r"^\s*([-*+]|\d+[.)])\s+", lineas[i])
                   and not lineas[i].strip().startswith(("#", "|", ">", "```"))
                   and lineas[i].startswith(" ")):
                cuerpo += " " + lineas[i].strip()
                i += 1
            out.append(f"<li>{inline(cuerpo)}</li>")
            continue

        buf = [txt]                                        # parrafo
        i += 1
        while (i < n and lineas[i].strip()
               and not re.match(r"^\s*([-*+]|\d+[.)])\s+|^\s*[#>|]|^\s*```", lineas[i])):
            buf.append(lineas[i].strip())
            i += 1
        cerrar_listas()
        out.append(f"<p>{inline(' '.join(buf))}</p>")

    cerrar_listas()
    return "\n".join(out)


def convertir_fichero(src: Path) -> None:
    md = src.read_text(encoding="utf-8")
    m = re.search(r"^#\s+(.*)$", md, re.M)
    titulo = m.group(1).strip() if m else src.stem.replace("_", " ")
    titulo = re.sub(r"[*`]", "", titulo)

    dst = src.with_suffix(".html")
    dst.write_text(
        PAGE.format(title=html.escape(titulo), css=CSS, body=convertir(md), src=src.name),
        encoding="utf-8", newline="\n")
    print(f"OK  {src.name} -> {dst.name}  ({dst.stat().st_size:,} bytes)")


def main() -> int:
    if len(sys.argv) < 2:
        print(__doc__)
        return 2

    objetivos: list[Path] = []
    for arg in sys.argv[1:]:
        p = Path(arg)
        if p.is_dir():
            encontrados = sorted(p.glob("*.md"))
            if not encontrados:
                print(f"AVISO: sin ficheros .md en {p}")
            objetivos.extend(encontrados)
        elif p.is_file():
            objetivos.append(p)
        else:
            print(f"ERROR: no existe {p}")
            return 1

    if not objetivos:
        print("ERROR: nada que convertir")
        return 1

    for src in objetivos:
        convertir_fichero(src)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
