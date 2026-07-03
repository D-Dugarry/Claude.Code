# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Qué hace este proyecto

App de escritorio (tkinter, Windows) que convierte **un** PDF "Listado de Liquidación de Tasas Académicas de Matrícula" (Universidad Permanente UPUA, Universidad de Alicante) a un fichero Excel con dos hojas:

- **Resumen** — una fila por alumno/expediente, con importes agregados.
- **Detalle** — una fila por línea de cobro (Referencia/Fecha/Importe/Plazo/Forma Pago) **más una fila adicional por alumno** con el importe administrativo en la columna `Imp.Adm.` (esa fila deja vacías las columnas de cobro, porque el cargo administrativo no es una línea de cobro). Así cada importe — cobros y administrativo — tiene su propia fila para conciliar uno a uno.

La app procesa un único PDF por ejecución (no hay cola multi-fichero).

## Comandos

```bash
pip install -r requirements.txt
python Convert_PDF_Liquid_EP_to_Excel.py      # lanzar la app
python -m py_compile *.py                      # comprobación rápida de sintaxis
```

No hay suite de tests automatizada. La lógica de parseo (`pdf_parser.py`) se validó manualmente contra los dos PDF de muestra de `PDF_Liquidaciones/` comparando la suma de importes netos calculados con la línea `IMPORTE TOTAL POR TASAS` que trae el propio PDF (coincide exactamente: 205.649,00 €, 1645 registros, 206 páginas, 0 líneas sin clasificar). Al tocar `pdf_parser.py`, repetir esa comprobación:

```python
from pdf_parser import parse_pdf
records, paginas = parse_pdf("PDF_Liquidaciones/_237_1365588_Tasas202526.pdf")
print(len(records), paginas, round(sum(r.importe_neto for r in records), 2))
# Esperado: 1645 206 205649.0
```

Y para la hoja Detalle (fila extra de Imp.Adm.), el número de filas debe ser exactamente `nº_referencias_totales + nº_registros + 1` (cabecera), y la suma de la columna `Imp.Adm.` debe coincidir con `sum(r.administrativo for r in records)`.

## Arquitectura

Tres módulos independientes (sin dependencias circulares) más la UI:

- **`pdf_parser.py`** — extrae los datos del PDF con `fitz.get_text("text", sort=True)` (PyMuPDF) y regex sobre las líneas resultantes. `Registro` (un alumno) contiene una lista de `Referencia` (una por línea de cobro) más `importe`, `administrativo` e `importe_total_pdf`. **`importe_neto` se calcula siempre como `importe - administrativo`**, nunca se confía en el campo "Importe Total" tal cual lo trae el PDF, porque ese campo viene en blanco en ~0,6 % de los registros (defecto real del informe origen, no de la extracción — verificado con `fitz`; cuando sí trae valor, coincide al 100 % con el cálculo). Solo expone `parse_pdf(path)` (un fichero); no hay variante multi-fichero.
- **`excel_export.py`** — vuelca una `list[Registro]` a `.xlsx` (hojas Resumen + Detalle) con `openpyxl`. No conoce nada de PDF ni de tkinter. La hoja Detalle escribe, por cada alumno, una fila por `Referencia` y luego una fila extra solo con `Exped/DNI/Nombre/Imp.Adm.` (columnas de cobro vacías).
- **`Convert_PDF_Liquid_EP_to_Excel.py`** — UI tkinter (layout de 3 bloques + panel de configuración, plantilla del skill `tkinter-app-design`; `data_table.py` y `help_tooltips.py` están vendorizados junto a este archivo, misma fuente que en Corrector_Orientacion_PDF y Mailing_Personalizado). Bloque "Selección" = un único PDF + carpeta destino (sin cola). Tabla 1 = "Vista previa (Detalle)" (espejo exacto de la hoja Detalle, incluida la fila de Imp.Adm.); Tabla 2 = "Vista previa (Resumen)". El botón "Ejecutar" parsea + exporta en un solo paso, en un hilo aparte (`threading`) para no bloquear la UI; las actualizaciones de widgets desde el hilo se despachan con `self.after(...)`. El bloque "Informe" tiene el log reducido a 4 líneas de alto (antes 6) para dejar más sitio a las tablas.

## Formato del PDF de entrada (importante si se toca `pdf_parser.py`)

Informe paginado, cabecera/pie repetidos en cada página, con la fila de columnas `Exped | Dni | Apellidos y Nombre | Referencia | Fecha Cobro | Importe Cobrado | Plazo | Forma Pago`. Por cada alumno:

```
<Exped>  <Dni>  <Apellidos y Nombre>
<Referencia 13 dígitos>  <dd/mm/aaaa>  <Importe>  <Plazo>  <Forma Pago>   (1..n líneas)
Importe:          <suma de los importes cobrados>
Administrativo:   <gastos administrativos>
Importe Total :   <importe neto> €        (puede venir en blanco)
```

Al final del documento: `IMPORTE TOTAL POR TASAS : <importe> €` (gran total).

### Peculiaridades reales del PDF (verificadas con PyMuPDF, no son artefactos de extracción)

- **Usar `get_text("text", sort=True)`, no el orden de lectura por defecto.** Sin `sort=True`, la cabecera de columnas sale mezclada (los fragmentos de texto se leen en el orden en que se dibujaron en el PDF, no en orden visual). Con `sort=True` el texto sale ya alineado por líneas visuales y es directamente parseable por regex — no hizo falta extracción por coordenadas (`get_text("words")`) pese a que ese fue el plan inicial.
- **Discriminador Exped vs Referencia**: la Referencia de cobro son siempre 13 dígitos (empieza por el año, p. ej. `2025985284015`); el Exped nunca llega a esa longitud (1-4 dígitos en los PDF de muestra). Es la clave para distinguir la línea de cabecera de alumno de una línea de cobro sin ambigüedad.
- **"Importe Total" en blanco**: en ~0,6 % de los registros el PDF no imprime el valor de esa celda (el resto de campos sí están bien). No se ha encontrado un patrón que lo explique (no depende de tener 1 o varias líneas de referencia). Por eso `Registro.importe_neto` se calcula siempre, no se lee directamente.
- **Ñ/tildes corruptas en el propio PDF** (no es un problema de extracción): algunos nombres traen glyphs sin mapeo Unicode válido en la fuente embebida del PDF origen. Ejemplos reales extraídos con `fitz`: `"IVA#EZ OLCINA, MONICA"` (Ñ → `#`) y `"CARDONA VALENCIA, BEGO<U+FFFD>A"` (Ñ → carácter de reemplazo). El conversor no intenta "arreglar" estos nombres — se exportan tal cual vienen, porque no hay forma fiable de saber qué carácter faltaba.
- **DNI heterogéneo**: NIF, NIE (`X`/`Y`/`Z` + dígitos + letra), pasaporte extranjero con letras, y algunos con el mismo tipo de corrupción de fuente que los nombres. Se trata siempre como texto libre, sin validar formato.
- El segundo PDF de muestra (`_237_1450473_ltimastasas202526.pdf`, 228 páginas, 1816 registros) sigue el mismo layout y lo parsea el mismo código sin cambios.

## Contexto del repositorio Git

- Esta carpeta **no es un repo Git propio**: forma parte del monorepo `___Claude.Code` (la raíz Git real está un nivel por encima, en `F:\__Dugarry UA\Dugarry Proyectos\___Claude.Code`).
- Todavía no aparece en el `.gitignore` de proyectos migrados a repos independientes (a diferencia de Norma43, Mailing_Personalizado, etc.), por lo que de momento se versiona dentro del monorepo.
- Si en el futuro se decide migrarlo a un repo independiente, seguir el runbook `___Claude.Code\MIGRAR_A_REPO.md`.

## Pendiente / decisiones abiertas

- No se ha probado el flujo completo de la UI de punta a punta con clics reales en el diálogo de archivos (sí se comprobó por script que `_load_tabla_detalle`/`_load_tabla_resumen` rellenan las tablas correctamente con datos reales, y que la app arranca y renderiza bien — capturas de pantalla vía `PrintWindow`, no `CopyFromScreen`, porque esta última mezcló contenido de otra ventana por un desajuste de coordenadas/DPI en un escritorio con varios monitores). Conviene hacer una pasada manual con clics reales antes de dar la app por terminada.
- No hay `.spec` de PyInstaller todavía (los otros proyectos del ecosistema sí lo tienen, p. ej. `Corrector_Orientacion_PDF.spec`) — añadirlo si se necesita distribuir como `.exe`.
