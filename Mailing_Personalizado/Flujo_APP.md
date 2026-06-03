# Mailing Personalizado — Flujo del Programa

## Visión general

Herramienta de escritorio (Python/tkinter, compilada a `.exe`) para el envío masivo de
correos personalizados con adjuntos a los departamentos de la Universidad de Alicante.
Cada correo se dirige a la dirección extraída del propio fichero Excel adjunto (hoja `Correo`, celda A1),
permite adjuntos comunes para todos los destinatarios, y soporta cuerpo bilingüe ES / VA con firma HTML.

---

## Pantalla principal — secciones

```
┌──────────────────────────────────────────────────────────────────────────┐
│  Carpetas de Adjuntos                                         [Dugarry]  │
│    Adj. Personalizados:  [ruta/carpeta/personalizados]      Seleccionar  │
│    Adj. Comunes:         [ruta/carpeta/comunes]             Seleccionar  │
├──────────────────────────────────────────────────┬───────────────────────┤
│  Ficheros Adj. Personalizados  N/Total Todos Ngno│ Adj. Comunes N/T T N  │
│  ┌────────────┬─────────────┬──────────┬─────────────┬─────────────┐    │
│  │ Correo     │ Fichero     │F. Modif. │ Observación │    Envío    │    │
│  │ dep@ua.es  │ dep.xlsx    │26/01/25  │             │      -      │    │
│  │ (vacío)    │ otro.xlsx   │26/01/25  │Sin Correo   │  No Enviado │    │
│  └────────────┴─────────────┴──────────┴─────────────┴─────────────┘    │
│                                           ┌──────────┬──────────┐        │
│                                           │ Fichero  │F. Modif. │        │
│                                           │ doc.pdf  │01/02/25  │        │
│                                           └──────────┴──────────┘        │
├──────────────────────────────────────────────────────────────────────────┤
│  Informe   [Ejecutar Mailing]  [Prueba Activada]        [Configuración]  │
│  > log de operaciones en fondo negro                          [Limpiar]  │
└──────────────────────────────────────────────────────────────────────────┘
```

---

## Caracolillo (logo Dugarry)

El logotipo de la barra **Carpetas de Adjuntos** es el control de ayuda de la aplicación:

- **Pasar el cursor** → muestra la firma *Dugarry'26* y un texto explicativo con la acción
  del control (*Mostrar / Ocultar Texto explicativo de Botones*, según el estado actual).
- **Clic izquierdo** (`Button-1`) → alterna la visibilidad de **todos** los textos
  explicativos (tooltips) de los botones del proyecto. El del propio caracolillo se ve siempre.
- **Clic derecho** (`Button-3`) → muestra u oculta el **panel Flujo de la Aplicación** (toggle).

## Textos explicativos (tooltips)

Cada botón tiene un tooltip de ayuda gestionado por `_DynTooltip`:

- Un **interruptor maestro** (atributo de clase `_DynTooltip.enabled`) enciende o apaga
  todos a la vez; se alterna con el clic izquierdo en el caracolillo.
- Los tooltips `always=True` (firma y explicación del caracolillo) ignoran el interruptor
  y se muestran siempre.
- El estado **de arranque** lo fija el check *Textos explicativos al iniciar APP* del panel
  de Configuración (clave de registro `CfgTipsInicio`).
- Texto **dinámico** (callable evaluado al mostrarse) y **esquinas redondeadas** (API DWM de
  Windows 11, helper `_round_corners`).

## Panel Flujo de la Aplicación

Overlay que cubre la zona de las dos tablas. Se construye la primera vez (lazy) y se reutiliza.

| Elemento             | Descripción                                                      |
|----------------------|------------------------------------------------------------------|
| Cabecera             | Barra morada (`BG_TABLA`) con título y botones                   |
| ⬇ Descargar en PDF  | Exporta `_FLUJO_MD` a PDF con `reportlab` → `_export_flujo_pdf`  |
| ✕ Cerrar             | Oculta el panel (`place_forget`)                                 |
| Cuerpo               | `ScrolledText` con markdown renderizado por `_render_md`         |

### Exportación a PDF (`_export_flujo_pdf`)

```
Pide ruta → asksaveasfilename
    default: "Flujo_Mailing_Personalizado.pdf"
│
├─ Parsea _FLUJO_MD línea a línea:
│       # / ## / ###  → Paragraph (Helvetica-Bold, colores de la app)
│       | tabla |     → Table (cabecera gris, filas alternas, grid 0.5 pt)
│       ``` bloque `` → Paragraph Courier fondo gris (chars caja → ASCII)
│       - bullet      → Paragraph con •
│       > cita        → Paragraph Helvetica-Oblique, indentado
│       ---           → HRFlowable
│
└─ SimpleDocTemplate(A4, márgenes 2 cm) → doc.build(story)
        → Informe ✔  +  messagebox "PDF generado"
```

---

## Flujo paso a paso

### 1 · Seleccionar Carpeta de Adj. Personalizados

- El usuario pulsa **Seleccionar** o escribe la ruta directamente en el campo.
- La ruta se guarda en el registro de Windows (`HKCU\Software\MailingPersonalizado\LastAdjPersonalizados`).
- Cambiar la ruta dispara `_on_adj_pers_change` → espera 300 ms (debounce) → llama a `_load_adj_pers`.

### 2 · Carga tabla Adj. Personalizados (`_load_adj_pers`)

```
_load_adj_pers()
    │
    ├─ Lee os.listdir(carpeta) → solo ficheros .xlsx (alfabéticamente)
    ├─ Para cada fichero:
    │       fmodif = fecha de modificación "DD/MM/YYYY"
    │       fsize  = tamaño formateado (B / KB / MB / GB)
    │       fila = ("", fichero, fmodif, fsize, "", "")
    │
    ├─ _populate_tree(tree1, …)
    │       ├─ Inserta filas en el Treeview con tags even/odd
    │       ├─ _on_sel_change()  → actualiza contador y colores
    │       └─ after(120ms) → _autosize_columns()
    │
    └─ Si openpyxl disponible y hay filas:
            → hilo _read_emails_thread(carpeta, files, gen)
```

### 3 · Lectura de emails en hilo (`_read_emails_thread`)

Se ejecuta en hilo secundario para no bloquear la UI. Usa un contador `gen`
para ignorar resultados de una carga anterior si la carpeta cambia mientras el hilo corre.

```
Para cada fichero .xlsx:
    │
    ├─ openpyxl.load_workbook(read_only=True)
    ├─ Busca hoja cuyo nombre (lower) == "correo"
    │
    ├─ [Sin hoja]       → obs = "Sin Hoja Correo"
    ├─ [Hoja vacía]     → obs = "Sin Correo"
    ├─ [Email inválido] → obs = "Correo Erróneo"
    └─ [Email válido]   → email = valor de A1, obs = ""
    │
    └─ after(0, _update_email_cell) → actualiza Correo y Observación en hilo principal
```

Al terminar todos los ficheros: autoajusta columnas, ordena la tabla por la columna **Correo** (ascendente) y escribe `"✔ Emails cargados."` en el Informe.

### 4 · Seleccionar Carpeta de Adj. Comunes (opcional)

- Igual mecanismo debounce → `_load_adj_com`.
- Lee **todos** los ficheros de la carpeta (no solo `.xlsx`).
- Tabla 2 solo tiene columnas `Fichero` y `F. Modif.`
- La ruta persiste en `HKCU\Software\MailingPersonalizado\LastAdjComunes`.

### 5 · Ajuste de columnas (`_autosize_columns`)

**Tabla 1 — Adj. Personalizados:**

| Columna       | Anchura                                                          | Alineación | Stretch |
|---------------|------------------------------------------------------------------|------------|---------|
| `F. Modif.`   | Fija: máx. entre título y `"DD/MM/YYYY"` +15 px                 | Centro     | No      |
| `Tamaño`      | Fija: máx. entre título y `"999.9 MB"` +15 px                   | Derecha    | No      |
| `Correo`      | Fija: máx. de todos los valores +15 px                          | Izquierda  | No      |
| `Fichero`     | Fija: máx. de todos los valores +15 px                          | Izquierda  | No      |
| `Observación` | Fija: máx. de los valores posibles (`"Sin Hoja Correo"`) +15 | Centro     | No      |
| `Envío`       | Fija: máx. de los valores posibles (`"Enviado correo 99"`) +15  | Centro     | No      |

**Tabla 2 — Adj. Comunes:**

| Columna     | Anchura                                | Alineación | Stretch |
|-------------|----------------------------------------|------------|---------|
| `F. Modif.` | Fija: igual que tabla 1                | Centro     | No      |
| `Tamaño`    | Fija: igual que tabla 1                | Derecha    | No      |
| `Fichero`   | Fija: máx. de todos los valores +15 px | Izquierda  | No      |

### 6 · Selección de filas

- **Todos** / **Ninguno**: selecciona o deselecciona todas las filas.
- **Clic individual**: selecciona una fila (deselecciona el resto).
- **Shift + clic** / **Ctrl + clic**: selección de rango o múltiple (modo `extended`).
- Cualquier cambio llama a `_on_sel_change`:
  - Actualiza el tag de cada fila (even / odd / even_sel / odd_sel).
  - Actualiza el contador `N / Total` en la cabecera del bloque.

> Cuando **Activado Modif. Correos** está activo, la tabla 1 pasa a `selectmode="none"`:
> ningún clic selecciona filas, Shift/Ctrl no tienen efecto. Al desactivar vuelve a `"extended"`.

### 6b · Activar Modif. Correos (`_toggle_modif_correos`)

Botón en la cabecera de la tabla 1. Arranca siempre **desactivado**.

| Estado | Texto | Estilo | Comportamiento tabla 1 |
|---|---|---|---|
| Desactivado | `Activar Modif. Correos` | `Sel.TButton` (gris) | `selectmode="extended"` — selección normal |
| Activado | `Activado Modif. Correos` | `ModifOn.TButton` (naranja) | `selectmode="none"` + deselecciona todo |

Al hacer **clic en una fila** con el modo activo → `_open_edit_correo(iid)`:

```
Abre overlay (gris) + panel centrado ✎ Modificar Correo
    │
    ├─ Muestra fichero (solo lectura) y correo actual en Entry
    ├─ Valida: convierte a minúsculas, comprueba formato
    │
    ├─ _write_correo_xlsx(filepath, email)
    │       Abre xlsx → busca hoja "Correo" (case-insensitive)
    │       Si no existe → crea hoja "Correo"
    │       Escribe email en A1  →  sheet_state = "hidden"
    │       Devuelve True si la hoja fue creada nueva
    │
    ├─ Actualiza columnas Correo y Observación en el Treeview
    ├─ Reordena tabla por Correo (ascendente)
    ├─ Autoajusta columna Correo
    │
    └─ Flash explicativo (3 s):
            "Correo Borrado"                         si email vacío
            "Correo Añadido en Hoja Creada 'Correo'" si hoja nueva
            "Correo Modificado"                      si había correo previo
            "Correo Añadido"                         si celda estaba vacía
```

### 7 · Panel ⚙ Configuración

Overlay a pantalla completa. Se abre y cierra con el botón **⚙ Configuración**.
Si hay cambios sin guardar al cerrar, solicita confirmación.

**Grupo izquierdo — SMTP y contenido:**

| Campo                       | Clave registro                  |
|-----------------------------|---------------------------------|
| Cuenta Gmail (SMTP)         | `CfgMailCta`                    |
| Remitente visible           | `CfgMailFrom`                   |
| Mail Password (oculto)      | `CfgMailClau`                   |
| Correo para testear         | `CfgDirPrueba`                  |
| Tipo de letra + Tamaño      | `CfgFontFamily` / `CfgFontSize` |
| ☑ Animación (cartero)       | `CfgAnimacion`                  |
| ☑ Textos explicativos al iniciar APP | `CfgTipsInicio`        |
| Asunto del Correo           | `CfgAsunto`                     |

> Los dos checkboxes van a la derecha de la fila *Tipo de letra / Tamaño*. **Animación**
> activa/desactiva el GIF del cartero durante el envío (por defecto activado). **Textos
> explicativos al iniciar APP** fija si los tooltips de los botones arrancan visibles
> (por defecto desactivado; en caliente se alternan con el clic izquierdo en el caracolillo).

**Grupo derecho — Variables de la firma:**

| Campo        | Clave registro |
|--------------|----------------|
| Extensión    | `CfgUserExt`   |
| Web ES       | `CfgWebEs`     |
| Web VA       | `CfgWebVa`     |
| Servicio     | `CfgServicio`  |
| Unidad       | `CfgUnidad`    |

**Cuerpo del correo (bilingüe):**

- Dos columnas: Español (izquierda) y Valencià (derecha).
- Cada columna tiene: **Saludo** (entry) | **Cuerpo** (ScrolledText) | **Despedida** (entry).
- Botones **Traducir** (ES→VA) disponibles en saludo, cuerpo y despedida usando `deep-translator`.
- Toggle **⇄ Español / Valencià**: define cuál idioma aparece primero en el correo. Se guarda en `CfgLangOrder`.

**Firma HTML:**

- Visible solo si el correo de prueba contiene "dugarry" (administrador).
- Se puede expandir/contraer con **▼ Visualizar Firma**.
- Los placeholders de la firma se sustituyen al generar el HTML con los valores de las variables.

**Acciones del panel:**

| Botón               | Acción                                                                     |
|---------------------|----------------------------------------------------------------------------|
| ¡Guardar!           | Persiste todos los campos en el registro de Windows                        |
| Borrar Cambios      | Revierte a la última salvaguardia (aparece solo si hay cambios pendientes) |
| 👁 Previsualizar    | Overlay con el correo montado (texto plano) en orden de idiomas activo     |
| ✉ Mandar 1 Muestra | Flash `"Enviando Mail..."` → envía correo sin adjuntos → flash `"¡ Mail enviado !"` 3 s |

### 8 · Modo Prueba / Producción

| Estado          | Botón             | Comportamiento                                                          |
|-----------------|-------------------|-------------------------------------------------------------------------|
| Prueba Activada | Rojo / PruebaOn   | Todos los correos se redirigen a `dir_prueba`; asunto lleva `[PRUEBA]` |
| Activar Prueba  | Gris / Prueba     | Correos se envían a las direcciones reales                             |

> El modo Prueba está **activado por defecto** al arrancar la aplicación.

### 9 · Ejecutar Mailing (`_ejecutar_mailing`)

**Validaciones previas:**
- Cuenta SMTP configurada.
- App Password configurado.
- Asunto no vacío.
- Al menos una fila seleccionada con email válido.

**Agrupación de destinatarios:**

```
Para cada fila en Tabla 1:
    │
    ├─ [No seleccionada]       → Envío = "no seleccionado"
    ├─ [Observación con error] → Envío = "No Enviado"
    └─ [Email válido]
            ├─ Agrupa por dirección única (groups[email])
            ├─ Asigna número de grupo correlativo
            └─ Añade ruta del fichero xlsx a la lista del grupo
```

**Doble confirmación** (modal con modo, nº destinatarios, nº adjuntos comunes).

**Hilo de trabajo (`_send_thread`)** — se ejecuta en hilo secundario:

```
Conecta SMTP_SSL smtp.gmail.com:465
│
└─ Para cada (email, adj_personalizados) en groups:
        │
        ├─ dest = dir_prueba  (si modo prueba)  |  email (si producción)
        ├─ Construye MIMEMultipart
        │       From / To / Subject
        │       Body HTML (_build_body_html)
        │       Adjuntos personalizados del grupo
        │       Adjuntos comunes seleccionados
        │
        ├─ server.sendmail()
        │       OK    → Envío = "Enviado correo N"  |  log ✔ verde
        └─ Exception → Envío = "Envío Fallido"      |  log ✖ rojo
```

Al terminar: reactiva el botón, muestra flash "¡ Mailing completado !" y el diálogo de resumen.

#### Animación del cartero (overlay `MailSent.gif`)

Mientras se envía, un overlay flotante y **sin fondo** (transparencia por `-transparentcolor`)
muestra el GIF del cartero:

- Por cada **fila de adjunto personalizado** procesada, el cartero **se desliza** suavemente
  (con desaceleración) hacia un punto aleatorio dentro de la ventana — `_move_sending_overlay`
  / `_glide_loop`.
- Sobre el sobre se dibuja el **número del correo** en curso, compuesto en la propia imagen
  con PIL (`_set_overlay_number`).
- Se puede desactivar con el check **Animación** del panel de Configuración (`CfgAnimacion`).

#### Diálogo de resumen (`_show_envio_done`)

Al acabar muestra un diálogo modal con los totales **alineados a la derecha**:

```
Correos enviados:                    6

No enviados:                         3      (naranja oscuro)
        Sin Correo:                  2
        Sin Hoja Correo:             1
[Errores:                            N]     (solo si hubo fallos SMTP)
```

El desglose de **No enviados** por categoría (Sin Correo / Sin Hoja Correo / Correo Erróneo)
se calcula en `_ejecutar_mailing` (`_no_enviado_counts`).

### 10 · Estructura del correo HTML (`_build_body_html`)

El cuerpo se monta siempre con **idioma secundario primero**, idioma principal después:

```
<html><body>
  Saludo    (idioma-1)
  Cuerpo    (idioma-1)
  Despedida (idioma-1)
  <hr>
  Saludo    (idioma-2)
  Cuerpo    (idioma-2)
  Despedida (idioma-2)
  [Firma HTML con placeholders sustituidos]
</body></html>
```

Los placeholders sustituidos en la firma:

| Placeholder                       | Variable   |
|-----------------------------------|------------|
| `ext. ____`                       | `user_ext` |
| `https://sc.ua.es/es/_____`       | `web_es`   |
| `https://sc.ua.es/va/_____`       | `web_va`   |
| `Servei _____`                    | `servicio` |
| `Unitat _____`                    | `unidad`   |

---

## Columna Envío — valores posibles

| Valor              | Condición                                          |
|--------------------|----------------------------------------------------|
| *(vacío)*          | Fila no procesada (estado inicial)                 |
| `no seleccionado`  | Fila no estaba seleccionada al ejecutar el mailing |
| `No Enviado`       | Email inválido, ausente o con observación de error |
| `Enviado correo N` | Enviado con éxito (N = número de grupo correlativo)|
| `Envío Fallido`    | Error SMTP al intentar enviar                      |

---

## Colores del Treeview

| Tag         | Fondo     | Texto    | Cuándo                      |
|-------------|-----------|----------|-----------------------------|
| `even`      | `#FFFFFF` | `#222222`| Fila par, no seleccionada   |
| `odd`       | `#E8EDF2` | `#222222`| Fila impar, no seleccionada |
| `even_sel`  | `#5B9BD5` | blanco   | Fila par, seleccionada      |
| `odd_sel`   | `#2E75B6` | blanco   | Fila impar, seleccionada    |

> Los tags anulan el color de selección nativo del tema clam (configurado a nivel Tcl con `ttk::style theme settings clam { ttk::style map Treeview -background {} }`).

---

## Colores de bloque (cabeceras)

| Constante    | Hex       | Sección                       |
|--------------|-----------|-------------------------------|
| `BG_CARPETA` | `#1A5276` | Carpetas de Adjuntos          |
| `BG_TABLA`   | `#7D3C98` | Tabla Adj. Personalizados     |
| `BG_COMUNES` | `#1E8449` | Tabla Adj. Comunes            |
| `BG_LOG_HDR` | `#5D6D7E` | Informe                       |
| `BG_CONFIG`  | `#1B2631` | Cabecera panel Configuración  |

---

## Persistencia (Registro de Windows)

Clave raíz: `HKCU\Software\MailingPersonalizado`

| Valor registro          | Contenido                                 |
|-------------------------|-------------------------------------------|
| `LastAdjPersonalizados` | Última carpeta de adjuntos personalizados |
| `LastAdjComunes`        | Última carpeta de adjuntos comunes        |
| `CfgMailCta`            | Cuenta Gmail SMTP                         |
| `CfgMailFrom`           | Remitente visible                         |
| `CfgMailClau`           | App Password (cifrada en Windows)         |
| `CfgMailFirm`           | Firma HTML completa                       |
| `CfgDirPrueba`          | Dirección de correo para pruebas          |
| `CfgFontFamily`         | Tipo de letra del cuerpo                  |
| `CfgFontSize`           | Tamaño de letra del cuerpo                |
| `CfgAnimacion`          | Animación del cartero (`1`/`0`, def. `1`) |
| `CfgTipsInicio`         | Tooltips visibles al iniciar (`1`/`0`, def. `0`) |
| `CfgAsunto`             | Asunto del correo                         |
| `CfgSaludoEs`           | Saludo en español                         |
| `CfgSaludoVa`           | Saludo en valencià                        |
| `CfgCuerpoEs`           | Cuerpo del mensaje en español             |
| `CfgCuerpoVa`           | Cuerpo del mensaje en valencià            |
| `CfgDespedidaEs`        | Despedida en español                      |
| `CfgDespedidaVa`        | Despedida en valencià                     |
| `CfgLangOrder`          | Orden de idiomas (`es_va` / `va_es`)      |
| `CfgUserExt`            | Extensión telefónica (firma)              |
| `CfgWebEs`              | URL web en español (firma)                |
| `CfgWebVa`              | URL web en valencià (firma)               |
| `CfgServicio`           | Nombre del servicio (firma)               |
| `CfgUnidad`             | Nombre de la unidad (firma)               |

---

## Archivos del proyecto

| Fichero                      | Descripción                        |
|------------------------------|------------------------------------|
| `Mailing_Personalizado.py`   | Código fuente principal            |
| `Mailing_Personalizado.exe`  | Ejecutable compilado (PyInstaller) |
| `build.bat`                  | Script de recompilación            |
| `requirements.txt`           | Dependencias Python (`openpyxl`, `deep-translator`, `reportlab`, `Pillow`) |
| `Mail.png` / `MailSent.gif`  | Iconos: logo del Informe y cartero animado del envío |
| `Flujo_APP.md`               | Este documento                     |
