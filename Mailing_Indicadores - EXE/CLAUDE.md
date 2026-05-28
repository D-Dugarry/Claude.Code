# Mailing Indicadores — Contexto del Proyecto

## Descripción

Aplicación Excel VBA para el envío masivo de correos electrónicos personalizados con
adjuntos a los departamentos de la **Universidad de Alicante (UA)**. Cada departamento
recibe su propio fichero de indicadores académicos como adjunto.

---

## 🚀 Primeros pasos para Claude Code

### Antes de trabajar en este proyecto:

1. ✅ **Revisar los ficheros principales**
   - `Mailing_Con_Adjunto.xlsm` — fichero Excel con la lógica VBA
   - `Departamentos.xlsx` — lista de departamentos con emails y nombres de adjunto
   - `Tablas_Indicadores/` — carpeta con los ~70 ficheros Excel de indicadores

2. ✅ **Entender el flujo principal**
   - El usuario ejecuta `Rut_Load_Tabla_Adjuntos` para cargar los ficheros en `Tb_Datos`
   - Pega los emails en Col 1 desde `Departamentos.xlsx`
   - Ejecuta `Enviar_Emails` — el motor CDO (M_800) lee credenciales de la hoja `Prog__APP`

3. ✅ **Si necesitas modificar VBA:**
   - Edita el `.bas` / `.cls` en `VBA_Moduls/`
   - Haz commit git con descripción clara
   - En Excel: Alt+F11 → clic derecho en módulo → "Quitar módulo" → Archivo → Importar fichero
   - Guarda el `.xlsm`

### Limitaciones actuales:
- El proyecto usa **VBA (Excel)**, no Python puro
- Las credenciales de Gmail están en la hoja `Prog__APP` (rangos nombrados) — no en el código
- Los `.bas` usan encoding **Windows-1252** — VS Code está configurado para abrirlos correctamente
- ⚠️ Al exportar módulos desde Excel, verificar que los acentos no se corrompan

---

## 📂 Rutas importantes

| Ruta | Descripción |
|------|-------------|
| `Mailing_Con_Adjunto.xlsm` | 📊 **Fichero principal (fuente de verdad)** |
| `Departamentos.xlsx` | 📋 Lista de destinatarios con email y nombre de adjunto |
| `Tablas_Indicadores/` | 📁 ~70 ficheros Excel de indicadores por departamento (año 2023) |
| `VBA_Moduls/` | 📝 Módulos VBA exportados (editar aquí, reimportar al Excel) |
| `.claude/` | ⚙️ Configuración de Claude Code |
| `.vscode/` | ⚙️ Configuración VS Code (encoding Windows-1252 para .bas/.cls) |

---

## 🛠️ Herramientas y dependencias

### Software requerido:
- **Excel 2019+** — Para abrir y ejecutar el `.xlsm`
- **Git** — Para control de versiones y commits
- **Cuenta Gmail con App Password** — Para el envío CDO (2FA obligatorio)

### Referencia COM requerida en Excel:
- **Microsoft CDO for Windows 2000 Library**
  - Herramientas → Referencias → activar la referencia

---

## 📋 Ficheros de datos

| Fichero | Descripción |
|---------|-------------|
| `Departamentos.xlsx` | Fuente de emails y nombres de fichero para poblar `Tb_Datos` |
| `Tablas_Indicadores/*.xlsx` | ~70 ficheros de indicadores, uno por departamento UA (año 2023) |

---

## 🏗️ Arquitectura VBA

### Módulos exportados en `VBA_Moduls/`

| Fichero | Función |
|---------|---------|
| `M_0_Ini_APP.bas` | Inicialización de la UI al abrir el libro (`Rut_Iniciar_APP`) |
| `M_1_Gestión_Correos.bas` | Lógica principal de envío masivo (`Enviar_Emails`) |
| `M_2_Inicializar_Tabla.bas` | Limpieza de `Tb_Datos` (`Rut_Inicializar_Tabla`) |
| `M_3_Load_Tabla_Adjuntos.bas` | Carga de ficheros en `Tb_Datos` (`Rut_Load_Tabla_Adjuntos`) |
| `M_800_Mail_Send_New.bas` | **Motor CDO**: envío, credenciales, errores (`Rut_Email_Send`) |
| `M_810_Mail_Valid.bas` | Validación de emails (`Fnc_Valid_Email`, `Fnc_Valid_Email_Multi`) |
| `M_815_Mail_HTML.bas` | Generación de HTML para cuerpos de correo (`Fnc_HTML_Tabla`) |
| `M_820_Range_TO_HTML.bas` | Conversión de rango Excel a HTML (`Fnc_RangeToHTML`) |
| `DatosCorreo.cls` | Sheet class — hoja con `Tb_Datos` y celdas de configuración |
| `ThisWorkbook.cls` | `Workbook_Open` → llama a `Rut_Iniciar_APP` |

### Tabla `Tb_Datos` (en hoja `DatosCorreo`)

| Columna | Contenido |
|---------|-----------|
| Col 1 | Email del destinatario |
| Col 2 | Nombre del fichero adjunto (sin ruta) |
| Col 3 | Estado: `"Enviado"` / `"Sin Destinatario"` / `"Email no válido"` / `"El Fichero NO Existe"` |

### Celdas de configuración (hoja `DatosCorreo`)

| Celda | Contenido |
|-------|-----------|
| `B3` | Asunto del correo |
| `B4` | Línea 1 del cuerpo |
| `B5` | Línea 2 del cuerpo |
| `B6` | Línea 3 del cuerpo (firma) |

### Rangos nombrados en hoja `Prog__APP` (Config_APP)

| Rango | Contenido |
|-------|-----------|
| `APP_MailCta` | Cuenta Gmail de autenticación SMTP (ej. `ingresos@gcloud.ua.es`) |
| `APP_MailFrom` | Remitente visible (ej. `ingresos@ua.es`). Si vacío, usa `APP_MailCta` |
| `APP_MailClau` | App Password de Google (16 chars, sin espacios) |
| `APP_MailFirm` | Plantilla HTML de la firma (con placeholders) |
| `APP_User_Ext` | Extensión telefónica |
| `APP_Web_es` | URL web en español |
| `APP_Web_va` | URL web en valenciano |
| `APP_Servicio` | Nombre del servicio |
| `APP_Unidad` | Nombre de la unidad |
| `SW_Test` | `TRUE` = modo prueba (redirige todos los correos a `dugarry@gcloud.ua.es`) |
| `SW_WB_Deactivate` | Control interno para `Fnc_RangeToHTML` — no modificar manualmente |

### Flujo de ejecución

```
Workbook_Open → Rut_Iniciar_APP (M_0) — configura UI

        ↓  [preparar envío]

Rut_Load_Tabla_Adjuntos (M_3):
  · FileDialog → carpeta con .xlsx
  · Filtra solo .xlsx y añade sus nombres a Col 2 de Tb_Datos

Usuario pega emails en Col 1 desde Departamentos.xlsx

        ↓  [envío]

Enviar_Emails (M_1):
  1. Doble confirmación
  2. FileDialog → carpeta de adjuntos
  3. Ordena Tb_Datos por Col1 (email) y Col2 (fichero)
  4. Por cada fila:
     · Si email vacío → "Sin Destinatario"
     · Si email inválido (M_810) → "Email no válido"
     · Si fichero no existe → "El Fichero NO Existe"
     · Si todo OK → llama a Rut_Email_Send (M_800) → "Enviado"
  5. MsgBox con tiempo y nº de emails enviados

        ↓  [motor CDO — M_800]

Rut_Email_Send:
  · Lee credenciales de Prog__APP (APP_MailCta, APP_MailClau, APP_MailFrom)
  · Si SW_Test=TRUE → redirige a dugarry@gcloud.ua.es y prefija "[PRUEBA]" al asunto
  · Construye CDO.Message (HTML body + firma + adjunto)
  · Configura SMTP: smtp.gmail.com:465 SSL
  · Envía; gestiona errores con mensajes descriptivos
```

---

## ⚠️ Seguridad e información crítica

### Email (CDO + Gmail)
- **Motor**: CDO (Collaboration Data Objects) con SMTP seguro
- **Servidor**: `smtp.gmail.com` (puerto 465, SSL — CDO no soporta STARTTLS/587)
- **Autenticación**: App Password de Google (requiere 2FA activado)
- **Credenciales**: en rango `APP_MailClau` de hoja `Prog__APP` — ⚠️ **NUNCA** en el código ni en commits
- **Modo prueba**: `SW_Test = TRUE` redirige todos los correos a `dugarry@gcloud.ua.es`

### Módulos M_800–M_820 (módulos compartidos)
Los módulos `M_800`, `M_810`, `M_815` y `M_820` son **idénticos** a los del proyecto
`Certificados_de_Pago`. Si se mejoran en uno, actualizar en el otro.

---

## 📌 Notas importantes

- Hay ~70 departamentos con su fichero de indicadores en `Tablas_Indicadores/` (año 2023)
- `M_3_Load_Tabla_Adjuntos` solo carga `.xlsx` (constante `EXT_FILTRO`); cambiar a `""` para todos los tipos
- La validación de email (M_810) usa RegExp — requiere `VBScript.RegExp` (disponible en Windows por defecto)
- `Fnc_RangeToHTML` (M_820) auto-detecta si Excel generó el HTML en UTF-8 o Windows-1252

---

**Última actualización**: 2026-05-28
**Responsable**: Dugarry
**Estado**: 🔄 En uso / revisión
