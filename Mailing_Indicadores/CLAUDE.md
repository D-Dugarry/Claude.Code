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
   - El Excel lee la lista de `Departamentos.xlsx`
   - Para cada fila, abre/adjunta el fichero correspondiente de `Tablas_Indicadores/`
   - Envía el correo personalizado con CDO (Gmail SMTP)

3. ✅ **Si necesitas modificar VBA:**
   - Edita el `.bas` / `.cls` en `VBA_Moduls/`
   - Haz commit git con descripción clara
   - En Excel: Alt+F11 → clic derecho en módulo → "Quitar módulo" → Archivo → Importar fichero
   - Guarda el `.xlsm`

### Limitaciones actuales:
- El proyecto usa **VBA (Excel)**, no Python puro
- La contraseña de Gmail se pide por `InputBox` en cada ejecución (no se almacena)
- Los adjuntos deben existir en la carpeta seleccionada con el nombre exacto de Col 2 en `Tb_Datos`
- ⚠️ Los `.bas` exportados pueden tener acentos corruptos (Windows-1252 vs UTF-8) — verificar tras editar

---

## 📂 Rutas importantes

| Ruta | Descripción |
|------|-------------|
| `Mailing_Con_Adjunto.xlsm` | 📊 **Fichero principal (fuente de verdad)** |
| `Departamentos.xlsx` | 📋 Lista de destinatarios con email y nombre de adjunto |
| `Tablas_Indicadores/` | 📁 ~70 ficheros Excel de indicadores por departamento (año 2023) |
| `.claude/` | ⚙️ Configuración de Claude Code |
| `CLAUDE.md` | 📖 Este archivo (documentación) |

---

## 🛠️ Herramientas y dependencias

### Software requerido:
- **Excel 2019+** — Para abrir y ejecutar el `.xlsm`
- **Git** — Para control de versiones y commits
- **Cuenta Gmail con App Password** — Para el envío CDO

---

## 📋 Ficheros de datos

| Fichero | Descripción |
|---------|-------------|
| `Departamentos.xlsx` | Fuente de emails y nombres de fichero para poblar `Tb_Datos` |
| `Tablas_Indicadores/*.xlsx` | ~70 ficheros de indicadores, uno por departamento UA (año 2023) |

---

## 🏗️ Arquitectura VBA

### Módulos exportados en `VBA_Moduls/`

| Fichero | Tipo | Función |
|---------|------|---------|
| `Mandar_Correos.bas` | Módulo estándar | Lógica principal de envío, carga de ficheros y UI |
| `DatosCorreo.cls` | Sheet class | Hoja que contiene la tabla `Tb_Datos` |
| `ThisWorkbook.cls` | Workbook class | Evento `Workbook_Open` → llama a `Rut_Iniciar` |

### Subrutinas en `Mandar_Correos.bas`

| Subrutina | Función |
|-----------|---------|
| `Enviar_Emails()` | Rutina principal: pide carpeta + contraseña, itera `Tb_Datos`, envía CDO |
| `Borrar_Datos_Tabla()` | Borra todos los datos de `Tb_Datos` con confirmación |
| `Añadir_Lista_Ficheros()` | Rellena `Tb_Datos` con los nombres de fichero de una carpeta seleccionada |
| `Rut_Iniciar()` | Configura la UI de Excel (pantalla completa, oculta barras y ribbon) |

### Tabla `Tb_Datos` (en hoja `DatosCorreo`)

| Columna | Contenido |
|---------|-----------|
| Col 1 | Email del destinatario |
| Col 2 | Nombre del fichero adjunto |
| Col 3 | Estado del envío ("Enviado" / "Sin Destinatario" / "El Fichero NO Existe") |

### Celdas de configuración (hoja `DatosCorreo`)

| Celda | Contenido |
|-------|-----------|
| `B2` | Dirección del emisor (From) |
| `B3` | Asunto del correo |
| `B4` | Línea 1 del cuerpo |
| `B5` | Línea 2 del cuerpo |
| `B6` | Línea 3 del cuerpo (firma) |

### Flujo de ejecución

```
Workbook_Open → Rut_Iniciar (configura UI)
        ↓
Añadir_Lista_Ficheros → rellena Col 2 de Tb_Datos con los .xlsx de Tablas_Indicadores/
        ↓  (el usuario rellena Col 1 con emails copiados de Departamentos.xlsx)
Enviar_Emails:
  1. Confirma doble con MsgBox
  2. FileDialog → elige carpeta de adjuntos
  3. Ordena Tb_Datos por Col 1 (email) y Col 2 (fichero)
  4. InputBox → pide App Password
  5. Itera filas → CDO.Message por destinatario
     • Si mismo email en filas consecutivas → adjunta múltiples ficheros en un solo correo
     • Marca Col 3 con resultado: "Enviado" / error
  6. MsgBox con tiempo y conteo de emails enviados
```

---

## ⚠️ Seguridad e información crítica

### Email (CDO + Gmail)
- **Motor**: CDO (Collaboration Data Objects) con SMTP seguro
- **Servidor**: `smtp.gmail.com` (puerto 465, SSL)
- **Autenticación**: App Password de Google (requiere 2FA activado en la cuenta)
- ⚠️ **NUNCA** commitear contraseñas o App Passwords al repositorio

### Correspondencia de ficheros adjuntos
- Los nombres de fichero en `Tablas_Indicadores/` deben coincidir exactamente con lo referenciado en `Departamentos.xlsx`
- Si un fichero no existe o el nombre no coincide, el envío fallará para ese destinatario

---

## 📌 Notas importantes

- Hay ~70 departamentos con su fichero de indicadores en `Tablas_Indicadores/` (año 2023)
- El mismo destinatario puede recibir varios adjuntos: basta con repetir su email en filas consecutivas de `Tb_Datos`
- `Añadir_Lista_Ficheros()` puebla automáticamente Col 2 — útil al preparar un nuevo envío
- ⚠️ Los `.bas` exportados contienen caracteres corruptos visibles (ej. `¿`, `ó`) — son acentos Windows-1252; no afectan al funcionamiento en Excel pero sí al editarlos fuera

---

**Última actualización**: 2026-05-26
**Responsable**: Dugarry
**Estado**: 🔄 En uso / revisión
