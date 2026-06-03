# Mailing Personalizado

## Descripción

Aplicación de escritorio Python/tkinter para el **envío masivo de correos personalizados
con adjuntos** a los departamentos de la Universidad de Alicante, compilable a `.exe`.

Origen del proyecto: fork del proyecto VBA `Mailing_Indicadores` (referencia funcional).

---

## 🚀 Estado actual

- **Fase**: en producción — código Python funcional, compilable a `.exe` con PyInstaller
- **Carpeta**: `F:\__Dugarry UA\Dugarry Proyectos\___Claude.Code\Mailing_Personalizado`
- Los archivos VBA (`VBA_Moduls/`, `.xlsm`) están presentes solo como **referencia histórica**

---

## 📂 Rutas importantes

| Ruta | Descripción |
|------|-------------|
| `Mailing_Personalizado.py` | 🐍 **Código fuente principal** |
| `Mailing_Personalizado.exe` | 📦 Ejecutable compilado (PyInstaller) |
| `build.bat` | ⚙️ Script de compilación |
| `requirements.txt` | 📋 Dependencias Python |
| `VBA_Moduls/` | 📝 Módulos VBA originales — **solo referencia** |
| `.claude/` | ⚙️ Configuración de Claude Code |

---

## 🛠️ Stack

- **Python 3.10+** + **tkinter** — UI de escritorio
- **openpyxl** — lectura de emails desde hojas `Correo` de los xlsx
- **smtplib** (SSL 465) — envío SMTP Gmail con App Password
- **deep-translator** — traducción automática ES → VA
- **Pillow (PIL)** — GIF del cartero (overlay animado) e iconos
- **reportlab** — exportación del Flujo de la App a PDF
- **PyInstaller** — compilación a `.exe`
- **Registro Windows** (`Software\MailingPersonalizado`) — persistencia de configuración
- **ctypes / DWM** — transparencia y esquinas redondeadas (Windows 11)

---

## 🏗️ Arquitectura Python

### Flujo principal

```
Seleccionar carpeta Adj. Personalizados
  → Lee xlsx, extrae email de hoja "Correo" A1
  → Observación: Sin Hoja / Sin Correo / Correo Erróneo / (vacío=OK)

Seleccionar carpeta Adj. Comunes (opcional)

Configurar (panel ⚙️):
  Cuenta SMTP, Mail Password, Dir. prueba
  Asunto, Saludo (es/va), Cuerpo (es/va), Despedida (es/va)
  Firma HTML, Variables de firma, Tipografía
  Checkboxes: Animación (cartero) · Textos explicativos al iniciar APP

Ejecutar Mailing:
  → 1 correo por dirección única válida
  → Adjuntos: todos los xlsx de esa dirección + comunes seleccionados
  → Overlay animado del cartero (se desliza, con el nº de correo sobre el sobre)
  → Columna "Envío": Enviado Nº / Envío Fallido / No Enviado
  → Diálogo final con desglose de "No Enviado" por categoría
```

### Ayuda contextual (caracolillo + tooltips)

- El logo **caracolillo** (barra Carpetas de Adjuntos) es el control de ayuda:
  hover = firma + explicación · **clic izq** = mostrar/ocultar tooltips de botones ·
  **clic der** = panel Flujo.
- Tooltips gestionados por `_DynTooltip` con interruptor maestro de clase
  (`_DynTooltip.enabled`) y esquinas redondeadas (DWM). Estado de arranque: `CfgTipsInicio`.

### Patrones extraídos a skills (reutilizables)

Dos comportamientos de UI de este proyecto se han generalizado como **skills globales**
de Claude Code (en `F:\.claude\skills\`), reutilizables en otros proyectos:

- **`tkinter-moving-overlay`** — overlay de imagen/GIF transparente que se mueve/desliza
  sobre la ventana durante una tarea, con texto superpuesto (el cartero del mailing).
- **`tkinter-help-tooltips`** — tooltips de ayuda con interruptor maestro (caracolillo +
  textos explicativos de los botones).

### Registro Windows

Clave raíz: `HKCU\Software\MailingPersonalizado`

---

**Última actualización**: 2026-06-03
**Responsable**: Dugarry
**Estado**: ✅ En producción
**Origen**: Fork de `Mailing_Indicadores` (VBA)
