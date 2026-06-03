"""
Mailing_Personalizado.py  v1
Envío masivo de correos personalizados con adjuntos. Universidad de Alicante.
Requiere: Python 3.9+, tkinter (stdlib), openpyxl
"""

import os
import random
import re
import ssl
import smtplib
import sys
import threading
import tkinter as tk
import tkinter.font as tkfont
from datetime import datetime
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from tkinter import filedialog, messagebox, scrolledtext, ttk
import ctypes, ctypes.wintypes
import winreg

try:
    import openpyxl
    _HAS_OPENPYXL = True
except ImportError:
    _HAS_OPENPYXL = False

try:
    from deep_translator import GoogleTranslator
    _HAS_DEEP_TRANSLATOR = True
except ImportError:
    _HAS_DEEP_TRANSLATOR = False

try:
    from PIL import Image, ImageTk
    _HAS_PIL = True
except ImportError:
    _HAS_PIL = False


# ── Registro Windows ──────────────────────────────────────────────────────────

_REG_KEY      = r"Software\MailingPersonalizado"
_REG_ADJ_PERS = "LastAdjPersonalizados"
_REG_ADJ_COM  = "LastAdjComunes"

# Claves de configuración (todas bajo _REG_KEY)
_CFG_KEYS = {
    "mail_cta":      "CfgMailCta",
    "mail_from":     "CfgMailFrom",
    "mail_clau":     "CfgMailClau",
    "mail_firm":     "CfgMailFirm",
    "user_ext":      "CfgUserExt",
    "web_es":        "CfgWebEs",
    "web_va":        "CfgWebVa",
    "servicio":      "CfgServicio",
    "unidad":        "CfgUnidad",
    "dir_prueba":    "CfgDirPrueba",
    "font_family":   "CfgFontFamily",
    "font_size":     "CfgFontSize",
    "tips_inicio":   "CfgTipsInicio",
    "animacion":     "CfgAnimacion",
    "asunto":        "CfgAsunto",
    "saludo_es":     "CfgSaludoEs",
    "saludo_va":     "CfgSaludoVa",
    "despedida_es":  "CfgDespedidaEs",
    "despedida_va":  "CfgDespedidaVa",
    "cuerpo_es":     "CfgCuerpoEs",
    "cuerpo_va":     "CfgCuerpoVa",
    "lang_order":    "CfgLangOrder",
}

_FONTS_ACCESIBLES = ["Arial", "Verdana", "Calibri", "Trebuchet MS", "Georgia"]
_FONT_SIZES       = ["12", "13", "14", "15", "16"]


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


def _cfg_read(name: str) -> str:
    return _reg_read(_CFG_KEYS[name])


def _cfg_write(name: str, value: str) -> None:
    _reg_write(_CFG_KEYS[name], value)


# ── SMTP ──────────────────────────────────────────────────────────────────────

_SMTP_SERVER = "smtp.gmail.com"
_SMTP_PORT   = 465
_DIR_PRUEBA_DEFAULT = "dugarry@gcloud.ua.es"
_EMAIL_RE = re.compile(r'^[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}$')


def _valid_email(email: str) -> bool:
    return bool(_EMAIL_RE.match(email.strip()))


def _fmt_size(n: int) -> str:
    if n < 1024:
        return f"{n} B"
    if n < 1_048_576:
        return f"{n / 1024:.0f} KB"
    if n < 1_073_741_824:
        return f"{n / 1_048_576:.0f} MB"
    return f"{n / 1_073_741_824:.1f} GB"


# ── Assets / constantes UI ────────────────────────────────────────────────────

_FIRMA_B64 = 'iVBORw0KGgoAAAANSUhEUgAAACgAAAAgCAYAAABgrToAAAAACXBIWXMAADqYAAA6mAGHJxjCAAAOb0lEQVR42o1YC4xc5Xn9JSgNOLaCzCukvIxDeNhFdUKcQLGiJtQkoCQ4L+K4tLQNpFGbIuqItqhQI9VCrahapXElsBKhGhyXxCiA3/Y+5j135r7mcR9z5877ubO79u56d3Zm587pubMGTJDaanT3zt479/+//3znnO/7rwAgPnAMIYbD4ej7cDD8wL2ZV5ui8HlXqJtVodyTEM7XLFF+piLUHeld8ZvkA9lr0ofdq9QTzlXJoHpjQtK26JO17aWjmd/PvTF+38SrZz5xdFt2vSZSGxThfCsnzv7JlJh745zw4K3Ot+LPNxzN/24MHwhg9eJwNciV1Yf6Cz1RejYnal9Ki9xfmyL/r65wHy2uzW5JH8x8UoXzZQfunhIax5qonyhh6u0Kv1dQPV6AfUxHs9lCtV2DezgP+3EL0Q0hOFcYL9e2lETrK3WR56e1syGWGucvCvL9mH4jOJ49Ro/V7+dC0yJ6ryyMp0oi92xatJ9yt5l/Kj/h7MrB+IkLR60h3y0iDwMpfkI8IshARZp/05B5XRtkkUCQ11JodRuonc6j9u0iKh8pI397HpXvuwemrmiKqZs7Yj46N5rbD3IUz/BCgBcHNzovQjgv2EK+PyCmAlNC3WPeqW63GvYfaWgRqQJykKAh7gezEoPWDyGxovBIQulJiPUVKEsTUPoSEgMdCS+GGO/HBgkGG+WCCigdKqJ4q43Zj3RgftqQCxvrwr20KBr76+8h6Qf7Ib4t15dF/tasKD5ui3KwKNSvh/bm/txGqTpLRPIMK4FMdxxyX0V0YEIdnkJiGEDUS8HsJWF3FaR6DhQkEUYA0vAoFyEj5mmYGAYRHHBRy2EuLYFiYQrOFytoXc6xNyfhfLx4fU90Rfvl1oiXXt9bRdAjcp7niaWZntDvUkTl8YyY2dt+0rw7Vc7ty6LE9CiYhNYNIr4SQmx4mkGkkRzkEPHOIMSJ44M0kYxCGcpIzSgIHA0iw/th7yisrs4zx1iRkO5mofLZ8CDKDMRhDgw4zxVwViygdXUNM+tm766LadH+VXuVgyPFXEDP3mWJ/KOOKL1QFNb6OKaPlWHBxfggBmMxw4lSkL0IuRZAkmk0Zizonok4v0sDmQHHERqEIa+oUP9RQk2qEkUdyeUE5IHCQ4MykMhTCUmPBOmFIXljKHKW6jer6Io+jBtzmL9y6fr2hrZYnusKMeh5IzjVp3PC+DNdVF6rXK2v11AP1EjtGE724kgs+2kzGITDVMmIDCIkvYaQNA7j7Qzqi01MegGo/Tjk5TifS6J8zMHJ544j6xUxucIr/TDRHYPVizFIlfykrHopZkQnVSIolGwUtlaxeMkKmltaDHZOqI9FxEgpU8GOCD+tilqgsy6+MRppnKlyenKly1SQ4NqAkw/GcZroSBzcWuSgXhS12QqCP3gH5aRDfgVgnpdhMoXKchTmnALtLQv1eYsBn4TSJdp9DQGOJ3vE1UsiPtQoGZ4JgMIll49UMHVpB1NX1NBZW3vWuVIRoj/sCe15XXQOTIvj2ymISR9wKnKZDzJVClNnzaXhLMtMLFPFwYyuRKLLFI2F8bFTcE/aSJFTypKG1CLTSzVreQmFfIHqjSAwCFFACei9DMaGvN+VkZ6NUeHjI/VHyNs41V9YKqD6YAOLYhrFzVZ76bLzQrg/KYvEjxTxzv2nROU/azSQCrLzVB0HjfrIrRjQT+sYezKM8mtF2CtZCkJCtKfBoAii0SjqUhtZhjtBdkW4CH2FvH0zjEa7RBnQaoYSkU8y2AS0YQLJrIXA7iRqAwvpRZn36ZNd37gM5N+qoHlZA9WNBcz99sIdQnvWFuNfCosz28ZerJbrI14kez5CcSKVpgDiTEcCJbWO8ENnYAUpChI/SuvIoYrk60lY36U9H6J4DgURejMKY08e2dsszB5soqHPwl7IQmNwsf4kyjCRDWWQ/y8TGS9HUaUolCDS/QTFGOI5h/oDTUxd1oa+Of2keP3+X4hjXzglsk+59RztJE4zlbjiyDBGBMmM4RmcHo6TKTIKLfrbIXJlyYQ9l0Hmp+QT1Vo4aMDOmaimp1CyqXvFgLvbhXWzhVNrVRj/lKdVMRM8nFQesb8NobhYwqkhzZw2pfeJ/0p0tIA0YyjvL6C8tg51W+qAiN6ti1/fe+xq9+clT6P0A4NJcsVCmAFGPB3ZUTVQEF3WR0Su2g1Efkz+/LcOKUmz7VY4pIXTTHCSwqIh8ZxnouMouC7sg22kPmug8dMyykoT0hcl5N/07/vWdALpHtHttvh/gN9p4FR1ueCgtKEG+boUxIk/nBCBHeGdmXdYO4fEqefzxaZq6VXkhtKzmeo8UzrBqbPIvJyB/A9MUdehHyo0XBozF2Gd01GcyML9Nwf50DSDTDMLYbqojdpbVbREHrmHJBhFl0iWaEkSMn1/LpOOwJpNm1KY4jCNW18w0b6tjfbHWhCHNr8uUptMyThnchWTMJesUf1UPN8FJ2iuKQomSow05F5NI7onAndYpDVEGDQH7Y0hN8Oi/3UHU6KN0lVlRO9i6P8SRnE5Q45JaL7dwKyoo/UzvyaxutAhDPqiRt4lPaaZ9qUyQGlFoyFlodM1Fm9YQGWdC5FcHxXWj5imnk4esEKs0EBJaF95YSKa7k8gRWuwnAJCXz2BYp3W4aWJgG8pMSIVQOHvGyiJKqovNlEpsWOxazAfpiLZfkWHadSO1zEvzqL6SomC05gl2hWFF2EVinl+w0G6DEmtZYqGlpObVbBwwxyM67MQtTWVJ5xMFbZnI72g4RStM84fyzTkqD8IraY2rCC8l9f3RmhCeUS6JlUZRbEXRSlbR+53Wpj6jwbR8W1XGgnKiCURfYPp8qpwX84T3UW4r8xC9zua3klSKIw4OR6nMY3DL4MhZOmxMjPnnM2iuKmI1HUMcHbNOWgm6+hQQXopzRVEaaYTVJa/yiSrgITceQOFRwrIuwVkiGaCFiPRgAtwkNrpwLk5gyIrRqg3jsxiip7mt1URZGei/I2B3K4spq500CnNEO84ieSzN84KpfOgPIiy4rEZIZVk8lY9R1f4vRwStyQh2mtrKKQchhEiJ+Ij5CSWoww5kmDdDfqQL2TgPuRCbWr8P4D8eYVKH6NmOfkPaaprae+sGgmGE13JEJEE7SmDaS6gEiE/iXDpoykYeyXUDzJbh004OZf12GW6U5zXF8kYs0VbJwXc2Qbcu2wc2fTmV8XspXU4aRJ7cIa8Yh1gB5wlegrrZZL2QqfjipJw/zkHcyHPmnuc7VIek4PTVCBRNKeQvzaPzo48mrkO8tMVaNN1lMNlOLvzcC5vw9zURHPyLIx32B09wyq/h3aSyWIMfk85OVJzmnMHyX13WEX+NdrQx0PQvxy5QTh/wP5YzpE3VBQ/fuMZZM+X4AMy+za5LyO6FIB+QEGux056GCK5La6WNXmJafQkmIkK0vepkK9IYOozVTRun0b78il0hIPpT7VQy1VpOhQhFRzz0WV6/W7IH0P35yKlfHrp/FWpbEPeEEDqXrVg/KUqRFSKi2P//kvIr8RRzbXgzhf5qF9Xg+SLPNpblLMGIvtopKy9iZVJEjxGU9Vgs3WSBmMkh4pW6xxKL3A7cE8d539rHu3rSigfYpLbeZh+B7Ps13ZaVs/koiN0ixi7b/aXg8iqx7IQ5DslKI+qyNxDVX9N2tH+RV2IidePiKxhrJF2653SHRmY2w0YL7qos/Vxx7KYO9aBvjWFzA9MKrgOadmvw4mRHSRXTHbJGqS+X0XGqeIYaUImlgpQXjqFerHCJSYQ7oXod1wujT+9bI06JBY2PsvGo09hsq2tniuj9tkqUut1ZL6nHi7udgV67KilfUmRfFAW1ftsoe7Uv5F7Oq+7D1p254FiqbRVayoiNlvZVVtUn9c9V/d3ahyQnUmQagv6PQzrqN1jwmg5MR56j15KjdZyFYT2B5Cd9b3NtysaMNswlWoNDVP8LhE98piLrM7NoL6jhZk1bdgPZxbz3zPWdc50hDe4sCfp1ruivq8srL/IicLDjug8Rtv9fkpUHymK3I6CMPfZIuNw+/ntcLvGhtRYUmAt0SL6yZHSox5ZRfWqfW6i+n4zG+fh95SsTufNVRGsxJEi+v7OLsX6m59ny0bsDDaylW9V4FxSQHq7AvMR7abpfc3VvZK/afIubNDff7PADVSfe5S+/4PB6JrzM0PU9qtCfy7+kvxd+h6bJonNhLJElyQSLHi+vY/qqR+w37JFByf9DQOTR0WucHO/RHS5PT3K4OPLaZj0U2e2iNJ3sihfwrbuMyy0N06IzM6U6J3tjeLyt8GrG2XvwmbZe39XP7ywX/avecW+kL8gieLbLaHcY87Ku5KYqk5Tk9aISxJR0dhhK6wuIX+LwE1UhDRILftWRdUPIyNjjzDF8dG23kR9og5ro4ziNVmkPp1YnLxq/KbsY1kx6Hsf2KOLD72b+Y3jXYQ7h5bExPqEKP3KvbH1V2VodybQ/LtpFCIktu9n/NCYRmIxyM4gg5fZyWSJoS+JBP9P08jP8vftXRnkPmEht9lF8UnrcOEb2XX1v3HEsDtYfTfz7gsEAvS/BvfuS5zRexo+1Pl5Q2Q3qyL3hCWqz1cOO7fUUbmmjupOG/ZLKmqn6IFGDbVsFfVMDY10A800fTDVQPpVVqNdDorXllBcYyNwS/Ds2O+GtyUe1MXsxNx7+/OLg/t/IfhesBceXCzMCecRW5Q+2hHGFvvW6h9b+9M3JZftz+WgP6BBu41t66dSSN+uIX2HBvWTrPF3s8mihVmXZvwGIHVkw9HHj2w4IsLfDIl5c36VZlTse285hsMPvzz6v1C8+I2XxwHPhRbEzBNNYW/MCPlzstAeVoV2Z/JyaW3kTulj0a/Eroz8ML4++kzs2sh37E36Vnurdc3JW4Pi6OfHhPpjU0xNzHxo8f4co+Oi+f8HLX3x4X4+BpkAAAAASUVORK5CYII='

BASE       = 14
FONT_UI    = ("Verdana", BASE)
FONT_BOLD  = ("Verdana", BASE, "bold")
FONT_TITLE = ("Verdana", BASE + 1, "bold")
FONT_ENTRY = ("Verdana", BASE - 1)
FONT_LOG   = ("Courier New", BASE)

BG_APP      = "#F2F3F4"
BG_CARPETA  = "#1A5276"
BG_TABLA    = "#7D3C98"
BG_COMUNES  = "#1E8449"
BG_LOG_HDR  = "#5D6D7E"
BG_CONFIG   = "#1B2631"

C_SELEC     = "#D5D8DC"
C_TODOS     = "#AED6F1"
C_NINGUNO   = "#F5B7B1"
C_EXPORTAR  = "#A9DFBF"
C_ACCION    = "#F0B27A"
C_PRUEBA_ON = "#E74C3C"
C_CONFIG    = "#AAB7B8"

_COL1_IDS   = ("correo", "fichero", "observacion", "envio", "fmodif", "fsize")
_COL1_NAMES = ("Correo", "Fichero", "Observación", "Envío", "F. Modif.", "Tam.")
_COL2_IDS   = ("fichero", "fmodif", "fsize")
_COL2_NAMES = ("Fichero", "F. Modif.", "Tam.")

_FIRMA_DEFAULT = """\
<div dir="ltr" class="gmail_signature" data-smartmail="gmail_signature"><div dir="ltr">
<table style="font-family:&quot;Times New Roman&quot;;width:750px"><tbody><tr><td colspan="2" style="width:750px"><a href="https://www.google.com/maps/place/Universitat+d'Alacant/@38.3852488,-0.5165048,17z/data=!3m1!4b1!4m5!3m4!1s0xd6236bb72bf619b:0x506e11c403138428!8m2!3d38.3852446!4d-0.5143161" rel="noopener" target="_blank"><img src="https://web.ua.es/es/accesibilidad/imagenes/gestadm/logo-ua.png" alt="Universitat d'Alacant / Universidad de Alicante" width="280"></a><br></td></tr>
<tr><td style="width:68px"><br></td><td style="width:682px">
<p><span style="font-family:&quot;trebuchet ms&quot;,sans-serif;color:rgb(211,116,40);font-size:13pt">Servei _____</span></p>
<p><span style="font-family:&quot;trebuchet ms&quot;,sans-serif;color:rgb(211,116,40);font-size:13pt">Unitat _____</span></p>
<p><span style="font-family:&quot;trebuchet ms&quot;,sans-serif;font-size:12pt"><span style="font-size:12pt"><font color="#444444">Tel. (+34) 965 90 34 00 ext. ____</font></span></span><span style="font-family:&quot;trebuchet ms&quot;,sans-serif;font-size:12pt"><br><font color="#444444"><span style="font-size:12pt">E-mail:&nbsp;</span><a href="mailto:ingresos@ua.es" target="_blank">Ingresos@ua.es</a><br></font></span><span style="font-family:&quot;trebuchet ms&quot;,sans-serif;font-size:12pt"><font color="#444444">Web:&nbsp;</font></span><span style="font-family:&quot;trebuchet ms&quot;,sans-serif;font-size:12pt"><a href="https://sc.ua.es/va/_____" target="_blank">Valencià</a>&nbsp;/&nbsp;<a href="https://sc.ua.es/es/_____" target="_blank">Español</a>&nbsp;</span></p></td></tr>
<tr><td style="text-align:right;width:68px"><img src="https://si.ua.es/va/imagenes/logos/no-imprimisques.png" alt="No imprimisques / No imprimas" style="font-family:&quot;trebuchet ms&quot;,sans-serif;font-size:12pt"><span style="font-family:&quot;trebuchet ms&quot;,sans-serif;font-size:12pt">&nbsp;</span></td><td style="width:682px"><font color="#666666"><font face="trebuchet ms, sans-serif"><span style="font-size:12pt">No imprimiu el correu si no cal. Per un ús responsable del paper.</span></font><br><font face="trebuchet ms, sans-serif"><span style="font-size:12pt">No imprimir el correo si no es necesario. Por un uso responsable del papel.</span></font></font></td></tr></tbody></table>
<table style="font-family:&quot;Times New Roman&quot;;width:750px"><tbody></tbody></table></div></div>"""

_FLUJO_MD = """\
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
- La ruta se guarda en el registro de Windows (`HKCU\\Software\\MailingPersonalizado\\LastAdjPersonalizados`).
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
- La ruta persiste en `HKCU\\Software\\MailingPersonalizado\\LastAdjComunes`.

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

Clave raíz: `HKCU\\Software\\MailingPersonalizado`

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
"""

_INLINE_RE = re.compile(r'\*\*(.+?)\*\*|`(.+?)`')

_C_CONTENIDO = "#D6EAF8"   # fondo etiquetas de contenido de correo
C_GOOGLE     = "#A9CCE3"   # botones Google (azul suave)
C_LANG_ES    = "#AED6F1"   # toggle idioma cuando ES primero
C_LANG_VA    = "#A9DFBF"   # toggle idioma cuando VA primero
C_PREVIEW    = "#D2B4DE"   # botón Previsualizar (lila)
C_MUESTRA    = "#F9E79F"   # botón Mandar 1 Muestra (amarillo suave)
_C_SALUDO_DEP = "#A9DFBF"  # fondo etiquetas y botones Saludo / Despedida (verde suave)

# Grupo izquierdo: cuenta SMTP + tipografía + asunto
_CFG_LEFT = [
    ("Cuenta Gmail (SMTP):",           "mail_cta",   False),
    ("Remitente visible:",              "mail_from",  False),
    ("Mail Password:",                  "mail_clau",  True),
    ("Correo para testear el Mailing:", "dir_prueba", False),
    # fila especial con dos comboboxes (marcador interno):
    ("__FONTS__",),
    None,
    ("Asunto del Correo:",              "asunto",     False, _C_CONTENIDO),
]
# Grupo derecho: variables de la firma
_CFG_RIGHT = [
    ("Extensión tel.:", "user_ext",  False),
    ("Web ES:",         "web_es",    False),
    ("Web VA:",         "web_va",    False),
    ("Servicio:",       "servicio",  False),
    ("Unidad:",         "unidad",    False),
]


def _insert_inline(txt: "tk.Text", text: str, base_tag: str) -> None:
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


class _RECT(ctypes.Structure):
    _fields_ = [("left",   ctypes.c_long), ("top",    ctypes.c_long),
                ("right",  ctypes.c_long), ("bottom", ctypes.c_long)]

class _MONITORINFO(ctypes.Structure):
    _fields_ = [("cbSize",    ctypes.c_ulong),
                ("rcMonitor", _RECT),
                ("rcWork",    _RECT),
                ("dwFlags",   ctypes.c_ulong)]


def _monitor_rect(x: int, y: int) -> tuple[int, int, int, int]:
    """Devuelve (left, top, right, bottom) del monitor que contiene (x, y)."""
    try:
        u32  = ctypes.windll.user32
        pt   = ctypes.wintypes.POINT(x, y)
        hmon = u32.MonitorFromPoint(pt, ctypes.c_uint(2))   # DEFAULTTONEAREST
        mi   = _MONITORINFO()
        mi.cbSize = ctypes.sizeof(_MONITORINFO)             # obligatorio antes de llamar
        if u32.GetMonitorInfoW(hmon, ctypes.byref(mi)):
            r = mi.rcWork
            return r.left, r.top, r.right, r.bottom
    except Exception:
        pass
    sm = ctypes.windll.user32.GetSystemMetrics
    return 0, 0, sm(0), sm(1)


def _round_corners(win: "tk.Misc", small: bool = True) -> None:
    """Redondea ligeramente las esquinas de una ventana sin borde (Windows 11 DWM)."""
    try:
        win.update_idletasks()
        u32 = ctypes.windll.user32
        u32.GetParent.restype  = ctypes.wintypes.HWND
        u32.GetParent.argtypes = [ctypes.wintypes.HWND]
        hwnd = u32.GetParent(win.winfo_id())
        _DWMWA_WINDOW_CORNER_PREFERENCE = 33
        _DWMWCP_ROUND      = 2   # redondeo estándar
        _DWMWCP_ROUNDSMALL = 3   # redondeo ligero
        pref = ctypes.c_int(_DWMWCP_ROUNDSMALL if small else _DWMWCP_ROUND)
        dwm = ctypes.windll.dwmapi
        dwm.DwmSetWindowAttribute.argtypes = [
            ctypes.wintypes.HWND, ctypes.wintypes.DWORD,
            ctypes.c_void_p, ctypes.wintypes.DWORD]
        dwm.DwmSetWindowAttribute(hwnd, _DWMWA_WINDOW_CORNER_PREFERENCE,
                                  ctypes.byref(pref), ctypes.sizeof(pref))
    except Exception:
        pass


def _base_dir() -> str:
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))

def _resource_path(filename: str) -> str:
    """Ruta a un recurso empaquetado: usa _MEIPASS en --onefile, _base_dir en desarrollo."""
    base = getattr(sys, "_MEIPASS", None) or _base_dir()
    return os.path.join(base, filename)


class _DynTooltip:
    """Tooltip cuyo texto se evalúa en tiempo de muestra (callable)."""
    enabled = False   # gate global: ¿se muestran los tooltips explicativos de botones?

    def __init__(self, widget: tk.Widget, text_fn, always: bool = False) -> None:
        self._w   = widget
        self._fn  = text_fn
        self._always = always   # True = ignora el gate global (siempre visible)
        self._win: tk.Toplevel | None = None
        _add = "+" if always else None   # add="+" solo donde debe coexistir con otro tooltip
        widget.bind("<Enter>",       self._show, add=_add)
        widget.bind("<Leave>",       self._hide, add=_add)
        widget.bind("<ButtonPress>", self._hide, add=_add)

    def _show(self, _=None) -> None:
        if self._win:
            return
        if not self._always and not _DynTooltip.enabled:
            return
        self._win = tk.Toplevel(self._w)
        self._win.wm_overrideredirect(True)
        self._win.configure(bg="white",
                            highlightbackground=BG_CARPETA, highlightthickness=1)
        tk.Label(self._win, text=self._fn(), bg="white", fg=BG_CARPETA,
                 relief="flat", bd=0, font=FONT_UI, padx=10, pady=5,
                 justify="left").pack(anchor="w")
        self._win.update_idletasks()
        tip_w = self._win.winfo_reqwidth()
        tip_h = self._win.winfo_reqheight()
        x     = self._w.winfo_rootx()
        y     = self._w.winfo_rooty() + self._w.winfo_height() + 4
        ml, mt, mr, mb = _monitor_rect(x, y)   # rectángulo del monitor actual
        if x + tip_w > mr - 4:
            x = mr - tip_w - 4
        if y + tip_h > mb - 4:
            y = self._w.winfo_rooty() - tip_h - 4   # aparece encima
        x = max(ml + 4, x)
        y = max(mt + 4, y)
        self._win.wm_geometry(f"+{x}+{y}")
        _round_corners(self._win)

    def _hide(self, _=None) -> None:
        if self._win:
            self._win.destroy()
            self._win = None


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
        _BG = BG_CARPETA
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

    def __init__(self):
        super().__init__()
        self.title("Mailing Personalizado")
        self.resizable(True, True)
        self.minsize(700, 540)
        self.configure(bg=BG_APP)

        self._adj_pers_var  = tk.StringVar()
        self._adj_com_var   = tk.StringVar()
        self._after_id_pers = None
        self._after_id_com  = None
        self._prueba_activa = True   # activo por defecto
        self._pers_gen      = 0
        self._config_dirty          = False
        self._modif_correos_activo  = False
        self._lang_order    = "es_va"   # "es_va" = Español primero | "va_es" = Valencià primero

        self._firma_img = tk.PhotoImage(data=_FIRMA_B64)

        self._mail_img = None
        if _HAS_PIL:
            try:
                _img_path = _resource_path("Mail.png")
                _pil = Image.open(_img_path).resize((36, 36), Image.LANCZOS)
                self._mail_img = ImageTk.PhotoImage(_pil)
            except Exception:
                pass

        self._gif_frames: list = []
        self._gif_delay:  int  = 80
        self._gif_base = None       # primer frame PIL, base para componer el número encima
        self._gif_lbl  = None
        self._gif_num_photo = None  # ref viva de la imagen compuesta (evita recolección)
        if _HAS_PIL:
            try:
                _gif_path = _resource_path("MailSent.gif")
                _gif = Image.open(_gif_path)
                for i in range(_gif.n_frames):
                    _gif.seek(i)
                    _frame = _gif.convert("RGBA")
                    if i == 0:
                        self._gif_base = _frame.copy()
                    self._gif_frames.append(ImageTk.PhotoImage(_frame))
                # duration puede venir como 0 → forzar mínimo sano (evita busy-loop)
                self._gif_delay = _gif.info.get("duration") or 80
            except Exception:
                pass

        self._sending_overlay: tk.Toplevel | None = None
        self._gif_after_id = None
        self._gif_overlay_wh = (0, 0)   # tamaño del overlay (se fija al mostrarlo)
        # Deslizamiento suave del overlay (trayectoria con desaceleración)
        self._glide_after_id = None
        self._gliding        = False
        self._glide_ease     = 0.10   # fracción del trayecto por frame (más bajo = más lento)
        self._glide_ms       = 16     # ~60 fps
        self._glide_pos      = (0.0, 0.0)
        self._glide_target   = (0.0, 0.0)

        self._build_styles()
        self._build_ui()
        self.update_idletasks()
        self.state("zoomed")
        self.protocol("WM_DELETE_WINDOW", self._on_close)

        def _dir() -> str:
            return _cfg_read("dir_prueba") or _DIR_PRUEBA_DEFAULT

        _DynTooltip(self._btn_prueba,
                    lambda: (f'Para mandar todo el Mailing a "{_dir()}"'
                             if not self._prueba_activa
                             else f'Se mandará todo el Mailing a "{_dir()}"'))

        _DynTooltip(self._btn_mailing,
                    lambda: ("Mandaré todo el Mailing a cada dirección de correo "
                             "especificada con su/sus adjuntos."
                             if not self._prueba_activa
                             else f'Mandaré todo el Mailing a "{_dir()}" con sus Adjuntos'))

        self._adj_pers_var.trace_add("write", self._on_adj_pers_change)
        self._adj_com_var.trace_add("write",  self._on_adj_com_change)

        last_pers = _reg_read(_REG_ADJ_PERS)
        if last_pers and os.path.isdir(last_pers):
            self._adj_pers_var.set(last_pers)
        last_com = _reg_read(_REG_ADJ_COM)
        if last_com and os.path.isdir(last_com):
            self._adj_com_var.set(last_com)

        # Estado inicial de los textos explicativos según la preferencia guardada
        _DynTooltip.enabled = (_cfg_read("tips_inicio") == "1")

    # ── Estilos ───────────────────────────────────────────────────────────────

    def _build_styles(self) -> None:
        s = ttk.Style(self)
        s.theme_use("clam")
        self.tk.eval(
            "ttk::style theme settings clam "
            "{ ttk::style map Treeview -background {} }"
        )
        s.configure(".",                font=FONT_UI,   background=BG_APP)
        s.configure("TFrame",           background=BG_APP)
        s.configure("TEntry",           font=FONT_UI,   fieldbackground="white")
        s.configure("TScrollbar",       background=BG_APP)
        s.configure("Sel.TButton",      font=FONT_UI,   background=C_SELEC)
        s.map("Sel.TButton",            background=[("active", "#BFC9CA")])
        s.configure("Todos.TButton",    font=FONT_UI,   background=C_TODOS)
        s.map("Todos.TButton",          background=[("active", "#85C1E9")])
        s.configure("Ninguno.TButton",  font=FONT_UI,   background=C_NINGUNO)
        s.map("Ninguno.TButton",        background=[("active", "#F1948A")])
        s.configure("Export.TButton",   font=FONT_BOLD, background=C_EXPORTAR)
        s.map("Export.TButton",         background=[("active", "#7DCEA0")])
        s.configure("Accion.TButton",   font=FONT_UI,   background=C_ACCION)
        s.map("Accion.TButton",         background=[("active", "#E59866")])
        # Botón inline (mismo alto que etiqueta)
        s.configure("Inline.TButton",   font=FONT_ENTRY, padding=(6, 0), background=C_SELEC)
        s.map("Inline.TButton",         background=[("active", "#BFC9CA")])
        # Botón Google (azul suave)
        s.configure("Google.TButton",   font=FONT_ENTRY, padding=(6, 0), background=C_GOOGLE)
        s.map("Google.TButton",         background=[("active", "#7FB3D3")])
        # Toggle idioma ES
        s.configure("LangES.TButton",   font=FONT_UI, background=C_LANG_ES)
        s.map("LangES.TButton",         background=[("active", "#85C1E9")])
        # Toggle idioma VA
        s.configure("LangVA.TButton",   font=FONT_UI, background=C_LANG_VA)
        s.map("LangVA.TButton",         background=[("active", "#7DCEA0")])
        # Previsualizar (lila)
        s.configure("Preview.TButton",  font=FONT_UI, background=C_PREVIEW)
        s.map("Preview.TButton",        background=[("active", "#BB8FCE")])
        s.configure("Muestra.TButton",   font=FONT_UI, background=C_MUESTRA)
        s.map("Muestra.TButton",         background=[("active", "#F4D03F")])
        s.configure("SaluDep.TButton",   font=FONT_ENTRY, padding=(6, 0),
                    background=_C_SALUDO_DEP)
        s.map("SaluDep.TButton",         background=[("active", "#7DCEA0")])
        # Guardar: limpio (verde) / pendiente (rosa)
        s.configure("SaveClean.TButton", font=FONT_UI, background=C_EXPORTAR)
        s.map("SaveClean.TButton",       background=[("active", "#7DCEA0")])
        s.configure("SaveDirty.TButton", font=FONT_BOLD, background="#F8BBD0")
        s.map("SaveDirty.TButton",       background=[("active", "#F48FB1")])
        s.configure("Prueba.TButton",   font=FONT_UI,   background=C_SELEC)
        s.map("Prueba.TButton",         background=[("active", "#BFC9CA")])
        s.configure("PruebaOn.TButton", font=FONT_BOLD, background=C_PRUEBA_ON,
                    foreground="white")
        s.map("PruebaOn.TButton",       background=[("active", "#C0392B")],
                                        foreground=[("active", "white")])
        s.configure("Config.TButton",   font=FONT_UI,   background=C_CONFIG)
        s.map("Config.TButton",         background=[("active", "#808B96")])
        s.configure("ModifOn.TButton",  font=FONT_BOLD, background="#F39C12",
                    foreground="white")
        s.map("ModifOn.TButton",        background=[("active", "#E67E22")],
                                        foreground=[("active", "white")])

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
                     foreground="#222222", rowheight=_row_h + 1, font=FONT_UI)
        s.configure("Treeview.Heading",
                     font=FONT_BOLD, background="#DDDDDD", relief="flat")
        s.map("Treeview", foreground=[("selected", "white")])

    # ── Bloque genérico ───────────────────────────────────────────────────────

    def _bloque(self, titulo: str, bg_titulo: str,
                expand: bool = False,
                parent: tk.Widget | None = None) -> tuple[tk.Frame, tk.Frame]:
        p = parent if parent is not None else self
        outer = tk.Frame(p, bg=BG_APP, bd=1, relief="solid",
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

    # ── UI principal ──────────────────────────────────────────────────────────

    def _build_ui(self) -> None:

        # ── Carpetas de Adjuntos ─────────────────────────────────────────────
        tb_carp, c_carp = self._bloque("Carpetas de Adjuntos", BG_CARPETA)
        tb_carp.columnconfigure(1, weight=1)
        tk.Label(tb_carp, text="", bg=BG_CARPETA).grid(row=0, column=1, sticky="ew")
        _firma_lbl = tk.Label(tb_carp, image=self._firma_img,
                              bg=BG_CARPETA, bd=0, cursor="hand2")
        _firma_lbl.grid(row=0, column=2, padx=(0, 8), pady=2)
        _firma_lbl.bind("<Button-3>", lambda _: self._show_flujo())
        _firma_lbl.bind("<Button-1>", lambda _: self._toggle_tips())
        _Tooltip(_firma_lbl, "Dugarry'26")
        _DynTooltip(_firma_lbl,
                    lambda: ("Ocultar Texto explicativo de Botones"
                             if _DynTooltip.enabled else
                             "Mostrar Texto explicativo de Botones"),
                    always=True)

        tk.Label(c_carp, text="Adj. Personalizados:", bg=BG_APP,
                 font=FONT_UI, anchor="w").grid(
            row=0, column=0, sticky="w", padx=(0, 10), pady=(4, 2))
        tk.Entry(c_carp, textvariable=self._adj_pers_var,
                 font=FONT_ENTRY, bg="white", fg="#222222",
                 relief="sunken", bd=1).grid(
            row=0, column=1, sticky="ew", pady=(4, 2), ipady=4)
        _btn_sel_pers = ttk.Button(c_carp, text="Seleccionar", style="Sel.TButton",
                                   command=self._pick_adj_pers)
        _btn_sel_pers.grid(row=0, column=2, padx=(8, 0), pady=(4, 2), ipadx=4)
        _DynTooltip(_btn_sel_pers, lambda: "Seleccionar la carpeta de los archivos Excel a adjuntar en su correo definido en el propio fichero.")

        tk.Label(c_carp, text="Adj. Comunes:", bg=BG_APP,
                 font=FONT_UI, anchor="w").grid(
            row=1, column=0, sticky="w", padx=(0, 10), pady=(2, 4))
        tk.Entry(c_carp, textvariable=self._adj_com_var,
                 font=FONT_ENTRY, bg="white", fg="#222222",
                 relief="sunken", bd=1).grid(
            row=1, column=1, sticky="ew", pady=(2, 4), ipady=4)
        _btn_sel_com = ttk.Button(c_carp, text="Seleccionar", style="Sel.TButton",
                                  command=self._pick_adj_com)
        _btn_sel_com.grid(row=1, column=2, padx=(8, 0), pady=(2, 4), ipadx=4)
        _DynTooltip(_btn_sel_com, lambda: "Seleccionar la carpeta de los archivos a adjuntar en todos los correos.")

        # ── Contenedor para las dos tablas + overlay de flujo ────────────────
        self._tables_frame = tk.Frame(self, bg=BG_APP)
        self._tables_frame.pack(fill="both", expand=True)
        self._tab_bloque = self._tables_frame

        # Sub-contenedor en dos columnas (personalizados más ancho)
        _tbl_grid = tk.Frame(self._tables_frame, bg=BG_APP)
        _tbl_grid.pack(fill="both", expand=True)
        _tbl_grid.columnconfigure(0, weight=3)   # personalizados: ~75%
        _tbl_grid.columnconfigure(1, weight=1)   # comunes: ~25%
        _tbl_grid.rowconfigure(0, weight=1)
        _col_pers = tk.Frame(_tbl_grid, bg=BG_APP)
        _col_pers.grid(row=0, column=0, sticky="nsew", padx=(0, 4))
        _col_com  = tk.Frame(_tbl_grid, bg=BG_APP)
        _col_com.grid(row=0, column=1, sticky="nsew", padx=(4, 0))

        # ── Tabla 1: Adjuntos Personalizados ──────────────────────────────────
        tb_tab1, c_tab1 = self._bloque(
            "Fich. Adjuntos Personalizados para cada Mail.",
            BG_TABLA, expand=True, parent=_col_pers)
        tb_tab1.columnconfigure(0, weight=1)
        self._lbl_count1 = tk.Label(tb_tab1, text="", bg=BG_TABLA, fg="white", font=FONT_UI)
        self._lbl_count1.grid(row=0, column=1, padx=(0, 6))
        _btn_todos1 = tk.Button(tb_tab1, text="Todos", font=FONT_UI,
                                bg=C_TODOS, activebackground="#85C1E9",
                                relief="flat", bd=1, padx=4, pady=1, cursor="hand2",
                                command=lambda: self._sel_all(self._tree1))
        _btn_todos1.grid(row=0, column=2, padx=(0, 2), pady=4)
        _DynTooltip(_btn_todos1, lambda: "Selecciona todos los ficheros Adj. Personalizados")
        _btn_ngno1 = tk.Button(tb_tab1, text="Ngno", font=FONT_UI,
                               bg=C_NINGUNO, activebackground="#F1948A",
                               relief="flat", bd=1, padx=4, pady=1, cursor="hand2",
                               command=lambda: self._sel_none(self._tree1))
        _btn_ngno1.grid(row=0, column=3, padx=(0, 6), pady=4)
        _DynTooltip(_btn_ngno1, lambda: "Deselecciona todos los ficheros Adj. Personalizados")
        self._btn_modif_correos = ttk.Button(
            tb_tab1, text="Activar Modif. Correos", style="Sel.TButton",
            command=self._toggle_modif_correos)
        self._btn_modif_correos.grid(row=0, column=4, padx=(0, 8), pady=4, ipadx=4)
        self._tip_modif_correos = _DynTooltip(self._btn_modif_correos, lambda: (
            "Selecciona la fila del fichero y se abrirá un cuadro de diálogo.\n"
            "Podrás modificar o añadir la dirección de correo en el fichero.\n"
            "Pulsa el botón para salir del modo 'Modificar Correo'."
            if self._modif_correos_activo else
            "Permite modificar o añadir, uno a uno, los correos dentro de los ficheros.\n"
            "Seleccionando la fila del fichero, se abrirá un cuadro de diálogo a tal efecto."))

        c_tab1.rowconfigure(0, weight=1)
        tf1 = tk.Frame(c_tab1, bg=BG_APP)
        tf1.grid(row=0, column=0, columnspan=3, sticky="nsew")
        tf1.rowconfigure(0, weight=1)
        tf1.columnconfigure(0, weight=1)

        self._tree1 = ttk.Treeview(tf1, columns=_COL1_IDS,
                                   show="headings", selectmode="extended")
        for col_id, col_name in zip(_COL1_IDS, _COL1_NAMES):
            anc = ("center" if col_id in ("fmodif", "observacion", "envio")
                   else "e" if col_id == "fsize" else "w")
            self._tree1.heading(col_id, text=col_name, anchor=anc,
                                command=lambda c=col_id: self._sort_col(
                                    self._tree1, self._sort_rev1,
                                    _COL1_IDS, self._lbl_count1, c))
            self._tree1.column(col_id, width=200, minwidth=60, stretch=True, anchor=anc)

        vsb1 = ttk.Scrollbar(tf1, orient="vertical",   command=self._tree1.yview)
        hsb1 = ttk.Scrollbar(tf1, orient="horizontal", command=self._tree1.xview)
        self._tree1.grid(row=0, column=0, sticky="nsew")
        vsb1.grid(row=0, column=1, sticky="ns")
        hsb1.grid(row=1, column=0, sticky="ew")
        self._tree1.configure(yscrollcommand=vsb1.set, xscrollcommand=hsb1.set)
        self._apply_tree_tags(self._tree1)
        self._tree1.bind("<<TreeviewSelect>>",
                         lambda e: self._on_sel_change(self._tree1, self._lbl_count1))
        self._tree1.bind("<MouseWheel>",
                         lambda e: self._tree1.yview_scroll(-1 if e.delta > 0 else 1, "units"))
        self._tree1.bind("<Button-1>", self._on_tree1_click)
        self._sort_rev1: dict[str, bool] = {c: False for c in _COL1_IDS}

        # ── Tabla 2: Adjuntos Comunes ─────────────────────────────────────────
        tb_tab2, c_tab2 = self._bloque(
            "Fich. Adjuntos Comunes para todos los Mails.",
            BG_COMUNES, expand=True, parent=_col_com)
        tb_tab2.columnconfigure(0, weight=1)
        self._lbl_count2 = tk.Label(tb_tab2, text="", bg=BG_COMUNES, fg="white", font=FONT_UI)
        self._lbl_count2.grid(row=0, column=1, padx=(0, 6))
        _btn_todos2 = tk.Button(tb_tab2, text="Todos", font=FONT_UI,
                                bg=C_TODOS, activebackground="#85C1E9",
                                relief="flat", bd=1, padx=4, pady=1, cursor="hand2",
                                command=lambda: self._sel_all(self._tree2))
        _btn_todos2.grid(row=0, column=2, padx=(0, 2), pady=4)
        _DynTooltip(_btn_todos2, lambda: "Selecciona todos los ficheros Adj. Comunes")
        _btn_ngno2 = tk.Button(tb_tab2, text="Ngno", font=FONT_UI,
                               bg=C_NINGUNO, activebackground="#F1948A",
                               relief="flat", bd=1, padx=4, pady=1, cursor="hand2",
                               command=lambda: self._sel_none(self._tree2))
        _btn_ngno2.grid(row=0, column=3, padx=(0, 6), pady=4)
        _DynTooltip(_btn_ngno2, lambda: "Deselecciona todos los ficheros Adj. Comunes")

        c_tab2.rowconfigure(0, weight=1)
        tf2 = tk.Frame(c_tab2, bg=BG_APP)
        tf2.grid(row=0, column=0, columnspan=3, sticky="nsew")
        tf2.rowconfigure(0, weight=1)
        tf2.columnconfigure(0, weight=1)

        self._tree2 = ttk.Treeview(tf2, columns=_COL2_IDS, show="headings",
                                   selectmode="extended")
        for col_id, col_name in zip(_COL2_IDS, _COL2_NAMES):
            anc = ("center" if col_id == "fmodif"
                   else "e" if col_id == "fsize" else "w")
            self._tree2.heading(col_id, text=col_name, anchor=anc,
                                command=lambda c=col_id: self._sort_col(
                                    self._tree2, self._sort_rev2,
                                    _COL2_IDS, self._lbl_count2, c))
            self._tree2.column(col_id, width=200, minwidth=60, stretch=True, anchor=anc)

        vsb2 = ttk.Scrollbar(tf2, orient="vertical",   command=self._tree2.yview)
        hsb2 = ttk.Scrollbar(tf2, orient="horizontal", command=self._tree2.xview)
        self._tree2.grid(row=0, column=0, sticky="nsew")
        vsb2.grid(row=0, column=1, sticky="ns")
        hsb2.grid(row=1, column=0, sticky="ew")
        self._tree2.configure(yscrollcommand=vsb2.set, xscrollcommand=hsb2.set)
        self._apply_tree_tags(self._tree2)
        self._tree2.bind("<<TreeviewSelect>>",
                         lambda e: self._on_sel_change(self._tree2, self._lbl_count2))
        self._tree2.bind("<MouseWheel>",
                         lambda e: self._tree2.yview_scroll(-1 if e.delta > 0 else 1, "units"))
        self._sort_rev2: dict[str, bool] = {c: False for c in _COL2_IDS}

        # ── Informe ───────────────────────────────────────────────────────────
        tb_log, c_log = self._bloque("Informe", BG_LOG_HDR)
        tb_log.columnconfigure(0, weight=1)   # espaciador izquierdo (título)
        tb_log.columnconfigure(3, weight=1)   # espaciador derecho — igual peso → botones centrados
        c_log.rowconfigure(0, weight=1)
        c_log.columnconfigure(0, weight=1)   # área de log expande
        c_log.columnconfigure(1, weight=0)   # columna Limpiar: estrecha

        _mailing_wrap = tk.Frame(tb_log, bg=BG_LOG_HDR)
        _mailing_wrap.grid(row=0, column=1, padx=(0, 100), pady=4)
        if self._mail_img:
            tk.Label(_mailing_wrap, image=self._mail_img,
                     bg=BG_LOG_HDR).pack(side="left", padx=(0, 14))
        self._btn_mailing = ttk.Button(
            _mailing_wrap, text="Ejecutar Mailing", style="Export.TButton",
            command=self._ejecutar_mailing)
        self._btn_mailing.pack(side="left", ipadx=6)

        self._btn_prueba = ttk.Button(
            tb_log, text="Prueba Activada", style="PruebaOn.TButton",
            command=self._toggle_prueba)
        self._btn_prueba.grid(row=0, column=2, padx=(0, 4), pady=4, ipadx=6)

        tk.Label(tb_log, text="", bg=BG_LOG_HDR).grid(row=0, column=3, sticky="ew")

        _btn_cfg = ttk.Button(tb_log, text="⚙  Configuración", style="Config.TButton",
                              command=self._show_config)
        _btn_cfg.grid(row=0, column=4, padx=(0, 8), pady=3, ipadx=4)
        _DynTooltip(_btn_cfg, lambda: "Permite modificar el texto del correo y los parámetros de la APP")

        # Text + Scrollbar manual para poder situar Limpiar a la derecha
        log_frame = tk.Frame(c_log, bg="#1E1E1E")
        log_frame.grid(row=0, column=0, sticky="nsew")
        log_frame.rowconfigure(0, weight=1)
        log_frame.columnconfigure(0, weight=1)

        self._log = tk.Text(
            log_frame, height=6, state="disabled", font=FONT_LOG, wrap="word",
            spacing1=2, spacing3=2,
            bg="#1E1E1E", fg="#D4D4D4", insertbackground="white",
            relief="flat", bd=0)
        log_vsb = ttk.Scrollbar(log_frame, orient="vertical", command=self._log.yview)
        self._log.configure(yscrollcommand=log_vsb.set)
        self._log.grid(row=0, column=0, sticky="nsew")
        log_vsb.grid(row=0, column=1, sticky="ns")
        self._log.bind("<MouseWheel>",
                       lambda e: self._log.yview_scroll(-1 if e.delta > 0 else 1, "units"))

        self._log.tag_configure("ok",    foreground="#4EC94E")
        self._log.tag_configure("error", foreground="#FF6B6B")
        self._log.tag_configure("info",  foreground="#85C1E9")
        self._log.tag_configure("warn",  foreground="#F0B27A")

        # Botón Limpiar: Canvas con texto girado 270° (top→bottom)
        _lf = tkfont.Font(family=FONT_UI[0], size=FONT_UI[1])
        _lw = _lf.metrics("linespace") + 14   # ancho = alto de fuente + padding
        _lcvs = tk.Canvas(c_log, width=_lw, bg=C_ACCION,
                          highlightthickness=1, highlightbackground="#D4895A",
                          cursor="hand2", relief="flat")
        _lcvs.grid(row=0, column=1, sticky="ns", padx=(4, 2))
        _ltxt = _lcvs.create_text(_lw // 2, 50,
                                  text="Limpiar", angle=270,
                                  font=FONT_UI, fill="#222222")
        _lcvs.bind("<Configure>",
                   lambda e, c=_lcvs, t=_ltxt, w=_lw:
                       c.coords(t, w // 2, e.height // 2))
        _lcvs.bind("<Enter>",    lambda e: _lcvs.configure(bg="#E59866"))
        _lcvs.bind("<Leave>",    lambda e: _lcvs.configure(bg=C_ACCION))
        _lcvs.bind("<Button-1>", lambda e: self._log_clear())
        _DynTooltip(_lcvs, lambda: "Borra el contenido de la ventana de Informe")

        tk.Frame(self, bg=BG_APP, height=8).pack()

    # ── Panel Flujo (overlay sobre ambas tablas) ──────────────────────────────

    def _toggle_tips(self) -> None:
        """Alterna la visibilidad de los tooltips explicativos de todos los botones
        del proyecto. El del caracolillo (always=True) no se ve afectado."""
        _DynTooltip.enabled = not _DynTooltip.enabled

    def _show_flujo(self) -> None:
        if hasattr(self, "_flujo_panel") and self._flujo_panel.winfo_ismapped():
            self._hide_flujo(); return
        if not hasattr(self, "_flujo_panel"):
            pan = tk.Frame(self._tab_bloque, bg=BG_APP)
            self._flujo_panel = pan
            hdr = tk.Frame(pan, bg=BG_TABLA)
            hdr.pack(fill="x")
            tk.Label(hdr, text="  Flujo de la Aplicación",
                     bg=BG_TABLA, fg="white", font=FONT_TITLE,
                     anchor="w", pady=4).pack(side="left", fill="x", expand=True)
            tk.Button(hdr, text="  ✕  ", command=self._hide_flujo,
                      bg=BG_TABLA, fg="white", activebackground="#C0392B",
                      activeforeground="white", relief="flat", bd=0,
                      font=FONT_BOLD, cursor="hand2").pack(side="right", padx=4)
            tk.Button(hdr, text="  ⬇  Descargar en PDF  ",
                      command=self._export_flujo_pdf,
                      bg=BG_TABLA, fg="#A9DFBF",
                      activebackground="#1E8449", activeforeground="white",
                      relief="flat", bd=0,
                      font=FONT_UI, cursor="hand2").pack(side="right", padx=(0, 12))
            self._flujo_txt = scrolledtext.ScrolledText(
                pan, font=FONT_UI, wrap="word", bg=BG_APP, fg="#222222",
                relief="flat", bd=0, spacing1=1, spacing3=1, padx=14, pady=6)
            self._flujo_txt.pack(fill="both", expand=True)
            self._render_md(self._flujo_txt, _FLUJO_MD)
        self._flujo_panel.place(x=0, y=0, relwidth=1, relheight=1)
        self._flujo_panel.lift()

    def _hide_flujo(self) -> None:
        if hasattr(self, "_flujo_panel"):
            self._flujo_panel.place_forget()

    def _export_flujo_pdf(self) -> None:
        try:
            from reportlab.platypus import (
                SimpleDocTemplate, Paragraph, Spacer,
                Table, TableStyle, HRFlowable)
            from reportlab.lib.styles import ParagraphStyle
            from reportlab.lib.pagesizes import A4
            from reportlab.lib.units import cm
            from reportlab.lib import colors
        except ImportError:
            messagebox.showwarning(
                "reportlab no disponible",
                "Instala la librería:\n  pip install reportlab",
                parent=self)
            return

        path = filedialog.asksaveasfilename(
            parent=self, title="Guardar Flujo en PDF",
            defaultextension=".pdf",
            filetypes=[("PDF", "*.pdf")],
            initialfile="Flujo_Mailing_Personalizado.pdf")
        if not path:
            return

        def _rl(text: str) -> str:
            text = (text.replace("&", "&amp;")
                        .replace("<", "&lt;")
                        .replace(">", "&gt;"))
            text = re.sub(r'\*\*(.+?)\*\*', r'<b>\1</b>', text)
            text = re.sub(r'`(.+?)`',
                          r'<font name="Courier" size="8">\1</font>', text)
            return text

        _BOX = str.maketrans({
            '│': '|',  '─': '-',  '┌': '+',  '┐': '+',
            '└': '+',  '┘': '+',  '├': '+',  '┤': '+',
            '┬': '+',  '┴': '+',  '┼': '+',  '✔': 'v',
            '✖': 'x',  '✉': '',   '👁': '',  '⚙': '',
            '⇄': '<>', '▲': '^',  '▼': 'v',  '•': '*',
        })

        try:
            PAGE_W = A4[0] - 4 * cm
            C_H1   = colors.HexColor('#7D3C98')
            C_H2   = colors.HexColor('#1A5276')
            C_H3   = colors.HexColor('#1E8449')
            C_TH   = colors.HexColor('#DDDDDD')
            C_TR2  = colors.HexColor('#F4F6F7')
            C_SEP  = colors.HexColor('#CCCCCC')
            C_CODE = colors.HexColor('#F0F2F3')
            C_QUO  = colors.HexColor('#5D6D7E')

            s_h1 = ParagraphStyle('h1', fontName='Helvetica-Bold', fontSize=16,
                                   textColor=C_H1, spaceBefore=14, spaceAfter=6)
            s_h2 = ParagraphStyle('h2', fontName='Helvetica-Bold', fontSize=13,
                                   textColor=C_H2, spaceBefore=10, spaceAfter=4)
            s_h3 = ParagraphStyle('h3', fontName='Helvetica-Bold', fontSize=11,
                                   textColor=C_H3, spaceBefore=8,  spaceAfter=3)
            s_n  = ParagraphStyle('n',  fontName='Helvetica', fontSize=10,
                                   spaceAfter=3)
            s_bl = ParagraphStyle('bl', fontName='Helvetica', fontSize=10,
                                   leftIndent=14, spaceAfter=2)
            s_q  = ParagraphStyle('q',  fontName='Helvetica-Oblique', fontSize=9,
                                   textColor=C_QUO, leftIndent=24, spaceAfter=3)
            s_cd = ParagraphStyle('cd', fontName='Courier', fontSize=8,
                                   leftIndent=8, spaceAfter=0, leading=11,
                                   backColor=C_CODE)
            s_th = ParagraphStyle('th', fontName='Helvetica-Bold', fontSize=9)
            s_td = ParagraphStyle('td', fontName='Helvetica',      fontSize=9)

            story: list = []
            in_code  = False
            code_buf: list[str] = []
            tbl_buf:  list[list[str]] = []

            def flush_table() -> None:
                if not tbl_buf:
                    return
                n = max(len(r) for r in tbl_buf)
                w = PAGE_W / n
                data = []
                for i, row in enumerate(tbl_buf):
                    sty = s_th if i == 0 else s_td
                    padded = row + [''] * (n - len(row))
                    data.append([Paragraph(_rl(c), sty) for c in padded])
                tbl = Table(data, colWidths=[w] * n,
                            repeatRows=1, hAlign='LEFT')
                tbl.setStyle(TableStyle([
                    ('BACKGROUND',    (0, 0), (-1, 0),  C_TH),
                    ('ROWBACKGROUNDS',(0, 1), (-1, -1), [colors.white, C_TR2]),
                    ('GRID',          (0, 0), (-1, -1), 0.5, C_SEP),
                    ('VALIGN',        (0, 0), (-1, -1), 'TOP'),
                    ('TOPPADDING',    (0, 0), (-1, -1), 3),
                    ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
                    ('LEFTPADDING',   (0, 0), (-1, -1), 4),
                ]))
                story.append(tbl)
                story.append(Spacer(1, 6))
                tbl_buf.clear()

            for line in _FLUJO_MD.split('\n'):
                stripped = line.strip()

                if stripped == '```':
                    if in_code:
                        story.append(Spacer(1, 2))
                        for cl in code_buf:
                            safe = (cl.translate(_BOX)
                                      .replace("&", "&amp;")
                                      .replace("<", "&lt;")
                                      .replace(">", "&gt;"))
                            story.append(Paragraph(safe, s_cd))
                        story.append(Spacer(1, 6))
                        code_buf.clear()
                        in_code = False
                    else:
                        flush_table()
                        in_code = True
                    continue

                if in_code:
                    code_buf.append(line)
                    continue

                if stripped.startswith('|'):
                    if all(c in '-|: ' for c in stripped):
                        continue
                    tbl_buf.append([c.strip() for c in stripped.split('|')[1:-1]])
                    continue
                else:
                    flush_table()

                if stripped in ('---', '***', '___'):
                    story.append(HRFlowable(width='100%', thickness=0.5,
                                            color=C_SEP))
                    story.append(Spacer(1, 4))
                elif line.startswith('# ') and not line.startswith('##'):
                    story.append(Paragraph(_rl(line[2:]), s_h1))
                elif line.startswith('## ') and not line.startswith('###'):
                    story.append(Paragraph(_rl(line[3:]), s_h2))
                elif line.startswith('### '):
                    story.append(Paragraph(_rl(line[4:]), s_h3))
                elif line.startswith('> '):
                    story.append(Paragraph(_rl(line[2:]), s_q))
                elif re.match(r'^( {0,4})-\s', line):
                    story.append(Paragraph(
                        '&#x2022;&nbsp;&nbsp;' + _rl(line.lstrip().lstrip('- ')),
                        s_bl))
                elif stripped:
                    story.append(Paragraph(_rl(stripped), s_n))
                else:
                    story.append(Spacer(1, 4))

            flush_table()

            doc = SimpleDocTemplate(
                path, pagesize=A4,
                rightMargin=2*cm, leftMargin=2*cm,
                topMargin=2*cm, bottomMargin=2*cm,
                title="Flujo — Mailing Personalizado",
                author="Dugarry UA")
            doc.build(story)
            self._log_write(f"  ✔  PDF guardado: {path}\n", "ok")
            messagebox.showinfo("PDF generado",
                                f"Guardado en:\n{path}", parent=self)

        except Exception as exc:
            messagebox.showerror("Error al generar PDF", str(exc), parent=self)

    # ── Panel Configuración (overlay sobre ventana completa) ──────────────────

    def _show_config(self) -> None:
        if hasattr(self, "_config_panel") and self._config_panel.winfo_ismapped():
            self._hide_config(); return
        if not hasattr(self, "_config_panel"):
            self._build_config_panel()
        for name, var in self._cfg_vars.items():
            saved = _cfg_read(name)
            if name == "dir_prueba" and not saved:
                saved = _DIR_PRUEBA_DEFAULT
            var.set(saved)
        self._cfg_cuerpo_es.delete("1.0", "end")
        self._cfg_cuerpo_es.insert("1.0", _cfg_read("cuerpo_es"))
        self._cfg_cuerpo_va.delete("1.0", "end")
        self._cfg_cuerpo_va.insert("1.0", _cfg_read("cuerpo_va"))
        # Valores por defecto para font y tamaño
        if not self._cfg_vars["font_family"].get():
            self._cfg_vars["font_family"].set(_cfg_read("font_family") or "Verdana")
        if not self._cfg_vars["font_size"].get():
            self._cfg_vars["font_size"].set(_cfg_read("font_size") or "13")
        if self._cfg_vars["tips_inicio"].get() not in ("0", "1"):
            self._cfg_vars["tips_inicio"].set(_cfg_read("tips_inicio") or "0")
        if self._cfg_vars["animacion"].get() not in ("0", "1"):
            self._cfg_vars["animacion"].set(_cfg_read("animacion") or "1")
        saved_firma = _cfg_read("mail_firm")
        self._cfg_firma.delete("1.0", "end")
        self._cfg_firma.insert("1.0", saved_firma if saved_firma else _FIRMA_DEFAULT)
        # Restaurar orden de idioma guardado
        saved_order = _cfg_read("lang_order")
        if saved_order in ("es_va", "va_es"):
            self._lang_order = saved_order
        self._btn_lang_order.config(text=self._lang_label())
        self._config_dirty = False
        self._update_guardar_btn()
        self._update_firma_visibility()
        self._config_panel.place(x=0, y=0, relwidth=1, relheight=1)
        self._config_panel.lift()
        self.after(60, self._adjust_firma_height)

    def _hide_config(self) -> None:
        if not hasattr(self, "_config_panel"):
            return
        if self._config_dirty:
            resp = messagebox.askyesnocancel(
                "Cambios sin guardar",
                "La configuración tiene cambios sin guardar.\n"
                "¿Guardar antes de cerrar?",
                parent=self)
            if resp is None:   # Cancelar → no cerrar
                return
            if resp:           # Sí → guardar y cerrar
                self._do_save_config()
        self._config_dirty = False
        self._config_panel.place_forget()

    def _save_config(self) -> None:
        self._do_save_config()
        self._config_dirty = False
        self._update_guardar_btn()

    def _do_save_config(self) -> None:
        for name, var in self._cfg_vars.items():
            _cfg_write(name, var.get().strip())
        _cfg_write("cuerpo_es", self._cfg_cuerpo_es.get("1.0", "end-1c"))
        _cfg_write("cuerpo_va", self._cfg_cuerpo_va.get("1.0", "end-1c"))
        _cfg_write("mail_firm", self._cfg_firma.get("1.0", "end-1c"))
        _cfg_write("lang_order", self._lang_order)
        self._log_write("  ✔  Configuración guardada.\n", "ok")

    def _build_config_panel(self) -> None:
        pan = tk.Frame(self, bg=BG_APP, bd=2, relief="raised")
        self._config_panel = pan
        self._firma_shown  = False   # oculta por defecto

        # ── Cabecera con toggle centrado ──────────────────────────────────────
        hdr = tk.Frame(pan, bg=BG_CONFIG)
        hdr.pack(fill="x")
        hdr.columnconfigure(1, weight=1)   # espaciador izquierdo
        hdr.columnconfigure(3, weight=1)   # espaciador derecho → toggle centrado

        tk.Label(hdr, text="  ⚙  Configuración APP",
                 bg=BG_CONFIG, fg="white", font=FONT_TITLE,
                 anchor="w", pady=4).grid(row=0, column=0, sticky="w")

        center_grp = tk.Frame(hdr, bg=BG_CONFIG)
        center_grp.grid(row=0, column=2, pady=3)
        self._btn_lang_order = ttk.Button(
            center_grp, text=self._lang_label(), style=self._lang_style(),
            command=self._toggle_lang_order)
        self._btn_lang_order.pack(side="left")
        _DynTooltip(self._btn_lang_order, lambda: "Selecciona el orden de visualización de los idiomas.")

        self._btn_muestra = ttk.Button(
            hdr, text="✉  Mandar 1 Muestra", style="Muestra.TButton",
            command=self._mandar_muestra)
        self._btn_muestra.grid(row=0, column=4, padx=(0, 4), pady=3, ipadx=6)
        # Tooltip dinámico con la dirección de prueba actual
        def _tip_text():
            d = (self._cfg_vars.get("dir_prueba", tk.StringVar()).get()
                 or _cfg_read("dir_prueba") or _DIR_PRUEBA_DEFAULT)
            return f'Mandar un correo de prueba a "{d}" sin adjunto.'
        _DynTooltip(self._btn_muestra, _tip_text)

        _btn_prev = ttk.Button(hdr, text="👁  Previsualizar", style="Preview.TButton",
                               command=self._show_preview)
        _btn_prev.grid(row=0, column=5, padx=(0, 4), pady=3, ipadx=6)
        _DynTooltip(_btn_prev, lambda: "Muestra el Cuerpo del Correo Montado.")
        self._btn_cancelar = ttk.Button(hdr, text="Borrar Cambios",
                                        style="Accion.TButton",
                                        command=self._revert_config)
        self._btn_cancelar.grid(row=0, column=6, padx=(0, 4), pady=3, ipadx=8)
        self._btn_cancelar.grid_remove()   # oculto hasta que haya cambios
        _DynTooltip(self._btn_cancelar,
                    lambda: "Restaura los valores de la última salvaguardia.")

        self._btn_guardar = ttk.Button(hdr, text="¡Guardado!",
                                       style="SaveClean.TButton",
                                       command=self._save_config)
        self._btn_guardar.grid(row=0, column=7, padx=(0, 4), pady=3, ipadx=8)
        _DynTooltip(self._btn_guardar,
                    lambda: ("Salvaguardar la Configuración."
                             if self._config_dirty
                             else "La Configuración actual está Salvaguardada."))

        tk.Button(hdr, text="  ✕  ", command=self._hide_config,
                  bg=BG_CONFIG, fg="white", activebackground="#C0392B",
                  activeforeground="white", relief="flat", bd=0,
                  font=FONT_BOLD, cursor="hand2").grid(row=0, column=8, padx=(0, 6))

        # ── Main ──────────────────────────────────────────────────────────────
        main = tk.Frame(pan, bg=BG_APP)
        main.pack(fill="both", expand=True, padx=10, pady=(8, 8))
        main.columnconfigure(0, weight=1)
        main.columnconfigure(1, weight=1)
        main.rowconfigure(2, weight=1)
        self._cfg_main = main

        # ── Fila 0: campos en dos grupos ──────────────────────────────────────
        self._cfg_vars: dict[str, tk.StringVar] = {}

        def _make_field_group(parent: tk.Frame, fields: list,
                              title: str = "") -> None:
            parent.columnconfigure(1, weight=1)
            r = 0
            if title:
                tk.Label(parent, text=title, bg=BG_APP,
                         font=FONT_BOLD, anchor="w", fg=BG_CARPETA).grid(
                    row=r, column=0, columnspan=2, sticky="w", pady=(0, 6))
                r += 1
            for field in fields:
                if field is None:
                    tk.Frame(parent, bg="#BBBBBB", height=1).grid(
                        row=r, column=0, columnspan=2, sticky="ew", pady=6)
                    r += 1; continue
                # Fila especial: Tipo de letra (col 0) + combo+tamaño (col 1)
                if field[0] == "__FONTS__":
                    tk.Label(parent, text="Tipo de letra:", bg=BG_APP,
                             font=FONT_UI, anchor="e").grid(
                        row=r, column=0, sticky="e", padx=(0, 8), pady=2)
                    inner = tk.Frame(parent, bg=BG_APP)
                    inner.grid(row=r, column=1, sticky="ew", pady=2)
                    vf = tk.StringVar(value=_FONTS_ACCESIBLES[1])
                    vf.trace_add("write", self._mark_config_dirty)
                    self._cfg_vars["font_family"] = vf
                    ttk.Combobox(inner, textvariable=vf, values=_FONTS_ACCESIBLES,
                                 state="readonly", width=13).pack(side="left")
                    tk.Label(inner, text="  Tamaño:", bg=BG_APP,
                             font=FONT_UI).pack(side="left", padx=(10, 4))
                    vs = tk.StringVar(value="13")
                    vs.trace_add("write", self._mark_config_dirty)
                    self._cfg_vars["font_size"] = vs
                    ttk.Combobox(inner, textvariable=vs, values=_FONT_SIZES,
                                 state="readonly", width=4).pack(side="left")
                    vt = tk.StringVar(value="0")
                    vt.trace_add("write", self._mark_config_dirty)
                    self._cfg_vars["tips_inicio"] = vt
                    tk.Checkbutton(
                        inner, text="Textos explicativos al iniciar APP",
                        variable=vt, onvalue="1", offvalue="0",
                        bg=BG_APP, activebackground=BG_APP, font=FONT_UI,
                        cursor="hand2").pack(side="right", padx=(14, 0))
                    # Animación del cartero (queda a la izquierda del check anterior)
                    va_anim = tk.StringVar(value="1")
                    va_anim.trace_add("write", self._mark_config_dirty)
                    self._cfg_vars["animacion"] = va_anim
                    tk.Checkbutton(
                        inner, text="Animación", variable=va_anim,
                        onvalue="1", offvalue="0",
                        bg=BG_APP, activebackground=BG_APP, font=FONT_UI,
                        cursor="hand2").pack(side="right", padx=(14, 0))
                    r += 1; continue
                lbl_text, key, pw = field[0], field[1], field[2]
                bg = field[3] if len(field) > 3 else BG_APP
                tk.Label(parent, text=lbl_text, bg=bg, font=FONT_UI,
                         anchor="e", padx=3).grid(
                    row=r, column=0, sticky="e", padx=(0, 8), pady=2)
                var = tk.StringVar()
                var.trace_add("write", self._mark_config_dirty)
                self._cfg_vars[key] = var
                tk.Entry(parent, textvariable=var, font=FONT_ENTRY,
                         bg="white", fg="#222222", relief="sunken", bd=1,
                         show="*" if pw else "", width=24).grid(
                    row=r, column=1, sticky="ew", pady=2, ipady=3)
                r += 1

        grp_left  = tk.Frame(main, bg=BG_APP)
        grp_right = tk.Frame(main, bg=BG_APP)
        grp_left.grid( row=0, column=0, sticky="nsew", padx=(0, 16))
        grp_right.grid(row=0, column=1, sticky="nsew")
        _make_field_group(grp_left,  _CFG_LEFT)
        _make_field_group(grp_right, _CFG_RIGHT,
                          title="Variables de la Firma del Correo")

        # ── Fila 1: separador ─────────────────────────────────────────────────
        tk.Frame(main, bg="#BBBBBB", height=1).grid(
            row=1, column=0, columnspan=2, sticky="ew", pady=(10, 4))

        # ── Fila 2: cuerpos lado a lado con rúbrica encima ───────────────────
        es_frame = tk.Frame(main, bg=BG_APP)
        va_frame = tk.Frame(main, bg=BG_APP)
        es_frame.grid(row=2, column=0, sticky="nsew", padx=(0, 6))
        va_frame.grid(row=2, column=1, sticky="nsew", padx=(6, 0))

        def _labeled_entry_row(parent: tk.Frame, row: int, label: str,
                               key: str, btns=None,
                               bg=_C_CONTENIDO, btn_style="Google.TButton") -> None:
            """Fila horizontal: etiqueta | entry | botones opcionales."""
            parent.columnconfigure(1, weight=1)
            tk.Label(parent, text=label, bg=bg, font=FONT_UI,
                     anchor="w", padx=4).grid(
                row=row, column=0, sticky="ew", pady=(2, 2))
            var = tk.StringVar()
            var.trace_add("write", self._mark_config_dirty)
            self._cfg_vars[key] = var
            tk.Entry(parent, textvariable=var, font=FONT_ENTRY,
                     bg="white", fg="#222222", relief="sunken", bd=1).grid(
                row=row, column=1, sticky="ew", pady=(2, 2), ipady=3)
            if btns:
                btn_frame = tk.Frame(parent, bg=BG_APP)
                btn_frame.grid(row=row, column=2, padx=(4, 0))
                for item_b in btns:
                    txt_b, cmd_b = item_b[0], item_b[1]
                    tip_b = item_b[2] if len(item_b) > 2 else None
                    _b = ttk.Button(btn_frame, text=txt_b, style=btn_style, command=cmd_b)
                    _b.pack(side="left", padx=2)
                    if tip_b:
                        _DynTooltip(_b, lambda t=tip_b: t)

        def _body_col(parent: tk.Frame, lbl_cuerpo: str,
                      key_saludo: str, lbl_saludo: str,
                      key_despedida: str, lbl_despedida: str,
                      saludo_btns=None, cuerpo_btns=None,
                      despedida_btns=None) -> scrolledtext.ScrolledText:
            parent.columnconfigure(1, weight=1)
            parent.rowconfigure(2, weight=1)   # cuerpo text expands

            # Fila 0: Saludo horizontal (etiqueta | entry [| Traducir])
            _labeled_entry_row(parent, 0, lbl_saludo, key_saludo, saludo_btns,
                               bg=_C_SALUDO_DEP, btn_style="SaluDep.TButton")

            # Fila 1: etiqueta Cuerpo ancho completo (justo encima del texto)
            hdr_c = tk.Frame(parent, bg=_C_CONTENIDO)
            hdr_c.grid(row=1, column=0, columnspan=3, sticky="ew", pady=(6, 0))
            hdr_c.columnconfigure(0, weight=1)
            tk.Label(hdr_c, text=lbl_cuerpo, bg=_C_CONTENIDO, font=FONT_UI,
                     anchor="w", padx=4).grid(row=0, column=0, sticky="ew")
            if cuerpo_btns:
                for ci, item_c in enumerate(cuerpo_btns, 1):
                    bt, bc = item_c[0], item_c[1]
                    tip_c = item_c[2] if len(item_c) > 2 else None
                    _bc = ttk.Button(hdr_c, text=bt, style="Google.TButton", command=bc)
                    _bc.grid(row=0, column=ci, padx=(4, 0))
                    if tip_c:
                        _DynTooltip(_bc, lambda t=tip_c: t)

            # Fila 2: ScrolledText del cuerpo (expande)
            txt = scrolledtext.ScrolledText(parent, font=FONT_ENTRY, bg="white",
                                            fg="#222222", relief="sunken", bd=1,
                                            wrap="word")
            txt.grid(row=2, column=0, columnspan=3, sticky="nsew")
            txt.bind("<Key>", self._mark_config_dirty)

            # Fila 3: Despedida horizontal (etiqueta | entry | botones)
            _labeled_entry_row(parent, 3, lbl_despedida, key_despedida, despedida_btns,
                               bg=_C_SALUDO_DEP, btn_style="SaluDep.TButton")
            return txt

        self._cfg_cuerpo_es = _body_col(
            es_frame,
            "Cuerpo del mensaje (es):",
            key_saludo="saludo_es",    lbl_saludo="Saludo (es):",
            key_despedida="despedida_es", lbl_despedida="Despedida (es):")

        self._cfg_cuerpo_va = _body_col(
            va_frame,
            "Cuerpo del missatge (va):",
            key_saludo="saludo_va",    lbl_saludo="Saludo (va):",
            key_despedida="despedida_va", lbl_despedida="Comiat (va):",
            saludo_btns=[("Traducir",
                lambda: self._translate_field_to_va("saludo_es", "saludo_va"),
                "Traduce del Español al Valencià.")],
            cuerpo_btns=[("Traducir", self._translate_google,
                "Traduce del Español al Valencià.")],
            despedida_btns=[("Traducir",
                lambda: self._translate_field_to_va("despedida_es", "despedida_va"),
                "Traduce del Español al Valencià.")])

        # ── Filas 3-4: Firma HTML (condicional) ───────────────────────────────
        firma_hdr = tk.Frame(main, bg=BG_APP)
        firma_hdr.columnconfigure(0, weight=1)
        tk.Frame(firma_hdr, bg="#BBBBBB", height=1).grid(row=0, column=0, sticky="ew")
        self._btn_firma_toggle = ttk.Button(
            firma_hdr, text="▼ Visualizar Firma", style="Sel.TButton",
            command=self._toggle_firma)
        self._btn_firma_toggle.grid(row=0, column=1, padx=(10, 0))
        self._firma_sep = firma_hdr

        self._firma_frame = tk.Frame(main, bg=BG_APP)
        self._firma_frame.rowconfigure(1, weight=1)
        self._firma_frame.columnconfigure(0, weight=1)
        tk.Label(self._firma_frame, text="Firma HTML:",
                 bg=_C_CONTENIDO, font=FONT_UI, anchor="w", padx=4).grid(
            row=0, column=0, sticky="ew", pady=(0, 2))
        self._cfg_firma = scrolledtext.ScrolledText(
            self._firma_frame, font=FONT_ENTRY, bg="white", fg="#222222",
            relief="sunken", bd=1, wrap="word")
        self._cfg_firma.grid(row=1, column=0, sticky="nsew")
        self._cfg_firma.bind("<Key>", self._mark_config_dirty)

        self._cfg_vars["dir_prueba"].trace_add("write", self._update_firma_visibility)

    def _update_firma_visibility(self, *_) -> None:
        if not hasattr(self, "_cfg_vars") or not hasattr(self, "_cfg_main"):
            return
        dir_p = self._cfg_vars.get("dir_prueba", tk.StringVar()).get()
        show  = "dugarry" in dir_p.lower()
        main  = self._cfg_main
        main.rowconfigure(3, weight=0)
        main.rowconfigure(4, weight=0)
        if show:
            self._firma_sep.grid(row=3, column=0, columnspan=2, sticky="ew", pady=(8, 4))
            if self._firma_shown:
                self._firma_frame.grid(row=4, column=0, columnspan=2, sticky="nsew")
                self.after(30, self._adjust_firma_height)
        else:
            self._firma_sep.grid_remove()
            self._firma_frame.grid_remove()

    def _toggle_firma(self) -> None:
        self._firma_shown = not self._firma_shown
        if self._firma_shown:
            self._firma_frame.grid(row=4, column=0, columnspan=2, sticky="nsew")
            self._btn_firma_toggle.config(text="▲ Ocultar Firma")
            self.after(30, self._adjust_firma_height)
        else:
            self._firma_frame.grid_remove()
            self._btn_firma_toggle.config(text="▼ Visualizar Firma")

    def _adjust_firma_height(self) -> None:
        if not hasattr(self, "_cfg_firma") or not hasattr(self, "_config_panel"):
            return
        self.update_idletasks()
        self._cfg_firma.configure(height=15)

    def _mandar_muestra(self) -> None:
        """Envía un correo de prueba al correo de testeo sin adjuntos."""
        cfg = self._get_config()
        # Sobreescribir con los valores actuales del panel (aunque no estén guardados)
        if hasattr(self, "_cfg_vars"):
            for k in ("font_family", "font_size", "asunto",
                      "saludo_es", "saludo_va", "despedida_es", "despedida_va"):
                v = self._cfg_vars.get(k, tk.StringVar()).get()
                if v:
                    cfg[k] = v
        if hasattr(self, "_cfg_cuerpo_es"):
            cfg["cuerpo_es"] = self._cfg_cuerpo_es.get("1.0", "end-1c")
            cfg["cuerpo_va"] = self._cfg_cuerpo_va.get("1.0", "end-1c")
            cfg["mail_firm"] = self._cfg_firma.get("1.0", "end-1c") or _FIRMA_DEFAULT
        dest = cfg.get("dir_prueba", "") or _DIR_PRUEBA_DEFAULT
        if not cfg.get("mail_cta") or not cfg.get("mail_clau"):
            messagebox.showerror("Configuración incompleta",
                                 "Configura la cuenta y contraseña antes de enviar.",
                                 parent=self); return
        self._btn_muestra.config(state="disabled")

        # Flash "Enviando Mail..." — mismo estilo y posición que _flash_message
        win_env = tk.Toplevel(self._btn_muestra)
        win_env.wm_overrideredirect(True)
        win_env.configure(bg="white",
                          highlightbackground=BG_CARPETA, highlightthickness=1)
        tk.Label(win_env, text="  Enviando Mail...  ", bg="white", fg=BG_CARPETA,
                 relief="flat", bd=0, font=FONT_UI, padx=10, pady=5).pack()
        win_env.update_idletasks()
        _tw = win_env.winfo_reqwidth()
        _th = win_env.winfo_reqheight()
        _bx = self._btn_muestra.winfo_rootx()
        _by = self._btn_muestra.winfo_rooty() + self._btn_muestra.winfo_height() + 4
        _ml, _mt, _mr, _mb = _monitor_rect(_bx, _by)
        if _bx + _tw > _mr - 4: _bx = _mr - _tw - 4
        if _by + _th > _mb - 4: _by = self._btn_muestra.winfo_rooty() - _th - 4
        _bx = max(_ml + 4, _bx)
        _by = max(_mt + 4, _by)
        win_env.wm_geometry(f"+{_bx}+{_by}")

        def _send():
            try:
                import ssl, smtplib
                from email.mime.multipart import MIMEMultipart
                from email.mime.text import MIMEText
                ctx = ssl.create_default_context()
                with smtplib.SMTP_SSL(_SMTP_SERVER, _SMTP_PORT, context=ctx) as srv:
                    srv.login(cfg["mail_cta"], cfg["mail_clau"])
                    msg = MIMEMultipart()
                    msg["From"]    = cfg.get("mail_from") or cfg["mail_cta"]
                    msg["To"]      = dest
                    msg["Subject"] = "[MUESTRA] " + cfg.get("asunto", "")
                    html = self._build_body_html(cfg)
                    msg.attach(MIMEText(html, "html", "utf-8"))
                    srv.sendmail(cfg["mail_cta"], dest, msg.as_string())
                self.after(0, lambda: (
                    win_env.destroy(),
                    self._log_write(f"  ✔  Muestra enviada a {dest}\n", "ok"),
                    self._btn_muestra.config(state="normal"),
                    self._flash_message(self._btn_muestra, "¡ Mail enviado !")))
            except Exception as exc:
                self.after(0, lambda: (
                    win_env.destroy(),
                    messagebox.showerror("Error al enviar", str(exc), parent=self),
                    self._btn_muestra.config(state="normal")))

        threading.Thread(target=_send, daemon=True).start()

    def _flash_message(self, widget: tk.Widget, text: str, ms: int = 3000) -> None:
        """Muestra un mensaje temporal bajo 'widget' con el mismo estilo que los tooltips."""
        win = tk.Toplevel(widget)
        win.wm_overrideredirect(True)
        win.configure(bg="white",
                      highlightbackground=BG_CARPETA, highlightthickness=1)
        tk.Label(win, text=text, bg="white", fg=BG_CARPETA,
                 relief="flat", bd=0, font=FONT_UI, padx=10, pady=5).pack()
        win.update_idletasks()
        tip_w = win.winfo_reqwidth()
        tip_h = win.winfo_reqheight()
        x     = widget.winfo_rootx()
        y     = widget.winfo_rooty() + widget.winfo_height() + 4
        ml, mt, mr, mb = _monitor_rect(x, y)
        if x + tip_w > mr - 4:
            x = mr - tip_w - 4
        if y + tip_h > mb - 4:
            y = widget.winfo_rooty() - tip_h - 4
        x = max(ml + 4, x)
        y = max(mt + 4, y)
        win.wm_geometry(f"+{x}+{y}")
        win.after(ms, win.destroy)

    def _translate_field_to_va(self, src_key: str, tgt_key: str) -> None:
        if not _HAS_DEEP_TRANSLATOR:
            messagebox.showwarning("deep-translator no disponible",
                                   "pip install deep-translator", parent=self); return
        text = self._cfg_vars.get(src_key, tk.StringVar()).get().strip()
        if not text:
            return
        try:
            translated = GoogleTranslator(source="es", target="ca").translate(text)
            self._cfg_vars[tgt_key].set(translated)
            self._mark_config_dirty()
        except Exception as exc:
            messagebox.showerror("Error de traducción", str(exc), parent=self)

    def _translate_google(self) -> None:
        if not _HAS_DEEP_TRANSLATOR:
            messagebox.showwarning(
                "deep-translator no disponible",
                "Instala la librería:\n  pip install deep-translator",
                parent=self)
            return
        text = self._cfg_cuerpo_es.get("1.0", "end-1c").strip()
        if not text:
            messagebox.showinfo("Sin texto",
                                "El cuerpo en español está vacío.", parent=self)
            return
        try:
            translated = GoogleTranslator(source="es", target="ca").translate(text)
            self._cfg_cuerpo_va.delete("1.0", "end")
            self._cfg_cuerpo_va.insert("1.0", translated)
            self._mark_config_dirty()
        except Exception as exc:
            messagebox.showerror("Error de traducción", str(exc), parent=self)

    def _lang_label(self) -> str:
        return ("⇄  Español / Valencià" if self._lang_order == "es_va"
                else "⇄  Valencià / Español")

    def _lang_style(self) -> str:
        return "LangES.TButton" if self._lang_order == "es_va" else "LangVA.TButton"

    def _toggle_lang_order(self) -> None:
        self._lang_order = "va_es" if self._lang_order == "es_va" else "es_va"
        self._btn_lang_order.config(text=self._lang_label(), style=self._lang_style())
        self._mark_config_dirty()

    def _show_preview(self) -> None:
        if not hasattr(self, "_preview_panel"):
            self._build_preview_panel()
        cfg = self._get_config()
        cfg["cuerpo_es"]  = self._cfg_cuerpo_es.get("1.0", "end-1c")
        cfg["cuerpo_va"]  = self._cfg_cuerpo_va.get("1.0", "end-1c")
        cfg["rubrica_es"] = self._cfg_vars.get("rubrica_es", tk.StringVar()).get()
        cfg["rubrica_va"] = self._cfg_vars.get("rubrica_va", tk.StringVar()).get()
        cfg["mail_firm"]  = self._cfg_firma.get("1.0", "end-1c") or _FIRMA_DEFAULT

        if self._lang_order == "es_va":
            pri, sec = "es", "va"
        else:
            pri, sec = "va", "es"

        t = self._preview_txt
        t.configure(state="normal")
        t.delete("1.0", "end")

        mail_from = cfg.get("mail_from") or cfg.get("mail_cta", "")
        t.insert("end", f"De:      {mail_from}\n", "meta")
        t.insert("end", f"Asunto:  {cfg.get('asunto','')}\n", "meta")
        t.insert("end", "─" * 64 + "\n\n", "sep")

        def _sec_block(lang):
            sal = cfg.get(f"saludo_{lang}", "")
            bod = cfg.get(f"cuerpo_{lang}", "")
            dep = cfg.get(f"despedida_{lang}", "")
            if sal:
                t.insert("end", sal + "\n\n", "rubrica")
            if bod:
                t.insert("end", bod + "\n", "body")
            if dep:
                t.insert("end", "\n" + dep + "\n", "rubrica")
            t.insert("end", "\n")

        # Sección primaria (idioma-1)
        _sec_block(pri)
        t.insert("end", "─" * 64 + "\n\n", "sep")

        # Sección secundaria (idioma-2)
        _sec_block(sec)
        t.insert("end", "─" * 64 + "\n", "sep")

        t.configure(state="disabled")
        self._preview_panel.place(x=0, y=0, relwidth=1, relheight=1)
        self._preview_panel.lift()

    def _build_preview_panel(self) -> None:
        pan = tk.Frame(self, bg="white", bd=1, relief="raised")
        self._preview_panel = pan

        hdr = tk.Frame(pan, bg=BG_TABLA)
        hdr.pack(fill="x")
        tk.Label(hdr, text="  👁  Previsualización del Correo",
                 bg=BG_TABLA, fg="white", font=FONT_TITLE,
                 anchor="w", pady=4).pack(side="left", fill="x", expand=True)
        tk.Button(hdr, text="  ✕  ", command=self._hide_preview,
                  bg=BG_TABLA, fg="white", activebackground="#C0392B",
                  activeforeground="white", relief="flat", bd=0,
                  font=FONT_BOLD, cursor="hand2").pack(side="right", padx=4)

        self._preview_txt = scrolledtext.ScrolledText(
            pan, font=("Georgia", BASE), wrap="word",
            bg="white", fg="#222222", padx=30, pady=20,
            relief="flat", bd=0, state="disabled",
            spacing1=3, spacing3=3)
        self._preview_txt.pack(fill="both", expand=True)

        self._preview_txt.tag_configure("meta",         font=("Courier New", BASE - 1),
                                        foreground="#666666")
        self._preview_txt.tag_configure("sep",          font=("Courier New", BASE - 2),
                                        foreground="#BBBBBB")
        self._preview_txt.tag_configure("link_sec",     font=("Verdana", BASE - 1),
                                        foreground=BG_CARPETA, justify="right",
                                        underline=True)
        self._preview_txt.tag_configure("link_pri",     font=("Verdana", BASE - 1),
                                        foreground=BG_COMUNES, justify="right",
                                        underline=True)
        self._preview_txt.tag_configure("rubrica",      font=("Georgia", BASE, "bold"))
        self._preview_txt.tag_configure("body",         font=("Georgia", BASE))
        self._preview_txt.tag_configure("firma_content",font=("Verdana", BASE - 2),
                                        foreground="#555555")

    def _hide_preview(self) -> None:
        if hasattr(self, "_preview_panel"):
            self._preview_panel.place_forget()

    def _mark_config_dirty(self, *_) -> None:
        self._config_dirty = True
        self._update_guardar_btn()

    def _update_guardar_btn(self) -> None:
        if not hasattr(self, "_btn_guardar"):
            return
        if self._config_dirty:
            self._btn_guardar.config(style="SaveDirty.TButton", text="¡Guardar!")
            self._btn_cancelar.grid()
        else:
            self._btn_guardar.config(style="SaveClean.TButton", text="¡Guardado!")
            self._btn_cancelar.grid_remove()

    def _revert_config(self) -> None:
        """Restaura los valores de la última salvaguardia sin cerrar el panel."""
        for name, var in self._cfg_vars.items():
            saved = _cfg_read(name)
            if name == "dir_prueba" and not saved:
                saved = _DIR_PRUEBA_DEFAULT
            if name in ("font_family",) and not saved:
                saved = "Verdana"
            if name == "font_size" and not saved:
                saved = "13"
            var.set(saved)
        self._cfg_cuerpo_es.delete("1.0", "end")
        self._cfg_cuerpo_es.insert("1.0", _cfg_read("cuerpo_es"))
        self._cfg_cuerpo_va.delete("1.0", "end")
        self._cfg_cuerpo_va.insert("1.0", _cfg_read("cuerpo_va"))
        saved_firma = _cfg_read("mail_firm")
        self._cfg_firma.delete("1.0", "end")
        self._cfg_firma.insert("1.0", saved_firma if saved_firma else _FIRMA_DEFAULT)
        saved_order = _cfg_read("lang_order")
        if saved_order in ("es_va", "va_es"):
            self._lang_order = saved_order
        self._btn_lang_order.config(text=self._lang_label(), style=self._lang_style())
        self._config_dirty = False
        self._update_guardar_btn()

    def _get_config(self) -> dict:
        cfg = {k: _cfg_read(k) for k in _CFG_KEYS}
        if not cfg["dir_prueba"]:
            cfg["dir_prueba"] = _DIR_PRUEBA_DEFAULT
        if not cfg["mail_from"]:
            cfg["mail_from"] = cfg["mail_cta"]
        if not cfg["mail_firm"]:
            cfg["mail_firm"] = _FIRMA_DEFAULT
        return cfg

    def _on_close(self) -> None:
        if self._config_dirty:
            if not messagebox.askyesno(
                    "Cambios sin guardar",
                    "La configuración tiene cambios sin guardar.\n"
                    "¿Cerrar sin guardar?",
                    parent=self):
                return
        self.destroy()

    # ── Markdown renderer ─────────────────────────────────────────────────────

    def _render_md(self, txt: scrolledtext.ScrolledText, content: str) -> None:
        txt.tag_configure("h1",       font=("Verdana", BASE + 4, "bold"),
                          foreground=BG_TABLA,   spacing1=12, spacing3=6)
        txt.tag_configure("h2",       font=("Verdana", BASE + 2, "bold"),
                          foreground=BG_CARPETA, spacing1=10, spacing3=4)
        txt.tag_configure("h3",       font=("Verdana", BASE, "bold"),
                          foreground=BG_COMUNES, spacing1=8,  spacing3=2)
        txt.tag_configure("code",     font=("Courier New", BASE - 1),
                          background="#E8EDF2", foreground="#2C3E50",
                          lmargin1=20, lmargin2=20, spacing1=1, spacing3=1)
        txt.tag_configure("table",    font=("Courier New", BASE - 1),
                          background="#F4F6F7", foreground="#2C3E50",
                          lmargin1=10, lmargin2=10)
        txt.tag_configure("tabsep",   font=("Courier New", BASE - 1),
                          background="#D5D8DC", foreground="#7F8C8D",
                          lmargin1=10, lmargin2=10)
        txt.tag_configure("bullet",   font=FONT_UI, foreground="#2C3E50",
                          lmargin1=24, lmargin2=40)
        txt.tag_configure("rule",     font=("Verdana", 4),
                          foreground=BG_TABLA, spacing1=6, spacing3=6)
        txt.tag_configure("quote",    font=("Verdana", BASE, "italic"),
                          foreground="#5D6D7E", lmargin1=30, lmargin2=30)
        txt.tag_configure("normal",   font=FONT_UI, foreground="#222222")
        txt.tag_configure("md_bold",  font=FONT_BOLD)
        txt.tag_configure("md_icode", font=("Courier New", BASE - 1), foreground="#C0392B")
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
                txt.insert("end", line + "\n", "code"); continue
            if stripped in ("---", "***", "___"):
                txt.insert("end", "─" * 72 + "\n", "rule"); continue
            if line.startswith("### "):
                _insert_inline(txt, line[4:] + "\n", "h3"); continue
            if line.startswith("## "):
                _insert_inline(txt, line[3:] + "\n", "h2"); continue
            if line.startswith("# "):
                _insert_inline(txt, line[2:] + "\n", "h1"); continue
            if line.startswith("|"):
                tag = "tabsep" if all(c in "-|: " for c in line) else "table"
                txt.insert("end", line + "\n", tag); continue
            if line.startswith("> "):
                _insert_inline(txt, line[2:] + "\n", "quote"); continue
            if re.match(r"^( {0,4})-\s", line):
                _insert_inline(txt, "  •  " + line.lstrip("- ").lstrip() + "\n", "bullet"); continue
            if stripped:
                _insert_inline(txt, line + "\n", "normal")
            else:
                txt.insert("end", "\n", "normal")
        txt.configure(state="disabled")

    # ── Selección de carpetas ─────────────────────────────────────────────────

    def _pick_adj_pers(self) -> None:
        init = (self._adj_pers_var.get()
                if self._adj_pers_var.get() and os.path.isdir(self._adj_pers_var.get())
                else os.path.expanduser("~"))
        folder = filedialog.askdirectory(
            parent=self, title="Carpeta de adjuntos personalizados", initialdir=init)
        if folder:
            self._adj_pers_var.set(folder)
            _reg_write(_REG_ADJ_PERS, folder)

    def _pick_adj_com(self) -> None:
        init = (self._adj_com_var.get()
                if self._adj_com_var.get() and os.path.isdir(self._adj_com_var.get())
                else os.path.expanduser("~"))
        folder = filedialog.askdirectory(
            parent=self, title="Carpeta de adjuntos comunes", initialdir=init)
        if folder:
            self._adj_com_var.set(folder)
            _reg_write(_REG_ADJ_COM, folder)

    # ── Debounce → carga ──────────────────────────────────────────────────────

    def _on_adj_pers_change(self, *_) -> None:
        if self._after_id_pers is not None:
            self.after_cancel(self._after_id_pers)
        self._after_id_pers = self.after(300, self._load_adj_pers)

    def _on_adj_com_change(self, *_) -> None:
        if self._after_id_com is not None:
            self.after_cancel(self._after_id_com)
        self._after_id_com = self.after(300, self._load_adj_com)

    # ── Carga tabla 1 (xlsx personalizados + email en hilo) ───────────────────

    def _load_adj_pers(self) -> None:
        self._after_id_pers = None
        carpeta = self._adj_pers_var.get().strip()
        self._pers_gen += 1
        gen = self._pers_gen
        self._clear_tree(self._tree1, self._lbl_count1)
        if not carpeta or not os.path.isdir(carpeta):
            return
        try:
            files = sorted(
                f for f in os.listdir(carpeta)
                if os.path.isfile(os.path.join(carpeta, f)) and f.lower().endswith(".xlsx"))
        except PermissionError as exc:
            self._log_write(f"  ERROR: {exc}\n", "error"); return

        rows = []
        for f in files:
            fp     = os.path.join(carpeta, f)
            fmodif = datetime.fromtimestamp(os.path.getmtime(fp)).strftime("%d/%m/%y")
            fsize  = _fmt_size(os.path.getsize(fp))
            rows.append(("", f, "", "", fmodif, fsize))

        self._populate_tree(self._tree1, _COL1_IDS, _COL1_NAMES,
                            self._sort_rev1, self._lbl_count1, rows)
        self._log_write(f"{len(rows)} fichero(s) Excel en  {carpeta}\n", "info")

        if _HAS_OPENPYXL and rows:
            self._log_write("  Leyendo emails de los ficheros…\n", "info")
            threading.Thread(
                target=self._read_emails_thread,
                args=(carpeta, files, gen),
                daemon=True).start()
        elif not _HAS_OPENPYXL:
            self._log_write("  ⚠  openpyxl no disponible — columna Correo vacía.\n", "warn")

    def _read_emails_thread(self, carpeta: str, files: list, gen: int) -> None:
        for i, fname in enumerate(files):
            if gen != self._pers_gen:
                return
            email = ""; obs = ""
            try:
                wb = openpyxl.load_workbook(
                    os.path.join(carpeta, fname), read_only=True, data_only=True)
                sheet_name = next(
                    (s for s in wb.sheetnames if s.lower() == "correo"), None)
                if not sheet_name:
                    obs = "Sin Hoja Correo"
                else:
                    val = wb[sheet_name]["A1"].value
                    if not val or not str(val).strip():
                        obs = "Sin Correo"
                    else:
                        email = str(val).strip()
                        if not _valid_email(email):
                            obs = "Correo Erróneo"
                wb.close()
            except Exception:
                obs = "Sin Hoja Correo"
            self.after(0, self._update_email_cell, str(i), email, obs, gen)
        self.after(0, self._log_write, "  ✔  Emails cargados.\n", "ok")
        self.after(0, lambda: self._autosize_columns(self._tree1, _COL1_IDS))
        self.after(0, lambda: self._sort_col(
            self._tree1, self._sort_rev1, _COL1_IDS, self._lbl_count1, "correo"))

    def _update_email_cell(self, iid: str, email: str, obs: str, gen: int) -> None:
        if gen != self._pers_gen:
            return
        try:
            self._tree1.set(iid, "correo",      email)
            self._tree1.set(iid, "observacion", obs)
        except Exception:
            pass

    # ── Modificación de correos en tabla 1 ───────────────────────────────────

    def _toggle_modif_correos(self) -> None:
        self._modif_correos_activo = not self._modif_correos_activo
        if self._modif_correos_activo:
            self._btn_modif_correos.config(
                text="Activado Modif. Correos", style="ModifOn.TButton")
            self._tree1.configure(selectmode="none")
            self._sel_none(self._tree1)
            self._log_write(
                "  ✎  Modo edición de correos activado — "
                "clic en una fila para editar su correo.\n", "warn")
            self.after(50, self._tip_modif_correos._show)
        else:
            self._btn_modif_correos.config(
                text="Activar Modif. Correos", style="Sel.TButton")
            self._tree1.configure(selectmode="extended")
            self._log_write("  Modo edición de correos desactivado.\n", "info")

    def _on_tree1_click(self, event) -> None:
        if not self._modif_correos_activo:
            return
        iid = self._tree1.identify_row(event.y)
        if not iid:
            return
        self._open_edit_correo(iid)

    def _open_edit_correo(self, iid: str) -> None:
        if not _HAS_OPENPYXL:
            messagebox.showwarning("openpyxl no disponible",
                                   "pip install openpyxl", parent=self)
            return
        try:
            vals = self._tree1.item(iid, "values")
            if len(vals) < 2:
                return
            fich          = str(vals[1]).strip()
            current_email = str(vals[0]).strip()
        except Exception:
            return

        carpeta  = self._adj_pers_var.get().strip()
        filepath = os.path.join(carpeta, fich)

        dlg = tk.Toplevel(self)
        try:
            dlg.title("Modificar Correo")
            dlg.resizable(False, False)
            dlg.configure(bg=BG_APP)
            dlg.transient(self)
            dlg.withdraw()

            # ── Cabecera ──────────────────────────────────────────────────
            hdr = tk.Frame(dlg, bg=BG_TABLA)
            hdr.pack(fill="x")
            tk.Label(hdr, text="  ✎  Modificar Correo",
                     bg=BG_TABLA, fg="white", font=FONT_TITLE,
                     anchor="w", pady=4).pack(side="left")

            # ── Cuerpo ────────────────────────────────────────────────────
            body = tk.Frame(dlg, bg=BG_APP, padx=20, pady=14)
            body.pack(fill="x")
            body.columnconfigure(1, weight=1)

            tk.Label(body, text="Fichero:", bg=BG_APP,
                     font=FONT_UI, anchor="e").grid(
                row=0, column=0, sticky="e", padx=(0, 10), pady=(0, 6))
            tk.Label(body, text=fich, bg=BG_APP, font=FONT_ENTRY,
                     fg=BG_CARPETA, anchor="w").grid(
                row=0, column=1, sticky="w", pady=(0, 6))

            tk.Label(body, text="Correo:", bg=BG_APP,
                     font=FONT_UI, anchor="e").grid(
                row=1, column=0, sticky="e", padx=(0, 10))
            var   = tk.StringVar(value=current_email)
            entry = tk.Entry(body, textvariable=var, font=FONT_UI,
                             bg="white", fg="#222222", relief="sunken",
                             bd=1, width=38)
            entry.grid(row=1, column=1, sticky="ew", ipady=4)

            lbl_err = tk.Label(body, text="", bg=BG_APP,
                               fg="#E74C3C", font=FONT_ENTRY)
            lbl_err.grid(row=2, column=1, sticky="w", pady=(3, 0))

            # ── Botones ───────────────────────────────────────────────────
            btn_bar = tk.Frame(dlg, bg=BG_APP, padx=20)
            btn_bar.pack(fill="x", pady=(0, 14))

            def _save() -> None:
                email = var.get().strip().lower()
                if email and not _valid_email(email):
                    lbl_err.config(text="Dirección no válida.")
                    entry.focus_set()
                    return
                try:
                    sheet_nueva = self._write_correo_xlsx(filepath, email)
                except PermissionError:
                    lbl_err.config(text="Fichero en uso — ciérralo en Excel.")
                    return
                except Exception as exc:
                    lbl_err.config(text=f"Error: {exc}")
                    return
                obs = "" if email else "Sin Correo"
                self._tree1.set(iid, "correo",      email)
                self._tree1.set(iid, "observacion", obs)
                self._sort_rev1["correo"] = False
                self.after(0, lambda: self._autosize_columns(self._tree1, _COL1_IDS))
                self.after(0, lambda: self._sort_col(
                    self._tree1, self._sort_rev1, _COL1_IDS,
                    self._lbl_count1, "correo"))
                accion = f"→ {email}" if email else "correo eliminado"
                self._log_write(f"  ✎  {fich}: {accion}\n", "ok")
                if not email:
                    msg = "Correo Borrado"
                elif sheet_nueva:
                    msg = "Correo Añadido en Hoja Creada 'Correo'"
                elif current_email:
                    msg = "Correo Modificado"
                else:
                    msg = "Correo Añadido"
                dlg.destroy()
                self._flash_correo(msg)

            ttk.Button(btn_bar, text="Guardar",  style="Export.TButton",
                       command=_save).pack(side="left", padx=(0, 8), ipadx=10)
            ttk.Button(btn_bar, text="Cancelar", style="Accion.TButton",
                       command=dlg.destroy).pack(side="left", ipadx=10)

            entry.bind("<Return>", lambda _: _save())
            dlg.bind("<Escape>",   lambda _: dlg.destroy())

            dlg.update_idletasks()
            w  = dlg.winfo_reqwidth()
            h  = dlg.winfo_reqheight()
            px = self.winfo_rootx() + self.winfo_width()  // 2
            py = self.winfo_rooty() + self.winfo_height() // 2
            x  = px - w // 2
            y  = py - h // 2
            ml, mt, mr, mb = _monitor_rect(px, py)
            x = max(ml + 8, min(x, mr - w - 8))
            y = max(mt + 8, min(y, mb - h - 8))
            dlg.geometry(f"+{x}+{y}")
            dlg.deiconify()
            dlg.grab_set()
            entry.select_range(0, "end")
            entry.focus_set()

        except Exception as exc:
            try:
                dlg.destroy()
            except Exception:
                pass
            messagebox.showerror("Error al abrir el editor de correo",
                                 str(exc), parent=self)

    @staticmethod
    def _write_correo_xlsx(filepath: str, email: str) -> bool:
        """Escribe el email en A1 de la hoja Correo. Devuelve True si la hoja fue creada."""
        wb = openpyxl.load_workbook(filepath)
        sheet_name  = next(
            (s for s in wb.sheetnames if s.lower() == "correo"), None)
        sheet_nueva = sheet_name is None
        ws = wb.create_sheet("Correo") if sheet_nueva else wb[sheet_name]
        ws["A1"] = email if email else None
        ws.sheet_state = "hidden"
        wb.save(filepath)
        wb.close()
        return sheet_nueva

    def _flash_correo(self, text: str) -> None:
        """Mensaje de confirmación como overlay hijo de self — siempre en el mismo monitor."""
        flash = tk.Frame(self, bg="white",
                         highlightbackground=BG_CARPETA, highlightthickness=1)
        tk.Label(flash, text=f"  {text}  ", bg="white", fg=BG_CARPETA,
                 font=FONT_UI, pady=8).pack()
        flash.place(relx=0.5, rely=0.5, anchor="center")
        flash.lift()
        flash.after(3000, flash.destroy)

    # ── Carga tabla 2 (comunes) ───────────────────────────────────────────────

    def _load_adj_com(self) -> None:
        self._after_id_com = None
        carpeta = self._adj_com_var.get().strip()
        self._clear_tree(self._tree2, self._lbl_count2)
        if not carpeta or not os.path.isdir(carpeta):
            return
        try:
            files = sorted(
                f for f in os.listdir(carpeta)
                if os.path.isfile(os.path.join(carpeta, f)))
        except PermissionError as exc:
            self._log_write(f"  ERROR: {exc}\n", "error"); return
        rows = []
        for f in files:
            fp     = os.path.join(carpeta, f)
            fmodif = datetime.fromtimestamp(os.path.getmtime(fp)).strftime("%d/%m/%y")
            fsize  = _fmt_size(os.path.getsize(fp))
            rows.append((f, fmodif, fsize))
        self._populate_tree(self._tree2, _COL2_IDS, _COL2_NAMES,
                            self._sort_rev2, self._lbl_count2, rows)
        self._log_write(f"{len(rows)} fichero(s) en  {carpeta}\n", "info")

    # ── Treeview helpers ──────────────────────────────────────────────────────

    def _apply_tree_tags(self, tree: ttk.Treeview) -> None:
        tree.tag_configure("even",     background="#FFFFFF", foreground="#222222")
        tree.tag_configure("odd",      background="#E8EDF2", foreground="#222222")
        tree.tag_configure("even_sel", background="#5B9BD5", foreground="white")
        tree.tag_configure("odd_sel",  background="#2E75B6", foreground="white")

    def _populate_tree(self, tree: ttk.Treeview, col_ids: tuple, col_names: tuple,
                       sort_rev: dict, lbl_count: tk.Label, rows: list) -> None:
        self._clear_tree(tree, lbl_count)
        sort_rev.update({c: False for c in col_ids})
        for col_id, col_name in zip(col_ids, col_names):
            anc = ("center" if col_id in ("fmodif", "observacion", "envio")
                   else "e" if col_id == "fsize" else "w")
            tree.heading(col_id, text=col_name, anchor=anc,
                         command=lambda c=col_id, t=tree, sr=sort_rev,
                                        ci=col_ids, lc=lbl_count:
                             self._sort_col(t, sr, ci, lc, c))
        for i, row in enumerate(rows):
            tag = "even" if i % 2 == 0 else "odd"
            tree.insert("", "end", iid=str(i), values=row, tags=(tag,))
        self._on_sel_change(tree, lbl_count)
        self.after(120, lambda t=tree, ci=col_ids: self._autosize_columns(t, ci))

    def _clear_tree(self, tree: ttk.Treeview, lbl_count: tk.Label) -> None:
        for iid in tree.get_children():
            tree.delete(iid)
        lbl_count.config(text="")

    def _sel_all(self, tree: ttk.Treeview) -> None:
        tree.selection_set(tree.get_children())

    def _sel_none(self, tree: ttk.Treeview) -> None:
        tree.selection_remove(tree.get_children())

    def _on_sel_change(self, tree: ttk.Treeview, lbl_count: tk.Label, _=None) -> None:
        selection = set(tree.selection())
        for iid in tree.get_children():
            even = tree.index(iid) % 2 == 0
            tag  = ("even_sel" if even else "odd_sel") if iid in selection \
                   else ("even" if even else "odd")
            tree.item(iid, tags=(tag,))
        sel = len(selection)
        tot = len(tree.get_children())
        lbl_count.config(text=f"{sel} / {tot}" if tot else "")

    def _sort_col(self, tree: ttk.Treeview, sort_rev: dict,
                  col_ids: tuple, lbl_count: tk.Label, col: str) -> None:
        items = [(tree.set(iid, col), iid) for iid in tree.get_children()]
        rev   = sort_rev.get(col, False)
        items.sort(key=lambda t: t[0].lower(), reverse=rev)
        for idx, (_, iid) in enumerate(items):
            tree.move(iid, "", idx)
        sort_rev[col] = not rev
        arrow = " ▲" if not rev else " ▼"
        for c in col_ids:
            raw = tree.heading(c, "text").rstrip(" ▲▼")
            tree.heading(c, text=raw + (arrow if c == col else ""))
        self._on_sel_change(tree, lbl_count)

    def _autosize_columns(self, tree: ttk.Treeview, col_ids: tuple) -> None:
        self.update_idletasks()
        font_n = tkfont.Font(family=FONT_UI[0],   size=FONT_UI[1])
        font_b = tkfont.Font(family=FONT_BOLD[0], size=FONT_BOLD[1], weight="bold")

        col_fmodif = max(font_b.measure("F. Modif."),
                         font_n.measure("DD/MM/YY")) + 15
        tree.column("fmodif", width=col_fmodif, minwidth=col_fmodif,
                    stretch=False, anchor="center")

        col_fsize = max(font_b.measure("Tam."),
                        font_n.measure("999 MB")) + 15
        tree.column("fsize", width=col_fsize, minwidth=col_fsize,
                    stretch=False, anchor="e")

        if "correo" in col_ids:
            col_correo = font_b.measure("Correo")
            col_fich   = font_b.measure("Fichero")
            for iid in tree.get_children():
                wc = font_n.measure(str(tree.set(iid, "correo")))
                wf = font_n.measure(str(tree.set(iid, "fichero")))
                if wc > col_correo: col_correo = wc
                if wf > col_fich:   col_fich   = wf
            tree.column("correo",  width=col_correo + 15, minwidth=80,
                        stretch=False, anchor="w")
            tree.column("fichero", width=col_fich   + 15, minwidth=80,
                        stretch=False, anchor="w")
            # Observación: fijo al máximo de los valores posibles
            col_obs = max(font_b.measure("Observación"),
                          *[font_n.measure(v) for v in
                            ("Sin Hoja Correo", "Sin Correo",  "Correo Erróneo")]) + 15
            tree.column("observacion", width=col_obs, minwidth=col_obs,
                        stretch=False, anchor="center")
            # Envío: fijo al máximo de los valores posibles
            col_env = max(font_b.measure("Envío"),
                          *[font_n.measure(v) for v in
                            ("Enviado 99º", "no seleccionado",
                             "Envío Fallido", "No Enviado")]) + 15
            tree.column("envio", width=col_env, minwidth=col_env,
                        stretch=False, anchor="center")
        else:
            col_fich = font_b.measure("Fichero")
            for iid in tree.get_children():
                w = font_n.measure(str(tree.set(iid, "fichero")))
                if w > col_fich: col_fich = w
            tree.column("fichero", width=col_fich + 15, minwidth=80,
                        stretch=False, anchor="w")

    # ── Mailing ───────────────────────────────────────────────────────────────

    def _toggle_prueba(self) -> None:
        self._prueba_activa = not self._prueba_activa
        if self._prueba_activa:
            self._btn_prueba.config(text="Prueba Activada", style="PruebaOn.TButton")
            cfg = self._get_config()
            self._log_write(
                f"  ⚠  Modo PRUEBA activado — correos redirigidos a: "
                f"{cfg['dir_prueba']}\n", "warn")
        else:
            self._btn_prueba.config(text="Activar Prueba", style="Prueba.TButton")
            self._log_write("  Modo prueba desactivado.\n", "info")

    def _confirm_envio(self, modo: str, n_dest: int, n_comunes: int) -> bool:
        result = [False]

        dlg = tk.Toplevel(self)
        dlg.withdraw()
        dlg.title("Confirmar envío")
        dlg.resizable(False, False)
        dlg.configure(bg=BG_APP)
        dlg.transient(self)

        bg_hdr = C_PRUEBA_ON if modo == "PRUEBA" else BG_COMUNES

        hdr = tk.Frame(dlg, bg=bg_hdr)
        hdr.pack(fill="x")
        tk.Label(hdr, text=f"  MODO  {modo}  ",
                 bg=bg_hdr, fg="white",
                 font=("Verdana", BASE + 6, "bold"),
                 anchor="w", pady=10).pack(side="left")

        body = tk.Frame(dlg, bg=BG_APP, padx=34, pady=18)
        body.pack(fill="x")
        fn = ("Verdana", BASE + 2)
        tk.Label(body, text=f"Destinatarios únicos:   {n_dest}",
                 bg=BG_APP, font=fn, anchor="w").pack(anchor="w", pady=3)
        tk.Label(body, text=f"Adjuntos comunes:       {n_comunes}",
                 bg=BG_APP, font=fn, anchor="w").pack(anchor="w", pady=3)
        tk.Label(body, text="",
                 bg=BG_APP).pack()
        tk.Label(body, text="¿Iniciar el envío?",
                 bg=BG_APP,
                 font=("Verdana", BASE + 3, "bold"),
                 anchor="w").pack(anchor="w")

        btn_bar = tk.Frame(dlg, bg=BG_APP, padx=34)
        btn_bar.pack(fill="x", pady=(4, 20))

        def _yes() -> None:
            result[0] = True
            dlg.destroy()

        ttk.Button(btn_bar, text="Sí, enviar", style="Export.TButton",
                   command=_yes).pack(side="left", padx=(0, 12), ipadx=14)
        ttk.Button(btn_bar, text="Cancelar", style="Accion.TButton",
                   command=dlg.destroy).pack(side="left", ipadx=14)

        dlg.bind("<Escape>", lambda _: dlg.destroy())

        dlg.update_idletasks()
        w  = dlg.winfo_reqwidth()
        h  = dlg.winfo_reqheight()
        px = self.winfo_rootx() + self.winfo_width()  // 2
        py = self.winfo_rooty() + self.winfo_height() // 2
        x, y = px - w // 2, py - h // 2
        ml, mt, mr, mb = _monitor_rect(px, py)
        x = max(ml + 8, min(x, mr - w - 8))
        y = max(mt + 8, min(y, mb - h - 8))
        dlg.geometry(f"+{x}+{y}")
        dlg.deiconify()
        dlg.grab_set()
        self.wait_window(dlg)
        return result[0]

    def _ejecutar_mailing(self) -> None:
        cfg = self._get_config()

        if not cfg["mail_cta"]:
            messagebox.showerror("Configuración incompleta",
                                 "Configura la cuenta Gmail (SMTP) antes de enviar.",
                                 parent=self); return
        if not cfg["mail_clau"]:
            messagebox.showerror("Configuración incompleta",
                                 "Configura el App Password antes de enviar.",
                                 parent=self); return
        if not cfg["asunto"]:
            messagebox.showerror("Configuración incompleta",
                                 "El asunto del correo está vacío.",
                                 parent=self); return

        carpeta_pers = self._adj_pers_var.get().strip()
        carpeta_com  = self._adj_com_var.get().strip()
        selected     = set(self._tree1.selection())

        # Agrupar filas, asignar números de grupo y actualizar columna Envío
        groups:       dict[str, list[str]] = {}
        group_nums:   dict[str, int]       = {}
        iid_by_email: dict[str, list]      = {}
        no_enviado:   dict[str, int]       = {}
        g_counter = 0

        for iid in self._tree1.get_children():
            email = str(self._tree1.set(iid, "correo")).strip()
            fich  = str(self._tree1.set(iid, "fichero")).strip()
            obs   = str(self._tree1.set(iid, "observacion")).strip()

            if iid not in selected:
                self._tree1.set(iid, "envio", "no seleccionado")
                continue
            if obs in ("Sin Hoja Correo", "Sin Correo", "Correo Erróneo") or \
               not email or not _valid_email(email):
                self._tree1.set(iid, "envio", "No Enviado")
                if obs in ("Sin Hoja Correo", "Sin Correo", "Correo Erróneo"):
                    cat = obs
                elif not email:
                    cat = "Sin Correo"
                else:
                    cat = "Correo Erróneo"
                no_enviado[cat] = no_enviado.get(cat, 0) + 1
                continue
            if email not in group_nums:
                g_counter += 1
                group_nums[email] = g_counter
            iid_by_email.setdefault(email, []).append(iid)
            groups.setdefault(email, []).append(os.path.join(carpeta_pers, fich))

        if not groups:
            messagebox.showwarning(
                "Sin destinatarios",
                "No hay filas seleccionadas con email válido.",
                parent=self); return

        # Adjuntos comunes seleccionados
        comunes: list[str] = []
        for iid in self._tree2.selection():
            fich = str(self._tree2.item(iid, "values")[0]).strip()
            if fich and carpeta_com:
                comunes.append(os.path.join(carpeta_com, fich))

        modo   = "PRUEBA" if self._prueba_activa else "REAL"
        n_dest = len(groups)
        if not self._confirm_envio(modo, n_dest, len(comunes)):
            return

        self._btn_mailing.config(state="disabled")
        self._no_enviado_counts = no_enviado
        self._log_write(f"  Iniciando envío — {n_dest} destinatario(s)…\n", "info")
        self._show_sending_overlay()

        threading.Thread(
            target=self._send_thread,
            args=(cfg, groups, comunes, group_nums, iid_by_email),
            daemon=True).start()

    def _send_thread(self, cfg: dict, groups: dict, comunes: list,
                     group_nums: dict, iid_by_email: dict) -> None:
        ok = 0; errors = 0
        try:
            ctx = ssl.create_default_context()
            with smtplib.SMTP_SSL(_SMTP_SERVER, _SMTP_PORT, context=ctx) as server:
                server.login(cfg["mail_cta"], cfg["mail_clau"])

                for email, adj_pers in groups.items():
                    dest   = cfg["dir_prueba"] if self._prueba_activa else email
                    asunto = ("[PRUEBA] " if self._prueba_activa else "") + cfg["asunto"]
                    n      = group_nums.get(email, 0)
                    self.after(0, self._set_overlay_number, n)   # nº del correo a enviar
                    try:
                        msg = MIMEMultipart()
                        msg["From"]    = cfg["mail_from"]
                        msg["To"]      = dest
                        msg["Subject"] = asunto
                        msg.attach(MIMEText(self._build_body_html(cfg), "html", "utf-8"))
                        for path in adj_pers:
                            self._attach_file(msg, path)
                            self.after(0, self._move_sending_overlay)   # mover por cada fila personalizada leída
                        for path in comunes:
                            self._attach_file(msg, path)
                        server.sendmail(cfg["mail_cta"], dest, msg.as_string())
                        ok += 1
                        dest_txt = f"{email} → {dest}" if self._prueba_activa else email
                        self.after(0, self._log_write, f"  ✔  {dest_txt}\n", "ok")
                        for iid in iid_by_email.get(email, []):
                            self.after(0, self._tree1.set, iid, "envio",
                                       f"Enviado {n}º")
                    except Exception as exc:
                        errors += 1
                        self.after(0, self._log_write,
                                   f"  ✖  {email}: {exc}\n", "error")
                        for iid in iid_by_email.get(email, []):
                            self.after(0, self._tree1.set, iid, "envio", "Envío Fallido")

        except smtplib.SMTPAuthenticationError:
            self.after(0, self._log_write,
                       "\n  ✖  Error de autenticación — verifica cuenta y App Password.\n",
                       "error")
            self.after(0, self._send_done, ok, errors + 1)
            return
        except Exception as exc:
            self.after(0, self._log_write, f"\n  ✖  Error SMTP: {exc}\n", "error")
            self.after(0, self._send_done, ok, errors + 1)
            return

        self.after(0, self._send_done, ok, errors)

    def _show_sending_overlay(self) -> None:
        if not self._gif_frames:
            return
        if _cfg_read("animacion") == "0":   # animación del cartero desactivada en config
            return
        _TC = "#FF00FF"   # color clave para transparencia (no aparece en el GIF)
        ov = tk.Toplevel(self)
        ov.wm_overrideredirect(True)
        ov.configure(bg=_TC)
        ov.wm_attributes("-transparentcolor", _TC)
        ov.attributes("-topmost", True)
        self._sending_overlay = ov
        lbl = tk.Label(ov, image=self._gif_frames[0], bg=_TC)
        lbl.pack()
        self._gif_lbl = lbl
        ov.update_idletasks()
        w = ov.winfo_reqwidth(); h = ov.winfo_reqheight()
        x = self.winfo_rootx() + (self.winfo_width()  - w) // 2
        y = self.winfo_rooty() + (self.winfo_height() - h) // 2
        ov.wm_geometry(f"+{x}+{y}")
        self._gif_overlay_wh = (w, h)
        self._glide_pos    = (float(x), float(y))   # punto de partida del deslizamiento
        self._glide_target = (float(x), float(y))
        self._gliding      = False
        self.update_idletasks()   # render inicial antes de arrancar el hilo

        # Anima los frames solo si el GIF tiene más de uno (evita busy-loop a 0 ms).
        if len(self._gif_frames) > 1:
            def _animate(idx: int = 0) -> None:
                if self._sending_overlay is None:
                    return
                self._gif_lbl.configure(image=self._gif_frames[idx])
                self._gif_after_id = self.after(
                    self._gif_delay,
                    _animate, (idx + 1) % len(self._gif_frames))
            _animate()

    def _rand_overlay_target(self) -> tuple[int, int]:
        """Punto aleatorio dentro de la ventana respetando un margen (con clamp)."""
        margin = 80
        ow, oh = self._gif_overlay_wh
        min_x = self.winfo_rootx() + margin
        min_y = self.winfo_rooty() + margin
        # clamp: si la ventana es pequeña, garantiza min ≤ max (siempre hay destino)
        max_x = max(self.winfo_rootx() + self.winfo_width()  - ow - margin, min_x)
        max_y = max(self.winfo_rooty() + self.winfo_height() - oh - margin, min_y)
        return random.randint(min_x, max_x), random.randint(min_y, max_y)

    def _move_sending_overlay(self) -> None:
        """Fija un destino aleatorio y desliza suavemente el overlay hacia él.
        Se invoca en el hilo principal vía after(0) tras cada fila personalizada.
        Si ya se está deslizando, redirige al nuevo destino sin saltos."""
        if self._sending_overlay is None:
            return
        self._glide_target = self._rand_overlay_target()
        if not self._gliding:           # un único bucle persigue siempre el target actual
            self._gliding = True
            self._glide_loop()

    def _glide_loop(self) -> None:
        ov = self._sending_overlay
        if ov is None:
            self._gliding = False
            return
        px, py = self._glide_pos
        tx, ty = self._glide_target
        dx, dy = tx - px, ty - py
        if abs(dx) < 1 and abs(dy) < 1:                 # llegó: fija y detén el bucle
            self._glide_pos = (float(tx), float(ty))
            ov.wm_geometry(f"+{int(tx)}+{int(ty)}")
            ov.update_idletasks()
            self._gliding = False
            return
        px += dx * self._glide_ease                     # easing exponencial (desacelera)
        py += dy * self._glide_ease
        self._glide_pos = (px, py)
        ov.wm_geometry(f"+{int(round(px))}+{int(round(py))}")
        ov.update_idletasks()
        self._glide_after_id = self.after(self._glide_ms, self._glide_loop)

    def _set_overlay_number(self, n: int) -> None:
        """Dibuja el número del correo sobre el sobre del cartero. El número forma parte
        de la imagen (no es un widget), para no romper la transparencia del overlay."""
        lbl = self._gif_lbl
        if self._sending_overlay is None or self._gif_base is None or lbl is None:
            return
        try:
            from PIL import ImageDraw, ImageFont
            img  = self._gif_base.copy()
            draw = ImageDraw.Draw(img)
            text = str(n)
            try:
                font = ImageFont.truetype("arialbd.ttf", 56)
            except Exception:
                font = ImageFont.load_default()
            W, H = img.size
            cx, cy = int(W * 0.43) + 40, int(H * 0.58) - 25   # sobre el sobre naranja
            bbox = draw.textbbox((0, 0), text, font=font, stroke_width=3)
            tw, th = bbox[2] - bbox[0], bbox[3] - bbox[1]
            draw.text((cx - tw / 2 - bbox[0], cy - th / 2 - bbox[1]), text,
                      font=font, fill=(255, 255, 255, 255),
                      stroke_width=3, stroke_fill=(120, 55, 0, 255))
            photo = ImageTk.PhotoImage(img)
            self._gif_num_photo = photo            # ref viva para evitar recolección
            lbl.configure(image=photo)
        except Exception:
            pass

    def _hide_sending_overlay(self) -> None:
        for aid in (self._gif_after_id, self._glide_after_id):
            if aid:
                self.after_cancel(aid)
        self._gif_after_id = None
        self._glide_after_id = None
        self._gliding = False
        if self._sending_overlay:
            self._sending_overlay.destroy()
            self._sending_overlay = None

    def _send_done(self, ok: int, errors: int) -> None:
        self._hide_sending_overlay()
        self._btn_mailing.config(state="normal")
        suffix = f", {errors} error(es)" if errors else ""
        self._log_write(
            f"  Completado: {ok} correo(s) enviado(s){suffix}.\n",
            "warn" if errors else "ok")
        self._flash_message(self._btn_mailing, "¡ Mailing completado !")
        self._show_envio_done(ok, errors)

    def _show_envio_done(self, ok: int, errors: int) -> None:
        dlg = tk.Toplevel(self)
        dlg.withdraw()
        dlg.title("Envío completado")
        dlg.resizable(False, False)
        dlg.configure(bg=BG_APP)
        dlg.transient(self)

        bg_hdr = C_PRUEBA_ON if errors else BG_COMUNES
        titulo = "COMPLETADO CON ERRORES" if errors else "ENVÍO COMPLETADO"

        hdr = tk.Frame(dlg, bg=bg_hdr)
        hdr.pack(fill="x")
        tk.Label(hdr, text=f"  {titulo}  ",
                 bg=bg_hdr, fg="white",
                 font=("Verdana", BASE + 6, "bold"),
                 anchor="w", pady=10).pack(side="left")

        body = tk.Frame(dlg, bg=BG_APP, padx=34, pady=18)
        body.pack(fill="x")
        fn     = ("Verdana", BASE + 2)
        fn_sub = ("Verdana", BASE)
        _C_NO_ENVIADO = "#A04000"   # naranja oscuro, legible sobre fondo claro

        rows = tk.Frame(body, bg=BG_APP)
        rows.pack(fill="x")
        rows.columnconfigure(0, weight=1)   # empuja los números a la columna derecha
        _r = 0

        def _row(label, value, *, lab_font=fn, fg="#222222", indent=0, pady=3):
            nonlocal _r
            tk.Label(rows, text=label, bg=BG_APP, font=lab_font, fg=fg,
                     anchor="w").grid(row=_r, column=0, sticky="w",
                                      padx=(indent, 24), pady=pady)
            tk.Label(rows, text=str(value), bg=BG_APP, font=lab_font, fg=fg,
                     anchor="e").grid(row=_r, column=1, sticky="e", pady=pady)
            _r += 1

        _row("Correos enviados:", ok)
        no_enviado = getattr(self, "_no_enviado_counts", {})
        if no_enviado:
            _row("No enviados:", sum(no_enviado.values()),
                 fg=_C_NO_ENVIADO, pady=(18, 3))
            for cat in ("Sin Correo", "Sin Hoja Correo", "Correo Erróneo"):
                if no_enviado.get(cat):
                    _row(f"{cat}:", no_enviado[cat],
                         lab_font=fn_sub, indent=28, pady=1)
        if errors:
            _row("Errores:", errors, fg=C_PRUEBA_ON, pady=(18, 3))
            tk.Label(body, text="Ver Informe para el detalle.",
                     bg=BG_APP, font=("Verdana", BASE + 1, "bold"),
                     anchor="w").pack(anchor="w", pady=(12, 0))

        btn_bar = tk.Frame(dlg, bg=BG_APP, padx=34)
        btn_bar.pack(fill="x", pady=(4, 20))
        ttk.Button(btn_bar, text="Aceptar", style="Export.TButton",
                   command=dlg.destroy).pack(side="left", ipadx=14)

        dlg.bind("<Escape>", lambda _: dlg.destroy())
        dlg.bind("<Return>", lambda _: dlg.destroy())

        dlg.update_idletasks()
        w  = dlg.winfo_reqwidth()
        h  = dlg.winfo_reqheight()
        px = self.winfo_rootx() + self.winfo_width()  // 2
        py = self.winfo_rooty() + self.winfo_height() // 2
        x, y = px - w // 2, py - h // 2
        ml, mt, mr, mb = _monitor_rect(px, py)
        x = max(ml + 8, min(x, mr - w - 8))
        y = max(mt + 8, min(y, mb - h - 8))
        dlg.geometry(f"+{x}+{y}")
        dlg.deiconify()
        dlg.grab_set()
        self.wait_window(dlg)

    def _build_body_html(self, cfg: dict) -> str:
        # Firma con placeholders sustituidos
        firma = cfg.get("mail_firm", "") or _FIRMA_DEFAULT
        firma = firma.replace("ext. ____",                 f"ext. {cfg.get('user_ext','')}")
        firma = firma.replace("https://sc.ua.es/es/_____", cfg.get("web_es", ""))
        firma = firma.replace("https://sc.ua.es/va/_____", cfg.get("web_va", ""))
        firma = firma.replace("Servei _____",              cfg.get("servicio", ""))
        firma = firma.replace("Unitat _____",              cfg.get("unidad", ""))

        # Orden: primary = primero del toggle, secondary = segundo
        if self._lang_order == "es_va":
            pri, sec      = "es", "va"
            lbl_pri, lbl_sec = "Español", "Valencià"
        else:
            pri, sec      = "va", "es"
            lbl_pri, lbl_sec = "Valencià", "Español"

        font_f = cfg.get("font_family", "Verdana") or "Verdana"
        font_s = cfg.get("font_size", "13") or "13"
        fstyle = f"font-family:'{font_f}',sans-serif;font-size:{font_s}pt;"

        def to_html(text: str) -> str:
            lines = [l for l in text.split("\n") if l.strip()]
            return "".join(f'<p style="{fstyle}">{l}</p>' for l in lines) or "<p>&nbsp;</p>"

        saludo_pri    = cfg.get(f"saludo_{pri}", "")
        saludo_sec    = cfg.get(f"saludo_{sec}", "")
        body_pri      = to_html(cfg.get(f"cuerpo_{pri}", ""))
        body_sec      = to_html(cfg.get(f"cuerpo_{sec}", ""))
        despedida_pri = cfg.get(f"despedida_{pri}", "")
        despedida_sec = cfg.get(f"despedida_{sec}", "")

        # Cabecera de cada sección:
        # - sección pri: anchor id=lang-pri + link al sec (texto lbl_sec)
        # - sección sec: anchor id=lang-sec + link al pri (texto lbl_pri)
        def p(text: str) -> str:
            return f'<p style="{fstyle}">{text}</p>' if text.strip() else ""

        # La firma mantiene su propio font (Times New Roman) — wrapper explícito
        firma_block = (
            '<div style="font-family:\'Times New Roman\',serif;font-size:medium;">'
            + firma +
            '</div>'
        )

        # Estructura: idioma-1 (pri) primero, idioma-2 (sec) segundo
        return (
            "<html><body>"
            + (p(saludo_pri) + "<p>&nbsp;</p>" if saludo_pri else "")
            + f"{body_pri}"
            + ("<p>&nbsp;</p>" + p(despedida_pri) if despedida_pri else "")
            + "<hr>"
            + (p(saludo_sec) + "<p>&nbsp;</p>" if saludo_sec else "")
            + f"{body_sec}"
            + ("<p>&nbsp;</p>" + p(despedida_sec) if despedida_sec else "")
            + f"<p>&nbsp;</p>"
            + firma_block
            + "</body></html>"
        )

    @staticmethod
    def _attach_file(msg: MIMEMultipart, path: str) -> None:
        if not os.path.isfile(path):
            return
        with open(path, "rb") as f:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(f.read())
        encoders.encode_base64(part)
        # Tupla (charset, language, value) → RFC 2231/5987, soporta caracteres no-ASCII
        part.add_header("Content-Disposition", "attachment",
                        filename=("utf-8", "", os.path.basename(path)))
        msg.attach(part)

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


if __name__ == "__main__":
    App().mainloop()
