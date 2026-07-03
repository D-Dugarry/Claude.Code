"""
help_tooltips.py — Tooltips de ayuda con interruptor maestro (tkinter + Windows).

Patrón: una imagen/logo (el "caracolillo") actúa como CONTROL MAESTRO de la ayuda:
  · pasar el cursor  → muestra tu firma + un tooltip que explica el propio control,
  · clic izquierdo   → activa/desactiva TODOS los tooltips explicativos de los botones,
  · clic derecho     → una acción tuya (abrir un panel de ayuda, el flujo, etc.).

Trampas no obvias que esto encapsula (descubiertas a base de depurar):

  1. GATE GLOBAL: los tooltips de botones consultan un único flag de clase
     (`Tooltip.enabled`). Así un solo control los enciende/apaga todos, SIN tocar
     cada `Tooltip`. Las excepciones (firma, el propio control) se crean con
     `always=True` y se muestran pase lo que pase.

  2. COEXISTENCIA en un mismo widget: el caracolillo lleva DOS tooltips (firma +
     explicativo). `widget.bind(seq, fn)` sin más REEMPLAZA el handler anterior; hay
     que usar `add="+"` para que ambos (y cualquier bind tuyo) convivan.

  3. ESQUINAS REDONDEADAS: en Windows 11 se redondean con la API DWM
     (`DWMWA_WINDOW_CORNER_PREFERENCE`) sobre el HWND `GetParent(winfo_id())` — no
     sobre `winfo_id()` directamente (da handle inválido). Todo en try/except: en
     Windows anteriores simplemente no redondea.

  4. POSICIÓN multi-monitor: el tooltip se reubica dentro del monitor que contiene el
     punto, para no salirse de pantalla con varios monitores.

  5. REDIBUJADO tras el toggle: el clic en el control cambia el texto del tooltip
     explicativo (solo dos valores: hint_on/hint_off) al instante, sin esperar a un
     <Enter> que no llegará porque el ratón no se movió. No basta con un flag de
     <Enter>/<Leave>: destruir/crear Toplevels en rápida sucesión dispara <Leave>
     espurios, así que se comprueba la posición real del puntero antes de redibujar.

Uso típico:

    from help_tooltips import Tooltip, attach_help_toggle

    Tooltip(boton1, "Envía el formulario")          # sujeto al interruptor
    Tooltip(boton2, lambda: f"{n} pendientes")      # texto dinámico (callable)

    attach_help_toggle(logo,                         # tu Label/imagen
                       signature="Dugarry'26",
                       on_secondary=mostrar_ayuda)   # clic derecho

Requiere Windows para esquinas redondeadas y posición multi-monitor (degrada solo).
"""

import tkinter as tk

try:
    import ctypes
    import ctypes.wintypes
    _HAS_CTYPES = True
except Exception:
    _HAS_CTYPES = False


# ── helpers Win32 ──────────────────────────────────────────────────────────────

if _HAS_CTYPES:
    class _RECT(ctypes.Structure):
        _fields_ = [("left", ctypes.c_long), ("top", ctypes.c_long),
                    ("right", ctypes.c_long), ("bottom", ctypes.c_long)]

    class _MONITORINFO(ctypes.Structure):
        _fields_ = [("cbSize", ctypes.c_ulong), ("rcMonitor", _RECT),
                    ("rcWork", _RECT), ("dwFlags", ctypes.c_ulong)]


def monitor_rect(x: int, y: int) -> tuple:
    """(left, top, right, bottom) del área de trabajo del monitor que contiene (x, y)."""
    if _HAS_CTYPES:
        try:
            u32 = ctypes.windll.user32
            pt = ctypes.wintypes.POINT(x, y)
            hmon = u32.MonitorFromPoint(pt, ctypes.c_uint(2))   # DEFAULTTONEAREST
            mi = _MONITORINFO()
            mi.cbSize = ctypes.sizeof(_MONITORINFO)
            if u32.GetMonitorInfoW(hmon, ctypes.byref(mi)):
                r = mi.rcWork
                return r.left, r.top, r.right, r.bottom
        except Exception:
            pass
        try:
            sm = ctypes.windll.user32.GetSystemMetrics
            return 0, 0, sm(0), sm(1)
        except Exception:
            pass
    return 0, 0, 1920, 1080


def round_corners(win: tk.Misc, small: bool = True) -> None:
    """Redondea ligeramente las esquinas de una ventana sin borde (Windows 11 DWM)."""
    if not _HAS_CTYPES:
        return
    try:
        win.update_idletasks()
        u32 = ctypes.windll.user32
        u32.GetParent.restype = ctypes.wintypes.HWND
        u32.GetParent.argtypes = [ctypes.wintypes.HWND]
        hwnd = u32.GetParent(win.winfo_id())
        _DWMWA_WINDOW_CORNER_PREFERENCE = 33
        _ROUND, _ROUNDSMALL = 2, 3
        pref = ctypes.c_int(_ROUNDSMALL if small else _ROUND)
        dwm = ctypes.windll.dwmapi
        dwm.DwmSetWindowAttribute.argtypes = [
            ctypes.wintypes.HWND, ctypes.wintypes.DWORD,
            ctypes.c_void_p, ctypes.wintypes.DWORD]
        dwm.DwmSetWindowAttribute(hwnd, _DWMWA_WINDOW_CORNER_PREFERENCE,
                                  ctypes.byref(pref), ctypes.sizeof(pref))
    except Exception:
        pass


# ── Tooltip ─────────────────────────────────────────────────────────────────────

class Tooltip:
    """Tooltip de texto fijo o dinámico (callable), con interruptor maestro.

    enabled (atributo de CLASE) es el gate global: los tooltips con `always=False`
    solo se muestran cuando enabled es True. Los `always=True` se muestran siempre.
    """
    enabled = False

    def __init__(self, widget: tk.Widget, text, *,
                 always: bool = False,
                 side: str = "bottom",                  # bottom|top|left|right
                 bg: str = "white",
                 fg: str = "#1A5276",
                 border: str = "#1A5276",
                 border_width: int = 1,
                 font=("Segoe UI", 10),
                 padx: int = 10, pady: int = 5,
                 rounded: bool = True,
                 transparent: bool = False) -> None:
        """text: str o callable() -> str (evaluado al mostrarse).
        transparent: efecto "solo texto" (sin recuadro) usando -transparentcolor;
                     útil para una firma. Anula borde y redondeo."""
        self._w = widget
        self._text = text
        self._always = always
        self._side = side
        self._bg, self._fg = bg, fg
        self._border, self._bw = border, border_width
        self._font = font
        self._padx, self._pady = padx, pady
        self._rounded = rounded and not transparent
        self._transparent = transparent
        self._win = None
        # add="+" SIEMPRE: no machaca otros binds del widget (trampa #2)
        widget.bind("<Enter>", self._show, add="+")
        widget.bind("<Leave>", self._hide, add="+")
        widget.bind("<ButtonPress>", self._hide, add="+")

    def _show(self, _=None) -> None:
        if self._win is not None:
            return
        if not self._always and not Tooltip.enabled:
            return
        try:
            text = self._text() if callable(self._text) else self._text
        except Exception:
            return   # un callable que falle no debe romper el handler <Enter>
        if not text:
            return
        try:
            self._win = tk.Toplevel(self._w)
            self._win.wm_overrideredirect(True)
            try:
                self._win.attributes("-topmost", True)
            except tk.TclError:
                pass
            if self._transparent:
                self._win.configure(bg=self._bg)
                try:
                    self._win.wm_attributes("-transparentcolor", self._bg)
                except tk.TclError:
                    pass
                tk.Label(self._win, text=text, bg=self._bg, fg=self._fg,
                         font=self._font, padx=self._padx, pady=self._pady,
                         justify="left").pack()
            else:
                self._win.configure(bg=self._bg, highlightbackground=self._border,
                                    highlightthickness=self._bw)
                tk.Label(self._win, text=text, bg=self._bg, fg=self._fg,
                         font=self._font, padx=self._padx, pady=self._pady,
                         bd=0, justify="left").pack()
            self._place(self._win)
        except tk.TclError:
            self._hide()   # el widget ancla pudo destruirse mientras se mostraba
            return
        if self._rounded:
            round_corners(self._win)

    def _place(self, win: tk.Toplevel) -> None:
        win.update_idletasks()
        tw, th = win.winfo_reqwidth(), win.winfo_reqheight()
        wx, wy = self._w.winfo_rootx(), self._w.winfo_rooty()
        ww, wh = self._w.winfo_width(), self._w.winfo_height()
        if self._side == "top":
            x, y = wx, wy - th - 4
        elif self._side == "left":
            x, y = wx - tw - 6, wy + (wh - th) // 2
        elif self._side == "right":
            x, y = wx + ww + 6, wy + (wh - th) // 2
        else:  # bottom
            x, y = wx, wy + wh + 4
        ml, mt, mr, mb = monitor_rect(wx, wy)
        x = max(ml + 4, min(x, mr - tw - 4))
        y = max(mt + 4, min(y, mb - th - 4))
        win.wm_geometry(f"+{x}+{y}")

    def _hide(self, _=None) -> None:
        if self._win is not None:
            self._win.destroy()
            self._win = None


# ── interruptor maestro (gate global) ────────────────────────────────────────────

def toggle_tooltips() -> bool:
    """Alterna la visibilidad de los tooltips 'gated'. Devuelve el nuevo estado."""
    Tooltip.enabled = not Tooltip.enabled
    return Tooltip.enabled


def set_tooltips_enabled(value: bool) -> None:
    Tooltip.enabled = bool(value)


def tooltips_enabled() -> bool:
    return Tooltip.enabled


# ── control maestro (el "caracolillo") ───────────────────────────────────────────

def attach_help_toggle(widget: tk.Widget, *,
                       signature: str = None,
                       on_secondary=None,
                       hint_on: str  = "Ocultar textos explicativos de los botones",
                       hint_off: str = "Mostrar textos explicativos de los botones",
                       signature_style: dict = None) -> None:
    """Convierte `widget` (un Label/imagen) en el control maestro de la ayuda:
      · hover          → firma (si se da) + tooltip dinámico (always),
      · clic izquierdo → toggle del interruptor global; el tooltip cambia de
                         texto al instante (mientras el cursor siga encima),
      · clic derecho   → on_secondary() (si se da).
    """
    try:
        widget.configure(cursor="hand2")
    except tk.TclError:
        pass

    sig_tip = None
    if signature:
        sty = dict(always=True, side="left", transparent=True,
                   fg="#C8A96E", bg="#1A5276", font=("Mistral", 18))
        if signature_style:
            sty.update(signature_style)
        sig_tip = Tooltip(widget, signature, **sty)

    # tooltip explicativo del propio control: siempre visible, refleja el estado
    hint_tip = Tooltip(widget, lambda: hint_on if Tooltip.enabled else hint_off,
                       always=True, side="bottom")

    def _pointer_still_over() -> bool:
        # Mientras el ratón no se mueva no llega un nuevo <Enter>, así que tras el
        # toggle hay que redibujar el tooltip "a mano" para que muestre el texto
        # del estado ya actualizado. No nos fiamos de un flag puesto por
        # <Enter>/<Leave>: destruir y crear el Toplevel del tooltip en rápida
        # sucesión dispara <Leave> espurios en el widget ancla (confunde el
        # seguimiento de puntero de Tk), así que comprobamos la posición real.
        try:
            x, y = widget.winfo_pointerxy()
            return widget.winfo_containing(x, y) is widget
        except tk.TclError:
            return False

    def _reshow():
        if _pointer_still_over():
            if sig_tip is not None:
                sig_tip._show()
            hint_tip._show()

    def _do_toggle(_e):
        # Forzamos el cierre aquí mismo (sin depender del <ButtonPress> genérico
        # de Tooltip._hide: en un clic real no siempre llega junto a <Button-1>,
        # así que si no se cierra a mano el tooltip viejo queda abierto y
        # _show() no hace nada por tener ya ventana, manteniendo el texto stale).
        if sig_tip is not None:
            sig_tip._hide()
        hint_tip._hide()
        toggle_tooltips()
        _reshow()

    widget.bind("<Button-1>", _do_toggle, add="+")
    if on_secondary is not None:
        widget.bind("<Button-3>", lambda _e: on_secondary(), add="+")
