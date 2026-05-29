# Rename Files — Flujo del Programa

## Visión general

Herramienta de escritorio (Python/tkinter, compilada a `.exe`) que renombra o copia-y-renombra ficheros de una carpeta aplicando una sustitución de texto simple en sus nombres.

---

## Pantalla principal — secciones

```
┌─────────────────────────────────────────────────────────┐
│  Carpeta de Ficheros   [ruta/carpeta/origen]  Seleccionar│
├─────────────────────────────────────────────────────────┤
│  Sustitución en nombres de fichero                       │
│    Texto a sustituir:  [_____________________]           │
│    Texto nuevo:        [_____________________]           │
├─────────────────────────────────────────────────────────┤
│  Ficheros en Carpeta          N / Total  Todos  Ninguno  │
│  ┌───────────────┬────────────────────┬──────────┐       │
│  │ Nombre Actual │ New Name           │ Resultado│       │
│  │ fichero_A.xlsx│ fichero_nuevo.xlsx │  -.-     │       │
│  │ fichero_B.xlsx│ fichero_nuevo.xlsx │  -.-     │       │
│  └───────────────┴────────────────────┴──────────┘       │
├─────────────────────────────────────────────────────────┤
│  Carpeta de Destino    [ruta/destino]  Examinar  Limpiar │
├─────────────────────────────────────────────────────────┤
│  Registro              Limpiar   [Renombrar / Copiar...] │
│  > log de operaciones                                    │
└─────────────────────────────────────────────────────────┘
```

---

## Flujo paso a paso

### 1 · Seleccionar Carpeta de Ficheros

- El usuario pulsa **Seleccionar** o escribe la ruta directamente en el campo.
- La ruta se guarda en el registro de Windows (`HKCU\Software\RenameFiles\LastCarpetaExcels`) para recordarla en la próxima sesión.
- Cambiar la ruta dispara `_on_params_change` → espera 300 ms (debounce) → llama a `_load_files`.

### 2 · Escribir la sustitución

- **Texto a sustituir**: cadena que se buscará en cada nombre de fichero.
- **Texto nuevo**: cadena por la que se reemplazará.
- Cada pulsación de tecla dispara `_on_params_change` → debounce 300 ms → `_load_files`.
- Si "Texto a sustituir" está vacío, `New Name` = `Nombre Actual` (sin cambio).

### 3 · Carga y visualización de ficheros (`_load_files`)

```
_load_files()
    │
    ├─ Lee os.listdir(carpeta) → solo ficheros (no subdirectorios)
    ├─ Ordena alfabéticamente
    ├─ Para cada fichero:
    │       new_name = nombre.replace(sustituir, nuevo)
    │       fila = (nombre, new_name, "-.-")
    │
    └─ _populate_tree(filas)
            │
            ├─ Inserta filas en el Treeview con tags even/odd
            ├─ Selecciona todas las filas por defecto
            ├─ _on_sel_change()  → aplica colores de selección
            └─ after(120ms) → _autosize_columns()
```

### 4 · Ajuste de columnas (`_autosize_columns`)

| Columna        | Anchura                                         | Stretch |
|----------------|-------------------------------------------------|---------|
| `Resultado`    | Fija: máx. entre título y valores posibles +15px | No      |
| `Nombre Actual`| Fija: máx. de todos los nombres de fichero +15px | No      |
| `New Name`     | Todo el espacio restante de la ventana           | Sí      |

### 5 · Selección de filas

- **Todos** / **Ninguno**: selecciona o deselecciona todas las filas.
- **Clic individual**: selecciona una fila (deselecciona el resto).
- **Shift + clic**: selección de rango.
- **Ctrl + clic**: añade/quita una fila a la selección.
- Cualquier cambio de selección llama a `_on_sel_change`:
  - Actualiza el tag de cada fila (even / odd / even_sel / odd_sel).
  - Actualiza el contador `N / Total` en la cabecera.

### 6 · Carpeta de Destino (opcional)

| Estado del campo | Comportamiento          | Texto del botón      |
|------------------|-------------------------|----------------------|
| Vacío            | Renombra en la carpeta origen | `Renombrar`    |
| Con ruta         | Copia los ficheros con el nuevo nombre a esa carpeta | `Copiar y Renombrar` |

### 7 · Operación: Renombrar / Copiar y Renombrar (`_do_rename`)

**Validaciones previas:**
- Carpeta de ficheros existe.
- Al menos una fila seleccionada.
- Si hay destino y no existe, se crea con `os.makedirs`.

**Hilo de trabajo (`_rename_thread`)** — se ejecuta en un hilo secundario para no bloquear la UI:

```
Para cada fila seleccionada:
    │
    ├─ Construye orig_path = carpeta / Nombre Actual
    ├─ Verifica que el fichero existe
    │
    ├─ [Modo Renombrar]
    │       os.rename(orig_path, carpeta / New Name)
    │       → Resultado = "Rename"
    │
    └─ [Modo Copiar y Renombrar]
            shutil.copy2(orig_path, destino / New Name)
            → Resultado = "CopyRename"

    Si cualquier paso falla:
        → Resultado = "Falló"
        → Mensaje en el Registro en rojo
```

**Actualización de la columna Resultado** (en el hilo principal vía `after(0, ...)`):

| Valor        | Significado                              |
|--------------|------------------------------------------|
| `-.-`        | Fila no procesada (no seleccionada o reset) |
| `Rename`     | Renombrado con éxito en origen           |
| `CopyRename` | Copiado y renombrado en destino          |
| `Falló`      | Error — ver detalle en el Registro       |

---

## Colores del Treeview

| Tag         | Fondo     | Texto    | Cuándo                        |
|-------------|-----------|----------|-------------------------------|
| `even`      | `#FFFFFF` | `#222222`| Fila par, no seleccionada     |
| `odd`       | `#E8EDF2` | `#222222`| Fila impar, no seleccionada   |
| `even_sel`  | `#5B9BD5` | blanco   | Fila par, seleccionada        |
| `odd_sel`   | `#2E75B6` | blanco   | Fila impar, seleccionada      |

> Los tags anulan el color de selección nativo del tema clam (configurado a nivel Tcl con `ttk::style theme settings clam { ttk::style map Treeview -background {} }`).

---

## Persistencia (Registro de Windows)

| Clave                              | Valor guardado          |
|------------------------------------|-------------------------|
| `HKCU\Software\RenameFiles\LastCarpetaExcels` | Última carpeta de ficheros usada |

---

## Archivos del proyecto

| Fichero             | Descripción                              |
|---------------------|------------------------------------------|
| `Rename_Files.py`   | Código fuente principal                  |
| `Rename_Files.exe`  | Ejecutable compilado (PyInstaller)       |
| `build.bat`         | Script de recompilación                  |
| `requirements.txt`  | Dependencias Python (`pyinstaller`)      |
| `paleta_colores.py` | Utilidad auxiliar para elegir colores    |
| `FLUJO.md`          | Este documento                           |
