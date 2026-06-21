# Migrar un proyecto a su propio repositorio independiente

Procedimiento para convertir una subcarpeta de este workspace (`___Claude.Code`) en un
repositorio git **independiente** con su propio repo **privado** en GitHub.
Resultado: cada proyecto tiene su rama propia y, al abrir su carpeta, VSCode/git
recuerdan y usan esa rama automáticamente (sin mezclarse con el monorepo).

---

## Requisitos previos (una sola vez por equipo)

```bash
# 1. GitHub CLI instalado
winget install --id GitHub.cli -e --accept-source-agreements --accept-package-agreements --silent

# 2. Autenticar gh (interactivo: GitHub.com -> HTTPS -> navegador)
gh auth login
gh auth status            # debe decir: Logged in to github.com account D-Dugarry

# 3. Disco F: sin "ownership": marcar todo como seguro (evita "dubious ownership")
git config --global --add safe.directory '*'
```

> En git-bash, `gh` puede no estar en el PATH: `export PATH="$PATH:/c/Program Files/GitHub CLI"`

---

## A · Proyecto NUEVO (la carpeta aún no tiene `.git`)

Desde dentro de la carpeta del proyecto:

```bash
cd "ruta/del/Proyecto"

git init -b main

# .gitignore: ignora artefactos y DATOS (ajusta según el proyecto)
printf '__pycache__/\n*.pyc\nbuild/\ndist/\n*.exe\n' > .gitignore
# Python con datos:   añade  *.xlsx  *.xlsm  carpetas de datos/
# VBA:                ignora *.xlsm *.xlsb *.frx  y los  *_BACKUP_*/

# Añade SOLO el código/docs/assets (no .vscode/.claude ni datos):
git add <archivos de código> .gitignore
#   o, si el .gitignore ya cubre todo lo no deseado:  git add -A

git commit -m "Import inicial"

# Crea el repo privado en GitHub y sube main de una vez:
gh repo create "Nombre-Del-Repo" --private --source="." --remote=origin --push
```

## B · Proyecto que YA tiene `.git` (p. ej. Certificados_de_Pago)

```bash
cd "ruta/del/Proyecto"
git branch -m master main        # (si está en master, uniformar a main)
git status                       # revisa cambios sin commitear (NO los toques si son trabajo en curso)
gh repo create "Nombre-Del-Repo" --private --source="." --remote=origin --push
```
`--push` sube solo lo ya **commiteado**; los cambios pendientes quedan locales.

## C · Limpiar el monorepo padre (`___Claude.Code`)

Para que el padre deje de "ver" la carpeta migrada (evita el repo anidado):

```bash
cd "f:/__Dugarry UA/Dugarry Proyectos/___Claude.Code"
# añade la carpeta al .gitignore del padre
echo "Nombre_De_La_Carpeta/" >> .gitignore
git add .gitignore && git commit -m "Ignorar <proyecto>: migrado a repo independiente"
```

---

## Convenciones

| Punto | Regla |
|---|---|
| Rama principal | `main` |
| Visibilidad | privado (`--private`); cambiar luego en GitHub si hace falta |
| Nombre del repo | sin espacios → guiones (`Rename_Files - Special` → `Rename_Files-Special`) |
| Qué versionar | código, docs, `.spec`, `build.bat`, `requirements.txt`, assets pequeños |
| Qué NO versionar | `*.exe`, `build/`, `dist/`, `__pycache__/`, datos (`*.xlsx`…), backups, `.vscode/`, `.claude/` |

## Gotchas observados

- **"dubious ownership"** en `F:` → `git config --global --add safe.directory '*'` (una vez).
- Si **`gh auth login` no es posible** y hay credencial guardada para push, se puede crear el
  repo por la API REST con ese token (scope `repo`):
  `Invoke-RestMethod -Uri https://api.github.com/user/repos -Method Post -Headers @{Authorization="token <TK>"} -Body '{"name":"X","private":true}'`
- Avisos **`LF will be replaced by CRLF`**: normales en Windows, no afectan.
- Antes de migrar, comprueba que el **script principal existe** (p. ej. el `.py` que cita el
  `.spec`). Rename_Files quedó sin migrar porque faltaba `Rename_Files.py`.

---

## Estado de la migración (2026-06-03)

| Proyecto | Repo `D-Dugarry/…` | Estado |
|---|---|---|
| Mailing_Personalizado | `Mailing_Personalizado` | ✅ migrado |
| Norma43 (desde Norma43.Code) | `Norma43` | ✅ migrado |
| Certificados_de_Pago | `Certificados_de_Pago` | ✅ migrado |
| Rename_Files - Special | `Rename_Files-Special` | ✅ migrado |
| Rename_Files | `Rename_Files` | ✅ migrado (Rename_Files.py recuperado) |
| PowerBI_PPUA | `PowerBI_PPUA` | ✅ migrado (2026-06-20, proyecto nuevo en main) |
| PPUA_VBA | `PPUA_VBA` | ✅ migrado (2026-06-21, proyecto nuevo en main; .xlsm fuera de git) |
