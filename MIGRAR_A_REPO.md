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

Si la carpeta **se versionaba en el monorepo** (no estaba ignorada), quítala antes del índice:
`git rm -r -q --cached -- "Nombre_De_La_Carpeta"` (los ficheros siguen en disco).

## D · Proyecto versionado DENTRO del monorepo, conservando su historia (2026-10-14)

Así se migraron PPub, Enseñanzas_Propias y Convert_PDF. El árbol de trabajo **no se toca**:
los cambios sin commitear siguen como estaban.

```bash
MONO="f:/__Dugarry UA/Dugarry Proyectos/___Claude.Code"
F="Nombre De La Carpeta"
cd "$MONO"
# 1. Historia de la carpeta en una rama aparte. core.quotepath=off es OBLIGATORIO si el
#    nombre lleva tildes o eñes: sin él, subtree revienta con «assertion failed».
git -c core.quotepath=off subtree split --prefix="$F" v3 -b split-tmp
# 2. Repo nuevo con esa historia, sin tocar los ficheros del disco.
cd "$MONO/$F"
git init -b main
git fetch "$MONO" split-tmp
git reset --mixed FETCH_HEAD
# 3. Comprobar ANTES de publicar: mismo árbol y mismo estado que en el monorepo.
git rev-parse 'HEAD^{tree}'                    # = git -C "$MONO" rev-parse "v3:$F"
git -c core.quotepath=off status --porcelain --untracked-files=all
#    (compáralo con el del monorepo para esa carpeta, quitando el prefijo y las comillas)
# 4. Publicar y sacar la carpeta del monorepo (apartado C).
gh repo create "Nombre-Del-Repo" --private --source="." --remote=origin --push
cd "$MONO" && git rm -r -q --cached -- "$F" && echo "$F/" >> .gitignore
git add .gitignore && git commit -m "$F pasa a su propio repo" && git branch -D split-tmp
```

⚠️ **La memoria de Claude va por repo git**: tras migrar, las sesiones del proyecto usan
`F:\.claude\projects\<ruta-codificada>\memory\` (vacía). Mueve allí su memoria de proyecto y
dale un `MEMORY.md` que enlace a las memorias comunes de la carpeta del monorepo.

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
| Robot_PPub_Conta | `Robot_PPub_Conta` | ✅ creado (2026-10-14, proyecto nuevo con `nueva-app`) |
| Convert_PDF_Liquid_EP_to_Excel | `Convert_PDF_Liquid_EP_to_Excel` | ✅ migrado (2026-10-14, apartado D, 13 commits) |
| Enseñanzas_Propias | `Ensenanzas_Propias` | ✅ migrado (2026-10-14, apartado D, 52 commits) |
| Precios Públicos - PPub | `Precios-Publicos-PPub` | ✅ migrado (2026-10-14, apartado D, 67 commits) |
| Corrector_Orientacion_PDF | `Corrector_Orientacion_PDF` | ✅ rama `master` → `main` (2026-10-14) |
| Sonido_Aviso_Claude | — | Movido a `F:\.claude\sounds\` (2026-10-14, config global) |
| _Config_Ordenadores | — | Se queda en este monorepo (decisión del usuario) |

Desde el 2026-10-14, que cada proyecto tenga su repo es **regla obligatoria** (`CLAUDE.md` global).
