# Chuleta — Git + Claude Code entre dos ordenadores

> **Situación**: Trabajas en el ordenador del **trabajo** y quieres continuar en **casa**
> (o al revés). Esta guía explica qué hacer antes de irte y al llegar.

---

## 📤 ANTES DE IRTE — Ordenador del trabajo

### 1. Actualizar el punto de continuación en CLAUDE.md

1. En VS Code, abre el archivo **CLAUDE.md** del proyecto
2. Pulsa `Ctrl+End` para ir al final del archivo
3. Añade este bloque (rellena los huecos):

```
---

## 🔖 Punto de continuación — FECHA DE HOY

- Último commit: [copia aquí el hash corto, ej. fcbf3e2]
- Estaba trabajando en: [describe qué estabas haciendo]
- Próximo paso: [qué ibas a hacer después]
- Pendiente / dudas: [lo que quedó sin resolver]
```

4. Guarda con `Ctrl+S`

> **¿Cómo sé el hash del último commit?**
> Abre el terminal (`Ctrl+ñ`) y escribe:
> ```
> git log --oneline -1
> ```
> Te aparece algo como `fcbf3e2 Descripción del commit`. El hash son los 7 primeros caracteres.

---

### 2. Hacer commit y push (subir a GitHub)

Abre el terminal en VS Code (`Ctrl+ñ`) y escribe estos comandos **uno a uno**:

```bash
git add CLAUDE.md
```
```bash
git commit -m "Punto de continuación"
```
```bash
git push
```

> Si git push te pide contraseña, usa tu **Token Personal de GitHub**
> (no tu contraseña normal de GitHub — GitHub no acepta passwords, solo tokens).

---

### 3. Verificar que se subió correctamente (opcional)

```bash
git status
```

Debería decir: `nothing to commit, working tree clean`

✅ **Ya puedes irte.** El código y el punto de continuación están en GitHub.

---

## 📥 AL LLEGAR — Ordenador de casa (o el otro equipo)

### 1. Descargar los cambios de GitHub

Abre el terminal en VS Code (`Ctrl+ñ`) y escribe:

```bash
git pull
```

> Si hay un error de rama, primero comprueba en qué rama estás:
> ```bash
> git branch
> ```
> La rama activa lleva un `*` delante. Si no es la correcta:
> ```bash
> git checkout nombre-de-la-rama
> git pull
> ```

---

### 2. Leer el punto de continuación

Abre **CLAUDE.md** y busca la sección `🔖 Punto de continuación` al final.
Lee qué estabas haciendo y qué tenías pendiente.

---

### 3. Arrancar Claude Code con contexto

Al abrir un nuevo chat en Claude Code, escribe algo como:

> *"Continúo el trabajo de ayer. Lee el CLAUDE.md, especialmente el punto de continuación al final, y dime si lo has entendido."*

Claude Code leerá el fichero y sabrá exactamente dónde lo dejaste.

---

### 4. Borrar la sección de continuación del CLAUDE.md

Una vez retomado el trabajo, borra el bloque `🔖` del CLAUDE.md:

1. Selecciona las líneas del bloque (desde `---` hasta el final)
2. Pulsa `Supr`
3. Guarda con `Ctrl+S`

Luego haz commit:

```bash
git add CLAUDE.md
git commit -m "Retomado trabajo — borrado punto de continuación"
```

---

## 🔄 Resumen visual del flujo completo

```
TRABAJO                        GITHUB                        CASA
   │                              │                             │
   │  1. Editar CLAUDE.md         │                             │
   │     (punto continuación)     │                             │
   │                              │                             │
   │  2. git add CLAUDE.md        │                             │
   │     git commit -m "..."      │                             │
   │     git push  ─────────────► │ sube el código              │
   │                              │ y el punto                  │
   │                              │ de continuación             │
   │                              │                             │
   │                              │ ◄─── git pull ──────────────│
   │                              │                             │
   │                              │      leer CLAUDE.md         │
   │                              │      arrancar Claude Code   │
   │                              │      continuar trabajo      │
```

---

## ⚠️ Errores frecuentes y soluciones

| Error | Causa | Solución |
|-------|-------|----------|
| `git push` pide contraseña | GitHub ya no acepta passwords | Usa el Token Personal |
| `git pull` falla con "divergent branches" | Hay cambios en ambos lados | `git pull --rebase` |
| `git push` dice "rejected" | Hay cambios en GitHub que no tienes | Haz `git pull` primero, luego `git push` |
| No encuentro el CLAUDE.md | Cada proyecto tiene el suyo | Está en la carpeta raíz del proyecto |

---

## 📋 Comandos de referencia rápida

```bash
git status              # Ver qué ficheros han cambiado
git log --oneline -5    # Ver los últimos 5 commits
git add CLAUDE.md       # Añadir un fichero al commit
git add .               # Añadir TODOS los ficheros cambiados
git commit -m "texto"   # Hacer commit con un mensaje
git push                # Subir a GitHub
git pull                # Descargar de GitHub
git branch              # Ver en qué rama estás
```

---

**Última actualización**: 2026-05-28
**Para**: Flujo de trabajo entre ordenador del trabajo y casa
