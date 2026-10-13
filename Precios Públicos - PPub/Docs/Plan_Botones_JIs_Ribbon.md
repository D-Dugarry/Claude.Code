# Plan: JI's del Informe de Recibos — `M_130` los trae de BDatos y un botón del Ribbon los pasa a BDatos

**Proyecto:** `PPub_BDatos_2026.xlsm` · **Fecha:** 2026-10-13 · **Afecta a:** `M_130` y `M_133` (cambiados), `M_132`
(sobra, D6), `M___RibbonUI` (un callback), `customUI14.xml` (un grupo con un botón), `Lo_RibbonUI` (dos filas), hoja
`Inf_Recibos_TIO` (borrar las dos formas), documentación.

**Etiquetas:** ✅ **Hecho** · 🧪 **Aplicado, falta probarlo en Excel** · ⏳ **Pendiente**

## Estado

- ✅ **`M_130` y `M_133` cambiados** y probados con una doble ejecución por COM en una copia del libro (sección 4);
  importados en el libro (comprobado con `olevba` el 2026-10-13).
- 🧪 **Botón del Ribbon puesto (2026-10-13)**, con el icono `Copy_JIs`: XML, callback y filas de `Lo_RibbonUI` (sección 5).
  **Falta** importar `M___RibbonUI` y probarlo en Excel; después, borrar la forma «Añadir JI's del Inf. a BD».
- ✅ Decidido (2026-10-13): las 3 hojas antiguas **se quedan** de momento (D5 = a); **`M_132` se retira** (D6); la
  **traza de tiempos de `M_130` se quita**. Hecho en el repo; la forma de `M_132` ya está borrada en el libro. Falta
  quitar el módulo `M_132_Anadir_Num_JIs_al_Inf` del libro e importar el `M_130` sin la traza.

## 1. Qué se pide (decisiones del usuario del 2026-10-13)

1. **«Añadir JI's de BD al Inf.» deja de ser un botón:** lo hace `M_130` al calcular el informe (ábaco). Había que
   estudiar si bastaba con llamar a `M_132` o era mejor hacerlo en la misma pasada de `M_130` (sección 2).
2. **«Añadir JI's del Inf. a BD»** pasa a un botón del Ribbon, **«JI's del Inf. a BD»** (D4), sin formulario de
   progreso (D2).
3. **Quitar la pregunta de Enseñanzas Propias** de `M_132` (D1).
4. **Arreglar `M_132` y `M_133`** (lo visto en la primera versión de este plan).

Sale de la sección 3.4 de `Docs/Plan_Mail_Inf_Recibos.md`, que preveía un grupo «Informe de Recibos» con el botón de
correo y los dos de JI's. Con este plan, ese grupo tendrá el botón de correo y uno de JI's; hay que actualizar aquella
3.4 cuando la sesión del correo no esté tocando el fichero.

## 2. Estudio: ¿llamar a `M_132` desde `M_130`, o hacerlo en la misma pasada?

| | Llamar a `M_132` al final de `M_130` | En la pasada en memoria de `M_130` (**elegido**) |
|---|---|---|
| Cómo | Ordena BDatos (dos ordenaciones) y, por cada una de las 15 filas del informe, filtra BDatos y recorre una a una las celdas visibles de las 6 columnas | Lee de una vez las 6 columnas de JI's de BDatos y junta los JI's en el mismo recorrido que ya hace para las sumas (`Rut_Inf_Recibos_Sumas`) |
| Tiempo medido (copia, COM) | M_130 1,77 s **+ M_132 3,95 s** | **1,81 s** en total (+0,04 s) |
| BDatos | Queda ordenada por `Tipo_Rec` + `Concepto` | No se ordena (como `M_130` desde el 2026-10-12) |
| Código | Dos módulos; `M_132` con sus fallos (sin control de errores, BDatos sin proteger, pregunta) | Todo en `M_130`, que ya tiene control de errores; `M_132` sobra |
| Resultado | — | **Idéntico** (doble ejecución, sección 4) |

Las reglas de `M_132` se reproducen tal cual:

- Por cada fila del informe, los recibos de BDatos con el **mismo** `Tipo_Rec` (igual, sin distinguir mayúsculas; las
  sumas, en cambio, usan «contiene») y el mismo concepto, **sin mirar fechas ni importes**.
- En cada columna, los valores distintos en el orden de BDatos, separados por «, ».
- **«·-·»** si hay recibos y ninguno tiene dato en esa columna; **vacía** si no hay recibos de ese tipo y concepto.
- Las filas de EURLE se siguen separando sin JI's (antes, `M_132` corría con EURLE ya fuera de la tabla).

Única diferencia, a propósito: los blancos, los errores y el propio «·-·» no cuentan como JI. Hoy no hay ninguno en
BDatos.

## 3. Cambios hechos (🧪, en el repo)

### 3.1 `M_130_Gen_Inf_Recibos`

- **JI's en la misma pasada:** `Rut_Inf_Recibos_Sumas` devuelve además `JIs(tipo, concepto, columna)` y la rutina
  principal los escribe al crear cada fila del informe.
- **Sin la pregunta de Enseñanzas Propias** (D1). No hace falta: el informe no tiene filas de conceptos AE4 (311 y 312).
- Nuevas, públicas porque las usa `M_133`: las constantes `JI_N_Col` (6) y `JI_Sin_Num` («·-·»), y la función
  `Fnc_JI_Con_Dato`.
- **BDatos queda protegida como estaba**, también si algo falla (`Rut_Prot_Save` al empezar, `Rut_Prot_Restore` en
  `Fin`). Antes se quedaba sin proteger hasta reabrir el libro.

### 3.2 `M_133_Anadir_Num_JIs_a_BDatos`

- **Control de errores** como `M_130`: `On Error GoTo Fin` con la variable `Paso`, y `Rut_Off_Functions`/
  `Rut_On_Functions`. Si algo falla: BDatos sin filtros, filas de totales visibles y un aviso que dice en qué paso.
- **BDatos queda protegida como estaba** (`Rut_Prot_Save`/`Rut_Prot_Restore`).
- **Mira las 6 columnas** de JI's para decidir si copia una fila. Antes miraba solo las 4 primeras, así que una fila con
  solo JI 4430 se saltaba.
- **«·-·» ya no se copia a BDatos.** Antes se copiaba como si fuera un JI a todas las celdas sin JI.
- **Una sola ordenación** de BDatos y del informe (antes, dos de cada). Da el mismo orden.
- Deja el resultado en `APP_Task_Inf`, por ejemplo: «13/10/2026 12:07  JI's del informe a BDatos: 6 filas del informe,
  104.979 recibos (1,6 s)». El callback del botón lo guardará en `Informe_Rut` (supertip de Boss).
- **Se mantiene:**
  - el filtro de Excel por `Tipo_Rec` y `Concepto`;
  - el salto del concepto 311;
  - que el informe manda: en una fila que se copia, una celda vacía (o con «·-·») vacía esa columna en sus recibos.

### 3.3 `M_132_Anadir_Num_JIs_al_Inf`

Sin tocar: con 3.1 ya no hace falta (D6). Mientras siga en el libro, su forma funciona como antes, con la pregunta.

## 4. Prueba: doble ejecución (2026-10-13, ✅)

En dos copias del libro guardado a las 11:46, cada una en una instancia de Excel aparte (el Excel del usuario, intacto):

1. **Semillas** iguales en las dos copias: JI's en BDatos de varios tipos.
   - Números `4711`/`4712` en Emitido 1303.
   - Textos «2026-100», «2026-101» y «2026-102» en Emitido 1310.
   - «4430-A» en Aplazado 1310.01.
   - Uno en EjeAnt 1310 y otro en EURLE.
2. **Viejo:** `M_130` y `M_132` (contestando No). **Nuevo:** `M_130`.
3. Los mismos **cambios a mano** en el informe:
   - en Aplazado 310.01, vaciar las 4 primeras columnas de JI's y poner «TEST-443» en JI 4430 Adm.;
   - en Emitido 303.00, poner «JI-NEW-1».
4. `M_133` (viejo o nuevo), y guardar.

**Resultado:**

- **Informe idéntico** en valores, rellenos y formatos (`Comparar_Libros_BD.py`).
- **BDatos:** mismo orden y mismas celdas fuera de las 6 columnas de JI's. En ellas, solo lo esperado:
  - **«·-·»:** el viejo lo escribía en todas las celdas sin JI de los 15 grupos (unas 816.000); el nuevo las deja
    vacías.
  - **Aplazado 1310.01** (la fila con solo JI 4430): el viejo se la saltaba. El nuevo la copia a sus 3.527 recibos:
    «TEST-443», «4430-A» y las AD vaciadas, como estaban en el informe.
- **Tiempos:**

  | Paso | Antes | Ahora |
  |---|---|---|
  | `M_130` | 1,77 s | 1,81 s (con los JI's) |
  | `M_132` | 3,95 s | — |
  | `M_133` | 2,77 s | 1,58 s |

- **Al acabar:** BDatos protegida (antes, sin proteger), el contador de `Rut_Off_Functions` a 0, sin filtros y con las
  filas de totales visibles.

## 5. El botón del Ribbon (🧪, puesto el 2026-10-13, falta probarlo en Excel)

**Hecho el 2026-10-13** (con Excel y el Custom UI Editor cerrados; copia previa en
`CopiaSeguridad/PPub_BDatos_2026_antes_Ribbon_Inf_Recibos_2026-10-13_2028.xlsm`):

- **XML:** grupo `Inf_Recibos_Group` detrás del ábaco con este botón (icono **`Copy_JIs`**, que ya estaba en el libro) y el
  de correo (`BtnMail_Inf_Recibos`, icono `Mail`). En la misma edición, las etiquetas «Prueba: NO» / «Prueba: SÍ» del
  plan del correo. Zip reconstruido y verificado: de 333 entradas solo cambia `customUI/customUI14.xml`.
- **Callback** `OnAct_JIs_Inf_a_BD` en `M___RibbonUI`, como el de abajo, con `App_Task_Inf` en vez de `APP_Task_Inf`.
- **Filas en `Lo_RibbonUI`**: las dos de la tabla de abajo y la de `Mail_Inf_Recibos` (`Rut_Mail_Inf_Recibos`), detrás
  de las del ábaco. Añadidas por COM en un Excel aparte (eventos y macros apagados).
- **Formas:** borrada «Añadir JI's de BD al Inf.» (`M_132`). Queda «Añadir JI's del Inf. a BD» hasta probar el botón.
- `Validar_Ribbon.py`: 0 errores (5 avisos que ya estaban).

Lo que se preveía:

⚠️ **`customUI14.xml`: aviso antes y después** (el Custom UI Editor no deja fichero de bloqueo).

- **XML**, en `TabUserMenu`, detrás del grupo del ábaco (`ReCalculate_ListObj_Sht_Group`). El botón de correo se
  añadirá a este grupo:

  ```xml
              <!-- Informe de Recibos: JI's (y, mas adelante, el correo) _______________ -->
                  <group id="Inf_Recibos_Group"                   tag="Inf_Recibos_Group"   label="Informe de Recibos"
                              getVisible="GetVsbl_CtrlTab">
                      <button id="BtnJIs_Inf_a_BD"                tag="JIs_Inf_a_BD"        image="(D3)"   size="large"
                              label="JI's del Inf. a BD"
                              getVisible="GetVsbl_CtrlTab"
                              getSupertip="getStip_CtrlTab"       screentip="Añadir a BDatos los JI's del informe"
                              onAction="OnAct_JIs_Inf_a_BD"          />
                  </group>
  ```

- **Callback** en `M___RibbonUI`, junto a `OnAct_ReCalculate_ListObj_Sht`. Se escribe cuando toque el XML, para no
  coincidir con la sesión del correo, que también va a tocar `M___RibbonUI`:

  ```vba
  Sub OnAct_JIs_Inf_a_BD(control As IRibbonControl)
  Debug.Print "OnAct_JIs_Inf_a_BD"
      Call Rut_Anadir_a_BDatos_JIs_de_Inf_Recibos             '- M_133: deja el resultado en APP_Task_Inf
      Call Rut_RibbonUI_Guardar_Informe(control.Tag)          '- APP_Task_Inf -> Informe_Rut (supertip de Boss)
      Call RefreshRibbon
  End Sub
  ```

- **Filas en `Lo_RibbonUI`**, las dos con `Usuario` = `Boss,SusanaC,FannyR;RafaG`, `SheetsNames` = `Inf_Recibos_TIO` y
  `Group-Tag` = `Inf_Recibos_Group`:

  | `Uribbon-Tags` | `Nom_Rut` | `Descripion` |
  |---|---|---|
  | `Inf_Recibos_Group` | | Botones del Informe de Recibos. |
  | `JIs_Inf_a_BD` | `Rut_Anadir_a_BDatos_JIs_de_Inf_Recibos` | Escribe en los recibos de BDatos los JI's y AD's del informe, por tipo de recibo y concepto. Solo las filas con algún JI; en ellas, una celda vacía o con «·-·» vacía esa columna en sus recibos. Los JI's de BDatos al informe los pone el cálculo del informe (ábaco). |

## 6. Orden de trabajo

**Estado (2026-10-13 noche):** pasos 1, 2 (salvo quitar `M_132` del libro) y 3 hechos, falta probar el botón; el
cierre del paso 4, hecho (`CLAUDE.md`, el plan del correo, volcado, validación y commit).

1. ✅ **(yo)** `M_130` y `M_133` en el repo, probados por COM (sección 4).
2. ✅ **(usuario)** D6 = retirar `M_132`. Después:
   - importar `M_130_Gen_Inf_Recibos` y `M_133_Anadir_Num_JIs_a_BDatos`;
   - si D6 = retirar, quitar `M_132_Anadir_Num_JIs_al_Inf`;
   - borrar la forma «Añadir JI's de BD al Inf.»;
   - Depuración → Compilar, y guardar.

   Para probarlo, pulsar el ábaco: salen los JI's en el informe, sin pregunta. Luego la forma «Añadir JI's del Inf. a
   BD», que sigue hasta que exista el botón: el resultado queda en `APP_Task_Inf` y BDatos sigue protegida.
3. ⏳ **D3 (icono)** → callback, XML (con el libro cerrado) y filas de `Lo_RibbonUI` → probar → borrar la otra forma.
4. ⏳ **Cierre:**
   - `CLAUDE.md`;
   - la 3.4 del plan del correo;
   - `CustomUI/Volcar_Tareas_y_Ribbon.py` y `Validar_Ribbon.py`;
   - commit.

## 7. Decisiones

1. ✅ **D1. Pregunta de Enseñanzas Propias:** quitada. `M_130` no pregunta nada.
2. ✅ **D2. Formulario de progreso:** no.
3. ✅ **D3. Icono del botón:** `Copy_JIs` (elegido por el usuario el 2026-10-13; ya estaba entre las imágenes del
   Ribbon del libro).
4. ✅ **D4. Rótulo:** «JI's del Inf. a BD».
5. ✅ **D5. Hojas antiguas del informe: a), se quedan como están** (el usuario las mirará más adelante, 2026-10-13).
   Ojo: la forma «Añadir JI's de BD al Inf.» de `Inf_Recibos_TIO (2)` dará «No se puede ejecutar la macro» en cuanto
   se quite `M_132` del libro. En el libro hay tres copias viejas del informe, **muy ocultas** (no salen en
   las pestañas). **Sí salen en el menú de hojas del Ribbon** («Select Sheets»), que lista todas las hojas cuyo
   CodeName empieza por `Sht__BD` o `Sht__Inf`, y «Show All» las enseña.

   | Hoja | CodeName | Calculada | Qué tiene |
   |---|---|---|---|
   | `Inf_Recibos_TIO (2)` | `Sht__Inf_Recibos_TIO1` | 17-feb-2026 | El informe de entonces y las **3 formas de entonces**: «Recalcular Tabla» (`M_130`), «Añadir JI's de BD al Inf.» (`M_132`) y «Añadir JI's del Inf. a BD» (`M_133`). No está en ninguna tabla de configuración ni tiene `.cls` en `VBA_Moduls/` |
   | `Inf_Recibos_TIO y EP` | `Sht__Inf_Recibos_TIOyEP` | 15-feb-2026 | El informe con las filas de Enseñanzas Propias, sin formas. Tiene fila en `Prog_Hojas_Names` y en `Lo_ListObjAPP` |
   | `Inf_RecibosJIs` | `Sht__Inf_RecibosJIs` | 23-dic-2025 | Un informe anterior, sin formas. Tiene fila en `Prog_Hojas_Names` y en `Lo_ListObjAPP` |

   Nada las lee: ni código, ni fórmulas, ni nombres. Son fotos de valores de aquellas fechas. Con este plan solo
   importa la primera, porque sus formas siguen vivas:
   - Si se abre desde el menú de hojas, sus formas se pueden pulsar. **Actúan sobre el informe bueno**, no sobre la copia:
     «Recalcular Tabla» recalcula `Inf_Recibos_TIO` y «Añadir JI's del Inf. a BD» copia a BDatos los JI's de
     `Inf_Recibos_TIO`.
   - Si se retira `M_132` (D6), su forma «Añadir JI's de BD al Inf.» dará «No se puede ejecutar la macro».

   Opciones:
   - a) **Dejarlas** como están.
   - b) **Borrar solo las 3 formas** de `Inf_Recibos_TIO (2)` y conservar las hojas como histórico.
   - c) **Borrar las tres hojas**, como se hizo con las de AE4, con sus filas en `Prog_Hojas_Names` y `Lo_ListObjAPP` y
     los `.cls` del repo (`Sht__Inf_Recibos_TIOyEP`, `Sht__Inf_RecibosJIs`).

   Recomendación: **b)** como mínimo, para que no quede ningún botón viejo que se pueda pulsar desde otra hoja; **c)** si
   no las quieres como histórico.
6. ✅ **D6. Retirar `M_132`: sí** (2026-10-13). `.bas` fuera del repo y su forma de `Inf_Recibos_TIO` borrada; el módulo
   lo quita del libro el usuario. Sin `M_132`, la pregunta de D1 desaparece con él.

## 8. Visto y sin tocar

- **`Rut_Lo_WrkSht_Preparar` desprotege la hoja que prepara y no la vuelve a proteger.** BDatos se queda sin proteger
  tras casi cualquier proceso (`M_110`, `M_210`…), no solo tras los de JI's. `M_130` y `M_133` ya la dejan como estaba,
  porque guardan su estado antes de llamar a `Preparar`; el resto, sin tocar.
- **`M_133` escribe los JI's con `.Value`, y las columnas de JI's de BDatos tienen formato General.** Un JI escrito
  como «2026-1» se convertiría en fecha: en BDatos, y ya en el propio informe al teclearlo, si la celda no tiene
  formato de texto. Los JI's de hoy son del tipo «2025-6217» y no se convierten.
- La traza de tiempos de `M_130` en Inmediato (`Rut_Inf_Recibos_Traza`) **se quitó el 2026-10-13** (decisión del
  usuario). Sus llamadas pasaron a `Paso = "..."`, que es lo que dice el aviso de error.
