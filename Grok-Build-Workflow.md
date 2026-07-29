# Grok Build - Flujo de Trabajo y Preferencias

**Ubicación:** raíz de `Grok Projects/` (aplica a todos los subproyectos de esta carpeta)  
**Dueño:** Saúl Marín González  
**Fecha de creación (plantilla original):** Abril 2026  
**Adaptado para este workspace:** Julio 2026  

**Objetivo general:** Documentar cómo trabajar de forma segura, cómoda y eficiente usando Grok Build, especialmente combinado con SuperGrok.

**Objetivo de este workspace:** Traer, ordenar y conservar **todas las conversaciones y el conocimiento de Grok web** en local, para que Grok Build sea el respaldo oficial y se pueda seguir trabajando sin depender del panel de archivos de grok.com.

**Primer proyecto migrado (ejemplo):** Arsenal Multioficio 2026 → carpeta `Herramientas/`  
**Proyecto Grok web de origen:**  
https://grok.com/project/ba4f9b70-109f-446e-92dd-0381806d0a94

---

## 0. Misión: de Grok web a Grok Build

### Por qué este workspace existe

- En grok.com el panel de archivos a veces falla (caché, pérdida de contexto).
- Grok Build **no carga sola** toda la conversación web anterior.
- Hay que **exportar y materializar** chats, decisiones e inventarios en archivos locales bajo `Grok Projects/`.

### Estructura del workspace

```
Grok Projects/
├── Grok-Build-Workflow.md     ← este archivo (preferencias globales)
└── Herramientas/              ← proyecto Arsenal Multioficio 2026
    ├── README.md
    ├── docs/
    │   ├── CONTEXTO.md
    │   └── conversaciones/    ← hilos exportados de Grok web
    ├── inventario/
    └── scripts/
```

### Qué se considera "traído" de Grok web (proyecto Herramientas)

| Tipo de contenido | Dónde vive en local | Estado |
|-------------------|---------------------|--------|
| Preferencias de trabajo con Grok | `Grok Projects/Grok-Build-Workflow.md` (este archivo) | Activo |
| Inventario de herramientas | `Herramientas/inventario/Arsenal_Multioficio_2026.xlsx` + `inventario_completo.md` | Importado |
| Contexto del proyecto / historial | `Herramientas/docs/CONTEXTO.md` | Importado (parcial) |
| Hilos / conversaciones del chat web | `Herramientas/docs/conversaciones/` | Pendiente de completar |
| Scripts y utilidades | `Herramientas/scripts/` | Local |

### Flujo para traer una conversación de Grok web

1. En **Grok web**, abrir el hilo o el proyecto.
2. Copiar el contenido relevante (o exportar si hay opción de descarga).
3. En **Grok Build**, pedir algo como:
   - *"Guardá esta conversación de Grok web en `Herramientas/docs/conversaciones/` con fecha y título claros."*
4. Yo propongo nombre de archivo + estructura (Markdown).
5. **El usuario confirma** → se guarda el archivo.
6. Actualizar `Herramientas/docs/CONTEXTO.md` (historial) si la conversación cambia decisiones importantes.
7. Commit intermedio (checkpoint) cuando el lote quede bien.

### Convención de nombres para chats importados

```
Herramientas/docs/conversaciones/
  YYYY-MM-DD_titulo-corto.md
```

Ejemplos:
- `2026-07-28_export-inventario-arsenal.md`
- `2026-07-28_enlaces-klein-y-pendientes.md`

Cada archivo debe incluir al inicio:

```markdown
# Título del hilo

**Origen:** Grok web  
**Proyecto:** ba4f9b70-109f-446e-92dd-0381806d0a94  
**Fecha importación:** YYYY-MM-DD  
**Resumen:** una o dos líneas de qué se decidió o se obtuvo.
```

### Prioridad de importación

1. Decisiones y listas que aún no estén en Excel/Markdown (inventario, pendientes, SKUs).
2. Hilos largos con contexto útil (comparaciones de herramientas, precios, PACKOUT).
3. Mensajes sueltos o capturas → anotar en CONTEXTO o en un solo archivo de "notas pendientes".
4. Evitar duplicar lo que ya es fuente de verdad en `Herramientas/inventario/`.

### Fuente de verdad después del import

| Tema | Fuente de verdad |
|------|------------------|
| Cómo trabajamos con Grok | Este `Grok Projects/Grok-Build-Workflow.md` |
| Qué herramientas tengo / compré / pendientes | `Herramientas/inventario/` (Excel + Markdown) |
| Por qué se tomó una decisión | `Herramientas/docs/CONTEXTO.md` + `conversaciones/*` |
| Chat "histórico" crudo | `Herramientas/docs/conversaciones/*` |

---

## 1. Git y Commits

### Preferencia actual del usuario:
- Normalmente hace commits **al final del día**.
- Con Grok Build quiere hacer **commits intermedios** cuando algo queda funcionando, para tener checkpoints seguros.

### Política de commits recomendada:

**Opción preferida (a confirmar con el usuario):**
- Yo (Grok) **puedo encargarme de hacer commits** cuando una tarea quede bien hecha (p. ej. un lote de conversaciones importadas, inventario actualizado, scripts corregidos).
- Los commits deben tener mensajes claros y descriptivos.
- El usuario puede pedir "no hagas commit" en cualquier momento.

**Alternativa:**
- Yo propongo el commit y pregunto antes de hacerlo.

**Regla general:**
- Antes de empezar una tarea importante, verificar estado de Git (`git status`) si el repo está inicializado.
- Usar commits pequeños y significativos en lugar de un gran commit al final del día.
- Al final del día el usuario puede squash o dejar el historial como está.
- Tras importar conversaciones de Grok web, un commit por lote es un buen checkpoint.

**Ejemplos de mensajes de commit:**
- `Importar conversación web: inventario arsenal 2026-07-28`
- `Actualizar pendientes PACKOUT y Klein`
- `Corregir enlaces oficiales Klein en Excel`

---

## 2. Cómo Solicitar Cambios (código, docs e imports)

### Regla fundamental:
**Yo nunca aplico cambios sin confirmación explícita del usuario.**

Proceso que debe seguirse:

1. El usuario describe lo que quiere (importar un chat, editar inventario, cambiar un script, etc.).
2. Yo analizo lo existente (leo archivos, busco, reviso estructura).
3. Yo propongo los cambios de forma clara (qué archivos se crean/modifican; before/after o diff si aplica).
4. Explico qué voy a modificar y por qué.
5. **El usuario debe decir explícitamente** algo como:
   - "Sí, hacelo"
   - "Aplicá los cambios"
   - "Dale, seguí"
   - "OK"
6. Solo entonces ejecuto ediciones (`search_replace`, creación de archivos, etc.).

### Buenas prácticas:
- Preferir cambios pequeños y precisos en lugar de grandes modificaciones de una sola vez.
- En tareas complejas (p. ej. "traer todas las conversaciones"), usar lista de tareas (`todo_write`).
- Después de cambios importantes, mostrar `git diff` (o un resumen de archivos tocados) para que el usuario revise.
- Al pegar un chat largo de Grok web, primero proponer estructura y nombre de archivo; no volcar todo a `CONTEXTO.md` sin orden.

---

## 3. Seguridad y Reversiones

- Git es la principal herramienta de seguridad.
- Comandos útiles:
  - `git diff` → ver exactamente qué se cambió
  - `git checkout -- archivo` → revertir un archivo específico
  - `git checkout -b nombre-rama` → crear rama antes de cambios grandes

**Recomendación:**
- Antes de hacer cambios grandes o riesgosos (reestructurar carpetas, regenerar Excel, import masivo), crear una rama temporal o un commit de checkpoint.
- Nunca asumir que "ya está bien". Siempre mostrar los cambios antes de pedir confirmación final.
- No borrar conversaciones importadas ni el Excel sin backup: son el historial que se sacó de Grok web.

---

## 4. SuperGrok + Grok Build

### Cómo combinar las dos herramientas:

| Herramienta | Mejor para | Cuándo usarla |
|-------------|------------|---------------|
| **SuperGrok / Grok web** | Charlas largas en el navegador, capturas de productos, comparar herramientas a ojo, hilos del proyecto web | Cuando el contenido aún está en la web o se necesita visión / chat exploratorio |
| **Grok Build** | Guardar conversaciones en disco, editar Excel/Markdown, scripts, Git, estructura de carpetas | Cuando hay que **materializar** lo de la web en archivos locales y trabajar con control |

### Flujo híbrido recomendado (migración + trabajo diario):

1. **Grok web / SuperGrok:** Trabajar o recuperar el hilo (chat, capturas, listas).
2. **Usuario:** Copiar el contenido o resumir lo importante.
3. **Grok Build:** Proponer dónde guardarlo (`Herramientas/docs/conversaciones/...`) y qué actualizar en inventario/CONTEXTO.
4. **Confirmación del usuario** → guardar y, si aplica, commit intermedio.
5. Si hace falta imagen de un producto o catálogo → SuperGrok con captura; el resultado se anota en local con Grok Build.
6. Si el panel de archivos de grok.com falla → **no depender de él**; este workspace es el respaldo oficial.

**Importante:** No intentar hacer todo en una sola herramienta. La web genera y explora; Build archiva y edita con seguridad.

---

## 5. Persistencia entre Sesiones

- Cuando se cierra Grok Build y se vuelve a abrir, **no se carga automáticamente** toda la conversación anterior (ni la de Grok web).
- Para mantener el contexto entre sesiones hay que **documentar** decisiones e imports en archivos.

### Recomendación:
- Mantener este archivo en la **raíz de `Grok Projects/`**.
- Mantener el `CONTEXTO.md` de cada subproyecto al día.
- Guardar cada hilo importante de Grok web en la carpeta `docs/conversaciones/` del subproyecto correspondiente.
- Al empezar una nueva sesión, indicarle a Grok:

  **"Leé el archivo Grok-Build-Workflow.md (en la raíz de Grok Projects) y seguí las preferencias que están ahí."**

  Si se trabaja en Herramientas:

  **"Leé Grok-Build-Workflow.md y Herramientas/docs/CONTEXTO.md."**

Esto permite retomar el trabajo sin tener que re-pegar todo el chat web.

---

## 6. Cómo Empezar una Sesión de Grok Build

Indicaciones útiles al abrir Grok Build en `Grok Projects/`:

- "Leé el archivo Grok-Build-Workflow.md"
- "Estado actual del proyecto y de Git"
- "Qué conversaciones de Grok web ya están importadas y cuáles faltan"
- "Vamos a importar este chat de Grok web: [pegar o describir]"
- "Actualizá el inventario de Herramientas con [compra / pendiente]"
- "Antes de tocar nada, mostrame qué archivos vas a crear o cambiar"

---

## 7. Resumen de Preferencias del Usuario

- Quiere **traer y conservar** el conocimiento de Grok web en local (no solo "chatear de nuevo").
- Quiere **seguridad** y facilidad para revertir cambios.
- Prefiere que Grok lo ayude con los **commits intermedios** (salvo que diga lo contrario).
- Le gusta que se le muestren los cambios **antes** de aplicarlos.
- Quiere combinar SuperGrok / Grok web (exploración y visión) + Grok Build (archivo, edición, Git).
- Valora **checkpoints** claros durante el trabajo y tras cada lote importado.
- Este archivo de preferencias vive a **nivel de `Grok Projects/`**, no dentro de cada subcarpeta.

---

## Notas

Este documento puede (y debe) actualizarse con el tiempo a medida que se ajusten las preferencias, se complete la importación de conversaciones o se descubran mejores formas de trabajar.

**Última actualización:** 28 de julio 2026

---

**Fin del documento**
