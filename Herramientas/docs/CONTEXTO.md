# Contexto – Arsenal Multioficio 2026

Continuidad local del proyecto Grok web:

https://grok.com/project/ba4f9b70-109f-446e-92dd-0381806d0a94

**Dueño:** Saúl Marín González  
**Carpeta local:** `Herramientas/`  
**Fuente de verdad:** `inventario/Arsenal_Multioficio_2026.xlsx` (+ `inventario/inventario_completo.md`)

---

## Objetivo del proyecto

1. **Migración:** Traer y conservar en local todas las conversaciones y el conocimiento del proyecto en Grok web (respaldo oficial cuando el panel de archivos de grok.com falla).
2. **Operación:** Mantener un inventario vivo del arsenal multioficio (fontanería, electricidad, corte, organización PACKOUT y seguridad), con estados **Ya tengo** / **Comprado** y una lista priorizada de **Pendientes**.

Flujo de trabajo con Grok: ver `../../Grok-Build-Workflow.md` en la raíz de `Grok Projects/` (incluye convención `docs/conversaciones/`).

## Ecosistema de herramientas

| Rol | Marcas / línea |
|-----|----------------|
| Plataforma principal | Milwaukee M12 + M18 FUEL |
| Prueba y eléctrico aislado | Klein Tools |
| Alicates / llaves | Knipex + Wera Joker |
| Organización | Milwaukee PACKOUT |
| Otros | Skil, Antive, Huepar, Franklin, LEXIVON |

## Estado al 28 jul 2026

- Inventario documentado: **40 ítems** (según el export del chat).
- Última compra registrada: **Milwaukee 2522-20** (M12 FUEL 3" Compact Cut Off Tool).
- PACKOUT actual: Backpack `48-22-8301` + Compact Organizer `48-22-8435`.
- M18: Sawzall `2719-20` + batería XC 5.0 + cargador 12-18V (kit comprado).

## Pendientes (orden de prioridad)

1. Klein **65400** – KNECT 15-Piece Pass Through Socket Set  
2. Klein **CL810** – Clamp Meter 600A (~₡80,000)  
3. PACKOUT **48-22-8420** – Rolling Drawer Tool Box  
4. PACKOUT **48-22-8415** – 2-Wheel Utility Cart  
5. Milwaukee **2572B-21** – M12 AIRSNAKE (~₡280,000)

## Observaciones al importar (revisar)

1. **Conteo:** al listar las filas del inventario pegado se cuentan **39** herramientas, no 40. Falta un ítem o el total hay que recalcularlo.
2. **Modelo duplicado en seguridad:** lentes y guantes aparecen ambos como `48-73-2013`. Es probable que los guantes tengan otro SKU (p. ej. línea 48-73-xxxx distinta). Conviene confirmar el modelo real de los guantes.
3. El panel de archivos de grok.com fallaba por caché; **este folder es el respaldo local oficial**.

## Flujos de trabajo sugeridos aquí

| Acción | Cómo |
|--------|------|
| Actualizar inventario | Editar `inventario_arsenal_2026.md` |
| Mover compra | Estado → `Comprado` y quitar de Pendientes |
| Añadir pendiente | Tabla Pendientes + prioridad |
| Exportar tabla | Pedir CSV/XLSX a partir del markdown |
| Kit por oficio | Generar listas Fontanería / Electricidad / Corte / PACKOUT |

## Próximos pasos posibles

- [ ] Corregir conteo (39 vs 40) y SKU de guantes  
- [ ] Generar `inventario.csv` o Excel para filtros  
- [ ] Listas de bolso por oficio (qué va en Backpack vs Organizer)  
- [ ] Comparar precios de pendientes (₡)  
- [ ] Checklist de recepción de compras nuevas  

## Estructura de carpetas

```
Grok Projects/
  Grok-Build-Workflow.md  → preferencias globales + migración Grok web → Build
  Herramientas/
    inventario/           → Excel + Markdown
    docs/                 → contexto, notas, conversaciones/
    scripts/              → regenerar / corregir Excel
```

## Historial de continuidad

| Fecha | Evento |
|-------|--------|
| 2026-07-28 | Export del chat web pegado; respaldo local creado |
| 2026-07-28 | Compra 2522-20 registrada en inventario |
| 2026-07-28 | Enlaces Klein corregidos a fichas oficiales |
| 2026-07-28 | Renombrado a `Arsenal-Multioficio-2026` y dividido en carpetas |
| 2026-07-28 | Contenido de Arsenal movido a la raíz de `Herramientas/` |
| 2026-07-28 | Creado `Grok-Build-Workflow.md` (flujo y preferencias) |
| 2026-07-28 | Workflow adaptado: misión traer conversaciones de Grok web a local |
| 2026-07-28 | `Grok-Build-Workflow.md` movido a la raíz de `Grok Projects/` |
