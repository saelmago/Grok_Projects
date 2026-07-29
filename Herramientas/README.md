# Arsenal Multioficio 2026

Inventario de herramientas de **Saúl Marín González**.  
Respaldo local del proyecto Grok web (antes `grok-project-ba4f9b70`).

**Misión:** traer y conservar aquí las conversaciones y el conocimiento de Grok web.  
Preferencias globales de Grok Build: **`../Grok-Build-Workflow.md`** (raíz de `Grok Projects/`).

## Estructura de carpetas

```
Grok Projects/
├── Grok-Build-Workflow.md    ← flujo y preferencias (nivel workspace)
└── Herramientas/
    ├── README.md             ← este índice
    ├── inventario/           ← datos del inventario
    │   ├── Arsenal_Multioficio_2026.xlsx
    │   └── inventario_completo.md
    ├── docs/                 ← contexto, notas, conversaciones
    │   ├── CONTEXTO.md
    │   └── conversaciones/
    └── scripts/              ← utilidades para regenerar el Excel
        ├── crear_excel.py
        └── fix_klein_links.py
```

| Carpeta / archivo | Contenido |
|-------------------|-----------|
| **../Grok-Build-Workflow.md** | Preferencias de trabajo con Grok Build (commits, confirmaciones, migración web) |
| **inventario/** | Excel con Resumen / Inventario / Pendientes + enlaces a marca; copia en Markdown |
| **docs/** | Contexto del proyecto, historial, conversaciones importadas |
| **scripts/** | Scripts para recrear o corregir el Excel |

## Archivo principal

Abre: **`inventario/Arsenal_Multioficio_2026.xlsx`**

- Hoja **Resumen** — totales, esquema de códigos y sitios de marca  
- Hoja **Inventario** — líneas con **Código** + **Cantidad** + enlaces  
- Hoja **Pendientes** — 5 compras con **código reservado**  
- Hoja **Consumibles** — kits a granel (`CNS-…`); **Clase** + tipo terminal + color; **Tengo** / **Reponer**  

Código = ID por **línea** (`MW-M12-001`, `CNS-SPD-001`, …), no por unidad física.  
Copia Markdown de consumibles: `inventario/consumibles.md`.  

## Ecosistema

- **Milwaukee** M12 / M18 + PACKOUT  
- **Klein Tools** — prueba y herramientas aisladas  
- **Knipex + Wera** — fontanería  
- Otros: Skil, Antive, Huepar, Franklin, LEXIVON  
