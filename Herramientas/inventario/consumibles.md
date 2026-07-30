# Consumibles – Arsenal Multioficio 2026

**Regla:** por **Clase** + **tipo de terminal** + **color** + **Ubicación** (código del contenedor).  
**Excel:** pestaña `Consumibles`  
**Códigos:** `CNS-*-###`

> **Pendiente (pri. 6):** `CNS-ORD-AMZ-001` — **Terminar de cargar la orden de consumibles de Amazon**  
> (Spade, solder seal, BHT crimp y non-insulated ya cargados; falta el resto de la orden cuando continúe.)

| Estado | Significado |
|--------|-------------|
| **Tengo** | Hay stock usable |
| **Reponer** | Bajo el umbral o se acabó |

**Ubicación** = código del lugar en el inventario.

| Código ubicación | Qué es |
|------------------|--------|
| MW-PKO-001 | PACKOUT Backpack |
| MW-PKO-002 | PACKOUT Compact Organizer |
| **MW-PKO-005** | **PACKOUT Organizer 19"** `48-22-8431` |

---

## Clases

| Clase | Qué va ahí | Tipo terminal |
|-------|------------|---------------|
| **Conector eléctrico** | Spade, ring, fork, butt, wire nut | Obligatorio |
| **Cinta / aislante** | Cinta eléctrica, vulcanizante… | — |
| **Fijación** | Tornillos, tuercas, cinchos | — |
| **Corte / desgaste** | Hojas, lijas, brocas de uso | — |
| **Químico / sellado** | Grasa, Loctite, sellador | — |
| **Otro** | Si no cabe | — |

---

## Stock actual

**Ubicación de estos conectores: `MW-PKO-005`** (PACKOUT Organizer 19")

### Tipo **Spade** (pala / quick disconnect)

Kit: [haisstronica 160PCS Heat Shrink Spade](https://www.amazon.com/dp/B0D1K6L6TL) (B0D1K6L6TL)

| Código | Tipo terminal | Color | AWG | Cant. | Umbral | Estado | Ubicación |
|--------|---------------|-------|-----|------:|-------:|--------|-----------|
| CNS-SPD-ROJ-001 | Spade (pala / quick disconnect) | Rojo | 22-16 | 60 | 15 | Tengo | MW-PKO-005 |
| CNS-SPD-AZU-001 | Spade (pala / quick disconnect) | Azul | 16-14 | 60 | 15 | Tengo | MW-PKO-005 |
| CNS-SPD-AMA-001 | Spade (pala / quick disconnect) | Amarillo | 12-10 | 40 | 10 | Tengo | MW-PKO-005 |
| | | **Total Spade** | | **160** | | | |

### Tipo **Butt** (empalme / solder seal)

Kit: [haisstronica 120PCS Solder Seal Wire Connectors](https://www.amazon.com/dp/B07C3NBTJ9) (B07C3NBTJ9)  
Cantidades según packaging del fabricante:

| Código | Tipo terminal | Color | AWG | mm² | Cant. | Umbral | Estado | Ubicación |
|--------|---------------|-------|-----|-----|------:|-------:|--------|-----------|
| CNS-BTT-BLA-001 | Butt (empalme / solder seal) | Blanco | 26-24 | 0.25-0.34 | **25** | 6 | Tengo | MW-PKO-005 |
| CNS-BTT-ROJ-001 | Butt (empalme / solder seal) | Rojo | 22-18 | 0.5-1.0 | **45** | 10 | Tengo | MW-PKO-005 |
| CNS-BTT-AZU-001 | Butt (empalme / solder seal) | Azul | 16-14 | 1.5-2.5 | **40** | 10 | Tengo | MW-PKO-005 |
| CNS-BTT-AMA-001 | Butt (empalme / solder seal) | Amarillo | 12-10 | 4.0-6.0 | **10** | 3 | Tengo | MW-PKO-005 |
| | | | | **Total solder seal** | **120** | | | |

### Tipo **Butt heat-shrink CRIMP** (presión + termocontraíble) — NO solder seal

Kit: [haisstronica 200PCS Heat Shrink Butt](https://www.amazon.com/dp/B07RX6QYX5) (B07RX6QYX5)  
Modelos BHT (núcleo de cobre):

| Código | Modelo | Color | AWG | Núcleo | Cant. aprox* | Umbral | Estado | Ubicación |
|--------|--------|-------|-----|--------|-------------:|-------:|--------|-----------|
| CNS-BHC-BLA-001 | BHT-0.5 | Blanco | 26-24 | 0.5 mm | 50 | 12 | Tengo | MW-PKO-005 |
| CNS-BHC-ROS-001 | BHT-1.25 | Rosa | 22-16 | 0.7 mm | 50 | 12 | Tengo | MW-PKO-005 |
| CNS-BHC-AZU-001 | BHT-2 | Azul | 16-14 | 0.8 mm | 50 | 12 | Tengo | MW-PKO-005 |
| CNS-BHC-AMA-001 | BHT-5.5 | Amarillo | 12-10 | 1.0 mm | 50 | 12 | Tengo | MW-PKO-005 |
| | | | | **Total crimp** | **200** | | | |

\*50 c/u por defecto (200÷4). Si el blister dice otro reparto, avisar.

### Tipo **Butt non-insulated** (empalme crimp sin aislante)

Kit: [haisstronica 50PCS Non-Insulated Butt AWG 12-10](https://www.amazon.com/dp/B0F1N2WH51) (B0F1N2WH51) — cobre estañado, color plata

| Código | Tipo terminal | AWG | Cant. | Umbral | Estado | Ubicación |
|--------|---------------|-----|------:|-------:|--------|-----------|
| CNS-BNI-1210-001 | Butt (empalme / non-insulated crimp) | 12-10 | **50** | 12 | Tengo | MW-PKO-005 |

Recomienda termocontraíble aparte para sellar.

### Tipo **Ring non-insulated** (ojal / crimp sin aislante + heat shrink en kit)

Kit: [haisstronica 100PCS Ring 5/16" + Heat Shrink 3:1 AWG 12-10](https://www.amazon.com/dp/B0F2T8LR9W) (B0F2T8LR9W)  
Contenido: **50** ring (stud 5/16") + **50** tubos termocontraíbles 3:1 rojo/negro. Una línea = 50 terminales (HST incluido en el kit).

| Código | Tipo terminal | Stud | AWG | Cant. | Umbral | Estado | Ubicación |
|--------|---------------|------|-----|------:|-------:|--------|-----------|
| CNS-RNI-516-001 | Ring (ojal / non-insulated crimp) | 5/16" | 12-10 | **50** | 12 | Tengo | MW-PKO-005 |

### Tipo **Fork non-insulated** (horquilla / crimp sin aislante + heat shrink en kit)

Kit: [haisstronica 100PCS Fork #10 + Heat Shrink 3:1 AWG 12-10](https://www.amazon.com/dp/B0F2T498YF) (B0F2T498YF)  
Contenido: **50** fork (stud #10) + **50** tubos termocontraíbles 3:1 rojo/negro. Una línea = 50 terminales (HST incluido en el kit).

| Código | Tipo terminal | Stud | AWG | Cant. | Umbral | Estado | Ubicación |
|--------|---------------|------|-----|------:|-------:|--------|-----------|
| CNS-FNI-10-001 | Fork (horquilla / non-insulated crimp) | #10 | 12-10 | **50** | 12 | Tengo | MW-PKO-005 |

| Prefijo | Significado |
|---------|-------------|
| `CNS-BTT-*` | Butt **solder seal** (calor + estaño) |
| `CNS-BHC-*` | Butt **heat-shrink crimp** (prensar + calor) |
| `CNS-BNI-*` | Butt **non-insulated** (prensar, sin forro; cobre desnudo/estañado) |
| `CNS-RNI-*` | Ring **non-insulated** (ojal + heat shrink en kit) |
| `CNS-FNI-*` | Fork **non-insulated** (horquilla + heat shrink en kit) |
