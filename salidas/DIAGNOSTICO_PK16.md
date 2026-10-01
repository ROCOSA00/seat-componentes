# PK-16 · Diagnóstico de fuentes y trazabilidad

Este documento acompaña al Excel maestro (`PK16_MAESTRO_CRITICIDAD_DIMENSIONAMIENTO.xlsx`) y al HTML interactivo (`PK16_MODELO_INTERACTIVO.html`).

**Principio de trabajo:** el PowerPoint final es el resultado oficial. El modelo lo reproduce con fórmulas y documenta, sin corregirlas, las diferencias que encuentra en las fuentes.

## 1. Archivos encontrados

**En la raíz del repositorio**

| Archivo | Qué aporta | Confianza |
|---|---|---|
| `DIMENSIONAMIENTO_MANTENIMIENTO_DATOS_FINAL.xlsx` | Supuestos (80 personas, 55 h, 1.450 h, 10 × 45 / 220, 60.000 / 6.300, 15–25 %), cálculo FTE y reparto por secciones | Media. La hoja `00_Resumen` está desactualizada |
| `MAESTRO_FABRICA_FINAL.xlsx` | Origen de A = 418; 75 críticos MEB; 264 filas de montaje; histórico de paros; 129 cuellos de botella; familias del parque (2.043) | Media. Mezcla filas agregadas e individuales |
| `LISTADO MAQUINAS SAP.xlsx` | Depuración previa: 2.078 equipos clasificados por ubicación raíz, de ellos 516 sin coincidencia; 779 eliminados con su motivo | Alta |
| `MAESTRO_PK16_PROFESIONAL_FINAL 1.xlsx` | Consolidación anterior de los dos primeros libros | Media. Es una versión anterior |

**En la release `ARCHIVOS`**

| Archivo | Qué aporta | Confianza |
|---|---|---|
| `PK-16.DATOS.MANTENIMIENTO.V2.pptx` | **Resultado oficial** (8 diapositivas, modificado el 10/09/2026) | Oficial |
| `Presentacion.Datos.mantenimiento.pptx` | Contexto: familias del parque, 9.000 planes, 50.000 puntos PM, 10.000 avisos, 3.000 intervenciones TPM | Media |
| `LISTADO_EQUIPOS_SAP.xlsx` | Exportación SAP completa (2.857 equipos) | Alta |
| `SAP.EXPORT.AVERIAS.LISTADO.xlsx` | 9.386 averías del 03/01/2025 al 27/07/2026; tiempo de paro en minutos; excluye montaje y TPM | Alta |
| `CUELLOS_DE_BOTELLA.xlsx` | 129 cuellos de botella con su OEE | Alta |
| `MAESTRO_CRITICIDAD_MAQUINAS.xlsx` | Criticidad anterior, puntos equivalentes (4 / 2,5 / 1,5 / 1) y rangos de ratio | Media |
| `MAESTRO_INTEGRAL_CRITICIDAD_MANTENIMIENTO.1.xlsx` | Asignación de cada cuello de botella a una máquina; escenarios de PM y TPM | Media |
| `data.1.xlsx` / `data.2.xlsx` | Paros por estación de las líneas de montaje 1 (MQ200) y 3 (MQ281) en 2025 | Media. Unidad «CP» sin confirmar |
| `TOP.ANALISIS.xlsx` | Exportación de Power BI vacía | — |

**No disponibles:**
- MAQRAIZ.
- `ANALISIS_PAROS_MAQUINAS_V2 3.xlsx`.
- `LINEAS DE MONTAJE GENERALES…`, aunque su contenido está en `04_Montaje_Completo`.
- Planes SAP PM con sus horas.
- Plantilla por especialidad y turno.
- Calendario de sábados y festivos.
- Plan de formación.
- Inventario de instalaciones generales (utilities).

## 2. PowerPoint oficial

Se ha tomado como oficial `PK-16.DATOS.MANTENIMIENTO.V2.pptx`:
- La diapositiva 6 está vacía.
- La diapositiva 8 conserva marcadores «XX».
- Los gráficos llevan tablas incrustadas con valores escritos a mano, sin vínculo a ningún Excel.

## 3. Cifras oficiales y si el modelo las reproduce

| KPI | Valor en el PPT | Diap. | Modelo | Estado |
|---|---|---|---|---|
| Parque | 2.043 | 4 · 8 | 2.043 | OK. Es una suma de familias, no un inventario fila a fila |
| A / B / C / D | 418 / 275 / 1.200 / 150 | 4 | Igual | A reconstruido; B/C/D con **trazabilidad incompleta** |
| Ratios | 12 / 20 / 35 / 60 (rangos 8–12, 12–20, 20–35, 35–60) | 3 · 4 | Igual | OK |
| FTE base | 85,37 | 4 | 85,37 | OK |
| Preventivos | +10 | 5 | 9,52 | OK (redondeo) |
| Formación | +3 | 5 | 3,03 | OK (redondeo) |
| Sábados / libranzas | +2 | 5 | 2,05 | OK (redondeo) |
| Cobertura 15–25 % | +13 / +21 | 5 | 15,0 / 25,0 | **Diferencia**: la base de cálculo no es la misma (ver §7) |
| Necesidad total | 115–125 (diap. 7) · 113–121 (diap. 5) | 5 · 7 | 114,97–124,97 | OK con la diap. 7; diferencia con la diap. 5 |
| Plantilla | ~80 | 5 · 7 | 80 | Pendiente de validar |
| GAP | +34 (diap. 7) · +33 (diap. 5) | 5 · 7 | 35,0 / 45,0 | **Diferencia** (ver §7) |
| Cobertura actual | 80 % | 5 · 7 | 80,0 % | OK |
| FTE por grupo | MQ200 11,35 · MQ281 9,25 · CONJ. DIF. 2,75 | 7 | Igual | OK |
| FTE por grupo | EBOX 4,66 · MANGUETAS 5,21 · KMM 4,22 · ALUMINIO 6,15 | 7 | 2,33 · 2,83 · 2,50 · 2,65 | **Diferencia** |
| Puntos de preventivo «yoY» | 15.000 | 8 | 50.000 usados en el cálculo | **Diferencia** |

## 4. Datos confirmados

- **Inventario SAP.** 2.857 equipos; la depuración previa excluye 779 periféricos.
- **Clasificación por ubicación.** La regla jerárquica del modelo (primero la ubicación más específica, después la superior) reproduce el 100 % de la depuración previa (2.078 registros).
- **Origen de A = 418.** Es el recuento de filas con criticidad A en `MAESTRO_FABRICA_FINAL/01_Maestro_Fabrica`: 129 cuellos de botella + 55 «A2 Alto paro» + 185 estaciones de montaje + 43 de montaje MEB + 6 filas MEB oficial.
- **Versión anterior.** El 68,1 FTE que aparece en `00_Resumen` sale de A = 418 con B/C/D de la fuente (119 / 189 / 1.317).
- **Tiempo de paro.** Está en minutos (comprobado con el histórico de paros).
- **MEB ECO.** 75 críticos confirmados por negocio: 56 estaciones de montaje y 19 máquinas de mecanizado.

## 5. Hipótesis (marcadas como supuesto en `12_SUPUESTOS`)

- **Horas productivas por FTE.** 1.450 h.
- **Formación.** 55 h por persona y año.
- **Sábados.** 10 personas × 45 sábados ÷ 220 jornadas.
- **Productividad de preventivo y avisos.** 6.300 unidades por FTE, unos 13,8 min por unidad. Puntos de preventivo y avisos no consumen el mismo tiempo.
- **Cobertura organizativa.** Entre el 15 % y el 25 %.
- **Valores de criterio por tipología.** Los criterios C1, C2 y C3 de cada activo salen de su tipología, no de un análisis de impacto activo a activo.
- **Umbrales de averías y obsolescencia.** Los del criterio C4 (averías) y C5 (obsolescencia por año de puesta en servicio) son regla del modelo.
- **Exclusiones ampliadas.** 231 palets, mesas, utillajes y maletas excluidos por el modelo.
- **Grupo propuesto.** 468 activos que no tenían coincidencia en la depuración previa se asignan a un grupo según la descripción de su ubicación.

## 6. Pendientes

Hay 27 pendientes en `15_PENDIENTES_VALIDACION`. Los de impacto alto son:
- Conciliar el parque oficial (2.043) con el inventario fila a fila (1.745).
- Documentar el origen del reparto B/C/D.
- Validar redundancia e impacto productivo de cada activo.
- Obtener las horas reales de preventivo.
- Confirmar los puntos de preventivo (15.000 frente a 50.000).
- Turnos, especialidades y plantilla de referencia.
- Reparto definitivo por grupos (la diap. 7 dice «PENDIENTE REALIZAR» y deja 41,8 FTE sin asignar).

## 7. Discrepancias (17 en `17_DIFERENCIAS`)

1. **Base de la cobertura.** El PPT aplica el 15–25 % sobre tres bases distintas:
   - Solo la base por criticidad → 113–121 y GAP +33 (diap. 5).
   - Base + preventivos → 114–124 y GAP +34 (Excel de dimensionamiento).
   - Subtotal → 115–125 (método de la diap. 2 y cifra de la diap. 7).

   El selector `P_COB_BASE` permite ver las tres; por defecto se usa el subtotal.
2. **A = 418 mezcla granularidades.**
   - Las 6 filas MEB representan 75 activos.
   - Los 129 cuellos corresponden a 94 máquinas SAP (107 entradas) + 22 entradas sin equivalencia (líneas o islas).
   - Hay máquinas contadas a la vez como cuello y como «A2».
3. **B/C/D declarados.** No existe un registro activo a activo que dé 275 / 1.200 / 150.
4. **Grupos MEB de la diap. 7.** Los valores no coinciden con su fuente, y la fuente suma dos veces los críticos MEB (columnas A y A_MEB).
5. **Cierre «RESTO PLANTA».** En la fuente, 51,7 FTE se asignan por diferencia para que cuadre el total.
6. **Horas de avería 2025→.** 29.627 h en el Excel de dimensionamiento frente a 7.234 h recalculadas desde el export SAP: el filtro sumaba el histórico 2022–2026 completo.
7. **Puntos de preventivo.** 15.000 en la diap. 8 frente a 50.000 usados en el cálculo.
8. **`00_Resumen` desactualizado.** Muestra 68,1 / 114 / 124 y la conclusión «+14 a +22», que contradice sus propias cifras (34–44).

## 8. Arquitectura entregada

- **`modelo/build.py`.** Regenera todo desde las fuentes; descarga de la release los ficheros que falten.
- **`pk16_datos.py`.** Inventario, jerarquía de ubicaciones, depuración y cruces (montaje, cuellos de botella, MEB, averías).
- **`pk16_excel.py`.** 21 hojas. Todas las fórmulas leen `12_SUPUESTOS` mediante nombres definidos (`P_…`). Recalculado con LibreOffice: 60.000 fórmulas y 0 errores.
- **`pk16_html.py` + `plantilla_html.html`.** HTML sin dependencias externas, con los mismos datos del Excel recalculado.

## 9. Formato de salida

- **Excel maestro.** Es la fuente técnica auditable: fórmulas visibles, supuestos centralizados, controles, trazabilidad y diferencias.
- **HTML interactivo.** Para explorar: buscador y filtros del inventario, ficha por activo con sus cinco criterios, escenarios editables, gráficos de capas y clases, trazabilidad y exportación a CSV.
- **No se genera un PowerPoint nuevo.** El oficial ya existe, y una segunda versión crearía un resultado alternativo.

## 10. Riesgos de trazabilidad

- **Parque y ratios.** El denominador (2.043) y el reparto B/C/D no se pueden recorrer hasta el activo. Pesan el 60 % del FTE base.
- **Tablas de los gráficos.** Están escritas a mano dentro del PPT, sin vínculo a ningún Excel.
- **Fuentes de criterio por activo.** No hay MAQRAIZ, AMFE, matriz de redundancia ni inventario de PLC/CNC. Por eso 451 activos quedan como «C provisional».
- **Cruce de cuellos de botella.** 22 entradas (líneas o islas) sin equivalencia en SAP.
- **Cruce de averías.** 279 averías sin activo SAP asignado.
- **Confidencialidad.** El repositorio es **público** y contiene datos industriales internos. Se recomienda pasarlo a privado.
