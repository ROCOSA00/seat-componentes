# PK-16 · Maestro de criticidad y dimensionamiento de mantenimiento

Respaldo técnico, trazable y auditable de la presentación final a Dirección (`PK-16.DATOS.MANTENIMIENTO.V2.pptx`).

> **POWERPOINT FINAL = RESULTADO OFICIAL PRESENTADO**
> **EXCEL MAESTRO = DATOS + CÁLCULOS + CRITERIOS + METODOLOGÍA + TRAZABILIDAD**

## Entregables (`salidas/`)

| Fichero | Uso |
|---|---|
| `PK16_MAESTRO_CRITICIDAD_DIMENSIONAMIENTO.xlsx` | Excel maestro: 21 hojas, fórmulas visibles que leen `12_SUPUESTOS`, controles OK/REVISAR/ERROR, trazabilidad y diferencias |
| `PK16_MODELO_INTERACTIVO.html` | Explorador sin conexión: inventario filtrable, ficha por activo, escenarios, gráficos, trazabilidad, exportación a CSV |
| `DIAGNOSTICO_PK16.md` | Diagnóstico de fuentes, cifras oficiales, hipótesis, pendientes, discrepancias y riesgos |

## Regenerar

```bash
pip install -r modelo/requirements.txt
python modelo/build.py
```

- `build.py` descarga de la release `ARCHIVOS` los ficheros pesados que falten (en `_release/`).
- Después recalcula el Excel con LibreOffice Calc si está instalado; Excel lo recalcula igualmente al abrirlo.

## Estructura del modelo (`modelo/`)

- `pk16_datos.py`: inventario, regla jerárquica de ubicaciones técnicas, depuración y cruces (montaje, cuellos de botella, MEB ECO, averías).
- `pk16_oficial.py`: cifras oficiales del PPT, con su número de diapositiva.
- `pk16_excel.py` y `pk16_xlsx_estilo.py`: generador del Excel maestro.
- `pk16_html.py` y `plantilla_html.html`: generador del HTML interactivo.

> ⚠️ El repositorio es **público** y contiene información industrial interna. Se recomienda pasarlo a privado.
