# -*- coding: utf-8 -*-
"""
Cifras OFICIALES del PowerPoint final presentado a Dirección (PK-16.DATOS.MANTENIMIENTO.V2.pptx).
Se transcriben literalmente con su diapositiva. El modelo nunca las modifica: las reproduce y, si no
puede, registra la diferencia en 17_DIFERENCIAS.
"""

PPT = "PK-16.DATOS.MANTENIMIENTO.V2.pptx"
PPT_CTX = "Presentacion.Datos.mantenimiento.pptx"

OF = {
    # Diapositiva 4 – Distribución A/B/C/D y cálculo base
    "PARQUE": (2043, "4", "Total activos (gráfico + tabla) · también diapositiva 8 'INSTALACIONES TOTALES 2043'"),
    "N_A": (418, "4", "Activos clase A · también diapositiva 8 'activos críticos 418'"),
    "N_B": (275, "4", "Activos clase B"),
    "N_C": (1200, "4", "Activos clase C"),
    "N_D": (150, "4", "Activos clase D"),
    "RATIO_A": (12, "4", "Ratio aplicado A (activos/FTE)"),
    "RATIO_B": (20, "4", "Ratio aplicado B"),
    "RATIO_C": (35, "4", "Ratio aplicado C"),
    "RATIO_D": (60, "4", "Ratio aplicado D"),
    "FTE_A": (34.83, "4", "FTE base clase A"),
    "FTE_B": (13.75, "4", "FTE base clase B"),
    "FTE_C": (34.29, "4", "FTE base clase C"),
    "FTE_D": (2.50, "4", "FTE base clase D"),
    "FTE_BASE": (85.37, "4", "FTE base total por criticidad"),
    # Diapositiva 3 – rangos de ratio
    "RANGO_A": ("8 a 12", "3", "Ratio A activos/FTE"),
    "RANGO_B": ("12 a 20", "3", "Ratio B"),
    "RANGO_C": ("20 a 35", "3", "Ratio C"),
    "RANGO_D": ("35 a 60", "3", "Ratio D"),
    # Diapositiva 4 – método de score
    "UMBRALES": ("A ≥ 8 | B 6–7,9 | C 3,5–5,9 | D < 3,5", "4", "Rangos de clase"),
    "CRITERIOS": ("Producción 0–4 · Seguridad/calidad 0–2 · Cuello de botella 0–2 · Averías y paro 2025→actual 0–1,5 · Obsolescencia/repuestos 0–0,5", "4", "De dónde sale la clase"),
    # Diapositiva 5 – waterfall y GAP
    "W_BASE": (85, "5", "FTE base criticidad (redondeado)"),
    "W_PM": (10, "5", "Preventivos"),
    "W_FORM": (3, "5", "Formación (estimación)"),
    "W_SAB": (2, "5", "Sábados / libranzas"),
    "W_COB_MIN": (13, "5", "Cobertura 15–25 %: +13"),
    "W_COB_MAX": (21, "5", "Cobertura 15–25 %: +21"),
    "W_NEC_MIN": (113, "5", "Necesidad total 113–121 (mínimo)"),
    "W_NEC_MAX": (121, "5", "Necesidad total 113–121 (máximo)"),
    "S5_REF_MIN": (115, "5", "Indicador 'base' 115 / 125 (mínimo)"),
    "S5_REF_MAX": (125, "5", "Indicador 'base' 115 / 125 (máximo)"),
    "S5_GAP": (33, "5", "PERSONAL / GAP +33"),
    "PLANTILLA": (80, "5 y 7", "Plantilla actual ~80 personas"),
    "COB_ACTUAL": (0.80, "5 y 7", "Cobertura actual estimada de un 80 %"),
    # Diapositiva 7 – referencia ponderada y reparto por grupos
    "S7_REF_MIN": (115, "7", "Referencia por carga ponderada 115-125 (mínimo)"),
    "S7_REF_MAX": (125, "7", "Referencia por carga ponderada 115-125 (máximo)"),
    "S7_GAP": (34, "7", "PERSONAL / GAP +34"),
    # Diapositiva 8
    "S8_PUNTOS_PM": (15000, "8", "Puntos de preventivo yoY 15.000"),
}

# Diapositiva 7: FTE MANTENIMIENTO POR GRUPO (gráfico; tabla incrustada con valores escritos a mano)
GRUPOS_S7 = [
    ("MQ200", 11.35), ("MQ281", 9.25), ("EBOX", 4.66), ("MANGUETAS", 5.21), ("KMM", 4.22),
    ("CONJ. DIF.", 2.75), ("ALUMINIO", 6.15), ("SERV. COMUNES", 0.0),
]
S7_NOTA = "Diapositiva 7: rótulo 'PENDIENTE REALIZAR · SIN MQ-281'. SERV. COMUNES = fórmula '=D7-D6' sobre celdas vacías (0)."

# DIMENSIONAMIENTO_MANTENIMIENTO_DATOS_FINAL.xlsx · 03_Secciones (filas 6-12)
SECCIONES_DIM = [
    # grupo, A, B, C, D, A_MEB, fila
    ("MQ200", 117, 19, 6, 29, 0, 6),
    ("MQ281", 65, 36, 38, 57, 0, 7),
    ("EBOX", 14, 0, 0, 0, 14, 8),
    ("MANGUETAS", 17, 0, 0, 0, 17, 9),
    ("KMM", 15, 0, 0, 0, 15, 10),
    ("CONJ. DIF.", 14, 0, 5, 17, 14, 11),
    ("ALUMINIO", 15, 3, 0, 0, 15, 12),
]

# Contexto de parque presentado (Presentacion.Datos.mantenimiento.pptx diap. 1 y 07_Contexto_Global)
FAMILIAS_PPT = [
    ("Mecanizado", 430, PPT_CTX + " diap. 1", "430 Máquinas de mecanizado (+90 fabricantes)"),
    ("Robots", 92, PPT_CTX + " diap. 1", "92 Robots (Kuka, Fanuc, ABB, Universal)"),
    ("Montaje", 264, "MAESTRO_FABRICA_FINAL.xlsx 07_Contexto_Global fila 6", "264 filas del fichero de montaje (la presentación de contexto decía 183 estaciones)"),
    ("Tratamientos y fundición", 58, PPT_CTX + " diap. 1", "58 instalaciones TT y fundición"),
    ("Calidad", 290, PPT_CTX + " diap. 1", "290 puestos de calidad"),
    ("Lavadoras", 74, PPT_CTX + " diap. 1", "74 lavadoras automáticas"),
    ("Transporte instalaciones", 355, PPT_CTX + " diap. 1", "355 instalaciones de transporte (Ewab, Bernais…)"),
    ("Medios de transporte", 386, PPT_CTX + " diap. 1", "386 medios de transporte (carros, apiladores, potros…)"),
    ("Auxiliares producción", 94, PPT_CTX + " diap. 1", "94 instalaciones auxiliares (ingrávidos, atemperadores, aspiradores…)"),
]

# Carga declarada (Presentacion.Datos.mantenimiento.pptx diap. 1 · 07_Contexto_Global)
CARGA_DECLARADA = [
    ("Planes de mantenimiento preventivo", 9000, "planes", "+ de 9.000 planes"),
    ("Puntos de preventivo", 50000, "puntos", "50.000 puntos a realizar"),
    ("Avisos de intervención en máquina (2024)", 10000, "avisos/año", "+ de 10.000 avisos"),
    ("Intervenciones TPM (2024)", 3000, "intervenciones/año", "+ de 3.000 intervenciones"),
    ("Horas de máquina parada", 10000, "h", "10.000 h de máquina parada"),
]

REFERENCIAS_EXTERNAS = [
    ("SMRP Best Practices – Maintenance & Reliability Body of Knowledge / Metrics & Guidelines",
     "https://smrp.org/Learning-Resources/SMRP-Library/Best-Practices-Metrics-Guidelines/",
     "Métricas estandarizadas y benchmarking de mantenimiento y fiabilidad (trabajo planificado, preventivo/correctivo, recursos)."),
    ("NASA NPR 8831.2F – Facilities Maintenance and Operations Management",
     "https://nodis3.gsfc.nasa.gov/displayDir.cfm?t=NPR&c=8831&s=2F",
     "Gestión del mantenimiento por requisitos de trabajo (work requirements), órdenes, planificación anual, criticidad y dotación de recursos (staffing)."),
]
