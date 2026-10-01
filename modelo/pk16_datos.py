# -*- coding: utf-8 -*-
"""
PK-16 · Carga, depuración y cruce de fuentes del Maestro de Criticidad y Dimensionamiento.

Principio: el PowerPoint final es el resultado oficial. Este módulo NO recalcula el resultado
oficial: prepara el inventario activo a activo, los criterios y las evidencias para que el Excel
reproduzca y justifique las cifras presentadas. Lo que no se puede reconstruir queda marcado
(DATO NO DISPONIBLE / PENDIENTE DE VALIDACIÓN / TRAZABILIDAD INCOMPLETA), nunca se rellena.

Todas las reglas de este fichero se describen en la hoja 13_METODOLOGIA del Excel.
"""
from __future__ import annotations

import re
import urllib.request
import warnings
from pathlib import Path

import pandas as pd

warnings.filterwarnings("ignore")

ROOT = Path(__file__).resolve().parent.parent
REL = ROOT / "_release"
RELEASE_URL = "https://github.com/ROCOSA00/seat-componentes/releases/download/ARCHIVOS/"

# --------------------------------------------------------------------------------------
# Fuentes
# --------------------------------------------------------------------------------------
F = {
    "PPT": REL / "PK-16.DATOS.MANTENIMIENTO.V2.pptx",
    "PPT_CONTEXTO": REL / "Presentacion.Datos.mantenimiento.pptx",
    "SAP_RAW": REL / "LISTADO_EQUIPOS_SAP.xlsx",
    "SAP_DEP": ROOT / "LISTADO MAQUINAS SAP.xlsx",
    "FAB": ROOT / "MAESTRO_FABRICA_FINAL.xlsx",
    "DIM": ROOT / "DIMENSIONAMIENTO_MANTENIMIENTO_DATOS_FINAL.xlsx",
    "PROF": ROOT / "MAESTRO_PK16_PROFESIONAL_FINAL 1.xlsx",
    "AVERIAS": REL / "SAP.EXPORT.AVERIAS.LISTADO.xlsx",
    "CB": REL / "CUELLOS_DE_BOTELLA.xlsx",
    "CRIT": REL / "MAESTRO_CRITICIDAD_MAQUINAS.xlsx",
    "INTEGRAL": REL / "MAESTRO_INTEGRAL_CRITICIDAD_MANTENIMIENTO.1.xlsx",
    "MONT_L1": REL / "data.1.xlsx",
    "MONT_L3": REL / "data.2.xlsx",
    "TOP": REL / "TOP.ANALISIS.xlsx",
}


def asegurar_release() -> None:
    """Descarga de la release 'ARCHIVOS' los ficheros pesados que no estén en local."""
    REL.mkdir(exist_ok=True)
    for key, path in F.items():
        if path.parent == REL and not path.exists():
            url = RELEASE_URL + urllib.request.quote(path.name)
            print(f"  descargando {path.name} ...")
            urllib.request.urlretrieve(url, path)


# --------------------------------------------------------------------------------------
# Utilidades
# --------------------------------------------------------------------------------------
def _txt(x) -> str:
    if x is None or (isinstance(x, float) and pd.isna(x)):
        return ""
    return str(x).strip()


def _up(x) -> str:
    return _txt(x).upper()


def inventario_de_equipo(equipo: str) -> str:
    """G00000006786324003 -> 6786324003 ; G000008536525031/1 -> 8536525031/1"""
    e = _txt(equipo)
    if e.startswith("G"):
        e = e[1:]
    return e.lstrip("0") or e


LINEA_ANTIGUA = {"90410": "7016420001", "90510": "7016420002"}  # 'Inventario antiguo' del fichero de montaje


def clave_estacion(texto: str):
    """Clave (línea, nº estación) para cruzar montaje/SAP/paros: 7016424001/10H -> ('7016424001', 10).
    Los inventarios antiguos de línea (90410, 90510) se traducen al nuevo (7016420001, 7016420002)."""
    t = _up(texto)
    m = re.search(r"(7016[34]\d{5}|90410|90510)\s*[/\-]\s*0*(\d{1,4})", t)
    if not m:
        return None
    return (LINEA_ANTIGUA.get(m.group(1), m.group(1)), int(m.group(2)))


# --------------------------------------------------------------------------------------
# Tabla de ubicaciones técnicas raíz (regla jerárquica: gana la raíz más larga/específica)
# --------------------------------------------------------------------------------------
# Campos: raiz, grupo_sap, linea, programa, producto, proceso, grupo_mant, origen_regla
_DEP = "Depuración SAP previa (LISTADO MAQUINAS SAP.xlsx)"
_PROP = "Propuesta del modelo por descripción de ubicación – PENDIENTE DE VALIDACIÓN"
_ESP = "Sub-raíz específica añadida por el modelo (hereda grupo; precisa proceso)"

ROOTS = [
    # --- Fundición y tratamientos
    ("TG00-1-0313", "FUNDICION MQ200/MQ281", "FUNDICION 313", "Convencional", "MQ200/MQ281", "Fundición", "FUNDICIÓN", _DEP),
    ("TG00-1-0311", "PENDIENTE DE CLASIFICACION", "HORNOS FUSIÓN 311", "Convencional", "MQ200/MQ281", "Fundición", "FUNDICIÓN", _PROP),
    ("TG00-1-0635", "TRATAMIENTOS TERMICOS MQ200/MQ281", "TRATAMIENTOS TERMICOS", "Convencional", "MQ200/MQ281", "Tratamientos térmicos", "TRATAMIENTOS TÉRMICOS", _DEP),
    ("TG00-1-0635-0100", "TRATAMIENTOS TERMICOS MQ200/MQ281", "SALA CONTROL", "Convencional", "MQ200/MQ281", "Tratamientos térmicos", "TRATAMIENTOS TÉRMICOS", _DEP),
    ("TG00-1-0635-0101", "TRATAMIENTOS TERMICOS MQ200/MQ281", "GRANALLAS", "Convencional", "MQ200/MQ281", "Tratamientos térmicos", "TRATAMIENTOS TÉRMICOS", _DEP),
    ("TG00-1-0635-0102", "TRATAMIENTOS TERMICOS MQ200/MQ281", "HORNOS HOLCROFT", "Convencional", "MQ200/MQ281", "Tratamientos térmicos", "TRATAMIENTOS TÉRMICOS", _DEP),
    ("TG00-1-0635-0103", "TRATAMIENTOS TERMICOS MQ200/MQ281", "IPSEN L1", "Convencional", "MQ200/MQ281", "Tratamientos térmicos", "TRATAMIENTOS TÉRMICOS", _DEP),
    ("TG00-1-0635-0104", "TRATAMIENTOS TERMICOS MQ200/MQ281", "IPSEN L2", "Convencional", "MQ200/MQ281", "Tratamientos térmicos", "TRATAMIENTOS TÉRMICOS", _DEP),
    ("TG00-1-0635-0105", "TRATAMIENTOS TERMICOS MQ200/MQ281", "NITRURACION", "Convencional", "MQ200/MQ281", "Tratamientos térmicos", "TRATAMIENTOS TÉRMICOS", _DEP),
    ("TG00-1-0635-0106", "TRATAMIENTOS TERMICOS MQ200/MQ281", "EFD HYNGASA", "Convencional", "MQ200/MQ281", "Tratamientos térmicos", "TRATAMIENTOS TÉRMICOS", _DEP),
    ("TG00-1-0635-0107", "TRATAMIENTOS TERMICOS MQ200/MQ281", "BEDA/BAUTERMIC", "Convencional", "MQ200/MQ281", "Tratamientos térmicos", "TRATAMIENTOS TÉRMICOS", _DEP),
    ("TG00-1-0635-0108", "TRATAMIENTOS TERMICOS MQ200/MQ281", "TRAT. CONJUNTO DIFERENCIAL", "Convencional", "MQ200/MQ281", "Tratamientos térmicos", "TRATAMIENTOS TÉRMICOS", _DEP),
    ("TG00-1-0635-0109", "TRATAMIENTOS TERMICOS MQ200/MQ281", "TRATAMIENTOS TERMICOS", "Convencional", "MQ200/MQ281", "Tratamientos térmicos", "TRATAMIENTOS TÉRMICOS", _DEP),
    ("TG00-1-0635-0110", "TRATAMIENTOS TERMICOS MQ200/MQ281", "AICHELIN", "Convencional", "MQ200/MQ281", "Tratamientos térmicos", "TRATAMIENTOS TÉRMICOS", _DEP),
    # --- MQ200 / MQ281 mecanizado
    ("TG00-1-0602-0001", "MQ200", "TORNEADO", "Convencional", "MQ200", "Mecanizado", "MQ200", _DEP),
    ("TG00-1-0602-0003", "MQ281", "TORNEADO", "Convencional", "MQ281", "Mecanizado", "MQ281", _DEP),
    ("TG00-1-0659-0401", "MQ200", "ENGRANAJES NEGROS", "Convencional", "MQ200", "Mecanizado", "MQ200", _DEP),
    ("TG00-1-0659-0402", "MQ200", "ENGRANAJES NEGROS", "Convencional", "MQ200", "Mecanizado", "MQ200", _DEP),
    ("TG00-1-0659-0403", "MQ200", "ENGRANAJES NEGROS", "Convencional", "MQ200", "Mecanizado", "MQ200", _DEP),
    ("TG00-1-0659-0404", "MQ281", "ENGRANAJES NEGROS", "Convencional", "MQ281", "Mecanizado", "MQ281", _DEP),
    ("TG00-1-0659-0405", "MQ281", "ENGRANAJES NEGROS", "Convencional", "MQ281", "Mecanizado", "MQ281", _DEP),
    ("TG00-1-0659-0406", "MQ281", "ENGRANAJES NEGROS", "Convencional", "MQ281", "Mecanizado", "MQ281", _DEP),
    ("TG00-1-0659-0407", "MQ200", "ENGRANAJES NEGROS", "Convencional", "MQ200", "Mecanizado", "MQ200", _DEP),
    ("TG00-1-0662-0009", "MQ200", "EJES BLANCOS", "Convencional", "MQ200", "Mecanizado", "MQ200", _DEP),
    ("TG00-1-0662-0019", "MQ200", "EJES BLANCOS", "Convencional", "MQ200", "Mecanizado", "MQ200", _DEP),
    ("TG00-1-0662-0020", "MQ200", "EJES NEGROS", "Convencional", "MQ200", "Mecanizado", "MQ200", _DEP),
    ("TG00-1-0662-0021", "MQ281", "EJES NEGROS", "Convencional", "MQ281", "Mecanizado", "MQ281", _DEP),
    ("TG00-1-0662-0022", "MQ281", "EJES BLANCOS", "Convencional", "MQ281", "Mecanizado", "MQ281", _DEP),
    ("TG00-1-0663-0005", "MQ200", "ENGRANAJES BLANCOS", "Convencional", "MQ200", "Mecanizado", "MQ200", _DEP),
    ("TG00-1-0663-0006", "MQ200", "ENGRANAJES BLANCOS", "Convencional", "MQ200", "Mecanizado", "MQ200", _DEP),
    ("TG00-1-0663-0201", "MQ200", "ENGRANAJES BLANCOS", "Convencional", "MQ200", "Mecanizado", "MQ200", _DEP),
    ("TG00-1-0663-0202", "MQ200", "ENGRANAJES BLANCOS", "Convencional", "MQ200", "Mecanizado", "MQ200", _DEP),
    ("TG00-1-0663-0203", "MQ200", "ENGRANAJES BLANCOS", "Convencional", "MQ200", "Mecanizado", "MQ200", _DEP),
    ("TG00-1-0663-0204", "MQ200", "ENGRANAJES BLANCOS", "Convencional", "MQ200", "Mecanizado", "MQ200", _DEP),
    ("TG00-1-0663-0205", "MQ281", "ENGRANAJES BLANCOS", "Convencional", "MQ281", "Mecanizado", "MQ281", _DEP),
    ("TG00-1-0663-0206", "MQ281", "ENGRANAJES BLANCOS", "Convencional", "MQ281", "Mecanizado", "MQ281", _DEP),
    ("TG00-1-0663-0207", "MQ281", "ENGRANAJES BLANCOS", "Convencional", "MQ281", "Mecanizado", "MQ281", _DEP),
    ("TG00-1-0664", "PENDIENTE DE CLASIFICACION", "SINCRO: BUJES Y MANGUITOS", "Convencional", "MQ200/MQ281", "Mecanizado", "SINCRO (BUJES/MANGUITOS)", _PROP),
    # --- Montaje convencional
    ("TG00-1-0634", "PENDIENTE DE CLASIFICACION", "MONTAJE CAMBIO (SECCIÓN GENERAL)", "Convencional", "MQ200/MQ281", "Montaje", "MONTAJE COMÚN", _PROP),
    ("TG00-1-0634-0025", "MQ200", "MONTAJE CAMBIO MQ200 LINEA 1", "Convencional", "MQ200", "Montaje", "MQ200", _DEP),
    ("TG00-1-0634-0030", "PENDIENTE DE CLASIFICACION", "MONTAJE CAMBIO MQ200 LINEA 2", "Convencional", "MQ200", "Montaje", "MQ200", _PROP),
    ("TG00-1-0634-0035", "MQ281", "MONTAJE CAMBIO MQ281 LINEA 3", "Convencional", "MQ281", "Montaje", "MQ281", _DEP),
    ("TG00-1-0667", "PENDIENTE DE CLASIFICACION", "MONTAJE KRAUSE 085 / REPARACIÓN", "Convencional", "MQ200", "Montaje", "MQ200", _PROP),
    # --- Mecanizado aluminio y calidad compartida
    ("TG00-1-0665", "MECANIZADO ALUMINIO MQ200/MQ281", "MECANIZADO CAJAS/CARCASAS", "Convencional", "MQ200/MQ281", "Mecanizado", "MECANIZADO ALUMINIO MQ", _DEP),
    ("TG00-1-0665-1100", "CALIDAD / MEDICION COMPARTIDA", "SALA MEDICION 665", "Común", "COMÚN PLANTA", "Calidad y medición", "CALIDAD Y MEDICIÓN", _DEP),
    ("TG00-1-0665-1104", "MECANIZADO ALUMINIO MQ200/MQ281", "MECANIZADO LINEA 4", "Convencional", "MQ200/MQ281", "Mecanizado", "MECANIZADO ALUMINIO MQ", _DEP),
    ("TG00-1-0665-1105", "MECANIZADO ALUMINIO MQ200/MQ281", "MECANIZADO LINEA 3", "Convencional", "MQ200/MQ281", "Mecanizado", "MECANIZADO ALUMINIO MQ", _DEP),
    ("TG00-1-0665-1106", "MECANIZADO ALUMINIO MQ200/MQ281", "MECANIZADO LINEA 5", "Convencional", "MQ200/MQ281", "Mecanizado", "MECANIZADO ALUMINIO MQ", _DEP),
    ("TG00-1-0665-1107", "MECANIZADO ALUMINIO MQ200/MQ281", "MECANIZADO LINEA 6", "Convencional", "MQ200/MQ281", "Mecanizado", "MECANIZADO ALUMINIO MQ", _DEP),
    # --- MEB ECO (nuevos productos)
    ("TG00-1-ALUM", "ALUMINIO MGH/GGH MEB ECO", "SECCION ALUMINIO", "MEB ECO", "ALUMINIO GGH/MGH", "Mecanizado", "ALUMINIO GGH/MGH (MEB)", _DEP),
    ("TG00-1-CDIF", "CONJUNTO DIFERENCIAL MEB ECO", "CONJUNTO DIFERENCIAL", "MEB ECO", "CONJ. DIFERENCIAL", "Mecanizado", "CONJ. DIFERENCIAL (MEB)", _DEP),
    ("TG00-1-CDIF-CONJ", "CONJUNTO DIFERENCIAL MEB ECO", "CDIF MECANIZADO / CELDA", "MEB ECO", "CONJ. DIFERENCIAL", "Mecanizado", "CONJ. DIFERENCIAL (MEB)", _ESP),
    ("TG00-1-CDIF-MONT", "CONJUNTO DIFERENCIAL MEB ECO", "CDIF LÍNEA MONTAJE", "MEB ECO", "CONJ. DIFERENCIAL", "Montaje", "CONJ. DIFERENCIAL (MEB)", _ESP),
    ("TG00-1-CDIF-MEDI", "CONJUNTO DIFERENCIAL MEB ECO", "CDIF MEDICIÓN / AUXILIARES", "MEB ECO", "CONJ. DIFERENCIAL", "Calidad y medición", "CONJ. DIFERENCIAL (MEB)", _ESP),
    ("TG00-1-EBOX", "EBOX MEB ECO", "LINEA EBOX", "MEB ECO", "EBOX", "Montaje", "EBOX (MEB)", _DEP),
    ("TG00-1-KMM0", "KMM MEB ECO", "LINEA KMM", "MEB ECO", "KMM", "Montaje", "KMM (MEB)", _DEP),
    ("TG00-1-MANG", "MANGUETAS MEB ECO", "LINEA MANGUETAS", "MEB ECO", "MANGUETAS", "Montaje", "MANGUETAS (MEB)", _DEP),
    # --- Calidad, taller, instalaciones, logística (propuestas sobre los 516 pendientes)
    ("TG00-1-0450", "PENDIENTE DE CLASIFICACION", "TALLER AUDITORÍA PRODUCTO", "Común", "COMÚN PLANTA", "Calidad y medición", "CALIDAD Y MEDICIÓN", _PROP),
    ("TG00-1-0454", "PENDIENTE DE CLASIFICACION", "CALIDAD LABORATORIO T.T.", "Común", "COMÚN PLANTA", "Calidad y medición", "CALIDAD Y MEDICIÓN", _PROP),
    ("TG00-1-0455", "PENDIENTE DE CLASIFICACION", "CALIDAD SALA MEDICIÓN Y METROLOGÍA", "Común", "COMÚN PLANTA", "Calidad y medición", "CALIDAD Y MEDICIÓN", _PROP),
    ("TG00-1-0458", "PENDIENTE DE CLASIFICACION", "CALIDAD BANCO SARA / PREPARACIÓN", "Común", "COMÚN PLANTA", "Calidad y medición", "CALIDAD Y MEDICIÓN", _PROP),
    ("TG00-1-0638-0002", "PENDIENTE DE CLASIFICACION", "MEDICIÓN FABRICACIÓN", "Común", "COMÚN PLANTA", "Calidad y medición", "CALIDAD Y MEDICIÓN", _PROP),
    ("TG00-1-0638", "PENDIENTE DE CLASIFICACION", "AFILADO HERRAMIENTAS", "Común", "COMÚN PLANTA", "Taller", "TALLER Y MANTENIMIENTO", _PROP),
    ("TG00-1-0298", "PENDIENTE DE CLASIFICACION", "MANTENIMIENTO MOLDES (MATRICERÍA)", "Común", "COMÚN PLANTA", "Taller", "TALLER Y MANTENIMIENTO", _PROP),
    ("TG00-1-0485", "PENDIENTE DE CLASIFICACION", "MANTENIMIENTO MÁQUINAS / CONSTRUCCIÓN MEDIOS", "Común", "COMÚN PLANTA", "Taller", "TALLER Y MANTENIMIENTO", _PROP),
    ("TG00-1-0486", "PENDIENTE DE CLASIFICACION", "MANTENIMIENTO INSTALACIONES", "Común", "COMÚN PLANTA", "Instalaciones", "INSTALACIONES Y SERVICIOS", _PROP),
    ("TG00-1-0481", "PENDIENTE DE CLASIFICACION", "ALMACENES AUXILIAR PRODUCCIÓN", "Común", "COMÚN PLANTA", "Manutención", "MANUTENCIÓN Y ALMACÉN", _PROP),
    ("TG00-1-0668", "PENDIENTE DE CLASIFICACION", "RECAMBIOS", "Común", "COMÚN PLANTA", "Manutención", "MANUTENCIÓN Y ALMACÉN", _PROP),
    ("TG00-1-0661", "PENDIENTE DE CLASIFICACION", "OPTIMIZACIÓN", "Común", "COMÚN PLANTA", "Otros", "PENDIENTE DE CLASIFICACIÓN", _PROP),
]
ROOTS_DF = pd.DataFrame(ROOTS, columns=["RAIZ", "GRUPO_SAP", "LINEA", "PROGRAMA", "PRODUCTO", "PROCESO", "GRUPO_MANT", "ORIGEN_REGLA"])


def clasificar_ubicacion(ubic: str):
    """Regla jerárquica: primero la raíz más específica (más larga), después la superior."""
    u = _up(ubic)
    best = None
    for r in ROOTS:
        raiz = r[0]
        if u == raiz or u.startswith(raiz + "-"):
            if best is None or len(raiz) > len(best[0]):
                best = r
    if best is None:
        return None, "SIN COINCIDENCIA"
    return best, ("EXACTA" if u == best[0] else "HEREDADA")


# --------------------------------------------------------------------------------------
# Tipologías de criticidad (valores base editables en 12_SUPUESTOS)
# --------------------------------------------------------------------------------------
TIPOLOGIAS = [
    # código, nombre, C1, C2, C3, datos suficientes, descripción del criterio
    ("T01", "Estación de montaje en línea serie", 4, 2, 2, "S", "Línea secuencial sin buffer: su fallo para el producto. Atornillados, prensados, test y trazabilidad. Regla acordada: montaje principal = A."),
    ("T02", "Estación standby / reserva / reparación", 1, 1, 0, "S", "Puesto de reserva o reparación: el flujo principal no depende de él."),
    ("T03", "Máquina de mecanizado principal (CNC)", 3, 1, 1, "S", "Máquina productiva de línea/grupo: pérdida importante de proceso. Calidad dimensional. Redundancia no confirmada salvo equivalente en el grupo."),
    ("T04", "Tratamiento térmico", 3, 2, 1, "S", "Horno o instalación de TT: pérdida de proceso, riesgo metalúrgico (calidad) y de atmósferas/gases (seguridad)."),
    ("T05", "Fundición (inyección / fusión)", 3, 2, 1, "S", "Isla de inyección u horno de fusión: pérdida de proceso y riesgo de metal fundido."),
    ("T06", "Periférico de isla de fundición", 2, 1, 1, "S", "Atemperador, lubrificador, campana, dosificador o prensa de rebaba: para la isla, con recuperación por sustitución."),
    ("T07", "Soldadura", 3, 2, 1, "S", "Soldadura (EB/PTR): proceso especial con impacto en calidad y seguridad."),
    ("T08", "Robot / automatización de carga", 2, 1, 1, "S", "Robot o automatización de célula: para la estación; alternativa manual limitada. Riesgo de seguridad de célula."),
    ("T09", "Lavadora industrial", 2, 1, 1, "S", "Lavado de piezas: para el flujo de la estación; impacto en limpieza (calidad)."),
    ("T10", "Calidad / medición / test en línea", 2, 2, 1, "S", "Control 100%, test, fugas, medición en línea o marcaje: riesgo de escape al cliente y trazabilidad."),
    ("T11", "Laboratorio / metrología (fuera de línea)", 1, 1, 0.5, "S", "Equipo de laboratorio o metrología: no para producción directamente; existe alternativa externa/otros equipos."),
    ("T12", "Transporte integrado en línea", 3, 0, 1, "S", "Transporte, satélite, highway o volteador integrado: su fallo detiene el flujo de la línea."),
    ("T13", "Manutención móvil (flota)", 1, 1, 0, "S", "Apiladores, plataformas, tractoras, fregadoras: flota con unidades equivalentes."),
    ("T14", "Instalación central / servicio único", 3, 1, 2, "S", "Central de filtración/emulsión/aspiración/refrigeración que da servicio a varias máquinas: único."),
    ("T15", "Auxiliar de producción", 1, 0, 1, "S", "Ingrávido, aspiración individual, atemperador, bomba, filtro: impacto menor."),
    ("T16", "Instalación / servicio de planta", 2, 1, 1, "S", "Transformadores, baterías de condensadores, clima, puertas, bombas de instalaciones, puentes grúa."),
    ("T17", "Taller / mantenimiento / utillaje", 0, 0, 0, "S", "Equipos de taller, matricería, afilado, construcción de medios y maletas de programación: sin impacto directo en producción."),
    ("T18", "No tipificado – información insuficiente", 1, 0, 1, "N", "Tipo de objeto genérico (OTROS/ESTAC fuera de línea): no hay datos para valorar impacto. Nunca se envía a D."),
    ("T19", "Seguridad / protección de planta", 1, 2, 1, "S", "Contraincendios, centrales de alarma, barreras ópticas: riesgo para las personas."),
]
TIPO_NOMBRE = {t[0]: t[1] for t in TIPOLOGIAS}

_MECANIZADO = {"RECTI", "TORNO", "DENTA", "CMECA", "BROCH", "LAMIN", "AFEIT", "BRUÑI", "DESPU", "TALLA", "MORTA",
               "FRESA", "TALAD", "REBAB", "RODAD", "ESMER", "ACHAF", "DESBA", "ROSCA", "COMPL", "PUNTE", "REMAC", "ATORN"}
_LAB = {"DUROM", "MICRO", "ESPEC", "PULID", "CORTA", "TRONZ", "VISCO", "PHME", "REFRA", "DETER", "CNTRI", "BAÑO",
        "PLACA", "TERMO", "DIOFR", "ELECT", "AMOLA"}
_CAL_LINEA = {"BANCO", "MEDID", "TRIDI", "PERFI", "ESMED", "BRAZO", "CAMAR", "MARCA", "CONTR"}
_TRANSP_LINEA = {"TRANS", "SATEL", "HIGHW", "CINTA", "VOLTE", "ELEVA", "CARRU", "KAVEM"}
_FLOTA = {"APILA", "PLATA", "MTRAC", "TRASP", "FREGA", "COMPA", "PATER", "NORIA"}
_AUX = {"ASPIR", "ATEMP", "POLIP", "FILTR", "CEMUL", "DOSIF", "REFRI", "FRIO", "BOMBA", "DEPOS", "VENTI", "SEPAR",
        "AGITA", "VIRUT", "ENFAR", "GRUPO", "CABIN"}
_INST = {"INSTALAC", "BATER", "ESTRA", "CLIMA", "PUERT", "CUADR", "ARMAR", "SINOP", "TRATA", "TLA", "COMPR", "ILUMI", "PUENT"}
_SEGUR = {"CNINC", "CTRAL", "BARRE"}
_TALLER = {"PLCN", "UTIL", "SIERR", "CEPIL", "ESCUA", "REGRU", "RELIJ", "AFILA", "ENGAS"}
_FUND_PERIF = {"ATEMP", "LUBRI", "ENGRS", "CAMPA", "DOSIF", "PRENS", "ESTUF", "CABIN", "POLIP"}

RX_ROBOT = re.compile(r"\bROBOT|\bROB\.|KUKA|FANUC|\bABB\b|A\.B\.B|COBOT|UNIVERSAL ROBOT|\bIS\d+ ROBOT|\bID\d+ ROBOT")
RX_FLOTA = re.compile(r"APILADOR|CARRETILLA|TRASPAL|PLAT\. ELEVADORA|TRACTORA|FREGADORA")
RX_STANDBY = re.compile(r"STAND[\s\-]?BY|RESERVA|REPARACI|\bS/R\b|\bRETRABAJO|REWORK|DESMONTAJE")
RX_SEGCAL = re.compile(r"TEST|FUGA|ESTANQ|CONTROL|MEDIC|ATORNILL|PRENSAD|SOLDA|LASER|LÁSER|GRABAD|HORNO|TEMPLE|"
                       r"NITRUR|CEMENT|INDUCCI|RAYOS X|VISION|VISIÓN|C[ÁA]MARA|TRIDIM|ALABEO|ABS\b")
RX_CENTRAL = re.compile(r"\bCENTRAL\b|CENTRALIZADA")
# Depuración ampliada (no estaba en LISTADO MAQUINAS SAP): palets, mesas, utillajes, maletas, carros
RX_EXCL_AMPL = re.compile(r"^PALETS?\b|CONJUNTO PALETS|PORTAPIEZA|UTIL DE MONTAJE|UTILLAJE|^MALETA|CONTENEDOR|^CARRO|ALMAC[ÉE]N MUELAS|"
                          r"^MESA DESMONTAJE|SOPORTE VISCOS")
RX_AGRUPADOR = re.compile(r"^GRUPO SECCI|GRUPO [A-D] MEDIOS|INFORMACI[ÓO]N COM[ÚU]N|^LINEA MONTAJE DIFERENCIAL 7016426001$|^LÍNEA MONTAJE DIFERENCIAL")
RX_TEST = re.compile(r"^TEST$")


def tecnologia_de(tp: str, desc: str) -> str:
    d = _up(desc)
    if RX_ROBOT.search(d):
        return "Robótica"
    if tp in _MECANIZADO:
        return "Mecanizado CNC"
    if tp in {"HORNO", "CUBA", "CALEN"}:
        return "Hornos / TT"
    if tp in {"INYEC", "TORRE"}:
        return "Fundición"
    if tp == "SOLDA":
        return "Soldadura"
    if tp == "LAVAD":
        return "Lavado"
    if tp in _CAL_LINEA or tp in _LAB:
        return "Visión / medición"
    if tp in _TRANSP_LINEA or tp in _FLOTA:
        return "Manutención / transporte"
    if tp in _INST or tp in _SEGUR:
        return "Instalaciones / servicios"
    if tp in _AUX or tp in _FUND_PERIF:
        return "Auxiliares"
    if tp == "ESTAC":
        return "Automatización / PLC (estación)"
    return "Otras / sin especificar"


RX_K_MAQ = re.compile(r"TORNO|RECTIF|DENTAD|DESPUNT|FRESAD|BROCHAD|LAMINAD|AFEITAD|BRU[ÑN]ID|TALLAD|MORTAJ|TALADR|"
                     r"GRANALLAD|REBABAD|ESTACI[ÓO]N PRENSADO|CENTRO MECANIZ|COMBINADA|RODADORA|SOLDADORA")
RX_K_CAL = re.compile(r"FUG[OÓ]M|MICROFUG|COMPROBADORA|FOTOGRAM|MEDICI[ÓO]N|TRIDIM|BANCO PRUEBA|ESTANQ|CONTROL|TEST|ATORNILLADOR")
RX_K_AUT = re.compile(r"C[ÉE]LULA|CELDA|P[ÓO]RTICO|MANIPULADOR|AUTOMATIZ|CARGADOR")
RX_K_LAV = re.compile(r"CUBA DE LAVADO|DESENGRASE|LAVADORA|LAVADO")
RX_K_FUND = re.compile(r"CRISOL|EQUIPO VAC[IÍ]O|CALENTADOR|COQUILLA|MOLDE")
RX_K_GRUA = re.compile(r"GR[ÚU]A|PUENTE GR")
RX_K_AUX = re.compile(r"HIDR[AÁ]UL|GRUPO HIDR|ARMARIO|CUADRO EL|DOSIFICADOR|ASPIRACI")
RX_K_TRANSP = re.compile(r"TRANSFERIZADO|\bLOOP\b|BUCLE|CADENA MONTAJE|\bVK\d?\b|TRANSPORTE|HIGHWAY|VOLTEADOR")
RX_K_SEG = re.compile(r"CARRIL DE SEGURIDAD|BARRERA|CONTRA ?INCENDIO|ALARMA")


def tipologia_de(rec: dict) -> str:
    """Tipología de criticidad: primero las reglas de negocio, luego el tipo SAP, después palabras clave."""
    tp = rec["TP"]
    d = _up(rec["DESC"])
    gm = rec["GRUPO_MANT"]
    proc = rec["PROCESO"]
    if rec["MONTAJE_PPAL"] == "S":
        return "T01"
    if _txt(rec.get("ESTADO", "")).startswith("Standby"):
        return "T02"
    if gm == "TALLER Y MANTENIMIENTO":
        return "T17"
    if gm == "MANUTENCIÓN Y ALMACÉN":
        return "T13"
    if tp in _SEGUR or RX_K_SEG.search(d):
        return "T19"
    if gm == "CALIDAD Y MEDICIÓN" or tp in _LAB:
        return "T11"
    if RX_ROBOT.search(d) or tp in {"CARGA", "ISLA"}:
        return "T08"
    if RX_FLOTA.search(d) or tp in _FLOTA:
        return "T13"
    if tp == "SOLDA" or "SOLDADORA" in d:
        return "T07"
    if tp in _MECANIZADO or (tp in {"ROBOT", "OTROS", "ESTAC", "GRANA", ""} and RX_K_MAQ.search(d)) or tp == "GRANA":
        return "T03"
    if tp in {"HORNO", "CUBA", "CALEN"} and proc == "Tratamientos térmicos":
        return "T04"
    if tp in {"HORNO", "TORRE", "INYEC"}:
        return "T05" if proc == "Fundición" else "T04"
    if proc == "Fundición" and (tp in _FUND_PERIF or RX_K_FUND.search(d) or tp in {"CALEN", "OTROS", ""}):
        return "T06"
    if tp == "LAVAD" or RX_K_LAV.search(d):
        return "T09"
    if tp in _CAL_LINEA or RX_K_CAL.search(d):
        return "T10"
    if tp == "PORTI" or RX_K_AUT.search(d):
        return "T08"
    if tp in _TRANSP_LINEA or RX_K_TRANSP.search(d):
        return "T12"
    if RX_CENTRAL.search(d) and (tp in _AUX or tp in {"OTROS", ""}):
        return "T14"
    if RX_K_GRUA.search(d):
        return "T16"
    if tp in _AUX or tp in _FUND_PERIF or RX_K_AUX.search(d):
        return "T15"
    if tp in _INST:
        return "T16"
    if tp in _TALLER:
        return "T17"
    return "T18"


FAMILIA_PPT = {
    # familia del inventario presentado (Presentacion.Datos.mantenimiento.pptx / 07_Contexto_Global)
    "T01": "Montaje", "T02": "Montaje", "T03": "Mecanizado", "T04": "Tratamientos y fundición",
    "T05": "Tratamientos y fundición", "T06": "Auxiliares producción", "T07": "Mecanizado", "T08": "Robots",
    "T09": "Lavadoras", "T10": "Calidad", "T11": "Calidad", "T12": "Transporte instalaciones",
    "T13": "Medios de transporte", "T14": "Auxiliares producción", "T15": "Auxiliares producción",
    "T16": "Instalaciones y servicios (no incluidas en 2.043)", "T17": "Taller y otros (no incluidos en 2.043)",
    "T18": "Sin familia (pendiente)", "T19": "Instalaciones y servicios (no incluidas en 2.043)",
}


# --------------------------------------------------------------------------------------
# Lecturas de fuentes
# --------------------------------------------------------------------------------------
def leer_sap():
    raw = pd.read_excel(F["SAP_RAW"], dtype=str)
    raw["FILA_ORIGEN"] = raw.index + 2
    dep = pd.read_excel(F["SAP_DEP"], sheet_name="01_MAESTRO_ORDENADO", dtype=str)
    elim = pd.read_excel(F["SAP_DEP"], sheet_name="14_ELIMINADOS", dtype=str)
    dep["FILA_DEP"] = dep.index + 2
    elim["FILA_ELIM"] = elim.index + 2
    return raw, dep, elim


def leer_averias():
    a = pd.read_excel(F["AVERIAS"], sheet_name="AVERIAS")
    a["FILA_ORIGEN"] = a.index + 2
    total = a[a["P"].astype(str) == "Total"]
    filtro = a[a["P"].astype(str).str.startswith("Filtros")]
    a["FECHA"] = pd.to_datetime(a["Primera fecha: IniAvería"], errors="coerce")
    es_dato = a["FECHA"].notna() & (a["P"].astype(str) != "Total") & ~a["P"].astype(str).str.startswith("Filtros")
    a = a[es_dato].copy()  # filas sin nº de orden se conservan (exportación Power BI con celdas combinadas)
    a["MIN"] = pd.to_numeric(a["Tiempo de Paro"], errors="coerce").fillna(0)
    meta = {
        "total_fila_min": float(total["Tiempo de Paro"].iloc[0]) if len(total) else None,
        "total_fila_desde": str(total["Primera fecha: IniAvería"].iloc[0])[:10] if len(total) else None,
        "filtro": _txt(filtro["P"].iloc[0]) if len(filtro) else "",
    }
    return a, meta


def leer_montaje_logs():
    out = []
    for key, linea, inv in [("MONT_L1", "MQ200 Línea 1 (Krause)", "7016420001"), ("MONT_L3", "MQ281 Línea 3 (Schmid)", "7016420006")]:
        d = pd.read_excel(F[key])
        d.columns = [c.strip() for c in d.columns]
        d["FILA_ORIGEN"] = d.index + 2
        d["FECHA"] = pd.to_datetime(d["FECHA"], errors="coerce")
        d["EST"] = d["EST"].map(_txt)
        d = d[d["EST"] != "0"].copy()
        d["CP"] = pd.to_numeric(d["CP"], errors="coerce").fillna(0)
        d["LINEA"] = linea
        d["INV_LINEA"] = inv
        d["ARCHIVO"] = F[key].name
        out.append(d)
    return pd.concat(out, ignore_index=True)


def leer_montaje_fichero():
    m = pd.read_excel(F["FAB"], sheet_name="04_Montaje_Completo", header=2, dtype=str)
    m["FILA_ORIGEN"] = m.index + 4
    return m


def leer_meb():
    of = pd.read_excel(F["FAB"], sheet_name="02_MEB_ECO_Oficial", header=2)
    of["FILA_ORIGEN"] = of.index + 4
    det = pd.read_excel(F["FAB"], sheet_name="03_MEB_ECO_Detalle", header=2, dtype=str)
    det["FILA_ORIGEN"] = det.index + 4
    return of, det


def leer_cb():
    cb = pd.read_excel(F["CB"], dtype=str)
    cb["FILA_ORIGEN"] = cb.index + 2
    mapa = pd.read_excel(F["INTEGRAL"], sheet_name="05_Cuellos_Botella", header=2, dtype=str)
    mapa["FILA_MAPA"] = mapa.index + 4
    cb = cb.reset_index(drop=True)
    mapa = mapa.reset_index(drop=True)
    cb["MAPEADO"] = mapa["Mapeado"]
    cb["ACTIVO_DESTINO"] = mapa["Activo destino"]
    cb["METODO_CRUCE"] = mapa["Método cruce"]
    cb["PRODUCTO_CB"] = mapa["Producto"]
    cb["FILA_MAPA"] = mapa["FILA_MAPA"]
    return cb


def leer_crit_previas():
    mc = pd.read_excel(F["CRIT"], sheet_name="Maestro Criticidad", header=2, dtype=str)
    it = pd.read_excel(F["INTEGRAL"], sheet_name="01_Maestro_Activos", header=2, dtype=str)
    return mc, it


def leer_fabrica_maestro():
    m = pd.read_excel(F["FAB"], sheet_name="01_Maestro_Fabrica", header=2, dtype=str)
    m["FILA_ORIGEN"] = m.index + 4
    return m


# --------------------------------------------------------------------------------------
# Construcción del inventario maestro
# --------------------------------------------------------------------------------------
def construir_inventario():
    raw, dep, elim = leer_sap()
    dep_by = dep.set_index("Equipo")
    elim_by = elim.set_index("Equipo")
    ubic_set = set(raw["Ubicac.técnica"].fillna("").str.upper())

    recs = []
    for _, r in raw.iterrows():
        eq = _txt(r["Equipo"])
        ub = _up(r["Ubicac.técnica"])
        desc = _txt(r["Descripción"])
        tp = _txt(r["Tp.objeto"]).replace("TG00_", "")
        root, coinc = clasificar_ubicacion(ub)
        anio = None
        f = _txt(r["Fe.puesta serv."])
        if f[:4].isdigit() and int(f[:4]) > 1900:
            anio = int(f[:4])
        rec = {
            "ID": eq,
            "EQUIPO": eq,
            "INVENTARIO": inventario_de_equipo(eq),
            "PLANO": _txt(r["Campo clasif."]) or "DATO NO DISPONIBLE",
            "UBIC": ub,
            "NIVEL": ub.count("-") + 1 if ub else None,
            "UBIC_PADRE": ub.rsplit("-", 1)[0] if ub.count("-") >= 3 else "",
            "DESC": desc,
            "DESC_UBIC": _txt(r["Descripción3"]),
            "TP": tp,
            "FABRICANTE": _txt(r["Fabricante"]),
            "MODELO": _txt(r["Denomin.tipo"]),
            "ANIO": anio,
            "SAP_ABC": _txt(r["Indicador ABC"]),
            "PTO_TRABAJO": _txt(r["Pto.tbjo.resp."]),
            "UBIC_RAIZ": root[0] if root else "",
            "TIPO_COINC": coinc,
            "GRUPO_SAP": root[1] if root else "PENDIENTE DE CLASIFICACION",
            "LINEA": root[2] if root else "PENDIENTE",
            "PROGRAMA": root[3] if root else "PENDIENTE",
            "PRODUCTO": root[4] if root else "PENDIENTE",
            "PROCESO": root[5] if root else "Pendiente",
            "GRUPO_MANT": root[6] if root else "PENDIENTE DE CLASIFICACIÓN",
            "ORIGEN_REGLA": root[7] if root else "Sin raíz aplicable",
            "FUENTE": F["SAP_RAW"].name,
            "HOJA": "Hoja1",
            "FILA": int(r["FILA_ORIGEN"]),
        }
        # Estado en la depuración SAP previa
        if eq in elim_by.index:
            rec["DEP_PREVIA"] = "ELIMINADO: " + _txt(elim_by.loc[eq, "Motivo eliminación"])
            rec["FILA_DEP"] = f"14_ELIMINADOS!{int(elim_by.loc[eq, 'FILA_ELIM'])}"
        elif eq in dep_by.index:
            g = _txt(dep_by.loc[eq, "Grupo asignado"])
            rec["DEP_PREVIA"] = "DEPURADO: " + g
            rec["FILA_DEP"] = f"01_MAESTRO_ORDENADO!{int(dep_by.loc[eq, 'FILA_DEP'])}"
            # Control: la regla jerárquica del modelo reproduce la asignación previa
            rec["CONTROL_JERARQUIA"] = "OK" if (g == rec["GRUPO_SAP"] or g == "PENDIENTE DE CLASIFICACION") else "DIFERENTE"
        else:
            rec["DEP_PREVIA"] = "NO ENCONTRADO EN DEPURACIÓN"
            rec["FILA_DEP"] = ""
        recs.append(rec)

    inv = pd.DataFrame(recs)
    inv["CONTROL_JERARQUIA"] = inv.get("CONTROL_JERARQUIA").fillna("N/A")

    # Equipo padre por ubicación (solo si la ubicación padre está por debajo de la raíz: estación -> sub-estación)
    by_ubic = inv.groupby("UBIC")["EQUIPO"].apply(list).to_dict()

    def padre(row):
        p = row["UBIC_PADRE"]
        if not p or p == row["UBIC_RAIZ"] or len(p) <= len(row["UBIC_RAIZ"]):
            return ""
        cands = [e for e in by_ubic.get(p, []) if e != row["EQUIPO"]]
        return cands[0] if cands else ""

    inv["EQUIPO_PADRE"] = inv.apply(padre, axis=1)
    return inv


def aplicar_estado(inv: pd.DataFrame, montaje_keys_sap: set) -> pd.DataFrame:
    """Estado del registro y pertenencia al parque principal. Nada se borra."""
    est, enp, mot, tipo = [], [], [], []
    desc_tp = inv[["DESC", "TP"]]
    for _, r in inv.iterrows():
        d = _up(r["DESC"])
        tp = r["TP"]
        dep = r["DEP_PREVIA"]
        if dep.startswith("ELIMINADO"):
            est.append("Excluido – auxiliar / periférico")
            enp.append("N")
            mot.append("Depuración SAP previa: " + dep.split(":", 1)[1].strip())
            tipo.append("Equipo auxiliar / periférico")
            continue
        if RX_TEST.search(_up(r["DESC_UBIC"])) or r["UBIC"].startswith("TG00-1-TG0N"):
            est.append("Excluido – registro de prueba")
            enp.append("N")
            mot.append("Ubicación TEST (TG00-1-TG0N-TEST)")
            tipo.append("Registro no válido")
            continue
        if RX_EXCL_AMPL.search(d) or tp in {"MESA", "UTIL", "PLCN"}:
            est.append("Excluido – auxiliar / utillaje (depuración ampliada)")
            enp.append("N")
            mot.append("Regla ampliada del modelo (palets, mesas, utillajes, maletas, carros) – PENDIENTE DE VALIDACIÓN")
            tipo.append("Utillaje / periférico")
            continue
        if RX_AGRUPADOR.search(d) or tp == "LINEA":
            est.append("Registro agrupador de línea")
            enp.append("N")
            mot.append("Objeto SAP que agrupa la línea; sus estaciones se cuentan individualmente (evita doble conteo)")
            tipo.append("Agrupador")
            continue
        if r["EQUIPO_PADRE"] and not RX_ROBOT.search(d):
            est.append("Componente integrado / sub-estación")
            enp.append("N")
            mot.append(f"Ubicación hija de la estación {r['EQUIPO_PADRE']} – se cuenta en el activo principal")
            tipo.append("Componente integrado")
            continue
        if RX_STANDBY.search(d):
            est.append("Standby / reserva / reparación")
            enp.append("S")
            mot.append("Descripción indica standby, reserva, reparación o retrabajo")
            tipo.append("Equipo en reserva / standby")
            continue
        if r["GRUPO_SAP"] == "PENDIENTE DE CLASIFICACION":
            est.append("Activo – clasificación propuesta (pendiente)")
            enp.append("S")
            mot.append("Sin coincidencia en la depuración previa; grupo propuesto por ubicación técnica")
            tipo.append("")
            continue
        est.append("Activo – parque principal")
        enp.append("S")
        mot.append("Clasificado por ubicación técnica (" + r["TIPO_COINC"].lower() + ")")
        tipo.append("")
    inv["ESTADO"] = est
    inv["EN_PARQUE"] = enp
    inv["MOTIVO_ESTADO"] = mot
    inv["TIPO_ACTIVO"] = tipo
    return inv


# --------------------------------------------------------------------------------------
# Cruces: montaje, cuellos de botella, MEB ECO, averías, criticidades previas
# --------------------------------------------------------------------------------------
LINEAS_CONV = {"MONTAJE CAMBIO MQ200 LINEA 1", "MONTAJE CAMBIO MQ200 LINEA 2", "MONTAJE CAMBIO MQ281 LINEA 3"}
LINEAS_MEB = {"LINEA KMM", "LINEA EBOX", "LINEA MANGUETAS", "CDIF LÍNEA MONTAJE"}
RX_NO_ESTACION = r"TRANSPORTE|PALET|MESA|SAT[ÉE]LITE EWAB|BUCLE|HIGHWAY|ARMARIO|VK SISTEMA|GRUPO [A-D] MEDIOS|^VK\b|CADENA MONTAJE"


def marcar_montaje(inv: pd.DataFrame, mont: pd.DataFrame):
    """Montaje principal (regla acordada = A):
    - Líneas convencionales (MQ200 L1/L2, MQ281 L3): estación ACTIVA del fichero de montaje actualizado
      (MAESTRO_FABRICA_FINAL/04_Montaje_Completo). Las estaciones SAP que no están en ese fichero quedan como
      'posible baja / no en fichero de montaje' fuera del parque (evita doble conteo con el recuento oficial 264).
    - Líneas MEB ECO (nuevas, fichero de montaje incompleto): estación SAP de nivel estación, activa, no componente.
    """
    inv["CLAVE_EST"] = inv.apply(lambda r: clave_estacion(inventario_de_equipo(r["EQUIPO"])) or clave_estacion(r["DESC"]), axis=1)
    inv["MONTAJE_PPAL"] = "N"
    inv["MONT_FILA"] = ""
    mont = mont.copy()
    mont["CLAVE_EST"] = mont["Inventario"].map(clave_estacion)
    mont["EN_SAP"] = "N"
    mont["ID_SAP"] = ""
    activo_parque = (inv["EN_PARQUE"] == "S") & (~inv["ESTADO"].str.startswith("Standby"))
    es_est = inv["TP"].isin(["ESTAC", "OTROS", "ATORN", "CONTR"]) & ~inv["DESC"].str.upper().str.contains(RX_NO_ESTACION, regex=True)
    cand = inv[(inv["PROCESO"] == "Montaje") & (inv["EN_PARQUE"] == "S")]
    by_key = {}
    for i, r in cand.iterrows():
        if r["CLAVE_EST"]:
            by_key.setdefault(r["CLAVE_EST"], []).append(i)
    usados = set()
    for j, m in mont.iterrows():
        k = m["CLAVE_EST"]
        if not k or k not in by_key:
            continue
        idxs = sorted(by_key[k], key=lambda i: (0 if inv.at[i, "TP"] == "ESTAC" else 1, i))
        mont.at[j, "EN_SAP"] = "S"
        mont.at[j, "ID_SAP"] = ", ".join(inv.at[i, "ID"] for i in idxs[:3])
        estado_m = _txt(m["Estado"])
        for i in idxs:
            inv.at[i, "MONT_FILA"] = (inv.at[i, "MONT_FILA"] + "," if inv.at[i, "MONT_FILA"] else "") + str(int(m["FILA_ORIGEN"]))
        if estado_m.startswith("Standby"):
            for i in idxs:
                if inv.at[i, "EN_PARQUE"] == "S" and inv.at[i, "TP"] == "ESTAC":
                    inv.at[i, "ESTADO"] = "Standby / reserva / reparación"
                    inv.at[i, "MOTIVO_ESTADO"] = f"Fichero de montaje fila {int(m['FILA_ORIGEN'])}: Standby/Reparación/Reserva"
                    inv.at[i, "TIPO_ACTIVO"] = "Equipo en reserva / standby"
        if estado_m == "Activo candidato":
            i0 = next((i for i in idxs if i not in usados and activo_parque[i]), None)
            if i0 is not None:
                inv.at[i0, "MONTAJE_PPAL"] = "S"
                usados.add(i0)
    # Cruce adicional por inventario simple (robots y equipos con nº de inventario propio)
    inv_idx = {}
    for i, r in inv.iterrows():
        if r["EN_PARQUE"] == "S" or r["ESTADO"].startswith("Componente"):
            inv_idx.setdefault(r["INVENTARIO"], i)
    for j, m in mont[mont["EN_SAP"] == "N"].iterrows():
        k = _txt(m["Inventario"]).lstrip("0")
        if k and k in inv_idx:
            i = inv_idx[k]
            mont.at[j, "EN_SAP"] = "S"
            mont.at[j, "ID_SAP"] = inv.at[i, "ID"]
            inv.at[i, "MONT_FILA"] = (inv.at[i, "MONT_FILA"] + "," if inv.at[i, "MONT_FILA"] else "") + str(int(m["FILA_ORIGEN"]))
    # MEB: estaciones SAP de nivel estación
    meb = (inv["LINEA"].isin(LINEAS_MEB)) & activo_parque & es_est
    inv.loc[meb, "MONTAJE_PPAL"] = "S"
    # Convencional: estaciones SAP no presentes en el fichero de montaje -> posible baja (fuera del parque)
    conv_sin = (inv["LINEA"].isin(LINEAS_CONV)) & activo_parque & es_est & (inv["MONTAJE_PPAL"] == "N") & (inv["MONT_FILA"] == "")
    inv.loc[conv_sin, "ESTADO"] = "Pendiente – estación SAP no presente en fichero de montaje (posible baja)"
    inv.loc[conv_sin, "EN_PARQUE"] = "N"
    inv.loc[conv_sin, "MOTIVO_ESTADO"] = ("La estación no figura en el fichero de montaje actualizado (264 filas). "
                                          "Se mantiene fuera del recuento para no duplicar; validar si es baja, renombrado o reserva")
    inv.loc[conv_sin, "TIPO_ACTIVO"] = "Posible baja"
    return inv, mont


def registros_montaje_fuera_sap(mont: pd.DataFrame) -> pd.DataFrame:  # noqa: C901
    """Estaciones del fichero de montaje que no se localizan en SAP: se añaden como registros a validar."""
    rows = []
    falta = mont[(mont["EN_SAP"] == "N")]
    for _, r in falta.iterrows():
        inv_ = _txt(r["Inventario"]) or f"SIN-INVENTARIO-FILA{int(r['FILA_ORIGEN'])}"
        estado_src = _txt(r["Estado"])
        standby = "Standby" in estado_src or bool(RX_STANDBY.search(_up(r["Comentario"])))
        comun = "Común" in estado_src
        prod = _txt(r["Producto"])
        prod_map = {"MQ-200": ("MQ200", "MQ200", "Convencional"), "MQ-281": ("MQ281", "MQ281", "Convencional"),
                    "Manguetas": ("MANGUETAS", "MANGUETAS (MEB)", "MEB ECO"), "KMM": ("KMM", "KMM (MEB)", "MEB ECO"),
                    "EBOX": ("EBOX", "EBOX (MEB)", "MEB ECO"), "Conjunto Diferencial": ("CONJ. DIFERENCIAL", "CONJ. DIFERENCIAL (MEB)", "MEB ECO")}
        p, gm, prog = prod_map.get(prod, ("PENDIENTE", "PENDIENTE DE CLASIFICACIÓN", "PENDIENTE"))
        d_up = _up(r["Comentario"]) + " " + _up(r["Modelo"])
        es_estacion = not re.search(RX_NO_ESTACION + r"|ARMARIO|HIDR[AÁ]UL|TRANSFERIZADO|\bLOOP\b|ROBOT|PALET", d_up)
        if comun:
            estado, enp, motivo, tipo = "Registro de datos comunes de línea", "N", "Fila 'Común/datos línea' del fichero de montaje", "Agrupador"
        elif RX_EXCL_AMPL.search(_up(r["Comentario"])) or re.search(r"\bPALET", d_up):
            estado, enp, motivo, tipo = "Excluido – auxiliar / utillaje (depuración ampliada)", "N", "Palet / utillaje en el fichero de montaje", "Utillaje / periférico"
        elif not _txt(r["Inventario"]):
            estado, enp, motivo, tipo = "Pendiente de confirmar (sin inventario)", "N", "Fila del fichero de montaje sin inventario", "Registro pendiente"
        else:
            estado = "Standby / reserva / reparación" if standby else "Activo – no localizado en SAP (pendiente)"
            enp, motivo, tipo = "S", "Estación del fichero de montaje actualizado sin equivalente SAP por (línea, estación)", ("Equipo en reserva / standby" if standby else "Estación")
        rows.append({
            "ID": f"MONT-{inv_}" + ("" if _txt(r["Inventario"]) and (mont["Inventario"] == r["Inventario"]).sum() == 1 else f"-F{int(r['FILA_ORIGEN'])}"), "EQUIPO": "", "INVENTARIO": inv_, "PLANO": "DATO NO DISPONIBLE", "UBIC": "",
            "NIVEL": None, "UBIC_PADRE": "", "EQUIPO_PADRE": "", "DESC": _txt(r["Comentario"]) or _txt(r["Modelo"]),
            "DESC_UBIC": _txt(r["Grupo"]), "TP": "ESTAC", "FABRICANTE": _txt(r["Fabricante"]), "MODELO": _txt(r["Modelo"]),
            "ANIO": int(r["Año"]) if _txt(r["Año"]).isdigit() else None, "SAP_ABC": "", "PTO_TRABAJO": "",
            "UBIC_RAIZ": "", "TIPO_COINC": "FICHERO MONTAJE", "GRUPO_SAP": "", "LINEA": "MONTAJE " + prod,
            "PROGRAMA": prog, "PRODUCTO": p, "PROCESO": "Montaje", "GRUPO_MANT": gm,
            "ORIGEN_REGLA": "Fichero de montaje (MAESTRO_FABRICA_FINAL 04_Montaje_Completo)",
            "FUENTE": F["FAB"].name, "HOJA": "04_Montaje_Completo", "FILA": int(r["FILA_ORIGEN"]),
            "DEP_PREVIA": "NO ESTÁ EN SAP", "FILA_DEP": "", "CONTROL_JERARQUIA": "N/A", "CLAVE_EST": r["CLAVE_EST"],
            "ESTADO": estado, "EN_PARQUE": enp, "MOTIVO_ESTADO": motivo, "TIPO_ACTIVO": tipo,
            "MONTAJE_PPAL": "S" if (enp == "S" and not standby and es_estacion) else "N",
        })
    return pd.DataFrame(rows)


def marcar_cb(inv: pd.DataFrame, cb: pd.DataFrame):
    by_desc = {}
    for i, r in inv.iterrows():
        by_desc.setdefault(_up(r["DESC"]), []).append(i)
    inv["CB"] = "N"
    inv["CB_DETALLE"] = ""
    cb = cb.copy()
    cb["ID_ASIGNADO"] = ""
    cb["ESTADO_CRUCE"] = ""
    for j, c in cb.iterrows():
        dest = _up(c["ACTIVO_DESTINO"])
        idxs = by_desc.get(dest, [])
        idxs_p = [i for i in idxs if inv.at[i, "EN_PARQUE"] == "S"] or idxs
        txt = f"{_txt(c['MODELO'])} / {_txt(c['LINEA'])} / {_txt(c['C.BOTELLA'])} (OEE {float(c['OEE']):.1%})"
        if idxs_p:
            i = idxs_p[0]
            inv.at[i, "CB"] = "S"
            inv.at[i, "CB_DETALLE"] = (inv.at[i, "CB_DETALLE"] + " | " if inv.at[i, "CB_DETALLE"] else "") + txt
            cb.at[j, "ID_ASIGNADO"] = inv.at[i, "ID"]
            cb.at[j, "ESTADO_CRUCE"] = "Enlazado a activo SAP" + (" (descripción repetida en SAP)" if len(idxs) > 1 else "")
        else:
            cb.at[j, "ESTADO_CRUCE"] = "PENDIENTE equivalencia SAP" if _txt(c["MAPEADO"]) != "Si" else "Mapeo sin activo SAP"
    return inv, cb


def registros_cb_pendientes(cb: pd.DataFrame) -> pd.DataFrame:
    rows = []
    for _, c in cb[cb["ID_ASIGNADO"] == ""].iterrows():
        rows.append({
            "ID": f"CB-{int(c['FILA_ORIGEN']):03d}", "EQUIPO": "", "INVENTARIO": "DATO NO DISPONIBLE", "PLANO": "DATO NO DISPONIBLE",
            "UBIC": "", "NIVEL": None, "UBIC_PADRE": "", "EQUIPO_PADRE": "",
            "DESC": f"CUELLO DE BOTELLA {_txt(c['C.BOTELLA'])} ({_txt(c['LINEA'])})", "DESC_UBIC": _txt(c["MODELO"]),
            "TP": "CB", "FABRICANTE": "", "MODELO": "", "ANIO": None, "SAP_ABC": "", "PTO_TRABAJO": "",
            "UBIC_RAIZ": "", "TIPO_COINC": "CUELLOS_DE_BOTELLA", "GRUPO_SAP": "", "LINEA": _txt(c["LINEA"]),
            "PROGRAMA": "MEB ECO" if "E-PRODUCTS" in _up(c["MODELO"]) else "Convencional",
            "PRODUCTO": "ALUMINIO GGH/MGH" if "E-PRODUCTS" in _up(c["MODELO"]) else (
                "MQ200" if "200" in _txt(c["MODELO"]) else ("MQ281" if "281" in _txt(c["MODELO"]) else "MQ200/MQ281")),
            "PROCESO": "Pendiente", "GRUPO_MANT": "CUELLOS SIN EQUIVALENCIA SAP",
            "ORIGEN_REGLA": "Cuello de botella sin equivalencia SAP (no se suma al parque para evitar doble conteo)",
            "FUENTE": F["CB"].name, "HOJA": "Export", "FILA": int(c["FILA_ORIGEN"]),
            "DEP_PREVIA": "NO ESTÁ EN SAP", "FILA_DEP": "", "CONTROL_JERARQUIA": "N/A", "CLAVE_EST": None,
            "ESTADO": "Pendiente de confirmar – CB sin equivalencia SAP", "EN_PARQUE": "N",
            "MOTIVO_ESTADO": "Identificador de línea/isla del fichero de cuellos; puede agrupar varias máquinas ya contadas en SAP",
            "TIPO_ACTIVO": "Registro pendiente", "MONTAJE_PPAL": "N", "CB": "S",
            "CB_DETALLE": f"{_txt(c['MODELO'])} / {_txt(c['LINEA'])} / {_txt(c['C.BOTELLA'])} (OEE {float(c['OEE']):.1%})",
        })
    return pd.DataFrame(rows)


def marcar_meb(inv: pd.DataFrame, meb_det: pd.DataFrame):
    """MEB crítico confirmado en mecanizado: rectificadoras Reishauer de CDIF. Montaje MEB: regla montaje principal."""
    inv["MEB_CRIT"] = "N"
    inv["MEB_DETALLE"] = ""
    rx_reis = (inv["GRUPO_MANT"] == "CONJ. DIFERENCIAL (MEB)") & inv["DESC"].str.upper().str.contains("REISHAUER") & (inv["EN_PARQUE"] == "S")
    inv.loc[rx_reis, "MEB_CRIT"] = "S"
    inv.loc[rx_reis, "MEB_DETALLE"] = "Rectificadora Reishauer CDIF – 4 críticas confirmadas por negocio (02_MEB_ECO_Oficial fila 7); identificación por nombre en SAP pendiente de confirmar"
    mont_meb = (inv["PROGRAMA"] == "MEB ECO") & (inv["MONTAJE_PPAL"] == "S")
    inv.loc[mont_meb, "MEB_DETALLE"] = "Estación de montaje MEB ECO (regla montaje principal)"
    # Cruce del detalle MEB (54 filas) con SAP / montaje
    keys = {}
    for _, r in inv.iterrows():
        k = r.get("CLAVE_EST")
        if k:
            keys.setdefault(k, []).append(r["ID"])
    meb_det = meb_det.copy()
    meb_det["CLAVE_EST"] = meb_det["Inventario / Activo"].map(clave_estacion)
    meb_det["ID_MODELO"] = meb_det["CLAVE_EST"].map(lambda k: ", ".join(keys.get(k, [])[:3]) if k else "")
    return inv, meb_det


def asignar_averias(inv: pd.DataFrame, av: pd.DataFrame, logs: pd.DataFrame, inicio: str):
    av = av.copy()
    av_per = av[av["FECHA"] >= pd.Timestamp(inicio)]
    by_desc = {}
    for i, r in inv.iterrows():
        by_desc.setdefault(_up(r["DESC"]), []).append(i)
    obj_map, amb = {}, {}
    for o in av_per["Denominación de objeto técnico"].dropna().unique():
        idxs = by_desc.get(_up(o), [])
        idxs_p = [i for i in idxs if inv.at[i, "EN_PARQUE"] == "S"] or idxs
        if idxs_p:
            obj_map[o] = inv.at[idxs_p[0], "ID"]
            if len(idxs) > 1:
                amb[o] = len(idxs)
    av["ID_ASIGNADO"] = av["Denominación de objeto técnico"].map(lambda o: obj_map.get(o, ""))
    av["CRUCE"] = av["Denominación de objeto técnico"].map(
        lambda o: ("Ambiguo (descripción repetida en SAP: %d)" % amb[o]) if o in amb else ("Exacto por descripción" if o in obj_map else "SIN ACTIVO SAP"))
    per = av[av["FECHA"] >= pd.Timestamp(inicio)]
    g = per[per["ID_ASIGNADO"] != ""].groupby("ID_ASIGNADO").agg(N=("MIN", "size"), MIN=("MIN", "sum"), MAX=("MIN", "max"))
    inv["AV_N"] = inv["ID"].map(g["N"]).fillna(0).astype(int)
    inv["AV_H"] = (inv["ID"].map(g["MIN"]).fillna(0) / 60).round(2)
    inv["AV_MAX"] = inv["ID"].map(g["MAX"]).fillna(0)
    inv["AV_FUENTE"] = inv["ID"].map(lambda i: F["AVERIAS"].name if i in g.index else "")
    # Montaje: registros de paro de línea (data.1 / data.2) por estación
    logs = logs.copy()
    logs["CLAVE_EST"] = logs.apply(lambda r: (r["INV_LINEA"], int(re.match(r"(\d+)", r["EST"]).group(1)))
                                   if re.match(r"^\d{1,4}[A-Z]?\b|^\d{1,4}[,.]\d", r["EST"]) and not r["EST"].startswith("9041") else None, axis=1)
    keymap = {}
    for _, r in inv.iterrows():
        k = r.get("CLAVE_EST")
        if k and r["EN_PARQUE"] == "S" and k not in keymap:
            keymap[k] = r["ID"]
    logs["ID_ASIGNADO"] = logs["CLAVE_EST"].map(lambda k: keymap.get(k, "") if k else "")
    gl = logs[logs["ID_ASIGNADO"] != ""].groupby("ID_ASIGNADO").agg(N=("CP", "count"), MIN=("CP", "sum"), MAX=("CP", "max"))
    sel = inv["ID"].isin(gl.index) & (inv["AV_N"] == 0)
    inv.loc[sel, "AV_N"] = inv.loc[sel, "ID"].map(gl["N"]).astype(int)
    inv.loc[sel, "AV_H"] = (inv.loc[sel, "ID"].map(gl["MIN"]) / 60).round(2)
    inv.loc[sel, "AV_MAX"] = inv.loc[sel, "ID"].map(gl["MAX"])
    inv.loc[sel, "AV_FUENTE"] = inv.loc[sel, "ID"].map(lambda i: "data.1/data.2 (paros línea montaje 2025, CP en min – unidad a validar)")
    return inv, av, logs


def criticidades_previas(inv: pd.DataFrame):
    mc, it = leer_crit_previas()
    m1 = dict(zip(mc["Máquina"].map(_up), mc["Criticidad"].map(_txt) + " (score " + mc["Score"].map(_txt) + ")"))
    m2 = dict(zip(it["Nombre Activo"].map(_up), it["Criticidad"].map(_txt)))
    inv["CRIT_PREV_MC"] = inv["DESC"].map(lambda d: m1.get(_up(d), ""))
    inv["CRIT_PREV_INT"] = inv["DESC"].map(lambda d: m2.get(_up(d), ""))
    return inv


def completar_tipologia(inv: pd.DataFrame):
    inv["TIPOLOGIA"] = inv.apply(lambda r: tipologia_de(r) if r["EN_PARQUE"] == "S" else "", axis=1)
    inv["TIPOLOGIA_NOMBRE"] = inv["TIPOLOGIA"].map(lambda t: TIPO_NOMBRE.get(t, ""))
    tec_tipo = {"T01": "Automatización / PLC (estación)", "T02": "Automatización / PLC (estación)", "T03": "Mecanizado CNC",
                "T04": "Hornos / TT", "T05": "Fundición", "T06": "Fundición", "T07": "Soldadura", "T09": "Lavado",
                "T10": "Visión / medición", "T11": "Visión / medición", "T12": "Manutención / transporte",
                "T13": "Manutención / transporte", "T14": "Instalaciones / servicios", "T15": "Auxiliares",
                "T16": "Instalaciones / servicios", "T17": "Taller", "T18": "Otras / sin especificar", "T19": "Instalaciones / servicios"}
    inv["TECNOLOGIA"] = inv.apply(lambda r: ("Robótica" if RX_ROBOT.search(_up(r["DESC"])) else "Automatización / PLC (estación)")
                                  if r["TIPOLOGIA"] == "T08" else tec_tipo.get(r["TIPOLOGIA"], tecnologia_de(r["TP"], r["DESC"])), axis=1)
    inv["IND_SEGCAL"] = inv["DESC"].map(lambda d: "S" if RX_SEGCAL.search(_up(d)) else "N")
    # Equivalentes: misma ubicación, mismo tipo SAP y mismo fabricante, dentro del parque
    p = inv[(inv["EN_PARQUE"] == "S") & inv["TIPOLOGIA"].isin(["T03", "T04", "T05", "T08", "T09", "T10", "T07"])]
    cnt = p.groupby(["UBIC", "TP", "FABRICANTE"]).size()
    inv["EQUIV_N"] = inv.apply(lambda r: int(cnt.get((r["UBIC"], r["TP"], r["FABRICANTE"]), 1)) - 1
                               if r["EN_PARQUE"] == "S" and r["FABRICANTE"] and r["UBIC"] else 0, axis=1)
    inv.loc[inv["CB"] == "S", "EQUIV_N"] = inv.loc[inv["CB"] == "S", "EQUIV_N"]
    # Tipo de activo (categoría pedida en el encargo) para los registros del parque
    def tipo(r):
        if r["TIPO_ACTIVO"]:
            return r["TIPO_ACTIVO"]
        t = r["TIPOLOGIA"]
        return {
            "T01": "Estación", "T02": "Equipo en reserva / standby", "T03": "Máquina principal", "T04": "Tratamiento",
            "T05": "Equipo de fundición", "T06": "Equipo de fundición (periférico)", "T07": "Máquina principal",
            "T08": "Robot", "T09": "Lavadora", "T10": "Equipo de calidad y medición", "T11": "Equipo de calidad y medición",
            "T12": "Sistema de manutención", "T13": "Sistema de manutención", "T14": "Instalación",
            "T15": "Equipo auxiliar", "T16": "Instalación", "T17": "Equipo de taller", "T18": "Registro pendiente de confirmar",
            "T19": "Instalación",
        }.get(t, "")
    inv["TIPO_ACTIVO"] = inv.apply(tipo, axis=1)
    inv["FAMILIA_PPT"] = inv.apply(lambda r: FAMILIA_PPT.get(r["TIPOLOGIA"], "") if r["EN_PARQUE"] == "S" else "", axis=1)
    inv["NUEVO_PRODUCTO"] = (inv["PROGRAMA"] == "MEB ECO").map({True: "S", False: "N"})
    # Regla forzada (texto que el Excel usa para fijar la clase)
    def forz(r):
        if r["EN_PARQUE"] != "S":
            return ""
        if r["CB"] == "S":
            return "A – Cuello de botella confirmado"
        if r["MEB_CRIT"] == "S":
            return "A – Activo crítico MEB ECO confirmado"
        if r["MONTAJE_PPAL"] == "S":
            return "A – Montaje principal (regla acordada)"
        return ""
    inv["REGLA_FORZADA"] = inv.apply(forz, axis=1)
    # Estado de validación
    def val(r):
        if r["EN_PARQUE"] != "S":
            return "PENDIENTE DE VALIDACIÓN" if ("Pendiente" in r["ESTADO"] or "ampliada" in r["ESTADO"]) else "EXCLUIDO CON TRAZABILIDAD"
        if r["TIPOLOGIA"] == "T18" or "propuesta" in r["ESTADO"] or "no localizado" in r["ESTADO"]:
            return "PENDIENTE DE VALIDACIÓN"
        if r["MEB_CRIT"] == "S":
            return "PENDIENTE DE VALIDACIÓN"
        if r["CB"] == "S" or r["MONTAJE_PPAL"] == "S":
            return "OK – clase A por regla de negocio (criterios por tipología)"
        return "PROVISIONAL – criterios por tipología"
    inv["VALIDACION"] = inv.apply(val, axis=1)
    return inv


# --------------------------------------------------------------------------------------
# Modelo completo
# --------------------------------------------------------------------------------------
PERIODO_INICIO = "2025-01-01"


def construir_modelo() -> dict:
    asegurar_release()
    inv = construir_inventario()
    mont = leer_montaje_fichero()
    inv = aplicar_estado(inv, set())
    inv, mont = marcar_montaje(inv, mont)
    extra_m = registros_montaje_fuera_sap(mont)
    cb = leer_cb()
    inv, cb = marcar_cb(inv, cb)
    extra_cb = registros_cb_pendientes(cb)
    inv = pd.concat([inv, extra_m, extra_cb], ignore_index=True)
    for c in ["CB", "MONTAJE_PPAL"]:
        inv[c] = inv[c].fillna("N")
    inv["CB_DETALLE"] = inv["CB_DETALLE"].fillna("")
    meb_of, meb_det = leer_meb()
    inv, meb_det = marcar_meb(inv, meb_det)
    av, av_meta = leer_averias()
    logs = leer_montaje_logs()
    inv, av, logs = asignar_averias(inv, av, logs, PERIODO_INICIO)
    inv = criticidades_previas(inv)
    inv = completar_tipologia(inv)
    fab = leer_fabrica_maestro()
    return {"inv": inv, "mont": mont, "cb": cb, "meb_of": meb_of, "meb_det": meb_det, "av": av, "av_meta": av_meta,
            "logs": logs, "fab": fab, "roots": ROOTS_DF}


if __name__ == "__main__":
    M = construir_modelo()
    inv = M["inv"]
    print(inv.shape)
    print(inv["ESTADO"].value_counts().to_string())
    print(inv[inv.EN_PARQUE == "S"]["GRUPO_MANT"].value_counts().to_string())
    print(inv[inv.EN_PARQUE == "S"]["TIPOLOGIA"].value_counts().to_string())
    print("montaje ppal", (inv.MONTAJE_PPAL == "S").sum(), "CB", (inv.CB == "S").sum(), "MEB crit", (inv.MEB_CRIT == "S").sum())
    print("control jerarquía", inv["CONTROL_JERARQUIA"].value_counts().to_string())
    print("averías asignadas", (inv.AV_N > 0).sum(), inv.AV_H.sum())
    print(M["mont"]["EN_SAP"].value_counts().to_string())


# --------------------------------------------------------------------------------------
# Parámetros (los mismos valores se escriben en 12_SUPUESTOS; el Excel recalcula con fórmulas)
# --------------------------------------------------------------------------------------
PARAM = {
    "UMB_A": 8.0, "UMB_B": 6.0, "UMB_C": 3.5,
    "C1_SAPA": 3,
    "C4_H15": 60, "C4_R15": 80, "C4_MAX15": 1500, "C4_H10": 20, "C4_R10": 30, "C4_H05": 5, "C4_R05": 10,
    "C5_Y05": 2000, "C5_Y025": 2012, "C5_UNK": 0.25,
    "PT": {"A": 4, "B": 2.5, "C": 1.5, "D": 1},
    "RATIO": {"A": 12, "B": 20, "C": 35, "D": 60},
}
TIPO_VAL = {t[0]: {"C1": t[2], "C2": t[3], "C3": t[4], "SUF": t[5]} for t in TIPOLOGIAS}


def puntuar(inv: pd.DataFrame, P: dict = PARAM) -> pd.DataFrame:
    """Réplica en Python de las fórmulas de 03_SCORE_CRITICIDAD (sirve para el HTML y para validar el Excel)."""
    s = inv[inv["EN_PARQUE"] == "S"].copy()

    def c(r):
        tv = TIPO_VAL[r["TIPOLOGIA"]]
        forz = r["CB"] == "S" or r["MONTAJE_PPAL"] == "S" or r["MEB_CRIT"] == "S"
        c1 = 4 if forz else max(tv["C1"], P["C1_SAPA"] if r["SAP_ABC"] == "A" else 0)
        c2 = 2 if (r["MONTAJE_PPAL"] == "S" or r["IND_SEGCAL"] == "S") else tv["C2"]
        c3 = 2 if forz else (min(tv["C3"], 0.5) if (r["EQUIV_N"] >= 1 and tv["C3"] > 0.5) else tv["C3"])
        h, n, mx = r["AV_H"], r["AV_N"], r["AV_MAX"]
        c4 = 1.5 if (h >= P["C4_H15"] or n >= P["C4_R15"] or mx >= P["C4_MAX15"]) else (
            1 if (h >= P["C4_H10"] or n >= P["C4_R10"]) else (0.5 if (h >= P["C4_H05"] or n >= P["C4_R05"]) else 0))
        y = r["ANIO"]
        c5 = P["C5_UNK"] if (y is None or pd.isna(y)) else (0.5 if y <= P["C5_Y05"] else (0.25 if y <= P["C5_Y025"] else 0))
        sc = c1 + c2 + c3 + c4 + c5
        cl = "A" if sc >= P["UMB_A"] else "B" if sc >= P["UMB_B"] else "C" if sc >= P["UMB_C"] else "D"
        suf = tv["SUF"] if "propuesta" not in r["ESTADO"] else "N"
        fin = r["REGLA_FORZADA"][:1] if r["REGLA_FORZADA"] else ("C" if (cl == "D" and suf != "S") else cl)
        prov = "N" if r["REGLA_FORZADA"] else ("S" if suf != "S" else "N")
        return pd.Series({"C1": c1, "C2": c2, "C3": c3, "C4": c4, "C5": c5, "SCORE": sc, "CLASE_SCORE": cl,
                          "DATOS_SUF": suf, "CLASE": fin, "PROVISIONAL": prov,
                          "PUNTOS": P["PT"][fin], "RATIO": P["RATIO"][fin], "FTE": 1 / P["RATIO"][fin]})

    return pd.concat([s, s.apply(c, axis=1)], axis=1)


if __name__ == "__main__":
    S = puntuar(M["inv"])
    print(S["CLASE"].value_counts().sort_index().to_string())
    print("provisional", (S.PROVISIONAL == "S").sum(), "FTE", S.FTE.sum())
    print(pd.crosstab(S["GRUPO_MANT"], S["CLASE"]).to_string())
    print(S["C4"].value_counts().to_string())
