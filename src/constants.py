"""Catálogos y constantes de negocio compartidas por toda la app."""

PLAZAS = ["Puebla", "Huauchinango", "Zacatlán"]

PROGRAMAS = [
    "Lic. Administración",
    "Lic. Contaduría",
    "Lic. Derecho",
    "Bachillerato",
    "Maestría Derecho Penal",
]

ESTRATEGIAS = [f"E{i}" for i in range(1, 11)]
CANALES = ESTRATEGIAS + ["Orgánico", "Recomendación histórica"]

ETAPAS = [
    "Prospecto",
    "Contactado",
    "Cita agendada",
    "Inscripción provisional",
    "Alumno",
    "Perdido",
]

MOTIVOS_PERDIDA = ["Precio", "Horario", "Competencia", "No responde", "Otro"]

RESPONSABLES = [
    "Alfonso", "Control Escolar", "Administración y Finanzas", "Seguimiento a prospectos",
    "Dirección Académica", "Mercadotecnia", "CEA Dirección", "Atención en sede",
    "Coordinación Sierra (ZAC/HUA)",
]

APROBADORES = ["Administración y Finanzas", "Dirección Académica", "CEA Dirección"]

TIPOS_DESCUENTO = [
    "Beca %",
    "Condonación adeudo",
    "Mensualidad 50% referente",
    "Inscripción 50% referido",
    "Inscripción gratis campaña",
]

VIGENCIA_DESCUENTO = ["Única vez", "Permanente"]

KR_IDS = [
    "KR1.1", "KR1.2", "KR1.3", "KR1.4",
    "KR2.1", "KR2.2", "KR2.3", "KR2.4", "KR2.5",
    "KR3.1", "KR3.2", "KR3.3", "KR3.4", "KR3.5",
    "KR4.1", "KR4.2", "KR4.3",
]

KPIS_MANUALES = [
    "Reseñas Google #",
    "Reseñas Google promedio",
    "NPS",
    "NPS tasa de respuesta %",
    "Mediana 1ª respuesta (min)",
    "Citas agendadas",
    "Citas que asistieron",
    "Gasto pauta MXN",
    "Bajas del mes",
]

NOMBRES_ESTRATEGIAS = {
    "E1": "Servicio que vende",
    "E2": "Regresa y termina",
    "E3": "Tubería interna",
    "E4": "Referidos",
    "E5": "Convenios",
    "E6": "Publicidad pagada",
    "E7": "Reseñas Google",
    "E8": "Radar de empresas",
    "E9": "Campo en la Sierra",
    "E10": "Ajuste +$150",
}

# Metas de negocio
ALUMNOS_HOY_DEMO = 110
PUNTO_EQUILIBRIO = 207
META_ALUMNOS = 250

# Semáforo AE del portafolio (Tab 2 y Tab 5)
AE_UMBRAL_VERDE = 0.90
AE_UMBRAL_AMARILLO = 0.85

# Umbral de CAC de E6
CAC_E6_UMBRAL = 2000

# Ventana de inscripción / vigencia de promociones
CIERRE_PROMOCIONES = "2026-08-28"

# Colores del semáforo estricto (verde >=100%, amarillo 70-99%, rojo <70%)
COLOR_VERDE = "#2E7D32"
COLOR_AMARILLO = "#F9A825"
COLOR_ROJO = "#C62828"


def color_semaforo(pct_cumplimiento: float) -> str:
    """pct_cumplimiento en fracción (1.0 = 100%)."""
    if pct_cumplimiento >= 1.0:
        return COLOR_VERDE
    if pct_cumplimiento >= 0.70:
        return COLOR_AMARILLO
    return COLOR_ROJO


def color_ae(ae: float) -> str:
    if ae >= AE_UMBRAL_VERDE:
        return COLOR_VERDE
    if ae >= AE_UMBRAL_AMARILLO:
        return COLOR_AMARILLO
    return COLOR_ROJO


COLOR_GRIS = "#9E9E9E"


def color_score_okr(avance: float) -> str:
    """Semáforo de score de KR: verde >=0.7, amarillo 0.3-0.69, rojo <0.3."""
    if avance >= 0.7:
        return COLOR_VERDE
    if avance >= 0.3:
        return COLOR_AMARILLO
    return COLOR_ROJO
