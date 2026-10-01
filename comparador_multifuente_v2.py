"""
Comparador multi-fuente - Revision economica de cotizaciones/licitaciones CAPEX
=================================================================================
Evalua una partida cotizada contra TODAS las fuentes de precio disponibles y
regresa un veredicto por fuente + un veredicto combinado.

Fuentes ya integradas (datos reales, con URL/origen verificable):
  1. NL historico (Tabulador Homologado NL) - contratos reales de obra publica
     de Nuevo Leon 2021-2025 (SIASI / Open Contracting Partnership).
  2. CDMX gobierno (Tabulador CDMX) - tabulador oficial de precios unitarios
     del Gobierno de la Ciudad de Mexico, edicion 2026.

Fuentes con interfaz lista pero SIN datos todavia (no se inventa nada):
  3. Historico Ragasa - requiere que Ragasa proporcione su propio historico
     de compras/ordenes de compra (concepto, unidad, precio, fecha).
  4. Comparacion entre proveedores de la misma licitacion - requiere las
     propuestas economicas reales de los demas participantes en la licitacion
     que se esta revisando.

Uso:
    from comparador_multifuente_v2 import ComparadorMultiFuente
    c = ComparadorMultiFuente("Base_Precios_Unitarios_NL_CDMX.xlsx")
    veredicto = c.evaluar(
        descripcion="Suministro y colocacion de acero de refuerzo en losas, varilla corrugada",
        unidad="KG",
        precio_cotizado=30.0,
    )
    print(veredicto)
"""
import re
import unicodedata
import pandas as pd
from rapidfuzz import fuzz, process

import ajuste_inflacion as _inflacion
from ajuste_inflacion import factor_ajuste, ajustar_precio, FUENTE as FUENTE_INPC
import revision_ia as _revision_ia
# Nota: ETIQUETA_ACTUAL se lee como _inflacion.ETIQUETA_ACTUAL (no se importa
# el nombre suelto) porque refrescar_nivel_actual() lo actualiza en tiempo de
# ejecucion; si se importara como valor suelto aqui, este modulo se quedaria
# con la etiqueta vieja aunque el numero (via factor_ajuste) si se actualice.

LEADING_CODE_RE = re.compile(r'^\s*[\d.]{1,15}\s+')
WS_RE = re.compile(r'\s+')

# Abreviaturas comunes en cotizaciones/licitaciones de CAPEX en Mexico.
# Se expanden a su forma completa ANTES de comparar texto, para que
# "SUM E INST DE..." y "SUMINISTRO E INSTALACION DE..." puntuen igual de
# bien que si estuvieran escritas exactamente igual. Se aplica por token
# completo (no substrings), asi que es seguro agregar mas entradas aqui
# sin miedo a corromper palabras que ya son correctas.
ABREVIATURAS_DESCRIPCION = {
    'SUM': 'SUMINISTRO', 'SUMIN': 'SUMINISTRO', 'SUMINIST': 'SUMINISTRO',
    'INST': 'INSTALACION', 'INSTAL': 'INSTALACION', 'INSTALAC': 'INSTALACION',
    'COL': 'COLOCACION', 'COLOC': 'COLOCACION',
    'TRANSP': 'TRANSPORTE',
    'INC': 'INCLUYE', 'INCL': 'INCLUYE', 'INCLUY': 'INCLUYE',
    'DEMOL': 'DEMOLICION',
    'EXC': 'EXCAVACION', 'EXCAV': 'EXCAVACION',
    'CONC': 'CONCRETO',
    'ACAB': 'ACABADO',
    'ELEC': 'ELECTRICO', 'ELECT': 'ELECTRICO',
    'ESTRUC': 'ESTRUCTURA', 'ESTR': 'ESTRUCTURA',
    'MOT': 'MOTOR',
    'MTO': 'MANTENIMIENTO',
    'IMPERM': 'IMPERMEABILIZACION',
    'PREF': 'PREFABRICADO',
    'GALV': 'GALVANIZADO',
}

# Prefijos tipo "C/", "S/", "P/" (con/sin/para) que aparecen pegados a la
# siguiente palabra en muchas cotizaciones ("C/CUÑA", "S/ANDAMIO"). Se
# expanden con regex de limite de palabra ANTES de quitar signos, porque
# el filtro de caracteres de normalize_text conserva la diagonal.
_PREFIJOS_SLASH = [
    (re.compile(r'\bC/'), 'CON '),
    (re.compile(r'\bS/'), 'SIN '),
    (re.compile(r'\bP/'), 'PARA '),
]

# Renglones genericos que aparecen sueltos en muchas cotizaciones (un
# "catch-all" de mano de obra o materiales varios, sin especificar de
# que material se trata) y que NO deben compararse contra la base de
# precios: como no describen un material especifico, cualquier "match"
# que el scorer encuentre para ellos es coincidencia de palabras
# genericas (mano/obra/material/equipo...) y no significa que sea el
# mismo concepto. Se comparan ya normalizados (mayusculas, sin acentos).
CONCEPTOS_GENERICOS_SIN_MATERIAL = {
    'MANO DE OBRA',
    'MANO DE OBRA ESPECIALIZADA',
    'MATERIALES',
    'MATERIALES Y MANO DE OBRA',
    'HERRAMIENTA',
    'HERRAMIENTAS',
    'EQUIPO',
    'VARIOS',
    'CONCEPTOS VARIOS',
    'TRABAJOS VARIOS',
}


def _strip_accents(s: str) -> str:
    # NFKD tambien "desarma" caracteres de compatibilidad como M² -> M2 y
    # M³ -> M3, ademas de quitar acentos. Por eso se usa tanto en texto
    # como en unidades: sin esto, "M2" y "M²" nunca hacian match aunque
    # fueran exactamente la misma unidad.
    return ''.join(c for c in unicodedata.normalize('NFKD', s) if not unicodedata.combining(c))


def _expandir_abreviaturas(texto: str) -> str:
    tokens = texto.split(' ')
    tokens = [ABREVIATURAS_DESCRIPCION.get(tok, tok) for tok in tokens]
    return ' '.join(tokens)


def normalize_text(raw: str) -> str:
    if not isinstance(raw, str):
        return ''
    t = raw.strip()
    t = LEADING_CODE_RE.sub('', t)
    t = _strip_accents(t.upper())
    for patron, reemplazo in _PREFIJOS_SLASH:
        t = patron.sub(reemplazo, t)
    t = re.sub(r'[^A-Z0-9%./\-\s]', ' ', t)
    t = WS_RE.sub(' ', t).strip()
    return _expandir_abreviaturas(t)


# Redaccion informal de obra (como la escribe un contratista pequeno) vs.
# redaccion de catalogo (como vienen las bases de NL y CDMX). Ejemplo real:
# "ESTUCO EN BARDA TANTO POR FUERA COMO POR DENTRO" no encontraba ningun
# match aunque la base si trae "APLANADO DE ESTUCO ... EN MUROS", porque
# las palabras de relleno (barda, por fuera, por dentro) hundian el score.
# Estas reglas generan una SEGUNDA consulta "de catalogo" que se busca
# ademas de la original; se conserva la que puntue mejor. Nunca reemplazan
# el texto que ve el usuario.
_RELLENO_OBRA = [
    r'\bTANTO POR FUERA COMO POR DENTRO\b',
    r'\bPOR FUERA Y POR DENTRO\b',
    r'\bPOR (AMBAS|LAS DOS|AMBOS) (CARAS|LADOS)\b',
    r'\bEN (LA )?(BARDA|AZOTEA)\b',
    r'\bPERIMETRO DE PLACA NUEVA\b',
    r'\bPERIMETR(O|AL)\b',
    r'\bA \d+(\.\d+)? ?(CM|CMS|M|MT|MTS) DE ALTURA\b',
    r'\bPARA SOSTENER (EL )?MURO( DE BLOCK)?\b',
    r'\bPARA AMARRAR (LA )?BARDA\b',
    r'\bBARDA\b',
]

_SINONIMOS_OBRA = [
    # Block "del numero 4/6/8" = espesor de 10/15/20 cm.
    (r'\bBLOCK (DEL |DE )?(NUM|NUMERO|NO|#)\.? ?6\b', 'BLOCK 15 X 20 X 40'),
    (r'\bBLOCK (DEL |DE )?(NUM|NUMERO|NO|#)\.? ?4\b', 'BLOCK 10 X 20 X 40'),
    (r'\bBLOCK (DEL |DE )?(NUM|NUMERO|NO|#)\.? ?8\b', 'BLOCK 20 X 20 X 40'),
    # "Columnas" de barda de block = castillos.
    (r'\bCOLUMNAS? (PARA AMARRAR|DE AMARRE|DE BARDA|DE CONFINAMIENTO)\b', 'CASTILLO'),
    (r'\bARMADO CON ARMEX\b', 'ARMEX'),
    (r'\bCONCRETO HECHO EN OBRA\b', 'CONCRETO HECHO EN OBRA'),
]


def consulta_catalogo(t: str) -> str:
    """Version "de catalogo" de un concepto ya normalizado (ver arriba)."""
    q = f' {t} '
    for patron, reemplazo in _SINONIMOS_OBRA:
        q = re.sub(patron, reemplazo, q)
    if re.search(r'\bESTUCO\b', q) and not re.search(r'\bAPLANADO\b', q):
        q = re.sub(r'\bESTUCO\b', 'APLANADO DE ESTUCO', q)
        if not re.search(r'\bMUROS?\b', q):
            q += ' EN MUROS'
    for patron in _RELLENO_OBRA:
        q = re.sub(patron, ' ', q)
    q = re.sub(r'(?<![0-9])[.,](?![0-9])', ' ', q)
    q = re.sub(r'\b(Y|DE|EN|CON|DEL|LA|EL)( (Y|DE|EN|CON|DEL|LA|EL))+\b', r'\1', q)
    q = WS_RE.sub(' ', q).strip()
    # Una consulta de una sola palabra ("CASTILLO") da 100% con
    # token_set_ratio contra cualquier concepto que la contenga (ej.
    # "ANCLAJE DE CASTILLO A COLUMNA"), asi que no aporta: se descarta.
    palabras = [p for p in q.split() if len(p) > 2 and p not in ('CON', 'DEL', 'LOS', 'LAS')]
    if len(palabras) < 2:
        return t
    return q


# Búsqueda por MATERIAL + TRABAJO + UNIDAD: el nombre que usa el
# contratista ("columnas para amarrar barda", "cerramiento") no es el que
# usan los tabuladores ("castillo de concreto", "dala/cadena de concreto").
# Cada regla agrega consultas de catálogo adicionales; la comparación sigue
# exigiendo la MISMA unidad y se rechaza cualquier referencia de otro
# elemento o material (ver _referencia_compatible).
_EQUIVALENCIAS_TRABAJO = [
    (r'\b(CERRAMIENTO|DALA|CADENA)S?\b', [
        'DALA DE CONCRETO ARMEX', 'CADENA DE CONCRETO ARMEX',
        'DALA DE CONCRETO REFORZADO CON VARILLAS Y ESTRIBOS',
        'CADENA DE CONCRETO REFORZADO CON VARILLAS Y ESTRIBOS',
    ]),
    (r'\b(CASTILLO|COLUMNAS? (PARA AMARRAR|DE AMARRE|DE BARDA|DE CONFINAMIENTO))', [
        'CASTILLO DE CONCRETO ARMEX', 'CASTILLO DE CONCRETO AHOGADO CON VARILLA',
        'CASTILLO DE CONCRETO 15 X 15',
    ]),
    (r'\bESTUCO\b', ['APLANADO DE ESTUCO EN MUROS', 'ESTUCO EN MUROS']),
    (r'\bMURO DE BLOCK\b', ['MURO DE BLOCK DE CONCRETO']),
]


def consultas_catalogo(t: str) -> list:
    """Todas las consultas a probar para un concepto normalizado: el texto
    original, su versión de catálogo y las equivalencias de trabajo."""
    consultas = [t, consulta_catalogo(t)]
    base = consulta_catalogo(t)
    espesor = re.search(r'\b\d+ X \d+( X \d+)?\b', base)
    try:
        from validacion_referencias import elemento_principal
        elemento = elemento_principal(t)
    except Exception:
        elemento = None
    for patron, alternas in _EQUIVALENCIAS_TRABAJO:
        # Solo el elemento PRINCIPAL de la partida: "cerramiento para
        # sostener muro de block" se busca como dala, no como muro.
        clave = {'cerramiento': 'CERRAMIENTO', 'castillo': 'CASTILLO', 'muro': 'MURO'}.get(elemento)
        if elemento and clave and clave not in patron and 'ESTUCO' not in patron:
            continue
        if re.search(patron, f' {t} '):
            for alterna in alternas:
                consultas.append(f'{alterna} {espesor.group(0)}' if espesor and 'BLOCK' in alterna else alterna)
    return [c for c in dict.fromkeys(consultas) if c]


def _referencia_compatible(cotizado: str, referencia: str) -> bool:
    """Rechaza materiales o elementos distintos (castillo vs cerramiento,
    columna metálica vs castillo de concreto...)."""
    try:
        from validacion_referencias import alcance_distinto
    except Exception:
        return True
    return not alcance_distinto(cotizado, referencia)


# Unidades equivalentes que en las bases reales aparecen escritas de
# formas distintas para la MISMA unidad fisica (confirmado revisando
# Base_Precios_Unitarios_Nuevo_Leon_REAL.xlsx: "pieza" aparece como PZA,
# PZ, PZAS, PZS, PIEZA, PIEZAS, PIESZA (typo) y UN por separado; "metro
# lineal" aparece como M, ML, MTS, MT, METRO, METROS y MI). Como el
# comparador exige unidad identica antes de comparar texto, esta
# fragmentacion hacia que miles de renglones nunca se cruzaran entre si
# aunque fueran el mismo material. Las llaves y valores ya deben venir en
# mayusculas/sin acentos (se aplica despues de _strip_accents).
UNIDAD_CANONICA = {
    # pieza
    'PZA': 'PZA', 'PZ': 'PZA', 'PZAS': 'PZA', 'PZS': 'PZA',
    'PIEZA': 'PZA', 'PIEZAS': 'PZA', 'PIESZA': 'PZA',
    'UN': 'PZA', 'UNIDAD': 'PZA', 'UNIDADES': 'PZA',
    # metro lineal (distinto de M2/M3, que ya se resuelven solos porque
    # _strip_accents convierte M² -> M2 y M³ -> M3)
    'M': 'ML', 'ML': 'ML', 'MTS': 'ML', 'MT': 'ML', 'MI': 'ML',
    'METRO': 'ML', 'METROS': 'ML', 'MTSL': 'ML', 'MTL': 'ML', 'MTO': 'ML',
    # peso
    'KG': 'KG', 'KGS': 'KG', 'KILOGRAMO': 'KG', 'KILOGRAMOS': 'KG', 'KILO': 'KG',
    'TON': 'TON', 'TONS': 'TON', 'TONELADA': 'TON', 'TONELADAS': 'TON',
    # volumen
    'L': 'LT', 'LT': 'LT', 'LTS': 'LT', 'LITRO': 'LT', 'LITROS': 'LT',
    # conjuntos
    'JGO': 'JGO', 'JUEGO': 'JGO', 'JUEGOS': 'JGO',
    'LOTE': 'LOTE', 'LOTES': 'LOTE',
    'GLOBAL': 'GLOBAL', 'GLB': 'GLOBAL', 'GL': 'GLOBAL',
    'KIT': 'KIT', 'KITS': 'KIT',
    # otros abreviados comunes en la base real
    'TRAMO': 'TRAMO', 'TMO': 'TRAMO',
    'BOBINA': 'BOBINA', 'BOB': 'BOBINA',
    'SERVICIO': 'SERVICIO', 'SERV': 'SERVICIO',
    'ESTUDIO': 'ESTUDIO', 'EST': 'ESTUDIO',
    'ROLLO': 'ROLLO', 'ROLLOS': 'ROLLO',
    'VIAJE': 'VIAJE', 'VIAJES': 'VIAJE',
    'JORNAL': 'JORNAL', 'JOR': 'JORNAL',
}


def normalize_unit(u: str) -> str:
    if not isinstance(u, str) or not u.strip():
        return 'SIN_UNIDAD'
    t = _strip_accents(u.strip().upper()).replace('.', '')
    t = WS_RE.sub(' ', t).strip()
    return UNIDAD_CANONICA.get(t, t)


def extraer_amperaje(texto: str):
    texto = normalize_text(texto)
    patrones = [
        r'\b\d+\s*X\s*(\d+(?:\.\d+)?)\s*(?:A|AMPS?\.?|AMPERES?)\b',
        r'\b(\d+(?:\.\d+)?)\s*(?:A|AMPS?\.?|AMPERES?)\b',
    ]
    for patron in patrones:
        coincidencia = re.search(patron, texto)
        if coincidencia:
            return float(coincidencia.group(1))
    return None


def extraer_polos(texto: str):
    texto = normalize_text(texto)
    patrones = [
        r'\b([1-4])\s*X\s*\d+(?:\.\d+)?\s*(?:A|AMPS?\.?|AMPERES?)\b',
        r'\b([1-4])\s*POLOS?\b',
    ]
    for patron in patrones:
        coincidencia = re.search(patron, texto)
        if coincidencia:
            return int(coincidencia.group(1))
    return None


def extraer_calibre_awg(texto: str):
    texto = normalize_text(texto)
    patrones = [
        r'\bCAL(?:IBRE)?\.?\s*(\d{1,3})\b',
        r'\b\d+\s*X\s*(\d{1,3})\s*AWG\b',
        r'\b(\d{1,3})\s*AWG\b',
    ]
    for patron in patrones:
        coincidencia = re.search(patron, texto)
        if coincidencia:
            return int(coincidencia.group(1))
    return None


def extraer_numero_conductores(texto: str):
    """
    Extrae la cantidad de conductores en expresiones como:
    3X14 AWG, 4 X 12 AWG.
    """
    texto = normalize_text(texto)
    coincidencia = re.search(
        r'\b(\d+)\s*X\s*\d{1,3}\s*AWG\b',
        texto,
    )
    if coincidencia:
        return int(coincidencia.group(1))
    return None


def extraer_diametro_mm(texto: str):
    texto = normalize_text(texto)
    coincidencia = re.search(
        r'\b(\d+(?:\.\d+)?)\s*MM\b',
        texto,
    )
    if coincidencia:
        return float(coincidencia.group(1))
    return None


def extraer_diametro_pulgadas(texto: str):
    texto = str(texto).upper()
    coincidencia = re.search(
        r'\b(\d+(?:\.\d+)?|\d+/\d+)\s*(?:"|PULG)',
        texto,
    )
    if coincidencia:
        return coincidencia.group(1)
    return None


def detectar_familia_producto(texto: str):
    texto = normalize_text(texto)
    familias = {
        "interruptor": [
            "INTERRUPTOR",
            "TERMOMAGNETICO",
            "DISYUNTOR",
            "BREAKER",
        ],
        "contacto": [
            "CONTACTO",
            "TOMACORRIENTE",
            "RECEPTACULO",
        ],
        "apagador": [
            "APAGADOR",
            "INTERRUPTOR DE LUZ",
        ],
        "caja_electrica": [
            "CAJA",
            "CHALUPA",
            "REGISTRO",
        ],
        "conector": [
            "CONECTOR",
            "GLANDULA",
        ],
        "cable": [
            "CABLE",
            "CONDUCTOR",
            "ALAMBRE",
            "ARMOFLEX",
        ],
        "tuberia_electrica": [
            "TUBO",
            "CONDUIT",
            "POLIFLEX",
            "LICUATITE",
        ],
        "condulet": [
            "CONDULET",
        ],
        "lampara": [
            "LAMPARA",
            "LUMINARIA",
            "REFLECTOR",
        ],
        "sensor": [
            "SENSOR",
        ],
        "relevador": [
            "RELEVADOR",
            "RELE",
        ],
        "ventilador": [
            "VENTILADOR",
            "EXTRACTOR",
        ],
        "canaleta": [
            "CANALETA",
        ],
        "poste": [
            "POSTE",
        ],
        "soporte": [
            "SOPORTE",
            "BASE METALICA",
            "PTR",
        ],
        "limpieza": [
            "LIMPIEZA",
            "ASEO",
        ],
        # Familias de obra civil. Se agregaron porque al hacer el
        # matching de texto mas flexible (scorer combinado + sinonimos de
        # unidad) crecio el riesgo de que, por ejemplo, "acero de refuerzo
        # en LOSAS" hiciera match con "acero de refuerzo en MUROS": el
        # texto es casi identico salvo el elemento estructural, y antes de
        # este cambio no habia ninguna validacion que lo detectara (la
        # unica familia de compatibilidad tecnica que existia era para
        # materiales electricos).
        "concreto": [
            "CONCRETO",
            "HORMIGON",
        ],
        "acero_refuerzo": [
            "ACERO DE REFUERZO",
            "VARILLA CORRUGADA",
            "VARILLA",
        ],
        "excavacion": [
            "EXCAVACION",
        ],
        "piso": [
            "PISO",
            "PORCELANATO",
            "PORCELANICO",
            "LOSETA",
            "AZULEJO",
        ],
        "pintura": [
            "PINTURA",
        ],
        "impermeabilizacion": [
            "IMPERMEABIL",
        ],
        "tablaroca": [
            "TABLAROCA",
            "DRYWALL",
        ],
    }
    for familia, palabras in familias.items():
        for palabra in palabras:
            if palabra in texto:
                return familia
    return None


def detectar_subtipo_cable(texto: str):
    texto = normalize_text(texto)
    if "COBRE DESNUDO" in texto or "CABLE DESNUDO" in texto:
        return "COBRE_DESNUDO"
    if "USO RUDO" in texto:
        return "USO_RUDO"
    if "ARMOFLEX" in texto:
        return "ARMOFLEX"
    if "THW-LS" in texto or "THW LS" in texto:
        return "THW_LS"
    return None


def detectar_subtipo_tuberia(texto: str):
    texto = normalize_text(texto)
    if "LICUATITE" in texto or "LIQUIDTIGHT" in texto:
        return "LICUATITE"
    if "POLIFLEX" in texto:
        return "POLIFLEX"
    if "CONDUIT" in texto:
        return "CONDUIT"
    if (
        "PVC SANITARIO" in texto
        or "TUBO SANITARIO" in texto
        or "ALBANAL" in texto
    ):
        return "PVC_SANITARIO"
    if "PVC" in texto:
        return "PVC"
    return None


def detectar_subtipo_caja(texto: str):
    texto = normalize_text(texto)
    if "ALBANAL" in texto:
        return "REGISTRO_ALBANAL"
    if "OCTAGONAL" in texto:
        return "CAJA_OCTAGONAL"
    if "CAJA FSCA" in texto:
        return "CAJA_FSCA"
    if re.search(r"\bCAJA\s+FS\b", texto):
        return "CAJA_FS"
    if "CAJA REGISTRO" in texto:
        return "REGISTRO_ELECTRICO"
    if "CHALUPA" in texto:
        return "CHALUPA"
    if "GALVANIZ" in texto:
        return "CAJA_GALVANIZADA"
    return None


def detectar_elemento_estructural(texto: str):
    """Para las familias 'concreto' y 'acero_refuerzo': en que elemento
    estructural va el material (losa, muro, zapata, columna, trabe/viga,
    cimentacion, castillo, dala, banqueta/guarnicion). Sin esto, "acero de
    refuerzo en LOSAS" y "acero de refuerzo en MUROS" son textualmente
    casi identicos y el matching por similitud los confundiria."""
    texto = normalize_text(texto)
    elementos = {
        "LOSA": ["LOSA"],
        "MURO": ["MURO"],
        "ZAPATA": ["ZAPATA"],
        "COLUMNA": ["COLUMNA"],
        "TRABE_VIGA": ["TRABE", "VIGA"],
        "CIMENTACION": ["CIMENTACION", "CIMIENTO"],
        "CASTILLO": ["CASTILLO"],
        "DALA": ["DALA"],
        "BANQUETA_GUARNICION": ["BANQUETA", "GUARNICION"],
        "PISO_FIRME": ["FIRME", "PISO"],
    }
    for elemento, palabras in elementos.items():
        for palabra in palabras:
            if palabra in texto:
                return elemento
    return None


def extraer_medida_caja(texto: str):
    texto = normalize_text(texto)
    coincidencia = re.search(
        r"\b(\d+(?:\.\d+)?)\s*X\s*(\d+(?:\.\d+)?)"
        r"(?:\s*X\s*(\d+(?:\.\d+)?))?\b",
        texto,
    )
    if not coincidencia:
        return None
    medidas = [
        valor
        for valor in coincidencia.groups()
        if valor is not None
    ]
    return "X".join(medidas)


def validar_compatibilidad_tecnica(
    descripcion_entrada: str,
    descripcion_referencia: str,
):
    familia_entrada = detectar_familia_producto(
        descripcion_entrada
    )
    familia_referencia = detectar_familia_producto(
        descripcion_referencia
    )
    texto_entrada = normalize_text(descripcion_entrada)
    texto_referencia = normalize_text(descripcion_referencia)
    entrada_es_registro_electrico = (
        "REGISTRO" in texto_entrada
        and (
            "ELECTRIC" in texto_entrada
            or "CAJA" in texto_entrada
        )
    )
    referencia_es_albanal = (
        "REGISTRO" in texto_referencia
        and "ALBANAL" in texto_referencia
    )
    entrada_es_albanal = (
        "REGISTRO" in texto_entrada
        and "ALBANAL" in texto_entrada
    )
    referencia_es_registro_electrico = (
        "REGISTRO" in texto_referencia
        and (
            "ELECTRIC" in texto_referencia
            or "CAJA" in texto_referencia
        )
    )
    if (
        entrada_es_registro_electrico
        and referencia_es_albanal
    ) or (
        entrada_es_albanal
        and referencia_es_registro_electrico
    ):
        return (
            False,
            "registro eléctrico incompatible con registro de albañal",
        )

    # La referencia debe pertenecer a la misma familia.
    if (
        familia_entrada is not None
        and familia_referencia != familia_entrada
    ):
        return (
            False,
            f"familia diferente: "
            f"{familia_entrada} vs {familia_referencia}",
        )
    # Validación específica para cables.
    if familia_entrada == "cable":
        subtipo_entrada = detectar_subtipo_cable(
            descripcion_entrada
        )
        subtipo_referencia = detectar_subtipo_cable(
            descripcion_referencia
        )
        if (
            subtipo_entrada is not None
            and subtipo_referencia != subtipo_entrada
        ):
            return (
                False,
                f"subtipo de cable diferente o no identificado: "
                f"{subtipo_entrada} vs {subtipo_referencia}",
            )
    # Validación específica para tuberías.
    if familia_entrada == "tuberia_electrica":
        subtipo_entrada = detectar_subtipo_tuberia(
            descripcion_entrada
        )
        subtipo_referencia = detectar_subtipo_tuberia(
            descripcion_referencia
        )
        if (
            subtipo_entrada is not None
            and subtipo_referencia != subtipo_entrada
        ):
            return (
                False,
                f"subtipo de tubería diferente o no identificado: "
                f"{subtipo_entrada} vs {subtipo_referencia}",
            )
    # Validación específica para concreto y acero de refuerzo: deben ser
    # del mismo elemento estructural (losa, muro, zapata, columna, etc.)
    # cuando se pueda identificar en ambos textos.
    if familia_entrada in ("concreto", "acero_refuerzo"):
        elemento_entrada = detectar_elemento_estructural(
            descripcion_entrada
        )
        elemento_referencia = detectar_elemento_estructural(
            descripcion_referencia
        )
        if (
            elemento_entrada is not None
            and elemento_referencia is not None
            and elemento_referencia != elemento_entrada
        ):
            return (
                False,
                f"elemento estructural diferente: "
                f"{elemento_entrada} vs {elemento_referencia}",
            )
    # Validación específica para cajas y registros.
    if familia_entrada == "caja_electrica":
        subtipo_entrada = detectar_subtipo_caja(
            descripcion_entrada
        )
        subtipo_referencia = detectar_subtipo_caja(
            descripcion_referencia
        )
        if (
            subtipo_entrada is not None
            and subtipo_referencia != subtipo_entrada
        ):
            return (
                False,
                f"subtipo de caja diferente o no identificado: "
                f"{subtipo_entrada} vs {subtipo_referencia}",
            )
        medida_entrada = extraer_medida_caja(
            descripcion_entrada
        )
        medida_referencia = extraer_medida_caja(
            descripcion_referencia
        )
        if (
            medida_entrada is not None
            and medida_referencia != medida_entrada
        ):
            return (
                False,
                f"medida de caja diferente o no identificada: "
                f"{medida_entrada} vs {medida_referencia}",
            )
    validaciones = [
        (
            "amperaje",
            extraer_amperaje(descripcion_entrada),
            extraer_amperaje(descripcion_referencia),
        ),
        (
            "número de conductores",
            extraer_numero_conductores(
                descripcion_entrada
            ),
            extraer_numero_conductores(
                descripcion_referencia
            ),
        ),
        (
            "polos",
            extraer_polos(descripcion_entrada),
            extraer_polos(descripcion_referencia),
        ),
        (
            "calibre AWG",
            extraer_calibre_awg(descripcion_entrada),
            extraer_calibre_awg(descripcion_referencia),
        ),
        (
            "diámetro en milímetros",
            extraer_diametro_mm(descripcion_entrada),
            extraer_diametro_mm(descripcion_referencia),
        ),
        (
            "diámetro en pulgadas",
            extraer_diametro_pulgadas(
                descripcion_entrada
            ),
            extraer_diametro_pulgadas(
                descripcion_referencia
            ),
        ),
    ]
    for nombre, entrada, referencia in validaciones:
        if (
            entrada is not None
            and referencia != entrada
        ):
            return (
                False,
                f"{nombre} diferente o no identificado: "
                f"{entrada} vs {referencia}",
            )
    return True, None


def clasificar(precio: float, low: float, high: float) -> str:
    if precio < low:
        return 'BAJO'
    if precio > high:
        return 'ALTO'
    return 'EN MERCADO'


# Margen fijo (+/-) alrededor del precio de referencia para considerarse
# "EN MERCADO" -- pedido explicito de direccion en vez de las bandas
# anteriores (percentiles p25-p75 en NL, +/-15% en CDMX, +/-10% en
# historico): un solo numero, igual en las tres fuentes, mas facil de
# explicar en una junta que "depende de que tan dispersos esten los
# precios historicos de ese concepto".
MARGEN_EN_MERCADO = 0.05


def banda_en_mercado(precio_referencia: float, margen: float = MARGEN_EN_MERCADO):
    """Regresa (banda_baja, banda_alta) = precio_referencia +/- margen."""
    return precio_referencia * (1 - margen), precio_referencia * (1 + margen)


# Si el precio cotizado esta a mas de este multiplo del precio de
# referencia (para arriba o para abajo), se manda a revision de IA aunque
# el match tenga confianza ALTA/MEDIA -- un texto parecido no garantiza
# que sea el mismo producto/tamano/calibre cuando el precio esta a este
# nivel de distancia (ej. una "TEE" generica de $115 contra una "TEE"
# industrial cotizada en $56,000: mismo nombre, escala totalmente distinta).
FACTOR_DIFERENCIA_EXTREMA = 5.0


def es_diferencia_extrema(precio_cotizado: float, precio_referencia: float,
                           factor: float = FACTOR_DIFERENCIA_EXTREMA) -> bool:
    if not precio_referencia or not precio_cotizado:
        return False
    return (
        precio_cotizado > precio_referencia * factor
        or precio_cotizado < precio_referencia / factor
    )


def score_combinado(s1, s2, *, processor=None, score_cutoff=None):
    """Scorer compuesto: en vez de depender de un solo algoritmo de
    similitud, calcula tres y se queda con el mejor. Esto ayuda cuando la
    cotizacion trae las palabras en otro orden, le faltan palabras que la
    base si tiene (o al reves), o solo coincide un fragmento largo. Un
    solo scorer (ej. token_set_ratio) falla distinto en cada uno de esos
    casos; combinarlos sube el recall sin bajar el umbral a ciegas."""
    a = processor(s1) if processor else s1
    b = processor(s2) if processor else s2
    return max(
        fuzz.token_set_ratio(a, b),
        fuzz.token_sort_ratio(a, b),
        fuzz.partial_token_sort_ratio(a, b),
    )


UMBRAL_CONFIANZA_ALTA = 90.0
UMBRAL_CONFIANZA_MEDIA = 78.0
UMBRAL_CONFIANZA_BAJA = 60.0  # piso absoluto: por debajo de esto no se ofrece nada


def nivel_confianza(score: float) -> str:
    if score >= UMBRAL_CONFIANZA_ALTA:
        return 'ALTA'
    if score >= UMBRAL_CONFIANZA_MEDIA:
        return 'MEDIA'
    return 'BAJA'


class ComparadorMultiFuente:
    def __init__(self, excel_path: str):
        self.nl = pd.read_excel(excel_path, sheet_name='Tabulador Homologado NL')
        # concepto_homologado ya viene normalizado (mayusculas, sin acentos) desde el homologador
        self.nl['concepto_norm'] = self.nl['concepto_homologado'].map(normalize_text)
        self.nl['unidad_norm'] = self.nl['unidad'].map(normalize_unit)

        self.cdmx = pd.read_excel(excel_path, sheet_name='Tabulador CDMX (gobierno)')
        self.cdmx['unidad_norm'] = self.cdmx['unidad'].map(normalize_unit)
        self.cdmx['concepto_norm'] = self.cdmx['concepto'].map(normalize_text)

        self._nl_pools = {
            u: df
            for u, df in self.nl.groupby('unidad_norm')
        }
        self._cdmx_pools = {
            u: df
            for u, df in self.cdmx.groupby('unidad_norm')
        }

        # Renglones crudos de NL (fecha, licitación, OCID) para actualizar cada
        # precio con el INPC de SU mes y dar evidencia por contrato. Se cargan
        # al primer uso.
        self._excel_path = excel_path
        self._nl_crudo = None
        self._nl_crudo_pools = {}

        # Fuentes pendientes de datos reales del usuario:
        self.ragasa = None
        self.competidores = None

    def registros_nl(self, row) -> pd.DataFrame:
        """Contratos (renglones crudos de NL) EQUIVALENTES al concepto `row`:
        misma unidad, texto casi idéntico (≥ 90), sin especificaciones en
        conflicto (p. ej. block 10x20 contra 15x20) y sin cifras atípicas
        (más de 5× o menos de 1/5 de la mediana del grupo). Con ellos se
        actualiza CADA precio con el INPC de su mes y luego se calcula la
        mediana."""
        try:
            from validacion_referencias import alcance_distinto, comparar_especificaciones
            if self._nl_crudo is None:
                crudo = pd.read_excel(self._excel_path, sheet_name='Precios Contratados (real)')
                crudo['unidad_norm'] = crudo['unidad'].map(normalize_unit)
                self._nl_crudo = crudo
            u = row['unidad_norm'] if 'unidad_norm' in row else normalize_unit(row['unidad'])
            if u not in self._nl_crudo_pools:
                pool = self._nl_crudo[self._nl_crudo['unidad_norm'] == u].copy()
                pool['concepto_norm'] = pool['concepto'].map(normalize_text)
                self._nl_crudo_pools[u] = pool
            pool = self._nl_crudo_pools[u]
            if pool.empty:
                return pool
            elegidos = process.extract(row['concepto_norm'], pool['concepto_norm'].tolist(),
                                       scorer=fuzz.token_set_ratio, score_cutoff=90, limit=None)
            idx = [i for _, _, i in elegidos]
            # Miembros seguros del grupo homologado: renglones cuyo precio ES
            # uno de sus estadísticos (mín, p25, mediana, p75, máx), dentro de
            # su periodo y con texto muy parecido (≥ 85).
            stats = [float(row[c]) for c in ('precio_min', 'precio_p25', 'precio_mediana', 'precio_p75', 'precio_max')
                     if c in row and pd.notna(row[c])]
            f = pool['fecha'].astype(str).str[:10]
            dentro = pool[(f >= str(row['fecha_min'])[:10]) & (f <= str(row['fecha_max'])[:10])]
            for i_pos, (pr, txt) in enumerate(zip(dentro['precio_unitario'], dentro['concepto_norm'])):
                if any(abs(float(pr) - st) < 0.005 for st in stats) and \
                        fuzz.token_set_ratio(row['concepto_norm'], txt) >= 85:
                    idx.append(pool.index.get_loc(dentro.index[i_pos]))
            grupo = pool.iloc[sorted(set(idx))]
            grupo = grupo[[not comparar_especificaciones(row['concepto_norm'], x)[1]
                           and not alcance_distinto(row['concepto_norm'], x) for x in grupo['concepto_norm']]]
            med = float(row['precio_mediana'] or 0)
            if med:
                grupo = grupo[(grupo['precio_unitario'] <= med * 5) & (grupo['precio_unitario'] >= med / 5)]
            return grupo
        except Exception:
            return pd.DataFrame()

    def fecha_de_precio(self, row, valor):
        """Fecha del renglón crudo de NL cuyo precio ES el valor indicado
        (p. ej. la mediana del grupo), para actualizarlo con el INPC de su
        propio mes. None si no se localiza un renglón con ese precio exacto."""
        try:
            if valor is None:
                return None
            self.registros_nl(row)   # carga los renglones crudos de esa unidad
            pool = self._nl_crudo_pools.get(row['unidad_norm'])
            if pool is None or pool.empty:
                return None
            fmin, fmax = str(row['fecha_min'])[:10], str(row['fecha_max'])[:10]
            f = pool['fecha'].astype(str).str[:10]
            cand = pool[(f >= fmin) & (f <= fmax) & ((pool['precio_unitario'] - float(valor)).abs() < 0.005)]
            if cand.empty:
                return None
            cand = cand.assign(_s=[fuzz.token_set_ratio(row['concepto_norm'], x) for x in cand['concepto_norm']])
            cand = cand[cand['_s'] >= 85].sort_values('_s', ascending=False)
            if cand.empty:
                return None
            c = cand.iloc[0]
            return {'fecha': str(c['fecha'])[:10], 'licitacion': c.get('licitacion_id'), 'ocid': c.get('ocid'),
                    'dependencia': c.get('dependencia'), 'proyecto': c.get('proyecto'), 'tipo': c.get('fuente'),
                    'precio': float(c['precio_unitario'])}
        except Exception:
            return None

    @staticmethod
    def periodo_de_precio(fecha, *identificadores):
        """Mes ('AAAA-MM') desde el que se actualiza un precio de NL.

        La columna 'fecha' de la base es la fecha en que el registro se
        PUBLICÓ en datos abiertos (hay decenas de miles de renglones con el
        mismo día), no la del concurso. El año real del precio viene en el
        número de licitación / OCID (termina en el año: ...-E126-2016).
        - Si el año de la licitación es anterior al de publicación, se usa
          julio de ese año (mitad del año; el mes no está disponible).
        - Si coincide, se usa el mes de publicación.
        Devuelve (periodo, año_de_licitación, nota)."""
        fecha = str(fecha or '')[:10]
        anio_pub = int(fecha[:4]) if fecha[:4].isdigit() else None
        anio_lic = None
        for ident in identificadores:
            m = re.findall(r'(?<!\d)(20[0-3]\d)(?!\d)', str(ident or ''))
            if m:
                anio_lic = int(m[-1])
                break
        if anio_lic and anio_pub and anio_lic < anio_pub:
            return (f'{anio_lic}-07', anio_lic,
                    f'año de la licitación ({anio_lic}) según su número; el registro se publicó en {fecha[:7]}. '
                    'Mes del concurso no disponible: se usa julio (mitad del año)')
        return fecha[:7], anio_lic or anio_pub, 'mes de publicación del registro (mismo año que la licitación)'

    def par_de_mediana(self, row):
        """Con número par de registros, la mediana es el promedio de dos
        renglones: los localiza (precio y fecha) para actualizar cada uno con
        el INPC de su mes."""
        try:
            self.registros_nl(row)
            pool = self._nl_crudo_pools.get(row['unidad_norm'])
            fmin, fmax = str(row['fecha_min'])[:10], str(row['fecha_max'])[:10]
            f = pool['fecha'].astype(str).str[:10]
            cand = pool[(f >= fmin) & (f <= fmax) & (pool['precio_unitario'] >= float(row['precio_min']) - 0.005)
                        & (pool['precio_unitario'] <= float(row['precio_max']) + 0.005)]
            cand = cand[[fuzz.token_set_ratio(row['concepto_norm'], x) >= 85 for x in cand['concepto_norm']]]
            regs = cand.to_dict('records')
            objetivo = float(row['precio_mediana'])
            for i in range(len(regs)):
                for j in range(i + 1, len(regs)):
                    if abs((regs[i]['precio_unitario'] + regs[j]['precio_unitario']) / 2 - objetivo) < 0.01:
                        return [regs[i], regs[j]]
        except Exception:
            pass
        return None

    def ajuste_nl(self, row, ajustar_inflacion=True) -> dict:
        """Mediana/p25/p75 de un concepto NL actualizados por inflación con el
        mes REAL de los contratos (por renglón, por el renglón que es la
        mediana, o por el par que la promedia); el mes intermedio solo como
        último recurso, marcado como aproximación."""
        anio_dato = str(row['fecha_max'])[:4]
        # Base de inflación: mes intermedio del periodo de los registros.
        periodo_base = _inflacion.periodo_medio(row['fecha_min'], row['fecha_max']) or anio_dato
        factor = factor_ajuste(periodo_base) if ajustar_inflacion else 1.0
        p25_uso = ajustar_precio(row['precio_p25'], periodo_base) if ajustar_inflacion else float(row['precio_p25'])
        p75_uso = ajustar_precio(row['precio_p75'], periodo_base) if ajustar_inflacion else float(row['precio_p75'])
        mediana_uso = ajustar_precio(row['precio_mediana'], periodo_base) if ajustar_inflacion else float(row['precio_mediana'])
        # Mejor: actualizar CADA renglón con el INPC de su propio mes y
        # luego sacar mediana/p25/p75 (sin aproximar con un mes intermedio).
        crudos = self.registros_nl(row)
        metodo_inflacion = 'mes intermedio del periodo (aproximación: no se localizaron los renglones)'
        detalle_registros = []
        registros_todos = []
        n_contratos = 0
        if ajustar_inflacion and not crudos.empty:
            # 1) Cada renglón con el INPC de SU mes (todo exportable).
            for idx_hoja, c in crudos.iterrows():
                periodo_c, anio_lic, nota_periodo = self.periodo_de_precio(
                    c['fecha'], c.get('licitacion_id'), c.get('ocid'))
                base, etiqueta, metodo_ind = _inflacion.indice_base(periodo_c)
                metodo_ind = f'{metodo_ind}. Periodo: {nota_periodo}'
                fac = _inflacion.NIVEL_ACTUAL / base
                registros_todos.append({
                    'fila_hoja': int(idx_hoja) + 2,   # fila en 'Precios Contratados (real)'
                    'ocid': c.get('ocid'), 'licitacion': c.get('licitacion_id'),
                    'dependencia': c.get('dependencia'), 'proyecto': c.get('proyecto'), 'tipo': c.get('fuente'),
                    'concepto': str(c.get('concepto'))[:200], 'fecha': str(c['fecha'])[:10],
                    'anio_licitacion': anio_lic, 'periodo_precio': periodo_c,
                    'precio_original': float(c['precio_unitario']),
                    'indice_mes': etiqueta, 'indice_base': base, 'metodo_indice': metodo_ind,
                    'indice_final_mes': _inflacion.ETIQUETA_ACTUAL, 'indice_final': _inflacion.NIVEL_ACTUAL,
                    'factor': round(fac, 6), 'precio_actualizado': round(float(c['precio_unitario']) * fac, 2),
                    # Inflación ACUMULADA (compuesta) del mes del renglón a hoy.
                    'inflacion_acumulada_pct': round((fac - 1) * 100, 2),
                })
            # 2) Un contrato (OCID) = una observación: la base no trae el
            #    identificador de partida, así que varios renglones del mismo
            #    contrato no se cuentan como contratos independientes.
            df_r = pd.DataFrame(registros_todos)
            df_r['contrato'] = df_r['ocid'].fillna(df_r['licitacion']).astype(str)
            # Con todos los decimales (igual que las fórmulas del Excel); el
            # redondeo a centavos se hace al final.
            df_r['_exacto'] = df_r['precio_original'] * _inflacion.NIVEL_ACTUAL / df_r['indice_base']
            por_contrato = df_r.groupby('contrato').agg(actualizado=('_exacto', 'median'),
                                                        original=('precio_original', 'median'))
            n_contratos = int(len(por_contrato))
            serie = por_contrato['actualizado']
            mediana_uso = round(float(serie.median()), 2)
            p25_uso = round(float(serie.quantile(0.25)), 2)
            p75_uso = round(float(serie.quantile(0.75)), 2)
            _med_orig = float(por_contrato['original'].median())
            factor = mediana_uso / _med_orig if _med_orig else factor
            metodo_inflacion = (f'{len(crudos)} renglones en {n_contratos} contrato(s): cada renglón con el INPC '
                                'de su mes; mediana por contrato y después mediana entre contratos')
        if ajustar_inflacion and crudos.empty:
            # Sin los renglones del grupo: se localiza el renglón cuyo precio
            # ES cada estadístico (mediana, p25, p75) y se usa SU mes.
            f_med = self.fecha_de_precio(row, row['precio_mediana'])
            if f_med:
                base_med = self.periodo_de_precio(f_med['fecha'], f_med.get('licitacion'), f_med.get('ocid'))[0]
                mediana_uso = ajustar_precio(row['precio_mediana'], base_med)
                f25 = self.fecha_de_precio(row, row['precio_p25'])
                f75 = self.fecha_de_precio(row, row['precio_p75'])
                _per = lambda x: self.periodo_de_precio(x['fecha'], x.get('licitacion'), x.get('ocid'))[0]
                p25_uso = ajustar_precio(row['precio_p25'], _per(f25) if f25 else base_med)
                p75_uso = ajustar_precio(row['precio_p75'], _per(f75) if f75 else base_med)
                factor = factor_ajuste(base_med)
                periodo_base = base_med
                metodo_inflacion = f'mes del renglón que es la mediana ({base_med})'
                detalle_registros.append(f_med)
            else:
                par = self.par_de_mediana(row)
                if par:
                    _pers = [self.periodo_de_precio(r_['fecha'], r_.get('licitacion_id'), r_.get('ocid'))[0]
                             for r_ in par]
                    ajust = [float(r_['precio_unitario']) * _inflacion.NIVEL_ACTUAL
                             / _inflacion.indice_base(pp)[0] for r_, pp in zip(par, _pers)]
                    mediana_uso = round(sum(ajust) / 2, 2)
                    factor = mediana_uso / float(row['precio_mediana'])
                    meses = sorted(set(_pers))
                    periodo_base = meses[0]
                    p25_uso = round(float(row['precio_p25']) * factor, 2)
                    p75_uso = round(float(row['precio_p75']) * factor, 2)
                    metodo_inflacion = (f"mediana = promedio de 2 renglones; cada uno con el INPC de su mes "
                                        f"({' y '.join(meses)})")
                    for r_ in par:
                        detalle_registros.append({
                            'fecha': str(r_['fecha'])[:10], 'precio': float(r_['precio_unitario']),
                            'licitacion': r_.get('licitacion_id'), 'ocid': r_.get('ocid'),
                            'dependencia': r_.get('dependencia'), 'proyecto': r_.get('proyecto'),
                            'tipo': r_.get('fuente')})
        for _, c in (crudos.sort_values('fecha', ascending=False).head(5) if not crudos.empty else crudos).iterrows():
            detalle_registros.append({
                'fecha': str(c['fecha'])[:10], 'precio': float(c['precio_unitario']),
                'licitacion': c.get('licitacion_id'), 'ocid': c.get('ocid'),
                'dependencia': c.get('dependencia'), 'proyecto': c.get('proyecto'), 'tipo': c.get('fuente'),
            })
        tipos = set(crudos['fuente'].astype(str)) if not crudos.empty and 'fuente' in crudos else set()
        if crudos.empty and detalle_registros:
            tipos = {str(detalle_registros[0].get('tipo'))}
        if not ajustar_inflacion:
            factor = 1.0
            p25_uso, p75_uso, mediana_uso = float(row['precio_p25']), float(row['precio_p75']), float(row['precio_mediana'])
            metodo_inflacion = 'sin ajuste por inflación'
        grupo_ok = (not crudos.empty) and ajustar_inflacion
        return {
            'mediana_original': (round(float(pd.DataFrame(registros_todos).assign(
                                    k=lambda d: d['ocid'].fillna(d['licitacion']).astype(str))
                                    .groupby('k')['precio_original'].median().median()), 2) if grupo_ok
                                 else float(row['precio_mediana'])),
            'n': n_contratos if grupo_ok else int(row['n_registros']),
            'n_renglones': int(len(crudos)) if grupo_ok else int(row['n_registros']),
            'registros_todos': registros_todos,
            'fecha_min': (str(crudos['fecha'].astype(str).min())[:10] if grupo_ok else str(row['fecha_min'])[:10]),
            'fecha_max': (str(crudos['fecha'].astype(str).max())[:10] if grupo_ok else str(row['fecha_max'])[:10]),
            'mediana': round(float(mediana_uso), 2), 'p25': round(float(p25_uso), 2), 'p75': round(float(p75_uso), 2),
            'factor': factor, 'periodo_base': periodo_base, 'metodo': metodo_inflacion,
            'registros_usados': int(len(crudos)), 'registros_detalle': detalle_registros,
            'tipos': ', '.join(sorted(tipos)),
            # Año del precio = año de la licitación más reciente del grupo
            # (no el de publicación del registro).
            'anio_dato': (str(max(r['periodo_precio'] for r in registros_todos))[:4] if grupo_ok else anio_dato),
            'periodo_precio_min': (min(r['periodo_precio'] for r in registros_todos) if grupo_ok else None),
            'periodo_precio_max': (max(r['periodo_precio'] for r in registros_todos) if grupo_ok else None),
        }

    def cargar_ragasa(self, df_o_ruta):
        """Conecta el historico real de compras de Ragasa cuando este disponible.
        Se espera columnas: concepto, unidad, precio_unitario, fecha (opcional)."""
        self.ragasa = df_o_ruta if isinstance(df_o_ruta, pd.DataFrame) else pd.read_excel(df_o_ruta)
        self.ragasa['concepto_norm'] = self.ragasa['concepto'].map(normalize_text)
        self.ragasa['unidad_norm'] = self.ragasa['unidad'].map(normalize_unit)

    def cargar_competidores(self, df_o_ruta):
        """Conecta las propuestas economicas reales de otros proveedores de la
        MISMA licitacion que se esta revisando (para comparar entre pares).
        Se espera columnas: proveedor, concepto, unidad, precio_unitario."""
        self.competidores = df_o_ruta if isinstance(df_o_ruta, pd.DataFrame) else pd.read_excel(df_o_ruta)
        self.competidores['concepto_norm'] = self.competidores['concepto'].map(normalize_text)
        self.competidores['unidad_norm'] = self.competidores[
            'unidad'
        ].map(normalize_unit)

    def _match_pool(
        self,
        pools,
        t,
        u,
        min_score,
        scorer,
        text_col='concepto_norm',
    ):
        """Busca con el texto original y con su version de catalogo
        (consulta_catalogo) y se queda con la mejor coincidencia."""
        resultado = None
        for consulta in consultas_catalogo(t):
            otro = self._match_pool_texto(pools, consulta, u, min_score, scorer, text_col, original=t)
            if otro and (resultado is None or otro[0] > resultado[0]):
                resultado = otro
        return resultado

    def _match_pool_texto(
        self,
        pools,
        t,
        u,
        min_score,
        scorer,
        text_col='concepto_norm',
        original=None,
    ):
        """Busca la mejor coincidencia dentro del pool de la unidad `u`.

        Hace dos pasadas: primero con `min_score` (coincidencia
        confiable), y si no encuentra nada, reintenta con un umbral mas
        bajo (UMBRAL_CONFIANZA_BAJA) para no dejar la partida totalmente
        sin dato solo porque la redaccion es muy distinta. La segunda
        pasada siempre se marca con nivel de confianza 'BAJA' para que
        quien revise sepa que debe confirmarla a mano; nunca se mezcla en
        silencio con las coincidencias fuertes.

        Regresa (score, row, nivel_confianza) o None si ni siquiera al
        umbral minimo hay algo tecnicamente compatible.
        """
        pool = pools.get(u)
        if pool is None or pool.empty:
            return None
        choices = pool[
            text_col
        ].fillna("").tolist()

        def _buscar(umbral):
            # Paso 1: preseleccion rapida con el scorer nativo de rapidfuzz
            # (implementado en C, corre sobre miles de filas en
            # milisegundos). Paso 2: el scorer combinado -mas lento por
            # ser Python puro, ya que corre 3 algoritmos- solo se aplica
            # sobre ese shortlist corto. Esto evita que comparar una
            # cotizacion con muchos renglones se sienta lento, sin perder
            # el beneficio del scorer combinado: si el texto real es
            # parecido, ya va a aparecer en la preseleccion (el propio
            # token_set_ratio es uno de los tres que se combinan despues).
            preseleccion = process.extract(
                t,
                choices,
                scorer=fuzz.token_set_ratio,
                score_cutoff=max(umbral - 25, 30),
                limit=40,
            )
            if not preseleccion:
                return None
            recalificados = sorted(
                (
                    (scorer(t, texto_candidato), indice)
                    for texto_candidato, _, indice in preseleccion
                ),
                key=lambda par: -par[0],
            )
            for score, indice in recalificados:
                if score < umbral:
                    continue
                row = pool.iloc[indice]
                descripcion_referencia = row.get(text_col, "")
                # Filtro de fragmentos demasiado cortos: cadenas de
                # referencia como "3/8", "1/2" o "1" (medidas/fracciones
                # sueltas que a veces quedan solas en la base de precios)
                # pueden marcar 100% con token_set_ratio contra CUALQUIER
                # cotizacion que mencione esa medida en cualquier parte de
                # su descripcion larga, sin que el material sea remotamente
                # el mismo (ej. "3/8" de una base de acero de refuerzo
                # marcando "match perfecto" contra tuberia de cobre para
                # refrigeracion que tambien mide 3/8"). Una referencia con
                # menos de 8 caracteres normalizados no trae suficiente
                # informacion como para confirmar que es el mismo concepto,
                # sin importar que tan alto haya salido el score.
                if len(str(descripcion_referencia).strip()) < 8:
                    continue
                compatible, _ = validar_compatibilidad_tecnica(
                    t,
                    descripcion_referencia,
                )
                if not compatible:
                    continue
                if not _referencia_compatible(original or t, descripcion_referencia):
                    continue
                return float(score), row
            return None

        resultado = _buscar(min_score)
        if resultado:
            score, row = resultado
            return score, row, nivel_confianza(score)

        if min_score > UMBRAL_CONFIANZA_BAJA:
            resultado = _buscar(UMBRAL_CONFIANZA_BAJA)
            if resultado:
                score, row = resultado
                return score, row, 'BAJA'

        return None

    # ------------------------------------------------------------------
    # Datos de mercado para comparar: varias referencias por fuente
    # ------------------------------------------------------------------
    def _top_pool(self, pools, t, u, k, text_col='concepto_norm', umbral=UMBRAL_CONFIANZA_BAJA):
        pool = pools.get(u)
        if pool is None or pool.empty:
            return []
        choices = pool[text_col].fillna("").tolist()
        resultados = {}
        for consulta in consultas_catalogo(t):
            pre = process.extract(consulta, choices, scorer=fuzz.token_set_ratio,
                                  score_cutoff=max(umbral - 25, 30), limit=60)
            for texto, _, indice in pre:
                score = score_combinado(consulta, texto)
                if score < umbral or len(str(texto).strip()) < 8:
                    continue
                if not validar_compatibilidad_tecnica(consulta, texto)[0]:
                    continue
                if not _referencia_compatible(t, texto):
                    continue
                if indice not in resultados or score > resultados[indice]:
                    resultados[indice] = score
        mejores = sorted(resultados.items(), key=lambda par: -par[1])[:k]
        return [(score, pool.iloc[indice]) for indice, score in mejores]

    def candidatos(self, descripcion: str, unidad: str, k: int = 3,
                   ajustar_inflacion: bool = True) -> list[dict]:
        """Hasta k referencias reales por fuente (NL y CDMX) para comparar.

        NL entrega mediana y rango p25-p75 de precios contratados, con número
        de registros y año; CDMX el precio del tabulador oficial 2026. Son
        datos reales publicados, pero NO validados como equivalentes: sirven
        para ubicar el precio cotizado frente al mercado.
        """
        t = normalize_text(descripcion)
        u = normalize_unit(unidad)
        salida = []
        for score, row in self._top_pool(self._nl_pools, t, u, k):
            aj_nl = self.ajuste_nl(row, ajustar_inflacion)
            med, p25, p75 = aj_nl['mediana'], aj_nl['p25'], aj_nl['p75']
            salida.append({
                'fuente': 'Nuevo León',
                'concepto': row['concepto_homologado'],
                'unidad': row['unidad'],
                'precio': round(float(med), 2),
                'rango_bajo': round(float(p25), 2),
                'rango_alto': round(float(p75), 2),
                'n_registros': aj_nl['n'],
                'fecha': f"{aj_nl['fecha_min'][:7]} a {aj_nl['fecha_max'][:7]}",
                'nota': 'mediana y rango p25-p75 de licitaciones SIASI (OCDS), ajustados por INPC'
                        if ajustar_inflacion else 'mediana y rango p25-p75 de licitaciones SIASI (OCDS)',
                'score': round(score, 1),
                'confianza': nivel_confianza(score),
            })
        for score, row in self._top_pool(self._cdmx_pools, t, u, k):
            salida.append({
                'fuente': 'CDMX',
                'concepto': row['concepto'],
                'unidad': row['unidad'],
                'precio': round(float(row['precio_unitario']), 2),
                'rango_bajo': None,
                'rango_alto': None,
                'n_registros': 1,
                'fecha': '2026-05 (tabulador oficial)',
                'nota': f"clave {row['clave']}",
                'score': round(score, 1),
                'confianza': nivel_confianza(score),
            })
        return salida

    def evaluar(self, descripcion: str, unidad: str, precio_cotizado: float,
                min_score: float = UMBRAL_CONFIANZA_MEDIA, scorer=score_combinado,
                ajustar_inflacion: bool = True, usar_ia: bool = False) -> dict:
        t = normalize_text(descripcion)
        u = normalize_unit(unidad)
        resultado = {
            'entrada': {'descripcion': descripcion, 'unidad': u, 'precio_cotizado': precio_cotizado},
            'fuentes': {},
        }

        if t.strip() in CONCEPTOS_GENERICOS_SIN_MATERIAL:
            motivo = (
                'esta partida no describe un material especifico (es un '
                'renglon generico de mano de obra/materiales varios), asi '
                'que no se compara contra la base de precios -- cualquier '
                'coincidencia seria casualidad de palabras genericas, no '
                'el mismo concepto.'
            )
            for fuente in ('nl_historico', 'cdmx_gobierno', 'ragasa_historico', 'comparacion_proveedores'):
                resultado['fuentes'][fuente] = {'match': None, 'motivo': motivo}
            resultado['veredicto_combinado'] = 'SIN DATOS SUFICIENTES'
            resultado['fuentes_consultadas'] = 0
            return resultado

        # --- Fuente 1: NL historico ---
        m = self._match_pool(self._nl_pools, t, u, min_score, scorer)
        if m:
            score, row, confianza = m
            aj_nl = self.ajuste_nl(row, ajustar_inflacion)
            anio_dato = aj_nl['anio_dato']
            periodo_base, factor = aj_nl['periodo_base'], aj_nl['factor']
            p25_uso, p75_uso, mediana_uso = aj_nl['p25'], aj_nl['p75'], aj_nl['mediana']
            metodo_inflacion, detalle_registros = aj_nl['metodo'], aj_nl['registros_detalle']
            crudos_n = aj_nl['registros_usados']
            # "EN MERCADO" = dentro de +/-5% del precio mediano de
            # referencia (ver MARGEN_EN_MERCADO), no de la banda p25-p75.
            banda_baja, banda_alta = banda_en_mercado(mediana_uso)
            veredicto = clasificar(precio_cotizado, banda_baja, banda_alta)
            resultado['fuentes']['nl_historico'] = {
                'match': row['concepto_homologado'], 'score': round(score, 1), 'unidad': row['unidad'],
                'fecha_min': aj_nl['fecha_min'], 'fecha_max': aj_nl['fecha_max'],
                'periodo_precio_min': aj_nl.get('periodo_precio_min'),
                'periodo_precio_max': aj_nl.get('periodo_precio_max'),
                'confianza': confianza,
                'precio_min': float(row['precio_min']), 'precio_p25': float(row['precio_p25']),
                'precio_mediana': aj_nl['mediana_original'], 'precio_p75': float(row['precio_p75']),
                'precio_max': float(row['precio_max']), 'n_registros': aj_nl['n'],
                'variabilidad': row['variabilidad'],
                'anio_dato_mas_reciente': anio_dato,
                'periodo_base_inflacion': periodo_base,
                'metodo_inflacion': metodo_inflacion,
                'registros_usados': crudos_n,
                'registros_todos': aj_nl.get('registros_todos', []),
                'n_renglones': aj_nl.get('n_renglones'),
                'registros_detalle': detalle_registros,
                'tipo_registros': aj_nl['tipos'],
                'ajuste_inflacion_aplicado': ajustar_inflacion,
                'factor_ajuste_inpc': round(factor, 4),
                'precio_p25_ajustado': p25_uso, 'precio_p75_ajustado': p75_uso,
                'precio_mediana_ajustada': mediana_uso,
                'banda_baja': round(banda_baja, 2), 'banda_alta': round(banda_alta, 2),
                'referencia_ajuste': f'INPC INEGI, {anio_dato} -> {_inflacion.ETIQUETA_ACTUAL} ({FUENTE_INPC})' if ajustar_inflacion else None,
                'clasificacion': veredicto,
            }
            precio_ref_nl = mediana_uso
            diferencia_extrema_nl = es_diferencia_extrema(precio_cotizado, precio_ref_nl)

            if confianza == 'BAJA':
                resultado['fuentes']['nl_historico']['motivo'] = (
                    'coincidencia debil (revisar a mano): el texto no es muy '
                    'parecido, confirma que sea el mismo concepto antes de '
                    'confiar en este precio de referencia.'
                )
            elif diferencia_extrema_nl:
                resultado['fuentes']['nl_historico']['motivo'] = (
                    'precio muy alejado de la referencia (revisar a mano): '
                    'el texto coincide bien, pero el precio esta a una '
                    'distancia inusual -- confirma que sea el mismo '
                    'producto, tamano y calibre antes de confiar en esta '
                    'clasificacion.'
                )

            # OJO: aqui YA NO se llama a la IA una partida a la vez -- eso
            # se hace en LOTE desde app_v2.py (ver revision_ia.
            # revisar_coincidencias_debiles_lote) despues de evaluar TODAS
            # las partidas de la cotizacion, para no tardar minutos ni
            # agotar el limite de solicitudes por minuto. Aqui solo se deja
            # marcado 'motivo' (arriba) para que quien orquesta sepa que
            # este match necesita revision.
        else:
            resultado['fuentes']['nl_historico'] = {'match': None, 'motivo': f'sin coincidencia ni relajada en unidad {u}'}

        # --- Fuente 2: CDMX gobierno ---
        m = self._match_pool(self._cdmx_pools, t, u, min_score, scorer)
        if m:
            score, row, confianza = m
            precio_ref = float(row['precio_unitario'])
            low, high = banda_en_mercado(precio_ref)
            resultado['fuentes']['cdmx_gobierno'] = {
                'match': row['concepto'], 'score': round(score, 1), 'clave': row['clave'],
                'pagina': (None if pd.isna(row.get('pagina')) else row.get('pagina')), 'unidad': row['unidad'],
                'confianza': confianza,
                'precio_referencia': precio_ref, 'banda_baja': round(low, 2), 'banda_alta': round(high, 2),
                'clasificacion': clasificar(precio_cotizado, low, high),
            }
            diferencia_extrema_cdmx = es_diferencia_extrema(precio_cotizado, precio_ref)

            if confianza == 'BAJA':
                resultado['fuentes']['cdmx_gobierno']['motivo'] = (
                    'coincidencia debil (revisar a mano): el texto no es muy '
                    'parecido, confirma que sea el mismo concepto antes de '
                    'confiar en este precio de referencia.'
                )
            elif diferencia_extrema_cdmx:
                resultado['fuentes']['cdmx_gobierno']['motivo'] = (
                    'precio muy alejado de la referencia (revisar a mano): '
                    'el texto coincide bien, pero el precio esta a una '
                    'distancia inusual -- confirma que sea el mismo '
                    'producto, tamano y calibre antes de confiar en esta '
                    'clasificacion.'
                )

            # Igual que en NL: la revision con IA de este match se hace en
            # LOTE desde app_v2.py, no aqui.
        else:
            resultado['fuentes']['cdmx_gobierno'] = {'match': None, 'motivo': f'sin coincidencia ni relajada en unidad {u}'}

        # --- Fuente 3: Ragasa (requiere datos reales del usuario) ---
        if self.ragasa is not None:
            pools = {uu: dfx for uu, dfx in self.ragasa.groupby('unidad_norm')}
            m = self._match_pool(pools, t, u, min_score, scorer)
            if m:
                score, row, confianza = m
                precio_hist_ragasa = float(row['precio_unitario'])
                resultado['fuentes']['ragasa_historico'] = {
                    'match': row['concepto'], 'score': round(score, 1),
                    'confianza': confianza,
                    'precio_historico': precio_hist_ragasa,
                    'clasificacion': clasificar(precio_cotizado, *banda_en_mercado(precio_hist_ragasa)),
                }
                diferencia_extrema_ragasa = es_diferencia_extrema(precio_cotizado, precio_hist_ragasa)
                if confianza == 'BAJA' or diferencia_extrema_ragasa:
                    resultado['fuentes']['ragasa_historico']['motivo'] = (
                        'coincidencia riesgosa (revisar a mano): confianza '
                        'baja o precio muy alejado de la referencia.'
                    )
                # La revision con IA de este match (si aplica) se hace en
                # LOTE desde app_v2.py, no aqui.
            else:
                resultado['fuentes']['ragasa_historico'] = {'match': None, 'motivo': 'sin coincidencia en historico Ragasa'}
        else:
            resultado['fuentes']['ragasa_historico'] = {
                'match': None,
                'motivo': 'PENDIENTE: no se ha cargado el historico de compras de Ragasa. '
                          'Usa comparador.cargar_ragasa(ruta_o_dataframe) con datos reales para activar esta fuente.'
            }

        # --- Fuente 4: comparacion entre proveedores de la misma licitacion ---
        if self.competidores is not None:
            comp = self.competidores[self.competidores['unidad_norm'] == u].copy()
            if not comp.empty:
                comp['_score'] = comp['concepto_norm'].map(lambda x: scorer(t, x))
                comp = comp[comp['_score'] >= min_score]
            if not comp.empty:
                resultado['fuentes']['comparacion_proveedores'] = {
                    'n_proveedores_comparables': int(comp['proveedor'].nunique()),
                    'precio_min': float(comp['precio_unitario'].min()),
                    'precio_mediana': float(comp['precio_unitario'].median()),
                    'precio_max': float(comp['precio_unitario'].max()),
                    'clasificacion': clasificar(precio_cotizado, comp['precio_unitario'].quantile(0.25), comp['precio_unitario'].quantile(0.75)),
                    'detalle': comp[['proveedor', 'precio_unitario']].to_dict('records'),
                }
            else:
                resultado['fuentes']['comparacion_proveedores'] = {'match': None, 'motivo': 'sin otras propuestas comparables para este concepto'}
        else:
            resultado['fuentes']['comparacion_proveedores'] = {
                'match': None,
                'motivo': 'PENDIENTE: no se han cargado las propuestas de otros proveedores de esta licitacion. '
                          'Usa comparador.cargar_competidores(ruta_o_dataframe) con las cotizaciones reales para activar esta fuente.'
            }

        # Si un match quedo marcado como riesgoso (confianza BAJA o
        # precio con diferencia extrema) y la IA lo rechazo, no estuvo
        # segura, o la revision nunca se completo (ej. limite de la
        # cuenta gratuita), no debe contar para el veredicto combinado --
        # ver _revision_ia.debe_descartarse().
        clasificaciones = [
            f['clasificacion'] for f in resultado['fuentes'].values()
            if isinstance(f, dict) and f.get('clasificacion')
            and not _revision_ia.debe_descartarse(f, usar_ia)[0]
        ]
        if clasificaciones:
            conteo = {c: clasificaciones.count(c) for c in set(clasificaciones)}
            resultado['veredicto_combinado'] = max(conteo, key=conteo.get)
            resultado['fuentes_consultadas'] = len(clasificaciones)
            resultado['coincidencia_entre_fuentes'] = len(set(clasificaciones)) == 1
        else:
            resultado['veredicto_combinado'] = 'SIN DATOS SUFICIENTES'
            resultado['fuentes_consultadas'] = 0
            # La opinion de IA para partidas sin ningun dato (si aplica) se
            # pide en LOTE desde app_v2.py, no aqui -- ver
            # revision_ia.opinar_sin_datos_lote().

        return resultado


if __name__ == '__main__':
    import sys
    import json

    excel = sys.argv[1] if len(sys.argv) > 1 else 'Base_Precios_Unitarios_NL_CDMX.xlsx'
    c = ComparadorMultiFuente(excel)
    ejemplos = [
        ("Suministro y colocacion de acero de refuerzo en losas, varilla corrugada", "KG", 30),
        ("Limpieza final de obra durante todo el periodo de ejecucion", "M2", 9),
        ("Anteproyecto de muro de contencion, primeros 100 m2", "m2", 90),
    ]
    for desc, unidad, precio in ejemplos:
        print('=' * 100)
        print('ENTRADA:', desc, '|', unidad, '| cotizado:', precio)
        print(json.dumps(c.evaluar(desc, unidad, precio), indent=2, ensure_ascii=False, default=str))
        print()
