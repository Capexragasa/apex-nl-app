"""
Validación de referencias de precio por fuente
==============================================

Este módulo es el que CONECTA la revisión con IA con los cálculos finales.
Antes, la IA podía decir RECHAZA y el precio rechazado seguía entrando al
promedio y al semáforo. Aquí cada fuente (Histórico Ragasa, Nuevo León,
CDMX, IA internet) recibe un ESTADO, y solo las referencias VALIDADAS se
usan para decidir el semáforo final y el precio de negociación.

Reglas (pedidas explícitamente por Compras CAPEX):

1. RECHAZA (IA)  -> la referencia se excluye del precio sugerido, de las
                    diferencias y del semáforo. Queda solo como evidencia.
2. NO_SEGURO (IA) -> se conserva como orientativa: estado NO CONCLUYENTE,
                    se muestra su precio y su % pero no decide nada.
3. CONFIRMA (IA) -> no basta con que el concepto sea el mismo: además se
                    verifica unidad, alcance, fecha del dato y escala del
                    precio antes de marcarla VALIDADA.
4. (En busqueda_mercado_ia.py) el precio web debe venir del mismo
   fragmento que menciona el concepto y la unidad.
5. Cada fuente se compara por separado. No se promedian fuentes: el precio
   de negociación sale de UNA referencia validada, con prioridad
   Histórico Ragasa > Nuevo León > CDMX > IA internet, y se dice cuál.

Estados posibles de una referencia:
    VALIDADA              concepto confirmado por IA + especificaciones de la
                          referencia declaradas por el proveedor y coincidentes
                          + precio vigente (≤ 12 meses). Decide semáforo y ahorro.
    EQUIVALENCIA PARCIAL  concepto confirmado, pero falta demostrar
                          especificación o vigencia: orientativa; su diferencia
                          se reporta como "pendiente de validar", nunca ahorro.
    NO CONCLUYENTE        la IA no pudo confirmarla, precio fuera de escala o
                          fragmento web sin verificar
    POR CONFIRMAR         la IA no ha confirmado el concepto
    RECHAZADA             la IA la rechazó, o material/especificación/alcance
                          distinto (no requiere IA)
    SIN DATO              la fuente no encontró nada comparable
"""

from __future__ import annotations

import datetime as _dt
import re
import unicodedata

VALIDADA = "VALIDADA"
EQUIVALENCIA_PARCIAL = "EQUIVALENCIA PARCIAL"
NO_CONCLUYENTE = "NO CONCLUYENTE"
POR_CONFIRMAR = "POR CONFIRMAR"
RECHAZADA = "RECHAZADA"
SIN_DATO = "SIN DATO"
# Concepto parecido, pero el precio no se puede comparar: otra escala, o
# precio por pieza con tamaño distinto o no declarado. No vota.
NO_COMPARABLE = "NO COMPARABLE"

ORIENTATIVAS = (EQUIVALENCIA_PARCIAL, NO_CONCLUYENTE, POR_CONFIRMAR)

# Semáforo final
ALTO = "ALTO"
EN_MERCADO = "EN MERCADO"
BAJO = "BAJO"
MIXTO = "MIXTO"
SIN_VALIDADA = "SIN DATOS SUFICIENTES"

MARGEN_EN_MERCADO = 0.05          # ±5 % alrededor de la referencia
FACTOR_ESCALA = 5.0               # filtro de cifras extremas (NO demuestra equivalencia)
VIGENCIA_MESES = 12               # un precio más viejo no demuestra precio vigente

PRIORIDAD_NEGOCIACION = ("historico", "nl", "cdmx", "ia")
NOMBRE_FUENTE = {
    "historico": "Histórico Ragasa",
    "nl": "Nuevo León",
    "cdmx": "CDMX",
    "ia": "IA internet",
}


def _plano(texto) -> str:
    t = unicodedata.normalize("NFKD", str(texto or "").lower())
    return "".join(c for c in t if not unicodedata.combining(c))


def clasificar(precio: float, referencia: float, margen: float = MARGEN_EN_MERCADO) -> str:
    if precio < referencia * (1 - margen):
        return BAJO
    if precio > referencia * (1 + margen):
        return ALTO
    return EN_MERCADO


def diferencia_pct(precio: float, referencia: float):
    if not referencia:
        return None
    return round((precio - referencia) / referencia * 100, 1)


_METAL = re.compile(r"\b(metalic[oa]s?|acero estructural|placa de acero|perfil(es)? de acero|ptr|viga ipr|soldad[oa])\b")
_MAMPOSTERIA = re.compile(r"\b(barda|block|bloque|tabique|concreto|castillos?|armex|mamposteria|dala|cerramiento)\b")


# Elementos estructurales: un castillo no es un cerramiento aunque ambos
# lleven Armex y concreto. Se compara el elemento PRINCIPAL (el primero que
# se menciona) de la partida contra el de la referencia.
_ELEMENTOS = (
    ("cerramiento", r"\b(cerramientos?|dalas?|cadenas?)\b"),
    ("castillo", r"\b(castillos?|columnas? (para amarrar|de amarre|de barda|de confinamiento))\b"),
    ("columna", r"\bcolumnas?\b"),
    ("trabe", r"\btrabes?\b"),
    ("zapata", r"\bzapatas?\b"),
    ("losa", r"\blosas?\b"),
    ("firme", r"\bfirmes?\b"),
    ("muro", r"\b(muros?|bardas?|pretil)\b"),
)


def elemento_principal(texto: str):
    t = _plano(texto)
    encontrados = []
    for nombre, patron in _ELEMENTOS:
        m = re.search(patron, t)
        if m:
            encontrados.append((m.start(), nombre))
    if not encontrados:
        return None
    nombre = min(encontrados)[1]
    return "castillo" if nombre == "columna" and re.search(_ELEMENTOS[1][1], t) else nombre


# Trabajo principal de acabados: un aplanado/estuco no se compara contra una
# pintura o un impermeabilizante aunque el texto mencione "aplanado".
_TRABAJOS = (
    ("aplanado", r"\b(estuco|aplanados?|zarpeo|repellados?|enjarres?|afine)\b"),
    ("pintura", r"\b(pintura|esmalte|vinilica|sellador)\b"),
    ("impermeabilizacion", r"\bimpermeabiliza\w*"),
    ("demolicion", r"\b(demolicion|desmantelamiento|retiro)\b"),
    ("cimbra", r"\b(cimbra|descimbra)\b"),
)


def trabajo_principal(texto: str):
    t = _plano(texto)
    encontrados = [(m.start(), nombre) for nombre, patron in _TRABAJOS
                   for m in [re.search(patron, t)] if m]
    return min(encontrados)[1] if encontrados else None


_VERBOS_OBRA = (r"suministro|colocacion|instalacion|aplicacion|fabricacion|construccion|elaboracion|"
                r"habilitado|armado|servicio|trabajos?|ejecucion|realizacion|renta|mano de obra|"
                r"suministrar|colocar|instalar|aplicar|fabricar|construir|elaborar")
_RELLENO = {"de", "del", "la", "el", "los", "las", "y", "e", "en", "para", "con", "por", "a", "al", "un", "una"}
# Formas de nombrar el mismo objeto (raíz -> sinónimos aceptados en la referencia).
_GRUPOS_OBJETO = {
    "columna": {"castillo", "columna"},
    "castillo": {"castillo", "columna"},
    "cerramiento": {"cerramiento", "dala", "cadena"},
    "dala": {"cerramiento", "dala", "cadena"},
    "cadena": {"cerramiento", "dala", "cadena"},
    "muro": {"muro", "barda", "pretil"},
    "barda": {"muro", "barda", "pretil"},
    "estuco": {"estuco", "aplanado"},
    "aplanado": {"aplanado", "repellado", "enjarre", "estuco", "zarpeo"},
    "grua": {"grua"},
}


def _cabeza(texto: str, n: int) -> str:
    """Primeras n palabras del objeto, sin códigos ni verbos de obra."""
    t = _plano(texto)
    # Quita códigos, claves entre paréntesis y unidades sueltas al inicio:
    # "3182000026 (M50300129) SUMINISTRO..." o "M3 13.47 SUMINISTRO...".
    for _ in range(4):
        t = re.sub(r"^\s*(\(?[a-z]{0,4}\d[\w.\-]*\)?|\b(m2|m3|ml|m|pza|kg|lote|jgo|total)\b)[\s,.:;-]*", "", t)
    t = re.sub(r"^[\W\d.\-]*", "", t)
    t = re.sub(rf"^((({_VERBOS_OBRA})\b[\s,]*)|(\b(y|e|de|del|o)\b\s*))+", "", t)
    palabras = [p for p in re.findall(r"[a-z0-9']+", t)]
    return " ".join(palabras[:n])


def clave_objeto(texto: str):
    """Sustantivo principal de lo que se cotiza (raíz en singular)."""
    for p in _cabeza(texto, 6).split():
        if p in _RELLENO or len(p) < 4 or p.isdigit():
            continue
        raiz = re.sub(r"(es|s)$", "", p) if p not in ("trabes",) else "trabe"
        for k in _GRUPOS_OBJETO:
            if raiz == k or p == k:
                return k
        return raiz
    return None


# Función del elemento: mismo objeto con distinta función NO es comparable
# ("muro de block" divisorio/barda vs "muro de contención"; "cerramiento" vs
# "dala de desplante"). Una etiqueta exclusiva presente en un lado y ausente
# en el otro descarta la referencia.
_FUNCION = (
    ("contención", r"contencion|muro de carga de tierra|talud"),
    ("desplante / cimentación", r"desplante|cimentacion|zapata|contratrabe|cimiento"),
    ("provisional", r"provisional|temporal|tapial|obra falsa"),
    ("demolición / retiro", r"demolicion|desmantelamiento|\bretiro de\b|desmontaje"),
    ("reparación", r"reparacion|resane|rehabilitacion de"),
    ("fachada / divisorio ligero", r"tablaroca|tablacemento|durock|panel de yeso|drywall|muro divisorio de panel"),
)
# Material principal: si ambos lo declaran, debe coincidir.
_MATERIAL = (
    ("mampostería", r"\b(block|bloque|tabique|ladrillo|tabicon|mamposteria)\b"),
    ("concreto armado", r"\b(muro|losa|castillo|dala|cadena|cerramiento|trabe|columna)s? de concreto\b|concreto armado|armex|varilla"),
    ("acero / metal", r"\b(metalic[oa]s?|acero estructural|ptr|perfil|lamina|ipr)\b"),
    ("madera", r"\bmadera\b|triplay"),
    ("tablaroca / panel", r"tablaroca|tablacemento|durock|panel"),
)


def funcion_material(texto: str) -> dict:
    # Solo la descripción del concepto: lo que viene después de "incluye"
    # (retiro de escombro, cimbra, acarreos...) es alcance, no la función.
    t = re.split(r"\bincluye\b|\bincluyendo\b", _plano(texto))[0]
    return {
        "funcion": {n for n, p in _FUNCION if re.search(p, t)},
        "material": {n for n, p in _MATERIAL if re.search(p, t)},
    }


def alcance_distinto(cotizado: str, referencia: str):
    """Regresa el motivo si el alcance o material de la referencia no es comparable.
    Revisa la descripción completa: objeto, función, material y alcance."""
    original = _plano(cotizado)
    candidato = _plano(referencia)
    if not original or not candidato:
        return None
    fc, fr = funcion_material(original), funcion_material(candidato)
    # En acabados (pintura, estuco...) el soporte no cambia el trabajo; ahí
    # solo cuentan demolición, reparación y provisional.
    estructural = clave_objeto(original) in {"muro", "barda", "columna", "castillo", "cerramiento", "dala",
                                              "cadena", "losa", "firme", "trabe", "zapata", "pretil"}
    if not estructural:
        generales = {"demolición / retiro", "reparación", "provisional"}
        fc = {**fc, "funcion": fc["funcion"] & generales}
        fr = {**fr, "funcion": fr["funcion"] & generales}
    diferencia_funcion = fc["funcion"] ^ fr["funcion"]
    if diferencia_funcion:
        lado = "la referencia" if diferencia_funcion & fr["funcion"] else "la partida"
        return f"función distinta: {lado} es {', '.join(sorted(diferencia_funcion))}"
    if fc["material"] and fr["material"] and not (fc["material"] & fr["material"]) \
            and not ({"concreto armado"} & fc["material"] and {"concreto armado"} & fr["material"]):
        return (f"material distinto: la partida es {', '.join(sorted(fc['material']))} y la referencia "
                f"{', '.join(sorted(fr['material']))}")
    e_cot, e_ref = elemento_principal(original), elemento_principal(candidato)
    if e_cot and e_ref and e_cot != e_ref and {e_cot, e_ref} != {"castillo", "columna"}:
        return f"elemento distinto: la partida es {e_cot} y la referencia es {e_ref}"
    if e_cot in ("castillo", "cerramiento", "columna") and not e_ref:
        return f"la referencia no es un {e_cot} (otro elemento)"
    # Qué se suministra o ejecuta: el OBJETO principal de la referencia debe
    # ser el mismo de la partida, no una palabra mencionada de paso
    # ("cónsula modular, anclaje a columna" no es una columna).
    clave = clave_objeto(original)
    if clave:
        sinonimos = _GRUPOS_OBJETO.get(clave, {clave})
        # Las dos primeras palabras con significado: el objeto, no un
        # complemento ("cónsula modular, anclaje a columna").
        cabeza = " ".join([w for w in _cabeza(candidato, 8).split() if w not in _RELLENO][:1])
        if not any(re.search(rf"\b{re.escape(x)}(e?s)?\b", cabeza) for x in sinonimos):
            return (f"lo que se suministra o ejecuta es otro: la partida es «{clave}» y la referencia es "
                    f"«{' '.join(_cabeza(candidato, 3).split())}…»")
    t_cot, t_ref = trabajo_principal(original), trabajo_principal(candidato)
    if t_cot and t_ref and t_cot != t_ref:
        return f"trabajo distinto: la partida es {t_cot} y la referencia es {t_ref}"
    if t_cot and not t_ref:
        return f"la referencia no es {t_cot} (otro trabajo)"
    if "estuco" in original and "estuco" not in candidato:
        return "material distinto: la partida es estuco y la referencia no"
    # Material distinto: p. ej. "columnas para amarrar barda" (concreto)
    # contra "bases para columnas metálicas" (acero). No requiere IA.
    if _METAL.search(candidato) and not _METAL.search(original) and _MAMPOSTERIA.search(original):
        return "material distinto: la referencia es metálica y la partida es de concreto/mampostería"
    if _METAL.search(original) and not _METAL.search(candidato) and _MAMPOSTERIA.search(candidato):
        return "material distinto: la partida es metálica y la referencia es de concreto/mampostería"
    instala = ("colocacion", "instalacion", "aplicacion", "bombeo", "mano de obra")
    if "suministro" in original and not any(p in original for p in instala) and any(
        p in candidato for p in instala
    ):
        return "la cotización es solo suministro y la referencia incluye instalación"
    if "mano de obra" in original and "material" in original and "solo mano de obra" in candidato:
        return "la referencia es solo mano de obra"
    if "estuco" in original and "estuco" not in candidato and re.search(r"aplanado|mortero|repellado|zarpeo", candidato):
        return "material distinto: la partida es estuco y la referencia es aplanado de mortero"
    if "premezclado" in original and "hecho en obra" in candidato:
        return "concreto premezclado vs. hecho en obra"
    if "vinilica" in original and "esmalte" in candidato and "vinilica" not in candidato:
        return "pintura vinílica vs. esmalte"
    return None


# ----------------------------------------------------------------------
# Especificaciones: la referencia puede traer sección, espesor, f'c,
# calibre o diámetro que el proveedor NO declaró. En ese caso el concepto
# puede ser el mismo, pero la equivalencia de precio no está demostrada.
# ----------------------------------------------------------------------
def _texto_especificaciones(texto: str) -> str:
    t = _plano(texto).upper()
    # Medidas comerciales equivalentes declaradas con otro nombre.
    t = re.sub(r"\bBLOCK (DEL |DE )?(NUM|NUMERO|NO|#)\.? ?6\b", "BLOCK 15 X 20 X 40", t)
    t = re.sub(r"\bBLOCK (DEL |DE )?(NUM|NUMERO|NO|#)\.? ?4\b", "BLOCK 10 X 20 X 40", t)
    t = re.sub(r"\bBLOCK (DEL |DE )?(NUM|NUMERO|NO|#)\.? ?8\b", "BLOCK 20 X 20 X 40", t)
    return t


def extraer_especificaciones(texto: str) -> dict:
    t = _texto_especificaciones(texto)
    specs = {}
    dims = re.findall(r"(\d+(?:\.\d+)?)\s*X\s*(\d+(?:\.\d+)?)(?:\s*X\s*(\d+(?:\.\d+)?))?", t)
    if dims:
        specs["sección / medidas"] = {"X".join(x for x in d if x) for d in dims}
    esp = re.findall(r"(\d+(?:\.\d+)?)\s*(MM|CM|CMS)\.?\s*DE\s+ESPESOR", t) + [
        (m[1], m[2]) for m in re.findall(r"(ESPESOR\s*(?:DE)?\s*)(\d+(?:\.\d+)?)\s*(MM|CM)", t)
    ]
    if esp:
        specs["espesor"] = {f"{float(v) / (10 if u == 'MM' else 1):g} CM" for v, u in esp}
    fc = re.findall(r"F\s*'?\s*C\s*=?\s*(\d{3})", t)
    if fc:
        specs["resistencia f'c"] = set(fc)
    cal = re.findall(r"\bCAL(?:IBRE)?\.?\s*(\d{1,2})\b", t)
    if cal:
        specs["calibre"] = set(cal)
    # Capacidad / potencia de equipos: dos bombas o dos grúas no se comparan
    # si no tienen la misma capacidad.
    cap = re.findall(r"(\d+(?:[.,]\d+)?)\s*(HP|H\.P\.|KW|KVA|WATTS?|W|TON|TONS|TONELADAS|BTU|LTS?|LITROS|"
                     r"GPM|LPS|M3/H|AMPERES|AMP|V|VOLTS?|KG/CM2)\b", t)
    # f'c va aparte; "V" solo como voltaje real (110/127/220/440).
    cap = [(v, u) for v, u in cap
           if u != "KG/CM2" and not (u == "V" and float(v.replace(",", ".")) < 100)]
    if cap:
        norm = {"H.P.": "HP", "TONS": "TON", "TONELADAS": "TON", "LT": "L", "LTS": "L", "LITROS": "L",
                "WATT": "W", "WATTS": "W", "AMPERES": "A", "AMP": "A", "VOLT": "V", "VOLTS": "V"}
        specs["capacidad"] = {f"{float(v.replace(',', '.')):g} {norm.get(u, u)}" for v, u in cap}
    dia = re.findall(r"(\d+(?:\s\d+)?/\d+|\d+(?:\.\d+)?)\s*(\"|”|''|PULG|PULGADAS)", t)
    dia += [(v, "MM") for v in re.findall(r"(\d+(?:\.\d+)?)\s*MM\.?\s*(?:DE\s+)?DIAMETRO", t)]
    if dia:
        specs["diámetro"] = {(v + " PULG") if u != "MM" else (v + " MM") for v, u in dia}
    return specs


# Especificaciones que definen el TAMAÑO de una pieza o equipo.
SPECS_TAMANO = ("sección / medidas", "capacidad", "diámetro")


def _dims_compatibles(a: str, b: str) -> bool:
    """'15X20' y '15X20X40' son compatibles (texto truncado); '10X20' no."""
    x, y = a.split("X"), b.split("X")
    corto, largo = (x, y) if len(x) <= len(y) else (y, x)
    return largo[:len(corto)] == corto


# Inclusiones que cambian el precio: si la referencia las trae y el
# proveedor no las menciona, la equivalencia no está demostrada.
_INCLUSIONES = (
    ("refuerzo horizontal", r"REFUERZO HORIZONTAL|ESCALERILLA|ESCALERA DE ACERO"),
    ("castillos ahogados", r"CASTILLOS? AHOGADOS?"),
    ("acabado aparente", r"ACABADO APARENTE|APARENTE"),
    ("cimbra", r"\bCIMBRA\b"),
)
_AZOTEA = r"AZOTEA|PLACA|LOSA EXISTENTE|SOBRE LOSA|PRETIL|EN ALTURA|ELEVACION"


def comparar_especificaciones(cotizado: str, referencia: str):
    """Regresa (faltantes, conflictos): specs de la referencia que el
    proveedor no declaró, y specs declaradas con valor distinto."""
    ref = extraer_especificaciones(referencia)
    cot = extraer_especificaciones(cotizado)
    faltantes, conflictos = [], []
    t_ref, t_cot = _texto_especificaciones(referencia), _texto_especificaciones(cotizado)
    for nombre, patron in _INCLUSIONES:
        if re.search(patron, t_ref) and not re.search(patron, t_cot):
            faltantes.append(nombre)
    if re.search(_AZOTEA, t_cot) and not re.search(_AZOTEA + r"|ALTURA|ELEVACIONES|ACARREO", t_ref):
        faltantes.append("condición de azotea/losa (la referencia no la contempla)")
    dims_cot = {p for d in cot.get("sección / medidas", set()) for p in d.split("X")}
    for campo, valores in ref.items():
        if campo == "espesor" and dims_cot and any(v.replace(" CM", "") in dims_cot for v in valores):
            continue  # "block 15x20x40" ya declara 15 cm de espesor
        if campo not in cot:
            faltantes.append(f"{campo} {', '.join(sorted(valores))}")
        elif campo == "sección / medidas" and any(_dims_compatibles(a, b) for a in valores for b in cot[campo]):
            continue
        elif not (valores & cot[campo]):
            conflictos.append(f"{campo}: cotizado {', '.join(sorted(cot[campo]))} vs referencia {', '.join(sorted(valores))}")
    return faltantes, conflictos


def _meses_desde(fecha_dato):
    if not fecha_dato:
        return None
    texto = str(fecha_dato)
    m = re.match(r"(\d{4})(?:-(\d{1,2}))?", texto)
    if not m:
        return None
    anio, mes = int(m.group(1)), int(m.group(2) or 6)
    hoy = _dt.date.today()
    return (hoy.year - anio) * 12 + (hoy.month - mes)


def evaluar_fuente(
    clave: str,
    fuente: dict | None,
    *,
    concepto: str,
    precio: float,
    usar_ia: bool,
    precio_referencia: float | None = None,
    unidad: str | None = None,
    fecha_dato=None,
    region: str | None = None,
    es_web: bool = False,
    web_verificada: bool = False,
    anio_dato=None,
    fecha_consulta=None,
) -> dict:
    """Decide el estado de UNA referencia y calcula su comparación.

    VALIDADA exige TODO lo siguiente (confirmar el concepto no basta):
      1. La IA confirmó que es el mismo concepto.
      2. Las especificaciones de la referencia (sección, espesor, f'c,
         calibre) están declaradas por el proveedor y coinciden.
      3. El dato es vigente (máx. VIGENCIA_MESES meses).
      4. En web: fuente y frase del precio verificadas.
    Si el concepto coincide pero falta alguna de 2-4: EQUIVALENCIA PARCIAL
    (orientativa). La escala (5×) solo filtra cifras extremas.
    """
    fecha_dato = fecha_dato or anio_dato
    salida = {
        "fuente": clave,
        "nombre": NOMBRE_FUENTE.get(clave, clave),
        "descripcion": None,
        "precio_referencia": None,
        "fecha": str(fecha_dato)[:10] if fecha_dato else None,
        "fecha_consulta": str(fecha_consulta)[:10] if fecha_consulta else None,
        "region": region,
        "estado": SIN_DATO,
        "motivo": "la fuente no encontró un concepto comparable",
        "clasificacion": None,
        "diferencia_pct": None,
        "diferencia_unitaria": None,
        "revision_ia": None,
        "confianza": None,
        "cuenta": False,
    }
    if not fuente or not fuente.get("match") or not precio_referencia or precio_referencia <= 0:
        if fuente and fuente.get("motivo"):
            salida["motivo"] = str(fuente["motivo"])
        return salida

    salida["descripcion"] = fuente.get("match")
    salida["precio_referencia"] = round(float(precio_referencia), 2)
    salida["confianza"] = fuente.get("confianza")
    salida["clasificacion"] = clasificar(precio, precio_referencia)
    salida["diferencia_pct"] = diferencia_pct(precio, precio_referencia)
    salida["diferencia_unitaria"] = round(precio - precio_referencia, 2)

    def _excluir(estado, motivo):
        salida.update(estado=estado, motivo=motivo, clasificacion=None,
                      diferencia_pct=None, diferencia_unitaria=None)
        return salida

    revision = fuente.get("revision_ia") or {}
    veredicto = str(revision.get("veredicto") or "").upper() if usar_ia else ""
    if veredicto:
        salida["revision_ia"] = f"{veredicto}: {revision.get('razon', '')}".strip()

    # Regla 1: RECHAZA excluye la referencia de todo cálculo.
    if veredicto == "RECHAZA":
        return _excluir(RECHAZADA, f"la IA la rechazó: {revision.get('razon', '')}".strip())

    motivo_alcance = alcance_distinto(concepto, fuente.get("match"))
    if motivo_alcance:
        return _excluir(RECHAZADA, motivo_alcance)

    faltantes, conflictos = comparar_especificaciones(concepto, fuente.get("match"))
    if conflictos:
        return _excluir(RECHAZADA, "especificación distinta: " + "; ".join(conflictos))

    if precio_referencia > precio * FACTOR_ESCALA or precio_referencia < precio / FACTOR_ESCALA:
        salida.update(estado=NO_COMPARABLE, clasificacion=None, diferencia_pct=None, diferencia_unitaria=None,
                      motivo="precio fuera de escala frente al cotizado (otra unidad o alcance)")
        return salida

    # Precio por pieza: dos piezas solo se comparan si se sabe su tamaño. Si
    # la referencia trae medidas y la partida no, o difieren, no se compara.
    # Piezas, equipos, juegos, lotes y servicios: dos "bombas" o dos "piezas"
    # pueden ser muy distintas. Si la referencia define tamaño o capacidad y
    # la cotización no lo declara, no se compara.
    u = _plano(unidad or "").replace(".", "").strip()
    if u in ("pza", "pz", "pzas", "pieza", "piezas", "un", "unidad", "jgo", "juego", "lote", "servicio",
             "serv", "jornal", "equipo", "kit", "sal", "salida"):
        spec_ref = extraer_especificaciones(fuente.get("match"))
        spec_cot = extraer_especificaciones(concepto)
        faltan = [f"{c} {', '.join(sorted(spec_ref[c]))}" for c in SPECS_TAMANO
                  if spec_ref.get(c) and not spec_cot.get(c)]
        sin_dato_ref = [f"{c} {', '.join(sorted(spec_cot[c]))}" for c in SPECS_TAMANO
                        if spec_cot.get(c) and not spec_ref.get(c)]
        if sin_dato_ref:
            salida.update(estado=NO_COMPARABLE, clasificacion=None, diferencia_pct=None, diferencia_unitaria=None,
                          motivo=f"precio por {u}: la cotización es de {'; '.join(sin_dato_ref)} y la referencia "
                                 "no indica ese tamaño/capacidad; no se puede comparar")
            return salida
        if faltan:
            salida.update(estado=NO_COMPARABLE, clasificacion=None, diferencia_pct=None, diferencia_unitaria=None,
                          motivo=f"precio por {u}: la referencia es de {'; '.join(faltan)} y la cotización no "
                                 "lo declara; pedir dimensiones/capacidad antes de comparar")
            return salida

    # Regla 2: NO_SEGURO se conserva como orientativa.
    if veredicto == "NO_SEGURO":
        salida.update(estado=NO_CONCLUYENTE,
                      motivo=f"la IA no pudo confirmarla: {revision.get('razon', '')}".strip())
        return salida

    if es_web:
        if not web_verificada:
            salida.update(estado=NO_CONCLUYENTE,
                          motivo="fragmento web: precio orientativo, sin validar alcance ni fecha")
            return salida
    elif veredicto != "CONFIRMA":
        confianza = str(fuente.get("confianza") or "").upper() or "ALTA"
        salida.update(
            estado=POR_CONFIRMAR,
            motivo=(
                f"coincidencia de texto {confianza}; la IA no ha confirmado el concepto"
                + ("" if usar_ia else " (revisión con IA desactivada)")
            ),
        )
        return salida

    # Regla 3: concepto confirmado. Falta demostrar especificación y vigencia.
    pendientes = []
    if faltantes:
        pendientes.append("el proveedor no declara " + "; ".join(faltantes))
    meses = _meses_desde(fecha_dato)
    if meses is None:
        pendientes.append("fecha del precio no disponible")
    elif meses > VIGENCIA_MESES:
        pendientes.append(
            f"precio de {str(fecha_dato)[:7]} ({meses} meses): el ajuste por INPC no demuestra vigencia"
        )
    confirmado = "concepto confirmado por IA" if veredicto == "CONFIRMA" else "fuente y frase web verificadas"
    if pendientes:
        salida.update(estado=EQUIVALENCIA_PARCIAL, motivo=f"{confirmado}, pero " + "; ".join(pendientes))
        return salida

    notas = [confirmado, "especificaciones compatibles", f"precio vigente ({str(fecha_dato)[:7]})"]
    if region and region not in ("NL", "RAGASA"):
        notas.append(f"región {region}: considerar diferencia regional")
    salida.update(estado=VALIDADA, motivo=", ".join(notas), cuenta=True)
    return salida


def resultado_final(evaluaciones: dict, precio: float, cantidad) -> dict:
    """Semáforo final por consenso de referencias VALIDADAS, sin promediar.

    - Validadas que coinciden        -> ese semáforo; la diferencia es ahorro respaldado.
    - Validadas en desacuerdo        -> MIXTO.
    - Sin validadas                  -> NO CONCLUYENTE; se reporta la
      "diferencia contra referencia" de la mejor referencia orientativa con
      concepto confirmado (EQUIVALENCIA PARCIAL), marcada como pendiente de
      validar, nunca como ahorro.
    """
    try:
        cant = float(cantidad)
    except (TypeError, ValueError):
        cant = None
    validadas = [evaluaciones[k] for k in PRIORIDAD_NEGOCIACION
                 if k in evaluaciones and evaluaciones[k]["estado"] == VALIDADA]
    parciales = [evaluaciones[k] for k in PRIORIDAD_NEGOCIACION
                 if k in evaluaciones and evaluaciones[k]["estado"] == EQUIVALENCIA_PARCIAL]
    orientativas = [e for e in evaluaciones.values() if e["estado"] in ORIENTATIVAS]

    salida = {
        "semaforo": SIN_VALIDADA,
        "fuentes_validadas": ", ".join(e["nombre"] for e in validadas) or None,
        "referencia_negociacion": None,
        "precio_negociacion": None,
        "diferencia_pct": None,
        "diferencia_importe": None,
        "respaldo": None,
        "ahorro_potencial": None,
        "detalle": None,
    }

    base = validadas[0] if validadas else (parciales[0] if parciales else None)
    if base:
        salida["referencia_negociacion"] = base["nombre"]
        salida["precio_negociacion"] = base["precio_referencia"]
        salida["diferencia_pct"] = diferencia_pct(precio, base["precio_referencia"])
        if cant:
            salida["diferencia_importe"] = round((precio - base["precio_referencia"]) * cant, 2)

    if validadas:
        clases = {e["clasificacion"] for e in validadas}
        salida["semaforo"] = clases.pop() if len(clases) == 1 else MIXTO
        salida["respaldo"] = "VALIDADA"
        # Solo es ahorro si la partida quedó ALTA (fuera del ±5 %); dentro de
        # mercado la diferencia es ruido, no un monto a negociar.
        if salida["diferencia_importe"] is not None:
            salida["ahorro_potencial"] = (
                max(0.0, salida["diferencia_importe"]) if salida["semaforo"] == ALTO else 0.0
            )
        if salida["semaforo"] == MIXTO:
            salida["detalle"] = " / ".join(f"{e['nombre']}: {e['clasificacion']}" for e in validadas)
    elif orientativas:
        salida["semaforo"] = NO_CONCLUYENTE
        salida["respaldo"] = "PENDIENTE DE VALIDAR" if base else None
        salida["detalle"] = (
            f"diferencia contra {base['nombre']} pendiente de validar: {base['motivo']}"
            if base else "solo hay referencias sin concepto confirmado; revisar antes de negociar"
        )
    return salida


# ----------------------------------------------------------------------
# Presentación común (pantalla, CSV y Excel): mismo texto en todos lados.
# ----------------------------------------------------------------------
_CLAS_TEXTO = {"ALTO": "caro", "EN MERCADO": "en precio", "BAJO": "barato"}


def estado_simple(ev: dict) -> str:
    """Validada · Orientativa · No comparable · Rechazada · Sin dato."""
    estado = ev.get("estado")
    if estado == VALIDADA:
        return "Validada"
    if estado in ORIENTATIVAS and ev.get("precio_referencia"):
        return "Orientativa"
    if estado == NO_COMPARABLE:
        return "No comparable"
    if estado == RECHAZADA:
        return "Rechazada"
    return "Sin dato"


def dictamen_texto(ev: dict) -> str:
    """'Caro' (validada), 'Posiblemente caro' (orientativa), 'No comparable'..."""
    simple = estado_simple(ev)
    clas = _CLAS_TEXTO.get(ev.get("clasificacion") or "", "")
    if simple == "Validada" and clas:
        return clas[0].upper() + clas[1:]
    if simple == "Orientativa" and clas:
        return f"Posiblemente {clas}"
    return simple


def confiabilidad_fuente(ev: dict) -> str:
    """Qué tan confiable es la FUENTE (no si el concepto es equivalente)."""
    if not ev.get("precio_referencia") and ev.get("estado") == SIN_DATO:
        return "—"
    clave = ev.get("fuente")
    if clave == "historico":
        return "Alta (interna Ragasa)"
    if clave == "nl":
        return "Alta (oficial, licitaciones SIASI/OCDS)"
    if clave == "cdmx":
        return "Alta (oficial, tabulador)"
    if clave == "ia":
        url = str(ev.get("url") or "").lower()
        if ev.get("origen") == "fragmento del buscador":
            return "Baja (solo fragmento, documento no verificado)"
        if ".gob.mx" in url or ".gob/" in url:
            return "Alta (documento oficial)"
        return "Media (página comercial)" if url else "Baja"
    return "—"


def confiabilidad_texto(ev: dict) -> str:
    """Compatibilidad: confiabilidad de la fuente."""
    return confiabilidad_fuente(ev)


def falta_confirmar(ev: dict, concepto: str) -> str:
    """Qué hay que confirmar con el proveedor para validar la referencia."""
    simple = estado_simple(ev)
    if simple == "Validada":
        return ""
    if simple in ("Rechazada", "No comparable", "Sin dato") or not ev.get("descripcion"):
        return str(ev.get("motivo") or "sin concepto comparable")
    faltantes, _ = comparar_especificaciones(concepto, ev.get("descripcion"))
    partes = list(faltantes)
    motivo = str(ev.get("motivo") or "")
    if ev.get("estado") == POR_CONFIRMAR:
        partes.append("que sea el mismo concepto (revisión con IA)")
    if clave_objeto(concepto) == "cerramiento" and re.match(r"\W*(dala|cadena)", _cabeza(ev.get("descripcion"), 2)):
        partes.append("que la dala/cadena de la referencia sea de cerramiento (superior) y no de desplante")
    if ev.get("fuente") == "ia":
        partes.append("alcance y fecha del precio de la página")
    elif "meses" in motivo or "fecha del precio no disponible" in motivo:
        partes.append("vigencia del precio")
    return "; ".join(dict.fromkeys(p for p in partes if p)) or motivo


def alcance_precio(ev: dict) -> str:
    """Qué incluye el precio de referencia, según su texto y su tipo de fuente."""
    t = _plano(ev.get("descripcion") or "")
    partes = []
    if "suministro" in t:
        partes.append("suministro")
    if re.search(r"colocacion|instalacion|aplicacion|fabricacion|construccion|elaborad", t):
        partes.append("colocación/instalación")
    if "mano de obra" in t:
        partes.append("mano de obra")
    inc = re.search(r"incluye:?\s*([^.]{0,140})", t)
    if inc:
        partes.append("incluye " + inc.group(1).strip())
    base = {
        "nl": (("precio de contratos adjudicados de obra pública de NL (registros OCDS marcados "
                "'contrato_adjudicado')" if "contrato_adjudicado" in str(ev.get("tipo_registros") or "")
                else "precio unitario de licitación de obra pública de NL (tipo de precio no documentado)")
               + "; P.U. de obra pública: costo directo + indirectos + utilidad, IVA aparte"),
        "cdmx": "precio unitario del tabulador oficial (costo directo + indirectos + utilidad), IVA aparte",
        "historico": "precio cotizado a Ragasa (cotización recibida), sin IVA",
        "ia": "precio publicado en la página, llevado a sin IVA; alcance no confirmado",
    }.get(ev.get("fuente"), "")
    return "; ".join(x for x in [", ".join(partes) if partes else "", base] if x)


def cuenta_filtro(ev: dict) -> bool:
    """El filtro aporta un dictamen: validada u orientativa, con precio comparable."""
    return estado_simple(ev) in ("Validada", "Orientativa") and bool(ev.get("clasificacion"))


def resultado_filtros(evaluaciones: dict) -> dict:
    """Lectura simple de los 4 filtros: cuántos dicen caro / en precio /
    barato y cuántos están validados. No promedia precios. Sin validadas,
    el resultado se presenta como 'Posiblemente...' (no es concluyente)."""
    usados = [e for e in evaluaciones.values() if cuenta_filtro(e)]
    conteo = {c: sum(1 for e in usados if e["clasificacion"] == c) for c in (ALTO, EN_MERCADO, BAJO)}
    validados = sum(1 for e in usados if e["estado"] == VALIDADA)
    n = len(usados)
    detalle = f"{validados} de {n} validada{'s' if n != 1 else ''}" if n else "ningún filtro comparable"
    if not n:
        return {"clave": None, "texto": "⚪ Sin datos", "texto_plano": "Sin datos", "detalle": detalle,
                "n": 0, "conteo": conteo, "validados": 0}
    mayor = max(conteo.values())
    ganadores = [c for c, k in conteo.items() if k == mayor]
    icono = {ALTO: "🔴", EN_MERCADO: "🟡", BAJO: "🟢"}
    if len(ganadores) > 1:
        partes = [f"{conteo[c]} {_CLAS_TEXTO[c]}" for c in (ALTO, EN_MERCADO, BAJO) if conteo[c]]
        clave, plano = MIXTO, "No coinciden: " + ", ".join(partes)
        texto = "🟠 " + plano
    else:
        clave = ganadores[0]
        palabra = _CLAS_TEXTO[clave]
        palabra = palabra[0].upper() + palabra[1:] if validados else f"Posiblemente {palabra}"
        plano = f"{palabra} en {mayor} de {n} filtro{'s' if n > 1 else ''}"
        texto = f"{icono[clave]} {plano}"
    return {"clave": clave, "texto": texto, "texto_plano": plano, "detalle": detalle,
            "n": n, "conteo": conteo, "validados": validados}


def evidencia(ev: dict, concepto: str) -> dict:
    """Campos de evidencia en el mismo orden para pantalla, CSV y Excel."""
    inf = ev.get("inflacion") or {}
    verificacion = ""
    if ev.get("fuente") == "ia" and ev.get("precio_referencia"):
        verificacion = ("solo fragmento del buscador (documento completo no verificado)"
                        if ev.get("origen") == "fragmento del buscador"
                        else "precio leído en el documento completo")
    elif ev.get("precio_referencia"):
        verificacion = "base de datos cargada en la app"
    return {
        "Estado": estado_simple(ev),
        "Motivo": ev.get("motivo") or "",
        "Dictamen": dictamen_texto(ev),
        "Confiabilidad de la fuente": confiabilidad_fuente(ev),
        "Falta confirmar": falta_confirmar(ev, concepto),
        "Descripción completa de la referencia": ev.get("descripcion") or "",
        "Unidad de la referencia": ev.get("unidad_ref") or "",
        "Qué incluye el precio": alcance_precio(ev) if ev.get("descripcion") else "",
        "Documento / fuente": ev.get("documento") or "",
        "Enlace": ev.get("url") or "",
        "Código": ev.get("codigo") or "",
        "Página": ev.get("pagina") or "",
        "Región": ev.get("region") or "",
        "Fecha / periodo del precio": ev.get("periodo") or ev.get("fecha") or (
            "no indicada" if ev.get("precio_referencia") else ""),
        "Fecha de consulta": ev.get("fecha_consulta") or "",
        "Verificación": verificacion,
        "Precio original": ev.get("precio_original"),
        "Índice de inflación": inf.get("indice", ""),
        "Periodo base": inf.get("periodo_base", ""),
        "Valor base": inf.get("valor_base"),
        "Periodo final": inf.get("periodo_final", ""),
        "Valor final": inf.get("valor_final"),
        "Factor": inf.get("factor"),
        "Precio usado (actualizado)": ev.get("precio_referencia"),
        "Justificación del periodo base": inf.get("justificacion", ""),
        "Nota de inflación": ("el ajuste por inflación no confirma vigencia comercial ni equivalencia técnica"
                              if inf else ""),
        "Registros de origen": (
            f"{ev.get('n_renglones')} renglones en "
            f"{len({(r.get('ocid') or r.get('licitacion')) for r in ev.get('registros_todos')})} contrato(s) (OCID); "
            "detalle completo en la hoja 'Inflación NL'" if ev.get("registros_todos") else "; ".join(
                f"{r.get('fecha')} · {r.get('licitacion') or ''} · {r.get('dependencia') or ''} · ${r.get('precio'):,.2f}"
                for r in (ev.get("registros_detalle") or [])[:5])),
        "Referencias rechazadas": "; ".join(
            f"{r.get('codigo') or ''} {str(r.get('concepto'))[:80]} (${r.get('precio') or 0:,.2f}"
            f"{', pág. ' + str(r.get('pagina')) if r.get('pagina') else ''}) — {r.get('motivo')}"
            for r in (ev.get("rechazadas") or [])[:5]),
        "Revisión IA": ev.get("revision_ia") or "",
        "Evidencia (frase)": ev.get("evidencia") or "",
    }
