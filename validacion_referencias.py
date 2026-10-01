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


def alcance_distinto(cotizado: str, referencia: str):
    """Regresa el motivo si el alcance o material de la referencia no es comparable."""
    original = _plano(cotizado)
    candidato = _plano(referencia)
    if not original or not candidato:
        return None
    e_cot, e_ref = elemento_principal(original), elemento_principal(candidato)
    if e_cot and e_ref and e_cot != e_ref and {e_cot, e_ref} != {"castillo", "columna"}:
        return f"elemento distinto: la partida es {e_cot} y la referencia es {e_ref}"
    if e_cot in ("castillo", "cerramiento", "columna") and not e_ref:
        return f"la referencia no es un {e_cot} (otro elemento)"
    t_cot, t_ref = trabajo_principal(original), trabajo_principal(candidato)
    if t_cot and t_ref and t_cot != t_ref:
        return f"trabajo distinto: la partida es {t_cot} y la referencia es {t_ref}"
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
    return specs


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
        salida.update(estado=NO_CONCLUYENTE,
                      motivo="precio fuera de escala frente al cotizado (otra unidad o alcance)")
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
_CLAS_TEXTO = {"ALTO": "caro", "EN MERCADO": "en mercado", "BAJO": "barato"}


def dictamen_texto(ev: dict) -> str:
    """'Validado · caro', 'Orientativo · caro', 'Rechazada', 'Sin referencia'."""
    estado, clas = ev.get("estado"), _CLAS_TEXTO.get(ev.get("clasificacion") or "", "")
    if estado == VALIDADA:
        return f"Validado · {clas}" if clas else "Validado"
    if estado in ORIENTATIVAS:
        return f"Orientativo · {clas}" if clas else "Orientativo"
    if estado == RECHAZADA:
        return "Rechazada"
    return "Sin referencia"


def confiabilidad_texto(ev: dict) -> str:
    """Confiabilidad del dato. Una referencia no validada nunca sale 'alta'."""
    if ev.get("estado") in (SIN_DATO, RECHAZADA) or not ev.get("precio_referencia"):
        return "—"
    if ev.get("estado") == VALIDADA:
        return "Alta (validada)"
    conf = str(ev.get("confianza") or "").upper()
    if ev.get("fuente") == "ia":
        return "Media" if conf == "MEDIA" else "Baja"
    return {"ALTA": "Media", "MEDIA": "Media", "BAJA": "Baja"}.get(conf, "Baja")


def falta_confirmar(ev: dict, concepto: str) -> str:
    """Qué hay que confirmar con el proveedor para validar la referencia."""
    if ev.get("estado") == VALIDADA:
        return ""
    if ev.get("estado") == RECHAZADA:
        return str(ev.get("motivo") or "")
    if ev.get("estado") == SIN_DATO or not ev.get("descripcion"):
        return str(ev.get("motivo") or "sin concepto comparable")
    faltantes, _ = comparar_especificaciones(concepto, ev.get("descripcion"))
    partes = list(faltantes)
    motivo = str(ev.get("motivo") or "")
    if ev.get("estado") == POR_CONFIRMAR:
        partes.append("que sea el mismo concepto (revisión con IA)")
    if "fuera de escala" in motivo:
        partes.append("unidad/alcance (precio fuera de escala)")
    if ev.get("fuente") == "ia":
        partes.append("alcance, fecha del precio e IVA de la página")
    elif "meses" in motivo or "fecha del precio no disponible" in motivo:
        partes.append("vigencia del precio")
    return "; ".join(dict.fromkeys(p for p in partes if p)) or motivo
