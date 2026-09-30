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
    VALIDADA        cuenta para el semáforo final y la negociación
    NO CONCLUYENTE  orientativa: la IA no pudo confirmarla, o no pasó
                    alguna verificación (escala, fecha, fuente web)
    POR CONFIRMAR   coincidencia de texto débil que nadie ha revisado
    RECHAZADA       la IA la rechazó o el alcance es distinto
    SIN DATO        la fuente no encontró nada comparable
"""

from __future__ import annotations

import datetime as _dt
import unicodedata

VALIDADA = "VALIDADA"
NO_CONCLUYENTE = "NO CONCLUYENTE"
POR_CONFIRMAR = "POR CONFIRMAR"
RECHAZADA = "RECHAZADA"
SIN_DATO = "SIN DATO"

# Semáforo final
ALTO = "ALTO"
EN_MERCADO = "EN MERCADO"
BAJO = "BAJO"
MIXTO = "MIXTO"
SIN_VALIDADA = "SIN DATOS SUFICIENTES"

MARGEN_EN_MERCADO = 0.05          # ±5 % alrededor de la referencia
FACTOR_ESCALA = 5.0               # más de 5× (o menos de 1/5) = otra escala
ANTIGUEDAD_MAXIMA_ANIOS = 5       # dato más viejo = orientativo aunque se ajuste por INPC

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


def alcance_distinto(cotizado: str, referencia: str):
    """Regresa el motivo si el alcance de la referencia no es comparable."""
    original = _plano(cotizado)
    candidato = _plano(referencia)
    if not original or not candidato:
        return None
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


def evaluar_fuente(
    clave: str,
    fuente: dict | None,
    *,
    concepto: str,
    precio: float,
    usar_ia: bool,
    precio_referencia: float | None = None,
    anio_dato=None,
    es_web: bool = False,
    web_verificada: bool = False,
) -> dict:
    """Decide el estado de UNA referencia y calcula su comparación.

    Regresa dict con: fuente, nombre, descripcion, precio_referencia, estado,
    motivo, clasificacion, diferencia_pct, revision_ia, cuenta.
    """
    salida = {
        "fuente": clave,
        "nombre": NOMBRE_FUENTE.get(clave, clave),
        "descripcion": None,
        "precio_referencia": None,
        "estado": SIN_DATO,
        "motivo": "la fuente no encontró un concepto comparable",
        "clasificacion": None,
        "diferencia_pct": None,
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

    revision = fuente.get("revision_ia") or {}
    veredicto = str(revision.get("veredicto") or "").upper() if usar_ia else ""
    if veredicto:
        salida["revision_ia"] = f"{veredicto}: {revision.get('razon', '')}".strip()

    # Regla 1: RECHAZA excluye la referencia de todo cálculo.
    if veredicto == "RECHAZA":
        salida.update(
            estado=RECHAZADA,
            motivo=f"la IA la rechazó: {revision.get('razon', '')}".strip(),
            clasificacion=None,
            diferencia_pct=None,
        )
        return salida

    motivo_alcance = alcance_distinto(concepto, fuente.get("match"))
    if motivo_alcance:
        salida.update(estado=RECHAZADA, motivo=f"alcance distinto: {motivo_alcance}",
                      clasificacion=None, diferencia_pct=None)
        return salida

    # Regla 2: NO_SEGURO se conserva como orientativa.
    if veredicto == "NO_SEGURO":
        salida.update(estado=NO_CONCLUYENTE,
                      motivo=f"la IA no pudo confirmarla: {revision.get('razon', '')}".strip())
        return salida

    confianza = str(fuente.get("confianza") or "").upper()
    if confianza in ("BAJA", "MEDIA") and veredicto != "CONFIRMA":
        salida.update(
            estado=POR_CONFIRMAR,
            motivo=(
                f"coincidencia de texto {confianza} sin confirmar por la IA"
                + ("" if usar_ia else " (revisión con IA desactivada)")
            ),
        )
        return salida

    # Regla 3: aun con CONFIRMA (o confianza ALTA/MEDIA), verificar
    # escala del precio, antigüedad del dato y, en web, la fuente.
    if precio_referencia > precio * FACTOR_ESCALA or precio_referencia < precio / FACTOR_ESCALA:
        salida.update(estado=NO_CONCLUYENTE,
                      motivo="precio fuera de escala frente al cotizado (otra unidad o alcance)")
        return salida

    if anio_dato:
        try:
            anio = int(str(anio_dato)[:4])
            if _dt.date.today().year - anio > ANTIGUEDAD_MAXIMA_ANIOS:
                salida.update(estado=NO_CONCLUYENTE,
                              motivo=f"dato de {anio}: demasiado antiguo aunque se ajuste por INPC")
                return salida
        except ValueError:
            pass

    if es_web and not web_verificada:
        salida.update(estado=NO_CONCLUYENTE,
                      motivo="fragmento web: precio orientativo, sin validar alcance ni fecha")
        return salida

    verificaciones = ["unidad igual"]
    if anio_dato:
        verificaciones.append(f"dato {str(anio_dato)[:4]}")
    verificaciones.append("escala coherente")
    if veredicto == "CONFIRMA":
        verificaciones.insert(0, "concepto confirmado por IA")
    else:
        verificaciones.insert(0, f"coincidencia de texto {confianza or 'ALTA'}")
    salida.update(estado=VALIDADA, motivo=", ".join(verificaciones), cuenta=True)
    return salida


def resultado_final(evaluaciones: dict, precio: float, cantidad) -> dict:
    """Semáforo final por consenso de referencias VALIDADAS, sin promediar.

    - Todas las validadas coinciden        -> ese semáforo.
    - Validadas en desacuerdo              -> MIXTO (se muestra cada una).
    - Ninguna validada, pero hay orientativas -> NO CONCLUYENTE.
    - Nada                                  -> SIN DATOS SUFICIENTES.

    Precio de negociación: la referencia VALIDADA de mayor prioridad
    (Histórico Ragasa > Nuevo León > CDMX > IA internet).
    """
    validadas = [evaluaciones[k] for k in PRIORIDAD_NEGOCIACION
                 if k in evaluaciones and evaluaciones[k]["estado"] == VALIDADA]
    orientativas = [e for e in evaluaciones.values()
                    if e["estado"] in (NO_CONCLUYENTE, POR_CONFIRMAR)]

    salida = {
        "semaforo": SIN_VALIDADA,
        "fuentes_validadas": ", ".join(e["nombre"] for e in validadas) or None,
        "referencia_negociacion": None,
        "precio_negociacion": None,
        "diferencia_pct": None,
        "ahorro_potencial": None,
        "detalle": None,
    }

    if validadas:
        clases = {e["clasificacion"] for e in validadas}
        salida["semaforo"] = clases.pop() if len(clases) == 1 else MIXTO
        base = validadas[0]
        salida["referencia_negociacion"] = base["nombre"]
        salida["precio_negociacion"] = base["precio_referencia"]
        salida["diferencia_pct"] = diferencia_pct(precio, base["precio_referencia"])
        try:
            cant = float(cantidad)
        except (TypeError, ValueError):
            cant = None
        if cant and precio > base["precio_referencia"]:
            salida["ahorro_potencial"] = round((precio - base["precio_referencia"]) * cant, 2)
        elif cant:
            salida["ahorro_potencial"] = 0.0
        if salida["semaforo"] == MIXTO:
            salida["detalle"] = " / ".join(
                f"{e['nombre']}: {e['clasificacion']}" for e in validadas
            )
    elif orientativas:
        salida["semaforo"] = NO_CONCLUYENTE
        salida["detalle"] = "solo hay referencias orientativas; revisar antes de negociar"

    return salida
