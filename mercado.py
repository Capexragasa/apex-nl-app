"""
Datos de mercado para comparar
==============================

Aunque ninguna referencia quede VALIDADA (la regla es estricta a propósito),
Compras necesita NÚMEROS REALES contra los cuales ubicar el precio cotizado.
Este módulo junta, por partida, varias referencias publicadas:

  - Nuevo León: mediana y rango p25-p75 de precios realmente contratados en
    licitaciones de obra pública (SIASI), con número de registros y periodo,
    ajustados por INPC.
  - CDMX: precio del Tabulador General de Precios Unitarios 2026.

y dice, para cada una, si es del MISMO material que la partida (por texto)
o solo un concepto relacionado, y dónde cae el precio cotizado.

Para castillos/columnas cotizados POR PIEZA (las bases los cobran por metro
lineal), calcula el equivalente por pieza = precio por ml × altura del
castillo (altura del muro + cerramiento), dejando el supuesto a la vista.

Todo es ORIENTATIVO: datos reales, pero sin equivalencia demostrada. La
validación formal vive en validacion_referencias.py.
"""

from __future__ import annotations

import re
import unicodedata

from validacion_referencias import FACTOR_ESCALA, alcance_distinto

MATERIALES_CLAVE = (
    "ESTUCO", "YESO", "MORTERO", "BLOCK", "TABIQUE", "LADRILLO", "TABICON",
    "CERRAMIENTO", "CASTILLO", "DALA", "CADENA", "TRABE", "LOSA", "FIRME",
    "PINTURA", "IMPERMEABILIZANTE", "PISO", "AZULEJO", "CONCRETO", "ACERO",
    "MALLA", "PANEL", "TABLAROCA", "PLAFON",
)
ALTURA_CERRAMIENTO = 0.20   # m, supuesto para el alto total del castillo


def _plano(texto) -> str:
    t = unicodedata.normalize("NFKD", str(texto or "").upper())
    return "".join(c for c in t if not unicodedata.combining(c))


def _materiales(texto: str) -> set:
    t = _plano(texto)
    t = re.sub(r"\bCOLUMNAS? (PARA AMARRAR|DE AMARRE|DE BARDA)\b", "CASTILLO", t)
    return {m for m in MATERIALES_CLAVE if re.search(rf"\b{m}(ES|S)?\b", t)}


def _relacion(concepto_cotizado: str, concepto_referencia: str) -> str:
    propios = _materiales(concepto_cotizado)
    ajenos = _materiales(concepto_referencia)
    if not propios:
        return "sin material identificable"
    principal = next((m for m in MATERIALES_CLAVE if m in propios and m != "CONCRETO"), None)
    if principal and principal in ajenos:
        return "mismo material"
    if propios & ajenos:
        return "mismo material"
    return "concepto relacionado (otro material)"


def _posicion(precio: float, ref: dict) -> str:
    bajo, alto, med = ref.get("rango_bajo"), ref.get("rango_alto"), ref.get("precio")
    if bajo and alto and alto > bajo:
        if precio > alto:
            return f"encima del rango ({(precio / alto - 1) * 100:+.0f}% vs p75)"
        if precio < bajo:
            return f"debajo del rango ({(precio / bajo - 1) * 100:+.0f}% vs p25)"
        return "dentro del rango"
    if med:
        return f"{(precio / med - 1) * 100:+.0f}% vs referencia"
    return ""


def datos_mercado(comparador, partidas: list[dict], altura_muro=None, k: int = 3) -> dict:
    """partidas: dicts con partida, concepto, unidad, cantidad, precio_unitario.

    Regresa {'referencias': [...filas...], 'resumen': [...una por partida...]}.
    """
    referencias, resumen = [], []
    for p in partidas:
        concepto = str(p.get("concepto") or "")
        unidad = str(p.get("unidad") or "")
        try:
            precio = float(p.get("precio_unitario"))
        except (TypeError, ValueError):
            continue
        filas = []
        for ref in comparador.candidatos(concepto, unidad, k=k):
            ref = dict(ref)
            ref["relacion"] = _relacion(concepto, ref["concepto"])
            ref["equivalencia"] = "misma unidad"
            filas.append(ref)

        # Castillos/columnas por pieza: convertir referencias por ml.
        es_pieza = unidad.strip().upper() in ("PZA", "PZ", "PIEZA", "PIEZAS", "PZAS")
        if es_pieza and "CASTILLO" in _materiales(concepto) and altura_muro:
            altura_total = round(float(altura_muro) + ALTURA_CERRAMIENTO, 2)
            for ref in comparador.candidatos("CASTILLO DE CONCRETO 15 X 15 ARMEX", "ML", k=k):
                ref = dict(ref)
                if "CASTILLO" not in _materiales(ref["concepto"]):
                    continue
                factor = altura_total
                ref["precio_ml"] = ref["precio"]
                ref["precio"] = round(ref["precio"] * factor, 2)
                for campo in ("rango_bajo", "rango_alto"):
                    if ref.get(campo):
                        ref[campo] = round(ref[campo] * factor, 2)
                ref["unidad"] = "PZA (escenario)"
                ref["relacion"] = "escenario por pieza"
                ref["equivalencia"] = (
                    f"${ref['precio_ml']:,.2f}/ml × {altura_total:g} m de alto "
                    f"(muro {float(altura_muro):g} m + cerramiento {ALTURA_CERRAMIENTO:g} m, supuesto)"
                )
                filas.append(ref)

        # Fuera: material/alcance incompatible (p. ej. columnas metálicas para
        # una barda de block) y cifras de otra escala (más de 5× o menos de 1/5
        # del cotizado: casi siempre otra unidad o un total).
        filas = [
            r for r in filas
            if not alcance_distinto(concepto, r["concepto"])
            and precio / FACTOR_ESCALA <= r["precio"] <= precio * FACTOR_ESCALA
        ]
        for ref in filas:
            ref["partida"] = p.get("partida")
            ref["concepto_cotizado"] = concepto
            ref["precio_cotizado"] = precio
            ref["posicion"] = _posicion(precio, ref)
            referencias.append(ref)

        # Los escenarios por pieza (castillo por ml × altura supuesta) se
        # muestran, pero NO entran al rango de mercado hasta confirmar
        # dimensiones con el proveedor.
        mismos = [r for r in filas if r["relacion"] == "mismo material"]
        base = mismos or [r for r in filas if r["relacion"] != "escenario por pieza"]
        if base:
            # Referencias con especificaciones distintas NO se mezclan en un
            # solo rango: se toma la MEJOR coincidencia de cada fuente y se
            # compara por separado. En NL, p25 y p75 son percentiles de
            # precios contratados (no mínimo ni máximo).
            nl = next((r for r in base if r["fuente"] == "Nuevo León"), None)
            cdmx = next((r for r in base if r["fuente"] == "CDMX"), None)
            partes = []
            color = "🟡"
            if nl and nl.get("rango_bajo") and nl.get("rango_alto"):
                if precio > nl["rango_alto"]:
                    partes.append(f"por encima del p75 de NL ({(precio / nl['rango_alto'] - 1) * 100:+.0f}%)")
                    color = "🔴"
                elif precio < nl["rango_bajo"]:
                    partes.append(f"por debajo del p25 de NL ({(precio / nl['rango_bajo'] - 1) * 100:+.0f}%)")
                    color = "🟢" if color != "🔴" else color
                else:
                    partes.append("entre p25 y p75 de NL")
            elif nl:
                partes.append(f"{(precio / nl['precio'] - 1) * 100:+.0f}% vs NL")
            if cdmx:
                dif = (precio / cdmx["precio"] - 1) * 100
                partes.append(f"{dif:+.0f}% vs CDMX")
            pos = f"{color} " + "; ".join(partes) if partes else "⚪ sin referencia comparable"
            resumen.append({
                "partida": p.get("partida"),
                "concepto": concepto,
                "unidad": unidad,
                "precio_cotizado": precio,
                "referencias": len(base),
                "tipo": "mismo material" if mismos else "solo conceptos relacionados",
                # Mejor coincidencia de NL (percentiles de precios contratados)
                "nl_concepto": nl["concepto"] if nl else None,
                "nl_p25": nl.get("rango_bajo") if nl else None,
                "nl_mediana": nl["precio"] if nl else None,
                "nl_p75": nl.get("rango_alto") if nl else None,
                "nl_registros": nl.get("n_registros") if nl else None,
                "nl_periodo": nl.get("fecha") if nl else None,
                # Mejor coincidencia de CDMX (precio único del tabulador)
                "cdmx_concepto": cdmx["concepto"] if cdmx else None,
                "cdmx_precio": cdmx["precio"] if cdmx else None,
                # Compatibilidad con pantallas y exportaciones existentes
                "minimo_etiqueta": "p25 NL" if nl and nl.get("rango_bajo") else None,
                "maximo_etiqueta": "p75 NL" if nl and nl.get("rango_alto") else None,
                "minimo": nl.get("rango_bajo") if nl else None,
                "mediana_referencias": nl["precio"] if nl else (cdmx["precio"] if cdmx else None),
                "maximo": nl.get("rango_alto") if nl else None,
                "vs_mediana_pct": (
                    round((precio / (nl["precio"] if nl else cdmx["precio"]) - 1) * 100, 1)
                ),
                "posicion": pos,
            })
        else:
            escenarios = [r for r in filas if r["relacion"] == "escenario por pieza"]
            resumen.append({
                "partida": p.get("partida"), "concepto": concepto, "unidad": unidad,
                "precio_cotizado": precio, "referencias": 0,
                "nl_concepto": None, "nl_p25": None, "nl_mediana": None, "nl_p75": None,
                "nl_registros": None, "nl_periodo": None, "cdmx_concepto": None, "cdmx_precio": None,
                "minimo_etiqueta": None, "maximo_etiqueta": None,
                "tipo": ("sin referencias por pieza (hay escenario por ml × altura supuesta)"
                         if escenarios else "sin referencias"),
                "minimo": None, "mediana_referencias": None, "maximo": None,
                "vs_mediana_pct": None,
                "posicion": ("⚪ sin referencia por pieza: confirmar dimensiones (ver escenario)"
                             if escenarios else "⚪ sin datos de mercado en NL/CDMX"),
            })
    return {"referencias": referencias, "resumen": resumen}
