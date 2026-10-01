"""
Revisión de cantidades y aritmética de una cotización
=====================================================

Complementa la comparación de precios: un precio unitario correcto no sirve
de nada si la cantidad cotizada está inflada. Caso real (barda en azotea):
16 m² de muro a 0.60 m de altura, pero 59 m² de estuco "por fuera y por
dentro" -- dos caras del muro dan 32 m²; aun sumando el cerramiento y la
corona no se llega a 59 m².

Qué revisa:
  1. Aritmética: cantidad × P.U. contra el importe de cada partida y la suma
     contra el total declarado.
  2. Geometría de bardas y muros: longitud implícita (m² ÷ altura) contra
     los metros lineales de cerramiento/dala, y m² de acabados "a dos caras"
     contra el área del muro, en tres escenarios.
  3. Castillos/columnas por pieza: separación implícita (solo como pregunta,
     nunca como conclusión).
  4. Alcances e impuestos a confirmar por escrito.

Nada de esto se da por hecho: son diferencias SUJETAS A ACLARACIÓN con el
generador de volúmenes del proveedor, no ahorros confirmados.
"""

from __future__ import annotations

import re
import unicodedata

TOLERANCIA_ARITMETICA = 1.0          # pesos
ALTURA_CERRAMIENTO_SUPUESTA = 0.20   # m, solo para el escenario E2
ANCHO_CORONA_SUPUESTO = 0.15         # m, solo para el escenario E3
IVA = 0.16


def _plano(texto) -> str:
    t = unicodedata.normalize("NFKD", str(texto or "").upper())
    return "".join(c for c in t if not unicodedata.combining(c))


def _num(valor):
    try:
        if valor is None or valor == "":
            return None
        n = float(str(valor).replace("$", "").replace(",", ""))
        return None if n != n else n
    except (TypeError, ValueError):
        return None


def _unidad(u) -> str:
    u = _plano(u).replace(".", "").strip()
    if u in ("M2", "MT2", "M²"):
        return "M2"
    if u in ("ML", "M", "MTS", "MT", "METRO", "METROS"):
        return "ML"
    if u in ("PZA", "PZ", "PZAS", "PIEZA", "PIEZAS", "UNIDAD", "PZ."):
        return "PZA"
    return u


def _altura_m(texto: str):
    t = _plano(texto)
    m = re.search(r"(\d+(?:[.,]\d+)?)\s*(CM|CMS|M|MT|MTS)\.?\s*(?:DE\s+)?ALTURA", t) or re.search(
        r"ALTURA\s*(?:DE|:)?\s*(\d+(?:[.,]\d+)?)\s*(CM|CMS|M|MT|MTS)", t
    )
    if not m:
        return None
    valor = float(m.group(1).replace(",", "."))
    return valor / 100 if m.group(2).startswith("CM") else valor


_MURO = re.compile(r"\b(MURO|BARDA|PRETIL)\b")
_MAMPOSTERIA = re.compile(r"\b(BLOCK|BLOQUE|TABIQUE|LADRILLO|TABICON)\b")
_LINEAL = re.compile(r"\b(CERRAMIENTO|DALA|CADENA|TRABE)\b")
_ACABADO = re.compile(r"\b(ESTUCO|APLANADO|ZARPEO|REPELLADO|ENJARRE|AFINE|PINTURA)\b")
_DOS_CARAS = re.compile(
    r"(AMBAS CARAS|AMBOS LADOS|DOS CARAS|2 CARAS|POR FUERA Y POR DENTRO|"
    r"TANTO POR FUERA COMO POR DENTRO|INTERIOR Y EXTERIOR|EXTERIOR E INTERIOR)"
)
_CASTILLO = re.compile(r"\b(CASTILLO|CASTILLOS|COLUMNA|COLUMNAS)\b")


def revisar(partidas: list[dict], total_declarado=None, contexto: str = "", iva_en_documento=None) -> dict:
    """partidas: dicts con partida, concepto, unidad, cantidad, precio_unitario, importe."""
    aritmetica = []
    total_calculado = 0.0
    suma_importes = 0.0
    for p in partidas:
        cant, pu, imp = _num(p.get("cantidad")), _num(p.get("precio_unitario")), _num(p.get("importe"))
        calc = round(cant * pu, 2) if cant is not None and pu is not None else None
        if calc is not None:
            total_calculado += calc
        if imp is not None:
            suma_importes += imp
        dif = round(imp - calc, 2) if imp is not None and calc is not None else None
        aritmetica.append({
            "partida": p.get("partida"), "concepto": p.get("concepto"), "unidad": p.get("unidad"),
            "cantidad": cant, "precio_unitario": pu, "importe_declarado": imp,
            "importe_calculado": calc, "diferencia": dif,
            "ok": dif is None or abs(dif) <= TOLERANCIA_ARITMETICA,
        })
    total_ref = _num(total_declarado) or (suma_importes or None)
    total_ok = total_ref is None or abs(total_ref - total_calculado) <= TOLERANCIA_ARITMETICA

    hallazgos = []

    muros = [p for p in partidas if _unidad(p.get("unidad")) == "M2"
             and _MURO.search(_plano(p.get("concepto"))) and _MAMPOSTERIA.search(_plano(p.get("concepto")))]
    lineales = [p for p in partidas if _unidad(p.get("unidad")) == "ML"
                and _LINEAL.search(_plano(p.get("concepto")))]
    acabados = [p for p in partidas if _unidad(p.get("unidad")) == "M2"
                and _ACABADO.search(_plano(p.get("concepto")))]
    castillos = [p for p in partidas if _unidad(p.get("unidad")) == "PZA"
                 and _CASTILLO.search(_plano(p.get("concepto")))]

    area_muro = sum(_num(p.get("cantidad")) or 0 for p in muros) or None
    altura = next((h for h in (_altura_m(p.get("concepto")) for p in muros) if h), None)
    longitud_lineal = max((_num(p.get("cantidad")) or 0 for p in lineales), default=0) or None
    longitud_implicita = round(area_muro / altura, 2) if area_muro and altura else None
    longitud = longitud_lineal or longitud_implicita

    if longitud_implicita and longitud_lineal:
        dif = longitud_lineal - longitud_implicita
        coherente = abs(dif) <= max(1.0, 0.10 * longitud_implicita)
        hallazgos.append({
            "tipo": "Coherencia muro – cerramiento",
            "nivel": "OK" if coherente else "REVISAR",
            "partidas": ", ".join(str(p.get("partida")) for p in muros + lineales),
            "detalle": (
                f"{area_muro:g} m² ÷ {altura:g} m de altura = {longitud_implicita:g} ml implícitos de muro, "
                f"contra {longitud_lineal:g} ml de {_LINEAL.search(_plano(lineales[0].get('concepto'))).group(1).lower()}. "
                + ("Diferencia menor, razonable por redondeo." if coherente else
                   f"Diferencia de {dif:+.2f} ml: pedir generador de volúmenes.")
            ),
            "escenarios": [],
        })

    for p in acabados:
        texto = _plano(p.get("concepto"))
        cant, pu = _num(p.get("cantidad")), _num(p.get("precio_unitario")) or 0
        if not (area_muro and cant and _DOS_CARAS.search(texto)):
            continue
        e1 = round(2 * area_muro, 2)
        escenarios = [{"nombre": "E1 – solo las 2 caras del muro", "formula": f"{area_muro:g} × 2", "m2": e1,
                       "base": "medida de la cotización"}]
        # Variante con la longitud del cerramiento (27 ml) y la altura del
        # muro declarada (0.60 m): también usa solo medidas de la cotización.
        if longitud_lineal and altura and abs(longitud_lineal * altura - area_muro) > 0.01:
            escenarios.append({
                "nombre": "E1b – 2 caras con la longitud del cerramiento",
                "formula": f"{longitud_lineal:g} × {altura:g} × 2",
                "m2": round(longitud_lineal * altura * 2, 2),
                "base": "medida de la cotización",
            })
        if longitud:
            e2 = round(e1 + 2 * longitud * ALTURA_CERRAMIENTO_SUPUESTA, 2)
            e3 = round(e2 + longitud * ANCHO_CORONA_SUPUESTO, 2)
            escenarios += [
                {"nombre": f"E2 – E1 + 2 caras del cerramiento (supuesto h={ALTURA_CERRAMIENTO_SUPUESTA:g} m)",
                 "formula": f"{e1:g} + {longitud:g}×{ALTURA_CERRAMIENTO_SUPUESTA:g}×2", "m2": e2,
                 "base": "incluye medidas supuestas"},
                {"nombre": f"E3 – E2 + corona (supuesto {ANCHO_CORONA_SUPUESTO:g} m)",
                 "formula": f"{e2:g} + {longitud:g}×{ANCHO_CORONA_SUPUESTO:g}", "m2": e3,
                 "base": "incluye medidas supuestas"},
            ]
        # Diferencia = cantidad cotizada − cantidad del escenario; importe =
        # esa diferencia × P.U. cotizado. Positivo = el proveedor cobra más m²
        # de los que da el escenario. Es un importe SUJETO A ACLARACIÓN, no
        # un ahorro.
        for e in escenarios:
            e["dif_m2"] = round(cant - e["m2"], 2)
            e["importe"] = round(e["dif_m2"] * pu, 2)
        maximo = max(e["m2"] for e in escenarios)
        excede = cant > maximo * 1.05
        # Los escenarios se presentan por separado, no como intervalo: E1 usa
        # solo medidas de la cotización; E2 y E3 agregan superficies con
        # medidas SUPUESTAS, así que no fijan un límite de la aclaración.
        lista = "; ".join(
            f"{e['nombre'].split(' –')[0]} ({e['base']}): {e['m2']:g} m², "
            f"{e['dif_m2']:g} m² de diferencia, ${e['importe']:,.2f} sujetos a aclaración"
            for e in escenarios
        )
        hallazgos.append({
            "tipo": "Cantidad de acabado a dos caras",
            "nivel": "REVISAR" if excede else "OK",
            "partidas": str(p.get("partida")),
            "detalle": (
                f"Se cotizan {cant:g} m² de {_ACABADO.search(texto).group(1).lower()} a dos caras sobre "
                f"{area_muro:g} m² de muro. "
                + (f"Ningún escenario llega a {cant:g} m². Escenarios por separado — {lista}. "
                   "No son ahorros. E2 y E3 dependen de medidas supuestas; pueden existir superficies no descritas "
                   "(pretil, remates): pedir generador." if excede else "La cantidad es coherente con el área del muro.")
            ),
            "escenarios": escenarios,
        })

    if castillos and longitud:
        n = sum(_num(p.get("cantidad")) or 0 for p in castillos)
        if n:
            hallazgos.append({
                "tipo": "Castillos / columnas",
                "nivel": "CONFIRMAR",
                "partidas": ", ".join(str(p.get("partida")) for p in castillos),
                "detalle": (
                    f"{n:g} piezas en {longitud:g} ml da una separación promedio de ~{longitud / n:.1f} m. "
                    "No se asume la separación real: pedir croquis acotado. Esta revisión no certifica "
                    "la seguridad estructural."
                ),
                "escenarios": [],
            })

    total_para_iva = total_ref or total_calculado
    alcances = [
        (f"IVA: la cotización lo desglosa; se compara el subtotal sin IVA (${total_para_iva:,.2f})."
         if iva_en_documento else
         f"IVA: la cotización no indica si sus precios incluyen IVA. Se toman como precios sin IVA "
         f"(criterio de Compras); si lo incluyeran, el subtotal sería ${total_para_iva / (1 + IVA):,.2f} "
         f"y cada P.U. bajaría 13.8 %. Confirmar con el proveedor."),
        "Vigencia de la cotización, condiciones de pago y tiempo de entrega.",
        "Retiro de escombro, limpieza final y curado de concreto.",
    ]
    texto_total = _plano(" ".join(str(p.get("concepto")) for p in partidas) + " " + contexto)
    if "AZOTEA" in texto_total or "LOSA" in texto_total or "PLACA" in texto_total:
        alcances += [
            "Acarreo de material a la azotea (elevación) y andamios.",
            "Protección de la losa/placa existente durante la obra.",
            "Anclaje o conector del muro a la losa (varilla, epóxico).",
        ]
    if muros:
        alcances.append("Generador de volúmenes y croquis acotado de la barda.")

    return {
        "aritmetica": aritmetica,
        "total_calculado": round(total_calculado, 2),
        "total_declarado": total_ref,
        "total_ok": total_ok,
        "partidas_con_error": sum(1 for a in aritmetica if not a["ok"]),
        "hallazgos": hallazgos,
        "alcances_confirmar": alcances,
        "altura_muro": altura,
        "longitud": longitud,
    }
