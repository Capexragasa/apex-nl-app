"""Dónde enfocarse para negociar.

Ordena las partidas de una cotización por el dinero que está en juego y por
qué tan firme es la evidencia de los 4 filtros. Todo sale de datos que ya
calculó la app (no promedia fuentes ni inventa precios):

- Precio: partidas donde uno o más filtros dicen «caro». El monto en juego
  es (precio cotizado − precio de referencia) × cantidad, por separado para
  cada filtro (se muestra el rango, no un promedio).
- Cantidad: partidas con hallazgo de cantidades (p. ej. m² de acabado que no
  cuadran con el muro); el monto es el del escenario, sujeto a aclaración.
- Especificación: partidas sin referencia comparable, donde primero hay que
  pedir medidas/alcance para poder comparar.

La redacción con IA (Gemini) es opcional: solo reescribe estos mismos hechos
y se descarta si menciona un importe que no está en los datos.
"""
from __future__ import annotations

import re

import validacion_referencias as v

NOMBRES = {"historico": "Histórico Ragasa", "nl": "Nuevo León", "cdmx": "CDMX", "ia": "IA internet"}
# Si la referencia de NL queda más de este porcentaje por debajo de la de
# CDMX para la misma partida, se avisa (se esperaría lo contrario).
UMBRAL_COHERENCIA = 0.10
# Para decidir dónde enfocarse (no para el color): una diferencia menor a
# este porcentaje se considera "prácticamente igual" y no se manda a negociar.
MINIMO_MATERIAL_PCT = 1.0


def _num(x):
    try:
        return float(str(x).replace(",", "").replace("$", ""))
    except (TypeError, ValueError):
        return 0.0


def nombre_corto(concepto, n=3) -> str:
    """Primeras palabras del concepto, sin terminar en preposición."""
    sueltas = ("de", "del", "la", "el", "en", "para", "y", "con", "a", "al", "los", "las", "por")
    palabras = re.sub(r"[^\wÁÉÍÓÚÑáéíóúñ ]", " ", str(concepto)).split()[:n]
    while palabras and (palabras[-1].lower() in sueltas or palabras[-1].isdigit()):
        palabras.pop()
    texto = " ".join(palabras).lower()
    return texto[:1].upper() + texto[1:] if texto else "Partida"


def coherencia_nl_cdmx(evaluaciones: dict):
    """Aviso cuando la referencia de Nuevo León queda por debajo de la de CDMX
    (se esperaría que NL fuera igual o más caro). None si no aplica."""
    nl, cdmx = evaluaciones.get("nl") or {}, evaluaciones.get("cdmx") or {}
    if not (v.cuenta_filtro(nl) and v.cuenta_filtro(cdmx)):
        return None
    p_nl, p_cdmx = nl.get("precio_referencia"), cdmx.get("precio_referencia")
    if not p_nl or not p_cdmx:
        return None
    brecha = p_nl / p_cdmx - 1
    if brecha >= -UMBRAL_COHERENCIA:
        return None
    acumulada = (nl.get("inflacion") or {}).get("acumulada_pct")
    return {
        "brecha_pct": round(brecha * 100, 1),
        "texto": (
            f"Nuevo León (${p_nl:,.2f}) queda {abs(brecha) * 100:.0f} % por debajo de CDMX (${p_cdmx:,.2f})"
            + (f", ya con inflación acumulada de +{acumulada:.1f} %" if acumulada is not None else "")
            + ". Se esperaría lo contrario. Causas probables: (1) son contratos de obra pública de gran volumen "
              "(precio unitario más bajo que en un trabajo chico); (2) la especificación del contrato puede ser "
              "distinta (referencia orientativa); (3) el índice general no sigue por fuerza el costo de la "
              "construcción. Úsala como PISO de negociación, no como precio objetivo."),
    }


def prioridades(filas: list, revision_cant: dict | None = None) -> dict:
    """Lista ordenada de partidas en las que conviene enfocarse."""
    total = sum(_num(f.get("Cantidad")) * _num(f.get("Precio cotizado")) for f in filas) or 0.0
    hallazgos = {}
    for h in (revision_cant or {}).get("hallazgos", []):
        if h.get("nivel") == "REVISAR" and h.get("escenarios"):
            montos = [abs(e.get("importe") or 0) for e in h["escenarios"] if e.get("importe")]
            if montos:
                hallazgos[str(h.get("partidas")).strip()] = {"min": min(montos), "max": max(montos),
                                                             "titulo": h.get("tipo") or "cantidad por aclarar"}
    items = []
    for i, f in enumerate(filas):
        evs = f.get("_evaluaciones") or {}
        cantidad, precio = _num(f.get("Cantidad")), _num(f.get("Precio cotizado"))
        importe = cantidad * precio
        usados = [(k, e) for k, e in evs.items() if v.cuenta_filtro(e)]
        caros = [(k, e) for k, e in usados if e.get("clasificacion") == v.ALTO
                 and (e.get("diferencia_pct") or 0) >= MINIMO_MATERIAL_PCT]
        otros = [(k, e) for k, e in usados if (k, e) not in caros]
        validado = any(e.get("estado") == v.VALIDADA for _, e in caros)
        montos = [max(0.0, (precio - e["precio_referencia"]) * cantidad) for _, e in caros]
        coher = coherencia_nl_cdmx(evs)
        cant = hallazgos.get(str(f.get("Partida", i + 1)).strip())
        item = {
            "partida": f.get("Partida", i + 1), "concepto": str(f.get("Concepto") or ""),
            "nombre": nombre_corto(f.get("Concepto")), "unidad": f.get("Unidad"),
            "cantidad": cantidad, "precio": precio, "importe": importe,
            "peso_pct": round(importe / total * 100, 1) if total else 0.0,
            "caros": [{"fuente": NOMBRES.get(k, k), "clave": k, "precio_ref": e["precio_referencia"],
                       "pct": e.get("diferencia_pct"), "validada": e.get("estado") == v.VALIDADA,
                       "monto": max(0.0, (precio - e["precio_referencia"]) * cantidad)} for k, e in caros],
            "otros": [{"fuente": NOMBRES.get(k, k), "clave": k, "precio_ref": e["precio_referencia"],
                       "pct": e.get("diferencia_pct"), "clasificacion": e.get("clasificacion"),
                       "casi_igual": e.get("clasificacion") == v.ALTO} for k, e in otros],
            "n_filtros": len(usados), "validado": validado,
            "monto_min": min(montos) if montos else 0.0, "monto_max": max(montos) if montos else 0.0,
            "cantidad_por_aclarar": cant, "coherencia": coher,
        }
        if caros and (validado or len(caros) >= 2 or not otros):
            item["tipo"], item["nivel"] = "precio", ("Negociar precio" if validado else "Negociar precio (pide desglose)")
        elif caros:
            item["tipo"], item["nivel"] = "precio_dudoso", "Revisar precio: los filtros no coinciden"
        elif not usados:
            item["tipo"], item["nivel"] = "especificacion", "Pedir especificación para poder comparar"
        else:
            item["tipo"], item["nivel"] = "ok", "Sin acción de precio"
        item["en_juego"] = max(item["monto_max"] if item["tipo"] in ("precio", "precio_dudoso") else 0.0,
                               (cant or {}).get("max", 0.0))
        item["texto"] = _texto(item)
        items.append(item)
    orden_tipo = {"precio": 0, "precio_dudoso": 1, "especificacion": 2, "ok": 3}
    items.sort(key=lambda it: (0 if it["en_juego"] > 0 else 1, -it["en_juego"], orden_tipo[it["tipo"]], -it["importe"]))
    foco = [it for it in items if it["en_juego"] > 0 or it["tipo"] == "especificacion"]
    peso_foco = sum(it["peso_pct"] for it in foco if it["en_juego"] > 0)
    return {
        "items": items, "foco": foco, "total": total,
        "resumen": (
            f"Enfócate en {sum(1 for it in foco if it['en_juego'] > 0)} partida(s) que concentran "
            f"{peso_foco:.0f} % del importe (${sum(it['importe'] for it in foco if it['en_juego'] > 0):,.0f} de "
            f"${total:,.0f})." if any(it["en_juego"] > 0 for it in foco) else
            "Ningún filtro marca un precio alto con dinero en juego; lo pendiente es completar especificaciones."),
    }


def _texto(it: dict) -> str:
    """Una línea por partida, en lenguaje simple y con sus números."""
    partes = []
    cant = it.get("cantidad_por_aclarar")
    if cant:
        rango = (f"${cant['min']:,.0f} a ${cant['max']:,.0f}" if abs(cant["max"] - cant["min"]) > 1
                 else f"${cant['max']:,.0f}")
        partes.append(f"primero aclara la cantidad ({it['cantidad']:g} {str(it['unidad']).lower()}): "
                      f"{rango} sujetos a aclaración; pide el generador de volúmenes")
    if it["tipo"] in ("precio", "precio_dudoso"):
        caros = " y ".join(
            f"{c['fuente']} (${c['precio_ref']:,.2f}, {c['pct']:+.1f} %{'' if c['validada'] else ', orientativa'})"
            for c in it["caros"])
        monto = (f"${it['monto_min']:,.0f} a ${it['monto_max']:,.0f}" if abs(it["monto_max"] - it["monto_min"]) > 1
                 else f"${it['monto_max']:,.0f}")
        frase = (f"{'alto' if it['validado'] else 'posible alto'} frente a {caros}; en juego "
                 f"{'' if it['validado'] or ' a ' in monto else 'hasta '}{monto}")
        if it["otros"]:
            frase += "; en cambio " + " y ".join(
                (f"queda prácticamente igual a {o['fuente']} (${o['precio_ref']:,.2f})" if o.get("casi_igual") else
                 f"{o['fuente']} lo da {v._CLAS_TEXTO.get(o['clasificacion'], '')} (${o['precio_ref']:,.2f})")
                for o in it["otros"])
        partes.append(frase)
        if it.get("coherencia"):
            cdmx = next((o for o in it["otros"] + it["caros"] if o["clave"] == "cdmx"), None)
            partes.append("ojo: Nuevo León queda por debajo de CDMX, tómalo como piso"
                          + (f" y usa CDMX (${cdmx['precio_ref']:,.2f}) como objetivo realista" if cdmx else ""))
        elif it["tipo"] == "precio":
            partes.append("pide el desglose del precio unitario (material, mano de obra, indirectos)"
                          if not it["validado"] else "negocia hacia la referencia validada")
        else:
            partes.append("pide la especificación para saber cuál referencia aplica")
    elif it["tipo"] == "especificacion":
        partes.append("ningún filtro tiene un concepto comparable: pide medidas y alcance para poder compararla")
    elif not cant:
        partes.append("los filtros la dan baja o prácticamente igual a la referencia")
    return "; ".join(partes) + "."


def hechos_para_ia(reco: dict, maximo=6) -> str:
    """Los mismos hechos, en texto plano, para que la IA los redacte."""
    lineas = [reco["resumen"]]
    for it in reco["items"][:maximo]:
        lineas.append(f"- Partida {it['partida']} {it['nombre']} (importe ${it['importe']:,.0f}, "
                      f"{it['peso_pct']:.0f} % del total, P.U. ${it['precio']:,.2f}): {it['texto']}")
    return "\n".join(lineas)


def _cifras(texto: str) -> set:
    """Importes en pesos que aparecen en un texto (sin centavos)."""
    return {str(int(round(_num(x)))) for x in re.findall(r"\$\s?([\d,]+(?:\.\d+)?)", str(texto))}


def _solo_texto(respuesta) -> str:
    """Deja solo el párrafo: algunos modelos lo devuelven envuelto en JSON
    ({"recomendacion": "..."}) o en un bloque de código."""
    texto = str(respuesta).strip()
    texto = re.sub(r"^```[a-zA-Z]*\s*|\s*```$", "", texto).strip()
    if texto.startswith("{"):
        try:
            import json
            datos = json.loads(texto)
            valores = [x for x in datos.values() if isinstance(x, str)] if isinstance(datos, dict) else []
            if valores:
                texto = max(valores, key=len)
        except Exception:
            m = re.search(r':\s*"(.+)"\s*}\s*$', texto, re.S)
            if m:
                texto = m.group(1)
    return re.sub(r"\s+", " ", texto).strip().strip('"').strip()


def redaccion_ia(reco: dict, llamar) -> str | None:
    """Pide a la IA un párrafo con la recomendación. `llamar(prompt)` regresa
    el texto o None. Se descarta si trae un importe que no está en los hechos
    (la IA no puede agregar precios)."""
    hechos = hechos_para_ia(reco)
    prompt = (
        "Eres analista de compras CAPEX en Monterrey. Con ÚNICAMENTE los hechos de abajo, escribe en español "
        "simple una recomendación de máximo 90 palabras: en qué partidas enfocarse al negociar, en qué orden y "
        "qué pedirle al proveedor. No inventes precios, porcentajes ni datos que no estén en los hechos; si una "
        "referencia es orientativa dilo como 'posible alto'. Responde solo con el párrafo, en texto simple: sin "
        "JSON, sin llaves, sin comillas, sin listas ni encabezados.\n\nHECHOS:\n" + hechos)
    try:
        texto = llamar(prompt)
    except Exception:
        return None
    if not texto:
        return None
    texto = _solo_texto(texto)
    if len(texto) < 40 or not _cifras(texto) <= _cifras(hechos):
        return None
    return texto


_CACHE_IA = {}   # hechos -> (momento, texto o None)


def redaccion_ia_con_cache(reco: dict, llamar, reintentar_s=600):
    """Igual que redaccion_ia, pero sin repetir la llamada en cada recarga de
    la pantalla. Un fallo se reintenta después de `reintentar_s` segundos."""
    import time
    clave = hechos_para_ia(reco)
    previo = _CACHE_IA.get(clave)
    if previo and (previo[1] or time.time() - previo[0] < reintentar_s):
        return previo[1]
    texto = redaccion_ia(reco, llamar)
    _CACHE_IA[clave] = (time.time(), texto)
    if len(_CACHE_IA) > 200:
        _CACHE_IA.pop(next(iter(_CACHE_IA)))
    return texto
