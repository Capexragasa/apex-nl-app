"""
Búsqueda web de respaldo (cuando Gemini no responde o se quedó sin cuota)
=========================================================================

Investiga CADA partida por separado y verifica el CONTENIDO de las páginas,
no solo el fragmento que devuelve el buscador:

  1. Arma 1-2 consultas por material + trabajo + unidad (p. ej. "columnas
     para amarrar barda" se busca como "castillo de concreto"; "cerramiento"
     como "dala de concreto con armex").
  2. Busca con Tavily (si hay 'tavily_api_key') o, sin clave, con el
     buscador público (paquete ddgs).
  3. Abre cada página (HTML o PDF) y localiza el renglón donde aparecen
     juntos el concepto, la unidad y el precio.
  4. Rechaza materiales, elementos o trabajos distintos (validacion_
     referencias.alcance_distinto) y cifras fuera de escala.

El resultado SIEMPRE es orientativo (verificado=False): se encontró un
precio real en una página real, pero la equivalencia exacta (especificación,
alcance, fecha, IVA) no está demostrada. La app nunca lo presenta como
validado.
"""

from __future__ import annotations

import re
import time
import unicodedata
from concurrent.futures import ThreadPoolExecutor, as_completed

TIEMPO_MAX_PARTIDA = 30      # s por partida (búsqueda + lectura de páginas)
PAGINAS_POR_CONSULTA = 5
FACTOR_ESCALA_WEB = 3.0
UMBRAL_TEXTO = 55            # similitud mínima concepto-renglón (0-100)

_EXCLUIR_URL = (
    "scribd.com", "pinterest.", "facebook.", "reddit.", "tiktok.", "youtube.",
    "youtu.be", "instagram.", "twitter.", "x.com/", "linkedin.", "wikipedia.",
)
_CABECERAS = {
    "User-Agent": ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                   "(KHTML, like Gecko) Chrome/126.0 Safari/537.36"),
    "Accept-Language": "es-MX,es;q=0.9",
}

_UNIDAD_TEXTO = {
    "M2": "m2", "M3": "m3", "ML": "metro lineal", "PZA": "pieza", "KG": "kg",
    "TON": "tonelada", "LT": "litro", "JORNAL": "jornal",
}
_UNIDAD_PATRON = {
    "M2": r"(m2|m²|mt2|metros? cuadrados?|/m2|x m2)",
    "M3": r"(m3|m³|mt3|metros? cubicos?)",
    "ML": r"(\bml\b|m\.l\.|metros? lineal(es)?|/ml|\bmts?\b\.?(?! ?[23²³])|por metro\b(?! cuadrado| cubico))",
    "PZA": r"(\bpzas?\b|\bpz\b|piezas?|\bc/u\b|cada un[oa]|\bunidad\b)",
    "KG": r"(\bkgs?\b|kilos?|kilogramos?)",
    "TON": r"(\bton\b|toneladas?)",
    "LT": r"(\blts?\b|litros?)",
    "JORNAL": r"(jornal|por dia|jornada)",
}
_PRECIO = re.compile(
    r"\$\s?(\d{1,3}(?:,\d{3})+(?:\.\d{1,2})?|\d+(?:\.\d{1,2})?)"
    r"|(\d{1,3}(?:,\d{3})+(?:\.\d{1,2})?|\d+(?:\.\d{1,2})?)\s?(?:mxn|pesos|m\.n\.)",
    re.I,
)
# En tablas de tabuladores el precio va sin "$", justo después de la unidad:
# "M2 219.12". Solo se acepta con dos decimales.
_PRECIO_TABLA = re.compile(r"\b(m2|m3|ml|m|pza|pieza|kg)\s+(\d{1,3}(?:,\d{3})*\.\d{2})\b", re.I)
_MONEDA_EXTRANJERA = re.compile(
    r"(₹|€|£|us\$|\busd\b|dolar|dollar|\beuros?\b|\bindia\b|\bcop\b|\bsoles\b|\bpies\b|\bfeet\b|sq ?ft)"
)
_ANIO = re.compile(r"\b(202[3-7])\b")


def _plano(texto) -> str:
    t = unicodedata.normalize("NFKD", str(texto or "").lower())
    return "".join(c for c in t if not unicodedata.combining(c))


def _unidad_canonica(unidad) -> str:
    try:
        from comparador_multifuente_v2 import normalize_unit
        return normalize_unit(unidad)
    except Exception:
        return str(unidad or "").strip().upper()


def consultas_para(descripcion: str, unidad: str) -> list[str]:
    """1-2 consultas de búsqueda por material + trabajo + unidad."""
    try:
        from comparador_multifuente_v2 import consultas_catalogo, normalize_text
        opciones = consultas_catalogo(normalize_text(descripcion))
    except Exception:
        opciones = [descripcion]
    # La primera es el texto original (largo y con relleno de obra): se
    # prefieren las versiones de catálogo y las equivalencias de trabajo.
    opciones = [o for o in opciones[1:] if len(o.split()) >= 2] or opciones[:1]
    u = _UNIDAD_TEXTO.get(_unidad_canonica(unidad), str(unidad or "").lower())
    vistas, consultas = set(), []
    for o in opciones:
        clave = " ".join(sorted(set(o.lower().split()))[:4])
        if clave in vistas:
            continue
        vistas.add(clave)
        consultas.append(f"precio unitario {o.lower()} por {u} México")
        if len(consultas) == 2:
            break
    return consultas


# ----------------------------------------------------------------------
# Buscadores
# ----------------------------------------------------------------------
def buscar_tavily(consulta: str, api_key: str) -> list[dict]:
    import requests
    r = requests.post(
        "https://api.tavily.com/search",
        json={"api_key": api_key, "query": consulta, "search_depth": "basic",
              "include_answer": False, "include_raw_content": True,
              "max_results": PAGINAS_POR_CONSULTA},
        timeout=20,
    )
    r.raise_for_status()
    return [
        {"url": x.get("url"), "titulo": x.get("title") or "", "resumen": x.get("content") or "",
         "texto": x.get("raw_content") or ""}
        for x in r.json().get("results") or []
    ]


def buscar_publico(consulta: str) -> list[dict]:
    from ddgs import DDGS
    resultados = DDGS(timeout=10).text(consulta, region="mx-es", max_results=PAGINAS_POR_CONSULTA + 3)
    return [
        {"url": x.get("href") or x.get("url"), "titulo": x.get("title") or "",
         "resumen": x.get("body") or "", "texto": ""}
        for x in resultados or []
    ]


def leer_pagina(url: str) -> str:
    """Texto de la página (HTML o PDF); '' si no se puede abrir."""
    import requests
    try:
        r = requests.get(url, headers=_CABECERAS, timeout=(5, 10), stream=True, allow_redirects=True)
        if r.status_code >= 400:
            return ""
        tipo = (r.headers.get("content-type") or "").lower()
        contenido = b""
        for bloque in r.iter_content(65536):
            contenido += bloque
            if len(contenido) > 4_000_000:
                break
    except Exception:
        return ""
    if "pdf" in tipo or url.lower().endswith(".pdf"):
        try:
            import io
            import pdfplumber
            with pdfplumber.open(io.BytesIO(contenido)) as pdf:
                return "\n".join((p.extract_text() or "") for p in pdf.pages[:25])
        except Exception:
            return ""
    try:
        from bs4 import BeautifulSoup
        sopa = BeautifulSoup(contenido, "html.parser")
        for etiqueta in sopa(["script", "style", "noscript", "nav", "footer", "header", "form", "svg"]):
            etiqueta.decompose()
        # Celdas de tabla en la misma línea: "concepto | unidad | precio".
        for fila in sopa.find_all("tr"):
            celdas = [c.get_text(" ", strip=True) for c in fila.find_all(["td", "th"])]
            fila.replace_with(" | ".join(c for c in celdas if c) + "\n")
        return sopa.get_text("\n")
    except Exception:
        return re.sub(r"<[^>]+>", " ", contenido.decode("utf-8", "ignore"))


# ----------------------------------------------------------------------
# Lectura del contenido: renglón con concepto + unidad + precio
# ----------------------------------------------------------------------
def _precios_en(linea: str) -> list[float]:
    valores = []
    for m in _PRECIO.finditer(linea):
        try:
            valores.append(float((m.group(1) or m.group(2)).replace(",", "")))
        except (TypeError, ValueError):
            pass
    for m in _PRECIO_TABLA.finditer(linea):
        try:
            valores.append(float(m.group(2).replace(",", "")))
        except ValueError:
            pass
    return valores


def extraer_referencias(texto: str, *, descripcion: str, unidad: str, precio_cotizado=None,
                        url: str = "", titulo: str = "", origen: str = "página") -> list[dict]:
    """Referencias (concepto, unidad, precio) encontradas en el texto."""
    from rapidfuzz import fuzz
    from validacion_referencias import alcance_distinto, elemento_principal, trabajo_principal

    if not texto:
        return []
    u = _unidad_canonica(unidad)
    patron_u = re.compile(_UNIDAD_PATRON.get(u, re.escape(_plano(unidad) or "-")), re.I)
    consultas = [_plano(c) for c in consultas_para(descripcion, unidad)]
    consultas = [re.sub(r"^precio unitario | por \S+( \S+)? mexico$", "", c) for c in consultas]
    elemento = elemento_principal(descripcion)
    trabajo = trabajo_principal(descripcion)

    lineas = [re.sub(r"\s+", " ", l).strip() for l in texto.splitlines()]
    lineas = [l for l in lineas if l]
    salida = []
    for i, linea in enumerate(lineas):
        precios = _precios_en(linea)
        if not precios:
            continue
        # Concepto: el MISMO renglón. Solo si el renglón no describe nada
        # (precio suelto en una ficha de producto) se usan los renglones
        # anteriores, sin cruzar a otro renglón con precio.
        palabras = re.findall(r"[a-záéíóúñ]{4,}", linea.lower())
        contexto = linea
        plano_linea = _plano(linea)
        # Renglones de desglose de un APU: no son el precio del concepto.
        if re.match(r"\W*(materiales|mano de obra|equipo|herramienta|subtotal|utilidad|indirecto|"
                    r"financiamiento|cargos? adicional|costo directo|sub-?total|iva|total)\b", plano_linea):
            continue
        if re.search(r"precio unitario|p\.u\.|importe unitario", plano_linea) and len(palabras) < 8:
            # Análisis de precio unitario (APU) en PDF de licitación: el
            # concepto y la unidad están arriba, en "Descripción ... Unidad: ML".
            inicio = None
            for j in range(i - 1, max(-1, i - 30), -1):
                pj = _plano(lineas[j])
                if re.search(r"precio unitario", pj):
                    break
                if re.search(r"^\W*(descripcion|concepto)\b", pj):
                    inicio = j
                    break
            if inicio is not None:
                desc = [lineas[inicio]] + [l for l in lineas[inicio + 1:inicio + 4] if not _precios_en(l)]
                cercanas = lineas[max(0, inicio - 3):i]
                unidad_l = next((l for l in cercanas if re.search(r"\bunidad\s*:", _plano(l))), "")
                contexto = " ".join(desc + [unidad_l, linea])
        elif len(palabras) < 2:
            previos = []
            for j in range(i - 1, max(-1, i - 3), -1):
                if _precios_en(lineas[j]):
                    break
                previos.insert(0, lineas[j])
            contexto = " ".join(previos + [linea])
        # Párrafos largos (blogs, calculadoras) mezclan varias cifras: no son
        # un renglón de catálogo.
        if len(contexto) > 500:
            continue
        ctx = re.sub(r"(\d)\s*x\s*(?=\d)", r"\1 x ", _plano(contexto))
        if _MONEDA_EXTRANJERA.search(ctx):
            continue
        if not patron_u.search(ctx):
            continue
        if elemento and elemento_principal(contexto) != elemento and not (
                elemento == "castillo" and elemento_principal(contexto) == "columna"):
            continue
        if trabajo and trabajo_principal(contexto) != trabajo:
            continue
        if alcance_distinto(descripcion, contexto):
            continue
        similitud = max((fuzz.token_set_ratio(c, ctx) for c in consultas), default=0)
        if similitud < UMBRAL_TEXTO:
            continue
        for precio in precios[:2]:
            if precio <= 0:
                continue
            if precio_cotizado and not (
                    precio_cotizado / FACTOR_ESCALA_WEB <= precio <= precio_cotizado * FACTOR_ESCALA_WEB):
                continue
            if not precio_cotizado and not (5 <= precio <= 200_000):
                continue
            anio = max(_ANIO.findall(contexto), default=None)
            salida.append({
                "precio": round(precio, 2),
                "concepto": contexto[:300],
                "unidad": u,
                "url": url,
                "titulo": titulo[:140],
                "similitud": round(similitud, 1),
                "unidad_en_renglon": bool(patron_u.search(_plano(linea))),
                "anio": anio,
                "origen": origen,
            })
            break
    return salida


def _confiabilidad(ref: dict) -> str:
    """Nunca ALTA: sin equivalencia demostrada el dato es orientativo."""
    if ref["origen"] == "página" and ref["similitud"] >= 75 and ref["unidad_en_renglon"]:
        return "MEDIA"
    return "BAJA"


def investigar_partida(item: dict, *, buscar, leer=leer_pagina, tiempo_max=TIEMPO_MAX_PARTIDA) -> dict:
    """item: {id, descripcion, unidad, precio (opcional)}. buscar(consulta)->lista."""
    inicio = time.monotonic()
    descripcion, unidad = item["descripcion"], item["unidad"]
    precio_cot = item.get("precio")
    referencias, consultas_hechas, paginas_leidas, errores = [], [], 0, []
    vistas = set()

    for consulta in consultas_para(descripcion, unidad):
        if time.monotonic() - inicio > tiempo_max:
            break
        consultas_hechas.append(consulta)
        try:
            resultados = buscar(consulta) or []
        except Exception as error:
            errores.append(str(error)[:160])
            continue
        resultados = [r for r in resultados if r.get("url") and r["url"] not in vistas
                      and not any(x in r["url"].lower() for x in _EXCLUIR_URL)][:PAGINAS_POR_CONSULTA]
        vistas.update(r["url"] for r in resultados)

        def _procesar(r):
            texto = r.get("texto") or ""
            if not texto and time.monotonic() - inicio < tiempo_max:
                texto = leer(r["url"])
            refs = extraer_referencias(texto, descripcion=descripcion, unidad=unidad,
                                       precio_cotizado=precio_cot, url=r["url"], titulo=r.get("titulo", ""))
            if not refs:
                # Sin contenido legible: el fragmento del buscador queda
                # como último recurso, con confiabilidad BAJA.
                refs = extraer_referencias(r.get("resumen") or "", descripcion=descripcion, unidad=unidad,
                                           precio_cotizado=precio_cot, url=r["url"],
                                           titulo=r.get("titulo", ""), origen="fragmento del buscador")
            return bool(texto), refs

        with ThreadPoolExecutor(max_workers=4) as ejecutor:
            for tarea in as_completed([ejecutor.submit(_procesar, r) for r in resultados]):
                try:
                    leida, refs = tarea.result()
                except Exception as error:
                    errores.append(str(error)[:160])
                    continue
                paginas_leidas += int(leida)
                referencias.extend(refs)
        if any(_confiabilidad(r) == "MEDIA" for r in referencias):
            break

    if not referencias:
        return {
            "precio_mxn": None, "tiene_dato": False, "verificado": False,
            "unidad_encontrada": "", "fuente_nombre": "", "fuente_url": "",
            "descripcion_encontrada": "",
            "nota": (f"No se pudo consultar el buscador de respaldo ({errores[0]})." if errores and not paginas_leidas
                     else f"Sin precio comprobable: se revisaron {paginas_leidas} página(s) con "
                     f"{len(consultas_hechas)} búsqueda(s) y ningún renglón tenía el mismo material, "
                     "trabajo y unidad."),
            "consultas": consultas_hechas, "paginas_revisadas": paginas_leidas,
        }

    for r in referencias:
        r["confiabilidad"] = _confiabilidad(r)
    referencias.sort(key=lambda r: (r["confiabilidad"] == "MEDIA", r["similitud"], r["anio"] or ""), reverse=True)
    mejor = referencias[0]
    otras = [r for r in referencias[1:] if r["url"] != mejor["url"]][:3]
    return {
        "precio_mxn": mejor["precio"],
        "tiene_dato": True,
        "verificado": False,             # orientativo: equivalencia sin demostrar
        "confiabilidad": mejor["confiabilidad"],
        "unidad_encontrada": mejor["unidad"],
        "descripcion_encontrada": mejor["concepto"],
        "fuente_nombre": mejor["titulo"] or mejor["url"],
        "fuente_url": mejor["url"],
        "fecha_fuente": mejor["anio"],
        "fragmento": mejor["concepto"],
        "nota": (f"Precio leído en el {mejor['origen']} (concepto, unidad y precio en el mismo renglón); "
                 f"similitud de concepto {mejor['similitud']:.0f}/100. Orientativo: falta confirmar "
                 "especificación, alcance, fecha e IVA."),
        "otras_referencias": [
            {"precio": r["precio"], "fuente": r["titulo"] or r["url"], "url": r["url"],
             "confiabilidad": r["confiabilidad"]} for r in otras
        ],
        "consultas": consultas_hechas,
        "paginas_revisadas": paginas_leidas,
    }


def motor_disponible(tavily_key=None):
    """('Tavily', función) o ('buscador público', función) o (None, None)."""
    if tavily_key:
        return "Tavily", (lambda q: buscar_tavily(q, tavily_key))
    try:
        import ddgs  # noqa: F401
        return "buscador público", buscar_publico
    except Exception:
        return None, None


def investigar_lote(items: list[dict], tavily_key=None, registrar_error=None) -> dict:
    nombre, buscar = motor_disponible(tavily_key)
    if not buscar:
        if registrar_error:
            registrar_error("Respaldo web no disponible: falta 'tavily_api_key' o el paquete 'ddgs'.")
        return {}
    salida = {}
    with ThreadPoolExecutor(max_workers=3) as ejecutor:
        tareas = {ejecutor.submit(investigar_partida, it, buscar=buscar): str(it["id"]) for it in items}
        for tarea in as_completed(tareas):
            id_ = tareas[tarea]
            try:
                dato = tarea.result()
            except Exception as error:
                if registrar_error:
                    registrar_error(f"Respaldo web ({nombre}): {error}")
                continue
            dato["motor"] = f"Respaldo web ({nombre}, contenido verificado)"
            salida[id_] = dato
    return salida
