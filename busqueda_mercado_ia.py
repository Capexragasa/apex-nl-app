"""
Busqueda de precios de mercado en internet con IA (4a fuente: "IA busca cotizaciones")
=========================================================================================
A diferencia de revision_ia.py (que solo JUZGA matches ya encontrados contra
las bases de NL/CDMX/historico, o da una opinion generica basada en el
conocimiento del modelo), este modulo hace una BUSQUEDA REAL en internet para
encontrar precios de mercado actuales para cada partida, con la fuente citada.

Tiene DOS motores, en orden de preferencia:

1. Gemini con "grounding" (Google Search integrado a la API) -- requiere
   'gemini_api_key' en Secrets. Es el mas completo: el propio modelo busca,
   lee varias fuentes y arma un JSON estructurado por partida.
2. Tavily (https://tavily.com) como RESPALDO GRATUITO -- requiere
   'tavily_api_key' en Secrets (plan gratuito: 1,000 busquedas/mes, sin
   tarjeta). Se usa automaticamente SOLO cuando Gemini no esta disponible o
   se quedo sin cuota (error 429 u otro), para que la 4a fuente no se caiga
   por completo solo porque se acabo la cuota de Gemini.

Nunca se inventa un precio: si ninguno de los dos motores encuentra una
fuente real con una busqueda real, la partida se marca como 'sin_dato' en
vez de rellenarse con un numero estimado de memoria.

Para no gastar la cuota de la API rapido (la preocupacion explicita de
direccion), se agrupan varias partidas por llamada cuando el motor lo
permite (Gemini soporta lote; Tavily se llama 1 vez por partida porque su
API no acepta varias preguntas en una sola llamada).
"""
import json
import os
import re

# gemini-2.5-flash quedo con acceso limitado (solo cuentas que ya lo usaban
# antes) -- una API key nueva creada en AI Studio puede no tener acceso y
# la llamada falla en silencio. gemini-3.5-flash es el modelo vigente
# recomendado para proyectos nuevos y soporta grounding con Google Search.
MODELO_POR_DEFECTO = "gemini-3.5-flash"

# Menos partidas por lote que revision_ia.TAMANO_LOTE (8): cada partida aqui
# implica que el modelo dispare una o mas busquedas reales en Google, asi
# que el prompt y la respuesta esperada son mas pesados por partida.
TAMANO_LOTE = 5

TAVILY_URL = "https://api.tavily.com/search"

# ----------------------------------------------------------------------
# Diagnostico: guarda el ULTIMO error real de la busqueda (Gemini o
# Tavily), para poder mostrarlo en la app (ej. "404 model not found",
# "403 permission denied: grounding requiere facturacion habilitada",
# "429 quota exceeded") en vez de solo decir "no encontro nada" sin
# explicar por que.
# ----------------------------------------------------------------------
_ultimo_error = {"mensaje": None}


def _registrar_error(error):
    _ultimo_error["mensaje"] = str(error)


def ultimo_error():
    """Regresa {'mensaje': str|None} con el ultimo error real que dio la
    busqueda en internet con IA en esta sesion, o None si no ha habido
    ninguno (o si nunca se ha llamado)."""
    return dict(_ultimo_error)


def _leer_secret(nombre):
    try:
        import streamlit as st

        if nombre in st.secrets:
            return st.secrets[nombre]
    except Exception:
        pass
    return None


_cliente_cache = {"cliente": None, "intentado": False}


def _obtener_cliente(api_key=None):
    usar_cache = api_key is None
    if usar_cache and _cliente_cache["intentado"]:
        return _cliente_cache["cliente"]

    cliente = None
    key = api_key or os.environ.get("GEMINI_API_KEY") or _leer_secret("gemini_api_key")
    if key:
        try:
            from google import genai

            cliente = genai.Client(api_key=key)
        except Exception as error:
            _registrar_error(error)
            cliente = None

    if usar_cache:
        _cliente_cache["intentado"] = True
        _cliente_cache["cliente"] = cliente

    return cliente


def _obtener_tavily_key(api_key=None):
    return api_key or os.environ.get("TAVILY_API_KEY") or _leer_secret("tavily_api_key")


def _tavily_disponible(api_key=None) -> bool:
    return bool(_obtener_tavily_key(api_key))


def busqueda_disponible(api_key=None) -> bool:
    """True si hay AL MENOS un motor de busqueda configurado (Gemini o
    Tavily como respaldo gratuito)."""
    return _obtener_cliente(api_key) is not None or _tavily_disponible()


def _extraer_json(texto):
    if not texto:
        return None
    inicio = texto.find("{")
    fin = texto.rfind("}")
    if inicio == -1 or fin == -1 or fin < inicio:
        return None
    try:
        return json.loads(texto[inicio:fin + 1])
    except (json.JSONDecodeError, ValueError):
        return None


def _fuentes_de_respuesta(respuesta):
    """Extrae las URLs reales que Google Search consulto (grounding
    metadata) para poder citarlas -- nunca se muestra un precio de
    'busqueda de IA' sin poder decir de donde salio."""
    urls = []
    try:
        candidatos = respuesta.candidates or []
        for candidato in candidatos:
            metadata = getattr(candidato, "grounding_metadata", None)
            if not metadata:
                continue
            for chunk in getattr(metadata, "grounding_chunks", None) or []:
                web = getattr(chunk, "web", None)
                if web and getattr(web, "uri", None):
                    urls.append(
                        {"titulo": getattr(web, "title", None) or web.uri, "url": web.uri}
                    )
    except Exception:
        pass
    return urls


def _tokens_relevantes(texto):
    """Palabras que ayudan a verificar que la fuente habla del mismo concepto."""
    return {
        palabra for palabra in re.findall(r"[a-z0-9]+", (texto or "").lower())
        if len(palabra) >= 5 and palabra not in {
            "suministro", "instalacion", "colocacion", "precio", "mexico",
            "material", "unidad", "marca", "modelo", "equipo", "servicio",
            "concreto", "limpieza", "general", "pieza", "obra",
        }
    }


def _referencia_equivalente(item, descripcion_fuente, unidad_fuente):
    original = str(item.get("descripcion") or "")
    encontrada = str(descripcion_fuente or "")
    if not original.strip() or not encontrada.strip():
        return False
    esperada = str(item.get("unidad") or "").strip().upper()
    hallada = str(unidad_fuente or "").strip().upper()
    if esperada and hallada and esperada != hallada:
        sinonimos = {
            "PZA": {"PIEZA", "UNIDAD", "PZA"},
            "M2": {"M2", "M²", "METRO CUADRADO"},
            "M3": {"M3", "M³", "METRO CUBICO", "METRO CÚBICO"},
            "KG": {"KG", "KILO", "KILOGRAMO"},
        }
        if hallada not in sinonimos.get(esperada, {esperada}):
            return False
    codigos = _codigos_distintivos(original)
    # En equipos con varios códigos, todos deben estar presentes: compartir
    # solo "2x40A" no convierte un Chint en un Square D FAL 22040.
    codigos |= {
        re.sub(r"\s+", "", c).upper()
        for c in re.findall(r"\b[A-Za-z]{2,6}\s+\d{3,6}\b", original)
    }
    texto_fuente = _texto_plano_normalizado(encontrada)
    if codigos and not all(codigo in texto_fuente for codigo in codigos):
        return False
    marca = re.search(r"\bMARCA\s+([A-Za-z]+(?:\s+[A-Za-z])?)", original, re.I)
    if marca and _texto_plano_normalizado(marca.group(1)) not in texto_fuente:
        return False
    claves = _tokens_relevantes(original)
    if claves and not claves.intersection(_tokens_relevantes(encontrada)):
        return False
    return True


def _buscar_precios_mercado_gemini_lote(items, api_key=None, modelo=None):
    """Intenta el motor principal (Gemini + Google Search grounding).
    Regresa {} si no esta disponible o si la llamada falla (revisa
    ultimo_error() para saber por que)."""
    cliente = _obtener_cliente(api_key)
    if not cliente or not items:
        return {}

    modelo = modelo or MODELO_POR_DEFECTO

    lineas = []
    for it in items:
        lineas.append(f'ID {it["id"]}: "{it["descripcion"]}" (unidad: {it["unidad"]})')

    prompt = (
        "Busca en internet (usa la busqueda de Google) precios de mercado "
        "ACTUALES en Mexico (idealmente Nuevo Leon/Monterrey; si no hay, "
        "usa el precio nacional promedio) para cada uno de estos materiales "
        "o servicios de construccion/obra/instalaciones:\n\n"
        + "\n".join(lineas) +
        "\n\nPara CADA partida, busca de verdad (no inventes ni calcules de "
        "memoria) y responde SOLO con un JSON (sin texto alrededor):\n"
        '{"resultados": [{"id": <mismo id>, "precio_mxn": <numero o null '
        'si no encontraste nada confiable>, "unidad_encontrada": "<unidad '
        'del precio que encontraste>", "fuente_nombre": "<nombre del sitio '
        'o proveedor>", "fuente_url": "<url real de donde salio, o vacio '
        'si no aplica>", "nota": "<1 frase en espanol, ej. rango de precios '
        'o contexto>"}, ...]}\n\n'
        "Solo acepta equivalencia de modelo, marca, capacidad, alcance y unidad; no uses un precio de otro producto parecido. Si no puedes comprobar el producto exacto en la fuente, usa null. "
        "Si no encuentras un precio real y verificable para una partida, "
        "pon precio_mxn en null -- NUNCA inventes un numero."
    )

    try:
        respuesta = cliente.models.generate_content(
            model=modelo,
            contents=prompt,
            config={"tools": [{"google_search": {}}]},
        )
    except Exception as error:
        _registrar_error(f"Gemini: {error}")
        return {}

    texto = getattr(respuesta, "text", None)
    datos = _extraer_json(texto)
    if not datos or "resultados" not in datos or not isinstance(datos["resultados"], list):
        _registrar_error(
            f"Gemini: la respuesta no traia el JSON esperado. Texto crudo: "
            f"{(texto or '(vacio)')[:300]}"
        )
        return {}

    fuentes_citadas = _fuentes_de_respuesta(respuesta)
    url_generica = fuentes_citadas[0]["url"] if fuentes_citadas else ""

    salida = {}
    for r in datos["resultados"]:
        if not isinstance(r, dict) or "id" not in r:
            continue
        id_ = r["id"]
        precio = r.get("precio_mxn")
        try:
            precio = float(precio) if precio is not None else None
        except (TypeError, ValueError):
            precio = None
        item_original = next((it for it in items if str(it["id"]) == str(id_)), None)
        descripcion_encontrada = str(r.get("descripcion_encontrada", "") or "")
        unidad_encontrada = str(r.get("unidad_encontrada", "") or "")
        fuente_url = str(r.get("fuente_url", "") or "")
        # No asociar un precio a la primera URL genérica del lote: puede
        # corresponder a otra partida distinta.
        fuente_verificada = bool(fuente_url and any(
            fuente_url == fuente["url"] for fuente in fuentes_citadas
        ))
        equivalente = bool(item_original and _referencia_equivalente(
            item_original, descripcion_encontrada, unidad_encontrada
        ))
        if not fuente_verificada or not equivalente:
            precio = None
        salida[str(id_)] = {
            "precio_mxn": precio,
            "unidad_encontrada": unidad_encontrada,
            "fuente_nombre": str(r.get("fuente_nombre", "") or ""),
            "fuente_url": fuente_url if fuente_verificada else "",
            "nota": (str(r.get("nota", "") or "") if precio is not None else
                     "Sin precio validado: fuente, concepto o unidad no comprobables."),
            "tiene_dato": precio is not None,
            "motor": "Gemini (Google Search)",
        }
    return salida


# ----------------------------------------------------------------------
# Motor de respaldo gratuito: Tavily. No requiere tarjeta, 1,000
# busquedas/mes gratis. A diferencia de Gemini, su API no agrupa varias
# preguntas en una sola llamada, asi que aqui se llama una vez por
# partida. Tavily puede regresar un resumen ya redactado (include_answer)
# citando fuentes reales, pero no fuerza un JSON estructurado por precio,
# asi que aqui se intenta extraer un numero en pesos del texto real que
# regreso -- si no se encuentra ningun numero, se deja precio_mxn en None
# en vez de inventarlo.
# ----------------------------------------------------------------------
# Dos formas de citar un precio en el resumen: "$650" (signo ANTES del
# numero, lo mas comun cuando Tavily cita en dolares/pesos con formato
# tipo "$650 to $1,400 per day") o "650 MXN"/"650 pesos" (indicador
# DESPUES del numero). Se prueban ambas formas -- se necesitaba el
# indicador (antes o despues) para no agarrar cualquier numero suelto
# del texto (ej. un anio, una medida) como si fuera un precio.
_PATRON_PRECIO_PREFIJO = re.compile(
    r"\$\s?(\d{1,3}(?:,\d{3})*(?:\.\d+)?)",
)
_PATRON_PRECIO_SUFIJO = re.compile(
    r"(\d{1,3}(?:[,.]\d{3})*(?:\.\d+)?)\s*(?:mxn|pesos|mx\$)",
    re.IGNORECASE,
)


# Tercer problema real encontrado en pruebas, distinto a los dos
# anteriores: el resumen puede mencionar MAS DE UN precio, cada uno con
# una unidad distinta (ej. "$80-$220 MXN por metro cuadrado" Y "$363.44
# MXN el jornal diario" en la misma respuesta, para una partida cotizada
# por JORNAL). Antes se tomaba el PRIMER numero que apareciera en el
# texto sin importar de que unidad hablaba, lo que comparo un precio por
# M2 contra un precio cotizado por JORNAL -- unidades distintas, mismo
# error de fondo que comparar peras con manzanas, y generaba un
# "carísimo" o "ahorro potencial" completamente irreal.
#
# Ahora se buscan TODOS los precios que aparecen en el texto (no solo el
# primero) junto con las palabras que los rodean, y se prefiere el que
# tenga cerca alguna palabra clave de la MISMA unidad que se cotizo
# (ej. "jornal"/"por dia" para JORNAL, "m2"/"metro cuadrado" para M2).
#
# OJO -- caso real visto en pruebas que obligo a este segundo ajuste:
# a veces el texto SOLO trae el precio de una unidad distinta (ej. solo
# menciona "$300-500 pesos por metro cuadrado" para una partida cotizada
# por JORNAL, sin mencionar ningun precio por dia). Antes, si ninguna
# oracion coincidia con la unidad cotizada, se usaba el primer precio
# como respaldo -- pero ese respaldo segui comparando peras con
# manzanas (m2 contra jornal). Si la unidad cotizada SI esta en el mapa
# de palabras clave (es decir, se sabe como reconocerla) pero ninguna
# oracion con precio la menciona, es mas seguro no dar ningun precio
# que arriesgarse a comparar unidades distintas -- se regresa None
# (sin dato) en vez de adivinar. El respaldo de "usar el primero" solo
# aplica cuando la unidad NI SIQUIERA esta en el mapa (no hay forma de
# verificarla de ningun modo, asi que una aproximacion es mejor que
# nada).
# Bilingues a proposito: los resumenes de Tavily a veces salen en
# ingles (caso real visto en pruebas: "per square meter" en vez de
# "metro cuadrado"), asi que cada unidad trae sus palabras clave en
# espanol Y en ingles.
_MAPA_UNIDAD_PALABRAS_CLAVE = {
    "JORNAL": [
        "jornal", "por dia", "por día", "diario", "al dia", "al día", "jornada",
        "per day", "daily", "day rate", "a day",
    ],
    "M2": [
        "m2", "m²", "metro cuadrado", "metros cuadrados",
        "square meter", "square meters", "sq m", "sqm", "per m2",
    ],
    "M3": [
        "m3", "m³", "metro cubico", "metro cúbico", "metros cubicos", "metros cúbicos",
        "cubic meter", "cubic meters",
    ],
    "ML": [
        "ml", "metro lineal", "metros lineales",
        "linear meter", "linear meters", "per meter", "per linear",
    ],
    "KG": [
        "kg", "kilogramo", "kilogramos", "kilo", "por kilo",
        "per kg", "per kilogram", "kilogram",
    ],
    "TON": ["ton", "tonelada", "toneladas", "per ton", "tonne"],
    "PZA": [
        "pza", "pieza", "unidad", "c/u", "cada uno", "por pieza",
        "per piece", "per unit", "each",
    ],
    "LOTE": ["lote", "per lot", "lot"],
    "SERVICIO": ["servicio", "per service", "service"],
    "GLOBAL": ["global", "por lote", "lump sum"],
    "LT": ["litro", "litros", "por litro", "per liter", "per litre"],
}


_SEPARADOR_ORACIONES = re.compile(r"(?<=[.!?;])\s+|\n+")


def _extraer_precio_de_texto(texto, unidad=None):
    if not texto:
        return None

    # Se busca primero DENTRO DE CADA ORACION por separado (no una
    # ventana de caracteres a ciegas): una ventana de caracteres fija
    # puede "ver" la palabra clave de la SIGUIENTE oracion y confundirse
    # (probado: con una ventana de 40 caracteres, "$220 MXN por metro
    # cuadrado. El jornal..." hacia que el precio por M2 se marcara como
    # si fuera el del JORNAL, solo porque la palabra "jornal" caia
    # dentro de la ventana de la oracion anterior). Restringir la
    # busqueda a la MISMA oracion evita ese arrastre.
    palabras_clave = _MAPA_UNIDAD_PALABRAS_CLAVE.get((unidad or "").strip().upper(), [])
    if palabras_clave:
        for oracion in _SEPARADOR_ORACIONES.split(texto):
            oracion_lower = oracion.lower()
            if not any(palabra in oracion_lower for palabra in palabras_clave):
                continue
            m = _PATRON_PRECIO_PREFIJO.search(oracion) or _PATRON_PRECIO_SUFIJO.search(oracion)
            if m:
                try:
                    return float(m.group(1).replace(",", ""))
                except ValueError:
                    continue

        # La unidad SI esta en el mapa (se sabe reconocerla) pero
        # ninguna oracion con precio la menciono -- no se adivina con
        # el primer precio que aparezca, porque lo mas probable es que
        # este en OTRA unidad (ver comentario arriba).
        return None

    # La unidad cotizada no esta en el mapa de arriba (no hay forma de
    # reconocerla en el texto): se usa el primer precio que aparecio en
    # todo el texto como aproximacion, ya que no hay nada mejor con qué
    # decidir.
    coincidencias = sorted(
        list(_PATRON_PRECIO_PREFIJO.finditer(texto))
        + list(_PATRON_PRECIO_SUFIJO.finditer(texto)),
        key=lambda m: m.start(),
    )
    if not coincidencias:
        return None
    try:
        return float(coincidencias[0].group(1).replace(",", ""))
    except ValueError:
        return None


# ------------------------------------------------------------------
# Segunda alucinacion real encontrada en pruebas, distinta a la del
# regex: aunque el resumen de Tavily SI trae un precio con formato
# correcto, el propio resumen puede estar mal -- afirma con seguridad
# que un modelo especifico (ej. "AquaFlow ZX-500", un modelo inventado
# a proposito para una prueba) cuesta tanto, citando una fuente que en
# realidad es sobre un producto GENERICO distinto (ej. "Bomba
# sumergible domestica 1Hp"), no sobre ese modelo. Tavily arma su
# resumen con un modelo de lenguaje propio a partir de los resultados,
# y ese modelo puede "completar" con seguridad algo que los resultados
# reales no respaldan.
#
# Para blindarse contra esto: si la descripcion de la partida trae un
# codigo/modelo distintivo (letras+numeros pegados, ej. "ZX-500",
# "2X40A"), ese codigo debe aparecer LITERALMENTE en el texto real que
# Tavily regreso (titulos/contenido de los resultados, o el propio
# resumen) antes de confiar en el precio. Si la partida no trae ningun
# codigo asi (descripciones genericas como "Mano de obra"), no aplica
# esta verificacion -- no hay nada distintivo que confirmar.
# ------------------------------------------------------------------
_PATRON_CODIGO_DISTINTIVO = re.compile(r"[A-Za-z]{1,10}-?\d{2,6}[A-Za-z]{0,3}")


def _codigos_distintivos(descripcion):
    return {
        re.sub(r"[\s\-]", "", codigo).upper()
        for codigo in _PATRON_CODIGO_DISTINTIVO.findall(descripcion or "")
    }


def _texto_plano_normalizado(*fragmentos):
    return re.sub(r"[\s\-]", "", " ".join(fragmentos)).upper()


def _buscar_precio_tavily_item(item, api_key=None):
    try:
        import requests
    except Exception as error:
        _registrar_error(f"Tavily (respaldo gratuito): falta la libreria 'requests' ({error})")
        return None

    key = _obtener_tavily_key(api_key)
    if not key:
        return None

    consulta = (
        f'precio de "{item["descripcion"]}" ({item["unidad"]}) en México, '
        "pesos MXN, proveedor o ferretería actual"
    )
    try:
        resp = requests.post(
            TAVILY_URL,
            json={
                "api_key": key,
                "query": consulta,
                "search_depth": "basic",
                "include_answer": True,
                "max_results": 5,
            },
            timeout=20,
        )
        resp.raise_for_status()
        datos = resp.json()
    except Exception as error:
        _registrar_error(f"Tavily (respaldo gratuito): {error}")
        return None

    resumen = datos.get("answer") or ""
    resultados = datos.get("results") or []
    primera_url = resultados[0].get("url", "") if resultados else ""
    primer_titulo = resultados[0].get("title", "") if resultados else ""
    codigos_consulta = _codigos_distintivos(item["descripcion"])
    if codigos_consulta:
        for fuente in resultados:
            titulo = fuente.get("title", "") or ""
            if codigos_consulta.intersection(_codigos_distintivos(titulo)):
                primera_url = fuente.get("url", "") or ""
                primer_titulo = titulo
                break

    # OJO -- bug real encontrado en pruebas: antes, si el resumen de
    # Tavily no traía un precio, se buscaba un numero con pinta de precio
    # en el CONTENIDO CRUDO de los resultados de busqueda (paginas que
    # Tavily regreso pero que pueden no tener nada que ver con la
    # partida real si la busqueda no encontro algo relevante). Eso
    # provoco un caso real: para un modelo de bomba inventado para
    # pruebas, Tavily no encontro nada relevante, pero el regex agarro
    # un numero de una pagina de criptomonedas (CoinGecko) que
    # coincidencialmente traia "... MXN" en el texto, y ese numero
    # se presento como si fuera el precio de mercado -- justo lo que
    # nunca se debe hacer (inventar/adivinar un precio). Ahora SOLO se
    # confia en el resumen que el propio Tavily redacto para responder
    # la pregunta (resumen); si ahi no hay un precio claro, la partida
    # se marca sin dato en vez de arriesgarse a un numero de contexto
    # equivocado.
    _frases_sin_precio = (
        "no disponible", "not available", "no encontr", "no se encontr",
        "no data", "sin informacion", "sin información", "no information",
        "could not find", "no pricing", "no price",
    )
    resumen_normalizado = resumen.lower()
    hay_indicio_de_sin_dato = any(
        frase in resumen_normalizado for frase in _frases_sin_precio
    )

    precio = (
        None
        if hay_indicio_de_sin_dato
        else _extraer_precio_de_texto(resumen, item.get("unidad"))
    )

    # Verificación del código/modelo distintivo (ver comentario arriba de
    # _codigos_distintivos): si la partida menciona un modelo específico
    # y ese código no aparece en NINGÚN lado del texto real que Tavily
    # regresó, no se confía en el precio aunque el resumen lo afirme con
    # seguridad -- es la señal más confiable de que el resumen está
    # atribuyendo el precio de un producto genérico al modelo específico
    # que se buscaba.
    # IMPORTANTE: el resumen (`resumen`) NO se incluye en esta
    # verificación a propósito, aunque parezca que "ya trae el código".
    # Tavily redacta su resumen repitiendo palabras de la PREGUNTA que se
    # le mandó (que ya incluye el código, porque es la partida cotizada)
    # -- así que el código *siempre* va a aparecer en el resumen, sea o
    # no real el dato. Comprobar contra el resumen sería revisar la
    # afirmación contra sí misma. Por eso solo cuentan los TÍTULOS y
    # CONTENIDO de los resultados de búsqueda reales (las páginas web que
    # Tavily de verdad encontró): si el código no aparece ahí, es que
    # ninguna página real habla de ese modelo específico.
    codigo_no_verificado = False
    if precio is not None:
        codigos = _codigos_distintivos(item["descripcion"])
        if codigos:
            texto_disponible = _texto_plano_normalizado(
                *[
                    f"{r.get('title', '')} {r.get('content', '')}"
                    for r in resultados
                ],
            )
            if not any(codigo in texto_disponible for codigo in codigos):
                codigo_no_verificado = True
                precio = None

    # Un número en el resumen no basta si el enlace seleccionado no
    # identifica el producto de la partida.
    if precio is not None and not _referencia_equivalente(
        item, primer_titulo, item.get("unidad")
    ):
        codigo_no_verificado = True
        precio = None

    if codigo_no_verificado:
        nota = (
            (resumen + " " if resumen else "")
            + "[verificación: el modelo/código exacto de esta partida no "
            "se encontró en ninguna fuente real consultada -- se "
            "descarta el precio por precaución, aunque el resumen lo "
            "mencionaba]"
        )
    else:
        nota = resumen or "no se encontró un resumen con precio claro en los resultados"

    return {
        "precio_mxn": precio,
        "unidad_encontrada": item["unidad"],
        "fuente_nombre": (primer_titulo or "Búsqueda web (Tavily, respaldo gratuito)")[:120],
        "fuente_url": primera_url or "",
        "nota": nota[:300],
        "tiene_dato": precio is not None,
        "motor": "Tavily (respaldo gratuito)",
    }


def _buscar_precios_mercado_tavily_lote(items, api_key=None):
    salida = {}
    for it in items:
        resultado = _buscar_precio_tavily_item(it, api_key=api_key)
        if resultado is not None:
            salida[it["id"]] = resultado
    return salida


def buscar_precios_mercado_lote(items, api_key=None, modelo=None, tavily_api_key=None):
    """
    items: lista de dicts {id, descripcion, unidad}

    Busca en internet un precio de mercado actual en Mexico/Nuevo Leon
    para cada partida. Intenta primero Gemini (grounding con Google
    Search); si Gemini no esta configurado o falla (ej. se acabo la
    cuota), cae automaticamente a Tavily como respaldo gratuito.

    Regresa dict {id: {'precio_mxn': float|None, 'unidad_encontrada': str,
    'fuente_nombre': str, 'fuente_url': str, 'nota': str,
    'tiene_dato': bool, 'motor': str}} -- solo incluye los ids que se
    lograron consultar. Si ningun motor esta disponible o ambos fallan,
    regresa {} (usa ultimo_error() para ver por que).
    """
    if not items:
        return {}

    salida = _buscar_precios_mercado_gemini_lote(items, api_key=api_key, modelo=modelo)
    pendientes = [it for it in items if not salida.get(str(it["id"]), {}).get("tiene_dato")]
    # Si Gemini devolvió una coincidencia no verificable, intentar Tavily
    # solo para esa partida y conservar los precios confirmados del lote.
    if pendientes and _tavily_disponible(tavily_api_key):
        respaldo = _buscar_precios_mercado_tavily_lote(pendientes, api_key=tavily_api_key)
        for id_, dato in respaldo.items():
            if dato.get("tiene_dato") or id_ not in salida:
                salida[id_] = dato
    return salida
