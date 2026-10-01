"""
Revision asistida con IA - capa OPCIONAL de segunda opinion
======================================================================
Este modulo NO reemplaza las fuentes de precio reales (NL, CDMX,
historico interno) -- esas siguen siendo la unica fuente de verdad
para el precio de referencia. Lo que hace esta capa es ayudar a
JUZGAR los casos que el comparador de texto (rapidfuzz) no puede
resolver bien por si solo:

  1. Coincidencias que se mandan a revision (por default, solo las de
     confianza BAJA o con precio muy alejado -- "dudosas"; si se activa
     "revisar todo" en la app, se manda CUALQUIER coincidencia con
     match, sin importar su confianza): el texto se parece pero no esta
     claro si es el mismo material, o simplemente se quiere una segunda
     opinion "por si acaso". Aqui la IA revisa el candidato y dice si de
     verdad es el mismo concepto o no, con una razon breve.

  2. Partidas SIN NINGUN dato de referencia (ningun match en NL, CDMX ni
     historico interno): la IA puede dar una opinion orientativa de si
     el precio cotizado suena razonable para ese tipo de material/
     equipo, basada en su conocimiento general -- pero esto NUNCA se
     presenta como un precio de mercado verificado, siempre queda
     marcado como "opinion de IA, no dato verificado".

SOBRE EL PROVEEDOR DE IA: en Ragasa el uso de Claude (Anthropic) esta
bloqueado por politica de la empresa, y todavia no esta confirmado si
se puede usar OpenAI/ChatGPT. Para no depender de una sola respuesta
que despues cambie, este modulo detecta AUTOMATICAMENTE cual proveedor
tiene una API key configurada y usa ese -- funciona igual sin importar
cual termine aprobado:

  - Gemini (Google):  Secrets -> gemini_api_key  (o env GEMINI_API_KEY)
  - OpenAI (ChatGPT):  Secrets -> openai_api_key  (o env OPENAI_API_KEY)

Si hay las dos, se prefiere Gemini, con respaldo automatico en OpenAI
si Gemini falla a la mitad de una revision (ej. se acabo la cuota). Si
no hay ninguna key, o falta instalar el paquete correspondiente
(`google-genai` / `openai`), o la llamada falla por cualquier motivo
(sin internet, rate limit, etc.), todas las funciones de este modulo
regresan None sin tronar la app -- exactamente el mismo patron que ya
usan ajuste_inflacion.py (token de INEGI) y historico_google_sheets.py
(credenciales de Google): la app sigue funcionando normal, solo sin
esta capa extra.
"""
import json
import os
import time

MODELO_GEMINI_POR_DEFECTO = "gemini-3.5-flash-lite"
MODELO_OPENAI_POR_DEFECTO = "gpt-5-mini"

# En cotizaciones grandes se hacen muchas llamadas seguidas a la IA (una
# por cada match que se manda a revision). El nivel gratuito de Gemini/
# OpenAI tiene un limite de solicitudes por minuto -- sin este reintento,
# en cuanto se topa ese limite a la mitad del lote, todas las llamadas
# restantes fallan silenciosamente y esos matches se quedan sin revisar
# (se ven identicos a uno ya confirmado, pero nadie los revisó).
_REINTENTOS_POR_RATE_LIMIT = 3
_ESPERA_BASE_SEGUNDOS = 3


def _es_error_rate_limit(excepcion) -> bool:
    texto = str(excepcion).lower()
    return any(
        pista in texto
        for pista in ("429", "rate limit", "rate_limit", "quota", "resource_exhausted", "resource exhausted", "too many requests")
    )


# ----------------------------------------------------------------------
# Diagnostico: guarda el ULTIMO error real de la revision con IA, para
# poder mostrarlo en la app (ej. "429 quota exceeded", "modelo no
# encontrado") en vez de solo decir "no se pudo revisar" sin explicar
# por que.
# ----------------------------------------------------------------------
_ultimo_error = {"mensaje": None, "proveedor": None}


def _registrar_error(error, proveedor=None):
    _ultimo_error["mensaje"] = str(error)
    _ultimo_error["proveedor"] = proveedor


def reiniciar_error():
    """Borra el error guardado: se llama al inicio de cada revisión para
    no mostrar errores viejos de una corrida anterior."""
    _ultimo_error["mensaje"] = None
    _ultimo_error["proveedor"] = None


def ultimo_error():
    """Regresa {'mensaje': str|None, 'proveedor': str|None} con el
    ultimo error real que dio la revision con IA en esta sesion, o
    {'mensaje': None, ...} si no ha habido ninguno."""
    return dict(_ultimo_error)


_proveedor_cache = {"proveedor": None, "cliente": None, "intentado": False}


def _leer_secret(nombre):
    try:
        import streamlit as st

        if nombre in st.secrets:
            return st.secrets[nombre]
    except Exception:
        pass
    return None


def _obtener_cliente(api_key=None, proveedor_forzado=None):
    """Detecta que proveedor de IA tiene una key configurada (Gemini
    primero, luego OpenAI) y regresa (proveedor, cliente). Si no hay
    ninguna key o falta el paquete correspondiente, regresa
    (None, None). Nunca lanza excepcion."""
    usar_cache = api_key is None and proveedor_forzado is None
    if usar_cache and _proveedor_cache["intentado"]:
        return _proveedor_cache["proveedor"], _proveedor_cache["cliente"]

    proveedor, cliente = None, None

    if proveedor_forzado in (None, "gemini"):
        key = api_key or os.environ.get("GEMINI_API_KEY") or _leer_secret("gemini_api_key")
        if key:
            try:
                from google import genai

                cliente = genai.Client(api_key=key)
                proveedor = "gemini"
            except Exception as error:
                _registrar_error(error, "gemini")
                cliente = None

    if cliente is None and proveedor_forzado in (None, "openai"):
        key = api_key or os.environ.get("OPENAI_API_KEY") or _leer_secret("openai_api_key")
        if key:
            try:
                import openai

                cliente = openai.OpenAI(api_key=key)
                proveedor = "openai"
            except Exception as error:
                _registrar_error(error, "openai")
                cliente = None

    if usar_cache:
        _proveedor_cache["intentado"] = True
        _proveedor_cache["proveedor"] = proveedor
        _proveedor_cache["cliente"] = cliente

    return proveedor, cliente


_clientes_cache = {"clientes": None, "intentado": False}


def _obtener_clientes(api_key_gemini=None, api_key_openai=None):
    """Regresa una lista ordenada de (proveedor, cliente) con TODOS los
    proveedores que tengan una API key configurada -- Gemini primero,
    OpenAI (ChatGPT de la empresa) despues. A diferencia de
    _obtener_cliente() (que solo regresa UNO), esta lista completa es
    la que permite el respaldo automatico: si Gemini se queda sin
    cuota a la mitad de una revision, se reintenta con OpenAI en vez
    de dejar esos matches sin revisar."""
    usar_cache = api_key_gemini is None and api_key_openai is None
    if usar_cache and _clientes_cache["intentado"]:
        return _clientes_cache["clientes"]

    clientes = []

    key_gemini = api_key_gemini or os.environ.get("GEMINI_API_KEY") or _leer_secret("gemini_api_key")
    if key_gemini:
        try:
            from google import genai

            clientes.append(("gemini", genai.Client(api_key=key_gemini)))
        except Exception as error:
            _registrar_error(error, "gemini")

    key_openai = api_key_openai or os.environ.get("OPENAI_API_KEY") or _leer_secret("openai_api_key")
    if key_openai:
        try:
            import openai

            clientes.append(("openai", openai.OpenAI(api_key=key_openai)))
        except Exception as error:
            _registrar_error(error, "openai")

    if usar_cache:
        _clientes_cache["intentado"] = True
        _clientes_cache["clientes"] = clientes

    return clientes


def proveedores_disponibles() -> list:
    """Nombres de los proveedores de IA activos ahora mismo (ej.
    ['gemini', 'openai']) -- lo usa la barra lateral de la app para
    mostrar si hay respaldo configurado."""
    return [proveedor for proveedor, _ in _obtener_clientes()]


def ia_disponible(api_key=None) -> bool:
    if api_key is not None:
        proveedor, cliente = _obtener_cliente(api_key)
        return cliente is not None
    return len(_obtener_clientes()) > 0


# Cuota DIARIA agotada (no el límite por minuto): reintentar no sirve y
# cada llamada fallida retrasa la revisión. Se "abre el circuito" 15 min y
# las demás partidas pasan de inmediato al siguiente proveedor o quedan
# como pendientes de validar.
_gemini_pausa = {"hasta": 0.0}


def _es_cuota_diaria(error) -> bool:
    texto = str(error).lower()
    return _es_error_rate_limit(error) and any(
        p in texto for p in ("per day", "perday", "daily", "limit: 0", "exceeded your current quota")
    )


def _llamar_gemini(cliente, prompt, modelo, max_tokens):
    if time.time() < _gemini_pausa["hasta"]:
        return None
    ultimo_error_local = None
    for intento in range(_REINTENTOS_POR_RATE_LIMIT):
        try:
            respuesta = cliente.models.generate_content(
                model=modelo,
                contents=prompt,
                config={
                    "response_mime_type": "application/json",
                    "max_output_tokens": max_tokens,
                },
            )
            return respuesta.text
        except Exception as error:
            ultimo_error_local = error
            if _es_cuota_diaria(error):
                _gemini_pausa["hasta"] = time.time() + 15 * 60
                _registrar_error(error, "gemini")
                return None
            if _es_error_rate_limit(error) and intento < _REINTENTOS_POR_RATE_LIMIT - 1:
                time.sleep(_ESPERA_BASE_SEGUNDOS * (intento + 1))
                continue
            break

    try:
        respuesta = cliente.models.generate_content(
            model=modelo,
            contents=prompt,
        )
        return respuesta.text
    except Exception as error:
        _registrar_error(error, "gemini")
        return None


def _llamar_openai(cliente, prompt, modelo, max_tokens):
    for intento in range(_REINTENTOS_POR_RATE_LIMIT):
        try:
            respuesta = cliente.chat.completions.create(
                model=modelo,
                max_completion_tokens=max_tokens,
                messages=[{"role": "user", "content": prompt}],
                response_format={"type": "json_object"},
            )
            return respuesta.choices[0].message.content
        except Exception as error:
            if _es_error_rate_limit(error) and intento < _REINTENTOS_POR_RATE_LIMIT - 1:
                time.sleep(_ESPERA_BASE_SEGUNDOS * (intento + 1))
                continue
            break

    try:
        respuesta = cliente.chat.completions.create(
            model=modelo,
            max_completion_tokens=max_tokens,
            messages=[{"role": "user", "content": prompt}],
        )
        return respuesta.choices[0].message.content
    except Exception as error:
        _registrar_error(error, "openai")
        return None


def _llamar_ia(proveedor, cliente, prompt, modelo=None, max_tokens=250):
    # Pequeña pausa antes de cada llamada para repartir las solicitudes
    # en el tiempo y no disparar el limite por minuto del nivel gratuito
    # cuando una cotizacion tiene muchos matches seguidos a revisar.
    time.sleep(1)

    if proveedor == "gemini":
        modelo = modelo or MODELO_GEMINI_POR_DEFECTO
        return _llamar_gemini(cliente, prompt, modelo, max_tokens), modelo
    if proveedor == "openai":
        modelo = modelo or MODELO_OPENAI_POR_DEFECTO
        return _llamar_openai(cliente, prompt, modelo, max_tokens), modelo
    return None, modelo


def _llamar_ia_con_respaldo(prompt, max_tokens=250, modelo=None):
    """Igual que _llamar_ia, pero si hay MAS DE UN proveedor de IA
    configurado (ej. Gemini y tambien el ChatGPT/OpenAI de la
    empresa), intenta primero con Gemini y, si falla -- ya agoto sus
    reintentos por rate limit, o cualquier otro error --, automatica-
    mente reintenta la MISMA peticion con el siguiente proveedor
    disponible en vez de rendirse. Antes, si Gemini se quedaba sin
    cuota a la mitad de una cotizacion grande, esos matches se
    quedaban sin revisar aunque hubiera una segunda API key
    configurada; con esto se aprovecha automaticamente.

    Regresa (texto, proveedor_usado, modelo_usado) -- proveedor_usado
    es None si ningun proveedor disponible logro responder."""
    clientes = _obtener_clientes()
    if not clientes:
        return None, None, None
    for proveedor, cliente in clientes:
        texto, modelo_usado = _llamar_ia(proveedor, cliente, prompt, modelo=modelo, max_tokens=max_tokens)
        if texto:
            return texto, proveedor, modelo_usado
    return None, None, None


def _extraer_json(texto):
    """La respuesta puede traer texto alrededor del JSON aunque se le
    pida limpio; esto saca el primer bloque {...} valido."""
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


def revisar_coincidencia_debil(
    descripcion_cotizada: str,
    unidad: str,
    descripcion_candidato: str,
    fuente: str,
    api_key=None,
    modelo=None,
):
    """
    Para un match ya encontrado: le pide a la IA que confirme o
    rechace si de verdad es el mismo material/concepto (no solo texto
    parecido).

    Regresa dict {'veredicto': 'CONFIRMA'|'RECHAZA'|'NO_SEGURO',
    'razon': str, 'modelo': str}, o None si la IA no esta disponible
    o la llamada fallo -- en ese caso quien llama debe seguir
    tratando el match como no revisado sin cambiar nada.
    """
    if not _obtener_clientes(api_key):
        return None

    prompt = (
        "Eres un experto en materiales y precios de construccion/obra "
        "en Mexico (civil, electrico, mecanico/HVAC, plomeria, "
        "acabados).\n\n"
        "Un sistema de busqueda por texto encontro esta posible "
        "coincidencia. Tu trabajo es decidir si de verdad es el MISMO "
        "concepto/material, no solo texto parecido.\n\n"
        f'Partida cotizada: "{descripcion_cotizada}" (unidad: {unidad})\n'
        f'Candidato encontrado en la base de {fuente}: '
        f'"{descripcion_candidato}"\n\n'
        "Responde SOLO con un JSON (sin texto alrededor) con este "
        "formato exacto:\n"
        '{"veredicto": "CONFIRMA" o "RECHAZA" o "NO_SEGURO", '
        '"razon": "explicacion breve en espanol, maximo 25 palabras"}\n\n'
        "CONFIRMA solo si estas genuinamente seguro de que es el "
        "mismo material/servicio (aunque este redactado distinto). "
        "RECHAZA si son materiales o servicios distintos aunque "
        "compartan palabras o medidas sueltas. NO_SEGURO si de "
        "verdad no se puede saber con la informacion dada."
    )

    texto, proveedor, modelo_usado = _llamar_ia_con_respaldo(prompt, max_tokens=200, modelo=modelo)
    datos = _extraer_json(texto)
    if not datos or "veredicto" not in datos:
        return None

    veredicto = str(datos.get("veredicto", "")).strip().upper()
    if veredicto not in {"CONFIRMA", "RECHAZA", "NO_SEGURO"}:
        veredicto = "NO_SEGURO"

    return {
        "veredicto": veredicto,
        "razon": str(datos.get("razon", "")).strip(),
        "modelo": f"{proveedor}:{modelo_usado}",
    }


def opinar_sin_datos(
    descripcion_cotizada: str,
    unidad: str,
    precio_cotizado: float,
    api_key=None,
    modelo=None,
):
    """
    Para una partida SIN NINGUN match en ninguna fuente: pide una
    opinion orientativa (NO un precio de mercado verificado) sobre si
    el precio cotizado suena razonable, basada en el conocimiento
    general del modelo.

    Regresa dict {'opinion': 'RAZONABLE'|'ALTO'|'BAJO'|
    'NO_PUEDO_OPINAR', 'razon': str, 'modelo': str}, o None si la IA
    no esta disponible o la llamada fallo.
    """
    if not _obtener_clientes(api_key):
        return None

    prompt = (
        "Eres un experto en materiales, equipo y precios de "
        "construccion/obra/instalaciones en Mexico (civil, "
        "electrico, mecanico/HVAC, plomeria, acabados), con "
        "conocimiento de precios de mercado aproximados en pesos "
        "mexicanos (MXN) a mediados de 2026.\n\n"
        "Esta partida NO tiene ninguna referencia en las bases de "
        "precios internas de la empresa (Nuevo Leon, CDMX, historico "
        "interno) -- por eso se te pide tu opinion general, sabiendo "
        "que NO es un dato verificado.\n\n"
        f'Partida: "{descripcion_cotizada}" (unidad: {unidad})\n'
        f"Precio unitario cotizado: ${precio_cotizado:,.2f} MXN\n\n"
        "Responde SOLO con un JSON (sin texto alrededor):\n"
        '{"opinion": "RAZONABLE" o "ALTO" o "BAJO" o '
        '"NO_PUEDO_OPINAR", "razon": "explicacion breve en espanol, '
        'maximo 30 palabras, menciona un rango aproximado si '
        'puedes"}\n\n'
        "Usa NO_PUEDO_OPINAR si el material/equipo es demasiado "
        "especifico (marca, modelo, capacidad exacta) como para dar "
        "una opinion responsable sin mas contexto."
    )

    texto, proveedor, modelo_usado = _llamar_ia_con_respaldo(prompt, max_tokens=200, modelo=modelo)
    datos = _extraer_json(texto)
    if not datos or "opinion" not in datos:
        return None

    opinion = str(datos.get("opinion", "")).strip().upper()
    if opinion not in {"RAZONABLE", "ALTO", "BAJO", "NO_PUEDO_OPINAR"}:
        opinion = "NO_PUEDO_OPINAR"

    return {
        "opinion": opinion,
        "razon": str(datos.get("razon", "")).strip(),
        "modelo": f"{proveedor}:{modelo_usado}",
    }


# Tamano de lote para las funciones "_lote": en vez de una llamada a la
# IA por cada partida a revisar (lento, y facil que se acabe el limite
# de solicitudes por minuto del nivel gratuito), se agrupan varias en un
# solo prompt y se pide un JSON con un resultado por cada una. Una
# cotizacion de 30 partidas con 20 matches a revisar pasa de ~20 llamadas
# a ~3, y de 2-4 minutos a unos 15-20 segundos.
TAMANO_LOTE = 8


def revisar_coincidencias_debiles_lote(items, api_key=None, modelo=None):
    """
    Version en lote de revisar_coincidencia_debil(): revisa varios
    matches en una sola llamada a la IA.

    items: lista de dicts con las claves:
        id (cualquier identificador unico, ej. un indice),
        descripcion_cotizada, unidad, descripcion_candidato, fuente

    Regresa dict {id: {'veredicto':..., 'razon':..., 'modelo':...}} --
    solo incluye los ids que la IA logro responder. Si la llamada
    completa falla, regresa {} (diccionario vacio, no None, para que
    quien llama pueda seguir iterando sin checar por None).
    """
    if not items or not _obtener_clientes(api_key):
        return {}

    lineas = []
    for it in items:
        lineas.append(
            f'ID {it["id"]}: partida cotizada = "{it["descripcion_cotizada"]}" '
            f'(unidad: {it["unidad"]}). Candidato encontrado en {it["fuente"]} = '
            f'"{it["descripcion_candidato"]}".'
        )

    prompt = (
        "Eres un experto en materiales y precios de construccion/obra "
        "en Mexico (civil, electrico, mecanico/HVAC, plomeria, "
        "acabados).\n\n"
        "Un sistema de busqueda por texto encontro estas posibles "
        "coincidencias. Para CADA una, decide si de verdad es el MISMO "
        "concepto/material, no solo texto parecido:\n\n"
        + "\n".join(lineas) +
        "\n\nResponde SOLO con un JSON (sin texto alrededor) con este "
        "formato exacto, un objeto por cada ID de la lista:\n"
        '{"resultados": [{"id": <mismo id de arriba>, "veredicto": '
        '"CONFIRMA" o "RECHAZA" o "NO_SEGURO", "razon": "explicacion '
        'breve en espanol, maximo 20 palabras"}, ...]}\n\n'
        "CONFIRMA solo si estas genuinamente seguro de que es el "
        "mismo material/servicio (aunque este redactado distinto). "
        "RECHAZA si son materiales o servicios distintos aunque "
        "compartan palabras o medidas sueltas. NO_SEGURO si de "
        "verdad no se puede saber con la informacion dada.\n\n"
        "Equivalencias de obra en Mexico que debes considerar el MISMO "
        "concepto: 'Armex' es acero de refuerzo prefabricado (varillas + "
        "estribos), equivale a un armado con varillas de la misma seccion; "
        "'block del numero 6' = block de 15 cm (15x20x40), 'del 4' = 10 cm, "
        "'del 8' = 20 cm; las 'columnas' de una barda de block son "
        "castillos; 'cerramiento' = cadena o dala de cerramiento; "
        "'concreto hecho en obra' = elaborado en obra. Diferencias menores "
        "de resistencia del concreto (f'c 150 vs 200) no cambian el "
        "concepto. Si la unica diferencia es de este tipo, CONFIRMA."
    )

    max_tokens = min(4000, 150 + 120 * len(items))
    texto, proveedor, modelo_usado = _llamar_ia_con_respaldo(prompt, max_tokens=max_tokens, modelo=modelo)
    datos = _extraer_json(texto)
    if not datos or "resultados" not in datos or not isinstance(datos["resultados"], list):
        if texto is None:
            _registrar_error(
                "no se pudo completar la revisión con IA (revisa el "
                "detalle técnico arriba para la causa real)",
                proveedor,
            )
        return {}

    salida = {}
    for r in datos["resultados"]:
        if not isinstance(r, dict) or "id" not in r:
            continue
        id_ = r["id"]
        veredicto = str(r.get("veredicto", "")).strip().upper()
        if veredicto not in {"CONFIRMA", "RECHAZA", "NO_SEGURO"}:
            veredicto = "NO_SEGURO"
        salida[id_] = {
            "veredicto": veredicto,
            "razon": str(r.get("razon", "")).strip(),
            "modelo": f"{proveedor}:{modelo_usado}",
        }
    return salida


def opinar_sin_datos_lote(items, api_key=None, modelo=None):
    """
    Version en lote de opinar_sin_datos(): da una opinion orientativa
    para varias partidas sin ningun dato de referencia, en una sola
    llamada a la IA.

    items: lista de dicts con las claves:
        id, descripcion_cotizada, unidad, precio_cotizado

    Regresa dict {id: {'opinion':..., 'razon':..., 'modelo':...}} --
    solo incluye los ids que la IA logro responder.
    """
    if not items or not _obtener_clientes(api_key):
        return {}

    lineas = []
    for it in items:
        lineas.append(
            f'ID {it["id"]}: "{it["descripcion_cotizada"]}" (unidad: '
            f'{it["unidad"]}), precio unitario cotizado: '
            f'${it["precio_cotizado"]:,.2f} MXN'
        )

    prompt = (
        "Eres un experto en materiales, equipo y precios de "
        "construccion/obra/instalaciones en Mexico (civil, "
        "electrico, mecanico/HVAC, plomeria, acabados), con "
        "conocimiento de precios de mercado aproximados en pesos "
        "mexicanos (MXN) a mediados de 2026.\n\n"
        "Estas partidas NO tienen ninguna referencia en las bases de "
        "precios internas de la empresa (Nuevo Leon, CDMX, historico "
        "interno) -- por eso se te pide tu opinion general de cada "
        "una, sabiendo que NO es un dato verificado:\n\n"
        + "\n".join(lineas) +
        "\n\nResponde SOLO con un JSON (sin texto alrededor) con este "
        "formato exacto, un objeto por cada ID de la lista:\n"
        '{"resultados": [{"id": <mismo id de arriba>, "opinion": '
        '"RAZONABLE" o "ALTO" o "BAJO" o "NO_PUEDO_OPINAR", "razon": '
        '"explicacion breve en espanol, maximo 25 palabras, menciona '
        'un rango aproximado si puedes"}, ...]}\n\n'
        "Usa NO_PUEDO_OPINAR si el material/equipo es demasiado "
        "especifico (marca, modelo, capacidad exacta) como para dar "
        "una opinion responsable sin mas contexto."
    )

    max_tokens = min(4000, 150 + 130 * len(items))
    texto, proveedor, modelo_usado = _llamar_ia_con_respaldo(prompt, max_tokens=max_tokens, modelo=modelo)
    datos = _extraer_json(texto)
    if not datos or "resultados" not in datos or not isinstance(datos["resultados"], list):
        if texto is None:
            _registrar_error(
                "no se pudo completar la opinión sobre partidas sin "
                "datos (revisa el detalle técnico arriba para la "
                "causa real)",
                proveedor,
            )
        return {}

    salida = {}
    for r in datos["resultados"]:
        if not isinstance(r, dict) or "id" not in r:
            continue
        id_ = r["id"]
        opinion = str(r.get("opinion", "")).strip().upper()
        if opinion not in {"RAZONABLE", "ALTO", "BAJO", "NO_PUEDO_OPINAR"}:
            opinion = "NO_PUEDO_OPINAR"
        salida[id_] = {
            "opinion": opinion,
            "razon": str(r.get("razon", "")).strip(),
            "modelo": f"{proveedor}:{modelo_usado}",
        }
    return salida


def debe_descartarse(fuente: dict, usar_ia: bool):
    """
    Decide si un match debe excluirse del resultado final.

    Se descarta si:
      - se mando a revision con IA (porque era dudoso, o porque se
        activo "revisar todo") Y la IA dijo RECHAZA o NO_SEGURO.

    Si usar_ia es False, o el match nunca se mando a revision (no
    tiene 'motivo' de duda y tampoco se le forzo una revision, asi que
    no hay 'revision_ia' en el diccionario), no se descarta -- se
    comporta igual que antes de tener esta capa.

    OJO: si SÍ se mando a revision pero la llamada nunca se completo
    (ej. se topo el limite de la cuenta gratuita), NO se descarta por
    eso solo -- es preferible seguir confiando en el match que el
    comparador de texto ya encontro, en vez de perder referencias de
    precio validas solo porque la IA no pudo confirmar (esto fue un
    problema real detectado en pruebas: partidas con match correcto
    terminaban en "SIN DATOS SUFICIENTES" nada mas porque la IA se
    quedo sin cuota a la mitad de la revision).

    Regresa (True/False, razon_para_mostrar_o_None).
    """
    si_se_mando_a_revision = bool(fuente.get('motivo')) or 'revision_ia' in fuente

    if not usar_ia or not si_se_mando_a_revision:
        return False, None

    revision = fuente.get('revision_ia')

    if revision is None:
        # Se mando a revisar pero la IA no respondio -- no se descarta
        # por precaución, se sigue tratando como un match valido.
        return False, None

    veredicto = revision.get('veredicto')

    if veredicto == 'RECHAZA':
        return True, 'la IA rechazó el match'

    if veredicto == 'NO_SEGURO':
        # "No estoy segura" solo descarta coincidencias que ya eran
        # debiles. Con "revisar TODAS" activado, la IA tambien revisa
        # coincidencias de confianza ALTA/MEDIA; un NO_SEGURO ahi no debe
        # tirar un match de texto bueno (caso real: cerramiento 15x20
        # hecho en obra descartado solo porque la cotizacion dice Armex).
        if str(fuente.get('confianza', '')).upper() in ('ALTA', 'MEDIA') and not fuente.get('motivo'):
            return False, None
        return True, 'la IA no pudo confirmar el match con seguridad'

    return False, None
