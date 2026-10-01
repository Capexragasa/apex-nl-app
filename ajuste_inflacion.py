"""
Ajuste por inflacion (INPC - INEGI) para precios historicos de Nuevo Leon
==========================================================================
La base de Nuevo Leon casi no tiene informacion de 2024 en adelante, porque
el gobierno del estado no ha publicado licitaciones mas recientes en SIASI
(ver hoja "Fuente y Metodologia"). Para no comparar un precio de 2021 contra
una cotizacion de 2026 "en crudo", este modulo escala los precios viejos a su
equivalente de hoy usando el Indice Nacional de Precios al Consumidor (INPC)
que publica el INEGI - el mismo tipo de indice que se usa en la industria de
la construccion en Mexico para "escalatorias" de contratos (Art. 58 de la Ley
de Obras Publicas y Servicios Relacionados con las Mismas).

Ningun numero esta inventado: los niveles de INPC son los publicados de forma
oficial por el INEGI en sus boletines de prensa mensuales (fuente:
https://www.inegi.org.mx/temas/inpc/). Boletines usados:

  dic-2021: 117.314  (Boletin 9/22,   INEGI, 7 ene 2022)
  dic-2022: 126.539  (Boletin 11/23,  INEGI, 9 ene 2023)
  dic-2023: 132.373  (Boletin s/n,    INEGI, 9 ene 2024)
  dic-2024: 137.977  (Boletin 7/25,   INEGI, 9 ene 2025)
  dic-2025: 143.042  (Boletin 6/26,   INEGI, 8 ene 2026)
  jun-2026: 145.131  (Boletin 417/26, INEGI, 9 jul 2026)  <- dato mas reciente disponible

Limitacion reconocida (importante para no sobre-vender la precision de esto):
el INPC es un indice de precios al consumidor (canasta de gasto de los
hogares: comida, renta, transporte, etc.), NO un indice especifico de
insumos de construccion. El INEGI si publica un indice especializado para
eso (INPP - "Insumos de Obras Publicas"), pero no existe un portal publico
que permita descargar su serie historica completa de forma sencilla y
gratuita como el INPC. Usar el INPC es la aproximacion estandar cuando no
se tiene acceso al INPP, pero puede sub-estimar o sobre-estimar el alza real
de materiales muy volatiles (acero, cobre, cemento, combustibles).
"""

# Nivel del INPC (base: 2a quincena de julio de 2018 = 100).
# Fuente: boletines de prensa del INEGI, https://www.inegi.org.mx/temas/inpc/
INPC_NIVEL_DICIEMBRE = {
    # 2012-2020: necesarios porque muchos contratos de NL se concursaron en
    # esos años (el año real está en el número de licitación), aunque el
    # registro se haya publicado después.
    2012: 80.568,
    2013: 83.770,
    2014: 87.189,
    2015: 89.047,
    2016: 92.039,
    2017: 98.273,
    2018: 103.020,
    2019: 105.934,
    2020: 109.271,
    2021: 117.314,
    2022: 126.539,
    2023: 132.373,
    2024: 137.977,
    2025: 143.042,
}

# Dato mensual mas reciente disponible. Estos dos valores son el RESPALDO:
# si la descarga automatica de INEGI (mas abajo) no esta configurada o
# falla, la app sigue funcionando con estos numeros tal como estan aqui.
# Se actualizan a mano cada vez que se corre este archivo con --actualizar
# o cuando corre la tarea programada mensual.
NIVEL_ACTUAL = 145.131
ETIQUETA_ACTUAL = "junio 2026"
# Índice con el que se actualizan los precios. Por defecto el INPC general.
# Si en Secrets se configura 'inegi_indicador_construccion' (clave de un
# índice de precios de construcción del Banco de Información Económica de
# INEGI, p. ej. el INPP de construcción), se usa ese índice en su lugar.
INDICE_NOMBRE = "INPC general, INEGI (base 2a quincena julio 2018 = 100)"
INDICE_ES_CONSTRUCCION = False
FUENTE = "INEGI, Indice Nacional de Precios al Consumidor (INPC): https://www.inegi.org.mx/temas/inpc/"


# ==========================================================================
# DESCARGA AUTOMATICA DEL DATO MAS RECIENTE (API del Banco de Indicadores)
# ==========================================================================
# Para activar esto necesitas un token gratuito de INEGI:
#   1. Registrate en https://www.inegi.org.mx/app/api/indicadores/desarrolladores/
#   2. Copia el token que te dan.
#   3. Configuralo como variable de entorno INEGI_API_TOKEN, o si corres la
#      app en Streamlit, agrega en Secrets: inegi_api_token = "tu_token"
#
# IMPORTANTE sobre INDICADOR_INPC_MENSUAL: es la clave del indicador "Precios
# al Consumidor" (Indice general, Nacional, mensual) dentro del Banco de
# Informacion Economica (BIE) de INEGI. Clave 334360, verificada a mano
# contra https://www.inegi.org.mx/app/indicadores/?tm=3 (Banco de
# Informacion Economica) el 29-jul-2026: el dato de 2026/06 que regresa la
# API (145.131) coincide exacto con el boletin oficial 417/26. IMPORTANTE:
# los indicadores de BIE necesitan la fuente "BIE-BISE" en la URL (no
# "BISE" ni "BIE" solos), o la API regresa "No se encontraron resultados"
# aunque el token y el indicador sean correctos -- este fue el bug que
# causaba que la app se quedara siempre en el valor de respaldo.
# Si INEGI reasigna esta clave en el futuro (pasa cuando cambia el ano
# base del indice), veriflcala de nuevo ahi mismo y corre este archivo como
# script (ver abajo) para comparar el numero que regresa contra el ultimo
# dato publicado que conozcas.
import json
import re
import os
import time
import urllib.error
import urllib.request

INDICADOR_INPC_MENSUAL = os.environ.get("INEGI_INDICADOR_INPC", "334360")

_INEGI_API_URL = (
    "https://www.inegi.org.mx/app/api/indicadores/desarrolladores/jsonxml/"
    "INDICATOR/{indicador}/es/00/false/BIE-BISE/2.0/{token}?type=json"
)

_cache_inegi = {"resultado": None, "timestamp": 0.0}
_CACHE_TTL_SEGUNDOS = 6 * 60 * 60  # 6 horas, para no golpear la API en cada rerun


def _obtener_token(token=None):
    if token:
        return token
    token = os.environ.get("INEGI_API_TOKEN")
    if token:
        return token
    try:
        import streamlit as st
        if "inegi_api_token" in st.secrets:
            return str(st.secrets["inegi_api_token"]).strip().strip('"').strip()
    except Exception:
        pass
    return None


def consultar_inpc_inegi(token=None, indicador=None, timeout=10):
    """
    Descarga directo de la API de INEGI el dato MENSUAL mas reciente del
    INPC general nacional.

    Regresa {"nivel": float, "periodo": "AAAA/MM", "fuente": url} o None si
    no se pudo obtener (sin token, sin internet, respuesta inesperada,
    etc.). Nunca lanza excepcion hacia afuera: si algo falla, regresa None
    para que quien llama use el valor de respaldo (NIVEL_ACTUAL de arriba).
    """
    token = _obtener_token(token)
    if not token:
        return None
    indicador = indicador or INDICADOR_INPC_MENSUAL
    url = _INEGI_API_URL.format(indicador=indicador, token=token)
    try:
        with urllib.request.urlopen(url, timeout=timeout) as respuesta:
            datos = json.loads(respuesta.read().decode("utf-8"))
    except (urllib.error.URLError, TimeoutError, ValueError, OSError):
        return None
    try:
        observaciones = datos["Series"][0]["OBSERVATIONS"]
        for o in observaciones:
            try:
                _SERIE_MENSUAL[str(o["TIME_PERIOD"])] = float(o["OBS_VALUE"])
            except (KeyError, TypeError, ValueError):
                pass
        ultimo = max(observaciones, key=lambda o: str(o.get("TIME_PERIOD")))
        nivel = float(ultimo["OBS_VALUE"])
        periodo = ultimo["TIME_PERIOD"]
    except (KeyError, IndexError, TypeError, ValueError):
        return None
    return {"nivel": nivel, "periodo": periodo, "fuente": url.split("?")[0]}


_MESES = {
    "01": "enero", "02": "febrero", "03": "marzo", "04": "abril",
    "05": "mayo", "06": "junio", "07": "julio", "08": "agosto",
    "09": "septiembre", "10": "octubre", "11": "noviembre", "12": "diciembre",
}


def _etiqueta_desde_periodo(periodo: str) -> str:
    try:
        anio, mes = periodo.split("/")
        return f"{_MESES.get(mes, mes)} {anio}"
    except (ValueError, AttributeError):
        return str(periodo)


def refrescar_nivel_actual(token=None, forzar=False):
    """
    Intenta traer el dato mas reciente de la API de INEGI y, si lo logra,
    actualiza NIVEL_ACTUAL/ETIQUETA_ACTUAL en memoria para el resto de la
    sesion (afecta a factor_ajuste/ajustar_precio de aqui en adelante).

    Si falla por cualquier motivo (sin token configurado, sin internet,
    INEGI caido, etc.) NO modifica nada: se queda con los valores de
    respaldo definidos arriba. Usa un cache en memoria de
    _CACHE_TTL_SEGUNDOS para no llamar a la API en cada rerun de Streamlit.

    Regresa True si actualizo con un dato en vivo de INEGI, False si se
    quedo con el valor de respaldo.
    """
    global NIVEL_ACTUAL, ETIQUETA_ACTUAL

    ahora = time.time()
    if not forzar and (ahora - _cache_inegi["timestamp"]) < _CACHE_TTL_SEGUNDOS:
        resultado = _cache_inegi["resultado"]
        if resultado is not None:
            NIVEL_ACTUAL = resultado["nivel"]
            ETIQUETA_ACTUAL = _etiqueta_desde_periodo(resultado["periodo"])
        return resultado is not None

    resultado = consultar_inpc_inegi(token=token)
    _cache_inegi["timestamp"] = ahora
    _cache_inegi["resultado"] = resultado

    if resultado is None:
        return False

    NIVEL_ACTUAL = resultado["nivel"]
    ETIQUETA_ACTUAL = _etiqueta_desde_periodo(resultado["periodo"])
    _activar_indice_construccion(token)
    return True


def _leer_secret(nombre):
    valor = os.environ.get(nombre.upper())
    if valor:
        return valor
    try:
        import streamlit as st
        if nombre in st.secrets:
            return str(st.secrets[nombre]).strip()
    except Exception:
        pass
    return None


def _activar_indice_construccion(token=None):
    """Si hay un indicador de construcción configurado y la API responde con
    una serie mensual suficiente, lo deja como índice activo (sin mezclarlo
    con el INPC)."""
    global NIVEL_ACTUAL, ETIQUETA_ACTUAL, INDICE_NOMBRE, INDICE_ES_CONSTRUCCION
    indicador = _leer_secret("inegi_indicador_construccion")
    token = _obtener_token(token)
    if not indicador or not token:
        return False
    url = _INEGI_API_URL.format(indicador=indicador, token=token)
    try:
        with urllib.request.urlopen(url, timeout=12) as respuesta:
            datos = json.loads(respuesta.read().decode("utf-8"))
        observaciones = datos["Series"][0]["OBSERVATIONS"]
        serie = {str(o["TIME_PERIOD"]): float(o["OBS_VALUE"]) for o in observaciones
                 if o.get("OBS_VALUE") not in (None, "")}
    except Exception:
        return False
    if len(serie) < 24:
        return False
    ultimo = max(serie)
    _SERIE_MENSUAL.clear()
    _SERIE_MENSUAL.update(serie)
    NIVEL_ACTUAL = serie[ultimo]
    ETIQUETA_ACTUAL = _etiqueta_desde_periodo(ultimo)
    INDICE_NOMBRE = (_leer_secret("inegi_nombre_indice_construccion")
                     or f"Índice de precios de construcción, INEGI (indicador {indicador})")
    INDICE_ES_CONSTRUCCION = True
    return True


# Serie mensual del INPC ("AAAA/MM" -> nivel), llenada desde la API de INEGI
# cuando hay token. Sin token se estima el mes interpolando entre diciembres.
_SERIE_MENSUAL = {}


def _anio_valido(anio) -> int:
    try:
        return int(str(anio)[:4])
    except (TypeError, ValueError):
        return max(INPC_NIVEL_DICIEMBRE)


def indice_base(periodo):
    """(valor, etiqueta, método) del INPC para un periodo 'AAAA' o 'AAAA-MM'.

    - Con mes y serie mensual de INEGI: el valor publicado de ese mes.
    - Con mes sin serie: interpolación lineal entre diciembre del año anterior
      y diciembre del año (aproximación, se indica así).
    - Solo año: diciembre de ese año."""
    texto = str(periodo or "")
    m = re.match(r"(\d{4})(?:[-/](\d{1,2}))?", texto)
    disponibles = sorted(INPC_NIVEL_DICIEMBRE)
    anio = int(m.group(1)) if m else disponibles[-1]
    mes = int(m.group(2)) if m and m.group(2) else None
    if INDICE_ES_CONSTRUCCION and _SERIE_MENSUAL:
        # Índice de construcción activo: solo se usa su propia serie.
        clave = f"{anio}/{(mes or 12):02d}"
        if clave in _SERIE_MENSUAL:
            return _SERIE_MENSUAL[clave], _etiqueta_desde_periodo(clave), "índice mensual publicado (INEGI)"
        cercano = min(_SERIE_MENSUAL, key=lambda k: abs((int(k[:4]) * 12 + int(k[5:7])) - (anio * 12 + (mes or 12))))
        return (_SERIE_MENSUAL[cercano], _etiqueta_desde_periodo(cercano),
                f"mes más cercano disponible del índice ({_etiqueta_desde_periodo(cercano)})")
    if mes == 12 and anio in INPC_NIVEL_DICIEMBRE and f"{anio}/12" not in _SERIE_MENSUAL:
        return INPC_NIVEL_DICIEMBRE[anio], f"diciembre {anio}", "INPC de diciembre publicado (INEGI)"
    if mes:
        clave = f"{anio}/{mes:02d}"
        if clave in _SERIE_MENSUAL:
            return _SERIE_MENSUAL[clave], f"{_MESES[f'{mes:02d}']} {anio}", "INPC mensual publicado (INEGI)"
        if anio - 1 in INPC_NIVEL_DICIEMBRE and anio in INPC_NIVEL_DICIEMBRE:
            a0, a1 = INPC_NIVEL_DICIEMBRE[anio - 1], INPC_NIVEL_DICIEMBRE[anio]
            return (round(a0 + (a1 - a0) * mes / 12, 3), f"{_MESES[f'{mes:02d}']} {anio}",
                    "estimado interpolando entre diciembre y diciembre (sin serie mensual)")
    a_usado = min(max(anio, disponibles[0]), disponibles[-1])
    if mes:
        return (INPC_NIVEL_DICIEMBRE[a_usado], f"diciembre {a_usado}",
                f"APROXIMACIÓN: no hay INPC de {_MESES[f'{mes:02d}']} {anio} disponible (falta la serie mensual "
                f"de INEGI); se usó diciembre {a_usado}")
    return INPC_NIVEL_DICIEMBRE[a_usado], f"diciembre {a_usado}", "INPC de diciembre del año del dato"


def periodo_medio(fecha_min, fecha_max):
    """Mes a la mitad del periodo de los registros ('AAAA-MM')."""
    def _m(f):
        x = re.match(r"(\d{4})-(\d{1,2})", str(f or ""))
        return int(x.group(1)) * 12 + int(x.group(2)) - 1 if x else None
    a, b = _m(fecha_min), _m(fecha_max)
    if a is None and b is None:
        return None
    a = a if a is not None else b
    b = b if b is not None else a
    medio = (a + b) // 2
    return f"{medio // 12}-{medio % 12 + 1:02d}"


def factor_ajuste(anio) -> float:
    """Factor para llevar un precio del periodo 'anio' ('AAAA' o 'AAAA-MM')
    al nivel más reciente del INPC (NIVEL_ACTUAL)."""
    return NIVEL_ACTUAL / indice_base(anio)[0]


def ajustar_precio(precio: float, anio) -> float:
    if precio is None:
        return None
    try:
        return round(float(precio) * factor_ajuste(anio), 2)
    except (TypeError, ValueError):
        return None


def detalle_ajuste(periodo, fecha_min=None, fecha_max=None) -> dict:
    """Datos para reproducir el ajuste: índice, periodos, valores, factor y
    por qué se eligió el periodo base. No confirma vigencia ni equivalencia."""
    valor, etiqueta, metodo = indice_base(periodo)
    rango = f"{str(fecha_min)[:7]} a {str(fecha_max)[:7]}" if fecha_min and fecha_max else ""
    return {
        "indice": INDICE_NOMBRE,
        "periodo_base": etiqueta,
        "valor_base": valor,
        "periodo_final": ETIQUETA_ACTUAL,
        "valor_final": NIVEL_ACTUAL,
        "factor": round(NIVEL_ACTUAL / valor, 4),
        "justificacion": (
            ((f"Los registros van de {rango}; la mediana mezcla varios meses, así que se toma el mes "
              f"intermedio del periodo ({etiqueta}) como base. "
              if str(fecha_min)[:7] != str(fecha_max)[:7] else
              f"Todos los registros son de {etiqueta}; ese mes es la base. ") if rango else "")
            + f"Valor base: {metodo}."
        ),
    }


def desglose_acumulado(periodo) -> dict:
    """Inflación ACUMULADA (compuesta) desde el periodo base hasta hoy, con
    el desglose año por año. factor = INPC final ÷ INPC base = producto de
    (1 + inflación de cada tramo); no es una suma de porcentajes."""
    valor_base, etiqueta_base, metodo = indice_base(periodo)
    m = re.match(r"(\d{4})(?:[-/](\d{1,2}))?", str(periodo or ""))
    anio = int(m.group(1)) if m else None
    tramos, previo, etq_previa = [], valor_base, etiqueta_base
    if anio:
        anio_final = int(re.search(r"(\d{4})", ETIQUETA_ACTUAL).group(1)) if re.search(r"(\d{4})", ETIQUETA_ACTUAL) else anio
        for a in range(anio, anio_final):
            v = _SERIE_MENSUAL.get(f"{a}/12") or (None if INDICE_ES_CONSTRUCCION else INPC_NIVEL_DICIEMBRE.get(a))
            if not v or f"diciembre {a}" == etq_previa or v == previo:
                continue
            tramos.append({"de": etq_previa, "a": f"diciembre {a}", "valor_de": previo, "valor_a": v,
                           "inflacion_pct": round((v / previo - 1) * 100, 2),
                           "acumulada_pct": round((v / valor_base - 1) * 100, 2)})
            previo, etq_previa = v, f"diciembre {a}"
    if NIVEL_ACTUAL != previo:
        tramos.append({"de": etq_previa, "a": ETIQUETA_ACTUAL, "valor_de": previo, "valor_a": NIVEL_ACTUAL,
                       "inflacion_pct": round((NIVEL_ACTUAL / previo - 1) * 100, 2),
                       "acumulada_pct": round((NIVEL_ACTUAL / valor_base - 1) * 100, 2)})
    return {
        "indice": INDICE_NOMBRE, "base": etiqueta_base, "valor_base": valor_base,
        "final": ETIQUETA_ACTUAL, "valor_final": NIVEL_ACTUAL,
        "factor": round(NIVEL_ACTUAL / valor_base, 6),
        "acumulada_pct": round((NIVEL_ACTUAL / valor_base - 1) * 100, 2),
        "tramos": tramos, "metodo_base": metodo,
    }
