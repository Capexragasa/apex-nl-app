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
import os
import time
import urllib.error
import urllib.request

INDICADOR_INPC_MENSUAL = os.environ.get("INEGI_INDICADOR_INPC", "334360")

_INEGI_API_URL = (
    "https://www.inegi.org.mx/app/api/indicadores/desarrolladores/jsonxml/"
    "INDICATOR/{indicador}/es/00/true/BIE-BISE/2.0/{token}?type=json"
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
            return st.secrets["inegi_api_token"]
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
        ultimo = observaciones[-1]
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
    return True


def _anio_valido(anio) -> int:
    try:
        return int(str(anio)[:4])
    except (TypeError, ValueError):
        return max(INPC_NIVEL_DICIEMBRE)


def factor_ajuste(anio) -> float:
    """Factor multiplicador para llevar un precio de 'anio' a su equivalente
    en NIVEL_ACTUAL (INPC general INEGI). Ej: un precio de 2021 se multiplica
    por ~1.237 para estimar su equivalente en junio de 2026."""
    anio = _anio_valido(anio)
    anios_disponibles = sorted(INPC_NIVEL_DICIEMBRE)
    if anio < anios_disponibles[0]:
        anio = anios_disponibles[0]
    elif anio > anios_disponibles[-1]:
        anio = anios_disponibles[-1]
    base = INPC_NIVEL_DICIEMBRE[anio]
    return NIVEL_ACTUAL / base


def ajustar_precio(precio: float, anio) -> float:
    if precio is None:
        return None
    try:
        return round(float(precio) * factor_ajuste(anio), 2)
    except (TypeError, ValueError):
        return None


def detalle_ajuste(anio) -> dict:
    """Datos para reproducir el ajuste: índice, periodos, valores y factor.
    El ajuste por inflación NO confirma vigencia comercial ni equivalencia."""
    a = _anio_valido(anio)
    disponibles = sorted(INPC_NIVEL_DICIEMBRE)
    a_usado = min(max(a, disponibles[0]), disponibles[-1])
    base = INPC_NIVEL_DICIEMBRE[a_usado]
    return {
        "indice": "INPC general, INEGI (base 2a quincena julio 2018 = 100)",
        "periodo_base": f"diciembre {a_usado}",
        "valor_base": base,
        "periodo_final": ETIQUETA_ACTUAL,
        "valor_final": NIVEL_ACTUAL,
        "factor": round(NIVEL_ACTUAL / base, 4),
    }
