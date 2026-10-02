"""
Revisor de cotizaciones CAPEX - Nuevo León
==========================================

La aplicación acepta:

- Excel .xlsx
- Excel .xls
- Excel .xlsm
- CSV
- PDF digital con tablas
e:

- La hoja correcta
La aplicación intenta detectar automáticament
- La fila donde empiezan los encabezados
- La columna de concepto
- La columna de unidad
- La columna de cantidad
- La columna de precio unitario
- La columna de importe

Después convierte los datos al formato interno:

concepto | unidad | precio_unitario

y realiza la evaluación contra:

1. Histórico de precios de Nuevo León
2. Tabulador de precios de CDMX
3. Histórico interno de Google Sheets
"""

import io
import re
import unicodedata
from pathlib import Path

import pandas as pd
import plotly.graph_objects as go
import streamlit as st

from comparador_multifuente_v2 import (
    ComparadorMultiFuente,
    banda_en_mercado,
    clasificar,
)
import ajuste_inflacion
import revision_ia
import validacion_referencias as validacion
import revision_cantidades
import recomendacion
import mercado
import exportar_revision
import busqueda_mercado_ia


# ==========================================================
# CONFIGURACIÓN GENERAL
# ==========================================================

st.set_page_config(
    page_title="Revisor de cotizaciones CAPEX - NL",
    layout="wide",
)

BASE_PATH = "Base_Precios_Unitarios_NL_CDMX.xlsx"

DEFAULT_SHEET_ID = (
    "13cqz5_MwOcDHwrQ4rNBb9NFI8odWKLEAV_vYQN2p70g"
)


st.title("Revisor de cotizaciones CAPEX - Nuevo León")

st.caption(
    "La base de precios de Nuevo León, CDMX y el histórico interno "
    "ya están integrados. Sube tu cotización o licitación."
)
st.caption(
    "Jose Carlos W/H "
)

import version as _version
st.caption(f"Versión {_version.VERSION} · actualizada {_version.FECHA}")


# ==========================================================
# CARGA DEL COMPARADOR
# ==========================================================

@st.cache_resource
def cargar_comparador():
    return ComparadorMultiFuente(BASE_PATH)


@st.cache_resource
def cargar_historico():
    """
    Intenta conectar el histórico interno mediante Google Sheets.

    Cuando no existen credenciales configuradas, devuelve None.
    La aplicación continúa funcionando con Nuevo León y CDMX.
    """
    try:
        from historico_google_sheets import HistoricoGoogleSheets

        if "gcp_service_account" not in st.secrets:
            return None

        sheet_id = st.secrets.get(
            "sheet_id",
            DEFAULT_SHEET_ID,
        )

        return HistoricoGoogleSheets(
            sheet_id=sheet_id,
            creds_dict=st.secrets["gcp_service_account"],
        )

    except Exception as error:
        st.warning(
            "No se pudo conectar el histórico interno "
            f"de Google Sheets: {error}"
        )

        return None


# Edición del Tabulador General de Precios Unitarios de la CDMX cargado en
# Base_Precios_Unitarios_NL_CDMX.xlsx (hoja "Fuente y Metodología").
FECHA_TABULADOR_CDMX = "2026-05"

comparador = cargar_comparador()


def _acumulada(d: dict, periodo=None, registros=None) -> dict:
    """Añade al detalle de inflación la inflación ACUMULADA (compuesta) que
    realmente se aplicó: el % total, el rango entre renglones y el desglose
    año por año (cuyo producto es el factor)."""
    if not d:
        return d
    factor = d.get("factor")
    d["indice"] = ajuste_inflacion.INDICE_NOMBRE
    d["acumulada_pct"] = round((float(factor) - 1) * 100, 2) if factor else None
    registros = [r for r in (registros or []) if r.get("inflacion_acumulada_pct") is not None]
    if registros:
        _per = lambda r: str(r.get("periodo_precio") or r.get("fecha"))[:7]
        viejo = min(registros, key=_per)
        nuevo_ = max(registros, key=_per)
        d["acumulada_rango"] = (
            f"de +{nuevo_['inflacion_acumulada_pct']:.1f} % (precio más reciente, {_per(nuevo_)}) "
            f"a +{viejo['inflacion_acumulada_pct']:.1f} % (precio más antiguo, {_per(viejo)})"
            if viejo is not nuevo_ and viejo["inflacion_acumulada_pct"] != nuevo_["inflacion_acumulada_pct"]
            else f"+{viejo['inflacion_acumulada_pct']:.1f} % ({_per(viejo)})")
        periodo = periodo or _per(viejo)
    if periodo:
        try:
            desg = ajuste_inflacion.desglose_acumulado(str(periodo)[:7])
            d["tramos"] = desg["tramos"]
            d["tramos_desde"] = desg["base"]
            d["tramos_acumulada_pct"] = desg["acumulada_pct"]
        except Exception:
            pass
    return d


def _inflacion_nl(nl: dict) -> dict:
    """Detalle reproducible del ajuste por inflación de la referencia NL."""
    usados = nl.get("registros_usados") or 0
    metodo = str(nl.get("metodo_inflacion") or "")
    regs = nl.get("registros_detalle") or []
    mensual = bool(ajuste_inflacion._SERIE_MENSUAL)
    fuente_indice = "serie mensual publicada por INEGI" if mensual else "sin serie mensual: estimado entre diciembres"
    if metodo.startswith("mes del renglón") and regs:
        d = ajuste_inflacion.detalle_ajuste(nl.get("periodo_base_inflacion"))
        r0 = regs[0]
        d["justificacion"] = (
            f"La mediana del grupo (${nl.get('precio_mediana'):,.2f}) es el precio del contrato "
            f"{r0.get('licitacion') or ''} ({r0.get('dependencia') or ''}) del {r0.get('fecha')}; se actualizó con "
            f"el índice de ese mes ({fuente_indice}). p25 y p75 con el mes de su propio renglón cuando se localiza.")
        return _acumulada(d, nl.get("periodo_base_inflacion"))
    if metodo.startswith("mediana = promedio") and regs:
        original = nl.get("precio_mediana") or 0
        ajustada = nl.get("precio_mediana_ajustada") or 0
        return _acumulada({
            "periodo_base": " y ".join(sorted({str(r.get('fecha'))[:7] for r in regs})),
            "valor_base": None,
            "periodo_final": ajuste_inflacion.ETIQUETA_ACTUAL,
            "valor_final": ajuste_inflacion.NIVEL_ACTUAL,
            "factor": round(ajustada / original, 4) if original else None,
            "justificacion": (
                "La mediana es el promedio de dos contratos: " + "; ".join(
                    f"{r.get('licitacion') or ''} del {r.get('fecha')} (${r.get('precio'):,.2f})" for r in regs)
                + f". Cada uno se actualizó con el índice de su mes ({fuente_indice}) y se promedió; factor efectivo."),
        }, min(str(r.get('fecha'))[:7] for r in regs))
    todos_ = nl.get("registros_todos") or []
    todos = [r for r in todos_ if r.get("usado", True)]
    if todos:
        original = nl.get("precio_mediana") or 0
        ajustada = nl.get("precio_mediana_ajustada") or 0
        # La explicación sale de lo que realmente se hizo con cada renglón.
        metodos = {}
        for r in todos:
            _m = str(r.get("metodo_indice") or "").split(". Periodo:")[0]
            metodos[_m] = metodos.get(_m, 0) + 1
        texto_metodos = "; ".join(f"{n} renglón(es): {m}" for m, n in metodos.items())
        meses = sorted({str(r.get("periodo_precio") or r.get("fecha", ""))[:7] for r in todos})
        _viejas = sum(1 for r in todos if str(r.get("periodo_precio"))[:4] < str(r.get("fecha"))[:4])
        return _acumulada({
            "periodo_base": f"periodo de cada precio ({meses[0]} a {meses[-1]})",
            "valor_base": None,
            "periodo_final": ajuste_inflacion.ETIQUETA_ACTUAL,
            "valor_final": ajuste_inflacion.NIVEL_ACTUAL,
            "factor": round(ajustada / original, 4) if original else None,
            "justificacion": (
                (f"De {nl.get('n_contratos_total')} contratos localizados se usan los {nl.get('n_registros')} de los "
                 f"años más recientes ({', '.join(nl.get('anios_usados') or [])}): un precio reciente necesita menos "
                 "ajuste por inflación y refleja mejor el mercado actual; se toman años completos, del más nuevo "
                 "hacia atrás, hasta reunir al menos 3 contratos. Los más antiguos se muestran pero no entran al "
                 "precio. " if (nl.get("n_contratos_total") or 0) > (nl.get("n_registros") or 0) else
                 ("Hay menos de 3 contratos: se usan todos. " if (nl.get("n_registros") or 0) < 3 else ""))
                + f"P.U. NL = mediana de las medianas por contrato (cada contrato pesa lo mismo). Se usaron "
                f"{len(todos)} renglones técnicamente equivalentes (mismo objeto, función, material, unidad y "
                f"especificaciones compatibles) de {nl.get('n_registros')} contrato(s) (OCID). A cada renglón se le "
                f"aplicó la inflación ACUMULADA (compuesta) desde su mes hasta {ajuste_inflacion.ETIQUETA_ACTUAL} "
                f"({texto_metodos}). "
                + (f"En {_viejas} renglón(es) el año del precio es el de la licitación (viene en su número), anterior "
                   "a la fecha en que se publicó el registro; como el mes del concurso no está en la base, se usa "
                   "julio de ese año. " if _viejas else "")
                + "Después, mediana dentro de cada contrato y "
                "mediana entre contratos. La base no trae número de partida, por eso varios renglones de un mismo "
                "contrato cuentan como una sola observación. Factor = efectivo (P.U. actualizado ÷ mediana de "
                "medianas original). Detalle y fórmulas en la hoja 'Inflación NL'."
            ),
        }, None, todos)
    d = ajuste_inflacion.detalle_ajuste(
        nl.get("periodo_base_inflacion") or nl.get("anio_dato_mas_reciente"), nl.get("fecha_min"), nl.get("fecha_max"))
    d["justificacion"] = "APROXIMACIÓN (no se localizó el renglón de la mediana): " + d["justificacion"]
    return _acumulada(d, nl.get("periodo_base_inflacion") or nl.get("anio_dato_mas_reciente"))
historico = cargar_historico()

# Revisa una sola vez si hay una API key de IA configurada -- acepta
# Gemini (Secrets: gemini_api_key) o OpenAI (Secrets: openai_api_key),
# lo que esté disponible primero. Si no hay ninguna, la casilla de
# "Activar revisión con IA" en la barra lateral se muestra
# deshabilitada y la app sigue funcionando normal sin esta capa
# opcional.
ia_disponible = revision_ia.ia_disponible()

# La 4a fuente ("IA busca cotizaciones en internet") es independiente de la
# revision_ia de arriba: no juzga matches existentes, busca en internet un
# precio de mercado real para cada partida. Usa
# Gemini + Google Search (Secrets: gemini_api_key) como motor principal y,
# si no esta configurado o se quedo sin cuota, cae automaticamente a
# Tavily (Secrets: tavily_api_key) como respaldo 100% gratuito (1,000
# busquedas/mes, sin tarjeta). Si ninguna de las dos esta configurada,
# esta fuente se omite sin tronar la app.
busqueda_ia_disponible = busqueda_mercado_ia.busqueda_disponible()

# Alias local: la logica de "cuando descartar un match riesgoso" vive en
# revision_ia.py (compartida con comparador_multifuente_v2.py) para no
# duplicarla en dos archivos.
_revision_ia_descarta = revision_ia.debe_descartarse

def _baja_sin_confirmar(fuente_dict):
    """Coincidencia de confianza BAJA que la IA no confirmó.

    Su precio se muestra como referencia "por confirmar", pero no vota en
    el semáforo final: un texto poco parecido puede ser otro concepto
    (ej. "columnas para amarrar barda" contra bases de columnas metálicas).
    """
    if not fuente_dict or not fuente_dict.get("match"):
        return False
    if str(fuente_dict.get("confianza", "")).upper() != "BAJA":
        return False
    revision = fuente_dict.get("revision_ia") or {}
    return revision.get("veredicto") != "CONFIRMA"


def _alcance_distinto(cotizado, referencia):
    """Descarta diferencias explícitas de alcance que alteran el precio unitario."""
    original = normalizar_texto(cotizado or "")
    candidato = normalizar_texto(referencia or "")
    if not original or not candidato:
        return False
    # El precio de instalar, bombear o colocar incluye trabajos que una
    # partida de solo suministro no está comprando.
    solo_suministro = "suministro" in original and not any(
        palabra in original for palabra in ("colocacion", "instalacion", "aplicacion", "bombeo")
    )
    if solo_suministro and any(
        palabra in candidato for palabra in ("colocacion", "instalacion", "aplicacion", "bombeo")
    ):
        return True
    if "premezclado" in original and "elaborado en obra" in candidato:
        return True
    if "vinilica" in original and "esmalte" in candidato and "vinilica" not in candidato:
        return True
    return False




# Intenta traer el dato mas reciente del INPC directo de la API de INEGI.
# Si no hay token configurado (Secrets: inegi_api_token) o falla la
# conexion, no truena nada: se queda con el valor de respaldo de
# ajuste_inflacion.py.
try:
    inpc_en_vivo = ajuste_inflacion.refrescar_nivel_actual()
except Exception:
    inpc_en_vivo = False


# ==========================================================
# SINÓNIMOS DE COLUMNAS
# ==========================================================

COLUMNAS_OBJETIVO = {
    "partida": [
        "partida",
        "item",
        "ítem",
        "renglon",
        "renglón",
        "numero",
        "número",
        "no",
        "num",
        "clave",
        "posición",
        "posicion",
    ],
    "concepto": [
        "concepto",
        "descripcion",
        "descripción",
        "descripcion de los trabajos",
        "descripción de los trabajos",
        "alcance",
        "servicio",
        "material",
        "partida descripcion",
        "partida descripción",
        "trabajo",
        "trabajos",
    ],
    "unidad": [
        "unidad",
        "um",
        "u m",
        "u.m",
        "unid",
        "unidad de medida",
        "medida",
    ],
    "cantidad": [
        "cantidad",
        "cant",
        "volumen",
        "vol",
        "qty",
        "cantidad solicitada",
    ],
    "precio_unitario": [
        "precio unitario",
        "p u",
        "p.u",
        "pu",
        "unitario",
        "precio",
        "unit price",
        "costo unitario",
        "precio por unidad",
    ],
    "importe": [
        "importe",
        "total",
        "monto",
        "importe total",
        "precio total",
        "subtotal",
    ],
}


# ==========================================================
# FUNCIONES DE LIMPIEZA
# ==========================================================

def normalizar_texto(valor) -> str:
    """
    Convierte un texto a una forma comparable.

    Ejemplo:
    'Precio Unitario' -> 'precio unitario'
    """
    if valor is None:
        return ""

    texto = str(valor).strip().lower()

    texto = unicodedata.normalize(
        "NFKD",
        texto,
    )

    texto = "".join(
        caracter
        for caracter in texto
        if not unicodedata.combining(caracter)
    )

    texto = re.sub(
        r"[\n\r\t]+",
        " ",
        texto,
    )

    texto = re.sub(
        r"[^a-z0-9]+",
        " ",
        texto,
    )

    texto = re.sub(
        r"\s+",
        " ",
        texto,
    ).strip()

    return texto


def convertir_numero(valor):
    if valor is None:
        return None

    if isinstance(valor, (list, tuple, dict, set)):
        if not valor:
            return None

        if isinstance(valor, (list, tuple)) and len(valor) == 1:
            valor = valor[0]
        else:
            valor = " ".join(str(x) for x in valor)

    if isinstance(valor, (int, float)):
        if pd.isna(valor):
            return None
        return float(valor)

    texto = str(valor).strip()

    if texto.lower() in {"", "none", "nan", "null", "[]", "-", "--"}:
        return None

    texto = texto.replace("$", "")
    texto = texto.replace("MXN", "")
    texto = texto.replace("USD", "")
    texto = texto.replace("EUR", "")
    texto = texto.replace(",", "")
    texto = texto.replace(" ", "")
    texto = texto.replace("(", "-")
    texto = texto.replace(")", "")

    texto = re.sub(r"[^0-9.\-]", "", texto)

    if texto in {"", "-", ".", "-."}:
        return None

    try:
        return float(texto)
    except (ValueError, TypeError):
        return None


def normalizar_unidad(valor) -> str:
    """
    Normaliza unidades frecuentes.
    """
    texto = normalizar_texto(valor)

    equivalencias = {
        "m3": "M3",
        "m 3": "M3",
        "metro cubico": "M3",
        "metros cubicos": "M3",
        "m2": "M2",
        "m 2": "M2",
        "metro cuadrado": "M2",
        "metros cuadrados": "M2",
        "ml": "M",
        "metro lineal": "M",
        "metros lineales": "M",
        "m": "M",
        "kg": "KG",
        "kilogramo": "KG",
        "kilogramos": "KG",
        "ton": "TON",
        "tonelada": "TON",
        "toneladas": "TON",
        "pza": "PZA",
        "pieza": "PZA",
        "piezas": "PZA",
        "lote": "LOTE",
        "servicio": "SERVICIO",
        "juego": "JGO",
        "jgo": "JGO",
    }

    if texto in equivalencias:
        return equivalencias[texto]

    return str(valor).strip().upper()


# ==========================================================
# DETECCIÓN DE COLUMNAS
# ==========================================================

def detectar_campo(nombre_columna: str):
    """
    Relaciona una columna del proveedor con el formato interno.
    """
    columna_normalizada = normalizar_texto(
        nombre_columna
    )

    mejor_campo = None
    mejor_puntaje = 0

    for campo, sinonimos in COLUMNAS_OBJETIVO.items():

        for sinonimo in sinonimos:
            sinonimo_normalizado = normalizar_texto(
                sinonimo
            )

            if (
                columna_normalizada
                == sinonimo_normalizado
            ):
                puntaje = 100

            elif (
                sinonimo_normalizado
                in columna_normalizada
            ):
                puntaje = 80

            elif (
                columna_normalizada
                in sinonimo_normalizado
                and columna_normalizada
            ):
                puntaje = 70

            else:
                puntaje = 0

            if puntaje > mejor_puntaje:
                mejor_puntaje = puntaje
                mejor_campo = campo

    if mejor_puntaje >= 70:
        return mejor_campo

    return None


def evaluar_fila_encabezado(fila) -> int:
    """
    Calcula cuántos encabezados reconocibles contiene una fila.
    """
    campos_detectados = set()

    for valor in fila:
        campo = detectar_campo(
            str(valor)
        )

        if campo:
            campos_detectados.add(
                campo
            )

    puntaje = len(
        campos_detectados
    )

    if "concepto" in campos_detectados:
        puntaje += 2

    if "precio_unitario" in campos_detectados:
        puntaje += 2

    if "unidad" in campos_detectados:
        puntaje += 1

    return puntaje


def encontrar_encabezado(
    df_crudo: pd.DataFrame,
    limite_filas: int = 50,
):
    """
    Busca la fila donde realmente comienza la tabla.
    """
    mejor_fila = None
    mejor_puntaje = 0

    limite = min(
        limite_filas,
        len(df_crudo),
    )

    for indice in range(limite):

        fila = df_crudo.iloc[
            indice
        ].tolist()

        puntaje = evaluar_fila_encabezado(
            fila
        )

        if puntaje > mejor_puntaje:
            mejor_puntaje = puntaje
            mejor_fila = indice

    if mejor_puntaje < 6:
        return None, mejor_puntaje

    return mejor_fila, mejor_puntaje


# ==========================================================
# NORMALIZACIÓN DE TABLAS
# ==========================================================

def normalizar_dataframe(
    df_crudo: pd.DataFrame,
    nombre_origen: str = "",
) -> pd.DataFrame:
    """
    Convierte una tabla irregular al formato interno.
    """
    if df_crudo.empty:
        return pd.DataFrame()

    fila_encabezado, puntaje = encontrar_encabezado(
        df_crudo
    )

    if fila_encabezado is None:
        return pd.DataFrame()

    encabezados = []

    for indice, valor in enumerate(
        df_crudo.iloc[
            fila_encabezado
        ].tolist()
    ):
        if (
            valor is not None
            and not pd.isna(valor)
        ):
            encabezado = str(
                valor
            ).strip()

        else:
            encabezado = (
                f"columna_{indice}"
            )

        encabezados.append(
            encabezado
        )

    df = df_crudo.iloc[
        fila_encabezado + 1:
    ].copy()

    df.columns = encabezados

    df = df.reset_index(
        drop=True
    )

    mapa_columnas = {}

    for columna in df.columns:

        campo = detectar_campo(
            columna
        )

        if (
            campo
            and campo not in mapa_columnas
        ):
            mapa_columnas[
                campo
            ] = columna

    if "concepto" not in mapa_columnas:
        return pd.DataFrame()

    if "precio_unitario" not in mapa_columnas:
        return pd.DataFrame()

    resultado = pd.DataFrame()

    for campo in COLUMNAS_OBJETIVO:

        columna_origen = mapa_columnas.get(
            campo
        )

        if columna_origen is not None:
            resultado[
                campo
            ] = df[
                columna_origen
            ]

        else:
            resultado[
                campo
            ] = None

    resultado[
        "origen"
    ] = nombre_origen

    resultado[
        "fila_encabezado"
    ] = fila_encabezado + 1

    resultado[
        "puntaje_deteccion"
    ] = puntaje

    resultado[
        "concepto"
    ] = (
        resultado[
            "concepto"
        ]
        .fillna("")
        .astype(str)
        .str.strip()
    )

    resultado[
        "unidad"
    ] = resultado[
        "unidad"
    ].apply(
        normalizar_unidad
    )

    resultado[
        "cantidad"
    ] = resultado[
        "cantidad"
    ].apply(
        convertir_numero
    )

    resultado[
        "precio_unitario"
    ] = resultado[
        "precio_unitario"
    ].apply(
        convertir_numero
    )

    resultado[
        "importe"
    ] = resultado[
        "importe"
    ].apply(
        convertir_numero
    )
    resultado["cantidad"] = pd.to_numeric(
        resultado["cantidad"],
        errors="coerce",
    )

    resultado["precio_unitario"] = pd.to_numeric(
        resultado["precio_unitario"],
        errors="coerce",
    )

    resultado["importe"] = pd.to_numeric(
        resultado["importe"],
        errors="coerce",
    )

    importe_calculado = (
        resultado["cantidad"]
        * resultado["precio_unitario"]
    )

    resultado["importe"] = resultado["importe"].fillna(
        importe_calculado
    )

    palabras_excluir = [
        "subtotal",
        "iva",
        "impuesto",
        "total general",
        "gran total",
        "condiciones de pago",
        "vigencia de la oferta",
        "forma de pago",
        "notas",
        "observaciones",
    ]

    concepto_normalizado = resultado[
        "concepto"
    ].apply(
        normalizar_texto
    )

    mascara_excluir = pd.Series(
        False,
        index=resultado.index,
    )

    for palabra in palabras_excluir:

        palabra_normalizada = normalizar_texto(
            palabra
        )

        mascara_excluir |= (
            concepto_normalizado
            == palabra_normalizada
        )

    resultado = resultado[
        ~mascara_excluir
        & resultado[
            "concepto"
        ].ne("")
        & resultado[
            "precio_unitario"
        ].notna()
        & resultado[
            "precio_unitario"
        ].gt(0)
    ]

    resultado = resultado.reset_index(
        drop=True
    )

    return resultado


# ==========================================================
# LECTURA DE EXCEL
# ==========================================================

def leer_excel(archivo):
    """
    Revisa todas las hojas del Excel y selecciona
    la tabla más probable.
    """
    contenido = archivo.getvalue()

    extension = Path(
        archivo.name
    ).suffix.lower()

    if extension == ".xls":
        motor = "xlrd"
    else:
        motor = "openpyxl"

    hojas = pd.read_excel(
        io.BytesIO(contenido),
        sheet_name=None,
        header=None,
        dtype=object,
        engine=motor,
    )

    candidatos = []

    for nombre_hoja, df_crudo in hojas.items():

        normalizado = normalizar_dataframe(
            df_crudo,
            nombre_origen=nombre_hoja,
        )

        if not normalizado.empty:

            candidatos.append(
                {
                    "hoja": nombre_hoja,
                    "datos": normalizado,
                    "partidas": len(
                        normalizado
                    ),
                    "puntaje": normalizado[
                        "puntaje_deteccion"
                    ].max(),
                }
            )

    if not candidatos:
        raise ValueError(
            "No se encontró una tabla con concepto, "
            "unidad y precio unitario en ninguna hoja."
        )

    candidatos.sort(
        key=lambda item: (
            item["puntaje"],
            item["partidas"],
        ),
        reverse=True,
    )

    mejor = candidatos[0]

    confianza = min(
        100,
        int(
            mejor["puntaje"]
            / 9
            * 100
        ),
    )

    metadatos = {
        "tipo": "Excel",
        "hoja_detectada": mejor["hoja"],
        "partidas": mejor["partidas"],
        "confianza": confianza,
        "hojas_candidatas": [
            item["hoja"]
            for item in candidatos
        ],
    }

    return mejor["datos"], metadatos


# ==========================================================
# LECTURA DE CSV
# ==========================================================

def leer_csv(archivo):
    contenido = archivo.getvalue()

    intentos = [
        {
            "sep": None,
            "engine": "python",
        },
        {
            "sep": ",",
        },
        {
            "sep": ";",
        },
        {
            "sep": "\t",
        },
    ]

    ultimo_error = None

    for configuracion in intentos:

        try:
            df_crudo = pd.read_csv(
                io.BytesIO(contenido),
                header=None,
                dtype=object,
                encoding="utf-8-sig",
                **configuracion,
            )

            normalizado = normalizar_dataframe(
                df_crudo,
                nombre_origen="CSV",
            )

            if not normalizado.empty:

                confianza = min(
                    100,
                    int(
                        normalizado[
                            "puntaje_deteccion"
                        ].max()
                        / 9
                        * 100
                    ),
                )

                metadatos = {
                    "tipo": "CSV",
                    "hoja_detectada": "CSV",
                    "partidas": len(
                        normalizado
                    ),
                    "confianza": confianza,
                }

                return normalizado, metadatos

        except Exception as error:
            ultimo_error = error

    raise ValueError(
        "No fue posible interpretar el CSV. "
        f"Detalle: {ultimo_error}"
    )


# ==========================================================
# LECTURA DE PDF
# ==========================================================

def leer_pdf(archivo):
    """
    Lee PDF digitales con distintos formatos:

    1. Tablas de 8 columnas.
    2. Tablas de 6 columnas.
    3. Cotizaciones con una sola partida global.
    4. Cotizaciones de obra con descripciones de varias líneas.

    No utiliza OCR.
    """
    try:
        import pdfplumber

    except ImportError as error:
        raise RuntimeError(
            "Falta instalar pdfplumber. "
            "Agrégalo al archivo requirements.txt."
        ) from error

    contenido = archivo.getvalue()

    filas_validas = []
    paginas_detectadas = set()

    # ======================================================
    # PRIMER INTENTO: TABLAS DE 6 U 8 COLUMNAS
    # ======================================================

    with pdfplumber.open(
        io.BytesIO(contenido)
    ) as pdf:

        for numero_pagina, pagina in enumerate(
            pdf.pages,
            start=1,
        ):
            tablas = pagina.extract_tables() or []

            for tabla in tablas:

                if not tabla:
                    continue

                for fila in tabla:

                    if not fila or len(fila) < 6:
                        continue

                    valores = list(fila)

                    # Formato de 8 columnas:
                    # PDA | DESCRIPCIÓN | UNIDAD | CANTIDAD |
                    # MATERIAL | M.O. | P.U. | IMPORTE
                    if len(valores) >= 8:
                        partida = valores[0]
                        concepto = valores[1]
                        unidad = valores[2]
                        cantidad = valores[3]
                        precio_unitario = valores[6]
                        importe = valores[7]

                    # Formato de 6 columnas:
                    # CLAVE | DESCRIPCIÓN | UNIDAD |
                    # CANTIDAD | P.U. | IMPORTE
                    else:
                        partida = valores[0]
                        concepto = valores[1]
                        unidad = valores[2]
                        cantidad = valores[3]
                        precio_unitario = valores[4]
                        importe = valores[5]

                    # Algunas cotizaciones (ej. proveedores que usan
                    # PARTIDA|DESCRIPTION|CANTIDAD|UNIDAD|...) traen las
                    # columnas CANTIDAD y UNIDAD en el orden contrario al
                    # que asumimos arriba. En vez de adivinar el orden por
                    # posición fija, se detecta por el CONTENIDO: la
                    # columna "cantidad" siempre debe poder leerse como
                    # número (1.00, 53.70...) y "unidad" nunca (PIEZA,
                    # METROS...). Si están al revés, se intercambian antes
                    # de seguir procesando. Esto evita que un documento con
                    # el orden invertido termine con cantidad=None y
                    # unidad="1.00" (y por lo tanto pierda esas partidas
                    # frente a los otros métodos de lectura).
                    if (
                        convertir_numero(cantidad) is None
                        and convertir_numero(unidad) is not None
                    ):
                        cantidad, unidad = unidad, cantidad

                    partida_texto = (
                        str(partida).strip()
                        if partida is not None
                        else ""
                    )

                    concepto_texto = (
                        str(concepto).strip()
                        if concepto is not None
                        else ""
                    )

                    unidad_texto = (
                        str(unidad).strip()
                        if unidad is not None
                        else ""
                    )

                    concepto_texto = re.sub(
                        r"\s+",
                        " ",
                        concepto_texto,
                    ).strip()

                    if normalizar_texto(partida_texto) in {
                        "pda",
                        "partida",
                        "item",
                        "clave",
                    }:
                        continue

                    if normalizar_texto(concepto_texto) in {
                        "descripcion",
                        "descripción",
                    }:
                        continue

                    cantidad_numero = convertir_numero(
                        cantidad
                    )

                    precio_numero = convertir_numero(
                        precio_unitario
                    )

                    importe_numero = convertir_numero(
                        importe
                    )

                    if not concepto_texto:
                        continue

                    if (
                        precio_numero is None
                        or precio_numero <= 0
                    ):
                        continue

                    coincidencia_partida = re.search(
                        r"\d+(?:\.\d+)?",
                        partida_texto,
                    )

                    if coincidencia_partida:
                        partida_limpia = (
                            coincidencia_partida.group()
                        )
                    else:
                        partida_limpia = str(
                            len(filas_validas) + 1
                        )

                    filas_validas.append(
                        {
                            "partida": partida_limpia,
                            "concepto": concepto_texto,
                            "unidad": normalizar_unidad(
                                unidad_texto
                            ),
                            "cantidad": cantidad_numero,
                            "precio_unitario": precio_numero,
                            "importe": importe_numero,
                            "origen": (
                                f"Página {numero_pagina}"
                            ),
                            "fila_encabezado": None,
                            "puntaje_deteccion": 9,
                        }
                    )

                    paginas_detectadas.add(
                        numero_pagina
                    )
    # Detectar cotizaciones de obra con partidas decimales.
    # En este formato, la extracción automática de tablas
    # puede mezclar las descripciones de las partidas

    with pdfplumber.open(
        io.BytesIO(contenido)
    ) as pdf:
        texto_para_detectar_formato = "\n".join(
            pagina.extract_text() or ""
            for pagina in pdf.pages
        )

    encabezado_normalizado = normalizar_texto(
        texto_para_detectar_formato
    )

    claves_obra_detectadas = re.findall(
        r"\b\d+\.\d+\b",
        texto_para_detectar_formato,
    )

    es_cotizacion_obra = (
        "precio unitario" in encabezado_normalizado
        and "unidad" in encabezado_normalizado
        and "cantidad" in encabezado_normalizado
        and "importe" in encabezado_normalizado
        and len(set(claves_obra_detectadas)) >= 2
    )

    # OJO: antes esto vaciaba filas_validas (los renglones que el PRIMER
    # INTENTO -- extracción real de tabla con bordes -- ya había leído
    # BIEN) solo porque el texto de la página contiene las palabras
    # "unidad/cantidad/precio unitario/importe" (que aparecen en el
    # encabezado de CUALQUIER tabla de cotización, no solo en el formato
    # de obra con descripciones multilínea) y hay 2+ números con decimales
    # (que aparece en CUALQUIER tabla de precios). Eso hacía que una
    # tabla perfectamente bien leída se descartara a favor del método de
    # texto línea por línea (menos confiable: pierde renglones y a veces
    # pega texto de títulos/encabezados a la primera partida). Ahora solo
    # se vacía cuando el PRIMER INTENTO de plano no encontró casi nada
    # (menos de 3 renglones) -- ahí sí es señal real de que la tabla no
    # se pudo leer con bordes y conviene probar el método de texto.
    # Cuando el PRIMER INTENTO sí encontró una tabla completa, se deja
    # intacta y es la lógica de "TABLA CONFIABLE" de más abajo (que
    # compara conteo de renglones y, si hay Subtotal declarado, también
    # el importe total) la que decide cuál de los métodos usar -- ese
    # mecanismo ya es más preciso que este atajo.
    if es_cotizacion_obra and len(filas_validas) < 3:
        filas_validas = []
        paginas_detectadas = set()



    # ======================================================
    # SEGUNDO INTENTO: UNA SOLA PARTIDA GLOBAL
    # ======================================================

    if not filas_validas:

        with pdfplumber.open(
            io.BytesIO(contenido)
        ) as pdf:
            texto_completo = "\n".join(
                pagina.extract_text() or ""
                for pagina in pdf.pages
            )

        patron_servicio_global = re.search(
            r"Descripción\s+Cantidad\s+Precio\s+unitario"
            r"\s+Impuestos\s+Importe\s+"
            r"(.+?)\s+"
            r"(\d+(?:\.\d+)?)\s+"
            r"([A-Za-zÁÉÍÓÚáéíóúÑñ]+)\s+"
            r"\$?\s*([\d,]+\.\d{2})\s+"
            r"IVA\s*\(\s*16%\s*\)\s+"
            r"\$?\s*([\d,]+\.\d{2})",
            texto_completo,
            flags=re.IGNORECASE | re.DOTALL,
        )

        if patron_servicio_global:

            descripcion = re.sub(
                r"\s+",
                " ",
                patron_servicio_global.group(1),
            ).strip()

            cantidad = convertir_numero(
                patron_servicio_global.group(2)
            )

            unidad = normalizar_unidad(
                patron_servicio_global.group(3)
            )

            precio_unitario = convertir_numero(
                patron_servicio_global.group(4)
            )

            importe = convertir_numero(
                patron_servicio_global.group(5)
            )

            filas_validas.append(
                {
                    "partida": "1",
                    "concepto": descripcion,
                    "unidad": unidad,
                    "cantidad": cantidad,
                    "precio_unitario": precio_unitario,
                    "importe": importe,
                    "origen": "Página 1",
                    "fila_encabezado": None,
                    "puntaje_deteccion": 8,
                }
            )

            paginas_detectadas.add(1)

    # ======================================================
    # TERCER INTENTO: OBRA CON DESCRIPCIONES MULTILÍNEA
    # ======================================================

    filas_tabla = list(filas_validas)
    filas_obra = []
    lineas_pdf = []

    with pdfplumber.open(
        io.BytesIO(contenido)
    ) as pdf:
        for numero_pagina, pagina in enumerate(
            pdf.pages,
            start=1,
        ):
            texto_pagina = pagina.extract_text() or ""

            for linea in texto_pagina.splitlines():
                linea = re.sub(
                    r"\s+",
                    " ",
                    linea,
                ).strip()

                if linea:
                    lineas_pdf.append(
                        (numero_pagina, linea)
                    )

    patron_renglon_precio = re.compile(
        r"^(?:(\d+(?:\.\d+)?)[.\-)]?\s+)?"
        r"(.*?)"
        r"\b(M2|M3|ML|M|PZA|PZAS|SERVICIO|LOTE|KG|TON)\b"
        r"\s+([\d,]+(?:\.\d+)?)"
        r"\s+\$?\s*([\d,]+\.\d{2})"
        r"\s+\$?\s*([\d,]+\.\d{2})$",
        flags=re.IGNORECASE,
    )

    descripcion_acumulada = []

    textos_ignorar = [
        "clave descripcion unidad cantidad",
        "saro construcciones",
        "www saroconstrucciones com",
        "fecha monterrey",
        "cliente",
        "nombre de la empresa",
        "proyecto",
        "no de cotizacion",
        "subtotal",
        "iva",
        "total",
        "notas y condiciones",
    ]

    # Encabezado de la tabla ("CONCEPTO UNIDAD CANTIDAD PRECIO ...").
    # Todo lo que venga antes (título del documento, datos del cliente)
    # no es descripción de ninguna partida y se descarta al llegar aquí.
    palabras_encabezado = (
        "concepto", "descripcion", "unidad", "cantidad",
        "precio", "importe", "p u", "total",
    )

    def _es_encabezado_tabla(texto_normalizado):
        return sum(
            palabra in texto_normalizado
            for palabra in palabras_encabezado
        ) >= 3 and not re.search(r"\d+\.\d{2}", texto_normalizado)

    for numero_pagina, linea in lineas_pdf:
        linea_normalizada = normalizar_texto(
            linea
        )

        if _es_encabezado_tabla(linea_normalizada):
            descripcion_acumulada = []
            continue

        # Los títulos de sección marcan una partida nueva.
        if re.match(
            r"^(PRELIMINARES|BANQUETA|LIMPIEZA FINA)"
            r"(?:\s+\$[\d,]+\.\d{2})?$",
            linea,
            flags=re.IGNORECASE,
        ):
            descripcion_acumulada = []
            continue

        coincidencia = patron_renglon_precio.match(
            linea
        )

        if coincidencia:
            clave = coincidencia.group(1)
            descripcion_en_linea = (
                coincidencia.group(2).strip()
            )

            partes = list(
                descripcion_acumulada
            )

            if descripcion_en_linea:
                partes.append(
                    descripcion_en_linea
                )

            descripcion = re.sub(
                r"\s+",
                " ",
                " ".join(partes),
            ).strip()

            unidad = normalizar_unidad(
                coincidencia.group(3)
            )

            cantidad = convertir_numero(
                coincidencia.group(4)
            )

            precio_unitario = convertir_numero(
                coincidencia.group(5)
            )

            importe = convertir_numero(
                coincidencia.group(6)
            )

            if (
                descripcion
                and precio_unitario is not None
                and precio_unitario > 0
            ):
                filas_obra.append(
                    {
                        "partida": (
                            clave
                            if clave
                            else str(len(filas_obra) + 1)
                        ),
                        "concepto": descripcion,
                        "unidad": unidad,
                        "cantidad": cantidad,
                        "precio_unitario": precio_unitario,
                        "importe": importe,
                        "origen": f"Página {numero_pagina}",
                        "fila_encabezado": None,
                        "puntaje_deteccion": 8,
                    }
                )

                paginas_detectadas.add(
                    numero_pagina
                )

            descripcion_acumulada = []
            continue

        if any(
            texto in linea_normalizada
            for texto in textos_ignorar
        ):
            continue



        descripcion_acumulada.append(
            linea
        )

    # ======================================================
    # CUARTO INTENTO: COTIZACIONES ESTILO ODOO/ZOHO/FACTURAMA
    # (cantidad y unidad juntas, con "IVA(XX%)" en medio del renglón)
    # ======================================================
    # Formato muy comun en cotizaciones que llegan de proveedores reales
    # (ej. "SUMINISTRO DE TEE ... 4.00 Pieza 462.00 IVA(16%) $ 1,848.00"):
    # no hay columna de "Unidad" separada (va pegada a la cantidad), el
    # orden es cantidad->unidad->precio->impuesto->importe (al reves de
    # "TERCER INTENTO", que espera unidad->cantidad), la unidad es una
    # palabra completa como "Pieza" (no una abreviatura de una lista
    # fija), y casi siempre hay una linea "ENTREGA: X SEMANAS" pegada
    # despues de cada renglon que hay que ignorar para que no se cuele
    # en la descripcion del siguiente renglon.
    #
    # Las lineas de encabezado/pie de pagina (razon social, direccion,
    # RFC, telefono, "Pagina X / Y", etc.) se repiten IDENTICAS en cada
    # pagina del PDF -- en vez de tratar de adivinar el nombre de cada
    # proveedor, se detectan automaticamente como cualquier linea que
    # aparezca en TODAS las paginas.
    #
    # IMPORTANTE: el umbral debe ser "en todas las paginas", no solo
    # "en 2 o mas". Con productos similares (ej. varios cuples/reducciones
    # con la misma norma) es comun que dos renglones DISTINTOS compartan
    # una especificacion identica como "S/C BE, B16.25, C 40, A420 WPL6,
    # ANSI B16.9" -- eso puede repetirse en 2 paginas sin ser encabezado,
    # y si se ignora se pierde contenido real de la descripcion.
    total_paginas_pdf = len(
        {numero_pagina for numero_pagina, _ in lineas_pdf}
    )
    paginas_por_linea = {}
    for numero_pagina, linea in lineas_pdf:
        paginas_por_linea.setdefault(linea, set()).add(numero_pagina)
    lineas_repetidas_en_paginas = {
        linea
        for linea, paginas in paginas_por_linea.items()
        if total_paginas_pdf >= 2 and len(paginas) >= total_paginas_pdf
    }

    patrones_ruido_factura = [
        re.compile(r"^RFC\s*:", re.IGNORECASE),
        re.compile(r"P[aá]gina\s+\d+\s*/\s*\d+", re.IGNORECASE),
        re.compile(r"^N[uú]mero de cotizaci[oó]n", re.IGNORECASE),
        re.compile(r"^Direcci[oó]n de (facturaci[oó]n|env[ií]o)", re.IGNORECASE),
        re.compile(r"^ENTREGA\s*:", re.IGNORECASE),
        re.compile(r"^T[eé]rminos\s+(y\s+condiciones|de\s+pago)", re.IGNORECASE),
        re.compile(r"^Subtotal\b", re.IGNORECASE),
        # Renglon de TOTALES de IVA (sin parentesis), distinto del
        # token "IVA(16%)" que va DENTRO de cada renglon de partida.
        re.compile(r"^IVA\s+\d{1,2}\s*%", re.IGNORECASE),
        re.compile(r"^Total\s*\$", re.IGNORECASE),
        re.compile(r"[\w.\-]+@[\w.\-]+\.\w+"),
        re.compile(r"https?://\S+"),
        re.compile(r"^\+?\d[\d\s]{8,}\d"),
        # Linea de encabezado de la tabla en si.
        re.compile(
            r"^Descripci[oó]n\s+Cantidad\s+Precio\s+unitario",
            re.IGNORECASE,
        ),
        # Fecha suelta (dd/mm/aaaa), tipica debajo de "Fecha de
        # cotizacion"/"Vencimiento" cuando la etiqueta y el valor
        # vienen en lineas separadas.
        re.compile(r"^\d{1,2}/\d{1,2}/\d{2,4}\s*$"),
    ]

    # Etiquetas cuyo VALOR viene en el renglon de abajo (no en la misma
    # linea), asi que hay que saltarse tambien esa siguiente linea -
    # ej. "Vendedor" seguido de "Javier Vega" en la linea de abajo. Sin
    # esto, el nombre del vendedor se cuela en la descripcion de la
    # siguiente partida real.
    patrones_etiqueta_valor_abajo = [
        re.compile(r"^Fecha de cotizaci[oó]n\s*$", re.IGNORECASE),
        re.compile(r"^Vencimiento\s*$", re.IGNORECASE),
        re.compile(r"^Vendedor\s*$", re.IGNORECASE),
    ]

    patron_fila_servicio = re.compile(
        r"^(.*?)"
        r"\b(\d+(?:[.,]\d+)?)\s+"
        r"([A-Za-zÁÉÍÓÚáéíóúÑñ]{2,20})\s+"
        r"\$?\s*([\d,]+\.\d{2})\s+"
        r"(?:IVA\s*\(\s*\d{1,2}\s*%\s*\)\s+)?"
        r"\$?\s*([\d,]+\.\d{2})\s*$",
        flags=re.IGNORECASE,
    )

    filas_odoo = []
    descripcion_acumulada = []
    saltar_siguiente_linea = False
    # En PDFs tipo Odoo, cuando la descripcion es larga el sobrante NO
    # queda ANTES de la linea con los numeros, sino DESPUES (el motor de
    # layout del PDF "recorta" la descripcion en la columna y el resto
    # cae en la siguiente linea, antes del "ENTREGA: ..."). Ejemplo real:
    #   'SUMINISTRO DE TEE REDUCC. SOLD. CED-40 DE 1-1/4" 4.00 Pieza
    #    462.00 IVA(16%) $ 1,848.00'
    #   'X 1/2" ASTM A-234 WPB'          <- esto sigue siendo la MISMA
    #                                       partida, no la siguiente.
    #   'ENTREGA: 2 A 3 SEMANAS'
    # Esta bandera indica "la ultima linea cerro un renglon con numeros;
    # si la siguiente linea no es ruido ni un renglon nuevo, es el
    # sobrante de la descripcion de ESE renglon, hay que pegarlo ahi".
    esperando_continuacion_descripcion = False

    for numero_pagina, linea in lineas_pdf:

        if saltar_siguiente_linea:
            saltar_siguiente_linea = False
            esperando_continuacion_descripcion = False
            continue

        if linea in lineas_repetidas_en_paginas:
            descripcion_acumulada = []
            esperando_continuacion_descripcion = False
            continue

        if any(
            patron.search(linea)
            for patron in patrones_etiqueta_valor_abajo
        ):
            descripcion_acumulada = []
            esperando_continuacion_descripcion = False
            saltar_siguiente_linea = True
            continue

        if any(
            patron.search(linea)
            for patron in patrones_ruido_factura
        ):
            descripcion_acumulada = []
            esperando_continuacion_descripcion = False
            continue

        coincidencia = patron_fila_servicio.match(linea)

        if coincidencia:
            descripcion_en_linea = coincidencia.group(1).strip()

            partes = list(descripcion_acumulada)
            if descripcion_en_linea:
                partes.append(descripcion_en_linea)

            descripcion = re.sub(
                r"\s+", " ", " ".join(partes)
            ).strip()

            cantidad = convertir_numero(coincidencia.group(2))
            unidad = normalizar_unidad(coincidencia.group(3))
            precio_unitario = convertir_numero(coincidencia.group(4))
            importe = convertir_numero(coincidencia.group(5))

            if (
                descripcion
                and precio_unitario is not None
                and precio_unitario > 0
            ):
                filas_odoo.append(
                    {
                        "partida": str(len(filas_odoo) + 1),
                        "concepto": descripcion,
                        "unidad": unidad,
                        "cantidad": cantidad,
                        "precio_unitario": precio_unitario,
                        "importe": importe,
                        "origen": f"Página {numero_pagina}",
                        "fila_encabezado": None,
                        "puntaje_deteccion": 8,
                    }
                )
                paginas_detectadas.add(numero_pagina)
                esperando_continuacion_descripcion = True
            else:
                esperando_continuacion_descripcion = False

            descripcion_acumulada = []
            continue

        if esperando_continuacion_descripcion and filas_odoo:
            filas_odoo[-1]["concepto"] = re.sub(
                r"\s+",
                " ",
                f"{filas_odoo[-1]['concepto']} {linea}",
            ).strip()
            esperando_continuacion_descripcion = False
            continue

        descripcion_acumulada.append(linea)

    # Para decidir cuál de los 3 métodos (tabla con bordes, obra
    # multilínea, facturas estilo Odoo) usar, primero se intenta contra
    # el SUBTOTAL declarado en el PDF -- comparar la suma de importes de
    # cada método contra ese número es mucho más confiable que solo
    # contar renglones, porque un método puede "ganar" por tener más
    # filas pero con datos rotos o mal alineados (columnas de una tabla
    # leídas en el orden equivocado, texto de factura cortado mal,
    # etc.). Ojo: se compara contra el Subtotal (antes de impuestos), NO
    # el Total, porque cada importe de renglón ya viene sin IVA.
    #
    # El PDF a veces trae el número del subtotal con un espacio suelto
    # en medio (ej. "$ 5 72,088.24" en vez de "$ 572,088.24" -- artefacto
    # de cómo el PDF codifica ese texto), así que el patrón permite
    # espacios/tabs sueltos dentro del número (pero no saltos de línea,
    # para no cruzarse con la línea del IVA que le sigue).
    coincidencia_subtotal = re.search(
        r"Subtotal[ \t]*\$?[ \t]*(\d[\d,\t ]*\.\d{2})",
        texto_para_detectar_formato,
        flags=re.IGNORECASE,
    )
    subtotal_declarado = (
        convertir_numero(coincidencia_subtotal.group(1))
        if coincidencia_subtotal
        else None
    )

    def _suma_importes(filas):
        total = 0.0
        for f in filas:
            importe = convertir_numero(f.get("importe"))
            if importe is None:
                cantidad = convertir_numero(f.get("cantidad"))
                precio = convertir_numero(f.get("precio_unitario"))
                if cantidad is not None and precio is not None:
                    importe = cantidad * precio
                else:
                    importe = 0.0
            total += importe
        return total

    metodo_por_subtotal = None
    if subtotal_declarado:
        mejor_diferencia = None
        for nombre, filas in (
            ("tabla", filas_tabla),
            ("obra", filas_obra),
            ("odoo", filas_odoo),
        ):
            if not filas:
                continue
            diferencia = abs(_suma_importes(filas) - subtotal_declarado)
            if mejor_diferencia is None or diferencia < mejor_diferencia:
                mejor_diferencia = diferencia
                metodo_por_subtotal = nombre
        # Tolerancia de $1 (redondeos de centavos). Si ni el que más se
        # acerca cuadra de verdad, no forzar esta señal: mejor caer al
        # criterio de respaldo de abajo.
        if mejor_diferencia is None or mejor_diferencia >= 1.0:
            metodo_por_subtotal = None

    # IMPORTANTE: el subtotal por sí solo NO basta para decidir. Un método
    # de texto (TERCER/CUARTO INTENTO) puede sumar exactamente el
    # subtotal correcto renglón por renglón (los números de cada línea se
    # leen bien) pero con las DESCRIPCIONES desfasadas una posición
    # (porque el texto de una celda que ocupa varias líneas se agrupó con
    # el renglón de números equivocado) -- en ese caso el resultado
    # "cuadra" en dinero pero cada partida queda con el texto de otra, lo
    # cual es inútil para buscar coincidencias en la base de precios. Por
    # eso primero se prioriza la extracción de tabla real (PRIMER
    # INTENTO, basada en las líneas de borde que trae el PDF), que
    # respeta la celda de descripción de cada renglón tal cual está en el
    # documento y no depende de heurísticas de texto: si encontró un
    # número de renglones razonable frente a los otros métodos (>=90%) Y
    # su suma no se aleja mucho del subtotal (dentro de 2%, para tolerar
    # como mucho 1-2 renglones que la tabla no haya podido leer por un
    # salto de página), se usa esa. Solo si la tabla no es confiable se
    # cae al criterio de subtotal exacto entre los métodos de texto, y si
    # tampoco hay subtotal, al conteo de renglones.
    max_otros = max(len(filas_obra), len(filas_odoo))
    tabla_confiable = False
    if filas_tabla and len(filas_tabla) >= max_otros * 0.9:
        if subtotal_declarado:
            diferencia_relativa = (
                abs(_suma_importes(filas_tabla) - subtotal_declarado)
                / subtotal_declarado
            )
            tabla_confiable = diferencia_relativa <= 0.02
        else:
            tabla_confiable = True

    if tabla_confiable:
        filas_validas = filas_tabla
    elif metodo_por_subtotal == "tabla":
        filas_validas = filas_tabla
    elif metodo_por_subtotal == "obra":
        filas_validas = filas_obra
    elif metodo_por_subtotal == "odoo":
        filas_validas = filas_odoo
    elif len(filas_odoo) > len(filas_obra) and len(filas_odoo) > len(filas_tabla):
        filas_validas = filas_odoo
    elif len(filas_obra) > len(filas_tabla):
        filas_validas = filas_obra
    else:
        filas_validas = filas_tabla

    # ======================================================
    # QUINTO INTENTO: PRESUPUESTOS DE OBRA TIPO OPUS/NEODATA
    # CON NÚMEROS "PARTIDOS" POR EL PDF
    # ======================================================
    #
    # Algunos PDFs digitales sí contienen texto, pero pdfplumber extrae
    # visualmente un número como 250.00 de esta manera:
    #
    #     "2 50.00"
    #
    # y un importe como 11,832.50 como:
    #
    #     "$ 1 1,832.50"
    #
    # Eso hacía fallar los patrones anteriores aunque el PDF NO estuviera
    # escaneado. Este fallback solo corre si los métodos anteriores no
    # encontraron partidas, por lo que no altera los formatos que ya
    # funcionaban.
    if not filas_validas:

        filas_presupuesto_obra = []

        patron_presupuesto_obra = re.compile(
            r"^(AR\s+\S+)\s+"                         # clave base
            r"(.+?)\s+"                              # descripción en la línea
            r"\b(M2|M3|ML|M|PZA|PZAS|SERVICIO|LOTE|KG|TON)\b"
            r"\s+([0-9][0-9 ,]*\.\d{2}|-)"          # cantidad (tolera espacios)
            r"\s+\$\s*([0-9][0-9 ,]*\.\d{2}|-)"     # precio unitario
            r"\s+\$\s*([0-9][0-9 ,]*\.\d{2}|-)\s*$",# importe
            flags=re.IGNORECASE,
        )

        patrones_ruido_presupuesto = [
            re.compile(r"^Descripci[oó]ndel proyecto:", re.IGNORECASE),
            re.compile(r"^Descripci[oó]n del proyecto:", re.IGNORECASE),
            re.compile(r"^Cliente:", re.IGNORECASE),
            re.compile(r"^Fecha:", re.IGNORECASE),
            re.compile(r"^PRESUPUESTO$", re.IGNORECASE),
            re.compile(
                r"^Clave\s+Descripci[oó]n\s+Unidad\s+Cantidad\s+P\s+Unitario\s+Importe$",
                re.IGNORECASE,
            ),
            # Títulos/subtotales de sección, por ejemplo:
            # "CIMENTACION $ 364,129.47"
            re.compile(
                r"^[A-ZÁÉÍÓÚÑ0-9 /().,\-]+"
                r"\s+\$\s*[0-9][0-9 ,]*\.\d{2}$",
                re.IGNORECASE,
            ),
        ]

        fila_actual = None

        def _numero_pdf_partido(valor):
            """
            Convierte números cuyo texto quedó partido por el layout del PDF.
            Ejemplos:
              '2 50.00'     -> 250.00
              '1 ,010.34'   -> 1010.34
              '1 1,832.50'  -> 11832.50
              '-'           -> None
            """
            if valor is None:
                return None

            texto = str(valor).strip()

            if texto in {"", "-", "--"}:
                return None

            texto = (
                texto.replace("$", "")
                .replace(",", "")
                .replace(" ", "")
                .strip()
            )

            try:
                return float(texto)
            except (TypeError, ValueError):
                return None

        for numero_pagina, linea in lineas_pdf:

            coincidencia = patron_presupuesto_obra.match(linea)

            if coincidencia:

                # Cierra la partida anterior después de haber acumulado
                # todas sus líneas de descripción.
                if fila_actual is not None:
                    filas_presupuesto_obra.append(fila_actual)

                cantidad = _numero_pdf_partido(coincidencia.group(4))
                precio_unitario = _numero_pdf_partido(coincidencia.group(5))
                importe = _numero_pdf_partido(coincidencia.group(6))

                fila_actual = {
                    "partida": coincidencia.group(1).strip(),
                    "concepto": coincidencia.group(2).strip(),
                    "unidad": normalizar_unidad(coincidencia.group(3)),
                    "cantidad": cantidad,
                    "precio_unitario": precio_unitario,
                    "importe": importe,
                    "origen": f"Página {numero_pagina}",
                    "fila_encabezado": None,
                    "puntaje_deteccion": 9,
                }

                paginas_detectadas.add(numero_pagina)
                continue

            if fila_actual is None:
                continue

            # Si aparece otra clave AR que no pudo cerrarse porque no trae
            # precio (por ejemplo "PZA - $ - $ -"), no debe pegarse a la
            # descripción de la partida anterior.
            if linea.upper().startswith("AR "):
                filas_presupuesto_obra.append(fila_actual)
                fila_actual = None
                continue

            if any(
                patron.search(linea)
                for patron in patrones_ruido_presupuesto
            ):
                continue

            # La descripción de estos presupuestos suele continuar debajo
            # de la línea que contiene unidad/cantidad/precio/importe.
            fila_actual["concepto"] = re.sub(
                r"\s+",
                " ",
                f"{fila_actual['concepto']} {linea}",
            ).strip()

        if fila_actual is not None:
            filas_presupuesto_obra.append(fila_actual)

        # Conserva únicamente partidas realmente cotizadas.
        filas_presupuesto_obra = [
            fila
            for fila in filas_presupuesto_obra
            if fila.get("concepto")
            and fila.get("precio_unitario") is not None
            and fila.get("precio_unitario") > 0
        ]

        if filas_presupuesto_obra:
            filas_validas = filas_presupuesto_obra


    # ======================================================
    # VALIDACIÓN Y LIMPIEZA FINAL
    # ======================================================

    if not filas_validas:
        raise ValueError(
            "No se detectaron partidas válidas en el PDF. "
            "El documento puede estar escaneado o tener "
            "una estructura diferente."
        )

    resultado = pd.DataFrame(
        filas_validas
    )

    resultado["cantidad"] = pd.to_numeric(
        resultado["cantidad"],
        errors="coerce",
    )

    resultado["precio_unitario"] = pd.to_numeric(
        resultado["precio_unitario"],
        errors="coerce",
    )

    resultado["importe"] = pd.to_numeric(
        resultado["importe"],
        errors="coerce",
    )

    importe_calculado = (
        resultado["cantidad"]
        * resultado["precio_unitario"]
    )

    resultado["importe"] = resultado["importe"].fillna(
        importe_calculado
    )

    resultado = resultado.drop_duplicates(
        subset=[
            "partida",
            "concepto",
            "unidad",
            "precio_unitario",
        ],
        keep="first",
    )

    resultado = resultado.sort_values(
        by="partida",
        key=lambda serie: pd.to_numeric(
            serie,
            errors="coerce",
        ),
    )

    resultado = resultado.reset_index(
        drop=True
    )

    metadatos = {
        "iva_en_documento": bool(re.search(r"\bI\.?\s?V\.?\s?A\b", texto_para_detectar_formato, re.I)),
        "tipo": "PDF digital",
        "hoja_detectada": (
            f"{len(paginas_detectadas)} páginas con partidas"
        ),
        "partidas": len(resultado),
        "confianza": (
            95 if len(resultado) > 0 else 0
        ),
    }

    # Verificación contra el SUBTOTAL del PDF (no el Total), reutilizando
    # el valor ya extraído arriba para decidir el método. El importe de
    # cada renglón es antes de impuestos -- el "IVA(16%)" que aparece
    # junto a cada partida es solo la tasa aplicable, no un monto ya
    # sumado al importe de esa fila -- así que la suma de todos los
    # importes debe cuadrar con el Subtotal de la cotización, NO con el
    # Total (que ya trae el IVA sumado). Comparar contra el Total aquí
    # daría una diferencia falsa del ~16% y haría parecer que la lectura
    # del PDF falló cuando en realidad está correcta.
    if subtotal_declarado:
        suma_importes = float(
            resultado["importe"].sum()
        )

        diferencia = abs(
            suma_importes - subtotal_declarado
        )

        metadatos["subtotal_declarado"] = subtotal_declarado
        metadatos["suma_importes_detectados"] = round(
            suma_importes, 2
        )
        # Tolerancia de $1: redondeos de centavos entre renglones.
        metadatos["coincide_con_subtotal"] = diferencia < 1.0

        if diferencia < 1.0:
            metadatos["confianza"] = max(
                metadatos["confianza"], 98
            )

    return resultado, metadatos

# ==========================================================
# IDENTIFICACIÓN DE FORMATO
# ==========================================================


def cargar_y_normalizar_archivo(archivo):
    extension = Path(
        archivo.name
    ).suffix.lower()

    if extension in {
        ".xlsx",
        ".xlsm",
        ".xls",
    }:
        return leer_excel(
            archivo
        )

    if extension == ".csv":
        return leer_csv(
            archivo
        )

    if extension == ".pdf":
        return leer_pdf(
            archivo
        )

    raise ValueError(
        f"El formato {extension} todavía no está soportado."
    )


@st.cache_data(show_spinner=False, max_entries=8)
def leer_archivo_en_cache(nombre, contenido):
    """Evita volver a extraer un PDF al cambiar controles de Streamlit."""
    archivo_en_memoria = io.BytesIO(contenido)
    archivo_en_memoria.name = nombre
    return cargar_y_normalizar_archivo(archivo_en_memoria)


# ==========================================================
# BARRA LATERAL
# ==========================================================

with st.sidebar:

    st.subheader(
        "Datos de esta cotización"
    )

    proveedor = st.text_input(
        "Proveedor",
        placeholder="Opcional",
    )

    proyecto = st.text_input(
        "Proyecto / licitación",
        placeholder="Opcional",
    )

    guardar_en_historico = st.checkbox(
        "Guardar esta cotización en el histórico",
        value=False,
        disabled=historico is None,
        help=(
            "Cada partida quedará guardada para "
            "comparaciones futuras."
        ),
    )

    # Los 4 filtros, en el mismo orden en que se pidió la app. Las opciones
    # técnicas quedan en "Opciones avanzadas" para no saturar la pantalla.
    st.markdown("**Los 4 filtros**")
    if historico is None:
        st.markdown("1. Histórico Ragasa — ⚠️ no conectado")
    else:
        _res_hist = historico.resumen()
        st.markdown(f"1. Histórico Ragasa — ✅ {_res_hist['total_renglones']} renglones · "
                    f"[abrir hoja](https://docs.google.com/spreadsheets/d/"
                    f"{st.secrets.get('sheet_id', DEFAULT_SHEET_ID)}/edit)")
    st.markdown("2. Nuevo León (licitaciones) — ✅")
    st.markdown("3. CDMX (tabulador 2026) — ✅")
    if busqueda_ia_disponible:
        st.markdown("4. IA en internet — ✅ " + (
            "Gemini" if busqueda_mercado_ia._obtener_cliente() is not None else "buscador"
        ))
    else:
        st.markdown("4. IA en internet — ⚠️ no conectada")

    with st.expander("Opciones avanzadas"):
        ajustar_inflacion = st.checkbox(
            "Aplicar inflación acumulada a los precios con fecha (NL e histórico)",
            value=True,
            help="Cada precio se lleva a hoy con la inflación acumulada (compuesta) desde su fecha: "
                 "factor = índice del último mes ÷ índice del mes del precio. En Nuevo León la fecha del "
                 "precio es el año de la licitación (viene en su número), no el día en que se publicó el registro.",
        )
        usar_ia = st.checkbox(
            "Revisión con IA (confirma que el concepto sea el mismo)",
            value=ia_disponible,
            disabled=not ia_disponible,
        )
        limite_busqueda_s = st.number_input(
            "Tiempo máximo total de búsqueda en internet (segundos)",
            min_value=30, max_value=1800, value=240, step=30,
            help="Límite para toda la revisión, además de 40 s por partida. Al llegar al límite se "
                 "conservan los resultados obtenidos y el resto queda como «tiempo agotado».",
        )
        actualizar_busquedas = st.checkbox(
            "Actualizar búsquedas de internet (no reutilizar las anteriores)",
            value=False,
            help="Por defecto se reutiliza una búsqueda del mismo concepto, unidad y región de los últimos "
                 "30 días; se muestra su fecha.",
        )
        st.caption(
            "Filtro 4: Gemini busca el precio con Google. Si se acaba su cuota, "
            "un buscador de respaldo abre las páginas y toma el precio solo del renglón "
            "donde aparecen el concepto, la unidad y el precio."
        )
        st.caption(
            f"{'Índice de construcción (INEGI)' if ajuste_inflacion.INDICE_ES_CONSTRUCCION else 'INPC'} "
            f"{'en vivo' if inpc_en_vivo else 'de respaldo'}: {ajuste_inflacion.ETIQUETA_ACTUAL}. "
            "La inflación se aplica acumulada (compuesta) desde la fecha de cada precio."
        )
    revisar_todo_con_ia = True
    buscar_precios_web = busqueda_ia_disponible


# ==========================================================
# INFORMACIÓN DEL FORMATO
# ==========================================================

with st.expander(
    "Formatos y estructura reconocida"
):

    st.write(
        "La aplicación detecta automáticamente la hoja, "
        "la fila de encabezados y los nombres de las columnas."
    )

    st.write(
        "Formatos admitidos:"
    )

    st.code(
        "Excel: .xlsx, .xls, .xlsm\n"
        "Datos: .csv\n"
        "Documento: .pdf digital",
        language="text",
    )

    st.write(
        "Campos internos utilizados:"
    )

    st.code(
        "partida | concepto | unidad | cantidad | "
        "precio_unitario | importe",
        language="text",
    )

    ejemplo = pd.DataFrame(
        [
            {
                "partida": "001",
                "concepto": (
                    "Suministro y colocación de acero "
                    "de refuerzo en losas"
                ),
                "unidad": "KG",
                "cantidad": 1000,
                "precio_unitario": 30,
                "importe": 30000,
            },
            {
                "partida": "002",
                "concepto": (
                    "Limpieza final de obra"
                ),
                "unidad": "M2",
                "cantidad": 500,
                "precio_unitario": 9,
                "importe": 4500,
            },
        ]
    )

    st.dataframe(
        ejemplo,
        width="stretch",
        hide_index=True,
    )


# ==========================================================
# CARGADOR DE ARCHIVOS
# ==========================================================

archivo = st.file_uploader(
    "Sube tu cotización o licitación",
    type=[
        "xlsx",
        "xls",
        "xlsm",
        "csv",
        "pdf",
    ],
    help=(
        "La aplicación buscará automáticamente "
        "la tabla de partidas."
    ),
)


# ==========================================================
# PROCESAMIENTO DEL ARCHIVO
# ==========================================================

if archivo is not None:

    try:

        with st.spinner(
            "Leyendo el archivo y detectando las partidas..."
        ):
            cotizacion, metadatos = (
                leer_archivo_en_cache(
                    archivo.name, archivo.getvalue()
                )
            )

        st.success(
            "Archivo interpretado correctamente. "
            f"Se detectaron {len(cotizacion)} partidas."
        )

        c1, c2, c3 = st.columns(
            3
        )

        c1.metric(
            "Tipo de archivo",
            metadatos.get(
                "tipo",
                "",
            ),
        )

        c2.metric(
            "Hoja o sección",
            metadatos.get(
                "hoja_detectada",
                "",
            ),
        )

        c3.metric(
            "Confianza de lectura",
            (
                f"{metadatos.get('confianza', 0)}%"
            ),
        )

        st.subheader(
            "Vista previa de las partidas detectadas"
        )

        columnas_vista = [
            "partida",
            "concepto",
            "unidad",
            "cantidad",
            "precio_unitario",
            "importe",
            "origen",
        ]

        st.dataframe(
            cotizacion[
                columnas_vista
            ].head(100),
            width="stretch",
            hide_index=True,
        )

        if metadatos.get(
            "confianza",
            0,
        ) < 80:

            st.warning(
                "La confianza de lectura es baja. "
                "Revisa la vista previa antes de continuar."
            )

        if "coincide_con_subtotal" in metadatos:

            if metadatos["coincide_con_subtotal"]:

                st.caption(
                    "✅ La suma de los renglones detectados "
                    "coincide con el Subtotal del documento "
                    f"(${metadatos['subtotal_declarado']:,.2f})."
                )

            else:

                st.warning(
                    "La suma de los renglones detectados "
                    f"(${metadatos['suma_importes_detectados']:,.2f}) "
                    "no coincide con el Subtotal del documento "
                    f"(${metadatos['subtotal_declarado']:,.2f}). "
                    "Puede faltar o sobrar alguna partida: revisa la "
                    "vista previa."
                )

        confirmar = st.checkbox(
            "Confirmo que las partidas detectadas son correctas",
            value=False,
        )

        if confirmar:

            with st.spinner(
                f"Revisando {len(cotizacion)} partidas "
                "contra NL, CDMX e histórico interno..."
            ):

                filas = []
                fila_registros = []
                partidas_sin_descripcion = 0
                revision_ia.reiniciar_error()
                busqueda_mercado_ia.reiniciar_error()

                for _, renglon in cotizacion.iterrows():

                    precio = convertir_numero(
                        renglon["precio_unitario"]
                    )

                    if precio is None or precio <= 0:
                        continue

                    concepto = str(renglon["concepto"]).strip()
                    if concepto.casefold() in ("", "none", "nan", "null"):
                        partidas_sin_descripcion += 1
                        continue
                    unidad = str(renglon["unidad"]).strip()

                    resultado = comparador.evaluar(
                        concepto,
                        unidad,
                        precio,
                        ajustar_inflacion=ajustar_inflacion,
                        usar_ia=usar_ia,
                    )

                    nl = resultado["fuentes"]["nl_historico"]
                    cdmx = resultado["fuentes"]["cdmx_gobierno"]

                    consulta_historico = None
                    if historico is not None:
                        consulta_historico = historico.consultar(
                            concepto,
                            unidad,
                            precio,
                            usar_ia=usar_ia,
                            excluir=((proveedor or "").strip() or "Sin proveedor",
                                     (proyecto or "").strip() or Path(archivo.name).stem),
                        )

                    fila_registros.append(
                        {
                            "renglon": renglon,
                            "precio": precio,
                            "concepto": concepto,
                            "unidad": unidad,
                            "nl": nl,
                            "cdmx": cdmx,
                            "consulta_historico": consulta_historico,
                        }
                    )

                if partidas_sin_descripcion:
                    st.warning(
                        f"{partidas_sin_descripcion} renglón(es) sin descripción se omitieron "
                        "del comparativo. Revisa la vista previa y corrige el archivo de origen."
                    )

                # ----------------------------------------------------------------
                # Revisión con IA en LOTE (no una llamada por partida): se juntan
                # todos los matches riesgosos (confianza BAJA o precio con
                # diferencia extrema, marcados con 'motivo' por comparador.evaluar()
                # / historico.consultar()) de TODA la cotización y se mandan en unos
                # pocos lotes de revision_ia.TAMANO_LOTE elementos. Esto es lo que
                # evita que revisar 20-25 partidas dudosas tome 2-4 minutos y se
                # agote el límite de solicitudes por minuto de la cuenta gratuita:
                # en vez de ~20 llamadas (una por partida), quedan ~3-4 (una por
                # lote).
                #
                # OJO: este bloque corre ANTES de decidir qué partidas mandar a la
                # 4ª fuente (búsqueda en internet) -- a propósito. Si una partida
                # tenía match en NL/CDMX/histórico pero la IA lo RECHAZA aquí, esa
                # partida se queda sin ninguna referencia válida, y debe poder
                # mandarse a buscar en internet igual que una partida que nunca
                # tuvo match. Si este bloque corriera después, esas partidas se
                # quedarían sin ninguna fuente de precio (bug ya visto en pruebas
                # reales: una partida con match rechazado por la IA se marcaba
                # "no hacía falta: ya hay precio de referencia de otra fuente"
                # cuando en realidad esa referencia ya había sido descartada).
                # ----------------------------------------------------------------
                if usar_ia:

                    items_revision = []

                    for indice, registro in enumerate(fila_registros):

                        fuentes_a_revisar = [
                            (
                                "nl",
                                registro["nl"],
                                "historico de Nuevo Leon",
                            ),
                            (
                                "cdmx",
                                registro["cdmx"],
                                "Tabulador CDMX (gobierno)",
                            ),
                        ]

                        if registro["consulta_historico"] is not None:
                            fuentes_a_revisar.append(
                                (
                                    "hist",
                                    registro["consulta_historico"],
                                    "historico interno guardado en Google Sheets",
                                )
                            )

                        for clave_fuente, fuente_dict, nombre_fuente in fuentes_a_revisar:
                            # Por default solo se manda a revisión la
                            # coincidencia si ya venía marcada como dudosa
                            # (confianza BAJA o diferencia de precio
                            # extrema). Con "revisar_todo_con_ia" activado,
                            # se manda CUALQUIER coincidencia con match,
                            # sin importar su confianza -- para que la IA
                            # confirme también las que ya parecían seguras.
                            # BAJA y MEDIA siempre se revisan: sin CONFIRMA de la
                            # IA no pueden quedar VALIDADAS (validacion_referencias).
                            if fuente_dict.get("match") and (
                                fuente_dict.get("motivo")
                                or str(fuente_dict.get("confianza", "")).upper() in ("BAJA", "MEDIA")
                                or revisar_todo_con_ia
                            ):
                                items_revision.append(
                                    {
                                        "id": f"{indice}:{clave_fuente}",
                                        "descripcion_cotizada": registro["concepto"],
                                        "unidad": registro["unidad"],
                                        "descripcion_candidato": fuente_dict["match"],
                                        "fuente": nombre_fuente,
                                    }
                                )

                    if items_revision:

                        resultados_revision = {}

                        # Ahorro de cuota: un veredicto ya obtenido para el mismo
                        # concepto + candidato + fuente se reutiliza en esta
                        # sesión en vez de volver a preguntarle a Gemini.
                        cache_revision = st.session_state.setdefault("veredictos_ia", {})

                        def _clave_revision(item):
                            return (
                                str(item["descripcion_cotizada"]).casefold().strip(),
                                str(item["unidad"]).casefold().strip(),
                                str(item["descripcion_candidato"]).casefold().strip(),
                                item["fuente"],
                            )

                        pendientes_revision = []
                        for item in items_revision:
                            guardado = cache_revision.get(_clave_revision(item))
                            if guardado:
                                resultados_revision[item["id"]] = guardado
                            else:
                                pendientes_revision.append(item)

                        for inicio in range(
                            0, len(pendientes_revision), revision_ia.TAMANO_LOTE
                        ):
                            lote = pendientes_revision[
                                inicio:inicio + revision_ia.TAMANO_LOTE
                            ]
                            respuesta_revision = revision_ia.revisar_coincidencias_debiles_lote(lote)
                            resultados_revision.update(respuesta_revision)
                            for item in lote:
                                if item["id"] in respuesta_revision:
                                    cache_revision[_clave_revision(item)] = respuesta_revision[item["id"]]

                        for item in items_revision:
                            veredicto = resultados_revision.get(item["id"])
                            if not veredicto:
                                continue
                            indice_texto, clave_fuente = item["id"].split(":", 1)
                            registro = fila_registros[int(indice_texto)]
                            if clave_fuente == "nl":
                                registro["nl"]["revision_ia"] = veredicto
                            elif clave_fuente == "cdmx":
                                registro["cdmx"]["revision_ia"] = veredicto
                            elif clave_fuente == "hist":
                                registro["consulta_historico"]["revision_ia"] = veredicto

                # ----------------------------------------------------------------
                # Una fuente cuenta como "referencia válida" solo si tiene
                # clasificación Y la IA no la rechazó (si la IA no está activada,
                # o simplemente no la revisó, sigue contando con su clasificación
                # original). Esto es lo que decide tanto si hace falta mandar la
                # partida a la 4ª fuente como el mensaje que se muestra después.
                # ----------------------------------------------------------------
                def _referencia_valida(fuente_dict):
                    if not fuente_dict or not fuente_dict.get("clasificacion"):
                        return False
                    descartar, _motivo = _revision_ia_descarta(fuente_dict, usar_ia)
                    return not descartar

                # ----------------------------------------------------------------
                # 4ª fuente: la IA busca en internet (Google Search vía Gemini, o
                # Tavily como respaldo gratuito) un precio de mercado para cada
                # partida.
                #
                # La fuente IA se consulta para cada partida, aun si otra base
                # encontró una referencia. La búsqueda externa cuesta cuota y
                # puede ser la parte más lenta de la revisión.
                # ----------------------------------------------------------------
                if busqueda_ia_disponible and buscar_precios_web:

                    items_busqueda_mercado = [
                        {
                            "id": str(indice),
                            "descripcion": registro["concepto"],
                            "unidad": registro["unidad"],
                            "precio": registro["precio"],
                        }
                        for indice, registro in enumerate(fila_registros)
                    ]

                    import time as _time
                    import busqueda_web_respaldo as _bwr
                    resultados_busqueda_mercado = {}
                    pendientes = []
                    for item in items_busqueda_mercado:
                        guardada = None if actualizar_busquedas else busqueda_mercado_ia.busqueda_guardada(item)
                        if guardada:
                            res_g, fecha_g = guardada
                            res_g = dict(res_g)
                            res_g["reutilizada"] = fecha_g.strftime("%Y-%m-%d %H:%M")
                            resultados_busqueda_mercado[item["id"]] = res_g
                        else:
                            pendientes.append(item)

                    _t0 = _time.monotonic()
                    _deadline = _t0 + float(limite_busqueda_s)
                    progreso_busqueda = st.progress(
                        0, text="Preparando búsqueda de precios externos..."
                    )
                    _total_items = max(len(items_busqueda_mercado), 1)
                    _hechas = len(items_busqueda_mercado) - len(pendientes)
                    for inicio in range(0, len(pendientes), busqueda_mercado_ia.TAMANO_LOTE):
                        lote = pendientes[inicio:inicio + busqueda_mercado_ia.TAMANO_LOTE]
                        if _time.monotonic() >= _deadline:
                            # Límite total alcanzado: se conservan los resultados ya
                            # obtenidos y el resto queda marcado como tiempo agotado.
                            for item in pendientes[inicio:]:
                                resultados_busqueda_mercado[item["id"]] = {
                                    "precio_mxn": None, "tiene_dato": False, "estado_busqueda": "tiempo agotado",
                                    "nota": (f"Tiempo agotado: se alcanzó el límite total de {limite_busqueda_s} s "
                                             "de búsqueda en esta revisión antes de buscar esta partida. No indica "
                                             "que no existan precios publicados; sube el límite en Opciones "
                                             "avanzadas o vuelve a revisar (se reutiliza lo ya buscado)."),
                                }
                            break
                        respuesta_lote = busqueda_mercado_ia.buscar_precios_mercado_lote(lote, deadline=_deadline)
                        _err = (busqueda_mercado_ia.ultimo_error() or {}).get("mensaje") or ""
                        for item in lote:
                            respuesta = respuesta_lote.get(item["id"])
                            if not respuesta:
                                estado_e = _bwr.clasificar_error(_err) if _err else "error de conexión"
                                respuesta = {
                                    "precio_mxn": None, "tiene_dato": False, "estado_busqueda": estado_e,
                                    "nota": (f"{estado_e.capitalize()}: el buscador no respondió para esta partida"
                                             + (f" ({_err[:160]})" if _err else "")
                                             + ". Fallo técnico; no indica que no existan precios publicados."),
                                }
                            respuesta["fecha_busqueda"] = pd.Timestamp.now(tz="America/Monterrey").strftime("%Y-%m-%d %H:%M")
                            resultados_busqueda_mercado[item["id"]] = respuesta
                            busqueda_mercado_ia.guardar_busqueda(item, respuesta)
                        _hechas += len(lote)
                        progreso_busqueda.progress(
                            min(_hechas / _total_items, 1.0),
                            text=(f"Búsqueda externa: {_hechas}/{len(items_busqueda_mercado)} partidas · "
                                  f"{_time.monotonic() - _t0:.0f} de {limite_busqueda_s} s"),
                        )

                    progreso_busqueda.empty()

                    for indice, registro in enumerate(fila_registros):
                        registro["busqueda_ia"] = resultados_busqueda_mercado.get(
                            str(indice)
                        )

                else:

                    for registro in fila_registros:
                        registro["busqueda_ia"] = None

                # ----------------------------------------------------------------
                # Segunda pasada: ya con la revisión de IA (si aplica) resuelta
                # para todas las partidas, se arma el resultado de cada renglón
                # exactamente igual que antes.
                # ----------------------------------------------------------------
                items_sin_datos = []

                for registro in fila_registros:

                    renglon = registro["renglon"]
                    precio = registro["precio"]
                    concepto = registro["concepto"]
                    unidad = registro["unidad"]
                    nl = registro["nl"]
                    cdmx = registro["cdmx"]
                    consulta_historico = registro["consulta_historico"]

                    # ------------------------------------------------------------
                    # Datos básicos + columnas de DETALLE de cada fuente (match
                    # de texto encontrado, confianza, precio de referencia).
                    # OJO: los 4 "Resultado" + "% Diferencia" de cada fuente NO
                    # se agregan aquí -- se calculan en variables locales y se
                    # agregan TODOS JUNTOS al final de este bloque, para que
                    # queden pegados unos con otros en las últimas columnas de
                    # la tabla en vez de repartidos entre el detalle de cada
                    # fuente -- pedido explícito: "que el resultado final de
                    # cada uno estén al final juntas pegadas".
                    # ------------------------------------------------------------
                    fila = {
                        "Partida": renglon.get("partida"),
                        "Concepto": concepto,
                        "Unidad": unidad,
                        "Cantidad": renglon.get("cantidad"),
                        "Precio cotizado": precio,
                        "Importe": renglon.get("importe"),
                        "Origen": renglon.get("origen"),
                        "Match NL": nl.get("match"),
                        "Confiabilidad NL": nl.get("confianza"),
                        "Año del dato NL": nl.get("anio_dato_mas_reciente"),
                        "Precio mediana NL (original)": nl.get("precio_mediana"),
                        "Precio mediana NL (ajustado hoy)": nl.get(
                            "precio_mediana_ajustada"
                        ),
                        "Match CDMX": cdmx.get("match"),
                        "Confiabilidad CDMX": cdmx.get("confianza"),
                        "Precio referencia CDMX": cdmx.get("precio_referencia"),
                    }

                    _ref_nl = nl.get("precio_mediana_ajustada")
                    _diff_nl = (
                        round((precio - _ref_nl) / _ref_nl * 100, 1) if _ref_nl else None
                    )
                    _ref_cdmx = cdmx.get("precio_referencia")
                    _diff_cdmx = (
                        round((precio - _ref_cdmx) / _ref_cdmx * 100, 1) if _ref_cdmx else None
                    )

                    # ------------------------------------------------------------
                    # 4ª fuente: precio de mercado encontrado por la IA en
                    # internet (búsqueda real con Google Search, no opinión de
                    # memoria). Si no encontró un precio verificable, o esta
                    # fuente no está configurada, se deja explícito "todavía no
                    # hay" en vez de dejar la columna en blanco sin explicación.
                    # ------------------------------------------------------------
                    busqueda_ia = registro.get("busqueda_ia")
                    clasificacion_ia_mercado = None
                    _diff_ia = None
                    _resultado_ia = "todavía no hay"

                    if (busqueda_ia and busqueda_ia.get("tiene_dato")
                            and convertir_numero(busqueda_ia.get("precio_mxn"))
                            and convertir_numero(busqueda_ia.get("precio_mxn")) > 0):

                        precio_mercado_ia = convertir_numero(busqueda_ia["precio_mxn"])
                        es_referencia_web = not (
                            str(busqueda_ia.get("motor", "")).startswith("Gemini")
                            and busqueda_ia.get("verificado")
                        )
                        # Prueba real (barda): Tavily trajo $9,500/m² para un
                        # muro de $550 y $218,693 por un castillo de $400 --
                        # números de otra partida o de un total dentro del
                        # fragmento. Un precio web a más de 3× (o menos de
                        # 1/3) del cotizado casi siempre es otra escala; con
                        # Gemini el margen es 5×.
                        _factor_escala = 3.0 if es_referencia_web else 5.0
                        _fuera_de_escala = not (
                            precio / _factor_escala <= precio_mercado_ia <= precio * _factor_escala
                        )
                        if _fuera_de_escala:
                            busqueda_ia = dict(busqueda_ia)
                            busqueda_ia["nota"] = (
                                f"Se descartó ${precio_mercado_ia:,.2f}: fuera de escala "
                                "frente al precio cotizado (probablemente otra unidad o un total)."
                            )
                            busqueda_ia["tiene_dato"] = False
                            _resultado_ia = "todavía no hay"
                        else:
                            banda_baja_ia, banda_alta_ia = banda_en_mercado(precio_mercado_ia)
                            _clasificacion_ia = clasificar(precio, banda_baja_ia, banda_alta_ia)
                            _diff_ia = round(
                                (precio - precio_mercado_ia) / precio_mercado_ia * 100, 1
                            )
                            # La columna IA se pinta con su semáforo. Solo un
                            # precio de Gemini (fuente y concepto validados)
                            # vota en el resultado final; un fragmento web de
                            # Tavily se muestra como referencia, sin votar.
                            _resultado_ia = _clasificacion_ia
                            if not es_referencia_web:
                                clasificacion_ia_mercado = _clasificacion_ia

                        fila["Motor IA"] = busqueda_ia.get("motor", "")
                        fila["Precio mercado (IA internet)"] = (
                            None if _fuera_de_escala else precio_mercado_ia
                        )
                        nombre_fuente = busqueda_ia.get("fuente_nombre") or ""
                        url_fuente = busqueda_ia.get("fuente_url") or ""
                        fila["Fuente IA (internet)"] = (
                            f"{nombre_fuente} ({url_fuente})" if nombre_fuente and url_fuente
                            else nombre_fuente or url_fuente
                        )
                        fila["Nota IA internet"] = busqueda_ia.get("nota", "")

                    elif busqueda_ia_disponible and not buscar_precios_web:

                        fila["Nota IA internet"] = "Búsqueda externa desactivada"

                    elif busqueda_ia_disponible:

                        fila["Nota IA internet"] = (
                            busqueda_ia.get("nota", "") if busqueda_ia else
                            "la IA no encontró un precio real verificable para esta partida"
                        )

                    else:

                        fila["Nota IA internet"] = (
                            "4ª fuente no conectada (falta gemini_api_key o "
                            "tavily_api_key en Secrets)"
                        )

                    revision_ia_nl = nl.get("revision_ia")
                    if revision_ia_nl:
                        fila["Revisión IA (match NL débil)"] = (
                            f"{revision_ia_nl['veredicto']}: "
                            f"{revision_ia_nl['razon']}"
                        )

                    revision_ia_cdmx = cdmx.get("revision_ia")
                    if revision_ia_cdmx:
                        fila["Revisión IA (match CDMX débil)"] = (
                            f"{revision_ia_cdmx['veredicto']}: "
                            f"{revision_ia_cdmx['razon']}"
                        )

                    # Si el match quedó marcado como riesgoso (confianza BAJA o
                    # precio con diferencia extrema) y la IA lo rechazó, no
                    # estuvo segura, O ni siquiera se pudo completar la revisión
                    # (ej. límite de la cuenta gratuita), no debe contar para el
                    # resultado final -- es más seguro tratarlo como no
                    # confirmado que confiar en un match que nunca se validó. Se
                    # deja visible en la columna final "Resultado NL"/"Resultado
                    # CDMX" (agregada más abajo) para que quede claro qué se
                    # descartó y por qué.
                    nl_rechazado_por_ia, motivo_descarte_nl = (
                        _revision_ia_descarta(nl, usar_ia)
                    )
                    cdmx_rechazado_por_ia, motivo_descarte_cdmx = (
                        _revision_ia_descarta(cdmx, usar_ia)
                    )
                    if _alcance_distinto(concepto, nl.get("match")):
                        nl_rechazado_por_ia, motivo_descarte_nl = True, "alcance distinto"
                    if _alcance_distinto(concepto, cdmx.get("match")):
                        cdmx_rechazado_por_ia, motivo_descarte_cdmx = True, "alcance distinto"

                    _resultado_nl = nl.get("clasificacion")
                    if nl_rechazado_por_ia and _resultado_nl:
                        _resultado_nl = f"{_resultado_nl} (descartado: {motivo_descarte_nl})"

                    _resultado_cdmx = cdmx.get("clasificacion")
                    if cdmx_rechazado_por_ia and _resultado_cdmx:
                        _resultado_cdmx = f"{_resultado_cdmx} (descartado: {motivo_descarte_cdmx})"

                    nl_cuenta = not nl_rechazado_por_ia and not _baja_sin_confirmar(nl)
                    cdmx_cuenta = not cdmx_rechazado_por_ia and not _baja_sin_confirmar(cdmx)
                    if nl.get("match") and _baja_sin_confirmar(nl) and not nl_rechazado_por_ia and _resultado_nl:
                        _resultado_nl = f"{_resultado_nl} (por confirmar)"
                    if cdmx.get("match") and _baja_sin_confirmar(cdmx) and not cdmx_rechazado_por_ia and _resultado_cdmx:
                        _resultado_cdmx = f"{_resultado_cdmx} (por confirmar)"

                    clasificaciones = [
                        valor
                        for valor in (
                            nl.get("clasificacion") if nl_cuenta else None,
                            cdmx.get("clasificacion") if cdmx_cuenta else None,
                        )
                        if valor
                    ]

                    # Precio(s) de referencia de las mismas fuentes que sí
                    # cuentan para el veredicto (se excluyen las que la IA
                    # rechazó) -- sirve para calcular el % de diferencia de la
                    # partida contra el mercado.
                    referencias_precio = [
                        valor
                        for valor in (
                            nl.get("precio_mediana_ajustada") if nl_cuenta else None,
                            cdmx.get("precio_referencia") if cdmx_cuenta else None,
                        )
                        if valor
                    ]

                    _resultado_historico = "todavía no hay"
                    _diff_historico = None

                    if consulta_historico is not None and consulta_historico.get(
                        "match"
                    ):

                        fila["Match histórico interno"] = consulta_historico["match"]

                        fila["Confiabilidad histórico interno"] = consulta_historico.get(
                            "confianza"
                        )

                        fila["Proveedores en histórico"] = ", ".join(
                            consulta_historico.get("proveedores", [])
                        )

                        fila["Precio mediana histórico"] = consulta_historico.get(
                            "precio_mediana"
                        )

                        veredicto_historico = consulta_historico.get("clasificacion")

                        revision_ia_historico = consulta_historico.get("revision_ia")

                        if revision_ia_historico:
                            fila["Revisión IA (match histórico débil)"] = (
                                f"{revision_ia_historico['veredicto']}: "
                                f"{revision_ia_historico['razon']}"
                            )

                        historico_rechazado_por_ia, motivo_descarte_historico = (
                            _revision_ia_descarta(consulta_historico, usar_ia)
                        )

                        if historico_rechazado_por_ia and veredicto_historico:
                            veredicto_historico = (
                                f"{veredicto_historico} "
                                f"(descartado: {motivo_descarte_historico})"
                            )

                        if veredicto_historico:
                            _resultado_historico = veredicto_historico

                        _ref_hist = consulta_historico.get("precio_mediana")
                        _diff_historico = (
                            round((precio - _ref_hist) / _ref_hist * 100, 1) if _ref_hist else None
                        )

                        if (
                            consulta_historico.get("clasificacion")
                            and not historico_rechazado_por_ia
                            and not _baja_sin_confirmar(consulta_historico)
                        ):

                            clasificaciones.append(
                                consulta_historico["clasificacion"]
                            )

                            if consulta_historico.get("precio_mediana"):
                                referencias_precio.append(
                                    consulta_historico["precio_mediana"]
                                )

                    else:

                        # Sin match en el histórico interno (todavía no hay
                        # suficientes datos guardados para este concepto, o el
                        # histórico ni siquiera está conectado): se deja
                        # explícito "todavía no hay" en vez de una columna en
                        # blanco sin explicación -- pedido directo de dirección.
                        fila["Match histórico interno"] = "todavía no hay"

                    # La 4ª fuente (IA búsqueda en internet) suma su voto al
                    # veredicto combinado igual que NL/CDMX/histórico, solo
                    # cuando de verdad encontró un precio real y verificable
                    # (nunca cuando dice "todavía no hay").
                    if clasificacion_ia_mercado:
                        clasificaciones.append(clasificacion_ia_mercado)
                        if busqueda_ia and busqueda_ia.get("precio_mxn"):
                            referencias_precio.append(convertir_numero(busqueda_ia["precio_mxn"]))

                    # ------------------------------------------------------------
                    # Los 4 resultados finales (uno por fuente) JUNTOS y PEGADOS
                    # al final, cada uno con su % de diferencia justo al lado --
                    # pedido explícito: en vez de repartir cada resultado entre
                    # las columnas de detalle de su propia fuente, se agrupan
                    # todos aquí para poder comparar los 4 de un vistazo.
                    # ------------------------------------------------------------
                    fila["Resultado NL"] = _resultado_nl
                    fila["% Diferencia NL"] = _diff_nl
                    fila["Resultado CDMX"] = _resultado_cdmx
                    fila["% Diferencia CDMX"] = _diff_cdmx
                    fila["Resultado histórico"] = _resultado_historico
                    fila["% Diferencia histórico"] = _diff_historico
                    fila["Resultado IA internet"] = _resultado_ia
                    fila["% Diferencia IA internet"] = _diff_ia

                    if clasificaciones:

                        conteo = {
                            clasificacion: clasificaciones.count(clasificacion)
                            for clasificacion in set(clasificaciones)
                        }

                        _maximo = max(conteo.values())
                        _empatados = [c for c, n in conteo.items() if n == _maximo]
                        if len(_empatados) == 1 or not referencias_precio:
                            fila["RESULTADO FINAL"] = _empatados[0]
                        else:
                            # Empate entre fuentes (ej. NL dice ALTO y CDMX
                            # EN MERCADO): se desempata contra el promedio de
                            # las referencias válidas, con el mismo ±5%.
                            _promedio_empate = sum(referencias_precio) / len(referencias_precio)
                            fila["RESULTADO FINAL"] = clasificar(
                                precio, *banda_en_mercado(_promedio_empate)
                            )

                    else:

                        fila["RESULTADO FINAL"] = "SIN DATOS SUFICIENTES"

                        if usar_ia:
                            items_sin_datos.append(
                                {
                                    "id": str(len(filas)),
                                    "descripcion_cotizada": concepto,
                                    "unidad": unidad,
                                    "precio_cotizado": precio,
                                }
                            )

                    # % de diferencia del precio cotizado contra el promedio de
                    # los precios de referencia que sí contaron para el
                    # veredicto (positivo = más caro que el mercado, negativo =
                    # más barato). Si no hubo ninguna fuente confiable, se deja
                    # en blanco -- no hay con qué comparar.
                    # ------------------------------------------------------------
                    # Precio sugerido para negociar + ahorro potencial: pensado
                    # para que quien negocie con el proveedor tenga, de un
                    # vistazo, contra qué precio comparar y cuánto se podría
                    # ahorrar si la partida está cara -- pedido explícito para
                    # que la tabla sirva de verdad para negociar, no solo para
                    # diagnosticar.
                    # ------------------------------------------------------------
                    if referencias_precio:

                        precio_referencia_promedio = (
                            sum(referencias_precio) / len(referencias_precio)
                        )

                        if precio_referencia_promedio:

                            fila["% Diferencia vs referencia"] = round(
                                (precio - precio_referencia_promedio)
                                / precio_referencia_promedio
                                * 100,
                                1,
                            )
                            fila["Precio sugerido (negociación)"] = round(
                                precio_referencia_promedio, 2
                            )

                            _cantidad = renglon.get("cantidad")
                            _diferencia_unitaria = precio - precio_referencia_promedio
                            if _cantidad and _diferencia_unitaria > 0:
                                fila["Ahorro potencial"] = round(
                                    _diferencia_unitaria * _cantidad, 2
                                )
                            else:
                                fila["Ahorro potencial"] = 0.0

                        else:

                            fila["% Diferencia vs referencia"] = None

                    else:

                        fila["% Diferencia vs referencia"] = None

                    # ------------------------------------------------------------
                    # VALIDACIÓN POR FUENTE (validacion_referencias.py): aquí la
                    # revisión con IA CONTROLA los cálculos. Cada fuente recibe
                    # un estado; solo las VALIDADAS deciden el semáforo final
                    # y el precio de negociación, sin promediar fuentes.
                    # ------------------------------------------------------------
                    _pseudo_ia = None
                    if busqueda_ia and busqueda_ia.get("tiene_dato") and convertir_numero(busqueda_ia.get("precio_mxn")):
                        _pseudo_ia = {
                            "match": busqueda_ia.get("descripcion_encontrada")
                            or busqueda_ia.get("fuente_nombre") or "precio web",
                            "confianza": busqueda_ia.get("confiabilidad") or (
                                "ALTA" if busqueda_ia.get("verificado") else "BAJA"),
                            "motivo": None,
                        }
                    elif busqueda_ia:
                        _pseudo_ia = {"match": None, "motivo": busqueda_ia.get("nota")}
                    _es_gemini = bool(busqueda_ia) and str(busqueda_ia.get("motor", "")).startswith("Gemini")
                    evaluaciones = {
                        "historico": validacion.evaluar_fuente(
                            "historico", consulta_historico, concepto=concepto, precio=precio, unidad=unidad,
                            usar_ia=usar_ia,
                            precio_referencia=(
                                (consulta_historico or {}).get("precio_mediana_actualizada")
                                if ajustar_inflacion and (consulta_historico or {}).get("precio_mediana_actualizada")
                                else (consulta_historico or {}).get("precio_mediana")),
                            fecha_dato=(consulta_historico or {}).get("fecha_dato"),
                            region="RAGASA",
                        ),
                        "nl": validacion.evaluar_fuente(
                            "nl", nl, concepto=concepto, precio=precio, usar_ia=usar_ia, unidad=unidad,
                            precio_referencia=nl.get("precio_mediana_ajustada"),
                            fecha_dato=nl.get("anio_dato_mas_reciente"),
                            region="NL",
                        ),
                        "cdmx": validacion.evaluar_fuente(
                            "cdmx", cdmx, concepto=concepto, precio=precio, usar_ia=usar_ia, unidad=unidad,
                            precio_referencia=cdmx.get("precio_referencia"),
                            fecha_dato=FECHA_TABULADOR_CDMX,
                            region="CDMX",
                        ),
                        "ia": validacion.evaluar_fuente(
                            "ia", _pseudo_ia, concepto=concepto, precio=precio, usar_ia=False, unidad=unidad,
                            precio_referencia=convertir_numero((busqueda_ia or {}).get("precio_mxn"))
                            if _pseudo_ia and _pseudo_ia.get("match") else None,
                            # La página web casi nunca dice de cuándo es el
                            # precio: no se inventa una fecha; se registra
                            # solo la fecha en que se consultó.
                            fecha_dato=(busqueda_ia or {}).get("fecha_fuente"),
                            fecha_consulta=((busqueda_ia or {}).get("reutilizada")
                                            or (busqueda_ia or {}).get("fecha_busqueda")
                                            or pd.Timestamp.now(tz="America/Monterrey").strftime("%Y-%m-%d %H:%M")),
                            region="WEB",
                            es_web=True,
                            web_verificada=_es_gemini and bool((busqueda_ia or {}).get("verificado")),
                        ),
                    }
                    # Evidencia verificable de cada referencia (documento, código,
                    # página, unidad, región, fechas, ajuste por inflación).
                    _ev_hist = evaluaciones["historico"]
                    _ev_hist.update({
                        "documento": "Histórico interno Ragasa (Google Sheets)" if historico is not None else "",
                        "unidad_ref": unidad, "region": "Ragasa",
                        "proveedores_ref": ", ".join((consulta_historico or {}).get("proveedores") or []),
                    })
                    _regs_h = (consulta_historico or {}).get("registros_inflacion") or []
                    if _regs_h and _ev_hist.get("precio_referencia"):
                        _ev_hist["precio_original"] = (consulta_historico or {}).get("precio_mediana")
                        _ev_hist["registros_historico"] = _regs_h
                        if ajustar_inflacion and consulta_historico.get("precio_mediana"):
                            _ev_hist["inflacion"] = _acumulada({
                                "periodo_base": "mes en que se guardó cada compra",
                                "valor_base": None,
                                "periodo_final": ajuste_inflacion.ETIQUETA_ACTUAL,
                                "valor_final": ajuste_inflacion.NIVEL_ACTUAL,
                                "factor": round(consulta_historico["precio_mediana_actualizada"]
                                                / consulta_historico["precio_mediana"], 4),
                                "justificacion": (
                                    f"Mediana de {len(_regs_h)} compra(s) del histórico; cada una se actualizó con la "
                                    "inflación acumulada desde el mes en que se guardó."),
                            }, None, _regs_h)
                    if historico is None:
                        _ev_hist["motivo"] = "histórico no conectado"
                    elif _ev_hist["estado"] == validacion.SIN_DATO:
                        _ev_hist["motivo"] = "conectado, sin coincidencias: " + str(_ev_hist.get("motivo") or "")
                    _anio_nl = nl.get("anio_dato_mas_reciente")
                    evaluaciones["nl"].update({
                        "documento": ("SIASI, Secretaría de Movilidad y Planeación Urbana de NL, publicado en formato OCDS "
                                      "(fuente primaria: si.nl.gob.mx/transparencia/publicaciones). En la app: "
                                      "Base_Precios_Unitarios_NL_CDMX.xlsx, hoja 'Tabulador Homologado NL' "
                                      f"(mediana de las medianas de los {nl.get('n_registros') or '—'} contrato(s) "
                                      f"más recientes de {nl.get('n_contratos_total') or nl.get('n_registros') or '—'}, "
                                      f"{nl.get('n_renglones') or nl.get('n_registros') or '—'} renglones; detalle en "
                                      "'Precios Contratados (real)')"),
                        "url": "https://data.open-contracting.org/en/publication/32",
                        "unidad_ref": nl.get("unidad"), "region": "Nuevo León",
                        "precio_original": nl.get("precio_mediana"),
                        "registros": nl.get("n_registros"),
                        "periodo": ((f"licitaciones de {nl['periodo_precio_min'][:4]} a {nl['periodo_precio_max'][:4]}; "
                                     if nl.get("periodo_precio_min") else "")
                                    + f"registros publicados {str(nl.get('fecha_min'))[:7]} a {str(nl.get('fecha_max'))[:7]}"
                                    if nl.get("fecha_min") else None),
                        "inflacion": (_inflacion_nl(nl) if (_anio_nl and ajustar_inflacion) else None),
                        "tipo_registros": nl.get("tipo_registros"),
                        "registros_detalle": nl.get("registros_detalle"),
                        "registros_todos": nl.get("registros_todos") or [],
                        "n_renglones": nl.get("n_renglones"),
                    })
                    evaluaciones["cdmx"].update({
                        "documento": ("Tabulador General de Precios Unitarios del Gobierno de la CDMX, edición 2026 "
                                      "(actualización mayo 2026), Secretaría de Obras y Servicios"),
                        "url": "https://www.obras.cdmx.gob.mx/normas-tabulador/tabulador-general-de-precios-unitarios",
                        "codigo": cdmx.get("clave"), "pagina": cdmx.get("pagina"),
                        "unidad_ref": cdmx.get("unidad"), "region": "CDMX",
                    })
                    if busqueda_ia:
                        evaluaciones["ia"].update({
                            "url": busqueda_ia.get("fuente_url"), "origen": busqueda_ia.get("origen"),
                            "codigo": busqueda_ia.get("codigo"), "pagina": busqueda_ia.get("pagina"),
                            "rechazadas": busqueda_ia.get("rechazadas") or [],
                            "estado_busqueda": busqueda_ia.get("estado_busqueda"),
                            "reutilizada": busqueda_ia.get("reutilizada"),
                            "documento": busqueda_ia.get("fuente_nombre"), "unidad_ref": busqueda_ia.get("unidad_encontrada"),
                            "region": busqueda_ia.get("region") or "México (web)",
                        })
                    if evaluaciones["ia"]["estado"] != validacion.SIN_DATO and busqueda_ia:
                        # Concepto encontrado y fuente (sitio + URL) por separado.
                        evaluaciones["ia"]["fuente_web"] = " · ".join(
                            x for x in (
                                busqueda_ia.get("fuente_nombre"),
                                busqueda_ia.get("fuente_url"),
                                busqueda_ia.get("motor"),
                            ) if x
                        )
                        evaluaciones["ia"]["evidencia"] = busqueda_ia.get("fragmento") or busqueda_ia.get("nota")
                    # Coherencia: se esperaría NL igual o más caro que CDMX.
                    _coh = recomendacion.coherencia_nl_cdmx(evaluaciones)
                    if _coh:
                        evaluaciones["nl"]["coherencia"] = _coh
                    final = validacion.resultado_final(
                        evaluaciones, precio, renglon.get("cantidad")
                    )
                    fila["_evaluaciones"] = evaluaciones
                    fila["_final"] = final
                    fila["RESULTADO FINAL"] = final["semaforo"]
                    fila["% Diferencia vs referencia"] = final["diferencia_pct"]
                    fila["Precio sugerido (negociación)"] = final["precio_negociacion"]
                    fila["Ahorro potencial"] = final["ahorro_potencial"]

                    filas.append(fila)

                # ----------------------------------------------------------------
                # Opinión de IA en LOTE para las partidas que quedaron totalmente
                # "SIN DATOS SUFICIENTES" (sin ningún match en ninguna fuente, o
                # con todos sus matches descartados por la revisión de arriba).
                # Igual que la revisión de matches débiles, se manda en unos
                # pocos lotes en vez de una llamada por partida.
                # ----------------------------------------------------------------
                if usar_ia and items_sin_datos:

                    opiniones = {}

                    for inicio in range(
                        0, len(items_sin_datos), revision_ia.TAMANO_LOTE
                    ):
                        lote = items_sin_datos[inicio:inicio + revision_ia.TAMANO_LOTE]
                        opiniones.update(
                            revision_ia.opinar_sin_datos_lote(lote)
                        )

                    for item in items_sin_datos:
                        opinion_ia = opiniones.get(item["id"])
                        if not opinion_ia:
                            continue
                        indice_fila = int(item["id"])
                        filas[indice_fila]["Opinión IA (sin datos verificados)"] = (
                            f"{opinion_ia['opinion']}: {opinion_ia['razon']}"
                        )


                tabla = pd.DataFrame(
                    filas
                )

                # Se arma desde "tabla" (no desde "cotizacion" cruda) para
                # que el histórico también guarde el RESULTADO FINAL y el
                # % de diferencia que ya se calcularon -- así, más
                # adelante, se puede ver qué tan alto o bajo llegó a estar
                # cotizado un concepto en su momento, no solo el precio
                # plano.
                cotizacion_historico = tabla[
                    [
                        "Concepto",
                        "Unidad",
                        "Precio cotizado",
                        "RESULTADO FINAL",
                        "% Diferencia vs referencia",
                    ]
                ].rename(
                    columns={
                        "Concepto": "concepto",
                        "Unidad": "unidad",
                        "Precio cotizado": "precio_unitario",
                        "RESULTADO FINAL": "resultado_final",
                        "% Diferencia vs referencia": "diferencia_pct",
                    }
                ).copy()

                if (
                    historico is not None
                    and guardar_en_historico
                ):

                    # Cotización única: proveedor y proyecto son opcionales. Si
                    # faltan, se guarda como "Sin proveedor" y el nombre del
                    # archivo, para poder identificarla y no duplicarla.
                    _prov_h = (proveedor or "").strip() or "Sin proveedor"
                    _proy_h = (proyecto or "").strip() or Path(archivo.name).stem
                    _nuevas = historico.ingerir(
                        cotizacion_historico,
                        proveedor=_prov_h,
                        proyecto=_proy_h,
                    )
                    st.success(
                        f"Cotización guardada en el histórico (sin IVA): {len(_nuevas)} partida(s) nuevas"
                        + (" · las demás ya estaban guardadas." if len(_nuevas) < len(cotizacion_historico) else ".")
                    )

            if tabla.empty:

                st.error(
                    "No quedaron partidas válidas "
                    "para revisar."
                )

            else:

                # ==========================================================
                # RESULTADOS -- pantalla mínima: resumen, comparativo, acción
                # principal y descarga. Todo el detalle va en desplegables y
                # en el Excel/CSV.
                # ==========================================================
                _FUENTES = (("historico", "Histórico"), ("nl", "Nuevo León"),
                            ("cdmx", "CDMX"), ("ia", "IA internet"))

                def _dinero(valor):
                    return f"${valor:,.2f}" if valor is not None and pd.notna(valor) else "—"

                revision_cant = revision_cantidades.revisar(
                    cotizacion.to_dict("records"),
                    total_declarado=metadatos.get("subtotal_declarado"),
                    iva_en_documento=metadatos.get("iva_en_documento"),
                    contexto=f"{archivo.name} {proyecto or ''}",
                )
                datos_mdo = mercado.datos_mercado(
                    comparador,
                    cotizacion.to_dict("records"),
                    altura_muro=revision_cant.get("altura_muro"),
                )
                resumen_mdo = {str(r["partida"]): r for r in datos_mdo["resumen"]}

                # ---------------- 1. Resumen ----------------
                _importe_total = float(sum(
                    (convertir_numero(f.get("Cantidad")) or 0) * f.get("Precio cotizado") for f in filas
                ))
                _validadas_rango = sum(
                    1 for f in filas
                    if f["_final"].get("respaldo") == "VALIDADA" and f["_final"]["semaforo"] == "EN MERCADO"
                )
                _validadas_caras = sum(
                    1 for f in filas
                    if f["_final"].get("respaldo") == "VALIDADA" and f["_final"]["semaforo"] == "ALTO"
                )
                _pendientes = len(filas) - _validadas_rango - _validadas_caras - sum(
                    1 for f in filas
                    if f["_final"].get("respaldo") == "VALIDADA" and f["_final"]["semaforo"] not in ("EN MERCADO", "ALTO")
                )
                _ahorro_respaldado = sum((f["_final"].get("ahorro_potencial") or 0) for f in filas)
                _pendiente_validar = sum(
                    max(0.0, f["_final"]["diferencia_importe"] or 0)
                    for f in filas if f["_final"].get("respaldo") == "PENDIENTE DE VALIDAR"
                )

                for f in filas:
                    f["_filtros"] = validacion.resultado_filtros(f["_evaluaciones"])
                _n_caras = sum(1 for f in filas if f["_filtros"]["conteo"].get(validacion.ALTO))
                _monto_caro = sum(
                    max(0.0, f["_final"]["diferencia_importe"] or 0) for f in filas
                    if f["_filtros"]["conteo"].get(validacion.ALTO)
                )
                st.subheader("Resultado de la revisión")
                m1, m2, m3 = st.columns(3)
                m1.metric("Total cotizado", f"${_importe_total:,.0f}")
                m2.metric("Partidas revisadas", len(filas))
                m3.metric(
                    "Partidas con algún filtro en alto", f"{_n_caras} de {len(filas)}",
                    help=("Partidas donde al menos un filtro marca rojo (precio arriba de la referencia). "
                          f"Diferencia contra la referencia: ${_monto_caro:,.0f}. "
                          f"Oportunidad contra referencias validadas: ${_ahorro_respaldado:,.0f} (potencial; el ahorro real es el que se negocie)."),
                )

                # Aviso de IA en una línea.
                _motores = [str((r.get("busqueda_ia") or {}).get("motor", ""))
                            for r in fila_registros if (r.get("busqueda_ia") or {}).get("tiene_dato")]
                _n_gemini = sum(1 for m in _motores if m.startswith("Gemini"))
                _n_respaldo = sum(1 for m in _motores if not m.startswith("Gemini"))
                _err_busqueda = (busqueda_mercado_ia.ultimo_error() or {}).get("mensaje")
                _err_revision = (revision_ia.ultimo_error() or {}).get("mensaje") if usar_ia else None
                _n_revisadas = sum(1 for f in filas for k in ("nl", "cdmx", "historico")
                                   if f["_evaluaciones"][k].get("revision_ia"))
                if usar_ia and not _n_revisadas and _err_revision:
                    st.caption("⚠️ La revisión con IA no se pudo hacer; las referencias quedan sin confirmar. "
                               "Detalle en «Diagnóstico de IA».")
                if busqueda_ia_disponible and buscar_precios_web:
                    if _n_gemini:
                        st.caption(f"Búsqueda de precios con IA: {_n_gemini} referencia(s) encontradas con fuente.")
                    elif _err_busqueda:
                        st.caption(
                            "⚠️ La búsqueda principal de IA falló."
                            + (f" Se encontraron {_n_respaldo} referencia(s) orientativa(s) con el buscador de respaldo."
                               if _n_respaldo else "")
                        )

                # ---------------- 2. Comparativo ----------------
                # Una columna por fuente; cada celda trae precio de referencia,
                # diferencia en pesos y %, dictamen (validado u orientativo),
                # confiabilidad y lo que falta confirmar. Orientativo nunca se
                # presenta como validado.
                import html as _html
                _COLOR_CELDA = {
                    # Solo dos colores: rojo = alto, verde = bajo. Fuerte =
                    # validado; claro = posible (referencia orientativa).
                    (True, "ALTO"): ("#f3a5a5", "#5c1414"), (True, "BAJO"): ("#a8dbb4", "#0f3d21"),
                    (False, "ALTO"): ("#fde4e4", "#8a2b2b"), (False, "BAJO"): ("#e3f4e7", "#1f6b3a"),
                }

                def _motivo_corto(ev):
                    m = _html.escape(str(ev.get("motivo") or ""), quote=False)
                    pm = m.lower()
                    eb = ev.get("estado_busqueda")
                    if eb:
                        txt = {"sin referencia localizada": "sin referencia localizada",
                               "sin referencia comparable": "sin referencia comparable",
                               "tiempo agotado": "⏱ tiempo agotado (fallo técnico)",
                               "cuota agotada": "⚠ cuota agotada (fallo técnico)",
                               "error de conexión": "⚠ error de conexión (fallo técnico)",
                               "búsqueda incompleta": "búsqueda incompleta"}.get(eb, eb)
                        if ev.get("reutilizada"):
                            txt += f" · búsqueda del {ev['reutilizada']}"
                        return txt
                    if "no conectado" in pm:
                        return "no conectado"
                    if "cuota" in pm or "429" in pm:
                        return "sin cuota de la IA"
                    if "no se pudo consultar" in pm or "error" in pm:
                        return "falló el buscador"
                    if "sin referencia comparable" in pm:
                        return "sin referencia comparable (se rechazaron conceptos distintos)"
                    if "conectado, sin coincidencias" in pm:
                        return "conectado, sin coincidencias"
                    if "desactivada" in pm:
                        return "búsqueda desactivada"
                    return "no se localizó referencia"

                def _celda_fuente(ev, concepto):
                    def e(x):
                        return _html.escape(str(x), quote=False)
                    simple = validacion.estado_simple(ev)
                    motivo = e(str(ev.get("motivo") or "")[:160])
                    if simple in ("Rechazada", "No comparable"):
                        precio_txt = f"<s>{_dinero(ev['precio_referencia'])}</s> · " if ev.get("precio_referencia") else ""
                        return ("color:#98a2b3", f'<span title="{motivo}">{precio_txt}{simple}<br>'
                                                 f'<small>{motivo[:90]}</small></span>')
                    if simple == "Sin dato":
                        _rech = ev.get("rechazadas") or []
                        if _rech:
                            _r0 = _rech[0]
                            return ("color:#98a2b3", f'<span title="{motivo}">Sin referencia comparable<br><small>'
                                    f'Rechazada: {e(_r0.get("codigo") or "")} {e(str(_r0.get("concepto"))[:45])}… '
                                    f'(<s>{_dinero(_r0.get("precio"))}</s>)</small></span>')
                        _tec = ev.get("estado_busqueda") in ("tiempo agotado", "cuota agotada", "error de conexión")
                        return ("color:#98a2b3" if not _tec else "color:#b54708",
                                f'<span title="{motivo}">{"Búsqueda no completada" if _tec else "Sin dato"}<br>'
                                f'<small>{_motivo_corto(ev)}</small></span>')
                    validada = simple == "Validada"
                    fondo, texto = _COLOR_CELDA.get((validada, ev.get("clasificacion")), ("#f2f4f7", "#344054"))
                    icono = {"ALTO": "🔴", "BAJO": "🟢"}.get(ev["clasificacion"], "🟢")
                    pct = ev.get("diferencia_pct")
                    dif = ev.get("diferencia_unitaria")
                    estilo = f"background:{fondo};color:{texto}" + (";font-weight:600" if validada else "")
                    return (estilo,
                            f"<b>{_dinero(ev['precio_referencia'])}</b><br>"
                            f"{icono} {e(validacion.dictamen_texto(ev))} {pct:+.1f}%<br>"
                            f"<small>{'+' if dif >= 0 else '-'}${abs(dif):,.2f} por unidad<br>"
                            f"{simple}{'' if validada else ' · falta confirmar'}<br>"
                            f"Fuente: {e(validacion.confiabilidad_fuente(ev).split(' (')[0].lower())}"
                            + (f"<br>Incluye inflación acumulada +{ev['inflacion']['acumulada_pct']:.1f}%"
                               if (ev.get("inflacion") or {}).get("acumulada_pct") else "")
                            + (f"<br>⚠ {abs(ev['coherencia']['brecha_pct']):.0f}% debajo de CDMX: úsala como piso"
                               if ev.get("coherencia") else "")
                            + "</small>")

                # ---- Detalle al hacer clic en un recuadro (sin recargar) ----
                def _kv(pares):
                    filas_ = "".join(
                        f"<tr><th>{_html.escape(str(k), quote=False)}</th><td>{val}</td></tr>"
                        for k, val in pares if val not in (None, "", "—"))
                    return f'<table class="kv">{filas_}</table>' if filas_ else ""

                def _detalle_fuente(ev, f, nombre_fuente):
                    def e(x):
                        return _html.escape(str(x), quote=False)
                    concepto = f.get("Concepto")
                    precio_c = f.get("Precio cotizado")
                    simple = validacion.estado_simple(ev)
                    evid = validacion.evidencia(ev, concepto)
                    inf = ev.get("inflacion") or {}
                    partes = [f'<div class="pt">{e(nombre_fuente)} · {e(str(concepto)[:110])}</div>']
                    # --- Por qué este color ---
                    if validacion.cuenta_filtro(ev):
                        pct, dif = ev.get("diferencia_pct") or 0, ev.get("diferencia_unitaria") or 0
                        regla = ("arriba de la referencia = alto (rojo)" if ev["clasificacion"] == "ALTO"
                                 else "igual o abajo de la referencia = bajo (verde)")
                        porque = (
                            f"Tu precio <b>{_dinero(precio_c)}</b> contra la referencia <b>{_dinero(ev['precio_referencia'])}</b> "
                            f"= <b>{pct:+.1f} %</b> ({'+' if dif >= 0 else '-'}${abs(dif):,.2f} por unidad). Regla: {regla}.")
                        if simple != "Validada":
                            porque += (" Dice «posible» y el color es claro porque la referencia es <b>orientativa</b>: "
                                       f"falta confirmar {e(evid['Falta confirmar'] or 'la especificación')}.")
                        else:
                            porque += " Color fuerte: la referencia está <b>validada</b>."
                    else:
                        porque = (f"Gris: <b>{e(simple)}</b>; no entra al resultado ni a los colores. "
                                  f"{e(_motivo_corto(ev)) if simple == 'Sin dato' else e(evid['Motivo'])}")
                    partes.append(f'<div class="ps">Por qué este color</div><div>{porque}</div>')
                    if ev.get("coherencia"):
                        partes.append(f'<div class="pa">⚠ {e(ev["coherencia"]["texto"])}</div>')
                    # --- De dónde sale el dato ---
                    enlace = evid.get("Enlace")
                    partes.append('<div class="ps">De dónde sale el dato</div>' + _kv([
                        ("Fuente / documento", e(evid["Documento / fuente"])),
                        ("Enlace", f'<a href="{_html.escape(str(enlace))}" target="_blank">{e(str(enlace)[:80])}</a>'
                         if enlace else ""),
                        ("Código", e(evid["Código"])), ("Página", e(evid["Página"])), ("Región", e(evid["Región"])),
                        ("Concepto de la referencia", e(evid["Descripción completa de la referencia"])),
                        ("Unidad de la referencia", e(evid["Unidad de la referencia"])),
                        ("Qué incluye el precio", e(evid["Qué incluye el precio"])),
                        ("Fecha / periodo del precio", e(evid["Fecha / periodo del precio"])),
                        ("Fecha de consulta", e(evid["Fecha de consulta"])),
                        ("Verificación", e(evid["Verificación"])),
                        ("Confiabilidad de la fuente", e(evid["Confiabilidad de la fuente"])),
                        ("Estado de la búsqueda", e(evid["Estado de la búsqueda"])),
                        ("Revisión con IA", e(evid["Revisión IA"])),
                        ("Frase de evidencia", e(str(evid["Evidencia (frase)"])[:400])),
                        ("Referencias rechazadas", e(evid["Referencias rechazadas"])),
                    ]))
                    # --- Inflación acumulada ---
                    if inf:
                        tramos = "".join(
                            f"<tr><td>{e(t['de'])} → {e(t['a'])}</td><td>{t['inflacion_pct']:+.2f} %</td>"
                            f"<td>{t['acumulada_pct']:+.2f} %</td></tr>" for t in (inf.get("tramos") or []))
                        partes.append(
                            '<div class="ps">Inflación acumulada aplicada</div>' + _kv([
                                ("Precio original", _dinero(ev.get("precio_original"))),
                                ("Precio actualizado (el que se compara)", f"<b>{_dinero(ev.get('precio_referencia'))}</b>"),
                                ("Inflación acumulada (efectiva)", (f"<b>+{inf['acumulada_pct']:.1f} %</b> (factor {inf.get('factor')})"
                                                         if inf.get("acumulada_pct") is not None else "")),
                                ("Por precio", e(inf.get("acumulada_rango") or "")),
                                ("Periodo", e(f"{inf.get('periodo_base')} → {inf.get('periodo_final')}")),
                                ("Índice", e(inf.get("indice"))),
                            ])
                            + (f'<table class="mini"><tr><th>Tramo (desde {e(inf.get("tramos_desde"))})</th>'
                               f'<th>Inflación del tramo</th><th>Acumulada</th></tr>{tramos}</table>'
                               '<div class="pn">La acumulada se compone (se multiplica año con año); no es la suma.</div>'
                               if tramos else "")
                            + f'<div class="pn">{e(inf.get("justificacion") or "")}</div>')
                    # --- Contratos / compras usados ---
                    todos = ev.get("registros_todos") or []
                    if todos:
                        grupos = {}
                        for r in todos:
                            grupos.setdefault(str(r.get("ocid") or r.get("licitacion")), []).append(r)
                        filas_c = []
                        for k_, g in sorted(grupos.items(), key=lambda kv: pd.Series(
                                [x["precio_actualizado"] for x in kv[1]]).median()):
                            orig = float(pd.Series([x["precio_original"] for x in g]).median())
                            act = float(pd.Series([x["precio_actualizado"] for x in g]).median())
                            usado_ = g[0].get("usado", True)
                            filas_c.append((usado_, (
                                f"<tr{'' if usado_ else ' class=nou'}><td>{e(k_.replace('ocds-7wj9x5-', ''))}</td>"
                                f"<td>{e(g[0].get('anio_licitacion') or '')}</td>"
                                f"<td>{len(g)}</td><td>${orig:,.2f}</td>"
                                f"<td>+{g[0].get('inflacion_acumulada_pct', 0):.1f} %</td><td><b>${act:,.2f}</b></td>"
                                f"<td>{'Sí' if usado_ else 'No (más antiguo)'}</td></tr>")))
                        n_us = sum(1 for u_, _ in filas_c if u_)
                        filas_c = [h_ for u_, h_ in filas_c if u_] + [h_ for u_, h_ in filas_c if not u_]
                        partes.append(
                            f'<div class="ps">Contratos: se usan los {n_us} más recientes de {len(grupos)} localizados</div>'
                            '<table class="mini"><tr><th>Contrato</th><th>Año</th><th>Renglones</th><th>Mediana original</th>'
                            '<th>Inflación acumulada</th><th>Mediana actualizada</th><th>¿Entra al precio?</th></tr>'
                            + "".join(filas_c[:16]) + "</table>"
                            + (f'<div class="pn">… y {len(filas_c) - 16} contrato(s) más en el Excel.</div>'
                               if len(filas_c) > 16 else "")
                            + f'<div class="pn">P.U. Nuevo León = mediana de las medianas de los {n_us} contratos que '
                              f'entran = <b>{_dinero(ev.get("precio_referencia"))}</b> (no es un promedio). Se toman los '
                              "años más recientes hasta reunir al menos 3 contratos; los más antiguos (en gris) quedan "
                              "solo como evidencia. Cada renglón, con su índice y su factor, está en la hoja "
                              "«Inflación NL» del Excel.</div>")
                    regs_h = ev.get("registros_historico") or []
                    if regs_h:
                        partes.append(
                            f'<div class="ps">Compras del histórico usadas ({len(regs_h)})</div>'
                            '<table class="mini"><tr><th>Fecha</th><th>Proveedor</th><th>Proyecto</th><th>Precio</th>'
                            '<th>Inflación acumulada</th><th>Actualizado</th></tr>' + "".join(
                                f"<tr><td>{e(r['fecha'])}</td><td>{e(r['proveedor'])}</td><td>{e(r['proyecto'])}</td>"
                                f"<td>${r['precio_original']:,.2f}</td><td>+{r['inflacion_acumulada_pct']:.1f} %</td>"
                                f"<td><b>${r['precio_actualizado']:,.2f}</b></td></tr>" for r in regs_h[:12])
                            + "</table>")
                    return "".join(partes)

                def _td_clic(estilo, visible, detalle, extra_td=""):
                    return (f'<td class="clic" style="{estilo}"{extra_td}><details name="cmpdet"><summary>{visible}'
                            '<span class="x">✕ Cerrar</span></summary>'
                            f'<div class="pop">{detalle}</div></details></td>')

                st.markdown("**Los 4 filtros por partida** · rojo = alto, verde = bajo · "
                            "haz clic en un recuadro para ver de dónde sale el dato")
                _cab = "".join(f"<th>{n}</th>" for n in ("Concepto", "Cantidad", "Precio cotizado")) + "".join(
                    f"<th>{k}. {n}</th>" for k, (_, n) in enumerate(_FUENTES, 1))
                _filas_html = []
                for f in filas:
                    _tds = [
                        f"<td>{_html.escape(str(f.get('Concepto'))[:90], quote=False)}</td>",
                        f"<td>{convertir_numero(f.get('Cantidad')) or 0:g} {_html.escape(str(f.get('Unidad')))}</td>",
                        f"<td><b>{_dinero(f.get('Precio cotizado'))}</b></td>",
                    ]
                    for clave, _ in _FUENTES:
                        _est, _txt = _celda_fuente(f["_evaluaciones"][clave], f.get("Concepto"))
                        _nom = next(f"{k}. {n}" for k, (c_, n) in enumerate(_FUENTES, 1) if c_ == clave)
                        _tds.append(_td_clic(_est, _txt, _detalle_fuente(f["_evaluaciones"][clave], f, _nom)))
                    _filas_html.append("<tr>" + "".join(_tds) + "</tr>")
                st.markdown(
                    "<style>.cmp{width:100%;border-collapse:collapse;font-size:0.82rem}"
                    ".cmp th{background:#1f3864;color:#fff;padding:6px;text-align:left;white-space:nowrap}"
                    ".cmp td{border:1px solid #d0d5dd;padding:6px;vertical-align:top}"
                    ".cmp small{opacity:.85}"
                    ".cmp td.clic{padding:0;cursor:pointer;height:1px}"
                    "@supports (-moz-appearance:none){.cmp td.clic{height:100%}}"
                    ".cmp td.clic>details{height:100%}"
                    ".cmp td.clic>details>summary{list-style:none;display:block;padding:6px;height:100%;"
                    "box-sizing:border-box}"
                    ".cmp td.clic>details>summary::-webkit-details-marker{display:none}"
                    ".cmp td.clic:hover{box-shadow:inset 0 0 0 2px #1f3864}"
                    ".cmp details[open]>summary{box-shadow:inset 0 0 0 3px #1f3864}"
                    ".cmp details[open]>summary::before{content:'';position:fixed;inset:0;background:rgba(16,24,40,.45);"
                    "z-index:999990;cursor:default}"
                    ".cmp summary .x{display:none}"
                    ".cmp details[open]>summary .x{display:block;position:fixed;z-index:999992;top:calc(7vh + 12px);"
                    "left:calc(50% + min(390px,47vw) - 96px);background:#1f3864;color:#fff;border-radius:6px;"
                    "padding:4px 10px;font-weight:600;font-size:.8rem}"
                    ".cmp .pop{position:fixed;z-index:999991;top:7vh;left:50%;transform:translateX(-50%);"
                    "width:min(780px,94vw);max-height:86vh;overflow:auto;background:#fff;color:#101828;"
                    "border-radius:10px;padding:18px 22px 20px;box-shadow:0 20px 50px rgba(0,0,0,.35);"
                    "font-weight:400;font-size:.84rem;line-height:1.45;cursor:auto;text-align:left}"
                    ".cmp .pop .pt{font-size:1rem;font-weight:700;color:#1f3864;padding-right:96px;margin-bottom:6px}"
                    ".cmp .pop .ps{font-weight:700;color:#1f3864;border-bottom:1px solid #d0d5dd;margin:14px 0 6px}"
                    ".cmp .pop .pn{color:#475467;font-size:.78rem;margin-top:5px}"
                    ".cmp .pop .pa{background:#fff4e5;border-left:4px solid #f79009;padding:6px 10px;margin-top:8px}"
                    ".cmp .pop table{border-collapse:collapse;width:100%;font-size:.8rem}"
                    ".cmp .pop table.kv th{background:none;color:#475467;font-weight:600;white-space:normal;"
                    "width:32%;padding:3px 8px 3px 0;vertical-align:top;text-align:left;border:0}"
                    ".cmp .pop table.kv td{border:0;padding:3px 0;word-break:break-word}"
                    ".cmp .pop table.mini th{background:#eef2f8;color:#1f3864;padding:4px 6px;white-space:normal}"
                    ".cmp .pop table.mini td{border:1px solid #e4e7ec;padding:3px 6px}"
                    ".cmp .pop table.mini tr.nou td{color:#98a2b3}"
                    "</style>"
                    f'<div style="overflow-x:auto"><table class="cmp"><tr>{_cab}</tr>{"".join(_filas_html)}</table></div>',
                    unsafe_allow_html=True,
                )
                # ---------------- 3. Dónde enfocarte para negociar ----------------
                _reco = recomendacion.prioridades(filas, revision_cant)
                _foco = _reco["foco"][:5]
                _texto_ia = None
                if _foco:
                    _lis = "".join(
                        f"<li><b>{_html.escape(it['nombre'], quote=False)}</b> · ${it['importe']:,.0f} "
                        f"({it['peso_pct']:.0f} % del total) — {_html.escape(it['texto'], quote=False)}</li>"
                        for it in _foco)
                    if usar_ia and ia_disponible:
                        _texto_ia = recomendacion.redaccion_ia_con_cache(
                            _reco, lambda prompt: revision_ia._llamar_ia_con_respaldo(prompt, max_tokens=400)[0])
                    # Bloque HTML puro y "$" como entidad: si no, Streamlit
                    # interpreta $...$ como fórmula matemática.
                    st.markdown(
                        ('<div style="font-size:0.88rem"><div style="font-size:1rem;font-weight:700;margin-bottom:4px">'
                         "Dónde enfocarte para negociar</div>"
                         f'{_html.escape(_reco["resumen"], quote=False)}'
                         f'<ol style="margin:4px 0 4px 0">{_lis}</ol>'
                         + (f'<div style="background:#eef2f8;border-left:4px solid #1f3864;padding:6px 10px">'
                            f'<b>Recomendación de la IA:</b> {_html.escape(_texto_ia, quote=False)}</div>'
                            if _texto_ia else "")
                         + "</div>").replace("$", "&#36;"),
                        unsafe_allow_html=True,
                    )
                    st.caption(
                        "Orden: por el dinero en juego de cada partida = (tu precio − precio de referencia) × cantidad, "
                        "por separado para cada filtro (se muestra el rango, no un promedio), más las cantidades por "
                        "aclarar. «Posible» = referencia orientativa. "
                        + ("La IA solo redacta estos mismos datos; no agrega precios."
                           if _texto_ia else "Recomendación calculada con las reglas de la app"
                           + (" (la IA no respondió)." if usar_ia and ia_disponible else ".")))

                # ---------------- 4. Acción principal ----------------
                def _nombre_corto(concepto):
                    palabras = re.sub(r"[^\wÁÉÍÓÚÑáéíóúñ ]", " ", str(concepto)).split()
                    return " ".join(palabras[:1]).lower() if palabras else "partida"

                _pedir_generador, _pedir_espec = [], []
                for h in revision_cant["hallazgos"]:
                    if h["nivel"] == "REVISAR" and h["escenarios"]:
                        a = next((x for x in revision_cant["aritmetica"] if str(x["partida"]) == str(h["partidas"])), None)
                        if a:
                            _pedir_generador.append(
                                f"generador de los {a['cantidad']:g} {str(a['unidad']).lower()} de {_nombre_corto(a['concepto'])}"
                            )
                    elif h["nivel"] == "CONFIRMAR":
                        for p in str(h["partidas"]).split(","):
                            a = next((x for x in revision_cant["aritmetica"] if str(x["partida"]) == p.strip()), None)
                            if a:
                                _pedir_espec.append(f"{_nombre_corto(a['concepto'])} (medidas y separación)")
                # Especificaciones que faltan, por partida, para poder validar
                # las referencias (lo que hay que pedir al proveedor).
                _specs_pedir = {}
                for f in filas:
                    for clave in ("historico", "nl", "cdmx", "ia"):
                        ev_ = f["_evaluaciones"][clave]
                        if validacion.estado_simple(ev_) in ("Orientativa", "No comparable") and ev_.get("descripcion"):
                            falt, _ = validacion.comparar_especificaciones(f.get("Concepto"), ev_.get("descripcion"))
                            sp = validacion.extraer_especificaciones(ev_.get("descripcion"))
                            falt += [c for c in validacion.SPECS_TAMANO if sp.get(c)
                                     and not validacion.extraer_especificaciones(f.get("Concepto")).get(c)]
                            for x in falt:
                                nombre_x = next((n_ for n_ in ("sección / medidas", "espesor", "resistencia f'c",
                                                               "calibre", "capacidad", "diámetro",
                                                               "refuerzo horizontal", "castillos ahogados",
                                                               "acabado aparente", "cimbra", "condición de azotea")
                                                 if x.startswith(n_)), x.split(" (")[0])
                                _specs_pedir.setdefault(_nombre_corto(f.get("Concepto")), []).append(nombre_x)
                for k_, v_ in _specs_pedir.items():
                    _pedir_espec.append(f"{k_} ({', '.join(list(dict.fromkeys(v_))[:4])})")
                _caras = [_nombre_corto(f.get("Concepto")) for f in filas
                          if f["_final"].get("respaldo") == "VALIDADA" and f["_final"]["semaforo"] == "ALTO"]
                _partes = []
                _errores_arit = [a for a in revision_cant["aritmetica"] if not a["ok"]]
                if _errores_arit:
                    _partes.append("corregir importes de " + ", ".join(
                        f"partida {a['partida']} ({'+' if a['diferencia'] >= 0 else '-'}${abs(a['diferencia']):,.2f})"
                        for a in _errores_arit))
                elif not revision_cant["total_ok"] and revision_cant.get("total_declarado"):
                    _partes.append(
                        f"aclarar el total: declarado ${revision_cant['total_declarado']:,.2f} vs calculado "
                        f"${revision_cant['total_calculado']:,.2f}")
                if _pedir_generador:
                    _partes.append("solicitar " + " y ".join(dict.fromkeys(_pedir_generador)))
                _agrupado = {}
                for e_ in _pedir_espec:
                    nom, _, det = e_.partition(" (")
                    _agrupado.setdefault(nom, [])
                    for d_ in det.rstrip(")").split(", "):
                        if d_ and d_ not in _agrupado[nom]:
                            _agrupado[nom].append(d_)
                _espec = [f"{n_} ({', '.join(d_[:4])})" if d_ else n_ for n_, d_ in _agrupado.items()]
                if _espec:
                    _partes.append("pedir especificaciones de " + "; ".join(_espec))
                if _caras:
                    _partes.append("negociar precio de " + ", ".join(dict.fromkeys(_caras)))
                if _partes:
                    _accion = "; ".join(_partes)
                    st.info(f"**Acción principal:** {_accion[0].upper() + _accion[1:]}.")
                else:
                    st.success("**Acción principal:** sin observaciones; los precios validados están en rango.")

                # ---------------- 4. Descarga ----------------
                configuracion_ia = (
                    f"Versión {_version.VERSION}. Configuración IA: revisión {'SÍ' if usar_ia else 'NO'} · "
                    f"búsqueda de precios {'SÍ' if (busqueda_ia_disponible and buscar_precios_web) else 'NO'} · "
                    f"generado {pd.Timestamp.now(tz='America/Monterrey'):%Y-%m-%d %H:%M}."
                )
                nombre_proveedor = re.sub(r"[^a-zA-Z0-9_-]+", "_", proveedor.strip() if proveedor else "proveedor")
                nombre_proyecto = re.sub(r"[^a-zA-Z0-9_-]+", "_", proyecto.strip() if proyecto else "proyecto")
                st.download_button(
                    "⬇️ Descargar análisis completo (Excel)",
                    data=exportar_revision.generar_excel(
                        tabla.to_dict("records"), proveedor=proveedor, proyecto=proyecto,
                        revision_cantidades=revision_cant, configuracion=configuracion_ia,
                        datos_mercado=datos_mdo, recomendacion={**_reco, "texto_ia": _texto_ia},
                    ),
                    file_name=f"revision_{nombre_proveedor}_{nombre_proyecto}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                )

                # ---------------- Detalle (desplegables) ----------------
                registros_csv = []
                for f in filas:
                    evs, fin_ = f["_evaluaciones"], f["_final"]
                    cant = convertir_numero(f.get("Cantidad"))
                    fila_csv = {
                        "#": f.get("Partida"), "Descripción cotizada": f.get("Concepto"),
                        "Unidad": f.get("Unidad"), "Cantidad": cant,
                        "P.U. cotizado": f.get("Precio cotizado"),
                        "Importe cotizado": round((cant or 0) * f.get("Precio cotizado"), 2),
                    }
                    for clave, nombre in _FUENTES:
                        ev = evs[clave]
                        fuente_txt = {
                            "historico": "Histórico interno Ragasa (Google Sheets)",
                            "nl": "Tabulador homologado de licitaciones NL (SIASI), ajustado INPC",
                            "cdmx": "Tabulador General de Precios Unitarios CDMX 2026",
                            "ia": ev.get("fuente_web") or ev.get("descripcion") or "Búsqueda web",
                        }[clave] if ev["estado"] != validacion.SIN_DATO else ""
                        fila_csv.update({
                            f"{nombre} · P.U. referencia": ev["precio_referencia"],
                            f"{nombre} · diferencia por unidad": ev.get("diferencia_unitaria"),
                            f"{nombre} · diferencia %": ev.get("diferencia_pct"),
                            f"{nombre} · diferencia importe": (
                                round(ev["diferencia_unitaria"] * cant, 2)
                                if ev.get("diferencia_unitaria") is not None and cant else None
                            ),
                        })
                        fila_csv.update({f"{nombre} · {k}": val for k, val in
                                         validacion.evidencia(ev, f.get("Concepto")).items()})
                    _mdo = resumen_mdo.get(str(f.get("Partida")), {})
                    fila_csv.update({
                        "Mercado · referencias": _mdo.get("referencias"),
                        "Mercado · tipo": _mdo.get("tipo"),
                        "Mercado NL · concepto": _mdo.get("nl_concepto"),
                        "Mercado NL · p25": _mdo.get("nl_p25"),
                        "Mercado NL · mediana": _mdo.get("nl_mediana"),
                        "Mercado NL · p75": _mdo.get("nl_p75"),
                        "Mercado NL · registros": _mdo.get("nl_registros"),
                        "Mercado NL · periodo": _mdo.get("nl_periodo"),
                        "Mercado CDMX · concepto": _mdo.get("cdmx_concepto"),
                        "Mercado CDMX · precio": _mdo.get("cdmx_precio"),
                        "Mercado · posición": _mdo.get("posicion"),
                        "Resultado de los 4 filtros": f["_filtros"]["texto_plano"],
                        "Referencias validadas": f["_filtros"]["detalle"],
                        "Referencia": fin_["referencia_negociacion"] or "",
                        "Respaldo": ("VALIDADA" if f["_filtros"]["validados"] else ("POR VALIDAR" if f["_filtros"]["n"] else "—")),
                        "% vs referencia": fin_["diferencia_pct"],
                        "Diferencia contra referencia": fin_["diferencia_importe"],
                        "Oportunidad validada (potencial)": fin_["ahorro_potencial"] or 0.0,
                        "Nota": fin_["detalle"] or "",
                        "Configuración": configuracion_ia,
                        "Versión de la app": _version.VERSION,
                    })
                    registros_csv.append(fila_csv)

                with st.expander("Ver detalle (evidencia, cantidades, mercado, metodología)"):
                    _tabs_detalle = st.tabs(['Qué falta confirmar y evidencia', 'Mercado', 'Cantidades', 'Cómo se lee', 'Estado de la IA', 'CSV'])
                    with _tabs_detalle[0]:
                        st.dataframe(pd.DataFrame([{
                            "#": f.get("Partida"), "Concepto": f.get("Concepto"), "Fuente": nombre,
                            "P.U. referencia": f["_evaluaciones"][clave]["precio_referencia"],
                            **validacion.evidencia(f["_evaluaciones"][clave], f.get("Concepto")),
                        } for f in filas for clave, nombre in _FUENTES]), width="stretch", hide_index=True)

                    with _tabs_detalle[1]:
                        st.dataframe(pd.DataFrame([{
                            "#": r["partida"], "Concepto": r["concepto"], "Cotizado": _dinero(r["precio_cotizado"]),
                            "Referencias": f"{r['referencias']} ({r['tipo']})",
                            "NL p25": _dinero(r.get("nl_p25")),
                            "NL mediana": _dinero(r.get("nl_mediana")),
                            "NL p75": _dinero(r.get("nl_p75")),
                            "CDMX (concepto aparte)": _dinero(r.get("cdmx_precio")),
                            "Posición del cotizado": r["posicion"],
                        } for r in datos_mdo["resumen"]]), width="stretch", hide_index=True)
                        st.dataframe(pd.DataFrame([{
                            "#": r["partida"], "Fuente": r["fuente"], "Concepto de referencia": r["concepto"],
                            "Relación": r["relacion"], "Unidad": r["unidad"], "Precio (mediana)": _dinero(r["precio"]),
                            "Rango p25–p75": (f"{_dinero(r['rango_bajo'])} – {_dinero(r['rango_alto'])}"
                                              if r.get("rango_bajo") else "—"),
                            "Registros": r["n_registros"], "Periodo / fecha": r["fecha"],
                            "Equivalencia": r["equivalencia"], "Cotizado vs referencia": r["posicion"],
                        } for r in datos_mdo["referencias"]]), width="stretch", hide_index=True)
                        st.caption("NL: precios de licitaciones de obra pública publicados por SIASI en formato OCDS (mediana y rango p25–p75, "
                                   "ajustados por INPC). CDMX: tabulador oficial 2026. Los escenarios por pieza "
                                   "(precio por ml × altura supuesta) no entran al rango hasta confirmar dimensiones.")

                    with _tabs_detalle[2]:
                        if revision_cant["partidas_con_error"] == 0 and revision_cant["total_ok"]:
                            st.markdown(f"🟢 Aritmética correcta: total ${revision_cant['total_calculado']:,.2f}.")
                        else:
                            st.markdown(f"🔴 {revision_cant['partidas_con_error']} partida(s) con diferencia aritmética.")
                            st.dataframe(pd.DataFrame([a for a in revision_cant["aritmetica"] if not a["ok"]]),
                                         width="stretch", hide_index=True)
                        _iconos = {"REVISAR": "🔴", "CONFIRMAR": "🟡", "OK": "🟢"}
                        for h in revision_cant["hallazgos"]:
                            st.markdown(f"{_iconos.get(h['nivel'], '•')} **{h['tipo']}** (partida {h['partidas']}): {h['detalle']}")
                            if h["escenarios"]:
                                st.dataframe(pd.DataFrame([{
                                    "Escenario": e["nombre"], "Fórmula": e["formula"], "Base": e.get("base", ""),
                                    "m² del escenario": e["m2"],
                                    "Diferencia m² (cotizado − escenario)": e["dif_m2"],
                                    "Diferencia × P.U. (sujeta a aclaración)": (
                                        f"${e['importe']:,.2f}" if e["importe"] >= 0 else f"-${-e['importe']:,.2f}"),
                                } for e in h["escenarios"]]), width="stretch", hide_index=True)
                        st.markdown("**Alcances e impuestos a confirmar por escrito**")
                        for a in revision_cant["alcances_confirmar"]:
                            st.markdown(f"- {a}")

                    with _tabs_detalle[3]:
                        st.markdown(
                            "- **Colores:** rojo = alto (tu precio está arriba de la referencia); verde = bajo (igual o "
                            "abajo). Color fuerte y «Alto / Bajo» = referencia validada; color claro y «Posible alto / "
                            "Posible bajo» = referencia orientativa; gris = no se puede comparar.\n"
                            "- **IVA:** "
                            + ("se usa el subtotal sin IVA de la cotización.\n" if metadatos.get("iva_en_documento") else
                               "la cotización no indica IVA; se supone que sus precios son antes de IVA (por confirmar "
                               "con el proveedor).\n")
                            + "- **Validada:** la IA confirma el concepto, el proveedor declara las especificaciones e "
                            "inclusiones de la referencia (sección, espesor, f'c, refuerzo, condición de azotea) y el "
                            "precio es vigente (≤ 12 meses).\n"
                            "- **Orientativo:** hay un precio real de un concepto parecido, pero falta demostrar especificación, "
                            "inclusiones, alcance o vigencia (se indica qué falta). Su diferencia se reporta como *pendiente de "
                            "validar*, nunca como ahorro.\n"
                            "- **Rechazada:** la IA la rechaza, o el material, el elemento (p. ej. castillo vs. cerramiento), "
                            "la especificación o el alcance son distintos.\n"
                            "- **Oportunidad validada:** diferencia contra una referencia validada cuando el cotizado está arriba de ella. Es potencial: el ahorro real es la reducción que se negocie o contrate.\n"
                            "- **Sin promedios:** cada fuente se compara por separado; la referencia de negociación sigue la "
                            "prioridad Histórico Ragasa > Nuevo León > CDMX > IA.\n"
                            "- **Precios web:** la fecha del precio solo se registra si la fuente la indica; aparte se guarda "
                            "la fecha de consulta."
                        )

                    with _tabs_detalle[4]:
                        st.write(f"Revisión con IA: {'activa' if usar_ia else 'apagada'} · "
                                 f"{_n_revisadas} referencia(s) revisada(s).")
                        st.write(f"Búsqueda de precios: Gemini {_n_gemini} · buscador de respaldo {_n_respaldo}.")
                        if _err_revision:
                            st.code(f"Revisión IA: {_err_revision}", language=None)
                        if _err_busqueda:
                            st.code(f"Búsqueda IA: {_err_busqueda}", language=None)
                        st.caption(configuracion_ia)

                    with _tabs_detalle[5]:
                        st.download_button(
                            "CSV comparativo completo (4 fuentes con evidencia)",
                            data=pd.DataFrame(registros_csv).to_csv(index=False).encode("utf-8-sig"),
                            file_name=f"comparativo_{nombre_proveedor}_{nombre_proyecto}.csv",
                            mime="text/csv",
                        )
                        st.download_button(
                            "CSV de referencias de mercado",
                            data=pd.DataFrame(datos_mdo["referencias"]).to_csv(index=False).encode("utf-8-sig"),
                            file_name=f"mercado_{nombre_proveedor}_{nombre_proyecto}.csv",
                            mime="text/csv",
                        )
                        _regs_nl = [
                            {"#": f.get("Partida"), "Concepto cotizado": f.get("Concepto"), **r}
                            for f in filas for r in (f["_evaluaciones"]["nl"].get("registros_todos") or [])
                        ]
                        st.download_button(
                            "CSV de contratos NL usados (inflación por registro)",
                            data=pd.DataFrame(_regs_nl).to_csv(index=False).encode("utf-8-sig"),
                            file_name=f"inflacion_nl_{nombre_proveedor}_{nombre_proyecto}.csv",
                            mime="text/csv",
                        )


    except Exception as error:

        st.error(
            "No fue posible interpretar "
            "automáticamente el archivo."
        )

        st.exception(
            error
        )

        st.info(
            "Revisa que el archivo contenga una tabla "
            "con concepto, unidad y precio unitario. "
            "Los PDF escaneados todavía requieren OCR."
        )


# ==========================================================
# INFORMACIÓN DE FUENTES
# ==========================================================

st.divider()

st.caption(
    "Fuentes: histórico de licitaciones de obra pública "
    "de Nuevo León, Tabulador General de Precios Unitarios "
    "del Gobierno de la Ciudad de México e histórico interno "
    "guardado en Google Sheets."
)

st.caption(
    "El ajuste por inflación utiliza el INPC como "
    "aproximación para actualizar precios antiguos. "
    "No sustituye un índice específico de construcción."
)
