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
st.caption("Versión del comparativo: 2026-09-26 · búsqueda por fuente 3")


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


comparador = cargar_comparador()
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
        r"^(?:(\d+(?:\.\d+)?)\s+)?"
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

    for numero_pagina, linea in lineas_pdf:
        linea_normalizada = normalizar_texto(
            linea
        )

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
        placeholder="Nombre del proveedor",
    )

    proyecto = st.text_input(
        "Proyecto / licitación",
        placeholder="Ej. Planta Norte 2026",
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

    ajustar_inflacion = st.checkbox(
        "Ajustar precios viejos de NL por inflación",
        value=True,
        help=(
            "Los precios históricos de Nuevo León "
            "se actualizan usando INPC."
        ),
    )

    usar_ia = st.checkbox(
        "Activar revisión con IA (segunda opinión)",
        value=ia_disponible,
        disabled=not ia_disponible,
        help=(
            "Para coincidencias de confianza BAJA, le pide a la IA "
            "que confirme si de verdad es el mismo material. Para "
            "partidas sin ningún dato de referencia, da una opinión "
            "orientativa (NO un precio verificado) sobre si el precio "
            "cotizado suena razonable. Requiere una API key de "
            "Gemini o de OpenAI configurada en Secrets."
        ),
    )

    if not ia_disponible:

        st.caption(
            "Revisión con IA no conectada. Falta configurar "
            "gemini_api_key u openai_api_key en Secrets."
        )

    else:

        _proveedores_ia = revision_ia.proveedores_disponibles()

        if len(_proveedores_ia) > 1:

            st.caption(
                "Revisión con IA: Gemini con respaldo automático en "
                "OpenAI (ChatGPT) si Gemini se queda sin cuota."
            )

        elif "openai" in _proveedores_ia:

            st.caption("Revisión con IA: OpenAI (ChatGPT).")

        else:

            st.caption(
                "Revisión con IA: Gemini. Agrega también "
                "'openai_api_key' en Secrets para tener respaldo "
                "automático si se acaba la cuota de Gemini."
            )

    revisar_todo_con_ia = st.checkbox(
        "Revisar con IA TODAS las coincidencias, no solo las dudosas",
        value=False,
        disabled=not (ia_disponible and usar_ia),
        help=(
            "Por default la IA solo revisa las coincidencias de "
            "confianza BAJA o con diferencia de precio extrema (así "
            "rinde más la cuota gratuita). Actívalo para que la IA "
            "confirme TODAS las coincidencias, incluso las de "
            "confianza ALTA/MEDIA, por si acaso -- usa muchas más "
            "llamadas a la IA y agota la cuota gratuita más rápido "
            "en cotizaciones grandes."
        ),
    )

    if busqueda_ia_disponible:

        buscar_precios_web = st.checkbox(
            "Buscar precio IA para todas las partidas",
            value=True,
            help=(
                "Puede tardar varios minutos en cotizaciones grandes. "
                "Consulta cada concepto, incluso si ya existe un precio en "
                "Ragasa, Nuevo León o CDMX. Requiere Gemini o Tavily."
            ),
        )

        st.caption(
            "4ª fuente activa: la IA busca en internet un precio de "
            "mercado real, con la fuente citada, para cada partida. "
            "Usa Gemini (Google Search) como motor principal y, si "
            "se acaba su cuota, cae automáticamente a Tavily como "
            "respaldo 100% gratuito (1,000 búsquedas/mes, sin "
            "tarjeta)."
        )

    else:

        buscar_precios_web = False

        st.caption(
            "4ª fuente ('IA busca cotizaciones en internet') no "
            "conectada. Configura 'gemini_api_key' en Secrets, o "
            "'tavily_api_key' (gratis, sin tarjeta, en tavily.com) "
            "como alternativa."
        )

    if historico is None:

        st.caption(
            "Histórico interno no conectado. "
            "Falta configurar Google Sheets en Secrets."
        )

    else:

        resumen = historico.resumen()

        st.caption(
            f"Histórico interno: "
            f"{resumen['total_renglones']} renglones · "
            f"{len(resumen['proveedores'])} proveedores · "
            f"{len(resumen['proyectos'])} proyectos"
        )

    if inpc_en_vivo:

        st.caption(
            f"INPC: dato en vivo de INEGI "
            f"({ajuste_inflacion.ETIQUETA_ACTUAL}, "
            f"nivel {ajuste_inflacion.NIVEL_ACTUAL})."
        )

    else:

        st.caption(
            f"INPC: usando valor de respaldo "
            f"({ajuste_inflacion.ETIQUETA_ACTUAL}, "
            f"nivel {ajuste_inflacion.NIVEL_ACTUAL}). "
            "Configura 'inegi_api_token' en Secrets "
            "para traerlo automático."
        )


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
        use_container_width=True,
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
            use_container_width=True,
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
                            if fuente_dict.get("match") and (
                                fuente_dict.get("motivo")
                                or str(fuente_dict.get("confianza", "")).upper() == "BAJA"
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

                        for inicio in range(
                            0, len(items_revision), revision_ia.TAMANO_LOTE
                        ):
                            lote = items_revision[
                                inicio:inicio + revision_ia.TAMANO_LOTE
                            ]
                            resultados_revision.update(
                                revision_ia.revisar_coincidencias_debiles_lote(lote)
                            )

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
                        }
                        for indice, registro in enumerate(fila_registros)
                    ]

                    resultados_busqueda_mercado = {}
                    cache_precios_ia = st.session_state.setdefault(
                        "precios_ia_verificados", {}
                    )
                    pendientes = []
                    for item in items_busqueda_mercado:
                        clave = ("fuente-verificada-v4", item["descripcion"].casefold().strip(), item["unidad"].casefold().strip())
                        if clave in cache_precios_ia:
                            resultados_busqueda_mercado[item["id"]] = cache_precios_ia[clave]
                        else:
                            pendientes.append(item)

                    progreso_busqueda = st.progress(
                        0, text="Preparando búsqueda de precios externos..."
                    )

                    for inicio in range(
                        0, len(pendientes), busqueda_mercado_ia.TAMANO_LOTE
                    ):
                        lote = pendientes[
                            inicio:inicio + busqueda_mercado_ia.TAMANO_LOTE
                        ]
                        respuesta_lote = busqueda_mercado_ia.buscar_precios_mercado_lote(lote)
                        resultados_busqueda_mercado.update(respuesta_lote)
                        for item in lote:
                            respuesta = respuesta_lote.get(item["id"])
                            if respuesta and respuesta.get("tiene_dato") and respuesta.get("precio_mxn"):
                                clave = (item["descripcion"].casefold().strip(), item["unidad"].casefold().strip())
                                cache_precios_ia[clave] = respuesta

                        completadas = min(
                            len(items_busqueda_mercado) - len(pendientes) + inicio + len(lote),
                            len(items_busqueda_mercado)
                        )
                        progreso_busqueda.progress(
                            completadas / len(items_busqueda_mercado),
                            text=(
                                f"Búsqueda externa: {completadas}/"
                                f"{len(items_busqueda_mercado)} partidas"
                            ),
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
                        banda_baja_ia, banda_alta_ia = banda_en_mercado(precio_mercado_ia)
                        clasificacion_ia_mercado = clasificar(
                            precio, banda_baja_ia, banda_alta_ia
                        )
                        _resultado_ia = clasificacion_ia_mercado
                        _diff_ia = round(
                            (precio - precio_mercado_ia) / precio_mercado_ia * 100, 1
                        )

                        fila["Precio mercado (IA internet)"] = precio_mercado_ia
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

                    _resultado_nl = nl.get("clasificacion")
                    if nl_rechazado_por_ia and _resultado_nl:
                        _resultado_nl = f"{_resultado_nl} (descartado: {motivo_descarte_nl})"

                    _resultado_cdmx = cdmx.get("clasificacion")
                    if cdmx_rechazado_por_ia and _resultado_cdmx:
                        _resultado_cdmx = f"{_resultado_cdmx} (descartado: {motivo_descarte_cdmx})"

                    clasificaciones = [
                        valor
                        for valor in (
                            nl.get("clasificacion")
                            if not nl_rechazado_por_ia
                            else None,
                            cdmx.get("clasificacion")
                            if not cdmx_rechazado_por_ia
                            else None,
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
                            nl.get("precio_mediana_ajustada")
                            if not nl_rechazado_por_ia
                            else None,
                            cdmx.get("precio_referencia")
                            if not cdmx_rechazado_por_ia
                            else None,
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
                            referencias_precio.append(busqueda_ia["precio_mxn"])

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

                        fila["RESULTADO FINAL"] = max(conteo, key=conteo.get)

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

                    if (
                        not proveedor
                        or not proyecto
                    ):

                        st.warning(
                            "Escribe el proveedor y el proyecto "
                            "en la barra lateral para guardar "
                            "la cotización en el histórico."
                        )

                    else:

                        historico.ingerir(
                            cotizacion_historico,
                            proveedor=proveedor,
                            proyecto=proyecto,
                        )

                        st.success(
                            "Cotización guardada en el "
                            "histórico interno. "
                            f"{len(cotizacion_historico)} "
                            "partidas agregadas."
                        )

            if tabla.empty:

                st.error(
                    "No quedaron partidas válidas "
                    "para revisar."
                )

            else:

                resumen_veredictos = tabla[
                    "RESULTADO FINAL"
                ].value_counts()

                _ahorro_potencial_total = (
                    tabla["Ahorro potencial"].sum()
                    if "Ahorro potencial" in tabla.columns
                    else 0
                )

                c1, c2, c3, c4, c5, c6 = st.columns(
                    6
                )

                c1.metric(
                    "Partidas revisadas",
                    len(tabla),
                )

                c2.metric(
                    "🔴 Caras",
                    int(
                        resumen_veredictos.get(
                            "ALTO",
                            0,
                        )
                    ),
                    help="Precio cotizado por encima del precio de mercado.",
                )

                c3.metric(
                    "🟢 Baratas",
                    int(
                        resumen_veredictos.get(
                            "BAJO",
                            0,
                        )
                    ),
                    help="Precio cotizado por debajo del precio de mercado.",
                )

                c4.metric(
                    "🟡 En rango",
                    int(
                        resumen_veredictos.get(
                            "EN MERCADO",
                            0,
                        )
                    ),
                    help="Precio cotizado dentro del rango normal de mercado.",
                )

                c5.metric(
                    "⚪ Sin datos",
                    int(
                        resumen_veredictos.get(
                            "SIN DATOS SUFICIENTES",
                            0,
                        )
                    ),
                    help="No se encontró ninguna referencia confiable para comparar esta partida.",
                )

                c6.metric(
                    "💰 Ahorro potencial",
                    f"${_ahorro_potencial_total:,.0f}",
                    help=(
                        "Suma de (precio cotizado - precio sugerido) × "
                        "cantidad, solo en las partidas donde el precio "
                        "cotizado está por encima del precio de "
                        "referencia -- lo que se podría negociar a la "
                        "baja si se logra el precio sugerido."
                    ),
                )

                # ----------------------------------------------------------
                # Diagnóstico real de la IA: si se activó la revisión con IA
                # o la búsqueda en internet y de verdad no ayudó en nada,
                # aquí se muestra el ÚLTIMO error real que dio Gemini/OpenAI
                # (ej. "modelo no encontrado", "permiso denegado", "se acabó
                # la cuota") en vez de solo un mensaje genérico -- así se
                # puede saber la causa real en lugar de adivinar.
                # ----------------------------------------------------------
                _error_revision = revision_ia.ultimo_error() if usar_ia else None
                _error_busqueda = (
                    busqueda_mercado_ia.ultimo_error()
                    if busqueda_ia_disponible and buscar_precios_web
                    else None
                )

                if (
                    (_error_revision and _error_revision.get("mensaje"))
                    or (_error_busqueda and _error_busqueda.get("mensaje"))
                ):

                    with st.expander(
                        "⚠️ La IA tuvo errores durante esta revisión "
                        "(clic para ver el detalle técnico)"
                    ):

                        if _error_revision and _error_revision.get("mensaje"):
                            st.caption(
                                f"Revisión con IA "
                                f"({_error_revision.get('proveedor') or '?'}): "
                                f"{_error_revision['mensaje']}"
                            )

                        if _error_busqueda and _error_busqueda.get("mensaje"):
                            st.caption(
                                f"Búsqueda en internet (4ª fuente): "
                                f"{_error_busqueda['mensaje']}"
                            )

                # Precio caro/en rango/bajo -- pedido explícito de dirección:
                # ALTO (caro) = rojo, EN MERCADO (en rango) = amarillo,
                # BAJO (por debajo del precio de referencia) = verde.
                def resaltar(valor):

                    if valor == "ALTO":
                        return (
                            "background-color: #f7c1c1; "
                            "color: #501313"
                        )

                    if valor == "EN MERCADO":
                        return (
                            "background-color: #ffe699; "
                            "color: #7f6000"
                        )

                    if valor == "BAJO":
                        return (
                            "background-color: #c6e0b4; "
                            "color: #006100"
                        )

                    if (
                        valor
                        == "SIN DATOS SUFICIENTES"
                    ):
                        return (
                            "background-color: #d9d9d9; "
                            "color: #595959"
                        )

                    return ""

                # Confianza del TEXTO encontrado (qué tan seguro está el
                # sistema de que el match es el mismo material) -- también
                # pedido de dirección, en sentido contrario al de arriba:
                # ALTA confianza = bueno = verde, BAJA confianza = hay que
                # revisar a mano = rojo.
                def resaltar_confianza(valor):

                    if valor == "ALTA":
                        return (
                            "background-color: #c6e0b4; "
                            "color: #006100"
                        )

                    if valor == "MEDIA":
                        return (
                            "background-color: #ffe699; "
                            "color: #7f6000"
                        )

                    if valor == "BAJA":
                        return (
                            "background-color: #f7c1c1; "
                            "color: #501313"
                        )

                    return ""

                def resaltar_diferencia(valor):

                    if valor is None or pd.isna(valor):
                        return ""

                    if valor > 0:
                        return "color: #c0392b"

                    if valor < 0:
                        return "color: #1e8449"

                    return ""

                # ------------------------------------------------------------
                # Vista simple (default) vs. vista avanzada: con 4 fuentes,
                # cada una con su match/confiabilidad/precio/resultado/
                # diferencia, la tabla completa llega a tener ~25-30
                # columnas -- demasiado para leer de un vistazo. Por default
                # se muestra solo lo que hace falta para decidir (veredicto
                # combinado + qué tan lejos está del mercado); quien quiera
                # ver POR QUÉ se llegó a ese veredicto (con qué coincidió
                # cada fuente, su confiabilidad, su precio) puede activar la
                # vista avanzada con el mismo detalle de siempre.
                # ------------------------------------------------------------
                st.markdown("#### Comparativo por partida")
                st.caption(
                    "🔴 Cotizado alto · 🟡 En mercado · 🟢 Cotizado bajo · "
                    "Gris: sin referencia validada. Los precios son unitarios; "
                    "el total corresponde a cantidad × precio cotizado."
                )

                def _columna(nombre):
                    return tabla[nombre] if nombre in tabla else pd.Series(
                        [None] * len(tabla), index=tabla.index
                    )

                comparativo = pd.DataFrame(index=tabla.index)
                comparativo["# Concepto"] = _columna("Partida")
                comparativo["Descripción"] = _columna("Concepto")
                comparativo["Cantidad"] = _columna("Cantidad")
                comparativo["Precio"] = _columna("Precio cotizado")
                comparativo["Precio total"] = _columna("Importe")
                _total_calculado = pd.to_numeric(comparativo["Cantidad"], errors="coerce") * pd.to_numeric(comparativo["Precio"], errors="coerce")
                comparativo["Precio total"] = comparativo["Precio total"].fillna(_total_calculado)
                comparativo["Descripción histórico Ragasa"] = _columna("Match histórico interno")
                comparativo["Precio histórico"] = _columna("Precio mediana histórico")
                comparativo["Confiabilidad histórico"] = _columna("Confiabilidad histórico interno")
                comparativo["Descripción histórico Ragasa"] = comparativo["Descripción histórico Ragasa"].replace(
                    {"todavía no hay": None, "no hay": None, "None": None}
                )
                comparativo["Descripción Nuevo León"] = _columna("Match NL")
                comparativo["Precio Nuevo León"] = _columna("Precio mediana NL (ajustado hoy)")
                comparativo["Confiabilidad Nuevo León"] = _columna("Confiabilidad NL")
                comparativo["Descripción CDMX"] = _columna("Match CDMX")
                comparativo["Precio CDMX"] = _columna("Precio referencia CDMX")
                comparativo["Confiabilidad CDMX"] = _columna("Confiabilidad CDMX")
                _nota_ia = _columna("Nota IA internet")
                _fuente_ia = _columna("Fuente IA (internet)")
                _opinion_ia = _columna("Opinión IA (sin datos verificados)")
                def _texto_visible(valor):
                    return "" if pd.isna(valor) or str(valor).strip().lower() in ("none", "nan") else str(valor).strip()
                comparativo["Descripción IA"] = [
                    " · ".join(x for x in (_texto_visible(fuente), _texto_visible(nota)) if x)
                    or _texto_visible(opinion)
                    for fuente, nota, opinion in zip(_fuente_ia, _nota_ia, _opinion_ia)
                ]
                comparativo["Precio IA"] = _columna("Precio mercado (IA internet)")
                for _precio_col, _resultado_col in (
                    ("Precio histórico", "Resultado histórico"),
                    ("Precio Nuevo León", "Resultado NL"),
                    ("Precio CDMX", "Resultado CDMX"),
                ):
                    if _resultado_col in tabla:
                        _descartados = tabla[_resultado_col].astype(str).str.contains(
                            "descartado", case=False, na=False
                        )
                        comparativo.loc[_descartados, _precio_col] = float("nan")
                for _precio_col, _confianza_col in (
                    ("Precio histórico", "Confiabilidad histórico"),
                    ("Precio Nuevo León", "Confiabilidad Nuevo León"),
                    ("Precio CDMX", "Confiabilidad CDMX"),
                ):
                    _baja = comparativo[_confianza_col].astype(str).str.upper().eq("BAJA")
                    comparativo.loc[_baja, _precio_col] = float("nan")

                comparativo["Confiabilidad IA"] = [
                    "Precio en fuente: revisar alcance" if pd.notna(valor) else
                    "Opinión sin precio verificable" if pd.notna(opinion) and str(opinion).strip() else
                    "Sin precio verificable"
                    for valor, opinion in zip(comparativo["Precio IA"], _opinion_ia)
                ]

                # Ausencia de referencia: mostrar un guion, sin textos técnicos como None.
                # La descripción original vacía se señala para evitar búsquedas sin concepto.
                comparativo["Descripción"] = comparativo["Descripción"].apply(
                    lambda v: _texto_visible(v) or "Descripción no detectada: revisar archivo de origen"
                )
                for _col in comparativo.columns:
                    if _col.startswith("Descripción ") or _col.startswith("Confiabilidad "):
                        comparativo[_col] = comparativo[_col].apply(
                            lambda v: _texto_visible(v) or "—"
                        )

                for _precio in ("Precio", "Precio total", "Precio histórico",
                                "Precio Nuevo León", "Precio CDMX", "Precio IA"):
                    comparativo[_precio] = pd.to_numeric(comparativo[_precio], errors="coerce")

                _resultado_por_precio = {
                    "Precio histórico": "Resultado histórico",
                    "Precio Nuevo León": "Resultado NL",
                    "Precio CDMX": "Resultado CDMX",
                    "Precio IA": "Resultado IA internet",
                }
                _formatos = {
                    col: lambda valor: f"${valor:,.2f}" if pd.notna(valor) else "—"
                    for col in ["Precio", "Precio total", *_resultado_por_precio]
                }
                def _estilo_precio(data):
                    estilos = pd.DataFrame("", index=data.index, columns=data.columns)
                    colores = {
                        "ALTO": "background-color: #f8c9c9; color: #712121; font-weight: 650",
                        "BAJO": "background-color: #ccebd2; color: #14532d; font-weight: 650",
                        "EN MERCADO": "background-color: #fff0bb; color: #705000; font-weight: 650",
                    }
                    for precio_col, resultado_col in _resultado_por_precio.items():
                        for indice in data.index:
                            resultado = str(tabla.at[indice, resultado_col]) if resultado_col in tabla else ""
                            valor = comparativo.at[indice, precio_col]
                            estilos.at[indice, precio_col] = (
                                "background-color: #f1f3f5; color: #60666d"
                                if pd.isna(valor) else colores.get(resultado, "")
                            )
                    return estilos

                comparativo_vista = comparativo.copy()
                for _precio in ("Precio", "Precio total", "Precio histórico",
                                "Precio Nuevo León", "Precio CDMX", "Precio IA"):
                    comparativo_vista[_precio] = comparativo[_precio].map(
                        lambda v: f"${v:,.2f}" if pd.notna(v) else "—"
                    )
                st.dataframe(
                    comparativo_vista.style.apply(_estilo_precio, axis=None),
                    use_container_width=True,
                    height=min(680, 48 + 36 * len(comparativo)),
                    hide_index=True,
                )
                with st.expander("Cómo se validó cada coincidencia"):
                    st.caption(
                        "La confiabilidad describe la equivalencia del concepto, no garantiza "
                        "que precio, alcance, unidad y fecha sean comparables. La IA revisa "
                        "coincidencias dudosas en lotes. Una búsqueda de internet solo aporta "
                        "precio si encontró una fuente; confirma su vigencia antes de negociar."
                    )
                    detalle = [c for c in tabla.columns if c.startswith("Revisión IA")]
                    if detalle:
                        st.dataframe(tabla[["Partida", "Concepto", *detalle]],
                                     use_container_width=True, hide_index=True)
                    else:
                        st.info("No hubo coincidencias dudosas que requirieran revisión adicional.")

                buffer = io.BytesIO()

                comparativo.replace({"None": "", "nan": ""}).to_excel(
                    buffer,
                    index=False,
                    engine="openpyxl",
                )

                nombre_proveedor = (
                    proveedor.strip()
                    if proveedor
                    else "proveedor"
                )

                nombre_proyecto = (
                    proyecto.strip()
                    if proyecto
                    else "proyecto"
                )

                nombre_proveedor = re.sub(
                    r"[^a-zA-Z0-9_-]+",
                    "_",
                    nombre_proveedor,
                )

                nombre_proyecto = re.sub(
                    r"[^a-zA-Z0-9_-]+",
                    "_",
                    nombre_proyecto,
                )

                st.download_button(
                    "Descargar resultado de la revisión",
                    data=buffer.getvalue(),
                    file_name=(
                        f"revision_{nombre_proveedor}_"
                        f"{nombre_proyecto}.xlsx"
                    ),
                    mime=(
                        "application/vnd.openxmlformats-"
                        "officedocument.spreadsheetml.sheet"
                    ),
                )

                # ==========================================================
                # GRÁFICA RESUMEN (mismo conteo de arriba, en dona con %)
                # ==========================================================

                orden_categorias = [
                    "ALTO",
                    "BAJO",
                    "EN MERCADO",
                    "SIN DATOS SUFICIENTES",
                ]

                nombres_categorias = {
                    "ALTO": "Alto",
                    "BAJO": "Bajo",
                    "EN MERCADO": "En mercado",
                    "SIN DATOS SUFICIENTES": "Sin datos",
                }

                # Mismos colores que ya se usan para resaltar la tabla,
                # asi la grafica se ve consistente con "RESULTADO FINAL".
                colores_categorias = {
                    "ALTO": "#f7c1c1",
                    "EN MERCADO": "#ffe699",
                    "BAJO": "#c6e0b4",
                    "SIN DATOS SUFICIENTES": "#d9d9d9",
                }

                etiquetas_grafica = []
                valores_grafica = []
                colores_grafica = []

                for categoria in orden_categorias:
                    conteo = int(
                        resumen_veredictos.get(
                            categoria, 0
                        )
                    )
                    if conteo > 0:
                        etiquetas_grafica.append(
                            nombres_categorias[categoria]
                        )
                        valores_grafica.append(conteo)
                        colores_grafica.append(
                            colores_categorias[categoria]
                        )

                if valores_grafica:

                    st.subheader(
                        "Distribución de resultados"
                    )

                    figura_resumen = go.Figure(
                        data=[
                            go.Pie(
                                labels=etiquetas_grafica,
                                values=valores_grafica,
                                marker=dict(
                                    colors=colores_grafica,
                                    line=dict(
                                        color="#ffffff",
                                        width=2,
                                    ),
                                ),
                                textinfo="label+percent+value",
                                hole=0.45,
                            )
                        ]
                    )

                    figura_resumen.update_layout(
                        showlegend=True,
                        margin=dict(
                            t=10, b=10, l=10, r=10
                        ),
                        height=380,
                    )

                    st.plotly_chart(
                        figura_resumen,
                        use_container_width=True,
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
