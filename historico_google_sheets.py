"""
Historico interno respaldado en Google Sheets (memoria persistente de verdad)
================================================================================

Misma idea que historico_interno.py, pero en vez de guardar en un Excel local
(que se pierde entre sesiones de Colab/Streamlit), lee y escribe directo en
una Google Sheet - asi la memoria sobrevive sin importar donde corras el
codigo (Colab, tu compu, o la app en Streamlit Cloud).

Requiere:
    pip install gspread google-auth

Como se autentica (dos formas, usa la que aplique):

  A) Local / Colab: descarga el JSON de la cuenta de servicio y pasa la RUTA
     del archivo:
         hist = HistoricoGoogleSheets(sheet_id="TU_SHEET_ID", creds_path="credenciales.json")

  B) Streamlit Cloud: el JSON se guarda como "secret" (Settings -> Secrets) y
     se lee como diccionario:
         hist = HistoricoGoogleSheets(sheet_id="TU_SHEET_ID", creds_dict=st.secrets["gcp_service_account"])

La Google Sheet debe:
  - tener una pestaña llamada "Historico"
  - tener en la fila 1 estos encabezados exactos:
    fecha_carga | proyecto | proveedor | concepto | concepto_norm | unidad | unidad_norm | precio_unitario | cluster_id
  - estar compartida (permiso Editor) con el correo de la cuenta de servicio
    (algo como capex-sheets-bot@TU-PROYECTO.iam.gserviceaccount.com)
"""

import re
from datetime import datetime

import gspread
import pandas as pd
from google.oauth2.service_account import Credentials
from rapidfuzz import process

# Se reutiliza la MISMA normalizacion de texto/unidades y el mismo scorer
# combinado que usa comparador_multifuente_v2.py (en vez de mantener una
# copia separada aqui). Antes este archivo tenia su propia version de
# normalize_text/normalize_unit sin el arreglo de sinonimos de unidad
# (PZA/PIEZA/PZ/UN, ML/M/MTS/METRO, M2/M², etc.) ni la expansion de
# abreviaturas ("SUM"->"SUMINISTRO"...), asi que el historico interno de
# Ragasa sufria el mismo problema de "no encuentra coincidencias" que las
# fuentes NL/CDMX. Tener una sola copia evita que se vuelvan a desincronizar.
from comparador_multifuente_v2 import (
    normalize_text,
    normalize_unit,
    score_combinado,
    nivel_confianza,
    UMBRAL_CONFIANZA_BAJA,
    es_diferencia_extrema,
    banda_en_mercado,
)
from rapidfuzz import fuzz  # se sigue usando fuzz.token_set_ratio en _homologar_uno
# Nota: ya no se importa revision_ia aqui -- la revision con IA de matches
# riesgosos de este historico se hace en LOTE desde app_v2.py, no adentro
# de consultar().

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

COLUMNAS = ['fecha_carga', 'proyecto', 'proveedor', 'concepto', 'concepto_norm',
            'unidad', 'unidad_norm', 'precio_unitario', 'cluster_id',
            'resultado_final', 'diferencia_pct']

# Columnas que se agregaron despues de que el histórico ya estaba en uso.
# Si algun llamador (o una hoja vieja) no las trae, se rellenan en blanco
# en vez de tronar -- asi una cotizacion guardada sin pasar por la
# revision completa (ALTO/BAJO/% diferencia) no rompe el guardado.
COLUMNAS_OPCIONALES = ['resultado_final', 'diferencia_pct']

NUM_RE = re.compile(r'\d+(?:\.\d+)?')


def _es_numero(v) -> bool:
    try:
        float(v)
        return v == v
    except (TypeError, ValueError):
        return False


def numeric_signature(t: str):
    nums = NUM_RE.findall(t)
    return tuple(sorted(round(float(n), 1) for n in nums))


def _to_native(v):
    """Convierte tipos numpy/pandas (int64, float64, Timestamp, NaN) a tipos
    nativos de Python para que gspread pueda serializarlos a JSON."""
    if v is None:
        return ""
    if isinstance(v, pd.Timestamp):
        return v.strftime("%Y-%m-%d")
    if hasattr(v, "item"):  # numpy int64/float64/bool_
        v = v.item()
    try:
        if pd.isna(v):
            return ""
    except (TypeError, ValueError):
        pass
    return v


def clasificar(precio, low, high):
    if precio < low:
        return 'BAJO'
    if precio > high:
        return 'ALTO'
    return 'EN MERCADO'


class HistoricoGoogleSheets:
    def __init__(self, sheet_id: str, creds_path: str = None, creds_dict: dict = None,
                 sim_threshold: float = 72.0):
        if creds_path:
            creds = Credentials.from_service_account_file(creds_path, scopes=SCOPES)
        elif creds_dict:
            creds = Credentials.from_service_account_info(dict(creds_dict), scopes=SCOPES)
        else:
            raise ValueError("Pasa creds_path (archivo JSON local) o creds_dict (st.secrets en Streamlit)")

        self.client = gspread.authorize(creds)
        self.sheet = self.client.open_by_key(sheet_id).worksheet("Historico")
        self.sim_threshold = sim_threshold
        self._cargar()

    def _cargar(self):
        # OJO: se pide el valor SIN formatear (UNFORMATTED_VALUE) a proposito.
        # gspread.get_all_records() por defecto trae el valor ya "formateado"
        # como texto y luego intenta convertirlo el mismo a numero -- pero
        # esa conversion asume formato tipo "11,785.5" (coma de miles, punto
        # decimal). Esta hoja esta en configuracion regional Espanol/Mexico,
        # donde un precio como 11785.5 se MUESTRA como "11785,5" (coma
        # decimal). gspread interpreta esa coma como separador de miles y la
        # descarta, convirtiendo "11785,5" en 117855 -- un precio 10 veces
        # mas alto que el real. Pedir el valor sin formatear evita ese
        # problema por completo: se recibe el numero real (11785.5) tal cual
        # esta guardado, sin pasar por ninguna interpretacion de texto.
        registros = self.sheet.get_all_records(
            value_render_option="UNFORMATTED_VALUE"
        )
        if registros:
            self.df = pd.DataFrame(registros)
            for col in COLUMNAS:
                if col not in self.df.columns:
                    self.df[col] = None
            self.df['cluster_id'] = pd.to_numeric(self.df['cluster_id'], errors='coerce').fillna(0).astype(int)
            self.df['precio_unitario'] = pd.to_numeric(self.df['precio_unitario'], errors='coerce')
            self.df['diferencia_pct'] = pd.to_numeric(self.df['diferencia_pct'], errors='coerce')
        else:
            self.df = pd.DataFrame(columns=COLUMNAS)
        self._next_cluster_id = (self.df['cluster_id'].max() + 1) if len(self.df) else 0

    def _homologar_uno(self, concepto_norm: str, unidad_norm: str) -> int:
        firma = numeric_signature(concepto_norm)
        candidatos = self.df[self.df['unidad_norm'] == unidad_norm]
        if not candidatos.empty:
            candidatos = candidatos.copy()
            candidatos['_firma'] = candidatos['concepto_norm'].map(numeric_signature)
            candidatos = candidatos[candidatos['_firma'] == firma]
        if candidatos.empty:
            cid = self._next_cluster_id
            self._next_cluster_id += 1
            return cid
        mejor_cid, mejor_score = None, -1
        for cid, grupo in candidatos.groupby('cluster_id'):
            rep = grupo['concepto_norm'].mode().iloc[0]
            score = fuzz.token_set_ratio(concepto_norm, rep)
            if score > mejor_score:
                mejor_score, mejor_cid = score, cid
        if mejor_score >= self.sim_threshold:
            return int(mejor_cid)
        cid = self._next_cluster_id
        self._next_cluster_id += 1
        return cid

    def ingerir(self, datos, proveedor: str, proyecto: str) -> pd.DataFrame:
        """Agrega las partidas de una cotizacion/licitacion nueva y las escribe
        de inmediato en la Google Sheet (persistente)."""
        if isinstance(datos, pd.DataFrame):
            nuevo = datos.copy()
        elif str(datos).lower().endswith('.csv'):
            nuevo = pd.read_csv(datos)
        else:
            nuevo = pd.read_excel(datos)

        faltantes = {'concepto', 'unidad', 'precio_unitario'} - set(nuevo.columns)
        if faltantes:
            raise ValueError(f"Faltan columnas obligatorias: {faltantes}")

        for col in COLUMNAS_OPCIONALES:
            if col not in nuevo.columns:
                nuevo[col] = None

        nuevo['concepto_norm'] = nuevo['concepto'].map(normalize_text)
        nuevo['unidad_norm'] = nuevo['unidad'].map(normalize_unit)
        nuevo['proveedor'] = proveedor
        nuevo['proyecto'] = proyecto
        nuevo['fecha_carga'] = datetime.now().strftime('%Y-%m-%d')

        filas_nuevas = []
        existentes = set()
        if not self.df.empty:
            existentes = {
                (str(a), str(b), str(c), str(d), round(float(e or 0), 2))
                for a, b, c, d, e in self.df[['proyecto', 'proveedor', 'concepto_norm', 'unidad_norm',
                                              'precio_unitario']].itertuples(index=False)
                if _es_numero(e)
            }
        for _, r in nuevo.iterrows():
            llave = (str(proyecto), str(proveedor), str(r['concepto_norm']), str(r['unidad_norm']),
                     round(float(r['precio_unitario'] or 0), 2) if _es_numero(r['precio_unitario']) else None)
            if llave in existentes:
                continue   # ya estaba guardada: no se duplica
            existentes.add(llave)
            cid = self._homologar_uno(r['concepto_norm'], r['unidad_norm'])
            fila = {**r.to_dict(), 'cluster_id': cid}
            self.df = pd.concat([self.df, pd.DataFrame([fila])[COLUMNAS]], ignore_index=True)
            filas_nuevas.append([_to_native(fila[c]) for c in COLUMNAS])

        # escribe solo las filas nuevas al final de la hoja (rapido, no reescribe todo)
        if filas_nuevas:
            self.sheet.append_rows(filas_nuevas, value_input_option="USER_ENTERED")
        return pd.DataFrame(filas_nuevas, columns=COLUMNAS)

    def consultar(self, descripcion: str, unidad: str, precio_cotizado: float = None,
                  min_score: float = 78.0, usar_ia: bool = False, excluir=None) -> dict:
        """excluir=(proveedor, proyecto): no compara la cotización contra sí misma."""
        t = normalize_text(descripcion)
        u = normalize_unit(unidad)
        pool = self.df[self.df['unidad_norm'] == u]
        if excluir and all(excluir) and not pool.empty:
            pool = pool[~((pool['proveedor'].astype(str) == str(excluir[0]))
                          & (pool['proyecto'].astype(str) == str(excluir[1])))]
        if pool.empty:
            return {'match': None, 'motivo': f'historico vacio o sin unidad {u} todavia'}

        choices = pool['concepto_norm'].tolist()
        # Misma búsqueda por material + trabajo + unidad que NL/CDMX: el
        # texto original, su versión de catálogo y las equivalencias
        # ("columnas para barda" = castillo; "cerramiento" = dala/cadena).
        # Se rechazan elementos, materiales o trabajos distintos.
        try:
            from comparador_multifuente_v2 import consultas_catalogo, _referencia_compatible
            consultas = consultas_catalogo(t)
        except Exception:
            consultas, _referencia_compatible = [t], (lambda a, b: True)

        def _mejor(umbral):
            mejor = None
            for consulta in consultas:
                for _, score, idx in process.extract(consulta, choices, scorer=score_combinado,
                                                     score_cutoff=umbral, limit=10):
                    if not _referencia_compatible(t, choices[idx]):
                        continue
                    if mejor is None or score > mejor[1]:
                        mejor = (choices[idx], score, idx)
                    break
            return mejor

        result = _mejor(min_score)
        confianza = None
        if result is None and min_score > UMBRAL_CONFIANZA_BAJA:
            # Segunda pasada con umbral relajado: mejor ofrecer una
            # referencia debil marcada "revisar" que ninguna.
            result = _mejor(UMBRAL_CONFIANZA_BAJA)
            confianza = 'BAJA' if result is not None else None
        if result is None:
            return {'match': None, 'motivo': f'sin coincidencia ni relajada en el historico (unidad {u})'}
        _, score, idx = result
        if confianza is None:
            confianza = nivel_confianza(float(score))
        cid = pool.iloc[idx]['cluster_id']
        grupo = self.df[self.df['cluster_id'] == cid]

        out = {
            'match': grupo['concepto'].mode().iloc[0], 'score': round(float(score), 1),
            'confianza': confianza,
            'n_registros': len(grupo), 'n_proveedores': grupo['proveedor'].nunique(),
            'proveedores': sorted(grupo['proveedor'].unique().tolist()),
            'precio_min': float(grupo['precio_unitario'].min()),
            'precio_mediana': float(grupo['precio_unitario'].median()),
            'precio_max': float(grupo['precio_unitario'].max()),
            'fecha_dato': str(grupo['fecha_carga'].dropna().astype(str).max() or '') if 'fecha_carga' in grupo else '',
        }
        if precio_cotizado is not None and len(grupo) >= 2:
            # "EN MERCADO" = dentro de +/-5% de la mediana del clúster (ver
            # MARGEN_EN_MERCADO en comparador_multifuente_v2.py), igual que
            # NL y CDMX -- antes esto usaba percentiles p25-p75, que daban
            # una banda distinta segun que tan dispersos estuvieran los
            # precios de ese concepto en el histórico.
            low, high = banda_en_mercado(out['precio_mediana'])
            out['clasificacion'] = clasificar(precio_cotizado, low, high)
        elif precio_cotizado is not None:
            out['clasificacion'] = None
            out['nota'] = 'solo 1 registro historico: referencia puntual, no banda estadistica'

        diferencia_extrema = (
            precio_cotizado is not None
            and es_diferencia_extrema(precio_cotizado, out['precio_mediana'])
        )

        if confianza == 'BAJA':
            out['motivo'] = (
                'coincidencia debil (revisar a mano): el texto no es muy '
                'parecido, confirma que sea el mismo concepto antes de '
                'confiar en este precio de referencia.'
            )
        elif diferencia_extrema:
            out['motivo'] = (
                'precio muy alejado de la referencia (revisar a mano): '
                'el texto coincide bien, pero el precio esta a una '
                'distancia inusual -- confirma que sea el mismo producto, '
                'tamano y calibre antes de confiar en esta clasificacion.'
            )

        # OJO: aqui YA NO se llama a la IA una partida a la vez -- eso se
        # hace en LOTE desde app_v2.py (ver revision_ia.
        # revisar_coincidencias_debiles_lote) despues de consultar TODAS
        # las partidas de la cotizacion, para no tardar minutos ni agotar
        # el limite de solicitudes por minuto. El parametro usar_ia se deja
        # en la firma por compatibilidad, pero ya no dispara ninguna
        # llamada aqui adentro.

        return out

    def resumen(self) -> dict:
        return {
            'total_renglones': len(self.df),
            'conceptos_homologados': self.df['cluster_id'].nunique() if len(self.df) else 0,
            'proveedores': sorted(self.df['proveedor'].dropna().unique().tolist()) if len(self.df) else [],
            'proyectos': sorted(self.df['proyecto'].dropna().unique().tolist()) if len(self.df) else [],
        }
