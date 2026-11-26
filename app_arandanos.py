import colorsys
import hashlib
import json
import re
import unicodedata
from pathlib import Path

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components

# =============================================================================
# CONFIGURACIÓN (todo lo que probablemente quieras ajustar está aquí arriba)
# =============================================================================
ARCHIVO_EXCEL = "MAPEO ARANDANOS.xlsx"       # También acepta MAPEO_ARANDANOS.xlsx (guion bajo)
ARCHIVO_RECEPCION = "BD_IN_26.xlsx"          # Base de recepción: lote, fecha guia, procesado, rechazo
# Pallets cuya columna "Tipo" contenga alguna de estas palabras NO pertenecen al mapeo de fruta embalada
# (IQF, IQF Jugo, Fresco sin proceso). Se excluyen y se informan en «Calidad de Datos».
PALABRAS_TIPO_EXCLUIDO = ["iqf", "fresco"]
# Valores de la columna "rechazo" de la base de recepción que significan "NO rechazado"
# (la columna puede venir vacía o con un guion en los pallets normales). Cualquier otro texto = rechazado.
RECHAZO_SIN_VALOR = {"", "-", "--", "---", ".", "nan", "none", "no", "n/a", "na", "ninguno", "sin rechazo", "0"}
PALABRA_BULK = "bulk"                        # Formato de recepción que contiene esto = bulk (no se procesa)
DIAS_ALERTA_DEFECTO = 7          # Alerta si la permanencia es MAYOR a este número de días
                                 # (el día de "Fecha Creacion" cuenta como día 1)
ARCHIVO_AJUSTES = "AJUSTES FECHA.csv"   # Fechas originales de pallets cuya Fecha Creacion fue cambiada
                                        # (se crea y edita desde la pestaña «Ajustes de Fecha»)

# Estructura física de las cámaras codificadas (C1, C2, C3 ...)
HILERAS = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M']  # 13 hileras
PROFUNDIDAD_MAX = 5   # módulos laterales
ALTURA_MAX = 4        # niveles de altura
SECTORES_NOMBRES = {'H': 'Sector H (Huape)', 'O': 'Sector O (Oro Verde)'}
POSICIONES_POR_CAMARA_DEFECTO = len(HILERAS) * len(SECTORES_NOMBRES) * PROFUNDIDAD_MAX * ALTURA_MAX  # 520
CAPACIDAD_POR_CAMARA = {}   # Opcional, ej: {"C2": 400} si alguna cámara tiene otra capacidad

# Nombres con los que se mostrarán ubicaciones especiales (lo que aparece en la columna Posicion)
ALIAS_UBICACIONES = {'X': 'X (Sin mapear)'}
ETIQUETA_SIN_POSICION = 'SIN POSICIÓN'

# Nombres posibles de cada columna del Excel (no distingue mayúsculas, tildes ni espacios extra)
CANDIDATOS_COLUMNAS = {
    'Posicion':    ['posicion', 'ubicacion'],
    'Pallet':      ['pallet', 'nro pallet', 'n pallet', 'numero pallet', 'n° pallet'],
    'Fecha':       ['fecha creacion', 'fecha de creacion', 'creacion'],
    'Exportadora': ['exportadora', 'cliente', 'productor'],
    'Formato':     ['formato', 'envase', 'embalaje', 'tipo envase'],
    'Variedad':    ['variedad'],
    'Lote':        ['lote'],
    'Tipo':        ['tipo'],
    'Cantidad':    ['cantidad', 'cajas'],
    'KN':          ['kn'],
}

PATRON_POSICION = re.compile(r'^C(\d+)([HO])([A-Z])(\d+)(\d+)$')  # ej: C1HA11, C2OA24

# =============================================================================
# CONFIGURACIÓN DE PÁGINA
# =============================================================================
st.set_page_config(
    page_title="WMS 3D - Mapeo de Arándanos",
    page_icon="🫐",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.title("🫐 Mapeo Arándanos FSJ 🫐")
st.markdown("Los latidos representan posiciones con más de un pallet · El borde naranja marca pallets con alerta de permanencia")


# =============================================================================
# FUNCIONES AUXILIARES
# =============================================================================
def norm(txt):
    """Normaliza un texto: sin tildes, minúsculas, sin espacios repetidos."""
    t = unicodedata.normalize('NFKD', str(txt)).encode('ascii', 'ignore').decode()
    return re.sub(r'\s+', ' ', t).strip().lower()


def limpiar_pallet(v):
    """Convierte el número de pallet a texto limpio (12345.0 -> '12345')."""
    if pd.isna(v):
        return ''
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return re.sub(r'\.0*$', '', str(v).strip())   # '2027000033.' -> '2027000033'


def es_rechazo(v):
    """True si la celda de la columna 'rechazo' indica un rechazo real (vacío o '-' = no rechazado)."""
    if pd.isna(v):
        return False
    return norm(v) not in RECHAZO_SIN_VALOR


def clasificar_posicion(pos):
    """Devuelve la cámara y, si la posición está codificada, sus coordenadas."""
    m = PATRON_POSICION.match(pos)
    if m:
        return {
            'Camara': f"C{int(m.group(1))}",
            'Codificada': True,
            'Sector': m.group(2),
            'Hilera': m.group(3),
            'Nivel': int(m.group(4)),
            'Modulo': int(m.group(5)),
        }
    etiqueta = ALIAS_UBICACIONES.get(pos, pos)
    return {'Camara': etiqueta, 'Codificada': False, 'Sector': None,
            'Hilera': None, 'Nivel': None, 'Modulo': None}


def orden_camara(c):
    """C1, C2, C3... primero; luego otras cámaras (SNP, R4...); al final X y sin posición."""
    if re.fullmatch(r'C\d+', c):
        return (0, int(c[1:]), c)
    if c in (ETIQUETA_SIN_POSICION,) or c in ALIAS_UBICACIONES.values():
        return (2, 0, c)
    return (1, 0, c)


def ordenar_camaras(lista):
    return sorted(lista, key=orden_camara)


def color_estable(texto):
    """Color fijo por texto (no cambia cada vez que se reinicia la app)."""
    h = int(hashlib.md5(str(texto).encode('utf-8')).hexdigest()[:8], 16)
    r, g, b = colorsys.hls_to_rgb((h % 360) / 360, 0.55, 0.65)
    return f"#{int(r * 255):02x}{int(g * 255):02x}{int(b * 255):02x}"


def mostrar_tabla(df, **kwargs):
    """Compatible con versiones nuevas y antiguas de Streamlit."""
    try:
        st.dataframe(df, width='stretch', **kwargs)
    except TypeError:
        st.dataframe(df, use_container_width=True, **kwargs)


def deduplicar_pallets(df):
    """
    Deja UNA sola fila por número de pallet (el mismo pallet puede venir repetido en el Excel).
    Si un pallet aparece varias veces se conserva, en este orden de preferencia:
      1) la fila con posición codificada (C1HA11...), 2) la de fecha de creación más antigua.
    Las filas sin número de pallet se excluyen del conteo.
    """
    d = df[df['Pallet'] != ''].copy()
    d['_cod'] = d['Codificada'].astype(int)
    d = d.sort_values(['_cod', 'Fecha'], ascending=[False, True], na_position='last')
    return d.drop_duplicates('Pallet', keep='first').drop(columns='_cod')


def ruta_ajustes():
    return Path(__file__).resolve().parent / ARCHIVO_AJUSTES


def cargar_ajustes():
    """Lee la lista de ajustes de fecha: Pallet | Fecha Original | Motivo."""
    vacio = pd.DataFrame({'Pallet': pd.Series(dtype=str),
                          'Fecha Original': pd.Series(dtype='datetime64[ns]'),
                          'Motivo': pd.Series(dtype=str)})
    ruta = ruta_ajustes()
    if not ruta.exists():
        return vacio
    try:
        a = pd.read_csv(ruta, dtype=str, encoding='utf-8-sig').fillna('')
    except Exception:
        return vacio
    a.columns = [norm(c) for c in a.columns]
    if 'pallet' not in a.columns or 'fecha original' not in a.columns:
        return vacio
    out = pd.DataFrame({
        'Pallet': a['pallet'].str.strip(),
        'Fecha Original': pd.to_datetime(a['fecha original'], errors='coerce', dayfirst=True).dt.normalize(),
        'Motivo': a['motivo'] if 'motivo' in a.columns else '',
    })
    out = out[(out['Pallet'] != '') & out['Fecha Original'].notna()]
    return out.drop_duplicates('Pallet', keep='last').reset_index(drop=True)


def guardar_ajustes(df):
    """Guarda la tabla de ajustes en el CSV. Devuelve cuántos ajustes quedaron guardados."""
    d = df.copy()
    d['Pallet'] = d['Pallet'].fillna('').astype(str).str.strip().str.replace(r'\.0$', '', regex=True)
    d['Fecha Original'] = pd.to_datetime(d['Fecha Original'], errors='coerce').dt.normalize()
    d['Motivo'] = d['Motivo'].fillna('').astype(str).str.strip()
    d = d[(d['Pallet'] != '') & d['Fecha Original'].notna()].drop_duplicates('Pallet', keep='last')
    salida = d.copy()
    salida['Fecha Original'] = salida['Fecha Original'].dt.strftime('%d/%m/%Y')
    salida.to_csv(ruta_ajustes(), index=False, encoding='utf-8-sig')
    return len(d)


# =============================================================================
# CARGA DE DATOS
# =============================================================================
def buscar_archivo(nombre):
    """Busca el archivo junto al .py o en la carpeta de trabajo (acepta espacios o guiones bajos)."""
    def clave(n):
        return re.sub(r'[\s_]+', ' ', norm(Path(n).stem))
    carpetas = [Path(__file__).resolve().parent, Path.cwd()]
    for c in carpetas:
        if (c / nombre).exists():
            return c / nombre
    objetivo, ext = clave(nombre), Path(nombre).suffix.lower()
    for c in carpetas:
        for f in sorted(c.iterdir()):
            if f.is_file() and f.suffix.lower() == ext and not f.name.startswith('~$') and clave(f.name) == objetivo:
                return f
    return None


def ubicar_excel():
    return buscar_archivo(ARCHIVO_EXCEL)


@st.cache_data(ttl=60)
def cargar_recepcion():
    """
    Lee la base de recepción y devuelve (lotes, aviso). Una fila por lote.
    - Los pallets rechazados no cuentan para las fechas del lote.
    - Si TODOS los pallets del lote están rechazados, el lote queda marcado Rechazado_Total.
    - Recepcion = fecha de recepción más antigua; Proceso = fecha de proceso más antigua (NaT si no hay).
    """
    ruta = buscar_archivo(ARCHIVO_RECEPCION)
    if ruta is None:
        return pd.DataFrame(), (f"No se encontró '{ARCHIVO_RECEPCION}'. Sin ella la antigüedad se calcula "
                                f"solo con la Fecha Creacion del mapeo.")
    try:
        raw = pd.read_excel(ruta)
    except Exception as e:
        return pd.DataFrame(), f"Error al leer '{ARCHIVO_RECEPCION}': {e}"

    cn = {norm(c): c for c in raw.columns}

    def col(*cands):
        for c in cands:
            if c in cn:
                return cn[c]
        return None

    c_lote = col('lote')
    c_rec = col('fecha guia', 'fecha recepcion', 'fecha de recepcion')
    c_proc = col('procesado', 'fecha proceso', 'fecha de proceso')
    if None in (c_lote, c_rec, c_proc):
        return pd.DataFrame(), (f"La base de recepción debe tener las columnas lote, fecha guia y procesado. "
                                f"Columnas encontradas: {list(raw.columns)}")
    c_rech, c_exp = col('rechazo', 'rechazado'), col('exportadora')
    c_for, c_var = col('formato/recip', 'formato'), col('variedad')

    def texto(c):
        if c is None:
            return pd.Series('Sin dato', index=raw.index)
        return raw[c].astype(str).str.strip().where(raw[c].notna(), 'Sin dato')

    r = pd.DataFrame({
        'Lote': raw[c_lote].map(limpiar_pallet),
        'Recepcion': pd.to_datetime(raw[c_rec], errors='coerce', dayfirst=True).dt.normalize(),
        'Proceso': pd.to_datetime(raw[c_proc], errors='coerce', dayfirst=True).dt.normalize(),
        'Rechazado': raw[c_rech].map(es_rechazo) if c_rech else pd.Series(False, index=raw.index),
        'Exportadora': texto(c_exp),
        'Variedad': texto(c_var),
        'Bulk': texto(c_for).str.lower().str.contains(PALABRA_BULK, regex=False),
    })
    r = r[r['Lote'] != '']

    todos = r.groupby('Lote').agg(Exportadora=('Exportadora', 'first'), Variedad=('Variedad', 'first'),
                                  Bulk=('Bulk', 'any'), Pallets_Total=('Lote', 'size'),
                                  Pallets_Rechazados=('Rechazado', 'sum'), Recepcion_Todos=('Recepcion', 'min'))
    ok = r[~r['Rechazado']].groupby('Lote').agg(Recepcion=('Recepcion', 'min'), Proceso=('Proceso', 'min'),
                                                Procesos_Distintos=('Proceso', 'nunique'),
                                                Pallets_OK=('Lote', 'size'))
    aviso = None
    if r['Rechazado'].mean() > 0.5:
        aviso = (f"⚠️ {r['Rechazado'].mean():.0%} de las filas de la base de recepción figuran como rechazadas. "
                 f"Revisa los valores de la columna 'rechazo' (valores sin rechazo reconocidos: "
                 f"{sorted(v for v in RECHAZO_SIN_VALOR if v)}).")
    lotes = todos.join(ok, how='left')
    lotes['Rechazado_Total'] = lotes['Pallets_OK'].isna()
    lotes['Recepcion'] = lotes['Recepcion'].fillna(lotes['Recepcion_Todos'])
    lotes['Pallets_OK'] = lotes['Pallets_OK'].fillna(0).astype(int)
    lotes['Procesos_Distintos'] = lotes['Procesos_Distintos'].fillna(0).astype(int)
    lotes['Pallets_Rechazados'] = lotes['Pallets_Rechazados'].astype(int)
    return lotes.drop(columns='Recepcion_Todos').reset_index(), aviso


@st.cache_data(ttl=60)
def cargar_datos():
    """Lee el Excel y devuelve (df, columnas_detectadas, avisos, error)."""
    ruta = ubicar_excel()
    if ruta is None:
        return pd.DataFrame(), {}, [], f"No se encontró '{ARCHIVO_EXCEL}'. Debe estar en la misma carpeta que este archivo .py"
    try:
        raw = pd.read_excel(ruta)
    except Exception as e:
        return pd.DataFrame(), {}, [], f"Error al leer el Excel: {e}"

    raw.columns = [str(c).strip() for c in raw.columns]
    cols_norm = {norm(c): c for c in raw.columns}
    mapa, avisos = {}, []
    for canon, candidatos in CANDIDATOS_COLUMNAS.items():
        for cand in candidatos:
            if cand in cols_norm:
                mapa[canon] = cols_norm[cand]
                break

    for obligatoria in ('Posicion', 'Pallet'):
        if obligatoria not in mapa:
            return pd.DataFrame(), mapa, avisos, (
                f"No se encontró la columna '{obligatoria}'. Columnas del Excel: {list(raw.columns)}")

    if 'Formato' not in mapa and 'Tipo' in mapa:
        mapa['Formato'] = mapa['Tipo']
        avisos.append("No existe columna 'Formato'; se usó la columna 'Tipo' como formato.")
    if 'Formato' not in mapa:
        avisos.append("No existe columna 'Formato' ni 'Tipo'; todos los pallets quedan como 'Sin formato'.")
    if 'Fecha' not in mapa:
        avisos.append("No existe columna 'Fecha Creacion'; las alertas de permanencia están desactivadas.")

    def texto(canon, defecto):
        if canon not in mapa:
            return pd.Series(defecto, index=raw.index)
        serie = raw[mapa[canon]]
        s = serie.astype(str).str.strip()
        s = s.where(serie.notna(), defecto)
        return s.replace({'': defecto, 'nan': defecto, 'NaT': defecto})

    pos_serie = raw[mapa['Posicion']]
    pos = pos_serie.astype(str).str.upper().str.replace(r'\s+', '', regex=True)
    pos = pos.where(pos_serie.notna() & ~pos.isin(['', 'NAN']), ETIQUETA_SIN_POSICION)

    df = pd.DataFrame(index=raw.index)
    df['Pallet'] = raw[mapa['Pallet']].map(limpiar_pallet)
    df['Posicion'] = pos
    df['Exportadora'] = texto('Exportadora', 'Sin exportadora')
    df['Formato'] = texto('Formato', 'Sin formato')
    df['Lote'] = raw[mapa['Lote']].map(limpiar_pallet).replace('', '—') if 'Lote' in mapa else '—'
    df['Tipo'] = texto('Tipo', '')
    variedad = texto('Variedad', '—')
    df['Variedad'] = variedad.str.split('(').str[0].str.strip()
    df['Cantidad'] = (pd.to_numeric(raw[mapa['Cantidad']], errors='coerce').fillna(0)
                      if 'Cantidad' in mapa else 0)
    df['KN'] = texto('KN', '')
    if 'Fecha' in mapa:
        df['Fecha'] = pd.to_datetime(raw[mapa['Fecha']], errors='coerce', dayfirst=True).dt.normalize()
    else:
        df['Fecha'] = pd.NaT

    info = pd.DataFrame(list(df['Posicion'].map(clasificar_posicion)), index=df.index)
    df = pd.concat([df, info], axis=1)

    # Posiciones codificadas con más de un pallet DISTINTO (sobre-apilado)
    cod = df[df['Codificada'] & (df['Pallet'] != '')]
    n_por_pos = cod.groupby('Posicion')['Pallet'].nunique()
    posiciones_multiples = set(n_por_pos[n_por_pos > 1].index)
    df['Sobre_apilado'] = df['Codificada'] & df['Posicion'].isin(posiciones_multiples)

    return df, mapa, avisos, None


df_all, mapa_columnas, avisos, error = cargar_datos()

if error:
    st.error(error)
    st.stop()
if df_all.empty:
    st.warning(f"No se encontraron registros en '{ARCHIVO_EXCEL}'.")
    st.stop()

# =============================================================================
# SIDEBAR: FILTROS GLOBALES
# =============================================================================
st.sidebar.header("🔍 Filtros")

exportadoras_disp = sorted(df_all['Exportadora'].unique())
exportadoras_sel = st.sidebar.multiselect("Exportadora", exportadoras_disp, default=exportadoras_disp)

dias_alerta = st.sidebar.number_input(
    "Alerta de permanencia (días mayor a)", min_value=1, max_value=60,
    value=DIAS_ALERTA_DEFECTO, step=1,
    help="El día de 'Fecha Creacion' cuenta como día 1. Con 7, la alerta se activa desde el día 8."
)

excluir_tipos = st.sidebar.checkbox(
    "Excluir IQF y fruta fresca sin proceso", value=True,
    help="Excluye los pallets cuyo 'Tipo' contiene: " + ", ".join(PALABRAS_TIPO_EXCLUIDO)
         + ". Si algún día necesitas verlos, desmarca esta opción."
)

with st.sidebar.expander("Columnas detectadas del Excel"):
    for canon, real in mapa_columnas.items():
        st.write(f"**{canon}** ← `{real}`")
for a in avisos:
    st.sidebar.warning(a)

if not exportadoras_sel:
    st.warning("Selecciona al menos una exportadora en el panel lateral.")
    st.stop()

# =============================================================================
# CÁLCULO DE PERMANENCIA Y ALERTAS
# =============================================================================
hoy = pd.Timestamp.now().normalize()

# --- Datos de recepción (llave: Lote) -------------------------------------------------
lotes_rec, aviso_rec = cargar_recepcion()
if aviso_rec:
    st.sidebar.warning(aviso_rec)

df_all['Fecha_Excel'] = df_all['Fecha']                        # Fecha Creacion tal como viene en el mapeo
if lotes_rec.empty:
    df_all['Lote_Rechazado'] = False
    df_all['Fecha_Recepcion'] = pd.NaT
    df_all['Fecha_Proceso'] = pd.NaT
else:
    L = lotes_rec.set_index('Lote')
    df_all['Lote_Rechazado'] = df_all['Lote'].map(L['Rechazado_Total']).fillna(False).astype(bool)
    df_all['Fecha_Recepcion'] = pd.to_datetime(df_all['Lote'].map(L['Recepcion']))
    df_all['Fecha_Proceso'] = pd.to_datetime(df_all['Lote'].map(L['Proceso']))

# Pallets de lotes 100% rechazados: se despachan, no deben figurar (se informan en "Calidad de Datos")
df_rech = df_all[df_all['Lote_Rechazado']].copy()
df_all = df_all[~df_all['Lote_Rechazado']].copy()

# Pallets de Tipo IQF / Fresco sin proceso: no pertenecen a este mapeo (se informan en "Calidad de Datos")
df_all['Excluido_Tipo'] = df_all['Tipo'].map(lambda t: any(w in norm(t) for w in PALABRAS_TIPO_EXCLUIDO))
if excluir_tipos:
    df_excl = df_all[df_all['Excluido_Tipo']].copy()
    df_all = df_all[~df_all['Excluido_Tipo']].copy()
else:
    df_excl = df_all.iloc[0:0].copy()

if df_all.empty:
    st.warning("No quedan registros después de excluir rechazados / IQF / fresco. "
               "Revisa la columna 'rechazo' de la base de recepción y la columna 'Tipo' del mapeo, "
               "o desmarca la exclusión en el panel lateral.")
    st.stop()

# Posiciones con más de un pallet DISTINTO (sobre-apilado), recalculado con los datos ya depurados
_cod = df_all[df_all['Codificada'] & (df_all['Pallet'] != '')]
_n = _cod.groupby('Posicion')['Pallet'].nunique()
df_all['Sobre_apilado'] = df_all['Codificada'] & df_all['Posicion'].isin(set(_n[_n > 1].index))

# Fecha base: proceso del lote; si está en blanco (bulk / pendiente) -> recepción
fecha_base = df_all['Fecha_Proceso'].fillna(df_all['Fecha_Recepcion'])
origen = pd.Series('Fecha Creación (lote sin datos de recepción)', index=df_all.index)
origen = origen.mask(df_all['Fecha_Recepcion'].notna(), 'Recepción del lote')
origen = origen.mask(df_all['Fecha_Proceso'].notna(), 'Proceso del lote')

# Si la Fecha Creacion del mapeo es ANTERIOR a la fecha base (posible error de digitación),
# se conserva la más antigua para no subestimar la antigüedad
excel_antes = df_all['Fecha_Excel'].notna() & fecha_base.notna() & (df_all['Fecha_Excel'] < fecha_base)
fecha_ef = fecha_base.fillna(df_all['Fecha_Excel']).where(~excel_antes, df_all['Fecha_Excel'])
origen = origen.mask(excel_antes, 'Fecha Creación (anterior al proceso/recepción)')

df_all['Fecha_Base'] = fecha_base
df_all['Excel_Antes_Base'] = excel_antes
df_all['Fecha'] = fecha_ef
df_all['Origen_Fecha'] = origen

# --- Ajustes manuales (excepciones): tienen prioridad sobre todo ---
ajustes = cargar_ajustes()
f_original = pd.to_datetime(df_all['Pallet'].map(dict(zip(ajustes['Pallet'], ajustes['Fecha Original']))))
df_all['Ajustada'] = f_original.notna()
df_all['Fecha'] = f_original.fillna(df_all['Fecha'])
df_all['Origen_Fecha'] = df_all['Origen_Fecha'].mask(df_all['Ajustada'], 'Ajuste manual')
df_all['Motivo_Ajuste'] = df_all['Pallet'].map(dict(zip(ajustes['Pallet'], ajustes['Motivo']))).fillna('')

df_all['Dias_Excel'] = (hoy - df_all['Fecha_Excel']).dt.days + 1   # Según Fecha Creacion del mapeo (referencia)
df_all['Dias'] = (hoy - df_all['Fecha']).dt.days + 1                # Efectiva: fecha base = día 1
df_all['Alerta'] = df_all['Dias'] > dias_alerta                     # NaN > n -> False

with st.sidebar.expander("Origen de la fecha de antigüedad"):
    _o = df_all.drop_duplicates('Pallet')['Origen_Fecha'].value_counts()
    for k, v in _o.items():
        st.write(f"**{v}** pallets · {k}")
if len(df_rech):
    st.sidebar.info(f"🚫 {df_rech['Pallet'].nunique()} pallet(s) de lotes rechazados excluidos")
if len(df_excl):
    st.sidebar.info(f"🧹 {df_excl['Pallet'].nunique()} pallet(s) IQF / fresco sin proceso excluidos")

# Filas filtradas por exportadora (todas las filas del Excel)
df_f = df_all[df_all['Exportadora'].isin(exportadoras_sel)].copy()

# Un solo registro por número de pallet (evita contar dos veces el mismo pallet)
df_u = deduplicar_pallets(df_all)
df_u = df_u[df_u['Exportadora'].isin(exportadoras_sel)].copy()

# =============================================================================
# PESTAÑAS
# =============================================================================
tab_dash, tab_3d, tab_alertas, tab_rec, tab_ajustes, tab_calidad = st.tabs(
    ["📊 Dashboard Resumen", "🧊 Mapa 3D", "⚠️ Alertas de Permanencia", "⏱️ Recepción → Proceso",
     "📝 Ajustes de Fecha", "🔎 Calidad de Datos"]
)

# -----------------------------------------------------------------------------
# TAB 1: DASHBOARD
# -----------------------------------------------------------------------------
with tab_dash:
    if len(df_excl) or len(df_rech):
        partes = []
        if len(df_excl):
            partes.append(f"{df_excl['Pallet'].nunique()} pallet(s) de tipo IQF / fresco sin proceso")
        if len(df_rech):
            partes.append(f"{df_rech['Pallet'].nunique()} pallet(s) de lotes rechazados")
        st.warning("🧹 Se excluyeron del análisis " + " y ".join(partes) +
                   ". Estos registros no deberían estar en el mapeo; revisa el detalle en «Calidad de Datos».")
    n_alertas = int(df_u['Alerta'].sum())
    if n_alertas > 0:
        st.error(f"⚠️ {n_alertas} pallet(s) superan los {dias_alerta} días de permanencia. "
                 f"Revisa la pestaña «Alertas de Permanencia».")

    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Pallets únicos", f"{len(df_u)}",
              help="Cada número de pallet se cuenta una sola vez, aunque aparezca repetido en el Excel.")
    k2.metric("En cámaras mapeadas", f"{int(df_u['Codificada'].sum())}",
              help="Pallets con posición codificada (C1HA11, C2OA24, etc.).")
    k3.metric("En otras ubicaciones", f"{int((~df_u['Codificada']).sum())}",
              help="SNP, X (sin mapear), R4, sin posición u otras ubicaciones sin coordenadas.")
    k4.metric("Con alerta de permanencia", f"{n_alertas}",
              delta=f"> {dias_alerta} días" if n_alertas else None, delta_color="inverse")

    if df_u.empty:
        st.info("No hay pallets para los filtros seleccionados.")
    else:
        st.markdown("### 🏭 Pallets por cámara / ubicación")
        camaras = ordenar_camaras(df_u['Camara'].unique())
        for i in range(0, len(camaras), 4):
            columnas = st.columns(4)
            for col, cam in zip(columnas, camaras[i:i + 4]):
                sub = df_u[df_u['Camara'] == cam]
                n_al = int(sub['Alerta'].sum())
                col.metric(cam, f"{len(sub)} pallets",
                           delta=f"{n_al} con alerta" if n_al else None, delta_color="inverse")

        st.markdown("### 📦 Pallets por cámara y formato")
        piv = df_u.pivot_table(index='Camara', columns='Formato', values='Pallet',
                               aggfunc='nunique', fill_value=0)
        piv = piv.reindex(ordenar_camaras(piv.index))
        piv.columns.name = None
        piv.index.name = 'Cámara'

        col_graf, col_tabla = st.columns([3, 2])
        with col_graf:
            st.bar_chart(piv)
        with col_tabla:
            piv_tabla = piv.copy()
            piv_tabla['TOTAL'] = piv_tabla.sum(axis=1)
            piv_tabla.loc['TOTAL GENERAL'] = piv_tabla.sum()
            mostrar_tabla(piv_tabla.astype(int))
            st.download_button("⬇️ Descargar tabla (CSV)",
                               piv_tabla.to_csv().encode('utf-8-sig'),
                               file_name="pallets_por_camara_formato.csv", mime="text/csv")

        st.markdown("### 🚢 Pallets por cámara y exportadora")
        piv_exp = df_u.pivot_table(index='Camara', columns='Exportadora', values='Pallet',
                                   aggfunc='nunique', fill_value=0)
        piv_exp = piv_exp.reindex(ordenar_camaras(piv_exp.index))
        piv_exp.columns.name = None
        piv_exp.index.name = 'Cámara'
        piv_exp['TOTAL'] = piv_exp.sum(axis=1)
        piv_exp.loc['TOTAL GENERAL'] = piv_exp.sum()
        mostrar_tabla(piv_exp.astype(int))

        # Ubicaciones sin coordenadas (SNP, X, R4, etc.)
        st.markdown("### 📍 Ubicaciones sin coordenadas (SNP, X, R4, otras)")
        st.caption("Cada número de pallet se cuenta una sola vez, aunque se repita en el Excel. "
                   "«Filas en Excel» muestra cuántas filas había realmente.")
        df_otras_u = df_u[~df_u['Codificada']]
        df_otras_f = df_f[~df_f['Codificada']]
        if df_otras_u.empty:
            st.info("No hay pallets en ubicaciones sin coordenadas.")
        else:
            resumen = (df_otras_u.groupby('Camara')
                       .agg(**{'Pallets únicos': ('Pallet', 'nunique'),
                               'Cajas': ('Cantidad', 'sum'),
                               'Con alerta': ('Alerta', 'sum')}))
            filas = df_otras_f.groupby('Camara').size().rename('Filas en Excel')
            resumen = resumen.join(filas).fillna(0)
            resumen['Filas repetidas'] = (resumen['Filas en Excel'] - resumen['Pallets únicos']).clip(lower=0)
            resumen = resumen.reindex(ordenar_camaras(resumen.index))
            resumen.index.name = 'Ubicación'
            mostrar_tabla(resumen.astype(int))

# -----------------------------------------------------------------------------
# TAB 2: MAPA 3D
# -----------------------------------------------------------------------------
HTML_TEMPLATE = """
<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8">
    <style>
        body { margin: 0; padding: 0; overflow: hidden; background-color: #111115; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; }
        #canvas-container { width: 100vw; height: 600px; position: relative; }
        #tooltip {
            position: absolute;
            background: rgba(15, 23, 42, 0.95);
            color: #f8fafc;
            padding: 14px 18px;
            border-radius: 8px;
            border: 1px solid #38bdf8;
            display: none;
            pointer-events: none;
            font-size: 13px;
            z-index: 100;
            box-shadow: 0 10px 15px -3px rgba(0, 0, 0, 0.5), 0 4px 6px -4px rgba(0, 0, 0, 0.5);
            line-height: 1.6;
        }
        #legend {
            position: absolute;
            bottom: 15px;
            left: 15px;
            background: rgba(15, 23, 42, 0.85);
            padding: 12px;
            color: #fff;
            font-size: 11px;
            border-radius: 6px;
            max-height: 220px;
            overflow-y: auto;
            border: 1px solid #334155;
        }
        .legend-item { display: flex; align-items: center; margin-bottom: 6px; }
        .legend-color { width: 12px; height: 12px; margin-right: 8px; border-radius: 3px; }
        .legend-note { margin-top: 8px; padding-top: 6px; border-top: 1px solid #334155; color: #cbd5e1; }
        #aisle-label {
            position: absolute;
            top: 15px;
            left: 50%;
            transform: translateX(-50%);
            background: rgba(15, 23, 42, 0.9);
            color: #38bdf8;
            padding: 8px 18px;
            border-radius: 20px;
            font-size: 13px;
            font-weight: bold;
            border: 1px solid #38bdf8;
            letter-spacing: 0.5px;
        }
    </style>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/three.js/r128/three.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/three@0.128.0/examples/js/controls/OrbitControls.js"></script>
</head>
<body>
    <div id="canvas-container">
        <div id="aisle-label">🚦 PASILLO CENTRAL LONGITUDINAL (LADO O ⟵  🛩️  ⟶ LADO H)</div>
        <div id="tooltip"></div>
        <div id="legend">
            <strong>__LEYENDA_TITULO__</strong>
            <div id="legend-items"></div>
            <div class="legend-note">
                🔴 Latido rojo: posición con más de un pallet<br/>
                🟠 Borde naranja: permanencia mayor a __DIAS_ALERTA__ días
            </div>
        </div>
    </div>

    <script>
        const palletsData = __THREEJS_DATA__;
        const container = document.getElementById('canvas-container');
        const scene = new THREE.Scene();
        scene.background = new THREE.Color(0x0f172a);

        const camera = new THREE.PerspectiveCamera(40, window.innerWidth / 600, 0.1, 1000);
        camera.position.set(0, 35, 75);

        const renderer = new THREE.WebGLRenderer({ antialias: true });
        renderer.setSize(window.innerWidth, 600);
        container.appendChild(renderer.domElement);

        const controls = new THREE.OrbitControls(camera, renderer.domElement);
        controls.enableDamping = true;
        controls.dampingFactor = 0.05;
        controls.maxPolarAngle = Math.PI / 2 - 0.05;
        controls.target.set(0, 4, 27);

        // Iluminación
        scene.add(new THREE.AmbientLight(0xffffff, 0.65));
        const dirLight1 = new THREE.DirectionalLight(0xffffff, 0.6);
        dirLight1.position.set(15, 60, 15);
        scene.add(dirLight1);

        const dirLight2 = new THREE.DirectionalLight(0x38bdf8, 0.4);
        dirLight2.position.set(-30, 20, -10);
        scene.add(dirLight2);

        // Grilla del piso
        const gridHelper = new THREE.GridHelper(150, 50, 0x38bdf8, 0x1e293b);
        gridHelper.position.y = -0.9;
        scene.add(gridHelper);

        // Pasillo central
        const laneGeo = new THREE.PlaneGeometry(7.5, 70);
        const laneMat = new THREE.MeshBasicMaterial({ color: 0x1e293b, side: THREE.DoubleSide });
        const lane = new THREE.Mesh(laneGeo, laneMat);
        lane.rotation.x = Math.PI / 2;
        lane.position.set(0, -0.88, 27);
        scene.add(lane);

        // Geometrías para racks y pallets
        const HILERAS_LIST = __HILERAS_JSON__;
        const MODULOS = __PROFUNDIDAD__;
        const NIVELES = __ALTURA__;
        const geometry = new THREE.BoxGeometry(2.2, 1.6, 2.2);
        const edgesGeometry = new THREE.EdgesGeometry(geometry);
        const rackGeometry = new THREE.BoxGeometry(2.5, 0.06, 2.5);
        const rackMaterial = new THREE.MeshStandardMaterial({ color: 0x475569, metalness: 0.8, roughness: 0.2 });

        function createTextTexture(text) {
            const canvas = document.createElement('canvas');
            canvas.width = 256;
            canvas.height = 128;
            const ctx = canvas.getContext('2d');
            ctx.fillStyle = '#1e293b';
            ctx.fillRect(0, 0, 256, 128);
            ctx.font = 'Bold 80px sans-serif';
            ctx.fillStyle = '#38bdf8';
            ctx.textAlign = 'center';
            ctx.textBaseline = 'middle';
            ctx.fillText(text, 128, 64);
            return new THREE.CanvasTexture(canvas);
        }

        // Racks vacíos
        HILERAS_LIST.forEach((hilera, hIdx) => {
            const z_coord = hIdx * 4.5;
            const sides = [
                { name: 'O', sign: -1 },
                { name: 'H', sign: 1 }
            ];

            sides.forEach(side => {
                for (let modulo = 1; modulo <= MODULOS; modulo++) {
                    const x_coord = side.sign * (modulo * 3.2 + 4.0);
                    for (let level = 1; level <= NIVELES; level++) {
                        const shelf = new THREE.Mesh(rackGeometry, rackMaterial);
                        shelf.position.set(x_coord, (level - 1) * 2.2 - 0.85, z_coord);
                        scene.add(shelf);
                    }
                }

                const labelTexture = createTextTexture(`${hilera}${side.name}`);
                const labelMaterial = new THREE.MeshBasicMaterial({ map: labelTexture, side: THREE.DoubleSide });
                const labelGeometry = new THREE.PlaneGeometry(2.2, 1.1);
                const labelMesh = new THREE.Mesh(labelGeometry, labelMaterial);
                const labelX = side.sign * 5.8;
                labelMesh.position.set(labelX, 0.2, z_coord);
                labelMesh.rotation.y = side.sign * (Math.PI / 2);
                scene.add(labelMesh);
            });
        });

        // Pallets
        const meshes = [];
        const legendMap = new Map();

        palletsData.forEach(p => {
            const materialParams = {
                color: new THREE.Color(p.color),
                roughness: 0.3,
                metalness: 0.1
            };

            if (p.duplicate) {
                materialParams.emissive = new THREE.Color(0xff3333);
                materialParams.emissiveIntensity = 0.3;
            } else if (p.alerta) {
                materialParams.emissive = new THREE.Color(0xff9900);
                materialParams.emissiveIntensity = 0.35;
            }

            const material = new THREE.MeshStandardMaterial(materialParams);
            const mesh = new THREE.Mesh(geometry, material);
            mesh.position.set(p.x, p.y, p.z);
            mesh.userData = p;

            if (p.alerta) {
                const edges = new THREE.LineSegments(
                    edgesGeometry,
                    new THREE.LineBasicMaterial({ color: 0xff9900 })
                );
                mesh.add(edges);
            }

            scene.add(mesh);
            meshes.push(mesh);
            legendMap.set(p.grupo, p.color);
        });

        // Leyenda dinámica
        const legendContainer = document.getElementById('legend-items');
        legendMap.forEach((color, grupo) => {
            const item = document.createElement('div');
            item.className = 'legend-item';
            item.innerHTML = `<div class="legend-color" style="background-color: ${color}"></div><span>${grupo}</span>`;
            legendContainer.appendChild(item);
        });

        // Tooltip
        const raycaster = new THREE.Raycaster();
        const mouse = new THREE.Vector2();
        const tooltip = document.getElementById('tooltip');

        function onMouseMove(event) {
            const rect = renderer.domElement.getBoundingClientRect();
            mouse.x = ((event.clientX - rect.left) / rect.width) * 2 - 1;
            mouse.y = -((event.clientY - rect.top) / rect.height) * 2 + 1;

            if (tooltip.style.display === 'block') {
                tooltip.style.left = (event.clientX + 15) + 'px';
                tooltip.style.top = (event.clientY + 15) + 'px';
            }
        }
        window.addEventListener('mousemove', onMouseMove, false);

        function animate() {
            requestAnimationFrame(animate);
            controls.update();

            const time = Date.now() * 0.005;
            const pulseScale = 1.0 + Math.sin(time) * 0.08;
            const pulseEmissive = 0.15 + (Math.sin(time) + 1) / 2 * 0.45;

            meshes.forEach(m => {
                if (m.userData.duplicate) {
                    m.scale.set(pulseScale, pulseScale, pulseScale);
                    if (m.material.emissive) {
                        m.material.emissiveIntensity = pulseEmissive;
                    }
                }
            });

            raycaster.setFromCamera(mouse, camera);
            const intersects = raycaster.intersectObjects(meshes, false);

            if (intersects.length > 0) {
                const t = intersects[0].object.userData;
                const diasTxt = (t.dias === null)
                    ? 'Sin fecha'
                    : (t.alerta
                        ? `<span style="color:#fb923c;font-weight:bold;">${t.dias} días ⚠️</span>`
                        : `${t.dias} días`);
                tooltip.style.display = 'block';
                tooltip.innerHTML = `
                    <div style="border-bottom: 1px solid #38bdf8; padding-bottom: 5px; margin-bottom: 8px;">
                        <strong style="color: #38bdf8; font-size: 14px;">📍 Ubicación: ${t.pos}</strong>
                    </div>
                    📦 <strong>N° de Pallet:</strong> ${t.pallet}<br/>
                    🚢 <strong>Exportadora:</strong> ${t.exportadora}<br/>
                    📐 <strong>Formato:</strong> ${t.formato}<br/>
                    🔢 <strong>Lote:</strong> ${t.lote}<br/>
                    🏷️ <strong>Variedad:</strong> ${t.variedad}<br/>
                    📊 <strong>Cantidad:</strong> ${t.cantidad} un.<br/>
                    📅 <strong>Fecha base:</strong> ${t.fecha} <span style="color:#94a3b8;">(${t.origen}${t.fecha !== t.fecha_excel ? '; Excel: ' + t.fecha_excel : ''})</span><br/>
                    ⏳ <strong>Permanencia:</strong> ${diasTxt}
                `;
                document.body.style.cursor = 'pointer';
            } else {
                tooltip.style.display = 'none';
                document.body.style.cursor = 'default';
            }

            renderer.render(scene, camera);
        }
        animate();

        window.addEventListener('resize', () => {
            camera.aspect = window.innerWidth / 600;
            camera.updateProjectionMatrix();
            renderer.setSize(window.innerWidth, 600);
        });
    </script>
</body>
</html>
"""

with tab_3d:
    df_cod = df_f[df_f['Codificada']].copy()
    # Evita dibujar dos veces el mismo pallet en la misma posición
    df_cod = pd.concat([
        df_cod[df_cod['Pallet'] != ''].drop_duplicates(['Pallet', 'Posicion']),
        df_cod[df_cod['Pallet'] == ''],
    ])

    if df_cod.empty:
        st.info("No hay pallets con posición codificada (C1HA11, C2OA24...) para los filtros seleccionados.")
    else:
        f1, f2, f3, f4 = st.columns(4)
        camaras_cod = ordenar_camaras(df_cod['Camara'].unique())
        camara_sel = f1.selectbox("Cámara", camaras_cod)
        df_camara = df_cod[df_cod['Camara'] == camara_sel]

        variedad_sel = f2.selectbox("Variedad", ["Todas"] + sorted(df_camara['Variedad'].unique()))
        rack_sel = f3.selectbox("Rack (hilera)", ["Todos"] + sorted(df_camara['Hilera'].unique()))
        color_por = f4.selectbox("Colorear por", ["Variedad", "Exportadora", "Formato"])
        solo_alertas = st.checkbox(f"Mostrar solo pallets con alerta (> {dias_alerta} días)")

        # KPIs de la cámara (antes de filtrar por variedad/rack)
        capacidad = CAPACIDAD_POR_CAMARA.get(camara_sel, POSICIONES_POR_CAMARA_DEFECTO)
        pallets_unicos = df_camara.loc[df_camara['Pallet'] != '', 'Pallet'].nunique()
        pos_ocupadas = df_camara['Posicion'].nunique()
        espacios_vacios = max(capacidad - pos_ocupadas, 0)
        cnt_pos = df_camara.groupby('Posicion')['Pallet'].nunique()
        sobre_apilados = int((cnt_pos[cnt_pos > 1] - 1).sum())
        pos_multiples = int((cnt_pos > 1).sum())
        alertas_camara = int(df_camara.drop_duplicates('Pallet')['Alerta'].sum())

        st.write("---")
        c1, c2, c3, c4, c5 = st.columns(5)
        c1.metric("Capacidad de la Cámara", f"{capacidad} slots",
                  help=f"{len(HILERAS)} racks x {len(SECTORES_NOMBRES)} lados x {PROFUNDIDAD_MAX} módulos x {ALTURA_MAX} niveles.")
        c2.metric("Pallets Únicos", f"{pallets_unicos} un.",
                  f"{(pos_ocupadas / capacidad * 100):.1f}% Ocupación")
        c3.metric("Espacios Vacíos", f"{espacios_vacios} un.",
                  f"{(espacios_vacios / capacidad * 100):.1f}% Disponible", delta_color="inverse")
        c4.metric("Misma Posición (Sobre-apilado)", f"{sobre_apilados} un.",
                  help=f"{pos_multiples} ubicaciones tienen más de un pallet distinto asignado.")
        c5.metric("Con alerta de permanencia", f"{alertas_camara} un.",
                  delta=f"> {dias_alerta} días" if alertas_camara else None, delta_color="inverse")

        # Filtros aplicados al render
        df_3d = df_camara
        if variedad_sel != "Todas":
            df_3d = df_3d[df_3d['Variedad'] == variedad_sel]
        if rack_sel != "Todos":
            df_3d = df_3d[df_3d['Hilera'] == rack_sel]
        if solo_alertas:
            df_3d = df_3d[df_3d['Alerta']]

        hilera_map = {h: i for i, h in enumerate(HILERAS)}
        datos_3d = []
        for r in df_3d.to_dict('records'):
            z = hilera_map.get(r['Hilera'], 0) * 4.5
            x_off = int(r['Modulo']) * 3.2 + 4.0
            x = -x_off if r['Sector'] == 'O' else x_off
            y = (int(r['Nivel']) - 1) * 2.2
            grupo = str(r[color_por])
            dias = None if pd.isna(r['Dias']) else int(r['Dias'])
            datos_3d.append({
                'pos': r['Posicion'],
                'pallet': r['Pallet'] or 'S/N',
                'lote': r['Lote'],
                'exportadora': r['Exportadora'],
                'formato': r['Formato'],
                'variedad': r['Variedad'],
                'cantidad': int(r['Cantidad']),
                'fecha': r['Fecha'].strftime('%d-%m-%Y') if pd.notna(r['Fecha']) else 'S/F',
                'origen': r['Origen_Fecha'],
                'fecha_excel': r['Fecha_Excel'].strftime('%d-%m-%Y') if pd.notna(r['Fecha_Excel']) else 'S/F',
                'dias': dias,
                'alerta': bool(r['Alerta']),
                'x': x, 'y': y, 'z': z,
                'grupo': grupo,
                'color': color_estable(grupo),
                'duplicate': bool(r['Sobre_apilado']),
            })

        html_final = (HTML_TEMPLATE
                      .replace("__THREEJS_DATA__", json.dumps(datos_3d))
                      .replace("__LEYENDA_TITULO__", f"{color_por}")
                      .replace("__DIAS_ALERTA__", str(int(dias_alerta)))
                      .replace("__HILERAS_JSON__", json.dumps(HILERAS))
                      .replace("__PROFUNDIDAD__", str(PROFUNDIDAD_MAX))
                      .replace("__ALTURA__", str(ALTURA_MAX)))
        components.html(html_final, height=620)

        st.markdown("### 📊 Inventario Físico")
        tabla = df_3d.copy()
        tabla['Sector'] = tabla['Sector'].map(SECTORES_NOMBRES)
        tabla['Fecha Creacion'] = tabla['Fecha'].dt.strftime('%d-%m-%Y')   # fecha base usada
        mostrar_tabla(tabla[['Posicion', 'Pallet', 'Exportadora', 'Formato', 'Lote', 'Variedad',
                             'Cantidad', 'KN', 'Sector', 'Fecha Creacion', 'Origen_Fecha', 'Dias', 'Alerta']])

# -----------------------------------------------------------------------------
# TAB 3: ALERTAS
# -----------------------------------------------------------------------------
with tab_alertas:
    st.markdown(f"### ⚠️ Pallets con permanencia mayor a {dias_alerta} días")
    st.caption("La antigüedad se cuenta desde la fecha de proceso del lote (recepción si no se procesa), que cuenta como día 1. "
               "Cada pallet se cuenta una sola vez.")

    if 'Fecha' not in mapa_columnas:
        st.warning("El Excel no tiene columna 'Fecha Creacion', no se pueden calcular alertas.")
    else:
        alertas = df_u[df_u['Alerta']].sort_values('Dias', ascending=False)
        if alertas.empty:
            st.success(f"✅ Ningún pallet supera los {dias_alerta} días de permanencia.")
        else:
            st.error(f"{len(alertas)} pallet(s) en alerta")

            camaras_al = ordenar_camaras(alertas['Camara'].unique())
            cam_filtro = st.multiselect("Filtrar por cámara / ubicación", camaras_al, default=camaras_al)
            alertas = alertas[alertas['Camara'].isin(cam_filtro)]

            a1, a2 = st.columns(2)
            with a1:
                st.markdown("**Por cámara / ubicación**")
                res_cam = alertas.groupby('Camara').size().rename('Pallets en alerta')
                res_cam = res_cam.reindex(ordenar_camaras(res_cam.index))
                mostrar_tabla(res_cam.to_frame())
            with a2:
                st.markdown("**Por exportadora**")
                mostrar_tabla(alertas.groupby('Exportadora').size().rename('Pallets en alerta').to_frame())

            por_ajuste = int((~(alertas['Dias_Excel'] > dias_alerta)).sum())
            if por_ajuste:
                st.warning(f"📅 {por_ajuste} de estos pallets NO estarían en alerta según la Fecha Creacion del mapeo; "
                           f"entran porque la antigüedad se calcula desde la fecha de proceso / recepción del lote "
                           f"(o un ajuste manual).")

            salida = alertas[['Pallet', 'Posicion', 'Camara', 'Exportadora', 'Formato', 'Variedad',
                              'Lote', 'Fecha_Excel', 'Fecha', 'Dias_Excel', 'Dias',
                              'Origen_Fecha', 'Motivo_Ajuste']].copy()
            salida['Fecha_Excel'] = salida['Fecha_Excel'].dt.strftime('%d-%m-%Y')
            salida['Fecha'] = salida['Fecha'].dt.strftime('%d-%m-%Y')
            salida['Dias'] = salida['Dias'].astype(int)
            salida['Dias_Excel'] = salida['Dias_Excel'].astype(int)
            salida['Días sobre el límite'] = salida['Dias'] - int(dias_alerta)
            salida = salida.rename(columns={
                'Fecha_Excel': 'Fecha Creacion (mapeo)',
                'Fecha': 'Fecha usada',
                'Dias_Excel': 'Días según mapeo',
                'Dias': 'Días de permanencia',
                'Origen_Fecha': 'Origen de la fecha',
                'Motivo_Ajuste': 'Motivo',
            })
            mostrar_tabla(salida, hide_index=True)
            st.download_button("⬇️ Descargar alertas (CSV)", salida.to_csv(index=False).encode('utf-8-sig'),
                               file_name="alertas_permanencia.csv", mime="text/csv")

# -----------------------------------------------------------------------------
# TAB 4: RECEPCIÓN → PROCESO
# -----------------------------------------------------------------------------
with tab_rec:
    st.markdown("### ⏱️ Días desde la recepción hasta el proceso (por lote)")
    if lotes_rec.empty:
        st.warning(aviso_rec or "No hay datos de recepción disponibles.")
    else:
        st.caption("Días = fecha de proceso − fecha de recepción (0 = mismo día). Si un lote se procesó en más de un día "
                   "se usa la fecha más antigua. Se excluyen los lotes rechazados y el bulk (no se procesa).")
        base_rec = lotes_rec[lotes_rec['Exportadora'].isin(exportadoras_sel)].copy()

        r1, r2 = st.columns([2, 1])
        solo_mapeo = r2.checkbox("Solo lotes que están en el mapeo", value=False)
        if solo_mapeo:
            base_rec = base_rec[base_rec['Lote'].isin(set(df_all['Lote']))]
        if base_rec['Recepcion'].notna().any():
            rango = r1.date_input("Rango de fecha de recepción",
                                  value=(base_rec['Recepcion'].min().date(), base_rec['Recepcion'].max().date()))
            if isinstance(rango, (tuple, list)) and len(rango) == 2:
                ini, fin = pd.Timestamp(rango[0]), pd.Timestamp(rango[1])
                base_rec = base_rec[(base_rec['Recepcion'] >= ini) & (base_rec['Recepcion'] <= fin)]

        rechazados = base_rec[base_rec['Rechazado_Total']]
        bulk = base_rec[~base_rec['Rechazado_Total'] & base_rec['Bulk']]
        normales = base_rec[~base_rec['Rechazado_Total'] & ~base_rec['Bulk']]
        procesados = normales[normales['Proceso'].notna()].copy()
        pendientes = normales[normales['Proceso'].isna()].copy()
        procesados['Dias'] = (procesados['Proceso'] - procesados['Recepcion']).dt.days

        m1, m2, m3, m4, m5 = st.columns(5)
        m1.metric("Lotes procesados", len(procesados))
        m2.metric("Promedio recepción → proceso",
                  f"{procesados['Dias'].mean():.1f} días" if len(procesados) else "—")
        m3.metric("Máximo", f"{int(procesados['Dias'].max())} días" if len(procesados) else "—")
        m4.metric("Pendientes de proceso", len(pendientes))
        m5.metric("Bulk / Rechazados", f"{len(bulk)} / {len(rechazados)}",
                  help="Bulk no se procesa; los lotes rechazados se despachan. Ninguno entra en la estadística.")

        if procesados.empty:
            st.info("No hay lotes procesados para los filtros seleccionados.")
        else:
            if (procesados['Dias'] < 0).any():
                st.error(f"⚠️ {int((procesados['Dias'] < 0).sum())} lote(s) tienen fecha de proceso ANTERIOR a la "
                         f"recepción (posible error de digitación). Revísalos en la tabla de detalle.")

            dist = procesados.groupby('Dias').agg(Lotes=('Lote', 'size'),
                                                  Pallets_recepcion=('Pallets_OK', 'sum')).reset_index()
            dist['% de lotes'] = (dist['Lotes'] / dist['Lotes'].sum() * 100).round(1)
            dist.insert(0, 'Recepción → proceso',
                        dist['Dias'].map(lambda d: 'Mismo día' if d == 0 else f"{int(d)} día(s)"))

            g1, g2 = st.columns([3, 2])
            with g1:
                st.markdown("**Lotes según días de recepción a proceso**")
                st.bar_chart(dist.set_index('Recepción → proceso')['Lotes'])
            with g2:
                st.markdown("**Cantidad de lotes por días**")
                mostrar_tabla(dist.drop(columns='Dias').rename(columns={'Pallets_recepcion': 'Pallets de recepción'}),
                              hide_index=True)

            st.markdown("**Por exportadora**")
            cruz = procesados.pivot_table(index='Dias', columns='Exportadora', values='Lote',
                                          aggfunc='count', fill_value=0)
            cruz.index = [('Mismo día' if d == 0 else f"{int(d)} día(s)") for d in cruz.index]
            cruz.index.name = 'Recepción → proceso'
            cruz.columns.name = None
            cruz['TOTAL'] = cruz.sum(axis=1)
            mostrar_tabla(cruz)

            st.markdown("**Detalle de lotes**")
            minimo = st.number_input("Mostrar lotes con este número de días o más", min_value=0, max_value=30,
                                     value=2, step=1)
            det = procesados[procesados['Dias'] >= minimo].sort_values(['Dias', 'Lote'], ascending=[False, True])
            det = det[['Lote', 'Exportadora', 'Variedad', 'Recepcion', 'Proceso', 'Dias',
                       'Procesos_Distintos', 'Pallets_OK']].copy()
            det['Recepcion'] = det['Recepcion'].dt.strftime('%d-%m-%Y')
            det['Proceso'] = det['Proceso'].dt.strftime('%d-%m-%Y')
            det = det.rename(columns={'Recepcion': 'Fecha recepción', 'Proceso': 'Fecha proceso (más antigua)',
                                      'Dias': 'Días', 'Procesos_Distintos': 'Fechas de proceso distintas',
                                      'Pallets_OK': 'Pallets'})
            st.caption(f"{len(det)} lote(s) con {int(minimo)} día(s) o más.")
            mostrar_tabla(det, hide_index=True)
            st.download_button("⬇️ Descargar detalle (CSV)", det.to_csv(index=False).encode('utf-8-sig'),
                               file_name="lotes_recepcion_a_proceso.csv", mime="text/csv")

        if not pendientes.empty:
            st.markdown("**Lotes recibidos aún sin fecha de proceso**")
            pend = pendientes[['Lote', 'Exportadora', 'Variedad', 'Recepcion', 'Pallets_OK']].copy()
            pend['Días desde recepción'] = (hoy - pend['Recepcion']).dt.days
            pend['Recepcion'] = pend['Recepcion'].dt.strftime('%d-%m-%Y')
            pend = pend.rename(columns={'Recepcion': 'Fecha recepción', 'Pallets_OK': 'Pallets'})
            mostrar_tabla(pend.sort_values('Días desde recepción', ascending=False), hide_index=True)

# -----------------------------------------------------------------------------
# TAB 5: AJUSTES DE FECHA
# -----------------------------------------------------------------------------
with tab_ajustes:
    st.markdown("### 📝 Ajustes de fecha de antigüedad")
    st.caption("Solo para EXCEPCIONES que la fecha de proceso/recepción del lote no resuelva. La antigüedad normal ya se "
               "calcula desde la base de recepción. Un ajuste manual tiene prioridad sobre todo; el pallet debe "
               "escribirse igual que en el mapeo.")

    if 'msg_ajustes' in st.session_state:
        st.success(st.session_state.pop('msg_ajustes'))

    editado = st.data_editor(
        ajustes, num_rows="dynamic", hide_index=True, key="editor_ajustes",
        column_config={
            'Pallet': st.column_config.TextColumn("N° Pallet", required=True),
            'Fecha Original': st.column_config.DateColumn("Fecha original de proceso",
                                                          format="DD/MM/YYYY", required=True),
            'Motivo': st.column_config.TextColumn("Motivo (opcional)"),
        },
    )
    if st.button("💾 Guardar ajustes", type="primary"):
        n_guardados = guardar_ajustes(editado)
        st.session_state['msg_ajustes'] = f"Se guardaron {n_guardados} ajuste(s) en '{ARCHIVO_AJUSTES}'."
        st.cache_data.clear()
        (getattr(st, 'rerun', None) or st.experimental_rerun)()

    if not ajustes.empty:
        st.markdown("**Estado de los ajustes guardados**")
        en_excel = set(df_all['Pallet'])
        estado = ajustes.copy()
        fechas_excel = df_all.drop_duplicates('Pallet').set_index('Pallet')['Fecha_Excel']
        estado['Fecha Creacion (Excel)'] = estado['Pallet'].map(fechas_excel)
        estado['Días ganados'] = (estado['Fecha Creacion (Excel)'] - estado['Fecha Original']).dt.days
        estado['Estado'] = estado['Pallet'].map(lambda p: 'En el Excel' if p in en_excel
                                                else '⚠️ No está en el Excel (¿ya salió? bórralo)')
        estado['Fecha Original'] = estado['Fecha Original'].dt.strftime('%d-%m-%Y')
        estado['Fecha Creacion (Excel)'] = estado['Fecha Creacion (Excel)'].dt.strftime('%d-%m-%Y')
        mostrar_tabla(estado, hide_index=True)
        st.caption("Si un número de pallet sale de la planta y se reutiliza después, borra su ajuste para que no "
                   "se aplique a un pallet distinto.")

# -----------------------------------------------------------------------------
# TAB 6: CALIDAD DE DATOS
# -----------------------------------------------------------------------------
with tab_calidad:
    st.markdown("### 🔎 Revisión del archivo Excel")
    st.caption("Esta pestaña analiza todo el Excel, sin aplicar el filtro de exportadora.")

    filas_total = len(df_all)
    sin_pallet = int((df_all['Pallet'] == '').sum())
    pallets_unicos_total = df_all.loc[df_all['Pallet'] != '', 'Pallet'].nunique()
    sin_fecha = int(df_all['Fecha'].isna().sum())
    fechas_futuras = int((df_all['Dias'] < 1).sum())

    q1, q2, q3, q4 = st.columns(4)
    q1.metric("Filas en el Excel", filas_total)
    q2.metric("Pallets únicos", pallets_unicos_total)
    q3.metric("Filas repetidas", filas_total - sin_pallet - pallets_unicos_total,
              help="Filas cuyo número de pallet ya apareció antes.")
    q4.metric("Sin N° de pallet", sin_pallet, help="Estas filas no se cuentan en el dashboard.")
    st.write(f"Filas sin fecha válida: **{sin_fecha}** · Filas con fecha futura: **{fechas_futuras}**")

    con_pallet = df_all[df_all['Pallet'] != '']
    g = con_pallet.groupby('Pallet').agg(
        Filas=('Pallet', 'size'),
        Posiciones_distintas=('Posicion', 'nunique'),
        Posiciones=('Posicion', lambda s: ', '.join(sorted(set(s)))),
        Exportadoras=('Exportadora', lambda s: ', '.join(sorted(set(s)))),
    )

    st.markdown("**Pallets repetidos en el Excel**")
    rep = g[g['Filas'] > 1].sort_values('Filas', ascending=False)
    if rep.empty:
        st.success("No hay pallets repetidos.")
    else:
        mostrar_tabla(rep)

    st.markdown("**Pallets registrados en más de una posición distinta**")
    multi = g[g['Posiciones_distintas'] > 1]
    if multi.empty:
        st.success("Ningún pallet aparece en posiciones distintas.")
    else:
        st.caption("En el dashboard se cuentan una sola vez (se prefiere la posición codificada y la fecha más antigua).")
        mostrar_tabla(multi)

    st.markdown("**Ubicaciones no codificadas encontradas en la columna Posicion**")
    nc = (df_all[~df_all['Codificada']].groupby('Posicion').size()
          .rename('Filas').sort_values(ascending=False))
    mostrar_tabla(nc.to_frame())


    # --- Controles que dependen de la base de recepción ---
    st.markdown("---")
    st.markdown("### 🔗 Cruce con la base de recepción (por Lote)")

    def resumen_lotes(d):
        t = d.groupby('Lote').agg(Exportadora=('Exportadora', 'first'), Pallets=('Pallet', 'nunique'),
                                  Fecha_Creacion_mapeo=('Fecha_Excel', 'min'), Fecha_Base=('Fecha_Base', 'min'))
        t['Diferencia (días)'] = (t['Fecha_Creacion_mapeo'] - t['Fecha_Base']).dt.days
        for c in ('Fecha_Creacion_mapeo', 'Fecha_Base'):
            t[c] = t[c].dt.strftime('%d-%m-%Y')
        return t.rename(columns={'Fecha_Creacion_mapeo': 'Fecha Creacion (mapeo)',
                                 'Fecha_Base': 'Fecha proceso / recepción'}).reset_index()

    if lotes_rec.empty:
        st.warning(aviso_rec or "No hay base de recepción.")
    else:
        sin_lote = df_all[df_all['Fecha_Recepcion'].isna()]
        st.markdown("**Lotes del mapeo que no se encontraron en la base de recepción**")
        if sin_lote.empty:
            st.success("Todos los lotes del mapeo existen en la base de recepción.")
        else:
            mostrar_tabla(sin_lote.groupby('Lote').agg(Exportadora=('Exportadora', 'first'),
                                                       Pallets=('Pallet', 'nunique')).reset_index(),
                          hide_index=True)

        st.markdown("**🚨 Fecha Creacion ANTERIOR al proceso/recepción del lote (posible error de digitación)**")
        err = df_all[df_all['Excel_Antes_Base']]
        if err.empty:
            st.success("No hay lotes con fecha de creación anterior al proceso.")
        else:
            st.caption("Un pallet no puede crearse antes de que se procese su lote. Para no subestimar la antigüedad "
                       "se usa la fecha más antigua.")
            mostrar_tabla(resumen_lotes(err), hide_index=True)

        st.markdown("**Lotes con Fecha Creacion POSTERIOR al proceso (re-fechados)**")
        refe = df_all[df_all['Fecha_Excel'].notna() & df_all['Fecha_Base'].notna()
                      & (df_all['Fecha_Excel'] > df_all['Fecha_Base'])
                      & (df_all['Fecha_Proceso'].notna())]
        if refe.empty:
            st.success("No hay lotes re-fechados.")
        else:
            st.caption("La antigüedad se calcula desde la fecha de proceso, por lo que estos cambios no restan días.")
            mostrar_tabla(resumen_lotes(refe), hide_index=True)

        st.markdown("**Lotes procesados en más de una fecha**")
        dos = lotes_rec[lotes_rec['Procesos_Distintos'] > 1]
        if dos.empty:
            st.success("Ningún lote tiene más de una fecha de proceso.")
        else:
            st.caption("Se usa la fecha más antigua.")
            mostrar_tabla(dos[['Lote', 'Exportadora', 'Recepcion', 'Proceso', 'Procesos_Distintos']]
                          .assign(Recepcion=lambda d: d['Recepcion'].dt.strftime('%d-%m-%Y'),
                                  Proceso=lambda d: d['Proceso'].dt.strftime('%d-%m-%Y'))
                          .rename(columns={'Proceso': 'Proceso (más antigua)',
                                           'Procesos_Distintos': 'Fechas distintas'}),
                          hide_index=True)

        st.markdown("**🚫 Pallets del mapeo en lotes rechazados (excluidos de todo el análisis)**")
        if df_rech.empty:
            st.success("No hay pallets del mapeo en lotes rechazados.")
        else:
            st.caption("Los lotes rechazados se despachan y no deberían estar en el mapeo.")
            mostrar_tabla(df_rech.groupby('Lote').agg(Exportadora=('Exportadora', 'first'),
                                                      Pallets=('Pallet', 'nunique'),
                                                      Posiciones=('Posicion', lambda s: ', '.join(sorted(set(s))))
                                                      ).reset_index(), hide_index=True)

    st.markdown("---")
    st.markdown("### 🧹 Pallets excluidos por Tipo (IQF / fresco sin proceso)")
    if not excluir_tipos:
        st.info("La exclusión está desactivada en el panel lateral: estos pallets se están incluyendo en el análisis.")
    elif df_excl.empty:
        st.success("No se encontraron pallets de tipo IQF o fresco en el mapeo.")
    else:
        st.caption("Estos pallets no pertenecen al mapeo de fruta embalada y no se cuentan en el dashboard, "
                   "las alertas ni el mapa 3D. Conviene sacarlos del Excel.")
        mostrar_tabla(df_excl.groupby(['Tipo', 'Lote']).agg(
            Exportadora=('Exportadora', 'first'), Formato=('Formato', 'first'),
            Pallets=('Pallet', 'nunique'),
            Posiciones=('Posicion', lambda x: ', '.join(sorted(set(x))))).reset_index(), hide_index=True)
