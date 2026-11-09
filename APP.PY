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
ARCHIVO_EXCEL = "MAPEO ARANDANOS.xlsx"
DIAS_ALERTA_DEFECTO = 7          # Alerta si la permanencia es MAYOR a este número de días
                                 # (el día de "Fecha Creacion" cuenta como día 1)

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
    return str(v).strip()


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


# =============================================================================
# CARGA DE DATOS
# =============================================================================
def ubicar_excel():
    for ruta in (Path(__file__).resolve().parent / ARCHIVO_EXCEL, Path.cwd() / ARCHIVO_EXCEL):
        if ruta.exists():
            return ruta
    return None


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
    df['Lote'] = texto('Lote', '—')
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
df_all['Dias'] = (hoy - df_all['Fecha']).dt.days + 1          # Fecha de creación = día 1
df_all['Alerta'] = df_all['Dias'] > dias_alerta               # NaN > n -> False

# Filas filtradas por exportadora (todas las filas del Excel)
df_f = df_all[df_all['Exportadora'].isin(exportadoras_sel)].copy()

# Un solo registro por número de pallet (evita contar dos veces el mismo pallet)
df_u = deduplicar_pallets(df_all)
df_u = df_u[df_u['Exportadora'].isin(exportadoras_sel)].copy()

# =============================================================================
# PESTAÑAS
# =============================================================================
tab_dash, tab_3d, tab_alertas, tab_calidad = st.tabs(
    ["📊 Dashboard Resumen", "🧊 Mapa 3D", "⚠️ Alertas de Permanencia", "🔎 Calidad de Datos"]
)

# -----------------------------------------------------------------------------
# TAB 1: DASHBOARD
# -----------------------------------------------------------------------------
with tab_dash:
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
                    📅 <strong>Creación:</strong> ${t.fecha}<br/>
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
        tabla['Fecha Creacion'] = tabla['Fecha'].dt.strftime('%d-%m-%Y')
        mostrar_tabla(tabla[['Posicion', 'Pallet', 'Exportadora', 'Formato', 'Lote', 'Variedad',
                             'Cantidad', 'KN', 'Sector', 'Fecha Creacion', 'Dias', 'Alerta']])

# -----------------------------------------------------------------------------
# TAB 3: ALERTAS
# -----------------------------------------------------------------------------
with tab_alertas:
    st.markdown(f"### ⚠️ Pallets con permanencia mayor a {dias_alerta} días")
    st.caption("El día de «Fecha Creacion» cuenta como día 1. Cada pallet se cuenta una sola vez.")

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

            salida = alertas[['Pallet', 'Posicion', 'Camara', 'Exportadora', 'Formato',
                              'Variedad', 'Lote', 'Fecha', 'Dias']].copy()
            salida['Fecha'] = salida['Fecha'].dt.strftime('%d-%m-%Y')
            salida['Dias'] = salida['Dias'].astype(int)
            salida['Días sobre el límite'] = salida['Dias'] - int(dias_alerta)
            salida = salida.rename(columns={'Fecha': 'Fecha Creacion', 'Dias': 'Días de permanencia'})
            mostrar_tabla(salida, hide_index=True)
            st.download_button("⬇️ Descargar alertas (CSV)", salida.to_csv(index=False).encode('utf-8-sig'),
                               file_name="alertas_permanencia.csv", mime="text/csv")

# -----------------------------------------------------------------------------
# TAB 4: CALIDAD DE DATOS
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