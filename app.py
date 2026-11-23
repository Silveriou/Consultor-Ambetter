import streamlit as st
import pandas as pd
import io
from datetime import datetime

# ─────────────────────────────────────────────
# CONFIGURACIÓN DE PÁGINA
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="Consultor Ambetter",
    page_icon="🏥",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─────────────────────────────────────────────
# ESTILOS CSS
# ─────────────────────────────────────────────
st.markdown("""
<style>
    /* Fondo general */
    .stApp { background-color: #f0f4f8; }

    /* Tarjetas de métricas */
    [data-testid="metric-container"] {
        background-color: #ffffff;
        border: 1px solid #e0e7ef;
        border-radius: 12px;
        padding: 16px 20px;
        box-shadow: 0 2px 8px rgba(0,0,0,0.06);
    }

    /* Encabezado principal */
    .titulo-principal {
        font-size: 2rem;
        font-weight: 700;
        color: #1a3c5e;
        margin-bottom: 0px;
    }
    .subtitulo {
        font-size: 1rem;
        color: #5a7a99;
        margin-bottom: 1.5rem;
    }

    /* Sidebar */
    [data-testid="stSidebar"] {
        background-color: #1a3c5e;
    }
    [data-testid="stSidebar"] * {
        color: #e8f0fa !important;
    }
    [data-testid="stSidebar"] .stFileUploader label {
        color: #e8f0fa !important;
    }

    /* Separador */
    hr { border-color: #d0dce8; }

    /* Badges de error */
    .badge-error {
        background-color: #ffe4e4;
        color: #c0392b;
        border-radius: 6px;
        padding: 2px 10px;
        font-weight: 600;
    }
    .badge-ok {
        background-color: #e4f8e4;
        color: #27ae60;
        border-radius: 6px;
        padding: 2px 10px;
        font-weight: 600;
    }
</style>
""", unsafe_allow_html=True)


# ─────────────────────────────────────────────
# FUNCIONES AUXILIARES
# ─────────────────────────────────────────────
@st.cache_data
def cargar_archivo(archivo):
    """Carga un archivo Excel o CSV y retorna un DataFrame."""
    nombre = archivo.name.lower()
    if nombre.endswith(".csv"):
        df = pd.read_csv(archivo, encoding="utf-8", on_bad_lines="skip")
    else:
        df = pd.read_excel(archivo, engine="openpyxl")
    return df


def detectar_errores(df):
    """Detecta filas duplicadas y celdas vacías en columnas clave."""
    duplicados = df[df.duplicated(keep=False)]
    vacios = df[df.isnull().any(axis=1)]
    return duplicados, vacios


def convertir_a_excel(df):
    """Convierte un DataFrame a bytes de Excel para descarga."""
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="Datos")
        # Formato de tabla
        workbook  = writer.book
        worksheet = writer.sheets["Datos"]
        fmt_header = workbook.add_format({
            "bold": True, "bg_color": "#1a3c5e",
            "font_color": "white", "border": 1
        })
        for col_num, col_name in enumerate(df.columns):
            worksheet.write(0, col_num, col_name, fmt_header)
            worksheet.set_column(col_num, col_num, max(15, len(str(col_name)) + 4))
    return buffer.getvalue()


# ─────────────────────────────────────────────
# BARRA LATERAL
# ─────────────────────────────────────────────
with st.sidebar:
    st.image("https://img.icons8.com/fluency/96/health-graph.png", width=70)
    st.markdown("## 🏥 Consultor Ambetter")
    st.markdown("---")
    st.markdown("### 📂 Cargar archivo")
    archivo = st.file_uploader(
        "Sube tu archivo de clientes",
        type=["xlsx", "xls", "csv"],
        help="Los datos NO se suben a ningún servidor. Todo se procesa localmente.",
    )
    st.markdown("---")
    st.markdown("### ℹ️ Información")
    st.info("🔒 **Privacidad garantizada**\nTus archivos se procesan solo en tu computadora. Nunca se envían a internet.", icon="🔐")
    st.caption(f"© {datetime.now().year} Consultor Ambetter")


# ─────────────────────────────────────────────
# PANTALLA DE BIENVENIDA
# ─────────────────────────────────────────────
if archivo is None:
    st.markdown('<p class="titulo-principal">🏥 Consultor Ambetter</p>', unsafe_allow_html=True)
    st.markdown('<p class="subtitulo">Plataforma segura de análisis de información de clientes</p>', unsafe_allow_html=True)
    st.markdown("---")

    col1, col2, col3 = st.columns(3)
    with col1:
        st.markdown("### 📊 Analiza")
        st.markdown("Visualiza métricas clave de tus clientes al instante.")
    with col2:
        st.markdown("### 🔍 Filtra")
        st.markdown("Busca y filtra por cualquier columna de tu archivo.")
    with col3:
        st.markdown("### 💾 Descarga")
        st.markdown("Exporta los datos filtrados o corregidos en Excel.")

    st.markdown("---")
    st.info("👈 **Para comenzar**, sube un archivo Excel o CSV desde la barra lateral izquierda.", icon="📂")
    st.stop()


# ─────────────────────────────────────────────
# CARGA DE DATOS
# ─────────────────────────────────────────────
try:
    df = cargar_archivo(archivo)
except Exception as e:
    st.error(f"❌ Error al leer el archivo: {e}")
    st.stop()

# ─────────────────────────────────────────────
# ENCABEZADO
# ─────────────────────────────────────────────
st.markdown('<p class="titulo-principal">🏥 Consultor Ambetter</p>', unsafe_allow_html=True)
st.markdown(f'<p class="subtitulo">Archivo cargado: <strong>{archivo.name}</strong></p>', unsafe_allow_html=True)
st.markdown("---")

# ─────────────────────────────────────────────
# MÉTRICAS RÁPIDAS
# ─────────────────────────────────────────────
duplicados_df, vacios_df = detectar_errores(df)
total_clientes    = len(df)
total_columnas    = len(df.columns)
total_duplicados  = len(df[df.duplicated()])
total_vacios      = df.isnull().sum().sum()

col1, col2, col3, col4 = st.columns(4)
col1.metric("👥 Total de Clientes",    f"{total_clientes:,}")
col2.metric("📋 Columnas Detectadas",  f"{total_columnas}")
col3.metric("🔁 Registros Duplicados", f"{total_duplicados:,}",
            delta=f"-{total_duplicados}" if total_duplicados > 0 else "✓ Sin duplicados",
            delta_color="inverse")
col4.metric("⚠️ Celdas Vacías",        f"{total_vacios:,}",
            delta=f"-{total_vacios}" if total_vacios > 0 else "✓ Sin vacíos",
            delta_color="inverse")

st.markdown("---")

# ─────────────────────────────────────────────
# PESTAÑAS PRINCIPALES
# ─────────────────────────────────────────────
tab1, tab2, tab3 = st.tabs(["📋 Datos y Filtros", "🚨 Detección de Errores", "💾 Descargar"])


# ══════════════════════════════════════════════
# TAB 1: DATOS Y FILTROS
# ══════════════════════════════════════════════
with tab1:
    st.markdown("### 🔍 Buscar y Filtrar")

    # Buscador global por texto
    busqueda = st.text_input(
        "🔎 Búsqueda general",
        placeholder="Escribe cualquier texto para buscar en todos los campos...",
    )

    # Filtros dinámicos por columna
    with st.expander("⚙️ Filtros avanzados por columna", expanded=False):
        filtros = {}
        cols_filtro = st.columns(min(3, total_columnas))
        for i, col in enumerate(df.columns[:9]):  # máx 9 columnas para no saturar
            with cols_filtro[i % 3]:
                if df[col].dtype == object or df[col].nunique() <= 30:
                    opciones = ["(Todos)"] + sorted(df[col].dropna().astype(str).unique().tolist())
                    sel = st.selectbox(f"**{col}**", opciones, key=f"filtro_{col}")
                    if sel != "(Todos)":
                        filtros[col] = sel
                else:
                    col_data = df[col].dropna()
                    if len(col_data) == 0:
                        st.caption(f"**{col}**: sin datos")
                    else:
                        min_val = float(col_data.min())
                        max_val = float(col_data.max())
                        if min_val == max_val:
                            st.caption(f"**{col}**: {min_val} (valor único)")
                        else:
                            rango = st.slider(
                                f"**{col}**",
                                min_value=min_val, max_value=max_val,
                                value=(min_val, max_val),
                                key=f"filtro_{col}"
                            )
                            if rango != (min_val, max_val):
                                filtros[col] = rango

    # Aplicar búsqueda global
    df_filtrado = df.copy()
    if busqueda:
        mask = df_filtrado.apply(
            lambda row: row.astype(str).str.contains(busqueda, case=False, na=False).any(),
            axis=1
        )
        df_filtrado = df_filtrado[mask]

    # Aplicar filtros por columna
    for col, valor in filtros.items():
        if isinstance(valor, tuple):
            df_filtrado = df_filtrado[
                df_filtrado[col].between(valor[0], valor[1])
            ]
        else:
            df_filtrado = df_filtrado[
                df_filtrado[col].astype(str) == valor
            ]

    # Resultado
    st.markdown(f"**Mostrando {len(df_filtrado):,} de {total_clientes:,} registros**")
    st.dataframe(df_filtrado, use_container_width=True, height=450)


# ══════════════════════════════════════════════
# TAB 2: DETECCIÓN DE ERRORES
# ══════════════════════════════════════════════
with tab2:
    st.markdown("### 🚨 Reporte de Calidad de Datos")

    col_a, col_b = st.columns(2)

    with col_a:
        st.markdown("#### 🔁 Registros Duplicados")
        if len(duplicados_df) > 0:
            st.warning(f"Se encontraron **{len(duplicados_df)}** filas duplicadas.")
            st.dataframe(duplicados_df, use_container_width=True, height=300)
        else:
            st.success("✅ No se encontraron registros duplicados.")

    with col_b:
        st.markdown("#### ⚠️ Filas con Datos Vacíos")
        if len(vacios_df) > 0:
            st.warning(f"Se encontraron **{len(vacios_df)}** filas con al menos una celda vacía.")
            st.dataframe(vacios_df, use_container_width=True, height=300)
        else:
            st.success("✅ No se encontraron celdas vacías.")

    st.markdown("---")
    st.markdown("#### 📊 Resumen de vacíos por columna")
    vacios_col = df.isnull().sum().reset_index()
    vacios_col.columns = ["Columna", "Celdas Vacías"]
    vacios_col = vacios_col[vacios_col["Celdas Vacías"] > 0].sort_values("Celdas Vacías", ascending=False)

    if not vacios_col.empty:
        st.dataframe(vacios_col, use_container_width=True)
        st.bar_chart(vacios_col.set_index("Columna"))
    else:
        st.success("✅ Todas las columnas tienen datos completos.")


# ══════════════════════════════════════════════
# TAB 3: DESCARGAR
# ══════════════════════════════════════════════
with tab3:
    st.markdown("### 💾 Exportar Datos")

    col_d1, col_d2 = st.columns(2)

    with col_d1:
        st.markdown("#### 📥 Datos Filtrados")
        st.markdown(f"Exportar **{len(df_filtrado):,}** registros con los filtros aplicados.")
        excel_filtrado = convertir_a_excel(df_filtrado)
        st.download_button(
            label="⬇️ Descargar datos filtrados (.xlsx)",
            data=excel_filtrado,
            file_name=f"datos_filtrados_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    with col_d2:
        st.markdown("#### 🧹 Datos Limpios (sin duplicados ni vacíos)")
        df_limpio = df.drop_duplicates().dropna()
        st.markdown(f"Exportar **{len(df_limpio):,}** registros sin errores.")
        excel_limpio = convertir_a_excel(df_limpio)
        st.download_button(
            label="⬇️ Descargar datos limpios (.xlsx)",
            data=excel_limpio,
            file_name=f"datos_limpios_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    st.markdown("---")
    st.info("🔒 Los archivos descargados se guardan en tu computadora. Nunca se envían a internet.", icon="🛡️")
