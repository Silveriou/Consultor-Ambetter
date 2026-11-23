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
    return analizar_polizas(df)


def analizar_polizas(df_raw):
    """Calcula estados de deuda, saldo y vigencia de cada póliza."""
    df = df_raw.copy()

    # 1. Limpieza de Member Responsibility (cuota mensual del cliente)
    if "Member Responsibility" in df.columns:
        df["Cuota Mensual ($)"] = (
            df["Member Responsibility"]
            .astype(str)
            .str.replace("$", "", regex=False)
            .str.replace(",", "", regex=False)
            .str.strip()
        )
        df["Cuota Mensual ($)"] = pd.to_numeric(df["Cuota Mensual ($)"], errors="coerce").fillna(0.0)
    else:
        df["Cuota Mensual ($)"] = 0.0

    # 2. Parsing de fechas clave
    for col_f in ["Paid Through Date", "Policy Effective Date", "Policy Term Date"]:
        if col_f in df.columns:
            df[f"_{col_f}_dt"] = pd.to_datetime(df[col_f], errors="coerce")
        else:
            df[f"_{col_f}_dt"] = pd.NaT

    hoy = pd.Timestamp.now().normalize()
    primer_dia_mes_actual = pd.Timestamp(year=hoy.year, month=hoy.month, day=1)

    # 3. Cálculo de Estado de Saldo y Deuda según reglas de Ambetter
    def calc_saldo(row):
        cuota = float(row.get("Cuota Mensual ($)", 0.0))
        paid_dt = row.get("_Paid Through Date_dt", pd.NaT)
        eff_dt = row.get("_Policy Effective Date_dt", pd.NaT)
        term_dt = row.get("_Policy Term Date_dt", pd.NaT)

        # Regla 1: Plan sin costo mensual (subsidio 100%)
        if cuota == 0.0:
            return "En Cero ($0 Cuota)", 0.0, "Plan con subsidio 100% / Sin costo mensual"

        # Regla 2: Si nunca se registró pago inicial (Binder)
        if pd.isna(paid_dt):
            return "Deuda (Pago Inicial Pendiente)", cuota, f"Falta primer pago de ${cuota:,.2f} para activar póliza"

        # Regla 3: Si Paid Through Date es IGUAL o MAYOR a Policy Term Date
        # ¡La póliza fue pagada completamente durante toda su vigencia! Saldo $0.00 exacto
        if pd.notna(term_dt) and paid_dt >= term_dt:
            return "Al Día (En Cero)", 0.0, f"Pagada al 100% hasta fin de vigencia ({paid_dt.strftime('%m/%d/%Y')})"

        # Regla 4: Pólizas vigentes que están al día o adelantadas
        if paid_dt >= primer_dia_mes_actual:
            if paid_dt >= primer_dia_mes_actual + pd.DateOffset(months=1):
                return "A Favor (Pagado Adelantado)", 0.0, f"Cubierto por adelantado hasta {paid_dt.strftime('%m/%d/%Y')}"
            else:
                return "Al Día (En Cero)", 0.0, f"Cuota al día (pagado hasta {paid_dt.strftime('%m/%d/%Y')})"

        # Regla 5: Póliza que terminó en el pasado
        limite_cobro = term_dt if (pd.notna(term_dt) and term_dt < primer_dia_mes_actual) else primer_dia_mes_actual
        if paid_dt >= limite_cobro:
            return "Al Día (En Cero)", 0.0, f"Pagada completamente hasta fin de vigencia ({paid_dt.strftime('%m/%d/%Y')})"

        # Si realmente quedó debiendo meses de vigencia activa
        meses_atraso = max(1, (limite_cobro.year - paid_dt.year) * 12 + (limite_cobro.month - paid_dt.month))
        deuda = round(meses_atraso * cuota, 2)
        return "Con Deuda (Atrasado)", deuda, f"Debe ~${deuda:,.2f} ({meses_atraso} mes(es) pendientes)"

    res_saldo = df.apply(calc_saldo, axis=1)
    df["Estado de Saldo"] = [r[0] for r in res_saldo]
    df["Monto Deuda Estimada ($)"] = [r[1] for r in res_saldo]
    df["Detalle de Deuda"] = [r[2] for r in res_saldo]

    # 4. Cálculo de Estatus de Póliza
    def calc_estatus(row):
        eff_dt = row.get("_Policy Effective Date_dt", pd.NaT)
        term_dt = row.get("_Policy Term Date_dt", pd.NaT)
        paid_dt = row.get("_Paid Through Date_dt", pd.NaT)
        estado_saldo = row.get("Estado de Saldo", "")

        if pd.isna(eff_dt) and pd.isna(paid_dt):
            return "Pendiente de Emisión / Binder"
        if pd.notna(term_dt) and term_dt < hoy:
            return "Terminada / Expirada"
        if estado_saldo in ["Con Deuda (Atrasado)", "Deuda (Pago Inicial Pendiente)"]:
            return "En Período de Gracia (Riesgo)"
        return "Activa"

    df["Estatus de Póliza"] = df.apply(calc_estatus, axis=1)

    # 5. Nombre Completo
    if "Insured First Name" in df.columns and "Insured Last Name" in df.columns:
        df["Cliente"] = (
            df["Insured First Name"].fillna("").astype(str) + " " +
            df["Insured Last Name"].fillna("").astype(str)
        ).str.strip()
    elif "Insured Last Name" in df.columns:
        df["Cliente"] = df["Insured Last Name"].fillna("").astype(str)
    else:
        df["Cliente"] = "Cliente"

    return df


def detectar_errores(df):
    """Detecta filas duplicadas y celdas vacías en columnas clave."""
    cols_base = [c for c in df.columns if not c.startswith("_")]
    df_clean = df[cols_base]
    duplicados = df_clean[df_clean.duplicated(keep=False)]
    vacios = df_clean[df_clean.isnull().any(axis=1)]
    return duplicados, vacios


def convertir_a_excel(df):
    """Convierte un DataFrame a bytes de Excel para descarga."""
    cols_export = [c for c in df.columns if not c.startswith("_")]
    df_export = df[cols_export].copy()

    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
        df_export.to_excel(writer, index=False, sheet_name="Datos")
        workbook  = writer.book
        worksheet = writer.sheets["Datos"]
        fmt_header = workbook.add_format({
            "bold": True, "bg_color": "#1a3c5e",
            "font_color": "white", "border": 1
        })
        for col_num, col_name in enumerate(df_export.columns):
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
total_columnas    = len([c for c in df.columns if not c.startswith("_")])
total_duplicados  = len(duplicados_df)

con_deuda_mask = df["Estado de Saldo"] == "Con Deuda (Atrasado)"
pendiente_binder_mask = df["Estado de Saldo"] == "Deuda (Pago Inicial Pendiente)"
al_dia_mask = df["Estado de Saldo"].isin(["Al Día (En Cero)", "En Cero ($0 Cuota)", "A Favor (Pagado Adelantado)"])

total_con_deuda = int(con_deuda_mask.sum())
total_pendiente_binder = int(pendiente_binder_mask.sum())
total_al_dia = int(al_dia_mask.sum())
monto_deuda_total = float(df["Monto Deuda Estimada ($)"].sum())

col1, col2, col3, col4 = st.columns(4)
col1.metric("👥 Total de Pólizas", f"{total_clientes:,}")
col2.metric(
    "🔴 Con Deuda / Atrasadas",
    f"{total_con_deuda:,}",
    delta=f"${monto_deuda_total:,.2f} deuda total" if total_con_deuda > 0 else "Sin deudas",
    delta_color="inverse"
)
col3.metric(
    "🟡 Pendiente Pago Inicial",
    f"{total_pendiente_binder:,}",
    delta="Binder requerido" if total_pendiente_binder > 0 else "Todo emitido",
    delta_color="off"
)
col4.metric(
    "🟢 Al Día / Cuota $0",
    f"{total_al_dia:,}",
    delta=f"{round((total_al_dia / total_clientes) * 100 if total_clientes > 0 else 0)}% del total",
    delta_color="normal"
)

st.markdown("---")

# ─────────────────────────────────────────────
# PESTAÑAS PRINCIPALES
# ─────────────────────────────────────────────
tab0, tab1, tab2, tab3 = st.tabs([
    "💰 Saldos, Deudas y Estatus",
    "📋 Todos los Datos",
    "🚨 Detección de Errores",
    "💾 Exportar Reportes"
])


# ══════════════════════════════════════════════
# TAB 0: SALDOS, DEUDAS Y ESTATUS
# ══════════════════════════════════════════════
with tab0:
    st.markdown("### 💳 Control Financiero y Estatus de Pólizas")
    st.caption("Identifica inmediatamente quién debe, quién está al día y qué pólizas están en riesgo.")

    # Filtros rápidos
    col_f1, col_f2, col_f3 = st.columns([1.5, 1.5, 2])
    with col_f1:
        opciones_saldo = ["(Todos los saldos)"] + sorted(df["Estado de Saldo"].dropna().unique().tolist())
        filtro_saldo = st.selectbox("🎯 Filtrar por Estado de Saldo:", opciones_saldo)
    with col_f2:
        opciones_estatus = ["(Todos los estatus)"] + sorted(df["Estatus de Póliza"].dropna().unique().tolist())
        filtro_estatus = st.selectbox("🛡️ Filtrar por Estatus de Póliza:", opciones_estatus)
    with col_f3:
        filtro_texto_fin = st.text_input("🔍 Buscar cliente, póliza o teléfono:", placeholder="Ej: Perez, U724..., 773...")

    # Aplicar filtros específicos de finanzas
    df_fin = df.copy()
    if filtro_saldo != "(Todos los saldos)":
        df_fin = df_fin[df_fin["Estado de Saldo"] == filtro_saldo]
    if filtro_estatus != "(Todos los estatus)":
        df_fin = df_fin[df_fin["Estatus de Póliza"] == filtro_estatus]
    if filtro_texto_fin:
        mask_t = df_fin.apply(
            lambda r: r.astype(str).str.contains(filtro_texto_fin, case=False, na=False).any(),
            axis=1
        )
        df_fin = df_fin[mask_t]

    # Columnas prioritarias para mostrar en la tabla financiera
    cols_deseadas = [
        "Cliente",
        "Policy Number",
        "Estado de Saldo",
        "Monto Deuda Estimada ($)",
        "Cuota Mensual ($)",
        "Paid Through Date",
        "Estatus de Póliza",
        "Autopay",
        "Member Phone Number",
        "Member Email",
        "Plan Name",
        "State",
        "County",
        "Renewal Type"
    ]
    cols_existentes = [c for c in cols_deseadas if c in df_fin.columns]

    st.markdown(f"**Mostrando {len(df_fin):,} pólizas filtradas** (Deuda acumulada en esta vista: **${df_fin['Monto Deuda Estimada ($)'].sum():,.2f}**)")

    st.dataframe(
        df_fin[cols_existentes],
        use_container_width=True,
        height=450
    )

    # Botón directo para descargar los deudores
    df_deudores = df[df["Estado de Saldo"].isin(["Con Deuda (Atrasado)", "Deuda (Pago Inicial Pendiente)"])]
    if len(df_deudores) > 0:
        excel_deudores = convertir_a_excel(df_deudores[cols_existentes])
        st.download_button(
            label=f"📥 Descargar lista de cobro ({len(df_deudores)} clientes con deuda o binder pendiente) en Excel",
            data=excel_deudores,
            file_name=f"clientes_con_deuda_ambetter_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )


# ══════════════════════════════════════════════
# TAB 1: DATOS Y FILTROS GENERALES
# ══════════════════════════════════════════════
with tab1:
    st.markdown("### 🔍 Explorador Completo de Datos")

    # Buscador global por texto
    busqueda = st.text_input(
        "🔎 Búsqueda general en todas las columnas",
        placeholder="Escribe cualquier texto para buscar...",
        key="busqueda_tab1"
    )

    # Columnas visibles (excluir internas que empiezan con _)
    cols_visibles = [c for c in df.columns if not c.startswith("_")]

    # Filtros dinámicos por columna
    with st.expander("⚙️ Filtros avanzados por columna", expanded=False):
        filtros = {}
        cols_filtro = st.columns(min(3, len(cols_visibles)))
        for i, col in enumerate(cols_visibles[:9]):  # máx 9 columnas
            with cols_filtro[i % 3]:
                try:
                    col_num = pd.to_numeric(df[col], errors="coerce")
                    es_numerico = (df[col].dtype != object) and col_num.notna().any()
                    if not es_numerico or df[col].nunique() <= 30:
                        opciones = ["(Todos)"] + sorted(df[col].dropna().astype(str).unique().tolist())
                        sel = st.selectbox(f"**{col}**", opciones, key=f"filtro_{col}")
                        if sel != "(Todos)":
                            filtros[col] = sel
                    else:
                        col_data = col_num.dropna()
                        if len(col_data) == 0:
                            st.caption(f"**{col}**: sin datos")
                        else:
                            min_val = float(col_data.min())
                            max_val = float(col_data.max())
                            if pd.isna(min_val) or pd.isna(max_val) or min_val == max_val:
                                st.caption(f"**{col}**: {min_val}")
                            else:
                                rango = st.slider(
                                    f"**{col}**",
                                    min_value=min_val, max_value=max_val,
                                    value=(min_val, max_val),
                                    key=f"filtro_{col}"
                                )
                                if rango != (min_val, max_val):
                                    filtros[col] = rango
                except Exception:
                    pass

    # Aplicar búsqueda global
    df_filtrado = df[cols_visibles].copy()
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
    cols_base = [c for c in df.columns if not c.startswith("_")]
    vacios_col = df[cols_base].isnull().sum().reset_index()
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
    st.markdown("### 💾 Exportar Datos y Reportes")

    col_d1, col_d2 = st.columns(2)

    with col_d1:
        st.markdown("#### 📥 Datos Filtrados con Cálculos")
        st.markdown(f"Exportar **{len(df_filtrado):,}** registros con el estado de saldo, deuda y vigencia calculados.")
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
        cols_clean = [c for c in df.columns if not c.startswith("_")]
        df_limpio = df[cols_clean].drop_duplicates().dropna()
        st.markdown(f"Exportar **{len(df_limpio):,}** registros sin celdas vacías.")
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
