import streamlit as st
import pandas as pd
import openpyxl
import io
import re

st.set_page_config(
    page_title="Pliego Tarifario SAESA",
    page_icon="⚡",
    layout="wide",
)

# ── Estilos ────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
    .stApp { background: #f5f7fa; }
    h1 { color: #4a0e8f; }
    .metric-card {
        background: white;
        border-radius: 12px;
        padding: 16px 20px;
        box-shadow: 0 2px 8px rgba(0,0,0,.08);
        text-align: center;
    }
    .metric-card .label { font-size: .8rem; color: #666; margin-bottom: 4px; }
    .metric-card .value { font-size: 1.4rem; font-weight: 700; color: #4a0e8f; }
    .tariff-header {
        background: linear-gradient(135deg, #4a0e8f, #7c3aed);
        color: white;
        padding: 8px 16px;
        border-radius: 8px;
        font-weight: 600;
        margin: 8px 0 4px 0;
    }
</style>
""", unsafe_allow_html=True)

# ── Procesamiento del Excel ────────────────────────────────────────────────────
@st.cache_data
def procesar_excel(file_bytes: bytes) -> tuple[pd.DataFrame, pd.DataFrame]:
    """
    Transforma el pliego tarifario pivoteado en una tabla normalizada (base de datos).
    Retorna (df_tarifas, df_inyecciones).
    """
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes))
    ws = wb.active

    max_col = ws.max_column  # 93

    # ── 1. Extraer localidades (fila 5, columnas pares desde D=4) ──────────────
    localidades = {}          # col_index -> nombre localidad
    for col in range(4, max_col + 1, 2):
        val = ws.cell(5, col).value
        if val and isinstance(val, str):
            localidades[col] = val.strip()

    # ── 2. Identificar grupos de tarifas (filas donde col A es int y col C es None) ──
    # Filas de encabezado de tarifa: col C es None pero col B tiene el nombre
    TARIFF_DATA_END = 210  # filas de datos de tarifas

    tariff_groups = {}   # row -> nombre_tarifa
    current_tariff = None

    rows = []

    for r in range(8, TARIFF_DATA_END):
        a = ws.cell(r, 1).value
        b = ws.cell(r, 2).value
        c = ws.cell(r, 3).value

        if b is None:
            continue

        b_str = str(b).strip()

        # Encabezado de bloque tarifario: sin unidad en col C
        if c is None and b_str:
            current_tariff = b_str
            continue

        if c is None:
            continue

        if current_tariff is None:
            continue

        # Fila de dato real
        concepto = b_str.strip()
        unidad = str(c).strip() if c else ""

        for loc_col, localidad in localidades.items():
            neto = ws.cell(r, loc_col).value
            civa = ws.cell(r, loc_col + 1).value

            # Separar comuna y tipo de suministro
            parts = localidad.rsplit(" - ", 1)
            comuna = parts[0] if len(parts) == 2 else localidad
            tipo_suministro = parts[1] if len(parts) == 2 else ""

            rows.append({
                "tarifa": current_tariff,
                "n_concepto": a,
                "concepto": concepto,
                "unidad": unidad,
                "localidad": localidad,
                "comuna": comuna,
                "tipo_suministro": tipo_suministro,
                "valor_neto": round(float(neto), 4) if isinstance(neto, (int, float)) else None,
                "valor_civa": round(float(civa), 4) if isinstance(civa, (int, float)) else None,
            })

    df_tarifas = pd.DataFrame(rows)

    # ── 3. Sección de inyecciones (filas 234-235) ──────────────────────────────
    inj_rows = []
    for r in [234, 235]:
        concepto = ws.cell(r, 2).value
        unidad = ws.cell(r, 3).value
        if not concepto:
            continue
        for loc_col, localidad in localidades.items():
            neto = ws.cell(r, loc_col).value
            parts = localidad.rsplit(" - ", 1)
            inj_rows.append({
                "concepto": str(concepto).strip(),
                "unidad": str(unidad).strip() if unidad else "",
                "localidad": localidad,
                "comuna": parts[0] if len(parts) == 2 else localidad,
                "tipo_suministro": parts[1] if len(parts) == 2 else "",
                "valor_neto": round(float(neto), 4) if isinstance(neto, (int, float)) else None,
            })

    df_inyecciones = pd.DataFrame(inj_rows)

    return df_tarifas, df_inyecciones


# ── UI ─────────────────────────────────────────────────────────────────────────
st.markdown("# ⚡ Pliego Tarifario SAESA")
st.markdown("Carga el Excel del pliego tarifario para consultar y filtrar tarifas por localidad.")

uploaded = st.file_uploader(
    "Sube el archivo Excel del pliego tarifario",
    type=["xlsx"],
    help="Archivo: Tarifas_de_Suministro_Regulado_*.xlsx"
)

if not uploaded:
    st.info("👆 Sube el archivo Excel para comenzar.")
    st.stop()

file_bytes = uploaded.read()

with st.spinner("Procesando pliego tarifario..."):
    df_tarifas, df_iny = procesar_excel(file_bytes)

# ── Métricas resumen ───────────────────────────────────────────────────────────
localidades_list = sorted(df_tarifas["localidad"].unique())
tarifas_list = sorted(df_tarifas["tarifa"].unique())
comunas_list = sorted(df_tarifas["comuna"].unique())

c1, c2, c3, c4 = st.columns(4)
with c1:
    st.markdown(f"""<div class="metric-card">
        <div class="label">Localidades</div>
        <div class="value">{len(localidades_list)}</div>
    </div>""", unsafe_allow_html=True)
with c2:
    st.markdown(f"""<div class="metric-card">
        <div class="label">Tarifas</div>
        <div class="value">{len(tarifas_list)}</div>
    </div>""", unsafe_allow_html=True)
with c3:
    st.markdown(f"""<div class="metric-card">
        <div class="label">Conceptos</div>
        <div class="value">{df_tarifas["concepto"].nunique()}</div>
    </div>""", unsafe_allow_html=True)
with c4:
    st.markdown(f"""<div class="metric-card">
        <div class="label">Registros totales</div>
        <div class="value">{len(df_tarifas):,}</div>
    </div>""", unsafe_allow_html=True)

st.markdown("---")

# ── Tabs ───────────────────────────────────────────────────────────────────────
tab1, tab2, tab3, tab4 = st.tabs([
    "🔍 Consulta por Localidad",
    "📊 Comparar Localidades",
    "🗄️ Base de Datos Completa",
    "🔋 Inyecciones de Energía",
])

# ══════════════════════════════════════════════════════════════════════════════
# TAB 1 – Consulta por localidad
# ══════════════════════════════════════════════════════════════════════════════
with tab1:
    st.subheader("Consulta de tarifas por localidad")

    col_f1, col_f2, col_f3 = st.columns([2, 2, 1])
    with col_f1:
        sel_localidad = st.selectbox("Localidad", localidades_list, key="loc1")
    with col_f2:
        sel_tarifa_opt = ["Todas"] + tarifas_list
        sel_tarifa = st.selectbox("Tarifa", sel_tarifa_opt, key="tar1")
    with col_f3:
        precio_tipo = st.radio("Precio", ["Neto", "C/IVA"], horizontal=True)

    precio_col = "valor_neto" if precio_tipo == "Neto" else "valor_civa"

    df_loc = df_tarifas[df_tarifas["localidad"] == sel_localidad].copy()
    if sel_tarifa != "Todas":
        df_loc = df_loc[df_loc["tarifa"] == sel_tarifa]

    if df_loc.empty:
        st.warning("No hay datos para la selección.")
    else:
        for tarifa_name, grp in df_loc.groupby("tarifa", sort=False):
            st.markdown(f'<div class="tariff-header">📋 {tarifa_name}</div>', unsafe_allow_html=True)
            show = grp[["concepto", "unidad", precio_col]].copy()
            show.columns = ["Concepto", "Unidad", f"Valor {precio_tipo} ($)"]
            show = show.reset_index(drop=True)
            st.dataframe(
                show.style.format({f"Valor {precio_tipo} ($)": lambda x: f"{x:,.3f}" if pd.notna(x) else "-"}),
                use_container_width=True,
                hide_index=True,
            )

# ══════════════════════════════════════════════════════════════════════════════
# TAB 2 – Comparar localidades
# ══════════════════════════════════════════════════════════════════════════════
with tab2:
    st.subheader("Comparación de tarifas entre localidades")

    col_a, col_b = st.columns(2)
    with col_a:
        locs_sel = st.multiselect(
            "Selecciona localidades (máx. 6)",
            localidades_list,
            default=localidades_list[:3],
            max_selections=6,
            key="locs_comp"
        )
    with col_b:
        tar_comp = st.selectbox("Tarifa a comparar", tarifas_list, key="tar_comp")
        precio_comp = st.radio("Precio", ["Neto", "C/IVA"], horizontal=True, key="precio_comp")

    if not locs_sel:
        st.info("Selecciona al menos una localidad.")
    else:
        precio_col2 = "valor_neto" if precio_comp == "Neto" else "valor_civa"
        df_comp = df_tarifas[
            (df_tarifas["tarifa"] == tar_comp) &
            (df_tarifas["localidad"].isin(locs_sel))
        ].copy()

        if df_comp.empty:
            st.warning("Sin datos para la selección.")
        else:
            pivot = df_comp.pivot_table(
                index=["concepto", "unidad"],
                columns="localidad",
                values=precio_col2,
                aggfunc="first"
            ).reset_index()
            pivot.columns.name = None

            # Formatear columnas numéricas
            num_cols = [c for c in pivot.columns if c not in ["concepto", "unidad"]]
            fmt = {c: lambda x: f"{x:,.3f}" if pd.notna(x) else "-" for c in num_cols}

            st.dataframe(
                pivot.style.format(fmt),
                use_container_width=True,
                hide_index=True,
            )

# ══════════════════════════════════════════════════════════════════════════════
# TAB 3 – Base de datos completa
# ══════════════════════════════════════════════════════════════════════════════
with tab3:
    st.subheader("Base de datos normalizada completa")

    col_f1, col_f2, col_f3 = st.columns(3)
    with col_f1:
        f_comunas = st.multiselect("Filtrar por comuna", comunas_list, key="f_com")
    with col_f2:
        f_tarifas = st.multiselect("Filtrar por tarifa", tarifas_list, key="f_tar")
    with col_f3:
        txt_busq = st.text_input("Buscar concepto", key="f_txt", placeholder="ej: energía, potencia...")

    df_full = df_tarifas.copy()
    if f_comunas:
        df_full = df_full[df_full["comuna"].isin(f_comunas)]
    if f_tarifas:
        df_full = df_full[df_full["tarifa"].isin(f_tarifas)]
    if txt_busq:
        df_full = df_full[df_full["concepto"].str.contains(txt_busq, case=False, na=False)]

    st.caption(f"Mostrando {len(df_full):,} registros")

    st.dataframe(
        df_full[[
            "tarifa", "concepto", "unidad",
            "comuna", "tipo_suministro",
            "valor_neto", "valor_civa"
        ]].style.format({
            "valor_neto": lambda x: f"{x:,.3f}" if pd.notna(x) else "-",
            "valor_civa": lambda x: f"{x:,.3f}" if pd.notna(x) else "-",
        }),
        use_container_width=True,
        hide_index=True,
        height=500,
    )

    # Descargar CSV
    csv_bytes = df_full.to_csv(index=False).encode("utf-8-sig")
    st.download_button(
        "⬇️ Descargar CSV filtrado",
        data=csv_bytes,
        file_name="tarifas_saesa_filtrado.csv",
        mime="text/csv",
    )

# ══════════════════════════════════════════════════════════════════════════════
# TAB 4 – Inyecciones
# ══════════════════════════════════════════════════════════════════════════════
with tab4:
    st.subheader("Precios para valorización de inyecciones de energía")
    st.caption("Vigentes desde el 01 de abril de 2026 – Valores netos (sin IVA)")

    if df_iny.empty:
        st.warning("No se encontraron datos de inyecciones.")
    else:
        f_loc_iny = st.multiselect(
            "Filtrar por localidad",
            sorted(df_iny["localidad"].unique()),
            key="f_iny",
        )
        df_iny_f = df_iny if not f_loc_iny else df_iny[df_iny["localidad"].isin(f_loc_iny)]

        pivot_iny = df_iny_f.pivot_table(
            index=["concepto", "unidad"],
            columns="localidad",
            values="valor_neto",
            aggfunc="first",
        ).reset_index()
        pivot_iny.columns.name = None

        num_cols_iny = [c for c in pivot_iny.columns if c not in ["concepto", "unidad"]]
        fmt_iny = {c: lambda x: f"{x:,.3f}" if pd.notna(x) else "-" for c in num_cols_iny}

        st.dataframe(
            pivot_iny.style.format(fmt_iny),
            use_container_width=True,
            hide_index=True,
        )
