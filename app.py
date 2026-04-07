import streamlit as st
import pandas as pd
import openpyxl
import io
import re
from datetime import datetime

st.set_page_config(
    page_title="Pliego Tarifario SAESA",
    page_icon="⚡",
    layout="wide",
)

st.markdown("""
<style>
    .stApp { background: #f5f7fa; }
    h1 { color: #4a0e8f; }
    h2, h3 { color: #3b0764; }
    .metric-card {
        background: white;
        border-radius: 12px;
        padding: 16px 20px;
        box-shadow: 0 2px 8px rgba(0,0,0,.08);
        text-align: center;
        margin-bottom: 8px;
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
    .alert-box {
        background: #fef3c7;
        border-left: 4px solid #f59e0b;
        padding: 10px 16px;
        border-radius: 0 8px 8px 0;
        margin: 4px 0;
    }
    .period-badge {
        display: inline-block;
        background: #ede9fe;
        color: #4a0e8f;
        border-radius: 20px;
        padding: 2px 12px;
        font-size: .8rem;
        font-weight: 600;
        margin: 2px;
    }
</style>
""", unsafe_allow_html=True)

# ── Meses en español para ordenar ─────────────────────────────────────────────
MESES = {
    "enero":1,"febrero":2,"marzo":3,"abril":4,"mayo":5,"junio":6,
    "julio":7,"agosto":8,"septiembre":9,"octubre":10,"noviembre":11,"diciembre":12
}

def extraer_periodo(ws) -> str:
    """Lee el título del pliego (B3) y extrae 'mes año', ej: 'abril 2026'."""
    titulo = ws.cell(3, 2).value or ""
    titulo = titulo.lower()
    for mes in MESES:
        if mes in titulo:
            m = re.search(r"\d{4}", titulo)
            year = m.group() if m else "?"
            return f"{mes} {year}"
    return "desconocido"

def periodo_sort_key(periodo: str):
    parts = periodo.lower().split()
    if len(parts) == 2:
        mes, year = parts
        return (int(year), MESES.get(mes, 0))
    return (9999, 0)

@st.cache_data(show_spinner=False)
def procesar_excel(file_bytes: bytes) -> tuple[pd.DataFrame, pd.DataFrame, str]:
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes))
    ws = wb.active
    max_col = ws.max_column
    periodo = extraer_periodo(ws)

    # Localidades: fila 5, columnas pares desde D=4
    localidades = {}
    for col in range(4, max_col + 1, 2):
        val = ws.cell(5, col).value
        if val and isinstance(val, str):
            localidades[col] = val.strip()

    TARIFF_DATA_END = 210
    current_tariff = None
    rows = []

    for r in range(8, TARIFF_DATA_END):
        a = ws.cell(r, 1).value
        b = ws.cell(r, 2).value
        c = ws.cell(r, 3).value
        if b is None:
            continue
        b_str = str(b).strip()
        if c is None and b_str:
            current_tariff = b_str
            continue
        if c is None or current_tariff is None:
            continue
        concepto = b_str.strip()
        unidad = str(c).strip() if c else ""
        for loc_col, localidad in localidades.items():
            neto = ws.cell(r, loc_col).value
            civa = ws.cell(r, loc_col + 1).value
            parts = localidad.rsplit(" - ", 1)
            rows.append({
                "periodo": periodo,
                "tarifa": current_tariff,
                "n_concepto": a,
                "concepto": concepto,
                "unidad": unidad,
                "localidad": localidad,
                "comuna": parts[0] if len(parts) == 2 else localidad,
                "tipo_suministro": parts[1] if len(parts) == 2 else "",
                "valor_neto": round(float(neto), 4) if isinstance(neto, (int, float)) else None,
                "valor_civa": round(float(civa), 4) if isinstance(civa, (int, float)) else None,
            })

    df_tarifas = pd.DataFrame(rows)

    # Inyecciones
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
                "periodo": periodo,
                "concepto": str(concepto).strip(),
                "unidad": str(unidad).strip() if unidad else "",
                "localidad": localidad,
                "comuna": parts[0] if len(parts) == 2 else localidad,
                "tipo_suministro": parts[1] if len(parts) == 2 else "",
                "valor_neto": round(float(neto), 4) if isinstance(neto, (int, float)) else None,
            })

    df_inyecciones = pd.DataFrame(inj_rows)
    return df_tarifas, df_inyecciones, periodo


# ══════════════════════════════════════════════════════════════════════════════
# SIDEBAR – Carga de archivos
# ══════════════════════════════════════════════════════════════════════════════
with st.sidebar:
    st.markdown("## ⚡ SAESA Tarifas")
    st.markdown("---")
    st.markdown("### 📂 Cargar pliegos")
    uploaded_files = st.file_uploader(
        "Sube uno o más archivos Excel",
        type=["xlsx"],
        accept_multiple_files=True,
        help="Cada archivo corresponde a un período tarifario distinto."
    )
    if uploaded_files:
        st.success(f"{len(uploaded_files)} archivo(s) cargado(s)")

st.markdown("# ⚡ Pliego Tarifario SAESA")

if not uploaded_files:
    st.info("👈 Sube uno o más archivos Excel del pliego tarifario en el panel lateral para comenzar.")
    st.stop()

# ── Procesar todos los archivos ────────────────────────────────────────────────
all_tarifas = []
all_inyecciones = []
periodos_cargados = []

progress = st.progress(0, text="Procesando archivos...")
for i, f in enumerate(uploaded_files):
    df_t, df_i, periodo = procesar_excel(f.read())
    all_tarifas.append(df_t)
    all_inyecciones.append(df_i)
    periodos_cargados.append(periodo)
    progress.progress((i + 1) / len(uploaded_files), text=f"Procesado: {periodo}")

progress.empty()

df_all = pd.concat(all_tarifas, ignore_index=True)
df_iny_all = pd.concat(all_inyecciones, ignore_index=True)

# Ordenar períodos cronológicamente
periodos_ordenados = sorted(set(df_all["periodo"].unique()), key=periodo_sort_key)
df_all["periodo"] = pd.Categorical(df_all["periodo"], categories=periodos_ordenados, ordered=True)
df_all = df_all.sort_values("periodo")

localidades_list  = sorted(df_all["localidad"].unique())
comunas_list      = sorted(df_all["comuna"].unique())
tarifas_list      = sorted(df_all["tarifa"].unique())

# ── Métricas ───────────────────────────────────────────────────────────────────
c1, c2, c3, c4, c5 = st.columns(5)
for col, label, val in zip(
    [c1, c2, c3, c4, c5],
    ["Períodos", "Localidades", "Tarifas", "Conceptos", "Registros"],
    [len(periodos_ordenados), df_all["localidad"].nunique(),
     df_all["tarifa"].nunique(), df_all["concepto"].nunique(), len(df_all)]
):
    col.markdown(f"""<div class="metric-card">
        <div class="label">{label}</div>
        <div class="value">{val:,}</div>
    </div>""", unsafe_allow_html=True)

# Períodos badge
badges = " ".join(f'<span class="period-badge">{p}</span>' for p in periodos_ordenados)
st.markdown(f"**Períodos cargados:** {badges}", unsafe_allow_html=True)
st.markdown("---")

# ══════════════════════════════════════════════════════════════════════════════
# TABS
# ══════════════════════════════════════════════════════════════════════════════
tab_labels = ["🔍 Consulta por Localidad", "📊 Comparar Localidades",
              "📈 Evolución Temporal", "🗄️ Base de Datos", "🔋 Inyecciones"]
tabs = st.tabs(tab_labels)


# ══════════════════════════════════════════════════════════════════════════════
# TAB 1 – Consulta por localidad
# ══════════════════════════════════════════════════════════════════════════════
with tabs[0]:
    st.subheader("Consulta de tarifas por localidad")
    col1, col2, col3, col4 = st.columns([2, 2, 1, 1])
    with col1:
        sel_loc = st.selectbox("Localidad", localidades_list, key="loc1")
    with col2:
        sel_tar = st.selectbox("Tarifa", ["Todas"] + tarifas_list, key="tar1")
    with col3:
        sel_per = st.selectbox("Período", ["Último"] + periodos_ordenados, key="per1")
    with col4:
        precio_tipo = st.radio("Precio", ["Neto", "C/IVA"], horizontal=True, key="pt1")

    precio_col = "valor_neto" if precio_tipo == "Neto" else "valor_civa"

    if sel_per == "Último":
        periodo_filtro = periodos_ordenados[-1]
    else:
        periodo_filtro = sel_per

    df_loc = df_all[
        (df_all["localidad"] == sel_loc) &
        (df_all["periodo"] == periodo_filtro)
    ].copy()
    if sel_tar != "Todas":
        df_loc = df_loc[df_loc["tarifa"] == sel_tar]

    if df_loc.empty:
        st.warning("No hay datos para la selección.")
    else:
        for tarifa_name, grp in df_loc.groupby("tarifa", sort=False):
            st.markdown(f'<div class="tariff-header">📋 {tarifa_name}</div>', unsafe_allow_html=True)
            show = grp[["concepto", "unidad", precio_col]].copy()
            show.columns = ["Concepto", "Unidad", f"Valor {precio_tipo} ($)"]
            st.dataframe(
                show.reset_index(drop=True).style.format(
                    {f"Valor {precio_tipo} ($)": lambda x: f"{x:,.3f}" if pd.notna(x) else "-"}
                ),
                use_container_width=True, hide_index=True,
            )


# ══════════════════════════════════════════════════════════════════════════════
# TAB 2 – Comparar localidades
# ══════════════════════════════════════════════════════════════════════════════
with tabs[1]:
    st.subheader("Comparación entre localidades")
    col1, col2, col3 = st.columns([3, 2, 1])
    with col1:
        locs_sel = st.multiselect("Localidades (máx 6)", localidades_list,
                                   default=localidades_list[:3], max_selections=6, key="locs2")
    with col2:
        tar_comp = st.selectbox("Tarifa", tarifas_list, key="tar2")
        per_comp = st.selectbox("Período", ["Último"] + periodos_ordenados, key="per2")
    with col3:
        precio_comp = st.radio("Precio", ["Neto", "C/IVA"], horizontal=True, key="pc2")

    precio_col2 = "valor_neto" if precio_comp == "Neto" else "valor_civa"
    periodo_comp = periodos_ordenados[-1] if per_comp == "Último" else per_comp

    if not locs_sel:
        st.info("Selecciona al menos una localidad.")
    else:
        df_comp = df_all[
            (df_all["tarifa"] == tar_comp) &
            (df_all["localidad"].isin(locs_sel)) &
            (df_all["periodo"] == periodo_comp)
        ]
        if df_comp.empty:
            st.warning("Sin datos para la selección.")
        else:
            pivot = df_comp.pivot_table(
                index=["concepto", "unidad"], columns="localidad",
                values=precio_col2, aggfunc="first"
            ).reset_index()
            pivot.columns.name = None
            num_cols = [c for c in pivot.columns if c not in ["concepto", "unidad"]]
            st.dataframe(
                pivot.style.format({c: lambda x: f"{x:,.3f}" if pd.notna(x) else "-" for c in num_cols}),
                use_container_width=True, hide_index=True,
            )


# ══════════════════════════════════════════════════════════════════════════════
# TAB 3 – Evolución temporal
# ══════════════════════════════════════════════════════════════════════════════
with tabs[2]:
    st.subheader("Evolución temporal de tarifas")

    if len(periodos_ordenados) < 2:
        st.info("Carga al menos 2 archivos de distintos períodos para ver la evolución.")
    else:
        subtab1, subtab2, subtab3, subtab4 = st.tabs([
            "📈 Gráfico de evolución",
            "📋 Tabla comparativa mes a mes",
            "📉 Variación porcentual",
            "🚨 Alertas de cambios",
        ])

        # Filtros comunes para evolución
        with st.container():
            ec1, ec2, ec3, ec4 = st.columns([2, 2, 2, 1])
            with ec1:
                evo_loc = st.selectbox("Localidad", localidades_list, key="evo_loc")
            with ec2:
                evo_tar = st.selectbox("Tarifa", tarifas_list, key="evo_tar")
            with ec3:
                conceptos_disp = sorted(
                    df_all[df_all["tarifa"] == evo_tar]["concepto"].unique()
                )
                evo_conc = st.selectbox("Concepto", conceptos_disp, key="evo_conc")
            with ec4:
                evo_precio = st.radio("Precio", ["Neto", "C/IVA"], horizontal=True, key="evo_precio")

        evo_col = "valor_neto" if evo_precio == "Neto" else "valor_civa"

        df_evo = df_all[
            (df_all["localidad"] == evo_loc) &
            (df_all["tarifa"] == evo_tar) &
            (df_all["concepto"] == evo_conc)
        ][["periodo", evo_col]].dropna().copy()
        df_evo = df_evo.sort_values("periodo")

        # ── Gráfico ────────────────────────────────────────────────────────────
        with subtab1:
            if df_evo.empty:
                st.warning("Sin datos para la combinación seleccionada.")
            else:
                import json
                labels = df_evo["periodo"].astype(str).tolist()
                values = df_evo[evo_col].tolist()
                unidad_str = df_all[
                    (df_all["tarifa"] == evo_tar) & (df_all["concepto"] == evo_conc)
                ]["unidad"].iloc[0]

                chart_data = json.dumps({"labels": labels, "values": values})
                title_str = f"{evo_conc} — {evo_loc}"
                ylabel_str = f"$ {evo_precio} ({unidad_str})"

                chart_html = f"""
<!DOCTYPE html>
<html>
<head>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>
</head>
<body style="margin:0;padding:10px;background:#fff;">
<canvas id="myChart" height="110"></canvas>
<script>
const data = {chart_data};
const ctx = document.getElementById('myChart').getContext('2d');
new Chart(ctx, {{
    type: 'line',
    data: {{
        labels: data.labels,
        datasets: [{{
            label: {json.dumps(ylabel_str)},
            data: data.values,
            borderColor: '#4a0e8f',
            backgroundColor: 'rgba(74,14,143,0.08)',
            borderWidth: 3,
            pointBackgroundColor: '#7c3aed',
            pointRadius: 6,
            pointHoverRadius: 9,
            tension: 0.3,
            fill: true,
        }}]
    }},
    options: {{
        responsive: true,
        plugins: {{
            title: {{
                display: true,
                text: {json.dumps(title_str)},
                font: {{ size: 14, weight: 'bold' }},
                color: '#3b0764'
            }},
            legend: {{ display: true }},
            tooltip: {{
                callbacks: {{
                    label: ctx => '$ ' + ctx.parsed.y.toLocaleString('es-CL', {{minimumFractionDigits:3}})
                }}
            }}
        }},
        scales: {{
            y: {{
                title: {{ display: true, text: {json.dumps(ylabel_str)} }},
                ticks: {{
                    callback: v => '$' + v.toLocaleString('es-CL', {{minimumFractionDigits:0}})
                }}
            }}
        }}
    }}
}});
</script>
</body>
</html>"""
                st.components.v1.html(chart_html, height=420)

                # Tabla mini bajo el gráfico
                df_show = df_evo.copy()
                df_show.columns = ["Período", f"Valor {evo_precio} ($)"]
                st.dataframe(
                    df_show.style.format({f"Valor {evo_precio} ($)": "{:,.3f}"}),
                    use_container_width=True, hide_index=True,
                )

        # ── Tabla comparativa mes a mes ────────────────────────────────────────
        with subtab2:
            st.markdown("**Todos los conceptos de la tarifa seleccionada a través del tiempo**")
            df_pivot_time = df_all[
                (df_all["localidad"] == evo_loc) &
                (df_all["tarifa"] == evo_tar)
            ].pivot_table(
                index=["concepto", "unidad"],
                columns="periodo",
                values=evo_col,
                aggfunc="first"
            ).reset_index()
            df_pivot_time.columns.name = None
            num_cols_t = [c for c in df_pivot_time.columns if c not in ["concepto", "unidad"]]
            if df_pivot_time.empty:
                st.warning("Sin datos.")
            else:
                st.dataframe(
                    df_pivot_time.style.format(
                        {c: lambda x: f"{x:,.3f}" if pd.notna(x) else "—" for c in num_cols_t}
                    ).highlight_null(color="#f3f4f6"),
                    use_container_width=True, hide_index=True,
                )

        # ── Variación porcentual ───────────────────────────────────────────────
        with subtab3:
            st.markdown("**Variación % entre períodos consecutivos**")
            if len(periodos_ordenados) < 2:
                st.info("Se necesitan al menos 2 períodos.")
            else:
                col_p1, col_p2 = st.columns(2)
                with col_p1:
                    per_base = st.selectbox("Período base", periodos_ordenados[:-1], key="pbase")
                with col_p2:
                    per_comp2 = st.selectbox(
                        "Período comparar",
                        [p for p in periodos_ordenados if periodo_sort_key(p) > periodo_sort_key(per_base)],
                        key="pcomp2"
                    )

                df_base = df_all[
                    (df_all["localidad"] == evo_loc) &
                    (df_all["tarifa"] == evo_tar) &
                    (df_all["periodo"] == per_base)
                ][["concepto", "unidad", evo_col]].rename(columns={evo_col: "base"})

                df_nuevo = df_all[
                    (df_all["localidad"] == evo_loc) &
                    (df_all["tarifa"] == evo_tar) &
                    (df_all["periodo"] == per_comp2)
                ][["concepto", "unidad", evo_col]].rename(columns={evo_col: "nuevo"})

                df_var = df_base.merge(df_nuevo, on=["concepto", "unidad"], how="outer")
                df_var["variacion_%"] = ((df_var["nuevo"] - df_var["base"]) / df_var["base"] * 100).round(2)

                def color_var(val):
                    if pd.isna(val): return ""
                    if val > 0: return "color: #dc2626; font-weight:600"
                    if val < 0: return "color: #16a34a; font-weight:600"
                    return "color: #6b7280"

                df_var.columns = ["Concepto", "Unidad",
                                   f"$ {per_base}", f"$ {per_comp2}", "Variación %"]
                st.dataframe(
                    df_var.style
                        .format({
                            f"$ {per_base}": lambda x: f"{x:,.3f}" if pd.notna(x) else "—",
                            f"$ {per_comp2}": lambda x: f"{x:,.3f}" if pd.notna(x) else "—",
                            "Variación %": lambda x: f"{x:+.2f}%" if pd.notna(x) else "—",
                        })
                        .applymap(color_var, subset=["Variación %"]),
                    use_container_width=True, hide_index=True,
                )

        # ── Alertas de cambios ─────────────────────────────────────────────────
        with subtab4:
            st.markdown("**Conceptos que cambiaron entre el penúltimo y último período**")

            umbral = st.slider("Umbral mínimo de variación (%)", 0.0, 20.0, 1.0, 0.5, key="umbral")

            per_ant = periodos_ordenados[-2]
            per_ult = periodos_ordenados[-1]

            df_ant_all = df_all[df_all["periodo"] == per_ant][
                ["localidad", "tarifa", "concepto", "unidad", evo_col]
            ].rename(columns={evo_col: "ant"})
            df_ult_all = df_all[df_all["periodo"] == per_ult][
                ["localidad", "tarifa", "concepto", "unidad", evo_col]
            ].rename(columns={evo_col: "ult"})

            df_alert = df_ant_all.merge(df_ult_all, on=["localidad", "tarifa", "concepto", "unidad"], how="outer")
            df_alert["var_%"] = ((df_alert["ult"] - df_alert["ant"]) / df_alert["ant"] * 100).round(2)
            df_alert = df_alert[df_alert["var_%"].abs() >= umbral].dropna(subset=["var_%"])
            df_alert = df_alert.sort_values("var_%", ascending=False)

            # Filtros de alerta
            al1, al2 = st.columns(2)
            with al1:
                alert_loc = st.selectbox("Filtrar localidad", ["Todas"] + comunas_list, key="al_loc")
            with al2:
                alert_tar = st.selectbox("Filtrar tarifa", ["Todas"] + tarifas_list, key="al_tar")

            if alert_loc != "Todas":
                df_alert = df_alert[df_alert["localidad"].str.startswith(alert_loc)]
            if alert_tar != "Todas":
                df_alert = df_alert[df_alert["tarifa"] == alert_tar]

            st.caption(f"Mostrando {len(df_alert)} conceptos con variación ≥ {umbral}% entre **{per_ant}** y **{per_ult}**")

            if df_alert.empty:
                st.success("✅ No se detectaron variaciones significativas.")
            else:
                df_alert_show = df_alert.rename(columns={
                    "localidad": "Localidad", "tarifa": "Tarifa",
                    "concepto": "Concepto", "unidad": "Unidad",
                    "ant": f"$ {per_ant}", "ult": f"$ {per_ult}", "var_%": "Variación %"
                })

                def color_alert(val):
                    if pd.isna(val): return ""
                    if val > 0: return "background-color: #fee2e2; color: #dc2626; font-weight:600"
                    if val < 0: return "background-color: #dcfce7; color: #16a34a; font-weight:600"
                    return ""

                st.dataframe(
                    df_alert_show.reset_index(drop=True).style
                        .format({
                            f"$ {per_ant}": lambda x: f"{x:,.3f}" if pd.notna(x) else "—",
                            f"$ {per_ult}": lambda x: f"{x:,.3f}" if pd.notna(x) else "—",
                            "Variación %": lambda x: f"{x:+.2f}%" if pd.notna(x) else "—",
                        })
                        .applymap(color_alert, subset=["Variación %"]),
                    use_container_width=True, hide_index=True, height=500,
                )

                csv_alert = df_alert_show.to_csv(index=False).encode("utf-8-sig")
                st.download_button("⬇️ Descargar alertas CSV", csv_alert,
                                   "alertas_tarifarias.csv", "text/csv")


# ══════════════════════════════════════════════════════════════════════════════
# TAB 4 – Base de datos completa
# ══════════════════════════════════════════════════════════════════════════════
with tabs[3]:
    st.subheader("Base de datos normalizada completa")
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        f_per = st.multiselect("Período", periodos_ordenados, key="f_per")
    with col2:
        f_com = st.multiselect("Comuna", comunas_list, key="f_com")
    with col3:
        f_tar = st.multiselect("Tarifa", tarifas_list, key="f_tar")
    with col4:
        txt_busq = st.text_input("Buscar concepto", placeholder="ej: energía...", key="f_txt")

    df_full = df_all.copy()
    if f_per: df_full = df_full[df_full["periodo"].isin(f_per)]
    if f_com: df_full = df_full[df_full["comuna"].isin(f_com)]
    if f_tar: df_full = df_full[df_full["tarifa"].isin(f_tar)]
    if txt_busq: df_full = df_full[df_full["concepto"].str.contains(txt_busq, case=False, na=False)]

    st.caption(f"Mostrando {len(df_full):,} registros")
    st.dataframe(
        df_full[["periodo", "tarifa", "concepto", "unidad",
                 "comuna", "tipo_suministro", "valor_neto", "valor_civa"
                 ]].reset_index(drop=True).style.format({
            "valor_neto": lambda x: f"{x:,.3f}" if pd.notna(x) else "—",
            "valor_civa": lambda x: f"{x:,.3f}" if pd.notna(x) else "—",
        }),
        use_container_width=True, hide_index=True, height=500,
    )

    csv_bytes = df_full.to_csv(index=False).encode("utf-8-sig")
    st.download_button("⬇️ Descargar CSV filtrado", csv_bytes,
                       "tarifas_saesa_filtrado.csv", "text/csv")


# ══════════════════════════════════════════════════════════════════════════════
# TAB 5 – Inyecciones
# ══════════════════════════════════════════════════════════════════════════════
with tabs[4]:
    st.subheader("Precios para valorización de inyecciones de energía")
    st.caption("Valores netos (sin IVA) — Art. 149 bis DFL N°4/2006")

    col1, col2 = st.columns(2)
    with col1:
        per_iny = st.selectbox("Período", ["Último"] + periodos_ordenados, key="per_iny")
    with col2:
        f_loc_iny = st.multiselect("Filtrar localidad", sorted(df_iny_all["localidad"].unique()), key="f_iny")

    periodo_iny = periodos_ordenados[-1] if per_iny == "Último" else per_iny
    df_iny_f = df_iny_all[df_iny_all["periodo"] == periodo_iny]
    if f_loc_iny:
        df_iny_f = df_iny_f[df_iny_f["localidad"].isin(f_loc_iny)]

    if df_iny_f.empty:
        st.warning("Sin datos.")
    else:
        pivot_iny = df_iny_f.pivot_table(
            index=["concepto", "unidad"], columns="localidad",
            values="valor_neto", aggfunc="first"
        ).reset_index()
        pivot_iny.columns.name = None
        num_cols_iny = [c for c in pivot_iny.columns if c not in ["concepto", "unidad"]]
        st.dataframe(
            pivot_iny.style.format({c: lambda x: f"{x:,.3f}" if pd.notna(x) else "—" for c in num_cols_iny}),
            use_container_width=True, hide_index=True,
        )

    if len(periodos_ordenados) >= 2:
        st.markdown("---")
        st.markdown("**Evolución de inyecciones por localidad**")
        loc_iny_evo = st.selectbox("Localidad", sorted(df_iny_all["localidad"].unique()), key="loc_iny_evo")
        df_iny_evo = df_iny_all[df_iny_all["localidad"] == loc_iny_evo].pivot_table(
            index=["concepto", "unidad"], columns="periodo",
            values="valor_neto", aggfunc="first"
        ).reset_index()
        df_iny_evo.columns.name = None
        num_cols_ie = [c for c in df_iny_evo.columns if c not in ["concepto", "unidad"]]
        st.dataframe(
            df_iny_evo.style.format({c: lambda x: f"{x:,.3f}" if pd.notna(x) else "—" for c in num_cols_ie}),
            use_container_width=True, hide_index=True,
        )
