import streamlit as st
import pandas as pd
import openpyxl
import io
import re
import json

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

MESES = {
    "enero": 1, "febrero": 2, "marzo": 3, "abril": 4,
    "mayo": 5, "junio": 6, "julio": 7, "agosto": 8,
    "septiembre": 9, "octubre": 10, "noviembre": 11, "diciembre": 12,
}


def periodo_sort_key(periodo: str):
    parts = periodo.lower().split()
    if len(parts) == 2:
        return (int(parts[1]), MESES.get(parts[0], 0))
    return (9999, 0)


# ══════════════════════════════════════════════════════════════════════════════
# PARSER ADAPTATIVO
# ══════════════════════════════════════════════════════════════════════════════

def _find_titulo_y_periodo(ws) -> str:
    """Busca el título del pliego en cualquier fila/columna y extrae el período."""
    for r in range(1, 25):
        for col in range(1, 10):
            val = ws.cell(r, col).value
            if not val or not isinstance(val, str):
                continue
            val_l = val.lower()
            if "suministro" in val_l and any(m in val_l for m in MESES):
                for mes in MESES:
                    if mes in val_l:
                        m = re.search(r"\d{4}", val_l)
                        year = m.group() if m else "?"
                        return f"{mes} {year}"
    return "desconocido"


def _find_comunas_row(ws) -> int:
    """Busca la fila donde aparecen las localidades."""
    for r in range(1, 25):
        for col in range(2, 15):
            val = ws.cell(r, col).value
            if val and isinstance(val, str) and (" - Aéreo" in val or " - Subterráneo" in val):
                return r
    return 5


def _extract_localidades(ws, comunas_row: int) -> dict:
    """Extrae {col_index: nombre_localidad} ignorando celdas 'Comunas'."""
    localidades = {}
    for col in range(4, ws.max_column + 1):
        val = ws.cell(comunas_row, col).value
        if val and isinstance(val, str):
            val_s = val.strip()
            if " - " in val_s and "Comunas" not in val_s:
                localidades[col] = val_s
    return localidades


def _find_data_start(ws, comunas_row: int) -> int:
    """Primera fila de datos reales (después de la fila de $ NETO/C/IVA)."""
    for r in range(comunas_row + 1, comunas_row + 6):
        for col in range(2, 8):
            val = ws.cell(r, col).value
            if val and isinstance(val, str) and "neto" in val.lower():
                return r + 1
    return comunas_row + 3


def _find_data_end(ws, data_start: int) -> int:
    """Última fila de datos de tarifas (antes de las notas al pie)."""
    for r in range(data_start + 5, ws.max_row):
        b = ws.cell(r, 2).value
        c = ws.cell(r, 3).value
        if b and isinstance(b, str) and c is None:
            b_l = b.lower()
            # Si no parece nombre de tarifa ni encabezado, es nota al pie
            if not any(x in b_l for x in ["tarifa", "bt", "at", "tr"]) or len(b) > 100:
                return r
    return ws.max_row - 30


def _find_injection_rows(ws) -> list:
    """Encuentra filas de inyecciones."""
    result = []
    for r in range(ws.max_row - 30, ws.max_row + 1):
        b = ws.cell(r, 2).value
        if b and isinstance(b, str) and "inyectada" in b.lower():
            result.append(r)
    return result


def _find_inj_comunas_row(ws, inj_rows: list) -> int | None:
    """Fila de localidades de la sección de inyecciones."""
    if not inj_rows:
        return None
    for r in range(max(1, inj_rows[0] - 5), inj_rows[0]):
        for col in range(2, 8):
            val = ws.cell(r, col).value
            if val and isinstance(val, str) and (" - Aéreo" in val or " - Subterráneo" in val):
                return r
    return None


@st.cache_data(show_spinner=False)
def procesar_excel(file_bytes: bytes) -> tuple:
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes))
    ws = wb.active

    periodo      = _find_titulo_y_periodo(ws)
    comunas_row  = _find_comunas_row(ws)
    data_start   = _find_data_start(ws, comunas_row)
    data_end     = _find_data_end(ws, data_start)
    localidades  = _extract_localidades(ws, comunas_row)
    inj_rows     = _find_injection_rows(ws)
    inj_com_row  = _find_inj_comunas_row(ws, inj_rows)

    # Parsear tarifas
    current_tariff = None
    rows = []

    for r in range(data_start, data_end):
        b = ws.cell(r, 2).value
        c = ws.cell(r, 3).value
        if not b:
            continue
        b_str = str(b).strip()
        if c is None:
            if b_str:
                current_tariff = b_str
            continue
        if current_tariff is None:
            continue
        concepto = b_str
        unidad   = str(c).strip()
        for loc_col, localidad in localidades.items():
            neto = ws.cell(r, loc_col).value
            civa = ws.cell(r, loc_col + 1).value
            parts = localidad.rsplit(" - ", 1)
            rows.append({
                "periodo":          periodo,
                "tarifa":           current_tariff,
                "concepto":         concepto,
                "unidad":           unidad,
                "localidad":        localidad,
                "comuna":           parts[0] if len(parts) == 2 else localidad,
                "tipo_suministro":  parts[1] if len(parts) == 2 else "",
                "valor_neto":       round(float(neto), 4) if isinstance(neto, (int, float)) else None,
                "valor_civa":       round(float(civa), 4) if isinstance(civa, (int, float)) else None,
            })

    df_tarifas = pd.DataFrame(rows)

    # Parsear inyecciones
    inj_loc = _extract_localidades(ws, inj_com_row) if inj_com_row else localidades
    inj_result = []
    for r in inj_rows:
        concepto = ws.cell(r, 2).value
        unidad   = ws.cell(r, 3).value
        if not concepto:
            continue
        for loc_col, localidad in inj_loc.items():
            neto = ws.cell(r, loc_col).value
            parts = localidad.rsplit(" - ", 1)
            inj_result.append({
                "periodo":         periodo,
                "concepto":        str(concepto).strip(),
                "unidad":          str(unidad).strip() if unidad else "",
                "localidad":       localidad,
                "comuna":          parts[0] if len(parts) == 2 else localidad,
                "tipo_suministro": parts[1] if len(parts) == 2 else "",
                "valor_neto":      round(float(neto), 4) if isinstance(neto, (int, float)) else None,
            })

    df_inyecciones = pd.DataFrame(inj_result)
    return df_tarifas, df_inyecciones, periodo


# ══════════════════════════════════════════════════════════════════════════════
# SIDEBAR
# ══════════════════════════════════════════════════════════════════════════════
with st.sidebar:
    st.markdown("## ⚡ SAESA Tarifas")
    st.markdown("---")
    st.markdown("### 📂 Cargar pliegos")
    uploaded_files = st.file_uploader(
        "Sube uno o más archivos Excel",
        type=["xlsx"],
        accept_multiple_files=True,
        help="Cada archivo corresponde a un período tarifario distinto.",
    )
    if uploaded_files:
        st.success(f"{len(uploaded_files)} archivo(s) cargado(s)")

st.markdown("# ⚡ Pliego Tarifario SAESA")

if not uploaded_files:
    st.info("👈 Sube uno o más archivos Excel del pliego tarifario en el panel lateral.")
    st.stop()

# ── Procesar archivos ──────────────────────────────────────────────────────────
all_tarifas, all_iny, errors = [], [], []

progress = st.progress(0, text="Procesando archivos...")
for i, f in enumerate(uploaded_files):
    try:
        df_t, df_i, periodo = procesar_excel(f.read())
        all_tarifas.append(df_t)
        all_iny.append(df_i)
        progress.progress((i + 1) / len(uploaded_files), text=f"✅ {periodo}")
    except Exception as e:
        errors.append(f"❌ {f.name}: {e}")
        progress.progress((i + 1) / len(uploaded_files))

progress.empty()
for err in errors:
    st.error(err)

if not all_tarifas:
    st.error("No se pudo procesar ningún archivo.")
    st.stop()

df_all     = pd.concat(all_tarifas, ignore_index=True)
df_iny_all = pd.concat(all_iny, ignore_index=True) if all_iny else pd.DataFrame()

periodos_ordenados = sorted(df_all["periodo"].unique(), key=periodo_sort_key)
df_all["periodo"]  = pd.Categorical(df_all["periodo"], categories=periodos_ordenados, ordered=True)
df_all             = df_all.sort_values("periodo")

localidades_list = sorted(df_all["localidad"].unique())
comunas_list     = sorted(df_all["comuna"].unique())
tarifas_list     = sorted(df_all["tarifa"].unique())

# ── Métricas ───────────────────────────────────────────────────────────────────
cols_m = st.columns(5)
for col, label, val in zip(
    cols_m,
    ["Períodos", "Localidades", "Tarifas", "Conceptos", "Registros"],
    [len(periodos_ordenados), df_all["localidad"].nunique(),
     df_all["tarifa"].nunique(), df_all["concepto"].nunique(), len(df_all)],
):
    col.markdown(f"""<div class="metric-card">
        <div class="label">{label}</div>
        <div class="value">{val:,}</div>
    </div>""", unsafe_allow_html=True)

badges = " ".join(f'<span class="period-badge">{p}</span>' for p in periodos_ordenados)
st.markdown(f"**Períodos cargados:** {badges}", unsafe_allow_html=True)
st.markdown("---")

# ══════════════════════════════════════════════════════════════════════════════
# TABS PRINCIPALES
# ══════════════════════════════════════════════════════════════════════════════
tabs = st.tabs([
    "🔍 Consulta por Localidad",
    "📊 Comparar Localidades",
    "📈 Evolución Temporal",
    "🗄️ Base de Datos",
    "🔋 Inyecciones",
])

# ── TAB 1: Consulta ────────────────────────────────────────────────────────────
with tabs[0]:
    st.subheader("Consulta de tarifas por localidad")
    c1, c2, c3, c4 = st.columns([2, 2, 1, 1])
    with c1: sel_loc = st.selectbox("Localidad", localidades_list, key="loc1")
    with c2: sel_tar = st.selectbox("Tarifa", ["Todas"] + tarifas_list, key="tar1")
    with c3: sel_per = st.selectbox("Período", ["Último"] + periodos_ordenados, key="per1")
    with c4: pt1     = st.radio("Precio", ["Neto", "C/IVA"], horizontal=True, key="pt1")

    pcol     = "valor_neto" if pt1 == "Neto" else "valor_civa"
    per_filt = periodos_ordenados[-1] if sel_per == "Último" else sel_per

    df_loc = df_all[(df_all["localidad"] == sel_loc) & (df_all["periodo"] == per_filt)].copy()
    if sel_tar != "Todas":
        df_loc = df_loc[df_loc["tarifa"] == sel_tar]

    if df_loc.empty:
        st.warning("No hay datos para la selección.")
    else:
        for tarifa_name, grp in df_loc.groupby("tarifa", sort=False):
            st.markdown(f'<div class="tariff-header">📋 {tarifa_name}</div>', unsafe_allow_html=True)
            show = grp[["concepto", "unidad", pcol]].copy()
            show.columns = ["Concepto", "Unidad", f"$ {pt1}"]
            st.dataframe(
                show.reset_index(drop=True).style.format(
                    {f"$ {pt1}": lambda x: f"{x:,.3f}" if pd.notna(x) else "—"}
                ),
                use_container_width=True, hide_index=True,
            )

# ── TAB 2: Comparar ────────────────────────────────────────────────────────────
with tabs[1]:
    st.subheader("Comparación entre localidades")
    c1, c2 = st.columns([3, 2])
    with c1:
        locs_sel = st.multiselect("Localidades (máx 6)", localidades_list,
                                   default=localidades_list[:3], max_selections=6, key="locs2")
    with c2:
        tar_comp = st.selectbox("Tarifa", tarifas_list, key="tar2")
        per_comp = st.selectbox("Período", ["Último"] + periodos_ordenados, key="per2")
        pc2      = st.radio("Precio", ["Neto", "C/IVA"], horizontal=True, key="pc2")

    pcol2    = "valor_neto" if pc2 == "Neto" else "valor_civa"
    per_c2   = periodos_ordenados[-1] if per_comp == "Último" else per_comp

    if not locs_sel:
        st.info("Selecciona al menos una localidad.")
    else:
        df_comp = df_all[
            (df_all["tarifa"] == tar_comp) &
            (df_all["localidad"].isin(locs_sel)) &
            (df_all["periodo"] == per_c2)
        ]
        if df_comp.empty:
            st.warning("Sin datos.")
        else:
            pivot = df_comp.pivot_table(
                index=["concepto", "unidad"], columns="localidad",
                values=pcol2, aggfunc="first",
            ).reset_index()
            pivot.columns.name = None
            nc = [c for c in pivot.columns if c not in ["concepto", "unidad"]]
            st.dataframe(
                pivot.style.format({c: lambda x: f"{x:,.3f}" if pd.notna(x) else "—" for c in nc}),
                use_container_width=True, hide_index=True,
            )

# ── TAB 3: Evolución ───────────────────────────────────────────────────────────
with tabs[2]:
    st.subheader("Evolución temporal de tarifas")

    if len(periodos_ordenados) < 2:
        st.info("Carga al menos 2 archivos de distintos períodos para ver la evolución.")
    else:
        ec1, ec2, ec3, ec4 = st.columns([2, 2, 2, 1])
        with ec1: evo_loc   = st.selectbox("Localidad", localidades_list, key="evo_loc")
        with ec2: evo_tar   = st.selectbox("Tarifa", tarifas_list, key="evo_tar")
        with ec3:
            conc_disp = sorted(df_all[df_all["tarifa"] == evo_tar]["concepto"].unique())
            evo_conc  = st.selectbox("Concepto", conc_disp, key="evo_conc")
        with ec4: evo_p     = st.radio("Precio", ["Neto", "C/IVA"], horizontal=True, key="evo_p")

        evo_col = "valor_neto" if evo_p == "Neto" else "valor_civa"

        df_evo = df_all[
            (df_all["localidad"] == evo_loc) &
            (df_all["tarifa"] == evo_tar) &
            (df_all["concepto"] == evo_conc)
        ][["periodo", evo_col]].dropna().sort_values("periodo").copy()

        st1, st2, st3, st4 = st.tabs(["📈 Gráfico", "📋 Tabla mes a mes", "📉 Variación %", "🚨 Alertas"])

        with st1:
            if df_evo.empty:
                st.warning("Sin datos para la combinación seleccionada.")
            else:
                unidad_str = df_all[
                    (df_all["tarifa"] == evo_tar) & (df_all["concepto"] == evo_conc)
                ]["unidad"].iloc[0]
                labels     = df_evo["periodo"].astype(str).tolist()
                values     = df_evo[evo_col].tolist()
                chart_data = json.dumps({"labels": labels, "values": values})
                title_str  = f"{evo_conc} — {evo_loc}"
                ylabel_str = f"$ {evo_p} ({unidad_str})"

                chart_html = f"""<!DOCTYPE html><html><head>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>
</head><body style="margin:0;padding:10px;background:#fff;">
<canvas id="c" height="110"></canvas>
<script>
const d={chart_data};
new Chart(document.getElementById('c').getContext('2d'),{{
  type:'line',
  data:{{labels:d.labels,datasets:[{{
    label:{json.dumps(ylabel_str)},data:d.values,
    borderColor:'#4a0e8f',backgroundColor:'rgba(74,14,143,0.08)',
    borderWidth:3,pointBackgroundColor:'#7c3aed',pointRadius:6,pointHoverRadius:9,
    tension:0.3,fill:true
  }}]}},
  options:{{responsive:true,plugins:{{
    title:{{display:true,text:{json.dumps(title_str)},font:{{size:14,weight:'bold'}},color:'#3b0764'}},
    tooltip:{{callbacks:{{label:c=>'$ '+c.parsed.y.toLocaleString('es-CL',{{minimumFractionDigits:3}})}}}}
  }},scales:{{y:{{
    title:{{display:true,text:{json.dumps(ylabel_str)}}},
    ticks:{{callback:v=>'$'+v.toLocaleString('es-CL',{{minimumFractionDigits:0}})}}
  }}}}}}
}});
</script></body></html>"""
                st.components.v1.html(chart_html, height=420)
                show_e = df_evo.copy()
                show_e.columns = ["Período", f"$ {evo_p}"]
                st.dataframe(
                    show_e.style.format({f"$ {evo_p}": "{:,.3f}"}),
                    use_container_width=True, hide_index=True,
                )

        with st2:
            st.markdown("**Todos los conceptos de la tarifa por período**")
            df_pt = df_all[
                (df_all["localidad"] == evo_loc) & (df_all["tarifa"] == evo_tar)
            ].pivot_table(
                index=["concepto", "unidad"], columns="periodo",
                values=evo_col, aggfunc="first",
            ).reset_index()
            df_pt.columns.name = None
            nc_t = [c for c in df_pt.columns if c not in ["concepto", "unidad"]]
            if df_pt.empty:
                st.warning("Sin datos.")
            else:
                st.dataframe(
                    df_pt.style.format(
                        {c: lambda x: f"{x:,.3f}" if pd.notna(x) else "—" for c in nc_t}
                    ).highlight_null(color="#f3f4f6"),
                    use_container_width=True, hide_index=True,
                )

        with st3:
            p1c, p2c = st.columns(2)
            with p1c: per_base = st.selectbox("Período base", periodos_ordenados[:-1], key="pbase")
            with p2c:
                ops = [p for p in periodos_ordenados if periodo_sort_key(p) > periodo_sort_key(per_base)]
                per_cmp = st.selectbox("Período a comparar", ops, key="pcmp")

            def get_vals(per):
                return df_all[
                    (df_all["localidad"] == evo_loc) &
                    (df_all["tarifa"] == evo_tar) &
                    (df_all["periodo"] == per)
                ][["concepto", "unidad", evo_col]]

            df_var = get_vals(per_base).rename(columns={evo_col: "base"}).merge(
                get_vals(per_cmp).rename(columns={evo_col: "nuevo"}),
                on=["concepto", "unidad"], how="outer",
            )
            df_var["var_%"] = ((df_var["nuevo"] - df_var["base"]) / df_var["base"] * 100).round(2)
            df_var.columns  = ["Concepto", "Unidad", f"$ {per_base}", f"$ {per_cmp}", "Variación %"]

            def color_var(val):
                if pd.isna(val): return ""
                if val > 0: return "color:#dc2626;font-weight:600"
                if val < 0: return "color:#16a34a;font-weight:600"
                return "color:#6b7280"

            st.dataframe(
                df_var.style
                    .format({
                        f"$ {per_base}": lambda x: f"{x:,.3f}" if pd.notna(x) else "—",
                        f"$ {per_cmp}":  lambda x: f"{x:,.3f}" if pd.notna(x) else "—",
                        "Variación %":   lambda x: f"{x:+.2f}%" if pd.notna(x) else "—",
                    })
                    .applymap(color_var, subset=["Variación %"]),
                use_container_width=True, hide_index=True,
            )

        with st4:
            st.markdown(f"**Cambios entre {periodos_ordenados[-2]} → {periodos_ordenados[-1]}**")
            umbral  = st.slider("Umbral mínimo de variación (%)", 0.0, 20.0, 1.0, 0.5, key="umbral")
            per_ant = periodos_ordenados[-2]
            per_ult = periodos_ordenados[-1]

            def get_all(per):
                return df_all[df_all["periodo"] == per][
                    ["localidad", "tarifa", "concepto", "unidad", evo_col]
                ]

            df_alert = get_all(per_ant).rename(columns={evo_col: "ant"}).merge(
                get_all(per_ult).rename(columns={evo_col: "ult"}),
                on=["localidad", "tarifa", "concepto", "unidad"], how="outer",
            )
            df_alert["var_%"] = ((df_alert["ult"] - df_alert["ant"]) / df_alert["ant"] * 100).round(2)
            df_alert = df_alert[df_alert["var_%"].abs() >= umbral].dropna(subset=["var_%"])
            df_alert = df_alert.sort_values("var_%", ascending=False)

            al1, al2 = st.columns(2)
            with al1: a_loc = st.selectbox("Filtrar localidad", ["Todas"] + comunas_list, key="al_loc")
            with al2: a_tar = st.selectbox("Filtrar tarifa",    ["Todas"] + tarifas_list, key="al_tar")
            if a_loc != "Todas": df_alert = df_alert[df_alert["localidad"].str.startswith(a_loc)]
            if a_tar != "Todas": df_alert = df_alert[df_alert["tarifa"] == a_tar]

            st.caption(f"{len(df_alert)} conceptos con variación ≥ {umbral}%")

            if df_alert.empty:
                st.success("✅ No se detectaron variaciones significativas.")
            else:
                df_alert = df_alert.rename(columns={
                    "localidad": "Localidad", "tarifa": "Tarifa",
                    "concepto": "Concepto",   "unidad": "Unidad",
                    "ant": f"$ {per_ant}",    "ult": f"$ {per_ult}",
                    "var_%": "Variación %",
                })

                def color_alert(val):
                    if pd.isna(val): return ""
                    if val > 0: return "background:#fee2e2;color:#dc2626;font-weight:600"
                    if val < 0: return "background:#dcfce7;color:#16a34a;font-weight:600"
                    return ""

                st.dataframe(
                    df_alert.reset_index(drop=True).style
                        .format({
                            f"$ {per_ant}": lambda x: f"{x:,.3f}" if pd.notna(x) else "—",
                            f"$ {per_ult}": lambda x: f"{x:,.3f}" if pd.notna(x) else "—",
                            "Variación %":  lambda x: f"{x:+.2f}%" if pd.notna(x) else "—",
                        })
                        .applymap(color_alert, subset=["Variación %"]),
                    use_container_width=True, hide_index=True, height=500,
                )
                st.download_button(
                    "⬇️ Descargar alertas CSV",
                    df_alert.to_csv(index=False).encode("utf-8-sig"),
                    "alertas_tarifarias.csv", "text/csv",
                )

# ── TAB 4: Base de datos ───────────────────────────────────────────────────────
with tabs[3]:
    st.subheader("Base de datos normalizada completa")
    c1, c2, c3, c4 = st.columns(4)
    with c1: f_per = st.multiselect("Período", periodos_ordenados, key="f_per")
    with c2: f_com = st.multiselect("Comuna",  comunas_list,       key="f_com")
    with c3: f_tar = st.multiselect("Tarifa",  tarifas_list,       key="f_tar")
    with c4: f_txt = st.text_input("Buscar concepto", placeholder="ej: energía...", key="f_txt")

    df_full = df_all.copy()
    if f_per: df_full = df_full[df_full["periodo"].isin(f_per)]
    if f_com: df_full = df_full[df_full["comuna"].isin(f_com)]
    if f_tar: df_full = df_full[df_full["tarifa"].isin(f_tar)]
    if f_txt: df_full = df_full[df_full["concepto"].str.contains(f_txt, case=False, na=False)]

    st.caption(f"Mostrando {len(df_full):,} registros")
    st.dataframe(
        df_full[["periodo", "tarifa", "concepto", "unidad",
                 "comuna", "tipo_suministro", "valor_neto", "valor_civa"]
                ].reset_index(drop=True).style.format({
            "valor_neto": lambda x: f"{x:,.3f}" if pd.notna(x) else "—",
            "valor_civa": lambda x: f"{x:,.3f}" if pd.notna(x) else "—",
        }),
        use_container_width=True, hide_index=True, height=500,
    )
    st.download_button(
        "⬇️ Descargar CSV filtrado",
        df_full.to_csv(index=False).encode("utf-8-sig"),
        "tarifas_saesa_filtrado.csv", "text/csv",
    )

# ── TAB 5: Inyecciones ─────────────────────────────────────────────────────────
with tabs[4]:
    st.subheader("Precios para valorización de inyecciones de energía")
    st.caption("Valores netos (sin IVA) — Art. 149 bis DFL N°4/2006")

    if df_iny_all.empty:
        st.warning("No se encontraron datos de inyecciones.")
    else:
        c1, c2 = st.columns(2)
        with c1: per_iny   = st.selectbox("Período", ["Último"] + periodos_ordenados, key="per_iny")
        with c2: f_loc_iny = st.multiselect("Filtrar localidad",
                                             sorted(df_iny_all["localidad"].unique()), key="f_iny")

        per_i   = periodos_ordenados[-1] if per_iny == "Último" else per_iny
        df_iny_f = df_iny_all[df_iny_all["periodo"] == per_i]
        if f_loc_iny:
            df_iny_f = df_iny_f[df_iny_f["localidad"].isin(f_loc_iny)]

        if df_iny_f.empty:
            st.warning("Sin datos.")
        else:
            pivot_iny = df_iny_f.pivot_table(
                index=["concepto", "unidad"], columns="localidad",
                values="valor_neto", aggfunc="first",
            ).reset_index()
            pivot_iny.columns.name = None
            nc_i = [c for c in pivot_iny.columns if c not in ["concepto", "unidad"]]
            st.dataframe(
                pivot_iny.style.format(
                    {c: lambda x: f"{x:,.3f}" if pd.notna(x) else "—" for c in nc_i}
                ),
                use_container_width=True, hide_index=True,
            )

        if len(periodos_ordenados) >= 2:
            st.markdown("---")
            st.markdown("**Evolución por localidad**")
            loc_iny_evo = st.selectbox(
                "Localidad", sorted(df_iny_all["localidad"].unique()), key="loc_iny_evo"
            )
            df_ie = df_iny_all[df_iny_all["localidad"] == loc_iny_evo].pivot_table(
                index=["concepto", "unidad"], columns="periodo",
                values="valor_neto", aggfunc="first",
            ).reset_index()
            df_ie.columns.name = None
            nc_ie = [c for c in df_ie.columns if c not in ["concepto", "unidad"]]
            st.dataframe(
                df_ie.style.format(
                    {c: lambda x: f"{x:,.3f}" if pd.notna(x) else "—" for c in nc_ie}
                ),
                use_container_width=True, hide_index=True,
            )
