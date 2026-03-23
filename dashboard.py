import streamlit as st
import gspread
import pandas as pd
import plotly.express as px
from collections import Counter

# ── Config ───────────────────────────────────────────────────
st.set_page_config(
    page_title="NPS Dashboard — Deuna Partnerships",
    page_icon="📊",
    layout="wide",
)

SPREADSHEET_ID = "1AGdEueo334gtGc82oq2kffjPdLrcJhFFd9uxcIDKQWk"
SHEET_NAME     = "NPS Answers"
CREDS_FILE     = "/Users/bsampaio/gcp-oauth.keys.json"
LOGO_PATH      = "/Users/bsampaio/Documents/Claude - Deuna/nps-parceiros/nps/deuna_logo.png"

ORANGE = "#FF5500"
TEAL   = "#0B9595"
RED    = "#FF614B"
BLUE   = "#3B5BDB"
AMBER  = "#B45309"
DARK   = "#1C1C1C"

COLS_HEADER = [
    'Data/Hora', 'Empresa', 'Nome', 'E-mail',
    'NPS (0–10)', 'Categoria NPS', 'Idioma', 'Resposta Condicional',
    'Suporte Técnico',
    'Com. Técnica — Velocidade', 'Com. Técnica — Qualidade', 'Com. Técnica — Proatividade',
    'Comunicação da Equipe', 'Resultados (1–10)',
    'Integração — Rapidez', 'Integração — Qualidade', 'Integração — Facilidade',
    'Aspectos Valorizados', 'Sugestões de Melhoria',
    'Clareza Valor Agregado (1–10)',
]

# ── Translations ─────────────────────────────────────────────
T = {
    "pt": {
        "title":          "NPS Dashboard — Deuna Partnerships",
        "filters":        "Filtros",
        "language":       "Idioma da interface",
        "filter_lang":    "Filtrar por idioma da resposta",
        "filter_company": "Empresa",
        "all_langs":      "Todos",
        "all_companies":  "Todas",
        "refresh":        "🔄 Atualizar dados",
        "loading":        "Carregando dados...",
        "no_data":        "Nenhuma resposta ainda.",
        "error":          "Erro ao carregar planilha",
        "tab_dash":       "📊  Dashboard",
        "tab_detail":     "📋  Respostas",
        # Dashboard
        "overview":       "Visão Geral",
        "total":          "Total de Respostas",
        "nps_score":      "NPS Score",
        "promoters":      "😊 Promotores (9–10)",
        "detractors":     "😟 Detratores (0–6)",
        "dist_title":     "Distribuição de Notas NPS",
        "dist_y":         "Respostas",
        "pie_title":      "Promotores · Neutros · Detratores",
        "pie_promoters":  "Promotores (9–10)",
        "pie_neutrals":   "Neutros (7–8)",
        "pie_detractors": "Detratores (0–6)",
        "sliders_title":  "Médias — Réguas (escala 1–10)",
        "slider_results": "Resultados da Parceria",
        "slider_valor":   "Clareza Valor Agregado",
        "slider_rap":     "Integração — Rapidez",
        "slider_qual":    "Integração — Qualidade",
        "slider_facil":   "Integração — Facilidade",
        "slider_chart":   "Comparativo de Médias por Régua",
        "slider_y":       "Média (1–10)",
        "comtec_title":   "Comunicação Técnica — Score Médio (😊=10 · 😐=5 · 😞=1)",
        "vel":            "⚡ Velocidade",
        "qual":           "✅ Qualidade",
        "proat":          "📣 Proatividade",
        "face_chart":     "Distribuição das Reações — Comunicação Técnica",
        "face_good":      "😊 Bom",
        "face_neutral":   "😐 Neutro",
        "face_bad":       "😞 Ruim",
        "face_dim":       "Dimensão",
        "face_reaction":  "Reação",
        "likert_title":   "Suporte e Comunicação da Equipe",
        "likert_chart":   "Avaliação de Suporte e Comunicação",
        "likert_suporte": "Suporte Técnico",
        "likert_com":     "Comunicação da Equipe",
        "likert_ct":      "Concordo totalmente",
        "likert_c":       "Concordo",
        "likert_d":       "Discordo",
        "likert_dt":      "Discordo totalmente",
        "aspects_title":  "O que os Parceiros Mais Valorizam",
        "aspects_chart":  "Aspectos mais mencionados",
        "aspects_col":    "Aspecto",
        "mentions_col":   "Menções",
        "footer":         "Dados atualizados a cada 5 min",
        "responses":      "resposta",
        "responses_pl":   "respostas",
        # Detail tab
        "individual":     "Respostas Individuais",
        "comment":        "Comentário",
        "suggestion":     "Sugestão",
        "values":         "Valoriza",
        "scores":         "Notas por Resposta",
        "export":         "Exportar",
        "download":       "⬇️  Baixar todas as respostas (.csv)",
        "filename":       "nps_deuna_respostas.csv",
        "criterion":      "Critério",
    },
    "es": {
        "title":          "NPS Dashboard — Deuna Partnerships",
        "filters":        "Filtros",
        "language":       "Idioma de la interfaz",
        "filter_lang":    "Filtrar por idioma de respuesta",
        "filter_company": "Empresa",
        "all_langs":      "Todos",
        "all_companies":  "Todas",
        "refresh":        "🔄 Actualizar datos",
        "loading":        "Cargando datos...",
        "no_data":        "Aún no hay respuestas.",
        "error":          "Error al cargar la hoja",
        "tab_dash":       "📊  Dashboard",
        "tab_detail":     "📋  Respuestas",
        "overview":       "Resumen General",
        "total":          "Total de Respuestas",
        "nps_score":      "NPS Score",
        "promoters":      "😊 Promotores (9–10)",
        "detractors":     "😟 Detractores (0–6)",
        "dist_title":     "Distribución de Notas NPS",
        "dist_y":         "Respuestas",
        "pie_title":      "Promotores · Neutros · Detractores",
        "pie_promoters":  "Promotores (9–10)",
        "pie_neutrals":   "Neutros (7–8)",
        "pie_detractors": "Detractores (0–6)",
        "sliders_title":  "Promedios — Escalas (1–10)",
        "slider_results": "Resultados de la Alianza",
        "slider_valor":   "Claridad Valor Agregado",
        "slider_rap":     "Integración — Rapidez",
        "slider_qual":    "Integración — Calidad",
        "slider_facil":   "Integración — Facilidad",
        "slider_chart":   "Comparativo de Promedios por Escala",
        "slider_y":       "Promedio (1–10)",
        "comtec_title":   "Comunicación Técnica — Score Promedio (😊=10 · 😐=5 · 😞=1)",
        "vel":            "⚡ Velocidad",
        "qual":           "✅ Calidad",
        "proat":          "📣 Proactividad",
        "face_chart":     "Distribución de Reacciones — Comunicación Técnica",
        "face_good":      "😊 Bueno",
        "face_neutral":   "😐 Neutro",
        "face_bad":       "😞 Malo",
        "face_dim":       "Dimensión",
        "face_reaction":  "Reacción",
        "likert_title":   "Soporte y Comunicación del Equipo",
        "likert_chart":   "Evaluación de Soporte y Comunicación",
        "likert_suporte": "Soporte Técnico",
        "likert_com":     "Comunicación del Equipo",
        "likert_ct":      "Totalmente de acuerdo",
        "likert_c":       "De acuerdo",
        "likert_d":       "En desacuerdo",
        "likert_dt":      "Totalmente en desacuerdo",
        "aspects_title":  "Lo que Más Valoran los Partners",
        "aspects_chart":  "Aspectos más mencionados",
        "aspects_col":    "Aspecto",
        "mentions_col":   "Menciones",
        "footer":         "Datos actualizados cada 5 min",
        "responses":      "respuesta",
        "responses_pl":   "respuestas",
        "individual":     "Respuestas Individuales",
        "comment":        "Comentario",
        "suggestion":     "Sugerencia",
        "values":         "Valora",
        "scores":         "Notas por Respuesta",
        "export":         "Exportar",
        "download":       "⬇️  Descargar todas las respuestas (.csv)",
        "filename":       "nps_deuna_respuestas.csv",
        "criterion":      "Criterio",
    },
    "en": {
        "title":          "NPS Dashboard — Deuna Partnerships",
        "filters":        "Filters",
        "language":       "Interface language",
        "filter_lang":    "Filter by response language",
        "filter_company": "Company",
        "all_langs":      "All",
        "all_companies":  "All",
        "refresh":        "🔄 Refresh data",
        "loading":        "Loading data...",
        "no_data":        "No responses yet.",
        "error":          "Error loading spreadsheet",
        "tab_dash":       "📊  Dashboard",
        "tab_detail":     "📋  Responses",
        "overview":       "Overview",
        "total":          "Total Responses",
        "nps_score":      "NPS Score",
        "promoters":      "😊 Promoters (9–10)",
        "detractors":     "😟 Detractors (0–6)",
        "dist_title":     "NPS Score Distribution",
        "dist_y":         "Responses",
        "pie_title":      "Promoters · Neutrals · Detractors",
        "pie_promoters":  "Promoters (9–10)",
        "pie_neutrals":   "Neutrals (7–8)",
        "pie_detractors": "Detractors (0–6)",
        "sliders_title":  "Averages — Scales (1–10)",
        "slider_results": "Partnership Results",
        "slider_valor":   "Value Clarity",
        "slider_rap":     "Integration — Speed",
        "slider_qual":    "Integration — Quality",
        "slider_facil":   "Integration — Ease",
        "slider_chart":   "Scale Averages Comparison",
        "slider_y":       "Average (1–10)",
        "comtec_title":   "Technical Communication — Avg Score (😊=10 · 😐=5 · 😞=1)",
        "vel":            "⚡ Speed",
        "qual":           "✅ Quality",
        "proat":          "📣 Proactivity",
        "face_chart":     "Reaction Distribution — Technical Communication",
        "face_good":      "😊 Good",
        "face_neutral":   "😐 Neutral",
        "face_bad":       "😞 Bad",
        "face_dim":       "Dimension",
        "face_reaction":  "Reaction",
        "likert_title":   "Support & Team Communication",
        "likert_chart":   "Support & Communication Assessment",
        "likert_suporte": "Technical Support",
        "likert_com":     "Team Communication",
        "likert_ct":      "Strongly agree",
        "likert_c":       "Agree",
        "likert_d":       "Disagree",
        "likert_dt":      "Strongly disagree",
        "aspects_title":  "What Partners Value Most",
        "aspects_chart":  "Most mentioned aspects",
        "aspects_col":    "Aspect",
        "mentions_col":   "Mentions",
        "footer":         "Data refreshed every 5 min",
        "responses":      "response",
        "responses_pl":   "responses",
        "individual":     "Individual Responses",
        "comment":        "Comment",
        "suggestion":     "Suggestion",
        "values":         "Values",
        "scores":         "Scores per Response",
        "export":         "Export",
        "download":       "⬇️  Download all responses (.csv)",
        "filename":       "nps_deuna_responses.csv",
        "criterion":      "Criterion",
    },
}

LIKERT_NORMALIZE = {
    "Concordo totalmente": "ct", "Strongly agree": "ct", "Totalmente de acuerdo": "ct",
    "Concordo": "c", "Agree": "c", "De acuerdo": "c",
    "Discordo": "d", "Disagree": "d", "En desacuerdo": "d",
    "Discordo totalmente": "dt", "Strongly disagree": "dt", "Totalmente en desacuerdo": "dt",
}
FACE_SCORE = {"good": 10, "neutral": 5, "bad": 1}

# ── CSS ───────────────────────────────────────────────────────
st.markdown("""
<style>
  [data-testid="stAppViewContainer"] { background: #FAFAFA; }
  [data-testid="stSidebar"] { background: #F5F5F5; }
  [data-testid="stSidebar"] img { mix-blend-mode: multiply; }
  [data-testid="stAppViewContainer"] > section:first-child img { mix-blend-mode: multiply; }
  [data-testid="stSidebar"] button { background: #FF5500 !important; color: white !important; border: none !important; border-radius: 8px !important; }
  .stTabs [data-baseweb="tab-list"] { border-bottom: 2px solid #FF5500; }
  .stTabs [data-baseweb="tab"] { font-weight: 600; color: #888; }
  .stTabs [aria-selected="true"] { color: #FF5500 !important; border-bottom: 2px solid #FF5500 !important; }
  .metric-card {
    background: white; border-radius: 12px; padding: 20px 16px;
    text-align: center; border: 1px solid #EBEBEB;
    box-shadow: 0 1px 4px rgba(0,0,0,.06);
  }
  .metric-label { font-size: 12px; font-weight: 600; color: #888;
                  text-transform: uppercase; letter-spacing: .5px; }
  .metric-value { font-size: 40px; font-weight: 800; margin: 4px 0 0; }
  .section-title { font-size: 13px; font-weight: 700; color: #888;
                   text-transform: uppercase; letter-spacing: .6px; margin: 32px 0 8px; }
  hr.divider { border: none; border-top: 1px solid #EBEBEB; margin: 4px 0 20px; }
  .response-card {
    background: white; border-radius: 12px; padding: 20px 24px;
    border: 1px solid #EBEBEB; margin-bottom: 16px;
    box-shadow: 0 1px 4px rgba(0,0,0,.04);
  }
  .badge { display: inline-block; padding: 3px 10px; border-radius: 20px;
           font-size: 12px; font-weight: 700; margin-right: 6px; }
  .badge-promotor  { background: #E6F7F7; color: #0B9595; }
  .badge-neutro    { background: #FFF8EC; color: #B45309; }
  .badge-detrator  { background: #FEF2F2; color: #FF614B; }
</style>
""", unsafe_allow_html=True)

# ── Load data ─────────────────────────────────────────────────
@st.cache_data(ttl=300)
def load_data():
    import json
    from google.oauth2.credentials import Credentials
    from google.auth.transport.requests import Request

    with open(CREDS_FILE) as f:
        creds_data = json.load(f)["installed"]
    with open("/Users/bsampaio/.mcp-google-sheets-token.json") as f:
        token_data = json.load(f)

    creds = Credentials(
        token=token_data.get("access_token"),
        refresh_token=token_data.get("refresh_token"),
        token_uri="https://oauth2.googleapis.com/token",
        client_id=creds_data["client_id"],
        client_secret=creds_data["client_secret"],
        scopes=["https://www.googleapis.com/auth/spreadsheets",
                "https://www.googleapis.com/auth/drive.file"],
    )
    if not creds.valid:
        creds.refresh(Request())

    gc = gspread.authorize(creds)
    rows = gc.open_by_key(SPREADSHEET_ID).worksheet(SHEET_NAME).get_all_values()
    if not rows:
        return pd.DataFrame()

    has_header = any(h in str(rows[0][0]) for h in ['Data', 'data', 'Empresa'])
    data_rows  = rows[1:] if has_header else rows
    if not data_rows:
        return pd.DataFrame()

    ncols     = max(len(r) for r in data_rows)
    headers   = COLS_HEADER[:ncols] + [f'col_{i}' for i in range(len(COLS_HEADER), ncols)]
    data_rows = [r + [''] * (ncols - len(r)) for r in data_rows]
    return pd.DataFrame(data_rows, columns=headers)

# ── Helpers ───────────────────────────────────────────────────
def card(label, value, color=DARK):
    return f"""<div class="metric-card">
      <div class="metric-label">{label}</div>
      <div class="metric-value" style="color:{color}">{value}</div>
    </div>"""

def section(title):
    st.markdown(f'<div class="section-title">{title}</div><hr class="divider">', unsafe_allow_html=True)

def avg_num(series):
    nums = pd.to_numeric(series, errors='coerce').dropna()
    return round(nums.mean(), 1) if len(nums) else 0

# ── Sidebar ───────────────────────────────────────────────────
# ── Password gate ────────────────────────────────────────────
if "authenticated" not in st.session_state:
    st.session_state.authenticated = False

if not st.session_state.authenticated:
    col_c = st.columns([1, 2, 1])[1]
    with col_c:
        st.markdown("<br><br>", unsafe_allow_html=True)
        st.image(LOGO_PATH, width=140)
        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown(f"<h2 style='color:{DARK};font-weight:800;margin-bottom:4px'>NPS Dashboard</h2>", unsafe_allow_html=True)
        st.markdown("<p style='color:#888;margin-bottom:24px'>Deuna Partnerships</p>", unsafe_allow_html=True)
        pwd = st.text_input("Password", type="password", placeholder="Enter password")
        if st.button("Enter", use_container_width=True):
            if pwd == "deuna2025":
                st.session_state.authenticated = True
                st.rerun()
            else:
                st.error("Incorrect password.")
    st.stop()

with st.sidebar:
    st.image(LOGO_PATH, width=120)
    st.markdown("<br>", unsafe_allow_html=True)
    lang = st.selectbox("🌐 Language / Idioma", ["pt", "es", "en"],
                        format_func=lambda x: {"pt": "🇧🇷 Português", "es": "🇪🇸 Español", "en": "🇺🇸 English"}[x])

t = T[lang]

col_logo, col_title = st.columns([1, 5])
with col_logo:
    st.image(LOGO_PATH, width=80)
with col_title:
    st.markdown(f"<h1 style='margin:16px 0 4px;color:{DARK};font-size:24px;font-weight:800'>{t['title']}</h1>", unsafe_allow_html=True)
st.markdown("<hr style='border:none;border-top:2px solid #FF5500;margin:8px 0 24px'>", unsafe_allow_html=True)

with st.spinner(t["loading"]):
    try:
        df = load_data()
    except Exception as e:
        st.error(f"{t['error']}: {e}")
        st.stop()

if df.empty:
    st.info(t["no_data"])
    st.stop()

with st.sidebar:
    st.markdown(f"<h3 style='color:{DARK}'>{t['filters']}</h3>", unsafe_allow_html=True)
    lang_label = {"pt": "🇧🇷 Português", "es": "🇪🇸 Español", "en": "🇺🇸 English"}
    raw_langs  = sorted(df["Idioma"].dropna().unique().tolist())
    lang_opts  = [t["all_langs"]] + [lang_label.get(l, l) for l in raw_langs]
    lang_sel   = st.selectbox(t["filter_lang"], lang_opts)
    idioma_sel = {"🇧🇷 Português": "pt", "🇪🇸 Español": "es", "🇺🇸 English": "en"}.get(lang_sel, lang_sel)
    empresa_sel = st.selectbox(t["filter_company"], [t["all_companies"]] + sorted(df["Empresa"].dropna().unique().tolist()))
    st.markdown("---")
    if st.button(t["refresh"]):
        st.cache_data.clear()
        st.rerun()

dff = df.copy()
if idioma_sel  != t["all_langs"]:     dff = dff[dff["Idioma"]  == idioma_sel]
if empresa_sel != t["all_companies"]: dff = dff[dff["Empresa"] == empresa_sel]

nps_col = "NPS (0–10)"
dff[nps_col] = pd.to_numeric(dff[nps_col], errors="coerce")
dff = dff.dropna(subset=[nps_col])

total      = len(dff)
promotores = int((dff[nps_col] >= 9).sum())
neutros    = int(((dff[nps_col] >= 7) & (dff[nps_col] <= 8)).sum())
detratores = int((dff[nps_col] <= 6).sum())
nps_score  = round(((promotores - detratores) / total) * 100) if total else 0
nps_color  = TEAL if nps_score >= 50 else (AMBER if nps_score >= 0 else RED)

# ── Tabs ──────────────────────────────────────────────────────
tab_dash, tab_details = st.tabs([t["tab_dash"], t["tab_detail"]])

# ════════════════════════════════════════════════════════════
# ABA 1 — DASHBOARD
# ════════════════════════════════════════════════════════════
with tab_dash:

    section(t["overview"])
    c1, c2, c3, c4 = st.columns(4)
    c1.markdown(card(t["total"],      total,      ORANGE), unsafe_allow_html=True)
    c2.markdown(card(t["nps_score"],  f"{'+' if nps_score > 0 else ''}{nps_score}", nps_color), unsafe_allow_html=True)
    c3.markdown(card(t["promoters"],  promotores, TEAL),   unsafe_allow_html=True)
    c4.markdown(card(t["detractors"], detratores, RED),    unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    ch1, ch2 = st.columns(2)
    with ch1:
        dist = dff[nps_col].value_counts().reindex(range(11), fill_value=0).reset_index()
        dist.columns = ["Nota", t["dist_y"]]
        colors = [RED if n <= 6 else (AMBER if n <= 8 else TEAL) for n in dist["Nota"]]
        fig = px.bar(dist, x="Nota", y=t["dist_y"], title=t["dist_title"],
                     color="Nota", color_discrete_sequence=colors)
        fig.update_layout(showlegend=False, plot_bgcolor="white", paper_bgcolor="white",
                          title_font_size=14, font_family="Arial")
        fig.update_xaxes(tickmode="linear", dtick=1)
        st.plotly_chart(fig, use_container_width=True)

    with ch2:
        seg = pd.DataFrame({
            "Categoria": [t["pie_promoters"], t["pie_neutrals"], t["pie_detractors"]],
            "Total": [promotores, neutros, detratores],
        })
        fig2 = px.pie(seg, names="Categoria", values="Total", title=t["pie_title"],
                      color="Categoria",
                      color_discrete_map={t["pie_promoters"]: TEAL, t["pie_neutrals"]: AMBER, t["pie_detractors"]: RED},
                      hole=0.5)
        fig2.update_layout(plot_bgcolor="white", paper_bgcolor="white",
                           title_font_size=14, font_family="Arial")
        st.plotly_chart(fig2, use_container_width=True)

    # Réguas
    section(t["sliders_title"])
    slider_cols = {
        t["slider_results"]: "Resultados (1–10)",
        t["slider_valor"]:   "Clareza Valor Agregado (1–10)",
        t["slider_rap"]:     "Integração — Rapidez",
        t["slider_qual"]:    "Integração — Qualidade",
        t["slider_facil"]:   "Integração — Facilidade",
    }
    avgs = {lbl: avg_num(dff[col]) for lbl, col in slider_cols.items() if col in dff.columns}
    cols_w = st.columns(len(avgs))
    for i, (lbl, val) in enumerate(avgs.items()):
        cols_w[i].markdown(card(lbl, val, BLUE), unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)
    fig_sl = px.bar(x=list(avgs.keys()), y=list(avgs.values()),
                    labels={"x": "", "y": t["slider_y"]},
                    title=t["slider_chart"],
                    color_discrete_sequence=[BLUE])
    fig_sl.update_layout(showlegend=False, plot_bgcolor="white", paper_bgcolor="white",
                         yaxis_range=[0, 10], title_font_size=14, font_family="Arial")
    st.plotly_chart(fig_sl, use_container_width=True)

    # Comunicação Técnica
    section(t["comtec_title"])
    face_cols = {
        t["vel"]:  "Com. Técnica — Velocidade",
        t["qual"]: "Com. Técnica — Qualidade",
        t["proat"]:"Com. Técnica — Proatividade",
    }
    face_avgs = {}
    for lbl, col in face_cols.items():
        if col in dff.columns:
            nums = dff[col].map(FACE_SCORE).dropna()
            face_avgs[lbl] = round(nums.mean(), 1) if len(nums) else 0

    f1, f2, f3 = st.columns(3)
    for w, (lbl, val) in zip([f1, f2, f3], face_avgs.items()):
        w.markdown(card(lbl, val, AMBER), unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)
    face_map_local = {"good": t["face_good"], "neutral": t["face_neutral"], "bad": t["face_bad"]}
    face_rows = []
    for lbl, col in face_cols.items():
        if col in dff.columns:
            for resp, cnt in dff[col].map(face_map_local).dropna().value_counts().items():
                face_rows.append({t["face_dim"]: lbl.split(" ", 1)[1], t["face_reaction"]: resp, "Total": cnt})
    if face_rows:
        df_face = pd.DataFrame(face_rows)
        fig_face = px.bar(df_face, x="Total", y=t["face_dim"], color=t["face_reaction"],
                          orientation="h", barmode="stack", title=t["face_chart"],
                          color_discrete_map={t["face_good"]: TEAL, t["face_neutral"]: AMBER, t["face_bad"]: RED})
        fig_face.update_layout(plot_bgcolor="white", paper_bgcolor="white",
                               title_font_size=14, font_family="Arial")
        st.plotly_chart(fig_face, use_container_width=True)

    # Likert
    section(t["likert_title"])
    likert_labels = {
        "ct": t["likert_ct"], "c": t["likert_c"], "d": t["likert_d"], "dt": t["likert_dt"]
    }
    likert_metrics = {t["likert_suporte"]: "Suporte Técnico", t["likert_com"]: "Comunicação da Equipe"}
    rows_lik = []
    for lbl, col in likert_metrics.items():
        if col in dff.columns:
            for raw, cnt in dff[col].value_counts().items():
                key = LIKERT_NORMALIZE.get(raw)
                if key:
                    rows_lik.append({t["criterion"]: lbl, "Resposta": likert_labels[key], "Total": cnt})
    if rows_lik:
        df_lik = pd.DataFrame(rows_lik)
        fig_lik = px.bar(df_lik, x="Total", y=t["criterion"], color="Resposta",
                         orientation="h", barmode="stack", title=t["likert_chart"],
                         color_discrete_map={
                             t["likert_ct"]: TEAL, t["likert_c"]: "#76B4E8",
                             t["likert_d"]: AMBER,  t["likert_dt"]: RED,
                         })
        fig_lik.update_layout(plot_bgcolor="white", paper_bgcolor="white",
                              title_font_size=14, font_family="Arial")
        st.plotly_chart(fig_lik, use_container_width=True)

    # Aspectos
    section(t["aspects_title"])
    if "Aspectos Valorizados" in dff.columns:
        all_asp = []
        for row in dff["Aspectos Valorizados"].dropna():
            all_asp.extend([a.strip() for a in str(row).split(",") if a.strip()])
        if all_asp:
            df_asp = pd.DataFrame(Counter(all_asp).most_common(10), columns=[t["aspects_col"], t["mentions_col"]])
            fig_asp = px.bar(df_asp, x=t["mentions_col"], y=t["aspects_col"], orientation="h",
                             title=t["aspects_chart"], color_discrete_sequence=[ORANGE])
            fig_asp.update_layout(yaxis={"categoryorder": "total ascending"},
                                  showlegend=False, plot_bgcolor="white", paper_bgcolor="white",
                                  title_font_size=14, font_family="Arial")
            st.plotly_chart(fig_asp, use_container_width=True)

    n_word = t["responses"] if total == 1 else t["responses_pl"]
    st.markdown(f"<p style='color:#CCC;font-size:11px;text-align:center'>{t['footer']} · {total} {n_word}</p>",
                unsafe_allow_html=True)


# ════════════════════════════════════════════════════════════
# ABA 2 — RESPOSTAS
# ════════════════════════════════════════════════════════════
with tab_details:

    section(t["individual"])

    for _, row in dff.sort_values("Data/Hora", ascending=False).iterrows():
        nps_val  = int(row.get("NPS (0–10)", 0))
        cat      = row.get("Categoria NPS", "")
        empresa  = row.get("Empresa", "—")
        nome     = row.get("Nome", "")
        email    = row.get("E-mail", "")
        data     = row.get("Data/Hora", "")
        cond     = row.get("Resposta Condicional", "")
        sugestao = row.get("Sugestões de Melhoria", "")
        aspectos = row.get("Aspectos Valorizados", "")

        badge_class = "badge-promotor" if cat == "Promotor" else ("badge-detrator" if cat == "Detrator" else "badge-neutro")
        nome_txt  = f" · {nome}"  if nome  else ""
        email_txt = f" · {email}" if email else ""

        comment_block  = f"<div style='background:#F9F9F9;border-radius:8px;padding:10px 14px;margin-bottom:10px;font-size:14px;color:#444'><b>{t['comment']}:</b> {cond}</div>"   if cond     else ""
        suggest_block  = f"<div style='background:#F9F9F9;border-radius:8px;padding:10px 14px;margin-bottom:10px;font-size:14px;color:#444'><b>{t['suggestion']}:</b> {sugestao}</div>" if sugestao else ""
        aspects_block  = f"<div style='font-size:13px;color:#888'><b>{t['values']}:</b> {aspectos}</div>"                                                                              if aspectos else ""

        st.markdown(f"""
        <div class="response-card">
          <div style="display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:12px">
            <div>
              <span style="font-size:18px;font-weight:800;color:{DARK}">{empresa}</span>
              <span style="font-size:13px;color:#888">{nome_txt}{email_txt}</span>
            </div>
            <div style="text-align:right">
              <span class="badge {badge_class}">NPS {nps_val} · {cat}</span>
              <div style="font-size:11px;color:#AAA;margin-top:4px">{data}</div>
            </div>
          </div>
          {comment_block}{suggest_block}{aspects_block}
        </div>
        """, unsafe_allow_html=True)

    section(t["scores"])
    detail_cols = ["Empresa", "NPS (0–10)", "Resultados (1–10)", "Clareza Valor Agregado (1–10)",
                   "Integração — Rapidez", "Integração — Qualidade", "Integração — Facilidade",
                   "Suporte Técnico", "Comunicação da Equipe"]
    available = [c for c in detail_cols if c in dff.columns]
    st.dataframe(dff[available].sort_values("NPS (0–10)", ascending=False),
                 use_container_width=True, hide_index=True)

    section(t["export"])
    csv = dff.to_csv(index=False).encode("utf-8")
    st.download_button(label=t["download"], data=csv,
                       file_name=t["filename"], mime="text/csv")
