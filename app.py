import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import openpyxl
from collections import defaultdict, Counter
import io
from datetime import date

# ── Config page ──────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Thales — Besoins & Candidatures",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ── CSS ──────────────────────────────────────────────────────────────────
st.markdown("""
<style>
    .main { padding-top: 1rem; }
    .metric-card {
        background: linear-gradient(135deg, #1F3864, #2E75B6);
        border-radius: 12px; padding: 20px; color: white;
        text-align: center; margin: 4px;
    }
    .metric-card.green  { background: linear-gradient(135deg, #375623, #4CAF50); }
    .metric-card.blue   { background: linear-gradient(135deg, #1a5276, #2E75B6); }
    .metric-card.orange { background: linear-gradient(135deg, #784212, #C55A11); }
    .metric-card.red    { background: linear-gradient(135deg, #7B0000, #C00000); }
    .metric-card h2 { font-size: 2.4rem; margin: 0; font-weight: 700; }
    .metric-card p  { margin: 4px 0 0; font-size: 0.85rem; opacity: 0.9; }
    .metric-card .pct { font-size: 0.8rem; opacity: 0.75; }
    .fournisseur-badge {
        display: inline-block; padding: 4px 14px; border-radius: 20px;
        font-weight: bold; font-size: 0.9rem; margin: 2px;
    }
    .badge-ri  { background: #D6EEF1; color: #1F6B75; }
    .badge-exp { background: #EAD1FF; color: #7030A0; }
    .stAlert { border-radius: 8px; }
    div[data-testid="stSidebarNav"] { display: none; }
</style>
""", unsafe_allow_html=True)

# ── Constantes ────────────────────────────────────────────────────────────
CAT_ORDER  = ['acceptee','a_selectionner','a_etudier','toutes_refusees','sans_cand']
CAT_LBL    = {
    'acceptee':       '✅ Acceptée',
    'a_selectionner': '🔵 A sélectionner',
    'a_etudier':      '🔵 A étudier',
    'toutes_refusees':'🟠 Toutes refusées',
    'sans_cand':      '🔴 Sans candidature',
}
CAT_COLORS = {
    'acceptee':       '#375623',
    'a_selectionner': '#2E75B6',
    'a_etudier':      '#5BA3E0',
    'toutes_refusees':'#C55A11',
    'sans_cand':      '#C00000',
}
CAT_PRIO = {'acceptee':0,'a_selectionner':1,'a_etudier':2,'toutes_refusees':3,'sans_cand':4}

# ── Chargement & traitement ───────────────────────────────────────────────
def load_edb(uploaded_file, idx_edb=None, idx_nb=14, idx_acc=19):
    wb = openpyxl.load_workbook(uploaded_file)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    data = [r for r in rows[1:] if r[0] and r[0] != 'Total général']
    edb_d = {}
    for r in data:
        num = r[9]
        if not num or num == 'N\\A': continue
        statut = r[idx_edb] if idx_edb is not None else None
        if num not in edb_d:
            edb_d[num] = {
                'statut': statut, 'sem': r[5], 'site': r[2], 'qual': r[7],
                'nb_cand': 0, 'acceptee_ci': 0, 'statuts_cand': set(), 'agences': set()
            }
        edb_d[num]['nb_cand']     = max(edb_d[num]['nb_cand'],     r[idx_nb]  or 0)
        edb_d[num]['acceptee_ci'] = max(edb_d[num]['acceptee_ci'], r[idx_acc] or 0)
        if r[10] and r[10] != '-': edb_d[num]['statuts_cand'].add(r[10])
        if r[3]: edb_d[num]['agences'].add(r[3])
    return edb_d

def categorise(e):
    sc = e['statuts_cand']; nb = e['nb_cand']
    if nb == 0:                    return 'sans_cand'
    elif 'Acceptée' in sc:         return 'acceptee'
    elif 'A sélectionner' in sc:   return 'a_selectionner'
    elif 'A étudier' in sc:        return 'a_etudier'
    else:                          return 'toutes_refusees'

def edb_to_df(edb_d, fournisseur):
    rows = []
    for num, e in edb_d.items():
        e['cat'] = categorise(e)
        rows.append({
            'num': num, 'site': e['site'], 'qual': e['qual'], 'sem': str(e['sem']) if e['sem'] else '',
            'statut_edb': e['statut'] or '', 'nb_cand': e['nb_cand'],
            'cat': e['cat'], 'cat_label': CAT_LBL[e['cat']],
            'agences': '; '.join(sorted(e['agences'])),
            'nb_agences': len(e['agences']),
            'fournisseur': fournisseur,
        })
    return pd.DataFrame(rows)

def consolider(df_ri, df_exp):
    all_nums = set(df_ri['num'].tolist()) | set(df_exp['num'].tolist())
    rows = []
    ri_map  = df_ri.set_index('num').to_dict('index') if df_ri is not None else {}
    exp_map = df_exp.set_index('num').to_dict('index') if df_exp is not None else {}
    for num in all_nums:
        ri  = ri_map.get(num)
        ex  = exp_map.get(num)
        base = ex if ex else ri
        cat_ri = ri['cat'] if ri else 'sans_cand'
        cat_ex = ex['cat'] if ex else 'sans_cand'
        cat_best = cat_ri if CAT_PRIO[cat_ri] <= CAT_PRIO[cat_ex] else cat_ex
        rows.append({
            'num': num, 'site': base['site'], 'qual': base['qual'], 'sem': base['sem'],
            'statut_edb': ex['statut_edb'] if ex else '',
            'nb_cand_ri': ri['nb_cand'] if ri else 0,
            'nb_cand_exp': ex['nb_cand'] if ex else 0,
            'nb_cand_total': (ri['nb_cand'] if ri else 0) + (ex['nb_cand'] if ex else 0),
            'cat_ri': cat_ri, 'cat_exp': cat_ex, 'cat': cat_best,
            'cat_label': CAT_LBL[cat_best],
            'fournisseur': ('RI+EXP' if ri and ex else ('RI' if ri else 'EXP')),
            'dans_ri': ri is not None, 'dans_exp': ex is not None,
        })
    return pd.DataFrame(rows)

# ── Composants UI ─────────────────────────────────────────────────────────
def kpi_row(df, couleur_principale):
    total = len(df)
    counts = df['cat'].value_counts()
    cols = st.columns(5)
    for i, (c_id, style) in enumerate(zip(CAT_ORDER,
        ['','green','blue','orange','red'])):
        val = counts.get(c_id, 0)
        pct = val/total*100 if total else 0
        with cols[i]:
            st.markdown(f"""
            <div class="metric-card {style}">
                <p>{CAT_LBL[c_id]}</p>
                <h2>{val}</h2>
                <p class="pct">{pct:.0f}% du total</p>
            </div>
            """, unsafe_allow_html=True)

def graphique_camembert(df, titre):
    counts = df['cat'].value_counts().reindex(CAT_ORDER, fill_value=0)
    labels = [CAT_LBL[c] for c in CAT_ORDER]
    values = counts.values
    colors = [CAT_COLORS[c] for c in CAT_ORDER]
    fig = go.Figure(go.Pie(
        labels=labels, values=values,
        hole=0.45,
        marker_colors=colors,
        textinfo='percent+label',
        textfont_size=11,
    ))
    fig.update_layout(
        title=titre, showlegend=False,
        height=320, margin=dict(t=40,b=10,l=10,r=10),
        font=dict(family='Arial')
    )
    return fig

def graphique_barres_semaine(df, titre):
    sem_cat = df.groupby(['sem','cat']).size().unstack(fill_value=0)
    for c in CAT_ORDER:
        if c not in sem_cat.columns: sem_cat[c]=0
    sem_cat = sem_cat[CAT_ORDER].sort_index()
    fig = go.Figure()
    for c_id in CAT_ORDER:
        fig.add_trace(go.Bar(
            name=CAT_LBL[c_id], x=sem_cat.index,
            y=sem_cat[c_id], marker_color=CAT_COLORS[c_id]
        ))
    fig.update_layout(
        barmode='stack', title=titre,
        xaxis_title='Semaine de diffusion', yaxis_title="Nb Expressions de Besoin",
        height=380, margin=dict(t=40,b=60,l=40,r=10),
        legend=dict(orientation='h', y=-0.25),
        font=dict(family='Arial'), plot_bgcolor='#FAFAFA'
    )
    fig.update_xaxes(tickangle=45)
    return fig

def graphique_sites(df, titre, top_n=15):
    site_cat = df.groupby(['site','cat']).size().unstack(fill_value=0)
    for c in CAT_ORDER:
        if c not in site_cat.columns: site_cat[c]=0
    site_cat['critique'] = site_cat.get('sans_cand',0) + site_cat.get('toutes_refusees',0)
    site_cat = site_cat.sort_values('critique', ascending=False).head(top_n)
    site_cat = site_cat[CAT_ORDER].iloc[::-1]
    fig = go.Figure()
    for c_id in CAT_ORDER:
        fig.add_trace(go.Bar(
            name=CAT_LBL[c_id], y=site_cat.index,
            x=site_cat[c_id], orientation='h',
            marker_color=CAT_COLORS[c_id]
        ))
    fig.update_layout(
        barmode='stack', title=titre,
        xaxis_title="Nb Expressions", yaxis_title='',
        height=max(380, top_n*28), margin=dict(t=40,b=40,l=10,r=10),
        legend=dict(orientation='h', y=-0.12),
        font=dict(family='Arial'), plot_bgcolor='#FAFAFA'
    )
    return fig

def tableau_critiques(df, statut_filtre='Diffusée'):
    if statut_filtre and 'statut_edb' in df.columns:
        df_crit = df[df['statut_edb']==statut_filtre]
    else:
        df_crit = df.copy()
    df_crit = df_crit[df_crit['cat'].isin(['sans_cand','toutes_refusees'])].copy()
    df_crit = df_crit.sort_values(['cat','site'])
    if df_crit.empty:
        st.success("✅ Aucune expression critique dans cette sélection.")
        return
    cols_show = ['num','site','qual','sem','cat_label']
    if 'nb_cand_ri' in df_crit.columns:
        cols_show += ['nb_cand_ri','nb_cand_exp']
    else:
        cols_show += ['nb_cand']
    rename = {
        'num':'N° Expression','site':'Site','qual':'Qualification',
        'sem':'Semaine','cat_label':'Situation',
        'nb_cand':'Nb Cand.','nb_cand_ri':'Cand. RI','nb_cand_exp':'Cand. EXP'
    }
    df_show = df_crit[[c for c in cols_show if c in df_crit.columns]].rename(columns=rename)

    def color_row(row):
        if 'Sans candidature' in str(row.get('Situation','')):
            return ['background-color: #FCE4D6']*len(row)
        elif 'refusées' in str(row.get('Situation','')).lower():
            return ['background-color: #FBE4D5']*len(row)
        return ['']*len(row)

    st.dataframe(
        df_show.style.apply(color_row, axis=1),
        use_container_width=True, hide_index=True,
        height=min(400, 40+len(df_show)*35)
    )

def export_excel_button(df, nom):
    """Génère un Excel téléchargeable"""
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    wb = Workbook(); ws = wb.active; ws.title = "Données"
    def tb():
        s=Side(style='thin',color="BFBFBF")
        return Border(left=s,right=s,top=s,bottom=s)

    cols = [c for c in ['num','site','qual','sem','statut_edb','cat_label',
                         'nb_cand','nb_cand_ri','nb_cand_exp','fournisseur'] if c in df.columns]
    rename = {'num':'N° Expression','site':'Site','qual':'Qualification','sem':'Semaine',
              'statut_edb':'Statut EdB','cat_label':'Situation','nb_cand':'Nb Cand.',
              'nb_cand_ri':'Cand. RI','nb_cand_exp':'Cand. EXP','fournisseur':'Fournisseur'}
    headers = [rename.get(c,c) for c in cols]

    for i,h in enumerate(headers,1):
        cell=ws.cell(1,i,h)
        cell.font=Font(name='Arial',bold=True,color='FFFFFF',size=9)
        cell.fill=PatternFill('solid',start_color='1F3864')
        cell.alignment=Alignment(horizontal='center',vertical='center',wrap_text=True)
        cell.border=tb()
        ws.column_dimensions[get_column_letter(i)].width=18
    ws.row_dimensions[1].height=28

    CAT_BG_XLSX={'acceptee':'E2EFDA','a_selectionner':'D6E4F0','a_etudier':'EBF3FB',
                 'toutes_refusees':'FBE4D5','sans_cand':'FCE4D6'}
    for ri,row in enumerate(df[cols].itertuples(index=False),2):
        cat_val=df.iloc[ri-2]['cat'] if 'cat' in df.columns else 'sans_cand'
        bg=CAT_BG_XLSX.get(cat_val,'FFFFFF')
        for ci,v in enumerate(row,1):
            cell=ws.cell(ri,ci,str(v) if v is not None else '')
            cell.font=Font(name='Arial',size=8)
            cell.fill=PatternFill('solid',start_color=bg)
            cell.alignment=Alignment(horizontal='left',vertical='center')
            cell.border=tb()
        ws.row_dimensions[ri].height=14

    buf=io.BytesIO(); wb.save(buf); buf.seek(0)
    st.download_button(
        f"⬇️ Télécharger {nom}.xlsx", buf,
        file_name=f"{nom}_{date.today().strftime('%Y%m%d')}.xlsx",
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        use_container_width=True
    )

# ══════════════════════════════════════════════════════════════════════════
# INTERFACE PRINCIPALE
# ══════════════════════════════════════════════════════════════════════════

# ── Sidebar ───────────────────────────────────────────────────────────────
with st.sidebar:
    st.image("https://upload.wikimedia.org/wikipedia/commons/thumb/1/16/Randstad_logo.svg/320px-Randstad_logo.svg.png", width=160)
    st.markdown("---")
    st.markdown("### 📁 Upload des fichiers")

    file_ri = st.file_uploader(
        "Fichier Randstad Intérim (.xlsx)",
        type=['xlsx'], key='ri',
        help="Fichier pivot Randstad Intérim — sans colonne Statut EdB"
    )
    file_exp = st.file_uploader(
        "Fichier Expectra (.xlsx)",
        type=['xlsx'], key='exp',
        help="Fichier pivot Expectra — avec colonne Statut EdB"
    )

    st.markdown("---")
    st.markdown("### 🔎 Filtres")

    if file_ri or file_exp:
        vue = st.radio(
            "Vue fournisseur",
            options=['🔀 Consolidé RI+EXP', '🏢 Randstad Intérim', '📊 Expectra'],
            disabled=(not file_ri or not file_exp) if True else False
        )
        if not file_ri and vue == '🏢 Randstad Intérim':
            st.warning("Uploadez le fichier RI")
        if not file_exp and vue == '📊 Expectra':
            st.warning("Uploadez le fichier EXP")
    else:
        vue = None

    st.markdown("---")
    st.caption(f"📅 Extraction : {date.today().strftime('%d/%m/%Y')}")
    st.caption("Thales — Besoins & Candidatures")

# ── Zone principale ───────────────────────────────────────────────────────
st.title("📊 Tableau de Bord — Thales Besoins & Candidatures")

if not file_ri and not file_exp:
    st.info("👈 **Commencez par uploader vos fichiers** dans la barre latérale à gauche.")
    col1, col2 = st.columns(2)
    with col1:
        st.markdown("""
        #### 📂 Fichiers attendus
        - **Randstad Intérim** : fichier pivot RI (20 colonnes)
        - **Expectra** : fichier pivot EXP (21 colonnes, avec Statut EdB)
        - Les deux peuvent être uploadés ensemble ou séparément
        """)
    with col2:
        st.markdown("""
        #### 📊 Ce que vous obtiendrez
        - KPIs par fournisseur et consolidés
        - Graphiques interactifs par semaine et par site
        - Tableau des expressions critiques
        - Export Excel téléchargeable
        """)
    st.stop()

# ── Chargement des données ────────────────────────────────────────────────
@st.cache_data
def charger_ri(f):
    return load_edb(f, idx_edb=None, idx_nb=14, idx_acc=19)

@st.cache_data
def charger_exp(f):
    return load_edb(f, idx_edb=11, idx_nb=15, idx_acc=20)

df_ri = df_exp = df_conso = None

if file_ri:
    try:
        edb_ri = charger_ri(file_ri)
        df_ri = edb_to_df(edb_ri, 'Randstad Intérim')
        st.sidebar.success(f"✅ RI chargé — {len(df_ri)} expressions")
    except Exception as ex:
        st.sidebar.error(f"❌ Erreur RI : {ex}")

if file_exp:
    try:
        edb_exp = charger_exp(file_exp)
        df_exp = edb_to_df(edb_exp, 'Expectra')
        st.sidebar.success(f"✅ EXP chargé — {len(df_exp)} expressions")
    except Exception as ex:
        st.sidebar.error(f"❌ Erreur EXP : {ex}")

if df_ri is not None and df_exp is not None:
    df_conso = consolider(df_ri, df_exp)

# ── Sélection vue active ──────────────────────────────────────────────────
if vue == '🏢 Randstad Intérim' and df_ri is not None:
    df_actif = df_ri
    nom_vue = "Randstad Intérim"
    couleur = "#1F6B75"
    has_statut = False
elif vue == '📊 Expectra' and df_exp is not None:
    df_actif = df_exp
    nom_vue = "Expectra"
    couleur = "#7030A0"
    has_statut = True
elif df_conso is not None:
    df_actif = df_conso
    nom_vue = "Consolidé RI + EXP"
    couleur = "#1F3864"
    has_statut = True
elif df_ri is not None:
    df_actif = df_ri; nom_vue = "Randstad Intérim"; couleur = "#1F6B75"; has_statut = False
else:
    df_actif = df_exp; nom_vue = "Expectra"; couleur = "#7030A0"; has_statut = True

# ── Filtres dynamiques ────────────────────────────────────────────────────
with st.sidebar:
    if df_actif is not None and len(df_actif):
        sites = sorted(df_actif['site'].dropna().unique())
        site_sel = st.multiselect("Site Thales", sites, default=[])

        sems = sorted(df_actif['sem'].dropna().unique())
        sem_sel = st.multiselect("Semaine de diffusion", sems, default=[])

        if has_statut and 'statut_edb' in df_actif.columns:
            statuts = sorted(df_actif['statut_edb'].dropna().unique())
            statut_sel = st.multiselect("Statut EdB", statuts, default=[])
        else:
            statut_sel = []

df_filtre = df_actif.copy() if df_actif is not None else pd.DataFrame()
if site_sel:   df_filtre = df_filtre[df_filtre['site'].isin(site_sel)]
if sem_sel:    df_filtre = df_filtre[df_filtre['sem'].isin(sem_sel)]
if statut_sel: df_filtre = df_filtre[df_filtre['statut_edb'].isin(statut_sel)]

# ── En-tête vue ───────────────────────────────────────────────────────────
st.markdown(f"### 🏷️ Vue : **{nom_vue}** — {len(df_filtre)} expressions de besoin")
if len(df_filtre) < len(df_actif):
    st.caption(f"🔍 Filtre actif — {len(df_actif)-len(df_filtre)} expressions masquées")

st.markdown("---")

# ── KPIs ──────────────────────────────────────────────────────────────────
kpi_row(df_filtre, couleur)
st.markdown("<br>", unsafe_allow_html=True)

# ── Graphiques ────────────────────────────────────────────────────────────
col_g1, col_g2 = st.columns([3, 2])
with col_g1:
    st.plotly_chart(
        graphique_barres_semaine(df_filtre, f"Expressions de Besoin par Semaine — {nom_vue}"),
        use_container_width=True
    )
with col_g2:
    st.plotly_chart(
        graphique_camembert(df_filtre, f"Répartition — {nom_vue}"),
        use_container_width=True
    )

# ── Graphique sites ───────────────────────────────────────────────────────
st.plotly_chart(
    graphique_sites(df_filtre, f"Top 15 Sites en tension — {nom_vue}"),
    use_container_width=True
)

# ── Comparaison RI vs EXP (si consolidé) ─────────────────────────────────
if nom_vue == "Consolidé RI + EXP" and df_ri is not None and df_exp is not None:
    st.markdown("---")
    st.markdown("#### 🔄 Comparaison Randstad Intérim vs Expectra")
    col_c1, col_c2 = st.columns(2)
    with col_c1:
        st.plotly_chart(graphique_camembert(df_ri, "Randstad Intérim"), use_container_width=True)
    with col_c2:
        st.plotly_chart(graphique_camembert(df_exp, "Expectra"), use_container_width=True)

    # Graphique comparatif groupé
    fig_cmp = go.Figure()
    for fournisseur, df_f, color in [("Randstad Intérim", df_ri, "#1F6B75"), ("Expectra", df_exp, "#7030A0")]:
        counts = df_f['cat'].value_counts().reindex(CAT_ORDER, fill_value=0)
        fig_cmp.add_trace(go.Bar(
            name=fournisseur, x=[CAT_LBL[c] for c in CAT_ORDER],
            y=counts.values, marker_color=color
        ))
    fig_cmp.update_layout(
        barmode='group', title="RI vs Expectra — Comparaison directe",
        height=350, font=dict(family='Arial'),
        xaxis_tickangle=-20, plot_bgcolor='#FAFAFA',
        legend=dict(orientation='h', y=-0.2)
    )
    st.plotly_chart(fig_cmp, use_container_width=True)

# ── Tableau critiques ─────────────────────────────────────────────────────
st.markdown("---")
st.markdown("#### 🚨 Expressions nécessitant une action")

tab1, tab2 = st.tabs([
    f"🔴 Sans candidature ({len(df_filtre[df_filtre['cat']=='sans_cand'])})",
    f"🟠 Toutes refusées ({len(df_filtre[df_filtre['cat']=='toutes_refusees'])})"
])
with tab1:
    df_sans = df_filtre[df_filtre['cat']=='sans_cand'].copy()
    if has_statut and 'statut_edb' in df_sans.columns:
        df_sans = df_sans[df_sans['statut_edb']=='Diffusée']
        st.caption(f"Filtrées sur statut Diffusée — {len(df_sans)} expressions")
    tableau_critiques(df_sans, statut_filtre=None)

with tab2:
    df_ref = df_filtre[df_filtre['cat']=='toutes_refusees'].copy()
    if has_statut and 'statut_edb' in df_ref.columns:
        df_ref = df_ref[df_ref['statut_edb']=='Diffusée']
        st.caption(f"Filtrées sur statut Diffusée — {len(df_ref)} expressions")
    tableau_critiques(df_ref, statut_filtre=None)

# ── Export ────────────────────────────────────────────────────────────────
st.markdown("---")
st.markdown("#### ⬇️ Export")
export_excel_button(df_filtre, f"Thales_{nom_vue.replace(' ','_').replace('+','')}")

