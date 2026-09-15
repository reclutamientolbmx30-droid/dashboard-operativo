import streamlit as st
import pandas as pd
import plotly.express as px
import os

# --- 1. CONFIGURACIÓN DE LA PÁGINA ---
st.set_page_config(page_title="Dashboard FGR | Operaciones", page_icon="⚙️", layout="wide")

# --- CSS PERSONALIZADO (Ajustes de espacio en Tarjetas) ---
st.markdown("""
    <style>
    .stApp { background-color: #f8fafc !important; }
    
    /* Reducimos un poco el 'padding' interno para dar más espacio a los números */
    div[data-testid="stMetric"] {
        background-color: #ffffff !important;
        border: 1px solid #f1f5f9 !important;
        padding: 15px 15px !important; 
        border-radius: 20px !important; 
        box-shadow: 0 10px 40px -10px rgba(0, 0, 0, 0.05) !important;
        transition: transform 0.2s ease, box-shadow 0.2s ease;
    }
    div[data-testid="stMetric"]:hover {
        transform: translateY(-3px);
        box-shadow: 0 15px 50px -10px rgba(0, 0, 0, 0.08) !important;
    }
    
    div[data-testid="stMetricLabel"] p, div[data-testid="stMetric"] label {
        color: #64748b !important;
        font-size: 12px !important;
        font-weight: 600 !important;
        text-transform: capitalize !important;
        letter-spacing: 0.3px;
    }
    
    /* Disminuimos la fuente de los números y forzamos a que no se corten */
    div[data-testid="stMetricValue"], div[data-testid="stMetricValue"] div {
        color: #1e293b !important;
        font-weight: 800 !important;
        font-size: 1.3rem !important; 
        margin-top: 5px;
        white-space: nowrap !important;
    }
    
    .plot-container > div {
        border-radius: 20px !important;
        box-shadow: 0 10px 40px -10px rgba(0, 0, 0, 0.04) !important;
        padding: 20px;
        background-color: #ffffff !important;
        border: 1px solid #f1f5f9 !important;
    }
    
    .stTabs [data-baseweb="tab-list"] { gap: 24px; }
    .stTabs [data-baseweb="tab"] {
        height: 50px;
        white-space: pre-wrap;
        background-color: transparent;
        border-radius: 8px 8px 0px 0px;
        padding-top: 10px;
        padding-bottom: 10px;
        color: #64748b;
    }
    .stTabs [aria-selected="true"] { color: #4f46e5 !important; font-weight: bold; }
    </style>
""", unsafe_allow_html=True)

# --- 2. EXTRACCIÓN, LIMPIEZA Y ABREVIATURAS ---
@st.cache_data(ttl=600)
def load_data():
    file_path = "DASHBOARD.xlsx"
    df = pd.read_excel(file_path, sheet_name="Dash")
    
    kpis = {
        "Cumplimiento": pd.to_numeric(df.iloc[3, 2], errors='coerce'),
        "Productividad": pd.to_numeric(df.iloc[3, 5], errors='coerce'),
        "Dias_Retraso": pd.to_numeric(df.iloc[3, 11], errors='coerce'),
        "Piezas_Prog": pd.to_numeric(df.iloc[6, 2], errors='coerce'),
        "Piezas_Prod": pd.to_numeric(df.iloc[6, 3], errors='coerce'),
        "Horas_Prog": pd.to_numeric(df.iloc[6, 5], errors='coerce'),
        "Horas_Prod": pd.to_numeric(df.iloc[6, 6], errors='coerce'),
        "WO_Atrasadas": pd.to_numeric(df.iloc[6, 9], errors='coerce')
    }
    
    # DICCIONARIO DE ABREVIATURAS DE CLIENTES
    abreviaturas = {
        'Maniobras y Servicios Industriales del Centro SA de CV': 'MSI Centro',
        'Sterling Products, Inc. dba ACS Group': 'Sterling / ACS',
        'FGR TRANSFORMACIONES METALICAS SA DE CV': 'FGR Transf.',
        'SIEMENS, S.A. DE C.V.': 'Siemens SA',
        'ABB MEXICO, S.A. DE C.V.': 'ABB México',
        'XNRGY Climate Systems (US) LLC': 'XNRGY',
        'Dana Automotive Manufacturing Inc': 'Dana Auto',
        'SIEMENS ENERGY GTO': 'Siemens Energy',
        'FGR PROYECTOS INTEGRALES E INDUSTRIALES SA DE CV': 'FGR Proyectos',
        'ASP QRO, S de R.L de C.V': 'ASP Qro'
    }
    
    df_clientes_horas = df.iloc[51:61, [2, 3, 4, 5]].copy()
    df_clientes_horas.columns = ['Cliente', 'Programado', 'Producido', '% Avance']
    df_clientes_horas.dropna(subset=['Cliente'], inplace=True)
    df_clientes_horas['Cliente'] = df_clientes_horas['Cliente'].replace(abreviaturas) # Aplica abreviaturas
    for col in ['Programado', 'Producido', '% Avance']:
        df_clientes_horas[col] = pd.to_numeric(df_clientes_horas[col], errors='coerce').fillna(0)
        
    df_clientes_piezas = df.iloc[65:75, [2, 3, 4, 5]].copy()
    df_clientes_piezas.columns = ['Cliente', 'Programado', 'Producido', '% Avance']
    df_clientes_piezas.dropna(subset=['Cliente'], inplace=True)
    df_clientes_piezas['Cliente'] = df_clientes_piezas['Cliente'].replace(abreviaturas) # Aplica abreviaturas
    for col in ['Programado', 'Producido', '% Avance']:
        df_clientes_piezas[col] = pd.to_numeric(df_clientes_piezas[col], errors='coerce').fillna(0)
        
    df_procesos = df.iloc[63:85, [7, 8, 9, 10]].copy()
    df_procesos.columns = ['Proceso', 'Programado', 'Producido', '% Avance']
    df_procesos.dropna(subset=['Proceso'], inplace=True)
    for col in ['Programado', 'Producido', '% Avance']:
        df_procesos[col] = pd.to_numeric(df_procesos[col], errors='coerce').fillna(0)
        
    return kpis, df_clientes_horas, df_clientes_piezas, df_procesos

try:
    kpis, df_clientes_horas, df_clientes_piezas, df_procesos = load_data()
except Exception as e:
    st.error(f"Error al cargar Excel: {e}")
    st.stop()

# --- 3. ENCABEZADO ---
col_logo, col_titulo = st.columns([1, 7])

with col_logo:
    if os.path.exists("logo_fgr.png"):
        st.image("logo_fgr.png", width=140)

with col_titulo:
    st.title("Dashboard Operativo")
    st.markdown("<p style='color:#64748b; font-size:15px; margin-top:-15px;'>Visión general y estatus de proyectos</p>", unsafe_allow_html=True)

st.write("")

# --- 4. BLOQUE DE KPIs ---
col1, col2, col3, col4, col5 = st.columns(5)

with col1:
    st.metric(label="Cumplimiento", value=f"{kpis['Cumplimiento'] * 100:.1f}%")
with col2:
    st.metric(label="Productividad", value=f"{kpis['Productividad']:.2f}")
with col3:
    st.metric(label="Piezas (Prod/Prog)", value=f"{kpis['Piezas_Prod']:,.0f} / {kpis['Piezas_Prog']:,.0f}")
with col4:
    st.metric(label="Horas (Prod/Prog)", value=f"{kpis['Horas_Prod']:,.0f} / {kpis['Horas_Prog']:,.0f}")
with col5:
    st.metric(label="Órdenes Atrasadas", value=f"{kpis['WO_Atrasadas']:,.0f}", delta=f"{kpis['Dias_Retraso']:.1f} d. prom.", delta_color="inverse")

st.markdown("<br>", unsafe_allow_html=True)

# --- 5. FILTROS ---
if os.path.exists("logo_fgr.png"):
    st.sidebar.image("logo_fgr.png", width=120)
    st.sidebar.markdown("<br>", unsafe_allow_html=True)
    
st.sidebar.markdown("### 🎛️ Filtros")
lista_clientes = df_clientes_horas['Cliente'].unique().tolist()
clientes_seleccionados = st.sidebar.multiselect(
    "Seleccionar Cliente(s):",
    options=lista_clientes,
    default=lista_clientes
)

df_ch_filtrado = df_clientes_horas[df_clientes_horas['Cliente'].isin(clientes_seleccionados)]
df_cp_filtrado = df_clientes_piezas[df_clientes_piezas['Cliente'].isin(clientes_seleccionados)]

# --- 6. GRÁFICOS ---
modern_colors = {'Programado': '#cbd5e1', 'Producido': '#4f46e5'} 

def style_plotly_fig(fig):
    fig.update_layout(
        plot_bgcolor='rgba(0,0,0,0)', 
        paper_bgcolor='rgba(0,0,0,0)',
        legend=dict(orientation="h", yanchor="bottom", y=1.05, xanchor="right", x=1, title=""),
        margin=dict(t=60, b=20, l=20, r=20),
        font=dict(color="#64748b")
    )
    # Se añade la rotación a 0 para que los nombres cortos queden completamente horizontales y legibles
    fig.update_xaxes(showgrid=False, showline=True, linecolor='#e2e8f0', title="", tickangle=0)
    fig.update_yaxes(showgrid=True, gridcolor='#f1f5f9', zeroline=False, title="")
    return fig

tab1, tab2 = st.tabs(["📊 Rendimiento por Cliente", "⚙️ Desglose por Procesos"])

with tab1:
    colA, colB = st.columns(2)
    
    with colA:
        df_melted_horas = df_ch_filtrado.melt(id_vars='Cliente', value_vars=['Programado', 'Producido'], var_name='Tipo', value_name='Horas')
        fig_horas = px.bar(df_melted_horas, x='Cliente', y='Horas', color='Tipo', barmode='group',
                           title="<b>Horas:</b> Programadas vs Producidas", 
                           color_discrete_map=modern_colors)
        fig_horas = style_plotly_fig(fig_horas)
        st.plotly_chart(fig_horas, use_container_width=True)

    with colB:
        df_melted_piezas = df_cp_filtrado.melt(id_vars='Cliente', value_vars=['Programado', 'Producido'], var_name='Tipo', value_name='Piezas')
        fig_piezas = px.bar(df_melted_piezas, x='Cliente', y='Piezas', color='Tipo', barmode='group',
                            title="<b>Piezas:</b> Programadas vs Producidas", 
                            color_discrete_map={'Programado': '#cbd5e1', 'Producido': '#0284c7'})
        fig_piezas = style_plotly_fig(fig_piezas)
        st.plotly_chart(fig_piezas, use_container_width=True)

with tab2:
    df_melted_procesos = df_procesos.melt(id_vars='Proceso', value_vars=['Programado', 'Producido'], var_name='Tipo', value_name='Horas')
    fig_procesos = px.area(df_melted_procesos, x='Proceso', y='Horas', color='Tipo', 
                          title="<b>Carga de Trabajo:</b> Distribución por Área Productiva", 
                          color_discrete_map=modern_colors, line_shape='spline')
    fig_procesos = style_plotly_fig(fig_procesos)
    fig_procesos.update_xaxes(tickangle=-45) # Aquí dejamos el ángulo en 45 porque los procesos pueden ser más largos
    st.plotly_chart(fig_procesos, use_container_width=True)