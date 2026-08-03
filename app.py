import streamlit as st
import pandas as pd
import plotly.express as px

# ==============================================================================
# 1. CONFIGURAÇÃO DA PÁGINA E CSS (FIX CIRÚRGICO DA BUSCA E DOWNLOAD DA TABELA)
# ==============================================================================
st.set_page_config(
    page_title="Torre de Controle de Frota e Risco (OnixSat)",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
<style>
    header[data-testid="stHeader"] { visibility: hidden; }
    .block-container { padding-top: 1rem !important; }
    .stApp { background-color: #0E1117; color: #FFFFFF !important; }
    section[data-testid="stSidebar"] { background-color: #161B22; border-right: 1px solid #30363D; }
    
    /* FIX DE COR PARA TÍTULOS, RÓTULOS, CHECKBOXES E RÁDIOS */
    h1, h2, h3, h4, h5, h6, label, p, span, .stMarkdown, .stRadio label {
        color: #FFFFFF !important;
    }
    
    /* AJUSTE DOS BOTÕES RÁDIO */
    div[data-testid="stRadio"] label span {
        color: #FFFFFF !important;
        font-weight: 500;
    }

    /* 🎯 FIX DA CAIXA FLUTUANTE DE BUSCA / DOWNLOAD DA TABELA (GLIDE DATA GRID) */
    div[class*="glide-data-grid-search"],
    div[class*="dvc-grid-search"],
    div[data-testid="stDataFrame"] input,
    div[data-testid="stDataFrame"] button {
        background-color: #21262D !important;
        color: #FFFFFF !important;
        border: 1px solid #30363D !important;
    }
    div[class*="glide-data-grid-search"] *,
    div[class*="data-grid-search"] * {
        color: #FFFFFF !important;
        fill: #FFFFFF !important;
    }

    /* CARDS DE MÉTRICAS DO TOPO */
    .metric-card {
        background-color: #161B22;
        border: 1px solid #30363D;
        border-radius: 8px;
        padding: 15px;
        text-align: center;
        box-shadow: 0 2px 8px rgba(0,0,0,0.2);
    }
    .metric-title { color: #8B949E !important; font-size: 0.85rem; font-weight: 600; margin-bottom: 5px; }
    .metric-value { color: #58A6FF !important; font-size: 2rem; font-weight: bold; }
    .metric-alert { color: #FF4500 !important; font-size: 2rem; font-weight: bold; }
    
    /* ESTILIZAÇÃO DAS ABAS */
    .stTabs [data-baseweb="tab-list"] { 
        gap: 6px; 
        background-color: #0E1117;
        padding-bottom: 5px;
    }
    .stTabs [data-baseweb="tab"] {
        background-color: #21262D !important;
        border: 1px solid #30363D !important;
        border-radius: 6px 6px 0px 0px;
        color: #FFFFFF !important;
        font-weight: 600 !important;
        font-size: 0.9rem !important;
        padding: 8px 14px !important;
    }
    .stTabs [aria-selected="true"] {
        background-color: #1F242D !important;
        color: #58A6FF !important;
        border-bottom: 3px solid #58A6FF !important;
    }
    .stTabs [aria-selected="true"] p {
        color: #58A6FF !important;
    }
</style>
""", unsafe_allow_html=True)

st.title("📊 Torre de Controle: Condução, Risco e Operação (OnixSat)")

# ==============================================================================
# 2. CARREGAMENTO E LEITURA INTELIGENTE DAS 4 PLANILHAS
# ==============================================================================
uploaded_files = st.file_uploader(
    "Envie as planilhas Excel (.xlsx ou .xls)", 
    type=["xlsx", "xls"], 
    accept_multiple_files=True
)

if uploaded_files:
    try:
        df_alertas_cam = None
        df_vel_motoristas = None
        df_log_pos = None
        df_bloqueios = None

        for file in uploaded_files:
            df_temp = pd.read_excel(file)
            df_temp = df_temp.loc[:, ~df_temp.columns.duplicated()].copy()

            if "Vídeo MDVR" in df_temp.columns or "Descrição Alerta" in df_temp.columns:
                df_alertas_cam = df_temp
            elif "Velocidade Excedida (Km/h)" in df_temp.columns:
                df_vel_motoristas = df_temp
            elif "Excesso de Tempo em Movimento" in df_temp.columns or "Tempo Parado" in df_temp.columns:
                df_log_pos = df_temp
            elif "Motivo Bloqueio" in df_temp.values or "Disparado por" in df_temp.values or "RelatMotivosBloq" in str(file.name):
                df_bloqueios = df_temp

        # --- PROCESSAMENTO 1: PLANILHA DE POSIÇÕES (LogPos_GERAL) ---
        df_pos_limpo = None
        df_km_resumo = None

        if df_log_pos is not None:
            df4_clean = df_log_pos.loc[:, ~df_log_pos.columns.duplicated()].copy()
            df_km_resumo = df4_clean.dropna(subset=['Placa'])[['Placa', 'Dist. Percorrida', 'Tempo Parado', 'Velocidade Média', 'Tempo em Movimento', 'Excesso de Tempo em Movimento']].copy()
            df_km_resumo['KM_Num'] = df_km_resumo['Dist. Percorrida'].astype(str).str.replace(' Km', '').str.replace('.', '').astype(float, errors='ignore')

            df4_clean['Placa'] = df4_clean['Placa'].where(df4_clean['Placa'].notna()).ffill()

            df_pos_limpo = df4_clean.dropna(subset=['Unnamed: 1', 'Unnamed: 3']).copy()
            df_pos_limpo = df_pos_limpo[df_pos_limpo['Unnamed: 1'] != 'Data Hora'].copy()
            df_pos_limpo.rename(columns={'Unnamed: 1': 'Data_Hora_Str', 'Unnamed: 3': 'Localizacao', 'Unnamed: 8': 'Km_h'}, inplace=True)
            df_pos_limpo['Data_Hora'] = pd.to_datetime(df_pos_limpo['Data_Hora_Str'], errors='coerce')
            df_pos_limpo['Data'] = df_pos_limpo['Data_Hora'].dt.date

            def extrair_estado(loc_str):
                s = str(loc_str).strip()
                if '(' in s and ')' in s: return s.split('(')[-1].replace(')', '').strip()
                elif '-' in s: return s.split('-')[0].strip()
                return 'Outros'

            def extrair_cidade(loc_str):
                s = str(loc_str).strip()
                if '(' in s: return s.split('(')[0].strip()
                elif '-' in s and len(s.split('-')) > 1: return s.split('-')[1].strip()
                return s

            df_pos_limpo['Estado_UF'] = df_pos_limpo['Localizacao'].apply(extrair_estado)
            df_pos_limpo['Cidade'] = df_pos_limpo['Localizacao'].apply(extrair_cidade)

        # --- PROCESSAMENTO 2: PLANILHA DE BLOQUEIOS (RelatMotivosBloq) ---
        df_bloq_limpo = None
        if df_bloqueios is not None:
            df3_clean = df_bloqueios.loc[:, ~df_bloqueios.columns.duplicated()].copy()
            df3_clean['Placa'] = df3_clean['Placa'].where(df3_clean['Placa'].notna()).ffill()
            df_bloq_limpo = df3_clean.dropna(subset=['Unnamed: 1']).copy()
            df_bloq_limpo = df_bloq_limpo[~df_bloq_limpo['Unnamed: 1'].astype(str).str.contains('Motivo Bloqueio', na=False)].copy()
            df_bloq_limpo.rename(columns={
                'Unnamed: 1': 'Motivo_Bloqueio',
                'Unnamed: 5': 'Saida_Acionada',
                'Unnamed: 6': 'Data_Hora'
            }, inplace=True)
            df_bloq_limpo['dt'] = pd.to_datetime(df_bloq_limpo['Data_Hora'], errors='coerce')

        # --- PROCESSAMENTO 3: CRUZAMENTO CÂMERA X VELOCIDADE ---
        if df_alertas_cam is not None and df_vel_motoristas is not None:
            col_dt_cam = "Data Hora Alerta" if "Data Hora Alerta" in df_alertas_cam.columns else "Data/Hora"
            col_dt_vel = "Data/Hora" if "Data/Hora" in df_vel_motoristas.columns else "Data Hora Alerta"

            df_alertas_cam['dt'] = pd.to_datetime(df_alertas_cam[col_dt_cam], errors='coerce')
            df_vel_motoristas['dt'] = pd.to_datetime(df_vel_motoristas[col_dt_vel], errors='coerce')

            df_cam_sorted = df_alertas_cam.dropna(subset=['dt', 'Placa']).sort_values('dt')
            df_vel_sorted = df_vel_motoristas.dropna(subset=['dt', 'Placa', 'Motorista']).sort_values('dt')

            df_raw = pd.merge_asof(
                df_cam_sorted,
                df_vel_sorted[['Placa', 'dt', 'Motorista']],
                on='dt', by='Placa', direction='nearest'
            )
            df_raw['Motorista'] = df_raw['Motorista'].fillna('Motorista Não Cadastrado')
        elif df_alertas_cam is not None:
            df_raw = df_alertas_cam
            df_raw['dt'] = pd.to_datetime(df_raw['Data Hora Alerta' if 'Data Hora Alerta' in df_raw.columns else 'Data/Hora'], errors='coerce')
            df_raw['Motorista'] = 'Consulte Velocidade'
        elif df_vel_motoristas is not None:
            df_raw = df_vel_motoristas
            df_raw['dt'] = pd.to_datetime(df_raw['Data/Hora' if 'Data/Hora' in df_raw.columns else 'Data Hora Alerta'], errors='coerce')
            if "Descrição Alerta" not in df_raw.columns and "Velocidade Excedida (Km/h)" in df_raw.columns:
                df_raw["Descrição Alerta"] = df_raw["Velocidade Excedida (Km/h)"]
        else:
            df_raw = pd.concat([pd.read_excel(f) for f in uploaded_files], ignore_index=True)
            df_raw['dt'] = pd.to_datetime(df_raw.iloc[:, 1], errors='coerce')

        df_raw = df_raw.loc[:, ~df_raw.columns.duplicated()].copy()

        cols = list(df_raw.columns)
        col_placa = "Placa" if "Placa" in cols else cols[0]
        col_infracao = "Descrição Alerta" if "Descrição Alerta" in cols else cols[1]
        col_motorista = "Motorista" if "Motorista" in cols else "Operador"
        col_km = "Km/h" if "Km/h" in cols else "Velocidade Excedida (Km/h)"

        df_raw[col_placa] = df_raw[col_placa].astype(str).str.strip()
        df_raw[col_infracao] = df_raw[col_infracao].astype(str).str.strip()
        df_raw[col_motorista] = df_raw[col_motorista].astype(str).str.strip()
        if col_km in df_raw.columns: df_raw[col_km] = pd.to_numeric(df_raw[col_km], errors='coerce').fillna(0)

        df_raw['Hora'] = df_raw['dt'].dt.hour.fillna(0).astype(int)

        if df_bloq_limpo is not None and df_vel_motoristas is not None:
            df_bloq_sorted = df_bloq_limpo.dropna(subset=['dt', 'Placa']).sort_values('dt')
            df_vel_sorted = df_vel_motoristas.dropna(subset=['dt', 'Placa', 'Motorista']).sort_values('dt')
            df_bloq_limpo = pd.merge_asof(
                df_bloq_sorted,
                df_vel_sorted[['Placa', 'dt', 'Motorista']],
                on='dt', by='Placa', direction='nearest'
            )
            df_bloq_limpo['Motorista'] = df_bloq_limpo['Motorista'].fillna('Motorista Não Mapeado')
            df_bloq_limpo = df_bloq_limpo.loc[:, ~df_bloq_limpo.columns.duplicated()].copy()

        # ----------------------------------------------------------------------
        # 3. SIDEBAR - FILTROS OPERACIONAIS COMPLETOS
        # ----------------------------------------------------------------------
        st.sidebar.header("🎯 Filtros Globais Operacionais")

        ocultar_rotina = st.sidebar.checkbox("⚠️ Exibir Apenas Infrações Reais (Ocultar Rotina)", value=False)
        filtro_velocidade = st.sidebar.selectbox(
            "🏎️ Filtrar por Velocidade Mínima:",
            options=["Todas as Velocidades", "90+ km/h", "100+ km/h", "110+ km/h", "120+ km/h", "130+ km/h"]
        )

        placas_universais = set()
        if df_raw is not None and col_placa in df_raw.columns:
            placas_universais.update(df_raw[col_placa].dropna().unique())
        if df_pos_limpo is not None and 'Placa' in df_pos_limpo.columns:
            placas_universais.update(df_pos_limpo['Placa'].dropna().unique())
        if df_bloq_limpo is not None and 'Placa' in df_bloq_limpo.columns:
            placas_universais.update(df_bloq_limpo['Placa'].dropna().unique())

        placas_disponiveis = sorted([str(p) for p in placas_universais if str(p) not in ['nan', 'None', '']])
        motoristas_disponiveis = sorted([str(m) for m in df_raw[col_motorista].unique() if str(m) not in ['nan', 'None', '']])
        tipos_disponiveis = sorted([str(t) for t in df_raw[col_infracao].unique() if str(t) not in ['nan', 'None', '']])

        motoristas_selecionados = st.sidebar.multiselect("👤 Filtrar por Motorista:", options=motoristas_disponiveis)
        placas_selecionadas = st.sidebar.multiselect("🚛 Filtrar por Placa:", options=placas_disponiveis)
        tipos_selecionados = st.sidebar.multiselect("⚠️ Filtrar por Tipo de Alerta:", options=tipos_disponiveis)

        df = df_raw.copy()

        if ocultar_rotina:
            df = df[~df[col_infracao].str.contains("Posição Normal|Heartbeat|Transmissão", case=False, na=False)]

        if col_km in df.columns:
            if filtro_velocidade == "90+ km/h": df = df[df[col_km] >= 90]
            elif filtro_velocidade == "100+ km/h": df = df[df[col_km] >= 100]
            elif filtro_velocidade == "110+ km/h": df = df[df[col_km] >= 110]
            elif filtro_velocidade == "120+ km/h": df = df[df[col_km] >= 120]
            elif filtro_velocidade == "130+ km/h": df = df[df[col_km] >= 130]

        if motoristas_selecionados: df = df[df[col_motorista].isin(motoristas_selecionados)]
        if placas_selecionadas: df = df[df[col_placa].isin(placas_selecionadas)]
        if tipos_selecionados: df = df[df[col_infracao].isin(tipos_selecionados)]

        # ==============================================================================
        # 4. MÉTRICAS DO TOPO
        # ==============================================================================
        total_alertas = len(df)
        total_veiculos_frota = len(placas_disponiveis)
        total_km_frota = int(df_km_resumo['KM_Num'].sum()) if df_km_resumo is not None else 0

        st.markdown("<br>", unsafe_allow_html=True)
        m1, m2, m3, m4 = st.columns(4)
        with m1: st.markdown(f'<div class="metric-card"><div class="metric-title">TOTAL DE ALERTAS</div><div class="metric-value">{total_alertas:,}</div></div>', unsafe_allow_html=True)
        with m2: st.markdown(f'<div class="metric-card"><div class="metric-title">TOTAL KM RODADOS (MÊS)</div><div class="metric-value">{total_km_frota:,} km</div></div>', unsafe_allow_html=True)
        with m3: st.markdown(f'<div class="metric-card"><div class="metric-title">VEÍCULOS EM OPERAÇÃO</div><div class="metric-value">{total_veiculos_frota}</div></div>', unsafe_allow_html=True)
        with m4: st.markdown(f'<div class="metric-card"><div class="metric-title">VELOCIDADE MÁXIMA</div><div class="metric-alert">{int(df[col_km].max()) if total_alertas > 0 else 0} km/h</div></div>', unsafe_allow_html=True)

        st.markdown("<br><hr style='border-color: #30363D;'><br>", unsafe_allow_html=True)

        # ==============================================================================
        # 5. ORGANIZAÇÃO EM 6 ABAS ESTRUTURADAS
        # ==============================================================================
        aba1, aba2, aba3, aba4, aba5, aba6 = st.tabs([
            "🗺️ Região & KM Rodado", 
            "📊 Visão Geral de Condução",
            "🔒 Motivos dos Bloqueios",
            "🕒 Pico por Horário & Perfis", 
            "🚨 Duplo Risco", 
            "🏆 Score de Segurança (0 a 100)"
        ])

        # ------------------------------------------------------------------------------
        # ABA 1: REGIÃO & KM RODADO
        # ------------------------------------------------------------------------------
        with aba1:
            st.markdown("### 🗺️ Controle de Quilometragem e Polos de Operação")
            if df_pos_limpo is not None:
                df_p = df_pos_limpo.copy()
                if placas_selecionadas: df_p = df_p[df_p['Placa'].isin(placas_selecionadas)]

                if 'Km_h' in df_p.columns:
                    df_p['Velocidade_Num'] = pd.to_numeric(df_p['Km_h'], errors='coerce').fillna(0)
                else:
                    df_p['Velocidade_Num'] = 0

                f_col1, f_col2 = st.columns(2)
                with f_col1:
                    estados_list = ["Todos os Estados"] + sorted([str(e) for e in df_p['Estado_UF'].unique() if str(e) not in ['nan', 'None', 'Outros']])
                    sel_estado = st.selectbox("📍 Filtrar por Estado (UF):", options=estados_list)
                
                df_p_filtrado = df_p.copy()
                if sel_estado != "Todos os Estados":
                    df_p_filtrado = df_p_filtrado[df_p_filtrado['Estado_UF'] == sel_estado]

                with f_col2:
                    cidades_list = ["Todas as Cidades"] + sorted([str(c) for c in df_p_filtrado['Cidade'].unique() if str(c) not in ['nan', 'None']])
                    sel_cidade = st.selectbox("🏢 Filtrar por Cidade / Região:", options=cidades_list)

                if sel_cidade != "Todas as Cidades":
                    df_p_filtrado = df_p_filtrado[df_p_filtrado['Cidade'] == sel_cidade]

                st.markdown("<br>", unsafe_allow_html=True)

                c1, c2 = st.columns(2)
                with c1:
                    estado_counts = df_p_filtrado['Estado_UF'].value_counts().reset_index()
                    estado_counts.columns = ['Estado (UF)', 'Registros de Presença']
                    fig_uf = px.pie(estado_counts, names='Estado (UF)', values='Registros de Presença', title="📍 Distribuição por Estado (UF)", hole=0.4, color_discrete_sequence=px.colors.qualitative.Pastel)
                    fig_uf.update_layout(template="plotly_dark", paper_bgcolor="#161B22", plot_bgcolor="#161B22", font=dict(color="#FFFFFF"))
                    st.plotly_chart(fig_uf, use_container_width=True)

                with c2:
                    df_operacao_real = df_p_filtrado[df_p_filtrado['Velocidade_Num'] <= 20]
                    if len(df_operacao_real) > 0:
                        cidade_counts = df_operacao_real['Cidade'].value_counts().head(10).reset_index()
                        title_grafico = "🏢 Top 10 Polos de Carga/Descarga (Velocidade ≤ 20 km/h)"
                    else:
                        cidade_counts = df_p_filtrado['Cidade'].value_counts().head(10).reset_index()
                        title_grafico = "🏢 Top 10 Cidades / Polos de Atuação (Geral)"

                    cidade_counts.columns = ['Cidade / Região', 'Registros Parado/Manobra']
                    fig_cidade = px.bar(cidade_counts, x='Registros Parado/Manobra', y='Cidade / Região', orientation='h', title=title_grafico, color='Registros Parado/Manobra', color_continuous_scale="Blues")
                    fig_cidade.update_layout(template="plotly_dark", paper_bgcolor="#161B22", plot_bgcolor="#161B22", font=dict(color="#FFFFFF"), yaxis={"categoryorder": "total ascending"})
                    st.plotly_chart(fig_cidade, use_container_width=True)

                st.markdown("### ⏱️ Detalhamento de Presença e Estadias")

                modo_tabela_regiao = st.radio(
                    "Selecione o Formato de Exibição da Tabela:",
                    options=["📊 Resumo Consolidado (Estado, Cidade e Qtd de Visitas)", "⏱️ Linha do Tempo Detalhada (Entrada, Saída e Dias Seguidos)"],
                    horizontal=True
                )

                lista_estadias = []
                for placa, df_grupo in df_p_filtrado.groupby('Placa'):
                    df_grupo = df_grupo.sort_values('Data_Hora').copy()
                    df_grupo['Cidade_Anterior'] = df_grupo['Cidade'].shift()
                    df_grupo['Mudou'] = (df_grupo['Cidade'] != df_grupo['Cidade_Anterior']).astype(int)
                    df_grupo['Estadia_ID'] = df_grupo['Mudou'].cumsum()

                    estadias = df_grupo.groupby(['Estadia_ID', 'Cidade', 'Estado_UF']).agg(
                        Data_Entrada=('Data', 'min'),
                        Data_Saida=('Data', 'max'),
                        Vel_Media=('Velocidade_Num', 'mean'),
                        Qtd_Registros=('Data', 'count')
                    ).reset_index()

                    estadias['Placa'] = placa
                    estadias['Dias_Seguidos'] = (pd.to_datetime(estadias['Data_Saida']) - pd.to_datetime(estadias['Data_Entrada'])).dt.days + 1
                    
                    def classificar_local(row):
                        if row['Vel_Media'] <= 25 or row['Dias_Seguidos'] > 1:
                            return "📦 Operação / Entrega (Parado/Manobra)"
                        else:
                            return "🛣️ Trânsito / Cidade de Passagem"

                    estadias['Tipo_Permanencia'] = estadias.apply(classificar_local, axis=1)
                    lista_estadias.append(estadias)

                if lista_estadias:
                    df_estadias_final = pd.concat(lista_estadias, ignore_index=True)

                    if "Resumo Consolidado" in modo_tabela_regiao:
                        resumo_cidade_qtd = df_estadias_final.groupby(['Placa', 'Estado_UF', 'Cidade', 'Tipo_Permanencia']).agg(
                            Qtd_Visitas_Estadias=('Estadia_ID', 'count'),
                            Total_Dias_Acumulados=('Dias_Seguidos', 'sum')
                        ).reset_index().sort_values(by='Qtd_Visitas_Estadias', ascending=False).reset_index(drop=True)
                        
                        resumo_cidade_qtd.index = resumo_cidade_qtd.index + 1
                        resumo_cidade_qtd.index.name = "Posição"

                        st.markdown("#### 🔢 Total de Visitas e Estadias por Cidade")
                        st.dataframe(resumo_cidade_qtd, use_container_width=True)
                    else:
                        df_estadias_final = df_estadias_final.sort_values(by=['Dias_Seguidos', 'Data_Entrada'], ascending=[False, False]).reset_index(drop=True)
                        df_estadias_final.index = df_estadias_final.index + 1
                        df_estadias_final.index.name = "Posição"

                        st.markdown("#### 📜 Linha do Tempo de Entrada e Saída Exata")
                        st.dataframe(df_estadias_final[['Placa', 'Estado_UF', 'Cidade', 'Tipo_Permanencia', 'Data_Entrada', 'Data_Saida', 'Dias_Seguidos']], use_container_width=True)

            elif df_km_resumo is not None:
                st.dataframe(df_km_resumo[['Placa', 'Identificador', 'Dist. Percorrida', 'Tempo Parado', 'Velocidade Média']], use_container_width=True)

        # ------------------------------------------------------------------------------
        # ABA 2: VISÃO GERAL DE CONDUÇÃO
        # ------------------------------------------------------------------------------
        with aba2:
            st.markdown("### 🏎️ Indicador de Gravidade por Faixa de Velocidade")
            def categorizar_velocidade(km):
                try: km = float(km)
                except: km = 0.0
                if km >= 130: return "130+ km/h"
                elif km >= 120: return "120 a 129 km/h"
                elif km >= 110: return "110 a 119 km/h"
                elif km >= 100: return "100 a 109 km/h"
                elif km >= 90: return "90 a 99 km/h"
                else: return "< 90 km/h"

            df['Faixa_Velocidade'] = df[col_km].apply(categorizar_velocidade)
            ordem_faixas = ["< 90 km/h", "90 a 99 km/h", "100 a 109 km/h", "110 a 119 km/h", "120 a 129 km/h", "130+ km/h"]
            df_faixas = df['Faixa_Velocidade'].value_counts().reindex(ordem_faixas, fill_value=0).reset_index()
            df_faixas.columns = ["Faixa de Velocidade", "Quantidade"]

            fig_km = px.bar(
                df_faixas, x="Faixa de Velocidade", y="Quantidade", text="Quantidade",
                title="Distribuição de Ocorrências por Categoria de Velocidade", color="Faixa de Velocidade",
                color_discrete_map={"< 90 km/h": "#2EA043", "90 a 99 km/h": "#58A6FF", "100 a 109 km/h": "#D29922", "110 a 119 km/h": "#DB6D28", "120 a 129 km/h": "#F85149", "130+ km/h": "#8B0000"}
            )
            fig_km.update_layout(
                template="plotly_dark", paper_bgcolor="#161B22", plot_bgcolor="#161B22", showlegend=False,
                font=dict(color="#FFFFFF", size=13),
                title_font=dict(color="#FFFFFF", size=16)
            )
            st.plotly_chart(fig_km, use_container_width=True)

            if total_alertas > 0:
                g1, g2 = st.columns(2)
                with g1:
                    alertas_por_tipo = df[col_infracao].value_counts().reset_index()
                    alertas_por_tipo.columns = ["Descrição Alerta", "Quantidade"]
                    
                    fig_infra_tipo = px.bar(
                        alertas_por_tipo, x="Quantidade", y="Descrição Alerta", orientation="h", 
                        title="⚠️ Total de Ocorrências por Tipo", color="Quantidade", 
                        color_continuous_scale=["#00BFFF", "#39D353"]
                    )
                    
                    fig_infra_tipo.update_layout(
                        template="plotly_dark", paper_bgcolor="#161B22", plot_bgcolor="#161B22", 
                        coloraxis_showscale=False,
                        font=dict(color="#FFFFFF", size=13),
                        title_font=dict(color="#FFFFFF", size=16),
                        yaxis={"categoryorder": "total ascending"}
                    )
                    fig_infra_tipo.update_yaxes(tickfont=dict(color="#FFFFFF", size=12))
                    fig_infra_tipo.update_xaxes(tickfont=dict(color="#FFFFFF", size=12))
                    st.plotly_chart(fig_infra_tipo, use_container_width=True)

                with g2:
                    modo_top5 = st.radio("Exibir Ranking por:", options=["Motorista", "Placa"], horizontal=True)
                    col_top5_alvo = col_motorista if modo_top5 == "Motorista" else col_placa
                    top5_itens = df[col_top5_alvo].value_counts().head(5).index
                    df_top5 = df[df[col_top5_alvo].isin(top5_itens)]
                    agrupado_top5 = df_top5.groupby([col_top5_alvo, col_infracao]).size().reset_index(name="Quantidade")
                    
                    fig_empilhado = px.bar(agrupado_top5, x="Quantidade", y=col_top5_alvo, color=col_infracao, orientation="h", title=f"🏆 Top 5 {modo_top5}s", barmode="stack")
                    fig_empilhado.update_layout(
                        template="plotly_dark", paper_bgcolor="#161B22", plot_bgcolor="#161B22", 
                        font=dict(color="#FFFFFF", size=13),
                        title_font=dict(color="#FFFFFF", size=16),
                        legend=dict(title=dict(text="Tipo", font=dict(color="#FFFFFF")), font=dict(color="#FFFFFF"), orientation="h", y=-0.3)
                    )
                    fig_empilhado.update_yaxes(tickfont=dict(color="#FFFFFF", size=12))
                    fig_empilhado.update_xaxes(tickfont=dict(color="#FFFFFF", size=12))
                    st.plotly_chart(fig_empilhado, use_container_width=True)

                st.markdown("### 🔍 Detalhamento Geral de Condução")
                resumo_detalhado = df.groupby([col_motorista, col_placa, col_infracao]).agg(
                    Total_Alertas=(col_infracao, 'count'), Velocidade_Maxima=(col_km, 'max')
                ).reset_index().sort_values(by="Total_Alertas", ascending=False).reset_index(drop=True)
                
                resumo_detalhado.index = resumo_detalhado.index + 1
                resumo_detalhado.index.name = "Posição"

                st.dataframe(resumo_detalhado, use_container_width=True)

        # ------------------------------------------------------------------------------
        # ABA 3: MOTIVOS DOS BLOQUEIOS
        # ------------------------------------------------------------------------------
        with aba3:
            st.markdown("### 🔒 Análise Executiva de Bloqueios e Alarmes Operacionais")
            if df_bloq_limpo is not None:
                df_b = df_bloq_limpo.copy()
                if motoristas_selecionados and "Motorista" in df_b.columns: df_b = df_b[df_b["Motorista"].isin(motoristas_selecionados)]
                if placas_selecionadas: df_b = df_b[df_b["Placa"].isin(placas_selecionadas)]

                b1, b2 = st.columns(2)
                with b1:
                    motivos_top = df_b['Motivo_Bloqueio'].value_counts().head(8).reset_index()
                    motivos_top.columns = ['Motivo do Bloqueio', 'Total']
                    fig_motivos = px.bar(motivos_top, x='Total', y='Motivo do Bloqueio', orientation='h', title="🚨 Principais Motivos de Disparo / Bloqueio", color='Total', color_continuous_scale="Reds")
                    fig_motivos.update_layout(
                        template="plotly_dark", paper_bgcolor="#161B22", plot_bgcolor="#161B22", 
                        font=dict(color="#FFFFFF", size=13),
                        title_font=dict(color="#FFFFFF", size=16),
                        yaxis={"categoryorder": "total ascending"}
                    )
                    st.plotly_chart(fig_motivos, use_container_width=True)

                with b2:
                    if "Motorista" in df_b.columns:
                        top_mot_bloq = df_b['Motorista'].value_counts().head(8).reset_index()
                        top_mot_bloq.columns = ['Motorista', 'Bloqueios']
                        fig_mot_b = px.bar(top_mot_bloq, x='Bloqueios', y='Motorista', orientation='h', title="👤 Top Motoristas Bloqueados", color='Bloqueios', color_continuous_scale="Oranges")
                        fig_mot_b.update_layout(
                            template="plotly_dark", paper_bgcolor="#161B22", plot_bgcolor="#161B22", 
                            font=dict(color="#FFFFFF", size=13),
                            title_font=dict(color="#FFFFFF", size=16),
                            yaxis={"categoryorder": "total ascending"}
                        )
                        st.plotly_chart(fig_mot_b, use_container_width=True)

                st.markdown("### 🔍 Investigação Individual por Motorista / Placa")
                mots_bloq = sorted([str(m) for m in df_b['Motorista'].unique() if str(m) not in ['nan', 'None', 'Motorista Não Mapeado']]) if 'Motorista' in df_b.columns else []
                
                sel_mot_bloq = st.selectbox("👤 Selecione o Motorista para auditar os bloqueios:", options=["Todos os Motoristas"] + mots_bloq)

                df_b_view = df_b.copy()
                if sel_mot_bloq != "Todos os Motoristas": 
                    df_b_view = df_b_view[df_b_view['Motorista'] == sel_mot_bloq]

                modo_vis_bloq = st.radio(
                    "Selecione o Formato de Exibição dos Bloqueios:",
                    options=["📊 Resumo Consolidado (Quantidade por Motivo)", "📜 Log Histórico Detalhado (Data e Hora Exata)"],
                    horizontal=True
                )

                if "Resumo Consolidado" in modo_vis_bloq:
                    grp_cols = ['Placa', 'Motivo_Bloqueio']
                    if 'Motorista' in df_b_view.columns: grp_cols.insert(1, 'Motorista')
                    
                    df_resumo_bloq = df_b_view.groupby(grp_cols).size().reset_index(name='Qtd_Vezes_Bloqueado')
                    df_resumo_bloq = df_resumo_bloq.sort_values(by='Qtd_Vezes_Bloqueado', ascending=False).reset_index(drop=True)
                    df_resumo_bloq.index = df_resumo_bloq.index + 1
                    df_resumo_bloq.index.name = "Posição"

                    st.markdown("#### 🔢 Total de Bloqueios Agrupados por Motivo")
                    st.dataframe(df_resumo_bloq, use_container_width=True)
                else:
                    cols_b_show = ['Placa', 'Motivo_Bloqueio', 'Saida_Acionada', 'Data_Hora']
                    if 'Motorista' in df_b_view.columns: cols_b_show.insert(1, 'Motorista')
                    
                    df_b_show_view = df_b_view[cols_b_show].reset_index(drop=True)
                    df_b_show_view.index = df_b_show_view.index + 1
                    df_b_show_view.index.name = "Posição"

                    st.markdown("#### 📜 Histórico de Transmissões e Ações")
                    st.dataframe(df_b_show_view, use_container_width=True)

            else:
                st.info("💡 Envie a planilha 'RelatMotivosBloq_FROTA COMPLETA.xls' para visualizar o ranking de bloqueios e disparos.")

        # ------------------------------------------------------------------------------
        # ABA 4: PICO POR HORÁRIO & RAIO-X DO MOTORISTA
        # ------------------------------------------------------------------------------
        with aba4:
            st.markdown("### 🕒 Mapeamento de Horários e Auditoria de Hábito de Condução")

            mots_pico = sorted([str(m) for m in df[col_motorista].unique() if str(m) not in ['nan', 'None']])
            sel_mot_pico = st.selectbox("👤 Selecione o Motorista para RAIO-X do Perfil de Horário:", options=["Todos os Motoristas"] + mots_pico)

            df_pico = df.copy()
            if sel_mot_pico != "Todos os Motoristas":
                df_pico = df_pico[df_pico[col_motorista] == sel_mot_pico]

            df_horas_full = df_pico['Hora'].value_counts().reindex(range(24), fill_value=0).reset_index()
            df_horas_full.columns = ['Hora_Dia', 'Total_Alertas']
            df_horas_full['Hora_Str'] = df_horas_full['Hora_Dia'].apply(lambda h: f"{h:02d}:00")

            fig_pico_full = px.bar(
                df_horas_full, x='Hora_Str', y='Total_Alertas', text='Total_Alertas',
                title=f"📊 Perfil de Atividade 24h: {sel_mot_pico}",
                color='Total_Alertas', color_continuous_scale="Reds"
            )
            fig_pico_full.update_layout(
                template="plotly_dark", paper_bgcolor="#161B22", plot_bgcolor="#161B22",
                coloraxis_showscale=False,
                font=dict(color="#FFFFFF", size=13),
                title_font=dict(color="#FFFFFF", size=16),
                xaxis_title="Hora do Dia", yaxis_title="Quantidade de Ocorrências / Registros"
            )
            st.plotly_chart(fig_pico_full, use_container_width=True)

            def categorizar_turno(h):
                if 0 <= h < 6: return "1. 🌙 Madrugada (00h-06h)"
                elif 6 <= h < 12: return "2. 🌅 Manhã (06h-12h)"
                elif 12 <= h < 18: return "3. ☀️ Tarde (12h-18h)"
                else: return "4. 🌌 Noite (18h-00h)"

            df_pico['Turno'] = df_pico['Hora'].apply(categorizar_turno)

            total_noturno = len(df_pico[df_pico['Hora'].isin([18, 19, 20, 21, 22, 23, 0, 1, 2, 3, 4, 5])])
            tot_geral = max(1, len(df_pico))
            pct_noturno = (total_noturno / tot_geral) * 100

            if pct_noturno >= 40 and sel_mot_pico != "Todos os Motoristas":
                st.error(f"🚨 **ALERTA DE PERFIL NOTURNO:** O motorista **{sel_mot_pico}** possui **{pct_noturno:.1f}%** das suas atividades concentradas no período da Noite/Madrugada (18h às 06h). Possível padrão de hora extra / baixo rendimento diurno.")
            elif sel_mot_pico != "Todos os Motoristas":
                st.success(f"✅ **PERFIL DIURNO REGULAR:** O motorista **{sel_mot_pico}** concentra **{(100 - pct_noturno):.1f}%** das suas atividades no horário comercial regular (06h às 18h).")

            st.markdown("---")

            st.markdown("### 🌅 Distribuição por Turnos de Trabalho (Manhã, Tarde, Noite e Madrugada)")
            
            resumo_turno_mot = df_pico.groupby(['Turno', col_motorista]).agg(
                Total_Alertas=(col_infracao, 'count'),
                Velocidade_Maxima=(col_km, 'max')
            ).reset_index().sort_values(by=['Turno', 'Total_Alertas'], ascending=[True, False]).reset_index(drop=True)

            resumo_turno_mot.index = resumo_turno_mot.index + 1
            resumo_turno_mot.index.name = "Posição"

            st.dataframe(
                resumo_turno_mot[['Turno', col_motorista, 'Total_Alertas', 'Velocidade_Maxima']], 
                use_container_width=True
            )

        # ------------------------------------------------------------------------------
        # ABA 5: DUPLO RISCO (> 100 KM/H)
        # ------------------------------------------------------------------------------
        with aba5:
            st.markdown("### 🚨 Análise de Duplo Risco (Cabine + Velocidade Acima de 100 km/h)")
            df_duplo = df[(df[col_km] >= 100) & (df[col_infracao].str.contains("Celular|Cinto|Fadiga|Distração|Bocejo", case=False, na=False))].copy()

            if len(df_duplo) > 0:
                resumo_duplo = df_duplo.groupby(col_motorista).agg(
                    Total_Duplo_Risco=(col_infracao, 'count'),
                    Velocidade_Maxima=(col_km, 'max')
                ).reset_index().sort_values(by='Total_Duplo_Risco', ascending=False).reset_index(drop=True)
                resumo_duplo.index = resumo_duplo.index + 1
                resumo_duplo.index.name = "Posição"

                st.error(f"⚠️ Atenção: Foram detectadas {len(df_duplo):,} ocorrências de Duplo Risco Crítico (> 100 km/h) em toda a frota!")
                st.markdown("#### 📊 Ranking Agrupado de Motoristas em Duplo Risco")
                st.dataframe(resumo_duplo, use_container_width=True)

                st.markdown("---")
                st.markdown("#### 🔍 Investigar Combos do Motorista")
                mot_duplo_sel = st.selectbox("Selecione um Motorista para detalhar os combos de risco:", options=resumo_duplo[col_motorista].unique())

                if mot_duplo_sel:
                    df_mot_d = df_duplo[df_duplo[col_motorista] == mot_duplo_sel]
                    combos = df_mot_d.groupby([col_infracao, col_km]).size().reset_index(name='Vezes').sort_values(by='Vezes', ascending=False).reset_index(drop=True)
                    combos.index = combos.index + 1
                    combos.index.name = "Posição"

                    st.markdown(f"**Combos Críticos de {mot_duplo_sel}:**")
                    st.dataframe(combos, use_container_width=True)
            else:
                st.success("✅ Nenhuma ocorrência de Duplo Risco (> 100 km/h) registrada para os filtros aplicados.")

        # ------------------------------------------------------------------------------
        # ABA 6: SCORE DE SEGURANÇA
        # ------------------------------------------------------------------------------
        with aba6:
            st.markdown("### 🏆 Ranking de Score de Segurança do Motorista (Nota 0 a 100)")
            
            st.info("""
            💡 **Regra de Cálculo da Nota (100 a 0):**  
            Todo motorista inicia com **100 Pontos**. Descontos por ocorrência:
            * 🟢 **Leve (-0,5 pt):** Bocejo / Distração leve.  
            * 🟡 **Média (-1,5 pt):** Sem Cinto / Vel. entre 90 e 100 km/h.  
            * 🔴 **Grave (-3,0 pts):** Vel. entre 101 e 110 km/h.  
            * 💥 **Crítica (-5,0 pts):** Vel. **> 110 km/h** OU **Uso de Celular** OU **Câmera Tampada / Obstruída** OU **Duplo Risco (>100 km/h)**.
            """)

            list_score = []
            for mot, df_mot in df.groupby(col_motorista):
                pontos_perdidos = 0
                c_celular = 0
                c_vel_110 = 0
                c_vel_100 = 0
                c_fadiga = 0
                c_duplo = 0
                c_cam_tampada = 0

                for _, row in df_mot.iterrows():
                    km_val = row[col_km]
                    txt = str(row[col_infracao]).lower()

                    is_duplo = (km_val >= 100) and ("celular" in txt or "fadiga" in txt or "cinto" in txt)
                    if is_duplo: c_duplo += 1

                    if "cobertura" in txt or "não detectado" in txt or "obstru" in txt or "tampada" in txt:
                        c_cam_tampada += 1
                        pontos_perdidos += 5.0

                    if "celular" in txt:
                        c_celular += 1
                        pontos_perdidos += 5.0
                    elif "fadiga" in txt or "bocejo" in txt:
                        c_fadiga += 1
                        pontos_perdidos += 0.5 if "bocejo" in txt else 5.0

                    if km_val > 110:
                        c_vel_110 += 1
                        pontos_perdidos += 5.0
                    elif km_val > 100:
                        c_vel_100 += 1
                        pontos_perdidos += 3.0

                nota_final = max(0.0, 100.0 - pontos_perdidos)
                status = "🟢 Excelência" if nota_final >= 90 else ("🟡 Atenção" if nota_final >= 70 else "🚨 Risco Crítico")
                
                list_score.append({
                    'Motorista': mot,
                    'Nota_Score': round(nota_final, 1),
                    'Total_Alertas': len(df_mot),
                    'Cam_Tampada': c_cam_tampada,
                    'Celular': c_celular,
                    'Vel_Acima_110': c_vel_110,
                    'Vel_101_110': c_vel_100,
                    'Fadiga_Bocejo': c_fadiga,
                    'Duplo_Risco': c_duplo,
                    'Status': status
                })

            df_score_table = pd.DataFrame(list_score).sort_values(by='Nota_Score', ascending=True).reset_index(drop=True)
            
            df_score_table.index = df_score_table.index + 1
            df_score_table.index.name = "Posição"
            
            st.dataframe(df_score_table, use_container_width=True)

    except Exception as e:
        st.error(f"Erro ao processar as planilhas: {e}")
else:
    st.info("Aguardando o envio das planilhas Excel para exibir o dashboard...")