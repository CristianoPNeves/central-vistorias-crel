import datetime
import io
import math
import os
import re
import sqlite3
import unicodedata
import folium
import pandas as pd
import pypdf
import requests
import streamlit as st
import streamlit.components.v1 as components
from streamlit_folium import st_folium
from streamlit_js_eval import get_geolocation

st.set_page_config(
    page_title="Central RIA & Vistorias - CREL",
    page_icon="🛗",
    layout="wide",
    initial_sidebar_state="expanded"
)

DB_NAME = "gestao_ria.db"
DIR_RELATORIOS = "relatorios_gerados"
ARQUIVO_GEO_BASE = "clientes_geolocalizados.xlsx"
PIN_ACESSO = "2026"

os.makedirs(DIR_RELATORIOS, exist_ok=True)


# -----------------------------------------------------------------------------
# 1. CONTROLE DE ACESSO POR PIN
# -----------------------------------------------------------------------------
def autenticar():
    if "auth" not in st.session_state:
        st.session_state["auth"] = False
    if not st.session_state["auth"]:
        st.markdown("""
        <div style="text-align: center; margin-top: 50px; margin-bottom: 20px;">
            <h2 style="color: #1B365D;">🔒 Central de Gestão e Vistorias RIA</h2>
            <p style="color: #666;">CREL Elevadores — Sistema Operacional Integrado</p>
        </div>
        """, unsafe_allow_html=True)
        _, col, _ = st.columns([1, 2, 1])
        with col:
            with st.form("pin_form"):
                pin = st.text_input("PIN de Acesso da Equipe:", type="password", placeholder="••••")
                if st.form_submit_button("Acessar Central", use_container_width=True):
                    if pin == PIN_ACESSO:
                        st.session_state["auth"] = True
                        st.rerun()
                    else:
                        st.error("PIN incorreto. Acesso restrito.")
        st.stop()


autenticar()


# -----------------------------------------------------------------------------
# 2. BANCO DE DADOS & FUNÇÕES GEOGRÁFICAS
# -----------------------------------------------------------------------------
def get_db():
    conexao = sqlite3.connect(DB_NAME)
    conexao.execute('''
        CREATE TABLE IF NOT EXISTS rotas_ativas (
            id TEXT PRIMARY KEY,
            mes_exec TEXT,
            dia_rota TEXT,
            ordem_parada INTEGER,
            cliente TEXT,
            cep TEXT,
            endereco TEXT,
            qtd_elev INTEGER,
            marca TEXT,
            vencto_ria TEXT,
            grupo_status TEXT,
            lat REAL,
            lon REAL,
            waze_url TEXT,
            status_vistoria TEXT,
            data_baixa TEXT,
            vistoriador TEXT,
            obs TEXT
        )
    ''')
    conexao.execute('''
        CREATE TABLE IF NOT EXISTS relatorios_vistorias (
            id_rota TEXT PRIMARY KEY,
            cliente TEXT,
            endereco TEXT,
            data_vistoria TEXT,
            vistoriador TEXT,
            caminho_arquivo TEXT,
            qtd_normas_orc INTEGER,
            qtd_pendencias_sup INTEGER
        )
    ''')
    return conexao


def normalizar_codigo(texto):
    t = str(texto or '').upper().replace(' ', '').replace('_', '').replace('-', '')
    return ''.join(c for c in unicodedata.normalize('NFD', t) if unicodedata.category(c) != 'Mn')


@st.cache_data
def carregar_base_georreferenciada():
    if os.path.exists(ARQUIVO_GEO_BASE):
        df_geo_local = pd.read_excel(ARQUIVO_GEO_BASE)
        df_geo_local['CLI_NORM'] = df_geo_local['CLIENTE'].apply(normalizar_codigo)
        return df_geo_local[['CLI_NORM', 'lat', 'lon']].drop_duplicates(subset=['CLI_NORM'])
    return pd.DataFrame()


def buscar_coords_fallback(cep, endereco_padrao=None):
    _ = endereco_padrao
    cep_limpo = re.sub(r'\D', '', str(cep or ''))
    if len(cep_limpo) == 8:
        try:
            resp = requests.get(f"https://cep.awesomeapi.com.br/json/{cep_limpo}", timeout=3)
            if resp.status_code == 200:
                dados_cep = resp.json()
                return float(dados_cep['lat']), float(dados_cep['lng'])
        except (requests.RequestException, ValueError, KeyError):
            pass
    return -23.5505, -46.6333


def calc_dist_km(lat1, lon1, lat2, lon2):
    raio_terra = 6371.0
    dif_lat = math.radians(lat2 - lat1)
    dif_lon = math.radians(lon2 - lon1)
    a_calc = (math.sin(dif_lat / 2) ** 2 +
              math.cos(math.radians(lat1)) * math.cos(math.radians(lat2)) * math.sin(dif_lon / 2) ** 2)
    return raio_terra * (2 * math.atan2(math.sqrt(a_calc), math.sqrt(1 - a_calc)))


# -----------------------------------------------------------------------------
# 3. EXTRATOR AUTOMÁTICO DE PDF DO SIGA
# -----------------------------------------------------------------------------
def extrair_dados_pdf_siga(pdf_bytes):
    reader = pypdf.PdfReader(io.BytesIO(pdf_bytes))
    texto_completo = ""
    for page in reader.pages:
        texto_completo += page.extract_text() or ""

    linhas = texto_completo.split('\n')
    registros = []
    start = False
    table_lines = []
    for linha in linhas:
        if 'Listados:' in linha:
            break
        if '------------------------------------------------------------------------------------------------------------------------------------' in linha:
            if not start:
                start = True
                continue
        if start:
            table_lines.append(linha)

    for l_str in table_lines:
        l_clean = l_str.strip()
        if not l_clean or 'CLIENTE' in l_clean or '---' in l_clean:
            continue

        match_linha = re.match(r'^([A-Z0-9_\s]{7,9})\s+(\d{8})\s+(.*)$', l_clean)
        if match_linha:
            cli = match_linha.group(1).strip()
            cep = match_linha.group(2).strip()
            resto = match_linha.group(3).strip()

            m_tech = re.search(r'\s+([A-Z])\s+(\d)\s+(\d{1,2})\s+([A-Z0-9/_-]+)\s*(.*)$', resto)
            if m_tech:
                end = resto[:m_tech.start()].strip()
                _tec_letra = m_tech.group(1)
                _com_val = m_tech.group(2)
                qtd = int(m_tech.group(3))
                marca = m_tech.group(4)
                datas_str = m_tech.group(5)

                datas = re.findall(r'\d{2}/\d{2}/\d{4}', datas_str)
                if len(datas) >= 3:
                    vencto = datas[2]
                elif len(datas) == 1:
                    vencto = datas[0]
                else:
                    vencto = "Sem Data"

                registros.append({
                    'CLIENTE': cli,
                    'CEP': f"{cep[:5]}-{cep[5:]}",
                    'ENDEREÇO': end,
                    'QTD': qtd,
                    'MARCA': marca,
                    'VENCTO RIA': vencto
                })
    return pd.DataFrame(registros)


# -----------------------------------------------------------------------------
# CONSTANTES DO FORMULÁRIO DE VISTORIA
# -----------------------------------------------------------------------------
OPCOES_ST = [
    "BO - Bom",
    "AC - Acompanhar",
    "FN - Fora Norma",
    "RU - Ruim",
    "NT - Não tem",
    "NA - Não se aplica"
]

OPCOES_PR = [
    "- (Nenhuma)",
    "OR - Orçar",
    "CA - Enviar carta"
]

ESTRUTURA_VISTORIA = {
    "Casa de Máquinas": {
        "Máquina de Tração / Unidade Hidráulica": [
            "Escova",
            "Freio (ajuste / funcionamento)",
            "Óleo (nível)",
            "Ruído",
            "Vazamento",
            "Limpeza",
            "Polia de tração"
        ],
        "Quadro de Comando": [
            "Isento de jumpers",
            "Relé, bobina, fusível, etc",
            "Limpeza"
        ],
        "Limitador de Velocidade": [
            "Contato elétrico",
            "Limpeza / Lubrificação"
        ]
    },
    "Pavimento": {
        "Portas de Pavimento": [
            "Trinco",
            "Corrediça",
            "Protetores de ilhós com tampa",
            "Limpeza / Lubrificação"
        ],
        "Sinalização": [
            "Funcionamento das botoeiras",
            "Funcionamento setas direção / IPD"
        ]
    },
    "Passadiço": {
        "Operador de Porta": [
            "Contato, roldana, rampa, etc",
            "Limpeza / Lubrificação"
        ],
        "Cabina": [
            "Botoeira de cabina",
            "Iluminação, acrílico, etc",
            "Nivelamento cabina X pavimento",
            "Aparelho de segurança cabina",
            "Luz de emergência"
        ],
        "Guias": [
            "Limpeza / Folgas",
            "Limpeza / Lubrificação"
        ],
        "Cabos de Aço Tração": [
            "Desgaste, equalização",
            "Distância contrapeso para-choque"
        ],
        "Estrutura Contrapeso / Cabina": [
            "Desgaste / Folga corrediças cabina",
            "Desgaste / Folga corrediças contrapeso"
        ]
    },
    "Poço": {
        "Poço & Segurança": [
            "Limpeza do poço",
            "Distância batedor contrapeso e para-choque",
            "Chave de emergência / PAP"
        ],
        "Polia Tensora": [
            "Polia tensora aspecto geral",
            "Contato elétrico",
            "Limpeza / Lubrificação"
        ]
    }
}

CATALOGO_NORMAS = [
    ("001", "Casa de máquinas", "Alçapão"),
    ("014", "Casa de máquinas", "Chave geral do quadro de força"),
    ("015", "Casa de máquinas", "Chave geral identificação nº. Elev."),
    ("016", "Casa de máquinas", "Disjuntor"),
    ("003", "Casa de máquinas", "Extintor de incêndio"),
    ("004", "Casa de máquinas", "Furação passagem de cabos"),
    ("018", "Casa de máquinas", "Fusível"),
    ("005", "Casa de máquinas", "Gancho carga (ausente/ident. Carga)"),
    ("328", "Casa de máquinas", "Guarda-corpo/escada casa maq."),
    ("006", "Casa de máquinas", "Identif. Equip. cab. Quadro maq."),
    ("008", "Casa de máquinas", "Infiltração"),
    ("326", "Casa de máquinas", "Intercomunicador casa máquinas"),
    ("329", "Casa de máquinas", "Luz de emergência casa máquinas"),
    ("347", "Casa de máquinas", "Marcação dos cabos de aço"),
    ("330", "Casa de máquinas", "Placa de aviso para alçapão"),
    ("332", "Casa de máquinas", "Placa de aviso para porta de inspeção"),
    ("331", "Casa de máquinas", "Placa de aviso porta casa máquinas"),
    ("009", "Casa de máquinas", "Porta de acesso casa de máquinas"),
    ("327", "Casa de máquinas", "Porta e fechadura porta casa maq."),
    ("046", "Casa de máquinas", "Protetor de polia de tração"),
    ("133", "Casa de máquinas", "Protetor de polia do limitador"),
    ("020", "Casa de máquinas", "Quadro de força"),
    ("011", "Casa de máquinas", "Ventilação casa de máquinas"),
    ("140", "Cabina", "Botoeira de inspeção"),
    ("166", "Cabina", "Braile para botão de cabina"),
    ("146", "Cabina", "Guarda-corpo"),
    ("170", "Cabina", "Intercomunicador cabina"),
    ("176", "Cabina", "Placa de capacidade"),
    ("175", "Cabina", "Placa de conservação CREL"),
    ("172", "Cabina", "Luz de emergência"),
    ("338", "Passadiço", "Iluminação do passadiço"),
    ("337", "Passadiço", "Kit alça de segurança"),
    ("186", "Passadiço", "Tapa vista"),
    ("228", "Pavimento", "Ilhós de porta de pavimento"),
    ("346", "Poço", "Coletor de óleo"),
    ("241", "Poço", "Contato elétrico polia tensora"),
    ("242", "Poço", "Escada marinheiro"),
    ("149", "Poço", "Lubrificador"),
    ("240", "Poço", "Chave PAP"),
    ("340", "Poço", "Placa aviso verifique elev. Parado pavto."),
    ("343", "Poço", "Placa de aviso chave PAP"),
    ("342", "Poço", "Placa de aviso dist. Contrapeso"),
    ("344", "Poço", "Placa de aviso escada acesso poço"),
    ("345", "Portaria", "Intercomunicador portaria")
]

# -----------------------------------------------------------------------------
# 4. NAVEGAÇÃO PRINCIPAL
# -----------------------------------------------------------------------------
menu = st.sidebar.radio("Navegação Principal", [
    "⚙️ 1. Importador & Roteirizador Logístico",
    "📄 2. Relatórios & Painel do Gestor",
    "📱 3. Terminal de Campo (Vistoriador)"
])

# =============================================================================
# ETAPA 1: IMPORTADOR, FILTROS E ROTEIRIZADOR
# =============================================================================
if menu == "⚙️ 1. Importador & Roteirizador Logístico":
    st.title("⚙️ Gerador Unificado e Roteirizador RIA")
    st.markdown("Importe o relatório (**PDF** ou **Excel**) para gerar jornadas sequenciadas por menor distância.")

    arquivo = st.file_uploader("Carregue o arquivo (PDF ou Excel):", type=["pdf", "xlsx", "xls"])

    if arquivo:
        if arquivo.name.lower().endswith(".pdf"):
            df = extrair_dados_pdf_siga(arquivo.read())
            col_cli, col_cep, col_end = 'CLIENTE', 'CEP', 'ENDEREÇO'
            df['QTD_CLEAN'] = df['QTD']
            df['MARCA_CLEAN'] = df['MARCA'].astype(str).str.upper().str.strip()
            df['VENC_CLEAN'] = df['VENCTO RIA']
        else:
            df_raw = pd.read_excel(arquivo)
            linha_cabecalho = None
            for idx_row, row_data in df_raw.head(15).iterrows():
                valores = [str(v).upper() for v in row_data.values]
                if any("CLIENTE" in v for v in valores) and any("ENDEREÇO" in v or "ENDERECO" in v for v in valores):
                    linha_cabecalho = idx_row
                    break

            if linha_cabecalho is not None:
                df_raw.columns = df_raw.iloc[linha_cabecalho]
                df = df_raw.iloc[linha_cabecalho + 1:].copy()
            else:
                df = df_raw.copy()

            df.columns = [str(c).strip().upper() for c in df.columns]
            col_cli = next((c for c in df.columns if "CLIENTE" in c), None)
            col_cep = next((c for c in df.columns if "CEP" in c), None)
            col_end = next((c for c in df.columns if "ENDEREÇO" in c or "ENDERECO" in c), None)
            col_qtd = next((c for c in df.columns if "QTDE" in c or "QTD" in c), None)
            col_marca = next((c for c in df.columns if "MARCA" in c), None)
            col_venc = next((c for c in df.columns if "MÊS" in c or "MES" in c or "VENCTO" in c or "VENC" in c), None)

            if not (col_cli and col_end):
                st.error("Não foi possível identificar as colunas de CLIENTE e ENDEREÇO.")
                st.stop()

            df = df[df[col_cli].notna() & (df[col_cli].astype(str).str.strip() != '')].copy()
            df['QTD_CLEAN'] = pd.to_numeric(df[col_qtd], errors='coerce').fillna(1).astype(int)
            df['MARCA_CLEAN'] = df[col_marca].astype(str).str.strip().str.upper() if col_marca else 'DIVERSAS'
            df['VENC_CLEAN'] = df[col_venc].astype(str).str.strip() if col_venc else 'Sem Data'

        df_geo = carregar_base_georreferenciada()
        df['CLI_NORM'] = df[col_cli].apply(normalizar_codigo)

        if not df_geo.empty:
            df = df.merge(df_geo[['CLI_NORM', 'lat', 'lon']], on='CLI_NORM', how='left')
        else:
            df['lat'] = None
            df['lon'] = None

        st.subheader("🎯 Filtros Operacionais")
        f1, f2 = st.columns(2)
        with f1:
            excluir_alfa = st.toggle("Excluir Marca ALFA", value=True)
        with f2:
            modo_venc = st.selectbox(
                "Escopo de Vencimento:",
                ["Todos (Vencimentos + Sem Data / Pendências)", "Apenas Vencimentos com Data",
                 "Apenas Sem Data / Pendências"]
            )

        df_f = df.copy()
        if excluir_alfa:
            df_f = df_f[~df_f['MARCA_CLEAN'].str.contains('ALFA', na=False)]


        def classificar_grupo(v):
            v_str = str(v).lower()
            if "sem data" in v_str or "//" in v_str or not re.search(r'\d{2}/\d{2}/\d{4}', v_str):
                return "Sem Data / Pendência"
            return "Vencimento Programado"


        df_f['GRUPO_STATUS'] = df_f['VENC_CLEAN'].apply(classificar_grupo)

        if modo_venc == "Apenas Vencimentos com Data":
            df_f = df_f[df_f['GRUPO_STATUS'] == "Vencimento Programado"]
        elif modo_venc == "Apenas Sem Data / Pendências":
            df_f = df_f[df_f['GRUPO_STATUS'] == "Sem Data / Pendência"]

        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Locais Filtrados", len(df_f))
        m2.metric("Com Vencimento", len(df_f[df_f['GRUPO_STATUS'] == 'Vencimento Programado']))
        m3.metric("Sem Data / Pendências", len(df_f[df_f['GRUPO_STATUS'] == 'Sem Data / Pendência']))
        m4.metric("Total de Elevadores", int(df_f['QTD_CLEAN'].sum()))

        st.markdown("---")
        st.subheader("☑️ Seleção Manual de Endereços")
        df_edit = df_f[[col_cli, col_cep, col_end, 'QTD_CLEAN', 'MARCA_CLEAN', 'VENC_CLEAN', 'GRUPO_STATUS']].copy()
        df_edit.insert(0, "Incluir", True)
        df_edit.columns = ["Incluir", "CLIENTE", "CEP", "ENDEREÇO", "QTD", "MARCA", "VENCTO RIA", "GRUPO"]

        grid_selecao = st.data_editor(df_edit, use_container_width=True, hide_index=True)
        selecionados = grid_selecao[grid_selecao["Incluir"] == True].copy()

        st.markdown("---")
        st.subheader("🚚 Configuração das Jornadas Diárias")
        c1, c2 = st.columns(2)
        with c1:
            max_locais_dia = st.slider("Máximo de Locais por Dia:", 3, 15, 7)
        with c2:
            max_elev_dia = st.slider("Máximo de Elevadores por Dia:", 5, 40, 14)

        if st.button("🚀 Otimizar Trajetos e Salvar para Campo", type="primary", use_container_width=True):
            if selecionados.empty:
                st.warning("Nenhum local selecionado.")
            else:
                with st.spinner("Otimizando trajetos por menor distância..."):
                    dados_processados = []
                    for _, r_sel in selecionados.iterrows():
                        norm = normalizar_codigo(r_sel['CLIENTE'])
                        row_orig = df_f[df_f['CLI_NORM'] == norm].iloc[0]
                        p_lat = row_orig['lat']
                        p_lon = row_orig['lon']

                        if pd.isna(p_lat) or p_lat is None:
                            p_lat, p_lon = buscar_coords_fallback(r_sel['CEP'], r_sel['ENDEREÇO'])

                        waze = f"https://waze.com/ul?ll={p_lat},{p_lon}&navigate=yes"
                        dados_processados.append({
                            'cliente': r_sel['CLIENTE'],
                            'cep': str(r_sel['CEP'] or ''),
                            'endereco': r_sel['ENDEREÇO'],
                            'qtd_elev': r_sel['QTD'],
                            'marca': r_sel['MARCA'],
                            'vencto_ria': r_sel['VENCTO RIA'],
                            'grupo_status': r_sel['GRUPO'],
                            'lat': float(p_lat),
                            'lon': float(p_lon),
                            'waze_url': waze
                        })

                    rotas_finais = []
                    pontos_restantes = dados_processados.copy()
                    dia_num = 1

                    while pontos_restantes:
                        dia_nome = f"Dia {dia_num}"
                        parada_num = 1
                        elev_acum = 0
                        locais_acum = 0

                        atual = pontos_restantes.pop(0)
                        atual['dia_rota'] = dia_nome
                        atual['ordem_parada'] = parada_num
                        elev_acum += atual['qtd_elev']
                        locais_acum += 1
                        rotas_finais.append(atual)

                        while pontos_restantes:
                            candidatos = [
                                p for p in pontos_restantes
                                if (locais_acum + 1 <= max_locais_dia) and (elev_acum + p['qtd_elev'] <= max_elev_dia)
                            ]

                            if not candidatos:
                                break

                            proximo = min(
                                candidatos,
                                key=lambda p_cand: calc_dist_km(atual['lat'], atual['lon'], p_cand['lat'],
                                                                p_cand['lon']),
                            )
                            pontos_restantes.remove(proximo)
                            parada_num += 1
                            locais_acum += 1
                            elev_acum += proximo['qtd_elev']
                            proximo['dia_rota'] = dia_nome
                            proximo['ordem_parada'] = parada_num
                            rotas_finais.append(proximo)
                            atual = proximo
                        dia_num += 1

                    conexao_db = get_db()
                    conexao_db.execute("DELETE FROM rotas_ativas")
                    for p_item in rotas_finais:
                        uid = f"{p_item['dia_rota']}_{p_item['ordem_parada']}_{p_item['cliente']}".replace(" ", "_")
                        conexao_db.execute('''
                            INSERT INTO rotas_ativas 
                            (id, mes_exec, dia_rota, ordem_parada, cliente, cep, endereco, qtd_elev, marca, vencto_ria, grupo_status, lat, lon, waze_url, status_vistoria)
                            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'Pendente')
                        ''', (uid, p_item['vencto_ria'], p_item['dia_rota'], p_item['ordem_parada'], p_item['cliente'],
                              p_item['cep'],
                              p_item['endereco'], p_item['qtd_elev'], p_item['marca'], p_item['vencto_ria'],
                              p_item['grupo_status'], p_item['lat'],
                              p_item['lon'], p_item['waze_url']))
                    conexao_db.commit()
                    conexao_db.close()

                    st.success(
                        f"✅ Roteirização concluída! {len(rotas_finais)} locais distribuídos em {dia_num - 1} jornadas diárias.")

# =============================================================================
# ETAPA 2: RELATÓRIOS & PAINEL DO GESTOR (COM IMPRESSÃO / SALVAR EM PDF)
# =============================================================================
elif menu == "📄 2. Relatórios & Painel do Gestor":
    st.title("📄 Painel do Gestor — Roteiros & Relatórios de Vistoria")

    st.markdown("""
    <style>
    @media print {
        header, footer, [data-testid="stSidebar"], [data-testid="stHeader"], 
        .stButton, [data-testid="stDownloadButton"], iframe, .no-print {
            display: none !important;
        }
        html, body, [data-testid="stAppViewContainer"], [data-testid="stMainBlockContainer"], section.main {
            overflow: visible !important;
            height: auto !important;
            position: static !important;
            padding: 0 !important;
            margin: 0 !important;
        }
        @page {
            size: A4 portrait;
            margin: 10mm;
        }
        table {
            width: 100% !important;
            border-collapse: collapse !important;
            page-break-inside: auto !important;
        }
        tr {
            page-break-inside: avoid !important;
        }
        th, td {
            border: 1px solid #777 !important;
            padding: 5px 7px !important;
            font-size: 11px !important;
        }
        th {
            background-color: #1B365D !important;
            color: white !important;
            -webkit-print-color-adjust: exact !important;
            print-color-adjust: exact !important;
        }
    }
    .tabela-rel-doc {
        width: 100%;
        border-collapse: collapse;
        margin-top: 10px;
        margin-bottom: 20px;
        font-family: Arial, sans-serif;
    }
    .tabela-rel-doc th, .tabela-rel-doc td {
        border: 1px solid #ccc;
        padding: 6px 8px;
        font-size: 12px;
    }
    .tabela-rel-doc th {
        background-color: #1B365D;
        color: white;
        text-align: left;
    }
    .tabela-rel-doc tr:nth-child(even) {
        background-color: #f8f9fa;
    }
    </style>
    """, unsafe_allow_html=True)

    conexao_gestor = get_db()
    df_rel = pd.read_sql("SELECT * FROM rotas_ativas", conexao_gestor)
    df_vistorias_db = pd.read_sql("SELECT * FROM relatorios_vistorias", conexao_gestor)
    conexao_gestor.close()

    tab_roteiro, tab_vistorias_salvas = st.tabs(
        ["🗺️ Roteiro Operacional (Impressão/Excel)", "📥 Relatórios Técnicos Concluídos"])

    with tab_roteiro:
        if not df_rel.empty:
            df_rel['dia_num'] = df_rel['dia_rota'].astype(str).str.extract(r'(\d+)').fillna(0).astype(int)
            df_rel = df_rel.sort_values(by=['dia_num', 'ordem_parada']).drop(columns=['dia_num']).reset_index(drop=True)

        if df_rel.empty:
            st.info("Nenhuma rota cadastrada no momento. Processe o arquivo na Etapa 1.")
        else:
            c1, c2, c3 = st.columns(3)
            c1.metric("Total de Endereços", len(df_rel))
            c2.metric("Total de Elevadores", int(df_rel['qtd_elev'].sum()))
            c3.metric("Dias Programados", df_rel['dia_rota'].nunique())

            col1, col2 = st.columns(2)
            with col1:
                buffer = pd.ExcelWriter("Relatorio_Consolidado_RIA.xlsx", engine='openpyxl')
                df_rel.to_excel(buffer, index=False, sheet_name="Roteiro_Geral")
                buffer.close()
                with open("Relatorio_Consolidado_RIA.xlsx", "rb") as f_excel:
                    st.download_button("📥 Baixar Roteiro em Excel", f_excel, "Relatorio_Consolidado_RIA.xlsx",
                                       use_container_width=True)

            with col2:
                components.html("""
                <button onclick="window.parent.print()" style="
                    width: 100%; height: 42px; background-color: #1B365D; color: #ffffff;
                    border: none; border-radius: 8px; font-size: 15px; font-weight: 600; cursor: pointer;
                ">🖨️ Imprimir / Salvar em PDF (Página Inteira)</button>
                """, height=50)


            def renderizar_tabela_html(df_sub):
                linhas = []
                for _, r_item in df_sub.iterrows():
                    linhas.append(
                        f"<tr>"
                        f"<td style='text-align:center; font-weight:bold;'>{r_item['dia_rota']}</td>"
                        f"<td style='text-align:center;'>#{r_item['ordem_parada']}</td>"
                        f"<td><b>{r_item['cliente']}</b></td>"
                        f"<td>{r_item['cep']}</td>"
                        f"<td>{r_item['endereco']}</td>"
                        f"<td style='text-align:center;'>{r_item['qtd_elev']}</td>"
                        f"<td>{r_item['marca']}</td>"
                        f"<td style='text-align:center;'>{r_item['vencto_ria']}</td>"
                        f"<td style='text-align:center;'>{r_item['status_vistoria']}</td>"
                        f"</tr>"
                    )
                corpo = "".join(linhas)
                return (
                    "<table style='width:100%; border-collapse:collapse; margin-bottom:20px; font-family:Arial; font-size:12px;'>"
                    "<thead><tr style='background-color:#1B365D; color:white;'>"
                    "<th style='padding:6px; border:1px solid #ddd;'>Dia</th>"
                    "<th style='padding:6px; border:1px solid #ddd;'>Parada</th>"
                    "<th style='padding:6px; border:1px solid #ddd;'>Cliente</th>"
                    "<th style='padding:6px; border:1px solid #ddd;'>CEP</th>"
                    "<th style='padding:6px; border:1px solid #ddd;'>Endereço</th>"
                    "<th style='padding:6px; border:1px solid #ddd;'>Qtd</th>"
                    "<th style='padding:6px; border:1px solid #ddd;'>Marca</th>"
                    "<th style='padding:6px; border:1px solid #ddd;'>Vencimento</th>"
                    "<th style='padding:6px; border:1px solid #ddd;'>Status</th>"
                    "</tr></thead>"
                    f"<tbody>{corpo}</tbody>"
                    "</table>"
                )


            df_out = df_rel[df_rel['grupo_status'] == 'Vencimento Programado']
            if not df_out.empty:
                st.markdown(f"#### 1. Vistorias com Vencimento Programado ({len(df_out)} locais)")
                st.html(renderizar_tabela_html(df_out))

            df_sem = df_rel[df_rel['grupo_status'] != 'Vencimento Programado']
            if not df_sem.empty:
                st.markdown(f"#### 2. Endereços sem Data Cadastrada / Pendências ({len(df_sem)} locais)")
                st.html(renderizar_tabela_html(df_sem))

    with tab_vistorias_salvas:
        st.subheader("📁 Vistorias Técnicas Concluídas")
        if df_vistorias_db.empty:
            st.info("Nenhuma vistoria técnica finalizada até o momento.")
        else:
            st.dataframe(df_vistorias_db[['cliente', 'endereco', 'data_vistoria', 'vistoriador', 'qtd_normas_orc',
                                          'qtd_pendencias_sup']], use_container_width=True)
            st.markdown("---")

            opcoes_vistorias = {
                f"{r['cliente']} - {r['data_vistoria']} ({r['id_rota']})": r['caminho_arquivo']
                for _, r in df_vistorias_db.iterrows()
            }

            escolha_rel = st.selectbox("Selecione a Vistoria para Visualizar, Imprimir ou Salvar em PDF:",
                                       list(opcoes_vistorias.keys()))
            caminho_escolhido = opcoes_vistorias[escolha_rel]

            if os.path.exists(caminho_escolhido):
                col_btn_print, col_btn_xls = st.columns(2)
                with col_btn_print:
                    components.html("""
                    <button onclick="window.parent.print()" style="
                        width: 100%; height: 42px; background-color: #1B365D; color: #ffffff;
                        border: none; border-radius: 8px; font-size: 15px; font-weight: 600; cursor: pointer;
                        box-shadow: 0 2px 4px rgba(0,0,0,0.15);
                    ">🖨️ Imprimir / Salvar Relatório em PDF</button>
                    """, height=50)

                with col_btn_xls:
                    with open(caminho_escolhido, "rb") as f_down:
                        st.download_button(
                            label="📥 Baixar Pasta Excel (.xlsx)",
                            data=f_down,
                            file_name=os.path.basename(caminho_escolhido),
                            use_container_width=True
                        )

                xls = pd.ExcelFile(caminho_escolhido)
                df_id_elev = pd.read_excel(xls, sheet_name="Elevadores_Identificacao")
                df_orc_normas = pd.read_excel(xls, sheet_name="Orcamento_Normas")
                df_prev_matriz = pd.read_excel(xls, sheet_name="Manutencao_Preventiva")
                df_sup_pend = pd.read_excel(xls, sheet_name="Supervisao_Pendencias")

                st.markdown(f"""
                <div style="border-bottom: 2px solid #1B365D; padding-bottom: 8px; margin-top: 20px; margin-bottom: 15px;">
                    <div style="display:flex; justify-content:space-between; align-items:center;">
                        <h2 style="color:#1B365D; margin:0;">RELATÓRIO TÉCNICO DE VISTORIA — RIA</h2>
                        <span style="font-weight:bold; color:#555;">CREL ELEVADORES</span>
                    </div>
                    <p style="margin:5px 0 0 0; color:#333; font-size:14px;">
                        <b>Edifício:</b> {escolha_rel.split(' - ')[0]} | <b>Data:</b> {escolha_rel.split(' - ')[1].split(' (')[0]}
                    </p>
                </div>
                """, unsafe_allow_html=True)

                st.markdown("##### 1. Identificação dos Elevadores")
                st.html(df_id_elev.to_html(classes="tabela-rel-doc", index=False))

                st.markdown("##### 2. Itens de Norma Técnica para Orçamento Comercial")
                if not df_orc_normas.empty:
                    st.html(df_orc_normas.to_html(classes="tabela-rel-doc", index=False))
                else:
                    st.info("Nenhuma não-conformidade de norma técnica que demande orçamento.")

                st.markdown("##### 3. Avaliação de Manutenção Preventiva")
                st.html(df_prev_matriz.to_html(classes="tabela-rel-doc", index=False))

                st.markdown("##### 4. Apontamentos e Pendências para Supervisão Técnica")
                if not df_sup_pend.empty:
                    st.html(df_sup_pend.to_html(classes="tabela-rel-doc", index=False))
                else:
                    st.info("Nenhuma pendência ou anomalia operacional apontada.")

# =============================================================================
# ETAPA 3: TERMINAL DE CAMPO & FORMULÁRIO DE VISTORIA DINÂMICO
# =============================================================================
elif menu == "📱 3. Terminal de Campo (Vistoriador)":
    conexao_campo = get_db()
    df_todos = pd.read_sql("SELECT * FROM rotas_ativas", conexao_campo)
    conexao_campo.close()

    if df_todos.empty:
        st.warning("Nenhuma vistoria atribuída no momento.")
        st.stop()

    if "vistoria_ativa_id" not in st.session_state:
        st.session_state["vistoria_ativa_id"] = None

    if st.session_state["vistoria_ativa_id"]:
        id_ativa = st.session_state["vistoria_ativa_id"]
        registro_cliente = df_todos[df_todos['id'] == id_ativa]

        if registro_cliente.empty:
            st.session_state["vistoria_ativa_id"] = None
            st.rerun()

        dados_cli = registro_cliente.iloc[0]

        st.button("⬅ Voltar para o Roteiro do Dia",
                  on_click=lambda: st.session_state.update({"vistoria_ativa_id": None}))

        st.markdown(f"""
        <div style="background:#1B365D; color:white; padding:12px; border-radius:8px; margin-bottom:15px;">
            <h3 style="margin:0; color:white;">🛗 Vistoria Técnica: {dados_cli['cliente']}</h3>
            <p style="margin:4px 0 0 0; font-size:13px; opacity:0.9;">
                📍 {dados_cli['endereco']} | Marca: {dados_cli['marca']} | Qtd Cadastrada: {dados_cli['qtd_elev']}
            </p>
        </div>
        """, unsafe_allow_html=True)

        with st.expander("📋 1. Identificação dos Elevadores (CREL & PMSP)", expanded=True):
            qtd_padrao = max(int(dados_cli['qtd_elev'] or 1), 1)
            qtd_carros = st.number_input(
                "Quantidade Total de Elevadores no Edifício:",
                min_value=1,
                max_value=30,
                value=qtd_padrao,
                step=1,
                key=f"num_elev_{id_ativa}"
            )

            # Lista ordenada de todos os carros: C1, C2, ..., Cn
            todos_carros = [f"C{i + 1}" for i in range(int(qtd_carros))]

            dados_elevadores_cad = []
            colunas_por_linha = 4 if len(todos_carros) >= 4 else len(todos_carros)
            cols_cad = st.columns(colunas_por_linha)

            for i, c_id in enumerate(todos_carros):
                with cols_cad[i % colunas_por_linha]:
                    st.markdown(f"**Elevador {c_id}**")
                    crel_val = st.text_input(f"Nº CREL ({c_id}):", key=f"crel_{id_ativa}_{c_id}")
                    pmsp_val = st.text_input(f"Nº PMSP ({c_id}):", key=f"pmsp_{id_ativa}_{c_id}")
                    dados_elevadores_cad.append({"Elevador": c_id, "Nº CREL": crel_val, "Nº PMSP": pmsp_val})

            col_tec, col_sup = st.columns(2)
            with col_tec:
                nome_vistoriador = st.text_input("Técnico Vistoriador:", value=dados_cli['vistoriador'] or "",
                                                 key=f"tec_{id_ativa}")
            with col_sup:
                _nome_supervisor = st.text_input("Supervisor Técnico:", key=f"sup_{id_ativa}")

        st.markdown("### 📋 2. Avaliação dos Componentes")
        st.caption(
            "Legenda ST: BO (Bom) | AC (Acompanhar) | FN (Fora Norma) | RU (Ruim) | NT (Não tem) | NA (Não se aplica)")
        st.caption("Legenda PR: - (Nenhuma) | OR (Orçar) | CA (Enviar carta)")

        abas_locais = st.tabs(list(ESTRUTURA_VISTORIA.keys()))
        itens_avaliados = []

        for idx_local, (local_nome, subitens) in enumerate(ESTRUTURA_VISTORIA.items()):
            with abas_locais[idx_local]:
                for item_nome, pontos in subitens.items():
                    with st.expander(f"⚙️ {item_nome}", expanded=True):
                        for ponto in pontos:
                            st.markdown(f"🔹 **{ponto}**")
                            for c_id in todos_carros:
                                col_st, col_pr, col_obs = st.columns([1, 1, 2])

                                k_st = f"st_{id_ativa}_{local_nome}_{item_nome}_{ponto}_{c_id}"
                                k_pr = f"pr_{id_ativa}_{local_nome}_{item_nome}_{ponto}_{c_id}"
                                k_ob = f"ob_{id_ativa}_{local_nome}_{item_nome}_{ponto}_{c_id}"

                                with col_st:
                                    s_val = st.selectbox(f"ST ({c_id}):", OPCOES_ST, index=0, key=k_st)
                                with col_pr:
                                    p_val = st.selectbox(f"PR ({c_id}):", OPCOES_PR, index=0, key=k_pr)
                                with col_obs:
                                    o_val = st.text_input(f"Obs ({c_id}):", key=k_ob, placeholder=f"Obs {c_id}")

                                itens_avaliados.append({
                                    "Local": local_nome,
                                    "Item": item_nome,
                                    "Ponto de Avaliação": ponto,
                                    "Elevador": c_id,
                                    "Situação (ST)": s_val.split(" - ")[0],
                                    "Providência (PR)": p_val.split(" - ")[0].replace("- (Nenhuma)", "-"),
                                    "Observações": o_val.strip()
                                })
                            st.markdown("<hr style='margin: 4px 0 8px 0;'>", unsafe_allow_html=True)

        with st.expander("📑 3. Itens de Adequação de Norma Técnica (Orçamento)", expanded=False):
            opcoes_n = [f"{cod} - {loc}: {desc}" for cod, loc, desc in CATALOGO_NORMAS]
            normas_sel = st.multiselect("Selecione itens de norma para gerar proposta de orçamento:", opcoes_n,
                                        key=f"normas_sel_{id_ativa}")

            lista_normas_avulsas = []
            for n_str in normas_sel:
                n_cod = n_str.split(" - ")[0]
                n_loc = n_str.split(" - ")[1].split(":")[0]
                n_desc = n_str.split(": ")[1]

                st.write(f"**Item {n_cod}: {n_desc}**")
                carros_af = st.multiselect(
                    f"Carros afetados ({n_cod}):",
                    todos_carros,
                    default=todos_carros,
                    key=f"c_af_{id_ativa}_{n_cod}"
                )
                c_st, c_pr = st.columns(2)
                with c_st:
                    st_n = st.selectbox(f"Situação ({n_cod}):", OPCOES_ST, index=2, key=f"st_n_{id_ativa}_{n_cod}")
                with c_pr:
                    pr_n = st.selectbox(f"Providência ({n_cod}):", OPCOES_PR, index=1, key=f"pr_n_{id_ativa}_{n_cod}")

                lista_normas_avulsas.append({
                    "Endereço": dados_cli['endereco'],
                    "Equipamento(s)": ", ".join(carros_af),
                    "Código Norma": n_cod,
                    "Local": n_loc,
                    "Descrição do Item": n_desc,
                    "Situação (ST)": st_n.split(" - ")[0],
                    "Providência (PR)": pr_n.split(" - ")[0],
                    "Destino": "Orçamento de Adequação de Norma"
                })

        with st.expander("🚨 4. Parecer Geral da Supervisão", expanded=False):
            parecer_sup = st.text_area("Comentários e Orientações para a Equipe de Manutenção:",
                                       key=f"parecer_{id_ativa}")

        st.markdown("---")
        if st.button("💾 Finalizar Vistoria e Salvar Relatório", type="primary", use_container_width=True):
            df_av = pd.DataFrame(itens_avaliados)

            df_norm_rotina = df_av[(df_av["Providência (PR)"] == "OR") | (df_av["Situação (ST)"] == "FN")].copy()
            df_norm_rotina["Endereço"] = dados_cli['endereco']
            df_norm_rotina["Destino"] = "Orçamento Comercial"
            df_norm_final = pd.concat([df_norm_rotina, pd.DataFrame(lista_normas_avulsas)], ignore_index=True)

            # 2. Relatório de Manutenção Preventiva (Matriz Dinâmica com Ordenação Real de C1 até Cn)
            df_prev_pivot = df_av.pivot_table(
                index=["Local", "Item", "Ponto de Avaliação"],
                columns="Elevador",
                values="Situação (ST)",
                aggfunc="first"
            ).reset_index()

            colunas_ordenadas = ["Local", "Item", "Ponto de Avaliação"] + [c for c in todos_carros if
                                                                           c in df_prev_pivot.columns]
            df_prev_pivot = df_prev_pivot.reindex(columns=colunas_ordenadas)

            df_sup = df_av[
                (df_av["Situação (ST)"].isin(["RU", "AC"])) |
                (df_av["Observações"] != "") |
                (df_av["Providência (PR)"] == "CA")
                ].copy()
            df_sup["Endereço"] = dados_cli['endereco']
            df_sup["Prioridade"] = df_sup["Situação (ST)"].apply(lambda s: "Alta" if s == "RU" else "Média")

            if parecer_sup.strip():
                linha_parecer = pd.DataFrame([{
                    "Local": "Geral", "Item": "Supervisão", "Ponto de Avaliação": "Parecer do Supervisor",
                    "Elevador": "Geral", "Situação (ST)": "RU", "Providência (PR)": "-",
                    "Observações": parecer_sup.strip(), "Endereço": dados_cli['endereco'], "Prioridade": "Alta"
                }])
                df_sup = pd.concat([linha_parecer, df_sup], ignore_index=True)

            dt_hoje = datetime.datetime.now().strftime("%Y%m%d_%H%M")
            nome_arquivo_excel = os.path.join(DIR_RELATORIOS,
                                              f"Vistoria_{normalizar_codigo(dados_cli['cliente'])}_{dt_hoje}.xlsx")
            with pd.ExcelWriter(nome_arquivo_excel, engine="openpyxl") as writer:
                pd.DataFrame(dados_elevadores_cad).to_excel(writer, sheet_name="Elevadores_Identificacao", index=False)
                df_norm_final.to_excel(writer, sheet_name="Orcamento_Normas", index=False)
                df_prev_pivot.to_excel(writer, sheet_name="Manutencao_Preventiva", index=False)
                df_sup.to_excel(writer, sheet_name="Supervisao_Pendencias", index=False)

            data_str = datetime.datetime.now().strftime("%d/%m/%Y %H:%M")
            conexao_salvar = get_db()
            conexao_salvar.execute('''
                UPDATE rotas_ativas 
                SET status_vistoria = 'Concluído', data_baixa = ?, vistoriador = ?, obs = ?
                WHERE id = ?
            ''', (data_str, nome_vistoriador, "Vistoria técnica preenchida via aplicativo.", id_ativa))

            conexao_salvar.execute('''
                INSERT OR REPLACE INTO relatorios_vistorias
                (id_rota, cliente, endereco, data_vistoria, vistoriador, caminho_arquivo, qtd_normas_orc, qtd_pendencias_sup)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            ''', (id_ativa, dados_cli['cliente'], dados_cli['endereco'], data_str, nome_vistoriador, nome_arquivo_excel,
                  len(df_norm_final), len(df_sup)))
            conexao_salvar.commit()
            conexao_salvar.close()

            st.success(f"✅ Vistoria concluída e salva com sucesso em: {nome_arquivo_excel}")
            st.session_state["vistoria_ativa_id"] = None
            st.rerun()

    else:
        st.title("📱 Roteiro de Campo - Vistoriador")

        dias_disp = sorted(df_todos['dia_rota'].unique(), key=lambda x: int(re.sub(r'\D', '', x) or 0))
        dia_sel = st.selectbox("Selecione a Jornada / Dia:", dias_disp)

        df_dia = df_todos[df_todos['dia_rota'] == dia_sel].sort_values("ordem_parada").copy()

        col_gps1, col_gps2 = st.columns([1, 2])
        with col_gps1:
            loc = get_geolocation()

        gps_ativo = False
        lat_user, lon_user = None, None
        if loc and 'coords' in loc and loc['coords']:
            lat_user = loc['coords']['latitude']
            lon_user = loc['coords']['longitude']
            gps_ativo = True
            with col_gps2:
                st.success(f"📍 GPS Ativo ({lat_user:.4f}, {lon_user:.4f})")

        with st.expander("🗺️ Ver Mapa do Dia", expanded=True):
            centro_lat = float(df_dia['lat'].mean())
            centro_lon = float(df_dia['lon'].mean())
            mapa = folium.Map(location=[centro_lat, centro_lon], zoom_start=13, tiles="OpenStreetMap")

            coords = []
            if gps_ativo and lat_user and lon_user:
                coords.append([lat_user, lon_user])
                folium.Marker([lat_user, lon_user], popup="Sua Posição", tooltip="Você está aqui",
                              icon=folium.Icon(color="purple", icon="user")).add_to(mapa)

            for _, r_mapa in df_dia.iterrows():
                p_coord = [float(r_mapa['lat']), float(r_mapa['lon'])]
                coords.append(p_coord)
                cor = "green" if r_mapa['status_vistoria'] == "Concluído" else (
                    "red" if r_mapa['status_vistoria'] == "Improdutiva" else "blue")
                folium.Marker(p_coord, popup=f"#{r_mapa['ordem_parada']} - {r_mapa['cliente']}<br>{r_mapa['endereco']}",
                              tooltip=f"#{r_mapa['ordem_parada']} {r_mapa['cliente']}",
                              icon=folium.Icon(color=cor)).add_to(mapa)

            if len(coords) > 1:
                folium.PolyLine(coords, color="#1B365D", weight=4, dash_array="6, 8").add_to(mapa)

            st_folium(mapa, width="100%", height=350, returned_objects=[])

        st.markdown("---")
        for _, row_parada in df_dia.iterrows():
            st_cor = "#28a745" if row_parada['status_vistoria'] == "Concluído" else (
                "#dc3545" if row_parada['status_vistoria'] == "Improdutiva" else "#1B365D")
            st.markdown(f"""
            <div style="background:#fff; border-left:6px solid {st_cor}; padding:12px; border-radius:8px; margin-bottom:12px; box-shadow:0 2px 5px rgba(0,0,0,0.08);">
                <div style="display:flex; justify-content:space-between;">
                    <b>PARADA #{row_parada['ordem_parada']} — {row_parada['cliente']}</b>
                    <span style="font-size:12px; font-weight:bold; color:{st_cor};">{row_parada['status_vistoria']}</span>
                </div>
                <div style="font-size:13px; color:#555; margin-top:4px;">
                    📍 {row_parada['endereco']} | CEP: {row_parada['cep']}<br>
                    🛗 {row_parada['qtd_elev']} elevador(es) — Marca: {row_parada['marca']}
                </div>
            </div>
            """, unsafe_allow_html=True)

            b1, b2, b3 = st.columns([1, 1, 1])
            with b1:
                st.link_button("🚗 Waze", row_parada['waze_url'], use_container_width=True)
            with b2:
                if st.button("📝 Fazer Vistoria", key=f"vistoria_btn_{row_parada['id']}", type="primary",
                             use_container_width=True):
                    st.session_state["vistoria_ativa_id"] = row_parada['id']
                    st.rerun()
            with b3:
                with st.expander("Baixa Rápida"):
                    with st.form(f"baixa_{row_parada['id']}"):
                        n_status = st.selectbox("Status:", ["Concluído", "Improdutiva", "Pendente"], index=0)
                        tec_resp = st.text_input("Vistoriador:", value=row_parada['vistoriador'] or "")
                        obs_baixa = st.text_area("Observações:", value=row_parada['obs'] or "")
                        if st.form_submit_button("Salvar Baixa", use_container_width=True):
                            dt_agora = datetime.datetime.now().strftime("%d/%m/%Y %H:%M")
                            conexao_bx = get_db()
                            conexao_bx.execute(
                                "UPDATE rotas_ativas SET status_vistoria = ?, data_baixa = ?, vistoriador = ?, obs = ? WHERE id = ?",
                                (n_status, dt_agora, tec_resp, obs_baixa, row_parada['id']))
                            conexao_bx.commit()
                            conexao_bx.close()
                            st.success("Baixa realizada!")
                            st.rerun()