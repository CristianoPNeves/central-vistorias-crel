import datetime
import io
import math
import os
import re
import unicodedata
import urllib.parse
import base64
import folium
import pandas as pd
import pypdf
import requests
import sqlalchemy
from sqlalchemy import text
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
# 2. BANCO DE DADOS HÍBRIDO (SUPABASE NA NUVEM / SQLITE LOCAL)
# -----------------------------------------------------------------------------
def get_engine():
    db_url = None
    try:
        if hasattr(st, "secrets") and "DATABASE_URL" in st.secrets:
            db_url = st.secrets["DATABASE_URL"]
    except Exception:
        db_url = None

    if not db_url:
        db_url = f"sqlite:///{DB_NAME}"

    if db_url.startswith("postgres://"):
        db_url = db_url.replace("postgres://", "postgresql+psycopg://", 1)
    elif db_url.startswith("postgresql://") and not db_url.startswith("postgresql+"):
        db_url = db_url.replace("postgresql://", "postgresql+psycopg://", 1)

    if "sqlite" in db_url:
        return sqlalchemy.create_engine(db_url)

    return sqlalchemy.create_engine(
        db_url,
        connect_args={"prepare_threshold": None},
        pool_pre_ping=True
    )


def inicializar_banco():
    engine = get_engine()
    with engine.begin() as conn:
        conn.execute(text('''
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
        '''))
        conn.execute(text('''
            CREATE TABLE IF NOT EXISTS relatorios_vistorias (
                id_rota TEXT PRIMARY KEY,
                cliente TEXT,
                endereco TEXT,
                data_vistoria TEXT,
                vistoriador TEXT,
                caminho_arquivo TEXT,
                qtd_normas_orc INTEGER,
                qtd_pendencias_sup INTEGER,
                arquivo_binario TEXT
            )
        '''))
        try:
            conn.execute(text("ALTER TABLE relatorios_vistorias ADD COLUMN IF NOT EXISTS arquivo_binario TEXT"))
        except Exception:
            pass


inicializar_banco()


def normalizar_codigo(texto):
    t = str(texto or '').upper().replace(' ', '').replace('_', '').replace('-', '')
    return ''.join(c for c in unicodedata.normalize('NFD', t) if unicodedata.category(c) != 'Mn')


@st.cache_data
def carregar_base_georreferenciada():
    if os.path.exists(ARQUIVO_GEO_BASE):
        try:
            df_geo_local = pd.read_excel(ARQUIVO_GEO_BASE)
            col_cli = next((c for c in df_geo_local.columns if "CLIENTE" in str(c).upper()), None)
            if col_cli and 'lat' in df_geo_local.columns and 'lon' in df_geo_local.columns:
                df_geo_local['CLI_NORM'] = df_geo_local[col_cli].apply(normalizar_codigo)
                return df_geo_local[['CLI_NORM', 'lat', 'lon']].dropna().drop_duplicates(subset=['CLI_NORM'])
        except Exception:
            pass
    return pd.DataFrame()


def buscar_coords_fallback(cep, endereco_padrao=None):
    cep_limpo = re.sub(r'\D', '', str(cep or ''))

    if len(cep_limpo) == 8:
        try:
            resp = requests.get(f"https://cep.awesomeapi.com.br/json/{cep_limpo}", timeout=3)
            if resp.status_code == 200:
                dados_cep = resp.json()
                lat = float(dados_cep.get('lat', 0))
                lng = float(dados_cep.get('lng', 0))
                if lat != 0 and lng != 0 and not (round(lat, 3) == -23.550 and round(lng, 3) == -46.633):
                    return lat, lng
        except Exception:
            pass

    if endereco_padrao and str(endereco_padrao).strip():
        try:
            headers = {'User-Agent': 'GestaoRIA_CREL_Logistica/3.0'}
            end_limpo = re.split(r'[-–,/]|\bapt\b|\bapto\b|\bbloco\b|\bconj\b|\bsala\b|\bandar\b',
                                 str(endereco_padrao), flags=re.IGNORECASE)[0].strip()
            consulta = f"{end_limpo}, São Paulo, SP, Brasil"
            url_nom = f"https://nominatim.openstreetmap.org/search?format=json&q={urllib.parse.quote(consulta)}&limit=1"
            resp_nom = requests.get(url_nom, headers=headers, timeout=4)
            if resp_nom.status_code == 200:
                dados_nom = resp_nom.json()
                if len(dados_nom) > 0:
                    lat_nom = float(dados_nom[0]['lat'])
                    lon_nom = float(dados_nom[0]['lon'])
                    if not (round(lat_nom, 3) == -23.550 and round(lon_nom, 3) == -46.633):
                        return lat_nom, lon_nom
        except Exception:
            pass

    return None, None


def gerar_link_waze(endereco, cep, lat, lon):
    coordenada_invalida = (
        lat is None or lon is None or pd.isna(lat) or pd.isna(lon) or
        (round(float(lat), 3) == -23.550 and round(float(lon), 3) == -46.633)
    )
    if coordenada_invalida:
        termo = f"{endereco or ''}, {cep or ''}, São Paulo, SP"
        return f"https://waze.com/ul?q={urllib.parse.quote(termo)}&navigate=yes"
    return f"https://waze.com/ul?ll={lat},{lon}&navigate=yes"


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
# 4. CONSTANTES E CATÁLOGOS MESTRES DE VISTORIA RIA
# -----------------------------------------------------------------------------
OPCOES_ST = [
    "FN - Fora Norma",
    "RU - Ruim",
    "AC - Acompanhar",
    "NT - Não tem",
    "NA - Não se aplica",
    "BO - Bom"
]

OPCOES_PR = [
    "OR - Orçar",
    "VM - Verificar-máquinas",
    "VT - Verificar-técnica",
    "VP - Verificar-preventiva",
    "CA - Enviar carta",
    "- (Nenhuma)"
]

ESTRUTURA_PREVENTIVA = {
    "Casa de Máquinas": {
        "Máquina de Tração / Centralina": [
            "Escova", "Freio (ajuste / funcionamento)", "Óleo (nível)", "Ruído", "Vazamento", "Limpeza",
            "Polia de tração"
        ],
        "Quadro de Comando": [
            "Isento de jumpers", "Relé, bobina, fusível, etc", "Limpeza"
        ],
        "Limitador de Velocidade": [
            "Contato elétrico", "Limpeza / Lubrificação"
        ]
    },
    "Pavimento": {
        "Portas de Pavimento": [
            "Trinco", "Corrediça", "Protetores de ilhós com tampa", "Limpeza / Lubrificação"
        ],
        "Sinalização": [
            "Funcionamento das botoeiras", "Funcionamento setas direção / IPD"
        ]
    },
    "Passadiço": {
        "Operador de Porta": [
            "Contato, roldana, rampa, etc", "Limpeza / Lubrificação"
        ],
        "Cabina": [
            "Botoeira de cabina", "Iluminação, acrílico, etc", "Nivelamento cabina X pavimento",
            "Aparelho de segurança cabina", "Luz de emergência"
        ],
        "Guias": [
            "Limpeza / Folgas", "Limpeza / Lubrificação"
        ],
        "Cabos de Aço Tração": [
            "Desgaste, equalização", "Distância contrapeso para-choque"
        ],
        "Estrutura Contrapeso / Cabina": [
            "Desgaste / Folga corrediças cabina", "Desgaste / Folga corrediças contrapeso"
        ]
    },
    "Poço": {
        "Poço & Segurança": [
            "Limpeza do poço", "Distância batedor contrapeso e para-choque", "Chave de emergência / PAP"
        ],
        "Polia Tensora": [
            "Polia tensora aspecto geral", "Contato elétrico", "Limpeza / Lubrificação"
        ]
    }
}

ESTRUTURA_NORMAS = {
    "Casa de Máquinas": [
        ("001", "Alçapão"),
        ("014", "Chave geral do quadro de força"),
        ("015", "Chave geral identificação nº. Elev."),
        ("016", "Disjuntor"),
        ("003", "Extintor de incêndio"),
        ("004", "Furação passagem de cabos"),
        ("018", "Fusível"),
        ("005", "Gancho carga (ausente/ident. Carga)"),
        ("328", "Guarda-corpo/escada casa maq."),
        ("006", "Identif. Equip. cab. Quadro maq."),
        ("008", "Infiltração"),
        ("326", "Intercomunicador casa máquinas"),
        ("329", "Luz de emergência casa máquinas"),
        ("347", "Marcação dos cabos de aço"),
        ("330", "Placa de aviso para alçapão"),
        ("332", "Placa de aviso para porta de inspeção"),
        ("331", "Placa de aviso porta casa máquinas"),
        ("009", "Porta de acesso casa de máquinas"),
        ("327", "Porta e fechadura porta casa maq."),
        ("046", "Protetor de polia de tração"),
        ("133", "Protetor de polia do limitador"),
        ("020", "Quadro de força"),
        ("011", "Ventilação casa de máquinas")
    ],
    "Cabina": [
        ("140", "Botoeira de inspeção"),
        ("166", "Braile para botão de cabina"),
        ("146", "Guarda-corpo"),
        ("170", "Intercomunicador cabina"),
        ("176", "Placa de capacidade"),
        ("175", "Placa de conservação CREL"),
        ("172", "Luz de emergência")
    ],
    "Passadiço": [
        ("338", "Iluminação do passadiço"),
        ("337", "Kit alça de segurança"),
        ("186", "Tapa vista")
    ],
    "Pavimento": [
        ("228", "Ilhós de porta de pavimento")
    ],
    "Poço": [
        ("346", "Coletor de óleo"),
        ("241", "Contato elétrico polia tensora"),
        ("242", "Escada marinheiro"),
        ("149", "Lubrificador"),
        ("240", "Chave PAP"),
        ("340", "Placa aviso verifique elev. Parado pavto."),
        ("343", "Placa de aviso chave PAP"),
        ("342", "Placa de aviso dist. Contrapeso"),
        ("344", "Placa de aviso escada acesso poço")
    ],
    "Portaria": [
        ("345", "Intercomunicador portaria")
    ]
}


# -----------------------------------------------------------------------------
# FUNÇÃO REUTILIZÁVEL: BOTÃO DE IMPRESSÃO / SALVAR EM PDF ISOLADO
# -----------------------------------------------------------------------------
def renderizar_botao_impressao_isolada(titulo_doc, subtitulo_doc, conteudo_html,
                                       titulo_btn="🖨️ Imprimir / Salvar em PDF", cor_fundo="#1B365D"):
    conteudo_sanitizado = conteudo_html.replace('"', '&quot;').replace('\n', ' ')
    titulo_sanitizado = titulo_doc.replace('"', '&quot;').replace('\n', ' ')
    sub_sanitizado = subtitulo_doc.replace('"', '&quot;').replace('\n', ' ')

    js_code = f"""
    <script>
    function dispararImpressao() {{
        var htmlConteudo = "{conteudo_sanitizado}";
        var titulo = "{titulo_sanitizado}";
        var sub = "{sub_sanitizado}";

        var win = window.open('', '_blank', 'width=950,height=750');
        if (!win) {{
            alert('Por favor, autorize a abertura de pop-ups neste navegador para imprimir o relatório.');
            return;
        }}

        win.document.open();
        win.document.write('<!DOCTYPE html><html><head><title>' + titulo + '</title>');
        win.document.write('<style>');
        win.document.write('@page {{ size: A4 portrait; margin: 12mm; }}');
        win.document.write('body {{ font-family: Arial, Helvetica, sans-serif; padding: 15px; color: #222; font-size: 11px; }}');
        win.document.write('.header-doc {{ border-bottom: 2px solid #1B365D; padding-bottom: 8px; margin-bottom: 12px; }}');
        win.document.write('.header-doc h2 {{ color: #1B365D; margin: 0; font-size: 17px; }}');
        win.document.write('.header-doc p {{ margin: 4px 0 0 0; color: #444; font-size: 12px; }}');
        win.document.write('h4 {{ color: #1B365D; margin-top: 14px; margin-bottom: 6px; font-size: 13px; border-bottom: 1px solid #ddd; padding-bottom: 3px; }}');
        win.document.write('table {{ width: 100%; border-collapse: collapse; margin-top: 6px; margin-bottom: 14px; }}');
        win.document.write('th, td {{ border: 1px solid #777; padding: 5px 7px; text-align: left; font-size: 10.5px; }}');
        win.document.write('th {{ background-color: #1B365D; color: white; -webkit-print-color-adjust: exact; print-color-adjust: exact; }}');
        win.document.write('tr:nth-child(even) {{ background-color: #f8fafc; -webkit-print-color-adjust: exact; }}');
        win.document.write('.quebra-doc {{ page-break-before: always; }}');
        win.document.write('</style></head><body>');
        win.document.write('<div class="header-doc">');
        win.document.write('<h2>CREL ELEVADORES — ' + titulo + '</h2>');
        win.document.write('<p>' + sub + '</p>');
        win.document.write('</div>');
        win.document.write(htmlConteudo);
        win.document.write('</body></html>');
        win.document.close();

        setTimeout(function() {{
            win.focus();
            win.print();
        }}, 400);
    }}
    </script>
    <button onclick="dispararImpressao()" style="
        width: 100%; height: 42px; background-color: {cor_fundo}; color: #ffffff;
        border: none; border-radius: 6px; font-size: 14px; font-weight: bold; cursor: pointer;
        box-shadow: 0 2px 4px rgba(0,0,0,0.15); margin-bottom: 12px;
    ">{titulo_btn}</button>
    """
    return components.html(js_code, height=55)


# -----------------------------------------------------------------------------
# 5. NAVEGAÇÃO PRINCIPAL
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
                valores = [str(v).upper() for v in row_data.values if pd.notna(v)]
                if any("CLIENTE" in v for v in valores) and any("ENDEREÇO" in v or "ENDERECO" in v for v in valores):
                    linha_cabecalho = idx_row
                    break

            if linha_cabecalho is not None:
                df_raw.columns = [str(c).strip().upper() for c in df_raw.iloc[linha_cabecalho]]
                df = df_raw.iloc[linha_cabecalho + 1:].copy().reset_index(drop=True)
            else:
                df = df_raw.copy()
                df.columns = [str(c).strip().upper() for c in df.columns]

            def remover_acentos(texto):
                return ''.join(c for c in unicodedata.normalize('NFD', str(texto)) if unicodedata.category(c) != 'Mn')

            def achar_coluna(termos_busca):
                for col in df.columns:
                    col_limpa = re.sub(r'[^a-zA-Z0-9]', '', remover_acentos(col)).lower()
                    for termo in termos_busca:
                        termo_limpo = re.sub(r'[^a-zA-Z0-9]', '', remover_acentos(termo)).lower()
                        if termo_limpo == col_limpa or termo_limpo in col_limpa:
                            return col
                return None

            col_cli = achar_coluna(['cliente', 'condominio', 'nome', 'razao', 'edificio'])
            col_end = achar_coluna(['endereco', 'logradouro', 'rua', 'local', 'end'])
            col_cep = achar_coluna(['cep', 'codigopostal'])
            col_qtd = achar_coluna(['qtd', 'elevadores', 'qtdelev', 'quantidade', 'equipamentos', 'qtde'])
            col_marca = achar_coluna(['marca', 'fabricante'])
            col_venc = achar_coluna(['mes', 'exec', 'vencto', 'vencimento', 'venc', 'data'])

            if not (col_cli and col_end):
                st.error("Não foi possível identificar as colunas de CLIENTE e ENDEREÇO.")
                st.stop()

            df = df[df[col_cli].notna() & (df[col_cli].astype(str).str.strip() != '')].copy()
            df['QTD_CLEAN'] = pd.to_numeric(df[col_qtd], errors='coerce').fillna(1).astype(int) if col_qtd else 1
            df['MARCA_CLEAN'] = df[col_marca].astype(str).str.strip().str.upper() if col_marca else 'DIVERSAS'
            df['VENC_CLEAN'] = df[col_venc].astype(str).str.strip() if col_venc else 'Sem Data'
            if not col_cep:
                df['CEP'] = ""
                col_cep = 'CEP'

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
            v_str = str(v).strip().lower()
            if not v_str or "sem data" in v_str or "//" in v_str or v_str in ["nan", "none", "-", ""]:
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
        df_edit = pd.DataFrame({
            "Incluir": True,
            "CLIENTE": df_f[col_cli].astype(str),
            "CEP": df_f[col_cep].astype(str),
            "ENDEREÇO": df_f[col_end].astype(str),
            "QTD": df_f['QTD_CLEAN'],
            "MARCA": df_f['MARCA_CLEAN'],
            "VENCTO RIA": df_f['VENC_CLEAN'],
            "GRUPO": df_f['GRUPO_STATUS']
        })

        grid_selecao = st.data_editor(
            df_edit,
            use_container_width=True,
            hide_index=True,
            column_config={
                "Incluir": st.column_config.CheckboxColumn("Incluir?", default=True),
                "QTD": st.column_config.NumberColumn("Qtd Elev", format="%d")
            },
            disabled=["CLIENTE", "CEP", "ENDEREÇO", "QTD", "MARCA", "VENCTO RIA", "GRUPO"]
        )
        selecionados = grid_selecao[grid_selecao["Incluir"] == True].copy()

        st.markdown("---")
        st.subheader("🚚 Configuração das Jornadas Diárias")
        c1, c2, c3 = st.columns(3)
        with c1:
            data_inicio_prog = st.date_input("Data de Início da Rota:", datetime.date.today())
        with c2:
            max_locais_dia = st.number_input("Máximo de Locais por Dia:", min_value=1, max_value=30, value=7, step=1)
        with c3:
            max_elev_dia = st.number_input("Máximo de Elevadores por Dia:", min_value=1, max_value=60, value=14, step=1)

        if st.button("🚀 Otimizar Trajetos e Salvar para Campo", type="primary", use_container_width=True):
            if selecionados.empty:
                st.warning("Nenhum local selecionado na tabela.")
            else:
                with st.spinner("Geolocalizando pontos e organizando jornadas por menor distância..."):
                    dados_processados = []
                    for _, r_sel in selecionados.iterrows():
                        norm = normalizar_codigo(r_sel['CLIENTE'])
                        matches = df_f[df_f['CLI_NORM'] == norm]
                        p_lat, p_lon = None, None
                        if not matches.empty:
                            row_orig = matches.iloc[0]
                            p_lat = row_orig['lat']
                            p_lon = row_orig['lon']

                        if pd.isna(p_lat) or pd.isna(p_lon) or p_lat is None:
                            p_lat, p_lon = buscar_coords_fallback(r_sel['CEP'], r_sel['ENDEREÇO'])

                        waze = gerar_link_waze(r_sel['ENDEREÇO'], r_sel['CEP'], p_lat, p_lon)

                        lat_val = float(p_lat) if (p_lat is not None and not pd.isna(p_lat)) else None
                        lon_val = float(p_lon) if (p_lon is not None and not pd.isna(p_lon)) else None

                        dados_processados.append({
                            'cliente': r_sel['CLIENTE'],
                            'cep': str(r_sel['CEP'] or ''),
                            'endereco': r_sel['ENDEREÇO'],
                            'qtd_elev': int(r_sel['QTD']),
                            'marca': r_sel['MARCA'],
                            'vencto_ria': r_sel['VENCTO RIA'],
                            'grupo_status': r_sel['GRUPO'],
                            'lat': lat_val,
                            'lon': lon_val,
                            'waze_url': waze
                        })

                    rotas_finais = []
                    pontos_restantes = dados_processados.copy()
                    dia_num = 1
                    dia_corrente = data_inicio_prog

                    def avancar_dia_util(d):
                        prox = d + datetime.timedelta(days=1)
                        while prox.weekday() >= 5:
                            prox += datetime.timedelta(days=1)
                        return prox

                    resumo_dias = []

                    while pontos_restantes:
                        dia_nome = f"Dia {dia_num} - {dia_corrente.strftime('%d/%m/%Y')}"
                        parada_num = 1
                        elev_acum = 0
                        locais_acum = 0
                        itens_deste_dia = []

                        atual = pontos_restantes.pop(0)
                        atual['dia_rota'] = dia_nome
                        atual['ordem_parada'] = parada_num
                        elev_acum += atual['qtd_elev']
                        locais_acum += 1
                        rotas_finais.append(atual)
                        itens_deste_dia.append(atual)

                        while pontos_restantes:
                            candidatos = [
                                p for p in pontos_restantes
                                if (locais_acum + 1 <= max_locais_dia) and (elev_acum + p['qtd_elev'] <= max_elev_dia)
                            ]

                            if not candidatos:
                                break

                            ref_lat = atual['lat'] if atual['lat'] is not None else -23.5505
                            ref_lon = atual['lon'] if atual['lon'] is not None else -46.6333

                            proximo = min(
                                candidatos,
                                key=lambda p_cand: calc_dist_km(
                                    ref_lat,
                                    ref_lon,
                                    p_cand['lat'] if p_cand['lat'] is not None else -23.5505,
                                    p_cand['lon'] if p_cand['lon'] is not None else -46.6333
                                )
                            )
                            pontos_restantes.remove(proximo)
                            parada_num += 1
                            locais_acum += 1
                            elev_acum += proximo['qtd_elev']
                            proximo['dia_rota'] = dia_nome
                            proximo['ordem_parada'] = parada_num
                            rotas_finais.append(proximo)
                            itens_deste_dia.append(proximo)
                            atual = proximo

                        resumo_dias.append({
                            "dia_nome": dia_nome,
                            "itens": itens_deste_dia,
                            "elev": elev_acum
                        })

                        dia_num += 1
                        dia_corrente = avancar_dia_util(dia_corrente)

                    engine_salvar = get_engine()
                    with engine_salvar.begin() as conn:
                        conn.execute(text("DELETE FROM rotas_ativas"))
                        for p_item in rotas_finais:
                            uid = f"{p_item['dia_rota']}_{p_item['ordem_parada']}_{normalizar_codigo(p_item['cliente'])}"
                            conn.execute(text('''
                                INSERT INTO rotas_ativas 
                                (id, mes_exec, dia_rota, ordem_parada, cliente, cep, endereco, qtd_elev, marca, vencto_ria, grupo_status, lat, lon, waze_url, status_vistoria, data_baixa, vistoriador, obs)
                                VALUES (:id, :mes, :dia, :ordem, :cli, :cep, :end, :qtd, :marca, :venc, :grupo, :lat, :lon, :waze, 'Pendente', '', '', '')
                            '''), {
                                "id": uid,
                                "mes": data_inicio_prog.strftime("%m/%Y"),
                                "dia": p_item['dia_rota'],
                                "ordem": p_item['ordem_parada'],
                                "cli": p_item['cliente'],
                                "cep": p_item['cep'],
                                "end": p_item['endereco'],
                                "qtd": p_item['qtd_elev'],
                                "marca": p_item['marca'],
                                "venc": p_item['vencto_ria'],
                                "grupo": p_item['grupo_status'],
                                "lat": p_item['lat'],
                                "lon": p_item['lon'],
                                "waze": p_item['waze_url']
                            })

                    st.success(
                        f"✅ Roteirização concluída! {len(rotas_finais)} locais distribuídos em {len(resumo_dias)} jornadas diárias.")

                    st.markdown("### 📅 Resumo dos Dias e Paradas Formadas:")
                    for b_dia in resumo_dias:
                        with st.expander(
                                f"📌 **{b_dia['dia_nome']}** — {len(b_dia['itens'])} paradas | {b_dia['elev']} elevadores",
                                expanded=True):
                            tab_res = pd.DataFrame([{
                                "Parada": it['ordem_parada'],
                                "Cliente": it['cliente'],
                                "Endereço": it['endereco'],
                                "CEP": it['cep'],
                                "Elevadores": it['qtd_elev'],
                                "Marca": it['marca']
                            } for it in b_dia['itens']])
                            st.table(tab_res)

# =============================================================================
# ETAPA 2: RELATÓRIOS & PAINEL DO GESTOR
# =============================================================================
elif menu == "📄 2. Relatórios & Painel do Gestor":
    st.title("📄 Painel do Gestor — Roteiros & Relatórios de Vistoria")

    st.markdown("""
    <style>
    @media print {
        header, footer, [data-testid="stSidebar"], [data-testid="stHeader"], 
        .stButton, [data-testid="stDownloadButton"], iframe, .no-print, [data-testid="stTabs"] {
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
        margin-top: 8px;
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

    engine_gestor = get_engine()
    try:
        df_rel = pd.read_sql("SELECT * FROM rotas_ativas", engine_gestor)
        df_vistorias_db = pd.read_sql("SELECT * FROM relatorios_vistorias", engine_gestor)
    except Exception:
        df_rel = pd.DataFrame()
        df_vistorias_db = pd.DataFrame()

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
                buffer_excel = io.BytesIO()
                with pd.ExcelWriter(buffer_excel, engine='openpyxl') as writer:
                    df_rel.to_excel(writer, index=False, sheet_name="Roteiro_Geral")
                st.download_button(
                    "📥 Baixar Roteiro em Excel",
                    buffer_excel.getvalue(),
                    "Relatorio_Consolidado_RIA.xlsx",
                    use_container_width=True
                )

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
        st.subheader("📁 Vistorias Técnicas Salvas na Nuvem")
        if df_vistorias_db.empty:
            st.info("Nenhuma vistoria técnica finalizada até o momento.")
        else:
            st.dataframe(df_vistorias_db[['cliente', 'endereco', 'data_vistoria', 'vistoriador', 'qtd_normas_orc',
                                          'qtd_pendencias_sup']], use_container_width=True)
            st.markdown("---")

            opcoes_vistorias = {
                f"{r['cliente']} - {r['data_vistoria']} ({r['id_rota']})": r['id_rota']
                for _, r in df_vistorias_db.iterrows()
            }

            escolha_rel = st.selectbox("Selecione a Vistoria para Visualizar, Imprimir ou Salvar em PDF:",
                                       list(opcoes_vistorias.keys()))
            id_escolhido = opcoes_vistorias[escolha_rel]
            registro_vistoria = df_vistorias_db[df_vistorias_db['id_rota'] == id_escolhido].iloc[0]

            xls = None
            bytes_excel_export = None

            if 'arquivo_binario' in registro_vistoria and pd.notna(registro_vistoria['arquivo_binario']) and str(registro_vistoria['arquivo_binario']).strip():
                try:
                    bytes_excel_export = base64.b64decode(registro_vistoria['arquivo_binario'])
                    xls = pd.ExcelFile(io.BytesIO(bytes_excel_export))
                except Exception:
                    xls = None

            if xls is None and registro_vistoria.get('caminho_arquivo') and os.path.exists(registro_vistoria['caminho_arquivo']):
                try:
                    xls = pd.ExcelFile(registro_vistoria['caminho_arquivo'])
                    with open(registro_vistoria['caminho_arquivo'], "rb") as f_b:
                        bytes_excel_export = f_b.read()
                except Exception:
                    xls = None

            if xls is not None:
                planilhas = xls.sheet_names

                df_id_elev = pd.read_excel(xls, sheet_name="Elevadores_Identificacao") if "Elevadores_Identificacao" in planilhas else pd.DataFrame()
                df_orcar = pd.read_excel(xls, sheet_name="Acoes_Orcamento") if "Acoes_Orcamento" in planilhas else (
                    pd.read_excel(xls, sheet_name="Orcamento_Normas") if "Orcamento_Normas" in planilhas else pd.DataFrame()
                )
                df_prev = pd.read_excel(xls, sheet_name="Acoes_Preventiva") if "Acoes_Preventiva" in planilhas else pd.DataFrame()
                df_tec = pd.read_excel(xls, sheet_name="Acoes_Tecnico") if "Acoes_Tecnico" in planilhas else (
                    pd.read_excel(xls, sheet_name="Supervisao_Pendencias") if "Supervisao_Pendencias" in planilhas else pd.DataFrame()
                )
                df_mec = pd.read_excel(xls, sheet_name="Acoes_Mecanica") if "Acoes_Mecanica" in planilhas else pd.DataFrame()
                df_prev_matriz = pd.read_excel(xls, sheet_name="Preventiva_Matriz_Integral") if "Preventiva_Matriz_Integral" in planilhas else (
                    pd.read_excel(xls, sheet_name="Manutencao_Preventiva") if "Manutencao_Preventiva" in planilhas else pd.DataFrame()
                )
                df_normas_comp = pd.read_excel(xls, sheet_name="Normas_Checklist_Integral") if "Normas_Checklist_Integral" in planilhas else pd.DataFrame()

                cli_nome = escolha_rel.split(' - ')[0]
                dt_rel = escolha_rel.split(' - ')[1].split(' (')[0]
                sub_cab = f"Edifício: {cli_nome} | Data: {dt_rel}"

                tabela_elev_html = df_id_elev.to_html(classes="tabela-rel-doc", index=False) if not df_id_elev.empty else ""

                html_dossie_geral = f"""
                <h4>1. Identificação dos Equipamentos</h4>
                {tabela_elev_html}
                <h4>2. Matriz Integral de Manutenção Preventiva</h4>
                {df_prev_matriz.to_html(classes="tabela-rel-doc", index=False) if not df_prev_matriz.empty else "<p>Sem dados.</p>"}
                <div class="quebra-doc" style="page-break-before: always;"></div>
                <h4>3. Checklist de Adequações de Normas Técnicas</h4>
                {df_normas_comp.to_html(classes="tabela-rel-doc", index=False) if not df_normas_comp.empty else (df_orcar.to_html(classes="tabela-rel-doc", index=False) if not df_orcar.empty else "<p>Sem pendências registradas.</p>")}
                <div style="margin-top: 35px; display: flex; justify-content: space-between;">
                    <div style="border-top: 1px solid #444; width: 45%; text-align: center; padding-top: 5px;">
                        <b>CREL Elevadores</b><br>Técnico Responsável
                    </div>
                    <div style="border-top: 1px solid #444; width: 45%; text-align: center; padding-top: 5px;">
                        <b>CREL Elevadores</b><br>Supervisão Técnica
                    </div>
                </div>
                """

                c_btn_dossie, c_btn_xls = st.columns([1.5, 1])
                with c_btn_dossie:
                    renderizar_botao_impressao_isolada(
                        "DOSSIÊ GERAL DE VISTORIA RIA (ARQUIVO COMPLETO)",
                        sub_cab,
                        html_dossie_geral,
                        "🖨️ Imprimir Dossiê Geral Completo (PDF)",
                        cor_fundo="#0F172A"
                    )
                with c_btn_xls:
                    if bytes_excel_export:
                        st.download_button(
                            label="📥 Baixar Pasta Excel (.xlsx)",
                            data=bytes_excel_export,
                            file_name=f"Vistoria_{normalizar_codigo(cli_nome)}.xlsx",
                            use_container_width=True
                        )

                st.markdown(f"""
                <div style="border-bottom: 2px solid #1B365D; padding-bottom: 8px; margin-top: 10px; margin-bottom: 15px;">
                    <h3 style="color:#1B365D; margin:0;">RELATÓRIO TÉCNICO DE VISTORIA — RIA</h3>
                    <p style="margin:4px 0 0 0; color:#444; font-size:13px;"><b>{sub_cab}</b></p>
                </div>
                """, unsafe_allow_html=True)

                t_orcar_s, t_prev_s, t_tec_s, t_mec_s, t_dossie_s = st.tabs([
                    "💰 1. Itens a Orçar",
                    "🔧 2. Itens Preventivo",
                    "🚨 3. Itens Técnico",
                    "⚙️ 4. Itens Mecânica",
                    "📑 5. Matriz Integral / Dossiê"
                ])

                with t_orcar_s:
                    html_orcar_tabela = df_orcar.to_html(classes="tabela-rel-doc", index=False) if not df_orcar.empty else "<p>Nenhum item registrado para orçamento comercial.</p>"
                    html_imp_orcar = f"<h4>Identificação dos Equipamentos</h4>{tabela_elev_html}<h4>Itens Apontados para Elaboração de Orçamento</h4>{html_orcar_tabela}"
                    renderizar_botao_impressao_isolada("RELATÓRIO: ITENS A ORÇAR", sub_cab, html_imp_orcar, "🖨️ Imprimir / Salvar Relatório de Orçamentos em PDF")
                    st.markdown("##### Identificação dos Equipamentos")
                    st.html(tabela_elev_html)
                    st.markdown("##### Itens para Orçamento Comercial")
                    if not df_orcar.empty:
                        st.html(html_orcar_tabela)
                    else:
                        st.info("Nenhum item registrado para orçamento comercial.")

                with t_prev_s:
                    html_prev_tabela = df_prev.to_html(classes="tabela-rel-doc", index=False) if not df_prev.empty else "<p>Nenhum item apontado para ação preventiva. Todos conformes.</p>"
                    html_imp_prev = f"<h4>Identificação dos Equipamentos</h4>{tabela_elev_html}<h4>Itens Apontados para Ação da Preventiva</h4>{html_prev_tabela}"
                    renderizar_botao_impressao_isolada("RELATÓRIO: ITENS PREVENTIVO", sub_cab, html_imp_prev, "🖨️ Imprimir / Salvar Relatório Preventivo em PDF")
                    st.markdown("##### Identificação dos Equipamentos")
                    st.html(tabela_elev_html)
                    st.markdown("##### Itens Apontados para Ação da Preventiva")
                    if not df_prev.empty:
                        st.html(html_prev_tabela)
                    else:
                        st.info("Nenhum item apontado para ação da preventiva. Todos os itens de rotina estão conformes.")

                with t_tec_s:
                    html_tec_tabela = df_tec.to_html(classes="tabela-rel-doc", index=False) if not df_tec.empty else "<p>Nenhuma pendência técnica registrada.</p>"
                    html_imp_tec = f"<h4>Identificação dos Equipamentos</h4>{tabela_elev_html}<h4>Itens Apontados para Ação Técnica / Suporte / Cartas</h4>{html_tec_tabela}"
                    renderizar_botao_impressao_isolada("RELATÓRIO: ITENS TÉCNICO", sub_cab, html_imp_tec, "🖨️ Imprimir / Salvar Relatório Técnico em PDF")
                    st.markdown("##### Identificação dos Equipamentos")
                    st.html(tabela_elev_html)
                    st.markdown("##### Pendências Técnicas e Apontamentos")
                    if not df_tec.empty:
                        st.html(html_tec_tabela)
                    else:
                        st.info("Nenhuma pendência técnica registrada.")

                with t_mec_s:
                    html_mec_tabela = df_mec.to_html(classes="tabela-rel-doc", index=False) if not df_mec.empty else "<p>Nenhuma intervenção mecânica solicitada.</p>"
                    html_imp_mec = f"<h4>Identificação dos Equipamentos</h4>{tabela_elev_html}<h4>Itens Apontados para Intervenção Mecânica (VM)</h4>{html_mec_tabela}"
                    renderizar_botao_impressao_isolada("RELATÓRIO: ITENS MECÂNICA", sub_cab, html_imp_mec, "🖨️ Imprimir / Salvar Relatório de Mecânica em PDF")
                    st.markdown("##### Identificação dos Equipamentos")
                    st.html(tabela_elev_html)
                    st.markdown("##### Apontamentos de Máquinas e Equipamentos Mecânicos")
                    if not df_mec.empty:
                        st.html(html_mec_tabela)
                    else:
                        st.info("Nenhuma intervenção mecânica solicitada.")

                with t_dossie_s:
                    st.markdown("##### 1. Matriz Integral de Manutenção Preventiva")
                    if not df_prev_matriz.empty:
                        st.html(df_prev_matriz.to_html(classes="tabela-rel-doc", index=False))
                    else:
                        st.info("Sem dados de manutenção preventiva disponíveis.")

                    st.markdown("##### 2. Checklist Integral de Adequações de Normas Técnicas")
                    if not df_normas_comp.empty:
                        st.html(df_normas_comp.to_html(classes="tabela-rel-doc", index=False))
                    elif not df_orcar.empty:
                        st.html(df_orcar.to_html(classes="tabela-rel-doc", index=False))
            else:
                st.info("O arquivo de vistoria física detalhada não está armazenado nesta instância (foi gerado antes da ativação da persistência na nuvem). Novas vistorias realizadas a partir de agora ficarão permanentemente armazenadas no banco.")

# =============================================================================
# ETAPA 3: TERMINAL DE CAMPO & FORMULÁRIO ENCAPSULADO (SEM TRAVAMENTOS)
# =============================================================================
elif menu == "📱 3. Terminal de Campo (Vistoriador)":
    engine_campo = get_engine()
    try:
        df_todos = pd.read_sql("SELECT * FROM rotas_ativas", engine_campo)
    except Exception:
        df_todos = pd.DataFrame()

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

        qtd_padrao = max(int(dados_cli['qtd_elev'] or 1), 1)
        qtd_carros = st.number_input(
            "Quantidade Total de Elevadores no Edifício:",
            min_value=1,
            max_value=30,
            value=qtd_padrao,
            step=1,
            key=f"num_elev_{id_ativa}"
        )

        todos_carros = [f"C{i + 1}" for i in range(int(qtd_carros))]

        st.info("ℹ️ **Modo Desconectado Ativo:** Preencha todo o checklist sem lentidão. Os dados só serão transmitidos para a nuvem quando você clicar no botão 'Salvar e Finalizar Vistoria' ao final da página.")

        # FORMULÁRIO ENCAPSULADO: Zero transmissões durante cliques em checkboxes/selects
        with st.form(f"form_execucao_{id_ativa}"):
            with st.expander("🏢 1. Identificação dos Elevadores (CREL & PMSP)", expanded=True):
                colunas_por_linha = min(len(todos_carros), 4)
                cols_cad = st.columns(colunas_por_linha)
                for i, c_id in enumerate(todos_carros):
                    with cols_cad[i % colunas_por_linha]:
                        st.markdown(f"**Elevador {c_id}**")
                        st.text_input(f"Nº CREL ({c_id}):", key=f"crel_{id_ativa}_{c_id}")
                        st.text_input(f"Nº PMSP ({c_id}):", key=f"pmsp_{id_ativa}_{c_id}")

                col_tec, col_sup = st.columns(2)
                with col_tec:
                    nome_vistoriador = st.text_input("Técnico Vistoriador:", value=dados_cli['vistoriador'] or "",
                                                     key=f"tec_{id_ativa}")
                with col_sup:
                    nome_supervisor = st.text_input("Supervisor Técnico:", key=f"sup_{id_ativa}")

            st.markdown("### 📋 2. Apontamento de Irregularidades")
            aba_norma, aba_prev, aba_sup = st.tabs([
                "💰 1. Adequações de Norma Técnica (Orçamento)",
                "🔧 2. Manutenção Preventiva (Rotina)",
                "🚨 3. Parecer da Supervisão Técnica"
            ])

            # ABA 1: ADEQUAÇÕES DE NORMA TÉCNICA
            with aba_norma:
                for zona_norma, itens_norma in ESTRUTURA_NORMAS.items():
                    with st.expander(f"📌 {zona_norma} ({len(itens_norma)} itens)", expanded=False):
                        for cod, desc in itens_norma:
                            chave_n = f"norm_{id_ativa}_{zona_norma}_{cod}"
                            c_chk, c_st, c_pr, c_car, c_obs = st.columns([3, 1.2, 1.5, 2, 2.5])
                            with c_chk:
                                st.checkbox(f"**[{cod}]** {desc}", key=f"chk_{chave_n}")
                            with c_st:
                                st.selectbox("Situação:", OPCOES_ST, index=0, key=f"st_{chave_n}", label_visibility="collapsed")
                            with c_pr:
                                st.selectbox("Providência:", OPCOES_PR, index=0, key=f"pr_{chave_n}", label_visibility="collapsed")
                            with c_car:
                                st.multiselect("Carros:", options=todos_carros, default=todos_carros, key=f"car_{chave_n}", label_visibility="collapsed")
                            with c_obs:
                                st.text_input("Obs:", placeholder="Nota da norma", key=f"obs_{chave_n}", label_visibility="collapsed")

            # ABA 2: MANUTENÇÃO PREVENTIVA
            with aba_prev:
                for local_nome, subs in ESTRUTURA_PREVENTIVA.items():
                    with st.expander(f"📍 {local_nome}", expanded=False):
                        for item_nome, pontos in subs.items():
                            st.markdown(f"**{item_nome}**")
                            for ponto in pontos:
                                chave_p = f"prev_{id_ativa}_{local_nome}_{item_nome}_{ponto}".replace(" ", "_")
                                c_chk, c_st, c_pr, c_car, c_obs = st.columns([3, 1.2, 1.5, 2, 2.5])
                                with c_chk:
                                    st.checkbox(f"Irregular: {ponto}", key=f"chk_{chave_p}")
                                with c_st:
                                    st.selectbox("ST:", OPCOES_ST, index=1, key=f"st_{chave_p}", label_visibility="collapsed")
                                with c_pr:
                                    st.selectbox("PR:", OPCOES_PR, index=3, key=f"pr_{chave_p}", label_visibility="collapsed")
                                with c_car:
                                    st.multiselect("Carros:", options=todos_carros, default=todos_carros, key=f"car_{chave_p}", label_visibility="collapsed")
                                with c_obs:
                                    st.text_input("Obs:", placeholder="Irregularidade", key=f"obs_{chave_p}", label_visibility="collapsed")

            # ABA 3: APONTAMENTOS GERAIS DA SUPERVISÃO
            with aba_sup:
                parecer_sup = st.text_area(
                    "Parecer e Orientações Globais do Supervisor:",
                    placeholder="Descreva pendências globais, orientações para equipe de manutenção, prazos ou necessidades urgentes...",
                    key=f"parecer_{id_ativa}"
                )

            st.markdown("---")
            btn_salvar_definitivo = st.form_submit_button("💾 Salvar e Finalizar Vistoria (Transmitir para Nuvem)", type="primary", use_container_width=True)

        if btn_salvar_definitivo:
            # 1. Recupera identificação dos carros preenchida no form
            dados_elevadores_cad = []
            for c_id in todos_carros:
                c_val = st.session_state.get(f"crel_{id_ativa}_{c_id}", "")
                p_val = st.session_state.get(f"pmsp_{id_ativa}_{c_id}", "")
                dados_elevadores_cad.append({"Elevador": c_id, "Nº CREL": c_val, "Nº PMSP": p_val})

            # 2. Coleta irregularidades de normas marcadas
            irregularidades_norma = []
            for zona_norma, itens_norma in ESTRUTURA_NORMAS.items():
                for cod, desc in itens_norma:
                    chave_n = f"norm_{id_ativa}_{zona_norma}_{cod}"
                    if st.session_state.get(f"chk_{chave_n}", False):
                        st_n = st.session_state.get(f"st_{chave_n}", "FN - Fora Norma").split(" - ")[0]
                        pr_n = st.session_state.get(f"pr_{chave_n}", "OR - Orçar").split(" - ")[0].replace("- (Nenhuma)", "-")
                        car_n = st.session_state.get(f"car_{chave_n}", todos_carros)
                        obs_n = str(st.session_state.get(f"obs_{chave_n}", "")).strip()

                        irregularidades_norma.append({
                            "Código": cod,
                            "Local": zona_norma,
                            "Descrição": desc,
                            "Carros": ", ".join(car_n),
                            "Carros_Lista": car_n,
                            "Situação (ST)": st_n,
                            "Providência (PR)": pr_n,
                            "Observações": obs_n
                        })

            # 3. Coleta irregularidades da preventiva marcadas
            irregularidades_preventiva = []
            for local_nome, subs in ESTRUTURA_PREVENTIVA.items():
                for item_nome, pontos in subs.items():
                    for ponto in pontos:
                        chave_p = f"prev_{id_ativa}_{local_nome}_{item_nome}_{ponto}".replace(" ", "_")
                        if st.session_state.get(f"chk_{chave_p}", False):
                            st_val = st.session_state.get(f"st_{chave_p}", "RU - Ruim").split(" - ")[0]
                            pr_val = st.session_state.get(f"pr_{chave_p}", "VP - Verificar-preventiva").split(" - ")[0].replace("- (Nenhuma)", "-")
                            car_val = st.session_state.get(f"car_{chave_p}", todos_carros)
                            obs_val = str(st.session_state.get(f"obs_{chave_p}", "")).strip()

                            irregularidades_preventiva.append({
                                "Local": local_nome,
                                "Componente": item_nome,
                                "Ponto": ponto,
                                "Carros": ", ".join(car_val),
                                "Carros_Lista": car_val,
                                "Situação (ST)": st_val,
                                "Providência (PR)": pr_val,
                                "Observações": obs_val
                            })

            # Montagem das planilhas técnicas
            dados_orcar = []
            for n in irregularidades_norma:
                if n["Providência (PR)"] == "OR" or n["Situação (ST)"] == "FN":
                    dados_orcar.append({
                        "Código": n["Código"],
                        "Local": n["Local"],
                        "Item / Adequação": n["Descrição"],
                        "Elevador(es)": n["Carros"],
                        "Situação (ST)": n["Situação (ST)"],
                        "Providência (PR)": n["Providência (PR)"],
                        "Observação Comercial": n["Observações"] or "Adequação de norma técnica necessária"
                    })
            for p in irregularidades_preventiva:
                if p["Providência (PR)"] == "OR" or p["Situação (ST)"] == "FN":
                    dados_orcar.append({
                        "Código": "MANUT",
                        "Local": p["Local"],
                        "Item / Adequação": f"{p['Componente']} - {p['Ponto']}",
                        "Elevador(es)": p["Carros"],
                        "Situação (ST)": p["Situação (ST)"],
                        "Providência (PR)": p["Providência (PR)"],
                        "Observação Comercial": p["Observações"] or "Item reprovado demandando orçamento"
                    })
            df_orcar = pd.DataFrame(dados_orcar)

            dados_preventivo = []
            for p in irregularidades_preventiva:
                if p["Providência (PR)"] == "VP" or (
                        p["Providência (PR)"] not in ["VM", "VT", "OR", "CA"] and p["Situação (ST)"] in ["RU", "AC"]):
                    dados_preventivo.append({
                        "Local": p["Local"],
                        "Componente / Ponto": f"{p['Componente']} - {p['Ponto']}",
                        "Elevador(es)": p["Carros"],
                        "Situação (ST)": p["Situação (ST)"],
                        "Providência (PR)": p["Providência (PR)"],
                        "Ação Preventiva Solicitada": p["Observações"] or "Ajuste e correção na próxima rota de manutenção"
                    })
            for n in irregularidades_norma:
                if n["Providência (PR)"] == "VP":
                    dados_preventivo.append({
                        "Local": n["Local"],
                        "Componente / Ponto": f"[{n['Código']}] {n['Descrição']}",
                        "Elevador(es)": n["Carros"],
                        "Situação (ST)": n["Situação (ST)"],
                        "Providência (PR)": n["Providência (PR)"],
                        "Ação Preventiva Solicitada": n["Observações"] or "Verificação e correção pela preventiva"
                    })
            df_preventivo = pd.DataFrame(dados_preventivo)

            dados_tecnico = []
            if parecer_sup and parecer_sup.strip():
                dados_tecnico.append({
                    "Origem": "Supervisão",
                    "Local": "Geral",
                    "Item / Componente": "Parecer Geral do Supervisor",
                    "Elevador(es)": "Todos",
                    "Situação (ST)": "-",
                    "Providência (PR)": "Orientação",
                    "Ação Técnica Necessária": parecer_sup.strip(),
                    "Prioridade": "Alta"
                })
            for p in irregularidades_preventiva:
                if p["Providência (PR)"] in ["VT", "CA"]:
                    dados_tecnico.append({
                        "Origem": "Preventiva",
                        "Local": p["Local"],
                        "Item / Componente": f"{p['Componente']} - {p['Ponto']}",
                        "Elevador(es)": p["Carros"],
                        "Situação (ST)": p["Situação (ST)"],
                        "Providência (PR)": p["Providência (PR)"],
                        "Ação Técnica Necessária": p["Observações"] or ("Verificação e diagnóstico técnico" if p["Providência (PR)"] == "VT" else "Enviar carta técnica ao condomínio"),
                        "Prioridade": "Alta" if p["Situação (ST)"] == "RU" else "Média"
                    })
            for n in irregularidades_norma:
                if n["Providência (PR)"] in ["VT", "CA"]:
                    dados_tecnico.append({
                        "Origem": "Norma Técnica",
                        "Local": n["Local"],
                        "Item / Componente": f"[{n['Código']}] {n['Descrição']}",
                        "Elevador(es)": n["Carros"],
                        "Situação (ST)": n["Situação (ST)"],
                        "Providência (PR)": n["Providência (PR)"],
                        "Ação Técnica Necessária": n["Observações"] or ("Verificação técnica de campo" if n["Providência (PR)"] == "VT" else "Enviar carta técnica ao condomínio"),
                        "Prioridade": "Alta"
                    })
            df_tecnico = pd.DataFrame(dados_tecnico)

            dados_mecanica = []
            for p in irregularidades_preventiva:
                if p["Providência (PR)"] == "VM":
                    dados_mecanica.append({
                        "Local": p["Local"],
                        "Conjunto / Ponto Mecânico": f"{p['Componente']} - {p['Ponto']}",
                        "Elevador(es)": p["Carros"],
                        "Situação (ST)": p["Situação (ST)"],
                        "Diagnóstico / Intervenção Mecânica": p["Observações"] or "Inspeção e intervenção mecânica requerida",
                        "Prioridade": "Alta" if p["Situação (ST)"] == "RU" else "Média"
                    })
            for n in irregularidades_norma:
                if n["Providência (PR)"] == "VM":
                    dados_mecanica.append({
                        "Local": n["Local"],
                        "Conjunto / Ponto Mecânico": f"[{n['Código']}] {n['Descrição']}",
                        "Elevador(es)": n["Carros"],
                        "Situação (ST)": n["Situação (ST)"],
                        "Diagnóstico / Intervenção Mecânica": n["Observações"] or "Intervenção mecânica de norma",
                        "Prioridade": "Alta"
                    })
            df_mecanica = pd.DataFrame(dados_mecanica)

            todos_itens_prev = []
            for local_nome, subs in ESTRUTURA_PREVENTIVA.items():
                for item_nome, pontos in subs.items():
                    for ponto in pontos:
                        linha = {"Local": local_nome, "Componente": item_nome, "Ponto de Avaliação": ponto}
                        irreg = next((x for x in irregularidades_preventiva if
                                      x["Local"] == local_nome and x["Componente"] == item_nome and x["Ponto"] == ponto),
                                     None)
                        for cid in todos_carros:
                            if irreg and (cid in irreg["Carros_Lista"]):
                                linha[cid] = f"{irreg['Situação (ST)']} ({irreg['Providência (PR)']})" if irreg['Providência (PR)'] != '-' else irreg['Situação (ST)']
                            else:
                                linha[cid] = "OK"
                        linha["Observações"] = irreg["Observações"] if irreg else ""
                        todos_itens_prev.append(linha)

            colunas_ordenadas_prev = ["Local", "Componente", "Ponto de Avaliação"] + todos_carros + ["Observações"]
            df_preventiva_matriz_completa = pd.DataFrame(todos_itens_prev).reindex(columns=colunas_ordenadas_prev)

            todos_itens_normas = []
            for zona_norma, itens_norma in ESTRUTURA_NORMAS.items():
                for cod, desc in itens_norma:
                    irreg_n = next((x for x in irregularidades_norma if x["Código"] == cod), None)
                    if irreg_n:
                        todos_itens_normas.append({
                            "Código": cod,
                            "Local": zona_norma,
                            "Item de Norma": desc,
                            "Condição": f"Não Conforme - {irreg_n['Situação (ST)']}",
                            "Providência": irreg_n["Providência (PR)"],
                            "Elevador(es)": irreg_n["Carros"],
                            "Anotação Técnica": irreg_n["Observações"] or "Adequação necessária"
                        })
                    else:
                        todos_itens_normas.append({
                            "Código": cod,
                            "Local": zona_norma,
                            "Item de Norma": desc,
                            "Condição": "Conforme",
                            "Providência": "Nenhuma",
                            "Elevador(es)": "Todos",
                            "Anotação Técnica": "Item dentro dos padrões de norma"
                        })
            df_normas_completo = pd.DataFrame(todos_itens_normas)

            dt_hoje = datetime.datetime.now().strftime("%Y%m%d_%H%M")
            nome_arquivo_excel = os.path.join(DIR_RELATORIOS, f"Vistoria_{normalizar_codigo(dados_cli['cliente'])}_{dt_hoje}.xlsx")

            buffer_excel_vistoria = io.BytesIO()
            with pd.ExcelWriter(buffer_excel_vistoria, engine="openpyxl") as writer:
                pd.DataFrame(dados_elevadores_cad).to_excel(writer, sheet_name="Elevadores_Identificacao", index=False)
                df_preventiva_matriz_completa.to_excel(writer, sheet_name="Preventiva_Matriz_Integral", index=False)
                df_normas_completo.to_excel(writer, sheet_name="Normas_Checklist_Integral", index=False)
                df_orcar.to_excel(writer, sheet_name="Acoes_Orcamento", index=False)
                df_preventivo.to_excel(writer, sheet_name="Acoes_Preventiva", index=False)
                df_tecnico.to_excel(writer, sheet_name="Acoes_Tecnico", index=False)
                df_mecanica.to_excel(writer, sheet_name="Acoes_Mecanica", index=False)

            conteudo_excel_bytes = buffer_excel_vistoria.getvalue()
            excel_b64 = base64.b64encode(conteudo_excel_bytes).decode('utf-8')

            try:
                with open(nome_arquivo_excel, "wb") as f_local:
                    f_local.write(conteudo_excel_bytes)
            except Exception:
                pass

            data_str = datetime.datetime.now().strftime("%d/%m/%Y %H:%M")
            engine_salvar = get_engine()
            with engine_salvar.begin() as conn:
                conn.execute(text('''
                    UPDATE rotas_ativas 
                    SET status_vistoria = 'Concluído', data_baixa = :data, vistoriador = :vist, obs = :obs 
                    WHERE id = :id
                '''), {
                    "data": data_str,
                    "vist": nome_vistoriador,
                    "obs": "Vistoria técnica preenchida via aplicativo.",
                    "id": id_ativa
                })

                conn.execute(text('''
                    INSERT INTO relatorios_vistorias
                    (id_rota, cliente, endereco, data_vistoria, vistoriador, caminho_arquivo, qtd_normas_orc, qtd_pendencias_sup, arquivo_binario)
                    VALUES (:id, :cli, :end, :data, :vist, :arq, :normas, :sup, :b64)
                    ON CONFLICT (id_rota) DO UPDATE SET
                        data_vistoria = EXCLUDED.data_vistoria,
                        vistoriador = EXCLUDED.vistoriador,
                        caminho_arquivo = EXCLUDED.caminho_arquivo,
                        qtd_normas_orc = EXCLUDED.qtd_normas_orc,
                        qtd_pendencias_sup = EXCLUDED.qtd_pendencias_sup,
                        arquivo_binario = EXCLUDED.arquivo_binario
                '''), {
                    "id": id_ativa,
                    "cli": dados_cli['cliente'],
                    "end": dados_cli['endereco'],
                    "data": data_str,
                    "vist": nome_vistoriador,
                    "arq": nome_arquivo_excel,
                    "normas": len(df_orcar),
                    "sup": len(df_tecnico) + len(df_preventivo) + len(df_mecanica),
                    "b64": excel_b64
                })

            st.success("✅ Vistoria finalizada e transmitida com sucesso para a nuvem!")
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
            pontos_validos = df_dia[
                df_dia['lat'].notna() &
                df_dia['lon'].notna() &
                ~((df_dia['lat'].round(3) == -23.550) & (df_dia['lon'].round(3) == -46.633))
            ].copy()

            if not pontos_validos.empty:
                centro_lat = float(pontos_validos['lat'].mean())
                centro_lon = float(pontos_validos['lon'].mean())
            elif gps_ativo and lat_user and lon_user:
                centro_lat, centro_lon = lat_user, lon_user
            else:
                centro_lat, centro_lon = -23.5505, -46.6333

            mapa = folium.Map(location=[centro_lat, centro_lon], zoom_start=12, tiles="OpenStreetMap")

            coords = []
            if gps_ativo and lat_user and lon_user:
                coords.append([lat_user, lon_user])
                folium.Marker([lat_user, lon_user], popup="Sua Posição", tooltip="Você está aqui",
                              icon=folium.Icon(color="purple", icon="user")).add_to(mapa)

            for _, r_mapa in pontos_validos.iterrows():
                p_coord = [float(r_mapa['lat']), float(r_mapa['lon'])]
                coords.append(p_coord)
                cor = "green" if r_mapa['status_vistoria'] == "Concluído" else (
                    "red" if r_mapa['status_vistoria'] == "Improdutiva" else "blue")
                folium.Marker(p_coord,
                              popup=f"#{r_mapa['ordem_parada']} - {r_mapa['cliente']}<br>{r_mapa['endereco']}",
                              tooltip=f"#{r_mapa['ordem_parada']} {r_mapa['cliente']}",
                              icon=folium.Icon(color=cor)).add_to(mapa)

            if len(coords) > 1:
                folium.PolyLine(coords, color="#1B365D", weight=4, dash_array="6, 8").add_to(mapa)

            st_folium(mapa, width="100%", height=350, returned_objects=[])

            total_sem_coord = len(df_dia) - len(pontos_validos)
            if total_sem_coord > 0:
                st.caption(
                    f"ℹ️ {total_sem_coord} local(is) sem GPS exato não foram desenhados no mapa acima. "
                    f"Para navegar até eles, utilize diretamente o botão **🚗 Waze** do card da parada correspondente."
                )

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

            waze_link_corrigido = gerar_link_waze(
                row_parada.get('endereco'),
                row_parada.get('cep'),
                row_parada.get('lat'),
                row_parada.get('lon')
            )

            b1, b2, b3 = st.columns([1, 1, 1])
            with b1:
                st.link_button("🚗 Waze", waze_link_corrigido, use_container_width=True)
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
                            engine_bx = get_engine()
                            with engine_bx.begin() as conn:
                                conn.execute(text('''
                                    UPDATE rotas_ativas 
                                    SET status_vistoria = :st, data_baixa = :dt, vistoriador = :tec, obs = :obs 
                                    WHERE id = :id
                                '''), {
                                    "st": n_status,
                                    "dt": dt_agora,
                                    "tec": tec_resp,
                                    "obs": obs_baixa,
                                    "id": row_parada['id']
                                })
                            st.success("Baixa realizada!")
                            st.rerun()