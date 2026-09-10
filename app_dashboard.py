import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import openpyxl
import datetime
import os

# ==============================================================================
# CONFIGURAÇÃO DA PÁGINA
# ==============================================================================
st.set_page_config(
    page_title="Dashboard Taborda | Investimentos",
    page_icon="📈",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ==============================================================================
# CSS EXECUTIVO PERSONALIZADO (REMOÇÃO DO DEPLOY, CABEÇALHO E VISUAL ELEGANTE)
# ==============================================================================
st.markdown("""
<style>
    /* Ocultar barra superior do Streamlit, botão Deploy e menus desnecessários */
    header[data-testid="stHeader"] {
        display: none !important;
        visibility: hidden !important;
        height: 0 !important;
    }
    .stDeployButton {
        display: none !important;
        visibility: hidden !important;
    }
    [data-testid="stToolbar"] {
        display: none !important;
        visibility: hidden !important;
    }
    #MainMenu {
        visibility: hidden !important;
        display: none !important;
    }
    footer {
        visibility: hidden !important;
        display: none !important;
    }
    
    /* Ajuste de espaçamento para visual em tela cheia de alto padrão */
    .block-container {
        padding-top: 1.2rem !important;
        padding-bottom: 2rem !important;
        padding-left: 2rem !important;
        padding-right: 2rem !important;
        max-width: 98% !important;
    }

    /* Cartões de KPIs Executivos com visual Dark Moderno */
    .kpi-container {
        background: linear-gradient(135deg, #1E293B 0%, #0F172A 100%);
        border: 1px solid #334155;
        border-radius: 14px;
        padding: 16px 18px;
        box-shadow: 0 4px 15px -1px rgba(0, 0, 0, 0.2);
        transition: transform 0.2s ease, border-color 0.2s ease;
    }
    .kpi-container:hover {
        transform: translateY(-2px);
        border-color: #3B82F6;
    }
    .kpi-title {
        font-size: 0.72rem;
        font-weight: 700;
        color: #94A3B8;
        text-transform: uppercase;
        letter-spacing: 0.06em;
        margin-bottom: 4px;
    }
    .kpi-value {
        font-size: 1.6rem;
        font-weight: 900;
        color: #F8FAFC;
        letter-spacing: -0.02em;
    }
    .kpi-badge-green {
        display: inline-flex;
        align-items: center;
        padding: 2px 8px;
        background: rgba(16, 185, 129, 0.15);
        color: #10B981;
        border: 1px solid rgba(16, 185, 129, 0.3);
        border-radius: 6px;
        font-size: 0.75rem;
        font-weight: 700;
        margin-top: 6px;
    }
    .kpi-badge-blue {
        display: inline-flex;
        align-items: center;
        padding: 2px 8px;
        background: rgba(59, 130, 246, 0.15);
        color: #60A5FA;
        border: 1px solid rgba(59, 130, 246, 0.3);
        border-radius: 6px;
        font-size: 0.75rem;
        font-weight: 700;
        margin-top: 6px;
    }
    .kpi-badge-amber {
        display: inline-flex;
        align-items: center;
        padding: 2px 8px;
        background: rgba(245, 158, 11, 0.15);
        color: #FBBF24;
        border: 1px solid rgba(245, 158, 11, 0.3);
        border-radius: 6px;
        font-size: 0.75rem;
        font-weight: 700;
        margin-top: 6px;
    }
    .kpi-badge-purple {
        display: inline-flex;
        align-items: center;
        padding: 2px 8px;
        background: rgba(167, 139, 250, 0.15);
        color: #A78BFA;
        border: 1px solid rgba(167, 139, 250, 0.3);
        border-radius: 6px;
        font-size: 0.75rem;
        font-weight: 700;
        margin-top: 6px;
    }
    .kpi-badge-yellow {
        display: inline-flex;
        align-items: center;
        padding: 2px 8px;
        background: rgba(252, 211, 77, 0.15);
        color: #FCD34D;
        border: 1px solid rgba(252, 211, 77, 0.3);
        border-radius: 6px;
        font-size: 0.75rem;
        font-weight: 700;
        margin-top: 6px;
    }
    .kpi-badge-white {
        display: inline-flex;
        align-items: center;
        padding: 2px 8px;
        background: rgba(248, 250, 252, 0.15);
        color: #F8FAFC;
        border: 1px solid rgba(248, 250, 252, 0.3);
        border-radius: 6px;
        font-size: 0.75rem;
        font-weight: 700;
        margin-top: 6px;
    }
    
    /* Caixa de taxas de mercado */
    .rate-pill {
        display: inline-block;
        padding: 4px 12px;
        background: #1E293B;
        border: 1px solid #334155;
        border-radius: 20px;
        font-size: 0.78rem;
        color: #CBD5E1;
        margin-right: 6px;
        margin-bottom: 6px;
    }
    .rate-pill strong {
        color: #F8FAFC;
    }
    
    /* Melhoria Visual das Abas (Tabs) */
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
        background-color: transparent;
    }
    .stTabs [data-baseweb="tab"] {
        background-color: #1E293B;
        border-radius: 8px 8px 0px 0px;
        border: 1px solid #334155;
        border-bottom: none;
        padding: 12px 20px;
        color: #94A3B8;
        transition: all 0.2s ease-in-out;
    }
    .stTabs [data-baseweb="tab"]:hover {
        background-color: #334155;
        color: #F8FAFC;
    }
    .stTabs [aria-selected="true"] {
        background-color: #0F172A !important;
        color: #3B82F6 !important;
        border-top: 3px solid #3B82F6 !important;
        font-weight: bold;
    }
    
    /* Transformar Radio Buttons em Segmented Controls (Botões Arredondados estilo Pill) */
    .stRadio [role="radiogroup"] {
        gap: 12px;
        flex-wrap: wrap;
    }
    .stRadio [role="radiogroup"] label {
        background-color: #1E293B;
        border: 1px solid #334155;
        border-radius: 30px !important;
        padding: 8px 18px;
        cursor: pointer;
        transition: all 0.3s ease;
    }
    .stRadio [role="radiogroup"] label:hover {
        background-color: #334155;
        border-color: #3B82F6;
    }
    .stRadio [role="radiogroup"] label:has(input[aria-checked="true"]) {
        background-color: #2563EB !important;
        border-color: #60A5FA !important;
        box-shadow: 0 4px 10px rgba(37, 99, 235, 0.3);
    }
    .stRadio [role="radiogroup"] label:has(input[aria-checked="true"]) p {
        color: #F8FAFC !important;
        font-weight: 700;
    }
    /* Ocultar a bolinha nativa do radio */
    .stRadio [role="radiogroup"] label div[data-testid="stMarkdownContainer"] {
        margin-left: 0px !important;
    }
    .stRadio [role="radiogroup"] label span[data-baseweb="radio"] {
        display: none !important;
    }
</style>
""", unsafe_allow_html=True)

EXCEL_FILE = 'Investimentos.xlsx'

# ==============================================================================
# CARREGAMENTO DINÂMICO E ROBUSTO DE DADOS
# ==============================================================================
def load_data():
    if not os.path.exists(EXCEL_FILE):
        st.error(f"Arquivo {EXCEL_FILE} não encontrado!")
        return None, None, None, None

    wb = openpyxl.load_workbook(EXCEL_FILE, data_only=False)

    # 1. Indicadores
    ws_ind = wb['Indicadores']
    cdi = float(ws_ind['B2'].value or 0.1415)
    selic = float(ws_ind['C2'].value or 0.1425)
    ipca = float(ws_ind['D2'].value or 0.0016)
    usd = float(ws_ind['E2'].value or 5.0566)
    gbp = float(ws_ind['F2'].value or 6.7900)
    btc = float(ws_ind['G2'].value or 335675.0)
    
    meta_val = ws_ind['I2'].value
    try:
        meta = float(meta_val) if meta_val else 300000.0
    except:
        meta = 300000.0

    rates = {'BRL': 1.0, 'USD': usd, 'GBP': gbp}
    indicadores = {
        'cdi': cdi, 'selic': selic, 'ipca': ipca,
        'usd': usd, 'gbp': gbp, 'btc': btc, 'meta': meta
    }

    # 2. Cadastro
    ws_cad = wb['Cadastro']
    cad_data = {}
    for r in list(ws_cad.iter_rows(values_only=True))[1:]:
        if r[1]: # Código
            cad_data[str(r[1]).strip()] = {
                'codigo': str(r[1]).strip(),
                'instituicao': str(r[2] or '').strip(),
                'classe': str(r[3] or '').strip(),
                'subclasse': str(r[4] or '').strip(),
                'ativo': str(r[5] or r[1]).strip(),
                'ticker': str(r[6] or '').strip(),
                'moeda': str(r[9] or 'BRL').strip(),
                'obs': str(r[11] or '').strip()
            }

    # 3. Carteira Atual (Varredura DINÂMICA completa até encontrar TOTAL)
    ws_cart = wb['Carteira Atual']
    rows = []
    r = 4
    while r <= ws_cart.max_row:
        code_cell = ws_cart[f'A{r}'].value
        if not code_cell:
            r += 1
            continue
        code_str = str(code_cell).strip()
        if "TOTAL" in code_str.upper():
            break
            
        code = code_str
        cad_info = cad_data.get(code, {})
        
        # Puxa informações: se não achou em Cadastro, usa o que tiver na própria linha da Carteira Atual
        instituicao = cad_info.get('instituicao') or str(ws_cart[f'C{r}'].value or 'Outro').strip()
        classe = cad_info.get('classe') or str(ws_cart[f'D{r}'].value or 'Outro').strip()
        subclasse = cad_info.get('subclasse') or str(ws_cart[f'E{r}'].value or 'Geral').strip()
        ativo_nome = cad_info.get('ativo') or str(ws_cart[f'B{r}'].value or code).strip()
        moeda = cad_info.get('moeda') or str(ws_cart[f'F{r}'].value or 'BRL').strip()
        
        cambio = rates.get(moeda, 1.0)
        
        try:
            val_aplicado_orig = float(ws_cart[f'H{r}'].value or 0)
        except:
            val_aplicado_orig = 0.0
            
        try:
            saldo_atual_orig = float(ws_cart[f'J{r}'].value or 0)
        except:
            saldo_atual_orig = 0.0
            
        try:
            proventos_rs = float(ws_cart[f'L{r}'].value or 0)
        except:
            proventos_rs = 0.0
            
        try:
            retiradas_rs = float(ws_cart[f'S{r}'].value or 0)
        except:
            retiradas_rs = 0.0
            
        val_aplicado_rs = val_aplicado_orig * cambio
        saldo_atual_rs = saldo_atual_orig * cambio
        # Ao resgatar, reduzimos o Saldo e o Aplicado. Assim o Rendimento é preservado naturalmente.
        # Retiradas serve apenas como coluna informativa de quanto já foi sacado.
        rendimento_rs = (saldo_atual_rs - val_aplicado_rs) + proventos_rs
        rentabilidade_pct = (rendimento_rs / val_aplicado_rs) if val_aplicado_rs > 0 else 0.0
        
        obs = str(ws_cart[f'R{r}'].value or cad_info.get('obs', ''))
        
        rows.append({
            'Linha_Excel': r,
            'Código': code,
            'Ativo': ativo_nome,
            'Instituição': instituicao,
            'Classe': classe,
            'Subclasse': subclasse,
            'Moeda': moeda,
            'Câmbio': cambio,
            'Valor Aplicado Original': val_aplicado_orig,
            'Valor Aplicado (R$)': val_aplicado_rs,
            'Saldo Atual Original': saldo_atual_orig,
            'Saldo Atual (R$)': saldo_atual_rs,
            'Proventos (R$)': proventos_rs,
            'Retiradas (R$)': retiradas_rs,
            'Rendimento (R$)': rendimento_rs,
            'Rentabilidade (%)': rentabilidade_pct,
            'Observação': obs
        })
        r += 1

    df = pd.DataFrame(rows)
    total_saldo = df['Saldo Atual (R$)'].sum() if not df.empty else 0.0
    if not df.empty and total_saldo > 0:
        df['% Carteira'] = df['Saldo Atual (R$)'] / total_saldo
    else:
        df['% Carteira'] = 0.0

    # 4. Histórico
    hist_df = pd.DataFrame()
    if 'Evolução Mensal' in wb.sheetnames:
        ws_evo = wb['Evolução Mensal']
        evo_rows = []
        for row in list(ws_evo.iter_rows(values_only=True))[1:]:
            if row[0] and row[6] is not None:
                evo_rows.append({
                    'Data': row[0],
                    'Competência': row[1],
                    'Instituição': row[2],
                    'Classe': row[3],
                    'Ativo': row[4],
                    'Código': row[5],
                    'Saldo': float(row[6] or 0),
                    'Aporte': float(row[7] or 0),
                    'Resgate': float(row[8] or 0),
                    'Observação': row[10] or ''
                })
        hist_df = pd.DataFrame(evo_rows)

    return df, indicadores, cad_data, hist_df

# ==============================================================================
# ATUALIZAÇÃO / APORTE EM ATIVO EXISTENTE
# ==============================================================================
def update_existing_asset(row_excel, novo_aplicado, novo_saldo, novos_proventos, retiradas=0.0):
    wb = openpyxl.load_workbook(EXCEL_FILE)
    ws_cart = wb['Carteira Atual']
    
    ws_cart[f'H{row_excel}'] = float(novo_aplicado)
    ws_cart[f'J{row_excel}'] = float(novo_saldo)
    ws_cart[f'L{row_excel}'] = float(novos_proventos)
    
    # Atualizar Retiradas e fórmula
    ws_cart[f'S{row_excel}'] = float(retiradas)
    ws_cart[f'M{row_excel}'] = f'=(K{row_excel}-I{row_excel})+L{row_excel}'
    
    wb.save(EXCEL_FILE)

def delete_asset(row_excel):
    wb = openpyxl.load_workbook(EXCEL_FILE)
    ws_cart = wb['Carteira Atual']
    ws_cart.delete_rows(row_excel)
    wb.save(EXCEL_FILE)

# ==============================================================================
# CADASTRO DE NOVO ATIVO COM AGRUPAMENTO NA CLASSE E INSTITUIÇÃO
# ==============================================================================
def add_new_asset_grouped(codigo, ativo, instituicao, classe, subclasse, moeda, valor_aplicado, saldo_atual, obs):
    wb = openpyxl.load_workbook(EXCEL_FILE)
    ws_cad = wb['Cadastro']
    ws_cart = wb['Carteira Atual']
    
    # 1. Adicionar ao Cadastro
    next_cad_row = ws_cad.max_row + 1
    new_id = next_cad_row - 1
    ws_cad.cell(row=next_cad_row, column=1, value=new_id)
    ws_cad.cell(row=next_cad_row, column=2, value=codigo)
    ws_cad.cell(row=next_cad_row, column=3, value=instituicao)
    ws_cad.cell(row=next_cad_row, column=4, value=classe)
    ws_cad.cell(row=next_cad_row, column=5, value=subclasse)
    ws_cad.cell(row=next_cad_row, column=6, value=ativo)
    ws_cad.cell(row=next_cad_row, column=7, value=codigo)
    ws_cad.cell(row=next_cad_row, column=8, value=valor_aplicado)
    ws_cad.cell(row=next_cad_row, column=9, value=datetime.date.today())
    ws_cad.cell(row=next_cad_row, column=10, value=moeda)
    ws_cad.cell(row=next_cad_row, column=11, value="Sim")
    ws_cad.cell(row=next_cad_row, column=12, value=obs)
    
    # 2. Localizar linha do TOTAL em Carteira Atual
    total_row = None
    for r in range(4, ws_cart.max_row + 10):
        val = ws_cart[f'A{r}'].value
        if val and "TOTAL" in str(val).upper():
            total_row = r
            break
    if not total_row:
        total_row = ws_cart.max_row + 1

    # Inserir nova linha exatamente antes do TOTAL
    ws_cart.insert_rows(total_row)
    nr = total_row
    
    ws_cart[f'A{nr}'] = codigo
    ws_cart[f'B{nr}'] = f'=IFERROR(VLOOKUP(A{nr}, Cadastro!$B$2:$G$100, 5, FALSE), "{ativo}")'
    ws_cart[f'C{nr}'] = f'=IFERROR(VLOOKUP(A{nr}, Cadastro!$B$2:$C$100, 2, FALSE), "{instituicao}")'
    ws_cart[f'D{nr}'] = f'=IFERROR(VLOOKUP(A{nr}, Cadastro!$B$2:$D$100, 3, FALSE), "{classe}")'
    ws_cart[f'E{nr}'] = f'=IFERROR(VLOOKUP(A{nr}, Cadastro!$B$2:$E$100, 4, FALSE), "{subclasse}")'
    ws_cart[f'F{nr}'] = f'=IFERROR(VLOOKUP(A{nr}, Cadastro!$B$2:$J$100, 9, FALSE), "{moeda}")'
    ws_cart[f'G{nr}'] = f'=IF(F{nr}="USD", Indicadores!$E$2, IF(F{nr}="GBP", Indicadores!$F$2, 1))'
    ws_cart[f'H{nr}'] = float(valor_aplicado)
    ws_cart[f'I{nr}'] = f'=H{nr}*G{nr}'
    ws_cart[f'J{nr}'] = float(saldo_atual)
    ws_cart[f'K{nr}'] = f'=J{nr}*G{nr}'
    ws_cart[f'L{nr}'] = 0.0
    ws_cart[f'M{nr}'] = f'=(K{nr}-I{nr})+L{nr}'
    ws_cart[f'N{nr}'] = f'=IF(I{nr}>0, M{nr}/I{nr}, 0)'
    ws_cart[f'O{nr}'] = f'=IF($K${total_row+1}>0, K{nr}/$K${total_row+1}, 0)'
    ws_cart[f'P{nr}'] = 0.03
    ws_cart[f'Q{nr}'] = f'=O{nr}-P{nr}'
    ws_cart[f'R{nr}'] = obs
    ws_cart[f'S{nr}'] = 0.0
    
    # Atualizar fórmula da linha de TOTAL
    new_total_row = total_row + 1
    ws_cart[f'I{new_total_row}'] = f'=SUM(I4:I{new_total_row-1})'
    ws_cart[f'K{new_total_row}'] = f'=SUM(K4:K{new_total_row-1})'
    ws_cart[f'L{new_total_row}'] = f'=SUM(L4:L{new_total_row-1})'
    ws_cart[f'M{new_total_row}'] = f'=SUM(M4:M{new_total_row-1})'
    ws_cart[f'S{new_total_row}'] = f'=SUM(S4:S{new_total_row-1})'
    
    wb.save(EXCEL_FILE)

# ==============================================================================
# CARREGAR DADOS ATUAIS
# ==============================================================================
df, ind, cad, hist_df = load_data()

if df is None or df.empty:
    st.error("Não foi possível carregar a base de investimentos.")
    st.stop()

# ==============================================================================
# SIDEBAR COM FILTROS RÁPIDOS
# ==============================================================================
with st.sidebar:
    st.markdown("### 🎛️ Filtros Rápidos")
    
    todas_instituicoes = sorted(df['Instituição'].unique())
    filtro_inst = st.multiselect("Instituição / Banco", options=todas_instituicoes, default=todas_instituicoes)

    todas_classes = sorted(df['Classe'].unique())
    filtro_classe = st.multiselect("Classe de Ativo", options=todas_classes, default=todas_classes)

    todas_moedas = sorted(df['Moeda'].unique())
    filtro_moeda = st.multiselect("Moeda", options=todas_moedas, default=todas_moedas)

    st.markdown("---")
    st.markdown("### 📈 Cotações & Benchmarks")
    st.markdown(f"""
    <div style="font-size: 0.85rem; line-height: 1.8;">
        <div>🔹 <b>CDI:</b> {ind['cdi']*100:.2f}% a.a.</div>
        <div>🔹 <b>Selic:</b> {ind['selic']*100:.2f}% a.a.</div>
        <div>🔹 <b>Dólar:</b> R$ {ind['usd']:.2f}</div>
        <div>🔹 <b>Libra:</b> R$ {ind['gbp']:.2f}</div>
        <div>🔹 <b>Bitcoin:</b> R$ {ind['btc']:,.0f}</div>
    </div>
    """, unsafe_allow_html=True)
    st.markdown("---")
    st.caption(f"💾 Arquivo: `{EXCEL_FILE}`")

# Filtrar dataframe com base na seleção
dff = df[
    (df['Instituição'].isin(filtro_inst)) &
    (df['Classe'].isin(filtro_classe)) &
    (df['Moeda'].isin(filtro_moeda))
]

# ==============================================================================
# CABEÇALHO EXECUTIVO (SEM BARRAS DE DEPLOY, 100% LIMPO)
# ==============================================================================
col_head1, col_head2, col_head3 = st.columns([5, 3, 1])

with col_head1:
    st.markdown("""
    <div style="padding-bottom: 4px;">
        <h1 style="font-size: 1.9rem; font-weight: 900; color: #F8FAFC; margin: 0; padding: 0; letter-spacing: -0.02em;">
            PORTFÓLIO DE INVESTIMENTOS | TABORDA
        </h1>
        <p style="font-size: 0.88rem; color: #94A3B8; margin-top: 3px;">
            Painel Executivo de Alocação Patrimonial, Performance e Rebalanceamento
        </p>
    </div>
    """, unsafe_allow_html=True)

with col_head2:
    st.markdown(f"""
    <div style="text-align: right; padding-top: 6px;">
        <span class="rate-pill">CDI: <strong>{ind['cdi']*100:.2f}%</strong></span>
        <span class="rate-pill">USD: <strong>R$ {ind['usd']:.2f}</strong></span><br>
        <span class="rate-pill">GBP: <strong>R$ {ind['gbp']:.2f}</strong></span>
        <span class="rate-pill">BTC: <strong>R$ {ind['btc']/1000:.0f}k</strong></span>
    </div>
    """, unsafe_allow_html=True)

with col_head3:
    st.markdown("<div style='padding-top: 12px;'></div>", unsafe_allow_html=True)
    if st.button("🔄 Atualizar", use_container_width=True, type="primary"):
        st.rerun()

# ==============================================================================
# BARRA DE KPIS EXECUTIVOS
# ==============================================================================
total_patrimonio = dff['Saldo Atual (R$)'].sum()
total_aplicado = dff['Valor Aplicado (R$)'].sum()
total_rendimento = dff['Rendimento (R$)'].sum()
total_retiradas = dff['Retiradas (R$)'].sum()
rentabilidade_global = (total_rendimento / total_aplicado) if total_aplicado > 0 else 0.0
total_proventos = dff['Proventos (R$)'].sum()

# Meta definition
meta_global = st.session_state.get('meta_patrimonio', ind.get('meta', 300000.0))
pct_meta = (total_patrimonio / meta_global) * 100 if meta_global > 0 else 0.0

kpi1, kpi2, kpi3, kpi4, kpi5, kpi6 = st.columns(6)

with kpi1:
    st.markdown(f"""
    <div class="kpi-container">
        <div class="kpi-title">Patrimônio Total</div>
        <div class="kpi-value" style="color:#60A5FA;">R$ {total_patrimonio:,.2f}</div>
        <div class="kpi-badge-blue">+{rentabilidade_global*100:.2f}% Retorno Total</div>
    </div>
    """, unsafe_allow_html=True)

with kpi2:
    st.markdown(f"""
    <div class="kpi-container">
        <div class="kpi-title">Total Investido</div>
        <div class="kpi-value">R$ {total_aplicado:,.2f}</div>
        <div class="kpi-badge-white">100.00% Base de Custo</div>
    </div>
    """, unsafe_allow_html=True)

with kpi3:
    rent_liq_pct = (total_rendimento / total_aplicado * 100) if total_aplicado > 0 else 0.0
    st.markdown(f"""
    <div class="kpi-container">
        <div class="kpi-title">Rendimento Líquido</div>
        <div class="kpi-value" style="color:#34D399;">R$ {total_rendimento:,.2f}</div>
        <div class="kpi-badge-green">+{rent_liq_pct:.2f}% de Lucro</div>
    </div>
    """, unsafe_allow_html=True)

with kpi4:
    ret_pct = (total_retiradas / total_aplicado * 100) if total_aplicado > 0 else 0.0
    st.markdown(f"""
    <div class="kpi-container">
        <div class="kpi-title">Lucro Realizado</div>
        <div class="kpi-value" style="color:#FCD34D;">R$ {total_retiradas:,.2f}</div>
        <div class="kpi-badge-yellow">{ret_pct:.2f}% Resgatado</div>
    </div>
    """, unsafe_allow_html=True)

with kpi5:
    yield_pct = (total_proventos / total_aplicado * 100) if total_aplicado > 0 else 0.0
    st.markdown(f"""
    <div class="kpi-container">
        <div class="kpi-title">Proventos (Yield)</div>
        <div class="kpi-value" style="color:#A78BFA;">R$ {total_proventos:,.2f}</div>
        <div class="kpi-badge-purple">+{yield_pct:.2f}% em Dividendos</div>
    </div>
    """, unsafe_allow_html=True)

with kpi6:
    st.markdown(f"""
    <div class="kpi-container">
        <div class="kpi-title">Progresso da Meta</div>
        <div class="kpi-value">{pct_meta:.1f}%</div>
        <div style="width: 100%; background-color: #334155; border-radius: 4px; height: 6px; margin-top: 8px;">
            <div style="width: {min(pct_meta, 100)}%; background-color: #3B82F6; height: 100%; border-radius: 4px;"></div>
        </div>
    </div>
    """, unsafe_allow_html=True)

st.markdown("<div style='height: 14px;'></div>", unsafe_allow_html=True)

# ==============================================================================
# OS 4 GRÁFICOS MAIS IMPORTANTES "EM UMA TELA" (GRID 2x2)
# ==============================================================================
chart_row1_col1, chart_row1_col2 = st.columns(2)

# Tema e Cores Plotly
PALETA_CORES = ['#3B82F6', '#10B981', '#8B5CF6', '#F59E0B', '#EC4899', '#06B6D4', '#64748B']

with chart_row1_col1:
    with st.container(border=True):
        # Gráfico 1: Alocação por Classe de Ativo (Donut Moderno)
        df_classe = dff.groupby('Classe')['Saldo Atual (R$)'].sum().reset_index()
        df_classe = df_classe.sort_values(by='Saldo Atual (R$)', ascending=False)
        
        fig_classe = px.pie(
            df_classe,
            values='Saldo Atual (R$)',
            names='Classe',
            hole=0.55,
            color_discrete_sequence=PALETA_CORES,
            title="<b>1. Alocação por Classe de Ativo</b>"
        )
        fig_classe.update_traces(
            textposition='inside',
            textinfo='percent',
            hovertemplate="<b>%{label}</b><br>Saldo: R$ %{value:,.2f}<br>Participação: %{percent}<extra></extra>",
            hoverlabel=dict(bgcolor="#0F172A", bordercolor="#3B82F6", font=dict(color="white", size=13)),
            marker=dict(line=dict(color='#0F172A', width=3))
        )
        # Adicionar o valor total no centro do gráfico de rosca
        total_grafico1 = df_classe['Saldo Atual (R$)'].sum()
        fig_classe.add_annotation(
            text=f"<span style='font-size:11px;color:#94A3B8'>Patrimônio</span><br><b style='font-size:16px;color:#F8FAFC'>R$ {total_grafico1/1000:,.0f}k</b>",
            x=0.5, y=0.5, showarrow=False
        )
        fig_classe.update_layout(
            template="plotly_dark",
            paper_bgcolor='rgba(0,0,0,0)',
            plot_bgcolor='rgba(0,0,0,0)',
            font=dict(color="#F8FAFC", size=12),
            margin=dict(t=45, b=20, l=15, r=15),
            height=330,
            showlegend=True,
            legend=dict(orientation="v", yanchor="middle", y=0.5, xanchor="left", x=0.9)
        )
        st.plotly_chart(fig_classe, use_container_width=True)

with chart_row1_col2:
    with st.container(border=True):
        # Gráfico 2: Posição por Instituição / Corretora (Barras Horizontais)
        df_inst = dff.groupby('Instituição')['Saldo Atual (R$)'].sum().reset_index()
        df_inst = df_inst.sort_values(by='Saldo Atual (R$)', ascending=True)
        
        fig_inst = px.bar(
            df_inst,
            x='Saldo Atual (R$)',
            y='Instituição',
            orientation='h',
            title="<b>2. Posição por Instituição / Banco (R$)</b>",
            color='Saldo Atual (R$)',
            color_continuous_scale=['#1E3A8A', '#2563EB', '#60A5FA'],
            text='Saldo Atual (R$)'
        )
        fig_inst.update_traces(
            texttemplate='R$ %{x:,.0f}',
            textposition='outside',
            hovertemplate="<b>%{y}</b><br>Saldo Total: R$ %{x:,.2f}<extra></extra>",
            hoverlabel=dict(bgcolor="#0F172A", bordercolor="#3B82F6", font=dict(color="white", size=13))
        )
        fig_inst.update_layout(
            template="plotly_dark",
            paper_bgcolor='rgba(0,0,0,0)',
            plot_bgcolor='rgba(0,0,0,0)',
            font=dict(color="#F8FAFC", size=12),
            margin=dict(t=45, b=20, l=15, r=40),
            height=330,
            coloraxis_showscale=False,
            xaxis=dict(showgrid=False, showticklabels=False, zeroline=False),
            yaxis=dict(showgrid=False)
        )
        st.plotly_chart(fig_inst, use_container_width=True)

chart_row2_col1, chart_row2_col2 = st.columns(2)

with chart_row2_col1:
    with st.container(border=True):
        # Gráfico 3: Rentabilidade (%) por Classe de Ativo
        df_rent_cls = dff.groupby('Classe').agg({
            'Valor Aplicado (R$)': 'sum',
            'Rendimento (R$)': 'sum'
        }).reset_index()
        df_rent_cls['Rentabilidade (%)'] = (df_rent_cls['Rendimento (R$)'] / df_rent_cls['Valor Aplicado (R$)']) * 100
        df_rent_cls = df_rent_cls.sort_values(by='Rentabilidade (%)', ascending=True)
        
        fig_rent = px.bar(
            df_rent_cls,
            x='Rentabilidade (%)',
            y='Classe',
            orientation='h',
            title="<b>3. Rentabilidade (%) por Classe</b>",
            color='Rentabilidade (%)',
            color_continuous_scale=['#10B981', '#34D399', '#6EE7B7'],
            text=df_rent_cls['Rentabilidade (%)'].apply(lambda x: f"{x:.1f}%")
        )
        fig_rent.update_traces(
            textposition='outside',
            hovertemplate="<b>%{y}</b><br>Rentabilidade: %{x:.2f}%<extra></extra>",
            hoverlabel=dict(bgcolor="#0F172A", bordercolor="#10B981", font=dict(color="white", size=13))
        )
        fig_rent.update_layout(
            template="plotly_dark",
            paper_bgcolor='rgba(0,0,0,0)',
            plot_bgcolor='rgba(0,0,0,0)',
            font=dict(color="#F8FAFC", size=12),
            margin=dict(t=45, b=20, l=15, r=40),
            height=310,
            coloraxis_showscale=False,
            xaxis=dict(showgrid=False, showticklabels=False, zeroline=False),
            yaxis=dict(showgrid=False)
        )
        st.plotly_chart(fig_rent, use_container_width=True)

with chart_row2_col2:
    with st.container(border=True):
        # Gráfico 4: Exposição Cambial (Moedas BRL vs USD vs GBP)
        df_moeda = dff.groupby('Moeda')['Saldo Atual (R$)'].sum().reset_index()
        fig_moeda = px.pie(
            df_moeda,
            values='Saldo Atual (R$)',
            names='Moeda',
            hole=0.55,
            title="<b>4. Exposição Cambial (Moedas)</b>",
            color='Moeda',
            color_discrete_map={'BRL': '#2563EB', 'USD': '#10B981', 'GBP': '#8B5CF6'}
        )
        fig_moeda.update_traces(
            textposition='inside',
            textinfo='percent',
            hovertemplate="<b>Moeda: %{label}</b><br>Total em R$: R$ %{value:,.2f}<br>Fração: %{percent}<extra></extra>",
            hoverlabel=dict(bgcolor="#0F172A", bordercolor="#3B82F6", font=dict(color="white", size=13)),
            marker=dict(line=dict(color='#0F172A', width=3))
        )
        
        # Adicionar total no centro
        total_grafico4 = df_moeda['Saldo Atual (R$)'].sum()
        fig_moeda.add_annotation(
            text=f"<span style='font-size:11px;color:#94A3B8'>Total</span><br><b style='font-size:16px;color:#F8FAFC'>100%</b>",
            x=0.5, y=0.5, showarrow=False
        )

        fig_moeda.update_layout(
            template="plotly_dark",
            paper_bgcolor='rgba(0,0,0,0)',
            plot_bgcolor='rgba(0,0,0,0)',
            font=dict(color="#F8FAFC", size=12),
            margin=dict(t=45, b=20, l=15, r=15),
            height=310,
            showlegend=True,
            legend=dict(orientation="v", yanchor="middle", y=0.5, xanchor="left", x=0.9)
        )
        st.plotly_chart(fig_moeda, use_container_width=True)

st.markdown("<div style='height: 10px;'></div>", unsafe_allow_html=True)

# ==============================================================================
# ABAS INFERIORES: TABELA AGRUPADA, LANÇAMENTOS E METAS
# ==============================================================================
tab_tabela, tab_lancamento, tab_metas = st.tabs([
    "📋 Carteira de Ativos (Visão Agrupada)",
    "➕ Lançar Aporte & Novo Ativo",
    "🎯 Rebalanceamento da Carteira (Metas 300k)"
])

# ------------------------------------------------------------------------------
# ABA 1: TABELA DETALHADA E AGRUPADA POR CLASSE
# ------------------------------------------------------------------------------
with tab_tabela:
    with st.container(border=True):
        col_tb1, col_tb2 = st.columns([3, 1])
        with col_tb1:
            modo_visao = st.radio(
                "Modo de Exibição:",
                options=["Por Classe de Ativo (Agrupado)", "Por Instituição (Agrupado)", "Lista Completa Detalhada"],
                horizontal=True
            )
        with col_tb2:
            termo_busca = st.text_input("🔍 Buscar ativo por código ou nome:", "")

    df_view = dff.copy()
    if termo_busca:
        df_view = df_view[
            df_view['Código'].str.contains(termo_busca, case=False, na=False) |
            df_view['Ativo'].str.contains(termo_busca, case=False, na=False) |
            df_view['Instituição'].str.contains(termo_busca, case=False, na=False)
        ]

    if modo_visao == "Por Classe de Ativo (Agrupado)":
        classes_presentes = sorted(df_view['Classe'].unique())
        for cls in classes_presentes:
            df_c = df_view[df_view['Classe'] == cls]
            sub_total_aplicado = df_c['Valor Aplicado (R$)'].sum()
            sub_total_saldo = df_c['Saldo Atual (R$)'].sum()
            sub_proventos = df_c['Proventos (R$)'].sum()
            sub_retiradas = df_c['Retiradas (R$)'].sum()
            sub_rendimento = df_c['Rendimento (R$)'].sum()
            sub_rent = (sub_rendimento / sub_total_aplicado * 100) if sub_total_aplicado > 0 else 0.0
            sub_share = (sub_total_saldo / total_patrimonio * 100) if total_patrimonio > 0 else 0.0
            
            with st.expander(f"📁 **{cls}** — {len(df_c)} ativos | Saldo: **R$ {sub_total_saldo:,.2f}** ({sub_share:.1f}%) | Retorno: **{sub_rent:+.2f}%**", expanded=True):
                sub_df = df_c[['Código', 'Ativo', 'Instituição', 'Subclasse', 'Moeda', 'Valor Aplicado (R$)', 'Saldo Atual (R$)', 'Proventos (R$)', 'Retiradas (R$)', 'Rendimento (R$)', 'Rentabilidade (%)', '% Carteira', 'Observação']].copy()
                
                # Adicionar linha de Total
                total_row = pd.DataFrame([{
                    'Código': 'TOTAL',
                    'Ativo': '',
                    'Instituição': '',
                    'Subclasse': '',
                    'Moeda': '',
                    'Valor Aplicado (R$)': sub_total_aplicado,
                    'Saldo Atual (R$)': sub_total_saldo,
                    'Proventos (R$)': sub_proventos,
                    'Retiradas (R$)': sub_retiradas,
                    'Rendimento (R$)': sub_rendimento,
                    'Rentabilidade (%)': sub_rent / 100.0,
                    '% Carteira': sub_share / 100.0,
                    'Observação': ''
                }])
                sub_df = pd.concat([sub_df, total_row], ignore_index=True)

                st.dataframe(
                    sub_df.style.format({
                        'Valor Aplicado (R$)': 'R$ {:,.2f}',
                        'Saldo Atual (R$)': 'R$ {:,.2f}',
                        'Proventos (R$)': 'R$ {:,.2f}',
                        'Retiradas (R$)': 'R$ {:,.2f}',
                        'Rendimento (R$)': 'R$ {:,.2f}',
                        'Rentabilidade (%)': '{:.2%}',
                        '% Carteira': '{:.2%}'
                    }),
                    use_container_width=True
                )

    elif modo_visao == "Por Instituição (Agrupado)":
        inst_presentes = sorted(df_view['Instituição'].unique())
        for inst in inst_presentes:
            df_i = df_view[df_view['Instituição'] == inst]
            sub_total_aplicado = df_i['Valor Aplicado (R$)'].sum()
            sub_total_saldo = df_i['Saldo Atual (R$)'].sum()
            sub_proventos = df_i['Proventos (R$)'].sum()
            sub_retiradas = df_i['Retiradas (R$)'].sum()
            sub_rendimento = df_i['Rendimento (R$)'].sum()
            sub_rent = (sub_rendimento / sub_total_aplicado * 100) if sub_total_aplicado > 0 else 0.0
            sub_share = (sub_total_saldo / total_patrimonio * 100) if total_patrimonio > 0 else 0.0
            
            with st.expander(f"🏛️ **{inst}** — {len(df_i)} ativos | Saldo: **R$ {sub_total_saldo:,.2f}** ({sub_share:.1f}%) | Retorno: **{sub_rent:+.2f}%**", expanded=True):
                sub_df = df_i[['Código', 'Ativo', 'Classe', 'Subclasse', 'Moeda', 'Valor Aplicado (R$)', 'Saldo Atual (R$)', 'Proventos (R$)', 'Retiradas (R$)', 'Rendimento (R$)', 'Rentabilidade (%)', '% Carteira', 'Observação']].copy()
                
                # Adicionar linha de Total
                total_row = pd.DataFrame([{
                    'Código': 'TOTAL',
                    'Ativo': '',
                    'Classe': '',
                    'Subclasse': '',
                    'Moeda': '',
                    'Valor Aplicado (R$)': sub_total_aplicado,
                    'Saldo Atual (R$)': sub_total_saldo,
                    'Proventos (R$)': sub_proventos,
                    'Retiradas (R$)': sub_retiradas,
                    'Rendimento (R$)': sub_rendimento,
                    'Rentabilidade (%)': sub_rent / 100.0,
                    '% Carteira': sub_share / 100.0,
                    'Observação': ''
                }])
                sub_df = pd.concat([sub_df, total_row], ignore_index=True)

                st.dataframe(
                    sub_df.style.format({
                        'Valor Aplicado (R$)': 'R$ {:,.2f}',
                        'Saldo Atual (R$)': 'R$ {:,.2f}',
                        'Proventos (R$)': 'R$ {:,.2f}',
                        'Retiradas (R$)': 'R$ {:,.2f}',
                        'Rendimento (R$)': 'R$ {:,.2f}',
                        'Rentabilidade (%)': '{:.2%}',
                        '% Carteira': '{:.2%}'
                    }),
                    use_container_width=True
                )

    else:
        # Lista Completa
        st.dataframe(
            df_view[['Código', 'Ativo', 'Instituição', 'Classe', 'Subclasse', 'Moeda', 'Valor Aplicado (R$)', 'Saldo Atual (R$)', 'Proventos (R$)', 'Retiradas (R$)', 'Rendimento (R$)', 'Rentabilidade (%)', '% Carteira', 'Observação']].style.format({
                'Valor Aplicado (R$)': 'R$ {:,.2f}',
                'Saldo Atual (R$)': 'R$ {:,.2f}',
                'Proventos (R$)': 'R$ {:,.2f}',
                'Retiradas (R$)': 'R$ {:,.2f}',
                'Rendimento (R$)': 'R$ {:,.2f}',
                'Rentabilidade (%)': '{:.2%}',
                '% Carteira': '{:.2%}'
            }),
            use_container_width=True,
            height=450
        )

# ------------------------------------------------------------------------------
# ABA 2: LANÇAMENTOS E APORTES (MELHORIA CRÍTICA)
# ------------------------------------------------------------------------------
with tab_lancamento:
    st.markdown("### 📝 Gestão de Investimentos & Lançamentos")
    st.markdown("Escolha se deseja **lançar um aporte / atualizar a posição de um ativo existente** ou **cadastrar um novo investimento** que será agrupado automaticamente junto com a sua classe.")

    with st.container(border=True):
        tipo_acao = st.radio(
            "O que você deseja fazer?",
            options=["1️⃣ Fazer Aporte ou Atualizar Saldo (Ativo Existente)", "2️⃣ Cadastrar um Novo Ativo", "3️⃣ Registrar Retirada ou Excluir Ativo"],
            horizontal=True
        )

    # ---------------------------------------------------------
    # OPÇÃO 1: APORTE OU ATUALIZAÇÃO EM ATIVO JÁ EXISTENTE
    # ---------------------------------------------------------
    if tipo_acao.startswith("1️⃣"):
        with st.container(border=True):
            st.markdown("#### 🔄 Atualizar Posição ou Realizar Novo Aporte")
            
            opcoes_ativos = [f"{row['Código']} | {row['Ativo']} ({row['Instituição']} - {row['Classe']})" for _, row in df.iterrows()]
            escolha_ativo = st.selectbox("Selecione o Ativo para Aporte/Atualização:", options=opcoes_ativos)
            
            if escolha_ativo:
                cod_selecionado = escolha_ativo.split(" | ")[0]
                item = df[df['Código'] == cod_selecionado].iloc[0]
                
                st.info(f"📍 **Ativo Selecionado:** {item['Ativo']}  \n🏛️ **Instituição:** {item['Instituição']} | 🏷️ **Classe:** {item['Classe']} | 💵 **Moeda:** {item['Moeda']}")
                
                tipo_operacao = st.radio(
                    "Tipo de Operação:",
                    options=["Novo Aporte (Soma ao aplicado e saldo)", "Atualizar Saldo de Mercado (Nova Cotação)", "Lançar Proventos (Dividendos/Cupons)"],
                    horizontal=True
                )
    
                col_op1, col_op2, col_op3 = st.columns(3)
                
                if "Novo Aporte" in tipo_operacao:
                    with col_op1:
                        aporte_valor = st.number_input(f"Valor do Novo Aporte ({item['Moeda']}):", min_value=0.0, step=500.0, format="%.2f")
                    with col_op2:
                        st.metric("Novo Valor Aplicado Total", f"R$ {(item['Valor Aplicado Original'] + aporte_valor)*item['Câmbio']:,.2f}")
                    with col_op3:
                        st.metric("Novo Saldo Estimado", f"R$ {(item['Saldo Atual Original'] + aporte_valor)*item['Câmbio']:,.2f}")
    
                    if st.button("💾 Gravar Novo Aporte na Planilha Excel", type="primary"):
                        if aporte_valor > 0:
                            novo_ap = item['Valor Aplicado Original'] + aporte_valor
                            novo_sal = item['Saldo Atual Original'] + aporte_valor
                            update_existing_asset(int(item['Linha_Excel']), novo_ap, novo_sal, item['Proventos (R$)'], item['Retiradas (R$)'])
                            st.success(f"✅ Aporte de {item['Moeda']} {aporte_valor:,.2f} registrado com sucesso em **{item['Código']}**!")
                            st.rerun()
                        else:
                            st.warning("Informe um valor de aporte maior que zero.")
    
                elif "Atualizar Saldo" in tipo_operacao:
                    with col_op1:
                        st.metric("Saldo Atual Registrado", f"{item['Moeda']} {item['Saldo Atual Original']:,.2f}")
                    with col_op2:
                        novo_saldo_input = st.number_input(f"Novo Saldo de Mercado ({item['Moeda']}):", min_value=0.0, value=float(item['Saldo Atual Original']), step=100.0, format="%.2f")
                    with col_op3:
                        novo_rend = (novo_saldo_input - item['Valor Aplicado Original'])*item['Câmbio'] + item['Proventos (R$)'] + item['Retiradas (R$)']
                        st.metric("Novo Rendimento Líquido", f"R$ {novo_rend:,.2f}")
    
                    if st.button("💾 Atualizar Saldo na Planilha Excel", type="primary"):
                        update_existing_asset(int(item['Linha_Excel']), item['Valor Aplicado Original'], novo_saldo_input, item['Proventos (R$)'], item['Retiradas (R$)'])
                        st.success(f"✅ Saldo de **{item['Código']}** atualizado com sucesso!")
                        st.rerun()
    
                else: # Proventos
                    with col_op1:
                        st.metric("Proventos Anteriores", f"R$ {item['Proventos (R$)']:,.2f}")
                    with col_op2:
                        add_proventos = st.number_input("Valor do Provento Recebido (R$):", min_value=0.0, step=50.0, format="%.2f")
                    with col_op3:
                        st.metric("Proventos Totais Acumulados", f"R$ {item['Proventos (R$)'] + add_proventos:,.2f}")
    
                    if st.button("💾 Gravar Provento na Planilha Excel", type="primary"):
                        if add_proventos > 0:
                            novos_prov = item['Proventos (R$)'] + add_proventos
                            update_existing_asset(int(item['Linha_Excel']), item['Valor Aplicado Original'], item['Saldo Atual Original'], novos_prov, item['Retiradas (R$)'])
                            st.success(f"✅ Provento de R$ {add_proventos:,.2f} adicionado a **{item['Código']}**!")
                            st.rerun()

    # ---------------------------------------------------------
    # OPÇÃO 3: REGISTRAR RETIRADA OU EXCLUIR ATIVO
    # ---------------------------------------------------------
    elif tipo_acao.startswith("3️⃣"):
        with st.container(border=True):
            st.markdown("#### 💸 Registrar Retirada (Resgate) ou ❌ Excluir Ativo")
            
            opcoes_ativos = [f"{row['Código']} | {row['Ativo']} ({row['Instituição']} - {row['Classe']})" for _, row in df.iterrows()]
            escolha_ativo = st.selectbox("Selecione o Ativo para Retirada/Exclusão:", options=opcoes_ativos)
            
            if escolha_ativo:
                cod_selecionado = escolha_ativo.split(" | ")[0]
                item = df[df['Código'] == cod_selecionado].iloc[0]
                
                st.warning(f"📍 **Ativo Selecionado:** {item['Ativo']}  \nSaldo Atual: **{item['Moeda']} {item['Saldo Atual Original']:,.2f}** | Retiradas Acumuladas: **R$ {item['Retiradas (R$)']:,.2f}**")
                
                tipo_operacao = st.radio(
                    "Qual ação deseja realizar com este ativo?",
                    options=["Registrar Retirada Parcial / Resgate", "Excluir Ativo Definitivamente da Carteira"],
                    horizontal=True
                )
                
                if "Retirada" in tipo_operacao:
                    st.info("💡 **Dica:** Registrar uma retirada reduz o saldo atual aplicado, mas soma o valor resgatado como 'Lucro/Rendimento' para não prejudicar seu histórico de rentabilidade global.")
                    col_r1, col_r2 = st.columns(2)
                    with col_r1:
                        valor_retirada = st.number_input(f"Valor do Resgate / Retirada ({item['Moeda']}):", min_value=0.0, step=100.0, format="%.2f")
                    with col_r2:
                        novo_saldo_ret = item['Saldo Atual Original'] - valor_retirada
                        novo_aplicado_ret = item['Valor Aplicado Original'] - valor_retirada
                        st.metric("Novo Saldo Restante Estimado", f"{item['Moeda']} {novo_saldo_ret:,.2f}")
                    
                    if st.button("💸 Gravar Retirada na Planilha Excel", type="primary"):
                        if valor_retirada > 0:
                            # A retirada reduz o valor aplicado e o saldo. Em reais, soma no acumulado de retiradas para garantir o cálculo de rentabilidade
                            retirada_em_rs = valor_retirada * item['Câmbio']
                            novas_retiradas_acumuladas = item['Retiradas (R$)'] + retirada_em_rs
                            
                            update_existing_asset(
                                row_excel=int(item['Linha_Excel']), 
                                novo_aplicado=novo_aplicado_ret if novo_aplicado_ret > 0 else 0, 
                                novo_saldo=novo_saldo_ret if novo_saldo_ret > 0 else 0, 
                                novos_proventos=item['Proventos (R$)'],
                                retiradas=novas_retiradas_acumuladas
                            )
                            st.success(f"✅ Retirada de {item['Moeda']} {valor_retirada:,.2f} registrada! O valor foi adicionado ao seu histórico de lucros.")
                            st.rerun()
                        else:
                            st.warning("Informe um valor de retirada maior que zero.")
                
                elif "Excluir" in tipo_operacao:
                    st.error("⚠️ **Atenção:** Esta ação excluirá permanentemente o ativo da sua 'Carteira Atual' no Excel. O histórico de lucros deste ativo deixará de contabilizar no painel.")
                    confirmar = st.checkbox("Eu entendo e quero excluir este ativo.")
                    if confirmar:
                        if st.button("❌ Confirmar Exclusão Definitiva", type="primary"):
                            delete_asset(int(item['Linha_Excel']))
                            st.success(f"✅ Ativo **{item['Código']}** foi excluído com sucesso da carteira!")
                            st.rerun()

    # ---------------------------------------------------------
    # OPÇÃO 2: CADASTRAR NOVO ATIVO COM AGRUPAMENTO
    # ---------------------------------------------------------
    else:
        with st.container(border=True):
            st.markdown("#### ➕ Cadastrar Novo Ativo na Carteira")
            st.markdown("Ao cadastrar, o ativo é **automaticamente inserido nas fórmulas e tabelas da mesma Classe e Instituição**, ficando perfeitamente agrupado em todos os gráficos.")
    
            with st.form("form_novo_ativo_grouped"):
                f_col1, f_col2 = st.columns(2)
                
                with f_col1:
                    # Seleciona de classes existentes para garantir agrupamento correto
                    classes_opcoes = ["Renda Fixa", "Renda Variável", "Exterior", "Tesouro Direto", "Cripto", "Caixa", "Moeda"]
                    n_classe = st.selectbox("Classe do Ativo (Para Agrupamento):", options=classes_opcoes)
                    
                    instituicoes_opcoes = ["XP", "Clear", "C6", "Avenue", "Binance", "Wise", "Mercado Pago", "Outro"]
                    n_inst = st.selectbox("Instituição / Banco:", options=instituicoes_opcoes)
                    
                    n_codigo = st.text_input("Código do Ativo (ex: XP-CDB0528, CLR-WEGE3, AVE-NVDA):")
                    n_nome = st.text_input("Nome Completo / Descrição do Ativo:")
    
                with f_col2:
                    n_subclasse = st.text_input("Subclasse (ex: CDB, Ações, FII, ETF, Cripto):")
                    n_moeda = st.selectbox("Moeda:", options=["BRL", "USD", "GBP"])
                    n_aplicado = st.number_input("Valor Aplicado Inicial:", min_value=0.0, step=500.0, format="%.2f")
                    n_saldo = st.number_input("Saldo Atual de Mercado:", min_value=0.0, step=500.0, format="%.2f")
                    n_obs = st.text_input("Observação / Objetivo (ex: 300k, Poupança, Carro):")
    
                submit_btn = st.form_submit_button("🚀 Cadastrar e Agrupar Automaticamente no Excel", type="primary")
                if submit_btn:
                    if not n_codigo or not n_nome:
                        st.error("Por favor, preencha pelo menos o Código e o Nome do ativo.")
                    else:
                        add_new_asset_grouped(
                            n_codigo.strip().upper(),
                            n_nome.strip(),
                            n_inst,
                            n_classe,
                            n_subclasse,
                            n_moeda,
                            n_aplicado,
                            n_saldo,
                            n_obs
                        )
                        st.success(f"🎉 Ativo **{n_codigo}** cadastrado e agrupado na classe **{n_classe}** com sucesso!")
                        st.rerun()

# ------------------------------------------------------------------------------
# ABA 3: REBALANCEAMENTO E METAS (CARTEIRA 300K)
# ------------------------------------------------------------------------------
with tab_metas:
    st.markdown("### 🎯 Rebalanceamento de Carteira (Alvo Estratégico)")
    
    with st.container(border=True):
        col_m1, col_m2 = st.columns([3, 1])
        with col_m1:
            alvo_patrimonio = st.number_input(
                "Definir Meta de Patrimônio Alvo (R$):", 
                min_value=10000.0, 
                value=float(st.session_state.get('meta_patrimonio', 300000.0)), 
                step=10000.0
            )
        with col_m2:
            st.markdown("<div style='padding-top: 28px;'></div>", unsafe_allow_html=True)
            if st.button("💾 Atualizar Meta Global", type="primary", use_container_width=True):
                st.session_state['meta_patrimonio'] = alvo_patrimonio
                # Também salvar na planilha para persistência se quisermos
                wb = openpyxl.load_workbook(EXCEL_FILE)
                ws_ind = wb['Indicadores']
                ws_ind['I2'] = alvo_patrimonio
                wb.save(EXCEL_FILE)
                st.success("Meta atualizada!")
                st.rerun()

    # Metas padrão da estratégia
    metas_dict = {
        'Renda Fixa': 35.0,
        'Exterior': 25.0,
        'Renda Variável': 20.0,
        'Tesouro Direto': 10.0,
        'Cripto': 5.0,
        'Caixa': 3.0,
        'Moeda': 2.0
    }

    df_reb = df.groupby('Classe')['Saldo Atual (R$)'].sum().reset_index()
    total_cart = df['Saldo Atual (R$)'].sum()
    df_reb['% Atual'] = (df_reb['Saldo Atual (R$)'] / total_cart) * 100 if total_cart > 0 else 0.0
    df_reb['% Meta'] = df_reb['Classe'].map(metas_dict).fillna(5.0)
    df_reb['Saldo Meta Ideal (R$)'] = alvo_patrimonio * (df_reb['% Meta'] / 100.0)
    df_reb['Diferença (R$)'] = df_reb['Saldo Meta Ideal (R$)'] - df_reb['Saldo Atual (R$)']
    
    def acao(diff):
        if diff > 1000:
            return f"🟢 Aportar R$ {diff:,.2f}"
        elif diff < -1000:
            return f"🟡 Sobrealocado (Excesso R$ {abs(diff):,.2f})"
        else:
            return "⚪ Equilibrado"

    df_reb['Ação Recomendada'] = df_reb['Diferença (R$)'].apply(acao)

    with st.container(border=True):
        fig_reb = go.Figure(data=[
            go.Bar(name='Saldo Atual', x=df_reb['Classe'], y=df_reb['Saldo Atual (R$)'], marker_color='#3B82F6', text=df_reb['Saldo Atual (R$)']),
            go.Bar(name='Meta Ideal', x=df_reb['Classe'], y=df_reb['Saldo Meta Ideal (R$)'], marker_color='#10B981', text=df_reb['Saldo Meta Ideal (R$)'])
        ])
        fig_reb.update_traces(
            texttemplate='R$ %{y:,.0f}',
            textposition='outside',
            hoverlabel=dict(bgcolor="#0F172A", bordercolor="#3B82F6", font=dict(color="white", size=13))
        )
        fig_reb.update_layout(
            template="plotly_dark",
            paper_bgcolor='rgba(0,0,0,0)',
            plot_bgcolor='rgba(0,0,0,0)',
            barmode='group',
            title="<b>Comparativo: Posição Atual vs. Meta Ideal (R$)</b>",
            height=360,
            margin=dict(t=45, b=20, l=15, r=15),
            yaxis=dict(showgrid=False, showticklabels=False, zeroline=False),
            legend=dict(orientation="h", yanchor="bottom", y=-0.2, xanchor="center", x=0.5)
        )
        st.plotly_chart(fig_reb, use_container_width=True)

    st.dataframe(
        df_reb.style.format({
            'Saldo Atual (R$)': 'R$ {:,.2f}',
            'Saldo Meta Ideal (R$)': 'R$ {:,.2f}',
            'Diferença (R$)': 'R$ {:,.2f}',
            '% Atual': '{:.1f}%',
            '% Meta': '{:.1f}%'
        }),
        use_container_width=True
    )
