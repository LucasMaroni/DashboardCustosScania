import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime, timedelta
import io

# Configuração da página
st.set_page_config(
    page_title="Dashboard Custos Scania",
    page_icon="🚛",
    layout="wide",
    initial_sidebar_state="expanded"
)

# CSS personalizado para estilo corporativo
st.markdown("""
<style>
    .main-header {
        font-size: 2.5rem;
        color: #1f77b4;
        text-align: center;
        margin-bottom: 1rem;
        font-weight: bold;
    }
    .sub-header {
        font-size: 1.2rem;
        color: #666;
        text-align: center;
        margin-bottom: 2rem;
    }
    .metric-card {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        padding: 1.5rem;
        border-radius: 15px;
        color: white;
        text-align: center;
    }
    .stButton button {
        width: 100%;
        border-radius: 10px;
        border: 2px solid #1f77b4;
        background-color: white;
        color: #1f77b4;
        font-weight: bold;
        transition: all 0.3s ease;
    }
    .stButton button:hover {
        background-color: #1f77b4;
        color: white;
    }
    .css-1d391kg {
        padding: 2rem 1rem;
    }
    /* Aumentar o tamanho do popup dos gráficos */
    .hoverlayer {
        font-size: 16px !important;
    }
    .xtooltip {
        font-size: 16px !important;
        padding: 12px !important;
    }
</style>
""", unsafe_allow_html=True)

@st.cache_data
def load_data():
    """
    Função para carregar os dados do arquivo consolidado.xlsx
    """
    try:
        # Carregar o arquivo Excel da página específica
        df = pd.read_excel("consolidado.xlsx", sheet_name="OFICIAL TABLE | DATE")
        
        # Verificar as colunas disponíveis (removido o display lateral)
        print(f"Colunas carregadas: {list(df.columns)}")
        
        # Mapear colunas baseado nas suas imagens - CORRIGINDO OS NOMES EXATOS
        column_mapping = {
            'CODIGO FORNECEDOR': 'COD_FORNECEDOR',
            'CATEGORIA SCANIA': 'CATEGORIA_SCANIA', 
            'NOME FORNECEDOR': 'NOME_FORNECEDOR',
            'DATA EMISSIO': 'DATA_EMISSAO',
            'N. DOCUMENTO': 'DOCUMENTO',
            'TIPO DOC': 'TIPO_DOC',
            'N.PARCELAS': 'PARCELAS',
            'DATA VENCIMENTO': 'DATA_VENCIMENTO',
            'VALOR': 'VALOR',
            'N. TAREFA': 'TAREFA',
            'N.CONTA TAREFA': 'CONTA_TAREFA',
            'NOME TAREFA': 'NOME_TAREFA',
            'TAREFA CONSOLIDADA': 'TAREFA_CONSOLIDADA'
        }
        
        # Renomear colunas se existirem no arquivo
        for old_col, new_col in column_mapping.items():
            if old_col in df.columns:
                df = df.rename(columns={old_col: new_col})
        
        # Tratamento das colunas de data - USANDO DATA VENCIMENTO PARA GRÁFICOS
        date_columns = ['DATA_EMISSAO', 'DATA_VENCIMENTO']
        for col in date_columns:
            if col in df.columns:
                df[col] = pd.to_datetime(df[col], errors='coerce', dayfirst=True)
        
        # Criar colunas derivadas para análise - USANDO DATA VENCIMENTO
        if 'DATA_VENCIMENTO' in df.columns:
            df['ANO'] = df['DATA_VENCIMENTO'].dt.year
            df['MES'] = df['DATA_VENCIMENTO'].dt.month
            df['MES_ANO'] = df['DATA_VENCIMENTO'].dt.strftime('%Y-%m')
            df['MES_NOME'] = df['DATA_VENCIMENTO'].dt.strftime('%b/%Y')
            df['ANO_MES'] = df['DATA_VENCIMENTO'].dt.strftime('%Y-%m')
        
        # Garantir que a coluna VALOR seja numérica
        if 'VALOR' in df.columns:
            # Remover caracteres não numéricos se necessário
            df['VALOR'] = pd.to_numeric(df['VALOR'], errors='coerce')
        
        # Remover linhas com valores nulos essenciais
        df = df.dropna(subset=['VALOR'], how='all')
        
        return df
        
    except Exception as e:
        st.error(f"Erro ao carregar dados: {e}")
        st.info("""
        **Verifique:**
        1. O arquivo 'consolidado.xlsx' está na mesma pasta do script
        2. Existe uma página chamada 'OFICIAL TABLE | DATE'
        3. As colunas estão com os nomes corretos
        """)
        return pd.DataFrame()

def create_header():
    """Cria o cabeçalho com logos lado a lado"""
    col1, col2, col3 = st.columns([2, 3, 1])
    
    with col1:
        # 🔧 LOGOS ATUALIZADAS COM WIDTH 250
        logo_col1, logo_col2 = st.columns(2)
        with logo_col1:
            st.image("maroni.png", width=250)  # Logo Scania
        with logo_col2:
            st.image("scania.png", width=250)  # Logo Maroni
    
    with col2:
        st.markdown("<h1 class='main-header'>RELATÓRIO DE CUSTOS | SCANIA</h1>", unsafe_allow_html=True)
        st.markdown("<div class='sub-header'>Análise de Custos por Concessionária Scania</div>", unsafe_allow_html=True)
    
    with col3:
        st.write(f"**Atualizado em:** {datetime.now().strftime('%d/%m/%Y')}")

def create_filters(df):
    """Cria os filtros interativos na sidebar"""
    st.sidebar.header("🎛️ Filtros Interativos")
    
    # Filtro de datas - USANDO DATA VENCIMENTO
    if 'DATA_VENCIMENTO' in df.columns and not df.empty:
        min_date = df['DATA_VENCIMENTO'].min().date()
        max_date = df['DATA_VENCIMENTO'].max().date()
    else:
        min_date = datetime(2024, 1, 1).date()
        max_date = datetime.now().date()
    
    col1, col2 = st.sidebar.columns(2)
    with col1:
        data_inicio = st.date_input(
            "Data Vencimento Início",
            value=min_date,
            min_value=min_date,
            max_value=max_date
        )
    
    with col2:
        data_fim = st.date_input(
            "Data Vencimento Fim", 
            value=max_date,
            min_value=min_date,
            max_value=max_date
        )
    
    # Filtro de ano - USANDO DATA VENCIMENTO
    if 'ANO' in df.columns and not df.empty:
        anos = sorted(df['ANO'].unique(), reverse=True)
        ano_selecionado = st.sidebar.selectbox(
            "📅 Selecione o Ano",
            options=["Todos"] + list(anos),
            index=0
        )
    else:
        ano_selecionado = "Todos"
    
    # NOVOS FILTROS ADICIONAIS
    st.sidebar.markdown("---")
    st.sidebar.markdown("### 🔧 Filtros Adicionais")
    
    # Filtro de Tipo de Documento
    if 'TIPO_DOC' in df.columns and not df.empty:
        tipos_doc = st.sidebar.multiselect(
            "📄 Tipo de Documento",
            options=sorted(df['TIPO_DOC'].unique()),
            default=df['TIPO_DOC'].unique()
        )
    else:
        tipos_doc = []
    
    # Filtro de Tarefa Consolidada
    if 'TAREFA_CONSOLIDADA' in df.columns and not df.empty:
        tarefas_consolidadas = st.sidebar.multiselect(
            "🏷️ Tarefa Consolidada",
            options=sorted(df['TAREFA_CONSOLIDADA'].unique()),
            default=df['TAREFA_CONSOLIDADA'].unique()
        )
    else:
        tarefas_consolidadas = []
    
    # Filtro de Nome da Tarefa (mantido para compatibilidade)
    if 'NOME_TAREFA' in df.columns and not df.empty:
        nomes_tarefa = st.sidebar.multiselect(
            "📋 Nome da Tarefa",
            options=sorted(df['NOME_TAREFA'].unique()),
            default=df['NOME_TAREFA'].unique()
        )
    else:
        nomes_tarefa = []
    
    # Botões interativos para CATEGORIA SCANIA
    st.sidebar.markdown("---")
    st.sidebar.markdown("### 🏢 Categorias Scania")
    
    if 'CATEGORIA_SCANIA' in df.columns and not df.empty:
        categorias_scania = sorted(df['CATEGORIA_SCANIA'].unique())
        
        # Botão "Todos"
        if st.sidebar.button("✅ Todas as Categorias", key="all_categories"):
            st.session_state.selected_categories = categorias_scania
        
        # Multiselect com todas as categorias
        if 'selected_categories' not in st.session_state:
            st.session_state.selected_categories = categorias_scania
            
        categorias_selecionadas = st.sidebar.multiselect(
            "Selecione as categorias:",
            options=categorias_scania,
            default=st.session_state.selected_categories,
            key="selected_categories_multiselect"
        )
    else:
        categorias_selecionadas = []
    
    # Filtro de categorias (CUSTO MANUTENÇÕES, SINISTRO, etc)
    if 'CATEGORIA' in df.columns and not df.empty:
        categorias = st.sidebar.multiselect(
            "🏷️ Tipo de Categoria",
            options=df['CATEGORIA'].unique(),
            default=df['CATEGORIA'].unique()
        )
    else:
        categorias = []
    
    return data_inicio, data_fim, ano_selecionado, categorias_selecionadas, categorias, tipos_doc, nomes_tarefa, tarefas_consolidadas

def apply_filters(df, data_inicio, data_fim, ano_selecionado, categorias_selecionadas, categorias, tipos_doc, nomes_tarefa, tarefas_consolidadas):
    """Aplica os filtros ao dataframe"""
    df_filtrado = df.copy()
    
    # Filtro de data - USANDO DATA VENCIMENTO
    if 'DATA_VENCIMENTO' in df_filtrado.columns and not df_filtrado.empty:
        df_filtrado = df_filtrado[
            (df_filtrado['DATA_VENCIMENTO'].dt.date >= data_inicio) & 
            (df_filtrado['DATA_VENCIMENTO'].dt.date <= data_fim)
        ]
    
    # Filtro de ano
    if ano_selecionado != "Todos" and 'ANO' in df_filtrado.columns:
        df_filtrado = df_filtrado[df_filtrado['ANO'] == ano_selecionado]
    
    # Filtro de categorias Scania
    if categorias_selecionadas and 'CATEGORIA_SCANIA' in df_filtrado.columns:
        df_filtrado = df_filtrado[df_filtrado['CATEGORIA_SCANIA'].isin(categorias_selecionadas)]
    
    # Filtro de categorias
    if categorias and 'CATEGORIA' in df_filtrado.columns:
        df_filtrado = df_filtrado[df_filtrado['CATEGORIA'].isin(categorias)]
    
    # NOVOS FILTROS APLICADOS
    # Filtro de Tipo de Documento
    if tipos_doc and 'TIPO_DOC' in df_filtrado.columns:
        df_filtrado = df_filtrado[df_filtrado['TIPO_DOC'].isin(tipos_doc)]
    
    # Filtro de Nome da Tarefa
    if nomes_tarefa and 'NOME_TAREFA' in df_filtrado.columns:
        df_filtrado = df_filtrado[df_filtrado['NOME_TAREFA'].isin(nomes_tarefa)]
    
    # Filtro de Tarefa Consolidada
    if tarefas_consolidadas and 'TAREFA_CONSOLIDADA' in df_filtrado.columns:
        df_filtrado = df_filtrado[df_filtrado['TAREFA_CONSOLIDADA'].isin(tarefas_consolidadas)]
    
    return df_filtrado

def create_metrics(df_filtrado):
    """Cria os cartões de métricas"""
    st.markdown("## 📊 Métricas Principais")
    
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        if not df_filtrado.empty and 'VALOR' in df_filtrado.columns:
            custo_total = df_filtrado['VALOR'].sum()
            qtd_registros = len(df_filtrado)
            st.metric(
                label="💰 Custo Total", 
                value=f"R$ {custo_total:,.2f}",
                delta=f"{qtd_registros} documentos"
            )
        else:
            st.metric("💰 Custo Total", "R$ 0,00")
    
    with col2:
        if not df_filtrado.empty:
            qtd_registros = len(df_filtrado)
            custo_total = df_filtrado['VALOR'].sum() if 'VALOR' in df_filtrado.columns else 0
            st.metric(
                label="📋 Qtd. Documentos", 
                value=qtd_registros,
                delta=f"R$ {custo_total/qtd_registros:,.2f} médio" if qtd_registros > 0 else "0"
            )
        else:
            st.metric("📋 Qtd. Documentos", "0")
    
    with col3:
        if not df_filtrado.empty and 'VALOR' in df_filtrado.columns:
            custo_total = df_filtrado['VALOR'].sum()
            qtd_registros = len(df_filtrado)
            custo_medio = custo_total / qtd_registros if qtd_registros > 0 else 0
            st.metric(
                label="⚖️ Custo Médio", 
                value=f"R$ {custo_medio:,.2f}",
                delta="Por documento"
            )
        else:
            st.metric("⚖️ Custo Médio", "R$ 0,00")
    
    with col4:
        if not df_filtrado.empty and 'VALOR' in df_filtrado.columns:
            maior_despesa = df_filtrado['VALOR'].max()
            st.metric(
                label="🔝 Maior Despesa", 
                value=f"R$ {maior_despesa:,.2f}",
                delta="Valor máximo"
            )
        else:
            st.metric("🔝 Maior Despesa", "R$ 0,00")

def create_charts(df_filtrado):
    """Cria os gráficos interativos"""
    
    if df_filtrado.empty:
        st.warning("⚠️ Nenhum dado encontrado com os filtros aplicados.")
        return
    
    # GRÁFICO 1: Custo por Mês (Colunas) - USANDO DATA VENCIMENTO
    st.markdown("### 📅 Custo por Mês (Data Vencimento)")
    
    if 'MES_ANO' in df_filtrado.columns and 'VALOR' in df_filtrado.columns:
        df_mensal = df_filtrado.groupby(['MES_ANO', 'MES_NOME', 'ANO', 'MES'])['VALOR'].sum().reset_index()
        df_mensal = df_mensal.sort_values(['ANO', 'MES'])
        
        if not df_mensal.empty:
            fig_mensal = px.bar(
                df_mensal,
                x='MES_NOME',
                y='VALOR',
                color='ANO',
                text=df_mensal['VALOR'].apply(lambda x: f"R$ {x:,.0f}"),
                color_discrete_sequence=px.colors.qualitative.Bold,
                title="Evolução Mensal de Custos por Data de Vencimento"
            )
            
            fig_mensal.update_traces(
                textposition='outside',
                textfont=dict(size=12, color='black'),
                marker=dict(line=dict(width=2, color='DarkSlateGrey')),
                hovertemplate='<b>%{x}</b><br>Valor Total: R$ %{y:,.2f}<br>Ano: %{customdata[0]}<extra></extra>',
                customdata=df_mensal[['ANO']]
            )
            
            # AUMENTAR O POPUP (hover) e melhorar layout
            fig_mensal.update_layout(
                xaxis_title="Mês/Ano",
                yaxis_title="Valor (R$)",
                height=500,
                showlegend=True,
                hovermode='x unified',
                hoverlabel=dict(
                    bgcolor="white",
                    font_size=16,
                    font_family="Arial"
                ),
                plot_bgcolor='rgba(0,0,0,0)',
                paper_bgcolor='rgba(0,0,0,0)',
                font=dict(size=14)
            )
            
            # Melhorar eixos
            fig_mensal.update_xaxes(showgrid=True, gridwidth=1, gridcolor='LightGray')
            fig_mensal.update_yaxes(showgrid=True, gridwidth=1, gridcolor='LightGray')
            
            st.plotly_chart(fig_mensal, use_container_width=True)
        else:
            st.info("📊 Não há dados para exibir o gráfico mensal.")
    else:
        st.info("📊 Colunas necessárias não encontradas para o gráfico mensal.")
    
    # GRÁFICO 2 e 3: Lado a lado
    col1, col2 = st.columns(2)
    
    with col1:
        # Gráfico de Pizza por Categoria - CORES EM AZUL
        st.markdown("### 🎯 Distribuição por Categoria")
        
        if 'CATEGORIA' in df_filtrado.columns and 'VALOR' in df_filtrado.columns:
            df_categoria = df_filtrado.groupby('CATEGORIA')['VALOR'].sum().reset_index()
            
            if not df_categoria.empty:
                # Paleta de cores em tons de azul
                cores_azul = ['#1f77b4', '#4a90e2', '#7eb0ff', '#a6c8ff', '#cfe2ff', '#1e3a8a', '#2563eb', '#3b82f6']
                
                fig_pizza = px.pie(
                    df_categoria,
                    values='VALOR',
                    names='CATEGORIA',
                    hole=0.4,
                    color_discrete_sequence=cores_azul,
                    title="Distribuição por Categoria"
                )
                
                fig_pizza.update_traces(
                    textinfo='percent+label',
                    hovertemplate='<b>%{label}</b><br>Valor Total: R$ %{value:,.2f}<br>Percentual: %{percent}<extra></extra>',
                    textfont=dict(size=14),
                    marker=dict(line=dict(color='white', width=2))
                )
                
                fig_pizza.update_layout(
                    height=500,
                    showlegend=True,
                    hoverlabel=dict(
                        bgcolor="white",
                        font_size=16,
                        font_family="Arial"
                    ),
                    legend=dict(
                        orientation="v",
                        yanchor="top",
                        y=1,
                        xanchor="left",
                        x=1.05
                    )
                )
                
                st.plotly_chart(fig_pizza, use_container_width=True)
            else:
                st.info("📊 Não há dados para exibir o gráfico de pizza.")
        else:
            st.info("📊 Colunas necessárias não encontradas para o gráfico de pizza.")
    
    with col2:
        # Gráfico de Barras por Concessionária
        st.markdown("### 🏢 Top Concessionárias")
        
        if 'CATEGORIA_SCANIA' in df_filtrado.columns and 'VALOR' in df_filtrado.columns:
            df_concessionaria = df_filtrado.groupby('CATEGORIA_SCANIA')['VALOR'].sum().reset_index()
            df_concessionaria = df_concessionaria.sort_values('VALOR', ascending=True)
            
            if not df_concessionaria.empty:
                fig_concessionaria = px.bar(
                    df_concessionaria,
                    x='VALOR',
                    y='CATEGORIA_SCANIA',
                    orientation='h',
                    color='VALOR',
                    text=df_concessionaria['VALOR'].apply(lambda x: f"R$ {x:,.2f}"),
                    color_continuous_scale='Blues',
                    title="Custo por Concessionária"
                )
                
                fig_concessionaria.update_layout(
                    yaxis_title="",
                    xaxis_title="Valor (R$)",
                    height=500,
                    showlegend=False,
                    hoverlabel=dict(
                        bgcolor="white",
                        font_size=16,
                        font_family="Arial"
                    ),
                    plot_bgcolor='rgba(0,0,0,0)',
                    paper_bgcolor='rgba(0,0,0,0)'
                )
                
                fig_concessionaria.update_traces(
                    textposition='outside',
                    textfont=dict(size=11),
                    hovertemplate='<b>%{y}</b><br>Valor Total: R$ %{x:,.2f}<extra></extra>'
                )
                
                # Melhorar eixos
                fig_concessionaria.update_xaxes(showgrid=True, gridwidth=1, gridcolor='LightGray')
                fig_concessionaria.update_yaxes(showgrid=True, gridwidth=1, gridcolor='LightGray')
                
                st.plotly_chart(fig_concessionaria, use_container_width=True)
            else:
                st.info("📊 Não há dados para exibir o gráfico de concessionárias.")
        else:
            st.info("📊 Colunas necessárias não encontradas para o gráfico de concessionárias.")

def create_tarefa_consolidada_chart(df_filtrado):
    """Cria o gráfico de custos por TAREFA CONSOLIDADA"""
    st.markdown("### 📊 Custo por Tarefa Consolidada")
    
    if 'TAREFA_CONSOLIDADA' in df_filtrado.columns and 'VALOR' in df_filtrado.columns and not df_filtrado.empty:
        df_tarefa = df_filtrado.groupby('TAREFA_CONSOLIDADA')['VALOR'].sum().reset_index()
        df_tarefa = df_tarefa.sort_values('VALOR', ascending=False).head(15)  # Top 15 tarefas consolidadas
        
        if not df_tarefa.empty:
            # Paleta de cores em tons de azul para o gráfico de tarefas consolidadas
            cores_azul = ['#1f77b4', '#4a90e2', '#7eb0ff', '#a6c8ff', '#cfe2ff', '#1e3a8a', '#2563eb', '#3b82f6']
            
            fig_tarefa = px.bar(
                df_tarefa,
                x='VALOR',
                y='TAREFA_CONSOLIDADA',
                orientation='h',
                color='VALOR',
                text=df_tarefa['VALOR'].apply(lambda x: f"R$ {x:,.2f}"),
                color_continuous_scale='Blues',  # Escala de azul
                title="Top 15 Custos por Tarefa Consolidada"
            )
            
            fig_tarefa.update_layout(
                yaxis_title="Tarefa Consolidada",
                xaxis_title="Valor (R$)",
                height=600,
                showlegend=False,
                hoverlabel=dict(
                    bgcolor="white",
                    font_size=14,
                    font_family="Arial"
                ),
                plot_bgcolor='rgba(0,0,0,0)',
                paper_bgcolor='rgba(0,0,0,0)'
            )
            
            fig_tarefa.update_traces(
                textposition='outside',
                textfont=dict(size=10),
                hovertemplate='<b>%{y}</b><br>Valor Total: R$ %{x:,.2f}<extra></extra>',
                marker=dict(line=dict(color='darkblue', width=1))
            )
            
            # Melhorar eixos
            fig_tarefa.update_xaxes(showgrid=True, gridwidth=1, gridcolor='LightGray')
            fig_tarefa.update_yaxes(showgrid=True, gridwidth=1, gridcolor='LightGray')
            
            st.plotly_chart(fig_tarefa, use_container_width=True)
        else:
            st.info("📊 Não há dados para exibir o gráfico de tarefas consolidadas.")
    else:
        st.info("📊 Colunas necessárias não encontradas para o gráfico de tarefas consolidadas.")

def create_tables(df_filtrado):
    """Cria as tabelas de dados"""
    
    if df_filtrado.empty:
        st.warning("⚠️ Nenhum dado para exibir nas tabelas.")
        return
    
    # NOVO GRÁFICO: Custo por Tarefa Consolidada (substituindo o de Nome Tarefa)
    create_tarefa_consolidada_chart(df_filtrado)
    
    st.markdown("---")
    
    # Tabela 1: Fornecedores e Valores
    st.markdown("### 📋 Ranking de Fornecedores")
    
    if 'NOME_FORNECEDOR' in df_filtrado.columns and 'VALOR' in df_filtrado.columns:
        df_fornecedores = df_filtrado.groupby('NOME_FORNECEDOR').agg({
            'VALOR': 'sum',
            'DOCUMENTO': 'count'
        }).reset_index()
        
        df_fornecedores = df_fornecedores.rename(columns={
            'DOCUMENTO': 'QTD_DOCUMENTOS'
        }).sort_values('VALOR', ascending=False)
        
        df_fornecedores['VALOR_FORMATADO'] = df_fornecedores['VALOR'].apply(lambda x: f"R$ {x:,.2f}")
        
        # Mostrar apenas as colunas NOME_FORNECEDOR e VALOR
        st.dataframe(
            df_fornecedores[['NOME_FORNECEDOR', 'VALOR_FORMATADO', 'QTD_DOCUMENTOS']].reset_index(drop=True),
            use_container_width=True,
            height=400
        )
    else:
        st.info("📋 Colunas necessárias não encontradas para a tabela de fornecedores.")
    
    # Tabela 2: TODOS OS DADOS DISPONÍVEIS
    st.markdown("### 📊 Tabela Completa de Dados")
    
    # Mostrar todas as colunas disponíveis
    if not df_filtrado.empty:
        # Formatar colunas para exibição
        df_display = df_filtrado.copy()
        
        # Formatar colunas de data
        date_columns = ['DATA_EMISSAO', 'DATA_VENCIMENTO']
        for col in date_columns:
            if col in df_display.columns:
                df_display[col] = df_display[col].dt.strftime('%d/%m/%Y')
        
        # Formatar coluna de valor
        if 'VALOR' in df_display.columns:
            df_display['VALOR'] = df_display['VALOR'].apply(lambda x: f"R$ {x:,.2f}")
        
        # Ordenar por data de vencimento (mais recente primeiro)
        if 'DATA_VENCIMENTO' in df_filtrado.columns:
            df_display = df_display.sort_values('DATA_VENCIMENTO', ascending=False)
        
        st.dataframe(
            df_display,
            use_container_width=True,
            height=600
        )
        
        # Botão de download em Excel
        st.markdown("---")
        col1, col2, col3 = st.columns([1, 2, 1])
        with col2:
            # Função para converter para Excel
            def convert_to_excel(df):
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df.to_excel(writer, index=False, sheet_name='Dados_Completos')
                return output.getvalue()
            
            excel_data = convert_to_excel(df_filtrado)
            
            st.download_button(
                label="📥 Baixar Tabela Completa (Excel)",
                data=excel_data,
                file_name=f"dados_completos_custos_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
    else:
        st.info("📋 Nenhuma coluna disponível para a tabela completa.")

def main():
    # Inicializar session state para categorias
    if 'selected_categories' not in st.session_state:
        st.session_state.selected_categories = []
    
    # Carregar dados
    df = load_data()
    
    if df.empty:
        st.error("""
        ❌ Não foi possível carregar os dados. Verifique:
        1. O arquivo 'consolidado.xlsx' está na mesma pasta
        2. A planilha 'OFICIAL TABLE | DATE' existe
        3. O arquivo não está corrompido
        """)
        return
    
    # Mostrar informações sobre os dados carregados (sem mostrar na sidebar)
    st.sidebar.success(f"✅ Dados carregados com sucesso!")
    
    # Criar cabeçalho
    create_header()
    
    # Criar filtros
    data_inicio, data_fim, ano_selecionado, categorias_selecionadas, categorias, tipos_doc, nomes_tarefa, tarefas_consolidadas = create_filters(df)
    
    # Aplicar filtros
    df_filtrado = apply_filters(df, data_inicio, data_fim, ano_selecionado, categorias_selecionadas, categorias, tipos_doc, nomes_tarefa, tarefas_consolidadas)
    
    # Mostrar métricas
    create_metrics(df_filtrado)
    
    st.markdown("---")
    
    # Mostrar gráficos
    create_charts(df_filtrado)
    
    st.markdown("---")
    
    # Mostrar tabelas
    create_tables(df_filtrado)

if __name__ == "__main__":

    main()
