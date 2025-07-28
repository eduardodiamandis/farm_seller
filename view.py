# view.py
import streamlit as st
from openpyxl import Workbook
import pandas as pd
from model import expandir_coluna_e_salvar_v3
import datetime
from io import BytesIO

# Configuração da página
st.set_page_config(page_title="Expansor de Dados Excel", page_icon="📊", layout="wide")
st.title(" Expansor de Dados para Farme_Seller")


# Divisão em abas
tab1, tab2 = st.tabs(["Configuração", "Resultados"])

# Aba de configuração
with tab1:
    # Upload do arquivo
    st.subheader("1. Carregar Arquivo Excel")
    uploaded_file = st.file_uploader("Selecione o arquivo Excel de entrada", type=["xlsx", "xls"])
    
    # Parâmetros básicos
    st.subheader("2. Configurações Básicas")
    col1, col2, col3 = st.columns(3)
    with col1:
        sheet_name = st.text_input("Nome da aba no Excel", "datas")
        repeticoes = st.number_input("Número de repetições", min_value=1, value=17)
        nome_aba_formulas = st.text_input("Nome da aba para fórmulas", "sbs_nc_21")
    with col2:
        col_data = st.number_input("Coluna de dados (número)", min_value=1, value=3)
        col_week = st.number_input("Coluna de semanas (número)", min_value=1, value=2)
        linha_base_valor = st.number_input("Linha base VALOR", min_value=1, value=4)
    with col3:
        intervalo_valor_start = st.text_input("Início intervalo VALOR", "C").upper()
        intervalo_valor_end = st.text_input("Fim intervalo VALOR", "S").upper()
        linha_base_percent = st.number_input("Linha base PERCENT", min_value=1, value=4)
    
    # Estados
    st.subheader("3. Configuração de Estados")
    
    # Inicializar lista de estados na session state
    if 'estados_lista' not in st.session_state:
        st.session_state.estados_lista = [
            "Mato Grosso", "MT N", "MT S", "MT O", "MT L", "Rio Grande", "Paraná",
            "Goiás", "M. T. do Sul", "Santa Catarina", "Minas Gerais", "São Paulo",
            "Bahia", "Tocantins", "Piauí", "Maranhão", "Others"
        ]
    
    col_estados1, col_estados2 = st.columns([3, 1])
    
    with col_estados1:
        # Editor de estados
        estados_text = st.text_area(
            "Lista de Estados (um por linha)", 
            value="\n".join(st.session_state.estados_lista),
            height=200,
            help="Digite um estado por linha. Você também pode copiar/colar uma lista"
        )
        
        # Atualizar lista quando o texto mudar
        if estados_text:
            novos_estados = [e.strip() for e in estados_text.split('\n') if e.strip()]
            st.session_state.estados_lista = novos_estados
    
    with col_estados2:
        st.markdown("**Ações Rápidas**")
        
        # Botão para adicionar estado
        novo_estado = st.text_input("Adicionar novo estado", key="novo_estado_input")
        if st.button("➕ Adicionar", key="add_estado_btn") and novo_estado:
            st.session_state.estados_lista.append(novo_estado.strip())
            st.rerun()
        
        st.divider()
        
        # Botão para resetar estados
        if st.button("🔄 Resetar para padrão", key="reset_estados_btn"):
            st.session_state.estados_lista = [
                "Mato Grosso", "MT N", "MT S", "MT O", "MT L", "Rio Grande", "Paraná",
                "Goiás", "M. T. do Sul", "Santa Catarina", "Minas Gerais", "São Paulo",
                "Bahia", "Tocantins", "Piauí", "Maranhão", "Others"
            ]
            st.rerun()
        
        st.divider()
        
        # Upload de lista de estados
        uploaded_estados = st.file_uploader("Carregar lista de estados", type=["txt", "csv"])
        if uploaded_estados:
            try:
                # Processar arquivo texto (um estado por linha)
                content = uploaded_estados.getvalue().decode("utf-8")
                novos_estados = [e.strip() for e in content.splitlines() if e.strip()]
                if novos_estados:
                    st.session_state.estados_lista = novos_estados
                    st.success(f"{len(novos_estados)} estados carregados!")
                    st.rerun()
            except Exception as e:
                st.error(f"Erro ao processar arquivo: {str(e)}")
    
    # Mostrar lista atual de estados
    st.markdown(f"**Estados configurados ({len(st.session_state.estados_lista)}):**")
    estados_str = ", ".join(st.session_state.estados_lista)
    st.caption(estados_str if len(estados_str) < 100 else estados_str[:100] + "...")
    
    # Intervalo PERCENT
    st.subheader("4. Intervalo PERCENT")
    col4, col5 = st.columns(2)
    with col4:
        intervalo_percent_start = st.text_input("Início intervalo PERCENT", "W").upper()
    with col5:
        intervalo_percent_end = st.text_input("Fim intervalo PERCENT", "AM").upper()
    
    # Botão de execução
    st.subheader("5. Executar Processamento")
    processar = st.button("▶️ Processar Dados", type="primary", use_container_width=True)

# Aba de resultados
with tab2:
    if 'processar' in locals() and processar and uploaded_file:
        with st.spinner("Processando dados... Isso pode levar alguns minutos"):
            try:
                # Criar arquivo temporário
                with open("temp_input.xlsx", "wb") as f:
                    f.write(uploaded_file.getbuffer())
                
                # Chamar a função de processamento
                result, total_rows = expandir_coluna_e_salvar_v3(
                    caminho_entrada="temp_input.xlsx",
                    aba=sheet_name,
                    coluna_index_data=col_data,
                    coluna_index_week_number=col_week,
                    repeticoes=repeticoes,
                    lista_estados=st.session_state.estados_lista,
                    intervalo_valor=(intervalo_valor_start, intervalo_valor_end),
                    intervalo_percent=(intervalo_percent_start, intervalo_percent_end),
                    linha_base_valor=linha_base_valor,
                    linha_base_percent=linha_base_percent,
                    nome_aba=nome_aba_formulas
                )
                
                st.success("✅ Processamento concluído com sucesso!")
                
                # Botão de download
                st.download_button(
                    label="⬇️ Baixar Arquivo Resultante",
                    data=result,
                    file_name="dados_expandidos.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
                # Prévia do arquivo
                st.subheader("Prévia dos Dados Gerados")
                try:
                    # Carregar uma pequena amostra para mostrar
                    wb = load_workbook(filename=BytesIO(result.getvalue()))
                    ws = wb.active
                    
                    # Converter para DataFrame
                    data = []
                    for row in ws.iter_rows(min_row=1, max_row=11, values_only=True):
                        data.append(row)
                    
                    df_preview = pd.DataFrame(data[1:], columns=data[0])
                    st.dataframe(df_preview, hide_index=True)
                    
                    # Mostrar estatísticas
                    col_info1, col_info2 = st.columns(2)
                    with col_info1:
                        st.metric("Total de Linhas Geradas", total_rows)
                    with col_info2:
                        st.metric("Total de Estados", len(st.session_state.estados_lista))
                    
                except Exception as e:
                    st.warning(f"Não foi possível mostrar prévia: {str(e)}")
                    
            except Exception as e:
                st.error(f"❌ Erro no processamento: {str(e)}")
                
    elif 'processar' in locals() and processar:
        st.warning("⏳ Aguardando processamento...")
    else:
        st.info("👆 Configure os parâmetros e clique em 'Processar Dados' para gerar o arquivo")

# Rodapé
st.divider()
st.caption("Ferramenta desenvolvida para expansão de dados regionais - v2.0")
