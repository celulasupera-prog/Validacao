import io
from datetime import datetime

import streamlit as st

from processador_eventos import ProcessadorEventosPeriodicos

st.set_page_config(page_title="Consolidador eSocial", page_icon="📊", layout="wide")

st.title("📊 Consolidador de Relatório de Status dos Eventos Periódicos")
st.caption("Faça upload da planilha .xlsx, processe os dados e baixe o consolidado.")

uploaded_file = st.file_uploader("Selecione a planilha de entrada", type=["xlsx"])

if uploaded_file:
    st.success(f"Arquivo carregado: {uploaded_file.name}")

    if st.button("Processar planilha", type="primary"):
        with st.spinner("Processando dados..."):
            processador = ProcessadorEventosPeriodicos(uploaded_file)
            processador.processar()

            if processador.dados_consolidados.empty:
                st.warning("Nenhum dado foi identificado para consolidação.")
            else:
                stats = processador.calcular_estatisticas()

                col1, col2, col3 = st.columns(3)
                col1.metric("Total de registros", stats["total_registros"])
                col2.metric("Validados", stats["total_validados"])
                col3.metric("Invalidados", stats["total_invalidados"])

                st.dataframe(processador.dados_consolidados, use_container_width=True)

                output = io.BytesIO()
                processador.exportar_excel(output)
                output.seek(0)

                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                nome_saida = f"relacao_eventos_periodicos_consolidado_{timestamp}.xlsx"

                st.download_button(
                    label="⬇️ Baixar consolidado em Excel",
                    data=output,
                    file_name=nome_saida,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

                with st.expander("Detalhamento por empresa"):
                    st.dataframe(stats["por_empresa"], use_container_width=True)
