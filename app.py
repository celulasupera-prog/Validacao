import io
import re
from io import StringIO
from datetime import datetime

import pandas as pd
import streamlit as st

from processador_eventos import ProcessadorEventosPeriodicos

st.set_page_config(page_title="Consolidador eSocial", page_icon="📊", layout="wide")

st.title("📊 Consolidador de Relatório de Status dos Eventos Periódicos")
st.caption("Faça upload da planilha .xlsx, processe os dados e baixe o consolidado.")

uploaded_file = st.file_uploader("Selecione a planilha de entrada", type=["xlsx"])
arquivo_afastados = st.file_uploader(
    "Lista de afastados (opcional: xlsx/csv/tsv)",
    type=["xlsx", "csv", "tsv", "txt"],
)
texto_afastados = st.text_area(
    "Ou cole os afastados aqui (Ctrl+C / Ctrl+V) - opcional",
    help="Cole linhas no formato: código empresa, nome empresa, código funcionário, nome funcionário.",
    height=140,
)


def carregar_lista_afastados(uploaded, texto):
    frames = []

    if uploaded is not None:
        nome = uploaded.name.lower()
        if nome.endswith(".xlsx"):
            frames.append(pd.read_excel(uploaded))
        elif nome.endswith(".csv"):
            frames.append(pd.read_csv(uploaded))
        else:
            frames.append(pd.read_csv(uploaded, sep="\t"))

    if texto and texto.strip():
        linhas = [l for l in texto.splitlines() if l.strip()]
        if linhas:
            registros = []
            for linha in linhas:
                linha = linha.strip()
                if not linha:
                    continue

                if "\t" in linha:
                    partes = [p.strip() for p in linha.split("\t")]
                elif ";" in linha:
                    partes = [p.strip() for p in linha.split(";")]
                elif "," in linha:
                    partes = [p.strip() for p in linha.split(",")]
                else:
                    # Exemplo aceito: 133 IGREJA ASSEMBLEIA 1 MARIA PASTORINA
                    match = re.match(r"^\s*(\d+)\s+(.+?)\s+(\d+)\s+(.+)\s*$", linha)
                    partes = list(match.groups()) if match else re.split(r"\s{2,}", linha)

                partes = [p for p in partes if str(p).strip()]
                if len(partes) >= 4:
                    registros.append(partes[:4])

            df_texto = pd.DataFrame(registros)

            if df_texto.shape[1] >= 4:
                df_texto = df_texto.iloc[:, :4]
                df_texto.columns = [
                    "codigo_empresa",
                    "empresa",
                    "codigo_funcionario",
                    "nome_funcionario",
                ]
            elif df_texto.shape[1] == 3:
                df_texto.columns = ["empresa", "codigo_funcionario", "nome_funcionario"]
            frames.append(df_texto)

    if not frames:
        return None
    return pd.concat(frames, ignore_index=True)

if uploaded_file:
    st.success(f"Arquivo carregado: {uploaded_file.name}")

    if st.button("Processar planilha", type="primary"):
        with st.spinner("Processando dados..."):
            processador = ProcessadorEventosPeriodicos(uploaded_file)
            processador.processar()

            df_afastados = carregar_lista_afastados(arquivo_afastados, texto_afastados)
            if df_afastados is not None:
                processador.marcar_afastados(df_afastados)

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
