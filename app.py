import io
import re
from datetime import datetime

import pandas as pd
import streamlit as st

from processador_eventos import ProcessadorEventosPeriodicos

st.set_page_config(page_title="Consolidador eSocial", layout="wide")

st.markdown(
    """
    <style>
    .main {
        background: linear-gradient(180deg, #f8fafc 0%, #eef2ff 100%);
    }
    .hero {
        border-radius: 16px;
        padding: 1.2rem 1.4rem;
        margin-bottom: 1rem;
        background: linear-gradient(135deg, #1d4ed8 0%, #4338ca 55%, #6d28d9 100%);
        color: #ffffff;
        box-shadow: 0 10px 20px rgba(30, 64, 175, 0.20);
    }
    .hero h1 {
        font-size: 1.45rem;
        margin: 0 0 0.35rem 0;
    }
    .hero p {
        margin: 0;
        opacity: 0.95;
    }
    .section-card {
        border: 1px solid #dbeafe;
        border-radius: 14px;
        background: #ffffff;
        color: #0f172a;
        padding: 0.85rem 1rem 0.3rem 1rem;
        margin-bottom: 0.8rem;
        box-shadow: 0 2px 10px rgba(15, 23, 42, 0.04);
    }
    .section-card strong, .section-card em, .section-card span, .section-card p {
        color: #0f172a !important;
    }
    .hero-title {
        display: flex;
        align-items: center;
        gap: 0.5rem;
    }
    .hero-icon {
        width: 22px;
        height: 22px;
        fill: none;
        stroke: #ffffff;
        stroke-width: 2;
        stroke-linecap: round;
        stroke-linejoin: round;
        flex-shrink: 0;
    }
    .label-with-icon {
        display: inline-flex;
        align-items: center;
        gap: 0.4rem;
        font-weight: 600;
        color: #0f172a;
    }
    .label-with-icon svg {
        width: 16px;
        height: 16px;
        stroke: #1d4ed8;
        fill: none;
        stroke-width: 2;
        stroke-linecap: round;
        stroke-linejoin: round;
    }
    div[data-testid="stMetric"] {
        border: 1px solid #e2e8f0;
        background-color: #ffffff;
        border-radius: 12px;
        padding: 0.55rem 0.7rem;
        box-shadow: 0 1px 6px rgba(15, 23, 42, 0.06);
    }
    .stDownloadButton button {
        width: 100%;
        border-radius: 10px;
        font-weight: 600;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

st.markdown(
    """
    <div class="hero">
        <h1 class="hero-title">
            <svg class="hero-icon" viewBox="0 0 24 24" aria-hidden="true">
                <path d="M3 3v18h18"></path>
                <path d="M7 14l3-3 3 2 4-5"></path>
                <circle cx="7" cy="14" r="1"></circle>
                <circle cx="10" cy="11" r="1"></circle>
                <circle cx="13" cy="13" r="1"></circle>
                <circle cx="17" cy="8" r="1"></circle>
            </svg>
            Consolidador de Relatório de Status dos Eventos Periódicos
        </h1>
        <p>Envie a planilha, marque afastados e baixe o consolidado em Excel com poucos cliques.</p>
    </div>
    """,
    unsafe_allow_html=True,
)

col_upload, col_hint = st.columns([2.3, 1.2])
with col_upload:
    st.markdown('<div class="section-card">', unsafe_allow_html=True)
    st.markdown(
        """
        <span class="label-with-icon">
            <svg viewBox="0 0 24 24" aria-hidden="true">
                <path d="M12 3v12"></path>
                <path d="M8 7l4-4 4 4"></path>
                <path d="M4 14v4a3 3 0 0 0 3 3h10a3 3 0 0 0 3-3v-4"></path>
            </svg>
            Selecione a planilha de entrada
        </span>
        """,
        unsafe_allow_html=True,
    )
    uploaded_file = st.file_uploader("Selecione a planilha de entrada", type=["xlsx"], label_visibility="collapsed")
    st.markdown("</div>", unsafe_allow_html=True)
with col_hint:
    st.markdown(
        """
        <div class="section-card">
            <strong>Dica rápida</strong><br/>
            Para afastados, você pode colar dados com <em>TAB</em>, <em>;</em> ou <em>,</em>.
        </div>
        """,
        unsafe_allow_html=True,
    )

st.markdown('<div class="section-card">', unsafe_allow_html=True)
texto_afastados = st.text_area(
    "Cole os afastados aqui (opcional)",
    help="Formato sugerido: código empresa, nome empresa, código funcionário, nome funcionário.",
    height=160,
    placeholder="133\tIGREJA ASSEMBLEIA\t1\tMARIA PASTORINA DE OLIVEIRA",
)
st.markdown("</div>", unsafe_allow_html=True)


def carregar_lista_afastados(texto):
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
            return df_texto
    return None

if uploaded_file:
    st.success(f"Arquivo carregado: {uploaded_file.name}")
    df_preview_afastados = carregar_lista_afastados(texto_afastados)

    st.markdown("#### Prévia dos afastados colados")
    if df_preview_afastados is not None and not df_preview_afastados.empty:
        st.dataframe(df_preview_afastados, use_container_width=True)
    elif texto_afastados and texto_afastados.strip():
        st.warning("Não foi possível interpretar os afastados. Verifique o formato das linhas coladas.")
    else:
        st.info("Cole os dados de afastados para visualizar a prévia antes do processamento.")

    if st.button("Processar planilha", type="primary"):
        with st.spinner("Processando dados..."):
            processador = ProcessadorEventosPeriodicos(uploaded_file)
            processador.processar()

            df_afastados = carregar_lista_afastados(texto_afastados)
            if df_afastados is not None:
                processador.marcar_afastados(df_afastados)

            if processador.dados_consolidados.empty:
                st.warning("Nenhum dado foi identificado para consolidação.")
            else:
                stats = processador.calcular_estatisticas()

                col1, col2, col3, col4 = st.columns(4)
                col1.metric("Total de registros", stats["total_registros"])
                col2.metric("Validados", stats["total_validados"])
                col3.metric("Invalidados", stats["total_invalidados"])
                col4.metric("Afastados", stats["total_afastados"])

                st.dataframe(processador.dados_consolidados, use_container_width=True)

                output = io.BytesIO()
                processador.exportar_excel(output)
                output.seek(0)

                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                nome_saida = f"relacao_eventos_periodicos_consolidado_{timestamp}.xlsx"

                st.download_button(
                    label="Baixar consolidado em Excel",
                    data=output,
                    file_name=nome_saida,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

                with st.expander("Detalhamento por empresa"):
                    st.dataframe(stats["por_empresa"], use_container_width=True)
