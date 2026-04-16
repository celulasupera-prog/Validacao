import io
import os
import re
from datetime import datetime

import pandas as pd
import streamlit as st

from processador_eventos import ProcessadorEventosPeriodicos
from supabase_client import SupabaseError, SupabaseRestClient

st.set_page_config(page_title="Consolidador eSocial", layout="wide")

st.markdown(
    """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=DM+Sans:ital,wght@0,300;0,400;0,500;0,700;1,400&family=Syne:wght@600;700;800&display=swap');
    :root {
        --bg: #080808;
        --panel: #111111;
        --panel-2: #181818;
        --border: rgba(255, 255, 255, 0.07);
        --border-hover: rgba(255, 255, 255, 0.18);
        --text: #f0f0f0;
        --muted: #888888;
        --cyan: #00e5ff;
        --purple: #c77dff;
        --lime: #c8ff5d;
        --primary: #00e5ff;
        --primary-hover: #29ebff;
        --shadow: 0 20px 50px rgba(0,0,0,0.55);
    }
    html, body, [class*="css"]  {
        font-family: "DM Sans", sans-serif;
        scroll-behavior: smooth;
    }
    .main {
        position: relative;
        overflow: hidden;
        font-family: "DM Sans", sans-serif;
        background:
            radial-gradient(circle at 16% 6%, rgba(0,229,255,0.1), transparent 26%),
            radial-gradient(circle at 84% 85%, rgba(199,125,255,0.1), transparent 28%),
            linear-gradient(180deg, #0a0a0a 0%, #080808 100%);
        color: var(--text);
    }
    .main::before {
        content: "";
        position: fixed;
        inset: 0;
        pointer-events: none;
        opacity: 0.04;
        z-index: 0;
        background-image:
            radial-gradient(circle at 20% 20%, rgba(255,255,255,0.8) 1px, transparent 1px),
            radial-gradient(circle at 80% 75%, rgba(255,255,255,0.65) 1px, transparent 1px);
        background-size: 3px 3px, 4px 4px;
    }
    .bg-blob {
        position: fixed;
        filter: blur(120px);
        opacity: 0.12;
        z-index: 0;
        pointer-events: none;
        border-radius: 999px;
    }
    .blob-cyan { width: 600px; height: 600px; top: -230px; left: -180px; background: var(--cyan); }
    .blob-pink { width: 500px; height: 500px; right: -160px; bottom: -220px; background: var(--purple); }
    .blob-lime { width: 400px; height: 400px; left: 50%; top: 38%; transform: translate(-50%, -50%); background: var(--lime); }
    div[data-testid="stVerticalBlock"]:has(.hero-shell-anchor) {
        border: 1px solid var(--border);
        border-radius: 24px;
        padding: 1.2rem 1.35rem 1.35rem;
        margin-bottom: 1rem;
        background: linear-gradient(135deg, rgba(17,17,17,0.95), rgba(24,24,24,0.94));
        color: var(--text);
        backdrop-filter: blur(14px);
        box-shadow: var(--shadow);
        position: relative;
        z-index: 1;
    }
    .hero-head h1 {
        font-family: "Syne", sans-serif;
        font-size: clamp(2.75rem, 7.3vw, 5.2rem);
        letter-spacing: -0.03em;
        line-height: 0.98;
        margin: 0 0 0.35rem 0;
        font-weight: 10;
    }
    .hero-head p {
        margin: 0 0 0.75rem;
        color: var(--muted);
        max-width: 56ch;
        font-size: 1.05rem;
    }
    .hero-divider {
        margin: 1rem 0 0.5rem 0;
        border: 0;
        border-top: 1px solid var(--border);
    }
    .hero-tip {
        border: 1px solid var(--border);
        border-radius: 20px;
        background: linear-gradient(180deg, #121212, #181818);
        padding: 0.9rem;
        height: 100%;
    }
    .section-card {
        border: 1px solid var(--border);
        border-radius: 20px;
        background: linear-gradient(180deg, #121212, #171717);
        color: var(--text);
        padding: 0.8rem 0.95rem 0.3rem 0.95rem;
        margin-bottom: 0.75rem;
        box-shadow: var(--shadow);
        transition: transform .3s ease, border-color .3s ease;
    }
    .section-card:hover {
        transform: translateY(-4px);
        border-color: var(--border-hover);
    }
    .section-card strong, .section-card em, .section-card span, .section-card p {
        color: var(--text) !important;
    }
    .hero-title {
        display: flex;
        align-items: flex-start;
        flex-direction: column;
        gap: 0.3rem;
    }
    .hero-icon {
        width: 26px;
        height: 26px;
        fill: none;
        stroke: var(--cyan);
        stroke-width: 2;
        stroke-linecap: round;
        stroke-linejoin: round;
        flex-shrink: 0;
        filter: drop-shadow(0 0 14px rgba(0,229,255,0.45));
    }
    .hero-highlight {
        background: linear-gradient(90deg, var(--cyan) 0%, #ff6ad5 100%);
        -webkit-background-clip: text;
        background-clip: text;
        color: transparent;
    }
    .hero-badge {
        display:inline-flex;
        align-items:center;
        gap:8px;
        padding:7px 12px;
        border-radius:999px;
        border:1px solid rgba(200,255,93,0.5);
        background:rgba(200,255,93,0.08);
        color:#ddff92;
        font-size:12px;
        font-family:"Syne",sans-serif;
        font-weight:600;
        letter-spacing:0.12em;
        text-transform:uppercase;
        margin-bottom:14px;
    }
    .fade-up {
        opacity: 0;
        transform: translateY(16px);
        animation: fadeUp .75s ease forwards;
    }
    .delay-1 { animation-delay: .08s; }
    .delay-2 { animation-delay: .18s; }
    .delay-3 { animation-delay: .28s; }
    .delay-4 { animation-delay: .38s; }
    @keyframes fadeUp {
        to { opacity: 1; transform: translateY(0); }
    }
    .feature-grid {
        display:grid;
        grid-template-columns:repeat(3,minmax(160px,1fr));
        gap:12px;
        min-width:300px;
    }
    .feature-card {
        border:1px solid var(--border);
        border-radius:20px;
        background:#121212;
        transition:transform .3s ease,border-color .3s ease;
        overflow:hidden;
    }
    .feature-card:hover {
        transform:translateY(-4px);
        border-color:var(--border-hover);
    }
    .feature-preview {
        height:88px;
        display:flex;
        align-items:center;
        justify-content:center;
        font-size:2rem;
        position:relative;
        background: radial-gradient(circle at 50% 45%, rgba(255,255,255,0.12) 0%, transparent 58%);
    }
    .feature-preview::before{
        content:"";
        position:absolute;
        width:74px;
        height:74px;
        border-radius:999px;
        filter:blur(22px);
        opacity:.38;
    }
    .preview-cyan::before{ background:var(--cyan);}
    .preview-purple::before{ background:var(--purple);}
    .preview-lime::before{ background:var(--lime);}
    .feature-body{ padding:10px 12px 12px; }
    .feature-body span { display:block;color:var(--muted);font-size:11px;letter-spacing:.08em;text-transform:uppercase; }
    .feature-body strong { font-family:"Syne",sans-serif;font-size:1rem;color:var(--text); }
    .label-with-icon {
        font-family: "Syne", sans-serif;
        letter-spacing: .02em;
    }
    .label-with-icon {
        display: inline-flex;
        align-items: center;
        gap: 0.4rem;
        font-weight: 600;
        color: var(--text);
    }
    .label-with-icon svg {
        width: 16px;
        height: 16px;
        stroke: #60a5fa;
        fill: none;
        stroke-width: 2;
        stroke-linecap: round;
        stroke-linejoin: round;
    }
    .stFileUploader label, .stTextArea label {
        color: var(--text) !important;
        font-weight: 600;
    }
    .stFileUploader small, .stTextArea small {
        color: var(--muted) !important;
    }
    div[data-testid="stMetric"] {
        border: 1px solid var(--border);
        background-color: rgba(255,255,255,0.03);
        border-radius: 12px;
        padding: 0.55rem 0.7rem;
        box-shadow: 0 8px 20px rgba(0,0,0,0.25);
    }
    div[data-testid="stMetricLabel"] p {
        font-weight: 600;
        color: var(--muted);
    }
    div[data-testid="stMetricValue"] {
        color: var(--text);
    }
    .stDataFrame {
        border: 1px solid var(--border);
        border-radius: 14px;
        overflow: hidden;
    }
    .stDownloadButton button {
        width: 100%;
        border-radius: 14px;
        font-weight: 700;
        background: var(--primary);
        color: #001318;
        border: 1px solid var(--primary);
    }
    .stDownloadButton button:hover {
        background: var(--primary-hover);
        border-color: var(--primary-hover);
    }
    .stButton button {
        border-radius: 14px;
        font-weight: 700;
        border: 1px solid var(--border);
        transition: all .3s ease;
    }
    .stButton button:hover {
        border-color: var(--border-hover);
        transform: translateY(-2px);
    }
    @media (max-width: 600px) {
        .feature-grid { grid-template-columns:1fr; }
        .hero-head p { font-size: .95rem; }
    }
    .stSuccess, .stWarning, .stInfo {
        border-radius: 12px;
    }
    </style>
    """,
    unsafe_allow_html=True,
)
st.markdown(
    """
    <div class="bg-blob blob-cyan"></div>
    <div class="bg-blob blob-pink"></div>
    <div class="bg-blob blob-lime"></div>
    <div class="hero-shell-anchor"></div>
    <div class="hero-head">
        <div style="display:flex;justify-content:space-between;gap:16px;align-items:flex-start;flex-wrap:wrap;">
        <div style="max-width:760px;">
            <div class="hero-badge fade-up">Consolidação inteligente</div>
        <h1 class="hero-title fade-up delay-1">
            <svg class="hero-icon" viewBox="0 0 24 24" aria-hidden="true">
                <path d="M3 3v18h18"></path>
                <path d="M7 14l3-3 3 2 4-5"></path>
                <circle cx="7" cy="14" r="1"></circle>
                <circle cx="10" cy="11" r="1"></circle>
                <circle cx="13" cy="13" r="1"></circle>
                <circle cx="17" cy="8" r="1"></circle>
            </svg>
            <span>Consolidador de Relatório de Status</span>
            <span class="hero-highlight">dos Eventos Periódicos</span>
        </h1>
        <p class="fade-up delay-2">Envie a planilha, marque afastados e baixe o consolidado em Excel com poucos cliques.</p>
        </div>
        <div class="feature-grid">
            <div class="feature-card fade-up delay-2">
                <div class="feature-preview preview-cyan">📄</div>
                <div class="feature-body"><span>Formato ↗</span><strong>.XLSX</strong></div>
            </div>
            <div class="feature-card fade-up delay-3">
                <div class="feature-preview preview-purple">⚡</div>
                <div class="feature-body"><span>Tamanho ↗</span><strong>200MB</strong></div>
            </div>
            <div class="feature-card fade-up delay-4">
                <div class="feature-preview preview-lime">🟢</div>
                <div class="feature-body"><span>Status ↗</span><strong>Pronto</strong></div>
            </div>
        </div>
        </div>
    </div>
    <hr class="hero-divider"/>
    """,
    unsafe_allow_html=True,
)

uploaded_file = None
col_upload, col_hint = st.columns([2.3, 1.2])
with col_upload:
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
    uploaded_file = st.file_uploader(
        "Selecione a planilha de entrada",
        type=["xlsx"],
        label_visibility="collapsed",
    )
    st.markdown("</div>", unsafe_allow_html=True)
with col_hint:
    st.markdown(
        """
        <div class="hero-tip">
            <strong>Dica rápida</strong><br/>
            Para afastados, você pode colar dados com <em>TAB</em>, <em>;</em> ou <em>,</em>.
        </div>
        <hr class="hero-divider"/>
        """,
        unsafe_allow_html=True,
        )

supabase_url = st.secrets.get("SUPABASE_URL") or os.getenv("SUPABASE_URL")
supabase_key = st.secrets.get("SUPABASE_SERVICE_ROLE_KEY") or os.getenv(
    "SUPABASE_SERVICE_ROLE_KEY"
)
supabase_client = None
if supabase_url and supabase_key:
    try:
        supabase_client = SupabaseRestClient(supabase_url, supabase_key)
        supabase_client.ensure_default_groups(["Supera", "Nova Era"])
    except SupabaseError as exc:
        st.warning(f"Não foi possível conectar ao Supabase: {exc}")
else:
    st.info(
        "Configure SUPABASE_URL e SUPABASE_SERVICE_ROLE_KEY em secrets/env para habilitar "
        "cadastros persistentes por grupo."
    )

grupos_disponiveis = []
grupo_id_selecionado = None
if supabase_client:
    try:
        grupos_disponiveis = [
            g for g in supabase_client.get_groups() if g.get("ativo", True)
        ]
    except SupabaseError as exc:
        st.error(f"Erro ao carregar grupos do Supabase: {exc}")

if grupos_disponiveis:
    nomes_grupos = [g["nome"] for g in grupos_disponiveis]
    col_grupo, col_grupo_hint = st.columns([1.2, 2.8])
    with col_grupo:
        grupo_nome_selecionado = st.selectbox("Grupo", nomes_grupos, index=0)
    with col_grupo_hint:
        st.markdown(
            """
            <div style="margin-top:1.9rem;color:#94a3b8;font-size:0.9rem;">
                Selecione o grupo para carregar e editar os cadastros fixos de
                <strong style="color:#cbd5e1;">Pro Labore</strong> e
                <strong style="color:#cbd5e1;">Domésticas</strong>.
            </div>
            """,
            unsafe_allow_html=True,
        )
    grupo_id_selecionado = next(
        (g["id"] for g in grupos_disponiveis if g["nome"] == grupo_nome_selecionado),
        None,
    )
else:
    st.warning("Nenhum grupo ativo encontrado no Supabase.")


def _carregar_registros_grupo(nome_tabela: str, grupo_id: int) -> pd.DataFrame:
    if not supabase_client or not grupo_id:
        return pd.DataFrame()

    registros = supabase_client.get_group_records(nome_tabela, grupo_id)
    if not registros:
        return pd.DataFrame(
            columns=[
                "id",
                "codigo_empresa",
                "nome_empresa",
                "codigo_empregado",
                "nome_empregado",
                "ativo",
            ]
        )
    return pd.DataFrame(registros)


def _renderizar_crud_grupo(nome_tabela: str, titulo: str, grupo_id: int):
    st.markdown(f"#### {titulo}")
    base_df = _carregar_registros_grupo(nome_tabela, grupo_id)
    if "ativo" not in base_df.columns:
        base_df["ativo"] = True

    coluna_nome = (
        "Nome do sócio" if nome_tabela == "pro_labore" else "Nome da doméstica"
    )
    visual_df = base_df.rename(columns={"nome_empregado": coluna_nome})
    edited_df = st.data_editor(
        visual_df,
        use_container_width=True,
        num_rows="dynamic",
        key=f"editor_{nome_tabela}_{grupo_id}",
        column_config={
            "id": st.column_config.NumberColumn("ID", disabled=True),
            "codigo_empresa": st.column_config.TextColumn("Código empresa"),
            "nome_empresa": st.column_config.TextColumn("Nome empresa"),
            "codigo_empregado": st.column_config.TextColumn("Código empregado"),
            coluna_nome: st.column_config.TextColumn(coluna_nome),
            "ativo": st.column_config.CheckboxColumn("Ativo"),
        },
    ).rename(columns={coluna_nome: "nome_empregado"})

    if st.button(f"Salvar {titulo}", key=f"btn_salvar_{nome_tabela}_{grupo_id}"):
        try:
            supabase_client.sync_group_records(
                nome_tabela, grupo_id, edited_df.to_dict("records")
            )
            st.success(f"{titulo} salvo com sucesso para o grupo selecionado.")
            st.rerun()
        except SupabaseError as exc:
            st.error(f"Erro ao salvar {titulo}: {exc}")


if supabase_client and grupo_id_selecionado:
    with st.expander("Gerenciar cadastros fixos do grupo (CRUD)", expanded=False):
        _renderizar_crud_grupo("pro_labore", "Pro Labore", grupo_id_selecionado)
        _renderizar_crud_grupo("domesticas", "Domésticas", grupo_id_selecionado)

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
                    partes = (
                        list(match.groups()) if match else re.split(r"\s{2,}", linha)
                    )

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
    df_preview_pro_labore = (
        _carregar_registros_grupo("pro_labore", grupo_id_selecionado)
        if grupo_id_selecionado
        else pd.DataFrame()
    )
    df_preview_domesticas = (
        _carregar_registros_grupo("domesticas", grupo_id_selecionado)
        if grupo_id_selecionado
        else pd.DataFrame()
    )

    st.markdown("#### Prévia dos afastados colados")
    if df_preview_afastados is not None and not df_preview_afastados.empty:
        st.dataframe(df_preview_afastados, use_container_width=True)
    elif texto_afastados and texto_afastados.strip():
        st.warning(
            "Não foi possível interpretar os afastados. Verifique o formato das linhas coladas."
        )
    else:
        st.info(
            "Cole os dados de afastados para visualizar a prévia antes do processamento."
        )

    st.markdown("#### Prévia dos cadastros de Pro Labore do grupo")
    if df_preview_pro_labore is not None and not df_preview_pro_labore.empty:
        st.dataframe(df_preview_pro_labore, use_container_width=True)
    else:
        st.info("Nenhum cadastro de pro labore encontrado para o grupo selecionado.")

    st.markdown("#### Prévia dos cadastros de Domésticas do grupo")
    if df_preview_domesticas is not None and not df_preview_domesticas.empty:
        st.dataframe(df_preview_domesticas, use_container_width=True)
    else:
        st.info("Nenhum cadastro de domésticas encontrado para o grupo selecionado.")

    if st.button("Processar planilha", type="primary"):
        with st.spinner("Processando dados..."):
            processador = ProcessadorEventosPeriodicos(uploaded_file)
            processador.processar()

            df_afastados = carregar_lista_afastados(texto_afastados)
            if df_afastados is not None:
                processador.marcar_afastados(df_afastados)
            df_pro_labore = df_preview_pro_labore
            if df_pro_labore is not None:
                processador.marcar_por_lista(df_pro_labore, "Pro Labore")
            df_domesticas = df_preview_domesticas
            if df_domesticas is not None:
                processador.marcar_por_lista(df_domesticas, "Doméstica")

            if processador.dados_consolidados.empty:
                st.warning("Nenhum dado foi identificado para consolidação.")
            else:
                stats = processador.calcular_estatisticas()

                col1, col2, col3, col4, col5, col6 = st.columns(6)
                col1.metric("Total de registros", stats["total_registros"])
                col2.metric("Validados", stats["total_validados"])
                col3.metric("Invalidados", stats["total_invalidados"])
                col4.metric("Afastados", stats["total_afastados"])
                col5.metric("Pro Labore", stats["total_pro_labore"])
                col6.metric("Doméstica", stats["total_domestica"])

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
