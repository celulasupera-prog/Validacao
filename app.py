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
    :root {
        --bg: #020817;
        --panel: #0b1220;
        --panel-2: #111827;
        --border: rgba(148, 163, 184, 0.14);
        --text: #f8fafc;
        --muted: #94a3b8;
        --primary: #3b82f6;
        --primary-hover: #2563eb;
        --shadow: 0 18px 50px rgba(0,0,0,0.35);
    }
    .main {
        background:
            radial-gradient(circle at top, rgba(59,130,246,0.14), transparent 28%),
            linear-gradient(180deg, #030712 0%, #020817 100%);
        color: var(--text);
    }
    .hero {
        border: 1px solid var(--border);
        border-radius: 22px;
        padding: 1.2rem 1.35rem;
        margin-bottom: 1rem;
        background: linear-gradient(135deg, rgba(15,23,42,0.96), rgba(9,17,33,0.95));
        color: var(--text);
        box-shadow: var(--shadow);
    }
    .hero h1 {
        font-size: 1.35rem;
        margin: 0 0 0.35rem 0;
    }
    .hero p {
        margin: 0;
        color: var(--muted);
    }
    .section-card {
        border: 1px solid var(--border);
        border-radius: 16px;
        background: linear-gradient(180deg, rgba(15,23,42,0.96), rgba(10,15,28,0.94));
        color: var(--text);
        padding: 0.8rem 0.95rem 0.3rem 0.95rem;
        margin-bottom: 0.75rem;
        box-shadow: var(--shadow);
    }
    .section-card strong, .section-card em, .section-card span, .section-card p {
        color: var(--text) !important;
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
        stroke: #bfdbfe;
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
        border-radius: 12px;
        font-weight: 600;
        background: var(--primary);
        color: #fff;
        border: 1px solid var(--primary);
    }
    .stDownloadButton button:hover {
        background: var(--primary-hover);
        border-color: var(--primary-hover);
    }
    .stButton button {
        border-radius: 12px;
        font-weight: 700;
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
    <div class="hero">
        <div style="display:flex;justify-content:space-between;gap:16px;align-items:flex-start;flex-wrap:wrap;">
        <div style="max-width:760px;">
            <div style="display:inline-flex;align-items:center;gap:8px;padding:6px 10px;border-radius:999px;border:1px solid rgba(96,165,250,0.22);background:rgba(59,130,246,0.08);color:#bfdbfe;font-size:12px;font-weight:600;margin-bottom:10px;">
                Consolidação inteligente
            </div>
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
        <div style="display:grid;grid-template-columns:repeat(3,minmax(96px,1fr));gap:8px;min-width:280px;">
            <div style="background:rgba(255,255,255,0.03);border:1px solid var(--border);padding:10px;border-radius:12px;"><span style="display:block;color:var(--muted);font-size:11px;">Formato</span><strong>.XLSX</strong></div>
            <div style="background:rgba(255,255,255,0.03);border:1px solid var(--border);padding:10px;border-radius:12px;"><span style="display:block;color:var(--muted);font-size:11px;">Máx.</span><strong>200MB</strong></div>
            <div style="background:rgba(255,255,255,0.03);border:1px solid var(--border);padding:10px;border-radius:12px;"><span style="display:block;color:var(--muted);font-size:11px;">Status</span><strong>Pronto</strong></div>
        </div>
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)

uploaded_file = None
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
    uploaded_file = st.file_uploader(
        "Selecione a planilha de entrada",
        type=["xlsx"],
        label_visibility="collapsed",
    )
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
