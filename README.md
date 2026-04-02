# Consolidador de Eventos Periódicos (eSocial)

Aplicação em Streamlit para upload de planilha `.xlsx`, tratamento dos blocos de empresas e download do consolidado em Excel.

## Executar localmente

```bash
python -m pip install -r requirements.txt
python -m streamlit run app.py
```

Depois acesse o endereço mostrado no terminal (normalmente `http://localhost:8501`).

## Deploy na Streamlit Community Cloud

1. Suba este projeto para um repositório no GitHub.
2. Acesse [share.streamlit.io](https://share.streamlit.io/) e clique em **New app**.
3. Selecione o repositório, branch e o arquivo principal: `app.py`.
4. Clique em **Deploy**.

A plataforma instala automaticamente as dependências de `requirements.txt`.

## Fluxo de uso

1. Faça upload da planilha de entrada.
2. Clique em **Processar planilha**.
3. Visualize dados consolidados e métricas.
4. Baixe o arquivo Excel final.
