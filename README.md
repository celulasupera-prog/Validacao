# Consolidador de Eventos Periódicos (eSocial)

Aplicação em Streamlit para upload de planilha `.xlsx`, tratamento dos blocos de empresas e download do consolidado em Excel.
Opcionalmente, você pode enviar também uma lista de afastados para marcar o status como **Afastado**.

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
2. (Opcional) Cole os dados de afastados direto no campo de texto (Ctrl+C / Ctrl+V).
3. Clique em **Processar planilha**.
4. Visualize dados consolidados e métricas.
5. Baixe o arquivo Excel final.

## Formato da lista de afastados (opcional)

Aceita texto colado com colunas equivalentes a:
- Empresa
- Código Empregado (ou Código Funcionário/Matrícula)
- Nome (ou Nome Funcionário)

Exemplo de texto colado:
`133<TAB>IGREJA ASSEMBLEIA<TAB>1<TAB>MARIA PASTORINA DE OLIVEIRA`
