# Consolidador de Eventos Periódicos (eSocial)

Você tem duas formas de usar:

## 1) Modo mais simples (local com Streamlit)

```bash
python -m pip install -r requirements.txt
python -m streamlit run app.py
```

Depois acesse o endereço mostrado no terminal (normalmente `http://localhost:8501`).

## 2) Deploy na Vercel (FastAPI)

Este repositório inclui uma API em `api/index.py` preparada para Vercel com upload de `.xlsx` e retorno do arquivo consolidado.

### Passos

1. Suba este repositório no GitHub.
2. Na Vercel, crie um novo projeto apontando para esse repositório.
3. A Vercel vai usar o `vercel.json` automaticamente.
4. Após deploy, abra a URL do projeto e envie o arquivo.

## Fluxo de processamento

1. Faça upload da planilha de entrada.
2. O sistema identifica blocos de empresas.
3. Consolida os dados e calcula estatísticas.
4. Retorna arquivo Excel final para download.
