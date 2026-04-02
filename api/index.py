import io
import sys
from pathlib import Path

from fastapi import FastAPI, File, UploadFile
from fastapi.responses import HTMLResponse, JSONResponse, StreamingResponse

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.append(str(ROOT))

from processador_eventos import ProcessadorEventosPeriodicos

app = FastAPI(title="Consolidador eSocial")


@app.get("/", response_class=HTMLResponse)
def home():
    return """
    <html>
      <head><title>Consolidador eSocial</title></head>
      <body style="font-family:Arial;max-width:900px;margin:40px auto;">
        <h2>Consolidador de Eventos Periódicos (eSocial)</h2>
        <p>Envie um arquivo .xlsx para processar e baixar o consolidado.</p>
        <form action="/processar" method="post" enctype="multipart/form-data">
          <input type="file" name="arquivo" accept=".xlsx" required />
          <button type="submit">Processar</button>
        </form>
      </body>
    </html>
    """


@app.post("/processar")
async def processar_arquivo(arquivo: UploadFile = File(...)):
    if not arquivo.filename.lower().endswith(".xlsx"):
        return JSONResponse(status_code=400, content={"erro": "Envie um arquivo .xlsx"})

    conteudo = await arquivo.read()
    entrada = io.BytesIO(conteudo)

    processador = ProcessadorEventosPeriodicos(entrada)
    processador.processar()

    if processador.dados_consolidados.empty:
        return JSONResponse(status_code=422, content={"erro": "Nenhum dado encontrado para consolidação."})

    saida = io.BytesIO()
    processador.exportar_excel(saida)
    saida.seek(0)

    return StreamingResponse(
        saida,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={
            "Content-Disposition": 'attachment; filename="relacao_eventos_periodicos_consolidado.xlsx"'
        },
    )
