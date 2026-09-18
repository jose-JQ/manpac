import os
import re
import json
import uuid
from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import StreamingResponse
import uvicorn

from pipeline import procesar_archivo_interno

app = FastAPI(title="API de Extracción Automotriz IA", version="1.0")
CARPETA_TEMP = "temp_api_uploads"


def nombre_seguro(nombre):
    base = os.path.basename(nombre or "documento")
    base = re.sub(r"[^\w.\- ]+", "_", base).strip("._ ") or "documento"
    return base[:180]


@app.post("/api/procesar-factura")
async def endpoint_procesar_factura(file: UploadFile = File(...)):
    """
    Recibe un archivo y devuelve un flujo NDJSON con cada factura detectada.
    """
    os.makedirs(CARPETA_TEMP, exist_ok=True)
    nombre = nombre_seguro(file.filename)
    ruta_temp = os.path.join(CARPETA_TEMP, f"{uuid.uuid4().hex}_{nombre}")

    contenido = await file.read()
    if not contenido:
        raise HTTPException(status_code=400, detail="El archivo está vacío.")

    with open(ruta_temp, "wb") as buffer:
        buffer.write(contenido)

    def generador_resultados():
        try:
            for resultado in procesar_archivo_interno(ruta_temp):
                yield json.dumps(resultado, ensure_ascii=False, default=str) + "\n"
        except Exception as e:
            yield json.dumps({
                "estado": f"REVISIÓN (Error de procesamiento: {e})",
                "archivo": nombre,
                "paginas_pdf": [],
                "error": str(e),
            }, ensure_ascii=False, default=str) + "\n"
        finally:
            if os.path.exists(ruta_temp):
                os.remove(ruta_temp)

    return StreamingResponse(generador_resultados(), media_type="application/x-ndjson")


@app.get("/api/health")
def health_check():
    return {"status": "Servidor Activo", "pipeline": "Conectado", "modelos": "Listos para cargar bajo demanda"}


if __name__ == "__main__":
    print("Iniciando API...")
    uvicorn.run(app, host="127.0.0.1", port=8000)
