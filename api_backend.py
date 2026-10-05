import os
import re
import json
import uuid
from pathlib import Path

from dotenv import load_dotenv
from fastapi import FastAPI, UploadFile, File, HTTPException, Body
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, StreamingResponse, HTMLResponse, Response
from fastapi.staticfiles import StaticFiles
from sqlalchemy import text
from sqlalchemy.orm import Session

from base import engine
from oracle_db import aceptar_registros
from pipeline import estado_ollama, liberar_modelos_ollama, procesar_archivo_interno, revalidar_registro
from reportes import (
    TIPOS_CARROCERIA,
    cargar_corrida,
    escribir_reciente,
    excel_exitos,
    excel_revision,
    fila_resumen,
    iniciar_corrida,
    modelo_en_catalogo,
    normalizar_tipo,
    registrar_modelo,
    registros_historico,
)

load_dotenv()

app = FastAPI(title="ManPAC Extractor Automotriz", version="2.0")
app.add_middleware(
    CORSMiddleware,
    allow_origins=["http://127.0.0.1:8000", "http://localhost:8000"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)
CARPETA_TEMP = "temp_api_uploads"
CARPETAS_DOCUMENTO = ("procesados_exito", "revision_manual")
DIR_WEB = Path(__file__).resolve().parent / "web"


def nombre_seguro(nombre):
    base = os.path.basename(nombre or "documento")
    base = re.sub(r"[^\w.\- ]+", "_", base).strip("._ ") or "documento"
    return base[:180]


def _es_exito(reg):
    est = str(reg.get("Estado") or reg.get("estado") or "")
    return "ÉXITO" in est or est.upper().startswith("EXITO")


def aplicar_edicion(registros, filas):
    fusion = []
    for i, fila in enumerate(filas):
        base = dict(registros[i]) if i < len(registros) else {}
        base["fecha_ejecucion"] = str(fila.get("Fecha_Ejecucion") or base.get("fecha_ejecucion") or "")
        base["marca"] = str(fila.get("Marca") or "")
        base["modelo"] = str(fila.get("Modelo") or "")
        base["cantidad"] = fila.get("cantidad") or 1
        base["beneficiario"] = str(fila.get("beneficiario") or "")
        base["ci"] = str(fila.get("ci") or "")
        base["costo_mas_alto"] = fila.get("costo_mas_alto") or 0
        base["costo_total"] = fila.get("costo_total") or 0
        base["marca_base"] = str(fila.get("marca_base") or "")
        base["modelo_base"] = str(fila.get("modelo_base") or "")
        base["tipo"] = normalizar_tipo(fila.get("Tipo") or base.get("tipo"))
        base["archivo"] = str(fila.get("archivo") or base.get("archivo") or "")
        estado_ui = str(fila.get("Estado") or "")
        if estado_ui.upper().startswith("EXITO"):
            base["estado"] = "ÉXITO"
        elif estado_ui:
            base["estado"] = estado_ui if estado_ui.upper().startswith("REVISION") else f"REVISIÓN ({estado_ui})"
        fusion.append(base)
    return fusion


@app.get("/", response_class=HTMLResponse)
def inicio():
    pagina = DIR_WEB / "templates" / "index.html"
    if not pagina.is_file():
        raise HTTPException(500, "No se encontró la interfaz web.")
    return HTMLResponse(pagina.read_text(encoding="utf-8"))


@app.post("/api/procesar-factura")
async def endpoint_procesar_factura(file: UploadFile = File(...)):
    os.makedirs(CARPETA_TEMP, exist_ok=True)
    nombre = nombre_seguro(file.filename)
    marca = uuid.uuid4().hex
    ruta_temp = os.path.join(CARPETA_TEMP, f"{marca}_{nombre}")

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
            fallo = {
                "estado": f"REVISIÓN (Error de procesamiento: {e})",
                "archivo": nombre,
                "paginas_pdf": [],
                "error": str(e),
            }
            try:
                from reportes import anexar_corrida
                anexar_corrida([fallo])
            except Exception:
                pass
            yield json.dumps(fallo, ensure_ascii=False, default=str) + "\n"
        finally:
            if os.path.exists(ruta_temp):
                os.remove(ruta_temp)

    return StreamingResponse(generador_resultados(), media_type="application/x-ndjson")


@app.get("/api/documento")
def ver_documento(nombre: str):
    limpio = os.path.basename(nombre or "")
    if not limpio or limpio in (".", "..") or limpio != (nombre or "").replace("\\", "/").split("/")[-1]:
        raise HTTPException(status_code=400, detail="Nombre de archivo no permitido.")
    for carpeta in CARPETAS_DOCUMENTO:
        base = os.path.abspath(carpeta)
        ruta = os.path.abspath(os.path.join(base, limpio))
        if os.path.normcase(os.path.commonpath([base, ruta])) != os.path.normcase(base):
            continue
        if os.path.isfile(ruta):
            ext = os.path.splitext(limpio)[1].lower()
            media = {
                ".pdf": "application/pdf",
                ".png": "image/png",
                ".jpg": "image/jpeg",
                ".jpeg": "image/jpeg",
                ".webp": "image/webp",
            }.get(ext, "application/octet-stream")
            return FileResponse(
                ruta,
                media_type=media,
                filename=limpio,
                content_disposition_type="inline",
            )
    raise HTTPException(status_code=404, detail="Documento no encontrado.")


@app.get("/api/health")
def health_check():
    return {"status": "ok", "pipeline": True, **estado_ollama()}


@app.post("/api/corrida/iniciar")
def api_iniciar_corrida():
    iniciar_corrida()
    return {"ok": True, "registros": []}


@app.post("/api/corrida/finalizar")
def api_finalizar_corrida():
    liberar_modelos_ollama()
    return {"ok": True, "n": len(cargar_corrida())}


@app.get("/api/corrida")
def api_corrida():
    return {"registros": cargar_corrida(), "tipos": list(TIPOS_CARROCERIA)}


@app.put("/api/corrida")
def api_guardar_corrida(payload: dict = Body(...)):
    actuales = cargar_corrida()
    filas = payload.get("filas") or payload.get("registros") or []
    if filas and isinstance(filas[0], dict) and "Marca" in filas[0]:
        fusion = aplicar_edicion(actuales, filas)
    else:
        fusion = filas or actuales
    escribir_reciente(fusion)
    return {"ok": True, "n": len(fusion)}


@app.post("/api/corrida/validar")
def api_validar_corrida(payload: dict = Body(...)):
    actuales = cargar_corrida()
    idx = payload.get("indice")
    nombre = os.path.basename(str(payload.get("archivo_guardado") or payload.get("archivo") or ""))
    if idx is None and nombre:
        for i, r in enumerate(actuales):
            guardado = os.path.basename(str(r.get("archivo_guardado") or r.get("archivo") or ""))
            if guardado == nombre:
                idx = i
                break
    if idx is None:
        raise HTTPException(404, "Registro no encontrado en la corrida.")
    idx = int(idx)
    if idx < 0 or idx >= len(actuales):
        raise HTTPException(404, "Registro no encontrado en la corrida.")
    base = dict(actuales[idx])
    mapa = {
        "Marca": "marca",
        "Modelo": "modelo",
        "Fecha_Ejecucion": "fecha_ejecucion",
        "Tipo": "tipo",
        "cantidad": "cantidad",
        "beneficiario": "beneficiario",
        "ci": "ci",
        "costo_mas_alto": "costo_mas_alto",
        "costo_total": "costo_total",
        "items": "items",
        "archivo": "archivo",
    }
    for clave, dest in mapa.items():
        if clave in payload:
            base[dest] = payload[clave]
        elif dest in payload:
            base[dest] = payload[dest]
    actualizado = revalidar_registro(base)
    actuales[idx] = actualizado
    escribir_reciente(actuales)
    return {"ok": True, "registro": actualizado, "exito": _es_exito(actualizado), "indice": idx}


@app.get("/api/corrida/excel")
def api_excel_corrida(tipo: str = "exitos"):
    registros = cargar_corrida()
    if tipo == "revision":
        data = excel_revision([r for r in registros if not _es_exito(r)])
        nombre = "reporte_revision_reciente.xlsx"
    else:
        data = excel_exitos([r for r in registros if _es_exito(r)])
        nombre = "reporte_exitos_reciente.xlsx"
    return Response(
        content=data,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{nombre}"'},
    )


@app.post("/api/corrida/aceptar")
def api_aceptar_corrida(payload: dict = Body(default=None)):
    registros = [r for r in cargar_corrida() if _es_exito(r)]
    if not registros:
        return {"ok": False, "mensaje": "No hay éxitos en la corrida reciente para aceptar."}
    ok, msg = aceptar_registros([fila_resumen(r) for r in registros])
    return {"ok": ok, "mensaje": msg}


@app.get("/api/historico")
def api_historico():
    ex, rev = registros_historico()
    return {"exitos": ex, "revision": rev}


@app.get("/api/historico/excel")
def api_excel_historico(tipo: str = "exitos"):
    ex, rev = registros_historico()
    if tipo == "revision":
        data = excel_revision(rev)
        nombre = "historico_revision.xlsx"
    else:
        data = excel_exitos(ex)
        nombre = "historico_exitos.xlsx"
    return Response(
        content=data,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{nombre}"'},
    )


@app.get("/api/catalogo")
def api_catalogo():
    with Session(engine) as session:
        marcas = [
            {"id_marca": r[0], "nombre": r[1], "origen": r[2] or ""}
            for r in session.execute(text("SELECT id_marca, nombre, origen FROM marca ORDER BY nombre")).all()
        ]
        modelos = [
            {
                "id_modelo": r[0],
                "id_marca": r[1],
                "nombre": r[2],
                "tipo_carroceria": r[3] or "",
            }
            for r in session.execute(
                text("SELECT id_modelo, id_marca, nombre, tipo_carroceria FROM modelo ORDER BY nombre")
            ).all()
        ]
    return {"marcas": marcas, "modelos": modelos, "tipos": list(TIPOS_CARROCERIA)}


@app.post("/api/catalogo/marcas")
def api_alta_marca(payload: dict = Body(...)):
    nombre = str(payload.get("nombre") or "").strip().upper()
    origen = str(payload.get("origen") or "").strip()
    if not nombre:
        raise HTTPException(400, "La marca es obligatoria.")
    with Session(engine) as session:
        session.execute(text("INSERT INTO marca (nombre, origen) VALUES (:n, :o)"), {"n": nombre, "o": origen})
        session.commit()
    return {"ok": True}


@app.put("/api/catalogo/marcas/{id_marca}")
def api_editar_marca(id_marca: int, payload: dict = Body(...)):
    with Session(engine) as session:
        session.execute(
            text("UPDATE marca SET nombre = :n, origen = :o WHERE id_marca = :id"),
            {
                "n": str(payload.get("nombre") or "").strip().upper(),
                "o": str(payload.get("origen") or "").strip(),
                "id": id_marca,
            },
        )
        session.commit()
    return {"ok": True}


@app.delete("/api/catalogo/marcas/{id_marca}")
def api_borrar_marca(id_marca: int):
    with Session(engine) as session:
        session.execute(text("DELETE FROM modelo WHERE id_marca = :id"), {"id": id_marca})
        session.execute(text("DELETE FROM marca WHERE id_marca = :id"), {"id": id_marca})
        session.commit()
    return {"ok": True}


@app.post("/api/catalogo/modelos")
def api_alta_modelo(payload: dict = Body(...)):
    try:
        tipo, aviso = registrar_modelo(payload.get("marca"), payload.get("nombre") or payload.get("modelo"), payload.get("tipo_carroceria") or payload.get("tipo"))
    except ValueError as exc:
        raise HTTPException(400, str(exc)) from exc
    return {"ok": True, "tipo": tipo, "aviso": aviso}


@app.put("/api/catalogo/modelos/{id_modelo}")
def api_editar_modelo(id_modelo: int, payload: dict = Body(...)):
    with Session(engine) as session:
        session.execute(
            text("UPDATE modelo SET nombre = :n, tipo_carroceria = :t WHERE id_modelo = :id"),
            {
                "n": str(payload.get("nombre") or "").strip().upper(),
                "t": normalizar_tipo(payload.get("tipo_carroceria") or payload.get("tipo")),
                "id": id_modelo,
            },
        )
        session.commit()
    return {"ok": True}


@app.delete("/api/catalogo/modelos/{id_modelo}")
def api_borrar_modelo(id_modelo: int):
    with Session(engine) as session:
        session.execute(text("DELETE FROM modelo WHERE id_modelo = :id"), {"id": id_modelo})
        session.commit()
    return {"ok": True}


@app.post("/api/catalogo/registrar")
def api_registrar_revision(payload: dict = Body(...)):
    try:
        tipo, aviso = registrar_modelo(payload.get("marca"), payload.get("modelo"), payload.get("tipo") or "OTRO")
    except ValueError as exc:
        raise HTTPException(400, str(exc)) from exc
    return {"ok": True, "tipo": tipo, "aviso": aviso, "en_catalogo": modelo_en_catalogo(payload.get("modelo"))}


if DIR_WEB.joinpath("static").is_dir():
    app.mount("/static", StaticFiles(directory=str(DIR_WEB / "static")), name="static")


if __name__ == "__main__":
    import uvicorn
    print("ManPAC UI + API en http://127.0.0.1:8000")
    uvicorn.run(app, host="127.0.0.1", port=8000)
