"""
Servicio de render de presupuestos (Scala Valencia y Vertical Projects).

Regenera el presupuesto con los datos que Miguel corrige en la aplicación, usando
exactamente el mismo código que la skill de Claude (carpeta presupuestos-scala-vertical,
copiada en la imagen), y lo convierte a PDF con Gotenberg.

Lo llama solo el servidor de la aplicación (Vercel), con la cabecera
    Authorization: Bearer <RENDER_TOKEN>
salvo las previsualizaciones, que el navegador abre con un enlace firmado que caduca.

Endpoints:
  GET  /salud                          Estado del servicio y de Gotenberg (sin autenticación).
  POST /extraer                        docx -> datos incrustados, totales y PDF del original.
  POST /render                         docx + datos editados -> presupuesto regenerado (PDF y docx).
  GET  /archivos/{id}/{nombre}         Descarga de un resultado (Bearer o enlace firmado).

Los resultados se guardan en DATOS_DIR durante CONSERVAR_HORAS y después se borran.
"""
from __future__ import annotations

import asyncio
import copy
import datetime as dt
import hashlib
import hmac
import json
import logging
import math
import os
import re
import shutil
import subprocess
import sys
import tempfile
import time
import uuid
from decimal import Decimal
from pathlib import Path
from typing import Any, Optional
from urllib.parse import quote
from zoneinfo import ZoneInfo

import httpx
from fastapi import Depends, FastAPI, File, Form, Header, HTTPException, UploadFile
from fastapi.responses import FileResponse, JSONResponse

# ---------------------------------------------------------------------------
# Configuración
# ---------------------------------------------------------------------------

SKILL = Path(os.environ.get("SKILL_DIR", "/srv/skill"))
sys.path.insert(0, str(SKILL / "scripts"))

import generar  # noqa: E402  (calcular: mismos importes que el documento)
from incrustar import VERSION_SKILL, ErrorRevision, leer, preparar_render  # noqa: E402

TOKEN = os.environ.get("RENDER_TOKEN", "")
if len(TOKEN) < 32:
    raise SystemExit("RENDER_TOKEN tiene que estar definido y tener al menos 32 caracteres.")

GOTENBERG_URL = os.environ.get("GOTENBERG_URL", "http://gotenberg:3000").rstrip("/")
URL_PUBLICA = os.environ.get("URL_PUBLICA", "https://render.advantys.ai").rstrip("/")
DATOS_DIR = Path(os.environ.get("DATOS_DIR", "/datos"))
# Orígenes que pueden mostrar la previsualización dentro de un iframe (separados por espacios).
ORIGENES_IFRAME = os.environ.get("ORIGENES_IFRAME", "").strip()

ZONA = ZoneInfo("Europe/Madrid")
CADUCIDAD_ENLACE_S = 30 * 60
CONSERVAR_HORAS = 24
MAX_DOCX_BYTES = 8 * 1024 * 1024
MAX_DATOS_BYTES = 1024 * 1024
TIEMPO_MAX_GENERAR_S = 120
TIEMPO_MAX_PDF_S = 150
NOMBRES_ARCHIVO = {"original.pdf", "presupuesto.pdf", "presupuesto.docx"}

DATOS_DIR.mkdir(parents=True, exist_ok=True)
logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
log = logging.getLogger("render")

# Dos renders a la vez como máximo (la instancia tiene 2 CPU y Gotenberg comparte máquina).
SEMAFORO = asyncio.Semaphore(2)

app = FastAPI(title="Render de presupuestos", docs_url=None, redoc_url=None, openapi_url=None)


# ---------------------------------------------------------------------------
# Utilidades
# ---------------------------------------------------------------------------

class ErrorDatos(Exception):
    def __init__(self, errores: list[str]):
        self.errores = errores


def autorizado(authorization: Optional[str] = Header(default=None)) -> None:
    esperado = f"Bearer {TOKEN}"
    if not authorization or not hmac.compare_digest(authorization.encode(), esperado.encode()):
        raise HTTPException(status_code=401, detail="No autorizado")


def calcular_firma(id_: str, nombre: str, expira: int) -> str:
    msg = f"{id_}/{nombre}:{expira}".encode()
    return hmac.new(TOKEN.encode(), msg, hashlib.sha256).hexdigest()[:40]


def enlace_firmado(id_: str, nombre: str) -> dict:
    expira = int(time.time()) + CADUCIDAD_ENLACE_S
    url = f"{URL_PUBLICA}/archivos/{id_}/{nombre}?expira={expira}&firma={calcular_firma(id_, nombre, expira)}"
    return {"url": url, "expira": dt.datetime.fromtimestamp(expira, ZONA).isoformat(timespec="seconds")}


def limpiar_antiguos() -> None:
    limite = time.time() - CONSERVAR_HORAS * 3600
    for d in DATOS_DIR.iterdir():
        try:
            if d.is_dir() and d.stat().st_mtime < limite:
                shutil.rmtree(d, ignore_errors=True)
        except FileNotFoundError:
            pass


def importe(v) -> Optional[str]:
    return None if v is None else f"{Decimal(v):.2f}"


def totales(datos: dict) -> dict:
    c = generar.calcular(datos)
    return {
        "capitulos": [{"codigo": x["codigo"], "titulo": x["titulo"], "subtotal": importe(x["subtotal"])}
                      for x in c["capitulos"]],
        "pem": importe(c["pem"]),
        "margen_porcentaje": importe(c["margen_pct"]),
        "margen_importe": importe(c["margen_importe"]),
        "base": importe(c["base"]),
        "iva_porcentaje": importe(c["iva_pct"]),
        "iva": importe(c["iva"]),
        "total": importe(c["total"]),
        "partidas_pendientes": c["pendientes"],
    }


async def leer_subida(archivo: UploadFile) -> bytes:
    contenido = await archivo.read(MAX_DOCX_BYTES + 1)
    if len(contenido) > MAX_DOCX_BYTES:
        raise HTTPException(status_code=413, detail="El documento supera el tamaño máximo admitido.")
    return contenido


async def a_pdf(docx: bytes) -> bytes:
    async with httpx.AsyncClient(timeout=TIEMPO_MAX_PDF_S) as cliente:
        r = await cliente.post(
            f"{GOTENBERG_URL}/forms/libreoffice/convert",
            files={"files": ("presupuesto.docx", docx,
                             "application/vnd.openxmlformats-officedocument.wordprocessingml.document")},
        )
    if r.status_code != 200 or not r.content.startswith(b"%PDF"):
        log.error("Gotenberg respondió %s: %s", r.status_code, r.text[:300])
        raise HTTPException(status_code=502, detail="No se ha podido convertir el documento a PDF.")
    return r.content


def nombre_descarga(datos: dict, extension: str) -> str:
    numero = str(datos.get("numero") or "").strip()
    comunidad = str((datos.get("cliente") or {}).get("comunidad") or "").strip()
    base = " - ".join(x for x in (f"Presupuesto {numero}".strip(), comunidad) if x) or "Presupuesto"
    base = re.sub(r'[\\/:*?"<>|\r\n]+', " ", base).strip()
    return f"{base}.{extension}"


# ---------------------------------------------------------------------------
# Validación de los datos editados
# ---------------------------------------------------------------------------

# Caracteres que no admite XML (los trae a veces el texto pegado desde Word o desde el móvil).
_INVALIDOS_XML = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f￾￿]")


def limpiar_texto(v: Any) -> Any:
    if isinstance(v, str):
        return _INVALIDOS_XML.sub("", v.replace("\r\n", "\n").replace("\r", "\n")).strip()
    if isinstance(v, list):
        return [limpiar_texto(x) for x in v]
    if isinstance(v, dict):
        return {k: limpiar_texto(x) for k, x in v.items()}
    return v


def numero(v: Any, donde: str, errores: list[str], minimo=0.0, maximo=1e8, nulo=False) -> Optional[float]:
    if v is None or v == "":
        if nulo:
            return None
        errores.append(f"{donde}: falta el valor.")
        return None
    if isinstance(v, bool):
        errores.append(f"{donde}: no es un número.")
        return None
    if isinstance(v, str):
        t = v.strip().replace(" ", "").replace("€", "")
        if "," in t:
            t = t.replace(".", "").replace(",", ".")  # 1.234,56 -> 1234.56
        try:
            v = float(t)
        except ValueError:
            errores.append(f"{donde}: «{v}» no es un número.")
            return None
    if not isinstance(v, (int, float)) or not math.isfinite(v):
        errores.append(f"{donde}: no es un número.")
        return None
    if v < minimo:
        errores.append(f"{donde}: no puede ser negativo." if minimo == 0 else f"{donde}: el mínimo es {minimo:g}.")
        return None
    if v > maximo:
        errores.append(f"{donde}: el máximo es {maximo:,.0f}.".replace(",", "."))
        return None
    return round(float(v), 4)


def texto(v: Any, donde: str, errores: list[str], maximo: int, obligatorio=False) -> str:
    if v is None:
        v = ""
    if not isinstance(v, str):
        errores.append(f"{donde}: tiene que ser texto.")
        return ""
    if obligatorio and not v:
        errores.append(f"{donde}: no puede quedar vacío.")
    if len(v) > maximo:
        errores.append(f"{donde}: es demasiado largo (máximo {maximo} caracteres).")
    return v


def validar(editados: Any, original: dict) -> dict:
    """Comprueba los datos editados y devuelve los datos con los que se regenera.

    Lo que Miguel no puede cambiar (empresa, número, cliente, fotos, portada, fecha) se toma
    siempre del original: así un error en la aplicación nunca puede alterarlo.
    """
    if not isinstance(editados, dict):
        raise ErrorDatos(["Los datos enviados no tienen el formato esperado."])
    d = limpiar_texto(copy.deepcopy(editados))
    errores: list[str] = []

    for campo in ("empresa", "numero", "cliente", "fotos", "foto_portada", "validez_dias"):
        if campo in original:
            d[campo] = copy.deepcopy(original[campo])
        else:
            d.pop(campo, None)
    d["fecha"] = dt.datetime.now(ZONA).strftime("%d/%m/%Y")  # D2: fecha de validación

    d["titulo"] = texto(d.get("titulo"), "Título del presupuesto", errores, 200, obligatorio=True)
    if "plazo" in d:
        d["plazo"] = texto(d.get("plazo"), "Plazo de ejecución", errores, 60)

    iva = numero(d.get("iva_porcentaje"), "IVA", errores, 0, 100)
    if iva is not None:
        d["iva_porcentaje"] = iva
    margen = d.get("margen") or {}
    if not isinstance(margen, dict):
        errores.append("Margen: formato no válido.")
        margen = {}
    pct = numero(margen.get("porcentaje", 0), "Margen", errores, 0, 100)
    base = str(margen.get("base") or "partida").lower()
    if base not in ("partida", "total"):
        errores.append("Margen: la base tiene que ser «partida» o «total».")
    margen.update(porcentaje=pct or 0, base=base)
    d["margen"] = margen

    capitulos = d.get("capitulos")
    if not isinstance(capitulos, list) or not capitulos:
        errores.append("El presupuesto tiene que tener al menos un capítulo.")
        capitulos = []
    codigos: set[str] = set()
    for i, c in enumerate(capitulos, 1):
        dc = f"Capítulo {i}"
        if not isinstance(c, dict):
            errores.append(f"{dc}: formato no válido.")
            continue
        c["codigo"] = texto(c.get("codigo"), f"{dc}, código", errores, 20, obligatorio=True)
        c["titulo"] = texto(c.get("titulo"), f"{dc}, título", errores, 200, obligatorio=True)
        partidas = c.get("partidas")
        if not isinstance(partidas, list) or not partidas:
            errores.append(f"{dc} ({c.get('titulo') or 'sin título'}): tiene que tener al menos una partida.")
            continue
        for j, p in enumerate(partidas, 1):
            dp = f"{dc}, partida {j}"
            if not isinstance(p, dict):
                errores.append(f"{dp}: formato no válido.")
                continue
            p.pop("titulo", None)
            p.pop("medicion", None)
            p["codigo"] = texto(p.get("codigo"), f"{dp}, código", errores, 20, obligatorio=True)
            if p["codigo"] in codigos:
                errores.append(f"{dp}: el código {p['codigo']} está repetido.")
            codigos.add(p["codigo"])
            p["resumen"] = texto(p.get("resumen"), f"{dp}, descripción corta", errores, 400, obligatorio=True)
            p["descripcion"] = texto(p.get("descripcion"), f"{dp}, descripción larga", errores, 4000)
            p["ud"] = texto(p.get("ud"), f"{dp}, unidad", errores, 20, obligatorio=True)
            cantidad = numero(p.get("cantidad"), f"{dp}, cantidad", errores, 0, 1e7)
            if cantidad is not None:
                p["cantidad"] = cantidad
            p["pendiente"] = bool(p.get("pendiente"))
            if p["pendiente"]:
                p["precio"] = None
            else:
                precio = numero(p.get("precio"), f"{dp}, precio", errores, 0, 1e7)
                if precio is not None:
                    p["precio"] = precio
    d["capitulos"] = capitulos

    if errores:
        raise ErrorDatos(errores)
    return d


# ---------------------------------------------------------------------------
# Endpoints
# ---------------------------------------------------------------------------

@app.exception_handler(ErrorRevision)
async def _error_revision(_, e: ErrorRevision):
    return JSONResponse(status_code=422, content={"error": str(e)})


@app.exception_handler(ErrorDatos)
async def _error_datos(_, e: ErrorDatos):
    return JSONResponse(status_code=422, content={"error": "Hay datos que corregir.", "errores": e.errores})


@app.get("/salud")
async def salud():
    try:
        async with httpx.AsyncClient(timeout=5) as cliente:
            g = (await cliente.get(f"{GOTENBERG_URL}/health")).json().get("status")
    except Exception:
        g = "down"
    return {"estado": "ok" if g == "up" else "degradado", "version_skill": VERSION_SKILL, "gotenberg": g}


@app.post("/extraer", dependencies=[Depends(autorizado)])
async def extraer(docx: UploadFile = File(...)):
    limpiar_antiguos()
    contenido = await leer_subida(docx)
    sobre = leer(contenido)
    datos = sobre["datos"]
    id_ = uuid.uuid4().hex
    async with SEMAFORO:
        pdf = await a_pdf(contenido)
    carpeta = DATOS_DIR / id_
    carpeta.mkdir()
    (carpeta / "original.pdf").write_bytes(pdf)
    (carpeta / "meta.json").write_text(json.dumps({"nombre_pdf": nombre_descarga(datos, "pdf")}))
    log.info("extraer id=%s empresa=%s numero=%s", id_, datos.get("empresa"), datos.get("numero"))
    return {
        "id": id_,
        "version_skill": sobre["version_skill"],
        "generado": sobre["generado"],
        "datos": datos,
        "totales": totales(datos),
        "previsualizacion": enlace_firmado(id_, "original.pdf"),
    }


@app.post("/render", dependencies=[Depends(autorizado)])
async def render(docx: UploadFile = File(...), datos: str = Form(...)):
    limpiar_antiguos()
    inicio = time.monotonic()
    original_docx = await leer_subida(docx)
    original = leer(original_docx)["datos"]  # valida que el original sigue siendo revisable
    if len(datos.encode()) > MAX_DATOS_BYTES:
        raise HTTPException(status_code=413, detail="Los datos enviados son demasiado grandes.")
    try:
        editados = json.loads(datos.lstrip("\ufeff"))  # algunos editores de Windows añaden BOM
    except ValueError:
        raise ErrorDatos(["Los datos enviados no son un JSON válido."])
    d = validar(editados, original)

    id_ = uuid.uuid4().hex
    carpeta = DATOS_DIR / id_
    carpeta.mkdir()
    async with SEMAFORO:
        with tempfile.TemporaryDirectory() as tmp:
            listos = preparar_render(original_docx, d, os.path.join(tmp, "fotos"))
            ruta_json = os.path.join(tmp, "datos.json")
            ruta_docx = os.path.join(tmp, "presupuesto.docx")
            with open(ruta_json, "w", encoding="utf-8") as fh:
                json.dump(listos, fh, ensure_ascii=False)
            try:
                r = await asyncio.to_thread(
                    subprocess.run,
                    [sys.executable, str(SKILL / "scripts" / "generar.py"), ruta_json, ruta_docx],
                    capture_output=True, text=True, timeout=TIEMPO_MAX_GENERAR_S,
                )
            except subprocess.TimeoutExpired:
                shutil.rmtree(carpeta, ignore_errors=True)
                raise HTTPException(status_code=504, detail="La generación del documento ha tardado demasiado.")
            if r.returncode != 0 or not os.path.exists(ruta_docx):
                shutil.rmtree(carpeta, ignore_errors=True)
                log.error("generar.py falló id=%s: %s", id_, (r.stderr or r.stdout)[-1500:])
                raise HTTPException(status_code=500, detail="No se ha podido generar el documento.")
            nuevo_docx = Path(ruta_docx).read_bytes()
        avisos = [x[len("AVISO:"):].strip() for x in r.stdout.splitlines() if x.startswith("AVISO:")]
        regenerado = leer(nuevo_docx)["datos"]  # comprueba que el resultado vuelve a ser revisable
        pdf = await a_pdf(nuevo_docx)

    (carpeta / "presupuesto.docx").write_bytes(nuevo_docx)
    (carpeta / "presupuesto.pdf").write_bytes(pdf)
    (carpeta / "meta.json").write_text(json.dumps({
        "nombre_pdf": nombre_descarga(regenerado, "pdf"),
        "nombre_docx": nombre_descarga(regenerado, "docx"),
    }))
    log.info("render id=%s empresa=%s numero=%s %.1fs", id_, regenerado.get("empresa"),
             regenerado.get("numero"), time.monotonic() - inicio)
    return {
        "id": id_,
        "fecha": regenerado["fecha"],
        "datos": regenerado,
        "totales": totales(regenerado),
        "avisos": avisos,
        "previsualizacion": enlace_firmado(id_, "presupuesto.pdf"),
        "nombre_pdf": nombre_descarga(regenerado, "pdf"),
        "nombre_docx": nombre_descarga(regenerado, "docx"),
    }


@app.get("/archivos/{id_}/{nombre}")
async def descargar(id_: str, nombre: str, expira: Optional[int] = None, firma: Optional[str] = None,
                    authorization: Optional[str] = Header(default=None)):
    if not re.fullmatch(r"[0-9a-f]{32}", id_) or nombre not in NOMBRES_ARCHIVO:
        raise HTTPException(status_code=404, detail="No encontrado")
    con_token = bool(authorization) and hmac.compare_digest(authorization.encode(), f"Bearer {TOKEN}".encode())
    con_firma = (
        nombre.endswith(".pdf") and expira is not None and firma is not None and expira >= time.time()
        and hmac.compare_digest(firma, calcular_firma(id_, nombre, expira))
    )
    if not (con_token or con_firma):
        raise HTTPException(status_code=403, detail="Enlace no válido o caducado")
    ruta = DATOS_DIR / id_ / nombre
    if not ruta.exists():
        raise HTTPException(status_code=404, detail="El archivo ya no está disponible. Vuelve a previsualizar.")
    try:
        meta = json.loads((DATOS_DIR / id_ / "meta.json").read_text())
    except Exception:
        meta = {}
    descarga = meta.get("nombre_docx" if nombre.endswith(".docx") else "nombre_pdf") or nombre
    tipo = ("application/pdf" if nombre.endswith(".pdf") else
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
    cabeceras = {
        "Content-Disposition": f"inline; filename*=UTF-8''{quote(descarga)}",
        "Cache-Control": "private, no-store",
        "X-Content-Type-Options": "nosniff",
        "Content-Security-Policy": f"frame-ancestors 'self' {ORIGENES_IFRAME}".strip(),
    }
    return FileResponse(ruta, media_type=tipo, headers=cabeceras)
