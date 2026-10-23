"""
Datos incrustados en el presupuesto (.docx) para poder revisarlo y regenerarlo
sin volver a pasar por Claude.

Al generar, se añaden al .docx dos partes:

  advantys/datos.json          Los datos con los que se ha generado, normalizados:
                               precios finales (con el margen por partida ya aplicado),
                               fecha resuelta y fotos referenciadas por su archivo
                               dentro de word/media/.
  advantys/portada-fuente.jpg  La foto de portada a resolución de portada (la del anexo
                               va más comprimida y en la portada se vería pixelada).

Word elimina estas partes si alguien abre y guarda el documento. Es intencionado:
un .docx retocado a mano no coincide con sus datos y no se puede revisar en la
aplicación. Para detectarlo se guarda además la huella de word/document.xml.

Funciones:
  portada_fuente()   Prepara la foto de portada que se incrusta.
  incrustar()        Añade las dos partes a un .docx recién generado (lo llama generar.py).
  leer()             Lee y valida los datos incrustados (aplicación y servicio de render).
  preparar_render()  Extrae las fotos a una carpeta y devuelve los datos listos para
                     volver a ejecutar generar.py (servicio de render).
"""
from __future__ import annotations

import copy
import datetime as dt
import hashlib
import io
import json
import os
import zipfile
from decimal import Decimal
from typing import Optional

from PIL import Image, ImageOps

VERSION_SKILL = "1.1.0"
ESQUEMA = 1

RUTA_DATOS = "advantys/datos.json"
RUTA_PORTADA = "advantys/portada-fuente.jpg"

# La portada se dibuja a 1860 px de ancho (portada.py, ANCHO_SALIDA).
LADO_PORTADA = 2000
CALIDAD_PORTADA = 82

TIPOS_CONTENIDO = {
    RUTA_DATOS: "application/json",
    RUTA_PORTADA: "image/jpeg",
}


class ErrorRevision(Exception):
    """El .docx no se puede revisar: el mensaje está pensado para mostrarse al usuario."""


# ---------------------------------------------------------------------------
# Generación
# ---------------------------------------------------------------------------

def portada_fuente(ruta: Optional[str]) -> Optional[bytes]:
    """JPEG de la foto de portada con el lado largo a LADO_PORTADA px como máximo.

    Si la foto ya es un JPEG de ese tamaño o menor (caso de una regeneración, que parte
    de la portada-fuente incrustada), se devuelve tal cual para no recomprimirla cada vez.
    """
    if not ruta or not os.path.exists(ruta):
        return None
    try:
        with Image.open(ruta) as im:
            if im.format == "JPEG" and max(im.size) <= LADO_PORTADA:
                with open(ruta, "rb") as fh:
                    return fh.read()
            im = ImageOps.exif_transpose(im).convert("RGB")
            im.thumbnail((LADO_PORTADA, LADO_PORTADA), Image.LANCZOS)
            buf = io.BytesIO()
            im.save(buf, "JPEG", quality=CALIDAD_PORTADA, optimize=True, progressive=True)
            return buf.getvalue()
    except Exception:
        return None


def _num(v: Optional[Decimal]):
    """Decimal -> número JSON (entero si no tiene decimales)."""
    if v is None:
        return None
    v = Decimal(v)
    return int(v) if v == v.to_integral_value() else float(v)


def normalizar(datos: dict, calc: dict, fechas: dict, fotos: list, ruta_media: dict,
               hay_portada: bool) -> dict:
    """Datos tal como se guardan en el .docx.

    - fecha: la de emisión resuelta (nunca vacía).
    - margen por partida: ya aplicado en cada precio; se guarda con porcentaje 0 y el
      porcentaje original en "aplicado_en_precios", para que al regenerar no se aplique dos veces.
    - margen sobre total: se mantiene igual (sale como línea aparte en el documento).
    - partidas: claves canónicas (resumen, cantidad, precio, pendiente).
    - fotos: solo las que han entrado en el anexo, con la ruta de su archivo en word/media/.
    """
    d = copy.deepcopy(datos)
    d["fecha"] = fechas["emision"]
    d.pop("iva", None)
    d["iva_porcentaje"] = _num(calc["iva_pct"])

    margen = dict(d.get("margen") or {})
    pct = margen.get("porcentaje") or 0
    base = (margen.get("base") or "partida").lower()
    if pct and base.startswith("partida"):
        d["margen"] = {"porcentaje": 0, "base": "partida", "aplicado_en_precios": pct}

    capitulos = []
    for c_orig, c_calc in zip(datos["capitulos"], calc["capitulos"]):
        c = {k: v for k, v in c_orig.items() if k not in ("partidas", "num")}
        c["codigo"] = c_calc["codigo"]
        c["titulo"] = c_calc["titulo"]
        partidas = []
        for p_orig, p_calc in zip(c_orig.get("partidas", []), c_calc["partidas"]):
            p = {k: v for k, v in p_orig.items() if k not in ("titulo", "medicion")}
            p.update(
                codigo=p_calc["codigo"],
                resumen=p_calc["resumen"],
                descripcion=p_calc["descripcion"],
                ud=p_calc["ud"],
                cantidad=_num(p_calc["cantidad"]),
                pendiente=bool(p_calc["pendiente"]),
                precio=None if p_calc["pendiente"] else _num(p_calc["precio"]),
            )
            partidas.append(p)
        c["partidas"] = partidas
        capitulos.append(c)
    d["capitulos"] = capitulos

    d["fotos"] = []
    for f in fotos:
        if f["ruta"] in ruta_media:
            nueva = {k: v for k, v in f.items() if k != "ruta"}
            nueva["ruta"] = ruta_media[f["ruta"]]
            d["fotos"].append(nueva)
    d["foto_portada"] = RUTA_PORTADA if hay_portada else None
    return d


def _con_tipos(content_types: bytes) -> bytes:
    """Declara las partes nuevas en [Content_Types].xml (si no, Word puede dar el archivo por dañado)."""
    xml = content_types.decode("utf-8")
    extra = ""
    for parte, tipo in TIPOS_CONTENIDO.items():
        if f'PartName="/{parte}"' not in xml:
            extra += f'<Override PartName="/{parte}" ContentType="{tipo}"/>'
    return xml.replace("</Types>", extra + "</Types>").encode("utf-8")


def incrustar(docx: bytes, datos: dict, calc: dict, fechas: dict, fotos: list,
              registro: dict, portada: Optional[bytes]) -> tuple[bytes, list[str]]:
    """Devuelve el .docx con los datos incrustados y una lista de avisos."""
    avisos: list[str] = []
    with zipfile.ZipFile(io.BytesIO(docx)) as z:
        entradas = [(info, z.read(info.filename)) for info in z.infolist()
                    if info.filename not in TIPOS_CONTENIDO]

    contenido = {info.filename: datos_ for info, datos_ in entradas}
    por_huella = {hashlib.sha1(b).hexdigest(): nombre
                  for nombre, b in contenido.items() if nombre.startswith("word/media/")}
    ruta_media = {ruta: por_huella[h] for ruta, h in registro.items() if h in por_huella}

    sin_localizar = [f["ruta"] for f in fotos if f["ruta"] not in ruta_media]
    if sin_localizar:
        avisos.append(f"{len(sin_localizar)} foto(s) del anexo no se han localizado dentro del .docx; "
                      "no estarán disponibles al regenerarlo en la aplicación.")

    sobre = {
        "esquema": ESQUEMA,
        "version_skill": VERSION_SKILL,
        "generado": dt.datetime.now(dt.timezone.utc).isoformat(timespec="seconds"),
        "huella_documento": hashlib.sha256(contenido["word/document.xml"]).hexdigest(),
        "datos": normalizar(datos, calc, fechas, fotos, ruta_media, portada is not None),
    }

    salida = io.BytesIO()
    with zipfile.ZipFile(salida, "w", zipfile.ZIP_DEFLATED) as z:
        for info, b in entradas:
            if info.filename == "[Content_Types].xml":
                b = _con_tipos(b)
            z.writestr(info, b)
        z.writestr(RUTA_DATOS, json.dumps(sobre, ensure_ascii=False, indent=1))
        if portada:
            z.writestr(RUTA_PORTADA, portada, compress_type=zipfile.ZIP_STORED)
    return salida.getvalue(), avisos


# ---------------------------------------------------------------------------
# Revisión y regeneración
# ---------------------------------------------------------------------------

def leer(docx: bytes) -> dict:
    """Lee y valida los datos incrustados. Lanza ErrorRevision con un mensaje para el usuario."""
    try:
        z = zipfile.ZipFile(io.BytesIO(docx))
    except zipfile.BadZipFile:
        raise ErrorRevision("El archivo no es un documento Word válido.")
    with z:
        nombres = set(z.namelist())
        if RUTA_DATOS not in nombres:
            raise ErrorRevision(
                "Este presupuesto no se puede revisar en la aplicación: se generó con una versión anterior "
                "de la skill o se ha abierto y guardado en Word. Hay que volver a generarlo en Claude.")
        try:
            sobre = json.loads(z.read(RUTA_DATOS).decode("utf-8"))
        except ValueError:
            raise ErrorRevision("Los datos del presupuesto están dañados. Hay que volver a generarlo en Claude.")
        if sobre.get("esquema") != ESQUEMA:
            raise ErrorRevision("Este presupuesto se generó con una versión de la skill que la aplicación no reconoce.")
        if hashlib.sha256(z.read("word/document.xml")).hexdigest() != sobre.get("huella_documento"):
            raise ErrorRevision(
                "El documento se ha modificado fuera de Claude y ya no coincide con sus datos. "
                "Hay que volver a generarlo en Claude.")
        faltan = [f["ruta"] for f in sobre["datos"].get("fotos", []) if f["ruta"] not in nombres]
        portada = sobre["datos"].get("foto_portada")
        if portada and portada not in nombres:
            faltan.append(portada)
        if faltan:
            raise ErrorRevision("Faltan fotos dentro del documento. Hay que volver a generarlo en Claude.")
    return sobre


def preparar_render(docx: bytes, datos: dict, carpeta: str) -> dict:
    """Extrae a `carpeta` las fotos que usan `datos` y devuelve una copia de `datos` con rutas locales.

    `datos` son los datos editados en la aplicación; las fotos se toman siempre del .docx original.
    """
    d = copy.deepcopy(datos)
    os.makedirs(carpeta, exist_ok=True)
    with zipfile.ZipFile(io.BytesIO(docx)) as z:
        nombres = set(z.namelist())

        def extraer(parte: str) -> str:
            if parte not in nombres or not (parte.startswith("word/media/") or parte == RUTA_PORTADA):
                raise ErrorRevision(f"La foto {parte} no está en el documento original.")
            destino = os.path.join(carpeta, parte.replace("/", "_"))
            with open(destino, "wb") as fh:
                fh.write(z.read(parte))
            return destino

        for f in d.get("fotos", []):
            f["ruta"] = extraer(f["ruta"])
        if d.get("foto_portada"):
            d["foto_portada"] = extraer(d["foto_portada"])
    return d
