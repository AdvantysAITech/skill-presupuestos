"""
Portada infográfica del presupuesto (página 1).

Imagen A4 a página completa: foto de la obra arriba con la marca y el título,
banda con obra / cliente / propuesta, tarjeta "Distribución por capítulos"
(anillo + leyenda), tarjeta "Importe por partida" (barras) con "La obra en
cifras", fila de totales y pie.

Es el mismo diseño que generaba la aplicación antigua (portada.svg.ts), pasado
a Pillow para no depender de un rasterizador de SVG. Las coordenadas están en
la retícula original de 1240 x 1754 (A4 a 150 ppp); se dibuja a 3x y se reduce
a la resolución de salida para que bordes y textos salgan suaves.
"""

from __future__ import annotations

import os
from typing import Optional

from PIL import Image, ImageDraw, ImageFont, ImageOps

W, H, M = 1240, 1754, 60
ALTO_BANDA = 136
ALTO_TOTALES = 106
ALTO_PIE = 60
HUECO_BANDA_TARJETAS = 30
HUECO_TARJETAS_TOTALES = 26
HUECO_TOTALES_PIE = 32
ALTO_FOTO_MAXIMO = 820
ALTO_TARJETAS_MINIMO = 480
CIFRAS_POR_FILA = 2
ALTO_CIFRA = 82
HUECO_CIFRA = 14
FILA_PARTIDA = 58
FILA_LEYENDA_HOLGADA = 104
FILA_LEYENDA_COMPACTA = 42
FILA_LEYENDA_COMPACTA_MAXIMA = 62

ESCALA_DIBUJO = 3  # se dibuja a 3x
ANCHO_SALIDA = 1860  # 225 ppp: nítida al imprimir y ligera en el Word

DIR_FUENTES = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "assets", "fonts")


def _hex(color: str, alfa: float = 1.0):
    c = color.lstrip("#")
    return (int(c[0:2], 16), int(c[2:4], 16), int(c[4:6], 16), int(round(255 * alfa)))


def a_frase(texto: str) -> str:
    """'IMPERMEABILIZACIONES Y CUBIERTAS' -> 'Impermeabilizaciones y cubiertas'."""
    limpio = " ".join((texto or "").split())
    if not limpio or limpio != limpio.upper():
        return limpio
    minus = limpio.lower()
    return minus[0].upper() + minus[1:]


def fmt_pct(v: float) -> str:
    return f"{v:.1f}".replace(".", ",")


class Lienzo:
    def __init__(self, fuentes: dict):
        self.s = ESCALA_DIBUJO
        self.img = Image.new("RGBA", (W * self.s, H * self.s), (255, 255, 255, 255))
        self.d = ImageDraw.Draw(self.img)
        self.rutas = fuentes
        self._cache: dict = {}

    # --- fuentes y medidas -------------------------------------------------
    def fuente(self, familia: str, negrita: bool, tamano: float):
        clave = (familia, negrita, round(tamano * self.s))
        if clave not in self._cache:
            ruta = self.rutas[familia]["negrita" if negrita else "regular"]
            self._cache[clave] = ImageFont.truetype(ruta, max(1, round(tamano * self.s)))
        return self._cache[clave]

    def ancho(self, texto: str, tamano: float, negrita=False, familia="texto", espaciado=0.0) -> float:
        f = self.fuente(familia, negrita, tamano)
        base = f.getlength(texto) / self.s
        return base + espaciado * len(texto)

    def truncar(self, texto: str, max_px: float, tamano: float, negrita=False, familia="texto") -> str:
        limpio = " ".join((texto or "").split())
        if self.ancho(limpio, tamano, negrita, familia) <= max_px:
            return limpio
        while limpio and self.ancho(limpio.rstrip() + "…", tamano, negrita, familia) > max_px:
            limpio = limpio[:-1]
        return limpio.rstrip(" ,.;:·-") + "…"

    def tamano_que_cabe(self, texto, max_px, tamano, minimo, negrita=False, familia="texto"):
        t = tamano
        while t > minimo and self.ancho(texto, t, negrita, familia) > max_px:
            t -= 1
        return t

    def envolver(self, texto, max_px, tamano, max_lineas, negrita=False, familia="texto"):
        palabras = " ".join((texto or "").split()).split(" ")
        lineas, actual = [], ""
        for p in palabras:
            tentativa = f"{actual} {p}" if actual else p
            es_ultima = len(lineas) == max_lineas - 1
            if not actual or es_ultima or self.ancho(tentativa, tamano, negrita, familia) <= max_px:
                actual = tentativa
            else:
                lineas.append(actual)
                actual = p
        if actual:
            lineas.append(actual)
        return [self.truncar(l, max_px, tamano, negrita, familia) for l in lineas]

    # --- primitivas --------------------------------------------------------
    def _capa(self):
        return Image.new("RGBA", self.img.size, (0, 0, 0, 0))

    def _fundir(self, capa):
        self.img.alpha_composite(capa)

    def rect(self, x, y, w, h, relleno, radio=0, alfa=1.0, borde=None, alfa_borde=1.0, grosor=0.0):
        if w <= 0 or h <= 0:
            return
        s = self.s
        caja = [round(x * s), round(y * s), round((x + w) * s) - 1, round((y + h) * s) - 1]
        translucido = alfa < 1 or (borde is not None and alfa_borde < 1)
        capa = self._capa() if translucido else None
        d = ImageDraw.Draw(capa) if capa else self.d
        kw = dict(fill=_hex(relleno, alfa))
        if borde:
            kw.update(outline=_hex(borde, alfa_borde), width=max(1, round(grosor * s)))
        if radio:
            d.rounded_rectangle(caja, radius=round(radio * s), **kw)
        else:
            d.rectangle(caja, **kw)
        if capa:
            self._fundir(capa)

    def linea(self, x1, y1, x2, y2, color, grosor=1.0, alfa=1.0):
        s = self.s
        capa = self._capa()
        ImageDraw.Draw(capa).line(
            [x1 * s, y1 * s, x2 * s, y2 * s], fill=_hex(color, alfa), width=max(1, round(grosor * s))
        )
        self._fundir(capa)

    def texto(self, x, y, contenido, tamano=16, color="#ffffff", negrita=False, anclaje="start",
              espaciado=0.0, alfa=1.0, familia="texto"):
        if not contenido:
            return
        s = self.s
        f = self.fuente(familia, negrita, tamano)
        total = self.ancho(contenido, tamano, negrita, familia, espaciado)
        if anclaje == "middle":
            x0 = x - total / 2
        elif anclaje == "end":
            x0 = x - total
        else:
            x0 = x
        capa = self._capa() if alfa < 1 else None
        d = ImageDraw.Draw(capa) if capa else self.d
        relleno = _hex(color, alfa)
        if espaciado:
            cx = x0
            for c in contenido:
                d.text((cx * s, y * s), c, font=f, fill=relleno, anchor="ls")
                cx += f.getlength(c) / s + espaciado
        else:
            d.text((x0 * s, y * s), contenido, font=f, fill=relleno, anchor="ls")
        if capa:
            self._fundir(capa)

    def texto_ajustado(self, x, y, contenido, max_px, minimo, tamano, negrita=False, color="#fff", familia="texto"):
        t = self.tamano_que_cabe(contenido, max_px, tamano, minimo, negrita, familia)
        self.texto(x, y, self.truncar(contenido, max_px, t, negrita, familia), t, color, negrita, familia=familia)

    def donut(self, cx, cy, r_ext, r_int, capitulos):
        s = self.s
        con_peso = [c for c in capitulos if c["porcentaje"] > 0]
        if not con_peso:
            return
        caja = [(cx - r_ext) * s, (cy - r_ext) * s, (cx + r_ext) * s, (cy + r_ext) * s]
        if len(con_peso) == 1:
            self.d.ellipse(caja, fill=_hex(con_peso[0]["color"]))
        else:
            acumulado = -90.0
            for c in con_peso:
                barrido = c["porcentaje"] / 100 * 360
                self.d.pieslice(caja, acumulado, acumulado + barrido, fill=_hex(c["color"]))
                acumulado += barrido
        self.d.ellipse([(cx - r_int) * s, (cy - r_int) * s, (cx + r_int) * s, (cy + r_int) * s],
                       fill=_hex(self.fondo_tarjeta))

    def foto(self, ruta: Optional[str], alto: int, oscuro: str, acento: str):
        s = self.s
        ancho_px, alto_px = W * s, alto * s
        if ruta:
            try:
                im = Image.open(ruta)
                im = ImageOps.exif_transpose(im).convert("RGB")
                im = ImageOps.fit(im, (ancho_px, alto_px), method=Image.LANCZOS, centering=(0.5, 0.5))
                self.img.paste(im, (0, 0))
            except Exception:
                ruta = None
        if not ruta:
            # Sin foto: degradado diagonal oscuro -> acento.
            a, b = _hex(oscuro), _hex(acento)
            grad = Image.linear_gradient("L").resize((ancho_px, alto_px))
            base = Image.new("RGBA", (ancho_px, alto_px), a)
            tope = Image.new("RGBA", (ancho_px, alto_px), b)
            self.img.paste(Image.composite(tope, base, grad), (0, 0))
        # Degradado de lectura: oscuro arriba (marca), claro en medio, casi opaco abajo (título).
        paradas = [(0.0, 0.78), (0.3, 0.2), (0.55, 0.28), (1.0, 0.96)]
        mascara = Image.new("L", (1, alto_px))
        for yy in range(alto_px):
            t = yy / max(1, alto_px - 1)
            for (t0, a0), (t1, a1) in zip(paradas, paradas[1:]):
                if t0 <= t <= t1:
                    a = a0 + (a1 - a0) * (t - t0) / (t1 - t0)
                    break
            mascara.putpixel((0, yy), int(round(a * 255)))
        mascara = mascara.resize((ancho_px, alto_px))
        sombra = Image.new("RGBA", (ancho_px, alto_px), _hex(oscuro))
        sombra.putalpha(mascara)
        capa = self._capa()
        capa.paste(sombra, (0, 0))
        self._fundir(capa)

    def guardar(self, ruta_salida: str, calidad=84):
        alto = round(ANCHO_SALIDA * H / W)
        final = self.img.convert("RGB").resize((ANCHO_SALIDA, alto), Image.LANCZOS)
        final.save(ruta_salida, "JPEG", quality=calidad, optimize=True, progressive=True, dpi=(225, 225))
        return ruta_salida


# ---------------------------------------------------------------------------
# Maqueta
# ---------------------------------------------------------------------------

def _alto_cifras(n):
    filas = -(-n // CIFRAS_POR_FILA)
    return 0 if filas == 0 else filas * ALTO_CIFRA + (filas - 1) * HUECO_CIFRA


def _maqueta(d):
    n = len(d["capitulos"])
    holgada = n <= 4
    alto_izq = 84 + max(276 if holgada else 232, n * (FILA_LEYENDA_HOLGADA if holgada else FILA_LEYENDA_COMPACTA)) + 36
    alto_der = (92 + len(d["partidas"]) * FILA_PARTIDA + (26 if d["partidas_restantes"] else 0)
                + (58 + _alto_cifras(len(d["cifras"])) if d["cifras"] else 0) + 26)
    alto_fijo = ALTO_BANDA + HUECO_BANDA_TARJETAS + HUECO_TARJETAS_TOTALES + ALTO_TOTALES + HUECO_TOTALES_PIE + ALTO_PIE
    alto_tarjetas = max(alto_izq, alto_der, ALTO_TARJETAS_MINIMO, H - alto_fijo - ALTO_FOTO_MAXIMO)
    alto_foto = H - alto_fijo - alto_tarjetas
    y_banda = alto_foto
    y_tarjetas = y_banda + ALTO_BANDA + HUECO_BANDA_TARJETAS
    return dict(alto_foto=alto_foto, y_banda=y_banda, y_tarjetas=y_tarjetas, alto_tarjetas=alto_tarjetas,
                y_totales=y_tarjetas + alto_tarjetas + HUECO_TARJETAS_TOTALES, y_pie=H - ALTO_PIE, holgada=holgada)


# ---------------------------------------------------------------------------
# Bloques
# ---------------------------------------------------------------------------

def _radio(e, r):
    return 0 if e["esquinas"] == "rectas" else r


def _rotulo(lz, x, y, etiqueta, color, e):
    tamano = 18 if e["marca"] == "lateral" else 15
    lz.texto(x, y, etiqueta, tamano, color, True, espaciado=2.4, familia="titulos")


def _bloque_foto(lz, d, p, e, m):
    lz.foto(d.get("foto_portada"), m["alto_foto"], p["oscuro"], p["acento"])
    ident = d["identidad"]
    lateral = e["marca"] == "lateral"
    if lateral:
        lz.rect(M, 62, 8, 124, p["acento"])
        lz.texto(M + 36, 122, ident["nombre_marca"], 60, p["texto_claro"], True, espaciado=9, familia="titulos")
        lz.texto(M + 38, 172, ident["claim"], 16, p["acento"], True, espaciado=2.8)
    else:
        esp = 12
        lz.texto(W / 2 + esp / 2, 108, ident["nombre_marca"], 50, p["texto_claro"], True, "middle", esp, familia="titulos")
        lz.linea(W / 2 - 80, 138, W / 2 + 80, 138, p["texto_claro"], 2, 0.9)
        lz.texto(W / 2 + 1.2, 178, ident["claim"], 16, p["texto_claro"], False, "middle", 2.4, 0.88)

    # Título, anclado abajo.
    tam = 58 if lateral else 46
    inter = 66 if lateral else 56
    lineas = lz.envolver(d["titulo"], W - 2 * M, tam, 2, True, "titulos")
    y_ultima = m["alto_foto"] - 42
    y_primera = y_ultima - (len(lineas) - 1) * inter
    for i, l in enumerate(lineas):
        lz.texto(M, y_primera + i * inter, l, tam, p["texto_claro"], True, familia="titulos")

    y_ante = y_primera - (72 if lateral else 62)
    lz.texto(M, y_ante, d["antetitulo"], 17, p["acento"] if lateral else p["texto_claro_tenue"], True, espaciado=5)

    # Etiquetas.
    alto_et = 40
    y_et = y_ante - (46 if lateral else 34) - alto_et
    x = M
    for etiqueta in d["etiquetas"]:
        t = lz.truncar(etiqueta, 360, 16, True)
        ancho = lz.ancho(t, 16, True) + 36
        if x + ancho > W - M:
            break
        if e["esquinas"] == "rectas":
            lz.rect(x, y_et, ancho, alto_et, p["oscuro"], 0, 0.82, p["acento"], 0.9, 1.5)
        else:
            lz.rect(x, y_et, ancho, alto_et, p["oscuro"], alto_et / 2, 0.82, p["texto_claro"], 0.28, 1)
        lz.texto(x + ancho / 2, y_et + 26, t, 16, p["texto_claro"], True, "middle")
        x += ancho + 12


def _banda(lz, d, p, e, m):
    y0 = m["y_banda"]
    anchos = [420, 370, W - 2 * M - 790]
    cliente = f"Adm.: {d['administrador']}" if d["administrador"] else d["comunidad"]
    fecha = f"{d['fecha']} · Válida hasta {d['fecha_validez']}" if d["fecha_validez"] else d["fecha"]
    columnas = [
        ("OBRA", d["comunidad"], d["localidad"]),
        ("CLIENTE", "Comunidad de Propietarios", cliente),
        ("PROPUESTA", f"Nº {d['numero']}", fecha),
    ]
    lz.rect(0, y0, W, ALTO_BANDA, p["oscuro"])
    if e["filete_acento"]:
        lz.rect(0, y0, W, 6, p["acento"])
    x0 = M
    for i, (etq, valor, detalle) in enumerate(columnas):
        x = x0 if i == 0 else x0 + 28
        util = anchos[i] - (56 if i == 1 else 28)
        if i > 0:
            lz.linea(x0, y0 + 30, x0, y0 + ALTO_BANDA - 30, p["texto_claro"], 1, 0.18)
        lz.texto(x, y0 + 44, etq, 14, p["acento"] if e["filete_acento"] else p["texto_claro_tenue"],
                 e["filete_acento"], espaciado=3)
        lz.texto_ajustado(x, y0 + 81, valor, util, 16, 21, True, p["texto_claro"])
        lz.texto_ajustado(x, y0 + 109, detalle, util, 13, 16, False, p["texto_claro_tenue"])
        x0 += anchos[i]


def _tarjeta_capitulos(lz, d, p, e, m, x0, ancho):
    y0 = m["y_tarjetas"]
    lz.rect(x0, y0, ancho, m["alto_tarjetas"], p["tarjeta"], _radio(e, 18))
    if e["filete_acento"]:
        lz.rect(x0, y0, 6, m["alto_tarjetas"], p["acento"])
    _rotulo(lz, x0 + 30, y0 + 52, "DISTRIBUCIÓN POR CAPÍTULOS", p["texto"], e)

    caps = d["capitulos"]
    n = len(caps)
    arriba, abajo = y0 + 84, y0 + m["alto_tarjetas"] - 36
    cy = (arriba + abajo) / 2
    holgada = m["holgada"]
    r_ext, r_int = (138, 92) if holgada else (116, 76)
    cx = x0 + 30 + r_ext
    lz.fondo_tarjeta = p["tarjeta"]
    lz.donut(cx, cy, r_ext, r_int, caps)

    tam_c = 26 if holgada else 21
    cabe = lz.ancho(d["pem_fmt"], tam_c, True, "titulos") <= 2 * r_int - 18
    lz.texto(cx, cy + 6, d["pem_fmt"] if cabe else d["pem_abreviado"], tam_c, p["texto"], True, "middle", familia="titulos")
    lz.texto(cx, cy + 31, "Base imponible", 14 if holgada else 12, p["texto_tenue"], False, "middle")

    lx = cx + r_ext + 32
    derecha = x0 + ancho - 26
    if holgada:
        y_ini = cy - (n * FILA_LEYENDA_HOLGADA) / 2 + 24
        for i, c in enumerate(caps):
            y = y_ini + i * FILA_LEYENDA_HOLGADA
            lz.rect(lx, y - 13, 14, 14, c["color"], _radio(e, 3))
            lz.texto(lx + 22, y, f"Cap. {c['codigo']}", 17, p["texto"], True)
            lz.texto(derecha, y, f"{fmt_pct(c['porcentaje'])} %", 14, p["texto_tenue"], anclaje="end")
            for j, l in enumerate(lz.envolver(a_frase(c["nombre"]), derecha - lx - 22, 14, 2)):
                lz.texto(lx + 22, y + 24 + j * 18, l, 14, p["texto_tenue"])
            lz.texto(lx + 22, y + 72, c["importe_fmt"], 19, p["texto"], True)
    else:
        fila = min(FILA_LEYENDA_COMPACTA_MAXIMA, max(FILA_LEYENDA_COMPACTA, (abajo - arriba) / n))
        tam = 15 if fila >= 52 else 13
        salto = 20 if fila >= 52 else 17
        y_ini = cy - (n * fila) / 2 + (fila - salto) / 2 + tam * 0.8
        for i, c in enumerate(caps):
            y = y_ini + i * fila
            lz.rect(lx, y - tam + 2, tam - 2, tam - 2, c["color"], _radio(e, 2))
            ancho_cod = lz.ancho(c["codigo"], tam, True) + 8
            lz.texto(lx + tam + 6, y, c["codigo"], tam, p["texto"], True)
            lz.texto(lx + tam + 6 + ancho_cod, y,
                     lz.truncar(a_frase(c["nombre"]), derecha - lx - tam - 6 - ancho_cod, tam), tam, p["texto_tenue"])
            lz.texto(lx + tam + 6, y + salto, c["importe_fmt"], tam, p["texto"], True)
            lz.texto(derecha, y + salto, f"{fmt_pct(c['porcentaje'])} %", tam, p["texto_tenue"], anclaje="end")


def _tarjeta_partidas(lz, d, p, e, m, x0, ancho):
    y0 = m["y_tarjetas"]
    izq, der = x0 + 30, x0 + ancho - 30
    util = der - izq
    lz.rect(x0, y0, ancho, m["alto_tarjetas"], p["tarjeta"], _radio(e, 18))
    if e["filete_acento"]:
        lz.rect(x0, y0, 6, m["alto_tarjetas"], p["acento"])
    _rotulo(lz, izq, y0 + 52, "IMPORTE POR PARTIDA", p["texto"], e)

    y = y0 + 92
    for pa in d["partidas"]:
        ancho_imp = lz.ancho(pa["importe_fmt"], 16, True)
        codigo = (pa["codigo"] or "").strip()
        ancho_cod = lz.ancho(codigo, 16, True) + 14 if codigo else 0
        if codigo:
            lz.texto(izq, y, codigo, 16, p["texto"], True)
        lz.texto(izq + ancho_cod, y, lz.truncar(pa["resumen"], util - ancho_cod - ancho_imp - 18, 15), 15, p["texto_tenue"])
        lz.texto(der, y, pa["importe_fmt"], 16, p["texto"], True, "end")
        lz.rect(izq, y + 12, util, 12, p["pista"], _radio(e, 6))
        lz.rect(izq, y + 12, max(12, util * pa["proporcion"]), 12, pa["color"], _radio(e, 6))
        y += FILA_PARTIDA

    if d["partidas_restantes"]:
        n, imp = d["partidas_restantes"]
        lz.texto(izq, y - 4, f"+ {n} {'partida más' if n == 1 else 'partidas más'} · {imp}", 14, p["texto_tenue"])
        y += 26

    if not d["cifras"]:
        return
    y_rot = y + 22
    y_cajas = y_rot + 22
    ancho_caja = (util - HUECO_CIFRA) / CIFRAS_POR_FILA
    _rotulo(lz, izq, y_rot, "LA OBRA EN CIFRAS", p["texto"], e)
    for i, c in enumerate(d["cifras"]):
        cx = izq + (i % CIFRAS_POR_FILA) * (ancho_caja + HUECO_CIFRA)
        cy = y_cajas + (i // CIFRAS_POR_FILA) * (ALTO_CIFRA + HUECO_CIFRA)
        lz.rect(cx, cy, ancho_caja, ALTO_CIFRA, p["fondo"], _radio(e, 12))
        if e["filete_acento"]:
            lz.rect(cx, cy, ancho_caja, 3, p["acento"])
        lz.texto_ajustado(cx + 20, cy + 40, c["valor"], ancho_caja - 40, 16, 26, True, p["texto"], "titulos")
        lz.texto(cx + 20, cy + 64, lz.truncar(c["etiqueta"], ancho_caja - 40, 14), 14, p["texto_tenue"])


def _totales(lz, d, p, e, m):
    y0 = m["y_totales"]
    hueco = 30
    anchos = [330, 330, W - 2 * M - 660 - 2 * hueco]
    cajas = [("BASE IMPONIBLE", d["pem_fmt"], False), (d["iva_etiqueta"].upper(), d["iva_fmt"], False),
             (d["total_etiqueta"], d["total_fmt"], True)]
    en_acento = e["total"] == "acento"
    relleno_total = p["acento"] if en_acento else p["oscuro"]
    etq_total = p["oscuro"] if en_acento else p["texto_claro_tenue"]
    val_total = p["oscuro"] if en_acento else p["texto_claro"]
    x = M
    for i, (etq, valor, dest) in enumerate(cajas):
        ancho = anchos[i]
        lz.rect(x, y0, ancho, ALTO_TOTALES, relleno_total if dest else p["tarjeta"], _radio(e, 16))
        if e["filete_acento"] and not dest:
            lz.rect(x, y0, 6, ALTO_TOTALES, p["acento"])
        lz.texto(x + 28, y0 + 40, etq, 14, etq_total if dest else p["texto_tenue"], dest and en_acento, espaciado=2.4)
        tam = 36 if dest else 29
        lz.texto(x + 28, y0 + 84, lz.truncar(valor, ancho - 56, tam, True, "titulos"), tam,
                 val_total if dest else p["texto"], True, familia="titulos")
        x += ancho + hueco


def _pie(lz, d, p, e, m):
    ident = d["identidad"]
    lz.rect(0, m["y_pie"], W, ALTO_PIE, p["oscuro"])
    if e["filete_acento"]:
        lz.rect(0, m["y_pie"], W, 4, p["acento"])
    lz.texto(M, m["y_pie"] + 37, ident["pie"], 15, p["texto_claro"], True, espaciado=1 if e["filete_acento"] else 0)
    lz.texto(W - M, m["y_pie"] + 37, ident["contacto"], 15, p["texto_claro_tenue"], anclaje="end")


def generar_portada(datos: dict, ruta_salida: str) -> str:
    """`datos` lo prepara generar.py (capítulos con color y %, partidas, cifras...)."""
    ident = datos["identidad"]
    p, e = ident["paleta"], ident["estilo"]
    fuentes = {
        "titulos": {k: os.path.join(DIR_FUENTES, v) for k, v in e["fuente_titulos"].items()},
        "texto": {k: os.path.join(DIR_FUENTES, v) for k, v in e["fuente_texto"].items()},
    }
    lz = Lienzo(fuentes)
    lz.fondo_tarjeta = p["tarjeta"]
    m = _maqueta(datos)
    _bloque_foto(lz, datos, p, e, m)
    _banda(lz, datos, p, e, m)
    ancho_izq, hueco = 540, 30
    _tarjeta_capitulos(lz, datos, p, e, m, M, ancho_izq)
    _tarjeta_partidas(lz, datos, p, e, m, M + ancho_izq + hueco, W - 2 * M - ancho_izq - hueco)
    _totales(lz, datos, p, e, m)
    _pie(lz, datos, p, e, m)
    return lz.guardar(ruta_salida)


# ---------------------------------------------------------------------------
# Portada de Vertical Projects: composición en columna
# ---------------------------------------------------------------------------
# Vertical no comparte maqueta con Scala para que no parezcan empresas
# relacionadas. En lugar de foto arriba + banda + tarjetas + fila de totales:
#   - columna grafito a la izquierda con la marca en vertical, el título y las
#     etiquetas;
#   - foto a sangre a la derecha, separada por un filete dorado;
#   - zona crema abajo con la ficha de la obra y el total a la izquierda, y la
#     distribución por capítulos en una barra apilada con su lista a la derecha;
#   - pie grafito.
# Misma paleta y tipografías que el documento de referencia (VRT-2026-0003).

COL_V = 430          # ancho de la columna grafito
ALTO_SUPERIOR_V = 1080
ALTO_PIE_V = 70


def _foto_derecha(lz, ruta, x0, ancho, alto, oscuro, acento):
    s = lz.s
    caja = (round(ancho * s), round(alto * s))
    im = None
    if ruta:
        try:
            im = ImageOps.exif_transpose(Image.open(ruta)).convert("RGB")
            im = ImageOps.fit(im, caja, method=Image.LANCZOS, centering=(0.5, 0.5))
        except Exception:
            im = None
    if im is None:
        grad = Image.linear_gradient("L").rotate(90).resize(caja)
        im = Image.composite(Image.new("RGB", caja, _hex(oscuro)[:3]), Image.new("RGB", caja, _hex(acento)[:3]), grad)
    lz.img.paste(im, (round(x0 * s), 0))
    # Sombra suave hacia la columna, para que el corte no sea plano.
    # Opaca junto al filete y transparente hacia la derecha.
    sombra = Image.linear_gradient("L").rotate(90, expand=True).resize((round(220 * s), caja[1]))
    if sombra.getpixel((0, 0)) < sombra.getpixel((sombra.width - 1, 0)):
        sombra = ImageOps.invert(sombra)
    sombra = sombra.point(lambda v: int(v * 0.5))
    capa = Image.new("RGBA", lz.img.size, (0, 0, 0, 0))
    negro = Image.new("RGBA", sombra.size, _hex(oscuro))
    negro.putalpha(sombra)
    capa.paste(negro, (round(x0 * s), 0))
    lz._fundir(capa)


def generar_portada_vertical(d: dict, ruta_salida: str) -> str:
    ident = d["identidad"]
    p, e = ident["paleta"], ident["estilo"]
    fuentes = {
        "titulos": {k: os.path.join(DIR_FUENTES, v) for k, v in e["fuente_titulos"].items()},
        "texto": {k: os.path.join(DIR_FUENTES, v) for k, v in e["fuente_texto"].items()},
    }
    lz = Lienzo(fuentes)
    oscuro, acento, crema = p["oscuro"], p["acento"], p["tarjeta"]
    claro, claro_t, texto, tenue = p["texto_claro"], p["texto_claro_tenue"], p["texto"], p["tenue"] if "tenue" in p else p["texto_tenue"]
    acento_osc = "#9a7434"

    # Fondo y foto.
    lz.rect(0, 0, W, H, oscuro)
    _foto_derecha(lz, d.get("foto_portada"), COL_V, W - COL_V, ALTO_SUPERIOR_V, oscuro, acento)
    lz.rect(COL_V - 3, 0, 6, ALTO_SUPERIOR_V, acento)

    # Marca en la columna.
    x, util = 60, COL_V - 60 - 44
    palabras = ident["nombre_marca"].split()
    y = 112
    for palabra in palabras:
        t = lz.tamano_que_cabe(palabra, util, 64, 30, True, "titulos")
        lz.texto(x, y, palabra, t, claro, True, espaciado=6, familia="titulos")
        y += 72
    lz.rect(x, y - 30, 70, 5, acento)
    y += 18
    for parte in [c.strip() for c in ident["claim"].split("·") if c.strip()]:
        lz.texto(x, y, lz.truncar(parte, util, 15, True), 15, acento, True, espaciado=2.2)
        y += 28

    # Título, anclado abajo de la columna.
    # El título entero en 4 líneas como mucho: se reduce el cuerpo hasta que cabe.
    tam_tit = 46
    while True:
        lineas = lz.envolver(d["titulo"], util, tam_tit, 4, True, "titulos")
        if not lineas[-1].endswith("…") or tam_tit <= 28:
            break
        tam_tit -= 2
    inter = round(tam_tit * 1.17)
    y_ult = ALTO_SUPERIOR_V - 70
    y_pri = y_ult - (len(lineas) - 1) * inter
    for i, l in enumerate(lineas):
        lz.texto(x, y_pri + i * inter, l, tam_tit, claro, True, familia="titulos")
    lz.texto(x, y_pri - 66, d["antetitulo"], 15, acento, True, espaciado=4)
    y_et = y_pri - 66 - 34
    for etiqueta in reversed(d["etiquetas"][:3]):
        t = lz.truncar(etiqueta.upper(), util - 20, 14, True)
        lz.texto(x + 18, y_et, t, 14, claro_t, True, espaciado=1.6)
        lz.rect(x, y_et - 10, 8, 8, acento)
        y_et -= 30

    # Zona crema.
    y0 = ALTO_SUPERIOR_V
    alto_crema = H - ALTO_PIE_V - y0
    lz.rect(0, y0, W, alto_crema, crema)
    lz.rect(0, y0, W, 6, acento)

    # Ficha de la obra (columna izquierda).
    ficha = [
        ("OBRA", d["comunidad"], d["localidad"]),
        ("CLIENTE", "Comunidad de Propietarios", f"Adm.: {d['administrador']}" if d["administrador"] else ""),
        ("PROPUESTA", f"Nº {d['numero']}", f"Emitida {d['fecha']}"),
        ("VALIDEZ", f"Hasta {d['fecha_validez']}" if d["fecha_validez"] else "", ""),
    ]
    y = y0 + 66
    for etq, valor, detalle in ficha:
        if not valor:
            continue
        lz.texto(x, y, etq, 13, acento_osc, True, espaciado=2.6)
        lz.texto_ajustado(x, y + 30, valor, util, 15, 20, True, texto)
        if detalle:
            lz.texto(x, y + 54, lz.truncar(detalle, util, 14), 14, tenue)
        y += 84 if detalle else 62

    # Total, abajo a la izquierda.
    y_tot = H - ALTO_PIE_V - 60
    lz.texto(x, y_tot - 118, f"Base imponible  ·  {d['base_fmt']}", 14, tenue)
    lz.texto(x, y_tot - 92, f"{d['iva_etiqueta']}  ·  {d['iva_fmt']}", 14, tenue)
    lz.linea(x, y_tot - 76, x + util, y_tot - 76, oscuro, 1.5)
    lz.texto(x, y_tot - 46, d["total_etiqueta"], 13, acento_osc, True, espaciado=2.4)
    tam = lz.tamano_que_cabe(d["total_fmt"], util, 46, 28, True, "titulos")
    lz.texto(x, y_tot + 6, d["total_fmt"], tam, texto, True, familia="titulos")

    # Distribución por capítulos (derecha).
    xd, wd = COL_V + 50, W - COL_V - 50 - 60
    lz.linea(COL_V, y0 + 40, COL_V, H - ALTO_PIE_V - 40, oscuro, 1, 0.15)
    lz.texto(xd, y0 + 66, "DISTRIBUCIÓN DEL PRESUPUESTO", 15, texto, True, espaciado=3, familia="titulos")
    caps = d["capitulos"]
    xb, yb = xd, y0 + 96
    for c in caps:
        ancho = wd * c["porcentaje"] / 100
        lz.rect(xb, yb, ancho, 30, c["color"])
        xb += ancho
    cifras = d["cifras"][:4] if len(caps) <= 6 else []
    limite = (y_tot - 160) if cifras else (H - ALTO_PIE_V - 50)
    y = yb + 30 + 50
    fila = 56
    max_filas = max(1, int((limite - y) // 34) + 1)
    visibles = caps[:max_filas] if len(caps) <= max_filas else caps[:max_filas - 1]
    if len(visibles) > 1:
        fila = max(34, min(56, (limite - y) / (len(visibles) - 1 + (1 if len(visibles) < len(caps) else 0))))
    for c in visibles:
        lz.rect(xd, y - 13, 12, 12, c["color"])
        cod_w = lz.ancho(c["codigo"], 16, True, "titulos") + 14
        lz.texto(xd + 24, y, c["codigo"], 16, texto, True, familia="titulos")
        lz.texto(xd + wd, y, c["importe_fmt"], 16, texto, True, "end")
        lz.texto(xd + wd - 170, y, f"{fmt_pct(c['porcentaje'])} %", 14, tenue, anclaje="end")
        lz.texto(xd + 24 + cod_w, y, lz.truncar(a_frase(c["nombre"]), wd - 24 - cod_w - 290, 15), 15, texto)
        lz.linea(xd, y + min(fila, 40) / 2 + 2, xd + wd, y + min(fila, 40) / 2 + 2, oscuro, 1, 0.12)
        y += fila
    if len(visibles) < len(caps):
        lz.texto(xd + 24, y, f"+ {len(caps) - len(visibles)} capítulos más en el desglose detallado", 14, tenue)

    # La obra en cifras: una fila de números grandes, sin tarjetas.
    if cifras:
        lz.texto(xd, y_tot - 112, "LA OBRA EN CIFRAS", 15, texto, True, espaciado=3, familia="titulos")
        ancho_c = wd / len(cifras)
        for i, c in enumerate(cifras):
            cx = xd + i * ancho_c
            if i:
                lz.rect(cx - 1, y_tot - 76, 3, 92, acento)
            xi = cx + (22 if i else 0)
            lz.texto_ajustado(xi, y_tot - 30, c["valor"], ancho_c - 34, 18, 34, True, texto, "titulos")
            for j, l in enumerate(lz.envolver(c["etiqueta"], ancho_c - 34, 13, 2)):
                lz.texto(xi, y_tot - 4 + j * 17, l, 13, tenue)

    # Pie.
    yp = H - ALTO_PIE_V
    lz.rect(0, yp, W, ALTO_PIE_V, oscuro)
    lz.texto(60, yp + 43, ident["pie"].upper(), 14, claro, True, espaciado=2)
    lz.texto(W - 60, yp + 43, ident["contacto"], 14, acento, anclaje="end")
    return lz.guardar(ruta_salida)
