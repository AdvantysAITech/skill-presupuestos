"""
Prueba de regeneración (B2.5): lee los datos incrustados en un presupuesto, lo vuelve a
generar con esos mismos datos y comprueba que el resultado coincide con el original.

Es exactamente lo que hará el servicio de render cuando Miguel valide sin cambiar nada.

Uso (desde la raíz del repo):
    python herramientas/probar_regeneracion.py ruta/al/presupuesto.docx

Deja el .docx regenerado junto al original, con el sufijo "-regenerado".
"""
import json
import os
import subprocess
import sys
import tempfile

RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SCRIPTS = os.path.join(RAIZ, "presupuestos-scala-vertical", "scripts")
sys.path.insert(0, SCRIPTS)

import generar  # noqa: E402
from incrustar import ErrorRevision, leer, preparar_render  # noqa: E402


def totales(datos):
    c = generar.calcular(datos)
    return {"PEM": c["pem"], "GG y BI": c["margen_importe"], "IVA": c["iva"], "TOTAL": c["total"],
            "capitulos": [(x["codigo"], x["subtotal"]) for x in c["capitulos"]], "pendientes": c["pendientes"]}


def sin_rutas(datos):
    d = json.loads(json.dumps(datos))
    for f in d.get("fotos", []):
        f.pop("ruta", None)
    d.pop("foto_portada", None)
    return d


def main():
    if len(sys.argv) != 2:
        sys.exit(__doc__)
    origen = os.path.abspath(sys.argv[1])
    with open(origen, "rb") as fh:
        docx = fh.read()
    try:
        sobre = leer(docx)
    except ErrorRevision as e:
        sys.exit(f"NO REVISABLE: {e}")
    print(f"Datos leídos: esquema {sobre['esquema']}, skill {sobre['version_skill']}, "
          f"{len(sobre['datos']['fotos'])} fotos, portada {'sí' if sobre['datos'].get('foto_portada') else 'no'}")

    destino = os.path.splitext(origen)[0] + "-regenerado.docx"
    with tempfile.TemporaryDirectory() as tmp:
        datos = preparar_render(docx, sobre["datos"], os.path.join(tmp, "fotos"))
        ruta_json = os.path.join(tmp, "datos.json")
        with open(ruta_json, "w", encoding="utf-8") as fh:
            json.dump(datos, fh, ensure_ascii=False)
        r = subprocess.run([sys.executable, os.path.join(SCRIPTS, "generar.py"), ruta_json, destino],
                           capture_output=True, text=True)
        print(r.stdout, end="")
        if r.returncode != 0:
            sys.exit(f"ERROR al regenerar:\n{r.stderr}")

    with open(destino, "rb") as fh:
        sobre2 = leer(fh.read())

    errores = 0
    t1, t2 = totales(sobre["datos"]), totales(sobre2["datos"])
    for k in t1:
        ok = t1[k] == t2[k]
        errores += not ok
        print(f"  {'OK   ' if ok else 'ERROR'} {k}: {t1[k]}" + ("" if ok else f"  ->  {t2[k]}"))
    ok = sin_rutas(sobre["datos"]) == sin_rutas(sobre2["datos"])
    errores += not ok
    print(f"  {'OK   ' if ok else 'ERROR'} datos incrustados idénticos (salvo rutas internas de las fotos)")
    ok = len(sobre2["datos"]["fotos"]) == len(sobre["datos"]["fotos"])
    errores += not ok
    print(f"  {'OK   ' if ok else 'ERROR'} número de fotos: {len(sobre2['datos']['fotos'])}")
    print(f"\nRegenerado en: {destino}")
    print("RESULTADO: " + ("CORRECTO" if not errores else f"{errores} DIFERENCIA(S)"))
    sys.exit(1 if errores else 0)


if __name__ == "__main__":
    main()
