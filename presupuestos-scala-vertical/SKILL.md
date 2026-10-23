---
name: presupuestos-scala-vertical
description: "Genera presupuestos de obra en Word para Scala Valencia o Vertical Projects con su diseño corporativo, portada con foto e infografía, desglose por capítulos y anexo fotográfico."
---

# Presupuestos Scala Valencia y Vertical Projects

Genera el presupuesto en Word (.docx) con el diseño corporativo de cada empresa:

1. **Portada** a página completa: la primera foto de la obra con la marca y el título, banda con obra / cliente / propuesta, gráfico de distribución por capítulos, importe por partida, "La obra en cifras" y totales.
2. **Presupuesto y mediciones**: objeto y alcance, diagnóstico del edificio, alertas técnicas, desglose económico por capítulos con totales, condiciones y validez, firmas.
3. **Desglose detallado**: una tabla por capítulo (código, resumen, ud, cantidad, precio, importe, total).
4. **Garantías y certificaciones, condiciones generales, próximo paso y firmas.**
5. **Anexo fotográfico** con índice por tipo de trabajo y fotos numeradas.

No lleva cronograma.

Las dos empresas llevan el mismo contenido pero **maquetas distintas**, para que un cliente no pueda relacionarlas. Scala: portada con la foto arriba y tarjetas, tablas en rejilla. Vertical: portada en columna grafito con la foto a la derecha y zona crema con la ficha, la distribución y el total; rótulos con número grande en dorado, tablas abiertas con filetes, garantías en tarjetas y anexo a dos columnas. No intentes que una se parezca a la otra.

El script calcula todos los importes (cantidad × precio, redondeo a céntimos), los subtotales, el IVA y el total. Tú solo pasas cantidades y precios unitarios.

## Antes de generar: preguntas en un solo mensaje

Lee el archivo de instrucciones del encargo (lo genera la aplicación de los comerciales). Después reúne en **un solo mensaje** todo lo que falte. No preguntes lo que ya esté en el archivo de instrucciones o en los adjuntos.

Pregunta siempre:
- **Número de propuesta.** No lo calcules ni empieces en 001: esta skill no guarda histórico entre conversaciones. Úsalo tal cual te lo den.
- **IVA**: 10 % (rehabilitación de vivienda), 21 % u otro, o sin IVA. Nunca lo asumas.
- **Margen**: porcentaje sobre los precios de origen (0 si los precios ya son finales) y si va sobre cada partida o sobre el total.
- **Título del presupuesto** (el tema que va en la portada), proponiendo dos o tres opciones según el alcance.
- **Plazo de ejecución** en días (rango orientativo, por ejemplo "8-12 días").
- **Alertas técnicas**: si las que hayas detectado van en el documento o son solo internas.

Pregunta solo si aplica:
- **Nº de propietarios**, solo si quieren mostrar el coste por propietario. Nunca lo supongas.
- **Capítulo gratis**, solo si el comercial lo menciona.
- Las mediciones o datos imprescindibles que falten, según `references/checklists/` (lee `checklist-comun.md` y solo la del tipo de trabajo: fachada, amianto, cubierta o verticales).

Las fotos adjuntas en la conversación van siempre al anexo, sin volver a preguntar. La primera foto es la portada (si el comercial indica otra, usa `foto_portada`).

## Precios

1. Busca la partida en `references/precios-referencia.md`, `references/precios-referencia-ampliado.md` y `references/mapa-partidas-cype.md` (con `references/urls-cype.md`).
2. Si no está, búscala en el generador de precios de CYPE (generadordeprecios.info) con la búsqueda web.
3. Si tampoco la encuentras, no inventes el precio: pon la partida con `"pendiente": true`. Sale como "Pendiente de precio", no suma en el total y el documento lo explica. Avísalo en el chat.

Usa precios realistas con dos decimales al estilo CYPE (22,47 €/m², no 22,00 €/m²) y mediciones reales. Nunca inventes mediciones.

## Generar

```bash
pip install python-docx pillow --break-system-packages   # solo si faltan
python3 scripts/generar.py datos.json "/mnt/user-data/outputs/<nombre del archivo>.docx"
```

El nombre del archivo es el que pide el encargo. Si no se pasa ruta, se guarda en `/mnt/user-data/outputs/` como "Presupuesto <número> - <comunidad>.docx".

El script imprime el PEM, el IVA, el total, el subtotal de cada capítulo, las partidas pendientes y cualquier aviso. **Revisa esa salida** antes de entregar: que los capítulos y el total cuadren con lo que esperas y que no haya avisos sin resolver.

El documento se mantiene por debajo de ~3,6 MB para poder subirlo a la aplicación: el script comprime las fotos automáticamente. Si avisa de que se pasa, quita fotos y vuelve a generarlo.

El script incrusta dentro del .docx los datos con los que se ha generado y la foto de portada (`scripts/incrustar.py`). Con ellos, dirección revisa y corrige el presupuesto en la aplicación sin volver a Claude. No hay que hacer nada para que se incrusten, pero sí respetar la regla de cambios de abajo.

## Estructura del JSON

Ejemplos completos en `examples/ejemplo_scala.json` y `examples/ejemplo_vertical.json`.

```json
{
  "empresa": "scala",
  "numero": "SV-2026-021",
  "fecha": "21/10/2026",
  "titulo": "Rehabilitación de cubiertas y reparaciones puntuales",
  "cliente": {
    "comunidad": "C.P. Benipeixcar 4",
    "localidad": "Valencia",
    "provincia": "Valencia",
    "administracion": "Fincas Iranzo",
    "administracion_localidad": "Valencia"
  },
  "tipos_trabajo": ["Cubiertas", "Reparaciones puntuales"],
  "objeto_intro": "",
  "objeto": [
    {"titulo": "Cubierta plana", "texto": "Qué se hace y cómo, en dos o tres frases."}
  ],
  "diagnostico": {
    "intro": "Tras la inspección visual realizada en Valencia, hemos identificado las siguientes patologías:",
    "patologias": [
      {"zona": "Cubierta plana (198 m²)", "descripcion": "...", "gravedad": "Urgente"}
    ],
    "conclusion": ""
  },
  "alertas": {"incluir_en_doc": true, "items": [{"titulo": "Estado del soporte no visible", "descripcion": "..."}]},
  "capitulos": [
    {"codigo": "01", "titulo": "Trabajos previos y demoliciones", "partidas": [
      {"codigo": "01.01", "resumen": "Limpieza del suelo de la cubierta", "descripcion": "", "ud": "m²", "cantidad": 198, "precio": 3.07},
      {"codigo": "01.02", "resumen": "Revisión de sumideros", "ud": "ud", "cantidad": 2, "pendiente": true}
    ]}
  ],
  "iva_porcentaje": 10,
  "margen": {"porcentaje": 0, "base": "partida"},
  "plazo": "8-12 días",
  "num_vecinos": null,
  "meses_financiacion": null,
  "capitulo_gratis": {"activo": false, "capitulo": "01", "mostrar_valor": true},
  "obras_referencia": [],
  "condiciones_adicionales": [],
  "lugar_firma": "Valencia",
  "foto_portada": null,
  "fotos": [
    {"ruta": "/mnt/user-data/uploads/IMG_0001.jpg", "grupo": "Cubiertas", "pie": ""}
  ]
}
```

Notas de los campos:
- `empresa`: `"scala"` o `"vertical"`. Decide logotipo, colores, tipografías, datos de empresa, garantías y condiciones (están en `assets/marcas.json`).
- `fecha`: fecha de emisión en dd/mm/aaaa (si falta, hoy). La validez se calcula sola: 30 días Scala, 15 días Vertical (`validez_dias` para cambiarla).
- `cliente`: solo la comunidad, su localidad y provincia y el **nombre del despacho** de la administración. Nunca el CIF de la comunidad ni datos personales del administrador (nombre de la persona, teléfono, email) ni de los vecinos.
- `tipos_trabajo`: hasta 3 etiquetas cortas para la portada.
- `objeto`: un bloque por zona o tipo de trabajo, en lenguaje claro para un administrador de fincas.
- `diagnostico`: opcional. `gravedad` es `Urgente`, `Moderada` o `Baja`. Distingue lo observado de lo supuesto.
- `alertas`: solo salen en el documento con `incluir_en_doc: true`.
- `capitulos[].partidas[]`: `cantidad` y `precio` (precio unitario sin IVA). `descripcion` es la descripción técnica larga (opcional). `pendiente: true` para las partidas sin precio.
- `margen.base`: `"partida"` sube cada precio unitario; `"total"` añade la línea "Gastos generales y beneficio industrial (X %)" antes del IVA.
- `iva_porcentaje`: 0 para presupuesto sin IVA.
- `num_vecinos` y `meses_financiacion`: opcionales; muestran el coste por propietario y la cuota orientativa.
- `capitulo_gratis`: oferta condicionada a aceptar todos los capítulos. No se descuenta del total.
- `obras_referencia`: obras reales que la dirección quiera citar en garantías. Nunca las inventes.
- `fotos`: rutas a las fotos de la conversación (normalmente en `/mnt/user-data/uploads/`). `grupo` agrupa el anexo por tipo de trabajo y alimenta el índice; `pie` es opcional. Admite JPG, PNG y HEIC (este último necesita `pip install pillow-heif`).

## Reglas fijas

- **Cualquier cambio** (precios, textos, partidas, IVA, fotos) se hace modificando el JSON y volviendo a ejecutar `generar.py`. Nunca edites el .docx generado, ni con python-docx ni de ninguna otra forma: un .docx retocado ya no coincide con sus datos incrustados y la aplicación no deja revisarlo.
- Un documento es de una sola empresa: nunca mezcles datos, colores ni textos de la otra. No menciones Scala Valencia en documentos de Vertical Projects.
- Vertical Projects no presupuesta retirada de amianto: si el encargo lo incluye, avísalo en el chat para derivarlo a Scala Valencia.
- Retirada de amianto (Scala): incluye capítulo de gestión de residuos; las condiciones ya citan RERA y el RD 396/2006.
- Trabajos verticales: incluye partida de seguridad y salud.
- Sin subvenciones ni ayudas: las empresas no las gestionan.
- Una sola opción de presupuesto, con un único importe.
- La forma de pago, la validez, la normativa y la firma de Miguel Baza Velasco, Director Técnico, ya vienen en la plantilla. No las cambies.
- Descuentos solo con autorización de Miguel Baza.
- No menciones en el documento que ha intervenido una inteligencia artificial.

## Al entregar

Entrega el .docx y recuerda al comercial que lo suba a la aplicación tal cual: puede abrirlo en Word para verlo, pero si lo guarda, la aplicación ya no podrá revisarlo. Escribe también en el chat un resumen breve: número, cliente, total con el IVA aplicado, margen y base, partidas pendientes de precio, qué precios vienen de CYPE y cuáles son estimados, alertas (y si van en el documento) y cualquier aviso del script. Ese resumen no va dentro del documento.

## Personalizar

- Textos de garantías, condiciones, próximo paso y datos de empresa: `assets/marcas.json`.
- Cabecera, pie y estilos de cada empresa: `templates/scala.docx` y `templates/vertical.docx` (se pueden abrir en Word; no toques los textos entre llaves `{{...}}`).
- Portada: `scripts/portada.py` (`generar_portada` para Scala, `generar_portada_vertical` para Vertical).
- Maqueta del cuerpo de Vertical: `scripts/vertical.py`. La de Scala está en `scripts/generar.py`.
