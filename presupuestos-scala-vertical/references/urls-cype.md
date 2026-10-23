# URLs útiles de CYPE

## Generador de Precios

URL base: `https://generadordeprecios.info/`

Para consultar un precio específico, navegar:
- Edificación → Rehabilitación → [tipo de obra]

## Capítulos directos más usados

- **Demolición fachadas**: `https://generadordeprecios.info/espacios_urbanos/calculaprecio.asp?Valor=2|0_0|2|0|2|RD|`
- **Fachadas exteriores**: `https://generadordeprecios.info/espacios_urbanos/calculaprecio.asp?Valor=2|0_0|2|0|2|FAC|`
- **Pinturas**: `https://generadordeprecios.info/espacios_urbanos/calculaprecio.asp?Valor=2|0_0|2|0|2|RPP|`
- **Cubiertas**: `https://generadordeprecios.info/espacios_urbanos/calculaprecio.asp?Valor=2|0_0|2|0|2|QQ|`

## Código de partida explicado

Los códigos CYPE siguen el formato `XYZNNN`:
- `XYZ` = capítulo (3 letras)
- `NNN` = número correlativo dentro del capítulo

Ejemplo: `RFP010` = **R**evestimientos, **F**achada, **P**intura, partida 010 (pintura plástica exterior dos manos).

## Procedimiento para añadir una partida nueva al sistema

1. Buscar la partida en generadordeprecios.info y copiar:
   - Código (ej. RFP010)
   - Descripción larga (incluir todos los detalles técnicos)
   - Unidad (m², m³, ml, ud, p.a., etc.)
   - Precio unitario (€/Ud)
2. Calcular el importe = medición × precio
3. Añadir al JSON del presupuesto en el capítulo correspondiente.

## Web alternativa

- **Cype CAD generador online**: `https://generador.cype.es/`
- **Base de datos ITeC** (Cataluña): `https://itec.cat/serveis/itecbim/`
