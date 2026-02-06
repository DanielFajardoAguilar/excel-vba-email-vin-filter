# Excel/VBA – Email + VIN Filter (Demo)

Macro en Excel + VBA que agrupa VINs por correo, elimina duplicados y exporta únicamente los correos que tienen exactamente N VINs únicos, en formato de tabla:
CORREO | VIN_1 | ... | VIN_N

Confidencialidad:
El desarrollo original se utilizó con información de una empresa (correos reales, identificadores/series y bases internas). Por lo mismo, no se publica el archivo original ni datos reales. Este repositorio contiene una versión demo con datos ficticios y lógica equivalente para fines de portafolio.

## Para qué sirve
Cuando una base contiene múltiples filas por correo (cliente) y VIN, esta macro permite:
- Consolidar VINs por correo
- Evitar VINs duplicados por correo
- Filtrar casos donde el correo tiene exactamente N VINs únicos
- Generar una salida lista para reporte/seguimiento

## Cómo funciona
1. Seleccionas la celda donde inicia la base (columna Correo).
   - Se asume que el VIN está en la columna inmediata a la derecha.
2. Ingresas el valor N (VINs exactos por correo a filtrar).
3. Seleccionas la celda donde iniciará la salida.
4. La macro agrupa por correo, elimina duplicados y exporta solo los correos con N VINs únicos.

## Estructura del repositorio
- src/ contiene el módulo VBA exportado (.bas)
- data/ contiene una base ficticia
- output/ incluye capturas del input/output 

## Uso
1. Abrir el archivo de demo en demo/
2. Importar src/correosXnumVINs.bas en el Editor VBA (Alt + F11)
3. Ejecutar la macro CorreosPorNumeroVINs
4. Seguir los cuadros de diálogo (rango base, N, celda de salida)

## Salida
Tabla con encabezados:
- CORREO
- VIN_1 ... VIN_N

## Supuestos y notas
- Se filtra por VINs únicos: si un VIN se repite para el mismo correo, cuenta una sola vez.
- La columna VIN debe estar inmediatamente a la derecha de la columna Correo.
