# Importador de PO de Trader Joe's

Proyecto de Google Apps Script independiente para la pestaña **Datos** de la hoja de producción. Lee correos de Gmail con la etiqueta `[TRJO]` y escribe una fila por producto, sin borrar datos existentes.

## Instalación

1. Haz una copia de seguridad de la hoja. En **Archivo > Configuración**, establece la zona horaria `America/Los_Angeles`.
2. Desde esa hoja abre **Extensiones > Apps Script** y crea un proyecto independiente. Copia `Code.gs` en el editor. Activa «Mostrar el archivo de manifiesto appsscript.json» en configuración y copia el contenido de `appsscript.json`.
3. Guarda, vuelve a cargar la hoja y autoriza los permisos solicitados al usar **Trader Joe’s > Importar PO de Gmail**. La cuenta que autoriza debe tener acceso a la etiqueta de Gmail y a la hoja.
4. Introduce las fechas inicial y final de **ARRIVE** en formato `AAAA-MM-DD`; ambas son inclusivas. El proyecto revisa únicamente los correos recibidos durante los últimos 10 días calendario, incluyendo hoy, según la zona horaria de la hoja.

Por ejemplo, si hoy es 23 de noviembre, se revisan correos recibidos del 14 al 23 de noviembre. Solo se importan los PO cuya fecha ARRIVE esté en el rango solicitado, incluso si esa fecha es posterior al correo. Un PO recibido antes de esos 10 días no se encontrará aunque su ARRIVE coincida; el resultado muestra el intervalo de correos revisado.

Apps Script requiere permiso completo de Gmail para `GmailApp`; este proyecto solo lee correos y no los modifica.

La primera fila de `Datos` debe contener, en este orden: `PROD_DATE`, `ARRIVE_DATE`, `LOCATION`, `PO_NUMBER`, `ORDER_TYPE`, `PRODUCT`, `QTY`, `UOM`, `NOTES`. Se agregan filas al final; no se modifican las demás pestañas ni los formatos de las filas existentes.

## Formato de las filas

Las nuevas filas copian la tipografía y altura de las filas 2 y 3 de `Datos`. Además, se aplican los colores alternos blanco y gris claro, bordes y alineaciones de la plantilla; las filas `Demo` se resaltan en amarillo y negrita. Las filas 2 y 3 deben ser pedidos `Regular` con la tipografía y altura deseadas; si no lo son, la importación avisa que no pudo aplicar el formato.

Para dar formato a filas ya importadas, actualiza `Code.gs` en Apps Script, guarda y recarga la hoja. En **Trader Joe’s > Dar formato a filas existentes**, indica la primera fila que quieres modificar (por ejemplo, `17`). Se aplica hasta la última fila con datos, sin cambiar valores, fórmulas, validaciones ni formatos numéricos. Se reemplaza el formato visual de esas filas en las columnas A:I y su altura; haz una copia de seguridad si tienen formato personalizado. Si cambias manualmente el tipo de pedido, vuelve a aplicar el formato para actualizar el resaltado.

## Reglas y revisión

- Un PO con menos de 8 productos es `Demo`; con 8 o más es `Regular`. Se cuenta cada SKU del correo, no las cajas. Un SKU repetido bloquea la importación para evitar una clasificación incorrecta.
- `PROD_DATE` es el día calendario anterior a `ARRIVE`; `Case QTY` se registra en `CASES` y `Each QTY` en `UNITS`. El número de PO se guarda como texto para conservar los ceros iniciales.
- Los nombres conocidos del ejemplo se convierten al formato de la hoja; los demás se conservan tal como vienen en el correo. Agrega equivalencias exactas en `TJ_PRODUCTS` si necesitas agrupar otros nombres. Revisa que coincidan con los nombres usados en los resúmenes de la hoja.
- Si el PO ya existe, no se añaden ni sobrescriben filas, aunque se hayan perdido ceros iniciales en la hoja. Los correos idénticos del mismo PO se procesan una sola vez. El resultado avisa los PO cuyas fechas de llegada, ubicaciones, tipos, productos, cantidades o unidades difieren. `PROD_DATE` y `NOTES` no participan en la comparación, para permitir su edición manual. Corrige cambios o cancelaciones manualmente.
- Si dos correos del mismo PO tienen datos diferentes, o un correo no se puede interpretar sin ambigüedad, se detiene la importación antes de escribir. Revisa el correo y vuelve a intentar. Solo se reconocen ubicaciones PALMDALE, FONTANA, STOCKTON, SACRAMENTO, LACEY, ARIZONA, TEXAS y COLORADO; agrega equivalencias verificadas al código si aparece otra.
- Gmail puede incluir varios mensajes en una conversación: solo se procesan los mensajes recibidos en los últimos 10 días que contienen `PURCHASE AGREEMENT` y cuyo ARRIVE coincide con el intervalo solicitado. Si no se puede determinar la fecha ARRIVE, la importación se detiene sin escribir. La regla presupone que cada mensaje contiene un PO completo. Revisa manualmente pedidos cancelados, modificaciones y correos reenviados con varios PO.

Prueba primero con una copia de la hoja y un intervalo pequeño, verificando los resultados y los totales antes de usarlo en producción.
