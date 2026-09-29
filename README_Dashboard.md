# Dashboard NV pendientes de despacho

Archivos a publicar en GitHub Pages: `index.html` y `dataCOM.json.gz`.

## Actualizar los datos (terminal)
```
python -m pip install openpyxl      # solo la primera vez
python actualizar_datos.py          # pide la clave; lee T:\Corte_OC_FIFO\Base\TEst_4_Corte_v4.xlsm
python actualizar_datos.py --push   # ademas hace git add / commit / push
python actualizar_datos.py "D:\otra\ruta\archivo.xlsm"
```
La clave se valida contra un hash SHA-256 guardado en el script (no aparece en texto plano).
Cada ejecucion agrega un punto al historial de tendencia.

## Publicar
1. Crear el repositorio en GitHub y subir `index.html`, `dataCOM.json.gz` y `actualizar_datos.py`.
2. Settings > Pages > Deploy from branch > `main` / root.
3. Compartir `https://<usuario>.github.io/<repo>/`.

## Ver en local
`python -m http.server 8000` y abrir http://localhost:8000 (con doble clic el navegador bloquea la lectura del .gz).

## Importante: privacidad
En un repositorio publico, `dataCOM.json.gz` (clientes, montos, vendedores) es descargable por cualquiera que tenga la URL. Use un repositorio privado (Pages privado requiere plan GitHub Team/Enterprise) o proteja el acceso.
