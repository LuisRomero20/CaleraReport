# Dependencia local de Excel

`xlsx.bundle.js` corresponde a `xlsx-js-style` 1.2.0 y se distribuye bajo su
licencia MIT. Se conserva localmente para que CALERA pueda abrirse sin conexión
y para impedir que los expedientes dependan de una carga desde un CDN.

Fuente: https://www.npmjs.com/package/xlsx-js-style

`jszip.min.js` corresponde a JSZip 3.10.1, también bajo licencia MIT. Se usa
únicamente al descargar libros desde el navegador para insertar de forma
compatible la inmovilización de la fila superior en el XML de las hojas.

Fuente: https://www.npmjs.com/package/jszip

`pdf.min.mjs` y `pdf.worker.min.mjs` corresponden a PDF.js (`pdfjs-dist`). Se
distribuyen bajo la licencia Apache 2.0 incluida en `pdfjs.LICENSE`. CALERA los
usa para extraer localmente el texto de cada PDF, incluidos sus flujos
comprimidos, sin enviar expedientes a servicios externos.

Fuente: https://www.npmjs.com/package/pdfjs-dist
