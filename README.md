# Generador de archivos planos SCC y OCN (versión web para GitHub Pages)

Esta versión corre 100% en el navegador — no necesita Python, ni Streamlit, ni ningún servidor. Solo son 2 archivos (`index.html` y `xlsx.full.min.js`) que puedes publicar gratis con **GitHub Pages** y tener un link real para usar la plataforma desde cualquier navegador.

## Privacidad

Todo el procesamiento (leer los Excel, cruzar proveedores y precios, generar el archivo final) ocurre dentro de tu propio navegador. Ningún archivo que subas se envía a ningún servidor — ni siquiera a GitHub. Los maestros de proveedores y precios que cargues en la pestaña "Maestros" quedan guardados solo en ese navegador/computador (usando almacenamiento local), nunca se suben al repositorio. Por eso este repositorio puede ser público sin ningún problema de privacidad.

## Cómo publicarlo con GitHub Pages (sin usar la terminal)

1. Crea un repositorio nuevo en GitHub (puede ser público).
2. En la página del repositorio, usa **"uploading an existing file"** (o "Add file" → "Upload files") y sube estos dos archivos: `index.html` y `xlsx.full.min.js`. Confirma el commit.
3. Ve a **Settings** (configuración del repositorio) → **Pages** (en el menú de la izquierda).
4. En "Build and deployment" → "Source", selecciona **"Deploy from a branch"**.
5. En "Branch", selecciona **main** y la carpeta **/ (root)**. Guarda.
6. Espera uno o dos minutos y recarga la página de Settings → Pages. Ahí va a aparecer el link donde quedó publicada, con esta forma:
   `https://<tu-usuario>.github.io/<nombre-del-repositorio>/`

Ese es el link real donde la plataforma funciona — lo puedes abrir desde cualquier computador o celular, guardarlo en favoritos, o compartirlo con tu equipo.

## Cómo usarla

Entra al link. La primera vez, ve a la pestaña **Maestros** y carga el Excel de proveedores y el de lista de precios (solo hace falta hacerlo una vez por navegador; queda guardado ahí). Luego usa el **Módulo SCC** o el **Módulo OCN** según lo que necesites: sube el archivo de solicitudes, revisa la vista previa, genera, y descarga el archivo resultante.

## Reglas de negocio ya definidas

- Solo se procesan filas con `Estado = Aprobado`; el resto se descarta.
- En SCC, el campo NOTAS se llena por defecto con el nombre del archivo subido (editable antes de generar).
- En OCN, `TERCERO COMPRADOR` es fijo (901906381 por defecto, editable en pantalla).
- Un documento OCN agrupa un centro con un solo proveedor; si un centro necesita comprarle a más de un proveedor en el mismo lote, se generan varios documentos numerados independientemente para ese centro.
- Si algún ítem/unidad de medida de una solicitud de compra no se encuentra en la lista de precios (o su proveedor no está en el maestro), esa fila no se incluye en el archivo final: queda registrada aparte en la hoja "No cruzados" para que se revise manualmente.

Para el detalle completo de cómo se construyó cada regla, ver el documento de especificación guardado en el proyecto de Claude "PLATAFORMA PARA GENERAR SCC Y OCN".

## Actualizar la plataforma más adelante

Si en el futuro quieres ajustar alguna regla, edita `index.html` (toda la lógica está comentada y organizada por módulo) y vuelve a subir el archivo actualizado a GitHub — Pages se actualiza solo en un par de minutos.
