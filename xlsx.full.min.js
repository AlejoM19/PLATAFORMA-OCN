# Generador de archivos planos SCC y OCN (versión web para GitHub Pages)

Esta versión corre 100% en el navegador — no necesita Python, ni Streamlit, ni ningún servidor. Solo son 2 archivos (`index.html` y `xlsx.full.min.js`) que puedes publicar gratis con **GitHub Pages** y tener un link real para usar la plataforma desde cualquier navegador.

## Privacidad y maestros compartidos con el equipo

Todo el procesamiento (leer los Excel, cruzar proveedores y precios, generar el archivo final) ocurre dentro del navegador de quien la usa. Hay dos formas de manejar los maestros de proveedores y precios que usa el módulo OCN:

- **Cada persona los carga una vez, en su propio navegador** (pestaña "Maestros"): nunca se suben a ningún servidor ni a GitHub, quedan guardados localmente. Esto es lo más privado, pero cada persona del equipo tiene que cargarlos por su cuenta la primera vez.
- **Compartidos automáticamente para todo el equipo**: si subes los archivos `data/proveedores.xlsx` y `data/lista_precios.xlsx` al repositorio, la página los carga sola al abrirse — nadie tiene que subir nada manualmente, y siempre ven la misma versión. La contra: esos dos archivos sí quedarían dentro del repositorio, así que **si usas esta opción, el repositorio debe ser privado** (contienen NIT y precios reales de proveedores). Ten en cuenta que GitHub Pages con repositorios privados normalmente requiere un plan de pago (GitHub Pro, Team o Enterprise) — revísalo en la configuración de tu cuenta antes de optar por esta ruta; si tu cuenta es gratuita, usa la primera opción.

Para activar la carga automática: crea dentro del repositorio una carpeta `data` y sube ahí tus dos archivos con esos nombres exactos (`proveedores.xlsx` y `lista_precios.xlsx`). Cuando quieras actualizarlos (nuevos proveedores, precios nuevos), simplemente reemplaza esos mismos archivos en el repositorio — todo el equipo va a ver la versión nueva la próxima vez que abran la página.

## Cómo publicarlo con GitHub Pages (sin usar la terminal)

1. Crea un repositorio nuevo en GitHub (público, o privado si vas a usar maestros compartidos — ver arriba).
2. En la página del repositorio, usa **"uploading an existing file"** (o "Add file" → "Upload files") y sube `index.html` y `xlsx.full.min.js` (y, si aplica, la carpeta `data` con los dos maestros). Confirma el commit.
3. Ve a **Settings** (configuración del repositorio) → **Pages** (en el menú de la izquierda).
4. En "Build and deployment" → "Source", selecciona **"Deploy from a branch"**.
5. En "Branch", selecciona **main** y la carpeta **/ (root)**. Guarda.
6. Espera uno o dos minutos y recarga la página de Settings → Pages. Ahí va a aparecer el link donde quedó publicada, con esta forma:
   `https://<tu-usuario>.github.io/<nombre-del-repositorio>/`

Ese es el link real donde la plataforma funciona — lo puedes abrir desde cualquier computador o celular, guardarlo en favoritos, o compartirlo con tu equipo. Es importante que todos entren siempre por ese mismo link (no abriendo el archivo local ni copias sueltas), porque cada dirección distinta guarda sus propios datos por separado.

## Cómo usarla

Entra al link. Si el repositorio trae los maestros compartidos, ya van a estar listos; si no, ve a la pestaña **Maestros** y carga el Excel de proveedores y el de lista de precios (una sola vez por navegador). Luego usa el **Módulo SCC** o el **Módulo OCN** según lo que necesites: sube el archivo de solicitudes, revisa la vista previa, genera, y descarga el archivo resultante.

## Reglas de negocio ya definidas

- Solo se procesan filas con `Estado = Aprobado`; el resto se descarta.
- En SCC, el campo NOTAS se llena por defecto con el nombre del archivo subido (editable antes de generar).
- En OCN, `TERCERO COMPRADOR` es fijo (901906381 por defecto, editable en pantalla).
- Un documento OCN agrupa un centro con un solo proveedor; si un centro necesita comprarle a más de un proveedor en el mismo lote, se generan varios documentos numerados independientemente para ese centro.
- Si algún ítem/unidad de medida de una solicitud de compra no se encuentra en la lista de precios (o su proveedor no está en el maestro), esa fila no se incluye en el archivo final: queda registrada aparte en la hoja "No cruzados" para que se revise manualmente.

Para el detalle completo de cómo se construyó cada regla, ver el documento de especificación guardado en el proyecto de Claude "PLATAFORMA PARA GENERAR SCC Y OCN".

## Actualizar la plataforma más adelante

Si en el futuro quieres ajustar alguna regla, edita `index.html` (toda la lógica está comentada y organizada por módulo) y vuelve a subir el archivo actualizado a GitHub — Pages se actualiza solo en un par de minutos.
