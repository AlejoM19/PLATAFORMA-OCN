<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Generador de archivos planos SCC y OCN</title>
<script src="xlsx.full.min.js"></script>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&display=swap" rel="stylesheet">
<style>
  :root {
    --azul-900: #0f2a44;
    --azul-700: #1f4e78;
    --azul-600: #2a628f;
    --azul-500: #3b7ab0;
    --azul-100: #eaf2fa;
    --azul-50: #f4f8fc;
    --verde: #15803d;
    --verde-bg: #eafaf0;
    --verde-borde: #a6f4c5;
    --rojo: #b42318;
    --rojo-bg: #fef3f2;
    --rojo-borde: #fda29b;
    --amarillo: #92650a;
    --amarillo-bg: #fffaeb;
    --amarillo-borde: #fec84b;
    --gris: #66738a;
    --gris-osc: #1d2939;
    --borde: #e1e6ee;
    --bg: #eef2f7;
    --radio: 14px;
    --sombra: 0 1px 2px rgba(16,24,44,0.04), 0 6px 20px -8px rgba(16,24,44,0.12);
    --sombra-sm: 0 1px 2px rgba(16,24,44,0.06);
  }
  * { box-sizing: border-box; }
  body {
    margin: 0;
    font-family: 'Plus Jakarta Sans', -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Arial, sans-serif;
    background: var(--bg);
    color: var(--gris-osc);
    -webkit-font-smoothing: antialiased;
  }
  .layout { display: flex; min-height: 100vh; }

  nav.sidebar {
    width: 240px;
    background: linear-gradient(180deg, var(--azul-900) 0%, var(--azul-700) 100%);
    color: white;
    padding: 26px 16px;
    flex-shrink: 0;
    display: flex;
    flex-direction: column;
    gap: 4px;
  }
  nav.sidebar .brand {
    display: flex;
    align-items: center;
    gap: 10px;
    padding: 0 8px 22px 8px;
    margin-bottom: 10px;
    border-bottom: 1px solid rgba(255,255,255,0.14);
  }
  nav.sidebar .brand .icono {
    font-size: 22px;
    width: 38px;
    height: 38px;
    display: flex;
    align-items: center;
    justify-content: center;
    background: rgba(255,255,255,0.12);
    border-radius: 10px;
  }
  nav.sidebar .brand .texto {
    font-size: 14.5px;
    font-weight: 700;
    line-height: 1.25;
    letter-spacing: 0.1px;
  }
  nav.sidebar button {
    display: flex;
    align-items: center;
    gap: 10px;
    width: 100%;
    text-align: left;
    padding: 11px 12px;
    background: none;
    border: none;
    border-radius: 10px;
    color: rgba(255,255,255,0.82);
    font-family: inherit;
    font-size: 14px;
    font-weight: 500;
    cursor: pointer;
    transition: background .15s ease, color .15s ease;
  }
  nav.sidebar button .ico { font-size: 16px; }
  nav.sidebar button.active {
    background: rgba(255,255,255,0.16);
    color: white;
    font-weight: 700;
    box-shadow: var(--sombra-sm);
  }
  nav.sidebar button:hover { background: rgba(255,255,255,0.1); color: white; }

  main { flex: 1; padding: 40px 44px; max-width: 940px; }
  .page { display: none; animation: aparecer .18s ease; }
  .page.active { display: block; }
  @keyframes aparecer { from { opacity: 0; transform: translateY(4px); } to { opacity: 1; transform: none; } }

  h2 {
    color: var(--azul-900);
    margin: 0 0 20px 0;
    font-size: 22px;
    font-weight: 800;
    letter-spacing: -0.2px;
  }
  h3 { color: var(--azul-900); font-weight: 700; }
  h4 { color: var(--azul-900); font-weight: 700; }

  .card {
    background: white;
    border: 1px solid var(--borde);
    border-radius: var(--radio);
    padding: 26px 28px;
    margin-bottom: 22px;
    box-shadow: var(--sombra);
  }
  .card p { line-height: 1.55; }
  .card ul { line-height: 1.6; padding-left: 20px; }

  label { display: block; font-weight: 600; margin-bottom: 7px; font-size: 13.5px; color: var(--gris-osc); }

  input[type="text"] {
    width: 100%;
    padding: 10px 13px;
    border: 1.5px solid var(--borde);
    border-radius: 9px;
    font-size: 14px;
    font-family: inherit;
    margin-bottom: 16px;
    background: var(--azul-50);
    transition: border-color .15s ease, box-shadow .15s ease;
  }
  input[type="text"]:focus {
    outline: none;
    border-color: var(--azul-500);
    box-shadow: 0 0 0 3px rgba(59,122,176,0.18);
    background: white;
  }

  input[type="file"] {
    margin-bottom: 16px;
    font-family: inherit;
    font-size: 13.5px;
  }
  input[type="file"]::file-selector-button {
    font-family: inherit;
    font-weight: 600;
    font-size: 13px;
    color: var(--azul-700);
    background: var(--azul-100);
    border: none;
    border-radius: 8px;
    padding: 8px 14px;
    margin-right: 12px;
    cursor: pointer;
    transition: background .15s ease;
  }
  input[type="file"]::file-selector-button:hover { background: #dbe9f6; }

  button.primary {
    background: linear-gradient(135deg, var(--azul-600), var(--azul-900));
    color: white;
    border: none;
    padding: 11px 24px;
    border-radius: 10px;
    font-family: inherit;
    font-size: 14px;
    font-weight: 700;
    letter-spacing: 0.1px;
    cursor: pointer;
    box-shadow: 0 4px 12px -4px rgba(15,42,68,0.45);
    transition: transform .12s ease, box-shadow .12s ease, opacity .12s ease;
  }
  button.primary:hover { transform: translateY(-1px); box-shadow: 0 8px 18px -6px rgba(15,42,68,0.5); }
  button.primary:active { transform: translateY(0); }

  button.secondary {
    background: white;
    color: var(--azul-700);
    border: 1.5px solid var(--azul-500);
    padding: 8px 16px;
    border-radius: 9px;
    font-family: inherit;
    font-size: 13px;
    font-weight: 600;
    cursor: pointer;
    transition: background .15s ease;
  }
  button.secondary:hover { background: var(--azul-100); }

  .alert {
    padding: 13px 16px;
    border-radius: 10px;
    margin: 12px 0;
    font-size: 14px;
    line-height: 1.5;
    border: 1px solid;
    box-shadow: var(--sombra-sm);
  }
  .alert.error { background: var(--rojo-bg); color: var(--rojo); border-color: var(--rojo-borde); }
  .alert.success { background: var(--verde-bg); color: var(--verde); border-color: var(--verde-borde); }
  .alert.warning { background: var(--amarillo-bg); color: var(--amarillo); border-color: var(--amarillo-borde); }
  .alert.info { background: var(--azul-100); color: var(--azul-700); border-color: #b6d4ea; }

  .metrics { display: flex; gap: 14px; margin: 18px 0; flex-wrap: wrap; }
  .metric {
    flex: 1;
    min-width: 130px;
    background: var(--azul-50);
    border: 1px solid var(--azul-100);
    border-radius: 12px;
    padding: 16px;
    text-align: center;
    transition: transform .12s ease;
  }
  .metric:hover { transform: translateY(-2px); }
  .metric .valor { font-size: 25px; font-weight: 800; color: var(--azul-900); letter-spacing: -0.3px; }
  .metric .etiqueta { font-size: 12px; color: var(--gris); margin-top: 5px; font-weight: 500; }

  table.preview { border-collapse: collapse; width: 100%; font-size: 12.5px; margin-top: 10px; }
  table.preview th, table.preview td {
    padding: 8px 10px;
    text-align: left;
    white-space: nowrap;
    border-bottom: 1px solid var(--borde);
  }
  table.preview th {
    background: var(--azul-50);
    color: var(--azul-900);
    font-weight: 700;
    font-size: 11.5px;
    text-transform: uppercase;
    letter-spacing: 0.3px;
    position: sticky;
    top: 0;
  }
  table.preview tbody tr:hover { background: #f7fafc; }

  .bar-bg { display: inline-block; width: 90px; height: 7px; background: #e4e7ec; border-radius: 4px; overflow: hidden; margin-left: 8px; vertical-align: middle; }
  .bar-fill { display: block; height: 100%; background: linear-gradient(90deg, var(--azul-500), var(--azul-900)); }
  .pct-num { font-weight: 700; color: var(--azul-900); }

  .preview-wrap {
    overflow: auto;
    max-height: 260px;
    border: 1px solid var(--borde);
    border-radius: 10px;
  }
  details summary {
    cursor: pointer;
    font-size: 13px;
    font-weight: 600;
    color: var(--azul-700);
    margin: 10px 0;
    list-style: none;
  }
  details summary::-webkit-details-marker { display: none; }
  details summary::before { content: '▸ '; }
  details[open] summary::before { content: '▾ '; }

  .estado-maestro { font-size: 13.5px; color: var(--gris); margin-bottom: 12px; }
  .privacidad {
    font-size: 12.5px;
    color: var(--azul-700);
    background: var(--azul-100);
    border-radius: 10px;
    padding: 12px 16px;
    margin-top: 18px;
    line-height: 1.5;
  }
  code { background: var(--azul-50); padding: 1px 6px; border-radius: 4px; font-size: 13px; color: var(--azul-900); }
  hr { border: none; border-top: 1px solid var(--borde); }
</style>
</head>
<body>

<div class="layout">
  <nav class="sidebar">
    <div class="brand">
      <span class="icono">📄</span>
      <span class="texto">Generador<br>SCC / OCN</span>
    </div>
    <button data-page="inicio" class="active"><span class="ico">🏠</span> Inicio</button>
    <button data-page="scc"><span class="ico">🧴</span> Módulo SCC</button>
    <button data-page="ocn"><span class="ico">🛒</span> Módulo OCN</button>
    <button data-page="maestros"><span class="ico">🗂️</span> Maestros</button>
  </nav>

  <main>

    <!-- INICIO -->
    <section id="page-inicio" class="page active">
      <h2>Generador de archivos planos SCC y OCN</h2>
      <div class="card">
        <p>Esta plataforma toma el reporte de solicitudes (Excel) y genera el archivo plano listo para cargar al ERP, en dos módulos independientes:</p>
        <ul>
          <li><strong>SCC</strong>: solicitudes de consumo/traslado interno entre bodegas.</li>
          <li><strong>OCN</strong>: órdenes de compra a proveedor, cruzando automáticamente el proveedor, el precio y la sucursal desde los maestros cargados.</li>
        </ul>
        <p>Usa el menú de la izquierda para ir al módulo que necesites. En <strong>Maestros</strong> puedes cargar la lista de proveedores y precios que usa el módulo OCN.</p>
        <div class="privacidad">
          🔒 Todo el procesamiento ocurre en tu propio navegador. Ningún archivo que subas (solicitudes, proveedores, precios) se envía a ningún servidor — ni siquiera a GitHub. Los maestros que cargues en la pestaña "Maestros" quedan guardados solo en este navegador/computador.
        </div>
      </div>
    </section>

    <!-- SCC -->
    <section id="page-scc" class="page">
      <h2>🧴 Módulo SCC</h2>
      <div class="card">
        <p>Sube el reporte de solicitudes aprobadas (una fila por ítem solicitado, con columnas como C.O. docto., Bodega, Item, U.M., Cant. solicitada, C.Costo, Motivo, Fecha solic., Estado, etc.).</p>
        <label for="scc-file">Archivo de solicitudes (Excel)</label>
        <input type="file" id="scc-file" accept=".xlsx,.xls">

        <div id="scc-config" style="display:none;">
          <label for="scc-notas">Texto para el campo NOTAS del documento</label>
          <input type="text" id="scc-notas">

          <details>
            <summary>Ver vista previa del archivo cargado</summary>
            <div class="preview-wrap"><div id="scc-preview"></div></div>
          </details>

          <div style="margin-top:16px;">
            <button class="primary" id="scc-generar">Generar archivo SCC</button>
          </div>
        </div>

        <div id="scc-resultado"></div>
      </div>
    </section>

    <!-- OCN -->
    <section id="page-ocn" class="page">
      <h2>🛒 Módulo OCN</h2>
      <div class="card">
        <p>Sube el reporte de solicitudes de compra (misma estructura que el de SCC). El proveedor, el precio unitario, la sucursal y la condición de pago se cruzan automáticamente con los maestros cargados en la sección <strong>Maestros</strong>.</p>

        <div id="ocn-sin-maestros" class="alert error" style="display:none;">
          Todavía no has cargado el maestro de proveedores y/o la lista de precios en esta sesión del navegador. Ve a la sección <strong>Maestros</strong> para subirlos antes de usar este módulo.
        </div>

        <div id="ocn-form" style="display:none;">
          <label for="ocn-file">Archivo de solicitudes de compra (Excel)</label>
          <input type="file" id="ocn-file" accept=".xlsx,.xls">

          <label for="ocn-comprador">TERCERO COMPRADOR (fijo)</label>
          <input type="text" id="ocn-comprador" value="901906381">

          <div id="ocn-config" style="display:none;">
            <details>
              <summary>Ver vista previa del archivo de solicitudes</summary>
              <div class="preview-wrap"><div id="ocn-preview"></div></div>
            </details>
            <div style="margin-top:16px;">
              <button class="primary" id="ocn-generar">Generar archivo OCN</button>
            </div>
          </div>
        </div>

        <div id="ocn-resultado"></div>
      </div>
    </section>

    <!-- MAESTROS -->
    <section id="page-maestros" class="page">
      <h2>🗂️ Maestros</h2>
      <div class="card">
        <p>Estos archivos alimentan el módulo OCN. Si el repositorio de esta página incluye <code>data/proveedores.xlsx</code> y <code>data/lista_precios.xlsx</code>, se cargan solos al abrir la página — así todo el equipo ve siempre los mismos maestros. Si no, se guardan únicamente en este navegador (nunca se suben a ningún servidor) y tendrás que cargarlos aquí manualmente cada vez que cambies de computador o de navegador.</p>
        <div id="aviso-file-protocol" class="alert warning" style="display:none;">
          Estás abriendo este archivo directamente en tu computador, no desde su link publicado. En este modo el navegador no permite guardar los maestros de forma permanente — vas a poder generar archivos en esta misma sesión, pero tendrás que volver a cargarlos cada vez que abras la página de nuevo. Para que queden guardados entre visitas, publícala con GitHub Pages y úsala desde ese link.
        </div>

        <h3>Maestro de proveedores</h3>
        <p class="estado-maestro" id="estado-proveedores">No hay maestro de proveedores cargado todavía.</p>
        <label for="prov-file">Cargar / reemplazar maestro de proveedores</label>
        <input type="file" id="prov-file" accept=".xlsx,.xls">
        <div id="prov-msg"></div>

        <hr style="margin:24px 0; border:none; border-top:1px solid var(--borde);">

        <h3>Lista de precios</h3>
        <p class="estado-maestro" id="estado-precios">No hay lista de precios cargada todavía.</p>
        <label for="precios-file">Cargar / reemplazar lista de precios</label>
        <input type="file" id="precios-file" accept=".xlsx,.xls">
        <div id="precios-msg"></div>
      </div>
    </section>

  </main>
</div>

<script>
/* ==========================================================================
   Navegación entre páginas
   ========================================================================== */
document.querySelectorAll('nav.sidebar button').forEach(btn => {
  btn.addEventListener('click', () => {
    document.querySelectorAll('nav.sidebar button').forEach(b => b.classList.remove('active'));
    document.querySelectorAll('.page').forEach(p => p.classList.remove('active'));
    btn.classList.add('active');
    document.getElementById('page-' + btn.dataset.page).classList.add('active');
    if (btn.dataset.page === 'ocn') actualizarEstadoOCN();
  });
});

/* ==========================================================================
   Utilidades compartidas (equivalentes a transformations/utils.py)
   ========================================================================== */
function stripStr(v) {
  if (v === null || v === undefined) return '';
  return String(v).trim();
}

function normalizarItem(v) {
  const s = stripStr(v);
  if (/^\d+$/.test(s)) {
    const sinCeros = s.replace(/^0+/, '');
    return sinCeros === '' ? '0' : sinCeros;
  }
  return s;
}

function toNumber(v) {
  if (v === null || v === undefined) return v;
  const s = String(v).trim();
  if (s === '') return v;
  const f = parseFloat(s);
  if (isNaN(f)) return v;
  return f;
}

function pad2(n) { return String(n).padStart(2, '0'); }

function formatFecha(v) {
  if (v === null || v === undefined || v === '') return '';
  if (v instanceof Date && !isNaN(v)) {
    return `${v.getFullYear()}${pad2(v.getMonth() + 1)}${pad2(v.getDate())}`;
  }
  const s = String(v).trim();
  if (/^\d{8}$/.test(s)) return s;
  let m;
  if ((m = s.match(/^(\d{4})-(\d{2})-(\d{2})/))) return `${m[1]}${m[2]}${m[3]}`;
  if ((m = s.match(/^(\d{2})\/(\d{2})\/(\d{4})$/))) return `${m[3]}${m[2]}${m[1]}`;
  if ((m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/))) return `${m[3]}${pad2(m[2])}${pad2(m[1])}`;
  return s;
}

function nitBase(v) {
  const s = stripStr(v);
  return s.split('-')[0].trim();
}

function trimKeys(rows) {
  return rows.map(r => {
    const o = {};
    for (const k in r) o[String(k).trim()] = r[k];
    return o;
  });
}

/* ==========================================================================
   Lectura de Excel con SheetJS
   ========================================================================== */
function parsearWorkbookArrayBuffer(buffer) {
  const data = new Uint8Array(buffer);
  const wb = XLSX.read(data, { type: 'array', cellDates: true });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(ws, { defval: '' });
  return trimKeys(rows);
}

function leerExcel(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        resolve(parsearWorkbookArrayBuffer(e.target.result));
      } catch (err) {
        reject(err);
      }
    };
    reader.onerror = () => reject(reader.error);
    reader.readAsArrayBuffer(file);
  });
}

/* SIESA espera el archivo plano con todas las celdas en formato Texto, no
   Número/General — de lo contrario puede truncar ceros a la izquierda,
   redondear decimales o interpretar mal ciertos códigos. Esta función arma
   la hoja a partir de una matriz de filas y fuerza cada celda (menos el
   encabezado) a texto, tanto en el tipo de dato interno como en el formato
   de celda visible en Excel. */
function hojaComoTexto(aoa) {
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  if (!ws['!ref']) return ws;
  const range = XLSX.utils.decode_range(ws['!ref']);
  for (let R = range.s.r; R <= range.e.r; R++) {
    for (let C = range.s.c; C <= range.e.c; C++) {
      const addr = XLSX.utils.encode_cell({ r: R, c: C });
      const cell = ws[addr];
      if (!cell) continue;
      if (cell.v !== undefined && cell.v !== null) {
        cell.v = String(cell.v);
      }
      cell.t = 's';
      cell.z = '@';
      delete cell.w;
    }
  }
  return ws;
}

function validarColumnas(rows, requeridas, nombre) {
  if (!rows.length) throw new Error(`El archivo de ${nombre} está vacío.`);
  const cols = Object.keys(rows[0]);
  const faltantes = requeridas.filter(c => !cols.includes(c));
  if (faltantes.length) {
    throw new Error(`El archivo de ${nombre} no tiene las columnas esperadas. Faltan: ${faltantes.join(', ')}`);
  }
}

/* ==========================================================================
   Módulo SCC (equivalente a transformations/scc.py)
   ========================================================================== */
const REQUIRED_SCC_COLUMNS = ["C.O. docto.", "C.O. movto.", "Bodega", "Motivo", "C.Costo", "U.M.", "Cant. solicitada", "Fecha solic.", "Item", "Estado"];

function generarSCC(rowsIn, notas) {
  validarColumnas(rowsIn, REQUIRED_SCC_COLUMNS, "solicitudes");
  const totalFilas = rowsIn.length;

  let rows = rowsIn.filter(r => stripStr(r["Estado"]).toLowerCase() === 'aprobado');
  const descartadasEstado = totalFilas - rows.length;
  if (!rows.length) throw new Error("No quedó ninguna fila con Estado = 'Aprobado' en el archivo de solicitudes.");

  rows = rows.map(r => ({
    "C.O. docto.": stripStr(r["C.O. docto."]),
    "C.O. movto.": stripStr(r["C.O. movto."]),
    "Bodega": stripStr(r["Bodega"]),
    "Motivo": toNumber(r["Motivo"]),
    "C.Costo": stripStr(r["C.Costo"]),
    "U.M.": stripStr(r["U.M."]),
    "Item": normalizarItem(r["Item"]),
    "Fecha solic.": formatFecha(r["Fecha solic."]),
    "Cant. solicitada": toNumber(r["Cant. solicitada"]),
  }));

  const gruposOrden = [];
  const vistos = new Set();
  for (const r of rows) {
    if (!vistos.has(r["C.O. docto."])) { vistos.add(r["C.O. docto."]); gruposOrden.push(r["C.O. docto."]); }
  }
  const consecutivoPorGrupo = {};
  gruposOrden.forEach((co, i) => consecutivoPorGrupo[co] = i + 1);

  const docsAOA = [["COMPANIA", "CENTRO DE OPERACION", "CONSECUTIVO", "FECHA  AAAAMMDD", "NOTAS", "DCTO REFERENCIA"]];
  let documentosGenerados = 0;
  for (const centro of gruposOrden) {
    const sub = rows.filter(r => r["C.O. docto."] === centro);
    docsAOA.push([1, centro, consecutivoPorGrupo[centro], sub[0]["Fecha solic."], notas, "SCC"]);
    documentosGenerados++;
  }

  const movAOA = [["COMPANIA", "CENTRO DE OPERACION CO", "CONSECUTIVO", "NRO MOVIMIENTOS", "BODEGA", "MOTIVO 01/52", "CO DE MOVIMIENTO", "CENTRO DE COSTO", "UM ORDEN", "CANT A SOLICITAR", "FECHA AAAAMMDD", "CODIGO ITEM"]];
  let movimientosGenerados = 0;
  for (const centro of gruposOrden) {
    const sub = rows.filter(r => r["C.O. docto."] === centro);
    const consecutivo = consecutivoPorGrupo[centro];
    sub.forEach((row, i) => {
      movAOA.push([1, row["C.O. movto."], consecutivo, i + 1, row["Bodega"], row["Motivo"], row["C.O. movto."], row["C.Costo"], row["U.M."], row["Cant. solicitada"], row["Fecha solic."], row["Item"]]);
      movimientosGenerados++;
    });
  }

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, hojaComoTexto([["COMPANIA"], [1]]), "Inicial");
  XLSX.utils.book_append_sheet(wb, hojaComoTexto(docsAOA), "Documentos");
  XLSX.utils.book_append_sheet(wb, hojaComoTexto(movAOA), "Movimientos");
  XLSX.utils.book_append_sheet(wb, hojaComoTexto([["COMPANIA"], [1]]), "Final");

  return {
    wb,
    resumen: {
      total_filas_entrada: totalFilas,
      descartadas_por_estado: descartadasEstado,
      documentos_generados: documentosGenerados,
      movimientos_generados: movimientosGenerados,
    },
  };
}

/* ==========================================================================
   Módulo OCN (equivalente a transformations/ocn.py)
   ========================================================================== */
const REQUIRED_SOLICITUD_COLUMNS = ["C.O. docto.", "C.O. movto.", "Bodega", "U.M.", "Cant. solicitada", "Fecha solic.", "Item", "Estado"];
const REQUIRED_PROVEEDOR_COLUMNS = ["Nit", "Sucursal", "Condicion de pago", "Numero identificación"];
const REQUIRED_PRECIO_COLUMNS = ["numero identificacion proveedor", "Item", "U.M.", "Precio unit."];

function generarOCN(rowsSolIn, rowsProvIn, rowsPrecIn, terceroComprador) {
  validarColumnas(rowsSolIn, REQUIRED_SOLICITUD_COLUMNS, "solicitudes de compra");
  validarColumnas(rowsProvIn, REQUIRED_PROVEEDOR_COLUMNS, "maestro de proveedores");
  validarColumnas(rowsPrecIn, REQUIRED_PRECIO_COLUMNS, "lista de precios");

  const totalFilas = rowsSolIn.length;
  let rows = rowsSolIn.filter(r => stripStr(r["Estado"]).toLowerCase() === 'aprobado');
  const descartadasEstado = totalFilas - rows.length;
  if (!rows.length) throw new Error("No quedó ninguna fila con Estado = 'Aprobado' en el archivo de solicitudes.");

  const tieneFechaEntrega = Object.keys(rowsSolIn[0] || {}).includes("Fecha entrega");

  rows = rows.map(r => ({
    "C.O. docto.": stripStr(r["C.O. docto."]),
    "C.O. movto.": stripStr(r["C.O. movto."]),
    "Bodega": stripStr(r["Bodega"]),
    "U.M.": stripStr(r["U.M."]),
    "Item": normalizarItem(r["Item"]),
    "Fecha solic.": formatFecha(r["Fecha solic."]),
    "Fecha entrega": tieneFechaEntrega ? formatFecha(r["Fecha entrega"]) : "",
    "Cant. solicitada": toNumber(r["Cant. solicitada"]),
  }));

  const precioLookup = new Map();
  rowsPrecIn.forEach(r => {
    const item = normalizarItem(r["Item"]);
    const um = stripStr(r["U.M."]);
    const nb = nitBase(r["numero identificacion proveedor"]);
    precioLookup.set(item + "|" + um, [nb, toNumber(r["Precio unit."])]);
  });

  const provLookup = new Map();
  const nombreProvLookup = new Map();
  rowsProvIn.forEach(r => {
    const nb = nitBase(r["Numero identificación"]);
    provLookup.set(nb, [stripStr(r["Nit"]), stripStr(r["Sucursal"]), stripStr(r["Condicion de pago"])]);
    if (!nombreProvLookup.has(nb)) nombreProvLookup.set(nb, stripStr(r["Razón social"]));
  });

  const descItemLookup = new Map();
  rowsPrecIn.forEach(r => {
    const item = normalizarItem(r["Item"]);
    if (!descItemLookup.has(item)) descItemLookup.set(item, stripStr(r["Desc. item"]));
  });

  const filasValidas = [];
  const filasSinCruce = [];
  for (const row of rows) {
    const item = row["Item"], um = row["U.M."];
    const key = item + "|" + um;
    if (!precioLookup.has(key)) {
      filasSinCruce.push({ motivo: "Item/U.M. no encontrado en lista de precios", item, um, bodega: row["Bodega"], proveedor: "" });
      continue;
    }
    const [nb, precio] = precioLookup.get(key);
    if (!provLookup.has(nb)) {
      filasSinCruce.push({ motivo: "Proveedor no encontrado en maestro de proveedores", item, um, bodega: row["Bodega"], proveedor: nb });
      continue;
    }
    const [nitSinGuion, sucursal, condicionPago] = provLookup.get(nb);
    filasValidas.push({
      centro: row["C.O. docto."],
      coMovto: row["C.O. movto."],
      bodega: row["Bodega"],
      um, cantidad: row["Cant. solicitada"],
      fechaSolic: row["Fecha solic."],
      fechaEntrega: row["Fecha entrega"] || row["Fecha solic."],
      item, precio,
      proveedor: nitSinGuion, sucursal, cPago: condicionPago,
    });
  }

  if (!filasValidas.length) {
    throw new Error("Ninguna fila del archivo de solicitudes pudo cruzarse con la lista de precios / maestro de proveedores.");
  }

  const consecutivoPorCentro = {};
  const grupoConsecutivo = new Map();
  const ordenGrupos = [];
  for (const row of filasValidas) {
    const gkey = row.centro + "||" + row.proveedor;
    if (!grupoConsecutivo.has(gkey)) {
      consecutivoPorCentro[row.centro] = (consecutivoPorCentro[row.centro] || 0) + 1;
      grupoConsecutivo.set(gkey, consecutivoPorCentro[row.centro]);
      ordenGrupos.push(gkey);
    }
  }

  const docsAOA = [["C.O", "TIPO DCTO", "NRO CONSECUTIVO", "FECHA", "TERCERO COMPRADOR", "TERCERO PROVEEDOR", "SUCURSAL", "C PAGO"]];
  const movAOA = [["CO", "TIPO DCTO", "NRO CONSECUTIVO", "NRO REGISTRO", "BODEGA", "C.O MOVIMIENTO", "UM", "CANTIDAD", "AAAAMMDD ENTREGA", "PRECIO UNITARIO", "ITEM"]];
  let documentosGenerados = 0, movimientosGenerados = 0;

  for (const gkey of ordenGrupos) {
    const [centro, proveedor] = gkey.split("||");
    const sub = filasValidas.filter(r => r.centro === centro && r.proveedor === proveedor);
    const consecutivo = grupoConsecutivo.get(gkey);
    docsAOA.push([centro, "OCN", consecutivo, sub[0].fechaSolic, terceroComprador, proveedor, sub[0].sucursal, sub[0].cPago]);
    documentosGenerados++;
    sub.forEach((row, i) => {
      movAOA.push([centro, "OCN", consecutivo, i + 1, row.bodega, row.coMovto, row.um, row.cantidad, row.fechaEntrega, row.precio, row.item]);
      movimientosGenerados++;
    });
  }

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, hojaComoTexto(docsAOA), "Documentos");
  XLSX.utils.book_append_sheet(wb, hojaComoTexto(movAOA), "Movimientos");

  if (filasSinCruce.length) {
    const errAOA = [["Motivo", "Item", "U.M.", "Bodega", "Proveedor"]];
    filasSinCruce.forEach(e => errAOA.push([e.motivo, e.item, e.um, e.bodega, e.proveedor]));
    XLSX.utils.book_append_sheet(wb, hojaComoTexto(errAOA), "No cruzados");
  }

  const analisis = construirAnalisisOCN(filasValidas, nombreProvLookup, descItemLookup);

  return {
    wb,
    resumen: {
      total_filas_entrada: totalFilas,
      descartadas_por_estado: descartadasEstado,
      filas_sin_cruce: filasSinCruce.length,
      documentos_generados: documentosGenerados,
      movimientos_generados: movimientosGenerados,
    },
    analisis,
  };
}

function construirAnalisisOCN(filasValidas, nombreProvLookup, descItemLookup) {
  let cantidadTotal = 0, valorTotal = 0;
  const porProveedor = new Map();
  const porCentro = new Map();
  const porItem = new Map();

  for (const row of filasValidas) {
    const cantidad = Number(row.cantidad) || 0;
    const valor = cantidad * (Number(row.precio) || 0);
    cantidadTotal += cantidad;
    valorTotal += valor;

    const p = porProveedor.get(row.proveedor) || { clave: row.proveedor, cantidad: 0, valor: 0 };
    p.cantidad += cantidad; p.valor += valor;
    porProveedor.set(row.proveedor, p);

    const c = porCentro.get(row.centro) || { clave: row.centro, cantidad: 0, valor: 0 };
    c.cantidad += cantidad; c.valor += valor;
    porCentro.set(row.centro, c);

    const i = porItem.get(row.item) || { clave: row.item, cantidad: 0, valor: 0 };
    i.cantidad += cantidad; i.valor += valor;
    porItem.set(row.item, i);
  }

  const conPct = (mapa, nombreFn) => Array.from(mapa.values())
    .map(x => ({ ...x, nombre: nombreFn(x.clave), pct: valorTotal > 0 ? (x.valor / valorTotal) * 100 : 0 }))
    .sort((a, b) => b.valor - a.valor);

  return {
    cantidadTotal,
    valorTotal,
    numProveedores: porProveedor.size,
    numItems: porItem.size,
    porProveedor: conPct(porProveedor, nb => nombreProvLookup.get(nb) || ''),
    porCentro: conPct(porCentro, c => ''),
    topItems: conPct(porItem, it => descItemLookup.get(it) || '').slice(0, 10),
  };
}

/* ==========================================================================
   Helpers de interfaz
   ========================================================================== */
function renderAlert(container, tipo, mensaje) {
  container.insertAdjacentHTML('beforeend', `<div class="alert ${tipo}">${mensaje}</div>`);
}

function renderMetrics(container, items) {
  const html = `<div class="metrics">${items.map(i => `
      <div class="metric"><div class="valor">${i.valor}</div><div class="etiqueta">${i.etiqueta}</div></div>
    `).join('')}</div>`;
  container.insertAdjacentHTML('beforeend', html);
}

function renderPreview(container, rows, maxRows = 15) {
  if (!rows.length) { container.innerHTML = '<p>Sin datos.</p>'; return; }
  const cols = Object.keys(rows[0]);
  let html = '<table class="preview"><thead><tr>' + cols.map(c => `<th>${c}</th>`).join('') + '</tr></thead><tbody>';
  rows.slice(0, maxRows).forEach(r => {
    html += '<tr>' + cols.map(c => `<td>${r[c] instanceof Date ? r[c].toLocaleDateString() : (r[c] ?? '')}</td>`).join('') + '</tr>';
  });
  html += '</tbody></table>';
  container.innerHTML = html;
}

function nombreSinExtension(nombre) {
  return nombre.replace(/\.[^/.]+$/, '');
}

function formatCantidad(n) {
  const v = Number(n) || 0;
  const redondeado = Number.isInteger(v) ? v : Math.round(v * 100) / 100;
  return redondeado.toLocaleString('es-CO');
}

function formatCOP(n) {
  return '$ ' + Math.round(Number(n) || 0).toLocaleString('es-CO');
}

function renderBarraPct(pct) {
  const p = Math.max(0, Math.min(100, pct));
  return `<span class="pct-num">${pct.toFixed(1)}%</span><span class="bar-bg"><span class="bar-fill" style="width:${p}%"></span></span>`;
}

function renderAnalisisOCN(container, an) {
  container.insertAdjacentHTML('beforeend', '<h3 style="margin-top:24px;">📊 Análisis rápido antes de descargar</h3>');
  renderMetrics(container, [
    { valor: formatCantidad(an.cantidadTotal), etiqueta: 'Cantidad total a comprar' },
    { valor: formatCOP(an.valorTotal), etiqueta: 'Valor total' },
    { valor: an.numProveedores, etiqueta: 'Proveedores' },
    { valor: an.numItems, etiqueta: 'Ítems distintos' },
  ]);

  let html = '<h4 style="margin:18px 0 6px;">Por proveedor</h4>';
  html += '<table class="preview"><thead><tr><th>NIT</th><th>Proveedor</th><th>Cantidad</th><th>Valor</th><th>% del valor total</th></tr></thead><tbody>';
  an.porProveedor.slice(0, 10).forEach(f => {
    html += `<tr><td>${f.clave}</td><td>${f.nombre || ''}</td><td>${formatCantidad(f.cantidad)}</td><td>${formatCOP(f.valor)}</td><td>${renderBarraPct(f.pct)}</td></tr>`;
  });
  html += '</tbody></table>';
  if (an.porProveedor.length > 10) {
    html += `<p style="font-size:12px;color:var(--gris);">Mostrando los 10 proveedores de mayor valor, de ${an.porProveedor.length} en total.</p>`;
  }

  html += '<h4 style="margin:18px 0 6px;">Por centro</h4>';
  html += '<table class="preview"><thead><tr><th>Centro</th><th>Cantidad</th><th>Valor</th><th>% del valor total</th></tr></thead><tbody>';
  an.porCentro.slice(0, 10).forEach(f => {
    html += `<tr><td>${f.clave}</td><td>${formatCantidad(f.cantidad)}</td><td>${formatCOP(f.valor)}</td><td>${renderBarraPct(f.pct)}</td></tr>`;
  });
  html += '</tbody></table>';
  if (an.porCentro.length > 10) {
    html += `<p style="font-size:12px;color:var(--gris);">Mostrando los 10 centros de mayor valor, de ${an.porCentro.length} en total.</p>`;
  }

  html += '<h4 style="margin:18px 0 6px;">Top 10 ítems por valor</h4>';
  html += '<table class="preview"><thead><tr><th>Item</th><th>Descripción</th><th>Cantidad</th><th>Valor</th><th>% del valor total</th></tr></thead><tbody>';
  an.topItems.forEach(f => {
    html += `<tr><td>${f.clave}</td><td>${f.nombre || ''}</td><td>${formatCantidad(f.cantidad)}</td><td>${formatCOP(f.valor)}</td><td>${renderBarraPct(f.pct)}</td></tr>`;
  });
  html += '</tbody></table>';

  container.insertAdjacentHTML('beforeend', html);
}

function fechaHoyCompacta() {
  const d = new Date();
  return `${d.getFullYear()}${pad2(d.getMonth() + 1)}${pad2(d.getDate())}`;
}

/* ==========================================================================
   Maestros: guardado en IndexedDB (nunca sale del navegador)
   IndexedDB tiene un límite de almacenamiento muchísimo más alto que
   localStorage (normalmente varios cientos de MB o más, según el navegador),
   así que un maestro de proveedores de unas pocas miles de filas no debería
   volver a toparse con un error de cuota.
   ========================================================================== */
const LS_PROVEEDORES = 'scc_ocn_maestro_proveedores';
const LS_PRECIOS = 'scc_ocn_maestro_precios';
const IDB_NOMBRE = 'scc_ocn_db';
const IDB_TIENDA = 'maestros';

// Respaldo en memoria: si el navegador no permite guardar de forma
// permanente -por ejemplo al abrir el archivo directamente con file:// en
// vez de publicarlo con GitHub Pages-, el maestro igual queda disponible
// mientras esta pestaña siga abierta, aunque no sobreviva a un cierre o
// recarga de la página.
const maestrosMemoria = {};
let idbPromesa = null;

function abrirIDB() {
  if (!window.indexedDB) return Promise.reject(new Error('IndexedDB no disponible'));
  if (idbPromesa) return idbPromesa;
  idbPromesa = new Promise((resolve, reject) => {
    const req = indexedDB.open(IDB_NOMBRE, 1);
    req.onupgradeneeded = () => {
      if (!req.result.objectStoreNames.contains(IDB_TIENDA)) {
        req.result.createObjectStore(IDB_TIENDA);
      }
    };
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
  return idbPromesa;
}

async function guardarMaestro(clave, rows, nombreArchivo) {
  const payload = { rows, nombreArchivo, fecha: formatFecha(new Date()) };
  maestrosMemoria[clave] = payload;
  try {
    const db = await abrirIDB();
    await new Promise((resolve, reject) => {
      const tx = db.transaction(IDB_TIENDA, 'readwrite');
      tx.objectStore(IDB_TIENDA).put(payload, clave);
      tx.oncomplete = () => resolve();
      tx.onerror = () => reject(tx.error);
      tx.onabort = () => reject(tx.error);
    });
    return { persistido: true };
  } catch (err) {
    return { persistido: false, error: err };
  }
}

async function cargarMaestro(clave) {
  if (maestrosMemoria[clave]) return maestrosMemoria[clave];
  try {
    const db = await abrirIDB();
    const payload = await new Promise((resolve, reject) => {
      const tx = db.transaction(IDB_TIENDA, 'readonly');
      const req = tx.objectStore(IDB_TIENDA).get(clave);
      req.onsuccess = () => resolve(req.result || null);
      req.onerror = () => reject(req.error);
    });
    if (payload) maestrosMemoria[clave] = payload;
    return payload;
  } catch (err) {
    return null;
  }
}

async function actualizarEstadoMaestros() {
  const prov = await cargarMaestro(LS_PROVEEDORES);
  const precios = await cargarMaestro(LS_PRECIOS);
  document.getElementById('estado-proveedores').textContent = prov
    ? `Cargado: ${prov.rows.length} proveedores (archivo: ${prov.nombreArchivo}).`
    : 'No hay maestro de proveedores cargado todavía.';
  document.getElementById('estado-precios').textContent = precios
    ? `Cargado: ${precios.rows.length} registros de precio (archivo: ${precios.nombreArchivo}).`
    : 'No hay lista de precios cargada todavía.';
}

async function actualizarEstadoOCN() {
  const prov = await cargarMaestro(LS_PROVEEDORES);
  const precios = await cargarMaestro(LS_PRECIOS);
  const listo = prov && precios;
  document.getElementById('ocn-sin-maestros').style.display = listo ? 'none' : 'block';
  document.getElementById('ocn-form').style.display = listo ? 'block' : 'none';
}

// Si el repositorio trae data/proveedores.xlsx y/o data/lista_precios.xlsx,
// se cargan solos al abrir la página -así todo el equipo ve siempre los
// mismos maestros sin que cada quien tenga que subirlos-. Si esos archivos
// no existen (o la página se abrió como archivo local, donde esto no
// funciona por restricciones del navegador), no pasa nada: sigue disponible
// la carga manual de siempre en la pestaña Maestros.
async function intentarCargarMaestrosDesdeRepo() {
  const [provRes, precRes] = await Promise.allSettled([
    fetch('data/proveedores.xlsx').then(r => { if (!r.ok) throw new Error('no encontrado'); return r.arrayBuffer(); }),
    fetch('data/lista_precios.xlsx').then(r => { if (!r.ok) throw new Error('no encontrado'); return r.arrayBuffer(); }),
  ]);

  if (provRes.status === 'fulfilled') {
    try {
      const rows = parsearWorkbookArrayBuffer(provRes.value);
      validarColumnas(rows, REQUIRED_PROVEEDOR_COLUMNS, "maestro de proveedores");
      await guardarMaestro(LS_PROVEEDORES, rows, 'data/proveedores.xlsx (repositorio)');
    } catch (err) { /* archivo del repo inválido: se ignora, queda la carga manual */ }
  }

  if (precRes.status === 'fulfilled') {
    try {
      const rows = parsearWorkbookArrayBuffer(precRes.value);
      validarColumnas(rows, REQUIRED_PRECIO_COLUMNS, "lista de precios");
      await guardarMaestro(LS_PRECIOS, rows, 'data/lista_precios.xlsx (repositorio)');
    } catch (err) { /* archivo del repo inválido: se ignora, queda la carga manual */ }
  }
}

function avisarResultadoGuardado(msg, resultado, etiqueta, filas) {
  if (resultado.persistido) {
    renderAlert(msg, 'success', `${etiqueta} guardado (${filas} filas). Quedará disponible la próxima vez que entres a esta página.`);
    return;
  }
  const motivoFile = location.protocol === 'file:'
    ? ' Esto suele pasar porque abriste el archivo directamente en tu computador (file://); publícala con GitHub Pages para que sí quede guardada entre visitas.'
    : ' El navegador no tiene espacio disponible en su almacenamiento local.';
  renderAlert(msg, 'warning', `${etiqueta} cargado (${filas} filas) y disponible mientras no cierres o recargues esta pestaña, pero no se pudo guardar de forma permanente.${motivoFile}`);
}

document.getElementById('prov-file').addEventListener('change', async (e) => {
  const file = e.target.files[0];
  const msg = document.getElementById('prov-msg');
  msg.innerHTML = '';
  if (!file) return;
  try {
    const rows = await leerExcel(file);
    validarColumnas(rows, REQUIRED_PROVEEDOR_COLUMNS, "maestro de proveedores");
    const resultado = await guardarMaestro(LS_PROVEEDORES, rows, file.name);
    avisarResultadoGuardado(msg, resultado, 'Maestro de proveedores', rows.length);
    await actualizarEstadoMaestros();
  } catch (err) {
    renderAlert(msg, 'error', err.message);
  }
});

document.getElementById('precios-file').addEventListener('change', async (e) => {
  const file = e.target.files[0];
  const msg = document.getElementById('precios-msg');
  msg.innerHTML = '';
  if (!file) return;
  try {
    const rows = await leerExcel(file);
    validarColumnas(rows, REQUIRED_PRECIO_COLUMNS, "lista de precios");
    const resultado = await guardarMaestro(LS_PRECIOS, rows, file.name);
    avisarResultadoGuardado(msg, resultado, 'Lista de precios', rows.length);
    await actualizarEstadoMaestros();
  } catch (err) {
    renderAlert(msg, 'error', err.message);
  }
});

/* ==========================================================================
   Módulo SCC: interfaz
   ========================================================================== */
let sccRowsActuales = null;

document.getElementById('scc-file').addEventListener('change', async (e) => {
  const file = e.target.files[0];
  const resultado = document.getElementById('scc-resultado');
  resultado.innerHTML = '';
  if (!file) return;
  try {
    sccRowsActuales = await leerExcel(file);
    document.getElementById('scc-notas').value = nombreSinExtension(file.name);
    renderPreview(document.getElementById('scc-preview'), sccRowsActuales);
    document.getElementById('scc-config').style.display = 'block';
  } catch (err) {
    renderAlert(resultado, 'error', 'No se pudo leer el archivo: ' + err.message);
  }
});

document.getElementById('scc-generar').addEventListener('click', () => {
  const resultado = document.getElementById('scc-resultado');
  resultado.innerHTML = '';
  const notas = document.getElementById('scc-notas').value || 'SCC';
  try {
    const { wb, resumen } = generarSCC(sccRowsActuales, notas);
    renderAlert(resultado, 'success', 'Archivo SCC generado correctamente.');
    renderMetrics(resultado, [
      { valor: resumen.total_filas_entrada, etiqueta: 'Filas leídas' },
      { valor: resumen.documentos_generados, etiqueta: 'Documentos generados' },
      { valor: resumen.movimientos_generados, etiqueta: 'Movimientos generados' },
    ]);
    if (resumen.descartadas_por_estado) {
      renderAlert(resultado, 'warning', `${resumen.descartadas_por_estado} fila(s) se descartaron por no tener Estado = 'Aprobado'.`);
    }
    const nombreArchivo = document.getElementById('scc-file').files[0].name;
    const nombreSalida = `SCC_${nombreSinExtension(nombreArchivo)}_${fechaHoyCompacta()}.xlsx`;
    const btn = document.createElement('button');
    btn.className = 'primary';
    btn.textContent = '⬇️ Descargar archivo plano SCC';
    btn.style.marginTop = '10px';
    btn.onclick = () => XLSX.writeFile(wb, nombreSalida);
    resultado.appendChild(btn);
  } catch (err) {
    renderAlert(resultado, 'error', err.message);
  }
});

/* ==========================================================================
   Módulo OCN: interfaz
   ========================================================================== */
let ocnRowsActuales = null;

document.getElementById('ocn-file').addEventListener('change', async (e) => {
  const file = e.target.files[0];
  const resultado = document.getElementById('ocn-resultado');
  resultado.innerHTML = '';
  if (!file) return;
  try {
    ocnRowsActuales = await leerExcel(file);
    renderPreview(document.getElementById('ocn-preview'), ocnRowsActuales);
    document.getElementById('ocn-config').style.display = 'block';
  } catch (err) {
    renderAlert(resultado, 'error', 'No se pudo leer el archivo: ' + err.message);
  }
});

document.getElementById('ocn-generar').addEventListener('click', async () => {
  const resultado = document.getElementById('ocn-resultado');
  resultado.innerHTML = '';
  const terceroComprador = document.getElementById('ocn-comprador').value || '901906381';
  try {
    const prov = await cargarMaestro(LS_PROVEEDORES);
    const precios = await cargarMaestro(LS_PRECIOS);
    if (!prov || !precios) {
      throw new Error('No hay maestro de proveedores y/o lista de precios cargados. Ve a la pestaña Maestros y súbelos primero.');
    }
    const { wb, resumen, analisis } = generarOCN(ocnRowsActuales, prov.rows, precios.rows, terceroComprador);
    renderAlert(resultado, 'success', 'Archivo OCN generado correctamente.');
    renderMetrics(resultado, [
      { valor: resumen.total_filas_entrada, etiqueta: 'Filas leídas' },
      { valor: resumen.documentos_generados, etiqueta: 'Documentos generados' },
      { valor: resumen.movimientos_generados, etiqueta: 'Movimientos generados' },
    ]);
    if (resumen.descartadas_por_estado) {
      renderAlert(resultado, 'warning', `${resumen.descartadas_por_estado} fila(s) se descartaron por no tener Estado = 'Aprobado'.`);
    }
    if (resumen.filas_sin_cruce) {
      renderAlert(resultado, 'warning', `${resumen.filas_sin_cruce} fila(s) no se pudieron cruzar con la lista de precios o el maestro de proveedores y quedaron en la hoja 'No cruzados' del archivo generado.`);
    }
    renderAnalisisOCN(resultado, analisis);
    const nombreArchivo = document.getElementById('ocn-file').files[0].name;
    const nombreSalida = `OCN_${nombreSinExtension(nombreArchivo)}_${fechaHoyCompacta()}.xlsx`;
    const btn = document.createElement('button');
    btn.className = 'primary';
    btn.textContent = '⬇️ Descargar archivo plano OCN';
    btn.style.marginTop = '10px';
    btn.onclick = () => XLSX.writeFile(wb, nombreSalida);
    resultado.appendChild(btn);
  } catch (err) {
    renderAlert(resultado, 'error', err.message);
  }
});

/* ==========================================================================
   Inicialización
   ========================================================================== */
if (location.protocol === 'file:') {
  document.getElementById('aviso-file-protocol').style.display = 'block';
}
(async () => {
  await intentarCargarMaestrosDesdeRepo();
  await actualizarEstadoMaestros();
  await actualizarEstadoOCN();
})();
</script>

</body>
</html>
