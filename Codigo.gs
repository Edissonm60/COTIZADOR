// ============================================================
//  SecureQuote Pro — Google Apps Script Backend
//  INSTRUCCIONES:
//  1. Abre tu Google Sheet
//  2. Ve a Extensiones → Apps Script
//  3. Borra todo el contenido y pega este código
//  4. Guarda (Ctrl+S) y ya funcionará correctamente
// ============================================================

// ─── Obtener SpreadSheet de forma segura ─────────────────────
function getSS() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

// Nombres de hojas
const SHEETS = {
  PRODUCTOS: 'Productos',
  COTIZACIONES: 'Cotizaciones',
  ITEMS: 'Items_Cotizacion',
  CONFIG: 'Config'
};

// ─── Inicializar hojas al abrir ───────────────────────────────
function onOpen() {
  initSheets();
}

function initSheets() {
  const ss = getSS();
  Object.values(SHEETS).forEach(name => {
    if (!ss.getSheetByName(name)) {
      ss.insertSheet(name);
    }
  });

  // Cabeceras hoja PRODUCTOS
  const sp = ss.getSheetByName(SHEETS.PRODUCTOS);
  if (sp.getLastRow() === 0) {
    sp.appendRow(['ID','Nombre','Categoria','SKU','Descripcion','Precio','Unidad','ImagenURL','FechaCreado']);
    sp.getRange(1,1,1,9).setFontWeight('bold').setBackground('#00d4aa').setFontColor('#000');
  }

  // Cabeceras hoja COTIZACIONES
  const sc = ss.getSheetByName(SHEETS.COTIZACIONES);
  if (sc.getLastRow() === 0) {
    sc.appendRow(['ID','Numero','Fecha','ValidaHasta','EmpresaNombre','EmpresaNIT','EmpresaContacto','EmpresaDireccion','ClienteNombre','ClienteNIT','ClienteContacto','ClienteDireccion','Moneda','IVA_Activo','Utilidad_Pct','Subtotal','Utilidad_Valor','IVA_Valor','Total','Notas','Estado','FechaCreado']);
    sc.getRange(1,1,1,22).setFontWeight('bold').setBackground('#0066ff').setFontColor('#fff');
  }

  // Cabeceras hoja ITEMS
  const si = ss.getSheetByName(SHEETS.ITEMS);
  if (si.getLastRow() === 0) {
    si.appendRow(['CotizacionID','ProductoID','Nombre','Descripcion','Cantidad','PrecioUnitario','Total','ImagenURL']);
    si.getRange(1,1,1,8).setFontWeight('bold').setBackground('#534AB7').setFontColor('#fff');
  }
}

// ─── Web App Entry Point ──────────────────────────────────────
function doGet(e) {
  return HtmlService.createHtmlOutput('SecureQuote Pro API running ✓')
    .setTitle('SecureQuote API');
}

function doPost(e) {
  const data = JSON.parse(e.postData.contents);
  const action = data.action;
  let result;

  try {
    switch(action) {
      case 'getProductos':     result = getProductos(); break;
      case 'saveProducto':     result = saveProducto(data.payload); break;
      case 'deleteProducto':   result = deleteProducto(data.payload.id); break;
      case 'getCotizaciones':  result = getCotizaciones(); break;
      case 'saveCotizacion':   result = saveCotizacion(data.payload); break;
      case 'deleteCotizacion': result = deleteCotizacion(data.payload.id); break;
      case 'getConfig':        result = getConfig(); break;
      case 'saveConfig':       result = saveConfig(data.payload); break;
      default: result = { ok: false, error: 'Acción no reconocida: ' + action };
    }
  } catch(err) {
    result = { ok: false, error: err.toString() };
  }

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ─── PRODUCTOS ────────────────────────────────────────────────
function getProductos() {
  const ss = getSS();
  const sh = ss.getSheetByName(SHEETS.PRODUCTOS);
  const data = sh.getDataRange().getValues();
  if (data.length <= 1) return { ok: true, data: [] };
  const headers = data[0];
  const rows = data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => obj[h] = row[i]);
    return obj;
  }).filter(p => p.ID);
  return { ok: true, data: rows };
}

function saveProducto(p) {
  const ss = getSS();
  const sh = ss.getSheetByName(SHEETS.PRODUCTOS);
  const id = p.ID || 'P-' + Date.now();
  const now = new Date().toISOString();

  const data = sh.getDataRange().getValues();
  let rowIndex = -1;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) { rowIndex = i + 1; break; }
  }

  const row = [id, p.Nombre, p.Categoria, p.SKU, p.Descripcion, p.Precio, p.Unidad, p.ImagenURL, now];

  if (rowIndex > 0) {
    sh.getRange(rowIndex, 1, 1, 9).setValues([row]);
  } else {
    sh.appendRow(row);
  }
  return { ok: true, id };
}

function deleteProducto(id) {
  const ss = getSS();
  const sh = ss.getSheetByName(SHEETS.PRODUCTOS);
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) {
      sh.deleteRow(i + 1);
      return { ok: true };
    }
  }
  return { ok: false, error: 'Producto no encontrado' };
}

// ─── COTIZACIONES ─────────────────────────────────────────────
function getCotizaciones() {
  const ss = getSS();
  const sc = ss.getSheetByName(SHEETS.COTIZACIONES);
  const si = ss.getSheetByName(SHEETS.ITEMS);

  const cData = sc.getDataRange().getValues();
  const iData = si.getDataRange().getValues();

  if (cData.length <= 1) return { ok: true, data: [] };

  const cHeaders = cData[0];
  const iHeaders = iData[0];

  const items = iData.slice(1).map(row => {
    const obj = {};
    iHeaders.forEach((h, i) => obj[h] = row[i]);
    return obj;
  });

  const cotizaciones = cData.slice(1).map(row => {
    const obj = {};
    cHeaders.forEach((h, i) => obj[h] = row[i]);
    obj.items = items.filter(it => String(it.CotizacionID) === String(obj.ID));
    return obj;
  }).filter(c => c.ID);

  return { ok: true, data: cotizaciones };
}

function saveCotizacion(c) {
  const ss = getSS();
  const sc = ss.getSheetByName(SHEETS.COTIZACIONES);
  const si = ss.getSheetByName(SHEETS.ITEMS);
  const id = c.ID || 'C-' + Date.now();
  const now = new Date().toISOString();

  const row = [
    id, c.Numero, c.Fecha, c.ValidaHasta,
    c.EmpresaNombre, c.EmpresaNIT, c.EmpresaContacto, c.EmpresaDireccion,
    c.ClienteNombre, c.ClienteNIT, c.ClienteContacto, c.ClienteDireccion,
    c.Moneda, c.IVA_Activo, c.Utilidad_Pct,
    c.Subtotal, c.Utilidad_Valor, c.IVA_Valor, c.Total,
    c.Notas, c.Estado || 'Borrador', now
  ];

  const cData = sc.getDataRange().getValues();
  let rowIndex = -1;
  for (let i = 1; i < cData.length; i++) {
    if (String(cData[i][0]) === String(id)) { rowIndex = i + 1; break; }
  }

  if (rowIndex > 0) {
    sc.getRange(rowIndex, 1, 1, 22).setValues([row]);
  } else {
    sc.appendRow(row);
  }

  // Eliminar items anteriores de esta cotización
  const iData = si.getDataRange().getValues();
  for (let i = iData.length - 1; i >= 1; i--) {
    if (String(iData[i][0]) === String(id)) si.deleteRow(i + 1);
  }

  // Insertar items nuevos
  if (c.items && c.items.length > 0) {
    c.items.forEach(it => {
      si.appendRow([id, it.ProductoID || '', it.Nombre, it.Descripcion || '', it.Cantidad, it.PrecioUnitario, it.Total, it.ImagenURL || '']);
    });
  }

  return { ok: true, id };
}

function deleteCotizacion(id) {
  const ss = getSS();
  const sc = ss.getSheetByName(SHEETS.COTIZACIONES);
  const si = ss.getSheetByName(SHEETS.ITEMS);

  const cData = sc.getDataRange().getValues();
  for (let i = 1; i < cData.length; i++) {
    if (String(cData[i][0]) === String(id)) { sc.deleteRow(i + 1); break; }
  }

  const iData = si.getDataRange().getValues();
  for (let i = iData.length - 1; i >= 1; i--) {
    if (String(iData[i][0]) === String(id)) si.deleteRow(i + 1);
  }

  return { ok: true };
}

// ─── CONFIG ───────────────────────────────────────────────────
function getConfig() {
  const ss = getSS();
  const sh = ss.getSheetByName(SHEETS.CONFIG);
  const data = sh.getDataRange().getValues();
  const config = {};
  data.forEach(row => { if (row[0]) config[row[0]] = row[1]; });
  return { ok: true, data: config };
}

function saveConfig(cfg) {
  const ss = getSS();
  const sh = ss.getSheetByName(SHEETS.CONFIG);
  sh.clearContents();
  Object.entries(cfg).forEach(([k, v]) => sh.appendRow([k, v]));
  return { ok: true };
}
