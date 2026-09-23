/* ============================================================
   MONITOREO TI · SOFITASA — Backend minimalista
   ============================================================ */

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Monitoreo TI · Sofitasa')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/* ---------- Helpers ---------- */
function num_(v) {
  if (v === '' || v === null || v === undefined) return 0;
  if (typeof v === 'number') return isFinite(v) ? v : 0;
  var x = Number(String(v).replace(/\./g, '').replace(',', '.'));
  return isFinite(x) ? x : 0;
}
function date_(v) {
  if (!v) return '';
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'dd/MM');
  return String(v);
}
function agg_(rows, kT, kA, kR) {
  var sum = function(k) { return rows.reduce(function(a, r) { return a + (Number(r[k]) || 0); }, 0); };
  var t = sum(kT), a = sum(kA), r = sum(kR);
  return {
    total: t, aprob: a, rech: r,
    tasaAprob: t ? a / t : 0,
    tasaRech: t ? r / t : 0
  };
}
function sheet_(name) {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
}

/* ---------- Readers ---------- */

function readerPOSPropios_() {
  var sh = sheet_('POS propios'); if (!sh) return { rows: [], kpi: null };
  var v = sh.getRange(3, 1, 30, 12).getValues();
  var rows = v.filter(function(r){ return r[0]; }).map(function(r){
    return {
      fecha: date_(r[0]), total: num_(r[1]), aprob: num_(r[2]), rech: num_(r[3]),
      claves: num_(r[4]), cedulas: num_(r[5]), chip: num_(r[6]), fondos: num_(r[7]),
      reversos: num_(r[8]), emisor: num_(r[9]), otros: num_(r[10]), timeout: num_(r[11])
    };
  });
  return { rows: rows, kpi: agg_(rows, 'total', 'aprob', 'rech') };
}

function readerPOSOtros_() {
  var sh = sheet_('POS otros bancos'); if (!sh) return { rows: [], kpi: null };
  var v = sh.getRange(4, 1, 30, 11).getValues();
  var rows = v.filter(function(r){ return r[0]; }).map(function(r){
    return {
      fecha: date_(r[0]), total: num_(r[1]), aprob: num_(r[2]), rech: num_(r[3]),
      claves: num_(r[4]), cedulas: num_(r[5]), chip: num_(r[6]), fondos: num_(r[7]),
      reversos: num_(r[8]), emisor: num_(r[9]), otros: num_(r[10])
    };
  });
  return { rows: rows, kpi: agg_(rows, 'total', 'aprob', 'rech') };
}

function readerP2P_() {
  var sh = sheet_('P2P'); if (!sh) return { rows: [], kpi: null };
  var v = sh.getRange(3, 1, 30, 35).getValues();
  var rows = v.filter(function(r){ return r[0]; }).map(function(r){
    return {
      fecha: date_(r[0]), total: num_(r[1]), emitidas: num_(r[2]), rech: num_(r[3]),
      tarjetaInv: num_(r[4]), fondo: num_(r[5]), ccNoCheq: num_(r[6]), tlfNoReg: num_(r[7]),
      aprob: num_(r[8]), errorCom: num_(r[9]), extraviada: num_(r[10]),
      retiroLim: num_(r[11]), restring: num_(r[12]), retiroFrec: num_(r[13]),
      bcoRecp: num_(r[14]), emisorInop: num_(r[15]),
      sofnet: num_(r[21]), sofnetOK: num_(r[22]), appMovil: num_(r[24]), appOK: num_(r[25])
    };
  });
  var kpi = agg_(rows, 'emitidas', 'aprob', 'rech');
  var sum = function(k){ return rows.reduce(function(a,r){return a+(Number(r[k])||0);},0); };
  kpi.sofnet   = sum('sofnet');
  kpi.appMovil = sum('appMovil');
  var tot = kpi.sofnet + kpi.appMovil;
  kpi.usoApp    = tot ? kpi.appMovil / tot : 0;
  kpi.usoSofnet = tot ? kpi.sofnet / tot : 0;
  return { rows: rows, kpi: kpi };
}

function readerTransferencias_() {
  var sh = sheet_('Transferencias'); if (!sh) return { rows: [], kpi: null };
  var v = sh.getRange(3, 1, 30, 36).getValues();
  var rows = v.filter(function(r){ return r[0] && (num_(r[1]) || num_(r[2])); }).map(function(r){
    return {
      fecha: date_(r[0]), emitidas: num_(r[1]), recibidas: num_(r[2]),
      cuentaCerrada: num_(r[3]), datosNo: num_(r[4]), receptorOff: num_(r[5]),
      aprob: num_(r[6]), idIncorrecto: num_(r[7]), rechTec: num_(r[8]),
      sinDesc: num_(r[10]), timeout: num_(r[11]), credProhib: num_(r[12]),
      cuentaBloq: num_(r[13]), montoExcede: num_(r[14]), saldoIns: num_(r[17]),
      afilInact: num_(r[18]), noAfiliacion: num_(r[19]), otpInc: num_(r[20]),
      emitidasJ: num_(r[23]), recibidasJ: num_(r[25]),
      emitidasN: num_(r[27]), recibidasN: num_(r[29]), rech: num_(r[32])
    };
  });
  return { rows: rows, kpi: agg_(rows, 'emitidas', 'aprob', 'rech') };
}

function readerCI_() {
  var sh = sheet_('Credito Inmediato'); if (!sh) return { rows: [], kpi: null };
  var v = sh.getRange(3, 1, 30, 30).getValues();
  var rows = v.filter(function(r){ return r[0] && (num_(r[1]) || num_(r[2])); }).map(function(r){
    return {
      fecha: date_(r[0]), emitidas: num_(r[1]), recibidas: num_(r[2]),
      cuentaCerrada: num_(r[3]), datosNo: num_(r[4]), receptorOff: num_(r[5]),
      aprob: num_(r[6]), idIncorrecto: num_(r[7]), rechTec: num_(r[8]),
      sinDesc: num_(r[10]), timeout: num_(r[11]), montoExcede: num_(r[14]),
      saldoIns: num_(r[18]), afilInact: num_(r[19]), noAfiliacion: num_(r[20]),
      emitidasJ: num_(r[21]), recibidasJ: num_(r[23]),
      emitidasN: num_(r[25]), recibidasN: num_(r[27]), rech: num_(r[29])
    };
  });
  return { rows: rows, kpi: agg_(rows, 'emitidas', 'aprob', 'rech') };
}

function readerDI_() {
  var sh = sheet_('Debito Inmediato'); if (!sh) return { rows: [], kpi: null };
  var v = sh.getRange(3, 1, 30, 34).getValues();
  var rows = v.filter(function(r){ return r[0] && (num_(r[1]) || num_(r[2])); }).map(function(r){
    return {
      fecha: date_(r[0]), emitidas: num_(r[1]), recibidas: num_(r[2]),
      cuentaCerrada: num_(r[3]), datosNo: num_(r[4]), receptorOff: num_(r[5]),
      aprob: num_(r[6]), idIncorrecto: num_(r[7]), rechTec: num_(r[8]),
      sinDesc: num_(r[10]), timeout: num_(r[11]), montoExcede: num_(r[14]),
      saldoIns: num_(r[18]), afilInact: num_(r[19]), noAfiliacion: num_(r[20]),
      otpInc: num_(r[21]), codSubprod: num_(r[22]), codBancoNo: num_(r[23]),
      emitidasJ: num_(r[25]), recibidasJ: num_(r[27]),
      emitidasN: num_(r[29]), recibidasN: num_(r[31]), rech: num_(r[33])
    };
  });
  return { rows: rows, kpi: agg_(rows, 'emitidas', 'aprob', 'rech') };
}

function readerEstDiaria_() {
  var sh = sheet_('Estadistica diaria'); if (!sh) return { rows: [] };
  var last = sh.getLastRow();
  if (last < 2) return { rows: [] };
  var v = sh.getRange(2, 1, last - 1, 5).getValues();
  var rows = v.filter(function(r){ return r[0] || r[1] || r[2] || r[3] || r[4]; }).map(function(r){
    return {
      titulo: String(r[0] || ''),
      estadistica: String(r[1] || ''),
      observaciones: String(r[2] || ''),
      cierreT1: String(r[3] || ''),
      cierreT2: String(r[4] || '')
    };
  });
  return { rows: rows };
}

function readerEstSemanal_() {
  var sh = sheet_('Est_Semales'); if (!sh) return { rows: [] };
  var last = sh.getLastRow();
  if (last < 1) return { rows: [] };
  var v = sh.getRange(1, 1, last, 7).getValues();
  var rows = v.filter(function(r){ return r[1] || r[2]; }).map(function(r){
    return {
      indicador: String(r[1] || ''),
      s1: r[2], s2: r[3], s3: r[4], s4: r[5], total: r[6]
    };
  });
  return { rows: rows };
}

function readerEstMensual_() {
  var sh = sheet_('Est_Trimestrales'); if (!sh) return { matrix: [] };
  var lastR = Math.max(sh.getLastRow(), 1);
  var lastC = Math.max(sh.getLastColumn(), 1);
  var v = sh.getRange(1, 1, lastR, lastC).getValues();
  var m = v.map(function(row){
    return row.map(function(c){
      if (c instanceof Date) return date_(c);
      if (c === null || c === undefined) return '';
      return c;
    });
  });
  return { matrix: m };
}

/* ---------- APIs expuestas al cliente ---------- */

function getPageData(pageId) {
  try {
    var map = {
      'pos-propios':        readerPOSPropios_,
      'pos-otros':          readerPOSOtros_,
      'p2p':                readerP2P_,
      'transferencias':     readerTransferencias_,
      'credito-inmediato':  readerCI_,
      'debito-inmediato':   readerDI_,
      'estadistica-diaria': readerEstDiaria_,
      'est-semanales':      readerEstSemanal_,
      'est-trimestrales':   readerEstMensual_
    };
    var fn = map[pageId];
    if (!fn) return { ok: false, error: 'Página desconocida: ' + pageId };
    return { ok: true, data: fn() };
  } catch (err) {
    return { ok: false, error: err.message, stack: err.stack };
  }
}

function getDashboardData() {
  try {
    var p = readerPOSPropios_();
    var o = readerPOSOtros_();
    var q = readerP2P_();
    var t = readerTransferencias_();
    return {
      ok: true,
      data: {
        posPropios: p.kpi, posOtros: o.kpi, p2p: q.kpi, transferencias: t.kpi,
        posRows: p.rows
      }
    };
  } catch (err) {
    return { ok: false, error: err.message };
  }
}

/* ---------- Debug ---------- */
function debugReaders() {
  Logger.log(JSON.stringify({
    pos:     readerPOSPropios_().rows.length,
    posOtros:readerPOSOtros_().rows.length,
    p2p:     readerP2P_().rows.length,
    transf:  readerTransferencias_().rows.length,
    ci:      readerCI_().rows.length,
    di:      readerDI_().rows.length,
    dia:     readerEstDiaria_().rows.length,
    sem:     readerEstSemanal_().rows.length,
    trim:    readerEstMensual_().matrix.length
  }, null, 2));
}
