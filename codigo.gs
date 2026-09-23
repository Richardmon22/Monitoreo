/* ============================================================
   MONITOREO TI · SOFITASA — Backend completo
   Cálculo correcto de Tasa de Aprobación por canal
   ============================================================ */

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Monitoreo TI · Sofitasa')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/* ============================================================
   HELPERS
   ============================================================ */
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

/**
 * Agregador de KPIs.
 * kT = campo total (denominador), kA = aprobadas (numerador), kR = rechazadas.
 * Si kT es null o 0, se calcula como kA + kR para asegurar tasa válida.
 */
function agg_(rows, kT, kA, kR) {
  var sum = function(k) {
    return rows.reduce(function(a, r) { return a + (Number(r[k]) || 0); }, 0);
  };
  var a = sum(kA);
  var r = sum(kR);
  var t = kT ? sum(kT) : (a + r);
  // Fallback: si el total viene en 0 pero hay datos, usar a + r
  if (!t && (a || r)) t = a + r;
  return {
    total: t,
    aprob: a,
    rech:  r,
    tasaAprob: t ? a / t : 0,
    tasaRech:  t ? r / t : 0
  };
}

function sheet_(name) {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
}

/* ============================================================
   READERS
   ============================================================ */

/* ---------- POS PROPIOS ----------
   Col B = Total (fórmula =C+D)
   Col C = Aprobadas
   Col D = Rechazadas (suma de E:L)
   → tasaAprob = C / B                                          */
function readerPOSPropios_() {
  var sh = sheet_('POS propios'); if (!sh) return { rows: [], kpi: null };
  var v = sh.getRange(3, 1, 30, 12).getValues();
  var rows = v.filter(function(r){ return r[0]; }).map(function(r){
    var aprob = num_(r[2]);
    var rech  = num_(r[3]);
    var total = num_(r[1]);
    if (!total) total = aprob + rech; // fallback si B viene vacío
    return {
      fecha: date_(r[0]),
      total: total,
      aprob: aprob,
      rech:  rech,
      tasaAprob: total ? aprob / total : 0,
      tasaRech:  total ? rech  / total : 0,
      claves:   num_(r[4]),
      cedulas:  num_(r[5]),
      chip:     num_(r[6]),
      fondos:   num_(r[7]),
      reversos: num_(r[8]),
      emisor:   num_(r[9]),
      otros:    num_(r[10]),
      timeout:  num_(r[11])
    };
  });
  return { rows: rows, kpi: agg_(rows, 'total', 'aprob', 'rech') };
}

/* ---------- POS OTROS BANCOS ---------- */
function readerPOSOtros_() {
  var sh = sheet_('POS otros bancos'); if (!sh) return { rows: [], kpi: null };
  var v = sh.getRange(4, 1, 30, 11).getValues();
  var rows = v.filter(function(r){ return r[0]; }).map(function(r){
    var aprob = num_(r[2]);
    var rech  = num_(r[3]);
    var total = num_(r[1]);
    if (!total) total = aprob + rech;
    return {
      fecha: date_(r[0]),
      total: total,
      aprob: aprob,
      rech:  rech,
      tasaAprob: total ? aprob / total : 0,
      tasaRech:  total ? rech  / total : 0,
      claves:   num_(r[4]),
      cedulas:  num_(r[5]),
      chip:     num_(r[6]),
      fondos:   num_(r[7]),
      reversos: num_(r[8]),
      emisor:   num_(r[9]),
      otros:    num_(r[10])
    };
  });
  return { rows: rows, kpi: agg_(rows, 'total', 'aprob', 'rech') };
}

/* ---------- P2P ----------
   Col C = Emitidas (denominador)
   Col I = Aprobadas (numerador)                                */
function readerP2P_() {
  var sh = sheet_('P2P'); if (!sh) return { rows: [], kpi: null };
  var v = sh.getRange(3, 1, 30, 35).getValues();
  var rows = v.filter(function(r){ return r[0]; }).map(function(r){
    var emitidas = num_(r[2]);
    var aprob    = num_(r[8]);
    var rech     = num_(r[3]);
    var total    = emitidas || (aprob + rech);
    return {
      fecha: date_(r[0]),
      total: total,
      emitidas: emitidas,
      rech: rech,
      aprob: aprob,
      tasaAprob: total ? aprob / total : 0,
      tasaRech:  total ? rech  / total : 0,
      tarjetaInv: num_(r[4]),
      fondo:      num_(r[5]),
      ccNoCheq:   num_(r[6]),
      tlfNoReg:   num_(r[7]),
      errorCom:   num_(r[9]),
      extraviada: num_(r[10]),
      retiroLim:  num_(r[11]),
      restring:   num_(r[12]),
      retiroFrec: num_(r[13]),
      bcoRecp:    num_(r[14]),
      emisorInop: num_(r[15]),
      sofnet:     num_(r[21]),
      sofnetOK:   num_(r[22]),
      appMovil:   num_(r[24]),
      appOK:      num_(r[25])
    };
  });
  var kpi = agg_(rows, 'emitidas', 'aprob', 'rech');
  var sum = function(k){ return rows.reduce(function(a,r){return a+(Number(r[k])||0);},0); };
  kpi.sofnet   = sum('sofnet');
  kpi.appMovil = sum('appMovil');
  var tot = kpi.sofnet + kpi.appMovil;
  kpi.usoApp    = tot ? kpi.appMovil / tot : 0;
  kpi.usoSofnet = tot ? kpi.sofnet   / tot : 0;
  return { rows: rows, kpi: kpi };
}

/* ---------- TRANSFERENCIAS ----------
   Col B = Emitidas (denominador)
   Col G = Aprobadas (numerador)                                */
function readerTransferencias_() {
  var sh = sheet_('Transferencias'); if (!sh) return { rows: [], kpi: null };
  var v = sh.getRange(3, 1, 30, 36).getValues();
  var rows = v.filter(function(r){ return r[0] && (num_(r[1]) || num_(r[2])); }).map(function(r){
    var emitidas = num_(r[1]);
    var aprob    = num_(r[6]);
    var rech     = num_(r[32]);
    var total    = emitidas || (aprob + rech);
    return {
      fecha: date_(r[0]),
      emitidas: emitidas,
      recibidas: num_(r[2]),
      total: total,
      aprob: aprob,
      rech:  rech,
      tasaAprob: total ? aprob / total : 0,
      tasaRech:  total ? rech  / total : 0,
      cuentaCerrada: num_(r[3]),
      datosNo:       num_(r[4]),
      receptorOff:   num_(r[5]),
      idIncorrecto:  num_(r[7]),
      rechTec:       num_(r[8]),
      sinDesc:       num_(r[10]),
      timeout:       num_(r[11]),
      credProhib:    num_(r[12]),
      cuentaBloq:    num_(r[13]),
      montoExcede:   num_(r[14]),
      saldoIns:      num_(r[17]),
      afilInact:     num_(r[18]),
      noAfiliacion:  num_(r[19]),
      otpInc:        num_(r[20]),
      codBancoNo:    num_(r[21]),
      codSubprod:    num_(r[22]),
      emitidasJ:     num_(r[23]),
      recibidasJ:    num_(r[25]),
      emitidasN:     num_(r[27]),
      recibidasN:    num_(r[29])
    };
  });
  return { rows: rows, kpi: agg_(rows, 'emitidas', 'aprob', 'rech') };
}

/* ---------- CRÉDITO INMEDIATO ---------- */
function readerCI_() {
  var sh = sheet_('Credito Inmediato'); if (!sh) return { rows: [], kpi: null };
  var v = sh.getRange(3, 1, 30, 30).getValues();
  var rows = v.filter(function(r){ return r[0] && (num_(r[1]) || num_(r[2])); }).map(function(r){
    var emitidas = num_(r[1]);
    var aprob    = num_(r[6]);
    var rech     = num_(r[29]);
    var total    = emitidas || (aprob + rech);
    return {
      fecha: date_(r[0]),
      emitidas: emitidas,
      recibidas: num_(r[2]),
      total: total,
      aprob: aprob,
      rech:  rech,
      tasaAprob: total ? aprob / total : 0,
      tasaRech:  total ? rech  / total : 0,
      cuentaCerrada: num_(r[3]),
      datosNo:       num_(r[4]),
      receptorOff:   num_(r[5]),
      idIncorrecto:  num_(r[7]),
      rechTec:       num_(r[8]),
      sinDesc:       num_(r[10]),
      timeout:       num_(r[11]),
      montoExcede:   num_(r[14]),
      saldoIns:      num_(r[18]),
      afilInact:     num_(r[19]),
      noAfiliacion:  num_(r[20]),
      emitidasJ:     num_(r[21]),
      recibidasJ:    num_(r[23]),
      emitidasN:     num_(r[25]),
      recibidasN:    num_(r[27])
    };
  });
  return { rows: rows, kpi: agg_(rows, 'emitidas', 'aprob', 'rech') };
}

/* ---------- DÉBITO INMEDIATO ---------- */
function readerDI_() {
  var sh = sheet_('Debito Inmediato'); if (!sh) return { rows: [], kpi: null };
  var v = sh.getRange(3, 1, 30, 34).getValues();
  var rows = v.filter(function(r){ return r[0] && (num_(r[1]) || num_(r[2])); }).map(function(r){
    var emitidas = num_(r[1]);
    var aprob    = num_(r[6]);
    var rech     = num_(r[33]);
    var total    = emitidas || (aprob + rech);
    return {
      fecha: date_(r[0]),
      emitidas: emitidas,
      recibidas: num_(r[2]),
      total: total,
      aprob: aprob,
      rech:  rech,
      tasaAprob: total ? aprob / total : 0,
      tasaRech:  total ? rech  / total : 0,
      cuentaCerrada: num_(r[3]),
      datosNo:       num_(r[4]),
      receptorOff:   num_(r[5]),
      idIncorrecto:  num_(r[7]),
      rechTec:       num_(r[8]),
      sinDesc:       num_(r[10]),
      timeout:       num_(r[11]),
      montoExcede:   num_(r[14]),
      saldoIns:      num_(r[18]),
      afilInact:     num_(r[19]),
      noAfiliacion:  num_(r[20]),
      otpInc:        num_(r[21]),
      codSubprod:    num_(r[22]),
      codBancoNo:    num_(r[23]),
      cobroNoPer:    num_(r[24]),
      emitidasJ:     num_(r[25]),
      recibidasJ:    num_(r[27]),
      emitidasN:     num_(r[29]),
      recibidasN:    num_(r[31])
    };
  });
  return { rows: rows, kpi: agg_(rows, 'emitidas', 'aprob', 'rech') };
}

/* ---------- ESTADÍSTICA DIARIA ---------- */
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

/* ---------- EST_SEMANALES ---------- */
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

/* ---------- EST_TRIMESTRALES ---------- */
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

/* ============================================================
   APIs expuestas al cliente
   ============================================================ */
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
        posPropios:     p.kpi,
        posOtros:       o.kpi,
        p2p:            q.kpi,
        transferencias: t.kpi,
        posRows:        p.rows
      }
    };
  } catch (err) {
    return { ok: false, error: err.message, stack: err.stack };
  }
}

/* ============================================================
   DEBUG
   ============================================================ */
function debugTasas() {
  var p = readerPOSPropios_();
  var o = readerPOSOtros_();
  var q = readerP2P_();
  var t = readerTransferencias_();
  var c = readerCI_();
  var d = readerDI_();

  Logger.log('=== TASAS GLOBALES ===');
  Logger.log('POS Propios:      ' + (p.kpi.tasaAprob * 100).toFixed(2) + '%  (' + p.kpi.aprob + ' / ' + p.kpi.total + ')');
  Logger.log('POS Otros:        ' + (o.kpi.tasaAprob * 100).toFixed(2) + '%  (' + o.kpi.aprob + ' / ' + o.kpi.total + ')');
  Logger.log('P2P:              ' + (q.kpi.tasaAprob * 100).toFixed(2) + '%  (' + q.kpi.aprob + ' / ' + q.kpi.total + ')');
  Logger.log('Transferencias:   ' + (t.kpi.tasaAprob * 100).toFixed(2) + '%  (' + t.kpi.aprob + ' / ' + t.kpi.total + ')');
  Logger.log('Crédito Inm:      ' + (c.kpi.tasaAprob * 100).toFixed(2) + '%  (' + c.kpi.aprob + ' / ' + c.kpi.total + ')');
  Logger.log('Débito Inm:       ' + (d.kpi.tasaAprob * 100).toFixed(2) + '%  (' + d.kpi.aprob + ' / ' + d.kpi.total + ')');

  Logger.log('\n=== DETALLE POS PROPIOS (primeros 3 días) ===');
  p.rows.slice(0, 3).forEach(function(r) {
    Logger.log('  ' + r.fecha + ' → total=' + r.total +
      ' | aprob=' + r.aprob +
      ' | rech='  + r.rech  +
      ' | tasa='  + (r.tasaAprob * 100).toFixed(2) + '%');
  });
}
