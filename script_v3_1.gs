// ===== SCRIPT WBR-EDSPs v3.1 =====
// Alteracoes v3.1:
// - processarCSV_FM: Receive/Stow adj (col 2.1/4.1), OTD via FM (col 13. OTD %)
// - preencherTodasAbas + upsert: OTD via FM como PRIMARIO; otd_cpt.csv como fallback
// Mantém APENAS o necessário para alimentar o HTML WBR-EDSP (buildHTML5_)
// Push: WBR-EDSPs/data.js
// Trigger: diário às 16h BRT

function criarAbas() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var abas = {
    "config": ["parametro","valor","descricao"],
    "regional_6w": ["Semana","Regional","Volume","Receive","Stow","Depart","Backlog","Backlog_Acc","PPP","Missort"],
    "nodes_6w": ["Semana","Node","Regional","Grupo","Tipo","Volume","Receive","Stow","Depart","Backlog","Backlog_Acc","PPP","Missort","Wrong_Node"],
    "nodes_w1": ["Semana","Node","Regional","Grupo","Tipo","Volume","Receive","Stow","Depart","Backlog","PPP"],
    "wbr_extras": ["Semana","Node","CPT_Success","OTD","Forecast","Desvio_Forecast","HC","TPH"],
    "tph": ["Semana","Node","Regional","TPH","HC","Horas"],
    "edsps_20w": ["Semana","Node","Grupo","Volume","Receive","Stow","Depart","Backlog","PPP"],
    "regional_ppp_comparativo": ["Semana","Regional","PPP"],
    "shipment_data": ["Semana","Node","Grupo","Tipo","Miss_Real","Miss_D1_Corrigido",
                      "Recv_Real","Recv_Corrigido","Stow_Real","Stow_Corrigido","Dep_Real","Dep_Corrigido"],
    "nodes_daily": ["Data","Node","Grupo","Tipo","Miss_Real","Recv_Miss","Stow_Miss","Dep_Miss","D1_Corrigido","Total"]
  };
  var nomes = Object.keys(abas);
  for (var i = 0; i < nomes.length; i++) {
    var nome  = nomes[i];
    var sheet = ss.getSheetByName(nome);
    if (!sheet) { sheet = ss.insertSheet(nome); }
    var headers = abas[nome];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
  var configSheet = ss.getSheetByName("config");
  var defaults = [
    ["semana_ref","","Semana de referencia (W-0)"],
    ["semana_w1","","Semana parcial (W+1)"],
    ["semanas_range","","Lista de 6 semanas exibidas"],
    ["data_atualizacao","","Data da ultima atualizacao (YYYY-MM-DD)"],
    ["hora_atualizacao","","Hora da ultima atualizacao (HH:MM)"],
    ["meta_ppp","175","Meta PPP"],
    ["meta_recv","98.5","Meta Receive %"],
    ["meta_stow","99.5","Meta Stow %"],
    ["meta_dep","99.5","Meta Departure %"],
    ["meta_tph","58.96","Meta TPH regional"],
    ["regional_nome","DF-SUL","Nome da regional"]
  ];
  configSheet.getRange(2, 1, defaults.length, 3).setValues(defaults);
  Browser.msgBox("Abas criadas com sucesso!");
}


// ============================================================================
// HELPERS
// ============================================================================

function findCol_(headers, name) {
  var names = (typeof name === "string") ? [name] : name;
  var clean = [];
  for (var i = 0; i < headers.length; i++) {
    clean.push(headers[i].replace(/^\uFEFF/, "").trim().toLowerCase());
  }
  for (var i = 0; i < clean.length; i++) {
    for (var n = 0; n < names.length; n++) {
      if (clean[i] === names[n].trim().toLowerCase()) return i;
    }
  }
  for (var i = 0; i < clean.length; i++) {
    for (var n = 0; n < names.length; n++) {
      var target = names[n].trim().toLowerCase();
      if (clean[i].indexOf(target) > -1 || target.indexOf(clean[i]) > -1) return i;
    }
  }
  return -1;
}

function toNumBR_(val) {
  if (val === "" || val === null || val === undefined || val === "Sem Dados") return null;
  var s = String(val).trim();
  s = s.replace(/%/g, "");
  if (s.indexOf(",") > -1 && s.indexOf(".") > -1) {
    s = s.replace(/\./g, "");
    s = s.replace(/,/g, ".");
  } else if (s.indexOf(",") > -1 && s.indexOf(".") === -1) {
    s = s.replace(/,/g, ".");
  }
  var n = parseFloat(s);
  if (isNaN(n)) return null;
  return n;
}

function toNum_(val) {
  if (val === "" || val === null || val === undefined) return null;
  var n = parseFloat(val);
  if (isNaN(n)) return null;
  return n;
}

function toWeekStr_(val) {
  var dt;
  if (val instanceof Date) {
    dt = val;
  } else {
    var s = String(val).trim();
    s = s.replace(/^\uFEFF/, "");
    if (/^W\d+$/i.test(s)) return normalizeWeek_(s);
    var weekMatch = s.match(/^Week\s+(\d+)$/i);
    if (weekMatch) return "W" + parseInt(weekMatch[1], 10);
    var brMatch = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/);
    if (brMatch) {
      var day = parseInt(brMatch[1], 10);
      var month = parseInt(brMatch[2], 10) - 1;
      var year = parseInt(brMatch[3], 10);
      if (year < 100) year += 2000;
      dt = new Date(year, month, day);
    } else {
      var shortMatch = s.match(/^(\d{1,2})-([A-Za-z]{3})$/);
      if (shortMatch) {
        var months = {jan:0,feb:1,mar:2,apr:3,may:4,jun:5,jul:6,aug:7,sep:8,oct:9,nov:10,dec:11};
        var mIdx = months[shortMatch[2].toLowerCase()];
        if (mIdx !== undefined) {
          dt = new Date(2026, mIdx, parseInt(shortMatch[1], 10));
        }
      }
      if (!dt) {
        var isoMatch = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
        if (isoMatch) {
          dt = new Date(parseInt(isoMatch[1],10), parseInt(isoMatch[2],10)-1, parseInt(isoMatch[3],10));
        } else {
          dt = new Date(s);
          if (isNaN(dt.getTime())) return s;
        }
      }
    }
  }
  var d = new Date(Date.UTC(dt.getFullYear(), dt.getMonth(), dt.getDate()));
  var dayNum = d.getUTCDay() || 7;
  d.setUTCDate(d.getUTCDate() + 4 - dayNum);
  var yearStart = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
  var weekNo = Math.ceil((((d - yearStart) / 86400000) + 1) / 7);
  return "W" + weekNo;
}

function normalizeWeek_(weekStr) {
  var m = String(weekStr).match(/W0*(\d+)/i);
  return m ? "W" + parseInt(m[1], 10) : weekStr;
}

function lerConfig_(ss) {
  var sheet = ss.getSheetByName("config");
  if (!sheet) return {};
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return {};
  var data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
  var result = {};
  for (var i = 0; i < data.length; i++) {
    var param = data[i][0];
    var valor = data[i][1];
    if (param && String(param).trim() !== "") {
      result[String(param).trim()] = String(valor).trim();
    }
  }
  return result;
}

function detectSeparator_(headerLine) {
  if (!headerLine || String(headerLine).trim() === "") return ",";
  var tabCount   = (headerLine.match(/\t/g)  || []).length;
  var commaCount = (headerLine.match(/,/g)   || []).length;
  var semiCount  = (headerLine.match(/;/g)   || []).length;
  if (semiCount >= tabCount && semiCount >= commaCount) return ";";
  if (tabCount >= commaCount) return "\t";
  return ",";
}

function fmToPercent_(val) {
  if (val === null || val === undefined || isNaN(val)) return null;
  if (val === 0) return null;
  if (val > 0 && val <= 1) return Math.round(val * 10000) / 100;
  if (val > 1 && val < 110) return Math.round(val * 100) / 100;
  if (val >= 1000 && val <= 11000) return Math.round(val) / 100;
  if (val >= 110 && val < 1000) return Math.round(val * 10) / 100;
  return null;
}

function normalizeWeekOTD_(raw) {
  if (!raw) return "";
  var s = raw.toString().trim();
  if (/^W\d+$/i.test(s)) return "W" + parseInt(s.replace(/\D/g, ""), 10);
  var match = s.match(/(?:week|semana)\s*(\d+)/i);
  if (match) return "W" + parseInt(match[1], 10);
  if (/^\d+$/.test(s)) return "W" + parseInt(s, 10);
  var nums = s.match(/\d+/);
  if (nums) return "W" + parseInt(nums[0], 10);
  return "";
}

function weightedAvg_(arr) {
  if (!arr || arr.length === 0) return null;
  var totalWeight = 0, weightedSum = 0;
  for (var i = 0; i < arr.length; i++) {
    var w = arr[i].vol || 1;
    weightedSum += arr[i].val * w;
    totalWeight += w;
  }
  return totalWeight > 0
    ? Math.round((weightedSum / totalWeight) * 100) / 100
    : null;
}

function dateToWeek_(input) {
  var d;
  if (input instanceof Date) {
    d = input;
  } else {
    var s = String(input).trim();
    if (s.indexOf("/") > -1) {
      var p = s.split("/");
      d = new Date(Number(p[2]), Number(p[1]) - 1, Number(p[0]));
    } else if (s.indexOf("-") > -1) {
      var p2 = s.split("-");
      d = new Date(Number(p2[0]), Number(p2[1]) - 1, Number(p2[2]));
    } else {
      d = new Date(s);
    }
  }
  if (!d || isNaN(d.getTime())) return "WNaN";
  var jan1 = new Date(d.getFullYear(), 0, 1);
  var days = Math.floor((d - jan1) / 86400000);
  return "W" + Math.ceil((days + jan1.getDay() + 1) / 7);
}


// ============================================================================
// FERIADOS NACIONAIS 2026 (COMPLETO)
// ============================================================================

function getHolidays2026_() {
  return [
    "2026-01-01","2026-02-16","2026-02-17","2026-04-03","2026-04-21",
    "2026-05-01","2026-06-04","2026-09-07","2026-10-12","2026-11-02",
    "2026-11-15","2026-11-20","2026-12-25"
  ];
}

function isHoliday_(d) {
  var holidays = getHolidays2026_();
  var ds = Utilities.formatDate(d, "GMT-3", "yyyy-MM-dd");
  return holidays.indexOf(ds) > -1;
}

function isBusinessDay_(d) {
  var dow = d.getDay();
  if (dow === 0 || dow === 6) return false;
  return !isHoliday_(d);
}

function getNextBusinessDay_(d) {
  var nxt = new Date(d.getTime() + 86400000);
  var tries = 0;
  while (!isBusinessDay_(nxt) && tries < 15) {
    nxt = new Date(nxt.getTime() + 86400000);
    tries++;
  }
  return nxt;
}

function hasHolidayInGap_(event202Date) {
  if (!event202Date || isNaN(event202Date.getTime())) return false;
  var nbd = getNextBusinessDay_(event202Date);
  var current = new Date(event202Date.getTime() + 86400000);
  while (current <= nbd) {
    if (isHoliday_(current)) return true;
    current = new Date(current.getTime() + 86400000);
  }
  return false;
}

function getStowD1Config_() {
  return { "ELO8": ["EVE8"], "ESE8": ["EME8", "EUM8"] };
}

function getDepD1Config_() {
  return ["EUM8", "EME8"];
}

function parseDate_(val) {
  if (!val || String(val).trim() === "") return null;
  var s = String(val).trim();
  if (s.match(/^\d{4}-\d{2}-\d{2}/)) {
    var parts = s.substring(0, 10).split("-");
    return new Date(parseInt(parts[0]), parseInt(parts[1]) - 1, parseInt(parts[2]));
  }
  if (s.match(/^\d{2}\/\d{2}\/\d{4}/)) {
    var p = s.substring(0, 10).split("/");
    return new Date(parseInt(p[2]), parseInt(p[1]) - 1, parseInt(p[0]));
  }
  var d = new Date(s);
  return isNaN(d.getTime()) ? null : d;
}

function dateOnly_(d) {
  if (!d) return null;
  return new Date(d.getFullYear(), d.getMonth(), d.getDate());
}


// ============================================================================
// NODE CONFIG (62 nodes)
// ============================================================================

function getNodeConfig_() {
  return {
    "ESE8": { regional:"DF-SUL", grupo:"Dialogo",       tipo:"SPC", owner:"Adriano",       cpt:"23:59", d1_stow:true,  d1_dep:false },
    "EIJ8": { regional:"DF-SUL", grupo:"Rede Frete",    tipo:"PN",  owner:"Adriano",       cpt:"19:00", d1_stow:false, d1_dep:false },
    "EUM8": { regional:"DF-SUL", grupo:"Rede Frete",    tipo:"PN",  owner:"Adriano",       cpt:"10:30", d1_stow:false, d1_dep:true  },
    "EVE8": { regional:"DF-SUL", grupo:"Rede Frete",    tipo:"PN",  owner:"Adriano",       cpt:"23:59", d1_stow:false, d1_dep:false },
    "EJO8": { regional:"DF-SUL", grupo:"Dialogo",       tipo:"PN",  owner:"Joao Pedro",    cpt:"22:30", d1_stow:false, d1_dep:false },
    "EMR8": { regional:"DF-SUL", grupo:"Dialogo",       tipo:"PN",  owner:"Joao Pedro",    cpt:"19:00", d1_stow:false, d1_dep:false },
    "ELO8": { regional:"DF-SUL", grupo:"Dialogo",       tipo:"SPC", owner:"Joao Pedro",    cpt:"23:59", d1_stow:false, d1_dep:false },
    "ESB8": { regional:"DF-SUL", grupo:"AMZL",          tipo:"PN",  owner:"Joao Pedro",    cpt:"23:59", d1_stow:false, d1_dep:false },
    "EGO8": { regional:"DF-SUL", grupo:"AMZL",          tipo:"PN",  owner:"Jose Procopio", cpt:"06:00", d1_stow:false, d1_dep:false },
    "ERS8": { regional:"DF-SUL", grupo:"Rede Frete",    tipo:"SPC", owner:"Kimberly",      cpt:"23:59", d1_stow:false, d1_dep:false },
    "EXS9": { regional:"DF-SUL", grupo:"AMZL_eXPT",    tipo:"PN",  owner:"Kimberly",      cpt:"18:00", d1_stow:false, d1_dep:false },
    "EME8": { regional:"DF-SUL", grupo:"Smolka",        tipo:"PN",  owner:"Vinicius",      cpt:"10:30", d1_stow:false, d1_dep:true  },
    "ECB8": { regional:"DF-SUL", grupo:"Smolka",        tipo:"SPC", owner:"Vinicius",      cpt:"22:00", d1_stow:true,  d1_dep:false },
    "ESS8": { regional:"RIMES-NE", grupo:"AMZL",          tipo:"SPC", owner:"Bruno",     cpt:"20:00", d1_stow:false, d1_dep:false },
    "ERJ1": { regional:"RIMES-NE", grupo:"DeLuna",         tipo:"SPC", owner:"Bruno",     cpt:"23:59", d1_stow:false, d1_dep:false },
    "ENF8": { regional:"RIMES-NE", grupo:"Facil Express",  tipo:"SPC", owner:"Bruno",     cpt:"20:00", d1_stow:false, d1_dep:false },
    "ERJ9": { regional:"RIMES-NE", grupo:"Favela Log",     tipo:"PN",  owner:"Bruno",     cpt:"19:00", d1_stow:false, d1_dep:false },
    "ESG8": { regional:"RIMES-NE", grupo:"Favela Log",     tipo:"PN",  owner:"Bruno",     cpt:"19:00", d1_stow:false, d1_dep:false },
    "ELZ8": { regional:"RIMES-NE", grupo:"AMZL",          tipo:"SPC", owner:"Gilberto",  cpt:"18:00", d1_stow:false, d1_dep:false },
    "EMG8": { regional:"RIMES-NE", grupo:"AMZL",          tipo:"SPC", owner:"Gilberto",  cpt:"21:00", d1_stow:false, d1_dep:false },
    "EGV8": { regional:"RIMES-NE", grupo:"Shippify",       tipo:"PN",  owner:"Gilberto",  cpt:"16:00", d1_stow:false, d1_dep:false },
    "EOU8": { regional:"RIMES-NE", grupo:"Shippify",       tipo:"PN",  owner:"Gilberto",  cpt:"20:30", d1_stow:false, d1_dep:false },
    "EPC8": { regional:"RIMES-NE", grupo:"Shippify",       tipo:"PN",  owner:"Gilberto",  cpt:"17:00", d1_stow:false, d1_dep:false },
    "EPL8": { regional:"RIMES-NE", grupo:"Shippify",       tipo:"PN",  owner:"Gilberto",  cpt:"17:00", d1_stow:false, d1_dep:false },
    "EUB8": { regional:"RIMES-NE", grupo:"Shippify",       tipo:"PN",  owner:"Gilberto",  cpt:"19:00", d1_stow:false, d1_dep:false },
    "EJF8": { regional:"RIMES-NE", grupo:"Shippify",       tipo:"SPC", owner:"Gilberto",  cpt:"19:00", d1_stow:false, d1_dep:false },
    "EPG8": { regional:"RIMES-NE", grupo:"Shippify",       tipo:"SPC", owner:"Gilberto",  cpt:"21:00", d1_stow:false, d1_dep:false },
    "EPE8": { regional:"RIMES-NE", grupo:"AMZL",          tipo:"SPC", owner:"Thiago",    cpt:"21:00", d1_stow:false, d1_dep:false },
    "EVT8": { regional:"RIMES-NE", grupo:"AMZL",          tipo:"SPC", owner:"Thiago",    cpt:"23:59", d1_stow:false, d1_dep:false },
    "ECO9": { regional:"RIMES-NE", grupo:"Facil Express",  tipo:"PN",  owner:"Thiago",    cpt:"18:00", d1_stow:false, d1_dep:false },
    "EUA8":       { regional:"SP CAP", grupo:"AMZL",       tipo:"PN",  owner:"Henrique",        cpt:"20:00", d1_stow:false, d1_dep:false },
    "ELP8":       { regional:"SP CAP", grupo:"AMZL",       tipo:"SPC", owner:"Talita",           cpt:"23:30", d1_stow:false, d1_dep:false },
    "Super ELP8": { regional:"SP CAP", grupo:"AMZL",       tipo:"SPC", owner:"Murilo",           cpt:"08:00", d1_stow:false, d1_dep:false },
    "ESD8":       { regional:"SP CAP", grupo:"Dominalog",  tipo:"SPC", owner:"Eduardo",          cpt:"16:30", d1_stow:false, d1_dep:false },
    "ESA8":       { regional:"SP CAP", grupo:"AMZL",       tipo:"SPC", owner:"Thayrini",         cpt:"23:59", d1_stow:false, d1_dep:false },
    "ELP3":       { regional:"SP CAP", grupo:"To Do Green",tipo:"PN",  owner:"Thayrini/Eduardo", cpt:"19:00", d1_stow:false, d1_dep:false },
    "ESP7": { regional:"SP INT", grupo:"AMZL_eXPT",   tipo:"PN",  owner:"Ana",     cpt:"19:00", d1_stow:false, d1_dep:false },
    "ESP9": { regional:"SP INT", grupo:"AMZL_eXPT",   tipo:"PN",  owner:"Ana",     cpt:"19:00", d1_stow:false, d1_dep:false },
    "ESP8": { regional:"SP INT", grupo:"GAFOR",        tipo:"SPC", owner:"Ana",     cpt:"02:00", d1_stow:false, d1_dep:false },
    "ELP7": { regional:"SP INT", grupo:"MORAIS",       tipo:"SPC", owner:"Ana",     cpt:"23:59", d1_stow:false, d1_dep:false },
    "ELT8": { regional:"SP INT", grupo:"To Do Green",  tipo:"PN",  owner:"Ana",     cpt:"19:00", d1_stow:false, d1_dep:false },
    "EOG8": { regional:"SP INT", grupo:"To Do Green",  tipo:"PN",  owner:"Ana",     cpt:"19:00", d1_stow:false, d1_dep:false },
    "ESP5": { regional:"SP INT", grupo:"To Do Green",  tipo:"PN",  owner:"Ana",     cpt:"19:00", d1_stow:false, d1_dep:false },
    "ESP6": { regional:"SP INT", grupo:"To Do Green",  tipo:"PN",  owner:"Ana",     cpt:"19:00", d1_stow:false, d1_dep:false },
    "EUI8": { regional:"SP INT", grupo:"To Do Green",  tipo:"PN",  owner:"Ana",     cpt:"19:00", d1_stow:false, d1_dep:false },
    "EUN8": { regional:"SP INT", grupo:"To Do Green",  tipo:"PN",  owner:"Ana",     cpt:"19:00", d1_stow:false, d1_dep:false },
    "ESJ8": { regional:"SP INT", grupo:"To Do Green",  tipo:"SPC", owner:"Ana",     cpt:"01:00", d1_stow:false, d1_dep:false },
    "ENC8": { regional:"SP INT", grupo:"To Do Green",  tipo:"PN",  owner:"Ricardo", cpt:"19:00", d1_stow:false, d1_dep:false },
    "EOS8": { regional:"SP INT", grupo:"To Do Green",  tipo:"PN",  owner:"Ricardo", cpt:"19:00", d1_stow:false, d1_dep:false },
    "EPF8": { regional:"SP INT", grupo:"To Do Green",  tipo:"PN",  owner:"Ricardo", cpt:"19:00", d1_stow:false, d1_dep:false },
    "ERN8": { regional:"SP INT", grupo:"To Do Green",  tipo:"PN",  owner:"Ricardo", cpt:"20:00", d1_stow:false, d1_dep:false },
    "ESC8": { regional:"SP INT", grupo:"To Do Green",  tipo:"PN",  owner:"Ricardo", cpt:"19:00", d1_stow:false, d1_dep:false },
    "ERP8": { regional:"SP INT", grupo:"To Do Green",  tipo:"SPC", owner:"Ricardo", cpt:"00:29", d1_stow:false, d1_dep:false },
    "EIO8": { regional:"SP INT", grupo:"To Do Green",  tipo:"PN",  owner:"Sidny",   cpt:"18:00", d1_stow:false, d1_dep:false },
    "ELN8": { regional:"SP INT", grupo:"To Do Green",  tipo:"PN",  owner:"Sidny",   cpt:"18:30", d1_stow:false, d1_dep:false },
    "EML8": { regional:"SP INT", grupo:"To Do Green",  tipo:"PN",  owner:"Sidny",   cpt:"18:00", d1_stow:false, d1_dep:false },
    "EPP8": { regional:"SP INT", grupo:"To Do Green",  tipo:"PN",  owner:"Sidny",   cpt:"15:30", d1_stow:false, d1_dep:false },
    "ETB8": { regional:"SP INT", grupo:"To Do Green",  tipo:"PN",  owner:"Sidny",   cpt:"17:30", d1_stow:false, d1_dep:false },
    "ETU8": { regional:"SP INT", grupo:"To Do Green",  tipo:"PN",  owner:"Sidny",   cpt:"18:00", d1_stow:false, d1_dep:false },
    "ESV8": { regional:"SP INT", grupo:"Total Express", tipo:"PN", owner:"Sidny",   cpt:"20:00", d1_stow:false, d1_dep:false },
    "EUR8": { regional:"SP INT", grupo:"To Do Green",  tipo:"SPC", owner:"Sidny",   cpt:"01:00", d1_stow:false, d1_dep:false }
  };
}


// ============================================================================
// PROCESSAMENTO DOS CSVs
// ============================================================================

function processarCSV_FM(csvContent) {
  if (!csvContent || csvContent.trim() === "") return { weekly: [], daily: [], isDiario: false };
  csvContent = csvContent.replace(/^\uFEFF/, "");
  csvContent = csvContent.replace(/\r\n/g, "\n").replace(/\r/g, "\n");
  var lines = csvContent.split("\n");
  if (lines.length < 2) return { weekly: [], daily: [], isDiario: false };

  var sep = detectSeparator_(lines[0]);
  Logger.log("FM - Separador: '" + sep + "'");
  var headers = lines[0].split(sep);
  for (var h = 0; h < headers.length; h++) {
    headers[h] = headers[h].replace(/"/g, "").replace(/^\uFEFF/, "").trim();
  }

  var colPeriod   = findCol_(headers, ["Period","period","Periodo"," Period"]);
  var colNode     = findCol_(headers, ["Node","node"]);
  var colNodeDef  = findCol_(headers, ["Node definition","node_definition","Node Definition"]);
  var colRegional = findCol_(headers, ["Regional","regional"]);
  var colRecv     = findCol_(headers, ["2. Receive success %","Receive success"]);
  var colStow     = findCol_(headers, ["4. Stow success %","Stow success"]);
  var colDepart   = findCol_(headers, ["5. Depart success %","Depart success"]);
  var colBacklog  = findCol_(headers, ["6. FM Backlog delay","FM Backlog delay"]);
  var colBackAcc  = findCol_(headers, ["7. FM Backlog delay (acc)","FM Backlog delay (acc)"]);
  var colPPP      = findCol_(headers, ["9. Packages per pallet","Packages per pallet"]);
  var colMissort  = findCol_(headers, ["10. Missort (DPMO)","Missort"]);
  var colWN       = findCol_(headers, ["8. Wrong node %","8. Wrong node","Wrong node"]);
  var colVol      = findCol_(headers, ["3. Processed","Processed"]);
  var colPickedUp = findCol_(headers, ["1. Picked up","Picked up"]);
  var colRecvAdj  = findCol_(headers, ["2.1. Receive success adj %","Receive success adj"]);
  var colStowAdj  = findCol_(headers, ["4.1. Stow succcess adj %","4.1. Stow success adj %","Stow success adj"]);
  var colOTD      = findCol_(headers, ["13. OTD %","OTD %","OTD"]);

  if (colPPP === -1 && headers.length > 22) colPPP = 22;
  if (colVol === -1) colVol = colPickedUp;

  Logger.log("FM cols: Period=" + colPeriod + " Node=" + colNode +
             " Vol=" + colVol + " PPP=" + colPPP);

  if (colPeriod === -1 || colNode === -1) {
    Logger.log("ERRO: colunas obrigatorias nao encontradas");
    return { weekly: [], daily: [], isDiario: false };
  }

  var isDiario = true;
  var weeklyMap = {};
  var dailyOutput = [];
  var lastPPP = {};

  for (var i = 1; i < lines.length; i++) {
    if (lines[i].trim() === "") continue;
    var cols = lines[i].split(sep);
    for (var c = 0; c < cols.length; c++) {
      cols[c] = cols[c].replace(/"/g, "").trim();
    }

    var period   = cols[colPeriod] || "";
    var node     = cols[colNode] || "";
    var nodeDef  = colNodeDef !== -1 ? (cols[colNodeDef] || "") : "";
    var regional = colRegional !== -1 ? (cols[colRegional] || "") : "";

    if (!node || !period) continue;

    var recvRaw    = colRecv    !== -1 ? toNumBR_(cols[colRecv])    : null;
    var recvAdjRaw = colRecvAdj !== -1 ? toNumBR_(cols[colRecvAdj]) : null;
    var stowRaw    = colStow    !== -1 ? toNumBR_(cols[colStow])    : null;
    var stowAdjRaw = colStowAdj !== -1 ? toNumBR_(cols[colStowAdj]) : null;
    var depRaw     = colDepart  !== -1 ? toNumBR_(cols[colDepart])  : null;
    var backlog    = colBacklog !== -1 ? toNumBR_(cols[colBacklog])  : null;
    var backAcc    = colBackAcc !== -1 ? toNumBR_(cols[colBackAcc])  : null;
    var ppp        = colPPP     !== -1 ? toNumBR_(cols[colPPP])     : null;
    var missort    = colMissort !== -1 ? toNumBR_(cols[colMissort])  : null;
    var wn         = colWN      !== -1 ? toNumBR_(cols[colWN])       : null;
    var volume     = colVol     !== -1 ? toNumBR_(cols[colVol])      : null;
    var pickedUp   = colPickedUp !== -1 ? toNumBR_(cols[colPickedUp]) : null;
    var otdRaw     = colOTD    !== -1 ? toNumBR_(cols[colOTD])       : null;

    // Receive: usa adj se disponivel e nao-placeholder (raw > 1); senao original
    var recv = (recvAdjRaw !== null && recvAdjRaw > 1)
              ? fmToPercent_(recvAdjRaw)
              : fmToPercent_(recvRaw);
    // Stow: mesma logica
    var stow = (stowAdjRaw !== null && stowAdjRaw > 1)
              ? fmToPercent_(stowAdjRaw)
              : fmToPercent_(stowRaw);
    var depart = fmToPercent_(depRaw);
    // OTD: raw <= 1 e placeholder -> null
    var otd = (otdRaw !== null && otdRaw > 1) ? fmToPercent_(otdRaw) : null;

    if (ppp !== null && ppp > 0) {
      while (ppp > 500) { ppp = ppp / 10; }
      lastPPP[node] = ppp;
    } else {
      ppp = null;
    }

    var semana = toWeekStr_(period);

    if (isDiario) {
      dailyOutput.push({
        data: period, node: node, nodeDef: nodeDef, regional: regional,
        recv: recv, stow: stow, depart: depart, backlog: backlog,
        ppp: ppp, volume: volume
      });
    }

    var wKey = semana + "|" + node;
    if (!weeklyMap[wKey]) {
      weeklyMap[wKey] = {
        semana: semana, node: node, nodeDef: nodeDef, regional: regional,
        sumVol: 0, countVol: 0,
        sumRecvW: 0, sumVolRecv: 0,
        sumStowW: 0, sumVolStow: 0,
        sumDepW: 0, sumVolDep: 0,
        sumOtdW: 0, sumVolOtd: 0,
        sumBack: 0, countBack: 0,
        sumBackAcc: 0, countBackAcc: 0,
        sumPPP: 0, countPPP: 0,
        sumMissort: 0, countMissort: 0,
        sumWN: 0, countWN: 0
      };
    }
    var entry = weeklyMap[wKey];

    if (volume !== null) { entry.sumVol += volume; entry.countVol++; }

    var peso = pickedUp || volume || 0;

    if (recv !== null && peso > 0) {
      entry.sumRecvW += recv * peso;
      entry.sumVolRecv += peso;
    }
    if (stow !== null && peso > 0) {
      entry.sumStowW += stow * peso;
      entry.sumVolStow += peso;
    }
    if (depart !== null && peso > 0) {
      entry.sumDepW += depart * peso;
      entry.sumVolDep += peso;
    }
    if (otd !== null && peso > 0) {
      entry.sumOtdW  += otd * peso;
      entry.sumVolOtd += peso;
    }

    if (backlog !== null) { entry.sumBack += backlog; entry.countBack++; }
    if (backAcc !== null) { entry.sumBackAcc += backAcc; entry.countBackAcc++; }
    if (ppp !== null) { entry.sumPPP += ppp; entry.countPPP++; }
    if (missort !== null) { entry.sumMissort += missort; entry.countMissort++; }
    if (wn !== null) { entry.sumWN += wn; entry.countWN++; }
  }

  var weeklyOutput = [];
  var keys = Object.keys(weeklyMap);
  for (var k = 0; k < keys.length; k++) {
    var e = weeklyMap[keys[k]];
    weeklyOutput.push({
      semana: e.semana, node: e.node, nodeDef: e.nodeDef, regional: e.regional,
      volume: e.sumVol,
      receive: e.sumVolRecv > 0 ? Math.round(e.sumRecvW / e.sumVolRecv * 100) / 100 : null,
      stow: e.sumVolStow > 0 ? Math.round(e.sumStowW / e.sumVolStow * 100) / 100 : null,
      depart: e.sumVolDep  > 0 ? Math.round(e.sumDepW  / e.sumVolDep  * 100) / 100 : null,
      otd:    e.sumVolOtd  > 0 ? Math.round(e.sumOtdW  / e.sumVolOtd  * 100) / 100 : null,
      backlog: e.countBack > 0 ? e.sumBack / e.countBack : null,
      backlogAcc: e.countBackAcc > 0 ? e.sumBackAcc / e.countBackAcc : null,
      ppp: e.countPPP > 0 ? Math.round(e.sumPPP / e.countPPP * 10) / 10 : null,
      missort: e.countMissort > 0 ? e.sumMissort / e.countMissort : null,
      wrongNode: e.countWN > 0 ? e.sumWN / e.countWN : null
    });
  }

  Logger.log("FM Metrics - Weekly: " + weeklyOutput.length +
             " | Daily: " + dailyOutput.length);
  return { weekly: weeklyOutput, daily: dailyOutput, isDiario: isDiario };
}


function processarExcel_OTD(csvContent) {
  if (!csvContent || csvContent.trim() === "") return [];
  csvContent = csvContent.replace(/\r\n/g, "\n").replace(/\r/g, "\n");
  csvContent = csvContent.replace(/^\uFEFF/, "");
  var lines = csvContent.split("\n");
  if (lines.length < 2) return [];

  var sep = detectSeparator_(lines[0]);
  Logger.log("OTD/CPT - Separador: '" + sep + "'");
  var headers = lines[0].split(sep);
  for (var h = 0; h < headers.length; h++) {
    headers[h] = headers[h].replace(/"/g, "").replace(/^\uFEFF/, "").trim();
  }

  var colWeek     = findCol_(headers, ["Week","week","Semana","Period"]);
  var colNode     = findCol_(headers, ["Node","node","Station"]);
  var colCPT      = findCol_(headers, ["CPT Success %","CPT Success","CPT_Success","cpt_success"]);
  var colOTD      = findCol_(headers, ["OTD (%)","OTD","OTD Success","otd_success","OTD_Success"]);
  var colForecast = findCol_(headers, ["Forecast","forecast","Volume Forecast"]);
  var colTPH      = findCol_(headers, ["TPH","tph"]);
  var colHC       = findCol_(headers, ["HC","hc","Headcount"]);
  var colHoras    = findCol_(headers, ["Horas Trabalhadas","horas_trabalhadas","Horas"]);
  var colVolProc  = findCol_(headers, ["Volume Processado","volume_processado"]);
  var colDesvio   = findCol_(headers, ["Desvio Forecast x Real","Desvio","desvio"]);

  if (colCPT === -1) {
    for (var h = 0; h < headers.length; h++) {
      var hLow = headers[h].toLowerCase();
      if (hLow.indexOf("cpt") > -1 && hLow.indexOf("success") > -1) {
        colCPT = h; break;
      }
    }
  }

  Logger.log("OTD cols: Week=" + colWeek + " Node=" + colNode +
             " OTD=" + colOTD + " CPT=" + colCPT + " TPH=" + colTPH +
             " HC=" + colHC + " Horas=" + colHoras +
             " VolProc=" + colVolProc + " Desvio=" + colDesvio);

  if (colWeek === -1 || colNode === -1) {
    Logger.log("ERRO OTD: colunas obrigatorias nao encontradas (Week=" +
               colWeek + " Node=" + colNode + ")");
    return [];
  }

  var output = [];
  for (var i = 1; i < lines.length; i++) {
    if (lines[i].trim() === "") continue;
    var cols = lines[i].split(sep);
    for (var c = 0; c < cols.length; c++) {
      cols[c] = cols[c].replace(/"/g, "").trim();
    }

    var weekRaw = cols[colWeek] || "";
    var node    = cols[colNode] || "";
    if (!weekRaw || !node) continue;

    var semana = normalizeWeekOTD_(weekRaw);
    if (!semana) continue;

    var otd      = colOTD !== -1 ? toNumBR_(cols[colOTD]) : null;
    var cpt      = colCPT !== -1 ? toNumBR_(cols[colCPT]) : null;
    var forecast = colForecast !== -1 ? toNumBR_(cols[colForecast]) : null;
    var tph      = colTPH !== -1 ? toNumBR_(cols[colTPH]) : null;
    var hc       = colHC !== -1 ? toNumBR_(cols[colHC]) : null;
    var horas    = colHoras !== -1 ? toNumBR_(cols[colHoras]) : null;
    var volProc  = colVolProc !== -1 ? toNumBR_(cols[colVolProc]) : null;
    var desvio   = colDesvio !== -1 ? toNumBR_(cols[colDesvio]) : null;

    output.push({
      semana: semana, node: node, otd: otd, cpt: cpt,
      forecast: forecast, desvio: desvio, tph: tph,
      hc: hc, horas: horas, volProc: volProc
    });
  }

  Logger.log("OTD/CPT - Registros: " + output.length);
  return output;
}


function processarCSV_Shipment(csvContent) {
  if (!csvContent || csvContent.trim() === "") return { weekly: [], daily: [] };
  csvContent = csvContent.replace(/\r\n/g, "\n").replace(/\r/g, "\n");
  var lines = csvContent.split("\n");
  if (lines.length < 2) return { weekly: [], daily: [] };

  var sep = detectSeparator_(lines[0]);
  Logger.log("Shipment - Separador: '" + sep + "'");
  var headers = lines[0].split(sep);
  for (var h = 0; h < headers.length; h++) {
    headers[h] = headers[h].replace(/"/g, "").trim();
  }

  var colDate     = findCol_(headers, ["Reference date","reference_date","Date"]);
  var colNode     = findCol_(headers, ["Node","node","Planned node"]);
  var colKPI      = findCol_(headers, ["KPI name","kpi_name","KPI"]);
  var colSuccess  = findCol_(headers, ["KPI success","kpi_success","Success"]);
  var colOrig     = findCol_(headers, ["Planned lane","planned_lane","Origin"]);
  var colEv201    = 27;
  var colEv202    = 28;
  var colSched    = 20;
  var colNextMileCode = findCol_(headers, ["Next mile ev. code","next_mile_ev_code"]);

  if (colDate === -1 || colKPI === -1) {
    Logger.log("ERRO Shipment: colunas obrigatorias nao encontradas");
    return { weekly: [], daily: [] };
  }

  var stowD1Cfg  = getStowD1Config_();
  var depD1Cfg   = getDepD1Config_();
  var nodeConfig = getNodeConfig_();

  var weeklyMap = {};
  var dailyMap  = {};
  var totalProcessed = 0;
  var stowD1Hits = 0;
  var depD1Hits  = 0;

  for (var i = 1; i < lines.length; i++) {
    if (lines[i].trim() === "") continue;
    var cols = lines[i].split(sep);
    for (var c = 0; c < cols.length; c++) {
      cols[c] = cols[c].replace(/"/g, "").trim();
    }

    var refDate = cols[colDate] || "";
    var node    = colNode !== -1 ? (cols[colNode] || "") : "";
    var kpi     = cols[colKPI] || "";
    var success = colSuccess !== -1 ? (cols[colSuccess] || "") : "";
    var origin  = colOrig !== -1 ? (cols[colOrig] || "") : "";

    if (!refDate || !kpi) continue;
    totalProcessed++;

    var dt = parseDate_(refDate);
    var semana = "";
    if (dt) { semana = toWeekStr_(dt); }
    if (!semana) continue;

    var kpiLower = kpi.toLowerCase();
    var isRecv = kpiLower.indexOf("receive") > -1;
    var isStow = kpiLower.indexOf("stow") > -1;
    var isDep  = kpiLower.indexOf("depart") > -1;
    if (!isRecv && !isStow && !isDep) continue;

    var isMiss = (success.toLowerCase() !== "true" &&
                  success !== "1" && success.toLowerCase() !== "success");

    var cfg   = nodeConfig[node];
    var grupo = cfg ? cfg.grupo : "";
    var tipo  = cfg ? cfg.tipo  : "";

    var wKey = semana + "|" + node;
    if (!weeklyMap[wKey]) {
      weeklyMap[wKey] = {
        semana: semana, node: node, grupo: grupo, tipo: tipo,
        recv_total: 0, recv_miss: 0,
        stow_total: 0, stow_miss: 0, stow_d1_hit: 0,
        dep_total: 0,  dep_miss: 0,  dep_d1_hit: 0
      };
    }
    var entry = weeklyMap[wKey];

    var dKey = refDate + "|" + node;
    if (!dailyMap[dKey]) {
      dailyMap[dKey] = {
        data: refDate, node: node, grupo: grupo, tipo: tipo,
        miss_real: 0, recv_miss: 0, stow_miss: 0, dep_miss: 0,
        d1_corrigido: 0, total: 0
      };
    }
    var dEntry = dailyMap[dKey];
    dEntry.total++;

    if (isRecv) {
      entry.recv_total++;
      if (isMiss) {
        entry.recv_miss++;
        dEntry.recv_miss++;
        dEntry.miss_real++;
      }
    } else if (isStow) {
      entry.stow_total++;
      if (isMiss) {
        var stowCorrected = false;
        var d1Origins = stowD1Cfg[node];
        if (d1Origins && d1Origins.indexOf(origin) > -1) {
          var ev202 = parseDate_(cols[colEv202]);
          if (ev202 && hasHolidayInGap_(ev202)) {
            var ev201 = parseDate_(cols[colEv201]);
            var prazo = getNextBusinessDay_(ev202);
            var hasEv201Code = true;
            if (colNextMileCode !== -1) {
              var evCode = (cols[colNextMileCode] || "").toUpperCase();
              hasEv201Code = evCode.indexOf("201") > -1;
            }
            if (ev201 && hasEv201Code) {
              var ev201d = dateOnly_(ev201);
              var prazod = dateOnly_(prazo);
              if (ev201d && prazod && ev201d.getTime() <= prazod.getTime()) {
                stowCorrected = true;
                entry.stow_d1_hit++;
                stowD1Hits++;
                dEntry.d1_corrigido++;
              }
            }
          }
        }
        if (!stowCorrected) {
          entry.stow_miss++;
          dEntry.stow_miss++;
          dEntry.miss_real++;
        }
      }
    } else if (isDep) {
      entry.dep_total++;
      if (isMiss) {
        var depCorrected = false;
        if (depD1Cfg.indexOf(node) > -1) {
          var ev201dep = parseDate_(cols[colEv201]);
          var ev202dep = parseDate_(cols[colEv202]);
          if (ev201dep && ev202dep) {
            var prazoDep = getNextBusinessDay_(ev201dep);
            var ev202d   = dateOnly_(ev202dep);
            var prazoDepD = dateOnly_(prazoDep);
            if (ev202d && prazoDepD && ev202d.getTime() <= prazoDepD.getTime()) {
              depCorrected = true;
              entry.dep_d1_hit++;
              depD1Hits++;
              dEntry.d1_corrigido++;
            }
          }
        }
        if (!depCorrected) {
          entry.dep_miss++;
          dEntry.dep_miss++;
          dEntry.miss_real++;
        }
      }
    }
  }

  Logger.log("Shipment processado: " + totalProcessed + " registros, " +
             stowD1Hits + " stow D+1 HITs, " + depD1Hits + " dep D+1 HITs");

  var weeklyOutput = [];
  var wKeys = Object.keys(weeklyMap);
  for (var k = 0; k < wKeys.length; k++) {
    var e = weeklyMap[wKeys[k]];
    weeklyOutput.push({
      semana: e.semana, node: e.node, grupo: e.grupo, tipo: e.tipo,
      miss_real: e.recv_miss + e.stow_miss + e.dep_miss,
      miss_d1_corrigido: e.stow_d1_hit + e.dep_d1_hit,
      recv_real: e.recv_miss, recv_corrigido: 0,
      stow_real: e.stow_miss, stow_corrigido: e.stow_d1_hit,
      dep_real:  e.dep_miss,  dep_corrigido:  e.dep_d1_hit
    });
  }

  var dailyOutput = [];
  var dKeys = Object.keys(dailyMap);
  for (var k = 0; k < dKeys.length; k++) {
    dailyOutput.push(dailyMap[dKeys[k]]);
  }

  Logger.log("Shipment - Weekly: " + weeklyOutput.length +
             " | Daily: " + dailyOutput.length);
  return { weekly: weeklyOutput, daily: dailyOutput };
}


function processarCSV_Regional(csvContent) {
  var LF = String.fromCharCode(10);
  var CR = String.fromCharCode(13);
  var lines = csvContent.split(CR).join("").split(LF);
  var validLines = [];
  for (var x = 0; x < lines.length; x++) {
    if (lines[x].trim() !== "") validLines.push(lines[x]);
  }
  lines = validLines;

  var headerLine = lines.shift();
  var sep = headerLine.indexOf(";") > -1 ? ";" : ",";
  lines.unshift(headerLine);

  Logger.log("REGIONAL CSV - Separador: '" + sep + "' | Linhas: " + lines.length);
  if (lines.length < 2) return [];

  var COL_DATE = 0, COL_REG = 1, COL_VOL = 2;
  var COL_RECV = 4, COL_STOW = 9, COL_DEP = 11, COL_BACKLOG = 13;

  var agg = {};
  var skipped = 0;

  for (var i = 1; i < lines.length; i++) {
    var cols = lines[i].split(sep);
    if (cols.length < 12) continue;

    var dateStr  = (cols[COL_DATE] || "").trim().replace(/^\uFEFF/, "");
    var regional = (cols[COL_REG]  || "").trim();
    if (!dateStr || !regional) continue;

    var parts = dateStr.split("/");
    if (parts.length !== 3) { skipped++; continue; }
    var dia = parseInt(parts[0], 10);
    var mes = parseInt(parts[1], 10);
    var ano = parseInt(parts[2], 10);
    if (isNaN(dia) || isNaN(mes) || isNaN(ano)) { skipped++; continue; }

    var d = new Date(ano, mes - 1, dia);
    if (isNaN(d.getTime())) { skipped++; continue; }

    var jan1 = new Date(d.getFullYear(), 0, 1);
    var days = Math.floor((d - jan1) / 86400000);
    var sem  = "W" + Math.ceil((days + jan1.getDay() + 1) / 7);

    var key = regional + "|" + sem;
    if (!agg[key]) {
      agg[key] = { regional: regional, semana: sem,
        volume: 0, recvSum: 0, recvCount: 0,
        stowSum: 0, stowCount: 0, depSum: 0, depCount: 0, backlog: 0 };
    }

    var vol = parseInt(cols[COL_VOL], 10) || 0;
    var recv = fmToPercent_(toNumBR_(cols[COL_RECV]));
    var stow = fmToPercent_(toNumBR_(cols[COL_STOW]));
    var dep  = fmToPercent_(toNumBR_(cols[COL_DEP]));
    var bkl  = parseInt(cols[COL_BACKLOG], 10) || 0;

    agg[key].volume += vol;
    agg[key].backlog += bkl;
    if (recv !== null && recv > 0) { agg[key].recvSum += recv; agg[key].recvCount++; }
    if (stow !== null && stow > 0) { agg[key].stowSum += stow; agg[key].stowCount++; }
    if (dep  !== null && dep  > 0) { agg[key].depSum  += dep;  agg[key].depCount++;  }
  }

  var results = [];
  var keys = Object.keys(agg);
  for (var k = 0; k < keys.length; k++) {
    var r = agg[keys[k]];
    if (r.regional === "OTHERS") continue;
    results.push({
      semana: r.semana, regional: r.regional, volume: r.volume,
      receive: r.recvCount > 0 ? Math.round(r.recvSum / r.recvCount * 100) / 100 : 0,
      stow:    r.stowCount > 0 ? Math.round(r.stowSum / r.stowCount * 100) / 100 : 0,
      depart:  r.depCount  > 0 ? Math.round(r.depSum  / r.depCount  * 100) / 100 : 0,
      ppp: 0, backlog: r.backlog
    });
  }

  var fmAgg = {};
  for (var k2 = 0; k2 < keys.length; k2++) {
    var r2 = agg[keys[k2]];
    if (!fmAgg[r2.semana]) {
      fmAgg[r2.semana] = { volume: 0, recvSum: 0, recvCount: 0,
        stowSum: 0, stowCount: 0, depSum: 0, depCount: 0, backlog: 0 };
    }
    fmAgg[r2.semana].volume    += r2.volume;
    fmAgg[r2.semana].backlog   += r2.backlog;
    fmAgg[r2.semana].recvSum   += r2.recvSum;
    fmAgg[r2.semana].recvCount += r2.recvCount;
    fmAgg[r2.semana].stowSum   += r2.stowSum;
    fmAgg[r2.semana].stowCount += r2.stowCount;
    fmAgg[r2.semana].depSum    += r2.depSum;
    fmAgg[r2.semana].depCount  += r2.depCount;
  }
  var fmSems = Object.keys(fmAgg);
  for (var f = 0; f < fmSems.length; f++) {
    var fm = fmAgg[fmSems[f]];
    results.push({
      semana: fmSems[f], regional: "First Mile", volume: fm.volume,
      receive: fm.recvCount > 0 ? Math.round(fm.recvSum / fm.recvCount * 100) / 100 : 0,
      stow:    fm.stowCount > 0 ? Math.round(fm.stowSum / fm.stowCount * 100) / 100 : 0,
      depart:  fm.depCount  > 0 ? Math.round(fm.depSum  / fm.depCount  * 100) / 100 : 0,
      ppp: 0, backlog: fm.backlog
    });
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var pppSheet = ss.getSheetByName("regional_ppp_comparativo");
  if (pppSheet) {
    var pppData = pppSheet.getDataRange().getValues();
    var pppMap  = {};
    for (var p = 1; p < pppData.length; p++) {
      var pSem = String(pppData[p][0]).trim();
      var pReg = String(pppData[p][1]).trim();
      var pVal = pppData[p][2];
      if (typeof pVal === "string") { pVal = parseFloat(pVal.replace(",", ".")) || 0; }
      if (pSem && pReg) pppMap[pSem + "|" + pReg] = pVal;
    }
    for (var r4 = 0; r4 < results.length; r4++) {
      var pppKey = results[r4].semana + "|" + results[r4].regional;
      if (pppMap[pppKey]) { results[r4].ppp = Math.round(pppMap[pppKey] * 100) / 100; }
    }
    for (var r5 = 0; r5 < results.length; r5++) {
      if (results[r5].regional === "First Mile") {
        var fmSem2 = results[r5].semana;
        var pppTotal = 0, pppN = 0;
        for (var r6 = 0; r6 < results.length; r6++) {
          if (results[r6].semana === fmSem2 && results[r6].regional !== "First Mile" && results[r6].ppp > 0) {
            pppTotal += results[r6].ppp; pppN++;
          }
        }
        results[r5].ppp = pppN > 0 ? Math.round(pppTotal / pppN * 100) / 100 : 0;
      }
    }
  }

  Logger.log("REGIONAL CSV - Output: " + results.length + " registros | Skipped: " + skipped);
  return results;
}


// ============================================================================
// IMPORTAR DADOS
// ============================================================================

function importarDados() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var pastas = DriveApp.getFoldersByName("_automacao_wbr");
  if (!pastas.hasNext()) { Logger.log("ERRO: Pasta _automacao_wbr nao encontrada"); return; }
  var pasta = pastas.next();

  var fmContent  = ""; var otdContent = ""; var shipContent = ""; var regContent = "";
  var fmFiles  = pasta.getFilesByName("fm_metrics.csv");  if (fmFiles.hasNext())  fmContent  = fmFiles.next().getBlob().getDataAsString("UTF-8");
  var otdFiles = pasta.getFilesByName("otd_cpt.csv");     if (otdFiles.hasNext()) otdContent = otdFiles.next().getBlob().getDataAsString("UTF-8");
  var shipFiles= pasta.getFilesByName("shipment.csv");    if (shipFiles.hasNext())shipContent= shipFiles.next().getBlob().getDataAsString("UTF-8");
  var regFiles = pasta.getFilesByName("REGIONAL.csv");
  if (regFiles.hasNext()) { regContent = regFiles.next().getBlob().getDataAsString("UTF-8"); Logger.log("REGIONAL.csv: " + regContent.length + " chars"); }
  else { Logger.log("REGIONAL.csv: nao encontrado"); }

  var dadosFM       = processarCSV_FM(fmContent);
  var dadosOTD      = processarExcel_OTD(otdContent);
  var dadosShipment = processarCSV_Shipment(shipContent);
  var dadosRegional = processarCSV_Regional(regContent);

  preencherTodasAbas(ss, dadosFM, dadosOTD, dadosShipment, dadosRegional);
  Logger.log("importarDados concluido!");
}


// ============================================================================
// PREENCHER TODAS AS ABAS
// ============================================================================

function preencherTodasAbas(ss, dadosFM, dadosOTD, dadosShipment, dadosRegional) {
  var nodeConfig = getNodeConfig_();

  function sortSemanas(a, b) {
    var na = parseInt(a.replace(/\D/g, ""), 10);
    var nb = parseInt(b.replace(/\D/g, ""), 10);
    return na - nb;
  }

  // --- REGIONAL 6W ---
  var regSheet = ss.getSheetByName("regional_6w");
  if (regSheet && dadosRegional && dadosRegional.length > 0) {
    regSheet.getRange(2, 1, regSheet.getMaxRows() - 1, regSheet.getMaxColumns()).clearContent();
    var regRows = [];
    for (var r = 0; r < dadosRegional.length; r++) {
      var reg = dadosRegional[r];
      regRows.push([reg.semana, reg.regional, reg.volume, reg.receive, reg.stow, reg.depart, reg.backlog, 0, reg.ppp, 0]);
    }
    if (regRows.length > 0) { regSheet.getRange(2, 1, regRows.length, 10).setValues(regRows); }
  }

  // --- NODES 6W ---
  var nodesSheet = ss.getSheetByName("nodes_6w");
  if (nodesSheet && dadosFM.weekly.length > 0) {
    var nodeRows = [];
    for (var i = 0; i < dadosFM.weekly.length; i++) {
      var r = dadosFM.weekly[i];
      if (!nodeConfig[r.node]) continue;
      var cfg = nodeConfig[r.node] || {};
      nodeRows.push([r.semana, r.node, r.regional || "DF-SUL", cfg.grupo || "", cfg.tipo || "",
        r.volume || 0, r.receive || 0, r.stow || 0, r.depart || 0,
        r.backlog || 0, r.backlogAcc || 0, r.ppp || 0, r.missort || 0, r.wrongNode || 0]);
    }
    nodeRows.sort(function(a, b) { return sortSemanas(a[0], b[0]); });
    if (nodeRows.length > 0) {
      var lastRowNodes = nodesSheet.getLastRow();
      if (lastRowNodes > 1) { nodesSheet.getRange(2, 1, lastRowNodes - 1, 14).clearContent(); }
      nodesSheet.getRange(2, 1, nodeRows.length, 14).setValues(nodeRows);
    }
  }

  // --- NODES W1 ---
  var w1Sheet = ss.getSheetByName("nodes_w1");
  if (w1Sheet && dadosFM.weekly.length > 0) {
    var allWeeks = [];
    for (var i = 0; i < dadosFM.weekly.length; i++) {
      if (allWeeks.indexOf(dadosFM.weekly[i].semana) === -1) allWeeks.push(dadosFM.weekly[i].semana);
    }
    allWeeks.sort(sortSemanas);
    var lastWeek = allWeeks[allWeeks.length - 1];
    var w1Rows = [];
    for (var i = 0; i < dadosFM.weekly.length; i++) {
      var r = dadosFM.weekly[i];
      if (r.semana !== lastWeek) continue;
      if (!nodeConfig[r.node]) continue;
      var cfg = nodeConfig[r.node] || {};
      w1Rows.push([r.semana, r.node, r.regional || "DF-SUL", cfg.grupo || "", cfg.tipo || "",
        r.volume || 0, r.receive || 0, r.stow || 0, r.depart || 0, r.backlog || 0, r.ppp || 0]);
    }
    if (w1Rows.length > 0) {
      var lastRowW1 = w1Sheet.getLastRow();
      if (lastRowW1 > 1) { w1Sheet.getRange(2, 1, lastRowW1 - 1, 11).clearContent(); }
      w1Sheet.getRange(2, 1, w1Rows.length, 11).setValues(w1Rows);
    }
  }

  // --- REGIONAL PPP COMPARATIVO ---
  var pppSheet = ss.getSheetByName("regional_ppp_comparativo");
  if (pppSheet && dadosFM.weekly.length > 0) {
    var pppMap = {};
    for (var i = 0; i < dadosFM.weekly.length; i++) {
      var r = dadosFM.weekly[i];
      if (r.ppp === null || r.ppp === 0) continue;
      if (!nodeConfig[r.node]) continue;
      var pk = r.semana + "|" + (r.regional || "DF-SUL");
      if (!pppMap[pk]) { pppMap[pk] = { semana: r.semana, regional: r.regional || "DF-SUL", sum: 0, count: 0 }; }
      pppMap[pk].sum += r.ppp; pppMap[pk].count++;
    }
    var pppRows = [];
    var pppKeys = Object.keys(pppMap);
    pppKeys.sort(function(a, b) { return sortSemanas(a.split("|")[0], b.split("|")[0]); });
    for (var k = 0; k < pppKeys.length; k++) {
      var e = pppMap[pppKeys[k]];
      pppRows.push([e.semana, e.regional, e.count > 0 ? e.sum / e.count : 0]);
    }
    if (pppRows.length > 0) {
      var lastRowPPP = pppSheet.getLastRow();
      if (lastRowPPP > 1) { pppSheet.getRange(2, 1, lastRowPPP - 1, 3).clearContent(); }
      pppSheet.getRange(2, 1, pppRows.length, 3).setValues(pppRows);
    }
  }

  // --- WBR EXTRAS ---
  var extrasSheet = ss.getSheetByName("wbr_extras");
  if (extrasSheet && dadosOTD.length > 0) {
    var exMap = {};
    for (var i = 0; i < dadosOTD.length; i++) {
      var r = dadosOTD[i];
      if (!r.semana || !r.node) continue;
      var eKey = r.semana + "|" + r.node;
      if (!exMap[eKey]) {
        exMap[eKey] = { semana: r.semana, node: r.node,
          sumCptVol: 0, sumOtdVol: 0, sumVol: 0, sumForecast: 0, sumDesvioVol: 0,
          sumHC: 0, cHC: 0, sumHoras: 0, sumTPH: 0, cTPH: 0, sumCpt: 0, cCpt: 0, sumOtd: 0, cOtd: 0 };
      }
      var e = exMap[eKey];
      var vol = (r.volProc && r.volProc > 0) ? r.volProc : 0;
      if (r.cpt !== null && r.cpt > 0) { if (vol > 0) { e.sumCptVol += r.cpt * vol; } e.sumCpt += r.cpt; e.cCpt++; }
      if (r.otd !== null && r.otd > 0) { if (vol > 0) { e.sumOtdVol += r.otd * vol; } e.sumOtd += r.otd; e.cOtd++; }
      e.sumVol += vol;
      if (r.forecast !== null && r.forecast > 0) e.sumForecast += r.forecast;
      if (r.desvio !== null && vol > 0) e.sumDesvioVol += r.desvio * vol;
      if (r.hc !== null && r.hc > 0) { e.sumHC += r.hc; e.cHC++; }
      if (r.horas !== null && r.horas > 0) e.sumHoras += r.horas;
      if (r.tph !== null && r.tph > 0) { e.sumTPH += r.tph; e.cTPH++; }
    }

    // Construir mapa OTD do FM — fonte primaria (mais confiavel que otd_cpt.csv)
    var fmOtdMap = {};
    for (var fi = 0; fi < dadosFM.weekly.length; fi++) {
      var fr = dadosFM.weekly[fi];
      if (fr.otd !== null && fr.otd !== undefined) {
        fmOtdMap[fr.semana + "|" + fr.node] = fr.otd;
      }
    }

    var extrasRows = [];
    var exKeys = Object.keys(exMap);
    for (var k = 0; k < exKeys.length; k++) {
      var e = exMap[exKeys[k]];
      var cptFinal = e.sumVol > 0 && e.sumCptVol > 0 ? e.sumCptVol / e.sumVol : (e.cCpt > 0 ? e.sumCpt / e.cCpt : 0);
      var otdFallback = e.sumVol > 0 && e.sumOtdVol > 0 ? e.sumOtdVol / e.sumVol : (e.cOtd > 0 ? e.sumOtd / e.cOtd : 0);
      // OTD: FM como primario; otd_cpt.csv como fallback
      var eKeyFM = e.semana + "|" + e.node;
      var otdFinal = (fmOtdMap[eKeyFM] !== undefined) ? fmOtdMap[eKeyFM] : otdFallback;
      var desvioFinal = e.sumVol > 0 ? e.sumDesvioVol / e.sumVol : 0;
      var hcFinal  = e.cHC  > 0 ? e.sumHC  / e.cHC  : 0;
      var tphFinal = e.cTPH > 0 ? e.sumTPH / e.cTPH : 0;
      extrasRows.push([e.semana, e.node,
        Math.round(cptFinal  * 100) / 100,
        Math.round(otdFinal  * 100) / 100,
        Math.round(e.sumForecast * 100) / 100,
        Math.round(desvioFinal   * 100) / 100,
        Math.round(hcFinal   * 100) / 100,
        Math.round(tphFinal  * 100) / 100]);
    }
    // Inserir nodes que tem OTD no FM mas nao existem no otd_cpt.csv
    var extrasKeySet = {};
    for (var ei3 = 0; ei3 < extrasRows.length; ei3++) {
      extrasKeySet[extrasRows[ei3][0] + "|" + extrasRows[ei3][1]] = true;
    }
    var fmOtdKeys = Object.keys(fmOtdMap);
    for (var fk = 0; fk < fmOtdKeys.length; fk++) {
      if (!extrasKeySet[fmOtdKeys[fk]]) {
        var fkParts = fmOtdKeys[fk].split("|");
        extrasRows.push([fkParts[0], fkParts[1], 0, fmOtdMap[fmOtdKeys[fk]], 0, 0, 0, 0]);
        Logger.log("OTD novo via FM: " + fmOtdKeys[fk] + " = " + fmOtdMap[fmOtdKeys[fk]] + "%");
      }
    }

    // OTD: FM e primario — ver bloco fmOtdMap acima
    extrasRows.sort(function(a, b) { return sortSemanas(a[0], b[0]); });
    if (extrasRows.length > 0) {
      var lastRowExtras = extrasSheet.getLastRow();
      if (lastRowExtras > 1) { extrasSheet.getRange(2, 1, lastRowExtras - 1, 8).clearContent(); }
      extrasSheet.getRange(2, 1, extrasRows.length, 8).setValues(extrasRows);
    }
  }

  // --- TPH ---
  var tphSheet = ss.getSheetByName("tph");
  if (tphSheet && dadosOTD.length > 0) {
    var tphRows = [];
    for (var i = 0; i < dadosOTD.length; i++) {
      var r = dadosOTD[i];
      if (!nodeConfig[r.node]) continue;
      tphRows.push([r.semana, r.node, "DF-SUL", r.tph || 0, r.hc || 0, r.horas || 0]);
    }
    tphRows.sort(function(a, b) { return sortSemanas(a[0], b[0]); });
    if (tphRows.length > 0) {
      var lastRowTPH = tphSheet.getLastRow();
      if (lastRowTPH > 1) { tphSheet.getRange(2, 1, lastRowTPH - 1, 6).clearContent(); }
      tphSheet.getRange(2, 1, tphRows.length, 6).setValues(tphRows);
    }
  }

  // --- SHIPMENT DATA ---
  var shipSheet = ss.getSheetByName("shipment_data");
  if (shipSheet && dadosShipment.weekly && dadosShipment.weekly.length > 0) {
    var shipRows = [];
    for (var i = 0; i < dadosShipment.weekly.length; i++) {
      var s = dadosShipment.weekly[i];
      shipRows.push([s.semana, s.node, s.grupo, s.tipo, s.miss_real, s.miss_d1_corrigido,
        s.recv_real, s.recv_corrigido, s.stow_real, s.stow_corrigido, s.dep_real, s.dep_corrigido]);
    }
    if (shipRows.length > 0) {
      var lastRowShip = shipSheet.getLastRow();
      if (lastRowShip > 1) { shipSheet.getRange(2, 1, lastRowShip - 1, 12).clearContent(); }
      shipSheet.getRange(2, 1, shipRows.length, 12).setValues(shipRows);
    }
  }

  // --- NODES DAILY ---
  var dailySheet = ss.getSheetByName("nodes_daily");
  if (dailySheet && dadosShipment.daily && dadosShipment.daily.length > 0) {
    var dailyRows = [];
    for (var i = 0; i < dadosShipment.daily.length; i++) {
      var d = dadosShipment.daily[i];
      dailyRows.push([d.data, d.node, d.grupo, d.tipo, d.miss_real, d.recv_miss, d.stow_miss, d.dep_miss, d.d1_corrigido, d.total]);
    }
    if (dailyRows.length > 0) {
      var lastRowDaily = dailySheet.getLastRow();
      if (lastRowDaily > 1) { dailySheet.getRange(2, 1, lastRowDaily - 1, 10).clearContent(); }
      dailySheet.getRange(2, 1, dailyRows.length, 10).setValues(dailyRows);
    }
  }

  // --- Atualizar config ---
  var configSheet = ss.getSheetByName("config");
  if (configSheet) {
    var now = new Date();
    var dataAtual = Utilities.formatDate(now, "GMT-3", "yyyy-MM-dd");
    var horaAtual = Utilities.formatDate(now, "GMT-3", "HH:mm");
    var cfgData = configSheet.getRange(2, 1, configSheet.getLastRow() - 1, 3).getValues();
    for (var i = 0; i < cfgData.length; i++) {
      if (cfgData[i][0] === "data_atualizacao") { configSheet.getRange(i + 2, 2).setValue(dataAtual); }
      if (cfgData[i][0] === "hora_atualizacao") { configSheet.getRange(i + 2, 2).setValue(horaAtual); }
    }
  }

  Logger.log("preencherTodasAbas concluido!");
}


// ============================================================================
// BUILD HTML 5 (WBR-EDSP) - NAO ALTERAR
// ============================================================================

function buildHTML5_(ss) {
  var t0 = new Date();
  var nodesSheet  = ss.getSheetByName("nodes_6w");
  var shipSheet   = ss.getSheetByName("shipment_data");
  var extrasSheet = ss.getSheetByName("wbr_extras");
  var dailySheet  = ss.getSheetByName("nodes_daily");
  var nodeConfig  = getNodeConfig_();

  if (!nodesSheet) { Logger.log("ERRO: nodes_6w nao encontrada"); return ""; }

  var semanasSet = {};
  var nDataRaw   = nodesSheet.getDataRange().getValues();
  for (var i = 1; i < nDataRaw.length; i++) {
    var sem = String(nDataRaw[i][0]).trim();
    if (sem) semanasSet[sem] = true;
  }
  var semanas = Object.keys(semanasSet);
  semanas.sort(function(a, b) {
    var na = parseInt(a.replace(/\D/g, ""), 10);
    var nb = parseInt(b.replace(/\D/g, ""), 10);
    return na - nb;
  });
  Logger.log("buildHTML5_: " + semanas.length + " semanas: " + semanas.join(", "));

  var nodesObj    = {};
  var allNodeKeys = Object.keys(nodeConfig);
  for (var ni = 0; ni < allNodeKeys.length; ni++) {
    var nd  = allNodeKeys[ni];
    var cfg = nodeConfig[nd];
    nodesObj[nd] = {
      regional: cfg.regional, grupo: cfg.grupo, tipo: cfg.tipo, owner: cfg.owner || "",
      d1_stow: cfg.d1_stow || false, d1_dep: cfg.d1_dep || false, semanas: {}
    };
  }

  for (var i = 1; i < nDataRaw.length; i++) {
    var row  = nDataRaw[i];
    var sem2 = String(row[0]).trim();
    var nd2  = String(row[1]).trim();
    if (nd2 === "EOZ8") continue;
    if (!sem2 || !nd2) continue;
    if (!nodesObj[nd2]) {
      nodesObj[nd2] = {
        regional: String(row[2] || ""), grupo: String(row[3] || ""),
        tipo: String(row[4] || ""), owner: "", d1_stow: false, d1_dep: false, semanas: {}
      };
    }
    nodesObj[nd2].semanas[sem2] = {
      volume:     Number(row[5])  || 0,
      receive:    row[6]  !== "" && row[6]  !== null ? Number(row[6])  : null,
      stow:       row[7]  !== "" && row[7]  !== null ? Number(row[7])  : null,
      depart:     row[8]  !== "" && row[8]  !== null ? Number(row[8])  : null,
      backlog:    Number(row[9])  || 0,
      backlogAcc: Number(row[10]) || 0,
      ppp: Number(row[11]) > 0 ? Number(row[11]) : null,
      missort:    Number(row[12]) || 0,
      wrongNode:  Number(row[13]) || 0
    };
  }

  var shipData = {};
  if (shipSheet) {
    var sData = shipSheet.getDataRange().getValues();
    for (var i = 1; i < sData.length; i++) {
      var sSem  = String(sData[i][0]).trim();
      var sNode = String(sData[i][1]).trim();
      if (!sSem || !sNode) continue;
      var sKey = sSem + "|" + sNode;
      shipData[sKey] = {
        stow_real:      Number(sData[i][8])  || 0,
        stow_corrigido: Number(sData[i][9])  || 0,
        dep_real:       Number(sData[i][10]) || 0,
        dep_corrigido:  Number(sData[i][11]) || 0
      };
    }
    Logger.log("shipData: " + Object.keys(shipData).length + " keys carregadas");
  }

  var d1AppliedCount = 0;
  var allNds = Object.keys(nodesObj);
  for (var ni2 = 0; ni2 < allNds.length; ni2++) {
    var ndKey = allNds[ni2];
    var ndCfg = nodesObj[ndKey];
    if (!ndCfg.d1_stow && !ndCfg.d1_dep) continue;

    for (var si = 0; si < semanas.length; si++) {
      var sSem3   = semanas[si];
      var shipKey = sSem3 + "|" + ndKey;
      var shipRec = shipData[shipKey];
      if (!shipRec) continue;
      var nodeWeek = ndCfg.semanas[sSem3];
      if (!nodeWeek) continue;

      if (ndCfg.d1_stow && shipRec.stow_corrigido > 0) {
        var volS    = nodeWeek.volume || 1;
        var newStow = (volS - shipRec.stow_real) / volS * 100;
        if (newStow > 100) newStow = 100;
        nodeWeek.stow = Math.round(newStow * 100) / 100;
        d1AppliedCount++;
      }
      if (ndCfg.d1_dep && shipRec.dep_corrigido > 0) {
        var volD   = nodeWeek.volume || 1;
        var newDep = (volD - shipRec.dep_real) / volD * 100;
        if (newDep > 100) newDep = 100;
        nodeWeek.depart = Math.round(newDep * 100) / 100;
        d1AppliedCount++;
      }
    }
  }
  Logger.log("D+1 aplicacoes total: " + d1AppliedCount);

  var extras = {};
  if (extrasSheet) {
    var exData = extrasSheet.getDataRange().getValues();
    for (var i = 1; i < exData.length; i++) {
      var eSem  = String(exData[i][0]).trim();
      var eNode = String(exData[i][1]).trim();
      if (!eSem || !eNode) continue;
      if (!extras[eNode]) extras[eNode] = {};
      extras[eNode][eSem] = {
        cpt:      exData[i][2] !== "" && exData[i][2] !== null ? Number(exData[i][2]) : null,
        otd:      exData[i][3] !== "" && exData[i][3] !== null ? Number(exData[i][3]) : null,
        forecast: Number(exData[i][4]) || 0,
        desvio:   Number(exData[i][5]) || 0,
        hc:       Number(exData[i][6]) || 0,
        tph:      Number(exData[i][7]) || 0
      };
    }
    Logger.log("extras (OTD/CPT/TPH): " + Object.keys(extras).length + " nodes carregados");
  }

  var dailyObj = {};
  if (dailySheet) {
    var dData = dailySheet.getDataRange().getValues();
    for (var i = 1; i < dData.length; i++) {
      var dDate = String(dData[i][0]).trim();
      var dNode = String(dData[i][1]).trim();
      if (!dDate || !dNode) continue;
      var dKey = dDate + "|" + dNode;
      dailyObj[dKey] = {
        miss_real:      Number(dData[i][4]) || 0,
        recv_miss:      Number(dData[i][5]) || 0,
        stow_miss_real: Number(dData[i][6]) || 0,
        dep_miss_real:  Number(dData[i][7]) || 0,
        d1_corrigido:   Number(dData[i][8]) || 0,
        total:          Number(dData[i][9]) || 0
      };
    }
  }

  var regSheet2    = ss.getSheetByName("regional_6w");
  var regionaisObj = {};
  if (regSheet2) {
    var regData = regSheet2.getDataRange().getValues();
    for (var i = 1; i < regData.length; i++) {
      var rSem = String(regData[i][0]).trim();
      var rReg = String(regData[i][1]).trim();
      if (!rSem || !rReg) continue;
      if (!regionaisObj[rReg]) regionaisObj[rReg] = { semanas: {} };
      regionaisObj[rReg].semanas[rSem] = {
        volume:  Number(regData[i][2]) || 0,
        receive: regData[i][3] !== "" && regData[i][3] !== null ? Number(regData[i][3]) : null,
        stow:    regData[i][4] !== "" && regData[i][4] !== null ? Number(regData[i][4]) : null,
        depart:  regData[i][5] !== "" && regData[i][5] !== null ? Number(regData[i][5]) : null,
        backlog: Number(regData[i][6]) || 0,
        ppp:     regData[i][8] !== "" && regData[i][8] !== null ? Number(regData[i][8]) : null
      };
    }
  }

  // D+1 NAS REGIONAIS
  var regNames = Object.keys(regionaisObj);
  for (var ri = 0; ri < regNames.length; ri++) {
    var regName     = regNames[ri];
    var regNodeList = [];
    for (var nj = 0; nj < allNds.length; nj++) {
      if (nodesObj[allNds[nj]].regional === regName) regNodeList.push(allNds[nj]);
    }
    if (regNodeList.length === 0) continue;
    for (var si2 = 0; si2 < semanas.length; si2++) {
      var sSem4 = semanas[si2];
      if (!regionaisObj[regName].semanas[sSem4]) continue;
      var sumStowW = 0, totalVolStow = 0, sumDepW = 0, totalVolDep = 0;
      for (var rn2 = 0; rn2 < regNodeList.length; rn2++) {
        var rndNode = regNodeList[rn2];
        var rndWeek = nodesObj[rndNode].semanas[sSem4];
        if (!rndWeek) continue;
        var vol = rndWeek.volume || 0;
        if (vol === 0) continue;
        if (rndWeek.stow !== null)   { sumStowW += rndWeek.stow   * vol; totalVolStow += vol; }
        if (rndWeek.depart !== null) { sumDepW  += rndWeek.depart * vol; totalVolDep  += vol; }
      }
      if (totalVolStow > 0) {
        regionaisObj[regName].semanas[sSem4].stow = Math.round((sumStowW / totalVolStow) * 100) / 100;
      }
      if (totalVolDep > 0) {
        regionaisObj[regName].semanas[sSem4].depart = Math.round((sumDepW / totalVolDep) * 100) / 100;
      }
    }
  }

  // TPH por regional
  if (extras && Object.keys(extras).length > 0) {
    for (var ri2 = 0; ri2 < regNames.length; ri2++) {
      var regName2  = regNames[ri2];
      var regNodes2 = [];
      for (var nj2 = 0; nj2 < allNds.length; nj2++) {
        if (nodesObj[allNds[nj2]].regional === regName2) regNodes2.push(allNds[nj2]);
      }
      for (var si3 = 0; si3 < semanas.length; si3++) {
        var sSem5 = semanas[si3];
        var tphSum = 0, tphCount = 0;
        for (var rn3 = 0; rn3 < regNodes2.length; rn3++) {
          var rnd3 = regNodes2[rn3];
          if (extras[rnd3] && extras[rnd3][sSem5] && extras[rnd3][sSem5].tph > 0) {
            tphSum += extras[rnd3][sSem5].tph; tphCount++;
          }
        }
        if (tphCount > 0 && regionaisObj[regName2] && regionaisObj[regName2].semanas[sSem5]) {
          regionaisObj[regName2].semanas[sSem5].tph = Math.round((tphSum / tphCount) * 100) / 100;
        }
      }
    }
  }

  var config = lerConfig_(ss);
  var dataObj = {
    lastUpdate: Utilities.formatDate(new Date(), "America/Sao_Paulo", "dd/MM/yyyy HH:mm"),
    semanas: semanas, nodes: nodesObj, regionais: regionaisObj, extras: extras, daily: dailyObj,
    metas: {
      recv: parseFloat(config["meta_recv"]) || 98.5,
      stow: parseFloat(config["meta_stow"]) || 99.5,
      dep:  parseFloat(config["meta_dep"])  || 99.5,
      ppp:  parseFloat(config["meta_ppp"])  || 175
    },
    labelConfig: { showAll: true, collisionDetection: true }
  };

  var js = "var D = " + JSON.stringify(dataObj) + ";";
  var elapsed = ((new Date() - t0) / 1000).toFixed(3);
  Logger.log("buildHTML5_: " + elapsed + "s | nodes=" + Object.keys(nodesObj).length +
             " | regionais=" + Object.keys(regionaisObj).length + " | semanas=" + semanas.length);
  return js;
}


// ============================================================================
// GRAVAR EDSPS 20W
// ============================================================================

function gravarEdsps20w_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var edspsSheet = ss.getSheetByName("edsps_20w");
  if (!edspsSheet) { Logger.log("edsps_20w: aba nao encontrada"); return; }
  var nodesSheet = ss.getSheetByName("nodes_6w");
  if (!nodesSheet || nodesSheet.getLastRow() < 2) { Logger.log("edsps_20w: nodes_6w vazia"); return; }

  var data     = nodesSheet.getDataRange().getValues();
  var spcNodes = ["EGO8","EIJ8","ESB8","ESE8","ECB8","ELO8"];
  var edspsRows = [];
  var COL_SEMANA=0, COL_NODE=1, COL_GRUPO=3, COL_VOLUME=5, COL_RECEIVE=6;
  var COL_STOW=7, COL_DEPART=8, COL_BACKLOG=9, COL_PPP=11;

  for (var i = 1; i < data.length; i++) {
    var row  = data[i];
    var node = String(row[COL_NODE]).trim();
    if (spcNodes.indexOf(node) === -1) continue;
    edspsRows.push([
      row[COL_SEMANA], node, row[COL_GRUPO] || "",
      Number(row[COL_VOLUME]) || 0, Number(row[COL_RECEIVE]) || 0,
      Number(row[COL_STOW]) || 0, Number(row[COL_DEPART]) || 0,
      Number(row[COL_BACKLOG]) || 0, Number(row[COL_PPP]) || 0
    ]);
  }
  edspsRows.sort(function(a, b) {
    return parseInt(String(a[0]).replace(/\D/g,""),10) - parseInt(String(b[0]).replace(/\D/g,""),10);
  });
  var lastRow = edspsSheet.getLastRow();
  if (lastRow > 1) edspsSheet.getRange(2, 1, lastRow - 1, 9).clearContent();
  if (edspsRows.length > 0) { edspsSheet.getRange(2, 1, edspsRows.length, 9).setValues(edspsRows); }
  Logger.log("edsps_20w FINAL: " + edspsRows.length + " linhas gravadas");
}


// ============================================================================
// GITHUB PUSH
// ============================================================================

function pushToGitHub_(token, owner, repo, path, content) {
  var apiUrl = "https://api.github.com/repos/" + owner + "/" + repo + "/contents/" + path;
  var getResp = UrlFetchApp.fetch(apiUrl, {
    headers: { "Authorization": "token " + token, "Accept": "application/vnd.github.v3+json" },
    muteHttpExceptions: true
  });
  var sha = "";
  if (getResp.getResponseCode() === 200) {
    sha = JSON.parse(getResp.getContentText()).sha;
  }
  var payload = {
    message: "Auto-update " + path + " via WBR script v3.1",
    content: Utilities.base64Encode(content, Utilities.Charset.UTF_8),
    branch:  "main"
  };
  if (sha) payload.sha = sha;
  var putResp = UrlFetchApp.fetch(apiUrl, {
    method: "put",
    headers: { "Authorization": "token " + token, "Accept": "application/vnd.github.v3+json" },
    contentType: "application/json", payload: JSON.stringify(payload), muteHttpExceptions: true
  });
  var code = putResp.getResponseCode();
  Logger.log("GitHub " + repo + "/" + path + ": " + code);
  if (code !== 200 && code !== 201) { throw new Error("GitHub push failed: " + code); }
}


// ============================================================================
// FUNCOES PRINCIPAIS (ORQUESTRADORAS)
// ============================================================================

function executarFluxoCompleto() {
  importarDados();
  atualizarTodos();
  gravarEdsps20w_();
}

function atualizarTodos() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var token = lerConfig_(ss)["github_token"] ||
              PropertiesService.getScriptProperties().getProperty("GITHUB_TOKEN");
  if (!token) { Logger.log("ERRO: Token GitHub nao configurado!"); return; }
  var owner = "adrianovie-rgb";
  try {
    var t0      = new Date().getTime();
    var content = buildHTML5_(ss);
    var elapsed = ((new Date().getTime() - t0) / 1000).toFixed(3);
    Logger.log("buildHTML5_: " + elapsed + "s");
    if (!content) { Logger.log("AVISO: buildHTML5_ retornou vazio"); return; }
    pushToGitHub_(token, owner, "WBR-EDSPs", "data.js", content);
  } catch (e) {
    Logger.log("ERRO WBR-EDSPs: " + e.message);
  }
}


// ============================================================================
// INCREMENTAL
// ============================================================================

function preencherTodasAbas_incremental(ss, dadosFM, dadosOTD, dadosShipment) {
  var nodeConfig  = getNodeConfig_();
  var nodesSheet  = ss.getSheetByName("nodes_6w");
  var regSheet    = ss.getSheetByName("regional_6w");

  var existingWeeks = [];
  if (nodesSheet) {
    var existData = nodesSheet.getDataRange().getValues();
    for (var i = 1; i < existData.length; i++) {
      var s = String(existData[i][0]);
      if (s && existingWeeks.indexOf(s) === -1) existingWeeks.push(s);
    }
  }

  var semanasNovas = [];
  for (var i = 0; i < dadosFM.weekly.length; i++) {
    var s = dadosFM.weekly[i].semana;
    if (existingWeeks.indexOf(s) === -1 && semanasNovas.indexOf(s) === -1) semanasNovas.push(s);
  }
  if (semanasNovas.length === 0) { Logger.log("Nenhuma semana nova para inserir"); return; }
  Logger.log("Semanas novas: " + semanasNovas.join(", "));

  var newFM   = dadosFM.weekly.filter(function(r) { return semanasNovas.indexOf(r.semana) > -1; });
  var newOTD  = dadosOTD.filter(function(r)       { return semanasNovas.indexOf(r.semana) > -1; });
  var newShip = dadosShipment.weekly.filter(function(r) { return semanasNovas.indexOf(r.semana) > -1; });

  if (nodesSheet && newFM.length > 0) {
    var lastRow = nodesSheet.getLastRow();
    var rows    = [];
    for (var i = 0; i < newFM.length; i++) {
      var r   = newFM[i];
      var cfg = nodeConfig[r.node] || {};
      rows.push([r.semana, r.node, r.regional || "DF-SUL", cfg.grupo || "", cfg.tipo || "",
        r.volume || 0, r.receive || 0, r.stow || 0, r.depart || 0,
        r.backlog || 0, r.backlogAcc || 0, r.ppp || 0, r.missort || 0, r.wrongNode || 0]);
    }
    nodesSheet.getRange(lastRow + 1, 1, rows.length, 14).setValues(rows);
  }

  var shipSheet = ss.getSheetByName("shipment_data");
  if (shipSheet && newShip.length > 0) {
    var lastRow3 = shipSheet.getLastRow();
    var shipRows = [];
    for (var i = 0; i < newShip.length; i++) {
      var s = newShip[i];
      shipRows.push([s.semana, s.node, s.grupo, s.tipo, s.miss_real, s.miss_d1_corrigido,
        s.recv_real, s.recv_corrigido, s.stow_real, s.stow_corrigido, s.dep_real, s.dep_corrigido]);
    }
    shipSheet.getRange(lastRow3 + 1, 1, shipRows.length, 12).setValues(shipRows);
  }

  var extrasSheet = ss.getSheetByName("wbr_extras");
  if (extrasSheet && newOTD.length > 0) {
    var lastRow4  = extrasSheet.getLastRow();
    var extrasRows = [];
    for (var i = 0; i < newOTD.length; i++) {
      var r = newOTD[i];
      extrasRows.push([r.semana, r.node, r.cpt || 0, r.otd || 0, r.forecast || 0, 0, 0, 0]);
    }
    extrasSheet.getRange(lastRow4 + 1, 1, extrasRows.length, 8).setValues(extrasRows);
  }

  Logger.log("Incremental concluido! Semanas: " + semanasNovas.join(", "));
}


// ============================================================================
// TRIGGER
// ============================================================================

function configurarTrigger() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === "executarFluxoCompleto") { ScriptApp.deleteTrigger(t); }
  });
  ScriptApp.newTrigger("executarFluxoCompleto")
    .timeBased().atHour(16).everyDays(1).inTimezone("America/Sao_Paulo").create();
  Logger.log("Trigger configurado: executarFluxoCompleto diario as 16h BRT");
}


// ============================================================================
// PREENCHER TODAS AS ABAS - UPSERT (v3.1)
// ============================================================================

function preencherTodasAbas_upsert(ss, dadosFM, dadosOTD, dadosShipment, dadosRegional) {
  var nodeConfig = getNodeConfig_();
  function sortSemanas(a, b) {
    return parseInt(a.replace(/\D/g,""),10) - parseInt(b.replace(/\D/g,""),10);
  }
  function upsertSheet_(sheet, newRows, keyCols, totalCols) {
    if (!sheet || newRows.length === 0) return 0;
    var lastRow = sheet.getLastRow();
    var existingData = [];
    var existingMap = {};
    if (lastRow > 1) {
      existingData = sheet.getRange(2, 1, lastRow - 1, totalCols).getValues();
      for (var i = 0; i < existingData.length; i++) {
        var key = "";
        for (var k = 0; k < keyCols.length; k++) { key += String(existingData[i][keyCols[k]]).trim() + "|"; }
        existingMap[key] = i;
      }
    }
    var inserted = 0, updated = 0;
    for (var n = 0; n < newRows.length; n++) {
      var newKey = "";
      for (var k = 0; k < keyCols.length; k++) { newKey += String(newRows[n][keyCols[k]]).trim() + "|"; }
      if (existingMap.hasOwnProperty(newKey)) { existingData[existingMap[newKey]] = newRows[n]; updated++; }
      else { existingData.push(newRows[n]); existingMap[newKey] = existingData.length - 1; inserted++; }
    }
    if (lastRow > 1) { sheet.getRange(2, 1, lastRow - 1, totalCols).clearContent(); }
    if (existingData.length > 0) { sheet.getRange(2, 1, existingData.length, totalCols).setValues(existingData); }
    Logger.log(sheet.getName() + ": " + updated + " atualizados, " + inserted + " inseridos (total: " + existingData.length + ")");
    return inserted + updated;
  }

  // --- REGIONAL 6W ---
  var regSheet = ss.getSheetByName("regional_6w");
  if (regSheet && dadosRegional && dadosRegional.length > 0) {
    var regRows = [];
    for (var r = 0; r < dadosRegional.length; r++) {
      var reg = dadosRegional[r];
      regRows.push([reg.semana, reg.regional, reg.volume, reg.receive, reg.stow, reg.depart, reg.backlog, 0, reg.ppp, 0]);
    }
    upsertSheet_(regSheet, regRows, [0, 1], 10);
  }

  // --- NODES 6W ---
  var nodesSheet = ss.getSheetByName("nodes_6w");
  if (nodesSheet && dadosFM.weekly.length > 0) {
    var nodeRows = [];
    for (var i = 0; i < dadosFM.weekly.length; i++) {
      var r = dadosFM.weekly[i];
      if (!nodeConfig[r.node]) continue;
      var cfg = nodeConfig[r.node] || {};
      nodeRows.push([r.semana, r.node, r.regional || "DF-SUL", cfg.grupo || "", cfg.tipo || "",
        r.volume || 0, r.receive || 0, r.stow || 0, r.depart || 0,
        r.backlog || 0, r.backlogAcc || 0, r.ppp || 0, r.missort || 0, r.wrongNode || 0]);
    }
    nodeRows.sort(function(a, b) { return sortSemanas(a[0], b[0]); });
    upsertSheet_(nodesSheet, nodeRows, [0, 1], 14);
  }

  // --- NODES W1 ---
  var w1Sheet = ss.getSheetByName("nodes_w1");
  if (w1Sheet && dadosFM.weekly.length > 0) {
    var allWeeks = [];
    for (var i = 0; i < dadosFM.weekly.length; i++) {
      if (allWeeks.indexOf(dadosFM.weekly[i].semana) === -1) allWeeks.push(dadosFM.weekly[i].semana);
    }
    allWeeks.sort(sortSemanas);
    var lastWeek = allWeeks[allWeeks.length - 1];
    var w1Rows = [];
    for (var i = 0; i < dadosFM.weekly.length; i++) {
      var r = dadosFM.weekly[i];
      if (r.semana !== lastWeek || !nodeConfig[r.node]) continue;
      var cfg = nodeConfig[r.node] || {};
      w1Rows.push([r.semana, r.node, r.regional || "DF-SUL", cfg.grupo || "", cfg.tipo || "",
        r.volume || 0, r.receive || 0, r.stow || 0, r.depart || 0, r.backlog || 0, r.ppp || 0]);
    }
    if (w1Rows.length > 0) {
      var lastRowW1 = w1Sheet.getLastRow();
      if (lastRowW1 > 1) { w1Sheet.getRange(2, 1, lastRowW1 - 1, 11).clearContent(); }
      w1Sheet.getRange(2, 1, w1Rows.length, 11).setValues(w1Rows);
    }
  }

  // --- REGIONAL PPP COMPARATIVO ---
  var pppSheet = ss.getSheetByName("regional_ppp_comparativo");
  if (pppSheet && dadosFM.weekly.length > 0) {
    var pppMap = {};
    for (var i = 0; i < dadosFM.weekly.length; i++) {
      var r = dadosFM.weekly[i];
      if (r.ppp === null || r.ppp === 0 || !nodeConfig[r.node]) continue;
      var pk = r.semana + "|" + (r.regional || "DF-SUL");
      if (!pppMap[pk]) { pppMap[pk] = { semana: r.semana, regional: r.regional || "DF-SUL", sum: 0, count: 0 }; }
      pppMap[pk].sum += r.ppp; pppMap[pk].count++;
    }
    var pppRows = [];
    var pppKeys = Object.keys(pppMap);
    pppKeys.sort(function(a, b) { return sortSemanas(a.split("|")[0], b.split("|")[0]); });
    for (var k = 0; k < pppKeys.length; k++) {
      var e = pppMap[pppKeys[k]];
      pppRows.push([e.semana, e.regional, e.count > 0 ? e.sum / e.count : 0]);
    }
    upsertSheet_(pppSheet, pppRows, [0, 1], 3);
  }

  // --- WBR EXTRAS ---
  var extrasSheet = ss.getSheetByName("wbr_extras");
  if (extrasSheet && dadosOTD.length > 0) {
    var exMap = {};
    for (var i = 0; i < dadosOTD.length; i++) {
      var r = dadosOTD[i];
      if (!r.semana || !r.node) continue;
      var eKey = r.semana + "|" + r.node;
      if (!exMap[eKey]) {
        exMap[eKey] = { semana: r.semana, node: r.node,
          sumCptVol: 0, sumOtdVol: 0, sumVol: 0, sumForecast: 0, sumDesvioVol: 0,
          sumHC: 0, cHC: 0, sumHoras: 0, sumTPH: 0, cTPH: 0,
          sumCpt: 0, cCpt: 0, sumOtd: 0, cOtd: 0 };
      }
      var e = exMap[eKey];
      var vol = (r.volProc && r.volProc > 0) ? r.volProc : 0;
      if (r.cpt !== null && r.cpt > 0) { if (vol > 0) { e.sumCptVol += r.cpt * vol; } e.sumCpt += r.cpt; e.cCpt++; }
      if (r.otd !== null && r.otd > 0) { if (vol > 0) { e.sumOtdVol += r.otd * vol; } e.sumOtd += r.otd; e.cOtd++; }
      e.sumVol += vol;
      if (r.forecast !== null && r.forecast > 0) e.sumForecast += r.forecast;
      if (r.desvio !== null && vol > 0) e.sumDesvioVol += r.desvio * vol;
      if (r.hc !== null && r.hc > 0) { e.sumHC += r.hc; e.cHC++; }
      if (r.horas !== null && r.horas > 0) e.sumHoras += r.horas;
      if (r.tph !== null && r.tph > 0) { e.sumTPH += r.tph; e.cTPH++; }
    }
    // Construir mapa OTD do FM — fonte primaria (mais confiavel que otd_cpt.csv)
    var fmOtdMap = {};
    for (var fi = 0; fi < dadosFM.weekly.length; fi++) {
      var fr = dadosFM.weekly[fi];
      if (fr.otd !== null && fr.otd !== undefined) {
        fmOtdMap[fr.semana + "|" + fr.node] = fr.otd;
      }
    }

    var extrasRows = [];
    var exKeys = Object.keys(exMap);
    for (var k = 0; k < exKeys.length; k++) {
      var e = exMap[exKeys[k]];
      var cptFinal = e.sumVol > 0 && e.sumCptVol > 0 ? e.sumCptVol / e.sumVol : (e.cCpt > 0 ? e.sumCpt / e.cCpt : 0);
      var otdFallback = e.sumVol > 0 && e.sumOtdVol > 0 ? e.sumOtdVol / e.sumVol : (e.cOtd > 0 ? e.sumOtd / e.cOtd : 0);
      // OTD: FM como primario; otd_cpt.csv como fallback
      var eKeyFM = e.semana + "|" + e.node;
      var otdFinal = (fmOtdMap[eKeyFM] !== undefined) ? fmOtdMap[eKeyFM] : otdFallback;
      var desvioFinal = e.sumVol > 0 ? e.sumDesvioVol / e.sumVol : 0;
      var hcFinal  = e.cHC  > 0 ? e.sumHC  / e.cHC  : 0;
      var tphFinal = e.cTPH > 0 ? e.sumTPH / e.cTPH : 0;
      extrasRows.push([e.semana, e.node,
        Math.round(cptFinal  * 100) / 100,
        Math.round(otdFinal  * 100) / 100,
        Math.round(e.sumForecast * 100) / 100,
        Math.round(desvioFinal   * 100) / 100,
        Math.round(hcFinal   * 100) / 100,
        Math.round(tphFinal  * 100) / 100]);
    }
    // Inserir nodes que tem OTD no FM mas nao existem no otd_cpt.csv
    var extrasKeySet = {};
    for (var ei3 = 0; ei3 < extrasRows.length; ei3++) {
      extrasKeySet[extrasRows[ei3][0] + "|" + extrasRows[ei3][1]] = true;
    }
    var fmOtdKeys = Object.keys(fmOtdMap);
    for (var fk = 0; fk < fmOtdKeys.length; fk++) {
      if (!extrasKeySet[fmOtdKeys[fk]]) {
        var fkParts = fmOtdKeys[fk].split("|");
        extrasRows.push([fkParts[0], fkParts[1], 0, fmOtdMap[fmOtdKeys[fk]], 0, 0, 0, 0]);
        Logger.log("OTD novo via FM: " + fmOtdKeys[fk] + " = " + fmOtdMap[fmOtdKeys[fk]] + "%");
      }
    }

    // OTD: FM e primario — ver bloco fmOtdMap acima
    extrasRows.sort(function(a, b) { return sortSemanas(a[0], b[0]); });
    upsertSheet_(extrasSheet, extrasRows, [0, 1], 8);
  }

  // --- TPH ---
  var tphSheet = ss.getSheetByName("tph");
  if (tphSheet && dadosOTD.length > 0) {
    var tphRows = [];
    for (var i = 0; i < dadosOTD.length; i++) {
      var r = dadosOTD[i];
      if (!nodeConfig[r.node]) continue;
      tphRows.push([r.semana, r.node, "DF-SUL", r.tph || 0, r.hc || 0, r.horas || 0]);
    }
    tphRows.sort(function(a, b) { return sortSemanas(a[0], b[0]); });
    upsertSheet_(tphSheet, tphRows, [0, 1], 6);
  }

  // --- SHIPMENT DATA ---
  var shipSheet = ss.getSheetByName("shipment_data");
  if (shipSheet && dadosShipment.weekly && dadosShipment.weekly.length > 0) {
    var shipRows = [];
    for (var i = 0; i < dadosShipment.weekly.length; i++) {
      var s = dadosShipment.weekly[i];
      shipRows.push([s.semana, s.node, s.grupo, s.tipo, s.miss_real, s.miss_d1_corrigido,
        s.recv_real, s.recv_corrigido, s.stow_real, s.stow_corrigido, s.dep_real, s.dep_corrigido]);
    }
    upsertSheet_(shipSheet, shipRows, [0, 1], 12);
  }

  // --- NODES DAILY ---
  var dailySheet = ss.getSheetByName("nodes_daily");
  if (dailySheet && dadosShipment.daily && dadosShipment.daily.length > 0) {
    var dailyRows = [];
    for (var i = 0; i < dadosShipment.daily.length; i++) {
      var d = dadosShipment.daily[i];
      dailyRows.push([d.data, d.node, d.grupo, d.tipo, d.miss_real, d.recv_miss, d.stow_miss, d.dep_miss, d.d1_corrigido, d.total]);
    }
    upsertSheet_(dailySheet, dailyRows, [0, 1], 10);
  }

  // --- Atualizar config ---
  var configSheet = ss.getSheetByName("config");
  if (configSheet) {
    var now = new Date();
    var dataAtual = Utilities.formatDate(now, "GMT-3", "yyyy-MM-dd");
    var horaAtual = Utilities.formatDate(now, "GMT-3", "HH:mm");
    var cfgData = configSheet.getRange(2, 1, configSheet.getLastRow() - 1, 3).getValues();
    for (var i = 0; i < cfgData.length; i++) {
      if (cfgData[i][0] === "data_atualizacao") { configSheet.getRange(i + 2, 2).setValue(dataAtual); }
      if (cfgData[i][0] === "hora_atualizacao") { configSheet.getRange(i + 2, 2).setValue(horaAtual); }
    }
  }

  Logger.log("preencherTodasAbas_upsert concluido!");
}


// ============================================================================
// AUTO-PREENCHER FERIADOS NA ABA CONFIG
// ============================================================================

function preencherFeriadosConfig_(ss) {
  var configSheet = ss.getSheetByName("config");
  if (!configSheet) return;
  var feriados = [
    { data: "2026-01-01", nome: "Confraternizacao Universal (01/01)" },
    { data: "2026-02-16", nome: "Carnaval (16/02)" },
    { data: "2026-02-17", nome: "Carnaval (17/02)" },
    { data: "2026-04-03", nome: "Sexta-feira Santa (03/04)" },
    { data: "2026-04-21", nome: "Tiradentes (21/04)" },
    { data: "2026-05-01", nome: "Dia do Trabalho (01/05)" },
    { data: "2026-06-04", nome: "Corpus Christi (04/06)" },
    { data: "2026-09-07", nome: "Independencia (07/09)" },
    { data: "2026-10-12", nome: "N. Sra. Aparecida (12/10)" },
    { data: "2026-11-02", nome: "Finados (02/11)" },
    { data: "2026-11-15", nome: "Proclamacao da Republica (15/11)" },
    { data: "2026-11-20", nome: "Consciencia Negra (20/11)" },
    { data: "2026-12-25", nome: "Natal (25/12)" }
  ];
  var feriadosPorSemana = {};
  for (var i = 0; i < feriados.length; i++) {
    var parts = feriados[i].data.split("-");
    var d  = new Date(parseInt(parts[0]), parseInt(parts[1]) - 1, parseInt(parts[2]));
    var dt = new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()));
    var dayNum = dt.getUTCDay() || 7;
    dt.setUTCDate(dt.getUTCDate() + 4 - dayNum);
    var yearStart = new Date(Date.UTC(dt.getUTCFullYear(), 0, 1));
    var weekNo = Math.ceil((((dt - yearStart) / 86400000) + 1) / 7);
    var weekKey = "W" + weekNo;
    if (feriadosPorSemana[weekKey]) { feriadosPorSemana[weekKey] += " + " + feriados[i].nome; }
    else { feriadosPorSemana[weekKey] = feriados[i].nome; }
  }
  var lastRow = configSheet.getLastRow();
  var cfgData = lastRow > 1 ? configSheet.getRange(2, 1, lastRow - 1, 3).getValues() : [];
  var existingRows = {};
  for (var i = 0; i < cfgData.length; i++) {
    existingRows[String(cfgData[i][0]).trim()] = i + 2;
  }
  var semanasComFeriado = Object.keys(feriadosPorSemana);
  semanasComFeriado.sort(function(a,b) { return parseInt(a.replace("W",""))-parseInt(b.replace("W","")); });
  var rowsToAdd = [];
  for (var s = 0; s < semanasComFeriado.length; s++) {
    var semana = semanasComFeriado[s];
    var weekNum = parseInt(semana.replace("W",""));
    var paramName = "feriado_W" + weekNum;
    var valor = feriadosPorSemana[semana];
    var descricao = "Feriado nacional (auto-preenchido)";
    if (existingRows[paramName]) {
      configSheet.getRange(existingRows[paramName], 2).setValue(valor);
      configSheet.getRange(existingRows[paramName], 3).setValue(descricao);
    } else { rowsToAdd.push([paramName, valor, descricao]); }
  }
  for (var param in existingRows) {
    if (param.indexOf("feriado_W") === 0 && !feriadosPorSemana["W" + param.replace("feriado_W","")]) {
      configSheet.getRange(existingRows[param], 2).setValue("");
      configSheet.getRange(existingRows[param], 3).setValue("Sem feriado");
    }
  }
  if (rowsToAdd.length > 0) {
    var appendRow = configSheet.getLastRow() + 1;
    configSheet.getRange(appendRow, 1, rowsToAdd.length, 3).setValues(rowsToAdd);
  }
  Logger.log("Feriados config: " + semanasComFeriado.length + " semanas preenchidas");
}


function importarDados_upsert(ss) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
  var pastas = DriveApp.getFoldersByName("_automacao_wbr");
  if (!pastas.hasNext()) { Logger.log("ERRO: Pasta _automacao_wbr nao encontrada"); return; }
  var pasta = pastas.next();
  var fmContent = ""; var otdContent = ""; var shipContent = ""; var regContent = "";
  var fmFiles  = pasta.getFilesByName("fm_metrics.csv");  if (fmFiles.hasNext())  fmContent  = fmFiles.next().getBlob().getDataAsString("UTF-8");
  var otdFiles = pasta.getFilesByName("otd_cpt.csv");     if (otdFiles.hasNext()) otdContent = otdFiles.next().getBlob().getDataAsString("UTF-8");
  var shipFiles= pasta.getFilesByName("shipment.csv");    if (shipFiles.hasNext())shipContent= shipFiles.next().getBlob().getDataAsString("UTF-8");
  var regFiles = pasta.getFilesByName("REGIONAL.csv");
  if (regFiles.hasNext()) { regContent = regFiles.next().getBlob().getDataAsString("UTF-8"); Logger.log("REGIONAL.csv: " + regContent.length + " chars"); }
  else { Logger.log("REGIONAL.csv: nao encontrado"); }
  var dadosFM       = processarCSV_FM(fmContent);
  var dadosOTD      = processarExcel_OTD(otdContent);
  var dadosShipment = processarCSV_Shipment(shipContent);
  var dadosRegional = processarCSV_Regional(regContent);
  preencherTodasAbas_upsert(ss, dadosFM, dadosOTD, dadosShipment, dadosRegional);
  Logger.log("importarDados_upsert concluido!");
}


function executarFluxoCompleto_upsert() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  importarDados_upsert(ss);
  preencherFeriadosConfig_(ss);
  atualizarTodos();
  gravarEdsps20w_();
}

function configurarTrigger_upsert() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    var fn = t.getHandlerFunction();
    if (fn === "executarFluxoCompleto" || fn === "executarFluxoCompleto_upsert") { ScriptApp.deleteTrigger(t); }
  });
  ScriptApp.newTrigger("executarFluxoCompleto_upsert")
    .timeBased().atHour(16).everyDays(1).inTimezone("America/Sao_Paulo").create();
  Logger.log("Trigger configurado: executarFluxoCompleto_upsert diario as 16h BRT");
}


// ============================================================================
// EXPORTAR JSON PARA WBR DASHBOARD
// ============================================================================

function onOpen() {
  SpreadsheetApp.getUi().createMenu('EDSP')
    .addItem('Gerar JSON para App', 'gerarJSON').addToUi();
}

function gerarJSON() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var abas = [
    { nome: 'config', tipo: 'kv' },
    { nome: 'shipment_data', tipo: 'tabela' }, { nome: 'regional_6w', tipo: 'tabela' },
    { nome: 'nodes_6w', tipo: 'tabela' }, { nome: 'edsps_20w', tipo: 'tabela' },
    { nome: 'nodes_w1', tipo: 'tabela' }, { nome: 'nodes_daily', tipo: 'tabela' },
    { nome: 'tph', tipo: 'tabela' }, { nome: 'wbr_extras', tipo: 'tabela' },
    { nome: 'regional_ppp_comparativo', tipo: 'tabela' }
  ];
  var result = {};
  for (var i = 0; i < abas.length; i++) {
    var aba = abas[i];
    var sheet = ss.getSheetByName(aba.nome);
    if (!sheet) {
      var sheets = ss.getSheets();
      for (var s = 0; s < sheets.length; s++) {
        if (sheets[s].getName().toLowerCase().replace(/\s+/g,"_") === aba.nome.toLowerCase()) { sheet = sheets[s]; break; }
      }
    }
    if (!sheet) continue;
    if (aba.tipo === 'kv') { result[aba.nome] = lerConfig_(sheet); }
    else { result[aba.nome] = lerTabela_(sheet); }
  }
  var json = JSON.stringify(result);
  var html = HtmlService.createHtmlOutput(
    '<style>body{font-family:Arial;padding:16px}textarea{width:100%;height:300px;font-size:11px;font-family:monospace}button{margin-top:12px;padding:10px 24px;background:#3b82f6;color:white;border:none;border-radius:6px;font-size:14px;cursor:pointer}button:hover{background:#2563eb}.info{font-size:12px;color:#666;margin-bottom:8px}</style>' +
    '<p class="info">JSON gerado com ' + Object.keys(result).length + ' abas. Tamanho: ' + (json.length / 1024).toFixed(1) + ' KB</p>' +
    '<textarea id="jsonArea">' + escapeHtml_(json) + '</textarea>' +
    '<button onclick="copiar()">Copiar para Clipboard</button>' +
    '<span id="status" style="margin-left:12px;color:#16a34a;font-size:13px"></span>' +
    '<script>function copiar(){var t=document.getElementById("jsonArea");t.select();document.execCommand("copy");document.getElementById("status").textContent="Copiado!";}<\/script>'
  ).setWidth(700).setHeight(450).setTitle('JSON para WBR Dashboard');
  ui.showModalDialog(html, 'JSON para WBR Dashboard');
}

function lerConfig_(sheet) {
  var data = sheet.getDataRange().getValues();
  var config = {};
  for (var i = 0; i < data.length; i++) {
    var chave = String(data[i][0] || '').trim();
    var valor = data[i][1];
    if (chave) config[chave] = formatValue_(valor);
  }
  return config;
}

function lerTabela_(sheet) {
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  var headers = [];
  for (var h = 0; h < data[0].length; h++) { headers.push(String(data[0][h] || '').trim()); }
  var rows = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var hasData = false;
    for (var j = 0; j < row.length; j++) {
      if (row[j] !== '' && row[j] !== null && row[j] !== undefined) { hasData = true; break; }
    }
    if (!hasData) continue;
    var obj = {};
    for (var c = 0; c < headers.length; c++) {
      if (headers[c]) obj[headers[c]] = formatValue_(row[c]);
    }
    rows.push(obj);
  }
  return rows;
}

function formatValue_(val) {
  if (val === null || val === undefined || val === '') return null;
  if (val instanceof Date) return Utilities.formatDate(val, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
  if (typeof val === 'number') return val;
  return String(val);
}

function escapeHtml_(text) {
  return text.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}

// ===== FIM DO SCRIPT WBR-EDSPs v3.1 =====
