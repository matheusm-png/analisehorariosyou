// Google Apps Script — Dashboard Leads Inbound (Análise de Horários)
// Deploy como Web App: Executar como "Eu", Acesso "Qualquer pessoa"

function doGet(e) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('LEADS GERAIS 2026');
    var data = sheet.getDataRange().getValues();
    var headers = data[0];

    var idx = {};
    for (var i = 0; i < headers.length; i++) {
      idx[String(headers[i]).trim()] = i;
    }

    var raw = [];
    for (var row = 1; row < data.length; row++) {
      var r = data[row];
      if (!r[idx['Data']]) continue; // pula linhas vazias

      raw.push({
        h: parseHour(r[idx['Hora']]),
        d: getDow(r[idx['Data']]),
        m: String(r[idx['Mês']] || 'NI').trim(),
        s: normStatus(r[idx['Status']]),
        o: normOrigem(r[idx['Origem']]),
        e: mapEmail(r[idx['Email Corporativo?']]),
        p: mapPorte(r[idx['NR de Colab']])
      });
    }

    return ContentService
      .createTextOutput(JSON.stringify(raw))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function parseHour(v) {
  if (!v) return 0;
  var s = String(v).replace('h', '').trim();
  var n = parseInt(s, 10);
  return isNaN(n) ? 0 : n;
}

// Apps Script: getDay() → 0=Dom, 1=Seg ... 6=Sáb
// Dashboard usa: 0=Seg ... 5=Sáb, 6=Dom  (igual ao Python weekday)
function getDow(v) {
  if (!v) return 0;
  var d;
  if (v instanceof Date) {
    d = v;
  } else {
    // Formato DD/MM/YYYY
    var parts = String(v).split('/');
    if (parts.length === 3) {
      d = new Date(parts[2], parts[1] - 1, parts[0]);
    } else {
      d = new Date(v);
    }
  }
  return (d.getDay() + 6) % 7; // converte para 0=Seg...6=Dom
}

function normStatus(v) {
  if (!v) return 'NI';
  var lut = {
    'vendido': 'VENDIDO',
    'agendado': 'Agendado',
    'descartado': 'Descartado',
    'em tentativa': 'Em tentativa',
    'agendamento q': 'Agendamento q',
    'agendamento nq': 'Agendamento nq'
  };
  return lut[String(v).trim().toLowerCase()] || String(v).trim();
}

function normOrigem(v) {
  if (!v) return 'NI';
  return String(v).trim().toLowerCase();
}

function mapEmail(v) {
  if (!v) return 'Invalido';
  var s = String(v).trim().toLowerCase();
  if (s === 'sim') return 'Sim';
  if (s === 'nao' || s === 'não') return 'Outro';
  return 'Invalido';
}

function mapPorte(v) {
  if (!v) return 'NI';
  var s = String(v).trim().replace(/\s/g, '');
  var lut = {
    '20-30': '20-30',
    '31-60': '31-60', '30-60': '31-60', '31-100': '31-60',
    '61-120': '61-120', '61-300': '61-120', '121-300': '61-120',
    '101-500': '101-500', '201-500': '101-500',
    '301-500': '301-500',
    '500+': '500+', '501+': '500+'
  };
  return lut[s] || 'NI';
}
