// ============================================================
// EdoTec Configurator — Google Apps Script (versiune actualizata)
// Inlocuieste tot codul din editorul Google Apps Script cu acesta
// ============================================================

const SALES_EMAIL = 'contact@edotec.ro';
const SALES_CC    = 'ionut.matei@edotec.ro';

// ID-ul spreadsheet-ului Google Sheets (cel existent)
const SHEET_ID = '1Xxy8_FZXaf4BxcZVKqgON1sBOy-fIIqZHCEAX0YtoVg';

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);

    // ── Actiune: logare comanda in Sheets ──
    if (data.action === 'logOrder') {
      logOrderToSheet(data);
      return ContentService.createTextOutput('logged').setMimeType(ContentService.MimeType.TEXT);
    }

    // ── Actiune default: trimitere email (comportament existent) ──
    if (data.subject && data.htmlBody && data.to) {
      GmailApp.sendEmail(data.to, data.subject, '', {
        htmlBody: data.htmlBody,
        cc: data.cc || '',
        name: 'EdoTec Configurator'
      });
    }

    return ContentService.createTextOutput('ok').setMimeType(ContentService.MimeType.TEXT);
  } catch(err) {
    return ContentService.createTextOutput('error: ' + err.message).setMimeType(ContentService.MimeType.TEXT);
  }
}

function doGet(e) {
  return ContentService.createTextOutput('EdoTec GAS OK').setMimeType(ContentService.MimeType.TEXT);
}

// ── Log comanda in foaia "Comenzi" ──
function logOrderToSheet(data) {
  const ss = SpreadsheetApp.openById(SHEET_ID);

  // Creaza foaia "Comenzi" daca nu exista
  let sheet = ss.getSheetByName('Comenzi');
  if (!sheet) {
    sheet = ss.insertSheet('Comenzi');
    // Header row
    sheet.appendRow([
      'Data', 'Ref', 'Masina', 'Caroserie',
      'Prenume', 'Nume', 'Telefon', 'Email',
      'Adresa', 'Oras', 'Judet', 'Oras Montaj',
      'Metoda Plata', 'Livrare',
      'Total Materiale RON', 'Total + Montaj RON',
      'Zone', 'Nr Zone'
    ]);
    // Formatare header
    sheet.getRange(1, 1, 1, 18).setFontWeight('bold').setBackground('#1a1a2e').setFontColor('#ffffff');
    sheet.setFrozenRows(1);
  }

  // Adauga rand cu datele comenzii
  sheet.appendRow([
    new Date(data.timestamp || Date.now()),
    data.refNum || '',
    data.car || '',
    data.bodyClass || '',
    data.fname || '',
    data.lname || '',
    data.phone || '',
    data.email || '',
    data.address || '',
    data.city || '',
    data.county || '',
    data.montajCity || '',
    data.paymentMethod || '',
    data.deliveryMethod || '',
    data.materialeRON || 0,
    data.grandTotal || 0,
    data.zones || '',
    data.zonesCount || 0
  ]);

  // Auto-resize coloane la prima comanda
  if (sheet.getLastRow() === 2) {
    sheet.autoResizeColumns(1, 18);
  }

  // ── Creaza/actualizeaza foaia "Raport Lunar" ──
  updateMonthlyReport(ss, data);
}

// ── Actualizeaza raportul lunar de zone/pachete ──
function updateMonthlyReport(ss, data) {
  let rSheet = ss.getSheetByName('Raport Zone');
  if (!rSheet) {
    rSheet = ss.insertSheet('Raport Zone');
    rSheet.appendRow(['Zona / Pachet', 'Nr Comenzi', 'Ultima comanda']);
    rSheet.getRange(1, 1, 1, 3).setFontWeight('bold').setBackground('#e74c3c').setFontColor('#ffffff');
    rSheet.setFrozenRows(1);
  }

  // Descompune zonele (separate prin virgula)
  const zones = (data.zones || '').split(', ');
  zones.forEach(zone => {
    if (!zone.trim()) return;
    const dataRange = rSheet.getDataRange().getValues();
    let found = false;
    for (let i = 1; i < dataRange.length; i++) {
      if (dataRange[i][0] === zone.trim()) {
        rSheet.getRange(i + 1, 2).setValue(dataRange[i][1] + 1);
        rSheet.getRange(i + 1, 3).setValue(new Date());
        found = true;
        break;
      }
    }
    if (!found) {
      rSheet.appendRow([zone.trim(), 1, new Date()]);
    }
  });
}
