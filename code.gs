/* ============================================================
   PRODUZIONE INTERNI CASSETTO - API REST (Google Apps Script)
   ------------------------------------------------------------
   Il frontend (index.html su GitHub Pages) chiama questa Web App
   via fetch() POST inviando un JSON { action, ...params }.
   Il routing avviene nello switch/case di doPost().

   Deploy: Distribuisci > Nuova distribuzione > App web
     - Esegui come: Me
     - Chi ha accesso: Chiunque
   Copia l'URL /exec e incollalo in API_URL dentro index.html.
   ============================================================ */

function doPost(e) {
  try {
    const request = JSON.parse(e.postData.contents);
    const action = request.action;

    let result;
    switch (action) {
      case 'getMacrobolleAperte':
        result = getMacrobolleAperte();
        break;
      case 'cercaMacrobolla':
        result = cercaMacrobolla(request.macrobolla);
        break;
      case 'aggiornaProduzione':
        result = aggiornaProduzione(request);
        break;
      case 'dichiaraOperatori':
        result = dichiaraOperatori(request);
        break;
      case 'getDashboardData':
        result = getDashboardData();
        break;
      default:
        throw new Error('Azione non riconosciuta: ' + action);
    }

    return jsonResponse(result);
  } catch (err) {
    return jsonResponse({ error: err.message });
  }
}

function doGet() {
  return jsonResponse({
    status: 'ok',
    message: 'API Produzione Interni Cassetto attiva'
  });
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/* =========================
   UTILITY
========================= */

function getMacrobolleSheet() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Macrobolle');
}

function formatDateForDisplay(value) {
  if (!value) return '';
  const d = new Date(value);
  if (isNaN(d.getTime())) return '';
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'dd/MM/yyyy');
}

function startOfDay(date) {
  const d = new Date(date);
  d.setHours(0, 0, 0, 0);
  return d;
}

/* =========================
   COLONNE FOGLIO MACROBOLLE
   A=0  Macrobolla
   B=1  Data inserimento
   C=2  Stabilimento
   D=3  Data di consegna
   E=4  Tipologia
   F=5  Quantità da produrre
   G=6  Quantità prodotta
   H=7  Stato
========================= */

/* =========================
   PRODUZIONE
========================= */

function getMacrobolleAperte() {
  const sheet = getMacrobolleSheet();
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  if (lastRow < 2) return [];

  const values = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  const aperte = [];

  for (let i = 0; i < values.length; i++) {
    const row = values[i];

    const macrobolla = String(row[0] || '').trim();      // A
    const stabilimento = String(row[2] || '').trim();    // C
    const dataConsegna = row[3];                         // D
    const tipologia = String(row[4] || '').trim();       // E
    const qtaDaProdurre = Number(row[5] || 0);           // F
    const qtaProdotta = Number(row[6] || 0);             // G
    const stato = String(row[7] || '').trim().toUpperCase(); // H

    if (stato === 'CHIUSO') continue;

    const residuo = qtaDaProdurre - qtaProdotta;
    if (residuo <= 0) continue;

    if (!(dataConsegna instanceof Date) || isNaN(dataConsegna.getTime())) continue;

    aperte.push({
      macrobolla: macrobolla,
      stabilimento: stabilimento,
      dataConsegna: Utilities.formatDate(dataConsegna, Session.getScriptTimeZone(), 'dd/MM'),
      dataConsegnaRaw: dataConsegna.getTime(),
      tipologia: tipologia,
      residuo: residuo
    });
  }

  aperte.sort(function(a, b) {
    return a.dataConsegnaRaw - b.dataConsegnaRaw;
  });

  return aperte;
}

function cercaMacrobolla(macrobolla) {
  const sheet = getMacrobolleSheet();
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  if (lastRow < 2) {
    return { ok: false, messaggio: 'Il foglio Macrobolle non contiene dati.' };
  }

  const values = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  const codice = String(macrobolla || '').trim();

  if (!codice) {
    return { ok: false, messaggio: 'Inserisci una macrobolla valida.' };
  }

  const trovati = [];

  for (let i = 0; i < values.length; i++) {
    const row = values[i];

    const macrobollaFoglio = String(row[0] || '').trim(); // A
    const stabilimento = String(row[2] || '').trim();     // C
    const dataConsegna = row[3];                          // D
    const tipologia = String(row[4] || '').trim();        // E
    const qtaDaProdurre = Number(row[5] || 0);            // F
    const qtaProdotta = Number(row[6] || 0);              // G
    const stato = String(row[7] || '').trim().toUpperCase(); // H

    if (macrobollaFoglio === codice) {
      const residuo = qtaDaProdurre - qtaProdotta;

      if (residuo > 0 && stato !== 'CHIUSO') {
        trovati.push({
          rowIndex: i + 2,
          macrobolla: macrobollaFoglio,
          stabilimento: stabilimento,
          dataConsegna: formatDateForDisplay(dataConsegna),
          tipologia: tipologia,
          qtaDaProdurre: qtaDaProdurre,
          qtaProdotta: qtaProdotta,
          residuo: residuo,
          stato: stato || 'APERTO'
        });
      }
    }
  }

  if (trovati.length === 0) {
    // Verifica se esiste ma è chiusa
    const esiste = values.some(row => String(row[0] || '').trim() === codice);
    if (esiste) {
      return {
        ok: false,
        messaggio: 'La macrobolla esiste ma risulta già chiusa o senza quantità residua.'
      };
    }
    return {
      ok: false,
      messaggio: 'Macrobolla non presente nel foglio.'
    };
  }

  return {
    ok: true,
    records: trovati
  };
}

function getConsuntivoSheet() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Consuntivo Macrobolle');
}

function aggiornaProduzione(payload) {
  const macrobolla = String(payload.macrobolla || '').trim();
  const tipologia = String(payload.tipologia || '').trim();
  const quantita = Number(payload.quantita);

  if (!macrobolla) {
    throw new Error('Macrobolla non valida.');
  }

  if (!tipologia) {
    throw new Error('Tipologia non valida.');
  }

  if (!quantita || quantita <= 0) {
    throw new Error('La quantità deve essere maggiore di zero.');
  }

  const sheet = getConsuntivoSheet();
  if (!sheet) {
    throw new Error('Foglio "Consuntivo Macrobolle" non trovato.');
  }

  const oggi = new Date();

  // Colonne: A=Data inserimento, B=Macrobolla, C=Tipologia, D=Cassetti prodotti
  sheet.appendRow([
    oggi,
    macrobolla,
    tipologia,
    quantita
  ]);

  return {
    ok: true,
    messaggio: 'Produzione registrata per la macrobolla ' + macrobolla + '.',
    quantitaRegistrata: quantita,
    dataRegistrazione: formatDateForDisplay(oggi)
  };
}

/* =========================
   OPERATORI
========================= */

function getProduzioneSheet() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Produzione');
}

function dichiaraOperatori(payload) {
  const dataStr = String(payload.data || '').trim();
  const numOperatori = Number(payload.numOperatori);
  const oreLavorate = Number(payload.oreLavorate);

  if (!dataStr) {
    throw new Error('Data non valida.');
  }

  if (!numOperatori || numOperatori <= 0) {
    throw new Error('Il numero di operatori deve essere maggiore di zero.');
  }

  if (!oreLavorate || oreLavorate <= 0) {
    throw new Error('Le ore lavorate devono essere maggiori di zero.');
  }

  const sheet = getProduzioneSheet();
  if (!sheet) {
    throw new Error('Foglio "Produzione" non trovato.');
  }

  // Parsing data dal formato yyyy-MM-dd (input type=date)
  const parti = dataStr.split('-');
  const dataTarget = new Date(Number(parti[0]), Number(parti[1]) - 1, Number(parti[2]));
  dataTarget.setHours(0, 0, 0, 0);

  // Cerca se esiste già una riga per questa data
  const lastRow = sheet.getLastRow();

  if (lastRow >= 2) {
    const dates = sheet.getRange(2, 1, lastRow - 1, 1).getValues();

    for (let i = 0; i < dates.length; i++) {
      const cellDate = dates[i][0];
      if (cellDate instanceof Date) {
        const d = new Date(cellDate);
        d.setHours(0, 0, 0, 0);
        if (d.getTime() === dataTarget.getTime()) {
          // Aggiorna riga esistente: B (Numero Operatori) e C (Ore totali lavorate)
          sheet.getRange(i + 2, 2).setValue(numOperatori);
          sheet.getRange(i + 2, 3).setValue(oreLavorate);
          return {
            ok: true,
            messaggio: 'Dati aggiornati per il ' + formatDateForDisplay(dataTarget) + ': ' + numOperatori + ' operatori, ' + oreLavorate + ' ore.',
            aggiornato: true
          };
        }
      }
    }
  }

  // Nuova riga: A (Data), B (Numero Operatori), C (Ore totali lavorate). D-E sono formule
  sheet.appendRow([dataTarget, numOperatori, oreLavorate]);

  return {
    ok: true,
    messaggio: 'Dati registrati per il ' + formatDateForDisplay(dataTarget) + ': ' + numOperatori + ' operatori, ' + oreLavorate + ' ore.',
    aggiornato: false
  };
}

/* =========================
   DASHBOARD
========================= */

function getIsoWeekInfo(date) {
  const d = startOfDay(date);
  d.setDate(d.getDate() + 4 - (d.getDay() || 7));
  const yearStart = new Date(d.getFullYear(), 0, 1);
  const weekNo = Math.ceil((((d - yearStart) / 86400000) + 1) / 7);
  return {
    year: d.getFullYear(),
    week: weekNo
  };
}

function getStartOfIsoWeek(date) {
  const d = startOfDay(date);
  const day = d.getDay() || 7;
  d.setDate(d.getDate() - day + 1);
  d.setHours(0, 0, 0, 0);
  return d;
}

function addDays(date, days) {
  const d = new Date(date);
  d.setDate(d.getDate() + days);
  return d;
}

function initBucketMap(labels) {
  const obj = {};
  labels.forEach(label => obj[label] = 0);
  return obj;
}

function ensureGroup(container, key, labels) {
  if (!container[key]) {
    container[key] = {
      daEvadere: initBucketMap(labels),
      daProdurre: initBucketMap(labels),
      prodotti: initBucketMap(labels)
    };
  }
}

function getDashboardColumnsInfo() {
  const today = startOfDay(new Date());
  const currentWeekStart = getStartOfIsoWeek(today);

  const weekWindows = [];
  const labels = ['Scad'];

  for (let i = 0; i < 6; i++) {
    const start = addDays(currentWeekStart, i * 7);
    const end = addDays(start, 7);
    const info = getIsoWeekInfo(start);
    const label = String(info.week);

    weekWindows.push({
      label: label,
      start: start,
      end: end
    });

    labels.push(label);
  }

  labels.push('Oltre');

  return {
    today: today,
    currentWeekStart: currentWeekStart,
    weekWindows: weekWindows,
    labels: labels,
    lastVisibleWeekEnd: weekWindows[weekWindows.length - 1].end
  };
}

function getDeliveryBucket(deliveryDate, residuo, stato, columnsInfo) {
  const consegna = startOfDay(deliveryDate);

  // Scaduto: data prima di oggi, residuo > 0 e stato APERTO
  if (consegna < columnsInfo.today && residuo > 0 && stato === 'APERTO') {
    return 'Scad';
  }

  // Settimana corrente + 5 successive
  for (let i = 0; i < columnsInfo.weekWindows.length; i++) {
    const w = columnsInfo.weekWindows[i];
    if (consegna >= w.start && consegna < w.end) {
      return w.label;
    }
  }

  // Tutto oltre l'ultima settimana visibile
  if (consegna >= columnsInfo.lastVisibleWeekEnd) {
    return 'Oltre';
  }

  return null;
}

function getDashboardData() {
  const sheet = getMacrobolleSheet();
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  const columnsInfo = getDashboardColumnsInfo();

  const result = {
    generatedAt: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm'),
    columns: columnsInfo.labels,
    byPlant: {},
    byMachine: {},
    summary: {
      daEvadere: initBucketMap(columnsInfo.labels),
      daProdurre: initBucketMap(columnsInfo.labels),
      prodotti: initBucketMap(columnsInfo.labels)
    }
  };

  if (lastRow < 2) {
    return result;
  }

  const values = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

  values.forEach(row => {
    const stabilimento = String(row[2] || '').trim();  // C
    const dataConsegna = row[3];                       // D
    const tipologia = String(row[4] || '').trim();     // E
    const qtaDaProdurre = Number(row[5] || 0);         // F
    const qtaProdotta = Number(row[6] || 0);           // G
    const stato = String(row[7] || '').trim().toUpperCase(); // H

    if (!(dataConsegna instanceof Date) || isNaN(dataConsegna.getTime())) return;
    if (!stabilimento || !tipologia) return;

    const residuo = Math.max(qtaDaProdurre - qtaProdotta, 0);
    const bucket = getDeliveryBucket(dataConsegna, residuo, stato, columnsInfo);

    if (!bucket) return;

    ensureGroup(result.byPlant, stabilimento, columnsInfo.labels);
    ensureGroup(result.byMachine, tipologia, columnsInfo.labels);

    result.byPlant[stabilimento].daEvadere[bucket] += qtaDaProdurre;
    result.byMachine[tipologia].daEvadere[bucket] += qtaDaProdurre;
    result.summary.daEvadere[bucket] += qtaDaProdurre;

    result.byPlant[stabilimento].daProdurre[bucket] += residuo;
    result.byMachine[tipologia].daProdurre[bucket] += residuo;
    result.summary.daProdurre[bucket] += residuo;

    result.byPlant[stabilimento].prodotti[bucket] += qtaProdotta;
    result.byMachine[tipologia].prodotti[bucket] += qtaProdotta;
    result.summary.prodotti[bucket] += qtaProdotta;
  });

  return result;
}
