const STG_CONFIG = {
  surveyId: '74a431a9-5960-41cc-af8f-30ed3270e5c2',
  apiUrl: 'https://api.dooblo.net/newapi',
  dataSheet: 'Raw_Data',
  logSheet: 'Sync_Log',
  chunkSize: 45000
};

function syncSTGSurvey() {
  const props = PropertiesService.getScriptProperties();
  const apiKey = props.getProperty('STG_API_KEY');
  const password = props.getProperty('STG_PASSWORD');

  if (!apiKey || !password) {
    throw new Error('Missing API credentials.');
  }

  const auth = 'Basic ' +
    Utilities.base64Encode(apiKey + ':' + password);

  // 1. Get interview IDs
  const idsUrl = STG_CONFIG.apiUrl +
    '/SurveyInterviewIDs?surveyIDs=' +
    encodeURIComponent(STG_CONFIG.surveyId);

  const idsResponse = UrlFetchApp.fetch(idsUrl, {
    method: 'get',
    headers: {
      Authorization: auth,
      Accept: 'application/json'
    },
    muteHttpExceptions: true
  });

  if (idsResponse.getResponseCode() !== 200) {
    throw new Error('Interview ID error: HTTP ' +
      idsResponse.getResponseCode());
  }

  const interviewIds = JSON.parse(idsResponse.getContentText());

  if (!Array.isArray(interviewIds)) {
    throw new Error('Unexpected Interview IDs response.');
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let dataSheet = ss.getSheetByName(STG_CONFIG.dataSheet);
  if (!dataSheet) {
    dataSheet = ss.insertSheet(STG_CONFIG.dataSheet);
  }

  const syncTime = new Date();
  const rows = [];

  // 2. Download in batches of 99
  //    (Sheet is NOT cleared yet, so the dashboard keeps reading the previous
  //     data while downloads run.)
  for (let i = 0; i < interviewIds.length; i += 99) {
    const batch = interviewIds.slice(i, i + 99);

    const dataUrl = STG_CONFIG.apiUrl +
      '/SurveyInterviewData?subjectIDs=' +
      encodeURIComponent(batch.join(',')) +
      '&surveyID=' + encodeURIComponent(STG_CONFIG.surveyId) +
      '&onlyHeaders=false&includeNulls=false';

    const response = UrlFetchApp.fetch(dataUrl, {
      method: 'get',
      headers: {
        Authorization: auth,
        Accept: 'text/xml'
      },
      muteHttpExceptions: true
    });

    if (response.getResponseCode() !== 200) {
      throw new Error('Data download failed: HTTP ' +
        response.getResponseCode());
    }

    const xml = response.getContentText();

    // 3. Split XML into chunks below 50,000 characters
    const batchNumber = Math.floor(i / 99) + 1;
    const chunkSize = STG_CONFIG.chunkSize;

    for (let j = 0; j < xml.length; j += chunkSize) {
      rows.push([
        syncTime,
        batchNumber,
        Math.floor(j / chunkSize) + 1,
        // FIX: leading apostrophe forces Sheets to store the chunk as plain text.
        // Without it, a chunk that happens to start with "=", "+", "-" or "@"
        // is parsed as a formula and becomes #ERROR!, which loses ~45,000
        // characters of XML. The apostrophe itself is not saved in the cell.
        "'" + xml.substring(j, j + chunkSize)
      ]);
    }

    Utilities.sleep(600);
  }

  // 4. Write chunks to Google Sheets
  //    FIX: clear only now, right before writing, so the sheet is empty for
  //    as short a time as possible.
  dataSheet.clearContents();
  dataSheet.getRange(1, 1, 1, 4).setValues([[
    'Sync Time',
    'Batch',
    'Chunk',
    'XML Data'
  ]]);

  if (rows.length > 0) {
    dataSheet.getRange(2, 1, rows.length, 4).setValues(rows);
    dataSheet.setFrozenRows(1);
    dataSheet.setColumnWidth(1, 160);
    dataSheet.setColumnWidth(2, 80);
    dataSheet.setColumnWidth(3, 80);
    dataSheet.setColumnWidth(4, 500);
  }

  // FIX: check that no chunk was turned into an error value.
  const written = rows.length > 0
    ? dataSheet.getRange(2, 4, rows.length, 1).getValues()
    : [];
  const badChunks = [];
  written.forEach(function (r, k) {
    if (String(r[0]).length !== rows[k][3].length - 1) {
      badChunks.push(rows[k][1] + '/' + rows[k][2]);
    }
  });

  // 5. Record sync log
  let logSheet = ss.getSheetByName(STG_CONFIG.logSheet);

  if (!logSheet) {
    logSheet = ss.insertSheet(STG_CONFIG.logSheet);
    logSheet.appendRow([
      'Sync Time',
      'Survey ID',
      'Interview Count',
      'XML Chunks',
      'Status'
    ]);
  }

  logSheet.appendRow([
    syncTime,
    STG_CONFIG.surveyId,
    interviewIds.length,
    rows.length,
    badChunks.length ? 'WARNING: bad chunks ' + badChunks.join(', ') : 'SUCCESS'
  ]);

  SpreadsheetApp.flush();

  Logger.log('Total interviews: ' + interviewIds.length);
  Logger.log('XML chunks saved: ' + rows.length);
  if (badChunks.length) Logger.log('Bad chunks: ' + badChunks.join(', '));
  Logger.log('Sync completed.');
}
