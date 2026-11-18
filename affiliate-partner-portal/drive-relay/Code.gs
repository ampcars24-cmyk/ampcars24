// Cars24 Affiliate Portal — Apps Script Drive relay
//
// This relay reads the latest Excel exports from a private Drive folder.
// The portal's OIDC server calls it with a server-only shared secret, then
// filters partner rows after it has validated the signed-in user's identity.
// Never put RELAY_SHARED_SECRET in index.html or any browser-delivered code.
//
// Required Script Properties:
//   RELAY_SHARED_SECRET       long random secret shared with the portal server
//   EXPORT_FOLDER_ID          Drive folder containing the daily .xlsx exports
// Optional Script Properties:
//   ADMIN_EXPORT_FILENAME     defaults to "Sampled X Orignal.xlsx"
//   PARTNER_EXPORT_FILENAME   defaults to "Report Level.xlsx"
//
// Enable the Advanced Drive service (v3) in Apps Script before deployment.

var SHEET_ID = '1-x0H80na3INSHd0_sXOuTdvOHzcMLUd7K8eXRc6KlfA';
var SHEET_NAME = 'Config';
var GOOGLE_SHEETS_MIME = 'application/vnd.google-apps.spreadsheet';

function doGet(e) {
  var p = (e && e.parameter) || {};
  var action = String(p.action || 'status');
  if (action === 'status') {
    return _json({ ok: true, status: 'Cars24 Drive relay active' });
  }
  if (!_authorized(p.token)) return _json({ error: 'Unauthorized' });
  if (action === 'getConfig') return _getConfig();
  if (action === 'saveConfig') return _saveConfig(p.configKey, p.configValue, p.ts);
  if (action === 'getData') return _getData(p.dataset);
  return _json({ error: 'Unsupported action' });
}

function doPost(e) {
  try {
    var body = {};
    try {
      body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    } catch (parseError) {
      return _json({ error: 'Request body must be valid JSON' });
    }

    if (!_authorized(body.token)) return _json({ error: 'Unauthorized' });
    switch (String(body.action || '')) {
      case 'getData':
        return _getData(body.dataset);
      case 'getConfig':
        return _getConfig();
      case 'saveConfig':
        return _saveConfig(body.configKey, body.configValue, body.ts);
      case 'upload':
        return _upload(body);
      default:
        return _json({ error: 'Unsupported action' });
    }
  } catch (err) {
    console.error('Relay request failed: ' + err.message);
    return _json({ error: 'Relay request failed' });
  }
}

function _authorized(candidate) {
  var expected = PropertiesService.getScriptProperties().getProperty('RELAY_SHARED_SECRET') || '';
  candidate = String(candidate || '');
  if (!expected || !candidate) return false;
  var left = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, expected, Utilities.Charset.UTF_8);
  var right = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, candidate, Utilities.Charset.UTF_8);
  var difference = 0;
  for (var i = 0; i < left.length; i++) difference |= left[i] ^ right[i];
  return difference === 0;
}

function _getData(dataset) {
  dataset = String(dataset || '');
  if (dataset !== 'admin' && dataset !== 'partner') return _json({ error: 'dataset must be admin or partner' });

  var props = PropertiesService.getScriptProperties();
  var folderId = props.getProperty('EXPORT_FOLDER_ID');
  if (!folderId) return _json({ error: 'EXPORT_FOLDER_ID is not configured' });

  var filename = dataset === 'admin'
    ? (props.getProperty('ADMIN_EXPORT_FILENAME') || 'Sampled X Orignal.xlsx')
    : (props.getProperty('PARTNER_EXPORT_FILENAME') || 'Report Level.xlsx');
  var file = _latestFile(folderId, filename);
  if (!file) return _json({ error: 'Export file not found: ' + filename });

  var cache = CacheService.getScriptCache();
  var cacheKey = 'export-v1-' + dataset + '-' + file.id + '-' + String(file.modifiedTime || '');
  var parsed = null;
  var cached = cache.get(cacheKey);
  if (cached) {
    try { parsed = JSON.parse(cached); } catch (cacheError) { parsed = null; }
  }
  if (!parsed) {
    parsed = _readExcel(file);
    try {
      var serialized = JSON.stringify(parsed);
      if (serialized.length <= 95000) cache.put(cacheKey, serialized, 900);
    } catch (cacheWriteError) {
      console.warn('Export cache skipped');
    }
  }
  return _json({
    ok: true,
    dataset: dataset,
    filename: file.name,
    updatedAt: file.modifiedTime || new Date().toISOString(),
    period: parsed.period,
    columns: parsed.columns,
    rows: parsed.rows,
    summary: parsed.summary
  });
}

function _latestFile(folderId, filename) {
  var safeName = String(filename).replace(/\\/g, '\\\\').replace(/'/g, "\\'");
  var query = "'" + folderId + "' in parents and name = '" + safeName + "' and trashed = false";
  var result = Drive.Files.list({
    q: query,
    pageSize: 100,
    orderBy: 'modifiedTime desc',
    fields: 'files(id,name,mimeType,modifiedTime,size)',
    supportsAllDrives: true,
    includeItemsFromAllDrives: true
  });
  var files = result.files || [];
  return files.length ? files[0] : null;
}

function _readExcel(file) {
  var convertedId = null;
  try {
    var converted = Drive.Files.create({
      name: '_cars24_import_' + Date.now(),
      mimeType: GOOGLE_SHEETS_MIME
    }, DriveApp.getFileById(file.id).getBlob(), { fields: 'id' });
    convertedId = converted.id;

    var spreadsheet = SpreadsheetApp.openById(convertedId);
    var sheet = spreadsheet.getSheets()[0];
    if (!sheet) throw new Error('Workbook has no worksheet');
    var values = sheet.getDataRange().getValues();
    var headerIndex = -1;
    for (var i = 0; i < values.length; i++) {
      if (_normalize(values[i][0]) === 'partner name') {
        headerIndex = i;
        break;
      }
    }
    if (headerIndex < 0) throw new Error('Export is missing the Partner Name header');

    var columns = values[headerIndex].map(function(value, index) {
      var label = String(value == null ? '' : value).replace(/^C2B Affiliate\s*\|\s*/i, '').trim();
      return label || 'Column ' + (index + 1);
    });
    var partnerColumn = -1;
    for (var c = 0; c < columns.length; c++) {
      if (_normalize(columns[c]) === 'partner name') {
        partnerColumn = c;
        break;
      }
    }
    if (partnerColumn < 0) throw new Error('Export is missing a partner name column');

    var dataRows = [];
    var summary = null;
    for (var r = headerIndex + 1; r < values.length; r++) {
      var row = values[r];
      var partnerName = _normalize(row[partnerColumn]);
      var hasValue = row.some(function(value) { return value !== '' && value !== null; });
      if (partnerName && partnerName !== 'partner name') {
        var item = {};
        columns.forEach(function(column, index) { item[column] = _jsonValue(row[index]); });
        dataRows.push(item);
      } else if (!partnerName && hasValue && !summary) {
        summary = {};
        columns.forEach(function(column, index) { summary[column] = _jsonValue(row[index]); });
      }
    }

    var period = null;
    for (var pr = 0; pr < headerIndex && !period; pr++) {
      for (var pc = 0; pc < values[pr].length; pc++) {
        var candidate = values[pr][pc];
        if (typeof candidate === 'string' && /^[A-Za-z]{3}-\d{2,4}$/.test(candidate.trim())) {
          period = candidate.trim();
          break;
        }
      }
    }
    return { columns: columns, rows: dataRows, summary: summary, period: period };
  } finally {
    if (convertedId) {
      try { Drive.Files.remove(convertedId); } catch (cleanupError) { console.warn('Temporary import cleanup failed'); }
    }
  }
}

function _getConfig() {
  try {
    var sheet = _getSheet();
    var data = sheet.getDataRange().getValues();
    var config = {};
    for (var i = 1; i < data.length; i++) {
      var key = String(data[i][0] || '').trim();
      if (!key) continue;
      var value = String(data[i][1] == null ? '' : data[i][1]);
      var timestamp = data[i][2];
      try { config[key] = JSON.parse(value); } catch (parseError) { config[key] = value; }
      if (timestamp) config['__ts_' + key] = Number(timestamp);
    }
    // Preserve the original relay's flat {key:value, __ts_key:timestamp} shape.
    return _json(config);
  } catch (err) {
    console.error('Config read failed: ' + err.message);
    return _json({ error: 'Config read failed' });
  }
}

function _saveConfig(key, value, ts) {
  try {
    key = String(key || '').trim();
    if (!key) return _json({ error: 'Missing configKey' });
    value = String(value == null ? '' : value);
    var sheet = _getSheet();
    var data = sheet.getDataRange().getValues();
    var rowNumber = -1;
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0] || '').trim() === key) { rowNumber = i + 1; break; }
    }
    var timestamp = ts ? Number(ts) : Date.now();
    if (!Number.isFinite(timestamp)) timestamp = Date.now();
    if (rowNumber > 0) sheet.getRange(rowNumber, 2, 1, 2).setValues([[value, timestamp]]);
    else sheet.appendRow([key, value, timestamp]);
    return _json({ ok: true, key: key, ts: timestamp });
  } catch (err) {
    console.error('Config write failed: ' + err.message);
    return _json({ error: 'Config write failed' });
  }
}

function _getSheet() {
  var spreadsheet = SpreadsheetApp.openById(SHEET_ID);
  var sheet = spreadsheet.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(SHEET_NAME);
    sheet.appendRow(['key', 'value', 'timestamp']);
    sheet.getRange(1, 1, 1, 3).setFontWeight('bold');
  }
  return sheet;
}

function _upload(body) {
  var base64Data = String(body.base64Data || '');
  if (!base64Data) return _json({ error: 'No base64Data provided' });
  if (base64Data.length > 14000000) return _json({ error: 'Upload exceeds the 10 MB limit' });
  var fileName = String(body.fileName || ('upload_' + Date.now())).replace(/[\\/:*?"<>|]/g, '_');
  var folderName = String(body.folderName || 'Cars24_Uploads');
  var folders = DriveApp.getFoldersByName(folderName);
  var folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);
  var mimeType = String(body.mimeType || 'application/octet-stream');
  var file = folder.createFile(Utilities.newBlob(Utilities.base64Decode(base64Data), mimeType, fileName));
  // Keep the uploaded invoice/proof private to the Drive account owner.
  return _json({ ok: true, success: true, fileId: file.getId(), fileName: fileName,
    fileUrl: 'https://drive.google.com/file/d/' + file.getId() + '/view', url: 'https://drive.google.com/file/d/' + file.getId() + '/view' });
}

function _normalize(value) {
  return String(value == null ? '' : value).trim().replace(/\s+/g, ' ').toLocaleLowerCase();
}

function _jsonValue(value) {
  return value instanceof Date ? value.toISOString() : value;
}

function _json(value) {
  return ContentService.createTextOutput(JSON.stringify(value)).setMimeType(ContentService.MimeType.JSON);
}
