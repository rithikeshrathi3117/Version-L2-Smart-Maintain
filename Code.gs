/**
 * SMART MAINTAIN – BPCL Gonda LPG Plant
 * Shared Google Sheets database + mandatory daily AMC report + expiry reminders.
 *
 * DEPLOYMENT CHECKLIST
 * 1. Update Code.gs in the EXISTING Apps Script project connected to the /exec URL in Index.html.
 * 2. Save, then run setupSmartMaintain() once and approve the requested permissions.
 * 3. Run diagnoseAmcSync() and check the execution log.
 * 4. Deploy > Manage deployments > Edit > New version > Deploy.
 * 5. Keep the deployment URL ending in /exec unchanged in Index.html.
 * 6. Deploy as "Execute as: Me". Restrict access to BPCL users if the organisation's
 *    Google Workspace settings permit it. If the endpoint is deployed as "Anyone",
 *    anyone who obtains the URL can submit changes; this is not an authenticated API.
 * 7. Set the Apps Script project time zone to Asia/Kolkata.
 *
 * NOTE: The website uses a cross-origin no-cors POST. Its response is opaque to the
 * browser, so Index.html confirms each sync by reading syncMeta.requestId back using
 * JSONP. A successful fetch() call alone is not treated as proof that data was saved.
 */

const RECIPIENTS = [
  'rithikeshrathi@bharatpetroleum.in',
  'gauravmadan@bharatpetroleum.in',
  'patilvishalravi@bharatpetroleum.in',
  'atharavsanjayka@bharatpetroleum.in',
  'rithikeshrathi3117@gmail.com',
  'anandrao@bharatpetroleum.in'
];

const TZ = 'Asia/Kolkata';
const PROP_KEY = 'SMART_MAINTAIN_SPREADSHEET_ID';
const REMINDER_DAYS_KEY = 'SMART_MAINTAIN_REMINDER_DAYS';
const REMINDERS_ENABLED_KEY = 'SMART_MAINTAIN_REMINDERS_ENABLED';
const REMINDER_DAYS_DEFAULT = [60, 45, 30, 15, 10, 5, 2, 0];

const SHEET_HEADERS = {
  'Equipment Register': ['Equipment ID', 'Equipment JSON'],
  'AMC Contracts': ['AMC ID', 'Contract JSON'],
  'Maintenance Logs': ['Log ID', 'Log JSON'],
  'Plant Data': ['Data Key', 'Data JSON'],
  'Engineering Advisories': ['Advisory ID', 'Advisory JSON'],
  'Sync Status': ['Key', 'Value'],
  'Email Log': ['Sent At (IST)', 'AMC ID', 'Contract Ref', 'Expiry Date', 'Reminder Days', 'Recipients']
};

const SYNCABLE_PLANT_KEYS = ['team', 'initiatives', 'training', 'plantGallery'];

function setupSmartMaintain() {
  const props = PropertiesService.getScriptProperties();
  let spreadsheetId = props.getProperty(PROP_KEY);
  let ss;

  if (spreadsheetId) {
    ss = SpreadsheetApp.openById(spreadsheetId);
  } else {
    ss = SpreadsheetApp.create('SMART MAINTAIN – Gonda LPG Plant AMC Register');
    props.setProperty(PROP_KEY, ss.getId());
  }

  Object.keys(SHEET_HEADERS).forEach(name => ensureSheet_(ss, name, SHEET_HEADERS[name]));

  if (!props.getProperty(REMINDER_DAYS_KEY)) {
    props.setProperty(REMINDER_DAYS_KEY, JSON.stringify(REMINDER_DAYS_DEFAULT));
  }
  if (!props.getProperty(REMINDERS_ENABLED_KEY)) {
    props.setProperty(REMINDERS_ENABLED_KEY, 'true');
  }

  installDailyReminderTrigger_();
  Logger.log('SMART MAINTAIN spreadsheet: ' + ss.getUrl());
  Logger.log('Equipment rows: ' + readJsonRows_(ss.getSheetByName('Equipment Register')).length);
  Logger.log('AMC contract rows: ' + readJsonRows_(ss.getSheetByName('AMC Contracts')).length);
  Logger.log('Daily mandatory email and retry triggers installed.');
  Logger.log('Configured recipients: ' + RECIPIENTS.join(', '));
  return ss.getUrl();
}

function ensureSheet_(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getDb_() {
  const id = PropertiesService.getScriptProperties().getProperty(PROP_KEY);
  if (!id) {
    throw new Error('Shared database is not initialized. Run setupSmartMaintain() once from the Apps Script editor.');
  }
  return SpreadsheetApp.openById(id);
}

function installDailyReminderTrigger_() {
  const handlers = ['sendDailyMandatoryAmcMail', 'retryDailyMandatoryAmcMail', 'checkAmcExpiry'];
  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (handlers.indexOf(trigger.getHandlerFunction()) >= 0) {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger('sendDailyMandatoryAmcMail')
    .timeBased().everyDays(1).atHour(8).inTimezone(TZ).create();

  // Retry checks every 15 minutes. Once the mandatory mail succeeds for the day,
  // the success property prevents a duplicate report.
  ScriptApp.newTrigger('retryDailyMandatoryAmcMail')
    .timeBased().everyMinutes(15).create();

  ScriptApp.newTrigger('checkAmcExpiry')
    .timeBased().everyDays(1).atHour(8).inTimezone(TZ).create();
}

function reinstallDailyReminderTrigger() {
  installDailyReminderTrigger_();
  Logger.log('Daily mandatory AMC report trigger and expiry reminder trigger reinstalled.');
}

/* --------------------------- HTTP API --------------------------- */

function jsonResponse_(value, callback) {
  const json = JSON.stringify(value);
  if (callback && /^[a-zA-Z_$][0-9a-zA-Z_$]*$/.test(callback)) {
    return ContentService.createTextOutput(callback + '(' + json + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  const params = (e && e.parameter) || {};
  const action = String(params.action || '');
  const callback = String(params.callback || '');

  try {
    if (action === 'get') {
      const ss = getDb_();
      const result = {
        ok: true,
        assets: readJsonRows_(ss.getSheetByName('Equipment Register')),
        contracts: readJsonRows_(ss.getSheetByName('AMC Contracts')),
        logs: readJsonRows_(ss.getSheetByName('Maintenance Logs')),
        plantData: readPlantData_(ss),
        syncMeta: getSyncMeta_(ss)
      };
      return jsonResponse_(result, callback);
    }

    if (action === 'health') {
      const ss = getDb_();
      const triggerNames = ScriptApp.getProjectTriggers().map(t => t.getHandlerFunction());
      return jsonResponse_({
        ok: true,
        databaseReady: true,
        spreadsheetUrl: ss.getUrl(),
        assetCount: readJsonRows_(ss.getSheetByName('Equipment Register')).length,
        contractCount: readJsonRows_(ss.getSheetByName('AMC Contracts')).length,
        logCount: readJsonRows_(ss.getSheetByName('Maintenance Logs')).length,
        syncMeta: getSyncMeta_(ss),
        timeZone: Session.getScriptTimeZone(),
        reminderDays: getReminderDays_(),
        remindersEnabled: remindersEnabled_(),
        mandatoryMailTriggerInstalled: triggerNames.indexOf('sendDailyMandatoryAmcMail') >= 0,
        retryTriggerInstalled: triggerNames.indexOf('retryDailyMandatoryAmcMail') >= 0,
        expiryReminderTriggerInstalled: triggerNames.indexOf('checkAmcExpiry') >= 0
      }, callback);
    }

    throw new Error('Unsupported GET action. Use action=get or action=health.');
  } catch (err) {
    return jsonResponse_({ ok: false, error: String(err && err.message || err) }, callback);
  }
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  let locked = false;

  try {
    lock.waitLock(25000);
    locked = true;

    const raw = e && e.postData && e.postData.contents ? e.postData.contents : '{}';
    const body = JSON.parse(raw);
    if (!body || typeof body !== 'object' || Array.isArray(body)) {
      throw new Error('Invalid JSON request body.');
    }

    const ss = getDb_();
    Object.keys(SHEET_HEADERS).forEach(name => ensureSheet_(ss, name, SHEET_HEADERS[name]));

    if (body.action === 'saveAdvisory') {
      if (!body.advisory || !body.advisory.id) {
        throw new Error('Advisory payload or stable advisory ID is missing.');
      }
      const sheet = ss.getSheetByName('Engineering Advisories');
      const merged = mergeRecords_(readJsonRows_(sheet), [body.advisory], 'id');
      writeJsonRows_(sheet, merged, 'id');
      return jsonResponse_({
        ok: true,
        action: 'saveAdvisory',
        savedId: String(body.advisory.id),
        count: merged.length
      });
    }

    if (body.action !== 'sync') {
      throw new Error('Unsupported POST action: ' + String(body.action || '(missing)'));
    }
    if (!Array.isArray(body.contracts)) {
      throw new Error('Sync rejected: contracts must be an array. Existing cloud data has not been changed.');
    }
    if (!Array.isArray(body.assets) || !Array.isArray(body.logs)) {
      throw new Error('Sync rejected: assets and logs must both be arrays.');
    }

    const assetsSheet = ss.getSheetByName('Equipment Register');
    const contractsSheet = ss.getSheetByName('AMC Contracts');
    const logsSheet = ss.getSheetByName('Maintenance Logs');

    const mergedAssets = mergeRecords_(readJsonRows_(assetsSheet), body.assets, 'id');
    const mergedContracts = mergeRecords_(readJsonRows_(contractsSheet), body.contracts, 'id');
    const mergedLogs = mergeRecords_(readJsonRows_(logsSheet), body.logs, 'id');

    const deletedLogIds = new Set(
      (Array.isArray(body.deletedLogIds) ? body.deletedLogIds : []).map(String)
    );
    const finalLogs = deletedLogIds.size
      ? mergedLogs.filter(row => !deletedLogIds.has(String(row.id || '')))
      : mergedLogs;

    // Keep the row arrays merged by stable identity and updatedAt. This prevents a
    // stale browser tab from replacing a more recently updated record.
    writeJsonRows_(assetsSheet, mergedAssets, 'id');
    writeJsonRows_(contractsSheet, mergedContracts, 'id');
    writeJsonRows_(logsSheet, finalLogs, 'id');

    const plantData = body.plantData && typeof body.plantData === 'object' ? body.plantData : {};
    if (Array.isArray(plantData.engineeringAdvisories)) {
      const advisorySheet = ss.getSheetByName('Engineering Advisories');
      const mergedAdvisories = mergeRecords_(
        readJsonRows_(advisorySheet),
        plantData.engineeringAdvisories,
        'id'
      );
      writeJsonRows_(advisorySheet, mergedAdvisories, 'id');
    }

    SYNCABLE_PLANT_KEYS.forEach(key => {
      if (Array.isArray(plantData[key])) {
        const existing = readPlantValue_(ss, key);
        const merged = mergePlantRows_(key, existing, plantData[key]);
        savePlantData_(ss, key, merged);
      }
    });

    if (body.reminderPrefs && typeof body.reminderPrefs === 'object') {
      if (body.reminderPrefs.days !== undefined) {
        const days = parseReminderDays_(body.reminderPrefs.days);
        PropertiesService.getScriptProperties().setProperty(REMINDER_DAYS_KEY, JSON.stringify(days));
      }
      if (body.reminderPrefs.enabled !== undefined) {
        const enabled = body.reminderPrefs.enabled === true ||
          String(body.reminderPrefs.enabled).toLowerCase() === 'true';
        PropertiesService.getScriptProperties().setProperty(REMINDERS_ENABLED_KEY, enabled ? 'true' : 'false');
      }
    }

    const storedAssetCount = readJsonRows_(assetsSheet).length;
    const storedContractCount = readJsonRows_(contractsSheet).length;
    const storedLogCount = readJsonRows_(logsSheet).length;
    const syncTimestamp = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss');

    saveSyncMeta_(ss, {
      requestId: String(body.requestId || ''),
      syncedAt: syncTimestamp,
      assets: storedAssetCount,
      contracts: storedContractCount,
      logs: storedLogCount,
      clientVersion: String(body.clientVersion || 'unspecified'),
      source: 'SMART MAINTAIN website sync'
    });

    return jsonResponse_({
      ok: true,
      action: 'sync',
      requestId: String(body.requestId || ''),
      assets: storedAssetCount,
      contracts: storedContractCount,
      logs: storedLogCount,
      syncMeta: getSyncMeta_(ss)
    });
  } catch (err) {
    console.error('SMART MAINTAIN POST failed: ' + String(err && err.stack || err));
    return jsonResponse_({
      ok: false,
      error: String(err && err.message || err)
    });
  } finally {
    if (locked) lock.releaseLock();
  }
}

/* --------------------------- Data storage --------------------------- */

function rowTimestamp_(row) {
  return Date.parse(row && row.updatedAt ? row.updatedAt : '') || 0;
}

function mergeRecords_(existingRows, incomingRows, keyField) {
  const map = new Map();
  const add = row => {
    if (!row || row[keyField] === undefined || row[keyField] === null || String(row[keyField]).trim() === '') return;
    const key = String(row[keyField]);
    const previous = map.get(key);
    if (!previous) {
      map.set(key, row);
      return;
    }
    // A strict greater-than comparison makes the already stored version win on
    // equal or missing timestamps. Normal website edits stamp updatedAt.
    if (rowTimestamp_(row) > rowTimestamp_(previous)) map.set(key, row);
  };

  (Array.isArray(existingRows) ? existingRows : []).forEach(add);
  (Array.isArray(incomingRows) ? incomingRows : []).forEach(add);
  return Array.from(map.values());
}

function plantRowKey_(key, row, index) {
  if (!row || typeof row !== 'object') return 'row:' + index;
  // Prefer the natural key for these collections so older rows without an ID
  // merge with newer rows after the website assigns stable IDs.
  if (key === 'team' && row.name) return 'name:' + String(row.name).trim().toLowerCase();
  if (key === 'initiatives' && row.title) {
    return 'title:' + String(row.title).trim().toLowerCase() + '|' + String(row.date || '');
  }
  if (key === 'training' && row.title) {
    return 'title:' + String(row.title).trim().toLowerCase() + '|' + String(row.date || '');
  }
  if (key === 'plantGallery' && row.caption) {
    return 'caption:' + String(row.caption).trim().toLowerCase() + '|' + String(row.date || '');
  }
  if (row.id !== undefined && row.id !== null && String(row.id).trim()) return 'id:' + String(row.id);
  if (row.name) return 'name:' + String(row.name).trim().toLowerCase();
  if (row.title) return 'title:' + String(row.title).trim().toLowerCase();
  return 'row:' + JSON.stringify(row);
}

function mergePlantRows_(key, existingRows, incomingRows) {
  const map = new Map();
  const add = (row, index) => {
    if (!row || typeof row !== 'object') return;
    const stableKey = plantRowKey_(key, row, index);
    const previous = map.get(stableKey);
    if (!previous) {
      map.set(stableKey, row);
      return;
    }
    if (rowTimestamp_(row) > rowTimestamp_(previous)) map.set(stableKey, row);
  };
  (Array.isArray(existingRows) ? existingRows : []).forEach(add);
  (Array.isArray(incomingRows) ? incomingRows : []).forEach(add);
  return Array.from(map.values());
}

function readJsonRows_(sheet) {
  if (!sheet || sheet.getLastRow() < 2) return [];
  const values = sheet.getRange(2, 2, sheet.getLastRow() - 1, 1).getValues();
  return values.map(row => {
    try { return JSON.parse(row[0]); } catch (_) { return null; }
  }).filter(row => row && typeof row === 'object');
}

function writeJsonRows_(sheet, rows, keyField) {
  const safeRows = Array.isArray(rows) ? rows.filter(row =>
    row && row[keyField] !== undefined && row[keyField] !== null && String(row[keyField]).trim() !== ''
  ) : [];
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) sheet.getRange(2, 1, lastRow - 1, 2).clearContent();
  if (!safeRows.length) return;
  const values = safeRows.map(row => [String(row[keyField]), JSON.stringify(row)]);
  sheet.getRange(2, 1, values.length, 2).setValues(values);
}

function readPlantData_(ss) {
  const data = {};
  const sheet = ss.getSheetByName('Plant Data');
  if (sheet && sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues().forEach(row => {
      const key = String(row[0] || '');
      if (!key) return;
      try { data[key] = JSON.parse(row[1]); } catch (_) { data[key] = row[1]; }
    });
  }
  data.engineeringAdvisories = readJsonRows_(ss.getSheetByName('Engineering Advisories'));
  return data;
}

function readPlantValue_(ss, key) {
  const sheet = ss.getSheetByName('Plant Data');
  if (!sheet || sheet.getLastRow() < 2) return [];
  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0]) === String(key)) {
      try { return JSON.parse(values[i][1]); } catch (_) { return values[i][1]; }
    }
  }
  return [];
}

function savePlantData_(ss, key, value) {
  const sheet = ensureSheet_(ss, 'Plant Data', SHEET_HEADERS['Plant Data']);
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    const keys = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
    const found = keys.findIndex(value => String(value) === String(key));
    if (found >= 0) {
      sheet.getRange(found + 2, 2).setValue(JSON.stringify(value));
      return;
    }
  }
  sheet.appendRow([String(key), JSON.stringify(value)]);
}

function getSyncMeta_(ss) {
  const sheet = ss.getSheetByName('Sync Status');
  const stored = {};
  if (sheet && sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues().forEach(row => {
      if (row[0] !== '' && row[0] != null) stored[String(row[0])] = row[1];
    });
  }
  return {
    requestId: String(stored.requestId || ''),
    updatedAt: String(stored.syncedAt || ''),
    assetCount: Number(stored.assets || 0),
    contractCount: Number(stored.contracts || 0),
    logCount: Number(stored.logs || 0),
    clientVersion: String(stored.clientVersion || ''),
    source: String(stored.source || 'No website sync recorded')
  };
}

function saveSyncMeta_(ss, metadata) {
  const sheet = ensureSheet_(ss, 'Sync Status', SHEET_HEADERS['Sync Status']);
  const rows = Object.keys(metadata).map(key => [key, metadata[key]]);
  if (sheet.getLastRow() > 1) sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).clearContent();
  if (rows.length) sheet.getRange(2, 1, rows.length, 2).setValues(rows);
}

/* --------------------------- AMC date and reminder helpers --------------------------- */

function parseReminderDays_(value) {
  const source = Array.isArray(value) ? value : String(value || '').split(',');
  const numbers = source.map(item => Number(String(item).trim()));
  if (!numbers.length || numbers.some(n => !Number.isInteger(n) || n < 0 || n > 3650)) {
    throw new Error('Reminder days must be whole numbers from 0 to 3650.');
  }
  return Array.from(new Set(numbers)).sort((a, b) => b - a);
}

function getReminderDays_() {
  const value = PropertiesService.getScriptProperties().getProperty(REMINDER_DAYS_KEY);
  if (!value) return REMINDER_DAYS_DEFAULT.slice();
  try { return parseReminderDays_(JSON.parse(value)); }
  catch (_) { return REMINDER_DAYS_DEFAULT.slice(); }
}

function remindersEnabled_() {
  return PropertiesService.getScriptProperties().getProperty(REMINDERS_ENABLED_KEY) !== 'false';
}

function parseExpiryDate_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, TZ, 'yyyy-MM-dd');
  }
  const raw = String(value || '').trim();
  if (!raw) return null;

  let match = raw.match(/^(\d{4})-(\d{1,2})-(\d{1,2})(?:$|T|\s)/);
  if (match) {
    return match[1] + '-' + String(match[2]).padStart(2, '0') + '-' + String(match[3]).padStart(2, '0');
  }
  match = raw.match(/^(\d{1,2})[-\/](\d{1,2})[-\/](\d{4})$/);
  if (match) {
    return match[3] + '-' + String(match[2]).padStart(2, '0') + '-' + String(match[1]).padStart(2, '0');
  }
  return null;
}

function dayDifference_(dateString, todayString) {
  const target = Date.parse(dateString + 'T00:00:00Z');
  const today = Date.parse(todayString + 'T00:00:00Z');
  return Math.round((target - today) / 86400000);
}

function countVisitsForContract_(logs, contractId) {
  return (Array.isArray(logs) ? logs : []).filter(log =>
    log && (log.type === 'AMC Visit' || log.type === 'Service Visit') &&
    String(log.contract || '') === String(contractId || '')
  ).length;
}

function visitsAgreed_(contract) {
  if (contract.totalVisits === '' || contract.totalVisits == null) return null;
  const count = Number(contract.totalVisits);
  return Number.isFinite(count) && count >= 0 ? count : null;
}

function formatContractStatus_(record) {
  if (!record.expiry || record.expiry === 'Not specified') return 'EXPIRY NOT ENTERED';
  if (record.days < 0) return 'EXPIRED';
  if (record.days === 0) return 'EXPIRES TODAY';
  return 'ACTIVE';
}

/* --------------------------- Expiry reminder mail --------------------------- */

function checkAmcExpiry() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) {
    Logger.log('AMC expiry reminder: another script execution holds the lock. This run will be skipped.');
    return;
  }

  try {
    if (!remindersEnabled_()) {
      Logger.log('AMC expiry reminders are disabled by the website preference.');
      return;
    }

    const ss = getDb_();
    const contracts = readJsonRows_(ss.getSheetByName('AMC Contracts'));
    const logs = readJsonRows_(ss.getSheetByName('Maintenance Logs'));
    const emailLog = ensureSheet_(ss, 'Email Log', SHEET_HEADERS['Email Log']);
    const today = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
    const reminderDays = getReminderDays_();
    const sentKeys = new Set();

    if (emailLog.getLastRow() > 1) {
      // Columns: timestamp, contract ID, contract ref, expiry date, reminder day, recipients.
      emailLog.getRange(2, 2, emailLog.getLastRow() - 1, 4).getValues().forEach(row => {
        sentKeys.add([row[0], row[2], row[3]].join('|'));
      });
    }

    const due = [];
    contracts.forEach(contract => {
      if (!contract || !contract.id) return;
      const expiry = parseExpiryDate_(contract.end);
      if (!expiry) return;
      const days = dayDifference_(expiry, today);

      // Before expiry: configured dates only. On expiry day and after: daily.
      if (days > 0 && reminderDays.indexOf(days) < 0) return;
      const key = [contract.id, expiry, days].join('|');
      if (sentKeys.has(key)) return;

      const visitsGiven = Math.max(
        Number(contract.visitsGivenBaseline || 0),
        countVisitsForContract_(logs, contract.id)
      );
      const visitsAgreed = visitsAgreed_(contract);

      due.push({
        id: contract.id,
        ref: contract.ref || contract.id,
        equipment: contract.equipment || 'Not specified',
        oem: contract.oem || contract.agency || 'Not specified',
        start: parseExpiryDate_(contract.start) || 'Not specified',
        expiry: expiry,
        days: days,
        po: contract.po || '',
        ola: contract.ola || '',
        visitsAgreed: visitsAgreed,
        visitsGiven: visitsGiven,
        visitsRemaining: visitsAgreed === null ? null : Math.max(0, visitsAgreed - visitsGiven)
      });
    });

    if (!due.length) {
      Logger.log('No AMC reminders due today (' + today + '). Contracts checked: ' + contracts.length);
      return;
    }

    const lines = [
      'Dear Team Gonda,',
      '',
      'SMART MAINTAIN – AMC EXPIRY REMINDER',
      'Gonda LPG Plant | ' + today,
      'The following AMC contract(s) have reached a configured reminder date.',
      ''
    ];

    due.forEach((item, index) => {
      lines.push((index + 1) + '. Equipment: ' + item.equipment);
      lines.push('   Contract reference: ' + item.ref);
      lines.push('   OEM / Agency: ' + item.oem);
      lines.push('   Validity start: ' + item.start);
      lines.push('   Exact validity end: ' + item.expiry);
      lines.push('   Days to expiry: ' + item.days);
      if (item.po) lines.push('   PO: ' + item.po);
      if (item.ola) lines.push('   OLA: ' + item.ola);
      if (item.visitsAgreed !== null) {
        lines.push('   Visits agreed / given / remaining: ' +
          item.visitsAgreed + ' / ' + item.visitsGiven + ' / ' + item.visitsRemaining);
      }
      lines.push('');
    });

    lines.push('Please review the contract and arrange the required action.', '', 'Regards,', 'SMART MAINTAIN – Gonda LPG Plant');

    const quota = MailApp.getRemainingDailyQuota();
    if (quota < RECIPIENTS.length) {
      throw new Error('Insufficient email recipient quota. Remaining quota: ' + quota);
    }

    MailApp.sendEmail({
      to: RECIPIENTS.join(','),
      subject: 'SMART MAINTAIN | AMC expiry reminder | Gonda LPG Plant',
      body: lines.join('\n'),
      name: 'SMART MAINTAIN – Gonda LPG Plant'
    });

    const sentAt = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss');
    const rows = due.map(item => [sentAt, item.id, item.ref, item.expiry, item.days, RECIPIENTS.join(', ')]);
    emailLog.getRange(emailLog.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);

    Logger.log('SUCCESS: AMC expiry reminder sent. Contracts checked: ' + contracts.length + '; reminders sent for: ' + due.length);
  } catch (err) {
    console.error('AMC expiry reminder failed: ' + String(err && err.stack || err));
    throw err;
  } finally {
    lock.releaseLock();
  }
}

/* --------------------------- Mandatory daily AMC report --------------------------- */

function sendDailyMandatoryAmcMail() {
  return sendDailyMandatoryAmcMail_(false);
}

function retryDailyMandatoryAmcMail() {
  return sendDailyMandatoryAmcMail_(true);
}

function sendDailyMandatoryAmcMail_(isRetry) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) {
    Logger.log('Mandatory AMC report: another execution is active; the retry trigger will try again.');
    return;
  }

  try {
    const props = PropertiesService.getScriptProperties();
    const today = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
    const sentProperty = 'SMART_MAINTAIN_MANDATORY_AMC_MAIL_SENT_' + today;

    if (props.getProperty(sentProperty) === 'true') {
      Logger.log('Mandatory AMC status report already sent successfully today: ' + today);
      return;
    }

    const ss = getDb_();
    const contracts = readJsonRows_(ss.getSheetByName('AMC Contracts'));
    const logs = readJsonRows_(ss.getSheetByName('Maintenance Logs'));
    const emailLog = ensureSheet_(ss, 'Email Log', SHEET_HEADERS['Email Log']);
    const syncMeta = getSyncMeta_(ss);

    const records = contracts.map(contract => {
      const expiry = parseExpiryDate_(contract.end);
      const days = expiry ? dayDifference_(expiry, today) : null;
      const visitsGiven = Math.max(
        Number(contract.visitsGivenBaseline || 0),
        countVisitsForContract_(logs, contract.id)
      );
      const visitsAgreed = visitsAgreed_(contract);
      return {
        id: contract.id || '',
        ref: contract.ref || contract.id || '',
        equipment: contract.equipment || 'Not specified',
        oem: contract.oem || contract.agency || 'Not specified',
        start: parseExpiryDate_(contract.start) || 'Not specified',
        expiry: expiry || 'Not specified',
        days: days,
        po: contract.po || '',
        ola: contract.ola || '',
        visitsAgreed: visitsAgreed,
        visitsGiven: visitsGiven,
        visitsRemaining: visitsAgreed === null ? null : Math.max(0, visitsAgreed - visitsGiven)
      };
    });

    const active = records.filter(row => row.expiry !== 'Not specified' && row.days >= 0);
    const expired = records.filter(row => row.expiry !== 'Not specified' && row.days < 0);
    const noExpiry = records.filter(row => row.expiry === 'Not specified');

    const sorted = records.slice().sort((a, b) => {
      if (a.expiry === 'Not specified' && b.expiry !== 'Not specified') return 1;
      if (b.expiry === 'Not specified' && a.expiry !== 'Not specified') return -1;
      if (a.expiry === 'Not specified') return String(a.equipment).localeCompare(String(b.equipment));
      return a.expiry.localeCompare(b.expiry);
    });

    const lines = [
      'Dear Team Gonda,',
      '',
      'MANDATORY DAILY AMC STATUS REPORT – GONDA LPG PLANT',
      'Report date: ' + today,
      'Scheduled issue: Every day at approximately 8:00 AM IST',
      'Last website sync: ' + (syncMeta.updatedAt || 'NEVER / NOT RECORDED'),
      'Last sync request ID: ' + (syncMeta.requestId || 'None'),
      'Last sync contract count: ' + syncMeta.contractCount,
      'Data source: SMART MAINTAIN shared Google Sheets database',
      '',
      'SUMMARY',
      'Total registered AMC contracts: ' + records.length,
      'Active / not yet expired: ' + active.length,
      'Expired contracts: ' + expired.length,
      'Contracts with missing / invalid expiry date: ' + noExpiry.length,
      '',
      'REGISTERED AMC EQUIPMENT – EXPIRY DATE LIST',
      'Equipment | AMC / Contract Reference | Valid Till | Status'
    ];

    if (!records.length) {
      lines.push('WARNING: ZERO AMC contracts are currently present in the shared database.');
      lines.push('This indicates that the website AMC register has not yet synchronized or the register is empty.');
      lines.push('Last successful website sync: ' + (syncMeta.updatedAt || 'NEVER'));
      lines.push('ACTION: Open SMART MAINTAIN, confirm that the AMC Management page contains the required records, and use Sync AMC records now.');
    } else {
      sorted.forEach((row, index) => {
        lines.push(
          (index + 1) + '. ' + row.equipment + ' | ' + row.ref + ' | ' + row.expiry + ' | ' +
          formatContractStatus_(row)
        );
      });

      lines.push('', 'FULL REGISTERED AMC CONTRACT DETAILS');
      records.forEach((row, index) => {
        lines.push('');
        lines.push((index + 1) + '. Equipment: ' + row.equipment);
        lines.push('   Contract / AMC reference: ' + row.ref);
        lines.push('   OEM / Agency: ' + row.oem);
        lines.push('   Valid from: ' + row.start);
        lines.push('   Valid till: ' + row.expiry);
        if (row.days !== null) {
          lines.push('   Days to expiry: ' + row.days + (row.days < 0 ? ' (EXPIRED)' : row.days === 0 ? ' (EXPIRES TODAY)' : ''));
        }
        if (row.po) lines.push('   PO: ' + row.po);
        if (row.ola) lines.push('   OLA: ' + row.ola);
        if (row.visitsAgreed !== null) {
          lines.push('   Visits agreed / recorded / remaining: ' +
            row.visitsAgreed + ' / ' + row.visitsGiven + ' / ' + row.visitsRemaining);
        }
      });
    }

    lines.push('', 'Please review the AMC position and take the necessary action wherever required.', '', 'Regards,', 'SMART MAINTAIN – Gonda LPG Plant');

    const quota = MailApp.getRemainingDailyQuota();
    if (quota < RECIPIENTS.length) {
      throw new Error('Insufficient email recipient quota. Remaining quota: ' + quota);
    }

    MailApp.sendEmail({
      to: RECIPIENTS.join(','),
      subject: 'MANDATORY DAILY AMC STATUS | ' + today + ' | Gonda LPG Plant',
      body: lines.join('\n'),
      name: 'SMART MAINTAIN – Gonda LPG Plant'
    });

    const sentAt = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss');
    emailLog.appendRow([
      sentAt,
      'DAILY-MANDATORY-REPORT',
      'Daily AMC Status',
      today,
      'MANDATORY DAILY',
      RECIPIENTS.join(', ')
    ]);

    props.setProperty(sentProperty, 'true');
    props.setProperty('SMART_MAINTAIN_MANDATORY_AMC_LAST_SUCCESS', today);
    props.setProperty('SMART_MAINTAIN_MANDATORY_AMC_LAST_SENT_AT', sentAt);
    props.deleteProperty('SMART_MAINTAIN_MANDATORY_AMC_LAST_FAILURE');

    Logger.log(
      'SUCCESS: mandatory daily AMC status report sent. Registered contracts: ' +
      records.length + '; retry execution: ' + Boolean(isRetry)
    );
  } catch (err) {
    const failure = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss') +
      ' | ' + String(err && err.message || err);
    PropertiesService.getScriptProperties().setProperty('SMART_MAINTAIN_MANDATORY_AMC_LAST_FAILURE', failure);
    console.error('Mandatory daily AMC email failed: ' + String(err && err.stack || err));
    throw err;
  } finally {
    lock.releaseLock();
  }
}

/* --------------------------- Diagnostics and authorisation --------------------------- */

function sendTestEmail() {
  const quota = MailApp.getRemainingDailyQuota();
  if (quota < RECIPIENTS.length) {
    throw new Error('Insufficient email recipient quota. Remaining quota: ' + quota);
  }

  MailApp.sendEmail({
    to: RECIPIENTS.join(','),
    subject: 'SMART MAINTAIN – TEST EMAIL – Gonda LPG Plant',
    body: [
      'Dear Team Gonda,',
      '',
      'This is a delivery test from SMART MAINTAIN – Gonda LPG Plant.',
      'If this email is received, MailApp is authorised for the configured recipient list.',
      '',
      'Regards,',
      'SMART MAINTAIN – Gonda LPG Plant'
    ].join('\n'),
    name: 'SMART MAINTAIN – Gonda LPG Plant'
  });

  Logger.log('Test email sent to: ' + RECIPIENTS.join(', '));
}

function diagnoseAmcSync() {
  const ss = getDb_();
  const contracts = readJsonRows_(ss.getSheetByName('AMC Contracts'));
  const meta = getSyncMeta_(ss);

  Logger.log('Shared spreadsheet: ' + ss.getUrl());
  Logger.log('Cloud AMC contracts currently stored: ' + contracts.length);
  Logger.log('Last website sync: ' + (meta.updatedAt || 'NEVER'));
  Logger.log('Last sync request ID: ' + (meta.requestId || 'None'));
  Logger.log('Last website sync contract count: ' + meta.contractCount);

  contracts.forEach((contract, index) => {
    Logger.log(
      (index + 1) + '. ' + (contract.equipment || 'Equipment not specified') +
      ' | ' + (contract.ref || contract.id || '') +
      ' | Expiry: ' + (contract.end || 'Not entered')
    );
  });
}

function diagnoseSmartMaintain() {
  const ss = getDb_();
  const props = PropertiesService.getScriptProperties();
  const today = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
  const triggers = ScriptApp.getProjectTriggers().map(trigger => trigger.getHandlerFunction());
  const contracts = readJsonRows_(ss.getSheetByName('AMC Contracts'));
  const meta = getSyncMeta_(ss);

  Logger.log('Spreadsheet: ' + ss.getUrl());
  Logger.log('Equipment rows: ' + readJsonRows_(ss.getSheetByName('Equipment Register')).length);
  Logger.log('AMC contracts: ' + contracts.length);
  Logger.log('Maintenance logs: ' + readJsonRows_(ss.getSheetByName('Maintenance Logs')).length);
  Logger.log('Last sync timestamp: ' + (meta.updatedAt || 'NEVER'));
  Logger.log('Last sync request ID: ' + (meta.requestId || 'None'));
  Logger.log('Last sync contract count: ' + meta.contractCount);
  Logger.log('Script time zone: ' + Session.getScriptTimeZone());
  Logger.log('Configured time zone: ' + TZ);
  Logger.log('Mandatory daily mail trigger installed: ' + (triggers.indexOf('sendDailyMandatoryAmcMail') >= 0));
  Logger.log('Retry trigger installed: ' + (triggers.indexOf('retryDailyMandatoryAmcMail') >= 0));
  Logger.log('Expiry reminder trigger installed: ' + (triggers.indexOf('checkAmcExpiry') >= 0));
  Logger.log('Mandatory report sent today: ' + (props.getProperty('SMART_MAINTAIN_MANDATORY_AMC_MAIL_SENT_' + today) === 'true'));
  Logger.log('Last mandatory mail success: ' + (props.getProperty('SMART_MAINTAIN_MANDATORY_AMC_LAST_SUCCESS') || 'None'));
  Logger.log('Last mandatory mail failure: ' + (props.getProperty('SMART_MAINTAIN_MANDATORY_AMC_LAST_FAILURE') || 'None'));
  Logger.log('Configured reminder days: ' + JSON.stringify(getReminderDays_()));
  Logger.log('Expiry reminders enabled: ' + remindersEnabled_());
  Logger.log('Remaining daily MailApp quota: ' + MailApp.getRemainingDailyQuota());
}
