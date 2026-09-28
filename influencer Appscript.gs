/**
 * Connectly — Google Apps Script backend
 * --------------------------------------
 * Influencers sheet ID (existing):
 *   12KG3BbkclwCiRI-DtoiikukJMvmoOpa450Vj1LU9SfQ
 * Campaigns sheet ID (Zudi Brand Campaign):
 *   1iaV1C07yqjzH7ppvzfOW_0EmAcY9_dcIflSdkbreQBI
 *
 * doGet:
 *   - default / no action  → influencer CSV
 *   - ?action=campaigns    → campaign CSV
 * doPost:
 *   - action: "join"       → append influencer
 *   - action: "campaign"   → append campaign
 */

var SHEET_ID = '12KG3BbkclwCiRI-DtoiikukJMvmoOpa450Vj1LU9SfQ';
var CAMPAIGN_SHEET_ID = '1iaV1C07yqjzH7ppvzfOW_0EmAcY9_dcIflSdkbreQBI';
var CAMPAIGN_SHEET_NAME = 'Zudi Brand Campaign';

// ---------- Campaign helpers ----------
function getCampaignSheet_() {
  var ss = SpreadsheetApp.openById(CAMPAIGN_SHEET_ID);
  var sh = ss.getSheetByName(CAMPAIGN_SHEET_NAME);
  if (!sh) {
    // fallback: first sheet if name not found
    sh = ss.getSheets()[0];
  }
  return sh;
}

function campaignsToCsv_() {
  var sh = getCampaignSheet_();
  var data = sh.getDataRange().getValues();
  if (!data.length) return '';
  return data.map(function(row) {
    return row.map(function(cell) {
      var s = (cell === null || cell === undefined) ? '' : String(cell);
      if (s.indexOf(',') >= 0 || s.indexOf('"') >= 0 || s.indexOf('\n') >= 0) {
        return '"' + s.replace(/"/g, '""') + '"';
      }
      return s;
    }).join(',');
  }).join('\n');
}

function appendCampaign_(obj) {
  var sh = getCampaignSheet_();
  sh.appendRow([
    obj.campaignId || '',
    obj.brandName || '',
    obj.title || '',
    obj.brief || '',
    obj.niche || '',
    obj.platform || '',
    obj.targetAudience || '',
    obj.location || '',
    obj.followerRange || '',
    obj.budget || '',
    obj.deliverables || '',
    obj.deadline || '',
    obj.email || '',
    obj.website || '',
    obj.status || 'Active',
    obj.createdDate || '',
    obj.createdTime || ''
  ]);
}

// ---------- Web app entry points ----------
function doGet(e) {
  e = e || { parameter: {} };
  var action = (e.parameter && e.parameter.action) || '';

  // Campaigns list for Open Campaigns page
  if (action === 'campaigns') {
    return ContentService
      .createTextOutput(campaignsToCsv_())
      .setMimeType(ContentService.MimeType.CSV);
  }

  // Default: influencer/creators CSV (unchanged)
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheets()[0];
  var data = sheet.getDataRange().getValues();

  var csv = data.map(function(row) {
    return row.map(function(cell) {
      var s = (cell === null || cell === undefined) ? '' : String(cell);
      if (s.indexOf(',') >= 0 || s.indexOf('"') >= 0 || s.indexOf('\n') >= 0) {
        return '"' + s.replace(/"/g, '""') + '"';
      }
      return s;
    }).join(',');
  }).join('\n');

  return ContentService
    .createTextOutput(csv)
    .setMimeType(ContentService.MimeType.CSV);
}

function doPost(e) {
  try {
    var body = {};
    if (e && e.postData && e.postData.contents) {
      body = JSON.parse(e.postData.contents);
    }

    // Brand campaign → Zudi Brand Campaign sheet
    if (body.action === 'campaign') {
      appendCampaign_(body);
      return jsonOut({ ok: true, message: 'Campaign added' });
    }

    // Creator join → influencer sheet (unchanged)
    if (body.action && body.action !== 'join') {
      return jsonOut({ ok: false, error: 'Unknown action' });
    }

    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheets()[0];

    var now = new Date();
    var timestamp = Utilities.formatDate(
      now,
      Session.getScriptTimeZone() || 'Asia/Kolkata',
      'dd/MM/yyyy HH:mm:ss'
    );

    var row = [
      timestamp,
      body.age || '',
      body.name || '',
      body.instagram || '',
      body.phone || '',
      body.city || '',
      body.followers || '',
      body.skill || '',
      body.income || '',
      body.working || '',
      body.email || '',
      body.gender || '',
      body.charges || '',
      '',
      'New (website form)',
      'Via Connectly website'
    ];

    sheet.appendRow(row);

    return jsonOut({ ok: true, message: 'Creator added' });
  } catch (err) {
    return jsonOut({ ok: false, error: String(err) });
  }
}

function jsonOut(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
