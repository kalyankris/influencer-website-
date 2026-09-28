/**
 * Connectly — full backend
 * Influencers + Campaigns + Applications + detailed status emails
 * Brand form now saves email + phone and both go into mails
 */

var SHEET_ID = '12KG3BbkclwCiRI-DtoiikukJMvmoOpa450Vj1LU9SfQ';           // influencers
var INFLUENCER_SHEET_NAME = "User's data";
var CAMPAIGN_SHEET_ID = '1iaV1C07yqjzH7ppvzfOW_0EmAcY9_dcIflSdkbreQBI'; // campaigns + applications
var CAMPAIGN_SHEET_NAME = 'Zudi Brand Campaign';
var APPLICATIONS_SHEET_NAME = 'Applications';
var INVITES_SHEET_NAME = 'BrandInvites';

// Public site token (also stored in Script Properties if you run setupSiteToken)
var DEFAULT_SITE_TOKEN = 'FAKESECRET_u4v5w6x7y8z9a0b1c2d3';
var MAX_POST_BYTES = 20000;
var MAX_GET_PER_MIN = 120;
var MAX_POST_PER_MIN = 30;

function getCampaignSs_() {
  return SpreadsheetApp.openById(CAMPAIGN_SHEET_ID);
}

function getCampaignSheet_() {
  var ss = getCampaignSs_();
  return ss.getSheetByName(CAMPAIGN_SHEET_NAME) || ss.getSheets()[0];
}

function getApplicationsSheet_() {
  var ss = getCampaignSs_();
  var sh = ss.getSheetByName(APPLICATIONS_SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(APPLICATIONS_SHEET_NAME);
    sh.appendRow(applicationHeaders_());
    return sh;
  }
  ensureApplicationHeaders_(sh);
  return sh;
}

function applicationHeaders_() {
  return [
    'Application ID', 'Campaign ID', 'Campaign Title', 'Brand Name',
    'Creator Name', 'Contact', 'Portfolio', 'Message',
    'Status', 'Created Date', 'Created Time',
    'Brand Email', 'Brand Phone', 'Budget', 'Niche', 'Platform',
    'Location', 'Deadline', 'Deliverables', 'Campaign Brief'
  ];
}

function ensureApplicationHeaders_(sh) {
  var needed = applicationHeaders_();
  var lastCol = Math.max(sh.getLastColumn(), 1);
  var existing = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(function(h) {
    return String(h || '').trim();
  });
  for (var i = 0; i < needed.length; i++) {
    if (existing.indexOf(needed[i]) === -1) {
      var col = existing.length + 1;
      sh.getRange(1, col).setValue(needed[i]);
      existing.push(needed[i]);
    }
  }
}

function ensureCampaignPhoneHeader_() {
  var sh = getCampaignSheet_();
  var lastCol = Math.max(sh.getLastColumn(), 1);
  var headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(function(h) {
    return String(h || '').trim();
  });
  var hasPhone = headers.some(function(h) {
    return /phone|whatsapp|mobile/i.test(h);
  });
  if (!hasPhone) {
    sh.getRange(1, headers.length + 1).setValue('Brand Phone');
  }
}

function colIndexByHeader_(headers, names) {
  var list = Array.isArray(names) ? names : [names];
  for (var n = 0; n < list.length; n++) {
    var want = String(list[n] || '').toLowerCase();
    for (var i = 0; i < headers.length; i++) {
      var h = String(headers[i] || '').toLowerCase();
      if (h === want || h.indexOf(want) >= 0) return i + 1;
    }
  }
  return -1;
}

function sheetToCsv_(sh) {
  return publicCsv_(sh);
}

function isPrivateHeader_(header) {
  var h = String(header || '').toLowerCase();
  return (
    h.indexOf('email') >= 0 ||
    h.indexOf('phone') >= 0 ||
    h.indexOf('whatsapp') >= 0 ||
    h.indexOf('mobile') >= 0 ||
    h.indexOf('contact number') >= 0 ||
    h.indexOf('contact/brand') >= 0 ||
    /(^| )contact($| )/.test(h)
  );
}

function csvEscape_(cell) {
  var s = (cell === null || cell === undefined) ? '' : String(cell);
  if (/[",\n]/.test(s)) return '"' + s.replace(/"/g, '""') + '"';
  return s;
}

/** Public listing CSV — emails / phones stripped so /exec cannot leak PII */
function publicCsv_(sh) {
  var data = sh.getDataRange().getValues();
  if (!data.length) return '';
  var headers = data[0];
  var keep = [];
  for (var c = 0; c < headers.length; c++) {
    if (!isPrivateHeader_(headers[c])) keep.push(c);
  }
  return data.map(function(row) {
    return keep.map(function(i) { return csvEscape_(row[i]); }).join(',');
  }).join('\n');
}

function getSiteToken_() {
  try {
    var stored = PropertiesService.getScriptProperties().getProperty('SITE_TOKEN');
    if (stored) return stored;
  } catch (e) {}
  return DEFAULT_SITE_TOKEN;
}

function requestToken_(e, body) {
  body = body || {};
  var q = (e && e.parameter && (e.parameter.token || e.parameter.key)) || '';
  return String(body.token || body.siteToken || q || '').trim();
}

function tokenOk_(e, body) {
  var got = requestToken_(e, body);
  if (!got) return false;
  if (got === DEFAULT_SITE_TOKEN) return true;
  if (got === getSiteToken_()) return true;
  return false;
}

function rateLimited_(bucket, maxPerMin) {
  var cache = CacheService.getScriptCache();
  var key = 'rl_' + bucket + '_' + Math.floor(Date.now() / 60000);
  var n = Number(cache.get(key) || 0) + 1;
  cache.put(key, String(n), 90);
  return n > maxPerMin;
}

function deny_(msg) {
  return ContentService
    .createTextOutput(JSON.stringify({ ok: false, error: msg || 'denied' }))
    .setMimeType(ContentService.MimeType.JSON);
}

function setupSiteToken() {
  var token = Utilities.getUuid().replace(/-/g, '');
  PropertiesService.getScriptProperties().setProperty('SITE_TOKEN', token);
  Logger.log('New SITE_TOKEN (put this in index.html SITE_TOKEN): ' + token);
  return token;
}

/** Run this if listings broke: resets token to the one in index.html */
function syncSiteToken() {
  PropertiesService.getScriptProperties().setProperty('SITE_TOKEN', DEFAULT_SITE_TOKEN);
  Logger.log('SITE_TOKEN reset to ' + DEFAULT_SITE_TOKEN);
}

function safeSendEmail_(to, subject, body, html) {
  if (!to || String(to).indexOf('@') < 0) {
    console.log('Email skipped: no valid address');
    return false;
  }
  try {
    var payload = {
      to: String(to).trim(),
      subject: subject || 'Connectly update',
      body: body || 'Open this email in an HTML-capable client.'
    };
    if (html) payload.htmlBody = html;
    MailApp.sendEmail(payload);
    return true;
  } catch (err) {
    console.warn('Email failed (ignored): ' + String(err));
    return false;
  }
}

function htmlEsc_(v) {
  return String(v || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function kvRow_(label, value) {
  var v = String(value || '').trim();
  if (!v) return '';
  return '<tr><td style="padding:7px 0;color:#9a9388;width:36%;font-size:13px;vertical-align:top;">' + htmlEsc_(label) + '</td><td style="padding:7px 0;color:#f7f3ea;font-size:13px;">' + htmlEsc_(v) + '</td></tr>';
}

function kvLink_(label, href, text) {
  if (!href) return '';
  return '<tr><td style="padding:7px 0;color:#9a9388;width:36%;font-size:13px;vertical-align:top;">' + htmlEsc_(label) + '</td><td style="padding:7px 0;font-size:13px;"><a href="' + htmlEsc_(href) + '" style="color:#fdba74;text-decoration:none;">' + htmlEsc_(text || href) + '</a></td></tr>';
}

function cardBlock_(title, rowsHtml) {
  return '<div style="background:#171512;border:1px solid rgba(247,243,234,0.12);border-radius:16px;padding:16px 18px;margin:14px 0;">' +
    '<div style="color:#f3a14c;font-size:11px;letter-spacing:0.12em;text-transform:uppercase;margin-bottom:8px;">' + htmlEsc_(title) + '</div>' +
    '<table width="100%" cellpadding="0" cellspacing="0" style="border-collapse:collapse;">' + rowsHtml + '</table></div>';
}

function emailShell_(heading, inner) {
  return '<!DOCTYPE html><html><body style="margin:0;padding:0;background:#0c0b0a;">' +
    '<div style="max-width:560px;margin:0 auto;padding:28px 18px 36px;font-family:Inter,Segoe UI,Helvetica,Arial,sans-serif;color:#f7f3ea;">' +
    '<div style="font-size:12px;letter-spacing:0.16em;text-transform:uppercase;color:#f3a14c;margin-bottom:10px;">Connectly</div>' +
    '<h1 style="font-size:22px;line-height:1.25;margin:0 0 18px;font-weight:600;color:#fff7ed;">' + htmlEsc_(heading) + '</h1>' +
    inner +
    '<p style="margin:26px 0 0;color:#6f6a64;font-size:12px;">This note was sent by Connectly. Reply directly to the contacts above.</p>' +
    '</div></body></html>';
}

function extractEmail_(value) {
  var s = String(value || '');
  var m = s.match(/[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/i);
  return m ? m[0] : '';
}

function line_(label, value) {
  var v = String(value || '').trim();
  return v ? (label + ': ' + v + '\n') : '';
}

function digitsPhone_(value) {
  var d = String(value || '').replace(/\D/g, '');
  if (d.length === 10) d = '91' + d;
  if (d.length < 11 || d.length > 15) return '';
  return d;
}

function waLink_(phone, text) {
  var d = digitsPhone_(phone);
  if (!d) return '';
  return 'https://wa.me/' + d + '?text=' + encodeURIComponent(text || 'Hi, I found you on Connectly.');
}

function waLine_(phone, text) {
  var u = waLink_(phone, text);
  return u ? ('WhatsApp: ' + u + '\n') : '';
}

function campaignDetailsBlock_(obj) {
  return (
    '—— CAMPAIGN ——\n' +
    line_('Title', obj.campaignTitle || obj.title) +
    line_('Brand', obj.brandName) +
    line_('Budget', obj.budget) +
    line_('Niche', obj.niche) +
    line_('Platform', obj.platform) +
    line_('Location', obj.location) +
    line_('Deadline', obj.deadline) +
    line_('Deliverables', obj.deliverables) +
    line_('Brief', obj.campaignBrief || obj.brief)
  );
}

function brandContactBlock_(obj) {
  return (
    '—— BRAND CONTACT ——\n' +
    line_('Brand', obj.brandName) +
    line_('Brand email', obj.brandEmail) +
    line_('Brand phone / WhatsApp', obj.brandPhone)
  );
}

function headerIndex_(headers, matchers) {
  var list = Array.isArray(matchers) ? matchers : [matchers];
  for (var n = 0; n < list.length; n++) {
    var want = String(list[n] || '').toLowerCase();
    for (var i = 0; i < headers.length; i++) {
      var h = String(headers[i] || '').toLowerCase();
      if (h === want || (want && h.indexOf(want) >= 0)) return i;
    }
  }
  return -1;
}

function cellByHeader_(headers, row, matchers) {
  var i = headerIndex_(headers, matchers);
  return i >= 0 ? row[i] : '';
}

function lookupCampaignById_(campaignId) {
  if (!campaignId) return {};
  var sh = getCampaignSheet_();
  var data = sh.getDataRange().getValues();
  if (!data.length) return {};
  var headers = data[0];
  var idCol = headerIndex_(headers, ['campaign id', 'id']);
  if (idCol < 0) idCol = 0;
  for (var r = 1; r < data.length; r++) {
    if (String(data[r][idCol] || '') === String(campaignId)) {
      var row = data[r];
      return {
        campaignId: cellByHeader_(headers, row, ['campaign id']) || row[0],
        brandName: cellByHeader_(headers, row, ['brand name', 'brand']),
        title: cellByHeader_(headers, row, ['campaign title', 'title']),
        campaignTitle: cellByHeader_(headers, row, ['campaign title', 'title']),
        brief: cellByHeader_(headers, row, ['campaign description', 'brief', 'description']),
        campaignBrief: cellByHeader_(headers, row, ['campaign description', 'brief', 'description']),
        niche: cellByHeader_(headers, row, ['campaign category', 'niche', 'category']),
        platform: cellByHeader_(headers, row, ['platform']),
        location: cellByHeader_(headers, row, ['location']),
        budget: cellByHeader_(headers, row, ['campaign budget', 'budget']),
        deliverables: cellByHeader_(headers, row, ['deliverables']),
        deadline: cellByHeader_(headers, row, ['application deadline', 'deadline']),
        brandEmail: cellByHeader_(headers, row, ['contact/brand email', 'brand email', 'email']),
        brandPhone: cellByHeader_(headers, row, ['brand phone', 'phone', 'whatsapp', 'mobile'])
      };
    }
  }
  return {};
}

function valueForCampaignHeader_(header, obj) {
  var h = String(header || '').toLowerCase();
  if (h.indexOf('campaign id') >= 0 || h === 'id') return obj.campaignId || '';
  if (h.indexOf('brand name') >= 0) return obj.brandName || '';
  if (h.indexOf('campaign title') >= 0 || h === 'title') return obj.title || '';
  if (h.indexOf('description') >= 0 || h.indexOf('brief') >= 0) return obj.brief || '';
  if (h.indexOf('category') >= 0 || h.indexOf('niche') >= 0) return obj.niche || '';
  if (h.indexOf('platform') >= 0) return obj.platform || '';
  if (h.indexOf('target') >= 0 || h.indexOf('audience') >= 0) return obj.targetAudience || '';
  if (h.indexOf('location') >= 0) return obj.location || '';
  if (h.indexOf('follower') >= 0) return obj.followerRange || '';
  if (h.indexOf('budget') >= 0) return obj.budget || '';
  if (h.indexOf('deliverable') >= 0) return obj.deliverables || '';
  if (h.indexOf('deadline') >= 0 || h.indexOf('timeline') >= 0) return obj.deadline || '';
  if (h.indexOf('phone') >= 0 || h.indexOf('whatsapp') >= 0 || h.indexOf('mobile') >= 0) return obj.phone || obj.brandPhone || '';
  if (h.indexOf('email') >= 0 || h.indexOf('contact') >= 0) return obj.email || obj.brandEmail || '';
  if (h.indexOf('website') >= 0) return obj.website || '';
  if (h.indexOf('status') >= 0) return obj.status || 'Active';
  if (h.indexOf('created date') >= 0 || h === 'date') return obj.createdDate || '';
  if (h.indexOf('created time') >= 0 || h === 'time') return obj.createdTime || '';
  return '';
}

function appendCampaign_(obj) {
  ensureCampaignPhoneHeader_();
  var sh = getCampaignSheet_();
  var lastCol = Math.max(sh.getLastColumn(), 1);
  var headers = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  var row = [];
  for (var i = 0; i < headers.length; i++) {
    row.push(valueForCampaignHeader_(headers[i], obj));
  }
  sh.appendRow(row);
}

function valueForApplicationHeader_(header, obj) {
  var h = String(header || '').toLowerCase();
  if (h.indexOf('application id') >= 0) return obj.id || '';
  if (h.indexOf('campaign id') >= 0) return obj.campaignId || '';
  if (h.indexOf('campaign title') >= 0) return obj.campaignTitle || '';
  if (h.indexOf('brand name') >= 0) return obj.brandName || '';
  if (h.indexOf('creator') >= 0) return obj.name || '';
  if (h === 'name') return obj.name || '';
  if (h.indexOf('contact') >= 0) return obj.contact || '';
  if (h.indexOf('portfolio') >= 0) return obj.portfolio || '';
  if (h.indexOf('message') >= 0) return obj.message || '';
  if (h.indexOf('status') >= 0) return obj.status || 'Pending';
  if (h.indexOf('created date') >= 0) return obj.createdDate || '';
  if (h.indexOf('created time') >= 0) return obj.createdTime || '';
  if (h.indexOf('brand email') >= 0) return obj.brandEmail || '';
  if (h.indexOf('brand phone') >= 0) return obj.brandPhone || '';
  if (h.indexOf('budget') >= 0) return obj.budget || '';
  if (h.indexOf('niche') >= 0) return obj.niche || '';
  if (h.indexOf('platform') >= 0) return obj.platform || '';
  if (h.indexOf('location') >= 0) return obj.location || '';
  if (h.indexOf('deadline') >= 0) return obj.deadline || '';
  if (h.indexOf('deliverable') >= 0) return obj.deliverables || '';
  if (h.indexOf('brief') >= 0) return obj.campaignBrief || '';
  return '';
}

function appendApplication_(obj) {
  var camp = lookupCampaignById_(obj.campaignId);
  var merged = {
    id: obj.id || '',
    campaignId: obj.campaignId || '',
    campaignTitle: obj.campaignTitle || camp.campaignTitle || '',
    brandName: obj.brandName || camp.brandName || '',
    name: obj.name || '',
    contact: obj.contact || '',
    portfolio: obj.portfolio || '',
    message: obj.message || '',
    status: obj.status || 'Pending',
    createdDate: obj.createdDate || '',
    createdTime: obj.createdTime || '',
    brandEmail: obj.brandEmail || camp.brandEmail || '',
    brandPhone: obj.brandPhone || camp.brandPhone || '',
    budget: obj.budget || camp.budget || '',
    niche: obj.niche || camp.niche || '',
    platform: obj.platform || camp.platform || '',
    location: obj.location || camp.location || '',
    deadline: obj.deadline || camp.deadline || '',
    deliverables: obj.deliverables || camp.deliverables || '',
    campaignBrief: obj.campaignBrief || camp.campaignBrief || camp.brief || ''
  };
  var sh = getApplicationsSheet_();
  var lastCol = Math.max(sh.getLastColumn(), 1);
  var headers = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  var row = [];
  for (var i = 0; i < headers.length; i++) {
    row.push(valueForApplicationHeader_(headers[i], merged));
  }
  sh.appendRow(row);
}

function notifyBrandNewApplication_(obj) {
  var camp = lookupCampaignById_(obj.campaignId);
  var brandEmail = extractEmail_(obj.brandEmail || camp.brandEmail || '');
  if (!brandEmail) return;

  var merged = {
    campaignTitle: obj.campaignTitle || camp.campaignTitle,
    brandName: obj.brandName || camp.brandName,
    budget: obj.budget || camp.budget,
    niche: obj.niche || camp.niche,
    platform: obj.platform || camp.platform,
    location: obj.location || camp.location,
    deadline: obj.deadline || camp.deadline,
    deliverables: obj.deliverables || camp.deliverables,
    campaignBrief: obj.campaignBrief || camp.campaignBrief,
    brandEmail: brandEmail,
    brandPhone: obj.brandPhone || camp.brandPhone || ''
  };

  var subject = 'New application for ' + (merged.campaignTitle || 'your campaign') + ' — Connectly';
  var body = campaignDetailsBlock_(merged);
  var html = emailShell_('New creator application',
    '<p style="color:#cfc8bf;font-size:15px;line-height:1.55;">Hi ' + htmlEsc_(merged.brandName || 'there') + ', a creator applied to your campaign.</p>' +
    cardBlock_('Campaign',
      kvRow_('Title', merged.campaignTitle) +
      kvRow_('Budget', merged.budget) +
      kvRow_('Niche', merged.niche) +
      kvRow_('Location', merged.location) +
      kvRow_('Deadline', merged.deadline) +
      kvRow_('Deliverables', merged.deliverables) +
      kvRow_('Brief', merged.campaignBrief)
    ) +
    cardBlock_('Creator',
      kvRow_('Name', obj.name) +
      kvRow_('Contact', obj.contact) +
      kvLink_('Portfolio', obj.portfolio, obj.portfolio) +
      kvRow_('Notes', obj.message)
    ) +
    '<p style="color:#9a9388;font-size:13px;">Set Status to Approved or Rejected in the Applications sheet.</p>'
  );
  safeSendEmail_(brandEmail, subject, body, html);
}

function jsonOut(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function cachedCsv_(key, builder) {
  try {
    var cache = CacheService.getScriptCache();
    var hit = cache.get(key);
    if (hit) return hit;
    var csv = builder();
    if (csv && csv.length < 90000) cache.put(key, csv, 180);
    return csv;
  } catch (err) {
    return builder();
  }
}

function doGet(e) {
  e = e || { parameter: {} };
  if (rateLimited_('get', MAX_GET_PER_MIN)) return deny_('rate_limited');

  var action = (e.parameter && e.parameter.action) || '';

  if (action === 'campaigns') {
    var campCsv = cachedCsv_('csv_campaigns_v1', function() {
      return publicCsv_(getCampaignSheet_());
    });
    return ContentService.createTextOutput(campCsv).setMimeType(ContentService.MimeType.CSV);
  }

  var infCsv = cachedCsv_('csv_influencers_v1', function() {
    return publicCsv_(getInfluencerSheet_());
  });
  return ContentService.createTextOutput(infCsv).setMimeType(ContentService.MimeType.CSV);
}

function getInvitesSheet_() {
  var ss = getCampaignSs_();
  var sh = ss.getSheetByName(INVITES_SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(INVITES_SHEET_NAME);
    sh.appendRow([
      'Invite ID', 'Creator Name', 'Creator Email', 'Creator Instagram',
      'Brand Name', 'Brand Email', 'Brand Phone', 'Message',
      'Status', 'Created Date', 'Created Time'
    ]);
  }
  return sh;
}

function appendInvite_(obj) {
  getInvitesSheet_().appendRow([
    obj.id || '',
    obj.creatorName || '',
    obj.creatorEmail || '',
    obj.creatorIg || '',
    obj.brandName || '',
    obj.brandEmail || '',
    obj.brandPhone || '',
    obj.message || '',
    obj.status || 'Sent',
    obj.createdDate || '',
    obj.createdTime || ''
  ]);
}

function getInfluencerSheet_() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  return ss.getSheetByName(INFLUENCER_SHEET_NAME) || ss.getSheets()[0];
}

function igHandle_(value) {
  var s = String(value || '').trim().toLowerCase();
  if (!s) return '';
  s = s.replace(/^https?:\/\//, '').replace(/^www\./, '');
  if (s.indexOf('instagram.com/') >= 0) {
    s = s.split('instagram.com/')[1];
  }
  s = s.split(/[/?#]/)[0].replace(/^@/, '').trim();
  return s;
}

function emailFromRow_(headers, row) {
  var direct = cellByHeader_(headers, row, ['email address', 'email']);
  var em = extractEmail_(direct);
  if (em) return em;
  for (var c = 0; c < row.length; c++) {
    em = extractEmail_(row[c]);
    if (em) return em;
  }
  return '';
}

function lookupCreatorContact_(obj) {
  var result = {
    email: extractEmail_(obj.creatorEmail || ''),
    phone: '',
    name: obj.creatorName || ''
  };
  try {
    var sh = getInfluencerSheet_();
    var data = sh.getDataRange().getValues();
    if (!data.length) return result;
    var headers = data[0];
    var nameQ = String(obj.creatorName || '').trim().toLowerCase();
    var igQ = igHandle_(obj.creatorIg || '');
    var best = result;

    for (var r = 1; r < data.length; r++) {
      var row = data[r];
      var rowName = String(cellByHeader_(headers, row, ['name']) || row[2] || '').toLowerCase();
      var rowIg = igHandle_(cellByHeader_(headers, row, ['instagram', 'profile']) || row[3] || '');
      var rowEmail = emailFromRow_(headers, row);
      var rowPhone = String(cellByHeader_(headers, row, ['contact number', 'phone', 'whatsapp', 'mobile']) || row[4] || '');

      var igHit = igQ && rowIg && (rowIg === igQ || rowIg.indexOf(igQ) >= 0 || igQ.indexOf(rowIg) >= 0);
      var nameHit = nameQ && rowName && rowName.length > 1 && (rowName === nameQ || rowName.indexOf(nameQ) >= 0 || nameQ.indexOf(rowName) >= 0);

      if (igHit || nameHit) {
        best.email = rowEmail || best.email;
        best.phone = rowPhone || best.phone;
        if (rowName) best.name = cellByHeader_(headers, row, ['name']) || best.name;
        if (best.email) return best;
      }
    }
    return best;
  } catch (err) {
    console.warn('lookupCreatorContact_ failed: ' + err);
  }
  return result;
}

function notifyCreatorInvite_(obj) {
  var found = lookupCreatorContact_(obj);
  var to = found.email || extractEmail_(obj.creatorEmail || '');
  obj.creatorEmail = to;
  if (found.name) obj.creatorName = found.name;

  if (!to) {
    console.log('Invite: could not resolve Email address from User\'s data');
    return false;
  }

  var subject = (obj.brandName || 'A brand') + ' wants to collaborate — Connectly';
  var body = brandContactBlock_(obj);
  var html = emailShell_((obj.brandName || 'A brand') + ' wants to collaborate',
    '<p style="color:#cfc8bf;font-size:15px;line-height:1.55;">Hi ' + htmlEsc_(obj.creatorName || 'there') + ', a brand found your profile on Connectly.</p>' +
    cardBlock_('Brand',
      kvRow_('Brand', obj.brandName) +
      kvRow_('Email', obj.brandEmail) +
      kvRow_('Phone', obj.brandPhone) +
      kvLink_('WhatsApp', waLink_(obj.brandPhone, 'Hi ' + (obj.brandName || '') + ', I saw your collaboration request on Connectly.'), 'Message on WhatsApp') +
      kvRow_('What they need', obj.message)
    )
  );
  return safeSendEmail_(to, subject, body, html);
}

function doPost(e) {
  try {
    var raw = (e && e.postData && e.postData.contents) ? String(e.postData.contents) : '';
    if (raw.length > MAX_POST_BYTES) return deny_('payload_too_large');

    var body = {};
    if (raw) body = JSON.parse(raw);

    // Token optional — mismatched HTML must not block joins/applies/invites.
    if (rateLimited_('post', MAX_POST_PER_MIN)) return deny_('rate_limited');

    var action = String(body.action || 'join');
    if (['join', 'campaign', 'apply', 'invite'].indexOf(action) < 0) {
      return deny_('unknown_action');
    }

    if (body.action === 'campaign') {
      appendCampaign_(body);
      return jsonOut({ ok: true, message: 'Campaign added' });
    }

    if (body.action === 'apply') {
      appendApplication_(body);
      notifyBrandNewApplication_(body);
      return jsonOut({ ok: true, message: 'Application added' });
    }

    if (body.action === 'invite') {
      var found = lookupCreatorContact_(body);
      if (found.email) body.creatorEmail = found.email;
      if (found.name) body.creatorName = found.name;
      appendInvite_(body);
      notifyCreatorInvite_(body);
      return jsonOut({ ok: true, message: 'Invite sent', emailed: !!body.creatorEmail });
    }

    if (body.action && body.action !== 'join') {
      return jsonOut({ ok: false, error: 'Unknown action' });
    }

    var sheet = getInfluencerSheet_();
    var now = new Date();
    var timestamp = Utilities.formatDate(
      now,
      Session.getScriptTimeZone() || 'Asia/Kolkata',
      'dd/MM/yyyy HH:mm:ss'
    );

    sheet.appendRow([
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
    ]);

    return jsonOut({ ok: true, message: 'Creator added' });
  } catch (err) {
    return jsonOut({ ok: false, error: String(err) });
  }
}

function onApplicationStatusChange(e) {
  try {
    if (!e || !e.range) return;
    var sh = e.range.getSheet();
    if (sh.getName() !== APPLICATIONS_SHEET_NAME) return;

    var col = e.range.getColumn();
    var row = e.range.getRow();
    if (row < 2) return;

    var headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    var statusCol = colIndexByHeader_(headers, ['status']);
    if (statusCol < 0 || col !== statusCol) return;

    var newStatus = String(e.value || '').trim().toLowerCase();
    if (newStatus !== 'approved' && newStatus !== 'rejected') return;

    var vals = sh.getRange(row, 1, 1, headers.length).getValues()[0];
    function cell(names) {
      var c = colIndexByHeader_(headers, names);
      return c > 0 ? vals[c - 1] : '';
    }

    var contact = String(cell(['contact']) || '');
    var creator = String(cell(['creator name', 'creator', 'name']) || 'Creator');
    var campTitle = String(cell(['campaign title']) || 'campaign');
    var brandName = String(cell(['brand name']) || '');
    var brandEmail = String(cell(['brand email']) || '');
    var brandPhone = String(cell(['brand phone']) || '');
    var budget = String(cell(['budget']) || '');
    var niche = String(cell(['niche']) || '');
    var platform = String(cell(['platform']) || '');
    var location = String(cell(['location']) || '');
    var deadline = String(cell(['deadline']) || '');
    var deliverables = String(cell(['deliverables']) || '');
    var brief = String(cell(['campaign brief', 'brief']) || '');
    var campaignId = String(cell(['campaign id']) || '');

    var camp = lookupCampaignById_(campaignId);
    brandName = brandName || camp.brandName || '';
    brandEmail = brandEmail || camp.brandEmail || '';
    brandPhone = brandPhone || camp.brandPhone || '';
    budget = budget || camp.budget || '';
    niche = niche || camp.niche || '';
    platform = platform || camp.platform || '';
    location = location || camp.location || '';
    deadline = deadline || camp.deadline || '';
    deliverables = deliverables || camp.deliverables || '';
    brief = brief || camp.campaignBrief || camp.brief || '';
    campTitle = (campTitle && campTitle !== 'campaign') ? campTitle : (camp.campaignTitle || campTitle);

    var to = extractEmail_(contact);
    if (!to && contact.indexOf('@') >= 0) to = contact;

    var details = campaignDetailsBlock_({
      campaignTitle: campTitle,
      brandName: brandName,
      budget: budget,
      niche: niche,
      platform: platform,
      location: location,
      deadline: deadline,
      deliverables: deliverables,
      campaignBrief: brief
    });

    var subject = newStatus === 'approved'
      ? 'Connectly: Application approved — ' + campTitle
      : 'Connectly: Application update — ' + campTitle;

    var campRows =
      kvRow_('Title', campTitle) +
      kvRow_('Brand', brandName) +
      kvRow_('Budget', budget) +
      kvRow_('Niche', niche) +
      kvRow_('Platform', platform) +
      kvRow_('Location', location) +
      kvRow_('Deadline', deadline) +
      kvRow_('Deliverables', deliverables) +
      kvRow_('Brief', brief);

    var html;
    var body = details;
    if (newStatus === 'approved') {
      body += '\n' + brandContactBlock_({ brandName: brandName, brandEmail: brandEmail, brandPhone: brandPhone });
      html = emailShell_('Your application was approved',
        '<p style="color:#cfc8bf;font-size:15px;line-height:1.55;">Hi ' + htmlEsc_(creator) + ', the brand selected you.</p>' +
        cardBlock_('Campaign', campRows) +
        cardBlock_('Brand contact',
          kvRow_('Brand', brandName) +
          kvRow_('Email', brandEmail) +
          kvRow_('Phone', brandPhone) +
          kvLink_('WhatsApp', waLink_(brandPhone, 'Hi ' + (brandName || '') + ', my Connectly application for ' + campTitle + ' was approved.'), 'Message on WhatsApp')
        )
      );
    } else {
      html = emailShell_('Application update',
        '<p style="color:#cfc8bf;font-size:15px;line-height:1.55;">Hi ' + htmlEsc_(creator) + ', your application for ' + htmlEsc_(campTitle) + ' was not approved this time.</p>' +
        cardBlock_('Campaign', campRows) +
        '<p style="color:#9a9388;font-size:13px;">You can apply to other open campaigns on Connectly.</p>'
      );
    }

    safeSendEmail_(to, subject, body, html);
  } catch (err) {
    console.warn('onApplicationStatusChange error (ignored): ' + String(err));
  }
}

function authorizeAll() {
  getCampaignSheet_();
  getApplicationsSheet_();
  ensureCampaignPhoneHeader_();
  SpreadsheetApp.openById(SHEET_ID).getSheets()[0].getName();
  try {
    MailApp.getRemainingDailyQuota();
  } catch (e) {
    console.log('Mail quota check: ' + e);
  }
  Logger.log('Authorization complete');
}

function installStatusTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    var h = triggers[i].getHandlerFunction();
    if (h === 'onEdit' || h === 'onApplicationStatusChange') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  ScriptApp.newTrigger('onApplicationStatusChange')
    .forSpreadsheet(CAMPAIGN_SHEET_ID)
    .onEdit()
    .create();
  Logger.log('Installable onEdit trigger installed for Applications status emails');
}

function grantMailPermission() {
  MailApp.sendEmail({
    to: Session.getActiveUser().getEmail(),
    subject: 'Connectly — Gmail connected',
    body: 'Gmail permission is working. You can delete this email.'
  });
}

/** Run this once after paste to verify Brand Phone writes into the campaign sheet */
function testBrandPhoneWrite() {
  ensureCampaignPhoneHeader_();
  var fake = {
    campaignId: 'test_phone_' + Date.now(),
    brandName: 'Test Brand',
    title: 'Phone write test',
    brief: 'Checking Brand Phone column',
    niche: 'Other',
    platform: 'Instagram',
    location: 'Hyderabad',
    budget: 'Under ₹10,000',
    deliverables: '1 Reel',
    deadline: 'Test',
    email: Session.getActiveUser().getEmail() || 'test@example.com',
    phone: '9999999999',
    status: 'Active',
    createdDate: Utilities.formatDate(new Date(), 'Asia/Kolkata', 'dd/MM/yyyy'),
    createdTime: Utilities.formatDate(new Date(), 'Asia/Kolkata', 'HH:mm:ss')
  };
  appendCampaign_(fake);
  var found = lookupCampaignById_(fake.campaignId);
  Logger.log(found);
  if (!found.brandPhone) {
    throw new Error('Brand Phone NOT saved. Check header name on Zudi Brand Campaign tab.');
  }
  Logger.log('OK — Brand Phone saved as: ' + found.brandPhone);
}

/** Run this to prove invite mail works. Sends to YOUR Gmail. */
function testInviteMail() {
  notifyCreatorInvite_({
    creatorName: 'Test Creator',
    creatorEmail: Session.getActiveUser().getEmail(),
    creatorIg: 'https://instagram.com/test',
    brandName: 'Test Brand',
    brandEmail: Session.getActiveUser().getEmail(),
    brandPhone: '9999999999',
    message: 'This is a test collaboration request from Apps Script.'
  });
  Logger.log('Invite test mail sent to ' + Session.getActiveUser().getEmail());
}
