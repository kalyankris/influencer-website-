
/**
 * Connectly — COMPLETE single backend (final)
 * Influencers + Campaigns + Applications + Deals + Invites + Emails
 * Anti-spam email headers + clean human timeline messages
 */

var SHEET_ID = '12KG3BbkclwCiRI-DtoiikukJMvmoOpa450Vj1LU9SfQ';
var INFLUENCER_SHEET_NAME = "User's data";
var CAMPAIGN_SHEET_ID = '1iaV1C07yqjzH7ppvzfOW_0EmAcY9_dcIflSdkbreQBI';
var CAMPAIGN_SHEET_NAME = 'Zudi Brand Campaign';
var APPLICATIONS_SHEET_NAME = 'Applications';
var INVITES_SHEET_NAME = 'BrandInvites';
var DEALS_SHEET_NAME = 'Deals';

// ========== EDIT ONLY THESE 3 LINES ==========
var CONNECTLY_UPI = 'kalyankris92-4@oksbi';
var CONNECTLY_PHONE = '9177791543';
var CONNECTLY_OPS_EMAIL = 'kalyankris92@gmail.com';
// =============================================

var DEFAULT_SITE_TOKEN = 'FAKESECRET_u4v5w6x7y8z9a0b1c2d3';
var MAX_POST_BYTES = 20000;
var MAX_GET_PER_MIN = 120;
var MAX_POST_PER_MIN = 30;

function getCampaignSs_() { return SpreadsheetApp.openById(CAMPAIGN_SHEET_ID); }
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
    'Application ID','Campaign ID','Campaign Title','Brand Name',
    'Creator Name','Contact','Portfolio','Message',
    'Status','Created Date','Created Time',
    'Brand Email','Brand Phone','Budget','Niche','Platform',
    'Location','Deadline','Deliverables','Campaign Brief'
  ];
}
function ensureApplicationHeaders_(sh) {
  var needed = applicationHeaders_();
  var lastCol = Math.max(sh.getLastColumn(), 1);
  var existing = sh.getRange(1,1,1,lastCol).getValues()[0].map(function(h){ return String(h||'').trim(); });
  for (var i=0; i<needed.length; i++) {
    if (existing.indexOf(needed[i]) === -1) {
      sh.getRange(1, existing.length+1).setValue(needed[i]);
      existing.push(needed[i]);
    }
  }
}
function ensureCampaignPhoneHeader_() {
  var sh = getCampaignSheet_();
  var lastCol = Math.max(sh.getLastColumn(),1);
  var headers = sh.getRange(1,1,1,lastCol).getValues()[0].map(function(h){ return String(h||'').trim(); });
  if (!headers.some(function(h){ return /phone|whatsapp|mobile/i.test(h); })) {
    sh.getRange(1, headers.length+1).setValue('Brand Phone');
  }
}
function colIndexByHeader_(headers, names) {
  var list = Array.isArray(names) ? names : [names];
  for (var n=0; n<list.length; n++) {
    var want = String(list[n]||'').toLowerCase();
    for (var i=0; i<headers.length; i++) {
      var h = String(headers[i]||'').toLowerCase();
      if (h === want || h.indexOf(want) >= 0) return i+1;
    }
  }
  return -1;
}
function isPrivateHeader_(header) {
  var h = String(header||'').toLowerCase();
  return h.indexOf('email')>=0 || h.indexOf('phone')>=0 || h.indexOf('whatsapp')>=0 ||
         h.indexOf('mobile')>=0 || h.indexOf('contact number')>=0 || h.indexOf('contact/brand')>=0 ||
         /(^| )contact($| )/.test(h);
}
function csvEscape_(cell) {
  var s = (cell===null||cell===undefined) ? '' : String(cell);
  if (/[",\n]/.test(s)) return '"' + s.replace(/"/g,'""') + '"';
  return s;
}
function publicCsv_(sh) {
  var data = sh.getDataRange().getValues();
  if (!data.length) return '';
  var headers = data[0];
  var keep = [];
  for (var c=0; c<headers.length; c++) if (!isPrivateHeader_(headers[c])) keep.push(c);
  return data.map(function(row){ return keep.map(function(i){ return csvEscape_(row[i]); }).join(','); }).join('\n');
}
function getSiteToken_() {
  try {
    var stored = PropertiesService.getScriptProperties().getProperty('SITE_TOKEN');
    if (stored) return stored;
  } catch(e){}
  return DEFAULT_SITE_TOKEN;
}
function rateLimited_(bucket, maxPerMin) {
  var cache = CacheService.getScriptCache();
  var key = 'rl_' + bucket + '_' + Math.floor(Date.now()/60000);
  var n = Number(cache.get(key)||0) + 1;
  cache.put(key, String(n), 90);
  return n > maxPerMin;
}
function deny_(msg) {
  return ContentService.createTextOutput(JSON.stringify({ok:false,error:msg||'denied'}))
    .setMimeType(ContentService.MimeType.JSON);
}
function syncSiteToken() {
  PropertiesService.getScriptProperties().setProperty('SITE_TOKEN', DEFAULT_SITE_TOKEN);
  Logger.log('SITE_TOKEN reset');
}

/**
 * Anti-spam email sender
 * - Always sends both plain text + HTML
 * - Uses a real display name "Connectly"
 * - Sets replyTo when possible
 * - Avoids spammy subjects
 */
function safeSendEmail_(to, subject, body, html, replyTo) {
  if (!to || String(to).indexOf('@') < 0) {
    console.log('Email skipped: no valid address → ' + to);
    return false;
  }
  try {
    var options = {
      name: 'Connectly',
      to: String(to).trim(),
      subject: String(subject || 'Update from Connectly').substring(0, 120),
      body: body || 'Please open this email in a mail app that supports HTML.',
      htmlBody: html || undefined
    };
    if (replyTo && String(replyTo).indexOf('@') >= 0) {
      options.replyTo = String(replyTo).trim();
    }
    // Prefer GmailApp when available (better deliverability headers)
    try {
      GmailApp.sendEmail(options.to, options.subject, options.body, {
        name: options.name,
        htmlBody: options.htmlBody,
        replyTo: options.replyTo || undefined
      });
    } catch (gErr) {
      MailApp.sendEmail(options);
    }
    return true;
  } catch (err) {
    console.warn('Email failed: ' + String(err));
    return false;
  }
}

function htmlEsc_(v) {
  return String(v||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}
function kvRow_(label, value) {
  var v = String(value||'').trim();
  if (!v) return '';
  return '<tr><td style="padding:7px 0;color:#9a9388;width:36%;font-size:13px;">'+htmlEsc_(label)+'</td><td style="padding:7px 0;color:#f7f3ea;font-size:13px;">'+htmlEsc_(v)+'</td></tr>';
}
function kvLink_(label, href, text) {
  if (!href) return '';
  return '<tr><td style="padding:7px 0;color:#9a9388;width:36%;font-size:13px;">'+htmlEsc_(label)+'</td><td style="padding:7px 0;font-size:13px;"><a href="'+htmlEsc_(href)+'" style="color:#fdba74;">'+htmlEsc_(text||href)+'</a></td></tr>';
}
function cardBlock_(title, rowsHtml) {
  return '<div style="background:#171512;border:1px solid rgba(247,243,234,0.12);border-radius:16px;padding:16px 18px;margin:14px 0;">'+
    '<div style="color:#f3a14c;font-size:11px;letter-spacing:0.12em;text-transform:uppercase;margin-bottom:8px;">'+htmlEsc_(title)+'</div>'+
    '<table width="100%" cellpadding="0" cellspacing="0">'+rowsHtml+'</table></div>';
}
function emailShell_(heading, inner) {
  return '<!DOCTYPE html><html><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width"></head>'+
    '<body style="margin:0;padding:0;background:#0c0b0a;">'+
    '<div style="max-width:560px;margin:0 auto;padding:28px 18px;font-family:Inter,Segoe UI,Helvetica,Arial,sans-serif;color:#f7f3ea;">'+
    '<div style="font-size:12px;letter-spacing:0.16em;text-transform:uppercase;color:#f3a14c;margin-bottom:10px;">Connectly</div>'+
    '<h1 style="font-size:22px;margin:0 0 18px;color:#fff7ed;font-weight:600;">'+htmlEsc_(heading)+'</h1>'+inner+
    '<p style="margin:26px 0 0;color:#6f6a64;font-size:12px;line-height:1.5;">This message was sent by Connectly, an influencer–brand marketplace in India. If this is unexpected, you can ignore it.</p>'+
    '</div></body></html>';
}
function extractEmail_(value) {
  var m = String(value||'').match(/[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/i);
  return m ? m[0] : '';
}
function line_(label, value) {
  var v = String(value||'').trim();
  return v ? (label+': '+v+'\n') : '';
}
function digitsPhone_(value) {
  var d = String(value||'').replace(/\D/g,'');
  if (d.length===10) d = '91'+d;
  if (d.length<11||d.length>15) return '';
  return d;
}
function waLink_(phone, text) {
  var d = digitsPhone_(phone);
  if (!d) return '';
  return 'https://wa.me/'+d+'?text='+encodeURIComponent(text||'Hi from Connectly');
}
function campaignDetailsBlock_(obj) {
  return '—— CAMPAIGN ——\n'+
    line_('Title', obj.campaignTitle||obj.title)+
    line_('Brand', obj.brandName)+
    line_('Budget', obj.budget)+
    line_('Niche', obj.niche)+
    line_('Platform', obj.platform)+
    line_('Location', obj.location)+
    line_('Deadline', obj.deadline)+
    line_('Deliverables', obj.deliverables)+
    line_('Brief', obj.campaignBrief||obj.brief);
}
function brandContactBlock_(obj) {
  return '—— BRAND CONTACT ——\n'+
    line_('Brand', obj.brandName)+
    line_('Brand email', obj.brandEmail)+
    line_('Brand phone', obj.brandPhone);
}
function headerIndex_(headers, matchers) {
  var list = Array.isArray(matchers) ? matchers : [matchers];
  for (var n=0; n<list.length; n++) {
    var want = String(list[n]||'').toLowerCase();
    for (var i=0; i<headers.length; i++) {
      var h = String(headers[i]||'').toLowerCase();
      if (h===want || (want && h.indexOf(want)>=0)) return i;
    }
  }
  return -1;
}
function cellByHeader_(headers, row, matchers) {
  var i = headerIndex_(headers, matchers);
  return i>=0 ? row[i] : '';
}
function lookupCampaignById_(campaignId) {
  if (!campaignId) return {};
  var sh = getCampaignSheet_();
  var data = sh.getDataRange().getValues();
  if (!data.length) return {};
  var headers = data[0];
  var idCol = headerIndex_(headers, ['campaign id','id']);
  if (idCol<0) idCol=0;
  for (var r=1; r<data.length; r++) {
    if (String(data[r][idCol]||'') === String(campaignId)) {
      var row = data[r];
      return {
        campaignId: cellByHeader_(headers,row,['campaign id'])||row[0],
        brandName: cellByHeader_(headers,row,['brand name','brand']),
        title: cellByHeader_(headers,row,['campaign title','title']),
        campaignTitle: cellByHeader_(headers,row,['campaign title','title']),
        brief: cellByHeader_(headers,row,['campaign description','brief','description']),
        campaignBrief: cellByHeader_(headers,row,['campaign description','brief','description']),
        niche: cellByHeader_(headers,row,['campaign category','niche','category']),
        platform: cellByHeader_(headers,row,['platform']),
        location: cellByHeader_(headers,row,['location']),
        budget: cellByHeader_(headers,row,['campaign budget','budget']),
        deliverables: cellByHeader_(headers,row,['deliverables']),
        deadline: cellByHeader_(headers,row,['application deadline','deadline']),
        brandEmail: cellByHeader_(headers,row,['contact/brand email','brand email','email']),
        brandPhone: cellByHeader_(headers,row,['brand phone','phone','whatsapp','mobile'])
      };
    }
  }
  return {};
}
function valueForCampaignHeader_(header, obj) {
  var h = String(header||'').toLowerCase();
  if (h.indexOf('campaign id')>=0 || h==='id') return obj.campaignId||'';
  if (h.indexOf('brand name')>=0) return obj.brandName||'';
  if (h.indexOf('campaign title')>=0 || h==='title') return obj.title||'';
  if (h.indexOf('description')>=0 || h.indexOf('brief')>=0) return obj.brief||'';
  if (h.indexOf('category')>=0 || h.indexOf('niche')>=0) return obj.niche||'';
  if (h.indexOf('platform')>=0) return obj.platform||'';
  if (h.indexOf('location')>=0) return obj.location||'';
  if (h.indexOf('budget')>=0) return obj.budget||'';
  if (h.indexOf('deliverable')>=0) return obj.deliverables||'';
  if (h.indexOf('deadline')>=0) return obj.deadline||'';
  if (h.indexOf('phone')>=0 || h.indexOf('whatsapp')>=0 || h.indexOf('mobile')>=0) return obj.phone||obj.brandPhone||'';
  if (h.indexOf('email')>=0 || h.indexOf('contact')>=0) return obj.email||obj.brandEmail||'';
  if (h.indexOf('status')>=0) return obj.status||'Active';
  if (h.indexOf('created date')>=0) return obj.createdDate||'';
  if (h.indexOf('created time')>=0) return obj.createdTime||'';
  return '';
}
function appendCampaign_(obj) {
  ensureCampaignPhoneHeader_();
  var sh = getCampaignSheet_();
  var lastCol = Math.max(sh.getLastColumn(),1);
  var headers = sh.getRange(1,1,1,lastCol).getValues()[0];
  var row = [];
  for (var i=0; i<headers.length; i++) row.push(valueForCampaignHeader_(headers[i], obj));
  sh.appendRow(row);
}
function valueForApplicationHeader_(header, obj) {
  var h = String(header||'').toLowerCase();
  if (h.indexOf('application id')>=0) return obj.id||'';
  if (h.indexOf('campaign id')>=0) return obj.campaignId||'';
  if (h.indexOf('campaign title')>=0) return obj.campaignTitle||'';
  if (h.indexOf('brand name')>=0) return obj.brandName||'';
  if (h.indexOf('creator')>=0 || h==='name') return obj.name||'';
  if (h.indexOf('contact')>=0) return obj.contact||'';
  if (h.indexOf('portfolio')>=0) return obj.portfolio||'';
  if (h.indexOf('message')>=0) return obj.message||'';
  if (h.indexOf('status')>=0) return obj.status||'Pending';
  if (h.indexOf('created date')>=0) return obj.createdDate||'';
  if (h.indexOf('created time')>=0) return obj.createdTime||'';
  if (h.indexOf('brand email')>=0) return obj.brandEmail||'';
  if (h.indexOf('brand phone')>=0) return obj.brandPhone||'';
  if (h.indexOf('budget')>=0) return obj.budget||'';
  if (h.indexOf('niche')>=0) return obj.niche||'';
  if (h.indexOf('platform')>=0) return obj.platform||'';
  if (h.indexOf('location')>=0) return obj.location||'';
  if (h.indexOf('deadline')>=0) return obj.deadline||'';
  if (h.indexOf('deliverable')>=0) return obj.deliverables||'';
  if (h.indexOf('brief')>=0) return obj.campaignBrief||'';
  return '';
}
function appendApplication_(obj) {
  var camp = lookupCampaignById_(obj.campaignId);
  var merged = {
    id: obj.id||'', campaignId: obj.campaignId||'',
    campaignTitle: obj.campaignTitle||camp.campaignTitle||'',
    brandName: obj.brandName||camp.brandName||'',
    name: obj.name||'', contact: obj.contact||'', portfolio: obj.portfolio||'',
    message: obj.message||'', status: obj.status||'Pending',
    createdDate: obj.createdDate||'', createdTime: obj.createdTime||'',
    brandEmail: obj.brandEmail||camp.brandEmail||'',
    brandPhone: obj.brandPhone||camp.brandPhone||'',
    budget: obj.budget||camp.budget||'', niche: obj.niche||camp.niche||'',
    platform: obj.platform||camp.platform||'', location: obj.location||camp.location||'',
    deadline: obj.deadline||camp.deadline||'', deliverables: obj.deliverables||camp.deliverables||'',
    campaignBrief: obj.campaignBrief||camp.campaignBrief||camp.brief||''
  };
  var sh = getApplicationsSheet_();
  var lastCol = Math.max(sh.getLastColumn(),1);
  var headers = sh.getRange(1,1,1,lastCol).getValues()[0];
  var row = [];
  for (var i=0; i<headers.length; i++) row.push(valueForApplicationHeader_(headers[i], merged));
  sh.appendRow(row);
}
function notifyBrandNewApplication_(obj) {
  var camp = lookupCampaignById_(obj.campaignId);
  var brandEmail = extractEmail_(obj.brandEmail||camp.brandEmail||'');
  if (!brandEmail) return;
  var merged = {
    campaignTitle: obj.campaignTitle||camp.campaignTitle,
    brandName: obj.brandName||camp.brandName,
    budget: obj.budget||camp.budget, niche: obj.niche||camp.niche,
    platform: obj.platform||camp.platform, location: obj.location||camp.location,
    deadline: obj.deadline||camp.deadline, deliverables: obj.deliverables||camp.deliverables,
    campaignBrief: obj.campaignBrief||camp.campaignBrief,
    brandEmail: brandEmail, brandPhone: obj.brandPhone||camp.brandPhone||''
  };
  // Soft subject — avoid spam triggers
  var subject = 'New creator applied to ' + (merged.campaignTitle || 'your campaign');
  var plain = 'Hi '+(merged.brandName||'there')+',\n\nA creator applied to your campaign on Connectly.\n\n'+
    campaignDetailsBlock_(merged)+'\n'+
    'Creator: '+(obj.name||'')+'\nContact: '+(obj.contact||'')+'\n\n'+
    'Open the Applications sheet to Approve or Reject.\n\n— Connectly';
  var html = emailShell_('New creator application',
    '<p style="color:#cfc8bf;font-size:15px;line-height:1.55;">Hi '+htmlEsc_(merged.brandName||'there')+', a creator applied to your campaign.</p>'+
    cardBlock_('Campaign', kvRow_('Title',merged.campaignTitle)+kvRow_('Budget',merged.budget)+kvRow_('Niche',merged.niche)+kvRow_('Deadline',merged.deadline))+
    cardBlock_('Creator', kvRow_('Name',obj.name)+kvRow_('Contact',obj.contact)+kvLink_('Portfolio',obj.portfolio,obj.portfolio)+kvRow_('Notes',obj.message))+
    '<p style="color:#9a9388;font-size:13px;">Set Status to Approved or Rejected in the Applications sheet.</p>'
  );
  safeSendEmail_(brandEmail, subject, plain, html, extractEmail_(obj.contact||''));
}
function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
function cachedCsv_(key, builder) {
  try {
    var cache = CacheService.getScriptCache();
    var hit = cache.get(key);
    if (hit) return hit;
    var csv = builder();
    if (csv && csv.length < 90000) cache.put(key, csv, 180);
    return csv;
  } catch(e){ return builder(); }
}

// ---------- Deals ----------
/** Unique numeric deal ID (e.g. 1728123456789) */
function makeDealId_() {
  return String(Date.now());
}

function getDealsSheet_() {
  var ss = getCampaignSs_();
  var sh = ss.getSheetByName(DEALS_SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(DEALS_SHEET_NAME);
    sh.appendRow([
      'Deal ID','Status','Creator Name','Creator Email','Creator Instagram',
      'Brand Name','Brand Email','Brand Phone','Price','Deliverable','Deadline','Brief',
      'Deliverable URL','Advance Paid','Final Paid','Payment Note','History',
      'Platform Paid','Creator Paid Out','Created Date','Created Time','Updated At'
    ]);
  } else {
    ensureDealHistoryCol_(sh);
    ensureDealExtraCols_(sh);
  }
  return sh;
}
function ensureDealHistoryCol_(sh) {
  var headers = sh.getRange(1,1,1,Math.max(sh.getLastColumn(),1)).getValues()[0];
  for (var i=0; i<headers.length; i++) if (String(headers[i]).toLowerCase()==='history') return i;
  sh.getRange(1, headers.length+1).setValue('History');
  return headers.length;
}
function ensureDealExtraCols_(sh) {
  var needed = ['Platform Paid','Creator Paid Out'];
  var lastCol = Math.max(sh.getLastColumn(),1);
  var headers = sh.getRange(1,1,1,lastCol).getValues()[0].map(function(h){ return String(h||'').trim(); });
  for (var i=0; i<needed.length; i++) {
    if (headers.indexOf(needed[i]) === -1) {
      sh.getRange(1, headers.length+1).setValue(needed[i]);
      headers.push(needed[i]);
    }
  }
}
function appendHistoryLine_(sh, row, histCol, line) {
  if (histCol < 0) return;
  var cell = sh.getRange(row, histCol+1);
  var prev = String(cell.getValue()||'').trim();
  var next = prev ? (prev + '\n' + line) : line;
  var parts = next.split('\n');
  if (parts.length > 40) next = parts.slice(-40).join('\n');
  cell.setValue(next);
}
function appendDeal_(obj) {
  var now = new Date();
  var id = obj.dealId || obj.id || makeDealId_();
  var stamp = Utilities.formatDate(now, 'Asia/Kolkata', "d MMM, h:mm a");
  var brand = obj.brandName || 'Brand';
  var brief = String(obj.message || obj.brief || '').trim();
  var price = String(obj.price || '').trim();
  var deliv = String(obj.deliverable || '').trim();
  var reqParts = [];
  if (price) reqParts.push('₹' + price);
  if (deliv) reqParts.push(deliv);
  if (brief) reqParts.push(brief);
  var reqText = reqParts.length ? reqParts.join(' · ') : 'collaboration request';
  // First timeline line = brand request (never "creator initial")
  var history = obj.history || (stamp + ' | MSG | brand | ' + brand + ' | sent collaboration request · ' + reqText);
  getDealsSheet_().appendRow([
    id, obj.status||'Proposed',
    obj.creatorName||'', obj.creatorEmail||'', obj.creatorIg||'',
    obj.brandName||'', obj.brandEmail||'', obj.brandPhone||'',
    obj.price||'', obj.deliverable||'', obj.deadline||'',
    obj.message||obj.brief||'', obj.deliverableUrl||'',
    obj.advancePaid||'No', obj.finalPaid||'No', obj.paymentNote||'',
    history, obj.platformPaid||'', obj.creatorPaidOut||'',
    obj.createdDate||Utilities.formatDate(now,'Asia/Kolkata','dd/MM/yyyy'),
    obj.createdTime||Utilities.formatDate(now,'Asia/Kolkata','HH:mm:ss'),
    now.toISOString()
  ]);
  return id;
}
function listDealsByEmail_(email) {
  email = String(email||'').trim().toLowerCase();
  if (!email || email.indexOf('@')<0) return [];
  var sh = getDealsSheet_();
  var data = sh.getDataRange().getValues();
  if (data.length < 2) return [];
  var headers = data[0];
  var out = [];
  var ops = String(CONNECTLY_OPS_EMAIL||'').trim().toLowerCase();
  for (var r=1; r<data.length; r++) {
    var row = data[r], obj = {};
    for (var c=0; c<headers.length; c++) obj[String(headers[c])] = row[c];
    var ce = String(obj['Creator Email']||'').toLowerCase();
    var be = String(obj['Brand Email']||'').toLowerCase();
    var isOps = ops && email === ops;
    if (ce===email || be===email || isOps) {
      var role = isOps ? 'ops' : ((ce===email && be===email) ? 'both' : (ce===email ? 'creator' : 'brand'));
      out.push({
        dealId: String(obj['Deal ID']||''),
        status: String(obj['Status']||''),
        creatorName: String(obj['Creator Name']||''),
        creatorEmail: String(obj['Creator Email']||''),
        creatorIg: String(obj['Creator Instagram']||''),
        brandName: String(obj['Brand Name']||''),
        brandEmail: String(obj['Brand Email']||''),
        brandPhone: String(obj['Brand Phone']||''),
        price: String(obj['Price']||''),
        deliverable: String(obj['Deliverable']||''),
        deadline: String(obj['Deadline']||''),
        brief: String(obj['Brief']||''),
        deliverableUrl: String(obj['Deliverable URL']||''),
        advancePaid: String(obj['Advance Paid']||''),
        finalPaid: String(obj['Final Paid']||''),
        paymentNote: String(obj['Payment Note']||''),
        history: String(obj['History']||''),
        platformPaid: String(obj['Platform Paid']||''),
        creatorPaidOut: String(obj['Creator Paid Out']||''),
        creatorConfirmed: /confirmed payment received/i.test(String(obj['History']||'')) ? 'Yes' :
          (/payment not received or wrong UPI/i.test(String(obj['History']||'')) ? 'Problem' : ''),
        role: role,
        connectlyUpi: CONNECTLY_UPI,
        connectlyPhone: CONNECTLY_PHONE
      });
    }
  }
  return out;
}
function updateDeal_(body) {
  var id = String(body.dealId||body.id||'').trim();
  if (!id) throw new Error('missing_deal_id');
  var sh = getDealsSheet_();
  ensureDealHistoryCol_(sh);
  ensureDealExtraCols_(sh);
  var data = sh.getDataRange().getValues();
  if (data.length < 2) throw new Error('no_deals');
  var headers = data[0], col = {};
  for (var i=0; i<headers.length; i++) col[String(headers[i]).toLowerCase()] = i;
  function C(name){ return col[name]!=null ? col[name] : -1; }
  var idCol = C('deal id');
  if (idCol < 0) throw new Error('bad_headers');
  // Human Indian-style stamp: "4 Oct, 9:36 pm"
  var stamp = Utilities.formatDate(new Date(), 'Asia/Kolkata', "d MMM, h:mm a");
  var actor = String(body.actorName||body.actor||'Someone').trim();
  var actorRole = String(body.actorRole||'').trim();

  for (var r=1; r<data.length; r++) {
    if (String(data[r][idCol]) !== id) continue;
    var row1 = r+1;
    var statusCol = C('status'), urlCol = C('deliverable url'), histCol = C('history');
    var prevStatus = statusCol>=0 ? String(data[r][statusCol]||'') : '';
    var prevUrl = urlCol>=0 ? String(data[r][urlCol]||'') : '';

    if (body.messageText) {
      appendHistoryLine_(sh, row1, histCol, stamp+' | MSG | '+actorRole+' | '+actor+' | '+String(body.messageText).replace(/\n/g,' '));
    }
    if (body.price!=null && C('price')>=0) sh.getRange(row1, C('price')+1).setValue(body.price);
    if (body.deliverable!=null && C('deliverable')>=0) sh.getRange(row1, C('deliverable')+1).setValue(body.deliverable);
    if (body.deadline!=null && C('deadline')>=0) sh.getRange(row1, C('deadline')+1).setValue(body.deadline);
    if (body.brief!=null && C('brief')>=0) sh.getRange(row1, C('brief')+1).setValue(body.brief);
    if (body.editNote) {
      appendHistoryLine_(sh, row1, histCol, stamp+' | EDIT | '+actorRole+' | '+actor+' | '+String(body.editNote).replace(/\n/g,' '));
    }

    // Brand asked for changes (revision)
    if (body.status==='Accepted' && prevStatus==='Submitted') {
      if (prevUrl) {
        appendHistoryLine_(sh, row1, histCol, stamp+' | FILE | brand | '+actor+' | Please revise the deliverable · previous file | '+prevUrl);
      } else {
        appendHistoryLine_(sh, row1, histCol, stamp+' | STATUS | brand | '+actor+' | Please revise the deliverable');
      }
      if (urlCol>=0) sh.getRange(row1, urlCol+1).setValue('');
    }

    // Creator submitted / resubmitted work
    if (body.status==='Submitted' && body.deliverableUrl) {
      var subLabel = prevStatus === 'Accepted' && /revision|asked for changes/i.test(String(data[r][histCol]||''))
        ? 'resubmitted work'
        : 'sent work to brand';
      appendHistoryLine_(sh, row1, histCol, stamp+' | FILE | creator | '+actor+' | '+subLabel+' | '+body.deliverableUrl);
      if (urlCol>=0) sh.getRange(row1, urlCol+1).setValue(body.deliverableUrl);
    }

    // Connectly rejected payment — brand must pay again
    if (body.rejectPayment || body.status==='PaymentRejected') {
      appendHistoryLine_(sh, row1, histCol, stamp+' | PAY | connectly | Connectly | payment rejected — brand must pay again');
      if (C('platform paid')>=0) sh.getRange(row1, C('platform paid')+1).setValue('No');
      if (statusCol>=0) sh.getRange(row1, statusCol+1).setValue('Accepted');
      body.status = null;
    }

    // Brand paid Connectly
    if (body.platformPaid==='Yes' || body.status==='BrandPaidPlatform') {
      var pref = String(body.paymentRef||body.paymentNote||'').trim();
      appendHistoryLine_(sh, row1, histCol, stamp+' | PAY | brand | '+actor+' | paid Connectly'+(pref?' · ref '+pref:''));
      if (C('platform paid')>=0) sh.getRange(row1, C('platform paid')+1).setValue('Yes'+(pref?' · '+pref:''));
      if (statusCol>=0 && (!body.status||body.status==='BrandPaidPlatform')) {
        sh.getRange(row1, statusCol+1).setValue('BrandPaidPlatform');
        body.status = null;
      }
    }

    // Connectly confirmed money received — clean human message
    if (body.status==='PlatformReceived') {
      appendHistoryLine_(sh, row1, histCol, stamp+' | PAY | connectly | Connectly | Payment confirmed by Connectly');
    }

    // Brand approved work
    if (body.status==='Approved') {
      appendHistoryLine_(sh, row1, histCol, stamp+' | STATUS | brand | '+actor+' | approved the work');
    }

    // Creator confirms payout received
    if (body.creatorConfirmed === 'Yes' || body.status === 'CreatorReceived') {
      appendHistoryLine_(sh, row1, histCol, stamp+' | PAY | creator | '+actor+' | confirmed payment received');
      if (statusCol>=0) sh.getRange(row1, statusCol+1).setValue('Complete');
      body.status = null;
    }
    if (body.creatorConfirmed === 'Problem' || body.status === 'CreatorNotReceived') {
      appendHistoryLine_(sh, row1, histCol, stamp+' | PAY | creator | '+actor+' | reported payment not received or wrong UPI');
      // stay Paid so ops can fix
      body.status = null;
    }

    // Connectly paid creator
    if (body.creatorPaidOut==='Yes' || body.status==='Paid') {
      var cref = String(body.payoutRef||body.paymentNote||'').trim();
      appendHistoryLine_(sh, row1, histCol, stamp+' | PAY | connectly | Connectly | paid the creator'+(cref?' · ref '+cref:''));
      if (C('creator paid out')>=0) sh.getRange(row1, C('creator paid out')+1).setValue('Yes'+(cref?' · '+cref:''));
      if (C('final paid')>=0) sh.getRange(row1, C('final paid')+1).setValue('Yes');
      if (statusCol>=0) sh.getRange(row1, statusCol+1).setValue('Paid');
      body.status = null;
    }

    if (body.status && statusCol>=0) {
      if (!(body.status==='Accepted' && prevStatus==='Submitted')) {
        if (body.status !== prevStatus && body.status !== 'Submitted') {
          appendHistoryLine_(sh, row1, histCol, stamp+' | STATUS | '+(actorRole||'user')+' | '+actor+' | '+body.status);
        }
      }
      sh.getRange(row1, statusCol+1).setValue(body.status);
    }
    if (body.deliverableUrl!=null && urlCol>=0 && body.status==='Submitted') {
      sh.getRange(row1, urlCol+1).setValue(body.deliverableUrl);
    }
    if (body.clearDeliverableUrl && urlCol>=0) sh.getRange(row1, urlCol+1).setValue('');
    if (body.advancePaid && C('advance paid')>=0) sh.getRange(row1, C('advance paid')+1).setValue(body.advancePaid);
    if (body.paymentNote!=null && C('payment note')>=0) sh.getRange(row1, C('payment note')+1).setValue(body.paymentNote);
    if (C('updated at')>=0) sh.getRange(row1, C('updated at')+1).setValue(new Date().toISOString());
    return true;
  }
  throw new Error('deal_not_found');
}

// ---------- Invites ----------
function getInvitesSheet_() {
  var ss = getCampaignSs_();
  var sh = ss.getSheetByName(INVITES_SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(INVITES_SHEET_NAME);
    sh.appendRow(['Invite ID','Creator Name','Creator Email','Creator Instagram','Brand Name','Brand Email','Brand Phone','Message','Status','Created Date','Created Time']);
  }
  return sh;
}
function appendInvite_(obj) {
  getInvitesSheet_().appendRow([
    obj.id||'', obj.creatorName||'', obj.creatorEmail||'', obj.creatorIg||'',
    obj.brandName||'', obj.brandEmail||'', obj.brandPhone||'', obj.message||'',
    obj.status||'Sent', obj.createdDate||'', obj.createdTime||''
  ]);
}
function getInfluencerSheet_() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  return ss.getSheetByName(INFLUENCER_SHEET_NAME) || ss.getSheets()[0];
}
function igHandle_(value) {
  var s = String(value||'').trim().toLowerCase();
  if (!s) return '';
  s = s.replace(/^https?:\/\//,'').replace(/^www\./,'');
  if (s.indexOf('instagram.com/')>=0) s = s.split('instagram.com/')[1];
  return s.split(/[/?#]/)[0].replace(/^@/,'').trim();
}
function emailFromRow_(headers, row) {
  var direct = cellByHeader_(headers, row, ['email address','email']);
  var em = extractEmail_(direct);
  if (em) return em;
  for (var c=0; c<row.length; c++) {
    em = extractEmail_(row[c]);
    if (em) return em;
  }
  return '';
}
function lookupCreatorContact_(obj) {
  var result = { email: extractEmail_(obj.creatorEmail||''), phone:'', name: obj.creatorName||'' };
  try {
    var sh = getInfluencerSheet_();
    var data = sh.getDataRange().getValues();
    if (!data.length) return result;
    var headers = data[0];
    var nameQ = String(obj.creatorName||'').trim().toLowerCase();
    var igQ = igHandle_(obj.creatorIg||'');
    for (var r=1; r<data.length; r++) {
      var row = data[r];
      var rowName = String(cellByHeader_(headers,row,['name'])||row[2]||'').toLowerCase();
      var rowIg = igHandle_(cellByHeader_(headers,row,['instagram','profile'])||row[3]||'');
      var rowEmail = emailFromRow_(headers, row);
      var igHit = igQ && rowIg && (rowIg===igQ || rowIg.indexOf(igQ)>=0 || igQ.indexOf(rowIg)>=0);
      var nameHit = nameQ && rowName && rowName.length>1 && (rowName===nameQ || rowName.indexOf(nameQ)>=0 || nameQ.indexOf(rowName)>=0);
      if (igHit || nameHit) {
        result.email = rowEmail || result.email;
        result.phone = String(cellByHeader_(headers,row,['contact number','phone','whatsapp','mobile'])||row[4]||'') || result.phone;
        if (rowName) result.name = cellByHeader_(headers,row,['name']) || result.name;
        if (result.email) return result;
      }
    }
  } catch(err){ console.warn(err); }
  return result;
}
function notifyCreatorInvite_(obj) {
  var found = lookupCreatorContact_(obj);
  var to = found.email || extractEmail_(obj.creatorEmail||'');
  obj.creatorEmail = to;
  if (found.name) obj.creatorName = found.name;
  if (!to) {
    console.warn('Invite email FAILED: no creator email resolved for ' + (obj.creatorName||obj.creatorIg||'unknown'));
    return false;
  }
  console.log('Invite email → ' + to);
  var brand = obj.brandName || 'A brand';
  var subject = brand + ' wants to work with you on Connectly';
  var plain = 'Hi '+(obj.creatorName||'there')+',\n\n'+brand+' found your profile on Connectly and wants to collaborate.\n\n'+
    brandContactBlock_(obj)+'\n'+
    line_('Price', obj.price)+line_('Deliverable', obj.deliverable)+line_('Deadline', obj.deadline)+
    line_('Brief', obj.message)+line_('Deal ID', obj.dealId)+
    '\nOpen Connectly → My Deals, enter your email to accept.\n\n— Connectly';
  var html = emailShell_(brand+' wants to collaborate',
    '<p style="color:#cfc8bf;font-size:15px;line-height:1.55;">Hi '+htmlEsc_(obj.creatorName||'there')+', a brand found you on Connectly.</p>'+
    cardBlock_('Offer',
      kvRow_('Brand',obj.brandName)+kvRow_('Email',obj.brandEmail)+kvRow_('Phone',obj.brandPhone)+
      kvLink_('WhatsApp',waLink_(obj.brandPhone,'Hi, I saw your request on Connectly'),'Message on WhatsApp')+
      kvRow_('Price',obj.price||'—')+kvRow_('Deliverable',obj.deliverable||'—')+
      kvRow_('Deadline',obj.deadline||'—')+kvRow_('Brief',obj.message)+
      kvRow_('Deal ID',obj.dealId||'')+
      '<p style="color:#9a9388;font-size:13px;margin-top:12px;">Open Connectly → My Deals and enter your email to accept.</p>'
    )
  );
  // Reply goes to brand so conversation can start
  return safeSendEmail_(to, subject, plain, html, extractEmail_(obj.brandEmail||''));
}

// ---------- Entry points ----------


// ========== USERS + AUTH + SHORTLIST (realtime) ==========
var USERS_SHEET_NAME = 'Users';
var SHORTLISTS_SHEET_NAME = 'Shortlists';

function getUsersSheet_() {
  var ss = getCampaignSs_();
  var sh = ss.getSheetByName(USERS_SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(USERS_SHEET_NAME);
    sh.appendRow([
      'Email', 'Role', 'Display Name', 'Brand Name', 'Website', 'Instagram',
      'Phone', 'City', 'Niche', 'Bio', 'Gender', 'Followers', 'Charges',
      'Profile Complete', 'Picture', 'Created', 'Last Login'
    ]);
  }
  return sh;
}

function getShortlistsSheet_() {
  var ss = getCampaignSs_();
  var sh = ss.getSheetByName(SHORTLISTS_SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHORTLISTS_SHEET_NAME);
    sh.appendRow(['Email', 'Creator Id', 'Creator Name', 'Creator Instagram', 'Saved At']);
  }
  return sh;
}

function userFromRow_(headers, row) {
  function g(names) {
    var i = headerIndex_(headers, names);
    return i >= 0 ? row[i] : '';
  }
  return {
    email: String(g(['email']) || '').trim().toLowerCase(),
    role: String(g(['role']) || '').trim().toLowerCase(),
    displayName: String(g(['display name', 'name']) || ''),
    brandName: String(g(['brand name']) || ''),
    website: String(g(['website']) || ''),
    instagram: String(g(['instagram']) || ''),
    phone: String(g(['phone']) || ''),
    city: String(g(['city', 'location']) || ''),
    niche: String(g(['niche']) || ''),
    bio: String(g(['bio']) || ''),
    gender: String(g(['gender']) || ''),
    followers: String(g(['followers']) || ''),
    charges: String(g(['charges']) || ''),
    profileComplete: String(g(['profile complete']) || '') === 'Yes',
    picture: String(g(['picture']) || ''),
    created: String(g(['created']) || ''),
    lastLogin: String(g(['last login']) || '')
  };
}

function findUserRow_(email) {
  email = String(email || '').trim().toLowerCase();
  if (!email || email.indexOf('@') < 0) return null;
  var sh = getUsersSheet_();
  var data = sh.getDataRange().getValues();
  if (data.length < 2) return null;
  var headers = data[0];
  var emailCol = headerIndex_(headers, ['email']);
  if (emailCol < 0) emailCol = 0;
  for (var r = 1; r < data.length; r++) {
    if (String(data[r][emailCol] || '').trim().toLowerCase() === email) {
      return { row: r + 1, headers: headers, values: data[r], user: userFromRow_(headers, data[r]) };
    }
  }
  return null;
}


function decodeGoogleCredential_(cred) {
  if (!cred) return null;
  try {
    var parts = String(cred).split('.');
    if (parts.length < 2) return null;
    var b64 = parts[1].replace(/-/g, '+').replace(/_/g, '/');
    while (b64.length % 4) b64 += '=';
    var json = Utilities.newBlob(Utilities.base64Decode(b64)).getDataAsString();
    var payload = JSON.parse(json);
    // basic checks
    if (payload.exp && payload.exp * 1000 < Date.now()) return null;
    if (!payload.email) return null;
    return {
      email: String(payload.email).toLowerCase(),
      name: payload.name || payload.given_name || '',
      picture: payload.picture || '',
      emailVerified: !!payload.email_verified,
      sub: payload.sub || ''
    };
  } catch (e) {
    console.warn('google decode ' + e);
    return null;
  }
}

function authLogin_(body) {
  var g = null;
  if (body.googleCredential) g = decodeGoogleCredential_(body.googleCredential);
  var email = extractEmail_((g && g.email) || body.email || body.googleEmail || '');
  if (!email) throw new Error('email_required');
  email = email.toLowerCase();
  var name = String((g && g.name) || body.name || body.displayName || '').trim();
  var picture = String((g && g.picture) || body.picture || '').trim();
  var now = Utilities.formatDate(new Date(), 'Asia/Kolkata', 'dd/MM/yyyy HH:mm:ss');
  var found = findUserRow_(email);
  if (!found) {
    getUsersSheet_().appendRow([
      email,
      '', // role chosen later
      name,
      '',
      '',
      '',
      '',
      '',
      '',
      '',
      '',
      '',
      '',
      'No',
      picture,
      now,
      now
    ]);
    found = findUserRow_(email);
  } else {
    // update last login + picture/name if empty
    var sh = getUsersSheet_();
    var h = found.headers;
    var li = headerIndex_(h, ['last login']);
    if (li >= 0) sh.getRange(found.row, li + 1).setValue(now);
    var ni = headerIndex_(h, ['display name', 'name']);
    if (ni >= 0 && name && !String(found.values[ni] || '').trim()) sh.getRange(found.row, ni + 1).setValue(name);
    var pi = headerIndex_(h, ['picture']);
    if (pi >= 0 && picture) sh.getRange(found.row, pi + 1).setValue(picture);
    found = findUserRow_(email);
  }
  return { ok: true, user: found.user };
}

function authProfile_(body) {
  var email = extractEmail_(body.email || '');
  if (!email) throw new Error('email_required');
  email = email.toLowerCase();
  var found = findUserRow_(email);
  if (!found) {
    authLogin_(body);
    found = findUserRow_(email);
  }
  var sh = getUsersSheet_();
  var h = found.headers;
  function set(names, val) {
    var i = headerIndex_(h, names);
    if (i >= 0 && val != null) sh.getRange(found.row, i + 1).setValue(val);
  }
  if (body.role) set(['role'], String(body.role).toLowerCase());
  if (body.displayName != null || body.name != null) set(['display name', 'name'], body.displayName || body.name || '');
  if (body.brandName != null) set(['brand name'], body.brandName);
  if (body.website != null) set(['website'], body.website);
  if (body.instagram != null) set(['instagram'], body.instagram);
  if (body.phone != null) set(['phone'], body.phone);
  if (body.city != null) set(['city', 'location'], body.city);
  if (body.niche != null) set(['niche'], body.niche);
  if (body.bio != null) set(['bio'], body.bio);
  if (body.gender != null) set(['gender'], body.gender);
  if (body.followers != null) set(['followers'], body.followers);
  if (body.charges != null) set(['charges'], body.charges);
  if (body.picture != null) set(['picture'], body.picture);

  // Mark complete
  var role = String(body.role || found.user.role || '').toLowerCase();
  var complete = 'No';
  if (role === 'brand') {
    var bn = body.brandName != null ? body.brandName : found.user.brandName;
    var ph = body.phone != null ? body.phone : found.user.phone;
    if (String(bn || '').trim() && String(ph || '').trim()) complete = 'Yes';
  } else if (role === 'creator') {
    var dn = body.displayName || body.name || found.user.displayName;
    var ig = body.instagram != null ? body.instagram : found.user.instagram;
    if (String(dn || '').trim() && String(ig || '').trim()) complete = 'Yes';
  }
  set(['profile complete'], complete);

  // If creator profile complete, also ensure influencer sheet has a row (join-style)
  if (role === 'creator' && complete === 'Yes') {
    try {
      syncCreatorToInfluencerSheet_(email, body, found.user);
    } catch (e) { console.warn('sync creator: ' + e); }
  }

  var updated = findUserRow_(email);
  return { ok: true, user: updated.user };
}

function syncCreatorToInfluencerSheet_(email, body, prev) {
  var sh = getInfluencerSheet_();
  var data = sh.getDataRange().getValues();
  var headers = data.length ? data[0] : [];
  var emailCol = headerIndex_(headers, ['email', 'email address']);
  var foundRow = -1;
  if (emailCol >= 0) {
    for (var r = 1; r < data.length; r++) {
      if (extractEmail_(data[r][emailCol]).toLowerCase() === email.toLowerCase()) {
        foundRow = r + 1;
        break;
      }
    }
  }
  var name = body.displayName || body.name || prev.displayName || '';
  var ig = body.instagram || prev.instagram || '';
  var phone = body.phone || prev.phone || '';
  var city = body.city || prev.city || '';
  var followers = body.followers || prev.followers || '';
  var niche = body.niche || prev.niche || '';
  var gender = body.gender || prev.gender || '';
  var charges = body.charges || prev.charges || '';
  var now = Utilities.formatDate(new Date(), 'Asia/Kolkata', 'dd/MM/yyyy HH:mm:ss');
  if (foundRow < 0) {
    // same shape as join form
    sh.appendRow([
      now, '', name, ig, phone, city, followers, niche, '', '',
      email, gender, charges, '', 'New (login profile)', 'Via Connectly login'
    ]);
  }
}

function shortlistGet_(email) {
  email = String(email || '').trim().toLowerCase();
  if (!email) return [];
  var sh = getShortlistsSheet_();
  var data = sh.getDataRange().getValues();
  if (data.length < 2) return [];
  var out = [];
  for (var r = 1; r < data.length; r++) {
    if (String(data[r][0] || '').trim().toLowerCase() === email) {
      out.push({
        creatorId: String(data[r][1] || ''),
        creatorName: String(data[r][2] || ''),
        creatorIg: String(data[r][3] || ''),
        savedAt: String(data[r][4] || '')
      });
    }
  }
  return out;
}

function shortlistToggle_(body) {
  var email = extractEmail_(body.email || '');
  if (!email) throw new Error('email_required');
  email = email.toLowerCase();
  var creatorId = String(body.creatorId || body.id || '').trim();
  if (!creatorId) throw new Error('creator_id_required');
  var sh = getShortlistsSheet_();
  var data = sh.getDataRange().getValues();
  var removeRows = [];
  for (var r = 1; r < data.length; r++) {
    if (String(data[r][0] || '').trim().toLowerCase() === email && String(data[r][1] || '') === creatorId) {
      removeRows.push(r + 1);
    }
  }
  if (removeRows.length) {
    for (var i = removeRows.length - 1; i >= 0; i--) sh.deleteRow(removeRows[i]);
    return { ok: true, saved: false, items: shortlistGet_(email) };
  }
  var now = Utilities.formatDate(new Date(), 'Asia/Kolkata', 'dd/MM/yyyy HH:mm');
  sh.appendRow([
    email,
    creatorId,
    body.creatorName || '',
    body.creatorIg || body.instagram || '',
    now
  ]);
  return { ok: true, saved: true, items: shortlistGet_(email) };
}

function shortlistSet_(body) {
  var email = extractEmail_(body.email || '');
  if (!email) throw new Error('email_required');
  email = email.toLowerCase();
  var items = body.items || body.ids || [];
  var sh = getShortlistsSheet_();
  var data = sh.getDataRange().getValues();
  // delete existing for email
  for (var r = data.length - 1; r >= 1; r--) {
    if (String(data[r][0] || '').trim().toLowerCase() === email) sh.deleteRow(r + 1);
  }
  var now = Utilities.formatDate(new Date(), 'Asia/Kolkata', 'dd/MM/yyyy HH:mm');
  for (var i = 0; i < items.length; i++) {
    var it = items[i];
    if (typeof it === 'string') {
      sh.appendRow([email, it, '', '', now]);
    } else {
      sh.appendRow([email, it.creatorId || it.id || '', it.creatorName || it.name || '', it.creatorIg || it.instagram || '', now]);
    }
  }
  return { ok: true, items: shortlistGet_(email) };
}


function doGet(e) {
  e = e || {parameter:{}};
  if (rateLimited_('get', MAX_GET_PER_MIN)) return deny_('rate_limited');
  var action = (e.parameter && e.parameter.action) || '';
  if (action === 'campaigns') {
    return ContentService.createTextOutput(cachedCsv_('csv_campaigns_v1', function(){ return publicCsv_(getCampaignSheet_()); }))
      .setMimeType(ContentService.MimeType.CSV);
  }
  if (action === 'deals') {
    var email = (e.parameter && e.parameter.email) || '';
    return jsonOut({ok:true, deals: listDealsByEmail_(email)});
  }
  return ContentService.createTextOutput(cachedCsv_('csv_influencers_v1', function(){ return publicCsv_(getInfluencerSheet_()); }))
    .setMimeType(ContentService.MimeType.CSV);
}
function doPost(e) {
  try {
    var raw = (e && e.postData && e.postData.contents) ? String(e.postData.contents) : '';
    if (raw.length > MAX_POST_BYTES) return deny_('payload_too_large');
    var body = raw ? JSON.parse(raw) : {};
    if (rateLimited_('post', MAX_POST_PER_MIN)) return deny_('rate_limited');
    var action = String(body.action || 'join');
    if (['join','campaign','apply','invite','deal_update','auth_login','auth_profile','shortlist_get','shortlist_set','shortlist_toggle'].indexOf(action) < 0) return deny_('unknown_action');

    if (action === 'campaign') { appendCampaign_(body); return jsonOut({ok:true,message:'Campaign added'}); }
    if (action === 'apply') { appendApplication_(body); notifyBrandNewApplication_(body); return jsonOut({ok:true,message:'Application added'}); }
    if (action === 'invite') {
      var found = lookupCreatorContact_(body);
      if (found.email) body.creatorEmail = found.email;
      if (found.name) body.creatorName = found.name;
      appendInvite_(body);
      var dealId = body.dealId || body.id || makeDealId_();
      body.dealId = dealId;
      body.status = body.status || 'Proposed';
      appendDeal_(body);
      notifyCreatorInvite_(body);
      return jsonOut({ok:true, message:'Invite sent', emailed:!!body.creatorEmail, dealId:dealId});
    }
    if (action === 'deal_update') { updateDeal_(body); return jsonOut({ok:true,message:'Deal updated'}); }

    if (action === 'auth_login') { return jsonOut(authLogin_(body)); }
    if (action === 'auth_profile') { return jsonOut(authProfile_(body)); }
    if (action === 'shortlist_get') {
      return jsonOut({ ok: true, items: shortlistGet_(body.email || '') });
    }
    if (action === 'shortlist_toggle') { return jsonOut(shortlistToggle_(body)); }
    if (action === 'shortlist_set') { return jsonOut(shortlistSet_(body)); }

    // join
    var sheet = getInfluencerSheet_();
    var now = new Date();
    var timestamp = Utilities.formatDate(now, Session.getScriptTimeZone()||'Asia/Kolkata', 'dd/MM/yyyy HH:mm:ss');
    sheet.appendRow([
      timestamp, body.age||'', body.name||'', body.instagram||'', body.phone||'',
      body.city||'', body.followers||'', body.skill||'', body.income||'', body.working||'',
      body.email||'', body.gender||'', body.charges||'', '', 'New (website form)', 'Via Connectly website'
    ]);
    return jsonOut({ok:true, message:'Creator added'});
  } catch(err) {
    return jsonOut({ok:false, error:String(err)});
  }
}
function onApplicationStatusChange(e) {
  try {
    if (!e || !e.range) return;
    var sh = e.range.getSheet();
    if (sh.getName() !== APPLICATIONS_SHEET_NAME) return;
    var col = e.range.getColumn(), row = e.range.getRow();
    if (row < 2) return;
    var headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
    var statusCol = colIndexByHeader_(headers, ['status']);
    if (statusCol < 0 || col !== statusCol) return;
    var newStatus = String(e.value||'').trim().toLowerCase();
    if (newStatus !== 'approved' && newStatus !== 'rejected') return;
    var vals = sh.getRange(row,1,1,headers.length).getValues()[0];
    function cell(names){ var c=colIndexByHeader_(headers,names); return c>0 ? vals[c-1] : ''; }
    var contact = String(cell(['contact'])||'');
    var creator = String(cell(['creator name','creator','name'])||'Creator');
    var campTitle = String(cell(['campaign title'])||'campaign');
    var brandName = String(cell(['brand name'])||'');
    var brandEmail = String(cell(['brand email'])||'');
    var brandPhone = String(cell(['brand phone'])||'');
    var budget = String(cell(['budget'])||'');
    var niche = String(cell(['niche'])||'');
    var platform = String(cell(['platform'])||'');
    var location = String(cell(['location'])||'');
    var deadline = String(cell(['deadline'])||'');
    var deliverables = String(cell(['deliverables'])||'');
    var brief = String(cell(['campaign brief','brief'])||'');
    var campaignId = String(cell(['campaign id'])||'');
    var camp = lookupCampaignById_(campaignId);
    brandName = brandName||camp.brandName||'';
    brandEmail = brandEmail||camp.brandEmail||'';
    brandPhone = brandPhone||camp.brandPhone||'';
    budget = budget||camp.budget||''; niche = niche||camp.niche||'';
    platform = platform||camp.platform||''; location = location||camp.location||'';
    deadline = deadline||camp.deadline||''; deliverables = deliverables||camp.deliverables||'';
    brief = brief||camp.campaignBrief||camp.brief||'';
    campTitle = (campTitle && campTitle!=='campaign') ? campTitle : (camp.campaignTitle||campTitle);
    var to = extractEmail_(contact);
    if (!to && contact.indexOf('@')>=0) to = contact;
    var subject = newStatus==='approved'
      ? 'Good news — your application was selected'
      : 'Update on your Connectly application';
    var campRows = kvRow_('Title',campTitle)+kvRow_('Brand',brandName)+kvRow_('Budget',budget)+kvRow_('Niche',niche)+kvRow_('Platform',platform)+kvRow_('Location',location)+kvRow_('Deadline',deadline)+kvRow_('Deliverables',deliverables)+kvRow_('Brief',brief);
    var bodyTxt = campaignDetailsBlock_({campaignTitle:campTitle,brandName:brandName,budget:budget,niche:niche,platform:platform,location:location,deadline:deadline,deliverables:deliverables,campaignBrief:brief});
    var html;
    if (newStatus==='approved') {
      bodyTxt += '\n'+brandContactBlock_({brandName:brandName,brandEmail:brandEmail,brandPhone:brandPhone});
      html = emailShell_('Your application was selected',
        '<p style="color:#cfc8bf;font-size:15px;line-height:1.55;">Hi '+htmlEsc_(creator)+', the brand selected you.</p>'+
        cardBlock_('Campaign',campRows)+
        cardBlock_('Brand contact', kvRow_('Brand',brandName)+kvRow_('Email',brandEmail)+kvRow_('Phone',brandPhone)+kvLink_('WhatsApp',waLink_(brandPhone,'Hi, my application was selected on Connectly'),'Message on WhatsApp'))
      );
    } else {
      html = emailShell_('Application update',
        '<p style="color:#cfc8bf;font-size:15px;line-height:1.55;">Hi '+htmlEsc_(creator)+', your application for '+htmlEsc_(campTitle)+' was not selected this time.</p>'+
        cardBlock_('Campaign',campRows)+
        '<p style="color:#9a9388;font-size:13px;">You can apply to other open campaigns on Connectly.</p>'
      );
    }
    safeSendEmail_(to, subject, bodyTxt, html, extractEmail_(brandEmail||''));
  } catch(err){ console.warn(err); }
}
function authorizeAll() {
  getCampaignSheet_(); getApplicationsSheet_(); ensureCampaignPhoneHeader_();
  getUsersSheet_(); getShortlistsSheet_();
  SpreadsheetApp.openById(SHEET_ID).getSheets()[0].getName();
  try { MailApp.getRemainingDailyQuota(); } catch(e){}
  try { GmailApp.getInboxUnreadCount(); } catch(e){}
  Logger.log('Authorization complete — also grant Gmail if prompted (better inbox delivery)');
}
function installStatusTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i=0; i<triggers.length; i++) {
    var h = triggers[i].getHandlerFunction();
    if (h==='onEdit' || h==='onApplicationStatusChange') ScriptApp.deleteTrigger(triggers[i]);
  }
  ScriptApp.newTrigger('onApplicationStatusChange').forSpreadsheet(CAMPAIGN_SHEET_ID).onEdit().create();
  Logger.log('Trigger installed');
}
/**
 * Run once: sends a test mail to yourself so Gmail learns the sender.
 * Check inbox + spam, then mark "Not spam" if needed.
 */
function warmUpSender() {
  var me = Session.getActiveUser().getEmail();
  safeSendEmail_(me,
    'Connectly is connected',
    'Hi,\n\nThis is a one-time test from your Connectly Apps Script.\nIf this landed in spam, open it and click "Not spam".\n\n— Connectly',
    emailShell_('Connectly is connected',
      '<p style="color:#cfc8bf;font-size:15px;">This is a one-time test from your Connectly script.</p>'+
      '<p style="color:#9a9388;font-size:13px;">If it landed in spam, open it and click <strong>Not spam</strong>. That helps future mails reach the inbox.</p>'
    )
  );
  Logger.log('Warm-up mail sent to ' + me);
}
