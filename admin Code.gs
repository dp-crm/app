const SHEET_ID = '1fPv8EJJTAkjl82eMcnyV34aLQJZ_BzAiXMRYL9ydwTc'; // Deep Pocket CRM — Admin
const ENQUIRIES_TAB = 'ENQUIRIES', FEEDBACK_TAB = 'FEEDBACK', ACTIVITY_TAB = 'ACTIVITY_LOG';
const VERSION = '2.0.0';
const STAGE_ORDER = ['Enquiry','Mail Sent','Prompt Generated','File Generated','Deployment Done'];

// ═══════════════════════════════════════════════
// ONE-TIME SETUP — run this once from the Apps Script editor (select
// setup from the function dropdown, click Run). It creates the 5-minute
// inbox-check trigger. Re-running it is safe; it won't create duplicates.
// Before running, set your Script Properties (Project Settings > Script
// Properties): GEMINI_API_KEY, CLIENT_FORM_URL, ADMIN_EMAIL.
// ═══════════════════════════════════════════════
function setup(){
  ScriptApp.getProjectTriggers().forEach(function(t){
    if(t.getHandlerFunction()==='checkInbox') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('checkInbox').timeBased().everyMinutes(5).create();
  Logger.log('Trigger installed: checkInbox will run every 5 minutes.');
}

function doGet(e){
  try{
    var p=e.parameter||{};
    if(p.ping||p.action==='version') return out({version:VERSION, ok:true});
    if(p.action==='syncNow') return out(syncMailNow());
    if(p.action==='syncStatus') return out(getSyncStatus_());
    if(p.payload) return writeData(JSON.parse(decodeURIComponent(p.payload)));
    if(p.action==='readTab') return readTab(p.tab||ENQUIRIES_TAB);
    return readTab(ENQUIRIES_TAB);
  }catch(err){ return out({error:err.message}); }
}
function doPost(e){
  try{
    var d;
    if(e.postData&&e.postData.contents) d=JSON.parse(e.postData.contents);
    else if(e.parameter&&e.parameter.payload) d=JSON.parse(decodeURIComponent(e.parameter.payload));
    if(d && d.action==='syncNow') return out(syncMailNow());
    return writeData(d);
  }catch(err){ return out({error:err.message}); }
}

// ═══════════════════════════════════════════════
// EMAIL — the single source of truth for prospect identity. Every match
// against ENQUIRIES goes through this, never firm name, unless email is
// genuinely unavailable (only the AI-detected email-body prompt-marker
// path, as a last resort).
// ═══════════════════════════════════════════════
function normalizeEmail(email){
  return String(email||'').trim().toLowerCase();
}

function readTab(t){
  var ss=SpreadsheetApp.openById(SHEET_ID);
  var sh=ss.getSheetByName(t);
  if(!sh) return out({rows:[]});
  var lr=sh.getLastRow(), lc=sh.getLastColumn();
  if(lr===0||lc===0) return out({rows:[]});
  var raw=sh.getRange(1,1,lr,lc).getValues();
  var cleaned=raw.map(function(row){
    return row.map(function(c){
      if(c===null||c===undefined||c==='') return '';
      if(c instanceof Date){ var y=c.getFullYear(),m=String(c.getMonth()+1).padStart(2,'0'),d=String(c.getDate()).padStart(2,'0'); return y+'-'+m+'-'+d; }
      return String(c).trim();
    });
  });
  return out({rows:cleaned});
}

function nextId(sheet,colIndex,prefix){
  var lr=sheet.getLastRow(); if(lr<2) return prefix+'1';
  var vals=sheet.getRange(2,colIndex,lr-1,1).getValues();
  var max=0; var re=new RegExp('^'+prefix+'(\\d+)$','i');
  vals.forEach(function(v){ var mm=String(v[0]||'').trim().match(re); if(mm){ var n=parseInt(mm[1],10); if(n>max)max=n; } });
  return prefix+(max+1);
}

// 15 columns. Column H stays intentionally blank (legacy — Enquiry Created
// Date in col E already covers it). Column O (Normalized Email) is new —
// added at the END so existing rows/columns are never disturbed.
function ensureEnquiriesHeader_(sh){
  if(sh.getLastRow()===0){
    var hdr=[['Record ID','Firm Name','Contact No','Email','Enquiry Created Date','Status','Notes','','Modified Date','Intake JSON','Thank You Mail Sent Date','Prompt Generated Date','File Generated Date','Deployment Done Date','Normalized Email']];
    var hr=sh.getRange(1,1,1,15); hr.setValues(hdr); hr.setFontWeight('bold').setBackground('#0A0E14').setFontColor('#F4B942');
    sh.setFrozenRows(1);
  } else if(sh.getLastColumn()<15){
    // Upgrading an existing sheet that predates the Normalized Email column.
    sh.getRange(1,15).setValue('Normalized Email').setFontWeight('bold').setBackground('#0A0E14').setFontColor('#F4B942');
    backfillNormalizedEmails_(sh);
  }
}

// One-time backfill for rows that existed before column O was added, so
// findProspectByEmail_ works correctly on old data without you having to
// touch anything by hand.
function backfillNormalizedEmails_(sh){
  var n=sh.getLastRow();
  if(n<2) return;
  var emails=sh.getRange(2,4,n-1,1).getValues();
  var norm=emails.map(function(r){ return [normalizeEmail(r[0])]; });
  sh.getRange(2,15,n-1,1).setValues(norm);
}

function ensureActivityHeader_(sh){
  if(sh.getLastRow()===0){
    var hdr=[['Timestamp','Record ID','Firm Name','Event','Source','Gmail Ref','Notes']];
    var hr=sh.getRange(1,1,1,7); hr.setValues(hdr); hr.setFontWeight('bold').setBackground('#0A0E14').setFontColor('#F4B942');
    sh.setFrozenRows(1);
    try{ sh.setColumnWidth(1,150); sh.setColumnWidth(3,180); sh.setColumnWidth(4,160); sh.setColumnWidth(7,400); }catch(e){}
  }
}

// Appends one row to ACTIVITY_LOG — every automated or manual workflow
// event goes through here, so Section 15's timeline is always complete
// rather than reconstructed after the fact from scattered timestamps.
function logAutomationEvent(ss, recordId, firmName, eventType, source, gmailRef, notes){
  try{
    var sh = ss.getSheetByName(ACTIVITY_TAB) || ss.insertSheet(ACTIVITY_TAB);
    ensureActivityHeader_(sh);
    var ts = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
    sh.appendRow([ts, recordId||'', firmName||'', eventType||'', source||'', gmailRef||'', notes||'']);
  }catch(e){ Logger.log('logAutomationEvent failed: '+e.message); }
}

// Stamps a stage-reached date into the given row, but only if that column
// is still empty — first-time-reached semantics, so re-syncing or moving a
// status back and forth doesn't overwrite genuine history.
function stampStageDateIfEmpty_(sh, rowNum, col, dateStr){
  var cur = sh.getRange(rowNum, col).getValue();
  if(!cur){ sh.getRange(rowNum, col).setValue(dateStr).setNumberFormat('@'); }
}

// The one place status transitions happen. Forward-only, always — an
// automated OR manual call to move a record to an earlier stage than it's
// already reached is silently ignored (the date-stamp columns are
// first-time-only anyway, so this just protects the Status cell itself).
// Returns true if the status actually changed.
function advanceStatusSafely(sh, rowNum, newStatus, dateCol, currentStatus){
  var idx = STAGE_ORDER.indexOf(newStatus);
  var curIdx = STAGE_ORDER.indexOf(currentStatus||'Enquiry');
  var today = todayStr_();
  if(dateCol) stampStageDateIfEmpty_(sh, rowNum, dateCol, today);
  sh.getRange(rowNum,9).setValue(today); // Modified Date, col I
  if(idx > curIdx){
    sh.getRange(rowNum,6).setValue(newStatus);
    return true;
  }
  return false;
}

function todayStr_(){ return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'); }

// ═══════════════════════════════════════════════
// PROSPECT LOOKUP — email-first, always. This is the one function every
// automated and manual matching path should call; nothing else should
// loop over the sheet independently searching by email or firm name.
// ═══════════════════════════════════════════════
function findProspectByEmail_(sh, email){
  var norm = normalizeEmail(email);
  if(!norm) return -1;
  var n=sh.getLastRow();
  if(n<2) return -1;
  var emails=sh.getRange(2,15,n-1,1).getValues(); // col O, Normalized Email
  for(var i=0;i<emails.length;i++){
    if(String(emails[i][0]||'').trim()===norm) return i+2; // 1-based row number
  }
  return -1;
}

// Fallback only — used solely when a genuinely unmatched-by-email prompt
// email carries no reliable email at all (shouldn't normally happen, since
// senderEmail is always available for inbound mail). Never used for the
// client-form submission path, which always has an email now.
function findProspectByFirmNameFallback_(sh, firmName){
  if(!firmName) return -1;
  var n=sh.getLastRow();
  if(n<2) return -1;
  var names=sh.getRange(2,2,n-1,1).getValues();
  var target=firmName.trim().toLowerCase();
  for(var i=0;i<names.length;i++){
    if(String(names[i][0]||'').trim().toLowerCase()===target) return i+2;
  }
  return -1;
}

function writeData(data){
  var _lock=LockService.getScriptLock();
  try{ _lock.waitLock(20000); }catch(e){ return out({error:'Server busy, retry'}); }
  try{ return _writeInner(data); } finally { try{SpreadsheetApp.flush();}catch(e){} try{_lock.releaseLock();}catch(e){} }
}

function _writeInner(data){
  var ss=SpreadsheetApp.openById(SHEET_ID);

  if(data.action==='appendEnquiry'){
    var sh=ss.getSheetByName(ENQUIRIES_TAB)||ss.insertSheet(ENQUIRIES_TAB);
    ensureEnquiriesHeader_(sh);
    var row=data.row||[];
    // Server-side duplicate guard, in addition to whatever check the admin
    // console already did client-side — the same prospect must not appear
    // twice just because two people (or two tabs) tried to add them at once.
    var email = row[3]||'';
    var existingRow = findProspectByEmail_(sh, email);
    if(existingRow>0){
      return out({success:false, duplicate:true, existingRecordId: String(sh.getRange(existingRow,1).getValue())});
    }
    if(!String(row[0]||'').trim()) row[0]=nextId(sh,1,'E');
    row[14] = normalizeEmail(email); // col O
    sh.appendRow(row);
    var lr=sh.getLastRow(); sh.getRange(lr,1,1,row.length).setNumberFormat('@');
    logAutomationEvent(ss, row[0], row[1], 'Enquiry created', 'Manual', '', '');
    return out({success:true, recordId:row[0]});
  }

  if(data.action==='updateEnquiry'){
    var sh=ss.getSheetByName(ENQUIRIES_TAB);
    if(sh && sh.getLastRow()>1){
      var id=String(data.recordId||'').trim();
      var n=sh.getLastRow();
      var ids=sh.getRange(2,1,n-1,1).getValues();
      for(var i=0;i<ids.length;i++){
        if(String(ids[i][0]||'').trim()===id){
          var row=data.row||[];
          row[14] = normalizeEmail(row[3]||'');
          sh.getRange(i+2,1,1,row.length).setValues([row]);
          sh.getRange(i+2,1,1,row.length).setNumberFormat('@');
          break;
        }
      }
    }
    return out({success:true});
  }

  if(data.action==='deleteEnquiry'){
    var sh=ss.getSheetByName(ENQUIRIES_TAB);
    if(sh && sh.getLastRow()>1){
      var id=String(data.recordId||'').trim();
      var n=sh.getLastRow();
      var ids=sh.getRange(2,1,n-1,1).getValues();
      for(var i=ids.length-1;i>=0;i--){
        if(String(ids[i][0]||'').trim()===id){ sh.deleteRow(i+2); break; }
      }
    }
    return out({success:true});
  }

  if(data.action==='appendFeedback'){
    var sh=ss.getSheetByName(FEEDBACK_TAB)||ss.insertSheet(FEEDBACK_TAB);
    if(sh.getLastRow()===0){
      var hdr=[['Firm Name','Category','Feedback']];
      var hr=sh.getRange(1,1,1,3); hr.setValues(hdr); hr.setFontWeight('bold').setBackground('#0A0E14').setFontColor('#F4B942');
      sh.setFrozenRows(1);
      try{ sh.setColumnWidth(1,180); sh.setColumnWidth(2,160); sh.setColumnWidth(3,600); }catch(e){}
    }
    var category = data.category || data.feedbackCategory || data.type || data.feedbackType || '';
    var msg = data.senderName ? '['+data.senderName+'] '+(data.message||'') : (data.message||'');
    sh.appendRow([data.appName||'', category, msg]);
    var lr=sh.getLastRow(); sh.getRange(lr,1,1,3).setNumberFormat('@').setWrap(true);
    return out({success:true});
  }

  // Older email-based intake path (kept working, superseded in practice by
  // notifySubmission below, which the current client form actually uses).
  if(data.action==='clientIntakeSubmitted'){
    var adminEmail = PropertiesService.getScriptProperties().getProperty('ADMIN_EMAIL') || 'deeppocket1610@gmail.com';
    var subject = 'Deep Pocket Customization Details - ' + (data.firmName || 'New Firm');
    var summary = buildIntakeSummaryText_(data);
    var body = summary
      + '\n\n---\nDo not remove the line below, it lets our system link this to your enquiry automatically---\n'
      + 'DEEPPOCKET_INTAKE_JSON: ' + JSON.stringify(data);
    var opts = {};
    if(data.email) opts.replyTo = data.email;
    GmailApp.sendEmail(adminEmail, subject, body, opts);
    return out({success:true});
  }

  // Called when Generate is pressed in the admin app's Intake tab. Saves
  // the 5 generated setup prompts into their own sheet tab, named after the
  // firm, for future reference. Re-generating for the same firm overwrites
  // that firm's row in place rather than creating duplicate tabs.
  if(data.action==='savePrompts'){
    var firmName = String(data.firmName||'').trim() || 'Untitled Firm';
    var tabName = saveFirmPromptTexts_(ss, firmName, data.prompts||[]);
    var driveSelf = {};
    try{ driveSelf = saveClientDriveDeliverables_(firmName, data.prompts||[], data.iconBase64||''); }catch(e){ driveSelf = {}; }
    return out({success:true, tab: tabName, folderUrl: driveSelf.folderUrl||'', docUrl: driveSelf.docUrl||'', iconUrl: driveSelf.iconUrl||''});
  }

  // Sent by the public client-facing form on submit, with the 5 finished
  // prompts already generated in the client's browser AND (as of this
  // version) their email address — the primary key for matching this to
  // the right enquiry. Firm name is used only as a last-resort fallback if
  // email is somehow blank.
  if(data.action==='notifySubmission'){
    var firmName2 = String(data.firmName||'').trim() || 'Untitled Firm';
    var submitEmail = String(data.email||'').trim();
    var promptTitles = ['Step 1 — Identity Replacement','Step 2 — Visual Identity','Step 3 — Feature Scope','Step 4 — Communications & Compliance','Step 5 — Technical & Handover'];
    var promptTexts = (data.prompts||[]).map(function(p){ return (p && typeof p==='object') ? (p.text||'') : String(p||''); });
    saveFirmPromptTexts_(ss, firmName2, promptTexts);
    matchProspectAndAdvance_(ss, submitEmail, firmName2, data.phone||'', 'Prompt Generated', 12, 'Client form submitted', 'Automated', '');

    try{
      var adminEmail2 = PropertiesService.getScriptProperties().getProperty('ADMIN_EMAIL') || 'deeppocket1610@gmail.com';
      var wordBlob = buildPromptsWordBlob_(firmName2, promptTexts);
      var attachments = [wordBlob];
      var hasIcon = (data.iconMode==='upload' && data.iconBase64);
      if(hasIcon){
        try{
          attachments.push(Utilities.newBlob(Utilities.base64Decode(data.iconBase64), 'image/png', firmName2+' icon-192.png'));
        }catch(e){ /* icon attach is best-effort — the prompts still go out either way */ }
      }
      var subject = firmName2 + ' CRM application prompt';

      var nBody = 'A new CRM customization submission has come in from '+firmName2+'.\n\n';
      for(var pi=0; pi<5; pi++){
        nBody += '───── '+promptTitles[pi]+' ─────\n\n'+(promptTexts[pi]||'(not generated)')+'\n\n';
      }
      nBody += (hasIcon ? 'The firm\'s uploaded logo is attached, along with a Word copy of these prompts.' : 'This firm asked for an auto-generated initials icon rather than uploading a logo. A Word copy of these prompts is attached.')
        + '\n\nAlso logged in the "'+sanitizeSheetName_(firmName2)+'" tab in your sheet for reference:\n'+ss.getUrl();

      var hBody = '<div style="font-family:Arial,Helvetica,sans-serif;font-size:13.5px;line-height:1.6;color:#222;max-width:600px">'
        + '<p>A new CRM customization submission has come in from <b>'+firmName2+'</b>.</p>';
      for(var hi=0; hi<5; hi++){
        hBody += '<div style="margin:18px 0;padding:14px 16px;background:#f4f4f4;border-left:3px solid #F4B942;border-radius:4px">'
          + '<div style="font-weight:bold;margin-bottom:8px">'+promptTitles[hi]+'</div>'
          + '<div style="white-space:pre-wrap;font-family:monospace;font-size:12px;color:#333">'+(promptTexts[hi]||'(not generated)').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;')+'</div>'
          + '</div>';
      }
      hBody += '<p>'+(hasIcon ? 'The firm\'s uploaded logo is attached, along with a Word copy of these prompts.' : 'This firm asked for an auto-generated initials icon rather than uploading a logo. A Word copy of these prompts is attached.')+'</p>'
        + '<p>Also logged in the "'+sanitizeSheetName_(firmName2)+'" tab in your sheet: <a href="'+ss.getUrl()+'">open sheet</a></p>'
        + '</div>';

      GmailApp.sendEmail(adminEmail2, subject, nBody, { attachments: attachments, htmlBody: hBody });
    }catch(e){ /* non-fatal — the sheet record above already succeeded either way */ }

    return out({success:true});
  }

  // Sent when a new enquiry is saved manually from the Dashboard/Enquiries
  // (the "+ New Enquiry" modal) — same welcome template as the automatic
  // inbox-detected path, sent as a fresh email since there's no thread to
  // reply to. Also stamps that row's Mail Sent Date and Status.
  if(data.action==='sendEnquiryWelcome'){
    var email = String(data.email||'').trim();
    if(email){
      GmailApp.sendEmail(email, 'Thank you for your enquiry', autoReplyBody_(data.firmName), {htmlBody: autoReplyBodyHtml_(data.firmName)});
    }
    if(data.recordId){
      var sh=ss.getSheetByName(ENQUIRIES_TAB);
      if(sh && sh.getLastRow()>1){
        var id=String(data.recordId).trim();
        var n=sh.getLastRow();
        var ids=sh.getRange(2,1,n-1,1).getValues();
        for(var i=0;i<ids.length;i++){
          if(String(ids[i][0]||'').trim()===id){
            var r=i+2;
            var curStatus = sh.getRange(r,6).getValue()||'Enquiry';
            advanceStatusSafely(sh, r, 'Mail Sent', 11, curStatus);
            logAutomationEvent(ss, id, data.firmName||'', 'Thank-you mail sent', 'Manual', '', '');
            break;
          }
        }
      }
    }
    return out({success:true});
  }

  if(data.action==='syncNow'){
    return out(syncMailNow());
  }

  return out({error:'Unknown action: '+data.action});
}

// Shared helper: matches a prospect by email (primary) with an optional
// firm-name fallback, advances their status forward-only, logs the event,
// and creates a fresh record only if truly nothing matches — used by both
// the direct form-submission path and the email-detected prompt path, so
// there's exactly one matching/advancement rule rather than two drifting
// implementations.
function matchProspectAndAdvance_(ss, email, firmName, phone, newStatus, dateCol, eventLabel, source, gmailRef){
  var sh=ss.getSheetByName(ENQUIRIES_TAB);
  if(!sh){ sh=ss.insertSheet(ENQUIRIES_TAB); }
  ensureEnquiriesHeader_(sh);

  var rowNum = findProspectByEmail_(sh, email);
  if(rowNum<0 && !email && firmName){
    // Only reached when there's genuinely no email to match on at all.
    rowNum = findProspectByFirmNameFallback_(sh, firmName);
  }

  if(rowNum>0){
    var curStatus = sh.getRange(rowNum,6).getValue()||'Enquiry';
    var changed = advanceStatusSafely(sh, rowNum, newStatus, dateCol, curStatus);
    var recId = String(sh.getRange(rowNum,1).getValue());
    logAutomationEvent(ss, recId, firmName||sh.getRange(rowNum,2).getValue(), eventLabel + (changed?'':' (status already at or beyond this stage)'), source, gmailRef, '');
    return { matched:true, recordId: recId, rowNum: rowNum };
  } else {
    var id=nextId(sh,1,'E');
    var today=todayStr_();
    var row=[id, firmName||'(unmatched submission)', phone||'', email||'', today, newStatus, 'Auto-created: could not match to an existing enquiry - check manually', '', today, '', '', '', '', '', normalizeEmail(email)];
    sh.appendRow(row);
    var newRowNum = sh.getLastRow();
    sh.getRange(newRowNum,1,1,row.length).setNumberFormat('@');
    stampStageDateIfEmpty_(sh, newRowNum, dateCol, today);
    logAutomationEvent(ss, id, firmName||'', eventLabel + ' (new record — no existing match found)', source, gmailRef, 'Check manually: auto-created without a matching prior enquiry');
    return { matched:false, recordId: id, rowNum: newRowNum };
  }
}

function sanitizeSheetName_(name){
  var clean = String(name).replace(/[\[\]\*\/\\\?:]/g,'').trim();
  if(!clean) clean = 'Untitled Firm';
  if(clean.length>90) clean = clean.substring(0,90);
  return clean;
}

// Shared by both 'savePrompts' (you press Generate yourself) and
// 'notifySubmission' (the client submits the public form) — one sheet tab
// per firm either way, so it doesn't matter which path produced the prompts.
function saveFirmPromptTexts_(ss, firmName, promptTexts){
  var tabName = sanitizeSheetName_(firmName);
  var sh = ss.getSheetByName(tabName);
  if(!sh){ sh = ss.insertSheet(tabName); }
  if(sh.getLastRow()===0){
    var hdr=[['Firm Name','Prompt 1','Prompt 2','Prompt 3','Prompt 4','Prompt 5']];
    var hr=sh.getRange(1,1,1,6); hr.setValues(hdr); hr.setFontWeight('bold').setBackground('#0A0E14').setFontColor('#F4B942');
    sh.setFrozenRows(1);
    try{ sh.setColumnWidth(1,180); for(var ci=2;ci<=6;ci++) sh.setColumnWidth(ci,420); }catch(e){}
  }
  var row = [firmName, promptTexts[0]||'', promptTexts[1]||'', promptTexts[2]||'', promptTexts[3]||'', promptTexts[4]||''];
  sh.getRange(2,1,1,6).setValues([row]);
  try{ sh.getRange(2,1,1,6).setWrap(true).setVerticalAlignment('top'); }catch(e){}
  return tabName;
}

function getOrCreateFolder_(name){
  var it = DriveApp.getFoldersByName(name);
  if(it.hasNext()) return it.next();
  return DriveApp.createFolder(name);
}

function getOrCreateNestedFolder_(parentName, childName){
  var parent = getOrCreateFolder_(parentName);
  var it = parent.getFoldersByName(childName);
  if(it.hasNext()) return it.next();
  return parent.createFolder(childName);
}

// Builds the 5 prompts into an actual Word file and returns it as a Blob —
// via a throwaway Google Doc, since Apps Script has no native .docx writer.
function buildPromptsWordBlob_(firmName, promptTexts){
  var docName = firmName+' - Setup Prompts.docx';
  var stepTitles = ['Step 1 — Identity Replacement','Step 2 — Visual Identity','Step 3 — Feature Scope','Step 4 — Communications & Compliance','Step 5 — Technical & Handover'];
  var tempDoc = DocumentApp.create('dp-temp-'+Utilities.getUuid());
  var body = tempDoc.getBody();
  body.clear();
  body.appendParagraph(firmName).setHeading(DocumentApp.ParagraphHeading.TITLE);
  body.appendParagraph('Deep Pocket CRM — Setup Prompts').setHeading(DocumentApp.ParagraphHeading.SUBTITLE);
  for(var i=0;i<5;i++){
    body.appendParagraph(stepTitles[i]).setHeading(DocumentApp.ParagraphHeading.HEADING2);
    body.appendParagraph(promptTexts[i]||'(not generated)');
    if(i<4) body.appendPageBreak();
  }
  tempDoc.saveAndClose();
  var tempFile = DriveApp.getFileById(tempDoc.getId());
  var wordBlob = tempFile.getAs(MimeType.MICROSOFT_WORD).setName(docName);
  tempFile.setTrashed(true);
  return wordBlob;
}

// One Drive folder per firm — the prompts as an actual Word file, and the
// firm's uploaded logo if given. Re-running replaces both files in place.
function saveClientDriveDeliverables_(firmName, promptTexts, iconBase64){
  var folder = getOrCreateNestedFolder_('Deep Pocket CRM Clients', sanitizeSheetName_(firmName));

  var docName = firmName+' - Setup Prompts.docx';
  var existingDocs = folder.getFilesByName(docName);
  while(existingDocs.hasNext()){ existingDocs.next().setTrashed(true); }

  var wordBlob = buildPromptsWordBlob_(firmName, promptTexts);
  var wordFile = folder.createFile(wordBlob);

  var iconUrl = '';
  if(iconBase64){
    var iconName = 'icon-192.png';
    var existingIcons = folder.getFilesByName(iconName);
    while(existingIcons.hasNext()){ existingIcons.next().setTrashed(true); }
    var iconBlob = Utilities.newBlob(Utilities.base64Decode(iconBase64), 'image/png', iconName);
    var iconFile = folder.createFile(iconBlob);
    try{ iconFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); }catch(e){}
    iconUrl = iconFile.getUrl();
  }

  return { folderUrl: folder.getUrl(), docUrl: wordFile.getUrl(), iconUrl: iconUrl };
}

function buildIntakeSummaryText_(s){
  var p=s.products||{};
  var prods=[];
  if(p.mf) prods.push('Mutual Funds'); if(p.life) prods.push('Term Life Insurance'); if(p.health) prods.push('Health Insurance');
  if(p.will) prods.push('WILL'); if(p.aif) prods.push('AIF'); if(p.pms) prods.push('PMS'); if(p.nps) prods.push('NPS');
  if(p.other) prods.push(p.other);
  return 'New CRM customization request\n\n'
    + 'Firm: ' + (s.firmName||'') + ' (' + (s.entityType||'') + ')\n'
    + 'Owner: ' + (s.ownerName||'') + ', ' + (s.ownerTitle||'') + '\n'
    + 'Phone: ' + (s.phone||'') + '\n'
    + 'Email: ' + (s.email||'') + '\n'
    + 'Business model: ' + (s.modelType||'') + '\n'
    + 'Products requested: ' + (prods.join(', ')||'-') + '\n';
}

function out(o){ return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function testConnection(){ Logger.log('OK: '+SpreadsheetApp.openById(SHEET_ID).getName()); }

// ═══════════════════════════════════════════════
// SYNC STATUS — backs the Settings tab's "Last Sync" display and the
// "Run Sync Now" button's result.
// ═══════════════════════════════════════════════
function getSyncStatus_(){
  var props = PropertiesService.getScriptProperties();
  return {
    lastSyncTime: props.getProperty('DP_LAST_SYNC_TIME') || '',
    lastSyncResult: props.getProperty('DP_LAST_SYNC_RESULT') || '',
    triggerInstalled: ScriptApp.getProjectTriggers().some(function(t){ return t.getHandlerFunction()==='checkInbox'; })
  };
}
function setSyncStatus_(resultText){
  var props = PropertiesService.getScriptProperties();
  props.setProperty('DP_LAST_SYNC_TIME', Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss'));
  props.setProperty('DP_LAST_SYNC_RESULT', resultText);
}

// Callable on demand from the admin console's "Run Sync Now" button, and
// also what the 5-minute trigger calls internally — one code path either way.
function syncMailNow(){
  var lock = LockService.getScriptLock();
  var gotLock = false;
  try{ gotLock = lock.tryLock(10000); }catch(e){ gotLock = false; }
  if(!gotLock){
    return { success:false, message:'Another sync is already in progress — try again shortly.' };
  }
  try{
    var result = processInbox_();
    processSentMail_();
    var summary = 'Processed '+result.processed+' message(s): '+result.enquiries+' enquiry, '+result.prompts+' prompt, '+result.ignored+' ignored, '+result.duplicates+' duplicate.';
    setSyncStatus_(summary);
    return { success:true, message: summary };
  }catch(e){
    setSyncStatus_('Error: '+e.message);
    return { success:false, message: 'Sync error: '+e.message };
  } finally {
    try{ lock.releaseLock(); }catch(e){}
  }
}

// checkInbox is what the time-driven trigger calls — thin wrapper so the
// trigger's handler function name stays stable even though the real work
// now lives in syncMailNow/processInbox_/processSentMail_.
function checkInbox(){
  syncMailNow();
}

// ═══════════════════════════════════════════════
// INBOX PROCESSING
// ═══════════════════════════════════════════════
function processInbox_(){
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName(ENQUIRIES_TAB) || ss.insertSheet(ENQUIRIES_TAB);
  ensureEnquiriesHeader_(sh);

  var counts = { processed:0, enquiries:0, prompts:0, ignored:0, duplicates:0 };

  // "Already processed" is tracked by message ID, completely independent of
  // Gmail's read/unread flag or thread labels — both of which you can
  // change yourself just by using your own inbox normally. Search is
  // time-bounded (not is:unread) so an email you've already read but not
  // yet processed is still found.
  var threads = GmailApp.search('in:inbox newer_than:3d', 0, 30);
  var lblEnquiry = getLabel_('DP-Enquiry-Processed');
  var lblPrompt = getLabel_('DP-Prompt-Processed');
  var lblIgnored = getLabel_('DP-Ignored');
  var lblDuplicate = getLabel_('DP-DuplicateEnquiry');
  var processedIds = getProcessedIds_();

  threads.forEach(function(thread){
    var msgs = thread.getMessages();
    msgs.forEach(function(msg){
      var msgId = msg.getId();
      if(processedIds.indexOf(msgId)>=0) return;

      try{
        var subject = msg.getSubject()||'';
        var body = msg.getPlainBody()||'';
        var senderEmail = extractEmail_(msg.getFrom());
        var gmailRef = thread.getId()+'/'+msgId;

        var alreadyKnown = findProspectByEmail_(sh, senderEmail) > 0;
        var cls = isRelevantWorkflowMessage_(subject, body, senderEmail, alreadyKnown);

        if(cls.type==='enquiry'){
          if(alreadyKnown){
            thread.addLabel(lblDuplicate);
            counts.duplicates++;
          } else {
            var id=nextId(sh,1,'E');
            var today=todayStr_();
            // Row is created at "Enquiry" first — even if the mail send
            // below fails, the prospect still exists and isn't lost.
            // Status only advances to "Mail Sent" once the send actually
            // succeeds (Section 8's explicit requirement), never before.
            var row=[id, cls.firmName||'', cls.contactNo||'', senderEmail, today, 'Enquiry', subject, '', today, '', '', '', '', '', normalizeEmail(senderEmail)];
            sh.appendRow(row);
            var newRowNum = sh.getLastRow();
            sh.getRange(newRowNum,1,1,row.length).setNumberFormat('@');
            thread.addLabel(lblEnquiry);
            try{
              sendAutoReply_(msg, senderEmail, cls.firmName);
              advanceStatusSafely(sh, newRowNum, 'Mail Sent', 11, 'Enquiry');
              logAutomationEvent(ss, id, cls.firmName||'', 'Enquiry detected + thank-you mail sent', 'Automated', gmailRef, subject);
            }catch(sendErr){
              logAutomationEvent(ss, id, cls.firmName||'', 'Enquiry detected but thank-you mail FAILED: '+sendErr.message, 'Automated', gmailRef, subject);
            }
            counts.enquiries++;
          }

        } else if(cls.type==='prompt'){
          var res = matchProspectAndAdvance_(ss, senderEmail, cls.firmName, '', 'Prompt Generated', 12, 'Prompt-related email detected', 'Automated', gmailRef);
          try{
            var note = '[Received via inbox email — not the 5 structured prompts, since those only exist once generated through the form. Raw email content below for reference.]\n\n' + body.slice(0,1800);
            saveFirmPromptTexts_(ss, cls.firmName || 'Unknown Firm', [note,'','','','']);
          }catch(e){}
          thread.addLabel(lblPrompt);
          counts.prompts++;

        } else {
          thread.addLabel(lblIgnored);
          counts.ignored++;
        }

        msg.markRead();
        markProcessed_(msgId);
        counts.processed++;
      }catch(err){
        Logger.log('processInbox_: error on message: '+err.message);
        // left unmarked — retried on the next pass
      }
    });
  });

  return counts;
}

// Classification — deterministic fast paths first (work even without
// GEMINI_API_KEY or if the AI call ever fails), AI fallback otherwise.
function isRelevantWorkflowMessage_(subject, body, senderEmail, alreadyKnown){
  if(/\b(prompt|generated prompt)\b/i.test(subject) && alreadyKnown){
    return { type:'prompt', firmName:'', intakeJson:null };
  }
  if(!alreadyKnown && (/\bcrm\b/i.test(subject) || /deep\s*pocket/i.test(subject))){
    // Covers "CRM Enquiry", "CRM Software Enquiry", "Interested in CRM",
    // "Deep Pocket CRM Enquiry", "Need a CRM for my business", "CRM Demo
    // Request", "CRM Application Request", "CRM Pricing", etc.
    return { type:'enquiry', firmName:'', contactNo:'' };
  }
  return classifyEmail_(subject, body, senderEmail);
}

// Persistent record of message IDs this script has already acted on,
// stored in Script Properties rather than relying on Gmail's read/unread
// state or thread labels. Capped at the most recent 500 IDs.
function getProcessedIds_(){
  var raw = PropertiesService.getScriptProperties().getProperty('DP_PROCESSED_MSG_IDS');
  if(!raw) return [];
  try{ return JSON.parse(raw); }catch(e){ return []; }
}
function markProcessed_(msgId){
  var ids = getProcessedIds_();
  ids.push(msgId);
  if(ids.length>500) ids = ids.slice(ids.length-500);
  PropertiesService.getScriptProperties().setProperty('DP_PROCESSED_MSG_IDS', JSON.stringify(ids));
}

// ═══════════════════════════════════════════════
// SENT-MAIL PROCESSING — File Generated detection
// ═══════════════════════════════════════════════
function processSentMail_(){
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName(ENQUIRIES_TAB) || ss.insertSheet(ENQUIRIES_TAB);
  ensureEnquiriesHeader_(sh);

  try{
    var threads = GmailApp.search('in:sent subject:(File Generated) -label:DP-FileGenerated-Processed', 0, 20);
    if(!threads.length) return;
    var lblDone = getLabel_('DP-FileGenerated-Processed');

    threads.forEach(function(thread){
      try{
        var msgs = thread.getMessages();
        var sentMsg = null;
        for(var i=msgs.length-1;i>=0;i--){
          if(/file generated/i.test(msgs[i].getSubject()||'')){ sentMsg = msgs[i]; break; }
        }
        if(!sentMsg){ thread.addLabel(lblDone); return; }
        var toEmail = extractEmail_(sentMsg.getTo());
        var gmailRef = thread.getId()+'/'+sentMsg.getId();

        matchProspectAndAdvance_(ss, toEmail, '', '', 'File Generated', 13, 'Customized files sent (detected from Sent folder)', 'Automated', gmailRef);
        thread.addLabel(lblDone);
      }catch(e){ Logger.log('processSentMail_: error on thread: '+e.message); }
    });
  }catch(e){ Logger.log('processSentMail_: search failed: '+e.message); }
}

function extractEmail_(from){
  var m=String(from).match(/<(.+?)>/);
  return m ? m[1].toLowerCase() : String(from).trim().toLowerCase();
}
function getLabel_(name){ return GmailApp.getUserLabelByName(name) || GmailApp.createLabel(name); }

function extractIntakeMarker_(body){
  var m = body.match(/DEEPPOCKET_INTAKE_JSON:\s*(\{[\s\S]*\})\s*$/);
  if(!m) return null;
  try{ return JSON.parse(m[1]); }catch(e){ return null; }
}

function classifyEmail_(subject, body, senderEmail){
  var marker = extractIntakeMarker_(body);
  if(marker){
    return { type:'prompt', firmName: marker.firmName||'', contactNo: marker.phone||'', intakeJson: marker };
  }

  var apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if(!apiKey){
    Logger.log('classifyEmail_: GEMINI_API_KEY is not set in Script Properties — skipping AI classification for subject "'+subject+'"');
    return { type:'none' };
  }

  var sys = 'You triage inbound emails for a CRM-software sales inbox called Deep Pocket. Classify each email into exactly one type:\n'
    + '- "enquiry": a new prospective client asking about the CRM product - look for words/ideas like CRM, CRM software, customization, pricing, demo, interested, enquiry, etc.\n'
    + '- "prompt": a client\'s customization/intake submission - usually contains structured firm details like firm name, business model, products, features, sent after they were given a customization link.\n'
    + '- "none": anything else (spam, unrelated, newsletters, personal mail).\n\n'
    + 'Respond with ONLY a compact JSON object, no other text: {"type":"enquiry|prompt|none","firmName":"...","contactNo":"..."}\n'
    + 'Leave firmName/contactNo as empty strings if not present or not applicable.';
  var userMsg = 'Subject: ' + subject + '\nFrom: ' + senderEmail + '\n\nBody:\n' + body.slice(0,4000);

  var model = PropertiesService.getScriptProperties().getProperty('GEMINI_MODEL') || 'gemini-3.5-flash-lite';
  var url = 'https://generativelanguage.googleapis.com/v1beta/models/' + model + ':generateContent?key=' + apiKey;
  var payload = {
    systemInstruction: { parts: [{ text: sys }] },
    contents: [{ role:'user', parts: [{ text: userMsg }] }],
    generationConfig: { responseMimeType: 'application/json', maxOutputTokens: 300 }
  };

  try{
    var res = UrlFetchApp.fetch(url, {
      method:'post', contentType:'application/json',
      payload: JSON.stringify(payload), muteHttpExceptions:true
    });
    if(res.getResponseCode()!==200){
      Logger.log('classifyEmail_: Gemini returned HTTP '+res.getResponseCode()+' for subject "'+subject+'": '+res.getContentText().slice(0,500));
      return { type:'none' };
    }
    var data = JSON.parse(res.getContentText());
    var text = (data.candidates && data.candidates[0] && data.candidates[0].content
      && data.candidates[0].content.parts && data.candidates[0].content.parts[0]
      && data.candidates[0].content.parts[0].text) || '{}';
    var clean = text.replace(/```json|```/g,'').trim();
    return JSON.parse(clean);
  }catch(e){
    Logger.log('classifyEmail_: error classifying subject "'+subject+'": '+e.message);
    return { type:'none' };
  }
}

function autoReplyBody_(firmName){
  var formUrl = PropertiesService.getScriptProperties().getProperty('CLIENT_FORM_URL') || 'https://dp-crm.github.io/app/deep-pocket-crm-form.html';
  var name = (firmName||'').trim() || 'Sir/Madam';
  return 'Dear '+name+',\n\n'
    + 'Thank you so much for showing interest in our CRM web & mobile application service. We provide a complete customised CRM app for your business. Please help us to understand your business by filling the form. Form link is given below. We shall process your request immediately and deliver the required customised files to you.\n\n'
    + formUrl + '\n\n'
    + 'Regards,\n'
    + 'Team Deep Pocket Fintech Pvt. Ltd.';
}

function autoReplyBodyHtml_(firmName){
  var formUrl = PropertiesService.getScriptProperties().getProperty('CLIENT_FORM_URL') || 'https://dp-crm.github.io/app/deep-pocket-crm-form.html';
  var name = (firmName||'').trim() || 'Sir/Madam';
  return '<div style="font-family:Arial,Helvetica,sans-serif;font-size:14px;line-height:1.7;color:#222;max-width:480px;margin:0 auto">'
    + '<p style="margin:0 0 16px">Dear '+name+',</p>'
    + '<p style="margin:0 0 16px">Thank you so much for showing interest in our CRM web &amp; mobile application service. We provide a complete customised CRM app for your business.</p>'
    + '<p style="margin:0 0 24px">Please help us understand your business by filling the form below. We shall process your request immediately and deliver the required customised files to you.</p>'
    + '<p style="text-align:center;margin:0 0 24px">'
    +   '<a href="'+formUrl+'" style="background:#F4B942;color:#0A0E14;padding:13px 30px;border-radius:8px;text-decoration:none;font-weight:bold;font-size:14px;display:inline-block">Fill Customization Form</a>'
    + '</p>'
    + '<p style="font-size:11.5px;color:#777;margin:0 0 28px;word-break:break-all">Or copy this link: <a href="'+formUrl+'" style="color:#777">'+formUrl+'</a></p>'
    + '<p style="margin:0">Regards,<br>Team Deep Pocket Fintech Pvt. Ltd.</p>'
    + '</div>';
}

function sendAutoReply_(msg, senderEmail, firmName){
  var body = autoReplyBody_(firmName);
  var html = autoReplyBodyHtml_(firmName);
  var subject = 'Thank you for your enquiry';
  // GmailMessage.reply() does NOT support overriding the subject line — it
  // always sends as "Re: {original subject}", so the exact requested
  // subject can only be guaranteed by sending directly.
  GmailApp.sendEmail(senderEmail, subject, body, {htmlBody: html});
}
