// Wealth Matrix CRM — Apps Script backend v2.8 (idempotent / duplicate-proof)
// Changes vs v2.7: appendLead & appendClient are now UPSERTS (match by mobile),
// appendLog dedupes by Raw ID, and all writes are serialized with LockService.
// These four changes make every write safe to retry — the root fix for duplicate rows.

const SHEET_ID='1oPQMj6KB39M20RnEoix9PCvSA8RFhVhOV4BG2M4TXlQ';
const LEADS_TAB='LEADS',CLIENTS_TAB='CLIENTS',LOG_TAB='ACTIVITY_LOG',SPECIAL_TAB='Special Days',REMINDERS_TAB='REMINDERS',TASKS_TAB='TASKS',INSIGHTS_TAB='Insights',ADMIN_TAB='Admin';
const VERSION='2.17.0';

// ── Generate the next sequential human-readable ID (L1, L2... or C1, C2...).
// Scans the whole ID column, finds the highest existing number for this prefix,
// returns max+1. Robust against gaps from deleted rows — never reuses a number. ──
// ── SHA-256 hex hashing — used for all password storage/comparison. The
// client hashes the password before it's ever sent, and this hashes again
// for comparison against what's stored — meaning the plain-text password
// itself is never transmitted, logged, or stored anywhere, including here. ──
function sha256Hex(text){
  var digest=Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(text), Utilities.Charset.UTF_8);
  return digest.map(function(b){
    var v=(b<0?b+256:b).toString(16);
    return v.length===1?'0'+v:v;
  }).join('');
}

function nextSequentialId(sheet, colIndex, prefix){
  var lr=sheet.getLastRow();
  if(lr<2) return prefix+'1';
  var vals=sheet.getRange(2,colIndex,lr-1,1).getValues();
  var maxNum=0;
  var re=new RegExp('^'+prefix+'(\\d+)$','i');
  for(var i=0;i<vals.length;i++){
    var v=String(vals[i][0]||'').trim();
    var m=v.match(re);
    if(m){ var n=parseInt(m[1],10); if(n>maxNum) maxNum=n; }
  }
  return prefix+(maxNum+1);
}

// ── Ensure the header label exists for a given column, without disturbing
// anything else. Used to self-heal the two NEW columns (LEADS!T, CLIENTS!AK)
// on a live, already-populated sheet where the normal "write header row on
// first use" path (which only fires when the sheet is totally empty) never
// runs again. Safe to call on every relevant write — no-ops once set. ──
function ensureColumnHeader(sheet, colIndex, label){
  try{
    var cell=sheet.getRange(1,colIndex);
    if(!String(cell.getValue()||'').trim()){
      cell.setValue(label);
      cell.setFontWeight('bold').setBackground('#C0392B').setFontColor('#FFFFFF');
    }
  }catch(e){ /* non-fatal — header cosmetics only */ }
}

// ═══════════════════════════════════════════════════════════════
// AUTOMATED BUSINESS INSIGHTS — a scheduled trigger reads a compact summary
// of the CRM data, sends it to Gemini for a genuine writeup (not raw stats),
// then logs the result to an "Insights" tab and emails it to the script owner.
// ═══════════════════════════════════════════════════════════════

// Gathers a CONCISE summary — never raw row dumps — so the prompt sent to
// Gemini stays focused and the resulting insight is genuinely useful rather
// than a restatement of numbers the person already has in front of them.
function buildInsightsSummaryText(){
  var ss=SpreadsheetApp.openById(SHEET_ID);
  var lines=[];
  var today=new Date();
  function daysAgo(n){ var d=new Date(today); d.setDate(d.getDate()-n); return d; }
  function dateOnly(d){ return d instanceof Date ? new Date(d.getFullYear(),d.getMonth(),d.getDate()) : null; }

  // ── PRODUCTIVITY: funnel velocity, from the daily-bucketed ANALYTICS sheet.
  // This tracks NEW leads reaching each status each day — not current totals —
  // so it answers "how much did I generate" for today / this week / this month. ──
  var anaSh=ss.getSheetByName('ANALYTICS');
  if(anaSh && anaSh.getLastRow()>1){
    var anaRows=anaSh.getRange(2,1,anaSh.getLastRow()-1,7).getValues();
    var todayStr0=Utilities.formatDate(today,Session.getScriptTimeZone(),'yyyy-MM-dd');
    var week0=daysAgo(7), month0=daysAgo(30);
    var sums={today:[0,0,0,0,0], week:[0,0,0,0,0], month:[0,0,0,0,0]}; // [Prospects,Contacted,InProgress,Converted,Dropped]
    anaRows.forEach(function(r){
      var rDate=String(r[0]||'').trim();
      var isToday=rDate===todayStr0;
      var rd=rDate?new Date(rDate):null;
      var inWeek=rd && rd>=week0;
      var inMonth=rd && rd>=month0;
      for(var i=0;i<5;i++){
        var v=parseInt(r[i+2])||0;
        if(v<=0) continue; // only count positive (new) movements, not corrections/removals
        if(isToday) sums.today[i]+=v;
        if(inWeek) sums.week[i]+=v;
        if(inMonth) sums.month[i]+=v;
      }
    });
    var labels=['Prospects','Contacted','In Progress','Converted','Dropped'];
    function fmtSums(arr){ return labels.map(function(l,i){return l+'='+arr[i];}).join(', '); }
    lines.push('PRODUCTIVITY (new movements into each stage):');
    lines.push('Today: '+fmtSums(sums.today)+'.');
    lines.push('Last 7 days: '+fmtSums(sums.week)+'.');
    lines.push('Last 30 days: '+fmtSums(sums.month)+'.');
  }

  // ── Leads funnel + Potential AUM/SIP ──
  var lSh=ss.getSheetByName(LEADS_TAB);
  var highValueStaleLeads=[];
  if(lSh && lSh.getLastRow()>1){
    var lRows=lSh.getRange(2,1,lSh.getLastRow()-1,20).getValues();
    var statusCounts={}, staleCount=0, hotStale=[];
    var totalPotAum=0, totalPotSip=0;
    lRows.forEach(function(r){
      var status=String(r[5]||'').trim().toUpperCase();
      if(status) statusCounts[status]=(statusCounts[status]||0)+1;
      var modDate=r[13]instanceof Date?r[13]:(r[13]?new Date(r[13]):null);
      var daysSince=modDate?Math.floor((today-modDate)/86400000):999;
      var potAum=parseFloat(r[16])||0, potSip=parseFloat(r[8])||0; // col Q=potentialAum, col I=potentialSip(labelled aum internally)
      totalPotAum+=potAum; totalPotSip+=potSip;
      if(['PROSPECT','CONTACTED','IN PROGRESS'].indexOf(status)>=0 && daysSince>=14){
        staleCount++;
        if(daysSince>=21) hotStale.push(String(r[0]||'')+' ('+daysSince+'d)');
      }
      // "Connect potentials at least monthly" — high potential value, not touched in 30+ days
      if(['PROSPECT','CONTACTED','IN PROGRESS'].indexOf(status)>=0 && daysSince>=30 && (potAum>0||potSip>0)){
        highValueStaleLeads.push({name:String(r[0]||''), days:daysSince, potAum:potAum, potSip:potSip});
      }
    });
    lines.push('');
    lines.push('LEADS: '+lRows.length+' total. By status: '+Object.keys(statusCounts).map(function(k){return k+'='+statusCounts[k];}).join(', ')+'.');
    lines.push('Leads not followed up in 14+ days: '+staleCount+(hotStale.length?' (notably: '+hotStale.slice(0,8).join('; ')+')':'.'));
    lines.push('Total Potential AUM across active leads: \u20B9'+totalPotAum.toFixed(1)+'L. Total Potential SIP: \u20B9'+totalPotSip.toFixed(1)+'.');
  } else {
    lines.push(''); lines.push('LEADS: none yet.');
  }

  // ── Clients + family + stale contacts + Potential AUM/SIP ──
  var cSh=ss.getSheetByName(CLIENTS_TAB);
  if(cSh && cSh.getLastRow()>1){
    var cRows=cSh.getRange(2,1,cSh.getLastRow()-1,37).getValues();
    var stale90=[], familyMap={};
    var cTotalPotAum=0, cTotalPotSip=0;
    cRows.forEach(function(r){
      var name=String(r[0]||''), lastContacted=r[30], createdDate=r[10];
      var refDate=lastContacted instanceof Date?lastContacted:(createdDate instanceof Date?createdDate:null);
      var daysSince=refDate?Math.floor((today-refDate)/86400000):999;
      if(daysSince>90) stale90.push(name+' ('+daysSince+'d)');
      var fam=String(r[35]||'').trim();
      if(fam){ familyMap[fam]=familyMap[fam]||[]; familyMap[fam].push(name); }
      var potAum=parseFloat(r[27])||0, potSip=parseFloat(r[33])||0; // col AB=potentialAum, col AH=potentialSip
      cTotalPotAum+=potAum; cTotalPotSip+=potSip;
      // "Connect potentials at least monthly" — same check for clients with open potential
      if(daysSince>=30 && (potAum>0||potSip>0)){
        highValueStaleLeads.push({name:name, days:daysSince, potAum:potAum, potSip:potSip});
      }
    });
    lines.push('');
    lines.push('CLIENTS: '+cRows.length+' total.');
    lines.push('Clients not contacted in 90+ days: '+stale90.length+(stale90.length?' (notably: '+stale90.slice(0,8).join('; ')+')':'.'));
    lines.push('Total Potential AUM across clients (top-up opportunity): \u20B9'+cTotalPotAum.toFixed(1)+'L. Total Potential SIP: \u20B9'+cTotalPotSip.toFixed(1)+'.');
    var famEntries=Object.keys(familyMap).filter(function(f){return familyMap[f].length>1;});
    if(famEntries.length){
      lines.push('Families with 2+ tagged members: '+famEntries.slice(0,10).map(function(f){return f+' ('+familyMap[f].length+')';}).join(', ')+'.');
    }

    // Upcoming this week: reviews, SIP dates
    var upcoming=[];
    cRows.forEach(function(r){
      var name=String(r[0]||'');
      var reviewDate=r[16];
      if(reviewDate instanceof Date){
        var d=Math.floor((reviewDate-today)/86400000);
        if(d>=0 && d<=7) upcoming.push(name+': review in '+d+'d');
      }
      var sipDay=parseInt(r[20]);
      if(sipDay>=1 && sipDay<=28){
        var sipDate=new Date(today.getFullYear(),today.getMonth(),sipDay);
        if(sipDate<today) sipDate=new Date(today.getFullYear(),today.getMonth()+1,sipDay);
        var sd=Math.floor((sipDate-today)/86400000);
        if(sd<=7) upcoming.push(name+': SIP due in '+sd+'d');
      }
    });
    if(upcoming.length) lines.push('Upcoming this week: '+upcoming.slice(0,15).join('; ')+'.');
  } else {
    lines.push(''); lines.push('CLIENTS: none yet.');
  }

  // ── "Connect potentials at least once a month" — the direct answer to that ask.
  // Sorted by combined potential value so the highest-stakes names surface first. ──
  if(highValueStaleLeads.length){
    highValueStaleLeads.sort(function(a,b){ return (b.potAum+b.potSip)-(a.potAum+a.potSip); });
    lines.push('');
    lines.push('HIGH-POTENTIAL, NOT TOUCHED IN 30+ DAYS (overdue for a monthly check-in): '
      +highValueStaleLeads.slice(0,10).map(function(x){
        return x.name+' (\u20B9'+x.potAum.toFixed(1)+'L AUM / \u20B9'+x.potSip.toFixed(1)+' SIP, '+x.days+'d since contact)';
      }).join('; ')+'.');
  }

  // ── Pending reminders ──
  var rSh=ss.getSheetByName(REMINDERS_TAB);
  if(rSh && rSh.getLastRow()>1){
    var rRows=rSh.getRange(2,1,rSh.getLastRow()-1,4).getValues();
    var pending=rRows.filter(function(r){
      var rd=r[3] instanceof Date?r[3]:(r[3]?new Date(r[3]):null);
      return rd && rd<=today;
    });
    lines.push('');
    lines.push('Reminders currently due/overdue: '+pending.length+'.');
  }

  // ── Touch points — activity volume AND type breakdown, last 7 and 30 days ──
  var aSh=ss.getSheetByName(LOG_TAB);
  if(aSh && aSh.getLastRow()>1){
    var aRows=aSh.getRange(2,1,aSh.getLastRow()-1,6).getValues();
    var week0b=daysAgo(7), month0b=daysAgo(30);
    var last7=[], last30=[], typeCounts7={};
    aRows.forEach(function(r){
      var ts=String(r[1]||'');
      var d=ts?new Date(ts.split(' ')[0]):null;
      if(!d) return;
      if(d>=month0b) last30.push(r);
      if(d>=week0b){
        last7.push(r);
        var activity=String(r[4]||'').trim();
        // Group by the activity's emoji+label prefix as a rough touch-point type
        var typeKey=activity.split('-')[0].trim()||'Other';
        typeCounts7[typeKey]=(typeCounts7[typeKey]||0)+1;
      }
    });
    lines.push('');
    lines.push('TOUCH POINTS: '+last7.length+' in the last 7 days, '+last30.length+' in the last 30 days.');
    var typeList=Object.keys(typeCounts7).sort(function(a,b){return typeCounts7[b]-typeCounts7[a];});
    if(typeList.length){
      lines.push('Touch point types this week: '+typeList.slice(0,6).map(function(t){return t+'='+typeCounts7[t];}).join(', ')+'.');
    }
  }

  return lines.join('\n');
}

// Calls the Gemini API directly. Requires GEMINI_API_KEY to be set in
// Script Properties (Project Settings → Script Properties in the Apps
// Script editor) — never hardcoded in source, never transmitted anywhere
// except directly from this script to Google's Generative Language API.
// ── Ensure the Stage column (AF, 32nd) in CLIENTS has a dropdown data validation
// in the actual Google Sheet — so manual edits directly in the sheet are also
// constrained to the same options as the App's Stage dropdown. Self-healing:
// safe to call on every client write, only actually sets it if not already present. ──
// ── ONE-TIME SETUP for an EXISTING deployment: run this once manually from the
// Apps Script editor (select it in the function dropdown, click Run) to add the
// Stage dropdown to your existing CLIENTS sheet. New sheets get it automatically. ──
function setupStageDropdownNow(){
  var ss=SpreadsheetApp.openById(SHEET_ID);
  var sh=ss.getSheetByName(CLIENTS_TAB);
  if(!sh){ Logger.log('CLIENTS tab not found'); return; }
  ensureStageDropdown(sh);
  Logger.log('✅ Stage dropdown applied to CLIENTS column AF.');
}

function ensureStageDropdown(sheet){
  try{
    var stageOptions=['KYC PENDING','DOCUMENT PENDING','PROFILE CREATION PENDING','MANDATE PENDING','INVESTMENT PENDING','GRATITUDE CALL PENDING','COMPLETED'];
    var lastRow=Math.max(sheet.getLastRow(),1000); // cover existing rows plus headroom for new ones
    var range=sheet.getRange(2,32,lastRow-1,1); // column AF, from row 2 down
    var rule=SpreadsheetApp.newDataValidation().requireValueInList(stageOptions,true).setAllowInvalid(true).build();
    range.setDataValidation(rule);
  }catch(e){ /* non-fatal — dropdown is a convenience, not a data-integrity requirement */ }
}

// ═══════════════════════════════════════════════
// UNIFIED GEMINI CALLER — replaces every Claude call in this project.
// One shared function instead of 13 separate copies of the same
// fetch/parse pattern, each previously duplicated with Claude's request
// shape. Supports everything the old Claude calls needed:
//   opts.prompt            - a single-turn prompt (most calls)
//   opts.messages           - multi-turn [{role,content}] history (chat)
//   opts.systemInstruction  - a system prompt (Gemini's own equivalent)
//   opts.useWebSearch       - enables the google_search grounding tool
//   opts.documents          - [{base64, mimeType}] inline PDF documents
//   opts.maxOutputTokens    - output length cap
// Returns the response text directly, or a diagnostic string if the model
// genuinely returned no text (e.g. hit its token limit) — never a silent
// empty string, so a caller always has something explainable to show.
// ═══════════════════════════════════════════════
// ═══════════════════════════════════════════════
// DATA PROTECTION — file ownership verification for AI document analysis.
//
// Every function that reads a client's CAS PDF and sends it to Gemini
// previously trusted whatever fileId(s) the frontend sent, with no
// server-side check that the file actually belonged to the client named in
// the request. DriveApp.getFileById() can open ANY file the script has
// access to — not just files inside that one client's folder — so a stale
// ID, a UI race condition, or a future code change could otherwise cause
// one client's financial document (PAN, holdings, balances) to be read and
// analyzed under a DIFFERENT client's name, and potentially surfaced back
// to that wrong client or into their own review report.
//
// The fix: look up the client's OWN recorded folderId directly from the
// CLIENTS sheet (the source of truth, never the request payload), and only
// pass along files that are genuinely located inside that folder. Anything
// that doesn't match is silently excluded rather than trusted — the
// function proceeds with fewer files rather than a wrong one.
// ═══════════════════════════════════════════════

// Looks up a client's row by mobile OR Record ID and returns their own
// recorded folderId from the sheet — never from request data, which is
// exactly the thing this exists to not trust.
function getClientFolderId(mobile, recordId){
  var ss=SpreadsheetApp.openById(SHEET_ID);
  var sh=ss.getSheetByName(CLIENTS_TAB);
  if(!sh || sh.getLastRow()<2) return '';
  var vals=sh.getRange(2,1,sh.getLastRow()-1,42).getValues(); // A:AP, full CLIENT_COLS width
  var mobileClean=String(mobile||'').trim();
  var recordIdClean=String(recordId||'').trim();
  for(var i=0;i<vals.length;i++){
    var rowMobile=String(vals[i][1]||'').trim();   // col B = mobile
    var rowRecordId=String(vals[i][36]||'').trim(); // col AK = Record ID (0-indexed 36)
    if((recordIdClean && rowRecordId===recordIdClean) || (mobileClean && rowMobile===mobileClean)){
      return String(vals[i][34]||'').trim(); // col AI = Folder ID (0-indexed 34)
    }
  }
  return '';
}

// Given requested fileIds and the client's own verified folderId, returns
// only the File objects that are genuinely located inside that folder.
// If expectedFolderId is empty (client has no recorded folder), returns
// nothing — an unverifiable request trusts nothing rather than everything.
function verifyClientFileOwnership(fileIds, expectedFolderId){
  var verified=[];
  if(!expectedFolderId || !fileIds || !fileIds.length) return verified;
  fileIds.forEach(function(id){
    try{
      var f=DriveApp.getFileById(id);
      var parents=f.getParents();
      var belongs=false;
      while(parents.hasNext()){
        if(parents.next().getId()===expectedFolderId){ belongs=true; break; }
      }
      if(belongs) verified.push(f);
      // else: silently excluded — this file is not confirmed to belong to
      // this client, so it is never read or sent to the AI.
    }catch(e){ /* file missing or inaccessible — excluded, not trusted */ }
  });
  return verified;
}

function callGemini(opts){
  opts = opts || {};
  var apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if(!apiKey) throw new Error('GEMINI_API_KEY not set in Script Properties');

  var payload = { generationConfig: { maxOutputTokens: opts.maxOutputTokens || 800 } };

  if(opts.messages){
    // Multi-turn history — Gemini uses 'model' where Claude used 'assistant'.
    payload.contents = opts.messages.map(function(m){
      return { role: (m.role==='assistant'?'model':'user'), parts:[{ text: m.content }] };
    });
  } else if(opts.mixedContent){
    // Accepts Claude's own content-array shape directly — [{type:'document',
    // source:{data}}, {type:'text', text}, ...] — converted here rather than
    // rewriting the (substantial, file-looping) code at each call site that
    // builds these arrays from uploaded CAS PDFs.
    var parts2 = opts.mixedContent.map(function(c){
      if(c.type==='document') return { inlineData: { mimeType: (c.source&&c.source.media_type)||'application/pdf', data: c.source&&c.source.data } };
      return { text: c.text||'' };
    });
    payload.contents = [{ role:'user', parts: parts2 }];
  } else {
    var parts = [];
    if(opts.documents){
      opts.documents.forEach(function(d){
        parts.push({ inlineData: { mimeType: d.mimeType||'application/pdf', data: d.base64 } });
      });
    }
    parts.push({ text: opts.prompt||'' });
    payload.contents = [{ role:'user', parts: parts }];
  }

  if(opts.systemInstruction){
    payload.systemInstruction = { parts:[{ text: opts.systemInstruction }] };
  }
  if(opts.useWebSearch){
    payload.tools = [{ google_search: {} }];
  }

  var resp = UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/models/gemini-3.5-flash-lite:generateContent', {
    method:'post', contentType:'application/json',
    headers:{ 'x-goog-api-key': apiKey },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });
  var json = JSON.parse(resp.getContentText());
  if(json.error) throw new Error(json.error.message || 'Gemini API error');
  var candidate = (json.candidates||[])[0];
  var textPart = candidate && candidate.content && candidate.content.parts
    ? candidate.content.parts.find(function(p){ return p.text; }) : null;
  if(textPart) return textPart.text.trim();
  return '(no text returned — finishReason: '+(candidate?candidate.finishReason:'unknown')+')';
}

function callGeminiAPI(summaryText){
  var prompt='You are a productivity analyst for Wealth Matrix Finserv, an Indian wealth-management advisory firm. '
    +'The advisor measures their own productivity by funnel velocity (how many prospects, contacted, and in-progress leads '
    +'are generated per day/week/month, and how many convert), how consistently they re-engage high-potential leads and '
    +'clients (at least once a month), and touch-point frequency. Below is a compact summary of the current CRM state. '
    +'Write a short, direct briefing (under 320 words) for the advisor, covering: '
    +'(1) funnel velocity \u2014 is this week/month\'s pace up or down, and where is the bottleneck stage; '
    +'(2) Potential AUM/SIP holders who have gone 30+ days without contact \u2014 name them, they are overdue for the monthly touch; '
    +'(3) touch-point pattern \u2014 volume and whether it is concentrated in one type or spread out; '
    +'(4) any family cross-sell opportunity visible in the data. '
    +'End with ONE specific, decisive recommendation labeled "This week, do this:" \u2014 the single highest-value action, '
    +'naming a specific person or group, not generic advice. '
    +'Be specific and reference names/numbers from the data throughout. Use short paragraphs or a tight bullet list, '
    +'no headers, no preamble.\n\nCRM SUMMARY:\n'+summaryText;
  return callGemini({ prompt: prompt, maxOutputTokens: 600 });
}

// Extracts a first name (Title Case) from a full name stored in UPPERCASE,
// and computes a time-aware greeting based on the CURRENT hour in IST —
// computed here deterministically rather than left for the AI to guess,
// since it has no reliable access to the real current time.
function getFirstName(fullName){
  var first=(fullName||'').trim().split(/\s+/)[0]||'';
  return first.charAt(0).toUpperCase()+first.slice(1).toLowerCase();
}
function getTimeAwareGreeting(){
  var hour=parseInt(Utilities.formatDate(new Date(), 'Asia/Kolkata', 'H'), 10);
  if(hour<12) return 'Good morning';
  if(hour<17) return 'Good afternoon';
  return 'Good evening'; // covers evening through late night, including after 8 PM
}

// The orchestrator — called by the weekly trigger, or manually for testing.
// Drafts a short, personalized WhatsApp message for one specific lead.
// Returns the draft text only — never sends anything itself. Sending always
// happens from the app after the advisor has reviewed (and possibly edited) it.
function draftLeadMessage(data){
  var today=new Date();
  var modDate=data.modifiedDate?new Date(data.modifiedDate):null;
  var daysSince=modDate?Math.floor((today-modDate)/86400000):null;

  // ── Activity log deep-dive: full history for this specific lead, not just
  // the last few entries — so the draft has real continuity of everything
  // that's actually been discussed, not a shallow recent snapshot. ──
  var historyText='';
  try{
    var ss=SpreadsheetApp.openById(SHEET_ID);
    var logSh=ss.getSheetByName(LOG_TAB);
    if(logSh && logSh.getLastRow()>1){
      var logRows=logSh.getRange(2,1,logSh.getLastRow()-1,7).getValues();
      var mobileClean=String(data.mobile||'').trim();
      var leadEntries=logRows.filter(function(r){
        return String(r[2]||'').trim()===mobileClean && String(r[3]||'').trim().toUpperCase()==='LEAD';
      });
      // Sort chronologically, oldest first — needed to identify the FIRST activity
      leadEntries.sort(function(a,b){ return new Date(a[0])-new Date(b[0]); });

      if(leadEntries.length>0){
        var firstEntry=leadEntries[0];
        var createdDate=data.createdDate||'';
        historyText+='Lead created: '+createdDate+'. First activity logged: '+firstEntry[0]+' ('+firstEntry[4]+(firstEntry[5]?' — '+firstEntry[5]:'')+').\n';

        // Full chronological history, capped at 20 entries to keep the prompt
        // focused — a genuine deep-dive without unbounded growth for very
        // long-running leads.
        var capped=leadEntries.slice(-20);
        historyText+='Full conversation history ('+leadEntries.length+' logged interactions'+(leadEntries.length>20?', showing most recent 20':'')+'):\n';
        capped.forEach(function(r){
          historyText+='- '+r[0]+': '+r[4]+(r[5]?' — '+r[5]:'')+'\n';
        });
      } else {
        historyText='Lead created: '+(data.createdDate||'')+'. No activity logged yet — this would be the first outreach.\n';
      }
    }
  }catch(e){
    console.warn('Activity history lookup failed:', e.message);
  }

  var context='Name: '+(data.name||'')+'\n'
    +'Status: '+(data.status||'')+'\n'
    +'Source: '+(data.source||'')+(data.referral?' (referred by '+data.referral+')':'')+'\n'
    +(data.country?'Country: '+data.country+'\n':'')
    +(data.products?'Interested products: '+data.products+'\n':'')
    +(data.potentialAum?'Potential AUM: \u20B9'+data.potentialAum+'L\n':'')
    +(data.potentialSip?'Potential SIP: \u20B9'+data.potentialSip+'\n':'')
    +(daysSince!==null?'Days since last contact: '+daysSince+'\n':'')
    +(data.remarks?'Advisor notes: '+data.remarks+'\n':'')
    +(data.familyName?'Family: '+data.familyName+' (already has family members with the firm)\n':'')
    +'\n'+historyText;

  var firstName=getFirstName(data.name);
  var greeting=getTimeAwareGreeting();
  var prompt='You are drafting a short WhatsApp message from a wealth management advisor at Wealth Matrix Finserv '
    +'to a LEAD (not yet a client) based on the context below, which includes their FULL conversation history.\n\n'
    +'FORMAT — this is mandatory: the very first line must be exactly \'Hi '+firstName+', '+greeting+'!\' followed by '
    +'a blank line, then the message itself.\n\n'
    +'TONE — this matters a lot, more than anything else here: the message must read as a real human texting on their '
    +'phone, not as AI-written text. Concretely: use contractions (I\'ll, don\'t, that\'s) the way people actually type. '
    +'Vary sentence length — a short fragment followed by a longer one is normal in texting, perfectly balanced sentences '
    +'are not. Avoid em-dashes and semicolons entirely — people don\'t text with them. Avoid AI-sounding openers and '
    +'transitions ("I hope this message finds you", "I wanted to reach out", "just checking in to see"). Short, casual, '
    +'natural everyday words — not corporate or formal phrasing. This must NOT read as a sales pitch or a broadcast/spam '
    +'message — it should feel fresh, personal, and specific to THIS person, like the advisor genuinely remembered '
    +'something about them. Keep the message itself (after the greeting) as short as a real manually-typed text — 2 to '
    +'3 short sentences at most, not a paragraph. Reference something SPECIFIC from what has actually been discussed '
    +'before — a concern they raised, a topic they asked about, timing they mentioned — do not write a generic template, '
    +'and do not repeat something already fully addressed. Do NOT promise any specific returns, performance, or give '
    +'investment advice. No emojis unless it reads naturally. If their Country is listed above and is outside India, '
    +'phrase any next step accordingly — a call or video meeting reads naturally, but do not casually suggest meeting '
    +'"in person" or dropping by the office as if they were local. End with a soft next step (a call, a quick meeting, or '
    +'sharing more info) — not a hard sell.\n\n'
    +'Output ONLY the greeting line, a blank line, then the message — nothing else, no preamble, no quotation marks, '
    +'no "Here is a draft:".'
    +'\n\nLEAD CONTEXT:\n'+context;

  return callGemini({ prompt: prompt, maxOutputTokens: 220 });
}

// Drafts a short, personalized WhatsApp message for one specific CLIENT.
// Mirrors draftLeadMessage but with client-specific context (actual AUM,
// SIP/review dates, tenure) and retention/engagement framing rather than
// conversion framing. Returns the draft text only — never sends anything.
function draftClientMessage(data){
  var today=new Date();
  var lastContacted=data.lastContacted?new Date(data.lastContacted):null;
  var daysSince=lastContacted?Math.floor((today-lastContacted)/86400000):null;
  var convertedDate=data.convertedDate?new Date(data.convertedDate):null;
  var tenureDays=convertedDate?Math.floor((today-convertedDate)/86400000):null;

  // ── Activity log deep-dive — same approach as leads, querying the CLIENT
  // side of the log so the draft has real continuity of the relationship. ──
  var historyText='';
  try{
    var ss=SpreadsheetApp.openById(SHEET_ID);
    var logSh=ss.getSheetByName(LOG_TAB);
    if(logSh && logSh.getLastRow()>1){
      var logRows=logSh.getRange(2,1,logSh.getLastRow()-1,7).getValues();
      var mobileClean=String(data.mobile||'').trim();
      var clientEntries=logRows.filter(function(r){
        return String(r[2]||'').trim()===mobileClean && String(r[3]||'').trim().toUpperCase()==='CLIENT';
      });
      clientEntries.sort(function(a,b){ return new Date(a[0])-new Date(b[0]); });

      if(clientEntries.length>0){
        var firstEntry=clientEntries[0];
        historyText+='Became a client: '+(data.convertedDate||'')+'. First activity logged as a client: '+firstEntry[0]+' ('+firstEntry[4]+(firstEntry[5]?' — '+firstEntry[5]:'')+').\n';
        var capped=clientEntries.slice(-20);
        historyText+='Full conversation history ('+clientEntries.length+' logged interactions'+(clientEntries.length>20?', showing most recent 20':'')+'):\n';
        capped.forEach(function(r){
          historyText+='- '+r[0]+': '+r[4]+(r[5]?' — '+r[5]:'')+'\n';
        });
      } else {
        historyText='Became a client: '+(data.convertedDate||'')+'. No activity logged yet as a client.\n';
      }
    }
  }catch(e){
    console.warn('Activity history lookup failed:', e.message);
  }

  var context='Name: '+(data.name||'')+'\n'
    +'Stage: '+(data.stage||'')+'\n'
    +(data.age?'Age: '+data.age+'\n':'')
    +(data.gender?'Gender: '+data.gender+'\n':'')
    +(data.country?'Country: '+data.country+'\n':'')
    +(data.products?'Products held: '+data.products+'\n':'')
    +(data.aum?'Current AUM: \u20B9'+data.aum+'L\n':'')
    +(data.potentialAum?'Potential AUM (top-up opportunity): \u20B9'+data.potentialAum+'L\n':'')
    +(data.potentialSip?'Potential SIP (not yet started): \u20B9'+data.potentialSip+'\n':'')
    +(daysSince!==null?'Days since last contact: '+daysSince+'\n':'')
    +(tenureDays!==null?'Client for: '+Math.floor(tenureDays/30)+' months\n':'')
    +(data.sipDate?'SIP date: '+data.sipDate+'th of the month\n':'')
    +(data.reviewDate?'Next review date: '+data.reviewDate+'\n':'')
    +(data.dob?'Date of birth: '+data.dob+'\n':'')
    +(data.anniversary?'Wedding anniversary: '+data.anniversary+'\n':'')
    +(data.healthInsurance?'Health insurance: '+data.healthInsurance+'\n':'')
    +(data.lifeInsurance?'Life insurance: '+data.lifeInsurance+'\n':'')
    +(data.willDone?'Will done: '+data.willDone+'\n':'')
    +(data.remarks?'Advisor notes (may include family/profession context): '+data.remarks+'\n':'')
    +(data.familyName?'Family: '+data.familyName+' (other family members also with the firm)\n':'')
    +'\n'+historyText;

  var firstName=getFirstName(data.name);
  var greeting=getTimeAwareGreeting();
  var prompt='You are drafting a short WhatsApp message from a wealth management advisor at Wealth Matrix Finserv '
    +'to an EXISTING CLIENT based on the RICH context below, which includes their FULL conversation history, personal '
    +'details, and protection status. This is about relationship maintenance, re-engagement, or a natural top-up/cross-sell '
    +'opportunity — NOT converting them, they are already a client.\n\n'
    +'FORMAT — this is mandatory: the very first line must be exactly \'Hi '+firstName+', '+greeting+'!\' followed by '
    +'a blank line, then the message itself.\n\n'
    +'TONE — this matters a lot, more than anything else here: the message must read as a real human texting on their '
    +'phone, not as AI-written text. Concretely: use contractions (I\'ll, don\'t, that\'s) the way people actually type. '
    +'Vary sentence length — a short fragment followed by a longer one is normal in texting, perfectly balanced sentences '
    +'are not. Avoid em-dashes and semicolons entirely — people don\'t text with them. Avoid AI-sounding openers and '
    +'transitions ("I hope this message finds you", "I wanted to reach out", "just checking in to see"). Short, casual, '
    +'natural everyday words — not corporate or formal phrasing. This must NOT read as a sales pitch or a broadcast/spam '
    +'message — it should feel fresh, personal, and specific to THIS person, like the advisor genuinely remembered '
    +'something about them. Keep the message itself (after the greeting) as short as a real manually-typed text — 2 to '
    +'3 short sentences at most, not a paragraph. Reference something SPECIFIC from what has actually been discussed '
    +'before, or a specific upcoming date (SIP, review) if relevant — do not write a generic template. Do NOT promise '
    +'any specific returns, performance, or give investment advice. No emojis unless it reads naturally. If their Country '
    +'is listed above and is outside India, phrase any next step accordingly — a call or video meeting reads naturally, '
    +'but do not casually suggest meeting "in person" or dropping by the office as if they were local. End with a '
    +'soft next step — not a hard sell.\n\n'
    +'Output ONLY the greeting line, a blank line, then the message — nothing else, no preamble, no quotation marks, '
    +'no "Here is a draft:".'
    +'\n\nCLIENT CONTEXT:\n'+context;

  return callGemini({ prompt: prompt, maxOutputTokens: 220 });
}

// ── AI Insights (per-person deep profile) — distinct from the WhatsApp
// drafting functions above: this is an internal analysis FOR the advisor,
// not a message TO the person. Pulls the fullest available picture: every
// tracked field plus the complete activity log, with special attention to
// the Remarks field since the advisor manually adds family/profession
// context there. ──
function buildFullActivityHistory(mobile, type){
  var ss=SpreadsheetApp.openById(SHEET_ID);
  var logSh=ss.getSheetByName(LOG_TAB);
  var out='';
  if(logSh && logSh.getLastRow()>1){
    var logRows=logSh.getRange(2,1,logSh.getLastRow()-1,7).getValues();
    var mobileClean=String(mobile||'').trim();
    var entries=logRows.filter(function(r){
      return String(r[2]||'').trim()===mobileClean && String(r[3]||'').trim().toUpperCase()===type.toUpperCase();
    });
    entries.sort(function(a,b){ return new Date(a[0])-new Date(b[0]); });
    if(entries.length>0){
      out+='First logged activity: '+entries[0][0]+' ('+entries[0][4]+(entries[0][5]?' — '+entries[0][5]:'')+').\n';
      out+='Full history ('+entries.length+' interactions'+(entries.length>25?', showing most recent 25':'')+'):\n';
      entries.slice(-25).forEach(function(r){
        out+='- '+r[0]+': '+r[4]+(r[5]?' — '+r[5]:'')+'\n';
      });
    } else {
      out='No activity logged yet.\n';
    }
  }
  return out;
}

function generateLeadProfileInsight(data){
  var history=buildFullActivityHistory(data.mobile, 'LEAD');
  var context='Name: '+(data.name||'')+'\n'
    +'Status: '+(data.status||'')+'\n'
    +'Source: '+(data.source||'')+(data.referral?' (referred by '+data.referral+')':'')+'\n'
    +(data.age?'Age: '+data.age+'\n':'')
    +(data.country?'Country: '+data.country+'\n':'')
    +(data.products?'Interested products: '+data.products+'\n':'')
    +(data.contactMode?'Preferred contact mode: '+data.contactMode+'\n':'')
    +(data.potentialAum?'Potential AUM: \u20B9'+data.potentialAum+'L\n':'')
    +(data.potentialSip?'Potential SIP: \u20B9'+data.potentialSip+'\n':'')
    +(data.geoTag?'Location: '+data.geoTag+'\n':'')
    +(data.familyName?'Family: '+data.familyName+' (has family members already with the firm)\n':'')
    +(data.remarks?'\nADVISOR REMARKS (may contain family, profession, and other personal context the advisor has noted):\n'+data.remarks+'\n':'')
    +'\nACTIVITY HISTORY:\n'+history;

  var prompt='You are helping a wealth management advisor at Wealth Matrix Finserv understand a LEAD (not yet converted) '
    +'in depth, using everything tracked about them below. This is an internal profile for the advisor\'s own understanding — '
    +'not a message to send. Write a genuine, specific profile covering: (1) who this person appears to be and what matters '
    +'to them, drawing on the remarks and activity history — infer their profession, family situation, or goals if the notes '
    +'suggest them, but do not invent anything not supported by the data; (2) their engagement pattern — responsive, hesitant, '
    +'going cold, etc.; (3) the clearest conversion opportunity and what seems to be the actual blocker, if any; (4) one '
    +'concrete suggestion for how to move this relationship forward. Under 350 words, specific, no generic advice, no headers, '
    +'no preamble.\n\nLEAD DATA:\n'+context;

  return callGemini({ prompt: prompt, maxOutputTokens: 600 });
}

function generateClientProfileInsight(data){
  var history=buildFullActivityHistory(data.mobile, 'CLIENT');
  var context='Name: '+(data.name||'')+'\n'
    +'Stage: '+(data.stage||'')+'\n'
    +(data.age?'Age: '+data.age+'\n':'')
    +(data.gender?'Gender: '+data.gender+'\n':'')
    +(data.country?'Country: '+data.country+'\n':'')
    +(data.products?'Products held: '+data.products+'\n':'')
    +(data.aum?'Current AUM: \u20B9'+data.aum+'L\n':'')
    +(data.potentialAum?'Potential AUM (top-up opportunity): \u20B9'+data.potentialAum+'L\n':'')
    +(data.potentialSip?'Potential SIP (not yet started): \u20B9'+data.potentialSip+'\n':'')
    +(data.convertedDate?'Client since: '+data.convertedDate+'\n':'')
    +(data.sipDate?'SIP date: '+data.sipDate+'th of the month\n':'')
    +(data.reviewDate?'Next review date: '+data.reviewDate+'\n':'')
    +(data.dob?'Date of birth: '+data.dob+'\n':'')
    +(data.anniversary?'Wedding anniversary: '+data.anniversary+'\n':'')
    +(data.healthInsurance?'Health insurance: '+data.healthInsurance+'\n':'')
    +(data.lifeInsurance?'Life insurance: '+data.lifeInsurance+'\n':'')
    +(data.willDone?'Will done: '+data.willDone+'\n':'')
    +(data.familyName?'Family: '+data.familyName+' (other family members also with the firm)\n':'')
    +(data.remarks?'\nADVISOR REMARKS (may contain family, profession, and other personal context the advisor has noted):\n'+data.remarks+'\n':'')
    +'\nACTIVITY HISTORY:\n'+history;

  var prompt='You are helping a wealth management advisor at Wealth Matrix Finserv understand an EXISTING CLIENT '
    +'in depth, using everything tracked about them below. This is an internal profile for the advisor\'s own understanding — '
    +'not a message to send. Write a genuine, specific profile covering: (1) who this person appears to be — profession, '
    +'family situation, life stage — drawing on the remarks and activity history; infer where the notes support it, but do '
    +'not invent anything unsupported; (2) protection and portfolio gaps worth addressing (missing insurance, no will, '
    +'unused potential AUM/SIP); (3) the relationship\'s health — engagement frequency, any drift; (4) one concrete, '
    +'specific next step for the advisor. Under 350 words, specific, no generic advice, no headers, no preamble.'
    +'\n\nCLIENT DATA:\n'+context;

  return callGemini({ prompt: prompt, maxOutputTokens: 600 });
}

// ── Lists candidate CAS PDFs in a client's folder with their upload dates,
// WITHOUT running any AI analysis — used to let the advisor pick which
// file(s) to actually analyze when more than one is found. ──
function listCasFilesInFolder(data){
  var folder=null;
  if(data.folderId){ try{ folder=DriveApp.getFolderById(data.folderId); }catch(fe){ folder=null; } }
  if(!folder) throw new Error('No folder found for this client');

  var files=[];
  var iter=folder.getFiles();
  while(iter.hasNext()){
    var f=iter.next();
    if(f.getName().toUpperCase().indexOf('CAS')>=0 && f.getMimeType()==='application/pdf'){
      files.push({
        id: f.getId(),
        name: f.getName(),
        uploadDate: Utilities.formatDate(f.getDateCreated(), Session.getScriptTimeZone(), 'dd MMM yyyy')
      });
    }
  }
  files.sort(function(a,b){ return new Date(b.uploadDate)-new Date(a.uploadDate); }); // newest first
  return files;
}

// ── AI Portfolio Email Draft — reads every CAS PDF in the client's folder
// (there may be more than one: CAS is generated PER EMAIL, not per PAN, so a
// client with holdings across multiple registered emails needs several files
// combined). Explicitly instructs Gemini to isolate only this client's PAN,
// since a single CAS file can contain OTHER investors who share that email
// (e.g. a spouse) — critical, or their holdings would wrongly get mixed in. ──
function draftCasEmailInsight(data){
  // The client's OWN recorded folder — looked up from the sheet, never
  // taken from the request. This is what every file below gets checked
  // against, regardless of what the frontend asked for.
  var verifiedFolderId=getClientFolderId(data.mobile, data.recordId);

  var casFiles=[];
  if(data.fileIds && data.fileIds.length){
    // Explicit file selection from the user — but still verified against
    // this client's own folder before any of them get read. A fileId that
    // doesn't belong to this client is silently excluded, not trusted.
    casFiles=verifyClientFileOwnership(data.fileIds, verifiedFolderId);
    if(casFiles.length===0) throw new Error('None of the selected files could be verified as belonging to this client — they may have been moved, deleted, or the request referenced the wrong client');
  } else {
    if(!verifiedFolderId) throw new Error('No folder found for this client');
    var folder=DriveApp.getFolderById(verifiedFolderId);

    var iter=folder.getFiles();
    while(iter.hasNext()){
      var f=iter.next();
      if(f.getName().toUpperCase().indexOf('CAS')>=0 && f.getMimeType()==='application/pdf'){
        casFiles.push(f);
      }
    }
    if(casFiles.length===0){
      throw new Error('No CAS PDF found in this client\'s folder — make sure the filename contains "CAS" (e.g. "CAS Report.pdf")');
    }
  }

  var content=[];
  casFiles.forEach(function(f){
    var b64=Utilities.base64Encode(f.getBlob().getBytes());
    content.push({type:'document', source:{type:'base64', media_type:'application/pdf', data:b64}});
    content.push({type:'text', text:'The above CAS document is named "'+f.getName()+'".'});
  });

  var instructions='You are analyzing CAMS/KFintech Consolidated Account Statement (CAS) document(s) for '
    +(data.clientName||'this investor')+' (PAN: '+(data.pan||'not provided')+') to give their wealth advisor a brief, '
    +'insightful portfolio summary. IMPORTANT: CAS documents are generated per registered EMAIL, not per '
    +'PAN — if multiple documents are provided above, they may represent the SAME investor\'s holdings split across '
    +'different registered emails and should be COMBINED into one consolidated view. ALSO IMPORTANT: a single CAS document '
    +'can contain MULTIPLE different investors (different PANs) if they share one registered email — you MUST identify and '
    +'use ONLY the holdings belonging to PAN '+(data.pan||'[not provided — use the name to identify the correct investor]')
    +'. Do not include any other investor\'s holdings in the analysis. PRIVACY REQUIREMENT: the PAN above is provided '
    +'solely so you can correctly identify which holdings belong to this investor within the document — it must be used '
    +'for that internal matching purpose ONLY. Do not reproduce the PAN, in full or in part, anywhere in your output '
    +'(neither the summary nor the table). Refer to the investor by name only.\n\n'
    +'From the correctly-isolated holdings, determine: total number of distinct funds/schemes, total invested/current value, '
    +'and for each fund: SIP active (yes/no), balance/value, and demat or non-demat. '
    +'If the number of distinct funds exceeds 10, explicitly recommend reducing the count — the ideal range is 4 to 7 funds, '
    +'since over-diversification can reduce long-term returns without meaningfully reducing risk. If 10 or fewer, do not force this point.\n\n'
    +'Output your response in EXACTLY this structure, with these exact markers on their own lines:\n'
    +'---SUMMARY---\n'
    +'(a brief, specific 3-5 sentence summary — total funds, total value, the fund-count observation if applicable, '
    +'one genuine insight. Professional, direct tone. No investment return promises.)\n'
    +'---TABLE---\n'
    +'(an HTML table, using this exact inline-styled format: '
    +'<table style="border-collapse:collapse;width:100%;font-family:Arial,sans-serif;font-size:13px">'
    +'<tr style="background:#1a1a1a;color:#fff"><th style="border:1px solid #ccc;padding:6px;text-align:left">Sl No</th>'
    +'<th style="border:1px solid #ccc;padding:6px;text-align:left">Investor Name</th>'
    +'<th style="border:1px solid #ccc;padding:6px;text-align:left">Fund Name</th>'
    +'<th style="border:1px solid #ccc;padding:6px;text-align:left">SIP Active</th>'
    +'<th style="border:1px solid #ccc;padding:6px;text-align:left">Balance</th>'
    +'<th style="border:1px solid #ccc;padding:6px;text-align:left">Demat/Non-Demat</th></tr>'
    +'then one <tr> with plain <td style="border:1px solid #ccc;padding:6px"> cells per fund, alternating row background '
    +'#ffffff and #f2f2f2 for readability.)\n'
    +'---END---\n'
    +'Output ONLY this structure, nothing before ---SUMMARY--- or after ---END---.';

  content.push({type:'text', text:instructions});

  var raw = callGemini({ mixedContent: content, maxOutputTokens: 1500 }).trim();

  var subjectMatch=raw.match(/SUBJECT:\s*(.+)/);
  var bodyMatch=raw.match(/---BODY---([\s\S]*?)---TABLE---/);
  var tableMatch=raw.match(/---TABLE---([\s\S]*?)---END---/);

  return {
    subject: subjectMatch?subjectMatch[1].trim():('Portfolio Summary — '+(data.clientName||'')),
    body: bodyMatch?bodyMatch[1].trim():'',
    tableHtml: tableMatch?tableMatch[1].trim():'',
    filesUsed: casFiles.map(function(f){return f.getName();})
  };
}

// ── Meeting Insights — a detailed, AI-generated pre-meeting briefing for one
// client. Combines: CAS PDF analysis (with a growth comparison if 2 CAS files
// are supplied), family cross-reference (count of other Leads/Clients sharing
// the same Family Name), engagement frequency (from ACTIVITY_LOG), portfolio
// gaps (missing Health/Life insurance, WILL), vintage (from Created Date), an
// approximate revenue estimate (₹600 per lakh AUM), and SIF/AIF cross-sell
// framing. Referral count comes from LEADS whose "Referred By" matches this
// client's name. ──
function countFamilyMembers(familyName){
  if(!familyName) return {leads:0, clients:0};
  var famUp=String(familyName).trim().toUpperCase();
  var ss=SpreadsheetApp.openById(SHEET_ID);
  var leadCount=0, clientCount=0;
  var lSh=ss.getSheetByName(LEADS_TAB);
  if(lSh && lSh.getLastRow()>1){
    var lVals=lSh.getRange(2,19,lSh.getLastRow()-1,1).getValues(); // column S = Family Name
    lVals.forEach(function(r){ if(String(r[0]||'').trim().toUpperCase()===famUp) leadCount++; });
  }
  var cSh=ss.getSheetByName(CLIENTS_TAB);
  if(cSh && cSh.getLastRow()>1){
    var cVals=cSh.getRange(2,36,cSh.getLastRow()-1,1).getValues(); // column AJ = Family Name
    cVals.forEach(function(r){ if(String(r[0]||'').trim().toUpperCase()===famUp) clientCount++; });
  }
  return {leads:leadCount, clients:clientCount};
}

function countReferralsByName(clientName){
  if(!clientName) return 0;
  var nameUp=String(clientName).trim().toUpperCase();
  var ss=SpreadsheetApp.openById(SHEET_ID);
  var count=0;
  var lSh=ss.getSheetByName(LEADS_TAB);
  if(lSh && lSh.getLastRow()>1){
    var refVals=lSh.getRange(2,8,lSh.getLastRow()-1,1).getValues(); // column H = Referred By
    refVals.forEach(function(r){ if(String(r[0]||'').trim().toUpperCase()===nameUp) count++; });
  }
  var cSh=ss.getSheetByName(CLIENTS_TAB);
  if(cSh && cSh.getLastRow()>1){
    var crefVals=cSh.getRange(2,6,cSh.getLastRow()-1,1).getValues(); // column F = Referred By
    crefVals.forEach(function(r){ if(String(r[0]||'').trim().toUpperCase()===nameUp) count++; });
  }
  return count;
}

function generateMeetingInsight(data){
  var today=new Date();
  var createdDate=data.createdDate?new Date(data.createdDate):null;
  var vintageDays=createdDate?Math.floor((today-createdDate)/86400000):null;
  var vintageText=vintageDays!==null?Math.floor(vintageDays/30)+' months ('+Math.floor(vintageDays/365)+'y '+Math.floor((vintageDays%365)/30)+'m)':'unknown';

  var family=countFamilyMembers(data.familyName);
  var referralCount=countReferralsByName(data.name);
  var history=buildFullActivityHistory(data.mobile, 'CLIENT');

  // Revenue estimate — ₹600 per lakh of AUM, AUM entered in lakhs
  var aumLakh=parseFloat(data.aum)||0;
  var estRevenue=aumLakh*600;

  // ── Read CAS PDF(s) if provided — 2 files means oldest+newest for a growth
  // comparison; 1 file means latest-only (outstanding balance pull).
  // Verified against this client's own recorded folder before reading —
  // an unverified fileId is excluded rather than trusted. ──
  var casContent=[];
  var casNote='No CAS document was available for analysis.';
  if(data.fileIds && data.fileIds.length){
    var verifiedFolderId=getClientFolderId(data.mobile, data.recordId);
    var casFiles=verifyClientFileOwnership(data.fileIds, verifiedFolderId);
    casFiles.forEach(function(f, idx){
      var b64=Utilities.base64Encode(f.getBlob().getBytes());
      casContent.push({type:'document', source:{type:'base64', media_type:'application/pdf', data:b64}});
      var label=casFiles.length===2?(idx===0?'OLDER (for comparison)':'LATEST'):'LATEST';
      casContent.push({type:'text', text:'The above CAS document ('+f.getName()+') is labeled: '+label+'.'});
    });
    casNote=casFiles.length===2
      ? 'TWO CAS documents provided — one older, one latest. Compare total AUM/outstanding balance between them to show growth or reduction.'
      : casFiles.length===1
        ? 'ONE CAS document provided (latest only). Pull the current outstanding balance across all mutual fund holdings — no comparison is possible.'
        : 'A CAS document was requested but could not be verified as belonging to this client, so none was included.';
  }

  var context='CLIENT PROFILE:\n'
    +'Name: '+(data.name||'')+'\n'
    +(data.age?'Age: '+data.age+'\n':'')
    +(data.familyName?'Family: '+data.familyName+' ('+family.clients+' other client(s), '+family.leads+' lead(s) tagged to this family)\n':'Family: not tagged\n')
    +'Referrals generated by this client (leads/clients referred by them): '+referralCount+'\n'
    +'AUM on file: \u20B9'+(data.aum||'0')+'L. Estimated advisory revenue at \u20B9600/lakh: \u20B9'+estRevenue.toFixed(0)+'\n'
    +(data.potentialAum?'Potential AUM (top-up opportunity, not yet invested): \u20B9'+data.potentialAum+'L\n':'')
    +(data.potentialSip?'Potential SIP (not yet started): \u20B9'+data.potentialSip+'\n':'')
    +'Products currently held: '+(data.products||'none recorded')+'\n'
    +'Health insurance: '+(data.healthInsurance||'not recorded')+'\n'
    +'Life insurance: '+(data.lifeInsurance||'not recorded')+'\n'
    +'WILL done: '+(data.willDone?'Yes':'No/not recorded')+'\n'
    +'Client since (vintage): '+vintageText+'\n'
    +(data.remarks?'\nADVISOR REMARKS (may reference family, kids, profession, income — use these to infer cross-sell potential):\n'+data.remarks+'\n':'')
    +'\nACTIVITY HISTORY (engagement frequency signal):\n'+history
    +'\nCAS NOTE: '+casNote;

  var panIsolationNote=casContent.length ? ('IMPORTANT — CAS documents are generated per registered EMAIL, not per PAN: a single '
    +'document can contain MULTIPLE different investors if they share one registered email. You MUST identify and use ONLY the '
    +'holdings belonging to PAN '+(data.pan||'[not provided — use the client name above to identify the correct investor, and '
    +'if the document contains more than one plausible match, say so explicitly rather than guessing]')+'. Do not include any '
    +'other investor\'s holdings anywhere in this briefing. PRIVACY REQUIREMENT: the PAN is provided solely for this internal '
    +'matching purpose — do not reproduce it, in full or in part, anywhere in your output.\n\n') : '';

  var promptText=panIsolationNote
    +'You are preparing a wealth management advisor at Wealth Matrix Finserv for an upcoming in-person or review '
    +'meeting with an EXISTING CLIENT. Produce a detailed, structured pre-meeting briefing using ONLY the data provided below '
    +'plus the attached CAS PDF(s) if present. This is an internal document for the advisor\'s own preparation — not a message '
    +'to the client.\n\n'
    +'Cover EXACTLY these sections, in this order, with clear headers:\n'
    +'1) CLIENT SNAPSHOT — name, age, family, vintage with the firm.\n'
    +'2) PORTFOLIO POSITION — from the CAS document(s): total AUM/outstanding balance, number of distinct funds/schemes, '
    +'SIP status per fund if visible. '+casNote+' If two CAS documents are provided, explicitly state whether total AUM has '
    +'grown or reduced between the older and latest statement, with approximate figures and the difference.\n'
    +'3) REVENUE — state the AUM on file and the estimated advisory revenue at \u20B9600 per lakh, calculated for you above; do not recompute it.\n'
    +'4) PROTECTION & PORTFOLIO GAPS — call out specifically which of Health Insurance, Life Insurance, and WILL are missing '
    +'or unconfirmed, and any unused Potential AUM/SIP.\n'
    +'5) CROSS-SELL OPPORTUNITY — explicitly consider whether Specialized Investment Funds (SIF) and Alternate Investment '
    +'Funds (AIF) are worth exploring with this client given their AUM level and profile, and say so directly if relevant.\n'
    +'6) FAMILY & REFERRALS — family members already with the firm, and referrals generated by this client so far.\n'
    +'7) ENGAGEMENT PATTERN — how frequently this client has been contacted, based on the activity history; flag if contact '
    +'has gone quiet.\n'
    +'8) FAMILY CONTEXT FROM REMARKS — if the remarks mention kids, their age, education, job, or income, infer future cross-'
    +'sell potential (e.g. a child nearing college age, a child starting to earn) — only state what is reasonably supported '
    +'by the notes, do not invent specifics.\n'
    +'9) RECOMMENDED TALKING POINTS — 3 to 5 specific, concrete points for this exact meeting, not generic advice.\n\n'
    +'Be specific and numeric wherever the data supports it. No preamble, no closing remarks — start directly with section 1.\n\n'
    +context;

  var content=casContent.slice();
  content.push({type:'text', text:promptText});

  return callGemini({ mixedContent: content, maxOutputTokens: 1800 });
}

// ── Standalone Productivity data — same daily-bucketed ANALYTICS logic already
// used in the weekly briefing, extracted so the on-demand Productivity icon
// can call it directly without generating the full weekly summary. ──
function buildProductivityData(){
  var ss=SpreadsheetApp.openById(SHEET_ID);
  var today=new Date();
  function daysAgo(n){ var d=new Date(today); d.setDate(d.getDate()-n); return d; }
  var anaSh=ss.getSheetByName('ANALYTICS');
  var sums={today:[0,0,0,0,0], week:[0,0,0,0,0], month:[0,0,0,0,0]};
  var labels=['Prospects','Contacted','In Progress','Converted','Dropped'];
  if(anaSh && anaSh.getLastRow()>1){
    var anaRows=anaSh.getRange(2,1,anaSh.getLastRow()-1,7).getValues();
    var todayStr0=Utilities.formatDate(today,Session.getScriptTimeZone(),'yyyy-MM-dd');
    var week0=daysAgo(7), month0=daysAgo(30);
    anaRows.forEach(function(r){
      var rDate=String(r[0]||'').trim();
      var isToday=rDate===todayStr0;
      var rd=rDate?new Date(rDate):null;
      var inWeek=rd && rd>=week0;
      var inMonth=rd && rd>=month0;
      for(var i=0;i<5;i++){
        var v=parseInt(r[i+2])||0;
        if(v<=0) continue;
        if(isToday) sums.today[i]+=v;
        if(inWeek) sums.week[i]+=v;
        if(inMonth) sums.month[i]+=v;
      }
    });
  }
  function fmtSums(arr){ return labels.map(function(l,i){return l+'='+arr[i];}).join(', '); }
  return 'PRODUCTIVITY (new movements into each pipeline stage):\n'
    +'Today: '+fmtSums(sums.today)+'.\n'
    +'Last 7 days: '+fmtSums(sums.week)+'.\n'
    +'Last 30 days ("last month"): '+fmtSums(sums.month)+'.';
}

function generateProductivityInsight(){
  var summary=buildProductivityData();
  var prompt='You are a productivity analyst for a wealth management advisor. Below is a compact summary of new leads '
    +'reaching each pipeline stage (Prospects, Contacted, In Progress, Converted, Dropped) — today, in the last 7 days, '
    +'and in the last 30 days. Write a short, direct briefing (under 220 words) covering: (1) today\'s pace at a glance, '
    +'(2) whether this week\'s pace is on track compared to the monthly average, (3) where the funnel bottleneck appears '
    +'to be (which stage has the fewest new movements relative to the stage before it), (4) the conversion rate implied '
    +'by these numbers (Converted vs total new Prospects over the last 30 days). End with ONE specific, decisive '
    +'recommendation labeled "Do this:" — the single highest-value action for the advisor this week. Be specific with '
    +'numbers throughout. No headers, no preamble, plain text only (this will be shown in a simple text pop-up).'
    +'\n\n'+summary;
  return callGemini({ prompt: prompt, maxOutputTokens: 500 });
}

// ── Engagement comparison — activity log volume in the last 7 and last 14
// days, compared against the SAME LENGTH windows exactly 30 days earlier (i.e.
// "the same period last month"), plus a per-client breakdown so it's clear
// WHO is being engaged with versus who's gone quiet, not just an aggregate. ──
function buildEngagementComparisonData(){
  var ss=SpreadsheetApp.openById(SHEET_ID);
  var today=new Date();
  function daysAgo(n){ var d=new Date(today); d.setDate(d.getDate()-n); d.setHours(0,0,0,0); return d; }
  var logSh=ss.getSheetByName(LOG_TAB);
  if(!logSh || logSh.getLastRow()<2) return 'No activity log entries recorded yet.';

  var rows=logSh.getRange(2,1,logSh.getLastRow()-1,7).getValues();
  var w7start=daysAgo(7), w14start=daysAgo(14);
  var w7prevStart=daysAgo(37), w7prevEnd=daysAgo(30);   // "same 7-day window, 30 days earlier"
  var w14prevStart=daysAgo(44), w14prevEnd=daysAgo(30); // "same 14-day window, 30 days earlier"

  var counts={last7:0,last14:0,prev7:0,prev14:0};
  var perClientLast14={}; // name -> count, for the "who's being engaged" breakdown

  rows.forEach(function(r){
    var ts=String(r[0]||'');
    var d=ts?new Date(ts.split(' ')[0]):null;
    if(!d) return;
    var name=String(r[1]||'').trim()||'(unnamed)';
    if(d>=w7start) counts.last7++;
    if(d>=w14start) counts.last14++;
    if(d>=w7prevStart && d<w7prevEnd) counts.prev7++;
    if(d>=w14prevStart && d<w14prevEnd) counts.prev14++;
    if(d>=w14start){ perClientLast14[name]=(perClientLast14[name]||0)+1; }
  });

  // Most-engaged and least-engaged (touched only once) in the last 14 days
  var sortedClients=Object.keys(perClientLast14).sort(function(a,b){return perClientLast14[b]-perClientLast14[a];});
  var mostEngaged=sortedClients.slice(0,5).map(function(n){return n+' ('+perClientLast14[n]+')';});
  var leastEngaged=sortedClients.filter(function(n){return perClientLast14[n]===1;}).slice(0,8);

  function pctChange(now,prev){
    if(prev===0) return now>0?'new activity (none in the comparable period last month)':'no change (both zero)';
    var pct=Math.round(((now-prev)/prev)*100);
    return (pct>=0?'+':'')+pct+'% vs the same period last month';
  }

  var lines=[];
  lines.push('ENGAGEMENT (activity log entries — calls, notes, WhatsApp sends, status changes, etc.):');
  lines.push('Last 7 days: '+counts.last7+' entries. Same 7-day window one month ago: '+counts.prev7+' entries. Change: '+pctChange(counts.last7,counts.prev7)+'.');
  lines.push('Last 14 days: '+counts.last14+' entries. Same 14-day window one month ago: '+counts.prev14+' entries. Change: '+pctChange(counts.last14,counts.prev14)+'.');
  lines.push('');
  lines.push('Most-engaged clients/leads in the last 14 days (name and entry count): '+(mostEngaged.length?mostEngaged.join(', '):'none'));
  lines.push('Clients/leads touched only ONCE in the last 14 days (may need a follow-up): '+(leastEngaged.length?leastEngaged.join(', '):'none'));
  return lines.join('\n');
}

function generateEngagementInsight(){
  var summary=buildEngagementComparisonData();
  var prompt='You are an engagement analyst for a wealth management advisor. Below is a comparison of how much client/lead '
    +'contact activity (calls, notes, WhatsApp messages, status changes) happened in the last 7 and last 14 days, versus '
    +'the SAME LENGTH window exactly one month earlier — plus which specific people got the most attention and which got '
    +'only a single touch. Write a short, direct briefing (under 220 words) covering: (1) whether engagement is trending '
    +'up or down compared to a month ago, with the actual percentage change stated, (2) whether the 7-day and 14-day '
    +'trends agree or tell different stories, (3) name the specific people who got only one touch in 14 days and may be '
    +'at risk of going cold. End with ONE specific, decisive recommendation labeled "Do this:" naming a specific person '
    +'or group where possible, not generic advice. No headers, no preamble, plain text only (shown in a simple text pop-up).'
    +'\n\n'+summary;
  return callGemini({ prompt: prompt, maxOutputTokens: 500 });
}

// ═══════════════════════════════════════════════
// BUSINESS INSIGHTS — replaces the old Quick Insights / Weekly Automation
// page. Each computes real data directly from the sheets, then has Gemini
// write a short, actionable narrative around it — same pattern as
// generateProductivityInsight/generateEngagementInsight above.
// ═══════════════════════════════════════════════

// Shared helper: calls Gemini with a prompt, returns plain text.
function callGeminiForInsight(prompt, maxTokens){
  return callGemini({ prompt: prompt, maxOutputTokens: maxTokens||500 });
}

// ── Gemini 3.5 Flash-Lite — used specifically for the Top 10 Business
// Insights AI-writing step (the least reasoning-heavy, non-client-facing
// call in the app), as a targeted cost experiment. Deliberately kept as its
// own separate function rather than modifying callGeminiForInsight, so
// nothing else in the app — chat, WhatsApp drafts, Meeting Insights, Review
// Reports — is touched by this at all. ──
function callGeminiFlashLite(prompt, maxOutputTokens){
  var apiKey=PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if(!apiKey) throw new Error('GEMINI_API_KEY not set in Script Properties');
  var resp=UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/models/gemini-3.5-flash-lite:generateContent',{
    method:'post', contentType:'application/json',
    headers:{'x-goog-api-key':apiKey},
    payload:JSON.stringify({
      contents:[{parts:[{text:prompt}]}],
      generationConfig:{maxOutputTokens:maxOutputTokens||500}
    }),
    muteHttpExceptions:true
  });
  var json=JSON.parse(resp.getContentText());
  if(json.error) throw new Error(json.error.message||'Gemini API error');
  var candidate=(json.candidates||[])[0];
  var textPart=candidate && candidate.content && candidate.content.parts
    ? candidate.content.parts.find(function(p){return p.text;}) : null;
  if(textPart) return textPart.text.trim();
  // Mirrors the same diagnostic principle used for the Claude fallback —
  // surface what actually happened (e.g. finishReason:'MAX_TOKENS') instead
  // of a bare, undiagnosable empty response.
  return '(no text returned — finishReason: '+(candidate?candidate.finishReason:'unknown')+')';
}

// ═══════════════════════════════════════════════
// TOP BUSINESS INSIGHTS — replaces the old three labeled categories with a
// single ranked list of individual findings, top 10 by priority. Each check
// below is a self-contained function returning scored candidates — kept
// deliberately modular so priority weights and check types can become
// configurable per-business later (this app is intended to be white-labeled
// for other firms eventually, per the CRM blueprint generator artifact).
// ═══════════════════════════════════════════════

// Check: clients with a birthday in 12-14 days AND AUM over 100 lakhs (₹1cr+).
// High priority — combines urgency (a real date) with high account value.
function checkBirthdayHighValue(ss){
  var out=[];
  var cSh=ss.getSheetByName(CLIENTS_TAB);
  if(!cSh || cSh.getLastRow()<2) return out;
  var vals=cSh.getRange(2,1,cSh.getLastRow()-1,41).getValues();
  var today=new Date(); today.setHours(0,0,0,0);
  vals.forEach(function(r){
    var name=String(r[0]||''); if(!name) return;
    var familyName=String(r[35]||''); // AJ = Family Name
    var dobRaw=r[12]; // M = DOB
    var aum=parseFloat(r[26])||0; // AA = AUM (in lakhs)
    if(!dobRaw || aum<=100) return;
    var dob=dobRaw instanceof Date?dobRaw:new Date(dobRaw);
    if(isNaN(dob.getTime())) return;
    var thisYear=new Date(today.getFullYear(), dob.getMonth(), dob.getDate());
    if(thisYear<today) thisYear=new Date(today.getFullYear()+1, dob.getMonth(), dob.getDate());
    var daysUntil=Math.round((thisYear-today)/86400000);
    if(daysUntil>=12 && daysUntil<=14){
      out.push({
        type:'birthday_high_value', name:name, familyName:familyName, priority:90,
        facts:name+' has a birthday coming up in '+daysUntil+' days, and manages an AUM of approximately \u20b9'+aum+' lakhs.'
      });
    }
  });
  return out;
}

// Check: dropped leads not contacted in 6+ months — surfaces 1-2 as a single
// "reconnect" prompt, not every dropped lead (that would be noise).
function checkDroppedReengage(ss){
  var out=[];
  var lSh=ss.getSheetByName(LEADS_TAB);
  if(!lSh || lSh.getLastRow()<2) return out;
  var vals=lSh.getRange(2,1,lSh.getLastRow()-1,21).getValues();
  var sixMonthsAgo=new Date(); sixMonthsAgo.setMonth(sixMonthsAgo.getMonth()-6);
  var candidates=[];
  vals.forEach(function(r){
    var status=String(r[5]||'').trim().toUpperCase(); // F = Status
    if(status!=='DROPPED') return;
    var modRaw=r[13]; // N = Modified Date
    var mod=modRaw instanceof Date?modRaw:(modRaw?new Date(modRaw):null);
    if(mod && !isNaN(mod.getTime()) && mod<=sixMonthsAgo){
      candidates.push({name:String(r[0]||''), familyName:String(r[18]||'')}); // S = Family Name
    }
  });
  if(candidates.length){
    var pick=candidates.slice(0,2);
    out.push({
      type:'dropped_reengage', name:pick.map(function(p){return p.name;}).join(' & '), familyName:'', priority:50,
      facts:'These dropped leads have not been contacted in over 6 months: '+pick.map(function(p){return p.name;}).join(', ')+'.'
    });
  }
  return out;
}

// Check: leads stuck in Prospect/Contacted/In Progress for 60+ days — reuses
// the existing computeStaleLeads(), surfaces up to 3 individually.
function checkStaleLeadsCandidates(){
  var stale=computeStaleLeads();
  return stale.slice(0,3).map(function(l){
    return { type:'stale_lead', name:l.name, familyName:'', priority: l.potentialAum?70:55,
      facts:l.name+' has been in '+l.status+' for '+l.daysStale+' days without progressing'+(l.potentialAum?' (Potential AUM: '+l.potentialAum+')':'')+'.' };
  });
}

// Check: clients missing health insurance, life insurance, or a WILL —
// surfaces up to 2, prioritizing higher-AUM clients.
function checkProtectionGapsCandidates(ss){
  var out=[];
  var cSh=ss.getSheetByName(CLIENTS_TAB);
  if(!cSh || cSh.getLastRow()<2) return out;
  var vals=cSh.getRange(2,1,cSh.getLastRow()-1,41).getValues();
  var gaps=[];
  vals.forEach(function(r){
    var name=String(r[0]||''); if(!name) return;
    var health=String(r[21]||'').trim().toUpperCase();
    var life=String(r[22]||'').trim().toUpperCase();
    var will=String(r[23]||'').trim();
    var missing=[];
    if(health!=='YES') missing.push('Health Insurance');
    if(life!=='YES') missing.push('Life Insurance');
    if(!will || will.toUpperCase()==='NO') missing.push('WILL');
    if(missing.length) gaps.push({name:name, familyName:String(r[35]||''), aum:parseFloat(r[26])||0, missing:missing.join(', ')});
  });
  gaps.sort(function(a,b){return b.aum-a.aum;});
  return gaps.slice(0,2).map(function(g){
    return { type:'protection_gap', name:g.name, familyName:g.familyName, priority:g.aum>100?75:60,
      facts:g.name+' is missing: '+g.missing+(g.aum?' (AUM: \u20b9'+g.aum+' lakhs)':'')+'.' };
  });
}

// Check: active leads with zero activity log entries — a process failure,
// surfaces up to 2.
function checkNeverContactedCandidates(ss){
  var out=[];
  var lSh=ss.getSheetByName(LEADS_TAB);
  if(!lSh || lSh.getLastRow()<2) return out;
  var vals=lSh.getRange(2,1,lSh.getLastRow()-1,21).getValues();
  var activeStatuses=['PROSPECT','CONTACTED','IN PROGRESS'];
  var leadsById={};
  vals.forEach(function(r){
    var status=String(r[5]||'').trim().toUpperCase();
    if(activeStatuses.indexOf(status)<0) return;
    var recordId=String(r[17]||'').trim();
    if(recordId) leadsById[recordId]={name:String(r[0]||''), familyName:String(r[18]||'')};
  });
  var logSh=ss.getSheetByName(LOG_TAB);
  if(logSh && logSh.getLastRow()>1){
    var logVals=logSh.getRange(2,8,logSh.getLastRow()-1,1).getValues();
    logVals.forEach(function(r){ var rid=String(r[0]||'').trim(); if(rid && leadsById[rid]) delete leadsById[rid]; });
  }
  var never=Object.keys(leadsById).map(function(k){return leadsById[k];});
  return never.slice(0,2).map(function(l){
    return { type:'never_contacted', name:l.name, familyName:l.familyName, priority:80,
      facts:l.name+' has zero activity logged since being added — no call, note, or message has ever been recorded.' };
  });
}

// ═══════════════════════════════════════════════
// PRACTICE CHAT — "Ask your practice anything." Replaces the fixed top-10
// insights list with a genuine, open-ended conversation grounded in the
// firm's real data. Full context is gathered once per chat session (sent
// as part of the first turn) and carried forward via conversation history
// on subsequent turns, the standard pattern for the Messages API.
// ═══════════════════════════════════════════════

// Returns TWO separate blocks, deliberately:
//   stable   — leads/clients/tasks rosters. Changes only when records are
//              actually added or edited, so between turns of a normal
//              conversation it stays byte-identical and the prompt cache
//              genuinely hits (~10% of input cost instead of 100%).
//   volatile — today's date + recent activity. Small, and changes often.
// Previously these were concatenated into ONE block with the date at the
// top, which meant the cached block was almost never byte-identical between
// calls — so it repeatedly paid the 1.25x cache-WRITE premium and rarely
// got the read discount. Splitting them is what makes caching actually pay.
function gatherPracticeContext(){
  var ss=SpreadsheetApp.openById(SHEET_ID);
  var parts=[];

  // Leads — compact one-line-per-lead summary
  var lSh=ss.getSheetByName(LEADS_TAB);
  if(lSh && lSh.getLastRow()>=2){
    var lVals=lSh.getRange(2,1,lSh.getLastRow()-1,21).getValues();
    var leadLines=lVals.filter(function(r){return r[0];}).map(function(r){
      return '- '+r[0]+' | status:'+r[5]+' | source:'+r[6]+' | potentialAUM:'+(r[16]||'-')+' | potentialSIP:'+(r[8]||'-')
        +' | createdDate:'+fmtCell(r[14])+' | modifiedDate:'+fmtCell(r[13])+' | family:'+(r[18]||'-');
    });
    parts.push('LEADS ('+leadLines.length+' total):\n'+leadLines.join('\n'));
  }

  // Clients — compact one-line-per-client summary
  var cSh=ss.getSheetByName(CLIENTS_TAB);
  if(cSh && cSh.getLastRow()>=2){
    var cVals=cSh.getRange(2,1,cSh.getLastRow()-1,41).getValues();
    var clientLines=cVals.filter(function(r){return r[0];}).map(function(r){
      return '- '+r[0]+' | stage:'+(r[31]||'-')+' | AUM:'+(r[26]||'-')+'L | potentialAUM:'+(r[27]||'-')+'L | products:'+(r[29]||'-')
        +' | health:'+(r[21]||'-')+' | life:'+(r[22]||'-')+' | WILL:'+(r[23]||'-')+' | DOB:'+fmtCell(r[12])
        +' | SIPdate:'+(r[20]||'-')+' | reviewDate:'+fmtCell(r[16])+' | family:'+(r[35]||'-')+' | createdDate:'+fmtCell(r[10]);
    });
    parts.push('CLIENTS ('+clientLines.length+' total):\n'+clientLines.join('\n'));
  }

  // Tasks — compact one-line-per-task summary. getTasksSheet() is nested
  // inside _writeDataInner and not reachable from here, so the same
  // TASKS/REMINDERS fallback lookup is inlined directly.
  var tSh=ss.getSheetByName(TASKS_TAB)||ss.getSheetByName(REMINDERS_TAB);
  if(tSh && tSh.getLastRow()>=2){
    var tVals=tSh.getRange(2,1,tSh.getLastRow()-1,6).getValues();
    var taskLines=tVals.filter(function(r){return r[1];}).map(function(r){
      return '- '+r[1]+' | details:'+r[2]+' | due:'+fmtCell(r[3])+' | status:'+r[4]+' | for:'+r[5];
    });
    parts.push('TASKS ('+taskLines.length+' total):\n'+taskLines.join('\n'));
  }

  // Activity Log — the volatile half. Trimmed from 300 to 150 entries: each
  // record is already capped to its own last 10 by the existing trim job, so
  // 150 still covers a wide slice of recent practice activity while roughly
  // halving what was by far the largest and most frequently-changing part of
  // the payload.
  var volatileParts=[];
  var logSh=ss.getSheetByName(LOG_TAB);
  if(logSh && logSh.getLastRow()>=2){
    var logVals=logSh.getRange(2,1,logSh.getLastRow()-1,6).getValues(); // A-F: Timestamp,Name,Mobile,Type,Activity,Notes
    var logEntries=logVals.filter(function(r){return r[1]&&r[4];}).map(function(r){
      return { ts:String(r[0]||''), name:String(r[1]||''), type:String(r[3]||''), activity:String(r[4]||''), notes:String(r[5]||'') };
    });
    logEntries.sort(function(a,b){ return b.ts.localeCompare(a.ts); }); // most recent first
    var recentLog=logEntries.slice(0,150);
    var logLines=recentLog.map(function(e){
      return '- ['+e.ts+'] '+e.name+' ('+e.type+'): '+e.activity+(e.notes?' — '+e.notes:'');
    });
    volatileParts.push('RECENT ACTIVITY LOG (most recent '+logLines.length+' of '+logEntries.length+' total interactions):\n'+logLines.join('\n'));
  }

  var today=Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return {
    stable: parts.join('\n\n'),
    volatile: 'TODAY\'S DATE: '+today+'\n\n'+volatileParts.join('\n\n')
  };
}

function fmtCell(v){
  if(!v) return '-';
  if(v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return String(v);
}

// ── Contemporary Insight — the one chat action that actually searches the
// live web, rather than reasoning only over the practice's own data. Kept as
// its own dedicated function (not folded into practiceChat) deliberately:
// enabling web search on every single chat message would add real latency
// and cost to questions that don't need it. This fires only when the
// advisor explicitly asks for it. ──
function generateContemporaryInsight(data){
  var prompt='Search for the most recent, genuinely significant financial and economic news relevant to a wealth management '
    +'advisor and their clients in India and the GCC region (UAE, Saudi Arabia, Qatar, Kuwait, Bahrain, Oman) — things like '
    +'Union Budget announcements, RBI monetary policy decisions (repo rate changes, policy stance), major market-moving '
    +'regulatory changes, or significant GCC economic developments. Focus on what has ACTUALLY happened most recently, not '
    +'general background. Write a short, sharp briefing (under 220 words): (1) the 2-3 most relevant recent developments, '
    +'stated plainly with the actual dates/figures involved, (2) for each, one direct sentence on why it might actually '
    +'matter for client conversations right now — not vague relevance, a real reason. Write like a sharp, well-read '
    +'colleague sharing what\'s actually worth knowing this week, not a news digest or a corporate report. If nothing '
    +'genuinely significant has happened recently, say so plainly rather than manufacturing importance. No headers, plain '
    +'text only.';

  var reply = callGemini({ prompt: prompt, useWebSearch: true, maxOutputTokens: 2000 });
  return reply || '(No answer came back. Try asking again.)';
}

function practiceChat(data){
  var history=data.history||[]; // [{role:'user'|'assistant', content:'...'}, ...] — plain Q&A only, never contains embedded context
  var messages=[];
  // Prior turns carried forward exactly as stored (capped, so a very long
  // session doesn't grow the payload unboundedly), then the new question.
  var capped=history.slice(-20);
  capped.forEach(function(h){ messages.push({role:h.role, content:h.content}); });
  messages.push({role:'user', content:data.message});

  // Gathered fresh on every call (not just the first turn) — a failed first
  // turn must not leave the rest of the session permanently blind, answering
  // like a generic assistant with no CRM connection at all.
  var context=gatherPracticeContext();
  var systemPrompt='You are a sharp, warm, genuinely knowledgeable practice-management assistant for a wealth management '
    +'advisor at Wealth Matrix Finserv. You have full visibility into their real leads, clients, tasks, AND their actual '
    +'logged activity history (calls, notes, status changes, messages sent) below — use it specifically, by name, not '
    +'generically. When asked what was discussed with someone, what happened recently, or how engaged a person has been, '
    +'draw on the actual activity log entries, not just their static status. Be genuinely insightful: notice patterns, '
    +'connect dots across records and interaction history, and offer real perspective, not just data lookup. Write like '
    +'a sharp, warm colleague talking through the practice with them — natural, direct, and specific — never like a '
    +'corporate report or a generic chatbot.\n\n'
    +'LENGTH — aim for under 200 words. Be genuinely selective rather than exhaustive: lead with the most important, '
    +'most useful point the data actually supports, then add only what genuinely strengthens it. Depth on the thing '
    +'that matters beats a shallow pass across everything. For a simple factual question, a two-line answer is the '
    +'right answer — don\'t pad it to fill space.\n\n'
    +'If the data doesn\'t support a confident answer, say so plainly rather than guessing. Never invent a name, '
    +'figure, or interaction that isn\'t in the data below.\n\nPRACTICE DATA (rosters):\n'+context.stable
    +'\n\nCURRENT CONTEXT (changes frequently):\n'+context.volatile;

  var reply = callGemini({ messages: messages, systemInstruction: systemPrompt, maxOutputTokens: 1500 });
  return reply || '(No answer came back. Try asking again, or a more specific question.)';
}

function generateTopBusinessInsights(){
  var ss=SpreadsheetApp.openById(SHEET_ID);
  var candidates=[]
    .concat(checkBirthdayHighValue(ss))
    .concat(checkDroppedReengage(ss))
    .concat(checkStaleLeadsCandidates())
    .concat(checkProtectionGapsCandidates(ss))
    .concat(checkNeverContactedCandidates(ss));

  if(!candidates.length) return [];

  candidates.sort(function(a,b){ return b.priority-a.priority; });
  var top10=candidates.slice(0,10);

  // One batched AI call for all of them — far cheaper and faster than one
  // call per finding, and the person using this app is cost-conscious about
  // exactly this (hence the usage-cost confirmation shown before this runs).
  var factsList=top10.map(function(c,i){ return (i+1)+'. '+c.facts; }).join('\n');
  var prompt='You are writing short, warm, actionable insight notes for a wealth advisor\'s CRM dashboard. Below are '+top10.length
    +' factual findings, numbered. For EACH one, write ONE short sentence (max 30 words) turning the fact into a natural, '
    +'actionable prompt for the advisor — conversational, not robotic. Examples of tone: "has a birthday coming up — worth '
    +'planning a small gesture?" or "hasn\'t been reached in months — a quick check-in could bring them back."\n\n'
    +'Respond with ONLY a JSON array of '+top10.length+' strings, one per numbered finding, in the same order. No other text, '
    +'no markdown, just the raw JSON array.\n\n'+factsList;

  var raw=callGeminiFlashLite(prompt, 900);
  var blurbs;
  try{
    var jsonMatch=raw.match(/\[[\s\S]*\]/);
    blurbs=JSON.parse(jsonMatch?jsonMatch[0]:raw);
  }catch(e){
    blurbs=top10.map(function(c){return c.facts;}); // fallback to the raw fact if parsing fails — never show nothing
  }

  return top10.map(function(c,i){
    return { name:c.name, familyName:c.familyName, insight: blurbs[i]||c.facts, type:c.type };
  });
}

// ── Stale Leads — Prospect/Contacted/In Progress for 60+ days, never
// converted or dropped. Returns both the raw list (for the delete-capable
// drill-down) and an AI-written narrative. ──
function computeStaleLeads(){
  var ss=SpreadsheetApp.openById(SHEET_ID);
  var lSh=ss.getSheetByName(LEADS_TAB);
  if(!lSh || lSh.getLastRow()<2) return [];
  var lr=lSh.getLastRow();
  var vals=lSh.getRange(2,1,lr-1,21).getValues();
  var today=new Date(); today.setHours(0,0,0,0);
  var cutoff=new Date(today); cutoff.setDate(cutoff.getDate()-60);
  var activeStatuses=['PROSPECT','CONTACTED','IN PROGRESS'];
  var out=[];
  vals.forEach(function(r){
    var status=String(r[5]||'').trim().toUpperCase(); // F = Status
    if(activeStatuses.indexOf(status)<0) return;
    var createdRaw=r[14]; // O = Created Date
    var created=createdRaw instanceof Date ? createdRaw : (createdRaw?new Date(createdRaw):null);
    if(!created || isNaN(created.getTime()) || created>cutoff) return;
    var daysStale=Math.floor((today-created)/86400000);
    out.push({
      name:String(r[0]||''), mobile:String(r[1]||''), status:status,
      recordId:String(r[17]||''), daysStale:daysStale,
      potentialAum:String(r[16]||''), potentialSip:String(r[8]||'')
    });
  });
  out.sort(function(a,b){ return b.daysStale-a.daysStale; });
  return out;
}

function generateStaleLeadsInsight(){
  var stale=computeStaleLeads();
  if(!stale.length) return {insight:'No stale leads right now — every active lead has moved in the last 60 days. Nothing to flag.', list:[]};
  var summary=stale.slice(0,25).map(function(l){
    return l.name+' — '+l.status+', '+l.daysStale+' days stale'+(l.potentialAum?' (Potential AUM: '+l.potentialAum+')':'');
  }).join('\n');
  var prompt='You are reviewing a wealth advisor\'s lead pipeline. Below are leads that have sat in Prospect, Contacted, or '
    +'In Progress for 60+ days without converting or being dropped — '+stale.length+' in total. Write a short, direct briefing '
    +'(under 180 words): (1) how many leads and how stale the worst ones are, (2) call out by name the 2-3 highest-value ones '
    +'(by Potential AUM, if any are notably high) worth one more real attempt, (3) end with ONE decisive recommendation labeled '
    +'"Do this:" — either a specific re-engagement action, or a suggestion to clean up leads that are clearly dead. No headers, '
    +'plain text only.\n\n'+summary;
  var insight=callGeminiForInsight(prompt, 450);
  return {insight:insight, list:stale};
}

// ── Never Contacted — active leads with zero activity log entries since
// creation. Worse than "stale": these never got a first touch at all. ──
function generateNeverContactedInsight(){
  var ss=SpreadsheetApp.openById(SHEET_ID);
  var lSh=ss.getSheetByName(LEADS_TAB);
  if(!lSh || lSh.getLastRow()<2) return 'No leads recorded yet.';
  var lr=lSh.getLastRow();
  var vals=lSh.getRange(2,1,lr-1,21).getValues();
  var activeStatuses=['PROSPECT','CONTACTED','IN PROGRESS'];
  var leadsById={};
  vals.forEach(function(r){
    var status=String(r[5]||'').trim().toUpperCase();
    if(activeStatuses.indexOf(status)<0) return;
    var recordId=String(r[17]||'').trim();
    if(!recordId) return;
    leadsById[recordId]={name:String(r[0]||''), recordId:recordId};
  });
  var logSh=ss.getSheetByName(LOG_TAB);
  if(logSh && logSh.getLastRow()>1){
    var logVals=logSh.getRange(2,8,logSh.getLastRow()-1,1).getValues(); // H = Record ID
    logVals.forEach(function(r){
      var rid=String(r[0]||'').trim();
      if(rid && leadsById[rid]) delete leadsById[rid]; // has at least one entry — remove from "never contacted"
    });
  }
  var neverContacted=Object.keys(leadsById).map(function(k){ return leadsById[k]; });
  if(!neverContacted.length) return 'Every active lead has at least one activity logged — nothing to flag here.';
  var summary=neverContacted.slice(0,30).map(function(l){ return l.name; }).join(', ');
  var prompt='You are reviewing a wealth advisor\'s lead pipeline. The following '+neverContacted.length+' active leads have '
    +'ZERO activity log entries since they were added — no call, note, or message has ever been logged for them: '+summary+'. '
    +'Write a short, direct briefing (under 150 words) flagging that these leads may be falling through the cracks entirely, '
    +'and end with ONE decisive recommendation labeled "Do this:" — a concrete next step. No headers, plain text only.';
  return callGeminiForInsight(prompt, 350);
}

// ── Protection Gaps — clients missing health insurance, life insurance, or
// a WILL. A genuine duty-of-care list, not a sales pitch. ──
function generateProtectionGapsInsight(){
  var ss=SpreadsheetApp.openById(SHEET_ID);
  var cSh=ss.getSheetByName(CLIENTS_TAB);
  if(!cSh || cSh.getLastRow()<2) return 'No clients recorded yet.';
  var cr=cSh.getLastRow();
  var vals=cSh.getRange(2,1,cr-1,41).getValues();
  var gaps=[];
  vals.forEach(function(r){
    var name=String(r[0]||''); if(!name) return;
    var health=String(r[21]||'').trim().toUpperCase(); // V = Health insurance
    var life=String(r[22]||'').trim().toUpperCase(); // W = Life insurance
    var will=String(r[23]||'').trim(); // X = WILL
    var missing=[];
    if(health!=='YES') missing.push('Health Insurance');
    if(life!=='YES') missing.push('Life Insurance');
    if(!will || will.toUpperCase()==='NO') missing.push('WILL');
    if(missing.length){
      gaps.push({name:name, aum:String(r[26]||''), missing:missing.join(', ')});
    }
  });
  if(!gaps.length) return 'Every client has health insurance, life insurance, and a WILL on file — no protection gaps to flag.';
  gaps.sort(function(a,b){ return (parseFloat(b.aum)||0)-(parseFloat(a.aum)||0); });
  var summary=gaps.slice(0,25).map(function(g){ return g.name+' — missing: '+g.missing+(g.aum?' (AUM: '+g.aum+')':''); }).join('\n');
  var prompt='You are reviewing a wealth advisor\'s client base for protection gaps. '+gaps.length+' clients are missing at least '
    +'one of: health insurance, life insurance, or a WILL. Write a short, direct briefing (under 180 words), framed as care for '
    +'these families\' security, never as a sales pitch: (1) how many clients and the most common gap, (2) call out by name the '
    +'2-3 highest-AUM clients with a gap, since they have the most to protect, (3) end with ONE decisive recommendation labeled '
    +'"Do this:". No headers, plain text only.\n\n'+summary;
  return callGeminiForInsight(prompt, 450);
}


// ── Client-facing Review Report — deliberately different in every way from
// generateMeetingInsight: warm, plain language, addressed directly to the
// client, and explicitly excludes anything internal (revenue estimates,
// referral counts, engagement/contact-frequency analysis, cross-sell framing).
// PAN is used only for CAS matching, same rule as everywhere else — never
// echoed in the output. ──
function generateClientReviewReport(data){
  var today=new Date();
  var createdDate=data.convertedDate?new Date(data.convertedDate):(data.createdDate?new Date(data.createdDate):null);
  var vintageDays=createdDate?Math.floor((today-createdDate)/86400000):null;
  var vintageYears=vintageDays!==null?Math.floor(vintageDays/365):null;

  var casContent=[];
  var casNote='No CAS document was available.';
  var panIsolationNote='';
  if(data.fileIds && data.fileIds.length){
    var verifiedFolderId=getClientFolderId(data.mobile, data.recordId);
    var casFiles=verifyClientFileOwnership(data.fileIds, verifiedFolderId);
    casFiles.forEach(function(f, idx){
      var b64=Utilities.base64Encode(f.getBlob().getBytes());
      casContent.push({type:'document', source:{type:'base64', media_type:'application/pdf', data:b64}});
      var label=casFiles.length===2?(idx===0?'OLDER (for comparison)':'LATEST'):'LATEST';
      casContent.push({type:'text', text:'The above statement ('+f.getName()+') is labeled: '+label+'.'});
    });
    casNote=casFiles.length===2
      ? 'TWO statements provided — one older (roughly from when the relationship with the client began, or the earliest available), one latest. A detailed, numeric comparison between them is the centerpiece of this letter.'
      : casFiles.length===1
        ? 'ONE statement provided (latest). Describe the current portfolio position only — no comparison is possible.'
        : 'A statement was requested but could not be verified as belonging to this client, so none was included — describe the relationship qualitatively without specific portfolio figures.';
    // This letter goes directly to the client, making correct isolation the
    // highest-stakes of anywhere this pattern is used — a mismatch here
    // means one client's financial data reaching a different client.
    if(casFiles.length){
      panIsolationNote='IMPORTANT — CAS documents are generated per registered EMAIL, not per PAN: a single document can contain '
        +'MULTIPLE different investors if they share one registered email. You MUST identify and use ONLY the holdings belonging '
        +'to PAN '+(data.pan||'[not provided — use the client name to identify the correct investor, and if the document contains '
        +'more than one plausible match, omit specific figures rather than guessing]')+'. This letter is being sent directly to '
        +'this client — including another investor\'s holdings would expose their private financial data. PRIVACY REQUIREMENT: '
        +'the PAN is provided solely for this internal matching purpose — never reproduce it, in full or in part, anywhere in the letter.\n\n';
    }
  }

  var context=panIsolationNote+'CLIENT: '+(data.name||'')+'\n'
    +'Relationship length with the firm: '+(vintageYears!==null?vintageYears+' year(s)':'not available')+'\n'
    +'Products currently held: '+(data.products||'not recorded')+'\n'
    +'Health insurance in place: '+(data.healthInsurance==='YES'?'Yes':data.healthInsurance==='NO'?'No':'Not confirmed')+'\n'
    +'Life insurance in place: '+(data.lifeInsurance==='YES'?'Yes':data.lifeInsurance==='NO'?'No':'Not confirmed')+'\n'
    +'WILL in place: '+(data.willDone?'Yes':'Not confirmed')+'\n'
    +(data.ourArnCode?'OUR FIRM\u2019S ARN/ADVISOR CODE: '+data.ourArnCode+'\n':'')
    +'CAS NOTE: '+casNote;

  var arnInstruction = data.ourArnCode ?
    ('FUND ATTRIBUTION — IMPORTANT: each fund/folio entry on a CAS statement typically shows a broker/ARN code next to it. Compare this '
    +'against our firm\u2019s own ARN code given above ('+data.ourArnCode+'). Only count a holding as "managed by us" if its ARN matches ours '
    +'exactly. Any holding with a different ARN, no ARN, or marked "Direct" was NOT brought in or is not being advised on by this firm — '
    +'do not credit its growth to the relationship with this advisor. If the client holds funds both with us and elsewhere, the growth '
    +'story in section 2 should focus specifically on the funds under our ARN, and you may briefly, matter-of-factly acknowledge the rest '
    +'exists without taking credit for it (e.g. "alongside the investments you hold with us..."). If every fund on the statement carries '
    +'our ARN, there\u2019s no need to raise this distinction at all — write the letter normally.\n\n')
    : '';

  var promptText='Write a warm, personal investment review letter addressed DIRECTLY to a wealth management client, from their advisor '
    +'Arjun K T M at Wealth Matrix Finserv LLP. This will be read BY THE CLIENT THEMSELVES — not an internal document. Tone: warm, '
    +'plain-language, genuinely personal, like a thoughtful note from someone who knows them — never salesy, never alarming, never '
    +'using financial jargon without explaining it simply.\n\n'
    +arnInstruction
    +(casFiles && casFiles.length===2 ?
      'BEFORE WRITING — extract these figures precisely from the two attached statements (do this quietly; the extraction itself should '
      +'never appear in the output, only the finished letter):\n'
      +'  a) Total portfolio value (sum of all mutual fund holdings) on the OLDER statement, and on the LATEST statement.\n'
      +'  b) The number of DISTINCT fund/scheme holdings on the OLDER statement, and on the LATEST statement.\n'
      +'  c) Absolute growth in rupees (latest minus older) and the percentage growth.\n'
      +'  d) Any fund names present in the LATEST statement that were NOT in the OLDER one (funds added since then).\n'
      +'  e) Any fund names present in the OLDER statement that are NOT in the LATEST one (funds fully exited or redeemed).\n'
      +'  f) If SIP status is visible on either statement, note whether SIPs are active/continuing.\n'
      +'Use these six extracted facts as the factual backbone of section 2 below — cite the actual rupee figures and fund counts, not '
      +'vague language like "your portfolio has grown nicely." Be numerically specific and correct.\n\n'
      +'HONESTY GUARDRAIL: growth in total value can come from market performance AND from new money added (fresh investments/SIP '
      +'contributions) — a CAS alone cannot cleanly separate the two. Do not claim the growth is purely "market gains" or "returns" unless '
      +'that is the only reasonable reading. Safer, honest framing: describe the growth in value and fund count as the combined result of '
      +'their continued investing discipline and market performance together, not attribute it to one cause alone.\n\n'
      : '')
    +'STRICT EXCLUSIONS — do not include any of the following, even implicitly: advisor revenue or commission figures, referral counts, '
    +'how often the client has been contacted, "cross-sell opportunity" framing, or any language that sounds like internal sales notes. '
    +'This is a message TO the client, not ABOUT them.\n\n'
    +'Structure, in this order:\n'
    +'1) A warm, genuine opening acknowledging their relationship with the firm'+(vintageYears!==null?' (they have been with the firm '+vintageYears+' year(s))':'')+'.\n'
    +'2) Their portfolio, explained simply: '+casNote+' Use plain language throughout — "your investments are now worth approximately ₹X, '
    +'up from about ₹Y" rather than technical terms. If two statements are available, this section is the heart of the letter: state the '
    +'actual growth in value (₹ and %), how the number of funds/schemes has changed, and mention by name any new fund(s) added or fully '
    +'exited, all using the figures extracted above. Make the progress feel tangible and earned, not just a number.\n'
    +'3) What\'s going well — genuinely acknowledge good financial habits reflected in their portfolio (consistency, diversification, '
    +'staying invested, whatever is actually supported by the data) — this must feel sincere, not generic flattery.\n'
    +'4) A gentle, caring note on protection — only if health insurance, life insurance, or a WILL is marked "No" or "Not confirmed" above, '
    +'mention it once, briefly, framed as care for their family\'s security, never as a sales pitch or a warning.\n'
    +'5) A warm closing line inviting them to reach out anytime, followed by:\nWarm regards,\nArjun K T M\nWealth Matrix Finserv LLP\n\n'
    +'Keep it under 400 words. No headers or section titles in the output, and no visible extraction/calculation steps — write it as a '
    +'single flowing, natural letter a person would actually enjoy receiving. Start directly with "Dear '+(data.name||'')+',"\n\n'
    +context;

  var content=casContent.slice();
  content.push({type:'text', text:promptText});

  return callGemini({ mixedContent: content, maxOutputTokens: 1200 });
}

function generateAndSendInsights(){
  var ss=SpreadsheetApp.openById(SHEET_ID);
  var iSh=ss.getSheetByName(INSIGHTS_TAB)||ss.insertSheet(INSIGHTS_TAB);
  if(iSh.getLastRow()===0){
    var iHdr=[['Date','Insight']];
    var iHr=iSh.getRange(1,1,1,2);
    iHr.setValues(iHdr);
    iHr.setFontWeight('bold').setBackground('#C0392B').setFontColor('#FFFFFF');
    iSh.setFrozenRows(1);
  }
  var stamp=Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
  try{
    var summary=buildInsightsSummaryText();
    var insight=callGeminiAPI(summary);
    iSh.appendRow([stamp, insight]);
    try{
      MailApp.sendEmail('arjun.wealthmatrix@gmail.com', 'Weekly CRM Insights — Wealth Matrix', insight);
    }catch(mailErr){
      console.warn('Insights email failed to send:', mailErr.message); // don't fail the whole run over email specifically
    }
    return {success:true, insight:insight};
  }catch(e){
    iSh.appendRow([stamp, 'ERROR: '+e.message]);
    return {success:false, error:e.message};
  }
}

// ── Trigger management — enable/disable the weekly schedule from the app ──
function setupInsightsTrigger(){
  removeInsightsTrigger(); // avoid duplicates on re-enable
  ScriptApp.newTrigger('generateAndSendInsights')
    .timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(8).create();
  return {success:true};
}
function removeInsightsTrigger(){
  var triggers=ScriptApp.getProjectTriggers();
  var removed=0;
  triggers.forEach(function(t){
    if(t.getHandlerFunction()==='generateAndSendInsights'){ ScriptApp.deleteTrigger(t); removed++; }
  });
  return {success:true, removed:removed};
}
function getInsightsTriggerStatus(){
  var triggers=ScriptApp.getProjectTriggers();
  var active=triggers.some(function(t){return t.getHandlerFunction()==='generateAndSendInsights';});
  return {enabled:active};
}

function doGet(e){try{var p=e.parameter||{};if(p.ping||p.action==='version')return out({version:VERSION,ok:true,deployed:true});if(p.action==='resolveUploadToken'&&p.token)return resolveUploadToken(p.token);if(p.action==='pullLog'&&(p.mobile||p.recordId))return pullLog(p.mobile,p.recordId);if(p.payload)return writeData(JSON.parse(decodeURIComponent(p.payload)));return readTab(p.tab||LEADS_TAB);}catch(err){return out({error:err.message});}}
function doPost(e){try{var d;if(e.postData&&e.postData.contents)d=JSON.parse(e.postData.contents);else if(e.parameter&&e.parameter.payload)d=JSON.parse(decodeURIComponent(e.parameter.payload));return writeData(d);}catch(err){return out({error:err.message});}}
function readTab(t){var ss=SpreadsheetApp.openById(SHEET_ID);var sh=ss.getSheetByName(t);if(!sh)return out({rows:[],error:'Tab not found:'+t});var lr=sh.getLastRow(),lc=sh.getLastColumn();if(lr===0||lc===0)return out({rows:[]});var raw=sh.getRange(1,1,lr,lc).getValues();var cleaned=raw.map(function(row){return row.map(function(c){if(c===null||c===undefined||c==='')return'';if(c instanceof Date){var y=c.getFullYear();var m=String(c.getMonth()+1).padStart(2,'0');var d=String(c.getDate()).padStart(2,'0');return y+'-'+m+'-'+d;}if(typeof c==='number')return String(c);return String(c).trim();});});return out({rows:cleaned});}

// ── Resolve a client upload-link token (read-only, no lock needed) ──
// Returns only the minimal fields the upload page needs — never the whole client list.
function resolveUploadToken(token){
  try{
    var ss=SpreadsheetApp.openById(SHEET_ID);
    var tok=String(token||'').trim();
    if(!tok) return out({found:false});
    // Check Clients first (existing behavior)
    var sh=ss.getSheetByName(CLIENTS_TAB);
    if(sh && sh.getLastRow()>=2){
      var lr=sh.getLastRow();
      var lc=Math.max(sh.getLastColumn(),41);
      var vals=sh.getRange(2,1,lr-1,lc).getValues();
      for(var i=0;i<vals.length;i++){
        var rowTok=String(vals[i][32]||'').trim(); // column AG (index 32) = Upload Token
        if(rowTok && rowTok===tok){
          var normD=function(c){
            if(c===null||c===undefined||c==='') return '';
            if(c instanceof Date){ var y=c.getFullYear(),m=String(c.getMonth()+1).padStart(2,'0'),d=String(c.getDate()).padStart(2,'0'); return y+'-'+m+'-'+d; }
            return String(c).trim();
          };
          return out({found:true, name:String(vals[i][0]||''), mobile:String(vals[i][1]||''), pan:String(vals[i][11]||''), recordType:'client', recordId:String(vals[i][36]||''), folderId:String(vals[i][34]||''),
            appInstalled:String(vals[i][18]||''), healthInsurance:String(vals[i][21]||''), lifeInsurance:String(vals[i][22]||''),
            dueDateHealth:normD(vals[i][24]), dueDateLife:normD(vals[i][25]), products:String(vals[i][29]||''),
            healthInsuranceCoverage:String(vals[i][39]||''), lifeInsuranceCoverage:String(vals[i][40]||'')});
        }
      }
    }
    // Not a client yet — check Leads. Column R (17) = Record ID, column T (19) = Folder ID,
    // column U (20) = Upload Token — three DISTINCT columns, never overloaded together.
    var lsh=ss.getSheetByName(LEADS_TAB);
    if(lsh && lsh.getLastRow()>=2){
      var llr=lsh.getLastRow();
      var llc=Math.max(lsh.getLastColumn(),21);
      var lvals=lsh.getRange(2,1,llr-1,llc).getValues();
      for(var j2=0;j2<lvals.length;j2++){
        var lRowTok2=String(lvals[j2][20]||'').trim(); // column U (index 20) = Upload Token
        if(lRowTok2 && lRowTok2===tok){
          return out({found:true, name:String(lvals[j2][0]||''), mobile:String(lvals[j2][1]||''), pan:'', recordType:'lead', recordId:String(lvals[j2][17]||''), folderId:String(lvals[j2][19]||'')});
        }
      }
    }
    return out({found:false});
  }catch(e){
    return out({found:false, error:e.message});
  }
}
// Returns ONLY the activity rows belonging to one record. Matching mirrors
// the frontend's rule exactly, so moving this filter server-side loses no
// accuracy: when BOTH the row and the target have a Record ID, match on
// Record ID alone (exact — correctly separates two leads sharing a mobile);
// fall back to mobile only for legacy rows that have no Record ID stored.
// A row that HAS a Record ID never matches a different target by mobile.
function pullLog(mobile, recordId){
  var ss=SpreadsheetApp.openById(SHEET_ID);
  var sh=ss.getSheetByName(LOG_TAB);
  if(!sh||sh.getLastRow()<=1) return out({rows:[]});
  var lr=sh.getLastRow(), lc=sh.getLastColumn();
  var raw=sh.getRange(2,1,lr-1,lc).getValues();
  var targetMobile=String(mobile||'').trim();
  var targetRecordId=String(recordId||'').trim();
  var filtered=raw.filter(function(r){
    var rowRecordId=String(r[7]||'').trim(); // column H = Record ID
    var rowMobile=String(r[2]||'').trim();   // column C = Mobile
    if(rowRecordId && targetRecordId) return rowRecordId===targetRecordId;
    if(!rowRecordId) return rowMobile===targetMobile;
    return false;
  });
  return out({rows:filtered});
}
// Actions that only READ the sheet and call an AI API — they never modify
// anything, so they must NOT sit behind the global write lock. Previously
// they did, which meant each chat message could wait up to 25s just to
// acquire the lock, then HOLD it for the entire multi-second AI call —
// blocking every other write in the app, and frequently blowing past the
// frontend's timeout before a reply ever came back. This is the single
// biggest cause of "no reply from the AI".
var READ_ONLY_AI_ACTIONS = {
  practiceChat:true,
  generateContemporaryInsight:true,
  generateTopBusinessInsights:true,
  draftLeadMessage:true,
  draftClientMessage:true,
  generateLeadProfileInsight:true,
  generateClientProfileInsight:true,
  generateMeetingInsight:true,
  generateEngagementInsight:true,
  generateProductivityInsight:true,
  generateStaleLeadsInsight:true,
  generateProtectionGapsInsight:true,
  generateNeverContactedInsight:true,
  draftCasEmailInsight:true
};
function writeData(data){
  if(data && data.action && READ_ONLY_AI_ACTIONS[data.action]){
    // Straight through — no lock acquired, no lock held during the API call.
    try{ return _writeDataInner(data); }
    catch(e){ return out({success:false, error:e.message}); }
  }
  var _lock=LockService.getScriptLock();
  try{_lock.waitLock(25000);}catch(_lkErr){return out({error:'Server busy, please retry'});}
  try{return _writeDataInner(data);}
  finally{try{SpreadsheetApp.flush();}catch(_fErr){}try{_lock.releaseLock();}catch(_rlErr){}}
}
function _writeDataInner(data){var ss=SpreadsheetApp.openById(SHEET_ID);
// EARLY DISPATCH — the AI chat actions are the highest-traffic, most
// latency-sensitive calls in the app, and previously sat ~1,450 lines deep,
// behind roughly a hundred string comparisons. Handling them here skips all
// of that. The original handlers further down are left untouched and simply
// become unreachable for these actions, so removing this block would restore
// the previous behaviour exactly rather than break anything.
if(data.action==='practiceChat'){
  try{ return out({success:true, reply:practiceChat(data)}); }
  catch(e){ return out({success:false, error:e.message}); }
}
if(data.action==='generateContemporaryInsight'){
  try{ return out({success:true, reply:generateContemporaryInsight(data)}); }
  catch(e){ return out({success:false, error:e.message}); }
}
if(data.action==='appendLog'){var logSh=ss.getSheetByName(LOG_TAB)||ss.insertSheet(LOG_TAB);if(logSh.getLastRow()===0){var hdr=[['Timestamp','Name','Mobile','Type','Activity','Notes','Raw ID','Record ID','Logged By']];var hr=logSh.getRange(1,1,1,9);hr.setValues(hdr);hr.setFontWeight('bold').setBackground('#C0392B').setFontColor('#FFFFFF');logSh.setFrozenRows(1);}else{ensureColumnHeader(logSh,8,'Record ID');ensureColumnHeader(logSh,9,'Logged By');}var logRow=data.row||[];var newRid=String(logRow[6]||'').trim();var logDup=false;if(newRid&&logSh.getLastRow()>1){var ridVals=logSh.getRange(2,7,logSh.getLastRow()-1,1).getValues();for(var _ri=0;_ri<ridVals.length;_ri++){if(String(ridVals[_ri][0]||'').trim()===newRid){logDup=true;break;}}}if(!logDup){logSh.appendRow(logRow);var lr3=logSh.getLastRow();logSh.getRange(lr3,1).setNumberFormat('@');logSh.getRange(lr3,7).setNumberFormat('@');logSh.getRange(lr3,8).setNumberFormat('@');logSh.getRange(lr3,9).setNumberFormat('@');}return out({success:true,skipped:logDup});}
if(data.action==='deleteLogById'){var lsh=ss.getSheetByName(LOG_TAB);if(lsh&&lsh.getLastRow()>1){var rid=String(data.rowId||'').trim();var dmob=String(data.mobile||'').trim();var dts=String(data.ts||'').trim();var dact=String(data.activity||'').trim();var n=lsh.getLastRow();var rng=lsh.getRange(2,1,n-1,7).getValues();for(var rj=0;rj<rng.length;rj++){var row=rng[rj];var gId=String(row[6]||'').trim();var rTs=String(row[0]||'').trim();var rMob=String(row[2]||'').trim();var rAct=String(row[4]||'').trim();var matchById=(rid!==''&&gId===rid);var matchByTs=(rid===''||gId==='')&&dts!==''&&rMob===dmob&&rTs===dts&&rAct===dact;if(matchById||matchByTs){lsh.deleteRow(rj+2);break;}}}return out({success:true,deleted:true});}if(data.action==='deleteLog'){var lsh2=ss.getSheetByName(LOG_TAB);if(lsh2&&lsh2.getLastRow()>1){var mob3=String(data.mobile||'').trim();var n2=lsh2.getLastRow();for(var rk=n2;rk>=2;rk--){if(String(lsh2.getRange(rk,3).getValue()||'').trim()===mob3){lsh2.deleteRow(rk);}}}return out({success:true});}

  // ── Trim activity log to the most recent 10 entries per lead/client — FIFO,
  // oldest deleted first. Runs silently in the background after the app loads,
  // no user confirmation. Groups by Record ID (column H) where present; falls
  // back to Mobile+Type together for legacy entries logged before Record ID
  // was added to this sheet. (Restored here — this was previously built but
  // found missing from the live file during the auth feature build, likely
  // lost in an earlier edit; re-added rather than left silently absent.) ──
  if(data.action==='trimActivityLogs'){
    try{
      var tlSh=ss.getSheetByName(LOG_TAB);
      if(!tlSh || tlSh.getLastRow()<2) return out({success:true, deleted:0});
      var tlLr=tlSh.getLastRow();
      var tlVals=tlSh.getRange(2,1,tlLr-1,8).getValues();
      var tlGroups={};
      var tlLimit=parseInt(data.limit)||10;
      for(var tli=0;tli<tlVals.length;tli++){
        var tlRecId=String(tlVals[tli][7]||'').trim();
        var tlMob=String(tlVals[tli][2]||'').trim();
        var tlType=String(tlVals[tli][3]||'').trim().toUpperCase();
        var tlKey=tlRecId ? ('R:'+tlRecId) : ('M:'+tlMob+':'+tlType);
        if(!tlKey || tlKey==='M::') continue;
        if(!tlGroups[tlKey]) tlGroups[tlKey]=[];
        tlGroups[tlKey].push({rowIndex:tli, ts:String(tlVals[tli][0]||'')});
      }
      var tlRowsToDelete=[];
      Object.keys(tlGroups).forEach(function(k){
        var entries=tlGroups[k];
        if(entries.length<=tlLimit) return;
        entries.sort(function(a,b){ return a.ts.localeCompare(b.ts); });
        var excess=entries.length-tlLimit;
        for(var ei=0;ei<excess;ei++){ tlRowsToDelete.push(entries[ei].rowIndex); }
      });
      tlRowsToDelete.sort(function(a,b){return b-a;});
      tlRowsToDelete.forEach(function(idx){ tlSh.deleteRow(idx+2); });
      return out({success:true, deleted:tlRowsToDelete.length});
    }catch(e){
      return out({success:false, error:e.message});
    }
  }

  // ═══════════════════════════════════════════════
  // AUTHENTICATION & USER MANAGEMENT — Admin sheet columns: A=Name, B=Username,
  // C=User Type (Maker/Checker/Authorizer/Admin), D=User ID, E=Password (SHA-256
  // hash, never plain text), F=Timestamp (last password-changed date; blank
  // means still using the admin-set initial password, forcing a change on login).
  // ═══════════════════════════════════════════════

  // Server-side role check — never trust a client-supplied "I'm an admin"
  // claim for a sensitive action; always re-verify against the sheet.
  function verifyIsAdmin(userId){
    var adSh=ss.getSheetByName(ADMIN_TAB);
    if(!adSh || adSh.getLastRow()<2) return false;
    var vals=adSh.getRange(2,1,adSh.getLastRow()-1,4).getValues(); // A-D
    for(var i=0;i<vals.length;i++){
      if(String(vals[i][3]||'').trim()===String(userId||'').trim()){
        return String(vals[i][2]||'').trim().toUpperCase()==='ADMIN';
      }
    }
    return false;
  }

  // Used to gate the task review actions (approve/send back) — Checker,
  // Authorizer, or Admin may review; Maker may not, including their own
  // pending-review tasks. Mirrors verifyIsAdmin exactly, never trusts a
  // role claim sent from the frontend.
  function verifyCanReview(userId){
    var crSh=ss.getSheetByName(ADMIN_TAB);
    if(!crSh || crSh.getLastRow()<2) return false;
    var crVals=crSh.getRange(2,1,crSh.getLastRow()-1,4).getValues(); // A-D
    for(var cri=0;cri<crVals.length;cri++){
      if(String(crVals[cri][3]||'').trim()===String(userId||'').trim()){
        var crRole=String(crVals[cri][2]||'').trim().toUpperCase();
        return crRole==='CHECKER' || crRole==='AUTHORIZER' || crRole==='ADMIN';
      }
    }
    return false;
  }

  if(data.action==='loginUser'){
    try{
      var luSh=ss.getSheetByName(ADMIN_TAB);
      if(!luSh || luSh.getLastRow()<2) return out({success:false, error:'No users set up yet'});
      var luVals=luSh.getRange(2,1,luSh.getLastRow()-1,6).getValues(); // A-F
      var luUser=String(data.username||'').trim().toLowerCase();
      var luHash=String(data.passwordHash||'').trim();
      for(var i=0;i<luVals.length;i++){
        if(String(luVals[i][1]||'').trim().toLowerCase()===luUser){
          if(String(luVals[i][4]||'').trim()!==luHash) return out({success:false, error:'Incorrect username or password'});
          return out({
            success:true, name:String(luVals[i][0]||''), username:String(luVals[i][1]||''),
            userType:String(luVals[i][2]||'').trim().toUpperCase(), userId:String(luVals[i][3]||''),
            passwordNeverChanged: !String(luVals[i][5]||'').trim()
          });
        }
      }
      return out({success:false, error:'Incorrect username or password'});
    }catch(e){
      return out({success:false, error:e.message});
    }
  }

  // User changes their OWN password — used both for the forced first-login
  // change and any later voluntary change from Settings.
  if(data.action==='changeOwnPassword'){
    try{
      var cpSh=ss.getSheetByName(ADMIN_TAB);
      if(!cpSh || cpSh.getLastRow()<2) return out({success:false, error:'No users set up yet'});
      var cpLr=cpSh.getLastRow();
      var cpIds=cpSh.getRange(2,4,cpLr-1,1).getValues(); // D = User ID
      for(var i=0;i<cpIds.length;i++){
        if(String(cpIds[i][0]||'').trim()===String(data.userId||'').trim()){
          var cpRow=i+2;
          // If an old-password hash was supplied, verify it before changing —
          // the forced first-login change skips this (they just authenticated).
          if(data.oldPasswordHash){
            var curHash=String(cpSh.getRange(cpRow,5).getValue()||'').trim();
            if(curHash!==String(data.oldPasswordHash).trim()) return out({success:false, error:'Current password is incorrect'});
          }
          cpSh.getRange(cpRow,5).setValue(String(data.newPasswordHash||'').trim());
          cpSh.getRange(cpRow,6).setValue(Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss'));
          return out({success:true});
        }
      }
      return out({success:false, error:'User not found'});
    }catch(e){
      return out({success:false, error:e.message});
    }
  }

  // Admin resets ANY user's password — does not require or expose the old
  // password. Server-verifies the requester is genuinely an Admin first.
  if(data.action==='resetUserPassword'){
    try{
      if(!verifyIsAdmin(data.requestingUserId)) return out({success:false, error:'Not authorized'});
      var rpSh=ss.getSheetByName(ADMIN_TAB);
      if(!rpSh || rpSh.getLastRow()<2) return out({success:false, error:'No users set up yet'});
      var rpLr=rpSh.getLastRow();
      var rpIds=rpSh.getRange(2,4,rpLr-1,1).getValues();
      for(var i=0;i<rpIds.length;i++){
        if(String(rpIds[i][0]||'').trim()===String(data.targetUserId||'').trim()){
          var rpRow=i+2;
          rpSh.getRange(rpRow,5).setValue(String(data.newPasswordHash||'').trim());
          rpSh.getRange(rpRow,6).setValue(''); // blank again — forces a change on the user's next login
          return out({success:true});
        }
      }
      return out({success:false, error:'User not found'});
    }catch(e){
      return out({success:false, error:e.message});
    }
  }

  // Admin creates a new user. Server-verifies the requester is genuinely an Admin.
  if(data.action==='createUser'){
    try{
      if(!verifyIsAdmin(data.requestingUserId)) return out({success:false, error:'Not authorized'});
      var cuSh=ss.getSheetByName(ADMIN_TAB)||ss.insertSheet(ADMIN_TAB);
      if(cuSh.getLastRow()===0){
        var cuHdr=[['Name','Username','User Type','User ID','Password','Timestamp']];
        var cuHr=cuSh.getRange(1,1,1,6);
        cuHr.setValues(cuHdr);
        cuHr.setFontWeight('bold').setBackground('#C0392B').setFontColor('#FFFFFF');
        cuSh.setFrozenRows(1);
      }
      var cuUsername=String(data.username||'').trim();
      if(!cuUsername) return out({success:false, error:'Username required'});
      if(cuSh.getLastRow()>1){
        var cuExisting=cuSh.getRange(2,2,cuSh.getLastRow()-1,1).getValues();
        for(var i=0;i<cuExisting.length;i++){
          if(String(cuExisting[i][0]||'').trim().toLowerCase()===cuUsername.toLowerCase()) return out({success:false, error:'That username is already taken'});
        }
      }
      var newUserId=nextSequentialId(cuSh,4,'U');
      cuSh.appendRow([data.name||'', cuUsername, String(data.userType||'').trim().toUpperCase(), newUserId, String(data.initialPasswordHash||'').trim(), '']);
      var cuLr=cuSh.getLastRow();
      cuSh.getRange(cuLr,1,1,6).setNumberFormat('@');
      return out({success:true, userId:newUserId});
    }catch(e){
      return out({success:false, error:e.message});
    }
  }

  // Admin lists all users — Name/Username/Type/UserID/last-changed date only,
  // password hash is never included in this response.
  if(data.action==='listUsers'){
    try{
      if(!verifyIsAdmin(data.requestingUserId)) return out({success:false, error:'Not authorized'});
      var luListSh=ss.getSheetByName(ADMIN_TAB);
      if(!luListSh || luListSh.getLastRow()<2) return out({success:true, users:[]});
      var luListVals=luListSh.getRange(2,1,luListSh.getLastRow()-1,6).getValues();
      var users=luListVals.map(function(r){
        return { name:String(r[0]||''), username:String(r[1]||''), userType:String(r[2]||'').trim().toUpperCase(),
          userId:String(r[3]||''), passwordChangedDate:String(r[5]||'') };
      }).filter(function(u){ return u.userId; });
      return out({success:true, users:users});
    }catch(e){
      return out({success:false, error:e.message});
    }
  }


  // ── Trim activity log history to the most recent 10 entries per lead/client.
  // Runs silently in the background (no confirmation) once per app load. Groups
  // by Record ID (column H) where available — falling back to type+mobile for
  // legacy rows logged before that column existed, same rule getActivity() uses
  // on the frontend. Rewrites the sheet ONCE with only the rows to keep, rather
  // than deleting rows one at a time — safer (no risk of index-shift bugs from
  // repeated deleteRow calls) and far fewer operations. ──
  if(data.action==='trimActivityLogHistory'){
    try{
      var trimSh=ss.getSheetByName(LOG_TAB);
      if(!trimSh || trimSh.getLastRow()<2) return out({success:true, removed:0});
      var trimLr=trimSh.getLastRow();
      var trimVals=trimSh.getRange(2,1,trimLr-1,8).getValues(); // A-H: Ts,Name,Mobile,Type,Activity,Notes,RowId,RecordId
      var groups={};
      trimVals.forEach(function(row,idx){
        var recId=String(row[7]||'').trim();
        var type=String(row[3]||'').trim().toUpperCase();
        var mobile=String(row[2]||'').trim();
        var key=recId ? ('R:'+recId) : ('M:'+type+':'+mobile);
        if(!groups[key]) groups[key]=[];
        groups[key].push({row:row, idx:idx});
      });
      var keepIdx={};
      var removedCount=0;
      Object.keys(groups).forEach(function(key){
        var entries=groups[key];
        if(entries.length<=10){
          entries.forEach(function(e){ keepIdx[e.idx]=true; });
          return;
        }
        // Sort by timestamp ascending (oldest first), keep the most recent 10
        entries.sort(function(a,b){ return String(a.row[0]||'').localeCompare(String(b.row[0]||'')); });
        var toRemove=entries.length-10;
        for(var ei=0;ei<entries.length;ei++){
          if(ei<toRemove){ removedCount++; }
          else { keepIdx[entries[ei].idx]=true; }
        }
      });
      if(removedCount===0) return out({success:true, removed:0});
      var keptRows=trimVals.filter(function(row,idx){ return keepIdx[idx]; });
      trimSh.getRange(2,1,trimLr-1,8).clearContent();
      if(keptRows.length>0){
        var trimRange=trimSh.getRange(2,1,keptRows.length,8);
        trimRange.setValues(keptRows);
        trimRange.setNumberFormat('@');
      }
      return out({success:true, removed:removedCount});
    }catch(e){
      return out({success:false, error:e.message});
    }
  }
if(data.action==='updateAnalytics'){var aS=ss.getSheetByName('ANALYTICS')||ss.insertSheet('ANALYTICS');if(aS.getLastRow()===0){aS.getRange(1,1,1,7).setValues([['Date','Lead Entries','Prospects','Contacted','In Progress','Converted','Dropped']]).setFontWeight('bold').setBackground('#C0392B').setFontColor('#FFFFFF');aS.setFrozenRows(1);}var ad=data.date,ac=parseInt(data.col)||2,adel=parseInt(data.delta)||1,alr=aS.getLastRow(),af=false;if(alr>1){var adts=aS.getRange(2,1,alr-1,1).getValues();for(var i=0;i<adts.length;i++){if(String(adts[i][0]).trim()===ad){var acel=aS.getRange(i+2,ac);acel.setValue((parseInt(acel.getValue())||0)+adel); /* negative values allowed */ af=true; break; }}}if(!af&&adel>0){var anr=[ad,0,0,0,0,0,0];anr[ac-1]=adel;aS.appendRow(anr);}return out({success:true});}
if(data.action==='appendLeads'||data.action==='appendClients'){var atab=data.action==='appendLeads'?LEADS_TAB:CLIENTS_TAB;var ash=ss.getSheetByName(atab);if(ash&&data.rows&&data.rows.length){data.rows.forEach(function(rw){ash.appendRow(rw);});var alr=ash.getLastRow();ash.getRange(2,1,alr-1,data.rows[0].length).setNumberFormat('@');}return out({success:true});}
// ── appendLead: UPSERT keyed on NAME+MOBILE together (not mobile alone) — two
// different leads (e.g. family members) are allowed to share one mobile number;
// only a genuine resubmit of the SAME name+mobile pair is treated as a retry. ──
if(data.action==='appendLead'){
  var aSh=ss.getSheetByName(LEADS_TAB)||ss.insertSheet(LEADS_TAB);
  if(aSh.getLastRow()===0){
    var lHdr=[['Name','Mobile','Email','Age','Country','Status','Source','Referred By','Potential SIP','Next Follow-up','Investment Products','Mode of Contact','Remarks','Modified Date','Created Date','Geo Tag','Potential AUM','Record ID','Family Name','Folder ID','Upload Token','User ID']];
    var lHr=aSh.getRange(1,1,1,22);lHr.setValues(lHdr);lHr.setFontWeight('bold').setBackground('#C0392B').setFontColor('#FFFFFF');aSh.setFrozenRows(1);
  }
  ensureColumnHeader(aSh,20,'Folder ID');
  ensureColumnHeader(aSh,21,'Upload Token');
  ensureColumnHeader(aSh,22,'User ID');
  var aRow=data.row||[];
  var aMob=String(aRow[1]||'').trim();
  var aName=String(aRow[0]||'').trim().toUpperCase();
  var aWrote=false;
  if(aMob&&aName&&aSh.getLastRow()>1){
    var aEx=aSh.getRange(2,1,aSh.getLastRow()-1,2).getValues(); // A=name, B=mobile
    for(var _ai=0;_ai<aEx.length;_ai++){
      if(String(aEx[_ai][1]||'').trim()===aMob && String(aEx[_ai][0]||'').trim().toUpperCase()===aName){
        var aRng=aSh.getRange(_ai+2,1,1,aRow.length);aRng.setValues([aRow]);aRng.setNumberFormat('@');aWrote=true;break;
      }
    }
  }
  if(!aWrote){
    if(!String(aRow[17]||'').trim())aRow[17]=nextSequentialId(aSh,18,'L'); // column R = Record ID (L1, L2...)
    aSh.appendRow(aRow);
    var aLr=aSh.getLastRow();
    aSh.getRange(aLr,1,1,aRow.length).setNumberFormat('@');
  }
  return out({success:true,action:'appendLead',upserted:aWrote,recordId:String(aRow[17]||'')});
}
// ── updateLead: matched by Record ID (column R) — the ONLY reliable key, since
// mobile numbers can legitimately repeat across different leads. Falls back to
// name+mobile only for legacy calls that don't yet send a recordId. ──
if(data.action==='updateLead'){
  var uSh=ss.getSheetByName(LEADS_TAB);
  var uFound=-1;
  if(uSh&&uSh.getLastRow()>1){
    var uRow=data.row||[];
    var uN=uSh.getLastRow();
    var uRecId=String(data.recordId||'').trim();
    if(uRecId){
      var uIds=uSh.getRange(2,18,uN-1,1).getValues(); // column R
      for(var ui=0;ui<uIds.length;ui++){ if(String(uIds[ui][0]||'').trim()===uRecId){ uFound=ui; break; } }
    }
    if(uFound<0){
      // Legacy fallback — match by name+mobile together (never mobile alone)
      var uMob=String(data.mobile||'').trim();
      var uName=String(uRow[0]||'').trim().toUpperCase();
      var uVals=uSh.getRange(2,1,uN-1,2).getValues();
      for(var ui2=0;ui2<uVals.length;ui2++){
        if(String(uVals[ui2][1]||'').trim()===uMob && String(uVals[ui2][0]||'').trim().toUpperCase()===uName){ uFound=ui2; break; }
      }
    }
    if(uFound>=0){
      // Preserve ownership (column V, index 21) if the incoming row doesn't
      // specify one — an edit should never silently reassign or wipe out who
      // originally owns this lead, regardless of what the save path sends.
      if(!String(uRow[21]||'').trim()){
        var uExistingOwner=uSh.getRange(uFound+2,22).getValue();
        uRow[21]=uExistingOwner||'';
      }
      var uRng=uSh.getRange(uFound+2,1,1,uRow.length);
      uRng.setValues([uRow]);
      uRng.setNumberFormat('@');
    }
  }
  // Honestly report whether a row was actually found and written — previously
  // this always returned success, silently doing nothing when neither Record ID
  // nor Name+Mobile matched, while the app showed a success toast. Now the
  // frontend's existing fallback (a full rewrite via pushLeads) correctly triggers.
  if(uFound>=0){
    return out({success:true,action:'updateLead'});
  }
  return out({success:false,action:'updateLead',error:'No matching lead row found (checked Record ID and Name+Mobile) — Record ID: '+String(data.recordId||'(none)')+', Mobile: '+String(data.mobile||'(none)')});
}
// ── deleteLead: matched by Record ID (column R) primarily; name+mobile fallback ──
if(data.action==='deleteLead'){
  var dSh=ss.getSheetByName(LEADS_TAB);
  var dFolderIds=[]; // collect any tagged folder(s) before the row(s) disappear
  if(dSh&&dSh.getLastRow()>1){
    var dRecId=String(data.recordId||'').trim();
    var dN=dSh.getLastRow();
    if(dRecId){
      var dIds=dSh.getRange(2,18,dN-1,1).getValues();
      for(var di=dIds.length-1;di>=0;di--){
        if(String(dIds[di][0]||'').trim()===dRecId){
          var dFid=String(dSh.getRange(di+2,20).getValue()||'').trim(); // column T = Folder ID
          if(dFid) dFolderIds.push(dFid);
          dSh.deleteRow(di+2);
        }
      }
    } else {
      var dMob=String(data.mobile||'').trim();
      var dName=String(data.name||'').trim().toUpperCase();
      var dVals=dSh.getRange(2,1,dN-1,2).getValues();
      for(var di2=dVals.length-1;di2>=0;di2--){
        if(String(dVals[di2][1]||'').trim()===dMob && (!dName || String(dVals[di2][0]||'').trim().toUpperCase()===dName)){
          var dFid2=String(dSh.getRange(di2+2,20).getValue()||'').trim();
          if(dFid2) dFolderIds.push(dFid2);
          dSh.deleteRow(di2+2);
        }
      }
    }
  }
  // Trash the Drive folder too — soft delete (recoverable from Drive trash for
  // ~30 days, same as individual document deletes elsewhere in this app), not
  // a permanent wipe. Non-fatal if the folder is already gone.
  dFolderIds.forEach(function(fid){
    try{ DriveApp.getFolderById(fid).setTrashed(true); }catch(fe){ /* already gone or inaccessible — skip */ }
  });
  return out({success:true,action:'deleteLead',foldersTrashed:dFolderIds.length});
}

// ── Tag a Drive folder reference to a LEAD — writes into column T (20th).
// Matched by Record ID (column R) — never mobile alone, since a lead's
// mobile may not be unique. A folder only exists here if the lead actually
// uploaded a document via their shared link; this never creates one. ──
if(data.action==='tagLeadFolder'){
  try{
    var tlSh=ss.getSheetByName(LEADS_TAB);
    if(!tlSh) return out({success:false, error:'No LEADS tab'});
    ensureColumnHeader(tlSh,20,'Folder ID');
    var tlN=tlSh.getLastRow();
    if(tlN>=2){
      var tlRecId=String(data.recordId||'').trim();
      var tlFound=-1;
      if(tlRecId){
        var tlIds=tlSh.getRange(2,18,tlN-1,1).getValues();
        for(var tli=0;tli<tlIds.length;tli++){ if(String(tlIds[tli][0]||'').trim()===tlRecId){ tlFound=tli; break; } }
      }
      if(tlFound<0){
        var tlMob=String(data.mobile||'').trim();
        var tlVals=tlSh.getRange(2,2,tlN-1,1).getValues();
        for(var tli2=0;tli2<tlVals.length;tli2++){ if(String(tlVals[tli2][0]||'').trim()===tlMob){ tlFound=tli2; break; } }
      }
      if(tlFound>=0){
        tlSh.getRange(tlFound+2,20).setValue(data.folderId||'');
        return out({success:true});
      }
    }
    return out({success:false, error:'Lead not found in sheet'});
  }catch(e){
    return out({success:false, error:e.message});
  }
}

  // ── Append ONE client row (new client) — UPSERT keyed on NAME+MOBILE together
  // (not mobile alone), same reasoning as leads. ──
  if(data.action==='appendClient'){
    var aShC=ss.getSheetByName(CLIENTS_TAB)||ss.insertSheet(CLIENTS_TAB);
    if(aShC.getLastRow()===0){
      var cHdr=[['Name','Mobile','Email','Country','Source','Referred By','Next Follow up date',
        'Age','Gender','Remarks','Created date','PAN','DOB','Spouse DOB','Anniversary',
        'Converted Date','Review Date','1st Anniversary','App Installed','Broadcast Added',
        'SIP date','Health insurance','Life insurance','WILL','Due date Health','Due date Life',
        'AUM','Potential AUM','Welcome Message send','Products','Last Contacted','Stage','Upload Token','Potential SIP','Record ID','Family Name','Client ID','SIP date 2','Modified Date','Health Insurance Coverage','Life Insurance Coverage','User ID']];
      var cHr=aShC.getRange(1,1,1,42);
      cHr.setValues(cHdr);cHr.setFontWeight('bold').setBackground('#C0392B').setFontColor('#FFFFFF');
      aShC.setFrozenRows(1);
      ensureStageDropdown(aShC); // new sheet — set the Stage dropdown once, upfront
    }
    ensureColumnHeader(aShC,37,'Client ID');
    ensureColumnHeader(aShC,40,'Health Insurance Coverage');
    ensureColumnHeader(aShC,41,'Life Insurance Coverage');
    ensureColumnHeader(aShC,38,'SIP date 2');
    ensureColumnHeader(aShC,39,'Modified Date');
    ensureColumnHeader(aShC,42,'User ID');
    var cRow=data.row||[];
    var cMob=String(cRow[1]||'').trim();
    var cName=String(cRow[0]||'').trim().toUpperCase();
    var cWrote=false;
    if(cMob&&cName&&aShC.getLastRow()>1){
      var cEx=aShC.getRange(2,1,aShC.getLastRow()-1,2).getValues(); // A=name, B=mobile
      for(var _ci=0;_ci<cEx.length;_ci++){
        if(String(cEx[_ci][1]||'').trim()===cMob && String(cEx[_ci][0]||'').trim().toUpperCase()===cName){
          var cRng=aShC.getRange(_ci+2,1,1,cRow.length);
          cRng.setValues([cRow]);cRng.setNumberFormat('@');cWrote=true;break;
        }
      }
    }
    if(!cWrote){
      if(!String(cRow[36]||'').trim())cRow[36]=nextSequentialId(aShC,37,'C'); // column AK = Client ID (C1, C2...)
      aShC.appendRow(cRow);
      var cLr=aShC.getLastRow();
      aShC.getRange(cLr,1,1,cRow.length).setNumberFormat('@');
    }
    return out({success:true,action:'appendClient',upserted:cWrote,recordId:String(cRow[36]||'')});
  }

  // ── Update ONE client row — matched by Client ID (column AK), never mobile
  // alone. Falls back to name+mobile for legacy calls without a recordId. ──
  if(data.action==='updateClient'){
    var uShC=ss.getSheetByName(CLIENTS_TAB);
    var uFoundC=-1;
    if(uShC&&uShC.getLastRow()>1){
      var uRowC=data.row||[];
      var uNC=uShC.getLastRow();
      var uRecIdC=String(data.recordId||'').trim();
      if(uRecIdC){
        var uIdsC=uShC.getRange(2,37,uNC-1,1).getValues(); // column AK
        for(var uic=0;uic<uIdsC.length;uic++){ if(String(uIdsC[uic][0]||'').trim()===uRecIdC){ uFoundC=uic; break; } }
      }
      if(uFoundC<0){
        var uMobC=String(data.mobile||'').trim();
        var uNameC=String(uRowC[0]||'').trim().toUpperCase();
        var uValsC=uShC.getRange(2,1,uNC-1,2).getValues();
        for(var uic2=0;uic2<uValsC.length;uic2++){
          if(String(uValsC[uic2][1]||'').trim()===uMobC && String(uValsC[uic2][0]||'').trim().toUpperCase()===uNameC){ uFoundC=uic2; break; }
        }
      }
      if(uFoundC>=0){
        // Preserve ownership (column AP, index 41) if the incoming row doesn't
        // specify one — same safeguard as updateLead.
        if(!String(uRowC[41]||'').trim()){
          var uExistingOwnerC=uShC.getRange(uFoundC+2,42).getValue();
          uRowC[41]=uExistingOwnerC||'';
        }
        var uRngC=uShC.getRange(uFoundC+2,1,1,uRowC.length);
        uRngC.setValues([uRowC]);
        uRngC.setNumberFormat('@');
      }
    }
    // Honestly report whether a row was actually found and written — returning
    // success unconditionally here previously masked real failures (no matching
    // Record ID AND no matching name+mobile), silently doing nothing while the
    // app showed a success toast. Now the frontend's existing fallback (a full
    // rewrite via pushClients) correctly triggers when this happens.
    if(uFoundC>=0){
      return out({success:true,action:'updateClient'});
    }
    return out({success:false,action:'updateClient',error:'No matching client row found (checked Record ID and Name+Mobile) — Client ID: '+String(data.recordId||'(none)')+', Mobile: '+String(data.mobile||'(none)')});
  }

  // ── Delete ONE client row — matched by Client ID (column AK) primarily ──
  if(data.action==='deleteClient'){
    var dShC=ss.getSheetByName(CLIENTS_TAB);
    var dFolderIdsC=[]; // collect any tagged folder(s) before the row(s) disappear
    if(dShC&&dShC.getLastRow()>1){
      var dRecIdC=String(data.recordId||'').trim();
      var dNC=dShC.getLastRow();
      if(dRecIdC){
        var dIdsC=dShC.getRange(2,37,dNC-1,1).getValues();
        for(var dic=dIdsC.length-1;dic>=0;dic--){
          if(String(dIdsC[dic][0]||'').trim()===dRecIdC){
            var dFidC=String(dShC.getRange(dic+2,35).getValue()||'').trim(); // column AI = Folder ID
            if(dFidC) dFolderIdsC.push(dFidC);
            dShC.deleteRow(dic+2);
          }
        }
      } else {
        var dMobC=String(data.mobile||'').trim();
        var dNameC=String(data.name||'').trim().toUpperCase();
        var dValsC=dShC.getRange(2,1,dNC-1,2).getValues();
        for(var dic2=dValsC.length-1;dic2>=0;dic2--){
          if(String(dValsC[dic2][1]||'').trim()===dMobC && (!dNameC || String(dValsC[dic2][0]||'').trim().toUpperCase()===dNameC)){
            var dFidC2=String(dShC.getRange(dic2+2,35).getValue()||'').trim();
            if(dFidC2) dFolderIdsC.push(dFidC2);
            dShC.deleteRow(dic2+2);
          }
        }
      }
    }
    // Trash the Drive folder too — soft delete (recoverable from Drive trash for
    // ~30 days), not a permanent wipe. Non-fatal if the folder is already gone.
    dFolderIdsC.forEach(function(fid){
      try{ DriveApp.getFolderById(fid).setTrashed(true); }catch(fe){ /* already gone or inaccessible — skip */ }
    });
    return out({success:true,action:'deleteClient'});
  }
  if(data.action==='backupNow'){ return manualBackupNow(); }
  // ── Get upcoming Google Calendar events (default calendar) ──
  if(data.action==='getCalendarEvents'){
    try{
      var TZ='Asia/Kolkata'; // force IST regardless of Apps Script project timezone setting
      // Pad the fetch window by 1 day on each side so no boundary event is ever missed,
      // then bucket every event to its correct IST calendar day using formatDate (not
      // getFullYear/getMonth/getDate, which use the SERVER's default timezone and can
      // silently shift events onto the wrong day if the project timezone isn't IST).
      var baseStart=data.start?new Date(data.start+'T00:00:00'):new Date();
      var baseEnd=data.end?new Date(data.end+'T23:59:59'):new Date(baseStart.getTime()+8*86400000);
      var evStart=new Date(baseStart.getTime()-86400000); // -1 day padding
      var evEnd=new Date(baseEnd.getTime()+86400000);     // +1 day padding
      var wantStart=data.start||Utilities.formatDate(new Date(),TZ,'yyyy-MM-dd');
      var wantEnd=data.end||Utilities.formatDate(new Date(baseStart.getTime()+8*86400000),TZ,'yyyy-MM-dd');

      var cal=CalendarApp.getDefaultCalendar();
      var events=cal.getEvents(evStart,evEnd);
      var evRows=events.map(function(e){
        var st=e.getStartTime();
        var dateStr=Utilities.formatDate(st,TZ,'yyyy-MM-dd');
        var timeStr=e.isAllDayEvent()?'All day':Utilities.formatDate(st,TZ,'h:mm a');
        return {date:dateStr,title:e.getTitle(),time:timeStr,type:'event'};
      }).filter(function(r){return r.date>=wantStart&&r.date<=wantEnd;}); // trim padding back to requested range

      // ── Also pull Birthdays (Google's auto-generated Contacts birthday calendar) ──
      var bdayRows=[];
      try{
        var allCals=CalendarApp.getAllCalendars();
        var bdayCal=null;
        for(var bi=0;bi<allCals.length;bi++){
          var cName=allCals[bi].getName().toLowerCase();
          var cId=allCals[bi].getId();
          if(cName.indexOf('birthday')>=0 || cId.indexOf('#contacts@group.v.calendar.google.com')>=0){
            bdayCal=allCals[bi]; break;
          }
        }
        if(bdayCal){
          var bdayEvents=bdayCal.getEvents(evStart,evEnd);
          bdayRows=bdayEvents.map(function(e){
            var st=e.getStartTime();
            var dateStr=Utilities.formatDate(st,TZ,'yyyy-MM-dd');
            return {date:dateStr,title:e.getTitle(),time:'',type:'birthday'};
          }).filter(function(r){return r.date>=wantStart&&r.date<=wantEnd;});
        }
      }catch(be){ /* birthday calendar not accessible — skip silently */ }

      return out({events:evRows.concat(bdayRows), calendarId:cal.getId(), calendarName:cal.getName(), birthdaysFound:bdayRows.length, scriptTimeZone:Session.getScriptTimeZone(), usedTimeZone:TZ});
    }catch(e){
      return out({events:[],error:e.message});
    }
  }

  // ── Diagnose calendar access — lists ALL calendars this account can see ──
  if(data.action==='diagnoseCalendars'){
    try{
      var allCals=CalendarApp.getAllCalendars();
      var calList=allCals.map(function(c){
        var isBday=c.getName().toLowerCase().indexOf('birthday')>=0||c.getId().indexOf('#contacts@group.v.calendar.google.com')>=0;
        return {name:c.getName(), id:c.getId(), isSelected:c.isSelected(), isOwned:c.isOwnedByMe(), isBirthdayCalendar:isBday};
      });
      var defCal=CalendarApp.getDefaultCalendar();
      return out({
        defaultCalendarId: defCal.getId(),
        defaultCalendarName: defCal.getName(),
        allCalendars: calList
      });
    }catch(e){
      return out({error:e.message});
    }
  }

  // ── Client Documents: get-or-create a client's Drive folder ──
  // Folder naming: "ClientName +Mobile" — name+mobile only, no PAN in the name/search.
  function clientFolderName(clientName, pan, mobile){
    return (clientName||'Unknown').trim()+' +'+(mobile||'').trim();
  }

  function getParentDocsFolder(){
    var parentName='Wealth Matrix Client Documents';
    var parentIter=DriveApp.getFoldersByName(parentName);
    return parentIter.hasNext()?parentIter.next():DriveApp.createFolder(parentName);
  }

  // Search for an existing client folder using name+mobile and name-only matching
  // (no PAN). Exact "Name +Mobile" match is always safe to use directly. A name-only
  // match is ambiguous — for this automated path (no human present, e.g. a client
  // uploading via their own link) it's only auto-resolved when there's EXACTLY ONE
  // candidate; two or more, or zero, means a fresh folder gets created instead.
  function findClientFolder(clientName, pan, mobile){
    var parent=getParentDocsFolder();
    var nameUp=(clientName||'').trim().toUpperCase();
    var newFormatUp=nameUp+' +'+(mobile||'').trim();
    var starts=[];
    var allFolders=parent.getFolders();
    while(allFolders.hasNext()){
      var f=allFolders.next();
      var fUp=f.getName().trim().toUpperCase();
      if(fUp===newFormatUp) return f;               // exact match — always safe
      if(fUp.indexOf(nameUp)===0) starts.push(f);    // name-only candidate
    }
    return starts.length===1 ? starts[0] : null;
  }

  function getOrCreateClientFolder(clientName, pan, mobile){
    var existing=findClientFolder(clientName, pan, mobile);
    if(existing) return existing;
    var parent=getParentDocsFolder();
    return parent.createFolder(clientFolderName(clientName, pan, mobile));
  }

  // ── QUICK REMINDERS ──
  // ── DPDP consent log — Column A: Name, B: Timestamp, C: exact consent text
  // the client agreed to. Self-heals the tab/header if not already created. ──
  if(data.action==='logDpdpConsent'){
    try{
      var dpSh=ss.getSheetByName('DPDP Consents')||ss.insertSheet('DPDP Consents');
      if(dpSh.getLastRow()===0){
        var dpHdr=[['Name','Timestamp','Consent Text']];
        var dpHr=dpSh.getRange(1,1,1,3);
        dpHr.setValues(dpHdr);
        dpHr.setFontWeight('bold').setBackground('#C0392B').setFontColor('#FFFFFF');
        dpSh.setFrozenRows(1);
      }
      var dpStamp=Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
      dpSh.appendRow([data.name||'', dpStamp, data.consentText||'']);
      var dpLr=dpSh.getLastRow();
      dpSh.getRange(dpLr,1).setNumberFormat('@');
      dpSh.getRange(dpLr,2).setNumberFormat('@');
      return out({success:true});
    }catch(e){
      return out({success:false, error:e.message});
    }
  }

  // ── TASKS (formerly Reminders) ── unique Task ID based, single-row targeted
  // writes throughout — no full-table scans, no redundant locking (the outer
  // writeData() wrapper already holds the script lock for this whole call).
  // Resilient to the sheet tab being named either "TASKS" (new) or "REMINDERS"
  // (the original name, if it was never renamed) — tries TASKS first, falls
  // back to REMINDERS so existing data is never silently invisible.
  function getTasksSheet(ss, createIfMissing){
    var sh=ss.getSheetByName(TASKS_TAB);
    if(sh) return sh;
    sh=ss.getSheetByName(REMINDERS_TAB);
    if(sh) return sh;
    return createIfMissing ? ss.insertSheet(TASKS_TAB) : null;
  }

  if(data.action==='addTask'){
    try{
      var tSh=getTasksSheet(ss, true);
      if(tSh.getLastRow()===0){
        var tHdr=[['Created Date','Name','Task Details','Due Date','Status','Task For','Task ID','Linked Record ID','Created By','Created By Role','Reviewed By','Reviewed Date','Review Comment']];
        var tHr=tSh.getRange(1,1,1,13);
        tHr.setValues(tHdr);
        tHr.setFontWeight('bold').setBackground('#C0392B').setFontColor('#FFFFFF');
        tSh.setFrozenRows(1);
      } else {
        // Existing sheet from before this system — repurpose the old, always-
        // empty "Assigned To" column as "Created By" (never actually used for
        // its original purpose), and append the new review-workflow columns.
        var tHdrCell=tSh.getRange(1,9);
        if(String(tHdrCell.getValue()||'').trim().toLowerCase()==='assigned to'){ tHdrCell.setValue('Created By'); }
        ensureColumnHeader(tSh,9,'Created By');
        ensureColumnHeader(tSh,10,'Created By Role');
        ensureColumnHeader(tSh,11,'Reviewed By');
        ensureColumnHeader(tSh,12,'Reviewed Date');
        ensureColumnHeader(tSh,13,'Review Comment');
      }
      var newId=nextSequentialId(tSh,7,'T'); // column G = Task ID
      tSh.appendRow([data.createdDate||'', data.name||'', data.taskDetails||'', data.dueDate||'', data.status||'Open', data.taskFor||'OTHERS', newId, data.linkedRecordId||'', data.createdBy||'', data.createdByRole||'', '', '', '']);
      var tLr=tSh.getLastRow();
      tSh.getRange(tLr,1,1,13).setNumberFormat('@');
      return out({success:true, taskId:newId});
    }catch(e){
      return out({success:false, error:e.message});
    }
  }

  // ── Fills in a Task ID (and a default "OTHERS" Task For) for any row that
  // predates this system — e.g. rows created back when this was a simple
  // Reminders sheet with no ID column. Called automatically by the frontend
  // whenever it notices ID-less rows, so old entries become fully manageable
  // (Delete/Finish/Save) instead of silently staying invisible. ──
  if(data.action==='backfillTaskIds'){
    try{
      var bSh=getTasksSheet(ss, false);
      if(!bSh || bSh.getLastRow()<2) return out({success:true, filled:0});
      var bLr=bSh.getLastRow();
      // Make sure every column has the right header, even on a sheet that
      // started life as the old 5-column (A-E) Reminders layout. Uses
      // ensureColumnHeader consistently (idempotent — only sets if blank)
      // rather than manual checks, so this function can never again drift
      // out of sync with what addTask actually sets up.
      ensureColumnHeader(bSh,6,'Task For');
      ensureColumnHeader(bSh,7,'Task ID');
      ensureColumnHeader(bSh,8,'Linked Record ID');
      ensureColumnHeader(bSh,9,'Created By');
      ensureColumnHeader(bSh,10,'Created By Role');
      ensureColumnHeader(bSh,11,'Reviewed By');
      ensureColumnHeader(bSh,12,'Reviewed Date');
      ensureColumnHeader(bSh,13,'Review Comment');

      var bVals=bSh.getRange(2,6,bLr-1,2).getValues(); // F=Task For, G=Task ID
      var filled=0;
      for(var bi=0;bi<bVals.length;bi++){
        var existingId=String(bVals[bi][1]||'').trim();
        if(!existingId){
          var bRow=bi+2;
          var newTaskId=nextSequentialId(bSh,7,'T');
          bSh.getRange(bRow,7).setValue(newTaskId);
          if(!String(bVals[bi][0]||'').trim()) bSh.getRange(bRow,6).setValue('OTHERS');
          filled++;
        }
      }
      return out({success:true, filled:filled});
    }catch(e){
      return out({success:false, error:e.message});
    }
  }

  if(data.action==='updateTask'){
    try{
      var uTSh=getTasksSheet(ss, false);
      if(!uTSh) return out({success:false, error:'No TASKS/REMINDERS tab yet'});
      var uTLr=uTSh.getLastRow();
      if(uTLr>=2){
        var uTIds=uTSh.getRange(2,7,uTLr-1,1).getValues(); // column G = Task ID
        for(var uti=0;uti<uTIds.length;uti++){
          if(String(uTIds[uti][0]||'').trim()===String(data.taskId||'').trim()){
            var uTRow=uti+2;
            // Batched into one read + one write (was up to 6 separate setValue
            // calls before) — columns B-H: Name, Task Details, Due Date,
            // Status, Task For, Task ID (untouched), Linked Record ID.
            var uTRange=uTSh.getRange(uTRow,2,1,7);
            var uTCur=uTRange.getValues()[0];
            if(data.name!==undefined) uTCur[0]=data.name;
            if(data.taskDetails!==undefined) uTCur[1]=data.taskDetails;
            if(data.dueDate!==undefined) uTCur[2]=data.dueDate;
            if(data.status!==undefined) uTCur[3]=data.status;
            if(data.taskFor!==undefined) uTCur[4]=data.taskFor;
            // uTCur[5] = Task ID (column G) — never modified here
            if(data.linkedRecordId!==undefined) uTCur[6]=data.linkedRecordId;
            uTRange.setValues([uTCur]);
            return out({success:true});
          }
        }
      }
      return out({success:false, error:'Task not found'});
    }catch(e){
      return out({success:false, error:e.message});
    }
  }

  // ── Task review workflow ──
  // submitTaskDone: the "Finish" button now calls this instead of directly
  // setting status='Finished'. Reads the task's OWN "Created By Role" column
  // (never trusts a role claim sent from the frontend) to decide the correct
  // next status — Maker-created tasks always route through review first,
  // regardless of who actually marks it done; Checker/Authorizer/Admin-
  // created tasks go straight to Finished, matching the confirmed design.
  // submitTaskDone: routes based on WHO IS PRESSING FINISH, not who created
  // the task — a Checker or Admin finishing a Maker-created task goes
  // straight to Finished (they're already trusted), only a Maker finishing
  // it routes to review. Verified server-side via verifyCanReview (never
  // trusts a role claim sent from the frontend), same as the review actions.
  if(data.action==='submitTaskDone'){
    try{
      var sTdSh=getTasksSheet(ss, false);
      if(!sTdSh) return out({success:false, error:'No TASKS tab yet'});
      var sTdLr=sTdSh.getLastRow();
      if(sTdLr>=2){
        var sTdIds=sTdSh.getRange(2,7,sTdLr-1,1).getValues();
        for(var sTdi=0;sTdi<sTdIds.length;sTdi++){
          if(String(sTdIds[sTdi][0]||'').trim()===String(data.taskId||'').trim()){
            var sTdRow=sTdi+2;
            var actorCanSkipReview=verifyCanReview(data.requestingUserId);
            var newStatus=actorCanSkipReview ? 'Finished' : 'Pending Review';
            sTdSh.getRange(sTdRow,5).setValue(newStatus); // column E = Status
            return out({success:true, newStatus:newStatus});
          }
        }
      }
      return out({success:false, error:'Task not found'});
    }catch(e){
      return out({success:false, error:e.message});
    }
  }

  // approveTaskReview: Checker confirms a Pending Review task is genuinely
  // done. Moves to Finished, records who approved it and when.
  if(data.action==='approveTaskReview'){
    try{
      if(!verifyCanReview(data.requestingUserId)) return out({success:false, error:'Not authorized to review tasks'});
      var aTrSh=getTasksSheet(ss, false);
      if(!aTrSh) return out({success:false, error:'No TASKS tab yet'});
      var aTrLr=aTrSh.getLastRow();
      if(aTrLr>=2){
        var aTrIds=aTrSh.getRange(2,7,aTrLr-1,1).getValues();
        for(var aTri=0;aTri<aTrIds.length;aTri++){
          if(String(aTrIds[aTri][0]||'').trim()===String(data.taskId||'').trim()){
            var aTrRow=aTri+2;
            aTrSh.getRange(aTrRow,5).setValue('Finished'); // column E
            aTrSh.getRange(aTrRow,11).setValue(data.reviewerName||''); // column K = Reviewed By
            aTrSh.getRange(aTrRow,12).setValue(Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss')); // column L = Reviewed Date
            return out({success:true});
          }
        }
      }
      return out({success:false, error:'Task not found'});
    }catch(e){
      return out({success:false, error:e.message});
    }
  }

  // sendBackTaskReview: Checker finds the work isn't actually done properly.
  // Returns it to Due (not Finished, not still Pending Review) with a
  // comment the Maker can see, so it's clearly back in their court to redo.
  if(data.action==='sendBackTaskReview'){
    try{
      if(!verifyCanReview(data.requestingUserId)) return out({success:false, error:'Not authorized to review tasks'});
      var sBSh=getTasksSheet(ss, false);
      if(!sBSh) return out({success:false, error:'No TASKS tab yet'});
      var sBLr=sBSh.getLastRow();
      if(sBLr>=2){
        var sBIds=sBSh.getRange(2,7,sBLr-1,1).getValues();
        for(var sBi=0;sBi<sBIds.length;sBi++){
          if(String(sBIds[sBi][0]||'').trim()===String(data.taskId||'').trim()){
            var sBRow=sBi+2;
            sBSh.getRange(sBRow,5).setValue('Due'); // column E
            sBSh.getRange(sBRow,11).setValue(''); // Reviewed By cleared — not actually approved
            sBSh.getRange(sBRow,12).setValue(''); // Reviewed Date cleared
            sBSh.getRange(sBRow,13).setValue(data.comment||''); // column M = Review Comment
            return out({success:true});
          }
        }
      }
      return out({success:false, error:'Task not found'});
    }catch(e){
      return out({success:false, error:e.message});
    }
  }

  if(data.action==='deleteTask'){
    try{
      var dTSh=getTasksSheet(ss, false);
      if(!dTSh) return out({success:true});
      var dTLr=dTSh.getLastRow();
      if(dTLr>=2){
        var dTIds=dTSh.getRange(2,7,dTLr-1,1).getValues();
        for(var dti=dTIds.length-1;dti>=0;dti--){
          if(String(dTIds[dti][0]||'').trim()===String(data.taskId||'').trim()){
            dTSh.deleteRow(dti+2);
          }
        }
      }
      return out({success:true});
    }catch(e){
      return out({success:false, error:e.message});
    }
  }

  // ── Bulk-mark overdue "Open" tasks as "Due" — called silently in the
  // background whenever the Task list loads, so the sheet itself reflects
  // "Due" for manual review without needing a scheduled trigger. ──
  if(data.action==='syncOverdueTasks'){
    try{
      var sTSh=getTasksSheet(ss, false);
      if(!sTSh || sTSh.getLastRow()<2) return out({success:true, updated:0});
      var sTLr=sTSh.getLastRow();
      var sTVals=sTSh.getRange(2,4,sTLr-1,2).getValues(); // D=Due Date, E=Status
      var todayStr=Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
      var updated=0;
      var statusColumn=[]; // whole column built in memory, written back in one call
      for(var sti=0;sti<sTVals.length;sti++){
        var dueRaw=sTVals[sti][0];
        var dueStr=dueRaw instanceof Date ? Utilities.formatDate(dueRaw, Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(dueRaw||'').trim();
        var status=String(sTVals[sti][1]||'').trim();
        if(status.toUpperCase()==='OPEN' && dueStr && dueStr<=todayStr){
          statusColumn.push(['Due']);
          updated++;
        } else {
          statusColumn.push([status]); // unchanged
        }
      }
      if(updated>0){
        sTSh.getRange(2,5,sTLr-1,1).setValues(statusColumn); // single batched write, only when actually needed
      }
      return out({success:true, updated:updated});
    }catch(e){
      return out({success:false, error:e.message});
    }
  }

  if(data.action==='addReminder'){
    try{
      var rSh=ss.getSheetByName(REMINDERS_TAB)||ss.insertSheet(REMINDERS_TAB);
      if(rSh.getLastRow()===0){
        var rHdr=[['Date','Client Name','Reminder','Reminder Date']];
        var rHr=rSh.getRange(1,1,1,4);
        rHr.setValues(rHdr);
        rHr.setFontWeight('bold').setBackground('#C0392B').setFontColor('#FFFFFF');
        rSh.setFrozenRows(1);
      }
      var lock=LockService.getScriptLock(); lock.waitLock(10000);
      try{
        rSh.appendRow([data.date||'', data.clientName||'', data.reminderText||'', data.reminderDate||'']);
      } finally { lock.releaseLock(); }
      return out({success:true});
    }catch(e){
      return out({success:false, error:e.message});
    }
  }

  function normDate(c){
    if(c===null||c===undefined||c==='') return '';
    if(c instanceof Date){
      var y=c.getFullYear(), m=String(c.getMonth()+1).padStart(2,'0'), d=String(c.getDate()).padStart(2,'0');
      return y+'-'+m+'-'+d;
    }
    return String(c).trim();
  }

  if(data.action==='updateReminder'){
    try{
      var rSh2=ss.getSheetByName(REMINDERS_TAB);
      if(!rSh2) return out({success:false, error:'No REMINDERS tab yet'});
      var rLr=rSh2.getLastRow();
      if(rLr>=2){
        var rVals=rSh2.getRange(2,1,rLr-1,4).getValues();
        for(var ri=0;ri<rVals.length;ri++){
          if(normDate(rVals[ri][0])===String(data.oldDate||'').trim()
            && String(rVals[ri][1]||'').trim()===String(data.oldClientName||'').trim()
            && String(rVals[ri][2]||'').trim()===String(data.oldReminderText||'').trim()
            && normDate(rVals[ri][3])===String(data.oldReminderDate||'').trim()){
            rSh2.getRange(ri+2,1,1,4).setValues([[data.newDate||'', data.newClientName||'', data.newReminderText||'', data.newReminderDate||'']]);
            break;
          }
        }
      }
      return out({success:true});
    }catch(e){
      return out({success:false, error:e.message});
    }
  }

  if(data.action==='deleteReminder'){
    try{
      var rSh3=ss.getSheetByName(REMINDERS_TAB);
      if(!rSh3) return out({success:true}); // nothing to delete
      var rLr3=rSh3.getLastRow();
      if(rLr3>=2){
        var rVals3=rSh3.getRange(2,1,rLr3-1,4).getValues();
        for(var ri3=rVals3.length-1;ri3>=0;ri3--){
          if(normDate(rVals3[ri3][0])===String(data.date||'').trim()
            && String(rVals3[ri3][1]||'').trim()===String(data.clientName||'').trim()
            && String(rVals3[ri3][2]||'').trim()===String(data.reminderText||'').trim()
            && normDate(rVals3[ri3][3])===String(data.reminderDate||'').trim()){
            rSh3.deleteRow(ri3+2);
          }
        }
      }
      return out({success:true});
    }catch(e){
      return out({success:false, error:e.message});
    }
  }

  // ── Automated Insights control actions ──
  if(data.action==='draftLeadMessage'){
    try{
      return out({success:true, message:draftLeadMessage(data)});
    }catch(e){
      return out({success:false, error:e.message});
    }
  }

  if(data.action==='draftClientMessage'){
    try{
      return out({success:true, message:draftClientMessage(data)});
    }catch(e){
      return out({success:false, error:e.message});
    }
  }

  if(data.action==='generateLeadProfileInsight'){
    try{
      return out({success:true, insight:generateLeadProfileInsight(data)});
    }catch(e){
      return out({success:false, error:e.message});
    }
  }

  if(data.action==='generateClientProfileInsight'){
    try{
      return out({success:true, insight:generateClientProfileInsight(data)});
    }catch(e){
      return out({success:false, error:e.message});
    }
  }

  if(data.action==='generateMeetingInsight'){
    try{
      return out({success:true, insight:generateMeetingInsight(data)});
    }catch(e){
      return out({success:false, error:e.message});
    }
  }

  // ── Generate the client-facing review report and save it so the client-
  // facing link can display it instantly (no live AI call when the client
  // opens their link — it's already sitting there, waiting). ──
  // ── Our firm's ARN/Advisor Code — a firm-level constant, asked once via
  // popup on first use of the Review Report feature, then reused for every
  // report going forward. Lets the AI distinguish funds actually managed by
  // this firm from funds a client holds elsewhere (bank, another distributor,
  // or direct), so growth isn't credited to funds never actually advised on. ──
  if(data.action==='getOurArnCode'){
    try{
      var arnVal=PropertiesService.getScriptProperties().getProperty('OUR_ARN_CODE')||'';
      return out({success:true, arnCode:arnVal});
    }catch(e){
      return out({success:false, error:e.message});
    }
  }
  if(data.action==='saveOurArnCode'){
    try{
      PropertiesService.getScriptProperties().setProperty('OUR_ARN_CODE', String(data.arnCode||'').trim());
      return out({success:true});
    }catch(e){
      return out({success:false, error:e.message});
    }
  }

  if(data.action==='generateAndSaveReviewReport'){
    try{
      var reportText=generateClientReviewReport(data);
      var rrSs=SpreadsheetApp.openById(SHEET_ID);
      var rrSh=rrSs.getSheetByName('REVIEW_REPORTS')||rrSs.insertSheet('REVIEW_REPORTS');
      if(rrSh.getLastRow()===0){
        var rrHdr=[['Record ID','Name','Mobile','Report Text','Generated Date']];
        var rrHr=rrSh.getRange(1,1,1,5);
        rrHr.setValues(rrHdr);
        rrHr.setFontWeight('bold').setBackground('#C0392B').setFontColor('#FFFFFF');
        rrSh.setFrozenRows(1);
      }
      var rrStamp=Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
      // One row per client — replace any previous report rather than piling up history
      var rrLr=rrSh.getLastRow();
      var rrFound=-1;
      if(rrLr>=2){
        var rrIds=rrSh.getRange(2,1,rrLr-1,1).getValues();
        for(var rri=0;rri<rrIds.length;rri++){
          if(String(rrIds[rri][0]||'').trim()===String(data.recordId||'').trim()){ rrFound=rri; break; }
        }
      }
      if(rrFound>=0){
        rrSh.getRange(rrFound+2,1,1,5).setValues([[data.recordId||'', data.name||'', data.mobile||'', reportText, rrStamp]]);
      } else {
        rrSh.appendRow([data.recordId||'', data.name||'', data.mobile||'', reportText, rrStamp]);
      }
      return out({success:true, report:reportText});
    }catch(e){
      return out({success:false, error:e.message});
    }
  }

  // ── Client-facing fetch — resolved by TOKEN only (never by client-supplied
  // mobile/recordId), same security rule as every other public-facing action,
  // since this endpoint has no authentication. ──
  if(data.action==='getClientReviewReport'){
    try{
      var gtTok=String(data.token||'').trim();
      if(!gtTok) return out({success:false, error:'Missing token'});
      var gtSs=SpreadsheetApp.openById(SHEET_ID);
      var gtClientSh=gtSs.getSheetByName(CLIENTS_TAB);
      if(!gtClientSh) return out({success:false, error:'No CLIENTS tab'});
      var gtLr=gtClientSh.getLastRow();
      var gtRecordId='';
      if(gtLr>=2){
        var gtVals=gtClientSh.getRange(2,1,gtLr-1,37).getValues();
        for(var gti=0;gti<gtVals.length;gti++){
          if(String(gtVals[gti][32]||'').trim()===gtTok){ gtRecordId=String(gtVals[gti][36]||'').trim(); break; } // AG=Upload Token, AK=Client ID
        }
      }
      if(!gtRecordId) return out({success:false, error:'Link not found or expired'});
      var gtReportSh=gtSs.getSheetByName('REVIEW_REPORTS');
      if(!gtReportSh || gtReportSh.getLastRow()<2) return out({success:false, error:'No report available yet'});
      var gtRows=gtReportSh.getRange(2,1,gtReportSh.getLastRow()-1,5).getValues();
      for(var gtj=0;gtj<gtRows.length;gtj++){
        if(String(gtRows[gtj][0]||'').trim()===gtRecordId){
          return out({success:true, name:String(gtRows[gtj][1]||''), report:String(gtRows[gtj][3]||''), generatedDate:String(gtRows[gtj][4]||'')});
        }
      }
      return out({success:false, error:'No report available yet'});
    }catch(e){
      return out({success:false, error:e.message});
    }
  }

  if(data.action==='listCasFilesInFolder'){
    try{
      return out({success:true, files:listCasFilesInFolder(data)});
    }catch(e){
      return out({success:false, error:e.message});
    }
  }

  if(data.action==='draftCasEmailInsight'){
    try{
      var casDraft=draftCasEmailInsight(data);
      // Audit trail: record that AI-assisted CAS analysis (which processes the
      // client's PAN for internal matching) actually ran, when, and for whom.
      // Logged to DPDP Consents (not Activity Log) so every DPDP-relevant event —
      // the original consent AND each instance it was relied upon — lives in one
      // clean, audit-ready sheet, rather than buried among unrelated activity.
      try{
        var casDpSh=ss.getSheetByName('DPDP Consents')||ss.insertSheet('DPDP Consents');
        if(casDpSh.getLastRow()===0){
          var casDpHdr=[['Name','Timestamp','Consent Text']];
          var casDpHr=casDpSh.getRange(1,1,1,3);
          casDpHr.setValues(casDpHdr);
          casDpHr.setFontWeight('bold').setBackground('#C0392B').setFontColor('#FFFFFF');
          casDpSh.setFrozenRows(1);
        }
        var casStamp=Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
        var casNote='Earlier consent to allow AI to handle PAN (for CAS matching only) — executed. Files analyzed: '+(casDraft.filesUsed||[]).join(', ');
        casDpSh.appendRow([data.clientName||'', casStamp, casNote]);
        var casDpLr=casDpSh.getLastRow();
        casDpSh.getRange(casDpLr,2).setNumberFormat('@');
      }catch(logErr){ /* non-fatal — never block the actual feature over an audit-log write */ }
      return out({success:true, subject:casDraft.subject, body:casDraft.body, tableHtml:casDraft.tableHtml, filesUsed:casDraft.filesUsed});
    }catch(e){
      return out({success:false, error:e.message});
    }
  }

  if(data.action==='generateProductivityInsight'){
    try{ return out({success:true, insight:generateProductivityInsight()}); }
    catch(e){ return out({success:false, error:e.message}); }
  }
  if(data.action==='generateEngagementInsight'){
    try{ return out({success:true, insight:generateEngagementInsight()}); }
    catch(e){ return out({success:false, error:e.message}); }
  }
  if(data.action==='generateStaleLeadsInsight'){
    try{ var r=generateStaleLeadsInsight(); return out({success:true, insight:r.insight, list:r.list}); }
    catch(e){ return out({success:false, error:e.message}); }
  }
  if(data.action==='generateNeverContactedInsight'){
    try{ return out({success:true, insight:generateNeverContactedInsight()}); }
    catch(e){ return out({success:false, error:e.message}); }
  }
  if(data.action==='generateProtectionGapsInsight'){
    try{ return out({success:true, insight:generateProtectionGapsInsight()}); }
    catch(e){ return out({success:false, error:e.message}); }
  }
  // ═══════════════════════════════════════════════
  // FIN DOC — Family Financial Information & Documents Organizer. Stored as
  // one JSON blob per client (not flat rows) in a dedicated FIN_DOC_DATA
  // tab, since the form has many variable-length repeating sections (family
  // members, bank accounts, policies...) that don't fit a normal row shape.
  // Reuses the SAME upload-token system already used for document uploads —
  // one token per lead/client, not a separate token scheme.
  // ═══════════════════════════════════════════════

  // Shared by all three Fin Doc actions — resolves a token to the owning
  // record (lead or client), same lookup resolveUploadToken does, but
  // callable from here since getOrCreateClientFolder only exists in this scope.
  function resolveTokenInThisScope(token){
    var tok=String(token||'').trim();
    if(!tok) return null;
    var cSh=ss.getSheetByName(CLIENTS_TAB);
    if(cSh && cSh.getLastRow()>=2){
      var cVals=cSh.getRange(2,1,cSh.getLastRow()-1,41).getValues();
      for(var ci=0;ci<cVals.length;ci++){
        if(String(cVals[ci][32]||'').trim()===tok){
          return {recordId:String(cVals[ci][36]||''), name:String(cVals[ci][0]||''), mobile:String(cVals[ci][1]||''), pan:String(cVals[ci][11]||''), recordType:'client'};
        }
      }
    }
    var lSh=ss.getSheetByName(LEADS_TAB);
    if(lSh && lSh.getLastRow()>=2){
      var lVals=lSh.getRange(2,1,lSh.getLastRow()-1,21).getValues();
      for(var li=0;li<lVals.length;li++){
        if(String(lVals[li][20]||'').trim()===tok){
          return {recordId:String(lVals[li][17]||''), name:String(lVals[li][0]||''), mobile:String(lVals[li][1]||''), pan:'', recordType:'lead'};
        }
      }
    }
    return null;
  }

  function getFinDocSheet(){
    var fdSh=ss.getSheetByName('FIN_DOC_DATA')||ss.insertSheet('FIN_DOC_DATA');
    if(fdSh.getLastRow()===0){
      var fdHdr=[['Record ID','Name','Mobile','Form Data (JSON)','Last Updated','PDF Generated Date']];
      var fdHr=fdSh.getRange(1,1,1,6);
      fdHr.setValues(fdHdr);
      fdHr.setFontWeight('bold').setBackground('#C0392B').setFontColor('#FFFFFF');
      fdSh.setFrozenRows(1);
    }
    return fdSh;
  }

  if(data.action==='getFinDocData'){
    try{
      var gfdRecord=resolveTokenInThisScope(data.token);
      if(!gfdRecord) return out({success:false, error:'Invalid or expired link'});
      var gfdSh=getFinDocSheet();
      if(gfdSh.getLastRow()>=2){
        var gfdVals=gfdSh.getRange(2,1,gfdSh.getLastRow()-1,6).getValues();
        for(var gfdi=0;gfdi<gfdVals.length;gfdi++){
          if(String(gfdVals[gfdi][0]||'').trim()===gfdRecord.recordId){
            return out({success:true, name:gfdRecord.name, formData:gfdVals[gfdi][3]||'{}', lastUpdated:String(gfdVals[gfdi][4]||'')});
          }
        }
      }
      return out({success:true, name:gfdRecord.name, formData:'{}', lastUpdated:''}); // first visit — nothing saved yet
    }catch(e){
      return out({success:false, error:e.message});
    }
  }

  if(data.action==='saveFinDocProgress'){
    try{
      var sfdRecord=resolveTokenInThisScope(data.token);
      if(!sfdRecord) return out({success:false, error:'Invalid or expired link'});
      var sfdSh=getFinDocSheet();
      var sfdNow=Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
      var sfdFound=-1;
      if(sfdSh.getLastRow()>=2){
        var sfdIds=sfdSh.getRange(2,1,sfdSh.getLastRow()-1,1).getValues();
        for(var sfdi=0;sfdi<sfdIds.length;sfdi++){
          if(String(sfdIds[sfdi][0]||'').trim()===sfdRecord.recordId){ sfdFound=sfdi+2; break; }
        }
      }
      var sfdRow=[sfdRecord.recordId, sfdRecord.name, sfdRecord.mobile, data.formData||'{}', sfdNow, ''];
      if(sfdFound>=0){
        // Preserve any existing PDF-generated date — saving progress shouldn't erase it
        var sfdExistingPdfDate=sfdSh.getRange(sfdFound,6).getValue();
        sfdRow[5]=sfdExistingPdfDate||'';
        sfdSh.getRange(sfdFound,1,1,6).setValues([sfdRow]);
      } else {
        sfdSh.appendRow(sfdRow);
      }
      return out({success:true, lastUpdated:sfdNow});
    }catch(e){
      return out({success:false, error:e.message});
    }
  }

  if(data.action==='generateFinDocPdf'){
    try{
      var gpRecord=resolveTokenInThisScope(data.token);
      if(!gpRecord) return out({success:false, error:'Invalid or expired link'});
      var gpSh=getFinDocSheet();
      var gpFormData=null, gpRow=-1;
      if(gpSh.getLastRow()>=2){
        var gpVals=gpSh.getRange(2,1,gpSh.getLastRow()-1,6).getValues();
        for(var gpi=0;gpi<gpVals.length;gpi++){
          if(String(gpVals[gpi][0]||'').trim()===gpRecord.recordId){ gpFormData=gpVals[gpi][3]; gpRow=gpi+2; break; }
        }
      }
      if(!gpFormData) return out({success:false, error:'No saved progress to generate from yet'});
      var gpData=JSON.parse(gpFormData||'{}');

      // Build a temporary Google Doc, formatted from whatever sections/rows
      // actually have data — empty sections are skipped entirely, never
      // rendered as blank rows, per the "partial report" requirement.
      var gpDoc=DocumentApp.create('TEMP_FinDoc_'+gpRecord.recordId+'_'+Date.now());
      var gpBody=gpDoc.getBody();
      gpBody.appendParagraph('Family Financial Information & Documents Organizer').setHeading(DocumentApp.ParagraphHeading.TITLE);
      gpBody.appendParagraph('Prepared for: '+gpRecord.name).setHeading(DocumentApp.ParagraphHeading.NORMAL);
      gpBody.appendParagraph('Generated: '+Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd MMM yyyy')).setHeading(DocumentApp.ParagraphHeading.NORMAL);

      var gpSections=[
        {key:'familyMembers', title:'Family Members', cols:['Couple Name','Dependent Children','Dependent Parents','Dependent Relatives'], fields:['coupleName','dependentChildren','dependentParents','dependentRelatives']},
        {key:'contacts', title:'Important Contacts', cols:['Particulars','Name','Mobile','Email'], fields:['particulars','name','mobile','email']},
        {key:'documents', title:'Important Financial Documents', cols:['Document','Holder Name','Doc Number','Physical Location','Virtual Location'], fields:['document','holderName','docNumber','physicalLocation','virtualLocation']},
        {key:'docValidity', title:'Documents Validity', cols:['Name','Document','Number','Valid Till','Changes Required'], fields:['name','document','number','validTill','changesRequired']},
        {key:'lockers', title:'Locker Details', cols:['Bank & Branch','Account Number','Locker No.','In Name Of','Code','Nominee'], fields:['bankBranch','accountNumber','lockerNo','inNameOf','code','nominee']},
        {key:'passwords', title:'Online Passwords (who knows them, not the password itself)', cols:['Login Kind','Person Aware'], fields:['loginKind','personAware']},
        {key:'lifeInsurance', title:'Life Insurance Policies', cols:['Company & Policy','Policy No.','Amount Insured','Valid Till','Premium','Nominee','Advisor','Advisor Contact'], fields:['companyPolicy','policyNo','amountInsured','validTill','premium','nominee','advisorName','advisorContact']},
        {key:'healthInsurance', title:'Health & Other Insurance', cols:['Type','Policy Name','Policy No.','Covered','Amount','Valid Till','Premium','Advisor','Advisor Contact'], fields:['type','policyName','policyNo','covered','amount','validTill','premium','advisorName','advisorContact']},
        {key:'bankAccounts', title:'Bank Account Details', cols:['Holder Name','Bank & Branch','Account No.','Type','Nominee','Email','Mobile','Username'], fields:['holderName','bankBranch','accountNumber','type','nominee','email','mobile','username']},
        {key:'cards', title:'Debit / Credit Card Details', cols:['Cardholder','Card Type','Card Number','Linked Account','ATM/Debit No.','Valid Till'], fields:['cardholderName','cardType','cardNumber','linkedAccount','atmDebitNo','validTill']},
        {key:'investments', title:'Investment Account Details', cols:['Type','Platform','Holder','Advisor','Account No.','Nominee'], fields:['type','platform','holderName','advisor','accountNo','nominee']},
        {key:'properties', title:'Property Details', cols:['Property Name','Area','Owners','Registration No.','Nominee','Papers Location'], fields:['propertyName','area','owners','registrationNo','nominee','papersLocation']}
      ];

      var gpAnySectionWritten=false;
      gpSections.forEach(function(sec){
        var rows=(gpData[sec.key]||[]).filter(function(r){
          return sec.fields.some(function(f){ return r[f] && String(r[f]).trim(); });
        });
        if(!rows.length) return; // entirely empty section — skip, never render blank rows
        gpAnySectionWritten=true;
        gpBody.appendParagraph(sec.title).setHeading(DocumentApp.ParagraphHeading.HEADING2);
        var gpTable=gpBody.appendTable();
        var gpHeaderRow=gpTable.appendTableRow();
        sec.cols.forEach(function(c){ gpHeaderRow.appendTableCell(c); });
        rows.forEach(function(r){
          var gpDataRow=gpTable.appendTableRow();
          sec.fields.forEach(function(f){ gpDataRow.appendTableCell(String(r[f]||'')); });
        });
        gpBody.appendParagraph(''); // spacing
      });

      if(!gpAnySectionWritten){
        gpBody.appendParagraph('No information has been entered yet.');
      }
      gpDoc.saveAndClose();

      // Export as PDF, save into the client's Drive folder (creating it if
      // needed via the same function every other document feature uses),
      // then delete the temporary Doc.
      var gpFolder=getOrCreateClientFolder(gpRecord.name, gpRecord.pan||'', gpRecord.mobile);
      var gpPdfBlob=DriveApp.getFileById(gpDoc.getId()).getAs('application/pdf');
      gpPdfBlob.setName('Fin Doc - '+gpRecord.name);
      // Overwrite any previous Fin Doc PDF for this client, rather than
      // accumulating multiple versions each time they return and regenerate.
      var gpOldFiles=gpFolder.getFilesByName('Fin Doc - '+gpRecord.name);
      while(gpOldFiles.hasNext()){ gpOldFiles.next().setTrashed(true); }
      gpFolder.createFile(gpPdfBlob);
      DriveApp.getFileById(gpDoc.getId()).setTrashed(true); // remove the temp Doc

      if(gpRow>=0){
        gpSh.getRange(gpRow,6).setValue(Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss'));
      }
      return out({success:true});
    }catch(e){
      return out({success:false, error:e.message});
    }
  }

  // ═══════════════════════════════════════════════
  // CLIENT FEEDBACK — a separate FEEDBACK tab (not part of CLIENTS), created
  // automatically on first use. Supports genuine anonymity: if the client
  // chooses it, their name/mobile/email are never written to the row at
  // all — not just hidden in the UI, actually absent from the sheet. A
  // referral, if given, still creates a real lead with "Referred By" set to
  // the actual client name (known via the token) regardless of whether the
  // feedback itself was anonymous — the referral only works if someone
  // knows who to thank for it.
  // ═══════════════════════════════════════════════

  if(data.action==='getFeedbackClientName'){
    try{
      var gfnRecord=resolveTokenInThisScope(data.token);
      if(!gfnRecord || gfnRecord.recordType!=='client') return out({success:false, error:'Invalid or expired link'});
      return out({success:true, name:gfnRecord.name});
    }catch(e){
      return out({success:false, error:e.message});
    }
  }

  if(data.action==='submitFeedback'){
    try{
      var sfRecord=resolveTokenInThisScope(data.token);
      if(!sfRecord || sfRecord.recordType!=='client') return out({success:false, error:'Invalid or expired link'});

      var sfSh=ss.getSheetByName('FEEDBACK')||ss.insertSheet('FEEDBACK');
      if(sfSh.getLastRow()===0){
        var sfHdr=[['Timestamp','Client Name','Mobile','Email','Query Resolution Rating','Transaction Ease Rating','RM Support Rating','Communication Rating','Product Transparency Rating','What They Value Most','Area To Improve','Recommend Score (0-10)','Referral Name','Referral Mobile','Is Anonymous']];
        var sfHr=sfSh.getRange(1,1,1,15);
        sfHr.setValues(sfHdr);
        sfHr.setFontWeight('bold').setBackground('#C0392B').setFontColor('#FFFFFF');
        sfSh.setFrozenRows(1);
      }
      var sfAnon=!!data.isAnonymous;
      var sfRow=[
        Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss'),
        sfAnon ? 'Anonymous' : sfRecord.name,
        sfAnon ? '' : sfRecord.mobile,
        sfAnon ? '' : (data.email||''),
        data.ratingQueryResolution||'', data.ratingTransactionEase||'', data.ratingRmSupport||'',
        data.ratingCommunication||'', data.ratingProductTransparency||'',
        data.valueMost||'', data.areaToImprove||'', data.npsScore||'',
        data.referralName||'', data.referralMobile||'',
        sfAnon ? 'YES' : 'NO'
      ];
      sfSh.appendRow(sfRow);
      var sfLr=sfSh.getLastRow();
      sfSh.getRange(sfLr,1,1,15).setNumberFormat('@');

      // Referral handling — creates a real lead if a name was given. Uses
      // the REAL client name for "Referred By", never "Anonymous", since a
      // referral with no identifiable source is useless to the advisor.
      var referralCreated=false;
      var refName=String(data.referralName||'').trim();
      if(refName){
        var rlSh=ss.getSheetByName(LEADS_TAB)||ss.insertSheet(LEADS_TAB);
        if(rlSh.getLastRow()===0){
          var rlHdr=[['Name','Mobile','Email','Age','Country','Status','Source','Referred By','Potential SIP','Next Follow-up','Investment Products','Mode of Contact','Remarks','Modified Date','Created Date','Geo Tag','Potential AUM','Record ID','Family Name','Folder ID','Upload Token','User ID']];
          var rlHr=rlSh.getRange(1,1,1,22);rlHr.setValues(rlHdr);rlHr.setFontWeight('bold').setBackground('#C0392B').setFontColor('#FFFFFF');rlSh.setFrozenRows(1);
        }
        var refMobile=String(data.referralMobile||'').trim();
        var newLeadId=nextSequentialId(rlSh,18,'L'); // column R = Record ID
        var refRow=[refName, refMobile, '', '', '', 'PROSPECT', 'Referral', sfRecord.name, '', '', '', '', 'Referred by '+sfRecord.name+' via feedback form', '', Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'), '', '', newLeadId, '', '', '', ''];
        rlSh.appendRow(refRow);
        var rlLr=rlSh.getLastRow();
        rlSh.getRange(rlLr,1,1,22).setNumberFormat('@');
        referralCreated=true;
      }

      return out({success:true, referralCreated:referralCreated});
    }catch(e){
      return out({success:false, error:e.message});
    }
  }

  if(data.action==='generateContemporaryInsight'){
    try{ return out({success:true, reply:generateContemporaryInsight(data)}); }
    catch(e){ return out({success:false, error:e.message}); }
  }
  if(data.action==='practiceChat'){
    try{ return out({success:true, reply:practiceChat(data)}); }
    catch(e){ return out({success:false, error:e.message}); }
  }
  if(data.action==='generateTopBusinessInsights'){
    try{ return out({success:true, insights:generateTopBusinessInsights()}); }
    catch(e){ return out({success:false, error:e.message}); }
  }
  if(data.action==='runInsightsNow'){
    return out(generateAndSendInsights());
  }
  if(data.action==='setupInsightsTrigger'){
    return out(setupInsightsTrigger());
  }
  if(data.action==='removeInsightsTrigger'){
    return out(removeInsightsTrigger());
  }
  if(data.action==='getInsightsTriggerStatus'){
    return out(getInsightsTriggerStatus());
  }

  // ── Check whether a client's folder already exists — read-only ──
  // ── Gather Info: client-facing update of Health/Life Insurance status,
  // coverage amounts, premium due dates, App Installed, and Interested
  // Products — resolved by TOKEN, not by any client-supplied mobile/recordId,
  // since this endpoint has no authentication and must not let a caller
  // update an arbitrary row by guessing an ID. Products are MERGED with
  // whatever is already there (never overwritten), and the dependents'
  // insurance answer — which has no dedicated column — is appended to
  // Remarks as a clearly-labeled note, per explicit instruction. ──
  if(data.action==='updateClientGatherInfo'){
    try{
      var giTok=String(data.token||'').trim();
      if(!giTok) return out({success:false, error:'Missing token'});
      var giSh=ss.getSheetByName(CLIENTS_TAB);
      if(!giSh) return out({success:false, error:'No CLIENTS tab'});
      var giLr=giSh.getLastRow();
      if(giLr>=2){
        var giVals=giSh.getRange(2,1,giLr-1,41).getValues();
        for(var gi=0;gi<giVals.length;gi++){
          if(String(giVals[gi][32]||'').trim()===giTok){ // column AG = Upload Token
            var giRow=gi+2;
            // Batched into one read (already have giVals[gi] from the bulk
            // read above) + one write at the end, instead of up to 9 separate
            // setValue calls — same fix applied to updateTask earlier.
            var giRowData=giVals[gi].slice(); // copy — mutate this, write once
            if(data.healthInsurance!==undefined)   giRowData[21]=data.healthInsurance||''; // V
            if(data.lifeInsurance!==undefined)     giRowData[22]=data.lifeInsurance||''; // W
            if(data.dueDateHealth!==undefined)     giRowData[24]=data.dueDateHealth||''; // Y
            if(data.dueDateLife!==undefined)       giRowData[25]=data.dueDateLife||''; // Z
            if(data.appInstalled!==undefined)      giRowData[18]=data.appInstalled?'YES':'NO'; // S
            if(data.healthInsuranceCoverage!==undefined) giRowData[39]=data.healthInsuranceCoverage||''; // AN
            if(data.lifeInsuranceCoverage!==undefined)   giRowData[40]=data.lifeInsuranceCoverage||''; // AO

            // Products — MERGE new selections into whatever's already there, never overwrite
            if(data.products && data.products.length){
              var existingProd=String(giVals[gi][29]||'').split(',').map(function(p){return p.trim().toUpperCase();}).filter(Boolean);
              data.products.forEach(function(p){
                var pUp=String(p||'').trim().toUpperCase();
                if(pUp && existingProd.indexOf(pUp)<0) existingProd.push(pUp);
              });
              giRowData[29]=existingProd.join(', '); // AD
            }

            // Dependents' insurance — no dedicated column, appended to Remarks as a labeled note
            if(data.dependentsInsurance){
              var stamp=Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd MMM yyyy');
              var noteLine='['+stamp+'] Dependents (parents/siblings) health insurance: '+data.dependentsInsurance
                +(data.dependentsCoverage?' (Coverage: '+data.dependentsCoverage+')':'');
              var existingRemarks=String(giVals[gi][9]||'').trim(); // J
              var newRemarks=existingRemarks ? (existingRemarks+'\n'+noteLine) : noteLine;
              giRowData[9]=newRemarks; // J
            }

            giSh.getRange(giRow,1,1,41).setValues([giRowData]); // single batched write

            return out({success:true});
          }
        }
      }
      return out({success:false, error:'Link not found or expired'});
    }catch(e){
      return out({success:false, error:e.message});
    }
  }

  if(data.action==='checkClientFolder'){
    try{
      var existingFolder=findClientFolder(data.clientName, data.pan, data.mobile);
      if(existingFolder){
        return out({exists:true, folderId:existingFolder.getId(), folderUrl:existingFolder.getUrl()});
      }
      return out({exists:false});
    }catch(e){
      return out({exists:false, error:e.message});
    }
  }

  // ── Interactive search: EXACT match plus every name-only "ambiguous" candidate.
  // Frontend uses this to power a human "create new vs rename" decision. ──
  if(data.action==='searchClientFolderCandidates'){
    try{
      var parentS=getParentDocsFolder();
      var nameUpS=(data.clientName||'').trim().toUpperCase();
      var newFormatUpS=nameUpS+' +'+(data.mobile||'').trim();
      var exactS=null, ambigS=[];
      var iterS=parentS.getFolders();
      while(iterS.hasNext()){
        var fS=iterS.next();
        var fUpS=fS.getName().trim().toUpperCase();
        if(fUpS===newFormatUpS){ exactS={id:fS.getId(), name:fS.getName(), url:fS.getUrl()}; }
        else if(nameUpS && fUpS.indexOf(nameUpS)>=0){ ambigS.push({id:fS.getId(), name:fS.getName(), url:fS.getUrl()}); }
      }
      return out({success:true, exact:exactS, ambiguous:ambigS});
    }catch(e){
      return out({success:false, error:e.message});
    }
  }

  // ── Permanently tag a folder to a CLIENT — writes into column AI (35th),
  // matched by Client ID (column AK) — never mobile alone. ──
  if(data.action==='tagClientFolder'){
    try{
      var cSh=ss.getSheetByName(CLIENTS_TAB);
      if(!cSh) return out({success:false, error:'No CLIENTS tab'});
      var cLr=cSh.getLastRow();
      if(cLr>=2){
        var cRecId=String(data.recordId||'').trim();
        var cFound=-1;
        if(cRecId){
          var cIds=cSh.getRange(2,37,cLr-1,1).getValues(); // column AK
          for(var cix=0;cix<cIds.length;cix++){ if(String(cIds[cix][0]||'').trim()===cRecId){ cFound=cix; break; } }
        }
        if(cFound<0){
          var cMobiles=cSh.getRange(2,2,cLr-1,1).getValues(); // column B = mobile — legacy fallback only
          for(var ci=0;ci<cMobiles.length;ci++){ if(String(cMobiles[ci][0]||'').trim()===String(data.mobile||'').trim()){ cFound=ci; break; } }
        }
        if(cFound>=0){
          cSh.getRange(cFound+2,35).setValue(data.folderId||''); // column AI = 35th column
          return out({success:true});
        }
      }
      return out({success:false, error:'Client not found in sheet'});
    }catch(e){
      return out({success:false, error:e.message});
    }
  }

  // ── Rename an existing folder to the current standard "Name +Mobile" format ──
  if(data.action==='renameClientFolder'){
    try{
      var folderR=DriveApp.getFolderById(data.folderId);
      var newNameR=clientFolderName(data.clientName, '', data.mobile);
      folderR.setName(newNameR);
      return out({success:true, folderId:folderR.getId(), folderUrl:folderR.getUrl(), newName:newNameR});
    }catch(e){
      return out({success:false, error:e.message});
    }
  }

  // ── Bulk-check folders for a whole list of clients in ONE call ──
  if(data.action==='checkAllClientFolders'){
    try{
      var parentB=getParentDocsFolder();
      var allF=[];
      var fIter=parentB.getFolders();
      while(fIter.hasNext()){
        var ff=fIter.next();
        allF.push({nameUp:ff.getName().trim().toUpperCase(), id:ff.getId(), name:ff.getName(), url:ff.getUrl()});
      }

      var missing=[], ambiguous=[];
      (data.clients||[]).forEach(function(c){
        var nameUp=(c.name||'').trim().toUpperCase();
        var newFormatUp=nameUp+' +'+(c.mobile||'').trim();
        var exactMatch=allF.some(function(f){return f.nameUp===newFormatUp;});
        if(exactMatch) return; // fine — nothing to do for this client
        var candidates=allF.filter(function(f){return f.nameUp.indexOf(nameUp)===0;})
          .map(function(f){return {id:f.id, name:f.name, url:f.url};});
        if(candidates.length>0){
          ambiguous.push({name:c.name||'', mobile:c.mobile||'', recordId:c.recordId||'', candidates:candidates});
        } else {
          missing.push({name:c.name||'', mobile:c.mobile||'', recordId:c.recordId||''});
        }
      });
      return out({success:true, missing:missing, ambiguous:ambiguous, totalChecked:(data.clients||[]).length, totalFoldersFound:allF.length});
    }catch(e){
      return out({success:false, error:e.message});
    }
  }

  // ── Bulk-create folders for a list of clients ──
  if(data.action==='bulkCreateClientFolders'){
    try{
      var createdList=[];
      (data.clients||[]).forEach(function(c){
        var cf=getOrCreateClientFolder(c.name||'', c.pan||'', c.mobile||'');
        createdList.push({name:c.name||'', mobile:c.mobile||'', folderUrl:cf.getUrl()});
      });
      return out({success:true, created:createdList});
    }catch(e){
      return out({success:false, error:e.message});
    }
  }

  // ── Create a client's Drive folder explicitly ──
  if(data.action==='createClientFolder'){
    try{
      var cFolder;
      if(data.force){
        var parentF=getParentDocsFolder();
        cFolder=parentF.createFolder(clientFolderName(data.clientName, data.pan, data.mobile));
      } else {
        cFolder=getOrCreateClientFolder(data.clientName, data.pan, data.mobile);
      }
      return out({success:true, folderId:cFolder.getId(), folderUrl:cFolder.getUrl()});
    }catch(e){
      return out({success:false, error:e.message});
    }
  }

  // ── Upload a document into a client's (or lead's) folder ──
  if(data.action==='uploadClientDocument'){
    try{
      var uFolder=null;
      if(data.folderId){
        try{ uFolder=DriveApp.getFolderById(data.folderId); }catch(fe){ uFolder=null; }
      }
      if(!uFolder){
        uFolder=getOrCreateClientFolder(data.clientName, data.pan, data.mobile);
      }
      var existingFiles=uFolder.getFiles();
      while(existingFiles.hasNext()){
        var ef=existingFiles.next();
        if(ef.getName().indexOf(data.docType+'.')===0 || ef.getName()===data.docType){
          ef.setTrashed(true);
        }
      }
      var bytes=Utilities.base64Decode(data.fileData);
      var blob=Utilities.newBlob(bytes, data.mimeType||'application/octet-stream', data.docType+(data.fileExt?'.'+data.fileExt:''));
      var newFile=uFolder.createFile(blob);
      return out({success:true, fileId:newFile.getId(), fileUrl:newFile.getUrl(), fileName:newFile.getName(), folderId:uFolder.getId()});
    }catch(e){
      return out({success:false, error:e.message});
    }
  }

  // ── Save nominee details as a plain text file in the client's folder ──
  if(data.action==='saveNomineeDetails'){
    try{
      var nFolder=null;
      if(data.folderId){ try{ nFolder=DriveApp.getFolderById(data.folderId); }catch(fe){ nFolder=null; } }
      if(!nFolder){ nFolder=getOrCreateClientFolder(data.clientName, data.pan, data.mobile); }
      var nExisting=nFolder.getFilesByName('Nominee Details.txt');
      while(nExisting.hasNext()){ nExisting.next().setTrashed(true); }
      var nContent='Nominee name: '+(data.nomineeName||'')+'\n'
        +'Nominee mobile: '+(data.nomineeMobile||'')+'\n'
        +'Nominee mail id: '+(data.nomineeEmail||'');
      var nFile=nFolder.createFile('Nominee Details.txt', nContent, MimeType.PLAIN_TEXT);
      return out({success:true, fileId:nFile.getId(), fileUrl:nFile.getUrl(), folderId:nFolder.getId()});
    }catch(e){
      return out({success:false, error:e.message});
    }
  }

  // ── Save bank account details as a plain text file in the client's folder ──
  if(data.action==='saveBankDetails'){
    try{
      var bFolder=null;
      if(data.folderId){ try{ bFolder=DriveApp.getFolderById(data.folderId); }catch(fe){ bFolder=null; } }
      if(!bFolder){ bFolder=getOrCreateClientFolder(data.clientName, data.pan, data.mobile); }
      var bExisting=bFolder.getFilesByName('Bank Details.txt');
      while(bExisting.hasNext()){ bExisting.next().setTrashed(true); }
      var bContent='Account Number: '+(data.accountNumber||'')+'\n'
        +'IFSC Code: '+(data.ifscCode||'');
      var bFile=bFolder.createFile('Bank Details.txt', bContent, MimeType.PLAIN_TEXT);
      return out({success:true, fileId:bFile.getId(), fileUrl:bFile.getUrl(), folderId:bFolder.getId()});
    }catch(e){
      return out({success:false, error:e.message});
    }
  }

  // ── List all documents currently uploaded for a client (or lead) ──
  if(data.action==='listClientDocuments'){
    try{
      var lFolder=null;
      if(data.folderId){ try{ lFolder=DriveApp.getFolderById(data.folderId); }catch(fe){ lFolder=null; } }
      if(!lFolder){
        // Read-only lookup only — do NOT create a folder just from checking/viewing.
        // Folder creation must only ever happen as a side effect of an actual upload.
        var foundFolder=findClientFolder(data.clientName, data.pan, data.mobile);
        if(!foundFolder) return out({success:true, files:[], folderUrl:'', folderId:''});
        lFolder=foundFolder;
      }
      var files=lFolder.getFiles();
      var fileList=[];
      while(files.hasNext()){
        var f=files.next();
        var fname=f.getName();
        var docType=fname.indexOf('.')>=0?fname.substring(0,fname.lastIndexOf('.')):fname;
        fileList.push({docType:docType, name:fname, url:f.getUrl(), id:f.getId()});
      }
      return out({success:true, files:fileList, folderUrl:lFolder.getUrl(), folderId:lFolder.getId()});
    }catch(e){
      return out({success:false, error:e.message, files:[]});
    }
  }

  // ── Delete a single client document ──
  if(data.action==='deleteClientDocument'){
    try{
      var dFile=DriveApp.getFileById(data.fileId);
      dFile.setTrashed(true);
      return out({success:true});
    }catch(e){
      return out({success:false, error:e.message});
    }
  }

  if(data.action==='readSpecialDays'){var spSh=ss.getSheetByName(SPECIAL_TAB);if(!spSh)return out({rows:[]});var spLr=spSh.getLastRow();if(spLr<2)return out({rows:[]});var spVals=spSh.getRange(2,1,spLr-1,2).getValues();var spRows=spVals.map(function(r){return[String(r[0]||'').trim(),String(r[1]||'').trim()];}).filter(function(r){return r[0]&&r[1];});return out({rows:spRows});}
  if(data.action==='pushSpecialDays'){
    var spSh3=ss.getSheetByName(SPECIAL_TAB)||ss.insertSheet(SPECIAL_TAB);
    var spDataRows=data.rows||[];
    var spLastRow=spSh3.getLastRow();
    if(spLastRow===0){
      spSh3.appendRow(['Date (MM-DD)','Remarks']);
      var spHdr=spSh3.getRange(1,1,1,2);
      spHdr.setFontWeight('bold').setBackground('#C0392B').setFontColor('#FFFFFF');
      spSh3.setFrozenRows(1);
    }
    var spExisting=spLastRow>1?spSh3.getRange(2,1,spLastRow-1,1).getValues().map(function(r){return String(r[0]||'').trim();}):[]; 
    spDataRows.forEach(function(row){
      if(row[0]==='Date (MM-DD)'||row[0]==='Day') return;
      if(spExisting.indexOf(String(row[0]||'').trim())>=0) return;
      spSh3.appendRow([row[0],row[1]]);
      spSh3.getRange(spSh3.getLastRow(),1,1,2).setNumberFormat('@');
    });
    return out({success:true,action:'pushSpecialDays'});
  }
  if(data.action==='appendSpecialDay'){var spSh2=ss.getSheetByName(SPECIAL_TAB)||ss.insertSheet(SPECIAL_TAB);if(spSh2.getLastRow()===0){var spH=spSh2.getRange(1,1,1,2);spH.setValues([['Date (MM-DD)','Remarks']]);spH.setFontWeight('bold').setBackground('#C0392B').setFontColor('#FFFFFF');spSh2.setFrozenRows(1);}spSh2.appendRow([data.date||'',data.speciality||'']);return out({success:true,action:'appendSpecialDay'});}
var tab=data.action==='pushLeads'?LEADS_TAB:data.action==='pushClients'?CLIENTS_TAB:null;if(!tab)return out({error:'Bad action:'+data.action});var rows=data.rows;if(!rows||rows.length===0)return out({error:'No rows'});var sh=ss.getSheetByName(tab)||ss.insertSheet(tab);sh.clearContents();var r=sh.getRange(1,1,rows.length,rows[0].length);r.setNumberFormat('@');r.setValues(rows);var h=sh.getRange(1,1,1,rows[0].length);h.setFontWeight('bold').setBackground('#C0392B').setFontColor('#FFFFFF');sh.setFrozenRows(1);return out({success:true,written:rows.length-1});}
function out(o){return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);}
function testConnection(){Logger.log('OK: '+SpreadsheetApp.openById(SHEET_ID).getName());}

// ═══════════════════════════════════════════
// AUTOMATED BACKUP SYSTEM
// ═══════════════════════════════════════════
const BACKUP_FOLDER_NAME='Wealth Matrix CRM Backups';
const MAX_BACKUPS_TO_KEEP=8; // ~2 months of weekly backups

function createBackupTrigger(){
  var triggers=ScriptApp.getProjectTriggers();
  triggers.forEach(function(t){
    if(t.getHandlerFunction()==='runWeeklyBackup') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('runWeeklyBackup')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.SUNDAY)
    .atHour(2)
    .create();
  Logger.log('✅ Weekly backup trigger created. Backups will run automatically every Sunday at 2 AM.');
}

function runWeeklyBackup(){
  try{
    var srcFile=DriveApp.getFileById(SHEET_ID);
    var folders=DriveApp.getFoldersByName(BACKUP_FOLDER_NAME);
    var folder=folders.hasNext()?folders.next():DriveApp.createFolder(BACKUP_FOLDER_NAME);

    var dateStr=Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'yyyy-MM-dd_HH-mm');
    var backupName='CRM Backup '+dateStr;
    srcFile.makeCopy(backupName, folder);
    Logger.log('✅ Backup created: '+backupName);

    var files=folder.getFiles();
    var fileList=[];
    while(files.hasNext()){
      var f=files.next();
      fileList.push({file:f, date:f.getDateCreated()});
    }
    fileList.sort(function(a,b){return b.date-a.date;});
    if(fileList.length>MAX_BACKUPS_TO_KEEP){
      for(var i=MAX_BACKUPS_TO_KEEP;i<fileList.length;i++){
        fileList[i].file.setTrashed(true);
      }
      Logger.log('🗑️ Cleaned up '+(fileList.length-MAX_BACKUPS_TO_KEEP)+' old backup(s)');
    }
  }catch(e){
    Logger.log('❌ Backup failed: '+e.message);
  }
}

function manualBackupNow(){
  runWeeklyBackup();
  return out({success:true,message:'Backup created'});
}
