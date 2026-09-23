#!/usr/bin/env node
// ═══════════════════════════════════════════════════════════════════════
// dp-build.js — Deep Pocket CRM white-label builder (runs inside Claude's
// sandbox, or on any PC with Node 18+). Deterministic, count-verified
// rebrand of the master files. Claude only does the judgment work on top.
//
//   node dp-build.js code             → out/Code.gs
//   node dp-build.js app <EXEC_URL>   → out/config.js, out/crm.html + 6 client pages
//   node dp-build.js check            → dry run on every rule (no files written)
//   node dp-build.js verify           → final checks on out/ + the deploy-zip command
//
// Reads (same folder): firm.json, crm.html, Code.gs (or crm_Code.gs),
// upload.html, gatherinfo.html, review.html, findoc.html, feedback.html,
// privacy-policy.html, icon-192.png, icon-512.png (optional).
// Prints a SHORT summary only; full detail goes to out/BUILD-REPORT-<stage>.md
// If a rule shows "mismatch", the master changed: check that anchor.
// ═══════════════════════════════════════════════════════════════════════
'use strict';
const fs=require('fs'), path=require('path');
const DIR=__dirname, OUT=path.join(DIR,'out');
const BUILDER_VERSION='4.0.0';
const CLIENT_PAGES=['upload.html','gatherinfo.html','review.html','findoc.html','feedback.html','privacy-policy.html'];
const RESIDUE_CHECKS=['Arjun','9745014475','Alappat','Ernakulam','wealthmatrix','wealthmetrics','wealth-matrix','Deep Pocket Fintech Pvt. Ltd.','Partner, Deep Pocket'];
const EXEC_RE=/^https:\/\/script\.google\.com\/macros\/s\/[\w-]+\/exec$/;
const PRODUCT_LABELS={'INVESTMENT':'Mutual Funds','LIFE INSURANCE':'Term Life Insurance','HEALTH INSURANCE':'Health Insurance','WILL':'Will Writing','AIF':'AIF','PMS':'PMS'};
const PAGE_FEATURES=['documents','gatherinfo','findoc','feedback','reviewreport','appinstalled','casanalysis'];

const MASTER_REF = {
  firmLegal: 'Deep Pocket Fintech Pvt. Ltd.',
  firmBrand: 'Deep Pocket Fintech',
  shortBrand: 'Deep Pocket',
  productName: 'Deep Pocket CRM',
  ownerFull: 'Arjun K T M',
  ownerFirst: 'Arjun',
  phone: '9745014475',
  website: 'www.wealthmatrix.in',
  storagePrefix: 'wm_',
  stages: ['KYC PENDING','DOCUMENT PENDING','PROFILE CREATION PENDING','MANDATE PENDING','INVESTMENT PENDING','GRATITUDE CALL PENDING'],
  products: ['INVESTMENT','LIFE INSURANCE','HEALTH INSURANCE','WILL','AIF','PMS'],
  theme: { bg:'#1B1512', s1:'#2A211D', s2:'#332924', s3:'#3C312B', border:'#3A2F2A', accent:'#F2E8DE', coral:'#E8734A', text:'#F2E8DE', t2:'#A99A8F', t3:'#786A60', green:'#4C9A6A', blue:'#4A8FC7', orange:'#C08A3E', red:'#C1553D' }
};

function jsSq_(s){ return String(s==null?'':s).replace(/\\/g,'\\\\').replace(/'/g,"\\'").replace(/\r?\n/g,'\\n'); }
function jsDq_(s){ return JSON.stringify(String(s==null?'':s)).slice(1,-1); }
function htmlEsc_(s){ return String(s==null?'':s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;'); }
function trim_(s){ return String(s==null?'':s).trim(); }
function slug_(s){ return String(s||'').toLowerCase().replace(/[^a-z0-9]+/g,'-').replace(/^-+|-+$/g,''); }
function hash4_(s){ var h=0; s=String(s||''); for(var i=0;i<s.length;i++){ h=((h<<5)-h+s.charCodeAt(i))|0; } return ('000'+(Math.abs(h)%46656).toString(36)).slice(-3); }
function hexToRgb_(hex){ var m=String(hex||'').replace('#','').match(/^([0-9a-f]{6})$/i); if(!m) return null; var n=parseInt(m[1],16); return [(n>>16)&255,(n>>8)&255,n&255]; }
function rgba_(hex,a){ var c=hexToRgb_(hex); return c?'rgba('+c[0]+','+c[1]+','+c[2]+','+a+')':null; }
function isHex_(h){ return /^#[0-9a-f]{6}$/i.test(String(h||'')); }
function intList_(s, lo, hi){
  var seen={}, outL=[];
  String(s||'').split(/[^0-9]+/).forEach(function(x){ var n=parseInt(x,10); if(n>=lo && n<=hi && !seen[n]){ seen[n]=1; outL.push(n); } });
  return outL.sort(function(a,b){ return b-a; });
}
function intOr_(s, lo, hi){ var n=parseInt(String(s||'').trim(),10); return (n>=lo && n<=hi) ? n : null; }

function stripLegalSuffix_(name){
  return String(name||'').replace(/[,\s]*(private\s+limited|pvt\.?\s*ltd\.?|limited|ltd\.?|llp|l\.l\.p\.?|opc)\s*$/i,'').replace(/[,\s.]+$/,'').trim();
}
function extractSheetId_(s){
  s=trim_(s); var m=s.match(/\/d\/([a-zA-Z0-9_-]{20,})/); if(m) return m[1];
  return /^[a-zA-Z0-9_-]{20,}$/.test(s) ? s : '';
}
function splitAddress_(addr){
  var a=trim_(addr); if(!a) return [];
  if(/\n/.test(a)) return a.split(/\r?\n/).map(trim_).filter(Boolean).slice(0,4);
  var parts=a.split(/\s*,\s*/).filter(Boolean);
  if(parts.length<=2) return [parts.join(', ')];
  var mid=Math.ceil(parts.length/2);
  return [parts.slice(0,mid).join(', '), parts.slice(mid).join(', ')];
}

// All identity values with fallbacks, computed once and reused by rules,
// prompts and the report.
function resolveValues_(intake){
  var id=intake.identity||{}, cm=intake.comms||{}, tc=intake.technical||{}, bz=intake.business||{};
  var legal=trim_(id.firmName)||'[FIRM NAME]';
  var brand=trim_(id.brandName)||stripLegalSuffix_(legal)||legal;
  var product=trim_(id.productName)||(brand+' CRM');
  var short=trim_(id.shortName)||(brand.length<=12?brand:brand.split(/\s+/)[0]).slice(0,12);
  var owner=trim_(id.ownerName)||'[OWNER NAME]';
  var signOff=trim_(id.signOffName)||owner.split(/\s+/)[0];
  var website=trim_(cm.websiteLink).replace(/^https?:\/\//i,'').replace(/\/+$/,'');
  var ns=(slug_(short).replace(/-/g,'').slice(0,10)||'firm')+hash4_(legal)+'_';
  return {
    legal:legal, brand:brand, product:product, short:short, owner:owner, signOff:signOff,
    title:trim_(id.ownerTitle), address:trim_(id.address), city:trim_(id.city),
    phone:trim_(id.phone), welcomePhone:trim_(id.welcomePhone)||trim_(id.phone), email:trim_(id.email),
    insightsEmail:trim_(id.insightsEmail)||trim_(id.email),
    website:website, hasOwnApp:cm.hasOwnApp!==false,
    android:trim_(cm.playStoreLink), ios:trim_(cm.appStoreLink),
    sheetId:extractSheetId_(tc.sheetId), calendarId:trim_(tc.calendarId),
    arn:trim_(bz.arn).toUpperCase(), ns:ns, deploySlug:slug_(brand)||'firm',
    scriptUrl: /^https:\/\/script\.google\.com\/macros\/s\/[\w-]+\/exec$/.test(trim_(tc.appsScriptUrl)) ? trim_(tc.appsScriptUrl) : ''
  };
}


function buildRebrandPlan_(intake, icons, masters){
  var v=resolveValues_(intake);
  var cm=intake.comms||{}, tc=intake.technical||{}, br=intake.brand||{};
  var R=MASTER_REF, H=[], G=[];
  function add(list, id, label, find, replace, expect, extra){ var r={id:id,label:label,find:find,replace:replace,expect:expect}; if(extra) for(var k in extra) r[k]=extra[k]; list.push(r); }

  // ---------- crm.html : messages & identity ----------
  var customWelcome=trim_(cm.customWelcome);
  if(customWelcome){
    var cw=customWelcome.replace(/\{name\}/g,'\u0002');
    var cwParts=cw.split('\u0002').map(function(p){ return "'"+jsSq_(p)+"'"; });
    var cwExpr=cwParts.join("+(name||'')+");
    add(H,'welcome-custom','Welcome message → firm\'s exact wording',
      /if\(type==='welcome'\)\{\s*return 'Dear '\+\(name\|\|''\)\+','[\s\S]*?\+'\\nMG Road, Ernakulam';\s*\}/,
      "if(type==='welcome'){\n    return "+cwExpr+";\n  }", 1);
  }
  var prospect=trim_(cm.prospectIntro) || ('Hi {name},\nHope you are doing great.\nThis is '+v.signOff+' from '+v.legal+(v.city?', '+v.city:'')+'. I would be glad to help you with your financial planning whenever convenient.\nRegards\n'+v.signOff);
  add(H,'defmsg-prospect','Default PROSPECT WhatsApp message',
    "PROSPECT:'Hi {name},\\nHope you are doing great.\\nI have started my own Financial Advisory / Wealth Management Firm - Deep Pocket Fintech Pvt. Ltd.. Office at Ernakulam.\\nRegards\\nArjun',",
    "PROSPECT:'"+jsSq_(prospect)+"',", 1);
  add(H,'signoff-regards','Default status messages sign-off ("Regards Arjun")', "\\nRegards\\nArjun'", "\\nRegards\\n"+jsSq_(v.signOff)+"'", 3);
  add(H,'signoff-thisis','Client call-intro WhatsApp', 'this is Arjun from Deep Pocket Fintech.', 'this is '+jsSq_(v.signOff)+' from '+jsSq_(v.brand)+'.', 1);
  add(H,'signoff-review','Review-report WhatsApp sign-off', "\\n\\nWarm regards\\nArjun';", "\\n\\nWarm regards\\n"+jsSq_(v.signOff)+"';", 1);
  add(H,'signoff-sipreturn','SIP-returned WhatsApp sign-off', "\\nThanks\\nArjun';", "\\nThanks\\n"+jsSq_(v.signOff)+"';", 1);
  add(H,'owner-full','Owner full name in templates', R.ownerFull, jsSq_(v.owner), customWelcome?13:14);
  add(H,'firm-legal','Firm legal name in templates', R.firmLegal, jsSq_(v.legal), 8);
  if(!customWelcome) add(H,'welcome-to','Welcome message greeting', 'Welcome to Deep Pocket Fintech — we', 'Welcome to '+jsSq_(v.brand)+' — we', 1);
  add(H,'mail-subject-welcome','Welcome email subject', "'Welcome to Deep Pocket Fintech'", "'Welcome to "+jsSq_(v.brand)+"'", 1);
  add(H,'mail-subject-default','Default client email subject', ": 'Deep Pocket Fintech';", ": '"+jsSq_(v.brand)+"';", 1);
  add(H,'mail-subject-mailto','Mail-button email subject', "encodeURIComponent('Deep Pocket Fintech')", "encodeURIComponent('"+jsSq_(v.brand)+"')", 1);
  if(!customWelcome){
    add(H,'owner-title','Owner title line in welcome', "+'\\nPartner, Deep Pocket'",
      v.title ? "+'\\n"+jsSq_(v.title+', '+v.brand)+"'" : "+''", 1);
    var addr=splitAddress_(v.address);
    add(H,'address','Office address in welcome',
      /\+'\\nMG Road, Ernakulam';/, addr.length ? addr.slice(1).map(function(l){ return "+'\\n"+jsSq_(l)+"'"; }).join('')+";" : ";", 1,
      {pre:{find:"+'\\n8th Floor, Alappat Heritage'", replace: addr.length ? "+'\\n"+jsSq_(addr[0])+"'" : "+''"}});
    add(H,'phone','Office phone in welcome', 'Ph: '+R.phone+'.', 'Ph: '+jsSq_(v.welcomePhone)+'.', 1);
    var sentence;
    if(v.hasOwnApp && v.website) sentence='You can track your investments anytime through our mobile app, or visit '+v.website+' to access your account online.';
    else if(v.hasOwnApp) sentence='You can track your investments anytime through our mobile app.';
    else if(v.website) sentence='You can visit '+v.website+' anytime to access your account online.';
    else sentence='';
    add(H,'website','Website / app sentence in welcome',
      "+'\\n\\nYou can track your investments anytime through our mobile app, or visit www.wealthmatrix.in to access your account online.'",
      sentence ? "+'\\n\\n"+jsSq_(sentence)+"'" : "+''", 1);
  }
  add(H,'app-android','Android app link (APP_LINK_ANDROID)', /var APP_LINK_ANDROID = '[^']*';/, "var APP_LINK_ANDROID = '"+jsSq_(v.hasOwnApp?v.android:'')+"';", 1);
  add(H,'app-ios','iOS app link (APP_LINK_IOS)', /var APP_LINK_IOS = '[^']*';/, "var APP_LINK_IOS = '"+jsSq_(v.hasOwnApp?v.ios:'')+"';", 1);

  // ---------- crm.html : product / app identity ----------
  add(H,'pwa-manifest','PWA manifest name + short_name', 'name:"Deep Pocket CRM",short_name:"Deep Pocket"', 'name:"'+jsDq_(v.product)+'",short_name:"'+jsDq_(v.short)+'"', 1);
  add(H,'ios-title','iOS home-screen title', 'name="apple-mobile-web-app-title" content="Deep Pocket">', 'name="apple-mobile-web-app-title" content="'+htmlEsc_(v.short)+'">', 1);
  add(H,'doc-title','Browser tab title', '<title>Deep Pocket CRM V ', '<title>'+htmlEsc_(v.product)+' V ', 1);
  add(H,'product-labels','Product name on login, lock, top bar, menu', '>Deep Pocket CRM<', '>'+htmlEsc_(v.product)+'<', 4);
  add(H,'settings-version','Settings footer version label', 'id="settings-app-version">Deep Pocket CRM V ', 'id="settings-app-version">'+htmlEsc_(v.product)+' V ', 1);
  add(H,'settings-version-js','Settings version label (JS)', "textContent='Deep Pocket CRM V '+APP_VERSION", "textContent='"+jsSq_(v.product)+" V '+APP_VERSION", 1);
  add(H,'sidebar-brand','Desktop sidebar brand', '>Deep Pocket</div><div style="font-size:9px;color:var(--t3);letter-spacing:.5px">', '>'+htmlEsc_(v.short)+'</div><div style="font-size:9px;color:var(--t3);letter-spacing:.5px">', 1);
  add(H,'logo-alt','Logo alt text', 'alt="Deep Pocket"', 'alt="'+htmlEsc_(v.brand)+'"', 4);
  add(H,'webauthn-rp','Biometric lock prompt name', "rp:{name:'Deep Pocket CRM'}", "rp:{name:'"+jsSq_(v.product)+"'}", 1);
  add(H,'webauthn-user','Biometric lock user label', "displayName:'Deep Pocket User'", "displayName:'"+jsSq_(v.short)+" User'", 1);
  add(H,'meet-title','Google Meet event title', "' - Deep Pocket');", "' - "+jsSq_(v.brand)+"');", 1);
  add(H,'backup-msg','Health-check backup folder message', '(see "Deep Pocket CRM Backups" folder in Drive)', '(see "'+jsSq_(v.legal)+' CRM Backup Folder" in Drive)', 1);

  // ---------- crm.html : Sheet link pre-filled in Settings ----------
  if(v.sheetId) add(H,'sheet-prefill','Settings → Google Sheet link pre-filled', 'id="sheet-url-input" placeholder=', 'id="sheet-url-input" value="https://docs.google.com/spreadsheets/d/'+htmlEsc_(v.sheetId)+'/edit" placeholder=', 1);

  // ---------- crm.html : master bug — white text on the cream --accent ----------
  H.push({id:'contrast-fix', label:'Readable text on accent-coloured buttons (master bug: white on cream)', expect:3, fn:fixAccentContrast_});

  // ---------- crm.html : per-firm storage namespace ----------
  add(H,'storage-ns','localStorage namespace (per-firm isolation on the shared dp-crm.github.io origin)', "'"+R.storagePrefix, "'"+v.ns, null);

  // ---------- crm.html : operational defaults ----------
  var fu=intOr_(tc.followUpDays,1,120);
  if(fu){ add(H,'followup-default','Lead follow-up window default', 'cfg_.followUpDays=14', 'cfg_.followUpDays='+fu, 1);
          add(H,'followup-uses','Lead follow-up window fallbacks', 'parseInt(cfg_.followUpDays)||14', 'parseInt(cfg_.followUpDays)||'+fu, 6); }
  var sipDays=intList_(tc.sipReminderDays,1,15);
  if(sipDays.length) add(H,'sip-reminder','SIP reminder days before due', 'd1days===5||d1days===2', sipDays.map(function(n){ return 'd1days==='+n; }).join('||'), 2);
  var bounce=intOr_(tc.sipBounceDays,1,25);
  if(bounce) add(H,'sip-bounce','SIP debit-verification delay', 'if(daysSince<7) return;', 'if(daysSince<'+bounce+') return;', 1);
  var gl=intOr_(tc.goalLeads,1,100000), gc=intOr_(tc.goalConversions,1,100000);
  if(gl){ add(H,'goal-leads-input','Monthly lead target (input)', 'id="s-goal-leads" type="number" style="width:70px;text-align:center" value="30"', 'id="s-goal-leads" type="number" style="width:70px;text-align:center" value="'+gl+'"', 1);
          add(H,'goal-leads-js','Monthly lead target (default)', "cfg_.goalLeads||'30'", "cfg_.goalLeads||'"+gl+"'", 1); }
  if(gc){ add(H,'goal-conv-input','Monthly conversion target (input)', 'id="s-goal-conv" type="number" style="width:70px;text-align:center" value="5"', 'id="s-goal-conv" type="number" style="width:70px;text-align:center" value="'+gc+'"', 1);
          add(H,'goal-conv-js','Monthly conversion target (default)', "cfg_.goalConv||'5'", "cfg_.goalConv||'"+gc+"'", 1); }
  var VOICE={'en-IN|English':1,'hi-IN|Hindi':1,'ml-IN|Malayalam':1,'ta-IN|Tamil':1,'te-IN|Telugu':1,'kn-IN|Kannada':1};
  if(cm.voiceLanguage && VOICE[cm.voiceLanguage] && cm.voiceLanguage!=='en-IN|English')
    add(H,'voice-default','Default voice-assistant language', "cfg_.voiceLanguage||'en-IN|English'", "cfg_.voiceLanguage||'"+cm.voiceLanguage+"'", 2);

  // ---------- crm.html : colours ----------
  var theme=br.theme||{}, changed={};
  Object.keys(R.theme).forEach(function(k){ if(isHex_(theme[k]) && theme[k].toUpperCase()!==R.theme[k].toUpperCase()) changed[k]=theme[k].toUpperCase(); });
  if(Object.keys(changed).length){
    H.push({id:'theme-root', label:'Theme colours in :root ('+Object.keys(changed).join(', ')+')', expect:null, fn:function(text){ return applyThemeToRoot_(text, changed); }});
    if(changed.bg){
      add(H,'theme-bg-outside','Background colour (:root, theme-color meta, manifest ×2)', R.theme.bg, changed.bg, 4);
      add(H,'nav-bg','Bottom nav background (was a hard-coded navy leftover)', 'rgba(12,20,32,.97)', rgba_(changed.bg,.97), 1);
    }
  }

  // ---------- crm.html : logo / icons ----------
  if(icons && icons.icon192){
    var m192=masters.html.match(/window\.DP_LOGO="(data:image\/png;base64,[^"]+)"/);
    var m512=masters.html.match(/window\.DP_LOGO_HD="(data:image\/png;base64,[^"]+)"/);
    if(m192) add(H,'logo-192','Logo 192px (favicon, apple-touch-icon, manifest, 4 in-app logos)', m192[1], 'data:image/png;base64,'+icons.icon192, 7);
    if(m512) add(H,'logo-512','Logo 512px (install icon)', m512[1], 'data:image/png;base64,'+(icons.icon512||icons.icon192), 1);
  }

  // ---------- Code.gs ----------
  add(G,'sheet-id','SHEET_ID', /const SHEET_ID='[^']*';/, "const SHEET_ID='"+(v.sheetId||'PASTE_YOUR_GOOGLE_SHEET_ID_HERE')+"';", 1);
  add(G,'review-author','Review-report author line (AI prompt)', 'Arjun K T M at Deep Pocket Fintech Pvt. Ltd..', jsSq_(v.owner)+' at '+jsSq_(/\.$/.test(v.legal)?v.legal:v.legal+'.'), 1);
  add(G,'owner-full','Owner name in review-report signature', R.ownerFull, jsSq_(v.owner), 1);
  add(G,'firm-legal','Firm legal name in review-report signature', R.firmLegal, jsSq_(v.legal), 1);
  add(G,'ai-advisor-at','Firm name inside AI system prompts', 'advisor at Deep Pocket Fintech', 'advisor at '+jsSq_(v.brand), 6);
  add(G,'docs-folder','Client documents parent Drive folder', "'Deep Pocket Client Documents'", "'"+jsSq_(v.legal)+" Client Documents'", 1);
  add(G,'backup-fallback','Backup folder fallback name', "(fn || 'Deep Pocket Fintech Pvt Ltd')+' CRM Backup Folder'", "(fn || '"+jsSq_(v.legal)+"')+' CRM Backup Folder'", 1);
  add(G,'firmname-default','Default firm name (client pages, insight mails, backups) — its Settings UI was removed', "getSettingValue('firmName','')", "getSettingValue('firmName','"+jsSq_(v.legal)+"')", 6);
  if(/^ARN-?\d+$/i.test(v.arn)) add(G,'arn-default','Default ARN code (skips first-use popup)', "getProperty('OUR_ARN_CODE')||''", "getProperty('OUR_ARN_CODE')||'"+jsSq_(v.arn)+"'", 1);
  
  if(icons && icons.icon192) add(G,'logo-default','Default firm logo for client pages (logo upload UI was removed)', "getSettingValue('logoDataUrl','')", "getSettingValue('logoDataUrl','data:image/png;base64,"+icons.icon192+"')", 3);

  // ---------- Client-facing pages: rewrite the FIRM SETTINGS block ----------
  var pages={}, block=dpConfigBlock_(intake, v, icons);
  Object.keys(masters.pages||{}).forEach(function(n){
    pages[n]=[{id:'dp-config', label:'FIRM SETTINGS block (DP_CONFIG)', expect:1, fn:function(text){
      var re=/\/\*DP_CONFIG_START\*\/[\s\S]*?\/\*DP_CONFIG_END\*\//g, c=(text.match(re)||[]).length;
      return {text:text.replace(re, function(){ return block; }), count:c};
    }}];
  });

  // Split "pre" rules (two-part address) into their own entries, before the main one.
  H=expandPreRules_(H);

  var masterSheetId=(masters.gs.match(/const SHEET_ID='([^']*)';/)||[])[1]||'';
  var masterLogoHead=((masters.html.match(/window\.DP_LOGO="data:image\/png;base64,([^"]{80,})"/)||[])[1]||'').slice(40,120);
  var masterScriptId='';
  
  return {
    html:H, gs:G, pages:pages, masterScriptId:masterScriptId, values:v, changedTheme:changed, masterSheetId:masterSheetId, masterLogoHead:(icons&&icons.icon192)?masterLogoHead:''
  };
}

function expandPreRules_(list){
  var outL=[];
  list.forEach(function(r){
    if(r.pre){ outL.push({id:r.id+'-line1', label:r.label+' (line 1)', find:r.pre.find, replace:r.pre.replace, expect:r.expect}); delete r.pre; }
    outL.push(r);
  });
  return outL;
}

function fixAccentContrast_(text){
  var count=0;
  text=text.replace(/(background:var\(--accent\);(?:border-color:var\(--accent\);)?color:)#fff\b|(color:)#fff(\s*!important)?(;background:var\(--accent\))/g,
    function(m,a,b,imp,c){ count++; return a ? a+'var(--bg)' : b+'var(--bg)'+(imp||'')+c; });
  return {text:text, count:count};
}

function applyThemeToRoot_(text, changed, includeBg){
  var start=text.indexOf(':root{'); if(start<0) return {text:text, count:0};
  var end=text.indexOf('}', start); if(end<0) return {text:text, count:0};
  var block=text.slice(start,end), count=0;
  function setVar(name, value){
    var re=new RegExp('(--'+name.replace(/[-]/g,'\\-')+':)[^;]*;');
    if(re.test(block)){ block=block.replace(re, function(_, p){ return p+value+';'; }); count++; }
  }
  Object.keys(changed).forEach(function(k){ if(k!=='bg' || includeBg) setVar(k, changed[k]); }); // crm.html: bg handled by a global rule
  if(changed.accent){ setVar('accent-light', changed.accent); setVar('gold', changed.accent); setVar('accent2', rgba_(changed.accent,.10)); setVar('gold2', rgba_(changed.accent,.10)); }
  if(changed.coral) setVar('coraldim', rgba_(changed.coral,.15));
  if(changed.green) setVar('gdim', rgba_(changed.green,.14));
  if(changed.blue) setVar('bdim', rgba_(changed.blue,.14));
  if(changed.orange) setVar('odim', rgba_(changed.orange,.14));
  if(changed.red) setVar('rdim', rgba_(changed.red,.15));
  if(changed.t3){ setVar('dropped', changed.t3); setVar('droppeddim', rgba_(changed.t3,.15)); }
  return {text:text.slice(0,start)+block+text.slice(end), count:count};
}

function applyRules_(text, rules){
  var tokens=[], results=[];
  rules.forEach(function(r,i){
    var count=0;
    if(r.fn){ var o=r.fn(text); text=o.text; count=o.count; }
    else if(r.find instanceof RegExp){
      var flags=r.find.flags.indexOf('g')>=0?r.find.flags:r.find.flags+'g';
      var re=new RegExp(r.find.source, flags);
      var m=text.match(re); count=m?m.length:0;
      if(count){ var tk='\u0001R'+i+'\u0001'; text=text.replace(re, function(){ return tk; }); tokens.push([tk,r.replace]); }
    } else {
      var parts=text.split(r.find); count=parts.length-1;
      if(count){ var tk2='\u0001R'+i+'\u0001'; text=parts.join(tk2); tokens.push([tk2,r.replace]); }
    }
    var status;
    if(r.expect==null) status = count>0 ? 'ok' : 'miss';
    else if(count===r.expect) status='ok';
    else if(count===0 && r.optional) status='skipped';
    else status='mismatch';
    results.push({id:r.id, label:r.label, count:count, expect:r.expect, status:status,
      find: (r.find instanceof RegExp) ? String(r.find) : (r.fn ? '(:root block)' : String(r.find)),
      replace: r.fn ? '(recomputed colour variables)' : String(r.replace)});
  });
  tokens.forEach(function(t){ text=text.split(t[0]).join(t[1]); });
  return {text:text, results:results};
}

function countOf_(text, s){ return s ? text.split(s).length-1 : 0; }
function linesContaining_(text, s){
  var outL=[]; var lines=text.split('\n');
  for(var i=0;i<lines.length;i++){ if(lines[i].indexOf(s)>=0){ var L=lines[i].replace(/data:image\/png;base64,[A-Za-z0-9+\/=]+/g,'data:…').trim(); outL.push((i+1)+': '+(L.length>150?L.slice(0,150)+'…':L)); } }
  return outL;
}
function scanResidue_(html, gs, plan, pages){
  var checks=RESIDUE_CHECKS.slice();
  if(plan.masterScriptId) checks.push(plan.masterScriptId);
  if(plan.masterSheetId) checks.push(plan.masterSheetId);
  if(plan.masterLogoHead) checks.push(plan.masterLogoHead);
  checks.push("'"+MASTER_REF.storagePrefix);
  var pageNames=Object.keys(pages||{});
  var rows=checks.map(function(s){
    var pc=0; pageNames.forEach(function(n){ pc+=countOf_(pages[n],s); });
    return {text:s.length>40?s.slice(0,20)+'…(master logo/ID)':s, html:countOf_(html,s), gs:countOf_(gs,s), pages:pc};
  });
  // Client pages have no vendor text at all, so any "Deep Pocket" left there is residue.
  if(pageNames.length){ var dp=0; pageNames.forEach(function(n){ dp+=countOf_(pages[n],'Deep Pocket'); }); rows.push({text:'Deep Pocket (client pages)', html:0, gs:0, pages:dp}); }
  return rows;
}


function executePlan_(masters, plan){
  var h=applyRules_(masters.html, plan.html);
  var g=applyRules_(masters.gs, plan.gs);
  var pageOut={}, pageResults={};
  Object.keys(plan.pages).forEach(function(n){ var r=applyRules_(masters.pages[n], plan.pages[n]); pageOut[n]=r.text; pageResults[n]=r.results; });
  var residue=scanResidue_(h.text, g.text, plan, pageOut);
  var vendorLines=linesContaining_(h.text,'Deep Pocket');
  var missingPages=CLIENT_PAGES.filter(function(n){ return !masters.pages[n]; });
  var report=buildReportMd_(null, plan, h.results, g.results, residue, vendorLines, pageResults, missingPages);
  var all=h.results.concat(g.results); Object.keys(pageResults).forEach(function(n){ all=all.concat(pageResults[n]); });
  var mismatches=all.filter(function(r){ return r.status==='mismatch'||r.status==='miss'; });
  var dirty=residue.filter(function(r){ return r.html+r.gs+r.pages>0; });
  return { html:h.text, gs:g.text, pages:pageOut, clean: mismatches.length===0 && dirty.length===0, values:plan.values,
    htmlResults:h.results, gsResults:g.results, pageResults:pageResults, missingPages:missingPages, residue:residue, vendorLines:vendorLines, report:report };
}


function buildReportMd_(intake, plan, hr, gr, residue, vendorLines, pageResults, missingPages){
  var v=plan.values, L=[];
  L.push('# Build report — '+v.legal, '');
  L.push('Product: **'+v.product+'** · Home-screen name: **'+v.short+'** · Storage namespace: `'+v.ns+'` · Suggested GitHub folder: `dp-crm/app/'+v.deploySlug+'/`', '');
  function table(title, rows){
    L.push('## '+title,'','| Status | Rule | Found | Expected |','|---|---|---|---|');
    rows.forEach(function(r){ L.push('| '+(r.status==='ok'?'✅ ok':r.status==='skipped'?'➖ skipped':'⚠️ '+r.status)+' | '+r.label+' | '+r.count+' | '+(r.expect==null?'any':r.expect)+' |'); });
    L.push('');
  }
  table('crm.html', hr); table('Code.gs', gr);
  Object.keys(pageResults||{}).forEach(function(n){ table(n, pageResults[n]); });
  if(missingPages && missingPages.length) L.push('⚠️ Master client pages not found (not built): '+missingPages.join(', '), '');
  L.push('config.js SCRIPT_URL: '+(v.scriptUrl?'`'+v.scriptUrl+'`':'not set (stage "code")'), '');
  L.push('## Residue check (must all be 0)','','| String | crm.html | Code.gs | client pages |','|---|---|---|---|');
  residue.forEach(function(r){ L.push('| `'+r.text+'` | '+r.html+' | '+r.gs+' | '+(r.pages||0)+' |'); });
  L.push('', '## Remaining "Deep Pocket" lines in crm.html (vendor references — expected to stay)', '');
  (vendorLines||[]).forEach(function(l){ L.push('- `'+l.replace(/`/g,"'")+'`'); });
  L.push('');
  L.push('A ⚠️ mismatch means the master changed since MASTER_REF was last updated: open that rule in dp-build.js → buildRebrandPlan_ and adjust the anchor or expected count.');
  return L.join('\n');
}


// ── FIRM SETTINGS block shared by the 6 client pages (window.DP_CONFIG) ──
function dpConfigBlock_(intake, v, icons){
  var br=intake.brand||{}, bz=intake.business||{}, cp=intake.compliance||{};
  var theme={}; Object.keys(MASTER_REF.theme).forEach(function(k){ var t=(br.theme||{})[k]; theme[k]=isHex_(t)?t.toUpperCase():MASTER_REF.theme[k]; });
  var fm={}; (intake.features||[]).forEach(function(f){ fm[f.key]=f.mode; });
  var features={}; PAGE_FEATURES.forEach(function(k){ features[k]=fm[k]!=='remove'; });
  if(!v.hasOwnApp) features.appinstalled=false;
  var codes=(bz.productList&&bz.productList.length)?bz.productList.slice():Object.keys(PRODUCT_LABELS);
  var products=codes.map(function(c){ return {code:c, label:PRODUCT_LABELS[c]||c}; });
  String(bz.otherProducts||'').split(',').map(trim_).filter(Boolean).forEach(function(p){ products.push({code:p.toUpperCase(), label:p}); });
  var custom=br.typographyMode==='custom';
  var today=new Date().toISOString().slice(0,10);
  var cfg={
    version:1, builtAt:today, storagePrefix:v.ns,
    firm:{ legalName:v.legal, brandName:v.brand, productName:v.product, shortName:v.short, ownerName:v.owner, ownerTitle:v.title,
      address:v.address, city:v.city, phone:v.welcomePhone||v.phone, email:v.email, website:v.website,
      modelType:bz.modelType||'MFD', arn:v.arn, sebiReg:trim_(bz.sebiReg).toUpperCase() },
    theme:theme,
    fonts:{ display:(custom&&trim_(br.displayFont))||'Fraunces', body:(custom&&trim_(br.bodyFont))||'Inter' },
    logo:(icons&&icons.icon192)?'data:image/png;base64,'+icons.icon192:'',
    app:{ hasOwnApp:v.hasOwnApp, android:v.hasOwnApp?v.android:'', ios:v.hasOwnApp?v.ios:'' },
    products:products, features:features,
    privacy:{ grievanceName:trim_(cp.grievanceName)||v.owner, grievanceTitle:trim_(cp.grievanceTitle)||'Grievance Officer',
      grievanceEmail:trim_(cp.grievanceEmail)||v.email, grievancePhone:trim_(cp.grievancePhone)||v.phone,
      lastUpdated:today, effectiveFrom:today, aiProvider:'Google Gemini', extraDisclosures:trim_(cp.extraDisclosures) }
  };
  return '/*DP_CONFIG_START*/\n/* ═══ FIRM SETTINGS — the only firm-specific part of this page. ═══\n'
    + '   Written by dp-build.js from the firm\'s client form; every client page\n'
    + '   carries the same block. The Apps Script URL is NOT here — it lives only\n'
    + '   in config.js. Re-run the build to change it, or edit identically in all six pages. */\n'
    + 'window.DP_CONFIG = '+JSON.stringify(cfg,null,2)+';\n/*DP_CONFIG_END*/';
}

function buildConfigJs_(v){
  return '// '+v.product+' — backend address for '+v.legal+'\n'
    + '// Read by crm.html and all six client pages. Keep this file in the SAME folder.\n'
    + '// If the Apps Script is ever deployed as a NEW deployment (new URL), change\n'
    + '// SCRIPT_URL here — no other file needs editing.\n'
    + (v.sheetId?'// Google Sheet ID: '+v.sheetId+'\n':'')
    + "var SCRIPT_URL = '"+v.scriptUrl+"';\n";
}

// ── CLI ──
function readOpt(n){ var p=path.join(DIR,n); return fs.existsSync(p)?fs.readFileSync(p,'utf8'):null; }
function readIcon(n){ var p=path.join(DIR,n); return fs.existsSync(p)?fs.readFileSync(p).toString('base64'):''; }
function main(){
  var stage=(process.argv[2]||'').toLowerCase(), url=String(process.argv[3]||'').trim().replace(/[?#].*$/,'');
  if(['code','app','check','verify'].indexOf(stage)<0){ console.log('Usage: node dp-build.js code | app <EXEC_URL> | check | verify'); process.exit(2); }
  var intake=JSON.parse(readOpt('firm.json')||'null'); if(!intake){ console.log('ERROR: firm.json not found next to dp-build.js'); process.exit(1); }
  if(stage==='verify') return verify_(intake);
  intake.technical=intake.technical||{};
  if(stage==='app'){
    if(/\/dev$/.test(url)){ console.log('ERROR: that is the /dev test URL. Use the Web app URL ending in /exec.'); process.exit(1); }
    if(!EXEC_RE.test(url)){ console.log('ERROR: not an Apps Script web-app URL: '+url); process.exit(1); }
    intake.technical.appsScriptUrl=url;
  }
  var gs=readOpt('Code.gs')||readOpt('crm_Code.gs')||readOpt('Code.js'), html=readOpt('crm.html');
  if(!gs||!html){ console.log('ERROR: master crm.html and Code.gs must be next to dp-build.js'); process.exit(1); }
  var pages={}; CLIENT_PAGES.forEach(function(n){ var t=readOpt(n); if(t) pages[n]=t; });
  var icons={icon192:readIcon('icon-192.png'), icon512:readIcon('icon-512.png')}; if(!icons.icon192) icons=null;
  var masters={html:html, gs:gs, pages:pages};
  var plan=buildRebrandPlan_(intake, icons, masters);
  var r=executePlan_(masters, plan), v=plan.values;
  if(!fs.existsSync(OUT)) fs.mkdirSync(OUT);
  var written=[];
  function put(n,t){ fs.writeFileSync(path.join(OUT,n), t); written.push(n); }
  if(stage==='code') put('Code.gs', r.gs);
  if(stage==='app'){ put('config.js', buildConfigJs_(v)); put('crm.html', r.html); Object.keys(r.pages).forEach(function(n){ put(n, r.pages[n]); }); }
  if(stage!=='check') fs.writeFileSync(path.join(OUT,'BUILD-REPORT-'+stage+'.md'), r.report);
  // Short summary (keeps Claude's context small).
  var scope = stage==='code' ? r.gsResults.map(function(x){ return ['Code.gs',x]; })
    : r.htmlResults.map(function(x){ return ['crm.html',x]; }).concat(stage==='check'?r.gsResults.map(function(x){ return ['Code.gs',x]; }):[]);
  Object.keys(r.pageResults).forEach(function(n){ if(stage!=='code') r.pageResults[n].forEach(function(x){ scope.push([n,x]); }); });
  var bad=scope.filter(function(p){ return p[1].status!=='ok' && p[1].status!=='skipped'; });
  var dirty=r.residue.filter(function(x){ var n=stage==='code'?x.gs:(stage==='app'?x.html+(x.pages||0):x.html+x.gs+(x.pages||0)); return n>0; });
  console.log('dp-build '+BUILDER_VERSION+' · stage '+stage+' · '+v.legal+' · namespace '+v.ns);
  if(written.length) console.log('Written to out/: '+written.join(', '));
  console.log('Rules: '+scope.length+' checked, '+(scope.length-bad.length)+' ok, '+bad.length+' problem(s)');
  bad.forEach(function(p){ console.log('  ⚠ '+p[0]+' · '+p[1].label+' · found '+p[1].count+', expected '+p[1].expect); });
  if(r.missingPages.length) console.log('  ⚠ missing master pages: '+r.missingPages.join(', '));
  console.log('Leftover master identity: '+(dirty.length?dirty.map(function(x){ return x.text; }).join(', '):'none'));
  if(stage==='code' && !v.sheetId) console.log('  ⚠ No Sheet ID in firm.json — SHEET_ID is a placeholder in Code.gs');
  if(!icons) console.log('  ⚠ icon-192.png not found — master logo kept');
}

// Final checks on the delivered files (stage 3). Prints problems only.
function verify_(intake){
  var v=resolveValues_(intake), probs=[], fm={};
  (intake.features||[]).forEach(function(f){ fm[f.key]=f.mode; });
  var PAGE_OF={documents:'upload.html', gatherinfo:'gatherinfo.html', findoc:'findoc.html', feedback:'feedback.html', reviewreport:'review.html'};
  var deploy=['config.js','crm.html'].concat(CLIENT_PAGES.filter(function(n){ return !Object.keys(PAGE_OF).some(function(k){ return PAGE_OF[k]===n && fm[k]==='remove'; }); }));
  function rd(n){ var p=path.join(OUT,n); return fs.existsSync(p)?fs.readFileSync(p,'utf8'):null; }
  deploy.forEach(function(n){ if(rd(n)==null) probs.push('missing out/'+n); });
  // Syntax: every inline <script> in HTML, plus config.js and Code.gs.
  deploy.concat(['Code.gs']).forEach(function(n){
    var t=rd(n); if(t==null) return;
    var chunks=/\.html$/.test(n) ? Array.from(t.matchAll(/<script(?![^>]*\bsrc=)[^>]*>([\s\S]*?)<\/script>/g)).map(function(m){ return m[1]; }) : [t];
    chunks.forEach(function(c,i){ try{ new Function(c); }catch(e){ probs.push('syntax error in '+n+(chunks.length>1?' script #'+(i+1):'')+': '+e.message); } });
  });
  var cj=rd('config.js')||'', m=cj.match(/SCRIPT_URL\s*=\s*'([^']*)'/);
  if(!m || !EXEC_RE.test(m[1])) probs.push('config.js has no valid SCRIPT_URL');
  // Leftover master identity.
  var checks=RESIDUE_CHECKS.concat(["'wm_", '"wm_"']);
  deploy.forEach(function(n){ var t=rd(n); if(t==null) return; checks.forEach(function(c){ var k=t.split(c).length-1; if(k) probs.push(n+': "'+c+'" ×'+k); }); });
  CLIENT_PAGES.forEach(function(n){ var t=rd(n); if(t && t.indexOf('Deep Pocket')>=0) probs.push(n+': contains "Deep Pocket" (client pages must have none)'); });
  var html=rd('crm.html')||'';
  var VENDOR=/Feedback to Deep Pocket|This goes to Deep Pocket|Contact Deep Pocket Fintech Pvt Ltd|contact Deep Pocket Fintech Pvt|Deep Pocket CRM — Diagnostic Report|sent to Deep Pocket|reach Deep Pocket|^\s*\/\/|Deep Pocket's support inbox|e\.g\. "Deep Pocket CRM"/;
  var odd=html.split('\n').map(function(l,i){ return [i+1,l]; }).filter(function(p){ return p[1].indexOf('Deep Pocket')>=0 && !VENDOR.test(p[1]); });
  odd.forEach(function(p){ probs.push('crm.html line '+p[0]+' (check: vendor or firm?): '+p[1].trim().slice(0,120)); });
  console.log('verify · '+v.legal+' · '+(probs.length?probs.length+' item(s) to look at:':'all checks passed'));
  probs.forEach(function(p){ console.log('  ⚠ '+p); });
  var zipName=v.deploySlug+'-deploy.zip';
  console.log('Deploy files: '+deploy.join(', '));
  console.log('Package: cd '+OUT+' && rm -f '+zipName+' && zip -q '+zipName+' '+deploy.join(' '));
  console.log('GitHub folder: dp-crm/app/'+v.deploySlug+'/   ·   App: https://dp-crm.github.io/app/'+v.deploySlug+'/crm.html');
}

if(require.main===module) main();
module.exports={buildRebrandPlan_:buildRebrandPlan_, executePlan_:executePlan_, resolveValues_:resolveValues_};
