/* config.js — the ONE file that differs from firm to firm.
   crm.html, Code.gs and the six client pages are the unchanged masters; everything below is what makes the app this firm's own.
   Every field is optional: leave one out and the app uses the default shown. Full reference: CONFIG_REFERENCE.md.
   Replace this file (same folder as crm.html) to change anything — open apps pick up the change at their next start
   *if configId changes*, so always set a new configId when you edit. */

var SCRIPT_URL = 'PASTE_YOUR_APPS_SCRIPT_EXEC_URL_HERE';   // Apps Script → Deploy → Manage deployments → Web app URL (…/exec)

window.DP_CONFIG = {
  "version": 2,                       // format version — keep 2
  "configId": "2026-10-09T10:00",     // CHANGE THIS every time the file changes (any text); apps reload once when it differs
  "builtAt": "2026-10-09",
  "storagePrefix": "acme_",           // 2–16 letters/digits + "_". Keeps firms on the same web address from seeing each other's saved data

  "firm": {
    "legalName": "Acme Wealth Advisors LLP",     // REQUIRED. Printed in messages, signatures, folder names
    "brandName": "Acme Wealth Advisors",         // default: legal name without "Pvt. Ltd."/"LLP"…
    "productName": "Acme Wealth Advisors CRM",   // default: brand + " CRM". Shown on login, menus, install name
    "shortName": "Acme",                         // default: brand if ≤12 letters, else its first word. Home-screen name
    "ownerName": "Ravi Menon",                   // REQUIRED. Message signatures
    "signOffName": "Ravi",                       // default: owner's first name. Short sign-off
    "ownerTitle": "Managing Partner",            // optional; shown under the name in the welcome message
    "address": "4th Floor, Marine Tower\nBeach Road, Kochi",   // \n = new line
    "city": "Kochi",
    "phone": "9000011111",                       // office phone printed in the welcome message
    "welcomePhone": "",                          // only if different from phone
    "email": "ravi@acme.example",
    "website": "www.acmewealth.example",
    "modelType": "MFD",
    "arn": "ARN-12345",                          // optional; also saved to the sheet
    "sebiReg": "",
    "gstNumber": ""
  },

  "theme": {                                     // each value must be #RRGGBB; leave any out for the default
    "bg": "#0B1F33",      "s1": "#102A44",   "s2": "#14304F",   "s3": "#193759",   "border": "#234A73",
    "accent": "#EAF1FA",  "coral": "#F4B942",
    "text": "#EAF1FA",    "t2": "#8FB4DC",   "t3": "#5D7A9E",
    "green": "#4C9A6A",   "blue": "#4A8FC7", "orange": "#C08A3E", "red": "#C1553D"
  },
  "fonts": { "display": "Playfair Display", "body": "Lato", "google": true },   // any Google Fonts family; google:false = use device fonts
  "logo": "",       // data:image/png;base64,… (192px square) or leave "" for an initials badge in the firm's colours
  "logoHD": "",     // 512px version for the install icon (optional)

  "app": { "hasOwnApp": true, "android": "https://play.google.com/store/apps/details?id=in.acme", "ios": "https://apps.apple.com/in/app/acme/id1" },
                    // hasOwnApp:false removes the app-link message, the "App Installed" tick and its row

  "stages": ["LEAD CHECK", "PAPERWORK", "FIRST INVESTMENT", "THANK YOU CALL"],   // client pipeline, in order (2–10). COMPLETED is added automatically
  "stageIcons": { "LEAD CHECK": "🔎" },          // optional per-stage emoji on the dashboard

  "products": [                                  // code = stored value (UPPER CASE); label = chip text in the CRM; pageLabel = wording on client pages
    { "code": "INVESTMENT", "label": "Investment", "icon": "📈", "lead": true,  "pageLabel": "Mutual Funds" },
    { "code": "PMS",        "label": "PMS",        "icon": "💼", "lead": false, "pageLabel": "PMS" },     // lead:false = client form only
    { "code": "WILL",       "label": "WILL",       "icon": "📜", "lead": true,  "pageLabel": "Will Writing" }
  ],

  "features": {                                  // only list what you switch OFF (false). Anything not listed is ON.
    // "family": false, "documents": false, "gatherinfo": false, "findoc": false, "feedback": false, "reviewreport": false,
    // "casanalysis": false, "meetinginsights": false, "practicechat": false, "topinsights": false, "contemporary": false,
    // "dailyinsights": false, "voiceagent": false, "whatsappdrafts": false, "sip": false, "calendar": false, "clientmail": false,
    // "applock": false, "appinstalled": false, "specialdays": false
    // (pipeline, tasks, multiuser, backups and dpdp are core and cannot be switched off)
  },

  "defaults": {
    "followUpDays": 14,                          // a lead untouched this long is flagged for follow-up (1–120)
    "sipReminderDays": [5, 2],                   // remind this many days before a SIP date (whole numbers 1–15)
    "sipBounceDays": 7,                          // check a SIP was debited this many days after its date (1–25)
    "goalLeads": 30, "goalConversions": 5,       // monthly targets
    "voiceLanguage": "en-IN|English"             // en-IN|English, hi-IN|Hindi, ml-IN|Malayalam, ta-IN|Tamil, te-IN|Telugu, kn-IN|Kannada
  },

  "messages": {                                  // all optional. {name} and {products} are filled in per client/lead.
    "prospect": "",                              // {{placeholders}} fill in the firm's details — see CONFIG_REFERENCE.md
    "contacted": "", "inProgress": "", "converted": "",
    "welcome": "",                               // the firm's exact welcome wording; leave "" for the standard one built from the details above
    "titleLine": ""                              // line under the owner's name in the standard welcome; default "<ownerTitle>, <brandName>"
  },

  "privacy": {                                   // printed in the privacy notice on client pages (DPDP) — name and email are REQUIRED
    "grievanceName": "Ravi Menon", "grievanceTitle": "Grievance Officer", "grievanceEmail": "ravi@acme.example", "grievancePhone": "9000011111",
    "lastUpdated": "2026-10-09", "effectiveFrom": "2026-10-09", "aiProvider": "Google Gemini", "extraDisclosures": ""
  }
};
