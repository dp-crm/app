/**
 * Deep Pocket CRM Customisation Intake — submission notifier.
 *
 * Deploy: Extensions > Apps Script (or script.google.com/New project) >
 * paste this in > Deploy > New deployment > Web app
 *   Execute as: Me
 *   Who has access: Anyone
 * Copy the /exec URL into SCRIPT_URL in deep-pocket-client_form.html.
 */

var NOTIFY_EMAIL = 'PASTE_YOUR_EMAIL_HERE'; // where the completed package is sent

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);

    if (data.action === 'notifySubmission') {
      sendImplementationPackageEmail(data);
    }

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function sendImplementationPackageEmail(data) {
  var firmName = data.firmName || '(no firm name given)';
  var summary = data.summary || {};
  var prompts = data.prompts || [];

  var subject = firmName + ' — Generated CRM Setup Prompt';

  // ---- Plain-text fallback ----
  var textBody = 'DEEP POCKET CRM — CLIENT CUSTOMISATION PACKAGE\n\n'
    + summaryTextBlock(summary)
    + '\n\n' + prompts.map(function (p) {
        return '──────────────────────\n' + p.title + '\n──────────────────────\n\n' + p.text + '\n';
      }).join('\n');

  // ---- HTML body ----
  var htmlBody = ''
    + '<div style="font-family:-apple-system,Segoe UI,Roboto,Arial,sans-serif;max-width:680px;margin:0 auto;color:#1a1a1a">'
    + '<div style="background:#0A1830;color:#fff;padding:16px 20px;border-radius:10px 10px 0 0">'
    + '<div style="font-size:11px;letter-spacing:.08em;text-transform:uppercase;color:#8FB4DC">Deep Pocket CRM</div>'
    + '<div style="font-size:18px;font-weight:800;margin-top:2px">Client Customisation Package</div>'
    + '</div>'
    + '<div style="border:1px solid #e0e0e0;border-top:none;border-radius:0 0 10px 10px;padding:20px">'
    + summaryHtmlBlock(summary)
    + '</div>'

    + '<div style="margin-top:24px;font-size:12px;font-weight:800;color:#0A1830;letter-spacing:.04em">IMPLEMENTATION PROMPTS \u2014 PASTE INTO CLAUDE ONE STAGE AT A TIME, IN ORDER</div>'

    + prompts.map(function (p) {
        return '<div style="margin:14px 0;border:1px solid #e0e0e0;border-radius:8px;overflow:hidden">'
          + '<div style="background:#0A1830;color:#fff;padding:10px 14px;font-weight:700;font-size:13px">' + escapeHtml(p.title) + '</div>'
          + '<pre style="white-space:pre-wrap;font-family:Menlo,Consolas,monospace;font-size:11.5px;line-height:1.6;padding:14px;margin:0;background:#f7f8fa">' + escapeHtml(p.text) + '</pre>'
          + '</div>';
      }).join('')
    + '</div>';

  var options = { htmlBody: htmlBody };

  // Attach the uploaded logo, if one was sent
  if (data.logoBase64) {
    var logoBlob = Utilities.newBlob(
      Utilities.base64Decode(data.logoBase64),
      'image/png',
      (firmName.replace(/[^a-z0-9]/gi, '_') || 'firm') + '_logo.png'
    );
    options.attachments = [logoBlob];
  }

  MailApp.sendEmail(NOTIFY_EMAIL, subject, textBody, options);
}

function summaryTextBlock(s) {
  return 'CLIENT SUMMARY\n'
    + 'Firm: ' + (s.firmName || '\u2014') + '\n'
    + 'CRM Name: ' + (s.crmName || '\u2014') + '\n'
    + 'Contact: ' + (s.contactName || '\u2014') + ' \u2014 ' + (s.designation || '\u2014') + '\n'
    + 'Phone: ' + (s.phone || '\u2014') + ' | Email: ' + (s.email || '\u2014') + '\n'
    + 'Address: ' + (s.address || '\u2014') + '\n'
    + 'Website: ' + (s.website || '\u2014') + '\n\n'
    + 'BRANDING\n'
    + 'Primary Colour: ' + (s.primaryColour || '\u2014') + ' | Accent Colour: ' + (s.accentColour || '\u2014') + '\n'
    + 'Logo: ' + (s.logo || '\u2014') + '\n\n'
    + 'BUSINESS PROFILE\n'
    + 'Type: ' + (s.businessTypes || '\u2014') + '\n'
    + 'ARN: ' + (s.arn || '\u2014') + ' | SEBI Reg: ' + (s.sebiReg || '\u2014') + '\n'
    + 'Primary Services: ' + (s.primaryServices || '\u2014') + '\n\n'
    + 'CONTACT & EMAIL MAPPING\n'
    + 'Support Phone: ' + (s.supportPhone || '\u2014') + ' | Support Email: ' + (s.supportEmail || '\u2014') + '\n'
    + 'Communication Style: ' + (s.commsStyle || '\u2014') + '\n\n'
    + 'PREMIUM FEATURES\n' + (s.premiumFeatures || '\u2014') + '\n\n'
    + 'AI MANAGEMENT\n' + (s.aiManagement || '\u2014');
}

function summaryHtmlBlock(s) {
  function row(k, v) {
    return '<tr><td style="padding:3px 12px 3px 0;color:#666;white-space:nowrap;vertical-align:top">' + escapeHtml(k) + '</td><td style="padding:3px 0;font-weight:600">' + escapeHtml(v || '\u2014') + '</td></tr>';
  }
  function section(title, rows) {
    return '<div style="margin-bottom:16px">'
      + '<div style="font-size:11px;font-weight:800;letter-spacing:.06em;color:#0A1830;text-transform:uppercase;margin-bottom:6px">' + escapeHtml(title) + '</div>'
      + '<table style="font-size:13px;border-collapse:collapse">' + rows + '</table></div>';
  }
  return section('Client Summary',
      row('Firm', s.firmName) + row('CRM Name', s.crmName) + row('Contact', (s.contactName || '') + ' — ' + (s.designation || ''))
      + row('Phone', s.phone) + row('Email', s.email) + row('Address', s.address) + row('Website', s.website))
    + section('Branding', row('Primary Colour', s.primaryColour) + row('Accent Colour', s.accentColour) + row('Logo', s.logo))
    + section('Business Profile', row('Type', s.businessTypes) + row('ARN', s.arn) + row('SEBI Reg.', s.sebiReg) + row('Primary Services', s.primaryServices))
    + section('Contact & Email Mapping', row('Support Phone', s.supportPhone) + row('Support Email', s.supportEmail) + row('Communication Style', s.commsStyle))
    + section('Premium Features', row('Selected', s.premiumFeatures))
    + section('AI Management', row('Model', s.aiManagement));
}

function escapeHtml(str) {
  return String(str || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}
