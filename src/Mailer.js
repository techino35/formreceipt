/**
 * Mailer.js - FormReceipt email sending module
 */

function sendReceiptEmail(recipientEmail, pdfBlob, data, licenseType) {
  const props = PropertiesService.getScriptProperties();
  const subject = _buildSubject(props, data);
  const body = _buildBody(props, data, licenseType);
  const options = { attachments: [pdfBlob], name: props.getProperty("FORMRECEIPT_SENDER_NAME") || "FormReceipt" };
  const replyTo = props.getProperty("FORMRECEIPT_REPLY_TO");
  if (replyTo) options.replyTo = replyTo;
  try {
    MailApp.sendEmail(recipientEmail, subject, body, options);
    console.log("FORMRECEIPT: Email sent -> " + recipientEmail);
  } catch (err) {
    console.error("FORMRECEIPT: Email send failed - " + err.message);
    throw err;
  }
}

function _buildSubject(props, data) {
  const template = props.getProperty("FORMRECEIPT_MAIL_SUBJECT") || "[Receipt] {{Receipt Number}} - Form submitted";
  return _interpolate(template, data);
}

function _buildBody(props, data, licenseType) {
  const defaultBody = ["Thank you for your submission.", "", "Receipt Number: {{Receipt Number}}", "Submission Date/Time: {{Submission Date/Time}}", "", "Please find the receipt PDF attached."].join("\n");
  const template = props.getProperty("FORMRECEIPT_MAIL_BODY") || defaultBody;
  let body = _interpolate(template, data);
  if (licenseType === "free") body += "\n\n--\nThis email was sent via FormReceipt (Free Plan)\nhttps://formreceipt.app";
  return body;
}

function _interpolate(template, data) {
  return template.replace(/\{\{([^}]+)\}\}/g, (match, key) => data[key.trim()] !== undefined ? String(data[key.trim()]) : match);
}
