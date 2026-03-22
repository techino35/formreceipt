/**
 * Mailer.js - FormReceipt メール送信モジュール
 */

function sendReceiptEmail(recipientEmail, pdfBlob, data, licenseType) {
  const props = PropertiesService.getScriptProperties();
  const subject = _buildSubject(props, data);
  const body = _buildBody(props, data, licenseType);
  const options = {
    attachments: [pdfBlob],
    name: props.getProperty("FORMRECEIPT_SENDER_NAME") || "FormReceipt",
  };
  const replyTo = props.getProperty("FORMRECEIPT_REPLY_TO");
  if (replyTo) options.replyTo = replyTo;
  try {
    MailApp.sendEmail(recipientEmail, subject, body, options);
    console.log("FORMRECEIPT: メール送信完了 -> " + recipientEmail);
  } catch (err) {
    console.error("FORMRECEIPT: メール送信失敗 - " + err.message);
    throw err;
  }
}

function _buildSubject(props, data) {
  const template = props.getProperty("FORMRECEIPT_MAIL_SUBJECT") || "【受領確認】{{受付番号}} フォームを受け付けました";
  return _interpolate(template, data);
}

function _buildBody(props, data, licenseType) {
  const defaultBody = [
    "この度はお申し込みいただきありがとうございます。",
    "",
    "受付番号: {{受付番号}}",
    "送信日時: {{送信日時}}",
    "",
    "受領確認書をPDFで添付しております。ご確認ください。",
    "",
    "ご不明な点がございましたら、お気軽にお問い合わせください。",
  ].join("\n");
  let body = _interpolate(props.getProperty("FORMRECEIPT_MAIL_BODY") || defaultBody, data);
  if (licenseType === "free") body += "\n\n--\nThis email was sent via FormReceipt (Free Plan)";
  return body;
}

function _interpolate(template, data) {
  return template.replace(/\{\{([^}]+)\}\}/g, (match, key) => {
    return data[key.trim()] !== undefined ? String(data[key.trim()]) : match;
  });
}