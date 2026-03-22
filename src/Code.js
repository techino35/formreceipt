/**
 * Code.js - FormReceipt エントリーポイント
 */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("FormReceipt")
    .addItem("設定を開く", "FORMRECEIPT_showSidebar")
    .addItem("トリガーを設定", "FORMRECEIPT_setupTrigger")
    .addItem("トリガーを削除", "FORMRECEIPT_removeTrigger")
    .addSeparator()
    .addItem("今月の利用件数を確認", "FORMRECEIPT_showUsage")
    .addToUi();
}

function FORMRECEIPT_showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile("Sidebar")
    .setTitle("FormReceipt 設定")
    .setWidth(320);
  SpreadsheetApp.getUi().showSidebar(html);
}

function FORMRECEIPT_setupTrigger() {
  FORMRECEIPT_removeTrigger();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ScriptApp.newTrigger("FORMRECEIPT_onFormSubmit")
    .forSpreadsheet(ss)
    .onFormSubmit()
    .create();
  SpreadsheetApp.getUi().alert("トリガーを設定しました。フォーム送信時に自動でPDFが生成されます。");
}

function FORMRECEIPT_removeTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === "FORMRECEIPT_onFormSubmit") {
      ScriptApp.deleteTrigger(trigger);
    }
  }
}

function FORMRECEIPT_onFormSubmit(e) {
  try {
    const processor = new FormReceiptProcessor(e);
    processor.run();
  } catch (err) {
    console.error("FORMRECEIPT_onFormSubmit error: " + err.message);
    FORMRECEIPT_notifyError(err);
  }
}

function FORMRECEIPT_notifyError(err) {
  const props = PropertiesService.getScriptProperties();
  const adminEmail = props.getProperty("FORMRECEIPT_ADMIN_EMAIL");
  if (!adminEmail) return;
  MailApp.sendEmail({
    to: adminEmail,
    subject: "[FormReceipt] エラーが発生しました",
    body: "エラー内容:\n" + err.message + "\n\nスタック:\n" + (err.stack || "なし"),
  });
}

function FORMRECEIPT_showUsage() {
  const namespace = "FORMRECEIPT_" + getYearMonthString();
  const props = PropertiesService.getScriptProperties();
  const count = parseInt(props.getProperty("SEQ_" + namespace) || "0", 10);
  const license = FORMRECEIPT_getLicenseType();
  const limit = license === "pro" ? "無制限" : "10件";
  SpreadsheetApp.getUi().alert(
    "今月の利用状況\n\nプラン: " + license.toUpperCase() + "\n送信件数: " + count + "件 / " + limit
  );
}

function FORMRECEIPT_saveConfig(config) {
  const props = PropertiesService.getScriptProperties();
  props.setProperty("FORMRECEIPT_TEMPLATE_DOC_ID", config.templateDocId || "");
  props.setProperty("FORMRECEIPT_ROOT_FOLDER_ID", config.rootFolderId || "");
  props.setProperty("FORMRECEIPT_ADMIN_EMAIL", config.adminEmail || "");
  props.setProperty("FORMRECEIPT_LICENSE_KEY", config.licenseKey || "");
  return "設定を保存しました";
}

function FORMRECEIPT_getConfig() {
  const props = PropertiesService.getScriptProperties();
  return {
    templateDocId: props.getProperty("FORMRECEIPT_TEMPLATE_DOC_ID") || "",
    rootFolderId: props.getProperty("FORMRECEIPT_ROOT_FOLDER_ID") || "",
    adminEmail: props.getProperty("FORMRECEIPT_ADMIN_EMAIL") || "",
    licenseKey: props.getProperty("FORMRECEIPT_LICENSE_KEY") || "",
    licenseType: FORMRECEIPT_getLicenseType(),
  };
}

// ────────────────────────────────────
// Shared utility helpers
// ────────────────────────────────────

function getYearMonthString() {
  return Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy-MM");
}

function getTodayString() {
  return Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyyMMdd");
}

function getNextSequence(namespace) {
  const props = PropertiesService.getScriptProperties();
  const key = "SEQ_" + namespace;
  const current = parseInt(props.getProperty(key) || "0", 10);
  const next = current + 1;
  props.setProperty(key, String(next));
  return next;
}

function replacePlaceholders(templateDocId, data) {
  const originalDoc = DriveApp.getFileById(templateDocId);
  const tempFile = originalDoc.makeCopy("[TEMP] " + originalDoc.getName());
  const tempDoc = DocumentApp.openById(tempFile.getId());
  const body = tempDoc.getBody();

  Object.keys(data).forEach(key => {
    body.replaceText("{{" + key + "}}", String(data[key] || ""));
  });

  tempDoc.saveAndClose();
  return tempFile.getId();
}

function exportToPdf(docId) {
  const file = DriveApp.getFileById(docId);
  const pdfBlob = file.getAs("application/pdf");
  DriveApp.getFileById(docId).setTrashed(true); // clean up temp
  return pdfBlob;
}

function getOrCreateFolder(parentId, path) {
  const parts = path.split("/");
  let current = DriveApp.getFolderById(parentId);
  for (const part of parts) {
    if (!part) continue;
    const iter = current.getFoldersByName(part);
    current = iter.hasNext() ? iter.next() : current.createFolder(part);
  }
  return current;
}

function saveToFolder(folderId, filename, blob) {
  const folder = DriveApp.getFolderById(folderId);
  blob.setName(filename);
  return folder.createFile(blob);
}