/**
 * Code.js - FormReceipt entry point
 */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("FormReceipt")
    .addItem("Open Settings", "FORMRECEIPT_showSidebar")
    .addItem("Setup Trigger", "FORMRECEIPT_setupTrigger")
    .addItem("Remove Trigger", "FORMRECEIPT_removeTrigger")
    .addSeparator()
    .addItem("Check Monthly Usage", "FORMRECEIPT_showUsage")
    .addToUi();
}

function FORMRECEIPT_showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile("Sidebar")
    .setTitle("FormReceipt Settings")
    .setWidth(320);
  SpreadsheetApp.getUi().showSidebar(html);
}

function FORMRECEIPT_setupTrigger() {
  FORMRECEIPT_removeTrigger();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ScriptApp.newTrigger("FORMRECEIPT_onFormSubmit").forSpreadsheet(ss).onFormSubmit().create();
  SpreadsheetApp.getUi().alert("Trigger set. PDF will be generated automatically on form submit.");
}

function FORMRECEIPT_removeTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === "FORMRECEIPT_onFormSubmit") ScriptApp.deleteTrigger(trigger);
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
  MailApp.sendEmail({ to: adminEmail, subject: "[FormReceipt] Error occurred", body: err.message + "\n\n" + (err.stack || "") });
}

function FORMRECEIPT_showUsage() {
  const namespace = "FORMRECEIPT_" + getYearMonthString();
  const props = PropertiesService.getScriptProperties();
  const count = parseInt(props.getProperty("SEQ_" + namespace) || "0", 10);
  const license = FORMRECEIPT_getLicenseType();
  const limit = license === "pro" ? "Unlimited" : "10";
  SpreadsheetApp.getUi().alert("Monthly Usage\n\nPlan: " + license.toUpperCase() + "\nSent: " + count + " / " + limit);
}

function FORMRECEIPT_saveConfig(config) {
  const props = PropertiesService.getScriptProperties();
  props.setProperty("FORMRECEIPT_TEMPLATE_DOC_ID", config.templateDocId || "");
  props.setProperty("FORMRECEIPT_ROOT_FOLDER_ID", config.rootFolderId || "");
  props.setProperty("FORMRECEIPT_ADMIN_EMAIL", config.adminEmail || "");
  props.setProperty("FORMRECEIPT_LICENSE_KEY", config.licenseKey || "");
  return "Settings saved";
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

function getYearMonthString() { return Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy-MM"); }
function getTodayString() { return Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyyMMdd"); }

function getNextSequence(namespace) {
  const props = PropertiesService.getScriptProperties();
  const key = "SEQ_" + namespace;
  const next = parseInt(props.getProperty(key) || "0", 10) + 1;
  props.setProperty(key, String(next));
  return next;
}

function getOrCreateFolder(parentFolderId, path) {
  let folder = DriveApp.getFolderById(parentFolderId);
  for (const part of path.split("/")) {
    const iter = folder.getFoldersByName(part);
    folder = iter.hasNext() ? iter.next() : folder.createFolder(part);
  }
  return folder;
}

function saveToFolder(folderId, filename, blob) {
  return DriveApp.getFolderById(folderId).createFile(blob.setName(filename));
}

function replacePlaceholders(templateDocId, data) {
  const tempCopy = DocumentApp.openById(templateDocId).makeCopy("FormReceipt_temp_" + Date.now());
  const tempDoc = DocumentApp.openById(tempCopy.getId());
  const body = tempDoc.getBody();
  for (const [key, value] of Object.entries(data)) body.replaceText("\\{\\{" + key + "\\}\\}", value);
  tempDoc.saveAndClose();
  return tempCopy.getId();
}

function exportToPdf(docId) {
  const blob = DriveApp.getFileById(docId).getAs(MimeType.PDF);
  DriveApp.getFileById(docId).setTrashed(true);
  return blob;
}
