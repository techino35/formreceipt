/**
 * Processor.js - FormReceipt main processing class
 */

class FormReceiptProcessor {
  constructor(e) {
    this.event = e;
    this.props = PropertiesService.getScriptProperties();
    this.templateDocId = this.props.getProperty("FORMRECEIPT_TEMPLATE_DOC_ID");
    this.rootFolderId = this.props.getProperty("FORMRECEIPT_ROOT_FOLDER_ID");
    this.licenseType = FORMRECEIPT_getLicenseType();
  }

  run() {
    if (!this.templateDocId) { console.warn("FORMRECEIPT: Template Doc ID not set"); return; }
    if (this.licenseType === "free" && !this._checkFreeLimit()) { console.warn("FORMRECEIPT: Free plan monthly limit reached"); return; }
    const data = this._extractFormData();
    const pdfBlob = this._generatePdf(data);
    const savedFile = this._saveToDrive(pdfBlob, data);
    const recipientEmail = this._getRecipientEmail(data);
    if (recipientEmail) sendReceiptEmail(recipientEmail, pdfBlob, data, this.licenseType);
    this._incrementCounter();
    console.log("FORMRECEIPT: Done - " + savedFile.getUrl());
  }

  _checkFreeLimit() {
    const count = parseInt(PropertiesService.getScriptProperties().getProperty("SEQ_FORMRECEIPT_" + getYearMonthString()) || "0", 10);
    return count < 10;
  }

  _extractFormData() {
    const data = {};
    for (const r of this.event.response.getItemResponses()) {
      const ans = r.getResponse();
      data[r.getItem().getTitle()] = Array.isArray(ans) ? ans.join(", ") : String(ans || "");
    }
    data["Submission Date/Time"] = Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy/MM/dd HH:mm");
    data["Submission Date"] = Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy/MM/dd");
    data["Receipt Number"] = this._generateReceiptNumber();
    return data;
  }

  _generateReceiptNumber() {
    const ym = getYearMonthString();
    const current = parseInt(PropertiesService.getScriptProperties().getProperty("SEQ_FORMRECEIPT_" + ym) || "0", 10);
    return ym + "-" + String(current + 1).padStart(4, "0");
  }

  _generatePdf(data) {
    const tempDocId = replacePlaceholders(this.templateDocId, data);
    if (this.licenseType === "pro") this._insertLogo(tempDocId);
    return exportToPdf(tempDocId);
  }

  _insertLogo(docId) {
    const logoFolderId = this.props.getProperty("FORMRECEIPT_LOGO_FOLDER_ID");
    if (!logoFolderId) return;
    try {
      const files = DriveApp.getFolderById(logoFolderId).getFilesByType(MimeType.PNG);
      if (!files.hasNext()) return;
      const logoBlob = files.next().getBlob();
      const doc = DocumentApp.openById(docId);
      const found = doc.getBody().findText("{{logo}}");
      if (found) { const p = found.getElement().getParent().asParagraph(); p.clear(); p.appendInlineImage(logoBlob); }
      doc.saveAndClose();
    } catch (err) { console.warn("FORMRECEIPT: Logo insert failed - " + err.message); }
  }

  _saveToDrive(pdfBlob, data) {
    if (!this.rootFolderId) this.rootFolderId = DriveApp.getRootFolder().getId();
    const formName = this._getFormName();
    const seq = String(getNextSequence("FORMRECEIPT_" + getYearMonthString())).padStart(4, "0");
    const filename = getTodayString() + "_" + seq + ".pdf";
    const folder = getOrCreateFolder(this.rootFolderId, "FormReceipt/" + formName);
    return saveToFolder(folder.getId(), filename, pdfBlob);
  }

  _getFormName() { try { return this.event.source.getTitle() || "Unknown"; } catch (e) { return "Unknown"; } }

  _getRecipientEmail(data) {
    const keys = Object.keys(data).filter((k) => /email|mail/i.test(k));
    if (!keys.length) return null;
    const email = data[keys[0]];
    return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email) ? email : null;
  }

  _incrementCounter() {
    const props = PropertiesService.getScriptProperties();
    const key = "SEQ_FORMRECEIPT_" + getYearMonthString();
    props.setProperty(key, String(parseInt(props.getProperty(key) || "0", 10) + 1));
  }
}
