/**
 * Processor.js - FormReceipt メイン処理クラス
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
    if (!this.templateDocId) { console.warn("FORMRECEIPT: テンプレートDocIDが未設定です"); return; }
    if (this.licenseType === "free" && !this._checkFreeLimit()) { console.warn("FORMRECEIPT: Free版の月10件制限に達しました"); return; }

    const data = this._extractFormData();
    const pdfBlob = this._generatePdf(data);
    const savedFile = this._saveToDrive(pdfBlob, data);

    const recipientEmail = this._getRecipientEmail(data);
    if (recipientEmail) sendReceiptEmail(recipientEmail, pdfBlob, data, this.licenseType);

    this._incrementCounter();
    console.log("FORMRECEIPT: 処理完了 - " + savedFile.getUrl());
  }

  _checkFreeLimit() {
    const namespace = "FORMRECEIPT_" + getYearMonthString();
    const count = parseInt(PropertiesService.getScriptProperties().getProperty("SEQ_" + namespace) || "0", 10);
    return count < 10;
  }

  _extractFormData() {
    const data = {};
    const responses = this.event.response.getItemResponses();
    for (const itemResponse of responses) {
      const title = itemResponse.getItem().getTitle();
      const answer = itemResponse.getResponse();
      data[title] = Array.isArray(answer) ? answer.join(", ") : String(answer || "");
    }
    data["送信日時"] = Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy/MM/dd HH:mm");
    data["送信日"] = Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy/MM/dd");
    data["受付番号"] = this._generateReceiptNumber();
    return data;
  }

  _generateReceiptNumber() {
    const ym = getYearMonthString();
    const namespace = "FORMRECEIPT_" + ym;
    const current = parseInt(PropertiesService.getScriptProperties().getProperty("SEQ_" + namespace) || "0", 10);
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
      const folder = DriveApp.getFolderById(logoFolderId);
      const files = folder.getFilesByType(MimeType.PNG);
      if (!files.hasNext()) return;
      const logoBlob = files.next().getBlob();
      const doc = DocumentApp.openById(docId);
      const body = doc.getBody();
      const logoPlaceholder = body.findText("{{ロゴ}}");
      if (logoPlaceholder) {
        const element = logoPlaceholder.getElement().getParent();
        element.asParagraph().clear();
        element.asParagraph().appendInlineImage(logoBlob);
      }
      doc.saveAndClose();
    } catch (err) { console.warn("FORMRECEIPT: ロゴ挿入失敗 - " + err.message); }
  }

  _saveToDrive(pdfBlob, data) {
    if (!this.rootFolderId) this.rootFolderId = DriveApp.getRootFolder().getId();
    const formName = this._getFormName();
    const today = getTodayString();
    const namespace = "FORMRECEIPT_" + getYearMonthString();
    const seq = String(getNextSequence(namespace)).padStart(4, "0");
    const filename = today + "_" + seq + ".pdf";
    const targetFolder = getOrCreateFolder(this.rootFolderId, "FormReceipt/" + formName);
    return saveToFolder(targetFolder.getId(), filename, pdfBlob);
  }

  _getFormName() {
    try { return this.event.source.getTitle() || "Unknown"; } catch (err) { return "Unknown"; }
  }

  _getRecipientEmail(data) {
    const emailKeys = Object.keys(data).filter(k => /メール|email|mail/i.test(k));
    if (emailKeys.length === 0) return null;
    const email = data[emailKeys[0]];
    return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email) ? email : null;
  }

  _incrementCounter() {
    const namespace = "FORMRECEIPT_" + getYearMonthString();
    const props = PropertiesService.getScriptProperties();
    const key = "SEQ_" + namespace;
    const current = parseInt(props.getProperty(key) || "0", 10);
    props.setProperty(key, String(current + 1));
  }
}