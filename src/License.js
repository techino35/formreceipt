/**
 * License.js - FormReceipt ライセンス管理
 */

function FORMRECEIPT_getLicenseType() {
  const props = PropertiesService.getScriptProperties();
  const licenseKey = props.getProperty("FORMRECEIPT_LICENSE_KEY");
  if (!licenseKey) return "free";
  return _validateLicenseKey(licenseKey) ? "pro" : "free";
}

function _validateLicenseKey(key) {
  if (!key) return false;
  return /^FR-[A-Z0-9]{4}-[A-Z0-9]{4}-[A-Z0-9]{4}$/.test(key.toUpperCase());
}

function FORMRECEIPT_isProUser() {
  return FORMRECEIPT_getLicenseType() === "pro";
}

function FORMRECEIPT_registerLicense(key) {
  if (!_validateLicenseKey(key)) {
    return { success: false, message: "無効なライセンスキーです", licenseType: "free" };
  }
  PropertiesService.getScriptProperties().setProperty("FORMRECEIPT_LICENSE_KEY", key.toUpperCase());
  return { success: true, message: "Proプランが有効化されました", licenseType: "pro" };
}