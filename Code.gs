// ============================================
// Avadh Millana Planning - Entry Form (Web App)
// Google Apps Script - Supersonic Group
// ============================================
//
// FUTURE ME CHANGES KARNE KE LIYE (isko hi edit karo, HTML nahi):
//   - Naya field add/remove/reorder karna ho  -> neeche FIELDS array edit karo
//   - Sheet ka tab naam badalna ho            -> CONFIG.SHEET_NAME edit karo
//
// FIELDS array hi form (kaunse box dikhenge) aur Sheet me column (kis order
// me data save hoga) dono control karta hai. Ek field add karoge to woh
// form me apne aap dikhega - bas Sheet me bhi wahi naam ka column add karna
// mat bhoolna (FIELDS jis order me hai, Sheet me A,B,C... us order me hi
// hone chahiye - is sheet me S.No column NAHI hai).

const CONFIG = {
  SHEET_NAME: "Sheet1",   // Sheet tab ka naam. Match na mile to script apne aap pehli tab use kar legi.
  FORM_TITLE: "Avadh Millana Planning - Entry Form",
  COMPANY_NAME: "Supersonic Group"
};

// Har field: key = Sheet column ka header + data ka naam, label = form pe dikhega,
// type = "text" | "textarea" | "number" | "date", required = true/false,
// voice = true matlab mic button milega (bolke type kar sakte ho).
const FIELDS = [
  { key: "weeklyPlanning",       label: "7 Days Planning",     type: "textarea", required: false, voice: true },
  { key: "fifteenDaysPlanning",  label: "15 Days Planning",    type: "textarea", required: false, voice: true },
  { key: "materialPurchase",     label: "Material Purchase",   type: "textarea", required: false, voice: true }
];

// ============================================
// WEB APP ENTRY POINTS
// ============================================
//
// Ye script 2 tarah se use hoti hai:
//  1) Seedha /exec URL kholo -> yahi doGet() wala form dikhta hai (typing works,
//     par is form ke andar voice/mic Google ki apni sandboxing ki wajah se kaam
//     nahi karta - Google ki limitation hai, hamare code ki nahi).
//  2) standalone/avadh-millana-form.html (GitHub Pages pe host hoti hai) -> wahan
//     mic bhi kaam karta hai, kyunki wo Google ke sandbox ke bahar hai. Wo page
//     is hi script se doGet(?action=fields) se form-fields mangata hai aur
//     doPost() se data save karta hai.

function doGet(e) {
  if (e && e.parameter && e.parameter.action === "fields") {
    return jsonOutput_({ fields: FIELDS, formTitle: CONFIG.FORM_TITLE, companyName: CONFIG.COMPANY_NAME });
  }

  var template = HtmlService.createTemplateFromFile("Form");
  template.formTitle = CONFIG.FORM_TITLE;
  template.companyName = CONFIG.COMPANY_NAME;
  template.fieldsJson = JSON.stringify(FIELDS);

  return template.evaluate()
    .setTitle(CONFIG.FORM_TITLE)
    .addMetaTag("viewport", "width=device-width, initial-scale=1")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// GitHub Pages wala standalone form isi ko call karta hai (fetch POST) data save karne ke liye.
function doPost(e) {
  try {
    var formData = JSON.parse(e.postData.contents);
    var result = saveEntry_(formData);
    return jsonOutput_(result);
  } catch (err) {
    return jsonOutput_({ success: false, error: err.message });
  }
}

function jsonOutput_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// HTML templating ke liye - Style.html / ClientScript.html ko Form.html me include karta hai.
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ============================================
// SHEET HELPER
// ============================================
function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(CONFIG.SHEET_NAME) || ss.getSheets()[0];
}

// ============================================
// FORM SUBMIT
// ============================================
// google.script.run se (Apps Script ke apne /exec form se) yahi call hota hai.
function submitForm(formData) {
  return saveEntry_(formData);
}

// Dono submitForm() aur doPost() (GitHub Pages wala standalone form) yahi
// shared logic use karte hai, taaki validation/Sheet-save ek hi jagah ho.
function saveEntry_(formData) {
  var sheet = getSheet_();

  // Required fields check.
  for (var i = 0; i < FIELDS.length; i++) {
    var f = FIELDS[i];
    if (f.required && (!formData[f.key] || String(formData[f.key]).trim() === "")) {
      throw new Error(f.label + " zaroori hai (required).");
    }
  }

  // Row banate hain: FIELDS ke order me, sheet ke A,B,C... columns match karte hain.
  var row = FIELDS.map(function (f) {
    return formData[f.key] || "";
  });

  sheet.appendRow(row);

  return {
    success: true,
    row: sheet.getLastRow()
  };
}
