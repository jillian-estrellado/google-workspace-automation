// --- Helper: format number as currency ---
function formatCurrency(value) {
  if (value == null || value === '') return '';
  return Number(value).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

// --- Fee calculator ---
function calculateFee(arv, propertyType) {
  const type = String(propertyType || '').trim().toLowerCase();
  if (!propertyType || type === '') return '';

  if (/multi[\s-]?unit|multiunit|multi[\s-]?family|4\+/.test(type)) return 1250;
  if (/land|vacant|undeveloped|lot|strip/.test(type)) return 500;

  const feeTable = {
    residential: [
      { max: 29000, fee: 650 },
      { min: 30000, max: 59999, fee: 800 },
      { min: 60000, max: 99999, fee: 900 },
      { min: 100000, max: 176999, fee: 1000 },
      { min: 177000, fee: 1200 }
    ],
    commercial: [
      { max: 99999, fee: 1000 },
      { min: 100000, max: 199999, fee: 1500 },
      { min: 200000, max: 1499999, fee: 2750 },
      { min: 1500000, max: 3499999, fee: 4000 },
      { min: 3500000, max: 5999999, fee: 6000 }
    ]
  };

  let category = null;
  if (type.includes('residential')) category = 'residential';
  else if (type.includes('commercial')) category = 'commercial';
  else return '';

  for (const { min = -Infinity, max = Infinity, fee } of feeTable[category]) {
    if (arv >= min && arv <= max) return fee;
  }

  return '';
}

// --- Main function ---
function createQuotation() {
  const ui = SpreadsheetApp.getUi();

  // --- Show loading screen ---
  const loadingHtml = HtmlService.createHtmlOutput(`
    <div style="font-family: Arial, sans-serif; text-align: center; padding: 20px;">
      <h3 style="color:#2e6c80;">Generating Quotation...</h3>
      <div class="loader" style="margin:15px auto;"></div>
      <ul style="list-style:none; padding:0; margin-top:20px; font-size:14px; text-align:left; display:inline-block;">
        <li>✔ Step 1/3: Preparing data</li>
        <li>✔ Step 2/3: Creating document</li>
        <li>✔ Step 3/3: Finalizing quotation</li>
      </ul>
      <style>
        .loader { border: 6px solid #f3f3f3; border-top: 6px solid #2e6c80; border-radius: 50%; width: 30px; height: 30px; animation: spin 1s linear infinite; }
        @keyframes spin { 0% { transform: rotate(0deg); } 100% { transform: rotate(360deg); } }
      </style>
    </div>`).setWidth(350).setHeight(350);
  ui.showModalDialog(loadingHtml, 'Please Wait');

  const TEMPLATE_ID = '1QTlHb24BL16WBDK4ojmuIvrbJLHGaiXNL59PXSO2YDQ';
  const FOLDER_ID = '1wostjuXGw0rfovEFCPAL_pax2HTYvK55';
  const SHEET_ID = '1vzTzGT9Xz2iluo-SUf2orCRQ-1Z9t5GqrpIcvSnoIwE';

  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Quotations');

  const lastRow = getLastDataRow(sheet, 5);
  if (lastRow === 0) return Logger.log('No data found in the sheet.');

  const data = sheet.getRange(lastRow, 1, 1, sheet.getLastColumn()).getValues()[0];

  const newQuote = data[0];
  const clientName = data[1];
  const clientDetail1 = data[2];
  const clientDetail2 = data[3];
  const completePropertyAddress = data[4];
  const basedOn = data[5];
  const scope = data[6];

  const today = new Date();
  const formattedDate = Utilities.formatDate(today, Session.getScriptTimeZone(), 'MMMM d, yyyy');

  const propertyColumns = [
    { property: 13, assessment: 8, arv: 9, tax: 11, description: 12 },
    { property: 19, assessment: 14, arv: 15, tax: 17, description: 18 },
    { property: 25, assessment: 20, arv: 21, tax: 23, description: 24 },
    { property: 31, assessment: 26, arv: 27, tax: 29, description: 30 },
    { property: 37, assessment: 32, arv: 33, tax: 35, description: 36 }
  ];

  const propertyTexts = ['', '', '', '', ''];
  let totalArv = 0;

  for (let i = 0; i < propertyColumns.length; i++) {
    const c = propertyColumns[i];
    const subjectProperty = data[c.property];
    if (!subjectProperty) continue;

    const assessmentNumber = data[c.assessment];
    const arvRaw = data[c.arv];
    const taxCode = data[c.tax];
    const description = data[c.description];

    const arvNumber = parseFloat(String(arvRaw).replace(/[^0-9.]/g, ''));
    if (!isNaN(arvNumber)) totalArv += arvNumber;

    const parts = [];
    if (!isNaN(arvNumber)) parts.push(`ARV: $${formatCurrency(arvNumber)}`);
    if (taxCode) parts.push(taxCode);
    if (description) parts.push(description);
    const arvFormatted = parts.join(' – ');

    let block = `Subject Property: ${subjectProperty}. ${completePropertyAddress}`;
    if (assessmentNumber) block += `\nAssessment No.: ${assessmentNumber}`;
    if (arvFormatted) block += `\n${arvFormatted}`;
    propertyTexts[i] = block.trim() + '\n\n';
  }

  // --- Use only calculated fee ---
  const calculatedFee = calculateFee(totalArv, scope);
  const feeText = calculatedFee ? '$' + formatCurrency(calculatedFee) : '';

  const fileName = `Appraisal Quotation - ${completePropertyAddress}`;
  const template = DriveApp.getFileById(TEMPLATE_ID);
  const newDoc = template.makeCopy(fileName, DriveApp.getFolderById(FOLDER_ID));
  const doc = DocumentApp.openById(newDoc.getId());
  const body = doc.getBody();

  const replacements = {
    '{{QUOTENUMBER}}': String(newQuote),
    '{{DATE}}': formattedDate,
    '{{CLIENTNAME}}': clientName,
    '{{CLIENTDETAIL1}}': clientDetail1,
    '{{CLIENTDETAIL2}}': clientDetail2,
    '{{COMPLETEPROPERTYADDRESS}}': completePropertyAddress,
    '{{BASEDON}}': basedOn,
    '{{SCOPE}}': scope,
    '{{FEE}}': feeText,
    '{{PROPERTYSET1}}': propertyTexts[0],
    '{{PROPERTYSET2}}': propertyTexts[1],
    '{{PROPERTYSET3}}': propertyTexts[2],
    '{{PROPERTYSET4}}': propertyTexts[3],
    '{{PROPERTYSET5}}': propertyTexts[4]
  };

  for (const key in replacements) {
    body.replaceText(key, replacements[key]);
  }

  doc.saveAndClose();
  const fileUrl = newDoc.getUrl();
  Logger.log('Quotation created: ' + fileUrl);
  sheet.getRange(lastRow, 39).setValue(fileUrl);

  // --- Show final message ---
  const resultHtml = HtmlService.createHtmlOutput(`
    <div style="font-family: Arial, sans-serif; text-align: center; padding: 10px;">
      <h2 style="color:#2e6c80; margin-bottom: 10px;">Quotation Generated.</h2>
      <p style="margin: 10px 0; font-size: 14px;">
        Click button to open.<br>
        Don't forget to proofread and remove extra spaces.
