/**
 * This script processes incoming emails with specific labels, extracts PDF attachments,
 * saves them to a designated Google Drive folder, and logs the activity to a Google Sheet.
 */

// --- Configuration ---
const GOOGLE_SHEET_ID = "1QWQS6kwsiFJX9XdDweQBVz1Lef5uR5fJOV4P814rd98";
const PROJECTS_SHEET_NAME = "Projets";
const LOGS_SHEET_NAME = "Logs_Ingestion";
const TARGET_DRIVE_FOLDER_ID = "1TkY23yZdrpCcJKLfxCBOVitQ1Syj3miO";
const GMAIL_LABEL_TO_PROCESS = "invoices/to-process";
const GMAIL_LABEL_PROCESSED_OK = "invoices/processed/ok";
const GMAIL_LABEL_PROCESSED_ERROR = "invoices/processed/error";
const GMAIL_LABEL_PROCESSED_AUTRES = "invoices/processed/autres";
// NEW: Central list of client prefixes for easy updating.
const CLIENT_PREFIXES = ["MR OU MME", "MME", "MR."];
// NEW: Keywords for PAC identification
const PAC_KEYWORDS = {
  "PAC": ["DAIKIN", "ATLANTIC"],
  "Portails": ["GARAGE", "PORTAIL"],
  "Fenêtres": ["FENETRE", "VOLET"]
};

const MONTH_NAMES = Object.freeze({
  "janvier": "01",
  "fevrier": "02",
  "février": "02",
  "mars": "03",
  "avril": "04",
  "mai": "05",
  "juin": "06",
  "juillet": "07",
  "aout": "08",
  "août": "08",
  "septembre": "09",
  "octobre": "10",
  "novembre": "11",
  "decembre": "12",
  "décembre": "12"
});

// --- Main Processing Function ---

/**
 * Processes incoming emails, categorizes them, extracts data, and logs the activity.
 */
function processIncomingInvoices() {
  Logger.log("--- Starting Invoice Processing Run ---");

  Logger.log(`Attempting to find Gmail label with name: "${GMAIL_LABEL_TO_PROCESS}"`);
  const label = GmailApp.getUserLabelByName(GMAIL_LABEL_TO_PROCESS);

  if (!label) {
    Logger.log(`CRITICAL ERROR: Gmail label '${GMAIL_LABEL_TO_PROCESS}' not found. Please ensure it exists.`);
    return;
  }
  Logger.log(`SUCCESS: Found the Gmail label object for "${label.getName()}".`);

  let targetFolder;
  try {
    targetFolder = DriveApp.getFolderById(TARGET_DRIVE_FOLDER_ID);
  } catch (e) {
    Logger.log(`Error: Google Drive folder with ID '${TARGET_DRIVE_FOLDER_ID}' not found or inaccessible.`);
    logIngestionError("N/A", "N/A", `Drive folder error: ${e.message}`);
    return;
  }

  const threads = label.getThreads();
  Logger.log(`Found exactly ${threads.length} email thread(s) with this label.`);

  if (threads.length === 0) {
    Logger.log("INFO: No new documents to process at this time. Function will now exit.");
    return;
  }

  threads.forEach(thread => {
    const messages = thread.getMessages();
    messages.forEach(message => {
      const messageId = message.getId();
      const senderEmail = message.getFrom();
      const recipientEmail = message.getTo();
      const subject = message.getSubject();
      const messageDate = message.getDate();
      const upperCaseSubject = subject.toUpperCase();
      let documentType;

      // --- Filtering Logic ---
      if (upperCaseSubject.includes("FACTURE D 'ACOMPTE")) {
        documentType = 'deposit_invoice';
      } else if (upperCaseSubject.includes("FACTURE")) {
        documentType = 'invoice';
      } else if (upperCaseSubject.includes("AVOIR")) {
        documentType = 'credit_note';
      } else {
        documentType = 'other';
      }

      // --- Handle based on document type ---
      if (documentType === 'other') {
        Logger.log(`Skipping message with non-standard subject: "${subject}"`);
        logIngestionSkipped(messageId, subject, senderEmail, recipientEmail, messageDate);

        const autresLabel = GmailApp.getUserLabelByName(GMAIL_LABEL_PROCESSED_AUTRES);
        if (autresLabel) {
          thread.addLabel(autresLabel);
        }
        thread.removeLabel(label);
        return;
      }

      Logger.log(`Processing document type '${documentType}'. Subject: "${subject}"`);
      const attachments = message.getAttachments();
      let pdfFound = false;

      attachments.forEach(attachment => {
        const attachmentContentType = attachment.getContentType();
        const attachmentName = attachment.getName();
        const lowerTrimmedName = attachmentName.toLowerCase().trim();
        const includesPdf = lowerTrimmedName.includes('.pdf');

        if (attachmentContentType === MimeType.PDF || (attachmentContentType === 'application/octet-stream' && includesPdf)) {
          pdfFound = true;
          try {
            const fileName = attachment.getName();
            const pdfFile = targetFolder.createFile(attachment);

            const convertedDoc = Drive.Files.insert({
              title: pdfFile.getName() + '_converted',
              mimeType: MimeType.GOOGLE_DOCS,
              parents: [{
                id: targetFolder.getId()
              }]
            }, pdfFile.getBlob());
            Utilities.sleep(2000);

            const doc = DocumentApp.openById(convertedDoc.id);
            const pdfTextContent = doc.getBody().getText();

            let extractedData;
            if (documentType === 'invoice') {
              extractedData = extractDataFromPdfContent(pdfTextContent);
            } else if (documentType === 'deposit_invoice') {
              extractedData = extractDataFromDepositInvoice(pdfTextContent);
            } else { // documentType === 'credit_note'
              extractedData = extractDataFromCreditNote(pdfTextContent);
            }
            Logger.log(`Parsed data: ${JSON.stringify(extractedData)}`);

            const invoiceNumber = extractedData['Numéro facture'];
            if (invoiceNumber && invoiceNumber !== 'N/A' && isDuplicateInvoice(invoiceNumber)) {
              const errorMessage = `Duplicate document error: N° ${invoiceNumber} already exists.`;
              Logger.log(errorMessage);
              logIngestionError(messageId, fileName, errorMessage, senderEmail, recipientEmail, messageDate);

              const errorLabel = GmailApp.getUserLabelByName(GMAIL_LABEL_PROCESSED_ERROR);
              if (errorLabel) {
                thread.addLabel(errorLabel);
                thread.removeLabel(label);
              }
              Drive.Files.remove(convertedDoc.id);
              Drive.Files.remove(pdfFile.getId());
              return;
            }

            appendToGoogleSheet(messageId, new Date(), extractedData);
            Drive.Files.remove(convertedDoc.id);
            Drive.Files.remove(pdfFile.getId());
            logIngestionSuccess(messageId, fileName, senderEmail, recipientEmail, messageDate, pdfFile.getId());

            const processedLabel = GmailApp.getUserLabelByName(GMAIL_LABEL_PROCESSED_OK);
            if (processedLabel) {
              thread.addLabel(processedLabel);
              thread.removeLabel(label);
            }
            return;
          } catch (e) {
            Logger.log(`Error during PDF processing: ${e.message}`);
            logIngestionError(messageId, attachmentName, `Full PDF processing error: ${e.message}`, senderEmail, recipientEmail, messageDate);

            const errorLabel = GmailApp.getUserLabelByName(GMAIL_LABEL_PROCESSED_ERROR);
            if (errorLabel) {
              thread.addLabel(errorLabel);
              thread.removeLabel(label);
            }
            return;
          }
        }
      });

      if (!pdfFound) {
        Logger.log(`No PDF attachment found in message ID: ${messageId}.`);
        logIngestionError(messageId, "N/A", "No PDF attachment found", senderEmail, recipientEmail, messageDate);

        const errorLabel = GmailApp.getUserLabelByName(GMAIL_LABEL_PROCESSED_ERROR);
        if (errorLabel) {
          thread.addLabel(errorLabel);
          thread.removeLabel(label);
        }
      }
    });
  });
  Logger.log("--- Finished Invoice Processing Run ---");
}


// --- Parsing Functions ---

/**
 * Extracts data from a standard invoice PDF content.
 */
function extractDataFromPdfContent(pdfContent) {
  const data = {};

  // Helper function to find a value that appears anywhere after a label
  function findValueAfterLabel(label, content) {
    const regex = new RegExp('(?:' + label + ')[\\s\\S]*?([\\d\\s]+,\\d{2})', 'i');
    const match = content.match(regex);
    if (match && match[1]) {
      return normalizeAmount(match[1]);
    }
    return 'N/A';
  }

  // --- Financial Totals Extraction ---
  Logger.log("--- Starting Financial Extraction ---");
  data['HT'] = findValueAfterLabel('(?:MONTANT TOTAL|Total) HT', pdfContent);
  data['TVA'] = findValueAfterLabel('(?:MONTANT TOTAL|Total) TVA', pdfContent);
  data['TTC'] = findValueAfterLabel('(?:MONTANT TOTAL|Total) TTC', pdfContent);
  Logger.log(`Extracted Totals: HT=${data['HT']}, TVA=${data['TVA']}, TTC=${data['TTC']}`);

  // --- NEW: VAT Rate Extraction based on Codes in "T" column ---
  Logger.log("--- Starting VAT Rate Code Extraction ---");
  const vatCodeMap = {
    '2': '5.5%',
    '6': '10%',
    '7': '20%'
  };
  const foundRates = new Set();
  // This regex finds single digits at the very end of a line, typical for the 'T' column.
  const codeRegex = /\s+(\d)\s*$/gm;
  const allCodeMatches = [...pdfContent.matchAll(codeRegex)];

  Logger.log(`Found ${allCodeMatches.length} potential VAT codes in the item list.`);
  for (const match of allCodeMatches) {
    const code = match[1];
    if (vatCodeMap[code]) {
      foundRates.add(vatCodeMap[code]);
    }
  }

  if (foundRates.size > 0) {
    // Sort the rates numerically for consistent output (e.g., 5.5% & 10%)
    const sortedRates = Array.from(foundRates).sort((a, b) => parseFloat(a) - parseFloat(b));
    data['Taux TVA'] = sortedRates.join(' & ');
  } else {
    data['Taux TVA'] = 'N/A';
  }
  Logger.log(`Extracted Rates from Codes: ${data['Taux TVA']}`);
  Logger.log("--- Finished Financial Extraction ---");

  // --- Extraction for other fields ---
  const clientRegex = new RegExp('(?:' + CLIENT_PREFIXES.join('|').replace(/\./g, '\\.') + ')\\s+(.*)', 'i');
  const regexMap = {
    "Numéro facture": /Facture N°\s*([A-Z0-9\/\-]+)/,
    "Client": clientRegex,
    "Vendeur": /Votre contact\s*:\s*(.*)/,
    "Acompte": /Montant versé\s*:\s*Chèque.*?de\s+([\d,.]+)\s*€/i,
    "Mode paiement": /Montant versé\s*:\s*(Chèque)/i,
    "Solde": /NET A PAYER\s*:\s*([\d\s,.]+)\s*€?/i,
    "InvoiceDate": /Le\s+(\d{1,2}\s+\w+\s+\d{4})/i,
  };

  const invoiceDateMatch = pdfContent.match(regexMap["InvoiceDate"]);
  if (invoiceDateMatch && invoiceDateMatch[1]) {
    const dateStringFr = invoiceDateMatch[1].trim();
    const parts = dateStringFr.split(' ');
    if (parts.length === 3) {
      const day = parts[0];
      const month = MONTH_NAMES[parts[1].toLowerCase()];
      const year = parts[2];
      if (day && month && year) {
        data["Échéance"] = convertDateToYYYYMMDD(`${day}/${month}/${year}`);
        data["Date paiement"] = convertDateToYYYYMMDD(`${day}/${month}/${year}`);
      }
    }
  } else {
    data["Échéance"] = "N/A";
    data["Date paiement"] = "N/A";
  }

  for (const field of ["Numéro facture", "Client", "Vendeur", "Mode paiement", "Solde"]) {
    const match = pdfContent.match(regexMap[field]);
    if (match && match[1]) data[field] = match[1].trim();
    else if (!data[field]) data[field] = "N/A";
  }

  const acompteMatch = pdfContent.match(regexMap["Acompte"]);
  if (acompteMatch && acompteMatch[1]) {
    let num = parseFloat(normalizeAmount(acompteMatch[1]));
    if (num > 0) num = -num;
    data['Acompte'] = num.toFixed(2);
  } else {
    data['Acompte'] = '0.00';
  }

  data['Solde'] = normalizeAmount(data['Solde']);
  data['Client'] = formatClientName(data['Client']);

  // Fallback calculation for TTC if it wasn't found directly
  if (data['TTC'] === 'N/A' && data['HT'] !== 'N/A' && data['TVA'] !== 'N/A') {
    data['TTC'] = (parseFloat(data['HT']) + parseFloat(data['TVA'])).toFixed(2);
  }

  data['PAC'] = determinePacValue(pdfContent);
  return data;
}





/**
 * Extracts data from a credit note ("Avoir") PDF content.
 */
function extractDataFromCreditNote(pdfContent) {
  const data = {};

  const clientRegex = new RegExp('(?:' + CLIENT_PREFIXES.join('|') + ')\\s+(.*)', 'i');
  const regexMap = {
    "Numéro facture": /Avoir N°\s*([A-Z0-9\/\-]+)/,
    "Client": clientRegex,
    "Vendeur": /Votre contact\s*:\s*(.*)/,
    "InvoiceDate": /Le\s+(\d{1,2}\s+\w+\s+\d{4})/i,
  };

  const invoiceDateMatch = pdfContent.match(regexMap["InvoiceDate"]);
  if (invoiceDateMatch && invoiceDateMatch[1]) {
    const dateStringFr = invoiceDateMatch[1].trim();
    const parts = dateStringFr.split(' ');
    if (parts.length === 3) {
      const day = parts[0];
      const month = MONTH_NAMES[parts[1].toLowerCase()];
      const year = parts[2];
      if (day && month && year) {
        data["InvoiceDate"] = convertDateToYYYYMMDD(`${day}/${month}/${year}`);
      } else {
        data["InvoiceDate"] = dateStringFr;
      }
    } else {
      data["InvoiceDate"] = dateStringFr;
    }
  } else {
    data["InvoiceDate"] = "N/A";
  }

  const summaryRegex = /\b7\s+([\d,.]+?)\s+(-?[\d\s,.]+?)\s+(-?[\d\s,.]+)/;
  const summaryMatch = pdfContent.match(summaryRegex);
  if (summaryMatch) {
    const rawVatRate = summaryMatch[1].trim();
    const normalizedRate = normalizeAmount(rawVatRate);
    const rateAsNumber = parseFloat(normalizedRate);
    data['Taux TVA'] = String(rateAsNumber) + '%';
    data['HT'] = normalizeAmount(summaryMatch[2].trim());
    data['TVA'] = normalizeAmount(summaryMatch[3].trim());
  } else {
    data['Taux TVA'] = 'N/A';
    data['HT'] = 'N/A';
    data['TVA'] = 'N/A';
  }

  for (const field in regexMap) {
    if (field === "InvoiceDate") continue;
    const match = pdfContent.match(regexMap[field]);
    if (match && match[1]) {
      data[field] = match[1].trim();
    } else if (!data[field]) {
      data[field] = "N/A";
    }
  }

  data['Client'] = formatClientName(data['Client']);

  if (data['HT'] !== 'N/A' && data['TVA'] !== 'N/A') {
    const htFloat = parseFloat(data['HT']);
    const tvaFloat = parseFloat(data['TVA']);
    if (!isNaN(htFloat) && !isNaN(tvaFloat)) {
      data['TTC'] = (htFloat + tvaFloat).toFixed(2);
    } else {
      data['TTC'] = 'N/A';
    }
  } else {
    data['TTC'] = 'N/A';
  }

  data['Acompte'] = '0';
  data['Solde'] = '0';
  data['Échéance'] = 'N/A';
  data['Date paiement'] = 'N/A';
  data['Mode paiement'] = 'Avoir';
  data['PAC'] = determinePacValue(pdfContent);
  return data;
}

/**
 * Extracts data from a deposit invoice ("Facture d'acompte") PDF content.
 */
function extractDataFromDepositInvoice(pdfContent) {
  const data = {};

  const clientRegex = new RegExp('(?:' + CLIENT_PREFIXES.join('|') + ')\\s+(.*)', 'i');
  const regexMap = {
    "Numéro facture": /Facture N°\s*([A-Z0-9\/\-]+)/,
    "Client": clientRegex,
    "TTC": /MONTANT TOTAL TTC\s*:\s*([\d\s,.]+) €/,
    "Vendeur": /Votre contact\s*:\s*(.*)/,
    "InvoiceDate": /Le\s+(\d{1,2}\s+\w+\s+\d{4})/i,
  };

  const invoiceDateMatch = pdfContent.match(regexMap["InvoiceDate"]);
  if (invoiceDateMatch && invoiceDateMatch[1]) {
    const dateStringFr = invoiceDateMatch[1].trim();
    const parts = dateStringFr.split(' ');
    if (parts.length === 3) {
      const day = parts[0];
      const month = MONTH_NAMES[parts[1].toLowerCase()];
      const year = parts[2];
      if (day && month && year) {
        data["InvoiceDate"] = convertDateToYYYYMMDD(`${day}/${month}/${year}`);
      } else {
        data["InvoiceDate"] = dateStringFr;
      }
    } else {
      data["InvoiceDate"] = dateStringFr;
    }
  } else {
    data["InvoiceDate"] = "N/A";
  }

  const htTvaRegex = /MONTANT TOTAL HT\s+MONTANT TVA\s+([\d\s,.]+) €\s+([\d\s,.]+) €/;
  const financialMatch = pdfContent.match(htTvaRegex);
  if (financialMatch) {
    data['HT'] = normalizeAmount(financialMatch[1].trim());
    data['TVA'] = normalizeAmount(financialMatch[2].trim());
  } else {
    data['HT'] = 'N/A';
    data['TVA'] = 'N/A';
  }

  const summaryRegex = /\b2\s+([\d,.]+)/;
  const summaryMatch = pdfContent.match(summaryRegex);
  if (summaryMatch) {
    const rawVatRate = summaryMatch[1].trim();
    const allowedVatRates = ['10,00', '10', '20,00', '20', '5,5', '5,50'];
    if (allowedVatRates.includes(rawVatRate)) {
      const normalizedRate = normalizeAmount(rawVatRate);
      const rateAsNumber = parseFloat(normalizedRate);
      data['Taux TVA'] = String(rateAsNumber) + '%';
    } else {
      data['Taux TVA'] = 'N/A';
    }
  } else {
    data['Taux TVA'] = 'N/A';
  }

  for (const field in regexMap) {
    if (field === "InvoiceDate") continue;
    const match = pdfContent.match(regexMap[field]);
    if (match && match[1]) {
      let value = match[1].trim();
      if (field === 'TTC') {
        value = normalizeAmount(value);
      }
      data[field] = value;
    } else if (!data[field]) {
      data[field] = "N/A";
    }
  }

  data['Client'] = formatClientName(data['Client']);

  data['Acompte'] = data['TTC'];
  data['Solde'] = 'N/A';
  data['Mode paiement'] = 'Acompte';
  data['Échéance'] = 'N/A';
  data['Date paiement'] = 'N/A';
  data['PAC'] = determinePacValue(pdfContent);
  return data;
}


// --- Sheet and Logging Functions ---

/**
 * Appends the extracted data to the 'Projets' Google Sheet.
 */
function appendToGoogleSheet(idMessage, date, data) {
  const spreadsheet = SpreadsheetApp.openById(GOOGLE_SHEET_ID);
  const sheet = spreadsheet.getSheetByName(PROJECTS_SHEET_NAME);
  if (!sheet) {
    Logger.log(`Error: '${PROJECTS_SHEET_NAME}' sheet not found.`);
    return;
  }

  const uuid = Utilities.getUuid();
  const dateOnlyOptions = {
    day: 'numeric',
    month: 'long',
    year: 'numeric',
    timeZone: 'UTC'
  };

  let avancement = "Dû";
  if (data['Mode paiement'] === 'Avoir') {
    avancement = 'Avoir';
  } else if (data['Mode paiement'] === 'Acompte') {
    avancement = 'Acompte';
  }

  function formatFrenchDate(dateString) {
    if (!dateString || dateString === 'N/A') return 'N/A';
    try {
      const dateObj = new Date(dateString);
      const formattedDate = dateObj.toLocaleDateString('fr-FR', dateOnlyOptions);
      return "'" + formattedDate;
    } catch (e) {
      return dateString;
    }
  }

  function formatNumberFrench(value) {
    if (value === null || value === undefined || value === 'N/A') return 'N/A';
    return String(value).replace('.', ',') + ' €';
  }

  const rowData = [
    uuid,
    "'" + date.toLocaleDateString('fr-FR', {
      day: 'numeric',
      month: 'long',
      year: 'numeric',
      timeZone: Session.getScriptTimeZone()
    }),
    data['Numéro facture'],
    data['Client'],
    String(data['Taux TVA']).replace('.', ','),
    formatNumberFrench(data['TTC']),
    formatNumberFrench(data['HT']),
    formatNumberFrench(data['TVA']),
    formatNumberFrench(data['Acompte']),
    formatNumberFrench(data['Solde']),
    formatFrenchDate(data['Échéance']),
    formatFrenchDate(data['Date paiement']),
    data['Mode paiement'],
    "", // Column N (Achats HT)
    "", // Column O (Marge HT)
    data['Vendeur'] || "Non Assigné",
    data['PAC'] || "", // Column Q (PAC ?)
    avancement
  ];

  sheet.appendRow(rowData);
  Logger.log(`Row appended to '${PROJECTS_SHEET_NAME}' sheet.`);

  const newRow = sheet.getLastRow();

  sheet.getRange(newRow, 2).setNumberFormat('@'); // Column B (Date)
  sheet.getRange(newRow, 5, 1, 8).setNumberFormat('@'); // Columns E through L
  Logger.log(`Forced Plain Text format for key columns in row ${newRow}.`);

  if (data['Mode paiement'] !== 'Avoir' && data['Mode paiement'] !== 'Acompte') {
    const margeHtCell = sheet.getRange(newRow, 15); // Column O
    margeHtCell.setFormula(`=G${newRow}-N${newRow}`);
    Logger.log(`Set Marge HT formula for row ${newRow}.`);
  } else {
    Logger.log(`Skipped Marge HT formula for special invoice types.`);
  }
}

/**
 * Logs successful ingestion to the Logs_Ingestion sheet.
 */
function logIngestionSuccess(messageId, fileName, senderEmail, recipientEmail, messageDate, driveFileId) {
  const spreadsheet = SpreadsheetApp.openById(GOOGLE_SHEET_ID);
  const logSheet = spreadsheet.getSheetByName(LOGS_SHEET_NAME);
  if (!logSheet) return;
  logSheet.appendRow([new Date(), messageId, fileName, senderEmail, recipientEmail, Utilities.formatDate(messageDate, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss"), driveFileId, "SUCCESS", ""]);
}

/**
 * Logs ingestion errors to the Logs_Ingestion sheet.
 */
function logIngestionError(messageId, fileName, errorMessage, senderEmail, recipientEmail, messageDate) {
  const spreadsheet = SpreadsheetApp.openById(GOOGLE_SHEET_ID);
  const logSheet = spreadsheet.getSheetByName(LOGS_SHEET_NAME);
  if (!logSheet) return;
  logSheet.appendRow([new Date(), messageId, fileName || "N/A", senderEmail || "N/A", recipientEmail || "N/A", messageDate ? Utilities.formatDate(messageDate, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss") : "N/A", "N/A", "ERROR", errorMessage]);
}

/**
 * Logs intentionally skipped emails to the Logs_Ingestion sheet.
 */
function logIngestionSkipped(messageId, subject, senderEmail, recipientEmail, messageDate) {
  const spreadsheet = SpreadsheetApp.openById(GOOGLE_SHEET_ID);
  const logSheet = spreadsheet.getSheetByName(LOGS_SHEET_NAME);
  if (!logSheet) return;
  logSheet.appendRow([new Date(), messageId, "N/A", senderEmail, recipientEmail, Utilities.formatDate(messageDate, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss"), "N/A", "SKIPPED", `Subject matched filter: "${subject}"`]);
}


// --- Helper Functions ---

/**
 * Checks if an invoice number already exists in the sheet.
 */
function isDuplicateInvoice(invoiceNumber) {
  const spreadsheet = SpreadsheetApp.openById(GOOGLE_SHEET_ID);
  const sheet = spreadsheet.getSheetByName(PROJECTS_SHEET_NAME);
  if (sheet) {
    const values = sheet.getRange("C2:C").getValues();
    for (let i = 0; i < values.length; i++) {
      if (values[i][0] == invoiceNumber) {
        return true;
      }
    }
  }
  return false;
}

/**
 * Formats the raw client name by removing prefixes, numbers, and limiting word count.
 * @param {string} rawName The client name extracted from the PDF.
 * @returns {string} The formatted client name.
 */
function formatClientName(rawName) {
  Logger.log(`--- Starting formatClientName ---`);
  Logger.log(`1. Raw input name: "${rawName}"`);

  if (!rawName || typeof rawName !== 'string') {
    Logger.log(`Input is invalid. Returning 'N/A'.`);
    return 'N/A';
  }

  let formattedName = rawName.trim();
  Logger.log(`2. After initial trim: "${formattedName}"`);

  for (const prefix of CLIENT_PREFIXES) {
    // CORRECTED: Escape any special regex characters (like '.') in the prefix.
    const escapedPrefix = prefix.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');

    const regex = new RegExp('^' + escapedPrefix + '\\s*', 'i');
    if (regex.test(formattedName)) {
      formattedName = formattedName.replace(regex, '');
      Logger.log(`3. After removing prefix "${prefix}": "${formattedName}"`);
      break;
    }
  }

  let words = formattedName.split(/\s+/);
  if (words.length > 3) {
    formattedName = words.slice(0, 3).join(' ');
    Logger.log(`4. After limiting to 3 words: "${formattedName}"`);
  }

  formattedName = formattedName.replace(/\d/g, '');
  Logger.log(`5. After removing numbers: "${formattedName}"`);

  words = formattedName.trim().split(/\s+/);
  if (words.length > 0) {
    formattedName = words[0];
    Logger.log(`6. After taking first word: "${formattedName}"`);
  }

  const finalName = formattedName.trim();
  Logger.log(`7. Final formatted name: "${finalName}"`);
  Logger.log(`--- Finished formatClientName ---`);
  return finalName;
}


/**
 * Converts a DD/MM/YYYY date string to YYYY-MM-DD.
 */
function convertDateToYYYYMMDD(dateString) {
  if (!dateString) return "";
  const parts = dateString.split('/');
  if (parts.length === 3) {
    return `${parts[2]}-${parts[1]}-${parts[0]}`;
  }
  return dateString;
}

/**
 * Normalizes an amount string for calculations.
 */
function normalizeAmount(amountString) {
  if (!amountString) return "";
  return amountString.replace(/\s/g, '').replace(',', '.').replace('€', '').trim();
}

/**
 * Determines the PAC value based on keywords in the PDF content.
 */
function determinePacValue(pdfContent) {
  const upperCaseContent = pdfContent.toUpperCase();
  for (const pacValue in PAC_KEYWORDS) {
    for (const keyword of PAC_KEYWORDS[pacValue]) {
      if (upperCaseContent.includes(keyword)) {
        return pacValue;
      }
    }
  }
  return "N/A"; // Default value if no keywords are found
}


// --- Whiteboard Functions (Not related to invoice processing) ---

function getWhiteboardConfig() {
  const ss = SpreadsheetApp.openById(GOOGLE_SHEET_ID);
  const sheet = ss.getSheetByName("Date affichage");
  if (!sheet) {
    return {
      affiche: true,
      duration: 60,
      week: null
    };
  }
  const datesData = sheet.getDataRange().getValues();
  const headers = datesData[0] || [];
  const row = datesData[1] || [];
  const idxAff = headers.indexOf('Affiché');
  const idxDur = headers.indexOf('Durée');
  const idxWeek = headers.indexOf('Semaine choisie');
  const affiche = idxAff >= 0 && (row[idxAff] === true || String(row[idxAff]).toUpperCase() === 'TRUE');
  const duration = idxDur >= 0 ? parseInt(row[idxDur], 10) || 60 : 60;
  const week = idxWeek >= 0 ? parseInt(row[idxWeek], 10) || null : null;
  return {
    affiche,
    duration,
    week
  };
}

function getAfficheFlag() {
  const ss = SpreadsheetApp.openById(GOOGLE_SHEET_ID);
  const sheet = ss.getSheetByName("Date affichage");
  if (!sheet) {
    return true;
  }
  const datesData = sheet.getDataRange().getValues();
  const headers = datesData[0] || [];
  const row = datesData[1] || [];
  const idxAff = headers.indexOf('Affiché');
  return idxAff >= 0 && (row[idxAff] === true || String(row[idxAff]).toUpperCase() === 'TRUE');
}

function fetchSheetData(sheetName) {
  const ss = SpreadsheetApp.openById(GOOGLE_SHEET_ID);
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error(`Sheet not found: ${sheetName}`);
  return sheet.getDataRange().getValues();
}

function getProjects() {
  const data = fetchSheetData(PROJECTS_SHEET_NAME);
  const headers = data[0];
  const colIndices = {
    date: headers.indexOf('Date'),
    client: headers.indexOf('Client'),
    ht: headers.indexOf('HT'),
    margeHt: headers.indexOf('Marge HT'),
    vendeur: headers.indexOf('Vendeur')
  };
  const projects = [];
  data.slice(1).forEach(r => {
    const projectDate = r[colIndices.date];
    const client = r[colIndices.client];
    const ht = parseFloat(String(r[colIndices.ht] || '0').replace(',', '.')) || 0;
    const margeHt = parseFloat(String(r[colIndices.margeHt] || '0').replace(',', '.')) || 0;
    const vendeur = r[colIndices.vendeur];
    if (projectDate instanceof Date && client && vendeur) {
      projects.push({
        date: normalizeDate(projectDate),
        client: client,
        ht: ht,
        margeHt: margeHt,
        vendeur: vendeur
      });
    }
  });
  return projects;
}

function getUniqueSalespeople() {
  const data = fetchSheetData(PROJECTS_SHEET_NAME);
  const headers = data[0];
  const vendeurIndex = headers.indexOf('Vendeur');
  if (vendeurIndex === -1) {
    return [];
  }
  const salespeople = new Set();
  data.slice(1).forEach(row => {
    const vendeur = row[vendeurIndex];
    if (vendeur && typeof vendeur === 'string' && vendeur.trim() !== '') {
      salespeople.add(vendeur.trim());
    }
  });
  return Array.from(salespeople).sort();
}

function getCurrentWeekMonday() {
  const today = new Date();
  const day = today.getDay();
  const diff = today.getDate() - day + (day === 0 ? -6 : 1);
  const monday = new Date(today.setDate(diff));
  return normalizeDate(monday);
}

function normalizeDate(d) {
  if (!(d instanceof Date) || isNaN(d)) return null;
  return new Date(d.getFullYear(), d.getMonth(), d.getDate());
}

const WEEK_DAYS = ['Lundi', 'Mardi', 'Mercredi', 'Jeudi', 'Vendredi'];

function buildWeeklyWhiteboardData(projects, salespeople) {
  const currentMonday = getCurrentWeekMonday();
  const weekDates = [];
  for (let i = 0; i < 5; i++) {
    const day = new Date(currentMonday);
    day.setDate(currentMonday.getDate() + i);
    weekDates.push(normalizeDate(day));
  }
  const matrix = {};
  WEEK_DAYS.forEach(day => {
    matrix[day] = {};
    salespeople.forEach(sp => {
      matrix[day][sp] = [];
    });
  });
  const marginTotals = {};
  salespeople.forEach(sp => {
    marginTotals[sp] = 0;
  });
  projects.forEach(p => {
    if (p.date) {
      const projectDayIndex = weekDates.findIndex(d => d.getTime() === p.date.getTime());
      if (projectDayIndex !== -1) {
        const dayName = WEEK_DAYS[projectDayIndex];
        if (matrix[dayName] && matrix[dayName][p.vendeur]) {
          matrix[dayName][p.vendeur].push({
            client: p.client,
            ht: p.ht
          });
          marginTotals[p.vendeur] += p.margeHt;
        }
      }
    }
  });
  return {
    matrix,
    marginTotals,
    weekDates
  };
}

function getWhiteboardData() {
  const config = getWhiteboardConfig();
  const projects = getProjects();
  const salespeople = getUniqueSalespeople();
  const {
    matrix,
    marginTotals,
    weekDates
  } = buildWeeklyWhiteboardData(projects, salespeople);
  return {
    config: config,
    salespeople: salespeople,
    weekDays: WEEK_DAYS,
    weekDates: weekDates.map(d => d.toLocaleDateString('fr-FR', {
      day: '2-digit',
      month: '2-digit'
    })),
    marginTotals: marginTotals,
    matrix: matrix,
    currentWeekNumber: getWeekNumber(new Date())
  };
}

function getWeekNumber(d) {
  d = new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()));
  d.setUTCDate(d.getUTCDate() + 4 - (d.getUTCDay() || 7));
  const yearStart = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
  const weekNo = Math.ceil((((d - yearStart) / 86400000) + 1) / 7);
  return weekNo;
}

function doGet(e) {
  const htmlTemplate = HtmlService.createTemplateFromFile('Whiteboard');
  const config = getWhiteboardConfig();
  htmlTemplate.initialAffiche = config.affiche;
  htmlTemplate.durationMin = config.duration;
  const htmlOutput = htmlTemplate.evaluate();
  htmlOutput.setTitle('Tableau Blanc Hebdomadaire');
  return htmlOutput;
}
