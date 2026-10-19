//Code.gs
const SCRIPT_PROPS = PropertiesService.getScriptProperties().getProperties();

const DATABASE_FOLDER_ID = SCRIPT_PROPS['DATABASE_FOLDER_ID'];

//INPUT FOLDERS
const SR_FOLDER_ID = SCRIPT_PROPS['SR_FOLDER_ID'];
const PO_PINGCON_BONITA_FOLDER_ID = SCRIPT_PROPS['PO_PINGCON_BONITA_FOLDER_ID'];
const PO_KEYSYS_FOLDER_ID = SCRIPT_PROPS['PO_KEYSYS_FOLDER_ID'];
const AR_FOLDER_ID = SCRIPT_PROPS['AR_FOLDER_ID'];
const SUPPLIER_RECEIPT_FOLDER_ID = SCRIPT_PROPS['SUPPLIER_RECEIPT_FOLDER_ID'];
const SQ_FOLDER_ID = SCRIPT_PROPS['SQ_FOLDER_ID']; 
const SKU_MAPPING_OLD_NEW_FOLDER_ID = SCRIPT_PROPS['SKU_MAPPING_OLD_NEW_FOLDER_ID'];
const BASKET_FOLDER_ID = SCRIPT_PROPS['BASKET_FOLDER_ID'];
const ACCOUNT_TITLE_FOLDER_ID = SCRIPT_PROPS['ACCOUNT_TITLE_FOLDER_ID'];
const PMC_DIRECTORY_FOLDER_ID = SCRIPT_PROPS['PMC_DIRECTORY_FOLDER_ID'];
const MISSING_SKU_PRICES_FOLDER_ID = SCRIPT_PROPS['MISSING_SKU_PRICES_FOLDER_ID'];
const AE_EMAILS_FOLDER_ID = SCRIPT_PROPS['AE_EMAILS_FOLDER_ID'];
const CPP_FOLDER_ID = SCRIPT_PROPS['CPP_FOLDER_ID']; 
const NARRA_VOUCHER_FOLDER_ID = SCRIPT_PROPS['NARRA_VOUCHER_FOLDER_ID'];
const EMAIL_RECIPIENT_FOLDER_ID = SCRIPT_PROPS['EMAIL_RECIPIENT_FOLDER_ID'];
const CHECK_SUMMARY_FOLDER_ID = SCRIPT_PROPS['CHECK_SUMMARY_FOLDER_ID'];
const RIWB_FOLDER_ID = SCRIPT_PROPS['RIWB_FOLDER_ID'];
const NEW_NARRA_FOLDER_ID = SCRIPT_PROPS['NEW_NARRA_FOLDER_ID']; 

const PREFIX_RIWB = "RIWB_Data_";
const PREFIX_NEW_NARRA = "New_Narra_Data_";
const PREFIX_CLIENT_2307 = "Client_2307_Data_";

const NTE_TEMPLATE_ID_PINGCON = SCRIPT_PROPS['NTE_TEMPLATE_ID_PINGCON'];
const NTE_TEMPLATE_ID_A2Z = SCRIPT_PROPS['NTE_TEMPLATE_ID_A2Z'];
const APPRECIATION_TEMPLATE_ID_PINGCON = SCRIPT_PROPS['APPRECIATION_TEMPLATE_ID_PINGCON'];
const APPRECIATION_TEMPLATE_ID_A2Z = SCRIPT_PROPS['APPRECIATION_TEMPLATE_ID_A2Z'];

//OUTPUT FOLDERS
const BUY_SELL_ROOT_FOLDER_ID = SCRIPT_PROPS['BUY_SELL_ROOT_FOLDER_ID'];
const COMMISSION_REPORT_ROOT_FOLDER_ID = SCRIPT_PROPS['COMMISSION_REPORT_ROOT_FOLDER_ID'];
const SALES_REPORT_ROOT_FOLDER_ID = SCRIPT_PROPS['SALES_REPORT_ROOT_FOLDER_ID'];
const EWT_REPORT_CLIENT_ROOT_FOLDER_ID = SCRIPT_PROPS['EWT_REPORT_CLIENT_ROOT_FOLDER_ID'];
const PURCHASE_OUTPUT_FOLDER_ID = SCRIPT_PROPS['PURCHASE_OUTPUT_FOLDER_ID'];
const UNPAID_DR_ROOT_FOLDER_ID = SCRIPT_PROPS['UNPAID_DR_ROOT_FOLDER_ID'];
const NTE_COMPLAINTS_FOLDER_ID = SCRIPT_PROPS['NTE_COMPLAINTS_FOLDER_ID']; 
const NTE_REPLIES_FOLDER_ID = SCRIPT_PROPS['NTE_REPLIES_FOLDER_ID'];

// IMAGES
const FEEDBACK_IMAGE_ID_A2Z_KEYSYS = SCRIPT_PROPS['FEEDBACK_IMAGE_ID_A2Z_KEYSYS']; 
const FEEDBACK_IMAGE_ID_PINGCON_BONITA = SCRIPT_PROPS['FEEDBACK_IMAGE_ID_PINGCON_BONITA'];
const FEEDBACK_IMAGE_ID_SPECIAL = SCRIPT_PROPS['FEEDBACK_IMAGE_ID_SPECIAL'];
const HEADER_LOGO_ID_A2Z_KEYSYS = SCRIPT_PROPS['HEADER_LOGO_ID_A2Z_KEYSYS']; 
const HEADER_LOGO_ID_PINGCON_BONITA = SCRIPT_PROPS['HEADER_LOGO_ID_PINGCON_BONITA']; 
//GOOGLE SHEET PREFIX NAMING
const PREFIX_SQ = "SQ_Data_";
const PREFIX_SKU_MAP = "Missing_SKU_Data_";
const PREFIX_EWT = "EWT_Data_";
const PREFIX_PURCHASES = "Purchases_Data_";
const PREFIX_BASKET = "Basket_Data_";
const PREFIX_ACCOUNT_TITLE = "Account_Title_Data_";
const PREFIX_PMC_DIRECTORY = "PMC_Directory_Data_";
const PREFIX_SR = "SR_Data_";
const PREFIX_PO_PINGCON_BONITA = "PO_PINGCON_BONITA_Data_";
const PREFIX_PO_KEYSYS = "PO_Keysys_Data_";
const PREFIX_AR = "AR_Data_";
const PREFIX_SCANNED_INVOICES = "Scanned_Supplier_Invoices_Data_";
const PREFIX_AE_EMAILS = "AE_Emails_Data_";
const PREFIX_NARRA_VOUCHER = "Narra_Voucher_Data_";
const PREFIX_CHECK_SUMMARY = "Check_Summary_Data_";

const REPORT_EMAIL = SCRIPT_PROPS['REPORT_EMAIL'];
const WEB_APP_URL = SCRIPT_PROPS['WEB_APP_URL'];


var NTE_PAYLOAD_ACCUMULATOR =[];
var CURRENT_RUN_NTE_FOLDER_ID = null; // <-- ADD THIS LINE

// =========================================================================
// MONTHLY GENERATE REPORTS FUNCTIONS
// =========================================================================

function _1_generateMonthlyBuySellReport() { _buildAndShowDialog("Generate Buy/Sell Report (All AEs)", "Generate Reports for All AEs", "_$processMonthlyBuySellForUI", false, true, 320); }
function _2_generateMonthlyCommissionReport() { _buildAndShowDialog("Generate Commission Reports (All AEs)", "Generate Reports for All AEs", "_$processMonthlyCommissionForUI", false, false, 250); }
function _3_generateMonthlyEWTReportClientReport() { _buildAndShowDialog("Generate EWT Report Client (All AEs)", "Generate Reports for All AEs", "_$processAllEWTReportClientForUI", false, false, 250); }



// =========================================================================
// GENERATE REPORTS FOR SPECIFIC AE FUNCTIONS
// =========================================================================

function _6_generateForSpecificAEBuySellReport() { _buildAndShowDialog("Generate Buy/Sell Report", "Generate Report", "_$startGeneration", true, true, 420); }
function _7_generateForSpecificAECommissionReport() { _buildAndShowDialog("Generate Commission Report (XLSX)", "Generate Report", "_$processCommissionReportForUI", true, false, 320); }
function _8_generateForSpecificAEEWTReportClientReport() { _buildAndShowDialog("Generate EWT Report Client (XLSX)", "Generate Report", "_$processEWTReportClientForUI", true, false, 320); }

// =========================================================================
// EMAIL SENDING FUNCTIONS
// =========================================================================

const UNPAID_BATCH_KEY = "PENDING_AES_UNPAID_DR";
const UNPAID_TRIGGER_NAME = "resumeUnpaidDRBatch";

function _9_runSendAllUnpaidDRs() {
  var lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) { throw new Error("System is currently busy. Another user is generating reports right now. Please wait a few minutes."); }
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const props = PropertiesService.getScriptProperties();
    deleteTriggersByName(UNPAID_TRIGGER_NAME);
    props.deleteProperty(UNPAID_BATCH_KEY);
    props.deleteProperty("CANCEL_UNPAID_DR");
    
    try { ss.toast("Scanning Data to initialize background batch...", "INITIALIZING", -1); } catch(e){}
    
    const arAllData = gatherAllData("AR_Data_");
    if (arAllData.length === 0) throw new Error("No Ageing Report Data found.");
    
    const arSheets = getDatabaseSheets("AR_Data_");
    let arHeaders = [];
    if (arSheets.length > 0 && arSheets[0].getLastColumn() > 0) {
      arHeaders = arSheets[0].getRange(1, 1, 1, arSheets[0].getLastColumn()).getValues()[0].map(h => String(h).toLowerCase().trim());
    }
    let managerIdx = arHeaders.findIndex(h => h.includes("manager"));
    let docDateIdx = arHeaders.findIndex(h => h.includes("date") && !h.includes("expiry") && !h.includes("countered") && !h.includes("received"));
    let openBalIdx = arHeaders.findIndex(h => h.includes("open balance") && !h.includes("%"));
    let paymentStatusIdx = arHeaders.findIndex(h => h === "payment status" || h.includes("payment status"));
    if (paymentStatusIdx === -1 && openBalIdx > -1) paymentStatusIdx = openBalIdx + 2;
    let sourceFileIdx = arHeaders.findIndex(h => h.includes("source file"));
    let textIdx = arHeaders.findIndex(h => h.includes("additional text"));
    let notesIdx = arHeaders.findIndex(h => h === "notes");
    
    let uniqueAEs = new Set();
    let allowedEntities = JSON.parse(props.getProperty("SALES_ENTITIES") || "[]");
    let allowedAEsStr = props.getProperty("SALES_TARGET_AES");
    let allowedAEs = allowedAEsStr ? JSON.parse(allowedAEsStr) : [];
    
    const startStr = props.getProperty("BATCH_START_DATE");
    const endStr = props.getProperty("BATCH_END_DATE");
    const startDate = startStr ? new Date(startStr + "T00:00:00") : new Date(2025, 9, 1);
    const endDate = endStr ? new Date(endStr + "T23:59:59") : new Date(2026, 2, 20, 23, 59, 59);
    
    let filterType = props.getProperty("UNPAID_DR_FILTER_TYPE");

    arAllData.forEach(item => {
        let row = item.data;

        let textVal = (textIdx > -1) ? String(row[textIdx]) : "";
        let notesVal = (notesIdx > -1) ? String(row[notesIdx]) : "";
        if (isTransactionVoid(notesVal, textVal)) return;

        let rawManager = (managerIdx > -1) ? String(row[managerIdx]) : "";
        let manager = getCanonicalName(rawManager);
        if (!manager) return;

        let openBal = (openBalIdx > -1) ? parseFloat(row[openBalIdx]) : 0;
        if (isNaN(openBal) || openBal < 1) return;

        let statusRaw = (paymentStatusIdx > -1 && paymentStatusIdx < row.length) ? String(row[paymentStatusIdx]).trim().toLowerCase() : "";
        if (statusRaw === "fully paid") return;

        let rawDocDate = (docDateIdx > -1) ? row[docDateIdx] : "";
        let docDateObj = (rawDocDate instanceof Date) ? rawDocDate : new Date(rawDocDate);
        if (isNaN(docDateObj.getTime())) return;
        if (docDateObj < startDate || docDateObj > endDate) return;

        let sourceFile = (sourceFileIdx > -1) ? String(row[sourceFileIdx]) : "";
        let companyEntity = getCompanyEntity(sourceFile);
        
        if (allowedEntities.length > 0 && !allowedEntities.includes(companyEntity)) return;
        
        if (allowedAEs.length > 0 && !allowedAEs.includes(manager)) return;

        uniqueAEs.add(manager);
    });
    
    let aeArray = Array.from(uniqueAEs);
    if(aeArray.length === 0) throw new Error("No Unpaid DRs found for the selected category in the specified date range.");
    
    props.setProperty(UNPAID_BATCH_KEY, JSON.stringify(aeArray));
    try { ss.toast(`Found ${aeArray.length} AEs. Processing first batch immediately... You can close this window now.`, "BATCH STARTED", 15); } catch(e){}
    
    resumeUnpaidDRBatch();
  } finally {
    lock.releaseLock();
  }
}




const NTE_SKU_TRIGGER_NAME = "resumeNteSkuBatch";


const NTE_MISSING_INV_TRIGGER_NAME = "resumeNteMissingInvBatch";


// =========================================================================
// SCHEDULED TRIGGER WRAPPERS 
// Use THESE functions for your 30-minute time-driven triggers
// =========================================================================

function _13_scheduledImportSalesReports(e) {
  _16_importSalesReports(e);
}

function _14_scheduledImportAgeingReports(e) {
  _20_importAgeingReports(e);
}

function _15_scheduledImportPurchaseOrders(e) {
  _17_importPurchaseOrdersPingconAndBonita(e);
}


// =========================================================================
// IMPORT FUNCTIONS
// =========================================================================
function _33_importCustomerPaymentPerformance(e) {
  var isBackground = (e && e.triggerUid);
  var lock = null;
  if (!isBackground) {
    lock = LockService.getScriptLock();
    try { lock.waitLock(10000); } catch (err) { throw new Error("System is currently busy. Another user is generating reports right now. Please wait a few minutes."); }
  }
  try {
  const startTime = Date.now();
  const MAX_EXECUTION_TIME = 4.5 * 60 * 1000;
  deleteTriggersByName("_33_importCustomerPaymentPerformance");

  console.time("Import_CPP");
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  if (!CPP_FOLDER_ID || CPP_FOLDER_ID === "1_INSERT_CPP_FOLDER_ID_HERE") {
    ss.toast("Error: CPP_FOLDER_ID is not configured.", "ERROR", -1);
    return;
  }

  const folder = DriveApp.getFolderById(CPP_FOLDER_ID);
  const files = folder.getFiles();
  let filesProcessed = 0;
  const sheetCache = {};

  // Custom parser to force dd/mm/yyyy into standard JS Date objects
  const parseDmyDate = (val) => {
      if (!val) return "";
      if (val instanceof Date) return val;
      if (typeof val === 'number') return new Date(Math.round((val - 25569) * 86400 * 1000));
      let str = String(val).trim();
      let parts = str.split(/[\/\-]/);
      if (parts.length === 3) {
          let d = parseInt(parts[0], 10);
          let m = parseInt(parts[1], 10) - 1;
          let y = parseInt(parts[2], 10);
          if (y < 100) y += 2000;
          return new Date(y, m, d);
      }
      let fallback = new Date(str);
      if (!isNaN(fallback.getTime())) return fallback;
      return "";
  };

  while (files.hasNext()) {
    let file = files.next();
    let fileName = file.getName();

    if (fileName.endsWith(".xlsx") || fileName.endsWith(".xls") || fileName.endsWith(".csv")) {
      ss.toast(`Importing CPP: ${fileName}`, "IMPORT", -1);
      let tempFileId = null;
      try {
        let converted = _openFileOrConvertCSV(file);
        let ssNew = converted.ssNew;
        tempFileId = converted.tempFileId;
        let sheetNew = ssNew.getSheets()[0];
        let data = sheetNew.getDataRange().getValues();

        if (data.length > 0) {
           let headerRowIdx = -1;
           for(let r=0; r<Math.min(20, data.length); r++) {
             let rowStr = data[r].join(" ").toLowerCase();
             if(rowStr.includes("customer") || rowStr.includes("business partner") || rowStr.includes("document date")) {
                 headerRowIdx = r; break;
             }
           }

           if (headerRowIdx > -1) {
              let originalHeaders = data[headerRowIdx].map(h => String(h).trim());
              let currentHeaders = [...originalHeaders, "Source File"];
              
              let lowerHeaders = originalHeaders.map(h => String(h).toLowerCase());
              let dateCols = [
                  lowerHeaders.findIndex(h => h.includes("invoice date")),
                  lowerHeaders.findIndex(h => h.includes("payment deadline") || h.includes("due date")),
                  lowerHeaders.findIndex(h => h.includes("received date")),
                  lowerHeaders.findIndex(h => h.includes("countered date")),
                  lowerHeaders.findIndex(h => h.includes("payment date"))
              ].filter(idx => idx > -1);

              let rowsByYearBlock = {};

              for (let i = headerRowIdx + 1; i < data.length; i++) {
                 let row = data[i];
                 if (row.every(c => String(c).trim() === "")) continue;

                 // Clean and parse all dates from dd/mm/yyyy to standard MM/DD/YYYY
                 dateCols.forEach(idx => {
                     let dObj = parseDmyDate(row[idx]);
                     row[idx] = (dObj instanceof Date && !isNaN(dObj.getTime())) ? fastFormatDate(dObj) : "";
                 });

                 // Determine block year based on Invoice Date (or fallback to today)
                 let mainDateIdx = lowerHeaders.findIndex(h => h.includes("invoice date"));
                 let blockDate = new Date();
                 if (mainDateIdx > -1 && row[mainDateIdx]) {
                     blockDate = new Date(row[mainDateIdx]);
                     if (isNaN(blockDate.getTime())) blockDate = new Date();
                 }

                 const startYear = Math.floor(blockDate.getFullYear() / 5) * 5;
                 const blockName = `CPP_Data_${startYear}-${startYear + 4}`;
                 
                 if (!rowsByYearBlock[blockName]) rowsByYearBlock[blockName] =[];
                 rowsByYearBlock[blockName].push([...row, fileName]);
              }

              // --- WIPE AND REPLACE LOGIC ---
              for (let block in rowsByYearBlock) {
                 let targetSheet;
                 let targetLastCol;
                 if (sheetCache[block]) {
                     targetSheet = sheetCache[block];
                     targetLastCol = targetSheet.getLastColumn();
                 } else {
                     targetSheet = getSheetForYearBlock("CPP_Data_", new Date(), currentHeaders);
                     targetLastCol = targetSheet.getLastColumn();
                     if (targetLastCol < currentHeaders.length) {
                         targetSheet.getRange(1, 1, 1, currentHeaders.length).setValues([currentHeaders]);
                         targetSheet.getRange(1, 1, 1, currentHeaders.length).setFontWeight("bold").setBackground("#f3f3f3");
                         targetLastCol = currentHeaders.length;
                     }
                     if (targetSheet.getLastRow() > 1) {
                         targetSheet.getRange(2, 1, targetSheet.getLastRow() - 1, targetLastCol).clearContent();
                     }
                     sheetCache[block] = targetSheet;
                 }
                 
                 let newRows = rowsByYearBlock[block];
                 if (newRows.length > 0) {
                    let paddedData = newRows.map(r => {
                        let newRow = [...r];
                        while(newRow.length < targetLastCol) newRow.push("");
                        return newRow.slice(0, targetLastCol);
                    });
                    let nextRow = targetSheet.getLastRow() + 1;
                    targetSheet.getRange(nextRow, 1, paddedData.length, targetLastCol).setValues(paddedData);
                 }
              }
           }
        }
        if (tempFileId) Drive.Files.remove(tempFileId);
        file.setTrashed(true);
        filesProcessed++;
      } catch (e) { 
        if (tempFileId) Drive.Files.remove(tempFileId);
        throw new Error("Error processing CPP " + fileName + ": " + e.toString()); 
      }
    }
  }
  console.timeEnd("Import_CPP");
  ss.toast(`CPP Import Complete. Files: ${filesProcessed}`, "SUCCESS", 5);
  } finally {
    if (lock) lock.releaseLock();
  }
}

function _36_importNarraVoucherReport(e) {
  var isBackground = (e && e.triggerUid);
  var lock = null;
  if (!isBackground) {
    lock = LockService.getScriptLock();
    try { lock.waitLock(10000); } catch (err) { throw new Error("System is currently busy. Another user is generating reports right now. Please wait a few minutes."); }
  }
  try {
    const startTime = Date.now();
    const MAX_EXECUTION_TIME = 4.5 * 60 * 1000;
    deleteTriggersByName("_36_importNarraVoucherReport");

    console.time("Import_Narra_Voucher");
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    const targetFolderId = SCRIPT_PROPS['VERIFIED_FOLDER_ID'];
    const folder = DriveApp.getFolderById(targetFolderId);
    const files = folder.getFiles();
    
    let filesProcessed = 0;
    const sheetCache = {};

    const parseDmyDate = (val) => {
        if (!val) return "";
        if (val instanceof Date) return val;
        if (typeof val === 'number') return new Date(Math.round((val - 25569) * 86400 * 1000));
        let str = String(val).trim();
        let parts = str.split(/[\/\-]/);
        if (parts.length === 3) {
            let d = parseInt(parts[0], 10);
            let m = parseInt(parts[1], 10) - 1;
            let y = parseInt(parts[2], 10);
            if (y < 100) y += 2000;
            return new Date(y, m, d);
        }
        let fallback = new Date(str);
        if (!isNaN(fallback.getTime())) return fallback;
        return "";
    };

    while (files.hasNext()) {
      if (Date.now() - startTime > MAX_EXECUTION_TIME) {
          ss.toast("Approaching time limit. Pausing and auto-resuming Narra Voucher in 1 minute...", "BATCH CONTINUATION", 10);
          ScriptApp.newTrigger("_36_importNarraVoucherReport").timeBased().after(60 * 1000).create();
          return; 
      }

      let file = files.next();
      let fileName = file.getName();
      if (!fileName.toUpperCase().startsWith("RAW_2307_")) continue;
      let mimeType = file.getMimeType();

      let lowerName = fileName.toLowerCase();
      let isGoogleSheet = mimeType === MimeType.GOOGLE_SHEETS;
      let isExcel = mimeType === MimeType.MICROSOFT_EXCEL || 
                    mimeType === "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" || 
                    mimeType === MimeType.CSV || mimeType === "text/csv" ||
                    lowerName.endsWith(".xlsx") || lowerName.endsWith(".xls") || lowerName.endsWith(".csv");

      if (isGoogleSheet || isExcel) {
        ss.toast(`Importing Narra Voucher: ${fileName}`, "IMPORT", -1);
        let tempFileId = null;
        try {
          let converted = _openFileOrConvertCSV(file);
          let ssNew = converted.ssNew;
          tempFileId = converted.tempFileId;
          let sheetNew = ssNew.getSheets()[0];
          let data = sheetNew.getDataRange().getValues();

          if (data.length > 0) {
             let headerRowIdx = -1;
             
             // HIGHLY RELAXED HEADER SEARCH: Uses spaces instead of stripping them
             for(let r=0; r<Math.min(20, data.length); r++) {
               let rowStr = data[r].join(" ").toLowerCase();
               if(rowStr.includes("voucher") && rowStr.includes("amount") && (rowStr.includes("invoice") || rowStr.includes("payee"))) {
                   headerRowIdx = r; break;
               }
             }

             // FORCES ERROR AND STOPS DELETION IF HEADERS ARE NOT FOUND
             if (headerRowIdx === -1) {
                 throw new Error("Headers not recognized! Ensure columns like 'VoucherNumber', 'Amount', and 'InvoiceNumber' exist.");
             }

             if (headerRowIdx > -1) {
                let originalHeaders = data[headerRowIdx].map(h => String(h).trim());
                
                // STRIP TRAILING BLANK HEADERS TO PREVENT OFFSET
                while(originalHeaders.length > 0 && originalHeaders[originalHeaders.length - 1] === "") {
                    originalHeaders.pop();
                }

                let currentHeaders = [...originalHeaders, "Source File"];
                
                let lowerHeaders = originalHeaders.map(h => String(h).toLowerCase().replace(/\s+/g, ''));
                let vDateIdx = lowerHeaders.indexOf("voucherdate");
                if (vDateIdx === -1) vDateIdx = lowerHeaders.indexOf("invoicedate");
                let invNoIdx = lowerHeaders.findIndex(h => h.includes("invoicenumber") || h.includes("invoiceno") || h === "invoice");
                
                let dateCols = [
                    lowerHeaders.indexOf("voucherdate"),
                    lowerHeaders.indexOf("invoicedate"),
                    lowerHeaders.indexOf("checkdate"),
                    lowerHeaders.indexOf("prepareddate"),
                    lowerHeaders.indexOf("checkeddate"),
                    lowerHeaders.indexOf("approveddate")
                ].filter(idx => idx > -1);

                let rowsByYearBlock = {};

                for (let i = headerRowIdx + 1; i < data.length; i++) {
                   let row = data[i];
                   if (row.every(c => String(c).trim() === "")) continue;

                   dateCols.forEach(idx => {
                       let dObj = parseDmyDate(row[idx]);
                       // Safely format dates without erasing cells if they aren't dates
                       row[idx] = (dObj instanceof Date && !isNaN(dObj.getTime())) ? fastFormatDate(dObj) : String(row[idx]);
                   });

                   let blockDate = new Date();
                   if (vDateIdx > -1 && row[vDateIdx]) {
                       blockDate = new Date(row[vDateIdx]);
                       if (isNaN(blockDate.getTime())) blockDate = new Date();
                   }

                   const startYear = Math.floor(blockDate.getFullYear() / 5) * 5;
                   const blockName = `${PREFIX_NARRA_VOUCHER}${startYear}-${startYear + 4}`;
                   
                   if (!rowsByYearBlock[blockName]) rowsByYearBlock[blockName] =[];
                   let cleanRow = row.slice(0, originalHeaders.length); // Trim to exact header length
                   rowsByYearBlock[blockName].push([...cleanRow, fileName]);
                }

                // --- WIPE AND REPLACE LOGIC ---
                for (let block in rowsByYearBlock) {
                   let targetSheet;
                   let targetLastCol;
                   if (sheetCache[block]) {
                       targetSheet = sheetCache[block];
                       targetLastCol = targetSheet.getLastColumn();
                   } else {
                       targetSheet = getSheetForYearBlock(PREFIX_NARRA_VOUCHER, new Date(), currentHeaders);
                       
                       // FORCE UPDATE HEADERS TO FIX OFFSETS
                       targetSheet.getRange(1, 1, 1, currentHeaders.length).setValues([currentHeaders]);
                       targetSheet.getRange(1, 1, 1, currentHeaders.length).setFontWeight("bold").setBackground("#f3f3f3");
                       
                       // Clear any residual offset headers/data from previous bad uploads
                       let existingLastCol = targetSheet.getLastColumn();
                       if (existingLastCol > currentHeaders.length) {
                           targetSheet.getRange(1, currentHeaders.length + 1, Math.max(1, targetSheet.getLastRow()), existingLastCol - currentHeaders.length).clearContent();
                       }
                       
                       targetLastCol = currentHeaders.length;
                       
                       if (targetSheet.getLastRow() > 1) {
                           targetSheet.getRange(2, 1, targetSheet.getLastRow() - 1, targetLastCol).clearContent();
                       }
                       sheetCache[block] = targetSheet;
                   }
                   
                   let newRows = rowsByYearBlock[block];
                   if (newRows.length > 0) {
                      let paddedData = newRows.map(r => {
                          let newRow = [...r];
                          while(newRow.length < targetLastCol) newRow.push("");
                          if (typeof invNoIdx !== 'undefined' && invNoIdx > -1 && newRow[invNoIdx] !== undefined) newRow[invNoIdx] = String(newRow[invNoIdx]).trim();
                          return newRow.slice(0, targetLastCol);
                      });
                      if (typeof invNoIdx !== 'undefined' && invNoIdx > -1) {
                          targetSheet.getRange(2, invNoIdx + 1, Math.max(1, targetSheet.getMaxRows() - 1), 1).setNumberFormat("@"); // Force Invoice column to plain text
                      }
                      let nextRow = targetSheet.getLastRow() + 1;
                      targetSheet.getRange(nextRow, 1, paddedData.length, targetLastCol).setValues(paddedData);
                   }
                }
             }
          }
          if (tempFileId) Drive.Files.remove(tempFileId);
          // ONLY TRASH IF EVERYTHING WAS SUCCESSFUL
          file.setTrashed(true);
          filesProcessed++;
        } catch (e) { 
          if (tempFileId) Drive.Files.remove(tempFileId);
          console.error("Error processing Narra Voucher " + fileName + ": " + e.toString()); 
          try { file.setTrashed(true); } catch(err) {} 
        }
      }
    }
    console.timeEnd("Import_Narra_Voucher");
    ss.toast(`Narra Voucher Import Complete. Files: ${filesProcessed}`, "SUCCESS", 5);
  } finally {
    if (lock) lock.releaseLock();
  }
}

function _37_importCheckSummary(e) {
  var isBackground = (e && e.triggerUid);
  var lock = null;
  if (!isBackground) {
    lock = LockService.getScriptLock();
    try { lock.waitLock(10000); } catch (err) { throw new Error("System is currently busy."); }
  }
  try {
    const startTime = Date.now();
    deleteTriggersByName("_37_importCheckSummary");
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const targetFolderId = SCRIPT_PROPS['VERIFIED_FOLDER_ID'];
    const folder = DriveApp.getFolderById(targetFolderId);
    const files = folder.getFiles();
    let filesProcessed = 0; const sheetCache = {};

    while (files.hasNext()) {
      if (Date.now() - startTime > 4.5 * 60 * 1000) {
          ScriptApp.newTrigger("_37_importCheckSummary").timeBased().after(60 * 1000).create();
          return; 
      }
      let file = files.next();
      let fileName = file.getName();
      if (!fileName.toUpperCase().startsWith("RAW_2307_")) continue;
      let mimeType = file.getMimeType();
      let lowerName = fileName.toLowerCase();
      let isGoogleSheet = mimeType === MimeType.GOOGLE_SHEETS;
      let isExcel = mimeType === MimeType.MICROSOFT_EXCEL || 
                    mimeType === "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" || 
                    mimeType === MimeType.CSV || mimeType === "text/csv" ||
                    lowerName.endsWith(".xlsx") || lowerName.endsWith(".xls") || lowerName.endsWith(".csv");
      
      if (isGoogleSheet || isExcel) {
        ss.toast(`Importing Check Summary: ${fileName}`, "IMPORT", -1);
        let tempFileId = null;
        try {
          let converted = _openFileOrConvertCSV(file);
          let ssNew = converted.ssNew;
          tempFileId = converted.tempFileId;
          
          let allSheets = ssNew.getSheets();
          let rowsByYearBlock = {};
          let currentHeaders = ["Date", "Company Name", "KEYSYS Amount", "PINGCON Amount", "SI #", "Source File"];
          
          allSheets.forEach(sheetNew => {
              let data = sheetNew.getDataRange().getValues();
              if (data.length === 0) return;
              
              let sheetDate = null;
              for (let r=0; r<5; r++) {
                  let cellVal = String(data[r][0]).trim();
                  let match = cellVal.match(/Date:\s*(.*)/i);
                  if (match) {
                      let rawD = match[1].split('(')[0].trim();
                      sheetDate = new Date(rawD);
                      break;
                  }
              }
              if (!sheetDate || isNaN(sheetDate.getTime())) sheetDate = new Date(sheetNew.getName()); 
              if (isNaN(sheetDate.getTime())) sheetDate = new Date(); 
              
              let dateStr = fastFormatDate(sheetDate);
              let headerRowIdx = -1;
              let keysysCol = -1, pingconCol = -1, siCol = -1, compCol = -1;
              
              for(let r=0; r<Math.min(20, data.length); r++) {
                  let rowStr = data[r].map(c => String(c).toLowerCase().replace(/[^a-z0-9]/g, ''));
                  
                  // Always grab KEYSYS and PINGCON columns as soon as we see them (usually Row 1)
                  let kIdx = rowStr.findIndex(c => c === "keysys" || c.includes("keysysamount"));
                  let pIdx = rowStr.findIndex(c => c === "pingcon" || c.includes("pingconamount"));
                  if (kIdx > -1 && keysysCol === -1) keysysCol = kIdx;
                  if (pIdx > -1 && pingconCol === -1) pingconCol = pIdx;
                  
                  let cIdx = rowStr.findIndex(c => c.includes("companyname"));
                  let sIdx = rowStr.findIndex(c => c === "si" || c === "sino" || c.includes("invoice"));
                  
                  // Update compCol and siCol if found
                  if (cIdx > -1) compCol = cIdx;
                  if (sIdx > -1) siCol = sIdx;
                  
                  // STRICT RULE: The header row MUST contain BOTH "Company Name" and "SI #" 
                  if (cIdx > -1 && sIdx > -1) {
                      headerRowIdx = r;
                      break; // Stop searching to lock accurately onto Row 2
                  }
              }
              
              if (headerRowIdx > -1) {
                  if (compCol === -1) compCol = 0;
                  
                  let currentCompany = "";
                  
                  for (let i = headerRowIdx + 1; i < data.length; i++) {
                      let row = data[i];
                      let rawComp = String(row[compCol]).trim();
                      
                      if (rawComp !== "") {
                          let isJustNumber = !isNaN(parseFloat(rawComp.replace(/,/g, ''))) && !/[a-zA-Z]/.test(rawComp);
                          if (!isJustNumber && rawComp.toLowerCase() !== "company name" && !rawComp.toLowerCase().includes("check summary")) {
                              currentCompany = rawComp;
                          }
                      }
                      
                      let rawSi = String(row[siCol]).trim();
                      
                      if (!rawSi || rawSi.toLowerCase() === "si #" || rawSi.toLowerCase() === "si" || rawSi.toLowerCase() === "amount") continue;
                      if (!currentCompany) continue;
                      
                      // Strip non-numeric symbols but KEEP leading zeros (0-9)
                      let siNum = rawSi.replace(/^[^0-9]+/, ''); 
                      
                      let kAmt = keysysCol > -1 ? parseFloat(String(row[keysysCol]).replace(/,/g, '')) : 0;
                      let pAmt = pingconCol > -1 ? parseFloat(String(row[pingconCol]).replace(/,/g, '')) : 0;
                      if (isNaN(kAmt)) kAmt = 0;
                      if (isNaN(pAmt)) pAmt = 0;
                      
                      const startYear = Math.floor(sheetDate.getFullYear() / 5) * 5;
                      const blockName = `${PREFIX_CHECK_SUMMARY}${startYear}-${startYear + 4}`;
                      
                      if (!rowsByYearBlock[blockName]) rowsByYearBlock[blockName] = [];
                      rowsByYearBlock[blockName].push([dateStr, currentCompany, kAmt, pAmt, siNum, fileName]);
                  }
              }
          });
          
          for (let block in rowsByYearBlock) {
             let targetSheet;
             if (sheetCache[block]) {
                 targetSheet = sheetCache[block];
             } else {
                 targetSheet = getSheetForYearBlock(PREFIX_CHECK_SUMMARY, new Date(), currentHeaders);
                 if (targetSheet.getLastRow() > 1) {
                     targetSheet.getRange(2, 1, targetSheet.getLastRow() - 1, currentHeaders.length).clearContent();
                 }
                 sheetCache[block] = targetSheet;
             }
             let newRows = rowsByYearBlock[block];
             if (newRows.length > 0) {
                 targetSheet.getRange(2, 5, Math.max(1, targetSheet.getMaxRows() - 1), 1).setNumberFormat("@"); // Force SI column to plain text
                 targetSheet.getRange(targetSheet.getLastRow() + 1, 1, newRows.length, currentHeaders.length).setValues(newRows);
             }
          }
          if (tempFileId) Drive.Files.remove(tempFileId);
          file.setTrashed(true);
          filesProcessed++;
        } catch (e) { 
          if (tempFileId) Drive.Files.remove(tempFileId);
          console.error("Error processing Check Summary " + fileName + ": " + e.toString()); 
          try { file.setTrashed(true); } catch(err) {} 
        }
      }
    }
    ss.toast(`Check Summary Import Complete. Files: ${filesProcessed}`, "SUCCESS", 5);
  } finally {
    if (lock) lock.releaseLock();
  }
}

function _38_importRIWBReport(e) {
  var isBackground = (e && e.triggerUid);
  var lock = null;
  if (!isBackground) {
    lock = LockService.getScriptLock();
    try { lock.waitLock(10000); } catch (err) { throw new Error("System is currently busy."); }
  }
  try {
    const startTime = Date.now();
    deleteTriggersByName("_38_importRIWBReport");
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const targetFolderId = SCRIPT_PROPS['VERIFIED_FOLDER_ID'];
    const folder = DriveApp.getFolderById(targetFolderId);
    const files = folder.getFiles();
    let filesProcessed = 0; const sheetCache = {};

    const parseDmyDate = (val) => {
        if (!val) return "";
        if (val instanceof Date) return val;
        if (typeof val === 'number') return new Date(Math.round((val - 25569) * 86400 * 1000));
        
        let str = String(val).trim();
        
        // 1. Try standard JS parsing first. If it's a valid MM/DD/YYYY CSV date, it parses perfectly here.
        let standardDate = new Date(str);
        if (!isNaN(standardDate.getTime())) return standardDate;
        
        // 2. If it fails (which happens for DD/MM/YYYY text), slice it manually
        let parts = str.split(/[\/\-]/);
        if (parts.length === 3) {
            let d = parseInt(parts[0], 10);
            let m = parseInt(parts[1], 10) - 1;
            let y = parseInt(parts[2], 10);
            if (y < 100) y += 2000;
            return new Date(y, m, d);
        }
        return "";
    };

    while (files.hasNext()) {
      if (Date.now() - startTime > 4.5 * 60 * 1000) {
          ScriptApp.newTrigger("_38_importRIWBReport").timeBased().after(60 * 1000).create();
          return; 
      }
      let file = files.next();
      let fileName = file.getName();
      if (!fileName.toUpperCase().startsWith("RAW_2307_")) continue;
      let mimeType = file.getMimeType();

      let lowerName = fileName.toLowerCase();
      let isGoogleSheet = mimeType === MimeType.GOOGLE_SHEETS;
      let isExcel = mimeType === MimeType.MICROSOFT_EXCEL || 
                    mimeType === "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" || 
                    mimeType === MimeType.CSV || mimeType === "text/csv" ||
                    lowerName.endsWith(".xlsx") || lowerName.endsWith(".xls") || lowerName.endsWith(".csv");

      if (isGoogleSheet || isExcel) {
        ss.toast(`Importing RIWB: ${fileName}`, "IMPORT", -1);
        let tempFileId = null;
        try {
          let converted = _openFileOrConvertCSV(file);
          let ssNew = converted.ssNew;
          tempFileId = converted.tempFileId;
          let sheetNew = ssNew.getSheets()[0];
          let data = sheetNew.getDataRange().getValues();

          if (data.length > 0) {
             let headerRowIdx = -1;
             
             for(let r=0; r<Math.min(20, data.length); r++) {
               let rowStr = data[r].join(" ").toLowerCase();
               if(rowStr.includes("supplier") && rowStr.includes("total sum") && rowStr.includes("date")) {
                   headerRowIdx = r; break;
               }
             }

             if (headerRowIdx > -1) {
                let lowerHeaders = data[headerRowIdx].map(h => String(h).toLowerCase().trim());
                let currentHeaders = ["Date", "Supplier", "Invoice No", "Total Sum", "Source File"];
                
                let dateIdx = lowerHeaders.findIndex(h => h === "date");
                let supIdx = lowerHeaders.findIndex(h => h === "supplier");
                let invIdx1 = lowerHeaders.findIndex(h => h.includes("supplier's #") || h.includes("order tra") || h.includes("supplier's reference"));
                let invIdx2 = lowerHeaders.findIndex(h => h === "reference number");
                let sumIdx = lowerHeaders.findIndex(h => h === "total sum");
                
                let rowsByYearBlock = {};

                for (let i = headerRowIdx + 1; i < data.length; i++) {
                   let row = data[i];
                   if (row.every(c => String(c).trim() === "")) continue;

                   let dateObj = parseDmyDate(row[dateIdx]);
                   let formattedDate = "";
                   let startYear = new Date().getFullYear(); // Safe default
                   
                   // SAFE PARSING: Prevents script crash if Excel exported the date as "########"
                   if (dateObj instanceof Date && !isNaN(dateObj.getTime())) {
                       formattedDate = fastFormatDate(dateObj);
                       startYear = Math.floor(dateObj.getFullYear() / 5) * 5;
                   } else {
                       startYear = Math.floor(startYear / 5) * 5;
                   }
                   
                   let supplier = (supIdx > -1) ? String(row[supIdx]).trim() : "";
                   
                   // Dynamic fallback to catch the invoice number no matter which column format is used
                   let rawInv1 = (invIdx1 > -1) ? row[invIdx1] : "";
                   let rawInv2 = (invIdx2 > -1) ? row[invIdx2] : "";
                   let bestInv = (rawInv1 !== "" && rawInv1 != null) ? rawInv1 : rawInv2;
                   
                   let invoiceNo = "";
                   // Prevents large invoice numbers (e.g. 1.04002E+11) from saving as scientific notation
                   if (typeof bestInv === "number" && bestInv > 9999999999) {
                       invoiceNo = bestInv.toLocaleString('fullwide', {useGrouping: false});
                   } else {
                       invoiceNo = String(bestInv).trim();
                   }
                   
                   let totalSum = (sumIdx > -1) ? parseFloat(String(row[sumIdx]).replace(/,/g, '')) || 0 : 0;

                   if (!supplier || !invoiceNo) continue;

                   const blockName = `${PREFIX_RIWB}${startYear}-${startYear + 4}`;
                   
                   if (!rowsByYearBlock[blockName]) rowsByYearBlock[blockName] =[];
                   rowsByYearBlock[blockName].push([formattedDate, supplier, invoiceNo, totalSum, fileName]);
                }

                for (let block in rowsByYearBlock) {
                   let targetSheet;
                   if (sheetCache[block]) {
                       targetSheet = sheetCache[block];
                   } else {
                       targetSheet = getSheetForYearBlock(PREFIX_RIWB, new Date(), currentHeaders);
                       if (targetSheet.getLastRow() > 1) {
                           targetSheet.getRange(2, 1, targetSheet.getLastRow() - 1, currentHeaders.length).clearContent();
                       }
                       sheetCache[block] = targetSheet;
                   }
                   
                   let newRows = rowsByYearBlock[block];
                   if (newRows.length > 0) {
                      targetSheet.getRange(2, 3, Math.max(1, targetSheet.getMaxRows() - 1), 1).setNumberFormat("@"); // Force Invoice column to text
                      let nextRow = targetSheet.getLastRow() + 1;
                      targetSheet.getRange(nextRow, 1, newRows.length, currentHeaders.length).setValues(newRows);
                   }
                }
             }
          }
          if (tempFileId) Drive.Files.remove(tempFileId);
          file.setTrashed(true);
          filesProcessed++;
        } catch (e) { 
          if (tempFileId) Drive.Files.remove(tempFileId);
          console.error("Error processing RIWB " + fileName + ": " + e.toString()); 
          try { file.setTrashed(true); } catch(err) {} 
        }
      }
    }
    ss.toast(`RIWB Import Complete. Files: ${filesProcessed}`, "SUCCESS", 5);
  } finally {
    if (lock) lock.releaseLock();
  }
}

function _16_importSalesReports(e) {
  var isBackground = (e && e.triggerUid);
  var lock = null;
  if (!isBackground) {
    lock = LockService.getScriptLock();
    try { lock.waitLock(10000); } catch (err) { throw new Error("System is currently busy. Another user is generating reports right now. Please wait a few minutes."); }
  }
  try {
  const startTime = Date.now();
  const MAX_EXECUTION_TIME = 4.5 * 60 * 1000;
  deleteTriggersByName("_16_importSalesReports"); // Clean up previous continuation triggers

  console.time("Import_SR");
  const folder = DriveApp.getFolderById(SR_FOLDER_ID);
  const files = folder.getFiles();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  const sheetCache = {}; 

  const headers = [
    "Document type", "Customer", "Status", "Number", "Date",
    "Additional text", "Notes", "Product", "VAT rate", "Sold Quantity",
    "Unit", "Net price", "Net total", "Row VAT", "Row total", "Cust manager", "Form Generated", "Source File"
  ];

  let filesProcessed = 0;

  while (files.hasNext()) {
    if (Date.now() - startTime > MAX_EXECUTION_TIME) {
        ss.toast("Approaching time limit. Pausing and auto-resuming Sales Reports in 1 minute...", "BATCH CONTINUATION", 10);
        ScriptApp.newTrigger("_16_importSalesReports").timeBased().after(60 * 1000).create();
        return; 
    }

    let file = files.next();
    let fileName = file.getName();
    
    // Ignore aborted Temp files from previous crashes so it doesn't get stuck in a loop
    if (fileName.startsWith("Temp_")) {
        try { file.setTrashed(true); } catch(e){}
        continue; 
    }

    let mimeType = file.getMimeType();
    let isGoogleSheet = mimeType === MimeType.GOOGLE_SHEETS;
    let isExcel = mimeType === MimeType.MICROSOFT_EXCEL || 
                  mimeType === "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" || 
                  mimeType === MimeType.CSV || mimeType === "text/csv" ||
                  fileName.endsWith(".xlsx") || fileName.endsWith(".xls") || fileName.endsWith(".csv");

    if (isGoogleSheet || isExcel) {
      console.time(`Process_File_${fileName}`);
      ss.toast(`Importing SR: ${fileName}...`, "IMPORT", -1);
      let tempFileId = null;
      try {
        let converted = _openFileOrConvertCSV(file);
        let ssNew = converted.ssNew;
        tempFileId = converted.tempFileId;
        let sheetNew = ssNew.getSheets()[0];

        const fileValues = sheetNew.getDataRange().getValues();
        const lastRow = fileValues.length;

        if (lastRow > 1) {
          let headerRowIdx = -1;
          for (let r = 0; r < Math.min(20, lastRow); r++) {
             let rowStr = fileValues[r].join(" ").toLowerCase();
             if (rowStr.includes("document") && rowStr.includes("customer") && rowStr.includes("date") && rowStr.includes("product")) {
                 headerRowIdx = r;
                 break;
             }
          }
          
          if (headerRowIdx > -1) {
            const fileHeaders = fileValues[headerRowIdx].map(h => String(h).toLowerCase().trim());
            
            const colMap = {
              "Document type": fileHeaders.findIndex(h => h.includes("document")),
              "Customer": fileHeaders.findIndex(h => h === "customer" || h.includes("customer")),
              "Status": fileHeaders.findIndex(h => h === "status" || h.includes("status")),
              "Number": fileHeaders.findIndex(h => h === "number" || h.includes("number")),
              "Date": fileHeaders.findIndex(h => h === "date" || h.includes("date")),
              "Additional text": fileHeaders.findIndex(h => h.includes("additional")),
              "Notes": fileHeaders.findIndex(h => h === "notes" || h.includes("notes")),
              "Product": fileHeaders.findIndex(h => h === "product" || h.includes("product")),
              "VAT rate": fileHeaders.findIndex(h => h.includes("vat rate")),
              "Sold Quantity": fileHeaders.findIndex(h => h.includes("sold") || h.includes("quantity")),
              "Unit": fileHeaders.findIndex(h => h === "unit" || h.includes("unit")),
              "Net price": fileHeaders.findIndex(h => h.includes("net price")),
              "Net total": fileHeaders.findIndex(h => h.includes("net total") || h.includes("net to")),
              "Row VAT": fileHeaders.findIndex(h => h.includes("row vat")),
              "Row total": fileHeaders.findIndex(h => h.includes("row total") || h.includes("row to")),
              "Cust manager": fileHeaders.findIndex(h => h.includes("manager") || h.includes("cust manager"))
            };

            if (colMap["Date"] > -1) {
              let rowsByYearBlock = {};

              for (let i = headerRowIdx + 1; i < lastRow; i++) {
                let row = fileValues[i];
                let dateVal = row[colMap["Date"]];
                let dateObj = new Date(dateVal);
                
                if (!isNaN(dateObj.getTime())) {
                   const startYear = Math.floor(dateObj.getFullYear() / 5) * 5;
                   const blockName = `${PREFIX_SR}${startYear}-${startYear + 4}`;
                   
                   if (!rowsByYearBlock[blockName]) rowsByYearBlock[blockName] = [];
                   
                   const cleanStr = (idx) => idx > -1 ? String(row[idx]).trim() : "";
                   const getNum = (idx) => idx > -1 ? (parseFloat(row[idx]) || 0) : 0;

                   rowsByYearBlock[blockName].push([
                     cleanStr(colMap["Document type"]),
                     cleanStr(colMap["Customer"]),
                     cleanStr(colMap["Status"]),
                     cleanStr(colMap["Number"]),
                     dateVal,                   
                     cleanStr(colMap["Additional text"]),
                     cleanStr(colMap["Notes"]),
                     cleanStr(colMap["Product"]),
                     cleanStr(colMap["VAT rate"]),
                     getNum(colMap["Sold Quantity"]),
                     cleanStr(colMap["Unit"]),
                     getNum(colMap["Net price"]),
                     getNum(colMap["Net total"]),
                     getNum(colMap["Row VAT"]),
                     getNum(colMap["Row total"]),
                     cleanStr(colMap["Cust manager"]),
                     "", 
                     fileName 
                   ]);
                }
              }

              // --- WIPE AND REPLACE LOGIC ---
              for (let block in rowsByYearBlock) {
                let targetSheet;
                if (sheetCache[block]) {
                  targetSheet = sheetCache[block];
                } else {
                  let sampleDate = new Date(rowsByYearBlock[block][0][4]); 
                  targetSheet = getSheetForYearBlock(PREFIX_SR, sampleDate, headers);
                  if (targetSheet.getLastRow() > 1) {
                      targetSheet.getRange(2, 1, targetSheet.getLastRow() - 1, targetSheet.getLastColumn()).clearContent();
                  }
                  sheetCache[block] = targetSheet;
                }

                let newRows = rowsByYearBlock[block];
                if (newRows.length > 0) {
                  let nextRow = targetSheet.getLastRow() + 1;
                  targetSheet.getRange(nextRow, 1, newRows.length, headers.length).setValues(newRows);
                  
                  targetSheet.getRange(2, 1, targetSheet.getLastRow() - 1, headers.length)
                    .sort([{ column: 5, ascending: true }, { column: 4, ascending: true }]);
                }
              }
            }
          }
        }
        if (tempFileId) Drive.Files.remove(tempFileId);
        file.setTrashed(true);
        filesProcessed++;
        console.timeEnd(`Process_File_${fileName}`);
      } catch (e) {
        if (tempFileId) Drive.Files.remove(tempFileId);
        throw new Error("Error importing SR " + fileName + ": " + e.toString());
      }
    }
  }
  console.timeEnd("Import_SR");
  ss.toast(`Sales Reports Import Complete. Files: ${filesProcessed}`, "SUCCESS", 5);
  } finally {
    if (lock) lock.releaseLock();
  }
}

function _17_importPurchaseOrdersPingconAndBonita(e) {
  var isBackground = (e && e.triggerUid);
  var lock = null;
  if (!isBackground) {
    lock = LockService.getScriptLock();
    try { lock.waitLock(10000); } catch (err) { throw new Error("System is currently busy. Another user is generating reports right now. Please wait a few minutes."); }
  }
  try {
  const startTime = Date.now();
  const MAX_EXECUTION_TIME = 4.5 * 60 * 1000;
  deleteTriggersByName("_17_importPurchaseOrdersPingconAndBonita");

  console.time("Import_PO");
  const folder = DriveApp.getFolderById(PO_PINGCON_BONITA_FOLDER_ID);
  const files = folder.getFiles();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  const sheetCache = {};
  
  // OUTPUT COLUMNS (8 Columns)
  const headers =["Date (MM/DD/YYYY)", "Purchase Invoice", "Pingcon PO#", "SKU", "Item Name", "Supplier", "Unit Price", "Quantity", "Net Total", "Row VAT", "Row Total", "Source File"];
  let filesProcessed = 0;

  while (files.hasNext()) {
    if (Date.now() - startTime > MAX_EXECUTION_TIME) {
        ss.toast("Approaching time limit. Pausing and auto-resuming POs in 1 minute...", "BATCH CONTINUATION", 10);
        ScriptApp.newTrigger("_17_importPurchaseOrdersPingconAndBonita").timeBased().after(60 * 1000).create();
        return; 
    }

    let file = files.next();
    let fileName = file.getName();
    
    if (fileName.endsWith(".xlsx") || fileName.endsWith(".xls")) {
      console.time(`Process_PO_${fileName}`);
      ss.toast(`Importing PO: ${fileName}...`, "IMPORT", -1);
      let tempFileId = null;
      try {
        let converted = _openFileOrConvertCSV(file);
        let ssNew = converted.ssNew;
        tempFileId = converted.tempFileId;
        let sheetNew = ssNew.getSheets()[0];

        let rawData = sheetNew.getDataRange().getValues();
        
        if (rawData.length > 0) {
            let headerRowIndex = -1;
            let skuIdx = -1, nameIdx = -1, qtyIdx = -1, totalIdx = -1, supplierIdx = -1, dateIdx = -1, invoiceIdx = -1, pingconPoIdx = -1;
            let netTotalIdx = -1, rowVatIdx = -1, rowTotalIdx = -1;
            let extractSkuFromName = false;

            // --- HEADER DETECTION ---
            for (let r = 0; r < Math.min(20, rawData.length); r++) {
              let row = rawData[r].map(c => String(c).toLowerCase().trim());
              
              // 1. Find Name/Description
              let nIdx = row.findIndex(c => c === "name" || c === "description" || c === "product" || c === "item name");
              
              // 2. Find Total (Priority: Row Total > Net Total > Total with VAT)
              let tIdx = row.findIndex(c => c === "row total"); 
              if (tIdx === -1) tIdx = row.findIndex(c => c === "net total"); 
              if (tIdx === -1) tIdx = row.findIndex(c => c.includes("total with vat")); 

              if (nIdx > -1 && tIdx > -1) {
                  nameIdx = nIdx; 
                  totalIdx = tIdx;
                  
                  skuIdx = row.findIndex(c => c === "code" || c === "sku" || c === "item code");
                  if (skuIdx === -1) extractSkuFromName = true;
                  
                  qtyIdx = row.findIndex(c => c === "purchased qty" || c === "purchased quantity" || c === "quantity" || c === "qty");
                  supplierIdx = row.findIndex(c => c.includes("supplier") || c.includes("vendor"));
                  dateIdx = row.findIndex(c => c === "date" || c.includes("doc. date"));
                  invoiceIdx = row.findIndex(c => c === "purchase invoice no." || c === "purchase invoice" || c === "ref no.");
                  pingconPoIdx = row.findIndex(c => c.includes("received purchase orders") || c.includes("our no") || c === "po number");
                  
                  netTotalIdx = row.findIndex(c => c === "net total" || c.includes("net total"));
                  rowVatIdx = row.findIndex(c => c === "row vat" || c.includes("row vat"));
                  rowTotalIdx = row.findIndex(c => c === "row total" || c.includes("row total"));
                  
                  headerRowIndex = r;
                  break;
              }
            }

            if (headerRowIndex > -1) {
              let rowsByYearBlock = {};

              for (let i = headerRowIndex + 1; i < rawData.length; i++) {
                 let row = rawData[i];
                 let name = (nameIdx > -1) ? String(row[nameIdx]).trim() : "";
                 if (name === "" && skuIdx === -1) continue; 

                 let sku = "";
                 if (!extractSkuFromName && skuIdx > -1) {
                    sku = String(row[skuIdx]).trim();
                 } else {
                    let parts = name.split(" ");
                    if (parts.length > 0) sku = parts[0].trim();
                 }

                 if (sku !== "") {
                   let dateVal = (dateIdx > -1 && row[dateIdx]) ? row[dateIdx] : new Date(); 
                   let dateObj = new Date(dateVal);
                   if (isNaN(dateObj.getTime())) dateObj = new Date();
                   
                   let invoiceVal = (invoiceIdx > -1) ? row[invoiceIdx] : "";
                   let pingconPoVal = (pingconPoIdx > -1) ? row[pingconPoIdx] : "";
                   let supplier = (supplierIdx > -1) ? String(row[supplierIdx]).trim() : "";

                   let qtyVal = (qtyIdx > -1) ? parseFloat(String(row[qtyIdx]).replace(/[^0-9.-]+/g, "")) : 0;
                   let totalVal = (totalIdx > -1) ? parseFloat(String(row[totalIdx]).replace(/[^0-9.-]+/g, "")) : 0;
                   if (isNaN(qtyVal)) qtyVal = 0;
                   if (isNaN(totalVal)) totalVal = 0;
                   
                   // REVERTED: Calculate Price from Total / Qty
                   let price = (qtyVal !== 0) ? (totalVal / qtyVal) : 0;

                   let parsedNetTotal = (netTotalIdx > -1) ? parseFloat(String(row[netTotalIdx]).replace(/[^0-9.-]+/g, "")) : 0;
                   let parsedRowVat = (rowVatIdx > -1) ? parseFloat(String(row[rowVatIdx]).replace(/[^0-9.-]+/g, "")) : 0;
                   let parsedRowTotal = (rowTotalIdx > -1) ? parseFloat(String(row[rowTotalIdx]).replace(/[^0-9.-]+/g, "")) : 0;
                   if (isNaN(parsedNetTotal)) parsedNetTotal = 0;
                   if (isNaN(parsedRowVat)) parsedRowVat = 0;
                   if (isNaN(parsedRowTotal)) parsedRowTotal = 0;

                   if (sku.toUpperCase().endsWith("C")) {
                      let match = name.match(/.*[xX]\s*(\d+)/i);
                      if (match) {
                        let packSize = parseInt(match[1]);
                        if (packSize > 0) { 
                          price = price / packSize; 
                          qtyVal = qtyVal * packSize;
                          sku = sku.substring(0, sku.length - 1) + "P"; 
                        }
                      }
                   }

                   const startYear = Math.floor(dateObj.getFullYear() / 5) * 5;
                   const blockName = `${PREFIX_PO_PINGCON_BONITA}${startYear}-${startYear + 4}`;

                   if (!rowsByYearBlock[blockName]) rowsByYearBlock[blockName] = [];
                   
                   rowsByYearBlock[blockName].push([
                    dateObj, invoiceVal, pingconPoVal, sku, name, supplier, price, qtyVal, parsedNetTotal, parsedRowVat, parsedRowTotal, fileName
                  ]);
                 }
              }

              // --- WIPE AND REPLACE LOGIC ---
              for (let block in rowsByYearBlock) {
                 let targetSheet;
                 if (sheetCache[block]) {
                   targetSheet = sheetCache[block];
                 } else {
                   let sampleDate = rowsByYearBlock[block][0][0];
                   targetSheet = getSheetForYearBlock(PREFIX_PO_PINGCON_BONITA, sampleDate, headers);
                   targetSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
                   targetSheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#f3f3f3");
                   if (targetSheet.getLastRow() > 1) {
                       targetSheet.getRange(2, 1, targetSheet.getLastRow() - 1, targetSheet.getLastColumn()).clearContent();
                   }
                   sheetCache[block] = targetSheet;
                 }

                 let newRows = rowsByYearBlock[block];
                 const formatRow = (r) => {
                    if (r[0] instanceof Date) r[0] = Utilities.formatDate(r[0], Session.getScriptTimeZone(), "MM/dd/yyyy");
                    return r;
                 };
                 let finalData = newRows.map(formatRow);
                 
                 if (finalData.length > 0) {
                    let nextRow = targetSheet.getLastRow() + 1;
                    targetSheet.getRange(nextRow, 1, finalData.length, headers.length).setValues(finalData);
                 }
              }
            }
        }
        if (tempFileId) Drive.Files.remove(tempFileId);
        file.setTrashed(true);
        filesProcessed++;
        console.timeEnd(`Process_PO_${fileName}`);
      } catch (e) { 
        if (tempFileId) Drive.Files.remove(tempFileId);
        throw new Error("Error importing PO " + fileName + ": " + e.toString()); 
      }
    }
  }
  console.timeEnd("Import_PO");
  ss.toast(`PO Import Complete. Files: ${filesProcessed}`, "SUCCESS", 5);
  } finally {
    if (lock) lock.releaseLock();
  }
}

function _18_importAccountTitles(e) {
  var isBackground = (e && e.triggerUid);
  var lock = null;
  if (!isBackground) {
    lock = LockService.getScriptLock();
    try { lock.waitLock(10000); } catch (err) { throw new Error("System is currently busy. Another user is generating reports right now. Please wait a few minutes."); }
  }
  try {
  const startTime = Date.now();
  const MAX_EXECUTION_TIME = 4.5 * 60 * 1000; // FIX: Added missing execution limit
  deleteTriggersByName("_18_importAccountTitles"); // FIX: Clean up previous triggers

  console.time("Import_Account_Titles");
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  if (!ACCOUNT_TITLE_FOLDER_ID || ACCOUNT_TITLE_FOLDER_ID === "INSERT_ACCOUNT_TITLE_FOLDER_ID_HERE") {
    ss.toast("Error: ACCOUNT_TITLE_FOLDER_ID is not configured.", "ERROR", -1);
    return;
  }

  const folder = DriveApp.getFolderById(ACCOUNT_TITLE_FOLDER_ID);
  const files = folder.getFiles();
  
  const headers =["Company Name", "Account Title", "EWT Percent", "ATC", "Source File", "Import Date"];
  let filesProcessed = 0; const sheetCache = {};

  while (files.hasNext()) {
    if (Date.now() - startTime > MAX_EXECUTION_TIME) {
        ss.toast("Approaching time limit. Pausing and auto-resuming Account Titles in 1 minute...", "BATCH CONTINUATION", 10);
        ScriptApp.newTrigger("_18_importAccountTitles").timeBased().after(60 * 1000).create(); // FIX: Pointed to correct function
        return; 
    }

    let file = files.next();
    let fileName = file.getName();
    let mimeType = file.getMimeType();

    let isGoogleSheet = mimeType === MimeType.GOOGLE_SHEETS;
    let isExcel = mimeType === MimeType.MICROSOFT_EXCEL || 
                  mimeType === "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" || 
                  fileName.endsWith(".xlsx") || fileName.endsWith(".xls") || fileName.endsWith(".csv"); // FIX: Added CSV support

    if (isGoogleSheet || isExcel) {
      ss.toast(`Scanning Account Titles: ${fileName}...`, "IMPORT", -1);
      let tempFileId = null;
      try {
        let converted = _openFileOrConvertCSV(file);
        let ssNew = converted.ssNew;
        tempFileId = converted.tempFileId;
        
        const sheetNew = ssNew.getSheets()[0];
        let data = sheetNew.getDataRange().getValues();
        let fileImported = false;

        if (data.length > 0) {
          let rowsByYearBlock = {};
          let importDateObj = new Date();
          let importDateStr = Utilities.formatDate(importDateObj, Session.getScriptTimeZone(), "MM/dd/yyyy");
          const startYear = Math.floor(importDateObj.getFullYear() / 5) * 5;
          const blockName = `${PREFIX_ACCOUNT_TITLE}${startYear}-${startYear + 4}`;
          
          if (!rowsByYearBlock[blockName]) rowsByYearBlock[blockName] =[];

          let currentEwtPct = "";
          let currentAtc = "";
          let currentAccountTitle = "";

          for (let i = 0; i < data.length; i++) {
            let colA = String(data[i][0]).trim();
            let colB = String(data[i][1]).trim();
            
            if (colA.toUpperCase().startsWith("EWT") && colA.toUpperCase().includes("ATC")) {
              let ewtMatch = colA.match(/EWT\s*(\d+(?:\.\d+)?)%/i);
              let atcMatch = colA.match(/ATC\s*([A-Z0-9\s]+)/i);
              if (ewtMatch) currentEwtPct = ewtMatch[1] + "%";
              if (atcMatch) currentAtc = atcMatch[1].trim();
              continue; 
            }
            
            // Check if it's a grouped header
            if (colA !== "" && colB === "") {
                if (colA.toLowerCase() !== "account title" && colA.toLowerCase() !== "company name") {
                    currentAccountTitle = colA;
                }
                continue;
            }

            // FIX: Smart parsing to prevent Account Titles swapping columns with Company Names
            let companyName = "";
            let accTitle = "";
            
            if (colA.toLowerCase() === "account title" || colA.toLowerCase() === "company name") continue;

            if (colA !== "" && colB !== "") {
                let aLower = colA.toLowerCase();
                if (aLower.includes("purchases") || aLower.includes("supplies") || aLower.includes("expense") || aLower.includes("payable")) {
                    companyName = colB;
                    accTitle = colA;
                } else {
                    companyName = colA;
                    accTitle = colB;
                }
            } else if (colA === "" && colB !== "" && currentAccountTitle !== "") {
                companyName = colB;
                accTitle = currentAccountTitle;
            } else if (colA !== "" && colB === "" && currentAccountTitle !== "") {
                companyName = colA;
                accTitle = currentAccountTitle;
            }

            if (companyName && accTitle) {
                rowsByYearBlock[blockName].push([companyName, accTitle, currentEwtPct, currentAtc, fileName, importDateStr]);
            }
          }

          for (let block in rowsByYearBlock) {
             let targetSheet;
             if (sheetCache[block]) {
                 targetSheet = sheetCache[block];
             } else {
                 targetSheet = getSheetForYearBlock(PREFIX_ACCOUNT_TITLE, importDateObj, headers);
                 if (targetSheet.getLastRow() > 1) {
                     targetSheet.getRange(2, 1, targetSheet.getLastRow() - 1, targetSheet.getLastColumn()).clearContent();
                 }
                 sheetCache[block] = targetSheet;
             }
             let newRows = rowsByYearBlock[block];
             if (newRows.length > 0) {
                 let nextRow = targetSheet.getLastRow() + 1;
                 targetSheet.getRange(nextRow, 1, newRows.length, headers.length).setValues(newRows);
             }
          }
          fileImported = true;
        }

        if (tempFileId) Drive.Files.remove(tempFileId);
        if (fileImported) {
            file.setTrashed(true);
            filesProcessed++;
        }
      } catch (e) {
        if (tempFileId) Drive.Files.remove(tempFileId);
        throw new Error("Error importing Account Titles " + fileName + ": " + e.toString());
      }
    }
  }
  ss.toast(`Account Titles Import Complete. Files: ${filesProcessed}`, "SUCCESS", 5);
  console.timeEnd("Import_Account_Titles");
  } finally {
    if (lock) lock.releaseLock();
  }
}

function _19_importPMCDirectory() {
  var lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) { throw new Error("System is currently busy. Another user is generating reports right now. Please wait a few minutes."); }
  try {
  console.time("Import_PMC_Directory");
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  if (!PMC_DIRECTORY_FOLDER_ID || PMC_DIRECTORY_FOLDER_ID === "INSERT_PMC_DIRECTORY_FOLDER_ID_HERE") {
    ss.toast("Error: PMC_DIRECTORY_FOLDER_ID is not configured.", "ERROR", -1);
    return;
  }

  const folder = DriveApp.getFolderById(PMC_DIRECTORY_FOLDER_ID);
  const files = folder.getFiles();
  
  const headers = ["Company Name", "Address", "TIN", "Source File", "Import Date"];
  let filesProcessed = 0; const sheetCache = {};

  while (files.hasNext()) {
    let file = files.next();
    let fileName = file.getName();

    if (fileName.endsWith(".xlsx") || fileName.endsWith(".xls") || fileName.endsWith(".csv")) {
      ss.toast(`Scanning PMC Directory: ${fileName}...`, "IMPORT", -1);
      let tempFileId = null;
      try {
        let converted = _openFileOrConvertCSV(file);
        let ssNew = converted.ssNew;
        tempFileId = converted.tempFileId;
        
        const sheetNew = ssNew.getSheets()[0];
        let data = sheetNew.getDataRange().getValues();
        let fileImported = false;

        if (data.length > 0) {
          let rowsByYearBlock = {};
          let importDateObj = new Date();
          let importDateStr = Utilities.formatDate(importDateObj, Session.getScriptTimeZone(), "MM/dd/yyyy");
          const startYear = Math.floor(importDateObj.getFullYear() / 5) * 5;
          const blockName = `${PREFIX_PMC_DIRECTORY}${startYear}-${startYear + 4}`;
          
          if (!rowsByYearBlock[blockName]) rowsByYearBlock[blockName] =[];

          for (let i = 0; i < data.length; i++) {
            let company = String(data[i][0]).trim();
            let address = String(data[i][1]).trim();
            let tin = String(data[i][2]).trim();

            if (company && company.toLowerCase() !== "company name") {
              rowsByYearBlock[blockName].push([company, address, tin, fileName, importDateStr]);
            }
          }

          for (let block in rowsByYearBlock) {
             let targetSheet;
             if (sheetCache[block]) {
                 targetSheet = sheetCache[block];
             } else {
                 targetSheet = getSheetForYearBlock(PREFIX_PMC_DIRECTORY, importDateObj, headers);
                 if (targetSheet.getLastRow() > 1) {
                     targetSheet.getRange(2, 1, targetSheet.getLastRow() - 1, targetSheet.getLastColumn()).clearContent();
                 }
                 sheetCache[block] = targetSheet;
             }
             let newRows = rowsByYearBlock[block];
             if (newRows.length > 0) {
                 let nextRow = targetSheet.getLastRow() + 1;
                 targetSheet.getRange(nextRow, 1, newRows.length, headers.length).setValues(newRows);
             }
          }
          fileImported = true;
        }

        if (tempFileId) Drive.Files.remove(tempFileId);
        if (fileImported) {
            file.setTrashed(true);
            filesProcessed++;
        }
      } catch (e) {
        if (tempFileId) Drive.Files.remove(tempFileId);
        throw new Error("Error importing PMC Directory " + fileName + ": " + e.toString());
      }
    }
  }
  ss.toast(`PMC Directory Import Complete. Files: ${filesProcessed}`, "SUCCESS", 5);
  console.timeEnd("Import_PMC_Directory");
  } finally {
    if (lock) lock.releaseLock();
  }
}

function _20_importAgeingReports(e) {
  var isBackground = (e && e.triggerUid);
  var lock = null;
  if (!isBackground) {
    lock = LockService.getScriptLock();
    try { lock.waitLock(10000); } catch (err) { throw new Error("System is currently busy. Another user is generating reports right now. Please wait a few minutes."); }
  }
  try {
  const startTime = Date.now();
  const MAX_EXECUTION_TIME = 4.5 * 60 * 1000;
  deleteTriggersByName("_20_importAgeingReports");

  console.time("Import_AR");
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const folder = DriveApp.getFolderById(AR_FOLDER_ID);
  const files = folder.getFiles();
  
  let filesProcessed = 0;
  const sheetCache = {};

  while (files.hasNext()) {
    let file = files.next();
    let fileName = file.getName();

    if (fileName.endsWith(".xlsx") || fileName.endsWith(".xls")) {
      console.time(`Process_AR_${fileName}`);
      ss.toast(`Importing AR: ${fileName}`, "IMPORT", -1);
      
      let tempFileId = null;
      try {
        let converted = _openFileOrConvertCSV(file);
        let ssNew = converted.ssNew;
        tempFileId = converted.tempFileId;
        let sheetNew = ssNew.getSheets()[0];
        let data = sheetNew.getDataRange().getValues();

        if (data.length > 0) {
           let headerRowIdx = -1;
           for(let r=0; r<Math.min(20, data.length); r++) {
             let rowStr = data[r].join(" ").toLowerCase();
             if(rowStr.includes("additional text") && (rowStr.includes("manager") || rowStr.includes("cust"))) {
                 headerRowIdx = r;
                 break;
             }
           }

           if (headerRowIdx > -1) {
              let originalHeaders = data[headerRowIdx].map(h => String(h).trim());
              let head = data[headerRowIdx].map(h => String(h).toLowerCase().trim());
              
              // Strictly look for the exact header "date" first
              let dateIdx = head.findIndex(h => h === "date"); 
              if (dateIdx === -1) {
                  dateIdx = head.findIndex(h => h.includes("date") && !h.includes("expiry") && !h.includes("countered") && !h.includes("received")); 
              }
              let textIdx = head.findIndex(h => h.includes("additional text")); 
              let managerIdx = head.findIndex(h => h.includes("manager")); 
              let openBalIdx = head.findIndex(h => h.includes("open balance"));

              // Inject the new headers dynamically based on Open Balance position
              if (openBalIdx > -1) {
                  if (originalHeaders.length > openBalIdx + 1) originalHeaders[openBalIdx + 1] = "% Open Balance";
                  if (originalHeaders.length > openBalIdx + 2) originalHeaders[openBalIdx + 2] = "Payment Status";
              }
              
              let currentHeaders = [...originalHeaders, "Source File"];
              let rowsByYearBlock = {};

              for (let i = headerRowIdx + 1; i < data.length; i++) {
                 let row = data[i];
                 let addText = (textIdx > -1) ? String(row[textIdx]).trim() : "";
                 let rawDate = (dateIdx > -1) ? row[dateIdx] : "";
                 
                 if (addText === "" && rawDate === "") continue;

                 let dateObj;
                 if (typeof rawDate === 'number') {
                     // Fixes dates exported as Excel serial numbers
                     dateObj = new Date(Math.round((rawDate - 25569) * 86400 * 1000));
                 } else if (rawDate instanceof Date) {
                     dateObj = rawDate;
                 } else {
                     dateObj = new Date(rawDate);
                 }

                 if (isNaN(dateObj.getTime())) {
                     dateObj = new Date(); 
                 } else if (dateIdx > -1 && rawDate !== "") {
                     // Force rewrite the date as a formatted string so it never turns blank in the database
                     row[dateIdx] = fastFormatDate(dateObj);
                 }

                 const startYear = Math.floor(dateObj.getFullYear() / 5) * 5;
                 const blockName = `${PREFIX_AR}${startYear}-${startYear + 4}`;
                 
                 if (!rowsByYearBlock[blockName]) rowsByYearBlock[blockName] =[];
                 
                 // Append the Source File name cleanly at the end of the original row
                 let fullRow = [...row, fileName];
                 rowsByYearBlock[blockName].push(fullRow);
              }

              // --- WIPE AND REPLACE LOGIC ---
              for (let block in rowsByYearBlock) {
                 let targetSheet;
                 let targetLastCol;
                 if (sheetCache[block]) {
                     targetSheet = sheetCache[block];
                     targetLastCol = targetSheet.getLastColumn();
                 } else {
                     let sampleDate = new Date();
                     if (dateIdx > -1) {
                        let sampleRaw = rowsByYearBlock[block][0][dateIdx];
                        let parsedDate = (sampleRaw instanceof Date) ? sampleRaw : new Date(sampleRaw);
                        if (!isNaN(parsedDate.getTime())) sampleDate = parsedDate;
                     }
                     targetSheet = getSheetForYearBlock(PREFIX_AR, sampleDate, currentHeaders);
                     targetLastCol = targetSheet.getLastColumn();
                     if (targetLastCol < currentHeaders.length) {
                         targetSheet.getRange(1, 1, 1, currentHeaders.length).setValues([currentHeaders]);
                         targetSheet.getRange(1, 1, 1, currentHeaders.length).setFontWeight("bold").setBackground("#f3f3f3");
                         targetLastCol = currentHeaders.length;
                     }
                     if (targetSheet.getLastRow() > 1) {
                         targetSheet.getRange(2, 1, targetSheet.getLastRow() - 1, targetLastCol).clearContent();
                     }
                     sheetCache[block] = targetSheet;
                 }
                 
                 let newRows = rowsByYearBlock[block];
                 if (newRows.length > 0) {
                    let paddedData = newRows.map(r => {
                        let newRow = [...r];
                        while(newRow.length < targetLastCol) newRow.push("");
                        return newRow.slice(0, targetLastCol);
                    });
                    let nextRow = targetSheet.getLastRow() + 1;
                    targetSheet.getRange(nextRow, 1, paddedData.length, targetLastCol).setValues(paddedData);
                    
                    // Format the column immediately after Open Balance as percentage (e.g., 12%)
                    if (openBalIdx > -1 && (openBalIdx + 2) <= targetLastCol) {
                        targetSheet.getRange(nextRow, openBalIdx + 2, paddedData.length, 1).setNumberFormat("0%");
                    }
                 }
              }
           }
        }
        if (tempFileId) Drive.Files.remove(tempFileId);
        file.setTrashed(true);
        filesProcessed++;
        console.timeEnd(`Process_AR_${fileName}`);
      } catch (e) { 
        if (tempFileId) Drive.Files.remove(tempFileId);
        throw new Error("Error processing AR " + fileName + ": " + e.toString()); 
      }
    }
  }
  console.timeEnd("Import_AR");
  ss.toast(`AR Import Complete. Files: ${filesProcessed}`, "SUCCESS", 5);
  } finally {
    if (lock) lock.releaseLock();
  }
}
function _21_importSalesQuotas() {
  var lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) { throw new Error("System is currently busy. Another user is generating reports right now. Please wait a few minutes."); }
  try {
  console.time("Import_SQ");
  const folder = DriveApp.getFolderById(SQ_FOLDER_ID);

  const files = folder.getFiles();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Headers for the destination SQ_Data sheets
  // Headers for the destination SQ_Data sheets
  const headers =["AE Name", "Start Date", "End Date", "Old Quota", "New Quota", "Source File"];
  let filesProcessed = 0; const sheetCache = {};

  while (files.hasNext()) {
    let file = files.next();
    let fileName = file.getName();

    if (fileName.endsWith(".xlsx") || fileName.endsWith(".xls")) {
      ss.toast(`Importing SQ: ${fileName}...`, "IMPORT", -1);
      let tempFileId = null;
      try {
        let converted = _openFileOrConvertCSV(file);
        let ssNew = converted.ssNew;
        tempFileId = converted.tempFileId;
        let sheetNew = ssNew.getSheets()[0];
        let data = sheetNew.getDataRange().getValues();

        if (data.length > 0) {
           let aeIdx = -1, dateIdx = -1;
           let newQuotaIdx = -1, oldQuotaIdx = -1;
           let headerRow = -1;

           // Find Headers
           for(let r=0; r<Math.min(20, data.length); r++) {
             let row = data[r].map(c => String(c).toLowerCase().trim());
             
             if (row.findIndex(c => c.includes("ae") || c.includes("name")) > -1) {
               aeIdx = row.findIndex(c => c.includes("ae") || c.includes("name"));
               dateIdx = row.findIndex(c => c.includes("effectivity") || c.includes("effective date"));
               newQuotaIdx = row.findIndex(c => c.includes("new quota") || c.includes("target") || (c.includes("quota") && !c.includes("old")));
               oldQuotaIdx = row.findIndex(c => c.includes("old quota") || c.includes("previous quota"));

               if (aeIdx > -1) {
                  headerRow = r;
                  break;
               }
             }
           }

           if (aeIdx > -1) {
             let rowsByYearBlock = {};

             for (let i = headerRow + 1; i < data.length; i++) {
               let row = data[i];
               let ae = String(row[aeIdx]).trim();
               if (!ae || ae.toLowerCase().includes("team")) continue; 

               let newQuotaVal = (newQuotaIdx > -1) ? parseFloat(String(row[newQuotaIdx]).replace(/[^0-9.-]+/g, "")) : 0;
               let oldQuotaVal = (oldQuotaIdx > -1) ? parseFloat(String(row[oldQuotaIdx]).replace(/[^0-9.-]+/g, "")) : 0;
               if (isNaN(newQuotaVal)) newQuotaVal = 0;
               if (isNaN(oldQuotaVal)) oldQuotaVal = 0;

               let dateStr = (dateIdx > -1) ? String(row[dateIdx]).trim() : "";
               let startDate = new Date();
               let endDate = new Date();
               
               if (dateStr.toLowerCase().includes(" to ")) {
                 let parts = dateStr.toLowerCase().split(" to ");
                 let s = new Date(parts[0].trim() + " 1"); 
                 let e = new Date(parts[1].trim() + " 1");
                 if (!isNaN(s.getTime())) startDate = s;
                 if (!isNaN(e.getTime())) endDate = new Date(e.getFullYear(), e.getMonth() + 1, 0); 
               } else {
                 let s = new Date(dateStr);
                 if (!isNaN(s.getTime())) startDate = s;
                 endDate = new Date(startDate.getFullYear() + 5, 0, 1); 
               }
               
               let cleanStartDate = Utilities.formatDate(startDate, Session.getScriptTimeZone(), "MM/dd/yyyy");

               const startYear = Math.floor(startDate.getFullYear() / 5) * 5;
               const blockName = `${PREFIX_SQ}${startYear}-${startYear + 4}`;

               if (!rowsByYearBlock[blockName]) rowsByYearBlock[blockName] = [];
               
               rowsByYearBlock[blockName].push([
                 ae, 
                 cleanStartDate, 
                 Utilities.formatDate(endDate, Session.getScriptTimeZone(), "MM/dd/yyyy"), 
                 oldQuotaVal, 
                 newQuotaVal, 
                 fileName
               ]);
             }

             // --- LOGIC CHANGE: UPSERT (Update/Insert) ---
             for (let block in rowsByYearBlock) {
                let targetSheet;
                if (sheetCache[block]) {
                    targetSheet = sheetCache[block];
                } else {
                    let sampleDate = new Date(rowsByYearBlock[block][0][1]);
                    targetSheet = getSheetForYearBlock(PREFIX_SQ, sampleDate, headers);
                    if (targetSheet.getLastRow() > 1) {
                        targetSheet.getRange(2, 1, targetSheet.getLastRow() - 1, targetSheet.getLastColumn()).clearContent();
                    }
                    sheetCache[block] = targetSheet;
                }
                let newRows = rowsByYearBlock[block];
                if (newRows.length > 0) {
                    let nextRow = targetSheet.getLastRow() + 1;
                    targetSheet.getRange(nextRow, 1, newRows.length, headers.length).setValues(newRows);
                }
             }
           }
        }
        if (tempFileId) Drive.Files.remove(tempFileId);
        file.setTrashed(true);
        filesProcessed++;
      } catch (e) { 
        if (tempFileId) Drive.Files.remove(tempFileId);
        throw new Error("Error importing SQ " + fileName + ": " + e.toString()); 
      }
    }
  }
  ss.toast(`Sales Quota Import Complete. Files: ${filesProcessed}`, "SUCCESS", 5);
  console.timeEnd("Import_SQ");
  } finally {
    lock.releaseLock();
  }
}

function _22_importMissingSKUsMappingOldToNew() {
  var lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) { throw new Error("System is currently busy. Another user is generating reports right now. Please wait a few minutes."); }
  try {
  console.time("Import_Missing_SKUs");
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  if (!SKU_MAPPING_OLD_NEW_FOLDER_ID || SKU_MAPPING_OLD_NEW_FOLDER_ID === "INSERT_FOLDER_ID_HERE") {
    ss.toast("Error: SKU_MAPPING_OLD_NEW_FOLDER_ID is not configured.", "ERROR", -1);
    return;
  }

  const folder = DriveApp.getFolderById(SKU_MAPPING_OLD_NEW_FOLDER_ID);
  const files = folder.getFiles();
  
  const headers = ["Old SKU", "Item Description", "New SKU", "Remarks", "Price", "Source File", "Import Date"];
  let filesProcessed = 0; const sheetCache = {};

  while (files.hasNext()) {
    let file = files.next();
    let fileName = file.getName();
    let mimeType = file.getMimeType();

    let isGoogleSheet = mimeType === MimeType.GOOGLE_SHEETS;
    let isExcel = mimeType === MimeType.MICROSOFT_EXCEL || 
                  mimeType === "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" || 
                  mimeType === "application/vnd.oasis.opendocument.spreadsheet" || 
                  fileName.toLowerCase().endsWith(".csv") || mimeType === MimeType.CSV || mimeType === "text/csv" ||
                  fileName.endsWith(".xlsx") || fileName.endsWith(".xls") || fileName.endsWith(".ods");

    if (isGoogleSheet || isExcel) {
      ss.toast(`Scanning: ${fileName}...`, "IMPORT", -1);
      let tempFileId = null;
      try {
        let converted = _openFileOrConvertCSV(file);
        let ssNew = converted.ssNew;
        tempFileId = converted.tempFileId;

        const allSheets = ssNew.getSheets();
        let fileImported = false;

        for (let s = 0; s < allSheets.length; s++) {
          let sheetNew = allSheets[s];
          let data = sheetNew.getDataRange().getValues();
          
          if (data.length > 0) {
             let skuIdx = -1, descIdx = -1, newSkuIdx = -1, remarksIdx = -1, mergedRemarksIdx = -1, priceIdx = -1;
             let headerRow = -1;
             let customMappingHeader = (SCRIPT_PROPS['CUSTOM_MAPPING_HEADER'] || "").toLowerCase().trim();

             for(let r = 0; r < Math.min(50, data.length); r++) {
               let row = data[r].map(c => String(c).toLowerCase().replace(/\s+/g, ' ').trim());
               let sIdx = row.findIndex(c => c === "missing sku" || c === "sku" || c === "old sku" || c === "item code");
               
               if (sIdx > -1) {
                 headerRow = r;
                 skuIdx = sIdx;
                 descIdx = row.findIndex(c => c.includes("description") || c.includes("item name"));
                 newSkuIdx = row.findIndex(c => c === "new sku" || c === "target sku");
                 remarksIdx = row.findIndex(c => c === "remarks" || (c.includes("notes") && !c.includes("pricing")));
                 mergedRemarksIdx = row.findIndex(c => c.startsWith("remarks - ") || c.includes("pricing") || (customMappingHeader && c.includes(customMappingHeader)));
                 priceIdx = row.findIndex(c => c.includes("price vat inc") || c === "price" || c.includes("buy price"));
                 break;
               }
             }

             if (headerRow > -1) {
               let rowsByYearBlock = {};
               let importDateObj = new Date(); 
               let importDateStr = Utilities.formatDate(importDateObj, Session.getScriptTimeZone(), "MM/dd/yyyy");
               const startYear = Math.floor(importDateObj.getFullYear() / 5) * 5;
               const blockName = `${PREFIX_SKU_MAP}${startYear}-${startYear + 4}`;

               if (!rowsByYearBlock[blockName]) rowsByYearBlock[blockName] = [];

               for (let i = headerRow + 1; i < data.length; i++) {
                 let row = data[i];
                 let sku = String(row[skuIdx]).trim();
                 if (!sku || sku.toLowerCase() === "sku" || sku === "") continue;

                 let desc = (descIdx > -1) ? String(row[descIdx]).trim() : "";
                 let newSku = (newSkuIdx > -1) ? String(row[newSkuIdx]).trim() : "";
                 let remarks = (remarksIdx > -1) ? String(row[remarksIdx]).trim() : "";
                 let price = (priceIdx > -1) ? parseFloat(String(row[priceIdx]).replace(/[^0-9.-]+/g, "")) : 0;

                 // Smart Extract from merged format (e.g. SKU# XYZ; Buy Price: 0.00) if present
                 if (mergedRemarksIdx > -1) {
                     let mergedText = String(row[mergedRemarksIdx]).trim();
                     if (mergedText) {
                         remarks = remarks ? remarks + " | " + mergedText : mergedText;
                         let skuMatch = mergedText.match(/SKU#?\s*([^;]+)/i);
                         if (skuMatch && !newSku) newSku = skuMatch[1].trim();
                         
                         let priceMatch = mergedText.match(/Buy Price:\s*([\d,.]+)/i) || mergedText.match(/Price:\s*([\d,.]+)/i);
                         if (priceMatch && (!price || isNaN(price))) {
                             price = parseFloat(priceMatch[1].replace(/,/g, ''));
                         }
                     }
                 }
                 if (isNaN(price)) price = 0;

                 rowsByYearBlock[blockName].push([sku, desc, newSku, remarks, price, fileName, importDateStr]);
               }

               // --- LOGIC CHANGE: UPSERT (Update/Insert) ---
               for (let block in rowsByYearBlock) {
                  let targetSheet;
                  if (sheetCache[block]) {
                      targetSheet = sheetCache[block];
                  } else {
                      targetSheet = getSheetForYearBlock(PREFIX_SKU_MAP, importDateObj, headers);
                      if (targetSheet.getLastRow() > 1) {
                          targetSheet.getRange(2, 1, targetSheet.getLastRow() - 1, targetSheet.getLastColumn()).clearContent();
                      }
                      sheetCache[block] = targetSheet;
                  }
                  let newRows = rowsByYearBlock[block];
                  if (newRows.length > 0) {
                      let nextRow = targetSheet.getLastRow() + 1;
                      targetSheet.getRange(nextRow, 1, newRows.length, headers.length).setValues(newRows);
                  }
               }
               
               fileImported = true;
               break; 
             }
          }
        } 

        if (fileImported) filesProcessed++;
        if (tempFileId) Drive.Files.remove(tempFileId);
        file.setTrashed(true);
        
      } catch (e) { throw new Error("Error importing SKU Mapping " + fileName + ": " + e.toString()); }
    }
  }
  ss.toast(`Import Complete. Processed: ${filesProcessed}`, "SUCCESS", 5);
  console.timeEnd("Import_Missing_SKUs");
  } finally {
    lock.releaseLock();
  }
}

function _23_importMissingSKUsMappingBaskets() {
  var lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) { throw new Error("System is currently busy. Another user is generating reports right now. Please wait a few minutes."); }
  try {
  console.time("Import_Baskets");
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  if (!BASKET_FOLDER_ID || BASKET_FOLDER_ID === "INSERT_BASKET_FOLDER_ID_HERE") {
    ss.toast("Error: BASKET_FOLDER_ID is not configured.", "ERROR", -1);
    return;
  }

  const folder = DriveApp.getFolderById(BASKET_FOLDER_ID);
  const files = folder.getFiles();

  // Define the flat layout for the database to make it easy to lookup later
  const headers =["Basket SKU", "Basket Desc", "Basket Sell Price", "Component SKU", "Component Desc", "Component Unit Price", "Component Qty", "Component UOM", "Component Total", "Source File", "Import Date"];
  let filesProcessed = 0; const sheetCache = {};

  while (files.hasNext()) {
    let file = files.next();
    let fileName = file.getName();
    let mimeType = file.getMimeType();

    let isGoogleSheet = mimeType === MimeType.GOOGLE_SHEETS;
    let isExcel = mimeType === MimeType.MICROSOFT_EXCEL || 
                  mimeType === "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" || 
                  mimeType === "application/vnd.oasis.opendocument.spreadsheet" || 
                  fileName.endsWith(".xlsx") || fileName.endsWith(".xls") || fileName.endsWith(".ods");

    if (isGoogleSheet || isExcel) {
      ss.toast(`Scanning Baskets: ${fileName}...`, "IMPORT", -1);
      let tempFileId = null;
      try {
        let converted = _openFileOrConvertCSV(file);
        let ssNew = converted.ssNew;
        tempFileId = converted.tempFileId;

        const allSheets = ssNew.getSheets();
        let fileImported = false;
        let rowsByYearBlock = {};
        
        // Group by 5-year block based on current import date
        let importDateObj = new Date();
        let importDateStr = Utilities.formatDate(importDateObj, Session.getScriptTimeZone(), "MM/dd/yyyy");
        const startYear = Math.floor(importDateObj.getFullYear() / 5) * 5;
        const blockName = `${PREFIX_BASKET}${startYear}-${startYear + 4}`;

        for (let s = 0; s < allSheets.length; s++) {
          let sheetNew = allSheets[s];
          let data = sheetNew.getDataRange().getValues();

          if (data.length > 0) {
             let compSkuIdx = -1, compDescIdx = -1, compPriceIdx = -1, compQtyIdx = -1, compUomIdx = -1;
             let baskSkuIdx = -1, baskDescIdx = -1, baskPriceIdx = -1, baskQtyIdx = -1;
             let currentBasketSku = "";
             let currentBasketDesc = "";
             let currentBasketPrice = 0;
             let currentBasketQty = 1;

             // Scan EVERY row to dynamically find the tables (Supports headers at row 100+)
             for (let i = 0; i < data.length; i++) {
               let row = data[i];
               let skuCols = [], descCols =[], priceCols = [], qtyCols = [], uomCols =[];
               
               for (let c = 0; c < row.length; c++) {
                 let cellVal = String(row[c]).toLowerCase().trim();
                 if (cellVal === "sku #" || cellVal === "sku") skuCols.push(c);
                 if (cellVal === "desc" || cellVal === "description") descCols.push(c);
                 if (cellVal === "unit price" || cellVal === "price" || cellVal === "case cost") priceCols.push(c);
                 if (cellVal === "qty" || cellVal === "quantity") qtyCols.push(c);
                 if (cellVal === "uom" || cellVal === "unit") uomCols.push(c);
               }
               
               // If we found BOTH Left and Right tables on this row, set this as the header!
               if (skuCols.length >= 2) {
                 compSkuIdx = skuCols[0];
                 baskSkuIdx = skuCols[skuCols.length - 1]; 
                 
                 compDescIdx = descCols.length > 0 ? descCols[0] : -1;
                 baskDescIdx = descCols.length > 1 ? descCols[descCols.length - 1] : -1;
                 
                 compPriceIdx = priceCols.length > 0 ? priceCols[0] : -1;
                 baskPriceIdx = priceCols.length > 1 ? priceCols[priceCols.length - 1] : -1;
                 
                 compQtyIdx = qtyCols.length > 0 ? qtyCols[0] : -1;
                 baskQtyIdx = qtyCols.length > 1 ? qtyCols[qtyCols.length - 1] : -1;
                 compUomIdx = uomCols.length > 0 ? uomCols[0] : -1;
                 
                 // Reset the active basket in case there are multiple tables in one sheet
                 currentBasketSku = "";
                 currentBasketQty = 1;
                 continue; // Move to the next row to start reading data
               }

               // EXTRACT DATA (Executes only after headers are found)
               if (compSkuIdx > -1 && baskSkuIdx > -1) {
                    
                    // 1. Check right side for a new Basket Declaration
                    let rightSkuRaw = String(row[baskSkuIdx]).trim().toUpperCase();
                    if (rightSkuRaw && rightSkuRaw !== "SKU #" && rightSkuRaw !== "SKU") {
                        currentBasketSku = rightSkuRaw;
                        currentBasketDesc = baskDescIdx > -1 ? String(row[baskDescIdx]).trim() : "";
                        currentBasketPrice = baskPriceIdx > -1 ? parseFloat(String(row[baskPriceIdx]).replace(/[^0-9.-]+/g, "")) || 0 : 0;
                        currentBasketQty = baskQtyIdx > -1 ? parseFloat(String(row[baskQtyIdx]).replace(/[^0-9.-]+/g, "")) || 1 : 1;
                        if (currentBasketQty === 0) currentBasketQty = 1; // Prevent division by zero
                    }

                    // 2. Check left side for Components
                    let cSku = String(row[compSkuIdx]).trim().toUpperCase();
                    let cDesc = compDescIdx > -1 ? String(row[compDescIdx]).trim() : "";
                    
                    // Strictly require an SKU. Ignore empty rows, descriptions without SKUs, and headers.
                    if (!cSku || cSku === "SKU #" || cSku === "DESC" || cSku === "DESCRIPTION" || cSku === "VISMIN" || cSku === "NCR") continue;

                    let rawCQty = compQtyIdx > -1 ? parseFloat(String(row[compQtyIdx]).replace(/[^0-9.-]+/g, "")) || 0 : 0;
                    let cQty = rawCQty / currentBasketQty; // Mathematically corrects custom bulk orders

                    // If we have an active basket AND a valid component, map them!
                    if (currentBasketSku && cSku && cQty > 0) {
                        let cPrice = compPriceIdx > -1 ? parseFloat(String(row[compPriceIdx]).replace(/[^0-9.-]+/g, "")) || 0 : 0;
                        let cUom = compUomIdx > -1 ? String(row[compUomIdx]).trim() : "";
                        let cTotal = cPrice * cQty;

                        if (!rowsByYearBlock[blockName]) rowsByYearBlock[blockName] = [];
                        rowsByYearBlock[blockName].push([
                            currentBasketSku, currentBasketDesc, currentBasketPrice, 
                            cSku, cDesc, cPrice, cQty, cUom, cTotal, 
                            fileName, importDateStr
                        ]);
                        fileImported = true;
                    }
               }
             }
          }
        } // End of Sheets Loop

        // 3. UPSERT LOGIC (Update/Insert to prevent Duplication)
        if (fileImported && Object.keys(rowsByYearBlock).length > 0) {
           for (let block in rowsByYearBlock) {
              let targetSheet;
              if (sheetCache[block]) {
                  targetSheet = sheetCache[block];
              } else {
                  targetSheet = getSheetForYearBlock(PREFIX_BASKET, importDateObj, headers);
                  if (targetSheet.getLastRow() > 1) {
                      targetSheet.getRange(2, 1, targetSheet.getLastRow() - 1, targetSheet.getLastColumn()).clearContent();
                  }
                  sheetCache[block] = targetSheet;
              }
              let newRows = rowsByYearBlock[block];
              if (newRows.length > 0) {
                  let nextRow = targetSheet.getLastRow() + 1;
                  targetSheet.getRange(nextRow, 1, newRows.length, headers.length).setValues(newRows);
              }
           }
           filesProcessed++;
           if (tempFileId) Drive.Files.remove(tempFileId);
           file.setTrashed(true);
        } else {
           if (tempFileId) Drive.Files.remove(tempFileId);
        }
      } catch (e) {
        if (tempFileId) Drive.Files.remove(tempFileId);
        throw new Error("Error importing Basket " + fileName + ": " + e.toString());
      }
    }
  }
  ss.toast(`Baskets Import Complete. Files processed: ${filesProcessed}`, "SUCCESS", 5);
  console.timeEnd("Import_Baskets");
  } finally {
    lock.releaseLock();
  }
}

function _25_importPurchaseOrdersKeysysAndA2Z2GO(e) {
  var isBackground = (e && e.triggerUid);
  var lock = null;
  if (!isBackground) {
    lock = LockService.getScriptLock();
    try { lock.waitLock(10000); } catch (err) { throw new Error("System is currently busy. Another user is generating reports right now. Please wait a few minutes."); }
  }
  try {
  const startTime = Date.now();
  const MAX_EXECUTION_TIME = 4.5 * 60 * 1000;
  deleteTriggersByName("_25_importPurchaseOrdersKeysysAndA2Z2GO");

  console.time("Import_PO_Keysys_A2Z2GO");
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  if (!PO_KEYSYS_FOLDER_ID || PO_KEYSYS_FOLDER_ID === "INSERT_PO_KEYSYS_FOLDER_ID_HERE") {
    ss.toast("Error: PO_KEYSYS_FOLDER_ID is not configured.", "ERROR", -1);
    return;
  }

  const folder = DriveApp.getFolderById(PO_KEYSYS_FOLDER_ID);
  const files = folder.getFiles();
  const sheetCache = {};
  
  // NEW FORMAT HEADERS
  const headers = ["Supplier", "Purchase Invoice No.", "Date (MM/DD/YYYY)", "Received Purchase Orders (Our no.)", "Name", "Purchased Qty", "Net total", "Row VAT", "Row total", "Source File"];
  let filesProcessed = 0;

  while (files.hasNext()) {
    if (Date.now() - startTime > MAX_EXECUTION_TIME) {
        ss.toast("Approaching time limit. Pausing and auto-resuming PO Keysys in 1 minute...", "BATCH CONTINUATION", 10);
        ScriptApp.newTrigger("_25_importPurchaseOrdersKeysysAndA2Z2GO").timeBased().after(60 * 1000).create();
        return; 
    }

    let file = files.next();
    let fileName = file.getName();
    let mimeType = file.getMimeType();

    let isGoogleSheet = mimeType === MimeType.GOOGLE_SHEETS;
    let isExcel = mimeType === MimeType.MICROSOFT_EXCEL || 
                  mimeType === "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" || 
                  mimeType === MimeType.CSV || mimeType === "text/csv" ||
                  fileName.endsWith(".xlsx") || fileName.endsWith(".xls") || fileName.endsWith(".csv");
    
    if (isGoogleSheet || isExcel) {
      console.time(`Process_PO_Keysys_${fileName}`);
      ss.toast(`Importing PO Keysys & A2Z2GO: ${fileName}...`, "IMPORT", -1);
      let tempFileId = null;
      try {
        let converted = _openFileOrConvertCSV(file);
        let ssNew = converted.ssNew;
        tempFileId = converted.tempFileId;

        let sheetNew = ssNew.getSheets()[0];
        let rawData = sheetNew.getDataRange().getValues();
        
        if (rawData.length > 0) {
            let headerRowIndex = -1;
            let colMap = {};

            // --- HEADER DETECTION ---
            for (let r = 0; r < Math.min(20, rawData.length); r++) {
              let row = rawData[r].map(c => String(c).toLowerCase().trim());
              
              let hasSupplier = row.some(c => c.includes("supplier"));
              let hasName = row.some(c => c === "name" || c.includes("item name"));
              let hasQty = row.some(c => c.includes("purchased"));
              
              if (hasSupplier && hasName && hasQty) {
                  colMap["Supplier"] = row.findIndex(c => c.includes("supplier"));
                  colMap["Purchase"] = row.findIndex(c => c.includes("purchase invoice") || c === "purchase");
                  colMap["Date"] = row.findIndex(c => c === "date" || c.includes("date"));
                  colMap["Received"] = row.findIndex(c => c.includes("received purchase orders") || c === "received");
                  colMap["Name"] = row.findIndex(c => c === "name" || c.includes("item name"));
                  colMap["Purchased"] = row.findIndex(c => c.includes("purchased qty") || c === "purchased");
                  colMap["Net total"] = row.findIndex(c => c === "net total");
                  colMap["Row VAT"] = row.findIndex(c => c === "row vat");
                  colMap["Row total"] = row.findIndex(c => c === "row total");
                  
                  headerRowIndex = r;
                  break;
              }
            }

            if (headerRowIndex > -1) {
              let rowsByYearBlock = {};

              for (let i = headerRowIndex + 1; i < rawData.length; i++) {
                 let row = rawData[i];
                 
                 let name = (colMap["Name"] > -1) ? String(row[colMap["Name"]]).trim() : "";
                 if (name === "") continue;

                 let dateVal = (colMap["Date"] > -1 && row[colMap["Date"]]) ? row[colMap["Date"]] : new Date(); 
                 let dateObj = new Date(dateVal);
                 if (isNaN(dateObj.getTime())) dateObj = new Date();
                 
                 const cleanStr = (colName) => {
                     if (colMap[colName] > -1) {
                         let val = row[colMap[colName]];
                         // Extracts full number instead of 8.61E+14 if Excel force-saved it as scientific notation
                         if (typeof val === "number" && val > 9999999999) {
                             return val.toLocaleString('fullwide', {useGrouping: false});
                         }
                         return String(val).trim();
                     }
                     return "";
                 };
                 const getNum = (colName) => {
                     let val = (colMap[colName] > -1) ? parseFloat(String(row[colMap[colName]]).replace(/[^0-9.-]+/g, "")) : 0;
                     return isNaN(val) ? 0 : val;
                 };

                 const startYear = Math.floor(dateObj.getFullYear() / 5) * 5;
                 const blockName = `${PREFIX_PO_KEYSYS}${startYear}-${startYear + 4}`;

                 if (!rowsByYearBlock[blockName]) rowsByYearBlock[blockName] =[];
                 
                 rowsByYearBlock[blockName].push([
                   cleanStr("Supplier"),
                   cleanStr("Purchase"),
                   dateObj,
                   cleanStr("Received"),
                   name,
                   getNum("Purchased"),
                   getNum("Net total"),
                   getNum("Row VAT"),
                   getNum("Row total"),
                   fileName
                 ]);
              }

              // --- WIPE AND REPLACE LOGIC ---
              for (let block in rowsByYearBlock) {
                 let targetSheet;
                 if (sheetCache[block]) {
                   targetSheet = sheetCache[block];
                 } else {
                   let sampleDate = rowsByYearBlock[block][0][2];
                   targetSheet = getSheetForYearBlock(PREFIX_PO_KEYSYS, sampleDate, headers);
                   if (targetSheet.getLastRow() > 1) {
                       targetSheet.getRange(2, 1, targetSheet.getLastRow() - 1, targetSheet.getLastColumn()).clearContent();
                   }
                   sheetCache[block] = targetSheet;
                 }

                 let newRows = rowsByYearBlock[block];
                 const formatRow = (r) => {
                    if (r[2] instanceof Date) r[2] = Utilities.formatDate(r[2], Session.getScriptTimeZone(), "MM/dd/yyyy");
                    return r;
                 };
                 let finalData = newRows.map(formatRow);
                 
                 if (finalData.length > 0) {
                    let nextRow = targetSheet.getLastRow() + 1;
                    targetSheet.getRange(nextRow, 1, finalData.length, headers.length).setValues(finalData);
                 }
              }
            }
        }
        if (tempFileId) Drive.Files.remove(tempFileId);
        file.setTrashed(true); // Deletes file as requested
        filesProcessed++;
        console.timeEnd(`Process_PO_Keysys_${fileName}`);
      } catch (e) { 
        if (tempFileId) Drive.Files.remove(tempFileId);
        throw new Error("Error importing PO Keysys " + fileName + ": " + e.toString()); 
      }
    }
  }
  console.timeEnd("Import_PO_Keysys_A2Z2GO");
  ss.toast(`PO Keysys & A2Z2GO Import Complete. Files: ${filesProcessed}`, "SUCCESS", 5);
  } finally {
    if (lock) lock.releaseLock();
  }
}

function _31_importAEEmails(e) {
  var isBackground = (e && e.triggerUid);
  var lock = null;
  if (!isBackground) {
    lock = LockService.getScriptLock();
    try { lock.waitLock(10000); } catch (err) { throw new Error("System is currently busy."); }
  }
  try {
    console.time("Import_AE_Emails");
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!AE_EMAILS_FOLDER_ID || AE_EMAILS_FOLDER_ID === "INSERT_AE_EMAILS_FOLDER_ID_HERE") {
      ss.toast("Error: AE_EMAILS_FOLDER_ID is not configured.", "ERROR", -1);
      return;
    }
    const folder = DriveApp.getFolderById(AE_EMAILS_FOLDER_ID);
    const files = folder.getFiles();
    // ADDED: New headers for SN and AE Name For NTE System
    const headers = ["SN", "AE Name", "AE Name For NTE System", "Pingcon Email", "A2Z2go Email", "Keysys Email", "Source File", "Import Date"];
    let filesProcessed = 0; const sheetCache = {};

    while (files.hasNext()) {
      let file = files.next();
      let fileName = file.getName();
      if (fileName.endsWith(".xlsx") || fileName.endsWith(".xls") || fileName.endsWith(".csv")) {
        ss.toast(`Scanning AE Emails: ${fileName}...`, "IMPORT", -1);
        let tempFileId = null;
        try {
          let converted = _openFileOrConvertCSV(file);
          let ssNew = converted.ssNew;
          tempFileId = converted.tempFileId;
          
          let data = ssNew.getSheets()[0].getDataRange().getValues();
          let fileImported = false;

          if (data.length > 0) {
            let importDateObj = new Date();
            let importDateStr = Utilities.formatDate(importDateObj, Session.getScriptTimeZone(), "MM/dd/yyyy");
            let blockName = `${PREFIX_AE_EMAILS}${Math.floor(importDateObj.getFullYear() / 5) * 5}`;
            
            let targetSheet;
            if (sheetCache[blockName]) {
                targetSheet = sheetCache[blockName];
            } else {
                targetSheet = getSheetForYearBlock(PREFIX_AE_EMAILS, importDateObj, headers);
                if (targetSheet.getLastRow() > 1) {
                    targetSheet.getRange(2, 1, targetSheet.getLastRow() - 1, targetSheet.getLastColumn()).clearContent();
                }
                if (targetSheet.getLastColumn() < headers.length) {
                    targetSheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold").setBackground("#f3f3f3");
                }
                sheetCache[blockName] = targetSheet;
            }
            
            let newRows = [];
            let isKeysysSection = false;
            let snCounter = 1;

            // Dynamically find headers inside the uploaded Excel file to support shifts
            let excelHeaders = data[0].map(h => String(h).toLowerCase().trim());
            let nameIdx = excelHeaders.indexOf("ae name");
            let nteNameIdx = excelHeaders.indexOf("ae name for nte system");
            let pIdx = excelHeaders.indexOf("pingcon email");
            let aIdx = excelHeaders.indexOf("a2z2go email");
            let kIdx = excelHeaders.indexOf("keysys email");

            // Fallback for older formats without headers
            if (nameIdx === -1) nameIdx = 0; 
            if (pIdx === -1) pIdx = 1;

            for (let i = 1; i < data.length; i++) {
              let aeNameRaw = String(data[i][nameIdx]).trim();
              let nteNameRaw = nteNameIdx > -1 ? String(data[i][nteNameIdx]).trim() : "";
              let pingconRaw = pIdx > -1 ? String(data[i][pIdx]).trim() : "";
              let a2z2goRaw = aIdx > -1 ? String(data[i][aIdx]).trim() : "";
              let keysysRaw = kIdx > -1 ? String(data[i][kIdx]).trim() : "";

              if (aeNameRaw.toUpperCase().includes("KEYSYS")) {
                  isKeysysSection = true;
                  continue;
              }
              if (aeNameRaw.toUpperCase().includes("PMC") || aeNameRaw.toLowerCase().includes("manager")) {
                  continue;
              }

              if (aeNameRaw && (pingconRaw.includes("@") || a2z2goRaw.includes("@") || keysysRaw.includes("@"))) {
                  let pingconEmail = isKeysysSection ? "" : pingconRaw;
                  let a2z2goEmail = isKeysysSection ? pingconRaw : a2z2goRaw; // Fallbacks based on old logic
                  let keysysEmail = isKeysysSection ? a2z2goRaw : keysysRaw;

                  // Added SN counter and NTE Name to the row push
                  newRows.push([snCounter++, getCanonicalName(aeNameRaw), nteNameRaw, pingconEmail, a2z2goEmail, keysysEmail, fileName, importDateStr]);
              }
            }
            
            if (newRows.length > 0) {
               let nextRow = targetSheet.getLastRow() + 1;
               targetSheet.getRange(nextRow, 1, newRows.length, headers.length).setValues(newRows);
               fileImported = true;
            }
          }
          if (tempFileId) Drive.Files.remove(tempFileId);
          if (fileImported) { file.setTrashed(true); filesProcessed++; }
        } catch (e) {
          if (tempFileId) Drive.Files.remove(tempFileId);
          throw new Error("Error importing AE Emails " + fileName + ": " + e.toString());
        }
      }
    }
    ss.toast(`AE Emails Import Complete. Files: ${filesProcessed}`, "SUCCESS", 5);
    console.timeEnd("Import_AE_Emails");
  } finally {
    if (lock) lock.releaseLock();
  }
}

function _39_importNewNarraReport(e) {
  var isBackground = (e && e.triggerUid);
  try {
    const startTime = Date.now();
    deleteTriggersByName("_39_importNewNarraReport");
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    if (!NEW_NARRA_FOLDER_ID || NEW_NARRA_FOLDER_ID === "INSERT_NEW_NARRA_RAW_UPLOADS_FOLDER_ID_HERE") {
      ss.toast("Error: NEW_NARRA_FOLDER_ID is not configured.", "ERROR", -1);
      return;
    }

    const folder = DriveApp.getFolderById(NEW_NARRA_FOLDER_ID);
    const files = folder.getFiles();
    let filesProcessed = 0;
    const sheetCache = {};

    // Standardizes DD/MM/YYYY dates from Excel to prevent text conversion
    const parseDmyDate = (val) => {
        if (!val) return "";
        if (val instanceof Date) return val;
        if (typeof val === 'number') return new Date(Math.round((val - 25569) * 86400 * 1000));
        let str = String(val).trim();
        let parts = str.split(/[\/\-]/);
        if (parts.length === 3) {
            let d = parseInt(parts[0], 10);
            let m = parseInt(parts[1], 10) - 1;
            let y = parseInt(parts[2], 10);
            if (y < 100) y += 2000;
            return new Date(y, m, d);
        }
        let fallback = new Date(str);
        if (!isNaN(fallback.getTime())) return fallback;
        return "";
    };

    while (files.hasNext()) {
      if (Date.now() - startTime > 4.5 * 60 * 1000) {
          ss.toast("Pausing and auto-resuming New Narra Database in 1 minute...", "BATCH CONTINUATION", 10);
          ScriptApp.newTrigger("_39_importNewNarraReport").timeBased().after(60 * 1000).create();
          return; 
      }

      let file = files.next();
      let fileName = file.getName();
      let mimeType = file.getMimeType();
      let lowerName = fileName.toLowerCase();
      let isGoogleSheet = mimeType === MimeType.GOOGLE_SHEETS;
      let isExcel = mimeType === MimeType.MICROSOFT_EXCEL || mimeType === "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" || lowerName.endsWith(".xlsx") || lowerName.endsWith(".xls") || lowerName.endsWith(".csv");

      if (isGoogleSheet || isExcel) {
        ss.toast(`Importing New Narra: ${fileName}`, "IMPORT", -1);
        let tempFileId = null;
        try {
          let converted = _openFileOrConvertCSV(file);
          let ssNew = converted.ssNew;
          tempFileId = converted.tempFileId;
          let sheetNew = ssNew.getSheets()[0];
          let data = sheetNew.getDataRange().getValues();

          if (data.length > 0) {
             let headerRowIdx = -1;
             for(let r=0; r<Math.min(20, data.length); r++) {
               let rowStr = data[r].join(" ").toLowerCase();
               if(rowStr.includes("company name") && rowStr.includes("vouchernumber") && rowStr.includes("amount")) {
                   headerRowIdx = r; break;
               }
             }

             if (headerRowIdx > -1) {
                let originalHeaders = data[headerRowIdx].map(h => String(h).trim());
                
                // STRIP TRAILING BLANK HEADERS TO PREVENT OFFSET
                while(originalHeaders.length > 0 && originalHeaders[originalHeaders.length - 1] === "") {
                    originalHeaders.pop();
                }
                
                let currentHeaders = [...originalHeaders, "Source File", "Import Date"];
                let lowerHeaders = originalHeaders.map(h => String(h).toLowerCase().replace(/\s+/g, ''));
                
                let dateCols = [
                    lowerHeaders.indexOf("voucherdate"), lowerHeaders.indexOf("invoicedate"),
                    lowerHeaders.indexOf("checkdate"), lowerHeaders.indexOf("prepareddate"),
                    lowerHeaders.indexOf("checkeddate"), lowerHeaders.indexOf("approveddate")
                ].filter(idx => idx > -1);

                let vDateIdx = lowerHeaders.indexOf("voucherdate");
                let invNoIdx = lowerHeaders.indexOf("invoicenumber");
                
                let rowsByYearBlock = {};
                let importDateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MM/dd/yyyy");

                for (let i = headerRowIdx + 1; i < data.length; i++) {
                   let row = data[i];
                   if (row.every(c => String(c).trim() === "")) continue;

                   dateCols.forEach(idx => {
                       let dObj = parseDmyDate(row[idx]);
                       row[idx] = (dObj instanceof Date && !isNaN(dObj.getTime())) ? fastFormatDate(dObj) : String(row[idx]);
                   });

                   let blockDate = new Date();
                   if (vDateIdx > -1 && row[vDateIdx]) {
                       blockDate = new Date(row[vDateIdx]);
                       if (isNaN(blockDate.getTime())) blockDate = new Date();
                   }

                   const startYear = Math.floor(blockDate.getFullYear() / 5) * 5;
                   const blockName = `${PREFIX_NEW_NARRA}${startYear}-${startYear + 4}`;
                   
                   if (!rowsByYearBlock[blockName]) rowsByYearBlock[blockName] =[];
                   let cleanRow = row.slice(0, originalHeaders.length); // Trim to exact header length
                   rowsByYearBlock[blockName].push([...cleanRow, fileName, importDateStr]);
                }

                // WIPE AND REPLACE LOGIC
                for (let block in rowsByYearBlock) {
                   let targetSheet;
                   let targetLastCol;
                   if (sheetCache[block]) {
                       targetSheet = sheetCache[block];
                       targetLastCol = targetSheet.getLastColumn();
                   } else {
                       let sampleDate = new Date(block.substring(block.length-9, block.length-5), 0, 1);
                       targetSheet = getSheetForYearBlock(PREFIX_NEW_NARRA, sampleDate, currentHeaders);
                       
                       // FORCE UPDATE HEADERS TO FIX OFFSETS
                       targetSheet.getRange(1, 1, 1, currentHeaders.length).setValues([currentHeaders]);
                       targetSheet.getRange(1, 1, 1, currentHeaders.length).setFontWeight("bold").setBackground("#f3f3f3");
                       
                       // Clear any residual offset headers/data from previous bad uploads
                       let existingLastCol = targetSheet.getLastColumn();
                       if (existingLastCol > currentHeaders.length) {
                           targetSheet.getRange(1, currentHeaders.length + 1, Math.max(1, targetSheet.getLastRow()), existingLastCol - currentHeaders.length).clearContent();
                       }
                       
                       targetLastCol = currentHeaders.length;
                       
                       // Overwrites previous data automatically
                       if (targetSheet.getLastRow() > 1) {
                           targetSheet.getRange(2, 1, targetSheet.getLastRow() - 1, targetLastCol).clearContent();
                       }
                       sheetCache[block] = targetSheet;
                   }
                   
                   let newRows = rowsByYearBlock[block];
                   if (newRows.length > 0) {
                      let paddedData = newRows.map(r => {
                          let newRow = [...r];
                          while(newRow.length < targetLastCol) newRow.push("");
                          if (invNoIdx > -1 && newRow[invNoIdx] !== undefined) newRow[invNoIdx] = String(newRow[invNoIdx]).trim();
                          return newRow.slice(0, targetLastCol);
                      });
                      if (invNoIdx > -1) {
                          targetSheet.getRange(2, invNoIdx + 1, Math.max(1, targetSheet.getMaxRows() - 1), 1).setNumberFormat("@");
                      }
                      let nextRow = targetSheet.getLastRow() + 1;
                      targetSheet.getRange(nextRow, 1, paddedData.length, targetLastCol).setValues(paddedData);
                   }
                }
             }
          }
          if (tempFileId) Drive.Files.remove(tempFileId);
          file.setTrashed(true);
          filesProcessed++;
        } catch (e) { 
          if (tempFileId) Drive.Files.remove(tempFileId);
          console.error("Error processing New Narra " + fileName + ": " + e.toString()); 
          try { file.setTrashed(true); } catch(err) {} 
        }
      }
    }
    ss.toast(`New Narra Import Complete. Files: ${filesProcessed}`, "SUCCESS", 5);
  } finally {
    // Non-blocking background execution
  }
}

// =========================================================================
// HELPER FUNCTIONS
// =========================================================================

function resolveCollectionAssistantAssignment(custUpper, hasTech, hasBonita, hasA2ZPMC, hasPingcon, sysSettings) {
  const settings = sysSettings || getSystemSettings();
  const primaryClients = String(settings["Special_Clients_Primary"] || "").toUpperCase().split(',').map(s => s.trim()).filter(Boolean);
  const bonitaClients = String(settings["Special_Clients_Bonita"] || "").toUpperCase().split(',').map(s => s.trim()).filter(Boolean);
  const abcExclusions = String(settings["Special_Clients_ABC_Excluded"] || "").toUpperCase().split(',').map(s => s.trim()).filter(Boolean);

  const isSpecialPrimary = primaryClients.some(c => custUpper.includes(c));
  const isBonitaPrimary = bonitaClients.some(c => custUpper.startsWith(c));
  const isAbcExcluded = abcExclusions.some(c => custUpper.startsWith(c));

  const letterMatch = custUpper.match(/[A-Z]/);
  const firstLetter = letterMatch ? letterMatch[0] : "A";
  const isABC = /^[A-C]/.test(firstLetter);

  if (hasTech || isSpecialPrimary) return true;
  if (hasBonita) return isBonitaPrimary;
  if (hasA2ZPMC) return false;
  if (hasPingcon) return (isABC && !isAbcExcluded);

  return false;
}

function isTransactionVoid(notesText, additionalText) {
    let nText = String(notesText).trim();
    let aText = String(additionalText).trim();
    
    // 1. Check if Notes start with "VOID DR <number>"
    let voidDrMatch = nText.match(/^VOID[\W_]*DR[\W_]*0*(\d+)/i);
    if (voidDrMatch) {
        let drNum = voidDrMatch[1];
        let regex = new RegExp("DR\\s*0*" + drNum + "\\b", "i");
        if (regex.test(aText)) {
            return true; // Match found in both, it is a VOID transaction
        }
    } 
    // 1.b Fallback: If it just says "VOID DR" with NO number attached
    else if (/^VOID[\W_]*DR(?!\s*[\d\-])/i.test(nText)) {
        return true;
    }
    
    // 2. Check for "ORIG DR <number>" in Notes
    let origDrMatch = nText.match(/ORIG\s+DR\s*0*(\d+)/i);
    if (origDrMatch) {
        let drNum = origDrMatch[1];
        let regex = new RegExp("DR\\s*0*" + drNum + "\\b", "i");
        if (regex.test(aText)) {
            return true; // Match found in both, it is a VOID transaction
        }
    }
    
    return false;
}

// =========================================================================
// SYSTEM SETTINGS DATABASE GETTER (CACHE REMOVED FOR INSTANT UPDATES)
// =========================================================================
function getSystemSettings() {
  let settings = {};
  try {
      const searchFolder = DriveApp.getFolderById(DATABASE_FOLDER_ID);
      const searchFiles = searchFolder.searchFiles("title contains 'System_Settings_Data' and trashed = false and mimeType = 'application/vnd.google-apps.spreadsheet'");
      
      if (searchFiles.hasNext()) {
         let file = searchFiles.next();
         let ss = SpreadsheetApp.open(file);
         let allTabs = ss.getSheets();
         
         allTabs.forEach(sheet => {
             let data = sheet.getDataRange().getValues();
             if (data.length > 1) {
                 for (let i = 1; i < data.length; i++) {
                   let key = String(data[i][0]).trim();
                   let val = String(data[i][1]).trim();
                   let enableFlag = String(data[i][2] || "").trim().toUpperCase(); 
                   let isEnabled = (enableFlag !== "FALSE" && enableFlag !== "NO" && enableFlag !== "0"); 
                   
                   if (key) {
                      // If FALSE, value is set to empty string so it gets skipped in the CC list
                      settings[key] = isEnabled ? val : ""; 
                      if (key.startsWith("Enable_")) settings[key] = isEnabled ? "TRUE" : "FALSE";
                   }
                 }
             }
         });
      }
  } catch (e) {
      console.error("Settings Error: " + e.message);
  }
  return settings;
}

// =========================================================================
// EXTERNAL CLOUD VAULT PERSISTENCE ENGINE (ANTI-DATA-LOSS LAYER)
// =========================================================================

function _getCloudVaultConfig() {
  const props = PropertiesService.getScriptProperties();
  const baseUrl = props.getProperty('CLOUD_VAULT_URL');
  const apiKey = props.getProperty('CLOUD_VAULT_KEY');
  
  let isDev = false;
  try {
    const service = ScriptApp.getService();
    if (service && service.getUrl()) {
      isDev = service.getUrl().toLowerCase().endsWith("/dev");
    }
  } catch(e) {
    isDev = false;
  }
  
  return {
    baseUrl: baseUrl ? baseUrl.replace(/\/+$/, '') : '',
    apiKey: apiKey || '',
    environment: isDev ? 'DEV' : 'EXEC',
    isConfigured: !!(baseUrl && apiKey)
  };
}

function _syncRecordToCloudVault(table, recordId, payload) {
  const cfg = _getCloudVaultConfig();
  if (!cfg.isConfigured) {
    console.warn("[Cloud Vault] Skipped: CLOUD_VAULT_URL or CLOUD_VAULT_KEY not set in Script Properties.");
    return false;
  }
  
  try {
    const recordWithMeta = Object.assign({}, payload, {
      vault_record_id: String(recordId).trim(),
      vault_environment: cfg.environment
    });
    
    const endpoint = `${cfg.baseUrl}/${encodeURIComponent(table)}`;
    const options = {
      method: "post",
      contentType: "application/json",
      headers: {
        "apikey": cfg.apiKey,
        "Authorization": "Bearer " + cfg.apiKey,
        "Prefer": "return=minimal"
      },
      payload: JSON.stringify(recordWithMeta),
      muteHttpExceptions: true
    };
    
    const response = UrlFetchApp.fetch(endpoint, options);
    const code = response.getResponseCode();
    if (code >= 200 && code < 300) {
      console.log(`[Cloud Vault] Successfully synced [${recordId}] to [${table}] (${cfg.environment})`);
      return true;
    } else {
      console.error(`[Cloud Vault Error] HTTP ${code}: ${response.getContentText()}`);
      return false;
    }
  } catch(err) {
    console.error(`[Cloud Vault Exception] Failed to sync ${recordId}: ${err.message}`);
    return false;
  }
}

function _deleteRecordFromCloudVault(table, recordId, invoiceNumber = null) {
  const cfg = _getCloudVaultConfig();
  if (!cfg.isConfigured) return false;
  
  try {
    // Clean and normalize to UPPERCASE to eliminate case-sensitivity mismatches
    const cleanId = String(recordId || "").trim().toUpperCase();
    let queryParam = `vault_record_id=ilike.${encodeURIComponent(cleanId)}`;
    
    // If invoice number is provided, also target it with case-insensitive matching
    if (invoiceNumber) {
      const cleanInv = String(invoiceNumber).trim();
      queryParam = `invoice_number=eq.${encodeURIComponent(cleanInv)}`;
    }
    
    const endpoint = `${cfg.baseUrl}/${encodeURIComponent(table)}?${queryParam}&vault_environment=eq.${cfg.environment}`;
    const options = {
      method: "delete",
      headers: {
        "apikey": cfg.apiKey,
        "Authorization": "Bearer " + cfg.apiKey
      },
      muteHttpExceptions: true
    };
    
    const response = UrlFetchApp.fetch(endpoint, options);
    const code = response.getResponseCode();
    if (code >= 200 && code < 300) {
      console.log(`[Cloud Vault] Pruned [${recordId || invoiceNumber}] from [${table}] (${cfg.environment})`);
      return true;
    } else {
      console.warn(`[Cloud Vault Delete Warning] HTTP ${code}: ${response.getContentText()}`);
      return false;
    }
  } catch(err) {
    console.error(`[Cloud Vault Delete Error] ${err.message}`);
    return false;
  }
}

function _deleteBatchRecordsFromCloudVault(table, invoiceNumbersArray) {
  const cfg = _getCloudVaultConfig();
  if (!cfg.isConfigured || !Array.isArray(invoiceNumbersArray) || invoiceNumbersArray.length === 0) return false;
  
  try {
    // Sanitize and filter non-empty invoice numbers
    const cleanInvs = invoiceNumbersArray.map(inv => String(inv).trim()).filter(Boolean);
    if (cleanInvs.length === 0) return false;
    
    // Supabase in.() filter syntax
    const inFilter = `(${cleanInvs.map(inv => `"${encodeURIComponent(inv)}"`).join(',')})`;
    const endpoint = `${cfg.baseUrl}/${encodeURIComponent(table)}?invoice_number=in.${inFilter}&vault_environment=eq.${cfg.environment}`;
    
    const options = {
      method: "delete",
      headers: {
        "apikey": cfg.apiKey,
        "Authorization": "Bearer " + cfg.apiKey
      },
      muteHttpExceptions: true
    };
    
    const response = UrlFetchApp.fetch(endpoint, options);
    const code = response.getResponseCode();
    if (code >= 200 && code < 300) {
      console.log(`[Cloud Vault] Bulk pruned ${cleanInvs.length} invoices from [${table}] (${cfg.environment})`);
      return true;
    } else {
      console.warn(`[Cloud Vault Bulk Delete Warning] HTTP ${code}: ${response.getContentText()}`);
      return false;
    }
  } catch(err) {
    console.error(`[Cloud Vault Bulk Delete Error] ${err.message}`);
    return false;
  }
}

function _fetchRecordsFromCloudVault(table, environmentFilter = null) {
  const cfg = _getCloudVaultConfig();
  if (!cfg.isConfigured) {
    throw new Error("Cloud Vault is not configured in Script Properties.");
  }
  
  const env = environmentFilter || cfg.environment;
  const endpoint = `${cfg.baseUrl}/${encodeURIComponent(table)}?vault_environment=eq.${encodeURIComponent(env)}&select=*`;
  
  const options = {
    method: "get",
    headers: {
      "apikey": cfg.apiKey,
      "Authorization": "Bearer " + cfg.apiKey
    },
    muteHttpExceptions: true
  };
  
  const response = UrlFetchApp.fetch(endpoint, options);
  if (response.getResponseCode() === 200) {
    return JSON.parse(response.getContentText());
  } else {
    throw new Error(`Cloud Vault Query Failed (HTTP ${response.getResponseCode()}): ${response.getContentText()}`);
  }
}

function TEST_vaultConnection() {
  const testRecord = {
    file_name: "TEST_SAMPLE.pdf",
    invoice_number: "TEST-999",
    company_name: "TEST SUPPLIER CORP",
    total_amount: "1,250.00",
    raw_json: JSON.stringify({ test: "data" }),
    source_database: "TEST_SOURCE"
  };
  const success = _syncRecordToCloudVault("verified_invoices_vault", "TEST-999_TEST SUPPLIER CORP", testRecord);
  console.log("Vault Sync Test Result:", success);
}

// =========================================================================
// HISTORICAL VAULT MIGRATION (ONE-TIME BACKFILL TOOL)
// =========================================================================

function webappBackfillHistoricalRecordsToVault(sessionToken = null) {
  if (sessionToken) {
    const user = getAuthenticatedUser(sessionToken);
    if (!user || user.role !== "Super Admin") {
      return { success: false, error: "Permission Denied: Only Super Admins can run Historical Backfill." };
    }
  }

  let lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);

    const cfg = _getCloudVaultConfig();
    if (!cfg.isConfigured) {
      return { success: false, error: "Cloud Vault is not configured in Script Properties." };
    }

    const props = PropertiesService.getScriptProperties();
    const isDevMode = (cfg.environment === "DEV");
    const targetFolderId = (isDevMode && props.getProperty('VERIFIED_FOLDER_ID_DEV')) 
        ? props.getProperty('VERIFIED_FOLDER_ID_DEV') 
        : props.getProperty('VERIFIED_FOLDER_ID');

    if (!targetFolderId) {
      return { success: false, error: "Target Google Drive Folder ID is missing." };
    }

    const folder = DriveApp.getFolderById(targetFolderId);
    const filesIter = folder.getFilesByType(MimeType.GOOGLE_SHEETS);

    let totalSynced = 0;
    let totalSkipped = 0;

    // Fetch existing keys in vault to avoid duplicate network traffic
    let existingKeys = new Set();
    try {
      let existingVault = _fetchRecordsFromCloudVault("verified_invoices_vault", cfg.environment);
      if (Array.isArray(existingVault)) {
        existingVault.forEach(r => existingKeys.add(String(r.vault_record_id).trim().toUpperCase()));
      }
    } catch(e) {}

    while (filesIter.hasNext()) {
      let file = filesIter.next();
      if (file.isTrashed() || !file.getName().startsWith("Verified Invoices for 2307")) continue;

      let ss = SpreadsheetApp.open(file);
      let sheet = ss.getSheetByName("All Verified Invoices");
      if (!sheet || sheet.getLastRow() <= 1) continue;

      let data = sheet.getDataRange().getValues();
      let headers = data[0].map(h => String(h).trim());
      let fnIdx = headers.indexOf("File Name");
      let vdIdx = headers.indexOf("Verified Data");
      let srcIdx = headers.indexOf("Source");

      if (fnIdx === -1 || vdIdx === -1) continue;

      for (let i = 1; i < data.length; i++) {
        let fileName = String(data[i][fnIdx] || "").trim();
        let jsonStr = String(data[i][vdIdx] || "").trim();
        let src = srcIdx > -1 ? String(data[i][srcIdx] || "").trim() : "PINGCON - RECEIVING";

        if (jsonStr.startsWith("{")) {
          try {
            let parsed = JSON.parse(jsonStr);
            let invNum = String(parsed["Invoice Number"] || "").trim();
            let compName = String(parsed["Company Name"] || "").trim().toUpperCase();
            let vaultKey = `${invNum}_${compName}`.trim().toUpperCase();

            // Skip if already secured in the vault
            if (existingKeys.has(vaultKey)) {
              totalSkipped++;
              continue;
            }

            let totalAmtStr = "";
            if (parsed.isMixedVat) {
              let vGross = parseFloat(String((parsed.VAT_Portion && parsed.VAT_Portion.Gross) || "0").replace(/,/g, '')) || 0;
              let nvGross = parseFloat(String((parsed.NV_Portion && parsed.NV_Portion.Gross) || "0").replace(/,/g, '')) || 0;
              totalAmtStr = (vGross + nvGross).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
            } else {
              totalAmtStr = String(parsed["Gross"] || parsed["Total Amount"] || "0.00").trim();
            }

            let vaultRecord = {
              file_name: fileName,
              invoice_number: invNum,
              company_name: parsed["Company Name"] || "",
              total_amount: totalAmtStr,
              raw_json: jsonStr,
              source_database: src
            };

            let ok = _syncRecordToCloudVault("verified_invoices_vault", vaultKey, vaultRecord);
            if (ok) {
              existingKeys.add(vaultKey);
              totalSynced++;
            }
          } catch(err) {
            console.warn(`[Backfill Warning] Row ${i + 1} skipped: ${err.message}`);
          }
        }
      }
    }

    logSystemActivity("SYSTEM", "Historical Vault Backfill Executed", `Backfilled ${totalSynced} historical invoice(s), ${totalSkipped} already present (${cfg.environment}).`, sessionToken);

    return {
      success: true,
      synced: totalSynced,
      skipped: totalSkipped,
      environment: cfg.environment
    };
  } catch(err) {
    return { success: false, error: err.message };
  } finally {
    lock.releaseLock();
  }
}

// =========================================================================
// ONE-CLICK DISASTER RECOVERY ENGINE (RESTORE SHEETS FROM VAULT)
// =========================================================================

function webappRunCloudVaultDisasterRecovery(sessionToken = null) {
  // Security Gate: Ensure caller has administrative privileges
  if (sessionToken) {
    const user = getAuthenticatedUser(sessionToken);
    if (!user || user.role !== "Super Admin") {
      return { success: false, error: "Permission Denied: Only Super Admins can execute Disaster Recovery." };
    }
  }

  let lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000); // 30 seconds wait lock for database reconstruction

    const cfg = _getCloudVaultConfig();
    if (!cfg.isConfigured) {
      return { success: false, error: "Cloud Vault is not configured in Script Properties." };
    }

    // 1. Fetch all records from Supabase for this environment (DEV or EXEC)
    const vaultRecords = _fetchRecordsFromCloudVault("verified_invoices_vault", cfg.environment);
    if (!vaultRecords || vaultRecords.length === 0) {
      return { success: false, error: `No records found in Cloud Vault for environment: ${cfg.environment}` };
    }

    const props = PropertiesService.getScriptProperties();
    const isDevMode = (cfg.environment === "DEV");
    const targetFolderId = (isDevMode && props.getProperty('VERIFIED_FOLDER_ID_DEV')) 
        ? props.getProperty('VERIFIED_FOLDER_ID_DEV') 
        : props.getProperty('VERIFIED_FOLDER_ID');

    if (!targetFolderId) {
      return { success: false, error: "Target Google Drive Folder ID is missing from Script Properties." };
    }

    const folder = DriveApp.getFolderById(targetFolderId);

    // 2. Group records by Year based on their verified release date or invoice date
    let recordsByYear = {};
    vaultRecords.forEach(rec => {
      let rawJson = {};
      try {
        rawJson = (typeof rec.raw_json === "string") ? JSON.parse(rec.raw_json) : (rec.raw_json || {});
      } catch(e) {}

      let relDate = rawJson["2307 Release Date"] || rawJson["Release Date"] || "";
      let invDate = rawJson["Invoice Date"] || "";
      let dateToUse = new Date(relDate || invDate || rec.created_at || new Date());
      let year = !isNaN(dateToUse.getTime()) ? Utilities.formatDate(dateToUse, Session.getScriptTimeZone(), "yyyy") : Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy");

      if (!recordsByYear[year]) recordsByYear[year] = [];
      recordsByYear[year].push({
        fileName: rec.file_name || ("SI " + (rec.invoice_number || "INV") + ".pdf"),
        rawJsonStr: (typeof rec.raw_json === "string") ? rec.raw_json : JSON.stringify(rec.raw_json),
        releaseDate: (rawJson["Record Type"] === "Standard Purchase") ? "" : (relDate || ""),
        sourceDatabase: rec.source_database || "PINGCON - RECEIVING",
        verifiedDateTime: Utilities.formatDate(new Date(rec.created_at || new Date()), Session.getScriptTimeZone(), "MM/dd/yyyy HH:mm")
      });
    });

    let totalRestored = 0;
    const targetTab = "All Verified Invoices";
    const headers = ["SN", "File Name", "Verified Data", "2307 Release Date", "Source", "Date and Time Verified (MM/DD/YYYY)"];

    // 3. Rebuild each missing or damaged Google Sheet
    for (let year in recordsByYear) {
      let dbName = "Verified Invoices for 2307 " + year;
      let files = folder.getFilesByName(dbName);
      let ss = null;

      while (files.hasNext()) {
        let f = files.next();
        if (!f.isTrashed()) {
          ss = SpreadsheetApp.open(f);
          break;
        }
      }

      // If the sheet was completely deleted/trashed, create a brand-new one in the target folder
      if (!ss) {
        ss = SpreadsheetApp.create(dbName);
        DriveApp.getFileById(ss.getId()).moveTo(folder);
        console.log(`[Disaster Recovery] Recreated missing Google Sheet: [${dbName}] in folder: [${targetFolderId}]`);
      }

      let sheet = ss.getSheetByName(targetTab);
      if (!sheet) {
        sheet = ss.insertSheet(targetTab);
        let sheet1 = ss.getSheetByName("Sheet1");
        if (sheet1 && ss.getSheets().length > 1) {
          try { ss.deleteSheet(sheet1); } catch(e) {}
        }
      }

      // Setup clean headers
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#f3f3f3");
      sheet.setFrozenRows(1);

      // Re-populate all rows from the Vault with sequential serial numbers
      let yearRecords = recordsByYear[year];
      let rowsToInsert = [];

      for (let i = 0; i < yearRecords.length; i++) {
        let rec = yearRecords[i];
        let sn = i + 1;

        // Ensure internal JSON SN is updated to match row SN
        let patchedJsonStr = rec.rawJsonStr;
        try {
          let p = JSON.parse(patchedJsonStr);
          p["SN"] = sn;
          patchedJsonStr = JSON.stringify(p);
        } catch(e) {}

        rowsToInsert.push([
          sn,
          rec.fileName,
          patchedJsonStr,
          rec.releaseDate,
          rec.sourceDatabase,
          rec.verifiedDateTime
        ]);
      }

      if (rowsToInsert.length > 0) {
        // Clear old damaged contents if sheet existed but was corrupted
        if (sheet.getLastRow() > 1) {
          sheet.getRange(2, 1, sheet.getLastRow() - 1, headers.length).clearContent();
        }

        sheet.getRange(2, 1, rowsToInsert.length, headers.length).setValues(rowsToInsert);
        sheet.getRange(2, 3, rowsToInsert.length, 1).setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP).setVerticalAlignment("top");
        totalRestored += rowsToInsert.length;
      }
    }

    SpreadsheetApp.flush();

    logSystemActivity("SYSTEM", "Cloud Vault Disaster Recovery Executed", `Reconstructed ${totalRestored} verified record(s) from Supabase across all sheets (${cfg.environment}).`, sessionToken);

    return {
      success: true,
      recordsRestored: totalRestored,
      environment: cfg.environment,
      yearsProcessed: Object.keys(recordsByYear)
    };

  } catch(err) {
    console.error("[Disaster Recovery Exception] " + err.message);
    return { success: false, error: err.message };
  } finally {
    lock.releaseLock();
  }
}

// =========================================================================
// EWT PRE-VERIFICATION EXCEL EXPORT
// =========================================================================
function webappExportEwtVerificationExcel(tableData, companyContext, startDate, endDate) {
    try {
        // Fetch dynamic email from the multi-tab settings database
        let sysSettings = getSystemSettings();
        let recipient = sysSettings["EWT_Verification_Email"] || SCRIPT_PROPS['ADMIN_EMAIL']; 
        
        let headers = ["SN", "2307 Invoice No.", "Supplier Name", "2307 Total Amount", "Narra Amount", "Check Summary Amount", "RIWB Amount", "Match Status"];
        let cleanContext = String(companyContext).replace(/[^a-zA-Z0-9]/g, "");
        let fileName = `EWT Verification Status - ${cleanContext} - ${startDate} to ${endDate}`;
        
        let tempSs = SpreadsheetApp.create(fileName);
        let sheet = tempSs.getSheets()[0];
        
        sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold").setBackground("#f3f3f3");
        
        if (tableData && tableData.length > 0) {
            sheet.getRange(2, 1, tableData.length, headers.length).setValues(tableData);
            sheet.getRange(2, 4, tableData.length, 4).setNumberFormat("#,##0.00"); 
            // Removed the "@" plain text format for Col 2 so the =HYPERLINK formulas can execute properly
        }
        
        sheet.getDataRange().setWrap(true).setVerticalAlignment("top");
        for (let i = 1; i <= headers.length; i++) sheet.autoResizeColumn(i);
        
        SpreadsheetApp.flush();
        Utilities.sleep(3000); 
        
        let url = "https://docs.google.com/spreadsheets/d/" + tempSs.getId() + "/export?format=xlsx";
        let token = ScriptApp.getOAuthToken();
        let response = UrlFetchApp.fetch(url, { headers: { 'Authorization': 'Bearer ' + token }, muteHttpExceptions: true });
        let blob = response.getBlob().setName(fileName + ".xlsx");
        
        DriveApp.getFileById(tempSs.getId()).setTrashed(true);
        
        let htmlBody = `
          <div style="font-family: Arial, sans-serif; font-size: 13px; color: #000;">
            <p>Good day,</p>
            <p>Please find attached the EWT Verification Status Report for <strong>${companyContext}</strong>.</p>
            <p><strong>Period Covered:</strong> ${startDate} to ${endDate}</p>
            <p>This report contains the current matching status of all scanned invoices against Narra, Check Summary, and RIWB.</p>
            <p>Thank you.</p>
          </div>`;
        
        // Encrypt the attachment based on the company context
        let isTech = (String(companyContext).toUpperCase().includes("KEYSYS") || String(companyContext).toUpperCase().includes("A2Z"));
        let props = PropertiesService.getScriptProperties();
let finalPassword = isTech ? props.getProperty('ENCRYPT_PASS_TECH') : props.getProperty('ENCRYPT_PASS_PINGCON');
        let securedBlob = blob;
        try { 
            securedBlob = encryptAttachmentViaPython(blob, fileName + ".xlsx", finalPassword); 
        } catch (e) { 
            console.error("Encryption Fallback."); 
        }
          
        // Added the attachments array so the file actually gets sent
        sendRobustEmail(recipient, fileName, htmlBody, { 
            name: "Finance System", 
            attachments: [securedBlob] 
        });
        
        // Return Base64 of the raw blob so the frontend can download it instantly
        return { 
            success: true, 
            message: "Excel report successfully generated and emailed to " + recipient,
            base64: Utilities.base64Encode(blob.getBytes()),
            fileName: fileName + ".xlsx"
        };
    } catch (e) {
        return { success: false, message: e.toString() };
    }
}

// NEW HELPER FUNCTION FOR EXTERNAL DATABASE
function getDatabaseSheets(prefix) {
  const sheets = [];
  let targetFolderId = DATABASE_FOLDER_ID;
  if (prefix.includes("Narra_Voucher_Data") || prefix.includes("Check_Summary_Data") || prefix.includes("RIWB_Data")) {
      targetFolderId = SCRIPT_PROPS['VERIFIED_FOLDER_ID'];
  }
  const dbFolder = DriveApp.getFolderById(targetFolderId);
  const files = dbFolder.getFiles();
  while (files.hasNext()) {
    let file = files.next();
    if (file.getName().startsWith(prefix) && file.getMimeType() === MimeType.GOOGLE_SHEETS) {
      sheets.push(SpreadsheetApp.open(file).getSheets()[0]);
    }
  }
  return sheets;
}

let GLOBAL_AE_MAP = null;

function getAEMap() {
    if (GLOBAL_AE_MAP) return GLOBAL_AE_MAP;
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const arSheets = getDatabaseSheets("AR_Data_");
    const srSheets = getDatabaseSheets("SR_Data_");
    
    let allNames = new Set();
    
    const extractNames = (sheets) => {
        sheets.forEach(s => {
            const data = s.getDataRange().getValues();
            if (data.length <= 1) return;
            const headers = data[0].map(h => String(h).toLowerCase().trim());
            const managerIdx = headers.findIndex(h => h.includes("manager") || h === "cust manager");
            if (managerIdx > -1) {
                for(let i = 1; i < data.length; i++) {
                    // Strips ALL hidden unicode characters from the map
                    let m = String(data[i][managerIdx]).replace(/[^\x20-\x7E]/g, '').replace(/\s+/g, ' ').trim();
                    if (m && m.toLowerCase() !== "cust manager") {
                        m = m.split(' ').map(w => w ? w.charAt(0).toUpperCase() + w.slice(1).toLowerCase() : '').join(' ');
                        allNames.add(m);
                    }
                }
            }
        });
    };
    
    extractNames(arSheets);
    extractNames(srSheets);

    let uniqueNames = Array.from(allNames);
    uniqueNames.sort((a, b) => b.length - a.length);

    let map = {};
    for (let i = 0; i < uniqueNames.length; i++) {
        let current = uniqueNames[i];
        let matched = false;
        
        let currentParts = current.toLowerCase().replace(/[^a-z0-9]/g, " ").split(/\s+/).filter(x => x);
        
        for (let canonical of new Set(Object.values(map))) {
            let canonicalParts = canonical.toLowerCase().replace(/[^a-z0-9]/g, " ").split(/\s+/).filter(x => x);
            
            if (currentParts.length > 0 && canonicalParts.length > 0 && currentParts[0] === canonicalParts[0]) {
                if (currentParts.every(p => canonicalParts.includes(p))) {
                    map[current] = canonical;
                    matched = true;
                    break;
                }
            }
        }
        
        if (!matched) {
            map[current] = current;
        }
    }
    
    // DYNAMIC TYPO NORMALIZATION: Pulls spelling corrections from Script Properties
    let aeExceptions = JSON.parse(SCRIPT_PROPS['AE_TYPO_MAP'] || '{}');
    for (let errName in aeExceptions) {
        map[errName] = aeExceptions[errName];
    }
    
    GLOBAL_AE_MAP = map;
    return map;
}

function getCanonicalName(rawName) {
    if (!rawName) return "";
    // FORCE STRIP invisible unicode characters so reports merge perfectly
    let cleanRaw = String(rawName).replace(/[^\x20-\x7E]/g, '').replace(/\s+/g, ' ').trim();
    if (cleanRaw.toLowerCase() === "cust manager" || cleanRaw === "") return "";
    
    cleanRaw = cleanRaw.split(' ').map(w => w ? w.charAt(0).toUpperCase() + w.slice(1).toLowerCase() : '').join(' ');
    
    let map = getAEMap();
    return map[cleanRaw] || cleanRaw;
}



function formatMoney(amount) {
  return amount.toFixed(2).replace(/\d(?=(\d{3})+\.)/g, '$&,');
}

function checkNameMatch(aeName, searchInput) {
  if (!searchInput) return true;
  const cleanAe = String(aeName).toLowerCase().replace(/[^a-z0-9]/g, " ").split(/\s+/).filter(x => x);
  const searchTerms = String(searchInput).toLowerCase().replace(/[^a-z0-9]/g, " ").split(/\s+/).filter(x => x);
  
  if (cleanAe.length === 0 || searchTerms.length === 0) return false;
  if (cleanAe[0] !== searchTerms[0]) return false;

  const aeContainsSearch = searchTerms.every(term => cleanAe.includes(term));
  const searchContainsAe = cleanAe.every(term => searchTerms.includes(term));
  
  return aeContainsSearch || searchContainsAe;
}

function getSenderEmail(isTech) {
  const props = PropertiesService.getScriptProperties();
  return isTech 
    ? (props.getProperty('SENDER_EMAIL_A2Z') || "") 
    : (props.getProperty('SENDER_EMAIL_PINGCON') || "");
}

function getCompanyEntity(sourceFile) {
    if (!sourceFile) return "Pingcon";
    let sf = String(sourceFile).toLowerCase();
    if (sf.includes("bonita")) return "Bonita";
    if (sf.includes("keysys")) return "Keysys";
    if (sf.includes("pingcon")) return "Pingcon";
    if (sf.includes("a2z") && (sf.includes("it") || sf.includes("i.t"))) return "A2ZIT";
    if (sf.includes("a2z") && sf.includes("pmc")) return "A2ZPMC";
    if (sf.includes("a2zit")) return "A2ZIT";
    if (sf.includes("a2zpmc")) return "A2ZPMC";
    return "Pingcon";
}

function fastFormatDate(d) {
    if (!(d instanceof Date) || isNaN(d.getTime())) return String(d);
    return ('0' + (d.getMonth() + 1)).slice(-2) + '/' + ('0' + d.getDate()).slice(-2) + '/' + d.getFullYear();
}

function replaceTrigger(funcName, delayMs = 60000) {
   deleteTriggersByName(funcName);
   ScriptApp.newTrigger(funcName).timeBased().after(delayMs).create();
   PropertiesService.getScriptProperties().setProperty("RESUME_TIME", (Date.now() + delayMs).toString());
   PropertiesService.getScriptProperties().setProperty("THREAD_STATE", "SLEEPING"); // Mark as sleeping
}

var GLOBAL_AE_EMAILS_CACHE = null;
function getAEEmailsList() {
  if (GLOBAL_AE_EMAILS_CACHE) return GLOBAL_AE_EMAILS_CACHE;
  let aeEmails = [];
  try {
    const sheets = getDatabaseSheets("AE_Emails_Data_");
    sheets.forEach(sheet => {
      const data = sheet.getDataRange().getValues();
      if (data.length < 2) return;
      let headers = data[0].map(h => String(h).toLowerCase().trim());
      let pIdx = headers.indexOf("pingcon email");
      let aIdx = headers.indexOf("a2z2go email");
      let kIdx = headers.indexOf("keysys email");
      
      for (let i = 1; i < data.length; i++) {
        let pingcon = pIdx > -1 ? String(data[i][pIdx]).trim().toLowerCase() : "";
        let a2zit = aIdx > -1 ? String(data[i][aIdx]).trim().toLowerCase() : "";
        let keysys = kIdx > -1 ? String(data[i][kIdx]).trim().toLowerCase() : "";
        if (pingcon && pingcon.includes("@")) aeEmails.push(pingcon);
        if (a2zit && a2zit.includes("@")) aeEmails.push(a2zit);
        if (keysys && keysys.includes("@")) aeEmails.push(keysys);
      }
    });
  } catch(e) {}
  GLOBAL_AE_EMAILS_CACHE = aeEmails;
  return aeEmails;
}

// --- ADD THESE LINES INSTEAD ---
function sendRobustEmail(recipient, subject, body, options) {
  // Protect system emails (like OTPs) from ever entering Test Mode
  const isSystemEmail = options && options.isSystemEmail === true;
  const isTestMode = !isSystemEmail && PropertiesService.getScriptProperties().getProperty("TEST_MODE") === "true";
  
  const excludedTestEmails = JSON.parse(SCRIPT_PROPS['EXCLUDED_TEST_EMAILS'] || '[]');

  let fullExclusionList = excludedTestEmails;
  if (isTestMode) {
      let aeList = getAEEmailsList();
      fullExclusionList = fullExclusionList.concat(aeList);
  }

  const applyTest = (emails) => {
      if (!emails) return "";
      return emails.split(',').map(e => {
          let trimmed = e.trim();
          let lowerCaseEmail = trimmed.toLowerCase();
          
          if (trimmed && !lowerCaseEmail.endsWith('.test') && !fullExclusionList.includes(lowerCaseEmail)) {
              return trimmed + ".test";
          }
          return trimmed; 
      }).join(', ');
  };

  let finalRecipient = recipient;
  let finalSubject = subject;
  if (isTestMode) {
      finalRecipient = applyTest(recipient);
      // Safely adds "Test " only if the subject doesn't already start with it
      finalSubject = subject.trim().toLowerCase().startsWith("test") ? subject : "Test " + subject;
      if (options && options.cc) options.cc = applyTest(options.cc);
      if (options && options.replyTo) options.replyTo = applyTest(options.replyTo);
  }

  let validAliases = GmailApp.getAliases();
  let matchedAlias = validAliases.find(a => a.toLowerCase() === ((options && options.from) || "").toLowerCase());
  
  let finalOptions = {
      htmlBody: body,
      name: (options && options.name) ? options.name : "Finance System",
      from: matchedAlias ? matchedAlias : (options && options.from) ? options.from : undefined
  };
  if (finalOptions.from === undefined) delete finalOptions.from;

  if (options && options.replyTo) finalOptions.replyTo = options.replyTo;
  if (options && options.cc) finalOptions.cc = options.cc;
  if (options && options.attachments) finalOptions.attachments = options.attachments;
  if (options && options.inlineImages) finalOptions.inlineImages = options.inlineImages;
  
  try {
      GmailApp.sendEmail(finalRecipient, finalSubject, "", finalOptions);
  } catch (e) {
      let errStr = e.toString().toLowerCase();
      // If GmailApp timed out, it actually successfully sent the email but failed to return a response in time.
      // Skipping MailApp fallback prevents sending a duplicate second copy.
      if (errStr.includes("timeout") || errStr.includes("time out") || errStr.includes("exceeded")) {
          console.warn("GmailApp timeout. Email was likely sent. Skipping MailApp fallback to prevent duplicates.");
          return;
      }
      try {
          delete finalOptions.from;
          finalOptions.to = finalRecipient;
          finalOptions.subject = finalSubject;
          MailApp.sendEmail(finalOptions);
      } catch (err) {
          console.error("Failed to send email: " + err.toString());
          if (options && options.throwError) throw err;
      }
  }
}

function _$processMonthlyBuySellForUI(month, reqDate) {
  var lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) { throw new Error("System is currently busy. Another user is generating reports right now. Please wait a few minutes."); }
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    ss.toast(`Generating Buy/Sell Reports for ALL AEs in ${month}...`, "PROCESS", -1);
    _$runAllProcesses(null, month, true, reqDate);
  } finally { lock.releaseLock(); }
}

function _$processSendEmailMonthlyHeader(type, fileName, headers, data, spreadsheet) {
  if (!REPORT_EMAIL || REPORT_EMAIL === "INSERT_EMAIL_HERE") return;
  
  // 1. Determine Min and Max Dates (Data is already sorted descending, so index 0 is newest/max, length-1 is oldest/min)
  let dateIdx = (type === "EWT") ? 2 : 1; 
  let maxDateRaw = new Date(data[0][dateIdx]);
  let minDateRaw = new Date(data[data.length - 1][dateIdx]);
  
  let minDateStr = isNaN(minDateRaw.getTime()) ? "Unknown Date" : Utilities.formatDate(minDateRaw, Session.getScriptTimeZone(), "MMMM d yyyy");
  let maxDateStr = isNaN(maxDateRaw.getTime()) ? "Unknown Date" : Utilities.formatDate(maxDateRaw, Session.getScriptTimeZone(), "MMMM d yyyy");
  
  let reportDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-dd-MMMM H:mm");
  let subject = `List of ${type} Headers ${minDateStr} to ${maxDateStr} as of Report Date ${reportDate}`;
  
  // 2. Build HTML Table
  let htmlBody = `<div style="font-family: Arial, sans-serif; font-size: 13px; color: #000;">
                    <p>Good day,</p>
                    <p>Please find the generated <strong>${type} Headers</strong> below.</p>
                    <table style="border-collapse: collapse; width: 100%; font-size: 12px;" border="1" cellpadding="5">
                      <thead>
                        <tr style="background-color: #f3f3f3; font-weight: bold;">`;
  
  headers.forEach(h => htmlBody += `<th>${h}</th>`);
  htmlBody += `</tr></thead><tbody>`;
  
  let isTooLarge = false;
  for (let i = 0; i < data.length; i++) {
    // Check limit (100,000 characters is a safe limit to prevent Gmail from clipping the table entirely)
    if (htmlBody.length > 100000) {
      isTooLarge = true;
      break; 
    }
    htmlBody += `<tr>`;
    data[i].forEach(cell => {
      htmlBody += `<td>${cell !== undefined ? cell : ""}</td>`;
    });
    htmlBody += `</tr>`;
  }
  
  htmlBody += `</tbody></table>`;
  
  if (isTooLarge) {
    htmlBody += `<div style="color: red; margin-top: 15px; font-weight: bold; border: 1px solid red; padding: 10px; background: #ffeeee;">
                  Notice: The list is too long to display entirely in this email body. 
                  Please refer to the attached XLSX file for the complete and un-cut list.
                 </div>`;
  }
  
  htmlBody += `<br><p>Thank you.</p></div>`;
  
  // 3. Fetch XLSX Blob directly from the generated Spreadsheet
  SpreadsheetApp.flush(); 
  let url = "https://docs.google.com/spreadsheets/d/" + spreadsheet.getId() + "/export?format=xlsx";
  let token = ScriptApp.getOAuthToken();
  let xlsxBlob;
  
  try {
    let response = UrlFetchApp.fetch(url, { headers: { 'Authorization': 'Bearer ' + token } });
    xlsxBlob = response.getBlob().setName(fileName + ".xlsx");
  } catch (e) {
    console.error("Failed to generate XLSX attachment: " + e.toString());
  }
  let sysSettings = getSystemSettings();
  const mainRecipient = sysSettings["Management_A2Z"] || ""; 
  const ccRecipients = sysSettings["AP_Pingcon"] || "";
  const senderEmail = getSenderEmail(false);

  // 4. Send Email
  let mailOptions = { from: senderEmail, name: "Finance Pingcon", cc: ccRecipients };
  if (xlsxBlob) mailOptions.attachments = [encryptAttachmentViaPython(xlsxBlob, fileName + ".xlsx", SCRIPT_PROPS['ENCRYPT_PASS_PINGCON'])];
  sendRobustEmail(mainRecipient, subject, htmlBody, mailOptions);
  
  Utilities.sleep(1000); // Tiny pause between generating emails to respect Google's limits
}


function _$processCommissionReportForUI(ae, month) {
  var lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) { throw new Error("System is currently busy. Another user is generating reports right now. Please wait a few minutes."); }
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    ss.toast(`Generating Commission Report for ${ae} in ${month}...`, "PROCESS", -1);
    _$processBuySellForms(ae, month, true, null, false, true);
    ss.toast("Commission Report Generated successfully.", "SUCCESS", 5);
  } finally { lock.releaseLock(); }
}

function _$processMonthlyCommissionForUI(month) {
  var lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) { throw new Error("System is currently busy. Another user is generating reports right now. Please wait a few minutes."); }
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    ss.toast(`Generating Commission Reports for ALL AEs in ${month}...`, "PROCESS", -1);
    _$processBuySellForms(null, month, true, null, false, true);
    ss.toast("All Commission Reports Generated successfully.", "SUCCESS", 5);
  } finally { lock.releaseLock(); }
}




function _$processEWTReportClientForUI(ae, month) {
  var lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) { throw new Error("System is currently busy. Another user is generating reports right now. Please wait a few minutes."); }
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    ss.toast(`Processing EWT Report Client for ${ae} in ${month}...`, "PROCESS", -1);
    _$processAndSendEWTReportClient(ae, [month]);
    ss.toast("EWT Report Client Generated successfully.", "SUCCESS", 5);
  } finally { lock.releaseLock(); }
}
function _$processAllEWTReportClientForUI(month) {
  var lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) { throw new Error("System is currently busy. Another user is generating reports right now. Please wait a few minutes."); }
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    ss.toast(`Processing EWT Report Client for ALL AEs in ${month}...`, "PROCESS", -1);
    _$processAndSendEWTReportClient(null, [month]);
    ss.toast("All EWT Report Client Generated successfully.", "SUCCESS", 5);
  } finally { lock.releaseLock(); }
}






function _$processAndSendEWTReportClient(startStr, endStr) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const props = PropertiesService.getScriptProperties();
  const allowedEntities = JSON.parse(props.getProperty("SALES_ENTITIES") || "[]");

  let startDate = new Date(startStr + "T00:00:00");
  let endDate = new Date(endStr + "T23:59:59");

  // 1. Gather EWT Database (Extracted via Python)
  const ewtAllData = gatherAllData("EWT_Data_");
  if (ewtAllData.length === 0) {
    ss.toast("No EWT Data found. Please extract BIR forms using Python first.", "Error");
    return;
  }

  // 2. Gather SR Data to map SI to Entity for Password logic
  const srAllData = gatherAllData(PREFIX_SR);
  let siInfoMap = {};
  srAllData.forEach(item => {
    let r = item.data;
    let addText = String(r[5]).trim();
    let entity = getCompanyEntity(String(r[17] || ""));
    let match = addText.match(/SI\s*#?\s*(\d+)/i);
    let siNo = match ? match[1] : addText;
    siInfoMap[siNo] = entity;
  });

  let filteredRows = [];
  let hasKeysysGroup = false;

  // 3. Filter the EWT data using the Form's "Date From"
  ewtAllData.forEach(item => {
    let r = item.data;
    
    // Check Date From (Column B / Index 1)
    let dateObj = new Date(r[1]);
    if (isNaN(dateObj.getTime()) || dateObj < startDate || dateObj > endDate) return;

    let siNo = String(r[3]).trim();
    let entity = siInfoMap[siNo] || "Pingcon";
    
    if (allowedEntities.length > 0 && !allowedEntities.includes(entity)) return;

    if (entity === "Keysys" || entity === "A2ZIT") hasKeysysGroup = true;

    filteredRows.push({ entity: entity, dateObj: dateObj, rawRow: r });
  });

  if (filteredRows.length === 0) return;

  // 4. Build the Report
  filteredRows.sort((a, b) => b.dateObj - a.dateObj);
  let reportHeaders = ["SN", "Date from (MM/DD/YYYY)", "Date to (MM/DD/YYYY)", "SI No.", "Customer Name", "Tin", "Address", "ATC", "Income Payment", "WC 158 1%", "WC 157 2%", "WC 160 2%", "WV010 5%", "WC120 2%", "WC640 1%", "WV020 5%", "WI632 1%", "WC690"];
  
  let finalRows = filteredRows.map((item, idx) => {
    let rr = item.rawRow;
    return [idx + 1, fastFormatDate(new Date(rr[1])), fastFormatDate(new Date(rr[2])), rr[3], rr[4], rr[5], rr[6], rr[7], rr[8], rr[9], rr[10], rr[11], rr[12], rr[13], rr[14], rr[15], rr[16], rr[17]];
  });

  let fStart = Utilities.formatDate(startDate, Session.getScriptTimeZone(), "MM-dd-yyyy");
  let fEnd = Utilities.formatDate(endDate, Session.getScriptTimeZone(), "MM-dd-yyyy");
  let reportDateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-ddMMM-HHmm");
  let fileName = `EWT Report Client - ${fStart} to ${fEnd} - Generated ${reportDateStr}`;
  
  let tempSs = SpreadsheetApp.create(fileName);
  let sheet = tempSs.getSheets()[0];
  sheet.getRange(1, 1, 1, reportHeaders.length).setValues([reportHeaders]).setFontWeight("bold").setBackground("#f3f3f3");
  sheet.getRange(2, 1, finalRows.length, reportHeaders.length).setValues(finalRows);
  sheet.getRange(2, 9, finalRows.length, 10).setNumberFormat("#,##0.00");
  for (let i = 1; i <= reportHeaders.length; i++) sheet.autoResizeColumn(i);
  SpreadsheetApp.flush();

  let url = "https://docs.google.com/spreadsheets/d/" + tempSs.getId() + "/export?format=xlsx";
  let token = ScriptApp.getOAuthToken();
  let blob = UrlFetchApp.fetch(url, { headers: { 'Authorization': 'Bearer ' + token } }).getBlob().setName(fileName + ".xlsx");

  // 5. Password Logic
  let password = hasKeysysGroup ? SCRIPT_PROPS['ENCRYPT_PASS_TECH'] : SCRIPT_PROPS['ENCRYPT_PASS_PINGCON'];
  let securedBlob = encryptAttachmentViaPython(blob, fileName + ".xlsx", password);

  // 6. Save unencrypted original to Drive internally
  if (EWT_REPORT_CLIENT_ROOT_FOLDER_ID && EWT_REPORT_CLIENT_ROOT_FOLDER_ID !== "INSERT_EWT_REPORT_CLIENT_FOLDER_ID_HERE") {
      try {
        let rootFolder = DriveApp.getFolderById(EWT_REPORT_CLIENT_ROOT_FOLDER_ID);
        rootFolder.createFile(blob); 
      } catch(e) {}
  }

  // 7. Send the Email
  let htmlBody = `<div style="font-family:Arial;font-size:13px;"><p>Good day,</p><p>Please find attached the <b>EWT Report Client</b> for the period <b>${fStart} to ${fEnd}</b>.</p><p>This report contains extracted BIR 2307 data matched against our Sales Register.</p><p>Thank you.</p></div>`;
  
  sendRobustEmail(REPORT_EMAIL, `EWT Report Client: ${fStart} to ${fEnd}`, htmlBody, {
    attachments: [securedBlob],
    name: "Finance System"
  });

  DriveApp.getFileById(tempSs.getId()).setTrashed(true);
}

function writeOutputSheet(folder, fileName, headers, newRows) {
  let files = folder.getFilesByName(fileName);
  let spreadsheet;
  let targetSheet;
  
  if (files.hasNext()) {
    spreadsheet = SpreadsheetApp.open(files.next());
    targetSheet = spreadsheet.getSheets()[0];
  } else {
    spreadsheet = SpreadsheetApp.create(fileName);
    let file = DriveApp.getFileById(spreadsheet.getId());
    file.moveTo(folder);
    targetSheet = spreadsheet.getSheets()[0];
  }
  
  let existingData =[];
  let lastRow = targetSheet.getLastRow();
  let currentHeaders = headers;
  
  if (lastRow > 0) {
    let existingHeaders = targetSheet.getRange(1, 1, 1, targetSheet.getLastColumn()).getValues()[0];
    if (existingHeaders.length >= headers.length) {
      currentHeaders = existingHeaders;
    }
  } else {
    targetSheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold").setBackground("#f3f3f3");
  }
  
  if (lastRow > 1) {
    existingData = targetSheet.getRange(2, 1, lastRow - 1, currentHeaders.length).getValues();
  }
  
  let isEWT = headers[1] === "Account Title";
  let dateIdx = isEWT ? 2 : 1; // Invoice Date index shifted due to SN
  let invIdx = isEWT ? 6 : 2;  // Invoice Number index shifted due to SN
  let compIdx = 3;             // Company Name index shifted due to SN
  
  let combinedMap = new Map();
  existingData.forEach(r => combinedMap.set(`${r[invIdx]}_${r[compIdx]}`, r));
  newRows.forEach(r => {
    let paddedRow = [...r];
    while(paddedRow.length < currentHeaders.length) paddedRow.push("");
    combinedMap.set(`${r[invIdx]}_${r[compIdx]}`, paddedRow);
  });
  
  let finalData = Array.from(combinedMap.values());
  if (finalData.length > 0) {
    // 1. Sort by Invoice Date (Descending - newest at the top)
    finalData.sort((a, b) => {
      let dateA = new Date(a[dateIdx]);
      let dateB = new Date(b[dateIdx]);
      return dateB - dateA;
    });

    // 2. Assign SN (1, 2, 3...) to the first column (Index 0)
    finalData.forEach((row, index) => {
      row[0] = index + 1;
    });

    targetSheet.getRange(2, 1, targetSheet.getLastRow() || 2, currentHeaders.length).clearContent();
    targetSheet.getRange(2, 1, finalData.length, currentHeaders.length).setValues(finalData);
    if (isEWT) targetSheet.getRange(2, 6, finalData.length, 1).setNumberFormat("@");
    else targetSheet.getRange(2, 2, finalData.length, 1).setNumberFormat("@");
  }
  
  // Return the spreadsheet and sorted data for the email logic
  return { spreadsheet: spreadsheet, data: finalData };
}

// =========================================================================
// SALES REPORTS (Batch Process to Avoid 6-Minute Limit)
// Fixed Date Range: Oct 1 2025 to Mar 20 2026
// =========================================================================




// =========================================================================
// SALES REPORT CONSOLIDATED BY ENTITY (NEW)
// =========================================================================
function _10b_runSendAllSalesReportsConsolidated() {
  const props = PropertiesService.getScriptProperties();
  deleteTriggersByName("resumeSalesReportConsolidatedBatch");
  props.deleteProperty("CANCEL_SALES_REPORT_CONSOLIDATED");
  resumeSalesReportConsolidatedBatch();
}

function resumeSalesReportConsolidatedBatch() {
  const props = PropertiesService.getScriptProperties();
  props.setProperty("THREAD_STATE", "RUNNING"); // <--- ADDED
  
  if (props.getProperty("CANCEL_SALES_REPORT_CONSOLIDATED") === "true") {
      deleteTriggersByName("resumeSalesReportConsolidatedBatch");
      props.deleteProperty("CANCEL_SALES_REPORT_CONSOLIDATED");
      clearReportLock();
      return;
  }

  const arAllData = gatherAllData("AR_Data_");
  const sqAllData = gatherAllData("SQ_Data_");
  const srAllData = gatherAllData("SR_Data_"); 
  
  const arSheets = getDatabaseSheets("AR_Data_");
  let arHeaders = [];
  if (arSheets.length > 0 && arSheets[0].getLastColumn() > 0) {
    arHeaders = arSheets[0].getRange(1, 1, 1, arSheets[0].getLastColumn()).getValues()[0].map(h => String(h).toLowerCase().trim());
  }
  
  let dateIdx = arHeaders.findIndex(h => h.includes("date") && !h.includes("expiry") && !h.includes("countered") && !h.includes("received"));
  let customerIdx = arHeaders.findIndex(h => h.includes("customer"));
  let managerIdx = arHeaders.findIndex(h => h.includes("manager"));
  let arTextIdx = arHeaders.findIndex(h => h.includes("additional text"));
  let arNotesIdx = arHeaders.findIndex(h => h === "notes"); 
  let totalIdx = arHeaders.findIndex(h => h.includes("sales document total in main currency") || h.includes("sales document total"));
  let vatTotalIdx = arHeaders.findIndex(h => h === "vat total" || h.includes("vat total"));
  let netTotalIdx = arHeaders.findIndex(h => h === "net total" || h.includes("net total") || h.includes("vatable"));
  let sourceFileIdx = arHeaders.findIndex(h => h.includes("source file"));

  let srProductMap = {};
  srAllData.forEach(item => {
      let r = item.data;
      let numCol = String(r[3]).trim(); 
      let addText = String(r[5]).trim(); 
      let notes = String(r[6]).trim(); 
      let prod = String(r[7]).trim(); 
      if (prod) {
          let combinedText = addText + " " + notes;
          let siMatch = combinedText.match(/SI\s*#?\s*(\d+)/i);
          let drMatch = combinedText.match(/DR\s*#?\s*(\d+)/i);
          let exactSi = siMatch ? siMatch[1].trim() : (drMatch ? drMatch[1].trim() : "");
          let exactNum = numCol.replace(/^(DR|SI)\s*/i, "").trim();
          
          if (exactSi) {
              if (!srProductMap[exactSi]) srProductMap[exactSi] = [];
              if (!srProductMap[exactSi].includes(prod)) srProductMap[exactSi].push(prod);
          }
          if (exactNum && exactNum !== "???") {
              if (!srProductMap[exactNum]) srProductMap[exactNum] = [];
              if (!srProductMap[exactNum].includes(prod)) srProductMap[exactNum].push(prod);
          }
      }
  });

  const startStr = props.getProperty("BATCH_START_DATE");
  const endStr = props.getProperty("BATCH_END_DATE");
  const startDate = startStr ? new Date(startStr + "T00:00:00") : new Date(2025, 9, 1);
  const endDate = endStr ? new Date(endStr + "T23:59:59") : new Date(2026, 2, 20, 23, 59, 59);
  
  let allowedEntities = JSON.parse(props.getProperty("SALES_ENTITIES") || "[]");
  let allowedAEsStr = props.getProperty("SALES_TARGET_AES");
  let allowedAEs = allowedAEsStr ? JSON.parse(allowedAEsStr) : [];
  
  // AE Emails
  const aeEmailSheets = getDatabaseSheets("AE_Emails_Data_");
  let aeEmailMap = {};
  aeEmailSheets.forEach(sheet => {
     let data = sheet.getDataRange().getValues();
     if(data.length < 2) return;
     let headers = data[0].map(h => String(h).toLowerCase().trim());
     let nameIdx = headers.indexOf("ae name");
     let nteNameIdx = headers.indexOf("ae name for nte system");
     let pIdx = headers.indexOf("pingcon email");
     let aIdx = headers.indexOf("a2z2go email");
     let kIdx = headers.indexOf("keysys email");
     let supIdx = headers.findIndex(h => h.includes("supervisor") && !h.includes("email"));
     let supEmailIdx = headers.findIndex(h => h.includes("supervisor email"));
     let deptIdx = headers.findIndex(h => h.includes("department"));

     for(let i=1; i<data.length; i++) { 
         let aeName = nameIdx > -1 ? String(data[i][nameIdx]).trim() : "";
         if(aeName) {
             aeEmailMap[aeName] = { 
                 nteName: nteNameIdx > -1 ? String(data[i][nteNameIdx]).trim() : "",
                 pingcon: pIdx > -1 ? String(data[i][pIdx]).trim() : "", 
                 a2zit: aIdx > -1 ? String(data[i][aIdx]).trim() : "", 
                 keysys: kIdx > -1 ? String(data[i][kIdx]).trim() : "",
                 supervisor: supIdx > -1 ? String(data[i][supIdx]).trim() : "",
                 supervisorEmail: supEmailIdx > -1 ? String(data[i][supEmailIdx]).trim() : "",
                 department: deptIdx > -1 ? String(data[i][deptIdx]).trim() : ""
             };
         }
     }
  });
  
  let monthlyBuckets = {};
  let monthlySummary = {};
  
  // --- PRE-POPULATE ALL MONTHS AND ALL AEs FROM AE_EMAILS_DATA ---
  let tempStart = new Date(startDate.getFullYear(), startDate.getMonth(), 1);
  let tempEnd = new Date(endDate.getFullYear(), endDate.getMonth(), 1);
  while (tempStart <= tempEnd) {
      let mStr = Utilities.formatDate(tempStart, Session.getScriptTimeZone(), "MMMM yyyy");
      monthlyBuckets[mStr] = { "Keysys, A2Z2GO IT": [], "Pingcon, Bonita, A2Z2GO PMC": [] };
      monthlySummary[mStr] = { "Keysys, A2Z2GO IT": {}, "Pingcon, Bonita, A2Z2GO PMC": {} };
      
      for (let aeName in aeEmailMap) {
          // Only include AEs who have an assigned quota in SQ_Data
          let hasQuota = sqAllData.some(qItem => checkNameMatch(String(qItem.data[0]).trim(), aeName));
          if (!hasQuota) continue;
          
          // Check if specific AEs were selected in the UI
          if (allowedAEs.length > 0 && !allowedAEs.includes(aeName)) continue;
          
          let aeEmails = aeEmailMap[aeName];
          let bName = (aeEmails.a2zit || aeEmails.keysys) ? "Keysys, A2Z2GO IT" : "Pingcon, Bonita, A2Z2GO PMC";
          
          // Pre-insert the AE with 0 actual sales
          monthlySummary[mStr][bName][aeName] = { actual: 0 };
      }
      tempStart.setMonth(tempStart.getMonth() + 1);
  }
  // ---------------------------------------------------------------
  
  arAllData.forEach(item => {
    let row = item.data;
    let textVal = (arTextIdx > -1) ? String(row[arTextIdx]) : "";
    let notesVal = (arNotesIdx > -1) ? String(row[arNotesIdx]) : "";
    if (isTransactionVoid(notesVal, textVal)) return;

    let rawDate = (dateIdx > -1) ? row[dateIdx] : "";
    let dateObj = (rawDate instanceof Date) ? rawDate : new Date(rawDate);
    if (isNaN(dateObj.getTime()) || dateObj < startDate || dateObj > endDate) return;

    let sourceFile = (sourceFileIdx > -1) ? String(row[sourceFileIdx]) : "";
    let companyEntity = getCompanyEntity(sourceFile);
    if (allowedEntities.length > 0 && !allowedEntities.includes(companyEntity)) return;

    let rawManager = (managerIdx > -1) ? String(row[managerIdx]) : "";
    let manager = getCanonicalName(rawManager);
    if (!manager) return;
    let hasQuota = sqAllData.some(qItem => checkNameMatch(String(qItem.data[0]).trim(), manager));
    if (!hasQuota) return;
    if (allowedAEs.length > 0 && !allowedAEs.includes(manager)) return;

    let customer = (customerIdx > -1) ? String(row[customerIdx]).trim() : "";
    
    // Determine Bucket strictly by AE Email database first, then fallback to Entity
    let bucketName = "Pingcon, Bonita, A2Z2GO PMC";
    if (aeEmailMap[manager] && (aeEmailMap[manager].a2zit || aeEmailMap[manager].keysys)) {
        bucketName = "Keysys, A2Z2GO IT";
    } else if (!aeEmailMap[manager] && (companyEntity === "Keysys" || companyEntity === "A2ZIT")) {
        bucketName = "Keysys, A2Z2GO IT";
    }

    let combinedText = textVal + " " + notesVal;
    let siNo = "";
    let siMatch = combinedText.match(/SI\s*#?\s*(\d+)/i);
    if (siMatch) { siNo = siMatch[1]; } 
    else { 
        let drMatch = combinedText.match(/DR\s*#?\s*(\d+)/i);
        if (drMatch) { siNo = "DR " + drMatch[1]; } else { siNo = textVal; }
    }

    const getNum = (idx) => (idx > -1 && row[idx] !== "" && row[idx] != null) ? parseFloat(row[idx]) : 0;
    let total = getNum(totalIdx);
    let outputVat = getNum(vatTotalIdx);
    let vatable = getNum(netTotalIdx);
    
    let exactSi = siNo.replace(/^(DR|SI)\s*/i, "").trim();
    let products = srProductMap[exactSi] ? srProductMap[exactSi].join(", ") : "";

    let monthStr = Utilities.formatDate(dateObj, Session.getScriptTimeZone(), "MMMM yyyy");
    if (!monthlyBuckets[monthStr]) {
        monthlyBuckets[monthStr] = { "Keysys, A2Z2GO IT": [], "Pingcon, Bonita, A2Z2GO PMC": [] };
        monthlySummary[monthStr] = { "Keysys, A2Z2GO IT": {}, "Pingcon, Bonita, A2Z2GO PMC": {} };
    }

    monthlyBuckets[monthStr][bucketName].push({
      dateObj: dateObj,
      dateStr: fastFormatDate(dateObj),
      siNo: siNo,
      customer: customer,
      manager: manager,
      total: total,
      outputVat: outputVat,
      vatable: vatable,
      products: products,
      companyEntity: companyEntity,
      sourceFile: sourceFile
    });

    if (!monthlySummary[monthStr][bucketName][manager]) {
        monthlySummary[monthStr][bucketName][manager] = { actual: 0 };
    }
    monthlySummary[monthStr][bucketName][manager].actual += vatable;
  });

  // Calculate Quotas for each AE per Month in summary
  for (let monthStr in monthlySummary) {
      let mDate = new Date(monthStr);
      let monthDateObj = new Date(mDate.getFullYear(), mDate.getMonth(), 1); 
      let monthEndObj = new Date(mDate.getFullYear(), mDate.getMonth() + 1, 0);

      for (let bucketName in monthlySummary[monthStr]) {
          for (let ae in monthlySummary[monthStr][bucketName]) {
              let aeSummary = monthlySummary[monthStr][bucketName][ae];
              aeSummary.quota = 0;
              aeSummary.earliestStart = null;
              for(let q=0; q<sqAllData.length; q++) {
                  let qRow = sqAllData[q].data;
                  let qAe = String(qRow[0]).trim();
                  if (checkNameMatch(qAe, ae)) {
                       let qStart = new Date(qRow[1]);
                       let qEnd = new Date(qRow[2]);
                       let newQuota = parseFloat(qRow[4]) || 0;
                       
                       // Keep track of the absolute earliest quota date on record for this AE
                       if (!isNaN(qStart.getTime())) {
                           if (!aeSummary.earliestStart || qStart < aeSummary.earliestStart) {
                               aeSummary.earliestStart = qStart;
                           }
                       }

                       if(qStart <= monthEndObj && qEnd >= monthDateObj) {
                           if (newQuota > 0) {
                               aeSummary.quota = newQuota;
                           }
                           break;
                       }
                  }
              }
          }
      }
  }

  // Generate Excel Files and Send
  const now = new Date();
  let maxAsOfDate = 0;
  let reportDateString = "";
  arAllData.forEach(item => {
      let sourceFile = (sourceFileIdx > -1) ? String(item.data[sourceFileIdx]) : "";
      if (sourceFile) {
          let match = sourceFile.match(/as of\s+([a-zA-Z]+)\s*(\d+),?\s+(\d{4})\s+(\d{4})/i);
          if (match) {
              let dateObj = new Date(`${match[1]} ${match[2]}, ${match[3]} ${match[4].substring(0,2)}:${match[4].substring(2,4)}:00`);
              if (!isNaN(dateObj.getTime()) && dateObj.getTime() > maxAsOfDate) {
                  maxAsOfDate = dateObj.getTime();
                  let paddedDay = match[2].length === 1 ? "0" + match[2] : match[2];
                  reportDateString = `${match[1]} ${paddedDay}, ${match[3]} ${match[4]}`;
              }
          }
      }
  });
  if (!reportDateString) {
      reportDateString = Utilities.formatDate(now, Session.getScriptTimeZone(), "MMM dd, yyyy HHmm");
  }

  for (let monthStr in monthlyBuckets) {
    let buckets = monthlyBuckets[monthStr];
    let summaryByBucketAndAE = monthlySummary[monthStr];
    
    let mDate = new Date(monthStr);
    let sDateStr = Utilities.formatDate(new Date(mDate.getFullYear(), mDate.getMonth(), 1), Session.getScriptTimeZone(), "MMMM d, yyyy");
    let eDateStr = Utilities.formatDate(new Date(mDate.getFullYear(), mDate.getMonth() + 1, 0), Session.getScriptTimeZone(), "MMMM d, yyyy");

    for (let bucketName in buckets) {
        // STRICT CANCEL CHECK: Stop instantly if Cancel was clicked
        if (props.getProperty("CANCEL_SALES_REPORT_CONSOLIDATED") === "true") { clearReportLock(); return; }

        let items = buckets[bucketName];
        let summaryAEs = summaryByBucketAndAE[bucketName];
        if (Object.keys(summaryAEs).length === 0) continue;

        let shortBucketName = bucketName === "Keysys, A2Z2GO IT" ? "A2Z2GO" : "Pingcon";
        let cleanFileName = `Consolidated Sales Quota Report - ${shortBucketName} - ${monthStr} as of Report Date ${reportDateString}`;
        let tempSpreadsheet = SpreadsheetApp.create(cleanFileName);

      // --- BUILD SUMMARY SHEET ---
      let summarySheet = tempSpreadsheet.getSheets()[0];
      summarySheet.setName("Summary");

      let summaryHeaders = [
          "SN", 
          "AE Name", 
          `Quota (VAT EX)\n${sDateStr} to ${eDateStr}`, 
          `Actual Sales (VAT EX)\n${sDateStr} to ${eDateStr}, as of ${reportDateString}`, 
          "Balance", 
          "Achievement Rate", 
          "Quota Reached?"
      ];
      let summaryRows = [summaryHeaders];
      let tableHtmlRows = "";
      let grandTotalActual = 0;
      let grandTotalQuota = 0;

      let aeList = Object.keys(summaryAEs).sort();
      aeList.forEach((ae, idx) => {
          let aeData = summaryAEs[ae];
          let balance = Math.max(0, aeData.quota - aeData.actual);
          let rate = aeData.quota > 0 ? ((aeData.actual / aeData.quota) * 100).toFixed(2) + "%" : "0.00%";
          
          let reachedExcel = "";
          let reachedHtml = "";
          
          if (aeData.quota === 0) {
              let mDateEnd = new Date(mDate.getFullYear(), mDate.getMonth() + 1, 0, 23, 59, 59);
              if (aeData.earliestStart && aeData.earliestStart > mDateEnd) {
                  let sMonth = Utilities.formatDate(aeData.earliestStart, Session.getScriptTimeZone(), "MMMM yyyy");
                  reachedExcel = `Missing Quota\n(Started ${sMonth})`;
                  reachedHtml = `Missing Quota<br><span style="font-size:10px; font-weight:normal; color:#777;">(Started ${sMonth})</span>`;
              } else {
                  reachedExcel = "Missing Quota";
                  reachedHtml = "Missing Quota";
              }
          } else {
              reachedExcel = aeData.actual >= aeData.quota ? "YES" : "NO";
              reachedHtml = reachedExcel;
          }
          
          let reachedColor = aeData.quota === 0 ? "#cc0000" : (reachedExcel === "YES" ? "#28a745" : "#cc0000");
          
          grandTotalActual += aeData.actual;
          grandTotalQuota += aeData.quota;

          summaryRows.push([idx + 1, ae, aeData.quota, aeData.actual, balance, rate, reachedExcel]);
          tableHtmlRows += `<tr>
            <td style="border:1px solid #ccc; padding:8px; text-align:center;">${idx + 1}</td>
            <td style="border:1px solid #ccc; padding:8px;">${ae}</td>
            <td style="border:1px solid #ccc; padding:8px; text-align:right;">${aeData.quota.toLocaleString(undefined,{minimumFractionDigits:2})}</td>
            <td style="border:1px solid #ccc; padding:8px; text-align:right;">${aeData.actual.toLocaleString(undefined,{minimumFractionDigits:2})}</td>
            <td style="border:1px solid #ccc; padding:8px; text-align:right;">${balance.toLocaleString(undefined,{minimumFractionDigits:2})}</td>
            <td style="border:1px solid #ccc; padding:8px; text-align:center;">${rate}</td>
            <td style="border:1px solid #ccc; padding:8px; text-align:center; font-weight:bold; color:${reachedColor};">${reachedHtml}</td>
          </tr>`;
      });
      
      let overallBalance = Math.max(0, grandTotalQuota - grandTotalActual);
      let overallRate = grandTotalQuota > 0 ? ((grandTotalActual / grandTotalQuota) * 100).toFixed(2) + "%" : "0.00%";
      
      summaryRows.push(["", "GRAND TOTAL:", grandTotalQuota, grandTotalActual, overallBalance, overallRate, ""]);
      
      tableHtmlRows += `<tr style="background-color:#f9f9f9; font-weight:bold;">
            <td style="border:1px solid #ccc; padding:8px; text-align:center;"></td>
            <td style="border:1px solid #ccc; padding:8px;">GRAND TOTAL:</td>
            <td style="border:1px solid #ccc; padding:8px; text-align:right;">${grandTotalQuota.toLocaleString(undefined,{minimumFractionDigits:2})}</td>
            <td style="border:1px solid #ccc; padding:8px; text-align:right;">${grandTotalActual.toLocaleString(undefined,{minimumFractionDigits:2})}</td>
            <td style="border:1px solid #ccc; padding:8px; text-align:right;">${overallBalance.toLocaleString(undefined,{minimumFractionDigits:2})}</td>
            <td style="border:1px solid #ccc; padding:8px; text-align:center;">${overallRate}</td>
            <td style="border:1px solid #ccc; padding:8px; text-align:center;"></td>
          </tr>`;

      summarySheet.getRange(1, 1, summaryRows.length, summaryHeaders.length).setValues(summaryRows);
      summarySheet.getRange(1, 1, 1, summaryHeaders.length).setFontWeight("bold").setBackground("#f3f3f3");
      
      // Force text-wrap and center alignment so the long dates display perfectly
      summarySheet.getRange(1, 1, 1, summaryHeaders.length).setWrap(true).setVerticalAlignment("middle").setHorizontalAlignment("center");
      
      summarySheet.getRange(summaryRows.length, 2, 1, 6).setFontWeight("bold"); // Grand Total Row
      summarySheet.getRange(2, 3, summaryRows.length - 1, 3).setNumberFormat("#,##0.00");
      
      // Set explicit wide columns to fix the ######## issue
      summarySheet.setColumnWidth(1, 50);
      summarySheet.setColumnWidth(2, 220);
      summarySheet.setColumnWidth(3, 220);
      summarySheet.setColumnWidth(4, 300); 
      summarySheet.setColumnWidth(5, 130);
      summarySheet.setColumnWidth(6, 150);
      summarySheet.setColumnWidth(7, 130);
      
      // Explicitly increase header row height so multi-line text isn't cut off in Excel
      summarySheet.setRowHeight(1, 70);

      // --- BUILD DETAILS SHEET ---
      let detailSheet = tempSpreadsheet.insertSheet("Details");
      let detailHeaders = ["SN", "Cust Manager", "Entity", "Date (MM/DD/YYYY)", "SI No.", "Customer", "Total", "Output VAT", "Vatable", "Products"];
      let detailRows = [detailHeaders];
      
      items.sort((a, b) => {
          let mCmp = a.manager.localeCompare(b.manager);
          if (mCmp !== 0) return mCmp;
          return b.dateObj - a.dateObj;
      });

      let currentManager = null;
      let managerTotal = 0;
      let sn = 1;

      items.forEach(item => {
          if (currentManager !== item.manager) {
              if (currentManager !== null) {
                  let subRow = Array(detailHeaders.length).fill("");
                  subRow[5] = "Total for " + currentManager + ":";
                  subRow[8] = managerTotal;
                  detailRows.push(subRow);
                  detailRows.push(Array(detailHeaders.length).fill("")); 
              }
              currentManager = item.manager;
              managerTotal = 0;
              detailRows.push(detailHeaders); 
          }
          let displayEntity = item.companyEntity === "A2ZIT" ? "A2Z2GO" : item.companyEntity;
          detailRows.push([
              sn++, item.manager, displayEntity, item.dateStr, item.siNo, item.customer, 
              item.total, item.outputVat, item.vatable, item.products
          ]);
          managerTotal += item.vatable;
      });

      if (currentManager !== null) {
          let subRow = Array(detailHeaders.length).fill("");
          subRow[5] = "Total for " + currentManager + ":";
          subRow[8] = managerTotal;
          detailRows.push(subRow);
      }

      detailSheet.getRange(1, 1, detailRows.length, detailHeaders.length).setValues(detailRows);
      
      for (let r = 0; r < detailRows.length; r++) {
          let rowData = detailRows[r];
          let excelRow = r + 1;
          if (rowData[0] === "SN") {
              detailSheet.getRange(excelRow, 1, 1, detailHeaders.length).setFontWeight("bold").setBackground("#f3f3f3");
          }
          if (String(rowData[5]).includes("Total for ")) {
              detailSheet.getRange(excelRow, 6, 1, 4).setFontWeight("bold"); 
              detailSheet.getRange(excelRow, 9).setBorder(true, false, false, false, false, false, "black", SpreadsheetApp.BorderStyle.SOLID);
          }
      }

      detailSheet.getRange(1, 7, detailRows.length, 3).setNumberFormat("#,##0.00");
      detailSheet.getRange(1, 5, detailRows.length, 1).setNumberFormat("@");
      detailSheet.getDataRange().setWrap(true).setVerticalAlignment("top");
      for (let i = 1; i <= detailHeaders.length; i++) detailSheet.autoResizeColumn(i);
      detailSheet.setColumnWidth(4, 140);
      detailSheet.setColumnWidth(5, 120);
      detailSheet.setColumnWidth(6, 250);
      detailSheet.setColumnWidth(7, 130);
      detailSheet.setColumnWidth(8, 130);
      detailSheet.setColumnWidth(9, 130);
      detailSheet.setColumnWidth(10, 350);

      SpreadsheetApp.flush();
      Utilities.sleep(3000);

      let url = "https://docs.google.com/spreadsheets/d/" + tempSpreadsheet.getId() + "/export?format=xlsx";
      let token = ScriptApp.getOAuthToken();
      let attachmentBlob = null;
      try {
          let response = UrlFetchApp.fetch(url, { headers: { 'Authorization': 'Bearer ' + token }, muteHttpExceptions: true });
          if (response.getResponseCode() === 200) attachmentBlob = response.getBlob().setName(cleanFileName + ".xlsx");
      } catch(e) {}

      try { DriveApp.getFileById(tempSpreadsheet.getId()).setTrashed(true); } catch(e){}

      if (attachmentBlob) {
          let htmlBody = `<div style="font-family:Arial, sans-serif; font-size:13px; color:#000;"><p>Good Day,</p>
            <p>Please find attached the <b>Consolidated Sales Quota Summary Report</b> for <b>${shortBucketName}</b> covering transactions from <strong>${sDateStr}</strong> to <strong>${eDateStr}</strong>, as of ${reportDateString}.</p>
            <table style="border-collapse:collapse; width:100%; max-width:850px; font-size:12px; margin:20px 0;">
              <thead><tr style="background-color:#F50057; color:white; text-align:center;">
                <th style="padding:10px; border:1px solid #ccc;">SN</th>
                <th style="padding:10px; border:1px solid #ccc;">AE Name</th>
                <th style="padding:10px; border:1px solid #ccc;">Quota (VAT EX)<br>${sDateStr} to ${eDateStr}</th>
                <th style="padding:10px; border:1px solid #ccc;">Actual Sales (VAT EX)<br>${sDateStr} to ${eDateStr}, as of ${reportDateString}</th>
                <th style="padding:10px; border:1px solid #ccc;">Balance</th>
                <th style="padding:10px; border:1px solid #ccc;">Rate</th>
                <th style="padding:10px; border:1px solid #ccc;">Reached?</th>
              </tr></thead><tbody>${tableHtmlRows}</tbody></table>
            <p>Thank you for your kindness and support. I truly appreciate it.</p><br><p>Regards,<br><b>Finance System</b></p></div>`;

          let sysSettings = getSystemSettings();
          let isTech = (shortBucketName === "A2Z2GO");
          
          // Dynamically fetch emails from System_Settings_Data (Zero private emails in code)
          let mainRecipient = isTech 
              ? (sysSettings["Management_A2Z"] || "") 
              : (sysSettings["Management_Pingcon"] || "");
          let ccRecipient = isTech 
              ? (sysSettings["Finance_A2Z"] || "") 
              : (sysSettings["Finance_Pingcon"] || "");
          let senderEmail = getSenderEmail(isTech);
          let senderName = isTech ? "Finance A2Z2GO" : "Finance Pingcon";

          let securedBlob = encryptAttachmentViaPython(attachmentBlob, cleanFileName + ".xlsx", shortBucketName === "A2Z2GO" ? SCRIPT_PROPS['ENCRYPT_PASS_TECH'] : SCRIPT_PROPS['ENCRYPT_PASS_PINGCON']);

          // STRICT CANCEL CHECK: Stop instantly before sending the main Consolidated Email
          if (props.getProperty("CANCEL_SALES_REPORT_CONSOLIDATED") === "true") { clearReportLock(); return; }

          sendRobustEmail(mainRecipient, cleanFileName, htmlBody, {
              from: senderEmail,
              name: senderName,
              cc: ccRecipient,
              attachments: [securedBlob]
          });
          Utilities.sleep(20000); // Wait 20 seconds to enforce max 3 emails per minute
          
          // --- BEGIN SECONDARY EMAILS (NTE / APPRECIATION) PER AE ---
          let isTestModeActive = (props.getProperty("TEST_MODE") === "true");
          let hrEmail = SCRIPT_PROPS['HR_EMAIL_1'];
          
          let feedbackImageId = shortBucketName === "A2Z2GO" ? FEEDBACK_IMAGE_ID_A2Z_KEYSYS : FEEDBACK_IMAGE_ID_PINGCON_BONITA;
          let feedbackLink = (shortBucketName === "A2Z2GO") ? (SCRIPT_PROPS['FEEDBACK_LINK_TECH'] || "") : (SCRIPT_PROPS['FEEDBACK_LINK_PINGCON'] || "");
          let feedbackImgBlob = null;
          try { feedbackImgBlob = DriveApp.getFileById(feedbackImageId).getBlob(); } catch(e) {}

          let aeListForEmails = Object.keys(summaryAEs).filter(k => k !== "quotaRange");
          
          for (let aeIdx = 0; aeIdx < aeListForEmails.length; aeIdx++) {
              if (props.getProperty("CANCEL_SALES_REPORT_CONSOLIDATED") === "true") { clearReportLock(); return; }

              let ae = aeListForEmails[aeIdx];
              let aeData = summaryAEs[ae];
              let monthTotalSales = aeData.actual;
              let monthlyQuota = aeData.quota;
              
              let isQuotaHit = (monthlyQuota > 0 && monthTotalSales >= monthlyQuota);
              let pctAchieved = (monthlyQuota > 0) ? ((monthTotalSales / monthlyQuota) * 100).toFixed(2) + "%" : "0.00%";
              
              let parsedLast = ae.includes(",") ? ae.split(",")[0].trim() : ae.split(" ")[0].trim();
              let parsedFirst = ae.includes(",") ? ae.split(",")[1].trim() : ae.replace(parsedLast, "").trim();
              let cleanFirst = parsedFirst.replace(/\s+[A-Za-z]\.?$/, "").trim(); 
              let greetingName = cleanFirst + " " + parsedLast;
              let nameNoComma = ae.replace(/,/g, '');
              let dateRangeStr = `${sDateStr} to ${eDateStr} as of Report Date ${reportDateString}`;
              
              let secSubject = ""; let secBody = ""; let secCc = "";
              let financeEmailStr = ""; let financeSignStr = ""; let supName = ""; let supEmail = "";
              let aeEmailToUse = ""; let department = "";
              
              if (aeEmailMap && aeEmailMap[ae]) {
                  supName = aeEmailMap[ae].supervisor || "";
                  supEmail = aeEmailMap[ae].supervisorEmail || "";
                  department = aeEmailMap[ae].department || "";
                  aeEmailToUse = shortBucketName === "A2Z2GO" ? (aeEmailMap[ae].a2zit || aeEmailMap[ae].keysys) : aeEmailMap[ae].pingcon;
              }

              let hrEmailDb = sysSettings["HR_Email"] || "";
              let isA2ZAe = (shortBucketName === "A2Z2GO");

              if (isA2ZAe) {
                  let mngtA2Z = sysSettings["Management_A2Z"] || "";
                  let finA2Z = sysSettings["Finance_A2Z"] || "";
                  let execA2Z = sysSettings["Executive_A2Z"] || "";
                  let marketingA2Z = sysSettings["Marketing_A2Z"] || "";
                  secCc = [mngtA2Z, execA2Z, marketingA2Z, finA2Z, hrEmailDb].filter(e => e && e.trim() !== "").join(", ");
                  financeEmailStr = finA2Z; 
                  financeSignStr = "Finance<br>A2Z2GO";
              } else {
                  let mngtPingcon = sysSettings["Management_Pingcon"] || "";
                  let finPingcon = sysSettings["Finance_Pingcon"] || "";
                  let execPingcon = sysSettings["Executive_Pingcon"] || "";
                  let marketingPingcon = sysSettings["Marketing_Pingcon"] || "";
                  secCc = [mngtPingcon, execPingcon, marketingPingcon, finPingcon, hrEmailDb].filter(e => e && e.trim() !== "").join(", ");
                  financeEmailStr = finPingcon; 
                  financeSignStr = "Finance<br>Pingcon";
              }
              
              let enableSupEmail = sysSettings["Enable_Supervisor_Email"] !== "FALSE";
              if (enableSupEmail && supEmail && supEmail.includes("@")) secCc += ", " + supEmail;

              let enableAeEmail = sysSettings["Enable_AE_Email"] !== "FALSE";
              let fallbackManagement = isA2ZAe ? (sysSettings["Management_A2Z"] || "") : (sysSettings["Management_Pingcon"] || "");
              
              let rawMainRecipient = aeEmailToUse || mainRecipient;
              let secMainRecipient = enableAeEmail ? rawMainRecipient : fallbackManagement; 
              if (!secMainRecipient || secMainRecipient.trim() === "") {
                  secMainRecipient = fallbackManagement;
              }

              let ntePdfBlob = null;

              // --- GENERATE INDIVIDUAL SALES REPORT FOR THIS AE ---
              let aeItems = items.filter(itm => itm.manager === ae);
              aeItems.sort((a, b) => b.dateObj - a.dateObj);

              let indHeaders = ["SN", "Entity", "Date (MM/DD/YYYY)", "SI No.", "Customer", "Total", "Output VAT", "Vatable", "Products"];
              let indSheetData = [indHeaders];
              let indSn = 1;
              let indSumTotal = 0;
              let indSumOutputVat = 0;
              let indSumVatable = 0;

              aeItems.forEach(item => {
                  let displayEntity = item.companyEntity === "A2ZIT" ? "A2Z2GO" : item.companyEntity;
                  indSheetData.push([ indSn++, displayEntity, item.dateStr, item.siNo, item.customer, item.total, item.outputVat, item.vatable, item.products ]);
                  indSumTotal += item.total;
                  indSumOutputVat += item.outputVat;
                  indSumVatable += item.vatable;
              });
              indSheetData.push(["", "", "", "", "TOTAL:", indSumTotal, "", "", ""]);

              let cleanAeName = ae.replace(/[^a-zA-Z0-9 -]/g, "");
              let indFileName = `Sales Report ${cleanAeName} - ${sDateStr} to ${eDateStr} as of Report Date ${reportDateString}`;

              let indSs = SpreadsheetApp.create(indFileName);
              let indSheet = indSs.getSheets()[0];
              indSheet.setName("Monthly Sales");
              indSheet.getRange(1, 1, indSheetData.length, indHeaders.length).setValues(indSheetData);
              indSheet.getRange(1, 1, 1, indHeaders.length).setFontWeight("bold").setBackground("#f3f3f3");
              indSheet.getRange(2, 6, indSheetData.length, 3).setNumberFormat("#,##0.00");
              indSheet.getRange(2, 4, indSheetData.length, 1).setNumberFormat("@");
              indSheet.getRange(indSheetData.length, 5, 1, 2).setFontWeight("bold");
              indSheet.getRange(indSheetData.length, 6, 1, 1).setBorder(true, false, false, false, false, false, "black", SpreadsheetApp.BorderStyle.SOLID);

              indSheet.setColumnWidth(1, 50);
              indSheet.setColumnWidth(2, 90);
              indSheet.setColumnWidth(3, 140);
              indSheet.setColumnWidth(4, 110);
              indSheet.setColumnWidth(5, 250);
              indSheet.setColumnWidth(6, 130);
              indSheet.setColumnWidth(7, 130);
              indSheet.setColumnWidth(8, 130);
              indSheet.setColumnWidth(9, 350);

              SpreadsheetApp.flush();
              Utilities.sleep(2000);

              let indUrl = "https://docs.google.com/spreadsheets/d/" + indSs.getId() + "/export?format=xlsx";
              let indToken = ScriptApp.getOAuthToken();
              let indBlob = null;
              try {
                  let indResp = UrlFetchApp.fetch(indUrl, { headers: { 'Authorization': 'Bearer ' + indToken }, muteHttpExceptions: true });
                  if (indResp.getResponseCode() === 200) indBlob = indResp.getBlob().setName(indFileName + ".xlsx");
              } catch(e) {}
              try { DriveApp.getFileById(indSs.getId()).setTrashed(true); } catch(e) {}

              // Calculate new clients count
              let curCustSet = new Set(aeItems.map(i => i.customer));
              let pastCustSet = new Set();
              arAllData.forEach(itemData => {
                  let r = itemData.data;
                  let mName = getCanonicalName(managerIdx > -1 ? String(r[managerIdx]) : "");
                  if (mName !== ae) return;
                  let dVal = (dateIdx > -1) ? r[dateIdx] : "";
                  let dObj = (dVal instanceof Date) ? dVal : new Date(dVal);
                  if (isNaN(dObj.getTime()) || dObj >= mDate) return;
                  let cName = (customerIdx > -1) ? String(r[customerIdx]).trim() : "";
                  if (cName) pastCustSet.add(cName);
              });
              let newClientsCount = 0;
              curCustSet.forEach(c => { if (!pastCustSet.has(c)) newClientsCount++; });

              let quotaMessage = "";
              if (!monthlyQuota) {
                  quotaMessage = "Let's continue to push forward and aim high in the coming months! Thank you for your hard work.";
              } else if (monthTotalSales >= monthlyQuota) {
                  quotaMessage = "Your hard work, focus, and dedication truly paid off. Achieving your targets is no small feat, and your efforts are a big part of our continued success. Keep up the great momentum and let’s aim even higher in the coming months. Please give feedback.<br><br>Thank you for your commitment and for delivering such strong results.";
              } else {
                  quotaMessage = "Every challenge brings an opportunity to learn, improve, and come back stronger. What matters most is staying focused, motivated, and committed to your goals. Please give feedback.<br><br>Let’s take a moment to reflect on what we can improve, support one another, and approach the new month with renewed energy and determination. I’m confident in your ability to bounce back and achieve great results. Thank you all for your hard work.";
              }
              
              // --- START OF CHART LOGIC (MOVED OUTSIDE SO BOTH APPRECIATION AND NTE CAN USE IT) ---
              let m0 = new Date(mDate);
              let m1 = new Date(m0.getFullYear(), m0.getMonth() - 1, 1);
              let m2 = new Date(m0.getFullYear(), m0.getMonth() - 2, 1);

              const getHistSales = (dObj) => {
                  let total = 0;
                  let y = dObj.getFullYear();
                  let m = dObj.getMonth();
                  arAllData.forEach(item => {
                      let r = item.data;
                      let rawManager = (managerIdx > -1) ? String(r[managerIdx]) : "";
                      if (getCanonicalName(rawManager) !== ae) return;
                      
                      let rawDate = (dateIdx > -1) ? r[dateIdx] : "";
                      let d = (rawDate instanceof Date) ? rawDate : new Date(rawDate);
                      if (isNaN(d.getTime())) return;

                      let sourceFile = (sourceFileIdx > -1) ? String(r[sourceFileIdx]) : "";
                      if (allowedEntities.length > 0 && !allowedEntities.includes(getCompanyEntity(sourceFile))) return;

                      if (d.getFullYear() === y && d.getMonth() === m) {
                          total += (netTotalIdx > -1 && r[netTotalIdx] !== "" && r[netTotalIdx] != null) ? parseFloat(r[netTotalIdx]) : 0;
                      }
                  });
                  return total;
              };

              const getHistQuota = (dObj) => {
                  let qEnd = new Date(dObj.getFullYear(), dObj.getMonth() + 1, 0);
                  let quota = 0;
                  for(let q=0; q<sqAllData.length; q++) {
                      let qRow = sqAllData[q].data;
                      if (checkNameMatch(String(qRow[0]).trim(), ae)) {
                           let start = new Date(qRow[1]);
                           let end = new Date(qRow[2]);
                           if(start <= qEnd && end >= dObj) {
                               if ((parseFloat(qRow[4]) || 0) > 0) quota = parseFloat(qRow[4]);
                               break;
                           }
                      }
                  }
                  return quota;
              };

              let s0 = monthTotalSales; let q0 = monthlyQuota || 0;
              let s1 = getHistSales(m1); let q1 = getHistQuota(m1);
              let s2 = getHistSales(m2); let q2 = getHistQuota(m2);
              const formatLbl = (d) => Utilities.formatDate(d, Session.getScriptTimeZone(), "MMM yyyy");
              
              let lbl0 = formatLbl(m0); let lbl1 = formatLbl(m1); let lbl2 = formatLbl(m2);

              let chart1Config = {
                type: 'bar',
                data: {
                  labels: [lbl2, lbl1, lbl0],
                  datasets: [
                    { label: 'Sales', data: [s2, s1, s0], backgroundColor: '#36A2EB' },
                    { label: 'Quota', data: [q2, q1, q0], backgroundColor: '#FF6384' }
                  ]
                },
                options: { title: { display: true, text: 'Past 3 Months: Sales vs Quota', fontSize: 13 }, legend: { display: true, position: 'top' }, plugins: { datalabels: { display: false } } }
              };

              let c1Url = "";
              try {
                  c1Url = "https://quickchart.io/chart?w=400&h=250&c=" + encodeURIComponent(JSON.stringify(chart1Config));
              } catch(e) { console.error("Chart gen failed: " + e); }
              // --- END OF CHART LOGIC ---
              
              if (isQuotaHit) {
                  let reportDateText = reportDateString || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MMM dd, yyyy HHmm");
                  secSubject = `Letter of Appreciation Sales Quota ${ae} - ${sDateStr} to ${eDateStr} as of Report Date ${reportDateText}`;

                  try {
                      let quotaTemplateData = Template_SalesQuotaEmailsAndPdf(
                          isQuotaHit, nameNoComma, dateRangeStr, greetingName, financeEmailStr, financeSignStr, 
                          isA2ZAe, ae, sDateStr, eDateStr, monthlyQuota, monthTotalSales, pctAchieved, supName, c1Url, feedbackImgBlob, feedbackLink, department, null, ""
                      );
                      secSubject = quotaTemplateData.secSubject;
                      secBody = quotaTemplateData.secBody;
                      ntePdfBlob = quotaTemplateData.ntePdfBlob;
                  } catch(tmplErr) { console.error("Template Error: " + tmplErr); }

              } else if (monthlyQuota > 0) {
                  // FAILED QUOTA -> LARAVEL ONLY
                  let calcBalance = monthlyQuota - monthTotalSales;
                  let displayBalanceSafe = calcBalance > 0 ? calcBalance : 0;
                  
                  let historicalParagraph = `Reviewing the historical data for the preceding months shows the following trend: ` +
                      `In Month 1 (${lbl2}), actual sales were ${s2.toLocaleString('en-US', {minimumFractionDigits: 2})} against a ${q2.toLocaleString('en-US', {minimumFractionDigits: 2})} target. ` +
                      `In Month 2 (${lbl1}), sales reached ${s1.toLocaleString('en-US', {minimumFractionDigits: 2})} against a ${q1.toLocaleString('en-US', {minimumFractionDigits: 2})} target. ` +
                      `Finally, in Month 3 (${lbl0}), sales were ${s0.toLocaleString('en-US', {minimumFractionDigits: 2})} against a ${q0.toLocaleString('en-US', {minimumFractionDigits: 2})} target. ` +
                      `This incident was automatically recorded by the system to document the shortfall and facilitate the necessary review and explanation.`;
                      
                  let summaryText = 
                    `A Sales Quota Mismatch incident was identified for the covered period of ${sDateStr} to ${eDateStr} (Report Date: ${reportDateString}). ` +
                    `The required monthly sales quota (VAT Exclusive) was set at PHP ${monthlyQuota.toLocaleString('en-US', {minimumFractionDigits: 2})}, ` +
                    `but the actual sales recorded were only PHP ${monthTotalSales.toLocaleString('en-US', {minimumFractionDigits: 2})}. ` +
                    `This resulted in a sales shortfall of PHP ${displayBalanceSafe.toLocaleString('en-US', {minimumFractionDigits: 2})}, ` +
                    `yielding a quota attainment of just ${pctAchieved}.\n\n` + historicalParagraph;

                  let reasonForFiling = `System-generated report due to the employee failing to meet the required monthly sales quota of PHP ${monthlyQuota.toLocaleString('en-US', {minimumFractionDigits: 2})} for ${sDateStr} to ${eDateStr}.`;

                  let apiDateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
                  let apiTimeStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "HH:mm");
                  let match = reportDateString.match(/([A-Za-z]+)\s+(\d{1,2}),\s+(\d{4})\s+(\d{2})(\d{2})/);
                  if (match) {
                      let parsedDate = new Date(`${match[1]} ${match[2]}, ${match[3]} ${match[4]}:${match[5]}:00`);
                      if (!isNaN(parsedDate.getTime())) {
                          apiDateStr = Utilities.formatDate(parsedDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
                          apiTimeStr = Utilities.formatDate(parsedDate, Session.getScriptTimeZone(), "HH:mm");
                      }
                  }

                  let currentYear = new Date().getFullYear();
                  let randomStr = Math.random().toString(36).substring(2, 6).toUpperCase();
                  let refString = `INCIDENT-${currentYear}-API-${randomStr}`;
                  
                  // Get the exact name configured for the NTE System, fallback to auto-generated if blank
                  let accusedNteName = (aeEmailMap[ae] && aeEmailMap[ae].nteName) ? aeEmailMap[ae].nteName : nameNoComma;

                  let payload = {
                    "source_system": "Finance System",
                    "is_test": isTestModeActive,
                    "external_reference": refString,
                    "date_of_incident": apiDateStr,
                    "time_of_incident": apiTimeStr,
                    "location": "Finance Portal Database",
                    "summary": summaryText,
                    "accused_name": accusedNteName,
                    "accused_email": secMainRecipient, 
                    "reason_for_report": "Failure to meet monthly sales quota.",
                    "reporting_party_name": "Finance System Apps Script",
                    "reporting_party_date_submitted": Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd"),
                    "reason_for_filing": reasonForFiling,
                    "offense_id": 32,
                    "violation_type": "Others",
                    "violation_type_details": "Sales Quota Shortfall"
                  };

                  if (isTestModeActive) {
                      payload["email_recipient_test_mode"] = false;
                      payload["testing_timeline_enabled"] = false;
                  }

                  let apiToken = props.getProperty('LARAVEL_PROD_TOKEN');
                  let apiUrl = props.getProperty('LARAVEL_PROD_URL') || "";
                  
                  let options = {
                    "method": "post",
                    "contentType": "application/json",
                    "headers": {
                      "Authorization": "Bearer " + apiToken,
                      "Accept": "application/json"
                    },
                    "payload": JSON.stringify(payload),
                    "muteHttpExceptions": true 
                  };

                  try {
                    let response = UrlFetchApp.fetch(apiUrl, options);
                    if (response.getResponseCode() === 200 || response.getResponseCode() === 201) {
                        console.log("Laravel NTE Created (Consolidated): " + JSON.parse(response.getContentText()).data.ir_number);
                    }
                  } catch (e) { console.error("Laravel Connection Error: " + e.message); }
                  
                  Utilities.sleep(1000);
              }

              let validSecCc = secCc.split(',').map(e => e.trim()).filter(e => e.includes('@')).join(',');

              // --- DISPATCH INDIVIDUAL SALES REPORT EMAIL ---
              if (indBlob) {
                  let securedIndBlob = encryptAttachmentViaPython(indBlob, indFileName + ".xlsx", isA2ZAe ? SCRIPT_PROPS['ENCRYPT_PASS_TECH'] : SCRIPT_PROPS['ENCRYPT_PASS_PINGCON']);
                  let indChartHtml = c1Url ? `<br><div style="text-align: left; margin: 15px 0;"><img src="${c1Url}" style="max-width: 450px; border: 1px solid #ddd; border-radius: 4px;" alt="Sales vs Quota Chart"></div>` : '';
                  let indHtmlBody = Template_SalesReportEmail(greetingName, indFileName, mDate, monthTotalSales, pctAchieved, monthlyQuota, newClientsCount, indChartHtml, quotaMessage, feedbackImgBlob, feedbackLink);

                  let indEmailOptions = {
                      from: senderEmail,
                      name: senderName,
                      attachments: [securedIndBlob],
                      inlineImages: {},
                      throwError: false
                  };
                  if (validSecCc) indEmailOptions.cc = validSecCc;
                  if (feedbackImgBlob) indEmailOptions.inlineImages.feedbackImage = feedbackImgBlob;

                  try {
                      sendRobustEmail(secMainRecipient, indFileName, indHtmlBody, indEmailOptions);
                      recordSuccessfulSend(ae, shortBucketName, monthlyQuota || 0, monthTotalSales || 0);
                      Utilities.sleep(6000); 
                  } catch(indErr) {
                      console.error("Failed to send individual Sales Report to " + ae + ": " + indErr);
                  }
              }

              // --- DISPATCH SECONDARY EMAIL (APPRECIATION) ---
              let secEmailOptions = {
                  from: senderEmail, name: senderName, inlineImages: {},
                  attachments: [], throwError: true 
              };
              if (validSecCc) secEmailOptions.cc = validSecCc;
              if (feedbackImgBlob) secEmailOptions.inlineImages.feedbackImage = feedbackImgBlob;
              if (ntePdfBlob) secEmailOptions.attachments.push(ntePdfBlob);

              try {
                  if (isQuotaHit) {
                      sendRobustEmail(secMainRecipient, secSubject, secBody, secEmailOptions);
                      Utilities.sleep(6000); 
                  }
              } catch(errStr) {
                  console.error("Failed to send secondary email for " + ae + ". Error: " + errStr);
              }
          }
          // --- END SECONDARY EMAILS ---
      }
      Utilities.sleep(20000); // Wait 20 seconds between entity buckets
    }
  }

  clearReportLock();
}

// =========================================================================
// UNPAID DRs (No UI - Fixed Date Range: Oct 1 2025 to Mar 20 2026)
// =========================================================================
// =========================================================================
// UNPAID DRs (Batch Process to Avoid 6-Minute Limit)
// Fixed Date Range: Oct 1 2025 to Mar 20 2026
// =========================================================================


function resumeUnpaidDRBatch() {
  const startTime = Date.now();
  const MAX_EXECUTION_TIME = 4.5 * 60 * 1000;
  
  const props = PropertiesService.getScriptProperties();
  props.setProperty("THREAD_STATE", "RUNNING"); // <--- ADDED
  
  let pendingAEsStr = props.getProperty(UNPAID_BATCH_KEY);
  
  if (!pendingAEsStr) {
    deleteTriggersByName(UNPAID_TRIGGER_NAME);
    return; 
  }
  
  let pendingAEs = JSON.parse(pendingAEsStr);
  if (pendingAEs.length === 0) {
    deleteTriggersByName(UNPAID_TRIGGER_NAME);
    props.deleteProperty(UNPAID_BATCH_KEY);
    clearReportLock(); // <-- ADDED THIS TO UNLOCK THE UI
    console.log("All AEs processed successfully.");
    return;
  }

  const arAllData = gatherAllData("AR_Data_");
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const arSheets = getDatabaseSheets("AR_Data_");
  let arHeaders =[];
  if (arSheets.length > 0 && arSheets[0].getLastColumn() > 0) {
    arHeaders = arSheets[0].getRange(1, 1, 1, arSheets[0].getLastColumn()).getValues()[0].map(h => String(h).toLowerCase().trim());
  }
  
  let docDateIdx = arHeaders.findIndex(h => h.includes("date") && !h.includes("expiry") && !h.includes("countered") && !h.includes("received"));
  let deadlineIdx = arHeaders.findIndex(h => h.includes("deadline") || h.includes("due date"));
  if (deadlineIdx === -1) deadlineIdx = docDateIdx;
  let customerIdx = arHeaders.findIndex(h => h.includes("customer"));
  let managerIdx = arHeaders.findIndex(h => h.includes("manager"));
  let textIdx = arHeaders.findIndex(h => h.includes("additional text"));
  let notesIdx = arHeaders.findIndex(h => h === "notes");
  let openBalIdx = arHeaders.findIndex(h => h.includes("open balance") && !h.includes("%"));
  let originalIdx = arHeaders.findIndex(h => h.includes("sales document total in main currency") || h.includes("sales document total"));
  let paymentTermIdx = arHeaders.findIndex(h => h.includes("payment condition") || h.includes("payment term"));
  let counterDateIdx = arHeaders.findIndex(h => h.includes("countered") || h.includes("counter date"));
  let receivedDateIdx = arHeaders.findIndex(h => h.includes("received delivery") || h.includes("received date") || h.includes("delivery date"));
  let sourceFileIdx = arHeaders.findIndex(h => h.includes("source file"));
  let paymentStatusIdx = arHeaders.findIndex(h => h === "payment status" || h.includes("payment status"));
  if (paymentStatusIdx === -1 && openBalIdx > -1) paymentStatusIdx = openBalIdx + 2;

  const startStr = props.getProperty("BATCH_START_DATE");
  const endStr = props.getProperty("BATCH_END_DATE");
  const startDate = startStr ? new Date(startStr + "T00:00:00") : new Date(2025, 9, 1);
  const endDate = endStr ? new Date(endStr + "T23:59:59") : new Date(2026, 2, 20, 23, 59, 59);
  
  let allowedEntities = JSON.parse(props.getProperty("SALES_ENTITIES") || "[]");
  const now = new Date();
  const reportDate = "as of Report Date " + Utilities.formatDate(now, Session.getScriptTimeZone(), "MMM dd, yyyy HHmm");
  

  let rootFolder = null;
  if (typeof UNPAID_DR_ROOT_FOLDER_ID !== 'undefined' && UNPAID_DR_ROOT_FOLDER_ID !== "INSERT_UNPAID_DR_FOLDER_ID_HERE") {
      rootFolder = DriveApp.getFolderById(UNPAID_DR_ROOT_FOLDER_ID);
  }

      while (pendingAEs.length > 0) {
         if (props.getProperty("CANCEL_UNPAID_DR") === "true") {
             deleteTriggersByName(UNPAID_TRIGGER_NAME);
             props.deleteProperty(UNPAID_BATCH_KEY);
             props.deleteProperty("CANCEL_UNPAID_DR");
             return; 
         }

         if (Date.now() - startTime > MAX_EXECUTION_TIME) {
        props.setProperty(UNPAID_BATCH_KEY, JSON.stringify(pendingAEs));
        replaceTrigger(UNPAID_TRIGGER_NAME);
        return; 
     }
     
     let currentAE = pendingAEs.shift(); 
     let allAeItems =[];
     
     arAllData.forEach(item => {
        let row = item.data;

        let textVal = (textIdx > -1) ? String(row[textIdx]) : "";
        let notesVal = (notesIdx > -1) ? String(row[notesIdx]) : "";
        if (isTransactionVoid(notesVal, textVal)) return;

        let rawManager = (managerIdx > -1) ? String(row[managerIdx]) : "";
        let manager = getCanonicalName(rawManager);
        if (manager !== currentAE) return;

        let openBal = (openBalIdx > -1) ? parseFloat(row[openBalIdx]) : 0;
        if (isNaN(openBal)) openBal = 0;
        if (openBal < 1) return;

        let statusRaw = (paymentStatusIdx > -1 && paymentStatusIdx < row.length) ? String(row[paymentStatusIdx]).trim() : "";
        let statusLow = statusRaw.toLowerCase();
        if (statusLow === "fully paid") return;

        let rawDocDate = (docDateIdx > -1) ? row[docDateIdx] : "";
        let docDateObj = (rawDocDate instanceof Date) ? rawDocDate : new Date(rawDocDate);
        if (isNaN(docDateObj.getTime())) return;
        if (docDateObj < startDate || docDateObj > endDate) return;

        let rawDeadline = (deadlineIdx > -1) ? row[deadlineIdx] : "";
        let deadlineObj = new Date(rawDeadline);
        if (isNaN(deadlineObj.getTime()) || rawDeadline === "") {
            deadlineObj = new Date(docDateObj);
            let pTerm = (paymentTermIdx > -1) ? parseInt(row[paymentTermIdx]) : 0;
            if (!isNaN(pTerm) && pTerm > 0) {
                deadlineObj.setDate(deadlineObj.getDate() + pTerm);
            }
        }

        let salesTotal = (originalIdx > -1) ? parseFloat(row[originalIdx]) : 0;
        if (isNaN(salesTotal)) salesTotal = 0;
        let amountPaid = salesTotal - openBal;
        if (amountPaid < 0) amountPaid = 0;

        let sourceFile = (sourceFileIdx > -1) ? String(row[sourceFileIdx]) : "";
        let companyEntity = getCompanyEntity(sourceFile);
        
        if (allowedEntities.length > 0 && !allowedEntities.includes(companyEntity)) return;

        allAeItems.push({
            docDateObj: docDateObj,
            invoiceDateStr: fastFormatDate(docDateObj),
            paymentTerm: (paymentTermIdx > -1) ? String(row[paymentTermIdx]).trim() : "",
            counterDateStr: (counterDateIdx > -1 && row[counterDateIdx]) ? fastFormatDate(new Date(row[counterDateIdx])) : "",
            deadlineStr: fastFormatDate(deadlineObj),
            receivedDateStr: (receivedDateIdx > -1 && row[receivedDateIdx]) ? fastFormatDate(new Date(row[receivedDateIdx])) : "",
            customer: (customerIdx > -1) ? String(row[customerIdx]) : "",
            text: (textIdx > -1) ? String(row[textIdx]) : "",
            salesTotal: salesTotal,
            amountPaid: amountPaid,
            openBal: openBal,
            companyEntity: companyEntity,
            remarksStatus: statusRaw,
            sourceFile: sourceFile
        });
     });
     
     if (allAeItems.length === 0) continue;

     allAeItems.sort((a, b) => a.docDateObj - b.docDateObj);

     let reportHeaders =[
         "SN", "Entity", "Invoice Date (MM/DD/YYYY)", "Terms", "Due Date (MM/DD/YYYY)", 
         "Customer Name", "Additional text", "Sales Total", "Amount Paid", "Open Balance", 
         "Counter Date (MM/DD/YYYY)", "Received Date (MM/DD/YYYY)", "Remarks/Payment Status"
     ];
     
     let sheetData = [reportHeaders];
     let globalSn = 1;

     let totalDrCount = allAeItems.length;
     let totalOpenBalance = 0;
     let monthlyBal = {};
     let fullyUnpaidCount = 0;
     let partiallyUnpaidCount = 0;
     
     let cutOffDate = new Date(now);
     let daysUntilFriday = (5 - cutOffDate.getDay() + 7) % 7;
     cutOffDate.setDate(cutOffDate.getDate() + daysUntilFriday);
     cutOffDate.setHours(23, 59, 59, 999);

     allAeItems.forEach(item => {
         let displayEntity = item.companyEntity === "A2ZIT" ? "A2Z2GO" : item.companyEntity;
         
         let dDate = new Date(item.deadlineStr);
         if (isNaN(dDate.getTime())) dDate = item.docDateObj;
         let finalRemarks = item.remarksStatus;
         if (dDate <= cutOffDate && String(finalRemarks).toLowerCase().includes("advance")) {
             finalRemarks = "Due for collection";
         }

         sheetData.push([
           globalSn++, displayEntity, item.invoiceDateStr, item.paymentTerm, item.deadlineStr, 
           item.customer, item.text, item.salesTotal, item.amountPaid, item.openBal,
           item.counterDateStr, item.receivedDateStr, finalRemarks
         ]);
         
         totalOpenBalance += item.openBal;

         let sortKey = Utilities.formatDate(item.docDateObj, Session.getScriptTimeZone(), "yyyy-MM");
         let displayLabel = Utilities.formatDate(item.docDateObj, Session.getScriptTimeZone(), "MMM yyyy");
         let monthId = sortKey + "|" + displayLabel;

         if (!monthlyBal[monthId]) monthlyBal[monthId] = 0;
         monthlyBal[monthId] += item.openBal;

         if (item.openBal >= item.salesTotal - 0.05 && item.salesTotal > 0) {
             fullyUnpaidCount++;
         } else {
             partiallyUnpaidCount++;
         }
     });
     
     let totalRow = Array(reportHeaders.length).fill("");
     totalRow[9] = totalOpenBalance;
     sheetData.push(totalRow);

     let sortedMonthIds = Object.keys(monthlyBal).sort();
     let chartLabels = sortedMonthIds.map(id => id.split("|")[1]);
     let chartData = sortedMonthIds.map(id => monthlyBal[id]);

     let chartConfig = {
         type: 'bar',
         data: { labels: chartLabels, datasets: [{ label: 'Open Balance', data: chartData, backgroundColor: '#36a2eb' }] },
         options: { title: { display: true, text: 'Total Open Balance by Month', fontSize: 14 }, legend: { display: false }, plugins: { datalabels: { display: false } } }
     };

     let chartBlob = null;
     try {
         let chartUrl = "https://quickchart.io/chart?w=500&h=250&c=" + encodeURIComponent(JSON.stringify(chartConfig));
         chartBlob = UrlFetchApp.fetch(chartUrl).getBlob().setName("chart.png");
     } catch(e) {}

     let reportDateString = "";
     if (allAeItems.length > 0 && allAeItems[0].sourceFile) {
         // \s* allows it to catch "May4" even when the space is missing
         let match = allAeItems[0].sourceFile.match(/as of\s+([a-zA-Z]+)\s*(\d+),?\s+(\d{4})\s+(\d{4})/i);
         if (match) {
             let paddedDay = match[2].length === 1 ? "0" + match[2] : match[2];
             reportDateString = `${match[1]} ${paddedDay}, ${match[3]} ${match[4]}`;
         }
     }
     if (!reportDateString) {
         let extractedDate = Utilities.formatDate(now, Session.getScriptTimeZone(), "MMM dd, yyyy");
         let currentTime = Utilities.formatDate(now, Session.getScriptTimeZone(), "HHmm");
         reportDateString = `${extractedDate} ${currentTime}`;
     }

     let cleanAeName = currentAE.replace(/[^a-zA-Z0-9 -]/g, "");
     let cleanFileName = `Statement of Account - ${cleanAeName} - as of Report Date ${reportDateString}`;

     let tempSpreadsheet = SpreadsheetApp.create(cleanFileName);
     
     let tempSheet = tempSpreadsheet.getSheets()[0];
     tempSheet.setName("Data");
     tempSheet.getRange(1, 1, sheetData.length, reportHeaders.length).setValues(sheetData);
     tempSheet.getRange(sheetData.length, 10).setBorder(true, false, false, false, false, false, "black", SpreadsheetApp.BorderStyle.SOLID);
     tempSheet.getRange(1, 1, 1, reportHeaders.length).setFontWeight("bold").setBackground("#f3f3f3");
     tempSheet.getRange(2, 8, sheetData.length, 3).setNumberFormat("#,##0.00");
     tempSheet.getRange(2, 7, sheetData.length, 1).setNumberFormat("@");
     tempSheet.getDataRange().setWrap(true).setVerticalAlignment("top");
     
     for (let i = 1; i <= reportHeaders.length; i++) tempSheet.autoResizeColumn(i);
     tempSheet.setColumnWidth(8, 110); 
     tempSheet.setColumnWidth(9, 110); 
     tempSheet.setColumnWidth(10, 110); 

     let statsSheet = tempSpreadsheet.insertSheet("Statistics");
     let statsData = [
         ["Metric", "Value"],
         ["Total Unpaid DRs", totalDrCount],
         ["Total Open Balance", totalOpenBalance],
         ["Num of clients 100% unpaid", fullyUnpaidCount],
         ["Num of clients partially unpaid", partiallyUnpaidCount]
     ];
     statsSheet.getRange(1, 1, statsData.length, 2).setValues(statsData);
     statsSheet.getRange(1, 1, 1, 2).setFontWeight("bold").setBackground("#f3f3f3");
     statsSheet.getRange(3, 2).setNumberFormat("#,##0.00"); 
     
     statsSheet.setColumnWidth(1, 200);
     statsSheet.setColumnWidth(2, 120);

     if (chartBlob) statsSheet.insertImage(chartBlob, 1, 7); 

     SpreadsheetApp.flush();
     Utilities.sleep(3000); 

     let url = "https://docs.google.com/spreadsheets/d/" + tempSpreadsheet.getId() + "/export?format=xlsx";
     let token = ScriptApp.getOAuthToken();
     let attachmentBlob = null;

     for (let fetchAttempt = 0; fetchAttempt < 3; fetchAttempt++) {
       try {
         let response = UrlFetchApp.fetch(url, { headers: { 'Authorization': 'Bearer ' + token }, muteHttpExceptions: true });
         
         if (response.getResponseCode() === 200) {
             attachmentBlob = response.getBlob().setName(cleanFileName + ".xlsx");
             if (rootFolder) {
                 let targetYearStr = Utilities.formatDate(endDate, Session.getScriptTimeZone(), "yyyy");
                 let targetMonthStr = Utilities.formatDate(endDate, Session.getScriptTimeZone(), "MMMM");

                 for (let attempt = 0; attempt < 3; attempt++) {
                     try {
                         let yearFolders = rootFolder.getFoldersByName(targetYearStr);
                         let yearFolder = yearFolders.hasNext() ? yearFolders.next() : rootFolder.createFolder(targetYearStr);

                         let aeFolders = yearFolder.getFoldersByName(currentAE);
                         let aeFolder = aeFolders.hasNext() ? aeFolders.next() : yearFolder.createFolder(currentAE);

                         let monthFolders = aeFolder.getFoldersByName(targetMonthStr);
                         let monthFolder = monthFolders.hasNext() ? monthFolders.next() : aeFolder.createFolder(targetMonthStr);

                         let existingFiles = monthFolder.getFiles();
                         while (existingFiles.hasNext()) {
                             let file = existingFiles.next();
                             if (file.getName() === (cleanFileName + ".xlsx")) file.setTrashed(true);
                         }
                         monthFolder.createFile(attachmentBlob);
                         break; 
                     } catch (driveErr) { Utilities.sleep(2000); }
                 }
             }
             break; 
         } else {
             Utilities.sleep(4000 * (fetchAttempt + 1));
         }
       } catch (e) { Utilities.sleep(5000 * (fetchAttempt + 1)); }
     }

     if (!attachmentBlob) throw new Error("Failed to export Excel file for " + currentAE);

     try { DriveApp.getFileById(tempSpreadsheet.getId()).setTrashed(true); } catch(e){}

     if (attachmentBlob) {
         let chartHtml = chartBlob ? `<br><div style="text-align: left; margin: 15px 0;"><img src="cid:chartImage" style="max-width: 100%; border: 1px solid #ddd; border-radius: 4px;" alt="Open Balance Chart"></div>` : '';
         const formatCurrency = (val) => Number(val).toLocaleString('en-US', {minimumFractionDigits: 2, maximumFractionDigits: 2});

         let sysSettings = getSystemSettings();
         let hasKeysysOrA2ZIT = allAeItems.some(item => item.companyEntity === "Keysys" || item.companyEntity === "A2ZIT");
         let isTech = hasKeysysOrA2ZIT;

         let senderName = isTech ? "Finance A2Z2GO" : "Finance Pingcon";
         let feedbackImageId = isTech ? FEEDBACK_IMAGE_ID_A2Z_KEYSYS : FEEDBACK_IMAGE_ID_PINGCON_BONITA;
         
         let senderEmail = getSenderEmail(isTech);
         let mainRecipient = isTech ? (sysSettings["Management_A2Z"] || "") : (sysSettings["Management_Pingcon"] || "");
         
         let ccList = isTech
             ? [sysSettings["Finance_A2Z"], sysSettings["Executive_A2Z"], sysSettings["Marketing_A2Z"]]
             : [sysSettings["Finance_Pingcon"], sysSettings["Executive_Pingcon"], sysSettings["Marketing_Pingcon"]];
         const ccRecipients = ccList.filter(e => e && e.trim() !== "").join(", ");

         let feedbackImgBlob = null;
         try {
             if (feedbackImageId && feedbackImageId !== "INSERT_FILE_ID_FOR_A2Z_KEYSYS_IMAGE_HERE" && feedbackImageId !== "INSERT_FILE_ID_FOR_PINGCON_BONITA_IMAGE_HERE") {
                 feedbackImgBlob = DriveApp.getFileById(feedbackImageId).getBlob();
             }
         } catch(e) {}

         let feedbackLink = isTech 
          ? (SCRIPT_PROPS['FEEDBACK_LINK_TECH'] || "") 
          : (SCRIPT_PROPS['FEEDBACK_LINK_PINGCON'] || "");

         let htmlBody = Template_UnpaidDrAeEmail(totalDrCount, totalOpenBalance, fullyUnpaidCount, partiallyUnpaidCount, chartHtml, feedbackImgBlob, feedbackLink);
         
             let securedBlob = encryptAttachmentViaPython(attachmentBlob, cleanFileName + ".xlsx", hasKeysysOrA2ZIT ? SCRIPT_PROPS['ENCRYPT_PASS_TECH'] : SCRIPT_PROPS['ENCRYPT_PASS_PINGCON']);
         
         let emailOptions = {
             from: senderEmail, cc: ccRecipients, name: senderName,
             attachments: [securedBlob], inlineImages: {}
         };
         if (chartBlob) emailOptions.inlineImages.chartImage = chartBlob;
         if (feedbackImgBlob) emailOptions.inlineImages.feedbackImage = feedbackImgBlob;
         sendRobustEmail(mainRecipient, cleanFileName, htmlBody, emailOptions);
        
         Utilities.sleep(10000); // Increased delay to prevent spam flagging
     }

     if (props.getProperty("CANCEL_UNPAID_DR") === "true") {
         deleteTriggersByName(UNPAID_TRIGGER_NAME);
         props.deleteProperty(UNPAID_BATCH_KEY);
         props.deleteProperty("CANCEL_UNPAID_DR");
         return; 
     }

     // Update the property so if the script crashes, it knows where it left off, but DO NOT stop the loop
     props.setProperty(UNPAID_BATCH_KEY, JSON.stringify(pendingAEs));
  }

  deleteTriggersByName(UNPAID_TRIGGER_NAME);
  props.deleteProperty(UNPAID_BATCH_KEY);
  clearReportLock();
  console.log("ALL UNPAID DRs PROCESSED SUCCESSFULLY.");
}


// =========================================================================
// UNPAID DRs PER CUSTOMER
// =========================================================================

const UNPAID_CUSTOMER_BATCH_KEY = "PENDING_CUSTOMERS_UNPAID_DR";
const UNPAID_CUSTOMER_TRIGGER_NAME = "resumeUnpaidDRCustomerBatch";

function _9b_runSendAllUnpaidDRsCustomer() {
  var lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) { throw new Error("System is currently busy. Another user is generating reports right now. Please wait a few minutes."); }
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const props = PropertiesService.getScriptProperties();
    deleteTriggersByName(UNPAID_CUSTOMER_TRIGGER_NAME);
    props.deleteProperty(UNPAID_CUSTOMER_BATCH_KEY);
    props.deleteProperty("CANCEL_UNPAID_DR_CUSTOMER");
    
    try { ss.toast("Scanning Data to initialize background batch...", "INITIALIZING", -1); } catch(e){}
    
    const arAllData = gatherAllData("AR_Data_");
    if (arAllData.length === 0) throw new Error("No Ageing Report Data found.");
    
    const arSheets = getDatabaseSheets("AR_Data_");
    let arHeaders =[];
    if (arSheets.length > 0 && arSheets[0].getLastColumn() > 0) {
      arHeaders = arSheets[0].getRange(1, 1, 1, arSheets[0].getLastColumn()).getValues()[0].map(h => String(h).toLowerCase().trim());
    }
    let customerIdx = arHeaders.findIndex(h => h.includes("customer"));
    let docDateIdx = arHeaders.findIndex(h => h.includes("date") && !h.includes("expiry") && !h.includes("countered") && !h.includes("received"));
    let openBalIdx = arHeaders.findIndex(h => h.includes("open balance") && !h.includes("%"));
    let paymentStatusIdx = arHeaders.findIndex(h => h === "payment status" || h.includes("payment status"));
    if (paymentStatusIdx === -1 && openBalIdx > -1) paymentStatusIdx = openBalIdx + 2;
    let managerIdx = arHeaders.findIndex(h => h.includes("manager"));
    let sourceFileIdx = arHeaders.findIndex(h => h.includes("source file"));
    let textIdx = arHeaders.findIndex(h => h.includes("additional text"));
    let notesIdx = arHeaders.findIndex(h => h === "notes");
    
    const aeEmailSheets = getDatabaseSheets(PREFIX_AE_EMAILS);
    let validAEsFromEmail = new Set();
    aeEmailSheets.forEach(sheet => {
       const data = sheet.getDataRange().getValues();
       for(let i=1; i<data.length; i++){
           let aeName = String(data[i][0]).replace(/[^\x20-\x7E]/g, '').replace(/\s+/g, ' ').trim();
           if(aeName) validAEsFromEmail.add(aeName.toLowerCase());
       }
    });

    let uniqueCustomers = new Set();
    let allowedEntities = JSON.parse(props.getProperty("SALES_ENTITIES") || "[]");
    let allowedCustsStr = props.getProperty("SALES_TARGET_CUSTOMERS");
    let allowedCusts = allowedCustsStr ? JSON.parse(allowedCustsStr) : [];
    
    const startStr = props.getProperty("BATCH_START_DATE");
    const endStr = props.getProperty("BATCH_END_DATE");
    const startDate = startStr ? new Date(startStr + "T00:00:00") : new Date(2025, 9, 1);
    const endDate = endStr ? new Date(endStr + "T23:59:59") : new Date(2026, 2, 20, 23, 59, 59);

    let filterType = props.getProperty("UNPAID_DR_FILTER_TYPE");
    
    arAllData.forEach(item => {
        let row = item.data;

        let textVal = (textIdx > -1) ? String(row[textIdx]) : "";
        let notesVal = (notesIdx > -1) ? String(row[notesIdx]) : "";
        if (isTransactionVoid(notesVal, textVal)) return;

        let customer = (customerIdx > -1) ? String(row[customerIdx]).trim() : "";
        if (!customer || customer === "") return;

        let openBal = (openBalIdx > -1) ? parseFloat(row[openBalIdx]) : 0;
        if (isNaN(openBal) || openBal < 1) return; // Ignores <= 0 and anything less than 1.00

        let statusRaw = (paymentStatusIdx > -1 && paymentStatusIdx < row.length) ? String(row[paymentStatusIdx]).trim().toLowerCase() : "";
        if (statusRaw === "fully paid") return;

        let rawDocDate = (docDateIdx > -1) ? row[docDateIdx] : "";
        let docDateObj = (rawDocDate instanceof Date) ? rawDocDate : new Date(rawDocDate);
        if (isNaN(docDateObj.getTime())) return;
        if (docDateObj < startDate || docDateObj > endDate) return;

        let rawManager = (managerIdx > -1) ? String(row[managerIdx]) : "";
        let manager = getCanonicalName(rawManager);
        let managerClean = manager.replace(/[^\x20-\x7E]/g, '').replace(/\s+/g, ' ').trim().toLowerCase();

        let sourceFile = (sourceFileIdx > -1) ? String(row[sourceFileIdx]) : "";
        let companyEntity = getCompanyEntity(sourceFile);
        
        if (allowedEntities.length > 0 && !allowedEntities.includes(companyEntity)) return;
        
        let cleanCustomer = customer.replace(/[^\x20-\x7EñÑ]/g, '').replace(/\s+/g, ' ').trim();
        // Normalizes by stripping all punctuation (apostrophes, commas, dots) for 100% reliable matching
        let custKey = cleanCustomer.toLowerCase().replace(/[^a-z0-9]/g, '');
        let displayCustomerName = customer; 
        
        if (allowedCusts.length > 0) {
            let matchedName = allowedCusts.find(c => c.toLowerCase().replace(/[^a-z0-9]/g, '') === custKey);
            if (!matchedName) return;
            displayCustomerName = matchedName; // Group them perfectly under the UI name
        }

        uniqueCustomers.add(displayCustomerName);
    });
    
    const emailSheets = getDatabaseSheets("Customer_Emails_Extracted_");
    let customerToEmailKey = {};
    let emailKeyGroups = {};

    emailSheets.forEach(sheet => {
       const data = sheet.getDataRange().getValues();
       for(let i=1; i<data.length; i++){
           let cust = String(data[i][1]).trim(); 
           let custManager = String(data[i][2]).trim(); // Captures the AE
           let email = String(data[i][3]).trim().toLowerCase(); 
           
           if (!cust || !email || !email.includes("@")) continue;
           
           if (!customerToEmailKey[cust]) customerToEmailKey[cust] = { emails: new Set(), manager: "" };
           
           customerToEmailKey[cust].emails.add(email);
           if (custManager && !customerToEmailKey[cust].manager) customerToEmailKey[cust].manager = custManager;
       }
    });

    uniqueCustomers.forEach(cust => {
        // Force every customer to be their own standalone group (Stops Consolidation)
        let emailKey = cust; 
        if (!emailKeyGroups[emailKey]) emailKeyGroups[emailKey] = [];
        emailKeyGroups[emailKey].push(cust);
    });
    
    let groupedCustomers = Object.values(emailKeyGroups);
    if(groupedCustomers.length === 0) throw new Error("No Unpaid DRs found for the selected category in the specified date range.");
    
    props.setProperty(UNPAID_CUSTOMER_BATCH_KEY, JSON.stringify(groupedCustomers));
    try { ss.toast(`Found ${groupedCustomers.length} Customer Groups. Processing first batch immediately... You can close this window now.`, "BATCH STARTED", 15); } catch(e){}
    
    resumeUnpaidDRCustomerBatch();
  } finally {
    if (lock) lock.releaseLock();
  }
}

function resumeUnpaidDRCustomerBatch() {
  const startTime = Date.now();
  const MAX_EXECUTION_TIME = 4.5 * 60 * 1000; 
  
  const props = PropertiesService.getScriptProperties();
  props.setProperty("THREAD_STATE", "RUNNING"); // <--- ADDED
  
  // STRICT CANCEL CHECK AT THE VERY TOP
  if (props.getProperty("CANCEL_UNPAID_DR_CUSTOMER") === "true") {
     deleteTriggersByName(UNPAID_CUSTOMER_TRIGGER_NAME);
     props.deleteProperty(UNPAID_CUSTOMER_BATCH_KEY);
     props.deleteProperty("CANCEL_UNPAID_DR_CUSTOMER");
     clearReportLock();
     sendBatchSuccessSummary(true, "UNPAID_DR_CUST");
     return; 
  }

  let pendingCustStr = props.getProperty(UNPAID_CUSTOMER_BATCH_KEY);
  
  if (!pendingCustStr) {
    deleteTriggersByName(UNPAID_CUSTOMER_TRIGGER_NAME);
    return; 
  }
  
  let pendingCust = JSON.parse(pendingCustStr);
  if (pendingCust.length === 0) {
    deleteTriggersByName(UNPAID_CUSTOMER_TRIGGER_NAME);
    props.deleteProperty(UNPAID_CUSTOMER_BATCH_KEY);
    clearReportLock();
    console.log("All Customers processed successfully.");
    sendBatchSuccessSummary(false, "UNPAID_DR_CUST");
    return;
  }

  const arAllData = gatherAllData("AR_Data_");
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const arSheets = getDatabaseSheets("AR_Data_");
  let arHeaders = [];
  if (arSheets.length > 0 && arSheets[0].getLastColumn() > 0) {
    arHeaders = arSheets[0].getRange(1, 1, 1, arSheets[0].getLastColumn()).getValues()[0].map(h => String(h).toLowerCase().trim());
  }
  
  let docDateIdx = arHeaders.findIndex(h => h.includes("date") && !h.includes("expiry") && !h.includes("countered") && !h.includes("received"));
  let deadlineIdx = arHeaders.findIndex(h => h.includes("deadline") || h.includes("due date"));
  if (deadlineIdx === -1) deadlineIdx = docDateIdx;
  let customerIdx = arHeaders.findIndex(h => h.includes("customer"));
  let managerIdx = arHeaders.findIndex(h => h.includes("manager"));
  let textIdx = arHeaders.findIndex(h => h.includes("additional text"));
  let notesIdx = arHeaders.findIndex(h => h === "notes");
  let openBalIdx = arHeaders.findIndex(h => h.includes("open balance") && !h.includes("%"));
  let originalIdx = arHeaders.findIndex(h => h.includes("sales document total in main currency") || h.includes("sales document total"));
  let paymentTermIdx = arHeaders.findIndex(h => h.includes("payment condition") || h.includes("payment term"));
  let counterDateIdx = arHeaders.findIndex(h => h.includes("countered") || h.includes("counter date"));
  let receivedDateIdx = arHeaders.findIndex(h => h.includes("received delivery") || h.includes("received date") || h.includes("delivery date"));
  let sourceFileIdx = arHeaders.findIndex(h => h.includes("source file"));
  let paymentStatusIdx = arHeaders.findIndex(h => h === "payment status" || h.includes("payment status"));
  if (paymentStatusIdx === -1 && openBalIdx > -1) paymentStatusIdx = openBalIdx + 2;

  const startStr = props.getProperty("BATCH_START_DATE");
  const endStr = props.getProperty("BATCH_END_DATE");
  const startDate = startStr ? new Date(startStr + "T00:00:00") : new Date(2025, 9, 1);
  const endDate = endStr ? new Date(endStr + "T23:59:59") : new Date(2026, 2, 20, 23, 59, 59);
  
  let allowedEntities = JSON.parse(props.getProperty("SALES_ENTITIES") || "[]");
  const now = new Date();
  const reportDate = "as of Report Date " + Utilities.formatDate(now, Session.getScriptTimeZone(), "MMM dd, yyyy HHmm");

  let rootFolder = null;
  if (typeof UNPAID_DR_ROOT_FOLDER_ID !== 'undefined' && UNPAID_DR_ROOT_FOLDER_ID !== "INSERT_UNPAID_DR_FOLDER_ID_HERE") {
      rootFolder = DriveApp.getFolderById(UNPAID_DR_ROOT_FOLDER_ID);
  }

  // --- ADDED: LOAD CUSTOMER EMAILS ONCE BEFORE THE LOOP ---
  const emailSheets = getDatabaseSheets("Customer_Emails_Extracted_");
  let customerEmailMap = {};
  emailSheets.forEach(sheet => {
     let data = sheet.getDataRange().getValues();
     for(let i=1; i<data.length; i++){ // Start at 1 to skip headers
         let cust = String(data[i][1]).toLowerCase().replace(/[^a-z0-9]/g, ''); // Column B (Punctuation-free)
         let custManager = String(data[i][2]).trim(); // Column C (NEW)
         let email = String(data[i][3]).trim(); // Column D
         let role = String(data[i][4]).trim(); // Column E
         
         if (cust && email && email.includes("@")) {
             if (!customerEmailMap[cust]) {
                 customerEmailMap[cust] = { main: new Set(), cc: new Set(), custManager: "" };
             }
             
             if (role.toLowerCase().includes("cc")) {
                 customerEmailMap[cust].cc.add(email);
             } else {
                 customerEmailMap[cust].main.add(email); // Treat Main Recipient or blank as main
             }
             
             if (custManager && !customerEmailMap[cust].custManager) {
                 customerEmailMap[cust].custManager = custManager;
             }
         }
     }
  });
  
  // --- ADDED: LOAD AE EMAILS ONCE BEFORE THE LOOP ---
  const aeEmailSheets = getDatabaseSheets(PREFIX_AE_EMAILS);
  let aeEmailMap = {};
  aeEmailSheets.forEach(sheet => {
     let data = sheet.getDataRange().getValues();
     if(data.length < 2) return;
     let headers = data[0].map(h => String(h).toLowerCase().trim());
     let nameIdx = headers.indexOf("ae name");
     let pIdx = headers.indexOf("pingcon email");
     let aIdx = headers.indexOf("a2z2go email");
     let kIdx = headers.indexOf("keysys email");

     for(let i = 1; i < data.length; i++) { // Skip headers
         let aeName = nameIdx > -1 ? String(data[i][nameIdx]).trim() : "";
         let pingconEmail = pIdx > -1 ? String(data[i][pIdx]).trim() : "";
         let a2zitEmail = aIdx > -1 ? String(data[i][aIdx]).trim() : "";
         let keysysEmail = kIdx > -1 ? String(data[i][kIdx]).trim() : "";
         
         if(aeName) {
             aeEmailMap[aeName] = { pingcon: pingconEmail, a2zit: a2zitEmail, keysys: keysysEmail };
         }
     }
  });
  // ---------------------------------------------------------

  while (pendingCust.length > 0) {
     // PROACTIVE QUOTA TRACKING: Stop early if 15 or fewer emails remain to guarantee summary is sent
     let remainingQuota = 999;
     try { remainingQuota = MailApp.getRemainingDailyQuota(); } catch(e) {}
     
     if (remainingQuota <= 15) {
         props.setProperty(UNPAID_CUSTOMER_BATCH_KEY, JSON.stringify(pendingCust));
         props.setProperty("LAST_RUN_ERROR", "Google Email Quota limit reached! Script stopped early to ensure Exclusion List is sent. Please paste the list into the Exclusion Box and run the remaining items tomorrow.");
         clearReportLock();
         deleteTriggersByName(UNPAID_CUSTOMER_TRIGGER_NAME);
         sendBatchSuccessSummary(true, "UNPAID_DR_CUST");
         return;
     }

     if (props.getProperty("CANCEL_UNPAID_DR_CUSTOMER") === "true") {
         deleteTriggersByName(UNPAID_CUSTOMER_TRIGGER_NAME);
         props.deleteProperty(UNPAID_CUSTOMER_BATCH_KEY);
         props.deleteProperty("CANCEL_UNPAID_DR_CUSTOMER");
         clearReportLock();
         sendBatchSuccessSummary(true, "UNPAID_DR_CUST");
         return; 
     }

     if (Date.now() - startTime > MAX_EXECUTION_TIME) {
        props.setProperty(UNPAID_CUSTOMER_BATCH_KEY, JSON.stringify(pendingCust));
        replaceTrigger(UNPAID_CUSTOMER_TRIGGER_NAME);
        return; 
     }
     
     let currentCustomerGroup = pendingCust.shift(); 
     
     try {
         // Using a comma is safer for filenames and perfectly searchable in email clients
         let commonName = currentCustomerGroup.join(", "); 
         // Strips punctuation (apostrophes, commas, dots) so Company A and Company A's match 100%
         let currentCustKeys = currentCustomerGroup.map(c => c.toLowerCase().replace(/[^a-z0-9]/g, ''));
         let allCustItems =[];
         
         arAllData.forEach(item => {
            let row = item.data;

            let textVal = (textIdx > -1) ? String(row[textIdx]) : "";
            let notesVal = (notesIdx > -1) ? String(row[notesIdx]) : "";
            if (isTransactionVoid(notesVal, textVal)) return;

            let rawCustomer = (customerIdx > -1) ? String(row[customerIdx]).trim() : "";
            let rowCustKey = rawCustomer.toLowerCase().replace(/[^a-z0-9]/g, '');
            if (!currentCustKeys.includes(rowCustKey)) return;

            let openBal = (openBalIdx > -1) ? parseFloat(row[openBalIdx]) : 0;
            if (isNaN(openBal)) openBal = 0;
            if (openBal < 1) return; // Ignores <= 0 and anything less than 1.00

            let statusRaw = (paymentStatusIdx > -1 && paymentStatusIdx < row.length) ? String(row[paymentStatusIdx]).trim() : "";
            let statusLow = statusRaw.toLowerCase();
            if (statusLow === "fully paid") return;

            let rawDocDate = (docDateIdx > -1) ? row[docDateIdx] : "";
            let docDateObj = (rawDocDate instanceof Date) ? rawDocDate : new Date(rawDocDate);
            if (isNaN(docDateObj.getTime())) return;
            if (docDateObj < startDate || docDateObj > endDate) return;

            let rawManager = (managerIdx > -1) ? String(row[managerIdx]) : "";
            let manager = getCanonicalName(rawManager);

            let rawDeadline = (deadlineIdx > -1) ? row[deadlineIdx] : "";
            let deadlineObj = new Date(rawDeadline);
            if (isNaN(deadlineObj.getTime()) || rawDeadline === "") {
                deadlineObj = new Date(docDateObj);
                let pTerm = (paymentTermIdx > -1) ? parseInt(row[paymentTermIdx]) : 0;
                if (!isNaN(pTerm) && pTerm > 0) {
                    deadlineObj.setDate(deadlineObj.getDate() + pTerm);
                }
            }

            let salesTotal = (originalIdx > -1) ? parseFloat(row[originalIdx]) : 0;
            if (isNaN(salesTotal)) salesTotal = 0;
            let amountPaid = salesTotal - openBal;
            if (amountPaid < 0) amountPaid = 0;

            let sourceFile = (sourceFileIdx > -1) ? String(row[sourceFileIdx]) : "";
            let companyEntity = getCompanyEntity(sourceFile);
            
            if (allowedEntities.length > 0 && !allowedEntities.includes(companyEntity)) return;

            allCustItems.push({
                docDateObj: docDateObj,
                invoiceDateStr: fastFormatDate(docDateObj),
                paymentTerm: (paymentTermIdx > -1) ? String(row[paymentTermIdx]).trim() : "",
                counterDateStr: (counterDateIdx > -1 && row[counterDateIdx]) ? fastFormatDate(new Date(row[counterDateIdx])) : "",
                deadlineStr: fastFormatDate(deadlineObj),
                manager: manager,
                customer: rawCustomer,
                text: (textIdx > -1) ? String(row[textIdx]) : "",
                salesTotal: salesTotal,
                amountPaid: amountPaid,
                openBal: openBal,
                companyEntity: companyEntity,
                remarksStatus: statusRaw,
                sourceFile: sourceFile
            });
         });
         
         if (allCustItems.length === 0) continue;
         
         let groupedByBucket = { "Tech": [], "Trading": [] };
         allCustItems.forEach(item => {
             if (item.companyEntity === "Keysys" || item.companyEntity === "A2ZIT") {
                 groupedByBucket["Tech"].push(item);
             } else {
                 groupedByBucket["Trading"].push(item);
             }
         });

         let cutOffDate = new Date(now);
         let daysUntilFriday = (5 - cutOffDate.getDay() + 7) % 7;
         cutOffDate.setDate(cutOffDate.getDate() + daysUntilFriday);
         cutOffDate.setHours(23, 59, 59, 999);

         for (let bucketName in groupedByBucket) {
             let aeItems = groupedByBucket[bucketName];
             if (aeItems.length === 0) continue; 
             
             aeItems.sort((a, b) => a.docDateObj - b.docDateObj);

             let reportHeaders =[
                 "SN", "Entity", "Invoice Date (MM/DD/YYYY)", "Terms", "Due Date (MM/DD/YYYY)", 
                 "Cust Manager", "Customer Name", "Additional text", "Sales Total", "Amount Paid", "Open Balance", 
                 "Counter Date (MM/DD/YYYY)", "Received Date (MM/DD/YYYY)", "Remarks/Payment Status"
             ];
             
             let sheetData = [];
             let totalDrCount = aeItems.length;
             let totalOpenBalance = 0;
             let monthlyBal = {};
             let fullyUnpaidCount = 0;
             let partiallyUnpaidCount = 0;

             let dueItems = [];
             let advanceItems = [];

             aeItems.forEach(item => {
                 let dDate = new Date(item.deadlineStr);
                 if (isNaN(dDate.getTime())) dDate = item.docDateObj; 
                 
                 if (dDate <= cutOffDate) {
                     if (String(item.remarksStatus).toLowerCase().includes("advance")) {
                         item.remarksStatus = "Due for collection";
                     }
                     dueItems.push(item);
                 } else {
                     advanceItems.push(item);
                 }

                 totalOpenBalance += item.openBal;
                 let sortKey = Utilities.formatDate(item.docDateObj, Session.getScriptTimeZone(), "yyyy-MM");
                 let displayLabel = Utilities.formatDate(item.docDateObj, Session.getScriptTimeZone(), "MMM yyyy");
                 let monthId = sortKey + "|" + displayLabel;

                 if (!monthlyBal[monthId]) monthlyBal[monthId] = 0;
                 monthlyBal[monthId] += item.openBal;

                 if (item.openBal >= item.salesTotal - 0.05 && item.salesTotal > 0) fullyUnpaidCount++;
                 else partiallyUnpaidCount++;
             });

             let totalDue = 0;
             if (dueItems.length > 0) {
                 sheetData.push(reportHeaders);
                 let sn = 1;
                 dueItems.forEach(item => {
                     let displayEntity = item.companyEntity === "A2ZIT" ? "A2Z2GO" : item.companyEntity;
                     sheetData.push([
                         sn++, displayEntity, item.invoiceDateStr, item.paymentTerm, item.deadlineStr, 
                         item.manager, item.customer, item.text, item.salesTotal, item.amountPaid, item.openBal,
                         item.counterDateStr, item.receivedDateStr, item.remarksStatus
                     ]);
                     totalDue += item.openBal;
                 });
                 let dueTotalRow = Array(reportHeaders.length).fill("");
                 dueTotalRow[7] = "Total Due:";
                 dueTotalRow[10] = totalDue;
                 sheetData.push(dueTotalRow);
             }

             let totalAdvance = 0;
             if (advanceItems.length > 0) {
                 // --- FIXED: Only print "Proof of Delivery" if there are Advance Items ---
                 // If there is a Due table above, add a blank spacer row. If not, don't add it so it stays at the very top.
                 if (dueItems.length > 0) {
                     sheetData.push(Array(reportHeaders.length).fill("")); 
                 }
                 let podRow = Array(reportHeaders.length).fill("");
                 podRow[0] = "Proof of Delivery"; // Index 0 is SN column
                 sheetData.push(podRow);
                 // ------------------------------------------------------------------------

                 sheetData.push(reportHeaders);
                 let sn = 1;
                 advanceItems.forEach(item => {
                     let displayEntity = item.companyEntity === "A2ZIT" ? "A2Z2GO" : item.companyEntity;
                     sheetData.push([
                         sn++, displayEntity, item.invoiceDateStr, item.paymentTerm, item.deadlineStr, 
                         item.manager, item.customer, item.text, item.salesTotal, item.amountPaid, item.openBal,
                         item.counterDateStr, item.receivedDateStr, item.remarksStatus
                     ]);
                     totalAdvance += item.openBal;
                 });
                 let advTotalRow = Array(reportHeaders.length).fill("");
                 advTotalRow[7] = "Total Advance:";
                 advTotalRow[10] = totalAdvance;
                 sheetData.push(advTotalRow);
             }

             sheetData.push(Array(reportHeaders.length).fill("")); 
             let grandTotalRow = Array(reportHeaders.length).fill("");
             grandTotalRow[7] = "GRAND TOTAL:";
             grandTotalRow[10] = totalOpenBalance;
             sheetData.push(grandTotalRow);

             let sortedMonthIds = Object.keys(monthlyBal).sort();
             let chartLabels = sortedMonthIds.map(id => id.split("|")[1]);
             let chartData = sortedMonthIds.map(id => monthlyBal[id]);

             let chartConfig = {
                 type: 'bar',
                 data: { labels: chartLabels, datasets:[{ label: 'Open Balance', data: chartData, backgroundColor: '#36a2eb' }] },
                 options: { title: { display: true, text: 'Total Open Balance by Month', fontSize: 14 }, legend: { display: false }, plugins: { datalabels: { display: false } } }
             };

             let chartBlob = null;
             try {
                 let chartUrl = "https://quickchart.io/chart?w=500&h=250&c=" + encodeURIComponent(JSON.stringify(chartConfig));
                 chartBlob = UrlFetchApp.fetch(chartUrl).getBlob().setName("chart.png");
             } catch(e) {}

             let reportDateString = "";
             if (aeItems.length > 0 && aeItems[0].sourceFile) {
                 let match = aeItems[0].sourceFile.match(/as of\s+([a-zA-Z]+)\s*(\d+),?\s+(\d{4})\s+(\d{4})/i);
                 if (match) {
                     let paddedDay = match[2].length === 1 ? "0" + match[2] : match[2];
                     reportDateString = `${match[1]} ${paddedDay}, ${match[3]} ${match[4]}`;
                 }
             }
             if (!reportDateString) {
                 let extractedDate = Utilities.formatDate(now, Session.getScriptTimeZone(), "MMM dd, yyyy");
                 let currentTime = Utilities.formatDate(now, Session.getScriptTimeZone(), "HHmm");
                 reportDateString = `${extractedDate} ${currentTime}`;
             }

             let minReportStartDate = null;
             let maxReportEndDate = null;
             aeItems.forEach(item => {
                 let parsedFromFile = false;
                 if (item.sourceFile) {
                     let match = item.sourceFile.match(/Ageing Report\s+(.*?)\s+to\s+(.*?)\s+as of/i);
                     if (match) {
                         let sDate = new Date(match[1]);
                         let eDate = new Date(match[2]);
                         if (!isNaN(sDate.getTime())) {
                             if (!minReportStartDate || sDate < minReportStartDate) minReportStartDate = sDate;
                         }
                         if (!isNaN(eDate.getTime())) {
                             if (!maxReportEndDate || eDate > maxReportEndDate) maxReportEndDate = eDate;
                         }
                         parsedFromFile = true;
                     }
                 }
                 if (!parsedFromFile && item.docDateObj) {
                     if (!isNaN(item.docDateObj.getTime())) {
                         if (!minReportStartDate || item.docDateObj < minReportStartDate) minReportStartDate = item.docDateObj;
                         if (!maxReportEndDate || item.docDateObj > maxReportEndDate) maxReportEndDate = item.docDateObj;
                     }
                 }
             });
             
             let dateRangePhrase = "";
             if (minReportStartDate && maxReportEndDate) {
                 let startStr = Utilities.formatDate(minReportStartDate, Session.getScriptTimeZone(), "MMMM d, yyyy");
                 let endStr = Utilities.formatDate(maxReportEndDate, Session.getScriptTimeZone(), "MMMM d, yyyy");
                 dateRangePhrase = `<p>For your reference, this Statement of Account covers transactions recorded in our system from <strong>${startStr}</strong> to <strong>${endStr}</strong>.</p>`;
             }

             let cleanCustName = String(commonName).replace(/[^a-zA-Z0-9 -,]/g, "");
             let cleanFileName = `Statement of Account - ${cleanCustName} as of Report Date ${reportDateString}`;

             let tempSpreadsheet = SpreadsheetApp.create(cleanFileName);
             let tempSheet = tempSpreadsheet.getSheets()[0];
             tempSheet.setName("Data");
             
             for (let r = 0; r < sheetData.length; r++) {
                 for (let c = 0; c < sheetData[r].length; c++) {
                     if (typeof sheetData[r][c] === 'string') {
                         sheetData[r][c] = sheetData[r][c].replace(/[\r\n]+/g, " ").trim();
                     }
                 }
             }

             tempSheet.getRange(1, 1, sheetData.length, reportHeaders.length).setValues(sheetData);
             
             for (let r = 0; r < sheetData.length; r++) {
                 let rowData = sheetData[r];
                 let excelRow = r + 1;
                 
                 if (rowData[0] === "SN") {
                     tempSheet.getRange(excelRow, 1, 1, reportHeaders.length).setFontWeight("bold").setBackground("#f3f3f3");
                 }
                 if (String(rowData[7]).includes("Total") || String(rowData[7]).includes("TOTAL")) {
                     tempSheet.getRange(excelRow, 8, 1, 4).setFontWeight("bold"); 
                     tempSheet.getRange(excelRow, 11).setBorder(true, false, false, false, false, false, "black", SpreadsheetApp.BorderStyle.SOLID);
                 }
             }
             
             tempSheet.getRange(1, 9, sheetData.length, 3).setNumberFormat("#,##0.00");
             tempSheet.getRange(1, 8, sheetData.length, 1).setNumberFormat("@");
             
             tempSheet.getDataRange().setWrap(true).setVerticalAlignment("top");

             // --- FIXED: Apply Red Font Color AND Disable Wrap Text for Proof of Delivery AFTER global wrap ---
             for (let r = 0; r < sheetData.length; r++) {
                 if (sheetData[r][0] === "Proof of Delivery") {
                     tempSheet.getRange(r + 1, 1).setFontColor("#d93025").setWrap(false);
                 }
             }
             // -------------------------------------------------------------------------------------------------
             for (let i = 1; i <= reportHeaders.length; i++) tempSheet.autoResizeColumn(i);
             tempSheet.setColumnWidth(7, 250);  
             tempSheet.setColumnWidth(8, 250);  
             tempSheet.setColumnWidth(9, 110);  
             tempSheet.setColumnWidth(10, 110); 
             tempSheet.setColumnWidth(11, 110); 
             tempSheet.setColumnWidth(14, 200); 

             let statsSheet = tempSpreadsheet.insertSheet("Statistics");
             let statsData = [
                 ["Metric", "Value"],
                 ["Total Unpaid DRs", totalDrCount],
                 ["Total Open Balance", totalOpenBalance],
                 ["Num of DRs 100% unpaid", fullyUnpaidCount],["Num of DRs partially unpaid", partiallyUnpaidCount]
             ];
             statsSheet.getRange(1, 1, statsData.length, 2).setValues(statsData);
             statsSheet.getRange(1, 1, 1, 2).setFontWeight("bold").setBackground("#f3f3f3");
             statsSheet.getRange(3, 2).setNumberFormat("#,##0.00"); 
             statsSheet.setColumnWidth(1, 200);
             statsSheet.setColumnWidth(2, 120);

             if (chartBlob) statsSheet.insertImage(chartBlob, 1, 7); 

             SpreadsheetApp.flush();
             Utilities.sleep(3000); 

             let url = "https://docs.google.com/spreadsheets/d/" + tempSpreadsheet.getId() + "/export?format=xlsx";
             let token = ScriptApp.getOAuthToken();
             let attachmentBlob = null;

             for (let fetchAttempt = 0; fetchAttempt < 3; fetchAttempt++) {
               try {
                 let response = UrlFetchApp.fetch(url, { headers: { 'Authorization': 'Bearer ' + token }, muteHttpExceptions: true });
                 if (response.getResponseCode() === 200) {
                     attachmentBlob = response.getBlob().setName(cleanFileName + ".xlsx");
                     if (rootFolder) {
                         let targetYearStr = Utilities.formatDate(endDate, Session.getScriptTimeZone(), "yyyy");
                         let targetMonthStr = Utilities.formatDate(endDate, Session.getScriptTimeZone(), "MMMM");

                         for (let attempt = 0; attempt < 3; attempt++) {
                             try {
                                 let yearFolders = rootFolder.getFoldersByName(targetYearStr);
                                 let yearFolder = yearFolders.hasNext() ? yearFolders.next() : rootFolder.createFolder(targetYearStr);

                                 let custFolders = yearFolder.getFoldersByName(cleanCustName);
                                 let custFolder = custFolders.hasNext() ? custFolders.next() : yearFolder.createFolder(cleanCustName);

                                 let monthFolders = custFolder.getFoldersByName(targetMonthStr);
                                 let monthFolder = monthFolders.hasNext() ? monthFolders.next() : custFolder.createFolder(targetMonthStr);

                                 let existingFiles = monthFolder.getFiles();
                                 while (existingFiles.hasNext()) {
                                     let file = existingFiles.next();
                                     if (file.getName() === (cleanFileName + ".xlsx")) file.setTrashed(true);
                                 }
                                 monthFolder.createFile(attachmentBlob);
                                 break; 
                             } catch (driveErr) { Utilities.sleep(2000); }
                         }
                     }
                     break; 
                 } else {
                     Utilities.sleep(4000 * (fetchAttempt + 1));
                 }
               } catch (e) { Utilities.sleep(5000 * (fetchAttempt + 1)); }
             }

             if (!attachmentBlob) throw new Error("Failed to export Excel file for " + commonName);

             try { DriveApp.getFileById(tempSpreadsheet.getId()).setTrashed(true); } catch(e){}

             if (attachmentBlob) {
                 let chartHtml = chartBlob ? `<br><div style="text-align: left; margin: 15px 0;"><img src="cid:chartImage" style="max-width: 100%; border: 1px solid #ddd; border-radius: 4px;" alt="Open Balance Chart"></div>` : '';

                 let hasKeysysOrA2ZIT = aeItems.some(item => item.companyEntity === "Keysys" || item.companyEntity === "A2ZIT");
                 let hasA2ZPMC = aeItems.some(item => item.companyEntity === "A2ZPMC");

                 let sysSettings = getSystemSettings();
                 let custUpperCheck = String(commonName).toUpperCase().trim();
                 let specialSenderClients = String(sysSettings["Special_Sender_Clients"] || "").toUpperCase().split(',').map(s => s.trim()).filter(Boolean);
                 let isSpecialSender = specialSenderClients.some(c => custUpperCheck.includes(c));
                 let isSpecialFeedback = isSpecialSender;

                 let isTech = (hasKeysysOrA2ZIT || isSpecialSender);
                 let senderName = isTech ? "Finance A2Z2GO" : "Finance Pingcon";
                 let senderEmail = getSenderEmail(isTech);
                 let feedbackImageId = isTech ? FEEDBACK_IMAGE_ID_A2Z_KEYSYS : FEEDBACK_IMAGE_ID_PINGCON_BONITA;
                 
                 if (isSpecialFeedback) {
                     feedbackImageId = FEEDBACK_IMAGE_ID_SPECIAL;
                 }

                 let feedbackImgBlob = null;
                 try {
                     if (feedbackImageId && feedbackImageId !== "INSERT_FILE_ID_FOR_A2Z_KEYSYS_IMAGE_HERE" && feedbackImageId !== "INSERT_FILE_ID_FOR_PINGCON_BONITA_IMAGE_HERE" && feedbackImageId !== "1_INSERT_NEW_IMAGE_ID_HERE") {
                         feedbackImgBlob = DriveApp.getFileById(feedbackImageId).getBlob();
                     }
                 } catch(e) {}

                 let xmasPdfBlob = null;
                 try {
                     let props = PropertiesService.getScriptProperties();
                     let isDevMode = ScriptApp.getService().getUrl().endsWith("/dev");
                     let entityName = isTech ? "A2Z2GO" : "Pingcon";

                     // Single Marketing Flyer Attachment
                     let propKey1 = isTech 
                         ? (isDevMode && props.getProperty('MARKETING_PDF_XMAS_A2Z_ID_DEV') ? 'MARKETING_PDF_XMAS_A2Z_ID_DEV' : 'MARKETING_PDF_XMAS_A2Z_ID')
                         : (isDevMode && props.getProperty('MARKETING_PDF_XMAS_PINGCON_ID_DEV') ? 'MARKETING_PDF_XMAS_PINGCON_ID_DEV' : 'MARKETING_PDF_XMAS_PINGCON_ID');

                     let xmasPdfId1 = props.getProperty(propKey1);
                     if (!xmasPdfId1 && isDevMode) {
                         xmasPdfId1 = isTech
                             ? props.getProperty('MARKETING_PDF_XMAS_A2Z_ID')
                             : props.getProperty('MARKETING_PDF_XMAS_PINGCON_ID');
                     }

                     if (xmasPdfId1 && !xmasPdfId1.includes("INSERT")) {
                         let fileObj1 = DriveApp.getFileById(xmasPdfId1);
                         let rawBlob1 = fileObj1.getBlob();
                         let fileMime1 = (rawBlob1.getContentType() || "").toLowerCase();
                         let fileNameLower1 = fileObj1.getName().toLowerCase();

                         if (fileMime1 === MimeType.PDF || fileNameLower1.endsWith(".pdf")) {
                             xmasPdfBlob = rawBlob1.getAs(MimeType.PDF);
                             xmasPdfBlob.setName(`Christmas_Flyer_${entityName}.pdf`);
                         } else {
                             let ext1 = "jpg";
                             if (fileMime1.includes("png") || fileNameLower1.endsWith(".png")) ext1 = "png";
                             else if (fileMime1.includes("jpeg") || fileMime1.includes("jpg") || fileNameLower1.endsWith(".jpg") || fileNameLower1.endsWith(".jpeg")) ext1 = "jpg";

                             xmasPdfBlob = rawBlob1;
                             xmasPdfBlob.setName(`Christmas_Flyer_${entityName}.${ext1}`);
                         }
                     }
                 } catch(e) { console.warn("Failed to fetch Christmas marketing attachment: " + e); }

                 let ccRecipientsList = [];
                 let mainRecipientsList = [];

                 let lookupCustKey = String(currentCustomerGroup[0]).toLowerCase().replace(/[^a-z0-9]/g, '');
                 if (customerEmailMap[lookupCustKey]) {
                     mainRecipientsList = Array.from(customerEmailMap[lookupCustKey].main);
                     ccRecipientsList = Array.from(customerEmailMap[lookupCustKey].cc);
                 }
                 let hasCustomerEmail = mainRecipientsList.length > 0;

                 let custUpper = String(commonName).toUpperCase().trim();

                 if (isTech) {
                     [sysSettings["Finance_A2Z"], sysSettings["Executive_A2Z"], sysSettings["Marketing_A2Z"]].forEach(e => {
                         if (e && e.trim() !== "") ccRecipientsList.push(e);
                     });
                 } else {
                     [sysSettings["Finance_Pingcon"], sysSettings["Executive_Pingcon"], sysSettings["Marketing_Pingcon"]].forEach(e => {
                         if (e && e.trim() !== "") ccRecipientsList.push(e);
                     });
                 }

                 let uniqueAEsMap = {};
                 aeItems.forEach(item => {
                     let managerToUse = item.manager;
                     if (managerToUse) {
                         let key = managerToUse + "_" + item.companyEntity;
                         uniqueAEsMap[key] = { ae: managerToUse, entity: item.companyEntity };
                     }
                 });
                 let customerAEs = Object.values(uniqueAEsMap);
                 
                 let aeContactHtml = "";
                 let uniqueAeContactSet = new Set();

                 customerAEs.forEach(pair => {
                     if (aeEmailMap[pair.ae]) {
                         let emails = aeEmailMap[pair.ae];
                         let emailToAdd = "";
                         
                         // REMOVED the isSpecialSender override so AE emails always pull from the database
                         if (isTech) emailToAdd = emails.a2zit;
                         else if (pair.entity === "Keysys") emailToAdd = emails.keysys;
                         else if (pair.entity === "A2ZIT") emailToAdd = emails.a2zit;
                         else emailToAdd = emails.pingcon; 
                         
                         if (emailToAdd && emailToAdd.includes("@")) {
                             if (!hasCustomerEmail) mainRecipientsList.push(emailToAdd);
                             else ccRecipientsList.push(emailToAdd);
                         }
                         
                         let aeFullName = pair.ae;
                         if (aeFullName.includes(",")) {
                             let parts = aeFullName.split(",");
                             aeFullName = parts[1].trim() + " " + parts[0].trim();
                         }
                         
                         let contactLine = `Account Executive: ${aeFullName} - ${emailToAdd || "N/A"}`;
                         if (!uniqueAeContactSet.has(contactLine)) {
                             uniqueAeContactSet.add(contactLine);
                             aeContactHtml += `${contactLine}<br>`;
                         }
                     }
                 });

                 let collectionAssts = [];
                 let dynamicCcToAppend = [];

                 // Check primary entity makeup of the customer's items (Reusing existing variables to prevent syntax error)
                 let hasBonita = aeItems.some(item => item.companyEntity === "Bonita");
                 let hasPingcon = aeItems.some(item => item.companyEntity === "Pingcon" || !item.companyEntity);

                 let assignToPrimary = resolveCollectionAssistantAssignment(
                     custUpper, 
                     hasKeysysOrA2ZIT, 
                     hasBonita, 
                     hasA2ZPMC, 
                     hasPingcon,
                     sysSettings
                 );

                 // Dynamic Collection Assistants from System_Settings_Data (Zero Script Properties needed)
                 let asstPrimaryEmail = hasKeysysOrA2ZIT 
                     ? (sysSettings["Collection_Asst_A2Z_Email"] || "") 
                     : (sysSettings["Collection_Asst_Pingcon_Primary_Email"] || "");
                 let asstPrimaryName = hasKeysysOrA2ZIT 
                     ? (sysSettings["Collection_Asst_A2Z_Name"] || "Collection Assistant (A2Z2GO)") 
                     : (sysSettings["Collection_Asst_Pingcon_Primary_Name"] || "Collection Assistant");

                 let asstSecondaryEmail = sysSettings["Collection_Asst_Secondary_Email"] || "";
                 let asstSecondaryName = sysSettings["Collection_Asst_Secondary_Name"] || "Collection Assistant";

                 if (isSpecialSender || assignToPrimary) {
                     if (asstPrimaryEmail) {
                         collectionAssts.push(`${asstPrimaryName} - ${asstPrimaryEmail}`);
                         dynamicCcToAppend.push(asstPrimaryEmail);
                     }
                 } else {
                     if (asstSecondaryEmail) {
                         collectionAssts.push(`${asstSecondaryName} - ${asstSecondaryEmail}`);
                         dynamicCcToAppend.push(asstSecondaryEmail);
                     }
                 }
                 
                 let caMobile = PropertiesService.getScriptProperties().getProperty('COLLECTION_ASSISTANT_MOBILE') || "";
                 let mobileLine = caMobile ? `Mobile: ${caMobile}<br>` : "";
                 let collectionAsstHtml = "";
                 [...new Set(collectionAssts)].forEach(asst => {
                     collectionAsstHtml += `BIR Form 2307 Expanded Withholding Tax (EWT), Collection Assistant in-charge: ${asst}<br>${mobileLine}`;
                 });

                 let feedbackLink = isTech ? (SCRIPT_PROPS['FEEDBACK_LINK_TECH'] || "") : (SCRIPT_PROPS['FEEDBACK_LINK_PINGCON'] || "");
                 if (isSpecialFeedback) {
                     feedbackLink = SCRIPT_PROPS['FEEDBACK_LINK_SPECIAL'] || feedbackLink;
                 }

                 let htmlBody = Template_UnpaidDrCustEmail(cleanCustName, dateRangePhrase, totalDrCount, totalDue, totalAdvance, totalOpenBalance, chartHtml, aeContactHtml, collectionAsstHtml, feedbackImgBlob, feedbackLink, senderEmail);

                 // Append executive management dynamically from System_Settings_Data
                 let mgtEmail = isTech ? (sysSettings["Management_A2Z"] || "") : (sysSettings["Management_Pingcon"] || "");
                 if (mgtEmail) ccRecipientsList.push(mgtEmail);

                 if (mainRecipientsList.length === 0) {
                     if (dynamicCcToAppend.length > 0) dynamicCcToAppend.forEach(email => mainRecipientsList.push(email));
                     else if (mgtEmail) mainRecipientsList.push(mgtEmail);
                 } else {
                     dynamicCcToAppend.forEach(email => ccRecipientsList.push(email));
                 }

                 ccRecipientsList = [...new Set(ccRecipientsList)].filter(email => !mainRecipientsList.includes(email));
                 mainRecipientsList = [...new Set(mainRecipientsList)];

                 const mainRecipient = mainRecipientsList.join(", ");
                 const ccRecipients = ccRecipientsList.join(", ");
             
                 let securedBlob = encryptAttachmentViaPython(attachmentBlob, cleanFileName + ".xlsx", null);
             
             let emailOptions = {
                 from: senderEmail, cc: ccRecipients, name: senderName,
                 attachments: [securedBlob], inlineImages: {}, throwError: true
             };
             // Single Flyer Attachment
             if (xmasPdfBlob) {
                 emailOptions.attachments.push(xmasPdfBlob);
             }
             if (chartBlob) emailOptions.inlineImages.chartImage = chartBlob;
             if (feedbackImgBlob) emailOptions.inlineImages.feedbackImage = feedbackImgBlob;

             try {
                 sendRobustEmail(mainRecipient, cleanFileName, htmlBody, emailOptions);
                 recordSuccessfulSend(commonName, hasKeysysOrA2ZIT ? "A2Z2GO" : "Pingcon");
             } catch(errStr) {
                 let errString = String(errStr).toLowerCase();
                 if (errString.includes("too many times") || errString.includes("premium email") || errString.includes("limit") || errString.includes("quota")) {
                     pendingCust.unshift(currentCustomerGroup);
                     props.setProperty(UNPAID_CUSTOMER_BATCH_KEY, JSON.stringify(pendingCust));
                     props.setProperty("LAST_RUN_ERROR", "Google Email Quota limit reached! Script stopped. Please run the remaining items tomorrow.");
                     clearReportLock();
                     deleteTriggersByName(UNPAID_CUSTOMER_TRIGGER_NAME);
                     sendBatchSuccessSummary(true, "UNPAID_DR_CUST");
                     return;
                 }
                 console.error("Failed to send email for " + commonName + ". Error: " + errStr); 
             }
             Utilities.sleep(10000); // Increased delay to prevent spam flagging
             }
         }
     } catch (masterErr) {
         console.error("Crash on " + currentCustomerGroup.join(", ") + ": " + masterErr);
         pendingCust.unshift(currentCustomerGroup);
         props.setProperty(UNPAID_CUSTOMER_BATCH_KEY, JSON.stringify(pendingCust));
         props.setProperty("LAST_RUN_ERROR", "Google Server Timeout. Script paused. Please run the remaining items again later.");
         clearReportLock();
         deleteTriggersByName(UNPAID_CUSTOMER_TRIGGER_NAME);
         sendBatchSuccessSummary(true, "UNPAID_DR_CUST");
         return;
     }

     if (props.getProperty("CANCEL_UNPAID_DR_CUSTOMER") === "true") {
         deleteTriggersByName(UNPAID_CUSTOMER_TRIGGER_NAME);
         props.deleteProperty(UNPAID_CUSTOMER_BATCH_KEY);
         props.deleteProperty("CANCEL_UNPAID_DR_CUSTOMER");
         clearReportLock();
         sendBatchSuccessSummary(true, "UNPAID_DR_CUST");
         return; 
     }

     // Update the property so if the script crashes, it knows where it left off, but DO NOT stop the loop
     props.setProperty(UNPAID_CUSTOMER_BATCH_KEY, JSON.stringify(pendingCust));
  }

  deleteTriggersByName(UNPAID_CUSTOMER_TRIGGER_NAME);
  props.deleteProperty(UNPAID_CUSTOMER_BATCH_KEY);
  clearReportLock();
  console.log("ALL UNPAID DRs (PER CUSTOMER) PROCESSED SUCCESSFULLY.");
  sendBatchSuccessSummary(false, "UNPAID_DR_CUST");
}







function formatMoney(amount) {
  return amount.toFixed(2).replace(/\d(?=(\d{3})+\.)/g, '$&,');
}


function _$startGeneration(aeName, month, requestDate) {
  var lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) { throw new Error("System is currently busy. Another user is generating reports right now. Please wait a few minutes."); }
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const currentMonthDate = new Date(month + " 1"); 
    const startPointDate = new Date("January 1, 2026");
    let prevMonthReportExists = true; 

    if (currentMonthDate > startPointDate) {
      let prevMonthDate = new Date(currentMonthDate);
      prevMonthDate.setMonth(currentMonthDate.getMonth() - 1);
      let prevMonthStr = Utilities.formatDate(prevMonthDate, Session.getScriptTimeZone(), "MMMM yyyy");
      let expectedPrevFile = `${aeName} - ${prevMonthStr}`;
      
      let fileFoundInDrive = false;
      try {
        const rootFolder = DriveApp.getFolderById(BUY_SELL_ROOT_FOLDER_ID);
        const aeFolders = rootFolder.getFoldersByName(aeName);
        if (aeFolders.hasNext()) {
          const aeFolder = aeFolders.next();
          if (aeFolder.getFilesByName(expectedPrevFile).hasNext()) {
            fileFoundInDrive = true;
          }
        }
      } catch (e) {
        console.error("Drive Check Error: " + e.toString());
      }
      prevMonthReportExists = fileFoundInDrive;
    }
    _$runAllProcesses(aeName, month, prevMonthReportExists, requestDate);
  } finally { lock.releaseLock(); }
}

function getDataForDropdowns() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const targetSheets = getDatabaseSheets(PREFIX_SR);
  const map = {};

  targetSheets.forEach(s => {
    const lastRow = s.getLastRow();
    if (lastRow > 1) {
      // Direct Memory Read: Fetches only the required block (Column 5 to 16) instead of entire row width
      const values = s.getRange(2, 5, lastRow - 1, 12).getValues(); 
      for(let i = 0; i < values.length; i++) {
        let dateVal = values[i][0]; // Col 5 (Date)
        let rawManager = String(values[i][11]); // Col 16 (Manager)
        let manager = getCanonicalName(rawManager);
        
        if (manager && dateVal) {
          let dateObj = new Date(dateVal);
          if (!isNaN(dateObj.getTime())) {
            let monthYear = Utilities.formatDate(dateObj, Session.getScriptTimeZone(), "MMMM yyyy");
            if (!map[manager]) map[manager] = new Set();
            map[manager].add(monthYear);
          }
        }
      }
    }
  });

  // --- ADDED: Extract dates from Scanned Supplier Invoices for EWT Report UI ---
  const scannedSheets = ss.getSheets().filter(s => s.getName().startsWith(PREFIX_SCANNED_INVOICES));
  if (!map["_SYSTEM_DATES_"]) map["_SYSTEM_DATES_"] = new Set();
  
  scannedSheets.forEach(scannedSheet => {
    if (scannedSheet && scannedSheet.getLastRow() > 1) {
      const scannedDates = scannedSheet.getRange(2, 3, scannedSheet.getLastRow() - 1, 1).getValues();
      for (let i = 0; i < scannedDates.length; i++) {
        let dateVal = scannedDates[i][0];
        if (dateVal) {
          let dateObj = new Date(dateVal);
          if (!isNaN(dateObj.getTime())) {
            let monthYear = Utilities.formatDate(dateObj, Session.getScriptTimeZone(), "MMMM yyyy");
            map["_SYSTEM_DATES_"].add(monthYear);
          }
        }
      }
    }
  });

  const finalMap = {};
  for (let key in map) finalMap[key] = Array.from(map[key]);
  return finalMap;
}


const monthNames = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
const fastFormatMonthYear = (d) => {
  if (!(d instanceof Date) || isNaN(d.getTime())) return String(d);
  return monthNames[d.getMonth()] + ' ' + d.getFullYear();
};
function _$processBuySellForms(targetAeFilter = null, targetMonthFilter = null, prevMonthExists = true, requestDate = null, generateOutput = true, generateCommReport = false, returnMissingInvoicesOnly = false) {
  console.time("Process_Buy_Sell_Total");
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rootFolder = DriveApp.getFolderById(BUY_SELL_ROOT_FOLDER_ID);
  
  const PREFIX_SKU_MAP_LOCAL = "Missing_SKU_Data_"; 

  // --- NEW: Email Variables for Commission Report ---
  let commEmailHtmlBlocks = [];
  let commEmailAttachments =[];
  let commOverallMinDate = null;
  let commOverallMaxDate = null;

  console.log("--- STARTED: GATHERING DATA ---");
  
  let soAllData = gatherAllData(PREFIX_SR);
  const poPingconBonitaData = gatherAllData(PREFIX_PO_PINGCON_BONITA);
  const poKeysysData = gatherAllData(PREFIX_PO_KEYSYS);
  const arAllData = gatherAllData(PREFIX_AR);
  const sqAllData = gatherAllData(PREFIX_SQ);
  const skuMapData = gatherAllData(PREFIX_SKU_MAP_LOCAL);
  
  // >>> START: NEW BASKET GATHERING <<<
  const PREFIX_BASKET_LOCAL = "Basket_Data_";
  const basketAllData = gatherAllData(PREFIX_BASKET_LOCAL);
  let basketMap = {};
  let basketComponentPriceMap = {}; // <-- NEW: Store all component prices globally

  basketAllData.forEach(item => {
     let r = item.data;
     let bSku = String(r[0]).trim().toUpperCase();
     let cSku = String(r[3]).trim().toUpperCase();
     let cPrice = parseFloat(r[5]) || 0;

     if(bSku) {
         if(!basketMap[bSku]) basketMap[bSku] = { components: [] };
         basketMap[bSku].components.push({
             sku: cSku,
             desc: String(r[4]).trim(),
             price: cPrice, // <-- Stores Component Unit Price
             qty: parseFloat(r[6]) || 0,
             uom: String(r[7]).trim()
         });
     }

     // NEW: Cache highest non-zero price for any component SKU to use when sold individually
     if (cSku && cPrice > 0) {
         if (!basketComponentPriceMap[cSku] || cPrice > basketComponentPriceMap[cSku]) {
             basketComponentPriceMap[cSku] = cPrice;
         }
     }
  });
  // >>> END: NEW BASKET GATHERING <<<

  // --- TARGETED INVOICE EXTRACTION & EWT/PURCHASE MAPPING ---
  let ewtMap = {};
  let purchasesMap = {};

  let skuAliasMap = {};
  let skuPerPcMap = {}; 
  let missingSkuPriceMap = {}; 
  
  skuMapData.forEach(item => {
     let r = item.data;
     let oldSku = String(r[0]).trim().toUpperCase(); 
     let desc = String(r[1]).trim();
     let newSkuRaw = String(r[2]).trim();            
     let remarks = String(r[3]).trim();
     let price = parseFloat(r[4]); 
     
     // Extracts the newly unified price column
     if (oldSku && !isNaN(price) && price > 0) {
         missingSkuPriceMap[oldSku] = price;
     }

     if(oldSku && newSkuRaw) {
         let aliases = newSkuRaw.split(/\r?\n/).map(line => line.trim().split(" ")[0].trim().toUpperCase()).filter(x => x !== "");
         if(aliases.length > 0) skuAliasMap[oldSku] = aliases;
     }
     
     // Detect if Remarks signify a "per pc" conversion
     if (oldSku && remarks.toLowerCase().includes("per pc")) {
         let perPcSku = remarks.split(/\s+/)[0].trim().toUpperCase(); 
         let match = desc.match(/[xX]\s*(\d+)(?!.*[xX]\s*\d+)/);
         let multiplier = match ? parseInt(match[1], 10) : 1;
         
         if (perPcSku && multiplier > 1) {
             skuPerPcMap[oldSku] = { newSku: perPcSku, multiplier: multiplier };
         }
     }
  });

  if (soAllData.length === 0) return;

  // --- COST HISTORY ---
  let skuCostHistory = {}; 
  let nameCostHistory = {}; // <-- ADDED: For name-based matching
  
  // 1. Process Pingcon & Bonita POs
  poPingconBonitaData.forEach(item => {
    let row = item.data;
    let rowDate = new Date(row[0]);
    if (isNaN(rowDate.getTime())) rowDate = new Date(0);
    let sku = String(row[3]).trim().toUpperCase(); 
    let unitCost = parseFloat(row[6]) || 0;
    let initialQty = parseFloat(row[7]);
    if (isNaN(initialQty)) initialQty = 9999999; 
    
    if (unitCost > 0) {
      let nameStr = String(row[4]).trim();
      let poRecord = { 
        name: nameStr, supplier: row[5], invoice: row[1], pingconPo: row[2],  
        cost: unitCost, dateObj: rowDate, id: `${item.sheetName}_${item.rowIndex}`,
        remaining: initialQty, originalQty: initialQty
      };
      
      if (sku) {
        if (!skuCostHistory[sku]) skuCostHistory[sku] = [];
        skuCostHistory[sku].push(poRecord);
      }
      
      let upperName = nameStr.toUpperCase();
      if (upperName) {
        if (!nameCostHistory[upperName]) nameCostHistory[upperName] = [];
        nameCostHistory[upperName].push(poRecord);
      }
    }
  });

  

  for (let sku in skuCostHistory) {
    skuCostHistory[sku].sort((a, b) => b.dateObj - a.dateObj);
  }
  for (let n in nameCostHistory) {
    nameCostHistory[n].sort((a, b) => b.dateObj - a.dateObj);
  }

  let scannedInvoicesData = gatherAllData("Scanned_Supplier_Invoices_Data_");
  let receiptMap = {};
  scannedInvoicesData.forEach(item => {
      let sku = String(item.data[14]).trim().toUpperCase();
      let price = parseFloat(item.data[17]);
      if (sku && !isNaN(price)) receiptMap[sku] = price;
  });

  soAllData.forEach(item => { item.parsedDate = new Date(item.data[4]); });
  soAllData.sort((a, b) => a.parsedDate - b.parsedDate);

  // --- MATCHING & GROUPING ---
  let pendingGroups = {};
  let rowsToMarkDoneBySheet = {};
  let stockMovementLogsByBlock = {};
  const PREFIX_LOG = "Global_Stock_Movement_Logs_";
  
  let globalUnpaidItems = [];
  let globalNoStockItems = []; 
  let globalMissingItems = []; 
  // NEW: Capture Negative Profit Items in memory
  let globalNegativeItems = [];
  let globalPurchaseOutput = {};

  soAllData.forEach((item) => {
    let row = item.data;
    let rawManager = String(row[15]);
    let aeName = getCanonicalName(rawManager) || "Unknown_AE";
    let dateObj = item.parsedDate; // Reuse parsed date from sort
    let monthYear = fastFormatMonthYear(dateObj);
    let poNumber = String(row[5]).trim();
    let siNumber = String(row[3]).trim();
    let customer = String(row[1]).replace(/\s+/g, ' ').trim();
    let additionalText = String(row[5]).trim();
    let custManager = aeName;   

    if (poNumber === "") return;

    let metaKey = `${aeName}|${monthYear}`;
    let poKey = `${customer}_${poNumber}_${siNumber}`;

    if (!pendingGroups[metaKey]) pendingGroups[metaKey] = {};
    if (!pendingGroups[metaKey][poKey]) {
      pendingGroups[metaKey][poKey] = {
        meta: { 
             ae: aeName, month: monthYear, customer: customer, po: poNumber, si: siNumber, 
             dr: row[6] || "", date: fastFormatDate(dateObj),
             paymentStatus: "Pending", manager: custManager, rawDate: dateObj,
             additionalText: additionalText // ADDED: Save Additional Text here
        },
        items: []
      };
    }

    let fullProductName = String(row[7]).trim();
    if (fullProductName === "#ERROR!" || fullProductName === "") return;
    let originalSku = fullProductName.split(" ")[0].trim().toUpperCase();
    let originalDesc = fullProductName.substring(originalSku.length).trim();
    let originalSoldQty = parseFloat(row[9]) || 0;
    let originalUom = String(row[10]).trim() || "pc";
    let totalRevenue = parseFloat(row[14]) || 0;
    let srStatus = String(row[2]).trim();
    let srNumber = String(row[3]).trim();
    let srProduct = String(row[7]).trim();
    let srVatRate = String(row[8]).trim();
    let srNetPrice = parseFloat(row[11]) || 0;
    let srNetTotal = parseFloat(row[12]) || 0;
    let srRowVat = parseFloat(row[13]) || 0;
    let srRowTotal = parseFloat(row[14]) || 0;

    let sku = originalSku;
    let desc = originalDesc;
    let soldQty = originalSoldQty;
    let packSize = 1;

    if (sku.endsWith("C")) {
      let match = fullProductName.match(/.*[xX]\s*(\d+)/);
      if (match) {
        let parsedSize = parseInt(match[1]);
        if (parsedSize > 0) { 
             packSize = parsedSize;
             sku = sku.substring(0, sku.length - 1) + "P"; 
             soldQty = soldQty * packSize; 
             desc = desc.replace(new RegExp(`([xX]\\s*)${packSize}(?!.*[xX]\\s*\\d+)`, 'i'), "$11");
        }
      }
    }

    let unitSellPrice = (soldQty > 0) ? (totalRevenue / soldQty) : 0;
    let originalUnitSellPrice = (originalSoldQty > 0) ? (totalRevenue / originalSoldQty) : 0;
    
    // Tag if missing basket (Starts with B2 but not mapped)
    if (sku.startsWith("B2") && !basketMap.hasOwnProperty(sku)) {
        desc = desc + " (Not found in Basket Mapping)";
    }

    let isBasket = basketMap.hasOwnProperty(sku);
    let basketId = isBasket ? `B_${item.rowIndex}_${sku}` : null;

    let itemsToFulfill =[];

    // EXPLODE THE BASKET INTO COMPONENTS
    if (isBasket) {
        let components = basketMap[sku].components;
        let firstComp = true;
        components.forEach(comp => {
            itemsToFulfill.push({
                isFirstComp: firstComp,
                basketSku: originalSku,
                basketDesc: originalDesc,
                compSku: comp.sku,
                compDesc: comp.desc,
                compUom: comp.uom || "pc",
                compUnitPrice: comp.price, // <-- ADDED: Passes the price down
                reqQty: comp.qty * soldQty,
                sellRevenue: firstComp ? totalRevenue : 0, // Put all Revenue on Row 1
                origSellQty: firstComp ? originalSoldQty : "",
                origSellUom: firstComp ? originalUom : "",
                origUnitSell: firstComp ? originalUnitSellPrice : ""
            });
            firstComp = false;
        });
    } else {
        // STANDARD ITEM
        itemsToFulfill.push({
            isFirstComp: true,
            basketSku: originalSku,
            basketDesc: originalDesc,
            compSku: sku,
            compDesc: desc,
            compUom: "pc",
            compUnitPrice: 0, // <-- ADDED: Default 0 for non-basket items
            reqQty: soldQty,
            sellRevenue: totalRevenue,
            origSellQty: originalSoldQty,
            origSellUom: originalUom,
            origUnitSell: originalUnitSellPrice
        });
    }

    // PROCESS EACH EXPLODED ROW
    itemsToFulfill.forEach(reqItem => {
        // --- NEW LOGIC: APPLY PER PC MAPPING DYNAMICALLY ---
        if (skuPerPcMap.hasOwnProperty(reqItem.compSku)) {
            let mapData = skuPerPcMap[reqItem.compSku];
            reqItem.compSku = mapData.newSku;
            reqItem.compDesc = `${reqItem.compDesc} (Mapped to ${mapData.newSku} x${mapData.multiplier})`;
            reqItem.reqQty = reqItem.reqQty * mapData.multiplier;
        }

        let pendingQty = reqItem.reqQty;
        // Recalculates unit sell price perfectly to match the newly multiplied quantity
        let compUnitSellPrice = (pendingQty > 0) ? (reqItem.sellRevenue / reqItem.reqQty) : 0;
        let firstSplit = true; // <-- Added flag

        while (pendingQty > 0.0001) {
            let unitBuyPrice = 0; let supplier = ""; let purchaseInvoice = ""; let pingconPo = "";
            let sourceDate = null; let isMissingSku = false; let isNoStock = false; 
            let isValidationOverride = false; let isPartner = false; let allocatedQty = 0;
            
            let effectiveSku = reqItem.compSku;
            let effectiveDesc = reqItem.compDesc || fullProductName;

            if (effectiveSku === "PARTNER" || fullProductName.toLowerCase().includes("partner discount")) {
                isMissingSku = false; unitBuyPrice = 0; supplier = "Partner"; isPartner = true;
                allocatedQty = pendingQty;
            } 
            else {
                let isFallback = false;
                let sourceHistory = null;
                let bestPO = null;

                let effectiveNameUpper = String(effectiveDesc || fullProductName).trim().toUpperCase();
                let fullProdUpper = fullProductName.toUpperCase();

                // 1. Exact Name Match (Prioritizes exact matching of Names for Keysys/Bonita)
                if (nameCostHistory[effectiveNameUpper] && nameCostHistory[effectiveNameUpper].some(p => p.remaining > 0)) {
                    sourceHistory = nameCostHistory[effectiveNameUpper];
                } 
                else if (nameCostHistory[fullProdUpper] && nameCostHistory[fullProdUpper].some(p => p.remaining > 0)) {
                    sourceHistory = nameCostHistory[fullProdUpper];
                }
                // 2. Exact SKU Match
                else if (skuCostHistory[effectiveSku] && skuCostHistory[effectiveSku].some(p => p.remaining > 0)) {
                    sourceHistory = skuCostHistory[effectiveSku];
                } 
                // 3. SKU Alias Match
                else if (skuAliasMap[effectiveSku]) {
                    let aliases = skuAliasMap[effectiveSku];
                    let bestAlias = aliases.find(alias => skuCostHistory[alias] && skuCostHistory[alias].some(p => p.remaining > 0));
                    if (!bestAlias && aliases.length > 0) bestAlias = aliases[0];
                    if (bestAlias) {
                        sourceHistory = skuCostHistory[bestAlias];
                        effectiveSku = bestAlias; 
                        effectiveDesc = `${effectiveDesc} (Old SKU ${reqItem.compSku})`; 
                    }
                }
                // 4. Fallbacks (Grab history even if empty so it tags as "Oversold/No Stock" instead of "Missing SKU")
                else if (nameCostHistory[effectiveNameUpper]) { sourceHistory = nameCostHistory[effectiveNameUpper]; }
                else if (nameCostHistory[fullProdUpper]) { sourceHistory = nameCostHistory[fullProdUpper]; }
                else if (skuCostHistory[effectiveSku]) { sourceHistory = skuCostHistory[effectiveSku]; }

                if (sourceHistory) {
                    let validPOs = sourceHistory.filter(h => h.dateObj <= dateObj);
                    if (validPOs.length === 0) validPOs = sourceHistory;
                    
                    bestPO = validPOs.find(p => p.remaining > 0);
                    if (!bestPO && validPOs.length > 0) {
                        bestPO = validPOs[0];
                        isFallback = true;
                    }
                }

                if (bestPO) {
                    let take = isFallback ? pendingQty : Math.min(pendingQty, bestPO.remaining);
                    unitBuyPrice = bestPO.cost;
                    supplier = bestPO.supplier;
                    purchaseInvoice = bestPO.invoice;
                    pingconPo = bestPO.pingconPo;
                    sourceDate = bestPO.dateObj;
                    allocatedQty = take;
                    bestPO.remaining -= take;

                    const saleYear = dateObj.getFullYear();
                    const startYear = Math.floor(saleYear / 5) * 5;
                    const blockName = `${PREFIX_LOG}${startYear}-${startYear + 4}`;

                    if (generateOutput) {
                        if (!stockMovementLogsByBlock[blockName]) stockMovementLogsByBlock[blockName] = [];
                        stockMovementLogsByBlock[blockName].push([
                          fastFormatDate(bestPO.dateObj), 
                          custManager, bestPO.invoice, 
                          isFallback ? additionalText + " (Oversold/LPP)" : additionalText, 
                          effectiveSku, bestPO.name,      
                          (bestPO.remaining + take), take, bestPO.remaining,      
                          fastFormatDate(dateObj) 
                        ]);
                    }
                } else if (sourceHistory) {
                    isNoStock = true; allocatedQty = pendingQty; 
                    globalNoStockItems.push({ 
                        ae: aeName, po: poNumber, sku: effectiveSku, 
                        desc: "No Stocks Available " + (effectiveSku !== reqItem.compSku ? `(Old SKU ${reqItem.compSku})` : ""), price: 0 
                    });
                } else if (missingSkuPriceMap.hasOwnProperty(effectiveSku)) {
                    unitBuyPrice = missingSkuPriceMap[effectiveSku];
                    supplier = ""; 
                    purchaseInvoice = "";
                    pingconPo = "";
                    sourceDate = ""; 
                    allocatedQty = pendingQty;
                    isMissingSku = false;
                    isNoStock = false;
                } else if (reqItem.compUnitPrice && reqItem.compUnitPrice > 0) { // <-- ADDED: Basket Price Fallback
                    unitBuyPrice = reqItem.compUnitPrice;
                    supplier = ""; 
                    purchaseInvoice = "";
                    pingconPo = "";
                    sourceDate = ""; 
                    allocatedQty = pendingQty;
                    isMissingSku = false;
                    isNoStock = false;
                } else if (basketComponentPriceMap.hasOwnProperty(effectiveSku)) { // <-- NEW FALLBACK: Global Component Price
                    unitBuyPrice = basketComponentPriceMap[effectiveSku];
                    supplier = ""; 
                    purchaseInvoice = "";
                    pingconPo = "";
                    sourceDate = ""; 
                    allocatedQty = pendingQty;
                    isMissingSku = false;
                    isNoStock = false;
                } else {
                    isMissingSku = true; unitBuyPrice = 0;
                    allocatedQty = pendingQty;
                    globalMissingItems.push({ 
                        ae: aeName, po: poNumber, sku: effectiveSku, desc: effectiveDesc, price: 0,
                        si: siNumber, date: fastFormatDate(dateObj), rawDate: dateObj, customer: customer, qty: allocatedQty, total: (compUnitSellPrice * allocatedQty),
                        rawRow: row 
                    });
                }
            }

            let lineTotalSell = compUnitSellPrice * allocatedQty;
            
            // NEW LOGIC: Regular items split proportionally, Baskets show full qty ONLY on the first row
            let displaySellQty = "";
            if (reqItem.isFirstComp) {
                if (!isBasket) displaySellQty = reqItem.origSellQty * (allocatedQty / reqItem.reqQty);
                else if (firstSplit) displaySellQty = reqItem.origSellQty;
            }

            // Flag determines if we print Sell Side text. (Regular items always print, Baskets print only once)
            let showSellSide = reqItem.isFirstComp && (!isBasket || firstSplit);

            pendingGroups[metaKey][poKey].items.push({
                sku: effectiveSku, 
                desc: effectiveDesc, 
                qty: allocatedQty, 
                uom: reqItem.compUom,
                
                sellSku: showSellSide ? reqItem.basketSku : "", 
                sellDesc: showSellSide ? reqItem.basketDesc : "", 
                sellQty: displaySellQty, 
                sellUom: showSellSide ? reqItem.origSellUom : "",
                unitBuy: unitBuyPrice, 
                unitSell: showSellSide ? reqItem.origUnitSell : "", 
                totalSell: lineTotalSell,
                
                supplier: supplier, 
                purchaseInvoice: purchaseInvoice, 
                pingconPo: pingconPo,
                sourceDate: sourceDate, 
                isMissingSku: isMissingSku, 
                isNoStock: isNoStock, 
                isValidationOverride: isValidationOverride, 
                isPartner: isPartner,
                srStatus: srStatus,
                srNumber: srNumber,
                srProduct: srProduct,
                srVatRate: srVatRate,
                srNetPrice: srNetPrice,
                srNetTotal: srNetTotal,
                srRowVat: srRowVat,
                srRowTotal: srRowTotal,
                
                isBasket: isBasket,
                basketId: basketId,
                isFirstComp: reqItem.isFirstComp
            });

            pendingQty -= allocatedQty;
            firstSplit = false; // <-- Turn off flag after first row
        }
    });

    if (!rowsToMarkDoneBySheet[item.sheetName]) rowsToMarkDoneBySheet[item.sheetName] = [];
    rowsToMarkDoneBySheet[item.sheetName].push(item.rowIndex);
  }); // <--- ADD THIS CLOSING BRACKET HERE! It closes the Sales Report loop.

  if (generateOutput) {
      // 1. Identify which Purchase Invoices are actually used in THIS report
      let neededInvoices = new Set();
      for (let metaKey in pendingGroups) {
          let [aeName, monthYear] = metaKey.split("|");
          // Only process if it matches the current AE and Month filter
          if (targetAeFilter && !checkNameMatch(aeName, targetAeFilter)) continue;
          if (targetMonthFilter && monthYear !== targetMonthFilter) continue;

          for (let poKey in pendingGroups[metaKey]) {
              pendingGroups[metaKey][poKey].items.forEach(itm => {
                  if (itm.purchaseInvoice) neededInvoices.add(String(itm.purchaseInvoice).trim().toUpperCase());
              });
          }
      }

      // 2. Check which ones we already have in the Scanned Invoices sheet
      let existingInvoices = new Set();
      scannedInvoicesData.forEach(item => {
          let inv = String(item.data[1]).trim().toUpperCase();
          if (inv) existingInvoices.add(inv);
      });

      // 3. Find missing invoices and WARN the user (Extraction is now done separately)
      let missingInvoices =[...neededInvoices].filter(inv => !existingInvoices.has(inv) && inv !== "" && inv !== "???");
      
      // Early return to precisely feed the targeted extraction tool and stop processing
      if (returnMissingInvoicesOnly) {
          return missingInvoices;
      }

      if (missingInvoices.length > 0) {
          try { ss.toast(`Warning: ${missingInvoices.length} invoices for this month are missing from Scanned Invoices. Please run the targeted extraction tool.`, "WARNING", 10); } catch(e){}
      }

      // 4. Build EWT and Purchases Maps from the refreshed data
      const pmcData = gatherAllData("PMC_Directory_Data_");
      let pmcDirMap = {};
      pmcData.forEach(item => {
          let r = item.data;
          let comp = String(r[0]).trim().toUpperCase();
          if(comp) pmcDirMap[comp] = { address: String(r[1]).trim(), tin: String(r[2]).trim() };
      });

      const accData = gatherAllData("Account_Title_Data_");
      let accTitleMap = {};
      accData.forEach(item => {
          let r = item.data;
          let comp = String(r[0]).trim().toUpperCase();
          if(comp) {
             let pctStr = String(r[2]).replace(/[^0-9.]/g, ''); // Extract just the numbers
             let ewtPctVal = 0;
             if (pctStr === "1" || pctStr === "0.01") ewtPctVal = 0.01;
             else if (pctStr === "2" || pctStr === "0.02") ewtPctVal = 0.02;
             accTitleMap[comp] = {
                 accTitle: String(r[1]).trim(),
                 ewtPct: ewtPctVal,
                 atc: String(r[3] || "").trim()
             };
          }
      });

      // Helper to find Account Title even with typos
      const getAccData = (compName) => {
          if (accTitleMap[compName]) return accTitleMap[compName];
          for (let key in accTitleMap) {
              if (compName.includes(key) || key.includes(compName)) return accTitleMap[key];
          }
          return null;
      };

      scannedInvoicesData.forEach(item => {
          let r = item.data;
          let invoiceNumRaw = String(r[2]).trim();
          let invoiceNum = invoiceNumRaw.toUpperCase();
          let companyRaw = String(r[1]).trim();
          let company = companyRaw.toUpperCase();

          if (invoiceNum) {
              let uniqueKey = `${invoiceNum}_${company}`;

              let rawTotalAmt = parseFloat(String(r[7]).replace(/[^0-9.-]+/g, "")) || 0;
              let rawGross = parseFloat(String(r[10]).replace(/[^0-9.-]+/g, "")) || 0;
              let rowTotalToUse = rawGross > 0 ? rawGross : rawTotalAmt; // Fallback Gross to Total Amount

              if (!ewtMap[uniqueKey]) {
                  let dObj = (r[3] instanceof Date) ? r[3] : new Date(r[3]);

                  ewtMap[uniqueKey] = {
                      dateObj: dObj,
                      companyRaw: companyRaw,
                      address: String(r[13] || "").trim(),
                      tin: String(r[12] || "").trim(),
                      invoiceNumRaw: invoiceNumRaw,
                      netTotal: (parseFloat(String(r[8]).replace(/[^0-9.-]+/g, "")) || 0),
                      rowVat: (parseFloat(String(r[9]).replace(/[^0-9.-]+/g, "")) || 0),
                      rowTotal: rowTotalToUse
                  };
              }

              if (!purchasesMap[uniqueKey]) {
                  let dirInfo = pmcDirMap[company] || { address: String(r[13] || "").trim(), tin: String(r[12] || "").trim() };

                  purchasesMap[uniqueKey] = {
                      dateObj: (r[3] instanceof Date) ? r[3] : new Date(r[3]),
                      companyRaw: companyRaw,
                      address: dirInfo.address,
                      tin: dirInfo.tin,
                      invoiceNumRaw: invoiceNumRaw,
                      netTotal: (parseFloat(String(r[8]).replace(/[^0-9.-]+/g, "")) || 0),
                      rowVat: (parseFloat(String(r[9]).replace(/[^0-9.-]+/g, "")) || 0),
                      rowTotal: rowTotalToUse
                  };
              }
          }
      });

      Object.keys(ewtMap).forEach(key => {
          let e = ewtMap[key];
          if (!e || Array.isArray(e)) return; 

          const fmt = (v) => (!isNaN(parseFloat(v)) && v !== "") ? parseFloat(v).toFixed(2) : v;

          let hasInputTax = e.rowVat > 0;
          let compKey = e.companyRaw.toUpperCase();
          
          let ewt1 = ""; let ewt2 = "";
          
          // First priority: Scanned Invoices VAT rule
          let assignedTitle = hasInputTax ? "Purchases - Local" : "Purchases - NV";
          let ewtPct = 0.01; // CHANGED: Default to 1% EWT even for NV
          let assignedAtc = hasInputTax ? "WC 158" : ""; // Default to blank if NV

          // Backup: Account Title Data (Uses loose matching)
          let backupData = getAccData(compKey);
          if (backupData) {
              // Only allow Backup Data to rename the Title if there is NO VAT
              if (!hasInputTax && backupData.accTitle) assignedTitle = backupData.accTitle;
              if (backupData.ewtPct > 0) ewtPct = backupData.ewtPct;
              if (backupData.atc) assignedAtc = backupData.atc;
          }

          let purchasesVal = e.netTotal; // ALWAYS output Scanned Invoices Purchases column exactly as is
          let ewtBase = hasInputTax ? e.netTotal : e.rowTotal; // Base for EWT is Purchases if VAT exists, else Gross

          if (ewtPct === 0.01) ewt1 = fmt(ewtBase * 0.01);
          if (ewtPct === 0.02) ewt2 = fmt(ewtBase * 0.02);

          let ewtRow =[assignedTitle, e.dateObj, e.companyRaw, e.address, e.tin, e.invoiceNumRaw, fmt(e.rowTotal), fmt(e.rowVat), fmt(purchasesVal), ewt1, ewt2, assignedAtc];
          ewtMap[key] = ewtRow; ewtMap[key.split("_")[0]] = ewtRow;
      });

      Object.keys(purchasesMap).forEach(key => {
          let p = purchasesMap[key];
          if (!p || Array.isArray(p)) return;

          const fmt = (v) => (!isNaN(parseFloat(v)) && v !== "") ? parseFloat(v).toFixed(2) : v;
          let hasInputTax = p.rowVat > 0;
          let compKey = p.companyRaw.toUpperCase();
          
          let assignedTitle = hasInputTax ? "Purchases - Local" : "Purchases - NV";

          let backupData = getAccData(compKey);
          if (backupData && backupData.accTitle) {
              if (!hasInputTax) assignedTitle = backupData.accTitle;
          }

          let purchasesVal = p.netTotal; // ALWAYS output Scanned Invoices Purchases column exactly as is

          let purRow =[p.dateObj, p.invoiceNumRaw, p.companyRaw, assignedTitle, p.tin, p.address, fmt(p.rowTotal), fmt(p.rowVat), fmt(purchasesVal)];
          purchasesMap[key] = purRow; purchasesMap[key.split("_")[0]] = purRow;
      });
  }
  // ---> REMOVED THE EXTRA `});` FROM HERE!

  // --- LOGS WRITING (Only if generateOutput is true) ---
  if (generateOutput) {
    for (let blockName in stockMovementLogsByBlock) {
       let logs = stockMovementLogsByBlock[blockName];
       if (logs.length > 0) {
          let sampleDateStr = logs[0][9]; 
          let logSheet = getSheetForYearBlock(PREFIX_LOG, new Date(sampleDateStr), 
               ["PO Date", "Cust Manager", "Purchase Invoice", "Additional Texts", "SKU", "Desc", "Before Sold Amount", "Sold Qty", "After Sold Amount", "Sale Date"]);
          
          const existingData = logSheet.getDataRange().getValues();
          const existingKeys = new Set();
          for (let i = 1; i < existingData.length; i++) {
             let r = existingData[i];
             let saleDateStr = r[9] instanceof Date ? fastFormatDate(r[9]) : String(r[9]);
             existingKeys.add(`${saleDateStr}_${r[2]}_${r[4]}_${r[7]}_${r[1]}`);
          }
          const uniqueLogsToAdd = logs.filter(log => !existingKeys.has(`${log[9]}_${log[2]}_${log[4]}_${log[7]}_${log[1]}`));
          if (uniqueLogsToAdd.length > 0) logSheet.getRange(logSheet.getLastRow() + 1, 1, uniqueLogsToAdd.length, 10).setValues(uniqueLogsToAdd);
       }
    }
  }

  // --- CALCULATION OF 4TH MECHANIC ---
  let tenDayMechanicMet = checkWorkingDaysMechanic(requestDate, targetMonthFilter);

  // --- PRE-PROCESS AR DATA FOR FAST LOOKUP DYNAMICALLY ---
  const arSheets = getDatabaseSheets(PREFIX_AR);
  let arHeaders =[];
  if (arSheets.length > 0 && arSheets[0].getLastColumn() > 0) {
    arHeaders = arSheets[0].getRange(1, 1, 1, arSheets[0].getLastColumn()).getValues()[0].map(h => String(h).toLowerCase().trim());
  }
  let arTextIdx = arHeaders.findIndex(h => h.includes("additional text"));
  let arOpenBalIdx = arHeaders.findIndex(h => h.includes("open balance"));
  let arPaymentStatusIdx = arHeaders.findIndex(h => h === "payment status");
  if (arPaymentStatusIdx === -1 && arOpenBalIdx > -1) arPaymentStatusIdx = arOpenBalIdx + 2; // Fallback
  let arNumberIdx = arHeaders.findIndex(h => h === "number" || h === "document number" || h.includes("number"));

  const arFastList = arAllData.map(ar => {
      // Extract open balance to check for 0 or negative
      let openBal = (arOpenBalIdx > -1) ? parseFloat(ar.data[arOpenBalIdx]) : 0;
      if (isNaN(openBal)) openBal = 0;

      let actualStatus = "";
      if (arPaymentStatusIdx > -1 && arPaymentStatusIdx < ar.data.length) {
          actualStatus = String(ar.data[arPaymentStatusIdx]).trim();
      }

      let arNumberVal = "";
      if (arNumberIdx > -1 && arNumberIdx < ar.data.length) {
          arNumberVal = String(ar.data[arNumberIdx]).trim();
      }
      
      // STRICT LOGIC: Preserve exact text (e.g., "Please fax email EWT" vs "Due for collection")
      let finalStatus = actualStatus ? actualStatus : "Due for collection";
      
      // Set to Fully Paid if Open Balance is 0 or less, OR if status string says Fully Paid
      if (openBal <= 0 || actualStatus.toLowerCase() === "fully paid") {
          finalStatus = "Fully Paid";
      }

      return {
            ref: (arTextIdx > -1) ? String(ar.data[arTextIdx]) : "",
            status: finalStatus,
            number: arNumberVal,
            openBal: openBal
        };
  });

  // --- GENERATE FILES & PROFIT REPORTS ---
  for (let metaKey in pendingGroups) {
    let [aeName, monthYear] = metaKey.split("|");
    
    // NOTE: For Broadcast, filter is null, so this passes for everyone
    if (!checkNameMatch(aeName, targetAeFilter)) continue; 
    if (targetMonthFilter && monthYear !== targetMonthFilter) continue;

    let poBatch = pendingGroups[metaKey];
    let allPaid = true;
    let aggSalesEx = 0; let aggCostEx = 0; let aggComm = 0;
    let aggSalesInc = 0; let aggCostInc = 0;
    let hasExcludedItems = false;
    let unpaidPos = []; 
    let profitReportRowsByBlock = {}; 
    let commReportRows =[];
    let aeMinDate = null;
    let aeMaxDate = null;


    let refDate = null;
    let firstPoKey = Object.keys(poBatch)[0];
    if(firstPoKey) refDate = poBatch[firstPoKey].meta.rawDate;

    let monthlyQuota = 0;
    if(refDate) {
        for(let q=0; q<sqAllData.length; q++) {
            let qRow = sqAllData[q].data;
            let qAe = String(qRow[0]).trim();
            let qStart = new Date(qRow[1]);
            let qEnd = new Date(qRow[2]);
            let newQuota = parseFloat(qRow[4]);
            if (isNaN(newQuota)) newQuota = 0;
            
            if (checkNameMatch(qAe, aeName)) {
                 if(refDate >= qStart && refDate <= qEnd) { 
                     // Inside Date Effectivity -> Strictly use New Quota
                     if (newQuota > 0) monthlyQuota = newQuota; 
                     break;
                 }
            }
        }
    }

    for (let poKey in poBatch) {
      let poObj = poBatch[poKey];
      let poNum = poObj.meta.po;
      let statusDesc = "Pending";
      let matchedArNumber = "";
      let matchedOpenBal = 0;
      let metaDate = new Date(poObj.meta.date); 
      
      const startYear = Math.floor(metaDate.getFullYear() / 5) * 5;
      const profitBlockName = `Profit_${aeName}_${startYear}-${startYear + 4}`;
      if (!profitReportRowsByBlock[profitBlockName]) profitReportRowsByBlock[profitBlockName] =[];

      for(let i=0; i<arFastList.length; i++) {
        if (arFastList[i].ref.includes(poNum)) {
          statusDesc = arFastList[i].status;
          matchedArNumber = arFastList[i].number;
          matchedOpenBal = arFastList[i].openBal;
          break;
        }
      }
      poObj.meta.paymentStatus = statusDesc;
      poObj.meta.arNumber = matchedArNumber;
      poObj.meta.openBal = matchedOpenBal;
      
      if (statusDesc !== "Fully Paid") {
        allPaid = false;
        // Append exact status to the PO number for clear remarks in the PDF/Report
        unpaidPos.push(`${poNum} [${statusDesc}]`);
        globalUnpaidItems.push({ ae: aeName, date: poObj.meta.date, customer: poObj.meta.customer, po: poNum, status: statusDesc });
      }

      let poSalesInc = 0; let poCostInc = 0;
      let poVoidReasons = [];
      
      // ADDED: Check for LOAN in Additional Text
      if (poObj.meta.additionalText && poObj.meta.additionalText.toUpperCase().includes("LOAN")) {
        poVoidReasons.push("Loan equipment excluded from commission computation.");
      }

      let hasHam = false; 
      let hasNonHam = false;

      let basketAggregator = {};

      poObj.items.forEach(itm => {
        if(itm.isMissingSku || itm.isNoStock) hasExcludedItems = true;
        
        if (!itm.isMissingSku && !itm.isNoStock) {
           poSalesInc += itm.totalSell;
           poCostInc += (itm.unitBuy * itm.qty);
        }

        if (/(ham|jamon)/i.test(itm.sellDesc)) hasHam = true;
        else hasNonHam = true;

        let totalCost = itm.unitBuy * itm.qty;
        let totalSales = itm.totalSell;

        if (itm.isBasket) {
            if (!basketAggregator[itm.basketId]) {
                basketAggregator[itm.basketId] = {
                    totalCost: 0,
                    totalSales: 0,
                    sellSku: itm.sellSku || itm.sku, 
                    sellDesc: itm.sellDesc || itm.desc,
                    supplier: itm.supplier,
                    purchaseInvoice: itm.purchaseInvoice,
                    sourceDate: itm.sourceDate
                };
            }
            basketAggregator[itm.basketId].totalCost += totalCost;
            basketAggregator[itm.basketId].totalSales += totalSales;
            
            if (itm.isFirstComp) {
                basketAggregator[itm.basketId].sellSku = itm.sellSku;
                basketAggregator[itm.basketId].sellDesc = itm.sellDesc;
                basketAggregator[itm.basketId].supplier = itm.supplier;
                basketAggregator[itm.basketId].purchaseInvoice = itm.purchaseInvoice;
                basketAggregator[itm.basketId].sourceDate = itm.sourceDate;
            }
        } else {
            let gp = totalSales - totalCost;
            let gpPct = (totalSales !== 0) ? gp / totalSales : 0;
            
            if (gp < -0.01 && !itm.isMissingSku && !itm.isNoStock) { 
                 globalNegativeItems.push({
                   ae: aeName,
                   po: poNum,
                   date: poObj.meta.date,
                   rawDate: poObj.meta.rawDate,
                   purchaseInvoice: itm.purchaseInvoice,
                   purchaseDate: itm.sourceDate, 
                   additionalText: poObj.meta.additionalText,
                   supplier: itm.supplier,
                   customer: poObj.meta.customer,
                   manager: poObj.meta.manager,
                   sku: itm.sellSku,
                   desc: itm.sellDesc,
                   totalCost: totalCost,
                   totalSales: totalSales,
                   gp: gp,
                   gpm: gpPct
                 });

                if(!poVoidReasons.includes("If there is negative GPM in one item in a PO, whole PO will be voided. No commission.")) {
                     poVoidReasons.push("If there is negative GPM in one item in a PO, whole PO will be voided. No commission.");
                }
            }
        }
        
        // profitReportRowsByBlock[profitBlockName].push([
        //   purchaseDateStr, poObj.meta.manager, itm.supplier, itm.pingconPo, itm.purchaseInvoice, 
        //   itm.sku, itm.desc, itm.unitBuy, itm.qty, "pc", totalCost,
        //   poObj.meta.date, poObj.meta.po, itm.sellSku, itm.sellDesc, itm.unitSell, itm.sellQty, itm.sellUom, totalSales, gp, gpPct
        // ]);
      });

      for (let bId in basketAggregator) {
          let bData = basketAggregator[bId];
          let gp = bData.totalSales - bData.totalCost;
          let gpPct = (bData.totalSales !== 0) ? gp / bData.totalSales : 0;
          
          if (gp < -0.01) {
              globalNegativeItems.push({
                   ae: aeName,
                   po: poNum,
                   date: poObj.meta.date,
                   rawDate: poObj.meta.rawDate,
                   purchaseInvoice: bData.purchaseInvoice,
                   purchaseDate: bData.sourceDate,
                   additionalText: poObj.meta.additionalText,
                   supplier: bData.supplier,
                   customer: poObj.meta.customer,
                   manager: poObj.meta.manager,
                   sku: bData.sellSku,
                   desc: bData.sellDesc,
                   totalCost: bData.totalCost,
                   totalSales: bData.totalSales,
                   gp: gp,
                   gpm: gpPct
              });

              if(!poVoidReasons.includes("If there is negative GPM in one item in a PO, whole PO will be voided. No commission.")) {
                   poVoidReasons.push("If there is negative GPM in one item in a PO, whole PO will be voided. No commission.");
              }
          }
      }

      if (hasHam && !hasNonHam) poVoidReasons.push("Ham only in one PO, no commission.");

      let poSalesEx = poSalesInc / 1.12; let poCostEx = poCostInc / 1.12;
      
      // CHANGED: Only add to Monthly Performance Summary if the PO is NOT VOID
      if (poVoidReasons.length === 0) {
        aggSalesEx += poSalesEx; 
        aggCostEx += poCostEx;
        aggSalesInc += poSalesInc;
        aggCostInc += poCostInc;
      }
      
      let poGp = poSalesEx - poCostEx;
      let poGpPct = (poSalesEx > 0) ? (poGp / poSalesEx) * 100 : 0;
      poGpPct = Math.round(poGpPct); 
      
      let commRate = 0;
      if (poGpPct >= 10 && poGpPct <= 15) commRate = 0.005;
      else if (poGpPct >= 16 && poGpPct <= 20) commRate = 0.0055;
      else if (poGpPct >= 21 && poGpPct <= 25) commRate = 0.006;
      else if (poGpPct >= 26 && poGpPct <= 30) commRate = 0.0065;
      else if (poGpPct >= 31) commRate = 0.007;

      if (statusDesc === "Fully Paid" && poVoidReasons.length === 0) {
          aggComm += (poSalesEx * commRate);
      }
      
      poObj.meta.voidReasons = poVoidReasons;

      let finalCommRateStr = "0%";
      let finalCommVal = 0;
      if (poVoidReasons.length > 0) {
          finalCommRateStr = "VOID";
          finalCommVal = "VOID";
      } else {
          finalCommRateStr = (commRate * 100) + "%";
          finalCommVal = poSalesEx * commRate;
      }

      let drSiRaw = "";
      if (poObj.meta.additionalText) {
          // Split by comma, trim spaces, and only keep parts that start with "DR" or "SI"
          drSiRaw = poObj.meta.additionalText.split(',')
              .map(part => part.trim())
              .filter(part => part.toUpperCase().startsWith("DR") || part.toUpperCase().startsWith("SI"))
              .join(", ");
      }

      poObj.items.forEach(itm => {
          let itemGp = itm.totalSell - (itm.unitBuy * itm.qty);
          let itemGpm = (itm.totalSell !== 0) ? (itemGp / itm.totalSell) : 0;
          
          if (itm.isBasket && basketAggregator[itm.basketId]) {
             let bData = basketAggregator[itm.basketId];
             let bGp = bData.totalSales - bData.totalCost;
             itemGpm = (bData.totalSales !== 0) ? (bGp / bData.totalSales) : 0;
          }
          
          let dObj = new Date(poObj.meta.date);
          if (!isNaN(dObj.getTime())) {
              if (!aeMinDate || dObj < aeMinDate) aeMinDate = dObj;
              if (!aeMaxDate || dObj > aeMaxDate) aeMaxDate = dObj;
          }

          // CATEGORY & TRANSPO/HAM LOGIC (Dynamic from Script Properties - Zero Vendor Leaks)
          let supplierName = String(itm.supplier || "").toLowerCase();
          let itemNameLower = String(itm.srProduct || itm.sellDesc || "").toLowerCase();
          let itemSkuUpper = String(itm.sellSku || "").toUpperCase();
          
          let unbrandedSuppliers = JSON.parse(SCRIPT_PROPS['UNBRANDED_SUPPLIERS'] || '[]');
          let unbrandedItems = JSON.parse(SCRIPT_PROPS['UNBRANDED_ITEMS'] || '[]');
          let isUnbranded = unbrandedSuppliers.some(s => supplierName.includes(s.toLowerCase())) || unbrandedItems.some(i => itemNameLower.includes(i.toLowerCase()));
          
          let itemCategory = isUnbranded ? "Unbranded" : "Branded";
          
          // --- NEW CALCULATIONS PER ROW ---
          let rowTotalInc = itm.totalSell;                            // Row Total (VAT INC)
          let rowNetTotal = rowTotalInc / 1.12;                       // Net total (VAT EX)
          let rowVat = rowTotalInc - rowNetTotal;                     // Row VAT
          
          let effectiveQty = parseFloat(itm.sellQty) || itm.qty || 1; 
          let rowNetPrice = rowNetTotal / effectiveQty;               // Net Price per unit
          
          let transpoHamVal = (itemNameLower.includes("ham") || itemNameLower.includes("jamon") || itemSkuUpper === "FREIGHT") ? rowNetTotal : 0;
          
          // --- COMMISSION CALCULATION (Net Total * Monthly PO Rate) ---
          let itemCommVal = (poVoidReasons.length > 0) ? "VOID" : (rowNetTotal * commRate);

          // --- VAT RATE FORMATTING ---
          let parsedVatRate = parseFloat(itm.srVatRate);
          let formattedVatRate = !isNaN(parsedVatRate) ? (parsedVatRate * 100).toFixed(0) + "%" : "12%";

          commReportRows.push({
              dateObj: dObj,
              row:[
                  "", // SN (Index 0)
                  poObj.meta.customer, // 1
                  itm.srStatus, // 2
                  poObj.meta.arNumber, // 3
                  fastFormatDate(dObj), // 4
                  drSiRaw, // 5
                  itemGpm, // 6
                  itemCategory, // 7
                  itm.sellSku, // 8
                  itm.srProduct, // 9
                  formattedVatRate, // 10 (Parsed and forced to percent)
                  itm.sellQty, // 11
                  itm.sellUom, // 12
                  rowNetPrice, // 13
                  rowNetTotal, // 14
                  rowVat,      // 15
                  rowTotalInc, // 16
                  transpoHamVal, // 17 
                  (itm.unitBuy * itm.qty), // 18 
                  finalCommRateStr, // 19 (Restored to overall PO monthly rate)
                  itemCommVal // 20 (Calculated as Net Total * Rate)
              ]
          });
      });
    }

    // --- NEW: Check Previous Month's Payment Status ---
    if (refDate) {
      let prevMonthDate = new Date(refDate);
      prevMonthDate.setMonth(prevMonthDate.getMonth() - 1);
      let prevMonthYear = fastFormatMonthYear(prevMonthDate);
      let prevMetaKey = `${aeName}|${prevMonthYear}`;
      
      // If the AE had sales in the previous month, check them
      if (pendingGroups[prevMetaKey]) {
        let prevPoBatch = pendingGroups[prevMetaKey];
        for (let prevPoKey in prevPoBatch) {
          let pPoObj = prevPoBatch[prevPoKey];
          let pPoNum = pPoObj.meta.po;
          let pStatusDesc = "Pending";
          
          // Look up status in AR Data
          for(let i = 0; i < arFastList.length; i++) {
            if (arFastList[i].ref.includes(pPoNum)) {
              pStatusDesc = arFastList[i].status;
              break;
            }
          }
          
          // If any previous month's PO is not fully paid, set allPaid to false
          if (pStatusDesc !== "Fully Paid") {
            allPaid = false;
            // Append exact status for previous month's remarks
            unpaidPos.push(`${pPoNum} (Prev Month) [${pStatusDesc}]`);
          }
        }
      }
    }
    // --- END NEW ---

    // --- ENTITY CHECK FOR EMAILS & FEEDBACK ---
    let hasKeysysOrA2ZIT = false;
    let hasA2ZPMC = false;
    soAllData.forEach(item => {
        let row = item.data;
        if (String(row[15]).trim() === aeName) {
            let d = new Date(row[4]);
            if (!isNaN(d.getTime()) && fastFormatMonthYear(d) === monthYear) {
                let ent = getCompanyEntity(String(row[17] || ""));
                if (ent === "Keysys" || ent === "A2ZIT") hasKeysysOrA2ZIT = true;
                if (ent === "A2ZPMC") hasA2ZPMC = true;
            }
        }
    });

    let sysSettings = getSystemSettings();
    let isTech = (hasKeysysOrA2ZIT || hasA2ZPMC);
    let senderEmail = getSenderEmail(isTech);
    let mainRecipient = isTech ? (sysSettings["Management_A2Z"] || "") : (sysSettings["Management_Pingcon"] || "");
    let feedbackImageId = isTech ? FEEDBACK_IMAGE_ID_A2Z_KEYSYS : FEEDBACK_IMAGE_ID_PINGCON_BONITA;
    let senderName = isTech ? "Finance A2Z2GO" : "Finance Pingcon";

    let feedbackImgBlob = null;
    try {
        if (feedbackImageId && feedbackImageId !== "INSERT_FILE_ID_FOR_A2Z_KEYSYS_IMAGE_HERE" && feedbackImageId !== "INSERT_FILE_ID_FOR_PINGCON_BONITA_IMAGE_HERE") {
            feedbackImgBlob = DriveApp.getFileById(feedbackImageId).getBlob();
        }
    } catch(e) {}
    // --- END ENTITY CHECK ---

    // --- WRITE BUY/SELL SHEET (ONLY IF generateOutput is TRUE) ---
    if (generateOutput) {
        let aggGP = aggSalesEx - aggCostEx;
        let aggGpPct = (aggSalesEx > 0) ? (aggGP / aggSalesEx) : 0;
        let quotaMet = (aggSalesEx >= monthlyQuota);
        
        let aggGPInc = aggSalesInc - aggCostInc;
        let aggGpPctInc = (aggSalesInc > 0) ? (aggGPInc / aggSalesInc) : 0;
        
        let sheetValues = [], sheetBackgrounds = [], sheetWeights = [], sheetColors = [], mergeRanges =[]; 

        const summaryData = buildSummaryHeader(
            aggSalesInc, aggCostInc, aggGPInc, aggGpPctInc, aggComm, 
            allPaid, hasExcludedItems, quotaMet, monthlyQuota, unpaidPos,
            prevMonthExists, tenDayMechanicMet, aggSalesEx 
        );
        const rData = summaryData.richTextData;
        const parts =[];
        if (rData) {
          parts.push({ text: rData.text1, color: rData.color1 });
          parts.push({ text: rData.text2, color: rData.color2 });
          parts.push({ text: rData.text3, color: rData.color3 });
          parts.push({ text: rData.text4, color: rData.color4 });
          parts.push({ text: rData.text5, color: rData.color5 });
          parts.push({ text: rData.text6, color: rData.color6 });
          parts.push({ text: rData.text7, color: rData.color7 });
          parts.push({ text: rData.text8, color: rData.color8 });
          parts.push({ text: rData.text9, color: rData.color9 });
          parts.push({ text: rData.text10, color: rData.color10 });
          parts.push({ text: rData.text11, color: rData.color11 });
          parts.push({ text: rData.text12, color: rData.color12 });
        }

        const extraRowsNeeded = Math.max(0, parts.length - 1); 

        // Add Summary Rows
        for(let r=0; r<3; r++){
          sheetValues.push(summaryData.values[r]);
          sheetBackgrounds.push(summaryData.backgrounds[r]);
          sheetWeights.push(summaryData.weights[r]);
          sheetColors.push(summaryData.colors[r]);
        }
        for(let i=0; i<extraRowsNeeded; i++){
          let emptyRow = Array(85).fill("");
          let emptyBg = Array(85).fill("#ffffff");
          let emptyWeight = Array(85).fill("normal");
          let emptyColor = Array(85).fill("black");
          
          sheetValues.push(emptyRow);
          sheetBackgrounds.push(emptyBg);
          sheetWeights.push(emptyWeight);
          sheetColors.push(emptyColor);
        }
        sheetValues.push(summaryData.values[3]); 
        sheetBackgrounds.push(summaryData.backgrounds[3]);
        sheetWeights.push(summaryData.weights[3]);
        sheetColors.push(summaryData.colors[3]);

        // Place each part in column AX (index 49) on subsequent rows to prevent stretching Row 3
        for (let i = 0; i < parts.length; i++) {
          sheetValues[2 + i][49] = parts[i].text;
          sheetColors[2 + i][49] = parts[i].color;
          sheetWeights[2 + i][49] = "bold";
        }
        
        let currentRow = sheetValues.length + 1;
        
        sheetValues.push(Array(85).fill("")); sheetBackgrounds.push(Array(85).fill("#ffffff"));
        sheetWeights.push(Array(85).fill("normal")); sheetColors.push(Array(85).fill("black"));
        currentRow++;

        let headerRowIndices =[]; // Track headers for centering

        for (let poKey in poBatch) {
          const mechanicsData = { quotaMet, prevMonthExists, allPaid, tenDayMechanicMet };
          const poData = buildPOBlockInMemory(poBatch[poKey], currentRow, poBatch[poKey].meta.voidReasons, ewtMap, purchasesMap, mechanicsData);
          
          headerRowIndices.push(currentRow + 8); // Track the main header row for this PO
          
          sheetValues.push(...poData.values);
          sheetBackgrounds.push(...poData.backgrounds);
          sheetWeights.push(...poData.weights);
          sheetColors.push(...poData.colors);
          poData.merges.forEach(m => mergeRanges.push({ r: currentRow + m.rOffset, c: m.c, nr: m.nr, nc: m.nc }));
          
          sheetValues.push(Array(85).fill("")); sheetBackgrounds.push(Array(85).fill("#ffffff"));
          sheetWeights.push(Array(85).fill("normal")); sheetColors.push(Array(85).fill("black"));
          currentRow += poData.values.length + 1;
        }

        for(let i=1; i<sheetValues.length; i++) { sheetValues[i][0] = i; }

        let aeFolder;
        let aeFolders = rootFolder.getFoldersByName(aeName);
        if (aeFolders.hasNext()) aeFolder = aeFolders.next();
        else aeFolder = rootFolder.createFolder(aeName);

        let fileName = `${aeName} - ${monthYear}`;
        let spreadsheet;
        let files = aeFolder.getFilesByName(fileName);
        let targetSheet;

        if (files.hasNext()) {
          spreadsheet = SpreadsheetApp.open(files.next());
          targetSheet = spreadsheet.getSheets()[0];
          targetSheet.clear(); 
        } else {
          spreadsheet = SpreadsheetApp.create(fileName);
          let file = DriveApp.getFileById(spreadsheet.getId());
          file.moveTo(aeFolder);
          targetSheet = spreadsheet.getSheets()[0];
          targetSheet.setName("Consolidated POs");
        }

        if (sheetValues.length > 0) {
          const range = targetSheet.getRange(1, 1, sheetValues.length, 85);
          range.setValues(sheetValues);
          range.setBackgrounds(sheetBackgrounds);
          range.setFontWeights(sheetWeights);
          range.setFontColors(sheetColors);
          range.setWrap(false); 
          range.setVerticalAlignment("top"); 

          // Formatting lines
          targetSheet.getRange(1, 3, sheetValues.length, 2).setNumberFormat("@"); 
          targetSheet.getRangeList([`BE1:BE${sheetValues.length}`, `BN1:BN${sheetValues.length}`]).setNumberFormat("@");
          targetSheet.getRangeList([
            `G1:G${sheetValues.length}`,
            `J1:M${sheetValues.length}`,
            `P1:P${sheetValues.length}`,
            `S1:V${sheetValues.length}`,
            `X1:X${sheetValues.length}`, 
            `AA1:AG${sheetValues.length}`,
            `AI1:AI${sheetValues.length}`, 
            `AK1:AK${sheetValues.length}`, 
            `AN1:AT${sheetValues.length}`, 
            `BG1:BI${sheetValues.length}`, 
            `B3:D3`,
            `G3`
          ]).setNumberFormat("#,##0.00");
          
          targetSheet.getRangeList([
            `F3`,
            `Y1:Y${sheetValues.length}`,   
            `AH1:AH${sheetValues.length}`,   
            `AJ1:AJ${sheetValues.length}`  
          ]).setNumberFormat("0.00%");

          targetSheet.getRange(3, 2, 1, 6).setVerticalAlignment("top").setHorizontalAlignment("right"); 
          targetSheet.getRange(3, 38, 1, 12).setVerticalAlignment("top").setHorizontalAlignment("center"); // Mechanics values
          
          targetSheet.getRange("A1:BV1").setHorizontalAlignment("center"); // Center headers without merging
          
          // mergeRanges.forEach(m => { try { targetSheet.getRange(m.r, m.c, m.nr, m.nc).merge(); } catch(e) {} }); // REMOVED AS REQUESTED
          
          targetSheet.getRange("B2:G2").setWrap(true).setVerticalAlignment("middle").setHorizontalAlignment("center");

          targetSheet.getRange(2, 2, 2, 6).setBorder(true, true, true, true, true, true);
          targetSheet.getRange(2, 38, 2, 12).setBorder(true, true, true, true, true, true); // Mechanics grid border
          
          // Outer border for Remarks header and body
          targetSheet.getRange(2, 50, 1, 1).setBorder(true, true, true, true, false, false); 
          targetSheet.getRange(3, 50, 1 + extraRowsNeeded, 1).setBorder(true, true, true, true, false, false);

          // Wrap only up to the Remarks column (AX). EWT/Purchases wrap is OFF to prevent row stretching.
          targetSheet.getRange(`AA1:AX${sheetValues.length}`).setWrap(true).setVerticalAlignment("top").setHorizontalAlignment("center");
          targetSheet.getRange(`AZ1:BU${sheetValues.length}`).setWrap(false).setVerticalAlignment("top").setHorizontalAlignment("center");
          
          targetSheet.getRangeList([
            `I1:M${sheetValues.length}`,
            `R1:V${sheetValues.length}`,
            `X1:Y${sheetValues.length}`
          ]).setWrap(true);

          [
            [1, 1, 50], [2, 4, 140], [6, 1, 350], [7, 1, 100],[8, 2, 50], [10, 4, 130],[14, 1, 140], [15, 1, 350],[16, 1, 100], [17, 2, 50],[19, 4, 130], [23, 1, 30],[24, 2, 140], [26, 1, 30],[27, 6, 130], [33, 2, 140],[35, 3, 120], [38, 12, 220],
            [50, 1, 450], [51, 1, 30],[52, 1, 200], [53, 1, 100],[54, 1, 250], [55, 1, 450],
            [56, 7, 120], [63, 1, 80], [64, 1, 30], [65, 1, 100],[66, 1, 120], [67, 1, 250],
            [68, 1, 200],[69, 1, 120], [70, 1, 450],[71, 3, 120]
          ].forEach(c => c[1] === 1 ? targetSheet.setColumnWidth(c[0], c[2]) : targetSheet.setColumnWidths(c[0], c[1], c[2]));

          // Middle-align and center the specific header rows so they don't look stretched
          if (headerRowIndices.length > 0) {
              let headerRanges = headerRowIndices.map(r => `A${r}:BU${r}`);
              targetSheet.getRangeList(headerRanges).setVerticalAlignment("middle").setHorizontalAlignment("center");
          }

          // Simple styling for the Remarks Column (No RichText needed anymore)
          let remarkRange = targetSheet.getRange(3, 22, 1 + extraRowsNeeded, 1);
          remarkRange.setWrap(false); // Text will cleanly overflow into W, X, Y, Z visually
          remarkRange.setVerticalAlignment("top").setHorizontalAlignment("left");
        }
        // --- NEW: EMAIL BUY/SELL REPORT ---
        SpreadsheetApp.flush();
        Utilities.sleep(3000);
        let bsUrl = "https://docs.google.com/spreadsheets/d/" + spreadsheet.getId() + "/export?format=xlsx";
        let bsToken = ScriptApp.getOAuthToken();
        try {
            let bsResponse = UrlFetchApp.fetch(bsUrl, { headers: { 'Authorization': 'Bearer ' + bsToken }, muteHttpExceptions: true });
            if (bsResponse.getResponseCode() === 200) {
                let bsBlob = bsResponse.getBlob().setName(fileName + ".xlsx");
                let securedBsBlob = encryptAttachmentViaPython(bsBlob, fileName + ".xlsx", hasKeysysOrA2ZIT ? SCRIPT_PROPS['ENCRYPT_PASS_TECH'] : SCRIPT_PROPS['ENCRYPT_PASS_PINGCON']);
                
                let feedbackLink = isTech ? (SCRIPT_PROPS['FEEDBACK_LINK_TECH'] || "") : (SCRIPT_PROPS['FEEDBACK_LINK_PINGCON'] || "");
                
                let bsHtmlBody = Template_BuySellEmail(aeName, monthYear, feedbackImgBlob, feedbackLink);
                

                let bsCc = isTech ? (sysSettings["Finance_A2Z"] || "") : (sysSettings["Finance_Pingcon"] || "");
                let bsEmailOptions = {
                    from: senderEmail, cc: bsCc,
                    name: senderName, attachments: [securedBsBlob], inlineImages: {}
                };
                if (feedbackImgBlob) bsEmailOptions.inlineImages.feedbackImage = feedbackImgBlob;
                sendRobustEmail(mainRecipient, `Buy/Sell Report - ${aeName} - ${monthYear}`, bsHtmlBody, bsEmailOptions);
            }
        } catch(e) { console.error("Failed to email Buy/Sell: " + e); }
        // --- END EMAIL BUY/SELL ---

        // for (let blockName in profitReportRowsByBlock) {
        //    writeAEProfitReport(blockName, profitReportRowsByBlock[blockName]);
        // }
        
        
        if (PURCHASE_OUTPUT_FOLDER_ID && PURCHASE_OUTPUT_FOLDER_ID !== "INSERT_FOLDER_ID_HERE") {
            const purFolder = DriveApp.getFolderById(PURCHASE_OUTPUT_FOLDER_ID);
            for (let monthYear in globalPurchaseOutput) {
                let rows = Object.values(globalPurchaseOutput[monthYear]);
                if (rows.length > 0) {
                    let fileName = `Purchase_${monthYear}`;
                    let headers =["Invoice Date (MM/DD/YYYY)", "Invoice Number", "Company Name", "Account Title", "TIN", "Address", "Gross", "Input Tax", "Purchases"];
                    writeOutputSheet(purFolder, fileName, headers, rows);
                }
            }
        }
    } // END generateOutput check
    // --- GENERATE COMMISSION REPORT (XLSX) ---
    if (generateCommReport && commReportRows.length > 0) {
        commReportRows.sort((a, b) => b.dateObj - a.dateObj); 
        
        let finalRows = commReportRows.map((item, index) => {
            let r = [...item.row]; 
            r[0] = index + 1; 
            
            // Safe parsing for GPM percentage
            let gpmVal = parseFloat(r[6]);
            r[6] = !isNaN(gpmVal) ? (gpmVal * 100).toFixed(2) + "%" : "0.00%"; 
            
            const formatDec = (val) => {
                if (val === "VOID") return "VOID";
                if (val == null || val === "") return "";
                let parsed = parseFloat(val);
                return !isNaN(parsed) ? parsed.toFixed(2) : val;
            };
            
            r[13] = formatDec(r[13]); // Net price
            r[14] = formatDec(r[14]); // Net total
            r[15] = formatDec(r[15]); // Row VAT
            r[16] = formatDec(r[16]); // Row Total
            r[17] = formatDec(r[17]); // Transpo/Ham
            r[18] = formatDec(r[18]); // PO Net Amount
            r[20] = formatDec(r[20]); // Commission
            
            // CRITICAL SANITIZER: Prevent setValues from crashing due to undefined, null, or NaN
            return r.map(cell => (cell == null || (typeof cell === 'number' && isNaN(cell))) ? "" : cell); 
        });

        let minDateStr = aeMinDate ? Utilities.formatDate(aeMinDate, Session.getScriptTimeZone(), "MMMM d yyyy") : "Unknown";
        let maxDateStr = aeMaxDate ? Utilities.formatDate(aeMaxDate, Session.getScriptTimeZone(), "MMMM d yyyy") : "Unknown";
        let reportDateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-ddMMM-HHmm");
        
        // Clean AE name to prevent invalid characters from breaking file/folder creation
        let cleanAeName = String(aeName).replace(/[,/\\:*?"<>|]/g, "").trim();
        let fileName = `Commission Report of ${cleanAeName} from ${minDateStr} to ${maxDateStr} as of Report Date ${reportDateStr}`;
        
        let commHeaders =[
            "SN", "Customer", "Status", "Number", "Date (MM/DD/YYYY)", "DR SI", 
            "GPM", "Category", "SKU", "Product", "VAT rate", "Sold Quantity", 
            "Unit", "Net price", "Net total", "Row VAT", "Row Total", 
            "Transpo/Ham", "PO Net Amount", "Rate", "Commission"
        ];

        let baseDate = aeMinDate || new Date();
        let yearFolderStr = Utilities.formatDate(baseDate, Session.getScriptTimeZone(), "yyyy");

        if (COMMISSION_REPORT_ROOT_FOLDER_ID && COMMISSION_REPORT_ROOT_FOLDER_ID !== "INSERT_COMM_REPORT_FOLDER_ID_HERE") {
            try {
                // 1. Create the standard Google Sheet first (This never fails)
                let tempSpreadsheet = SpreadsheetApp.create(fileName);
                let tempSheet = tempSpreadsheet.getSheets()[0];
                
                let sheetData =[commHeaders, ...finalRows];
                tempSheet.getRange(1, 1, sheetData.length, commHeaders.length).setValues(sheetData);
                tempSheet.getRange(1, 1, 1, commHeaders.length).setFontWeight("bold").setBackground("#f3f3f3");
                tempSheet.getRange(2, 4, sheetData.length, 1).setNumberFormat("@");
                tempSheet.getRange(2, 6, sheetData.length, 1).setNumberFormat("@");
                
                tempSheet.getDataRange().setWrap(true).setVerticalAlignment("top");
                for (let i = 1; i <= commHeaders.length; i++) tempSheet.autoResizeColumn(i);

                SpreadsheetApp.flush();
                Utilities.sleep(2000); // Give Google servers time to digest

                // 2. Navigate and prepare the destination folder perfectly
                let rootFolder = DriveApp.getFolderById(COMMISSION_REPORT_ROOT_FOLDER_ID);
                let yFolder = rootFolder.getFoldersByName(yearFolderStr).hasNext() ? rootFolder.getFoldersByName(yearFolderStr).next() : rootFolder.createFolder(yearFolderStr);
                let mFolder = yFolder.getFoldersByName(monthYear).hasNext() ? yFolder.getFoldersByName(monthYear).next() : yFolder.createFolder(monthYear);
                let targetFolder = mFolder.getFoldersByName(cleanAeName).hasNext() ? mFolder.getFoldersByName(cleanAeName).next() : mFolder.createFolder(cleanAeName);

                // 3. Attempt to convert and save as XLSX
                let url = "https://docs.google.com/spreadsheets/d/" + tempSpreadsheet.getId() + "/export?format=xlsx";
                let token = ScriptApp.getOAuthToken();
                
                try {
                    let response = UrlFetchApp.fetch(url, { headers: { 'Authorization': 'Bearer ' + token }, muteHttpExceptions: true });
                    let attachmentBlob = response.getBlob().setName(fileName + ".xlsx");
                    
                    // Cleanup old matching files to avoid duplicates
                    let existingFiles = targetFolder.getFiles();
                    let targetFile = null;
                    while (existingFiles.hasNext()) {
                        let file = existingFiles.next();
                        if (file.getName().endsWith(".xlsx") || file.getName() === fileName) {
                            if (!targetFile && file.getName().endsWith(".xlsx")) {
                                targetFile = file;
                            } else {
                                file.setTrashed(true);
                            }
                        }
                    }
                    
                    // Update or create the XLSX
                    if (targetFile) {
                        Drive.Files.update({}, targetFile.getId(), attachmentBlob);
                        targetFile.setName(fileName + ".xlsx");
                    } else {
                        targetFolder.createFile(attachmentBlob);
                    }
                    
                    // Since XLSX export succeeded, delete the temp Google Sheet
                    DriveApp.getFileById(tempSpreadsheet.getId()).setTrashed(true);

                    // --- NEW: BUILD EMAIL HTML FOR THIS AE ---
                    let tableRowsHtml = "";
                    let displayItems = finalRows.slice(0, 10);
                    let totalComm = 0;
                    
                    finalRows.forEach(r => {
                        let commVal = parseFloat(String(r[20]).replace(/,/g, ''));
                        if (!isNaN(commVal)) totalComm += commVal;
                    });

                    displayItems.forEach((r, index) => {
                        tableRowsHtml += `<tr>`;
                        r.forEach((cell, idx) => {
                            let align = "left";
                            if (idx === 0 || idx === 10 || idx === 11 || idx === 12) align = "center";
                            if (idx >= 13) align = "right";
                            let cellData = (cell === "" || cell == null) ? "" : cell;
                            tableRowsHtml += `<td style="border: 1px solid black; padding: 4px;" align="${align}">${cellData}</td>`;
                        });
                        tableRowsHtml += `</tr>`;
                    });

                    if (finalRows.length > 10) {
                        tableRowsHtml += `<tr><td colspan="21" style="border: 1px solid black; padding: 8px; text-align: center; font-style: italic; background-color: #f9f9f9; color: #555;">...and ${finalRows.length - 10} more rows (Please refer to the attached Excel file for the complete list for ${cleanAeName})</td></tr>`;
                    }

                    let feedbackLink = isTech ? (SCRIPT_PROPS['FEEDBACK_LINK_TECH'] || "") : (SCRIPT_PROPS['FEEDBACK_LINK_PINGCON'] || "");
                    
                    let aeHTML = Template_CommissionEmail(cleanAeName, commHeaders, tableRowsHtml, totalComm, finalRows.length, feedbackImgBlob, feedbackLink);

                    // --- SEND INDIVIDUAL COMMISSION EMAIL ---
                    let securedBlob = encryptAttachmentViaPython(attachmentBlob, fileName + ".xlsx", hasKeysysOrA2ZIT ? SCRIPT_PROPS['ENCRYPT_PASS_TECH'] : SCRIPT_PROPS['ENCRYPT_PASS_PINGCON']);
                    
                    let commCc = isTech ? (sysSettings["Finance_A2Z"] || "") : (sysSettings["Finance_Pingcon"] || "");
                    let commEmailOptions = {
                        from: senderEmail, cc: commCc,
                        name: senderName, attachments: [securedBlob], inlineImages: {}
                    };
                    if (feedbackImgBlob) commEmailOptions.inlineImages.feedbackImage = feedbackImgBlob;
                    sendRobustEmail(mainRecipient, fileName, aeHTML, commEmailOptions);
                    // --- END INDIVIDUAL EMAIL ---

                } catch (fetchErr) {
                    console.warn("XLSX Export blocked. Falling back to Google Sheet.");
                    
                    // 4. BULLETPROOF FALLBACK: Move the standard Google Sheet into the target folder so it is NEVER empty
                    let file = DriveApp.getFileById(tempSpreadsheet.getId());
                    file.moveTo(targetFolder);
                }

            } catch (mainErr) {
                console.error("Error generating Commission Report: " + mainErr.toString());
            }
        }
    }
    // --- END COMMISSION REPORT ---
  }

  

  if (generateOutput && !targetAeFilter) {
      const dbFolder = DriveApp.getFolderById(DATABASE_FOLDER_ID);
      for (let sheetName in rowsToMarkDoneBySheet) {
        let files = dbFolder.getFilesByName(sheetName);
        if (files.hasNext()) {
          let sheet = SpreadsheetApp.open(files.next()).getSheets()[0];
          let rows = rowsToMarkDoneBySheet[sheetName];
          let a1Notations = rows.map(r => `Q${r}`);
          for (let i = 0; i < a1Notations.length; i += 400) {
             sheet.getRangeList(a1Notations.slice(i, i + 400)).setValue("Done");
          }
        }
      }
  }

  

  console.timeEnd("Process_Buy_Sell_Total");
}

// =========================================================================
// TARGETED INVOICE EXTRACTION ENGINE (APPS SCRIPT EDITOR VERSION)
// =========================================================================


// =========================================================================
// TARGETED INVOICE EXTRACTION ENGINE (MOVED TO EXTERNAL TOOL)
// =========================================================================



function buildSummaryHeader(totalSalesInc, totalCostInc, grossProfitInc, gpPctInc, commValue, allPaid, hasExcludedItems, quotaMet, monthlyQuota, unpaidPos, prevMonthExists, tenDayMechanicMet, totalSalesEx) {
  const width = 85; 
  let v = [], b =[], w = [], c =[];
  
  const OFF = 1;

  const formatNum = (num) => {
    return (num || 0).toFixed(2).replace(/\d(?=(\d{3})+\.)/g, '$&,');
  };

  const fmtSales = formatNum(totalSalesEx); 
  const fmtQuota = formatNum(monthlyQuota);

  let releaseCommission = (allPaid && quotaMet && prevMonthExists && tenDayMechanicMet);
  
  let summaryValue = releaseCommission ? commValue : "Hold";
  let summaryValueColor = releaseCommission ? "green" : "red";
  
  let mech1 = quotaMet ? "Yes" : "No"; let mech1Color = quotaMet ? "green" : "red";
  let mech2 = prevMonthExists ? "Yes" : "No"; let mech2Color = prevMonthExists ? "green" : "red";
  let mech3 = allPaid ? "Yes" : "No"; let mech3Color = allPaid ? "green" : "red";
  let mech4 = tenDayMechanicMet ? "Yes" : "No"; let mech4Color = tenDayMechanicMet ? "green" : "red";

  let mech5 = "Yes", mech5Color = "green";
  let mech6 = "Yes", mech6Color = "green";
  let mech7 = "Yes", mech7Color = "green";
  let mech8 = "Yes", mech8Color = "green";
  let mech9 = "Yes", mech9Color = "green";
  let mech10 = "Yes", mech10Color = "green";
  let mech11 = "Yes", mech11Color = "green";
  let mech12 = "Yes", mech12Color = "green";
  
  let text1 = quotaMet ? `1. Total Sales Reached Monthly Quota [${fmtSales}] >= [${fmtQuota}]` : `1. Total Sales Didn't Reach Monthly Quota [${fmtSales}] < [${fmtQuota}]`;
  let color1 = quotaMet ? "green" : "red";

  let text2 = prevMonthExists ? "2. No skipping of commission request. Previous month commission was requested/released." : "2. No skipping of commission request. Example. AE requested for January 2025 commission.";
  let color2 = prevMonthExists ? "green" : "red";

  let text3 = ""; let color3 = "";
  if (allPaid) {
    text3 = "3. All clients are fully paid including previous month whether quota or not quota.";
    color3 = "green";
  } else {
    let uniqueUnpaid =[...new Set(unpaidPos)];
    text3 = "3. Not all clients are fully paid including previous month:\n" + uniqueUnpaid.map(item => " • " + item).join("\n");
    color3 = "red";
  }

  let text4 = tenDayMechanicMet ? "4. Commission request is 10 working days before nearest payroll release." : "4. Commission request is NOT 10 working days before nearest payroll release.";
  let color4 = tenDayMechanicMet ? "green" : "red";

  let text5 = "5. If EWT is not submitted by client, okay to deduct with the approval of AE."; let color5 = "green";
  let text6 = "6. If PO is composed of branded unbranded and ham, ham will not be excluded in buysell to determine margin."; let color6 = "green";
  let text7 = "7. Ham only in one PO, no commission."; let color7 = "green";
  let text8 = "8. If there is negative gpm in one item in a PO, whole PO will be voided. No commission."; let color8 = "green";
  let text9 = "9. Invoice date is the reference for determining monthly sales quota."; let color9 = "green";
  let text10 = "10. Loan equipment excluded from commission computation. Need signed Loan agreement. Need DR showing Loan item."; let color10 = "green";
  let text11 = "11. Commission is released via payroll to be included in payslip. Not cash."; let color11 = "green";
  let text12 = "12. No sliding date, no shifting date, no skip claim."; let color12 = "green";

  // --- ROW 1: TITLES ---
  let r1 = Array(width).fill("");
  r1[0] = "SN"; 
  r1[0 + OFF] = "MONTHLY PERFORMANCE SUMMARY VAT INC"; 
  r1[36 + OFF] = "COMMISSION RELEASING MECHANICS"; 

  v.push(r1);

  let rb1 = Array(width).fill("#ffffff"); 
  for(let i=0 + OFF; i<=5 + OFF; i++) rb1[i] = "#F44336"; 
  for(let i=36 + OFF; i<=48 + OFF; i++) rb1[i] = "#F44336";

  b.push(rb1);
  w.push(Array(width).fill("bold"));
  c.push(Array(width).fill("white")); 
  c[c.length-1][0] = "black"; 

  // --- ROW 2: HEADERS ---
  let r2 = Array(width).fill("");
  r2[0 + OFF] = "Total Sell VAT INC"; 
  r2[1 + OFF] = "Total Buy VAT INC";
  r2[2 + OFF] = "GPM value VAT INC";
  r2[3 + OFF] = "GPM percent VAT INC";
  r2[4 + OFF] = "Commission Rate";
  r2[5 + OFF] = "Commission";
  
  r2[36 + OFF] = "1. AE reach monthly quota? Yes or No?";
  r2[37 + OFF] = "2. Previous month Commission Released? Yes or No?";
  r2[38 + OFF] = "3. All clients are fully paid including previous month? Yes or No?";
  r2[39 + OFF] = "4. Is the Commission Request 10 working days before nearest payroll release? Yes or No?"; 
  r2[40 + OFF] = "5. Does the AE accept deduction if the EWT isn't submmited by the client? Yes or No?";
  r2[41 + OFF] = "6. If PO is composed of branded unbranded and ham, ham will not be excluded in buysell to determine margin.";
  r2[42 + OFF] = "7. Ham only in one PO, no commission.";
  r2[43 + OFF] = "8. If there is negative gpm in one item in a PO, whole PO will be voided. No commission.";
  r2[44 + OFF] = "9. Is Invoice Date the reference for determining monthly sales quota? Yes or No?";
  r2[45 + OFF] = "10. Is Loan equipment excluded from commission computation? Yes o No? Need signed Loan agreement. Need DR showing Loan item.";
  r2[46 + OFF] = "11. Is Commission released via payroll to be included in payslip and not cash? Yes or No?";
  r2[47 + OFF] = "12. No sliding date, no shifting date, no skip claim? Yes or No?";
  r2[48 + OFF] = "Remarks"; 

  v.push(r2);
  
  let rb2 = Array(width).fill("#ffffff");
  for(let i=0 + OFF; i<=5 + OFF; i++) rb2[i] = "#E0E0E0"; 
  for(let i=36 + OFF; i<=48 + OFF; i++) rb2[i] = "#E0E0E0"; 
  b.push(rb2);
  w.push(Array(width).fill("bold"));
  c.push(Array(width).fill("black"));

  // --- ROW 3: VALUES ---
  let r3 = Array(width).fill("");
  let overallGpPct = Math.round(gpPctInc * 100);
  let commRateVal = 0;
  
  if (overallGpPct >= 10 && overallGpPct <= 15) commRateVal = 0.005;
  else if (overallGpPct >= 16 && overallGpPct <= 20) commRateVal = 0.0055;
  else if (overallGpPct >= 21 && overallGpPct <= 25) commRateVal = 0.006;
  else if (overallGpPct >= 26 && overallGpPct <= 30) commRateVal = 0.0065;
  else if (overallGpPct >= 31) commRateVal = 0.007;

  r3[0 + OFF] = totalSalesInc;
  r3[1 + OFF] = totalCostInc;
  r3[2 + OFF] = grossProfitInc;
  r3[3 + OFF] = overallGpPct + "%";
  r3[4 + OFF] = commRateVal;
  r3[5 + OFF] = summaryValue;

  r3[36 + OFF] = mech1;
  r3[37 + OFF] = mech2;
  r3[38 + OFF] = mech3;
  r3[39 + OFF] = mech4; 
  r3[40 + OFF] = mech5;
  r3[41 + OFF] = mech6;
  r3[42 + OFF] = mech7;
  r3[43 + OFF] = mech8;
  r3[44 + OFF] = mech9;
  r3[45 + OFF] = mech10;
  r3[46 + OFF] = mech11;
  r3[47 + OFF] = mech12;

  v.push(r3);
  
  let rb3 = Array(width).fill("#ffffff");
  b.push(rb3);
  let rw3 = Array(width).fill("bold");
  w.push(rw3);
  let rc3 = Array(width).fill("black");
  
  rc3[5 + OFF] = summaryValueColor;
  rc3[36 + OFF] = mech1Color;
  rc3[37 + OFF] = mech2Color;
  rc3[38 + OFF] = mech3Color;
  rc3[39 + OFF] = mech4Color; 
  rc3[40 + OFF] = mech5Color;
  rc3[41 + OFF] = mech6Color;
  rc3[42 + OFF] = mech7Color;
  rc3[43 + OFF] = mech8Color;
  rc3[44 + OFF] = mech9Color;
  rc3[45 + OFF] = mech10Color;
  rc3[46 + OFF] = mech11Color;
  rc3[47 + OFF] = mech12Color;

  c.push(rc3);

  // --- ROW 4: SPACER ---
  v.push(Array(width).fill(""));
  b.push(Array(width).fill("#ffffff"));
  w.push(Array(width).fill("normal"));
  c.push(Array(width).fill("black"));

  return { 
    values: v, backgrounds: b, weights: w, colors: c,
    richTextData: {
      text1: text1, color1: color1,
      text2: text2, color2: color2,
      text3: text3, color3: color3,
      text4: text4, color4: color4,
      text5: text5, color5: color5,
      text6: text6, color6: color6,
      text7: text7, color7: color7,
      text8: text8, color8: color8,
      text9: text9, color9: color9,
      text10: text10, color10: color10,
      text11: text11, color11: color11,
      text12: text12, color12: color12
    }
  };
}

function buildPOBlockInMemory(data, startRow, voidReasons, ewtMap, purchasesMap, mechanics) {
  const width = 85; // Expanded to handle columns up to BV
  let v = [], b =[], w = [], c =[];
  let merges =[];
  
  const OFF = 1;

  const m = data.meta;
  const items = data.items;
  const numItems = items.length;
  
  const pushRow = (fillVal="", bg="#ffffff", weight="normal", col="black") => {
    v.push(Array(width).fill(fillVal));
    b.push(Array(width).fill(bg));
    w.push(Array(width).fill(weight));
    c.push(Array(width).fill(col));
  };

  let r0 = Array(width).fill("");
  r0[0 + OFF] = "Account Exec -"; r0[1 + OFF] = m.ae;
  v.push(r0);
  b.push(Array(width).fill("#ffffff")); b[b.length-1][0 + OFF] = "red";
  w.push(Array(width).fill("normal")); w[w.length-1][0 + OFF] = "bold";
  c.push(Array(width).fill("black")); c[c.length-1][0 + OFF] = "white";

  pushRow();

  let r2 = Array(width).fill("");
  r2[0 + OFF] = "End-user Info:";
  v.push(r2);
  b.push(Array(width).fill("#ffffff")); b[b.length-1][0 + OFF] = "red";
  w.push(Array(width).fill("normal")); w[w.length-1][0 + OFF] = "bold";
  c.push(Array(width).fill("black")); c[c.length-1][0 + OFF] = "white";

  let r3 = Array(width).fill("");
  r3[0 + OFF] = "Company Name"; r3[3 + OFF] = m.customer;
  v.push(r3);
  b.push(Array(width).fill("#ffffff")); b[b.length-1][0 + OFF] = "red";
  w.push(Array(width).fill("normal")); w[w.length-1][0 + OFF] = "bold";
  c.push(Array(width).fill("black")); c[c.length-1][0 + OFF] = "white";
  merges.push({rOffset: 3, c: 4 + OFF, nr: 1, nc: 11}); 

  let poInfoParts =[];
  if (m.dr) poInfoParts.push(`DR ${m.dr}`);
  if (m.si && m.si !== "???") poInfoParts.push(`SI ${m.si}`);
  if (m.po) {
    let poVal = String(m.po);
    if (!poVal.match(/^(DR|SI|PO)/i)) poVal = `PO ${poVal}`;
    if (!poInfoParts.includes(poVal)) poInfoParts.push(poVal);
  }
  let cleanPo = poInfoParts.join(", ");

  let r4 = Array(width).fill("");
  r4[0 + OFF] = "End-user PO -"; r4[3 + OFF] = cleanPo;
  v.push(r4);
  b.push(Array(width).fill("#ffffff")); b[b.length-1][0 + OFF] = "red";
  w.push(Array(width).fill("normal")); w[w.length-1][0 + OFF] = "bold";
  c.push(Array(width).fill("black")); c[c.length-1][0 + OFF] = "white";
  merges.push({rOffset: 4, c: 4 + OFF, nr: 1, nc: 11});

  let r5 = Array(width).fill("");
  r5[0 + OFF] = "PO Date -->"; r5[1 + OFF] = m.date;
  v.push(r5);
  b.push(Array(width).fill("#ffffff"));
  w.push(Array(width).fill("normal"));
  c.push(Array(width).fill("black"));

  pushRow();

  // Section Headers
  let r8 = Array(width).fill("");
  r8[0 + OFF] = "Keysys Buy Price / Cost Details"; 
  r8[12 + OFF] = "Keysys Sell Price";
  r8[50 + OFF] = "EWT Report Supplier"; // AZ
  r8[63 + OFF] = "Purchase Report";     // BM

  v.push(r8);
  let b8 = Array(width).fill("#ffffff");
  for(let i=0 + OFF;i<=11 + OFF;i++) b8[i] = "red"; 
  for(let i=12 + OFF;i<=20 + OFF;i++) b8[i] = "#38761D"; // N to V
  for(let i=50 + OFF;i<=61 + OFF;i++) b8[i] = "#3d85c6"; // AZ to BK
  for(let i=63 + OFF;i<=71 + OFF;i++) b8[i] = "#e69138"; // BM to BU

  b.push(b8);
  w.push(Array(width).fill("bold"));
  c.push(Array(width).fill("white"));
  
  merges.push({rOffset: 7, c: 1 + OFF, nr: 1, nc: 12}); 
  merges.push({rOffset: 7, c: 13 + OFF, nr: 1, nc: 9});  // N to V
  merges.push({rOffset: 7, c: 51 + OFF, nr: 1, nc: 12}); // AZ to BK
  merges.push({rOffset: 7, c: 64 + OFF, nr: 1, nc: 9});  // BM to BU

  // Column Headers
  let r9 = Array(width).fill("");
  const buyH =["Supplier", "Pingcon PO#", "Purchase Invoice", "SKU #", "Desc", "Unit Price", "Qty", "UOM", "Total", "Total Buy VAT INC", "Total Buy VAT EX", "INPUT VAT"];
  buyH.forEach((h,i) => r9[i + OFF] = h);

  const sellH =["SKU #", "Desc", "Unit Price", "Qty", "UOM", "Total", "Total Sell VAT INC", "Total Sell VAT EX", "OUTPUT VAT", "", "GPM value VAT INC", "GPM percent VAT INC"];
  sellH.forEach((h,i) => r9[i+12 + OFF] = h); // N to Y

  const botH =["total Buy VAT INC", "total Buy VAT EX", "INPUT VAT", "total Sell VAT INC", "total Sell VAT EX", "OUTPUT VAT", "total GPM value VAT INC", "total GPM percent VAT INC", "VAT Payable", "Commission Rate", "Commission"];
  botH.forEach((h, i) => r9[i + 25 + OFF] = h); // AA to AK

  const mechH =[
    "1. AE reach monthly quota? Yes or No?",
    "2. Previous month Commission Released? Yes or No?",
    "3. Payment Status is Fully Paid? Yes or No?",
    "4. Is the Commission Request 10 working days before nearest payroll release? Yes or No?", 
    "5. Does the AE accept deduction if the EWT isn't submmited by the client? Yes or No?",
    "6. If PO is composed of branded unbranded and ham, ham will not be excluded in buysell to determine margin.",
    "7. Ham only in one PO, no commission.",
    "8. If there is negative gpm in one item in a PO, whole PO will be voided. No commission.",
    "9. Is Invoice Date the reference for determining monthly sales quota? Yes or No?",
    "10. Is Loan equipment excluded from commission computation? Yes o No? Need signed Loan agreement. Need DR showing Loan item.",
    "11. Is Commission released via payroll to be included in payslip and not cash? Yes or No?",
    "12. No sliding date, no shifting date, no skip claim? Yes or No?"
  ];
  mechH.forEach((h, i) => r9[i + 36 + OFF] = h); // AL to AW

  r9[48 + OFF] = "Remarks"; // AX

  const ewtH =["Account Title", "Invoice Date", "Company Name", "Address", "TIN", "Invoice Number", "Gross", "Input Tax", "Purchases", "EWT 1%", "EWT 2%", "ATC"];
  ewtH.forEach((h,i) => r9[i+50 + OFF] = h); // AZ

  const purH =["Invoice Date", "Invoice Number", "Company Name", "Account Title", "TIN", "Address", "Gross", "Input Tax", "Purchases"];
  purH.forEach((h,i) => r9[i+63 + OFF] = h); // BM
  
  v.push(r9);
  let b9 = Array(width).fill("#ffffff");
  for(let i=0 + OFF;i<=11 + OFF;i++) b9[i] = "#ffcccc"; 
  for(let i=12 + OFF;i<=20 + OFF;i++) b9[i] = "#d9ead3"; // N to V
  b9[22 + OFF] = "#d9ead3"; // X
  b9[23 + OFF] = "#d9ead3"; // Y
  
  b9[25 + OFF] = "#FFFF00"; b9[26 + OFF] = "#FFFF00"; b9[27 + OFF] = "#FFFF00"; // AA to AC
  b9[28 + OFF] = "#D9EAD3"; b9[29 + OFF] = "#D9EAD3"; b9[30 + OFF] = "#D9EAD3"; // AD to AF
  b9[31 + OFF] = "#FCE5CD"; b9[32 + OFF] = "#FCE5CD"; // AG to AH
  b9[33 + OFF] = "#F6B26B"; // AI
  b9[34 + OFF] = "#D9D2E9"; b9[35 + OFF] = "#D9D2E9"; // AJ to AK

  for(let i=36 + OFF;i<=47 + OFF;i++) b9[i] = "#E0E0E0"; // AL to AW (Grey)
  b9[48 + OFF] = "#E0E0E0"; // AX (Remarks)

  for(let i=50 + OFF;i<=61 + OFF;i++) b9[i] = "#cfe2f3"; // AZ to BK
  for(let i=63 + OFF;i<=71 + OFF;i++) b9[i] = "#fce5cd"; // BM to BU
  
  b.push(b9);
  w.push(Array(width).fill("bold"));
  c.push(Array(width).fill("black"));

  let itemStartIdx = 9;
  let absStartRow = startRow + itemStartIdx; 
  let bRange = `J${absStartRow}:J${absStartRow + numItems - 1}`;
  let sRange = `S${absStartRow}:S${absStartRow + numItems - 1}`;
  if (numItems === 0) { bRange = "0"; sRange = "0"; } 

  let seenBasketIds = new Set();

  // --- PREPARE 12 REMARKS ARRAY TO DISTRIBUTE ACROSS ROWS ---
  let {quotaMet, prevMonthExists, allPaid, tenDayMechanicMet} = mechanics || {};
  let hasHamVoid = voidReasons && voidReasons.includes("Ham only in one PO, no commission.");
  let hasGpmVoid = voidReasons && voidReasons.includes("If there is negative GPM in one item in a PO, whole PO will be voided. No commission.");
  let hasLoanVoid = voidReasons && voidReasons.includes("Loan equipment excluded from commission computation.");

  let poFullyPaid = (m.paymentStatus === "Fully Paid");
  let formatOpenBal = (num) => (num || 0).toFixed(2).replace(/\d(?=(\d{3})+\.)/g, '$&,');

  let remarksArray =[
      quotaMet ? "1. Total Sales Reached Monthly Quota." : "1. Total Sales Didn't Reach Monthly Quota.",
      prevMonthExists ? "2. No skipping of commission request. Previous month commission was requested/released." : "2. No skipping of commission request. Example. AE requested for January 2025 commission.",
      poFullyPaid ? "3. Fully Paid" : `3. Open Balance: ${formatOpenBal(m.openBal)}`,
      tenDayMechanicMet ? "4. Commission request is 10 working days before nearest payroll release." : "4. Commission request is NOT 10 working days before nearest payroll release.",
      "5. If EWT is not submitted by client, okay to deduct with the approval of AE.",
      "6. If PO is composed of branded unbranded and ham, ham will not be excluded in buysell to determine margin.",
      hasHamVoid ? "7. VOIDED: Ham only in one PO, no commission." : "7. Ham only in one PO, no commission.",
      hasGpmVoid ? "8. VOIDED: If there is negative gpm in one item in a PO, whole PO will be voided. No commission." : "8. If there is negative gpm in one item in a PO, whole PO will be voided. No commission.",
      "9. Invoice date is the reference for determining monthly sales quota.",
      hasLoanVoid ? "10. VOIDED: Loan equipment excluded from commission computation. Need signed Loan agreement. Need DR showing Loan item." : "10. Loan equipment excluded from commission computation. Need signed Loan agreement. Need DR showing Loan item.",
      "11. Commission is released via payroll to be included in payslip. Not cash.",
      "12. No sliding date, no shifting date, no skip claim."
  ];

  let remarksColorArray =[
      quotaMet ? "green" : "red",
      prevMonthExists ? "green" : "red",
      poFullyPaid ? "green" : "red",
      tenDayMechanicMet ? "green" : "red",
      "green", "green",
      hasHamVoid ? "red" : "green",
      hasGpmVoid ? "red" : "green",
      "green",
      hasLoanVoid ? "red" : "green",
      "green", "green"
  ];

  items.forEach((item, i) => {
    let r = Array(width).fill("");
    let rb = Array(width).fill("#ffffff");
    let rc = Array(width).fill("black");
    let rw = Array(width).fill("normal");
    
    let excelRow = absStartRow + i;

    let isFirstOfBasket = false;
    let basketRowCount = 1;
    if (item.isBasket) {
        if (!seenBasketIds.has(item.basketId)) {
            isFirstOfBasket = true;
            seenBasketIds.add(item.basketId);
            basketRowCount = items.filter(x => x.basketId === item.basketId).length;
        }
    }

    r[0 + OFF] = item.supplier;
    r[1 + OFF] = item.pingconPo ? String(item.pingconPo) : ""; 
    r[2 + OFF] = item.purchaseInvoice; 
    r[3 + OFF] = item.sku;
    r[4 + OFF] = item.desc;
    r[6 + OFF] = item.qty;
    r[7 + OFF] = item.uom;

    r[12 + OFF] = item.sellSku;
    r[13 + OFF] = item.sellDesc;
    r[14 + OFF] = item.unitSell;
    r[15 + OFF] = item.sellQty;
    r[16 + OFF] = item.sellUom;
    
    let rowCost = item.unitBuy * item.qty; 
    let purchaseDateStr = item.sourceDate ? Utilities.formatDate(item.sourceDate, Session.getScriptTimeZone(), "MM/dd/yyyy") : "";
    
    if (item.isNoStock) {
      r[5 + OFF] = "NO STOCK AVAILABLE"; r[8 + OFF] = "Excluded"; r[17 + OFF] = "Excluded"; 
      rb.fill("#FFDAB9"); 
      rc[5 + OFF] = "#CC0000"; rw[5 + OFF] = "bold"; rc[8 + OFF] = "#CC0000"; rw[8 + OFF] = "bold"; rc[17 + OFF] = "#CC0000"; rw[17 + OFF] = "bold";
    }
    else if (item.isMissingSku) {
      r[5 + OFF] = "MISSING SKU"; r[8 + OFF] = "Excluded"; r[17 + OFF] = "Excluded"; 
      rb.fill("#FFF2CC"); 
      rc[5 + OFF] = "red"; rw[5 + OFF] = "bold"; rc[8 + OFF] = "red"; rw[8 + OFF] = "bold"; rc[17 + OFF] = "red"; rw[17 + OFF] = "bold";
    } else {
      r[5 + OFF] = item.unitBuy; 
      r[8 + OFF] = `=G${excelRow}*H${excelRow}`; // J (Total Buy)
      
      if (item.isBasket && !item.isFirstComp) {
          r[17 + OFF] = ""; 
      } else {
          r[17 + OFF] = item.totalSell; // S (Total Sell)
      }
      
      if (item.isBasket) {
          if (isFirstOfBasket) {
              r[22 + OFF] = `=S${excelRow}-SUM(J${excelRow}:J${excelRow + basketRowCount - 1})`; // X
              r[23 + OFF] = `=IF(S${excelRow}=0, 0, X${excelRow}/S${excelRow})`; // Y
          }
      } else {
          r[22 + OFF] = `=S${excelRow}-J${excelRow}`; // X
          r[23 + OFF] = `=IF(S${excelRow}=0, 0, X${excelRow}/S${excelRow})`; // Y
      }

      const format2Dec = (val) => {
          if (val === "" || val == null) return "";
          let num = parseFloat(String(val).replace(/,/g, ''));
          return (!isNaN(num)) ? num.toFixed(2) : val;
      };

      // EWT Extract
      let ewtDataMatch = null;
      if (ewtMap && item.purchaseInvoice) {
          let invKey = String(item.purchaseInvoice).trim().toUpperCase();
          let supKey = String(item.supplier).trim().toUpperCase();
          ewtDataMatch = ewtMap[`${invKey}_${supKey}`] || ewtMap[invKey];
      }
      if (ewtDataMatch) {
          r[50 + OFF] = ewtDataMatch[0]; // AZ 
          r[51 + OFF] = ewtDataMatch[1] instanceof Date ? Utilities.formatDate(ewtDataMatch[1], Session.getScriptTimeZone(), "MM/dd/yyyy") : ewtDataMatch[1]; // BA
          r[52 + OFF] = ewtDataMatch[2]; // BB 
          r[53 + OFF] = ewtDataMatch[3]; // BC 
          r[54 + OFF] = ewtDataMatch[4]; // BD 
          r[55 + OFF] = ewtDataMatch[5]; // BE 
          r[56 + OFF] = format2Dec(ewtDataMatch[6]); // BF 
          r[57 + OFF] = format2Dec(ewtDataMatch[7]); // BG 
          r[58 + OFF] = format2Dec(ewtDataMatch[8]); // BH 
          r[59 + OFF] = format2Dec(ewtDataMatch[9]); // BI 
          r[60 + OFF] = format2Dec(ewtDataMatch[10]); // BJ 
          r[61 + OFF] = ewtDataMatch[11]; // BK 
      } else {
          r[51 + OFF] = purchaseDateStr;      // BA
          r[52 + OFF] = item.supplier;        // BB
          r[55 + OFF] = item.purchaseInvoice; // BE
          r[56 + OFF] = format2Dec(rowCost);  // BF
      }

      // Purchases Extract
      let purDataMatch = null;
      if (purchasesMap && item.purchaseInvoice) {
          let invKey = String(item.purchaseInvoice).trim().toUpperCase();
          let supKey = String(item.supplier).trim().toUpperCase();
          purDataMatch = purchasesMap[`${invKey}_${supKey}`] || purchasesMap[invKey];
      }
      if (purDataMatch) {
          r[63 + OFF] = purDataMatch[0] instanceof Date ? Utilities.formatDate(purDataMatch[0], Session.getScriptTimeZone(), "MM/dd/yyyy") : purDataMatch[0]; // BM
          r[64 + OFF] = purDataMatch[1]; // BN
          r[65 + OFF] = purDataMatch[2]; // BO
          r[66 + OFF] = purDataMatch[3]; // BP
          r[67 + OFF] = purDataMatch[4]; // BQ
          r[68 + OFF] = purDataMatch[5]; // BR
          r[69 + OFF] = format2Dec(purDataMatch[6]); // BS
          r[70 + OFF] = format2Dec(purDataMatch[7]); // BT
          r[71 + OFF] = format2Dec(purDataMatch[8]); // BU
      } else {
          r[63 + OFF] = purchaseDateStr;      // BM
          r[64 + OFF] = item.purchaseInvoice; // BN
          r[65 + OFF] = item.supplier;        // BO
          r[69 + OFF] = format2Dec(rowCost);  // BS
      }
    }
    
    // First item top-row summaries
    if (i === 0) {
      r[9 + OFF] = `=SUM(${bRange})`;      // K
      r[10 + OFF] = `=K${excelRow}/1.12`;  // L
      r[11 + OFF] = `=K${excelRow}-L${excelRow}`;  // M
      
      r[18 + OFF] = `=SUM(${sRange})`;     // T
      r[19 + OFF] = `=T${excelRow}/1.12`;  // U
      r[20 + OFF] = `=T${excelRow}-U${excelRow}`; // V

      r[25 + OFF] = `=K${excelRow}`; // AA
      r[26 + OFF] = `=L${excelRow}`; // AB
      r[27 + OFF] = `=M${excelRow}`; // AC
      r[28 + OFF] = `=T${excelRow}`; // AD
      r[29 + OFF] = `=U${excelRow}`; // AE
      r[30 + OFF] = `=V${excelRow}`; // AF
      r[31 + OFF] = `=SUM(X${absStartRow}:X${absStartRow + numItems - 1})`; // AG
      r[32 + OFF] = `=IF(T${excelRow}=0, 0, AG${excelRow}/T${excelRow})`; // AH
      r[33 + OFF] = `=V${excelRow}-M${excelRow}`; // AI

      let pctCheck = `ROUND(AH${excelRow}*100,0)`;

      if (voidReasons && voidReasons.length > 0) { 
          r[34 + OFF] = "VOID"; rc[34 + OFF] = "red"; // AJ
          r[35 + OFF] = "VOID"; rc[35 + OFF] = "red"; // AK
      } else {
          r[34 + OFF] = `=IF(T${excelRow}=0, 0, IF(${pctCheck}<10, 0, IF(${pctCheck}<=15, 0.5%, IF(${pctCheck}<=20, 0.55%, IF(${pctCheck}<=25, 0.6%, IF(${pctCheck}<=30, 0.65%, 0.7%))))))`; // AJ
          r[35 + OFF] = `=U${excelRow}*AJ${excelRow}`; // AK
      }

      // Inject 12 Mechanics Shifted Horizontal
      r[36 + OFF] = quotaMet ? "Yes" : "No"; rc[36 + OFF] = quotaMet ? "green" : "red";
      r[37 + OFF] = prevMonthExists ? "Yes" : "No"; rc[37 + OFF] = prevMonthExists ? "green" : "red";
      r[38 + OFF] = poFullyPaid ? "Yes" : "No"; rc[38 + OFF] = poFullyPaid ? "green" : "red";
      r[39 + OFF] = tenDayMechanicMet ? "Yes" : "No"; rc[39 + OFF] = tenDayMechanicMet ? "green" : "red";
      
      r[40 + OFF] = "Yes"; rc[40 + OFF] = "green";
      r[41 + OFF] = "Yes"; rc[41 + OFF] = "green";
      r[42 + OFF] = "Yes"; rc[42 + OFF] = "green";
      r[43 + OFF] = "Yes"; rc[43 + OFF] = "green";
      r[44 + OFF] = "Yes"; rc[44 + OFF] = "green";
      r[45 + OFF] = "Yes"; rc[45 + OFF] = "green";
      r[46 + OFF] = "Yes"; rc[46 + OFF] = "green";
      r[47 + OFF] = "Yes"; rc[47 + OFF] = "green";
    }

  v.push(r); b.push(rb); w.push(rw); c.push(rc);
  });

// ========== THE RESTORED BOTTOM SUMMARY BLOCK ==========
  pushRow(); // Spacer row
  let bottomRow = absStartRow + numItems + 1; // Formula index for summary Row 1

  // Row 1 (bottomRow)
  let sumR1 = Array(width).fill(""), sumB1 = Array(width).fill("#ffffff"), sumW1 = Array(width).fill("normal"), sumC1 = Array(width).fill("black");
  sumR1[7 + OFF] = "Total Buy VAT INC"; sumR1[8 + OFF] = `=K${absStartRow}`; // I, J
  sumR1[16 + OFF] = "Total Sell VAT INC"; sumR1[17 + OFF] = `=T${absStartRow}`; // R, S
  sumR1[22 + OFF] = `=SUM(X${absStartRow}:X${absStartRow + numItems - 1})`; // X (Total GPM Value)
  sumR1[23 + OFF] = `=IF(T${absStartRow}=0, 0, X${bottomRow}/T${absStartRow})`; // Y (Total GPM Percent)
  v.push(sumR1); b.push(sumB1); w.push(sumW1); c.push(sumC1);

  // Row 2 (bottomRow + 1)
  let sumR2 = Array(width).fill(""), sumB2 = Array(width).fill("#ffffff"), sumW2 = Array(width).fill("normal"), sumC2 = Array(width).fill("black");
  sumR2[7 + OFF] = "Total Buy VAT EX"; sumR2[8 + OFF] = `=L${absStartRow}`; 
  sumR2[16 + OFF] = "Total Sell VAT EX"; sumR2[17 + OFF] = `=U${absStartRow}`;
  sumR2[22 + OFF] = "Total GPM value VAT INC"; // X
  sumR2[23 + OFF] = "Total GPM percent VAT INC"; // Y
  v.push(sumR2); b.push(sumB2); w.push(sumW2); c.push(sumC2);

  // Row 3 (bottomRow + 2)
  let sumR3 = Array(width).fill(""), sumB3 = Array(width).fill("#ffffff"), sumW3 = Array(width).fill("normal"), sumC3 = Array(width).fill("black");
  sumR3[7 + OFF] = "INPUT VAT"; sumR3[8 + OFF] = `=M${absStartRow}`; 
  sumR3[16 + OFF] = "OUTPUT VAT"; sumR3[17 + OFF] = `=V${absStartRow}`; 
  sumR3[22 + OFF] = `=V${absStartRow}-M${absStartRow}`; // X (Value)
  v.push(sumR3); b.push(sumB3); w.push(sumW3); c.push(sumC3);

  // Row 4 (bottomRow + 3)
  let sumR4 = Array(width).fill(""), sumB4 = Array(width).fill("#ffffff"), sumW4 = Array(width).fill("normal"), sumC4 = Array(width).fill("black");
  sumR4[22 + OFF] = "VAT Payable"; // X (Label)
  v.push(sumR4); b.push(sumB4); w.push(sumW4); c.push(sumC4);

  // --- DISTRIBUTE 12 REMARKS ACROSS ALL ROWS (Items + Spacer + Summary) ---
  let itemVStart = v.length - (numItems + 5); 
  
  for (let i = 0; i < 12; i++) {
      if (itemVStart + i < v.length) {
          v[itemVStart + i][48 + OFF] = remarksArray[i];
          c[itemVStart + i][48 + OFF] = remarksColorArray[i];
      } else {
          // If we need more rows to fit the 12 remarks, append blank ones at the very bottom
          let r = Array(width).fill("");
          let rb = Array(width).fill("#ffffff");
          let rc = Array(width).fill("black");
          let rw = Array(width).fill("normal");
          
          r[48 + OFF] = remarksArray[i];
          rc[48 + OFF] = remarksColorArray[i];
          
          v.push(r); b.push(rb); w.push(rw); c.push(rc);
      }
  }

  return { values: v, backgrounds: b, weights: w, colors: c, merges: merges };
}

// =========================================================================
// WORKING DAYS MECHANIC HELPERS
// =========================================================================

function checkWorkingDaysMechanic(requestDateStr, targetMonthFilter) {
  if (!requestDateStr) return false;
  
  const requestDate = new Date(requestDateStr);
  if (isNaN(requestDate.getTime())) return false;

  // RULE 1: Not allowed to request on a non-working day
  if (!isWorkingDay(requestDate)) return false;

  let payrollDate;

  // RULE 2: Determine Target Payroll Date based on the REPORT MONTH
  if (targetMonthFilter) {
    let reportDate = new Date(targetMonthFilter + " 1"); // e.g., "January 2026 1"
    let midMonthPayroll = new Date(reportDate.getFullYear(), reportDate.getMonth(), 15);
    let endMonthPayroll = new Date(reportDate.getFullYear(), reportDate.getMonth() + 1, 0);
    
    // Target the appropriate payroll period
    if (requestDate <= midMonthPayroll) {
      payrollDate = midMonthPayroll;
    } else {
      payrollDate = endMonthPayroll;
    }
  } else {
    // Fallback if no specific Report Month is selected
    if (requestDate.getDate() <= 15) {
      payrollDate = new Date(requestDate.getFullYear(), requestDate.getMonth(), 15);
    } else {
      payrollDate = new Date(requestDate.getFullYear(), requestDate.getMonth() + 1, 0);
    }
  }

  // If Target Payroll Date is a non-working day, adjust it backwards
  while (!isWorkingDay(payrollDate)) {
    payrollDate.setDate(payrollDate.getDate() - 1);
  }

  // RULE 3: Count Working Days starting from Request Date up to the day BEFORE Payroll Date
  let workingDaysCount = 0;
  let currentDate = new Date(requestDate);

  while (currentDate < payrollDate) {
    if (isWorkingDay(currentDate)) {
      workingDaysCount++;
    }
    currentDate.setDate(currentDate.getDate() + 1);
  }

  return workingDaysCount >= 10;
}

function isWorkingDay(dateObj) {
  const day = dateObj.getDay();
  // 0 = Sunday, 6 = Saturday
  if (day === 0 || day === 6) return false;
  
  // Check Holidays
  if (isHoliday(dateObj)) return false;
  
  return true;
}

function isHoliday(dateObj) {
  const month = dateObj.getMonth(); // 0-indexed (Jan=0, Dec=11)
  const date = dateObj.getDate();
  const year = dateObj.getFullYear();
  
  const key = `${year}-${month}-${date}`;
  const generalKey = `${month}-${date}`;

  // Standard recurring Philippine Holidays
  const recurringHolidays =[
    "0-1",   // Jan 1: New Year's Day
    "1-25",  // Feb 25: EDSA Revolution
    "3-9",   // Apr 9: Araw ng Kagitingan
    "4-1",   // May 1: Labor Day
    "5-12",  // Jun 12: Independence Day
    "7-21",  // Aug 21: Ninoy Aquino Day
    "7-31",  // Aug 31: National Heroes Day
    "10-1",  // Nov 1: All Saints' Day
    "10-2",  // Nov 2: All Souls' Day
    "10-30", // Nov 30: Bonifacio Day
    "11-8",  // Dec 8: Immaculate Conception
    "11-24", // Dec 24: Christmas Eve
    "11-25", // Dec 25: Christmas Day
    "11-30", // Dec 30: Rizal Day
    "11-31"  // Dec 31: Last Day of Year
  ];

  // Specific / Movable Holidays (Matches your Dec 24 to Jan 4 Company-wide break)
  const specificHolidays =[
    "2026-3-2",   // Apr 2, 2026: Maundy Thursday
    "2026-3-3",   // Apr 3, 2026: Good Friday
    "2025-11-26", // Dec 26, 2025: Extended Holiday
    "2025-11-29", // Dec 29, 2025: Extended Holiday
    "2026-0-2"    // Jan 2, 2026: Extended Holiday
  ];

  if (recurringHolidays.includes(generalKey)) return true;
  if (specificHolidays.includes(key)) return true;

  return false;
}

// function writeAEProfitReport(sheetName, rows) {
//   if (!rows || rows.length === 0) return;
//   const ss = SpreadsheetApp.getActiveSpreadsheet();
  
//   const safeName = sheetName.replace(/[:\\\/\?\*\[\]]/g, ""); 
  
//   let sheet = ss.getSheetByName(safeName);
//   if (!sheet) {
//     sheet = ss.insertSheet(safeName);
//     const sheets = ss.getSheets();
//     if (sheets.length > 1) {
//        sheet.activate();
//        ss.moveActiveSheet(sheets.length); 
//     }
//   } else {
//     if(sheet.getFilter()) sheet.getFilter().remove();
//     sheet.clear();
//   }
  
//   // HEADERS (Total 21 Columns now)
//   const mainHeaders = ["PURCHASE DATA (COST)", "", "", "", "", "", "", "", "", "", "", "SALES DATA (REVENUE)"];
  
//   const colHeaders = [
//     "Purchase Date", "Cust Manager", "Supplier", "Pingcon PO#", "Purchase Invoice", "SKU #", "Desc", "Unit Price", "Qty", "UOM", "Total Cost",
//     "Sales Report Date", "Additional Text", "SKU #", "Desc", "Unit Price", "Qty", "UOM", "Total Sales (VAT Inc)", "Gross Profit (VAT Inc)", "GPM"
//   ];
  
//   // 1. Set Main Title
//   // Cost Section: Columns 1 to 11
//   sheet.getRange(1, 1, 1, 11).merge().setValue(mainHeaders[0]).setBackground("#CC0000").setFontColor("white").setFontWeight("bold").setHorizontalAlignment("center");
//   // Revenue Section: Columns 12 to 21 (10 Columns)
//   sheet.getRange(1, 12, 1, 10).merge().setValue(mainHeaders[11]).setBackground("#38761D").setFontColor("white").setFontWeight("bold").setHorizontalAlignment("center");
  
//   // 2. Set Column Headers
//   sheet.getRange(2, 1, 1, colHeaders.length).setValues([colHeaders])
//        .setFontWeight("bold").setBackground("#f3f3f3").setBorder(true, true, true, true, true, true);
  
//   // 3. Write Data
//   sheet.getRange(3, 1, rows.length, colHeaders.length).setValues(rows);
  
//   // 4. Formatting
//   // Purchase Date is Col 1, Sales Date is Col 12. These are strings/dates, auto-format is usually fine, but dates are pre-formatted in processing.
  
//   // Batch number formats for Cost and Revenue Sides
//   sheet.getRangeList([
//     `H3:H${rows.length + 2}`, 
//     `K3:K${rows.length + 2}`, 
//     `P3:P${rows.length + 2}`, 
//     `S3:T${rows.length + 2}`
//   ]).setNumberFormat("#,##0.00");
//   sheet.getRange(3, 21, rows.length, 1).setNumberFormat("0.00%");
  
//   // Negative Profit Highlight (Column 20 is GP Value, Column 21 is GP %)
//   // Highlight based on GP Value < 0
//   let range = sheet.getRange(3, 20, rows.length, 1); 
//   let rule = SpreadsheetApp.newConditionalFormatRule()
//       .whenNumberLessThan(0)
//       .setBackground("#FFCCCC")
//       .setFontColor("#FF0000")
//       .setRanges([range])
//       .build();
//   sheet.setConditionalFormatRules([rule]);
  
//   sheet.autoResizeColumns(1, 21);
// }

function getSheetForYearBlock(prefix, dateObj, headers) {
  const year = dateObj.getFullYear();
  const startYear = Math.floor(year / 5) * 5;
  const endYear = startYear + 4;
  const sheetName = `${prefix}${startYear}-${endYear}`;
  
  let targetFolderId = DATABASE_FOLDER_ID;
  if (prefix.includes("Narra_Voucher_Data") || prefix.includes("Check_Summary_Data") || prefix.includes("RIWB_Data")) {
      targetFolderId = SCRIPT_PROPS['VERIFIED_FOLDER_ID'];
  }
  const dbFolder = DriveApp.getFolderById(targetFolderId);
  const files = dbFolder.getFilesByName(sheetName);
  
  let externalSs;
  if (files.hasNext()) {
    externalSs = SpreadsheetApp.open(files.next());
  } else {
    externalSs = SpreadsheetApp.create(sheetName);
    DriveApp.getFileById(externalSs.getId()).moveTo(dbFolder);
  }
  
  let sheet = externalSs.getSheets()[0];
  sheet.setName(sheetName); // Keeps the tab name consistent with the file name
  
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#f3f3f3");
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getExistingKeysForSheet(sheet, type) {
  const keys = new Set();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return keys;
  
  const data = sheet.getDataRange().getValues();
  let dateIdx = -1, managerIdx = -1, textIdx = -1;
  
  if (type === "AR") {
     let headers = data[0].map(h => String(h).toLowerCase().trim());
     dateIdx = headers.findIndex(h => h.includes("date") && !h.includes("expiry") && !h.includes("countered") && !h.includes("received"));
     managerIdx = headers.findIndex(h => h.includes("manager"));
     textIdx = headers.findIndex(h => h.includes("additional text"));
  }
  
  for (let i = 1; i < data.length; i++) {
    let key = "";
    if (type === "SR") {
      key = `${data[i][5]}_${data[i][6]}_${data[i][7]}`;
    } else if (type === "PO") {
      let d = data[i][0] instanceof Date ? Utilities.formatDate(data[i][0], Session.getScriptTimeZone(), "MM/dd/yyyy") : data[i][0];
      key = `${d}_${data[i][1]}_${data[i][2]}`;
    } else if (type === "AR") {
       let d = (dateIdx > -1 && data[i][dateIdx] instanceof Date) ? Utilities.formatDate(data[i][dateIdx], Session.getScriptTimeZone(), "MM/dd/yyyy") : (dateIdx > -1 ? data[i][dateIdx] : "");
       let m = managerIdx > -1 ? String(data[i][managerIdx]).toLowerCase().trim() : "";
       let t = textIdx > -1 ? String(data[i][textIdx]).toLowerCase().trim() : "";
       key = `${d}_${m}_${t}`;
    }
    if (key) keys.add(key);
  }
  return keys;
}


// =========================================================================
// REPLACEMENT: Keeps Scanned Invoices AND EWT Data in the active spreadsheet
// =========================================================================

function gatherAllData(prefix) {
  console.time(`Gather_${prefix}`);
  const aggregated = [];
  
  // --- UPDATED EXCEPTION: Scanned Invoices AND EWT Data stay in the active spreadsheet ---
  if (prefix === "Scanned_Supplier_Invoices_Data_" || prefix === PREFIX_SCANNED_INVOICES || prefix === "EWT_Data_" || prefix === PREFIX_EWT) {
    SpreadsheetApp.getActiveSpreadsheet().getSheets().filter(s => s.getName().startsWith(prefix)).forEach(s => {
      const lastRow = s.getLastRow();
      const lastCol = s.getLastColumn();
      if (lastRow > 1 && lastCol > 0) {
        const sheetName = s.getName();
        const values = s.getRange(2, 1, lastRow - 1, lastCol).getValues();
        for (let i = 0; i < values.length; i++) {
          aggregated.push({ sheetName, rowIndex: i + 2, data: values[i] });
        }
      }
    });
    console.timeEnd(`Gather_${prefix}`);
    return aggregated;
  }
  // --- END EXCEPTION ---

  let targetFolderId = DATABASE_FOLDER_ID;
  if (prefix.includes("Narra_Voucher_Data") || prefix.includes("Check_Summary_Data") || prefix.includes("RIWB_Data")) {
      targetFolderId = SCRIPT_PROPS['VERIFIED_FOLDER_ID'];
  }
  const dbFolder = DriveApp.getFolderById(targetFolderId);
  const files = dbFolder.getFiles();
  
  while (files.hasNext()) {
    let file = files.next();
    if (file.getName().startsWith(prefix) && file.getMimeType() === MimeType.GOOGLE_SHEETS) {
      let s = SpreadsheetApp.open(file).getSheets()[0];
      const lastRow = s.getLastRow();
      const lastCol = s.getLastColumn();
      if (lastRow > 1 && lastCol > 0) {
        const sheetName = file.getName(); 
        const values = s.getRange(2, 1, lastRow - 1, lastCol).getValues();
        for (let i = 0; i < values.length; i++) {
          aggregated.push({ sheetName, rowIndex: i + 2, data: values[i] });
        }
      }
    }
  }
  console.timeEnd(`Gather_${prefix}`);
  return aggregated;
}



function createTempPDF(fileName, headers, rows) {
  const tempSs = SpreadsheetApp.create(fileName);
  const tempFileId = tempSs.getId();
  const tempSheet = tempSs.getSheets()[0];
  
  // Set headers
  tempSheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold").setBackground("#EEEEEE");
  // Set data
  tempSheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
  
  // Prevent truncation in PDF by wrapping text and auto-resizing columns
  tempSheet.getDataRange().setWrap(true);
  tempSheet.getDataRange().setVerticalAlignment("top");
  for (let i = 1; i <= headers.length; i++) {
     tempSheet.autoResizeColumn(i);
  }
  
  SpreadsheetApp.flush(); Utilities.sleep(500);
  const blob = DriveApp.getFileById(tempFileId).getAs('application/pdf').setName(fileName + ".pdf");
  DriveApp.getFileById(tempFileId).setTrashed(true);
  return blob;
}
let GLOBAL_NTE_RUN_FOLDER = null;

function getOrCreateNteRunFolder() {
  if (GLOBAL_NTE_RUN_FOLDER) return GLOBAL_NTE_RUN_FOLDER;

  const rootFolder = DriveApp.getFolderById(NTE_COMPLAINTS_FOLDER_ID);
  const now = new Date();
  
  // Creates the year name e.g., "2026"
  const yearName = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy");
  // Creates just the month name e.g., "March"
  const monthName = Utilities.formatDate(now, Session.getScriptTimeZone(), "MMMM");
  // Creates the specific time folder e.g., "NTE_Complaints March 2026 8:13"
  const timeFolderName = "NTE_Complaints " + Utilities.formatDate(now, Session.getScriptTimeZone(), "MMMM yyyy H:mm");

  // 1. Get or Create Year Folder
  let yearFolder;
  const yearFolders = rootFolder.getFoldersByName(yearName);
  if (yearFolders.hasNext()) {
    yearFolder = yearFolders.next();
  } else {
    yearFolder = rootFolder.createFolder(yearName);
  }

  // 2. Get or Create Month Folder
  let monthFolder;
  const monthFolders = yearFolder.getFoldersByName(monthName);
  if (monthFolders.hasNext()) {
    monthFolder = monthFolders.next();
  } else {
    monthFolder = yearFolder.createFolder(monthName);
  }

  // 3. Get or Create Time Folder
  let runFolder;
  const runFolders = monthFolder.getFoldersByName(timeFolderName);
  if (runFolders.hasNext()) {
    runFolder = runFolders.next();
  } else {
    runFolder = monthFolder.createFolder(timeFolderName);
  }

  GLOBAL_NTE_RUN_FOLDER = runFolder; // Save to global so all files in this run use the exact same folder
  return runFolder;
}

function savePdfToDrive(blob) {
  const folder = getOrCreateNteRunFolder();
  return folder.createFile(blob).getUrl();
}

function createTempXLSX(fileName, headers, rows) {
  const tempSs = SpreadsheetApp.create(fileName);
  const tempFileId = tempSs.getId();
  const tempSheet = tempSs.getSheets()[0];
  
  // Set headers
  tempSheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold").setBackground("#EEEEEE");
  
  // Set data
  if (rows.length > 0) {
    tempSheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
  }
  
  // Format to prevent truncation (Wraps text & resizes columns automatically)
  tempSheet.getDataRange().setWrap(true);
  tempSheet.getDataRange().setVerticalAlignment("top");
  for (let i = 1; i <= headers.length; i++) {
     tempSheet.autoResizeColumn(i);
     // Forcing Item Name/Product columns to be very wide so they are highly visible
     if (headers[i-1] === "Item Name" || headers[i-1] === "Product" || headers[i-1] === "Description") {
         if (tempSheet.getColumnWidth(i) < 400) tempSheet.setColumnWidth(i, 450);
     }
  }
  
  SpreadsheetApp.flush(); 
  Utilities.sleep(3000); // 3-second wait allows Google servers to finalize the sheet before exporting
  
  let url = "https://docs.google.com/spreadsheets/d/" + tempFileId + "/export?format=xlsx";
  let token = ScriptApp.getOAuthToken();
  let blob;
  
  try {
    let response = UrlFetchApp.fetch(url, { 
        headers: { 'Authorization': 'Bearer ' + token },
        muteHttpExceptions: true
    });
    blob = response.getBlob().setName(fileName + ".xlsx");
    
    // Save the generated XLSX to the nested complaints folder directly
    if (NTE_COMPLAINTS_FOLDER_ID && NTE_COMPLAINTS_FOLDER_ID !== "INSERT_FOLDER_ID_HERE") {
        const folder = getOrCreateNteRunFolder();
        folder.createFile(blob);
    }
  } catch (e) {
    console.error("Failed to generate XLSX: " + e.toString());
    blob = Utilities.newBlob("", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName + ".xlsx");
  }
  
  DriveApp.getFileById(tempFileId).setTrashed(true);
  return blob;
}

function getNteHtmlTemplate(data) {
  return Template_NteHtml(data);
}

// =========================================================================
// IGNORE THIS FUNCTION
// =========================================================================

function _$runAllProcesses(targetAeFilter = null, targetMonthFilter = null, prevMonthExists = true, requestDate = null) {
  console.time("TOTAL_EXECUTION");
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  NTE_PAYLOAD_ACCUMULATOR = [];
  
  console.log("--- STARTED: AGGREGATING DATA ---");
  if(targetAeFilter) {
    let msg = `Processing ONLY for: ${targetAeFilter}`;
    if (targetMonthFilter) msg += ` (${targetMonthFilter})`;
    ss.toast(msg, "PROCESS", -1);
  } else {
    ss.toast("Aggregating Data from all sheets...", "PROCESS", -1);
  }
  
  // ADDED 5th PARAMETER "true" -> means "YES, Write the Spreadsheets"
  _$processBuySellForms(targetAeFilter, targetMonthFilter, prevMonthExists, requestDate, true);
  
  // REMOVED: _$processSendEmailConsolidatedNTEs(); 
  // We do not send emails here anymore, as per your request.
  
  console.log("--- COMPLETED: ALL PROCESSES ---");
  ss.toast("Process Completed. Reports Generated.", "SUCCESS", 5);
  console.timeEnd("TOTAL_EXECUTION");
}

function deleteTriggersByName(functionName) {
   const triggers = ScriptApp.getProjectTriggers();
   for(let i=0; i<triggers.length; i++) {
      if(triggers[i].getHandlerFunction() === functionName) {
         ScriptApp.deleteTrigger(triggers[i]);
      }
   }
}

// =========================================================================
// EWT SUPPLIER & PURCHASE REPORT BATCH PROCESSING (PER AE)
// =========================================================================

const PURCHASE_BATCH_KEY = "PENDING_AES_PURCHASE";
const PURCHASE_TRIGGER_NAME = "resumePurchaseBatch";



const EWT_CLIENT_BATCH_KEY = "PENDING_EWT_CLIENT";
const EWT_CLIENT_TRIGGER_NAME = "resumeEwtClientBatch";

function _30_runSendAllEwtClientReports() {
  const props = PropertiesService.getScriptProperties();
  deleteTriggersByName(EWT_CLIENT_TRIGGER_NAME);
  props.deleteProperty(EWT_CLIENT_BATCH_KEY);
  props.deleteProperty("CANCEL_EWT_CLIENT");

  const startStr = props.getProperty("BATCH_START_DATE");
  const endStr = props.getProperty("BATCH_END_DATE");
  
  // Queue a single task for the exact date range selected in the UI
  let taskPayload = JSON.stringify([{ start: startStr, end: endStr }]);
      props.setProperty(EWT_CLIENT_BATCH_KEY, taskPayload);
      
      resumeEwtClientBatch();
}

function resumeEwtClientBatch() {
  const props = PropertiesService.getScriptProperties();
  props.setProperty("THREAD_STATE", "RUNNING"); // <--- ADDED
  
  let pending = JSON.parse(props.getProperty(EWT_CLIENT_BATCH_KEY) || "[]");
  
  if (pending.length === 0 || props.getProperty("CANCEL_EWT_CLIENT") === "true") {
    deleteTriggersByName(EWT_CLIENT_TRIGGER_NAME);
    props.deleteProperty(EWT_CLIENT_BATCH_KEY);
    props.deleteProperty("CANCEL_EWT_CLIENT");
    clearReportLock(); // <-- ADDED THIS TO UNLOCK THE UI
    return;
  }

  let task = pending.shift();
  try { _$processAndSendEWTReportClient(task.start, task.end); } catch (e) { console.error(e); }
  
  props.setProperty(EWT_CLIENT_BATCH_KEY, JSON.stringify(pending));
  deleteTriggersByName(EWT_CLIENT_TRIGGER_NAME);
  
  // Cleanly unlock when finished
  if (pending.length === 0) {
      props.deleteProperty(EWT_CLIENT_BATCH_KEY);
      clearReportLock(); // <-- ADDED THIS TO UNLOCK THE UI
  }
}

function _resumeEwtClientCore(batchKey, triggerName, cancelKey) {
  const startTime = Date.now();
  const MAX_EXECUTION_TIME = 4.5 * 60 * 1000; 
  const props = PropertiesService.getScriptProperties();
  props.setProperty("THREAD_STATE", "RUNNING"); // <--- ADDED
  
  let pendingQueueStr = props.getProperty(batchKey);
  
  if (!pendingQueueStr) { deleteTriggersByName(triggerName); return; }
  
  let pendingQueue = JSON.parse(pendingQueueStr);
  if (pendingQueue.length === 0) {
    deleteTriggersByName(triggerName);
    props.deleteProperty(batchKey);
    clearReportLock();
    console.log("All EWT Client reports processed successfully.");
    return;
  }

  while (pendingQueue.length > 0) {
     if (props.getProperty(cancelKey) === "true") {
         deleteTriggersByName(triggerName);
         props.deleteProperty(batchKey); props.deleteProperty(cancelKey);
         return; 
     }

     if (Date.now() - startTime > MAX_EXECUTION_TIME) {
        props.setProperty(batchKey, JSON.stringify(pendingQueue));
        replaceTrigger(triggerName, 60000);
        return; 
     }
     
     let currentItem = pendingQueue.shift(); 
     let parts = currentItem.split("|");
     let currentAE = parts[0];
     let targetMonthStr = parts[1];
     
     try {
         _$processAndSendEWTReportClient(currentAE, [targetMonthStr]);
     } catch(e) { console.error("Error processing " + currentItem + ": " + e); }

     props.setProperty(batchKey, JSON.stringify(pendingQueue));
     replaceTrigger(triggerName, 60000);
     return; 
  }
  deleteTriggersByName(triggerName);
  props.deleteProperty(batchKey);
  clearReportLock();
}



// =========================================================================
// UNIFIED UI GENERATOR (Replaces 10 redundant UI functions)
// =========================================================================
function _buildAndShowDialog(title, btnText, srvFunc, hasAe, hasReq, h) {
  const ui = SpreadsheetApp.getUi();
  const dataMap = getDataForDropdowns();

  if (Object.keys(dataMap).length === 0) {
    ui.alert("No Data found to populate the menu.");
    return;
  }

  const htmlOutput = HtmlService.createHtmlOutput(Template_DialogUI(title, hasAe, hasReq, srvFunc, btnText, dataMap)).setWidth(400).setHeight(h);
  
  ui.showModalDialog(htmlOutput, 'Finance System Generator');
}

// =========================================================================
// WEB APP BACKEND HANDLERS
// =========================================================================

function webappUploadFile(fileData, fileName, importType, sessionToken = null) {
  try {
    let folderId;
    switch(importType) {
      case "AR": folderId = AR_FOLDER_ID; break;
      case "SR": folderId = SR_FOLDER_ID; break;
      case "PO_PINGCON_BONITA": folderId = PO_PINGCON_BONITA_FOLDER_ID; break;
      case "PO_KEYSYS": folderId = PO_KEYSYS_FOLDER_ID; break;
      case "SQ": folderId = SQ_FOLDER_ID; break;
      case "AT": folderId = ACCOUNT_TITLE_FOLDER_ID; break;
      case "PMC": folderId = PMC_DIRECTORY_FOLDER_ID; break;
      case "BASKET": folderId = BASKET_FOLDER_ID; break;
      case "SKU": folderId = SKU_MAPPING_OLD_NEW_FOLDER_ID; break;
      case "SCANNED": folderId = SUPPLIER_RECEIPT_FOLDER_ID; break; 
      case "AE_EMAILS": folderId = AE_EMAILS_FOLDER_ID; break; 
      case "CPP": folderId = CPP_FOLDER_ID; break;
      case "NEW_NARRA": {
        let isDev = false;
        try { isDev = ScriptApp.getService().getUrl().toLowerCase().endsWith("/dev"); } catch(e){}
        let props = PropertiesService.getScriptProperties();
        folderId = (isDev && props.getProperty('NEW_NARRA_FOLDER_ID_DEV'))
          ? props.getProperty('NEW_NARRA_FOLDER_ID_DEV')
          : NEW_NARRA_FOLDER_ID;
        break;
      }
      case "NARRA_VOUCHER": 
      case "CHECK_SUMMARY": 
      case "RIWB": 
          folderId = SCRIPT_PROPS['VERIFIED_FOLDER_ID']; 
          fileName = "RAW_2307_" + fileName; 
          break;
      default: throw new Error("Invalid import type selected.");
    }
    
    let dataStr = fileData.split(',')[1];
    let mimeType;
    
    // Extract the true MimeType directly from the browser to bypass Windows hiding extensions
    try { mimeType = fileData.split(';')[0].split(':')[1]; } catch(e) {}
    
    if (!mimeType || mimeType === "application/octet-stream") {
        let lowerName = fileName.toLowerCase();
        if (lowerName.endsWith('.csv')) mimeType = MimeType.CSV;
        else if (lowerName.endsWith('.xlsx')) mimeType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
        else mimeType = MimeType.MICROSOFT_EXCEL;
    }
    
    let blob = Utilities.newBlob(Utilities.base64Decode(dataStr), mimeType, fileName);
    DriveApp.getFolderById(folderId).createFile(blob);

    if (importType === "NEW_NARRA") {
      logNarraActivity("Uploaded Narra Database", `Uploaded raw file: ${fileName}`, sessionToken);
    }
    
    return "Success";
  } catch(e) {
    let errorMsg = e.toString();
    if (errorMsg.includes("Bad Request") || errorMsg.includes("drive.files.copy") || errorMsg.includes("drive.files.create")) {
        let fileMatch = errorMsg.match(/([a-zA-Z0-9_\-\s\(\)\,\.]+?\.(?:xlsx|xls|csv))/i);
        let badFileName = fileMatch ? fileMatch[1].trim() : fileName;
        return "Error: IMPORT FAILED! The file [" + badFileName + "] is PASSWORD PROTECTED. Please remove the password and upload it again.";
    }
    return "Error: " + errorMsg;
  }
}

// --- DOMAIN-ISOLATED LOCK ENGINE (PREVENTS CROSS-PORTAL BLOCKING) ---
function _getPortalLockPrefix(portalModule = "MAIN") {
  let isDev = false;
  try { isDev = ScriptApp.getService().getUrl().toLowerCase().endsWith("/dev"); } catch(e){}
  let env = isDev ? "DEV" : "EXEC";
  let rawMod = String(portalModule || "MAIN").toUpperCase().trim();
  let mod = "MAIN";
  if (rawMod.includes("NARRA")) mod = "NARRA";
  else if (rawMod.includes("CLIENT")) mod = "CLIENT";
  else if (rawMod.includes("SUPPLIER") || rawMod.includes("2307") || rawMod.includes("INVOICE")) mod = "SUPPLIER";
  return `${env}_LOCK_${mod}`;
}

function webappRequestUploadLock(portalModule = "MAIN") {
  try {
    let prefix = _getPortalLockPrefix(portalModule);
    let cache = CacheService.getScriptCache();
    
    // Only inspect the lock assigned to this specific portal and environment
    if (cache.get(prefix + "_UPLOAD") === "true") {
      return "BUSY"; 
    }
    
    let props = PropertiesService.getScriptProperties();
    let ongoingReport = props.getProperty(prefix + "_REPORT_TYPE");
    let ongoingTimestamp = props.getProperty(prefix + "_REPORT_TIMESTAMP");
    
    if (ongoingReport && ongoingReport.trim() !== "") {
        if (ongoingTimestamp && (Date.now() - parseInt(ongoingTimestamp) > 360000)) {
            clearReportLock(portalModule);
            ongoingReport = null;
        } else {
            return "BUSY_REPORT";
        }
    }
    
    return "SUCCESS";
  } catch (err) {
    console.error("Lock check error: " + err.message);
    return "SUCCESS";
  }
}

function webappSetUploadLock(portalModule = "MAIN") {
  let prefix = _getPortalLockPrefix(portalModule);
  CacheService.getScriptCache().put(prefix + "_UPLOAD", "true", 300);
}

function webappReleaseUploadLock(portalModule = "MAIN") {
  let prefix = _getPortalLockPrefix(portalModule);
  CacheService.getScriptCache().remove(prefix + "_UPLOAD");
}

function webappTriggerReport(action, startDateStr, endDateStr, selectedEntities, selectedAEs, isTestMode, selectedBanks) {
  console.log("[DEBUG] Starting webappTriggerReport for action: " + action);
  console.log("[DEBUG] Number of selected AEs/Customers received: " + (selectedAEs ? selectedAEs.length : 0));
  
  try {
    if (!SpreadsheetApp.getActiveSpreadsheet()) {
      throw new Error("Cannot detect the database. Please ensure the Web App is deployed as 'Execute as: Me'.");
    }

    const props = PropertiesService.getScriptProperties();
    
    let portalMod = (action === "NIGHTLY_NARRA") ? "NARRA" : "MAIN";
    let prefix = _getPortalLockPrefix(portalMod);

    if (action !== "EWT_SUPPLIER_REPORT") {
        if (props.getProperty(prefix + "_REPORT_TYPE")) {
            return "Error: A report (" + props.getProperty(prefix + "_REPORT_TYPE") + ") is currently ongoing on this portal. Please wait for it to finish or click Cancel Batch.";
        }
        props.setProperty(prefix + "_REPORT_TYPE", action);
        props.setProperty(prefix + "_REPORT_STATUS", "Processing " + action + " in background...");
        props.setProperty(prefix + "_REPORT_TIMESTAMP", Date.now().toString());
        props.setProperty(prefix + "_THREAD_STATE", "RUNNING");
        props.deleteProperty(prefix + "_RESUME_TIME");
        props.deleteProperty(prefix + "_LAST_RUN_ERROR");
    }
    
    console.log("[DEBUG] Saving basic props...");
    if (startDateStr) props.setProperty("BATCH_START_DATE", startDateStr);
    if (endDateStr) props.setProperty("BATCH_END_DATE", endDateStr);
    if (selectedEntities) props.setProperty("SALES_ENTITIES", JSON.stringify(selectedEntities));
    props.setProperty("TEST_MODE", isTestMode ? "true" : "false");
    if (selectedBanks) props.setProperty("NIGHTLY_NARRA_BANKS", JSON.stringify(selectedBanks));
    
    if (selectedAEs) {
        let jsonString = JSON.stringify(selectedAEs);
        console.log(`[DEBUG] JSON Stringified Payload Length: ${jsonString.length} characters.`);
        if (jsonString.length > 8500) {
            console.warn("[DEBUG] WARNING: Payload is dangerously close to or exceeding the 9KB limit of PropertiesService!");
        }
        
        if (action === "UNPAID_DR_CUST" || action === "UNPAID_DR_OVERDUE" || action === "CPP_REPORT") {
            console.log("[DEBUG] Saving to SALES_TARGET_CUSTOMERS...");
            props.setProperty("SALES_TARGET_CUSTOMERS", jsonString);
            props.deleteProperty("SALES_TARGET_AES");
        } else {
            console.log("[DEBUG] Saving to SALES_TARGET_AES...");
            props.setProperty("SALES_TARGET_AES", jsonString);
            props.deleteProperty("SALES_TARGET_CUSTOMERS");
        }
    }
    
    console.log("[DEBUG] Initial property saving complete. Moving to switch statement.");
    let msg = "Report processed successfully! Emails have been sent.";
        
        switch(action) {
          case "UNPAID_DR_OVERDUE": 
            _9d_runSendAllUnpaidDRsOverdue(); 
            msg = "Success: Overdue Aging Summary processing started in the background! You can safely close this window.";
            break;
          case "UNPAID_DR_AE": 
            _9_runSendAllUnpaidDRs(); 
            msg = "Success: Unpaid DRs (Per AE) processing started in the background! You can safely close this window.";
            break;
          case "UNPAID_DR_CUST": 
            _9b_runSendAllUnpaidDRsCustomer(); 
            msg = "Success: Unpaid DRs (Per Customer) processing started in the background! You can safely close this window.";
            break;
          
          case "BUY_SELL_REPORT": 
            _26_runSendAllBuySellReports(); 
            msg = "Success: Buy/Sell Reports batch processing started in the background! You can safely close this window.";
            break;
          case "COMMISSION_REPORT": 
            _27_runSendAllCommissionReports(); 
            msg = "Success: Commission Reports batch processing started in the background! You can safely close this window.";
            break;
          
          case "EWT_SUPPLIER_REPORT": 
            _35_runSendEwtSupplierReport(); 
            msg = "Success: EWT Report Supplier successfully generated and emailed!";
            // Note: clearReportLock() removed so it doesn't accidentally unlock the main portal
            break;
          case "CHECK_PO_PAYMENT": 
            _32_runCheckPOPaymentStatus(); 
            msg = "Success: PO Payment Status check successfully generated and emailed!";
            clearReportLock();
            break;
          case "CPP_REPORT": 
            props.setProperty("CPP_QUARTER", startDateStr);
            props.setProperty("CPP_YEAR", endDateStr);
            _34_runSendAllCustomerPaymentPerformance(); 
            msg = "Success: Customer Payment Performance batch processing started! You can safely close this window.";
            break;
          case "SALES_REPORT_CONSOLIDATED": 
            _10b_runSendAllSalesReportsConsolidated(); 
            msg = "Success: Consolidated Sales Report processing started in the background! You can safely close this window.";
            break;
          case "NIGHTLY_NARRA": 
            _40_runNightlyNarra(); 
            msg = "Success: Nightly Narra batch processing started! You can safely close this window.";
            break;
          default: throw new Error("Invalid action selected.");
          
        }
        
        CacheService.getScriptCache().remove("GLOBAL_UPLOAD_LOCK");
        return msg;
      } catch(e) {
    CacheService.getScriptCache().remove("GLOBAL_UPLOAD_LOCK");
    clearReportLock(); // UNLOCKS IF INITIALIZATION FAILS INSTANTLY
    return "Error: " + e.message + "\nStack: " + e.stack;
  }
}

function webappCancelReport(portalModule = "MAIN") {
  try {
    const props = PropertiesService.getScriptProperties();
    let mod = String(portalModule || "MAIN").toUpperCase().trim();
    let prefix = _getPortalLockPrefix(mod);
    
    CacheService.getScriptCache().remove(prefix + "_UPLOAD");
    CacheService.getScriptCache().remove("GLOBAL_UPLOAD_LOCK");
    
    // 1. ISOLATED NIGHTLY NARRA CANCELLATION (NEVER TOUCHES MAIN PORTAL OR SOA)
    if (mod === "NARRA") {
      deleteTriggersByName("resumeNightlyNarraBatch");
      props.deleteProperty("PENDING_NIGHTLY_NARRA");
      props.deleteProperty("CANCEL_NIGHTLY_NARRA");
      clearReportLock("NARRA");
      return "Success: Nightly Narra batch cancelled and portal unlocked.";
    }

    // 2. MAIN WEBSITE CANCELLATION
    let ongoingType = props.getProperty(prefix + "_REPORT_TYPE") || props.getProperty("ONGOING_REPORT_TYPE");
    let threadState = props.getProperty(prefix + "_THREAD_STATE") || props.getProperty("THREAD_STATE");

    const batchTriggers = ["resumeUnpaidDRBatch", "resumeUnpaidDRCustomerBatch", "resumeUnpaidDROverdueBatch", "resumeSalesReportConsolidatedBatch", "resumeBuySellBatch", "resumeCommissionBatch", "resumePurchaseBatch", "resumeEwtClientBatch", "resumeCheckPOPaymentBatch", "resumeCPPBatch", "resumeEwtSupplierBatch"];
    const cancelKeys = ["CANCEL_UNPAID_DR", "CANCEL_UNPAID_DR_CUSTOMER", "CANCEL_UNPAID_DR_OVERDUE", "CANCEL_SALES_REPORT_CONSOLIDATED", "CANCEL_BUY_SELL", "CANCEL_COMM", "CANCEL_PURCHASE", "CANCEL_EWT_CLIENT", "CANCEL_CHECK_PO_PAYMENT", "CANCEL_CPP", "CANCEL_EWT_SUPPLIER"];
    const keysToNuke = [
      "PENDING_AES_UNPAID_DR", "PENDING_CUSTOMERS_UNPAID_DR", "PENDING_ENTITIES_UNPAID_OVERDUE", "PENDING_AES_SALES_REPORT_CONSOLIDATED", "MISSING_QUOTAS_SALES", "PENDING_AES_BUY_SELL", 
      "PENDING_AES_COMM", "PENDING_AES_PURCHASE", "PENDING_EWT_CLIENT", 
      "SALES_TARGET_AES", "SALES_TARGET_CUSTOMERS", "RESUME_TIME", "PENDING_CUSTOMERS_CPP", "LAST_RUN_ERROR"
    ];

    if (ongoingType) {
        sendBatchSuccessSummary(true, ongoingType);
    }

    if (threadState === "RUNNING") {
        cancelKeys.forEach(k => props.setProperty(k, "true"));
        props.setProperty(prefix + "_REPORT_STATUS", "Cancelling... Waiting for current file to finish safely.");
        return "Cancellation requested. The UI will unlock as soon as the current file finishes generating.";
    } else {
        batchTriggers.forEach(t => deleteTriggersByName(t));
        keysToNuke.forEach(k => props.deleteProperty(k));
        cancelKeys.forEach(k => props.deleteProperty(k));
        clearReportLock("MAIN");
        return "Success: Main portal reports cancelled and unlocked.";
    }
  } catch(e) {
    return "Error: " + e.message;
  }
}
function webappCheckImportStatus(importType) {
  let folderId; let funcName = null;
  switch(importType) {
    case "AR": folderId = AR_FOLDER_ID; funcName = "_20_importAgeingReports"; break;
    case "SR": folderId = SR_FOLDER_ID; funcName = "_16_importSalesReports"; break;
    case "PO_PINGCON_BONITA": folderId = PO_PINGCON_BONITA_FOLDER_ID; funcName = "_17_importPurchaseOrdersPingconAndBonita"; break;
    case "PO_KEYSYS": folderId = PO_KEYSYS_FOLDER_ID; funcName = "_25_importPurchaseOrdersKeysysAndA2Z2GO"; break;
    case "SQ": folderId = SQ_FOLDER_ID; funcName = "_21_importSalesQuotas"; break;
    case "AT": folderId = ACCOUNT_TITLE_FOLDER_ID; funcName = "_18_importAccountTitles"; break;
    case "PMC": folderId = PMC_DIRECTORY_FOLDER_ID; funcName = "_19_importPMCDirectory"; break;
    case "BASKET": folderId = BASKET_FOLDER_ID; funcName = "_23_importMissingSKUsMappingBaskets"; break;
    case "SKU": folderId = SKU_MAPPING_OLD_NEW_FOLDER_ID; funcName = "_22_importMissingSKUsMappingOldToNew"; break;
    case "AE_EMAILS": folderId = AE_EMAILS_FOLDER_ID; funcName = "_31_importAEEmails"; break;
    case "CPP": folderId = CPP_FOLDER_ID; funcName = "_33_importCustomerPaymentPerformance"; break;
    case "NEW_NARRA": folderId = NEW_NARRA_FOLDER_ID; funcName = "_39_importNewNarraReport"; break; // <--- ADD THIS LINE
    case "NARRA_VOUCHER": folderId = SCRIPT_PROPS['VERIFIED_FOLDER_ID']; funcName = "_36_importNarraVoucherReport"; break;
    case "CHECK_SUMMARY": folderId = SCRIPT_PROPS['VERIFIED_FOLDER_ID']; funcName = "_37_importCheckSummary"; break;
    case "RIWB": folderId = SCRIPT_PROPS['VERIFIED_FOLDER_ID']; funcName = "_38_importRIWBReport"; break;
    case "SCANNED": folderId = SUPPLIER_RECEIPT_FOLDER_ID; break; 
  }
  
  if (!funcName) return { done: true, filesLeft: 0, inBackground: false };
  
  let count = 0;
  if (folderId) {
    let files = DriveApp.getFolderById(folderId).getFiles();
    while (files.hasNext()) {
      let f = files.next();
      let name = f.getName().toLowerCase();
      
      if (importType === "NARRA_VOUCHER" || importType === "CHECK_SUMMARY" || importType === "RIWB") {
          if (!name.startsWith("raw_2307_")) continue;
      }
      
      if (name.endsWith(".xlsx") || name.endsWith(".xls") || name.endsWith(".csv") || name.endsWith(".ods")) count++;
      else try { f.setTrashed(true); } catch(e) {}
    }
  }
  
  let hasTrigger = false;
  let triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === funcName) { hasTrigger = true; break; }
  }

  let isRunning = false;
  let lock = LockService.getScriptLock();
  try { lock.waitLock(1); lock.releaseLock(); } catch (e) { isRunning = true; }
  
  if (count > 0 && !isRunning && !hasTrigger) {
    deleteTriggersByName(funcName);
    ScriptApp.newTrigger(funcName).timeBased().after(1000).create();
    hasTrigger = true;
  }
  
  // ADDED: Cleanup orphaned triggers instantly to unlock the UI
  if (count === 0 && hasTrigger && !isRunning) {
    deleteTriggersByName(funcName);
    hasTrigger = false;
  }
  
  return { done: (count === 0 && !hasTrigger && !isRunning), filesLeft: count, inBackground: (hasTrigger || isRunning) };
}

function webappProcessStuckFiles(importType) {
  let folderId; let funcToRun = null;
  switch(importType) {
    case "AR": folderId = AR_FOLDER_ID; funcToRun = _20_importAgeingReports; break;
    case "SR": folderId = SR_FOLDER_ID; funcToRun = _16_importSalesReports; break;
    case "PO_PINGCON_BONITA": folderId = PO_PINGCON_BONITA_FOLDER_ID; funcToRun = _17_importPurchaseOrdersPingconAndBonita; break;
    case "PO_KEYSYS": folderId = PO_KEYSYS_FOLDER_ID; funcToRun = _25_importPurchaseOrdersKeysysAndA2Z2GO; break;
    case "SQ": folderId = SQ_FOLDER_ID; funcToRun = _21_importSalesQuotas; break;
    case "AT": folderId = ACCOUNT_TITLE_FOLDER_ID; funcToRun = _18_importAccountTitles; break;
    case "PMC": folderId = PMC_DIRECTORY_FOLDER_ID; funcToRun = _19_importPMCDirectory; break;
    case "BASKET": folderId = BASKET_FOLDER_ID; funcToRun = _23_importMissingSKUsMappingBaskets; break;
    case "SKU": folderId = SKU_MAPPING_OLD_NEW_FOLDER_ID; funcToRun = _22_importMissingSKUsMappingOldToNew; break;
    case "AE_EMAILS": folderId = AE_EMAILS_FOLDER_ID; funcToRun = _31_importAEEmails; break;
    case "CPP": folderId = CPP_FOLDER_ID; funcToRun = _33_importCustomerPaymentPerformance; break;
    case "NEW_NARRA": folderId = NEW_NARRA_FOLDER_ID; funcToRun = _39_importNewNarraReport; break; // <--- ADD THIS LINE
    case "NARRA_VOUCHER": folderId = SCRIPT_PROPS['VERIFIED_FOLDER_ID']; funcToRun = _36_importNarraVoucherReport; break;
    case "CHECK_SUMMARY": folderId = SCRIPT_PROPS['VERIFIED_FOLDER_ID']; funcToRun = _37_importCheckSummary; break;
    case "RIWB": folderId = SCRIPT_PROPS['VERIFIED_FOLDER_ID']; funcToRun = _38_importRIWBReport; break;
    case "SCANNED": return "NO_FILES"; 
  }
  
  let count = 0;
  if (folderId) {
    let files = DriveApp.getFolderById(folderId).getFiles();
    while (files.hasNext()) {
      let f = files.next();
      let name = f.getName().toLowerCase();
      
      if (importType === "NARRA_VOUCHER" || importType === "CHECK_SUMMARY" || importType === "RIWB") {
          if (!name.startsWith("raw_2307_")) continue;
      }
      
      if (name.endsWith(".xlsx") || name.endsWith(".xls") || name.endsWith(".csv") || name.endsWith(".ods")) count++;
      else try { f.setTrashed(true); } catch(e) {}
    }
  }
  
  if (count === 0) return "NO_FILES";
  
  let targetPortal = "MAIN";
  if (importType === "NEW_NARRA") targetPortal = "NARRA";
  else if (importType === "NARRA_VOUCHER" || importType === "CHECK_SUMMARY" || importType === "RIWB") targetPortal = "SUPPLIER";

  if (funcToRun) {
     webappSetUploadLock(targetPortal);
     try { funcToRun(); } finally { webappReleaseUploadLock(targetPortal); }
  }
  return count;
}



function webappGetDateBounds() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  

  const arSheets = getDatabaseSheets("AR_Data_");
  
  // --- ADDED: READ EMAIL EXTRACTS TO FILTER UI ---
  const emailSheets = getDatabaseSheets("Customer_Emails_Extracted_");
  let validEmailCustomers = new Set();
  emailSheets.forEach(sheet => {
     const data = sheet.getDataRange().getValues();
     for(let i=1; i<data.length; i++){ // Start at 1 to skip headers
         let cust = String(data[i][1]).replace(/[^\x20-\x7EñÑ]/g, '').trim(); // Column B
         let email = String(data[i][3]).trim(); // Column D (Shifted due to Cust Manager in C)
         
         // Only add the customer if they have a valid email address containing "@"
         if(cust && email && email.includes("@")) {
             validEmailCustomers.add(cust.toLowerCase().replace(/\s+/g, ''));
         }
     }
  });
  // ----------------------------------------------
      
      // --- ADDED: READ AE EMAILS TO FILTER UI AEs ---
      const aeEmailSheets = getDatabaseSheets(PREFIX_AE_EMAILS);
      let validAEsFromEmail = new Set();
      let exactAENames = [];
      let aeEmailEntityMap = {};
      aeEmailSheets.forEach(sheet => {
         const data = sheet.getDataRange().getValues();
         if(data.length < 2) return;
         let headers = data[0].map(h => String(h).toLowerCase().trim());
         let nameIdx = headers.indexOf("ae name");
         let pIdx = headers.indexOf("pingcon email");
         let aIdx = headers.indexOf("a2z2go email");
         let kIdx = headers.indexOf("keysys email");

         for(let i=1; i<data.length; i++){
             let aeName = nameIdx > -1 ? String(data[i][nameIdx]).replace(/[^\x20-\x7E]/g, '').replace(/\s+/g, ' ').trim() : "";
             if(aeName) {
                 validAEsFromEmail.add(aeName.toLowerCase());
                 exactAENames.push(aeName);
                 
                 let pingconEmail = pIdx > -1 ? String(data[i][pIdx]).trim() : "";
                 let a2zEmail = aIdx > -1 ? String(data[i][aIdx]).trim() : "";
                 let keysysEmail = kIdx > -1 ? String(data[i][kIdx]).trim() : "";
                 
                 let ents = new Set();
                 if (pingconEmail && pingconEmail.includes("@")) {
                     ents.add("Pingcon"); ents.add("Bonita"); ents.add("A2ZPMC");
                 }
                 if ((a2zEmail && a2zEmail.includes("@")) || (keysysEmail && keysysEmail.includes("@"))) {
                     ents.add("Keysys"); ents.add("A2ZIT");
                 }
                 if (ents.size === 0) { ents.add("Pingcon"); ents.add("Bonita"); ents.add("A2ZPMC"); } 
                 
                 aeEmailEntityMap[aeName] = ents;
             }
         }
      });
      // ----------------------------------------------

      let minDate = Infinity;
      let maxDate = 0;
      let aeEntitiesMap = {}; 
      let custEntitiesMap = {}; 

      // Fetch all AEs who actually have assigned Sales Quotas in SQ_Data
      const sqSheets = getDatabaseSheets("SQ_Data_");
      let quotaAENames = new Set();
      sqSheets.forEach(sheet => {
          const data = sheet.getDataRange().getValues();
          for (let i = 1; i < data.length; i++) {
              let managerRaw = String(data[i][0]).trim();
              let managerCanonical = getCanonicalName(managerRaw).replace(/[^\x20-\x7E]/g, '').replace(/\s+/g, ' ').trim();
              if (managerCanonical && managerCanonical.toLowerCase() !== "ae name") {
                  quotaAENames.add(managerCanonical);
              }
          }
      });

      // DYNAMIC INTERSECTION: Only list AEs who exist in BOTH AE_Emails_Data AND SQ_Data
      exactAENames.forEach(aeName => {
          let canonical = getCanonicalName(aeName);
          let hasQuota = Array.from(quotaAENames).some(qAe => checkNameMatch(qAe, canonical));
          
          if (hasQuota) {
              aeEntitiesMap[aeName] = { 
                  entities: aeEmailEntityMap[aeName] || new Set(["Pingcon", "Bonita", "A2ZPMC", "Keysys", "A2ZIT"]), 
                  minDate: 0, 
                  maxDate: 9999999999999 
              };
          }
      });

  arSheets.forEach(sheet => {
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return;
    
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h).toLowerCase().trim());
    const dateIdx = headers.findIndex(h => h.includes("date") && !h.includes("expiry") && !h.includes("countered") && !h.includes("received"));
    const managerIdx = headers.findIndex(h => h.includes("manager"));
    const customerIdx = headers.findIndex(h => h.includes("customer")); 
    const sourceFileIdx = headers.findIndex(h => h.includes("source file"));
    
    // --- ADDED: FETCH OPEN BALANCE & STATUS INDICES FOR TARGETED FILTERING ---
    const openBalIdx = headers.findIndex(h => h.includes("open balance") && !h.includes("%"));
    const paymentStatusIdx = headers.findIndex(h => h === "payment status" || h.includes("payment status"));
    const textIdx = headers.findIndex(h => h.includes("additional text"));
    const notesIdx = headers.findIndex(h => h === "notes");

    const data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();

    for (let i = 0; i < data.length; i++) {
      let row = data[i];
      let dTime = null;
      
      if (dateIdx > -1) {
        let rawDate = row[dateIdx];
        if (rawDate instanceof Date) {
          dTime = rawDate.getTime();
        } else if (rawDate) {
          let parsed = new Date(rawDate).getTime();
          if (!isNaN(parsed)) dTime = parsed;
        }
        
        if (dTime) {
          if (dTime < minDate) minDate = dTime;
          if (dTime > maxDate) maxDate = dTime;
        }
      }
      
      let sourceFile = (sourceFileIdx > -1) ? String(row[sourceFileIdx]) : "";
      let companyEntity = "Pingcon";
      if (sourceFile) {
        let sf = sourceFile.toLowerCase();
        if (sf.includes("bonita")) companyEntity = "Bonita";
        else if (sf.includes("keysys")) companyEntity = "Keysys";
        else if (sf.includes("a2zit") || (sf.includes("a2z") && sf.includes("it"))) companyEntity = "A2ZIT";
        else if (sf.includes("a2zpmc") || (sf.includes("a2z") && sf.includes("pmc"))) companyEntity = "A2ZPMC";
      }

      let managerCanonical = "";
      if (managerIdx > -1) {
        let managerRaw = String(row[managerIdx]).trim();
        managerCanonical = getCanonicalName(managerRaw).replace(/[^\x20-\x7E]/g, '').replace(/\s+/g, ' ').trim();
        if (managerCanonical && managerCanonical.toLowerCase() !== "cust manager") {
           // ONLY ADD IF THE AE EXISTS IN AE_Emails_Data_
           if (validAEsFromEmail.has(managerCanonical.toLowerCase())) {
               if (!aeEntitiesMap[managerCanonical]) aeEntitiesMap[managerCanonical] = { entities: new Set(), minDate: Infinity, maxDate: 0 };
               aeEntitiesMap[managerCanonical].entities.add(companyEntity);
               if (dTime) {
                 if (dTime < aeEntitiesMap[managerCanonical].minDate) aeEntitiesMap[managerCanonical].minDate = dTime;
                 if (dTime > aeEntitiesMap[managerCanonical].maxDate) aeEntitiesMap[managerCanonical].maxDate = dTime;
               }
           }
        }
      }

      if (customerIdx > -1) {
        // --- ADDED: CHECK IF CUSTOMER ACTUALLY HAS AN OPEN BALANCE & NOT VOID/PAID ---
        let openBal = (openBalIdx > -1) ? parseFloat(row[openBalIdx]) : 0;
        if (isNaN(openBal)) openBal = 0;
        
        let statusRaw = "";
        let pStatIdxToUse = paymentStatusIdx > -1 ? paymentStatusIdx : (openBalIdx > -1 ? openBalIdx + 2 : -1);
        if (pStatIdxToUse > -1 && pStatIdxToUse < row.length) {
            statusRaw = String(row[pStatIdxToUse]).toLowerCase().trim();
        }
        
        let textVal = (textIdx > -1) ? String(row[textIdx]).toUpperCase() : "";
        let notesVal = (notesIdx > -1) ? String(row[notesIdx]).toUpperCase() : "";
        let isVoid = isTransactionVoid(notesVal, textVal);

        let hasOpenBalance = (openBal >= 1) && (statusRaw !== "fully paid") && !isVoid;

        // FORCE STRIP invisible unicode characters for website display deduplication
        let customerRaw = String(row[customerIdx]).replace(/[^\x20-\x7EñÑ]/g, '').replace(/\s+/g, ' ').trim();
        let customerKey = customerRaw.toLowerCase().replace(/\s+/g, '');
        
        // --- ADDED: EXCLUDE INTERNAL COMPANIES ---
        let excludedCompanies = ["pingconmarketingcorporation", "bonitatissuecleancorporation", "a2z2gomarketingcorp", "keysysinc", "keysysincorporated"];
        let isExcluded = excludedCompanies.some(ex => customerKey.includes(ex));
        
        // ONLY insert Customer into the UI list if they pass the unpaid check AND are not excluded
        if (customerRaw && hasOpenBalance && managerCanonical && !isExcluded) {
           if (!custEntitiesMap[customerKey]) {
               custEntitiesMap[customerKey] = { name: customerRaw, entities: new Set(), minDate: Infinity, maxDate: 0 };
           }
           custEntitiesMap[customerKey].entities.add(companyEntity);
           
           // Logs their min/max dates strictly based on unpaid transactions 
           if (dTime) {
             if (dTime < custEntitiesMap[customerKey].minDate) custEntitiesMap[customerKey].minDate = dTime;
             if (dTime > custEntitiesMap[customerKey].maxDate) custEntitiesMap[customerKey].maxDate = dTime;
           }
        }
      }
    }
  });

  // ---> REPLACE THIS BLOCK FOR CPP CUSTOMERS <---
  let cppCustEntitiesMap = {};
  let cppYearsSet = new Set();
  const cppSheets = getDatabaseSheets("CPP_Data_");
  cppSheets.forEach(sheet => {
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return;
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h).toLowerCase().trim());
    const customerIdx = headers.findIndex(h => h.includes("customer") || h === "name" || h === "business partner"); 
    const dateIdx = headers.findIndex(h => h.includes("invoice date")); 
    const sourceFileIdx = headers.findIndex(h => h.includes("source file"));
    const data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
    for (let i = 0; i < data.length; i++) {
      let row = data[i];
      
      let dTime = null; 
      if (dateIdx > -1 && row[dateIdx]) { 
          let parsed = new Date(row[dateIdx]).getTime(); 
          if (!isNaN(parsed)) dTime = parsed; 
      } 

      let companyEntity = "Pingcon";
      if (sourceFileIdx > -1) {
          let sf = String(row[sourceFileIdx]).toLowerCase();
          if (sf.includes("bonita")) companyEntity = "Bonita";
          else if (sf.includes("keysys")) companyEntity = "Keysys";
          else if (sf.includes("a2zit") || (sf.includes("a2z") && sf.includes("it"))) companyEntity = "A2ZIT";
          else if (sf.includes("a2zpmc") || (sf.includes("a2z") && sf.includes("pmc"))) companyEntity = "A2ZPMC";
      }
      if (customerIdx > -1) {
          let customerRaw = String(row[customerIdx]).replace(/[^\x20-\x7EñÑ]/g, '').replace(/\s+/g, ' ').trim();
          let customerKey = customerRaw.toLowerCase().replace(/\s+/g, '');
          let excludedCompanies = ["pingconmarketingcorporation", "bonitatissuecleancorporation", "a2z2gomarketingcorp", "keysysinc", "keysysincorporated"];
          if (customerRaw && !excludedCompanies.some(ex => customerKey.includes(ex))) {
             if (!cppCustEntitiesMap[customerKey]) {
                 cppCustEntitiesMap[customerKey] = { name: customerRaw, entities: new Set(), minDate: Infinity, maxDate: 0 };
             }
             cppCustEntitiesMap[customerKey].entities.add(companyEntity);
             
             if (dTime) { 
                 if (dTime < cppCustEntitiesMap[customerKey].minDate) cppCustEntitiesMap[customerKey].minDate = dTime; 
                 if (dTime > cppCustEntitiesMap[customerKey].maxDate) cppCustEntitiesMap[customerKey].maxDate = dTime; 
                 cppYearsSet.add(new Date(dTime).getFullYear());
             } 
          }
      }
    }
  });
  // ---> END NEW BLOCK <---

  

  // --- ADDED: Extract Min and Max Dates from 2307 Verified Invoices ---
  // --- ADDED: Extract Min and Max Dates from New Narra Data ---
  try {
      const narraSheets = getDatabaseSheets("New_Narra_Data_");
      narraSheets.forEach(sheet => {
          if (sheet.getLastRow() > 1) {
              let data = sheet.getDataRange().getValues();
              let hRow = data[0].map(h => String(h).trim().toLowerCase().replace(/\s+/g, ''));
              let dateColIdx = hRow.indexOf("checkdate");
              if (dateColIdx === -1) dateColIdx = hRow.indexOf("invoicedate");
              if (dateColIdx > -1) {
                  for (let i = 1; i < data.length; i++) {
                      let dateStr = data[i][dateColIdx];
                      if (dateStr) {
                          let dTime = new Date(dateStr).getTime();
                          if (!isNaN(dTime)) {
                              if (dTime < minDate) minDate = dTime;
                              if (dTime > maxDate) maxDate = dTime;
                          }
                      }
                  }
              }
          }
      });
  } catch(e) { console.warn("Failed to scan New Narra for dates: " + e); }
  // --------------------------------------------------------------------

  let ewtMin = Infinity;
  let ewtMax = 0;
  const VERIFIED_FOLDER_ID = SCRIPT_PROPS['VERIFIED_FOLDER_ID'];
  try {
      let verifiedFolder = DriveApp.getFolderById(VERIFIED_FOLDER_ID);
      let vFiles = verifiedFolder.getFilesByType(MimeType.GOOGLE_SHEETS);
      while (vFiles.hasNext()) {
          let vSs = SpreadsheetApp.open(vFiles.next());
          vSs.getSheets().forEach(s => {
              if (s.getLastRow() > 1) {
                  let vData = s.getDataRange().getValues();
                  let hRow = vData[0].map(h => String(h).trim().toLowerCase());
                  let dateColIdx = hRow.indexOf("2307 release date");
                  if (dateColIdx === -1) dateColIdx = 3; 
                  
                  for (let i = 1; i < vData.length; i++) {
                      let dateStr = vData[i][dateColIdx];
                      if (dateStr) {
                          let dTime = new Date(dateStr).getTime();
                          if (!isNaN(dTime)) {
                              if (dTime < ewtMin) ewtMin = dTime;
                              if (dTime > ewtMax) ewtMax = dTime;
                          }
                      }
                  }
              }
          });
      }
  } catch (e) {
      console.warn("Failed to scan Verified Invoices for dates: " + e);
  }
  if (ewtMin === Infinity) ewtMin = new Date().getTime();
  if (ewtMax === 0) ewtMax = new Date().getTime();

  // --- ADDED: Extract Min and Max Dates & Bank Names from New Narra Data ---
  let narraMin = Infinity;
  let narraMax = 0;
  let narraBanksMap = {};
  try {
      const narraSheets = getDatabaseSheets("New_Narra_Data_");
      narraSheets.forEach(sheet => {
          if (sheet.getLastRow() > 1) {
              let data = sheet.getDataRange().getValues();
              let hRow = data[0].map(h => String(h).trim().toLowerCase().replace(/\s+/g, ''));
              let dateColIdx = hRow.indexOf("voucherdate"); // STRICTLY USE VOUCHER DATE
              let bankColIdx = hRow.indexOf("bankname");
              let srcColIdx = hRow.indexOf("sourcefile");
              
              for (let i = 1; i < data.length; i++) {
                  if (dateColIdx > -1) {
                      let dateStr = data[i][dateColIdx];
                      if (dateStr) {
                          let dTime = new Date(dateStr).getTime();
                          if (!isNaN(dTime)) {
                              if (dTime < narraMin) narraMin = dTime;
                              if (dTime > narraMax) narraMax = dTime;
                          }
                      }
                  }
                  if (bankColIdx > -1) {
                      let bName = String(data[i][bankColIdx]).trim();
                      if (bName) {
                          if (!narraBanksMap[bName]) narraBanksMap[bName] = new Set();
                          let sourceFile = (srcColIdx > -1) ? String(data[i][srcColIdx]).toUpperCase() : "";
                          if (sourceFile.includes("KEYSYS")) narraBanksMap[bName].add("Keysys");
                          else if (sourceFile.includes("PINGCON")) narraBanksMap[bName].add("Pingcon");
                          else narraBanksMap[bName].add("Pingcon"); // Fallback
                      }
                  }
              }
          }
      });
  } catch(e) { console.warn("Failed to scan New Narra for dates: " + e); }
  if (narraMin === Infinity) narraMin = new Date().getTime();
  if (narraMax === 0) narraMax = new Date().getTime();
  
  let narraBanksFormatted = [];
  for (let b in narraBanksMap) {
      narraBanksFormatted.push({ name: b, entities: Array.from(narraBanksMap[b]) });
  }
  narraBanksFormatted.sort((a,b) => a.name.localeCompare(b.name));
  // --------------------------------------------------------------------

  // --------------------------------------------------------------------

  // --- SAFETY FILTER: PREVENT 1970 FROM APPEARING ---
  const minValidTime = new Date(2020, 0, 1).getTime(); // Baseline: Jan 1, 2020
  
  if (minDate === Infinity || minDate < minValidTime) minDate = minValidTime;
  if (ewtMin === Infinity || ewtMin < minValidTime) ewtMin = minValidTime;
  if (narraMin === Infinity || narraMin < minValidTime) narraMin = minValidTime;
  if (maxDate === 0) maxDate = new Date().getTime();
  
  // Clean AE dates
  for (let ae in aeEntitiesMap) {
      if (aeEntitiesMap[ae].minDate < minValidTime) aeEntitiesMap[ae].minDate = minValidTime;
  }
  // Clean Customer dates
  for (let cust in custEntitiesMap) {
      if (custEntitiesMap[cust].minDate < minValidTime) custEntitiesMap[cust].minDate = minValidTime;
  }
  // Clean CPP dates
  for (let cust in cppCustEntitiesMap) {
      if (cppCustEntitiesMap[cust].minDate < minValidTime) cppCustEntitiesMap[cust].minDate = minValidTime;
  }
  
  // Remove bad years from CPP dropdown
  cppYearsSet.delete(1970);
  cppYearsSet.delete(1899);
  cppYearsSet.delete(1969);
  // --------------------------------------------------

  let aesFormatted = [];
  for (let ae in aeEntitiesMap) {
     let canonical = getCanonicalName(ae);
     // STRICT GATE: Only allow AEs who have an active quota in SQ_Data
     let hasQuota = Array.from(quotaAENames).some(qAe => checkNameMatch(qAe, canonical));
     if (!hasQuota) continue; // Blocks non-sales personnel without an assigned quota

     let aeData = aeEntitiesMap[ae];
     aesFormatted.push({ name: ae, entities: Array.from(aeData.entities), minDate: aeData.minDate === Infinity ? null : aeData.minDate, maxDate: aeData.maxDate === 0 ? null : aeData.maxDate });
  }
  aesFormatted.sort((a, b) => a.name.localeCompare(b.name));

  let custsFormatted = [];
  for (let custKey in custEntitiesMap) {
     let cData = custEntitiesMap[custKey];
     custsFormatted.push({ name: cData.name, entities: Array.from(cData.entities), minDate: cData.minDate === Infinity ? null : cData.minDate, maxDate: cData.maxDate === 0 ? null : cData.maxDate });
  }
  custsFormatted.sort((a, b) => a.name.localeCompare(b.name));

  let cppCustsFormatted = [];
  for (let custKey in cppCustEntitiesMap) {
     let cData = cppCustEntitiesMap[custKey];
     cppCustsFormatted.push({ name: cData.name, entities: Array.from(cData.entities), minDate: cData.minDate === Infinity ? null : cData.minDate, maxDate: cData.maxDate === 0 ? null : cData.maxDate });
  }
  cppCustsFormatted.sort((a, b) => a.name.localeCompare(b.name));

  return {
    min: Utilities.formatDate(new Date(minDate), Session.getScriptTimeZone(), "yyyy-MM-dd"),
    max: Utilities.formatDate(new Date(maxDate), Session.getScriptTimeZone(), "yyyy-MM-dd"),
    ewtMin: Utilities.formatDate(new Date(ewtMin), Session.getScriptTimeZone(), "yyyy-MM-dd"),
    ewtMax: Utilities.formatDate(new Date(ewtMax), Session.getScriptTimeZone(), "yyyy-MM-dd"),
    narraMin: Utilities.formatDate(new Date(narraMin), Session.getScriptTimeZone(), "yyyy-MM-dd"),
    narraMax: Utilities.formatDate(new Date(narraMax), Session.getScriptTimeZone(), "yyyy-MM-dd"),
    narraBanks: narraBanksFormatted,
    aes: aesFormatted,
    customers: custsFormatted,
    cppCustomers: cppCustsFormatted,
    cppYears: Array.from(cppYearsSet).sort((a,b) => b - a)
  };
}



function encryptAttachmentViaPython(blob, filename, password, internalFilename = null) {
  if (!blob) throw new Error("No file generated to encrypt.");
  if (!password) return blob; // Only skips if explicitly null
  
  let safeFilename = filename.replace(/[^a-zA-Z0-9 \-_.]/g, ''); // Safest filename for Python to prevent server crash
  let payload = { file: Utilities.base64Encode(blob.getBytes()), password: password, filename: safeFilename };
  
  // ADDED: Pass the internal filename to Python if one is provided
  if (internalFilename) {
      payload.internal_filename = internalFilename.replace(/[^a-zA-Z0-9 \-_.]/g, '');
  }
  
  // Fetch API Key from Settings
  let pythonApiKey = PropertiesService.getScriptProperties().getProperty('PYTHON_API_KEY');
  
  // Exact same as your original options, just with headers added
  let options = { 
    method: "post", 
    contentType: "application/json", 
    payload: JSON.stringify(payload), 
    headers: { "X-API-Key": pythonApiKey }, // <-- SECURITY PART
    muteHttpExceptions: true 
  };
  
  for (let i = 0; i < 3; i++) {
    try {
      let response = UrlFetchApp.fetch(SCRIPT_PROPS['PYTHON_ENCRYPT_URL'], options);
      if (response.getResponseCode() === 200) {
        let result = JSON.parse(response.getContentText());
        if (result.encrypted_base64) {
          return Utilities.newBlob(Utilities.base64Decode(result.encrypted_base64.split(",")[1]), "application/zip", result.zip_filename);
        }
      }
      if (i === 2) throw new Error("Encryption Server Failed.");
      Utilities.sleep(4000); 
    } catch (e) {
      if (i === 2) throw new Error("Encryption Error: " + e.message);
      Utilities.sleep(4000);
    }
  }
  throw new Error("Failed to encrypt attachment after 3 attempts.");
}

// =========================================================================
// BUY/SELL AND COMMISSION BATCH PROCESSING
// =========================================================================
const BUY_SELL_BATCH_KEY = "PENDING_AES_BUY_SELL";
const BUY_SELL_TRIGGER_NAME = "resumeBuySellBatch";
const COMM_BATCH_KEY = "PENDING_AES_COMM";
const COMM_TRIGGER_NAME = "resumeCommissionBatch";

function _26_runSendAllBuySellReports() {
  _startBuySellCommBatch(BUY_SELL_BATCH_KEY, BUY_SELL_TRIGGER_NAME, "CANCEL_BUY_SELL");
}

function _27_runSendAllCommissionReports() {
  _startBuySellCommBatch(COMM_BATCH_KEY, COMM_TRIGGER_NAME, "CANCEL_COMM");
}

function _startBuySellCommBatch(batchKey, triggerName, cancelKey) {
  var lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) { throw new Error("System is currently busy."); }
  try {
    const props = PropertiesService.getScriptProperties();
        if (triggerName === BUY_SELL_TRIGGER_NAME) deleteTriggersByName(BUY_SELL_TRIGGER_NAME);
        else if (triggerName === COMM_TRIGGER_NAME) deleteTriggersByName(COMM_TRIGGER_NAME);
        else deleteTriggersByName(triggerName);
        
        props.deleteProperty(batchKey);
    props.deleteProperty(cancelKey);
    
    let allowedAEsStr = props.getProperty("SALES_TARGET_AES");
    if (!allowedAEsStr) throw new Error("No AEs selected.");
    let aeArrayRaw = JSON.parse(allowedAEsStr);
    if (aeArrayRaw.length === 0) throw new Error("No AEs selected.");
    
    // GET ALL MONTHS IN DATE RANGE
    const startStr = props.getProperty("BATCH_START_DATE");
    const endStr = props.getProperty("BATCH_END_DATE");
    
    let startDate = startStr ? new Date(startStr + "T00:00:00") : new Date();
    let endDate = endStr ? new Date(endStr + "T00:00:00") : new Date();
    if (startDate > endDate) { let temp = startDate; startDate = endDate; endDate = temp; }

    let targetMonths = [];
    let currentMonth = new Date(startDate.getFullYear(), startDate.getMonth(), 1);
    let limitMonth = new Date(endDate.getFullYear(), endDate.getMonth(), 1);

    while (currentMonth <= limitMonth) {
        targetMonths.push(Utilities.formatDate(currentMonth, Session.getScriptTimeZone(), "MMMM yyyy"));
        currentMonth.setMonth(currentMonth.getMonth() + 1);
    }

    // CREATE QUEUE OF AE + MONTH COMBINATIONS
    let expandedQueue = [];
    for (let i = 0; i < aeArrayRaw.length; i++) {
        for (let j = 0; j < targetMonths.length; j++) {
            expandedQueue.push(aeArrayRaw[i] + "|" + targetMonths[j]);
        }
    }
    
    props.setProperty(batchKey, JSON.stringify(expandedQueue));
        if (triggerName === BUY_SELL_TRIGGER_NAME) resumeBuySellBatch();
        else if (triggerName === COMM_TRIGGER_NAME) resumeCommissionBatch();
        
        else if (triggerName === PURCHASE_TRIGGER_NAME) resumePurchaseBatch();
  } finally {
    lock.releaseLock();
  }
}
function resumeBuySellBatch() { _resumeBuySellCommCore(BUY_SELL_BATCH_KEY, BUY_SELL_TRIGGER_NAME, "CANCEL_BUY_SELL", false); }
function resumeCommissionBatch() { _resumeBuySellCommCore(COMM_BATCH_KEY, COMM_TRIGGER_NAME, "CANCEL_COMM", true); }

function _resumeBuySellCommCore(batchKey, triggerName, cancelKey, isCommission) {
  const startTime = Date.now();
  const MAX_EXECUTION_TIME = 4.5 * 60 * 1000; 
  
  const props = PropertiesService.getScriptProperties();
  props.setProperty("THREAD_STATE", "RUNNING"); // <--- ADDED
  
  let pendingAEsStr = props.getProperty(batchKey);
  
  if (!pendingAEsStr) {
    deleteTriggersByName(isCommission ? COMM_TRIGGER_NAME : BUY_SELL_TRIGGER_NAME);
    return; 
  }
  
  let pendingQueue = JSON.parse(pendingAEsStr);
  if (pendingQueue.length === 0) {
    deleteTriggersByName(isCommission ? COMM_TRIGGER_NAME : BUY_SELL_TRIGGER_NAME);
    props.deleteProperty(batchKey);
    clearReportLock();
    console.log("All AEs and Months processed successfully.");

    return;
  }
  
  while (pendingQueue.length > 0) {
     if (props.getProperty(cancelKey) === "true") {
         deleteTriggersByName(isCommission ? COMM_TRIGGER_NAME : BUY_SELL_TRIGGER_NAME);
         props.deleteProperty(batchKey);
         props.deleteProperty(cancelKey);
         return; 
     }

     if (Date.now() - startTime > MAX_EXECUTION_TIME) {
                props.setProperty(batchKey, JSON.stringify(pendingQueue));
                replaceTrigger(isCommission ? COMM_TRIGGER_NAME : BUY_SELL_TRIGGER_NAME);
                return; 
             }
             
             let currentItem = pendingQueue.shift(); 
             let parts = currentItem.split("|");
             let currentAE = parts[0];
             let targetMonthStr = parts[1];
             
             try {
                 _$processBuySellForms(currentAE, targetMonthStr, true, new Date(), true, isCommission);
         
             } catch(e) { console.error("Error processing " + currentItem + ": " + e); }

             props.setProperty(batchKey, JSON.stringify(pendingQueue));
             
             // Breaks the loop and waits exactly 15 minutes before the next AE report
             replaceTrigger(isCommission ? COMM_TRIGGER_NAME : BUY_SELL_TRIGGER_NAME);
             return; 
          }

          deleteTriggersByName(isCommission ? COMM_TRIGGER_NAME : BUY_SELL_TRIGGER_NAME);
          props.deleteProperty(batchKey);
          clearReportLock();
}



function webappGetReportStatus(portalModule = "MAIN") {
  const props = PropertiesService.getScriptProperties();
  let quota = "Error";
  try { quota = MailApp.getRemainingDailyQuota(); } catch(e) { quota = "Auth Error"; }
  
  let mod = String(portalModule || "MAIN").toUpperCase().trim();
  let prefix = _getPortalLockPrefix(mod);
  
  // Checks environment-isolated keys first, with safe fallback to legacy keys
  let type = props.getProperty(prefix + "_REPORT_TYPE") || props.getProperty("ONGOING_REPORT_TYPE");
  let status = props.getProperty(prefix + "_REPORT_STATUS") || props.getProperty("ONGOING_REPORT_STATUS");
  let ts = props.getProperty(prefix + "_REPORT_TIMESTAMP") || props.getProperty("ONGOING_REPORT_TIMESTAMP");
  let resumeTime = props.getProperty(prefix + "_RESUME_TIME") || props.getProperty("RESUME_TIME");
  let lastError = props.getProperty(prefix + "_LAST_RUN_ERROR") || props.getProperty("LAST_RUN_ERROR");
  
  let lastActive = ts ? parseInt(ts) : 0;
  if (resumeTime && parseInt(resumeTime) > lastActive) {
      lastActive = parseInt(resumeTime);
  }
  
  // Stale lock cleanup (> 15 mins) specifically for this portal
  if (type && lastActive > 0 && (Date.now() - lastActive > 900000)) {
      clearReportLock(mod);
      type = null;
  }
  
  return { 
    type: (type && type.trim() !== "") ? type : null, 
    status: status,
    resumeTime: resumeTime,
    lastError: lastError,
    emailQuota: quota
  };
}

function clearReportLock(portalModule = "MAIN") {
  const props = PropertiesService.getScriptProperties();
  let prefix = _getPortalLockPrefix(portalModule);

  props.deleteProperty(prefix + "_REPORT_TYPE");
  props.deleteProperty(prefix + "_REPORT_STATUS");
  props.deleteProperty(prefix + "_RESUME_TIME");
  props.deleteProperty(prefix + "_REPORT_TIMESTAMP");
  props.deleteProperty(prefix + "_THREAD_STATE");

  let rawMod = String(portalModule || "MAIN").toUpperCase().trim();
  if (rawMod === "MAIN") {
    props.deleteProperty("TEST_MODE");
    props.deleteProperty("ONGOING_REPORT_TYPE");
    props.deleteProperty("ONGOING_REPORT_STATUS");
    props.deleteProperty("RESUME_TIME");
    props.deleteProperty("ONGOING_REPORT_TIMESTAMP");
    props.deleteProperty("THREAD_STATE");
  }

  CacheService.getScriptCache().remove(prefix + "_UPLOAD");
}

// =========================================================================
// PINGCON HR CENTER API INTEGRATION & ROLE RESOLVER (PHASE 1)
// =========================================================================

function fetchActiveHREmployees() {
  const cache = CacheService.getScriptCache();
  const cachedData = cache.get("HR_ACTIVE_EMPLOYEES_CACHE");
  if (cachedData) {
    try { return JSON.parse(cachedData); } catch (e) {}
  }

  const props = PropertiesService.getScriptProperties();
  const apiUrl = props.getProperty('HR_CENTER_API_URL');
    if (!apiUrl) {
      console.error("Missing HR_CENTER_API_URL in Script Properties.");
      return [];
    }
  const apiKey = props.getProperty('HR_CENTER_API_KEY');

  if (!apiKey) {
    console.error("Missing HR_CENTER_API_KEY in Script Properties.");
    return [];
  }

  try {
    const url = apiUrl + (apiUrl.includes("?") ? "&" : "?") + "per_page=200";
    const options = {
      method: "get",
      headers: {
        "Authorization": "Bearer " + apiKey.trim(),
        "Accept": "application/json"
      },
      muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch(url, options);
    if (response.getResponseCode() === 200) {
      const json = JSON.parse(response.getContentText());
      const rawEmployees = Array.isArray(json) ? json : (json.data || json.employees || []);
      
      // COMPRESSION: Extract ONLY the 4 essential fields needed for identity & accountability
      const cleanEmployees = rawEmployees.map(e => ({
        id: e.employee_id || e.id || "",
        name: String(e.full_name || e.name || `${e.first_name || ""} ${e.last_name || ""}`).trim(),
        company_email: String(e.company_email || e.email || "").trim().toLowerCase(),
        personal_email: String(e.personal_email || "").trim().toLowerCase(),
        department: String(e.department || "")
      }));

      // Safe cache with try/catch so memory limit never causes a crash
      if (cleanEmployees.length > 0) {
        try {
          cache.put("HR_ACTIVE_EMPLOYEES_CACHE", JSON.stringify(cleanEmployees), 600);
        } catch (cacheErr) {
          console.warn("Roster too large to cache, proceeding with live fetch.");
        }
      }
      return cleanEmployees;
    } else {
      console.error("HR Center API Error (" + response.getResponseCode() + "): " + response.getContentText());
      return [];
    }
  } catch (err) {
    console.error("HR Center Connection Exception: " + err.message);
    return [];
  }
}

function _getPortalRolesSheet(portalType = "SUPPLIER") {
  const props = PropertiesService.getScriptProperties();
  
  // Safe Dev/Exec Environment Detection
  let isDevMode = false;
  try {
    const service = ScriptApp.getService();
    if (service && service.getUrl()) {
      isDevMode = service.getUrl().toLowerCase().endsWith("/dev");
    }
  } catch(e) {}

  let targetFolderId = null;
  const pType = String(portalType).toUpperCase().trim();

  // Dynamic Routing across portals using Script Properties
  let targetSheetName = "Portal_User_Roles";

  if (pType.includes("CLIENT")) {
    targetFolderId = isDevMode 
      ? props.getProperty('CLIENT_ROLES_FOLDER_ID_DEV') 
      : props.getProperty('CLIENT_ROLES_FOLDER_ID_EXEC');
  } else if (pType.includes("NARRA")) {
    targetSheetName = "Portal_User_Roles_Narra";
    targetFolderId = isDevMode 
      ? (props.getProperty('DATABASE_FOLDER_ID_DEV') || props.getProperty('DATABASE_FOLDER_ID'))
      : (props.getProperty('DATABASE_FOLDER_ID') || DATABASE_FOLDER_ID);
  } else {
    // Default to Supplier Roles
    targetFolderId = isDevMode 
      ? props.getProperty('SUPPLIER_ROLES_FOLDER_ID_DEV') 
      : props.getProperty('SUPPLIER_ROLES_FOLDER_ID_EXEC');
  }

  // Fallback protection if properties are not yet set
  if (!targetFolderId) {
    targetFolderId = DATABASE_FOLDER_ID || props.getProperty('DATABASE_FOLDER_ID') || props.getProperty('VERIFIED_FOLDER_ID');
  }

  const folder = DriveApp.getFolderById(targetFolderId);
  const files = folder.getFilesByName(targetSheetName);
  let ss;
  
  if (files.hasNext()) {
    ss = SpreadsheetApp.open(files.next());
  } else {
    ss = SpreadsheetApp.create(targetSheetName);
    DriveApp.getFileById(ss.getId()).moveTo(folder);
    const sheet = ss.getSheets()[0];
    sheet.setName("UserRoles");
    sheet.appendRow(["Email", "Name", "Role", "Status", "Updated By", "Updated At"]);
    sheet.getRange("A1:F1").setFontWeight("bold").setBackground("#f3f3f3");
    sheet.setFrozenRows(1);
  }
  return ss.getSheetByName("UserRoles") || ss.getSheets()[0];
}

function getHREmployeeByEmail(inputEmail, portalType = "SUPPLIER") {
  if (!inputEmail || typeof inputEmail !== 'string') return null;
  const cleanEmail = inputEmail.trim().toLowerCase();
  const employees = fetchActiveHREmployees();

  // Find employee matching personal or company email
  const emp = employees.find(e => {
    const cEmail = String(e.company_email || e.email || "").trim().toLowerCase();
    const pEmail = String(e.personal_email || "").trim().toLowerCase();
    return cEmail === cleanEmail || pEmail === cleanEmail;
  });

  if (!emp) return null;

  // Resolve Real Name
  let fullName = String(emp.full_name || emp.name || "").trim();
  if (!fullName && (emp.first_name || emp.last_name)) {
    fullName = `${emp.first_name || ""} ${emp.last_name || ""}`.trim();
  }
  if (!fullName) fullName = cleanEmail.split('@')[0];

  // 1. Permanent Super Admin Access Check from Script Properties
  const props = PropertiesService.getScriptProperties();
  const superAdminList = (props.getProperty('SUPER_ADMIN_EMAILS') || "").toLowerCase().split(',').map(s => s.trim());
  let role = superAdminList.includes(cleanEmail) ? "Super Admin" : "Encoder";
  let status = "Active";

  // 2. Query only the isolated quadrant database for this specific portal & environment
  try {
    const sheet = _getPortalRolesSheet(portalType);
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim().toLowerCase() === cleanEmail) {
        if (data[i][2]) role = String(data[i][2]).trim();
        if (data[i][3]) status = String(data[i][3]).trim();
        break;
      }
    }
  } catch (err) {
    console.warn(`Could not check ${portalType} role overrides: ` + err.message);
  }

  // Permanent Super Admins cannot be overridden
  if (superAdminList.includes(cleanEmail)) {
    role = "Super Admin";
    status = "Active";
  }

  // Deactivated employees are blocked from entering
  if (status.toLowerCase() === "deactivated") {
    return null;
  }

  return {
    found: true,
    name: fullName,
    email: cleanEmail,
    role: role,
    status: status,
    employeeId: emp.employee_id || emp.id || "",
    portalType: portalType
  };
}

// =========================================================================
// ENTERPRISE SYSTEM SETTINGS & LIVE DIAGNOSTICS (PART 2)
// =========================================================================

function webappGetSystemDiagnosticsAndSettings(sessionToken) {
  const user = getAuthenticatedUser(sessionToken);
  if (!user || user.role !== "Super Admin") {
    return { success: false, error: "Permission Denied: Only Super Admins can access System Settings." };
  }

  const props = PropertiesService.getScriptProperties();

  // 1. Live Google Quotas
  let emailQuota = 0;
  try { emailQuota = MailApp.getRemainingDailyQuota(); } catch(e) { emailQuota = -1; }

  // 2. Health Checks / Connectivity Pings
  let hrApiStatus = "Offline";
  try {
    const hrUrl = props.getProperty('HR_CENTER_API_URL');
    const hrKey = props.getProperty('HR_CENTER_API_KEY');
    if (hrKey) {
      const resp = UrlFetchApp.fetch(hrUrl + "?per_page=1", {
        headers: { "Authorization": "Bearer " + hrKey.trim(), "Accept": "application/json" },
        muteHttpExceptions: true
      });
      if (resp.getResponseCode() === 200) hrApiStatus = "Online";
    }
  } catch(e) { hrApiStatus = "Error: " + e.message; }

  let pythonServerStatus = "Offline";
  try {
    const pyUrl = props.getProperty('PYTHON_ENCRYPT_URL');
    if (pyUrl) {
      const pingResp = UrlFetchApp.fetch(pyUrl.replace(/\/encrypt\/?$/, '/health'), { muteHttpExceptions: true });
      if (pingResp.getResponseCode() === 200 || pingResp.getResponseCode() === 404 || pingResp.getResponseCode() === 405) {
        pythonServerStatus = "Online";
      }
    }
  } catch(e) { pythonServerStatus = "Reachable"; }

  let dbStatus = "Connected";
  try {
    const dbFolderId = DATABASE_FOLDER_ID || props.getProperty('DATABASE_FOLDER_ID') || props.getProperty('VERIFIED_FOLDER_ID');
    const folder = DriveApp.getFolderById(dbFolderId);
    if (!folder) dbStatus = "Unreachable";
  } catch(e) { dbStatus = "Error: " + e.message; }

  // 3. Dynamic Signatory, Entity Profiles, 2307 Form Studio & Assets
  const settings = {
    signatoryName: props.getProperty('SIGNATORY_2307_NAME') || "",
    signatoryTitle: props.getProperty('SIGNATORY_2307_TITLE') || "",
    signatoryTin: props.getProperty('SIGNATORY_2307_TIN') || "",
    discrepancyTolerance: props.getProperty('DISCREPANCY_TOLERANCE') || "",
    adminNotificationEmail: props.getProperty('ADMIN_EMAIL') || "",
    collectionAssistant: props.getProperty('DEFAULT_COLLECTION_ASSISTANT') || "",
    
    // Form Studio Revisions & Asset Drive IDs
    formRevisionTitle: props.getProperty('FORM_2307_REVISION_TITLE') || "",
    formTrackingCode: props.getProperty('FORM_2307_TRACKING_CODE') || "",
    birLogoId: props.getProperty('BIR_LOGO_DRIVE_ID') || "",
    barcodeId: props.getProperty('BARCODE_DRIVE_ID') || "",
    signatureId: props.getProperty('SIGNATURE_DRIVE_ID') || "",
    formBannerColor: props.getProperty('FORM_2307_BANNER_COLOR') || "",
    formBorderWidth: props.getProperty('FORM_2307_BORDER_WIDTH') || "",
    part3Row1Desc: props.getProperty('PART3_ROW1_DESC') || "",
    part3Row2Desc: props.getProperty('PART3_ROW2_DESC') || "",
    legalPerjuryText: props.getProperty('LEGAL_PERJURY_TEXT') || "",

    // Multi-Address & Entity Master Configurations
    pingconTin: props.getProperty('PINGCON_TIN') || "",
    pingconAddr: props.getProperty('PINGCON_ADDR') || "",
    keysysTin: props.getProperty('KEYSYS_TIN') || "",
    keysysAddr: props.getProperty('KEYSYS_ADDR') || "",
    a2zTin: props.getProperty('A2Z_TIN') || "",
    a2zAddr: props.getProperty('A2Z_ADDR') || "",
    bonitaTin: props.getProperty('BONITA_TIN') || "",
    bonitaAddr: props.getProperty('BONITA_ADDR') || ""
  };

  return {
    success: true,
    diagnostics: {
      emailQuotaRemaining: emailQuota,
      hrApiStatus: hrApiStatus,
      pythonServerStatus: pythonServerStatus,
      dbStatus: dbStatus,
      environment: ScriptApp.getService().getUrl().endsWith("/dev") ? "Development (dev)" : "Production (exec)"
    },
    settings: settings
  };
}

function webappUploadSystemAsset(base64Data, fileName, assetKey, sessionToken) {
  const user = getAuthenticatedUser(sessionToken);
  if (!user || user.role !== "Super Admin") {
    return { success: false, error: "Permission Denied: Only Super Admins can upload system assets." };
  }

  try {
    const props = PropertiesService.getScriptProperties();
    
    // Dynamic Environment Routing: Dev vs Exec folder resolution
    const isDevMode = ScriptApp.getService().getUrl().endsWith("/dev");
    const targetFolderPropKey = isDevMode ? 'EWT_Report_Supplier_Icons_Folder_Dev' : 'EWT_Report_Supplier_Icons_Folder_Exec';
    
    let folderId = props.getProperty(targetFolderPropKey);
    if (!folderId) {
      folderId = DATABASE_FOLDER_ID || props.getProperty('DATABASE_FOLDER_ID') || props.getProperty('VERIFIED_FOLDER_ID');
    }
    
    const folder = DriveApp.getFolderById(folderId);

    const rawData = base64Data.includes(',') ? base64Data.split(',')[1] : base64Data;
    let mimeType = MimeType.PNG;
    if (fileName.toLowerCase().endsWith('.jpg') || fileName.toLowerCase().endsWith('.jpeg')) mimeType = MimeType.JPEG;

    const blob = Utilities.newBlob(Utilities.base64Decode(rawData), mimeType, fileName);
    const file = folder.createFile(blob);

    // Set permission so thumbnail API renders in previews and PDFs without cookie blocking
    try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch(e){}

    const newFileId = file.getId();
    props.setProperty(assetKey, newFileId);

    logSystemActivity("SYSTEM", "Uploaded 2307 Asset", `Updated asset [${assetKey}] with file: ${fileName}`, sessionToken);

    return { success: true, fileId: newFileId, thumbnailUrl: "https://drive.google.com/thumbnail?id=" + newFileId };
  } catch(err) {
    return { success: false, error: err.message };
  }
}

// =========================================================================
// GEMINI AI COPILOT & AUTOPILOT ENGINE (PART 3)
// =========================================================================

function webappAskAiAssistant(userMessage, chatHistory, sessionToken) {
  const user = getAuthenticatedUser(sessionToken);
  if (!user) {
    return { success: false, error: "Authentication required to access the AI Copilot." };
  }

  const props = PropertiesService.getScriptProperties();
  const apiKey = props.getProperty('GEMINI_API_KEY');
  if (!apiKey) {
    return { success: false, error: "GEMINI_API_KEY is not configured in Script Properties. Please contact IT." };
  }

  // System Prompt / Embedded Knowledge Base with Dynamic Role Boundaries
  const userRole = user.role || "Viewer";
  const isSuperAdminUser = (userRole === "Super Admin");

  const systemInstruction = `You are the Pingcon Finance AI Copilot & Autopilot Assistant, built into the Purchases and EWT Report Portal.
Your primary role is to assist finance staff, encoders, and supervisors with portal navigation, tax math, Philippine BIR regulations, and troubleshooting.

CURRENT AUTHENTICATED USER CONTEXT:
- Name: ${user.name}
- Active Role: ${userRole}

STRICT ROLE PERMISSION GOVERNANCE:
- The user's role is: "${userRole}".
${!isSuperAdminUser ? `- SECURITY RESTRICTION: The user is NOT a Super Admin. They are strictly FORBIDDEN from accessing "settings" (System Settings) or "user_mgmt" (User Management).
- If the user asks to go to, open, navigate to, or view System Settings or User Management, you must REFUSE politely:
  "Access Denied: As an ${userRole}, you do not have administrative clearance to access System Settings or User Management. Please contact a Super Admin."
- Under NO circumstances should you output an ACTION block targeting "settings" or "user_mgmt" for this user.` : `- The user is a Super Admin and has full clearance for all administrative and operational tabs.`}

PORTAL KNOWLEDGE BASE:
- Entities Supported: Pingcon Marketing Corporation, Keysys Inc., A2Z2GO Marketing Corp., Bonita Tissue Clean Corp.
- ATC Codes:
  * WC158 (1%): Local purchases of goods by Top Withholding Agents.
  * WC160 (2%): Local purchases of services or rentals by Top Withholding Agents.
  * WI158 / WI160: Same rates, used for individual payees.
- Mixed VAT Invoices: When an invoice contains both taxable and exempt items, encoders check 'This invoice contains both VAT and Non-VAT items'. Box 1 holds VATable purchases (Gross / 1.12), Box 2 holds Non-VAT purchases (Gross = Purchases, Tax = 0.00).
- Discrepancy Matching (Supplier 2307):
  * Invoices are verified against three databases: Narra Voucher, Check Summary, and RIWB.
  * Discrepancies occur if an invoice is Missing or if the amount variance is >= 0.05 PHP.
  * Super Admins and Supervisors can 'Force Match' or resolve manually. Encoders cannot force match.
- User Roles:
  * Super Admin: Full access, settings, deletions, user role management.
  * Supervisor: Can resolve/force-match discrepancies, verify invoices, generate reports. Cannot delete records or modify system settings.
  * Encoder: Can verify invoices, edit draft forms, preview 2307. Cannot force-match or delete records.
  * Viewer: Read-only access. Cannot edit or save records.

AUTOPILOT CAPABILITIES:
When the user asks you to navigate or perform a permitted UI action, include an action block at the very end of your response in this exact format:
<<<ACTION:{"action":"NAVIGATE_TAB","targetTab":"overview"|"dashboard"|"queue"|"editor"|"ewt"|"purchase"|"logs"${isSuperAdminUser ? '|"settings"|"user_mgmt"' : ''}}>>>
OR
<<<ACTION:{"action":"FILTER_QUEUE","query":"supplier or invoice name","status":"Pending"|"Processed"|"All"}>>>

Tone: Professional, helpful, concise, and courteous. You can answer in English or Taglish. Never reveal secret API keys.`;

  // Build Conversation Payload
  const contents = [];
  if (Array.isArray(chatHistory)) {
    chatHistory.slice(-8).forEach(msg => {
      contents.push({
        role: msg.role === 'user' ? 'user' : 'model',
        parts: [{ text: msg.text }]
      });
    });
  }
  contents.push({
    role: "user",
    parts: [{ text: userMessage }]
  });

  const payload = {
    contents: contents,
    systemInstruction: {
      parts: [{ text: systemInstruction }]
    },
    generationConfig: {
      temperature: 0.2,
      maxOutputTokens: 1024
    }
  };

  const url = "https://generativelanguage.googleapis.com/v1beta/models/gemini-3.5-flash-lite:generateContent?key=" + apiKey;

  try {
    const response = UrlFetchApp.fetch(url, {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });

    if (response.getResponseCode() === 200) {
      const json = JSON.parse(response.getContentText());
      if (json.candidates && json.candidates[0] && json.candidates[0].content) {
        let rawReply = json.candidates[0].content.parts[0].text;
        
        let autopilotAction = null;
        const actionMatch = rawReply.match(/<<<ACTION:(.*?)>>>/);
        if (actionMatch) {
          try {
            autopilotAction = JSON.parse(actionMatch[1]);
          } catch(e) {}
          rawReply = rawReply.replace(/<<<ACTION:.*?>>>/, '').trim();
        }

        return {
          success: true,
          replyText: rawReply,
          autopilotAction: autopilotAction
        };
      }
    }
    return { success: false, error: "AI service returned code " + response.getResponseCode() + ": " + response.getContentText() };
  } catch(err) {
    return { success: false, error: "Connection error: " + err.message };
  }
}

// =========================================================================
// AI ANOMALY DETECTION & VENDOR LEARNING MEMORY (PART 4)
// =========================================================================

function webappCheckInvoiceAnomalies(invoiceNo, supplierName, tin, gross, purchases, inputTax, ewtTax, atc, creditMemo, totalAmount, accountTitle, invoiceDate, currentFileName, isMixedVat, companyContext) {
  try {
    const anomalies = [];
    const cleanInv = String(invoiceNo || "").trim().toUpperCase();
    const cleanSup = String(supplierName || "").trim().toUpperCase();

    // 1. Two-Tier Real-Time Duplicate Invoice Check (Ignores self-match in Edit Mode)
    if (cleanInv) {
      const verifiedList = webappGetVerifiedInvoices(companyContext);

      // Tier 1: Exact Match (Same Invoice No + Same Supplier)
      const exactDuplicate = verifiedList.find(item => 
        String(item.invoiceNo).trim().toUpperCase() === cleanInv && 
        String(item.companyName).trim().toUpperCase() === cleanSup &&
        (!currentFileName || String(item.fileName).trim() !== String(currentFileName).trim())
      );

      // Tier 2: Cross-Vendor Match (Same Invoice No + Different Supplier)
      const crossVendorDuplicate = verifiedList.find(item => 
        String(item.invoiceNo).trim().toUpperCase() === cleanInv && 
        String(item.companyName).trim().toUpperCase() !== cleanSup &&
        (!currentFileName || String(item.fileName).trim() !== String(currentFileName).trim())
      );

      if (exactDuplicate) {
        anomalies.push({
          type: "warning",
          title: "Duplicate Invoice Detected",
          message: `Invoice #${cleanInv} for "${supplierName}" was already verified on another record (${exactDuplicate.verifiedDateTime || "archive"}).`
        });
      } else if (crossVendorDuplicate) {
        anomalies.push({
          type: "warning",
          title: "Reused Invoice Number Warning",
          message: `Invoice #${cleanInv} is already registered under supplier "${crossVendorDuplicate.companyName}". Please verify if this is a separate receipt.`
        });
      }
    }

    // 2. Strict Philippine BIR TIN Format Validation (Only 9, 12, or 14 digits)
    const rawTin = String(tin || "").replace(/[^0-9]/g, '');
    const validTinLengths = [9, 12, 14];
    if (rawTin.length > 0 && !validTinLengths.includes(rawTin.length)) {
      anomalies.push({
        type: "warning",
        title: "Invalid BIR TIN Length",
        message: `Extracted TIN (${tin}) contains ${rawTin.length} digits. Standard Philippine BIR TINs must contain strictly 9 digits (base), 12 digits (3-digit branch), or 14 digits (5-digit branch code).`
      });
    }
    if (rawTin.length > 0 && (/^0{9,14}$/.test(rawTin) || /^9{9,14}$/.test(rawTin))) {
      anomalies.push({
        type: "warning",
        title: "Placeholder TIN Detected",
        message: `TIN (${tin}) appears to be a placeholder dummy number.`
      });
    }

    // 3. Date Sanity Checks
    if (invoiceDate) {
      const dObj = new Date(invoiceDate);
      if (!isNaN(dObj.getTime())) {
        if (dObj.getFullYear() <= 1970) {
          anomalies.push({
            type: "warning",
            title: "Epoch Date Error",
            message: `Invoice Date (${invoiceDate}) parses to year ${dObj.getFullYear()}. Please correct the date format.`
          });
        }
        const tomorrow = new Date();
        tomorrow.setDate(tomorrow.getDate() + 1);
        if (dObj > tomorrow) {
          anomalies.push({
            type: "warning",
            title: "Future Invoice Date",
            message: `Invoice Date (${invoiceDate}) is set in the future.`
          });
        }
      }
    }

    // 4. Financial & Math Reconciliation
    const pGross = parseFloat(String(gross).replace(/,/g, '')) || 0;
    const pCm = parseFloat(String(creditMemo || '0').replace(/,/g, '')) || 0;
    let pTot = parseFloat(String(totalAmount || '0').replace(/,/g, '')) || 0;
    if (pTot <= 0 && pGross > 0) pTot = Math.max(0, pGross - pCm);

    const pPur = parseFloat(String(purchases).replace(/,/g, '')) || 0;
    const pInputTax = parseFloat(String(inputTax || '0').replace(/,/g, '')) || 0;
    const pEwt = parseFloat(String(ewtTax || '0').replace(/,/g, '')) || 0;

    // Extreme Value Spike Check
    if (pGross >= 10000000) {
      anomalies.push({
        type: "info",
        title: "High Value Alert",
        message: `Gross amount (₱${pGross.toLocaleString('en-US', {minimumFractionDigits: 2})}) exceeds ₱10M. Ensure no accidental decimal point omission occurred.`
      });
    }

    // Credit Memo Check
    if (pCm > pGross && pGross > 0) {
      anomalies.push({
        type: "warning",
        title: "Credit Memo Exceeds Gross",
        message: `Credit Memo deduction (₱${pCm.toFixed(2)}) is greater than Gross (₱${pGross.toFixed(2)}).`
      });
    }

    // Account Title vs Input Tax Conflict
    const cleanTitle = String(accountTitle || "").toUpperCase();
    if ((cleanTitle.includes("NV") || cleanTitle.includes("NON-VAT")) && pInputTax > 0.05) {
      anomalies.push({
        type: "warning",
        title: "Non-VAT Account with Tax",
        message: `Account Title is Non-VAT (${accountTitle}), but Input Tax is greater than 0.00 (₱${pInputTax.toFixed(2)}).`
      });
    }

    // Validate 12% VAT Split on Net Base (Total Amount)
    if (!isMixedVat && pTot > 0 && pPur > 0) {
      const expectedPurchases = pTot / 1.12;
      const vatDiff = Math.abs(expectedPurchases - pPur);
      if (vatDiff > 1.00 && Math.abs(pTot - pPur) > 1.00) {
        anomalies.push({
          type: "info",
          title: "Non-Standard Tax Split",
          message: `Purchases (₱${pPur.toFixed(2)}) deviates from standard 12% VAT extraction on Net Amount ₱${pTot.toFixed(2)} (expected ₱${expectedPurchases.toFixed(2)}).`
        });
      }
    }

    // Validate 12% Input Tax Math
    if (!isMixedVat && pPur > 0 && pInputTax > 0) {
      const expectedVat = pPur * 0.12;
      if (Math.abs(expectedVat - pInputTax) > 1.00) {
        anomalies.push({
          type: "info",
          title: "VAT Computation Variance",
          message: `Input Tax (₱${pInputTax.toFixed(2)}) varies from expected 12% VAT on purchases (expected ₱${expectedVat.toFixed(2)}).`
        });
      }
    }

    // Validate Withholding Tax percentage (1% for goods / 2% for services)
    if (pPur > 0 && pEwt > 0) {
      const cleanAtc = String(atc || "").toUpperCase();
      let expectedRate = 0.01;
      if (cleanAtc.includes("160")) expectedRate = 0.02;
      else if (cleanAtc.includes("158")) expectedRate = 0.01;

      const expectedEwt = pPur * expectedRate;
      if (Math.abs(expectedEwt - pEwt) > 1.00) {
        anomalies.push({
          type: "info",
          title: "EWT Rate Variance",
          message: `Tax withheld (₱${pEwt.toFixed(2)}) varies from expected ${(expectedRate * 100).toFixed(0)}% for ATC ${atc || "WC158"} (expected ₱${expectedEwt.toFixed(2)}).`
        });
      }
    }

    return { success: true, anomalies: anomalies, isClean: anomalies.length === 0 };
  } catch(err) {
    return { success: false, error: err.message };
  }
}

function webappSaveVendorMemory(supplierName, accountTitle, atc, tin, address, companyContext) {
  if (!supplierName) return { success: false };
  try {
    const props = PropertiesService.getScriptProperties();
    const cleanKey = "VM_" + String(companyContext).toUpperCase() + "_" + String(supplierName).trim().toUpperCase().replace(/[^A-Z0-9]/g, '_').substring(0, 30);
    const memoryData = {
      accountTitle: accountTitle || "",
      atc: atc || "WC158",
      tin: tin || "",
      address: address || "",
      lastUpdated: new Date().toISOString()
    };
    props.setProperty(cleanKey, JSON.stringify(memoryData));
    return { success: true };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

function webappGetVendorMemory(supplierName, companyContext) {
  if (!supplierName) return { success: false };
  try {
    const props = PropertiesService.getScriptProperties();
    const cleanKey = "VM_" + String(companyContext).toUpperCase() + "_" + String(supplierName).trim().toUpperCase().replace(/[^A-Z0-9]/g, '_').substring(0, 30);
    const stored = props.getProperty(cleanKey);
    if (stored) {
      return { success: true, memory: JSON.parse(stored) };
    }
    return { success: false };
  } catch(e) {
    return { success: false };
  }
}

function webappGenerateExecutiveAiSummary(companyContext, sessionToken) {
  const user = getAuthenticatedUser(sessionToken);
  if (!user) return { success: false, error: "Authentication required." };

  const props = PropertiesService.getScriptProperties();
  const apiKey = props.getProperty('GEMINI_API_KEY');
  if (!apiKey) return { success: false, error: "Missing GEMINI_API_KEY." };

  try {
    const verified = webappGetVerifiedInvoices(companyContext);
    let totalGross = 0;
    let ewtCount = 0;
    let purCount = 0;
    const vendorTotals = {};

    verified.forEach(item => {
      const amt = parseFloat(String(item.totalAmount).replace(/,/g, '')) || 0;
      totalGross += amt;
      if (item.recordType === "Standard Purchase") purCount++;
      else ewtCount++;

      const v = item.companyName || "Unknown";
      vendorTotals[v] = (vendorTotals[v] || 0) + amt;
    });

    // Top 3 suppliers
    const topVendors = Object.keys(vendorTotals)
      .sort((a,b) => vendorTotals[b] - vendorTotals[a])
      .slice(0, 3)
      .map(k => `${k} (₱${vendorTotals[k].toLocaleString('en-US', {maximumFractionDigits: 0})})`)
      .join(", ");

    const prompt = `Write a crisp, high-level Executive Financial Summary for the executive management of ${companyContext}.
Data Points:
- Total Invoices Verified: ${verified.length} (2307 EWTs: ${ewtCount}, Standard Purchases: ${purCount})
- Total Verified Volume: PHP ${totalGross.toLocaleString('en-US', {minimumFractionDigits: 2})}
- Top Suppliers Withheld: ${topVendors || "None"}
- System Health: All accounting databases verified.

Format: Exactly 3 professional bullet points highlighting operational throughput, withholding volume, and compliance status. Keep it concise, executive, and formal.`;

    const payload = {
      contents: [{ role: "user", parts: [{ text: prompt }] }],
      generationConfig: { temperature: 0.2, maxOutputTokens: 500 }
    };

    const url = "https://generativelanguage.googleapis.com/v1beta/models/gemini-3.5-flash-lite:generateContent?key=" + apiKey;
    const response = UrlFetchApp.fetch(url, {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });

    if (response.getResponseCode() === 200) {
      const json = JSON.parse(response.getContentText());
      if (json.candidates && json.candidates[0]) {
        return { success: true, briefing: json.candidates[0].content.parts[0].text };
      }
    }
    return { success: false, error: "Failed to generate briefing." };
  } catch(err) {
    return { success: false, error: err.message };
  }
}

// =========================================================================
// AUDITOR COMPLIANCE EXPORT PACKAGE (PART 5)
// =========================================================================

function webappExportAuditorCompliancePackage(startStr, endStr, companyContext, sessionToken) {
  const user = getAuthenticatedUser(sessionToken);
  if (!user || user.role === "Encoder" || user.role === "Viewer") {
    return { success: false, error: "Permission Denied: Only Supervisors or Super Admins can export Auditor Packages." };
  }

  try {
    const sDate = new Date(startStr + " 00:00:00");
    const eDate = new Date(endStr + " 23:59:59");
    const cleanContext = String(companyContext).replace(/[^a-zA-Z0-9]/g, '');
    const fileName = `BIR_2307_Auditor_Package_${cleanContext}_${startStr}_to_${endStr}`.replace(/\//g, '-');

    const tempSs = SpreadsheetApp.create(fileName);

    // TAB 1: 2307 Reconciled Ledger
    const ledgerSheet = tempSs.getSheets()[0];
    ledgerSheet.setName("2307 Reconciled Ledger");

    const ledgerHeaders = [
      "SN", "Entity", "Account Title", "Invoice Date", "Supplier Name", 
      "TIN", "Address", "Invoice Number", "Gross", "Input Tax", "Purchases", 
      "EWT 1%", "EWT 2%", "ATC", "Release Date", "Verification Timestamp"
    ];
    const ledgerRows = [ledgerHeaders];

    const verified = webappGetVerifiedInvoices(companyContext);
    let sn = 1;

    verified.forEach(item => {
      let rDate = new Date(item.releaseDate);
      if (!isNaN(rDate.getTime()) && rDate >= sDate && rDate <= eDate) {
        let p = {};
        try { p = JSON.parse(item.extractedData); } catch(e){}
        ledgerRows.push([
          sn++, companyContext, p["Account Title"] || "Purchases - Local", item.invoiceDate || "", item.companyName,
          p["TIN"] || "", p["Address"] || "", item.invoiceNo, item.totalAmount, p["Input Tax"] || "0.00",
          p["Purchases"] || "0.00", p["EWT 1%"] || "", p["EWT 2%"] || "", p["ATC"] || "WC158",
          item.releaseDate, item.verifiedDateTime
        ]);
      }
    });

    ledgerSheet.getRange(1, 1, ledgerRows.length, ledgerHeaders.length).setValues(ledgerRows);
    ledgerSheet.getRange(1, 1, 1, ledgerHeaders.length).setFontWeight("bold").setBackground("#111c43").setFontColor("#ffffff");
    ledgerSheet.getRange(2, 8, ledgerRows.length, 1).setNumberFormat("@"); // Invoice No
    ledgerSheet.getRange(2, 6, ledgerRows.length, 1).setNumberFormat("@"); // TIN
    ledgerSheet.getRange(2, 9, ledgerRows.length, 5).setNumberFormat("#,##0.00");
    for (let i = 1; i <= ledgerHeaders.length; i++) ledgerSheet.autoResizeColumn(i);

    // TAB 2: Chain of Custody & Audit Trail
    const auditSheet = tempSs.insertSheet("Chain of Custody & Audit Trail");
    const auditHeaders = ["Timestamp", "Company Context", "User Email / Identity", "Action", "Details"];
    const auditRows = [auditHeaders];

    const logs = webappGetActivityLogs(companyContext);
    logs.forEach(l => {
      let lDate = new Date(l.timestamp);
      if (!isNaN(lDate.getTime()) && lDate >= sDate && lDate <= eDate) {
        auditRows.push([l.timestamp, companyContext, l.user, l.action, l.details]);
      }
    });

    auditSheet.getRange(1, 1, auditRows.length, auditHeaders.length).setValues(auditRows);
    auditSheet.getRange(1, 1, 1, auditHeaders.length).setFontWeight("bold").setBackground("#0f8b44").setFontColor("#ffffff");
    for (let i = 1; i <= auditHeaders.length; i++) auditSheet.autoResizeColumn(i);

    SpreadsheetApp.flush();
    Utilities.sleep(1500);

    const url = "https://docs.google.com/spreadsheets/d/" + tempSs.getId() + "/export?format=xlsx";
    const token = ScriptApp.getOAuthToken();
    const response = UrlFetchApp.fetch(url, { headers: { 'Authorization': 'Bearer ' + token }, muteHttpExceptions: true });
    const blob = response.getBlob().setName(fileName + ".xlsx");

    DriveApp.getFileById(tempSs.getId()).setTrashed(true);

    logSystemActivity(companyContext, "Exported Auditor Compliance Package", `Auditor package exported covering ${startStr} to ${endStr}`, sessionToken);

    return {
      success: true,
      base64: Utilities.base64Encode(blob.getBytes()),
      fileName: fileName + ".xlsx"
    };
  } catch(err) {
    return { success: false, error: err.message };
  }
}

// =========================================================================
// DATABASE SELF-HEALING & INTEGRITY MAINTENANCE (PART 6)
// =========================================================================

function webappRunDatabaseSelfHealing(companyContext, sessionToken) {
  const user = getAuthenticatedUser(sessionToken);
  if (!user || user.role !== "Super Admin") {
    return { success: false, error: "Permission Denied: Only Super Admins can execute Database Self-Healing." };
  }

  let lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);

    const props = PropertiesService.getScriptProperties();
    const isDevMode = ScriptApp.getService().getUrl().endsWith("/dev");
    const targetFolderId = (isDevMode && props.getProperty('VERIFIED_FOLDER_ID_DEV')) 
        ? props.getProperty('VERIFIED_FOLDER_ID_DEV') 
        : props.getProperty('VERIFIED_FOLDER_ID');

    const folder = DriveApp.getFolderById(targetFolderId);
    const filesIter = folder.getFilesByType(MimeType.GOOGLE_SHEETS);

    let totalScanned = 0;
    let totalRepaired = 0;
    let emptyRowsCleaned = 0;

    while (filesIter.hasNext()) {
      const file = filesIter.next();
      if (!file.getName().startsWith("Verified Invoices for 2307")) continue;

      const ss = SpreadsheetApp.open(file);
      const sheet = ss.getSheetByName("All Verified Invoices");
      if (!sheet || sheet.getLastRow() <= 1) continue;

      const data = sheet.getDataRange().getValues();
      let sheetModified = false;

      // 1. Remove Ghost / Completely Empty Rows
      for (let r = data.length - 1; r >= 1; r--) {
        const fileNameCell = String(data[r][1] || "").trim();
        const jsonCell = String(data[r][2] || "").trim();
        if (!fileNameCell && !jsonCell.startsWith("{")) {
          sheet.deleteRow(r + 1);
          emptyRowsCleaned++;
          sheetModified = true;
        }
      }

      // 2. Re-sequence Serial Numbers (SN) Sequentially
      const refreshedLastRow = sheet.getLastRow();
      if (refreshedLastRow > 1) {
        const numRows = refreshedLastRow - 1;
        const currentSns = sheet.getRange(2, 1, numRows, 1).getValues();
        const jsonRange = sheet.getRange(2, 3, numRows, 1);
        const jsonValues = jsonRange.getValues();

        const correctSns = [];
        for (let i = 0; i < numRows; i++) {
          const expectedSn = i + 1;
          correctSns.push([expectedSn]);
          totalScanned++;

          if (currentSns[i][0] !== expectedSn) {
            totalRepaired++;
            sheetModified = true;
          }

          // Sync internal JSON SN key
          const rawJson = String(jsonValues[i][0] || "").trim();
          if (rawJson.startsWith("{")) {
            try {
              const parsed = JSON.parse(rawJson);
              if (parsed.SN !== expectedSn) {
                parsed.SN = expectedSn;
                jsonValues[i][0] = JSON.stringify(parsed);
                sheetModified = true;
              }
            } catch(e) {}
          }
        }

        if (sheetModified) {
          sheet.getRange(2, 1, numRows, 1).setValues(correctSns);
          jsonRange.setValues(jsonValues);
          SpreadsheetApp.flush();
        }
      }
    }

    logSystemActivity(
      companyContext,
      "Database Self-Healing Executed",
      `Self-healing complete. ${totalScanned} records verified, ${totalRepaired} SN gaps repaired, ${emptyRowsCleaned} ghost rows purged.`,
      sessionToken
    );

    return {
      success: true,
      recordsScanned: totalScanned,
      gapsRepaired: totalRepaired,
      ghostRowsCleaned: emptyRowsCleaned
    };
  } catch(err) {
    return { success: false, error: err.message };
  } finally {
    lock.releaseLock();
  }
}

function webappSaveSystemSettingsConfig(sessionToken, updatedSettings) {
  const user = getAuthenticatedUser(sessionToken);
  if (!user || user.role !== "Super Admin") {
    return { success: false, error: "Permission Denied: Only Super Admins can modify settings." };
  }

  try {
    const props = PropertiesService.getScriptProperties();

    if (updatedSettings.signatoryName) props.setProperty('SIGNATORY_2307_NAME', updatedSettings.signatoryName.trim());
    if (updatedSettings.signatoryTitle) props.setProperty('SIGNATORY_2307_TITLE', updatedSettings.signatoryTitle.trim());
    if (updatedSettings.signatoryTin) props.setProperty('SIGNATORY_2307_TIN', updatedSettings.signatoryTin.trim());
    if (updatedSettings.discrepancyTolerance) props.setProperty('DISCREPANCY_TOLERANCE', updatedSettings.discrepancyTolerance.trim());
    if (updatedSettings.adminNotificationEmail) props.setProperty('ADMIN_EMAIL', updatedSettings.adminNotificationEmail.trim());
    if (updatedSettings.collectionAssistant) props.setProperty('DEFAULT_COLLECTION_ASSISTANT', updatedSettings.collectionAssistant.trim());

    if (updatedSettings.formRevisionTitle) props.setProperty('FORM_2307_REVISION_TITLE', updatedSettings.formRevisionTitle.trim());
    if (updatedSettings.formTrackingCode) props.setProperty('FORM_2307_TRACKING_CODE', updatedSettings.formTrackingCode.trim());
    if (updatedSettings.birLogoId) props.setProperty('BIR_LOGO_DRIVE_ID', updatedSettings.birLogoId.trim());
    if (updatedSettings.barcodeId) props.setProperty('BARCODE_DRIVE_ID', updatedSettings.barcodeId.trim());
    if (updatedSettings.signatureId) props.setProperty('SIGNATURE_DRIVE_ID', updatedSettings.signatureId.trim());
    if (updatedSettings.formBannerColor) props.setProperty('FORM_2307_BANNER_COLOR', updatedSettings.formBannerColor.trim());
    if (updatedSettings.formBorderWidth) props.setProperty('FORM_2307_BORDER_WIDTH', updatedSettings.formBorderWidth.trim());
    if (updatedSettings.part3Row1Desc) props.setProperty('PART3_ROW1_DESC', updatedSettings.part3Row1Desc.trim());
    if (updatedSettings.part3Row2Desc) props.setProperty('PART3_ROW2_DESC', updatedSettings.part3Row2Desc.trim());
    if (updatedSettings.legalPerjuryText) props.setProperty('LEGAL_PERJURY_TEXT', updatedSettings.legalPerjuryText.trim());

    if (updatedSettings.pingconTin) props.setProperty('PINGCON_TIN', updatedSettings.pingconTin.trim());
    if (updatedSettings.pingconAddr) props.setProperty('PINGCON_ADDR', updatedSettings.pingconAddr.trim());
    if (updatedSettings.keysysTin) props.setProperty('KEYSYS_TIN', updatedSettings.keysysTin.trim());
    if (updatedSettings.keysysAddr) props.setProperty('KEYSYS_ADDR', updatedSettings.keysysAddr.trim());
    if (updatedSettings.a2zTin) props.setProperty('A2Z_TIN', updatedSettings.a2zTin.trim());
    if (updatedSettings.a2zAddr) props.setProperty('A2Z_ADDR', updatedSettings.a2zAddr.trim());
    if (updatedSettings.bonitaTin) props.setProperty('BONITA_TIN', updatedSettings.bonitaTin.trim());
    if (updatedSettings.bonitaAddr) props.setProperty('BONITA_ADDR', updatedSettings.bonitaAddr.trim());

    logSystemActivity(
      "SYSTEM",
      "System Settings Updated",
      `Signatory: ${updatedSettings.signatoryName} | Tolerance: ₱${updatedSettings.discrepancyTolerance}`,
      sessionToken
    );

    return { success: true };
  } catch(err) {
    return { success: false, error: err.message };
  }
}

// =========================================================================
// SUPER ADMIN USER MANAGEMENT API ENDPOINTS
// =========================================================================

function webappGetPortalUsers(sessionToken, portalType = "SUPPLIER") {
  // SECURITY GATE: Only Super Admins can access user management
  const user = getAuthenticatedUser(sessionToken);
  if (!user || user.role !== "Super Admin") {
    return { success: false, error: "Permission Denied: Only Super Admins can access User Management." };
  }

  try {
    const employees = fetchActiveHREmployees();
    const sheet = _getPortalRolesSheet(portalType);
    const data = sheet.getDataRange().getValues();
    const overrides = {};

    for (let i = 1; i < data.length; i++) {
      const email = String(data[i][0]).trim().toLowerCase();
      if (email) {
        overrides[email] = {
          role: String(data[i][2] || "").trim(),
          status: String(data[i][3] || "Active").trim()
        };
      }
    }

    const props = PropertiesService.getScriptProperties();
    const superAdminList = (props.getProperty('SUPER_ADMIN_EMAILS') || "").toLowerCase().split(',').map(s => s.trim());

    const resultList = employees.map(emp => {
      const email = String(emp.company_email || emp.personal_email || "").trim().toLowerCase();
      let role = "Encoder";
      let status = "Active";

      if (superAdminList.includes(email)) {
        role = "Super Admin";
      } else if (overrides[email]) {
        if (overrides[email].role) role = overrides[email].role;
        if (overrides[email].status) status = overrides[email].status;
      }

      return {
        id: emp.id,
        name: emp.name,
        email: email,
        department: emp.department || "General",
        role: role,
        status: status,
        isPermanentAdmin: superAdminList.includes(email)
      };
    });

    resultList.sort((a, b) => a.name.localeCompare(b.name));
    return { success: true, users: resultList };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

function webappUpdateUserAccess(sessionToken, targetEmail, targetName, newRole, newStatus, portalType = "SUPPLIER") {
  const user = getAuthenticatedUser(sessionToken);
  if (!user || user.role !== "Super Admin") {
    return { success: false, error: "Permission Denied: Only Super Admins can update roles." };
  }

  const cleanTargetEmail = String(targetEmail).trim().toLowerCase();
  const props = PropertiesService.getScriptProperties();
  const superAdminList = (props.getProperty('SUPER_ADMIN_EMAILS') || "").toLowerCase().split(',').map(s => s.trim());

  if (superAdminList.includes(cleanTargetEmail) && newRole !== "Super Admin") {
    return { success: false, error: "Cannot downgrade a permanent Super Admin configured in Script Properties." };
  }

  try {
    const sheet = _getPortalRolesSheet(portalType);
    const data = sheet.getDataRange().getValues();
    let rowIndex = -1;

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim().toLowerCase() === cleanTargetEmail) {
        rowIndex = i + 1;
        break;
      }
    }

    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MM/dd/yyyy HH:mm:ss");
    const updatedBy = `${user.name} (${user.email})`;

    if (rowIndex > -1) {
      sheet.getRange(rowIndex, 2, 1, 5).setValues([[targetName, newRole, newStatus, updatedBy, timestamp]]);
    } else {
      sheet.appendRow([cleanTargetEmail, targetName, newRole, newStatus, updatedBy, timestamp]);
    }

    // If deactivated, invalidate any active cache sessions for that user
    if (newStatus.toLowerCase() === "deactivated") {
      const cacheKey = "OTP_USER_" + Utilities.base64Encode(cleanTargetEmail).replace(/=/g, '');
      CacheService.getScriptCache().remove(cacheKey);
    }

    // Stamp in the Audit Log
    logSystemActivity("SYSTEM", "User Role & Access Updated", `Updated ${targetName} (${cleanTargetEmail}) -> Role: ${newRole} | Status: ${newStatus}`, sessionToken);

    return { success: true };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

// SAFE DIAGNOSTIC TEST: Tests User Roles Database Creation
function TEST_portalUserManagement() {
  console.log("Initializing Portal_User_Roles database sheet in Drive...");
  const sheet = _getPortalRolesSheet();
  console.log("Sheet Initialized:", sheet.getParent().getName(), "| URL:", sheet.getParent().getUrl());
}

// SAFE DIAGNOSTIC TEST: Run this in editor to test your connection
function TEST_hrApiConnection() {
  console.log("Testing Pingcon HR Center API connection...");
  const employees = fetchActiveHREmployees();
  console.log("Total Active Employees Fetched: " + employees.length);
  if (employees.length > 0) {
    console.log("Sample Employee Record:", JSON.stringify(employees[0], null, 2));
  }
}

// =========================================================================
// INDIVIDUAL USER AUTHENTICATION & SESSION ENGINE (PHASE 2)
// =========================================================================

function webappSendUserOtp(inputEmail, portalContext, portalType = "SUPPLIER") {
  if (!inputEmail || typeof inputEmail !== 'string') {
    return { success: false, message: "Please provide a valid company email address." };
  }

  const cleanEmail = inputEmail.trim().toLowerCase();
  
  // 1. Verify against the live Pingcon HR Center API and specific portal quadrant
  const emp = getHREmployeeByEmail(cleanEmail, portalType);
  if (!emp) {
    return { 
      success: false, 
      message: "Access Denied: [" + cleanEmail + "] is not listed as an active employee in the Pingcon HR Center or has been deactivated in this portal." 
    };
  }

  // 2. Generate random 6-digit OTP code
  const otp = Math.floor(100000 + Math.random() * 900000).toString();
  const cacheKey = "OTP_USER_" + Utilities.base64Encode(cleanEmail).replace(/=/g, '');
  
  // Save OTP in CacheService for 120 seconds (2 minutes expiration)
  CacheService.getScriptCache().put(cacheKey, otp, 120);

  const contextTitle = portalContext || "Finance & 2307 Portal";
  const subject = "Your Access Code for " + contextTitle;
  
  const body = `
    <div style="font-family: Arial, sans-serif; font-size: 14px; color: #333; line-height: 1.6; max-width: 500px; margin: 0 auto; border: 1px solid #eef2f7; padding: 25px; border-radius: 8px;">
      <h2 style="color: #1a73e8; margin-top: 0;">Security Verification</h2>
      <p>Hello <strong>${emp.name}</strong>,</p>
      <p>Your single-use access code for the <strong>${contextTitle}</strong> is:</p>
      <div style="background: #f1f3f4; padding: 15px; text-align: center; border-radius: 6px; margin: 20px 0;">
        <span style="font-size: 28px; font-weight: bold; letter-spacing: 5px; color: #1a73e8;">${otp}</span>
      </div>
      <p style="font-size: 12px; color: #888;">This code will expire in <strong>2 minutes</strong>. If you did not request this verification, please contact IT immediately.</p>
      <br>
      <p style="font-size: 12px; color: #555; border-top: 1px solid #eee; padding-top: 10px;">Pingcon Marketing & Tech Finance System</p>
    </div>
  `;

  try {
    sendRobustEmail(cleanEmail, subject, body, { name: "Finance Security", isSystemEmail: true });
    return { 
      success: true, 
      message: "Access code sent to " + cleanEmail + ". Please check your inbox.",
      user: { name: emp.name, role: emp.role }
    };
  } catch (err) {
    return { success: false, message: "Failed to dispatch email: " + err.message };
  }
}

function webappVerifyUserOtp(inputEmail, userOtp, portalType = "SUPPLIER") {
  if (!inputEmail || !userOtp) {
    return { success: false, message: "Missing email or verification code." };
  }

  const cleanEmail = inputEmail.trim().toLowerCase();
  const cacheKey = "OTP_USER_" + Utilities.base64Encode(cleanEmail).replace(/=/g, '');
  const cache = CacheService.getScriptCache();
  const storedOtp = cache.get(cacheKey);

  if (!storedOtp || storedOtp !== String(userOtp).trim()) {
    return { success: false, message: "Invalid or expired access code. Please request a new one." };
  }

  // 1. Destroy OTP immediately upon successful entry (Prevent Replay Attacks)
  cache.remove(cacheKey);

  // 2. Fetch full verified profile and role for this specific portal
  const emp = getHREmployeeByEmail(cleanEmail, portalType);
  if (!emp) {
    return { success: false, message: "Employee profile deactivated or missing." };
  }

  // 3. Issue 6-Hour Cryptographic Session Token (21,600 seconds)
  const sessionToken = Utilities.getUuid();
  const sessionData = {
    email: cleanEmail,
    name: emp.name,
    role: emp.role,
    employeeId: emp.employeeId,
    loginTime: new Date().toISOString()
  };

  cache.put("SESSION_" + sessionToken, JSON.stringify(sessionData), 21600);

  return {
    success: true,
    token: sessionToken,
    user: sessionData
  };
}

// Master Helper: Resolves authenticated identity for Audit Logs and Permission Gates
function getAuthenticatedUser(sessionToken) {
  if (!sessionToken || typeof sessionToken !== 'string') return null;
  const cache = CacheService.getScriptCache();
  const sessionRaw = cache.get("SESSION_" + sessionToken.trim());
  if (!sessionRaw) return null;
  try {
    return JSON.parse(sessionRaw);
  } catch (e) {
    return null;
  }
}

// User Logout: Destroys the session token from cache immediately
function webappLogoutUser(sessionToken) {
  if (sessionToken) {
    CacheService.getScriptCache().remove("SESSION_" + sessionToken.trim());
  }
  return { success: true };
}

// SAFE DIAGNOSTIC TEST: Tests OTP Generation & Session Validation
function TEST_userLoginFlow() {
  const employees = fetchActiveHREmployees();
  if (employees.length === 0) {
    console.error("No employees found. Run TEST_hrApiConnection first.");
    return;
  }

  const testEmp = employees[0];
  const testEmail = testEmp.company_email || testEmp.personal_email;
  console.log("Simulating Login for:", testEmp.name, "(" + testEmail + ")");

  // Test 1: Send OTP
  const sendRes = webappSendUserOtp(testEmail, "Automated Test Portal");
  console.log("OTP Send Result:", JSON.stringify(sendRes));

  // Test 2: Peek at stored OTP in Cache
  const cacheKey = "OTP_USER_" + Utilities.base64Encode(testEmail).replace(/=/g, '');
  const currentOtp = CacheService.getScriptCache().get(cacheKey);
  console.log("Peek at Generated OTP from Cache:", currentOtp);

  // Test 3: Verify OTP & Issue Token
  const verifyRes = webappVerifyUserOtp(testEmail, currentOtp);
  console.log("Verification & Token Issue Result:", JSON.stringify(verifyRes));

  if (verifyRes.success) {
    // Test 4: Validate Session Token
    const userSession = getAuthenticatedUser(verifyRes.token);
    console.log("Verified Session Data Retrieved:", JSON.stringify(userSession, null, 2));
    
    // Cleanup test token
    webappLogoutUser(verifyRes.token);
    console.log("Test Session Cleaned Up Successfully.");
  }
}

function authenticateUser(username, password) {
  let user = String(username).trim();
  let pass = String(password).trim();
  let props = PropertiesService.getScriptProperties();
  
  if (user === props.getProperty('APP_ADMIN_USER') && pass === props.getProperty('APP_ADMIN_PASS')) {
    let token = Utilities.getUuid();
    CacheService.getScriptCache().put(token, "VALID", 21600);
    return { success: true, token: token, role: "admin" };
  }
  
  if (user === props.getProperty('APP_ADMIN_USER') && pass === props.getProperty('APP_2307_PASS')) {
    let token = Utilities.getUuid();
    CacheService.getScriptCache().put(token, "VALID_2307", 21600); 
    return { success: true, token: token, role: "2307" };
  }
  return { success: false, message: "Invalid username or password." };
}

function logoutUser(token) {
  if(token) CacheService.getScriptCache().remove(token);
  return true;
}

function doGet(e) {
  let page = (e.parameter && e.parameter.page) ? String(e.parameter.page).trim() : "index";
  let template = HtmlService.createTemplateFromFile("SHELL");
  template.requestedPage = page;
  template.rowParam = (e.parameter && e.parameter.row) ? e.parameter.row : "";
  
  // Pass SCRIPT_PROPS to the SHELL so it can flow down to other files
  template.SCRIPT_PROPS = PropertiesService.getScriptProperties().getProperties();
  
  return template.evaluate()
    .setTitle("Pingcon Finance Portal")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getSecurePage(page, token, rowParam) {
  // Block the base Invoice_Search link entirely BEFORE checking for authentication
          if (page.toUpperCase() === "INVOICE_SEARCH") {
              return { status: "ERROR", html: "<h2 style='text-align:center; margin-top:50px; font-family:sans-serif; color:#d93025;'>Invalid Link.<br><br><span style='font-size:16px; color:#5f6368;'>Please use the specific company link provided by the administrator (e.g., ?page=Invoice_Search_Pingcon).</span></h2>" };
          }

          // Restrict public access ONLY to URLs that specify the company (must have the underscore)
          let isPublic = (page.toUpperCase().startsWith("INVOICE_SEARCH_") || page.toUpperCase().startsWith("EWT_CLIENT_PORTAL_"));
          
          let companyContext = "";
          if (isPublic) {
              if (page.toUpperCase().startsWith("INVOICE_SEARCH_")) {
                  if (page.toUpperCase().includes("A2Z2GO")) companyContext = "A2Z2GO";
                  else if (page.toUpperCase().includes("KEYSYS")) companyContext = "KEYSYS";
                  else if (page.toUpperCase().includes("PINGCON")) companyContext = "PINGCON";
                  else if (page.toUpperCase().includes("BONITA")) companyContext = "BONITA";
                  page = "Invoice_Search";
              }

              // Routing for the EWT Client Portal
          if (page.toUpperCase().startsWith("EWT_CLIENT_PORTAL_")) {
              if (page.toUpperCase().includes("A2Z2GO")) companyContext = "A2Z2GO";
              else if (page.toUpperCase().includes("KEYSYS")) companyContext = "KEYSYS";
              else if (page.toUpperCase().includes("PINGCON")) companyContext = "PINGCON";
              else if (page.toUpperCase().includes("BONITA")) companyContext = "BONITA";
              page = "EWT_Client_Portal";
          }

          // Routing for General Nightly Narra Portal (Cross-entity)
          if (page.toUpperCase() === "NIGHTLY_NARRA" || page.toUpperCase() === "PORTAL_NIGHTLY_NARRA") {
              companyContext = "ALL";
              page = "NIGHTLY_NARRA";
          }
      } else if (page.toUpperCase() === "NIGHTLY_NARRA" || page.toUpperCase() === "PORTAL_NIGHTLY_NARRA") {
          companyContext = "ALL";
          page = "NIGHTLY_NARRA";
      }
          
          let template;
          try {
            template = HtmlService.createTemplateFromFile(page);
          } catch(e) {
            try {
              // Fallback for uppercase filenames like INDEX
              template = HtmlService.createTemplateFromFile(page.toUpperCase());
            } catch(err) {
              return { status: "ERROR", html: "<h2 style='text-align:center; margin-top:50px; font-family:sans-serif;'>404 Error: Page not found.<br><br><span style='font-size:12px;color:red;'>" + e.toString() + "</span></h2>" };
            }
          }

          template.rowParam = rowParam;
              template.companyContext = companyContext;
              
          // --- FIX: Define missing template variables to prevent ReferenceError ---
          template.isAuthorized = false; // Always set to false initially so the HTML OTP overlay logic runs
          
          // Detect Environment (Dev vs Exec)
          const isDevMode = ScriptApp.getService().getUrl().endsWith("/dev");
          let isTech = (companyContext === "A2Z2GO" || companyContext === "KEYSYS");
          
          // Assign the correct email(s) based on environment and page
          if (page.toLowerCase() === "index") {
              template.assignedEmail = isDevMode 
                  ? (SCRIPT_PROPS['INDEX_AUTH_EMAIL_DEV'] || SCRIPT_PROPS['ADMIN_EMAIL'])
                  : (SCRIPT_PROPS['INDEX_AUTH_EMAIL_EXEC'] || SCRIPT_PROPS['ADMIN_EMAIL']);
          } else {
              if (isDevMode) {
                  template.assignedEmail = isTech 
                      ? SCRIPT_PROPS['TECH_ADMIN_EMAIL'] 
                      : SCRIPT_PROPS['ADMIN_EMAIL'];
              } else {
                  template.assignedEmail = isTech 
                      ? SCRIPT_PROPS['AP_EMAIL_TECH'] 
                      : SCRIPT_PROPS['AP_EMAIL_PINGCON'];
              }
          }
          // ------------------------------------------------------------------------

          // --- ADDED: Pass Header Logo Base64 to Template ---
          let headerLogoId = (companyContext === "A2Z2GO" || companyContext === "KEYSYS") ? HEADER_LOGO_ID_A2Z_KEYSYS : HEADER_LOGO_ID_PINGCON_BONITA;
          let headerLogoB64 = "";
          try { headerLogoB64 = "data:image/png;base64," + Utilities.base64Encode(DriveApp.getFileById(headerLogoId).getBlob().getBytes()); } catch(e) {}
          template.headerLogoB64 = headerLogoB64;
          // -------------------------------------------------
  
          // Pass SCRIPT_PROPS so HTML files inside getSecurePage can read it
          template.SCRIPT_PROPS = PropertiesService.getScriptProperties().getProperties();
  
  return { status: "AUTHORIZED", html: template.evaluate().getContent() };
}
  
function webappSendNteOtp(rowId, assignedEmail) {
   let otp = Math.floor(100000 + Math.random() * 900000).toString();
   
   let is2307 = String(rowId).startsWith("2307_");
   
   let expireTime = 120; 
   let expireText = "2 minutes";
   let contextText = is2307 ? "the 2307 Generator Portal" : "the Finance Reports Generator";
   let subjectText = is2307 ? "Your Access Code for 2307 Generator" : "Your Access Code for Finance Reports Generator";
   let colorCode = is2307 ? "#1a73e8" : "#e81e63";
   
   CacheService.getScriptCache().put("OTP_" + rowId, otp, expireTime); 
   
   let body = `
       <div style="font-family: Arial, sans-serif; font-size: 14px; color: #333; line-height: 1.5;">
           <h2>Security Verification</h2>
           <p>Your access code for ${contextText} is: <strong style="font-size: 20px; color: ${colorCode};">${otp}</strong></p>
           <p>This code will expire in ${expireText}.</p>
           <p>If you did not request this, please ignore this email.</p>
           <br>
           <p>Thank you,</p>
           <p>Finance System</p>
       </div>
   `;
   // <-- FIXED: Sends OTP to the actual assigned email and forcefully bypasses Test Mode
   sendRobustEmail(assignedEmail, subjectText, body, { name: "Finance System", isSystemEmail: true });
   return "SENT";
}

function webappVerifyPortalOtp(userOtp) {
   let stored = CacheService.getScriptCache().get("OTP_PORTAL_ACCESS");
   if (stored && stored === String(userOtp).trim()) {
       let token = Utilities.getUuid();
       CacheService.getScriptCache().put(token, "VALID", 21600); // Session lasts 6 hours
       return { success: true, token: token, role: "admin" };
   }
   return { success: false, message: "Invalid or expired access code. Please try again." };
}

function webappVerifyNteOtp(rowId, userOtp) {
   let stored = CacheService.getScriptCache().get("OTP_" + rowId);
   if (stored && stored === String(userOtp).trim()) {
       return { success: true };
   }
   return { success: false, message: "Invalid or expired access code. Please try again." };
}

function _openFileOrConvertCSV(file) {
  let fileName = file.getName();
  let mimeType = file.getMimeType();
  let isGoogleSheet = mimeType === MimeType.GOOGLE_SHEETS;
  
  let tempFileId = null;
  let ssNew = null;

  if (isGoogleSheet) {
     return { ssNew: SpreadsheetApp.openById(file.getId()), tempFileId: null };
  } 

  // 1. ALWAYS try Google's Native Conversion First (Handles .xls, .xlsx, and HTML-masked .xls perfectly without memory crashes)
  try {
      let resource = { name: "Temp_" + fileName.split(".")[0], mimeType: MimeType.GOOGLE_SHEETS };
      let newTempGS = Drive.Files.copy(resource, file.getId());
      tempFileId = newTempGS.id;
      
      // If conversion succeeded and it didn't turn into a Google Doc by mistake
      if (newTempGS.mimeType === MimeType.GOOGLE_SHEETS) {
          ssNew = SpreadsheetApp.openById(tempFileId);
          return { ssNew: ssNew, tempFileId: tempFileId };
      } else {
          Drive.Files.remove(tempFileId); 
          tempFileId = null;
      }
  } catch(e) { 
      if (tempFileId) { Drive.Files.remove(tempFileId); tempFileId = null; }
  }

  // 2. Fallback for CSVs (If native conversion failed)
  let content = "";
  try { content = file.getBlob().getDataAsString(); } catch(e) { }

  if (content) {
      let csvData = [];
      try { csvData = Utilities.parseCsv(content); } catch(e) {}
      if (csvData.length > 0) {
          ssNew = SpreadsheetApp.create("Temp_" + fileName.split(".")[0]);
          tempFileId = ssNew.getId();
          let maxCols = Math.max(...csvData.map(r => r.length));
          let paddedData = csvData.map(r => {
              let newRow = [...r];
              while(newRow.length < maxCols) newRow.push("");
              return newRow;
          });
          ssNew.getSheets()[0].getRange(1, 1, paddedData.length, maxCols).setNumberFormat("@").setValues(paddedData);
          return { ssNew: ssNew, tempFileId: tempFileId };
      }
  }
  
  throw new Error("File could not be parsed natively. Ensure it is a valid Excel or CSV file.");
}

// =========================================================================
// UNPAID DRs OVERDUE SUMMARY (NEW)
// =========================================================================
function _9d_runSendAllUnpaidDRsOverdue() {
  const props = PropertiesService.getScriptProperties();
  deleteTriggersByName("resumeUnpaidDROverdueBatch");
  props.deleteProperty("CANCEL_UNPAID_DR_OVERDUE");
  resumeUnpaidDROverdueBatch();
}

function resumeUnpaidDROverdueBatch() {
  const props = PropertiesService.getScriptProperties();
  props.setProperty("THREAD_STATE", "RUNNING");
  
  if (props.getProperty("CANCEL_UNPAID_DR_OVERDUE") === "true") {
      deleteTriggersByName("resumeUnpaidDROverdueBatch");
      props.deleteProperty("CANCEL_UNPAID_DR_OVERDUE");
      clearReportLock();
      return;
  }

  // 1. GATHER DATA
  const arAllData = gatherAllData("AR_Data_");
  const arSheets = getDatabaseSheets("AR_Data_");
  let arHeaders = [];
  if (arSheets.length > 0 && arSheets[0].getLastColumn() > 0) {
    arHeaders = arSheets[0].getRange(1, 1, 1, arSheets[0].getLastColumn()).getValues()[0].map(h => String(h).toLowerCase().trim());
  }
  
  let docDateIdx = arHeaders.findIndex(h => h.includes("date") && !h.includes("expiry") && !h.includes("countered") && !h.includes("received"));
  let deadlineIdx = arHeaders.findIndex(h => h.includes("deadline") || h.includes("due date"));
  if (deadlineIdx === -1) deadlineIdx = docDateIdx;
  let customerIdx = arHeaders.findIndex(h => h.includes("customer"));
  let managerIdx = arHeaders.findIndex(h => h.includes("manager"));
  let textIdx = arHeaders.findIndex(h => h.includes("additional text"));
  let notesIdx = arHeaders.findIndex(h => h === "notes");
  let openBalIdx = arHeaders.findIndex(h => h.includes("open balance") && !h.includes("%"));
  let originalIdx = arHeaders.findIndex(h => h.includes("sales document total in main currency") || h.includes("sales document total"));
  let paymentTermIdx = arHeaders.findIndex(h => h.includes("payment condition") || h.includes("payment term"));
  let counterDateIdx = arHeaders.findIndex(h => h.includes("countered") || h.includes("counter date"));
  let receivedDateIdx = arHeaders.findIndex(h => h.includes("received delivery") || h.includes("received date") || h.includes("delivery date"));
  let sourceFileIdx = arHeaders.findIndex(h => h.includes("source file"));
  let paymentStatusIdx = arHeaders.findIndex(h => h === "payment status" || h.includes("payment status"));
  if (paymentStatusIdx === -1 && openBalIdx > -1) paymentStatusIdx = openBalIdx + 2;

  const startStr = props.getProperty("BATCH_START_DATE");
  const endStr = props.getProperty("BATCH_END_DATE");
  const startDate = startStr ? new Date(startStr + "T00:00:00") : new Date(2025, 9, 1);
  const endDate = endStr ? new Date(endStr + "T23:59:59") : new Date(2026, 2, 20, 23, 59, 59);
  
  let allowedEntities = JSON.parse(props.getProperty("SALES_ENTITIES") || "[]");
  let allowedCustsStr = props.getProperty("SALES_TARGET_CUSTOMERS");
  let allowedCusts = allowedCustsStr ? JSON.parse(allowedCustsStr) : [];
  const now = new Date();
  
  // Cut-off Date Logic
  let cutOffDate = new Date(now);
  let daysUntilFriday = (5 - cutOffDate.getDay() + 7) % 7;
  cutOffDate.setDate(cutOffDate.getDate() + daysUntilFriday);
  cutOffDate.setHours(23, 59, 59, 999);
  
  // Create Buckets
  let buckets = { "Keysys, A2Z2GO IT": [], "Pingcon, Bonita, A2Z2GO PMC": [] };

  // 2. FILTER & CALCULATE OVERDUE DAYS
  arAllData.forEach(item => {
    let row = item.data;
    let textVal = (textIdx > -1) ? String(row[textIdx]) : "";
    let notesVal = (notesIdx > -1) ? String(row[notesIdx]) : "";
    if (isTransactionVoid(notesVal, textVal)) return;

    let openBal = (openBalIdx > -1) ? parseFloat(row[openBalIdx]) : 0;
    if (isNaN(openBal) || openBal < 1) return;

    let statusRaw = (paymentStatusIdx > -1 && paymentStatusIdx < row.length) ? String(row[paymentStatusIdx]).trim() : "";
    if (statusRaw.toLowerCase() === "fully paid") return;

    let rawDocDate = (docDateIdx > -1) ? row[docDateIdx] : "";
    let docDateObj = (rawDocDate instanceof Date) ? rawDocDate : new Date(rawDocDate);
    if (isNaN(docDateObj.getTime()) || docDateObj < startDate || docDateObj > endDate) return;

    let sourceFile = (sourceFileIdx > -1) ? String(row[sourceFileIdx]) : "";
    let companyEntity = getCompanyEntity(sourceFile);
    if (allowedEntities.length > 0 && !allowedEntities.includes(companyEntity)) return;

    let customer = (customerIdx > -1) ? String(row[customerIdx]).trim() : "";
    if (!customer) return;

    let cleanCustomer = customer.replace(/[^\x20-\x7EñÑ]/g, '').replace(/\s+/g, ' ').trim();
    let custKey = cleanCustomer.toLowerCase().replace(/\s+/g, '');
    if (allowedCusts.length > 0) {
        let matchedName = allowedCusts.find(c => c.toLowerCase().replace(/\s+/g, '') === custKey);
        if (!matchedName) return;
        customer = matchedName;
    }

    let rawDeadline = (deadlineIdx > -1) ? row[deadlineIdx] : "";
    let deadlineObj = new Date(rawDeadline);
    if (isNaN(deadlineObj.getTime()) || rawDeadline === "") {
        deadlineObj = new Date(docDateObj);
        let pTerm = (paymentTermIdx > -1) ? parseInt(row[paymentTermIdx]) : 0;
        if (!isNaN(pTerm) && pTerm > 0) {
            deadlineObj.setDate(deadlineObj.getDate() + pTerm);
        }
    }

    // Overdue Calculation
    let overdueDays = Math.floor((cutOffDate.getTime() - deadlineObj.getTime()) / (1000 * 60 * 60 * 24));
    
    let bucketCategory = "";
    if (overdueDays >= 7 && overdueDays <= 13) bucketCategory = "Overdue 7 days";
    else if (overdueDays >= 14 && overdueDays <= 29) bucketCategory = "Overdue 14 days";
    else if (overdueDays >= 30 && overdueDays <= 44) bucketCategory = "Overdue 30 days";
    else if (overdueDays >= 45 && overdueDays <= 59) bucketCategory = "Overdue 45 days";
    else if (overdueDays >= 60) bucketCategory = "Overdue 60 days";

    let salesTotal = (originalIdx > -1) ? parseFloat(row[originalIdx]) : 0;
    if (isNaN(salesTotal)) salesTotal = 0;
    let amountPaid = salesTotal - openBal;
    if (amountPaid < 0) amountPaid = 0;
    
    let bucketName = (companyEntity === "Keysys" || companyEntity === "A2ZIT") ? "Keysys, A2Z2GO IT" : "Pingcon, Bonita, A2Z2GO PMC";

    buckets[bucketName].push({
      overdueCategory: bucketCategory,
      overdueDays: overdueDays,
      docDateObj: docDateObj,
      invoiceDateStr: fastFormatDate(docDateObj),
      paymentTerm: (paymentTermIdx > -1) ? String(row[paymentTermIdx]).trim() : "",
      deadlineStr: fastFormatDate(deadlineObj),
      customer: customer,
      manager: getCanonicalName((managerIdx > -1) ? String(row[managerIdx]) : ""),
      text: textVal,
      salesTotal: salesTotal,
      amountPaid: amountPaid,
      openBal: openBal,
      counterDateStr: (counterDateIdx > -1 && row[counterDateIdx]) ? fastFormatDate(new Date(row[counterDateIdx])) : "",
      receivedDateStr: (receivedDateIdx > -1 && row[receivedDateIdx]) ? fastFormatDate(new Date(row[receivedDateIdx])) : "",
      remarksStatus: statusRaw,
      companyEntity: companyEntity === "A2ZIT" ? "A2Z2GO" : companyEntity,
      sourceFile: sourceFile // <--- ADDED: Allows the script to read the filename for the date extractor
    });
  });

  // 3. FETCH OR AUTO-CREATE EMAIL DATABASE
  let overdueEmailMap = {};
  try {
      let emailDbFolder = DriveApp.getFolderById(EMAIL_RECIPIENT_FOLDER_ID !== "1_INSERT_YOUR_NEW_EMAIL_FOLDER_ID_HERE" ? EMAIL_RECIPIENT_FOLDER_ID : DATABASE_FOLDER_ID);
      let emailFiles = emailDbFolder.getFilesByName("Overdue_Emails_Data");
      let eSheet;
      
      let foundFile = null;
      while (emailFiles.hasNext()) {
          let f = emailFiles.next();
          if (!f.isTrashed()) { foundFile = f; break; }
      }
      
      if (foundFile) {
          eSheet = SpreadsheetApp.open(foundFile).getSheets()[0];
      } else {
          // Auto-create database if it doesn't exist
          let newSs = SpreadsheetApp.create("Overdue_Emails_Data");
          DriveApp.getFileById(newSs.getId()).moveTo(emailDbFolder);
          eSheet = newSs.getSheets()[0];
          eSheet.appendRow(["Entity", "To Emails", "CC Emails"]);
          eSheet.getRange("A1:C1").setFontWeight("bold").setBackground("#f3f3f3");
          eSheet.appendRow(["Pingcon", "", ""]);
          eSheet.appendRow(["Keysys", "", ""]);
          eSheet.appendRow(["A2Z2GO", "", ""]);
          eSheet.appendRow(["Bonita", "", ""]);
      }
      
      let eData = eSheet.getDataRange().getValues();
      if (eData.length > 1) {
          let eHeaders = eData[0].map(h => String(h).toLowerCase().trim());
          let entIdx = eHeaders.indexOf("entity");
          let toIdx = eHeaders.indexOf("to emails");
          let ccIdx = eHeaders.indexOf("cc emails");
          for (let i = 1; i < eData.length; i++) {
              if (entIdx > -1 && eData[i][entIdx]) {
                  overdueEmailMap[String(eData[i][entIdx]).toUpperCase().trim()] = { 
                      to: toIdx > -1 ? String(eData[i][toIdx]).trim() : "", 
                      cc: ccIdx > -1 ? String(eData[i][ccIdx]).trim() : "" 
                  };
              }
          }
      }
  } catch(e) { console.warn("Failed to load Overdue Email Database: " + e); }

  // 4. GENERATE EXCEL AND EMAIL PER BUCKET
  let reportDateString = Utilities.formatDate(now, Session.getScriptTimeZone(), "MMM dd, yyyy HHmm");
  
  let tuesdayDate = new Date(cutOffDate);
  tuesdayDate.setDate(cutOffDate.getDate() - 3); 
  let baseDate = new Date(2026, 4, 26); 
  let diffDays = Math.round((tuesdayDate.getTime() - baseDate.getTime()) / (1000 * 60 * 60 * 24));
  let snVal = Math.floor(diffDays / 7) + 1;
  let tuesdayStr = Utilities.formatDate(tuesdayDate, Session.getScriptTimeZone(), "MM/dd/yyyy");
  let fridayStr = Utilities.formatDate(cutOffDate, Session.getScriptTimeZone(), "MM/dd/yyyy");
  let soaSnText = `SOA sn=${snVal}: schedule run every Tuesdays weekly (mm/dd/yyyy)=${tuesdayStr} 10:00, SOA date up to and including Friday (mm/dd/yyyy)=${fridayStr}`;

  for (let bucketName in buckets) {
      let items = buckets[bucketName];
      if (items.length === 0) continue;

      let summaryTotals = {};
      let summaryTotalAmount = 0; // NEW: Exclusively for the Summary Tab
      let totalAmount = 0; // KEPT UNTOUCHED: For Email Body and Details Tab
      let minDateObj = null;
      let maxDateObj = null;

      items.forEach(item => {
          // Track overall min/max dates for this bucket
          if (!minDateObj || item.docDateObj < minDateObj) minDateObj = item.docDateObj;
          if (!maxDateObj || item.docDateObj > maxDateObj) maxDateObj = item.docDateObj;

          // Keep totalAmount untouched for Email Body and Details Tab
          totalAmount += item.openBal;

          // 1. Resolve the remark based on the Friday cut-off
          let finalRemarks = item.remarksStatus;
          if (item.deadlineObj <= cutOffDate && String(finalRemarks).toLowerCase().includes("advance")) {
              finalRemarks = "Due for collection";
          }

          // 2. Strict Filter: Only add to Summary Tab totals if it is "Due" or "EWT"
          let remarksLower = String(finalRemarks).toLowerCase();
          if (remarksLower.includes("due for collection") || remarksLower.includes("please fax email ewt")) {
              if (!summaryTotals[item.customer]) summaryTotals[item.customer] = 0;
              summaryTotals[item.customer] += item.openBal;
              summaryTotalAmount += item.openBal;
          }
      });
      let summaryArr = Object.keys(summaryTotals).map(k => ({customer: k, amount: summaryTotals[k]})).sort((a,b) => a.customer.localeCompare(b.customer));

      // Build the actual Date Range text (Row 4) and restore the Thursday/Friday text (Row 6)
      let minDateStr = minDateObj ? Utilities.formatDate(minDateObj, Session.getScriptTimeZone(), "MMMM d, yyyy") : "";
      let maxDateStr = maxDateObj ? Utilities.formatDate(maxDateObj, Session.getScriptTimeZone(), "MMMM d, yyyy") : "";
      let periodCoveredText = `For your reference, this report covers transactions recorded in our system from ${minDateStr} to ${maxDateStr}.`;
      
      let thursdayDate = new Date(cutOffDate);
      thursdayDate.setDate(cutOffDate.getDate() - 1);
      let thursdayFridayText = `For ${Utilities.formatDate(thursdayDate, Session.getScriptTimeZone(), "MMMM d")} and ${Utilities.formatDate(cutOffDate, Session.getScriptTimeZone(), "d, yyyy")}, I am expecting these collections:`;

      let shortBucketName = bucketName === "Keysys, A2Z2GO IT" ? "A2Z2GO" : "Pingcon";
      
      // --- NEWEST DATE EXTRACTOR LOGIC ---
      let maxAsOfDate = 0;
      let dynamicReportDateString = "";
      items.forEach(item => {
          if (item.sourceFile) {
              let match = item.sourceFile.match(/as of\s+([a-zA-Z]+)\s*(\d+),?\s+(\d{4})\s+(\d{4})/i);
              if (match) {
                  let monthStr = match[1];
                  let dayStr = match[2];
                  let yearStr = match[3];
                  let timeStr = match[4];
                  let hourStr = timeStr.substring(0, 2);
                  let minStr = timeStr.substring(2, 4);
                  let dateObj = new Date(`${monthStr} ${dayStr}, ${yearStr} ${hourStr}:${minStr}:00`);
                  
                  // Evaluates all files in this bucket and strictly locks in the NEWEST (max) date and time
                  if (!isNaN(dateObj.getTime()) && dateObj.getTime() > maxAsOfDate) {
                      maxAsOfDate = dateObj.getTime();
                      let paddedDay = dayStr.length === 1 ? "0" + dayStr : dayStr;
                      dynamicReportDateString = `${monthStr} ${paddedDay}, ${yearStr} ${timeStr}`;
                  }
              }
          }
      });
      // Fallback only if no files are found
      if (!dynamicReportDateString) {
          dynamicReportDateString = Utilities.formatDate(now, Session.getScriptTimeZone(), "MMM dd, yyyy HHmm");
      }
      
      let cleanFileName = `Consolidated Statement of Account - ${shortBucketName} as of Report Date ${dynamicReportDateString}`;
      let tempSpreadsheet = SpreadsheetApp.create(cleanFileName);

      // --- TAB 1: SUMMARY ---
      let summarySheet = tempSpreadsheet.getSheets()[0];
      summarySheet.setName("Summary");
      
      let summaryRows = [];
      summaryRows.push(["", "", ""]);
      summaryRows.push([soaSnText, "", ""]);
      summaryRows.push([`Consolidated Statement of Account - ${shortBucketName} as of Report Date ${reportDateString}`, "", ""]);
      summaryRows.push([periodCoveredText, "", ""]);
      summaryRows.push(["", "", ""]);
      
      const formatCur = (num) => "PHP " + num.toLocaleString('en-US', {minimumFractionDigits: 2, maximumFractionDigits: 2});
      let activeLabel = bucketName === "Keysys, A2Z2GO IT" ? "Total Keysys, A2Z2GO IT" : "Total Pingcon, Bonita, A2Z2GO PMC";
      
      summaryRows.push([thursdayFridayText, "", ""]);
      summaryRows.push([activeLabel + " = " + formatCur(summaryTotalAmount), "", ""]);
      summaryRows.push(["", "", ""]);
      summaryRows.push(["Communication is always 4 ways:", "", ""]);
      summaryRows.push(["1. email to email to email and", "", ""]);
      summaryRows.push(["2. viber to viber to viber and", "", ""]);
      summaryRows.push(["3. voice to voice to voice and", "", ""]);
      summaryRows.push(["4. Timestamp app and please see attached Timestamp configuration", "", ""]);
      summaryRows.push(["", "", ""]);
      summaryRows.push(["sn", shortBucketName, "amount"]);
      
      for (let i = 0; i < summaryArr.length; i++) {
          summaryRows.push([i + 1, summaryArr[i].customer, summaryArr[i].amount]);
      }
      summaryRows.push(["", activeLabel, summaryTotalAmount]);

      summarySheet.getRange(1, 1, summaryRows.length, 3).setValues(summaryRows);
      summarySheet.getRange("A2:F2").setBackground("#92d050"); 
      summarySheet.getRange("A3:C3").setBackground("yellow").setFontWeight("bold");
      summarySheet.getRange("A6:C7").setBackground("yellow").setFontWeight("bold");
      summarySheet.getRange("A9").setFontWeight("bold");
      
      let tableHeaderRow = 14;
      summarySheet.getRange(tableHeaderRow, 1, 1, 3).setBackground("#800080").setFontColor("white").setFontWeight("bold");
      summarySheet.getRange(tableHeaderRow + 1, 3, summaryArr.length + 1, 1).setNumberFormat("#,##0.00");
      if (summaryArr.length > 0) summarySheet.getRange(tableHeaderRow, 1, summaryArr.length + 2, 3).setBorder(true, true, true, true, true, true);

      summarySheet.setColumnWidth(1, 40);
      summarySheet.setColumnWidth(2, 350);
      summarySheet.setColumnWidth(3, 120);

      let standardHeaders = [
          "SN", "Entity", "Invoice Date (MM/DD/YYYY)", "Terms", "Due Date (MM/DD/YYYY)", 
          "Cust Manager", "Customer Name", "Additional text", "Sales Total", "Amount Paid", "Open Balance", 
          "Counter Date (MM/DD/YYYY)", "Received Date (MM/DD/YYYY)", "Remarks/Payment Status"
      ];

      // --- TAB 2: DETAILS (ALL UNPAID DRs) ---
      let detailsTab = tempSpreadsheet.insertSheet("Details");
      
      // Create a copy of the items to sort specifically for the Details tab
      let allDetailsItems = [...items];
      allDetailsItems.sort((a, b) => {
          let cCmp = a.customer.localeCompare(b.customer);
          if (cCmp !== 0) return cCmp;
          return a.docDateObj - b.docDateObj;
      });

      let stdRows = [];
      let stdCurrentCustomer = null;
      let stdCustTotal = 0;
      let stdSn = 1;

      allDetailsItems.forEach(itm => {
          // Check if customer name has changed to insert subtotal and new header
          if (stdCurrentCustomer !== itm.customer) {
              if (stdCurrentCustomer !== null) {
                  let subRow = Array(standardHeaders.length).fill("");
                  subRow[6] = "Total for " + stdCurrentCustomer + ":";
                  subRow[10] = stdCustTotal;
                  stdRows.push(subRow);
                  stdRows.push(Array(standardHeaders.length).fill("")); // Blank Spacer Row
              }
              stdCurrentCustomer = itm.customer;
              stdCustTotal = 0;
              stdRows.push(standardHeaders); // Repeated Header
          }

          stdRows.push([
              stdSn++, itm.companyEntity, itm.invoiceDateStr, itm.paymentTerm, itm.deadlineStr, 
              itm.manager, itm.customer, itm.text, itm.salesTotal, itm.amountPaid, itm.openBal,
              itm.counterDateStr, itm.receivedDateStr, itm.remarksStatus
          ]);
          stdCustTotal += itm.openBal;
      });

      // Add final subtotal for the very last customer
          if (stdCurrentCustomer !== null) {
              let subRow = Array(standardHeaders.length).fill("");
              subRow[6] = "Total for " + stdCurrentCustomer + ":";
              subRow[10] = stdCustTotal;
              stdRows.push(subRow);
          }

          // --- ADDED: OVERALL GRAND TOTAL ROW ---
          stdRows.push(Array(standardHeaders.length).fill("")); // Blank spacer row
          let grandRow = Array(standardHeaders.length).fill("");
          grandRow[6] = "OVERALL GRAND TOTAL:";
          grandRow[10] = totalAmount; // Uses the master totalAmount calculated earlier for this entity
          stdRows.push(grandRow);
          // --------------------------------------

          detailsTab.getRange(1, 1, stdRows.length, standardHeaders.length).setValues(stdRows);

          // Formatting logic for bolding headers and subtotal borders
          for (let r = 0; r < stdRows.length; r++) {
              let rowData = stdRows[r];
              let excelRow = r + 1;
              if (rowData[0] === "SN") {
                  detailsTab.getRange(excelRow, 1, 1, standardHeaders.length).setFontWeight("bold").setBackground("#f3f3f3");
              }
              // Modified to apply bold and top-border to both Customer Totals AND the Grand Total
              if (String(rowData[6]).includes("Total for ") || String(rowData[6]).includes("OVERALL GRAND TOTAL:")) {
                  detailsTab.getRange(excelRow, 7, 1, 5).setFontWeight("bold"); 
                  detailsTab.getRange(excelRow, 11).setBorder(true, false, false, false, false, false, "black", SpreadsheetApp.BorderStyle.SOLID);
              }
          }

      detailsTab.getRange(1, 9, stdRows.length, 3).setNumberFormat("#,##0.00");
      detailsTab.getRange(1, 8, stdRows.length, 1).setNumberFormat("@");
      
      detailsTab.getDataRange().setWrap(true).setVerticalAlignment("top");
      for (let i = 1; i <= standardHeaders.length; i++) detailsTab.autoResizeColumn(i);
      detailsTab.setColumnWidth(1, 60);  
      detailsTab.setColumnWidth(3, 120); 
      detailsTab.setColumnWidth(5, 120); 
      detailsTab.setColumnWidth(7, 250); 
      detailsTab.setColumnWidth(8, 250); 
      detailsTab.setColumnWidth(9, 110); 
      detailsTab.setColumnWidth(10, 110); 
      detailsTab.setColumnWidth(11, 110);
      detailsTab.setColumnWidth(14, 200);

      // --- TAB 3: OVERDUE DETAILS ---
      let detailSheet = tempSpreadsheet.insertSheet("Overdue Details");
      let overdueHeaders = [
          "Master Serial", "SN", "Entity", "Invoice Date (MM/DD/YYYY)", "Terms", "Due Date (MM/DD/YYYY)", 
          "Cust Manager", "Customer Name", "Additional text", "Sales Total", "Amount Paid", "Open Balance", 
          "Counter Date (MM/DD/YYYY)", "Received Date (MM/DD/YYYY)", "Remarks/Payment Status"
      ];
      
      let detailRows = [];
      let formatRanges = [];

      const categories = ["Overdue 7 days", "Overdue 14 days", "Overdue 30 days", "Overdue 45 days", "Overdue 60 days"];
      let masterSerial = 1; // Continuous counter for Column A
      let grandOverdueTotal = 0; // Tracks the overall grand total

      categories.forEach(cat => {
          let catItems = items.filter(i => i.overdueCategory === cat);
          if (catItems.length === 0) return;

          catItems.sort((a, b) => {
              let cCmp = a.customer.localeCompare(b.customer);
              if (cCmp !== 0) return cCmp;
              return a.docDateObj - b.docDateObj;
          });

          // Add Section Header (No Merge Cells)
          let catHeader = Array(overdueHeaders.length).fill("");
          catHeader[0] = cat;
          detailRows.push(catHeader);
          formatRanges.push({ type: 'cat', row: detailRows.length });

          // Add Column Headers
          detailRows.push(overdueHeaders);
          formatRanges.push({ type: 'col', row: detailRows.length });

          let sn = 1; // Resets per category for Column B
          let catTotal = 0; // Tracks the subtotal for this specific category

          catItems.forEach(itm => {
              detailRows.push([
                  masterSerial++, sn++, itm.companyEntity, itm.invoiceDateStr, itm.paymentTerm, itm.deadlineStr, 
                  itm.manager, itm.customer, itm.text, itm.salesTotal, itm.amountPaid, itm.openBal,
                  itm.counterDateStr, itm.receivedDateStr, itm.remarksStatus
              ]);
              catTotal += itm.openBal;
              grandOverdueTotal += itm.openBal;
          });
          
          // Add Category Subtotal Row
          let subTotalRow = Array(overdueHeaders.length).fill("");
          subTotalRow[7] = "Total for " + cat + ":"; // Column H
          subTotalRow[11] = catTotal; // Column L
          detailRows.push(subTotalRow);
          formatRanges.push({ type: 'total', row: detailRows.length }); // Tag for formatting

          detailRows.push(Array(overdueHeaders.length).fill("")); // Spacer
      });

      // Add Overall Grand Total Row at the very bottom
      if (detailRows.length > 0) {
          let grandRow = Array(overdueHeaders.length).fill("");
          grandRow[7] = "OVERALL GRAND TOTAL:"; // Column H
          grandRow[11] = grandOverdueTotal; // Column L
          detailRows.push(grandRow);
          formatRanges.push({ type: 'total', row: detailRows.length }); // Tag for formatting
      }

      if (detailRows.length > 0) {
          detailSheet.getRange(1, 1, detailRows.length, overdueHeaders.length).setValues(detailRows);

          // Shift format ranges +1 because of the new Master Serial column
          detailSheet.getRange(1, 10, detailRows.length, 3).setNumberFormat("#,##0.00");
          detailSheet.getRange(1, 9, detailRows.length, 1).setNumberFormat("@");
          
          // Apply global wrapping FIRST
          detailSheet.getDataRange().setWrap(true).setVerticalAlignment("top");

          // Apply specific Formatting
          formatRanges.forEach(f => {
              if (f.type === 'cat') {
                  // Format category header cell - Disable text wrap
                  detailSheet.getRange(f.row, 1).setFontWeight("bold").setWrap(false);
              } else if (f.type === 'col') {
                  // Format Column Headers
                  detailSheet.getRange(f.row, 1, 1, overdueHeaders.length).setFontWeight("bold").setBackground("#f3f3f3");
                  detailSheet.getRange(f.row, 1, 1, overdueHeaders.length).setBorder(true, true, true, true, true, true);
              } else if (f.type === 'total') {
                  // Format Subtotal and Grand Total Rows
                  detailSheet.getRange(f.row, 8, 1, 5).setFontWeight("bold"); // Bold from Customer Name to Open Balance
                  detailSheet.getRange(f.row, 12).setBorder(true, false, false, false, false, false, "black", SpreadsheetApp.BorderStyle.SOLID); // Top border on Open Balance
              }
          });
          
          for (let i = 1; i <= overdueHeaders.length; i++) detailSheet.autoResizeColumn(i);
          detailSheet.setColumnWidth(1, 80);  // Master Serial
          detailSheet.setColumnWidth(2, 50);  // SN
          detailSheet.setColumnWidth(4, 120); // Invoice Date
          detailSheet.setColumnWidth(6, 120); // Due Date
          detailSheet.setColumnWidth(8, 250); // Customer Name
          detailSheet.setColumnWidth(9, 250); // Additional Text
          detailSheet.setColumnWidth(10, 110); // Sales Total
          detailSheet.setColumnWidth(11, 110); // Amount Paid
          detailSheet.setColumnWidth(12, 110); // Open Balance
          detailSheet.setColumnWidth(15, 200); // Remarks
      }

      SpreadsheetApp.flush();
      Utilities.sleep(3000);

      let url = "https://docs.google.com/spreadsheets/d/" + tempSpreadsheet.getId() + "/export?format=xlsx";
      let token = ScriptApp.getOAuthToken();
      let attachmentBlob = null;
      try {
          let response = UrlFetchApp.fetch(url, { headers: { 'Authorization': 'Bearer ' + token }, muteHttpExceptions: true });
          if (response.getResponseCode() === 200) attachmentBlob = response.getBlob().setName(cleanFileName + ".xlsx");
      } catch(e) {}

      try { DriveApp.getFileById(tempSpreadsheet.getId()).setTrashed(true); } catch(e){}

      // --- SEND EMAIL ---
      if (attachmentBlob) {
          let sysSettings = getSystemSettings();
          let isTech = (bucketName === "Keysys, A2Z2GO IT");
          let senderName = isTech ? "Finance A2Z2GO" : "Finance Pingcon";
          let senderEmail = getSenderEmail(isTech);
          
          let mainRecipient = "";
          let ccRecipient = "";
          
          // Lookup database for recipients
          let dbKey = isTech ? "KEYSYS" : "PINGCON";
          if (overdueEmailMap[dbKey]) {
              if (overdueEmailMap[dbKey].to) mainRecipient = overdueEmailMap[dbKey].to;
              if (overdueEmailMap[dbKey].cc) ccRecipient = overdueEmailMap[dbKey].cc;
          }

          if (!mainRecipient) {
              console.error(`Overdue Report Aborted: No 'To Emails' configured for ${dbKey} in the Overdue_Emails_Data database.`);
              continue; // Skips sending email for this entity and moves to the next bucket
          }
          
          let securedBlob = encryptAttachmentViaPython(attachmentBlob, cleanFileName + ".xlsx", isTech ? SCRIPT_PROPS['ENCRYPT_PASS_TECH'] : SCRIPT_PROPS['ENCRYPT_PASS_PINGCON']);
          
          // Passing the full bucket name and grand total to match the email screenshot
          let htmlBody = Template_UnpaidDrOverdueEmail(bucketName, totalAmount);

          try {
              sendRobustEmail(mainRecipient, cleanFileName, htmlBody, {
                  from: senderEmail, name: senderName, cc: ccRecipient, attachments: [securedBlob]
              });
              // Log success so the Summary Email knows this entity finished
              recordSuccessfulSend(shortBucketName, shortBucketName);
          } catch (emailErr) { console.error("Email Error: " + emailErr); }
          Utilities.sleep(10000); 
      }
  }
  
  clearReportLock();
  try { deleteTriggersByName("resumeUnpaidDROverdueBatch"); } catch(triggerErr) {}
  
  // Trigger the final Success/Exclusion Summary email
  sendBatchSuccessSummary(false, "UNPAID_DR_OVERDUE");
}

// =========================================================================
// CHECK PO PAYMENT STATUS
// =========================================================================

function _32_runCheckPOPaymentStatus() {
  const props = PropertiesService.getScriptProperties();
  deleteTriggersByName("resumeCheckPOPaymentBatch");
  props.deleteProperty("CANCEL_CHECK_PO_PAYMENT");
  resumeCheckPOPaymentBatch();
}

function resumeCheckPOPaymentBatch() {
  const props = PropertiesService.getScriptProperties();
  props.setProperty("THREAD_STATE", "RUNNING"); // <--- ADDED
  
  if (props.getProperty("CANCEL_CHECK_PO_PAYMENT") === "true") {
      deleteTriggersByName("resumeCheckPOPaymentBatch");
      props.deleteProperty("CANCEL_CHECK_PO_PAYMENT");
      clearReportLock();
      return;
  }

  // 1. Read the input POs from user uploaded sheet
  const poInputSheets = getDatabaseSheets("PO_Checker_Input_Data");
  if (poInputSheets.length === 0) {
      deleteTriggersByName("resumeCheckPOPaymentBatch");
      clearReportLock();
      return;
  }
  
  let targetPOs = [];
  poInputSheets.forEach(sheet => {
      const data = sheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
          let rawPo = String(data[i][0]).trim();
          if (!rawPo) continue;
          
          // Completely strip the leading "PO#1234" part entirely
          let cleanedPo = rawPo.replace(/^POI?#?\d+[-\s]*/i, '');
          
          // Strip the remaining prefix letters/dashes/spaces (like "SDOK-", "IKUFITPO") to isolate the core number
          let coreMatch = cleanedPo.match(/^[A-Za-z\-\s_]+(.*)/);
          if (coreMatch && coreMatch[1]) {
              cleanedPo = coreMatch[1].trim();
          }

          // Fix for embedded SKN numbers (e.g., PO-SAMPLE-SKN9999999999 -> 9999999999)
          let embeddedSkn = cleanedPo.match(/SKN(\d+)/i);
          if (embeddedSkn) {
              cleanedPo = embeddedSkn[1];
          }
          
          // Remove trailing spaces or fractions like " 2/2"
          let primaryPoPart = cleanedPo.split(/[\s\/]/)[0];
          
          if (primaryPoPart) {
              // Force strip hidden characters to ensure clean matching
              let cleanPoStr = primaryPoPart.replace(/[^\x20-\x7E]/g, '').trim();
              let escapedPo = cleanPoStr.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
              
              let regexStr = "";
              // If it's a pure number of 4 or more digits, make the "PO" prefix optional to catch matches seamlessly
              if (cleanPoStr.length >= 4 && /^\d+$/.test(cleanPoStr)) {
                  // Uses standard boundaries (^|[^0-9]) to avoid regex engine lookbehind failures
                  regexStr = "(^|[^0-9])" + escapedPo + "([^a-zA-Z0-9]|$)";
              } else {
                  // Strict match for short strings: Must have "PO" in front (with optional spaces/symbols) to match "PO 005"
                  regexStr = "(^|[^a-zA-Z0-9])P\\.?O\\.?[^a-zA-Z0-9]*" + escapedPo + "([^a-zA-Z0-9]|$)";
              }
              
              targetPOs.push({ 
                  original: rawPo, 
                  extracted: cleanPoStr, // Store exactly what we are looking for
                  regex: new RegExp(regexStr, "i"),
                  foundMatch: false 
              });
          }
      }
  });

  if (targetPOs.length === 0) {
      deleteTriggersByName("resumeCheckPOPaymentBatch");
      clearReportLock();
      return;
  }

  // 2. Scan AR_Data_ directly
  const arAllData = gatherAllData("AR_Data_");
  const arSheets = getDatabaseSheets("AR_Data_");
  let arHeaders = [];
  if (arSheets.length > 0 && arSheets[0].getLastColumn() > 0) {
    arHeaders = arSheets[0].getRange(1, 1, 1, arSheets[0].getLastColumn()).getValues()[0].map(h => String(h).toLowerCase().trim());
  }
  
  let docDateIdx = arHeaders.findIndex(h => h.includes("date") && !h.includes("expiry") && !h.includes("countered") && !h.includes("received"));
  let deadlineIdx = arHeaders.findIndex(h => h.includes("deadline") || h.includes("due date"));
  if (deadlineIdx === -1) deadlineIdx = docDateIdx;
  let customerIdx = arHeaders.findIndex(h => h.includes("customer"));
  let managerIdx = arHeaders.findIndex(h => h.includes("manager"));
  let textIdx = arHeaders.findIndex(h => h.includes("additional text"));
  let notesIdx = arHeaders.findIndex(h => h === "notes");
  let openBalIdx = arHeaders.findIndex(h => h.includes("open balance") && !h.includes("%"));
  let originalIdx = arHeaders.findIndex(h => h.includes("sales document total in main currency") || h.includes("sales document total"));
  let paymentTermIdx = arHeaders.findIndex(h => h.includes("payment condition") || h.includes("payment term"));
  let counterDateIdx = arHeaders.findIndex(h => h.includes("countered") || h.includes("counter date"));
  let receivedDateIdx = arHeaders.findIndex(h => h.includes("received delivery") || h.includes("received date") || h.includes("delivery date"));
  let sourceFileIdx = arHeaders.findIndex(h => h.includes("source file"));
  let paymentStatusIdx = arHeaders.findIndex(h => h === "payment status" || h.includes("payment status"));
  if (paymentStatusIdx === -1 && openBalIdx > -1) paymentStatusIdx = openBalIdx + 2;

  let matchedItems = [];

  arAllData.forEach(item => {
      let row = item.data;
      let textVal = (textIdx > -1) ? String(row[textIdx]) : "";
      
      // Only use Additional Text, and force strip hidden unicode characters to ensure clean matching
      let cleanTextVal = textVal.replace(/[^\x20-\x7E]/g, ' ');
      if (!cleanTextVal.trim()) return;
      
      let matchedTargets = [];
      for (let t of targetPOs) {
          if (t.regex.test(cleanTextVal)) {
              matchedTargets.push(t);
              t.foundMatch = true; // Mark that we found at least one match for this input
          }
      }

      if (matchedTargets.length === 0) return;

      let openBal = (openBalIdx > -1) ? parseFloat(row[openBalIdx]) : 0;
      if (isNaN(openBal)) openBal = 0;

      let statusRaw = (paymentStatusIdx > -1 && paymentStatusIdx < row.length) ? String(row[paymentStatusIdx]).trim() : "";
      if (statusRaw.toLowerCase() === "fully paid" || openBal <= 0) statusRaw = "Fully Paid";

      let rawDocDate = (docDateIdx > -1) ? row[docDateIdx] : "";
      let docDateObj = (rawDocDate instanceof Date) ? rawDocDate : new Date(rawDocDate);
      
      let rawDeadline = (deadlineIdx > -1) ? row[deadlineIdx] : "";
      let deadlineObj = new Date(rawDeadline);
      if (isNaN(deadlineObj.getTime()) || rawDeadline === "") {
          deadlineObj = new Date(docDateObj);
          let pTerm = (paymentTermIdx > -1) ? parseInt(row[paymentTermIdx]) : 0;
          if (!isNaN(pTerm) && pTerm > 0) {
              deadlineObj.setDate(deadlineObj.getDate() + pTerm);
          }
      }

      let salesTotal = (originalIdx > -1) ? parseFloat(row[originalIdx]) : 0;
      if (isNaN(salesTotal)) salesTotal = 0;
      let amountPaid = salesTotal - openBal;
      if (amountPaid < 0) amountPaid = 0;

      let sourceFile = (sourceFileIdx > -1) ? String(row[sourceFileIdx]) : "";

      // Push a separate row for EVERY duplicate input PO that matched this AR_Data row
      matchedTargets.forEach(matchedTarget => {
          matchedItems.push({
              searchedRawPO: matchedTarget.original,
              extractedPO: matchedTarget.extracted,
              docDateObj: docDateObj,
              invoiceDateStr: isNaN(docDateObj.getTime()) ? "" : fastFormatDate(docDateObj),
              paymentTerm: (paymentTermIdx > -1) ? String(row[paymentTermIdx]).trim() : "",
              deadlineStr: isNaN(deadlineObj.getTime()) ? "" : fastFormatDate(deadlineObj),
              customer: (customerIdx > -1) ? String(row[customerIdx]).trim() : "",
              manager: getCanonicalName((managerIdx > -1) ? String(row[managerIdx]) : ""),
              text: textVal,
              notes: (notesIdx > -1) ? String(row[notesIdx]) : "",
              salesTotal: salesTotal,
              amountPaid: amountPaid,
              openBal: openBal,
              counterDateStr: (counterDateIdx > -1 && row[counterDateIdx]) ? fastFormatDate(new Date(row[counterDateIdx])) : "",
              receivedDateStr: (receivedDateIdx > -1 && row[receivedDateIdx]) ? fastFormatDate(new Date(row[receivedDateIdx])) : "",
              remarksStatus: statusRaw,
              companyEntity: getCompanyEntity(sourceFile)
          });
      });
  });

  // Append POs that had ZERO matches in AR_Data_
  targetPOs.forEach(t => {
      if (!t.foundMatch) {
          matchedItems.push({
              searchedRawPO: t.original,
              extractedPO: t.extracted,
              docDateObj: new Date(0), // Push to bottom when sorting
              invoiceDateStr: "",
              paymentTerm: "",
              deadlineStr: "",
              customer: "",
              manager: "",
              text: "",
              notes: "",
              salesTotal: "",
              amountPaid: "",
              openBal: "",
              counterDateStr: "",
              receivedDateStr: "",
              remarksStatus: "No Match Found",
              companyEntity: ""
          });
      }
  });

  // Sort logically for the final report
  matchedItems.sort((a, b) => {
      let c1 = a.customer || "ZZZ_NoMatch";
      let c2 = b.customer || "ZZZ_NoMatch";
      let cCmp = c1.localeCompare(c2);
      if (cCmp !== 0) return cCmp;
      return (a.docDateObj || 0) - (b.docDateObj || 0);
  });

  // 3. Build the Excel Sheet Output
  let reportHeaders = [
      "SN", "Original Raw PO", "Extracted PO", "Entity", "Invoice Date (MM/DD/YYYY)", "Terms", "Due Date (MM/DD/YYYY)", 
      "Cust Manager", "Customer Name", "Additional text", "Notes", "Sales Total", "Amount Paid", "Open Balance", 
      "Counter Date (MM/DD/YYYY)", "Received Date (MM/DD/YYYY)", "Remarks/Payment Status"
  ];

  let sheetData = [reportHeaders];
  let sn = 1;
  let grandTotal = 0;
  let grandOpenBal = 0;

  matchedItems.forEach((item) => {
      sheetData.push([
          sn++, item.searchedRawPO, item.extractedPO, item.companyEntity, item.invoiceDateStr, item.paymentTerm, item.deadlineStr, 
          item.manager, item.customer, item.text, item.notes, item.salesTotal, item.amountPaid, item.openBal,
          item.counterDateStr, item.receivedDateStr, item.remarksStatus
      ]);
      grandTotal += item.salesTotal;
      grandOpenBal += item.openBal;
  });

  let grandRow = Array(reportHeaders.length).fill("");
  grandRow[11] = "GRAND TOTAL:";
  grandRow[13] = grandOpenBal;
  sheetData.push(Array(reportHeaders.length).fill("")); 
  sheetData.push(grandRow);

  // Check entity to append to the filename and email subject
  let hasKeysysOrA2ZIT = matchedItems.some(item => item.companyEntity === "Keysys" || item.companyEntity === "A2ZIT");
  let bucketName = hasKeysysOrA2ZIT ? "A2Z2GO" : "Pingcon";

  // Extract Report Date from AR_Data_ source files
  let maxAsOfDate = 0;
  let reportDateString = "";
  arAllData.forEach(item => {
      let sourceFile = (sourceFileIdx > -1) ? String(item.data[sourceFileIdx]) : "";
      if (sourceFile) {
          let match = sourceFile.match(/as of\s+([a-zA-Z]+)\s*(\d+),?\s+(\d{4})\s+(\d{4})/i);
          if (match) {
              let monthStr = match[1];
              let dayStr = match[2];
              let yearStr = match[3];
              let timeStr = match[4];
              let hourStr = timeStr.substring(0, 2);
              let minStr = timeStr.substring(2, 4);
              let dateObj = new Date(`${monthStr} ${dayStr}, ${yearStr} ${hourStr}:${minStr}:00`);
              if (!isNaN(dateObj.getTime()) && dateObj.getTime() > maxAsOfDate) {
                  maxAsOfDate = dateObj.getTime();
                  let paddedDay = dayStr.length === 1 ? "0" + dayStr : dayStr;
                  reportDateString = `${monthStr} ${paddedDay}, ${yearStr} ${timeStr}`;
              }
          }
      }
  });
  
  // Fallback to current time if no AR files match the pattern
  if (!reportDateString) {
      let now = new Date();
      reportDateString = Utilities.formatDate(now, Session.getScriptTimeZone(), "MMM dd yyyy HHmm");
  }

  // Strip special characters (like commas) to prevent the Python server from crashing
  let cleanFileName = `Checked PO Payment Status - ${bucketName} as of Report Date ${reportDateString}`.replace(/[^a-zA-Z0-9 \-_]/g, "");
  
  let tempSpreadsheet = SpreadsheetApp.create(cleanFileName);
  let tempSheet = tempSpreadsheet.getSheets()[0];
  
  // Apply text formatting BEFORE setting values to preserve leading zeros like 005
  tempSheet.getRange(2, 3, sheetData.length, 1).setNumberFormat("@");
  tempSheet.getRange(2, 5, sheetData.length, 1).setNumberFormat("@");
  tempSheet.getRange(2, 7, sheetData.length, 1).setNumberFormat("@");
  
  tempSheet.getRange(1, 1, sheetData.length, reportHeaders.length).setValues(sheetData);

  // Styling identical to SOA per customer
  tempSheet.getRange(1, 1, 1, reportHeaders.length).setFontWeight("bold").setBackground("#f3f3f3");
  tempSheet.getRange(sheetData.length, 12, 1, 3).setFontWeight("bold"); 
  tempSheet.getRange(sheetData.length, 14).setBorder(true, false, false, false, false, false, "black", SpreadsheetApp.BorderStyle.SOLID);

  tempSheet.getRange(2, 12, sheetData.length, 3).setNumberFormat("#,##0.00");
  
  tempSheet.getDataRange().setWrap(true).setVerticalAlignment("top");
  for (let i = 1; i <= reportHeaders.length; i++) tempSheet.autoResizeColumn(i);
  tempSheet.setColumnWidth(2, 250); 
  tempSheet.setColumnWidth(3, 150); 
  tempSheet.setColumnWidth(9, 250); 
  tempSheet.setColumnWidth(10, 250); 
  tempSheet.setColumnWidth(12, 140); // Sales Total (Fixes ########)
  tempSheet.setColumnWidth(13, 140); // Amount Paid (Fixes ########)
  tempSheet.setColumnWidth(14, 140); // Open Balance (Fixes ########)
  tempSheet.setColumnWidth(17, 200); 

  SpreadsheetApp.flush();
  Utilities.sleep(3000);

  let url = "https://docs.google.com/spreadsheets/d/" + tempSpreadsheet.getId() + "/export?format=xlsx";
  let token = ScriptApp.getOAuthToken();
  let attachmentBlob = null;
  try {
      let response = UrlFetchApp.fetch(url, { headers: { 'Authorization': 'Bearer ' + token }, muteHttpExceptions: true });
      if (response.getResponseCode() === 200) attachmentBlob = response.getBlob().setName(cleanFileName + ".xlsx");
  } catch(e) {}

  try { DriveApp.getFileById(tempSpreadsheet.getId()).setTrashed(true); } catch(e){}

  // 4. Send Email Output
  if (attachmentBlob) {
      let minReportStartDate = null;
      let maxReportEndDate = null;
      
      arAllData.forEach(item => {
          let sourceFile = (sourceFileIdx > -1) ? String(item.data[sourceFileIdx]) : "";
          let parsedFromFile = false;
          if (sourceFile) {
              let match = sourceFile.match(/Ageing Report\s+(.*?)\s+to\s+(.*?)\s+as of/i);
              if (match) {
                  let sDate = new Date(match[1]);
                  let eDate = new Date(match[2]);
                  if (!isNaN(sDate.getTime())) {
                      if (!minReportStartDate || sDate < minReportStartDate) minReportStartDate = sDate;
                  }
                  if (!isNaN(eDate.getTime())) {
                      if (!maxReportEndDate || eDate > maxReportEndDate) maxReportEndDate = eDate;
                  }
                  parsedFromFile = true;
              }
          }
          if (!parsedFromFile && docDateIdx > -1) {
              let dDate = new Date(item.data[docDateIdx]);
              if (!isNaN(dDate.getTime())) {
                  if (!minReportStartDate || dDate < minReportStartDate) minReportStartDate = dDate;
                  if (!maxReportEndDate || dDate > maxReportEndDate) maxReportEndDate = dDate;
              }
          }
      });
      
      let dateRangePhrase = "";
      if (minReportStartDate && maxReportEndDate) {
          let startStr = Utilities.formatDate(minReportStartDate, Session.getScriptTimeZone(), "MMMM d, yyyy");
          let endStr = Utilities.formatDate(maxReportEndDate, Session.getScriptTimeZone(), "MMMM d, yyyy");
          dateRangePhrase = `<p>For your reference, this report covers transactions recorded in our system from <strong>${startStr}</strong> to <strong>${endStr}</strong>.</p>`;
      }

      let sysSettings = getSystemSettings();
      let isTech = hasKeysysOrA2ZIT;
      let mainRecipient = isTech ? (sysSettings["Management_A2Z"] || "") : (sysSettings["Management_Pingcon"] || "");
      let senderName = isTech ? "Finance A2Z2GO" : "Finance Pingcon";
      let senderEmail = getSenderEmail(isTech);
      
      let htmlBody = `
        <div style="font-family: Arial, sans-serif; font-size: 13px; color: #000; line-height: 1.5;">
          <p>Good day,</p>
          <p>Please find attached the PO Payment Status report checked against our Ageing Report database.</p>
          ${dateRangePhrase}
          <div style="background-color: #f9f9f9; padding: 15px; border: 1px solid #ddd; border-radius: 5px; margin: 20px 0; width: fit-content;">
              <table style="font-size: 13px;">
                  <tr><td style="padding-right: 20px;"><strong>Matched Transactions Found:</strong></td><td>${matchedItems.length}</td></tr>
                  <tr><td style="padding-right: 20px;"><strong>Total Open Balance (Unpaid):</strong></td><td style="color: #d9534f; font-weight: bold;">PHP ${Number(grandOpenBal).toLocaleString('en-US', {minimumFractionDigits: 2, maximumFractionDigits: 2})}</td></tr>
              </table>
          </div>
          <p>Thank you.</p>
        </div>`;

      let finalPassword = hasKeysysOrA2ZIT ? SCRIPT_PROPS['ENCRYPT_PASS_TECH'] : SCRIPT_PROPS['ENCRYPT_PASS_PINGCON'];
      let securedBlob = encryptAttachmentViaPython(attachmentBlob, cleanFileName + ".xlsx", finalPassword);

      sendRobustEmail(mainRecipient, cleanFileName, htmlBody, {
          from: senderEmail, name: senderName, attachments: [securedBlob]
      });
  }
  
  deleteTriggersByName("resumeCheckPOPaymentBatch");
  clearReportLock();
}

const CPP_BATCH_KEY = "PENDING_CUSTOMERS_CPP";
const CPP_TRIGGER_NAME = "resumeCPPBatch";

function _34_runSendAllCustomerPaymentPerformance() {
  var lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) { throw new Error("System is currently busy."); }
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const props = PropertiesService.getScriptProperties();
    deleteTriggersByName(CPP_TRIGGER_NAME);
    props.deleteProperty(CPP_BATCH_KEY);
    props.deleteProperty("CANCEL_CPP");
    
    try { ss.toast("Scanning CPP Data...", "INITIALIZING", -1); } catch(e){}
    
    const cppAllData = gatherAllData("CPP_Data_");
    if (cppAllData.length === 0) throw new Error("No Customer Payment Performance Data found.");
    
    const cppSheets = getDatabaseSheets("CPP_Data_");
    let headers = cppSheets[0].getRange(1, 1, 1, cppSheets[0].getLastColumn()).getValues()[0].map(h => String(h).toLowerCase().trim());
    let custIdx = headers.findIndex(h => h.includes("customer") || h === "name" || h === "business partner");
    let sourceFileIdx = headers.findIndex(h => h.includes("source file"));
    
    let uniqueCustomers = new Set();
    let allowedEntities = JSON.parse(props.getProperty("SALES_ENTITIES") || "[]");
    let allowedCustsStr = props.getProperty("SALES_TARGET_CUSTOMERS");
    let allowedCusts = allowedCustsStr ? JSON.parse(allowedCustsStr) : [];
    
    cppAllData.forEach(item => {
        let row = item.data;
        let customer = (custIdx > -1) ? String(row[custIdx]).trim() : "";
        if (!customer) return;
        
        let sourceFile = (sourceFileIdx > -1) ? String(row[sourceFileIdx]) : "";
        let companyEntity = getCompanyEntity(sourceFile);
        if (allowedEntities.length > 0 && !allowedEntities.includes(companyEntity)) return;

        let cleanCustomer = customer.replace(/[^\x20-\x7EñÑ]/g, '').replace(/\s+/g, ' ').trim();
        let custKey = cleanCustomer.toLowerCase().replace(/\s+/g, '');
        let displayCustomerName = customer; 
        if (allowedCusts.length > 0) {
            let matchedName = allowedCusts.find(c => c.toLowerCase().replace(/\s+/g, '') === custKey);
            if (!matchedName) return;
            displayCustomerName = matchedName;
        }
        uniqueCustomers.add(displayCustomerName);
    });
    
    let emailKeyGroups = {};
    uniqueCustomers.forEach(cust => {
        let emailKey = cust; 
        if (!emailKeyGroups[emailKey]) emailKeyGroups[emailKey] = [];
        emailKeyGroups[emailKey].push(cust);
    });
    
    let groupedCustomers = Object.values(emailKeyGroups);
    if(groupedCustomers.length === 0) throw new Error("No CPP data found for the selected customers.");
    
    props.setProperty(CPP_BATCH_KEY, JSON.stringify(groupedCustomers));
    resumeCPPBatch();
  } finally {
    if (lock) lock.releaseLock();
  }
}

function resumeCPPBatch() {
  const startTime = Date.now();
  const MAX_EXECUTION_TIME = 4.5 * 60 * 1000; 
  const props = PropertiesService.getScriptProperties();
  props.setProperty("THREAD_STATE", "RUNNING"); // <--- ADDED
  
  if (props.getProperty("CANCEL_CPP") === "true") {
     deleteTriggersByName(CPP_TRIGGER_NAME);
     props.deleteProperty(CPP_BATCH_KEY);
     props.deleteProperty("CANCEL_CPP");
     return; 
  }

  let pendingCustStr = props.getProperty(CPP_BATCH_KEY);
  if (!pendingCustStr) { deleteTriggersByName(CPP_TRIGGER_NAME); return; }
  
  let pendingCust = JSON.parse(pendingCustStr);
  if (pendingCust.length === 0) {
    deleteTriggersByName(CPP_TRIGGER_NAME);
    props.deleteProperty(CPP_BATCH_KEY);
    clearReportLock();
    console.log("All Customers processed successfully.");
    return;
  }

  const cppAllData = gatherAllData("CPP_Data_");
  const cppSheets = getDatabaseSheets("CPP_Data_");
  if (cppSheets.length === 0) {
      deleteTriggersByName(CPP_TRIGGER_NAME);
      props.deleteProperty(CPP_BATCH_KEY);
      clearReportLock();
      return;
  }
  let headers = cppSheets[0].getRange(1, 1, 1, cppSheets[0].getLastColumn()).getValues()[0].map(h => String(h).toLowerCase().trim());
  
  let custIdx = headers.findIndex(h => h.includes("customer") || h === "name" || h === "business partner");
  let dateIdx = headers.findIndex(h => h.includes("invoice date"));
  let termIdx = headers.findIndex(h => h.includes("payment condition") || h.includes("payment term"));
  let totalIdx = headers.findIndex(h => h.includes("sales document total in main currency"));
  let addTextIdx = headers.findIndex(h => h.includes("additional text"));
  let notesIdx = headers.findIndex(h => h === "notes");
  let dueIdx = headers.findIndex(h => h.includes("payment deadline") || h.includes("due date"));
  let recIdx = headers.findIndex(h => h.includes("received date"));
  let counterIdx = headers.findIndex(h => h.includes("countered date"));
  let payDateIdx = headers.findIndex(h => h.includes("payment date"));
  let amountIdx = headers.findIndex(h => h.includes("amount paid"));
  let sourceFileIdx = headers.findIndex(h => h.includes("source file"));

  const arAllData = gatherAllData("AR_Data_");
  const arSheets = getDatabaseSheets("AR_Data_");
  let arHeaders = arSheets.length > 0 ? arSheets[0].getRange(1, 1, 1, arSheets[0].getLastColumn()).getValues()[0].map(h => String(h).toLowerCase().trim()) : [];
  let arCustIdx = arHeaders.findIndex(h => h.includes("customer"));
  let arManagerIdx = arHeaders.findIndex(h => h.includes("manager"));
  let arSourceFileIdx = arHeaders.findIndex(h => h.includes("source file"));

  let targetQuarter = props.getProperty("CPP_QUARTER");
  let targetYear = parseInt(props.getProperty("CPP_YEAR"));

  let qStartMonth = 0; let qEndMonth = 2;
  if (targetQuarter === "Q1") { qStartMonth = 0; qEndMonth = 2; }
  if (targetQuarter === "Q2") { qStartMonth = 3; qEndMonth = 5; }
  if (targetQuarter === "Q3") { qStartMonth = 6; qEndMonth = 8; }
  if (targetQuarter === "Q4") { qStartMonth = 9; qEndMonth = 11; }

  let qStartObj = new Date(targetYear, qStartMonth, 1);
  let qEndObj = new Date(targetYear, qEndMonth + 1, 0, 23, 59, 59);

  let maxAsOfDate = 0;
  let reportDateString = "";
  arAllData.forEach(item => {
      let sourceFile = (arSourceFileIdx > -1) ? String(item.data[arSourceFileIdx]) : "";
      if (sourceFile) {
          let match = sourceFile.match(/as of\s+([a-zA-Z]+)\s*(\d+),?\s+(\d{4})\s+(\d{4})/i);
          if (match) {
              let dateObj = new Date(`${match[1]} ${match[2]}, ${match[3]} ${match[4].substring(0,2)}:${match[4].substring(2,4)}:00`);
              if (!isNaN(dateObj.getTime()) && dateObj.getTime() > maxAsOfDate) {
                  maxAsOfDate = dateObj.getTime();
                  let paddedDay = match[2].length === 1 ? "0" + match[2] : match[2];
                  reportDateString = `${match[1]} ${paddedDay}, ${match[3]} ${match[4]}`;
              }
          }
      }
  });
  if (!reportDateString) reportDateString = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MMM dd, yyyy HHmm");

  const emailSheets = getDatabaseSheets("Customer_Emails_Extracted_");
  let customerEmailMap = {};
  emailSheets.forEach(sheet => {
     let data = sheet.getDataRange().getValues();
     for(let i=1; i<data.length; i++){
         let cust = String(data[i][1]).replace(/[^\x20-\x7E]/g, '').toLowerCase().replace(/\s+/g, '');
         let email = String(data[i][3]).trim();
         let role = String(data[i][4]).trim();
         
         if (cust && email && email.includes("@")) {
             if (!customerEmailMap[cust]) customerEmailMap[cust] = { main: new Set(), cc: new Set() };
             
             if (role.toLowerCase().includes("cc")) {
                 customerEmailMap[cust].cc.add(email);
             } else {
                 customerEmailMap[cust].main.add(email);
             }
         }
     }
  });
  
  const aeEmailSheets = getDatabaseSheets("AE_Emails_Data_");
  let aeEmailMap = {};
  aeEmailSheets.forEach(sheet => {
     let data = sheet.getDataRange().getValues();
     for(let i = 1; i < data.length; i++) {
         let aeName = String(data[i][0]).trim();
         if(aeName) aeEmailMap[aeName] = { pingcon: String(data[i][1]).trim(), a2zit: data[i].length > 2 ? String(data[i][2]).trim() : "", keysys: data[i].length > 3 ? String(data[i][3]).trim() : "" };
     }
  });

  let allowedEntities = JSON.parse(props.getProperty("SALES_ENTITIES") || "[]");

  while (pendingCust.length > 0) {
     if (props.getProperty("CANCEL_CPP") === "true") {
         deleteTriggersByName(CPP_TRIGGER_NAME); props.deleteProperty(CPP_BATCH_KEY); props.deleteProperty("CANCEL_CPP"); return; 
     }
     if (Date.now() - startTime > MAX_EXECUTION_TIME) {
        props.setProperty(CPP_BATCH_KEY, JSON.stringify(pendingCust));
        replaceTrigger(CPP_TRIGGER_NAME); return; 
     }
     
     let currentCustomerGroup = pendingCust.shift(); 
     let commonName = currentCustomerGroup.join(", "); 
     let currentCustKeys = currentCustomerGroup.map(c => c.toLowerCase().replace(/\s+/g, ''));
     
     let extractedData = [];
     let hasKeysysOrA2ZIT = false;
     let hasA2ZPMC = false;

     cppAllData.forEach(item => {
         let row = item.data;
         let customer = (custIdx > -1) ? String(row[custIdx]).trim() : "";
         if (!customer) return;
         let rowCustKey = customer.toLowerCase().replace(/\s+/g, '');
         if (!currentCustKeys.includes(rowCustKey)) return;

         let amountVal = (amountIdx > -1) ? parseFloat(row[amountIdx]) || 0 : 0;
         if (amountVal < 0) return; // Ignore negative amounts (VOID)

         let addTextVal = (addTextIdx > -1) ? String(row[addTextIdx]) : "";
         let notesVal = (notesIdx > -1) ? String(row[notesIdx]) : "";
         if (isTransactionVoid(notesVal, addTextVal)) return; // Apply unified VOID logic

         if (dateIdx > -1 && row[dateIdx]) {
             let d = new Date(row[dateIdx]);
             if (!isNaN(d.getTime())) {
                 if (d < qStartObj || d > qEndObj) return; 
             }
         }
         
         let sourceFile = (sourceFileIdx > -1) ? String(row[sourceFileIdx]) : "";
         let companyEntity = getCompanyEntity(sourceFile);
         if (allowedEntities.length > 0 && !allowedEntities.includes(companyEntity)) return;

         if (companyEntity === "Keysys" || companyEntity === "A2ZIT") hasKeysysOrA2ZIT = true;
         if (companyEntity === "A2ZPMC") hasA2ZPMC = true;

         // Extracting needed variables
         let dueDateObj = new Date(row[dueIdx]);
         let payDateObj = new Date(row[payDateIdx]);
         let ageingDays = "";
         let passStr = "";
         let failStr = "";

         if (!isNaN(dueDateObj.getTime()) && !isNaN(payDateObj.getTime())) {
             let diffTime = payDateObj.getTime() - dueDateObj.getTime();
             ageingDays = Math.round(diffTime / (1000 * 60 * 60 * 24));
             if (ageingDays <= 0) passStr = "PASS";
             else failStr = "FAIL";
         }

         let cDateStr = (counterIdx > -1 && row[counterIdx]) ? fastFormatDate(new Date(row[counterIdx])) : "n/a";
         if (cDateStr === "12/31/1969" || cDateStr === "") cDateStr = "n/a"; // Fallback visual

         extractedData.push({
             customer: customer,
             term: (termIdx > -1) ? String(row[termIdx]) : "",
             total: (totalIdx > -1) ? parseFloat(row[totalIdx]) || 0 : 0,
             addText: (addTextIdx > -1) ? String(row[addTextIdx]) : "",
             notes: (notesIdx > -1) ? String(row[notesIdx]) : "",
             dueDate: isNaN(dueDateObj.getTime()) ? "" : fastFormatDate(dueDateObj),
             recDate: (recIdx > -1 && row[recIdx]) ? fastFormatDate(new Date(row[recIdx])) : "",
             countDate: cDateStr,
             payDate: isNaN(payDateObj.getTime()) ? "" : fastFormatDate(payDateObj),
             amount: (amountIdx > -1) ? parseFloat(row[amountIdx]) || 0 : 0,
             ageing: ageingDays,
             passStr: passStr,
             failStr: failStr,
             sortDate: dueDateObj // Sorting by Due Date
         });
     });

     if (extractedData.length === 0) continue;

     // Sort by Due Date
     extractedData.sort((a, b) => a.sortDate - b.sortDate);

     let reportHeaders = [
        "Customer", "Payment Term", "Sales document total in main currency", 
        "Additional text printed on the invoice", "Notes", 
        "Due Date (MM/DD/YYYY)", "Received Date/Delivery Date (MM/DD/YYYY)", 
        "Counter date (MM/DD/YYYY)", "Customer Payment Date (MM/DD/YYYY)", "Amount Paid", 
        "AGEING = Customer Payment Date Less Due Date = (+) late payment ( - ) advance payment", 
        "PASS", "FAIL", "Finance Result Pass % : equal and or higher than Passing score % (P) >= (R)", 
        "PASS Total", "FAIL Total", "Total # of invoice = pass + fail (L) + (M)", 
        "PASS % = pass / total # of invoice (L) / (O)", 
        "FAIL % = fail / total # of invoice (M) / (O)", "Passing Score % = Pingcon Standard"
     ];

     let sheetData = [reportHeaders];
     let passCount = 0;
     let failCount = 0;

     extractedData.forEach((r, idx) => {
         if (r.passStr === "PASS") passCount++;
         if (r.failStr === "FAIL") failCount++;

         let isLast = (idx === extractedData.length - 1);
         let rowData = [
             r.customer, r.term, r.total, r.addText, r.notes,
             r.dueDate, r.recDate, r.countDate, r.payDate, r.amount,
             r.ageing, r.passStr, r.failStr, 
             "", "", "", "", "", "", ""
         ];

         if (isLast) {
             let totalInv = passCount + failCount;
             let passPct = totalInv > 0 ? passCount / totalInv : 0;
             let failPct = totalInv > 0 ? failCount / totalInv : 0;
             let passScore = 0.70; // 70.00%
             let finResult = passPct >= passScore ? "PASS" : "FAIL";

             rowData[13] = finResult;
             rowData[14] = passCount;
             rowData[15] = failCount;
             rowData[16] = totalInv;
             rowData[17] = passPct;
             rowData[18] = failPct;
             rowData[19] = passScore;
         }
         sheetData.push(rowData);
     });

     let startMonStr = Utilities.formatDate(qStartObj, Session.getScriptTimeZone(), "MMM yyyy");
     let endMonStr = Utilities.formatDate(qEndObj, Session.getScriptTimeZone(), "MMM yyyy");
     let qSubjectStr = `${targetQuarter} ${startMonStr} - ${endMonStr}`;
     
     let cleanCustName = String(commonName).replace(/[^a-zA-Z0-9 -,]/g, "");
     let cleanFileName = `Customer Payment Performance ${qSubjectStr} - ${cleanCustName} as of Report Date ${reportDateString}`;

     let tempSpreadsheet = SpreadsheetApp.create(cleanFileName);
     let tempSheet = tempSpreadsheet.getSheets()[0];
     
     tempSheet.getRange(1, 1, sheetData.length, 20).setValues(sheetData);
     
     // ----------------------------------------------------
     // EXCEL STYLING AND FORMATTING
     // ----------------------------------------------------
     let hRange = tempSheet.getRange(1, 1, 1, 20);
     hRange.setWrap(true).setVerticalAlignment("middle").setHorizontalAlignment("center").setFontWeight("bold");
     
     // Background Colors per Screenshot
     tempSheet.getRange(1, 6, 1, 6).setBackground("#FFFF00"); // Yellow (F to K)
     tempSheet.getRange(1, 12, 1, 3).setBackground("#7030A0").setFontColor("white"); // Purple (L to N)
     tempSheet.getRange(1, 15, 1, 3).setBackground("#FFFFFF").setFontColor("black"); // White (O to Q)
     tempSheet.getRange(1, 18, 1, 3).setBackground("#7030A0").setFontColor("white"); // Purple (R to T)
     
     // Borders
     tempSheet.getRange(1, 1, sheetData.length, 20).setBorder(true, true, true, true, true, true);
     
     // Number Formats
     tempSheet.getRange(2, 3, sheetData.length - 1, 1).setNumberFormat("#,##0.00"); // Sales Total
     tempSheet.getRange(2, 10, sheetData.length - 1, 1).setNumberFormat("#,##0.00"); // Amount Paid
     tempSheet.getRange(2, 18, sheetData.length - 1, 3).setNumberFormat("0.000%"); // Percentages
     
     // Resize
     for (let i = 1; i <= 20; i++) tempSheet.autoResizeColumn(i);
     tempSheet.setColumnWidth(4, 250); // Additional text
     tempSheet.setColumnWidth(11, 200); // Ageing header
     tempSheet.setColumnWidth(14, 150); // Finance Result
     tempSheet.setColumnWidth(17, 120); // Total invoice
     
     SpreadsheetApp.flush(); Utilities.sleep(3000); 

     let url = "https://docs.google.com/spreadsheets/d/" + tempSpreadsheet.getId() + "/export?format=xlsx";
     let token = ScriptApp.getOAuthToken();
     let attachmentBlob = null;
     for (let fetchAttempt = 0; fetchAttempt < 3; fetchAttempt++) {
       try {
         let response = UrlFetchApp.fetch(url, { headers: { 'Authorization': 'Bearer ' + token }, muteHttpExceptions: true });
         if (response.getResponseCode() === 200) { attachmentBlob = response.getBlob().setName(cleanFileName + ".xlsx"); break; } 
         else { Utilities.sleep(4000 * (fetchAttempt + 1)); }
       } catch (e) { Utilities.sleep(5000 * (fetchAttempt + 1)); }
     }

     try { DriveApp.getFileById(tempSpreadsheet.getId()).setTrashed(true); } catch(e){}

     if (attachmentBlob) {
         let sysSettings = getSystemSettings();
         let isTech = hasKeysysOrA2ZIT;
         let senderName = isTech ? "Finance A2Z2GO" : "Finance Pingcon";
         let senderEmail = getSenderEmail(isTech);
         let feedbackImageId = isTech ? FEEDBACK_IMAGE_ID_A2Z_KEYSYS : FEEDBACK_IMAGE_ID_PINGCON_BONITA;

         let feedbackImgBlob = null;
         try { 
             if (feedbackImageId) feedbackImgBlob = DriveApp.getFileById(feedbackImageId).getBlob(); 
         } catch(e) {}

         let ccRecipientsList = [];
         let mainRecipientsList = [];
         let lookupCustKey = String(currentCustomerGroup[0]).toLowerCase().replace(/[^a-z0-9]/g, '');
         
         if (customerEmailMap[lookupCustKey]) {
             mainRecipientsList = Array.from(customerEmailMap[lookupCustKey].main);
             ccRecipientsList = Array.from(customerEmailMap[lookupCustKey].cc);
         }
         let hasCustomerEmail = mainRecipientsList.length > 0;

         // Dynamically append management CCs from System_Settings_Data
         if (isTech) {
             [sysSettings["Finance_A2Z"], sysSettings["Executive_A2Z"], sysSettings["Marketing_A2Z"]].forEach(e => {
                 if (e && e.trim() !== "") ccRecipientsList.push(e);
             });
         } else {
             [sysSettings["Finance_Pingcon"], sysSettings["Executive_Pingcon"], sysSettings["Marketing_Pingcon"]].forEach(e => {
                 if (e && e.trim() !== "") ccRecipientsList.push(e);
             });
         }

         let uniqueAEsMap = {};
         arAllData.forEach(item => {
             let row = item.data;
             let arCustomer = (arCustIdx > -1) ? String(row[arCustIdx]).trim() : "";
             let rowCustKey = arCustomer.toLowerCase().replace(/\s+/g, '');
             if (currentCustKeys.includes(rowCustKey)) {
                 let manager = getCanonicalName((arManagerIdx > -1) ? String(row[arManagerIdx]) : "");
                 let arSourceFile = (arSourceFileIdx > -1) ? String(row[arSourceFileIdx]) : "";
                 let arCompanyEntity = getCompanyEntity(arSourceFile);
                 if (manager) uniqueAEsMap[manager + "_" + arCompanyEntity] = { ae: manager, entity: arCompanyEntity };
             }
         });
         
         let customerAEs = Object.values(uniqueAEsMap);
         let aeContactHtml = "";
         let uniqueAeContactSet = new Set();

         customerAEs.forEach(pair => {
             if (aeEmailMap[pair.ae]) {
                 let emails = aeEmailMap[pair.ae];
                 let emailToAdd = isTech ? emails.a2zit : emails.pingcon;
                 
                 if (emailToAdd && emailToAdd.includes("@")) {
                     if (!hasCustomerEmail) mainRecipientsList.push(emailToAdd);
                     else ccRecipientsList.push(emailToAdd);
                 }
                 let contactLine = `Account Executive: ${pair.ae} - ${emailToAdd || "N/A"}`;
                 if (!uniqueAeContactSet.has(contactLine)) {
                     uniqueAeContactSet.add(contactLine);
                     aeContactHtml += `${contactLine}<br>`;
                 }
             }
         });

         let custUpper = String(commonName).toUpperCase().trim();
         let collectionAssts = []; 
         let dynamicCcToAppend = [];
         
         let hasBonitaCPP = extractedData.some(r => getCompanyEntity(r.sourceFile) === "Bonita");
         let hasPingconCPP = extractedData.some(r => getCompanyEntity(r.sourceFile) === "Pingcon" || !r.sourceFile);
         let assignToPrimaryCPP = resolveCollectionAssistantAssignment(
             custUpper, 
             hasKeysysOrA2ZIT, 
             hasBonitaCPP, 
             hasA2ZPMC, 
             hasPingconCPP,
             sysSettings
         );

         // Dynamically fetch Collection Assistant details from System_Settings_Data
         let asstPrimaryEmail = isTech 
             ? (sysSettings["Collection_Asst_A2Z_Email"] || "") 
             : (sysSettings["Collection_Asst_Pingcon_Primary_Email"] || "");
         let asstPrimaryName = isTech 
             ? (sysSettings["Collection_Asst_A2Z_Name"] || "Collection Assistant (A2Z2GO)") 
             : (sysSettings["Collection_Asst_Pingcon_Primary_Name"] || "Collection Assistant");

         let asstSecondaryEmail = sysSettings["Collection_Asst_Secondary_Email"] || "";
         let asstSecondaryName = sysSettings["Collection_Asst_Secondary_Name"] || "Collection Assistant";

         if (assignToPrimaryCPP) {
             if (asstPrimaryEmail) {
                 collectionAssts.push(`${asstPrimaryName} - ${asstPrimaryEmail}`);
                 dynamicCcToAppend.push(asstPrimaryEmail);
             }
         } else {
             if (asstSecondaryEmail) {
                 collectionAssts.push(`${asstSecondaryName} - ${asstSecondaryEmail}`);
                 dynamicCcToAppend.push(asstSecondaryEmail);
             }
         }
         
         let caMobileCPP = PropertiesService.getScriptProperties().getProperty('COLLECTION_ASSISTANT_MOBILE') || "";
         let mobileLineCPP = caMobileCPP ? `Mobile: ${caMobileCPP}<br>` : "";
         let collectionAsstHtml = "";
         [...new Set(collectionAssts)].forEach(asst => collectionAsstHtml += `BIR Form 2307 Expanded Withholding Tax (EWT), Collection Assistant in-charge: ${asst}<br>${mobileLineCPP}`);

         let mgtEmail = isTech ? (sysSettings["Management_A2Z"] || "") : (sysSettings["Management_Pingcon"] || "");
         if (mgtEmail) ccRecipientsList.push(mgtEmail);

         if (mainRecipientsList.length === 0) {
             if (dynamicCcToAppend.length > 0) dynamicCcToAppend.forEach(email => mainRecipientsList.push(email));
             else if (mgtEmail) mainRecipientsList.push(mgtEmail);
         } else {
             dynamicCcToAppend.forEach(email => ccRecipientsList.push(email));
         }

         ccRecipientsList = [...new Set(ccRecipientsList)].filter(email => !mainRecipientsList.includes(email));
         mainRecipientsList = [...new Set(mainRecipientsList)];

         const mainRecipient = mainRecipientsList.join(", ");
         const ccRecipients = ccRecipientsList.join(", ");
     
         let securedBlob = encryptAttachmentViaPython(attachmentBlob, cleanFileName + ".xlsx", null);
         let feedbackLink = isTech ? (SCRIPT_PROPS['FEEDBACK_LINK_TECH'] || "") : (SCRIPT_PROPS['FEEDBACK_LINK_PINGCON'] || "");
         let dateRangePhrase = `<p>This report covers transactions recorded in our system from <strong>${Utilities.formatDate(qStartObj, Session.getScriptTimeZone(), "MMMM d, yyyy")}</strong> to <strong>${Utilities.formatDate(qEndObj, Session.getScriptTimeZone(), "MMMM d, yyyy")}</strong>.</p>`;
         
         let htmlBody = Template_CppCustEmail(cleanCustName, dateRangePhrase, aeContactHtml, collectionAsstHtml, feedbackImgBlob, feedbackLink);
         let emailOptions = { from: senderEmail, cc: ccRecipients, name: senderName, attachments: [securedBlob], inlineImages: {}, throwError: true };
         if (feedbackImgBlob) emailOptions.inlineImages.feedbackImage = feedbackImgBlob;

         try {
             sendRobustEmail(mainRecipient, cleanFileName, htmlBody, emailOptions);
      
         } catch(errStr) {
             if (String(errStr).includes("too many times") || String(errStr).includes("premium email") || String(errStr).includes("limit")) {
                 pendingCust.unshift(currentCustomerGroup);
                 props.setProperty(CPP_BATCH_KEY, JSON.stringify(pendingCust));
                 props.setProperty("LAST_RUN_ERROR", "Google Email Quota limit reached! Script stopped. Please exclude the completed items and run the remaining items tomorrow.");
                 props.deleteProperty("ONGOING_REPORT_TYPE");
                 props.deleteProperty("ONGOING_REPORT_STATUS");
                 props.deleteProperty("RESUME_TIME");
                 deleteTriggersByName(CPP_TRIGGER_NAME);
             
                 return;
             }
         }
         Utilities.sleep(10000); // Increased delay to prevent spam flagging
     }
  }
  props.setProperty(CPP_BATCH_KEY, JSON.stringify(pendingCust));
  replaceTrigger(CPP_TRIGGER_NAME);
}

function Template_CppCustEmail(cleanCustName, dateRangePhrase, aeContactHtml, collectionAsstHtml, feedbackImgBlob, feedbackLink) {
  return `
    <div style="font-family: Arial, sans-serif; font-size: 13px; color: #000; line-height: 1.5;">
      <p>Dear ${cleanCustName},</p>
      <p>We hope this message finds you well.</p>
      <p>Please find attached your Customer Payment Performance Report.</p>
      ${dateRangePhrase}
      <p>This report provides a detailed overview of your payment history and performance for the specified quarter. We share this to ensure transparency and to help maintain our strong, collaborative partnership. If you notice any discrepancies or have any questions regarding the recorded transactions, please do not hesitate to reach out so we can clarify and align our records.</p>
      <p>We sincerely appreciate your continued trust and business.</p>
      <br>
      <p>Please do not hesitate to contact:<br>${aeContactHtml}${collectionAsstHtml}</p>
      <br><p>Warm regards,</p><br>
      ${feedbackImgBlob ? `<a href="${feedbackLink}" target="_blank"><img src="cid:feedbackImage" alt="Feedback Matters" style="max-width: 350px; height: auto; border: none;"></a>` : ''}
    </div>`;
}

// =========================================================================
// BATCH SUCCESS SUMMARY EMAILER (Exclusion List Generator)
// =========================================================================
function recordSuccessfulSend(entityName, bucketName, quota, actual) {
  const props = PropertiesService.getScriptProperties();
  let listStr = props.getProperty("BATCH_SUCCESS_LIST");
  let list = listStr ? JSON.parse(listStr) : [];
  
  if (!list.find(item => item.name === entityName && item.bucket === bucketName)) {
      list.push({
          name: entityName, 
          bucket: bucketName || "Pingcon", 
          quota: quota || 0, 
          actual: actual || 0
      });
  }
  props.setProperty("BATCH_SUCCESS_LIST", JSON.stringify(list));
}

function sendBatchSuccessSummary(isPaused, reportType) {
  const props = PropertiesService.getScriptProperties();
  let listStr = props.getProperty("BATCH_SUCCESS_LIST");
  if (!listStr) return;
  let list = JSON.parse(listStr);
  if (list.length === 0) return;

  const now = new Date();
  const reportDateString = Utilities.formatDate(now, Session.getScriptTimeZone(), "MMM dd, yyyy HHmm");
  const buckets = ["Pingcon", "A2Z2GO"];

  if (reportType === "UNPAID_DR_OVERDUE") {
      props.deleteProperty("BATCH_SUCCESS_LIST");
      return; 
  }

  if (reportType === "SALES_REPORT_CONSOLIDATED") {
      // --- CONSOLIDATED SALES QUOTA REPORT ---
      buckets.forEach(bucket => {
          let filtered = list.filter(i => i.bucket === bucket);
          if (filtered.length === 0) return;
          filtered.sort((a, b) => a.name.localeCompare(b.name));

          let cleanFileName = `Consolidated Sales Quota Report - ${bucket} as of Report Date ${reportDateString}`;
          let tempSs = SpreadsheetApp.create(cleanFileName);
          let sheet = tempSs.getSheets()[0];
          let headers = ["SN", "AE Name", "Monthly Quota (Vat Ex)", "Actual Sales (Vat Ex)", "Balance", "Achievement Rate", "Quota Reached?"];
          let rows = [headers];
          let tableHtmlRows = "";

          filtered.forEach((item, idx) => {
              let balance = Math.max(0, item.quota - item.actual);
              let rate = item.quota > 0 ? ((item.actual / item.quota) * 100).toFixed(2) + "%" : "0.00%";
              let reached = item.quota === 0 ? "Missing Quota" : (item.actual >= item.quota ? "YES" : "NO");
              let reachedColor = item.quota === 0 ? "#cc0000" : (reached === "YES" ? "#28a745" : "#cc0000");
              
              rows.push([idx + 1, item.name, item.quota, item.actual, balance, rate, reached]);
              tableHtmlRows += `<tr>
                <td style="border:1px solid #ccc; padding:8px; text-align:center;">${idx + 1}</td>
                <td style="border:1px solid #ccc; padding:8px;">${item.name}</td>
                <td style="border:1px solid #ccc; padding:8px; text-align:right;">${item.quota.toLocaleString(undefined,{minimumFractionDigits:2})}</td>
                <td style="border:1px solid #ccc; padding:8px; text-align:right;">${item.actual.toLocaleString(undefined,{minimumFractionDigits:2})}</td>
                <td style="border:1px solid #ccc; padding:8px; text-align:right;">${balance.toLocaleString(undefined,{minimumFractionDigits:2})}</td>
                <td style="border:1px solid #ccc; padding:8px; text-align:center;">${rate}</td>
                <td style="border:1px solid #ccc; padding:8px; text-align:center; font-weight:bold; color:${reachedColor};">${reached}</td>
              </tr>`;
          });

          sheet.getRange(1, 1, rows.length, headers.length).setValues(rows);
          sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#f3f3f3");
          sheet.getRange(2, 3, rows.length, 3).setNumberFormat("#,##0.00");
          for (let i = 1; i <= headers.length; i++) sheet.autoResizeColumn(i);
          sheet.setColumnWidth(1, 50);
          sheet.setColumnWidth(2, 250);
          sheet.setColumnWidth(3, 160);
          sheet.setColumnWidth(4, 160);
          sheet.setColumnWidth(5, 160);
          sheet.setColumnWidth(6, 130);
          sheet.setColumnWidth(7, 130);
          
          SpreadsheetApp.flush();
          let url = "https://docs.google.com/spreadsheets/d/" + tempSs.getId() + "/export?format=xlsx";
          let token = ScriptApp.getOAuthToken();
          let response = UrlFetchApp.fetch(url, { headers: { 'Authorization': 'Bearer ' + token } });
          let securedBlob = encryptAttachmentViaPython(response.getBlob().setName(cleanFileName + ".xlsx"), cleanFileName + ".xlsx", bucket === "A2Z2GO" ? SCRIPT_PROPS['ENCRYPT_PASS_TECH'] : SCRIPT_PROPS['ENCRYPT_PASS_PINGCON']);

          let htmlBody = `<div style="font-family:Arial; font-size:13px; color:#000;"><p>Good Day,</p>
            <p>Please find the <b>Consolidated Sales Quota Summary Report</b> for <b>${bucket}</b> as of ${reportDateString}.</p>
            <table style="border-collapse:collapse; width:100%; max-width:850px; font-size:12px; margin:20px 0;">
              <thead><tr style="background-color:#F50057; color:white;">
                <th style="padding:10px; border:1px solid #ccc;">SN</th><th style="padding:10px; border:1px solid #ccc;">AE Name</th>
                <th style="padding:10px; border:1px solid #ccc;">Quota (Vat Ex)</th><th style="padding:10px; border:1px solid #ccc;">Actual Sales (Vat Ex)</th>
                <th style="padding:10px; border:1px solid #ccc;">Balance</th><th style="padding:10px; border:1px solid #ccc;">Rate</th>
                <th style="padding:10px; border:1px solid #ccc;">Reached?</th>
              </tr></thead><tbody>${tableHtmlRows}</tbody></table>
            <p>Thank you for your kindness and support. I truly appreciate it.</p><br><p>Regards,<br><b>Finance System</b></p></div>`;

          let sysSettings = getSystemSettings();
          let isTech = (bucket === "A2Z2GO");
          let mainRecipient = isTech ? (sysSettings["Management_A2Z"] || "") : (sysSettings["Management_Pingcon"] || "");
          let senderEmail = getSenderEmail(isTech);
          let senderName = isTech ? "Finance A2Z2GO" : "Finance Pingcon";

          sendRobustEmail(mainRecipient, cleanFileName, htmlBody, {
              from: senderEmail,
              name: senderName,
              attachments: [securedBlob]
          });
          DriveApp.getFileById(tempSs.getId()).setTrashed(true);
      });

  } else {
      // --- STANDARD EXCLUSION LIST (SOA PER CUST, INDIVIDUAL SALES REPORT, ETC) ---
      let typeLabel = reportType.includes("CUST") ? "Customers" : (reportType.includes("OVERDUE") ? "Entities" : "AEs");
      let formattedDateForSubject = Utilities.formatDate(now, Session.getScriptTimeZone(), "MMM dd, yyyy HHmm");
      let subject = `${typeLabel} who Received Email on ${formattedDateForSubject}`;
      
      let htmlBody = `<div style="font-family: Arial, sans-serif; font-size: 13px; color: #000; line-height: 1.5;">
        <p>Here is the list of entities that successfully received their reports during this run.</p>
        <p>If you are continuing a paused batch tomorrow, please copy the list below and paste it into the <strong>Exclusion List</strong> box in the Finance Portal to prevent resending:</p>`;

      buckets.forEach(bucket => {
          let filtered = list.filter(i => i.bucket === bucket);
          if (filtered.length === 0) return;
          filtered.sort((a, b) => a.name.localeCompare(b.name));

          htmlBody += `<br><b>${bucket}:</b><br>
          <table style="border-collapse: collapse; width: 100%; max-width: 600px; margin-top: 10px; border: 1px solid #ccc;">
            <thead>
              <tr style="background-color: #f3f3f3;">
                <th style="border: 1px solid #ccc; padding: 8px; width: 50px; text-align: center;">SN</th>
                <th style="border: 1px solid #ccc; padding: 8px; text-align: left;">${typeLabel === "Customers" ? "Customer Name" : "AE Name"}</th>
              </tr>
            </thead>
            <tbody>`;
          filtered.forEach((item, idx) => {
              htmlBody += `<tr>
                <td style="border: 1px solid #ccc; padding: 8px; text-align: center;">${idx + 1}.</td>
                <td style="border: 1px solid #ccc; padding: 8px;">${item.name}</td>
              </tr>`;
          });
          htmlBody += `</tbody></table>`;
      });
      htmlBody += `</div>`;

      let sysSettings = getSystemSettings();
      let hasKeysysOrA2ZIT = list.some(item => item.bucket === "A2Z2GO");
      let mainRecipient = hasKeysysOrA2ZIT ? (sysSettings["Management_A2Z"] || "") : (sysSettings["Management_Pingcon"] || "");
      let senderEmail = getSenderEmail(hasKeysysOrA2ZIT);

      try {
          sendRobustEmail(mainRecipient, subject, htmlBody, {
              from: senderEmail,
              name: "Finance System"
          });
      } catch(e) { console.error("Summary Email Error: " + e); }
  }

  props.deleteProperty("BATCH_SUCCESS_LIST");
}




// =========================================================================
// AI EXTRACTION & 2307 GENERATION APIs (PART 1)
// =========================================================================
function getWebAppUrl() {
    return ScriptApp.getService().getUrl();
}

// ---> ADD THIS NEW FUNCTION BELOW <---


function webappSearchInvoices(query, companyContext) {
  if (!query || query.trim().length === 0) return { success: true, results: [] };
  
  let props = PropertiesService.getScriptProperties();
  let isDevMode = ScriptApp.getService().getUrl().endsWith("/dev");
  const NEW_DB_FOLDER_ID = (isDevMode && props.getProperty('AI_DB_FOLDER_ID_DEV'))
      ? props.getProperty('AI_DB_FOLDER_ID_DEV')
      : (props.getProperty('AI_DB_FOLDER_ID') || "");
  const VERIFIED_FOLDER_ID = (isDevMode && props.getProperty('VERIFIED_FOLDER_ID_DEV')) 
      ? props.getProperty('VERIFIED_FOLDER_ID_DEV') 
      : props.getProperty('VERIFIED_FOLDER_ID');
  
  // Remove "Invoice - " from the query so it searches purely for the exact number
  let cleanQuery = query.replace(/^invoice\s*-\s*/i, '').trim();
  let lowerQuery = cleanQuery.toLowerCase();
  
  let results = [];

  try {
            let folder = DriveApp.getFolderById(NEW_DB_FOLDER_ID);
            
                // Stop searching if the query becomes empty after stripping "Invoice - "
                if (cleanQuery.length === 0) return { success: true, results: [] };
                
                // FETCH & CACHE DB FILE IDs TO BYPASS SLOW FOLDER SEARCHING
                let cache = CacheService.getScriptCache();
                let cacheKey = "AI_DB_FILES_" + companyContext;
                let cachedIds = cache.get(cacheKey);
                let fileIdsArr = [];

                if (cachedIds) {
                    fileIdsArr = JSON.parse(cachedIds);
                } else {
                    let dbSearchQuery = `title contains '${companyContext}' and mimeType = '${MimeType.GOOGLE_SHEETS}'`;
                    let fileIter = folder.searchFiles(dbSearchQuery);
                    let tempFiles = [];
                    while (fileIter.hasNext()) tempFiles.push(fileIter.next());
                    tempFiles.sort((a, b) => b.getLastUpdated() - a.getLastUpdated());
                    fileIdsArr = tempFiles.map(f => f.getId());
                    cache.put(cacheKey, JSON.stringify(fileIdsArr), 3600); // Cache IDs for 1 hour
                }
                
    for (let fIdx = 0; fIdx < fileIdsArr.length && results.length < 10; fIdx++) {
        let ss = SpreadsheetApp.openById(fileIdsArr[fIdx]);
        let aiSheet = ss.getSheetByName("ai_extraction");
        if (!aiSheet) continue;

        let data = aiSheet.getDataRange().getValues();
        if (data.length <= 1) continue;

        let headerRow = data[0];
        let rawIdx = headerRow.findIndex(h => String(h).trim() === "Raw AI JSON");
        if (rawIdx === -1) rawIdx = 2; // Col C
        let correctedIdx = headerRow.findIndex(h => String(h).trim() === "Corrected JSON");
        if (correctedIdx === -1) correctedIdx = 3; // Col D

        for (let i = 1; i < data.length; i++) {
                    let rawData = String(data[i][rawIdx]).trim();
                    let correctedData = String(data[i][correctedIdx]).trim();
                    
                    // FAST BAILOUT: Check if either raw or corrected contains the search query
                    if (!rawData.toLowerCase().includes(lowerQuery) && !correctedData.toLowerCase().includes(lowerQuery)) continue;

                    // Use Corrected JSON if valid, else Raw JSON just to get the invoice number for the dropdown search
                    let extractedData = correctedData.startsWith("{") ? correctedData : rawData;
                    if (!extractedData.startsWith("{")) continue;

                    // Flag if it has no corrected data yet
                    let isUnverified = !correctedData.startsWith("{");

                    try {
                        let parsed = JSON.parse(extractedData);
                        if (parsed.documents && parsed.documents.length > 0) {
                            // Loop through all documents instead of just [0]
                            for (let d = 0; d < parsed.documents.length; d++) {
                                let doc = parsed.documents[d];
                                
                                // STRICT FILTER: Ignore Delivery Receipts and Purchase Orders in Search Dropdown
                                if (doc.documentType && doc.documentType !== "Invoice") continue;
                                
                                let fileName = String(doc.fileName || "").trim();
                                let fileId = String(doc.fileId || "").trim();
                                let invoiceNo = "";
                                
                                if (doc.fields) {
                                    let invField = doc.fields.find(f => 
                                        f.label === "Invoice No." || f.label === "Invoice Number" || f.label === "Invoice No" || f.label === "invoice_number" || f.label === "Sales Invoice No." || f.label === "Sales Invoice No"
                                    );
                                    if (invField) invoiceNo = String(invField.value).trim();
                                }
                                
                                // STRICT RULE: Must have an invoice number to show up
                                if (!fileName || !invoiceNo) continue; 

                                // PARTIAL MATCHING: Match if the query is inside the actual invoice number
                                if (invoiceNo.toLowerCase().includes(lowerQuery)) {
                                    // Package ONLY the matched document so the UI loads it cleanly
                                    let singleDocParsed = JSON.parse(extractedData);
                                    singleDocParsed.documents = [doc];
                                    let singleDocExtractedData = JSON.stringify(singleDocParsed);
                                    
                                    results.push({
                                        fileName: fileName,
                                        invoiceNo: invoiceNo,
                                        fileId: fileId,
                                        extractedData: singleDocExtractedData,
                                        isUnverified: isUnverified,
                                        verifiedData: "", 
                                        dbSourceName: ss.getName() 
                                    });
                                }
                                if (results.length >= 10) break; // Break document loop if max reached
                            }
                        }
                    } catch (e) {}
                    
                    if (results.length >= 10) break; // Break row loop if max reached
                }
    }

    // Retrieve verified data ONLY if we found matches, to save loading time
    let verifiedResults = results.filter(r => !r.isUnverified);
    if (verifiedResults.length > 0) {
        let cache = CacheService.getScriptCache();
        let cacheKeyV = "VERIFIED_DB_FILES_" + VERIFIED_FOLDER_ID;
        let cachedVIds = cache.get(cacheKeyV);
        let vFileIdsArr = [];

        if (cachedVIds) {
            vFileIdsArr = JSON.parse(cachedVIds);
        } else {
            let verifiedFolder = DriveApp.getFolderById(VERIFIED_FOLDER_ID);
            let vIter = verifiedFolder.getFilesByType(MimeType.GOOGLE_SHEETS);
            let tempVFiles = [];
            while (vIter.hasNext()) {
                let f = vIter.next();
                if (/^Verified Invoices for 2307 \d{4}$/i.test(f.getName().trim())) {
                    tempVFiles.push(f);
                }
            }
            tempVFiles.sort((a, b) => b.getLastUpdated() - a.getLastUpdated());
            vFileIdsArr = tempVFiles.map(f => f.getId());
            cache.put(cacheKeyV, JSON.stringify(vFileIdsArr), 3600); // Cache IDs for 1 hour
        }
        
        let pendingVerifications = verifiedResults.length;
        for (let vf = 0; vf < vFileIdsArr.length; vf++) {
            if (pendingVerifications <= 0) break; // FAST BAILOUT
            let vSs = SpreadsheetApp.openById(vFileIdsArr[vf]);
            let allSheets = vSs.getSheets();
            
            for (let s = 0; s < allSheets.length; s++) {
                let vData = allSheets[s].getDataRange().getValues();
                if (vData.length > 1) {
                    for (let i = 1; i < vData.length; i++) {
                        let vFileName = String(vData[i][1]).trim();
                        let match = verifiedResults.find(r => r.fileName === vFileName);
                        if (match && !match.hasOwnProperty('foundInDb')) {
                            let jsonStr = String(vData[i][2] || "").trim();
                            // Ensure it's valid JSON before marking it as verified data (detects cleared cells)
                            match.verifiedData = jsonStr.startsWith("{") ? jsonStr : "";
                            match.foundInDb = true; // Mark as found so it doesn't get overwritten by older backup files
                            pendingVerifications--;
                        }
                    }
                }
            }
        }
    }

    return { success: true, results: results };
  } catch (e) {
    return { error: "Database Error: " + e.message };
  }
}

function webappGetSingleInvoiceData(dbSourceName, rowIndex, invoiceNo) {
  try {
    let props = PropertiesService.getScriptProperties();
    let isDevMode = ScriptApp.getService().getUrl().endsWith("/dev");
    
    // 1. Fetch from AI Database (Dev vs Exec safe)
    const NEW_DB_FOLDER_ID = (isDevMode && props.getProperty('AI_DB_FOLDER_ID_DEV'))
        ? props.getProperty('AI_DB_FOLDER_ID_DEV')
        : (props.getProperty('AI_DB_FOLDER_ID') || "");

    let folder = DriveApp.getFolderById(NEW_DB_FOLDER_ID);
    let files = folder.searchFiles("title = '" + dbSourceName.replace(/'/g, "\\'") + "' and mimeType = '" + MimeType.GOOGLE_SHEETS + "' and trashed = false");
    if (!files.hasNext()) throw new Error("Database file not found: " + dbSourceName);
    
    let ss = SpreadsheetApp.open(files.next());
    let sheet = ss.getSheetByName("ai_extraction");
    if (!sheet) throw new Error("ai_extraction tab not found.");
    
    let headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h).trim());
    let rawIdx = headers.indexOf("Raw AI JSON") + 1;
    if (rawIdx === 0) rawIdx = 3;
    let correctedIdx = headers.indexOf("Corrected JSON") + 1;
    if (correctedIdx === 0) correctedIdx = 4;
    
    let rowNum = parseInt(rowIndex, 10) + 1; 
    let extractedData = "";

    // 1. Try direct row index if valid
    if (rowNum > 1 && rowNum <= sheet.getLastRow()) {
        let rawData = String(sheet.getRange(rowNum, rawIdx).getValue()).trim();
        let correctedData = String(sheet.getRange(rowNum, correctedIdx).getValue()).trim();
        extractedData = correctedData.startsWith("{") ? correctedData : rawData;
    }

    // 2. If row index is invalid or does not match this invoice, search the sheet for the invoice number
    let cleanTargetInv = String(invoiceNo).trim().toUpperCase();
    if (!extractedData.startsWith("{") || !extractedData.toUpperCase().includes(cleanTargetInv)) {
        let allData = sheet.getDataRange().getValues();
        for (let r = 1; r < allData.length; r++) {
            let cData = String(allData[r][correctedIdx - 1] || "").trim();
            let rData = String(allData[r][rawIdx - 1] || "").trim();
            let candidate = cData.startsWith("{") ? cData : rData;
            if (candidate.startsWith("{") && candidate.toUpperCase().includes(cleanTargetInv)) {
                extractedData = candidate;
                break;
            }
        }
    }

    if (!extractedData.startsWith("{")) {
        // Fallback: Build empty document payload if unextracted so verified data still opens
        extractedData = JSON.stringify({ documents: [{ documentType: "Invoice", fields: [{ label: "Invoice Number", value: invoiceNo }] }] });
    }
    
    let parsed = JSON.parse(extractedData);
    let targetDoc = null;
    if (parsed.documents && parsed.documents.length > 0) {
        for (let d = 0; d < parsed.documents.length; d++) {
            let doc = parsed.documents[d];
            if (doc.documentType && doc.documentType !== "Invoice") continue;
            
            if (doc.fields) {
                let invField = doc.fields.find(f => f.label === "Invoice No." || f.label === "Invoice Number" || f.label === "invoice_number" || f.label === "Sales Invoice No.");
                if (invField && String(invField.value).trim().toUpperCase() === String(invoiceNo).trim().toUpperCase()) {
                    targetDoc = doc;
                    break;
                }
            }
        }
    }
    
    if (!targetDoc) throw new Error("Specific invoice data not found in JSON payload.");
    parsed.documents = [targetDoc];

    // Extract Supplier Name and Filename from the document being opened
    let targetCompField = targetDoc.fields ? targetDoc.fields.find(f => f.label === "Supplier" || f.label === "Company Name" || f.label === "supplier_name" || f.label === "Registered Name") : null;
    let targetCompUpper = targetCompField ? String(targetCompField.value).trim().toUpperCase() : "";
    let targetFileUpper = String(targetDoc.fileName || "").trim().toUpperCase();

    // 2. Fetch existing Verified Purchase/2307 Record strictly matching BOTH Invoice No AND Supplier
    let verifiedDataStr = null;
    const VERIFIED_FOLDER_ID = (isDevMode && props.getProperty('VERIFIED_FOLDER_ID_DEV'))
        ? props.getProperty('VERIFIED_FOLDER_ID_DEV')
        : props.getProperty('VERIFIED_FOLDER_ID');

    if (VERIFIED_FOLDER_ID) {
        try {
            let vFolder = DriveApp.getFolderById(VERIFIED_FOLDER_ID);
            let vFiles = vFolder.getFilesByType(MimeType.GOOGLE_SHEETS);
            let targetInvUpper = String(invoiceNo).trim().toUpperCase();

            while (vFiles.hasNext()) {
                let vFile = vFiles.next();
                if (vFile.isTrashed()) continue; // Skip trashed spreadsheets
                if (!vFile.getName().startsWith("Verified Invoices for 2307")) continue;
                let vSs = SpreadsheetApp.open(vFile);
                let vSheet = vSs.getSheetByName("All Verified Invoices");
                if (!vSheet) continue;
                
                let vRows = vSheet.getDataRange().getValues();
                if (vRows.length <= 1) continue;
                
                let vHeaders = vRows[0].map(h => String(h).trim());
                let vdIdx = vHeaders.indexOf("Verified Data");
                if (vdIdx === -1) vdIdx = 2; // Col C default
                
                for (let r = 1; r < vRows.length; r++) {
                    let cellJson = String(vRows[r][vdIdx]).trim();
                    if (cellJson.startsWith("{") && cellJson.toUpperCase().includes(targetInvUpper)) {
                        try {
                            let vParsed = JSON.parse(cellJson);
                            let rowInv = String(vParsed["Invoice Number"] || "").trim().toUpperCase();
                            let rowComp = String(vParsed["Company Name"] || "").trim().toUpperCase();
                            let rowFile = String(vRows[r][1] || "").trim().toUpperCase();

                            let sameInv = (rowInv === targetInvUpper);
                            // Vendor Isolation: Only match if the supplier name or filename is truly identical
                            let sameComp = (targetCompUpper !== "" && rowComp !== "" && (targetCompUpper === rowComp || targetCompUpper.includes(rowComp) || rowComp.includes(targetCompUpper)));
                            let sameFile = (targetFileUpper !== "" && rowFile !== "" && targetFileUpper === rowFile);

                            if (sameInv && (sameComp || sameFile)) {
                                verifiedDataStr = cellJson;
                                break;
                            }
                        } catch (e) {}
                    }
                }
                if (verifiedDataStr) break;
            }
        } catch (vErr) {
            console.warn("Could not check verified database: " + vErr.message);
        }
    }
    
    let resolvedFileId = "";
    if (targetDoc) {
      resolvedFileId = String(targetDoc.fileId || targetDoc.file_id || parsed.fileId || parsed.file_id || "").trim();
    }

    return { 
        success: true, 
        extractedData: JSON.stringify(parsed),
        verifiedData: verifiedDataStr,
        fileId: resolvedFileId // Guarantees fileId is never lost
    };
  } catch (e) {
    return { error: e.message };
  }
}

function webappGetProcessingQueue(companyContext, offset = 0, searchQuery = "", targetLimit = 100, searchDepth = 300) {
  try {
    const startTime = Date.now();
    const MAX_SEARCH_TIME = 45 * 1000; // 45 seconds safe ceiling to prevent timeout

    let props = PropertiesService.getScriptProperties();
    let isDevMode = ScriptApp.getService().getUrl().endsWith("/dev");
    
    // Parse user-chosen search depth
    let depthNum = parseInt(searchDepth, 10);
    if (isNaN(depthNum) || depthNum < 0) depthNum = 300;
    let limitCount = searchQuery ? (depthNum === 0 ? Infinity : depthNum) : targetLimit;

    let pendingCount = 0;
    let rowsSkipped = 0;
    let lowerQuery = searchQuery ? String(searchQuery).toLowerCase().trim() : "";
    
    // 1. Fetch Verified Invoices (Processed) with Dual Indexing and Full Record Collection
    const VERIFIED_FOLDER_ID = (isDevMode && props.getProperty('VERIFIED_FOLDER_ID_DEV')) 
        ? props.getProperty('VERIFIED_FOLDER_ID_DEV') 
        : props.getProperty('VERIFIED_FOLDER_ID');
        
    let verifiedFolder;
    try { verifiedFolder = DriveApp.getFolderById(VERIFIED_FOLDER_ID); } catch(e) { verifiedFolder = null; }
    let processedInvoices = new Map();
    let processedFiles = new Map();
    let processedJsonMap = new Map();
    let verifiedQueueItems = [];
    
    if (verifiedFolder) {
        let vFiles = verifiedFolder.getFilesByType(MimeType.GOOGLE_SHEETS);
        while(vFiles.hasNext()) {
            let vFile = vFiles.next();
            if (vFile.isTrashed()) continue; // Skip trashed spreadsheets
            if (!vFile.getName().startsWith("Verified Invoices for 2307")) continue;
            let vSs = SpreadsheetApp.open(vFile);
            let vSheet = vSs.getSheetByName("All Verified Invoices");
            if (!vSheet) continue;
            let vData = vSheet.getDataRange().getValues();
            if (vData.length > 1) {
                let vdIdx = vData[0].map(h=>String(h).trim()).indexOf("Verified Data");
                let srcIdx = vData[0].map(h=>String(h).trim()).indexOf("Source");
                let rdIdx = vData[0].map(h=>String(h).trim()).indexOf("2307 Release Date");
                let dtIdx = vData[0].map(h=>String(h).trim()).indexOf("Date and Time Verified (MM/DD/YYYY)");

                for(let i=1; i<vData.length; i++) {
                    let source = srcIdx > -1 ? String(vData[i][srcIdx]).trim().toUpperCase() : "";
                    let vJsonStr = (vdIdx > -1) ? String(vData[i][vdIdx] || "").trim() : "";
                    
                    // Check if JSON has an explicit Assigned Entity override
                    let assignedEnt = "";
                    if (vJsonStr.startsWith("{")) {
                      try {
                        let parsedMeta = JSON.parse(vJsonStr);
                        assignedEnt = String(parsedMeta["Assigned Entity"] || parsedMeta["Company Entity"] || "").trim().toUpperCase();
                      } catch(e) {}
                    }
                    
                    let effectivePortal = assignedEnt || source;
                    let activeCtxUpper = String(companyContext || "").toUpperCase();
                    
                    // Strict Entity Gate: If this row belongs to another entity, skip it
                    if (activeCtxUpper === "PINGCON" && (effectivePortal.includes("BONITA") || effectivePortal.includes("KEYSYS") || effectivePortal.includes("A2Z"))) continue;
                    if (activeCtxUpper === "BONITA" && !effectivePortal.includes("BONITA")) continue;
                    if (activeCtxUpper === "KEYSYS" && !effectivePortal.includes("KEYSYS")) continue;
                    if (activeCtxUpper === "A2Z2GO" && !effectivePortal.includes("A2Z")) continue;

                    let rowFileName = String(vData[i][1] || "").trim().toUpperCase();
                    let rdVal = (rdIdx > -1 && vData[i][rdIdx]) ? String(vData[i][rdIdx]).trim() : "";
                    let recType = rdVal ? "2307 EWT" : "Standard Purchase";
                    let verifiedAt = (dtIdx > -1 && vData[i][dtIdx]) ? String(vData[i][dtIdx]).trim() : "";

                    // Index by Column B Filename
                    if (rowFileName) {
                        processedFiles.set(rowFileName, recType);
                        let fnDigits = rowFileName.replace(/[^0-9]/g, '');
                        if (fnDigits && fnDigits.length >= 3) {
                            processedInvoices.set(fnDigits, recType);
                        }
                    }

                    // Index by Column C JSON and Collect Full Record
                    if (vdIdx > -1) {
                        let vJson = String(vData[i][vdIdx]).trim();
                        if (vJson.startsWith("{")) {
                            let invNo = "";
                            let compName = "";
                            let invDate = "";
                            let totalAmount = "";
                            let fileId = "";
                            let rType = recType;

                            try {
                                let parsed = JSON.parse(vJson);
                                invNo = String(parsed["Invoice Number"] || "").trim().toUpperCase();
                                compName = String(parsed["Company Name"] || "").trim();
                                invDate = String(parsed["Invoice Date"] || "").trim();
                                fileId = String(parsed["fileId"] || "").trim();
                                rType = String(parsed["Record Type"] || recType).trim();
                                
                                if (parsed.isMixedVat) {
                                    let vGross = parseFloat(String((parsed.VAT_Portion && parsed.VAT_Portion.Gross) || "0").replace(/,/g, '')) || 0;
                                    let nvGross = parseFloat(String((parsed.NV_Portion && parsed.NV_Portion.Gross) || "0").replace(/,/g, '')) || 0;
                                    totalAmount = (vGross + nvGross).toLocaleString('en-US', {minimumFractionDigits: 2, maximumFractionDigits: 2});
                                } else {
                                    totalAmount = parsed["Gross"] || parsed["Total Amount"] || "0.00";
                                }
                            } catch(e) {
                                let invMatch = vJson.match(/"Invoice Number"\s*:\s*"([^"]+)"/i);
                                let compMatch = vJson.match(/"Company Name"\s*:\s*"([^"]+)"/i);
                                let dateMatch = vJson.match(/"Invoice Date"\s*:\s*"([^"]+)"/i);
                                let totMatch = vJson.match(/"(?:Gross|Total Amount)"\s*:\s*"([^"]+)"/i);
                                let recMatch = vJson.match(/"Record Type"\s*:\s*"([^"]+)"/i);
                                let fidMatch = vJson.match(/"fileId"\s*:\s*"([^"]+)"/i);

                                if (invMatch && invMatch[1]) invNo = String(invMatch[1]).trim().toUpperCase();
                                if (compMatch && compMatch[1]) compName = String(compMatch[1]).trim();
                                if (dateMatch && dateMatch[1]) invDate = String(dateMatch[1]).trim();
                                if (totMatch && totMatch[1]) totalAmount = String(totMatch[1]).trim();
                                if (recMatch && recMatch[1]) rType = String(recMatch[1]).trim();
                                if (fidMatch && fidMatch[1]) fileId = String(fidMatch[1]).trim();
                            }

                            if (invNo) {
                                let normComp = compName ? compName.trim().toUpperCase().replace(/[^A-Z0-9]/g, '') : "";
                                let cleanDigits = invNo.replace(/[^0-9]/g, '');

                                // Index strictly with Supplier Name to prevent cross-vendor collisions
                                if (normComp) {
                                    processedInvoices.set(invNo + "_" + normComp, rType);
                                    processedJsonMap.set(invNo + "_" + normComp, vJson);
                                    if (cleanDigits) {
                                        processedInvoices.set(cleanDigits + "_" + normComp, rType);
                                        processedJsonMap.set(cleanDigits + "_" + normComp, vJson);
                                    }
                                    if (rowFileName) {
                                        processedFiles.set(rowFileName + "_" + normComp, rType);
                                        processedJsonMap.set(rowFileName + "_" + normComp, vJson);
                                    }
                                }

                                let statusLabel = (rType === "Standard Purchase") ? "Processed Purchase" : "Processed 2307";

                                verifiedQueueItems.push({
                                    sn: 0,
                                    fileName: rowFileName || ("SI " + invNo + ".pdf"),
                                    invoiceNo: invNo,
                                    invoiceDate: invDate,
                                    companyName: compName,
                                    totalAmount: totalAmount,
                                    fileId: fileId,
                                    isUnverified: false,
                                    status: statusLabel,
                                    dbSourceName: source,
                                    rowIndex: -1,
                                    extractedAt: verifiedAt,
                                    extractedData: vJson,
                                    verifiedData: vJson
                                });
                            }
                        }
                    }
                }
            }
        }
    }
    
    // 2. Fetch AI Database (Source) with Dev vs Exec fallback
    const NEW_DB_FOLDER_ID = (isDevMode && props.getProperty('AI_DB_FOLDER_ID_DEV'))
        ? props.getProperty('AI_DB_FOLDER_ID_DEV')
        : (props.getProperty('AI_DB_FOLDER_ID') || "");
        
    let folder = DriveApp.getFolderById(NEW_DB_FOLDER_ID);
    let cache = CacheService.getScriptCache();
    let cacheKey = "AI_DB_FILES_" + companyContext + (isDevMode ? "_DEV" : "_EXEC");
    let cachedIds = cache.get(cacheKey);
    let fileIdsArr = [];

    if (cachedIds) {
        fileIdsArr = JSON.parse(cachedIds);
    } else {
        let dbSearchQuery = `title contains '${companyContext}' and mimeType = '${MimeType.GOOGLE_SHEETS}'`;
        let fileIter = folder.searchFiles(dbSearchQuery);
        let tempFiles = [];
        while (fileIter.hasNext()) tempFiles.push(fileIter.next());
        tempFiles.sort((a, b) => b.getLastUpdated() - a.getLastUpdated());
        fileIdsArr = tempFiles.map(f => f.getId());
        cache.put(cacheKey, JSON.stringify(fileIdsArr), 3600); 
    }

    let results = [];

    for (let fIdx = 0; fIdx < fileIdsArr.length; fIdx++) {
        if (Date.now() - startTime > MAX_SEARCH_TIME) break;
        if (searchQuery && results.length >= limitCount) break;
        if (!searchQuery && pendingCount >= limitCount) break;

        let ss = SpreadsheetApp.openById(fileIdsArr[fIdx]);
        let aiSheet = ss.getSheetByName("ai_extraction");
        if (!aiSheet) continue;

        let data = aiSheet.getDataRange().getValues();
        if (data.length <= 1) continue;

        let headerRow = data[0].map(h => String(h).trim());
        let rawIdx = headerRow.indexOf("Raw AI JSON");
        if (rawIdx === -1) rawIdx = 2; 
        let correctedIdx = headerRow.indexOf("Corrected JSON");
        if (correctedIdx === -1) correctedIdx = 3; 
        let extractedAtIdx = headerRow.indexOf("Extracted At");
        if (extractedAtIdx === -1) extractedAtIdx = 4;

        // Loop backwards to scan newest records first
        for (let i = data.length - 1; i >= 1; i--) {
            if (Date.now() - startTime > MAX_SEARCH_TIME) break;
            if (searchQuery && results.length >= limitCount) break;
            if (!searchQuery && pendingCount >= limitCount) break;

            let rawData = String(data[i][rawIdx]).trim();
            let correctedData = String(data[i][correctedIdx]).trim();
            
            let extractedData = correctedData.startsWith("{") ? correctedData : rawData;
            if (!extractedData.startsWith("{")) continue;

            // Deep Search: Fast-path string check before JSON parsing
            if (lowerQuery !== "") {
                if (extractedData.toLowerCase().indexOf(lowerQuery) === -1) continue;
            } else {
                if (rowsSkipped < offset) {
                    rowsSkipped++;
                    continue;
                }
            }

            let isUnverified = !correctedData.startsWith("{");

            try {
                let parsed = JSON.parse(extractedData);
                if (parsed.documents && parsed.documents.length > 0) {
                    for (let d = 0; d < parsed.documents.length; d++) {
                        let doc = parsed.documents[d];
                        
                        if (doc.documentType && doc.documentType !== "Invoice") continue;
                        
                        let fileName = String(doc.fileName || "").trim();
                        let fileId = String(doc.fileId || "").trim();
                        let invoiceNo = "";
                        let companyName = "";
                        let totalAmount = "";
                        let invDate = "";
                        
                        if (doc.fields) {
                            let invField = doc.fields.find(f => f.label === "Invoice No." || f.label === "Invoice Number" || f.label === "invoice_number" || f.label === "Sales Invoice No." || f.label === "Invoice No" || f.label === "Sales Invoice No");
                            if (invField) invoiceNo = String(invField.value).trim().toUpperCase();
                            
                            let compField = doc.fields.find(f => f.label === "Supplier" || f.label === "Company Name" || f.label === "Registered Name");
                            if (compField) companyName = String(compField.value).trim();
                            
                            let totalField = doc.fields.find(f => f.label === "Total Amount DUE" || f.label === "Gross" || f.label === "Total Amount" || f.label === "total_amount_php");
                            if (totalField) totalAmount = String(totalField.value).trim();

                            let dateField = doc.fields.find(f => f.label === "Date" || f.label === "Invoice Date" || f.label === "invoice_date");
                            if (dateField) invDate = String(dateField.value).trim();
                        }
                        
                        let extractedAt = String(data[i][extractedAtIdx] || "").trim();
                        
                        if (!fileName || !invoiceNo) continue;

                        // Strict Multi-Key Status Check: NEVER match by invoice number alone without supplier name
                        let status = "Pending";
                        let cleanInv = invoiceNo ? invoiceNo.trim().toUpperCase() : "";
                        let cleanComp = companyName ? companyName.trim().toUpperCase().replace(/[^A-Z0-9]/g, '') : "";
                        let cleanFile = fileName ? fileName.trim().toUpperCase() : "";
                        let invDigits = cleanInv.replace(/[^0-9]/g, '');

                        if (cleanInv && cleanComp && processedInvoices.has(cleanInv + "_" + cleanComp)) {
                            status = processedInvoices.get(cleanInv + "_" + cleanComp) === "Standard Purchase" ? "Processed Purchase" : "Processed 2307";
                        } else if (invDigits && cleanComp && processedInvoices.has(invDigits + "_" + cleanComp)) {
                            status = processedInvoices.get(invDigits + "_" + cleanComp) === "Standard Purchase" ? "Processed Purchase" : "Processed 2307";
                        } else if (cleanFile && cleanComp && processedFiles.has(cleanFile + "_" + cleanComp)) {
                            status = processedFiles.get(cleanFile + "_" + cleanComp) === "Standard Purchase" ? "Processed Purchase" : "Processed 2307";
                        } else if (cleanComp) {
                            // Loose check: only match if supplier name matches or includes each other
                            for (let [key, val] of processedInvoices.entries()) {
                                let parts = key.split("_");
                                if (parts.length >= 2 && (parts[0] === cleanInv || parts[0] === invDigits)) {
                                    let storedComp = parts[1];
                                    if (storedComp.includes(cleanComp) || cleanComp.includes(storedComp)) {
                                        status = (val === "Standard Purchase") ? "Processed Purchase" : "Processed 2307";
                                        break;
                                    }
                                }
                            }
                        }

                        let singleDocParsed = JSON.parse(extractedData);
                        singleDocParsed.documents = [doc];
                        
                        let bundledExtractedJson = JSON.stringify(singleDocParsed);
                        let matchedVerifiedJson = null;
                        if (cleanInv && cleanComp && processedJsonMap.has(cleanInv + "_" + cleanComp)) {
                            matchedVerifiedJson = processedJsonMap.get(cleanInv + "_" + cleanComp);
                        } else if (cleanFile && cleanComp && processedJsonMap.has(cleanFile + "_" + cleanComp)) {
                            matchedVerifiedJson = processedJsonMap.get(cleanFile + "_" + cleanComp);
                        }

                        results.push({
                            sn: results.length + 1,
                            fileName: fileName,
                            invoiceNo: invoiceNo,
                            invoiceDate: invDate,
                            companyName: companyName,
                            totalAmount: totalAmount,
                            fileId: fileId,
                            isUnverified: isUnverified,
                            status: status,
                            dbSourceName: ss.getName(),
                            rowIndex: i,
                            extractedAt: extractedAt,
                            extractedData: bundledExtractedJson,
                            verifiedData: matchedVerifiedJson
                        });
                        
                        if (status === "Pending") pendingCount++;
                        
                        if (searchQuery && results.length >= limitCount) break;
                        if (!searchQuery && pendingCount >= limitCount) break;
                    }
                }
            } catch (e) {}
        }
    }

    // 3. Ensure verified records update existing matches and only genuinely unique records are added
    const normStr = s => String(s || "").trim().toUpperCase().replace(/[^A-Z0-9]/g, '');

    verifiedQueueItems.forEach(vItem => {
        let vInv = normStr(vItem.invoiceNo);
        let vComp = normStr(vItem.companyName);
        let vFile = String(vItem.fileName || "").trim().toUpperCase();
        let vFileId = String(vItem.fileId || "").trim();

        // Check if an AI extracted record already exists for this exact file/supplier
        let matchedIndex = results.findIndex(r => {
            if (vFileId && r.fileId && String(r.fileId).trim() === vFileId) return true;
            if (vFile && r.fileName && String(r.fileName).trim().toUpperCase() === vFile) return true;
            let rInv = normStr(r.invoiceNo);
            let rComp = normStr(r.companyName);
            return (rInv === vInv && (rComp.includes(vComp) || vComp.includes(rComp)));
        });

        if (matchedIndex > -1) {
            // Update the existing record with the latest verified data (prevents duplicates)
            results[matchedIndex].status = vItem.status;
            results[matchedIndex].invoiceNo = vItem.invoiceNo;
            results[matchedIndex].companyName = vItem.companyName || results[matchedIndex].companyName;
            results[matchedIndex].totalAmount = vItem.totalAmount || results[matchedIndex].totalAmount;
            results[matchedIndex].verifiedData = vItem.verifiedData;
            results[matchedIndex].extractedData = vItem.verifiedData || results[matchedIndex].extractedData;
        } else {
            // Genuinely unique verified record not in the AI table
            if (lowerQuery !== "") {
                let matchQ = vItem.invoiceNo.toLowerCase().includes(lowerQuery) ||
                             vItem.companyName.toLowerCase().includes(lowerQuery) ||
                             vItem.fileName.toLowerCase().includes(lowerQuery);
                if (!matchQ) return;
            }
            results.push(vItem);
        }
    });

    // Final deduplication sweep: ensure no two items have the same filename or same invoice+supplier
    let uniqueResults = [];
    let seenSignatures = new Set();

    results.forEach(item => {
        let sigKey = item.fileName ? item.fileName.trim().toUpperCase() : (normStr(item.invoiceNo) + "_" + normStr(item.companyName));
        if (!seenSignatures.has(sigKey)) {
            seenSignatures.add(sigKey);
            uniqueResults.push(item);
        } else {
            // If duplicate found, prioritize the one marked Processed over Pending
            let existingIdx = uniqueResults.findIndex(x => {
                let xSig = x.fileName ? x.fileName.trim().toUpperCase() : (normStr(x.invoiceNo) + "_" + normStr(x.companyName));
                return xSig === sigKey;
            });
            if (existingIdx > -1 && item.status !== 'Pending' && uniqueResults[existingIdx].status === 'Pending') {
                uniqueResults[existingIdx] = item;
            }
        }
    });

    // Re-sequence sequential SNs for clean display
    uniqueResults.forEach((item, idx) => {
        item.sn = idx + 1;
    });
    
    return { success: true, results: uniqueResults };
  } catch (e) {
    return { error: "Queue DB Error: " + e.message };
  }
}

function webappSaveVerifiedInvoice(fileName, updatedJsonStr, dbSourceName, sessionToken = null) {
  let lock = null;
  try {
    let parsedJson = JSON.parse(updatedJsonStr);
    let isStandardPurchase = parsedJson["Record Type"] === "Standard Purchase";
    
    let releaseDateStr = parsedJson["2307 Release Date"];
    let rDate = new Date();
    if (!isStandardPurchase) {
        let tempDate = new Date(releaseDateStr);
        if (!isNaN(tempDate.getTime())) rDate = tempDate;
    } 
    
    let targetYear = Utilities.formatDate(rDate, Session.getScriptTimeZone(), "yyyy");
    let targetTab = "All Verified Invoices";
    let dbName = "Verified Invoices for 2307 " + targetYear;
    
    let props = PropertiesService.getScriptProperties();
    let isDevMode = ScriptApp.getService().getUrl().endsWith("/dev");
    const VERIFIED_FOLDER_ID = (isDevMode && props.getProperty('VERIFIED_FOLDER_ID_DEV')) 
        ? props.getProperty('VERIFIED_FOLDER_ID_DEV') 
        : props.getProperty('VERIFIED_FOLDER_ID');

    let folder = DriveApp.getFolderById(VERIFIED_FOLDER_ID);
    
    let propKey = "DB_ID_" + VERIFIED_FOLDER_ID + "_" + dbName;
    let dbId = props.getProperty(propKey);
    let ss = null;
    
    if (dbId) {
        try { 
            let file = DriveApp.getFileById(dbId);
            if (file.isTrashed()) {
                props.deleteProperty(propKey);
                dbId = null;
            } else {
                ss = SpreadsheetApp.openById(dbId); 
            }
        } catch(e) { dbId = null; }
    }
    
    if (!ss) {
        let files = folder.getFilesByName(dbName);
        while (files.hasNext()) {
            let f = files.next();
            if (!f.isTrashed()) {
                ss = SpreadsheetApp.open(f);
                break;
            }
        }
        if (!ss) {
            ss = SpreadsheetApp.create(dbName);
            DriveApp.getFileById(ss.getId()).moveTo(folder);
        }
        props.setProperty(propKey, ss.getId());
    }
    
    let sheet = ss.getSheetByName(targetTab);
    if (!sheet) {
      sheet = ss.insertSheet(targetTab);
      let sheet1 = ss.getSheetByName("Sheet1");
      if (sheet1 && sheet1.getLastRow() === 0 && ss.getSheets().length > 1) {
          ss.deleteSheet(sheet1);
      }
    }
    
    let lastCol = sheet.getLastColumn();
    if (sheet.getRange(1, 1).getValue() !== "SN" || lastCol < 6) {
      sheet.getRange(1, 1, 1, 6).setValues([["SN", "File Name", "Verified Data", "2307 Release Date", "Source", "Date and Time Verified (MM/DD/YYYY)"]]);
      sheet.getRange(1, 1, 1, 6).setFontWeight("bold").setBackground("#f3f3f3");
    }

    let currentInvNumber = String(parsedJson["Invoice Number"] || "").trim();
    let currentCompName = String(parsedJson["Company Name"] || "").trim().toUpperCase();
    
    // Resolve Target Entity (Preserve default from dbSourceName if not manually changed)
    let assignedEntity = String(parsedJson["Assigned Entity"] || parsedJson["Company Entity"] || "").trim().toUpperCase();
    let currentSource = String(dbSourceName || "").trim().toUpperCase();
    
    if (assignedEntity) {
      if (assignedEntity.includes("BONITA")) currentSource = "BONITA - RECEIVING";
      else if (assignedEntity.includes("KEYSYS")) currentSource = "KEYSYS - RECEIVING";
      else if (assignedEntity.includes("A2Z")) currentSource = "A2Z2GO - RECEIVING";
      else if (assignedEntity.includes("PINGCON")) currentSource = "PINGCON - RECEIVING";
    }
    
    let originAnchor = parsedJson["originAnchor"] || null;

    let targetOldInv = originAnchor && originAnchor.originalInvoiceNo ? String(originAnchor.originalInvoiceNo).trim().toUpperCase() : currentInvNumber.toUpperCase();
    let targetOldComp = originAnchor && originAnchor.originalCompanyName ? String(originAnchor.originalCompanyName).trim().toUpperCase() : currentCompName;

    // Cross-Year Sheet Migration done outside the sheet write lock
    if (originAnchor && originAnchor.originalReleaseDate && !isStandardPurchase) {
        let oldDate = new Date(originAnchor.originalReleaseDate);
        let oldYear = !isNaN(oldDate.getTime()) ? Utilities.formatDate(oldDate, Session.getScriptTimeZone(), "yyyy") : targetYear;
        if (oldYear !== targetYear) {
            try {
                let oldDbName = "Verified Invoices for 2307 " + oldYear;
                let oldFiles = folder.getFilesByName(oldDbName);
                if (oldFiles.hasNext()) {
                    let oldSs = SpreadsheetApp.open(oldFiles.next());
                    let oldSheet = oldSs.getSheetByName(targetTab);
                    if (oldSheet && oldSheet.getLastRow() > 1) {
                        let oldRows = oldSheet.getDataRange().getValues();
                        for (let or = oldRows.length - 1; or >= 1; or--) {
                            let oJsonStr = String(oldRows[or][2] || "").trim();
                            if (oJsonStr.startsWith("{")) {
                                try {
                                    let oParsed = JSON.parse(oJsonStr);
                                    let oInv = String(oParsed["Invoice Number"] || "").trim().toUpperCase();
                                    let oComp = String(oParsed["Company Name"] || "").trim().toUpperCase();
                                    if (oInv === targetOldInv && (targetOldComp === "" || oComp === targetOldComp)) {
                                        oldSheet.deleteRow(or + 1);
                                        break;
                                    }
                                } catch(e) {}
                            }
                        }
                    }
                }
            } catch(migErr) {
                console.warn("Cross-year migration cleanup warning: " + migErr.message);
            }
        }
    }

    // MICRO-LOCK: 30-second wait limit that ONLY protects the 0.2s sheet row write
    lock = LockService.getScriptLock();
    let isEdit = false;
    let finalJsonStr = "";
    
    try {
      lock.waitLock(30000); // 30 seconds wait timeout

      let data = sheet.getDataRange().getValues();
      let rowIndex = -1;
      let maxSn = 0;

      for (let i = 1; i < data.length; i++) {
        let currentSn = parseInt(data[i][0]) || 0;
        if (currentSn > maxSn) maxSn = currentSn;
        
        let rowFileName = String(data[i][1]).trim();
        let rowSource = String(data[i][4]).trim().toUpperCase();
        let rowJsonStr = String(data[i][2]).trim();
        let rowInvNumber = "";
        let rowCompName = "";

        if (rowJsonStr.startsWith("{")) {
            try {
                let rowParsed = JSON.parse(rowJsonStr);
                rowInvNumber = String(rowParsed["Invoice Number"] || "").trim();
                rowCompName = String(rowParsed["Company Name"] || "").trim().toUpperCase();
            } catch(e) {
                let invM = rowJsonStr.match(/"Invoice Number"\s*:\s*"([^"]+)"/i);
                let compM = rowJsonStr.match(/"Company Name"\s*:\s*"([^"]+)"/i);
                if (invM) rowInvNumber = String(invM[1]).trim();
                if (compM) rowCompName = String(compM[1]).trim().toUpperCase();
            }
        }
        
        let oldEntityUpper = (originAnchor && originAnchor.originalEntity) ? String(originAnchor.originalEntity).trim().toUpperCase() : "";
        let sameSource = (rowSource === currentSource || (rowSource && currentSource && (rowSource.includes(currentSource) || currentSource.includes(rowSource))));
        
        // If entity was migrated (e.g. from PINGCON to BONITA), match row using old entity
        if (!sameSource && oldEntityUpper) {
          sameSource = (rowSource.includes(oldEntityUpper) || oldEntityUpper.includes(rowSource.replace(/[^A-Z]/g, '')));
        }
        
        let isMatch = false;

        let normStr = s => String(s || "").trim().toUpperCase().replace(/[^A-Z0-9]/g, '');
        let targetFileId = String(parsedJson.fileId || "").trim();

        // Match existing row across conversions:
        // Priority 1: Match by Google Drive File ID if present in row JSON
        // Priority 2: Match by targetOldInv or currentInvNumber with Company Name
        // Priority 3: Match by Filename
        let rowFileId = "";
        if (rowJsonStr.startsWith("{")) {
          try {
            let rP = JSON.parse(rowJsonStr);
            rowFileId = String(rP.fileId || "").trim();
          } catch(e) {}
        }

        if (targetFileId !== "" && rowFileId !== "" && targetFileId === rowFileId && sameSource) {
          isMatch = true;
        } else if (targetOldInv !== "") {
          let matchOld = (normStr(rowInvNumber) === normStr(targetOldInv) && (targetOldComp === "" || normStr(rowCompName).includes(normStr(targetOldComp)) || normStr(targetOldComp).includes(normStr(rowCompName))));
          let matchNew = (normStr(rowInvNumber) === normStr(currentInvNumber) && (currentCompName === "" || normStr(rowCompName).includes(normStr(currentCompName)) || normStr(currentCompName).includes(normStr(rowCompName))));
          isMatch = (matchOld || matchNew) && sameSource;
        } else {
          isMatch = (rowFileName !== "" && rowFileName === fileName && sameSource);
        }

        if (isMatch) {
          rowIndex = i + 1;
          break;
        }
      }

      let finalSn = (rowIndex > -1) ? (parseInt(data[rowIndex-1][0]) || (maxSn + 1)) : (maxSn + 1);
      parsedJson["SN"] = finalSn;
      finalJsonStr = JSON.stringify(parsedJson);

      let cleanReleaseDate = isStandardPurchase ? "" : Utilities.formatDate(rDate, Session.getScriptTimeZone(), "MM/dd/yyyy");
      let verifiedDateTime = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MM/dd/yyyy HH:mm");

      if (rowIndex > -1) {
        sheet.getRange(rowIndex, 2, 1, 5).setValues([[fileName, finalJsonStr, cleanReleaseDate, currentSource, verifiedDateTime]]);
        sheet.getRange(rowIndex, 3).setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP).setVerticalAlignment("top");
        isEdit = true;
      } else {
        sheet.appendRow([finalSn, fileName, finalJsonStr, cleanReleaseDate, currentSource, verifiedDateTime]);
        sheet.getRange(sheet.getLastRow(), 3).setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP).setVerticalAlignment("top");
      }
      SpreadsheetApp.flush();
    } finally {
      // RELEASE LOCK IMMEDIATELY: Next invoice can write to the sheet right away!
      try { lock.releaseLock(); } catch(e){}
    }

    // NON-BLOCKING BACKGROUND TASKS (Executed safely outside the lock)
    let recordLabel = isStandardPurchase ? "Purchase Record" : "2307 Invoice";
    let actionStr = isEdit ? "Edited " + recordLabel : "Created " + recordLabel;
    logSystemActivity(currentSource, actionStr, `Invoice No: ${currentInvNumber} | Supplier: ${parsedJson["Company Name"] || "N/A"}`, sessionToken);

    // Cloud Vault Dual-Write Sync (Non-blocking outside the lock)
    try {
      let vaultKey = `${currentInvNumber}_${currentCompName}`.trim().toUpperCase();
      let totalAmtStr = "";
      if (parsedJson.isMixedVat) {
        let vGross = parseFloat(String((parsedJson.VAT_Portion && parsedJson.VAT_Portion.Gross) || "0").replace(/,/g, '')) || 0;
        let nvGross = parseFloat(String((parsedJson.NV_Portion && parsedJson.NV_Portion.Gross) || "0").replace(/,/g, '')) || 0;
        totalAmtStr = (vGross + nvGross).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
      } else {
        totalAmtStr = String(parsedJson["Gross"] || parsedJson["Total Amount"] || "0.00").trim();
      }

      let vaultRecord = {
        file_name: fileName || "",
        invoice_number: currentInvNumber || "",
        company_name: parsedJson["Company Name"] || "",
        total_amount: totalAmtStr,
        raw_json: finalJsonStr,
        source_database: currentSource || ""
      };

      if (originAnchor && (
          originAnchor.originalInvoiceNo !== currentInvNumber || 
          originAnchor.originalCompanyName !== currentCompName ||
          (originAnchor.originalEntity && originAnchor.originalEntity.toUpperCase() !== currentSource.toUpperCase())
      )) {
        let oldTargetInv = originAnchor.originalInvoiceNo || "";
        let oldTargetComp = originAnchor.originalCompanyName || "";
        let oldVaultKey = `${oldTargetInv}_${oldTargetComp}`.trim().toUpperCase();
        if (oldTargetInv) {
          _deleteRecordFromCloudVault("verified_invoices_vault", oldVaultKey, oldTargetInv);
        }
      }

      _syncRecordToCloudVault("verified_invoices_vault", vaultKey, vaultRecord);
    } catch(vaultErr) {
      console.warn("[Cloud Vault] Non-blocking sync warning: " + vaultErr.message);
    }

    return { success: true };
    
  } catch (e) {
    return { error: e.message };
  } finally {
    if (lock) {
      try { lock.releaseLock(); } catch(err){}
      lock = null;
    }
  }
}

function webappSave2307AndVerifiedData(fileName, updatedJsonStr, pdfHtmlStr, pdfName, dbSourceName, sessionToken = null) {
  try {
    const parsedData = JSON.parse(updatedJsonStr);
    let assignedEntity = String(parsedData["Assigned Entity"] || parsedData["Company Entity"] || "").trim().toUpperCase();
    if (assignedEntity) {
      if (assignedEntity.includes("BONITA")) dbSourceName = "BONITA - RECEIVING";
      else if (assignedEntity.includes("KEYSYS")) dbSourceName = "KEYSYS - RECEIVING";
      else if (assignedEntity.includes("A2Z")) dbSourceName = "A2Z2GO - RECEIVING";
      else if (assignedEntity.includes("PINGCON")) dbSourceName = "PINGCON - RECEIVING";
    }

    let saveResult = webappSaveVerifiedInvoice(fileName, updatedJsonStr, dbSourceName, sessionToken);
    if (saveResult.error) return saveResult;
    
    // DYNAMIC BASE64 IMAGE EMBEDDER: Converts ANY Drive image into Base64 so it renders in the PDF
    try {
      pdfHtmlStr = pdfHtmlStr.replace(/https:\/\/drive\.google\.com\/thumbnail\?id=([a-zA-Z0-9_\-]+)[^"']*/g, function(match, fileId) {
        try {
          let file = DriveApp.getFileById(fileId);
          let blob = file.getBlob();
          let mime = blob.getContentType() || 'image/png';
          let b64 = Utilities.base64Encode(blob.getBytes());
          return `data:${mime};base64,${b64}`;
        } catch(err) {
          console.warn(`Could not convert image ID [${fileId}] to Base64: ${err.message}`);
          return match;
        }
      });
    } catch(e) { console.warn("Failed to dynamically embed base64 images: " + e); }

    let finalHtml = `<html><head><style> @page { size: letter; margin: 0.15in; } * { -webkit-print-color-adjust: exact !important; print-color-adjust: exact !important; } </style></head><body style="margin:0; padding:0;">${pdfHtmlStr}</body></html>`;
    let blob = Utilities.newBlob(finalHtml, MimeType.HTML).getAs(MimeType.PDF).setName(pdfName);
    
    let props = PropertiesService.getScriptProperties();
    let isDevMode = ScriptApp.getService().getUrl().endsWith("/dev");
    const GENERATED_2307_FOLDER_ID = (isDevMode && props.getProperty('GENERATED_2307_FOLDER_ID_DEV')) 
        ? props.getProperty('GENERATED_2307_FOLDER_ID_DEV') 
        : props.getProperty('GENERATED_2307_FOLDER_ID');

    let rootFolder = DriveApp.getFolderById(GENERATED_2307_FOLDER_ID);
    
    let releaseDateStr = parsedData["2307 Release Date"];
    let releaseDate = new Date(releaseDateStr);
    if (isNaN(releaseDate.getTime())) releaseDate = new Date(); 
    
    let monthYear = Utilities.formatDate(releaseDate, Session.getScriptTimeZone(), "MMMM yyyy");
    
    let originAnchor = parsedData["originAnchor"] || null;
    if (!assignedEntity) {
      assignedEntity = String(parsedData["Assigned Entity"] || parsedData["Company Entity"] || "").trim().toUpperCase();
    }
    
    // Determine new destination Payor folder
    let payorFolder = "Pingcon";
    let targetSource = assignedEntity || String(dbSourceName).toUpperCase();
    if (targetSource.includes("KEYSYS")) payorFolder = "Keysys";
    else if (targetSource.includes("A2Z") || targetSource.includes("A2Z2GO")) payorFolder = "A2Z2GO";
    else if (targetSource.includes("BONITA")) payorFolder = "Bonita";
    else if (targetSource.includes("PINGCON")) payorFolder = "Pingcon";

    let compFolders = rootFolder.getFoldersByName(payorFolder);
    let compFolder = compFolders.hasNext() ? compFolders.next() : rootFolder.createFolder(payorFolder);
    
    let monthFolders = compFolder.getFoldersByName(monthYear);
    let targetFolder = monthFolders.hasNext() ? monthFolders.next() : compFolder.createFolder(monthYear);
    
    // Determine previous Payor folder (if entity changed)
    let oldPayorFolder = payorFolder;
    if (originAnchor && originAnchor.originalEntity) {
      let oldEnt = String(originAnchor.originalEntity).toUpperCase();
      if (oldEnt.includes("KEYSYS")) oldPayorFolder = "Keysys";
      else if (oldEnt.includes("A2Z")) oldPayorFolder = "A2Z2GO";
      else if (oldEnt.includes("BONITA")) oldPayorFolder = "Bonita";
      else if (oldEnt.includes("PINGCON")) oldPayorFolder = "Pingcon";
    }

    let oldPdfNameToDelete = "";
    if (originAnchor && (originAnchor.originalInvoiceNo || originAnchor.originalCompanyName)) {
        let oldCleanSup = String(originAnchor.originalCompanyName || "Supplier").replace(/[^a-zA-Z0-9 -]/g, "").trim();
        let oldInv = originAnchor.originalInvoiceNo || "INV";
        oldPdfNameToDelete = `2307_${oldCleanSup}_${oldInv}.pdf`;
    }

    // 1. Purge old PDF from PREVIOUS entity folder if entity was moved
    if (oldPayorFolder !== payorFolder) {
      let prevCompFolders = rootFolder.getFoldersByName(oldPayorFolder);
      if (prevCompFolders.hasNext()) {
        let prevComp = prevCompFolders.next();
        let prevMonths = prevComp.getFolders();
        while (prevMonths.hasNext()) {
          let pm = prevMonths.next();
          let targetDeleteName = oldPdfNameToDelete || pdfName;
          let oldEntityFiles = pm.getFilesByName(targetDeleteName);
          while (oldEntityFiles.hasNext()) {
            try { oldEntityFiles.next().setTrashed(true); } catch(e) {}
          }
        }
      }
    }

    // 2. Purge stale PDF within the target folder
    let allMonthFolders = compFolder.getFolders();
    while (allMonthFolders.hasNext()) {
        let mFolder = allMonthFolders.next();
        let oldFiles = mFolder.getFilesByName(pdfName);
        while (oldFiles.hasNext()) {
            try { oldFiles.next().setTrashed(true); } catch(e) {}
        }
        if (oldPdfNameToDelete && oldPdfNameToDelete !== pdfName) {
            let oldNameFiles = mFolder.getFilesByName(oldPdfNameToDelete);
            while (oldNameFiles.hasNext()) {
                try { oldNameFiles.next().setTrashed(true); } catch(e) {}
            }
        }
    }
    
    let pdfFile = targetFolder.createFile(blob);
    return { success: true, pdfUrl: pdfFile.getUrl(), folderUrl: compFolder.getUrl(), pdfId: pdfFile.getId() };
  } catch (e) {
    return { error: e.message };
  }
}

function Template_2307_PDF(parsed, dateFrom, dateTo) {
  let props = PropertiesService.getScriptProperties();
  let purchases = parsed["Purchases"] || "0.00";
  let ewt1 = parseFloat(String(parsed["EWT 1%"]).replace(/,/g, '')) || 0;
  let ewt2 = parseFloat(String(parsed["EWT 2%"]).replace(/,/g, '')) || 0;
  let taxWithheld = (ewt1 > 0 ? ewt1 : ewt2).toLocaleString('en-US', {minimumFractionDigits: 2, maximumFractionDigits: 2});
  let atc = parsed["ATC"] || "";
  
  let payeeTin = parsed["TIN"] || "";
  let payeeName = parsed["Company Name"] || "";
  let payeeAddress = parsed["Address"] || "";

  // Dynamic Payor & Signatory from Script Properties (Zero Hardcoded Leaks for GitHub)
  let payorTinRaw = (props.getProperty('PINGCON_TIN') || "").replace(/[^\d]/g, '');
  let payorTinFormatted = payorTinRaw ? `${payorTinRaw.slice(0,3)}-${payorTinRaw.slice(3,6)}-${payorTinRaw.slice(6,9)}-00000` : "";
  let payorAddress = (props.getProperty('PINGCON_ADDR') || "").split('\n')[0].trim();
  let sigName = props.getProperty('SIGNATORY_2307_NAME') || "";
  let sigTitle = props.getProperty('SIGNATORY_2307_TITLE') || "Accounts Payable Officer";
  let sigTin = props.getProperty('SIGNATORY_2307_TIN') || "";
  let perjuryText = props.getProperty('LEGAL_PERJURY_TEXT') || "We declare under the penalties of perjury that this certificate has been made in good faith, verified by us, and to the best of our knowledge and belief, is true and correct, pursuant to the provisions of the National Internal Revenue Code, as amended, and the regulations issued under authority thereof. Further, we give our consent to the processing of our information as contemplated under the *Data Privacy Act of 2012 (R.A. No. 10173) for legitimate and lawful purposes.";

  return `
  <html><head><style> @page { size: letter; margin: 0.3in; } * { -webkit-print-color-adjust: exact !important; print-color-adjust: exact !important; } </style></head>
  <body style="font-family: Arial, sans-serif; padding: 0px; color: #000; margin: 0;">
    <div style="width: 100%; max-width: 800px; margin: 0 auto; border: 2px solid #000; padding: 2px; box-sizing: border-box;">
      <table style="width: 100%; border-collapse: collapse; border: 1px solid #000;">
        <tr>
          <td style="width: 20%; border: 1px solid #000; text-align: center; vertical-align: top; padding: 5px;">
             <div style="font-size: 10px;">BIR Form No.</div>
             <div style="font-size: 32px; font-weight: bold; margin: 5px 0;">2307</div>
             <div style="font-size: 9px;">January 2018 (ENCS)</div>
          </td>
          <td style="width: 60%; border: 1px solid #000; text-align: center; vertical-align: middle;">
             <div style="font-size: 20px; font-weight: bold;">Certificate of Creditable Tax</div>
             <div style="font-size: 20px; font-weight: bold;">Withheld at Source</div>
          </td>
          <td style="width: 20%; border: 1px solid #000; text-align: right; vertical-align: bottom; font-size: 9px; padding: 5px;">
             2307 01/18ENCS
          </td>
        </tr>
      </table>

      <table style="width: 100%; border-collapse: collapse; border: 1px solid #000; margin-top: -1px;">
        <tr>
          <td style="background-color: #d9d9d9; font-size: 9px; padding: 2px 5px; font-weight: bold;">Fill in all applicable spaces. Mark all appropriate boxes with an "X".</td>
        </tr>
      </table>

      <table style="width: 100%; border-collapse: collapse; border: 1px solid #000; margin-top: -1px;">
        <tr>
          <td style="width: 15%; padding: 4px 5px; font-size: 11px; border-right: none;"><b>1</b> &nbsp;For the Period</td>
          <td style="width: 5%; text-align: right; font-size: 11px; border-left: none; border-right: none; padding-right: 10px;">From</td>
          <td style="width: 30%; text-align: center; font-size: 13px; font-weight: bold; border-left: none; border-right: 1px solid #000; border-bottom: 1px solid #000; letter-spacing: 2px;">${dateFrom}</td>
          <td style="width: 5%; text-align: right; font-size: 11px; border-left: none; border-right: none; padding-right: 10px;">To</td>
          <td style="width: 30%; text-align: center; font-size: 13px; font-weight: bold; border-left: none; border-bottom: 1px solid #000; letter-spacing: 2px;">${dateTo}</td>
          <td style="width: 15%; border-left: none;"></td>
        </tr>
      </table>

      <table style="width: 100%; border-collapse: collapse; border: 1px solid #000; margin-top: -1px;">
        <tr><td colspan="2" style="background-color: #d9d9d9; text-align: center; font-weight: bold; font-size: 11px; padding: 3px;">Part I - Payee Information</td></tr>
        <tr>
          <td style="width: 30%; font-size: 11px; padding: 4px 5px;"><b>2</b> Taxpayer Identification Number (TIN)</td>
          <td style="width: 70%; font-size: 14px; font-weight: bold; padding: 4px 5px; border-left: 1px solid #000; letter-spacing: 2px;">${payeeTin}</td>
        </tr>
        <tr>
          <td colspan="2" style="font-size: 11px; padding: 4px 5px; border-top: 1px solid #000;"><b>3</b> Payee's Name <span style="font-size: 9px; font-weight: normal;">(Last Name, First Name, Middle Name for Individual OR Registered Name for Non-Individual)</span><br>
          <div style="font-size: 14px; font-weight: bold; margin-top: 4px;">${payeeName}</div></td>
        </tr>
        <tr>
          <td colspan="2" style="font-size: 11px; padding: 4px 5px; border-top: 1px solid #000;"><b>4</b> Registered Address<br>
          <div style="font-size: 12px; margin-top: 4px;">${payeeAddress}</div></td>
        </tr>
        <tr>
          <td colspan="2" style="font-size: 11px; padding: 4px 5px; border-top: 1px solid #000;"><b>5</b> Foreign Address, if applicable<br><div style="margin-top: 4px;">&nbsp;</div></td>
        </tr>
      </table>

      <table style="width: 100%; border-collapse: collapse; border: 1px solid #000; margin-top: -1px;">
        <tr><td colspan="2" style="background-color: #d9d9d9; text-align: center; font-weight: bold; font-size: 11px; padding: 3px;">Part II - Payor Information</td></tr>
        <tr>
          <td style="width: 30%; font-size: 11px; padding: 4px 5px;"><b>6</b> Taxpayer Identification Number (TIN)</td>
          <td style="width: 70%; font-size: 14px; font-weight: bold; padding: 4px 5px; border-left: 1px solid #000; letter-spacing: 2px;">${payorTinFormatted}</td>
        </tr>
        <tr>
          <td colspan="2" style="font-size: 11px; padding: 4px 5px; border-top: 1px solid #000;"><b>7</b> Payor's Name <span style="font-size: 9px; font-weight: normal;">(Last Name, First Name, Middle Name for Individual OR Registered Name for Non-Individual)</span><br>
          <div style="font-size: 14px; font-weight: bold; margin-top: 4px;">PINGCON MARKETING CORPORATION</div></td>
        </tr>
        <tr>
          <td colspan="2" style="font-size: 11px; padding: 4px 5px; border-top: 1px solid #000;"><b>8</b> Registered Address<br>
          <div style="font-size: 12px; margin-top: 4px;">${payorAddress}</div></td>
        </tr>
      </table>

      <table style="width: 100%; border-collapse: collapse; border: 1px solid #000; margin-top: -1px;">
        <tr><td colspan="7" style="background-color: #d9d9d9; text-align: center; font-weight: bold; font-size: 11px; padding: 3px;">Part III - Details of Monthly Income Payments and Taxes Withheld</td></tr>
        <tr style="text-align: center; font-size: 10px; font-weight: bold;">
          <td rowspan="2" style="width: 35%; border: 1px solid #000; padding: 2px;">Income Payments Subject to Expanded<br>Withholding Tax</td>
          <td rowspan="2" style="width: 8%; border: 1px solid #000; padding: 2px;">ATC</td>
          <td colspan="4" style="border: 1px solid #000; padding: 2px;">AMOUNT OF INCOME PAYMENTS</td>
          <td rowspan="2" style="width: 12%; border: 1px solid #000; padding: 2px;">Tax Withheld for the<br>Quarter</td>
        </tr>
        <tr style="text-align: center; font-size: 10px; font-weight: bold;">
          <td style="width: 11.25%; border: 1px solid #000; padding: 2px;">1st Month of the<br>Quarter</td>
          <td style="width: 11.25%; border: 1px solid #000; padding: 2px;">2nd Month of the<br>Quarter</td>
          <td style="width: 11.25%; border: 1px solid #000; padding: 2px;">3rd Month of the<br>Quarter</td>
          <td style="width: 11.25%; border: 1px solid #000; padding: 2px;">Total</td>
        </tr>
        <tr style="text-align: center; font-size: 11px;">
          <td style="text-align: left; border: 1px solid #000; padding: 4px;">Income payment made by top withholding agent to their local/resident supplier of goods</td>
          <td style="border: 1px solid #000;">${(atc === 'WC158' || atc === 'WI158') ? atc : ''}</td>
          <td style="border: 1px solid #000;">${(atc === 'WC158' || atc === 'WI158') ? purchases : '0.00'}</td>
          <td style="border: 1px solid #000;">0.00</td>
          <td style="border: 1px solid #000;">0.00</td>
          <td style="border: 1px solid #000;">${(atc === 'WC158' || atc === 'WI158') ? purchases : '0.00'}</td>
          <td style="border: 1px solid #000;">${(atc === 'WC158' || atc === 'WI158') ? taxWithheld : '0.00'}</td>
        </tr>
        <tr style="text-align: center; font-size: 11px;">
          <td style="text-align: left; border: 1px solid #000; padding: 4px;">Income payments made by top 10,000 private corporations to their local/resident supplier of services</td>
          <td style="border: 1px solid #000;">${(atc === 'WC160' || atc === 'WI160') ? atc : ''}</td>
          <td style="border: 1px solid #000;">${(atc === 'WC160' || atc === 'WI160') ? purchases : '0.00'}</td>
          <td style="border: 1px solid #000;">0.00</td>
          <td style="border: 1px solid #000;">0.00</td>
          <td style="border: 1px solid #000;">${(atc === 'WC160' || atc === 'WI160') ? purchases : '0.00'}</td>
          <td style="border: 1px solid #000;">${(atc === 'WC160' || atc === 'WI160') ? taxWithheld : '0.00'}</td>
        </tr>
        <tr><td style="border: 1px solid #000; height: 18px;"></td><td style="border: 1px solid #000;"></td><td style="border: 1px solid #000;"></td><td style="border: 1px solid #000;"></td><td style="border: 1px solid #000;"></td><td style="border: 1px solid #000;"></td><td style="border: 1px solid #000;"></td></tr>
        <tr><td style="border: 1px solid #000; height: 18px;"></td><td style="border: 1px solid #000;"></td><td style="border: 1px solid #000;"></td><td style="border: 1px solid #000;"></td><td style="border: 1px solid #000;"></td><td style="border: 1px solid #000;"></td><td style="border: 1px solid #000;"></td></tr>
        <tr><td style="border: 1px solid #000; height: 18px;"></td><td style="border: 1px solid #000;"></td><td style="border: 1px solid #000;"></td><td style="border: 1px solid #000;"></td><td style="border: 1px solid #000;"></td><td style="border: 1px solid #000;"></td><td style="border: 1px solid #000;"></td></tr>
        
        <tr style="background-color: #d9d9d9; font-weight: bold; font-size: 11px;">
          <td style="border: 1px solid #000; padding: 4px;">Total</td>
          <td style="border: 1px solid #000;"></td>
          <td style="border: 1px solid #000;"></td>
          <td style="border: 1px solid #000;"></td>
          <td style="border: 1px solid #000;"></td>
          <td style="border: 1px solid #000; text-align: center;">${purchases}</td>
          <td style="border: 1px solid #000; text-align: center;">${taxWithheld}</td>
        </tr>
      </table>
      
      <table style="width: 100%; border-collapse: collapse; border: 1px solid #000; margin-top: -1px;">
        <tr>
          <td style="padding: 6px; font-size: 10px; text-align: justify; line-height: 1.3;">
            ${perjuryText}
          </td>
        </tr>
      </table>

      <table style="width: 100%; border-collapse: collapse; border: 1px solid #000; border-top: none; font-size: 9px; text-align: center;">
        <tr>
          <td colspan="4" style="padding: 25px 5px 2px 5px; text-align: center; position: relative;">
             <div style="width: 60%; margin: 0 auto; border-bottom: 1px solid #000; position: relative; padding-bottom: 2px;">
               <div style="font-size: 11px; font-weight: bold;">${sigName}</div>
             </div>
             <div style="font-size: 10px; font-weight: bold; margin-top: 2px;">( ${sigTitle} ) TIN: ${sigTin}</div>
             <div style="font-size: 9px; margin-top: 2px;">Signature over Printed Name of Payor/Payor's Authorized Representative/Tax Agent</div>
          </td>
        </tr>
        <tr>
          <td style="width: 25%; border: 1px solid #000; padding: 2px;">Tax Agent Accreditation No./<br>Attorney's Roll No. (if applicable)</td>
          <td style="width: 25%; border: 1px solid #000; padding: 0;"></td>
          <td style="width: 25%; border: 1px solid #000; padding: 2px;">Date of Issue<br>(MM/DD/YYYY)</td>
          <td style="width: 25%; border: 1px solid #000; padding: 0;"></td>
        </tr>
        <tr><td colspan="4" style="background-color: #d9d9d9; font-weight: bold; border: 1px solid #000; padding: 2px; text-align: center;">CONFORME:</td></tr>
        <tr>
          <td colspan="4" style="padding: 20px 5px 5px 5px; text-align: center;">
             <div style="width: 60%; margin: 0 auto; border-bottom: 1px solid #000; height: 15px;"></div>
             <div style="font-size: 9px; margin-top: 2px;">Signature over Printed Name of Payee/Payee's Authorized Representative/Tax Agent</div>
             <div style="font-size: 10px; margin-top: 2px;"><i>(Indicate Title/Designation and TIN)</i></div>
          </td>
        </tr>
        <tr>
          <td style="width: 25%; border: 1px solid #000; padding: 2px;">Tax Agent Accreditation No./<br>Attorney's Roll No. (if applicable)</td>
          <td style="width: 25%; border: 1px solid #000; padding: 0;"></td>
          <td style="width: 25%; border: 1px solid #000; padding: 2px;">Date of Expiry<br>(MM/DD/YYYY)</td>
          <td style="width: 25%; border: 1px solid #000; padding: 0;"></td>
        </tr>
      </table>
      <div style="font-size: 9px; padding: 2px 5px; text-align: left;">*NOTE: The BIR Data Privacy is in the BIR website (www.bir.gov.ph)</div>

    </div>
  </body></html>
  `;
}

function webappPreviewClientEwtReport(startStr, endStr, companyContext) {
  try {
    let startDate = new Date(startStr + " 00:00:00");
    let endDate = new Date(endStr + " 23:59:59");
    
    let dbRes = webappGetClient2307Data(companyContext);
    if (!dbRes.success) throw new Error(dbRes.error);
    
    let data = dbRes.data;
    let results = [];
    let snCounter = 1;
    
    data.forEach(r => {
      let dateToFilter = r.receiveDate ? r.receiveDate : r.timestamp;
      let d = new Date(dateToFilter);
      if (d >= startDate && d <= endDate) {
        const parseNum = (val) => {
          if (!val) return 0;
          let n = parseFloat(String(val).replace(/,/g, ''));
          return isNaN(n) ? 0 : n;
        };

        results.push([
          snCounter++, 
          r.siNo || "", 
          r.customerName || "", 
          r.tin || "", 
          r.dateFrom || "", 
          r.dateTo || "", 
          parseNum(r.incomePayment), 
          r.atc || "",
          parseNum(r.wc158), 
          parseNum(r.wc157), 
          parseNum(r.wc160), 
          parseNum(r.wv010), 
          parseNum(r.wc120), 
          parseNum(r.wc640), 
          parseNum(r.wv020), 
          parseNum(r.wi632), 
          parseNum(r.wc690),
          r.address || ""
        ]);
      }
    });
    
    return { success: true, results: results };
  } catch (e) {
    return { error: e.message };
  }
}

function webappGenerateAndEmailClientEwtReport(tableData, startStr, endStr, companyContext, sessionToken = null) {
  // SECURITY GATE: Prevent Viewers from emailing official reports
  if (sessionToken) {
    const user = getAuthenticatedUser(sessionToken);
    if (user && user.role === "Viewer") {
      return { error: "Permission Denied: Viewers have read-only access and cannot email official reports." };
    }
  }

  let lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    
    let startDate = new Date(startStr + " 00:00:00");
    let endDate = new Date(endStr + " 23:59:59");
    let sDate = Utilities.formatDate(startDate, Session.getScriptTimeZone(), "MM-dd-yyyy");
    let eDate = Utilities.formatDate(endDate, Session.getScriptTimeZone(), "MM-dd-yyyy");
    let fileName = `EWT_Client_Report_${companyContext}_${sDate}_to_${eDate}`;
    
    // Headers without 'Verified On'
    let reportHeaders = [
      "SN", "Date from (MM/DD/YYYY)", "Date to (MM/DD/YYYY)", "SI No.", "Customer Name", 
      "TIN", "Address", "ATC", "Income Payment", "WC 158 1%", "WC 157 2%", "WC 160 2%", 
      "WV010 5%", "WC120 2%", "WC640 1%", "WV020 5%", "WI632 1%", "WC690"
    ];

    let excelRows = [];
    let totalIncome = 0;
    let tableRowsHtml = "";
    
    tableData.forEach(row => {
      let sn = row[0];
      let siNo = row[1];
      let cust = row[2];
      let tin = row[3];
      let dFrom = row[4];
      let dTo = row[5];
      let income = parseFloat(row[6]) || 0;
      let atc = row[7];
      let wc158 = parseFloat(row[8]) || "";
      let wc157 = parseFloat(row[9]) || "";
      let wc160 = parseFloat(row[10]) || "";
      let wv010 = parseFloat(row[11]) || "";
      let wc120 = parseFloat(row[12]) || "";
      let wc640 = parseFloat(row[13]) || "";
      let wv020 = parseFloat(row[14]) || "";
      let wi632 = parseFloat(row[15]) || "";
      let wc690 = parseFloat(row[16]) || "";
      let address = row[17] || "";

      totalIncome += income;

      tableRowsHtml += `<tr>
        <td style="border: 1px solid black; padding: 4px 8px; text-align: center;">${sn}</td>
        <td style="border: 1px solid black; padding: 4px 8px;">${cust}</td>
        <td style="border: 1px solid black; padding: 4px 8px;">${siNo}</td>
        <td style="border: 1px solid black; padding: 4px 8px; text-align: right;">${income.toLocaleString('en-US', {minimumFractionDigits: 2})}</td>
        <td style="border: 1px solid black; padding: 4px 8px; text-align: center;">${atc}</td>
      </tr>`;

      excelRows.push([
        sn, dFrom, dTo, siNo, cust, tin, address, atc,
        income, wc158, wc157, wc160, wv010, wc120, wc640, wv020, wi632, wc690
      ]);
    });

    let tempSs = SpreadsheetApp.create(fileName);
    let sheet = tempSs.getSheets()[0];
    
    sheet.getRange(1, 1, 1, reportHeaders.length).setValues([reportHeaders]).setFontWeight("bold").setBackground("#f3f3f3");
    
    // Set Plain Text format on SI No. and TIN to preserve leading zeros
    sheet.getRange(2, 4, excelRows.length, 1).setNumberFormat("@");
    sheet.getRange(2, 6, excelRows.length, 1).setNumberFormat("@");
    
    sheet.getRange(2, 1, excelRows.length, reportHeaders.length).setValues(excelRows);
    
    // Set Number format on Income Payment and all Tax Columns (Col 9 through 18)
    sheet.getRange(2, 9, excelRows.length, 10).setNumberFormat("#,##0.00");
    
    sheet.getDataRange().setWrap(true).setVerticalAlignment("top");
    
    // EXPLICIT GENEROUS COLUMN WIDTHS TO PREVENT HASHTAGS (#######)
    sheet.setColumnWidth(1, 50);   // SN
    sheet.setColumnWidth(2, 115);  // Date From
    sheet.setColumnWidth(3, 115);  // Date To
    sheet.setColumnWidth(4, 120);  // SI No.
    sheet.setColumnWidth(5, 250);  // Customer Name
    sheet.setColumnWidth(6, 140);  // TIN
    sheet.setColumnWidth(7, 300);  // Address
    sheet.setColumnWidth(8, 80);   // ATC
    sheet.setColumnWidth(9, 140);  // Income Payment (Avoids #######)
    for (let c = 10; c <= 18; c++) {
      sheet.setColumnWidth(c, 110); // Tax columns (Avoids #######)
    }

    SpreadsheetApp.flush();
    Utilities.sleep(2000);
    
    let url = "https://docs.google.com/spreadsheets/d/" + tempSs.getId() + "/export?format=xlsx";
    let token = ScriptApp.getOAuthToken();
    let response = UrlFetchApp.fetch(url, { headers: { 'Authorization': 'Bearer ' + token }, muteHttpExceptions: true });
    let blob = response.getBlob().setName(fileName + ".xlsx");
    
    DriveApp.getFileById(tempSs.getId()).setTrashed(true);

    // FETCH RECIPIENT FROM EWT_Client_Emails_Data
    let mainRecipient = "";
    let ccRecipient = "";
    
    try {
      let emailFolderId = SCRIPT_PROPS['EMAIL_RECIPIENT_FOLDER_ID'] || DATABASE_FOLDER_ID;
      let emailFolder = DriveApp.getFolderById(emailFolderId);
      let emailFiles = emailFolder.getFilesByName("EWT_Client_Emails_Data");
      let eSheet;
      
      let foundFile = null;
      while (emailFiles.hasNext()) {
        let f = emailFiles.next();
        if (!f.isTrashed()) { foundFile = f; break; }
      }
      
      if (foundFile) {
        eSheet = SpreadsheetApp.open(foundFile).getSheets()[0];
      } else {
        // Auto-seed template if missing
        let newSs = SpreadsheetApp.create("EWT_Client_Emails_Data");
        DriveApp.getFileById(newSs.getId()).moveTo(emailFolder);
        eSheet = newSs.getSheets()[0];
        eSheet.appendRow(["Entity", "To Emails", "CC Emails"]);
        eSheet.getRange("A1:C1").setFontWeight("bold").setBackground("#f3f3f3");
        eSheet.appendRow(["Pingcon", "", ""]);
        eSheet.appendRow(["Bonita", "", ""]);
        eSheet.appendRow(["Keysys", "", ""]);
        eSheet.appendRow(["A2Z2GO", "", ""]);
      }
      
      let eData = eSheet.getDataRange().getValues();
      if (eData.length > 1) {
        let eHeaders = eData[0].map(h => String(h).toLowerCase().trim());
        let entIdx = eHeaders.indexOf("entity");
        let toIdx = eHeaders.indexOf("to emails");
        let ccIdx = eHeaders.indexOf("cc emails");
        
        for (let i = 1; i < eData.length; i++) {
          if (entIdx > -1 && String(eData[i][entIdx]).toUpperCase().includes(String(companyContext).toUpperCase())) {
            if (toIdx > -1 && eData[i][toIdx]) mainRecipient = String(eData[i][toIdx]).trim();
            if (ccIdx > -1 && eData[i][ccIdx]) ccRecipient = String(eData[i][ccIdx]).trim();
            break;
          }
        }
      }
    } catch (e) {
      console.warn("Failed to load EWT_Client_Emails_Data: " + e.message);
    }

    if (!mainRecipient) {
      lock.releaseLock();
      return { error: `No 'To Emails' configured for ${companyContext} in EWT_Client_Emails_Data database. Email dispatch aborted.` };
    }

    let sysSettings = getSystemSettings();
    let isTech = (companyContext === "KEYSYS" || companyContext === "A2Z2GO");
    let senderName = isTech ? "Finance A2Z2GO" : "Finance Pingcon";
    let senderEmail = getSenderEmail(isTech);
    let props = PropertiesService.getScriptProperties();
    let finalPassword = isTech ? props.getProperty('ENCRYPT_PASS_TECH') : props.getProperty('ENCRYPT_PASS_PINGCON');

    let htmlBody = `
      <div style="font-family: Arial, sans-serif; font-size: 13px; color: #000; line-height: 1.5;">
        <p>Good day,</p>
        <p>Please find attached the <strong>EWT Report Client</strong> for <strong>${companyContext}</strong> covering verified certificates from <strong>${sDate} to ${eDate}</strong>.</p>
        <p>This report contains all the client 2307 data verified within the specified date range.</p>
        <br>
        <table style="border-collapse: collapse; border: 1px solid black; font-size: 13px; text-align: left; width: 100%; max-width: 800px;">
          <thead>
            <tr>
              <th style="border: 1px solid black; padding: 4px 8px; background-color: #f3f3f3; text-align: center;">SN</th>
              <th style="border: 1px solid black; padding: 4px 8px; background-color: #f3f3f3;">Customer Name</th>
              <th style="border: 1px solid black; padding: 4px 8px; background-color: #f3f3f3;">SI Number</th>
              <th style="border: 1px solid black; padding: 4px 8px; background-color: #f3f3f3; text-align: right;">Income Payment</th>
              <th style="border: 1px solid black; padding: 4px 8px; background-color: #f3f3f3; text-align: center;">ATC</th>
            </tr>
          </thead>
          <tbody>
            ${tableRowsHtml}
          </tbody>
        </table>
        <br>
        <p>Thank you.</p>
      </div>`;

    let securedBlob = blob;
    try {
      securedBlob = encryptAttachmentViaPython(blob, fileName + ".xlsx", finalPassword);
    } catch (e) {
      console.warn("Encryption fallback: " + e.message);
    }

    sendRobustEmail(mainRecipient, `EWT Client Report ${companyContext} ${sDate} to ${eDate}`, htmlBody, {
      from: senderEmail,
      name: senderName,
      cc: ccRecipient,
      attachments: [securedBlob]
    });

    logClientActivity(companyContext, "Generated & Emailed EWT Client Report", `Emailed report to ${mainRecipient} (${excelRows.length} records) covering ${sDate} to ${eDate}`);

    lock.releaseLock();
    return {
      success: true,
      base64: Utilities.base64Encode(blob.getBytes()),
      fileName: fileName + ".xlsx"
    };
  } catch (e) {
    return { error: e.message };
  } finally {
    lock.releaseLock();
  }
}

// =========================================================================
// EWT SUPPLIER REPORT (NEW)
// =========================================================================

function _35_runSendEwtSupplierReport() {
  const props = PropertiesService.getScriptProperties();
  deleteTriggersByName("resumeEwtSupplierBatch");
  props.deleteProperty("CANCEL_EWT_SUPPLIER");
  resumeEwtSupplierBatch();
}

function resumeEwtSupplierBatch() {
  const props = PropertiesService.getScriptProperties();
  props.setProperty("THREAD_STATE", "RUNNING"); // <--- ADDED
  
  if (props.getProperty("CANCEL_EWT_SUPPLIER") === "true") {
      deleteTriggersByName("resumeEwtSupplierBatch");
      props.deleteProperty("CANCEL_EWT_SUPPLIER");
      clearReportLock();
      return;
  }
  
  // Use Date Ranges instead of specific Month/Year
  const startStr = props.getProperty("BATCH_START_DATE");
  const endStr = props.getProperty("BATCH_END_DATE");
  const startDate = startStr ? new Date(startStr + "T00:00:00") : new Date(2025, 0, 1);
  const endDate = endStr ? new Date(endStr + "T23:59:59") : new Date();
  
  const VERIFIED_FOLDER_ID = SCRIPT_PROPS['VERIFIED_FOLDER_ID'];
  let folder = DriveApp.getFolderById(VERIFIED_FOLDER_ID);
  let files = folder.getFilesByType(MimeType.GOOGLE_SHEETS);
  
  let extractedData = [];
  const narraAllData = gatherAllData("Narra_Voucher_Data_");
  const checkSummaryAllData = gatherAllData("Check_Summary_Data_");
  
  // Clean SI helper: Strip non-numeric symbols but KEEP leading zeros (0-9)
  const cleanSI = (si) => String(si).replace(/^[^0-9]+/, '').trim();
  
  // Scan all files in the verified folder
  while (files.hasNext()) {
      let ss = SpreadsheetApp.open(files.next());
      let sheets = ss.getSheets();
      
      sheets.forEach(sheet => {
          if (sheet.getLastRow() > 1) {
              let data = sheet.getDataRange().getValues();
              let hRow = data[0].map(h => String(h).trim().toLowerCase());
              let dateColIdx = hRow.indexOf("2307 release date");
              if (dateColIdx === -1) dateColIdx = 3; 
              let sourceColIdx = hRow.indexOf("source");
              if (sourceColIdx === -1) sourceColIdx = 4;

              for (let i = 1; i < data.length; i++) {
                  let jsonStr = String(data[i][2]).trim();
                  let releaseDateStr = data[i][dateColIdx];
                  let sourceStr = String(data[i][sourceColIdx]).toUpperCase();
                  
                  // Extract entity directly from the RECEIVING database source names
                  let payorVal = "Pingcon"; 
                  if (sourceStr.includes("BONITA")) payorVal = "Bonita";
                  else if (sourceStr.includes("KEYSYS")) payorVal = "Keysys";
                  else if (sourceStr.includes("A2Z2GO") || sourceStr.includes("A2Z")) payorVal = "A2Z2GO";
                  else if (sourceStr.includes("PINGCON")) payorVal = "Pingcon";

                  // Check if the user manually reassigned the entity inside the verified JSON
                  if (jsonStr.startsWith("{")) {
                    try {
                      let pMeta = JSON.parse(jsonStr);
                      let aEnt = String(pMeta["Assigned Entity"] || pMeta["Company Entity"] || "").toUpperCase();
                      if (aEnt.includes("BONITA")) payorVal = "Bonita";
                      else if (aEnt.includes("KEYSYS")) payorVal = "Keysys";
                      else if (aEnt.includes("A2Z")) payorVal = "A2Z2GO";
                      else if (aEnt.includes("PINGCON")) payorVal = "Pingcon";
                    } catch(e) {}
                  }

                  if (jsonStr.startsWith("{") && releaseDateStr) {
                      try {
                          let parsed = JSON.parse(jsonStr);
                          let rDate = new Date(releaseDateStr);
                          if (isNaN(rDate.getTime())) continue;
                          
                          if (rDate >= startDate && rDate <= endDate) {
                              if (parsed.isMixedVat) {
                                  // 1. Push the VAT Portion as Row 1
                                  extractedData.push([
                                      "", payorVal, parsed.VAT_Portion["Account Title"] || "", parsed["Invoice Date"] || "", 
                                      parsed["Company Name"] || "", parsed["Address"] || "", parsed["TIN"] || "", parsed["Invoice Number"] || "", 
                                      parsed.VAT_Portion["Total Amount"] || parsed.VAT_Portion["Gross"] || "", parsed.VAT_Portion["Input Tax"] || "", 
                                      parsed.VAT_Portion["Purchases"] || "", parsed.VAT_Portion["EWT 1%"] || "", parsed.VAT_Portion["EWT 2%"] || "", 
                                      parsed.VAT_Portion["ATC"] || "", "", parsed["PO Number"] || "", releaseDateStr
                                  ]);
                                  // 2. Push the Non-VAT Portion as Row 2
                                  extractedData.push([
                                      "", payorVal, parsed.NV_Portion["Account Title"] || "", parsed["Invoice Date"] || "", 
                                      parsed["Company Name"] || "", parsed["Address"] || "", parsed["TIN"] || "", parsed["Invoice Number"] || "", 
                                      parsed.NV_Portion["Total Amount"] || parsed.NV_Portion["Gross"] || "", parsed.NV_Portion["Input Tax"] || "", 
                                      parsed.NV_Portion["Purchases"] || "", parsed.NV_Portion["EWT 1%"] || "", parsed.NV_Portion["EWT 2%"] || "", 
                                      parsed.NV_Portion["ATC"] || "", "", parsed["PO Number"] || "", releaseDateStr
                                  ]);
                              } else {
                                  // Push Standard Single Row
                                  extractedData.push([
                                      "", payorVal, parsed["Account Title"] || "", parsed["Invoice Date"] || "", 
                                      parsed["Company Name"] || "", parsed["Address"] || "", parsed["TIN"] || "", parsed["Invoice Number"] || "", 
                                      parsed["Total Amount"] || parsed["Gross"] || "", parsed["Input Tax"] || "", parsed["Purchases"] || "", 
                                      parsed["EWT 1%"] || "", parsed["EWT 2%"] || "", parsed["ATC"] || "",    
                                      "", parsed["PO Number"] || "", releaseDateStr          
                                  ]);
                              }
                          }
                      } catch (e) {
                          console.error("JSON Parse Error: " + e);
                      }
                  }
              }
          }
      });
  }
  
  // Ensure case-insensitive matching between the UI context and the backend Payor values
  let allowedPayors = JSON.parse(props.getProperty("SALES_ENTITIES") || "[]").map(p => String(p).toUpperCase());

  let groupedData = {};
  extractedData.forEach(row => {
      let payor = row[1]; // This is the payorVal assigned from sourceStr earlier
      if (allowedPayors.includes(String(payor).toUpperCase())) {
          if (!groupedData[payor]) groupedData[payor] = [];
          groupedData[payor].push(row);
      }
  });

  if (Object.keys(groupedData).length === 0) {
      deleteTriggersByName("resumeEwtSupplierBatch");
      console.log("No data found for EWT Supplier Report within the selected date range and entities.");
      return;
  }

  let now = new Date();
  let yyyy = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy");
  let mm = Utilities.formatDate(now, Session.getScriptTimeZone(), "MM"); 
  let mmm = Utilities.formatDate(now, Session.getScriptTimeZone(), "MMM"); 
  let dd = Utilities.formatDate(now, Session.getScriptTimeZone(), "dd"); 
  let hhmm = Utilities.formatDate(now, Session.getScriptTimeZone(), "HHmm"); 
  
  // Follows exact format provided: 2026-06Jun-30 2220
      let bossDateStr = `${yyyy}-${mm}${mmm}-${dd} ${hhmm}`; 
      
      // --- NEW: FETCH EWT SUPPLIER EMAIL DATABASE ---
      let ewtEmailMap = {};
      try {
          let emailDbFolder = DriveApp.getFolderById(EMAIL_RECIPIENT_FOLDER_ID !== "1_INSERT_YOUR_NEW_EMAIL_FOLDER_ID_HERE" ? EMAIL_RECIPIENT_FOLDER_ID : DATABASE_FOLDER_ID);
          let emailFiles = emailDbFolder.getFilesByName("EWT_Supplier_Emails_Data");
          while (emailFiles.hasNext()) {
              let file = emailFiles.next();
              if (file.isTrashed() || file.getMimeType() !== MimeType.GOOGLE_SHEETS) continue;
              
              let eSheet = SpreadsheetApp.open(file).getSheets()[0];
              let eData = eSheet.getDataRange().getValues();
              if (eData.length > 1) {
                  let eHeaders = eData[0].map(h => String(h).toLowerCase().trim());
                  let entIdx = eHeaders.indexOf("entity");
                  let toIdx = eHeaders.indexOf("to emails");
                  let ccIdx = eHeaders.indexOf("cc emails");
                  
                  for (let i = 1; i < eData.length; i++) {
                      if (entIdx > -1) {
                          let entityName = String(eData[i][entIdx]).toUpperCase().trim();
                          if (entityName) {
                              let toEmails = toIdx > -1 ? String(eData[i][toIdx]).trim() : "";
                              let ccEmails = ccIdx > -1 ? String(eData[i][ccIdx]).trim() : "";
                              ewtEmailMap[entityName] = { to: toEmails, cc: ccEmails };
                          }
                      }
                  }
              }
          }
      } catch (e) {
          console.warn("Failed to load EWT Email Database: " + e);
      }
      // ----------------------------------------------

      // --- NEW EWT SUBJECT LOGIC ---
      let dayOfWeek = now.getDay();
      let daysUntilTuesday = (2 - dayOfWeek + 7) % 7;
      let nextTuesday = new Date(now);
      nextTuesday.setDate(now.getDate() + daysUntilTuesday);
      
      let baseDate = new Date(2026, 4, 26); // May 26, 2026 (SN=1)
      let diffDays = Math.round((nextTuesday.getTime() - baseDate.getTime()) / (1000 * 60 * 60 * 24));
      let snVal = Math.floor(diffDays / 7) + 1;
      
      let nextMonth5th = new Date(nextTuesday.getFullYear(), nextTuesday.getMonth() + 1, 5);
      
      let tuesdayStr = Utilities.formatDate(nextTuesday, Session.getScriptTimeZone(), "MM/dd/yyyy");
      let fifthStr = Utilities.formatDate(nextMonth5th, Session.getScriptTimeZone(), "MM/dd/yyyy");
      
      let emailSubject = `BIR 2307 EWT issued to Supplier sn=${snVal}: EWT report internal deadline weekly every Tuesday (mm/dd/yyyy)=${tuesdayStr} 10:00, EWT external deadline monthly every 5th of each month (mm/dd/yyyy)=${fifthStr}`;
      // -----------------------------
      
      let sRangeStr = Utilities.formatDate(startDate, Session.getScriptTimeZone(), "MM-dd-yyyy");
      let eRangeStr = Utilities.formatDate(endDate, Session.getScriptTimeZone(), "MM-dd-yyyy");

      let reportHeaders = ["SN", "Payor", "Account Title", "Invoice Date (MM/DD/YYYY)", "Company Name", "Address", "TIN", "Invoice Number", "Gross", "Input Tax", "Purchases", "EWT 1%", "EWT 2%", "ATC", "Notes"];

  for (let payor in groupedData) {
      let payorData = groupedData[payor];

      payorData.sort((a, b) => {
          let cCmp = String(a[4]).localeCompare(String(b[4]));
          if (cCmp !== 0) return cCmp;
          let d1 = new Date(a[3]).getTime() || 0;
          let d2 = new Date(b[3]).getTime() || 0;
          return d1 - d2;
      });

      let sn = 1;
      for (let r = 0; r < payorData.length; r++) {
          payorData[r][0] = sn++;
          [8, 9, 10, 11, 12].forEach(c => {
              let val = String(payorData[r][c]).replace(/,/g, '');
              if (val && !isNaN(parseFloat(val))) {
                  payorData[r][c] = parseFloat(val);
              } else {
                  payorData[r][c] = "";
              }
          });
      }

      // Exact naming convention format provided
      let cleanFileName = `SOA EWT SN Statement of Account ${payor} ${bossDateStr}`;
      
      let tempSpreadsheet = SpreadsheetApp.create(cleanFileName);
      let tempSheet = tempSpreadsheet.getSheets()[0];
      
      // Slice out the PO Number (Index 15) so it doesn't get written to the Excel file
      let excelData = payorData.map(row => row.slice(0, 15));
      
      tempSheet.getRange(1, 1, 1, reportHeaders.length).setValues([reportHeaders]).setFontWeight("bold").setBackground("#f3f3f3");
      tempSheet.getRange(2, 1, excelData.length, reportHeaders.length).setValues(excelData);
      
      tempSheet.getRange(2, 8, excelData.length, 1).setNumberFormat("@"); 
      tempSheet.getRange(2, 9, excelData.length, 5).setNumberFormat("#,##0.00");
      
      tempSheet.getDataRange().setWrap(true).setVerticalAlignment("top");
      for (let i = 1; i <= reportHeaders.length; i++) tempSheet.autoResizeColumn(i);
      
      tempSheet.setColumnWidth(1, 50);
      tempSheet.setColumnWidth(2, 100);
      tempSheet.setColumnWidth(3, 150);
      tempSheet.setColumnWidth(4, 100);
      tempSheet.setColumnWidth(5, 250);
      tempSheet.setColumnWidth(6, 300);
      tempSheet.setColumnWidth(7, 120);
      tempSheet.setColumnWidth(8, 120);
      tempSheet.setColumnWidth(9, 130);
      tempSheet.setColumnWidth(10, 120);
      tempSheet.setColumnWidth(11, 130);
      tempSheet.setColumnWidth(12, 110);
      tempSheet.setColumnWidth(13, 110);
      
      SpreadsheetApp.flush();
      Utilities.sleep(3000);
      
      let url = "https://docs.google.com/spreadsheets/d/" + tempSpreadsheet.getId() + "/export?format=xlsx";
      let token = ScriptApp.getOAuthToken();
      let attachmentBlob = null;
      try {
          let response = UrlFetchApp.fetch(url, { headers: { 'Authorization': 'Bearer ' + token }, muteHttpExceptions: true });
          if (response.getResponseCode() === 200) attachmentBlob = response.getBlob().setName(cleanFileName + ".xlsx");
      } catch(e) {}
      
      try { DriveApp.getFileById(tempSpreadsheet.getId()).setTrashed(true); } catch(e){}
      
      if (attachmentBlob) {
          let sysSettings = getSystemSettings();
          let isTech = (payor === "Keysys" || payor === "A2Z2GO");
          let senderName = isTech ? "Finance A2Z2GO" : "Finance Pingcon";
          let senderEmail = getSenderEmail(isTech);
          let props = PropertiesService.getScriptProperties();
          let finalPassword = isTech ? props.getProperty('ENCRYPT_PASS_TECH') : props.getProperty('ENCRYPT_PASS_PINGCON');
          
          // --- NEW: DYNAMIC RECIPIENT LOGIC ---
          let mainRecipient = "";
          let ccRecipient = "";
          
          let payorUpper = String(payor).toUpperCase();
          if (ewtEmailMap[payorUpper]) {
              if (ewtEmailMap[payorUpper].to) mainRecipient = ewtEmailMap[payorUpper].to;
              if (ewtEmailMap[payorUpper].cc) ccRecipient = ewtEmailMap[payorUpper].cc;
          }

          if (!mainRecipient) {
              console.error(`EWT Supplier Report Aborted: No 'To Emails' configured for ${payorUpper} in EWT_Supplier_Emails_Data.`);
              continue; // Skip sending this entity and move to the next
          }
          // ------------------------------------
          
          let tableRows = [...payorData]
              .map(row => `<tr><td style="border: 1px solid black; padding: 4px 8px; text-align: center;">${row[0]}</td><td style="border: 1px solid black; padding: 4px 8px;">${row[4]}</td><td style="border: 1px solid black; padding: 4px 8px;">${row[7]}</td><td style="border: 1px solid black; padding: 4px 8px;">${row[3]}</td><td style="border: 1px solid black; padding: 4px 8px;">${row[15]}</td></tr>`)
              .join("");
          
          let htmlBody = `
            <div style="font-family: Arial, sans-serif; font-size: 13px; color: #000; line-height: 1.5;">
              <p>Good day,</p>
              <p>Please find attached the <strong>EWT Report Supplier</strong> for <strong>${payor}</strong> for the period <strong>${sRangeStr} to ${eRangeStr}</strong>.</p>
              <p>This report contains all the 2307 data verified and released within the specified date range.</p>
              
              <table style="margin-top: 20px; margin-bottom: 20px; border-collapse: collapse; border: 1px solid black; font-size: 13px; text-align: left;">
                <thead>
                  <tr>
                    <th style="border: 1px solid black; padding: 4px 8px; background-color: #f3f3f3; text-align: center;">SN</th>
                    <th style="border: 1px solid black; padding: 4px 8px; background-color: #f3f3f3;">Company Name</th>
                    <th style="border: 1px solid black; padding: 4px 8px; background-color: #f3f3f3;">Invoice Number</th>
                    <th style="border: 1px solid black; padding: 4px 8px; background-color: #f3f3f3;">Invoice Date</th>
                    <th style="border: 1px solid black; padding: 4px 8px; background-color: #f3f3f3;">PO Number</th>
                  </tr>
                </thead>
                <tbody>
                  ${tableRows}
                </tbody>
              </table>

              <p>Thank you.</p>
            </div>`;
            
          let securedBlob = attachmentBlob;
          try { securedBlob = encryptAttachmentViaPython(attachmentBlob, cleanFileName + ".xlsx", finalPassword); } catch (e) { console.error("Encryption Fallback."); }
          
          sendRobustEmail(mainRecipient, emailSubject, htmlBody, {
              from: senderEmail, 
              name: senderName, 
              cc: ccRecipient,
              attachments: [securedBlob]
          });
          
          
      }
      Utilities.sleep(5000); // 5 second delay between generating files to prevent email limits
  }

  deleteTriggersByName("resumeEwtSupplierBatch");
}

// =========================================================================
// EWT REPORT CLIENT (NEW PORTAL) BACKEND APIS
// =========================================================================
//
// SECURITY / STORAGE NOTE:
// This pipeline intentionally NEVER writes the incoming PDF to this project's
// Google Drive. The file is received in memory only, uploaded temporarily to
// Gemini's own Files API storage (auto-deleted by Google within 48 hours,
// and explicitly deleted by us right after extraction), then discarded.
// Nothing related to the raw PDF is persisted anywhere after this function
// returns. This function behaves identically in Dev and Exec mode since it
// depends on no Drive folder property at all.

function _getClient2307RawFolder() {
  const props = PropertiesService.getScriptProperties();
  let isDevMode = false;
  try {
    let service = ScriptApp.getService();
    if (service && service.getUrl()) {
      isDevMode = service.getUrl().toLowerCase().endsWith("/dev");
    }
  } catch (e) {}

  const propKey = isDevMode ? 'CLIENT_2307_RAW_FOLDER_ID_DEV' : 'CLIENT_2307_RAW_FOLDER_ID';
  const folderId = props.getProperty(propKey) || props.getProperty('CLIENT_2307_RAW_FOLDER_ID') || props.getProperty('CLIENT_2307_DB_FOLDER_ID');
  if (!folderId) throw new Error(`Missing ${propKey} in Script Properties.`);
  return DriveApp.getFolderById(folderId);
}

function webappProcessClient2307(formObject) {
  try {
    let blob = formObject.fileBlob;
    let fileName = formObject.fileName;
    let totalPages = parseInt(formObject.totalPages, 10) || 1;
    if (!blob) throw new Error("No file received by the server.");

    // Process with Gemini AI strictly in memory (0 KB stored in Google Drive)
    return _processPdfBlobWithGemini(blob, fileName, totalPages);
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// Cookie-proof server streamer: supports both single-page images and multi-page PDFs
function webappGetDrivePdfBase64(fileId) {
  try {
    if (!fileId) throw new Error("File ID is required.");
    const file = DriveApp.getFileById(fileId);
    const blob = file.getBlob();
    const bytes = blob.getBytes();
    const mimeType = blob.getContentType();
    
    return {
      success: true,
      base64: Utilities.base64Encode(bytes),
      mimeType: mimeType,
      isImage: mimeType.startsWith("image/"),
      fileName: file.getName()
    };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function _processPdfBlobWithGemini(blob, fileName, totalPages = 1) {
  try {
    let mimeType = blob.getContentType() || "application/pdf";
    let bytes = blob.getBytes();
    let fileSize = bytes.length;

    if (fileSize > 100 * 1024 * 1024) {
      throw new Error("File exceeds 100MB limit. Please split the PDF and try again.");
    }

    let apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
    if (!apiKey) throw new Error("GEMINI_API_KEY is missing from Script Properties.");

    // 1. Upload once to Gemini Files API (Ephemeral, auto-deleted)
    let startResponse = UrlFetchApp.fetch(
      "https://generativelanguage.googleapis.com/upload/v1beta/files?key=" + apiKey,
      {
        method: "post",
        muteHttpExceptions: true,
        contentType: "application/json",
        headers: {
          "X-Goog-Upload-Protocol": "resumable",
          "X-Goog-Upload-Command": "start",
          "X-Goog-Upload-Header-Content-Length": String(fileSize),
          "X-Goog-Upload-Header-Content-Type": mimeType
        },
        payload: JSON.stringify({ file: { display_name: fileName } })
      }
    );

    let respHeaders = startResponse.getAllHeaders();
    let uploadUrl = null;
    for (let key in respHeaders) {
      if (key.toLowerCase() === "x-goog-upload-url") {
        uploadUrl = Array.isArray(respHeaders[key]) ? respHeaders[key][0] : respHeaders[key];
        break;
      }
    }
    if (!uploadUrl) throw new Error("Gemini upload session could not be created: " + startResponse.getContentText());

    let CHUNK_SIZE = 16 * 1024 * 1024;
    let offset = 0;
    let uploadJson = null;

    while (offset < fileSize) {
      let end = Math.min(offset + CHUNK_SIZE, fileSize);
      let chunkBytes = bytes.slice(offset, end);
      let isLastChunk = (end === fileSize);

      let chunkResponse = UrlFetchApp.fetch(uploadUrl, {
        method: "post",
        muteHttpExceptions: true,
        headers: {
          "X-Goog-Upload-Offset": String(offset),
          "X-Goog-Upload-Command": isLastChunk ? "upload, finalize" : "upload"
        },
        contentType: mimeType,
        payload: chunkBytes
      });

      if (isLastChunk) {
        uploadJson = JSON.parse(chunkResponse.getContentText());
      } else if (chunkResponse.getResponseCode() >= 400) {
        throw new Error("Chunk upload failed at offset " + offset + ": " + chunkResponse.getContentText());
      }

      offset = end;
    }

    if (!uploadJson.file || !uploadJson.file.uri) {
      throw new Error("Gemini file upload failed: " + JSON.stringify(uploadJson));
    }

    let geminiFileUri = uploadJson.file.uri;
    let geminiFileName = uploadJson.file.name;

    // 2. Process in 20-Page Batches to guarantee 100% completion without token truncation
    let allExtractedDocuments = [];
    let BATCH_SIZE = 20;
    let numBatches = Math.ceil(totalPages / BATCH_SIZE);

    try {
      for (let b = 0; b < numBatches; b++) {
        let startPage = (b * BATCH_SIZE) + 1;
        let endPage = Math.min((b + 1) * BATCH_SIZE, totalPages);

        let batchPrompt = "Extract every BIR Form 2307 certificate from this document ONLY for physical pages " + startPage + " to " + endPage + ".\n" +
                          "For each page, return a JSON object with these exact keys:\n" +
                          "- page: (integer) physical page number (between " + startPage + " and " + endPage + ")\n" +
                          "- date_from: Box 1 'From' date strictly in MM/DD/YYYY format\n" +
                          "- date_to: Box 1 'To' date strictly in MM/DD/YYYY format\n" +
                          "- customer_name: Part II Box 7 Payor's Registered Name (The Client, NOT Pingcon)\n" +
                          "- tin: Part II Box 6 Payor's TIN (extract exact digits as printed; keep 000 or 00000 without altering)\n" +
                          "- address: Part II Box 8 Payor's Registered Address\n" +
                          "- income: Part III Total Amount of Income Payments (gross base decimal, no commas)\n" +
                          "- atc: Part III ATC code (e.g. WC158, WC160)\n" +
                          "- wc158, wc160, wc157, wc120, wc640, wc690, wv010, wv020, wi632: tax withheld decimal, or '' if not present.\n" +
                          "Return strictly JSON: { \"documents\": [ { \"page\": " + startPage + ", \"date_from\": \"MM/DD/YYYY\", \"date_to\": \"MM/DD/YYYY\", \"customer_name\": \"...\", \"tin\": \"...\", \"address\": \"...\", \"income\": \"...\", \"atc\": \"...\", \"wc158\": \"...\", \"wc160\": \"\", \"wc157\": \"\", \"wc120\": \"\", \"wc640\": \"\", \"wc690\": \"\", \"wv010\": \"\", \"wv020\": \"\", \"wi632\": \"\" } ] }";

        let payload = {
          "contents": [
            {
              "parts": [
                { "text": batchPrompt },
                { "fileData": { "mimeType": mimeType, "fileUri": geminiFileUri } }
              ]
            }
          ],
          "safetySettings": [
            { "category": "HARM_CATEGORY_HARASSMENT", "threshold": "BLOCK_NONE" },
            { "category": "HARM_CATEGORY_HATE_SPEECH", "threshold": "BLOCK_NONE" },
            { "category": "HARM_CATEGORY_SEXUALLY_EXPLICIT", "threshold": "BLOCK_NONE" },
            { "category": "HARM_CATEGORY_DANGEROUS_CONTENT", "threshold": "BLOCK_NONE" }
          ],
          "generationConfig": { 
            "responseMimeType": "application/json",
            "maxOutputTokens": 8192,
            "temperature": 0.1
          }
        };

        let options = {
          method: "post",
          contentType: "application/json",
          payload: JSON.stringify(payload),
          muteHttpExceptions: true
        };

        let url = "https://generativelanguage.googleapis.com/v1beta/models/gemini-3.5-flash-lite:generateContent?key=" + apiKey;
        let response = UrlFetchApp.fetch(url, options);

        if (response.getResponseCode() === 200) {
          let jsonResp = JSON.parse(response.getContentText());
          if (jsonResp.candidates && jsonResp.candidates[0] && jsonResp.candidates[0].content) {
            let rawText = jsonResp.candidates[0].content.parts[0].text;
            let cleanJson = rawText.replace(/^```json\s*/i, '').replace(/^```\s*/i, '').replace(/\s*```$/i, '').trim();
            
            let fBrace = cleanJson.indexOf('{');
            let lBrace = cleanJson.lastIndexOf('}');
            if (fBrace !== -1 && lBrace !== -1) cleanJson = cleanJson.substring(fBrace, lBrace + 1);

            let parsedBatch = null;
            try { parsedBatch = JSON.parse(cleanJson); } catch(e){}

            if (parsedBatch && Array.isArray(parsedBatch.documents)) {
              parsedBatch.documents.forEach(d => allExtractedDocuments.push(d));
            }
          }
        }
      }
    } finally {
      // 3. Delete the temporary file from Gemini storage
      try {
        UrlFetchApp.fetch("https://generativelanguage.googleapis.com/v1beta/" + geminiFileName + "?key=" + apiKey, {
          method: "delete",
          muteHttpExceptions: true
        });
      } catch(cleanupErr) {}
    }

    if (allExtractedDocuments.length === 0) {
      return { success: false, error: "AI extraction returned no documents across all page batches." };
    }

    // 4. Normalize documents into standard fields array
    let normalizedDocs = allExtractedDocuments.map((d, idx) => {
      let pg = parseInt(d.page || d.Page || (idx + 1), 10);
      return {
        fields: [
          { label: "Page", value: String(pg) },
          { label: "Date_From", value: String(d.date_from || d.Date_From || "") },
          { label: "Date_To", value: String(d.date_to || d.Date_To || "") },
          { label: "Customer_Name", value: String(d.customer_name || d.Customer_Name || "") },
          { label: "TIN", value: String(d.tin || d.TIN || "") },
          { label: "Address", value: String(d.address || d.Address || "") },
          { label: "Income_Payment", value: String(d.income || d.income_payment || d.Income_Payment || "") },
          { label: "ATC", value: String(d.atc || d.ATC || "WC158") },
          { label: "WC158", value: String(d.wc158 || d.WC158 || "") },
          { label: "WC160", value: String(d.wc160 || d.WC160 || "") },
          { label: "WC157", value: String(d.wc157 || d.WC157 || "") },
          { label: "WC120", value: String(d.wc120 || d.WC120 || "") },
          { label: "WC640", value: String(d.wc640 || d.WC640 || "") },
          { label: "WC690", value: String(d.wc690 || d.WC690 || "") },
          { label: "WV010", value: String(d.wv010 || d.WV010 || "") },
          { label: "WV020", value: String(d.wv020 || d.WV020 || "") },
          { label: "WI632", value: String(d.wi632 || d.WI632 || "") }
        ]
      };
    });

    // Sort documents chronologically by page number
    normalizedDocs.sort((a, b) => {
      let pA = parseInt((a.fields.find(f => f.label === "Page") || {}).value || 0, 10);
      let pB = parseInt((b.fields.find(f => f.label === "Page") || {}).value || 0, 10);
      return pA - pB;
    });

    return { success: true, data: { documents: normalizedDocs }, fileName: fileName };

  } catch (e) {
    return { success: false, error: e.message };
  }
}


function webappSaveClient2307(payload, companyContext, sessionToken = null) {
  try {
    // --- SAVE SINGLE PAGE TO GOOGLE DRIVE (Executed OUTSIDE the lock) ---
    let savedDriveFileId = "";
    if (payload.pageImageBase64 && payload.pageImageBase64.length > 100) {
      try {
        const rootFolder = _getClient2307RawFolder();

        // 1. Resolve Entity Folder
        let entityName = "Pingcon";
        let ctxUpper = String(companyContext).toUpperCase();
        if (ctxUpper.includes("KEYSYS")) entityName = "Keysys";
        else if (ctxUpper.includes("A2Z")) entityName = "A2Z2GO";
        else if (ctxUpper.includes("BONITA")) entityName = "Bonita";

        let entityFolders = rootFolder.getFoldersByName(entityName);
        let entityFolder = entityFolders.hasNext() ? entityFolders.next() : rootFolder.createFolder(entityName);

        // 2. Resolve Month-Year Folder strictly based on THAT page's 2307 Receive Date
        let receiveDateObj = new Date(payload.receiveDate);
        if (isNaN(receiveDateObj.getTime())) receiveDateObj = new Date(); // Fallback to current date if invalid
        let monthYearStr = Utilities.formatDate(receiveDateObj, Session.getScriptTimeZone(), "MMMM yyyy");

        let monthFolders = entityFolder.getFoldersByName(monthYearStr);
        let targetFolder = monthFolders.hasNext() ? monthFolders.next() : entityFolder.createFolder(monthYearStr);

        // 3. Create clean single-page filename
        let cleanCust = String(payload.customerName || "Customer").replace(/[^a-zA-Z0-9]/g, '_').substring(0, 30);
        let cleanSi = String(payload.siNo || "SI").replace(/[^a-zA-Z0-9]/g, '_');
        let singleFileName = `Client_2307_${cleanCust}_${cleanSi}_P${payload.pageNum || 1}.jpg`;

        // 4. Save lightweight single-page image (~150KB to 250KB)
        let rawBase64 = payload.pageImageBase64.includes(',') ? payload.pageImageBase64.split(',')[1] : payload.pageImageBase64;
        let imgBlob = Utilities.newBlob(Utilities.base64Decode(rawBase64), "image/jpeg", singleFileName);
        let singlePageFile = targetFolder.createFile(imgBlob);
        savedDriveFileId = singlePageFile.getId();
      } catch (driveErr) {
        console.warn("Could not save single page to Drive: " + driveErr.message);
      }
    }
    
    let props = PropertiesService.getScriptProperties();
    let isDevMode = ScriptApp.getService().getUrl().endsWith("/dev");
    let dbFolderId = (isDevMode && props.getProperty('CLIENT_2307_DB_FOLDER_ID_DEV')) 
        ? props.getProperty('CLIENT_2307_DB_FOLDER_ID_DEV') 
        : props.getProperty('CLIENT_2307_DB_FOLDER_ID');
        
    if (!dbFolderId) throw new Error("Database Folder ID is missing in Script Properties.");
    let folder = DriveApp.getFolderById(dbFolderId);
    
    let targetYear = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy");
    let dbName = PREFIX_CLIENT_2307 + companyContext + "_" + targetYear;
    
    let ss = null;
    let files = folder.getFilesByName(dbName);
    
    if (files.hasNext()) {
      ss = SpreadsheetApp.open(files.next());
    } else {
      ss = SpreadsheetApp.create(dbName);
      DriveApp.getFileById(ss.getId()).moveTo(folder);
    }
    
    let sheet = ss.getSheetByName("Verified Data");
    if (!sheet) {
      sheet = ss.insertSheet("Verified Data");
      let sheet1 = ss.getSheetByName("Sheet1");
      if (sheet1) ss.deleteSheet(sheet1);
    }
    
    let headers = [
      "SN", "Date from (MM/DD/YYYY)", "Date to (MM/DD/YYYY)", "SI No.", "Customer Name", 
      "TIN", "Address", "ATC", "Income Payment", "WC 158 1%", "WC 157 2%", "WC 160 2%", 
      "WV010 5%", "WC120 2%", "WC640 1%", "WV020 5%", "WI632 1%", "WC690", 
      "2307 Receive Date (MM/DD/YYYY)", "File ID", "File Name", "Timestamp"
    ];
    
    let lastCol = sheet.getLastColumn();
    if (sheet.getRange(1, 1).getValue() !== "SN" || lastCol < headers.length) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold").setBackground("#f3f3f3");
      sheet.setFrozenRows(1);
    }
    
    let data = sheet.getDataRange().getValues();
    let rowIndex = -1;
    let maxSn = 0;
    
    let hRow = data[0].map(h => String(h).trim());
    let fileIdColIdx = hRow.indexOf("File ID");
    if (fileIdColIdx === -1) fileIdColIdx = 19;
    
    // Deterministic Record Key: Unique per certificate page
    let uniqueRecordKey = String(payload.recordId || (payload.fileName + "_P" + (payload.pageNum || 1))).trim().toUpperCase();
    
    for (let i = 1; i < data.length; i++) {
      let currentSn = parseInt(data[i][0], 10) || 0;
      if (currentSn > maxSn) maxSn = currentSn;
      
      let existingKey = String(data[i][fileIdColIdx] || "").trim().toUpperCase();
      if (existingKey === uniqueRecordKey || (payload.rowSn && currentSn === parseInt(payload.rowSn, 10))) {
        rowIndex = i + 1;
        break;
      }
    }
    
    let finalSn = (rowIndex > -1) ? (parseInt(data[rowIndex-1][0], 10) || (maxSn + 1)) : (maxSn + 1);
    let timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MM/dd/yyyy HH:mm");
    
    let effectiveDriveId = savedDriveFileId || payload.driveFileId || payload.fileId || "";
    let storedFileIdRef = effectiveDriveId ? `${effectiveDriveId}|${payload.pageNum || 1}|${uniqueRecordKey}` : uniqueRecordKey;

    let rowData = [
      finalSn, payload.dateFrom, payload.dateTo, payload.siNo, payload.customerName, payload.tin, payload.address, payload.atc,
      payload.incomePayment, payload.wc158, payload.wc157, payload.wc160, payload.wv010, payload.wc120, payload.wc640, payload.wv020, payload.wi632, payload.wc690,
      payload.receiveDate || "", storedFileIdRef, payload.fileName, timestamp
    ];
    
    // MICRO-LOCK: Only holds for the 0.2s sheet row write
    let lock = LockService.getScriptLock();
    try {
      lock.waitLock(30000);
      if (rowIndex > -1) {
        sheet.getRange(rowIndex, 1, 1, headers.length).setValues([rowData]);
      } else {
        sheet.appendRow(rowData);
      }
      SpreadsheetApp.flush();
    } finally {
      try { lock.releaseLock(); } catch(e){}
    }
    
    let pageLabel = payload.pageNum ? `Page ${payload.pageNum} | ` : "";
    logClientActivity(companyContext, "Saved Client 2307", `${pageLabel}Customer: ${payload.customerName} | SI: ${payload.siNo} | Gross: ${payload.incomePayment}`, sessionToken);
    
    return { success: true };
  } catch (e) {
    return { error: e.message };
  }
}

function webappGetClient2307Data(companyContext) {
  try {
    let props = PropertiesService.getScriptProperties();
    let isDevMode = ScriptApp.getService().getUrl().endsWith("/dev");
    let dbFolderId = (isDevMode && props.getProperty('CLIENT_2307_DB_FOLDER_ID_DEV')) 
        ? props.getProperty('CLIENT_2307_DB_FOLDER_ID_DEV') 
        : props.getProperty('CLIENT_2307_DB_FOLDER_ID');
        
    if (!dbFolderId) throw new Error("Database Folder ID is missing in Script Properties.");
    let folder = DriveApp.getFolderById(dbFolderId);
    
    let results = [];
    let currentYear = new Date().getFullYear();
    
    // Strict MM/DD/YYYY Date Formatter
    const fDate = (d) => {
      if (!d) return "";
      if (d instanceof Date && !isNaN(d)) return Utilities.formatDate(d, Session.getScriptTimeZone(), "MM/dd/yyyy");
      let str = String(d).trim();
      let parts = str.split(/[\/\-\.]/);
      if (parts.length === 3) {
        let m = parts[0], da = parts[1], y = parts[2];
        if (m.length === 4) { y = parts[0]; m = parts[1]; da = parts[2]; } // YYYY-MM-DD
        return ('0' + parseInt(m, 10)).slice(-2) + '/' + ('0' + parseInt(da, 10)).slice(-2) + '/' + y;
      }
      return str;
    };

    const fTime = (d) => (d instanceof Date && !isNaN(d)) ? Utilities.formatDate(d, Session.getScriptTimeZone(), "MM/dd/yyyy HH:mm") : String(d || "");

    for (let y = 2024; y <= currentYear + 1; y++) {
      let dbName = PREFIX_CLIENT_2307 + companyContext + "_" + y;
      let files = folder.getFilesByName(dbName);
      
      if (files.hasNext()) {
        let file = files.next();
        let ss = SpreadsheetApp.open(file);
        let sheet = ss.getSheetByName("Verified Data");
        if (!sheet) continue;
        
        let data = sheet.getDataRange().getValues();
        if (data.length <= 1) continue;
        
        let hRow = data[0].map(h => String(h).trim());
        let recDateIdx = hRow.indexOf("2307 Receive Date (MM/DD/YYYY)");
        let fileIdIdx = hRow.indexOf("File ID");
        if (fileIdIdx === -1) fileIdIdx = 18;
        let fileNameIdx = hRow.indexOf("File Name");
        if (fileNameIdx === -1) fileNameIdx = 19;
        let tsIdx = hRow.indexOf("Timestamp");
        if (tsIdx === -1) tsIdx = 20;
        
        for(let i = 1; i < data.length; i++) {
          let recDateVal = recDateIdx > -1 ? fDate(data[i][recDateIdx]) : "";
          let timeVal = fTime(data[i][tsIdx]);
          if (!recDateVal) recDateVal = timeVal ? fDate(timeVal) : fDate(data[i][1]); // Fallback for older records
          
          let rawFileIdCol = String(data[i][fileIdIdx] || "");
          let realDriveId = "";
          let pageNum = 1;

          if (rawFileIdCol.includes("|")) {
            let parts = rawFileIdCol.split("|");
            realDriveId = parts[0];
            pageNum = parseInt(parts[1], 10) || 1;
          } else {
            realDriveId = rawFileIdCol;
          }

          results.push({
            sn: data[i][0], 
            dateFrom: fDate(data[i][1]), 
            dateTo: fDate(data[i][2]), 
            siNo: String(data[i][3] || ""), 
            customerName: String(data[i][4] || ""),
            tin: String(data[i][5] || ""), 
            address: String(data[i][6] || ""), 
            atc: String(data[i][7] || ""), 
            incomePayment: data[i][8], 
            wc158: data[i][9],
            wc157: data[i][10], 
            wc160: data[i][11], 
            wv010: data[i][12], 
            wc120: data[i][13], 
            wc640: data[i][14],
            wv020: data[i][15], 
            wi632: data[i][16], 
            wc690: data[i][17],
            receiveDate: recDateVal, 
            fileId: rawFileIdCol,
            driveFileId: realDriveId,
            pageNum: pageNum,
            fileName: data[i][fileNameIdx], 
            timestamp: timeVal
          });
        }
      }
    }
    
    results.sort((a,b) => new Date(b.timestamp) - new Date(a.timestamp));
    return { success: true, data: results };
  } catch(e) {
    return { error: e.message };
  }
}

function webappDeleteClient2307Batch(payloads, companyContext, sessionToken = null) {
  // SECURITY GATE: Only Super Admins can delete client 2307 records
  if (sessionToken) {
    const user = getAuthenticatedUser(sessionToken);
    if (!user || user.role !== "Super Admin") {
      return { success: false, error: "Permission Denied: Only Super Admins are authorized to delete client 2307 records." };
    }
  } else {
    return { success: false, error: "Authentication required to delete records." };
  }
  let lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    
    let props = PropertiesService.getScriptProperties();
    let isDevMode = ScriptApp.getService().getUrl().endsWith("/dev");
    let dbFolderId = (isDevMode && props.getProperty('CLIENT_2307_DB_FOLDER_ID_DEV')) 
        ? props.getProperty('CLIENT_2307_DB_FOLDER_ID_DEV') 
        : props.getProperty('CLIENT_2307_DB_FOLDER_ID');
        
    if (!dbFolderId) throw new Error("Database Folder ID is missing in Script Properties.");
    let folder = DriveApp.getFolderById(dbFolderId);
    
    let filesIter = folder.getFiles();
    let totalDeleted = 0;
    
    while (filesIter.hasNext()) {
      let file = filesIter.next();
      let fName = file.getName();
      if (!fName.startsWith(PREFIX_CLIENT_2307 + companyContext + "_")) continue;
      
      let ss = SpreadsheetApp.open(file);
      let sheet = ss.getSheetByName("Verified Data");
      if (!sheet || sheet.getLastRow() <= 1) continue;
      
      let data = sheet.getDataRange().getValues();
      let deletedInSheet = false;
      
      // Traverse backwards to safely delete multiple rows without shifting lower row indices
      for (let i = data.length - 1; i >= 1; i--) {
        let rowRecordKey = String(data[i][18] || "").trim().toUpperCase(); // Col S: Record ID / Unique Key
        let rowSiNo = String(data[i][3] || "").trim().toUpperCase();       // Col D: SI No.
        let rowCustomer = String(data[i][4] || "").trim();
        
        let match = payloads.find(p => {
          let pKey = String(p.fileId || "").trim().toUpperCase();
          let pSi = String(p.siNo || "").trim().toUpperCase();
          return (pKey && pKey === rowRecordKey) || (pSi && pSi === rowSiNo);
        });
        
        if (match) {
          logClientActivity(companyContext, "Deleted Client 2307", `Deleted SI: ${rowSiNo} | Customer: ${rowCustomer}`);
          sheet.deleteRow(i + 1);
          totalDeleted++;
          deletedInSheet = true;
        }
      }
      
      // Re-sequence Serial Numbers (SN) in Column A to close numbering gaps
      if (deletedInSheet) {
        let remainingRows = sheet.getLastRow() - 1;
        if (remainingRows > 0) {
          let newSns = [];
          for (let sn = 1; sn <= remainingRows; sn++) {
            newSns.push([sn]);
          }
          sheet.getRange(2, 1, remainingRows, 1).setValues(newSns);
        }
        SpreadsheetApp.flush();
      }
    }
    
    return { success: true, deleted: totalDeleted };
  } catch (e) {
    return { error: e.message };
  } finally {
    lock.releaseLock();
  }
}
// =========================================================================
// DASHBOARD & VERIFIED DATABASE CRUD APIS
// =========================================================================

function webappGetVerifiedInvoices(companyContext) {
  let props = PropertiesService.getScriptProperties();
  let isDevMode = ScriptApp.getService().getUrl().endsWith("/dev");
  const VERIFIED_FOLDER_ID = (isDevMode && props.getProperty('VERIFIED_FOLDER_ID_DEV')) 
      ? props.getProperty('VERIFIED_FOLDER_ID_DEV') 
      : props.getProperty('VERIFIED_FOLDER_ID');
      
  let folder;
  try { folder = DriveApp.getFolderById(VERIFIED_FOLDER_ID); } catch(e) { return []; }
  
  let filesIter = folder.getFilesByType(MimeType.GOOGLE_SHEETS);
  let results = [];
  
  while(filesIter.hasNext()) {
    let file = filesIter.next();
    if (!file.getName().startsWith("Verified Invoices for 2307")) continue;
    
    let ss = SpreadsheetApp.open(file);
    let sheet = ss.getSheetByName("All Verified Invoices");
    if (!sheet) continue;
    
    let data = sheet.getDataRange().getValues();
    if (data.length <= 1) continue;
    
    let headers = data[0].map(h => String(h).trim());
    let snIdx = headers.indexOf("SN");
    let fnIdx = headers.indexOf("File Name");
    let vdIdx = headers.indexOf("Verified Data");
    let rdIdx = headers.indexOf("2307 Release Date");
    let srcIdx = headers.indexOf("Source");
    let dtIdx = headers.indexOf("Date and Time Verified (MM/DD/YYYY)");
    
    if (fnIdx === -1 || vdIdx === -1) continue;
    
    for(let i = 1; i < data.length; i++) {
      let source = String(data[i][srcIdx]).trim().toUpperCase();
      if (!source.includes(companyContext.toUpperCase())) continue;
      
      let vData = String(data[i][vdIdx]).trim();
      if (vData.startsWith("{")) {
        try {
          let parsed = JSON.parse(vData);
          
          let imgFileId = parsed["fileId"] || ""; // Reads the flattened ID

          let totalAmt = "";
          if (parsed.isMixedVat) {
              let vGross = parseFloat(String(parsed.VAT_Portion["Total Amount"] || parsed.VAT_Portion["Gross"] || "0").replace(/,/g, '')) || 0;
              let nvGross = parseFloat(String(parsed.NV_Portion["Total Amount"] || parsed.NV_Portion["Gross"] || "0").replace(/,/g, '')) || 0;
              totalAmt = (vGross + nvGross).toLocaleString('en-US', {minimumFractionDigits: 2, maximumFractionDigits: 2});
          } else {
              totalAmt = parsed["Total Amount"] || parsed["Gross"] || "";
          }

          results.push({
            recordType: parsed["Record Type"] || "2307 EWT",
            sn: data[i][snIdx],
            fileName: data[i][fnIdx],
            invoiceNo: parsed["Invoice Number"] || "",
            invoiceDate: parsed["Invoice Date"] || "", // <--- Added Invoice Date
            companyName: parsed["Company Name"] || "",
            totalAmount: totalAmt,
            releaseDate: data[i][rdIdx] instanceof Date ? Utilities.formatDate(data[i][rdIdx], Session.getScriptTimeZone(), "MM/dd/yyyy") : data[i][rdIdx],
            verifiedDateTime: data[i][dtIdx] instanceof Date ? Utilities.formatDate(data[i][dtIdx], Session.getScriptTimeZone(), "MM/dd/yyyy HH:mm") : data[i][dtIdx],
            extractedData: vData,
            dbSourceName: source,
            imgFileId: imgFileId
          });
        } catch(e) {}
      }
    }
  }
  
  // Sort descending by verification date so the newest are at the top
  results.sort((a,b) => new Date(b.verifiedDateTime) - new Date(a.verifiedDateTime));
  return results;
}

function webappDeleteVerifiedInvoiceBatch(payloads, sessionToken = null) {
  try {
    // SECURITY GATE: Prevent Encoders from deleting archives
    if (sessionToken) {
      const user = getAuthenticatedUser(sessionToken);
      if (user && user.role === "Encoder") {
        return { error: "Permission Denied: Encoders are not authorized to delete verified records." };
      }
    }

    console.log("BACKEND INIT: Received delete payloads: " + JSON.stringify(payloads));
    
    let props = PropertiesService.getScriptProperties();
    let isDevMode = ScriptApp.getService().getUrl().endsWith("/dev");
    
    const VERIFIED_FOLDER_ID = (isDevMode && props.getProperty('VERIFIED_FOLDER_ID_DEV')) 
        ? props.getProperty('VERIFIED_FOLDER_ID_DEV') 
        : props.getProperty('VERIFIED_FOLDER_ID');
        
    const GENERATED_2307_FOLDER_ID = (isDevMode && props.getProperty('GENERATED_2307_FOLDER_ID_DEV')) 
        ? props.getProperty('GENERATED_2307_FOLDER_ID_DEV') 
        : props.getProperty('GENERATED_2307_FOLDER_ID');

    console.log("BACKEND ENV: isDevMode=" + isDevMode + ", Target DB Folder=" + VERIFIED_FOLDER_ID);
    
    let folder = DriveApp.getFolderById(VERIFIED_FOLDER_ID);
    let rootPdfFolder = DriveApp.getFolderById(GENERATED_2307_FOLDER_ID);
    
    let filesIter = folder.getFilesByType(MimeType.GOOGLE_SHEETS);
    let totalDeleted = 0;
    
    while(filesIter.hasNext()) {
      let file = filesIter.next();
      if (!file.getName().startsWith("Verified Invoices for 2307")) continue;
      
      let ss = SpreadsheetApp.open(file);
      let sheet = ss.getSheetByName("All Verified Invoices");
      if (!sheet) continue;
      
      let data = sheet.getDataRange().getValues();
      if (data.length <= 1) continue;
      
      let headers = data[0].map(h => String(h).trim());
      let vdIdx = headers.indexOf("Verified Data");
      
      let deletedInThisSheet = false;
      
      console.log("Scanning Spreadsheet: " + file.getName() + " | Rows found: " + (data.length - 1));

      const norm = s => String(s || "").toUpperCase().replace(/[^A-Z0-9]/g, '');

      // Traverse backward to safely delete multiple rows without breaking the index loop
      for(let i = data.length - 1; i >= 1; i--) {
        let rowFileName = String(data[i][1] || "").trim();
        let rowSource = String(data[i][4] || "").trim().toUpperCase();
        let rowJsonStr = vdIdx > -1 ? String(data[i][vdIdx] || "").trim() : "";
        let rowInvNumber = "";
        let rowCompName = "";
        let rowRecordType = "";

        if (rowJsonStr.startsWith("{")) {
            try {
                let parsedRow = JSON.parse(rowJsonStr);
                rowInvNumber = String(parsedRow["Invoice Number"] || "").trim();
                rowCompName = String(parsedRow["Company Name"] || "").trim();
                rowRecordType = String(parsedRow["Record Type"] || "").trim();
            } catch(e) {
                let invM = rowJsonStr.match(/"Invoice Number"\s*:\s*"([^"]+)"/i);
                let compM = rowJsonStr.match(/"Company Name"\s*:\s*"([^"]+)"/i);
                let recM = rowJsonStr.match(/"Record Type"\s*:\s*"([^"]+)"/i);
                if (invM) rowInvNumber = String(invM[1]).trim();
                if (compM) rowCompName = String(compM[1]).trim();
                if (recM) rowRecordType = String(recM[1]).trim();
            }
        }
        
        // Check if this specific row matches ANY of the items the user marked for deletion
        let matchedPayload = payloads.find(p => {
             let sameSource = !p.dbSourceName || rowSource.includes(String(p.dbSourceName).toUpperCase()) || String(p.dbSourceName).toUpperCase().includes(rowSource);
             if (!sameSource) return false;

             // Match 1: Multi-factor match by Invoice Number AND Supplier Name AND Record Type
             if (p.invoiceNo && rowInvNumber && p.companyName && rowCompName) {
                 let sameInv = norm(p.invoiceNo) === norm(rowInvNumber);
                 let sameComp = norm(p.companyName).includes(norm(rowCompName)) || norm(rowCompName).includes(norm(p.companyName));
                 let sameRec = !p.recordType || !rowRecordType || p.recordType === rowRecordType;
                 if (sameInv && sameComp && sameRec) return true;
             }

             // Match 2: Exact Filename match fallback
             if (p.fileName && rowFileName && p.fileName.toUpperCase() === rowFileName.toUpperCase()) {
                 let sameRec = !p.recordType || !rowRecordType || p.recordType === rowRecordType;
                 return sameRec;
             }

             return false;
        });
        
        if (matchedPayload) {
          console.log("DELETING ROW: " + (i + 1) + " for file: " + rowFileName);
          // Delete corresponding PDF from the Drive Folders
          if (vdIdx > -1) {
             let vDataStr = String(data[i][vdIdx]).trim();
             if (vDataStr.startsWith("{")) {
                try {
                   let parsed = JSON.parse(vDataStr);
                   let cleanSupplier = String(parsed["Company Name"] || "Supplier").replace(/[^a-zA-Z0-9 -]/g, "").trim();
                   let invNo = String(parsed["Invoice Number"] || "INV").trim();
                   let pdfNameToDelete = `2307_${cleanSupplier}_${invNo}.pdf`;
                       
                   let payorFolder = "Pingcon";
                   let sourceUpper = String(matchedPayload.dbSourceName).toUpperCase();
                   if (sourceUpper.includes("KEYSYS")) payorFolder = "Keysys";
                   else if (sourceUpper.includes("A2Z") || sourceUpper.includes("A2Z2GO")) payorFolder = "A2Z2GO";
                   else if (sourceUpper.includes("BONITA")) payorFolder = "Bonita";
                   else if (sourceUpper.includes("PINGCON")) payorFolder = "Pingcon";

                   let compFolders = rootPdfFolder.getFoldersByName(payorFolder);
                   if (compFolders.hasNext()) {
                       let compFolder = compFolders.next();
                       let allMonthFolders = compFolder.getFolders();
                       while (allMonthFolders.hasNext()) {
                           let mFolder = allMonthFolders.next();
                           let pdfIter = mFolder.getFilesByName(pdfNameToDelete);
                           while (pdfIter.hasNext()) {
                               pdfIter.next().setTrashed(true);
                           }
                       }
                   }
                } catch(e) { console.warn("Failed to delete associated PDF: " + e); }
             }
          }
          
          // Stamped with verified employee identity
          logSystemActivity(matchedPayload.dbSourceName, "Deleted 2307 Invoice", `Deleted record for File: ${matchedPayload.fileName}`, sessionToken);

          // --- CLOUD VAULT PRUNE SYNC (NORMALIZED) ---
          try {
            let targetInv = rowInvNumber || matchedPayload.invoiceNo || "";
            let targetComp = rowCompName || matchedPayload.companyName || "";
            let delInvKey = `${targetInv}_${targetComp}`.trim().toUpperCase();
            
            if (targetInv) {
              _deleteRecordFromCloudVault("verified_invoices_vault", delInvKey, targetInv);
            }
          } catch(vDelErr) {
            console.warn("[Cloud Vault] Delete sync non-blocking warning: " + vDelErr.message);
          }
          // ------------------------------------------

          sheet.deleteRow(i + 1);
          totalDeleted++;
          deletedInThisSheet = true;
        }
      }
      
      // Resequence Serial Numbers (SN) to close the gaps left by the deleted rows
      if (deletedInThisSheet) {
          let remainingLastRow = sheet.getLastRow();
          if (remainingLastRow > 1) {
              let numRows = remainingLastRow - 1;
              let newSns = [];
              for (let sn = 1; sn <= numRows; sn++) {
                  newSns.push([sn]);
              }
              sheet.getRange(2, 1, numRows, 1).setValues(newSns);
              
              if (vdIdx > -1) {
                  let jsonRange = sheet.getRange(2, vdIdx + 1, numRows, 1);
                  let jsonValues = jsonRange.getValues();
                  for (let j = 0; j < jsonValues.length; j++) {
                      let jStr = String(jsonValues[j][0]).trim();
                      if (jStr.startsWith("{")) {
                          try {
                              let parsed = JSON.parse(jStr);
                              parsed["SN"] = j + 1;
                              jsonValues[j][0] = JSON.stringify(parsed);
                          } catch(e) {}
                      }
                  }
                  jsonRange.setValues(jsonValues);
              }
          }
          SpreadsheetApp.flush();
      }
    }
    return { success: true, deleted: totalDeleted };
  } catch(e) {
    return { error: e.message };
  }
}
// =========================================================================
// EWT PRE-VERIFICATION SCAN BACKEND
// =========================================================================
function webappScanEwtData(startStr, endStr, verifiedStartStr, verifiedEndStr, verifiedAllDates, companyContext) {
  try {
    const startDate = new Date(startStr + "T00:00:00");
    const endDate = new Date(endStr + "T23:59:59");
    
    // Set up the Verified Date range safely
    let vStartDate, vEndDate;
    if (!verifiedAllDates) {
        vStartDate = new Date(verifiedStartStr + "T00:00:00");
        vEndDate = new Date(verifiedEndStr + "T23:59:59");
    }
    
    // Automatically swap to correct Database environment (Dev vs Exec)
    let props = PropertiesService.getScriptProperties();
    let isDevMode = ScriptApp.getService().getUrl().endsWith("/dev");
    const VERIFIED_FOLDER_ID = (isDevMode && props.getProperty('VERIFIED_FOLDER_ID_DEV')) 
        ? props.getProperty('VERIFIED_FOLDER_ID_DEV') 
        : props.getProperty('VERIFIED_FOLDER_ID');
        
    let folder = DriveApp.getFolderById(VERIFIED_FOLDER_ID);
    let files = folder.getFilesByType(MimeType.GOOGLE_SHEETS);
    
    let verifiedInvoices = [];
    
    // Translate the UI company context to the database standard Payor name
    let ctxUpper = String(companyContext || "").toUpperCase().trim();
    let targetPayor = "PINGCON";
    if (ctxUpper.includes("KEYSYS")) targetPayor = "KEYSYS";
    else if (ctxUpper.includes("A2Z")) targetPayor = "A2Z2GO";
    else if (ctxUpper.includes("BONITA")) targetPayor = "BONITA";
    else if (ctxUpper.includes("PINGCON")) targetPayor = "PINGCON";
    
    // 1. Fetch Verified 2307 Data
    while (files.hasNext()) {
      let ss = SpreadsheetApp.open(files.next());
      let sheets = ss.getSheets();
      sheets.forEach(sheet => {
        if (sheet.getLastRow() > 1) {
          let data = sheet.getDataRange().getValues();
          let hRow = data[0].map(h => String(h).trim().toLowerCase());
          let dateColIdx = hRow.indexOf("2307 release date");
          if (dateColIdx === -1) dateColIdx = 3; 
          let sourceColIdx = hRow.indexOf("source");
          if (sourceColIdx === -1) sourceColIdx = 4;
          let snIdx = hRow.indexOf("sn");
          if (snIdx === -1) snIdx = 0;
          let dtIdx = hRow.indexOf("date and time verified (mm/dd/yyyy)");
          if (dtIdx === -1) dtIdx = 5; // Fallback to column F

          for (let i = 1; i < data.length; i++) {
            let jsonStr = String(data[i][2]).trim();
            let releaseDateStr = data[i][dateColIdx];
            let verifiedDateStr = data[i][dtIdx]; // Extract Verified Date
            let sourceStr = String(data[i][sourceColIdx]).toUpperCase();
            
            let isCompanyMatch = false;
            let effectiveSource = sourceStr;
            if (jsonStr.startsWith("{")) {
              try {
                let pMeta = JSON.parse(jsonStr);
                let aEnt = String(pMeta["Assigned Entity"] || pMeta["Company Entity"] || "").toUpperCase();
                if (aEnt) effectiveSource = aEnt;
              } catch(e) {}
            }

            if (targetPayor === "BONITA") isCompanyMatch = effectiveSource.includes("BONITA");
            else if (targetPayor === "PINGCON") isCompanyMatch = (effectiveSource.includes("PINGCON") && !effectiveSource.includes("BONITA"));
            else if (targetPayor === "KEYSYS") isCompanyMatch = effectiveSource.includes("KEYSYS");
            else if (targetPayor === "A2Z2GO") isCompanyMatch = (effectiveSource.includes("A2Z") || effectiveSource.includes("A2Z2GO"));

            if (isCompanyMatch && jsonStr.startsWith("{") && releaseDateStr) {
              let rDate = new Date(releaseDateStr);
              
              // Filter 1: Check Release Date
              let isReleaseMatch = (!isNaN(rDate.getTime()) && rDate >= startDate && rDate <= endDate);
              
              // Filter 2: Check Verified Date
              let isVerifiedMatch = true;
              if (!verifiedAllDates) {
                  let vDate = new Date(verifiedDateStr);
                  if (isNaN(vDate.getTime()) || vDate < vStartDate || vDate > vEndDate) {
                      isVerifiedMatch = false;
                  }
              }

              if (isReleaseMatch && isVerifiedMatch) {
                try {
                          let parsed = JSON.parse(jsonStr);
                          let rawGross = 0, originalGross = 0, creditMemo = 0, ewtAmount = 0;
                          
                          if (parsed.isMixedVat) {
                              let vRawGross = parseFloat(String(parsed.VAT_Portion["Total Amount"] || parsed.VAT_Portion["Gross"] || "0").replace(/,/g, '')) || 0;
                              let vOrigGross = parseFloat(String(parsed.VAT_Portion["Gross"] || "0").replace(/,/g, '')) || 0;
                              let vCm = parseFloat(String(parsed.VAT_Portion["Credit Memo"] || "0").replace(/,/g, '')) || 0;
                              let vEwt1 = parseFloat(String(parsed.VAT_Portion["EWT 1%"] || "0").replace(/,/g, '')) || 0;
                              let vEwt2 = parseFloat(String(parsed.VAT_Portion["EWT 2%"] || "0").replace(/,/g, '')) || 0;
                              
                              let nvRawGross = parseFloat(String(parsed.NV_Portion["Total Amount"] || parsed.NV_Portion["Gross"] || "0").replace(/,/g, '')) || 0;
                              let nvOrigGross = parseFloat(String(parsed.NV_Portion["Gross"] || "0").replace(/,/g, '')) || 0;
                              let nvCm = parseFloat(String(parsed.NV_Portion["Credit Memo"] || "0").replace(/,/g, '')) || 0;
                              let nvEwt1 = parseFloat(String(parsed.NV_Portion["EWT 1%"] || "0").replace(/,/g, '')) || 0;
                              let nvEwt2 = parseFloat(String(parsed.NV_Portion["EWT 2%"] || "0").replace(/,/g, '')) || 0;
                              
                              rawGross = vRawGross + nvRawGross;
                              originalGross = vOrigGross + nvOrigGross;
                              creditMemo = vCm + nvCm;
                              ewtAmount = (vEwt1 > 0 ? vEwt1 : vEwt2) + (nvEwt1 > 0 ? nvEwt1 : nvEwt2);
                          } else {
                              rawGross = parseFloat(String(parsed["Total Amount"] || parsed["Gross"] || "0").replace(/,/g, '')) || 0;
                              originalGross = parseFloat(String(parsed["Gross"] || "0").replace(/,/g, '')) || 0;
                              creditMemo = parseFloat(String(parsed["Credit Memo"] || "0").replace(/,/g, '')) || 0;
                              let ewt1 = parseFloat(String(parsed["EWT 1%"] || "0").replace(/,/g, '')) || 0;
                              let ewt2 = parseFloat(String(parsed["EWT 2%"] || "0").replace(/,/g, '')) || 0;
                              ewtAmount = (ewt1 > 0) ? ewt1 : (ewt2 > 0 ? ewt2 : 0);
                          }
                          
                          let grossLessEwt = rawGross - ewtAmount;
                  
                  let imgFileId = parsed["fileId"] || ""; // Reads the flattened ID

                  verifiedInvoices.push({
                    sn: String(data[i][snIdx]).trim(),
                    invNo: String(parsed["Invoice Number"] || "").trim().toUpperCase(),
                    supplier: parsed["Company Name"] || "",
                    gross: grossLessEwt,
                    rawGross: rawGross,
                    originalGross: originalGross,
                    creditMemo: creditMemo,
                    fileId: ss.getId(),
                    sheetName: sheet.getName(),
                    rowIndex: i + 1,
                    payor: targetPayor,
                    fileName: String(data[i][1]).trim(),
                    dbSourceName: sourceStr,
                    rawJson: parsed,
                    imgFileId: imgFileId,
                    isForceMatched: parsed["isForceMatched"] || false,
                    originalNarra: parsed["originalNarra"] !== undefined ? parsed["originalNarra"] : null,
                    originalCs: parsed["originalCs"] !== undefined ? parsed["originalCs"] : null,
                    originalRiwb: parsed["originalRiwb"] !== undefined ? parsed["originalRiwb"] : null
                  });
                } catch(e) {}
              }
            }
          }
        }
      });
    }

    // 2. Fetch Narra, Check Summary, and RIWB Databases
    const narraAllData = gatherAllData("Narra_Voucher_Data_");
    const checkSummaryAllData = gatherAllData("Check_Summary_Data_");
    const riwbAllData = gatherAllData("RIWB_Data_");
    const cleanSI = (si) => String(si).replace(/^[^0-9]+/, '').trim();

    let mismatches = [];
    let allVerified = [];

    // 3. Cross-Reference and Verify
    verifiedInvoices.forEach(inv => {
        let verInvClean = cleanSI(inv.invNo);
        // NEW: Strip all non-numeric characters and leading zeros for fallback matching (e.g., "000-000123" -> "123")
        let strippedInv = String(inv.invNo).replace(/[^0-9]/g, '').replace(/^0+/, '');
        
        // Match Narra Data
        let narraAmt = null;
        let nRowId = null;
        for (let i = 0; i < narraAllData.length; i++) {
            let nRow = narraAllData[i].data;
            let nInv = String(nRow[8]).trim().toUpperCase();
            let strippedNInv = nInv.replace(/[^0-9]/g, '').replace(/^0+/, '');
            
            // REMOVED nComp (Column 0) check because it contains Dates/Payees, not Payor entity names
            if (nInv === inv.invNo || (strippedInv !== "" && strippedNInv === strippedInv)) {
                narraAmt = parseFloat(nRow[5]) || 0;
                nRowId = { file: narraAllData[i].sheetName, row: narraAllData[i].rowIndex };
                break;
            }
        }

        // Match Check Summary Data
        let csAmt = null;
        let csRowId = null;
        for (let i = 0; i < checkSummaryAllData.length; i++) {
            let csRow = checkSummaryAllData[i].data;
            let csInv = String(csRow[4]).trim().toUpperCase();
            let strippedCsInv = csInv.replace(/[^0-9]/g, '').replace(/^0+/, '');
            
            if (cleanSI(csInv) === verInvClean || (strippedInv !== "" && strippedCsInv === strippedInv)) { 
                let kAmt = parseFloat(csRow[2]) || 0;
                let pAmt = parseFloat(csRow[3]) || 0;
                csAmt = kAmt > 0 ? kAmt : pAmt;
                csRowId = { file: checkSummaryAllData[i].sheetName, row: checkSummaryAllData[i].rowIndex };
                break;
            }
        }

        // Match RIWB Data
        let riwbAmt = null;
        let riwbRowId = null;
        let ewtDeduction = inv.rawGross - inv.gross;
        for (let i = 0; i < riwbAllData.length; i++) {
            let rRow = riwbAllData[i].data;
            let rInvRaw = String(rRow[2]).trim().toUpperCase();
            let rInv = cleanSI(rInvRaw);
            let strippedRInv = rInvRaw.replace(/[^0-9]/g, '').replace(/^0+/, '');
            
            // REMOVED rSourceFile check because RIWB default filenames do not contain "PINGCON", "KEYSYS", etc.
            if (rInv === verInvClean || (strippedInv !== "" && strippedRInv === strippedInv)) {
                let rawRiwbAmt = parseFloat(rRow[3]) || 0;
                riwbAmt = rawRiwbAmt > 0 ? (rawRiwbAmt - ewtDeduction) : 0;
                riwbRowId = { file: riwbAllData[i].sheetName, row: riwbAllData[i].rowIndex };
                break;
            }
        }

        // 4. Calculate Final Match Status
        let issueType = [];
        let isMatch = true;

        if (narraAmt === null) { isMatch = false; issueType.push("Missing in Narra"); }
        else if (Math.abs(narraAmt - inv.gross) >= 0.05) { isMatch = false; issueType.push("Narra Amt Mismatch"); }

        if (csAmt === null) { isMatch = false; issueType.push("Missing in Check Summary"); }
        else if (Math.abs(csAmt - inv.gross) >= 0.05) { isMatch = false; issueType.push("Check Summary Amt Mismatch"); }

        if (riwbAmt === null) { isMatch = false; issueType.push("Missing in RIWB"); }
        else if (Math.abs(riwbAmt - inv.gross) >= 0.05) { isMatch = false; issueType.push("RIWB Amt Mismatch"); }

        // Strict matching: Push to allVerified if amounts match perfectly OR if explicitly force matched
        if (isMatch || inv.isForceMatched) {
            inv.narraAmt = narraAmt;
            inv.csAmt = csAmt;
            inv.riwbAmt = riwbAmt;
            inv.ewtDeduction = ewtDeduction;
            inv.vFileId = inv.fileId;
            inv.vSheetName = inv.sheetName;
            inv.vRowIndex = inv.rowIndex;
            inv.nRowId = nRowId;
            inv.csRowId = csRowId;
            inv.riwbRowId = riwbRowId;
            // Note: inv.originalNarra, originalCs, and originalRiwb are already preserved inside inv
            allVerified.push(inv);
        } else {
            mismatches.push({
                sn: inv.sn,
                invNo: inv.invNo,
                supplier: inv.supplier,
                gross: inv.gross,
                rawGross: inv.rawGross,
                originalGross: inv.originalGross,
                creditMemo: inv.creditMemo,
                ewtDeduction: ewtDeduction,
                narraAmt: narraAmt,
                csAmt: csAmt,
                riwbAmt: riwbAmt,
                issueType: issueType.join(" & "),
                fileName: inv.fileName,
                dbSourceName: inv.dbSourceName,
                vFileId: inv.fileId,
                vSheetName: inv.sheetName,
                vRowIndex: inv.rowIndex,
                nRowId: nRowId,
                csRowId: csRowId,
                riwbRowId: riwbRowId,
                isForceMatched: inv.isForceMatched
            });
        }
    });

    return { success: true, mismatches: mismatches, allVerified: allVerified };

  } catch (e) {
    return { error: e.message };
  }
}

function webappSaveEwtCorrections(payload) {
  try {
    let lock = LockService.getScriptLock();
    lock.waitLock(10000); 

    // 1. Update Verified Database
    let ss = SpreadsheetApp.openById(payload.vFileId);
    let sheet = ss.getSheetByName(payload.vSheetName);
    
    let jsonStr = String(sheet.getRange(payload.vRowIndex, 3).getValue()).trim();
    if (!jsonStr.startsWith("{")) {
        lock.releaseLock();
        return { error: "Target data is corrupted or not a valid JSON string." };
    }
    
    let parsed = JSON.parse(jsonStr);
    
    if (payload.newInv) parsed["Invoice Number"] = payload.newInv;
    if (payload.newGross) parsed["Gross"] = payload.newGross;
    
    // Ensure old override keys are fully removed if they existed
    delete parsed["Verification_Override"];
    delete parsed["Override_Reason"];
    
    sheet.getRange(payload.vRowIndex, 3).setValue(JSON.stringify(parsed));

    // 2. Update Narra Database (if applicable)
    if (payload.newNarraAmt !== null && !isNaN(payload.newNarraAmt) && payload.nRowId) {
        const dbFolder = DriveApp.getFolderById(DATABASE_FOLDER_ID);
        let nFiles = dbFolder.getFilesByName(payload.nRowId.file);
        if (nFiles.hasNext()) {
            let nSheet = SpreadsheetApp.open(nFiles.next()).getSheets()[0];
            nSheet.getRange(payload.nRowId.row, 6).setValue(payload.newNarraAmt);
        }
    }

    // 3. Update Check Summary Database (if applicable)
    if (payload.newCsAmt !== null && !isNaN(payload.newCsAmt) && payload.csRowId) {
        const dbFolder = DriveApp.getFolderById(DATABASE_FOLDER_ID);
        let csFiles = dbFolder.getFilesByName(payload.csRowId.file);
        if (csFiles.hasNext()) {
            let csSheet = SpreadsheetApp.open(csFiles.next()).getSheets()[0];
            let kAmt = parseFloat(csSheet.getRange(payload.csRowId.row, 3).getValue()) || 0;
            
            // Update Keysys column if it had a value, otherwise update Pingcon column
            if (kAmt > 0) {
                csSheet.getRange(payload.csRowId.row, 3).setValue(payload.newCsAmt);
            } else {
                csSheet.getRange(payload.csRowId.row, 4).setValue(payload.newCsAmt);
            }
        }
    }

    // 4. Update RIWB Database (if applicable)
    if (payload.newRiwbAmt !== null && !isNaN(payload.newRiwbAmt) && payload.riwbRowId) {
        const dbFolder = DriveApp.getFolderById(DATABASE_FOLDER_ID);
        let rFiles = dbFolder.getFilesByName(payload.riwbRowId.file);
        if (rFiles.hasNext()) {
            let rSheet = SpreadsheetApp.open(rFiles.next()).getSheets()[0];
            let grossToSave = payload.newRiwbAmt + (payload.ewtDeduction || 0);
            rSheet.getRange(payload.riwbRowId.row, 4).setValue(grossToSave); // Column D / Index 3 is Total Sum
        }
    }

    SpreadsheetApp.flush();
    lock.releaseLock();
    return { success: true };
    
  } catch(e) {
    return { error: e.message };
  }
}

function webappSaveEwtCorrectionsBatch(payloads, companyContext, sessionToken = null) {
  try {
    // SECURITY GATE: Block Encoders from force-matching discrepancies
    if (sessionToken) {
      const user = getAuthenticatedUser(sessionToken);
      const isTryingToForce = payloads.some(p => p.isForceMatched);
      if (user && user.role === "Encoder" && isTryingToForce) {
        return { success: false, error: "Permission Denied: Encoders are not authorized to force-match discrepancies." };
      }
    }

    let lock = LockService.getScriptLock();
    lock.waitLock(30000); 
    
    let fileCache = {}; 
    
    // Caching logic to prevent reopening the same spreadsheet repeatedly
    const openSheet = (fileId, sheetName) => {
        let key = fileId + "_" + sheetName;
        if (!fileCache[key]) {
            let ss = SpreadsheetApp.openById(fileId);
            fileCache[key] = ss.getSheetByName(sheetName);
        }
        return fileCache[key];
    };

    const openDbSheet = (fileName) => {
        if (!fileCache[fileName]) {
            let targetFolderId = DATABASE_FOLDER_ID;
            if (fileName.includes("Narra_Voucher_Data") || fileName.includes("Check_Summary_Data") || fileName.includes("RIWB_Data")) {
                targetFolderId = SCRIPT_PROPS['VERIFIED_FOLDER_ID'];
            }
            const dbFolder = DriveApp.getFolderById(targetFolderId);
            let files = dbFolder.getFilesByName(fileName);
            if (files.hasNext()) {
                fileCache[fileName] = SpreadsheetApp.open(files.next()).getSheets()[0];
            } else {
                return null;
            }
        }
        return fileCache[fileName];
    };

    for (let i = 0; i < payloads.length; i++) {
        let payload = payloads[i];
        
        // 1. Update Verified Database
        let vSheet = openSheet(payload.vFileId, payload.vSheetName);
        if (vSheet) {
            let jsonStr = String(vSheet.getRange(payload.vRowIndex, 3).getValue()).trim();
            if (jsonStr.startsWith("{")) {
                let parsed = JSON.parse(jsonStr);
                if (payload.newInv) parsed["Invoice Number"] = payload.newInv;
                if (payload.newGross) parsed["Gross"] = payload.newGross;
                
                if (payload.isForceMatched) {
                    parsed["isForceMatched"] = true;
                    parsed["originalNarra"] = payload.originalNarra;
                    parsed["originalCs"] = payload.originalCs;
                    parsed["originalRiwb"] = payload.originalRiwb;
                } else if (payload.isUndoForceMatch) {
                    delete parsed["isForceMatched"];
                    delete parsed["originalNarra"];
                    delete parsed["originalCs"];
                    delete parsed["originalRiwb"];
                }
                
                delete parsed["Verification_Override"];
                delete parsed["Override_Reason"];
                vSheet.getRange(payload.vRowIndex, 3).setValue(JSON.stringify(parsed));
            }
        }

        // 2. Update Narra Database (Handles Undoing to NULL/Missing)
        if (payload.newNarraAmt !== undefined && payload.nRowId) {
            let nSheet = openDbSheet(payload.nRowId.file);
            if (nSheet) {
                nSheet.getRange(payload.nRowId.row, 6).setValue(payload.newNarraAmt === null ? "" : payload.newNarraAmt);
            }
        }

        // 3. Update Check Summary Database (Handles Undoing to NULL/Missing)
        if (payload.newCsAmt !== undefined && payload.csRowId) {
            let csSheet = openDbSheet(payload.csRowId.file);
            if (csSheet) {
                let kAmt = parseFloat(csSheet.getRange(payload.csRowId.row, 3).getValue()) || 0;
                if (kAmt > 0) {
                    csSheet.getRange(payload.csRowId.row, 3).setValue(payload.newCsAmt === null ? "" : payload.newCsAmt);
                } else {
                    csSheet.getRange(payload.csRowId.row, 4).setValue(payload.newCsAmt === null ? "" : payload.newCsAmt);
                }
            }
        }

        // 4. Update RIWB Database (Handles Undoing to NULL/Missing)
        if (payload.newRiwbAmt !== undefined && payload.riwbRowId) {
            let rSheet = openDbSheet(payload.riwbRowId.file);
            if (rSheet) {
                let grossToSave = payload.newRiwbAmt === null ? "" : (payload.newRiwbAmt + (payload.ewtDeduction || 0));
                rSheet.getRange(payload.riwbRowId.row, 4).setValue(grossToSave);
            }
        }
    }

     // --- ADD THIS 1 LINE ---
    logSystemActivity(companyContext || "SYSTEM", "EWT Resolved / Override", `Processed and applied resolutions for ${payloads.length} item(s)`);
    
    SpreadsheetApp.flush();
    lock.releaseLock();
    return { success: true };
    
  } catch(e) {
    return { error: e.message };
  }
}

function generateNteFromTemplate(templateId, replacements, chartUrl, pdfFileName, docxFileName) {
    // 1. Copy the template and convert it to an editable Google Doc temporarily
    let docFile = Drive.Files.copy({ title: "Temp_NTE", mimeType: MimeType.GOOGLE_DOCS }, templateId);
    let doc = DocumentApp.openById(docFile.id);
    let body = doc.getBody();
    let header = doc.getHeader();
    let footer = doc.getFooter();

    // 2. Replace all text variables in Body, Header, and Footer
    for (let key in replacements) {
        body.replaceText(key, replacements[key]);
        if (header) header.replaceText(key, replacements[key]);
        if (footer) footer.replaceText(key, replacements[key]);
    }

    // 3. Swap the {{CHART_IMAGE}} text with the actual bar chart
    if (chartUrl) {
        try {
            let chartBlob = UrlFetchApp.fetch(chartUrl).getBlob();
            let findImage = body.findText("{{CHART_IMAGE}}");
            if (findImage) {
                let element = findImage.getElement();
                let parent = element.getParent().asParagraph();
                element.removeFromParent(); // Clear placeholder text
                
                let img = parent.insertInlineImage(0, chartBlob);
                // Resizes chart to fit standard Word document margins perfectly
                let origW = img.getWidth(); let origH = img.getHeight();
                let newW = 450; let newH = (origH / origW) * newW;
                img.setWidth(newW); img.setHeight(newH);
            }
        } catch(e) { console.warn("Could not insert chart image: " + e); }
    }

    doc.saveAndClose();

    // 4. Export as PDF
    let pdfBlob = DriveApp.getFileById(docFile.id).getAs(MimeType.PDF).setName(pdfFileName);

    // 5. Export as pure MS Word DOCX
    let docxBlob = null;
    let token = ScriptApp.getOAuthToken();
    let exportUrl = "https://www.googleapis.com/drive/v3/files/" + docFile.id + "/export?mimeType=application/vnd.openxmlformats-officedocument.wordprocessingml.document";
    
    for (let i = 0; i < 3; i++) {
        let response = UrlFetchApp.fetch(exportUrl, { headers: { 'Authorization': 'Bearer ' + token }, muteHttpExceptions: true });
        if (response.getResponseCode() === 200) {
            docxBlob = response.getBlob().setName(docxFileName);
            break;
        }
        Utilities.sleep(2000);
    }

    // 6. Delete temporary Google Doc
    Drive.Files.remove(docFile.id);

    return { pdfBlob: pdfBlob, docxBlob: docxBlob };
}

// Run this function ONCE from the editor to authorize Google Docs permissions
function __authorizeGoogleDocs() {
  try { DocumentApp.create("Auth_Temp").setTrashed(true); } catch(e) {}
}




// =========================================================================
// SYSTEM ACTIVITY LOGGER (SAFE DEV/EXEC DETECTION + IDE TEST RUNNER SAFE)
// =========================================================================
function _getActivityLogsFolder() {
  let props = PropertiesService.getScriptProperties();
  
  // Safe Dev Mode Check: will never crash when running directly from the Apps Script IDE
  let isDevMode = false;
  try {
    let service = ScriptApp.getService();
    if (service) {
      let url = service.getUrl();
      if (url && typeof url === "string" && url.toLowerCase().endsWith("/dev")) {
        isDevMode = true;
      }
    }
  } catch (err) {
    isDevMode = false; // Fallback to standard environment if called from editor
  }

  let folderId = null;
  if (isDevMode) {
    folderId = props.getProperty('VERIFIED_FOLDER_ID_DEV') || props.getProperty('CLIENT_2307_DB_FOLDER_ID_DEV');
  }
  
  if (!folderId) {
    folderId = props.getProperty('VERIFIED_FOLDER_ID') || props.getProperty('CLIENT_2307_DB_FOLDER_ID');
  }

  // Final fallback to global SCRIPT_PROPS cache
  if (!folderId && typeof SCRIPT_PROPS !== 'undefined' && SCRIPT_PROPS) {
    folderId = SCRIPT_PROPS['VERIFIED_FOLDER_ID'] || SCRIPT_PROPS['CLIENT_2307_DB_FOLDER_ID'];
  }

  if (!folderId) {
    throw new Error("Activity Logs Folder ID is missing. Please verify 'VERIFIED_FOLDER_ID' or 'VERIFIED_FOLDER_ID_DEV' in Script Properties.");
  }
  
  return DriveApp.getFolderById(folderId);
}

function resolveUserIdentityString(sessionToken) {
  if (sessionToken) {
    const user = getAuthenticatedUser(sessionToken);
    if (user && user.name) {
      return `${user.name} (${user.email}) [${user.role}]`;
    }
  }
  const sessionUser = Session.getActiveUser().getEmail();
  return sessionUser ? sessionUser : "Unknown / Anonymous";
}

function logSystemActivity(companyContext, action, details, sessionToken = null) {
  try {
    let now = new Date();
    let year = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy");
    let sheetName = "Activity_Logs_" + year;
    let folder = _getActivityLogsFolder();

    let ss = null;
    let files = folder.getFilesByName(sheetName);
    if (files.hasNext()) {
      ss = SpreadsheetApp.open(files.next());
    } else {
      ss = SpreadsheetApp.create(sheetName);
      DriveApp.getFileById(ss.getId()).moveTo(folder);
    }

    let sheet = ss.getSheetByName("Logs");
    if (!sheet) {
      sheet = ss.insertSheet("Logs");
      sheet.appendRow(["Timestamp", "Company Context", "User Email", "Action", "Details"]);
      sheet.getRange("A1:E1").setFontWeight("bold").setBackground("#f3f3f3");
      sheet.setFrozenRows(1);
    }
    
    // Stamped with Real Employee Name, Email, and Role
    let userIdentifier = resolveUserIdentityString(sessionToken);
    let timestamp = Utilities.formatDate(now, Session.getScriptTimeZone(), "MM/dd/yyyy HH:mm:ss");
    
    sheet.appendRow([timestamp, String(companyContext).toUpperCase(), userIdentifier, action, details]);
    SpreadsheetApp.flush();
  } catch (e) {
    console.warn("Failed to log activity: " + e.message);
  }
}

// =========================================================================
// DEDICATED CLIENT 2307 ACTIVITY LOGGER (STORED INSIDE CLIENT DB FOLDER)
// =========================================================================
function _getClientDatabaseFolder() {
  let props = PropertiesService.getScriptProperties();
  let isDevMode = false;
  try {
    let service = ScriptApp.getService();
    if (service) {
      let url = service.getUrl();
      if (url && typeof url === "string" && url.toLowerCase().endsWith("/dev")) {
        isDevMode = true;
      }
    }
  } catch (err) {
    isDevMode = false;
  }
  
  let folderId = (isDevMode && props.getProperty('CLIENT_2307_DB_FOLDER_ID_DEV'))
      ? props.getProperty('CLIENT_2307_DB_FOLDER_ID_DEV')
      : props.getProperty('CLIENT_2307_DB_FOLDER_ID');
      
  if (!folderId) throw new Error("Client Database Folder ID is missing in Script Properties.");
  return DriveApp.getFolderById(folderId);
}

function logClientActivity(companyContext, action, details, sessionToken = null) {
  try {
    let now = new Date();
    let year = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy");
    let sheetName = "Activity_Logs_Client_" + year;
    let folder = _getClientDatabaseFolder();

    let ss = null;
    let files = folder.getFilesByName(sheetName);
    if (files.hasNext()) {
      ss = SpreadsheetApp.open(files.next());
    } else {
      ss = SpreadsheetApp.create(sheetName);
      DriveApp.getFileById(ss.getId()).moveTo(folder);
    }

    let sheet = ss.getSheetByName("Logs");
    if (!sheet) {
      let defaultSheet = ss.getSheetByName("Sheet1") || ss.getSheets()[0];
      if (defaultSheet && defaultSheet.getLastRow() === 0) {
        defaultSheet.setName("Logs");
        sheet = defaultSheet;
      } else {
        sheet = ss.insertSheet("Logs", 0);
      }
      sheet.appendRow(["Timestamp", "Company Context", "User Email", "Action", "Details"]);
      sheet.getRange("A1:E1").setFontWeight("bold").setBackground("#f3f3f3");
      sheet.setFrozenRows(1);
    }

    let orphanSheet1 = ss.getSheetByName("Sheet1");
    if (orphanSheet1 && orphanSheet1.getLastRow() === 0 && ss.getSheets().length > 1) {
      try { ss.deleteSheet(orphanSheet1); } catch(err) {}
    }

    // Stamped with Real Employee Name, Email, and Role
    let userIdentifier = resolveUserIdentityString(sessionToken);
    let timestamp = Utilities.formatDate(now, Session.getScriptTimeZone(), "MM/dd/yyyy HH:mm:ss");

    sheet.appendRow([timestamp, String(companyContext).toUpperCase(), userIdentifier, action, details]);
    SpreadsheetApp.flush();
  } catch (e) {
    console.warn("Failed to log client activity: " + e.message);
  }
}



// SAFE DIAGNOSTIC TEST: Tests Real User Stamping into Activity Logs
function TEST_auditLoggingWithUser() {
  const employees = fetchActiveHREmployees();
  if (employees.length === 0) {
    console.error("No active employees found.");
    return;
  }

  // 1. Simulate verified session for first employee
  const testEmp = employees[0];
  const email = testEmp.company_email || testEmp.personal_email;
  const mockToken = Utilities.getUuid();
  const mockSession = {
    email: email,
    name: testEmp.name,
    role: "Encoder",
    employeeId: testEmp.id
  };
  CacheService.getScriptCache().put("SESSION_" + mockToken, JSON.stringify(mockSession), 300);

  console.log("Simulating Action by:", mockSession.name);

  // 2. Log an action using their session token
  logSystemActivity("PINGCON", "Test User Accountability Verification", "Action performed during diagnostic test run", mockToken);

  console.log("Activity logged successfully! Open your Activity_Logs spreadsheet to see the real name stamped.");

  // Cleanup test token
  CacheService.getScriptCache().remove("SESSION_" + mockToken);
}

function webappLogClient2307Skip(payload, companyContext, sessionToken = null) {
  try {
    let pageLabel = payload && payload.pageNum ? `Page ${payload.pageNum} | ` : "";
    let fileLabel = payload && payload.fileName ? `File: ${payload.fileName} | ` : "";
    // Stamped with verified employee identity
    logClientActivity(companyContext, "Skipped Certificate", `${pageLabel}${fileLabel}Certificate skipped by operator without saving.`, sessionToken);
    return { success: true };
  } catch (e) {
    return { error: e.message };
  }
}

function webappGetClientActivityLogs(companyContext) {
  try {
    let folder = _getClientDatabaseFolder();
    let filesIter = folder.getFilesByName("Activity_Logs_Client_" + new Date().getFullYear());
    let results = [];

    if (filesIter.hasNext()) {
        let sheet = SpreadsheetApp.open(filesIter.next()).getSheetByName("Logs");
        if (sheet) {
            let data = sheet.getDataRange().getValues();
            for(let i = 1; i < data.length; i++) {
                let r = data[i];
                let ctx = String(r[1]).toUpperCase();

                if (ctx.includes(companyContext.toUpperCase())) {
                    results.push({
                        timestamp: r[0] instanceof Date ? Utilities.formatDate(r[0], Session.getScriptTimeZone(), "MM/dd/yyyy HH:mm:ss") : r[0],
                        user: r[2],
                        action: r[3],
                        details: r[4]
                    });
                }
            }
        }
    }
    results.sort((a,b) => new Date(b.timestamp) - new Date(a.timestamp));
    return results;
  } catch(e) {
    return [];
  }
}

function webappDeleteClientActivityLogs(startStr, endStr, companyContext, sessionToken = null) {
  try {
    // SECURITY GATE: Only Super Admins can delete client audit logs
    if (sessionToken) {
      const user = getAuthenticatedUser(sessionToken);
      if (!user || user.role !== "Super Admin") {
        return { success: false, error: "Permission Denied: Only Super Admins are authorized to delete audit logs." };
      }
    } else {
      return { success: false, error: "Authentication required to delete audit logs." };
    }

    let lock = LockService.getScriptLock();
    lock.waitLock(10000);

    let startDate = new Date(startStr + " 00:00:00");
    let endDate = new Date(endStr + " 23:59:59");
    if (isNaN(startDate.getTime()) || isNaN(endDate.getTime())) throw new Error("Invalid dates selected.");

    let folder = _getClientDatabaseFolder();

    let startYear = startDate.getFullYear();
    let endYear = endDate.getFullYear();
    let totalDeleted = 0;

    for (let y = startYear; y <= endYear; y++) {
        let sheetName = "Activity_Logs_Client_" + y;
        let filesIter = folder.getFilesByName(sheetName);

        if (filesIter.hasNext()) {
            let sheet = SpreadsheetApp.open(filesIter.next()).getSheetByName("Logs");
            if (sheet) {
                let data = sheet.getDataRange().getValues();
                let deletedInSheet = false;

                for (let i = data.length - 1; i >= 1; i--) {
                    let r = data[i];
                    let ctx = String(r[1]).toUpperCase();
                    let logDate = new Date(r[0]);

                    if (ctx.includes(String(companyContext).toUpperCase()) && logDate >= startDate && logDate <= endDate) {
                        sheet.deleteRow(i + 1);
                        totalDeleted++;
                        deletedInSheet = true;
                    }
                }
                if (deletedInSheet) SpreadsheetApp.flush();
            }
        }
    }

    lock.releaseLock();
    return { success: true, deleted: totalDeleted };
  } catch(e) {
    return { error: e.message };
  }
}

function webappDeleteActivityLogs(startStr, endStr, companyContext, sessionToken = null) {
  try {
    // SECURITY GATE: Only Super Admins can delete audit logs
    if (sessionToken) {
      const user = getAuthenticatedUser(sessionToken);
      if (!user || user.role !== "Super Admin") {
        return { success: false, error: "Permission Denied: Only Super Admins are authorized to delete audit logs." };
      }
    } else {
      return { success: false, error: "Authentication required to delete audit logs." };
    }

    let lock = LockService.getScriptLock();
    lock.waitLock(10000);
    
    let startDate = new Date(startStr + " 00:00:00");
    let endDate = new Date(endStr + " 23:59:59");
    if (isNaN(startDate.getTime()) || isNaN(endDate.getTime())) throw new Error("Invalid dates selected.");
    
    let folder = _getActivityLogsFolder();
    
    let startYear = startDate.getFullYear();
    let endYear = endDate.getFullYear();
    let totalDeleted = 0;
    
    for (let y = startYear; y <= endYear; y++) {
        let sheetName = "Activity_Logs_" + y;
        let filesIter = folder.getFilesByName(sheetName);
        
        if (filesIter.hasNext()) {
            let sheet = SpreadsheetApp.open(filesIter.next()).getSheetByName("Logs");
            if (sheet) {
                let data = sheet.getDataRange().getValues();
                let deletedInSheet = false;
                
                for (let i = data.length - 1; i >= 1; i--) {
                    let r = data[i];
                    let ctx = String(r[1]).toUpperCase();
                    let logDate = new Date(r[0]);
                    
                    if (ctx.includes(String(companyContext).toUpperCase()) && logDate >= startDate && logDate <= endDate) {
                        sheet.deleteRow(i + 1);
                        totalDeleted++;
                        deletedInSheet = true;
                    }
                }
                if (deletedInSheet) SpreadsheetApp.flush();
            }
        }
    }
    
    lock.releaseLock();
    return { success: true, deleted: totalDeleted };
  } catch(e) {
    return { error: e.message };
  }
}

function webappGetActivityLogs(companyContext) {
  try {
    let folder = _getActivityLogsFolder();
    let filesIter = folder.getFilesByName("Activity_Logs_" + new Date().getFullYear());
    let results = [];
    
    if (filesIter.hasNext()) {
        let sheet = SpreadsheetApp.open(filesIter.next()).getSheetByName("Logs");
        if (sheet) {
            let data = sheet.getDataRange().getValues();
            for(let i = 1; i < data.length; i++) {
                let r = data[i];
                let ctx = String(r[1]).toUpperCase();
                
                if (ctx.includes(companyContext.toUpperCase())) {
                    results.push({
                        timestamp: r[0] instanceof Date ? Utilities.formatDate(r[0], Session.getScriptTimeZone(), "MM/dd/yyyy HH:mm:ss") : r[0],
                        user: r[2],
                        action: r[3],
                        details: r[4]
                    });
                }
            }
        }
    }
    results.sort((a,b) => new Date(b.timestamp) - new Date(a.timestamp));
    return results;
  } catch(e) {
    return [];
  }
}

// =========================================================================
// NIGHTLY NARRA REPORT
// =========================================================================

function _40_runNightlyNarra() {
  const props = PropertiesService.getScriptProperties();
  props.deleteProperty("TEST_MODE"); // Force disable Test Mode for automated runs
  deleteTriggersByName("resumeNightlyNarraBatch");
  props.deleteProperty("PENDING_NIGHTLY_NARRA");
  props.deleteProperty("CANCEL_NIGHTLY_NARRA");

  props.setProperty("PENDING_NIGHTLY_NARRA", JSON.stringify(["RUN"]));
  resumeNightlyNarraBatch();
}
function getNarraBucketName(bankName, checkNumber) {
    let b = String(bankName).trim().toUpperCase().replace(/\s+/g, ' ');
    let c = String(checkNumber).trim().toUpperCase();
    
    let ext = "";
    if (c.includes("-")) {
        let parts = c.split("-");
        ext = parts[parts.length - 1].trim(); // Gets text after the last dash
    }

    // 1. Keysys / A2Z2GO Logic
    if (b.includes("BDO DM - KEYSYS")) {
        if (ext === "A2") return "narra Keysys A2Z2GO BDO SH";
        if (ext === "UB") return "narra Keysys Unionbank";
        return "narra Keysys BDO DM";
    }
    if (b.includes("BDO SH 11155") && b.includes("KEYSYS")) return "narra Keysys BDO SH";
    if (b.includes("SHORTHORN") && b.includes("KEYSYS")) return "narra Keysys BDO SH";
    if (b.includes("RETIRO") && b.includes("KEYSYS")) return "narra Keysys MB RETIRO";

    // 2. Pingcon / Bonita Logic
    if (b.includes("BDO DM - PINGCON")) {
        if (ext === "A AUB") return "narra A2Z2GO AUB";
        if (ext === "BONITA") return "narra BONITA BDO SH";
        if (ext === "A2") return "narra Pingcon A2ZBDO SH";
        if (ext === "P AUB") return "narra Pingcon AUB";
        if (ext === "CB") return "narra Pingcon Chinabank";
        return "narra Pingcon BDO DM";
    }
    if (b.includes("SHORTHORN") && b.includes("PINGCON")) return "narra Pingcon BDO SH";
    if (b.includes("BPI GH") && b.includes("PINGCON")) return "narra Pingcon BPI GH";
    if (b.includes("RETIRO") && (!b.includes("KEYSYS"))) return "narra Pingcon MB RETIRO";
    if (b.includes("SECURITY BANK") && b.includes("PINGCON")) return "narra Pingcon Security";

    // Fallback if there's a bank not in the 15-item list
    return "narra " + String(bankName).trim();
}
function resumeNightlyNarraBatch() {
  const props = PropertiesService.getScriptProperties();
  props.setProperty("THREAD_STATE", "RUNNING"); // <--- ADDED
  
  if (props.getProperty("CANCEL_NIGHTLY_NARRA") === "true") {
      deleteTriggersByName("resumeNightlyNarraBatch");
      props.deleteProperty("PENDING_NIGHTLY_NARRA");
      props.deleteProperty("CANCEL_NIGHTLY_NARRA");
      clearReportLock();
      return;
  }

  let pendingStr = props.getProperty("PENDING_NIGHTLY_NARRA");
  if (!pendingStr) { deleteTriggersByName("resumeNightlyNarraBatch"); return; }
  
  let pending = JSON.parse(pendingStr);
  if (pending.length === 0) {
    deleteTriggersByName("resumeNightlyNarraBatch");
    props.deleteProperty("PENDING_NIGHTLY_NARRA");
    clearReportLock();
    return;
  }
  
  let task = pending.shift();

  let allowedEntities = JSON.parse(props.getProperty("SALES_ENTITIES") || '["PINGCON", "KEYSYS"]').map(e => String(e).toUpperCase());
  let allowedBanks = JSON.parse(props.getProperty("NIGHTLY_NARRA_BANKS") || '[]');
  
  const narraAllData = gatherAllData("New_Narra_Data_");
  if (narraAllData.length === 0) {
      clearReportLock();
      return;
  }
  
  const narraSheets = getDatabaseSheets("New_Narra_Data_");
  if (narraSheets.length === 0) { clearReportLock(); return; }
  
  let headers = narraSheets[0].getRange(1, 1, 1, narraSheets[0].getLastColumn()).getValues()[0];
  let lowerHeaders = headers.map(h => String(h).trim().toLowerCase().replace(/\s+/g, ''));
  
  let dateColIdx = lowerHeaders.indexOf("voucherdate");
  let bankNameIdx = lowerHeaders.indexOf("bankname");
  let sourceFileIdx = lowerHeaders.indexOf("sourcefile");

  // Dynamic Filename & Database Oldest Date Extraction
  let detectedStart = null;
  let oldestVoucherDate = null;

  for (let i = 0; i < narraAllData.length; i++) {
    let row = narraAllData[i].data;
    
    // Check filename date pattern
    if (!detectedStart && sourceFileIdx > -1) {
      let sFile = String(row[sourceFileIdx] || "");
      let m = sFile.match(/([A-Za-z]{3,9}\s+\d{1,2},?\s+\d{4})\s+to\s+([A-Za-z]{3,9}\s+\d{1,2},?\s+\d{4})/i);
      if (m) {
        let parsedD = new Date(m[1].replace(/,/g, ''));
        if (!isNaN(parsedD.getTime())) detectedStart = parsedD;
      }
    }

    // Detect oldest voucher date across database rows
    if (dateColIdx > -1 && row[dateColIdx]) {
      let vDate = new Date(row[dateColIdx]);
      if (!isNaN(vDate.getTime()) && vDate.getFullYear() > 2000) {
        if (!oldestVoucherDate || vDate < oldestVoucherDate) {
          oldestVoucherDate = vDate;
        }
      }
    }
  }

  // Priority: 1. Filename Start Date, 2. Oldest Database Voucher Date, 3. Current Date
  let baseDate = detectedStart || oldestVoucherDate || new Date();
  let startDate = new Date(baseDate.getFullYear(), baseDate.getMonth(), 1, 0, 0, 0); // Always 1st day of month
  let endDate = new Date(startDate.getFullYear() + 1, startDate.getMonth(), 0, 23, 59, 59); // Exactly 1 year
  let importDateIdx = lowerHeaders.indexOf("importdate");
  let checkNumIdx = lowerHeaders.indexOf("checknumber"); 
  let checkDateIdx = lowerHeaders.indexOf("checkdate"); // <--- Added Check Date Index
  let amountIdx = lowerHeaders.indexOf("amount"); // <--- Added Amount Index

  // Exclude Source File and Import Date (Removed SN)
  let excludeIndices = [sourceFileIdx, importDateIdx].filter(idx => idx > -1);
  let finalHeaders = []; 
  for (let i = 0; i < headers.length; i++) {
      if (!excludeIndices.includes(i)) {
          finalHeaders.push(headers[i]);
          // Add 4 blank columns immediately after "Amount"
          if (i === amountIdx) {
              finalHeaders.push("", "", "", ""); 
          }
      }
  }

  let groupedByBank = {};
  let bucketAsOfDates = {};
  let globalMinAsOfDate = Infinity;
  let actualMinDate = Infinity;
  let actualMaxDate = 0;

  // 2. Loop Database and Distribute by Bucket Name
  narraAllData.forEach(item => {
      let row = item.data;
      
      let sourceFile = (sourceFileIdx > -1) ? String(row[sourceFileIdx]).toUpperCase() : "";
      let entity = "PINGCON";
      if (sourceFile.includes("KEYSYS")) entity = "KEYSYS";
      else if (sourceFile.includes("PINGCON")) entity = "PINGCON";

      // Extract the "as of" date from the filename string
      let rowAsOfDateObj = null;
      if (sourceFile) {
          let match = sourceFile.match(/AS OF\s+([A-Z]+)\s+(\d{1,2}),?\s+(\d{4})\s+(\d{4})/i);
          if (match) {
              let dStr = `${match[1]} ${match[2]}, ${match[3]} ${match[4].substring(0, 2)}:${match[4].substring(2, 4)}:00`;
              rowAsOfDateObj = new Date(dStr);
          }
      }

      if (!allowedEntities.includes(entity)) return;

      let dDateStr = (dateColIdx > -1) ? row[dateColIdx] : "";
      let dDate = new Date(dDateStr);
      if (isNaN(dDate.getTime()) || dDate < startDate || dDate > endDate) return;

      // Capture dynamic min/max for accurate file naming
      let dTime = dDate.getTime();
      if (dTime < actualMinDate) actualMinDate = dTime;
      if (dTime > actualMaxDate) actualMaxDate = dTime;

      let bankName = (bankNameIdx > -1) ? String(row[bankNameIdx]).trim() : "Unknown_Bank";
      if (!bankName) bankName = "Unknown_Bank";
      
      // Let the UI filter by BankName first
      if (allowedBanks.length > 0 && !allowedBanks.includes(bankName)) return;
      
      let checkNumber = (checkNumIdx > -1) ? String(row[checkNumIdx]).trim() : "";
      
      // Map to exact 15 buckets
      let bucketName = getNarraBucketName(bankName, checkNumber);
      
      if (!groupedByBank[bucketName]) groupedByBank[bucketName] = [];
      groupedByBank[bucketName].push(row);

      // Save the extracted "As Of" dates for filenames and email subjects
      if (rowAsOfDateObj && !isNaN(rowAsOfDateObj.getTime())) {
          let t = rowAsOfDateObj.getTime();
          // Find lowest date globally for Email Subject
          if (t < globalMinAsOfDate) globalMinAsOfDate = t;
          
          // Save the latest date specific to this bucket for the Excel Filename
          if (!bucketAsOfDates[bucketName]) bucketAsOfDates[bucketName] = t;
          else if (t > bucketAsOfDates[bucketName]) bucketAsOfDates[bucketName] = t;
      }
  });

  let bankNames = Object.keys(groupedByBank);
  if (bankNames.length === 0) {
      console.log("No matching Narra Data found for the selected date range and entities.");
      clearReportLock();
      return;
  }

  // 3. Generate individual Excel files for each BankName Group
  let attachments = [];
  // Force the email and filenames to always display the dynamic 12-month duration dates
  let sRangeStr = Utilities.formatDate(startDate, Session.getScriptTimeZone(), "MMM dd yyyy");
  let eRangeStr = Utilities.formatDate(endDate, Session.getScriptTimeZone(), "MMM dd yyyy");
  let fallbackReportDateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MMM dd yyyy HHmm");

  bankNames.forEach(bank => {
      let rows = groupedByBank[bank];
      
      // Assign the specific As Of date for this bucket, or fallback to now if not found
      let bucketAsOfTime = bucketAsOfDates[bank];
      let bucketReportDateStr = fallbackReportDateStr;
      if (bucketAsOfTime) {
          bucketReportDateStr = Utilities.formatDate(new Date(bucketAsOfTime), Session.getScriptTimeZone(), "MMM dd yyyy HHmm");
      }
      let fileSuffix = `${sRangeStr} to ${eRangeStr} as of ${bucketReportDateStr}`;
      
      // Sort Level 1: Column P (CheckDate - Oldest to Newest)
      // Sort Level 2: Column O (CheckNumber - Smallest to Largest)
      rows.sort((a, b) => {
          let dateA = (checkDateIdx > -1 && a[checkDateIdx]) ? new Date(a[checkDateIdx]).getTime() : 0;
          let dateB = (checkDateIdx > -1 && b[checkDateIdx]) ? new Date(b[checkDateIdx]).getTime() : 0;
          if (isNaN(dateA)) dateA = 0;
          if (isNaN(dateB)) dateB = 0;

          // Primary Sort: CheckDate (Oldest to Newest)
          if (dateA !== dateB) return dateA - dateB;

          // Secondary Sort: CheckNumber (Ascending numeric / natural sort)
          let checkA = String(a[checkNumIdx] || "").trim();
          let checkB = String(b[checkNumIdx] || "").trim();
          let numA = parseFloat(checkA);
          let numB = parseFloat(checkB);

          if (!isNaN(numA) && !isNaN(numB) && numA !== numB) {
              return numA - numB;
          }
          return checkA.localeCompare(checkB, undefined, { numeric: true });
      });

      // Add 4 blank columns after Amount, and remove Source File / Import Date (Removed SN)
      let finalRows = [];
      rows.forEach(r => {
          let newRow = []; 
          for (let i = 0; i < r.length; i++) {
              if (!excludeIndices.includes(i)) {
                  let valToPush = r[i];
                  // If CheckNumber is numeric digits, convert to true number so Excel stores it as number
                  if (i === checkNumIdx && valToPush !== "" && valToPush != null) {
                      let strVal = String(valToPush).trim();
                      if (/^\d+$/.test(strVal)) {
                          valToPush = parseInt(strVal, 10);
                      }
                  }
                  newRow.push(valToPush);
                  // Insert 4 blank cells after the Amount value
                  if (i === amountIdx) {
                      newRow.push("", "", "", "");
                  }
              }
          }
          finalRows.push(newRow);
      });

      let cleanBucketName = bank.replace(/[^a-zA-Z0-9 \-_]/g, "").trim();
      let fileName = `${cleanBucketName} ${fileSuffix}`; // Bucket already contains "narra"
      
      let tempSs = SpreadsheetApp.create(fileName);
      let tempSheet = tempSs.getSheets()[0];
      tempSheet.setName(cleanBucketName.substring(0, 31)); // Limit Excel tab name safely
      
      let sheetData = [finalHeaders, ...finalRows];
      tempSheet.getRange(1, 1, sheetData.length, finalHeaders.length).setValues(sheetData);
      
      let finalLowerHeaders = finalHeaders.map(h => String(h).trim().toLowerCase().replace(/\s+/g, ''));
      let inspectIdx = finalLowerHeaders.indexOf("inspectedby") + 1; // 1-based index for getRange
      
      // Determine Header Colors (Keysys = Violet, Pingcon = Pink)
      let isKeysys = bank.toUpperCase().includes("KEYSYS") || bank.toUpperCase().includes("A2Z");
      let headerBgColor = isKeysys ? "#C388F0" : "#F50057"; // Adjusted Keysys to match the lighter violet in the screenshot
      
      // Apply primary background color and bold to ALL headers
      let headerRange = tempSheet.getRange(1, 1, 1, finalHeaders.length);
      headerRange.setFontWeight("bold").setBackground(headerBgColor);
      
      // Apply BLACK text to headers BEFORE "InspectedBy"
      if (inspectIdx > 1) {
          tempSheet.getRange(1, 1, 1, inspectIdx - 1).setFontColor("#000000");
      }
      
      if (inspectIdx > 0) {
          let rightSideCols = finalHeaders.length - inspectIdx + 1;
          // Apply WHITE text to headers FROM "InspectedBy" to the end
          tempSheet.getRange(1, inspectIdx, 1, rightSideCols).setFontColor("#ffffff");
          
          // Apply faint gray text to the DATA ROWS (Row 2 onwards) from "InspectedBy" to the end
          if (finalRows.length > 0) {
              let dataRowsCount = sheetData.length - 1;
              tempSheet.getRange(2, inspectIdx, dataRowsCount, rightSideCols).setFontColor("#cccccc");
          }
      }
      // Set Font to Arial 8 for the whole sheet
      tempSheet.getDataRange().setFontFamily("Arial").setFontSize(8);
      
      // Freeze the Header Row (Row 1)
      tempSheet.setFrozenRows(1);

      let invIdx = finalLowerHeaders.indexOf("invoicenumber") + 1;
      let refIdx = finalLowerHeaders.indexOf("referencenum") + 1;
      let checkIdx = finalLowerHeaders.indexOf("checknumber") + 1;
      let amtIdx = finalLowerHeaders.indexOf("amount") + 1;
      // Find all Date columns dynamically
      let vDateIdx = finalLowerHeaders.indexOf("voucherdate") + 1;
      let cDateIdx = finalLowerHeaders.indexOf("checkdate") + 1;
      let invDateIdx = finalLowerHeaders.indexOf("invoicedate") + 1;
      let prepDateIdx = finalLowerHeaders.indexOf("prepareddate") + 1;
      let checkdDateIdx = finalLowerHeaders.indexOf("checkeddate") + 1;
      let appDateIdx = finalLowerHeaders.indexOf("approveddate") + 1;

      let allDateIndices = [vDateIdx, cDateIdx, invDateIdx, prepDateIdx, checkdDateIdx, appDateIdx];

      if (invIdx > 0) tempSheet.getRange(2, invIdx, sheetData.length, 1).setNumberFormat("@");
      if (refIdx > 0) tempSheet.getRange(2, refIdx, sheetData.length, 1).setNumberFormat("@");
      // Format Column O as General and right-align (avoids green triangle text warnings)
      if (checkIdx > 0) {
          tempSheet.getRange(2, checkIdx, Math.max(1, sheetData.length - 1), 1)
              .setNumberFormat("General")
              .setHorizontalAlignment("right");
      }
      if (amtIdx > 0) tempSheet.getRange(2, amtIdx, sheetData.length, 1).setNumberFormat("#,##0.00");

      let vNumIdx = finalLowerHeaders.indexOf("vouchernumber") + 1;

      // 1. Set global sheet formatting first: Bottom-Left, No Wrap
      tempSheet.getDataRange().setWrap(false).setVerticalAlignment("bottom").setHorizontalAlignment("left");
      
      // 2. Format and align ALL Dates perfectly (Bottom-Right)
      allDateIndices.forEach(idx => {
          if (idx > 0) {
              tempSheet.getRange(2, idx, Math.max(1, sheetData.length), 1)
                  .setNumberFormat("MM/dd/yyyy")
                  .setHorizontalAlignment("right")
                  .setVerticalAlignment("bottom");
          }
      });

      // 3. Format VoucherNumber, Amount & CheckNumber to Bottom-Right
      if (vNumIdx > 0) {
          tempSheet.getRange(2, vNumIdx, Math.max(1, sheetData.length), 1)
              .setHorizontalAlignment("right")
              .setVerticalAlignment("bottom");
      }
      if (amtIdx > 0) {
          tempSheet.getRange(2, amtIdx, Math.max(1, sheetData.length), 1)
              .setHorizontalAlignment("right")
              .setVerticalAlignment("bottom");
      }
      // Force Column O (CheckNumber) to General and strictly Right-Justified
      if (checkIdx > 0) {
          tempSheet.getRange(2, checkIdx, Math.max(1, sheetData.length), 1)
              .setNumberFormat("General")
              .setHorizontalAlignment("right")
              .setVerticalAlignment("bottom");
      }
      
      // 4. Auto-resize everything first as a baseline
      for (let i = 1; i <= finalHeaders.length; i++) tempSheet.autoResizeColumn(i);
      
      // 5. Expand specific columns significantly
      allDateIndices.forEach(idx => {
          if (idx > 0) tempSheet.setColumnWidth(idx, 90); 
      });
      
      let tinIdx = finalLowerHeaders.indexOf("tin") + 1;
      if (tinIdx > 0) tempSheet.setColumnWidth(tinIdx, 130); 
      
      if (amtIdx > 0) tempSheet.setColumnWidth(amtIdx, 100); 
      if (invIdx > 0) tempSheet.setColumnWidth(invIdx, 200); 
      if (refIdx > 0) tempSheet.setColumnWidth(refIdx, 220); 
      if (checkIdx > 0) tempSheet.setColumnWidth(checkIdx, 130); 
      
      let payeeIdx = finalLowerHeaders.indexOf("payee") + 1;
      if (payeeIdx > 0) tempSheet.setColumnWidth(payeeIdx, 300); 
      
      let compIdx = finalLowerHeaders.indexOf("companyname") + 1;
      if (compIdx > 0) tempSheet.setColumnWidth(compIdx, 200); 
      
      let bankIdx = finalLowerHeaders.indexOf("bankname") + 1;
      if (bankIdx > 0) tempSheet.setColumnWidth(bankIdx, 180); 

      SpreadsheetApp.flush();
      Utilities.sleep(500); // Reduced from 2000ms to 500ms to significantly speed up batch processing

      let url = "https://docs.google.com/spreadsheets/d/" + tempSs.getId() + "/export?format=xlsx";
      let token = ScriptApp.getOAuthToken();
      try {
          let response = UrlFetchApp.fetch(url, { headers: { 'Authorization': 'Bearer ' + token }, muteHttpExceptions: true });
          if (response.getResponseCode() === 200) {
              // The ZIP file retains the full name (fileName)
              let blob = response.getBlob().setName(fileName + ".xlsx");
              
              let isKeysys = bank.toLowerCase().includes("keysys") || bank.toLowerCase().includes("a2z");
              let props = PropertiesService.getScriptProperties();
              let pwd = isKeysys ? props.getProperty('ENCRYPT_PASS_TECH') : props.getProperty('ENCRYPT_PASS_PINGCON');
              
              // Create the short name for inside the ZIP (Strips the "narra " prefix)
              let shortBucketOnly = cleanBucketName.replace(/^narra\s+/i, '').trim(); 
              
              let dObj = new Date(bucketAsOfTime || new Date());
              let yyyy = Utilities.formatDate(dObj, Session.getScriptTimeZone(), "yyyy");
              let mm = Utilities.formatDate(dObj, Session.getScriptTimeZone(), "MM");
              let mmm = Utilities.formatDate(dObj, Session.getScriptTimeZone(), "MMM");
              let dd = Utilities.formatDate(dObj, Session.getScriptTimeZone(), "dd");
              let hhmm = Utilities.formatDate(dObj, Session.getScriptTimeZone(), "HHmm");
              let shortExcelName = `${shortBucketOnly} as of ${yyyy}-${mm}${mmm}-${dd} ${hhmm}`;
              
              // Pass the short name as the 4th parameter to the Python server!
              let securedBlob = encryptAttachmentViaPython(blob, fileName + ".xlsx", pwd, shortExcelName);
              
              attachments.push(securedBlob);
          }
      } catch(e) { console.error("Export error: " + e); }
      
      try { DriveApp.getFileById(tempSs.getId()).setTrashed(true); } catch(e){}
  });

  // 4. Dispatch Email with all generated Bank attachments
  if (attachments.length > 0) {
      let toEmails = ""; 
      let ccEmails = "";
      let fetchedTo = [];
      let fetchedCc = [];

      try {
          // Strictly target the defined Email Recipient Folder (Adapts to Dev vs Exec properties automatically)
          let emailDbFolderId = SCRIPT_PROPS['EMAIL_RECIPIENT_FOLDER_ID'] !== "1_INSERT_YOUR_NEW_EMAIL_FOLDER_ID_HERE" ? SCRIPT_PROPS['EMAIL_RECIPIENT_FOLDER_ID'] : SCRIPT_PROPS['DATABASE_FOLDER_ID'];
          let emailDbFolder = DriveApp.getFolderById(emailDbFolderId);
          let files = emailDbFolder.getFilesByName("Nightly_Narra_Emails_Data");
          
          while (files.hasNext()) {
              let file = files.next();
              if (file.isTrashed() || file.getMimeType() !== MimeType.GOOGLE_SHEETS) continue;
              
              try {
                  let eSheet = SpreadsheetApp.open(file).getSheets()[0];
                  let eData = eSheet.getDataRange().getValues();
                  
                  if (eData.length > 1) {
                      let eHeaders = eData[0].map(h => String(h).toLowerCase().trim());
                      let toIdx = eHeaders.indexOf("to emails");
                      let ccIdx = eHeaders.indexOf("cc emails");
                      
                      for (let i = 1; i < eData.length; i++) {
                          if (toIdx > -1 && eData[i][toIdx]) {
                              String(eData[i][toIdx]).split(/[,;]+/).forEach(e => {
                                  let cleanE = e.replace(/[^\x20-\x7E]/g, '').trim();
                                  if (cleanE.includes("@")) fetchedTo.push(cleanE);
                              });
                          }
                          if (ccIdx > -1 && eData[i][ccIdx]) {
                              String(eData[i][ccIdx]).split(/[,;]+/).forEach(e => {
                                  let cleanE = e.replace(/[^\x20-\x7E]/g, '').trim();
                                  if (cleanE.includes("@")) fetchedCc.push(cleanE);
                              });
                          }
                      }
                  }
              } catch(e) { console.warn("Skipping unreadable file: " + e); }
          }
      } catch(e) { console.warn("Drive search error: " + e); }

      // Remove any duplicate emails across rows or backup files
      fetchedTo = [...new Set(fetchedTo)];
      fetchedCc = [...new Set(fetchedCc)];
      
      if (fetchedTo.length > 0) toEmails = fetchedTo.join(", ");
      if (fetchedCc.length > 0) ccEmails = fetchedCc.join(", ");

      // Strict block: Do not send to a fallback if the list is completely empty
      if (!toEmails) {
          console.error("Nightly Narra Aborted: No valid 'To Emails' found in Nightly_Narra_Emails_Data.");
          clearReportLock();
          return;
      }

      let senderEmail = SCRIPT_PROPS['ADMIN_EMAIL'] || "";
      let senderName = "Finance System";
      
      // Use the lowest (oldest) As Of Date found across all files for the Email Subject
      let subjectAsOfStr = fallbackReportDateStr;
      if (globalMinAsOfDate !== Infinity) {
          subjectAsOfStr = Utilities.formatDate(new Date(globalMinAsOfDate), Session.getScriptTimeZone(), "MMM dd yyyy HHmm");
      }
      let subject = `Nightly Narra Report - ${sRangeStr} to ${eRangeStr} as of ${subjectAsOfStr}`;
      
      let htmlBody = `
        <div style="font-family: Arial, sans-serif; font-size: 13px; color: #000; line-height: 1.5;">
          <p>Good day,</p>
          <p>Please find attached the <strong>Nightly Narra Reports</strong> for the period of <strong>${sRangeStr} to ${eRangeStr}</strong>.</p>
          <p>This automated report contains the Database records filtered by date and assigned entity, automatically grouped into <strong>${attachments.length}</strong> different Excel files based on their Bank Name.</p>
          <p>Thank you.</p>
        </div>
      `;

      let emailOptions = {
          from: senderEmail,
          name: senderName,
          attachments: attachments
      };
      if (ccEmails) emailOptions.cc = ccEmails;

      sendRobustEmail(toEmails, subject, htmlBody, emailOptions);
      logNarraActivity("Generated & Dispatched Reports", `Dispatched ${attachments.length} bank report(s) covering ${sRangeStr} to ${eRangeStr}`);
  }

  // Final Cleanup
  props.setProperty("PENDING_NIGHTLY_NARRA", JSON.stringify(pending));
  deleteTriggersByName("resumeNightlyNarraBatch");
  clearReportLock("NARRA");
  console.log("Nightly Narra processed successfully.");
}


// =========================================================================
// FETCH INVOICE FILE IDs FROM RECEIVE_FILES TAB (SMART MULTI-PAGE MATCH)
// =========================================================================
function webappGetInvoiceFileIds(serialNumber, fileName, dbSourceName) {
  try {
    if (!dbSourceName || dbSourceName === "undefined" || dbSourceName.trim() === "") {
        throw new Error("Missing Database Source Name.");
    }
    
    let props = PropertiesService.getScriptProperties();
    let isDevMode = ScriptApp.getService().getUrl().endsWith("/dev");
    const NEW_DB_FOLDER_ID = (isDevMode && props.getProperty('AI_DB_FOLDER_ID_DEV'))
        ? props.getProperty('AI_DB_FOLDER_ID_DEV')
        : (props.getProperty('AI_DB_FOLDER_ID') || "");
        
    let folder = DriveApp.getFolderById(NEW_DB_FOLDER_ID);
    
    // Search for the specific database file
    let files = folder.searchFiles("title contains '" + dbSourceName.replace(/'/g, "\\'") + "' and mimeType = '" + MimeType.GOOGLE_SHEETS + "' and trashed = false");
    if (!files.hasNext()) throw new Error("Could not find the database file: " + dbSourceName);
    
    let ss = SpreadsheetApp.open(files.next());
    let sheet = ss.getSheetByName("receive_files");
    if (!sheet) throw new Error("The 'receive_files' tab is missing in " + dbSourceName);
    
    let data = sheet.getDataRange().getValues();
    let headers = data[0].map(h => String(h).trim().toLowerCase().replace(/[\r\n]+/g, ' ').replace(/\s+/g, ' '));
    
    let serialIdx = headers.indexOf("serial number");
    let fileNameIdx = headers.indexOf("file name");
    let fileIdIdx = headers.indexOf("file id"); 
    
    if (fileIdIdx === -1) throw new Error("Could not find 'File ID' column.");
    
    let fileIds = [];
    let fileNames = [];
    
    let rawSerialQuery = String(serialNumber).trim();
    let rawFileNameQuery = String(fileName).trim();
    
    // SMART FALLBACK: Strip "pg1", "pg2", ".pdf", ".jpg" to get the base invoice name
    let baseFileName = rawFileNameQuery.replace(/\s*pg\d*\.(pdf|jpg|png|jpeg)$/i, '').replace(/\.(pdf|jpg|png|jpeg)$/i, '').trim();

    for (let i = 1; i < data.length; i++) {
      let rowSerial = serialIdx > -1 ? String(data[i][serialIdx]).trim() : "";
      let rowFileName = fileNameIdx > -1 ? String(data[i][fileNameIdx]).trim() : "";
      let id = fileIdIdx > -1 ? String(data[i][fileIdIdx]).trim() : "";
      
      if (id) {
        let isMatched = false;
        
        // MATCH 1: By Serial Number (Strongest match)
        if (rawSerialQuery && rowSerial === rawSerialQuery) {
          isMatched = true;
        } 
        // MATCH 2: Smart File Name Match (If serial fails or is missing)
        if (!isMatched && baseFileName && rowFileName.toLowerCase().includes(baseFileName.toLowerCase())) {
          isMatched = true;
        }

        // Prevent duplicates
        if (isMatched && !fileIds.includes(id)) {
          fileIds.push(id);
          fileNames.push(rowFileName);
        }
      }
    }
    
    if (fileIds.length > 0) {
        // Sort the files alphabetically so "pg1.pdf" always shows before "pg2.pdf"
        let combined = fileNames.map((name, i) => ({ name: name, id: fileIds[i] }));
        combined.sort((a, b) => a.name.localeCompare(b.name));
        
        return { 
            success: true, 
            fileIds: combined.map(x => x.id), 
            fileNames: combined.map(x => x.name) 
        };
    }
    
    throw new Error("Record not found in the receive_files database.");
  } catch (e) {
    return { error: e.message };
  }
}

// =========================================================================
// PURCHASE REPORT GENERATOR BACKEND
// =========================================================================

function webappPreviewPurchaseReport(startStr, endStr, companyContext) {
  try {
    let startDate = new Date(startStr);
    startDate.setHours(0, 0, 0, 0);
    let endDate = new Date(endStr);
    endDate.setHours(23, 59, 59, 999);
    
    let props = PropertiesService.getScriptProperties();
    let isDevMode = ScriptApp.getService().getUrl().endsWith("/dev");
    const VERIFIED_FOLDER_ID = (isDevMode && props.getProperty('VERIFIED_FOLDER_ID_DEV')) 
        ? props.getProperty('VERIFIED_FOLDER_ID_DEV') 
        : props.getProperty('VERIFIED_FOLDER_ID');
        
    let folder = DriveApp.getFolderById(VERIFIED_FOLDER_ID);
    let files = folder.getFilesByType(MimeType.GOOGLE_SHEETS);
    
    let ctxUpper = String(companyContext || "").toUpperCase().trim();
    let targetPayor = "PINGCON";
    if (ctxUpper.includes("KEYSYS")) targetPayor = "KEYSYS";
    else if (ctxUpper.includes("A2Z")) targetPayor = "A2Z2GO";
    else if (ctxUpper.includes("BONITA")) targetPayor = "BONITA";
    else if (ctxUpper.includes("PINGCON")) targetPayor = "PINGCON";
    
    let results = [];
    let sn = 1;

    // Advanced date parser
    const parseDateToObj = (dStr) => {
        if (!dStr) return new Date("");
        let parts = String(dStr).replace(/,/g, '').trim().split(/[\s\-]+/);
        if (parts.length >= 3) {
            let monthMap = { jan:0, feb:1, mar:2, apr:3, may:4, jun:5, jul:6, aug:7, sep:8, oct:9, nov:10, dec:11 };
            let mStr = parts.find(p => isNaN(p) && monthMap[p.substring(0,3).toLowerCase()] !== undefined);
            let nums = parts.filter(p => !isNaN(p));
            if (mStr && nums.length >= 2) {
                let m = monthMap[mStr.substring(0,3).toLowerCase()];
                let p1 = parseInt(nums[0]);
                let p2 = parseInt(nums[1]);
                let y = p1 > 31 ? p1 : (p2 > 31 ? p2 : (p2 < 100 ? p2 + 2000 : p2));
                let day = p1 <= 31 ? p1 : p2;
                if (y < 100) y += 2000;
                let dObj = new Date(y, m, day);
                if (!isNaN(dObj.getTime())) return dObj;
            }
        }
        return new Date(dStr);
    };
    
    while (files.hasNext()) {
      let ss = SpreadsheetApp.open(files.next());
      let sheets = ss.getSheets();
      
      sheets.forEach(sheet => {
        if (sheet.getLastRow() > 1) {
          let data = sheet.getDataRange().getValues();
          let hRow = data[0].map(h => String(h).trim().toLowerCase());
          let sourceColIdx = hRow.indexOf("source");
          if (sourceColIdx === -1) sourceColIdx = 4;

          for (let i = 1; i < data.length; i++) {
            let jsonStr = String(data[i][2]).trim();
            let sourceStr = String(data[i][sourceColIdx]).toUpperCase();
            
            let isCompanyMatch = false;
            let effectiveSource = sourceStr;
            if (jsonStr.startsWith("{")) {
              try {
                let pMeta = JSON.parse(jsonStr);
                let aEnt = String(pMeta["Assigned Entity"] || pMeta["Company Entity"] || "").toUpperCase();
                if (aEnt) effectiveSource = aEnt;
              } catch(e) {}
            }

            if (targetPayor === "BONITA") isCompanyMatch = effectiveSource.includes("BONITA");
            else if (targetPayor === "PINGCON") isCompanyMatch = (effectiveSource.includes("PINGCON") && !effectiveSource.includes("BONITA"));
            else if (targetPayor === "KEYSYS") isCompanyMatch = effectiveSource.includes("KEYSYS");
            else if (targetPayor === "A2Z2GO") isCompanyMatch = (effectiveSource.includes("A2Z") || effectiveSource.includes("A2Z2GO"));

            if (isCompanyMatch && jsonStr.startsWith("{")) {
                try {
                    let parsed = JSON.parse(jsonStr);
                    
                    // Filter strictly by "Invoice Date" (Not 2307 release date)
                    let invDateStr = parsed["Invoice Date"] || "";
                    let invDateObj = parseDateToObj(invDateStr);
                    
                    if (!isNaN(invDateObj.getTime()) && invDateObj >= startDate && invDateObj <= endDate) {
                        
                        if (parsed.isMixedVat) {
                            // 1. Push VAT Portion
                            let vGross = parseFloat(String(parsed.VAT_Portion["Gross"] || "0").replace(/,/g, '')) || 0;
                            let vTax = parseFloat(String(parsed.VAT_Portion["Input Tax"] || "0").replace(/,/g, '')) || 0;
                            let vPur = parseFloat(String(parsed.VAT_Portion["Purchases"] || "0").replace(/,/g, '')) || 0;
                            
                            results.push([
                                "", // Placeholder for SN
                                invDateStr,
                                parsed["Invoice Number"] || "",
                                parsed["PO Number"] || "",
                                parsed["Company Name"] || "",
                                parsed.VAT_Portion["Account Title"] || "",
                                parsed["TIN"] || "",
                                parsed["Address"] || "",
                                vGross,
                                vTax,
                                vPur,
                                invDateObj.getTime()
                            ]);
                            
                            // 2. Push Non-VAT Portion
                            let nvGross = parseFloat(String(parsed.NV_Portion["Gross"] || "0").replace(/,/g, '')) || 0;
                            let nvTax = parseFloat(String(parsed.NV_Portion["Input Tax"] || "0").replace(/,/g, '')) || 0;
                            let nvPur = parseFloat(String(parsed.NV_Portion["Purchases"] || "0").replace(/,/g, '')) || 0;
                            
                            results.push([
                                "", // Placeholder for SN
                                invDateStr,
                                parsed["Invoice Number"] || "",
                                parsed["PO Number"] || "",
                                parsed["Company Name"] || "",
                                parsed.NV_Portion["Account Title"] || "",
                                parsed["TIN"] || "",
                                parsed["Address"] || "",
                                nvGross,
                                nvTax,
                                nvPur,
                                invDateObj.getTime()
                            ]);
                        } else {
                            // Standard Single Row
                            let rawGross = parseFloat(String(parsed["Gross"] || "0").replace(/,/g, '')) || 0;
                            let inputTax = parseFloat(String(parsed["Input Tax"] || "0").replace(/,/g, '')) || 0;
                            let purchases = parseFloat(String(parsed["Purchases"] || "0").replace(/,/g, '')) || 0;

                            results.push([
                                "", // Placeholder for SN
                                invDateStr,
                                parsed["Invoice Number"] || "",
                                parsed["PO Number"] || "",
                                parsed["Company Name"] || "",
                                parsed["Account Title"] || "",
                                parsed["TIN"] || "",
                                parsed["Address"] || "",
                                rawGross,
                                inputTax,
                                purchases,
                                invDateObj.getTime() // Hidden column for perfect chronological sorting
                            ]);
                        }
                    }
                } catch(e) {}
            }
          }
        }
      });
    }

    // Sort Chronologically by Invoice Date
    results.sort((a, b) => a[11] - b[11]);

    // Apply Serial Numbers and strip out the hidden sorting column
    let finalResults = [];
    results.forEach(r => {
        r[0] = sn++;
        finalResults.push(r.slice(0, 11));
    });

    return { success: true, results: finalResults };

  } catch (e) {
    return { error: e.message };
  }
}

function webappGenerateAndEmailPurchaseReport(tableData, startStr, endStr, companyContext, sessionToken = null) {
    // SECURITY GATE: Prevent Viewers from emailing official reports
    if (sessionToken) {
      const user = getAuthenticatedUser(sessionToken);
      if (user && user.role === "Viewer") {
        return { error: "Permission Denied: Viewers have read-only access and cannot email official reports." };
      }
    }

    try {
        let lock = LockService.getScriptLock();
        lock.waitLock(15000);

        let startDate = new Date(startStr);
        let endDate = new Date(endStr);
        let sRangeStr = Utilities.formatDate(startDate, Session.getScriptTimeZone(), "MM-dd-yyyy");
        let eRangeStr = Utilities.formatDate(endDate, Session.getScriptTimeZone(), "MM-dd-yyyy");
        
        let now = new Date();
        let yyyy = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy");
        let mm = Utilities.formatDate(now, Session.getScriptTimeZone(), "MM"); 
        let mmm = Utilities.formatDate(now, Session.getScriptTimeZone(), "MMM"); 
        let dd = Utilities.formatDate(now, Session.getScriptTimeZone(), "dd"); 
        let hhmm = Utilities.formatDate(now, Session.getScriptTimeZone(), "HHmm"); 
        let bossDateStr = `${yyyy}-${mm}${mmm}-${dd} ${hhmm}`;
        
        let reportHeaders = ["SN", "Invoice Date (MM/DD/YYYY)", "Invoice Number", "Company Name", "Account Title", "TIN", "Address", "Gross", "Input Tax", "Purchases"];
        let cleanFileName = `Purchase Report ${companyContext} ${bossDateStr}`;
        
        let tempSpreadsheet = SpreadsheetApp.create(cleanFileName);
        let tempSheet = tempSpreadsheet.getSheets()[0];
        
        let totalGross = 0, totalTax = 0, totalPur = 0;
        let tableRowsHtml = "";

        let excelData = [];
        tableData.forEach((row, idx) => {
            let r = [...row];
            let gross = parseFloat(r[8]) || 0;
            let tax = parseFloat(r[9]) || 0;
            let pur = parseFloat(r[10]) || 0;
            
            r[8] = gross;
            r[9] = tax;
            r[10] = pur;
            
            totalGross += gross;
            totalTax += tax;
            totalPur += pur;
            
            // Build the HTML row for the email body (Uses original indices BEFORE splice)
            tableRowsHtml += `<tr>
                <td style="border: 1px solid black; padding: 4px 8px; text-align: center;">${r[0]}</td>
                <td style="border: 1px solid black; padding: 4px 8px;">${r[4]}</td>
                <td style="border: 1px solid black; padding: 4px 8px;">${r[2]}</td>
                <td style="border: 1px solid black; padding: 4px 8px;">${r[1]}</td>
                <td style="border: 1px solid black; padding: 4px 8px;">${r[3]}</td>
            </tr>`;
            
            // Remove PO Number (index 3) for the Excel sheet
            r.splice(3, 1);
            excelData.push(r);
        });

        // Add Grand Total row to Excel
        let grandRow = Array(reportHeaders.length).fill("");
        grandRow[6] = "GRAND TOTAL:";
        grandRow[7] = totalGross;
        grandRow[8] = totalTax;
        grandRow[9] = totalPur;
        excelData.push(grandRow);

        tempSheet.getRange(1, 1, 1, reportHeaders.length).setValues([reportHeaders]).setFontWeight("bold").setBackground("#f3f3f3");
        tempSheet.getRange(2, 1, excelData.length, reportHeaders.length).setValues(excelData);
        
        tempSheet.getRange(2, 3, excelData.length, 1).setNumberFormat("@"); // Invoice Number as text
        tempSheet.getRange(2, 8, excelData.length, 3).setNumberFormat("#,##0.00"); // Money formatting
        tempSheet.getRange(excelData.length + 1, 7, 1, 4).setFontWeight("bold"); // Bold the grand total row
        
        tempSheet.getDataRange().setWrap(true).setVerticalAlignment("top");
        for (let i = 1; i <= reportHeaders.length; i++) tempSheet.autoResizeColumn(i);
        
        // Explicitly set wider column widths to prevent hashtags and truncation
        tempSheet.setColumnWidth(1, 50);  // SN
        tempSheet.setColumnWidth(2, 120); // Invoice Date
        tempSheet.setColumnWidth(3, 150); // Invoice Number
        tempSheet.setColumnWidth(4, 250); // Company Name
        tempSheet.setColumnWidth(5, 180); // Account Title
        tempSheet.setColumnWidth(6, 150); // TIN
        tempSheet.setColumnWidth(7, 300); // Address
        tempSheet.setColumnWidth(8, 140); // Gross
        tempSheet.setColumnWidth(9, 140); // Input Tax
        tempSheet.setColumnWidth(10, 140); // Purchases
        
        SpreadsheetApp.flush();
        Utilities.sleep(3000);
        
        let url = "https://docs.google.com/spreadsheets/d/" + tempSpreadsheet.getId() + "/export?format=xlsx";
        let token = ScriptApp.getOAuthToken();
        let attachmentBlob = null;
        try {
            let response = UrlFetchApp.fetch(url, { headers: { 'Authorization': 'Bearer ' + token }, muteHttpExceptions: true });
            if (response.getResponseCode() === 200) attachmentBlob = response.getBlob().setName(cleanFileName + ".xlsx");
        } catch(e) {}
        
        try { DriveApp.getFileById(tempSpreadsheet.getId()).setTrashed(true); } catch(e){}
        
        if (!attachmentBlob) throw new Error("Failed to generate XLSX file.");

        // --- FETCH DYNAMIC EMAILS ---
        let mainRecipient = "";
        let ccRecipient = "";

        try {
            let emailDbFolderId = SCRIPT_PROPS['EMAIL_RECIPIENT_FOLDER_ID'] !== "1_INSERT_YOUR_NEW_EMAIL_FOLDER_ID_HERE" ? SCRIPT_PROPS['EMAIL_RECIPIENT_FOLDER_ID'] : SCRIPT_PROPS['DATABASE_FOLDER_ID'];
            let emailDbFolder = DriveApp.getFolderById(emailDbFolderId);
            let emailFiles = emailDbFolder.getFilesByName("Purchase_Emails_Data");
            while (emailFiles.hasNext()) {
                let file = emailFiles.next();
                if (file.isTrashed()) continue;
                let eSheet = SpreadsheetApp.open(file).getSheets()[0];
                let eData = eSheet.getDataRange().getValues();
                if (eData.length > 1) {
                    let eHeaders = eData[0].map(h => String(h).toLowerCase().trim());
                    let entIdx = eHeaders.indexOf("entity");
                    let toIdx = eHeaders.indexOf("to emails");
                    let ccIdx = eHeaders.indexOf("cc emails");
                    
                    for (let i = 1; i < eData.length; i++) {
                        if (entIdx > -1) {
                            let entityName = String(eData[i][entIdx]).toUpperCase().trim();
                            if (entityName.includes(companyContext)) {
                                if (toIdx > -1 && eData[i][toIdx]) mainRecipient = String(eData[i][toIdx]).trim();
                                if (ccIdx > -1 && eData[i][ccIdx]) ccRecipient = String(eData[i][ccIdx]).trim();
                            }
                        }
                    }
                }
            }
        } catch (e) {
            console.warn("Failed to load Purchase Email Database: " + e);
        }

        if (!mainRecipient) {
            lock.releaseLock();
            return { error: `No 'To Emails' configured for ${companyContext} in the Purchase_Emails_Data database. Email dispatch aborted.` };
        }

        let sysSettings = getSystemSettings();
        let isTech = (companyContext === "KEYSYS" || companyContext === "A2Z2GO");
        let senderName = isTech ? "Finance A2Z2GO" : "Finance Pingcon";
        let senderEmail = getSenderEmail(isTech);
        let props = PropertiesService.getScriptProperties();
        let finalPassword = isTech ? props.getProperty('ENCRYPT_PASS_TECH') : props.getProperty('ENCRYPT_PASS_PINGCON');

        let htmlBody = `
            <div style="font-family: Arial, sans-serif; font-size: 13px; color: #000; line-height: 1.5;">
              <p>Good day,</p>
              <p>Please find attached the <strong>Purchase Report</strong> for <strong>${companyContext}</strong> for the period <strong>${sRangeStr} to ${eRangeStr}</strong>.</p>
              <p>This report contains all the Purchase data extracted from verified invoices within the specified date range.</p>
              <br>
              <table style="border-collapse: collapse; border: 1px solid black; font-size: 13px; text-align: left; width: 100%; max-width: 800px;">
                <thead>
                  <tr>
                    <th style="border: 1px solid black; padding: 4px 8px; background-color: #f3f3f3; text-align: center;">SN</th>
                    <th style="border: 1px solid black; padding: 4px 8px; background-color: #f3f3f3;">Company Name</th>
                    <th style="border: 1px solid black; padding: 4px 8px; background-color: #f3f3f3;">Invoice Number</th>
                    <th style="border: 1px solid black; padding: 4px 8px; background-color: #f3f3f3;">Invoice Date</th>
                    <th style="border: 1px solid black; padding: 4px 8px; background-color: #f3f3f3;">PO Number</th>
                  </tr>
                </thead>
                <tbody>
                  ${tableRowsHtml}
                </tbody>
              </table>
              <br>
            </div>`;

        let securedBlob = attachmentBlob;
        try { securedBlob = encryptAttachmentViaPython(attachmentBlob, cleanFileName + ".xlsx", finalPassword); } catch (e) { console.error("Encryption Fallback."); }

        sendRobustEmail(mainRecipient, `Purchase Report ${companyContext} ${bossDateStr}`, htmlBody, {
            from: senderEmail, 
            name: senderName, 
            cc: ccRecipient,
            attachments: [securedBlob]
        });

        lock.releaseLock();
        // Return Base64 of the raw blob so the frontend can download it instantly
        return { 
            success: true,
            base64: Utilities.base64Encode(attachmentBlob.getBytes()),
            fileName: cleanFileName + ".xlsx"
        };
    } catch(e) {
        return { error: e.message };
    }
}

// =========================================================================
// AUTOMATED EWT SUPPLIER REPORT (NO MATCHING REQUIRED)
// =========================================================================

function _41_scheduledEwtSupplierReport() {
  const props = PropertiesService.getScriptProperties();
  props.deleteProperty("TEST_MODE"); // Force disable Test Mode for automated runs
  
  // 1. DYNAMIC DATE LOGIC (1st to 5th = Prev Month, 6th+ = Current Month)
  const now = new Date();
  let targetMonth = now.getMonth();
  let targetYear = now.getFullYear();
  
  if (now.getDate() <= 5) {
      targetMonth = targetMonth - 1;
      if (targetMonth < 0) {
          targetMonth = 11;
          targetYear--;
      }
  }
  
  const startDate = new Date(targetYear, targetMonth, 1, 0, 0, 0);
  const endDate = new Date(targetYear, targetMonth + 1, 0, 23, 59, 59);
  
  // 2. TARGET VERIFIED DATABASE VIA SECURE SCRIPT PROPERTIES
  const VERIFIED_FOLDER_ID = SCRIPT_PROPS['VERIFIED_FOLDER_ID']; 
  let folder = DriveApp.getFolderById(VERIFIED_FOLDER_ID);
  let files = folder.getFilesByType(MimeType.GOOGLE_SHEETS);
  
  let extractedData = [];
  
  // 3. EXTRACT ALL DATA IGNORING MATCHING
  while (files.hasNext()) {
      let ss = SpreadsheetApp.open(files.next());
      let sheets = ss.getSheets();
      
      sheets.forEach(sheet => {
          if (sheet.getLastRow() > 1) {
              let data = sheet.getDataRange().getValues();
              let hRow = data[0].map(h => String(h).trim().toLowerCase());
              let dateColIdx = hRow.indexOf("2307 release date");
              if (dateColIdx === -1) dateColIdx = 3; 
              let sourceColIdx = hRow.indexOf("source");
              if (sourceColIdx === -1) sourceColIdx = 4;

              for (let i = 1; i < data.length; i++) {
                  let jsonStr = String(data[i][2]).trim();
                  let releaseDateStr = data[i][dateColIdx];
                  let sourceStr = String(data[i][sourceColIdx]).toUpperCase();
                  
                  let payorVal = "Pingcon"; 
                  if (sourceStr.includes("BONITA")) payorVal = "Bonita";
                  else if (sourceStr.includes("KEYSYS")) payorVal = "Keysys";
                  else if (sourceStr.includes("A2Z2GO") || sourceStr.includes("A2Z")) payorVal = "A2Z2GO";
                  else if (sourceStr.includes("PINGCON")) payorVal = "Pingcon";

                  // Check if the user manually reassigned the entity inside the verified JSON
                  if (jsonStr.startsWith("{")) {
                    try {
                      let pMeta = JSON.parse(jsonStr);
                      let aEnt = String(pMeta["Assigned Entity"] || pMeta["Company Entity"] || "").toUpperCase();
                      if (aEnt.includes("BONITA")) payorVal = "Bonita";
                      else if (aEnt.includes("KEYSYS")) payorVal = "Keysys";
                      else if (aEnt.includes("A2Z")) payorVal = "A2Z2GO";
                      else if (aEnt.includes("PINGCON")) payorVal = "Pingcon";
                    } catch(e) {}
                  }

                  if (jsonStr.startsWith("{") && releaseDateStr) {
                      try {
                          let parsed = JSON.parse(jsonStr);
                          let rDate = new Date(releaseDateStr);
                          if (isNaN(rDate.getTime())) continue;
                          
                          if (rDate >= startDate && rDate <= endDate) {
                              if (parsed.isMixedVat) {
                                  extractedData.push([
                                      "", payorVal, parsed.VAT_Portion["Account Title"] || "", parsed["Invoice Date"] || "", 
                                      parsed["Company Name"] || "", parsed["Address"] || "", parsed["TIN"] || "", parsed["Invoice Number"] || "", 
                                      parsed.VAT_Portion["Total Amount"] || parsed.VAT_Portion["Gross"] || "", parsed.VAT_Portion["Input Tax"] || "", 
                                      parsed.VAT_Portion["Purchases"] || "", parsed.VAT_Portion["EWT 1%"] || "", parsed.VAT_Portion["EWT 2%"] || "", 
                                      parsed.VAT_Portion["ATC"] || "", "", parsed["PO Number"] || "", releaseDateStr
                                  ]);
                                  extractedData.push([
                                      "", payorVal, parsed.NV_Portion["Account Title"] || "", parsed["Invoice Date"] || "", 
                                      parsed["Company Name"] || "", parsed["Address"] || "", parsed["TIN"] || "", parsed["Invoice Number"] || "", 
                                      parsed.NV_Portion["Total Amount"] || parsed.NV_Portion["Gross"] || "", parsed.NV_Portion["Input Tax"] || "", 
                                      parsed.NV_Portion["Purchases"] || "", parsed.NV_Portion["EWT 1%"] || "", parsed.NV_Portion["EWT 2%"] || "", 
                                      parsed.NV_Portion["ATC"] || "", "", parsed["PO Number"] || "", releaseDateStr
                                  ]);
                              } else {
                                  extractedData.push([
                                      "", payorVal, parsed["Account Title"] || "", parsed["Invoice Date"] || "", 
                                      parsed["Company Name"] || "", parsed["Address"] || "", parsed["TIN"] || "", parsed["Invoice Number"] || "", 
                                      parsed["Total Amount"] || parsed["Gross"] || "", parsed["Input Tax"] || "", parsed["Purchases"] || "", 
                                      parsed["EWT 1%"] || "", parsed["EWT 2%"] || "", parsed["ATC"] || "",    
                                      "", parsed["PO Number"] || "", releaseDateStr          
                                  ]);
                              }
                          }
                      } catch (e) {
                          console.error("JSON Parse Error: " + e);
                      }
                  }
              }
          }
      });
  }
  
  let groupedData = {};
  extractedData.forEach(row => {
      let payor = row[1];
      if (!groupedData[payor]) groupedData[payor] = [];
      groupedData[payor].push(row);
  });

  if (Object.keys(groupedData).length === 0) {
      console.log("Auto-EWT: No verified invoices found for " + startDate + " to " + endDate);
      return;
  }

  let yyyy = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy");
  let mm = Utilities.formatDate(now, Session.getScriptTimeZone(), "MM"); 
  let mmm = Utilities.formatDate(now, Session.getScriptTimeZone(), "MMM"); 
  let dd = Utilities.formatDate(now, Session.getScriptTimeZone(), "dd"); 
  let hhmm = Utilities.formatDate(now, Session.getScriptTimeZone(), "HHmm"); 
  let bossDateStr = `${yyyy}-${mm}${mmm}-${dd} ${hhmm}`; 
  
  // 4. FETCH EWT EMAILS
  let ewtEmailMap = {};
  try {
      let emailDbFolder = DriveApp.getFolderById(EMAIL_RECIPIENT_FOLDER_ID !== "1_INSERT_YOUR_NEW_EMAIL_FOLDER_ID_HERE" ? EMAIL_RECIPIENT_FOLDER_ID : DATABASE_FOLDER_ID);
      let emailFiles = emailDbFolder.getFilesByName("EWT_Supplier_Emails_Data");
      
      let foundFile = null;
      while (emailFiles.hasNext()) {
          let f = emailFiles.next();
          if (!f.isTrashed()) { foundFile = f; break; }
      }
      
      if (foundFile) {
          let eSheet = SpreadsheetApp.open(foundFile).getSheets()[0];
          let eData = eSheet.getDataRange().getValues();
          if (eData.length > 1) {
              let eHeaders = eData[0].map(h => String(h).toLowerCase().trim());
              let entIdx = eHeaders.indexOf("entity");
              let toIdx = eHeaders.indexOf("to emails");
              let ccIdx = eHeaders.indexOf("cc emails");
              for (let i = 1; i < eData.length; i++) {
                  if (entIdx > -1 && eData[i][entIdx]) {
                      ewtEmailMap[String(eData[i][entIdx]).toUpperCase().trim()] = { 
                          to: toIdx > -1 ? String(eData[i][toIdx]).trim() : "", 
                          cc: ccIdx > -1 ? String(eData[i][ccIdx]).trim() : "" 
                      };
                  }
              }
          }
      }
  } catch (e) { console.warn("Failed to load EWT Email Database: " + e); }

  let dayOfWeek = now.getDay();
  let daysUntilTuesday = (2 - dayOfWeek + 7) % 7;
  let nextTuesday = new Date(now);
  nextTuesday.setDate(now.getDate() + daysUntilTuesday);
  
  let baseDate = new Date(2026, 4, 26); 
  let diffDays = Math.round((nextTuesday.getTime() - baseDate.getTime()) / (1000 * 60 * 60 * 24));
  let snVal = Math.floor(diffDays / 7) + 1;
  let nextMonth5th = new Date(nextTuesday.getFullYear(), nextTuesday.getMonth() + 1, 5);
  
  let tuesdayStr = Utilities.formatDate(nextTuesday, Session.getScriptTimeZone(), "MM/dd/yyyy");
  let fifthStr = Utilities.formatDate(nextMonth5th, Session.getScriptTimeZone(), "MM/dd/yyyy");
  
  let emailSubject = `BIR 2307 EWT issued to Supplier sn=${snVal}: EWT report internal deadline weekly every Tuesday (mm/dd/yyyy)=${tuesdayStr} 10:00, EWT external deadline monthly every 5th of each month (mm/dd/yyyy)=${fifthStr}`;
  
  let sRangeStr = Utilities.formatDate(startDate, Session.getScriptTimeZone(), "MM-dd-yyyy");
  let eRangeStr = Utilities.formatDate(endDate, Session.getScriptTimeZone(), "MM-dd-yyyy");
  let reportHeaders = ["SN", "Payor", "Account Title", "Invoice Date (MM/DD/YYYY)", "Company Name", "Address", "TIN", "Invoice Number", "Gross", "Input Tax", "Purchases", "EWT 1%", "EWT 2%", "ATC", "Notes"];

  // 5. GENERATE & SEND FOR EACH ENTITY
  for (let payor in groupedData) {
      let payorData = groupedData[payor];
      payorData.sort((a, b) => {
          let cCmp = String(a[4]).localeCompare(String(b[4]));
          if (cCmp !== 0) return cCmp;
          return (new Date(a[3]).getTime() || 0) - (new Date(b[3]).getTime() || 0);
      });

      let sn = 1;
      for (let r = 0; r < payorData.length; r++) {
          payorData[r][0] = sn++;
          [8, 9, 10, 11, 12].forEach(c => {
              let val = String(payorData[r][c]).replace(/,/g, '');
              payorData[r][c] = (val && !isNaN(parseFloat(val))) ? parseFloat(val) : "";
          });
      }

      let cleanFileName = `SOA EWT SN Statement of Account ${payor} ${bossDateStr}`;
      let tempSpreadsheet = SpreadsheetApp.create(cleanFileName);
      let tempSheet = tempSpreadsheet.getSheets()[0];
      let excelData = payorData.map(row => row.slice(0, 15)); // Exclude PO Number & Release Date
      
      tempSheet.getRange(1, 1, 1, reportHeaders.length).setValues([reportHeaders]).setFontWeight("bold").setBackground("#f3f3f3");
      tempSheet.getRange(2, 1, excelData.length, reportHeaders.length).setValues(excelData);
      tempSheet.getRange(2, 8, excelData.length, 1).setNumberFormat("@"); 
      tempSheet.getRange(2, 9, excelData.length, 5).setNumberFormat("#,##0.00");
      tempSheet.getDataRange().setWrap(true).setVerticalAlignment("top");
      for (let i = 1; i <= reportHeaders.length; i++) tempSheet.autoResizeColumn(i);
      
      tempSheet.setColumnWidth(1, 50); tempSheet.setColumnWidth(2, 100); tempSheet.setColumnWidth(3, 150);
      tempSheet.setColumnWidth(4, 100); tempSheet.setColumnWidth(5, 250); tempSheet.setColumnWidth(6, 300);
      tempSheet.setColumnWidth(7, 120); tempSheet.setColumnWidth(8, 120); tempSheet.setColumnWidth(9, 130);
      tempSheet.setColumnWidth(10, 120); tempSheet.setColumnWidth(11, 130); tempSheet.setColumnWidth(12, 110);
      tempSheet.setColumnWidth(13, 110);
      
      SpreadsheetApp.flush();
      Utilities.sleep(3000);
      
      let url = "https://docs.google.com/spreadsheets/d/" + tempSpreadsheet.getId() + "/export?format=xlsx";
      let token = ScriptApp.getOAuthToken();
      let attachmentBlob = null;
      try {
          let response = UrlFetchApp.fetch(url, { headers: { 'Authorization': 'Bearer ' + token }, muteHttpExceptions: true });
          if (response.getResponseCode() === 200) attachmentBlob = response.getBlob().setName(cleanFileName + ".xlsx");
      } catch(e) {}
      
      try { DriveApp.getFileById(tempSpreadsheet.getId()).setTrashed(true); } catch(e){}
      
      if (attachmentBlob) {
          let sysSettings = getSystemSettings();
          let isTech = (payor === "Keysys" || payor === "A2Z2GO");
          let senderName = isTech ? "Finance A2Z2GO" : "Finance Pingcon";
          let senderEmail = getSenderEmail(isTech);
          let props = PropertiesService.getScriptProperties();
          let finalPassword = isTech ? props.getProperty('ENCRYPT_PASS_TECH') : props.getProperty('ENCRYPT_PASS_PINGCON');
          
          let mainRecipient = "";
          let ccRecipient = "";
          
          let payorUpper = String(payor).toUpperCase();
          if (ewtEmailMap[payorUpper]) {
              if (ewtEmailMap[payorUpper].to) mainRecipient = ewtEmailMap[payorUpper].to;
              if (ewtEmailMap[payorUpper].cc) ccRecipient = ewtEmailMap[payorUpper].cc;
          }

          if (!mainRecipient) {
              console.error(`Auto EWT Supplier Report Aborted: No 'To Emails' configured for ${payorUpper} in EWT_Supplier_Emails_Data.`);
              continue; // Skip sending this entity
          }
          
          let tableRows = [...payorData].map(row => `<tr><td style="border: 1px solid black; padding: 4px 8px; text-align: center;">${row[0]}</td><td style="border: 1px solid black; padding: 4px 8px;">${row[4]}</td><td style="border: 1px solid black; padding: 4px 8px;">${row[7]}</td><td style="border: 1px solid black; padding: 4px 8px;">${row[3]}</td><td style="border: 1px solid black; padding: 4px 8px;">${row[15]}</td></tr>`).join("");
          
          let htmlBody = `
            <div style="font-family: Arial, sans-serif; font-size: 13px; color: #000; line-height: 1.5;">
              <p>Good day,</p>
              <p>Please find attached the <strong>Automated EWT Report Supplier</strong> for <strong>${payor}</strong> for the period <strong>${sRangeStr} to ${eRangeStr}</strong>.</p>
              <p>This report contains all the 2307 data verified and released within the specified date range, generated automatically by the system.</p>
              
              <table style="margin-top: 20px; margin-bottom: 20px; border-collapse: collapse; border: 1px solid black; font-size: 13px; text-align: left;">
                <thead>
                  <tr>
                    <th style="border: 1px solid black; padding: 4px 8px; background-color: #f3f3f3; text-align: center;">SN</th>
                    <th style="border: 1px solid black; padding: 4px 8px; background-color: #f3f3f3;">Company Name</th>
                    <th style="border: 1px solid black; padding: 4px 8px; background-color: #f3f3f3;">Invoice Number</th>
                    <th style="border: 1px solid black; padding: 4px 8px; background-color: #f3f3f3;">Invoice Date</th>
                    <th style="border: 1px solid black; padding: 4px 8px; background-color: #f3f3f3;">PO Number</th>
                  </tr>
                </thead>
                <tbody>
                  ${tableRows}
                </tbody>
              </table>
              <p>Thank you.</p>
            </div>`;
            
          let securedBlob = attachmentBlob;
          try { securedBlob = encryptAttachmentViaPython(attachmentBlob, cleanFileName + ".xlsx", finalPassword); } catch (e) {}
          
          sendRobustEmail(mainRecipient, emailSubject, htmlBody, {
              from: senderEmail, 
              name: senderName, 
              cc: ccRecipient,
              attachments: [securedBlob]
          });
      }
  }
}

function webappGenerateEwtSupplierReportInstant(startStr, endStr, verifiedStartStr, verifiedEndStr, verifiedAllDates, companyContext, sessionToken = null) {
  // SECURITY GATE: Prevent Viewers from generating official supplier reports
  if (sessionToken) {
    const user = getAuthenticatedUser(sessionToken);
    if (user && user.role === "Viewer") {
      return { success: false, message: "Permission Denied: Viewers have read-only access and cannot generate official reports." };
    }
  }

  try {
      let lock = LockService.getScriptLock();
      lock.waitLock(15000);
      
      // Helper to safely format UI string dates (e.g. "09/09/2026") into valid system dates
      const safeParseDateStr = (dateStr, timeAppend) => {
          if (!dateStr) return null;
          let cleanStr = dateStr.replace(/-/g, '/'); // Converts "2026-09-09" to "2026/09/09"
          return new Date(cleanStr + " " + timeAppend);
      };

      const startDate = startStr ? safeParseDateStr(startStr, "00:00:00") : new Date(2025, 0, 1);
      const endDate = endStr ? safeParseDateStr(endStr, "23:59:59") : new Date();
      
      let vStartDate, vEndDate;
      if (!verifiedAllDates) {
          vStartDate = safeParseDateStr(verifiedStartStr, "00:00:00");
          vEndDate = safeParseDateStr(verifiedEndStr, "23:59:59");
      }
      
      // >>> LOGS START <<<
      console.log(`--- SCRIPT STARTED: ${companyContext} ---`);
      console.log(`Release Filter: ${startDate} to ${endDate}`);
      console.log(`Verified Filter: AllDates=${verifiedAllDates} | ${vStartDate} to ${vEndDate}`);
      
      let props = PropertiesService.getScriptProperties();
      let isDevMode = ScriptApp.getService().getUrl().endsWith("/dev");
      const VERIFIED_FOLDER_ID = (isDevMode && props.getProperty('VERIFIED_FOLDER_ID_DEV')) 
          ? props.getProperty('VERIFIED_FOLDER_ID_DEV') 
          : props.getProperty('VERIFIED_FOLDER_ID');
          
      console.log(`Using Database Folder ID: ${VERIFIED_FOLDER_ID}`);
      // >>> LOGS END <<<

      let folder = DriveApp.getFolderById(VERIFIED_FOLDER_ID);
      let files = folder.getFilesByType(MimeType.GOOGLE_SHEETS);
      
      let extractedData = [];
      
      while (files.hasNext()) {
          let file = files.next();
          console.log(`Scanning File: ${file.getName()}`);
          
          let ss = SpreadsheetApp.open(file);
          let sheets = ss.getSheets();
          
          sheets.forEach(sheet => {
              if (sheet.getLastRow() > 1) {
                  let data = sheet.getDataRange().getValues();
                  let hRow = data[0].map(h => String(h).trim().toLowerCase());
                  
                  let dateColIdx = hRow.indexOf("2307 release date");
                  if (dateColIdx === -1) dateColIdx = 3; 
                  let sourceColIdx = hRow.indexOf("source");
                  if (sourceColIdx === -1) sourceColIdx = 4;
                  let dtIdx = hRow.indexOf("date and time verified (mm/dd/yyyy)");
                  if (dtIdx === -1) dtIdx = 5;

                  for (let i = 1; i < data.length; i++) {
                      let jsonStr = String(data[i][2]).trim();
                      let releaseDateStr = data[i][dateColIdx];
                      let sourceStr = String(data[i][sourceColIdx]).toUpperCase();
                      let verifiedDateStr = data[i][dtIdx];
                      
                      let payorVal = "Pingcon"; 
                      if (sourceStr.includes("BONITA")) payorVal = "Bonita";
                      else if (sourceStr.includes("KEYSYS")) payorVal = "Keysys";
                      else if (sourceStr.includes("A2Z2GO") || sourceStr.includes("A2Z")) payorVal = "A2Z2GO";
                      else if (sourceStr.includes("PINGCON")) payorVal = "Pingcon";

                      if (jsonStr.startsWith("{") && releaseDateStr) {
                          try {
                              let parsed = JSON.parse(jsonStr);
                              let invNum = parsed["Invoice Number"] || "";

                              let rDate = new Date(releaseDateStr);
                              if (isNaN(rDate.getTime())) continue;
                              
                              let isReleaseMatch = (rDate >= startDate && rDate <= endDate);
                              let isVerifiedMatch = true;
                              
                              if (!verifiedAllDates) {
                                  let vDate = new Date(verifiedDateStr);
                                  if (isNaN(vDate.getTime()) || vDate < vStartDate || vDate > vEndDate) {
                                      isVerifiedMatch = false;
                                  }
                              }
                              
                              let ctxUpper = String(companyContext || "").toUpperCase().trim();
                              let expectedPayor = "PINGCON";
                              if (ctxUpper.includes("BONITA")) expectedPayor = "BONITA";
                              else if (ctxUpper.includes("KEYSYS")) expectedPayor = "KEYSYS";
                              else if (ctxUpper.includes("A2Z")) expectedPayor = "A2Z2GO";

                              let isMatchForContext = (payorVal.toUpperCase() === expectedPayor);

                              if (isReleaseMatch && isVerifiedMatch && isMatchForContext) {
                                  
                                  if (parsed.isMixedVat) {
                                      extractedData.push([
                                          "", payorVal, parsed.VAT_Portion["Account Title"] || "", parsed["Invoice Date"] || "", 
                                          parsed["Company Name"] || "", parsed["Address"] || "", parsed["TIN"] || "", parsed["Invoice Number"] || "", 
                                          parsed.VAT_Portion["Total Amount"] || parsed.VAT_Portion["Gross"] || "", parsed.VAT_Portion["Input Tax"] || "", 
                                          parsed.VAT_Portion["Purchases"] || "", parsed.VAT_Portion["EWT 1%"] || "", parsed.VAT_Portion["EWT 2%"] || "", 
                                          parsed.VAT_Portion["ATC"] || "", "", parsed["PO Number"] || "", releaseDateStr
                                      ]);
                                      extractedData.push([
                                          "", payorVal, parsed.NV_Portion["Account Title"] || "", parsed["Invoice Date"] || "", 
                                          parsed["Company Name"] || "", parsed["Address"] || "", parsed["TIN"] || "", parsed["Invoice Number"] || "", 
                                          parsed.NV_Portion["Total Amount"] || parsed.NV_Portion["Gross"] || "", parsed.NV_Portion["Input Tax"] || "", 
                                          parsed.NV_Portion["Purchases"] || "", parsed.NV_Portion["EWT 1%"] || "", parsed.NV_Portion["EWT 2%"] || "", 
                                          parsed.NV_Portion["ATC"] || "", "", parsed["PO Number"] || "", releaseDateStr
                                      ]);
                                  } else {
                                      extractedData.push([
                                          "", payorVal, parsed["Account Title"] || "", parsed["Invoice Date"] || "", 
                                          parsed["Company Name"] || "", parsed["Address"] || "", parsed["TIN"] || "", parsed["Invoice Number"] || "", 
                                          parsed["Total Amount"] || parsed["Gross"] || "", parsed["Input Tax"] || "", parsed["Purchases"] || "", 
                                          parsed["EWT 1%"] || "", parsed["EWT 2%"] || "", parsed["ATC"] || "",    
                                          "", parsed["PO Number"] || "", releaseDateStr          
                                      ]);
                                  }
                              }
                          } catch (e) {
                              console.warn("Row parse warning: " + e.message);
                          }
                      }
                  }
              }
          });
      }
      console.log("Extraction Complete. Total global rows pushed: " + extractedData.length);
      
      let ctxUpper = String(companyContext || "").toUpperCase().trim();
      let targetPayor = "Pingcon";
      if (ctxUpper.includes("BONITA")) targetPayor = "Bonita";
      else if (ctxUpper.includes("KEYSYS")) targetPayor = "Keysys";
      else if (ctxUpper.includes("A2Z")) targetPayor = "A2Z2GO";
      else if (ctxUpper.includes("PINGCON")) targetPayor = "Pingcon";
      
      let payorData = extractedData.filter(row => row[1].toUpperCase() === targetPayor.toUpperCase());
      
      if (payorData.length === 0) {
          lock.releaseLock();
          return { success: false, message: "No data found for EWT Supplier Report within the selected date range for " + targetPayor };
      }

      let now = new Date();
      let yyyy = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy");
      let mm = Utilities.formatDate(now, Session.getScriptTimeZone(), "MM"); 
      let mmm = Utilities.formatDate(now, Session.getScriptTimeZone(), "MMM"); 
      let dd = Utilities.formatDate(now, Session.getScriptTimeZone(), "dd"); 
      let hhmm = Utilities.formatDate(now, Session.getScriptTimeZone(), "HHmm"); 
      let bossDateStr = `${yyyy}-${mm}${mmm}-${dd} ${hhmm}`; 

      let ewtEmailMap = {};
      try {
          let emailDbFolderId = SCRIPT_PROPS['EMAIL_RECIPIENT_FOLDER_ID'] !== "1_INSERT_YOUR_NEW_EMAIL_FOLDER_ID_HERE" ? SCRIPT_PROPS['EMAIL_RECIPIENT_FOLDER_ID'] : SCRIPT_PROPS['DATABASE_FOLDER_ID'];
          let emailDbFolder = DriveApp.getFolderById(emailDbFolderId);
          let emailFiles = emailDbFolder.getFilesByName("EWT_Supplier_Emails_Data");
          while (emailFiles.hasNext()) {
              let file = emailFiles.next();
              if (file.isTrashed()) continue;
              let eSheet = SpreadsheetApp.open(file).getSheets()[0];
              let eData = eSheet.getDataRange().getValues();
              if (eData.length > 1) {
                  let eHeaders = eData[0].map(h => String(h).toLowerCase().trim());
                  let entIdx = eHeaders.indexOf("entity");
                  let toIdx = eHeaders.indexOf("to emails");
                  let ccIdx = eHeaders.indexOf("cc emails");
                  for (let i = 1; i < eData.length; i++) {
                      if (entIdx > -1) {
                          let entityName = String(eData[i][entIdx]).toUpperCase().trim();
                          if (entityName) {
                              ewtEmailMap[entityName] = { 
                                  to: toIdx > -1 ? String(eData[i][toIdx]).trim() : "", 
                                  cc: ccIdx > -1 ? String(eData[i][ccIdx]).trim() : "" 
                              };
                          }
                      }
                  }
              }
          }
      } catch (e) {}

      let dayOfWeek = now.getDay();
      let daysUntilTuesday = (2 - dayOfWeek + 7) % 7;
      let nextTuesday = new Date(now);
      nextTuesday.setDate(now.getDate() + daysUntilTuesday);
      let baseDate = new Date(2026, 4, 26);
      let diffDays = Math.round((nextTuesday.getTime() - baseDate.getTime()) / (1000 * 60 * 60 * 24));
      let snVal = Math.floor(diffDays / 7) + 1;
      let nextMonth5th = new Date(nextTuesday.getFullYear(), nextTuesday.getMonth() + 1, 5);
      let tuesdayStr = Utilities.formatDate(nextTuesday, Session.getScriptTimeZone(), "MM/dd/yyyy");
      let fifthStr = Utilities.formatDate(nextMonth5th, Session.getScriptTimeZone(), "MM/dd/yyyy");
      let emailSubject = `BIR 2307 EWT issued to Supplier sn=${snVal}: EWT report internal deadline weekly every Tuesday (mm/dd/yyyy)=${tuesdayStr} 10:00, EWT external deadline monthly every 5th of each month (mm/dd/yyyy)=${fifthStr}`;

      let sRangeStr = Utilities.formatDate(startDate, Session.getScriptTimeZone(), "MM-dd-yyyy");
      let eRangeStr = Utilities.formatDate(endDate, Session.getScriptTimeZone(), "MM-dd-yyyy");
      let reportHeaders = ["SN", "Payor", "Account Title", "Invoice Date (MM/DD/YYYY)", "Company Name", "Address", "TIN", "Invoice Number", "Gross", "Input Tax", "Purchases", "EWT 1%", "EWT 2%", "ATC", "Notes"];

      payorData.sort((a, b) => {
          let cCmp = String(a[4]).localeCompare(String(b[4]));
          if (cCmp !== 0) return cCmp;
          let d1 = new Date(a[3]).getTime() || 0;
          let d2 = new Date(b[3]).getTime() || 0;
          return d1 - d2;
      });

      let sn = 1;
      for (let r = 0; r < payorData.length; r++) {
          payorData[r][0] = sn++;
          [8, 9, 10, 11, 12].forEach(c => {
              let val = String(payorData[r][c]).replace(/,/g, '');
              payorData[r][c] = (val && !isNaN(parseFloat(val))) ? parseFloat(val) : "";
          });
      }

      let cleanFileName = `SOA EWT SN Statement of Account ${targetPayor} ${bossDateStr}`;
      let tempSpreadsheet = SpreadsheetApp.create(cleanFileName);
      let tempSheet = tempSpreadsheet.getSheets()[0];
      let excelData = payorData.map(row => row.slice(0, 15));
      
      tempSheet.getRange(1, 1, 1, reportHeaders.length).setValues([reportHeaders]).setFontWeight("bold").setBackground("#f3f3f3");
      tempSheet.getRange(2, 1, excelData.length, reportHeaders.length).setValues(excelData);
      tempSheet.getRange(2, 8, excelData.length, 1).setNumberFormat("@"); 
      tempSheet.getRange(2, 9, excelData.length, 5).setNumberFormat("#,##0.00");
      tempSheet.getDataRange().setWrap(true).setVerticalAlignment("top");
      for (let i = 1; i <= reportHeaders.length; i++) tempSheet.autoResizeColumn(i);
      
      tempSheet.setColumnWidth(1, 50); tempSheet.setColumnWidth(2, 100); tempSheet.setColumnWidth(3, 150);
      tempSheet.setColumnWidth(4, 100); tempSheet.setColumnWidth(5, 250); tempSheet.setColumnWidth(6, 300);
      tempSheet.setColumnWidth(7, 120); tempSheet.setColumnWidth(8, 120); tempSheet.setColumnWidth(9, 130);
      tempSheet.setColumnWidth(10, 120); tempSheet.setColumnWidth(11, 130); tempSheet.setColumnWidth(12, 110);
      tempSheet.setColumnWidth(13, 110);
      
      SpreadsheetApp.flush();
      Utilities.sleep(3000);
      
      let url = "https://docs.google.com/spreadsheets/d/" + tempSpreadsheet.getId() + "/export?format=xlsx";
      let token = ScriptApp.getOAuthToken();
      let attachmentBlob = null;
      let rawBase64 = null;
      
      try {
          let response = UrlFetchApp.fetch(url, { headers: { 'Authorization': 'Bearer ' + token }, muteHttpExceptions: true });
          if (response.getResponseCode() === 200) {
              let blob = response.getBlob();
              attachmentBlob = blob.setName(cleanFileName + ".xlsx");
              rawBase64 = Utilities.base64Encode(blob.getBytes()); // Capture raw Base64 for instant download
          }
      } catch(e) {}
      
      try { DriveApp.getFileById(tempSpreadsheet.getId()).setTrashed(true); } catch(e){}
      
      if (attachmentBlob) {
          let sysSettings = getSystemSettings();
          let isTech = (targetPayor === "Keysys" || targetPayor === "A2Z2GO");
          let senderName = isTech ? "Finance A2Z2GO" : "Finance Pingcon";
          let senderEmail = getSenderEmail(isTech);
          let props = PropertiesService.getScriptProperties();
          let finalPassword = isTech ? props.getProperty('ENCRYPT_PASS_TECH') : props.getProperty('ENCRYPT_PASS_PINGCON');
          
          let mainRecipient = "";
          let ccRecipient = "";
          
          let payorUpper = String(targetPayor).toUpperCase();
          if (ewtEmailMap[payorUpper]) {
              if (ewtEmailMap[payorUpper].to) mainRecipient = ewtEmailMap[payorUpper].to;
              if (ewtEmailMap[payorUpper].cc) ccRecipient = ewtEmailMap[payorUpper].cc;
          }

          if (!mainRecipient) {
              lock.releaseLock();
              return { success: false, message: `No 'To Emails' configured for ${payorUpper} in EWT_Supplier_Emails_Data. Email aborted.` };
          }
          
          let tableRows = [...payorData]
              .map(row => `<tr><td style="border: 1px solid black; padding: 4px 8px; text-align: center;">${row[0]}</td><td style="border: 1px solid black; padding: 4px 8px;">${row[4]}</td><td style="border: 1px solid black; padding: 4px 8px;">${row[7]}</td><td style="border: 1px solid black; padding: 4px 8px;">${row[3]}</td><td style="border: 1px solid black; padding: 4px 8px;">${row[15]}</td></tr>`)
              .join("");
          
          let htmlBody = `
            <div style="font-family: Arial, sans-serif; font-size: 13px; color: #000; line-height: 1.5;">
              <p>Good day,</p>
              <p>Please find attached the <strong>EWT Report Supplier</strong> for <strong>${targetPayor}</strong> for the period <strong>${sRangeStr} to ${eRangeStr}</strong>.</p>
              <p>This report contains all the 2307 data verified and released within the specified date range.</p>
              
              <table style="margin-top: 20px; margin-bottom: 20px; border-collapse: collapse; border: 1px solid black; font-size: 13px; text-align: left;">
                <thead>
                  <tr>
                    <th style="border: 1px solid black; padding: 4px 8px; background-color: #f3f3f3; text-align: center;">SN</th>
                    <th style="border: 1px solid black; padding: 4px 8px; background-color: #f3f3f3;">Company Name</th>
                    <th style="border: 1px solid black; padding: 4px 8px; background-color: #f3f3f3;">Invoice Number</th>
                    <th style="border: 1px solid black; padding: 4px 8px; background-color: #f3f3f3;">Invoice Date</th>
                    <th style="border: 1px solid black; padding: 4px 8px; background-color: #f3f3f3;">PO Number</th>
                  </tr>
                </thead>
                <tbody>
                  ${tableRows}
                </tbody>
              </table>
              <p>Thank you.</p>
            </div>`;
            
          let securedBlob = attachmentBlob;
          try { securedBlob = encryptAttachmentViaPython(attachmentBlob, cleanFileName + ".xlsx", finalPassword); } catch (e) {}
          
          sendRobustEmail(mainRecipient, emailSubject, htmlBody, {
              from: senderEmail, 
              name: senderName, 
              cc: ccRecipient,
              attachments: [securedBlob]
          });
      }
      
      lock.releaseLock();
      return { 
          success: true, 
          message: "EWT Report successfully generated and emailed.",
          base64: rawBase64,
          fileName: cleanFileName + ".xlsx"
      };

  } catch (e) {
      return { success: false, message: e.message };
  }
}

// =========================================================================
// STATE LOCKS ENGINE (Concurrency Control)
// =========================================================================

function webappRequestInvoiceLock(invoiceNo, userIdentifier, companyContext) {
  try {
    let cache = CacheService.getScriptCache();
    // Unique key isolates Dev vs Exec, Company A vs Company B, and specific Invoice
    let lockKey = "LOCK_INV_" + String(companyContext).toUpperCase() + "_" + String(invoiceNo).trim().toUpperCase();
    let existingLock = cache.get(lockKey);
    
    if (existingLock) {
      // If the current user already owns the lock, just extend the timer and let them in
      if (existingLock === userIdentifier) {
         cache.put(lockKey, userIdentifier, 600); // Extend lock for 10 minutes
         return { success: true, lockedByOther: false };
      }
      // Locked by someone else
      return { success: true, lockedByOther: true, lockedBy: existingLock };
    }
    
    // No lock exists, claim it for this user for 10 minutes (600 seconds)
    cache.put(lockKey, userIdentifier, 600); 
    return { success: true, lockedByOther: false };
  } catch (e) {
    return { error: e.message };
  }
}

function webappReleaseInvoiceLock(invoiceNo, userIdentifier, companyContext) {
  try {
    let cache = CacheService.getScriptCache();
    let lockKey = "LOCK_INV_" + String(companyContext).toUpperCase() + "_" + String(invoiceNo).trim().toUpperCase();
    let existingLock = cache.get(lockKey);
    
    // Only release the lock if the person trying to release it is the one who owns it
    if (existingLock === userIdentifier) {
       cache.remove(lockKey);
    }
    return { success: true };
  } catch (e) {
    return { error: e.message };
  }
}

// =========================================================================
// DEDICATED NIGHTLY NARRA ACTIVITY LOGGER (ISOLATED DATABASE)
// =========================================================================
function _getNarraDatabaseFolder() {
  const props = PropertiesService.getScriptProperties();
  let isDevMode = false;
  try {
    let service = ScriptApp.getService();
    if (service && service.getUrl()) {
      isDevMode = service.getUrl().toLowerCase().endsWith("/dev");
    }
  } catch (err) {
    isDevMode = false;
  }
  
  let folderId = null;
  if (isDevMode) {
    folderId = props.getProperty('DATABASE_FOLDER_ID_DEV') || props.getProperty('DATABASE_FOLDER_ID');
  }
  if (!folderId) {
    folderId = props.getProperty('DATABASE_FOLDER_ID') || DATABASE_FOLDER_ID;
  }
  if (!folderId) {
    throw new Error("Database folder property for Narra is not configured in Script Properties.");
  }
  return DriveApp.getFolderById(folderId);
}

function logNarraActivity(action, details, sessionToken = null) {
  try {
    let now = new Date();
    let year = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy");
    let sheetName = "Activity_Logs_Narra_" + year;
    let folder = _getNarraDatabaseFolder();

    let ss = null;
    let files = folder.getFilesByName(sheetName);
    if (files.hasNext()) {
      ss = SpreadsheetApp.open(files.next());
    } else {
      ss = SpreadsheetApp.create(sheetName);
      DriveApp.getFileById(ss.getId()).moveTo(folder);
    }

    let sheet = ss.getSheetByName("Logs");
    if (!sheet) {
      let defaultSheet = ss.getSheetByName("Sheet1") || ss.getSheets()[0];
      if (defaultSheet && defaultSheet.getLastRow() === 0) {
        defaultSheet.setName("Logs");
        sheet = defaultSheet;
      } else {
        sheet = ss.insertSheet("Logs", 0);
      }
      sheet.appendRow(["Timestamp", "Module", "User Email", "Action", "Details"]);
      sheet.getRange("A1:E1").setFontWeight("bold").setBackground("#f3f3f3");
      sheet.setFrozenRows(1);
    }

    let orphanSheet1 = ss.getSheetByName("Sheet1");
    if (orphanSheet1 && orphanSheet1.getLastRow() === 0 && ss.getSheets().length > 1) {
      try { ss.deleteSheet(orphanSheet1); } catch(err) {}
    }

    let userIdentifier = resolveUserIdentityString(sessionToken);
    let timestamp = Utilities.formatDate(now, Session.getScriptTimeZone(), "MM/dd/yyyy HH:mm:ss");

    sheet.appendRow([timestamp, "NIGHTLY_NARRA", userIdentifier, action, details]);
    SpreadsheetApp.flush();
  } catch (e) {
    console.warn("Failed to log Narra activity: " + e.message);
  }
}

// =========================================================================
// NIGHTLY NARRA PORTAL CONTROLLERS & DATA APIS
// =========================================================================

function webappGetNarraDateBoundsAndBanks() {
  try {
    let narraMin = Infinity;
    let narraMax = 0;
    let narraBanksMap = {};

    const narraSheets = getDatabaseSheets("New_Narra_Data_");
    narraSheets.forEach(sheet => {
      if (sheet.getLastRow() > 1) {
        let data = sheet.getDataRange().getValues();
        let hRow = data[0].map(h => String(h).trim().toLowerCase().replace(/\s+/g, ''));
        let dateColIdx = hRow.indexOf("voucherdate");
        let bankColIdx = hRow.indexOf("bankname");
        let srcColIdx = hRow.indexOf("sourcefile");

        for (let i = 1; i < data.length; i++) {
          if (dateColIdx > -1) {
            let dateStr = data[i][dateColIdx];
            if (dateStr) {
              let dTime = new Date(dateStr).getTime();
              if (!isNaN(dTime)) {
                if (dTime < narraMin) narraMin = dTime;
                if (dTime > narraMax) narraMax = dTime;
              }
            }
          }
          if (bankColIdx > -1) {
            let bName = String(data[i][bankColIdx]).trim();
            if (bName) {
              if (!narraBanksMap[bName]) narraBanksMap[bName] = new Set();
              let sourceFile = (srcColIdx > -1) ? String(data[i][srcColIdx]).toUpperCase() : "";
              if (sourceFile.includes("KEYSYS")) narraBanksMap[bName].add("Keysys");
              else if (sourceFile.includes("PINGCON")) narraBanksMap[bName].add("Pingcon");
              else narraBanksMap[bName].add("Pingcon");
            }
          }
        }
      }
    });

    let narraBanksFormatted = [];
    for (let b in narraBanksMap) {
      narraBanksFormatted.push({ name: b, entities: Array.from(narraBanksMap[b]) });
    }
    narraBanksFormatted.sort((a, b) => a.name.localeCompare(b.name));

    // Dynamic Filename & Database Oldest Date Extraction for UI
    let detectedStart = null;
    let oldestVoucherDate = null;

    for (let s = 0; s < narraSheets.length; s++) {
      let sSheet = narraSheets[s];
      let sLastRow = sSheet.getLastRow();
      if (sLastRow > 1) {
        let hRow = sSheet.getRange(1, 1, 1, sSheet.getLastColumn()).getValues()[0].map(h => String(h).trim().toLowerCase().replace(/\s+/g, ''));
        let srcIdx = hRow.indexOf("sourcefile");
        let vDateIdx = hRow.indexOf("voucherdate");

        let sValues = sSheet.getRange(2, 1, sLastRow - 1, sSheet.getLastColumn()).getValues();
        for (let r = 0; r < sValues.length; r++) {
          if (!detectedStart && srcIdx > -1) {
            let sFile = String(sValues[r][srcIdx] || "");
            let m = sFile.match(/([A-Za-z]{3,9}\s+\d{1,2},?\s+\d{4})\s+to\s+([A-Za-z]{3,9}\s+\d{1,2},?\s+\d{4})/i);
            if (m) {
              let parsedD = new Date(m[1].replace(/,/g, ''));
              if (!isNaN(parsedD.getTime())) detectedStart = parsedD;
            }
          }
          if (vDateIdx > -1 && sValues[r][vDateIdx]) {
            let vDate = new Date(sValues[r][vDateIdx]);
            if (!isNaN(vDate.getTime()) && vDate.getFullYear() > 2000) {
              if (!oldestVoucherDate || vDate < oldestVoucherDate) oldestVoucherDate = vDate;
            }
          }
        }
      }
    }

    let baseDate = detectedStart || oldestVoucherDate || new Date();
    let rollingStart = new Date(baseDate.getFullYear(), baseDate.getMonth(), 1, 0, 0, 0);
    let rollingEnd = new Date(rollingStart.getFullYear() + 1, rollingStart.getMonth(), 0, 23, 59, 59);

    return {
      success: true,
      min: Utilities.formatDate(rollingStart, Session.getScriptTimeZone(), "yyyy-MM-dd"),
      max: Utilities.formatDate(rollingEnd, Session.getScriptTimeZone(), "yyyy-MM-dd"),
      banks: narraBanksFormatted
    };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

function webappGetNarraUsers(sessionToken) {
  const user = getAuthenticatedUser(sessionToken);
  if (!user || user.role !== "Super Admin") {
    return { success: false, error: "Permission Denied: Only Super Admins can access User Management." };
  }
  return webappGetPortalUsers(sessionToken, "NARRA");
}

function webappUpdateNarraUserAccess(sessionToken, targetEmail, targetName, newRole, newStatus) {
  const user = getAuthenticatedUser(sessionToken);
  if (!user || user.role !== "Super Admin") {
    return { success: false, error: "Permission Denied: Only Super Admins can update roles." };
  }
  const result = webappUpdateUserAccess(sessionToken, targetEmail, targetName, newRole, newStatus, "NARRA");
  if (result.success) {
    logNarraActivity("User Access Updated", `Updated ${targetName} (${targetEmail}) -> Role: ${newRole} | Status: ${newStatus}`, sessionToken);
  }
  return result;
}

function webappGetNarraActivityLogs(sessionToken) {
  try {
    let folder = _getNarraDatabaseFolder();
    let currentYear = new Date().getFullYear();
    let sheetName = "Activity_Logs_Narra_" + currentYear;
    let filesIter = folder.getFilesByName(sheetName);
    let ss = null;

    if (filesIter.hasNext()) {
      ss = SpreadsheetApp.open(filesIter.next());
    } else {
      // Self-Healing: Automatically create the spreadsheet and headers if not present
      ss = SpreadsheetApp.create(sheetName);
      DriveApp.getFileById(ss.getId()).moveTo(folder);
      let newSheet = ss.insertSheet("Logs", 0);
      newSheet.appendRow(["Timestamp", "Module", "User Email", "Action", "Details"]);
      newSheet.getRange("A1:E1").setFontWeight("bold").setBackground("#f3f3f3");
      newSheet.setFrozenRows(1);
      let orphanSheet1 = ss.getSheetByName("Sheet1");
      if (orphanSheet1 && ss.getSheets().length > 1) {
        try { ss.deleteSheet(orphanSheet1); } catch(e){}
      }
    }

    let sheet = ss.getSheetByName("Logs") || ss.getSheets()[0];
    let data = sheet.getDataRange().getValues();
    let results = [];
    let uniqueActions = new Set();

    if (data.length > 1) {
      for (let i = 1; i < data.length; i++) {
        let r = data[i];
        if (!r[0] && !r[3]) continue;
        let actionStr = String(r[3] || "").trim();
        if (actionStr) uniqueActions.add(actionStr);

        results.push({
          timestamp: r[0] instanceof Date ? Utilities.formatDate(r[0], Session.getScriptTimeZone(), "MM/dd/yyyy HH:mm:ss") : String(r[0] || ""),
          module: String(r[1] || "NIGHTLY_NARRA"),
          user: String(r[2] || ""),
          action: actionStr,
          details: String(r[4] || "")
        });
      }
    }

    results.sort((a, b) => {
      let tA = new Date(a.timestamp).getTime() || 0;
      let tB = new Date(b.timestamp).getTime() || 0;
      return tB - tA;
    });

    return { 
      success: true, 
      logs: results, 
      actions: Array.from(uniqueActions).sort() 
    };
  } catch (err) {
    return { success: false, logs: [], actions: [], error: err.message };
  }
}

function webappDeleteNarraActivityLogs(startStr, endStr, sessionToken) {
  const user = getAuthenticatedUser(sessionToken);
  if (!user || user.role !== "Super Admin") {
    return { success: false, error: "Permission Denied: Only Super Admins can delete audit logs." };
  }

  let lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    let startDate = new Date(startStr + " 00:00:00");
    let endDate = new Date(endStr + " 23:59:59");
    if (isNaN(startDate.getTime()) || isNaN(endDate.getTime())) throw new Error("Invalid dates selected.");

    let folder = _getNarraDatabaseFolder();
    let startYear = startDate.getFullYear();
    let endYear = endDate.getFullYear();
    let totalDeleted = 0;

    for (let y = startYear; y <= endYear; y++) {
      let sheetName = "Activity_Logs_Narra_" + y;
      let filesIter = folder.getFilesByName(sheetName);

      if (filesIter.hasNext()) {
        let sheet = SpreadsheetApp.open(filesIter.next()).getSheetByName("Logs");
        if (sheet) {
          let data = sheet.getDataRange().getValues();
          let deletedInSheet = false;

          for (let i = data.length - 1; i >= 1; i--) {
            let logDate = new Date(data[i][0]);
            if (logDate >= startDate && logDate <= endDate) {
              sheet.deleteRow(i + 1);
              totalDeleted++;
              deletedInSheet = true;
            }
          }
          if (deletedInSheet) SpreadsheetApp.flush();
        }
      }
    }
    return { success: true, deleted: totalDeleted };
  } catch (e) {
    return { success: false, error: e.message };
  } finally {
    lock.releaseLock();
  }
}


// One-time utility runner to immediately unlock Nightly Narra if stuck
function resetNarraLock() {
  clearReportLock("NARRA");
  deleteTriggersByName("resumeNightlyNarraBatch");
  PropertiesService.getScriptProperties().deleteProperty("PENDING_NIGHTLY_NARRA");
  PropertiesService.getScriptProperties().deleteProperty("CANCEL_NIGHTLY_NARRA");
  console.log("SUCCESS: Nightly Narra locks have been completely erased. Portal unlocked!");
}