//TEMPLATES.GS

// =========================================================================
// HTML TEMPLATES FOR EMAILS, PDFS, AND UI
// =========================================================================

function Template_ConsolidatedNTE(todayStr, REPORT_EMAIL, NTE_PAYLOAD_ACCUMULATOR, WEB_APP_URL) {
  let companyAddr = (PropertiesService.getScriptProperties().getProperty('PINGCON_ADDR') || PropertiesService.getScriptProperties().getProperty('KEYSYS_ADDR') || "").replace(/[\r\n]+/g, '<br>');
  let htmlContent = `
    <div style="font-family: Arial, sans-serif; color: #000; max-width: 800px; margin: 0 auto; padding: 20px;">
      <h2 style="color: #F50057; font-family: 'Dancing Script', cursive, Georgia, serif; margin: 0; font-size: 28px;">Pingcon Marketing Corporation</h2>
      <div style="font-size: 11px; margin-top: 2px;">${companyAddr}</div>
      <h3 style="text-align: center; margin-top: 30px; margin-bottom: 30px; font-weight: bold;">NOTICE TO EXPLAIN</h3>
      <table style="width: 100%; border-collapse: collapse; font-size: 14px; margin-bottom: 10px;">
        <tr><td style="width: 80px; padding-bottom: 5px;">Date:</td><td>${todayStr}</td></tr>
        <tr><td style="padding-bottom: 5px;">Attention:</td><td>${REPORT_EMAIL}</td></tr>
        <tr><td style="padding-bottom: 5px;">From:</td><td>System Administrator</td></tr>
        <tr><td style="padding-bottom: 5px;">Subject:</td><td>CONSOLIDATED NTE REPORT: System Processing Issues</td></tr>
      </table>
      <div style="border-top: 2px solid #000; margin-bottom: 20px;"></div>
      <p style="font-size: 14px;">Good day!</p>
      <p style="font-size: 14px; margin-bottom: 20px;">This is a written notice regarding the following issues identified during the system process:</p>
  `;

  NTE_PAYLOAD_ACCUMULATOR.forEach(entry => {
    htmlContent += `<div style="border: 1px solid #ccc; padding: 15px; margin-bottom: 30px; background-color: #f9f9f9; box-shadow: 2px 2px 5px rgba(0,0,0,0.1);">
      <h4 style="margin-top: 0; color: #cc0000; font-size: 16px;">${entry.subject}</h4>
      <p><strong>Issue:</strong> ${entry.details}</p>`;

    if (entry.subject.toString().toLowerCase().includes("missing sku")) {
      htmlContent += `<table style="width:100%; border-collapse:collapse; font-size:12px; border:1px solid #ddd;">
        <tr style="background-color: #eee; font-weight:bold;">
          <td style="border:1px solid #ddd; padding:4px; text-align:center;">SN</td>
          <td style="border:1px solid #ddd; padding:4px;">Missing SKU</td>
          <td style="border:1px solid #ddd; padding:4px;">Item Name</td>
          <td style="border:1px solid #ddd; padding:4px;">Remarks</td>
        </tr>`;

      let uniqueSkus = {};
      entry.items.forEach(item => {
         let sku = String(item.sku || "").trim();
         if (sku !== "" && !uniqueSkus.hasOwnProperty(sku)) {
            let noteParts =[];
            let desc = String(item.desc || "");
            if (desc.includes("Not found in Basket Mapping")) noteParts.push(`${sku} Not found in Basket Mapping`);
            else noteParts.push("Not found in PO Data");
            let match = desc.match(/(Old SKU.*?\))/i) || desc.match(/(Old SKU.*)/i); 
            if (match) noteParts.push(match[0].replace(/\)$/, ''));
            uniqueSkus[sku] = { remarks: noteParts.join(", "), desc: desc };
         }
      });

      let sn = 1; let skuKeys = Object.keys(uniqueSkus); let previewSkus = skuKeys.slice(0, 10);
      previewSkus.forEach(sku => {
         let info = uniqueSkus[sku];
         htmlContent += `<tr>
           <td style="border:1px solid #ddd; padding:4px; text-align:center;">${sn++}</td>
           <td style="border:1px solid #ddd; padding:4px; font-weight:bold;">${sku}</td>
           <td style="border:1px solid #ddd; padding:4px;">${info.desc}</td>
           <td style="border:1px solid #ddd; padding:4px; color: #E91E63;">${info.remarks}</td>
         </tr>`;
      });
      if (skuKeys.length > 10) htmlContent += `<tr><td colspan="4" style="text-align:center; padding:4px; font-style:italic;">...and ${skuKeys.length - 10} more unique SKUs (see attached Excel file)</td></tr>`;
      htmlContent += `</table>`;
           
      htmlContent += `<h5 style="margin-top: 20px; margin-bottom: 10px; color: #333; font-size: 14px;">Sales Report Details for Missing SKUs</h5>
      <div style="overflow-x:auto;">
      <table style="width:100%; border-collapse:collapse; font-size:12px; border:1px solid #ddd; white-space: nowrap;">
        <tr style="background-color: #eee; font-weight:bold;">
           <td style="border:1px solid #ddd; padding:4px; text-align:center;">SN</td>
           <td style="border:1px solid #ddd; padding:4px;">Document type</td>
           <td style="border:1px solid #ddd; padding:4px;">Customer</td>
           <td style="border:1px solid #ddd; padding:4px;">Status</td>
           <td style="border:1px solid #ddd; padding:4px;">Number</td>
           <td style="border:1px solid #ddd; padding:4px;">Date (MM/DD/YYYY)</td>
           <td style="border:1px solid #ddd; padding:4px;">Additional text</td>
           <td style="border:1px solid #ddd; padding:4px;">Notes</td>
           <td style="border:1px solid #ddd; padding:4px;">Product</td>
           <td style="border:1px solid #ddd; padding:4px;">VAT rate</td>
           <td style="border:1px solid #ddd; padding:4px;">Sold Quantity</td>
           <td style="border:1px solid #ddd; padding:4px;">Unit</td>
           <td style="border:1px solid #ddd; padding:4px;">Net price</td>
           <td style="border:1px solid #ddd; padding:4px;">Net total</td>
           <td style="border:1px solid #ddd; padding:4px;">Row VAT</td>
           <td style="border:1px solid #ddd; padding:4px;">Row total</td>
           <td style="border:1px solid #ddd; padding:4px;">Cust manager</td>
        </tr>`;

      let srSn = 1; let previewSr = entry.items.slice(0, 10);
      previewSr.forEach(item => {
         let r = item.rawRow ||[];
         htmlContent += `<tr>
           <td style="border:1px solid #ddd; padding:4px; text-align:center;">${srSn++}</td>
           <td style="border:1px solid #ddd; padding:4px;">${r[0]||""}</td>
           <td style="border:1px solid #ddd; padding:4px;">${r[1]||""}</td>
           <td style="border:1px solid #ddd; padding:4px;">${r[2]||""}</td>
           <td style="border:1px solid #ddd; padding:4px;">${r[3]||""}</td>
           <td style="border:1px solid #ddd; padding:4px;">${item.date||"-"}</td>
           <td style="border:1px solid #ddd; padding:4px;">${r[5]||""}</td>
           <td style="border:1px solid #ddd; padding:4px;">${r[6]||""}</td>
           <td style="border:1px solid #ddd; padding:4px;">${r[7]||""}</td>
           <td style="border:1px solid #ddd; padding:4px;">${r[8]||""}</td>
           <td style="border:1px solid #ddd; padding:4px;">${r[9]||0}</td>
           <td style="border:1px solid #ddd; padding:4px;">${r[10]||""}</td>
           <td style="border:1px solid #ddd; padding:4px;">${r[11]||0}</td>
           <td style="border:1px solid #ddd; padding:4px;">${r[12]||0}</td>
           <td style="border:1px solid #ddd; padding:4px;">${r[13]||0}</td>
           <td style="border:1px solid #ddd; padding:4px;">${r[14]||0}</td>
           <td style="border:1px solid #ddd; padding:4px;">${r[15]||""}</td>
         </tr>`;
      });
      if (entry.items.length > 10) htmlContent += `<tr><td colspan="17" style="text-align:center; padding:4px; font-style:italic;">...and ${entry.items.length - 10} more rows (see attached XLSX file)</td></tr>`;
      htmlContent += `</table></div>`;

    } else if (entry.subject.toString().toLowerCase().includes("negative profit")) {
      entry.items.sort((a, b) => new Date(b.rawDate) - new Date(a.rawDate));
      htmlContent += `<table style="width:100%; border-collapse:collapse; font-size:12px; border:1px solid #ddd;">
        <tr style="background-color: #eee; font-weight:bold;">
          <td style="border:1px solid #ddd; padding:4px; text-align:center;">SN</td>
          <td style="border:1px solid #ddd; padding:4px;">Purchase Invoice</td>
          <td style="border:1px solid #ddd; padding:4px;">Purchase Date</td>
          <td style="border:1px solid #ddd; padding:4px;">Additional Text</td>
          <td style="border:1px solid #ddd; padding:4px;">Sales Date</td>
          <td style="border:1px solid #ddd; padding:4px;">Supplier</td>
          <td style="border:1px solid #ddd; padding:4px;">Customer</td>
          <td style="border:1px solid #ddd; padding:4px;">SKU</td>
          <td style="border:1px solid #ddd; padding:4px;">Description</td>
          <td style="border:1px solid #ddd; padding:4px;">Cost</td>
          <td style="border:1px solid #ddd; padding:4px;">Sales</td>
          <td style="border:1px solid #ddd; padding:4px;">GP</td>
          <td style="border:1px solid #ddd; padding:4px;">GPM</td>
        </tr>`;
      const previewItems = entry.items.slice(0, 10);
      previewItems.forEach((item, index) => {
         let pDateStr = item.purchaseDate ? Utilities.formatDate(new Date(item.purchaseDate), Session.getScriptTimeZone(), "MM/dd/yyyy") : '-';
         htmlContent += `<tr>
           <td style="border:1px solid #ddd; padding:4px; text-align:center;">${index + 1}</td>
           <td style="border:1px solid #ddd; padding:4px;">${item.purchaseInvoice || '-'}</td>
           <td style="border:1px solid #ddd; padding:4px;">${pDateStr}</td>
           <td style="border:1px solid #ddd; padding:4px;">${item.additionalText || '-'}</td>
           <td style="border:1px solid #ddd; padding:4px;">${item.date || '-'}</td>
           <td style="border:1px solid #ddd; padding:4px;">${item.supplier || '-'}</td>
           <td style="border:1px solid #ddd; padding:4px;">${item.customer || '-'}</td>
           <td style="border:1px solid #ddd; padding:4px;">${item.sku}</td>
           <td style="border:1px solid #ddd; padding:4px;">${item.desc}</td>
           <td style="border:1px solid #ddd; padding:4px;">${parseFloat(item.totalCost).toFixed(2)}</td>
           <td style="border:1px solid #ddd; padding:4px;">${parseFloat(item.totalSales).toFixed(2)}</td>
           <td style="border:1px solid #ddd; padding:4px; color:#cc0000;">${parseFloat(item.gp).toFixed(2)}</td>
           <td style="border:1px solid #ddd; padding:4px;">${(parseFloat(item.gpm) * 100).toFixed(2)}%</td>
         </tr>`;
      });
      if (entry.items.length > 10) htmlContent += `<tr><td colspan="13" style="text-align:center; padding:4px; font-style:italic;">...and ${entry.items.length - 10} more items (see attached PDF)</td></tr>`;
      htmlContent += `</table>`;
    } else {
      htmlContent += `<table style="width:100%; border-collapse:collapse; font-size:12px; border:1px solid #ddd;">
        <tr style="background-color: #eee; font-weight:bold;">
          <td style="border:1px solid #ddd; padding:4px; text-align:center;">SN</td>
          <td style="border:1px solid #ddd; padding:4px;">PO/Ref</td>
          <td style="border:1px solid #ddd; padding:4px;">SKU</td>
          <td style="border:1px solid #ddd; padding:4px;">Description</td>
          <td style="border:1px solid #ddd; padding:4px;">Value</td>
        </tr>`;
      const previewItems = entry.items.slice(0, 10);
      previewItems.forEach((item, index) => {
         htmlContent += `<tr>
           <td style="border:1px solid #ddd; padding:4px; text-align:center;">${index + 1}</td>
           <td style="border:1px solid #ddd; padding:4px;">${item.po || '-'}</td>
           <td style="border:1px solid #ddd; padding:4px;">${item.sku}</td>
           <td style="border:1px solid #ddd; padding:4px;">${item.desc}</td>
           <td style="border:1px solid #ddd; padding:4px;">${parseFloat(item.price).toFixed(2)}</td>
         </tr>`;
      });
      if (entry.items.length > 10) htmlContent += `<tr><td colspan="5" style="text-align:center; padding:4px; font-style:italic;">...and ${entry.items.length - 10} more items (see attached PDF)</td></tr>`;
      htmlContent += `</table>`;
    }

    htmlContent += `<div style="margin-top: 15px; text-align: right;">
        <a href="${WEB_APP_URL}?page=NTE_Reply&row=${entry.rowId}" style="background-color: #F50057; color: white; padding: 10px 20px; text-decoration: none; display: inline-block; font-size: 14px; font-weight: bold; border-radius: 4px;">REPLY TO THIS ISSUE</a>
      </div></div>`;
  });

  htmlContent += `
      <p style="font-size: 14px; text-align: justify; line-height: 1.5; margin-top: 30px;">You are hereby given a chance within five (5) calendar days to explain your side for each issue above. Click the respective <strong>"REPLY TO THIS ISSUE"</strong> button to submit your explanation and update missing prices (if required). If no explanation within the given period, you are clearly waiving your rights to be heard and give your side. With this, you are allowing Pingcon to proceed to the next step of this Notice to Explain.</p>
      <p style="font-size: 14px; margin-bottom: 40px;">For your strict compliance.</p>
      <div style="border-top: 2px solid #000; padding-top: 5px; margin-top: 40px;">
        <p style="font-size: 12px; font-style: italic; margin-top: 0;">Received. Read and understood.</p>
        <table style="width: 100%; margin-top: 40px; text-align: center; font-size: 12px;">
          <tr>
            <td style="width: 30%; border-top: 1px solid #000; padding-top: 5px;">Printed Name</td>
            <td style="width: 5%;"></td>
            <td style="width: 30%; border-top: 1px solid #000; padding-top: 5px;">Signature</td>
            <td style="width: 5%;"></td>
            <td style="width: 30%; border-top: 1px solid #000; padding-top: 5px;">Date</td>
          </tr>
        </table>
      </div>
    </div>`;
  return htmlContent;
}

function Template_SalesReportEmail(firstName, cleanFileName, monthDateObj, monthTotalSales, pctAchieved, monthlyQuota, newClientsCount, chartHtml, quotaMessage, feedbackImgBlob, feedbackLink) {
  let formatDec = (amount) => amount.toFixed(2).replace(/\d(?=(\d{3})+\.)/g, '$&,');
  
  // Dynamically fetch HR details from Script Properties
  let props = PropertiesService.getScriptProperties();
  let hr1Name = props.getProperty('HR_REP_1_NAME') || '';
  let hr1Email = props.getProperty('HR_REP_1_EMAIL') || '';
  let hr2Name = props.getProperty('HR_REP_2_NAME') || '';
  let hr2Email = props.getProperty('HR_EMAIL_2') || '';
  let hrContact = props.getProperty('HR_CONTACT') || '';

  return `
    <div style="font-family: Arial, sans-serif; font-size: 14px; color: #000; line-height: 1.6; max-width: 900px;">
      <p>Dear ${firstName},</p>
      <p>I'm sharing this sales report. Can you please check the attached file?</p>
      <p><strong>${cleanFileName}</strong></p>
      
      <div style="display: flex; gap: 20px; flex-wrap: wrap;">
        <div style="background-color: #fcfcfc; padding: 20px; border: 1px solid #e0e0e0; border-radius: 6px; margin: 25px 0; width: fit-content;">
            <h3 style="margin-top: 0; margin-bottom: 15px; color: #333; font-size: 16px;">Sales Insights - ${Utilities.formatDate(monthDateObj, Session.getScriptTimeZone(), "MMM yyyy")}</h3>
            <table style="font-size: 14px; line-height: 2;">
                <tr><td style="padding-right: 25px; font-weight: bold;">Peso Achieved:</td><td>PHP ${formatDec(monthTotalSales)}</td></tr>
                <tr><td style="padding-right: 25px; font-weight: bold;">% Achieved:</td><td>${pctAchieved}</td></tr>
                <tr><td style="padding-right: 25px; font-weight: bold;">Sales Quota:</td><td>PHP ${monthlyQuota ? formatDec(monthlyQuota) : '0.00'}</td></tr>
                <tr><td style="padding-right: 25px; font-weight: bold;">Number of New Clients:</td><td>${newClientsCount}</td></tr>
            </table>
        </div>
      </div>
      
      ${chartHtml}
      <p>${quotaMessage}</p>
      <p>Thank you for your kindness and support.</p>
      <p>Please do not hesitate to contact:<br>
      ${(!hr1Name || hr1Name.trim().toUpperCase() === 'HR') ? 'HR:' : 'HR: ' + hr1Name} &lt;<a href="mailto:${hr1Email}">${hr1Email}</a>&gt;<br>
      HR: ${hr2Name} &lt;<a href="mailto:${hr2Email}">${hr2Email}</a>&gt;<br>
      HR CP #: ${hrContact}</p>
      <br>
      ${feedbackImgBlob ? `<a href="${feedbackLink}" target="_blank"><img src="cid:feedbackImage" alt="Feedback Matters" style="max-width: 350px; height: auto; border: none;"></a>` : ''}
    </div>`;
}

function Template_SalesQuotaEmailsAndPdf(
  isQuotaHit, nameNoComma, dateRangeStr, greetingName, financeEmailStr, financeSignStr, 
  isA2ZAe, aeName, startDateStr, endDateStr, monthlyQuota, monthTotalSales, pctAchieved, 
  supName, c1Url, feedbackImgBlob, feedbackLink, department, incidentRef, replyLink
) {
    let secSubject = ""; let secBody = "";
    let ntePdfBlob = null; let nteDocxBlob = null;
    let props = PropertiesService.getScriptProperties();
    let hr1Name = props.getProperty('HR_REP_1_NAME') || '';
    let hr1Email = props.getProperty('HR_REP_1_EMAIL') || '';
    let hr2Name = props.getProperty('HR_REP_2_NAME') || '';
    let hr2Email = props.getProperty('HR_EMAIL_2') || '';
    let hrContact = props.getProperty('HR_CONTACT') || '';

    let balanceVal = (monthlyQuota || 0) - (monthTotalSales || 0);
    let displayBalance = balanceVal > 0 ? balanceVal : 0;
    
    const formatCur = (num) => "₱" + Number(num).toLocaleString('en-US', {minimumFractionDigits: 2, maximumFractionDigits: 2});
    const todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MMMM d, yyyy");
    
    // Set dynamic text for email
        let deptText = department ? ` (${department})` : "";
        let currentYear = new Date().getFullYear();
        
        // Incident Ref is now safely constructed conditionally to prevent errors if null
        let refString = (incidentRef && incidentRef.paddedSn) ? `INCIDENT-${currentYear}-${incidentRef.paddedSn}-${incidentRef.randomStr}` : "";

        if (isQuotaHit) {
            let reportDateText = dateRangeStr.split('as of Report Date ')[1] || todayStr;
            
            // NO INCIDENT REF IN SUBJECT
            secSubject = `Letter of Appreciation Sales Quota ${aeName} - ${startDateStr} to ${endDateStr} as of Report Date ${reportDateText}`;
        secBody = `
        <div style="font-family: Arial, sans-serif; font-size: 14px; color: #000; line-height: 1.6;">
            <p>Dear ${greetingName},</p>
            <p><b>Congratulations!</b></p>
            <p>The Management would like to formally recognize and commend your exceptional performance as an Account Executive for successfully achieving your assigned sales quota for the period of ${startDateStr} to ${endDateStr}.</p>
            <p>Your dedication, professionalism, and commitment to delivering results have significantly contributed to the Company's operations and overall business objectives. Your ability to consistently meet performance expectations reflects not only your strong work ethic but also your valuable contribution to the team and the organization.</p>
            <p>The report reflects the following:</p>
            <ul style="margin-top: 0;">
                <li><b>Required Monthly Sales Quota (VAT Exclusive):</b> ${formatCur(monthlyQuota)}</li>
                <li><b>Actual Sales (VAT Exclusive):</b> ${formatCur(monthTotalSales)}</li>
                <li><b>Sales Shortfall:</b> ${formatCur(displayBalance)}</li>
                <li><b>Quota Attainment (%):</b> ${pctAchieved}</li>
            </ul>
            <br>
            <img src="${c1Url}" alt="Sales Chart" style="max-width: 450px; height: auto; border: 1px solid #ccc; border-radius: 4px; margin-top: 10px; margin-bottom: 20px;">
            <br>
            <p>We sincerely appreciate your continued efforts, positive attitude, and commitment to excellence. Your achievement is highly recognized and serves as an encouragement for continued growth and outstanding performance.</p>
            <p>Thank you for your hard work and dedication. We look forward to your continued success with the Company.</p>
            <p>Thank you for your kindness and support.</p>
            <p>Please do not hesitate to contact:<br>
            ${(!hr1Name || hr1Name.trim().toUpperCase() === 'HR') ? 'HR:' : 'HR: ' + hr1Name} &lt;<a href="mailto:${hr1Email}">${hr1Email}</a>&gt;<br>
            HR: ${hr2Name} &lt;<a href="mailto:${hr2Email}">${hr2Email}</a>&gt;<br>
            HR CP #: ${hrContact}</p>
            <br>
            ${feedbackImgBlob ? `<a href="${feedbackLink}" target="_blank"><img src="cid:feedbackImage" alt="Feedback Matters" style="max-width: 350px; height: auto; border: none;"></a>` : ''}
        </div>`;

        // 1. SELECT THE CORRECT TEMPLATE ID
        let templateIdToUseAppr = isA2ZAe ? APPRECIATION_TEMPLATE_ID_A2Z : APPRECIATION_TEMPLATE_ID_PINGCON;

        try {
            // 2. MAKE A TEMPORARY COPY OF THE TEMPLATE
            let tempDocId = DriveApp.getFileById(templateIdToUseAppr).makeCopy("Temp_Appreciation_" + aeName).getId();
            let doc = DocumentApp.openById(tempDocId);
            let body = doc.getBody();

            // 3. REPLACE PLACEHOLDERS
            let replacements = {
                "{{TODAY_DATE}}": Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd MMMM yyyy"),
                "{{AE_NAME}}": nameNoComma,
                "{{INCIDENT_REF}}": "", // Ensure this is completely blank for Appreciation
                "{{START_DATE}}": startDateStr,
                "{{END_DATE}}": endDateStr,
                "{{EMAIL_SUBJECT}}": "Recognition of Sales Quota Achievement",
                "{{FIRST_NAME}}": greetingName,
                "{{QUOTA}}": Number(monthlyQuota).toLocaleString('en-US', {minimumFractionDigits: 2, maximumFractionDigits: 2}),
                "{{ACTUAL}}": Number(monthTotalSales).toLocaleString('en-US', {minimumFractionDigits: 2, maximumFractionDigits: 2}),
                "{{SHORTFALL}}": Number(displayBalance).toLocaleString('en-US', {minimumFractionDigits: 2, maximumFractionDigits: 2}),
                "{{ATTAINMENT}}": pctAchieved.replace('%', '')
            };

            let header = doc.getHeader();
            let footer = doc.getFooter();

            for (let key in replacements) {
                body.replaceText(key, replacements[key]);
                if (header) header.replaceText(key, replacements[key]);
                if (footer) footer.replaceText(key, replacements[key]);
            }

            // COMPLETELY REMOVE THE INCIDENT ROW FOR APPRECIATION LETTERS
            try {
                body.replaceText("(?i)Incident.*\\n", "");
                if (header) header.replaceText("(?i)Incident.*\\n", "");
            } catch(e) {}

            // Insert the Chart Image safely
            try {
                let chartBlob = UrlFetchApp.fetch(c1Url).getBlob();
                let findImage = body.findText("{{CHART_IMAGE}}");
                if (findImage) {
                    let element = findImage.getElement();
                    let parent = element.getParent().asParagraph();
                    element.removeFromParent(); // Completely remove the placeholder text
                    
                    let img = parent.insertInlineImage(0, chartBlob);
                    // Resize chart to fit document margins perfectly
                    img.setWidth(450); 
                    img.setHeight(281);
                }
            } catch (e) {
                body.replaceText("{{CHART_IMAGE}}", ""); // Clear placeholder if image fails to avoid errors
            }

            doc.saveAndClose();
            Utilities.sleep(3000); // Wait for Drive to sync changes

            // 4. EXPORT AS PDF AND DOCX
            let token = ScriptApp.getOAuthToken();
            let cleanAeNameStr = aeName.replace(/[^a-zA-Z0-9]/g, "");
            
            // Export PDF
            let pdfUrl = "https://docs.google.com/document/d/" + tempDocId + "/export?format=pdf";
            let pdfResponse = UrlFetchApp.fetch(pdfUrl, { headers: { 'Authorization': 'Bearer ' + token }, muteHttpExceptions: true });
            let pageCount = 1;
            if (pdfResponse.getResponseCode() === 200) {
                ntePdfBlob = pdfResponse.getBlob().setName(`Appreciation_Letter_${cleanAeNameStr}.pdf`);
                try {
                    let pdfString = ntePdfBlob.getDataAsString("ISO-8859-1");
                    let match = pdfString.match(/\/Type\s*\/Pages[^\>]*?\/Count\s+(\d+)/i) || pdfString.match(/\/Count\s+(\d+)/i);
                    if (match) {
                        let parsed = parseInt(match[1], 10);
                        if (parsed > 0 && parsed < 100) pageCount = parsed;
                    }
                } catch(e) {}
            }

            

            // 5. TRASH TEMPORARY FILE
            DriveApp.getFileById(tempDocId).setTrashed(true);

        } catch (docErr) {
            console.error("Appreciation Document Generation Failed: " + docErr);
        }
    } else {
        // NTE Email (Requires Attachment Generation)
        secSubject = `[INCIDENT ALERT] ${aeName}${deptText} - Ref# ${refString}`;
        
        // 1. SELECT THE CORRECT TEMPLATE ID
        let templateIdToUse = isA2ZAe ? NTE_TEMPLATE_ID_A2Z : NTE_TEMPLATE_ID_PINGCON;

        try {
            // 2. MAKE A TEMPORARY COPY OF THE TEMPLATE
            let tempDocId = DriveApp.getFileById(templateIdToUse).makeCopy("Temp_NTE_" + aeName).getId();
            let doc = DocumentApp.openById(tempDocId);
            let body = doc.getBody();

            // 3. REPLACE PLACEHOLDERS WITH ACTUAL DATA
            let baseDate = new Date();
            const getDayDate = (addDays) => {
                let d = new Date(baseDate);
                d.setDate(d.getDate() + addDays);
                return Utilities.formatDate(d, Session.getScriptTimeZone(), "MM/dd/yyyy");
            };

            let replacements = {
                "{{TODAY_DATE}}": todayStr,
                "{{AE_NAME}}": nameNoComma,
                "{{INCIDENT_REF}}": refString,
                "{{START_DATE}}": startDateStr,
                "{{END_DATE}}": endDateStr,
                "{{EMAIL_SUBJECT}}": `Notice to Explain - Sales Performance for ${startDateStr} to ${endDateStr}`,
                "{{FIRST_NAME}}": greetingName,
                "{{REPORT_DATE}}": dateRangeStr.split('as of Report Date ')[1] || todayStr,
                "{{REPORT_DATE_TOP}}": getDayDate(0),
                "{{QUOTA}}": Number(monthlyQuota).toLocaleString('en-US', {minimumFractionDigits: 2, maximumFractionDigits: 2}),
                "{{ACTUAL}}": Number(monthTotalSales).toLocaleString('en-US', {minimumFractionDigits: 2, maximumFractionDigits: 2}),
                "{{SHORTFALL}}": Number(displayBalance).toLocaleString('en-US', {minimumFractionDigits: 2, maximumFractionDigits: 2}),
                "{{ATTAINMENT}}": pctAchieved.replace('%', '')
            };
            
            // Generate the dates for DAY_0 to DAY_15
            for (let i = 0; i <= 15; i++) {
                replacements[`{{DAY_${i}}}`] = getDayDate(i);
            }

            let header = doc.getHeader();
            let footer = doc.getFooter();

            // Replace in Body, Header, and Footer safely
            for (let key in replacements) {
                body.replaceText(key, replacements[key]);
                if (header) header.replaceText(key, replacements[key]);
                if (footer) footer.replaceText(key, replacements[key]);
            }

            // COMPLETELY REMOVE THE INCIDENT ROW FOR APPRECIATION LETTERS
            try {
                body.replaceText("(?i)Incident.*\\n", "");
                if (header) header.replaceText("(?i)Incident.*\\n", "");
            } catch(e) {}

            // Insert the Chart Image safely
            try {
                let chartBlob = UrlFetchApp.fetch(c1Url).getBlob();
                let findImage = body.findText("{{CHART_IMAGE}}");
                if (findImage) {
                    let element = findImage.getElement();
                    let parent = element.getParent().asParagraph();
                    element.removeFromParent(); // Completely remove the placeholder text
                    
                    let img = parent.insertInlineImage(0, chartBlob);
                    // Resize chart to fit document margins perfectly
                    img.setWidth(450); 
                    img.setHeight(281);
                }
            } catch (e) {
                body.replaceText("{{CHART_IMAGE}}", ""); // Clear placeholder if image fails to avoid errors
            }

            // Save and Close the document to ensure all changes are written to Drive
            doc.saveAndClose();
            
            // CRITICAL FIX: Force Google Drive to wait 3 seconds before exporting to PDF/DOCX
            // This prevents the "cut off" pages and ensures the image finishes saving inside the file
            Utilities.sleep(3000);

            // 4. EXPORT AS PDF AND DOCX
            let token = ScriptApp.getOAuthToken();
            
            // Export PDF
            let pdfUrl = "https://docs.google.com/document/d/" + tempDocId + "/export?format=pdf";
            let pdfResponse = UrlFetchApp.fetch(pdfUrl, { headers: { 'Authorization': 'Bearer ' + token }, muteHttpExceptions: true });
            let pageCount = 5; // Default fallback in case extraction fails
            if (pdfResponse.getResponseCode() === 200) {
                ntePdfBlob = pdfResponse.getBlob().setName(`${refString}.pdf`);
                // Extract total page count from the generated PDF to mirror it into the DOCX
                try {
                    let pdfString = ntePdfBlob.getDataAsString("ISO-8859-1");
                    let match = pdfString.match(/\/Type\s*\/Pages[^\>]*?\/Count\s+(\d+)/i) || pdfString.match(/\/Count\s+(\d+)/i);
                    if (match) {
                        let parsed = parseInt(match[1], 10);
                        if (parsed > 0 && parsed < 100) pageCount = parsed;
                    }
                } catch(e) { console.warn("Could not extract PDF page count: " + e); }
            }
            /*
            // --- CRITICAL FIX FOR DOCX PAGE COUNT PREVIEW ISSUE ---
            // Gmail's DOCX previewer fails to render Google Docs PAGE_COUNT fields (shows as "Page 1 of ").
            // We reopen the doc, locate the PAGE_COUNT element, replace it with the static page count 
            // extracted from the PDF, and save before exporting to DOCX.
            try {
                let fixDoc = DocumentApp.openById(tempDocId);
                let fixH = fixDoc.getHeader();
                let fixF = fixDoc.getFooter();
                
                let fixPageCount = (element) => {
                    if (!element) return;
                    let paragraphs = element.getParagraphs();
                    for (let i = 0; i < paragraphs.length; i++) {
                        let p = paragraphs[i];
                        let text = p.getText();
                        if (text.match(/Page\s+.*of/i)) {
                            let numChildren = p.getNumChildren();
                            for (let j = numChildren - 1; j >= 0; j--) {
                                let child = p.getChild(j);
                                // PAGE_COUNT usually resolves as UNSUPPORTED in Apps Script
                                if (child.getType() === DocumentApp.ElementType.UNSUPPORTED) {
                                    let isPageCount = false;
                                    for (let k = j - 1; k >= 0; k--) {
                                        let prev = p.getChild(k);
                                        if (prev.getType() === DocumentApp.ElementType.TEXT) {
                                            let t = prev.getText();
                                            if (t.trim() === "of" || t.endsWith("of ")) {
                                                isPageCount = true;
                                                break;
                                            } else if (t.trim() !== "") {
                                                break;
                                            }
                                        }
                                    }
                                    if (isPageCount) {
                                        p.insertText(j, String(pageCount)).setBold(true);
                                        child.removeFromParent();
                                    }
                                }
                            }
                        }
                    }
                };
                
                fixPageCount(fixH);
                fixPageCount(fixF);
                fixDoc.saveAndClose();
                Utilities.sleep(3000); // Wait for changes to propagate to Drive
            } catch(e) {
                console.warn("Could not replace PAGE_COUNT for DOCX: " + e);
            }
            // --- END DOCX FIX ---

            // Export DOCX
            let docxUrl = "https://docs.google.com/document/d/" + tempDocId + "/export?format=docx";
            let docxResponse = UrlFetchApp.fetch(docxUrl, { headers: { 'Authorization': 'Bearer ' + token }, muteHttpExceptions: true });
            if (docxResponse.getResponseCode() === 200) {
                nteDocxBlob = docxResponse.getBlob().setName(`${refString}.docx`);
            }
            */
            // 5. TRASH THE TEMPORARY FILE
            DriveApp.getFileById(tempDocId).setTrashed(true);

        } catch (docErr) {
            console.error("Document Generation Failed: " + docErr);
        }

        let reportDateText = dateRangeStr.split('as of Report Date ')[1] || todayStr;
        
        secBody = `
        <div style="font-family: Arial, sans-serif; font-size: 14px; color: #000; line-height: 1.6; max-width: 800px;">
            <p>Dear ${greetingName},</p>
            <p>Good day!</p>
            <p>Greetings.</p>
            <p>Based on the Incident Report and the Sales Performance Report covering the period <b>${startDateStr} to ${endDateStr}</b> (Report Date: ${reportDateText}), it was noted that your sales performance for the month did not meet the required sales quota under the Company's Sales Policy.</p>
            <p>The report reflects the following:</p>
            <ul>
                <li><b>Required Monthly Sales Quota (VAT Exclusive):</b> ${formatCur(monthlyQuota)}</li>
                <li><b>Actual Sales (VAT Exclusive):</b> ${formatCur(monthTotalSales)}</li>
                <li><b>Sales Shortfall:</b> ${formatCur(displayBalance)}</li>
                <li><b>Quota Attainment (%):</b> ${pctAchieved}</li>
            </ul>
            <br>
            <img src="${c1Url}" alt="Sales Chart" style="max-width: 450px; height: auto; border: 1px solid #ccc; border-radius: 4px; margin-top: 10px; margin-bottom: 20px;">
            <br>
            <p>As part of our standard process, we would like to understand the circumstances that may have affected your sales performance during the period. You are therefore requested to submit a written explanation describing the factors that contributed to the result, as well as any challenges encountered and any actions you took to improve or recover your sales.</p>
            <p>If there were circumstances beyond your control that affected your ability to achieve the required quota, please include them in your explanation together with any supporting information you believe may be helpful.</p>
            <p>Kindly submit your written explanation within the period prescribed under company policy to allow Management to review the matter fairly and consider all relevant circumstances before determining the appropriate next steps.</p>
            <p>This Notice to Explain is issued to provide you with an opportunity to share your side and to ensure that the matter is reviewed objectively and fairly.</p>
            <p>Should you have any questions regarding this notice, you may coordinate with your immediate supervisor or the Human Resources Department.</p>
            <p>You are hereby given five (5) calendar days from the date of this letter to submit your written explanation.</p>
            <div style="background-color: #ffe6e6; border-left: 4px solid #cc0000; padding: 10px; margin: 15px 0;">
                <strong style="color: #cc0000;">ACTION REQUIRED:</strong><br>
                <b>Please reply directly to this email with your explanation. Do not change the subject line.</b> If you have any supporting documents (e.g., medical certificates or screenshots), please attach them to your reply.
            </div>
            <p>If no written explanation is received within the prescribed period, the Company will proceed with Steps 3 and 4 of the process based on the available records.</p>
            <p>Thank you for your kindness and support.</p>
            <p>Please do not hesitate to contact:<br>
            ${(!hr1Name || hr1Name.trim().toUpperCase() === 'HR') ? 'HR:' : 'HR: ' + hr1Name} &lt;<a href="mailto:${hr1Email}">${hr1Email}</a>&gt;<br>
            HR: ${hr2Name} &lt;<a href="mailto:${hr2Email}">${hr2Email}</a>&gt;<br>
            HR CP #: ${hrContact}</p>
        </div>`;
    }

    return {
        secSubject: secSubject,
        secBody: secBody,
        ntePdfBlob: ntePdfBlob,
        nteDocxBlob: nteDocxBlob
    };
}

function htmlToDocxBlob(htmlStr, filename) {
    try {
        // Strip complex CSS that causes Google Docs conversion to crash
        let safeHtml = htmlStr.replace(/@page\s*{[^}]+}/gi, "");

        let blob = Utilities.newBlob(safeHtml, MimeType.HTML, "temp.html");
        
        // Force Google Drive to parse HTML into a native Google Doc
        let file = Drive.Files.insert({title: "TempDoc", mimeType: MimeType.GOOGLE_DOCS}, blob, {convert: true});
        
        Utilities.sleep(3000); // Wait for Drive to finish processing
        
        let docxBlob = null;
        let token = ScriptApp.getOAuthToken();
        let lastStatus = 0;
        
        // OFFICIAL API ENDPOINT: This guarantees a proper DOCX export without login/token rejections
        let exportUrl = "https://www.googleapis.com/drive/v3/files/" + file.id + "/export?mimeType=application/vnd.openxmlformats-officedocument.wordprocessingml.document";
        
        for (let i = 0; i < 3; i++) {
            let response = UrlFetchApp.fetch(exportUrl, { 
                headers: { 'Authorization': 'Bearer ' + token }, 
                muteHttpExceptions: true 
            });
            
            lastStatus = response.getResponseCode();
            if (lastStatus === 200) {
                docxBlob = response.getBlob().setName(filename);
                break;
            }
            Utilities.sleep(3000);
        }
        
        // Cleanup temp file
        Drive.Files.remove(file.id);
        
        if (!docxBlob) throw new Error("API Export Failed with status: " + lastStatus);
        return docxBlob;
    } catch(e) {
        console.error("Docx generation failed: " + e);
        return null;
    }
}
function createNteExcelAttachment(refId, dateStr) {
    let filename = `NTE schedule ${refId}`;
    let ss = SpreadsheetApp.create(filename);
    
    // Tab 1: 00readme
    let sheet1 = ss.getSheets()[0];
    sheet1.setName("00readme");
    sheet1.getRange("B4").setValue("06/14/2026");
    sheet1.getRange("B6").setValue("What is the company policy? Please follow the company policy? Please scan the policy and email back to me. Please send update.");
    sheet1.getRange("B9").setValue("Communication is always 4 ways:");
    sheet1.getRange("B10").setValue("1. email to email to email and");
    sheet1.getRange("B11").setValue("2. viber to viber to viber and");
    sheet1.getRange("B12").setValue("3. voice to voice to voice and");
    sheet1.getRange("B13").setValue("4. Timestamp app and please see attached Timestamp configuration");
    sheet1.getRange("B19").setValue("Please scan the NTE and Disciplinary action and all the other relevant documents. To ensure proper documentation and compliance, please scan and email");
    sheet1.getRange("C21").setValue("1 Proof of the NTE itself along with the explanation");
    sheet1.getRange("C22").setValue("2 Proof of discussion with the employee (face-to-face or during an administrative hearing)");
    sheet1.getRange("C23").setValue("3 All the other documents");
    sheet1.getRange("C24").setValue("4 Proof that the employee received the disciplinary action");
    sheet1.getRange("C25").setValue("5 Please provide a complete set of documents because each NTE requires a black-and-white copy for closure, and all supporting documents must be properly filed in 201. Having everything organized and complete helps avoid future discrepancies, ensures transparency, and allows for proper tracking of employee records.");
    
    // Tab 2: Notice to Explain
    let sheet2 = ss.insertSheet("Notice to Explain");
    sheet2.getRange("B4").setValue("4. AWOL (Absent Without Official Leave)").setFontWeight("bold");
    sheet2.getRange("B5").setValue("The Company understands that there may be times when employees are unable to report for work due to unexpected or personal circumstances. We encourage open communication so we can provide appropriate support whenever possible.");
    sheet2.getRange("B6").setValue("Definition:").setFontWeight("bold");
    sheet2.getRange("B7").setValue("AWOL occurs when an employee is absent from work and is unable to notify their supervisor within a reasonable period or does not provide an explanation for their absence.");
    sheet2.getRange("B8").setValue("Guidelines:").setFontWeight("bold");
    sheet2.getRange("B10").setValue("Employees are encouraged to inform their supervisor as early as possible if they cannot report for work.");
    sheet2.getRange("B11").setValue("In emergency situations where prior notice is not possible, employees should reach out as soon as they are able.");
    sheet2.getRange("B12").setValue("If an employee is absent without communication for three consecutive days, the Company may temporarily classify the absence as AWOL while efforts are made to establish contact.");
    sheet2.getRange("B13").setValue("Our Approach:").setFontWeight("bold");
    sheet2.getRange("B15").setValue("The Company will make a sincere effort to check in and reach out to the employee to understand their situation.");
    sheet2.getRange("B16").setValue("We recognize that circumstances such as illness, family matters, or unforeseen events may affect attendance.");
    sheet2.getRange("B17").setValue("Employees will always be given a chance to explain before any conclusions are made.");
    sheet2.getRange("B18").setValue("Fair Process:").setFontWeight("bold");
    sheet2.getRange("B19").setValue("Consistent with standards of the Department of Labor and Employment (DOLE), the Company follows a fair and respectful process:");
    sheet2.getRange("C21").setValue("1. The employee will be contacted or issued a Notice to Explain (NTE)");
    sheet2.getRange("C22").setValue("2. The employee is given time to share their explanation");
    sheet2.getRange("C23").setValue("3. The Company will carefully review all information before making a decision");
    sheet2.getRange("B24").setValue("Resolution:").setFontWeight("bold");
    sheet2.getRange("B25").setValue("The Company aims to resolve AWOL cases with understanding and fairness. Depending on the circumstances, outcomes may include clarification of records, support measures, or, when necessary, appropriate disciplinary action.");

    // Tab 3: Incident, Steps 1, 2, 3, 4
    let sheet3 = ss.insertSheet("Incident, Steps 1, 2, 3, 4");
    sheet3.getRange("B2").setValue("How to use:");
    sheet3.getRange("D2").setValue("Change the Date of the incident report. Its highlighted in yellow text. Cell D5");
    sheet3.getRange("B5").setValue("Date Incident report (mm/dd/yyyy)").setFontWeight("bold").setBorder(true,true,true,true,false,false);
    
    // Explicitly set as plain text using "@" so Excel doesn't misinterpret the date format
    let d5 = sheet3.getRange("D5");
    d5.setNumberFormat("@").setValue(dateStr).setBackground("yellow").setBorder(true,true,true,true,false,false);
    
    sheet3.getRange("B6").setValue("Incident").setFontWeight("bold").setBorder(true,true,true,true,false,false);
    sheet3.getRange("D6").setValue(refId).setBorder(true,true,true,true,false,false);
    sheet3.getRange("B7").setValue("Due Process").setFontWeight("bold").setBorder(true,true,true,true,false,false);
    sheet3.getRange("D7").setValue("Schedule - Step 1 - Notice To Explain, Step 2 - Employee explanation, Step 3 - Conference, Step 4 - Closure and or Disciplinary Action").setBorder(true,true,true,true,false,false);
    sheet3.getRange("B8").setValue("Due Date per step").setFontWeight("bold").setBorder(true,true,true,true,false,false);
    sheet3.getRange("D8").setValue("Incident Report follows a four-step process. Each succeeding step may begin earlier if the maximum allotted number of days for the current step is not fully utilized.").setBorder(true,true,true,true,false,false);
    
    let headers = ["Day", "Calendar Date (mm/dd/yyyy)", "Days Elapsed", "Incident Report Due Process Steps 1, 2, 3, 4"];
    sheet3.getRange("A10:D10").setValues([headers]).setFontWeight("bold").setHorizontalAlignment("center").setBorder(true,true,true,true,true,true);
    
    let daysData = [
      ["Day 0", "=D5", 1, "Step 1 Supervisor HR issue Notice to Explain", "#f4b084"],
      ["Day 1", "=B11+1", 1, "Step 2 Receive employee explanation within five calendar days no extension (weekends/holidays included).", "#92d050"],
      ["Day 2", "=B12+1", 2, "Step 2 Receive employee explanation within five calendar days no extension (weekends/holidays included).", "#92d050"],
      ["Day 3", "=B13+1", 3, "Step 2 Receive employee explanation within five calendar days no extension (weekends/holidays included).", "#92d050"],
      ["Day 4", "=B14+1", 4, "Step 2 Receive employee explanation within five calendar days no extension (weekends/holidays included).", "#92d050"],
      ["Day 5", "=B15+1", 5, "Step 2 Receive employee explanation within five calendar days no extension (weekends/holidays included).", "#92d050"],
      ["Day 6", "=B16+1", 1, "Step 3 Conference within five calendar days no extension (weekends/holidays included).", "#d9d9d9"],
      ["Day 7", "=B17+1", 2, "Step 3 Conference within five calendar days no extension (weekends/holidays included).", "#d9d9d9"],
      ["Day 8", "=B18+1", 3, "Step 3 Conference within five calendar days no extension (weekends/holidays included).", "#d9d9d9"],
      ["Day 9", "=B19+1", 4, "Step 3 Conference within five calendar days no extension (weekends/holidays included).", "#d9d9d9"],
      ["Day 10", "=B20+1", 5, "Step 3 Conference within five calendar days no extension (weekends/holidays included).", "#d9d9d9"],
      ["Day 11", "=B21+1", 1, "Step 4 Closure - accept employee explanation and or issue Disciplinary action within five calendar days no extension (weekends/holidays included).", "#f4b084"],
      ["Day 12", "=B22+1", 2, "Step 4 Closure - accept employee explanation and or issue Disciplinary action within five calendar days no extension (weekends/holidays included).", "#f4b084"],
      ["Day 13", "=B23+1", 3, "Step 4 Closure - accept employee explanation and or issue Disciplinary action within five calendar days no extension (weekends/holidays included).", "#f4b084"],
      ["Day 14", "=B24+1", 4, "Step 4 Closure - accept employee explanation and or issue Disciplinary action within five calendar days no extension (weekends/holidays included).", "#f4b084"],
      ["Day 15", "=B25+1", 5, "Step 4 Closure - accept employee explanation and or issue Disciplinary action within five calendar days no extension (weekends/holidays included).", "#f4b084"]
    ];
    
    for (let i = 0; i < daysData.length; i++) {
        let row = 11 + i;
        sheet3.getRange(`A${row}`).setValue(daysData[i][0]).setBackground(daysData[i][4]).setBorder(true,true,true,true,true,true).setHorizontalAlignment("center");
        sheet3.getRange(`B${row}`).setFormula(daysData[i][1]).setNumberFormat("mm/dd/yyyy").setBackground(daysData[i][4]).setBorder(true,true,true,true,true,true).setHorizontalAlignment("center");
        sheet3.getRange(`C${row}`).setValue(daysData[i][2]).setBackground(daysData[i][4]).setBorder(true,true,true,true,true,true).setHorizontalAlignment("center");
        sheet3.getRange(`D${row}`).setValue(daysData[i][3]).setBackground(daysData[i][4]).setBorder(true,true,true,true,true,true);
    }
    
    sheet3.setColumnWidth(1, 80);
    sheet3.setColumnWidth(2, 180);
    sheet3.setColumnWidth(3, 100);
    sheet3.setColumnWidth(4, 600);
    
    SpreadsheetApp.flush();
    
    let url = "https://docs.google.com/spreadsheets/d/" + ss.getId() + "/export?format=xlsx";
    let token = ScriptApp.getOAuthToken();
    let blob = UrlFetchApp.fetch(url, { headers: { 'Authorization': 'Bearer ' + token }, muteHttpExceptions: true }).getBlob().setName(filename + ".xlsx");
    DriveApp.getFileById(ss.getId()).setTrashed(true);
    return blob;
}

function Template_UnpaidDrAeEmail(totalDrCount, totalOpenBalance, fullyUnpaidCount, partiallyUnpaidCount, chartHtml, feedbackImgBlob, feedbackLink) {
  let formatCurrency = (val) => Number(val).toLocaleString('en-US', {minimumFractionDigits: 2, maximumFractionDigits: 2});
  return `
    <div style="font-family: Arial, sans-serif; font-size: 13px; color: #000; line-height: 1.5;">
      <p>Good day!</p>
      <p>We hope you are doing well. We are reaching out to share a summary of your account with us, which includes recent payments received as well as any outstanding balances that may still require your attention.</p>
      <div style="background-color: #f9f9f9; padding: 15px; border: 1px solid #ddd; border-radius: 5px; margin: 20px 0; width: fit-content;">
          <h3 style="margin-top: 0; color: #333;">Account Summary Statistics</h3>
          <table style="font-size: 13px;">
              <tr><td style="padding-right: 20px;"><strong>Total Unpaid DRs:</strong></td><td>${totalDrCount}</td></tr>
              <tr><td style="padding-right: 20px;"><strong>Total Open Balance:</strong></td><td style="color: #d9534f; font-weight: bold;">PHP ${formatCurrency(totalOpenBalance)}</td></tr>
              <tr><td style="padding-right: 20px;"><strong>Num of clients 100% unpaid:</strong></td><td>${fullyUnpaidCount}</td></tr>
              <tr><td style="padding-right: 20px;"><strong>Num of clients partially unpaid:</strong></td><td>${partiallyUnpaidCount}</td></tr>
          </table>
      </div>
      ${chartHtml}
      <p>Please note that some payments may have been made but are not yet reflected in our records at the time of this report. If you have already settled any of the listed items, kindly disregard those entries or let us know so we can update our system accordingly.</p>
      <p>Should you have any questions or need assistance in reviewing your account, we would be more than happy to help.</p>
      <p>Thank you for your continued partnership, and we appreciate your time and attention to this matter.</p>
      <br>
      ${feedbackImgBlob ? `<a href="${feedbackLink}" target="_blank"><img src="cid:feedbackImage" alt="Feedback Matters" style="max-width: 350px; height: auto; border: none;"></a>` : ''}
    </div>`;
}

function Template_UnpaidDrCustEmail(cleanCustName, dateRangePhrase, totalDrCount, totalDue, totalAdvance, totalOpenBalance, chartHtml, aeContactHtml, collectionAsstHtml, feedbackImgBlob, feedbackLink, senderEmail) {
  let formatCurrency = (val) => Number(val).toLocaleString('en-US', {minimumFractionDigits: 2, maximumFractionDigits: 2});
  
  let holidayNotice = "";
    if (Date.now() <= new Date("August 31, 2026 23:59:59").getTime()) {
        holidayNotice = `
        <p style="font-weight: bold; color: #d93025; margin-bottom: 5px;">COMPANY BREAK NOTICE</p>
        <p style="margin-top: 0; margin-bottom: 5px;">Please be advised that our company will be on break on the following date:</p>
        <ul style="margin-top: 5px; margin-bottom: 10px;">
            <li><strong>August 31, 2026</strong> &mdash; National Heroes Day</li>
        </ul>
        <p style="margin-top: 0;">Please expect our office to be closed on this date. Regular operations will resume on the next business day.</p>`;
    }

  let bottomFeedbackHtml = feedbackImgBlob ? `<br><a href="${feedbackLink}" target="_blank"><img src="cid:feedbackImage" alt="Feedback Matters" style="max-width: 350px; height: auto; border: none;"></a>` : '';

  // Dynamically generate the text block based on the entity sending the email
  let catalogTextHtml = "";
  let a2zSenderEmail = (PropertiesService.getScriptProperties().getProperty('SENDER_EMAIL_A2Z') || "").toLowerCase().trim();
  let isA2ZSender = (senderEmail && a2zSenderEmail && senderEmail.toLowerCase().trim() === a2zSenderEmail);

  if (isA2ZSender) {
      catalogTextHtml = `
      <div style="margin-top: 20px; margin-bottom: 20px; font-family: Arial, sans-serif; font-size: 13px;">
          ------------<br>
          Happy Holidays!<br>
          <strong>Pingcon Basket Catalogue</strong><br>
          <em>Brought to you by Pingcon, an A2Z2GO sister company.</em><br>
          ------------
      </div>`;
  } else {
      catalogTextHtml = `
      <div style="margin-top: 20px; margin-bottom: 20px; font-family: Arial, sans-serif; font-size: 13px;">
          ----------<br>
          <strong>Pingcon Basket Catalogue</strong><br>
          <em>Your Trusted Noche Buena Partner</em><br>
          ----------
      </div>`;
  }

  return `
    <div style="font-family: Arial, sans-serif; font-size: 13px; color: #000; line-height: 1.5;">
      <p>Dear ${cleanCustName},</p>
      <p>We hope this message finds you well.</p>
      <p>Our report indicates that some payments have already been received, while others may still be outstanding. If payment has been made, please accept our sincere thanks—we will update our records accordingly.</p>
      <p>If any payments are still pending, we would appreciate hearing from you regarding any concerns or issues. Please feel free to reach out at your convenience; we are always happy to assist.</p>
      <p>Please review the report, which outlines the list of deliverables. Kindly confirm whether your records align with ours. The PO number can be found under Column H (“Additional Text”).</p>
      ${dateRangePhrase}
      <div style="background-color: #f9f9f9; padding: 15px; border: 1px solid #ddd; border-radius: 5px; margin: 20px 0; width: fit-content;">
          <h3 style="margin-top: 0; color: #333;">Account Summary Statistics</h3>
          <table style="font-size: 13px;">
              <tr><td style="padding-right: 20px;"><strong>Total Unpaid DRs:</strong></td><td>${totalDrCount}</td></tr>
              <tr><td style="padding-right: 20px;"><strong>Due For Collection:</strong></td><td style="font-weight: bold;">PHP ${formatCurrency(totalDue)}</td></tr>
              <tr><td style="padding-right: 20px;"><strong>Advance Collection:</strong></td><td style="font-weight: bold;">PHP ${formatCurrency(totalAdvance)}</td></tr>
              <tr><td style="padding-right: 20px;"><strong>Total Open Balance:</strong></td><td style="color: #d9534f; font-weight: bold;">PHP ${formatCurrency(totalOpenBalance)}</td></tr>
          </table>
      </div>
      ${holidayNotice}
      ${chartHtml}
      <p>Thank you for your continued support. We truly value your business.</p>
      
      ${catalogTextHtml}
      
      <p>Please do not hesitate to contact:<br>${aeContactHtml}${collectionAsstHtml}</p>
      <br><p>Warm regards,</p>
      ${bottomFeedbackHtml}
    </div>`;
}

function Template_CommissionEmail(cleanAeName, commHeaders, tableRowsHtml, totalComm, rowsLength, feedbackImgBlob, feedbackLink) {
  return `
    <div style="margin-bottom: 50px;">
    <p>Hi ma'am/sir <strong>${cleanAeName}</strong>,</p>
    <p>Please see Commission Report below.<br>Kindly review the computations for the generated invoices.</p>
    <table style="border-collapse: collapse; width: 100%; font-size: 12px;">
        <thead>
        <tr style="background-color: #71b9c9; color: black;">
            ${commHeaders.map(h => `<th style="border: 1px solid black; padding: 4px;">${h}</th>`).join("")}
        </tr>
        </thead>
        <tbody>${tableRowsHtml}</tbody>
        <tfoot>
        <tr>
            <td colspan="19" style="border: none;"></td>
            <td style="border-bottom: 3px double black; border-top: 1px solid black; padding: 4px; font-weight: bold; text-align: right;">Total</td>
            <td style="border-bottom: 3px double black; border-top: 1px solid black; padding: 4px; font-weight: bold; color: red; text-align: right;">${Number(totalComm).toLocaleString('en-US', {minimumFractionDigits: 2, maximumFractionDigits: 2})}</td>
        </tr>
        </tfoot>
    </table><br>
    ${feedbackImgBlob ? `<a href="${feedbackLink}" target="_blank"><img src="cid:feedbackImage" alt="Feedback Matters" style="max-width: 350px; height: auto; border: none;"></a>` : ''}
    </div>`;
}

function Template_BuySellEmail(aeName, monthYear, feedbackImgBlob, feedbackLink) {
  return `
    <div style="font-family: Arial, sans-serif; font-size: 13px; color: #000; line-height: 1.5;">
      <p>Good day ${aeName},</p>
      <p>Please find attached your Buy/Sell Report for the month of <strong>${monthYear}</strong>.</p>
      <p>Kindly review the details. Should you have any questions, please let us know.</p>
      <p>Thank you.</p>
      <br>
      ${feedbackImgBlob ? `<a href="${feedbackLink}" target="_blank"><img src="cid:feedbackImage" alt="Feedback Matters" style="max-width: 350px; height: auto; border: none;"></a>` : ''}
    </div>`;
}

function Template_NteHtml(data) {
  const { date, recipient, subject, problemDetails, link, isReminder, mode, explanation, signature, printedName, newPrices } = data;
  const isPdfMode = mode === 'pdf';
  const pinkColor = "#F50057";
  
  let itemsTableHtml = "";
  if (newPrices && newPrices.length > 0) {
    let rows = newPrices.map((p, i) => `<tr><td style="border:1px solid #000; text-align:center; padding:4px;">${i + 1}</td><td style="border:1px solid #000; padding:4px;">${p.po||'-'}</td><td style="border:1px solid #000; padding:4px;">${p.sku}</td><td style="border:1px solid #000; padding:4px;">${p.desc}</td><td style="border:1px solid #000; padding:4px;">${parseFloat(p.price).toFixed(2)}</td></tr>`).join("");
    itemsTableHtml = `<table style="width:100%; border-collapse:collapse; font-size:12px; margin-top:10px; margin-bottom:15px;"><tr><th style="border:1px solid #000; padding:4px;">SN</th><th style="border:1px solid #000; padding:4px;">PO</th><th style="border:1px solid #000; padding:4px;">SKU</th><th style="border:1px solid #000; padding:4px;">Desc</th><th style="border:1px solid #000; padding:4px;">Price</th></tr>${rows}</table>`;
  }

  return `
     <!DOCTYPE html><html><head><style>
         @import url('https://fonts.googleapis.com/css2?family=Dancing+Script:wght@700&family=Arial&display=swap');
         body { font-family: Arial, sans-serif; color: #000; padding: 20px; }
         .container { background-color: #ffffff; width: 100%; max-width: 800px; margin: 0 auto; padding: 40px; }
         .company-name { color: ${pinkColor}; font-family: 'Dancing Script', cursive, Georgia, serif; font-size: 28px; font-weight: bold; margin: 0; }
         .company-address { font-size: 11px; margin-top: 2px; }
         .main-title { text-align: center; font-weight: bold; font-size: 18px; margin: 30px 0; }
         .meta-table { width: 100%; border-collapse: collapse; font-size: 14px; margin-bottom: 10px; }
         .meta-table td { padding-bottom: 5px; }
         .explanation-container { border: 1px solid #000; min-height: 120px; margin: 20px 0; padding: 15px; }
         .cta-button { background-color: ${pinkColor}; color: white; padding: 10px 20px; text-decoration: none; display: inline-block; margin-top: 20px; }
         .footer-line { border-top: 2px solid #000; padding-top: 5px; margin-top: 40px; }
       </style></head><body style="background-color: #f2f2f2;">
       <div class="container">
         <? let nteAddr = (PropertiesService.getScriptProperties().getProperty('PINGCON_ADDR') || PropertiesService.getScriptProperties().getProperty('KEYSYS_ADDR') || "").replace(/[\r\n]+/g, '<br>'); ?>
         <div class="company-name">Pingcon Marketing Corporation</div>
         <div class="company-address">${(PropertiesService.getScriptProperties().getProperty('PINGCON_ADDR') || PropertiesService.getScriptProperties().getProperty('KEYSYS_ADDR') || "").replace(/[\r\n]+/g, '<br>')}</div>
         <div class="main-title">NOTICE TO EXPLAIN ${isReminder ? "(REMINDER)" : ""}</div>
         <table class="meta-table">
           <tr><td style="width: 80px;">Date:</td><td>${date}</td></tr>
           <tr><td>Attention:</td><td>${recipient}</td></tr>
           <tr><td>From:</td><td>System Administrator</td></tr>
           <tr><td>Subject:</td><td>${subject}</td></tr>
         </table>
         <div style="border-top: 2px solid #000; margin-bottom: 20px;"></div>
         <p style="font-size: 14px;">Good day!</p>
         <p style="font-size: 14px;">This is a written notice regarding: <strong>${problemDetails}</strong></p>
         ${itemsTableHtml}
         <div class="explanation-container">
           <p style="margin-top: 0; font-weight: bold;">Explanation:</p>
           ${isPdfMode ? explanation : `<div style="margin-top: 20px;"><a href="${link}" class="cta-button">CLICK HERE TO REPLY</a></div>`}
         </div>
         <p style="font-size: 14px; text-align: justify; line-height: 1.5;">You are hereby given a chance within five (5) calendar days to explain your side. If no explanation within the given period, you are clearly waiving your rights to be heard and give your side. With this, you are allowing Pingcon to proceed to the next step of this Notice to Explain.</p>
         <p style="font-size: 14px; margin-bottom: 40px;">For your strict compliance.</p>
         <div class="footer-line">
           <p style="font-size: 12px; font-style: italic; margin-top: 0;">Received. Read and understood.</p>
           <table style="width: 100%; margin-top: 40px; text-align: center; font-size: 12px;">
             <tr>
               <td style="width: 30%; border-top: 1px solid #000; padding-top: 5px;">${isPdfMode && printedName ? printedName : 'Printed Name'}</td>
               <td style="width: 5%;"></td>
               <td style="width: 30%; border-top: 1px solid #000; padding-top: 5px;">${isPdfMode && signature ? signature : 'Signature'}</td>
               <td style="width: 5%;"></td>
               <td style="width: 30%; border-top: 1px solid #000; padding-top: 5px;">Date</td>
             </tr>
           </table>
         </div>
         <div style="margin-top: 60px; font-size: 14px;">
           <p>Thank you.</p>
           <table style="width: 100%; margin-top: 40px; font-size: 14px;">
             <tr><td style="width: 50%;">Prepared by:</td><td style="width: 50%;">Approved by:</td></tr>
           </table>
         </div>
       </div></body></html>`;
}

function Template_DialogUI(title, hasAe, hasReq, srvFunc, btnText, validDataMap) {
  return `
    <!DOCTYPE html><html><head><base target="_top"><style>
      body{font-family:Arial,sans-serif;padding:10px;background-color:#f3f3f3;}
      .form-group{margin-bottom:15px;} label{display:block;font-weight:bold;margin-bottom:5px;}
      select,input[type="date"]{width:100%;padding:8px;border-radius:4px;border:1px solid #ccc;box-sizing:border-box;}
      button{background-color:#4CAF50;color:white;padding:10px 20px;border:none;border-radius:4px;cursor:pointer;width:100%;font-size:16px;}
      button:hover{background-color:#45a049;}
      .error{color:red;font-size:12px;margin-top:5px;display:none;}
      #loading{display:none;text-align:center;margin-top:10px;color:#666;}
    </style></head><body>
      <h3>${title}</h3>
      ${hasAe ? `<div class="form-group"><label>Select Account Executive:</label><select id="aeSelect" onchange="updateMonths()"><option value="" disabled selected>-- Choose AE --</option></select></div>` : ''}
      <div class="form-group"><label>Select Month to Process:</label><select id="monthSelect"><option value="" disabled selected>-- Choose Month --</option></select><div id="monthError" class="error">No available data for this AE in the selected month.</div></div>
      ${hasReq ? `<div class="form-group"><label>Commission Request Date:</label><input type="date" id="requestDate"></div>` : ''}
      <button onclick="process()">${btnText}</button>
      <div id="serverError" class="error" style="text-align:center; margin-top:10px; font-weight:bold;"></div>
      <div id="loading">Processing... Please wait.<br>This window will close automatically.</div>
      <script>
        var validData = ${JSON.stringify(validDataMap)};
        var hasAe = ${hasAe}; var hasReq = ${hasReq}; var srvFunc = "${srvFunc}";
        window.onload = function() {
          if(hasAe) {
            var aeSelect = document.getElementById("aeSelect");
            Object.keys(validData).sort().forEach(function(ae) { if (ae !== "_SYSTEM_DATES_") aeSelect.add(new Option(ae, ae)); });
          }
          populateMonthOptions();
        };
        function populateMonthOptions() {
          var mSel = document.getElementById("monthSelect");
          while(mSel.options.length > 1) mSel.remove(1);
          var allM = new Set();
          for(var ae in validData) {
            if (ae === "_SYSTEM_DATES_" && !srvFunc.includes("EWTReportSupplier") && !srvFunc.includes("PurchaseReport")) continue;
            validData[ae].forEach(function(m) { allM.add(m); });
          }
          Array.from(allM).sort(function(a,b) { return new Date(a) - new Date(b); }).forEach(function(m) { mSel.add(new Option(m, m)); });
        }
        function updateMonths() {
          if (!hasAe) return;
          var mSel = document.getElementById("monthSelect");
          var ae = document.getElementById("aeSelect").value;
          var currentVal = mSel.value;
          while(mSel.options.length > 1) mSel.remove(1);
          var months = validData[ae] ||[];
          months.sort(function(a,b) { return new Date(a) - new Date(b); }).forEach(function(m) { mSel.add(new Option(m, m)); });
          if(months.includes(currentVal)) mSel.value = currentVal;
        }
        function process() {
          var m = document.getElementById("monthSelect").value;
          var ae = hasAe ? document.getElementById("aeSelect").value : null;
          var req = hasReq ? document.getElementById("requestDate").value : null;
          if (hasAe && document.getElementById("monthError")) document.getElementById("monthError").style.display = "none";
          if (!m || (hasAe && !ae)) { alert("Please select required fields."); return; }
          if (hasReq && !req) { alert("Please select the Commission Request Date."); return; }
          if (hasAe) {
            var available = validData[ae] ||[];
            if (!available.includes(m)) { if(document.getElementById("monthError")) document.getElementById("monthError").style.display = "block"; return; }
          }
          document.querySelector("button").disabled = true; document.querySelector("button").innerText = "Running...";
          document.getElementById("loading").style.display = "block"; document.getElementById("serverError").style.display = "none";
          var runner = google.script.run.withFailureHandler(function(err){
            document.getElementById("loading").style.display = "none";
            var btn = document.querySelector("button"); btn.disabled = false; btn.innerText = "Try Again";
            var errDiv = document.getElementById("serverError"); errDiv.innerText = "Error: " + (err.message || err); errDiv.style.display = "block";
          }).withSuccessHandler(function(){ google.script.host.close(); });
          if (hasAe && hasReq) runner[srvFunc](ae, m, req); else if (!hasAe && hasReq) runner[srvFunc](m, req);
          else if (hasAe && !hasReq) runner[srvFunc](ae, m); else runner[srvFunc](m);
        }
      </script>
    </body></html>`;
}



function Template_UnpaidDrOverdueEmail(bucketName, totalAmount) {
  let formatCurrency = (val) => Number(val).toLocaleString('en-US', {minimumFractionDigits: 2, maximumFractionDigits: 2});
  return `
    <div style="font-family: Arial, sans-serif; font-size: 13px; color: #000; line-height: 1.5;">
      <p>Good day,</p>
      <p>Please find attached the consolidated Statement of Account for <strong>${bucketName}</strong>.</p>
      <p>This report contains all unpaid balances for all customers within this entity, sorted by Customer Name and Invoice Date.</p>
      <div style="background-color: #f9f9f9; padding: 15px; border: 1px solid #ddd; border-radius: 5px; margin: 20px 0; width: fit-content;">
          <table style="font-size: 13px;">
              <tr>
                  <td style="padding-right: 20px;"><strong>Overall Grand Total:</strong></td>
                  <td style="color: #d93025; font-weight: bold;">PHP ${formatCurrency(totalAmount)}</td>
              </tr>
          </table>
      </div>
      <p>Thank you.</p>
    </div>`;
}
