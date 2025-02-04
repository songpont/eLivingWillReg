/**
 * Main function to create PDF from form submission and send via email
 * This script takes data from a Google Sheet (populated by form responses)
 * and fills it into a Google Doc template, then converts to PDF and emails it
 * Design By Songpon Tulata (songpont@gmail.com)
 */
function createAndSendPDF() {
  try {
    // Configuration - Sheet settings
    const sheetName = "answer01";  // Name of the sheet containing form responses
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    
    // Validate sheet exists
    if (!sheet) {
      throw new Error('Sheet not found: ' + sheetName);
    }

    // Get all data from the sheet
    const range = sheet.getDataRange();
    const values = range.getValues();
    
    // Check if sheet has data
    if (values.length === 0) {
      throw new Error('No data found in the sheet: ' + sheetName);
    }

    // Get or create the destination folder
    const folderName = "eLivingWillForm/eLivingWill_RegisterPDF";
    let folder = DriveApp.getFoldersByName(folderName).hasNext() ? 
                DriveApp.getFoldersByName(folderName).next() : 
                DriveApp.createFolder(folderName);
    
    // Template document settings
    const templateId = '124PE6QFfPWe-FNBPRyAfVBKXgrh-pX89qqPNGGdy6Ms'; // Replace with your template ID
    
    // Create a copy of the template
    const docId = DriveApp.getFileById(templateId).makeCopy().getId();
    const doc = DocumentApp.openById(docId);
    const body = doc.getBody();

    // Get the most recent form submission (last row)
    const row = values[values.length - 1];
    Logger.log('Processing row data: ' + JSON.stringify(row));

    // Validate row has enough data
    if (row.length < 39) {
      throw new Error('Row data does not have enough columns. Expected 39, got ' + row.length);
    }

/**
 * Helper function to safely replace placeholders in the document
 * @param {DocumentApp.Body} body - The document body
 * @param {string} placeholder - The placeholder text to replace
 * @param {any} value - The value to replace with
 */
function safeReplace(body, placeholder, value) {
  try {
    // Handle blank, null, or undefined values
    let stringValue = "n/a";
    
    if (value != null && value !== undefined) {
      // Convert to string and trim whitespace
      stringValue = String(value).trim();
      // If after trimming it's empty, use "n/a"
      if (stringValue === "") {
        stringValue = "n/a";
      }
    }
    
    Logger.log('Replacing ' + placeholder + ' with: ' + stringValue);
    
    // Use replaceText to preserve formatting
    body.replaceText(placeholder, stringValue);
  } catch(e) {
    Logger.log('Error replacing ' + placeholder + ': ' + e.toString());
    // On error, try to replace with "n/a"
    try {
      body.replaceText(placeholder, "n/a");
    } catch(innerError) {
      Logger.log('Failed to replace with default value: ' + innerError.toString());
    }
  }
}

    // Define placeholders that match the template
    const placeholders = [
      '{{timestamp}}',
      '{{email}}',
      '{{hospital}}',
      '{{hcode}}',
      '{{subdistrict}}',
      '{{district}}',
      '{{province}}',
      '{{name}}',
      '{{tel}}',
      '{{hposition}}',
      '{{hname}}',
      '{{admin_prefix}}',
      '{{admin_name}}',
      '{{admin_pid}}',
      '{{admin_position}}',
      '{{admin_department}}',
      '{{admin_tel}}',
      '{{admin_email}}',
      '{{coor_prefix}}',
      '{{coor_name}}',
      '{{coor_pid}}',
      '{{coor_position}}',
      '{{coor_department}}',
      '{{coor_tel}}',
      '{{coor_email}}',
      '{{u1_prefix}}',
      '{{u1_name}}',
      '{{u1_pid}}',
      '{{u1_position}}',
      '{{u1_department}}',
      '{{u1_tel}}',
      '{{u1_email}}',
      '{{u2_prefix}}',
      '{{u2_name}}',
      '{{u2_pid}}',
      '{{u2_position}}',
      '{{u2_department}}',
      '{{u2_tel}}',
      '{{u2_email}}'
    ];

    // Replace each placeholder with corresponding data
    placeholders.forEach((placeholder, index) => {
      Logger.log(`Replacing ${placeholder} with value: ${row[index]}`);
      safeReplace(body, placeholder, row[index]);
    });

    // Save and close the document before converting to PDF
    doc.saveAndClose();

    // Convert the document to PDF
    const pdfFile = DriveApp.getFileById(docId);
    const pdf = pdfFile.getAs('application/pdf');

    // Get timestamp and hcode from the row data
    const timestamp = row[0];
    const hcode = row[36];

    // Format timestamp
    const formattedTimestamp = Utilities.formatDate(new Date(timestamp), "GMT+7", "yyyyMMdd_HHmmss");

    // Get and increment file number from Properties Service
    const properties = PropertiesService.getScriptProperties();
    let fileNumber = Number(properties.getProperty('fileNumber') || 0) + 1;
    properties.setProperty('fileNumber', fileNumber.toString());

    // Format the file number with leading zeros
    const paddedNumber = fileNumber.toString().padStart(3, '0');

    // Create the filename
    const fileName = `${paddedNumber}_${formattedTimestamp}_${hcode}.pdf`;

    // Create file with the new name in the specified folder
    const file = DriveApp.createFile(pdf).setName(fileName);
    folder.addFile(file);
    DriveApp.getRootFolder().removeFile(file); // Remove from root folder

    // Get recipient email from the form response
    const email = row[1]; // Assuming email is in the second column
    Logger.log('Sending email to: ' + email);

    // Configure email settings
    const subject = 'แบบฟอร์มการขอขึ้นทะเบียนสถานพยาบาลระบบ e-Living Will';
    const message = 'ท่านสามารถดาวน์โหลดไฟล์แบบฟอร์ม PDF ได้จากไฟล์แนบในอีเมลฉบับนี้ หลังจากปรินท์ให้ผู้บริหารลงนามแล้ว ขอความกรุณาส่งหนังสือตอบรับการขึ้นทะเบียนของสถานพยาบาลของท่าน ได้ที่อีเมลสารบรรณกลาง สำนักงานคณะกรรมการสุขภาพแห่งชาติ (สช.) nhco@saraban.mail.go.th และสำเนาอีเมลถึง kanokwan@nationalhealth.or.th';

    // Send email with PDF attachment
    GmailApp.sendEmail(email, subject, message, {
      attachments: [file.getAs(MimeType.PDF)]
    });

    Logger.log('PDF created and sent successfully to: ' + email);

    // Clean up: Delete the temporary Google Doc
    DriveApp.getFileById(docId).setTrashed(true);

  } catch (e) {
    // Log any errors that occur during execution
    Logger.log('Error in createAndSendPDF: ' + e.toString());
    throw e; // Re-throw the error for the Apps Script dashboard
  }
}

/**
 * Optional: Reset the file number counter
 * Run this function if you want to start the numbering from 1 again
 */
function resetFileCounter() {
  const properties = PropertiesService.getScriptProperties();
  properties.setProperty('fileNumber', '0');
  Logger.log('File counter reset to 0');
}
