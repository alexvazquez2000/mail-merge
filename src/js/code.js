
/* on a google spreadsheed, go to extensions, then add this code */

//your template ID of the Word Doc -
const TEMPLATE_ID = '1mSF1x0R-ysMKCEkbK2fzGaHVe83xshUE457yIDOQZoU';
//Attendance percentage -  100% is perfect attendance
const INCLUDE_LESS_THAN = 80

/* This is the entry point. Make sure google is setup to run this function */
function myFunction() {
  let data = readSpreadsheetData();
  //Validate the data
  //for (let i = 0; i < data.length; i++) {
  //  var row = data[i];
  //  Logger.log("--Name: " + row.name + ", Grade: " + row.grade + " Per:" + row.percent + " days:" + row.days);
  //}
  createDocumentFromTemplate(data)
}

/**
 * Reads all data from the active sheet and logs it to the execution log.
 */
function readSpreadsheetData() {
  // Get the active spreadsheet and the active sheet
  let sheet = SpreadsheetApp.getActiveSheet();
  
  // Get the range containing all data
  // getDataRange() automatically finds the last row and column with content
  let range = sheet.getDataRange();
  
  // Get the values as a 2D array (e.g., [[header1, header2], [data1, data2], ...])
  let values = range.getValues();
  
  // Log the values to the execution log (View -> Execution log)
  // Logger.log(values);
  let data = [];
  // Example: Iterate through the data and access individual cells
  // The first row (index 0) is usually the header - in our case the header is in about row 8 or 9
  for (let i = 1; i < values.length; i++) { // Start from index 1 to skip headers
    let row = values[i];
    let name = row[0].trim(); // Assuming ID is in the first column (index 0)
    let grade = row[1]; // Assuming Event is in the second column (index 1)
    let percent = Number(row[28]);
    let days = row[29];
    if (name.length >0 && name.length < 50 &&  typeof percent === 'number'  && Number.isFinite(percent) ) {
      if (percent < INCLUDE_LESS_THAN) {
        Logger.log("Adding Name: " + name + ", Grade: " + grade + " Per:" + percent + " days:" + days);
        data.push( {name, grade, percent, days });
      }
    }
  }
  return data;
}

function createDocumentFromTemplate(data) {

  const templateDoc = DocumentApp.openById(TEMPLATE_ID);
  const templateBody = templateDoc.getBody();
  const numElements = templateBody.getNumChildren();

  const templateFile = DriveApp.getFileById(TEMPLATE_ID);
  let copy = templateFile.makeCopy(`mergedLetter`); // Create copy in the same folder as the template

  const copyId = copy.getId();
  const copyDoc = DocumentApp.openById(copyId);
  const copyBody = copyDoc.getBody();

  // Replace placeholders in the copied document
  for (let i = 0; i < data.length; i++) {
    let row = data[i];
    Logger.log((i+1) +"/" + data.length + " Creating letter for --Name: " + row.name + ", Grade: " + row.grade + " Per:" + row.percent + " days:" + row.days);

    let studentCopy = templateBody.copy();
    studentCopy.replaceText('<<Student Name>>', row.name);
    studentCopy.replaceText('<<Grade>>', row.grade);
    studentCopy.replaceText('<<Attendance>>', row.percent);
    studentCopy.replaceText('<<Attendance  Days Absent>>', row.days);
    Logger.log("fixed text");
    for (let i = 0; i < numElements; i++) {
      const element = studentCopy.getChild(i);
      const elementType = element.getType();
      
      // Copy elements by type, as you can only copy elements within the same document type
      if (elementType === DocumentApp.ElementType.PARAGRAPH) {
        // Use copy() to create a new instance that can be moved to a different doc
        copyBody.appendParagraph(element.copy());
      } else if (elementType === DocumentApp.ElementType.TABLE) {
        copyBody.appendTable(element.copy());
      } else if (elementType === DocumentApp.ElementType.LIST_ITEM) {
        copyBody.appendListItem(element.copy());
      } else {
        Logger.log("Missing " + elementType)
      }
      // Add more else if blocks for other element types (e.g., INLINE_IMAGE, HORIZONTAL_RULE) as needed
    }
    Logger.log("done with row");
  }


  //Save and close the document
  copyDoc.saveAndClose();

  Logger.log(`Created new document: ${copyDoc.getUrl()}`);
  // Returns the URL of the new document
  return copyDoc.getUrl();
}
