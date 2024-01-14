/**
 * Represents the management of sheets, including creating, modifying, and setting up editors.
 * It handles operations related to task sheets in a Google Spreadsheet environment.
 */
class TaskSheetManager {
    /**
     * Initializes a new instance of the TaskSheetManager class.
     * Sets up the necessary properties and retrieves existing staff data.
     */
    constructor() {
        this.spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
        this.scriptProperties = PropertiesService.getScriptProperties();
        let indexSheetInfo = JSON.parse(this.scriptProperties.getProperty(SCRIPT_PROPERTY_INDEX_SHEET));
        let indexSheet = this.spreadsheet.getSheetByName(indexSheetInfo.ongoingTaskSheetName);
        this.indexSheetURL = indexSheet ? this.spreadsheet.getUrl() + "#gid=" + indexSheet.getSheetId() : "";
        this.backToIndexPhrase = indexSheetInfo.backToIndexPhrase;
        let existingStaffDataString = this.scriptProperties.getProperty(SCRIPT_PROPERTY_KEY_STAFF);
        this.existingStaffData = JSON.parse(existingStaffDataString || '[]');
    }

    /**
     * Displays any encountered error in a user-friendly format.
     * Useful for debugging and user notifications.
     *
     * @param {Error} e - The error object caught during execution.
     */
    static displayError(e) {
        console.log(`The following error occurred: ${e.message}\nStack Trace:\n${e.stack}`);
        Browser.msgBox(`The following error occurred: ${e.message}\nStack Trace:\n${e.stack}`);
    }

    /**
     * Displays a modal dialog using a specified HTML template.
     * @param {string} htmlFileName - HTML file name for the modal dialog.
     * @param {string} title - Title for the modal dialog.
     * @param {string} action - The action type (e.g., 'create', 'modify').
     * @param {Array} [currentEditors=[]] - Current editors for the sheet.
     * @param {string} [sheetName=""] - Name of the sheet.
     */
    showModal(htmlFileName, title, action, currentEditors = [], sheetName = "") {
        let htmlTemplate = HtmlService.createTemplateFromFile(htmlFileName);
        htmlTemplate.action = action;
        htmlTemplate.existingStaffData = this.existingStaffData;
        console.log(this.existingStaffData);
        htmlTemplate.currentEditors = currentEditors;
        htmlTemplate.sheetName = sheetName;

        let htmlOutput = htmlTemplate.evaluate().setWidth(400).setHeight(600);
        SpreadsheetApp.getUi().showModalDialog(htmlOutput, title);
    }

    /**
     * Processes the checked staff information from the HTML modal and extracts names and emails.
     * 
     * @param {Array} checkedStaff - Array containing the staff information in 'name|email' format.
     * @returns {Object} An object containing arrays of staff names and emails.
     */
    static submitStaff(checkedStaff) {
        console.log(`Submitted values from the html model is ${checkedStaff}`);
        let staffNames = [];
        let staffEmails = [];

        if (!Array.isArray(checkedStaff)) {
          console.error("checkedStaff is not an array:", checkedStaff);
          return;
        }

        for (let staff of checkedStaff) {
          let [name, email] = staff.split('|');
          staffNames.push(name);
          staffEmails.push(email);
        }
        console.log({
          staffName: staffNames,
          staffEmail: staffEmails
        });
        
        return {
          staffNames: staffNames,
          staffEmails: staffEmails
        };
    }

    /**
     * Prompts the user to input a sheet name following a specific format (e.g., 'Category Name: Task Name').
     * Validates the input format and returns the sheet name if it matches the required pattern.
     * 
     * @returns {string|null} The validated sheet name or null if the input is canceled or does not match the format.
     */
    static getSheetNameInput() {
        let sheetName = Browser.inputBox("Enter the sheet name (e.g., Category Name: Task Name) *INCLUDE ' : ' and a space (Half size)");

        // Regular expression to match the structure "Category Name: Task Name"
        let formatRegex = /^[^:]+:[^:]+$/;

        // Check if the sheetName matches the required format
        if (formatRegex.test(sheetName.trim())) {
            return sheetName;
        } else {
            // If format does not match, return the error message
            let invalidMsg = Browser.msgBox("The sheet name is not set as instructed 'Category Name: Task Name'. Try again.",Browser.Buttons.OK_CANCEL);
            if(invalidMsg === 'ok'){
              TaskSheetManager.getSheetNameInput();
              return
            } else {
              Browser.msgBox("Inputting a sheet name was cancelled.");
              return;
            }
        }
    }

    /**
     * Prompts the user to input the number of rows for setting up dropdowns in the sheet.
     * Parses the input as an integer.
     * 
     * @returns {number} The number of rows entered by the user.
     */
    static getRowNumInput() {
        let rowNumInput = Browser.inputBox("Enter the number of rows to set the pulldown for staff name");
        return parseInt(rowNumInput);
    }

    /**
     * Handles the creation of a new sheet.
     *
     * @param {Array.<Object>} checkedStaff - An array of selected staff objects, each containing name and email.
     */
    createNewSheet(checkedStaff) {
        let sheetName = TaskSheetManager.getSheetNameInput();

        // Check if sheetName is empty or does not include ":"
        if (!sheetName || sheetName.indexOf(":") === -1) {
            Browser.msgBox('The data is not input or does not include " : " (Half size). Try again.');
            return;
        }

        // Check for duplicate sheet name
        let sheets = this.spreadsheet.getSheets();
        for (let i = 0; i < sheets.length; i++) {
            if (sheets[i].getName() === sheetName) {
                Browser.msgBox('The sheet name already exists. Please input a different name.');
                return;
            }
        }

        let rowNum = TaskSheetManager.getRowNumInput();
        if (isNaN(rowNum)) {
            Browser.msgBox("The input data is not a number. Please try again.");
            return;
        }

        try {
          let staffData = TaskSheetManager.submitStaff(checkedStaff);
          console.log(`stafftData is ${staffData.staffNames}`);

          // Check if designated emails are current editors
          let currentEditors = this.spreadsheet.getEditors().map(editor => editor.getEmail());
          
          // Find emails that are not current editors
          let nonEditorEmails = staffData.staffEmails.filter(email => !currentEditors.includes(email));
          
          if (nonEditorEmails.length > 0) {
              let nonEditorEmailsStr = nonEditorEmails.join(', ');
              Browser.msgBox(`The following email(s) are not current editors of the spreadsheet: ${nonEditorEmailsStr}. Operation cancelled.`);
              return;
          }

          let options = staffData.staffNames; // This will be used for dropdown

          let newSheet = this.spreadsheet.insertSheet(sheetName,2);
          
          // Setup headers for columns B to F
          let headerRange = newSheet.getRange("B1:F1");
          headerRange.setValues([['Item', 'Summary', 'Date', 'Staff', 'Complete']])
                      .setBackground("#D3D3D3")
                      .setFontWeight("bold")
                      .setHorizontalAlignment("center")

          // Set hyperlink, background, fontWeight, and horizontalAlignment for "Back to Index" in A1
          newSheet.getRange("A1").setFormula(`=HYPERLINK("${this.indexSheetURL}", "${this.backToIndexPhrase}")`)
                                .setBackground("#FFCCCC")
                                .setFontWeight("bold")
                                .setHorizontalAlignment("center")

          // Set data format as date for "Date" column
          let dateRange = newSheet.getRange(`D2:D`);
          dateRange.setNumberFormat('yy/M/d (ddd)');

          // Set data validation for "Date" column to allow only dates
          let dateValidationRule = SpreadsheetApp.newDataValidation()
            .requireDate()
            .setAllowInvalid(false)
            .build();
          dateRange.setDataValidation(dateValidationRule);

          // Set dropdown for "Staff" column
          let dropdownRule = SpreadsheetApp.newDataValidation().requireValueInList(options, true).build();
          newSheet.getRange(`E2:E${rowNum}`).setDataValidation(dropdownRule);

          // Insert checkboxes for column F from F2 onwards
          newSheet.getRange(`F2:F${rowNum}`).insertCheckboxes();

          newSheet.getRange("A:F").setVerticalAlignment("middle");
          newSheet.getRange(`B1:F${rowNum}`).setBorder(true,true,true,true,true,true).setFontSize(11);

          newSheet.setColumnWidth(1, 120)
                  .setColumnWidth(2, 300)
                  .setColumnWidth(3, 600);

          newSheet.getRange(1, 1, newSheet.getMaxRows(), newSheet.getMaxColumns()).createFilter();

          newSheet.getRange("B:C").setWrap(true);
          newSheet.getRange("B:B").setHorizontalAlignment("center");
          newSheet.getRange("D:F").setHorizontalAlignment("center");
          newSheet.setFrozenRows(1);
          
          // Set protection for the new sheet to allow editing only by selected staff emails
          let protection = newSheet.protect().setDescription('Sheet protection');
          protection.removeEditors(protection.getEditors());
          protection.addEditors(staffData.staffEmails);

          Browser.msgBox(`New sheet ${sheetName} was created successfully.`);
      } catch(e) {
          this.displayError(e);
          return;
      }
    }

    /**
     * Modifies the editors for a specified sheet.
     *
     * @param {Array.<Object>} checkedStaff - An array of selected staff objects, each containing name and email.
     * @param {string} sheetName - The name of the sheet to modify.
     */
    modifyEditors(checkedStaff, sheetName) {
      try {
        let staffData = TaskSheetManager.submitStaff(checkedStaff);
        console.log(`staffData is ${staffData.staffNames}`);

        let options = staffData.staffNames; // This will be used for dropdown

        let sheet = this.spreadsheet.getSheetByName(sheetName);

        // Fetch all data validations in column E at once
        let lastRow = sheet.getLastRow();
        let dataValidations = sheet.getRange(`E1:E${lastRow}`).getDataValidations();

        // Find the last row with data validation
        let lastRowWithDataValidation = dataValidations.findIndex(cellValidation => cellValidation == null);

        // If we didn't find a null validation, it means all rows have data validation
        if (lastRowWithDataValidation === -1) {
          lastRowWithDataValidation = lastRow;
        }

        // Set dropdown for "Staff" column based on previously found number of rows
        let dropdownRule = SpreadsheetApp.newDataValidation().requireValueInList(options, true).build();
        sheet.getRange(`E2:E${lastRowWithDataValidation}`).setDataValidation(dropdownRule);

        // Set protection for the new sheet to allow editing only by selected staff emails
        let protection = sheet.protect().setDescription('Sheet protection');
        protection.removeEditors(protection.getEditors());
        protection.addEditors(staffData.staffEmails);

        Browser.msgBox(`Editors for ${sheetName} were modified successfully.`);
      } catch (e) {
        this.displayError(e);
        return;
      }
    }

    /**
     * Updates task completion status in the spreadsheet based on the content of Google Documents.
     */
    updateCompletionStatusToSheet() {
        // Retrieve the URLs from script properties
        let generalReminderUrls = JSON.parse(this.scriptProperties.getProperty(SCRIPT_PROPERTY_KEY_GENERAL_REM_DOC_URL) || '{}');
        let staffBasedReminderData = JSON.parse(this.scriptProperties.getProperty(SCRIPT_PROPERTY_KEY_STAFFBASED_REM_DATA) || '[]');

        // Process General Reminder URL if set
        if (generalReminderUrls.generalTodayReminderDocUrl) {
            this.processDocument(generalReminderUrls.generalTodayReminderDocUrl);
        }

        // Process Staff-Based Reminder URLs if set
        staffBasedReminderData.forEach(staffObj => {
            let staffName = Object.keys(staffObj)[0];
            let staffInfo = staffObj[staffName];
            if (staffInfo.todayReminderUrl) {
                this.processDocument(staffInfo.todayReminderUrl);
            }
        });
    }

    /**
     * Processes a Google Document to update the completion status in the corresponding Google Sheet.
     * @param {string} docUrl - The URL of the Google Document to be processed.
     */
    processDocument(docUrl) {
        let docId = ReminderManager.extractDocIdFromUrl(docUrl);
        let doc = DocumentApp.openById(docId);
        let body = doc.getBody();
        let sheetName;
        let numElements = body.getNumChildren();

        for (let i = 0; i < numElements; i++) {
            let element = body.getChild(i);
            if (element.getType() === DocumentApp.ElementType.PARAGRAPH) {
                let paragraph = element.asParagraph();
                if (paragraph.getHeading() === DocumentApp.ParagraphHeading.HEADING1) {
                    sheetName = paragraph.getText();
                    continue;
                }
            }

            if (element.getType() === DocumentApp.ElementType.TABLE && sheetName) {
                let table = element.asTable();
                this.updateSheetWithTableData(table, sheetName);
            }
        }
    }

    /**
     * Updates the completion status in a Google Sheet based on the data from a Google Document table.
     * @param {Table} table - The table element from the Google Document.
     * @param {string} sheetName - The name of the Google Sheet to update.
     */
    updateSheetWithTableData(table, sheetName) {
        let targetSheet = this.spreadsheet.getSheetByName(sheetName);
        if (!targetSheet) return;

        for (let j = 1; j < table.getNumRows(); j++) {
            let completionCell = table.getCell(j, 4);
            if (completionCell.getText() === "C") {
                let taskFlagInDoc = table.getCell(j, 0).getText() + table.getCell(j, 1).getText();
                console.log(`${taskFlagInDoc} is completed.`);
                let columnB = targetSheet.getRange("B1:B" + targetSheet.getLastRow()).getValues();
                let columnC = targetSheet.getRange("C1:C" + targetSheet.getLastRow()).getValues();

                for (let k = 0; k < columnB.length; k++) {
                  let taskFlagInSheet = columnB[k][0] + columnC[k][0]
                  if (taskFlagInDoc === taskFlagInSheet) {
                      targetSheet.getRange(k + 1, 6).setValue(true);
                      console.log(`Status for ${taskFlagInDoc} has been changed from incomplete to completed.`);
                      break;
                  }
                }
            }
        }
    }

    /**
     * Deletes displayed reminders in the Google Document. Currently not set as a global function
     */
    deleteRemindersInDoc(){
      let docId;
      let doc;
      if(this.period === 'today' && this.target === 'general'){
          docId = ReminderManager.extractDocIdFromUrl(TODAY_REM_DOC_URL);
          doc = DocumentApp.openById(docId);
          doc.getBody().clear();
      } else if (this.period === 'today' && this.target === 'staffBased'){
          return;
      }
    }
}

/**
 * Opens a modal dialog for creating a new task sheet.
 * Utilizes the TaskSheetManager to handle the creation process.
 */
function createNewSheetModal() {
    let manager = new TaskSheetManager();
    manager.showModal('show-editor-choice', 'Choose relevant staff as editors', 'create');
}

/**
 * Wrapper function for creating a new sheet.
 * It receives checked staff information and passes it to the TaskSheetManager.
 * @param {Array} checkedStaff - Array of selected staff members.
 */
// Create a new sheet using the TaskSheetManager class
function createNewSheetWrapper(checkedStaff) {
    let manager = new TaskSheetManager();
    manager.createNewSheet(checkedStaff);
}

/**
 * Opens a modal dialog for modifying sheet editors.
 * This function is used when changing permissions for a specific sheet.
 */
function modifyEditorsModal() {
    let manager = new TaskSheetManager();
    let sheet = manager.spreadsheet.getActiveSheet();
    let sheetName = sheet.getName();
    let protection = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET)[0]; // Assuming there's at least one protection
    let currentEditors = protection.getEditors().map(editor => editor.getEmail());
    manager.showModal('show-editor-choice', `Modify Editors of ${sheetName}`, 'modify', currentEditors, sheetName);
}

/**
 * Wrapper function for modifying editors of a sheet.
 * @param {Array} checkedStaff - Array of selected staff members.
 * @param {string} sheetName - The name of the sheet to modify.
 */
function modifyEditorsWrapper(checkedStaff, sheetName) {
    let manager = new TaskSheetManager();
    manager.modifyEditors(checkedStaff, sheetName);
}

/**
 * Updates the completion status of tasks in the spreadsheet.
 * This function retrieves data from reminder documents and updates the task completion status.
 */
function updateCompletionStatusToSheet(){
  const manager = new TaskSheetManager();
  manager.updateCompletionStatusToSheet();
}

/**
 * Displays an error message related to task sheet execution.
 * @param {string} action - The action attempted when the error occurred.
 */
function displayExecutionError(action){
  Browser.msgBox(`Action was ${action}. But neither create a new sheet nor modify editors of the sheeet has been executed. Check the codes and try again.`);
  return;
}