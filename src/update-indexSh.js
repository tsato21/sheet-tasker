/**
 * Updates index sheets for ongoing and completed tasks.
 * This function organizes tasks into categories and updates the corresponding index sheets.
 */
function updateAllTaskIndexSheets() {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let indexSheetInfo = JSON.parse(PropertiesService.getScriptProperties().getProperty(SCRIPT_PROPERTY_INDEX_SHEET));
  let ongoingTaskIndexSh = ss.getSheetByName(indexSheetInfo.ongoingTaskSheetName);
  let completedTaskIndexSh = ss.getSheetByName(indexSheetInfo.completedTaskSheetName);
  let completionFlag = indexSheetInfo.completionFlag;
  try {
    let allSheets = ss.getSheets();
    let ongoingTasks = {};
    let completedTasks = {};

    // Fetching and organizing tasks
    allSheets.forEach(sheet => {
      let sheetName = sheet.getName();
      sortTaskSheetByDate(sheet, sheetName);
      if (sheetName.includes(":") && !sheet.isSheetHidden()) {
        let category, task;
        let sheetGID = sheet.getSheetId();
        let sheetURL = `${ss.getUrl()}#gid=${sheetGID}`;
        let taskInfo = { url: sheetURL };

        if (sheetName.includes(completionFlag)) {
          [category, task] = sheetName.replace(completionFlag,"").split(":").map(part => part.trim());
          completedTasks[category] = completedTasks[category] || [];
          taskInfo.task = task;
          completedTasks[category].push(taskInfo);
        } else {
          [category, task] = sheetName.split(":").map(part => part.trim());
          ongoingTasks[category] = ongoingTasks[category] || [];
          taskInfo.task = task;
          ongoingTasks[category].push(taskInfo);
        }
      }
    });

    // Update the ongoing task index sheet
    updateSheetWithTaskData_(ongoingTaskIndexSh, ongoingTasks, "#FF8C00");

    // Update the completed task index sheet
    updateSheetWithTaskData_(completedTaskIndexSh, completedTasks, "#696969");

  } catch (error) {
    Logger.log("Error updating task index sheets: " + error.message);
    Logger.log("Stack Trace: " + error.stack);
  }
}


/**
 * Updates a specified sheet with task data, formatting, and hyperlinks.
 * It adjusts the number of columns as needed and applies formatting to display tasks categorically.
 * 
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheetToUpdate - The sheet to be updated with task data.
 * @param {Object} categoryData - An object containing task data categorized.
 * @param {string} tabColor - The color code for the sheet's tab.
 */
function updateSheetWithTaskData_(sheetToUpdate, categoryData, tabColor) {
  sheetToUpdate.clear();
  let lastColNum = sheetToUpdate.getMaxColumns();
  let needColNum = Object.keys(categoryData).length;

  // Log current status
  // console.log(`${sheetToUpdate.getSheetName()}: lastColNum is ${lastColNum}, needColNum is ${needColNum}`);

  // Check if additional columns are needed
  if (lastColNum < needColNum){
    // Insert enough columns to meet the requirement
    let columnsToInsert = needColNum - lastColNum;
    sheetToUpdate.insertColumnsAfter(lastColNum, columnsToInsert);
    // console.log(`${columnsToInsert} columns were inserted.`);
  }

  let currentCol = 1;
  for (let category in categoryData) {
      let updates = [];
      updates.push([category]);

      let hyperlinkUpdates = [['']];  // Top cell is the category, no hyperlink for it.
      
      for (let taskInfo of categoryData[category]) {
        updates.push([taskInfo.task]);
        hyperlinkUpdates.push(['=HYPERLINK("' + taskInfo.url + '","' + taskInfo.task + '")']);
      }

      // Apply the updates in batches
      sheetToUpdate.getRange(1, currentCol, updates.length).setValues(updates);
      sheetToUpdate.getRange(2, currentCol, hyperlinkUpdates.length - 1).setFormulas(hyperlinkUpdates.slice(1)).setFontSize(11);

      // Formatting 
      sheetToUpdate.getRange(1, currentCol).setBackground("#D3D3D3")
                                          .setFontSize(16)
                                          .setFontWeight("bold")
                                          .setWrap(true);

      let lastRow = updates.length;
      sheetToUpdate.getRange(1, currentCol, lastRow).setBorder(true, true, true, true, true, true)
                                                    .setWrap(true)
                                                    .setVerticalAlignment("middle")
                                                    .setHorizontalAlignment("center");
      
      sheetToUpdate.setColumnWidth(currentCol, 150);
      currentCol += 1;
    }

    // Setting the tab color of the target sheet
    sheetToUpdate.setTabColor(tabColor);
  
}

/**
 * Sorts a given task sheet by date.
 * Only sorts sheets that are not index sheets and have more than one row of data.
 * 
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The task sheet to be sorted.
 * @param {string} sheetName - The name of the sheet.
 */
function sortTaskSheetByDate(sheet, sheetName) {
  let lastRow = sheet.getLastRow();
  let lastCol = sheet.getLastColumn();
  let indexSheetInfo = JSON.parse(PropertiesService.getScriptProperties().getProperty(SCRIPT_PROPERTY_INDEX_SHEET));

  if (sheetName !== indexSheetInfo.ongoingTaskSheetName && sheetName !== indexSheetInfo.completedTaskSheetName && lastRow > 1) {
    // console.log(`sheetName is ${sheetName}`);
    let range = sheet.getRange(2, 1, lastRow - 1, lastCol);
    range.sort({ column: 4, ascending: true });
  }
}
