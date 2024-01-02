/**
 * Represents a reminder with task details.
 * Includes information about the task item, notes, due date, and assigned staff member.
 */
class Reminder {
    constructor(item, note, date, staff) {
        this.item = item;
        this.note = note;
        this.date = date;
        this.staff = staff;
    }
}

/**
 * Represents a reminder for a specific sheet.
 * Contains the name and URL of the sheet, along with an array of Reminder objects related to tasks.
 */
class SheetReminder {
    constructor(sheetName, sheetURL, taskData) {
        this.sheetName = sheetName;
        this.sheetURL = sheetURL;
        this.taskData = taskData;  // Array of Reminder objects
    }
}

/**
 * Manages reminders, including gathering, formatting, and sending.
 * Supports reminders for general tasks or staff-based tasks, with a focus on either today or a future period (e.g., next week).
 */
class ReminderManager {
    /**
     * @param {string} target - The target audience for the reminder ('general' or 'staffBased').
     * @param {string} period - The period for the reminder ('today' or 'week').
     */
    constructor(target, period) {
        this.target = target;
        this.period = period;
        this.reminderData = [];
        this.ss = SpreadsheetApp.getActiveSpreadsheet();
        this.scriptProperties = PropertiesService.getScriptProperties();
    }

    /*
    Use either Japanese version or English version to display date in Reminder Docs and Gmail Subject
    */
    /**
     * Formats a given Date object into a Japanese date string.
     *
     * @param {Date} dateObj - The Date object to format.
     * @returns {string} - The formatted date string in Japanese format.
     *
     * @example
     * let date = new Date(2023, 4, 5); // 5th May 2023
     * console.log(ReminderManager.formatJapaneseDate(date)); // Outputs: "5月5日(金)"
     */
    static formatJapaneseDate(dateObj) {
        let months = ["1月", "2月", "3月", "4月", "5月", "6月", "7月", "8月", "9月", "10月", "11月", "12月"];
        let days = ["日", "月", "火", "水", "木", "金", "土"];
      
        let month = months[dateObj.getMonth()];
        let day = days[dateObj.getDay()];
        let date = dateObj.getDate();
      
        return `${month}${date}日(${day})`;
    }

    /**
     * Formats a given Date object into an English date string.
     *
     * @param {Date} dateObj - The Date object to format.
     * @returns {string} - The formatted date string in English format.
     *
     * @example
     * let date = new Date(2023, 4, 5); // 5th May 2023
     * console.log(ReminderManager.formatEnglishDate(date)); // Outputs: "Friday, May 5, 2023"
     */
    static formatEnglishDate(dateObj) {
        let months = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
        let days = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];

        let month = months[dateObj.getMonth()];
        let day = days[dateObj.getDay()];
        let date = dateObj.getDate();
        let year = dateObj.getFullYear();

        return `${day}, ${month} ${date}, ${year}`;
    }


    /**
     * Extracts the Google Document ID from a given URL.
     *
     * @param {string} url - The Google Document URL.
     * @returns {string|null} - The extracted document ID or null if not found.
     *
     * @example
     * let url = 'https://docs.google.com/document/d/XXXXXXXX/edit';
     * let docId = ReminderManager.extractDocIdFromUrl(url);
     * console.log(docId);  // Outputs: 1AFA8L0sKMcMVgagXMLqbInaNk06pZ4SZnjkqT-hi2_8
     */
    static extractDocIdFromUrl(url) {
        let regex = /\/d\/(.*?)\//;
        let match = url.match(regex);
        if (match && match[1]) {
            return match[1];
        } else {
            return null;  // or throw an error if you prefer
        }
    }

    /**
     * Gathers reminder data from the spreadsheet.
     * Processes each sheet to extract tasks and organize them into reminders, considering the set period (today or week) and type (general or staff-based).
     * Handles timeouts by saving progress and scheduling a follow-up execution.
     * 
     * @returns {Array<SheetReminder>} An array of SheetReminder objects containing the reminder data.
     * The structure of returns:
     *  [
     *    {
     *      sheetName: "Project A Tasks",
     *      sheetUrl: "https://docs.google.com/spreadsheets/d/12345/edit#gid=67890",
     *      taskData: [
     *        {
     *          item: "Complete budget report",
     *          note: "Include projections for next quarter",
     *          dateInfo: "2023-07-21",
     *          staff: "John Doe"
     *        },
     *        {
     *          item: "Update project timeline",
     *          note: "Reflect changes in deliverable dates",
     *          dateInfo: "2023-07-22",
     *          staff: "Jane Smith"
     *        }
     *      ]
     *    },
     *    {
     *      sheetName: "Project B Tasks",
     *      sheetUrl: "https://docs.google.com/spreadsheets/d/54321/edit#gid=09876",
     *      taskData: [
     *        {
     *          item: "Review codebase for errors",
     *          note: "Focus on the authentication module",
     *          dateInfo: "2023-07-23",
     *          staff: "John Doe"
     *        },
     *      ]
     *    }
     *  ]
     */
    getReminderData() {
        console.log('Starting getReminderData...');
        let startTime = new Date().getTime();
        let sheets = this.ss.getSheets();

        // Use unique keys for stored reminders and current sheet index
        let storedDataKey = `SCRIPT_PROPERTY_KEY_STORED_REMINDERS_${this.target.toUpperCase()}_${this.period.toUpperCase()}`;
        let currentSheetIndexKey = `SCRIPT_PROPERTY_KEY_CURRENT_SHEET_INDEX_${this.target.toUpperCase()}_${this.period.toUpperCase()}`;
        let storedDataStr = this.scriptProperties.getProperty(storedDataKey);
        let completionStatusKey = this.getCompletionStatusKey();
        this.scriptProperties.setProperty(completionStatusKey,'NOT YET');


        if (storedDataStr) {
            try {
                this.reminderData = JSON.parse(storedDataStr);
                console.log(`There are some reminderData already stored in the previous execution.`);
            } catch (e) {
                console.error('Error parsing stored reminders. Defaulting to an empty array.', e);
            }
        }

        let currentSheetIndex = parseInt(this.scriptProperties.getProperty(currentSheetIndexKey) || '0');
        let today = new Date();
        today.setHours(0, 0, 0, 0);

        for (let i = currentSheetIndex; i < sheets.length; i++) {
          let sheet = sheets[i];
          let sheetName = sheet.getName();
          console.log(`Start reading ${sheetName}`);

          // Simulate a long-running process for testing purposes by sleeping for 10 seconds
          // Utilities.sleep(10000);

          // Check elapsed time, if more than 5 minutes (3000 seconds for a buffer), save progress and exit
          // On timeout, trigger sendGeneralReminder
          // if ((new Date().getTime() - startTime) > 300000) {
          if ((new Date().getTime() - startTime) > 700) {
           // Save progress with the unique key
            this.scriptProperties.setProperty(storedDataKey, JSON.stringify(this.reminderData));
            this.scriptProperties.setProperty(currentSheetIndexKey, i.toString());
            this.createOneTimeTrigger();

            console.log(`Timeout detected, saving progress and setting a trigger for continuation.`);

            return;
            
          } else {
            console.log(`Execusion continues when reading the data in ${sheetName}`);
          }

          if (sheetName !== ONGOING_TASKS_INDEX_SHEET_NAME && sheetName !== COMPLETED_TASKS_INDEX_SHEET_NAME) {

            let lastRow = sheet.getRange("B" + sheet.getMaxRows()).getNextDataCell(SpreadsheetApp.Direction.UP).getRow();
            let lastCol = sheet.getLastColumn();
            // console.log(`The number of the data in ${sheetName} is ${lastRow}`);

            if (lastRow ===0  || lastRow === 1 && lastCol === 0){
              continue;
            }
            // Retrieve data up to the last filled cell in column B
            let data = sheet.getRange(1, 1, lastRow, lastCol).getValues();
            let taskData = [];

            let validDates = [new Date(today)]; // Today is always a valid date
            let nextDate = new Date(today.getTime()); // Start with the base date

            for (let i = 1; i <= 5; i++) {
                nextDate.setDate(nextDate.getDate() + 1); // Increment by one day initially

                while (nextDate.getDay() === 0 || nextDate.getDay() === 6) {
                    nextDate.setDate(nextDate.getDate() + 1); // Skip weekends
                }

                validDates.push(new Date(nextDate.getTime())); // Add the valid date to the array
            }
          
            for (let i = 1; i < data.length; i++) {
              let checkbox = data[i][5];
              let staff = data[i][4];
              let dateStr = data[i][3];
              // console.log(`dateStr is ${dateStr}`);
              // If dateStr is empty, skip the current iteration
              if (!dateStr) {
                  // console.log(`Date is not input for this event, which is not subject to the reminder.`);
                  continue;
              }
              let dateInfo = ReminderManager.formatEnglishDate(dateStr);
              let item = data[i][1];
              let note = data[i][2];
              
              let date = new Date(dateStr);
              
              let isFutureValidDate = validDates.some(function(validDate) {
                return validDate.getTime() === date.getTime();
              });
              
              if (this.period === 'today'){
                if (!checkbox && date <= today) {
                  taskData.push(new Reminder(item, note, dateInfo, staff));
                  // console.log(`remiderRecords for ${sheetName} are ${taskData}`)
                }
              } else if (this.period === 'week'){
                if (!checkbox && (date <= today || isFutureValidDate)) {
                  taskData.push(new Reminder(item, note, dateInfo, staff));
                }
              }
            }

            if (taskData.length > 0) {
              let sheetGID = sheet.getSheetId();
              let spreadsheetURL = this.ss.getUrl();
              let sheetURL = spreadsheetURL + "#gid=" + sheetGID;
              this.reminderData.push(new SheetReminder(sheet.getName(), sheetURL, taskData));
            }
          }
        }        
        // Clear the stored reminders and current sheet index after successful processing
        this.scriptProperties.deleteProperty(storedDataKey);
        this.scriptProperties.deleteProperty(currentSheetIndexKey);
        this.deleteOneTimeTriggers();
        this.scriptProperties.setProperty(completionStatusKey,'COMPLETED');


        console.log('Completed processing all sheets in getReminderData.');
        return this.reminderData;
    }

    /**
     * Shares reminders through a Google Document.
     * Gathers reminder data, creates or updates a Google Document with reminder information, and sends an email notification with the document link.
     */
    shareRemindersByDoc() {
        try {
          let reminderData = this.getReminderData();
          let completionStatusKey = this.getCompletionStatusKey();
          // Fetch the completion status using the constructed key
          let completionStatus = this.scriptProperties.getProperty(completionStatusKey);

          if (completionStatus === 'NOT YET') {
              console.log(`getReminderData is still in the middle of processing. Thus, shareRemindersByDoc is not continued.`);
              return;
          } else if (completionStatus === 'COMPLETED') {
              if (!Array.isArray(reminderData) || reminderData.length === 0) {
                  console.log(`There is no reminder data to display or reminderData is not an array.`);
                  return;
              }
              console.log('shareRemindersByDoc has started being executed to creating a reminder doc.');
              let docId, title, body, displayDocUrl, successOrFailure;

              let generalReminderEmails = JSON.parse(this.scriptProperties.getProperty(SCRIPT_PROPERTY_KEY_GENERAL_REMINDER_EMAILS));
              let generalReminderDocsUrls = JSON.parse(this.scriptProperties.getProperty(SCRIPT_PROPERTY_KEY_GENERAL_REM_DOC_URL));
              let staffBasedReminderData = JSON.parse(this.scriptProperties.getProperty(SCRIPT_PROPERTY_KEY_STAFFBASED_REM_DATA));
              /*
              If staffBasedReminderData is not set, it is null.
              If staffBasedReminderData is set, it is as follows:
              [
                {
                  "AA": {
                    "email": "aa@demo.co.jp",
                    "todayReminderUrl": "xxx",
                    "nextWeekReminderUrl": null
                  }
                },
                {
                  "BB": {
                    "email": "bb@demo.co.jp",
                    "todayReminderUrl": "xxx",
                    "nextWeekReminderUrl": null
                  }
                }
              ];
              */

              // Send an email based on period and target
              if (this.target === 'general'){
                if(generalReminderEmails !== null && generalReminderDocsUrls !== null) {
                  let generalTodayReminderDocUrl = generalReminderDocsUrls.generalTodayReminderDocUrl;
                  let generalWeekReminderDocUrl = generalReminderDocsUrls.generalWeekReminderDocUrl;

                    if(this.period === 'today'){
                      title = `Today's General Reminder on ${ReminderManager.formatEnglishDate(new Date())}`;
                      if(generalTodayReminderDocUrl !== null){
                        docId = ReminderManager.extractDocIdFromUrl(generalTodayReminderDocUrl);
                        body = this.presetInDoc(docId, title);
                        this.createReminderTablesInDoc(body, reminderData);
                        displayDocUrl = generalTodayReminderDocUrl;
                        successOrFailure = "success";
                        this.sendEmail(generalReminderEmails,title,successOrFailure,displayDocUrl);
                        console.log(`Today's general reminders were successfully shared by email.`);
                        return true;
                      } else {
                        successOrFailure = "failure";
                        this.sendEmail(generalReminderEmails,title,successOrFailure);
                        console.log(`Today's general reminders could not be shared since the Google Doc is not set, which was informed by email.`);
                        return true;
                      }
                    } else if(this.period === 'week') {
                      title = `Next Week's General Reminder on ${ReminderManager.formatEnglishDate(new Date())}`;
                      if(generalWeekReminderDocUrl !== null){
                        docId = ReminderManager.extractDocIdFromUrl(generalWeekReminderDocUrl);
                        body = this.presetInDoc(docId, title);
                        this.createReminderTablesInDoc(body, reminderData);
                        displayDocUrl = generalWeekReminderDocUrl;
                        successOrFailure = "success";
                        this.sendEmail(generalReminderEmails,title,successOrFailure,displayDocUrl);
                        console.log(`Next week's general reminders were successfully shared by email.`);
                        return true;
                      } else {
                        successOrFailure = "failure";
                        this.sendEmail(generalReminderEmails,title,successOrFailure);
                        console.log(`Next week's general reminders could not be shared since the Google Doc is not set, which was informed by email.`);
                        return true;
                      }
                    }
                } else {
                    let email = Session.getActiveUser().getEmail();
                    let body = `Necessary information such as emails and Google Doc URLs is not set in the setting. Go to "Setting" from Custom Menu and conduct necessary setting.`;
                    let subject = "Error on Sharing General Reminders (Today or Next Week)";
                    console.log(`Error on Sharing General Reminders (Today or Next Week).`);
                    GmailApp.sendEmail(email,subject,body);
                    return true;
                }
              }

              if(this.target === 'staffBased'){
                if(staffBasedReminderData !== null) {
                  staffBasedReminderData.forEach(staffObject => {
                    let staffName = Object.keys(staffObject)[0];
                    let staffInfo = staffObject[staffName];
                    let email = staffInfo.email;
                    let staffSpecificReminders;

                    if(this.period === 'today'){
                      title = `Today's Reminder for ${staffName} on ${ReminderManager.formatEnglishDate(new Date())}`;
                      if(staffInfo.todayReminderUrl){
                        docId = ReminderManager.extractDocIdFromUrl(staffInfo.todayReminderUrl);
                        body = this.presetInDoc(docId, title);
                        staffSpecificReminders = this.filterRemindersForStaff(reminderData, staffName);
                        this.createReminderTablesInDoc(body, staffSpecificReminders);
                        displayDocUrl = staffInfo.todayReminderUrl;
                        successOrFailure = "success";
                        this.sendEmail(email,title,successOrFailure,displayDocUrl);
                        console.log(`Today's reminders were successfully shared with ${staffName} by email.`);
                        return true;
                      } else {
                        successOrFailure = "failure";
                        this.sendEmail(email,title,successOrFailure);
                        console.log(`Today's reminders could not be shared with ${staffName} since the Google Doc is not set, which was informed by email.`);
                        return true;
                      }
                    } else if (this.period === 'week'){
                      title = `Next Week's Reminder for ${staffName} on ${ReminderManager.formatEnglishDate(new Date())}`;
                      if(staffInfo.nextWeekReminderUrl){
                        docId = ReminderManager.extractDocIdFromUrl(staffInfo.nextWeekReminderUrl);
                        body = this.presetInDoc(docId, title);
                        staffSpecificReminders = this.filterRemindersForStaff(reminderData, staffName);
                        this.createReminderTablesInDoc(body, staffSpecificReminders);
                        displayDocUrl = staffInfo.nextWeekReminderUrl;
                        successOrFailure = "success";
                        this.sendEmail(email,title,successOrFailure,displayDocUrl);
                        console.log(`Next week's reminders were successfully shared with ${staffName} by email.`);
                      } else {
                        successOrFailure = "failure";
                        this.sendEmail(email,title,successOrFailure);
                        console.log(`Next week's reminders could not be shared with ${staffName} since the Google Doc is not set, which was informed by email.`);
                        return true;
                      }
                    }
                  });
                } else {
                    let email = Session.getActiveUser().getEmail();
                    let body = `Necessary information such as emails and Google Doc URLs is not set in the setting. Go to "Setting" from Custom Menu and conduct necessary setting.`;
                    let subject = "Error on Sharing General Reminders (Today or Next Week)";
                    console.log(`Error on Sharing General Reminders (Today or Next Week).`);
                    GmailApp.sendEmail(email,subject,body);
                    return true;
                }
              }
            this.scriptProperties.deleteProperty(completionStatusKey);
          }
        } catch (e) {
            console.error(`Error in displayRemindersInDoc: ${e.toString()} at ${e.stack}`);
        }
    }

    /**
     * Prepares a Google Document for displaying reminders.
     * Clears the existing content and sets a new title for the document.
     *
     * @param {string} docId - The ID of the Google Document.
     * @param {string} docTitle - The new title for the document.
     * @returns {GoogleAppsScript.Document.Body} The body element of the Google Document.
     */
    presetInDoc(docId, docTitle) {
        let doc = DocumentApp.openById(docId);
        let body = doc.getBody();
        body.clear();
        doc.setName(docTitle);
        
        if (this.period === 'today') {
            let introParagraph = body.appendParagraph(`*Once the item is completed, input "C"!`);
            introParagraph.editAsText().setForegroundColor("#FF0000");
            introParagraph.setBold(false);
        }

        return body;
    }

    /**
     * Creates tables in a Google Document for each sheet's reminder data.
     * Each table contains tasks and related information from a specific sheet.
     *
     * @param {GoogleAppsScript.Document.Body} body - The body element of the Google Document.
     * @param {Array<SheetReminder>} reminderData - Array of SheetReminder objects containing the reminder data.
     */
    createReminderTablesInDoc(body, reminderData) {
        reminderData.forEach(sheetReminder => {
            let title = body.appendParagraph(sheetReminder.sheetName);
            title.setHeading(DocumentApp.ParagraphHeading.HEADING1);
            title.setLinkUrl(sheetReminder.sheetURL);
            title.setBold(true).setFontSize(12);

            // Define headers based on period
            let headers = this.period === 'today' ? ["Item", "Summary", "Date", "Staff", "Complete"] : ["Item", "Summary", "Date", "Staff"];
            this.createEachTable(body, sheetReminder.taskData, headers);
        });
    }

    /**
     * Creates a table in a Google Document for the tasks of a single sheet.
     * Sets up headers and populates the table with task data.
     *
     * @param {GoogleAppsScript.Document.Body} body - The body element of the Google Document.
     * @param {Array<Reminder>} taskData - Array of Reminder objects containing tasks for the specific sheet.
     * @param {Array<string>} headers - Array of header titles for the table.
     */
    createEachTable(body, taskData, headers) {
        let numRows = taskData.length + 1; // +1 for header row
        let numCols = headers.length;
        let table = body.appendTable(new Array(numRows).fill(0).map(row => new Array(numCols).fill('')));

            // Format the header row
            //Adjust the columnWidths with your preference
            let columnWidths = this.period === 'today' ? [100, 200, 70, 50, 70] : [100, 250, 70, 70];
            let headerRow = table.getRow(0);
            for (let i = 0; i < headers.length; i++) {
                headerRow.getCell(i).setText(headers[i]).setWidth(columnWidths[i]).setBold(true).setFontSize(10);
            }

            // Fill in the table content
            for (let i = 0; i < taskData.length; i++) {
              //Adjust the font size with your preference
              table.getRow(i + 1).getCell(0).setText(taskData[i].item).setPaddingLeft(10).setBold(false).setFontSize(9);
              table.getRow(i + 1).getCell(1).setText(taskData[i].note).setPaddingLeft(10).setBold(false).setFontSize(7);
              table.getRow(i + 1).getCell(2).setText(taskData[i].date).setPaddingLeft(10).setBold(false).setFontSize(8);
              table.getRow(i + 1).getCell(3).setText(taskData[i].staff).setPaddingLeft(10).setBold(false).setFontSize(8);
                
                // Only add completion status if the column exists
                if (numCols > 4) {
                    table.getRow(i + 1).getCell(4).setText("").setPaddingLeft(10).setBold(false).setFontSize(10);
                }
            }
    }

    /**
     * Sends an email with a reminder.
     * Uses a template file for the HTML body and includes details about the reminder.
     *
     * @param {string} email - The email address to send the reminder to.
     * @param {string} subject - The subject of the email.
     * @param {string} successOrFailure - Indicator of whether the reminder was successfully created or not.
     * @param {string} [displayDocUrl=""] - The URL of the Google Document containing the reminder, if applicable.
     */
    sendEmail(email,subject,successOrFailure,displayDocUrl){
        let template = HtmlService.createTemplateFromFile('reminder-share-email');
        template.displayDocUrl = displayDocUrl;
        template.period = this.period;
        template.successOrFailure = successOrFailure;
        if(successOrFailure === "failure"){
          template.type = this.type;
          template.target = this.target;
          template.spreadSheetUrl = this.ss.getUrl();
        }
        let htmlBody = template.evaluate().getContent();
        GmailApp.sendEmail(email,subject,"",{
            htmlBody: htmlBody,
        });
    }
    
    /**
     * Filters the reminder data for a specific staff member.
     * Returns a modified array of SheetReminder objects containing only tasks assigned to the specified staff.
     *
     * @param {Array<SheetReminder>} reminderData - Array of SheetReminder objects containing the reminder data.
     * @param {string} staffName - The name of the staff member to filter reminders for.
     * @returns {Array<SheetReminder>} An array of SheetReminder objects with tasks for the specified staff.
     */
    filterRemindersForStaff(reminderData, staffName) {
      // Filter and map the reminders for the specific staff
      return reminderData.filter(sheetReminder => {
        // Check if the tasks array exists and has tasks assigned to the staff
        return sheetReminder.taskData.some(task => task.staff === staffName);
      }).map(sheetReminder => {
        // Return a new SheetReminder with only the tasks for this staff
        return new SheetReminder(
          sheetReminder.sheetName,
          sheetReminder.sheetURL,
          sheetReminder.taskData.filter(task => task.staff === staffName)
        );
      });
    }

    /**
     * Constructs the key used for storing the completion status of reminder data in script properties.
     * 
     * @returns {string} - The constructed completion status key.
     */
    getCompletionStatusKey() {
        return `SCRIPT_PROPERTY_KEY_COMPLETION_STATUS_${this.target.toUpperCase()}_${this.period.toUpperCase()}`;
    }

    /**
     * Creates a one-time trigger for a specific reminder function based on the target and period.
     * The trigger is set to execute after a specified delay.
     * A unique identifier for the trigger is stored in script properties along with target and period information.
     */
    createOneTimeTrigger() {
        // Set a trigger for continuation
        let triggerFunctionName = `run${this.target.charAt(0).toUpperCase() + this.target.slice(1)}Reminder${this.period.charAt(0).toUpperCase() + this.period.slice(1)}`;
        let trigger = ScriptApp.newTrigger(triggerFunctionName)
                              .timeBased()
                              .after(10000) // For example, 10 seconds
                              .create();

        let triggerInfo = {
            id: trigger.getUniqueId(),
            target: this.target, // e.g., 'general'
            period: this.period  // e.g., 'today'
        };

        let scriptProperties = PropertiesService.getScriptProperties();
        scriptProperties.setProperty(trigger.getUniqueId(), JSON.stringify(triggerInfo));
        console.log(`Trigger for the function, ${triggerFunctionName} is set. triggerInfo is ${JSON.stringify(triggerInfo)}`);
    }

    /**
     * Deletes one-time triggers that match the specific target and period of the ReminderManager instance.
     * It retrieves each trigger's information from script properties and checks if it matches
     * the target and period before deletion.
     */
    deleteOneTimeTriggers() {
          let allTriggers = ScriptApp.getProjectTriggers();
          let scriptProperties = PropertiesService.getScriptProperties();

          for (let i = 0; i < allTriggers.length; i++) {
              let triggerId = allTriggers[i].getUniqueId();
              let triggerInfoStr = scriptProperties.getProperty(triggerId);

              if (triggerInfoStr) {
                  let triggerInfo = JSON.parse(triggerInfoStr);

                  // Check if the trigger is a one-time trigger for the specific target and period
                  if (triggerInfo.target === this.target && triggerInfo.period === this.period) {
                      ScriptApp.deleteTrigger(allTriggers[i]);
                      scriptProperties.deleteProperty(triggerId); // Clean up the property
                      console.log(`Following trigger and script properties for that trigger were deleted: trigger_${triggerId}/ script properties_${triggerInfoStr}`);
                  }
              }
          }
    }

}

/**
 * Triggers the function to display today's general reminder.
 * It collects data for today's tasks and updates the Google Doc with reminder information.
 */
function runGeneralReminderToday() {
    let generalReminderToday = new ReminderManager('general', 'today');
    generalReminderToday.shareRemindersByDoc();
}

/**
 * Triggers the function to display next week's general reminder.
 * It collects data for next week's tasks and updates the Google Doc with reminder information.
 */
function runGeneralReminderWeek() {
    let generalReminderWeek = new ReminderManager('general', 'week');
    generalReminderWeek.shareRemindersByDoc();
}

/**
 * Triggers the function to display today's reminder for each of designated staff.
 * It collects data for today's tasks and updates the Google Doc with reminder information.
 */
function runStaffBasedReminderToday() {
    let staffBasedReminderToday = new ReminderManager('staffBased', 'today');
    staffBasedReminderToday.shareRemindersByDoc();
}

/**
 * Triggers the function to display next week's reminder for each of designated staff.
 * It collects data for next week's tasks and updates the Google Doc with reminder information.
 */
function runStaffBasedReminderNextWeek() {
    let staffBasedReminderWeek = new ReminderManager('staffBased', 'week');
    staffBasedReminderWeek.shareRemindersByDoc();
}

// function deleteAllTriggers() {
//     let allTriggers = ScriptApp.getProjectTriggers();

//     for (let i = 0; i < allTriggers.length; i++) {
//         ScriptApp.deleteTrigger(allTriggers[i]);
//     }
// }



