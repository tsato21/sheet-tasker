/*
Constant Variables ~ Sheet Names
ONGOING_TASKS_INDEX_SHEET_NAME: Name of the sheet that lists ongoing tasks.
COMPLETED_TASKS_INDEX_SHEET_NAME: Name of the sheet that lists completed tasks.
BUTTON_TO_INDEX_SHEET: Label for the button used to navigate back to the index sheet.
*/
const ONGOING_TASKS_INDEX_SHEET_NAME = "ongoing-task-index";
const COMPLETED_TASKS_INDEX_SHEET_NAME = "completed-task-index";
const BUTTON_TO_INDEX_SHEET = "Back to Index";

/*
Constant Variables ~ Script Property Keys
SCRIPT_PROPERTY_KEY_STAFF: Key for storing staff information in script properties.
SCRIPT_PROPERTY_KEY_GENERAL_REMINDER_EMAILS: Key for storing email addresses for general reminders.
SCRIPT_PROPERTY_KEY_DESIG_STAFF: Key for storing designated staff information for staff-based reminders.
SCRIPT_PROPERTY_KEY_GENERAL_REM_DOC_URL: Key for storing URLs of Google Docs for general reminders.
SCRIPT_PROPERTY_KEY_STAFFBASED_REM_DATA: Key for storing data related to staff-based reminders.
SCRIPT_PROPERTY_KEY_CURRENT_SHEET_INDEX: Key to store the index of the current sheet being processed.
SCRIPT_PROPERRY_KEY_STORED_REMINDERS: Key to store temporarily saved reminder data.
SCRIPT_PROPERRY_KEY_COMPLETION_STATUS: Key to track the completion status of a task or operation.
*/
const SCRIPT_PROPERTY_KEY_STAFF = 'STAFF_DATA';
const SCRIPT_PROPERTY_KEY_GENERAL_REMINDER_EMAILS = 'GENERAL_REMINDER_EMAILS';
const SCRIPT_PROPERTY_KEY_DESIG_STAFF = 'DESIG_STAFF';
const SCRIPT_PROPERTY_KEY_GENERAL_REM_DOC_URL = 'GENERAL_REM_DOC_URL';
const SCRIPT_PROPERTY_KEY_STAFFBASED_REM_DATA = 'STAFFBASED_REM_DATA';
const SCRIPT_PROPERTY_KEY_CURRENT_SHEET_INDEX = 'CURRENT_SHEET_INDEX';
const SCRIPT_PROPERRY_KEY_STORED_REMINDERS = 'STORED_REMINDERS';
const SCRIPT_PROPERRY_KEY_COMPLETION_STATUS = 'COMPLETION_STATUS';

/* Lookup object for script property keys
  This object maps the string identifiers (used in client-side interactions)
  to the actual constant values representing script property keys.
  It's used to dynamically retrieve the correct property key based on a string identifier,
  which helps in efficiently managing the deletion or manipulation of script properties
  without hardcoding multiple if-else conditions or switch cases.
*/
const PROPERTY_KEYS = {
    'SCRIPT_PROPERTY_KEY_STAFF': SCRIPT_PROPERTY_KEY_STAFF, // Maps to the property key for staff data
    'SCRIPT_PROPERTY_KEY_DESIG_STAFF': SCRIPT_PROPERTY_KEY_DESIG_STAFF, // Maps to the property key for designated staff data
    'SCRIPT_PROPERTY_KEY_GENERAL_REMINDER_EMAILS': SCRIPT_PROPERTY_KEY_GENERAL_REMINDER_EMAILS, // Maps to the property key for general reminder emails
    'SCRIPT_PROPERTY_KEY_GENERAL_REM_DOC_URL': SCRIPT_PROPERTY_KEY_GENERAL_REM_DOC_URL,
    'SCRIPT_PROPERTY_KEY_STAFFBASED_REM_DATA': SCRIPT_PROPERTY_KEY_STAFFBASED_REM_DATA
};