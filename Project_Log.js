/**
 *@fileoverview This Google Apps SCript automates the process of extracting specific data
 *from Google Sheet files located in a designated Google Drive folder and appending it
 *to a central Status Log Google Sheet. It also includes functionality to log manual edits
 *made directly within the Status Log spreadsheet.
 *
 *This version expands on the initial script to include:
 *-Conversion of new Excel (.xlsx) files to Google Sheets format.
 *-Automatic setting of "Created Date".
 *-Insertion of new entries at the top of the log.
 *-Comprehensive data extraction from muliple sheets within source files.
 *-Sequiential unique reference number generation
 *-Prevention of duplicate entries by tracking processed file IDs.
 *-Integration of a robust edit logging function ('handleEdit').
 *-**Time-driven trigger for processing files, as a workaround for Drive trigger UI issue.**
/*

// -Configuration ---

/**
* @constant {string} STATUS_LOG_ID tHE id of your main Status Log Google Sheet.
* You can find this in the URL of your spreadsheet (e.g., https://docs.google.com/spreadsheets/d/THIS_ID_IS_HERE/edit).
*/
const STATUS_LOG_ID = 'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX'; //User provided id of spreadsheet

/**

