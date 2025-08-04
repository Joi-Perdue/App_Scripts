/*
This code can be used to help automate tasks in Google Workspace
Many people use Microsoft Office products, while others use Google Workspace products
This code converts an .xlsx spreadsheet to a Google Spreadsheet. No manual opening and saving is necessary. T
Once the data is in a Google spreadsheet further automation can occur
*/

function convertExceltoGoogleSheets() {
  const sourcFolderID  = "xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx"; // User update Google folderID
  const targetFolderID = "xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx"; // User update Google folderID

  const sourceFiles = DriveApp.getFolderById(sourceFolderID).getFilesByType(MimeType.MICROSOFT_EXCEL);
  const targetFileNames = new Set();

  let targetFile = DriveApp.getFolderById(targetFolderId).getFilesByType(MimeType.GOOGLE_SHEETS);
  while (targetFile.hasNext()) {
    targetFilesNames.add(targetFile.net().getName());
  }

  const convertedFileNames = [];

  while (sourceFiles.hasNext()) {
    const excelFile = sourceFiles.net();
    const originalFileName = excelFile.getName();
    const NewSpreadsheetName = originalFileName.replace(/\.xlsx?$/i, '');

    if (!targetFileNames.has(newSpreadsheetName)) {
      try {
        const fileMetadata = {
          name: newSpreadsheetName,
          mimeType: MimeType.GOOGLE_SHEETS,
          parents: [targetFolderId]
        };
      
        const newSpreadsheet = Drive.Files.create( // Using Drive.Files.create
          fileMetadata,
          excelFile.getBlob(),
          { convert: true }
        );
        convertedFileNames.push(origialFileName);

        console.log("Converted:', originalFilename);
      } catch (error) {
        console.error("Error converting:", originalFileName, error);
      }
    } else {
      conosole.log("Already converted:", originalFilename);
    }
  }

  if (convertedFileNames.length === 0) {
    console.log("No new Excel files found to convert.");
  } else {
    console.log("Files converted:", convertedFileNames.join(", "));
  }
}    
