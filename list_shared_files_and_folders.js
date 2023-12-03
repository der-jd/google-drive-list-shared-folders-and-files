// Script iterates recursively over all folders and files in Google Drive to find shared folders and files.
// The results are written to a spreadsheet.
// The iteration itself needs multiple script executions because the predefined maximum runtime for scripts is too short to check all folders and files.
// To allow future script executions to continue the last iteration, the recursive iterator is persisted in user properties of this script.

// --- Start folder for iteration ---

// FORCE_NEW_ITERATION: true
// START_FOLDER_PATH:   {doesn't matter}
// Recursive iterator:  {doesn't matter}
// --> Start new iteration from Google Drive root folder in any case

// FORCE_NEW_ITERATION: false
// START_FOLDER_PATH:   "..."
// Recursive iterator:  {doesn't matter}
// --> Start new iteration from given folder path

// FORCE_NEW_ITERATION: false
// START_FOLDER_PATH:   ""
// Recursive iterator:  []
// --> Start new iteration from Google Drive root folder

// FORCE_NEW_ITERATION: false
// START_FOLDER_PATH:   ""
// Recursive iterator:  [...]
// --> Resume iteration from last iterator

// --- CONFIGURATION ---

// Specify a folder path in Google Drive to use as starting point for the script execution.
// Notice: Start the path without root folder!
// --> I.e. to start from "My Drive/my/subfolder" enter "my/subfolder"
const START_FOLDER_PATH = "";

// Force a new iteration starting from the root folder.
// ATTENTION: If true, this setting will start a new iteration in any case! A set START_FOLDER_PATH and a persisted iterator are ignored!
const FORCE_NEW_ITERATION = false;

// Call script again automatically if iteration is not finished
// Notice: This setting only works if START_FOLDER_PATH is empty and FORCE_NEW_ITERATION is false! Otherwise there would be an infinite loop of calls.
const CALL_SCRIPT_AGAIN = true;

// Name of the report spreadsheet in the Google Drive root folder to save the results.
// If the spreadsheet doesn't exist yet, a new one is created.
const REPORT_SPREADSHEET_NAME = "shared_files_and_folders_report";

// Cell in the report spreadsheet to store the start date and time of the last run
const REPORT_SPREADSHEET_CELL_LAST_RUN = "B1";

// Cell in the report spreadsheet to store the number of total runs for this iteration.
const REPORT_SPREADSHEET_CELL_ITERATIONS = "B2";

// Cell in the report spreadsheet to indicate if the last iteration has been finished or if the iteration needs to be resumed.
const REPORT_SPREADSHEET_CELL_ITERATION_FINISHED = "B3";

// Stop script execution after specified runtime and save current iteration state for next execution.
// This avoids exceeding the maximum execution time predefined by Google and ensures that the current progress is persisted.
const MAX_EXECUTION_TIME_IN_MS = 5 * 60 * 1000; // 5 min

// Key for the recursive iterator in the user properties of the script.
// Used to persist the last iteration progress between multiple script executions.
const RECURSIVE_ITERATOR_KEY = "RECURSIVE_ITERATOR_KEY";



// Define log levels
const LOG_LEVEL = {
  DEBUG: 1,
  INFO: 2,
  WARN: 3,
  ERROR: 4
};

// Set the current log level (adjust as needed)
const currentLogLevel = LOG_LEVEL.DEBUG; // Change this to control log level

// Log messages based on log level
function logMessage(message, logLevel) {
  if (logLevel >= currentLogLevel) {
    switch (logLevel) {
      case LOG_LEVEL.DEBUG:
        console.log("[DEBUG] " + message);
        break;
      case LOG_LEVEL.INFO:
        console.info("[INFO] " + message);
        break;
      case LOG_LEVEL.WARN:
        console.warn("[WARN] " + message);
        break;
      case LOG_LEVEL.ERROR:
        console.error("[ERROR] " + message);
        break;
      default:
        console.log(message);
        break;
    }
  }
}

// TODO Delete after test
function temp() {
  let userProperties = PropertiesService.getUserProperties();
  //userProperties.deleteAllProperties()
  logMessage("All set user properties for the current user and script:\n" + JSON.stringify(userProperties.getProperties()), LOG_LEVEL.INFO);
  console.log(DriveApp.getRootFolder().getName());
  console.log((new Date()).toLocaleString());
}


function main() {
  const startTime = (new Date()).getTime();

  let userProperties = PropertiesService.getUserProperties();

  formattedPropertiesForPrinting = JSON.stringify(userProperties.getProperties()).replace(/\\/g, ""); // Remove all backslashes "\" (used as escape character for quotes)
  formattedPropertiesForPrinting = formattedPropertiesForPrinting.replace(/\"\[/g, "["); // Remove all quotes at the beginning of arrays
  formattedPropertiesForPrinting = formattedPropertiesForPrinting.replace(/\]\"/g, "]"); // Remove all quotes at the end of the arrays
  formattedPropertiesForPrinting = JSON.stringify(JSON.parse(formattedPropertiesForPrinting), null, 2);
  logMessage("All set user properties for the current user and script:\n" + formattedPropertiesForPrinting, LOG_LEVEL.INFO);
  logMessage(`End script execution after ${MAX_EXECUTION_TIME_IN_MS/1000} s`, LOG_LEVEL.INFO);

  logMessage("Create list of shared files and folders...", LOG_LEVEL.INFO);

  let spreadsheet = prepareSpreadsheetForReport();
  let recursiveIterator = prepareIteration(spreadsheet);
  let sheet = spreadsheet.getActiveSheet();
  logMessage("List files and folders and populate the spreadsheet...", LOG_LEVEL.INFO);

  while (recursiveIterator.length > 0) {
    recursiveIterator = listFilesAndFolders(recursiveIterator, sheet);

    let currentTime = (new Date()).getTime();
    let elapsedTimeInMS = currentTime - startTime;
    let timeLimitExceeded = elapsedTimeInMS >= MAX_EXECUTION_TIME_IN_MS;
    if (timeLimitExceeded) {
      userProperties.setProperty(RECURSIVE_ITERATOR_KEY, JSON.stringify(recursiveIterator));
      sheet.getRange(REPORT_SPREADSHEET_CELL_ITERATION_FINISHED).setValue("no");
      logMessage(`Stop iteration after "${elapsedTimeInMS/1000}" seconds.`, LOG_LEVEL.INFO);

      if (CALL_SCRIPT_AGAIN === true) {
        if (FORCE_NEW_ITERATION === true || START_FOLDER_PATH !== "") {
          logMessage("Script can't be called automatically again! The configuration FORCE_NEW_ITERATION must be false and START_FOLDER_PATH must be empty.", LOG_LEVEL.WARN);
        }
        else {
          logMessage("Call script again automatically to continue iteration.", LOG_LEVEL.INFO);
          // todo enter api call link.
          return;
        }
      }

      logMessage("Run script again manually to resume iteration.", LOG_LEVEL.INFO);
      return;
    }
  }

  userProperties.deleteProperty(RECURSIVE_ITERATOR_KEY);
  sheet.getRange(REPORT_SPREADSHEET_CELL_ITERATION_FINISHED).setValue("yes");
  logMessage("Iteration finished!", LOG_LEVEL.INFO);
}


function prepareSpreadsheetForReport() {
  logMessage(`Save results in spreadsheet "${REPORT_SPREADSHEET_NAME}"`, LOG_LEVEL.INFO);

  let existingFiles = DriveApp.getFilesByName(REPORT_SPREADSHEET_NAME);
  if (existingFiles.hasNext()) {
    logMessage("Use existing spreadsheet...", LOG_LEVEL.INFO);
    return SpreadsheetApp.open(existingFiles.next());
  }
  else {
    logMessage("Create new spreadsheet...", LOG_LEVEL.INFO);
    return SpreadsheetApp.create(REPORT_SPREADSHEET_NAME);
  }
}


function prepareIteration(spreadsheet) {
  // [{folderName: String, fileIteratorContinuationToken: String?, folderIteratorContinuationToken: String}]
  // Each folder has its own entry in the iterator array.
  // Example: "path/to/folder" --> [{folderName: "path", ...}, {folderName: "to", ...}, {folderName: "folder", ...}]
  let recursiveIterator = JSON.parse(PropertiesService.getUserProperties().getProperty(RECURSIVE_ITERATOR_KEY));

  if (FORCE_NEW_ITERATION === true) {
    logMessage("[Force new iteration] Start new iteration from Google Drive root folder...", LOG_LEVEL.INFO);
    recursiveIterator = [];
    recursiveIterator.push(makeIterationFromFolder(DriveApp.getRootFolder()));

    createNewSheet(spreadsheet);
  }
  else {
    if (START_FOLDER_PATH !== "") {
      logMessage(`[Start folder given] Start new iteration from folder "${START_FOLDER_PATH}"...`, LOG_LEVEL.INFO);
      recursiveIterator = [];
      recursiveIterator.push(makeIterationFromFolder(getFolderByPath(START_FOLDER_PATH)));

      createNewSheet(spreadsheet);
    }
    else {
      if (recursiveIterator !== null) {
        const folderPath = recursiveIterator.map(function(iteration) { return iteration.folderName; }).join("/");
        logMessage(`[Use persisted iterator] Resume iteration from folder "${folderPath}"...`, LOG_LEVEL.INFO);

        updateSheet(spreadsheet);
      }
      else {
        logMessage("[Persisted iterator empty] Start new iteration from Google Drive root folder...", LOG_LEVEL.INFO);
        recursiveIterator = [];
        recursiveIterator.push(makeIterationFromFolder(DriveApp.getRootFolder()));

        createNewSheet(spreadsheet);
      }
    }
  }

  return recursiveIterator;
}


function createNewSheet(spreadsheet) {
  logMessage("Insert new sheet for iteration...", LOG_LEVEL.INFO);
  let sheet = spreadsheet.insertSheet(0); // Insert a new sheet as the first one
  sheet.appendRow(["Start time of the last run", (new Date()).toLocaleString()]);
  sheet.appendRow(["Number of runs", 1]);
  sheet.appendRow(["Iteration finished", "running"]);
  sheet.appendRow(["Name", "Type", "Shared"]);
}


function updateSheet(spreadsheet) {
  logMessage("Use and update sheet from last iteration...", LOG_LEVEL.INFO);
  let sheet = spreadsheet.getActiveSheet();
  sheet.getRange(REPORT_SPREADSHEET_CELL_LAST_RUN).setValue((new Date()).toLocaleString());
  sheet.getRange(REPORT_SPREADSHEET_CELL_ITERATIONS).setValue(sheet.getRange(REPORT_SPREADSHEET_CELL_ITERATIONS).getValue() + 1);
  sheet.getRange(REPORT_SPREADSHEET_CELL_ITERATION_FINISHED).setValue("running");
}


function getFolderByPath(path) {
  var folders = path.split('/');
  var currentFolder = DriveApp.getRootFolder();

  for (var i = 0; i < folders.length; i++) {
    var folderName = folders[i];
    var subfolders = currentFolder.getFoldersByName(folderName);

    if (subfolders.hasNext()) {
      currentFolder = subfolders.next();
    }
    else {
      return null; // Folder not found
    }
  }

  return currentFolder;
}


function makeIterationFromFolder(folder) {
  return {
    folderName: folder.getName(), 
    fileIteratorContinuationToken: folder.getFiles().getContinuationToken(),
    folderIteratorContinuationToken: folder.getFolders().getContinuationToken()
  };
}


function listFilesAndFolders(recursiveIterator, sheet) {
  let currentIteration = recursiveIterator[recursiveIterator.length-1];

  if (currentIteration.fileIteratorContinuationToken !== null) {
    let fileIterator = DriveApp.continueFileIterator(currentIteration.fileIteratorContinuationToken);

    if (fileIterator.hasNext()) {
      // Process the next file
      let path = recursiveIterator.map(function(iteration) { return iteration.folderName; }).join("/");
      let file = fileIterator.next();
      let access = getSharingAccess(file);
      if (access != "Private") {
        logMessage(`Add shared file "${path + file.getName()}" to sheet...`, LOG_LEVEL.INFO);
        sheet.appendRow([path + file.getName(), "File", access]);
      }

      currentIteration.fileIteratorContinuationToken = fileIterator.getContinuationToken();
      recursiveIterator[recursiveIterator.length-1] = currentIteration;
      return recursiveIterator;
    }
    else {
      // Done processing files
      currentIteration.fileIteratorContinuationToken = null;
      recursiveIterator[recursiveIterator.length-1] = currentIteration;
      return recursiveIterator;
    }
  }

  if (currentIteration.folderIteratorContinuationToken !== null) {
    let folderIterator = DriveApp.continueFolderIterator(currentIteration.folderIteratorContinuationToken);

    if (folderIterator.hasNext()) {
      // Process the next folder
      let path = recursiveIterator.map(function(iteration) { return iteration.folderName; }).join("/");
      let folder = folderIterator.next();
      let access = getSharingAccess(folder);
      if (access != "Private") {
        logMessage(`Add shared folder "${path + folder.getName()}" to sheet...`, LOG_LEVEL.INFO);
        sheet.appendRow([path + folder.getName(), "Folder", access]);
      }

      recursiveIterator[recursiveIterator.length-1].folderIteratorContinuationToken = folderIterator.getContinuationToken();
      recursiveIterator.push(makeIterationFromFolder(folder));
      return recursiveIterator;
    }
    else {
      // Done processing subfolders
      recursiveIterator.pop();
      return recursiveIterator;
    }
  }

  logMessage("Iterator failure!", LOG_LEVEL.ERROR);
  throw "Should never get here. Iterator failure!";
}


// TODO delete
//function listFilesAndFolders(folder, folderPath, sheet) {
//  let subfolders = folder.getFolders();
//  while (subfolders.hasNext()) {
//    let subfolder = subfolders.next();
//
//    let access = getSharingAccess(subfolder);
//    if (access != "Private") {
//      logMessage(`Add shared folder "${folderPath + subfolder.getName()}" to sheet...`, LOG_LEVEL.INFO);
//      sheet.appendRow([folderPath + subfolder.getName(), "Folder", access]);
//    }
//  
//    listFilesAndFolders(subfolder, folderPath + subfolder.getName() + "/", sheet);
//  }
//
//  let files = folder.getFiles();
//  while (files.hasNext()) {
//    let file = files.next();
//    let access = getSharingAccess(file);
//    if (access != "Private") {
//      logMessage(`Add shared file "${folderPath + file.getName()}" to sheet...`, LOG_LEVEL.INFO);
//      sheet.appendRow([folderPath + file.getName(), "File", access]);
//    }
//  }
//}


function getSharingAccess(item) {
  logMessage(item.getName(), LOG_LEVEL.DEBUG);

  // Check if any individual user, who might not be me, has access
  if (item.getSharingAccess() == DriveApp.Access.PRIVATE) {

    effectiveUserMail = Session.getEffectiveUser().getEmail();
    if (item.getOwner().getEmail() != effectiveUserMail) {
      return "Shared"; // Others have access
    }

    viewers = item.getViewers();
    for (let i = 0; i < viewers.length; i++) {
      if (viewers[i].getEmail() != effectiveUserMail) {
        return "Shared"; // Others have access
      }
    }

    editors = item.getEditors();
    for (let i = 0; i < editors.length; i++) {
      if (editors[i].getEmail() != effectiveUserMail) {
        return "Shared"; // Others have access
      }
    }

    return "Private"; // Only you have access
  }
  else {
    return "Shared"; // Access is not set to PRIVATE, so it's shared
  }
}
