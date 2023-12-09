// Script iterates recursively over all folders and files in Google Drive to find shared folders and files.
// The results are written to a spreadsheet.
// The iteration itself needs multiple script executions because the predefined maximum runtime for scripts is too short to check all folders and files.
// To allow future script executions to continue the last iteration, the recursive iterator is persisted in user properties of this script.
// NOTICE:
// Consider enabling an automatic, time-based trigger for the script in the Google AppScript interface.
// On this way the script will run repeatedly resuming the last iteration. You need to check manually in the result sheet if the whole iteration has been finished
// so that you can turn off the automatic trigger again.

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
// NOTICE:
// If an automatic, time-based trigger is used for the script (see comment above), the maximum execution time needs to be shorter than the trigger interval!
// Otherwise multiple script executions are started in parallel all using the same persisted iterator from the last finished execution.
// However all script executions must be strictly sequential so that the iterator is persisted and the next execution can continue the iteration.
const MAX_EXECUTION_TIME_IN_MS = 4 * 60 * 1000; // 4 min

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
const currentLogLevel = LOG_LEVEL.INFO; // Change this to control log level

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



function debug() {
  let userProperties = PropertiesService.getUserProperties();
  prettyPrintUserProperties(userProperties);
}



function main() {
  const startTime = (new Date()).getTime();

  let userProperties = PropertiesService.getUserProperties();

  prettyPrintUserProperties(userProperties);
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
      logMessage(`Stop iteration after "${elapsedTimeInMS/1000}" seconds. Run script again to resume iteration.`, LOG_LEVEL.INFO);
      prettyPrintUserProperties(userProperties);
      return;
    }
  }

  userProperties.deleteProperty(RECURSIVE_ITERATOR_KEY);
  sheet.getRange(REPORT_SPREADSHEET_CELL_ITERATION_FINISHED).setValue("yes");
  logMessage("Iteration finished!", LOG_LEVEL.INFO);
}


function prettyPrintUserProperties(userProperties) {
  let formattedProperties = JSON.stringify(userProperties.getProperties()).replace(/\\/g, ""); // Remove all backslashes "\" (used as escape character for quotes)

  formattedProperties = formattedProperties.replace("\"[", "["); // Remove quotes at the beginning of the first array

  // Remove quotes at the end of the last array
  const lastIndexOfSquareBracket = formattedProperties.lastIndexOf("]\"");
  if (lastIndexOfSquareBracket !== -1) {
    formattedProperties = formattedProperties.substring(0, lastIndexOfSquareBracket) + formattedProperties.substring(lastIndexOfSquareBracket).replace("]\"", "]");
  }

  formattedProperties = JSON.stringify(JSON.parse(formattedProperties), null, 2); // Format with indentation
  logMessage("All set user properties for the current user and script:\n" + formattedProperties, LOG_LEVEL.INFO);
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
