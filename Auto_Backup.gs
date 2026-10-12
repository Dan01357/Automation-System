//Auto_Backup.gs
// =========================================================================
// AUTOMATED MULTI-NODE SYSTEM BACKUP (TIMEOUT-PROOF BACKGROUND PROCESS)
// =========================================================================

// --- CONFIGURATION ---
// Note: All sensitive Folder IDs are now securely pulled from PropertiesService to protect your GitHub.

// Safe limits to prevent "INTERNAL" server memory crash
const MAX_RUN_TIME = 150000; // 2.5 minutes
const MAX_FILES_PER_CHUNK = 30; // Strictly process 30 items per run

/**
 * THIS IS THE MAIN FUNCTION TO RUN OR TRIGGER AT MIDNIGHT
 */
function executeMultiNodeBackup() {
  const startTime = Date.now();
  const scriptProperties = PropertiesService.getScriptProperties();
  
  let stateString = scriptProperties.getProperty('BACKUP_STATE');
  let state;
  
  if (stateString) {
    state = JSON.parse(stateString);
    console.log("Resuming background backup...");
  } else {
    console.log("Starting a fresh backup cycle. Cleaning old backups first...");
    state = initializeNewState();
    if (!state) return; // Exit if properties are missing
  }

  let itemsProcessedThisRun = 0;

  // Run until time is up OR we hit the strict 30-file chunk limit
  while ((Date.now() - startTime < MAX_RUN_TIME) && (itemsProcessedThisRun < MAX_FILES_PER_CHUNK)) {
    
    // --- A. Process File Copies ---
    if (state.filesToCopy.length > 0) {
      let fileTask = state.filesToCopy.shift();
      try {
        let file = DriveApp.getFileById(fileTask.src);
        let destFolder = DriveApp.getFolderById(fileTask.dest);
        file.makeCopy(file.getName(), destFolder);
        
        state.totalCopied++;
        itemsProcessedThisRun++;
        
        // Hint the garbage collector to release memory
        file = null; 
        destFolder = null;
      } catch(e) {
        console.error(`Failed to copy file ${fileTask.src}: ${e.message}`);
      }
      continue; 
    }

    // --- B. Process Folders (Mapping structure) ---
    if (state.foldersToProcess.length > 0) {
      let folderTask = state.foldersToProcess.shift();
      try {
        let sourceFolder = DriveApp.getFolderById(folderTask.src);
        let destFolder = DriveApp.getFolderById(folderTask.dest);

        let subFolders = sourceFolder.getFolders();
        while (subFolders.hasNext()) {
          let subF = subFolders.next();
          let newDest = destFolder.createFolder(subF.getName());
          state.foldersToProcess.push({ src: subF.getId(), dest: newDest.getId() });
        }

        let files = sourceFolder.getFiles();
        while (files.hasNext()) {
          state.filesToCopy.push({ src: files.next().getId(), dest: destFolder.getId() });
        }
        
        itemsProcessedThisRun++;
        
        // Hint the garbage collector to release memory
        sourceFolder = null;
        destFolder = null;
      } catch(e) {
         console.error(`Failed to map folder ${folderTask.src}: ${e.message}`);
      }
      continue; 
    }

    // --- C. If queues are empty, WE ARE COMPLETELY DONE ---
    console.log(`Backup Finished! Successfully backed up ${state.totalCopied} files.`);
    scriptProperties.deleteProperty('BACKUP_STATE');
    deleteContinuationTriggers();
    return; 
  }

  // --- D. TIME OR CHUNK LIMIT REACHED: Safe Pause ---
  console.log(`Limit reached (${itemsProcessedThisRun} items). Pausing safely to prevent crash and scheduling continuation...`);
  scriptProperties.setProperty('BACKUP_STATE', JSON.stringify(state));
  scheduleContinuation();
}

// =========================================================================
// HELPER FUNCTIONS
// =========================================================================

function initializeNewState() {
  const scriptProperties = PropertiesService.getScriptProperties();
  
  // NOTE: Ensure the key 'BACKUP_FOLDER_ID' matches the name you saved in your Project Settings!
  let backupRootId = scriptProperties.getProperty('BACKUP_FOLDER_ID') || scriptProperties.getProperty('MASTER_BACKUP_FOLDER_ID'); 
  if (!backupRootId) {
     console.error("CRITICAL: Backup Folder ID property is missing in Script Properties.");
     return null;
  }
  
  const dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
  let backupRoot = DriveApp.getFolderById(backupRootId);

  // --- 1. SWEEP OLD BACKUP FOLDERS AND PREVIOUS CODE BACKUP ---
  let oldFolders = backupRoot.getFolders();
  while (oldFolders.hasNext()) {
    let f = oldFolders.next();
    if (f.getName().startsWith("System_Backup_")) {
      f.setTrashed(true);
    }
  }
  
  // Clean previous CODE_BACKUP_ files in the backup folder to guarantee only 1 remains
  let oldFiles = backupRoot.getFiles();
  while (oldFiles.hasNext()) {
    let file = oldFiles.next();
    if (file.getName().startsWith("CODE_BACKUP_")) {
      file.setTrashed(true);
      console.log("Purged old code backup: " + file.getName());
    }
  }

  // Secondary sweep by direct ID reference
  const previousCodeBackupId = scriptProperties.getProperty('LATEST_CODE_BACKUP_ID');
  if (previousCodeBackupId) {
    try {
      DriveApp.getFileById(previousCodeBackupId).setTrashed(true);
    } catch (e) {
      // Already cleaned
    }
  }

  // --- 2. CREATE NEW BACKUP FOLDER & ROTATING CODE BACKUP ---
  let todayFolder = backupRoot.createFolder("System_Backup_" + dateStr);
  let liveSheetId = SpreadsheetApp.getActiveSpreadsheet().getId();
  
  // Creates single rotating backup file containing all code intact
  let codeBackupFile = DriveApp.getFileById(liveSheetId).makeCopy("CODE_BACKUP_" + dateStr, backupRoot);
  scriptProperties.setProperty('LATEST_CODE_BACKUP_ID', codeBackupFile.getId());
  
  let state = {
    dateStr: dateStr,
    foldersToProcess: [],
    filesToCopy: [],
    totalCopied: 0,
    currentDestRootId: todayFolder.getId()
  };
  
  // --- 3. QUEUE DATABASES TO THE SINGLE LOCATION ---
  // Fetching Source IDs securely from Script Properties
  const mainDataId = scriptProperties.getProperty('MAIN_FOLDER_ID');
  const devDbId = scriptProperties.getProperty('VERIFIED_FOLDER_ID_DEV');
  const execDbId = scriptProperties.getProperty('VERIFIED_FOLDER_ID');
  
  if (!mainDataId || !devDbId || !execDbId) {
     console.error("CRITICAL: Source Database Folder IDs are missing in Script Properties.");
     return null;
  }
  
  const sourcesToBackup = [
      { id: mainDataId, name: "Main_System_Data" },
      { id: devDbId, name: "Dev_Database" },
      { id: execDbId, name: "Exec_Database" }
  ];
  
  for (let i = 0; i < sourcesToBackup.length; i++) {
      let sourceSubFolder = todayFolder.createFolder(sourcesToBackup[i].name);
      state.foldersToProcess.push({
        src: sourcesToBackup[i].id,
        dest: sourceSubFolder.getId()
      });
  }
  
  return state;
}

// --- TRIGGER MANAGEMENT SYSTEM ---

function resumeMultiNodeBackup() {
  executeMultiNodeBackup();
}

function scheduleContinuation() {
  deleteContinuationTriggers(); 
  ScriptApp.newTrigger("resumeMultiNodeBackup")
           .timeBased()
           .after(60 * 1000) // Triggers EXACTLY 1 minute later
           .create();
}

function deleteContinuationTriggers() {
  let triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === "resumeMultiNodeBackup" && 
        triggers[i].getEventType() === ScriptApp.EventType.CLOCK) {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}

/**
 * FAST TEST: Verifies that CODE_BACKUP_ is created, older backups are deleted,
 * and exactly 1 backup remains in backupRoot—finishes in 3 seconds.
 */
function testCodeBackupRotationOnly() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const backupRootId = scriptProperties.getProperty('BACKUP_FOLDER_ID') || scriptProperties.getProperty('MASTER_BACKUP_FOLDER_ID');
  
  if (!backupRootId) {
    console.error("BACKUP_FOLDER_ID is missing in Script Properties.");
    return;
  }
  
  const backupRoot = DriveApp.getFolderById(backupRootId);
  const dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");

  console.log("1. Scanning and purging existing CODE_BACKUP_ files in backupRoot...");
  let oldFiles = backupRoot.getFiles();
  let purgedCount = 0;
  while (oldFiles.hasNext()) {
    let file = oldFiles.next();
    if (file.getName().startsWith("CODE_BACKUP_")) {
      file.setTrashed(true);
      console.log("Purged old code backup: " + file.getName());
      purgedCount++;
    }
  }
  console.log("Total old backups purged: " + purgedCount);

  console.log("2. Creating single fresh CODE_BACKUP_" + dateStr + "...");
  let liveSheetId = SpreadsheetApp.getActiveSpreadsheet().getId();
  let newBackup = DriveApp.getFileById(liveSheetId).makeCopy("CODE_BACKUP_" + dateStr, backupRoot);
  scriptProperties.setProperty('LATEST_CODE_BACKUP_ID', newBackup.getId());
  
  console.log("SUCCESS! Created fresh backup: " + newBackup.getName() + " (ID: " + newBackup.getId() + ")");
  console.log("Only 1 code backup now exists in your backup folder.");
}

/**
 * RESET HELPER: Clears any lingering paused queue or state
 * so the next midnight run starts fresh with the new code backup.
 */
function resetBackupStateToFresh() {
  const scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.deleteProperty('BACKUP_STATE');
  deleteContinuationTriggers();
  console.log("SUCCESS: Backup state cleared. Tonight's midnight trigger will run a 100% fresh cycle.");
}

function displayAllScriptProperties() {
  const props = PropertiesService.getScriptProperties().getProperties();
  
  console.log("=== ALL SCRIPT PROPERTIES ===");
  
  // Loop through and print every single property and its value
  for (let key in props) {
    console.log(`Key: ${key}  --->  Value: ${props[key]}`);
  }
  
  // Also log the total count
  console.log(`TOTAL PROPERTIES: ${Object.keys(props).length}`);
}

