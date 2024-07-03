/**
 * Extracts a zip file and creates a folder structure in Google Drive
 * @param {string} fileId - The ID of the zip file in Google Drive
 * @param {string} destinationFolderId - The ID of the destination folder in Google Drive
 * @param {string} newFolderName - The name for the new main folder (optional)
 * @returns {string} JSON string containing the extraction results
 */
function extractZipAndCreateFolders(fileId, destinationFolderId, newFolderName) {
  var startTime = new Date();
  var success = true;
  var errors = [];
  var totalFiles = 0;
  var totalSize = 0;
  var mainFolderName, mainFolderId, folderIdMap;

  try {
    // Step 1: Get the destination folder in Google Drive
    var destinationFolder;
    try {
      destinationFolder = DriveApp.getFolderById(destinationFolderId);
    } catch (e) {
      throw new Error("Invalid destination folder ID: " + destinationFolderId);
    }

    // Step 2: Download the .zip file as a blob
    var zipFile;
    try {
      zipFile = DriveApp.getFileById(fileId);
    } catch (e) {
      throw new Error("Invalid zip file ID: " + fileId);
    }
    var blob = zipFile.getBlob();

    // Step 3: Determine the main folder name
    mainFolderName = newFolderName || zipFile.getName().replace('.zip', '');

    // Step 4: Ensure the folder name is unique
    mainFolderName = getUniqueFolderName(destinationFolder, mainFolderName);

    // Step 5: Create the main folder
    var mainFolder;
    try {
      mainFolder = destinationFolder.createFolder(mainFolderName);
    } catch (e) {
      throw new Error("Failed to create main folder: " + e.message);
    }
    mainFolderId = mainFolder.getId();

    // Step 6: Unzip the blob into individual files
    var unzippedFiles;
    try {
      unzippedFiles = Utilities.unzip(blob);
    } catch (e) {
      throw new Error("Failed to unzip file: " + e.message);
    }

    // Step 7: Initialize the folder structure object
    var folderStructure = {
      name: mainFolderName,
      id: mainFolderId,
      subFolders: []
    };

    folderIdMap = {};
    folderIdMap[mainFolderId] = folderStructure;

    var createdFiles = [];

    // Step 8: First pass - Create all necessary folders
    for (var i = 0; i < unzippedFiles.length; i++) {
      var file = unzippedFiles[i];
      var filePath = file.getName();

      // Skip the root folder name if it exists
      if (filePath.indexOf('/') !== -1) {
        filePath = filePath.substring(filePath.indexOf('/') + 1);
      }

      var folderPath = filePath.substring(0, filePath.lastIndexOf('/'));
      if (folderPath) {
        try {
          createFolderStructure(mainFolder, folderPath, folderStructure, folderIdMap);
        } catch (e) {
          errors.push("Failed to create folder structure for path: " + folderPath + ". Error: " + e.message);
        }
      }
    }

    // Step 9: Second pass - Create files in the target folders
    for (var i = 0; i < unzippedFiles.length; i++) {
      var file = unzippedFiles[i];
      var filePath = file.getName();

      // Skip the root folder name if it exists
      if (filePath.indexOf('/') !== -1) {
        filePath = filePath.substring(filePath.indexOf('/') + 1);
      }

      var folderPath = filePath.substring(0, filePath.lastIndexOf('/'));
      var fileName = filePath.substring(filePath.lastIndexOf('/') + 1);

      // Skip entries that represent folders or the root folder
      if (fileName === "" || filePath === "") continue;

      // Get the target folder where the file will be placed
      var targetFolder;
      try {
        targetFolder = folderPath ? DriveApp.getFolderById(getFolderIdFromPath(folderPath, folderIdMap)) : mainFolder;
      } catch (e) {
        errors.push("Failed to get target folder for path: " + folderPath + ". Error: " + e.message);
        continue;
      }

      // Create the file in the target folder
      var newFile;
      try {
        if (fileName.endsWith('.zip')) {
          // If the file is a zip, create it without extracting
          newFile = targetFolder.createFile(file.setName(fileName));
        } else {
          newFile = targetFolder.createFile(file);
        }
      } catch (e) {
        errors.push("Failed to create file: " + fileName + ". Error: " + e.message);
        continue;
      }
      
      // Update total file count and size
      totalFiles++;
      totalSize += newFile.getSize();

      // Add file information to the createdFiles array
      createdFiles.push({
        name: newFile.getName(),
        id: newFile.getId(),
        path: folderPath,
        folderId: targetFolder.getId(),
        size: newFile.getSize(),
        mimeType: newFile.getMimeType(),
        createdTime: newFile.getDateCreated().toISOString()
      });
    }

  } catch (error) {
    // If a major error occurs, set success to false and log the error
    success = false;
    errors.push(error.toString());
    console.error("Error in extractZipAndCreateFolders: " + error);
  }

  // Step 10: Prepare return data
  var returnData = {
    success: success,
    timestamp: startTime.toISOString(),
    mainFolder: {
      name: mainFolderName,
      id: mainFolderId
    },
    summary: {
      totalFiles: totalFiles,
      totalSize: totalSize,
      folderCount: Object.keys(folderIdMap || {}).length
    },
    folderStructure: folderStructure,
    files: createdFiles,
    errors: errors
  };

  // Step 11: Return the result as a JSON string
  return JSON.stringify(returnData);
}

/**
 * Creates a folder structure based on the given path
 * @param {Folder} parentFolder - The parent folder object
 * @param {string} folderPath - The path of folders to create
 * @param {Object} currentStructure - The current folder structure object
 * @param {Object} folderIdMap - A map of folder IDs to their structure objects
 * @returns {string} The ID of the last created folder
 */
function createFolderStructure(parentFolder, folderPath, currentStructure, folderIdMap) {
  var folders = folderPath.split('/');
  var currentFolder = parentFolder;
  var currentStructureNode = currentStructure;

  for (var i = 0; i < folders.length; i++) {
    var folderName = folders[i];
    var existingFolder = currentFolder.getFoldersByName(folderName);

    // If the folder exists, use it; otherwise, create a new one
    if (existingFolder.hasNext()) {
      currentFolder = existingFolder.next();
    } else {
      currentFolder = currentFolder.createFolder(folderName);
    }

    // Update the folder structure and ID map
    if (!folderIdMap[currentFolder.getId()]) {
      var newStructureNode = {
        name: folderName,
        id: currentFolder.getId(),
        subFolders: []
      };
      currentStructureNode.subFolders.push(newStructureNode);
      folderIdMap[currentFolder.getId()] = newStructureNode;
      currentStructureNode = newStructureNode;
    } else {
      currentStructureNode = folderIdMap[currentFolder.getId()];
    }
  }

  return currentFolder.getId();
}

/**
 * Gets the folder ID from a given path
 * @param {string} folderPath - The path of the folder
 * @param {Object} folderIdMap - A map of folder IDs to their structure objects
 * @returns {string} The ID of the folder at the given path
 */
function getFolderIdFromPath(folderPath, folderIdMap) {
  var folders = folderPath.split('/');
  var currentNode = folderIdMap[Object.keys(folderIdMap)[0]]; // Start from the root

  for (var i = 0; i < folders.length; i++) {
    var folderName = folders[i];
    var found = false;
    for (var j = 0; j < currentNode.subFolders.length; j++) {
      if (currentNode.subFolders[j].name === folderName) {
        currentNode = currentNode.subFolders[j];
        found = true;
        break;
      }
    }
    if (!found) {
      throw new Error("Folder not found: " + folderPath);
    }
  }

  return currentNode.id;
}

/**
 * Generates a unique folder name by appending a timestamp if necessary
 * @param {Folder} parentFolder - The parent folder object
 * @param {string} baseName - The base name for the folder
 * @returns {string} A unique folder name
 */
function getUniqueFolderName(parentFolder, baseName) {
  var name = baseName;
  var counter = 1;
  while (parentFolder.getFoldersByName(name).hasNext()) {
    name = baseName + ' (' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd-MMM-yyyy HH:mm:ss") + ')';
    counter++;
  }
  return name;
}
