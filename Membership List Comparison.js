function onNewEmail(e) {
  function groupCsv(csv, key) {  // turns the "list of dicts" csv output into a dict that is easier to use, using the specified key
    var header = csv.shift();  // remove header row
    var index = 0;
    for (var i = 0; i < header.length; i++) {
      if (header[i] == key) {
        index = i;
        break;
      }
    }

    var sortedCsv = {};
    for (var i = 0; i < csv.length; i++) {
      var row = csv[i];
      keyVal = row[index];
      sortedCsv[keyVal] = row;
    }

    return sortedCsv;

  }

  function getColumnLetters(columnIndexStartFromOne) {  // retrieves the column letter for Sheets given an index number
    const ALPHABETS = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z']

    if (columnIndexStartFromOne < 27) {
      return ALPHABETS[columnIndexStartFromOne - 1]
    } else {
      var res = columnIndexStartFromOne % 26
      var div = Math.floor(columnIndexStartFromOne / 26)
      if (res === 0) {
        div = div - 1
        res = 26
      }
      return getColumnLetters(div) + ALPHABETS[res - 1]
    }
  }

  // TODO: Create a label in gmail for your membership list emails and make sure they are auto-sorted into this label. Update the label name here if it differs
  var threads = GmailApp.search('label:membership-actionkit-membership-lists');
  var messages = threads[0].getMessages();
  
  // TODO: Update with the name of your chapter, as used in the membership list file name (should be all lowercase)
  var chapterName = "";
  
  // TODO: Update the folderId, archiveZipFolderId, and archiveFolderId with your folder IDs. See the README to understand the folder structure I used
  var folderId = "";
  var archiveZipFolderId = "";
  var archiveFolderId = "";
  var folder = DriveApp.getFolderById(folderId);
  var archiveZipFolder = DriveApp.getFolderById(archiveZipFolderId);
  var archiveFolder = DriveApp.getFolderById(archiveFolderId);

  // TODO: Update the sheetsId with the ID of the sheet you will use
  var sheetsId = "";
  var sheet = SpreadsheetApp.openById(sheetsId);

  const d = new Date();
  const lastWeek = new Date();
  lastWeek.setDate(d.getDate() - 7);

  // Take the attachment from the membership email and store it in Drive
  for (var i = 0; i < messages.length; i++) {
    var message = messages[i];
    
    if (message.getAttachments().length > 0) {
      var attachments = message.getAttachments();
      
      for (var j = 0; j < attachments.length; j++) {
        var attachment = attachments[j];
        var attachmentHash = Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, attachment.getBytes()));
        var existingFiles = folder.getFiles();
        var isDuplicate = false;

        while (existingFiles.hasNext()) {
          var existingFile = existingFiles.next();
          var existingFileHash = Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, existingFile.getBlob().getBytes()));

          if (attachmentHash === existingFileHash) {
            isDuplicate = true;
            break;
          }
        }

        if (!isDuplicate) {
          folder.createFile(attachment);
        }
      }
    }
  }

  Logger.log("Zip file moved from email to drive");

  var newFileName = chapterName + "_membership_list_"+d.toDateString()+".csv"
  var oldFileName = chapterName + "_membership_list_"+lastWeek.toDateString()+".csv"

  var file = folder.getFilesByName(chapterName + "_membership_list.zip").next();
  var fileBlob = file.getBlob();
  fileBlob.setContentType("application/zip");
  var unzippedFile = Utilities.unzip(fileBlob);
  var newFile = folder.createFile(unzippedFile[0]);
  
  newFile.setName(newFileName);
  Logger.log("CSV extracted");
  file.moveTo(archiveZipFolder);
  Logger.log("Zip file moved to archive");

  var oldFile = folder.getFilesByName(oldFileName);
  if (oldFile.hasNext()) {
    oldFile = oldFile.next();

    var csvOld = Utilities.parseCsv(oldFile.getBlob().getDataAsString());
    var csvNew = Utilities.parseCsv(newFile.getBlob().getDataAsString());

    // group both by email
    var newData = groupCsv(csvNew, "email");
    var oldData = groupCsv(csvOld, "email");

	// TODO: Confirm these tabs are in your Sheets file
    var newRecords = sheet.getSheetByName("Additions");
    var changedRecords = sheet.getSheetByName("Modifications");
    var deletedRecords = sheet.getSheetByName("Deletions");

    // iterate through new file
    Logger.log("Comparing last list with most recent...");
    for (let email in newData) {
      rowN = newData[email];
      rowN.pop();  // last col is the list date, don't want this in comparison
      if (email in oldData) {
        var rowO = oldData[email];
        rowO.pop();
		// remove the row from the old list so we know what remains at the end
        delete oldData[email];

        var colsToUpdate = [];
        for (var i = 0; i < rowN.length; i++) {
          if (i === 11) {  // xdate col, ignore because this updates frequently for monthly members
            continue;
          }
          // if the same cell contains different values, then a change occurred
          if (rowN[i] !== rowO[i]) {
            colsToUpdate.push(getColumnLetters(i+2));  // add 1 for letter number, then add 1 because we add a List Date col in front
          }
        }

        if (colsToUpdate.length !== 0) {
          changedRecords.appendRow([d.toDateString()].concat(rowN));
          var allVals = changedRecords.getRange("A1:A").getValues();
          var rowNum = allVals.filter(String).length;
          for (i = 0; i < colsToUpdate.length; i++) {
            var cell = colsToUpdate[i] + rowNum.toString();
            changedRecords.getRange(cell + ':' + cell).setBackground("#f9cb9c");  // just a nice orange
          }
        }
      } else {
        // if email isn't in old data, it's a new addition
        newRecords.appendRow([d.toDateString()].concat(rowN));
      }
    }

    for (let email in oldData) {  // any leftover rows were removed from the new file
      var row = oldData[email];
      deletedRecords.appendRow([d.toDateString()].concat(row));
    }

    Logger.log("Archiving the older list...");
    oldFile.moveTo(archiveFolder);
    Logger.log("Done");
  }
}