function SheetUI() {
  SpreadsheetApp.getUi()
    .createMenu("Wizard")
    .addItem("Upload", "openUploadDialog")
    .addToUi();
}

function openUploadDialog() {
  var html = HtmlService.createHtmlOutputFromFile("upload");
  SpreadsheetApp.getUi().showModalDialog(html, "Upload File to Images folder");
}

function uploadFile(data, type, name) {
  const thisFileId = SpreadsheetApp.getActive().getId();
  const thisFile = DriveApp.getFileById(thisFileId);

  const parents = thisFile.getParents();
  let f;
  while (parents.hasNext()) {
    f = parents.next(); //f is a single 2020 folder.
  }
  const imagesFolders = f.getFoldersByName("Images");
  let imageFolder;
  if (!imagesFolders.hasNext()) {
    imageFolder = f.createFolder("Images");
  } else {
    imageFolder = imagesFolders.next();
  }
  const imageBlob = Utilities.newBlob(
    Utilities.base64Decode(data.split(",")[1]),
    type
  );
  const fileName = SpreadsheetApp.getActiveSheet().getActiveCell().getValue();
  imageBlob.setName(fileName);
  const newFile = imageFolder.createFile(imageBlob);
}
