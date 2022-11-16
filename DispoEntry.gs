const ss = SpreadsheetApp.getActiveSpreadsheet()
const formWS = ss.getSheetByName("DataEntry")
const dataWS = ss.getSheetByName("DataBase")
const podataWS = ss.getSheetByName("PO with TBA")

// If the Database is located on other Spreadsheet(For example: Here the database is located in a spreadsheet of id:"1yxbHRhfDj8DwA28OE0_grwUGldGwisvqrzWv9OjR7M8") like a central Database.

// const ss2 = SpreadsheetApp.openById("1yxbHRhfDj8DwA28OE0_grwUGldGwisvqrzWv9OjR7M8")
// const dataWS = ss2.getSheetByName("DataBase")

// Defining fieldrange for Data Entry Sheet
const fieldRange = ["B3","B4","B5","B6","B7","B8","B9","B10","B11","B12","B13","B14","B15","B16","B17","B18","B19","B20","E3","E4","E5","E6","E7","E8","E9","E10","E11","E12","E13","E14","E15","E16","E17","E18","E19","E20","H3","H4","H5","H6","H7","H8","H9","H10","H11","H12","H13","H14","H15","H16","H17","H18","H19","H20"]

// Defining the Global variables
const fieldValues = fieldRange.map(f => formWS.getRange(f).getValue())
const searchCell = formWS.getRange("B1")
const posearchCell = formWS.getRange("H1")

// Function to add functionalities in the menubar
function onOpen(){
  const ui = SpreadsheetApp.getUi()
  const menu = ui.createMenu("Data Entry")
  menu.addItem("Search Dipo", "search")
  menu.addItem("Search PO", "searchPO")
  menu.addSeparator()
  menu.addItem("Clear", "clear")
  menu.addSeparator()
  menu.addItem("Save PO", "savePO")
  menu.addItem("Save Dispo", "saveRecord")
  menu.addSeparator()
  menu.addItem("Delete", "deleteRecord")
  menu.addToUi()
}

// Function to clear the fields of Data entry sheet
function clear() {
  fieldRange.forEach(f => formWS.getRange(f).clearContent())
  searchCell.clearContent()
  posearchCell.clearContent()
}

// Function to create a new record (Here a Dispo)
function createNewRecord() {
  // Defining variable to find the existing Dispo in the Dispo Database
  const id = fieldValues[14]
  const cellFound = dataWS.getRange("O:O").createTextFinder(id).matchEntireCell(true).findNext()
  if (cellFound){
    ss.toast("This Dispo is already in Data Base.")
    return
  }
  dataWS.appendRow(fieldValues)
  ss.toast("New Diso Has Been Created")
}


// Function to create a new record (Here a PO)
function createNewPO() {
  // Defining variable to find the existing PO in the Search Colimn of the PO Database
  const id = [fieldValues[0]," -T.Number:",fieldValues[13]," -Color:",fieldValues[8]," -PO Qty(Yds):",fieldValues[20]].join("").toUpperCase()
  const cellFound = podataWS.getRange("BC:BC").createTextFinder(id).matchEntireCell(true).findNext()
  if (cellFound){
    ss.toast("This PO is already in Data Base.")
    return
  }
  podataWS.appendRow(fieldValues)
  ss.toast("New PO Has Been Created")
}


// Function for saving or editing a record (Here a PO)
function savePO() {
  const id = posearchCell.getValue()

  if(id == "") {
    createNewPO()
    return
  }
  const cellFound = podataWS.getRange("BC:BC").createTextFinder(id).matchEntireCell(true).findNext()
  if(!cellFound) return
  const row = cellFound.getRow()
  podataWS.getRange(row,1,1,fieldValues.length).setValues([fieldValues])
  posearchCell.clearContent()
  ss.toast("This PO Has Been Edited")
  newRecord()
}

// Helper function to edit the PO database as well when a Dispo is finally created from the PO data.
function editPOwithDispo() {
  const id = posearchCell.getValue()
  if(id == "") return

  const dispo = fieldValues[14]
  const cellFoundwithDispo = podataWS.getRange("O:O").createTextFinder(dispo).matchEntireCell(true).findNext()
  if(cellFoundwithDispo) return

  const cellFound = podataWS.getRange("BC:BC").createTextFinder(id).matchEntireCell(true).findNext()
  if(!cellFound) return
  const row = cellFound.getRow()
  podataWS.getRange(row,1,1,fieldValues.length).setValues([fieldValues])
  posearchCell.clearContent()
  ss.toast("This PO Has Been Edited")
  // newRecord()
}


// Function for saving or editing a record (Here a Dispo)
function saveRecord() {
  const id = searchCell.getValue()

  if(id == "") {
    createNewRecord()
    editPOwithDispo()
    newRecord()
    return
  }
  const cellFound = dataWS.getRange("O:O").createTextFinder(id).matchEntireCell(true).findNext()
  if(!cellFound) return
  const row = cellFound.getRow()
  dataWS.getRange(row,1,1,fieldValues.length).setValues([fieldValues])
  searchCell.clearContent()
  ss.toast("This Diso Has Been Edited")
  newRecord()
}


// Function for searching a Dispo from the Dispo database
function search() {
  const searchValue = searchCell.getValue()
  const data = dataWS.getRange("A2:BB").getValues()
  const recordFound = data.filter(r => r[14] == searchValue)
  if(recordFound.length === 0) return
  fieldRange.forEach((f,i) => formWS.getRange(f).setValue(recordFound[0][i]))
}


// Function for searching a PO from the PO database
function searchPO() {
  const searchValue = posearchCell.getValue()
  const data = podataWS.getRange("A2:BC").getValues()
  const recordFound = data.filter(r => r[54] == searchValue)
  if(recordFound.length === 0) return
  fieldRange.forEach((f,i) => formWS.getRange(f).setValue(recordFound[0][i]))
}

// Function for deleting a Dispo from the Dispo database
function deleteRecord() {
  const id = searchCell.getValue()
  if(id == "") return
  const cellFound = dataWS.getRange("O:O").createTextFinder(id).matchEntireCell(true).findNext()
  if(!cellFound) return
  const row = cellFound.getRow()
  dataWS.deleteRow(row)
  ss.toast("This Diso Has Been Deleted")
  newRecord()
}
