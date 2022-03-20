const ss = SpreadsheetApp.getActiveSpreadsheet()
const formWS = ss.getSheetByName("DataEntry")
const dataWS = ss.getSheetByName("DataBase")
const podataWS = ss.getSheetByName("PO with TBA")

const fieldRange = ["B3","B4","B5","B6","B7","B8","B9","B10","B11","B12","B13","B14","B15","B16","B17","B18","B19","B20","E3","E4","E5","E6","E7","E8","E9","E10","E11","E12","E13","E14","E15","E16","E17","E18","E19","E20","H3","H4","H5","H6","H7","H8","H9","H10","H11","H12","H13","H14","H15","H16","H17","H18","H19","H20"]

const fieldValues = fieldRange.map(f => formWS.getRange(f).getValue())
const searchCell = formWS.getRange("B1")
const posearchCell = formWS.getRange("H1")


function newRecord() {
  fieldRange.forEach(f => formWS.getRange(f).clearContent())
  searchCell.clearContent()
  posearchCell.clearContent()
}


function createNewRecord() {
  dataWS.appendRow(fieldValues)
  ss.toast("New Diso Has Been Created")
}


function createNewPO() {
  podataWS.appendRow(fieldValues)
  ss.toast("New PO Has Been Created")
}


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


function editPOwithDispo() {
  const id = posearchCell.getValue()

  if(id == "") return
  const cellFound = podataWS.getRange("BC:BC").createTextFinder(id).matchEntireCell(true).findNext()
  if(!cellFound) return
  const row = cellFound.getRow()
  podataWS.getRange(row,1,1,fieldValues.length).setValues([fieldValues])
  posearchCell.clearContent()
  ss.toast("This PO Has Been Edited")
  // newRecord()
}



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



function search() {
  const searchValue = searchCell.getValue()
  const data = dataWS.getRange("A2:BD").getValues()
  const recordFound = data.filter(r => r[14] == searchValue)
  if(recordFound.length === 0) return
  fieldRange.forEach((f,i) => formWS.getRange(f).setValue(recordFound[0][i]))
}


function searchPO() {
  const searchValue = posearchCell.getValue()
  const data = podataWS.getRange("A2:BC").getValues()
  const recordFound = data.filter(r => r[54] == searchValue)
  if(recordFound.length === 0) return
  fieldRange.forEach((f,i) => formWS.getRange(f).setValue(recordFound[0][i]))
}

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
