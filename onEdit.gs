function onEdit(e) {
  if((e.range.getA1Notation() !== "B1") && (e.range.getA1Notation() !== "H1")) return
  if(e.source.getSheetName() !== "DataEntry") return
  if(e.range.getA1Notation() == "H1") {
    searchCell.clearContent()
    searchPO()
    return
  }
  if(e.range.getA1Notation() == "B1") {
    posearchCell.clearContent()
    search()
    return
  }
}
