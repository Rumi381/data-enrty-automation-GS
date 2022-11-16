function onEdit(e) {
  if((e.range.getA1Notation() !== "C1") && (e.range.getA1Notation() !== "BK11")) return
  if((e.source.getSheetName() !== "Calculation Sheet to Dispo R&D") && (e.source.getSheetName() !== "Print Area R&D")) return
  if((e.range.getA1Notation() == "C1") && (e.source.getSheetName() == "Calculation Sheet to Dispo R&D")) {
    // searchCell.clearContent()
    searchInCalculation()
    return
  }
  if((e.range.getA1Notation() == "BK11") && (e.source.getSheetName() == "Print Area R&D")) {
    // posearchCell.clearContent()
    searchInPrintArea()
    return
  }
}
