const ss = SpreadsheetApp.getActiveSpreadsheet()
// const ss2 = SpreadsheetApp.openById("1zMi2RPL6K4cY-9kFo0FR5IZ0_QV5avZ2OPv0CACOT20")
const formWS = ss.getSheetByName("Calculation Sheet to Dispo R&D")
const dataWS = ss.getSheetByName("Dispo Data Base")
const printAreaWS = ss.getSheetByName("Print Area R&D")
const testWS = ss.getSheetByName("Test")

// Defining Constant for Search cell in Calculation Sheet
const searchDispoInCalculation = formWS.getRange("C1")
// Defining Constant for Search cell in Print Area Sheet
const searchDispoToPrint = printAreaWS.getRange("BK11")

// Field Range for Special Calculation Area
const spcalFieldRange = ["BT4","BT5","BT6","BT7","BT8","BT9","BT10","BT11","BT12","BT13","BT14","BT15","BT16","BT17","BT18","BT19","BT20","BT21","BT22","BT23","BT24","BT25","BT26","BV4","BV5","BV6","BV7","BV8","BV9","BX17","BX18","AL75","AL77"]

// Field Values for Special Calculation Area
const spcalFieldValues = spcalFieldRange.map(f => formWS.getRange(f).getValue())

// Field Range to paste Special Calculation Data while Dispo Creation
const spdataFieldRange = ["P2","P3","AJ3","AJ4","AJ5","AJ6","AJ7","AR2","AR3","AS3","AS4","AN3","AS5","AS6","AS7","AS8","AS9","AS10","AS11","BT33","BT34","AM8","AM9","BT27","BT28","BT29","BT30","BT31","BT32","BC22","AQ39","BV22","BV23"]


// Field Range for Calculation Area
const calfieldRange = ["E2","E4","E31","E32","E3","C61","D61","C62","D62","E6","E7","E5","E8","E9","E10","E11","E12","E13","E14","E23","E24","E33","E30","E29","E28","E27","E26","E25","E15","E16","E17","E18","P4","P5","P2","P3","Y9","Y10","Y11","Z2","Z3","Z4","Z5","Z6","AJ3","AJ4","AJ5","AJ6","E21","E22","AS5","AS6","BC22","AU22","P7","P8","AS4","AN3","AR2","AR3","AS3","BT27","BT28","BT29","BT30","BT31","BT32","AS7","AJ7","AS8","AQ39","AK39","Z6","AJ10","AJ11","AS9","AS10","AS11","M14","O14","S14","AB14","AD14","AG14","AJ14","Z14","M15","O15","S15","AB15","AD15","AG15","AJ15","Z15","M16","O16","S16","AB16","AD16","AG16","AJ16","Z16","M17","O17","S17","AB17","AD17","AG17","AJ17","Z17","M18","O18","S18","AB18","AD18","AG18","AJ18","Z18","M19","O19","S19","AB19","AD19","AG19","AJ19","Z19","M20","O20","S20","AB20","AD20","AG20","AJ20","Z20","M21","O21","S21","AB21","AD21","AG21","AJ21","Z21","M22","O22","S22","AB22","AD22","AG22","AJ22","Z22","M23","O23","S23","AB23","AD23","AG23","AJ23","Z23","M24","O24","S24","AB24","AD24","AG24","AJ24","Z24","M25","O25","S25","AB25","AD25","AG25","AJ25","Z25","M30","O30","S30","AB30","AD30","AG30","AJ30","Z30","M31","O31","S31","AB31","AD31","AG31","AJ31","Z31","M32","O32","S32","AB32","AD32","AG32","AJ32","Z32","M33","O33","S33","AB33","AD33","AG33","AJ33","Z33","M34","O34","S34","AB34","AD34","AG34","AJ34","Z34","M35","O35","S35","AB35","AD35","AG35","AJ35","Z35","M39","O39","S39","AB39","AD39","W39","Z39","M40","O40","S40","AB40","AD40","W40","Z40","M41","O41","S41","AB41","AD41","W41","Z41","M42","O42","S42","AB42","AD42","W42","Z42","M43","O43","S43","AB43","AD43","W43","Z43","M44","O44","S44","AB44","AD44","W44","Z44","M45","O45","S45","AB45","AD45","W45","Z45","M46","O46","S46","AB46","AD46","W46","Z46","A68","C68","E68","H68","K68","A69","C69","E69","H69","K69","A70","C70","E70","H70","K70","A71","C71","E71","H71","K71","S65","O65","V65","Y65","AF65","S66","O66","V66","Y66","AF66","S67","O67","V67","Y67","AF67","S68","O68","V68","Y68","AF68","O72","S72","V72","Y72","AA72","AC72","O73","S73","V73","Y73","AA73","AC73","O74","S74","V74","Y74","AA74","AC74","O75","S75","V75","Y75","AA75","AC75","O76","S76","V76","Y76","AA76","AC76","O77","S77","V77","Y77","AA77","AC77","AJ65","AN65","AP65","AR65","AJ66","AN66","AP66","AR66","AJ67","AN67","AP67","AR67","A54","A55","D54","D55","G54","G55","J54","J55","D58","D59","J64","E33","E45","E47","E42","E36","E38","E40","E35","E34","E46","E48","E43","E44","E37","E39","E41","Q49","Q50","Q51","Q52","Q53","Q54","Q55","Q56","Q57","Q58","Q59","Q60","Q61","Q62","A74","BV22","BV23","E19","I19","E20","I20","P9","P10","P11","Z7","Z8","AU24","AK24","P6","BT33","BT34","AM8","AM9","W14","W15","W16","W17","W18","W19","W20","W21","W22","W23","W24","W25","W30","W31","W32","W33","W34","W35","BH6","BI6","BJ6","BK6","BL6","BM6","BN6","BO6","BH7","BI7","BJ7","BK7","BL7","BM7","BN7","BO7","BH8","BI8","BJ8","BK8","BL8","BM8","BN8","BO8","BH9","BI9","BJ9","BK9","BL9","BM9","BN9","BO9","BH10","BI10","BJ10","BK10","BL10","BM10","BN10","BO10","BH11","BI11","BJ11","BK11","BL11","BM11","BN11","BO11","BH12","BI12","BJ12","BK12","BL12","BM12","BN12","BO12","BH13","BI13","BJ13","BK13","BL13","BM13","BN13","BO13","BH14","BI14","BJ14","BK14","BL14","BM14","BN14","BO14","BH15","BI15","BJ15","BK15","BL15","BM15","BN15","BO15","BH16","BI16","BJ16","BK16","BL16","BM16","BN16","BO16","BH17","BI17","BJ17","BK17","BL17","BM17","BN17","BO17","BH30","BI30","BJ30","BK30","BL30","BM30","BN30","BO30","BH31","BI31","BJ31","BK31","BL31","BM31","BN31","BO31","BH32","BI32","BJ32","BK32","BL32","BM32","BN32","BO32","BH33","BI33","BJ33","BK33","BL33","BM33","BN33","BO33","BH34","BI34","BJ34","BK34","BL34","BM34","BN34","BO34","BH35","BI35","BJ35","BK35","BL35","BM35","BN35","BO35","AK72","AM72","AO72","AP72","AQ72","AR72","AK73","AM73","AO73","AP73","AQ73","AR73","N14","N15","N16","N17","N18","N19","N20","N21","N22","N23","N24","N25","N30","N31","N32","N33","N34","N35","N39","N40","N41","N42","N43","N44","N45","N46"]

// Field Values for Whole Calculation Area (It is the main array to write the database)
const calFieldValues = calfieldRange.map(f => formWS.getRange(f).getValue())

function test (){
  const id = calFieldValues[27]
  const cellFound = dataWS.getRange("AB:AB").createTextFinder(id).matchEntireCell(true).findNext()
  if (cellFound){
    ss.toast("This Diso is already in Data Base. If you edited this dispo, please try 'Save as Edited Dispo' button.")
    return
  }
  // const row = cellFound.getRow()
  // console.log(row)
}
// Field Range for the Dispo Print area
const printfieldRange = ["AD20","AU20","AU21","AU22","K5","V5","AC5","P7","W7","J15","O15","J14","J16","J17","J18","J19","J20","J21","J22","J25","J26","I141","F23","R23","AD23","AD24","J27","J28","AE14","AE15","AE16","AE17","AU18","AY18","AU19","AY19","U9","U10","AI9","AZ28","AZ27","AZ29","AZ30","AX38","AU27","AU28","AU29","AU30","AC31","AM31","AV31","AC32","AM32","AW32","J30","N30","J31","J32","U30","U31","U32","A34","A35","AA34","AA35","N35","AN35","T39","AN38","AN39","AN40","AX39","AX38","AU56","AU78","AH42","AN42","AU42","F44","K44","S44","AA44","AE44","AP44","AK44","AU44","F45","K45","S45","AA45","AE45","AP45","AK45","AU45","F46","K46","S46","AA46","AE46","AP46","AK46","AU46","F47","K47","S47","AA47","AE47","AP47","AK47","AU47","F48","K48","S48","AA48","AE48","AP48","AK48","AU48","F49","K49","S49","AA49","AE49","AP49","AK49","AU49","F50","K50","S50","AA50","AE50","AP50","AK50","AU50","F51","K51","S51","AA51","AE51","AP51","AK51","AU51","F52","K52","S52","AA52","AE52","AP52","AK52","AU52","F53","K53","S53","AA53","AE53","AP53","AK53","AU53","F54","K54","S54","AA54","AE54","AP54","AK54","AU54","F55","K55","S55","AA55","AE55","AP55","AK55","AU55","F60","K60","S60","AA60","AE60","AP60","AK60","AU60","F61","K61","S61","AA61","AE61","AP61","AK61","AU61","F62","K62","S62","AA62","AE62","AP62","AK62","AU62","F63","K63","S63","AA63","AE63","AP63","AK63","AU63","F64","K64","S64","AA64","AE64","AP64","AK64","AU64","F65","K65","S65","AA65","AE65","AP65","AK65","AU65","F70","K70","S70","AA70","AE70","AK70","AP70","F71","K71","S71","AA71","AE71","AK71","AP71","F72","K72","S72","AA72","AE72","AK72","AP72","F73","K73","S73","AA73","AE73","AK73","AP73","F74","K74","S74","AA74","AE74","AK74","AP74","F75","K75","S75","AA75","AE75","AK75","AP75","F76","K76","S76","AA76","AE76","AK76","AP76","F77","K77","S77","AA77","AE77","AK77","AP77","AS88","N88","I88","AM88","AE88","AS89","N89","I89","AM89","AE89","AS90","N90","I90","AM90","AE90","AS91","N91","I91","AM91","AE91","E108","K108","S108","W108","AM108","E109","K109","S109","W109","AM109","E110","K110","S110","W110","AM110","E111","K111","S111","W111","AM111","K96","F96","S96","AA96","AE96","AK96","K97","F97","S97","AA97","AE97","AK97","K98","F98","S98","AA98","AE98","AK98","K99","F99","S99","AA99","AE99","AK99","K100","F100","S100","AA100","AE100","AK100","K101","F101","S101","AA101","AE101","AK101","M116","X116","AH116","AX116","M120","X120","AH120","AX120","M124","X124","AH124","AX124","C129","J129","C130","P130","AC129","AJ129","AC130","AP130","A135","AM135","AT152","I141","P142","P144","P145","P146","P147","P148","AK140","AK141","AR142","AR144","AJ145","AU145","AR146","AR147","AR148","A1","A159","A160","A161","A164","A163","A162","A158","A150","A165","A166","AL153","AL154","AL155","A41","K57","AE57"]


// Function for Special Calculation
function specialCalculation(){
  spdataFieldRange.forEach((f,i) => formWS.getRange(f).setValue(spcalFieldValues[i]))
}

// Function for Yarn Calculation
function yarnCalculation (){
  // Pasting the collected calculated Warp Dyed Yarn Data into Designated Area
  formWS.getRange("AB14:AB25").setValues(formWS.getRange("AM14:AM25").getValues())

  // Pasting the collected calculated Warp Grey Yarn Data into Designated Area
  formWS.getRange("AD14:AD25").setValues(formWS.getRange("AT14:AT25").getValues())

  // Pasting the collected calculated Warp No. of Cone Data into Designated Area
  formWS.getRange("AG14:AG25").setValues(formWS.getRange("AN14:AN25").getValues())

  // Pasting the collected calculated Warp Cone Length Data into Designated Area
  formWS.getRange("AJ14:AJ25").setValues(formWS.getRange("AO14:AO25").getValues())

  // Pasting the collected calculated Seer Sucker Dyed Yarn Data into Designated Area
  formWS.getRange("AB30:AB35").setValues(formWS.getRange("AM30:AM35").getValues())

  // Pasting the collected calculated Seer Sucker Grey Yarn Data into Designated Area
  formWS.getRange("AD30:AD35").setValues(formWS.getRange("AT30:AT35").getValues())

  // Pasting the collected calculated Seer Sucker No. Of Cone Data into Designated Area
  formWS.getRange("AG30:AG35").setValues(formWS.getRange("AN30:AN35").getValues())

  // Pasting the collected calculated Seer Sucker Cone Length Data into Designated Area
  formWS.getRange("AJ30:AJ35").setValues(formWS.getRange("AO30:AO35").getValues())

  // Pasting the collected calculated Weft Dyed Yarn Data into Designated Area
  formWS.getRange("AB39:AB46").setValues(formWS.getRange("AG39:AG46").getValues())

  // Pasting the collected calculated Weft Grey Yarn Data into Designated Area
  formWS.getRange("AD39:AD46").setValues(formWS.getRange("AH39:AH46").getValues())
}

// Function For clearing calculation sheet and search cells.
function clearField() {
  calfieldRange.forEach(f => formWS.getRange(f).clearContent())
  searchDispoInCalculation.clearContent()
  searchDispoToPrint.clearContent()
}


// Function for creating new Dispo
function createNewRecord() {
  const id = calFieldValues[27]
  const cellFound = dataWS.getRange("AB:AB").createTextFinder(id).matchEntireCell(true).findNext()
  if (cellFound){
    ss.toast("This Diso is already in Data Base. If you edited this dispo, please try 'Save as Edited Dispo' button.")
    return
  }
  dataWS.appendRow(calFieldValues)
  ss.toast("New Diso Has Been Created")
}

// Function for Editing Searched Dispo
function editDispo() {
  const id = searchDispoInCalculation.getValue()
  if(id == ""){
    const dispo = calFieldValues[27]
    const cellFound = dataWS.getRange("AB:AB").createTextFinder(dispo).matchEntireCell(true).findNext()
  if (!cellFound){
    ss.toast("This Diso doesn't exit in Data Base. Please save the Dispo first")
    return
  }
  const row = cellFound.getRow()
  dataWS.getRange(row,1,1,calFieldValues.length).setValues([calFieldValues])
  // searchDispoInCalculation.clearContent()
  ss.toast("This Diso Has Been Edited")
  return
  }
  const cellFound = dataWS.getRange("AB:AB").createTextFinder(id).matchEntireCell(true).findNext()
  if(!cellFound) return
  const row = cellFound.getRow()
  dataWS.getRange(row,1,1,calFieldValues.length).setValues([calFieldValues])
  // searchDispoInCalculation.clearContent()
  ss.toast("This Diso Has Been Edited")
}

// Function for Searching Dispo in Calculation Sheet
function searchInCalculation() {
  const searchValue = searchDispoInCalculation.getValue()
  const data = dataWS.getRange("A3:XB").getValues()
  const recordFound = data.filter(r => r[27] == searchValue)
  if(recordFound.length === 0) return
  calfieldRange.forEach((f,i) => formWS.getRange(f).setValue(recordFound[0][i]))
}

// Function for Searching Dispo in Print Area Sheet
function searchInPrintArea() {
  const searchValue = searchDispoToPrint.getValue()
  const data = dataWS.getRange("A3:OT").getValues()
  const recordFound = data.filter(r => r[27] == searchValue)
  if(recordFound.length === 0) return
  printfieldRange.forEach((f,i) => printAreaWS.getRange(f).setValue(recordFound[0][i]))
}


function deleteRecord() {
  const id = searchDispoInCalculation.getValue()
  if(id == "") return
  const cellFound = dataWS.getRange("AB:AB").createTextFinder(id).matchEntireCell(true).findNext()
  if(!cellFound) return
  const row = cellFound.getRow()
  dataWS.deleteRow(row)
  ss.toast("This Diso Has Been Deleted")
  newRecord()
}
