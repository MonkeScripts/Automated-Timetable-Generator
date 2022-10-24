var Headers =[["WHO","NEXT STEP","DAYS","START","END"]]
var Steps =["Job Start","Creative Development","Creative Internal Review & Revision","Development","Internal review and Staging","DISKOUT"]
var ExtraDaysDistribution=[]
var timezone = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
var CurrentDate = Utilities.formatDate(new Date(), SpreadsheetApp.getActive().getSpreadsheetTimeZone(), "EEE, d MMM yyyy HH:mm:ss")
function generateTimeline() {

  var TIMELINEGENERATOR = SpreadsheetApp.getActiveSpreadsheet()
  var Prompt = SpreadsheetApp.getActiveSheet()
  var DataEntered = Prompt.getRange("B1:B7").getValues()//[column][row]
  var ProjectName = DataEntered[0][0] + " Timeline Generated at " + CurrentDate
  var StartDate = DataEntered[1][0]
  var EndDate = DataEntered[2][0]
  var RevCreative = DataEntered[3][0]
  var RevTech = DataEntered[4][0]
  var Content = DataEntered[5][0]
  var Stage = DataEntered[6][0]
  //Generates new timeline sheet
  TIMELINEGENERATOR.insertSheet(ProjectName)
  var Template = SpreadsheetApp.getActiveSheet()
  //Insert prompt inputs
  var PromptInputs =[[ProjectName,""],["StartDate",StartDate],["EndDate",EndDate],["NetWorkDays",""]] 
  Template.getRange("A1:B4").setValues(PromptInputs)
  Template.getRange("B4").setFormula('=NETWORKDAYS(B2,B3,HOLIDAYS!A2:A100)')
  //Insert Headers
  Template.getRange("A7:E7").setValues(Headers)
// Checks number of revisions
  //Creative
  var CreaIndex = Steps.indexOf("Creative Development")
  switch(RevCreative){
    case 0:
        if (Content == true){
        Steps.splice(1,0,"Content Framework","Content Internal Review & Revision")
        }
      break
    case 1:
      if (Content == true){
        Steps.splice(1,0,"Content Framework","Content Internal Review & Revision")
        Steps.splice(CreaIndex+4,0,"Client Presentation Final","Creative Client Feedback Final")
      }
      else{
        Steps.splice(CreaIndex+2,0,"Client Presentation Final","Creative Client Feedback Final")
      }
      break
    case 2:
      if (Content == true){
        Steps.splice(1,0,"Content Framework","Content Internal Review & Revision","Client Presentation R1","Creative Client Feedback R1")
        Steps.splice(CreaIndex+5,0,"Client Presentation Final","Creative Client Feedback Final")
      }
      else{
        Steps.splice(CreaIndex+2,0,"Client Presentation R1","Creative Client Feedback R1","Creative Revision Final")
        Steps.splice(CreaIndex+5,0,"Client Presentation Final","Creative Client Feedback Final")
      }
      break
    case 3:
    if (Content == true){
      Steps.splice(1,0,"Content Framework","Content Internal Review & Revision","Client Presentation R1","Creative Client Feedback R1")
      Steps.splice(CreaIndex+5,0,"Client Presentation R2","Creative Client Feedback R2","Creative Revision Final")
      Steps.splice(CreaIndex+8,0,"Client Presentation Final","Creative Client Feedback Final")
    }
    else{
      Steps.splice(CreaIndex+2,0,"Client Presentation R1","Creative Client Feedback R1","Creative Revision R1")
      Steps.splice(CreaIndex+5,0,"Client Presentation R2","Creative Client Feedback R2","Creative Revision Final")
      Steps.splice(CreaIndex+8,0,"Client Presentation Final","Creative Client Feedback Final")
    }
    break
  }
  //Tech
  var devIndex = Steps.indexOf("Development")
  switch(RevTech){
    case 0:
      break
    case 1:
      Steps.splice(devIndex+2,0,"Client Presentation Final","Tech Client Feedback Final","Tech Revision Final")
      break
    case 2:
      Steps.splice(devIndex+2,0,"Client Presentation R1","Tech Client Feedback R1","Tech Revision R1")
      Steps.splice(devIndex+5,0,"Client Presentation Final","Tech Client Feedback Final","Tech Revision Final")
      break
    case 3:
      Steps.splice(devIndex+2,0,"Client Presentation R1","Tech Client Feedback R1","Tech Revision R1")
      Steps.splice(devIndex+5,0,"Client Presentation R2","Tech Client Feedback R2","Tech Revision R2")
      Steps.splice(devIndex+8,0,"Client Presentation Final","Tech Client Feedback Final","Tech Revision Final")
      break
  }
  // Drop down list, remove steps that have passed
  var Start = Steps.indexOf(Stage)
  Steps= Steps.slice(Start,Steps.length)
  // Add WHO based on steps 
  var WHO = WhoArray()
  // Add WHO and Steps into table
  for(var i=0;i<=Steps.length;i++){
    var WHORange = Template.getRange("A8").offset(i,0)
    var StepsRange = Template.getRange("B8").offset(i,0)
    var ColourRange = Template.getRange("A8:E8").offset(i,0)
    WHORange.setValue(WHO[i])
    StepsRange.setValue(Steps[i])
    // Add coloring 
    switch(WHORange.getValue()){
      case "SVC/Creative":
        ColourRange.setBackgroundRGB(255, 153, 255)
        break
      case "Creative":
        ColourRange.setBackgroundRGB(255, 153, 255)
        break
      case "SVC/Client":
        ColourRange.setBackgroundRGB(153, 255, 153)
        break
      case "Client":
        ColourRange.setBackgroundRGB(153, 255, 153)
        break
      case "Tech":
        ColourRange.setBackgroundRGB(0, 204, 255)
        break
      case "END":
        ColourRange.setBackgroundRGB(204, 153, 255)
        break
    }
  }
  // Add date 
  var DaysAvailable = Template.getRange("B4").getValue()
  var RemainingDays = CalculateDays(DaysAvailable)
  AddExtraDays(RemainingDays)
  // Add start date in the first cell
  Template.getRange("D8").setFormula('=B2')
  // Add rest of dates in table
  for(var i=0;i<Steps.length;i++){
    if(i<Steps.length-1){
      Template.getRange("D8").offset(i+1,0).setFormulaR1C1('=WORKDAY(R[-1]C,IF((R[-1]C[1]="EOD"),ROUNDUP(R[-1]C[-1]),ROUNDDOWN(R[-1]C[-1])),HOLIDAYS!R2C1:R100C1)')
    }
    Template.getRange("E8").offset(i,0).setFormulaR1C1('=IF((MOD(RC[-2],1))=0.5,IF(R[-1]C="MID","EOD","MID"),IF(R[-1]C="MID","MID","EOD"))');
  }
  // Add function for end date
  Template.getRange("B3").setFormula("=OFFSET(D8,"+(Steps.length-1)+",0)")
  //Auto resize steps column
  Template.autoResizeColumns(2, 1);
}

function WhoArray(){
  var dict={
    // creative 
    "Job Start" : "Creative",
    "Content Framework" : "Creative",
    "Content Internal Review & Revision" : "SVC/Creative",
    "Creative Internal Review & Revision" : "SVC/Creative",
    "Creative Revision R1" : "Creative",
    "Creative Revision Final": "Creative",
    "Creative Revision" : "Creative",
    "Creative Development" : "Creative",
    //SVC/Client 
    "Client Presentation R1" : "SVC/Client",
    "Client Presentation R2" : "SVC/Client",
    "Client Presentation Final" : "SVC/Client",
    "Creative Client Feedback R1" : "Client",
    "Creative Client Feedback R2" : "Client",
    "Creative Client Feedback Final": "Client",
    "Tech Client Feedback R1" : "Client",
    "Tech Client Feedback R2" : "Client",
    "Tech Client Feedback Final" : "Client",
    // Tech
    "Development" : "Tech",
    "Tech Revision R1" : "Tech",
    "Tech Revision R2" : "Tech",
    "Tech Revision Final" : "Tech",
    "Internal review and Staging" : "Tech",
    "DISKOUT" : "END" 
  }
  var newPerson=[]
  for(var i=0;i<Steps.length;i++){
    newPerson.push(dict[Steps[i]])
  }
  return(newPerson)
}
// Adds 0.5 to every step
function CalculateDays(DaysAvailable){
  var Template = SpreadsheetApp.getActiveSheet()
  Counter = 0
  Days = 0.5
  while(Counter != Steps.length){
    var StepsRange = Template.getRange("B8").offset(Counter,0)
    var DateRange = Template.getRange("C8").offset(Counter,0)
    if(DaysAvailable>=0){
      if(StepsRange.getValue()=="Job Start" || StepsRange.getValue().substring(0,19)=="Client Presentation"){
      DateRange.setValue(0)
      Counter +=1
      }
      else{
      DateRange.setValue(Days)
      DaysAvailable -= 0.5
      Counter += 1
      }
    }
    else{
      addImage()
      break
    }
  }
  return(DaysAvailable)
}
  
function AddExtraDays(RemainingDays){
  var Template = SpreadsheetApp.getActiveSheet()
  var PriorityList=["Development","Creative Development","Content Framework","Content Internal Review & Revision","Tech Revision R1","Creative Internal Review & Revision","Tech Revision R2","Creative Revision R1","Creative Revision Final","Internal review and Staging","Tech Revision Final","Creative Client Feedback R1","Tech Client Feedback R1","Creative Client Feedback R2","Tech Client Feedback R2","Creative Client Feedback Final","Tech Client Feedback Final","DISKOUT"]
  //Filters list,removing indexes not in step
  var X
  var FilteredList = PriorityList.filter(X => Steps.includes(X))

  // generate multiplier 
  var Multiplier =0
  if (RemainingDays <=5){
    Multiplier =0.5
  }
  else{
    Multiplier = 0.25
  }
  DistributeDays(RemainingDays,Multiplier)
  //special case when client give too much time
  if(FilteredList.length<ExtraDaysDistribution.length){
    Multiplier = 0.6
    ExtraDaysDistribution=[]
    DistributeDays(RemainingDays,Multiplier)
  }
  // i is index of extradays array
  // j is index of filtered prio list
  var i =0
  var j =0
  var DayDict={}
  // Match priorityList to days
  while(j<FilteredList.length){
    DayDict[FilteredList[j]] = ExtraDaysDistribution[i]
    i+=1
    j+=1
  }
  // Add days 
  for(var k=0;k<Steps.length;k++){
    var Stage = Template.getRange("B8").offset(k,0).getValue()
    if(FilteredList.includes(Stage)){
      var Content = Template.getRange("C8").offset(k,0).getValue()
      if(DayDict[Stage]!= null){
        Template.getRange("C8").offset(k,0).setValue(Content + DayDict[Stage])
      }
    }
  }
}
// recursion multiplier
function DistributeDays(Total,Multiplier){
  if (Total <=0){
    return
  }
  ExtraDays = SpecialRound(Total*Multiplier,0.5)
  ExtraDaysDistribution.push(ExtraDays)
  Total -= ExtraDays
  return DistributeDays(Total,Multiplier)
}

function SpecialRound(value, step) {
    var inv = 1.0 / step
    if ((Math.round(value*inv)/inv) <= 0 ){
      return 0.5
    }
    else{
      return (Math.round(value * inv) / inv)
    }
}
function addImage() {
  var html = `
  <center><img src="https://media2.giphy.com/media/vnOuXywbQ73Gw/giphy.gif?cid=ecf05e47tsgn224psnm5h4n3qz69072yfnz3htznsyl28zjz&rid=giphy.gif&ct=g" /></center>
  <p class="body" style="font-family: sans-serif;font-size:2rem; color:black; text-align:center">
    You do not have enough days!! :(
  </p>

  `
  var htmlOutput = HtmlService
      .createHtmlOutput(html)
      .setWidth(700)
      .setHeight(600);

  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'BIG SAD :(');
}


