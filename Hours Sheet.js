var resEdit;

function onSubmit1(e) {
  // Below is a dirty fix for script permission error which resulted in undefined property 'response'
  // FormApp.getActiveForm();
  
  var response = e.response;
  var responses = response.getItemResponses();
  
  var mainSheet = getHoursSheet();
  var hoursSheet = mainSheet.getSheetByName(responses[0].getResponse());
  
  if (hoursSheet == null)
    emailError("Hours Sheet Error",
               "An employee's name was submitted that has not been entered into the template spreadsheet. (Ask Shalom for help.)");
  
  // Checking if this is an edit...
  var props = PropertiesService.getScriptProperties();
  var lastRes = props.getProperty("LAST_RESPONSE");
  if (response.getId() == lastRes)
  {
    resEdit = true;
    var sheets = mainSheet.getSheets();
    var sheet2Del = sheets[0];
    var lastRw = sheet2Del.getLastRow();
    var lastDate;
    if (lastRw > 1)
      lastDate = new Date(sheet2Del.getRange(lastRw, 1).getValue());
    else
      lastDate = null;
    
    for (var i = 1; i < sheets.length; i++)
    {
      lastRw = sheets[i].getLastRow();
      if (lastRw > 1 && (lastDate == null || (new Date(sheets[i].getRange(lastRw, 1).getValue())) > lastDate))
        sheet2Del = sheets[i];
    }
    sheet2Del.deleteRow(sheet2Del.getLastRow());
  }
  else
  {
    props.setProperty("LAST_RESPONSE", response.getId());
  }
  
  
  // Entering data into hours sheet...
  var tDate = new Date(new Date() - 25200000);
  var strVal = tDate.toDateString() + " ";
  strVal = strVal + (tDate.toLocaleTimeString()).substr(0, 11);
  var vals = [[strVal, responses[1].getResponse(), responses[2].getResponse(), null]];
  
  var hrsRange = hoursSheet.getRange(hoursSheet.getLastRow() + 1, 1, 1, 4).setValues(vals);
  hrsRange.getCell(1, 1).setHorizontalAlignment("left");
  var totalHrsCell = hrsRange.getCell(1, hrsRange.getLastColumn());
  totalHrsCell.setFormula('(' + hrsRange.getCell(1, 3).getA1Notation() + "-" + hrsRange.getCell(1, 2).getA1Notation() + ")*24").setNumberFormat("0.00");
  
  
  // Calling second submit method...
  onSubmit2(e);
}



// Fetches the hour sheet or creates a new one if 2 weeks have passed since last one
function getHoursSheet()
{
  var day = 86400000;
  var date = new Date(new Date() - 25200000);
  var daysBack = [3, 4, 5, 6, 0, 1, 2];
  date.setTime(date - daysBack[date.getDay()]*day);
  var dir = DriveApp.getFolderById("0B88HlOjbQh4rb2lVa1FtNjlTb1U");
  var iter;
  for (var i = 0; i < 2; i++)
  {
    iter = dir.searchFiles("title contains " + "'" + Number(date.getMonth()+1) + "/" + date.getDate() + "/" + date.getFullYear() + "'");
    if (iter.hasNext())
      break;
    else if(i == 1)
      iter = null;
    else
      date.setTime(date - 7*day);
  }
  
  if (iter != null)
  {
    // open existing file here
    var sheet = SpreadsheetApp.open(iter.next());
    return sheet;
  }
  else
  {
    // create new sheet here
    date = new Date(new Date() - 25200000);
    date.setTime(date - daysBack[date.getDay()]*day);
    
    var templateSheet = DriveApp.getFileById("1pMnR6G7p_W-SzFLeFpX4E8-mjfTY8E7hmjmU40nn8VU");
    var newSheet = templateSheet.makeCopy(Number(date.getMonth() + 1) + "/" + date.getDate() + "/" + date.getFullYear(), dir);
    return SpreadsheetApp.open(newSheet);
  }
}



// An email error message
function emailError(sub, msg, link)
{
  if (link != null)
    msg += "\nA link to the spreadsheet: " + link;
  GmailApp.sendEmail(["ljst789@gmail.com"], sub, msg);
}
