function doGet(){
  // set a template html code instance on Index.html and create it as a HTML service
  return HtmlService.createTemplateFromFile("Index").evaluate();
}
function include(filename){
  //allow include statement to define javascript file in different file and  output it on the instance of Index.html
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

//Fetch list of Agents in a Spreadsheet
function fetchAgents(){
  //get agent spreadsheet instance
  var ss=SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/1vt5LDudAYKvTmVRIH0TkLxf-XQwwhxKWO6VSOsCrdIo/edit#gid=0');
  var s=ss.getSheetByName('Agent');// get data on sheet name Agent
  allValues=s.getRange('A2:A'+s.getLastRow()).getValues();//getting scope or range of values
  return allValues;
}
//myFunction handles duplicates in spreadsheet
function myFunction(req) {
  var row=0,len=0;
  //get ticket spreadsheet instance 
  var ss = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/1KL0obpDOXx9hfDASmanPIPCduhFXn4p52TISkhcUq3w/edit?resourcekey#gid=922390028');
  var month= Utilities.formatDate(new Date, Session.getScriptTimeZone(), "MMM");
  var datetime=Utilities.formatDate(new Date(), "GMT", "MM-dd-yyyy");
  //determine what sheet name to use
      s = ss.getSheetByName('Form responses 1');
      //append data or put data in last row of excel
      s.appendRow([
        new Date(),
        req.agents,
        req.ticket,
        month,
        datetime
      ])
      //============================
      lastRow = s.getLastRow();//get last row number 
      D_lastRow=lastRow-1;//last row numerical number zero based
      lastValues = s.getRange('A'+lastRow+':C'+lastRow).getValues();//get sample range of values
      name = lastValues[0][1];//get last row value of name 
      ticket=lastValues[0][2];//get last row value in ticket number
      allNames = s.getRange('B2:B'+D_lastRow).getValues();//get all values above lastrow
      allTickets=s.getRange('C2:C'+D_lastRow).getValues();//get all values above last row
      console.log(allNames);
  // TRY AND FIND EXISTING NAME
  var response="Record Saved..."//set default response as valid
  if(lastRow>2){
  for (row = 0; row < allNames.length; row++){
    if (allNames[row][0] == req.agents && allTickets[row][0]==req.ticket) {
      // compare all values above newly inputted value if inputted name and ticket is same for a certain row it is duplicated then delete last row else leave it there. 
      response="Record has a Duplicate...";//set default response as duplicate
      s.deleteRow(lastRow);
      break;
      }
  }
}
  return response;
}
