var log_counter = 2; //begin logging from the second row 

function main() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var cur_sheet = sheet.getSheetByName('Data');
  var range = cur_sheet.getDataRange(); 
  var ID_cell = 0;
  var status_cell = 4; 
  var input =  SpreadsheetApp.getActiveSpreadsheet().getRangeByName('INPUT').getValue();
  var start_date = SpreadsheetApp.getActiveSpreadsheet().getRangeByName('STARTDATE').getValue();
  var end_date = SpreadsheetApp.getActiveSpreadsheet().getRangeByName('ENDDATE').getValue();
  var MILLIS_PER_DAY = 1000 * 60 * 60 * 24;
  
  //formatting dates
  start_date = Utilities.formatDate(start_date, "GMT", "yyyy-MM-dd");
  end_date = Utilities.formatDate(end_date, "GMT", "yyyy-MM-dd");

  //clear up persistent global memory
  var stateChangeStat = {}
  
  //get direct children of input 
  var FILTER = 'issueFunction in linkedIssuesOfRecursive(\"issue = '+input+'\",\"is a Parent of\") AND issuetype = Topic';
  var childIssues = JIRA.createService('JIRA').list('story', FILTER);
  var inputs = [];
  var titles = [];
  for (var c = 0; c < childIssues.json_.length; c++) {
    inputs.push(childIssues.json_[c]['key']);
    titles.push(childIssues.json_[c]['fields']['summary']);
  }

  //timer-start
  var t0 = new Date().getTime();
  
  FILTER = '(project = PLAT) AND issueFunction in linkedIssuesOfRecursive(\"issue = '+input+'\", \"is a Parent of\") AND status changed DURING ( \"' + start_date + '\", \"' + end_date + '\")'; ;
  var values = JIRA.createService('JIRA').list('story', FILTER);
  
  for (var r=0; r < values.json_.length; r++) {

    var rt = ''
    var name = values.json_[r]['key']
    //SpreadsheetApp.getActiveSheet().getRange('G'+log_counter.toString()).setValue(name);
    //log_counter++;
    
     //get latest change status
    var new_status = determineStatusChange(name, start_date, end_date); 
   
    FILTER = '(issueFunction in linkedIssuesOfRecursive(\"issue = '+name+'\",\"is a Child of\") and issuetype = Topic) AND issueFunction in linkedIssuesOfRecursive(\"issue = '+input+'\", \"is a Parent of\") '; //search all upper node
    var parentIssues = JIRA.createService('JIRA').list('story', FILTER)
    
    for (var c = 0; c < parentIssues.json_.length; c++) {
     var key = parentIssues.json_[c]['key'];
     var parentIdx = -1;
     /*
     * Check if element in list 
     */ 
     for (var i = 0 ; i < inputs.length ; i++) {
       if (inputs[i] == key) {
         parentIdx = i;
         rt = titles[parentIdx];
         break;
       }
     }
    }
    if (rt.length > 0) {

    //updating global variable dict
    for(var j = 0 ; j < new_status.length ; j++){

      if (stateChangeStat[rt] && (new_status[j].length < 15)){
        // the ticket with no status outputs a long message
        if (stateChangeStat[rt][new_status[j]]) 
          stateChangeStat[rt][new_status[j]] +=1;
        else 
          stateChangeStat[rt][new_status[j]] = 1; 
      } else if(new_status[j].length < 15){

          stateChangeStat[rt] = {};
          stateChangeStat[rt][new_status[j]] = 1; 
        }
     }
    }
  }
  
  //write to outut cell
  cur_sheet = sheet.getSheetByName('Output');
 
  //clear content 
  cur_sheet.getRange('A2:H23').clearContent();
  
  //write dictionary to output cell and spreadsheet
  SpreadsheetApp.getActiveSheet().getRange('A8').setValue(stateChangeStat)
  var y = 2; 
  for (var z = 0 ; z < titles.length;  z++) {
    if (stateChangeStat[titles[z]]) {
    cur_sheet.getRange('A'+ y.toString()).setValue(titles[z]);
    cur_sheet.getRange('B'+ (y).toString()).setValue(stateChangeStat[titles[z]]['Unstarted']?stateChangeStat[titles[z]]['Unstarted']:0); 
    cur_sheet.getRange('C'+ (y).toString()).setValue(stateChangeStat[titles[z]]['Started']?stateChangeStat[titles[z]]['Started']:0); 
    cur_sheet.getRange('D'+ (y).toString()).setValue(stateChangeStat[titles[z]]['New Unstarted']?stateChangeStat[titles[z]]['New Unstarted']:0); 
    cur_sheet.getRange('E'+ (y).toString()).setValue(stateChangeStat[titles[z]]['New Started']?stateChangeStat[titles[z]]['New Started']:0); 
    cur_sheet.getRange('F'+ (y).toString()).setValue(stateChangeStat[titles[z]]['New Finished']?stateChangeStat[titles[z]]['New Finished']: 0 ); 
    cur_sheet.getRange('G'+ (y).toString()).setValue(stateChangeStat[titles[z]]['New Closed']?stateChangeStat[titles[z]]['New Closed']: 0 ); 
    cur_sheet.getRange('H'+ (y).toString()).setValue('=SUM('+'D'+ (y).toString() + ':G'+ (y).toString()+')'); 
    } else {
    cur_sheet.getRange('A'+ y.toString()).setValue(titles[z]);
    cur_sheet.getRange('B'+ (y).toString()).setValue(0); 
    cur_sheet.getRange('C'+ (y).toString()).setValue(0); 
    cur_sheet.getRange('D'+ (y).toString()).setValue(0); 
    cur_sheet.getRange('E'+ (y).toString()).setValue(0); 
    cur_sheet.getRange('F'+ (y).toString()).setValue(0); 
    cur_sheet.getRange('G'+ (y).toString()).setValue(0); 
    cur_sheet.getRange('H'+ (y).toString()).setValue('=SUM('+'D'+ (y).toString() + ':G'+ (y).toString()+')');  
    }
    y = y + 1;
  }
  
  y = y + 1; 
  cur_sheet.getRange('A'+ y.toString()).setValue("Totals:");
  cur_sheet.getRange('B'+ (y).toString()).setValue("=SUM(B2:B" + (y-2).toString() + ")"); 
  cur_sheet.getRange('C'+ (y).toString()).setValue("=SUM(C2:C" + (y-2).toString() + ")"); 
  cur_sheet.getRange('D'+ (y).toString()).setValue("=SUM(D2:D" + (y-2).toString() + ")"); 
  cur_sheet.getRange('E'+ (y).toString()).setValue("=SUM(E2:E" + (y-2).toString() + ")"); 
  cur_sheet.getRange('F'+ (y).toString()).setValue("=SUM(F2:F" + (y-2).toString() + ")"); 
  cur_sheet.getRange('G'+ (y).toString()).setValue("=SUM(G2:G" + (y-2).toString() + ")"); 
  cur_sheet.getRange('H'+ (y).toString()).setValue("=SUM(H2:H" + (y-2).toString() + ")");    
  
  //timer-end
  var t1 = new Date().getTime();
  
  SpreadsheetApp.getActiveSheet().getRange('A9').setValue("Process took " + (t1 - t0) + " milliseconds."); 
}


function determineStatusChange(ticket, from, to) {
  var ret = [];
  var last_change = findLatestStatusChange(ticket, from, to);

  if (last_change == 'Unstarted' || last_change == 'Started')
    ret.push(last_change);
   
  ret.push("New "+last_change); 
  
  //return issues.json_.length;
  i = 23; 
  return ret;
}

function findLatestStatusChange(ticket, from, to) {
  var MILLIS_PER_DAY = 1000 * 60 * 60 * 24;
  var statuses = ['Started', 'Unstarted', 'Finished', 'Closed']
  var c = 1; 

  //look for last change starting from 23:00 to 00:00
  for (var i = 23 ; i > 0 ; i = i - 5) {
    for (var j = 0; j < statuses.length ; j++) {
      to_h = to + " "  + i.toString() + ":00";
      from_h = to + " "  + (i-6).toString() + ":00";
      var FILTER = 'key = '+ ticket +' AND status changed to '+ statuses[j] +' DURING ( \"' + from_h + '\", \"' + to_h + '\")';  
      //SpreadsheetApp.getActiveSheet().getRange('C'+ i.toString()).setValue(FILTER);
      c+=1;
    //if (ticket == "PLAT-15352"){
    //  SpreadsheetApp.getActiveSheet().getRange('G'+log_counter.toString()).setValue("Not found To has value" + FILTER);
    //  log_counter++;
    //}
      issues = JIRA.createService('JIRA').list('story', FILTER)
      switch(issues.json_.length == 1) {
        case true: 
          return statuses[j]; 
          break; 
        case false: 
          break;
      } 
    }
  }
  
  // fall back one day
  to = new Date(new Date(to) - 1*MILLIS_PER_DAY)
  to = Utilities.formatDate(to, 'Etc/GMT', "yyyy-MM-dd")

  if (to >= from ) {
    return findLatestStatusChange(ticket, from, to); 
  }

  return 'status not changed in given time range!'; 
}

