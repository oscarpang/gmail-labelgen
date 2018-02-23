// Simple script to parse mailbox and generate stats to google spreadsheet
// Have to live with a spreadsheet.

var MAX_ITER = 100;

function parseEmail(from) {
    // Parse email address and count them
    from = from.match(/[a-zA-Z0-9.!#$%&'*+/=?^_`{|}~-]+@[a-zA-Z0-9](?:[a-zA-Z0-9-]{0,61}[a-zA-Z0-9])?(?:\.[a-zA-Z0-9](?:[a-zA-Z0-9-]{0,61}[a-zA-Z0-9])?)*/g);
    var email = "";
    if ( from && from.length ) {            
      email = from[0];
      email = email.replace(">", "");
      email = email.replace("<", "");
    }
    return email;
}

function parseDomain(email_addr) {
    var domain = email_addr.replace(/.*@/, "");
    return domain;
}

// Sweep through inbox and create stat
function processInbox_(sheet, addr_sheet, domain_sheet) {
  var threads = GmailApp.getInboxThreads();

  var i = 0;
  var from_dict = {};
  var addr_dict = {};
  var domain_dict = {};
  // Process each one in turn, assuming there's only a single
  // message in each thread
  for (var t in threads) {
    var thread = threads[t];

    var from = thread.getMessages()[0].getFrom();
    if(!(from in from_dict)) {
      from_dict[from]=1;
    } else {
      from_dict[from]++;
    }

    var email = parseEmail(from);
    if(!(email in addr_dict)) {
      addr_dict[email] = 1;
    } else {
      addr_dict[email]++;
    }
    
    var domain = parseDomain(email);
     if(!(domain in domain_dict)) {
      domain_dict[domain] = 1;
    } else {
      domain_dict[domain]++;
    }
    
    
    i = i+1;
    if (i > MAX_ITER) {
      break;
    }
  }
  
  for (var key in from_dict) {
    sheet.appendRow([key, from_dict[key]]);
  }
  
  for (var key in addr_dict) {
    addr_sheet.appendRow([key, addr_dict[key]]);
  }
    
  for(var key in domain_dict) {
    domain_sheet.appendRow([key, domain_dict[key]]);
  }
  sheet.sort(2,false);
  addr_sheet.sort(2, false);
  domain_sheet.sort(2, false);
}

function main_deleteAllLabels() {
  var labels = GmailApp.getUserLabels();
  for(var i = 0; i < labels.length; i++) {
    GmailApp.deleteLabel(labels[i]);
  }
}

// main()
// Starter function; from be scheduled regularly
function main_emailDatafromSpreadsheet() {
  // Get the active spreadsheet and make sure the first
  // sheet is the active one
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.setActiveSheet(ss.getSheets()[0]);
  
  var from = ss.getSheetByName("From");
  var from_addr = ss.getSheetByName("From_Addr");
  var domain_addr = ss.getSheetByName("Domain");
  
  
  if(from != null){
    from = from.clear();
  } else {
    from = ss.insertSheet("From");
  }


  if(from_addr != null){
    from_addr = from_addr.clear();
  } else {
    from_addr = ss.insertSheet("From_Addr");
  }
    
  if(domain_addr != null) {
    domain_addr = domain_addr.clear();
  } else {
    domain_addr = ss.insertSheet("Domain");
  }

  // Process the pending emails
  processInbox_(from, from_addr, domain_addr);
}
    
function generateRule(domain_sheet) {
  var range = domain_sheet.getDataRange();
  var rule = {};
  for(var i = 1; i < range.getNumRows(); i++) {
    if(domain_sheet.getDataRange().getCell(i,3).getValue() != "") {
       rule[domain_sheet.getDataRange().getCell(i,1).getDisplayValue()] = domain_sheet.getDataRange().getCell(i,3).getDisplayValue();
    }
  }
    
  return rule;
}    
    
function main_generateLabel() {
    
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.setActiveSheet(ss.getSheets()[0]);
  
  var from = ss.getSheetByName("From");
  var from_addr = ss.getSheetByName("From_Addr");
  var domain_addr = ss.getSheetByName("Domain");
  
  var domain_rule = generateRule(domain_addr);
  var addr_rule = generateRule(from_addr);
  
  // add user label
  for(var key in domain_rule) {
    if(GmailApp.getUserLabelByName(domain_rule[key]) == null) {
      GmailApp.createLabel(domain_rule[key]);
    }
  }
  
  for(var key in addr_rule) {
    if(GmailApp.getUserLabelByName(addr_rule[key]) == null) {
      GmailApp.createLabel(addr_rule[key]);
    }
  }
  
  
    
  var threads = GmailApp.getInboxThreads();
  var i = 0;
  for(var t in threads) {
    var thread = threads[t];
    
    var from = thread.getMessages()[0].getFrom();
    var email = parseEmail(from);
    var domain = parseDomain(email);

    if(domain in domain_rule) {
      thread.addLabel(GmailApp.getUserLabelByName(domain_rule[domain]));
    }
    
    if(email in addr_rule) {
      thread.addLabel(GmailApp.getUserLabelByName(addr_rule[email]));
    }
    
    i++;
    i = i+1;
    if (i > MAX_ITER) {
      break;
    }
  }
}
