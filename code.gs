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
  Logger.log("ParseDomain");
  Logger.log(email_addr);
  Logger.log(domain);

  return domain;
}

function processThread(from_dict, addr_dict, domain_dict, to_dict, thread) {
  var from = thread.getMessages()[0].getFrom();
  if (!(from in from_dict)) {
    from_dict[from] = 1;
  } else {
    from_dict[from]++;
  }

  var to = thread.getMessages()[0].getTo();
  if (!(to in to_dict)) {
    to_dict[to] = 1;
  } else {
    to_dict[to]++;
  }

  var email = parseEmail(from);
  if (!(email in addr_dict)) {
    addr_dict[email] = 1;
  } else {
    addr_dict[email]++;
  }

  var domain = parseDomain(email);
  if (!(domain in domain_dict)) {
    domain_dict[domain] = 1;
  } else {
    domain_dict[domain]++;
  }
}

// Sweep through inbox and create stat
function processInbox_(from_sheet, from_addr_sheet, from_domain_sheet,
                       to_sheet) {
  var threads = GmailApp.getInboxThreads(0, MAX_ITER);

  var from_dict = {};
  var addr_dict = {};
  var domain_dict = {};
  var to_dict = {};

  threads.forEach(function(t) {
    processThread(from_dict, addr_dict, domain_dict, to_dict, t);
  });

  for (var key in from_dict) {
    from_sheet.appendRow([ key, from_dict[key] ]);
  }

  for (var key in to_dict) {
    to_sheet.appendRow([ key, to_dict[key] ]);
  }

  for (var key in addr_dict) {
    from_addr_sheet.appendRow([ key, addr_dict[key] ]);
  }

  for (var key in domain_dict) {
    from_domain_sheet.appendRow([ key, domain_dict[key] ]);
  }

  from_sheet.sort(2, false);
  from_addr_sheet.sort(2, false);
  from_domain_sheet.sort(2, false);
}

function main_deleteAllLabels() {
  var labels = GmailApp.getUserLabels();
  for (var i = 0; i < labels.length; i++) {
    GmailApp.deleteLabel(labels[i]);
  }
}

// Generate clean sheet for use, and preserve the old rules
function generateSheets(ss) {
  var from_sheet = ss.getSheetByName("From");
  var from_addr_sheet = ss.getSheetByName("From_Addr");
  var from_domain_sheet = ss.getSheetByName("From_Domain");
  var to_sheet = ss.getSheetByName("To");

  var from_domain_rule = {};
  var from_addr_rule = {};
  var from_rule = {};
  var to_rule = {};

  if (from_sheet != null) {
    from_rule = generateRule(from_sheet);
    from_sheet = from_sheet.clear();
  } else {
    from_sheet = ss.insertSheet("From");
  }

  if (to_sheet != null) {
    to_rule = generateRule(to_sheet);
    to_sheet = to_sheet.clear();
  } else {
    to_sheet = ss.insertSheet("To");
  }

  if (from_addr_sheet != null) {
    from_addr_rule = generateRule(from_addr_sheet);
    from_addr_sheet = from_addr_sheet.clear();
  } else {
    from_addr_sheet = ss.insertSheet("From_Addr");
  }

  if (from_domain_sheet != null) {
    from_domain_rule = generateRule(from_domain_sheet);
    from_domain_sheet = from_domain_sheet.clear();
  } else {
    from_domain_sheet = ss.insertSheet("From_Domain");
  }

  return {
    from_sheet:from_sheet, 
    from_addr_sheet : from_addr_sheet,
    from_domain_sheet : from_domain_sheet,
    to_sheet : to_sheet,
    from_rule: from_rule,
    from_addr_rule : from_addr_rule,
    from_domain_rule : from_domain_rule,
    to_rule : to_rule
  }
}

function restoreRuleOnSheet(sheet, rules) {
  var range = sheet.getDataRange();
  var range = sheet.getRange(range.getRow(), range.getColumn(),
                             range.getNumRows(), range.getNumColumns() + 1);

  for (var i = 0; i < range.getNumRows(); i++) {
    var row_idx = i + range.getRow();
    var key = range.getCell(row_idx, 1).getDisplayValue();
    if (key in rules) {
      range.getCell(row_idx, 3).setValue(rules[key]);
      delete rules[key];
    }
  }

  // whatever left we just append
  for (var key in rules) {
    sheet.appendRow([ key, 1, rules[key] ]);
  }
}

function restoreOldRules(sheets_and_rules) {
  restoreRuleOnSheet(sheets_and_rules.from_sheet, sheets_and_rules.from_rule);
  restoreRuleOnSheet(sheets_and_rules.from_addr_sheet,
                     sheets_and_rules.from_addr_rule);
  restoreRuleOnSheet(sheets_and_rules.from_domain_sheet,
                     sheets_and_rules.from_domain_rule);
  restoreRuleOnSheet(sheets_and_rules.to_sheet, sheets_and_rules.to_rule);
}

// main()
function main_emailDatafromSpreadsheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var sheets_and_rules = generateSheets(ss);
  // Process the pending emails
  processInbox_(sheets_and_rules.from_sheet, sheets_and_rules.from_addr_sheet,
                sheets_and_rules.from_domain_sheet, sheets_and_rules.to_sheet);
  // Put old rules back
  restoreOldRules(sheets_and_rules);
}

function generateRule(domain_sheet) {
  var range = domain_sheet.getDataRange();

  var rule = {};
  if (range.getNumColumns() < 3) {
    return rule;
  }
  for (var i = 1; i < range.getNumRows(); i++) {
    if (range.getCell(i, 3).getValue() != "") {
      rule[range.getCell(i, 1).getDisplayValue()] =
          range.getCell(i, 3).getDisplayValue();
    }
  }

  return rule;
}

// If the corresponding sheet is not existed, return null
function generateRules(ss) {
  var from = ss.getSheetByName("From");
  var from_addr = ss.getSheetByName("From_Addr");
  var domain_addr = ss.getSheetByName("From_Domain");

  if (from == null || from_addr == null || domain_addr == null) {
    return null;
  }

  var domain_rule = generateRule(domain_addr);
  var addr_rule = generateRule(from_addr);
  var from_rule = generateRule(from);

  return {
  from_rule:
    from_rule, domain_rule : domain_rule, addr_rule : addr_rule
  }
}

function main_generateLabel() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var rules = generateRules(ss);

  // add user label
  for (var key in rules.domain_rule) {
    if (GmailApp.getUserLabelByName(rules.domain_rule[key]) == null) {
      GmailApp.createLabel(rules.domain_rule[key]);
    }
  }

  for (var key in rules.addr_rule) {
    if (GmailApp.getUserLabelByName(rules.addr_rule[key]) == null) {
      GmailApp.createLabel(rules.addr_rule[key]);
    }
  }

  var threads = GmailApp.getUserLabelByName("unprocessed").getThreads();
  var i = 0;
  for (var t in threads) {
    var thread = threads[t];

    var from = thread.getMessages()[0].getFrom();
    var email = parseEmail(from);
    var domain = parseDomain(email);

    if (domain in rules.domain_rule) {
      thread.addLabel(GmailApp.getUserLabelByName(rules.domain_rule[domain]));
    }

    if (email in rules.addr_rule) {
      thread.addLabel(GmailApp.getUserLabelByName(rules.addr_rule[email]));
    }
    thread.removeLabel(GmailApp.getUserLabelByName("unprocessed"));
  }
}
