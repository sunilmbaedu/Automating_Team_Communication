var spreadSheet = SpreadsheetApp.getActiveSpreadsheet();

var SHEETS = {
  STATUS: {
    NAME: "Status",
    COLUMNS: {
      TASK: 0,
      OWNER: 1,
      STATUS: 2,
      UPDATE: 3
    }
  },
  STAKEHOLDERS: {
    NAME: "Stakeholders",
    COLUMNS: {
      EMAIL: 2,
      COMMUNICATION_RULE: 5,
      SUBJECT: 7,
      HEADER: 8,
      FOOTER: 9
    }
  },
  COMMUNICATION_LOG: {
    NAME: "Communication Log"
  }
};

var COMMUNICATION_RULES = {
  CORE_TEAM: {
    NAME: "Core team",
    LAST_COLUMN: SHEETS.STATUS.COLUMNS.UPDATE
  },
  MONITOR: {
    NAME: "Monitor",
    LAST_COLUMN: SHEETS.STATUS.COLUMNS.STATUS
  },
  KEEP_SATISFIED: {
    NAME: "Keep satisfied",
    LAST_COLUMN: SHEETS.STATUS.COLUMNS.UPDATE
  },
  KEEP_INFORMED: {
    NAME: "Keep informed",
    LAST_COLUMN: SHEETS.STATUS.COLUMNS.STATUS
  },
  MANAGE_CLOSELY: {
    NAME:"Manage closely",
    LAST_COLUMN: SHEETS.STATUS.COLUMNS.UPDATE
  },
  URGENT: {
    NAME:"Urgent",
    LAST_COLUMN: SHEETS.STATUS.COLUMNS.UPDATE
  },
};

function sendEmail(required_communication_rule) {
  var emailRange = spreadSheet.getSheetByName(SHEETS.STAKEHOLDERS.NAME).getDataRange().getValues();

  var stakeholdersToEmail = emailRange.filter(
    function(stakeholdersrow) {
      return stakeholdersrow[SHEETS.STAKEHOLDERS.COLUMNS.COMMUNICATION_RULE] === required_communication_rule.NAME
    }
  );

  stakeholdersToEmail.forEach(function(stakeholdersrow) {
    var email_address = stakeholdersrow[SHEETS.STAKEHOLDERS.COLUMNS.EMAIL];
    var subject = stakeholdersrow[SHEETS.STAKEHOLDERS.COLUMNS.SUBJECT];
    var header = stakeholdersrow[SHEETS.STAKEHOLDERS.COLUMNS.HEADER];
    var footer = stakeholdersrow[SHEETS.STAKEHOLDERS.COLUMNS.FOOTER];
    
    var statusSheet = spreadSheet.getSheetByName(SHEETS.STATUS.NAME);
    var messageBodyRange = statusSheet.getRange(1, 1, statusSheet.getLastRow(), required_communication_rule.LAST_COLUMN + 1)

    var html_content = `
    <html>
      <body>
        <p>${header}</p>
        ${convertRange2html(messageBodyRange)}
        <p>${footer}</p>
      </body>
    </html>
    `;
      
    MailApp.sendEmail(email_address, subject, '', {htmlBody : html_content});

    var communication_log = spreadSheet.getSheetByName(SHEETS.COMMUNICATION_LOG.NAME);
    var date_time = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MM-dd-yyyy | HH:mm:ss");
    communication_log.appendRow([date_time, email_address, subject]);
  });
}

function sendEmailCoreTeam() {
  var required_communication_rule = COMMUNICATION_RULES.CORE_TEAM;
  sendEmail(required_communication_rule);
}

function sendEmailMonitor() {
  var required_communication_rule = COMMUNICATION_RULES.MONITOR;
  sendEmail(required_communication_rule);
}

function sendEmailKeepSatisfied() {
  var required_communication_rule = COMMUNICATION_RULES.KEEP_SATISFIED;
  sendEmail(required_communication_rule);
}

function sendEmailKeepInformed() {
  var required_communication_rule = COMMUNICATION_RULES.KEEP_INFORMED;
  sendEmail(required_communication_rule);
}

function sendEmailManageClosely() {
  var required_communication_rule = COMMUNICATION_RULES.MANAGE_CLOSELY;
  sendEmail(required_communication_rule);
}

function sendEmailUrgent() {
  var required_communication_rule = COMMUNICATION_RULES.URGENT;
  sendEmail(required_communication_rule);
}

function onEdit(e) {
  if (e.source.getSheetName() === SHEETS.STATUS.NAME) {
    sendEmailCoreTeam();
  }
}
