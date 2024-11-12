// Add a custom menu to the Google Sheets ribbon
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("Email Automation") // Name of the menu
    .addItem("Send Emails", "sendEmail") // Option to trigger the sendEmail function
    .addToUi();
}

function sendEmail() {
  var ui = SpreadsheetApp.getUi();

  // Confirmation dialog
  var response = ui.alert(
    "Send Emails",
    "This action will send emails to all recipients in the EmailList sheet. Do you want to proceed?",
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) {
    ui.alert("Email sending canceled.");
    return;
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("EmailList"); // Use "EmailList" instead of "Sheet1"
  var logSheet = ss.getSheetByName("Log"); // Logging sheet

  var emailRange = sheet.getRange(2, 1, sheet.getLastRow() - 1); // Emails start from row 2
  var emailAddresses = emailRange.getValues().flat();

  var htmlTemplate = HtmlService.createTemplateFromFile('email'); // Load email template
  var maxEmailsPerDay = 100; // Adjust this based on your account limits
  var emailCount = 0;
  var successfulSends = 0;
  var failedSends = 0;

  for (var i = 0; i < emailAddresses.length; i++) {
    if (emailCount >= maxEmailsPerDay) {
      Logger.log("Daily email limit reached.");
      break;
    }

    var currentRow = i + 2;
    var email = emailAddresses[i];
    if (!email) continue; // Skip blank rows

    // Retrieve dynamic fields
    var firstName = sheet.getRange(currentRow, 2).getValue();
    var lastName = sheet.getRange(currentRow, 3).getValue();
    var company = sheet.getRange(currentRow, 4).getValue();
    var language = sheet.getRange(currentRow, 5).getValue();

    // Set template variables
    htmlTemplate.firstName = firstName || "Valued";
    htmlTemplate.lastName = lastName || "Customer";
    htmlTemplate.company = company || "your organization";
    htmlTemplate.greeting = language === "French" ? "Bonjour" : "Hello";

    var emailHtml = htmlTemplate.evaluate().getContent();

    try {
      MailApp.sendEmail({
        to: email,
        subject: `Hello ${firstName || 'there'} from ${company || 'our team'}`,
        htmlBody: emailHtml
      });
      logSheet.appendRow([email, "Sent", new Date(), ""]);
      emailCount++;
      successfulSends++;
    } catch (error) {
      logSheet.appendRow([email, "Error", new Date(), error.message]);
      failedSends++;
    }
  }

  // Progress Notification
  if (successfulSends > 0 || failedSends > 0) {
    ui.alert(
      "Email Sending Complete",
      `Emails sent successfully: ${successfulSends}\nEmails failed to send: ${failedSends}`,
      ui.ButtonSet.OK
    );
  } else {
    ui.alert("No emails were sent. Please check your EmailList sheet.");
  }
}
