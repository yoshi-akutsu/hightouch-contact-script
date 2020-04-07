// The below variables can be changed to reflect which spreadsheet you want to use, what message you want to send, and how far back you want to check
// Make sure to format the spreadsheet with headers and put in column A- first name and last name, column B - student emails, column C and on - any number of emails to be cc'ed
// Use this as an example: https://docs.google.com/spreadsheets/d/1ExS3ypMqc3BbIXpNiJdD4udHscOayl1106rzgUcKVtE/edit#gid=2076019858 
// The spreadsheet ID is in the URL of the spreadsheet
// If you change the message, be sure to know that \n represents a new line and that the names in the greeting are pulled from the spreadsheet
// If you want to check in a different timeframe you can modify daysAgoToCheck to your desired timeframe
// It is currently set to 15 days
// I'd be wary of making it too large because the program can only check your past 500 email threads
// When you run it, make sure to run myFunction()
// Let me know if you have any questions: yoshi@collegeliftoff.com

var spreadsheetId = "1ExS3ypMqc3BbIXpNiJdD4udHscOayl1106rzgUcKVtE";
var msg = "Just writing to check in on you as we haven't been in contact or met in a few weeks. Just remember that I am here to help and please let me know if you would like to schedule a meeting or have any questions.\n\nBest,\nYoshi Akutsu";
var daysAgoToCheck = 15;

//
// Don't make changes below unless you know what you're doing
//

var now = new Date();
var twoWeeksAgo = new Date(now.getTime() - ((24* daysAgoToCheck)* 60 * 60 * 1000));
var emailedInLastTwoWeeks = [];
var eventInLastTwoWeeks = [];
var clientEmails = [];

// Adds all client emails from a spreadsheet to an array
function checkClientEmails() {
  var sheet = SpreadsheetApp.openById(spreadsheetId);
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][1] != "") {
      var allFamilyEmails = [];
      for (var j = 0; j < data[i].length; j++) {
        allFamilyEmails.push(data[i][j].toLowerCase())
      }
      clientEmails.push(allFamilyEmails)
    }
  }
}

// max smith -> Max Smith
function fixFormatting(name) {
  var array = name.split(" ");
  var firstName = array[0].split("");
  var lastName = array[1].split("");
  firstName[0] = firstName[0].toUpperCase();
  lastName[0] = lastName[0].toUpperCase();
  return firstName.join("") + " " + lastName.join("");
}

// Creates email drafts, ccing the appropriate parties
function createDraft(emailArray) {
  var ccEmails = "";
  for (var i = 2; i < emailArray.length; i++) {
    if (emailArray.length > 3) {
      ccEmails = ccEmails.concat(", ", emailArray[i])
    }
    else {
      ccEmails = emailArray[i];
    }
    
  }
  GmailApp.createDraft(emailArray[1], "Just checking in" , "Hey " + fixFormatting(emailArray[0]) + " and family,\n\n" + msg, {cc: ccEmails});
}

// Checks google calendar for all emails of people with events in past two weeks
function checkCalendar() {
  var events = CalendarApp.getDefaultCalendar().getEvents(twoWeeksAgo, now);
  for (var i = 0; i < events.length; i++) {
    var guests = events[i].getGuestList(true)
    for (var j = 0; j < guests.length; j++) {
      eventInLastTwoWeeks.push(guests[j].getEmail().toLowerCase());
    }
  }
}

// Takes Google message to, from, cc, and bcc labels and formats to email only
function cleanStringToEmails(stringEmails) {
  // Example input: Aaron Greene <aaron@collegeliftoff.com>, Alexandra Greene <alexandra@collegeliftoff.com>, Emma Mote <emma@collegeliftoff.org>
  var emails = [];
  var arrayUncleanedEmails = stringEmails.split(",");
  for (var i = 0; i < arrayUncleanedEmails.length; i++) {
    var emailTuple = arrayUncleanedEmails[i].split("<");
    if (emailTuple[1] === undefined){
    }
    else {
      var email = emailTuple[1].substring(0, emailTuple[1].length - 1);
      emails.push(email.toLowerCase());
    }
  }
  return emails;
}

// Adds emails to array that have recieved, sent, been cc'ed or bcc'ed in the past two weeks for the last 500 email threads
function checkEmail() {
  var threads = GmailApp.getInboxThreads(0, 500);
  // messages is an array of arrays of messages
  var messages = GmailApp.getMessagesForThreads(threads);
  for (var i = 0; i < messages.length; i++) {
    var date = messages[i][0].getDate();
    if (date > twoWeeksAgo) {
      if (messages[i][0].getFrom() != "") {
        for (var j = 0; j < cleanStringToEmails(messages[i][0].getFrom()).length; j ++) {
          emailedInLastTwoWeeks.push(cleanStringToEmails(messages[i][0].getFrom())[j])
        }
      }
      if (messages[i][0].getTo() != "") {
        for (var j = 0; j < cleanStringToEmails(messages[i][0].getTo()).length; j ++) {
          emailedInLastTwoWeeks.push(cleanStringToEmails(messages[i][0].getTo())[j])
        }
      }
      if (messages[i][0].getCc() != "") {
        for (var j = 0; j < cleanStringToEmails(messages[i][0].getCc()).length; j ++) {
          emailedInLastTwoWeeks.push(cleanStringToEmails(messages[i][0].getCc())[j])
        }
      }
      if (messages[i][0].getBcc() != "") {
        for (var j = 0; j < cleanStringToEmails(messages[i][0].getBcc()).length; j ++) {
          emailedInLastTwoWeeks.push(cleanStringToEmails(messages[i][0].getBcc())[j])
        }
      }
    }
  } 
}

function myFunction() {
  checkEmail();
  checkCalendar();
  checkClientEmails();
  var finalListEmails = emailedInLastTwoWeeks.concat(eventInLastTwoWeeks);
  // Clears duplicates
  finalListEmails = [... new Set(eventInLastTwoWeeks)];
  var mismatch = clientEmails.filter(element => !finalListEmails.includes(element[1]));
  for (var i = 0; i < mismatch.length; i++) {
    createDraft(mismatch[i]);
  }
}