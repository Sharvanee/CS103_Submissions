function parseSpreadsheet() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getActiveSheet();
  var dataRange = sheet.getDataRange();
  var data = dataRange.getValues();
  data = removeDuplicateRows(data);

  // Create an object to store the count of applicants for each location and subject
  var locationQueue = {};

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var timestamp = row[0];
    var name = row[1];
    var contact = row[3];
    var subjectfilled = row[4];
    var location = row[5];
    var email = row[2];
    
    // If the location is not in the locationQueue object, initialize it
    if (!locationQueue.hasOwnProperty(location)) {
      locationQueue[location] = {};
    }
    
    var locationData = locationQueue[location];

    // If the subject is not in the locationData object, initialize it
    if (!locationData.hasOwnProperty(subjectfilled)) {
      locationData[subjectfilled] = {
        count: 0,
        slots: 6
      };
    }

    var subjectData = locationData[subjectfilled];

    // Check if the current location and subject have available appointment slots
    if (subjectData.slots > 0) {
      // Increment the count for the current location and subject, and generate token number
      subjectData.count++;
      var tokenNumber = subjectData.count;

      // Generate time slot based on token number
      var slotStartHour = 10 + Math.floor((tokenNumber - 1) / 2);
      var slotStartMinute = (tokenNumber % 2 === 1) ? '00' : '30';
      var slotEndHour = slotStartHour;
      var slotEndMinute = (tokenNumber % 2 === 1) ? '30' : '00';

      if (tokenNumber % 2 === 0) {
        slotEndHour += 1;
        slotEndMinute = '00';
      }

      // Formatting the time values
      var formattedStartTime = slotStartHour.toString().padStart(2, '0') + ':' + slotStartMinute;
      var formattedEndTime = slotEndHour.toString().padStart(2, '0') + ':' + slotEndMinute;
      var startTimePeriod = (slotStartHour >= 12) ? 'PM' : 'AM';
      var endTimePeriod = (slotEndHour >= 12) ? 'PM' : 'AM';
      if(tokenNumber<6){var timeSlot = formattedStartTime + ' ' + startTimePeriod + ' - ' + formattedEndTime + ' ' + endTimePeriod};
      if(tokenNumber==6){formattedEndTime =1;
        var timeSlot = formattedStartTime + ' ' + startTimePeriod + ' - ' + formattedEndTime + ' ' + endTimePeriod};

      console.log('Time slot:', timeSlot);
      Logger.log("Recipient Email: " + email);
      Logger.log(tokenNumber)
      // Get the current date
      var currentDate = new Date();
      var options = { day: 'numeric', month: 'numeric', year: 'numeric' };
      var formattedDate = currentDate.toLocaleDateString('en-IN', options);


      // Compose the email
      var subject = "Appointment confirmation with BT faculty under doubt clearing service";
      var message = "Dear " + name + ",\n\n";
      message += "Greetings from Bakliwal Tutorials. Hope your JEE preparation is going well. Being a part of the personal doubt clearing system will definitely benefit you in your JEE journey." + "\n"
      message += "Your appointment for the personal doubt session is scheduled at: " + location + " with the " + subjectfilled + " faculty at BT." + "\n";
      message += "Your token number is: " + tokenNumber + "\n";
      message += "Your appointment time slot is: " + timeSlot + " on: " + formattedDate + "\n";
      message += "Collect all your doubts in the subject of " + subjectfilled + " and be there on time to utilize the session to the fullest." + "\n"
      message += "Thank you" + "\n\n" + "Regards" + ",\n" + "Bakliwal Tutorials | Where Best Students Meet Best Teachers"

      // Send the email only if the token number is up to 6
      if (tokenNumber <= 6) {
        MailApp.sendEmail(email, subject, message);
        Logger.log(message);
      }
      // Reduce the available slots for the current location and subject
      subjectData.slots--;
    } else {
      // If there are no available appointment slots for the current location and subject, send a notification email
      var subject = "Appointment Slot Unavailable with BT faculty under doubt clearing service";
      var message = "Dear " + name + ",\n\n";
      message += "Greetings from Bakliwal Tutorials. Hope your JEE preparation is going well. Being a part of the personal doubt clearing system will definitely benefit you in your JEE journey." + "\n"
      message += "Token number : -"+"\n"
      message += "Appointment slot : Not available"+"\n"
      message += "Unfortunately, all appointment slots for " + location + " with the " + subjectfilled + " faculty are already filled on a first-come, first-served basis.\n";
      message += "If you prefer to schedule at a later date, you can fill out the form again on that particular day between 8 am to 8:30 am.\n";
      message += "Thank you" + "\n\n" + "Regards" + ",\n" + "Bakliwal Tutorials | Where Best Students Meet Best Teachers"

      // Send the denial email with confirmation link
      MailApp.sendEmail(email, subject, message);
      Logger.log(message);
    }
  }
}
function feedbackform(){
// Check if the current time is 1:30 PM
      var currentDate = new Date();
      var currentHour = currentDate.getHours();
      var currentMinute = currentDate.getMinutes();
      if (currentHour === 13 && currentMinute === 30) {
        // Send the feedback form email
        var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
        var sheet = spreadsheet.getActiveSheet();
        var dataRange = sheet.getDataRange();
        var data = dataRange.getValues();
        data = removeDuplicateRows(data);
        for (var i = 1; i < data.length; i++) {
         var row = data[i];
         var name = row[1];
         var email = row[2];
         var feedbackSubject = "Feedback form for BT faculty under doubt clearing service";
         var feedbackMessage = "Dear " + name + ",\n\n";
         feedbackMessage += "We hope the doubt clearing session with our faculty at Bakliwal Tutorials was helpful for you." + "\n";
         feedbackMessage += "Your feedback is valuable to us, and we would appreciate it if you could take a few minutes to complete the feedback form linked below." + "\n";
         feedbackMessage += "Feedback Form: https://forms.gle/efp5Vv4P5Hm143d48" + "\n";
         feedbackMessage += "Thank you for your time and contribution." + "\n\n";
         feedbackMessage += "Regards," + "\n" + "Bakliwal Tutorials | Where Best Students Meet Best Teachers";

        // Send the feedback form email
        MailApp.sendEmail(email, feedbackSubject, feedbackMessage);
        Logger.log(feedbackMessage);
      }
      }
}
function removeDuplicateRows(data) {
  var uniqueKeys = {};

  for (var i = 1; i < data.length; i++) {
    var key = data[i][2];

    if (!uniqueKeys.hasOwnProperty(key)) {
      uniqueKeys[key] = true; // Store the unique key
    } else {
      data.splice(i, 1); // Remove the duplicate row
      i--; // Decrement the loop counter since the array length has changed
    }
  }

  return data;
}