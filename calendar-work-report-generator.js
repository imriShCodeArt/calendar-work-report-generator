function summarizeWorkEventsReport() {
    // Retrieve the calendar named "Project Manager [WH]"
    var name = "Project Manager [WH]"
    var calendars = CalendarApp.getCalendarsByName(name);
    if (calendars.length === 0) {
      Logger.log("No calendar found with the name '" + name + "'");
      return;
    }
    var calendar = calendars[0];
  
    // Define the time range (e.g., March 2025)
    var startDate = new Date("2025-05-01T00:00:00");
    var endDate = new Date("2025-06-01T00:00:00");
  
    // Fetch events from the specified calendar within the time range.
    var events = calendar.getEvents(startDate, endDate);
  
    // Define the target color: Empty string ("") will be renamed as "Work"
    var targetColors = { "": "Work" };
  
    // Array to hold detailed info for each relevant event.
    var eventDetails = [];
    
    // Variable to accumulate total work hours across all events.
    var totalWorkHours = 0;
  
    // Loop through events, processing only those that match the target color.
    events.forEach(function(event) {
      var colorId = event.getColor();
      if (targetColors.hasOwnProperty(colorId)) {
        var durationHours = (event.getEndTime() - event.getStartTime()) / (1000 * 60 * 60);
        
        // Sum up the overall total work hours.
        totalWorkHours += durationHours;
        
        // Clean the event description.
        var rawDescription = event.getDescription();
        var cleanedDescription = cleanHtmlDescription(rawDescription);
        
        // Gather event details.
        eventDetails.push({
          title: event.getTitle(),
          description: cleanedDescription,
          startTime: event.getStartTime(),
          endTime: event.getEndTime(),
          duration: durationHours
        });
      }
    });
  
    // Sort events by start time.
    eventDetails.sort(function(a, b) {
      return a.startTime - b.startTime;
    });
  
    var timeZone = Session.getScriptTimeZone();
  
    // Build the comprehensive events table.
    // The table includes:
    // - A top row displaying the total work hours (colspan all 6 columns).
    // - A header row for each new day (displaying the date in DD.MM.YYYY format).
    // - A row header for event detail columns.
    // - Rows for each event.
    // - A daily summary row showing the total work hours for that day.
    // - A divider row between days.
    var eventHtml = '<table border="1" cellpadding="5" cellspacing="0" style="border-collapse: collapse; width: 100%;">';
    
    // Add a top row for total work hours.
    eventHtml += '<tr style="background-color: #cccccc;"><td colspan="6"><strong>Total Work Hours: ' 
                 + totalWorkHours.toFixed(2) + ' hours</strong></td></tr>';
    
    // Initialize variables for day grouping.
    var currentDay = "";
    var dailyTotal = 0;
    var eventCounter = 0;
  
    eventDetails.forEach(function(detail) {
      // Format the event day as DD.MM.YYYY.
      var dayStr = Utilities.formatDate(detail.startTime, timeZone, "dd.MM.yyyy");
  
      // When the day changes, finish the previous day (if any), then add a header for the new day.
      if (dayStr !== currentDay) {
        // If it's not the first day, add the daily summary and a divider row.
        if (currentDay !== "") {
          eventHtml += '<tr style="background-color: #f0f0f0;"><td colspan="6"><strong>Daily Work Hours Summary: ' 
                       + dailyTotal.toFixed(2) + ' hours</strong></td></tr>';
          eventHtml += '<tr><td colspan="6" style="border-bottom: 2px solid #000;"></td></tr>';
        }
        currentDay = dayStr;
        dailyTotal = 0;
        eventCounter = 0;
        // Add the header row for the new day.
        eventHtml += '<tr style="background-color: #d0d0d0;"><td colspan="6"><strong>Date: ' 
                     + currentDay + '</strong></td></tr>';
        // Add a row with column headers for event details.
        eventHtml += '<tr style="font-weight: bold;">'
                   + '<td>Event #</td><td>Title</td><td>Description</td><td>Start Time</td><td>End Time</td><td>Duration (hours)</td>'
                   + '</tr>';
      }
  
      eventCounter++;
      dailyTotal += detail.duration;
  
      // Add the event row.
      eventHtml += '<tr>';
      eventHtml += '<td>' + eventCounter + '</td>';
      eventHtml += '<td>' + detail.title + '</td>';
      // Replace newline characters with <br> tags for HTML formatting.
      eventHtml += '<td>' + detail.description.replace(/\n/g, "<br>") + '</td>';
      eventHtml += '<td>' + Utilities.formatDate(detail.startTime, timeZone, "HH:mm") + '</td>';
      eventHtml += '<td>' + Utilities.formatDate(detail.endTime, timeZone, "HH:mm") + '</td>';
      eventHtml += '<td>' + detail.duration.toFixed(2) + '</td>';
      eventHtml += '</tr>';
    });
  
    // Add the daily summary for the last day.
    if (currentDay !== "") {
      eventHtml += '<tr style="background-color: #f0f0f0;"><td colspan="6"><strong>Daily Work Hours Summary: ' 
                  + dailyTotal.toFixed(2) + ' hours</strong></td></tr>';
    }
    
    eventHtml += '</table>';
  
    // Log the HTML output (for debugging or preview).
    Logger.log(eventHtml);
  
    // Send the comprehensive report via email using the HTML body.
    var emailAddress = "imriwain@gmail.com"; // Replace with your email address.
    var subject = "Comprehensive Work Report for " + Utilities.formatDate(startDate, timeZone, "MMMM yyyy");
    MailApp.sendEmail({
      to: emailAddress,
      subject: subject,
      htmlBody: eventHtml
    });
  }
  
  // Helper function to clean HTML tags from the description.
  // It replaces <br> with newlines, removes <span> tags,
  // and converts <ul>/<li> into a plain list with bullet points.
  function cleanHtmlDescription(html) {
    if (!html) return "";
    return html
      .replace(/<br\s*\/?>/gi, "\n")
      .replace(/<span>/gi, "")
      .replace(/<\/span>/gi, "")
      .replace(/<ul>/gi, "")
      .replace(/<\/ul>/gi, "")
      .replace(/<li>/gi, "• ")
      .replace(/<\/li>/gi, "\n");
  }
  