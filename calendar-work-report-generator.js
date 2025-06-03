// Configuration object for calendar and report settings
const CONFIG = {
    calendarName: "Project Manager [WH]",         // Name of the calendar to extract events from
    startDate: new Date("2025-05-01T00:00:00"),   // Start date for event filtering
    endDate: new Date("2025-06-01T00:00:00"),     // End date for event filtering
    emailAddress: "imriwain@gmail.com",           // Recipient email for the report
    targetColors: { "": "Work" },                 // Filter events by calendar color ("" = default)
  };
  
  // Main function that generates and sends the work events report
  function summarizeWorkEventsReport() {
    const calendar = getCalendarByName(CONFIG.calendarName);
    if (!calendar) return; // Abort if calendar is not found
  
    const events = getFilteredEvents(calendar, CONFIG.startDate, CONFIG.endDate, CONFIG.targetColors);
    const htmlReport = generateHtmlReport(events);
    sendEmailReport(CONFIG.emailAddress, htmlReport);
  }
  
  // Retrieves a calendar object by its name
  function getCalendarByName(name) {
    const calendars = CalendarApp.getCalendarsByName(name);
    if (calendars.length === 0) {
      Logger.log(`No calendar found with name '${name}'`);
      return null;
    }
    return calendars[0]; // Return the first matching calendar
  }
  
  // Filters calendar events based on the configured date range and target colors
  function getFilteredEvents(calendar, startDate, endDate, colorMap) {
    const events = calendar.getEvents(startDate, endDate);
    return events
      .filter(event => colorMap.hasOwnProperty(event.getColor())) // Match only events with specified color
      .map(event => ({
        title: event.getTitle(),
        description: cleanHtmlDescription(event.getDescription()),
        startTime: event.getStartTime(),
        endTime: event.getEndTime(),
        duration: (event.getEndTime() - event.getStartTime()) / (1000 * 60 * 60), // Duration in hours
      }));
  }
  
  // Generates an HTML table report from the filtered events list
  function generateHtmlReport(events) {
    const timeZone = Session.getScriptTimeZone();
    const sortedEvents = [...events].sort((a, b) => a.startTime - b.startTime);
    let html = '<table border="1" cellpadding="5" cellspacing="0" style="border-collapse: collapse; width: 100%;">';
  
    // Calculate and display total work hours
    const totalWorkHours = sortedEvents.reduce((sum, e) => sum + e.duration, 0);
    html += `<tr style="background-color: #cccccc;"><td colspan="6"><strong>Total Work Hours: ${totalWorkHours.toFixed(2)} hours</strong></td></tr>`;
  
    let currentDay = "", dailyTotal = 0, eventCounter = 0;
  
    sortedEvents.forEach((event, index) => {
      const dayStr = Utilities.formatDate(event.startTime, timeZone, "dd.MM.yyyy");
  
      // New day block
      if (dayStr !== currentDay) {
        if (currentDay) {
          // End summary for the previous day
          html += `<tr style="background-color: #f0f0f0;"><td colspan="6"><strong>Daily Work Hours Summary: ${dailyTotal.toFixed(2)} hours</strong></td></tr>`;
          html += `<tr><td colspan="6" style="border-bottom: 2px solid #000;"></td></tr>`;
        }
  
        // Initialize new day
        currentDay = dayStr;
        dailyTotal = 0;
        eventCounter = 0;
  
        // Day header and table columns
        html += `<tr style="background-color: #d0d0d0;"><td colspan="6"><strong>Date: ${currentDay}</strong></td></tr>`;
        html += `<tr style="font-weight: bold;"><td>Event #</td><td>Title</td><td>Description</td><td>Start Time</td><td>End Time</td><td>Duration (hours)</td></tr>`;
      }
  
      // Append event row
      eventCounter++;
      dailyTotal += event.duration;
      html += `<tr>
        <td>${eventCounter}</td>
        <td>${event.title}</td>
        <td>${event.description.replace(/\n/g, "<br>")}</td>
        <td>${Utilities.formatDate(event.startTime, timeZone, "HH:mm")}</td>
        <td>${Utilities.formatDate(event.endTime, timeZone, "HH:mm")}</td>
        <td>${event.duration.toFixed(2)}</td>
      </tr>`;
    });
  
    // Final day summary
    if (currentDay) {
      html += `<tr style="background-color: #f0f0f0;"><td colspan="6"><strong>Daily Work Hours Summary: ${dailyTotal.toFixed(2)} hours</strong></td></tr>`;
    }
  
    html += '</table>';
    return html;
  }
  
  // Sends the HTML report via email
  function sendEmailReport(to, htmlBody) {
    const subject = `Comprehensive Work Report for ${Utilities.formatDate(CONFIG.startDate, Session.getScriptTimeZone(), "MMMM yyyy")}`;
    MailApp.sendEmail({ to, subject, htmlBody });
  }
  
  // Cleans up HTML descriptions from calendar events into plain text format
  function cleanHtmlDescription(html) {
    return html
      ?.replace(/<br\s*\/?>/gi, "\n")
      .replace(/<span[^>]*>/gi, "")
      .replace(/<\/span>/gi, "")
      .replace(/<ul>/gi, "")
      .replace(/<\/ul>/gi, "")
      .replace(/<li>/gi, "• ")
      .replace(/<\/li>/gi, "\n") || "";
  }
  