/**
 * === Configuration Constants ===
 * Define core settings like calendar name, date range, email recipient, and report format.
 */
const FORMAT = ["minimal", "robust"];

const CONFIG = {
    calendarName: "Project Manager [WH]",
    startDate: new Date("2025-05-01T00:00:00"),
    endDate: new Date("2025-06-01T00:00:00"),
    emailAddress: "EMAIL_ADDRESS@gmail.com",
    targetColors: { "": "Work" }, // Events with this color are considered "Work"
    reportMode: FORMAT[0], // Options: "robust" or "minimal"
  };
  
  /**
   * === Entry Point ===
   * Main function to generate and send a work report email based on calendar events.
   */
  function summarizeWorkEventsReport() {
    const calendar = getCalendarByName(CONFIG.calendarName);
    if (!calendar) return;
  
    const events = getFilteredEvents(calendar, CONFIG.startDate, CONFIG.endDate, CONFIG.targetColors);
  
    const htmlReport =
      CONFIG.reportMode === "minimal"
        ? generateMinimalReport(events)
        : generateRobustReport(events);
  
    sendEmailReport(CONFIG.emailAddress, htmlReport);
  }
  
  /**
   * === Get Calendar by Name ===
   * @param {string} name - Name of the calendar to retrieve
   * @returns {Calendar} Calendar instance or null if not found
   */
  function getCalendarByName(name) {
    const calendars = CalendarApp.getCalendarsByName(name);
    if (calendars.length === 0) {
      Logger.log(`No calendar found with name '${name}'`);
      return null;
    }
    return calendars[0];
  }
  
  /**
   * === Get Filtered Events ===
   * @param {Calendar} calendar - Calendar object
   * @param {Date} startDate - Start of time range
   * @param {Date} endDate - End of time range
   * @param {Object} colorMap - Map of accepted color IDs
   * @returns {Array<Object>} Array of filtered event objects
   */
  function getFilteredEvents(calendar, startDate, endDate, colorMap) {
    const events = calendar.getEvents(startDate, endDate);
    return events
      .filter(event => colorMap.hasOwnProperty(event.getColor()))
      .map(event => ({
        title: event.getTitle(),
        description: cleanHtmlDescription(event.getDescription()),
        startTime: event.getStartTime(),
        endTime: event.getEndTime(),
        duration: (event.getEndTime() - event.getStartTime()) / (1000 * 60 * 60),
      }));
  }
  
  /**
   * === Generate Robust Report (Detailed) ===
   * Outputs a full HTML table with event details, grouped by day.
   * @param {Array<Object>} events - Filtered events
   * @returns {string} HTML string
   */
  function generateRobustReport(events) {
    const timeZone = Session.getScriptTimeZone();
    const sortedEvents = [...events].sort((a, b) => a.startTime - b.startTime);
    let html = '<table border="1" cellpadding="5" cellspacing="0" style="border-collapse: collapse; width: 100%;">';
  
    const totalWorkHours = sortedEvents.reduce((sum, e) => sum + e.duration, 0);
    html += `<tr style="background-color: #cccccc;"><td colspan="6"><strong>Total Work Hours: ${totalWorkHours.toFixed(2)} hours</strong></td></tr>`;
  
    let currentDay = "", dailyTotal = 0, eventCounter = 0;
  
    sortedEvents.forEach(event => {
      const dayStr = Utilities.formatDate(event.startTime, timeZone, "dd.MM.yyyy");
  
      if (dayStr !== currentDay) {
        if (currentDay) {
          html += `<tr style="background-color: #f0f0f0;"><td colspan="6"><strong>Daily Work Hours Summary: ${dailyTotal.toFixed(2)} hours</strong></td></tr>`;
          html += `<tr><td colspan="6" style="border-bottom: 2px solid #000;"></td></tr>`;
        }
        currentDay = dayStr;
        dailyTotal = 0;
        eventCounter = 0;
  
        html += `<tr style="background-color: #d0d0d0;"><td colspan="6"><strong>Date: ${currentDay}</strong></td></tr>`;
        html += `<tr style="font-weight: bold;"><td>Event #</td><td>Title</td><td>Description</td><td>Start Time</td><td>End Time</td><td>Duration (hours)</td></tr>`;
      }
  
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
  
    if (currentDay) {
      html += `<tr style="background-color: #f0f0f0;"><td colspan="6"><strong>Daily Work Hours Summary: ${dailyTotal.toFixed(2)} hours</strong></td></tr>`;
    }
  
    html += '</table>';
    return html;
  }
  
  /**
   * === Generate Minimal Report (Summary) ===
   * Outputs a basic table of daily work totals and the monthly total.
   * @param {Array<Object>} events - Filtered events
   * @returns {string} HTML string
   */
  function generateMinimalReport(events) {
    const timeZone = Session.getScriptTimeZone();
    const dayMap = {};
  
    events.forEach(event => {
      const dayStr = Utilities.formatDate(event.startTime, timeZone, "dd.MM.yyyy");
      dayMap[dayStr] = (dayMap[dayStr] || 0) + event.duration;
    });
  
    const sortedDays = Object.keys(dayMap).sort((a, b) =>
      new Date(a.split('.').reverse().join('-')) - new Date(b.split('.').reverse().join('-'))
    );
  
    const totalHours = Object.values(dayMap).reduce((a, b) => a + b, 0);
  
    let html = '<table border="1" cellpadding="5" cellspacing="0" style="border-collapse: collapse; width: 100%;">';
    html += `<tr style="background-color: #cccccc;"><td colspan="2"><strong>Total Work Hours: ${totalHours.toFixed(2)} hours</strong></td></tr>`;
    html += '<tr style="font-weight: bold;"><td>Date</td><td>Work Hours</td></tr>';
  
    sortedDays.forEach(date => {
      html += `<tr><td>${date}</td><td>${dayMap[date].toFixed(2)}</td></tr>`;
    });
  
    html += '</table>';
    return html;
  }
  
  /**
   * === Send Email Report ===
   * @param {string} to - Recipient email address
   * @param {string} htmlBody - The HTML content of the report
   */
  function sendEmailReport(to, htmlBody) {
    const subject = `Work Report (${CONFIG.reportMode}) for ${Utilities.formatDate(CONFIG.startDate, Session.getScriptTimeZone(), "MMMM yyyy")}`;
    MailApp.sendEmail({ to, subject, htmlBody });
  }
  
  /**
   * === Clean HTML Description ===
   * Converts rich HTML descriptions to clean text with bullets and newlines.
   * @param {string} html - The raw HTML from calendar event description
   * @returns {string} Cleaned plain-text version
   */
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
  
