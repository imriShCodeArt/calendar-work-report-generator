# Calendar Work Report Generator

Google Apps Script for summarizing Google Calendar events tagged as "Work" into a comprehensive HTML report. The script calculates total and daily work hours, formats events into an HTML table, and emails the report.

## Features

- Fetches events from a specified calendar and time range
- Filters events by specific color (e.g., "" for "Work")
- Cleans and formats event descriptions
- Sorts events chronologically
- Groups events by date with daily totals
- Sends a styled HTML report via email

## Example Output

The email contains:
- Total work hours at the top
- Daily grouped sections with:
  - Event number
  - Title
  - Description
  - Start/End times
  - Duration
- Daily summaries

## Installation

1. Open [Google Apps Script](https://script.google.com/)
2. Create a new project
3. Copy the script code into the editor
4. Replace the calendar name and email address as needed
5. Save and run the function `summarizeWorkEventsReport`

## Customization

- **Calendar Name:** Change the `name` variable to match your calendar.
- **Target Color(s):** Modify the `targetColors` object to include different color IDs and labels.
- **Date Range:** Adjust `startDate` and `endDate` to the desired reporting period.

## License

MIT License
