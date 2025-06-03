# 📅 Google Calendar Work Report Generator

This script generates and emails a work report based on Google Calendar events. Events are filtered by color and reported in either **robust** or **minimal** format.

---

## 🔧 Configuration

Defined in the `CONFIG` object:

- `calendarName`: Name of the Google Calendar to pull events from
- `startDate` / `endDate`: Date range to filter events
- `emailAddress`: Recipient of the final report
- `targetColors`: Color key for filtering relevant events (e.g. `{ "": "Work" }`)
- `reportMode`: `"robust"` for detailed reports or `"minimal"` for daily/hour summaries

---

## 📊 Report Modes

### 🔍 Robust Report
Detailed HTML table with:
- Event title, description, start/end time
- Duration
- Grouped by day with daily summaries
- Total hours at the top

### 🧾 Minimal Report
Condensed HTML table with:
- Total work hours for the entire month
- Per-day work hour summary

---

## 📤 Output

The report is emailed using Google Apps Script's `MailApp.sendEmail` with HTML formatting.

---

## 🔁 Flow Overview

1. `summarizeWorkEventsReport()` – Main entry point
2. `getCalendarByName()` – Finds the calendar by name
3. `getFilteredEvents()` – Filters events by date and color
4. `generateMinimalReport()` or `generateRobustReport()` – Depending on selected mode
5. `sendEmailReport()` – Emails the formatted HTML report

---

## 🧹 HTML Cleanup

Descriptions are sanitized using `cleanHtmlDescription()`:
- Converts `<br>`, `<li>`, and other elements to plaintext
- Removes span and list containers

---

## 📌 Requirements

- Google Apps Script environment
- Access to the target Google Calendar
- Valid email address for delivery

---

## ✅ Example Usage

To generate a minimal report:

```js
CONFIG.reportMode = "minimal";
summarizeWorkEventsReport();

---

To switch to a robust report:

js
Copy
Edit


```js
CONFIG.reportMode = "robust";
summarizeWorkEventsReport();
---