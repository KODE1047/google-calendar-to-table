/**
 * Creates a custom menu in the Google Sheet when the file is opened.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Events to Table')
    .addItem('Sync Events to Sheet', 'syncCalendarEvents')
    .addToUi();
}

/**
 * Main function to fetch calendar events and populate the sheet.
 * The week is defined as starting on the most recent Saturday.
 */
function syncCalendarEvents() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();

  // 1. Prepare the Sheet
  sheet.clear();
  const headers = ["Day", "Event", "Start Time", "End Time"];
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers]);
  headerRange.setFontWeight("bold");

  // 2. Calculate Date Range (Week starting Saturday)
  const today = new Date();
  const dayOfWeek = today.getDay(); // Sunday=0, Monday=1, ..., Saturday=6

  // Calculate days to rewind to get to the most recent Saturday.
  // (dayOfWeek + 1) % 7 gives the number of days past Saturday.
  // e.g., Friday (5): (5 + 1) % 7 = 6 days ago.
  // e.g., Saturday (6): (6 + 1) % 7 = 0 days ago.
  // e.g., Sunday (0): (0 + 1) % 7 = 1 day ago.
  const daysToRewind = (dayOfWeek + 1) % 7;
  
  const startDate = new Date(today.getFullYear(), today.getMonth(), today.getDate() - daysToRewind);
  startDate.setHours(0, 0, 0, 0); // Set to the beginning of that Saturday

  const endDate = new Date(startDate);
  endDate.setDate(startDate.getDate() + 7); // Get 7 full days (Sat, Sun, Mon, Tue, Wed, Thu, Fri)

  Logger.log(`Fetching events from ${startDate} to ${endDate}`);

  // 3. Fetch Calendar Events
  const calendar = CalendarApp.getDefaultCalendar();
  const events = calendar.getEvents(startDate, endDate);
  
  // Get the calendar's timezone for accurate date formatting
  const timezone = calendar.getTimeZone();

  // 4. Process and Format Events
  const data = [];
  for (const event of events) {
    const title = event.getTitle();
    const startTime = event.getStartTime();
    const endTime = event.getEndTime();

    // Format the date/time strings using the calendar's timezone
    const dayStr = Utilities.formatDate(startTime, timezone, 'EEEE'); // EEEE = Full day name
    const startTimeStr = Utilities.formatDate(startTime, timezone, 'HH:mm');
    const endTimeStr = Utilities.formatDate(endTime, timezone, 'HH:mm');

    data.push([dayStr, title, startTimeStr, endTimeStr]);
  }

  // 5. Write Data to Sheet (Batch operation)
  if (data.length > 0) {
    // Start writing from row 2 (since row 1 is headers)
    sheet.getRange(2, 1, data.length, data[0].length).setValues(data);
    
    // 6. Final Formatting
    // Set columns C and D (Start/End Time) to display as Time
    sheet.getRange(2, 3, data.length, 2).setNumberFormat("HH:mm");
    
    // Auto-resize all columns for readability
    sheet.autoResizeColumns(1, headers.length);
  } else {
    sheet.getRange(2, 1).setValue("No events found for this period.");
  }
}
