function refreshCalendarAndAddEvents() {
  generateActiveBidCalendar();  // Draw calendar in "Active Bids"
  addBidDueEventsToCalendar();  // Add events to Google Calendar
  SpreadsheetApp.getUi().alert("Refresh successful");
}

// ===== DRAW CALENDAR IN "Active Bids" =====
function generateActiveBidCalendar() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Active Bids");

  const startRow = 30; // CHANGED: Calendar now starts at Row 30
  const baseStartCol = 1;
  const data = sheet.getDataRange().getValues();
  const timeZone = Session.getScriptTimeZone();

  // CHANGED: Use current month and next month based on run date
  const runDate = new Date();
  const year = runDate.getFullYear();
  const month0 = runDate.getMonth();               // current month (0-based)
  const month1Date = new Date(year, month0 + 1, 1); // next month (handles year rollover)
  const targetMonths = [month0, month1Date.getMonth()];
  const targetYears = [year, month1Date.getFullYear()];

  const monthBids = {};

  for (let i = 1; i < data.length; i++) {
    const client = data[i][0];
    const bidNum = data[i][1];
    const status = data[i][3];      // Column D
    let closeDate = data[i][10];    // Column K
    const pbcDate = data[i][19];    // Column T

    // For PBC, use Column T if it has a valid date
    if (status === "PBC") {
      if (pbcDate instanceof Date && !isNaN(pbcDate)) {
        closeDate = pbcDate;
      } else if (pbcDate === "ASAP") {
        continue; // Skip ASAP PBC
      }
    }

    if ((status === "IN PROGRESS" || status === "PBC") && closeDate instanceof Date && !isNaN(closeDate)) {
      let label;
      if (status === "PBC") {
        label = `PBC DUE - ${bidNum} ${client} @ ${Utilities.formatDate(closeDate, timeZone, "h:mm a")}`;
      } else {
        label = `${bidNum} - ${client} @ ${Utilities.formatDate(closeDate, timeZone, "h:mm a")}`;
      }

      const m = closeDate.getMonth();
      const y = closeDate.getFullYear();

      // CHANGED: match either (currentYear,currentMonth) or (nextYear,nextMonth)
      const isTarget =
        (y === targetYears[0] && m === targetMonths[0]) ||
        (y === targetYears[1] && m === targetMonths[1]);

      if (isTarget) {
        // Key by year-month so Jan/Feb across year boundary doesn’t collide
        const key = `${y}-${m}`;
        if (!monthBids[key]) monthBids[key] = [];
        monthBids[key].push({ date: closeDate, label, status });
      }
    }
  }

  // Render current month and next month side-by-side
  const monthPairs = [
    { y: targetYears[0], m: targetMonths[0] },
    { y: targetYears[1], m: targetMonths[1] }
  ];

  monthPairs.forEach((mm, index) => {
    const startCol = baseStartCol + index * 5;
    const key = `${mm.y}-${mm.m}`;
    renderCalendar(sheet, startRow, startCol, mm.y, mm.m, monthBids[key] || []);
  });
}

// ===== RENDER CALENDAR HELPER =====
function renderCalendar(sheet, startRow, startCol, year, month, bids) {
  const weekdays = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday'];
  sheet.getRange(startRow, startCol, 1, weekdays.length).setValues([weekdays]);

  const timeZone = Session.getScriptTimeZone();
  const now = new Date();
  sheet.getRange(startRow + 1, startCol, 42, 5).clearContent().setBackground(null);

  const firstDate = new Date(year, month, 1);
  let row = startRow + 1;
  let currentDate = new Date(firstDate);

  while (currentDate.getMonth() === month) {
    const weekday = currentDate.getDay();
    if (weekday !== 0 && weekday !== 6) { // Skip weekends
      const colOffset = weekday - 1;
      const label = Utilities.formatDate(currentDate, timeZone, "MMM d");
      sheet.getRange(row, startCol + colOffset).setValue(label);

      const matching = bids.filter(b =>
        b.date.getDate() === currentDate.getDate() &&
        b.date.getMonth() === currentDate.getMonth() &&
        b.date.getFullYear() === currentDate.getFullYear()
      );

      if (matching.length > 0) {
        const cells = sheet.getRange(row + 1, startCol + colOffset, matching.length, 1);
        const values = matching.map(m => [m.label]);
        cells.setValues(values);

        for (let i = 0; i < matching.length; i++) {
          const due = matching[i].date;
          const status = matching[i].status;
          const diffDays = Math.floor((due - now) / (1000 * 60 * 60 * 24));
          let color = "#cfe2f3"; // Default
          if (status === "PBC") color = "#b6d7a8"; // PBC green
          else if (diffDays === 0) color = "#ea9999"; // Due today red
          else if (diffDays <= 7 && diffDays > 0) color = "#f4cccc"; // Soon
          cells.getCell(i + 1, 1).setBackground(color);
        }
      }
      if (weekday === 5) row += 6; // Move to next week
    }
    currentDate.setDate(currentDate.getDate() + 1);
  }

  sheet.setColumnWidths(startCol, 5, 160);
  sheet.autoResizeRows(startRow, 40);
}

// ===== ADD EVENTS TO GOOGLE CALENDAR =====
function addBidDueEventsToCalendar() {
  const sourceId = "TrackingsheetID";
  const range = "A1:K15";
  const sourceSheetName = "Active Bids";
  const values = SpreadsheetApp.openById(sourceId).getSheetByName(sourceSheetName).getRange(range).getValues();

  const calendarName = "Bid Due";
  const calendar = CalendarApp.getCalendarsByName(calendarName)[0];
  if (!calendar) throw new Error(`Calendar '${calendarName}' not found.`);

  values.forEach((row, idx) => {
    if (idx === 0) return; // skip header row

    const status = row[3];       // Column D
    const clientName = row[0];   // Column A
    const bidNumber = row[1];    // Column B
    let dueDate = row[10];       // Column K
    const pbcDate = row[19];     // Column T

    if (status === "PBC") {
      if (pbcDate instanceof Date && !isNaN(pbcDate)) {
        dueDate = pbcDate;
      } else if (pbcDate === "ASAP") return; // Skip ASAP
    }

    if (!clientName || !bidNumber || !(dueDate instanceof Date)) return;

    const title = status === "PBC" ? `PBC DUE - ${bidNumber} ${clientName}` : `BID DUE - ${bidNumber} ${clientName}`;
    const start = new Date(dueDate.getTime() - 60 * 60 * 1000); // 1 hour before
    const end = dueDate;

    // Prevent duplicate events
    const existingEvents = calendar.getEvents(start, end, { search: title });
    if (existingEvents.length > 0) return;

    const event = calendar.createEvent(title, start, end, {
      description: `Bid closing for ${clientName}`
    });

    event.setColor(status === "PBC" ? CalendarApp.EventColor.BLUE : CalendarApp.EventColor.RED);
    event.addPopupReminder(30);
  });
}
