function onFormSubmit(e) {
  const calendarId = 'calendarID@group.calendar.google.com';
  const calendar = CalendarApp.getCalendarById(calendarId);

  const responses = e.values;

  const name = responses[1];
  const absenceType = responses[2];
  const startDateStr = responses[3];
  const startTimeStr = responses[4];
  const endTimeStr = responses[5];
  const endDateStr = responses[6];
  const note = responses[7];

  if (!name || !absenceType || !startDateStr || !endDateStr) {
    Logger.log("Missing required fields. Event not created.");
    return;
  }

  const colorMap = {
    "Member 1": CalendarApp.EventColor.BLUE,
    "Member 2": CalendarApp.EventColor.YELLOW,
    "Member 3": CalendarApp.EventColor.PALE_RED,
    "Member 4": CalendarApp.EventColor.ORANGE,
    "Member 5": CalendarApp.EventColor.GREEN,
  };

  const startDate = new Date(startDateStr);
  const endDate = new Date(endDateStr);
  const isSameDay = startDateStr === endDateStr;

  const eventTitle = `${name} - ${absenceType}`;
  const description = note || "";
  let event;

  try {
    if (isSameDay && startTimeStr && endTimeStr) {
      const startDateTime = parseTimeOnDate(startDate, startTimeStr);
      const endDateTime = parseTimeOnDate(endDate, endTimeStr);

      if (!startDateTime || !endDateTime || startDateTime >= endDateTime) {
        Logger.log("Invalid or reversed time range. Skipping event.");
        return;
      }

      event = calendar.createEvent(eventTitle, startDateTime, endDateTime, {
        description,
      });

    } else {
      const allDayEnd = new Date(endDate);
      allDayEnd.setDate(allDayEnd.getDate() + 1);
      event = calendar.createAllDayEvent(eventTitle, startDate, allDayEnd, {
        description,
      });
    }

    const color = colorMap[name] || CalendarApp.EventColor.GRAY;
    event.setColor(color);

  } catch (err) {
    Logger.log("Event creation failed: " + err);
  }
}

// 🕐 Helper: Parse "h:mm AM/PM" time strings into a Date on the given date
function parseTimeOnDate(dateObj, timeStr) {
  try {
    const [time, modifier] = timeStr.split(' ');
    let [hours, minutes] = time.split(':').map(Number);

    if (modifier === 'PM' && hours < 12) hours += 12;
    if (modifier === 'AM' && hours === 12) hours = 0;

    const newDate = new Date(dateObj);
    newDate.setHours(hours, minutes, 0, 0);
    return newDate;
  } catch (err) {
    Logger.log("Failed to parse time string: " + timeStr);
    return null;
  }
}
