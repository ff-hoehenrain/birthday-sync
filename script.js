function getRawData() {
  Logger.log("Get raw data from spreadsheet...")

  const values = SpreadsheetApp.getActive().getSheetByName("Aktuell").getDataRange().getValues();
  values.shift();

  return values;
}

function getActiveBirthdays(rawData) {
  Logger.log("Filtering active birthdays out...");

  var activeBirthdays = []
  rawData.forEach(value => {
    const name = "Geburtstag: " + value[4] + " " + value[3];
    const date = value[8];
    const state = value[6];

    if(state === 'aktiv') {
      const birthday = {
        name : name,
        date : date
      }
      activeBirthdays.push(birthday);
    }
  });

  return activeBirthdays;
}

function getBirthdayCalendar() {
  Logger.log("Get birthday calendar...");

  const calendar = CalendarApp.getCalendarById('c_j010drga8ov53i6ce1imrjie54@group.calendar.google.com');

  return calendar;
}

function syncBirthdayWithCalendar(calendar, birthdays) {
  Logger.log("Sync birthdays with calendar...");

  const currentDate = new Date();

  birthdays.forEach(birthday => {
    name = birthday.name;
    date = birthday.date;

    const year = currentDate.getFullYear();
    const month = date.getUTCMonth();
    const day = date.getUTCDate();
    const birthdayDate = new Date(year, month, day, 3, 0);
    const endBirthdayDate = new Date(year, month, day, 3, 0);

    const recurrence = CalendarApp.newRecurrence().addYearlyRule();

    Logger.log("Currently processing birthday with title=" + name + " at " + birthdayDate);

    const events = calendar.getEventsForDay(birthdayDate, {search: name});
    if(events.length === 0) {
      calendar.createEventSeries(name, birthdayDate, endBirthdayDate, recurrence);
    }

    Logger.log("Successfully proccessed event.");
  });
}

function birthdayReminder() {
  Logger.log("Processing birthday reminder...");

  const rawData = getRawData();
  const activeBirthdays = getActiveBirthdays(rawData);
  const calendar = getBirthdayCalendar();

  Logger.log("Size of birthdays=" + activeBirthdays.length);

  syncBirthdayWithCalendar(calendar, activeBirthdays);

  Logger.log("Successfully executed birthday reminder!");
}
