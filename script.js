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
    const firstname = value[4];
    const lastname = value[3];
    const date = value[8];
    const state = value[6];

    if(state === 'aktiv' || state === 'jugend') {
      const birthday = {
        firstname : firstname,
        lastname : lastname,
        date : date
      }
      activeBirthdays.push(birthday);
    }
  });

  return activeBirthdays;
}

function getPassiveBirthdays(rawData) {
  Logger.log("Filtering passive birthdays out...");

  var passiveBirthdays = []
  rawData.forEach(value => {
    const firstname = value[4];
    const lastname = value[3];
    const date = value[8];
    const state = value[6];

    if(state === 'passiv') {
      const birthday = {
        firstname : firstname,
        lastname : lastname,
        date : date
      }
      passiveBirthdays.push(birthday);
    }
  });

  return passiveBirthdays;
}

function getActiveBirthdayCalendar() {
  Logger.log("Get active birthday calendar...");

  const calendar = CalendarApp.getCalendarById('c_j010drga8ov53i6ce1imrjie54@group.calendar.google.com');

  return calendar;
}

function getPassiveBirthdayCalendar() {
  Logger.log("Get passive birthday calendar...");

  const calendar = CalendarApp.getCalendarById('c_jcrml7eh3t9t0lp8jk8bpct614@group.calendar.google.com');

  return calendar;
}

function syncBirthdayWithCalendar(calendar, birthdays) {
  Logger.log("Sync birthdays with calendar...");

  const currentDate = new Date();

  birthdays.forEach(birthday => {
    const date = birthday.date;
    const currentYear = currentDate.getFullYear();
    const month = date.getUTCMonth();
    const day = date.getUTCDate();
    const birthdayDate = new Date(currentYear, month, day, 3, 0);
    const age = currentYear - date.getFullYear();
    const endBirthdayDate = new Date(currentYear, month, day, 3, 0);

    const name = age + ". Geburtstag von " + birthday.firstname + " " + birthday.lastname;

    Logger.log("Currently processing birthday with title=" + name + " at " + birthdayDate);

    const events = calendar.getEventsForDay(birthdayDate, {search: name});
    if(events.length === 0) {
      calendar.createEvent(name, birthdayDate, endBirthdayDate);
    }

    Logger.log("Successfully proccessed event.");
  });
}

function birthdayReminder() {
  Logger.log("Processing birthday reminder...");

  const rawData = getRawData();
  const activeBirthdays = getActiveBirthdays(rawData);
  const passiveBirthdays = getPassiveBirthdays(rawData);
  const activeBirthdayCalendar = getActiveBirthdayCalendar();
  const passiveBirthdayCalendar = getPassiveBirthdayCalendar();

  Logger.log("Size of active birthdays=" + activeBirthdays.length);
  syncBirthdayWithCalendar(activeBirthdayCalendar, activeBirthdays);
  
  Logger.log("Size of passive birthdays=" + passiveBirthdays.length);
  syncBirthdayWithCalendar(passiveBirthdayCalendar, passiveBirthdays);

  Logger.log("Successfully executed birthday reminder!");
}

function cleanup() {
  const calendar = getActiveBirthdayCalendar();

  calendar.getEvents(new Date('Jan 01 2022'), new Date('Dec 31 2022')).forEach(event => {
    Logger.log("Currently deleting event at" + event.toString());
    
    event.deleteEvent();
  })
}
