function clean_calendar() {
  const ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('建立學期課表');
  const data = ss.getDataRange().getValues();
  const startDate = new Date(data[0][1]);
  const endDate = new Date(data[0][3] + 1000 * 60 * 60 * 24);
  const object = get_all_calendarID();
  for (let index in object) {
    Logger.log(object[index]);
    let calendar = CalendarApp.getCalendarById(object[index]);
    let events = calendar.getEvents(startDate, endDate);
    events.forEach(element => element.deleteEvent())
  }
  Logger.log('Deleted all events.');
}

function modify_class() {
  const ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('修改資料');
  const data = ss.getDataRange().getValues();
  const lastRow = ss.getLastRow();
  const semesterStartDate = new Date(data[0][1]);
  const semesterEndDate = new Date(data[0][3]);
  try {
    for (let i = 3; i < lastRow; i++) {
      // 【修改前】 is empty.
      if (data[i][0] == '' && data[i][1] == '' && data[i][8] != '' && data[i][9] != '') {
        // Create new event.
        let newData = [];
        for (let newDataIdx = 8; newDataIdx < ss.getLastColumn(); newDataIdx++) {
          newData.push(data[i][newDataIdx]);
        }
        create_calendar_event(newData, semesterStartDate, semesterEndDate);
        ss.getRange(i + 1, 1, 1, 16).setBackground('#c4beb5');  // Set background color to gray.
        add_curriculum(newData); // Add new course data to 【建立學期課表】.
        continue;
      }
      // 【修改前】 contains values.
      const day = data[i][3];
      const startLesson = data[i][4];
      const lessonLength = data[i][5];
      const space = data[i][6];
      let calendar = getCalendar(space);
      let lessonDateTime = getLessonDateTime(day, semesterStartDate, startLesson, lessonLength);
      let lessonStartDateTime = lessonDateTime.startDateTime;
      let lessonEndDateTime = lessonDateTime.endDateTime;

      // Delete event from calendar.
      for (let startDateTime = semesterStartDate; startDateTime < new Date(semesterEndDate + 1000 * 60 * 60 * 24); startDateTime = new Date(startDateTime.getTime() + 1000 * 60 * 60 * 24 * 7)) {
        delete_from_calendar(calendar, lessonStartDateTime, lessonEndDateTime);
        lessonStartDateTime = new Date(lessonStartDateTime.getTime() + 1000 * 60 * 60 * 24 * 7);
        lessonEndDateTime = new Date(lessonEndDateTime.getTime() + 1000 * 60 * 60 * 24 * 7);
      }

      // Create new event in calendar.
      let oldData = [];
      let newData = [];
      for (let oldDataIdx = 0, newDataIdx = 8; newDataIdx < ss.getLastColumn(); oldDataIdx++, newDataIdx++) {
        oldData.push(data[i][oldDataIdx]);
        newData.push(data[i][newDataIdx]);
      }
      create_calendar_event(newData, semesterStartDate, semesterEndDate);
      Logger.log('created');
      ss.getRange(i + 1, 1, 1, 16).setBackground('#c4beb5');  // Set background color to gray.
      Logger.log('set background color');
      update_curriculum(oldData, newData);  // Updte old course data in 【建立學期課表】.
      Logger.log('refreshed');
    }
  } catch (e) {
    Logger.log('modifyClass:: ' + e);
  }
}

function delete_from_calendar(calendar, startDateTime, endDateTime) {
  try {
    let events = calendar.getEvents(startDateTime, endDateTime);
    events[0].deleteEvent();
  } catch (e) {
    Logger.log('delete_from_calendar:: ' + e);
  }
}

function update_curriculum(oldData, newData) {
  let oldSS = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('建立學期課表');
  let data = oldSS.getDataRange().getValues();
  let oldSSLastRow = oldSS.getLastRow();
  // Search old data and replace it with new data.
  for (let i = 2; i < oldSSLastRow; i++) {
    if (data[i][5] == oldData[5] && data[i][2] == oldData[2] && data[i][3] == oldData[3] && data[i][4] == oldData[4]) {
      oldSS.getRange(i + 1, 1, 1, 8).setValues([newData]);  
    }
  }
}

function add_curriculum(newData) {
  let oldSS = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('建立學期課表');
  let oldSSLastRow = oldSS.getLastRow();
  oldSS.getRange(oldSSLastRow + 1, 1, 1, 8).setValues([newData]);
}

function get_all_calendarID() {
  return {
    "B2-101": "c_g62jno9rpp2ofd34b801lsa80g@group.calendar.google.com",
    "B2-201": "c_qo31q6u3fpp7b4bkjiu3t1qvcs@group.calendar.google.com",
    "B2-202": "c_259r0o259ch71smi5srihn3brs@group.calendar.google.com",
    "B2-203": "c_3s5hk709bapn9gbpjse611g4ps@group.calendar.google.com",
    "B2-204": "c_ejjjhaqndrbkpiau78cogcbki0@group.calendar.google.com",
    "B2-205": "c_aufsutsbtddmvfmsjfrq00h04g@group.calendar.google.com",
    "B2-206": "c_4r8dvfml3m8da5rb4jm2o54odk@group.calendar.google.com",
    "B2-213": "c_1jeb5n0gq4sp029fnvcilbvsno@group.calendar.google.com",
    "B2-214": "c_7q9k761fp81g53d909u9jn9kok@group.calendar.google.com",
    "B2-215": "c_m9p6euenahljtht0uu7876b6qs@group.calendar.google.com",
    "B2-216": "c_uufepo395piq3kh2q2eb1gq9h4@group.calendar.google.com"
  };
}