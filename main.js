function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("日曆功能")
    .addItem("建立學期日曆事件", "main")
    .addSeparator()
    .addItem("修改教室與課程資料", "modify_class")
    .addSeparator()
    .addSubMenu(ui.createMenu("清除所有教室日曆的事件").addItem("確定", "clean_calendar"))
    .addItem("顯示使用說明", "show_sideBar")
    .addToUi();
}

/**
 * Show instruction.
 */
function show_sideBar() {
  const htmlOutput = HtmlService.createHtmlOutputFromFile("description").setTitle("使用說明");
  SpreadsheetApp.getUi().showSidebar(htmlOutput);
}

function main() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("建立學期課表");
  const drive = new ControlDrive(sheet);
  const spreadsheetId = drive.convert_excel_to_googlesheet();
  const courseSheet = SpreadsheetApp.openById(spreadsheetId).getActiveSheet();
  const courseHeads = courseSheet.getRange(1, 1, 1, courseSheet.getLastColumn()).getValues();
  const courseHeadColumns = courseHeads[0]
    .map((value, index) => {
      if (value === "cls_name" || value === "sub_name" || value === "crk_name" || value === "scr_period")
        return index + 1;
    })
    .filter((value) => value != undefined);

  const courseSheetLastRow = courseSheet.getLastRow();
  let courseDataArray = [];
  for (let col = 0; col < courseHeadColumns.length; col++) {
    const colData = courseSheet
      .getRange(2, courseHeadColumns[col], courseSheetLastRow - 1, 1)
      .getValues()
      .map((value) => value[0]);
    courseDataArray.push(colData);
  }

  let courseDataObjects = [];
  for (let i = 0; i < courseSheetLastRow - 1; i++) {
    // otherImformation: e.g. '(三)03-04 B2-206 (四)03-04 B2-206 林垂彩'
    const otherImformation = courseDataArray[3][i].split(" ").filter((value) => value != "");
    // The classroom under than floor 3 in B2.
    if (otherImformation[1].slice(0, 2) === "B2" && parseInt(otherImformation[1].charAt(3)) < 3) {
      // 1 class period.
      const tmpObj = {
        cls_name: courseDataArray[0][i].slice(3),
        sub_name: courseDataArray[1][i],
        crk_name: courseDataArray[2][i],
        scr_period: {
          day: get_course_day(otherImformation[0].charAt(1)).ordinal(),
          startLesson: parseInt(otherImformation[0].slice(3).split("-")[0]),
          endLesson:
            otherImformation[0].slice(3).split("-")[1] === undefined
              ? parseInt(otherImformation[0].slice(3).split("-")[0])
              : parseInt(otherImformation[0].slice(3).split("-")[1]),
          lessonLength:
            otherImformation[0].slice(3).split("-")[1] === undefined
              ? 1
              : parseInt(otherImformation[0].slice(3).split("-")[1]) -
                parseInt(otherImformation[0].slice(3).split("-")[0]) +
                1,
          space: otherImformation[1],
          teacher:
            otherImformation[otherImformation.length - 1].length > 3
              ? "待聘"
              : otherImformation[otherImformation.length - 1],
        },
      };
      courseDataObjects.push(tmpObj);

      // 2 class periods.
      if (otherImformation.length == 5) {
        const tmpObj = {
          cls_name: courseDataArray[0][i].slice(3),
          sub_name: courseDataArray[1][i],
          crk_name: courseDataArray[2][i],
          scr_period: {
            day: get_course_day(otherImformation[2].charAt(1)).ordinal(),
            startLesson: parseInt(otherImformation[2].slice(3).split("-")[0]),
            endLesson:
              otherImformation[2].slice(3).split("-")[1] === undefined
                ? parseInt(otherImformation[2].slice(3).split("-")[0])
                : parseInt(otherImformation[2].slice(3).split("-")[1]),
            lessonLength:
              otherImformation[2].slice(3).split("-")[1] === undefined
                ? 1
                : parseInt(otherImformation[2].slice(3).split("-")[1]) -
                  parseInt(otherImformation[2].slice(3).split("-")[0]) +
                  1,
            space: otherImformation[3],
            teacher:
              otherImformation[otherImformation.length - 1].length > 3
                ? "待聘"
                : otherImformation[otherImformation.length - 1],
          },
        };
        courseDataObjects.push(tmpObj);
      }
      // 3 class periods.
      if (otherImformation.length == 7) {
        const tmpObj = {
          cls_name: courseDataArray[0][i].slice(3),
          sub_name: courseDataArray[1][i],
          crk_name: courseDataArray[2][i],
          scr_period: {
            day: get_course_day(otherImformation[4].charAt(1)).ordinal(),
            startLesson: parseInt(otherImformation[4].slice(3).split("-")[0]),
            endLesson:
              otherImformation[4].slice(3).split("-")[1] === undefined
                ? parseInt(otherImformation[4].slice(3).split("-")[0])
                : parseInt(otherImformation[4].slice(3).split("-")[1]),
            lessonLength:
              otherImformation[4].slice(3).split("-")[1] === undefined
                ? 1
                : parseInt(otherImformation[4].slice(3).split("-")[1]) -
                  parseInt(otherImformation[4].slice(3).split("-")[0]) +
                  1,
            space: otherImformation[5],
            teacher:
              otherImformation[otherImformation.length - 1].length > 3
                ? "待聘"
                : otherImformation[otherImformation.length - 1],
          },
        };
        courseDataObjects.push(tmpObj);
      }
    }
  }

  courseDataArray.length = 0;
  courseDataArray = courseDataObjects.map((obj) => {
    let rowArray = [];
    for (let key in obj) {
      if (typeof obj[key] === "string") {
        rowArray.push(obj[key]);
        continue;
      }
      for (let objInObj in obj[key]) {
        rowArray.push(obj[key][objInObj]);
      }
    }
    return rowArray;
  });

  // Write all courses into google sheet.
  sheet.getRange(3, 1, courseDataArray.length, courseDataArray[0].length).setValues(courseDataArray);

  let semesterStartDate = sheet.getRange(1, 2).getValue();
  let semesterEndDate = sheet.getRange(1, 4).getValue();
  if (semesterStartDate === "") {
    SpreadsheetApp.getUi().alert("請在B1輸入學期開始日期");
  } else if (semesterEndDate == "") {
    SpreadsheetApp.getUi().alert("請在D1輸入學期結束日期");
  } else {
    semesterStartDate = new Date(semesterStartDate);
    semesterEndDate = new Date(semesterEndDate);
    // Create calendar event one by one.
    for (let i = 0; i < courseDataArray.length; i++) {
        create_calendar_event(courseDataArray[i], semesterStartDate, semesterEndDate);
    }
    SpreadsheetApp.getUi().alert("學期課表已輸入日曆");
  }

  drive.deleteFile(spreadsheetId);
}

/**
 * Filter spaces to control calendars.
 * @param {string} space
 * @returns boolean
 */
function getToDoCalendar(space) {
  switch (space) {
    case "B2-101":
    case "B2-201":
    case "B2-202":
    case "B2-203":
    case "B2-204":
    case "B2-205":
    case "B2-206":
    case "B2-213":
    case "B2-214":
    case "B2-215":
    case "B2-216":
      return true;
    default:
      return false;
  }
}

/**
 * Add an event to the corresponding calendar.
 * @param {array} newEvent
 * @param {string} semesterStartDate
 * @param {string} semesterEndDate
 */
function create_calendar_event(newEvent, semesterStartDate, semesterEndDate) {
  let studentClass = newEvent[0];
  let course = newEvent[1];
  let courseType = newEvent[2];
  let day = newEvent[3];
  let startLesson = newEvent[4];
  let lessonLength = newEvent[6];
  let space = newEvent[7];
  let teacher = newEvent[8];

  try {
    let calendar = getCalendar(space);
    let dayNextSemesterEndDate = new Date(semesterEndDate.getTime() + 1000 * 60 * 60 * 24); // 期末最後日隔天

    let lessonDateTime = get_lesson_date_time(day, semesterStartDate, startLesson, lessonLength);
    // 建立日曆事件 createEventSeries(title, startTime, endTime, recurrence, options)
    let event = calendar.createEventSeries(
      studentClass + " (" + courseType + ")",
      lessonDateTime.startDateTime,
      lessonDateTime.endDateTime,
      CalendarApp.newRecurrence()
        .addWeeklyRule()
        .onlyOnWeekdays([get_course_day(day)])
        .until(dayNextSemesterEndDate),
      { description: course + "\n授課教師 : " + teacher }
    );
    Logger.log(space + " " + event.getDescription() + " event created successfully.");
  } catch (e) {
    Logger.log("create_calendar_event::" + e);
  }
}

/**
 * Get the corresponding calendar object.
 * @param {string} space
 * @returns {Calendar Object}
 */
function getCalendar(space) {
  switch (space) {
    case "B2-101":
      return CalendarApp.getCalendarById("CALENDAR_ID");
    case "B2-201":
      return CalendarApp.getCalendarById("CALENDAR_ID");
    case "B2-202":
      return CalendarApp.getCalendarById("CALENDAR_ID");
    case "B2-203":
      return CalendarApp.getCalendarById("CALENDAR_ID");
    case "B2-204":
      return CalendarApp.getCalendarById("CALENDAR_ID");
    case "B2-205":
      return CalendarApp.getCalendarById("CALENDAR_ID");
    case "B2-206":
      return CalendarApp.getCalendarById("CALENDAR_ID");
    case "B2-213":
      return CalendarApp.getCalendarById("CALENDAR_ID");
    case "B2-214":
      return CalendarApp.getCalendarById("CALENDAR_ID");
    case "B2-215":
      return CalendarApp.getCalendarById("CALENDAR_ID");
    case "B2-216":
      return CalendarApp.getCalendarById("CALENDAR_ID");
    default:
      return -1;
  }
}

/**
 * Get object includes startDateTime & endDateTime.
 * @param {number} day
 * @param {Date} semesterStartDate
 * @param {number} startLesson
 * @param {number} lessonLength
 * @returns {Object}
 */
function get_lesson_date_time(day, semesterStartDate, startLesson, lessonLength) {
  // 節次開始時間 startDateTime
  let courseDay = 1000 * 60 * 60 * 24 * (day - 1);
  let lessonHour = get_lesson_time(startLesson).hour * 1000 * 60 * 60;
  let lessonMinute = get_lesson_time(startLesson).minute * 1000 * 60;
  let startDateTime = new Date(semesterStartDate.getTime() + courseDay + lessonHour + lessonMinute);

  // 節次結束時間 endDateTime
  lessonHour = get_lesson_time(startLesson + lessonLength - 1).hour * 1000 * 60 * 60;
  lessonMinute = get_lesson_time(startLesson + lessonLength - 1).minute * 1000 * 60;
  let endDateTime = new Date(semesterStartDate.getTime() + courseDay + lessonHour + lessonMinute + 1000 * 60 * 50);
  return { startDateTime: startDateTime, endDateTime: endDateTime };
}

/**
 * Get object includes hour & minute from startLesson.
 * @param {number} startLesson
 * @returns {Object}
 */
function get_lesson_time(startLesson) {
  switch (startLesson) {
    case 1:
      return { hour: 8, minute: 0 };
    case 2:
      return { hour: 9, minute: 0 };
    case 3:
      return { hour: 10, minute: 10 };
    case 4:
      return { hour: 11, minute: 10 };
    case 5:
      return { hour: 13, minute: 0 };
    case 6:
      return { hour: 14, minute: 0 };
    case 7:
      return { hour: 15, minute: 10 };
    case 8:
      return { hour: 16, minute: 10 };
    case 9:
      return { hour: 17, minute: 10 };
    default:
      return -1;
  }
}

/**
 * Get calendar weekday from number.
 * @param {number} day
 * @returns {CalendarApp.Weekday}
 */
function get_course_day(day) {
  switch (day) {
    case 1:
    case "一":
      return CalendarApp.Weekday.MONDAY;
    case 2:
    case "二":
      return CalendarApp.Weekday.TUESDAY;
    case 3:
    case "三":
      return CalendarApp.Weekday.WEDNESDAY;
    case 4:
    case "四":
      return CalendarApp.Weekday.THURSDAY;
    case 5:
    case "五":
      return CalendarApp.Weekday.FRIDAY;
    default:
      return -1;
  }
}

class ControlDrive {
  constructor(sheet) {
    this.sheet = sheet;
  }
  /**
   * Get converted google sheet id.
   * @returns {string}
   */
  convert_excel_to_googlesheet() {
    const spreadsheetID = this.sheet.getParent().getId();
    this.folderID = DriveApp.getFileById(spreadsheetID).getParents().next().getId();
    this.folder = DriveApp.getFolderById(this.folderID);
    const files = this.folder.getFiles();
    while (files.hasNext()) {
      let file = files.next();
      if (file.getName() != "建立學期課表日曆") {
        const resourse = {
          title: file.getName(),
          mimeType: MimeType.GOOGLE_SHEETS,
          parents: [{ id: file.getParents().next().getId() }],
        };
        const spreadsheet = Drive.Files.insert(resourse, file.getBlob());
        Drive.Files.remove(file.getId());
        return spreadsheet.id;
      }
    }
  }
  /**
   * Remove file from drive folder.
   * @param {string} fileId
   */
  deleteFile(fileId) {
    Drive.Files.remove(fileId);
  }
}
