// ============================
// Constants
// ============================

const ONE_DAY_IN_SECONDS = 24 * 60 * 60 * 1000;
const TIMEZONE_ASIA_TOKYO = 'Asia/Tokyo';

// ============================
// Date utilities
// ============================

function copyDate(date) {
  return new Date(date.getTime())
}

function getMinDate(dateA, dateB) {
  return new Date(Math.min(dateA, dateB));
}

function getMaxDate(dateA, dateB) {
  return new Date(Math.max(dateA, dateB));
}

function isSameDay(dateA, dateB) {
  return dateA.getDate() == dateB.getDate()
}

function getDaysDiff(startTime, endTime) {
  return Math.ceil((endTime - startTime) / ONE_DAY_IN_SECONDS)
}

function getMinutesDiff(startTime, endTime) {
  return Math.ceil((endTime - startTime) / (1000 * 60))
}

function toStartOfNextDay(date) {
  const copied = copyDate(date);
  const nextDay = new Date(copied.setDate(date.getDate() + 1));
  return new Date(nextDay.setHours(0, 0, 0))
}

function toEndOfDay(date) {
  const copied = copyDate(date);
  return new Date(copied.setHours(23, 59, 59))
}

function getJPDay(date) {
  const days = ['日', '月', '火', '水', '木', '金', '土'];
  return days[date.getDay()];
}

function isHoliday(date) {
  const day = date.getDay();
  return day === 0 || day === 6
}

// ============================
// GAS utilities
// ============================

function getValueFromCell(spreadsheet, cell) {
  return spreadsheet.getRange(cell).getValue()
}

function writeValueToSpreadsheet(spreadsheetId, cell, value) {
  const ss = SpreadsheetApp.openById(spreadsheetId);
  const range = ss.getRange(cell);
  range.setValue(value);
}

function listScheduleTimes(calendarId, startTime, endTime) {
  const calendar = CalendarApp.getCalendarById(calendarId);
  const events = calendar.getEvents(startTime, endTime);
  return events.map(event => [event.getStartTime(), event.getEndTime()]);
}

function removeAllDaySchedule(times) {
  return times.filter(([start, end]) => (end - start) % ONE_DAY_IN_SECONDS !== 0)
}

// ============================
// Business logics
// ============================

function makeBusinessDate(baseDate, businessTime) {
  const copied = copyDate(baseDate);
  const [hour, minute, second] = [businessTime.getHours(), businessTime.getMinutes(), 0];
  return new Date(copied.setHours(hour, minute, second))
}

function minutesIsMoreThan(startTime, endTime, minutes) {
  const minutesDiff = getMinutesDiff(startTime, endTime);
  return minutesDiff >= minutes
}

function formatHeadTime(date) {
  const day = getJPDay(date);
  return Utilities.formatDate(date, TIMEZONE_ASIA_TOKYO, `M月d日（${day}） HH:mm`);
}

function formatTime(date) {
  return Utilities.formatDate(date, TIMEZONE_ASIA_TOKYO, 'HH:mm')
}

function readParametersFromSpreadsheet(config) {
  const ss = SpreadsheetApp.openById(config.spreadsheetId);
  const calendarIdSeparator = ',';

  return {
    calendarIds: getValueFromCell(ss, config.calendarIdCell).split(calendarIdSeparator),
    businessStartTime: new Date(getValueFromCell(ss, config.businessStartTimeCell)),
    businessEndTime: new Date(getValueFromCell(ss, config.businessEndTimeCell)),
    interviewStartTime: new Date(getValueFromCell(ss, config.interviewStartTimeCell)),
    interviewEndTime: new Date(getValueFromCell(ss, config.interviewEndTimeCell)),
    minMinutes: getValueFromCell(ss, config.minMinutesCell),
  }
}

function mergeBusyTimes(times) {
  const result = [];

  for (let i = 0; i < times.length - 1; i++) {
    const [currStart, currEnd] = times[i];
    const [nextStart, nextEnd] = times[i + 1];
    const isNotOverlap = currStart < nextStart && currEnd < nextStart;

    if (isNotOverlap) {
      result.push([currStart, currEnd])
    } else {
      times[i + 1] = [getMinDate(currStart, nextStart), getMaxDate(currEnd, nextEnd)];
    }

    const isLastLoop = i + 1 === times.length - 1;
    if (isLastLoop) result.push(times[i + 1])
  }

  return result
}

function listFreeTimes(calendarId, interviewStartTime, interviewEndTime) {
  const result = [];

  let busyTimes;
  busyTimes = listScheduleTimes(calendarId, interviewStartTime, interviewEndTime);
  busyTimes = removeAllDaySchedule(busyTimes);
  busyTimes = mergeBusyTimes(busyTimes);

  if (busyTimes.length === 0) {
    result.push([interviewStartTime, interviewEndTime]);
    return result;
  }

  busyTimes = [[null, interviewStartTime], ...busyTimes, [interviewEndTime, null]];

  for (let i = 0; i < busyTimes.length - 1; i++) {
    const currBusyTo = busyTimes[i][1];
    const nextBusyFrom = busyTimes[i + 1][0];
    const invalidEvent = currBusyTo === null || nextBusyFrom === null;
    const isOverlap = currBusyTo >= nextBusyFrom;
    if (invalidEvent || isOverlap) continue;
    result.push([currBusyTo, nextBusyFrom]);
  }

  return result;
}

function getIntersecFreeTimes(timesA, timesB) {
  const result = [];

  timesA.forEach(([startTimeA, endTimeA]) => {
    timesB.forEach(([startTimeB, endTimeB]) => {
      if ((endTimeA <= startTimeB) || (endTimeB <= startTimeA)) return
      const time = [getMaxDate(startTimeA, startTimeB), getMinDate(endTimeA, endTimeB)];
      result.push(time)
    })
  })

  return result;
}

function splitFreeTimesByDay(freeTimes) {
  const result = [];

  freeTimes.forEach(([startAt, endAt]) => {
    if (isSameDay(startAt, endAt)) {
      result.push([startAt, endAt])
      return
    }

    let start = copyDate(startAt);
    const daysDiff = getDaysDiff(startAt, endAt);

    [...Array(daysDiff + 1).keys()].forEach(_ => {
      const endOfDay = toEndOfDay(start);
      const end = endOfDay <= endAt ? endOfDay : endAt;
      result.push([start, end])
      if (endOfDay <= endAt) start = toStartOfNextDay(start);
    })
  })

  return result;
}

function filterWithInBusinessTime(freeTimes, businessStartTime, businessEndTime, minMinutes) {
  const result = [];
  let times = [];

  freeTimes.forEach(([startTime, endTime]) => {
    if (isHoliday(startTime)) return

    const businessStart = makeBusinessDate(startTime, businessStartTime);
    const businessEnd = makeBusinessDate(startTime, businessEndTime);
    const beforeBusinessHours = (startTime < businessStart) && (endTime <= businessStart);
    const afterBusinessHours = (startTime >= businessEnd) && (endTime > businessEnd);

    if (beforeBusinessHours) return
    if (afterBusinessHours) {
      if (times.length > 0) result.push(times)
      times = [];
      return
    }

    const start = getMaxDate(businessStart, startTime);
    const end = endTime < businessEnd ? endTime : businessEnd;

    if (minutesIsMoreThan(start, end, minMinutes)) times.push([start, end])
    if (endTime >= businessEnd && times.length > 0) {
      result.push(times)
      times = [];
    }
  })

  if (times.length > 0) result.push(times)

  return result
}

function formatTimes(freeTimesByDay) {
  const result = [];

  freeTimesByDay.forEach(freeTimes => {
    const times = [];
    freeTimes.forEach(([startTime, endTime], index) => {
      const [start, end] = index === 0 ? [formatHeadTime(startTime), formatTime(endTime)] : [startTime, endTime].map(formatTime)
      times.push(`${start}〜${end}まで`)
    })
    result.push(times.join('、'))
  })

  return result
}

function searchFreeTimeFromCalendar(parameters) {
  const freeTimesArray = parameters.calendarIds.map(calendarId => {
    return listFreeTimes(calendarId, parameters.interviewStartTime, parameters.interviewEndTime)
  })

  const freeTimes = freeTimesArray.length > 1 ? freeTimesArray.reduce(getIntersecFreeTimes) : freeTimesArray[0];
  const freeTimesByDay = splitFreeTimesByDay(freeTimes);
  const freeTimesInBusinessTime = filterWithInBusinessTime(freeTimesByDay, parameters.businessStartTime, parameters.businessEndTime, parameters.minMinutes);

  return formatTimes(freeTimesInBusinessTime).join('\n')
}

// ============================
// Triggered
// ============================

function main() {
  const inputConfig = {
    spreadsheetId: 'dummy',
    calendarIdCell: 'test!C2',
    businessStartTimeCell: 'test!C3',
    businessEndTimeCell: 'test!C4',
    interviewStartTimeCell: 'test!C5',
    interviewEndTimeCell: 'test!C6',
    minMinutesCell: 'test!C7'
  }

  const outputConfig = {
    spreadsheetId: 'dummy',
    cell: 'test!C8',
  }

  const parameters = readParametersFromSpreadsheet(inputConfig);
  const freeTimeString = searchFreeTimeFromCalendar(parameters)
  writeValueToSpreadsheet(outputConfig.spreadsheetId, outputConfig.cell, freeTimeString)
}
