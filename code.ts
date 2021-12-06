import GmailMessage = GoogleAppsScript.Gmail.GmailMessage;

const CALENDAR_NAME = 'English Lesson'

function getCalendar(calendarName: string) {
  const calendars = CalendarApp.getCalendarsByName(calendarName)

  if (!calendars.length) {
    return
  }

  return calendars[0]
}

function getUnreadReservationThread() {
  const threads = GmailApp.search(
    'is:unread from:noreply@eikaiwa.dmm.com subject:"レッスン予約"'
  )

  if (!threads.length) {
    return
  }

  return threads[0]
}

function extractTextsFromMessage(message: GmailMessage) {
  const body = message.getBody()
  const linkToLesson = /https:\/\/eikaiwa.dmm.com\/app\/lesson-booking\/[0-9a-f-]{36}/.exec(body)?.[0].toString() ?? ""

  // XXX様、2021/12/06 22:00のAmandaとのレッスン予約が完了しました。
  const startPhrase = /様、.*の/.exec(body)?.[0].toString() ?? ""
  const startText = startPhrase.match(/\d{4}\/\d{2}\/\d{2} \d{2}:\d{2}/g)?.[0].toString() ?? ""
  const teacherPhrase = /の\w*とのレッスン/.exec(body)?.[0].toString() ?? ""
  const teacher = teacherPhrase.substring(1, teacherPhrase.length - 6)

  return {linkToLesson, startText, teacher}
}

function main() {
  const calendar = getCalendar(CALENDAR_NAME)
  if (!calendar) {
    Logger.log(`Calendar not found: ${CALENDAR_NAME}`)
    return
  }

  const thread = getUnreadReservationThread()
  if (!thread) {
    Logger.log("Reservation thread not found.")
    return
  }

  const messages = thread.getMessages()
  const latestMessage = messages[messages.length - 1]
  const {linkToLesson, startText, teacher} = extractTextsFromMessage(latestMessage)

  const start = new Date(startText)
  const end = new Date(start.getTime() + 30 * 60000)

  calendar.createEvent(teacher, start, end, {location: linkToLesson})
  latestMessage.markRead()
}