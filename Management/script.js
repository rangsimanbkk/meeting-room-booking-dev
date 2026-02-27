// --- CONFIGURATION ---
const sheetId = "xxxxxxxxxxxxxxxxxxxxxxx"; //Replace with your Google Sheet ID
const sheetName = "xxxxxxxxxxxxxxx"; //Replace with your Google Sheet name
const LOG_SHEET_NAME = "log";// Your log sheet name

const CALENDAR_IDS = {// Calendar IDs for each meeting room
  "room1Name": "calendarIdRoom1@group.calendar.google.com", //Replace with your Room name and Google Calendar Id
  "room2Name": "calendarIdRoom2@group.calendar.google.com", //Replace with your Room name and Google Calendar Id
};

// Column headers in the Google Sheet sorted by column index
const HEADERS = ['Date', 'Section', 'Topic', 'Room', 'StartTime', 'EndTime', 'Items', 'Name', 'Timestamp', 'BookingStatus', 'BookingCode'];
//const HEADERS = ['Timestamp', 'Section', 'Topic', 'Room', 'Date', 'StartTime', 'EndTime', 'Items', 'Name', 'BookingStatus', 'BookingCode'];//For Testing

// --- WEB APP ---
function doGet(e) {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Meeting Room Booking Management')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// --- DATA FUNCTIONS ---

/**
 * Finds a booking by its code and returns its data.
 * @param {string} bookingCode The unique code for the booking.
 * @returns {object|null} An object with booking data or null if not found.
 */
function getBookingData(bookingCode) {
  if (!bookingCode) return null;

  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  const codeColumnIndex = HEADERS.indexOf('BookingCode');

  for (let i = 1; i < data.length; i++) {
    if (data[i][codeColumnIndex] == bookingCode) {
      const booking = {};
      HEADERS.forEach((header, index) => {
        booking[header] = data[i][index]; // Assign raw value first
      });

      // --- Robust Date Processing ---
      const dateValue = booking.Date;
      if (dateValue) {
        let dateObj;
        if (dateValue instanceof Date) {
          dateObj = dateValue;
        } else { // It's a string, likely "DD-MM-YYYY" from previous edits
          const parts = String(dateValue).split('-');
          if (parts.length === 3) {
            // new Date(year, monthIndex, day). Month is 0-indexed.
            dateObj = new Date(parts[2], parts[1] - 1, parts[0]);
          } else {
             // Fallback for other formats or invalid string
             dateObj = new Date(dateValue);
          }
        }
        
        // Ensure dateObj is valid before formatting
        if (dateObj && !isNaN(dateObj)) {
            // Format for the HTML <input type="date"> value (YYYY-MM-DD)
            booking.htmlDate = Utilities.formatDate(dateObj, Session.getScriptTimeZone(), "yyyy-MM-dd");
            // Format for display in the results table (DD-MM-YYYY)
            booking.Date = Utilities.formatDate(dateObj, Session.getScriptTimeZone(), "dd-MM-yyyy");
        }
      }

      // --- Process Other Date/Time Fields ---
      if (booking.Timestamp instanceof Date) {
          booking.Timestamp = Utilities.formatDate(booking.Timestamp, Session.getScriptTimeZone(), "dd-MM-yyyy HH:mm:ss");
      }
      if (booking.StartTime instanceof Date) {
          booking.StartTime = Utilities.formatDate(booking.StartTime, Session.getScriptTimeZone(), "HH:mm");
      }
      if (booking.EndTime instanceof Date) {
          booking.EndTime = Utilities.formatDate(booking.EndTime, Session.getScriptTimeZone(), "HH:mm");
      }
      
      booking.row = i + 1; // Add row number for easier updates
      return booking;
    }
  }
  return null; // Not found
}


/**
 * Checks if a specific time slot is available in a calendar, excluding a specified event.
 * @param {string} calendarId The ID of the calendar to check.
 * @param {Date} startTime The start time of the potential new event.
 * @param {Date} endTime The end time of the potential new event.
 * @param {string} excludedBookingCode The booking code of the event to ignore (for updates).
 * @returns {boolean} True if the time slot is available, false otherwise.
 */
function isTimeSlotAvailable(calendarId, startTime, endTime, excludedBookingCode) {
  if (!calendarId) return false;

  const calendar = CalendarApp.getCalendarById(calendarId);
  if (!calendar) return false;

  const events = calendar.getEvents(startTime, endTime);

  for (const event of events) {
    const eventDescription = event.getDescription();
    if (eventDescription && eventDescription.includes(excludedBookingCode)) {
      // This is the event we are trying to update, so it's not a conflict.
      continue;
    }
    // If the event title is not empty, it is considered a conflict.
    if (event.getTitle()) {
      return false; // Time slot is not available.
    }
  }
  return true; // Time slot is available.
}

/**
 * Deletes a booking from the sheet and calendar.
 * @param {string} bookingCode The code of the booking to delete.
 * @param {string} requesterName The name of the person requesting the deletion.
 * @returns {string} A success message.
 */
function deleteBooking(bookingCode, requesterName) {
  const bookingData = getBookingData(bookingCode);
  if (!bookingData) {
    throw new Error("Booking not found.");
  }
  
  // Use the correctly formatted htmlDate for creating the Date object
  const startTime = new Date(`${bookingData.htmlDate}T${bookingData.StartTime}`);
  const endTime = new Date(`${bookingData.htmlDate}T${bookingData.EndTime}`);

  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);

  // --- Delete Calendar Event ---
  const calendarId = CALENDAR_IDS[bookingData.Room];
  if (calendarId) {
    const calendar = CalendarApp.getCalendarById(calendarId);
    const events = calendar.getEvents(startTime, endTime, { search: bookingData.Topic });

    for (const event of events) {
      if (event.getDescription() && event.getDescription().includes(bookingCode)) {
        event.deleteEvent();
        let messageCancel = `
        ⛔️ มีการยกเลิกการจอง
        เรื่อง : ${bookingData.Topic}
        ฝ่าย : ${bookingData.Section}
        ห้อง : ${bookingData.Room}
        วันที่ : ${bookingData.Date.split('-').reverse().join('/')}
        เวลาเริ่ม : ${bookingData.StartTime} น.
        เวลาสิ้นสุด : ${bookingData.EndTime} น.
        ผู้จอง : ${bookingData.Name}
        ผู้ยกเลิก : ${requesterName}
        Booking Code : ${bookingData.BookingCode}
          `;
        sendMessageToTelegram(messageCancel);
        sendDiscordMessage(messageCancel);
        break;
      }
    }
  }

  // --- Update Google Sheet ---
  const row = bookingData.row;
  sheet.getRange(row, HEADERS.indexOf('BookingStatus') + 1).setValue("ยกเลิกการจอง");
  //sheet.deleteRow(bookingData.row);

  // --- Log Action ---
  const logDetails = `Deleted booking for ${bookingData.Topic} on ${bookingData.Date}.`;
  logAction('DELETE', bookingCode, requesterName, logDetails);

  return `Booking ${bookingCode} deleted successfully.`;
}

/**
 * Updates a booking in the sheet and calendar.
 * @param {object} formData The updated booking data from the web app form.
 * @returns {string} A success message.
 */
function updateBooking(formData) {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
  const originalBooking = getBookingData(formData.BookingCode);

  if (!originalBooking) {
    throw new Error("Booking not found. Cannot update.");
  }
  
  const oldStartTime = new Date(`${originalBooking.htmlDate}T${originalBooking.StartTime}`);
  const oldEndTime = new Date(`${originalBooking.htmlDate}T${originalBooking.EndTime}`);
  const newStartTime = new Date(`${formData.Date}T${formData.StartTime}`);
  const newEndTime = new Date(`${formData.Date}T${formData.EndTime}`);

  // Check if the room has changed
  if (formData.Room !== originalBooking.Room) {
    // Scenario 2: Room has changed
    const newCalendarId = CALENDAR_IDS[formData.Room];
    if (!isTimeSlotAvailable(newCalendarId, newStartTime, newEndTime, "")) {
      throw new Error("The new time slot is not available in the selected room.");
    }

    // 1. Delete the old event from the original calendar
    const oldCalendarId = CALENDAR_IDS[originalBooking.Room];
    if (oldCalendarId) {
      const oldCalendar = CalendarApp.getCalendarById(oldCalendarId);
      const events = oldCalendar.getEvents(oldStartTime, oldEndTime, { search: originalBooking.BookingCode });
      if (events.length > 0) {
        events[0].deleteEvent();
      }
    }

    // 2. Create a new event in the new calendar
    if (newCalendarId) {
      const newCalendar = CalendarApp.getCalendarById(newCalendarId);
      const newEventTitle = `${formData.Topic} (${formData.Section})`;
      const newEventDescription = `ผู้จอง : ${formData.Name}\nอุปกรณ์ที่ใช้ : ${formData.Items}\nBooking Code : ${formData.BookingCode}`;
      newCalendar.createEvent(newEventTitle, newStartTime, newEndTime, {
        description: newEventDescription,
        location: formData.Room
      });
    }
    let messageChangeRoom = `
    ⚠️ มีการแก้ไขการจอง
    ◀️จาก
    เรื่อง : ${originalBooking.Topic}
    ฝ่าย : ${originalBooking.Section}
    ห้อง : ${originalBooking.Room}
    วันที่ : ${originalBooking.Date.split('-').reverse().join('/')}
    เวลาเริ่ม : ${originalBooking.StartTime} น.
    เวลาสิ้นสุด : ${originalBooking.EndTime} น.
    อุปกรณ์ที่ใช้ : ${originalBooking.Items}
    ผู้จอง : ${originalBooking.Name}
    Booking Code : ${originalBooking.BookingCode}
    ---------------------------------
    ▶️เป็น
    เรื่อง : ${formData.Topic}
    ฝ่าย : ${formData.Section}
    ห้อง : ${formData.Room}
    วันที่ : ${formData.Date.split('-').reverse().join('/')}
    เวลาเริ่ม : ${formData.StartTime} น.
    เวลาสิ้นสุด : ${formData.EndTime} น.
    อุปกรณ์ที่ใช้ : ${formData.Items}
    ผู้จอง : ${formData.Name}
    ผู้แก้ไข : ${formData.requesterName}
    Booking Code : ${formData.BookingCode}
      `;
    sendMessageToTelegram(messageChangeRoom);
    sendDiscordMessage(messageChangeRoom);

  } else {
    // Scenario 1: Room has not changed
    const calendarId = CALENDAR_IDS[originalBooking.Room];
    if (!isTimeSlotAvailable(calendarId, newStartTime, newEndTime, originalBooking.BookingCode)) {
      throw new Error("The new time slot is not available.");
    }

    const calendar = CalendarApp.getCalendarById(calendarId);
    const eventsToDelete = calendar.getEvents(oldStartTime, oldEndTime, { search: originalBooking.BookingCode });
    if (eventsToDelete.length > 0) {
      eventsToDelete[0].deleteEvent();
    }

    const newEventTitle = `${formData.Topic} (${formData.Section})`;
    const newEventDescription = `ผู้จอง : ${formData.Name}\nอุปกรณ์ที่ใช้ : ${formData.Items}\nBooking Code : ${formData.BookingCode}`;
    calendar.createEvent(newEventTitle, newStartTime, newEndTime, {
      description: newEventDescription,
      location: formData.Room
    });
    let messageNotChangeRoom = `
    ⚠️ มีการแก้ไขการจอง
    ◀️จาก
    เรื่อง : ${originalBooking.Topic}
    ฝ่าย : ${originalBooking.Section}
    ห้อง : ${originalBooking.Room}
    วันที่ : ${originalBooking.Date.split('-').reverse().join('/')}
    เวลาเริ่ม : ${originalBooking.StartTime} น.
    เวลาสิ้นสุด : ${originalBooking.EndTime} น.
    อุปกรณ์ที่ใช้ : ${originalBooking.Items}
    ผู้จอง : ${originalBooking.Name}
    Booking Code : ${originalBooking.BookingCode}
    ---------------------------------
    ▶️เป็น
    เรื่อง : ${formData.Topic}
    ฝ่าย : ${formData.Section}
    ห้อง : ${formData.Room}
    วันที่ : ${formData.Date.split('-').reverse().join('/')}
    เวลาเริ่ม : ${formData.StartTime} น.
    เวลาสิ้นสุด : ${formData.EndTime} น.
    อุปกรณ์ที่ใช้ : ${formData.Items}
    ผู้จอง : ${formData.Name}
    ผู้แก้ไข : ${formData.requesterName}
    Booking Code : ${formData.BookingCode}
      `;
    sendMessageToTelegram(messageNotChangeRoom);
    sendDiscordMessage(messageNotChangeRoom);
  }

  // --- Update Google Sheet ---
  const displayDate = formData.Date.split('-').reverse().join('-');
  const row = originalBooking.row;
  sheet.getRange(row, HEADERS.indexOf('Section') + 1).setValue(formData.Section);
  sheet.getRange(row, HEADERS.indexOf('Topic') + 1).setValue(formData.Topic);
  sheet.getRange(row, HEADERS.indexOf('Room') + 1).setValue(formData.Room);
  sheet.getRange(row, HEADERS.indexOf('Date') + 1).setValue(displayDate);
  sheet.getRange(row, HEADERS.indexOf('StartTime') + 1).setValue(formData.StartTime);
  sheet.getRange(row, HEADERS.indexOf('EndTime') + 1).setValue(formData.EndTime);
  sheet.getRange(row, HEADERS.indexOf('Items') + 1).setValue(formData.Items);
  sheet.getRange(row, HEADERS.indexOf('Name') + 1).setValue(formData.Name);
  sheet.getRange(row, HEADERS.indexOf('BookingStatus') + 1).setValue("มีการแก้ไข");


  // --- Log Action ---
  const logDetails = `Updated fields for booking ${formData.BookingCode}.`;
  logAction('EDIT', formData.BookingCode, formData.requesterName, logDetails);

  return `Booking ${formData.BookingCode} updated successfully.`;
}


/**
 * Logs an action to the 'log' sheet.
 * @param {string} action The action taken (e.g., 'EDIT', 'DELETE').
 * @param {string} bookingCode The affected booking code.
 * @param {string} requesterName The person performing the action.
 * @param {string} details A description of the action.
 */
function logAction(action, bookingCode, requesterName, details) {
  try {
    const logSheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(LOG_SHEET_NAME);
    if (!logSheet) { // Create log sheet if it doesn't exist
      const newSheet = SpreadsheetApp.openById(SHEET_ID).insertSheet(LOG_SHEET_NAME);
      newSheet.appendRow(['Log Timestamp', 'Action', 'Booking Code', 'Requester', 'Details']);
    }
    logSheet.appendRow([new Date(), action, bookingCode, requesterName, details]);
  } catch (e) {
    console.error("Could not write to log sheet: " + e.toString());
  }
}

// --- TEST FUNCTION ---
function testLoadDataByCode() {
  const bookingCode = "TOG90A"; // Use the test code provided
  const data = getBookingData(bookingCode);
  Logger.log(data);
}

function testDeleteDataByCode() {
  const bookingCode = "AXIC5T"; // Use the test code provided
  const data = getBookingData(bookingCode);
  Logger.log(data);
  deleteBooking(bookingCode, "rangsiman")
}
// --- TEST FUNCTION ---

// --- Message FUNCTION ---
function sendMessageToTelegram(message) {
  const botToken = "xxxxxxxxxxxxxxxxxxxx"; // Replace with your Telegram bot token
  const chatId = "-xxxxxxxxxxxx"; // Replace with your Telegram group chat ID (include the "-" if it starts with one)

  const url = `https://api.telegram.org/bot${botToken}/sendMessage`;

  const payload = {
    chat_id: chatId,
    text: message.trim(), // Trim excess whitespace
    parse_mode: "HTML", // Optional: Allows HTML formatting
  };

  const options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    console.log(`Message sent: ${response.getContentText()}`);
  } catch (error) {
    console.error(`Error sending message: ${error}`);
  }
}

function sendDiscordMessage(messageText) {
  const webhookUrl = 'https://discordapp.com/api/webhooks/xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx'; // Replace with your webhook URL

  const payload = {
    content: messageText,
    username: "Pathumwan Meeting Room Bot" //Edit the bot name that will appear in Discord
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload)
  };

  try {
    const response = UrlFetchApp.fetch(webhookUrl, options);
    Logger.log("Message sent successfully: " + response.getContentText());
  } catch (error) {
    Logger.log("Failed to send message: " + error);
  }
}
// --- Message FUNCTION ---