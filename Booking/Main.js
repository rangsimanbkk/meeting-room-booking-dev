const sheetId = "xxxxxxxxxxxxxxxxxxxxxxx"; //Replace with your Google Sheet ID
const sheetName = "xxxxxxxxxxxxxxx"; //Replace with your Google Sheet name
const sheet = SpreadsheetApp.openById(sheetId).getSheetByName(sheetName);
const lastRow = sheet.getLastRow(); //Get the last row with data in the sheet

//Set data for Google Calendar you can adjust the column number as your form structure, currently it is set for column A to J (1-10)
const date = sheet.getRange(lastRow, 1).getValue(); //Get date from column A
const section = sheet.getRange(lastRow, 2).getValue(); //Get section from column B
const topic = sheet.getRange(lastRow, 3).getValue(); //Get topic from column C
const room = sheet.getRange(lastRow, 4).getValue(); //Get room from column D
const startTime = sheet.getRange(lastRow, 5).getDisplayValue(); //Get start time from column E
const endTime = sheet.getRange(lastRow, 6).getDisplayValue(); //Get end time from column F
const item = sheet.getRange(lastRow, 7).getValue(); //Get item from column G
const bookedName = sheet.getRange(lastRow, 8).getValue(); //Get booked person from column H
const status = sheet.getRange(lastRow, 10).getValue(); //Get status from column J
let adDate = date;
const roomCalendars = {
  // Define your room calendar IDs here
  "room1Name": "calendarIdRoom1@group.calendar.google.com", //Replace with your Room name and Google Calendar Id
  "room2Name": "calendarIdRoom2@group.calendar.google.com", //Replace with your Room name and Google Calendar Id
  // You can add more rooms as you needed
};
const calendarId = roomCalendars[room];

//Function to generate unique booking code for managing bookings afterwards, such as cancellation or modification.
function generateUniqueCode(sheet, length) {
  const existingCodes = sheet.getRange("K2:K" + sheet.getLastRow()).getValues().flat(); //Set Column K for storing booking codes
  const chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789";
  do {
    var code = "";
    for (let i = 0; i < length; i++) {
      code += chars.charAt(Math.floor(Math.random() * chars.length));
    }
  } while (existingCodes.includes(code));
  return code;
}

//Function to convert time string to Date object
function convertTimeToDateObject(formDate, formTime) {
  const timePart = formTime.split(":");
  Logger.log("timePart: " + timePart);
  //use parseInt convert string to number, first parameter is the string to convert, second parameter is the radix (base)
  const hours = parseInt(timePart[0], 10);
  const minutes = parseInt(timePart[1], 10);
  let formatTime = new Date(formDate);
  formatTime.setHours(hours);
  formatTime.setMinutes(minutes);
  Logger.log("formatTime: " + formatTime);
  return formatTime;
}

//Function to format time for message sending
function formatTime(dateObj) {
  return Utilities.formatDate(dateObj, Session.getScriptTimeZone(), "HH:mm");
}

//Test Function for checking if the data is correctly retrieved from the sheet and the last row is correct, you can run this function in the Apps Script editor to see the logs.
function Test_Last_Row() {
  Logger.log(lastRow);
  Logger.log(date);
  Logger.log(topic);
}

//Main Function
function meetingRoomBooking() {
  //Validate required fields
  if (!room || !date || !startTime || !endTime) {
    console.error(`ข้อมูลไม่ครบถ้วน`);
    return;
  }

  const rawYear = date.getFullYear();
  let convertedYear = rawYear;

  //Convert 2 digit year to 4 digit year and convert BE to AD if necessary
  if (rawYear < 100) {
    //2 Digit Year
    if (rawYear >= 68) {
      convertedYear = rawYear + 1957; //2 Digit BE to 4 Digit AD
    } else if (rawYear >= 25) {
      convertedYear = rawYear + 2000; //2 Digit AD to 4 Digit AD
    }
  } else {
    //4 Digit Year
    if (rawYear >= 2568) {
      convertedYear = rawYear - 543; //BE to AD
    } else if (rawYear >= 2025) {
      convertedYear = rawYear; //Still AD
    }
  }

  adDate = new Date(date);
  adDate.setFullYear(convertedYear); //Update the year of the date object to the converted year
  sheet.getRange(lastRow, 1).setValue(adDate); //Update adjusted date to the sheet
  const formattedDate = Utilities.formatDate(adDate, Session.getScriptTimeZone(), "dd/MM/yyyy"); //Format date for message sending

  const eventStart = convertTimeToDateObject(adDate, startTime); //Assign start time to eventStart variable and convert it to Date object
  const formattedStartTime = formatTime(eventStart);
  Logger.log("eventStart: " + eventStart);

  const eventEnd = convertTimeToDateObject(adDate, endTime); //Assign end time to eventEnd variable and convert it to Date object
  const formattedEndTime = formatTime(eventEnd);
  Logger.log("eventEnd: " + eventEnd);

  const bookingCode = generateUniqueCode(sheet, 6); //Generate a unique booking code with 6 characters, you can adjust the length as needed
  const calendar = CalendarApp.getCalendarById(calendarId); //Get the calendar by ID based on the room selected
  const titles = topic + " (" + section + ")"; //Set the event title with topic and section, you can adjust the format as needed
  const descriptions = "ผู้จอง : " + bookedName + "\nอุปกรณ์ที่ใช้ : " + item + "\nBooking Code : " + bookingCode; //Set the event description with booked name, item and booking code, you can adjust the format as needed
  let events = calendar.getEvents(eventStart, eventEnd); //Get events from the calendar in the specified time range to check for conflicts
  let message;

  //Set message templates for different scenarios, you can adjust the content and format as needed
  const message1 = `
    📝 การจองใหม่
    เรื่อง : ${topic}
    ฝ่าย : ${section}
    ห้อง : ${room}
    วันที่ : ${formattedDate}
    เวลาเริ่ม : ${formattedStartTime} น.
    เวลาสิ้นสุด : ${formattedEndTime} น.
    อุปกรณ์ที่ใช้ : ${item}
    ผู้จอง : ${bookedName}
    Booking Code : ${bookingCode}
    `;

  const message2 = `
    ❌❌ มีการจองซ้ำ
    เรื่อง : ${topic}
    ฝ่าย : ${section}
    ห้อง : ${room}
    วันที่ : ${formattedDate}
    เวลาเริ่ม : ${formattedStartTime} น.
    เวลาสิ้นสุด : ${formattedEndTime} น.
    ผู้จอง : ${bookedName}
    `;

  const message3 = `
    ⚠️⚠️ จองไม่สำเร็จ
    เรื่อง : ${topic}
    ฝ่าย : ${section}
    ห้อง : ${room}
    วันที่ : ${formattedDate}
    เวลาเริ่ม : ${formattedStartTime} น.
    เวลาสิ้นสุด : ${formattedEndTime} น.
    ผู้จอง : ${bookedName}
    `;

  if (events.length > 0) { //check if there are any events in the specified time range, if there are, it means there is a conflict and the room is already booked
    console.log(`${room} ถูกจองแล้วในช่วงเวลานี้.`); //Log the conflict in the console for debugging purposes
    sheet.getRange(lastRow, 10).setValue("การจองซ้ำ"); //Update the status to column J
    message = message2; //Set the message to the booking conflict template

    //uncomment below to send message to Line
    //sendMessageToLine(message);

    //uncomment below to send message to Telegram
    //sendMessageToTelegram(message);

    //uncomment below to send message to Discord
    sendMessageToDiscord(message);

  } else {
    //(Event title, Start time, End time, description)
    calendar.createEvent(titles, eventStart, eventEnd, { description: descriptions, location: room }); //Create the event in the calendar, you can adjust the parameters as needed
    if (events) { //check if the event is created successfully
      sheet.getRange(lastRow, 10).setValue("ลงจองแล้ว"); //Update the status to column J
      //Logger.log(`Booking Code: ${bookingCode}`);
      sheet.getRange(lastRow, 11).setValue(bookingCode); //Store the booking code in column K
      message = message1; //Set the message to the booking confirmation template

      //uncomment below to send message to Line
      //sendMessageToLine(message);

      //uncomment below to send message to Telegram
      //sendMessageToTelegram(message);

      //uncomment below to send message to Discord
      sendMessageToDiscord(message);

    } else {
      sheet.getRange(lastRow, 10).setValue("จองไม่สำเร็จ"); //Update the status to column J
      message = message3; //Set the message to the booking failure template

      //uncomment below to send message to Line
      //sendMessageToLine(message);

      //uncomment below to send message to Telegram
      //sendMessageToTelegram(message);

      //uncomment below to send message to Discord
      sendMessageToDiscord(message);
    }
  }
}


function sendMessageToLine(message) {
  const lineAccessToken = "xxxxxxxxxxxxxxxxxxxxxxxx"; // Replace with your Channel Access Token
  const lineGroupId = "xxxxxxxxxxxxxxxxxx"; // Replace with your Line group ID

  const url = "https://api.line.me/v2/bot/message/push";

  const payload = {
    to: lineGroupId,
    messages: [
      {
        type: "text",
        text: message.trim(), // Trim excess whitespace
      },
    ],
  };

  const options = {
    method: "post",
    headers: {
      "Content-Type": "application/json",
      Authorization: `Bearer ${lineAccessToken}`,
    },
    payload: JSON.stringify(payload),
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    console.log(`Message sent: ${response.getContentText()}`);
  } catch (error) {
    console.error(`Error sending message: ${error}`);
  }
}

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

function sendMessageToDiscord(messageText) {
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