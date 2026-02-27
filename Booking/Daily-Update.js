function DailyUpdateMeetingRoomBooking() {
    var calendarID1 = "calendarIdRoom1@group.calendar.google.com"; // Replace with your Google Calendar ID
    var calendarID2 = "calendarIdRoom2@group.calendar.google.com"; // Replace with your Google Calendar ID
    var calendar1 = CalendarApp.getCalendarById(calendarID1);
    var calendar2 = CalendarApp.getCalendarById(calendarID2);
    var today = new Date(); // Get today's date
    var customDate = new Date(today.getFullYear(),today.getMonth(),today.getDate()+2); //For update any date, change the +2 to the number of days you want to check (e.g., +1 for tomorrow, +3 for the day after tomorrow, etc.)
    var events1 = calendar1.getEventsForDay(today); // Get events in calendar 1 for today and change today to customDate if you want to check for any date
    //Logger.log(events1);
    var events2 = calendar2.getEventsForDay(today); // Get events in calendar 2 for today and change today to customDate if you want to check for any date
    //Logger.log(events2);

    // Create the message to send
    var message = "📅 รายการจองห้องประชุมสำหรับวันนี้\n";
    message += "\nห้องประชุม 1 :\n";
    for (var i = 0; i < events1.length; i++) {
      var desc = extractImportantInfo(events1[i].getDescription());
      Logger.log(desc);
      message += `${i + 1}. ${events1[i].getTitle()}  ${formatTime(events1[i].getStartTime())}-${formatTime(events1[i].getEndTime())} น.\n${desc}\n`;
    }

    message += "\nห้องประชุม 2 :\n";
    for (var j = 0; j < events2.length; j++) {
      var desc = extractImportantInfo(events2[j].getDescription());
      Logger.log(desc);
      message += `${j + 1}. ${events2[j].getTitle()}  ${formatTime(events2[j].getStartTime())}-${formatTime(events2[j].getEndTime())} น.\n${desc}\n`;
    }
    
    //Logger.log(message);
    sendMessageToLine(message);
    sendMessageToTelegram(message);
    sendMessageToDiscord(message);
}

function formatTime(time) {
    if (time instanceof Date) { // Check if time is a Date object
        // Format the time as HH:mm
        return Utilities.formatDate(time, Session.getScriptTimeZone(), "HH:mm");
    } else {
        // If time is not a Date object, return it as is
        return time;
    }
}

function extractImportantInfo(description) {
  if (!description) return ""; // If description is empty or null, return an empty string
  
  // Split description into lines
  var lines = description.split('\n');
  
  // Find the line that starts with 'อุปกรณ์ที่ใช้ :'
  for (var i = 0; i < lines.length; i++) {
    if (lines[i].trim().startsWith("อุปกรณ์ที่ใช้ :")) {
      return lines[i].trim(); // Return that full line
    }
  }
  
  return ""; // If not found, return empty string
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