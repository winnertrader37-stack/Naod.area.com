/*******************************************************
 * LOCATION SHARING BACKEND
 * Google Apps Script - Code.gs
 *******************************************************/

const SHARE_ID =
  "AKfycbyO6ith0ZzQRKIHD8xfQPhKnSH91w9fAp46LhjC1GWPLHpIAIQN-q4C_3aSs7BK2edW";

const WEB_APP_URL =
  "https://script.google.com/macros/s/AKfycbzPp9A96uilwISd_ftxAh4u_r5o01rQv2WL7ZyePGdOaiLDmhOwRZe49Fnx1sDhnem6ag/exec";

const NOTIFICATION_EMAIL = "winnertrader37@gmail.com";

const SHEET_NAME = "Locations";


/*******************************************************
 * GET REQUEST
 *******************************************************/
function doGet(e) {
  try {
    const params = e && e.parameter ? e.parameter : {};

    const action = params.action || "open";
    const id = params.id || SHARE_ID;

    if (id !== SHARE_ID) {
      return jsonResponse({
        success: false,
        error: "Invalid Share ID"
      });
    }

    if (action === "location") {
      return getLatestLocation_();
    }

    if (action === "status") {
      return getStatus_();
    }

    return HtmlService.createHtmlOutput(
      "<h2>Location Sharing Server</h2>" +
      "<p>Server is running correctly.</p>"
    );

  } catch (error) {
    return jsonResponse({
      success: false,
      error: String(error)
    });
  }
}


/*******************************************************
 * POST REQUEST
 *******************************************************/
function doPost(e) {
  try {
    const data = parseRequest_(e);

    if (data.id !== SHARE_ID) {
      return jsonResponse({
        success: false,
        error: "Invalid Share ID"
      });
    }

    if (data.action === "location") {
      return saveLocation_(data);
    }

    return jsonResponse({
      success: false,
      error: "Unknown action"
    });

  } catch (error) {
    return jsonResponse({
      success: false,
      error: String(error)
    });
  }
}


/*******************************************************
 * SAVE LOCATION
 *******************************************************/
function saveLocation_(data) {

  const lat = Number(data.latitude);
  const lng = Number(data.longitude);
  const accuracy = Number(data.accuracy || 0);

  if (!isFinite(lat) || !isFinite(lng)) {
    return jsonResponse({
      success: false,
      error: "Invalid latitude or longitude"
    });
  }

  const sheet = getSheet_();

  const now = new Date();

  const device =
    data.device ||
    "Unknown Device";

  const userAgent =
    data.userAgent ||
    "";

  const previous = getLastSavedLocation_();

  sheet.appendRow([
    now,
    SHARE_ID,
    lat,
    lng,
    accuracy,
    device,
    userAgent
  ]);

  /*
   * Send notification only when the first valid
   * location is received for this Share ID.
   */
  if (!previous) {
    sendLocationEmail_(
      lat,
      lng,
      accuracy,
      device,
      now
    );
  }

  PropertiesService
    .getScriptProperties()
    .setProperty(
      "LAST_LOCATION",
      JSON.stringify({
        latitude: lat,
        longitude: lng,
        accuracy: accuracy,
        device: device,
        userAgent: userAgent,
        time: now.toISOString()
      })
    );

  return jsonResponse({
    success: true,
    message: "Location received",
    latitude: lat,
    longitude: lng,
    accuracy: accuracy,
    time: now.toISOString()
  });
}


/*******************************************************
 * GET LATEST LOCATION
 *******************************************************/
function getLatestLocation_() {

  const props =
    PropertiesService.getScriptProperties();

  const raw =
    props.getProperty("LAST_LOCATION");

  if (!raw) {
    return jsonResponse({
      success: true,
      available: false
    });
  }

  let location;

  try {
    location = JSON.parse(raw);
  } catch (e) {
    return jsonResponse({
      success: true,
      available: false
    });
  }

  return jsonResponse({
    success: true,
    available: true,
    location: location
  });
}


/*******************************************************
 * STATUS
 *******************************************************/
function getStatus_() {

  const props =
    PropertiesService.getScriptProperties();

  const raw =
    props.getProperty("LAST_LOCATION");

  if (!raw) {
    return jsonResponse({
      success: true,
      online: false
    });
  }

  let location;

  try {
    location = JSON.parse(raw);
  } catch (e) {
    return jsonResponse({
      success: true,
      online: false
    });
  }

  const lastTime =
    new Date(location.time).getTime();

  const now =
    Date.now();

  const age =
    now - lastTime;

  /*
   * User is considered online when the last update
   * was received within 20 seconds.
   */
  const online =
    age <= 20000;

  return jsonResponse({
    success: true,
    online: online,
    lastUpdate: location.time
  });
}


/*******************************************************
 * SEND EMAIL
 *******************************************************/
function sendLocationEmail_(
  lat,
  lng,
  accuracy,
  device,
  time
) {

  const mapUrl =
    "https://www.google.com/maps?q=" +
    encodeURIComponent(lat + "," + lng);

  const subject =
    "📍 New Location Shared";

  const body =
    "A user has allowed location sharing.\n\n" +

    "Share ID:\n" +
    SHARE_ID +
    "\n\n" +

    "Latitude:\n" +
    lat +
    "\n\n" +

    "Longitude:\n" +
    lng +
    "\n\n" +

    "Accuracy:\n" +
    accuracy +
    " meters\n\n" +

    "Device:\n" +
    device +
    "\n\n" +

    "Time:\n" +
    time +
    "\n\n" +

    "Google Maps:\n" +
    mapUrl +
    "\n\n" +

    "Apps Script:\n" +
    WEB_APP_URL;

  MailApp.sendEmail({
    to: NOTIFICATION_EMAIL,
    subject: subject,
    body: body
  });
}


/*******************************************************
 * GET SHEET
 *******************************************************/
function getSheet_() {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  let sheet =
    ss.getSheetByName(SHEET_NAME);

  if (!sheet) {

    sheet =
      ss.insertSheet(SHEET_NAME);

    sheet.appendRow([
      "Timestamp",
      "Share ID",
      "Latitude",
      "Longitude",
      "Accuracy",
      "Device",
      "User Agent"
    ]);
  }

  return sheet;
}


/*******************************************************
 * LAST SAVED LOCATION
 *******************************************************/
function getLastSavedLocation_() {

  const props =
    PropertiesService.getScriptProperties();

  const raw =
    props.getProperty("LAST_LOCATION");

  if (!raw) {
    return null;
  }

  try {
    return JSON.parse(raw);
  } catch (e) {
    return null;
  }
}


/*******************************************************
 * PARSE REQUEST
 *******************************************************/
function parseRequest_(e) {

  if (!e) {
    return {};
  }

  if (e.postData && e.postData.contents) {

    const content =
      e.postData.contents;

    try {
      return JSON.parse(content);
    } catch (error) {
      // Continue with normal parameters
    }
  }

  return e.parameter || {};
}


/*******************************************************
 * JSON RESPONSE
 *******************************************************/
function jsonResponse(data) {

  return ContentService
    .createTextOutput(
      JSON.stringify(data)
    )
    .setMimeType(
      ContentService.MimeType.JSON
    );
}


/*******************************************************
 * TEST EMAIL
 *
 * Run this function manually once from Apps Script
 * to authorize MailApp.
 *******************************************************/
function testEmail() {

  MailApp.sendEmail({
    to: NOTIFICATION_EMAIL,
    subject: "Location Sharing Test",
    body:
      "Your location-sharing email notification " +
      "system is working correctly."
  });

}
