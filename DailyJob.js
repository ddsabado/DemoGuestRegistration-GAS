function execute() {
  Logger.log("Execution started");
  sortSheet();
  Logger.log("Sheet sorted");
  
  if (validations()) {
    let reservationDetails = getAllDetails();
    Logger.log("Reservation details: " + JSON.stringify(reservationDetails));
    generateAuthLetter(reservationDetails);
    sendAuthLetter(reservationDetails);
    cleanup(reservationDetails);
  } else {
    Logger.log("Validation failed");
  }
}

/**
 * validation with main purpose of allowing sending letter for next day check-in
 * 
 */
function validations() {
  let validate = getValidationDetails()

  if(validate.checkInDate === null || validate.checkInDate === '') {
    console.log("Missing checkInDate (no reservation or bad data)")
    return false
  }

  if (calculateDaysBetween(validate.checkInDate, Date.now()) > 1) {
    console.log("Too early to send Authorization Letter (check-in is more than 1 day away)")
    return false
  }

  if (calculateDaysBetween(validate.checkInDate, Date.now()) < 0) {
    console.log("Check-in date is in the past; skipping Authorization Letter")
    return false
  }  
  return true
}

function getValidationDetails() {
  const guestRegSS = SpreadsheetApp.openById(REGISTRATION_SPREADSHEET_ID)
  const reservationSumSheet = guestRegSS.getSheetByName(GUEST_LIST_SHEET_NAME)

  let [checkInDate, checkOutDate] = reservationSumSheet.getRange(SHEET_RANGES.checkinCheckout).getValues().flat()

  return {
    checkInDate: checkInDate,
    checkOutDate: checkOutDate
  }
}


function sortSheet() {
  const guestRegSS = SpreadsheetApp.openById(REGISTRATION_SPREADSHEET_ID)
  const formRespSheet = guestRegSS.getSheetByName(GUEST_LIST_SHEET_NAME)

  formRespSheet.sort(3,true)
}

function getAllDetails() {
  const reservation = getReservation();
  reservation.guestList = getGuestList([]);
  return reservation;
}

function getReservation() {
  const guestRegSS = SpreadsheetApp.openById(REGISTRATION_SPREADSHEET_ID)
  const reservationSumSheet = guestRegSS.getSheetByName(GUEST_LIST_SHEET_NAME)
  // const formRespSheet = guestRegSS.getSheetByName(currentConfig.FORM_RESPONSE_SHEET_NAME)

  let [email, checkInDate, checkOutDate] = reservationSumSheet.getRange(SHEET_RANGES.reservation).getValues().flat()
  let reservation = {email, checkInDate, checkOutDate}

  reservation.checkInDate = formatDate(checkInDate);
  reservation.checkOutDate = formatDate(checkOutDate);

  let plateNumber = reservationSumSheet.getRange(SHEET_RANGES.plateNumber).getValue();
  reservation.plateNumber = (plateNumber ? plateNumber.toString().toUpperCase() : "");

  reservation.parking = getParkingShort(reservationSumSheet.getRange(SHEET_RANGES.parking).getValue()) 

  return reservation
}

function getGuestList(guestList) {
  const guestRegSS = SpreadsheetApp.openById(REGISTRATION_SPREADSHEET_ID)
  const guestListSheet = guestRegSS.getSheetByName(GUEST_LIST_SHEET_NAME)

  //guestList = guestListSheet.getRange('B2:E2').getValues()
  GUEST_DETAIL_RANGE.forEach(guestRange => {
    let [name, age, idType, idLink] = guestListSheet.getRange(guestRange).getValues().flat()
    name = name.toUpperCase()
    let guestDetail = {name, age, idType, idLink}
    guestList.push(guestDetail)
  })

  return guestList
}

function generateAuthLetter(reservationDetails) {
  const tempFile = getTempFile();
  const tempDocFile = DocumentApp.openById(tempFile.getId());
  const body = tempDocFile.getBody();

  const placeholders = {
    "{dateTodayLetter}": formatDate(new Date()),
    "{stayDuration}": calculateDaysBetween(reservationDetails.checkOutDate, reservationDetails.checkInDate),
    "{checkincheckoutDate}": formatReservationDateRange(reservationDetails.checkInDate, reservationDetails.checkOutDate),
    "{plateNumber}": reservationDetails.plateNumber,
    "{parking}": reservationDetails.parking,
  };

  // Replace text dynamically
  for (const [placeholder, value] of Object.entries(placeholders)) {
    body.replaceText(placeholder, value);
  }

  reservationDetails.guestList.forEach((guest, index) => {
    body.replaceText(`{guest${index + 1}Name}`, guest.name);
    body.replaceText(`{guest${index + 1}Age}`, guest.age);
    body.replaceText(`{guest${index + 1}IDProof}`, guest.idType);
  });

  tempDocFile.saveAndClose();
  const authLetterPDFBlob = tempFile.getAs(MimeType.PDF);
  savePdfToDrive(authLetterPDFBlob);
}

function getTempFile() {
  const authLetterTemplate = DriveApp.getFileById(AUTH_LETTER_TEMPLATE_FILE);
  const tempFolder = DriveApp.getFolderById(TEMP_FOLDER);
  const tempFile = authLetterTemplate.makeCopy(tempFolder);

  return tempFile
}

function savePdfToDrive(authLetterPDFBlob) {
  DriveApp.getFolderById(FILEDROP_FOLDER_ID).createFile(authLetterPDFBlob).setName(AUTH_LETTER_FILE_NAME);
}

/**
 * sendAuthLetter
 * Gathers details from registration sheet and initiates sending of email
 */
function sendAuthLetter(reservationDetails) {
    let recipient = reservationDetails.email
    let subject = formatReservationDateRange(reservationDetails.checkInDate, reservationDetails.checkOutDate) + " – T2 - 314";
    let attachments = getAttachments(reservationDetails)
    let options = {
      attachments: attachments,
      // bcc: currentConfig.BCC_RECIPIENTS,
      bcc: "ddacsat.business@gmail.com",
      htmlBody: HtmlService.createHtmlOutputFromFile("AuthLetterBody.html").getContent()
    }
    GmailApp.sendEmail(recipient, subject, "", { ...options, from: 'ddacsat.business@gmail.com' })
}

function formatDate(date) {
  return Utilities.formatDate(new Date(date), "GMT+8", "MMMM dd, yyyy");
}

function getParkingShort(parkingOption) {
  
  switch(parkingOption.trim().split(' ')[0]) {
    case "Free":
      return "OFFSITE PARKING"
    case "Basement":
      return "BASEMENT PARKING"
    case "Condo":
      return "CONDO PARKING"
    default:
      return "NO PARKING"
  }
}

function formatReservationDateRange(checkInDate, checkOutDate, { month = 'short', upper = true } = {}) {
  const ci = new Date(checkInDate);
  const co = new Date(checkOutDate);

  // Guard against invalid dates
  if (isNaN(ci) || isNaN(co)) return '';

  const sameYear  = ci.getFullYear() === co.getFullYear();
  const sameMonth = sameYear && ci.getMonth() === co.getMonth();

  const monthFmt = new Intl.DateTimeFormat('en-US', { month }); // 'short' -> "Oct", 'long' -> "October"
  const mon = d => {
    const m = monthFmt.format(d);
    return upper ? m.toUpperCase() : m;
  };
  const dd = d => String(d.getDate()).padStart(2, '0'); // remove padStart if you don't want leading zero

  if (sameMonth) {
    // e.g., "OCT 31–02, 2025" or "OCTOBER 31–NOVEMBER 02, 2025" (if you switch to long)
    return `${mon(ci)} ${dd(ci)}–${dd(co)}, ${ci.getFullYear()}`;
  }
  // e.g., "OCT 31, 2025 - NOV 02, 2025"
  return `${mon(ci)} ${dd(ci)}, ${ci.getFullYear()} - ${mon(co)} ${dd(co)}, ${co.getFullYear()}`;
}

function cleanup(reservationDetails) {
  /**
   * 1. delete pdf auth letter
   * 2. delete ids
   */
  clearPdfFromDrive()
  clearFirstResponse()
}

function clearPdfFromDrive() {
  // var fileIter = DriveApp.getFolderById(FILEDROP_FOLDER_NAME).getFiles()
  const fileIter = DriveApp.getFolderById(FILEDROP_FOLDER_ID).getFiles()
  while(fileIter.hasNext()) {
    let file = fileIter.next()
    console.log("removing " + file.getName())
    file.setTrashed(true)
  }
}

function clearFirstResponse() {
  const guestRegSS = SpreadsheetApp.openById(REGISTRATION_SPREADSHEET_ID);
  const formRespSheet = guestRegSS.getSheetByName(GUEST_LIST_SHEET_NAME);
  
  sortSheet();

  if (formRespSheet.getLastRow() > 1) {
    formRespSheet.deleteRow(2); // Delete the second row (first actual response)
  }
}

function calculateDaysBetween(compDate, refDate) {
  let convertedCIDate = new Date(compDate);
  let convertedRefDate = new Date(refDate);
  
  let timeDiff = convertedCIDate.getTime() - convertedRefDate.getTime();

  return Math.ceil(timeDiff / (MS_PER_DAY));
}

function getAttachments(reservationDetails) {

  // AUTH LETTER
  if(DriveApp.getFolderById(FILEDROP_FOLDER_ID).getFilesByName(AUTH_LETTER_FILE_NAME).hasNext()) {
    letterFile = DriveApp.getFolderById(FILEDROP_FOLDER_ID).getFilesByName(AUTH_LETTER_FILE_NAME).next();
  } else {
    console.log("No Authorization Letter found in File Drop folder")
    //TODO: stop flow
  }

  let attachments = [letterFile.getAs(MimeType.PDF)]

  // IDs
  reservationDetails.guestList.forEach( guest => {
    if(guest.idLink != null && guest.idLink !== '') {
      // let match = guest.idLink.match((/(?<=id=)[^&]+/))
      const match = guest.idLink.match(/\/file\/d\/([-\w]+)/);
      let idFile = DriveApp.getFileById(match[1])
      // const fileId = match ? match[1] : null;
      attachments.push(idFile.getAs(idFile.getMimeType()));
    }
  })
  return attachments
}