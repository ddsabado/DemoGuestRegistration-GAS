function doGet() {
  return HtmlService.createTemplateFromFile('RegistrationForm')
    .evaluate()
    .setTitle(APP_NAME)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function uploadGuestIdFiles(files) {
  if (!Array.isArray(files)) {
    throw new Error('Invalid upload payload.');
  }

  const folder = getUploadFolder_();
  return files.map(function(file) {
    validateUpload_(file);

    const safeName = sanitizeFileName_(file.name || 'guest-id');
    const bytes = Utilities.base64Decode(file.dataBase64);
    const blob = Utilities.newBlob(bytes, file.mimeType || 'application/octet-stream', safeName);
    const savedFile = folder.createFile(blob);

    return {
      id: savedFile.getId(),
      name: savedFile.getName(),
      url: savedFile.getUrl(),
      mimeType: savedFile.getMimeType(),
    };
  });
}

function submitGuestRegistration(payload) {
  const registration = normalizeRegistration_(payload);
  appendRegistrationRow_(registration);
  execute();
  const recipient = registration.email;

  GmailApp.sendEmail(recipient, 'Guest Registration Confirmation', '', {
    htmlBody: buildConfirmationHtml_(registration),
    name: APP_NAME,
    from: 'ddacsat.business@gmail.com',
  });

  execute();
  return {
    ok: true,
    message: 'Registration submitted. A confirmation email has been sent to ' + recipient + '.',
  };
}

function appendRegistrationRow_(registration) {
  const sheet = getRegistrationSheet_();
  ensureRegistrationHeader_(sheet);

  const headers = getRegistrationHeadersFromSheet_(sheet);
  const row = buildSubmissionRowFromHeaders_(headers, registration);
  sheet.appendRow(row);
}

function getRegistrationSheet_() {
  const spreadsheet = SpreadsheetApp.openById(REGISTRATION_SPREADSHEET_ID);
  const sheets = spreadsheet.getSheets();

  for (let i = 0; i < sheets.length; i++) {
    if (sheets[i].getSheetId() === REGISTRATION_SHEET_GID) {
      return sheets[i];
    }
  }

  throw new Error('Could not find the registrations sheet tab in the spreadsheet.');
}

function ensureRegistrationHeader_(sheet) {
  if (sheet.getLastRow() > 0) {
    return;
  }

  sheet.appendRow(getRegistrationHeaders_());
}

function getRegistrationHeaders_() {
  return SHEET_HEADERS;
}

function getRegistrationHeadersFromSheet_(sheet) {
  const lastColumn = Math.max(sheet.getLastColumn(), 1);
  if (lastColumn === 1 && !cleanString_(sheet.getRange(1, 1).getValue())) {
    return getRegistrationHeaders_();
  }

  const row = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
  const headers = row.map(function(value) {
    return cleanString_(value);
  }).filter(function(value) {
    return value !== '';
  });

  return headers.length ? headers : getRegistrationHeaders_();
}

function buildSubmissionRowFromHeaders_(headers, registration) {
  const valuesByHeader = buildSubmissionValues_(registration);
  return headers.map(function(header) {
    return Object.prototype.hasOwnProperty.call(valuesByHeader, header)
      ? valuesByHeader[header]
      : '';
  });
}

function buildSubmissionValues_(registration) {
  const values = {};
  values['Timestamp'] = new Date();
  values['Email Address'] = registration.email;
  values['Check-in date'] = registration.checkInDate;
  values['Checkout Date'] = registration.checkOutDate;
  values['Choose your parking below'] = registration.parkingOption;
  values['Plate Number'] = registration.plateNumber || '';

  for (let i = 0; i < 4; i++) {
    const guest = registration.guests[i] || {};
    const n = i + 1;
    const name = cleanString_(guest.name);
    const age = cleanString_(guest.age);
    const idType = cleanString_(guest.idType);
    const fileUrls = joinFileUrls_(guest.idFiles);

    values['Guest ' + n + ' Full Name'] = name;
    values['Guest ' + n + ' Age'] = age;
    values['Guest ' + n + ' Valid ID Type'] = idType;
    values['Guest ' + n + ' Valid ID'] = fileUrls;
  }

  return values;
}

function joinFileUrls_(files) {
  if (!Array.isArray(files) || !files.length) {
    return '';
  }

  return files.map(function(file) {
    return file.url;
  }).join('\n');
}

function joinAllFileUrls_(guests) {
  if (!Array.isArray(guests) || !guests.length) {
    return '';
  }

  const urls = [];
  guests.forEach(function(guest) {
    (guest.idFiles || []).forEach(function(file) {
      if (file && file.url) {
        urls.push(file.url);
      }
    });
  });
  return urls.join('\n');
}

function getUploadFolder_() {
  const props = PropertiesService.getScriptProperties();
  const configuredFolderId = props.getProperty('UPLOAD_FOLDER_ID');

  if (configuredFolderId) {
    try {
      return DriveApp.getFolderById(configuredFolderId);
    } catch (err) {
      props.deleteProperty('UPLOAD_FOLDER_ID');
    }
  }

  const existing = DriveApp.getFoldersByName(UPLOAD_ROOT_FOLDER_NAME);
  const folder = existing.hasNext()
    ? existing.next()
    : DriveApp.createFolder(UPLOAD_ROOT_FOLDER_NAME);

  props.setProperty('UPLOAD_FOLDER_ID', folder.getId());
  return folder;
}

function validateUpload_(file) {
  if (!file || typeof file !== 'object') {
    throw new Error('Invalid file upload.');
  }
  if (!file.dataBase64) {
    throw new Error('Uploaded file is empty.');
  }

  const mimeType = String(file.mimeType || '').toLowerCase();
  if (!(mimeType.indexOf('image/') === 0 || mimeType === 'application/pdf')) {
    throw new Error('Only image and PDF uploads are accepted.');
  }
}

function normalizeRegistration_(payload) {
  if (!payload || typeof payload !== 'object') {
    throw new Error('Invalid registration payload.');
  }

  const registration = {
    email: cleanString_(payload.email).toLowerCase(),
    plateNumber: cleanString_(payload.plateNumber).toUpperCase(),
    checkInDate: cleanString_(payload.checkInDate),
    checkOutDate: cleanString_(payload.checkOutDate),
    parkingOption: cleanString_(payload.parkingOption),
    guests: normalizeGuests_(payload.guests),
  };

  if (!registration.email || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(registration.email)) {
    throw new Error('A valid email address is required.');
  }
  if (!isIsoDate_(registration.checkInDate)) {
    throw new Error('A valid check-in date is required.');
  }
  if (!isIsoDate_(registration.checkOutDate)) {
    throw new Error('A valid checkout date is required.');
  }
  if (dateOnlyMs_(registration.checkOutDate) < dateOnlyMs_(registration.checkInDate)) {
    throw new Error('Checkout date cannot be earlier than check-in date.');
  }
  if (dateOnlyMs_(registration.checkInDate) < dateOnlyMs_(todayIsoDate_())) {
    throw new Error('Check-in date cannot be in the past.');
  }
  if (dateOnlyMs_(registration.checkOutDate) > dateOnlyMs_(registration.checkInDate) + 30 * MS_PER_DAY) {
    throw new Error('Checkout date cannot be more than 30 days from check-in date.');
  }

  const allowedParking = Object.keys(PARKING).map(function(key) {
    return PARKING[key];
  });
  if (allowedParking.indexOf(registration.parkingOption) === -1) {
    throw new Error('Please choose a valid parking option.');
  }

  if (isSameDay_(registration.checkInDate)) {
    const sameDayAllowed = [PARKING.CONDO, PARKING.NONE];
    if (sameDayAllowed.indexOf(registration.parkingOption) === -1) {
      throw new Error('Same-day bookings only allow Condo Pay Parking or No Parking.');
    }
  }

  validateGuestPayload_(payload.guests);
  return registration;
}

function validateGuestPayload_(guests) {
  if (!Array.isArray(guests)) {
    throw new Error('Guest 1 full name, age, valid ID type, and valid ID upload are required.');
  }

  const limitedGuests = guests.slice(0, 4);
  const guest1 = limitedGuests[0] || {};

  if (!cleanString_(guest1.name)) {
    throw new Error('Guest 1 full name is required.');
  }
  if (!cleanString_(guest1.age)) {
    throw new Error('Guest 1 age is required.');
  }
  if (!cleanString_(guest1.idType)) {
    throw new Error('Guest 1 valid ID type is required.');
  }
  if (!Array.isArray(guest1.idFiles) || guest1.idFiles.length !== 1) {
    throw new Error('Guest 1 valid ID upload is required.');
  }

  limitedGuests.forEach(function(guest, index) {
    if (!guest || typeof guest !== 'object') {
      return;
    }

    const guestNumber = index + 1;
    const name = cleanString_(guest.name);
    const ageText = cleanString_(guest.age);
    const idType = cleanString_(guest.idType);
    const idFiles = Array.isArray(guest.idFiles) ? guest.idFiles : [];
    const hasAnyValue = name || ageText || idType || idFiles.length;

    if (!hasAnyValue) {
      return;
    }
    if (name.length > 50) {
      throw new Error('Guest ' + guestNumber + ' full name must be 50 characters or fewer.');
    }
    if (idFiles.length > 1) {
      throw new Error('Guest ' + guestNumber + ' can only upload 1 valid ID file.');
    }
    if (ageText) {
      const age = Number(ageText);
      if (!Number.isInteger(age) || age < 0 || age >= 100) {
        throw new Error('Guest ' + guestNumber + ' age must be a whole number below 100.');
      }
    }
  });
}

function normalizeGuests_(guests) {
  if (!Array.isArray(guests)) {
    return [];
  }

  return guests.slice(0, 4).map(function(guest) {
    const normalized = {
      name: cleanString_(guest && guest.name).toUpperCase(),
      age: cleanString_(guest && guest.age),
      idType: cleanString_(guest && guest.idType),
      idFiles: normalizeUploadedFiles_(guest && guest.idFiles),
    };

    return normalized;
  }).filter(function(guest) {
    return guest.name !== '';
  });
}

function normalizeUploadedFiles_(files) {
  if (!Array.isArray(files)) {
    return [];
  }

  return files.map(function(file) {
    return {
      id: cleanString_(file && file.id),
      name: cleanString_(file && file.name),
      url: cleanString_(file && file.url),
      mimeType: cleanString_(file && file.mimeType),
    };
  }).filter(function(file) {
    return file.id && file.name && file.url;
  });
}

function buildConfirmationHtml_(registration) {
  const rows = [
    ['Email', registration.email],
    ['Check-in', registration.checkInDate],
    ['Checkout', registration.checkOutDate],
    ['Parking', registration.parkingOption],
    ['Plate Number', registration.plateNumber || 'N/A'],
  ].map(function(row) {
    return '<tr><td style="padding:6px 14px 6px 0;color:#555;">' + escapeHtml_(row[0]) +
      '</td><td style="padding:6px 0;">' + escapeHtml_(row[1]) + '</td></tr>';
  }).join('');

  const guests = registration.guests.length
    ? registration.guests.map(function(guest, index) {
      const files = guest.idFiles.length
        ? '<ul>' + guest.idFiles.map(function(file) {
          return '<li>' + escapeHtml_(file.name) + '</li>';
        }).join('') + '</ul>'
        : '<div style="color:#777;">No ID files uploaded</div>';

      return '<h3 style="margin:18px 0 6px;">Guest ' + (index + 1) + '</h3>' +
        '<div><strong>' + escapeHtml_(guest.name) + '</strong></div>' +
        '<div>Age: ' + escapeHtml_(guest.age || 'N/A') + '</div>' +
        '<div>ID Type: ' + escapeHtml_(guest.idType || 'N/A') + '</div>' +
        '<div style="margin-top:6px;">ID Uploads:' + files + '</div>';
    }).join('')
    : '<p style="color:#777;">No guests entered.</p>';

  return '<div style="font-family:Arial,sans-serif;line-height:1.45;">' +
    '<h2 style="margin:0 0 12px;">Guest Registration Confirmation</h2>' +
    '<p>Your registration details were received.</p>' +
    '<table style="border-collapse:collapse;">' + rows + '</table>' +
    '<h2 style="margin:24px 0 8px;">Guest Details</h2>' +
    guests +
    '</div>';
}

function buildConfirmationText_(registration) {
  const lines = [
    'Guest Registration Confirmation',
    '',
    'Email: ' + registration.email,
    'Check-in: ' + registration.checkInDate,
    'Checkout: ' + registration.checkOutDate,
    'Parking: ' + registration.parkingOption,
    'Plate Number: ' + (registration.plateNumber || 'N/A'),
    '',
    'Guest Details:',
  ];

  if (!registration.guests.length) {
    lines.push('No guests entered.');
  } else {
    registration.guests.forEach(function(guest, index) {
      lines.push('');
      lines.push('Guest ' + (index + 1) + ': ' + guest.name);
      lines.push('Age: ' + (guest.age || 'N/A'));
      lines.push('ID Type: ' + (guest.idType || 'N/A'));

      if (guest.idFiles.length) {
        guest.idFiles.forEach(function(file) {
          lines.push('ID File: ' + file.name);
        });
      } else {
        lines.push('ID Files: None uploaded');
      }
    });
  }

  return lines.join('\n');
}

function isSameDay_(isoDateStr) {
  const today = new Date();
  const selected = parseLocalDate_(isoDateStr);

  return selected.getFullYear() === today.getFullYear() &&
    selected.getMonth() === today.getMonth() &&
    selected.getDate() === today.getDate();
}

function todayIsoDate_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function dateOnlyMs_(isoDateStr) {
  const date = parseLocalDate_(isoDateStr);
  return new Date(date.getFullYear(), date.getMonth(), date.getDate()).getTime();
}

function parseLocalDate_(isoDateStr) {
  const parts = String(isoDateStr).split('-').map(Number);
  return new Date(parts[0], parts[1] - 1, parts[2]);
}

function isIsoDate_(value) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(value))) {
    return false;
  }

  const date = parseLocalDate_(value);
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd') === value;
}

function sanitizeFileName_(name) {
  return String(name).replace(/[\\/:*?"<>|#%{}~&]/g, '_').slice(0, 160);
}

function cleanString_(value) {
  return String(value == null ? '' : value).trim();
}

function escapeHtml_(value) {
  return cleanString_(value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}
