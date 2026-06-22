const APP_NAME = 'Guest Registration';

// Google Drive
const UPLOAD_ROOT_FOLDER_NAME = 'Guest Registration Uploads';

const _props = PropertiesService.getScriptProperties();
const FILEDROP_FOLDER_ID = _props.getProperty('FILEDROP_FOLDER_ID');
const AUTH_LETTER_TEMPLATE_FILE = _props.getProperty('AUTH_LETTER_TEMPLATE_FILE');
const TEMP_FOLDER = _props.getProperty('TEMP_FOLDER');
const REGISTRATION_SPREADSHEET_ID = _props.getProperty('REGISTRATION_SPREADSHEET_ID');
const REGISTRATION_SHEET_GID = Number(_props.getProperty('REGISTRATION_SHEET_GID'));

// Google Docs
const AUTH_LETTER_FILE_NAME = "Authorization Letter.pdf";
const GUEST_LIST_SHEET_NAME = 'Registration';
const GUEST_DETAIL_RANGE= ["e2:h2", "i2:l2", "m2:p2", "q2:t2"];
const SHEET_RANGES = {
    reservation: 'B2:D2',
    checkinCheckout: 'C2:D2',
    parking: 'U2',
    plateNumber: 'V2',
  }

const MS_PER_DAY = 24 * 60 * 60 * 1000;

const PARKING = {
  OFFSITE: 'Free Offsite Parking',
  BASEMENT: 'Basement Pay Parking (₱250 - 24hrs/multiple entries)',
  CONDO: 'Condo Pay Parking (₱400 - 24hrs/multiple entries)',
  NONE: "I won't be availing any parking",
};

const SHEET_HEADERS = [
  'Timestamp',
  'Email Address',
  'Check-in date',
  'Checkout Date',
  'Guest 1 Full Name',
  'Guest 1 Age',
  'Guest 1 Valid ID Type',
  'Guest 1 Valid ID',
  'Guest 2 Full Name',
  'Guest 2 Age',
  'Guest 2 Valid ID Type',
  'Guest 2 Valid ID',
  'Guest 3 Full Name',
  'Guest 3 Age',
  'Guest 3 Valid ID Type',
  'Guest 3 Valid ID',
  'Guest 4 Full Name',
  'Guest 4 Age',
  'Guest 4 Valid ID Type',
  'Guest 4 Valid ID',
  'Choose your parking below',
  'Plate Number'
];