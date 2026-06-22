# DemoGuestRegistration

An automated guest registration system for short-term rental properties, built entirely on Google Workspace and Apps Script. No external backend required.

## Live Demo

- **Registration Form:** https://sites.google.com/view/demo-guest-reg/registration
- **Google Sheet:** https://docs.google.com/spreadsheets/d/17513m1poTglv9oMVPT0dora-gPvucdBMIitmVmeGCWc

## Features

- Custom web registration form for pre-check-in data collection
- Collects guest details (name, age, valid ID) for up to 4 guests
- ID image uploads stored directly to Google Drive
- Automatic confirmation email sent to guest upon form submission
- Personalized authorization letter generated from a Google Docs template and emailed as a PDF attachment
- Daily time-based trigger for automated letter dispatch each morning
- Parking selection handling with formatted output in the authorization letter
- Emails sent from a configured alias address

## How It Works

1. Guest fills out the registration form
2. Submission is appended to the Google Sheet
3. A confirmation email is sent immediately to the guest
4. The `execute()` function runs — it validates the check-in date, generates the authorization letter from the Docs template, and emails it with guest ID attachments
5. A daily trigger also runs `execute()` each morning for same-day check-ins

## Apps Script Files

| File | Purpose |
|---|---|
| `RegistrationFormCode.js` | `doGet` / `doPost`, form submission handler, confirmation email |
| `DailyJob.js` | Auth letter generation, email sending, sheet sorting, cleanup |
| `GuestConfig.js` | All configuration constants (IDs, folder references, sheet structure) |
| `RegistrationForm.html` | Guest-facing registration form UI |
| `AuthLetterBody.html` | Email body template for authorization letter email |
| `Styles.html` | Shared CSS for the form |

## Sheet Structure

| Sheet | Purpose |
|---|---|
| Registration | Guest submissions — timestamp, email, check-in/out dates, up to 4 guests with ID info, parking selection |

## Configuration

Update `GuestConfig.js` before deploying:

| Constant | Description |
|---|---|
| `REGISTRATION_SPREADSHEET_ID` | Google Sheet ID |
| `REGISTRATION_SHEET_GID` | Sheet tab GID |
| `AUTH_LETTER_TEMPLATE_FILE` | Google Doc template file ID |
| `TEMP_FOLDER` | Drive folder ID for temp letter files and PDF output |
| `FILEDROP_FOLDER_ID` | Drive folder ID for generated PDFs |

## Deployment

Uses [clasp](https://github.com/google/clasp) for local development.

```bash
clasp push --force
clasp deploy --deploymentId <id> --description "description"
```
