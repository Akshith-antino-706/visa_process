/**
 * Full data model for a GDRFA visa application.
 * Maps directly to the actual form fields observed on the GDRFA Smart Channels portal.
 * Source: Video walkthrough analysis + GDFRA Required Fields.xlsx
 *
 * Form URL: smart.gdrfad.gov.ae/SmartChannels/EntryPermitTourism.aspx
 * Three tabs: Application Information | Attachments | Fees
 * Application Information sub-sections: Host/Submitter, Visit Details, Passport Details, Applicant Details, Contact Details
 */

// ─── Host / Submitter ────────────────────────────────────────────────────────
// The sponsoring travel agency details — usually pre-filled per account
export interface HostSubmitter {
  establishmentNameEN: string;
  establishmentNameAR?: string;     // Auto-translated by portal
  establishmentNo: string;
  emirate: string;
  activity: string;
  addressEN: string;
  addressAR?: string;               // Auto-translated by portal
  poBox?: string;
  email: string;
  mobileNumber: string;
}

// ─── Visit Details ────────────────────────────────────────────────────────────
export interface VisitDetails {
  purposeOfVisit: 'Tourism';
  dateOfArrival: string;            // DD-MM-YYYY
  dateOfDeparture: string;          // DD-MM-YYYY
  portOfEntry: string;
  accommodationType: string;
  hotelOrPlaceOfStay: string;
}

// ─── Passport Details ────────────────────────────────────────────────────────
export interface PassportDetails {
  passportType: PassportType;
  passportNumber: string;           // Filled from OCR
  currentNationality: string;       // e.g. "349 - SOUTH AFRICA"
  previousNationality?: string;

  // Manually entered (OCR extracts these; portal may also auto-fill some)
  fullNameEN: string;
  firstName: string;                // Given name (first word of givenNames)
  middleName?: string;              // Remaining given names (if any)
  lastName: string;                 // Surname from MRZ
  fullNameAR?: string;              // Auto-transliterated by portal
  dateOfBirth: string;              // DD-MM-YYYY
  birthCountry: string;
  birthPlaceEN: string;
  birthPlaceAR?: string;
  gender: 'Male' | 'Female';
  passportIssueCountry: string;
  passportIssueDate: string;        // DD-MM-YYYY
  passportExpiryDate: string;       // DD-MM-YYYY
  passportPlaceOfIssueEN: string;
  passportPlaceOfIssueAR?: string;
}

export type PassportType =
  | 'Normal'
  | 'Diplomatic'
  | 'Official Passport'
  | 'Service Passport'
  | 'Foreign Passport'
  | 'Private'
  | 'Public Affairs'
  | 'Assignment'
  | 'Egyptian Travel Doc'
  | 'Lebanese Travel Doc'
  | 'Syrian Travel Doc'
  | 'Travel Doc';

// ─── Applicant Details ────────────────────────────────────────────────────────
export interface ApplicantDetails {
  isInsideUAE: boolean;
  motherNameEN: string;
  motherNameAR?: string;
  maritalStatus: MaritalStatus;
  relationshipToHost: string;       // Default: 'Not Related'
  religion: Religion;
  faith: string;                    // e.g. 'Unknown', depends on religion
  education: string;                // e.g. 'UNIVERSITY DEGREE'
  profession: string;               // e.g. 'SALES EXECUTIVE'
  firstLanguage: string;            // e.g. 'ENGLISH'
  comingFromCountry: string;        // e.g. '349 - SOUTH AFRICA'
}

export type MaritalStatus =
  | 'Single'
  | 'Married'
  | 'Divorced'
  | 'Widow'
  | 'Deceased'
  | 'UNSPECIFIC'
  | 'Child';

export type Religion =
  | 'MUSLIM'
  | 'CHRISTIAN'
  | 'HINDU'
  | 'JEWISH'
  | 'BUDDHIST'
  | 'SIKH'
  | 'BAHAEI'
  | 'KADIANI'
  | 'NON RELIGIOUS'
  | 'OTHER RELIGION'
  | 'UNKNOWN';

// ─── Contact Details (includes UAE address) ───────────────────────────────────
export interface ContactDetails {
  email: string;
  mobileNumber: string;
  approvalEmailCopy?: string;
  preferredSMSLanguage: 'ENGLISH' | 'ARABIC';
  // Address Inside UAE
  uaeEmirate: string;               // Select: ABU DHABI, DUBAI, SHARJAH, etc.
  uaeCity: string;                  // Select: populated after Emirate AJAX
  uaeArea?: string;                 // Select: populated after City AJAX
  uaeStreet?: string;               // Text input
  uaeBuilding?: string;             // Text input (Building/Villa)
  uaeFloor?: string;                // Text input
  uaeFlat?: string;                 // Text input (Flat/Villa no.)
  // Address Outside UAE
  outsideCountry?: string;          // Select: cmbApplicantOutsideCountry (country code or name)
  outsideMobile?: string;           // Text input: inpAddressOutsideMobileNumber
  outsideCity?: string;             // Text input: inpAddressOutsideCity
  outsideAddress?: string;          // Text input: inpAddressOutsideAddress1
}

// ─── Documents to Upload ──────────────────────────────────────────────────────
// Slots match EXACTLY the labels on the Attachments tab in the portal.
// File constraints: jpg/pdf/png, max 1000 KB each.
// IMPORTANT: File names must NOT contain: \ / : * ? " < > |
export interface ApplicationDocuments {
  // Required (*) slots
  hotelReservationPage1: string;    // "Hotel reservation/Place of stay - Page 1"
  passportExternalCoverPage: string;// "Passport External Cover Page"
  personalPhoto: string;            // "Personal Photo"
  returnAirTicketPage1: string;     // "Return air ticket - Page 1"
  sponsoredPassportPage1: string;   // "Sponsored Passport page 1" — also used for OCR

  // Optional slots
  hotelReservationPage2?: string;   // "Hotel reservation/Place of stay - Page 2"
  othersPage1?: string;             // "Others Page 1"
  returnAirTicketPage2?: string;
  sponsoredPassportPages?: string[]; // Page 2 to 5
}

// ─── Full Application ─────────────────────────────────────────────────────────
export interface VisaApplication {
  hostSubmitter: HostSubmitter;
  visit: VisitDetails;
  passport: PassportDetails;
  applicant: ApplicantDetails;
  contact: ContactDetails;
  documents: ApplicationDocuments;
}
