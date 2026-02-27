/**
 * Reads applicant rows from the applications Excel file and maps each row
 * back into a VisaApplication object that the automation can consume.
 */

import * as XLSX from 'xlsx';
import * as path from 'path';
import * as fs from 'fs';
import { VisaApplication, ApplicationDocuments } from '../types/application-data';

/**
 * Reads the Excel workbook at `filePath` and returns one VisaApplication per row.
 * Column headers must match those produced by scripts/json-to-excel.ts.
 */
export function readApplicationsFromExcel(filePath: string): VisaApplication[] {
  // { cellDates: false } prevents xlsx from parsing date-looking cells into
  // JS Date objects, keeping them as the raw string/number the user typed.
  const wb = XLSX.readFile(filePath, { cellDates: false });
  const ws = wb.Sheets[wb.SheetNames[0]];
  // raw: false → use formatted text from cell (respects the display format),
  // so dates stay as "13/10/1973" instead of becoming serial numbers.
  const rawRows = XLSX.utils.sheet_to_json<Record<string, unknown>>(ws, { defval: '', raw: false });

  // xlsx may return numbers for date-like cells (Excel serial dates).
  // Coerce every cell value to a string so .replace() etc. always work.
  const rows = rawRows.map(raw => {
    const r: Record<string, string> = {};
    for (const [k, v] of Object.entries(raw)) {
      r[k] = String(v ?? '');
    }
    return r;
  });

  return rows.map(r => ({
    hostSubmitter: {
      establishmentNameEN: r['Establishment'] ?? '',
      establishmentNo:     '',
      emirate:             r['UAE Emirate'] ?? 'Dubai',
      activity:            '',
      addressEN:           '',
      poBox:               '',
      email:               r['Email'] ?? '',
      mobileNumber:        r['Mobile Number'] ?? '',
    },
    visit: {
      purposeOfVisit:      (r['Purpose of Visit'] || 'Tourism') as 'Tourism',
      dateOfArrival:       r['Date of Arrival'] ?? '',
      dateOfDeparture:     r['Date of Departure'] ?? '',
      portOfEntry:         r['Port of Entry'] ?? '',
      accommodationType:   r['Accommodation Type'] ?? '',
      hotelOrPlaceOfStay:  r['Hotel / Place of Stay'] ?? '',
    },
    passport: {
      passportType:            (r['Passport Type'] || 'Normal') as any,
      passportNumber:          r['Passport Number'] ?? '',
      currentNationality:      r['Current Nationality'] ?? '',
      previousNationality:     r['Previous Nationality'] ?? '',
      fullNameEN:              r['Full Name'] ?? '',
      firstName:               r['First Name'] ?? '',
      middleName:              r['Middle Name'] || undefined,
      lastName:                r['Last Name'] ?? '',
      dateOfBirth:             r['Date of Birth'] ?? '',
      birthCountry:            r['Birth Country'] ?? '',
      birthPlaceEN:            r['Birth Place'] ?? '',
      gender:                  (r['Gender'] || 'Male') as 'Male' | 'Female',
      passportIssueCountry:    r['Passport Issue Country'] ?? '',
      passportIssueDate:       r['Passport Issue Date'] ?? '',
      passportExpiryDate:      r['Passport Expiry Date'] ?? '',
      passportPlaceOfIssueEN:  r['Passport Place of Issue'] ?? '',
    },
    applicant: {
      isInsideUAE:         r['Inside UAE'] === 'Yes',
      motherNameEN:        r['Mother Name'] ?? '',
      maritalStatus:       (r['Marital Status'] || 'UNSPECIFIC') as any,
      relationshipToHost:  r['Relationship to Host'] || 'Not Related',
      religion:            (r['Religion'] || 'UNKNOWN') as any,
      faith:               r['Faith'] ?? '',
      education:           r['Education'] ?? '',
      profession:          r['Profession'] ?? '',
      firstLanguage:       r['First Language'] ?? '',
      comingFromCountry:   r['Coming From Country'] ?? '',
    },
    contact: {
      email:                 r['Email'] ?? '',
      mobileNumber:          r['Mobile Number'] ?? '',
      approvalEmailCopy:     '',
      preferredSMSLanguage:  (r['SMS Language'] || 'ENGLISH') as 'ENGLISH' | 'ARABIC',
      uaeEmirate:            r['UAE Emirate'] || 'Dubai',
      uaeCity:               r['UAE City'] || 'Dubai',
      uaeArea:               r['UAE Area'] || undefined,
      uaeStreet:             r['UAE Street'] || undefined,
      uaeBuilding:           r['UAE Building'] || undefined,
      uaeFloor:              r['UAE Floor'] || undefined,
      uaeFlat:               r['UAE Flat'] || undefined,
      outsideCountry:        r['Outside Country'] || undefined,
      outsideMobile:         r['Outside Mobile'] || undefined,
      outsideCity:           r['Outside City'] || undefined,
      outsideAddress:        r['Outside Address'] || undefined,
    },
    documents: resolveDocuments(r['Documents Folder'] ?? ''),
  }));
}

/**
 * Resolves document file paths from a folder.
 * Looks for files matching the portal's document type names.
 */
function resolveDocuments(folder: string): ApplicationDocuments {
  const empty: ApplicationDocuments = {
    documentsFolder:            '',
    sponsoredPassportPage1:   '',
    passportExternalCoverPage: '',
    personalPhoto:             '',
    hotelReservationPage1:     '',
    returnAirTicketPage1:      '',
  };

  if (!folder) return empty;

  const absFolder = path.resolve(folder);
  if (!fs.existsSync(absFolder)) {
    return { ...empty, documentsFolder: absFolder };
  }

  const files = fs.readdirSync(absFolder).filter(f => !f.startsWith('.') && f !== 'desktop.ini');

  // Find file matching any of the given keywords (case-insensitive)
  const findAny = (...patterns: string[]): string => {
    for (const pattern of patterns) {
      const match = files.find(f => f.toLowerCase().includes(pattern.toLowerCase()));
      if (match) return path.join(absFolder, match);
    }
    return '';
  };

  return {
    documentsFolder:            absFolder,
    sponsoredPassportPage1:    findAny('Sponsored Passport page 1', 'Sponsored Passport', 'passport page 1', 'passport front'),
    passportExternalCoverPage: findAny('Passport External Cover', 'cover page', 'passport cover'),
    personalPhoto:             findAny('Personal Photo', 'photo'),
    hotelReservationPage1:     findAny('Hotel reservation', 'hotel', 'tenancy contract', 'tenancy', 'accommodation'),
    returnAirTicketPage1:      findAny('Return air ticket', 'flight ticket', 'air ticket', 'indigo', 'boarding pass', 'itinerary', 'ticket', 'flight'),
    hotelReservationPage2:     findAny('Hotel reservation') && files.find(f => f.toLowerCase().includes('hotel') && f.toLowerCase().includes('page 2'))
                                  ? path.join(absFolder, files.find(f => f.toLowerCase().includes('hotel') && f.toLowerCase().includes('page 2'))!)
                                  : undefined,
  };
}
