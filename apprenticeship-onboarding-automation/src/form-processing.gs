// ============================================================
// FORM PROCESSING
// ============================================================
//
// Handles incoming Google Forms submissions (learner + employer),
// turns them into structured Learner/Employer records, keeps
// derived fields (age, end dates, tenure) up to date as dates are
// edited, and reacts to a learner's "Paperwork Status" being
// changed to trigger the next document-generation step.

/**
 * Triggered when a learner submits the Learner enrolment form.
 * Reads the raw form response, resolves the learner's programme,
 * matches or creates an Employer record, runs duplicate-learner
 * protection, generates the next Learner ID, and writes a new row
 * to the Learners sheet.
 */
function onFormSubmit(e) {

    const submittedSheetName =
    e.range.getSheet().getName();

  if (
    submittedSheetName !== 'Learner Form Responses - Raw'
  ) {
    return;
  }


  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {

    const ss = SpreadsheetApp.getActiveSpreadsheet();

    const learnerSheet = ss.getSheetByName('Learners');
    const employerSheet = ss.getSheetByName('Employers');
    const programmeSheet = ss.getSheetByName('Programmes');
    const settingsSheet = ss.getSheetByName('Settings');

    if (!learnerSheet || !employerSheet || !programmeSheet || !settingsSheet) {
      throw new Error('One or more required sheets could not be found.');
    }


    // ------------------------------------------------
    // 1. Read exact submitted form response
    // ------------------------------------------------

    const response = {};

    Object.keys(e.namedValues).forEach(key => {
      const value = e.namedValues[key];
      response[key.trim()] = Array.isArray(value) ? value[0] : value;
    });


    // ------------------------------------------------
    // 2. Read Settings
    // ------------------------------------------------

    const settings = getSettingsMap_(ss);

    const learnerPrefix =
      settings['LEARNER_ID_PREFIX'] || 'LRN';

    const employerPrefix =
      settings['EMPLOYER_ID_PREFIX'] || 'EMP';


    // ------------------------------------------------
    // 3. Course answer → Programme Code
    //    (see resolveProgrammeCode_() in template-selection.gs)
    // ------------------------------------------------

    const courseAnswer =
      response['Which programme would you like to enrol on?'] || '';

    const programmeCode = resolveProgrammeCode_(courseAnswer);

    if (!programmeCode) {
      throw new Error(
        `No programme mapping exists for "${courseAnswer}".`
      );
    }


    // ------------------------------------------------
    // 4. Find programme record
    // ------------------------------------------------

    const programmeHeaders = programmeSheet
      .getRange(1, 1, 1, programmeSheet.getLastColumn())
      .getDisplayValues()[0];

    const programmeRows =
      programmeSheet.getLastRow() >= 2
        ? programmeSheet
            .getRange(
              2,
              1,
              programmeSheet.getLastRow() - 1,
              programmeSheet.getLastColumn()
            )
            .getDisplayValues()
        : [];

    let programme = null;

    for (const row of programmeRows) {

      if (String(row[0]).trim() === programmeCode) {

        programme = {};

        programmeHeaders.forEach((header, index) => {
          if (header) {
            programme[header.trim()] = row[index];
          }
        });

        break;
      }
    }

    if (!programme) {
      throw new Error(
        `Programme "${programmeCode}" was not found in Programmes.`
      );
    }


    // ------------------------------------------------
    // 5. EMPLOYER: read headers and existing records
    // ------------------------------------------------

    const employerHeaders = employerSheet
      .getRange(1, 1, 1, employerSheet.getLastColumn())
      .getDisplayValues()[0];

    const employerHeaderIndex = buildHeaderIndex_(employerHeaders);

    const employerRows =
      employerSheet.getLastRow() >= 2
        ? employerSheet
            .getRange(
              2,
              1,
              employerSheet.getLastRow() - 1,
              employerSheet.getLastColumn()
            )
            .getDisplayValues()
        : [];


    // ------------------------------------------------
    // 6. Match or create employer
    // ------------------------------------------------

    const submittedCompany =
      String(response['Name of Employer'] || '').trim();

    if (!submittedCompany) {
      throw new Error('No employer name was supplied on the form.');
    }

// ============================================================
// 3. EMPLOYER + PROGRAMME LOOKUPS
// ============================================================

    let employerId = '';
    let employerWasCreated = false;

    for (const row of employerRows) {

      const existingCompany =
        String(row[employerHeaderIndex['Company Name']] || '').trim();

      if (
        existingCompany.toLowerCase() ===
        submittedCompany.toLowerCase()
      ) {
        employerId =
          String(row[employerHeaderIndex['Employer ID']] || '').trim();

        break;
      }
    }


    if (!employerId) {

      let highestEmployerNumber = 0;

      employerRows.forEach(row => {

        const id =
          String(row[employerHeaderIndex['Employer ID']] || '');

        const match = id.match(
          new RegExp(`^${employerPrefix}-(\\d+)$`)
        );

        if (match) {
          highestEmployerNumber =
            Math.max(
              highestEmployerNumber,
              parseInt(match[1], 10)
            );
        }
      });

      employerId =
        `${employerPrefix}-${String(highestEmployerNumber + 1).padStart(5, '0')}`;

      const now = new Date();

const employerRecord = {
  'Employer ID': employerId,
  'Company Name': submittedCompany,

  'Main Contact First Name':
    response['Line Manager First Name'] || '',

  'Main Contact Surname':
    response['Line Manager Surname'] || '',

  'Main Contact Email':
    response["Employer's Email Address"] || '',

  'Main Contact Telephone':
    response["Employer's Contact Number"] || '',

  'Employer Status': 'Active',
  'Last Updated': now
};

      const employerNewRow =
        new Array(employerHeaders.length).fill('');

      Object.keys(employerRecord).forEach(field => {

        if (employerHeaderIndex[field] !== undefined) {
          employerNewRow[employerHeaderIndex[field]] =
            employerRecord[field];
        }
      });

      const newEmployerRowNumber =
  employerSheet.getLastRow() + 1;

if (
  employerHeaderIndex['Main Contact Telephone'] !== undefined
) {

  employerSheet
    .getRange(
      newEmployerRowNumber,
      employerHeaderIndex['Main Contact Telephone'] + 1
    )
    .setNumberFormat('@');
}

employerSheet
  .getRange(
    newEmployerRowNumber,
    1,
    1,
    employerNewRow.length
  )
  .setValues([employerNewRow]);

      employerWasCreated = true;
    }


    // ------------------------------------------------
    // 7. LEARNER: headers and existing records
    // ------------------------------------------------

    const learnerHeaders = learnerSheet
      .getRange(1, 1, 1, learnerSheet.getLastColumn())
      .getDisplayValues()[0];

    const learnerHeaderIndex = buildHeaderIndex_(learnerHeaders);

    const existingLearners =
      learnerSheet.getLastRow() >= 2
        ? learnerSheet
            .getRange(
              2,
              1,
              learnerSheet.getLastRow() - 1,
              learnerSheet.getLastColumn()
            )
            .getDisplayValues()
        : [];


    // ------------------------------------------------
    // 8. Duplicate learner protection
    // ------------------------------------------------

    const firstName =
      String(response['First Name'] || '').trim();

    const surname =
      String(response['Surname'] || '').trim();

    const email =
      String(response['Email Address'] || '').trim();

    const dob =
      String(response['Date of Birth'] || '').trim();

    for (const row of existingLearners) {

      const sameFirst =
        String(
          row[learnerHeaderIndex['First Name']] || ''
        ).toLowerCase() === firstName.toLowerCase();

      const sameSurname =
        String(
          row[learnerHeaderIndex['Surname']] || ''
        ).toLowerCase() === surname.toLowerCase();

      const sameEmail =
        String(
          row[learnerHeaderIndex['Email Address']] || ''
        ).toLowerCase() === email.toLowerCase();

      const sameDob =
        String(
          row[learnerHeaderIndex['Date of Birth']] || ''
        ) === dob;

      if (sameFirst && sameSurname && sameEmail && sameDob) {
        throw new Error(
          'Possible duplicate learner detected. No learner was created.'
        );
      }
    }


    // ------------------------------------------------
    // 9. Generate next learner ID
    // ------------------------------------------------

    let highestLearnerNumber = 0;

    existingLearners.forEach(row => {

      const id =
        String(
          row[learnerHeaderIndex['Learner ID']] || ''
        );

      const match = id.match(
        new RegExp(`^${learnerPrefix}-(\\d+)$`)
      );

      if (match) {
        highestLearnerNumber =
          Math.max(
            highestLearnerNumber,
            parseInt(match[1], 10)
          );
      }
    });

    const learnerId =
      `${learnerPrefix}-${String(highestLearnerNumber + 1).padStart(5, '0')}`;


    // ------------------------------------------------
    // 10. Build learner record
    // ------------------------------------------------

    const now = new Date();

    const drivingLicenceFront =
  response['Driving Licence - Front'] || '';

    const drivingLicenceBack =
  response['Driving Licence - Back'] || '';

    const drivingLicencePhotos =
  [drivingLicenceFront, drivingLicenceBack]
    .filter(Boolean)
    .join('\n');

    const learnerRecord = {

      'Learner ID': learnerId,

      'First Name': firstName,
      'Surname': surname,
      'Full Name': `${firstName} ${surname}`.trim(),

      'Date of Birth':
        response['Date of Birth'] || '',

      'NI Number':
        response['National Insurance Number'] || '',

      'Email Address':
        response['Email Address'] || '',

      'Telephone Number':
        response['Mobile Number'] || '',

      'Address Line 1':
        response['Address Line 1'] || '',

      'Address Line 2':
        response['Address Line 2'] || '',

      'Town / City':
        response['Town / City'] || '',

      'County':
        response['County'] || '',

      'Postcode':
        response['Post Code'] || '',

      'Normal Working Hours':
        response['Contracted Hours'] || '',

      'Work Address':
        response['Registered Work Address'],

      'Driving Licence Number':
        response['Driving Licence Number'],
      
      

      'Driving Licence Photos':
        drivingLicencePhotos,

      'Employer ID':
        employerId,

      'Company Name':
        submittedCompany,

      'Position in Company':
        response['Position in Company'],

'Line Manager First Name':
  response['Line Manager First Name'] || '',

'Line Manager Surname':
  response['Line Manager Surname'] || '',

'Line Manager Email':
  response["Employer's Email Address"] || '',

'Line Manager Telephone':
  response["Employer's Contact Number"] || '',

      'Employment Start Date':
        response['Start Date of Employment'] || '',

      'Job Title':
        response['Position in Company'] || '',

      'Programme Code':
        programmeCode,

      'Apprenticeship Standard':
        programme['Apprenticeship Standard'] || '',

      'Apprenticeship Level':
        programme['Apprenticeship Level'] || '',

      'Application Status':
        'Submitted',

      'Eligibility Status':
        'Not Checked',

      'Paperwork Status':
        'Not Started',

      'Apprenticeship Agreement Status':
        'Not Generated',

      'Enrolment Pack Status':
        'Not Generated',

      'Record Created':
        now,

      'Last Updated':
        now
    };


    // ------------------------------------------------
    // 11. Write learner row by header name
    // ------------------------------------------------

    const learnerNewRow =
      new Array(learnerHeaders.length).fill('');

    Object.keys(learnerRecord).forEach(field => {

      if (learnerHeaderIndex[field] !== undefined) {
        learnerNewRow[learnerHeaderIndex[field]] =
          learnerRecord[field];
      }
    });

    const newLearnerRowNumber =
  learnerSheet.getLastRow() + 1;

// Force fields that must preserve leading zeroes to Plain text
const textFields = [
  'Telephone Number',
  'Line Manager Telephone',
  'Driving Licence Number',
  'NI Number'
];

textFields.forEach(field => {

  if (learnerHeaderIndex[field] !== undefined) {

    learnerSheet
      .getRange(
        newLearnerRowNumber,
        learnerHeaderIndex[field] + 1
      )
      .setNumberFormat('@');
  }

});

// Write the entire learner row
learnerSheet
  .getRange(
    newLearnerRowNumber,
    1,
    1,
    learnerNewRow.length
  )
  .setValues([learnerNewRow]);


// ============================================================
// 4. LEARNER DRIVE FOLDER MANAGEMENT
// ============================================================

        // ------------------------------------------------
    // 12. Create learner Drive folder structure
    // ------------------------------------------------

    const rootFolderId =
      settings['ROOT_FOLDER_ID'];

    if (!rootFolderId) {
      throw new Error(
        'ROOT_FOLDER_ID is not configured in Settings.'
      );
    }

    const rootFolder =
      DriveApp.getFolderById(rootFolderId);

    const learnerFolderName =
      `${learnerId} - ${firstName} ${surname}`.trim();

    const learnerFolder =
      rootFolder.createFolder(learnerFolderName);

    const subfolders = [
      '01 - Application',
      '02 - Eligibility',
      '03 - Apprenticeship Agreement',
      '04 - Enrolment Pack',
      '05 - Supporting Documents'
    ];

    const createdSubfolders = {};

subfolders.forEach(name => {
  createdSubfolders[name] =
    learnerFolder.createFolder(name);
});

// ------------------------------------------------
// Copy driving licence uploads into
// 05 - Supporting Documents
// ------------------------------------------------

const supportingDocumentsFolder =
  createdSubfolders['05 - Supporting Documents'];

const copiedLicenceLinks = [];

const copyUploadedFile = (fileUrl, label) => {

  if (!fileUrl) {
    return;
  }

  // Extract Google Drive file ID from Forms upload URL
  const match =
    String(fileUrl).match(/[-\w]{25,}/);

  if (!match) {
    console.log(
      `Could not extract file ID from ${label}: ${fileUrl}`
    );
    return;
  }

  const sourceFile =
    DriveApp.getFileById(match[0]);

  const copiedFile =
    sourceFile.makeCopy(
      `${learnerId} - ${firstName} ${surname} - ${label}`,
      supportingDocumentsFolder
    );

  copiedLicenceLinks.push(
    copiedFile.getUrl()
  );
};


copyUploadedFile(
  drivingLicenceFront,
  'Driving Licence - Front'
);

copyUploadedFile(
  drivingLicenceBack,
  'Driving Licence - Back'
);

    const learnerFolderId =
      learnerFolder.getId();

    const learnerFolderUrl =
      learnerFolder.getUrl();

    // Because appendRow() added the learner at the bottom,
    // this is now the learner's row number.
    const learnerRowNumber =
  newLearnerRowNumber;

if (
  learnerHeaderIndex['Driving Licence Photos'] !== undefined &&
  copiedLicenceLinks.length
) {

  learnerSheet
    .getRange(
      learnerRowNumber,
      learnerHeaderIndex['Driving Licence Photos'] + 1
    )
    .setValue(
      copiedLicenceLinks.join('\n')
    );
}


    learnerSheet
      .getRange(
        learnerRowNumber,
        learnerHeaderIndex['Drive Folder ID'] + 1
      )
      .setValue(learnerFolderId);

    learnerSheet
      .getRange(
        learnerRowNumber,
        learnerHeaderIndex['Drive Folder Link'] + 1
      )
      .setValue(learnerFolderUrl);

    learnerSheet
      .getRange(
        learnerRowNumber,
        learnerHeaderIndex['Last Updated'] + 1
      )
      .setValue(new Date());

    // ------------------------------------------------
    // 13. Create Enrolment Details record
    // ------------------------------------------------

    const enrolmentSheet =
      ss.getSheetByName('Enrolment Details');

    if (!enrolmentSheet) {
      throw new Error(
        'Could not find "Enrolment Details".'
      );
    }

    const enrolmentHeaders = enrolmentSheet
      .getRange(
        1,
        1,
        1,
        enrolmentSheet.getLastColumn()
      )
      .getDisplayValues()[0];

    const enrolmentHeaderIndex = buildHeaderIndex_(enrolmentHeaders);

    const enrolmentRecord = {

      'Learner ID':
        learnerId,

      'Driving Licence Number':
        response['Driving Licence Number'] || '',

      'UK Resident 3 Years':
        response[
          'Have you lived in the UK for the past three years?'
        ] || '',

      'E-Signature':
        response['E-Signature'] || '',

      'Work Address':
        response['Registered Work Address'] || '',

      'Employer Contact Email':
        response["Employer's Email Address"] || '',

      'Employer Contact Telephone':
        response["Employer's Contact Number"] || '',

      'Form Submitted':
        response['Timestamp'] || '',

      'Record Created':
        now,

      'Last Updated':
        now
    };

    const enrolmentNewRow =
      new Array(enrolmentHeaders.length).fill('');

    Object.keys(enrolmentRecord).forEach(field => {

      if (enrolmentHeaderIndex[field] !== undefined) {
        enrolmentNewRow[
          enrolmentHeaderIndex[field]
        ] = enrolmentRecord[field];
      }

    });

    enrolmentSheet.appendRow(enrolmentNewRow);

    // ------------------------------------------------
    // 14. Log outcome
    // ------------------------------------------------

    console.log(
      `Learner ${learnerId} created. Employer ${employerId} ${
        employerWasCreated ? 'created' : 'matched'
      }.`
    );

  }

  finally {
    lock.releaseLock();
  }
}



/**
 * Triggered when an employer submits the Employer confirmation
 * form. Matches the submission back to the correct learner (by
 * learner name + employer name) and fills in the learner's line
 * manager details on the Learners sheet.
 */
function handleEmployerFormSubmit(e) {

  const submittedSheet =
    e.range.getSheet();

  if (
    submittedSheet.getName() !== 'Employer Form Responses - Raw'
  ) {
    return;
  }

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const learnerSheet =
    ss.getSheetByName('Learners');

  if (!learnerSheet) {
    throw new Error('Learners sheet not found.');
  }

  // ---------------------------------------------
  // Read submitted employer form response
  // ---------------------------------------------

  const response =
    e.namedValues || {};

  const learnerName =
    String(
      response["Learner/Apprentice's name"]?.[0] || ''
    ).trim();

  const companyName =
    String(
      response['Name of Employer']?.[0] || ''
    ).trim();

  const firstName =
    String(
      response['First Name']?.[0] || ''
    ).trim();

  const surname =
    String(
      response['Surname']?.[0] || ''
    ).trim();

  const mobile =
    String(
      response['Mobile Number']?.[0] || ''
    ).trim();

  const email =
    String(
      response['Email Address']?.[0] || ''
    ).trim();


  if (!learnerName || !companyName) {
    throw new Error(
      'Learner name or employer name is missing from Employer Form submission.'
    );
  }


  // ---------------------------------------------
  // Read Learners sheet
  // ---------------------------------------------

  const data =
    learnerSheet
      .getDataRange()
      .getDisplayValues();

  const headers =
    data[0];

  const headerIndex = buildHeaderIndex_(headers);


  const requiredHeaders = [
    'Full Name',
    'Company Name',
    'Line Manager First Name',
    'Line Manager Surname',
    'Line Manager Email',
    'Line Manager Telephone'
  ];

  requiredHeaders.forEach(header => {
    if (headerIndex[header] === undefined) {
      throw new Error(
        `Required Learners column "${header}" was not found.`
      );
    }
  });


  // ---------------------------------------------
  // Normalise values for safe matching
  // (see normaliseForMatch_() in utilities.gs)
  // ---------------------------------------------

  const targetLearner =
    normaliseForMatch_(learnerName);

  const targetCompany =
    normaliseForMatch_(companyName);


  // ---------------------------------------------
  // Find exact learner + employer match
  // ---------------------------------------------

  const matches = [];

  for (let i = 1; i < data.length; i++) {

    const rowLearner =
      normaliseForMatch_(
        data[i][headerIndex['Full Name']]
      );

    const rowCompany =
      normaliseForMatch_(
        data[i][headerIndex['Company Name']]
      );

    if (
      rowLearner === targetLearner &&
      rowCompany === targetCompany
    ) {
      matches.push(i + 1);
    }
  }


  if (matches.length === 0) {
    throw new Error(
      `No learner found matching "${learnerName}" at "${companyName}".`
    );
  }

  if (matches.length > 1) {
    throw new Error(
      `Multiple learner records matched "${learnerName}" at "${companyName}".`
    );
  }


  const learnerRow =
    matches[0];


  // ---------------------------------------------
  // Update existing learner record
  // ---------------------------------------------

  learnerSheet
    .getRange(
      learnerRow,
      headerIndex['Line Manager First Name'] + 1
    )
    .setValue(firstName);

  learnerSheet
    .getRange(
      learnerRow,
      headerIndex['Line Manager Surname'] + 1
    )
    .setValue(surname);

  learnerSheet
    .getRange(
      learnerRow,
      headerIndex['Line Manager Email'] + 1
    )
    .setValue(email);

  learnerSheet
    .getRange(
      learnerRow,
      headerIndex['Line Manager Telephone'] + 1
    )
    .setValue(mobile);

  SpreadsheetApp.flush();

  console.log(
    `Employer form matched and updated learner: ${learnerName}`
  );
}



/**
 * Looks up the most recent Employer Form Responses row for a
 * given learner/company pairing. Used by the document-generation
 * workflows to pull in the employer contact's details and
 * e-signature when preparing the Apprenticeship Agreement.
 */
function getLatestEmployerAgreement_(learnerName, companyName) {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const sheet =
    ss.getSheetByName('Employer Form Responses - Raw');

  if (!sheet) {
    throw new Error(
      'Employer Form Responses - Raw sheet not found.'
    );
  }

  const data =
    sheet.getDataRange().getDisplayValues();

  if (data.length < 2) {
    return null;
  }

  const headers =
    data[0];

  const headerIndex = {};

  headers.forEach((header, index) => {
    if (header) {
      headerIndex[String(header).trim()] = index;
    }
  });

  const requiredHeaders = [
    "Learner/Apprentice's name",
    'Name of Employer',
    'First Name',
    'Surname',
    'Mobile Number',
    'Email Address',
    'Position in Company',
    'E-Signature'
  ];

  requiredHeaders.forEach(header => {
    if (headerIndex[header] === undefined) {
      throw new Error(
        `Employer Form column "${header}" was not found.`
      );
    }
  });

  const targetLearner =
    normaliseForMatch_(learnerName);

  const targetCompany =
    normaliseForMatch_(companyName);

  // Search newest submission first
  for (let i = data.length - 1; i >= 1; i--) {

    const rowLearner =
      normaliseForMatch_(
        data[i][
          headerIndex["Learner/Apprentice's name"]
        ]
      );

    const rowCompany =
      normaliseForMatch_(
        data[i][
          headerIndex['Name of Employer']
        ]
      );

    if (
      rowLearner === targetLearner &&
      rowCompany === targetCompany
    ) {

      return {
        firstName:
          data[i][headerIndex['First Name']] || '',

        surname:
          data[i][headerIndex['Surname']] || '',

        mobile:
          data[i][headerIndex['Mobile Number']] || '',

        email:
          data[i][headerIndex['Email Address']] || '',

        position:
          data[i][headerIndex['Position in Company']] || '',

        signature:
          data[i][headerIndex['E-Signature']] || ''
      };
    }
  }

  return null;
}



/**
 * Utility for re-running onFormSubmit() against any raw form
 * responses that don't yet have a matching Learners row — for
 * example if the trigger failed or was temporarily disabled.
 * Safe to re-run: entries that already have a matching learner
 * are skipped.
 */
function backfillMissedLearners() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const rawSheet =
    ss.getSheetByName('Learner Form Responses - Raw');

  const learnerSheet =
    ss.getSheetByName('Learners');

  if (!rawSheet || !learnerSheet) {
    throw new Error(
      'Could not find Learner Form Responses - Raw or Learners.'
    );
  }

  // ---------------------------------------------
  // Read raw form responses
  // ---------------------------------------------

  const rawData =
    rawSheet.getDataRange().getDisplayValues();

  const rawHeaders =
    rawData[0].map(header =>
      String(header || '').trim()
    );

  const rawHeaderIndex = buildHeaderIndex_(rawHeaders);

  // ---------------------------------------------
  // Read existing learners
  // ---------------------------------------------

  const learnerData =
    learnerSheet.getDataRange().getDisplayValues();

  const learnerHeaders =
    learnerData[0].map(header =>
      String(header || '').trim()
    );

  const learnerHeaderIndex = buildHeaderIndex_(learnerHeaders);

  // Note: this local matcher intentionally does NOT collapse
  // internal whitespace like normaliseForMatch_() does elsewhere —
  // duplicate protection here matches the exact-field comparison
  // used in onFormSubmit's own duplicate check.
  const normalise = value =>
    String(value || '')
      .trim()
      .toLowerCase();

  // Match the same fields your normal duplicate
  // protection uses.
  const makeKey = (
    firstName,
    surname,
    email,
    dob
  ) =>
    [
      normalise(firstName),
      normalise(surname),
      normalise(email),
      normalise(dob)
    ].join('|');

  const existingLearners = new Set();

  for (let i = 1; i < learnerData.length; i++) {

    const row = learnerData[i];

    const key = makeKey(
      row[learnerHeaderIndex['First Name']],
      row[learnerHeaderIndex['Surname']],
      row[learnerHeaderIndex['Email Address']],
      row[learnerHeaderIndex['Date of Birth']]
    );

    existingLearners.add(key);
  }

  // ---------------------------------------------
  // Re-process only missing form responses
  // ---------------------------------------------

  let created = 0;
  let skipped = 0;
  let failed = 0;

  for (let i = 1; i < rawData.length; i++) {

    const row = rawData[i];

    // Ignore completely empty rows
    if (
      row.every(value =>
        String(value || '').trim() === ''
      )
    ) {
      continue;
    }

    const key = makeKey(
      row[rawHeaderIndex['First Name']],
      row[rawHeaderIndex['Surname']],
      row[rawHeaderIndex['Email Address']],
      row[rawHeaderIndex['Date of Birth']]
    );

    if (existingLearners.has(key)) {
      skipped++;
      continue;
    }

    // Recreate the same namedValues structure
    // produced by a real Google Forms submission.
    const namedValues = {};

    rawHeaders.forEach((header, columnIndex) => {

      if (header) {
        namedValues[header] = [
          row[columnIndex] || ''
        ];
      }

    });

    const fakeEvent = {
      range:
        rawSheet.getRange(i + 1, 1),
      namedValues:
        namedValues
    };

    try {

      onFormSubmit(fakeEvent);

      existingLearners.add(key);
      created++;

      console.log(
        `Backfilled raw response row ${i + 1}.`
      );

    } catch (error) {

      failed++;

      console.error(
        `Backfill failed on raw row ${i + 1}: ${error.message}`
      );

    }
  }

  console.log(
    `Backfill finished. Created: ${created}. ` +
    `Already existed: ${skipped}. ` +
    `Failed: ${failed}.`
  );
}



/**
 * Recalculates the fields that depend on a learner's dates
 * (age at enrolment, practical period end date, apprenticeship
 * end date, length of employment at enrolment) whenever one of
 * the underlying date fields is edited. Called from onEdit().
 */
function calculateLearnerDerivedFields_(sheet, row) {

  const headers = sheet
    .getRange(1, 1, 1, sheet.getLastColumn())
    .getDisplayValues()[0];

  const headerIndex = {};

  headers.forEach((header, index) => {
    if (header) {
      headerIndex[header.trim()] = index + 1;
    }
  });

  const dob =
    sheet.getRange(row, headerIndex['Date of Birth']).getValue();

  const employmentStart =
    sheet.getRange(row, headerIndex['Employment Start Date']).getValue();

  const apprenticeshipStart =
    sheet.getRange(row, headerIndex['Apprenticeship Start Date']).getValue();

  const practicalStart =
     sheet.getRange(
     row,
     headerIndex['Practical Period Start Date']
    ).getValue();


  // Age at Enrolment
  if (
    dob instanceof Date &&
    !isNaN(dob) &&
    apprenticeshipStart instanceof Date &&
    !isNaN(apprenticeshipStart)
  ) {

    let age =
      apprenticeshipStart.getFullYear() - dob.getFullYear();

    const monthDifference =
      apprenticeshipStart.getMonth() - dob.getMonth();

    if (
      monthDifference < 0 ||
      (
        monthDifference === 0 &&
        apprenticeshipStart.getDate() < dob.getDate()
      )
    ) {
      age--;
    }

    sheet
      .getRange(row, headerIndex['Age at Enrolment'])
      .setValue(age);

    // Practical Period End Date = 10 months after Practical Period Start Date
if (
  practicalStart instanceof Date &&
  !isNaN(practicalStart) &&
  headerIndex['Practical Period End Date']
) {

  const practicalEnd =
    new Date(practicalStart);

  practicalEnd.setMonth(
    practicalEnd.getMonth() + 10
  );

  sheet
    .getRange(
      row,
      headerIndex['Practical Period End Date']
    )
    .setValue(practicalEnd);
}


// Apprenticeship End Date = 13 months after Apprenticeship Start Date
if (
  apprenticeshipStart instanceof Date &&
  !isNaN(apprenticeshipStart) &&
  headerIndex['Apprenticeship End Date']
) {

  const apprenticeshipEnd =
    new Date(apprenticeshipStart);

  apprenticeshipEnd.setMonth(
    apprenticeshipEnd.getMonth() + 13
  );

  sheet
    .getRange(
      row,
      headerIndex['Apprenticeship End Date']
    )
    .setValue(apprenticeshipEnd);
}
  }


  // Length of Employment at Enrolment (Months)
  if (
    employmentStart instanceof Date &&
    !isNaN(employmentStart) &&
    apprenticeshipStart instanceof Date &&
    !isNaN(apprenticeshipStart)
  ) {

    let months =
      (apprenticeshipStart.getFullYear() -
        employmentStart.getFullYear()) * 12;

    months +=
      apprenticeshipStart.getMonth() -
      employmentStart.getMonth();

    if (
      apprenticeshipStart.getDate() <
      employmentStart.getDate()
    ) {
      months--;
    }

    sheet
      .getRange(
        row,
        headerIndex['Length of Employment at Enrolment (Months)']
      )
      .setValue(Math.max(months, 0));
  }

}



/**
 * Simple trigger: watches the Learners sheet for edits to any
 * of the date fields that other fields are derived from, and
 * recalculates those derived fields when they change.
 */
function onEdit(e) {

  const sheet = e.range.getSheet();

  if (sheet.getName() !== 'Learners') {
    return;
  }

  const row = e.range.getRow();

  if (row < 2) {
    return;
  }

  const headers = sheet
    .getRange(1, 1, 1, sheet.getLastColumn())
    .getDisplayValues()[0];

  const editedHeader =
    headers[e.range.getColumn() - 1];

  if (
  editedHeader === 'Date of Birth' ||
  editedHeader === 'Employment Start Date' ||
  editedHeader === 'Apprenticeship Start Date' ||
  editedHeader === 'Practical Period Start Date'
) {
  calculateLearnerDerivedFields_(sheet, row);
}

}



/**
 * Simple trigger: watches the Learners sheet for two kinds of
 * status changes and reacts to each:
 *  - "Signature Request Status" changes update the learner's
 *    overall enrolment status, and finalise the enrolment once
 *    signatures are marked Completed.
 *  - "Paperwork Status" being set to "Ready to Generate" kicks
 *    off all three document-generation workflows (Agreement,
 *    Enrolment Pack, Learner Diagnostics) for that learner.
 */
function handlePaperworkStatusEdit(e) {

  const range = e.range;
  const sheet = range.getSheet();

  if (sheet.getName() !== 'Learners') {
    return;
  }

  const row = range.getRow();
  const column = range.getColumn();

  if (row < 2) {
    return;
  }

  const headers = sheet
    .getRange(1, 1, 1, sheet.getLastColumn())
    .getDisplayValues()[0];

  const headerIndex = {};

  headers.forEach((header, index) => {
    if (header) {
      headerIndex[String(header).trim()] = index + 1;
    }
  });

  const editedHeader =
    String(headers[column - 1] || '').trim();

  const newValue =
    String(e.value || '').trim();

  // ------------------------------------------------
  // 1. eSignature workflow status
  // ------------------------------------------------

  if (editedHeader === 'Signature Request Status') {

    const overallStatusColumn =
      headerIndex['Overall Enrolment Status'];

    if (!overallStatusColumn) {
      throw new Error(
        'Overall Enrolment Status column could not be found.'
      );
    }

    if (newValue === 'Ready for eSignature') {
      sheet
        .getRange(row, overallStatusColumn)
        .setValue('Documents Ready');

      if (headerIndex['Last Updated']) {
        sheet
          .getRange(row, headerIndex['Last Updated'])
          .setValue(new Date());
      }

      return;
    }

    if (newValue === 'Sent for eSignature') {
      sheet
        .getRange(row, overallStatusColumn)
        .setValue('Awaiting eSignatures');

      if (headerIndex['Last Updated']) {
        sheet
          .getRange(row, headerIndex['Last Updated'])
          .setValue(new Date());
      }

      return;
    }

    if (newValue === 'Completed') {

      const learnerIdColumn =
        headerIndex['Learner ID'];

      if (!learnerIdColumn) {
        throw new Error('Learner ID column could not be found.');
      }

      const learnerId =
        String(
          sheet
            .getRange(row, learnerIdColumn)
            .getDisplayValue()
        ).trim();

      if (!learnerId) {
        throw new Error(`Learner ID is missing on row ${row}.`);
      }

      finaliseEnrolment_(learnerId);
      return;
    }

    return;
  }

  // ------------------------------------------------
  // 2. Document-generation controls
  // ------------------------------------------------

  const isDocumentGeneration =
  editedHeader === 'Paperwork Status';

if (!isDocumentGeneration) {
  return;
}

  if (newValue !== 'Ready to Generate') {
    return;
  }

  const learnerIdColumn =
    headerIndex['Learner ID'];

  if (!learnerIdColumn) {
    throw new Error('Learner ID column could not be found.');
  }

  const learnerId =
    String(
      sheet
        .getRange(row, learnerIdColumn)
        .getDisplayValue()
    ).trim();

  if (!learnerId) {
    if (headerIndex['Automation Notes']) {
      sheet
        .getRange(row, headerIndex['Automation Notes'])
        .setValue('Generation failed: No Learner ID found.');
    }
    return;
  }

    // ------------------------------------------------
  // Generate complete enrolment document set
  // ------------------------------------------------

  if (headerIndex['Automation Notes']) {
    sheet
      .getRange(
        row,
        headerIndex['Automation Notes']
      )
      .setValue(
        `Generating Agreement, Enrolment Pack and Learner Diagnostics for ${learnerId}...`
      );
  }

  SpreadsheetApp.flush();

  try {

    // 1. Apprenticeship Agreement
    generateAgreementWorkflow_(
      learnerId
    );

    // 2. Enrolment Pack
    generateEnrolmentPackWorkflow_(
      learnerId
    );

    // 3. Learner Diagnostics
    generateLearnerDiagnosticsWorkflow_(
      learnerId
    );

    SpreadsheetApp.flush();


    // ------------------------------------------------
    // Confirm Agreement + Enrolment Pack generated
    // ------------------------------------------------

    const agreementStatus =
      headerIndex[
        'Apprenticeship Agreement Status'
      ]
        ? sheet
            .getRange(
              row,
              headerIndex[
                'Apprenticeship Agreement Status'
              ]
            )
            .getDisplayValue()
            .trim()
        : '';


    const enrolmentPackStatus =
      headerIndex[
        'Enrolment Pack Status'
      ]
        ? sheet
            .getRange(
              row,
              headerIndex[
                'Enrolment Pack Status'
              ]
            )
            .getDisplayValue()
            .trim()
        : '';


    if (
      agreementStatus === 'Generated' &&
      enrolmentPackStatus === 'Generated'
    ) {

      if (
        headerIndex[
          'Overall Enrolment Status'
        ]
      ) {
        sheet
          .getRange(
            row,
            headerIndex[
              'Overall Enrolment Status'
            ]
          )
          .setValue(
            'Documents Ready'
          );
      }


      if (
        headerIndex[
          'Signature Request Status'
        ]
      ) {
        sheet
          .getRange(
            row,
            headerIndex[
              'Signature Request Status'
            ]
          )
          .setValue(
            'Ready for eSignature'
          );
      }
    }


    if (
      headerIndex['Automation Notes']
    ) {
      sheet
        .getRange(
          row,
          headerIndex[
            'Automation Notes'
          ]
        )
        .setValue(
          `Agreement, Enrolment Pack and Learner Diagnostics generated successfully for ${learnerId}.`
        );
    }


    if (
      headerIndex['Last Updated']
    ) {
      sheet
        .getRange(
          row,
          headerIndex[
            'Last Updated'
          ]
        )
        .setValue(
          new Date()
        );
    }


    SpreadsheetApp.flush();


  } catch (error) {

    const errorMessage =
      error && error.message
        ? error.message
        : String(error);


    if (
      headerIndex['Automation Notes']
    ) {
      sheet
        .getRange(
          row,
          headerIndex[
            'Automation Notes'
          ]
        )
        .setValue(
          `Document generation failed: ${errorMessage}`
        );
    }


    SpreadsheetApp.flush();


    console.error(
      `Document generation failed for ${learnerId}: ${errorMessage}`
    );

  }
    }

