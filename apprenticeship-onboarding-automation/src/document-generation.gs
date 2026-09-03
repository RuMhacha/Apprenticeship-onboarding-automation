// ============================================================
// DOCUMENT GENERATION
// ============================================================
//
// Generates the three onboarding documents for a learner
// (Apprenticeship Agreement, Enrolment Pack, Learner Diagnostics)
// by copying the correct programme template, replacing
// {{PLACEHOLDER}} tokens with the learner's data, exporting a PDF,
// and logging the result to the Document Logs sheet. Also handles
// marking a learner's enrolment as finalised once all documents
// and signatures are in place.

/**
 * Generates a learner's Apprenticeship Agreement: validates the
 * learner record is complete, selects the correct template for
 * their programme, copies it into their Drive folder, fills in
 * the {{PLACEHOLDER}} tokens, exports a PDF, updates the Learners
 * row, and appends a Document Logs entry.
 *
 * Idempotent — if an Agreement already exists for this learner,
 * the function returns without creating a duplicate.
 */
function generateAgreementWorkflow_(learnerId) {

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {

    const ss = SpreadsheetApp.getActiveSpreadsheet();

    const learnerSheet = ss.getSheetByName('Learners');
    const settingsSheet = ss.getSheetByName('Settings');
    const logSheet = ss.getSheetByName('Document Logs');

    if (!learnerSheet || !settingsSheet || !logSheet) {
      throw new Error(
        'Could not find Learners, Settings or Document Logs.'
      );
    }


    // ------------------------------------------------
    // 1. Read learner
    // ------------------------------------------------

    const learnerData =
      learnerSheet.getDataRange().getValues();

    const headers =
      learnerData[0];

    const headerIndex = buildHeaderIndex_(headers);

    let learnerRow = null;
    let learnerRowNumber = null;

    for (let i = 1; i < learnerData.length; i++) {

      if (
        String(
          learnerData[i][headerIndex['Learner ID']]
        ).trim() === learnerId
      ) {

        learnerRow =
          learnerData[i];

        learnerRowNumber =
          i + 1;

        break;
      }
    }

    if (!learnerRow) {
      throw new Error(
        `Learner ${learnerId} was not found.`
      );
    }


    const getValue = (header) => {

      const index =
        headerIndex[header];

      if (index === undefined) {
        return '';
      }

      return learnerRow[index] || '';
    };


    // ------------------------------------------------
    // 2. Readiness checks
    // ------------------------------------------------

    const requiredFields = [
      'Full Name',
      'Company Name',
      'Programme Code',
      'Apprenticeship Standard',
      'Apprenticeship Level',
      'Apprenticeship Start Date',
      'Apprenticeship End Date',
      'Practical Period Start Date',
      'Practical Period End Date',
      'Planned OTJ Hours',
      'Drive Folder ID'
    ];

    const missingFields =
      requiredFields.filter(field => !getValue(field));

    if (missingFields.length) {

      throw new Error(
        `Agreement cannot be generated. Missing: ${missingFields.join(', ')}`
      );
    }


    // ------------------------------------------------
    // 3. Idempotency check
    // ------------------------------------------------

    const existingDocumentId =
      getValue('Agreement Document ID');

    const existingPdfLink =
      getValue('Agreement PDF Link');

    if (existingDocumentId || existingPdfLink) {

      console.log(
        `Agreement already exists for ${learnerId}. Nothing created.`
      );

      return;
    }


    // ------------------------------------------------
    // 4. Read Settings and select the right template
    //    for this learner's programme
    // ------------------------------------------------

    const settings = getSettingsMap_(ss);

    const programmeCode =
      getValue('Programme Code');

    const templateId =
      selectTemplateId_(programmeCode, 'AGREEMENT', settings);



    // ------------------------------------------------
    // 6. Build replacements
    //    (see formatDate_() in utilities.gs)
    // ------------------------------------------------
    // ------------------------------------------------
    const employerAgreement =
      getLatestEmployerAgreement_(
      getValue('Full Name'),
      getValue('Company Name')
  );

    const replacements = {

      '{{LEARNER_ID}}':
        getValue('Learner ID'),

      '{{LEARNER_FULL_NAME}}':
        getValue('Full Name'),

      '{{EMPLOYER_NAME}}':
        getValue('Company Name'),

      '{{PROGRAMME_CODE}}':
        getValue('Programme Code'),

      '{{APPRENTICESHIP_STANDARD}}':
        getValue('Apprenticeship Standard'),

      '{{WORK_ADDRESS}}':
        getValue('Work Address'),

      '{{DRIVING_LICENCE_NUMBER}}':
        getValue('Driving Licence Number'),

      '{{APPRENTICESHIP_LEVEL}}':
        getValue('Apprenticeship Level'),

      '{{APPRENTICESHIP_TITLE}}':
        programmeDisplayName_(
          getValue('Programme Code'),
          getValue('Apprenticeship Standard')
        ),

      '{{APPRENTICESHIP_START_DATE}}':
        formatDate_(
          getValue('Apprenticeship Start Date')
        ),

      '{{APPRENTICESHIP_END_DATE}}':
        formatDate_(
          getValue('Apprenticeship End Date')
        ),

      '{{PRACTICAL_PERIOD_START_DATE}}':
        formatDate_(
          getValue('Practical Period Start Date')
        ),

      '{{PRACTICAL_PERIOD_END_DATE}}':
        formatDate_(
          getValue('Practical Period End Date')
        ),

      '{{PLANNED_OTJ_HOURS}}':
        getValue('Planned OTJ Hours'),
        
      '{{EMPLOYER_CONTACT_FIRST_NAME}}':
        employerAgreement?.firstName || getValue('Line Manager First Name'),

      '{{EMPLOYER_CONTACT_SURNAME}}':
        employerAgreement?.surname || getValue('Line Manager Surname'),

      '{{EMPLOYER_CONTACT_EMAIL}}':
        employerAgreement?.email || getValue('Line Manager Email'),

      '{{EMPLOYER_CONTACT_TELEPHONE}}':
        employerAgreement?.mobile || getValue('Line Manager Telephone'),

      '{{EMPLOYER_CONTACT_POSITION}}':
        employerAgreement?.position || '',

      '{{EMPLOYER_E_SIGNATURE}}':
        employerAgreement?.signature || '',

      };


    // ------------------------------------------------
    // 7. Find agreement folder
    // ------------------------------------------------


    const learnerFolder =
      DriveApp.getFolderById(
        getValue('Drive Folder ID')
      );

    const agreementFolders =
      learnerFolder.getFoldersByName(
        '03 - Apprenticeship Agreement'
      );

    if (!agreementFolders.hasNext()) {
      throw new Error(
        '03 - Apprenticeship Agreement folder not found.'
      );
    }

    const agreementFolder =
      agreementFolders.next();


    // ------------------------------------------------
    // 8. Create Google Doc
    // ------------------------------------------------

    const documentName =
      `${learnerId} - ${getValue('Full Name')} - Apprenticeship Agreement`;

    const templateFile =
      DriveApp.getFileById(templateId);

    const documentFile =
      templateFile.makeCopy(
        documentName,
        agreementFolder
      );

    const doc =
      DocumentApp.openById(
        documentFile.getId()
      );

    const body =
      doc.getBody();

    Object.keys(replacements)
      .forEach(placeholder => {

        body.replaceText(
          escapeRegex_(placeholder),
          String(
            replacements[placeholder] ?? ''
          )
        );

      });

    doc.saveAndClose();


    // ------------------------------------------------
    // 9. Create PDF
    // ------------------------------------------------

    const pdfName =
      `${documentName}.pdf`;

    const pdfBlob =
      documentFile
        .getAs(MimeType.PDF)
        .setName(pdfName);

    const pdfFile =
      agreementFolder.createFile(
        pdfBlob
      );


    // ------------------------------------------------
    // 10. Update learner record
    // ------------------------------------------------

    if (
      headerIndex['Agreement Document ID'] !== undefined
    ) {

      learnerSheet
        .getRange(
          learnerRowNumber,
          headerIndex['Agreement Document ID'] + 1
        )
        .setValue(
          documentFile.getId()
        );

    }


    if (
      headerIndex['Agreement PDF Link'] !== undefined
    ) {

      learnerSheet
        .getRange(
          learnerRowNumber,
          headerIndex['Agreement PDF Link'] + 1
        )
        .setValue(
          pdfFile.getUrl()
        );

    }


    if (
      headerIndex['Apprenticeship Agreement Status'] !== undefined
    ) {

      learnerSheet
        .getRange(
          learnerRowNumber,
          headerIndex['Apprenticeship Agreement Status'] + 1
        )
        .setValue('Generated');

    }


    if (
      headerIndex['Paperwork Status'] !== undefined
    ) {

      learnerSheet
        .getRange(
          learnerRowNumber,
          headerIndex['Paperwork Status'] + 1
        )
        .setValue('Generated');

    }


    if (
      headerIndex['Last Updated'] !== undefined
    ) {

      learnerSheet
        .getRange(
          learnerRowNumber,
          headerIndex['Last Updated'] + 1
        )
        .setValue(new Date());

    }


    // ------------------------------------------------
    // 11. Create Document Log
    // ------------------------------------------------

    const logHeaders =
      logSheet
        .getRange(
          1,
          1,
          1,
          logSheet.getLastColumn()
        )
        .getDisplayValues()[0];


    const existingLogIds =
      logSheet.getLastRow() >= 2
        ? logSheet
            .getRange(
              2,
              1,
              logSheet.getLastRow() - 1,
              1
            )
            .getDisplayValues()
            .flat()
        : [];


    let highestLogNumber = 0;

    existingLogIds.forEach(id => {

      const match =
        String(id).match(
          /^LOG-(\d+)$/
        );

      if (match) {

        highestLogNumber =
          Math.max(
            highestLogNumber,
            parseInt(match[1], 10)
          );

      }

    });


    const logId =
      `LOG-${String(highestLogNumber + 1).padStart(5, '0')}`;


    const logRecord = {

      'Log ID':
        logId,

      'Learner ID':
        learnerId,

      'Learner Name':
        getValue('Full Name'),

      'Document Type':
        'Apprenticeship Agreement',

      'Document Status':
        'Generated',

      'Google Doc ID':
        documentFile.getId(),

      'Google Doc Link':
        documentFile.getUrl(),

      'PDF ID':
        pdfFile.getId(),

      'PDF Link':
        pdfFile.getUrl(),

      'Generated Date':
        new Date(),

      'Generated By':
        Session.getEffectiveUser().getEmail(),

      'Template Used':
        'Apprenticeship Agreement',

      'Notes':
        'Generated by production agreement workflow'

    };


    const logRow =
      logHeaders.map(header =>
        logRecord[
          String(header).trim()
        ] ?? ''
      );


    logSheet.appendRow(logRow);


    // ------------------------------------------------
    // 12. Complete
    // ------------------------------------------------

    console.log(
      `Production agreement workflow completed for ${learnerId}.`
    );

  }

  finally {

    lock.releaseLock();

  }

}



/**
 * Generates a learner's Enrolment Pack. Mirrors
 * generateAgreementWorkflow_() (readiness checks, template
 * selection, placeholder replacement, PDF export, logging) for
 * the Enrolment Pack document type.
 */
function generateEnrolmentPackWorkflow_(learnerId) {

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {

    const ss = SpreadsheetApp.getActiveSpreadsheet();

    const learnerSheet = ss.getSheetByName('Learners');
    const settingsSheet = ss.getSheetByName('Settings');
    const logSheet = ss.getSheetByName('Document Logs');

    if (!learnerSheet || !settingsSheet || !logSheet) {
      throw new Error(
        'Could not find Learners, Settings or Document Logs.'
      );
    }


    // ------------------------------------------------
    // 1. Read learner data
    // ------------------------------------------------

    const learnerData =
      learnerSheet.getDataRange().getValues();

    const headers =
      learnerData[0];

    const headerIndex = buildHeaderIndex_(headers);

    let learnerRow = null;
    let learnerRowNumber = null;

    for (let i = 1; i < learnerData.length; i++) {

      if (
        String(
          learnerData[i][headerIndex['Learner ID']]
        ).trim() === learnerId
      ) {

        learnerRow = learnerData[i];
        learnerRowNumber = i + 1;
        break;
      }
    }


    if (!learnerRow) {
      throw new Error(
        `Learner ${learnerId} was not found.`
      );
    }


    const getValue = (header) => {

      const index =
        headerIndex[header];

      if (index === undefined) {
        return '';
      }

      return learnerRow[index] || '';
    };


    const setLearnerValue = (header, value) => {

      const index =
        headerIndex[header];

      if (index === undefined) {
        return;
      }

      learnerSheet
        .getRange(
          learnerRowNumber,
          index + 1
        )
        .setValue(value);
    };


    // ------------------------------------------------
    // 2. Duplicate protection
    // ------------------------------------------------

    const existingDocumentId =
      getValue('Enrolment Document ID');

    const existingPdfLink =
      getValue('Enrolment PDF Link');


    if (existingDocumentId || existingPdfLink) {

      console.log(
        `Enrolment Pack already exists for ${learnerId}. Nothing created.`
      );

      return;
    }


    // ------------------------------------------------
    // 3. Check essential information
    // ------------------------------------------------

    const requiredFields = [
      'Full Name',
      'Drive Folder ID',
      'Programme Code',
      'Apprenticeship Standard',
      'Apprenticeship Level',
      'Apprenticeship Start Date',
      'Practical Period Start Date',
      'Practical Period End Date',
      'Apprenticeship End Date',
      'Planned OTJ Hours'
    ];


    const missingFields =
      requiredFields.filter(
        field => !getValue(field)
      );


    if (missingFields.length) {

      throw new Error(
        `Enrolment Pack cannot be generated. Missing: ${missingFields.join(', ')}`
      );
    }


    // ------------------------------------------------
    // 4. Read Settings and select the right template
    //    for this learner's programme
    // ------------------------------------------------

    const settings = getSettingsMap_(ss);

    const programmeCode =
      getValue('Programme Code');

    const templateId =
      selectTemplateId_(programmeCode, 'ENROLMENT', settings);



    // ------------------------------------------------
    // 6. Build replacements
    //    (see formatDate_() and escapeRegex_() in utilities.gs)
    // ------------------------------------------------

    const employerAgreement =
      getLatestEmployerAgreement_(
        getValue('Full Name'),
        getValue('Company Name')
      );

    const replacements = {

      '{{EMPLOYER_CONTACT_FIRST_NAME}}':
        employerAgreement?.firstName || getValue('Line Manager First Name'),

      '{{EMPLOYER_CONTACT_SURNAME}}':
        employerAgreement?.surname || getValue('Line Manager Surname'),

      '{{EMPLOYER_CONTACT_EMAIL}}':
        employerAgreement?.email || getValue('Line Manager Email'),

      '{{EMPLOYER_CONTACT_TELEPHONE}}':
        employerAgreement?.mobile || getValue('Line Manager Telephone'),

      '{{EMPLOYER_CONTACT_POSITION}}':
        employerAgreement?.position || '',

      '{{EMPLOYER_E_SIGNATURE}}':
        employerAgreement?.signature || '',

      '{{LEARNER_ID}}':
        getValue('Learner ID'),

      '{{APPRENTICESHIP_TITLE}}':
        programmeDisplayName_(
          getValue('Programme Code'),
          getValue('Apprenticeship Standard')
        ),

      '{{FIRST_NAME}}':
        getValue('First Name'),

      '{{SURNAME}}':
        getValue('Surname'),

      '{{LEARNER_FULL_NAME}}':
        getValue('Full Name'),

      '{{DATE_OF_BIRTH}}':
        formatDate_(
          getValue('Date of Birth')
        ),

      '{{AGE_AT_ENROLMENT}}':
        getValue('Age at Enrolment'),

      '{{NI_NUMBER}}':
        getValue('NI Number'),

      '{{ULN}}':
        getValue('ULN'),

      '{{EMAIL_ADDRESS}}':
        getValue('Email Address'),

      '{{TELEPHONE_NUMBER}}':
        getValue('Telephone Number'),

      '{{ADDRESS_LINE_1}}':
        getValue('Address Line 1'),

      '{{ADDRESS_LINE_2}}':
        getValue('Address Line 2'),

      '{{TOWN_CITY}}':
        getValue('Town / City'),

      '{{COUNTY}}':
        getValue('County'),

      '{{POSTCODE}}':
        getValue('Postcode'),

      '{{WORK_ADDRESS}}':
        getValue('Work Address'),

      '{{DRIVING_LICENCE_NUMBER}}':
        getValue('Driving Licence Number'),

      '{{EMPLOYER_ID}}':
        getValue('Employer ID'),

      '{{EMPLOYER_NAME}}':
        getValue('Company Name'),

      '{{LINE_MANAGER_FIRST_NAME}}':
        getValue('Line Manager First Name'),

      '{{LINE_MANAGER_SURNAME}}':
        getValue('Line Manager Surname'),

      '{{LINE_MANAGER_EMAIL}}':
        getValue('Line Manager Email'),

      '{{LINE_MANAGER_TELEPHONE}}':
        getValue('Line Manager Telephone'),

      '{{EMPLOYMENT_START_DATE}}':
        formatDate_(
          getValue('Employment Start Date')
        ),

      '{{JOB_TITLE}}':
        getValue('Job Title'),

      '{{CONTRACTED_HOURS}}':
        getValue('Normal Working Hours'),

      '{{LENGTH_OF_EMPLOYMENT}}':
        getValue('Length of Employment at Enrolment (Months)'),

      '{{PROGRAMME_CODE}}':
        getValue('Programme Code'),

      '{{APPRENTICESHIP_STANDARD}}':
        getValue('Apprenticeship Standard'),

      '{{APPRENTICESHIP_LEVEL}}':
        getValue('Apprenticeship Level'),

      '{{APPRENTICESHIP_START_DATE}}':
        formatDate_(
          getValue('Apprenticeship Start Date')
        ),

      '{{PRACTICAL_PERIOD_START_DATE}}':
        formatDate_(
          getValue('Practical Period Start Date')
        ),

      '{{PRACTICAL_PERIOD_END_DATE}}':
        formatDate_(
          getValue('Practical Period End Date')
        ),

      '{{APPRENTICESHIP_END_DATE}}':
        formatDate_(
          getValue('Apprenticeship End Date')
        ),

      '{{PLANNED_OTJ_HOURS}}':
        getValue('Planned OTJ Hours'),

      '{{NORMAL_WORKING_HOURS}}':
        getValue('Normal Working Hours'),

      '{{TRAINING_WEEK_HOURS}}':
        getValue('Training Week Hours')
    };


    // ------------------------------------------------
    // 7. Find 04 - Enrolment Pack folder
    // ------------------------------------------------

    const learnerFolder =
      DriveApp.getFolderById(
        getValue('Drive Folder ID')
      );


    const enrolmentFolders =
      learnerFolder.getFoldersByName(
        '04 - Enrolment Pack'
      );


    if (!enrolmentFolders.hasNext()) {
      throw new Error(
        'Could not find 04 - Enrolment Pack folder.'
      );
    }


    const enrolmentFolder =
      enrolmentFolders.next();


    // ------------------------------------------------
    // 8. Create populated Google Doc
    // ------------------------------------------------

    const learnerName =
      getValue('Full Name');


    const documentName =
      `${learnerId} - ${learnerName} - Enrolment Pack`;


    const templateFile =
      DriveApp.getFileById(
        templateId
      );


    const documentFile =
      templateFile.makeCopy(
        documentName,
        enrolmentFolder
      );


    const doc =
      DocumentApp.openById(
        documentFile.getId()
      );


    const body =
      doc.getBody();


    Object.keys(replacements)
      .forEach(placeholder => {

        body.replaceText(
          escapeRegex_(placeholder),
          String(
            replacements[placeholder] ?? ''
          )
        );

      });


    doc.saveAndClose();


    // ------------------------------------------------
    // 9. Create PDF
    // ------------------------------------------------

    const pdfName =
      `${documentName}.pdf`;


    const pdfBlob =
      documentFile
        .getAs(MimeType.PDF)
        .setName(pdfName);


    const pdfFile =
      enrolmentFolder.createFile(
        pdfBlob
      );


    // ------------------------------------------------
    // 10. Update learner record
    // ------------------------------------------------

    setLearnerValue(
      'Enrolment Document ID',
      documentFile.getId()
    );


    setLearnerValue(
      'Enrolment PDF Link',
      pdfFile.getUrl()
    );


    setLearnerValue(
      'Enrolment Pack Status',
      'Generated'
    );


    setLearnerValue(
      'Last Updated',
      new Date()
    );


    setLearnerValue(
      'Automation Notes',
      `Enrolment Pack generated successfully for ${learnerId}.`
    );


    // ------------------------------------------------
    // 11. Generate Document Log ID
    // ------------------------------------------------

    const existingLogIds =
      logSheet.getLastRow() >= 2
        ? logSheet
            .getRange(
              2,
              1,
              logSheet.getLastRow() - 1,
              1
            )
            .getDisplayValues()
            .flat()
        : [];


    let highestLogNumber = 0;


    existingLogIds.forEach(id => {

      const match =
        String(id).match(
          /^LOG-(\d+)$/
        );

      if (match) {

        highestLogNumber =
          Math.max(
            highestLogNumber,
            Number(match[1])
          );

      }

    });


    const logId =
      `LOG-${String(
        highestLogNumber + 1
      ).padStart(5, '0')}`;


    // ------------------------------------------------
    // 12. Add Document Log entry
    // ------------------------------------------------

    const logHeaders =
      logSheet
        .getRange(
          1,
          1,
          1,
          logSheet.getLastColumn()
        )
        .getDisplayValues()[0];


    const logRecord = {

      'Log ID':
        logId,

      'Learner ID':
        learnerId,

      'Learner Name':
        learnerName,

      'Document Type':
        'Enrolment Pack',

      'Document Status':
        'Generated',

      'Google Doc ID':
        documentFile.getId(),

      'Google Doc Link':
        documentFile.getUrl(),

      'PDF ID':
        pdfFile.getId(),

      'PDF Link':
        pdfFile.getUrl(),

      'Generated Date':
        new Date(),

      'Generated By':
        Session.getEffectiveUser().getEmail(),

      'Template Used':
        'Enrolment Pack',

      'Notes':
        'Generated by production Enrolment Pack workflow'
    };


    const logRow =
      logHeaders.map(header =>

        logRecord[
          String(header).trim()
        ] ?? ''

      );


    logSheet.appendRow(
      logRow
    );


    // ------------------------------------------------
    // 13. Finished
    // ------------------------------------------------

    console.log(
      `Production Enrolment Pack workflow completed for ${learnerId}.`
    );

  }

  finally {
    lock.releaseLock();
  }

}



/**
 * Generates a learner's Learner Diagnostics document. Unlike the
 * Agreement and Enrolment Pack, this document type is not
 * programme-specific — it uses a single shared template — but
 * otherwise follows the same generate/export/log pattern.
 */
function generateLearnerDiagnosticsWorkflow_(learnerId) {

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {

    const ss =
      SpreadsheetApp.getActiveSpreadsheet();

    const learnerSheet =
      ss.getSheetByName('Learners');

    const settingsSheet =
      ss.getSheetByName('Settings');

    const logSheet =
      ss.getSheetByName('Document Logs');

    if (
      !learnerSheet ||
      !settingsSheet ||
      !logSheet
    ) {
      throw new Error(
        'Could not find Learners, Settings or Document Logs.'
      );
    }


    // ------------------------------------------------
    // 1. Read learner
    // ------------------------------------------------

    const learnerData =
      learnerSheet.getDataRange().getValues();

    const headers =
      learnerData[0];

    const headerIndex = buildHeaderIndex_(headers);

    let learnerRow = null;

    for (
      let i = 1;
      i < learnerData.length;
      i++
    ) {

      if (
        String(
          learnerData[i][
            headerIndex['Learner ID']
          ]
        ).trim() === learnerId
      ) {

        learnerRow =
          learnerData[i];

        break;
      }
    }


    if (!learnerRow) {
      throw new Error(
        `Learner ${learnerId} was not found.`
      );
    }


    const getValue = header => {

      const index =
        headerIndex[header];

      if (index === undefined) {
        return '';
      }

      return learnerRow[index] || '';
    };


    // ------------------------------------------------
    // 2. Validate required information
    // ------------------------------------------------

    const requiredFields = [
      'Full Name',
      'Drive Folder ID',
      'Programme Code'
    ];

    const missingFields =
      requiredFields.filter(
        field => !getValue(field)
      );

    if (missingFields.length) {

      throw new Error(
        `Learner Diagnostics cannot be generated. Missing: ${missingFields.join(', ')}`
      );
    }


    // ------------------------------------------------
    // 3. Read Settings
    //    (this document type isn't programme-specific,
    //    so it reads a single template ID directly
    //    rather than going through selectTemplateId_())
    // ------------------------------------------------

    const settings = getSettingsMap_(ss);

    const templateId =
      settings['LEARNER_DIAGNOSTICS_TEMPLATE_ID'];

    if (!templateId) {
      throw new Error(
        'LEARNER_DIAGNOSTICS_TEMPLATE_ID is not configured in Settings.'
      );
    }


    // ------------------------------------------------
    // 5. Employer form details
    //    (see formatDate_() and escapeRegex_() in utilities.gs)
    // ------------------------------------------------

    const employerAgreement =
      getLatestEmployerAgreement_(
        getValue('Full Name'),
        getValue('Company Name')
      );


    // ------------------------------------------------
    // 6. Build replacements
    // ------------------------------------------------

    const replacements = {

      '{{LEARNER_ID}}':
        getValue('Learner ID'),

      '{{FIRST_NAME}}':
        getValue('First Name'),

      '{{SURNAME}}':
        getValue('Surname'),

      '{{LEARNER_FULL_NAME}}':
        getValue('Full Name'),

      '{{DATE_OF_BIRTH}}':
        formatDate_(
          getValue('Date of Birth')
        ),

      '{{AGE_AT_ENROLMENT}}':
        getValue('Age at Enrolment'),

      '{{NI_NUMBER}}':
        getValue('NI Number'),

      '{{ULN}}':
        getValue('ULN'),

      '{{EMAIL_ADDRESS}}':
        getValue('Email Address'),

      '{{TELEPHONE_NUMBER}}':
        getValue('Telephone Number'),

      '{{ADDRESS_LINE_1}}':
        getValue('Address Line 1'),

      '{{ADDRESS_LINE_2}}':
        getValue('Address Line 2'),

      '{{TOWN_CITY}}':
        getValue('Town / City'),

      '{{COUNTY}}':
        getValue('County'),

      '{{POSTCODE}}':
        getValue('Postcode'),

      '{{WORK_ADDRESS}}':
        getValue('Work Address'),

      '{{DRIVING_LICENCE_NUMBER}}':
        getValue('Driving Licence Number'),

      '{{EMPLOYER_ID}}':
        getValue('Employer ID'),

      '{{EMPLOYER_NAME}}':
        getValue('Company Name'),

      '{{LINE_MANAGER_FIRST_NAME}}':
        getValue(
          'Line Manager First Name'
        ),

      '{{LINE_MANAGER_SURNAME}}':
        getValue(
          'Line Manager Surname'
        ),

      '{{LINE_MANAGER_EMAIL}}':
        getValue(
          'Line Manager Email'
        ),

      '{{LINE_MANAGER_TELEPHONE}}':
        getValue(
          'Line Manager Telephone'
        ),

      '{{EMPLOYER_CONTACT_FIRST_NAME}}':
        employerAgreement?.firstName ||
        getValue(
          'Line Manager First Name'
        ),

      '{{EMPLOYER_CONTACT_SURNAME}}':
        employerAgreement?.surname ||
        getValue(
          'Line Manager Surname'
        ),

      '{{EMPLOYER_CONTACT_EMAIL}}':
        employerAgreement?.email ||
        getValue(
          'Line Manager Email'
        ),

      '{{EMPLOYER_CONTACT_TELEPHONE}}':
        employerAgreement?.mobile ||
        getValue(
          'Line Manager Telephone'
        ),

      '{{EMPLOYER_CONTACT_POSITION}}':
        employerAgreement?.position || '',

      '{{EMPLOYER_E_SIGNATURE}}':
        employerAgreement?.signature || '',

      '{{EMPLOYMENT_START_DATE}}':
        formatDate_(
          getValue(
            'Employment Start Date'
          )
        ),

      '{{JOB_TITLE}}':
        getValue('Job Title'),

      '{{CONTRACTED_HOURS}}':
        getValue(
          'Normal Working Hours'
        ),

      '{{NORMAL_WORKING_HOURS}}':
        getValue(
          'Normal Working Hours'
        ),

      '{{LENGTH_OF_EMPLOYMENT}}':
        getValue(
          'Length of Employment at Enrolment (Months)'
        ),

      '{{PROGRAMME_CODE}}':
        getValue('Programme Code'),

      '{{APPRENTICESHIP_STANDARD}}':
        getValue(
          'Apprenticeship Standard'
        ),

      '{{APPRENTICESHIP_LEVEL}}':
        getValue(
          'Apprenticeship Level'
        ),

      '{{APPRENTICESHIP_TITLE}}':
        programmeDisplayName_(
          getValue('Programme Code'),
          getValue('Apprenticeship Standard')
        ),

      '{{APPRENTICESHIP_START_DATE}}':
        formatDate_(
          getValue(
            'Apprenticeship Start Date'
          )
        ),

      '{{PRACTICAL_PERIOD_START_DATE}}':
        formatDate_(
          getValue(
            'Practical Period Start Date'
          )
        ),

      '{{PRACTICAL_PERIOD_END_DATE}}':
        formatDate_(
          getValue(
            'Practical Period End Date'
          )
        ),

      '{{APPRENTICESHIP_END_DATE}}':
        formatDate_(
          getValue(
            'Apprenticeship End Date'
          )
        ),

      '{{PLANNED_OTJ_HOURS}}':
        getValue(
          'Planned OTJ Hours'
        ),

      '{{TRAINING_WEEK_HOURS}}':
        getValue(
          'Training Week Hours'
        )

    };


    // ------------------------------------------------
    // 7. Find 05 - Supporting Documents
    // ------------------------------------------------

    const learnerFolder =
      DriveApp.getFolderById(
        getValue('Drive Folder ID')
      );


    const supportingFolders =
      learnerFolder.getFoldersByName(
        '05 - Supporting Documents'
      );


    if (!supportingFolders.hasNext()) {

      throw new Error(
        'Could not find 05 - Supporting Documents folder.'
      );
    }


    const supportingFolder =
      supportingFolders.next();


    // ------------------------------------------------
    // 8. Duplicate protection
    // ------------------------------------------------

    const learnerName =
      getValue('Full Name');

    const documentName =
      `${learnerId} - ${learnerName} - Learner Diagnostics`;


    const existingDocs =
      supportingFolder
        .getFilesByName(
          documentName
        );


    const existingPdfs =
      supportingFolder
        .getFilesByName(
          `${documentName}.pdf`
        );


    if (
      existingDocs.hasNext() ||
      existingPdfs.hasNext()
    ) {

      console.log(
        `Learner Diagnostics already exists for ${learnerId}. Nothing created.`
      );

      return;
    }


    // ------------------------------------------------
    // 9. Create populated Google Doc
    // ------------------------------------------------

    const templateFile =
      DriveApp.getFileById(
        templateId
      );


    const documentFile =
      templateFile.makeCopy(
        documentName,
        supportingFolder
      );


    const doc =
      DocumentApp.openById(
        documentFile.getId()
      );


    const body =
      doc.getBody();


    Object.keys(replacements)
      .forEach(placeholder => {

        body.replaceText(
          escapeRegex_(placeholder),
          String(
            replacements[
              placeholder
            ] ?? ''
          )
        );

      });


    doc.saveAndClose();


    // ------------------------------------------------
    // 10. Create PDF copy
    // ------------------------------------------------

    const pdfName =
      `${documentName}.pdf`;


    const pdfBlob =
      documentFile
        .getAs(MimeType.PDF)
        .setName(pdfName);


    const pdfFile =
      supportingFolder.createFile(
        pdfBlob
      );


    // ------------------------------------------------
    // 11. Add Document Log
    // ------------------------------------------------

    const logHeaders =
      logSheet
        .getRange(
          1,
          1,
          1,
          logSheet.getLastColumn()
        )
        .getDisplayValues()[0];


    const existingLogIds =
      logSheet.getLastRow() >= 2
        ? logSheet
            .getRange(
              2,
              1,
              logSheet.getLastRow() - 1,
              1
            )
            .getDisplayValues()
            .flat()
        : [];


    let highestLogNumber = 0;


    existingLogIds.forEach(id => {

      const match =
        String(id).match(
          /^LOG-(\d+)$/
        );

      if (match) {

        highestLogNumber =
          Math.max(
            highestLogNumber,
            Number(match[1])
          );
      }

    });


    const logId =
      `LOG-${String(
        highestLogNumber + 1
      ).padStart(5, '0')}`;


    const logRecord = {

      'Log ID':
        logId,

      'Learner ID':
        learnerId,

      'Learner Name':
        learnerName,

      'Document Type':
        'Learner Diagnostics',

      'Document Status':
        'Generated',

      'Google Doc ID':
        documentFile.getId(),

      'Google Doc Link':
        documentFile.getUrl(),

      'PDF ID':
        pdfFile.getId(),

      'PDF Link':
        pdfFile.getUrl(),

      'Generated Date':
        new Date(),

      'Generated By':
        Session
          .getEffectiveUser()
          .getEmail(),

      'Template Used':
        'Learner Diagnostics',

      'Notes':
        'Generated by Learner Diagnostics workflow'

    };


    const logRow =
      logHeaders.map(header =>
        logRecord[
          String(header).trim()
        ] ?? ''
      );


    logSheet.appendRow(
      logRow
    );


    console.log(
      `Learner Diagnostics generated successfully for ${learnerId}.`
    );

  }

  finally {

    lock.releaseLock();

  }

}



/**
 * Marks a learner's enrolment as complete once all required
 * documents have been generated and signed. Called from
 * handlePaperworkStatusEdit() when the Signature Request Status
 * is set to "Completed".
 */
function finaliseEnrolment_(learnerId) {

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const learnerSheet = ss.getSheetByName('Learners');
    const logSheet = ss.getSheetByName('Document Logs');

    if (!learnerSheet || !logSheet) {
      throw new Error(
        'Could not find Learners or Document Logs sheet.'
      );
    }

    const data =
      learnerSheet.getDataRange().getValues();

    const headers = data[0];
    const headerIndex = buildHeaderIndex_(headers);

    let learnerRow = null;
    let learnerRowNumber = null;

    for (let i = 1; i < data.length; i++) {
      if (
        String(
          data[i][headerIndex['Learner ID']]
        ).trim() === learnerId
      ) {
        learnerRow = data[i];
        learnerRowNumber = i + 1;
        break;
      }
    }

    if (!learnerRow) {
      throw new Error(`Learner ${learnerId} not found.`);
    }

    const getValue = (header) => {
      const index = headerIndex[header];
      if (index === undefined) {
        return '';
      }
      return learnerRow[index] || '';
    };

    const setValue = (header, value) => {
      const index = headerIndex[header];
      if (index === undefined) {
        return;
      }
      learnerSheet
        .getRange(learnerRowNumber, index + 1)
        .setValue(value);
    };

    const completedDate =
      getValue('Enrolment Completed Date');

    if (completedDate) {
      setValue('Signature Request Status', 'Completed');
      setValue('Overall Enrolment Status', 'Completed');
      return;
    }

    // Google Workspace eSignature is the signing system of record.
    // At this point staff should only select Completed after the
    // completed signed documents have been returned by Google.

    const agreementDocumentId =
      String(getValue('Agreement Document ID')).trim();

    const enrolmentDocumentId =
      String(getValue('Enrolment Document ID')).trim();

    const missing = [];

    if (!agreementDocumentId) {
      missing.push('Agreement Document ID');
    }

    if (!enrolmentDocumentId) {
      missing.push('Enrolment Document ID');
    }

    if (missing.length) {
      throw new Error(
        `Cannot complete ${learnerId}. Missing: ${missing.join(', ')}`
      );
    }

    const completedAt = new Date();

    setValue('Signature Request Status', 'Completed');
    setValue('Overall Enrolment Status', 'Completed');
    setValue('Enrolment Completed Date', completedAt);
    setValue('Last Updated', completedAt);
    setValue(
      'Automation Notes',
      `Enrolment marked completed for ${learnerId}. Google Workspace eSignature records are the signature audit trail.`
    );

    SpreadsheetApp.flush();

    const logHeaders =
      logSheet
        .getRange(1, 1, 1, logSheet.getLastColumn())
        .getDisplayValues()[0];

    const existingLogIds =
      logSheet.getLastRow() >= 2
        ? logSheet
            .getRange(
              2,
              1,
              logSheet.getLastRow() - 1,
              1
            )
            .getDisplayValues()
            .flat()
        : [];

    let highestLogNumber = 0;

    existingLogIds.forEach(id => {
      const match = String(id).match(/^LOG-(\d+)$/);
      if (match) {
        highestLogNumber = Math.max(
          highestLogNumber,
          Number(match[1])
        );
      }
    });

    const logId =
      `LOG-${String(highestLogNumber + 1).padStart(5, '0')}`;

    const logRecord = {
      'Log ID': logId,
      'Learner ID': learnerId,
      'Learner Name': String(getValue('Full Name')).trim(),
      'Document Type': 'Enrolment Completion',
      'Document Status': 'Completed',
      'Generated Date': completedAt,
      'Generated By': Session.getEffectiveUser().getEmail(),
      'Template Used': '',
      'Notes': 'Enrolment marked completed after Google Workspace eSignature workflow.'
    };

    const logRow =
      logHeaders.map(header =>
        logRecord[String(header).trim()] ?? ''
      );

    logSheet.appendRow(logRow);

    console.log(
      `Enrolment marked completed for ${learnerId}.`
    );

  } finally {
    lock.releaseLock();
  }
}

