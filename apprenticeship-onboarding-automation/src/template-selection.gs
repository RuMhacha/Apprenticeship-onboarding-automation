// ============================================================
// TEMPLATE SELECTION
// ============================================================
//
// Each apprenticeship programme has its own set of onboarding
// document templates. This is the core piece of conditional
// business logic in the project: given a learner's programme code
// and which document is being produced, look up the correct
// Google Docs template ID from Settings/config.
//
// The three document-generation workflows (Agreement, Enrolment
// Pack, Learner Diagnostics) all call this instead of repeating
// their own if/else programme checks.

/**
 * Maps a raw "which programme would you like to enrol on?" form
 * answer to the internal programme code used everywhere else in
 * the system (Learners sheet, template selection, etc).
 *
 * More than one form answer can map to the same programme code —
 * this is how alternate wording on the form ("Programme A" vs
 * "Programme A (Alternate)") still resolves to a single programme.
 */
function resolveProgrammeCode_(courseAnswer) {

  const courseMap = {
    'Programme A': 'PROG-A',
    'Programme A (Alternate)': 'PROG-A',
    'Programme B': 'PROG-B'
  };

  return courseMap[courseAnswer];
}

/**
 * Selects the appropriate onboarding document template ID for a
 * given programme code and document type, reading from the
 * Settings/config values.
 *
 * documentType is one of: 'AGREEMENT', 'ENROLMENT'
 */
function selectTemplateId_(programmeCode, documentType, settings) {

  const settingsKeysByDocumentType = {
    AGREEMENT: {
      'PROG-A': 'AGREEMENT_TEMPLATE_ID_PROGRAMME_A',
      'PROG-B': 'AGREEMENT_TEMPLATE_ID_PROGRAMME_B'
    },
    ENROLMENT: {
      'PROG-A': 'ENROLMENT_TEMPLATE_ID_PROGRAMME_A',
      'PROG-B': 'ENROLMENT_TEMPLATE_ID_PROGRAMME_B'
    }
  };

  const settingsKey =
    settingsKeysByDocumentType[documentType]?.[programmeCode];

  const templateId = settingsKey ? settings[settingsKey] : '';

  if (!templateId) {
    throw new Error(
      `No ${documentType} template configured for programme ${programmeCode}.`
    );
  }

  return templateId;
}

/**
 * Returns the display name to use for a programme in generated
 * documents (falls back to the learner's stored apprenticeship
 * standard for any programme code not explicitly named here).
 */
function programmeDisplayName_(programmeCode, fallback) {

  const displayNames = {
    'PROG-A': 'Programme A',
    'PROG-B': 'Programme B'
  };

  return displayNames[programmeCode] || fallback;
}
