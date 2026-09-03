// ============================================================
// CONFIGURATION
// ============================================================
//
// In the live system, configuration values (Drive folder IDs,
// Google Docs template IDs, ID prefixes, etc.) are stored in a
// "Settings" sheet inside the spreadsheet, rather than hardcoded
// in the script. This lets non-technical staff update which
// template is used for a programme without touching code.
//
// The CONFIG object below documents the same set of keys with
// placeholder values, so this public version is runnable against
// your own spreadsheet/Drive setup after you fill in real IDs.
// It is also used as a fallback if no "Settings" sheet is found.
//
// Replace every placeholder value with your own before use.

const CONFIG = {
  // Google Drive folder that contains all learner sub-folders
  ROOT_FOLDER_ID: 'ROOT_FOLDER_ID',

  // Prefixes used when generating human-readable record IDs
  LEARNER_ID_PREFIX: 'LRN',
  EMPLOYER_ID_PREFIX: 'EMP',
  LOG_ID_PREFIX: 'LOG',

  // Google Docs template IDs, one pair per apprenticeship programme.
  // Programme A / Programme B stand in for the organisation's real
  // apprenticeship pathways (renamed for this public repository).
  AGREEMENT_TEMPLATE_ID_PROGRAMME_A: 'PROGRAMME_A_AGREEMENT_TEMPLATE_ID',
  AGREEMENT_TEMPLATE_ID_PROGRAMME_B: 'PROGRAMME_B_AGREEMENT_TEMPLATE_ID',
  ENROLMENT_TEMPLATE_ID_PROGRAMME_A: 'PROGRAMME_A_ENROLMENT_TEMPLATE_ID',
  ENROLMENT_TEMPLATE_ID_PROGRAMME_B: 'PROGRAMME_B_ENROLMENT_TEMPLATE_ID',
  LEARNER_DIAGNOSTICS_TEMPLATE_ID: 'LEARNER_DIAGNOSTICS_TEMPLATE_ID',

  ACADEMIC_YEAR: '2026-27',
  SYSTEM_VERSION: '1.0'
};

/**
 * Reads configuration values from the spreadsheet's "Settings"
 * sheet into a plain key/value object, falling back to the CONFIG
 * placeholders above if no Settings sheet exists yet.
 *
 * Centralising this lookup avoids repeating the same "read two
 * columns starting at row 2" logic in every workflow function.
 */
function getSettingsMap_(ss) {

  const settingsSheet = ss.getSheetByName('Settings');

  if (!settingsSheet) {
    return CONFIG;
  }

  const settingsRows = settingsSheet
    .getRange(
      2,
      1,
      Math.max(settingsSheet.getLastRow() - 1, 1),
      2
    )
    .getDisplayValues();

  const settings = {};

  settingsRows.forEach(row => {
    if (row[0]) {
      settings[row[0].trim()] = row[1];
    }
  });

  return Object.keys(settings).length ? settings : CONFIG;
}
