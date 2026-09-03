// ============================================================
// SHARED UTILITIES
// ============================================================

/**
 * Converts a 1-based column number to its spreadsheet letter
 * (e.g. 1 -> 'A', 28 -> 'AB').
 */
function columnNumberToLetter(column) {

  let letter = '';

  while (column > 0) {
    const remainder = (column - 1) % 26;
    letter = String.fromCharCode(65 + remainder) + letter;
    column = Math.floor((column - 1) / 26);
  }

  return letter;
}

/**
 * Escapes regex special characters so a literal string (e.g. a
 * {{PLACEHOLDER}} token) can be safely passed to
 * Body.replaceText(), which expects a regular expression.
 */
function escapeRegex_(text) {

  return text.replace(
    /[.*+?^${}()|[\]\\]/g,
    '\\$&'
  );
}

/**
 * Formats a spreadsheet Date value as dd/MM/yyyy for use in
 * generated documents. Non-Date values are returned unchanged
 * (as a string) so blank or already-formatted cells don't error.
 */
function formatDate_(value) {

  if (!value) {
    return '';
  }

  if (value instanceof Date) {
    return Utilities.formatDate(
      value,
      Session.getScriptTimeZone(),
      'dd/MM/yyyy'
    );
  }

  return String(value);
}

/**
 * Builds a { headerName: columnIndex } lookup from a header row,
 * so the rest of the code can read/write sheet values by column
 * name instead of by position. Used throughout the form-processing
 * and document-generation workflows.
 */
function buildHeaderIndex_(headers) {

  const headerIndex = {};

  headers.forEach((header, index) => {
    if (header) {
      headerIndex[String(header).trim()] = index;
    }
  });

  return headerIndex;
}

/**
 * Lowercases, trims, and collapses internal whitespace so that
 * two values entered inconsistently by different people (extra
 * spaces, mixed case) can still be matched as "the same" record.
 * Used when matching a learner's employer-form submission back to
 * their original learner-form submission.
 */
function normaliseForMatch_(value) {

  return String(value || '')
    .trim()
    .toLowerCase()
    .replace(/\s+/g, ' ');
}
