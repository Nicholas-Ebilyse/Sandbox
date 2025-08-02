/**
 * Central configuration for IDs and labels.
 * Values are loaded from script properties so they can be updated without code changes.
 */
function getConfig() {
  const props = PropertiesService.getScriptProperties();
  return {
    GOOGLE_SHEET_ID: props.getProperty('GOOGLE_SHEET_ID'),
    PROJECTS_SHEET_NAME: props.getProperty('PROJECTS_SHEET_NAME'),
    LOGS_SHEET_NAME: props.getProperty('LOGS_SHEET_NAME'),
    TARGET_DRIVE_FOLDER_ID: props.getProperty('TARGET_DRIVE_FOLDER_ID'),
    GMAIL_LABEL_TO_PROCESS: props.getProperty('GMAIL_LABEL_TO_PROCESS'),
    GMAIL_LABEL_PROCESSED_OK: props.getProperty('GMAIL_LABEL_PROCESSED_OK'),
    GMAIL_LABEL_PROCESSED_ERROR: props.getProperty('GMAIL_LABEL_PROCESSED_ERROR'),
    GMAIL_LABEL_PROCESSED_AUTRES: props.getProperty('GMAIL_LABEL_PROCESSED_AUTRES'),
  };
}
