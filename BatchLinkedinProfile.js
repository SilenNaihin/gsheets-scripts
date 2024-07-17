/**
 * Initializes the batch processing and sets up a trigger to continue processing.
 * @param {number} startRowLinkedinProfile The starting row number for processing.
 * @param {number} endRowLinkedinProfile The ending row number for processing.
 * @param {string} inColLinkedinProfile The column letter where the name of the person is.
 * @param {string} inColLinkedinProfile2 The column letter where the name if the company is.
 * @param {string} outColLinkedinProfile The column letter where the results should be written.
 */
function startLinkedinProfileProcessing(
  startRowLinkedinProfile = 2,
  endRowLinkedinProfile = 100,
  inColLinkedinProfile = 'C',
  inColLinkedinProfile2 = 'F',
  outColLinkedinProfile = 'AI'
) {
  deleteExistingTriggers('continueLinkedinProfileProcessing');

  const batchSize = 30; // Adjust this number based on the typical processing time per batch to stay under the 6-minute script runtime limit
  let currentendRowLinkedinProfile = startRowLinkedinProfile + batchSize - 1;
  if (currentendRowLinkedinProfile > endRowLinkedinProfile)
    currentendRowLinkedinProfile = endRowLinkedinProfile;

  // Store settings in script properties
  const properties = PropertiesService.getScriptProperties();
  properties.setProperties({
    inColLinkedinProfile: inColLinkedinProfile,
    inColLinkedinProfile2: inColLinkedinProfile2,
    outColLinkedinProfile: outColLinkedinProfile,
    endRowLinkedinProfile: endRowLinkedinProfile.toString(),
    lastRowProcessedLinkedinProfile: currentendRowLinkedinProfile.toString(),
  });

  // Start the first batch
  processLinkedinProfileBatches(
    startRowLinkedinProfile,
    currentendRowLinkedinProfile,
    inColLinkedinProfile,
    inColLinkedinProfile2,
    outColLinkedinProfile
  );

  if (currentendRowLinkedinProfile < endRowLinkedinProfile) {
    ScriptApp.newTrigger('continueLinkedinProfileProcessing')
      .timeBased()
      .after(10000) // Wait 10 seconds before continuing
      .create();
  }
}

/**
 * Processes batches of prompts and writes the responses to a specified output column.
 * @param {number} startRowLinkedinProfile The starting row number for processing.
 * @param {number} endRowLinkedinProfile The ending row number for processing.
 * @param {string} inColLinkedinProfile The column letter where the name of the person is.
 * @param {string} inColLinkedinProfile2 The column letter where the name if the company is.
 * @param {string} outColLinkedinProfile The column letter where the results should be written.
 */
function processLinkedinProfileBatches(
  startRowLinkedinProfile,
  endRowLinkedinProfile,
  inColLinkedinProfile,
  inColLinkedinProfile2,
  outColLinkedinProfile
) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const promptsRange = `${inColLinkedinProfile}${startRowLinkedinProfile}:${inColLinkedinProfile}${endRowLinkedinProfile}`;
  const prompts = sheet.getRange(promptsRange).getValues();
  const promptsRange2 = `${inColLinkedinProfile2}${startRowLinkedinProfile}:${inColLinkedinProfile2}${endRowLinkedinProfile}`;
  const prompts2 = sheet.getRange(promptsRange2).getValues();
  let results = [];

  prompts.forEach((row, index) => {
    if (row[0] !== '') {
      try {
        let response = LinkedinProfile(row[0], prompts2[index][0]);
        results.push([response]);
      } catch (e) {
        console.error('Error processing prompt: ' + row[0], e);
        results.push(['Error: ' + e.toString()]);
      }
    } else {
      results.push(['']);
    }
  });

  const outputRange = `${outColLinkedinProfile}${startRowLinkedinProfile}:${outColLinkedinProfile}${
    startRowLinkedinProfile + results.length - 1
  }`;
  sheet.getRange(outputRange).setValues(results);
}

/**
 * Continues processing the next batch after a delay.
 */
function continueLinkedinProfileProcessing() {
  const properties = PropertiesService.getScriptProperties();
  const inColLinkedinProfile = properties.getProperty('inColLinkedinProfile');
  const inColLinkedinProfile2 = properties.getProperty('inColLinkedinProfile2');
  const outColLinkedinProfile = properties.getProperty('outColLinkedinProfile');
  const endRowLinkedinProfile = parseInt(
    properties.getProperty('endRowLinkedinProfile')
  );
  let lastRow = parseInt(
    properties.getProperty('lastRowProcessedLinkedinProfile')
  );

  const batchSize = 30;
  const startRowLinkedinProfile = lastRow + 1;
  let nextendRowLinkedinProfile = startRowLinkedinProfile + batchSize - 1;
  if (nextendRowLinkedinProfile > endRowLinkedinProfile)
    nextendRowLinkedinProfile = endRowLinkedinProfile;

  // Process the next batch
  processLinkedinProfileBatches(
    startRowLinkedinProfile,
    nextendRowLinkedinProfile,
    inColLinkedinProfile,
    inColLinkedinProfile2,
    outColLinkedinProfile
  );

  // Update the last processed row
  properties.setProperty(
    'lastRowProcessedLinkedinProfile',
    nextendRowLinkedinProfile.toString()
  );

  if (nextendRowLinkedinProfile < endRowLinkedinProfile) {
    deleteExistingTriggers('continueLinkedinProfileProcessing');
    ScriptApp.newTrigger('continueLinkedinProfileProcessing')
      .timeBased()
      .after(10000) // Wait 10 seconds before continuing again
      .create();
  } else {
    // All batches are processed, delete the trigger
    deleteExistingTriggers('continueLinkedinProfileProcessing');
    console.log('All rows have been processed.');
  }
}

/**
 * Deletes all triggers associated with a given function name.
 * @param {string} functionName The name of the function whose triggers need to be deleted.
 */
function deleteExistingTriggers(functionName) {
  var allTriggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < allTriggers.length; i++) {
    if (allTriggers[i].getHandlerFunction() === functionName) {
      ScriptApp.deleteTrigger(allTriggers[i]);
    }
  }
}
