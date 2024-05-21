/**
 * Initializes the batch processing and sets up a trigger to continue processing.
 * @param {number} startRowGPT The starting row number for processing.
 * @param {number} endRowGPT The ending row number for processing.
 * @param {string} inColGPT The column letter where the prompts are.
 * @param {string} outColGPT The column letter where the results should be written.
 */
function startGPTProcessing(
  startRowGPT = 2,
  endRowGPT = 100,
  inColGPT = 'AO',
  outColGPT = 'AP'
) {
  deleteExistingTriggers('continueGPTProcessing');

  const batchSize = 10; // Adjust this number based on the typical processing time per batch to stay under the 6-minute script runtime limit
  let currentendRowGPT = startRowGPT + batchSize - 1;
  if (currentendRowGPT > endRowGPT) currentendRowGPT = endRowGPT;

  // Store settings in script properties
  const properties = PropertiesService.getScriptProperties();
  properties.setProperties({
    inColGPT: inColGPT,
    outColGPT: outColGPT,
    endRowGPT: endRowGPT.toString(),
    lastRowProcessedGPT: currentendRowGPT.toString(),
  });

  // Start the first batch
  processGPTBatches(startRowGPT, currentendRowGPT, inColGPT, outColGPT);

  if (currentendRowGPT < endRowGPT) {
    ScriptApp.newTrigger('continueGPTProcessing')
      .timeBased()
      .after(10000) // Wait 10 seconds before continuing
      .create();
  }
}

/**
 * Processes batches of prompts and writes the responses to a specified output column.
 * @param {number} startRowGPT The starting row number for processing.
 * @param {number} endRowGPT The ending row number for processing.
 * @param {string} inColGPT The column letter where the prompts are.
 * @param {string} outColGPT The column letter where the results should be written.
 */
function processGPTBatches(startRowGPT, endRowGPT, inColGPT, outColGPT) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const promptsRange = `${inColGPT}${startRowGPT}:${inColGPT}${endRowGPT}`;
  const prompts = sheet.getRange(promptsRange).getValues();
  let results = [];

  prompts.forEach((row, index) => {
    if (row[0] !== '') {
      try {
        let response = GPT(row[0]); // Simulate an API call
        results.push([response]);
      } catch (e) {
        console.error('Error processing prompt: ' + row[0], e);
        results.push(['Error: ' + e.toString()]);
      }
    } else {
      results.push(['']);
    }
  });

  const outputRange = `${outColGPT}${startRowGPT}:${outColGPT}${
    startRowGPT + results.length - 1
  }`;
  sheet.getRange(outputRange).setValues(results);
}

/**
 * Continues processing the next batch after a delay.
 */
function continueGPTProcessing() {
  const properties = PropertiesService.getScriptProperties();
  const inColGPT = properties.getProperty('inColGPT');
  const outColGPT = properties.getProperty('outColGPT');
  const endRowGPT = parseInt(properties.getProperty('endRowGPT'));
  let lastRow = parseInt(properties.getProperty('lastRowProcessedGPT'));

  const batchSize = 10;
  const startRowGPT = lastRow + 1;
  let nextendRowGPT = startRowGPT + batchSize - 1;
  if (nextendRowGPT > endRowGPT) nextendRowGPT = endRowGPT;

  // Process the next batch
  processGPTBatches(startRowGPT, nextendRowGPT, inColGPT, outColGPT);

  // Update the last processed row
  properties.setProperty('lastRowProcessedGPT', nextendRowGPT.toString());

  if (nextendRowGPT < endRowGPT) {
    deleteExistingTriggers('continueGPTProcessing');
    ScriptApp.newTrigger('continueGPTProcessing')
      .timeBased()
      .after(10000) // Wait 10 seconds before continuing again
      .create();
  } else {
    // All batches are processed, delete the trigger
    deleteExistingTriggers('continueGPTProcessing');
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
