/**
 * Processes batches of prompts and writes the responses to a specified output column.
 * @param {number} startRowGoogleName The starting row number for processing.
 * @param {number} endRowGoogleName The ending row number for processing.
 * @param {string} inColGoogleName The column letter where the prompts are.
 * @param {string} inColGoogleName2 The column letter where the prompts are.
 * @param {string} outColGoogleName The column letter where the results should be written.
 */
function processGoogleNameBatches(startRowGoogleName, endRowGoogleName, inColGoogleName, inColGoogleName2, outColGoogleName) {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const promptsRange = `${inColGoogleName}${startRowGoogleName}:${inColGoogleName}${endRowGoogleName}`;
    const prompts = sheet.getRange(promptsRange).getValues();
    const promptsRange2 = `${inColGoogleName2}${startRowGoogleName}:${inColGoogleName2}${endRowGoogleName}`;
    const prompts2 = sheet.getRange(promptsRange2).getValues();
    let results = [];
  
    prompts.forEach((row, index) => {
      if (row[0] !== "") {
        try {
          let response = GoogleName(row[0], prompts2[index][0]); // Simulate an API call
          results.push([response]);
        } catch (e) {
          console.error('Error processing prompt: ' + row[0], e);
          results.push(["Error: " + e.toString()]);
        }
      } else {
        results.push([""]);
      }
    });
  
    const outputRange = `${outColGoogleName}${startRowGoogleName}:${outColGoogleName}${(startRowGoogleName + results.length - 1)}`;
    sheet.getRange(outputRange).setValues(results);
  }
  
  /**
   * Initializes the batch processing and sets up a trigger to continue processing.
   * @param {number} startRowGoogleName The starting row number for processing.
   * @param {number} endRowGoogleName The ending row number for processing.
   * @param {string} inColGoogleName The column letter where the prompts are.
   * @param {string} inColGoogleName2 The column letter where the prompts are.
   * @param {string} outColGoogleName The column letter where the results should be written.
   */
  function startGoogleNameProcessing(startRowGoogleName=51, endRowGoogleName=3973, inColGoogleName="C", inColGoogleName2="F", outColGoogleName="AI") {
    deleteExistingTriggers('continueGoogleNameProcessing');
  
    const batchSize = 30; // Adjust this number based on the typical processing time per batch to stay under the 6-minute script runtime limit
    let currentendRowGoogleName = startRowGoogleName + batchSize - 1;
    if (currentendRowGoogleName > endRowGoogleName) currentendRowGoogleName = endRowGoogleName;
  
    // Store settings in script properties
    const properties = PropertiesService.getScriptProperties();
    properties.setProperties({
      'inColGoogleName': inColGoogleName,
      'inColGoogleName2': inColGoogleName2,
      'outColGoogleName': outColGoogleName,
      'endRowGoogleName': endRowGoogleName.toString(),
      'lastRowProcessedGoogleName': currentendRowGoogleName.toString()
    });
  
    // Start the first batch
    processGoogleNameBatches(startRowGoogleName, currentendRowGoogleName, inColGoogleName, inColGoogleName2, outColGoogleName);
  
    if (currentendRowGoogleName < endRowGoogleName) {
      ScriptApp.newTrigger("continueGoogleNameProcessing")
        .timeBased()
        .after(10000) // Wait 10 seconds before continuing
        .create();
    }
  }
  
  /**
   * Continues processing the next batch after a delay.
   */
  function continueGoogleNameProcessing() {
    const properties = PropertiesService.getScriptProperties();
    const inColGoogleName = properties.getProperty('inColGoogleName');
    const inColGoogleName2 = properties.getProperty('inColGoogleName2');
    const outColGoogleName = properties.getProperty('outColGoogleName');
    const endRowGoogleName = parseInt(properties.getProperty('endRowGoogleName'));
    let lastRow = parseInt(properties.getProperty('lastRowProcessedGoogleName'));
  
    const batchSize = 30;
    const startRowGoogleName = lastRow + 1;
    let nextendRowGoogleName = startRowGoogleName + batchSize - 1;
    if (nextendRowGoogleName > endRowGoogleName) nextendRowGoogleName = endRowGoogleName;
  
    // Process the next batch
    processGoogleNameBatches(startRowGoogleName, nextendRowGoogleName, inColGoogleName, inColGoogleName2, outColGoogleName);
  
    // Update the last processed row
    properties.setProperty('lastRowProcessedGoogleName', nextendRowGoogleName.toString());
  
    if (nextendRowGoogleName < endRowGoogleName) {
      deleteExistingTriggers('continueGoogleNameProcessing');
      ScriptApp.newTrigger("continueGoogleNameProcessing")
        .timeBased()
        .after(10000) // Wait 10 seconds before continuing again
        .create();
    } else {
      // All batches are processed, delete the trigger
      deleteExistingTriggers('continueGoogleNameProcessing');
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
  