/**
 * Processes batches of prompts and writes the responses to a specified output column.
 * @param {number} startRowUrlDesc The starting row number for processing.
 * @param {number} endRowUrlDesc The ending row number for processing.
 * @param {string} inColUrlDesc The column letter where the prompts are.
 * @param {string} outColUrlDesc The column letter where the results should be written.
 */
function processUrlDescBatches(startRowUrlDesc, endRowUrlDesc, inColUrlDesc, outColUrlDesc) {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const promptsRange = `${inColUrlDesc}${startRowUrlDesc}:${inColUrlDesc}${endRowUrlDesc}`;
    const prompts = sheet.getRange(promptsRange).getValues();
    let results = [];
  
    prompts.forEach((row, index) => {
      if (row[0] !== "") {
        try {
          let response = UrlDescription(row[0]); // Simulate an API call
          results.push([response]);
        } catch (e) {
          console.error('Error processing prompt: ' + row[0], e);
          results.push(["Error: " + e.toString()]);
        }
      } else {
        results.push([""]);
      }
    });
  
    const outputRange = `${outColUrlDesc}${startRowUrlDesc}:${outColUrlDesc}${(startRowUrlDesc + results.length - 1)}`;
    sheet.getRange(outputRange).setValues(results);
  }
  
  /**
   * Initializes the batch processing and sets up a trigger to continue processing.
   * @param {number} startRowUrlDesc The starting row number for processing.
   * @param {number} endRowUrlDesc The ending row number for processing.
   * @param {string} inColUrlDesc The column letter where the prompts are.
   * @param {string} outColUrlDesc The column letter where the results should be written.
   */
  function startUrlDescProcessing(startRowUrlDesc=1321, endRowUrlDesc=3973, inColUrlDesc="N", outColUrlDesc="AH") {
    deleteExistingTriggers('continueUrlDescProcessing');
  
    const batchSize = 30; // Adjust this number based on the typical processing time per batch to stay under the 6-minute script runtime limit
    let currentendRowUrlDesc = startRowUrlDesc + batchSize - 1;
    if (currentendRowUrlDesc > endRowUrlDesc) currentendRowUrlDesc = endRowUrlDesc;
  
    // Store settings in script properties
    const properties = PropertiesService.getScriptProperties();
    properties.setProperties({
      'inColUrlDesc': inColUrlDesc,
      'outColUrlDesc': outColUrlDesc,
      'endRowUrlDesc': endRowUrlDesc.toString(),
      'lastRowProcessedUrlDesc': currentendRowUrlDesc.toString()
    });
  
    // Start the first batch
    processUrlDescBatches(startRowUrlDesc, currentendRowUrlDesc, inColUrlDesc, outColUrlDesc);
  
    if (currentendRowUrlDesc < endRowUrlDesc) {
      ScriptApp.newTrigger("continueUrlDescProcessing")
        .timeBased()
        .after(10000) // Wait 10 seconds before continuing
        .create();
    }
  }
  
  /**
   * Continues processing the next batch after a delay.
   */
  function continueUrlDescProcessing() {
    const properties = PropertiesService.getScriptProperties();
    const inColUrlDesc = properties.getProperty('inColUrlDesc');
    const outColUrlDesc = properties.getProperty('outColUrlDesc');
    const endRowUrlDesc = parseInt(properties.getProperty('endRowUrlDesc'));
    let lastRow = parseInt(properties.getProperty('lastRowProcessedUrlDesc'));
  
    const batchSize = 30;
    const startRowUrlDesc = lastRow + 1;
    let nextendRowUrlDesc = startRowUrlDesc + batchSize - 1;
    if (nextendRowUrlDesc > endRowUrlDesc) nextendRowUrlDesc = endRowUrlDesc;
  
    // Process the next batch
    processUrlDescBatches(startRowUrlDesc, nextendRowUrlDesc, inColUrlDesc, outColUrlDesc);
  
    // Update the last processed row
    properties.setProperty('lastRowProcessedUrlDesc', nextendRowUrlDesc.toString());
  
    if (nextendRowUrlDesc < endRowUrlDesc) {
      deleteExistingTriggers('continueUrlDescProcessing');
      ScriptApp.newTrigger("continueUrlDescProcessing")
        .timeBased()
        .after(10000) // Wait 10 seconds before continuing again
        .create();
    } else {
      // All batches are processed, delete the trigger
      deleteExistingTriggers('continueUrlDescProcessing');
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
  