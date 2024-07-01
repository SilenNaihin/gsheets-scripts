/**
 * Function to get data from the spreadsheet.
 * @return {Array} The spreadsheet data.
 */
function getSpreadsheetData() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getValues();
  return data;
}

/**
 * Custom function to handle user queries and get the relevant column data.
 * @param {string} userQuery The user's natural language query.
 * @return The response from Azure OpenAI.
 * @customfunction
 */
function ASK(
  userQuery = 'is the total freight cost or insurance cost higher?'
) {
  try {
    const data = getSpreadsheetData();
    const headers = data[0];
    const rows = data.slice(1);

    const columnNames = headers.join(', ');
    const instruction = `From the following query, identify the corresponding column name from the dataset and if further analysis is required on the data once it is retrieved. The response format will be a json { "name": ... , "analysis": true }.
  
  You must choose a single column that is most relevant. For example if the query is asking about the 'cost' and there is some marker of generic or total cost, assume that column. Extrapolate this thinking to other scenarios. 
  
  For example if the query is "get the cost" and the headers are "Total Cost | Freight Cost | Insurance Cost", your response should be: { "name": "Total Cost" , "analysis": false }
  If the query is "get the average cost of freight" and the headers are "Total Cost | Freight Cost | Insurance Cost", your response should be { "name": "Freight Cost" , "analysis": true }
  
  Make sure that you respond with nothing but proper json that can be parsed through JSON.parse(). Do not respond with \`\`\`json or any backticks at all.
  
  Available columns are: ${columnNames}. 
  
  Query: ${userQuery}`;
    const response = AZURE(instruction);

    console.log(response);

    const parsedResponse = JSON.parse(response);
    const columnName = parsedResponse.name;
    const needsAnalysis = parsedResponse.analysis;

    console.log(columnName, needsAnalysis);

    let columnIndex = -1;
    for (let i = 0; i < headers.length; i++) {
      if (headers[i].toLowerCase() === columnName.toLowerCase()) {
        columnIndex = i;
        break;
      }
    }

    if (columnIndex === -1) {
      throw new Error('Column not found');
    }

    const columnData = rows.map((row) => row[columnIndex]);
    const columnLetter = String.fromCharCode(65 + columnIndex);
    const rowNumbers = rows.map((_, index) => index + 2); // Data starts from row 2
    let result =
      `Here is the data for the column ${columnName}:\n` +
      columnData.join('\n');

    console.log(result);

    if (needsAnalysis) {
      const analysisPrompt = `You need to provide a suitable Google Sheets function to analyze the data in column ${columnLetter} from row ${
        rowNumbers[0]
      } to row ${rowNumbers[rowNumbers.length - 1]} for the query: ${userQuery}.
        
  Please just return a spreadsheet formula. For example, if the query is asking for an average, just return =AVERAGE(O2:O41) with nothing else. This formula will be directly executed so make sure there are no trailing commas or backticks.`;
      const analysisResponse = AZURE(analysisPrompt);

      const formula = analysisResponse.trim();
      console.log('formula', formula);

      // Evaluate the formula using Google Sheets
      const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
      const resultCell = sheet.getRange('Z1'); // Temporarily use cell Z1 to evaluate the formula
      resultCell.setFormula(formula);
      formulaResult = resultCell.getValue();
      resultCell.clear(); // Clear the temporary cell

      const finalPrompt = `You are looking at data from column ${columnLetter} from row ${
        rowNumbers[0]
      } to row ${rowNumbers[rowNumbers.length - 1]} for the query: ${userQuery}.
  
  The last function that was run was ${formula}.
  
  Now respond to the query given that the result of executing that formula was ${formulaResult}.`;
      result = AZURE(finalPrompt);
    }

    console.log(result);
    return result;
  } catch (error) {
    return `Error: ${error.message}`;
  }
}
