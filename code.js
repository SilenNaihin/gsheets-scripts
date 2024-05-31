function BatchRun() {
  /** 
   Initializes the batch processing.
   Options: 'BatchGPT', 'BatchUrlDesc', 'BatchGoogleName'
   * @param {number} startRow The starting row number for processing.
   * @param {number} endRow The ending row number for processing.
   * @param {string} inCol The column letter where the prompts are.
   * (BatchGoogleName only) @param {string} inCol2 The column letter where the prompts are.
   * @param {string} outCol The column letter where the results should be written.
  */
  startProcessing('BatchGPT', 7, 12, 'Q', 'R');
}

// Common function to handle API calls with exponential backoff
function fetchWithExponentialBackoff(url, params) {
  const maxAttempts = 5;
  let attempt = 0;
  let sleep = 1000; // Start with 1 second

  while (attempt < maxAttempts) {
    try {
      const response = UrlFetchApp.fetch(url, params);
      return JSON.parse(response.getContentText());
    } catch (e) {
      if (attempt < maxAttempts - 1) {
        Utilities.sleep(sleep);
        sleep *= 2; // Double the sleep time for the next attempt
      } else {
        console.error('Final attempt failed:', e);
        throw new Error('Error after multiple attempts: ' + e.toString());
      }
    }
    attempt++;
  }
}

const OPENAI_API_KEY = 'sk-';

/**
 * Custom Google Sheets function to call OpenAI's ChatGPT (gpt-4o).
 * @param {string} prompt The prompt to send to OpenAI.
 * @return The response from OpenAI.
 * @customfunction
 */
function GPT(prompt, model = 'gpt-4o') {
  if (prompt === 'Skip' || !prompt) {
    return '';
  }

  const url = 'https://api.openai.com/v1/chat/completions';

  const params = {
    method: 'post',
    contentType: 'application/json',
    headers: {
      Authorization: `Bearer ${OPENAI_API_KEY}`,
    },
    muteHttpExceptions: true,
    payload: JSON.stringify({
      model: model,
      messages: [{ role: 'system', content: prompt }],
    }),
  };

  const json = fetchWithExponentialBackoff(url, params);
  return json.choices[0].message.content.trim();
}

const options = {
  method: 'get',
  muteHttpExceptions: true,
  followRedirects: true,
  headers: {
    'User-Agent':
      'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.3',
  },
};

function fetchAndFilterHtml(url) {
  // Fetch the webpage content
  var response = UrlFetchApp.fetch(url, options);
  var htmlContent = response.getContentText();

  // Remove <script> and <style> tags and their content
  htmlContent = htmlContent.replace(
    /<script\b[^<]*(?:(?!<\/script>)<[^<]*)*<\/script>/gi,
    ''
  );
  htmlContent = htmlContent.replace(
    /<style\b[^<]*(?:(?!<\/style>)<[^<]*)*<\/style>/gi,
    ''
  );

  // Remove specific attributes globally
  htmlContent = htmlContent.replace(
    /\s*(src|href|target|id|class|style|value|type|role|tabindex|data-ved|data-dropdown)="[^"]*"/gi,
    ''
  );

  // Remove self-closing tags that are typically used empty
  htmlContent = htmlContent.replace(/<\s*(input|br|hr|img)\s*\/?>/gi, '');

  // Remove all tags
  htmlContent = htmlContent.replace(/<[^>]+>/g, '');

  // Remove empty container tags
  // htmlContent = htmlContent.replace(/<\s*(div|p|li|a|span|ul|ol|section|article)\s*>\s*<\/\s*\1\s*>/gi, "");

  // Remove all doubled linebreaks
  htmlContent = htmlContent
    .replace(/\r?\n|\r/g, '\n')
    .replace(/\n+/g, '\n')
    .trim();

  // Isolate <head> content
  var headContentMatch = htmlContent.match(/<head[^>]*>([\s\S]*?)<\/head>/i);
  var headContent = headContentMatch ? headContentMatch[1] : '';

  // Filter <head> to keep title and meta description
  var titleMatch = headContent.match(/<title>[\s\S]*?<\/title>/i);
  var metaDescriptionMatch = headContent.match(
    /<meta\s+name="description"\s+content="[^"]*"/i
  );
  // Rebuild head content with only the title and meta description
  var filteredHeadContent =
    (titleMatch ? titleMatch[0] : '') +
    (metaDescriptionMatch ? metaDescriptionMatch[0] : '');

  // Replace original <head> with filtered content
  htmlContent = htmlContent.replace(
    /<head[^>]*>[\s\S]*?<\/head>/i,
    `<head>${filteredHeadContent}</head>`
  );

  // Remove class and style attributes, if not already done
  htmlContent = htmlContent.replace(/\s*class="[^"]*"/gi, '');
  htmlContent = htmlContent.replace(/\s*style="[^"]*"/gi, '');

  // Log or return the cleaned HTML
  Logger.log(htmlContent);
  return htmlContent;
}

/**
 * Custom Google Sheets function to get the SEO description of a website.
 * @param {string} url The url of the website.
 * @return The SEO description.
 * @customfunction
 */
function WebSEO(url) {
  let htmlContent;
  try {
    // Fetch the webpage content
    const response = UrlFetchApp.fetch(url, options);
    htmlContent = response.getContentText();
  } catch (e) {
    return '';
  }

  // Regular expression to find the meta description content
  const metaDescriptionMatch = htmlContent.match(
    /<meta\s+name="description"\s+content="([^"]*)"/i
  );

  // Check if the meta description tag was found and extract the content
  const seoDescription = metaDescriptionMatch ? metaDescriptionMatch[1] : '';

  // Log or return the SEO description
  Logger.log(seoDescription);
  return seoDescription;
}

/**
 * Custom Google Sheets function to explain about a url.
 * @param {string} url The url of the website.
 * @return A short summary of the website.
 * @customfunction
 */
function UrlDescription(url) {
  let htmlContent;
  try {
    htmlContent = fetchAndFilterHtml(url);
    // Truncate htmlContent to 60,000 characters if it's longer.
    if (htmlContent.length > 60000) {
      htmlContent = htmlContent.substring(0, 60000);
    }
  } catch (e) {
    return '';
  }

  const response = GPT(
    `Please take a look at this html content (with all of the tags removed) from the ${url} company website and write a short 60 word summary of what the company does ${htmlContent}. Include any facts that you already know that may not be included on the website. I am looking for things I can reference in an email, so any recent announcements or anything specific to their company is valuable (taglines, sayings, priorities). Do not mention anything about how they can be contacted.`
  );

  return response;
}

/**
 * Custom Google Sheets function to explain the search results for a name.
 * @param {string} name The name of the person.
 * @param {string} company The company the person works at. Defaults to 'their company'.
 * @return A single string summary of the person's name description based on search results.
 * @customfunction
 */
function GoogleName(name, company) {
  var apiKey = '';
  var searchEngineId = '';

  var url =
    'https://www.googleapis.com/customsearch/v1?key=' +
    apiKey +
    '&cx=' +
    searchEngineId +
    '&q=' +
    encodeURIComponent(name + ' ' + company);

  try {
    var response = UrlFetchApp.fetch(url); // Make the API request
    var json = JSON.parse(response.getContentText());
    console.log(json);
    var searchResults = [];

    // Parse the search results and format them into a list of strings
    if (json.items) {
      json.items.forEach(function (item, index) {
        var title = item.title;
        var snippet = item.snippet.replace(/\n/g, ' '); // Remove newline characters from snippet
        searchResults.push(index + 1 + '. ' + title + ' - ' + snippet);
      });
    }

    console.log(searchResults);

    // Combine the search results into a single string, separated by new lines
    var resultsString = searchResults.join('\n');

    console.log(resultsString);

    const responseSummary =
      GPT(`Please take a look at this content from a Google search "${name} ${company}", and give a short summary about this person, and a couple
    of specific facts about them. Please focus on PERSONAL FACTS ABOUT THEMSELVES, and ignore anything related to the company - this is for a cold email so try to extract
    information I can reference within it. 
Search results:
${resultsString}

DO NOT mention the company (x company the work at...). DO NOT mention how to contact them (you can reach them at...) because I already know all of those.
Be sure to look at the dates that the search results are associated with. Do not say 'recent' if a promotion happened in 2021 for example.
Try to mention things like "this person wrote an article about xyz" or "this person enjoys to do x on their free time" or "this person joined ${company} because xyz" or other things like that which I can reference in a cold outreach email. Extract as much value as you can from the search results. 
AVOID SAYING ANYTHING LIKE 'unfortunately there are no personal facts available about...' - just return the useful content in the response.
DO NOT mention how many connections they have on Linkedin.
Respond with nothing but a short single paragraph of 60 words about this person. `);
    return responseSummary;
  } catch (e) {
    // Handle errors (e.g., API rate limit exceeded, network issues)
    console.error('Error fetching search results: ' + e.toString());
    return 'Error fetching search results. Please try again later.';
  }
}

function startProcessing(
  batchFunction,
  startRow,
  endRow,
  inCol,
  outCol,
  inCol2
) {
  switch (batchFunction) {
    case 'BatchGPT':
      startGPTProcessing(startRow, endRow, inCol, outCol);
      break;
    case 'BatchUrlDesc':
      startUrlDescProcessing(startRow, endRow, inCol, outCol);
      break;
    case 'BatchGoogleName':
      startGoogleNameProcessing(startRow, endRow, inCol, inCol2, outCol);
      break;
    default:
      throw new Error('Invalid batch function selected');
  }
}
