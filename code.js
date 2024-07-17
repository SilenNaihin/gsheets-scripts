function BatchRun() {
  /** 
   Initializes the batch processing.
   Options: 'BatchGPT', 'BatchUrlDesc', 'BatchLinkedinProfile'
   * @param {number} startRow The starting row number for processing.
   * @param {number} endRow The ending row number for processing.
   * @param {string} inCol The column letter where the prompts are.
   * (BatchLinkedinProfile only) @param {string} inCol2 The column letter where the prompts are.
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
 * Custom Google Sheets function to summarize results from a person's LinkedIn profile.
 * @param {string} profileUrl The LinkedIn profile URL of the person.
 * @return A single string summary based on results.
 * @customfunction
 */
function LinkedinProfile(
  profileUrl = 'https://www.linkedin.com/in/silen-naihin/'
) {
  var apiKey = '';
  var apiEndpoint = 'https://api.leadmagic.io/profile-search';
  var payload = {
    profile_url: profileUrl,
  };

  var options = {
    method: 'post',
    contentType: 'application/json',
    headers: {
      accept: 'application/json',
      'content-type': 'application/json',
      'X-API-Key': apiKey, // Adjusted based on typical API key formats
    },
    // muteHttpExceptions: true,
    payload: JSON.stringify(payload),
  };

  try {
    var response = UrlFetchApp.fetch(apiEndpoint, options);
    var data = JSON.parse(response.getContentText());

    // Building a summary from the profile data
    var summary = [];
    if (data.headline) summary.push('Headline: ' + data.headline);
    if (data.location) summary.push('Location: ' + data.location);
    if (data.company_industry)
      summary.push('Industry: ' + data.company_industry);
    if (data.company_website) summary.push('Website: ' + data.company_website);
    if (data.about) summary.push("About (IMPORTANT): '" + data.about + "'\n");
    if (data.experiences && data.experiences.length > 0) {
      var recentExperiences = data.experiences
        .slice(0, 2)
        .map((exp, index) => {
          if (
            exp.subComponents &&
            exp.subComponents.length > 0 &&
            exp.subComponents[0].hasOwnProperty('title')
          ) {
            // Handle complex job histories with multiple roles
            var rolesDetails = exp.subComponents
              .map((role, roleIndex) => {
                var roleDetails = `Role ${roleIndex + 1}: ${role.title}`;
                if (role.subtitle) {
                  roleDetails += `, ${role.subtitle}`;
                }
                if (role.caption) {
                  roleDetails += ` (${role.caption})`;
                }
                if (role.description && role.description.length > 0) {
                  var descriptions = role.description
                    .map((desc) => desc.text)
                    .join(' ');
                  roleDetails += `, Description: ${descriptions}`;
                }
                return roleDetails;
              })
              .join('.\n');
            return `Work experience ${index + 1} at ${
              exp.title
            } \n ${rolesDetails}`;
          } else if (exp.subComponents && exp.subComponents.length > 0) {
            // Handle simpler job entries or single descriptions without role details
            var experienceDetails = `Work experience ${index + 1}: ${
              exp.title
            }`;
            if (exp.subtitle) {
              experienceDetails += `, ${exp.subtitle}`;
            }
            if (exp.caption) {
              experienceDetails += ` (${exp.caption})`;
            }
            if (
              exp.subComponents[0].description &&
              exp.subComponents[0].description.length > 0
            ) {
              var descriptionText = exp.subComponents[0].description
                .map((desc) => desc.text)
                .join(' ');
              experienceDetails += `, Description: ${descriptionText}`;
            }
            return experienceDetails;
          }
        })
        .join(';\n\n');
      summary.push('Experience: ' + recentExperiences);
    }

    summary = summary.join('\n');

    console.log(summary);

    return summary;
  } catch (error) {
    const errorMsg = 'Failed to fetch profile: ' + error.toString();
    console.log(errorMsg);
    return '';
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
    case 'BatchLinkedinProfile':
      startLinkedinProfileProcessing(startRow, endRow, inCol, inCol2, outCol);
      break;
    default:
      throw new Error('Invalid batch function selected');
  }
}
