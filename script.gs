const OPENAI_API_KEY = "sk-"

/**
 * Custom Google Sheets function to call OpenAI's ChatGPT (gpt-3.5-turbo).
 * @param {string} prompt The prompt to send to OpenAI.
 * @return The response from OpenAI.
 * @customfunction
 */
function GPT(prompt, model = "gpt-3.5-turbo") {
  const url = "https://api.openai.com/v1/chat/completions";

  const params = {
    method: "post",
    contentType: "application/json",
    headers: {
      "Authorization": `Bearer ${OPENAI_API_KEY}`
    },
    muteHttpExceptions: true,
    payload: JSON.stringify({
      model: model,
      messages: [
        { role: "system", content: prompt }
      ],
    })
  };

  try {
    const response = UrlFetchApp.fetch(url, params);
    const json = JSON.parse(response.getContentText());
    return json.choices[0].message.content.trim();
  } catch (e) {
    return "Error calling the OpenAI API: " + e.toString();
  }
}

const options = {
  "method": "get",
  "muteHttpExceptions": true,
  "followRedirects": true,
  "headers": {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.3"
  }
};


function fetchAndFilterHtml(url) {
  // Fetch the webpage content
  var response = UrlFetchApp.fetch(url, options);
  var htmlContent = response.getContentText();
  
  // Remove <script> and <style> tags and their content
  htmlContent = htmlContent.replace(/<script\b[^<]*(?:(?!<\/script>)<[^<]*)*<\/script>/gi, "");
  htmlContent = htmlContent.replace(/<style\b[^<]*(?:(?!<\/style>)<[^<]*)*<\/style>/gi, "");

  // Remove specific attributes globally
  htmlContent = htmlContent.replace(/\s*(src|href|target|id|class|style|value|type|role|tabindex|data-ved|data-dropdown)="[^"]*"/gi, "");

  // Remove self-closing tags that are typically used empty
  htmlContent = htmlContent.replace(/<\s*(input|br|hr|img)\s*\/?>/gi, "");

  // Remove all tags
  htmlContent = htmlContent.replace(/<[^>]+>/g, "");

  // Remove empty container tags
  // htmlContent = htmlContent.replace(/<\s*(div|p|li|a|span|ul|ol|section|article)\s*>\s*<\/\s*\1\s*>/gi, "");

  // Remove all doubled linebreaks
  htmlContent = htmlContent.replace(/\r?\n|\r/g, "\n").replace(/\n+/g, "\n").trim();

  // Isolate <head> content
  var headContentMatch = htmlContent.match(/<head[^>]*>([\s\S]*?)<\/head>/i);
  var headContent = headContentMatch ? headContentMatch[1] : "";
  
  // Filter <head> to keep title and meta description
  var titleMatch = headContent.match(/<title>[\s\S]*?<\/title>/i);
  var metaDescriptionMatch = headContent.match(/<meta\s+name="description"\s+content="[^"]*"/i);
  // Rebuild head content with only the title and meta description
  var filteredHeadContent = (titleMatch ? titleMatch[0] : "") + (metaDescriptionMatch ? metaDescriptionMatch[0] : "");
  
  // Replace original <head> with filtered content
  htmlContent = htmlContent.replace(/<head[^>]*>[\s\S]*?<\/head>/i, `<head>${filteredHeadContent}</head>`);
  
  // Remove class and style attributes, if not already done
  htmlContent = htmlContent.replace(/\s*class="[^"]*"/gi, "");
  htmlContent = htmlContent.replace(/\s*style="[^"]*"/gi, "");

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
    return "";
  }
  
  // Regular expression to find the meta description content
  const metaDescriptionMatch = htmlContent.match(/<meta\s+name="description"\s+content="([^"]*)"/i);
  
  // Check if the meta description tag was found and extract the content
  const seoDescription = metaDescriptionMatch ? metaDescriptionMatch[1] : "";
  
  // Log or return the SEO description
  Logger.log(seoDescription);
  return seoDescription;
}

/**
 * Custom Google Sheets function to explain the search results for an email with exponential backoff.
 * @param {string} email The email of the person.
 * @param {string} company The company the person works at.
 * @return Email description of the person.
 * @customfunction
 */
function GoogleEmail(email, company='their company') {
  const baseDelay = 1000; // Start with a 1 second delay
  let attempt = 0; // Number of attempts made
  
  while (attempt < 5) { // Try up to 5 times
    const searchUrl = "https://www.google.com/search?q=" + encodeURIComponent(email);
    const htmlContent = fetchAndFilterHtml(searchUrl); // You need to implement this function
    
    // Check if the htmlContent contains the rate limit warning
    if (htmlContent.includes("Our systems have detected unusual traffic from your computer network.") &&
        htmlContent.includes("This page appears when Google automatically")) {
      
      // Calculate the delay with exponential backoff
      const delay = baseDelay * Math.pow(2, attempt);
      Utilities.sleep(delay); // Wait before retrying
      attempt++; // Increment attempt counter
      console.log("Retrying, attempt:", attempt, " ", baseDelay/1000, " seconds left")
    } else {
      // If there's no rate limit warning, process the content
      // Assuming processHtmlContent is a function you've defined to parse and process the HTML content
      const response = GPT(`Please take a look at this html content from a Google search with the persons email, and give a short summary about this person, and a couple of specific facts about them ${htmlContent}. If there is nothing specific about the person and it is just about ${company}, just say 'company' and nothing else. 
It should be in the following format: 
"60 words MAX description
- fact 1
- fact 2"
Keep it short and concise, no fluff at all please. I'm looking less for ideally things that happened recently, or things that I can reference in an email without feeling weird about it. For the two facts that you provide be specific about things regarding that person. USE THE SPECIFIC VERBAGE THAT THE PERSON WROTE DOWN for any posts. 
Do not mention the company (x company the work at...) or mention how to contact them (you can reach them at...) because I already know that.
Please just respond with 'company' if it is only information about the company or there is no specific information available about the person associated with the email address. Don't say anything about ${company}. However, if a person posted something about ${company}, that's ok.`)
      return response;
    }
  }
  
  // If the function reaches this point, it means all attempts have failed
  return "FAILED";
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
    htmlContent = fetchAndFilterHtml(url)
    // Truncate htmlContent to 60,000 characters if it's longer.
    if (htmlContent.length > 60000) {
      htmlContent = htmlContent.substring(0, 60000);
    }
  } catch (e) {
    return "";
  }

  const response = GPT(`Please take a look at this html content (with all of the tags removed) from the ${url} company website and write a short 60 word summary of what the company does ${htmlContent}. Include any facts that you already know that may not be included on the website. I am looking for things I can reference in an email, so any recent announcements or anything specific to their company is valuable (taglines, sayings, priorities). Do not mention anything about how they can be contacted.`)

  return response
}

const response = WebSEO("https://www.endflowautomation.com/")

console.log(response)




