// if you prefer AZURE

// Replace YOUR_API_KEY with your actual Azure OpenAI API key
const AZURE_API_KEY = '';
// Azure OpenAI endpoint for your specific deployment
const AZURE_ENDPOINT = 'https://google-sheets.openai.azure.com/openai/deployments/';

/**
 * Custom Google Sheets function to call Azure's OpenAI endpoint.
 * @param {string} prompt The user's prompt to send to OpenAI.
 * @param {number} temperature of the completion.
 * @return The response from OpenAI.
 * @customfunction
 */
function AZURE(prompt, temperature=0.1) {
  const params = {
    method: "post",
    contentType: "application/json",
    headers: {
      "Content-Type": "application/json",
      "api-key": AZURE_API_KEY
    },
    muteHttpExceptions: true,
    payload: JSON.stringify({
      messages: [
        { role: "system", content: prompt }
      ],
      max_tokens: 1200,
      temperature: temperature,
      frequency_penalty: 0,
      presence_penalty: 0,
      top_p: 0.95,
      stop: null
    })
  };

  const json = fetchWithExponentialBackoff(AZURE_ENDPOINT, params);
  return json.choices[0].message.content.trim();
}

// Replace YOUR_API_KEY with your actual Azure OpenAI API key
const AZURE_35_API_KEY = '';
// Azure OpenAI endpoint for your specific deployment
const AZURE_35_ENDPOINT = 'https://google-sheets.openai.azure.com/openai/deployments';


/**
 * Custom Google Sheets function to call Azure's OpenAI endpoint with gpt-3.5.
 * @param {string} prompt The user's prompt to send to OpenAI.
 * @param {number} temperature of the completion.
 * @return The response from OpenAI.
 * @customfunction
 */
function AZURE_35(prompt, temperature=0.1) {
  const params = {
    method: "post",
    contentType: "application/json",
    headers: {
      "Content-Type": "application/json",
      "api-key": AZURE_35_API_KEY
    },
    muteHttpExceptions: true,
    payload: JSON.stringify({
      messages: [
        { role: "system", content: prompt }
      ],
      max_tokens: 1200,
      temperature: temperature,
      frequency_penalty: 0,
      presence_penalty: 0,
      top_p: 0.95,
      stop: null
    })
  };

  const json = fetchWithExponentialBackoff(AZURE_35_ENDPOINT, params);
  return json.choices[0].message.content.trim();
}