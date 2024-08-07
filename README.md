These are scripts I use in my Google sheet in order to enrich lead information and craft custom emails using GPT.

### Usage

- Open a Google Sheets
- Open Extensions
- Go to Apps Script
- Paste in the code
- Add your OpenAI key and any other needed keys
- Save
- Run and authorize for access to external scripts (needed to run Triggers and use Google libraries)

### There are 4 added functions

`=WebSEO(D2)`

- SEO description of the website.

`=UrlDescription(D2)`

- Description of the website contents.
- Uses GPT, requires OpenAI key.

`=LinkedinProfile(B2)`

- Gets Linkedin info on a person in this format:

```
Headline: General Manager at Acme
Location: San Francisco, California, United States
Industry: machinery
Website: acme.com
About (IMPORTANT): 'I have been with Acme for...'
Experience: Work experience 1 at Acme, Inc.
Role 1: Director of Project Management and Field Service, Full-time (Nov 2019 - Present · 4 yrs 9 mos).
Role 2: Project Manager (May 2011 - Present · 13 yrs 3 mos)
Work experience 2: Senior Manager, JimBob Systems
```

How to

1. Leadmagic.io -> get an API key -> paste it into apiKey
2. Go to your GSheet -> extensions -> Apps Script -> paste the code with the API key -> click run button at the top once and allow permissions if it pops up -> `=LinkedinProfile(B2)` (B2 would have the profile in it)

`=GPT("prompt", "optional-model")`
I've added exponential backoff, but cells have a 30 second time limit so be careful with token limits.

Here's a super basic template, you'll see you need to be very specific with your instructions (ex. just write the email, no signature), provide a template, describe your company and offer, etc.

`=GPT(""Write me a customized email using this info: 
Name:"" & A2 & 
""
Title:"" & F2 &
""
Company name:"" & G2 &
""
Company SEO description:"" & O2 &
""
Company overview description:"" & P2 &
""
Person description:"" & Q2 &
""
The email should start with a greeting, don't be robotic, I am selling lead generation, offer is industry grade magnets. Be short and concise."", ""gpt-4"")`

#### NOTE: If you're running these formulas in cells raw, every time you reorder rows, delete a row, or refresh/reopen the spreadsheet the cells will rerun if it is a cell you ran a formula in. This can be costly. Ensure that you crtl+c then crtl+shit+v to just paste the content of the cells without the formulas once they are done.

Now lets say you're preparing 4000 emails with AI to send in Google Sheets. This will literally take days of time dragging the formula down, waiting for it to run, fixing any errors with cell timeouts or waiting for token limits to refresh. Trust me, I've spent days doing it. On the other side, Clay.com costs an insane amount of money, and doesn't even allow you to run code natively. You need to set up a Lambda function and call it through HTTP or a webhook.

Crazy thing is - Google Sheets allows you to execute long running code without needing wifi to run and bypassing all of the issues with the above. I've run these functions for 24+ hours automatically, perfectly. No cell limits. No timeouts. No token issues. No need to refresh. No costs outside of tokens.

`Dialog.html` and `Menu.js` are not working in their current form, so the command parameters need to be updated in the function itself, and the function needs to be run manually from the Apps Script (select the `BatchRun` function at the top of the .gs file and click run).

```javascript
function BatchRun() {
  /** 
   Initializes the batch processing.
   Options: 'BatchGPT' (in: string), 'BatchUrlDesc' (in: website), 'BatchGoogleName' (in: name, company name)
   * @param {number} startRow The starting row number for processing.
   * @param {number} endRow The ending row number for processing.
   * @param {string} inCol The column letter where the prompts are.
   * (BatchGoogleName only) @param {string} inCol2 The column letter where the prompts are.
   * @param {string} outCol The column letter where the results should be written.
  */
  startProcessing('BatchGPT', 7, 12, 'Q', 'R');
}
```

This exists for UrlDescription, GoogleName, and GPT functions. It can quite easily be set up for anything else that needs to be run for a while. More details:

- `batchSize` is manually set, the main constraint is that every batch has to complete in under 6 minutes. Don't forget that exponential backoff with the token limits can be a factor.
- The only reason I have time added between triggers is so that the token limits have time to refresh.
- You can parallel process these different long running functions if token limits allow.
