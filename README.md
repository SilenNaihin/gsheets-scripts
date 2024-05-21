These are scripts I use in my Google sheet in order to enrich lead information and craft custom emails using GPT.

### Usage

- Open a Google Sheets
- Open Extensions
- Go to Apps Script
- Paste in the code above
- Add your OpenAI key
- Save
- Run and authorize for access to external scripts (needed to run Triggers and use Google libraries)

### NOTE: Every time you reorder rows, delete a row, or refresh/reopen the spreadsheet the cells will rerun if it is a cell you ran a formula in. This can be costly. Ensure that you crtl+c then crtl+shit+v to just paste the content of the cells without the formulas once they are done.

### There are 4 added functions

`=WebSEO(D2)`

- SEO description of the website.

`=UrlDescription(D2)`

- Description of the website contents.
- Uses GPT, requires OpenAI key.

`=GoogleName(C2,B2)`

- 'Full Name + Company Name' is the Google search.
- Uses GPT.
- Requires Google `apiKey` and `searchEngineId`.

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


``

In order to work around the limit, 
`Dialog.html` and `Menu.js` are not working in their current form, so the command parameters needs to be updated in the 