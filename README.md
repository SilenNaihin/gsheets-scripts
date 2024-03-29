These are scripts I use in my Google sheet in order to enrich lead information and craft custom emails using GPT.

### Usage
- Open a Google Sheets
- Open Extensions
- Go to Apps Script
- Paste in the code above
- Add your OpenAI key
- Save
- Run and authorize for access to external scripts (needed for grabbing websites

### NOTE: Every time you reorder rows, delete a row, or refresh/reopen the spreadsheet the cells will rerun. Ensure that you crtl+c then crtl+shit+v to just paste the content of the cells without the formulas once they are done.

Here's a link to an example spreadsheet you can copy and use yourself : https://docs.google.com/spreadsheets/d/1mGNvG-PXau_HDDUIhUJe0kngPsxQQTnRr7O5KLZcDp0

### There are 4 added functions

`=WebSEO(D2)`
Doesn't use GPT.

`=UrlDescription(D2)`
Uses GPT.

`=GoogleEmail(C2,D2)`
Uses GPT. Often gets rate limited by Google, doesn't work to drag and drop for hundreds of rows. I've added exponential backoff, but cells have a 30 second time limit and end up ERRORing. 

`=GPT("prompt", "optional-model")`

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
