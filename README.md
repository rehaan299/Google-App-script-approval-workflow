This repository exists because of a team meeting I had last week.

We realized our process for approving training requests and expenses was a disaster. It was a mix of email threads, lost attachments, and people asking "Did you see my request?" in chat.

I looked at the existing solutions on the Google Workspace Marketplace. There are plenty of "Form Approval" add-ons, but they all require monthly subscriptions or "per-user" fees. I didn't want to pay for a SaaS subscription just to route a few emails, and I definitely didn't want to wait for IT to approve a new vendor.

So I built this instead. It’s a custom engine running on Google Apps Script (which is free) that does exactly what the paid tools do, but with more flexibility.

What it actually does
This script turns a Google Form into a multi-stage approval app.
It reads the room: When someone submits a form, the script checks the "Department" and "Team" they selected.
It knows the chain: Based on that selection, it pulls a specific list of approvers (e.g., Manager -> VP).
It sends emails: It emails the first person. They can click "Approve" or "Reject" right inside the email.
It has a UI: Because digging through emails is annoying, I built a web dashboard where approvers can see everything in one place.
Why this is better than a spreadsheet formula
Most people try to do this with spreadsheet formulas or Zapier. The problem is "State Management." If you have a 3-step approval process, you usually need 6 different columns in your sheet to track who approved what and when. It gets messy fast.

The Fix: This script stores the entire approval history in a single cell using JSON. It keeps your spreadsheet clean and makes the app much faster.

The Files
Code.gs: The backend logic. It handles the routing and talks to Gmail.
dashboard.html: The main view for managers to approve requests.
approval_email.html: The template for the email notifications.
index.html: The view for a single task.
portal.html: A simple login/landing page.
How to use this (The Setup)
You can't just clone this repo and run it. You need to hook it up to a Google Sheet.

1. The Form & Sheet
Create a Google Form with the questions you need (e.g., Cost, Description, Team). Go to the "Responses" tab in the Form and click the green Sheets icon to create a linked Spreadsheet.

2. The Code
Open that new Spreadsheet. Go to Extensions > Apps Script. Copy the files from this repo and paste them into that script editor. You will need to create the HTML files manually in the editor and paste the code in.

3. The Configuration
Look at the top of Code.gs. You will see a variable called FLOWS. This is where the magic happens. You need to map your Form dropdown answers to real email addresses.

It looks like this:

JavaScript
"Engineering|Backend": [
    { email: "manager@company.com", name: "Alice", title: "Manager" },
    { email: "director@company.com", name: "Bob", title: "Director" }
]
Make sure the keys (like "Engineering|Backend") match your form answers exactly.

4. Deployment (Important)
For the email links and the dashboard to work, this needs to be published as a Web App.
Click Deploy -> New Deployment.
Select Type: Web App.
Execute as: Me.
Who has access: Anyone within my organization (or Anyone with a Google Account).
Copy the URL it gives you. Paste that URL into the this.url variable in Code.gs.
5. Turn it on
Run the function called createTrigger once from the script editor. This tells Google to run the script every time a form is submitted.

A Note on "Silent Mode"
I added a feature specifically for my Director. They get too many emails and tend to ignore them. If you title someone "Director" or "VP" in the config, the script intentionally skips sending them an instant email for every single request. Instead, there is a function called sendWeeklyDigest that you can schedule (using Triggers) to send them one summary email on Monday mornings.

Contributing
If you find a bug or want to make the UI look even better, feel free to fork this. Just don't sell it as a paid add-on, we have enough of those.
