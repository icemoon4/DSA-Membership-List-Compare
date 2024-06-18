# DSA-Membership-List-Compare
If your chapter uses Gmail and Gdrive, this script (run in Google Apps Script) will copy the membership list attachment from your email and move it to Drive, then compare it with last week's list to see what has changed
As a disclaimer, I'm not fond of JavaScript, but it's what Apps Script uses which was the best tool for the job. Sorry in advance for the chicken scratch code. Also, this was not rigorously tested :) It was informally run through a few scenarios by me and only me. So there's probably issues. I'm also from a small-ish chapter, so I don't know if the emailed membership lists come in any other format than the single zip file containing a single csv file. Have fun!

# How to use this script
You'll need to configure some things in your Gmail and Gdrive to utilize this script. I've outlined the structure that our chapter uses for this, but obviously you can make changes as needed to fit your own structure.

## Gmail
Create a filter for the membership lists that are emailed to your chapter from ActionKit. Apply some label to these emails. Make sure that no other emails will get filtered under this label.

## Google Drive
The folder structure here is up to you, but I will lay out how I have organized this in our chapter so you can tell what's happening in the code.
In our Drive, we have a folder titled Membership Lists. Note that the ID for the folder appears in the URL, after `/folders/`. You will need this ID in the script.
Within the Membership Lists folder, we have another folder titled Archive. This is where we move all old membership files after we finish processing them. The CSV files will be in the main folder, and then we have another folder under Archive titled Zips, where we move the zip file from the email.
Back in the main Membership Lists folder, we have a Google Sheet called Membership Updates. In it, there are three tabs: Additions, Modifications, Deletions.
Copy the header row from the membership list and paste it into these sheets. Move the last column to the front. This is the list date. I moved it to the front so that it would be easier to see when these changes occurred.

## Apps Script
Launch Apps Script from your chapter email account: https://script.google.com
You may need to enable it. Create a new script and copy in the script from this repository.
On the left hand side, you'll see triggers. Add a new trigger. We have a time-based trigger for every Friday at 2:30 PM EST. Membership lists usually come in every Friday around 10 or 11 AM EST. I give a couple hours in between in case of any potential lag.
Back in the script, here's how it works:

Automatically takes the attachment from the membership email, unzips it, and stores the latest CSV in the Membership Lists folder. The CSV file is renamed with the current date on the end.
The zip is moved into Archive > Zips
The script compares the rows in the week-old membership list with the fresh membership list, using the email field as the key.
Updates are moved to the Google Sheet, with the list date moved to the front so we can easily see when the data changes occur
If an email is found in the new list that does not exist in the old list, it will update the Additions tab with the row. This is a new member to the chapter
If an email is found in both lists, and any of the fields are different between the two rows (excluding xdate, because this field updates frequently for monthly paying members and we don’t really need to see this), then it will add the row to the Modifications tab and highlight whichever cell was updated.
If an email exists in the old list that is not found in the new list, it will add the row to the Deletions tab. This member was removed from the list
The week-old file is moved to Archive. The current file stays in the folder until next week when the script is run again.

The script needs to be updated with the IDs from your own Drive. See the TODO comments for where to make updates.

## Getting Started
After setting up your folder structure, manually copy your most recent membership list and move it into the Membership Lists folder. Append the filename with last Friday's date in JavaScript Date String format. Ex: worcester_membership_list_Fri Jun 14 2024.csv
There needs to be a file in here when the script runs next and it needs to include the date from 7 days ago. Assuming you want to run this on Friday's, it should be a Friday date.
