# Email_script
Automating the process of applying to jobs.

While applying for jobs I found one approach was to send out emails directly to recruiters. However this was tedious compared to the more conventional approach of going to a job board and applying to positions. I had to first identify positions I wished to apply to via a job portal and then record that in an excel sheet, find the contacts/emails of relevant recruiters, send out emails to each recruiter while customizing the message/email body for each email. I decided to use Python to help me automate some of the repitative tasks. 

My script can read an excel file(.xlsx), parsing it to locate email addresses and name of the contact, login into your email service provider, compose an email, fill in the subject, attach files and write a customized message(using Python template strings) and send out multiple emails.

To run the program "email_script.py", you require the following libraries :-

* openpyxl - to parse xlsx spreadsheets
	`pip install openpyxl==2.6.2`
* smtplib
* email
* yagmail 0.14.245 (if sending email via Gmail)

This script only supports email with plain text content as of now.

Before running the script, be sure to check the values of the following variables : -

- start_row : The row number from where parsing should start in spreadsheet
- end_row : The row number where parsing should end in spreadsheet
- sheet_num : The sheet number within the workbook to be used by the script
- filepath_template : The path to the file containing the template for the email body
- filepath_contacts : The path to the workbook with the contacts
- filepath_resume : The path to the Resume to be attached with Email
- email_id : sender email address
- password : sender email login password
- host_name (SMTP details - go to your email's settings page and search for "POP and IMAP" or a similar keyword.)
- port_number (SMTP details - go to your email's settings page and search for "POP and IMAP" or a similar keyword.)

## REFERENCES
I used the following guide/web page as a starting point for my script : -
https://www.freecodecamp.org/news/send-emails-using-code-4fcea9df63f/

For Gmail : -
https://blog.mailtrap.io/yagmail-tutorial/
