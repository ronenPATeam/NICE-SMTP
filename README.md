# NICE-SMTP
NICE dll that adds the ability to send emails to SMTP directly, with or without attachments, plain text or html

Main features:

- Type 'Email Message' 
-- Recipients (list of)
-- CC Recipients (list of)
-- Attachments (list of)
-- Subject
-- Body
-- Sender
-- Server Details
-- Username
-- Password
-- Port
-- ESSL
-- isHTML
-- use Default Credentials
-- Send Timeout

- Send email using SMTP (using the 'Email Message' Type

Requirements:
- Port to SMTP server needs to be open

Install:
- Copy Direct.SMTP.Library.dll to NICE Designer and NICE Client installation directory
- Reference Dll in project
For reference see: 
http://apa-onlinehelp.s3-website-us-west-2.amazonaws.com/72/content/library%20objects%20sdk/installing%20the%20library%20objects%20project.htm

Verified Compatibility:

- NICE RTS 6.6
- NICE APA 6.7
- NICE APA 7.0
- NICE APA 7.1
- NICE APA 7.2


Disclaimer: this is a product of PAteam meant for the NICE community and is not created or supported by NICE