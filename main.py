import smtplib
from simplegmail import Gmail
from deletion_logic import delete_from_excel_sheet

gmail = Gmail()

gmail.maxResults = 50

# Unread messages in your inbox
emails = gmail.get_messages()

deleted = delete_from_excel_sheet("SifterSheetExample.xlsx", emails)

print("Your inbox has been sifted. {0} emails have been deleted".format(deleted))