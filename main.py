import smtplib
from simplegmail import Gmail
from deletion_logic import delete_from_excel_sheet

gmail = Gmail()

gmail.maxResults = 50

deleted = delete_from_excel_sheet("SifterSheetExample.xlsx", gmail)

print("Your inbox has been sifted. {0} emails have been deleted".format(deleted))