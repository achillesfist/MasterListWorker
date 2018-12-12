from collections import OrderedDict

import pandas as pd

file = 'MasterList.xlsx'
outputFile = 'output.xlsx'


xl = pd.ExcelFile(file)
dataframe = xl.parse('Report')
outDF = pd.DataFrame(columns=dataframe.keys())

emailRow = OrderedDict()
crmRow = OrderedDict()

"""
If the email's match, copy the email status over.
Otherwise, replace the main email with the email in Email Validation row, copy email status. 
Add original email to Email 2 or 3.
Mark the row's appropriately. Valid for any emailstatus other than invalid.  
"""


def mergeRows(mainRow, validationRow):
    newRow = mainRow
    if mainRow['Email'] == validationRow['Email']:
        newRow['EmailStatus'] = validationRow['EmailStatus']

    else:
        # Move email to email2 or email3 assuming they are empty
        if newRow['Email2'] == "":
            newRow['Email2'] = mainRow['Email']
        elif newRow['Email3'] == "":
            newRow['Email3'] = mainRow['Email']
        else:
            print("No space for original email")

        # Move email validation email to main email, and copy over email status
        newRow['Email'] = validationRow['Email']
        newRow['EmailStatus'] = validationRow['EmailStatus']

    # Mark the row's status appropriately
    validationRow['STATUS'] = "Delete"
    if validationRow['EmailStatus'] == "Invalid":
        newRow['STATUS'] = "Invalid"
    else:
        newRow['STATUS'] = "Valid"

    return newRow


for row in dataframe.itertuples():
    if row.Source == "Ed\'s CRM":
        if emailRow == {}:
            crmRow = row._asdict()
        elif emailRow['FullName'] == row.FullName:
            print("Email Row Matched on " + row.FullName)
            outDF = outDF.append(mergeRows(row._asdict(), emailRow), ignore_index=True)
            print("Row's merged, reverting Email and CRM row")
            emailRow = OrderedDict()
            crmRow = OrderedDict()
        else:
            print("No Match, reverting Email Row")
            emailRow = OrderedDict()
            crmRow = row._asdict()
    elif row.Source == "Email Validation":
        if crmRow == {}:
            emailRow = row._asdict()
        elif crmRow['FullName'] == row.FullName:
            print("CRM Row matched on " + row.FullName)
            outDF = outDF.append(mergeRows(crmRow, row._asdict()), ignore_index=True)
            print("Row's merged, reverting Email and CRM row")
            emailRow = OrderedDict()
            crmRow = OrderedDict()
        else:
            print("No Match, reverting CRM Row")
            crmRow = OrderedDict()
            emailRow = row._asdict()
    else:
        print("Row ignored for now")


print(outDF)

"""
writes the output Data Frame to a new file under the sheet called "Merged"
"""
writer = pd.ExcelWriter(outputFile)
outDF.to_excel(writer, 'Merged')
writer.save()