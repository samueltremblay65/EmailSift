from openpyxl import load_workbook

def open_workbook(path):
    workbook = load_workbook(filename=path)
    return workbook

def delete_from_excel_sheet(excel_file, emails):

    # Obtain data from excel sheet
    workbook = open_workbook(excel_file)

    for worksheet in workbook.sheetnames:
        sheet = workbook[worksheet]
        deletions = process_sheet(emails, sheet)

        print(deletions)

def process_sheet(emails, sheet):
    deletion_candidates = []

    # Get from list
    from_list = []
    i = 2
    while sheet.cell(i, 1).value is not None:
        from_list.append(sheet.cell(i, 1).value)
        i += 1

    # Get delete list
    delete_list = []
    i = 2
    while sheet.cell(i, 2).value is not None:
        delete_list.append(sheet.cell(i, 2).value)
        i += 1

    # Get keep list
    keep_list = []
    i = 2
    while sheet.cell(i, 3).value is not None:
        keep_list.append(sheet.cell(i, 3).value)
        i += 1

    # Get keep senders
    keep_senders = []
    i = 2
    while sheet.cell(i, 4).value is not None:
        keep_senders.append(sheet.cell(i, 4).value)
        i += 1

    print(from_list)
    print(delete_list)
    print(keep_list)
    print(keep_senders)

    for email in emails:
        for word in delete_list:
            if not (email.html is None) and word in email.html:
                deletion_candidates.append(email)
            elif not (email.plain is None) and word in email.plain:
                deletion_candidates.append(email)
            elif not (email.subject is None) and word in email.subject:
                deletion_candidates.append(email)

    deletions = apply_keep_logic(deletion_candidates, keep_list, keep_senders)

    return deletions

def apply_keep_logic(candidates, keep_list, keep_senders):
    deletions = [email for email in candidates if (not contains_keep_word(email, keep_list) and not from_keep_sender(email, keep_senders))]
    return deletions

def contains_keep_word(email, keep_list):
    for word in keep_list:
        if not (email.html is None) and word in email.html:
            return True
        elif not (email.plain is None) and word in email.plain:
            return True
        elif not (email.subject is None) and word in email.subject:
            return True
    return False

def from_keep_sender(email, keep_senders):
    for sender in keep_senders:
        if email.sender == sender:
            return True
    return False
