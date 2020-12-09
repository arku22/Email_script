import yagmail
from string import Template  # library to use template strings for easy subs
import openpyxl  # library to parse xlsx files
import sys


def read_template(filepath):
    """This function reads a txt file containing the template and
        returns the template.
    """
    with open(filepath, 'r', encoding='utf-8') as f:
        content = f.read()

    return Template(content)


def get_contact_dict(filepath, start_row, end_row, sheet_num):
    """This function parses the xlsx file to return a dictionary with
        the fields required to send out an email to the recruiter/contact.
        Following is the structure of the returned dictionary :-
        {company1: {pairs:[(name1,email1), (name2,email2),...],
                    pos_name: __,
                    pos_loc: __,}
        .
        .
        .
        companyn: {___}
        }
        """
    wb = openpyxl.load_workbook(filepath)

    # select relevant worksheet through index, default = 0
    wb.active = sheet_num
    sheet = wb.active

    # Initial Program Checks
    prompt = 'Did you check the following :-\n\u2022 start_row' \
        '\n\u2022 end_row' \
        '\n\u2022 sheet_num' \
        '\n\u2022 filepaths (y/n) ? = '
    check = input(prompt)
    text = "***Correct the required values and run script again!***"
    text += "\nSee lines 78-98 in script"
    if check != 'y':
        sys.exit(text)

    data = {}
    for r in range(start_row, end_row+1):
        print('Sending Emails to company ' + str(r) + '...')
        company = sheet['A' + str(r)].value
        pos_name = sheet['B' + str(r)].value
        pos_loc = sheet['C' + str(r)].value

        rec_names = sheet['F' + str(r)].value
        try:
            rec_names = rec_names.strip().split(',')
            rec_emails = sheet['G' + str(r)].value

            rec_emails = rec_emails.strip().split(',')
            temp = list(zip(rec_names, rec_emails))

            temp_dict = {}
            temp_dict['pos_name'] = pos_name
            temp_dict['pos_loc'] = pos_loc
            temp_dict['pairs'] = temp

            data[company] = temp_dict

        except AttributeError:
            continue

    return data


# ------------------------------------------------------------------------------
# select what rows to consider to send email
# start_row = 1
# end_row = 2
# sheet_num = 0
# Path to email body template
# filepath_template =
# filepath_template +=
# Path to xlsx spreadsheet
# filepath_contacts =
# filepath_contacts +=
# Path to Resume
# filepath_resume =
# Set your email address and password
# email_id =
# password =
# set smtp details
# host_name = 'smtp.office365.com'
# port_number = 587
# ------------------------------------------------------------------------------

# Function Calls
# Read in email body template
template = read_template(filepath_template)
# Get the dictionary with the info for each company
my_dict = get_contact_dict(filepath_contacts, start_row, end_row, sheet_num)

for key, value in my_dict.items():

    pos_name = value['pos_name']
    pos_loc = value['pos_loc']
    for pair in value['pairs']:
        # NOTE : Here 'pair' is a tuple
        rec_first_name = pair[0].split()[0].title()
        rec_email = pair[1]
        subs_mapping = {
            'COMPANY_NAME': key,
            'POSITION_NAME': pos_name,
            'POSITION_LOCATION': pos_loc,
            'REC_FIRST_NAME': rec_first_name}
        message = template.substitute(subs_mapping)

        yag = yagmail.SMTP(user='yourgmail@address',
                           password='emailpassword')
        # sending the email
        yag.send(to=rec_email,
                 subject='Subject',
                 contents=message,
                 # attachments='')

print(".")
print(".")
print(".")
print("--------Emails sent successfully!--------")
