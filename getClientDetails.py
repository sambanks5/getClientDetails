import requests
import json
import time
import gspread
from oauth2client.service_account import ServiceAccountCredentials

## Load list of Client usernames
with open('clientss.txt') as f:
    clients = f.read().split('\n')

## Load Pipedrive API Key
with open('src/creds.json') as f:
    data = json.load(f)

## Various API URLs & Credentials
pipedrive_api_token = data['pipedrive_api_key']
pipedrive_api_url = f'https://api.pipedrive.com/v1/itemSearch?api_token={pipedrive_api_token}'
scope = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
credentials = ServiceAccountCredentials.from_json_keyfile_dict(data, scope)
gc = gspread.authorize(credentials)

# Open Google Sheet for input
spreadsheet = gc.open('Management Tool')
worksheet = spreadsheet.get_worksheet(3)


print(len(clients))


## Format Functions
# def formatName(name):
#     firstname = name.split(' ')[0]
#     lastname = name.split(' ')[1]
#     return firstname, lastname

def formatAddress(address):
    parts = address.split(',', 5)
    parts += [None] * (6 - len(parts))

    add1, add2, add3, add4, add5, add6 = [part.strip() if part else None for part in parts]

    return add1, add2, add3, add4, add5, add6


# Get all values in column A (column index 1)
usernames_in_sheet = worksheet.col_values(1)

# Create a dictionary mapping usernames to their indices
username_to_index = {username: index for index, username in enumerate(usernames_in_sheet, start=1)}  # +1 because gspread row indices start at 1


## Main Loop. Retrieve Pipedrive ID from each username, use ID to retrieve full client detail, format and write to Google Sheet
for username in clients:
    params = {
        'term': username,
        'fields': 'custom_fields',
        'exact_match': 'true',
        'item_type': 'person'
    }

    # Get ID from Pipedrive
    response = requests.get(pipedrive_api_url, params=params)

    if response.status_code == 200:
        data = response.json()

        if 'data' in data and 'items' in data['data'] and data['data']['items']:
            person_id = data['data']['items'][0]['item']['id']

            # Get Client Details from Pipedrive
            search = requests.get(f'https://api.pipedrive.com/v1/persons/{person_id}?api_token={pipedrive_api_token}')

            # If search is successful (finds client in Pipedrive), format and write to Google Sheet
            if search.status_code == 200:
                data = search.json()
                print("Found User", username)

                if 'data' in data:
                    person = data['data']

                    firstname = person['first_name']
                    lastname = person['last_name']
                    email = person['email'][0]['value'] if person['email'] else None
                    phone = person['phone'][0]['value'] if person['phone'] else None
                    dob = person['92dd61ea1eede066a7a0ed21f88b2010bf7acc6c']
                    post_code = person['580978d2def7ff57b73b15bbbcbb951808d270b2_postal_code']


                    address = person['580978d2def7ff57b73b15bbbcbb951808d270b2'] if '580978d2def7ff57b73b15bbbcbb951808d270b2' in person else None
                    add1, add2, add3, add4, add5, add6 = formatAddress(address) if address else None
                    depositlimit = person['40a26c77465eaf870b6d79060506707a83bd0c20']
                    depositperiod = person['4db8c878dd2b8f5cf923f5a4a078fe738cfe19be']
                    dateregistered = person['d20f5843b941084d535fe0bc4377160ff2d3aceb'] if 'd20f5843b941084d535fe0bc4377160ff2d3aceb' in person and person['d20f5843b941084d535fe0bc4377160ff2d3aceb'] else person.get('3cf81a9598fbad65a194fad42cd50542ff26f475', None)            


                    print(add1, add2, add3, add4, depositlimit, depositperiod, dateregistered)
                    # Get all values in column A (column index 1)
                    # usernames_in_sheet = worksheet.col_values(1)

                    # If username is found in the sheet
                    if username in usernames_in_sheet:
                        # Get the row index
                        row_index = username_to_index[username]
                        print("Printing data to sheet")
                        # Update the required data points in the sheet
                        worksheet.update_cell(row_index, 2, firstname)  # Column B
                        worksheet.update_cell(row_index, 3, lastname)  # Column C
                        worksheet.update_cell(row_index, 13, add1)  # Column M
                        worksheet.update_cell(row_index, 14, add2)  # Column N
                        worksheet.update_cell(row_index, 15, add3)  # Column O
                        worksheet.update_cell(row_index, 16, add4)  # Column P
                        worksheet.update_cell(row_index, 17, post_code)  # Column Q
                        worksheet.update_cell(row_index, 19, email)  # Column S
                        worksheet.update_cell(row_index, 20, phone)  # Column T
                        worksheet.update_cell(row_index, 21, depositlimit)  # Column U
                        worksheet.update_cell(row_index, 22, depositperiod)  # Column V
                        worksheet.update_cell(row_index, 23, dob)  # Column W
                        worksheet.update_cell(row_index, 24, dateregistered)  # Column X
                        print("Successfully added user to sheet", username)
                        
                        time.sleep(3)

                    else:
                        print(f"Failed to find {username} in the sheet, skipping.")
            else:
                print(f"Failed to retrieve client details for {username}")

        else:
            print(f"Failed to find {username} in Pipedrive")

# for username in clients:

#     params = {
#     'term': username,
#     'fields' : 'custom_fields',
#     'exact_match': 'true',
#     'item_type': 'person'
#     }

#     ## Get ID from Pipedrive
#     response = requests.get(pipedrive_api_url, params=params)

#     if response.status_code == 200:
#         data = response.json()

#         if 'data' in data and 'items' in data['data'] and data['data']['items']:
#             person_id = data['data']['items'][0]['item']['id']
#             #print(person_id)


#     ## Get Client Details from Pipedrive
#     search = requests.get(f'https://api.pipedrive.com/v1/persons/{person_id}?api_token={pipedrive_api_token}')
    
#     ## If search is successful (finds client in Pipedrive), format and write to Google Sheet
#     if search.status_code == 200:
#         data = search.json()
#         #print(data)

#         if 'data' in data:
#             person = data['data']

#             #username = person['c1f84d7067cae06931128f22af744701a07b29c6']
#             #dob = person['92dd61ea1eede066a7a0ed21f88b2010bf7acc6c']
#             firstname = person['first_name']
#             lastname = person['last_name']
#             email = person['email'][0]['value'] if person['email'] else None
#             phone = person['phone'][0]['value'] if person['phone'] else None
#             post_code = person['580978d2def7ff57b73b15bbbcbb951808d270b2_postal_code']

#             print(username, firstname, lastname, email, phone, post_code)  

#             worksheet.append_row([username, firstname, lastname, email, phone, post_code])


#             # formatted_person = [username, firstname, lastname, add1, add2, add3, add4, add5, add6, email, phone, depositlimit, depositperiod, dob, dateregistered]
#             # print(formatted_person)

#             # worksheet.append_row(formatted_person)


