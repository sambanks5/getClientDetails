import requests
import json
import gspread
from oauth2client.service_account import ServiceAccountCredentials

## Load list of Client usernames
with open('clients.txt') as f:
    clients = f.read().split(',')

## Load Pipedrive API Key
with open('src/creds.json') as f:
    data = json.load(f)

## Various API URLs & Credentials
pipedrive_api_token = data['pipedrive_api_key']
pipedrive_api_url = f'https://api.pipedrive.com/v1/itemSearch?api_token={pipedrive_api_token}'
scope = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
credentials = ServiceAccountCredentials.from_json_keyfile_dict(data, scope)
gc = gspread.authorize(credentials)

## Open Google Sheet for output
sheet = gc.open('formattedclients').sheet1

## Format Functions
def formatName(name):
    firstname = name.split(' ')[0]
    lastname = name.split(' ')[1]
    return firstname, lastname

def formatAddress(address):
    parts = address.split(',', 5)
    parts += [None] * (6 - len(parts))

    add1, add2, add3, add4, add5, add6 = [part.strip() if part else None for part in parts]

    return add1, add2, add3, add4, add5, add6


## Main Loop. Retrieve Pipedrive ID from each username, use ID to retrieve full client detail, format and write to Google Sheet
for username in clients:

    params = {
    'term': username,
    'fields' : 'custom_fields',
    'exact_match': 'true',
    'item_type': 'person'
    }

    ## Get ID from Pipedrive
    response = requests.get(pipedrive_api_url, params=params)

    if response.status_code == 200:
        data = response.json()

        if 'data' in data and 'items' in data['data'] and data['data']['items']:
            person_id = data['data']['items'][0]['item']['id']
            print(person_id)


    ## Get Client Details from Pipedrive
    search = requests.get(f'https://api.pipedrive.com/v1/persons/{person_id}?api_token={pipedrive_api_token}')
    
    ## If search is successful (finds client in Pipedrive), format and write to Google Sheet
    if search.status_code == 200:
        data = search.json()

        if 'data' in data:
            person = data['data']

            username = person['c1f84d7067cae06931128f22af744701a07b29c6']
            dob = person['92dd61ea1eede066a7a0ed21f88b2010bf7acc6c']
            firstname = person['first_name']
            lastname = person['last_name']
            email = person['email'][0]['value'] if person['email'] else None
            phone = person['phone'][0]['value'] if person['phone'] else None
            address = person['580978d2def7ff57b73b15bbbcbb951808d270b2'] if '580978d2def7ff57b73b15bbbcbb951808d270b2' in person else None
            add1, add2, add3, add4, add5, add6 = formatAddress(address) if address else None
            depositlimit = person['40a26c77465eaf870b6d79060506707a83bd0c20']
            depositperiod = person['4db8c878dd2b8f5cf923f5a4a078fe738cfe19be']
            dateregistered = person['d20f5843b941084d535fe0bc4377160ff2d3aceb'] if 'd20f5843b941084d535fe0bc4377160ff2d3aceb' in person and person['d20f5843b941084d535fe0bc4377160ff2d3aceb'] else person.get('3cf81a9598fbad65a194fad42cd50542ff26f475', None)            

            formatted_person = [username, firstname, lastname, add1, add2, add3, add4, add5, add6, email, phone, depositlimit, depositperiod, dob, dateregistered]
            print(formatted_person)

            sheet.append_row(formatted_person)


