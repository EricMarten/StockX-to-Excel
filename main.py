from __future__ import print_function
import requests
import json
import openpyxl
import re
import os
import string

class Stockx():
    API_BASE = 'https://stockx.com/api'

    def __init__(self):
        self.customer_id = None
        self.headers = None

    def __api_query(self, request_type, command, data=None):
        endpoint = self.API_BASE + command
        response = None
        if request_type == 'GET':
            response = requests.get(endpoint, params=data, headers=self.headers)
        elif request_type == 'POST':
            response = requests.post(endpoint, json=data, headers=self.headers)
        elif request_type == 'DELETE':
            response = requests.delete(endpoint, json=data, headers=self.headers)
        return response.json()

    def __get(self, command, data=None):
        return self.__api_query('GET', command, data)

    def __post(self, command, data=None):
        return self.__api_query('POST', command, data)

    def __delete(self, command, data=None):
        return self.__api_query('DELETE', command, data)

    def authenticate(self, email, password):
        endpoint = self.API_BASE + '/login'
        payload = {
            'email': email,
            'password': password
        }
        response = requests.post(endpoint, json=payload)
        customer = response.json().get('Customer', None)
        if customer is None:
            raise ValueError('Authentication failed, check username/password')
        self.customer_id = response.json()['Customer']['id']
        self.headers = {
            'JWT-Authorization': response.headers['jwt-authorization']
        }
        return True

    def selling(self):
        command = '/customers/{0}/selling'.format(self.customer_id)
        response = self.__get(command)
        return response['PortfolioItems']

stockx = Stockx()

def settings(config):
	try:
		with open(config, 'r') as myfile:
			return json.load(myfile)
	except IOError:
		with open(config, 'w') as myfile:
			json.dump({}, myfile)
		return {}

def header(title):
	os.system('cls' if os.name == 'nt' else 'clear')
	print('\n{}\n{}\n'.format(title, '-' * len(title)))

def json_to_title(json):
	return re.sub(r'(\w)([A-Z])', r'\1 \2', json).title().replace('Shoe Size', 'Size')

def setup_workbook(attributes, market_attributes, workbook_name):
	wb = openpyxl.Workbook()
	ws = wb.active
	headers = [json_to_title(item) for item in attributes + market_attributes]
	for i in range(0, len(headers)):
		cell = ws[list(string.ascii_uppercase)[i + list(string.ascii_uppercase).index('B')] + str(2)]
		cell.value = headers[i]
		cell.alignment = openpyxl.styles.Alignment(horizontal='center', vertical='center')
	wb.save(workbook_name)

def write_workbook(email, password, attributes, market_attributes, workbook_name):
	wb = openpyxl.load_workbook(workbook_name)
	ws = wb.active
	if stockx.authenticate(email, password):
		i = 0
		for item in stockx.selling():
			if item['text'] == 'Asking':
				product = item['product']
				for k in range(0, len(attributes)):
					cell = ws[list(string.ascii_uppercase)[k + list(string.ascii_uppercase).index('B')] + str(3 + i)]
					cell.value = product[attributes[k]]
					cell.alignment = openpyxl.styles.Alignment(horizontal='center', vertical='center')
				if market_attributes != []:
					market = product['market']
					for k in range(len(attributes), len(attributes) + len(market_attributes)):
						cell = ws[list(string.ascii_uppercase)[k + list(string.ascii_uppercase).index('B')] + str(3 + i)]
						cell.value = market[market_attributes[k - len(attributes)]]
						cell.alignment = openpyxl.styles.Alignment(horizontal='center', vertical='center')
				i += 1
	wb.save(workbook_name)

header('StockX to Excel by @DefNotAvg')

print('Loading config.json...\r', end='')
config = settings('config.json')
email = config['email']
password = config['password']
attributes = config['attributes']
market_attributes = config['marketAttributes']
workbook_name = config['workbookName']
print('Successfully loaded config.json!!')

print('Setting up the Excel Workbook...\r', end='')
setup_workbook(attributes, market_attributes, workbook_name)
print('Successfully set up the Excel Workbook!!')

print('Writing StockX data to {}...\r'.format(workbook_name), end='')
write_workbook(email, password, attributes, market_attributes, workbook_name)
print('Successfully exported StockX data to {}!!'.format(workbook_name))