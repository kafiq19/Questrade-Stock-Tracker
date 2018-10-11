import requests
import sys

def get_access_code(refresh):
	url = 'https://login.questrade.com/oauth2/token?grant_type=refresh_token&refresh_token={}'.format(refresh)
	headers = {}
	rt = requests.get(url, headers=headers)
	if rt.status_code == 200:
		response = rt.json()
		access_code = response.get('access_token')
		uri = response.get('api_server')
		print "Successfully obtained response from Questrade"
		return access_code, uri
	else:
		sys.exit('ERROR: Access Code Not Returned - Try Creating New Refresh Code')

def establish_connection(access_code, url):
	url = url + "v1/accounts"
	headers = {'Authorization': 'Bearer {}'.format(access_code)}
	rt = requests.get(url, headers=headers)
	response = rt.json()
	return response

def get_account_values(access_code, url, account):
	url = url + "v1/accounts/{}/balances".format(account)
	headers = {'Authorization': 'Bearer {}'.format(access_code)}
	rt = requests.get(url, headers=headers)
	response = rt.json()
	return response

def get_account_positions(access_code, url, account):
	url = url + "v1/accounts/{}/positions".format(account)
	headers = {'Authorization': 'Bearer {}'.format(access_code)}
	rt = requests.get(url, headers=headers)
	response = rt.json()
	return response

def get_ticker_id(access_code, url, ticker):
	url = url + "v1/symbols/search?prefix={}".format(ticker)
	headers = {'Authorization': 'Bearer {}'.format(access_code)}
	rt = requests.get(url, headers=headers)
	response = rt.json()
	return response

#returns bitcoin price in CAD
def get_btc_price():
	url = 'https://api.cbix.ca/v1/index'
	headers = {}
	rt = requests.get(url, headers=headers)
	response = rt.json()
	value = float(response.get('index').get('value'))
	return value