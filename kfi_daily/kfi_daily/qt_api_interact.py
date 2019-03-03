import requests
import os
import sys

#helper: checks if auth code exists in pickle.txt
def pickle_check():
	pickle_file = 'pickle.txt'
	pickle_path = os.path.join(os.path.dirname(__file__), pickle_file)

	try:
		pickle_file = open(pickle_path,"r")
	except:
		print "Pickle File Does Not Exists"
		return False

	pickle_contents = pickle_file.readlines()
	pickle_file.close()
	access_code = pickle_contents[0]
	uri = pickle_contents[1]
	print "Successfully loaded pickle contents"
	return (access_code, uri)

#helper: saves auth code and uri to pickle.txt
def pickler(access_code, uri):
	pickle_file = 'pickle.txt'
	pickle_path = os.path.join(os.path.dirname(__file__), pickle_file)
	pickle_file = open(pickle_path, 'w')
	pickle_file.write(access_code + '\n')
	pickle_file.write(uri)
	print "Successfully created pickle file"
	pickle_file.close()	

#helper: removes pick file if access code expired
def depickler():
	pickle_file = 'pickle.txt'
	os.remove(pickle_file)

def get_access_code(refresh):
	pickle = pickle_check()
	if pickle:
		return pickle[0][:-1], pickle[1]
	else:
		url = 'https://login.questrade.com/oauth2/token?grant_type=refresh_token&refresh_token={}'.format(refresh)
		headers = {}
		rt = requests.get(url, headers=headers)
		if rt.status_code == 200:
			response = rt.json()
			access_code = response.get('access_token')
			uri = response.get('api_server')
			pickler(access_code, uri)
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
	#import pdb; pdb.set_trace()
	response = rt.json()
	value = float(response.get('index').get('value'))
	return value

#backup BTC price
# def get_btc_price():
# 	url = 'https://blockchain.info/ticker'
# 	headers = {}
# 	rt = requests.get(url, headers=headers)
# 	response = rt.json()
# 	btc_CAD = response['CAD']['sell']
# 	return btc_CAD

import pdb; pdb.set_trace()
