# -*- coding: utf-8 -*-
from bs4 import BeautifulSoup
from openpyxl import load_workbook
import pandas as pd
import requests
import urllib
import json
import os
# 查詢簡單地址的里 https://www.ris.gov.tw/app/portal/3053




def get_address(address_whole, county_get = False):
	'''將地址分解成區、街（路）、弄、之與號
	Args:
		address_whole: 完整地址
		county_get: 是否要抓「區」

	Return:
		address_list: 將 address_whole 分解後的地址 list。
		若有包含 county，[county, road, lane, number, number1]，沒包含 county，[road, lane, number, number1]
	'''

	address_list = []
	
	
	print address_whole.decode("utf-8").encode("big5")
	
	address = address_whole.split("區")

	##### 地址已含有村、里 #####
	# 地址含有里直接 return
	if len(address) == 2:
		# 地址含區，避免抓到「八里區」和「萬里區」，split 區後，里含在 list 第二個
		if "里" in address[1]:
			village = address[1].split("里")[1] + "里"
			address_dist = [village]
			return address_dist
	# 地址不含區
	else:
		if "里" in address_whole:
			village = address[0].split("里")[0] + "里"

	# 地址含有村直接 return
	if "村" in address_whole:
		if len(address) == 2:
			village = address[1].split("村")[1] + "村"

		else:
			village = address[0].split("村")[0] + "村"
			
		address_dist = village
		
	##### 區 #####
	if county_get == True:
		if len(address) == 2:	
			county = address[0] + "區"
		address = address[1]
	else:
		if county_get == True:
			county = df.iat[i, 1] + "區"
		address = address[0]

	##### 路、街 #####
	try:			
		if len(address.split("街")) == 1:
			address = address.split("路")
			road = address[0] + "路"
			address = address[1] 
		else:
			address = address.split("街")
			road = address[0] + "街"
			address = address[1]
	except:
		address_dist[address_whole] = []
		
	##### 巷 #####
	lane = address.split("巷")
	if len(lane) != 1:
		# 半形轉全形
		lane = halfwidth_to_fullwidth(lane[0])
		address = address.split("巷")[1]
	else:
		lane = ""

	##### 弄 #####
	alley = address.split("弄")
	if len(alley) != 1:
		alley = halfwidth_to_fullwidth(alley[0])
		address = address.split("弄")[1]
	else:
		alley = ""

	##### 之 #####
	number1 = address.split("之")
	if len(number1) != 1 and ("號" in number1[1]):
		# 半形轉全形
		number1 = halfwidth_to_fullwidth(number1[0])
		address = address.split("之")[1]
	else:
		number1 = ""

	if number1 == "":
		number1 = address.split("-")
		if len(number1) != 1 and ("號" in number1[1]):
			# 半形轉全形
			number1 = halfwidth_to_fullwidth(number1[0])
			address = address.split("-")[1]
		else:
			number1 = ""

	##### 號 #####
	number_str = str(address.split("號")[0])
	if len(number_str.split("-")) != 1:
		number_str = number_str.split("-")[1]

	# 半形轉全形
	number = halfwidth_to_fullwidth(number_str)

	# return
	if county_get == True:
		address_list = [county, road, lane, number, number1, alley]
	else:
		address_list = [road, lane, number, number1, alley]

	return address_list



def halfwidth_to_fullwidth(number):
	'''將數字半形轉全形
	Args:
		number: 要轉成全形的半形數字

	Return:
		number: 全形數字
	'''
	number_str = str(number)
	number = ""
	for i in range(len(number_str)):
		number = number + unichr(int(number_str[i]) + 48+65248).encode("utf-8")
	return number




def get_city_code():
	# 全台縣市代碼
	url = "https://www.ris.gov.tw/info-doorplate/app/doorplate/map?searchType=doorplate"
	res = requests.get(url)
	soup =  BeautifulSoup(res.text, "html.parser")

	city_dic = {}
	for i in soup.select(".col-xs-12")[1].select("area"):
		city_name = i["alt"].encode("utf-8").replace("資料", "")
		city_code = i["onclick"].replace("toQuery('", "").replace("');", "")
		city_dic[city_name] = city_code
	return city_dic



def get_county_code(city_code):
	# 縣市的各區代碼
	url = "https://www.ris.gov.tw/info-doorplate/app/doorplate/query"
	payload = {
		"cityCode": city_code,
		"searchType": "doorplate"
	}

	res = requests.post(url, data = payload)
	soup = BeautifulSoup(res.text, "html.parser")
	county_dic = {}
	for i in soup.select("#areaCode")[0].select("option")[1:]:
		county_name = i.text.encode("utf-8") 
		county_code = i["value"]
		county_dic[county_name] = county_code

	return county_dic



def get_village(df, i, city_dic):
	address_whole = df.iat[i, 4].encode("utf-8")

	### city ### 
	city = df.iat[i, 0].encode("utf-8")

	### county ###
	try:
		county = df.iat[i, 1].encode("utf-8")[:6] + "區"
		county_get = False
	except:
		county_get = True


	try:
		result = get_address(address_whole, county_get = county_get)
	except:
		return ""

	if county_get:
		county = result[0]
		road = result[1]
		lane = result[2]
		if result[4] != "":
			number = result[4]
			number1 = result[3]
		else:
			number = result[3]
			number1 = result[4]
		alley = result[5]
	
	else:
		road = result[0]
		lane = result[1]
		if result[3] != "":
			number = result[3]
			number1 = result[2]
		else:
			number = result[2]
			number1 = result[3]
		alley = result[4]

	# print city, county, road, lane, alley, number, number1

	
	city_keys = city_dic.keys()

	city_code = ""
	county_code = ""

	for i in range(len(city_keys)):
		# print i
		city_name = city_keys[i]
		if city in city_name:
			city_code = city_dic[city_name]
			# 區代碼
			county_dic = get_county_code(city_code)
			county_keys = county_dic.keys()

			for j in range(len(county_keys)):
				county_name = county_keys[j]
				# print county_name
				if county == county_name:
					# print county_name
					county_code = county_dic[county_name]
					# print county_code
					break

		if county_code != "":
			break

	# 抓取地址查詢結果
	road = urllib.quote(road)
	lane = urllib.quote(lane)
	alley = urllib.quote(alley)
	number = urllib.quote(number)
	number1 = urllib.quote(number1)
	url = "https://www.ris.gov.tw/info-doorplate/app/doorplate/doorplateQuery?searchType=doorplate&cityCode=" + city_code + "&tkt=-1&areaCode=" + county_code + "&village=&neighbor=&street=" + road + "&section=&lane=" + lane + "&alley=" + alley + "&number=" + number + "&number1=" + number1 + "&floor=&ext=&_search=false&nd=1545341958716&rows=20&page=1&sidx=&sord=asc&token="
	# print url
	res = requests.get(url)
	address_json = json.loads(res.text.encode("utf-8"))

	# 處理使用「-」分隔連續號的地址
	if address_json["rows"] == []:
		number1 = ""
		url = "https://www.ris.gov.tw/info-doorplate/app/doorplate/doorplateQuery?searchType=doorplate&cityCode=" + city_code + "&tkt=-1&areaCode=" + county_code + "&village=&neighbor=&street=" + road + "&section=&lane=" + lane + "&alley=" + alley + "&number=" + number + "&number1=" + number1 + "&floor=&ext=&_search=false&nd=1545341958716&rows=20&page=1&sidx=&sord=asc&token="
		res = requests.get(url)
		address_json = json.loads(res.text.encode("utf-8"))
		
	try:
		address = address_json["rows"][0]["v1"].encode("utf-8")
	except:
		return ""

	village = address.split("區")[1].split("里")[0] + "里"
	return village
	


def main():

	path = "file/"
	files = os.listdir(path)
	city_dic = get_city_code()
	# print files
	for file in files:
		print "開始抓取檔案-".decode("utf-8").encode("big5") + file
		file_name = file.replace(".xlsx", "") + "_result.xlsx"

		df = pd.read_excel(path + file, header = None)
		rows = df.shape[0] - 1

		for i in range(1, rows):
			try:
				# 檢查第一行是否為空白
				if pd.isnull(df.iloc[i, 0]):
					break
				print i

				try:
					village = get_village(df, i, city_dic)
					df.iloc[i, 2] = village
				except:
					writer = pd.ExcelWriter(file_name, engine = "openpyxl")
					sub_df = df[:i]
					sub_df.to_excel(writer, sheet_name = "Sheet1", index = False, header = False)
					writer.save()
					print "Save file!"
					continue
				
				if i % 500 == 0:
					try:
						writer = pd.ExcelWriter(file_name, engine = "openpyxl")
						sub_df = df[:i]
						sub_df.to_excel(writer, sheet_name = "Sheet1", index = False, header = False)
						writer.save()
					except:
						continue
			except:
				continue
			
		writer = pd.ExcelWriter(file_name, engine = "openpyxl")
		df.to_excel(writer, sheet_name = "Sheet1", index = False, header = False)
		writer.save()

		print "檔案 ".decode("utf-8").encode("big5") + file + " 完成".decode("utf-8").encode("big5")

	print "全部抓取完成!".decode("utf-8").encode("big5")
	input()






if __name__ == '__main__':
    main()
    input()









