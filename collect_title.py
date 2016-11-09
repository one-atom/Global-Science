# coding: utf-8
import urllib2,urllib
import docx2txt
import os
import pandas as pd
import re
import string
import pdb
import csv
import xlwt
import sys
import argparse
import docx
from xlrd import open_workbook
from xlutils.copy import copy

pattern = r'https://www\.scientificamerican\.com/podcast/episode/.+?</a>'
title_pattern = r'>[A-Z0-9].+?<'
url_pattern = r'https://.+?\"'


def analysis_single_record(text):
	title = re.search(title_pattern, text).group()[1:-1]
	url = re.search(url_pattern, text).group()[:-1]
	return title, url

def search_page(i):
	headers = {'User-Agent':'Mozilla/5.0 (Windows; U; Windows NT 6.1; en-US; rv:1.9.1.6) Gecko/20091201 Firefox/3.5.6'}
	search_url = 'https://www.scientificamerican.com/podcast/60-second-science/?page=' + str(i)
	req = urllib2.Request(search_url, headers=headers)  
	content = urllib2.urlopen(req).read()
	search_result = re.findall(pattern, content)
	titles = []
	for item in search_result:
		title,url = analysis_single_record(item)
		titles.append(title)
	return titles

def main():
	# opening and writing xlsx file refers to http://m.blog.csdn.net/article/details?id=8052746
	#rb = open_workbook('任务列表.xlsx')
	#rs = rb.sheet_by_index(0)
	#wb = copy(rb)

	dictionary = []
	for i in range(1,3):
		new_dict = search_page(i)
		dictionary += new_dict
	open('titles.txt','w').write('\n'.join(dictionary))
	

if __name__ == '__main__':
	main()