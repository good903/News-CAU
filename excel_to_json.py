# -*- coding:utf-8 -*-
import io
import sys
import pandas as pd
import json
from collections import OrderedDict

class NewsInfo :
   def __init__(self, publisher, title, date, link, lod, mods, subject, types, coverage, original_keywords, keywords):
	self.publisher = publisher
	self.title = title
        self.date = date
        self.link = link
        self.lod = lod
        self.mods = mods
        self.subject = subject
        self.types = types
        self.coverage = coverage
        self.original_keywords = original_keywords
        self.keywords = keywords


excel_file_name = "news.xlsx"
json_file_name = "json_news/n"
xlsx = pd.read_excel(open(excel_file_name.decode('utf-8'), 'r'), sheet_name = "Sheet1")

# Create NewsInfo Array(=list)
ni_list = list()
for i in xlsx.index:
	publisher = xlsx[u"publisher"][i]
	title = xlsx[u"title"][i]
	date = xlsx[u"date"][i]
	link = xlsx[u"link"][i]
	lod = xlsx[u"lod"][i]
	mods = xlsx[u"mods"][i]
	try:	
		subject = xlsx[u"subject"][i].strip()
	except (AttributeError) as detail:
		print(detail)
		subject = ""

	try:
		types = xlsx[u"type"][i].strip()
	except (AttributeError) as detail:
		print(detail)
		types = ""

	try:
		coverage = xlsx[u"coverage"][i].strip()
	except (AttributeError) as detail:
		print(detail)
		coverage = ""
	
	try:
		original_keywords = xlsx[u"original keywords"][i].split(",")
		keywords = xlsx[u"keywords"][i].split(",")
	except (AttributeError) as detail:
		print(detail)
		original_keywords = []
		keywords = []	

	ni = NewsInfo(publisher, title, date, link, lod, mods, subject, types, coverage, original_keywords, keywords)
	nis = json.dumps(ni.__dict__, indent=4).decode('unicode-escape').encode('utf8')
	with open(json_file_name + str(i).zfill(7), 'w') as make_file:
		make_file.write(nis + '\n')





