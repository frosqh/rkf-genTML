#!/usr/bin/env python
#-*- coding:utf-8 -*-

import codecs
import os
import random as r
import sys
import time
import webbrowser
import xlrd

from colorama improt Fore, Style, Init

CRLF = '\r\n'

init()

def toInt(v):
	try:v=int(v) if int(v)==v else v
	except:pass
	return v

def overwrite(s,c,t):
	done='▓'*c
	todo='-'*(t-c-1)
	g='\r'+s+Fore.YELLOW+'<'+done+todo+'>'+Fore.GREEN
	print(g,end='')

try:
	print(Fore.GREEN)
	p="Récupération des données depuis le classeur xls "
	print(p,end='')
	workbook=xlrd.open_workbook("../Classeur1.xls",encoding_override="cp1252")
	l=["Niveau de Gamme","Famille","Produit","Forme","TailleDimension","Tissage","Matière","Garnissage","Grammage","Coloris","Spécificitésé","Confectionneurs"]
	m=[]
	c=0
	for s in l:
		overwrite(p,c,len(l))
		ra=r.random()*0.5
		time.sleep(ra)
		k=[]
		worksheet=workbook.sheet_by_name(s)
		b=1
		i=1
		while b:
			try:
				v=toInt(worksheet.cell(i,1).value)
				v2=toInt(worksheet.cell(i,2).value)
				i+=1
				k.append("<tr><td>"+str(v)+"</td><td>"+str(v2)+"</td></tr>\n")
			except:b=0
		m.append(k)
		c+=1
	print(CRLF)
	c=0
	t="Rendu du template "
	print(t,end='')
	with codecs.open("../template.html","rb",encoding="utf-8") as f:data=f.readlines()
	with codecs.open("../output.html","w",encoding="utf-8") as f:
		for d in data:
			if "%tabular%" in p:
				overwrite(t,c,len(l))
				ra=r.random()*0.25
				time.sleep(ra)
				k=m[c]
				for s in k:f.write('\t'+s)
				c+=1
			else: f.write(p[:-1])
		f.write('>')
	print(CRLF)
	print("Rendu complété, ficier généré : "+Fore.BLUE+"../output.html"+Style.RESET_ALL)
	os.system("pause")
	webbrowser.open("../output.html")
except Exception as e:
	print(CRLF+CRLF+Fore.Red+'Error : '+str(e))
	os.system("pause")
	exit(-1)