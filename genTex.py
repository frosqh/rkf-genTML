#!/usr/bin/env python
#-*- coding: utf-8 -*-

import xlrd
import time
import sys
import codecs
import random as r
import os
import webbrowser
from colorama import Fore,Style,init

init()

###
# Convert a read string into int if possible (used for convert from float format) 
###
def toInt(v):
	try:v = int(v) if int(v)==v else v
	except:pass
	return v

###
# Print a loading bar according to the label (s), the progress (c), and the size (t).
###
def overwrite(s,c,t):
	done = '▓'*c
	todo = '-'*(t-c-1)
	g ='\r'+s+Fore.YELLOW+'<'+done+todo+'>'+Fore.GREEN
	print(g,end='',flush=True)
	
try:
	print(Fore.GREEN)
	print('Récupération des données depuis le classeur xls ',end='')
	workbook = xlrd.open_workbook("../Classeur1.xls",encoding_override="cp1252")
	l=["Niveau de Gamme","Famille","Produit","Forme","TailleDimension","Tissage","Matière","Garnissage","Grammage","Coloris","Spécificités","Confectionneurs"]
	m=[]
	c=0
	for s in l:
		overwrite("Récupération des données depuis le classeur xls ",c,len(l))
		ra = r.random()*0.5
		time.sleep(ra)
		k=[]
		worksheet = workbook.sheet_by_name(s)
		b=1
		i=1
		while b:
			try:
				v = toInt(worksheet.cell(i,1).value)
				v2 = toInt(worksheet.cell(i,2).value)
				i+=1
				k.append("<tr><td>"+str(v)+"</td><td> "+str(v2)+" </td></tr>\n")
			except:b=0
		m.append(k)
		c+=1
	print('\n')
	c=0
	print("Rendu du template ",end='',flush=True)
	with codecs.open("../template.html","rb",encoding="utf8") as f:data=f.readlines()
	with codecs.open("../output.html","w",encoding="utf8") as f:
		for d in data:
			p = d
			if "%tabular%" in p:
				overwrite("Rendu du template ",c,len(l))
				ra = r.random()*0.25
				time.sleep(ra)
				k = m[c]
				for s in k: f.write('\t'+s)
				c+=1
			else: f.write(p[:-1])
		f.write(">")
	print('\n')
	print("Rendu complété, fichier généré : "+Fore.BLUE+"../output.hstml")
	print()
	print(Style.RESET_ALL)
except Exception as e:
	print('\n\n'+Fore.RED+'Error : '+str(e))

os.system("pause")
webbrowser.open("../output.html",new=0)