#!/usr/bin/env python

import os
import sys
import re
import tika
import openpyxl

E = 'EXTRAVERSION'
I = 'INTROVERSION'
N = 'INTUITION'
S = 'SENSING'
T = 'THINKING'
F = 'FEELING'
J = 'JUDGEING'
P = 'PERCEIVING'
#n=len(sys.argv)

tika.initVM()
from tika import parser
parsed = parser.from_file(sys.argv[1])
content=parsed["content"]
#print(content)
name=['First','Last']
dirtyname=(sys.argv[1].split("-"))
cleanname=[re.sub(r'[^a-zA-Z0-9]','',string) for string in dirtyname]
namedict=dict(zip(name, cleanname[0:2]))	
print("processing:",namedict)
def find_between(s, first, last):
	try:
		start = s.index(first) + len(first)
		end = s.index(last, start)
		return s[start:end]
	except ValueError:
		return 

COP = find_between(content, "CLARITY OF YOUR PREFERENCES: ", '\n')

def intensity():
	try:
		intensity1 = [int(i) for i in find_between(content, "INTROVERSION  |", '\n').split() if i.isdigit()]
		return intensity1
	except:
		pass
	try:
		intensity2 = [int(i) for i in find_between(content, "EXTRAVERSION  |", '\n').split() if i.isdigit()]
		return intensity2
	except:
	
		pass
	
#print(COP)
#try:
#	print(intensity1)
#except:
#	pass
#try:
#	print(intensity2)
#except:
#	pass

coplist=[]
if COP[0] == "E":
	coplist.append(E)
else:
	coplist.append(I)
if COP[1] == "N":
	coplist.append(N)
else:
	coplist.append(S)
if COP[2] == "T":
	coplist.append(T)
else:
	coplist.append(F)
if COP[3] == "J":
	coplist.append(J)
else:
	coplist.append(P)


data1=intensity()
datadict=dict(zip(coplist, data1))
#print(datadict)

from openpyxl import load_workbook

workbook=load_workbook(filename="Class_Summary.xlsx")
sheet=workbook.active

sheet_rows = tuple(sheet.rows)

print(sheet_rows)
row_count=len(sheet_rows)
next_row=row_count+1
#print('row count'+' ', row_count)
sheet["A"+str(next_row)] = namedict['First']
sheet["B"+str(next_row)] = namedict['Last']
sheet["C"+str(next_row)] = COP
try:
	sheet["D"+str(next_row)] = datadict['EXTRAVERSION']
except:
	pass
try:
	sheet["E"+str(next_row)] = datadict['INTROVERSION']
except:
	pass
try:
	sheet["F"+str(next_row)] = datadict['INTUITION']
except:
	pass
try:
	sheet["G"+str(next_row)] = datadict['SENSING']
except:
	pass
try:
	sheet["H"+str(next_row)] = datadict['THINKING']
except:
	pass
try:
	sheet["I"+str(next_row)] = datadict['FEELING']
except:
	pass
try:
	sheet["J"+str(next_row)] = datadict['JUDGEING']
except:
	pass
try:
	sheet["K"+str(next_row)] = datadict['PERCEIVING']
except:
	pass
	

workbook.save(filename="Class_Summary.xlsx")
