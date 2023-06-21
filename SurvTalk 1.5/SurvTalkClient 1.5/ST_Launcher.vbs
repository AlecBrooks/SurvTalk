on error resume next

Dim fso, f, FSO2
dim WshShell, strCurDir, obj
dim CopyFrom, CopyTo, RandomName
dim nlow, nhigh

Set fso = CreateObject("Scripting.FileSystemObject")
Set f = fso.OpenTextFile("ST_ConfigFile.conf",1)
ReadAllTextFile = f.ReadAll
indexStart  = InStr(1, ReadAllTextFile, "=")
ReadAllTextFile = Right(ReadAllTextFile, Len(ReadAllTextFile) - indexStart)
x = Instr(ReadAllTextFile,";")
ServerPath = Left(ReadAllTextFile ,x-1)
indexStart  = InStr(1, ReadAllTextFile, "=")
ReadAllTextFile = Right(ReadAllTextFile, Len(ReadAllTextFile) - indexStart)
x = Instr(ReadAllTextFile,";")
versionNumber = Left(ReadAllTextFile ,x-1)

nLow = 10
nHigh = 99

randomize
RandomName = int((nHigh - Nlow + 1) * Rnd + nLow)

set WshShell = createobject("Wscript.shell")
strCurDir = WshShell.CurrentDirectory

CopyFrom = strCurDir & "\ST_" & versionNumber & ".xlsm"
CopyTo =  WshShell.ExpandEnvironmentStrings("%APPDATA%") & "\Client_" & RandomName & ".xlsm"

set FSO2 = CreateObject("Scripting.FileSystemObject")
FSO2.CopyFile CopyFrom,CopyTo

clientPath = strCurDir & "\"

set objexcel = createobject("Excel.application")
objexcel.visible = false
set objWorkbook = objexcel.workbooks.Open(CopyTo)
objexcel.Cells(2, 1).Value = ServerPath
objexcel.Cells(5, 1).Value = versionNumber 
objexcel.Cells(8, 1).Value = clientPath
objexcel.Run "ST_Controller.boot" 

