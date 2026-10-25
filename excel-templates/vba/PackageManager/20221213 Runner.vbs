'=================================================
'The purpose of this script is to call the RSDB Package Manager 
'and run one or all the items that are scripted in it.
'NOTE: This will copy from the the previous quarter if the PM is missing
'=================================================
'User Define Veriables
yr = Wscript.Arguments(0) 
quarter = Wscript.Arguments.Item(1)
server = Wscript.Arguments.Item(2)
scriptinp = Wscript.Arguments.Item(3)
db = "ReinsuranceSettlements" 'Wscript.Arguments.Item(4)

'Scripts to run
dim script(9)
dim names(9)
dim printer(10)'+1 above
script(0) = "a":	names(0) = "ALL"
script(1) = "s":	names(1) = "S2"
script(2) = "sw":	names(2) = "Swiss"
script(3) = "c":	names(3) = "CLife"
script(4) = "hr":	names(4) = "HLRRBC"
script(5) = "mr":	names(5) = "MUNICHRBC"
script(6) = "h3":	names(6) = "HLR03"
script(7) = "h5":	names(7) = "HLR05"
script(8) = "h6":	names(8) = "HLR06"
script(9) = "gaap": names(9) = "GAAP"

'Help logging
if scriptinp = "" then
	msgbox "Please input script code to be run! h for help."
	Wscript.Quit
else 
	if scriptinp = "h" then
		printer(0) = "Script Process Codes:"
		For x =lbound(script) to ubound(script)
		printer(x+1) = script(x) & " - " & names(x)
		Next
	MsgBox Join(printer, vbCrLf)
	Wscript.Quit
	else
		for k = lbound(script) to ubound(script)
			if scriptinp = script(k) then
				' msgbox "HELP ME!!!!!"
				exit for
			else
				if k >= ubound(script) then
					msgbox scriptinp & " is not a specified input!"
					Wscript.Quit
				end if 
			end if
		next
	end if
end if

'Package Manager saved location
PMpath = "J:\Acctng\QuarterClose\" & yr & "\Q" & quarter & "\Assumed Settlements\"
PM = "Package Manager.xlsm"

'If PM does not exist copy from previous quarter
If FileExists(PMpath & PM) Then
'NOTHING
Else
  if quarter = 1 then 
     priorpath = "J:\Acctng\QuarterClose\" & yr-1 & "\Q4\Assumed Settlements\"
  else
     priorpath = "J:\Acctng\QuarterClose\" & yr & "\Q" & quarter-1 & "\Assumed Settlements\"
  end if
  call CopyFiles(priorpath & PM, PMpath)
End If


'Open PM
Set objExcel = CreateObject("Excel.Application")
Set wb1 = objExcel.Workbooks.Open(PMpath & PM)
objExcel.Application.Visible = True
objExcel.DisplayAlerts = False

'Call Macro to update PM to current period
MacroPath = "Internal.update"
objExcel.Run MacroPath, Cstr(yr), cstr(quarter), cstr(server),cstr(db)

'Run All
if scriptinp = script(0) then
	MacroPath = "pm.RunAll"
	msgbox MacroPath
	wb1.Save
	objExcel.DisplayAlerts = True
	wb1.Close
	objExcel.Quit
	Wscript.Quit
end if
'Run one
For x =lbound(script) to ubound(script)
	if scriptinp = script(x) then
		MacroPath = "pm.RunRow"
		'r = wb1.Range("StartRuns").Offset(x,0).Value
		r = "A" & x+11
		Wscript.Echo "About to run"
		objExcel.Run MacroPath, Cstr(r), "TRUE"
		wb1.Save
		objExcel.DisplayAlerts = True
		wb1.Close
		objExcel.Quit
		Wscript.Echo "Just finished running"
		Wscript.Quit
	end if
Next

wb1.Save
objExcel.DisplayAlerts = True
wb1.Close
objExcel.Quit
msgbox "Script Code out of Range! END!"


'=========================================================================================
'Functions
'=========================================================================================
Function FileExists(FilePath)
  Set fso = CreateObject("Scripting.FileSystemObject")
  If fso.FileExists(FilePath) Then
    FileExists=CBool(1)
  Else
    FileExists=CBool(0)
  End If
End Function
'=========================================================================================
Function CopyFiles(FiletoCopy,DestinationFolder)
   Dim fso
                Dim Filepath,WarFileLocation
                Set fso = CreateObject("Scripting.FileSystemObject")
                If  Right(DestinationFolder,1) <>"\"Then
                    DestinationFolder=DestinationFolder&"\"
                End If
    fso.CopyFile FiletoCopy,DestinationFolder,True
                FiletoCopy = Split(FiletoCopy,"\")

End Function