' **************
' Process
' **************
sub Main(pmWb as workbook)
	dim path as string: path = pm.quarterPath & "Data\MAG\" & pm.year & "Q" & pm.quarter & " Magnastar Settlement PLA.xlsx"
	Dim workingDir as string: workingDir = PM.scriptPath & "MAG\"
	Dim wb As Workbook: Set wb = PMUtil.GetWorkbook(path)
	Dim ws as Worksheet: Set ws = wb.Worksheets("YRT Premiums")
	Dim QData as Range: Set QData = PMUtil.GetNamedRange("Q" & pm.quarter &"Data", wb)
	Dim MagYRT As Range:  Set MagYRT = PMUtil.GetNamedRange("MagYRT", wb).offset(0,(pm.quarter-1)*5)
	Dim MagTrialBalance As Range:  Set MagTrialBalance = PMUtil.GetNamedRange("MagTrialBalance", wb)
	Dim MagOverhead As Range:  Set MagOverhead = PMUtil.GetNamedRange("MagOverhead", wb)
	
	'Trial Balance data goes in same place every quarter
	MagTrialBalance.Clear

	'###################Yusheng is working####################
    'Clean overhead data that is on or after this quarter for rerunning
    Dim cell As Range
    For Each cell In Sheets("Allocations").Range("C5:C200")
        If cell.Value >= quarter Then
            Sheets("Allocations").Range("A" & cell.row, "I" & cell.row).ClearContents
        End If
    Next cell    
    'Cleanup Qq database before rerunning
	Qdata.Clear
	'Update links
	wb.sheets("YTD Settlement").Range("D78").Formula = "='J:\Acctng\QuarterClose\" & pm.year & "\Q" & pm.quarter & "\Data\MAG\[" & pm.year & "Q" & pm.quarter & " PLA DAC Tax.xlsx]M Summary'!$C$4"
	'Update YTD Database formula
	if pm.quarter <> 1 then
		wb.Sheets("YTD Database").Range("Z:AC").Replace What:="Q" & pm.quarter - 1, Replacement:="Q" & pm.quarter, LookAt:=xlPart, _
			SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
			ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
	End if
	'#########################################################


	Dim overheadrow as Integer
	'Overhead data goes below existing data
	If pm.quarter <> 1 Then
		overheadrow = MagOverhead.End(xldown).Row -4
		Set MagOverhead = MagOverhead.offset(overheadrow,0)
	End If
	
	Call SQL(workingdir & "MagYRTdb11.sqL", pm.server2,pm.database2,"PLA",MagYRT)
	'PL includes margin sharing column
	Call SQL(workingdir & "MagCombined_PL.sqL", pm.server,pm.database,"PLA",MagTrialBalance,MagOverhead,QData)
	
	For Each x in MagYRT
		If Isnumeric(x.value) and not isempty(x.value) Then
			x.value = 1* x.value
		End If
	Next x
	
	For Each x in QData
		If Isnumeric(x.value) and not isempty(x.value) Then
			x.value = 1* x.value
		End If
	Next x
	
	wb.Save
	wb.Close
end sub

Public Function SQL(path As String, srvr As String, db As String, carrierID As String, ParamArray destinations() As Variant) As Boolean
    If Not debugMode Then
        On Error GoTo error
    End If
    Dim fso As FileSystemObject: Set fso = New FileSystemObject
    Dim ts As TextStream: Set ts = fso.OpenTextFile(path, ForReading, False)
    Dim header As String: header = BuildSqlHeader(carrierID)
    Dim qry As String: qry = header & ts.ReadAll
    
    SQL = Query(qry, srvr, db, destinations)
	Exit Function
error:
    SQL = True
End Function

Public Function BuildSqlHeader(carrierID As String) As String
    BuildSqlHeader = "declare @year int; set @year = " & year & ";" & Chr(13)
    BuildSqlHeader = BuildSqlHeader & "declare @quarter int; set @quarter = " & quarter & ";" & Chr(13)
    BuildSqlHeader = BuildSqlHeader & "declare @carrierID varchar(3); set @carrierID = '" & carrierID & "';" & Chr(13)
    BuildSqlHeader = BuildSqlHeader & Chr(13)
End Function


Public Function Query(qry As String, srvr As String, db As String, ParamArray destinations() As Variant) As Boolean
    If Not debugMode Then
        On Error GoTo error
    End If
    Dim dst As Variant
    dst = IIf(IsArray(destinations(LBound(destinations))), destinations(LBound(destinations)), destinations)
    
    Dim conn As ADODB.Connection: Set conn = New ADODB.Connection
    conn.connectionString = "Driver={SQL Server};" & _
                        "Server=" & srvr & ";" & _
                        "Database=" & db & ";" & _
                        "Trusted_Connection=True;"
    conn.CommandTimeout = 0
    conn.Open
    Dim records As ADODB.recordset
    Set records = conn.Execute(qry)
    
    For Each d In dst
        Do While records.State = adStateClosed
            Set records = records.NextRecordset
            If records Is Nothing Then
                Exit For
            End If
        Loop
        
        d.CopyFromRecordset records
        Set records = records.NextRecordset
        If records Is Nothing Then
            Exit For
        End If
    Next d
    conn.Close
    Query = False
    
    Exit Function

error:
    conn.Close
    Query = True
End Function



' **************
' Wrappers
' **************
public function Wrapper(pmWb as workbook) as String
	if not pm.debugMode then
		on error goto error
	end if
	Call Main(pmWb)
	on error goto 0
	Wrapper = "Success"
	exit function
error:
	Wrapper = "Failure within module"
end function

' public sub RunQuery(path as string, dst as range) 
	' if pm.runquery(path, dst) then
		' Err.Raise vbObjectError + 513, "PM::RunQuery()", "An error occurred in pm.RunQuery."
	' end if
' end sub