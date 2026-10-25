' **************
' Process
' **************
sub Main(pmWb as workbook)
	Dim path as string: path = pm.quarterPath & "Data\MAG\" & pm.year & "Q" & pm.quarter & " Magnastar Fees SNY.xlsx"
	Dim workingDir as string: workingDir = PM.scriptPath & "MAG\"
	Dim wb As Workbook: Set wb = PMUtil.GetWorkbook(path)
	Dim ws as Worksheet: Set ws = wb.Worksheets("YTD Fees")
	Dim QData as Range: Set QData = PMUtil.GetNamedRange("Q" & pm.quarter &"Data", wb)
	
	Call SQL(workingdir & "MagSun.sqL", pm.server,pm.database,"SNY",QData)
	
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

public sub RunQuery(path as string, dst as range) 
	if pm.runquery(path, dst) then
		Err.Raise vbObjectError + 513, "PM::RunQuery()", "An error occurred in pm.RunQuery."
	end if
end sub