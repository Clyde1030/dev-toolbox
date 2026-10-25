Attribute VB_Name = "utils"
Option Explicit

Public Function ExportSht(SaveAsPath As String, CopiedRange As Range, Optional wb As Workbook) As Boolean
    Application.DisplayAlerts = False
    If wb Is Nothing Then
        Set wb = ThisWorkbook
    End If
    Dim ExportRange As Range: Set ExportRange = CopiedRange
    Dim wb2 As Workbook: Set wb2 = Workbooks.Add
    ExportRange.CurrentRegion.Copy
    With wb2.ActiveSheet.Range("A1")
        .cells(1).PasteSpecial xlPasteColumnWidths
        .cells(1).PasteSpecial xlPasteValues
        .cells(1).PasteSpecial xlPasteFormats
    End With

    wb2.SaveAs SaveAsPath
    Application.DisplayAlerts = True
    wb2.Close
    ExportSht = False
    Exit Function

End Function



Public Function SQL(path As String, srvr As String, db As String, year As Integer, month As String, ParamArray destinations() As Variant) As Boolean
    
    On Error GoTo error
    
    Dim fso As FileSystemObject: Set fso = New FileSystemObject
    Dim ts As TextStream: Set ts = fso.OpenTextFile(path, ForReading, False)
    Dim header As String: header = BuildSqlHeader(year, month)
    Dim qry As String: qry = header & ts.ReadAll
    
    SQL = Query(qry, srvr, db, destinations)
    Exit Function

error:
    SQL = True
End Function



Public Function BuildSqlHeader(year As Integer, month As String) As String
    BuildSqlHeader = "declare @year nvarchar(4); set @year = " & year & ";" & Chr(13)
    BuildSqlHeader = BuildSqlHeader & "declare @month nvarchar(2); set @month = '" & month & "';" & Chr(13)
    BuildSqlHeader = BuildSqlHeader & Chr(13)
End Function



Public Function Query(qry As String, srvr As String, db As String, ParamArray destinations() As Variant) As Boolean
    On Error GoTo error
    Dim d
    Dim dst As Variant
    dst = IIf(IsArray(destinations(LBound(destinations))), destinations(LBound(destinations)), destinations)
    
    Dim conn As ADODB.Connection: Set conn = New ADODB.Connection
    conn.ConnectionString = "Driver={SQL Server};" & _
                        "Server=" & srvr & ";" & _
                        "Database=" & db & ";" & _
                        "Trusted_Connection=True;"
    conn.CommandTimeout = 0
    conn.Open
    Dim records As ADODB.Recordset
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



Public Function GetNamedRange(rangeName As String, Optional wb As Workbook) As Range
    If wb Is Nothing Then
        Set wb = ThisWorkbook
    End If
    
    Dim address As String: address = wb.Names.Item(rangeName)
    Dim wsName As String: wsName = Left(address, InStr(1, address, "!"))
    Dim cells As String: cells = Mid(address, InStr(1, address, "!"), 1000)
    
    If (InStr(1, cells, ",") > 0) Then
        Exit Function
    End If
    
    cells = Replace(cells, "!", "")
    wsName = Replace(wsName, "'", "")
    wsName = Replace(wsName, "!", "")
    wsName = Replace(wsName, "=", "")
    
    Set GetNamedRange = wb.Worksheets(wsName).Range(address)
End Function




