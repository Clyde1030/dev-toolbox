Attribute VB_Name = "GetData"
Option Explicit

Sub Get_Data()

    Dim wb1 As Workbook
    Dim wb2 As Workbook
    Dim network As String
    Dim file As String
    Dim last_row_pivot As Variant
    Dim mben_row_start As Integer
    Dim mben_row_end As Integer
    Dim dest As String
    Dim filename As String
    
    Application.DisplayAlerts = False

    'Add a backslash to the end of idir if needed
    network = [idir]
    If Right(network, 1) <> "\" Then
        network = network & "\"
    End If
    file = [ifile]
    
    'Clear previous month data - Considering truncating all rows down from A2
    Set wb1 = ThisWorkbook
    last_row_pivot = wb1.Worksheets("Original Data").Range("A" & Rows.Count).End(xlUp).Row
    wb1.Worksheets("Original Data").Range("A3:W" & last_row_pivot).ClearContents
    
    'Processing non-M Ben entries - Select non Mben entries and paste as value to [Original Data] tab
    Workbooks.Open filename:=network & file
    Set wb2 = Workbooks(file)
    wb2.Sheets(1).Range("A1").AutoFilter
    wb2.Sheets(1).Range("A1:Y" & 9999).AutoFilter Field:=1, Criteria1:="<>M Benefit Solutions", Operator:=xlFilterValues
    wb2.Sheets(1).Range("A2:W9999").SpecialCells(xlCellTypeVisible).Copy
    wb1.Worksheets("Original Data").Range("A3").PasteSpecial Paste:=xlPasteValues
    
    'Processing m-ben entries - Select Mben entries and paste as value to [Original Data] tab
    mben_row_start = wb1.Worksheets("Original Data").Range("A" & Rows.Count).End(xlUp).Row + 1
    wb2.Sheets(1).Range("A1:Y" & 9999).AutoFilter Field:=1, Criteria1:="M Benefit Solutions", Operator:=xlFilterValues
    wb2.Sheets(1).Range("A2:W9999").SpecialCells(xlCellTypeVisible).Copy
    wb1.Worksheets("Original Data").Range("A" & mben_row_start).PasteSpecial Paste:=xlPasteValues
    mben_row_end = wb1.Worksheets("Original Data").Range("A" & Rows.Count).End(xlUp).Row
    wb2.Close savechanges:=False

    'Fill the color and formulas for the green area(Non-Mben entries)
    wb1.Sheets("Original Data").Activate
    Range("X3").Formula = "=SUMIF(B:B,B3,O:O)"
    Range("Y3").Formula = "=O3/X3"
    Range("Z3").Formula = "=VLOOKUP(B3,$B$" & mben_row_start & ":$Z$" & mben_row_end & ",25,FALSE)"
    Range("AA3").Formula = "=IF(Z3=1,X3,O3)"
    Range("X3:Y" & mben_row_end).FillDown
    Range("Z3:AA" & mben_row_start - 1).FillDown
    Range("A3:AA" & mben_row_start - 1).Interior.Color = 14348258
    
    'Fill the color and formulas for the blue area(Mben entries)
    Range("Z" & mben_row_start).Formula = "=COUNTIF($B$3:$B$" & mben_row_start - 1 & ",B" & mben_row_start & ")"
    'Range("AA" & mben_row_start).Formula = "=IF(OR(Z" & mben_row_start & "=0,Z" & mben_row_start & "=2,Z" & mben_row_start & "=4),O" & mben_row_start & ",0)"
    Range("AA" & mben_row_start).Formula = _
        "=IF(OR(Z" & mben_row_start & "=0,COUNTIFS($A$3:$A$" & mben_row_end & "," & Chr(34) & "*McMillan, Russell*" & Chr(34) & ",$B$3:$B$" & mben_row_end & ",$B" & mben_row_start & " )>0),O" & mben_row_start & ",0)"
    Range("Z" & mben_row_start & ":AA" & mben_row_end).FillDown
    Range("A" & mben_row_start & ":AA" & mben_row_end).Interior.Color = 16247773
    
    'Replace producer 's name with member firm
    Range("AB3").Formula = "=IFERROR(VLOOKUP($A3,ProducerTable,2,0),"""")"
    Range("AB3:AB" & mben_row_start - 1).FillDown
    Dim i As Long
    For i = 3 To mben_row_start - 1
        If Range("AB" & i).Value <> "" Then
            Range("A" & i).Value = Range("AB" & i)
        End If
    Next i
    Range("AB3", Range("AB3").End(xlDown)).ClearContents
    Range("A1").Select
    
    'Refresh Pivot Table
    Call RefreshPivot
    wb1.Save
    Application.DisplayAlerts = True
    
End Sub
   
   
Sub RefreshPivot()
    Sheets("Pivot Table").Select
    Range("D11").Select
    ActiveSheet.PivotTables("PivotTable1").PivotCache.Refresh
End Sub

       
Sub Export()

    Dim wb1 As Workbook: Set wb1 = ThisWorkbook
    Dim fdir As String: fdir = [dest]
    Dim Name As String: Name = [filename]
    Dim PrintRange As Range: Set PrintRange = Sheets("Pivot Table").Range("A1:J36")
    Application.DisplayAlerts = False
    wb1.SaveAs filename:=backslash(fdir) & Name, FileFormat:=xlWorkbookDefault, CreateBackup:=False
    PrintRange.ExportAsFixedFormat Type:=xlTypePDF, filename:=backslash(fdir) & Name & ".pdf"
    Application.DisplayAlerts = True

End Sub

Function backslash(dir As String)
    If Right(dir, 1) <> "\" Then
        dir = dir & "\"
    End If
    backslash = dir
End Function

Sub test()

    Dim mben_row_start As Integer: mben_row_start = 629
    Dim mben_row_end As Integer: mben_row_end = 1005
    ActiveCell.Formula = "=IF(OR(Z" & mben_row_start & "=0,COUNTIFS($A$3:$A$" & mben_row_end & "," & Chr(34) & "*McMillan, Russell*" & Chr(34) & ",$B$3:$B$" & mben_row_end & ",$B" & mben_row_start & " )>0),O" & mben_row_start & ",0)"
    '=IF(OR(Z629=0,COUNTIFS($A$3:$A$1005,"*McMillan, Russell*",$B$3:$B$1005,$B629)>0),O629,0)
End Sub


'Sub test()
'
'    Dim ProducersArray() As Variant
'    Dim AgencyName() As Variant
'    Dim counter As Long, AgencyNameLength As Long 'Counter for array Agencyname
'
'    Sheets("Settings").Activate
'    ProducersArray = Range("L6", Range("L6").End(xlDown))
'    Sheets("Original Data").Activate
'    AgencyNameLength = Range("A3", Range("A3").End(xlDown)).Rows.Count
'    ReDim AgencyName(1 To AgencyNameLength, 1 To 2)
'    For counter = 1 To AgencyNameLength
'        AgencyName(counter, 1) = Range("A2").Offset(counter, 0).Value
'        If IsInArray(AgencyName(counter, 1), ProducersArray) = True Then
'            AgencyName(counter, 2) = Application.WorksheetFunction.VLookup(AgencyName(counter, 1), Sheets("Settings").Range("L5", Range("L5").End(xlToRight).End(xlDown)), 2, 0)
'        Else
'            AgencyName(counter, 2) = AgencyName(counter, 1)
'        End If
'    Next counter
'    Range("A2", Range("A2").Offset(AgencyNameLength, 0)).Value = AgencyName
'
'End Sub

'Function IsInArray(StringToBeFound As String, arr As Variant) As Boolean
'    Dim i
'    For i = LBound(arr) To UBound(arr)
'        If arr(i) = StringToBeFound Then
'            IsInArray = True
'            Exit Function
'        End If
'    Next i
'    IsInArray = False
'End Function
'
'Function ElementOrder(StringToBeFound As String, arr As Variant) As Long
'    Dim i
'    For i = LBound(arr) To UBound(arr)
'        If arr(i) = StringToBeFound Then
'            ElementOrder = i
'            Exit Function
'        End If
'    Next i
'End Function







