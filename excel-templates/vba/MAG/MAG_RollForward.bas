' **************
' Process
' **************

sub Main(pmWb as workbook)
	
    Dim fso As Scripting.FileSystemObject
    Dim NewQdir As String: NewQdir = quarterPath
    Dim OldQdir As String
    Dim BlankPath As String: BlankPath = "J:\MLife\Magnastar\Scripts\Blanks\"
    Dim FileName As String
    Set fso = New Scripting.FileSystemObject
    
	
	'Make New Directory for MAG files if not existed
    If Not fso.FolderExists(NewQdir) Then
        fso.CreateFolder NewQdir
    End If
    If Not fso.FolderExists(NewQdir & "Data") Then
        fso.CreateFolder NewQdir & "Data"
    End If
    If Not fso.FolderExists(NewQdir & "Data\MAG") Then
        fso.CreateFolder NewQdir & "Data\MAG"
    End If
    
	
	'Copy over all files in Data\MAG and rename with new year and quarter
	If pm.quarter = 1 Then
	
	
		'For Q1, take new blank settlement sheets from "J:\MLife\Magnastar\Scripts\Blanks\"
		FileName = Dir(BlankPath & "*")		
		Do While Len(FileName) >0
            FileCopy Source:=BlankPath & FileName, Destination:=NewQdir & "Data\MAG\" & pm.year & FileName
            FileName = Dir
		Loop
        
		'Roll over DAC Tax Workbooks from prior year Q4
        OldQdir = Replace(Replace(pm.quarterPath, pm.year, pm.year - 1), "Q" & pm.quarter, "Q4") & "Data\MAG\"
        FileName = Dir(OldQdir & "*DAC Tax.xlsx*")
        Do While Len(FileName) > 0
            FileCopy Source:=OldQdir & FileName, Destination:=NewQdir & "Data\MAG\" & Replace(Replace(FileName, pm.year - 1, pm.year), "Q4", "Q1")
            FileName = Dir
        Loop
	
	Else
		
		'For Q2, Q3 and Q4, roll over everything with .xlsx file extension
		OldQdir = Replace(pm.QuarterPath, "Q" & pm.quarter, "Q" &pm.quarter - 1) & "Data\MAG\"
		FileName = Dir(OldQdir & "*xlsx*")
		Do While Len(FileName) > 0
            FileCopy Source:= OldQdir & FileName, Destination:=NewQdir & "Data\MAG\" & Replace(FileName, "Q" & pm.quarter - 1, "Q" & pm.quarter)
			FileName = Dir
		Loop
		
		'Reset settlement workbooks for each carrier
		'SLD
		Call ResetSettlement("SLD")
		
		'NYL
		Call ResetSettlement("NYL")
		
		'JH - need to also update Magnastar Fees column reference in CWS
		Call ResetSettlement("JH")
		Dim path As String: path = pm.quarterPath & "Data\MAG\" & pm.year & "Q" & pm.quarter & " Magnastar Settlement JH.xlsx"
		Dim wb As Workbook: Set wb = PMUtil.GetWorkbook(path)
		Dim col as String
			If pm.quarter = 2 then
				col = "E"
			ElseIf pm.quarter = 3 then
				col = "H"
			ElseIf pm.quarter = 4 then 
				col = "K"
			End If
		wb.Sheets("CWS").Range("E34").Formula = "=-SUM('Net Magnastar Fees'!" & col & "3:" & col & "6)"
		wb.Sheets("CWS").Range("E35").Formula = "=SUM('Net Magnastar Fees'!" & col & "21:" & col & "24)"	
		wb.Close SaveChanges:=True
		
		'PL
		Call ResetSummaryPL("PLA")
		Call ResetSettlement("PLA")
		Call ResetSummaryPL("PLN")
		Call ResetSettlement("PLN")

		'PRU
		Call ResetSummaryPRU("PR1")
		Call ResetSettlement("PR1")
		Call ResetSettlement("PR1 - M")
		Call ResetSettlement("PR1 - Swiss")

	End If
	
	Set fso = Nothing

End Sub


'Below are subroutines that reset summary and Qq settlement tabs, edited in advance and call in main subroutine
'##########################################################################################################################

Sub ResetSettlement(Carrier As String) 
    Dim path As String: path = pm.quarterPath & "Data\MAG\" & pm.year & "Q" & pm.quarter & " Magnastar Settlement " & Carrier & ".xlsx"
    Dim wb As Workbook: Set wb = PMUtil.GetWorkbook(path)
	Dim ws As Worksheet
	application.Calculation = xlCalculationAutomatic
	'Roll over the Qq Settlement sheet, reset Net Amount Due to 0
	For i = 1 to 3
		Set ws = wb.Worksheets("Q" & i & " Settlement")		
		ws.Activate
		Range("I6").Formula = "=F6+G6"
		Range("I6:I80").FillDown	
		Range("I6:I80").Copy
		Range("F6:F80").PasteSpecial Paste:=xlPasteValues
		Range("I6:I80").ClearContents
	Next i
	wb.Worksheets("YTD Settlement").Range("A3").Value = quarter
	wb.Worksheets("Q" & pm.quarter & " Settlement").Visible = True
	wb.Worksheets("Q" & pm.quarter & " Database").Visible = True
	application.Calculation = xlCalculationManual
	wb.Close SaveChanges:=True
End Sub


'Summary tab - PL only
Sub ResetSummaryPL(Carrier As String)
    Dim path As String: path = pm.quarterPath & "Data\MAG\" & pm.year & "Q" & pm.quarter & " Magnastar Settlement " & Carrier & ".xlsx"
    Dim wb As Workbook: Set wb = PMUtil.GetWorkbook(path)
    Dim ws As Worksheet: Set ws = wb.Worksheets("Summary")
    'Roll over the Qq Settlement sheet, reset Net Amount Due to 0
    ws.Activate
    application.Calculation = xlCalculationAutomatic
    Range("AA6").Formula = "=C6+I6"
	Range("AA6").Copy
    Range("AA6:AE74").PasteSpecial Paste:=xlPasteFormulas
	Range("AA6:AE74").Copy
    Range("O6:S74").PasteSpecial Paste:=xlPasteValues
    Range("AA6:AE74").ClearContents
    application.Calculation = xlCalculationManual
    wb.Close SaveChanges:=True
End Sub


'Summary tab - PRU only
Sub ResetSummaryPRU(Carrier As String)  
    Dim path As String: path = pm.quarterPath & "Data\MAG\" & pm.year & "Q" & pm.quarter & " Magnastar Settlement " & Carrier & ".xlsx"
    Dim wb As Workbook: Set wb = PMUtil.GetWorkbook(path)
    Dim ws As Worksheet: Set ws = wb.Worksheets("Summary")
    'Roll over the Qq Settlement sheet, reset Net Amount Due to 0
    ws.Activate
    application.Calculation = xlCalculationAutomatic
    Range("AA6").Formula = "=C6+H6"
	Range("AA6").Copy
    Range("AA6:AD77").PasteSpecial Paste:=xlPasteFormulas
	Range("AA6:AD77").Copy
    Range("M6:P77").PasteSpecial Paste:=xlPasteValues
    Range("AA6:AD77").ClearContents
    application.Calculation = xlCalculationManual
    wb.Close SaveChanges:=True
End Sub

'##########################################################################################################################



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