Attribute VB_Name = "Module1"
Option Explicit

'Need to have Tools -> "Microsoft Scripting Runtime" library checked
'Using autoInstancingVaraible
Sub autoInstancingVaraible()
    Dim fso As New Scripting.FileSystemObject
End Sub
    
'Does not need to have Tools -> "Microsoft Scripting Runtime" library checked
'But there is no intellisense
Sub UsingCreateObject()
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
End Sub

'Need to have Tools -> "Microsoft Scripting Runtime" library checked
Sub UsingTheScriptingRunTimeLibraryv1()
    'Creating a fileSystem Object
    Dim fso As Scripting.FileSystemObject
    Dim fil As Scripting.File
    Dim newFolderPath As String
    Dim oldFolderPath As String

    'User Profile
    newFolderPath = Environ("UserProfile") & "\Desktop\Python\VBA Tutorial\Wise Owls"
    oldFolderPath = "J:\Systems\Production & Override\Production\Premium Data\Current Month"
    Set fso = New Scripting.FileSystemObject

    'Copy to new dir if file exists
    If fso.FileExists(oldFolderPath & "\M Financial - March Sales Report.xlsx") Then
'        fso.CopyFile _
'            Source:=oldFolderPath & "\M Financial - May Sales Report.xlsx" _
'            , Destination:=newFolderPath & "\M Financial - May Sales Report.xlsx" _
'            , OverWriteFiles:=True
        Set fil = fso.GetFile(oldFolderPath & "\M Financial - March Sales Report.xlsx")
        If fil.Size > 20000 Then
            fil.Copy newFolderPath & "\" & fil.Name
        End If
    End If
    'This is a great practice to reset the variable and release the memory
    Set fso = Nothing
End Sub

'Looping over all files in a folder
Sub UsingTheScriptingRunTimeLibraryv2()

    'Creating a fileSystem Object
    Dim fso As Scripting.FileSystemObject
    Dim fil As Scripting.File
    Dim oldfolder As Scripting.Folder
    Dim newFolderPath As String
    Dim oldFolderPath As String

    'User Profile
    newFolderPath = Environ("UserProfile") & "\Desktop\Python\VBA Tutorial\Wise Owls"
    oldFolderPath = "J:\Systems\Production & Override\Production\Premium Data\Current Month"
    Set fso = New Scripting.FileSystemObject

    If fso.FolderExists(oldFolderPath) Then
        Set oldfolder = fso.GetFolder(oldFolderPath)
        'Test if the file has existed
        If Not fso.FolderExists(newFolderPath) Then
            fso.CreateFolder newFolderPath
        End If
        For Each fil In oldfolder.Files
            If Left(fso.GetExtensionName(fil.Path), 2) = "xl" Then
                fil.Copy newFolderPath & "\" & fil.Name
            End If
        Next fil
    End If
    Set fso = Nothing
End Sub

'Looping through all files and subfolders in a folder

Dim fso As Scripting.FileSystemObject
Dim newFolderPath As String

Sub UsingTheScriptingRunTimeLibraryv3()
    
    Dim oldFolderPath As String
    
    newFolderPath = Environ("UserProfile") & "\Desktop\Python\VBA Tutorial\Wise Owls"
    oldFolderPath = "J:\Systems\Production & Override\Production\Premium Data\Current Month"
    
    Set fso = New Scripting.FileSystemObject
    
    If fso.FolderExists(oldFolderPath) Then
        If Not fso.FolderExists(newFolderPath) Then
            fso.CreateFolder newFolderPath
        End If
        Call copyExcelFiles(oldFolderPath)
    End If
        
    Set fso = Nothing
    
End Sub

Sub copyExcelFiles(startFolderPath As String)
    
    Dim fil As Scripting.File
    Dim oldfolder As Scripting.Folder
    Dim subdir As Scripting.Folder
    
    Set oldfolder = fso.GetFolder(startFolderPath)
    
    For Each fil In oldfolder.Files
        If Left(fso.GetExtensionName(fil.Path), 2) = "xl" Then
            fil.Copy newFolderPath & "\" & fil.Name
        End If
    Next fil
    
    'Recursive Programming - a subroutine call itself
    For Each subdir In oldfolder.SubFolders
        Call copyExcelFiles(subdir.Path)
    Next subdir

End Sub
















