Attribute VB_Name = "fileFunc"
Option Explicit

'個人のブックのパス取得
Public Function getFilePath() As String()
    Dim myfilename As String, i As Integer, names() As String, arr() As String, loopCnt As Long, j As Long
    Dim subfolders() As String
    Dim FILE_PATH As String
    ReDim subfolders(getname_length - 1) As String
    ReDim names(getname_length - 1) As String
    Dim fso As Object
    Dim TARGET As Object
    Dim TEMP As Object
    Dim myFile As Object
    Dim myFolder As Object
    
    FILE_PATH = ThisWorkbook.Path
    names = getname
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set TARGET = fso.GetFolder(FILE_PATH).subfolders
    
    i = 0
    For Each TEMP In TARGET
        If IsNumeric(Left(TEMP.name, 3)) Then
            subfolders(i) = TEMP.name
            i = i + 1
        End If
    Next
    
    i = 0
    For j = 0 To UBound(subfolders)
        Set myFolder = fso.GetFolder(ThisWorkbook.Path & "\" & subfolders(j))
        For Each myFile In myFolder.Files
            If i = 0 Then
                i = 1
                ReDim Preserve arr(i)
                arr(0) = ThisWorkbook.Path & "\" & subfolders(j) & "\" & fso.GetFileName(myFile.Path)
            Else
                arr(i) = ThisWorkbook.Path & "\" & subfolders(j) & "\" & fso.GetFileName(myFile.Path)
                i = i + 1
                ReDim Preserve arr(i)
            End If
            
        Next myFile
    Next j
    
    getFilePath = arr()
    
End Function


