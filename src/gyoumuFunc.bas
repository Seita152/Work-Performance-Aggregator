Attribute VB_Name = "gyoumuFunc"
Option Explicit

'業務テーブルの値取得
Function getgyoumu() As String()
    Dim i As Long
    Dim arr(79) As String
    
    For i = 3 To 82
        arr(i - 3) = Workbooks("2024年度稼働実績集計ツール.xlsm").Worksheets("テーブル").Range("E" & i).Value
    Next i
    getgyoumu = arr
End Function

'業務テーブルの値更新
Sub AddGyoumu()
    Dim arr As Variant
    Dim i As Long, j As Long
    Dim ExcelApp As New Application
    Dim Wb As Workbook
    Dim ReadFolderFullPath  As String
    Dim filePath() As String
    Dim names() As String
    
    
    arr = getgyoumu
    names = getname
    
    ExcelApp.Visible = False
    ExcelApp.DisplayAlerts = False
    
    filePath = getFilePath
    
    '個人ブックの業務テーブル更新
    For i = 1 To getname_length
        ReadFolderFullPath = filePath(i - 1)
        Set Wb = ExcelApp.Workbooks.Open(ReadFolderFullPath, , False)
        For j = 0 To 79
            Wb.Worksheets("テーブル").Range("E" & j + 3).Value = arr(j)
        Next j
        Wb.Save
        Wb.Close
        DoEvents
    Next i
    
    '原本ブックの業務テーブルの値更新
    ReadFolderFullPath = ThisWorkbook.Path & "\" & "2024年度日報（原本).xlsx"
    Set Wb = ExcelApp.Workbooks.Open(ReadFolderFullPath, , False)
    For j = 0 To 79
        Wb.Worksheets("テーブル").Range("E" & j + 3).Value = arr(j)
    Next j
    Wb.Save
    Wb.Close
    
    '個人ブックの担当者と合計と原本のシートのセル更新
    For i = 1 To getname_length
        For j = 0 To 79
            Workbooks("2024年度稼働実績集計ツール.xlsm").Worksheets(names(i - 1)).Range("A" & 13 + j).Value = arr(j)
        Next j
    Next i
    
    For j = 0 To 79
            Workbooks("2024年度稼働実績集計ツール.xlsm").Worksheets("合計").Range("A" & 13 + j).Value = arr(j)
    Next j
    
    For j = 0 To 79
            Workbooks("2024年度稼働実績集計ツール.xlsm").Worksheets("原本").Range("A" & 13 + j).Value = arr(j)
    Next j
    
    
    ExcelApp.DisplayAlerts = True
    ExcelApp.Quit
    DoEvents
    Set ExcelApp = Nothing
    Debug.Print "完了"
End Sub