Attribute VB_Name = "memberfunc"
Option Explicit

'担当者名取得
Public Function getname() As String()
    Dim i As Long, j As Long, elemc As Long, name As String
    elemc = getname_length
    Dim arr() As String
    ReDim arr(getname_length) As String
    i = 3
    j = 0
    'Worksheets("テーブル").cell(3, 7).End(xlDown).Select
    Do Until Worksheets("テーブル").Range("G" & i).Value = ""
        name = Worksheets("テーブル").Range("G" & i).Value
        arr(j) = name
        j = j + 1
        'Debug.Print name
        i = i + 1
    Loop
    getname = arr()
End Function

'担当者の数取得
Public Function getname_length() As Long
    Dim elemc As Long
    elemc = 0
    elemc = Worksheets("テーブル").Range("G3").End(xlDown).Row
    elemc = elemc - 2
    getname_length = elemc
End Function

'所属名取得
Public Function getTeamName() As String
    Dim startRange As Range
    Dim endRow As Double
    Dim endRange As Range
    Dim address As String
    Dim arrAddress As Variant
    Dim targetColumnAlphabet As String
    Dim wf As WorksheetFunction
    Dim arrData As Variant
    Dim data As Variant
    Dim arr() As String
    Dim i As Long
    
    Set startRange = Worksheets("テーブル").Range("H3")
    
    endRow = startRange.End(xlDown).Row
    
    Set endRange = Worksheets("テーブル").Cells(endRow, startRange.Row)
    
    address = startRange.address(RowAbsolute:=True, ColumnAbsolute:=False)
    arrAddress = Split(address, "$")
    targetColumnAlphabet = arrAddress(0)

    Set wf = Application.WorksheetFunction
    
    arrData = wf.Transpose(Range(targetColumnAlphabet & startRange.Row & ":" & targetColumnAlphabet & endRange.Row))
    i = 0
    ReDim Preserve arr(100)
    For Each data In arrData
        arr(i) = data
        i = i + 1
    Next
    ReDim Preserve arr(i)
    getTeamName = arr
End Function

'所属名の数取得
Public Function getTeamName_length() As Long
    Dim elemc As Long
    elemc = 0
    elemc = Worksheets("テーブル").Range("H3").End(xlDown).Row
    elemc = elemc - 2
    getTeamName_length = elemc
End Function
