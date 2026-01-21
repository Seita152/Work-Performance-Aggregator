Attribute VB_Name = "nippouFunc"
'日報書き込み
Public Function PutNippou(days As String, filename As String, names As String) As Boolean
    Dim i As Long, j As Long, Row As Long, col As Long, k As Long, q As Long
    Dim kinmum As Long, kinmut As Long, kinmua As Long, comment As String, bodynum As String, busy As String, kinmuave As Long
    Dim searchRange As Range
    Dim searchObj As Range
    Dim keyWord As String
    Dim daysA As String, TEMP As String
    Dim ws As Worksheet, flag As Boolean

    Dim ExcelApp As New Application
    Dim Wb As Workbook
    Dim ReadFolderFullPath  As String
    Dim nippou_start As Long
    
    
    j = 0
    kinmum = 0
    kinmut = 0
    kinmua = 0
    kinmuave = 0
    nippou_start = 5

    comment = ""
    bodynum = ""
    busy = ""
    
    Row = 2
    col = 0
    
    ExcelApp.Visible = False
    ExcelApp.DisplayAlerts = False
    
    ReadFolderFullPath = filename
    Set Wb = ExcelApp.Workbooks.Open(ReadFolderFullPath, , True)
    
    '日付の列を探索
    Set myRange = Workbooks("2024年度稼働実績集計ツール.xlsm").Worksheets(names).Range("B1:JB1")
    keyWord = days
    Set myObj = myRange.Find(DateValue(keyWord), LookAt:=xlWhole, LookIn:=xlFormulas)
    col = myObj.Column
    
    '目的シート探索
    daysA = Format(days, "yyyymmdd")
    For Each ws In Wb.Worksheets
        If ws.name = daysA Then flag = True
    Next ws
    
    'シートがなかったら
    If flag = False Then
        MsgBox daysA & "のシートがありません"
        ExcelApp.DisplayAlerts = True
        ExcelApp.Quit
        Set ExcelApp = Nothing
        PutNippou = False
        Exit Function
    End If
    
    '未記入だったら
    If IsEmpty(Wb.Worksheets(daysA).Cells(nippou_start, 12).Value) Then
        PutNippou = False
        Exit Function
    End If
    
    '書き込み開始
    For q = 0 To 4
    
        '体調について
        TEMP = Call_ColConv(col + q)
        Workbooks("2024年度稼働実績集計ツール.xlsm").Worksheets(names).Range(TEMP & "2:" & TEMP & "93").ClearContents
        
        kinmum = Wb.Worksheets(daysA).Cells(nippou_start + j, 3).Value
        Workbooks("2024年度稼働実績集計ツール.xlsm").Worksheets(names).Cells(2, col + q).Value = kinmum
        j = j + 1
        
        kinmut = Wb.Worksheets(daysA).Cells(nippou_start + j, 3).Value
        Workbooks("2024年度稼働実績集計ツール.xlsm").Worksheets(names).Cells(3, col + q).Value = kinmut
        j = j + 1
        
        kinmua = Wb.Worksheets(daysA).Cells(nippou_start + j, 3).Value
        Workbooks("2024年度稼働実績集計ツール.xlsm").Worksheets(names).Cells(4, col + q).Value = kinmua
        j = j + 10
        
        kinmuave = (kinmum + kinmut + kinmua) / 3
        Workbooks("2024年度稼働実績集計ツール.xlsm").Worksheets(names).Cells(5, col + q).Value = kinmuave
        
        comment = Wb.Worksheets(daysA).Cells(nippou_start + j, 3).MergeArea.Item(1).Value
        Workbooks("2024年度稼働実績集計ツール.xlsm").Worksheets(names).Cells(6, col + q).Value = comment
        j = j + 1
        

        bodynum = Wb.Worksheets(daysA).Cells(nippou_start + j, 3).Value
        Workbooks("2024年度稼働実績集計ツール.xlsm").Worksheets(names).Cells(7, col + q).Value = bodynum
        j = j + 1
        
        busy = Wb.Worksheets(daysA).Cells(nippou_start + j, 3).Value
        Workbooks("2024年度稼働実績集計ツール.xlsm").Worksheets(names).Cells(8, col + q).Value = busy
        
        '作業時間書き込み
        For i = 0 To 13
            If Wb.Worksheets(daysA).Cells(nippou_start + i, 12).Value <> "" Then
                For k = 0 To 79
                     If Workbooks("2024年度稼働実績集計ツール.xlsm").Worksheets(names).Cells(13 + k, 1).Value = Wb.Worksheets(daysA).Cells(nippou_start + i, 12).Value Then
                        If Workbooks("2024年度稼働実績集計ツール.xlsm").Worksheets(names).Cells(13 + k, col + q).Value = "" Then
                            Workbooks("2024年度稼働実績集計ツール.xlsm").Worksheets(names).Cells(13 + k, col + q).Value = Wb.Worksheets(daysA).Cells(nippou_start + i, 11).Value
                            Workbooks("2024年度稼働実績集計ツール.xlsm").Worksheets(names).Cells(13 + k, col + q).NumberFormatLocal = "[h]:mm"
                        Else
                            Workbooks("2024年度稼働実績集計ツール.xlsm").Worksheets(names).Cells(13 + k, col + q).Value = Workbooks("2024年度稼働実績集計ツール.xlsm").Worksheets(names).Cells(13 + k, col + q).Value + Wb.Worksheets(daysA).Cells(nippou_start + i, 11).Value
                            Workbooks("2024年度稼働実績集計ツール.xlsm").Worksheets(names).Cells(13 + k, col + q).NumberFormatLocal = "[h]:mm"
                        End If
                        Exit For
                     End If
                Next k
            End If
        Next i
        
        nippou_start = nippou_start + 15
        j = 0
        kinmum = 0
        kinmut = 0
        kinmua = 0
        kinmuave = 0
        comment = ""
        bodynum = ""
        busy = ""
        Workbooks("2024年度稼働実績集計ツール.xlsm").Worksheets(names).Cells(12, col + q).NumberFormatLocal = "[h]:mm"
        Workbooks("2024年度稼働実績集計ツール.xlsm").Worksheets(names).Cells(12, col + q).Formula = "=SUM(" & TEMP & "13:" & TEMP & "92)"
        '次の個人ブックへ
    Next q
    
    ExcelApp.DisplayAlerts = True
    ExcelApp.Quit
    Set ExcelApp = Nothing
    PutNippou = True
End Function

'行番号をアルファベットに変換
Public Function Call_ColConv(var As Variant) As Variant
    Dim TEMP As String
     
    If IsNumeric(var) = True Then
      '■数値の場合はAddress取得し、アルファベットのみを返す
      TEMP = Cells(1, CLng(var)).address(RowAbsolute:=False, ColumnAbsolute:=False)
      Call_ColConv = Left(TEMP, Len(TEMP) - 1)
    Else
      '■アルファベットだったら行番号を返す
      Call_ColConv = Range(var & "1").Column
    End If
End Function

'シートが存在してるか判断するための関数
Function IsExistsSheet(sheetName) As Boolean
    Dim b As Boolean
    b = False
    
    Dim sh
    For Each sh In Worksheets
        If sh.name = sheetName Then
            b = True
            Exit For
        End If
    Next sh
    
    IsExistsSheet = b
End Function

'最新時刻シートの追加
Sub AddNippouSheet()
    Dim ExcelApp As New Application
    Dim Wb As Workbook
    Dim filePath() As String
    Dim ReadFolderFullPath  As String
    Dim i, j As Long
    Dim BsheetName, AsheetName As String
    Dim Ddata As Variant
    Dim names() As String
    Dim dt  As Date
    Dim startRow As Long
    
    startRow = 5
    
    ReDim names(getname_length - 1) As String
    names = getname
    
    ExcelApp.Visible = False
    ExcelApp.DisplayAlerts = False
    
    '001の最新時刻シート名取得
    filePath = getFilePath
    ReadFolderFullPath = filePath(0)
    Set Wb = ExcelApp.Workbooks.Open(ReadFolderFullPath, , True)
    BsheetName = Wb.Worksheets(Wb.Sheets.Count).name
    Ddata = CDate(Format(BsheetName, "@@@@/@@/@@"))
    dt = DateAdd("d", 7, Ddata)
    AsheetName = Format(dt, "yyyymmdd")
    Wb.Save
    Wb.Close
    DoEvents
    
    '最新時刻シート生成
    For i = 1 To getname_length
        ReadFolderFullPath = filePath(i - 1)
        Set Wb = ExcelApp.Workbooks.Open(ReadFolderFullPath, , False)
        Wb.Worksheets("原本").Copy After:=Wb.Worksheets(Wb.Worksheets.Count)
        Wb.Worksheets(Wb.Worksheets.Count).name = AsheetName
        Wb.Worksheets(AsheetName).Range("M1").Value = WorksheetFunction.Replace(names(i - 1), 4, 1, " ")
        For j = 0 To 4
            Wb.Worksheets(AsheetName).Range("A" & (startRow + 15 * j)).Value = DateAdd("d", j, dt)
        Next j
        Wb.Save
        Wb.Close
        DoEvents
    Next i

    ExcelApp.DisplayAlerts = True
    ExcelApp.Quit
    Set ExcelApp = Nothing
End Sub

'合計シートの値入力
Sub NippouSum(days As String)
    Dim i As Long, col As Long, k As Long, q As Long
    Dim keyWord As String
    Dim names() As String
    
    Set myRange = Workbooks("2024年度稼働実績集計ツール.xlsm").Worksheets("合計").Range("B1:JB1")
    keyWord = days
    Set myObj = myRange.Find(DateValue(keyWord), LookAt:=xlWhole, LookIn:=xlFormulas)
    col = myObj.Column
    
    Workbooks("2024年度稼働実績集計ツール.xlsm").Worksheets("合計").Range(Call_ColConv(col) & "2:" & Call_ColConv(col + 4) & "92").ClearContents
    names = getname
    For q = 0 To 4
        For i = 1 To getname_length
            For k = 0 To 79
                If Workbooks("2024年度稼働実績集計ツール.xlsm").Worksheets("合計").Cells(13 + k, col + q).Value = "" And Workbooks("2024年度稼働実績集計ツール.xlsm").Worksheets(names(i - 1)).Cells(13 + k, col + q) <> "" Then
                    Workbooks("2024年度稼働実績集計ツール.xlsm").Worksheets("合計").Cells(13 + k, col + q).Value = Workbooks("2024年度稼働実績集計ツール.xlsm").Worksheets(names(i - 1)).Cells(13 + k, col + q).Value
                ElseIf Workbooks("2024年度稼働実績集計ツール.xlsm").Worksheets("合計").Cells(13 + k, col + q).Value <> "" And Workbooks("2024年度稼働実績集計ツール.xlsm").Worksheets(names(i - 1)).Cells(13 + k, col + q) <> "" Then
                    Workbooks("2024年度稼働実績集計ツール.xlsm").Worksheets("合計").Cells(13 + k, col + q).Value = Workbooks("2024年度稼働実績集計ツール.xlsm").Worksheets("合計").Cells(13 + k, col + q).Value + Workbooks("2024年度稼働実績集計ツール.xlsm").Worksheets(names(i - 1)).Cells(13 + k, col + q).Value
                End If
            Next k
        Next i
        Workbooks("2024年度稼働実績集計ツール.xlsm").Worksheets("合計").Cells(12, col + q).Formula = "=SUM(" & Call_ColConv(col + q) & "13:" & Call_ColConv(col + q) & "92)"
        Workbooks("2024年度稼働実績集計ツール.xlsm").Worksheets("合計").Range(Call_ColConv(col + q) & "13", Call_ColConv(col + q) & "92").NumberFormatLocal = "[h]:mm"
        Workbooks("2024年度稼働実績集計ツール.xlsm").Worksheets("合計").Cells(12, col + q).NumberFormatLocal = "[h]:mm"
    Next q
End Sub

