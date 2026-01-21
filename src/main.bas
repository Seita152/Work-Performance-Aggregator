Attribute VB_Name = "main"
Sub NippouButton()
    Dim ans As String
    Dim filePath() As String, i As Long, names() As String, elemMsg As String
    Dim checkResult As Boolean
    
    'インプットボックス表示
    ans = InputBox("日報転記・集計したい週の月曜日の日付をyyyy/mm/ddの形式で入力し" & vbLf & "OKボタンを押してください", "日報転記・集計")
    elemMsg = "処理が完了しました"
    
    '入力チェック
    If StrPtr(ans) = 0 Then
        Exit Sub
    End If
    
    If ans = "" Then
        Exit Sub
    End If
    
    ans = Format(ans, "yyyy/m/d")
    
    
    If IsDate(ans) = False Then
        MsgBox "日付ではありません"
        Exit Sub
    End If
    
    If year(ans) < 2024 Then
        MsgBox "2024年の日付ではありません"
        Exit Sub
    End If
    
    If year(ans) = 2024 And Month(ans) < 4 Then
        MsgBox "2024年4月以降の日付を入力してください"
        Exit Sub
    End If
    
    If Weekday(ans) <> 2 Then
        MsgBox "月曜日の日付を入力してください"
        Exit Sub
    End If
    
    filePath = getFilePath
    names = getname
    
    '日報書き込み開始
    For i = 0 To UBound(filePath) - 1
        checkResult = PutNippou(ans, filePath(i), names(i))
        If checkResult = False Then
            elemMsg = "処理を中断しました"
            Exit For
        End If
    Next i
    If checkResult <> False Then
        Call NippouSum(ans)
    End If
    Workbooks("2024年度稼働実績集計ツール.xlsm").Save
    
    '書き込み終了
    MsgBox elemMsg
    Debug.Print elemMsg
End Sub

Sub shainButton()

    If MsgBox("担当追加をしてよろしいですか", vbYesNo) = vbYes Then
        UserForm1.Show vbModeless
    Else
        Exit Sub
    End If
     
End Sub

Sub gyoumuButton()

    If MsgBox("業務追加をしてよろしいですか", vbYesNo) = vbYes Then
        Call AddGyoumu
        MsgBox "処理が完了しました"
    Else
        Exit Sub
    End If
     
End Sub

Sub AddNippouButton()

    If MsgBox("日報シート追加をしてよろしいですか", vbYesNo) = vbYes Then
        Call AddNippouSheet
        MsgBox "処理が完了しました"
    Else
        Exit Sub
    End If
     
End Sub