Option Explicit

' =================================================================
' プロシージャ名: Workbook_Open
' 機能: このExcelブックが開かれたときに自動実行される
' =================================================================
Private Sub Workbook_Open()

    ' --- 変数の宣言 ---
    Dim lastRow As Long
    Dim i As Long
    Dim deadline As Date
    Dim taskName As String
    Dim reminderMsg As String
    
    ' --- 初期設定 ---
    Call setObj
    lastRow = wsList.Cells(wsList.Rows.Count, col_name).End(xlUp).Row
    reminderMsg = "【期限が3日以内に迫っているタスクがあります！】" & vbCrLf & vbCrLf
    
    ' --- タスクリストを1行ずつチェックするループ処理 ---
    ' 2行目から最終行まで繰り返します。
    For i = 2 To lastRow
        
        ' --- 期限が入力されていて、かつ進捗が「完了」でないタスクをチェック ---
        If wsList.Cells(i, col_deadline).Value <> "" _
            And wsList.Cells(i, col_progress).Value <> "完了" Then
            
            ' D列（期限）の日付を取得
            deadline = wsList.Cells(i, col_deadline).Value
            
            ' --- 期限が今日から3日以内かを判定 ---
            ' (期限 - 今日の日付) が 3以下かつ0以上の場合
            If (deadline - Date) <= 3 And (deadline - Date) >= 0 Then
                
                ' B列（タスク名）を取得
                taskName = wsList.Cells(i, col_name).Value
                
                ' メッセージ文字列にタスク情報を追加
                reminderMsg = reminderMsg & "・" & taskName _
                & " (期限: " & Format(deadline, "yyyy/mm/dd") & ")" & vbCrLf
                
            End If
        End If
    Next i
    
    ' --- 期限切れのタスクが1件でもあればメッセージボックスで通知 ---
    ' Len関数でメッセージの長さをチェックし、初期文字列より長ければタスクがあったと判断
    If Len(reminderMsg) > Len("【期限が3日以内に迫っているタスクがあります！】" _
        & vbCrLf & vbCrLf) Then
        MsgBox reminderMsg, vbExclamation, "期限アラート"
    End If

    ' オブジェクト開放
    Call releaseObj

End Sub
