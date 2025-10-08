Option Explicit

Private Sub CommandButton1_Click()
    Call setObj
    Call RegisterTask
    Call releaseObj
End Sub

Private Sub CommandButton2_Click()
    Call setObj
    Call SortByPriority
    Call releaseObj
End Sub

Private Sub CommandButton3_Click()
    Call setObj
    Call SortByDeadline
    Call releaseObj
End Sub

' =================================================================
' プロシージャ名: Worksheet_Change
' 機能: シート上のセルが変更されたときに自動実行される
' =================================================================
Private Sub Worksheet_Change(ByVal Target As Range)

    ' --- 変数の宣言 ---
    Dim archiveRow As Long
    
    ' --- 変更されたセルがE列（進捗列）か、複数セルでないかをチェック ---
    ' Target.Column = 5 はE列を意味します。
    ' Target.Count > 1 は複数セルが同時に変更された場合を弾くための条件です。
    If Target.Column <> col_progress Or Target.Count > 1 Then
        Exit Sub ' E列以外が変更された場合は、ここで処理を終了
    End If
    
    ' --- 見出し行（1行目）が変更された場合は処理しない ---
    If Target.Row = 1 Then
        Exit Sub
    End If

    ' --- 初期設定 ---
    Call setObj

    ' --- E列の値が「完了」に変更された場合の処理 ---
    If Target.Value = "完了" Then
        
        ' 画面の更新を一時的に停止（処理を高速化するため）
        Application.ScreenUpdating = False
        
        ' --- 完了タスクシートの最終行を取得 ---
        archiveRow = wsArchive.Cells(wsArchive.Rows.Count, 1).End(xlUp).Row + 1
        
        ' --- タスク行を完了タスクシートにコピー ---
        wsList.Rows(Target.Row).Copy Destination:=wsArchive.Rows(archiveRow)
        
        ' --- 完了タスクシートのG列に完了日を記録 ---
        wsArchive.Cells(archiveRow, col_complete).Value = Date
        
        ' --- 元のタスクリストから行を削除 ---
        wsList.Rows(Target.Row).Delete
        
        ' 罫線追加
        With wsArchive
        Call setGrid(.Range(.Cells(1, 1), .Cells(archiveRow, 7)))
        End With

        ' 画面の更新を再開
        Application.ScreenUpdating = True
        
        MsgBox "タスクを完了リストへ移動しました。お疲れ様でした！", vbInformation
        
    End If
    
    ' オブジェクト開放
    Call releaseObj
End Sub
