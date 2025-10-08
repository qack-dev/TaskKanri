Option Explicit

' グローバル定数
Public Const col_id As Integer = 1
Public Const col_name As Integer = 2
Public Const col_priority As Integer = 3
Public Const col_deadline As Integer = 4
Public Const col_progress As Integer = 5
Public Const col_register As Integer = 6
Public Const col_complete As Integer = 7
' グローバル変数
Public wsList As Worksheet
Public wsArchive As Worksheet

' グローバル変数代入
Public Sub setObj()
    Set wsList = ThisWorkbook.Worksheets("タスクリスト")
    Set wsArchive = ThisWorkbook.Worksheets("完了タスク")
End Sub
' グローバル変数開放
Public Sub releaseObj()
    Set wsList = Nothing
    Set wsArchive = Nothing
End Sub

' =================================================================
' プロシージャ名: RegisterTask
' 機能: 新しいタスクをタスクリストに登録する
' =================================================================
Sub RegisterTask()

    ' --- 変数の宣言 ---
    ' VBAでデータを扱うための「箱」を用意します。
    Dim lastRow As Long
    Dim newRow As Long

    With wsList
        ' --- 最終行の取得 ---
        ' A列の最終行を調べて、次に入力する行を決定します。
        ' Cells(Rows.Count, "A")はA列の最後のセルを指します。
        ' .End(xlUp)で、データが入力されている一番上のセルまで移動します。
        lastRow = .Cells(.Rows.Count, 1).End(xlUp).Row
        newRow = lastRow + 1
        
        ' --- タスク情報の書き込み ---
        ' 新しい行の各セルに、データを書き込んでいきます。
        .Cells(newRow, col_id).Value = makeTaskId     ' タスクID
        .Cells(newRow, col_priority).Value = "中"           ' 優先度（デフォルト値）
        .Cells(newRow, col_progress).Value = "未着手"      ' 進捗（デフォルト値）
        .Cells(newRow, col_register).Value = Date          ' 登録日（今日の日付）
        Call setGrid(.Range(Cells(1, 1), Cells(newRow, col_complete)))
    End With
    ' --- ユーザーへの通知 ---
    MsgBox "新しいタスクを登録しました！" _
        & vbCrLf & "タスク名や期限を入力してください。", vbInformation
    Cells(newRow, col_name).Select
End Sub

' タスクIDを採番
Private Function makeTaskId() As Long
    ' 変数宣言
    Dim tmpId As Long: tmpId = 0
    Dim r As Range
    ' タスクリストシートをloop
    With wsList
        ' タスクIDの最大値を捜索
        For Each r In .Range(.Cells(1, col_id), .Cells(.Rows.Count, col_id).End(xlUp))
            If 1 < r.Row And tmpId < r.Value Then
                tmpId = r.Value
            End If
        Next
    End With
    ' 完了タスクシートをloop
    With wsArchive
        ' タスクIDの最大値を捜索
        For Each r In .Range(.Cells(1, col_id), .Cells(.Rows.Count, col_id).End(xlUp))
            If 1 < r.Row And tmpId < r.Value Then
                tmpId = r.Value
            End If
        Next
    End With
    ' 戻り値
    makeTaskId = tmpId + 1
End Function

' =================================================================
' プロシージャ名: SortByPriority
' 機能: タスクリストを優先度順（降順）で並べ替える
' =================================================================
Sub SortByPriority()

    ' --- 変数の宣言 ---
    Dim sortRange As Range
    Dim lastRow As Long

    ' --- 初期設定 ---
    lastRow = wsList.Cells(wsList.Rows.Count, 1).End(xlUp).Row
    
    ' 見出し行しかない場合は処理を終了
    If lastRow <= 1 Then
        MsgBox "並べ替えるタスクがありません。", vbExclamation
        Exit Sub
    End If

    ' 並べ替えの対象範囲（A1からデータの最終行まで）を設定
    Set sortRange = wsList.Range("A1:G" & lastRow)
    
    ' --- 並べ替えの実行 ---
    With wsList.Sort
        .SortFields.Clear ' 既存の並べ替え条件をクリア
        ' C列（優先度）をキーに設定。ここでは「高,中,低」の順にしたいのでカスタムリストを使う
        .SortFields.Add Key:=wsList.Range("C1"), _
                          SortOn:=xlSortOnValues, _
                          Order:=xlAscending, _
                          CustomOrder:="高,中,低", _
                          DataOption:=xlSortNormal
        .SetRange sortRange
        .Header = xlYes ' 1行目は見出しとして扱う
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply ' 並べ替えを適用
    End With
    
    MsgBox "優先度順に並べ替えました。", vbInformation
End Sub

' =================================================================
' プロシージャ名: SortByDeadline
' 機能: タスクリストを期限順（昇順）で並べ替える
' =================================================================
Sub SortByDeadline()

    ' --- 変数の宣言 ---
    Dim sortRange As Range
    Dim lastRow As Long

    ' --- 初期設定 ---
    lastRow = wsList.Cells(wsList.Rows.Count, 1).End(xlUp).Row
    
    ' 見出し行しかない場合は処理を終了
    If lastRow <= 1 Then
        MsgBox "並べ替えるタスクがありません。", vbExclamation
        Exit Sub
    End If

    ' 並べ替えの対象範囲（A1からデータの最終行まで）を設定
    Set sortRange = wsList.Range("A1:G" & lastRow)

    ' --- 並べ替えの実行 ---
    With wsList.Sort
        .SortFields.Clear ' 既存の並べ替え条件をクリア
        ' D列（期限）をキーに設定。Order:=xlAscending で昇順（古い順）にする
        .SortFields.Add Key:=wsList.Range("D1"), _
                          SortOn:=xlSortOnValues, _
                          Order:=xlAscending, _
                          DataOption:=xlSortNormal
        .SetRange sortRange
        .Header = xlYes ' 1行目は見出しとして扱う
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply ' 並べ替えを適用
    End With

    MsgBox "期限の昇順に並べ替えました。", vbInformation
End Sub

' 表に罫線を追加などの調整
Public Sub setGrid(targetRange As Range)
    ' 変数宣言
    Dim r As Range
    ' targetRangeの全セルをloop
    For Each r In targetRange
        ' 年月日入力セルなら
        If 1 < r.Row And ( _
            r.Column = 4 Or r.Column = 6 Or r.Column = 7 _
        ) Then
            ' 書式設定を「yyyy/mm/dd」へ変更
            r.NumberFormat = "yyyy/mm/dd"
        End If
    Next
    ' 列幅自動調整
    targetRange.EntireColumn.AutoFit
    ' 罫線
    targetRange.Borders.LineStyle = xlContinuous
End Sub
