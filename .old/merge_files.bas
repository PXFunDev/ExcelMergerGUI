Attribute VB_Name = "merge_files"
Option Explicit

' -----------------------------------------------------------
' プログラム名: merge_files.bas
' 作成者: あなたの名前
' 作成日: 2025年4月9日
' バージョン: 1.0
' 説明: このプログラムは、〇〇を実行するためのものです。
' 使用方法: 実行方法の説明をここに記載
' -----------------------------------------------------------
' 変更履歴:
' 日付        バージョン    変更内容
' ----------  ----------  -----------------------------------
' 2025/04/09  1.0         初版作成
'
' -----------------------------------------------------------

' -----------------------------------------------------------
' ## 要件定義
' - 転記元のExcelファイルを手動選択する
' - 転記先のExcelファイルは、マクロが実行されているブックとする
' - 対象のシートは「オンサイト」「センドバック」「Nパッケージ」の3つ
' - 転記の基準は、B列の登録番号とする
' - 転記元の４行目以降を参照
' - 転記元のB列移行を参照
' - 転記元のオンサイト：AR列、センドバック：AP列、Nパッケージ：AP列までを転記
' - 転記元のデータを閉じる
' - マクロ実行のログを記録する
' -----------------------------------------------------------
'************************************************************
' タイトル欄
'************************************************************
Public Sub merge_files()
  
  On Error GoTo ErrorHandler

    ' マクロを実行するかどうか確認
    Dim response As VbMsgBoxResult
    response = MsgBox("マクロを実行しますか？", vbYesNo + vbQuestion, "確認")
    If response = vbYes Then

        '************************************************************
        ' 事前準備
        '************************************************************
        ' 処理開始時間を記録
        Dim T As Double
        T = Timer
        ' 自動計算＆画面更新停止
        Application.Calculation = xlCalculationManual
        Application.ScreenUpdating = False

        '************************************************************
        ' メインの処理
        '************************************************************
        ' 転記元のExcelファイル
        Dim SrcWb As Workbook
        Dim SrcWs_Onsite As Worksheet
        Dim SrcWs_Sendback As Worksheet
        Dim SrcWs_NPackage As Worksheet

        ' 任意のExcelファイルを開く(手動選択)
        Set SrcWb = Application.Workbooks.Open(Application.GetOpenFilename("Excel Files (*.xls; *.xlsx), *.xls; *.xlsx", , "対象ファイルを選択してください"))
        If SrcWb Is Nothing Then
            MsgBox "ファイルが選択されていません。処理を中止します。", vbExclamation
            Exit Sub
        End If

        ' 対象ファイルにオンサイト、センドバック、Nパッケージが存在するか確認
        Set SrcWs_Onsite = SrcWb.Worksheets("オンサイト")
        Set SrcWs_Sendback = SrcWb.Worksheets("センドバック")
        Set SrcWs_NPackage = SrcWb.Worksheets("Nパッケージ")

        ' 転記元のフィルター解除
        SrcWs_Onsite.AutoFilterMode = False
        SrcWs_Sendback.AutoFilterMode = False
        SrcWs_NPackage.AutoFilterMode = False

        ' 転記先のExcelファイル（マクロが実行されているブック）
        Dim dstWb           As Workbook
        Dim dstWs_Onsite    As Worksheet
        Dim dstWs_Sendback  As Worksheet
        Dim dstWs_NPackage  As Worksheet

        Set dstWb = ThisWorkbook ' マクロが実行されているブック
        Set dstWs_Onsite = dstWb.Worksheets("オンサイト") ' 転記先のオンサイトシート
        Set dstWs_Sendback = dstWb.Worksheets("センドバック") ' 転記先のセンドバックシート
        Set dstWs_NPackage = dstWb.Worksheets("Nパッケージ") ' 転記先のNパッケージシート

        ' 対象ファイルからオンサイトのデータを転記
        ' オンサイトシートが存在しない場合は、エラーメッセージを表示
        If SrcWs_Onsite Is Nothing Then
            MsgBox "対象ファイルにオンサイトシートが存在しません。", vbExclamation
            SrcWb.Close False
            Exit Sub
        Else
            ' オンサイトシートが存在する場合は、転記を実行
            Call CheckDataExists(SrcWs_Onsite, dstWs_Onsite)
        End If

        ' 対象ファイルからセンドバックのデータを転記
        ' センドバックシートが存在しない場合は、エラーメッセージを表示
        If SrcWs_Sendback Is Nothing Then
            MsgBox "対象ファイルにセンドバックシートが存在しません。", vbExclamation
            SrcWb.Close False
            Exit Sub
        Else
            ' センドバックシートが存在する場合は、転記を実行
            Call CheckDataExists(SrcWs_Sendback, dstWs_Sendback)
        End If

        ' 対象ファイルからNパッケージのデータを転記
        ' Nパッケージシートが存在しない場合は、エラーメッセージを表示
        If SrcWs_NPackage Is Nothing Then
            MsgBox "対象ファイルにNパッケージシートが存在しません。", vbExclamation
            SrcWb.Close False
            Exit Sub
        Else
            ' Nパッケージシートが存在する場合は、転記を実行
            Call CheckDataExists(SrcWs_NPackage, dstWs_NPackage)
        End If

        ' 転記元のデータを閉じる(セーブしない)
        Workbooks(SrcWb.Name).Close False

        '************************************************************
        ' 残作業
        '************************************************************
        ' 自動計算＆画面更新再開
        Application.Calculation = xlCalculationAutomatic
        Application.ScreenUpdating = True
        ' ログ: マクロ実行の成功を記録
        Call LogMacroExecution("merge_files", "成功")
        ' 処理完了メッセージ
        MsgBox "マクロを実行しました。" & vbCrLf & "処理時間: " & Format(Timer - T, "0.00") & " 秒"
    Else
        ' ログ: マクロ実行のキャンセルを記録
        Call LogMacroExecution("merge_files", "キャンセル")
        ' キャンセルメッセージ
        MsgBox "マクロの実行をキャンセルしました。"

    End If
    Exit Sub

'************************************************************
' エラーハンドリング
'************************************************************
ErrorHandler:
    ' 転記元のデータを閉じる(セーブしない)
    Workbooks(SrcWb.Name).Close False
    ' ログ：マクロ失敗時、エラーメッセージをログに記録
    Call LogMacroExecution("merge_files", "失敗 - " & Err.Description)
    ' エラーメッセージを表示
    MsgBox "エラーが発生しました。管理者に連絡ください。" & vbCrLf & "エラー内容: " & Err.Description, vbCritical
    ' 自動計算＆画面更新再開
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

End Sub

' ************************************************************
' 転記先に転記元のデータがないか検証するサブルーチン
Sub CheckDataExists(srcWs As Worksheet, dstWs As Worksheet)
    ' 変数宣言を集約
    Dim i As Long
    Dim key As Variant
    Dim srcRegNo As String
    Dim dstRegNo As String
    Dim lastRow As Long
    Dim dstLastRow As Long
    Dim srcDict As Object
    Dim DstDict As Object
    Set srcDict = CreateObject("Scripting.Dictionary")
    Set DstDict = CreateObject("Scripting.Dictionary")

    ' フィルターを解除
    srcWs.AutoFilterMode = False
    dstWs.AutoFilterMode = False

    ' 転記元の登録番号を集計
    lastRow = srcWs.Cells(srcWs.Rows.Count, 2).End(xlUp).Row
    For i = 4 To lastRow
        srcRegNo = srcWs.Cells(i, 2).Value
        If srcRegNo <> "" Then
            If Not srcDict.Exists(srcRegNo) Then
                srcDict.Add srcRegNo, 1
            Else
                srcDict(srcRegNo) = srcDict(srcRegNo) + 1
            End If
        End If
    Next i

    ' 転記先の登録番号を集計
    dstLastRow = dstWs.Cells(dstWs.Rows.Count, 2).End(xlUp).Row
    For i = 4 To dstLastRow
        dstRegNo = dstWs.Cells(i, 2).Value
        If dstRegNo <> "" Then
            If Not DstDict.Exists(dstRegNo) Then
                DstDict.Add dstRegNo, 1
            Else
                DstDict(dstRegNo) = DstDict(dstRegNo) + 1
            End If
        End If
    Next i

    ' 登録番号ごとに個数比較
    Dim needUpdate As Boolean: needUpdate = False
    Dim warnMsg As String: warnMsg = ""
    For Each key In srcDict.Keys
        If Not DstDict.Exists(key) Then
            needUpdate = True
            warnMsg = "転記先に登録番号 [" & key & "] がありません。"
            Exit For
        ElseIf srcDict(key) <> DstDict(key) Then
            needUpdate = True
            warnMsg = "登録番号 [" & key & "] の個数が一致しません。" & vbCrLf & _
                        "転記元: " & srcDict(key) & "件, 転記先: " & DstDict(key) & "件"
            Exit For
        End If
    Next key
    For Each key In DstDict.Keys
        If Not srcDict.Exists(key) Then
            needUpdate = True
            warnMsg = "転記先に余分な登録番号 [" & key & "] があります。"
            Exit For
        End If
    Next key

    If needUpdate Then
        MsgBox "転記元と転記先でデータの過不足があります。" & vbCrLf & _
                warnMsg & vbCrLf & _
                "手動で転記先のデータを修正してください。", vbExclamation
        Exit Sub
    End If
    ' 個数がすべて一致：何もしない
    Exit Sub


    ' 転記元の該当行のみ転記
    Call CopyData(srcWs, dstWs, 4, lastRow - 3)
End Sub

' ************************************************************
' データを転記するサブルーチン

Private Sub CopyData(srcWs As Worksheet, _
            dstWs As Worksheet, _
            targetRow As Long, _
            copyRowCount As Long _
            )

    Dim lastRow As Long
    Dim i As Long
    Dim col As Long

    ' 転記先のシートの最終行を取得（B列基準）
    lastRow = dstWs.Cells(dstWs.Rows.Count, 2).End(xlUp).Row + 1

    ' 要件定義に基づき、B列以降を転記
    Dim colStart As Long, colEnd As Long
    ' シート名で転記範囲を切り替え
    Select Case dstWs.Name
        Case "オンサイト"
            colStart = 2 ' B列
            colEnd = 44  ' AR列
        Case "センドバック"
            colStart = 2 ' B列
            colEnd = 42  ' AP列
        Case "Nパッケージ"
            colStart = 2 ' B列
            colEnd = 42  ' AP列
    End Select

    For i = targetRow To targetRow + copyRowCount - 1
        For col = colStart To colEnd
            dstWs.Cells(lastRow, col).Value = srcWs.Cells(i, col).Value
        Next col
        lastRow = lastRow + 1
    Next i
End Sub



