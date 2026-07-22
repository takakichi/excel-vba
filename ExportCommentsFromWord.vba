Option Explicit

' ==============================================================================
' 1. Wordからコメントを抽出するマクロ（Word -> Excel）
' ==============================================================================
Sub ExportCommentsFromWord()
    Dim wordFile As Variant
    Dim wdApp As Object, wdDoc As Object, cmt As Object
    Dim xlSheet As Worksheet
    Dim i As Long

    On Error GoTo ErrorHandler

    ' -------------------------------------------------------------------------
    ' 1. 対象のWordファイルを選択
    ' -------------------------------------------------------------------------
    wordFile = Application.GetOpenFilename( FileFilter:="Word Files (*.docx; *.docm), *.docx; *.docm", Title:="Wordファイルを選択してください")
    If wordFile = False Then Exit Sub

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ' -------------------------------------------------------------------------
    ' 2. Wordファイルを開く
    ' -------------------------------------------------------------------------
    Set wdApp = CreateObject("Word.Application")
    wdApp.Visible = False
    Set wdDoc = wdApp.Documents.Open(Filename:=wordFile, ReadOnly:=True)

    ' コメントの存在チェック
    If wdDoc.Comments.Count = 0 Then
        MsgBox "選択されたWord文書にはコメントが存在しません。", vbInformation
        GoTo CleanUp
    End If

    ' -------------------------------------------------------------------------
    ' 3. 現在のアクティブシートを初期化してヘッダーを設定
    ' -------------------------------------------------------------------------
    Set xlSheet = ActiveSheet
    xlSheet.Cells.Clear

    Dim headers As Variant
    headers = Array("ID", "Page", "Line", "Author", "Date", "Context", "Comment", "Status", "修正回答(ここに入力)")

    With xlSheet
        .Range("A1").Resize(1, UBound(headers) + 1).Value = headers
        .Range("A1").Resize(1, UBound(headers) + 1).Font.Bold = True
        .Range("A1").Resize(1, UBound(headers) + 1).Interior.Color = RGB(0, 112, 192) ' 濃い青
        .Range("A1").Resize(1, UBound(headers) + 1).Font.Color = RGB(255, 255, 255) ' 白文字
        .Activate
        ActiveWindow.FreezePanes = False
        ActiveWindow.FreezePanes = True ' 先頭行を固定
    End With

    ' -------------------------------------------------------------------------
    ' 4. データの書き出し
    ' -------------------------------------------------------------------------
    i = 2
    For Each cmt In wdDoc.Comments
        With xlSheet
            .Cells(i, 1).Value = cmt.Index ' コメントID
            .Cells(i, 2).Value = cmt.Reference.Information(3) ' ページ番号
            .Cells(i, 3).Value = cmt.Reference.Information(10) ' 行番号
            .Cells(i, 4).Value = cmt.Author
            .Cells(i, 5).Value = Format(cmt.Date, "yyyy/mm/dd HH:mm")
            .Cells(i, 6).Value = Replace(cmt.Scope.Text, vbCr, " ") ' 該当本文（改行除去）
            .Cells(i, 7).Value = cmt.Range.Text ' コメント本文

            If cmt.Done Then
                .Cells(i, 8).Value = "解決済み"
                .Cells(i, 8).Font.Color = RGB(128, 128, 128)
            Else
                .Cells(i, 8).Value = "未完了"
                .Cells(i, 8).Font.Bold = True
                .Cells(i, 8).Font.Color = RGB(255, 0, 0)
            End If
        End With
        i = i + 1
    Next cmt

    ' -------------------------------------------------------------------------
    ' 5. 書式とレイアウトの調整
    ' -------------------------------------------------------------------------
    With xlSheet
        .Columns("A:H").AutoFit
        .Columns("I:I").ColumnWidth = 30
        .Columns("I:I").Interior.Color = RGB(255, 255, 200) ' 修正回答欄を薄黄色に
        .Columns("F:G").ColumnWidth = 40
        .Columns("F:G").Style.WrapText = True ' 折り返し表示
        .Range("A1").CurrentRegion.Borders.LineStyle = 1 ' 罫線
        .Range("A1").AutoFilter ' オートフィルター設置
    End With

    ' -------------------------------------------------------------------------
    ' 6. 正常終了メッセージ
    ' -------------------------------------------------------------------------
    MsgBox "全 " & (i - 2) & " 件のコメント抽出が完了しました。" & vbCrLf & _
           "I列（黄色）に修正回答を入力してください。", vbInformation, "処理完了"

CleanUp:
    ' -------------------------------------------------------------------------
    ' 7. 終了処理
    ' -------------------------------------------------------------------------
    On Error Resume Next
    If Not wdDoc Is Nothing Then wdDoc.Close SaveChanges:=False
    If Not wdApp Is Nothing Then wdApp.Quit
    Set wdDoc = Nothing
    Set wdApp = Nothing
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Exit Sub

ErrorHandler:
    ' 予期せぬエラー発生時の処理
    MsgBox "エラーが発生しました。" & vbCrLf & _
           "エラー番号: " & Err.Number & vbCrLf & _
           "エラー内容: " & Err.Description, vbCritical, "エラー"
    Resume CleanUp
End Sub
