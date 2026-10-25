Option Explicit

' ==========================================
' 設定項目
' ==========================================
Private Const TARGET_ROOT_FOLDER As String = "C:\Users\username\Desktop\TargetFolder"
Private Const TARGET_SHEET_NAME As String = "Sheet1"
Private Const SHEET_PASSWORD As String = "あなたのパスワード"
Private Const TARGET_CELL As String = "A1"
Private Const NEW_VALUE As String = "書き換え後の値"
Private Const LOG_SHEET_NAME As String = "更新結果"

' 処理結果
Private successCount As Long
Private noChangeCount As Long
Private errorCount As Long
Private logWs As Worksheet
Private logRow As Long

' ==========================================
' メイン処理
' ==========================================
Public Sub RunUpdateAllSubfolders()

    Dim fso As Object
    Dim originalSecurity As MsoAutomationSecurity
    Dim originalEvents As Boolean
    Dim originalAlerts As Boolean
    Dim originalScreen As Boolean
    Dim originalAskLinks As Boolean
    Dim errorMessage As String
    Dim resultMessage As String
    Dim settingsSaved As Boolean

    On Error GoTo ErrorHandler

    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FolderExists(TARGET_ROOT_FOLDER) Then
        MsgBox "対象フォルダが存在しません。" & vbCrLf & _
               TARGET_ROOT_FOLDER, vbCritical
        Exit Sub
    End If

    ' Excelの設定を退避
    originalSecurity = Application.AutomationSecurity
    originalEvents = Application.EnableEvents
    originalAlerts = Application.DisplayAlerts
    originalScreen = Application.ScreenUpdating
    originalAskLinks = Application.AskToUpdateLinks
    settingsSaved = True

    successCount = 0
    noChangeCount = 0
    errorCount = 0

    ' 更新結果シートを準備
    Set logWs = GetLogWorksheet()
    PrepareLogWorksheet

    ' マクロ実行防止・画面更新停止
    Application.AutomationSecurity = msoAutomationSecurityForceDisable
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    Application.AskToUpdateLinks = False

    ' フォルダを再帰処理
    ProcessFolder fso.GetFolder(TARGET_ROOT_FOLDER)

    GoTo CleanUp

ErrorHandler:
    errorMessage = Err.Description

CleanUp:
    ' Excelの設定を復元
    On Error Resume Next

    If settingsSaved Then
        Application.AutomationSecurity = originalSecurity
        Application.EnableEvents = originalEvents
        Application.DisplayAlerts = originalAlerts
        Application.ScreenUpdating = originalScreen
        Application.AskToUpdateLinks = originalAskLinks
    End If

    If Not logWs Is Nothing Then
        With logWs
            .Rows(1).Font.Bold = True
            .Columns("A:H").AutoFit
            .Range("A1:H1").AutoFilter
        End With
    End If

    On Error GoTo 0

    resultMessage = "処理結果" & vbCrLf & _
                    "更新成功: " & successCount & " 件" & vbCrLf & _
                    "変更なし: " & noChangeCount & " 件" & vbCrLf & _
                    "更新失敗: " & errorCount & " 件"

    If Len(errorMessage) > 0 Then
        resultMessage = resultMessage & vbCrLf & _
                        "処理中断: " & errorMessage
    End If

    MsgBox resultMessage, vbInformation

End Sub

' ==========================================
' 更新結果シート取得
' ==========================================
Private Function GetLogWorksheet() As Worksheet

    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(LOG_SHEET_NAME)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add( _
            After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = LOG_SHEET_NAME
    End If

    Set GetLogWorksheet = ws

End Function

' ==========================================
' 更新結果シート初期化
' ==========================================
Private Sub PrepareLogWorksheet()

    With logWs
        If .AutoFilterMode Then .AutoFilterMode = False
        .Cells.Clear

        .Range("A1:H1").Value = Array( _
            "No.", "ファイルパス", "シート", "セル", _
            "変更前", "変更後", "結果", "エラー内容")

        .Columns("B").NumberFormat = "@"
        .Columns("E:F").NumberFormat = "@"
    End With

    logRow = 2

End Sub

' ==========================================
' 更新結果記録
' ==========================================
Private Sub WriteUpdateLog( _
    ByVal filePath As String, _
    ByVal oldValue As Variant, _
    ByVal newValue As Variant, _
    ByVal resultStatus As String, _
    ByVal errorMessage As String)

    With logWs
        .Cells(logRow, 1).Value = logRow - 1
        .Cells(logRow, 2).Value = filePath
        .Cells(logRow, 3).Value = TARGET_SHEET_NAME
        .Cells(logRow, 4).Value = TARGET_CELL

        ' エラー値・Nullなどを考慮して文字列化
        .Cells(logRow, 5).Value = SafeToString(oldValue)
        .Cells(logRow, 6).Value = SafeToString(newValue)

        .Cells(logRow, 7).Value = resultStatus
        .Cells(logRow, 8).Value = errorMessage
    End With

    logRow = logRow + 1

End Sub

' ==========================================
' ログ出力用の文字列変換
' ==========================================
Private Function SafeToString(ByVal value As Variant) As String

    If IsError(value) Then
        SafeToString = "(Excelエラー値)"
    ElseIf IsNull(value) Then
        SafeToString = "(Null)"
    ElseIf IsEmpty(value) Then
        SafeToString = ""
    Else
        SafeToString = CStr(value)
    End If

End Function

' ==========================================
' フォルダ再帰処理
' ==========================================
Private Sub ProcessFolder(ByVal currentFolder As Object)

    Dim fileItem As Object
    Dim subFolderItem As Object

    For Each fileItem In currentFolder.Files

        If LCase$(Right$(fileItem.Name, 4)) = ".xls" And _
           Left$(fileItem.Name, 2) <> "~$" And _
           StrComp(fileItem.Path, ThisWorkbook.FullName, vbTextCompare) <> 0 Then

            UpdateWorkbook CStr(fileItem.Path)

        End If

    Next fileItem

    For Each subFolderItem In currentFolder.SubFolders
        ProcessFolder subFolderItem
    Next subFolderItem

End Sub

' ==========================================
' Excelファイル更新処理
' ==========================================
Private Sub UpdateWorkbook(ByVal filePath As String)

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim oldValue As Variant
    Dim resultStatus As String
    Dim errorMessage As String

    ' 元の保護状態
    Dim wasProtected As Boolean
    Dim protectDrawingObjects As Boolean
    Dim protectContents As Boolean
    Dim protectScenarios As Boolean

    ' 元の保護オプション
    Dim allowFormattingCells As Boolean
    Dim allowFormattingColumns As Boolean
    Dim allowFormattingRows As Boolean
    Dim allowInsertingColumns As Boolean
    Dim allowInsertingRows As Boolean
    Dim allowInsertingHyperlinks As Boolean
    Dim allowDeletingColumns As Boolean
    Dim allowDeletingRows As Boolean
    Dim allowSorting As Boolean
    Dim allowFiltering As Boolean
    Dim allowUsingPivotTables As Boolean

    On Error GoTo ErrorHandler

    resultStatus = "更新失敗"

    ' マクロ無効・外部リンク更新なしで開く
    Set wb = Workbooks.Open( _
        Filename:=filePath, _
        UpdateLinks:=0, _
        ReadOnly:=False, _
        IgnoreReadOnlyRecommended:=True, _
        AddToMru:=False)

    If wb.ReadOnly Then
        Err.Raise vbObjectError + 1000, , _
                  "読み取り専用で開かれました。"
    End If

    ' 対象シート取得
    On Error Resume Next
    Set ws = wb.Worksheets(TARGET_SHEET_NAME)
    On Error GoTo ErrorHandler

    If ws Is Nothing Then
        Err.Raise vbObjectError + 1001, , _
                  "対象シートが存在しません。"
    End If

    ' 変更前の値を取得
    oldValue = ws.Range(TARGET_CELL).Value2

    ' 同一値なら更新しない
    If Not IsError(oldValue) Then
        If VarType(oldValue) = vbString Then
            If StrComp(CStr(oldValue), NEW_VALUE, vbBinaryCompare) = 0 Then
                resultStatus = "変更なし"
                GoTo Finally
            End If
        End If
    End If

    ' 元の保護状態を取得
    protectDrawingObjects = ws.ProtectDrawingObjects
    protectContents = ws.ProtectContents
    protectScenarios = ws.ProtectScenarios

    wasProtected = protectDrawingObjects Or _
                   protectContents Or protectScenarios

    If wasProtected Then

        With ws.Protection
            allowFormattingCells = .AllowFormattingCells
            allowFormattingColumns = .AllowFormattingColumns
            allowFormattingRows = .AllowFormattingRows
            allowInsertingColumns = .AllowInsertingColumns
            allowInsertingRows = .AllowInsertingRows
            allowInsertingHyperlinks = .AllowInsertingHyperlinks
            allowDeletingColumns = .AllowDeletingColumns
            allowDeletingRows = .AllowDeletingRows
            allowSorting = .AllowSorting
            allowFiltering = .AllowFiltering
            allowUsingPivotTables = .AllowUsingPivotTables
        End With

        ' 保護解除
        ws.Unprotect Password:=SHEET_PASSWORD

    End If

    ' セルを書き換える
    ws.Range(TARGET_CELL).Value = NEW_VALUE

    ' 元の保護状態に戻す
    If wasProtected Then
        ws.Protect _
            Password:=SHEET_PASSWORD, _
            DrawingObjects:=protectDrawingObjects, _
            Contents:=protectContents, _
            Scenarios:=protectScenarios, _
            AllowFormattingCells:=allowFormattingCells, _
            AllowFormattingColumns:=allowFormattingColumns, _
            AllowFormattingRows:=allowFormattingRows, _
            AllowInsertingColumns:=allowInsertingColumns, _
            AllowInsertingRows:=allowInsertingRows, _
            AllowInsertingHyperlinks:=allowInsertingHyperlinks, _
            AllowDeletingColumns:=allowDeletingColumns, _
            AllowDeletingRows:=allowDeletingRows, _
            AllowSorting:=allowSorting, _
            AllowFiltering:=allowFiltering, _
            AllowUsingPivotTables:=allowUsingPivotTables
    End If

    ' 更新成功時のみ保存
    wb.Save
    resultStatus = "更新成功"

    GoTo Finally

ErrorHandler:
    errorMessage = Err.Description
    resultStatus = "更新失敗"

Finally:
    ' ブックを閉じる
    On Error Resume Next

    If Not wb Is Nothing Then
        wb.Close SaveChanges:=False
    End If

    On Error GoTo 0

    ' 件数集計
    Select Case resultStatus
        Case "更新成功"
            successCount = successCount + 1
        Case "変更なし"
            noChangeCount = noChangeCount + 1
        Case Else
            errorCount = errorCount + 1
    End Select

    ' 更新結果シートに記録
    WriteUpdateLog filePath, oldValue, NEW_VALUE, _
                   resultStatus, errorMessage

End Sub
End Sub
