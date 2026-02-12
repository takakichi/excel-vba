' CreateSheetsInActiveWorkbookWithLinks.bas
' 選択範囲の値を元に ActiveWorkbook にワークシートを作成し、
' 元のセルに作成したシートへのハイパーリンクを貼るマクロ
'
' - 同じ値が複数ある場合、2回目以降は "_1", "_2", ... を付ける
' - すべての候補名を検証してから一括作成する
' - 作成先は ActiveWorkbook
' - 元の選択セルにハイパーリンクを追加する（保存されていない ActiveWorkbook へはリンクが機能しない場合があります）
'
' このマクロは生成AIで作成しています。
Option Explicit

' 定数定義
Private Const MAX_SHEET_NAME_LEN As Long = 31
Private Const FORBIDDEN_CHARS As String = ":/\?*[]"
Private Const SUFFIX_SEPARATOR As String = "_"

' メイン処理
Sub CreateSheetsInActiveWorkbookWithLinks()
    Dim sel As Range
    Dim destWB As Workbook
    Dim sheetNames As Object ' Dictionary: CellAddress -> ProposedName
    Dim validationError As String
    
    ' 1. 選択範囲の取得とチェック
    If Not GetSelection(sel) Then Exit Sub
    
    ' 2. 作成先ブックの決定
    Set destWB = ActiveWorkbook
    
    ' 3. シート名候補の生成（重複処理含む）
    Set sheetNames = GenerateSheetNames(sel)
    If sheetNames Is Nothing Then Exit Sub ' 空セルなどのエラー時は終了
    
    ' 4. 事前バリデーション（全件チェック）
    validationError = ValidateSheetNames(destWB, sheetNames)
    If validationError <> "" Then
        MsgBox "以下の理由により処理を中止しました:" & vbCrLf & vbCrLf & validationError, vbCritical, "作成不可"
        Exit Sub
    End If
    
    ' 5. シート作成とリンク設定（トランザクション的な実行）
    ProcessCreation sel, destWB, sheetNames
    
End Sub

' ---------------------------------------------------------
' Helper Functions
' ---------------------------------------------------------

' 選択範囲を取得し、基本的なチェックを行う
Private Function GetSelection(ByRef sel As Range) As Boolean
    On Error Resume Next
    Set sel = Application.Selection
    On Error GoTo 0
    
    If sel Is Nothing Then
        MsgBox "選択範囲が見つかりません。セルを選択して実行してください。", vbExclamation
        GetSelection = False
        Exit Function
    End If
    
    If sel.Count = 0 Then ' 通常あり得ないが念のため
        MsgBox "選択範囲が空です。", vbExclamation
        GetSelection = False
        Exit Function
    End If
    
    GetSelection = True
End Function

' セル範囲からシート名候補を生成する（重複時は _n を付与）
' 戻り値: Dictionary (Key: セルアドレス, Value: 作成予定のシート名)
Private Function GenerateSheetNames(ByVal rng As Range) As Object
    Dim cell As Range
    Dim dict As Object
    Dim nameCounts As Object
    Dim baseName As String, finalName As String, lowerKey As String
    
    Set dict = CreateObject("Scripting.Dictionary")
    Set nameCounts = CreateObject("Scripting.Dictionary")
    
    For Each cell In rng.Cells
        baseName = Trim(CStr(cell.Value))
        
        ' 空文字チェック
        If baseName = "" Then
            MsgBox "選択範囲に空のセルが含まれています（" & cell.Address(False, False) & "）。処理を中止します。", vbExclamation
            Set GenerateSheetNames = Nothing
            Exit Function
        End If
        
        ' 重複処理 (_n の付与)
        lowerKey = LCase(baseName)
        If Not nameCounts.Exists(lowerKey) Then
            nameCounts(lowerKey) = 1
            finalName = baseName
        Else
            finalName = baseName & SUFFIX_SEPARATOR & nameCounts(lowerKey)
            nameCounts(lowerKey) = nameCounts(lowerKey) + 1
        End If
        
        dict.Add cell.Address, finalName
    Next cell
    
    Set GenerateSheetNames = dict
End Function

' シート名の妥当性を検証する
' 戻り値: エラーメッセージ（問題なければ空文字）
Private Function ValidateSheetNames(ByVal wb As Workbook, ByVal candidates As Object) As String
    Dim errMsg As String
    Dim existingSheets As Object
    Dim ws As Worksheet
    Dim cellAddr As Variant
    Dim sheetName As String
    Dim i As Long
    Dim char As String
    
    errMsg = ""
    
    ' 既存シート名の取得
    Set existingSheets = CreateObject("Scripting.Dictionary")
    For Each ws In wb.Worksheets
        existingSheets(LCase(ws.Name)) = True
    Next ws
    
    ' 候補リスト内の重複チェック用（GenerateSheetNamesで処理済みだが、念のため全候補の最終確認）
    Dim batchCheck As Object
    Set batchCheck = CreateObject("Scripting.Dictionary")
    
    For Each cellAddr In candidates.Keys
        sheetName = candidates(cellAddr)
        
        ' 文字数チェック
        If Len(sheetName) > MAX_SHEET_NAME_LEN Then
            errMsg = errMsg & "- 文字数超過 (" & Len(sheetName) & "文字): " & sheetName & vbCrLf
        End If
        
        ' 禁止文字チェック
        For i = 1 To Len(FORBIDDEN_CHARS)
            char = Mid(FORBIDDEN_CHARS, i, 1)
            If InStr(sheetName, char) > 0 Then
                errMsg = errMsg & "- 禁止文字 '" & char & "' を含む: " & sheetName & vbCrLf
                Exit For
            End If
        Next i
        
        ' 既存シートとの重複チェック
        If existingSheets.Exists(LCase(sheetName)) Then
            errMsg = errMsg & "- 既存シートと重複: " & sheetName & vbCrLf
        End If
        
        ' 候補内での重複チェック（ロジック上発生しないはずだが安全策）
        If batchCheck.Exists(LCase(sheetName)) Then
            errMsg = errMsg & "- 作成予定リスト内で重複: " & sheetName & vbCrLf
        Else
            batchCheck(LCase(sheetName)) = True
        End If
    Next cellAddr
    
    ValidateSheetNames = errMsg
End Function

' シート作成とリンク貼り付けの実行部
Private Sub ProcessCreation(ByVal sourceRange As Range, ByVal destWB As Workbook, ByVal sheetNames As Object)
    Application.ScreenUpdating = False
    On Error GoTo ErrorHandler
    
    Dim cell As Range
    Dim newName As String
    Dim newWS As Worksheet
    Dim createdCount As Long
    Dim cellAddr As String
    
    createdCount = 0
    
    For Each cell In sourceRange.Cells
        cellAddr = cell.Address
        If sheetNames.Exists(cellAddr) Then
            newName = sheetNames(cellAddr)
            
            ' シート作成
            Set newWS = destWB.Worksheets.Add(After:=destWB.Worksheets(destWB.Worksheets.Count))
            newWS.Name = newName
            
            ' リンク作成
            AddHyperlinkToCell cell, destWB, newName
            
            createdCount = createdCount + 1
        End If
    Next cell
    
    Application.ScreenUpdating = True
    MsgBox createdCount & " 枚のシートを作成しました。", vbInformation, "完了"
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "予期せぬエラーが発生しました: " & Err.Number & vbCrLf & Err.Description, vbCritical
End Sub

' セルにハイパーリンクを設定する
Private Sub AddHyperlinkToCell(ByVal anchorCell As Range, ByVal targetWB As Workbook, ByVal sheetName As String)
    Dim subAddress As String
    subAddress = "'" & sheetName & "'!A1"
    
    ' 既存のハイパーリンクを削除
    If anchorCell.Hyperlinks.Count > 0 Then anchorCell.Hyperlinks.Delete
    
    ' リンク先のアドレス決定（同ブックか別ブックか）
    Dim address As String
    If anchorCell.Worksheet.Parent Is targetWB Then
        address = ""
    Else
        ' 別ブックの場合、保存されていないと機能しない可能性があるためフルパスを取得しようとする
        On Error Resume Next
        address = targetWB.FullName
        On Error GoTo 0
    End If
    
    anchorCell.Hyperlinks.Add _
        Anchor:=anchorCell, _
        address:=address, _
        subAddress:=subAddress, _
        TextToDisplay:=sheetName
End Sub
