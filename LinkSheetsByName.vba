Sub LinkSheetsByName( ByVal startCellPostion As String)
    Dim ws As Worksheet
    Dim targetCell As Range
    Dim sheetNames As Collection
    Dim sheetName As String
    Dim startCell As Range
    
    ' 開始点の設定
    Set startCell = ActiveSheet.Range(startCellPostion)
    
    ' 高速化設定
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' 1. 全シート名を保存
    Set sheetNames = New Collection
    On Error Resume Next
    For Each ws In ThisWorkbook.Worksheets
        sheetNames.Add ws.Name, ws.Name
    Next ws

On Error GoTo 0
    Set targetCell = startCell
    On Error GoTo ErrorHandler
    Do While targetCell.Value <> ""
        sheetName = CStr(targetCell.Value)
        If ExistsInCollection(sheetNames, sheetName) Then
            ' シート名が存在する場合は、ハイパーリンクを設定
            ActiveSheet.Hyperlinks.Add _
                Anchor:=targetCell, _
                Address:="", _
                SubAddress:="'" & sheetName & "'!A1", _
                ScreenTip:="シート「" & sheetName & "」へ移動", _
                TextToDisplay:=sheetName
        Else
            ' ハイパーリンクを解除し、書式を標準に戻す
            targetCell.Hyperlinks.Delete
            targetCell.Style = "Normal"
        End If
        Set targetCell = targetCell.Offset(1, 0)
    Loop

CleanUp:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    MsgBox "処理が完了しました。", vbInformation
    Exit Sub

ErrorHandler:
    MsgBox "エラーが発生しました: " & Err.Description, vbCritical
    Resume CleanUp
End Sub

' コレクション内にキーが存在するか判定するヘルパー関数
Function ExistsInCollection(col As Collection, key As String) As Boolean
    On Error Resume Next
    col.Item key
    ExistsInCollection = (Err.Number = 0)
    On Error GoTo 0
End Function
