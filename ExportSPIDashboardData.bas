Attribute VB_Name = "ExportSPIDashboardData"
Option Explicit

Public Sub ExportSPIDashboardJson()
    Dim targetFolder As String
    Dim outputPath As String
    Dim jsonText As String

    targetFolder = GetDashboardFolder()
    If Len(targetFolder) = 0 Then Exit Sub

    EnsureFolder JoinPath(targetFolder, "data")
    outputPath = JoinPath(JoinPath(targetFolder, "data"), "spi-dashboard-data.json")

    jsonText = "{""workbookName"":""" & JsonEscape(ThisWorkbook.Name) & """," & _
               """generatedAt"":""" & Format$(Now, "yyyy-mm-dd\Thh:nn:ss") & """," & _
               """sheets"":{" & _
               """SPI Master"":" & SheetToJsonArray(ThisWorkbook.Worksheets("SPI Master")) & "," & _
               """Monthly Data Entry"":" & SheetToJsonArray(ThisWorkbook.Worksheets("Monthly Data Entry")) & _
               "}}"

    WriteTextFile outputPath, jsonText

    MsgBox "SPI dashboard data exported:" & vbCrLf & outputPath, vbInformation, "SPI Dashboard Export"
End Sub

Private Function GetDashboardFolder() As String
    Dim defaultFolder As String
    Dim answer As Variant

    defaultFolder = "/Users/naufaniyaz/Documents/Codex/2026-06-04/files-mentioned-by-the-user-spi"
    answer = Application.InputBox( _
        Prompt:="Paste the dashboard project folder path." & vbCrLf & _
                "The macro will write data/spi-dashboard-data.json inside that folder.", _
        Title:="SPI Dashboard Export Folder", _
        Default:=defaultFolder, _
        Type:=2)

    If VarType(answer) = vbBoolean Then
        GetDashboardFolder = vbNullString
    Else
        GetDashboardFolder = Trim$(CStr(answer))
    End If
End Function

Private Function SheetToJsonArray(ByVal ws As Worksheet) As String
    Dim lastRow As Long
    Dim lastCol As Long
    Dim headers As Variant
    Dim values As Variant
    Dim rowIndex As Long
    Dim colIndex As Long
    Dim rowJson As String
    Dim result As String

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    If lastRow < 2 Or lastCol < 1 Then
        SheetToJsonArray = "[]"
        Exit Function
    End If

    headers = ws.Range(ws.Cells(1, 1), ws.Cells(1, lastCol)).Value2
    values = ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, lastCol)).Value2

    result = "["
    For rowIndex = 1 To UBound(values, 1)
        rowJson = "{"
        For colIndex = 1 To UBound(values, 2)
            If Len(CStr(headers(1, colIndex))) > 0 Then
                If Len(rowJson) > 1 Then rowJson = rowJson & ","
                rowJson = rowJson & """" & JsonEscape(CStr(headers(1, colIndex))) & """:" & JsonValue(values(rowIndex, colIndex))
            End If
        Next colIndex
        rowJson = rowJson & "}"

        If Len(result) > 1 Then result = result & ","
        result = result & rowJson
    Next rowIndex

    SheetToJsonArray = result & "]"
End Function

Private Function JsonValue(ByVal value As Variant) As String
    If IsError(value) Or IsEmpty(value) Then
        JsonValue = "null"
    ElseIf VarType(value) = vbBoolean Then
        JsonValue = LCase$(CStr(value))
    ElseIf IsNumeric(value) Then
        JsonValue = Replace(CStr(value), ",", ".")
    ElseIf IsDate(value) Then
        JsonValue = """" & Format$(CDate(value), "yyyy-mm-dd") & """"
    ElseIf Len(CStr(value)) = 0 Then
        JsonValue = "null"
    Else
        JsonValue = """" & JsonEscape(CStr(value)) & """"
    End If
End Function

Private Function JsonEscape(ByVal text As String) As String
    Dim index As Long
    Dim code As Long
    Dim ch As String
    Dim result As String

    For index = 1 To Len(text)
        ch = Mid$(text, index, 1)
        code = AscW(ch)
        If code < 0 Then code = code + 65536

        Select Case code
            Case 34
                result = result & Chr$(92) & Chr$(34)
            Case 92
                result = result & Chr$(92) & Chr$(92)
            Case 9
                result = result & "\t"
            Case 10
                result = result & "\n"
            Case 13
                result = result & "\n"
            Case 0 To 31, Is > 126
                result = result & "\u" & Right$("0000" & Hex$(code), 4)
            Case Else
                result = result & ch
        End Select
    Next index

    JsonEscape = result
End Function

Private Sub EnsureFolder(ByVal folderPath As String)
    If Len(Dir(folderPath, vbDirectory)) = 0 Then MkDir folderPath
End Sub

Private Sub WriteTextFile(ByVal filePath As String, ByVal text As String)
    Dim fileNumber As Integer

    fileNumber = FreeFile
    Open filePath For Output As #fileNumber
    Print #fileNumber, text
    Close #fileNumber
End Sub

Private Function JoinPath(ByVal leftPart As String, ByVal rightPart As String) As String
    Dim separator As String

    If InStr(1, Application.OperatingSystem, "Mac", vbTextCompare) > 0 Then
        separator = "/"
    Else
        separator = Application.PathSeparator
    End If

    If Right$(leftPart, 1) = "/" Or Right$(leftPart, 1) = "\" Or Right$(leftPart, 1) = ":" Then
        JoinPath = leftPart & rightPart
    Else
        JoinPath = leftPart & separator & rightPart
    End If
End Function
