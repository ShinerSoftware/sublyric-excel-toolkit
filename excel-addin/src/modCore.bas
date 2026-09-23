Attribute VB_Name = "modCore"
Option Explicit

' ============================================================
' 韵词字幕 · Excel 字幕工具箱 - 核心模块
' 功能：中英文空格清洗 / 结尾符号清理 / 表格美化 /
'       双语&中文 SRT 导出 / 双语 Word 导出
' ============================================================


' ---------- 1. 中英文之间自动加空格 ----------
Sub AddEnglishSpaces()
    Dim cell As Range
    Dim regex As Object
    Dim s As String

    If TypeName(Selection) <> "Range" Then
        MsgBox "请先选中包含文本的单元格区域。", vbExclamation
        Exit Sub
    End If

    Set regex = CreateObject("VBScript.RegExp")
    With regex
        .Global = True
        .MultiLine = True
        .IgnoreCase = True
    End With

    Application.ScreenUpdating = False

    For Each cell In Selection
        If Not IsEmpty(cell.Value) Then
            s = CStr(cell.Value)

            regex.Pattern = "([^\x00-\xff])([a-zA-Z])"      ' 中文后接英文
            s = regex.Replace(s, "$1 $2")
            regex.Pattern = "([a-zA-Z])([^\x00-\xff])"      ' 英文后接中文
            s = regex.Replace(s, "$1 $2")

            regex.Pattern = "(\d)([a-zA-Z])"                ' 数字后接英文（不加空格）
            s = regex.Replace(s, "$1$2")
            regex.Pattern = "([a-zA-Z])(\d)"                ' 英文后接数字（不加空格）
            s = regex.Replace(s, "$1$2")

            regex.Pattern = "([!@#$%^&*(),.?`~])([a-zA-Z])" ' 符号后接英文
            s = regex.Replace(s, "$1$2")

            regex.Pattern = "([a-zA-Z]) ([!@#$%^&*(),.?`~])" ' 英文+空格+符号 -> 去空格
            s = regex.Replace(s, "$1$2")

            regex.Pattern = "^ +| +$"                       ' 句首句尾空格
            s = regex.Replace(s, "")

            regex.Pattern = " {2,}"                         ' 合并多空格
            s = regex.Replace(s, " ")

            cell.Value = s
        End If
    Next cell

    Application.ScreenUpdating = True
    MsgBox "中英文空格处理完成！", vbInformation
End Sub


' ---------- 2. 去除句末多余符号（保留问号/感叹号） ----------
Sub RemoveEndingSymbols()
    Dim rng As Range
    Dim cell As Range
    Dim textVal As String
    Dim lastChar As String
    Dim symbolsToDelete As String

    symbolsToDelete = "。,.;:，；：、．,]"  ' 需要删除的符号，可自行增补

    On Error Resume Next
    Set rng = Application.Selection
    On Error GoTo 0

    If rng Is Nothing Or TypeName(rng) <> "Range" Then
        MsgBox "请先选择要处理的单元格区域。", vbExclamation
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    For Each cell In rng
        If Not IsEmpty(cell.Value) Then
            textVal = CStr(cell.Value)

            Do While Len(textVal) > 0
                lastChar = Right$(textVal, 1)

                If InStr("?!？！", lastChar) > 0 Then Exit Do   ' 保留的符号

                If InStr(symbolsToDelete, lastChar) > 0 Then
                    textVal = Left$(textVal, Len(textVal) - 1)
                Else
                    Exit Do
                End If
            Loop

            cell.Value = textVal
        End If
    Next cell

    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    MsgBox "结尾符号处理完成！", vbInformation
End Sub


' ---------- 3. 表格美化（表头/隔行底色/文本列格式/首字母大写） ----------
Sub FormatSheet()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim cell As Range
    Dim firstChar As String

    Set ws = ActiveSheet
    Application.ScreenUpdating = False

    On Error Resume Next
    lastRow = ws.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    On Error GoTo 0

    If lastRow = 0 Then
        Application.ScreenUpdating = True
        MsgBox "当前工作表没有数据。", vbExclamation
        Exit Sub
    End If

    ws.Rows("1:" & lastRow).RowHeight = 21

    With ws.Rows(1)
        .Interior.Color = RGB(128, 128, 128)
        .Font.Color = RGB(255, 255, 255)
    End With

    For i = 2 To lastRow Step 2
        ws.Rows(i).Interior.Color = RGB(240, 240, 240)
    Next i

    With ws.Range("B:C")
        .NumberFormat = "@"
        For Each cell In .Cells
            If Not IsEmpty(cell) Then cell.Value = CStr(cell.Value)
        Next
    End With

    For Each cell In ws.Range("D1:E" & lastRow)
        If Not IsEmpty(cell) Then cell.Value = Trim(cell.Value)
    Next

    For Each cell In ws.Range("D1:D" & lastRow)
        If Not IsEmpty(cell) And Len(cell.Value) > 0 Then
            firstChar = Left(cell.Value, 1)
            If firstChar Like "[a-z]" Then
                cell.Value = UCase(firstChar) & Mid(cell.Value, 2)
            End If
        End If
    Next

    Application.ScreenUpdating = True
    MsgBox "表格格式设置完成！", vbInformation
End Sub


' ---------- 4. 统一的时间格式转换（供导出宏调用） ----------
' 兼容两种输入：
'   1) Excel 时间序列值（数字，如单元格设为"时间"格式）
'   2) 文本时间字符串，如 "00:00:05.5" 或 "00:00:05.500"
' 输出统一为 SRT 标准格式："HH:MM:SS,mmm"
Function FormatTimeToSRT(timeValue As Variant) As String
    Dim totalTime As Double
    Dim hours As Long, minutes As Long, seconds As Long, milliseconds As Long
    Dim s As String
    Dim parts() As String

    If IsNumeric(timeValue) Then
        totalTime = CDbl(timeValue) * 86400

        hours = Int(totalTime / 3600)
        minutes = Int((totalTime - hours * 3600) / 60)
        seconds = Int(totalTime - hours * 3600 - minutes * 60)
        milliseconds = CLng(Round((totalTime - Int(totalTime)) * 1000))

        FormatTimeToSRT = Format(hours, "00") & ":" & _
                           Format(minutes, "00") & ":" & _
                           Format(seconds, "00") & "," & _
                           Format(milliseconds, "000")
    Else
        s = CStr(timeValue)
        If InStr(s, ".") > 0 Then
            parts = Split(s, ".")
            s = parts(0) & "," & Left(parts(1) & "000", 3)
        Else
            s = Replace(s, ".", ",")
        End If
        FormatTimeToSRT = s
    End If
End Function


' ---------- 5. 导出双语 SRT ----------
' 期望列结构：A=序号  B=开始时间  C=结束时间  D=英文字幕  E=中文字幕
Sub ExportBilingualSRT()
    Dim fs As Object, ts As Object
    Dim i As Long, lastRow As Long
    Dim filePath As String

    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then
        MsgBox "没有可导出的数据（需要从第 2 行开始填写）。", vbExclamation
        Exit Sub
    End If

    filePath = Application.GetSaveAsFilename( _
        FileFilter:="SRT Files (*.srt), *.srt", _
        Title:="保存双语字幕文件")
    If filePath = "False" Then Exit Sub

    Set fs = CreateObject("Scripting.FileSystemObject")
    Set ts = fs.CreateTextFile(filePath, True, True)

    For i = 2 To lastRow
        ts.WriteLine Cells(i, 1).Value
        ts.WriteLine FormatTimeToSRT(Cells(i, 2).Value) & " --> " & FormatTimeToSRT(Cells(i, 3).Value)
        ts.WriteLine Cells(i, 4).Value ' 英文
        ts.WriteLine Cells(i, 5).Value ' 中文
        ts.WriteLine ""
    Next i

    ts.Close
    MsgBox "双语字幕导出完成！" & vbCrLf & "保存位置：" & filePath, vbInformation
End Sub


' ---------- 6. 导出中文 SRT（仅 E 列） ----------
Sub ExportChineseSRT()
    Dim fs As Object, ts As Object
    Dim i As Long, lastRow As Long
    Dim filePath As String

    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then
        MsgBox "没有可导出的数据（需要从第 2 行开始填写）。", vbExclamation
        Exit Sub
    End If

    filePath = Application.GetSaveAsFilename( _
        FileFilter:="SRT Files (*.srt), *.srt", _
        Title:="保存中文字幕文件")
    If filePath = "False" Then Exit Sub

    Set fs = CreateObject("Scripting.FileSystemObject")
    Set ts = fs.CreateTextFile(filePath, True, True)

    For i = 2 To lastRow
        ts.WriteLine Cells(i, 1).Value
        ts.WriteLine FormatTimeToSRT(Cells(i, 2).Value) & " --> " & FormatTimeToSRT(Cells(i, 3).Value)
        ts.WriteLine Cells(i, 5).Value ' 仅中文
        ts.WriteLine ""
    Next i

    ts.Close
    MsgBox "中文字幕导出完成！" & vbCrLf & "保存位置：" & filePath, vbInformation
End Sub


' ---------- 7. 导出双语字幕为 Word 文档 ----------
Sub ExportBilingualToWord()
    Dim wordApp As Object, wordDoc As Object
    Dim i As Long, lastRow As Long
    Dim filePath As String, defaultName As String

    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then
        MsgBox "没有可导出的数据（需要从第 2 行开始填写）。", vbExclamation
        Exit Sub
    End If

    defaultName = Left(ThisWorkbook.Name, InStrRev(ThisWorkbook.Name, ".") - 1) & "_字幕"

    filePath = Application.GetSaveAsFilename( _
        InitialFileName:=defaultName, _
        FileFilter:="Word Documents (*.docx), *.docx", _
        Title:="保存双语字幕Word文档")
    If filePath = "False" Then Exit Sub

    If Not LCase(Right(filePath, 5)) = ".docx" Then
        filePath = filePath & ".docx"
    End If

    On Error Resume Next
    Set wordApp = CreateObject("Word.Application")
    On Error GoTo 0

    If wordApp Is Nothing Then
        MsgBox "无法创建 Word 应用程序，请确保已安装 Microsoft Word。", vbCritical
        Exit Sub
    End If

    wordApp.Visible = False
    Set wordDoc = wordApp.Documents.Add

    For i = 2 To lastRow
        With wordDoc
            .content.InsertAfter Cells(i, 1).Value & vbCrLf
            .content.InsertAfter FormatTimeToSRT(Cells(i, 2).Value) & " --> " & FormatTimeToSRT(Cells(i, 3).Value) & vbCrLf
            .content.InsertAfter Cells(i, 4).Value & vbCrLf  ' 英文
            .content.InsertAfter Cells(i, 5).Value & vbCrLf  ' 中文
            .content.InsertAfter vbCrLf
        End With
    Next i

    wordDoc.SaveAs2 filePath
    wordDoc.Close
    wordApp.Quit

    Set wordDoc = Nothing
    Set wordApp = Nothing

    MsgBox "双语字幕 Word 文档导出完成！" & vbCrLf & "保存位置：" & filePath, vbInformation
End Sub
