Attribute VB_Name = "modReader"
Option Explicit

' ============================================================
' 韵词字幕 · Excel 字幕工具箱 - 朗读校对模块
'
' 设计说明：
' Excel VBA 原生没有办法在功能区里嵌入一个"真正的"网页侧边栏
' （那需要 Office.js 任务窗格加载项，是完全不同的技术栈，
'  不是普通 .xlam 加载项能做到的）。
'
' 这里采用的折中方案：调用系统上真正安装的 Edge（Chromium 内核），
' 以 --app 模式打开 reader.html（无地址栏/无标签页），并用
' --window-position / --window-size 参数让它贴靠在屏幕右侧、
' 高度撑满整个屏幕——视觉效果很接近"侧边栏"，同时因为用的是真正
' 的 Edge 内核，Web Speech API 能拿到 Edge 的"自然语音"，
' 比旧版 WebBrowser 控件（IE 内核）效果好得多。
'
' 关键点：给这个朗读窗口用独立的 --user-data-dir（独立的 Edge
' 用户配置目录），不跟你平时正常浏览用的 Edge 共用同一个实例。
' 如果不这样做，Chromium 的"单实例"机制会把我们的启动请求转发给
' 已经在跑的 Edge，导致我们传的命令行参数（文件访问权限、窗口位置
' 等）全部被忽略——这正是"文本填不进去 / 每次都开新窗口"的根因。
'
' "不新开窗口，直接替换内容"是这样实现的：
'   1. 每次点击"朗读校对"，先把选中文本写入一个固定的临时文件
'      （%TEMP%\SublyricReaderText.txt）
'   2. 遍历当前所有可见窗口，标题里包含"韵词朗读校对"关键字的
'      就认为朗读窗口已经开着：
'        - 有：不新开窗口，只要把文件内容更新了，页面里的轮询
'          逻辑会在 1 秒内自动发现文件变了，自动替换文本框内容
'          并重新开始朗读（用新内容覆盖旧的）
'        - 没有：正常走 Shell 启动一个新窗口，reader.html 加载
'          时也是走同一套"读取临时文件"的逻辑，逻辑保持统一
' ============================================================

#If VBA7 Then
    Private Declare PtrSafe Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
    Private Declare PtrSafe Function EnumWindows Lib "user32" (ByVal lpEnumFunc As LongPtr, ByVal lParam As LongPtr) As Long
    Private Declare PtrSafe Function GetWindowTextW Lib "user32" (ByVal hwnd As LongPtr, ByVal lpString As LongPtr, ByVal cch As Long) As Long
    Private Declare PtrSafe Function GetWindowTextLengthW Lib "user32" (ByVal hwnd As LongPtr) As Long
    Private Declare PtrSafe Function IsWindowVisible Lib "user32" (ByVal hwnd As LongPtr) As Long
    Private Declare PtrSafe Function SetForegroundWindow Lib "user32" (ByVal hwnd As LongPtr) As Long
#Else
    Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
    Private Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
    Private Declare Function GetWindowTextW Lib "user32" (ByVal hwnd As Long, ByVal lpString As Long, ByVal cch As Long) As Long
    Private Declare Function GetWindowTextLengthW Lib "user32" (ByVal hwnd As Long) As Long
    Private Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
    Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
#End If

Private Const SM_CXSCREEN As Long = 0
Private Const SM_CYSCREEN As Long = 1

' 朗读窗口标题里一定会包含的关键字（reader.html 的 <title> 里也有这几个字）
Private Const READER_WINDOW_KEYWORD As String = "韵词朗读校对"

#If VBA7 Then
    Private foundReaderHwnd As LongPtr
#Else
    Private foundReaderHwnd As Long
#End If


Sub OpenReaderSidebar()
    Dim selectedText As String
    Dim cell As Range
    Dim htmlFilePath As String
    Dim edgePath As String
    Dim textFilePath As String
    Dim profileDir As String
    #If VBA7 Then
        Dim hwnd As LongPtr
    #Else
        Dim hwnd As Long
    #End If
    Dim url As String
    Dim shellCommand As String
    Dim screenW As Long, screenH As Long
    Dim winW As Long, winH As Long
    Dim winX As Long, winY As Long

    If TypeName(Selection) <> "Range" Then
        MsgBox "请先选中要朗读校对的单元格。", vbExclamation
        Exit Sub
    End If

    For Each cell In Selection
        If Len(cell.Value) > 0 Then
            selectedText = selectedText & cell.Value & vbLf
        End If
    Next cell
    selectedText = Trim(selectedText)

    If Len(selectedText) = 0 Then
        MsgBox "选中的单元格里没有文本。", vbExclamation
        Exit Sub
    End If

    ' 把选中文本写入固定的临时文件，供 reader.html 轮询读取
    textFilePath = Environ("TEMP") & "\SublyricReaderText.txt"
    WriteTextFileUTF8 textFilePath, selectedText

    ' 已经有朗读窗口在开着的话，不新开窗口，只更新文件内容，
    ' 页面会在 1 秒内自动检测到变化并替换朗读
    hwnd = FindReaderWindowHandle()
    If hwnd <> 0 Then
        SetForegroundWindow hwnd
        Exit Sub
    End If

    htmlFilePath = ThisWorkbook.Path & "\reader.html"
    If Dir(htmlFilePath) = "" Then
        MsgBox "找不到 reader.html，请确认它和本加载项文件放在同一目录下：" & _
               vbCrLf & htmlFilePath, vbExclamation
        Exit Sub
    End If

    edgePath = FindEdgePath()
    If edgePath = "" Then
        MsgBox "无法找到 Microsoft Edge 浏览器。", vbExclamation
        Exit Sub
    End If

    ' 独立的 Edge 用户配置目录，避免被你平时用的 Edge"单实例"合并、
    ' 导致命令行参数被忽略
    profileDir = Environ("TEMP") & "\SublyricEdgeProfile"
    If Dir(profileDir, vbDirectory) = "" Then
        On Error Resume Next
        MkDir profileDir
        On Error GoTo 0
    End If

    ' 计算贴靠屏幕右侧的窗口位置和大小
    screenW = GetSystemMetrics(SM_CXSCREEN)
    screenH = GetSystemMetrics(SM_CYSCREEN)
    winW = 420
    If winW > screenW Then winW = screenW
    winH = screenH
    winX = screenW - winW
    winY = 0

    url = "file:///" & Replace(htmlFilePath, "\", "/") & _
          "?textfile=" & URLEncodeUTF8(textFilePath) & "&autostart=1"

    shellCommand = """" & edgePath & """ --app=""" & url & """" & _
                   " --user-data-dir=""" & profileDir & """" & _
                   " --window-size=" & winW & "," & winH & _
                   " --window-position=" & winX & "," & winY & _
                   " --allow-file-access-from-files" & _
                   " --autoplay-policy=no-user-gesture-required" & _
                   " --no-first-run --no-default-browser-check"

    Shell shellCommand, vbNormalFocus
End Sub


' 遍历所有可见顶层窗口，标题包含 READER_WINDOW_KEYWORD 的就是朗读窗口
#If VBA7 Then
Function FindReaderWindowHandle() As LongPtr
#Else
Function FindReaderWindowHandle() As Long
#End If
    foundReaderHwnd = 0
    EnumWindows AddressOf EnumWindowsCallback, 0
    FindReaderWindowHandle = foundReaderHwnd
End Function


#If VBA7 Then
Public Function EnumWindowsCallback(ByVal hwnd As LongPtr, ByVal lParam As LongPtr) As Long
#Else
Public Function EnumWindowsCallback(ByVal hwnd As Long, ByVal lParam As Long) As Long
#End If
    Dim length As Long
    Dim buffer As String
    Dim title As String

    If IsWindowVisible(hwnd) = 0 Then
        EnumWindowsCallback = 1 ' 继续枚举
        Exit Function
    End If

    length = GetWindowTextLengthW(hwnd)
    If length = 0 Then
        EnumWindowsCallback = 1
        Exit Function
    End If

    buffer = String(length + 1, vbNullChar)
    GetWindowTextW hwnd, StrPtr(buffer), length + 1
    title = Left$(buffer, length)

    If InStr(title, READER_WINDOW_KEYWORD) > 0 Then
        foundReaderHwnd = hwnd
        EnumWindowsCallback = 0 ' 找到了，停止枚举
    Else
        EnumWindowsCallback = 1 ' 继续枚举
    End If
End Function


Function FindEdgePath() As String
    Dim possiblePaths(2) As String
    Dim i As Integer

    possiblePaths(0) = Environ("PROGRAMFILES(X86)") & "\Microsoft\Edge\Application\msedge.exe"
    possiblePaths(1) = Environ("PROGRAMFILES") & "\Microsoft\Edge\Application\msedge.exe"
    possiblePaths(2) = Environ("LOCALAPPDATA") & "\Microsoft\Edge\Application\msedge.exe"

    For i = 0 To UBound(possiblePaths)
        If Dir(possiblePaths(i)) <> "" Then
            FindEdgePath = possiblePaths(i)
            Exit Function
        End If
    Next i

    FindEdgePath = ""
End Function


' 把文本以 UTF-8（无 BOM）写入指定文件，供 reader.html 用 fetch() 读取
Sub WriteTextFileUTF8(ByVal filePath As String, ByVal content As String)
    Dim srcStream As Object, outStream As Object
    Dim bytes() As Byte
    Dim startIdx As Long
    Dim i As Long

    Set srcStream = CreateObject("ADODB.Stream")
    srcStream.Type = 2 ' 文本流
    srcStream.Charset = "utf-8"
    srcStream.Open
    srcStream.WriteText content
    srcStream.Position = 0
    srcStream.Type = 1 ' 切到二进制读取字节，方便去掉 BOM
    bytes = srcStream.Read
    srcStream.Close

    startIdx = 0
    If (UBound(bytes) - LBound(bytes) + 1) >= 3 Then
        If bytes(0) = &HEF And bytes(1) = &HBB And bytes(2) = &HBF Then
            startIdx = 3 ' 跳过 UTF-8 BOM
        End If
    End If

    Set outStream = CreateObject("ADODB.Stream")
    outStream.Type = 1
    outStream.Open

    If startIdx > 0 Then
        Dim trimmed() As Byte
        ReDim trimmed(0 To UBound(bytes) - startIdx)
        For i = startIdx To UBound(bytes)
            trimmed(i - startIdx) = bytes(i)
        Next i
        outStream.Write trimmed
    Else
        outStream.Write bytes
    End If

    outStream.SaveToFile filePath, 2 ' adSaveCreateOverWrite
    outStream.Close
End Sub


' 把文本转换成 UTF-8 百分号编码，用于拼在 URL 查询参数里
' （VBA 没有内置的 UTF-8 URL 编码函数，借助 ADODB.Stream 转码）
Function URLEncodeUTF8(ByVal strInput As String) As String
    Dim objStream As Object
    Dim bytes() As Byte
    Dim i As Long
    Dim b As Byte
    Dim startIdx As Long
    Dim result As String

    Set objStream = CreateObject("ADODB.Stream")
    objStream.Type = 2 ' 文本流
    objStream.Charset = "utf-8"
    objStream.Open
    objStream.WriteText strInput
    objStream.Position = 0
    objStream.Type = 1 ' 切到二进制流读取字节
    bytes = objStream.Read
    objStream.Close

    ' ADODB.Stream 以 utf-8 写文本时会带 BOM (EF BB BF)，需要跳过
    startIdx = 0
    If (UBound(bytes) - LBound(bytes) + 1) >= 3 Then
        If bytes(0) = &HEF And bytes(1) = &HBB And bytes(2) = &HBF Then
            startIdx = 3
        End If
    End If

    For i = startIdx To UBound(bytes)
        b = bytes(i)
        Select Case b
            Case 48 To 57, 65 To 90, 97 To 122, 45, 46, 95, 126 ' 0-9 A-Z a-z - . _ ~
                result = result & Chr(b)
            Case Else
                result = result & "%" & Right("0" & Hex(b), 2)
        End Select
    Next i

    URLEncodeUTF8 = result
End Function
