Attribute VB_Name = "modRibbon"
Option Explicit

' ============================================================
' 韵词字幕 · Excel 字幕工具箱 - 功能区回调模块
' 这里的每个 Ribbon_XXX 过程名，需与 customUI14.xml 中
' 对应按钮的 onAction 属性完全一致（区分大小写建议保持一致）
' ============================================================

Public Sub Ribbon_OnLoad(ribbon As IRibbonUI)
    ' 加载项载入时触发，目前无需特殊处理，保留以便未来扩展
End Sub

Public Sub Ribbon_AddEnglishSpaces(control As IRibbonControl)
    On Error GoTo ErrHandler
    Call AddEnglishSpaces
    Exit Sub
ErrHandler:
    MsgBox "执行失败：" & Err.Description, vbCritical
End Sub

Public Sub Ribbon_RemoveEndingSymbols(control As IRibbonControl)
    On Error GoTo ErrHandler
    Call RemoveEndingSymbols
    Exit Sub
ErrHandler:
    MsgBox "执行失败：" & Err.Description, vbCritical
End Sub

Public Sub Ribbon_FormatSheet(control As IRibbonControl)
    On Error GoTo ErrHandler
    Call FormatSheet
    Exit Sub
ErrHandler:
    MsgBox "执行失败：" & Err.Description, vbCritical
End Sub

Public Sub Ribbon_OpenReaderSidebar(control As IRibbonControl)
    On Error GoTo ErrHandler
    Call OpenReaderSidebar
    Exit Sub
ErrHandler:
    MsgBox "执行失败：" & Err.Description, vbCritical
End Sub

Public Sub Ribbon_ExportBilingualSRT(control As IRibbonControl)
    On Error GoTo ErrHandler
    Call ExportBilingualSRT
    Exit Sub
ErrHandler:
    MsgBox "执行失败：" & Err.Description, vbCritical
End Sub

Public Sub Ribbon_ExportChineseSRT(control As IRibbonControl)
    On Error GoTo ErrHandler
    Call ExportChineseSRT
    Exit Sub
ErrHandler:
    MsgBox "执行失败：" & Err.Description, vbCritical
End Sub

Public Sub Ribbon_ExportToWord(control As IRibbonControl)
    On Error GoTo ErrHandler
    Call ExportBilingualToWord
    Exit Sub
ErrHandler:
    MsgBox "执行失败：" & Err.Description, vbCritical
End Sub
