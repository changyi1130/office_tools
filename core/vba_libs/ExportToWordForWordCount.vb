Sub ExportToWordForWordCount()
    Dim wdApp As Object
    Dim wdDoc As Object
    Dim ws As Worksheet
    Dim dataRange As Range
    Dim lastRow As Long, lastCol As Long
    Dim excelPath As String
    Dim excelName As String
    Dim savePath As String
    Dim logPath As String          ' 错误日志路径
    Dim errorMsg As String         ' 错误信息暂存

    ' 获取当前工作簿的路径和名称
    excelPath = ThisWorkbook.Path
    excelPath = ConvertSharePointUrlToLocalPath(excelPath)
    excelName = Left(ThisWorkbook.Name, InStrRev(ThisWorkbook.Name, ".") - 1)

    ' 构建保存路径
    If excelPath <> "" Then
        savePath = excelPath & "\" & excelName & "-导出文本.docx"
        logPath = excelPath & "\" & excelName & "-导出文本-错误说明.txt"
    Else
        ' 如果工作簿未保存，使用桌面路径
        Dim desktopPath As String
        desktopPath = CreateObject("WScript.Shell").SpecialFolders("Desktop")
        savePath = desktopPath & "\" & excelName & "-导出文本.docx"
        logPath = desktopPath & "\" & excelName & "-导出文本-错误说明.txt"
    End If

    ' 创建Word应用
    On Error Resume Next
    Set wdApp = GetObject(, "Word.Application")
    If Err.Number <> 0 Then
        Set wdApp = CreateObject("Word.Application")
    End If
    On Error GoTo ErrHandler

    ' 如果没有创建成功，静默退出
    If wdApp Is Nothing Then
        Exit Sub
    End If

    ' 创建新文档
    Set wdDoc = wdApp.Documents.Add
    wdApp.Visible = True  ' 可根据需要改为 False

    ' 遍历所有工作表
    For Each ws In ActiveWorkbook.Worksheets
        ' 跳过隐藏工作表
        If ws.Visible = xlSheetVisible Then
            ' 添加工作表名称标题
            wdApp.Selection.TypeText Text:="[" & ws.Name & "]"
            wdApp.Selection.TypeParagraph
            wdApp.Selection.TypeParagraph

            ' 检查工作表是否有内容
            On Error Resume Next
            lastRow = ws.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
            lastCol = ws.Cells.Find("*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
            On Error GoTo ErrHandler

            If lastRow > 0 And lastCol > 0 Then
                Set dataRange = ws.Range("A1", ws.Cells(lastRow, lastCol))

                ' ==== 局部错误捕获：专门处理复制失败（如合并单元格 + 筛选）====
                On Error Resume Next
                dataRange.Copy
                If Err.Number <> 0 Then
                    ' 构造错误信息
                    If Err.Number = 1004 And InStr(1, Err.Description, "合并", vbTextCompare) > 0 Then
                        errorMsg = "工作表 [" & ws.Name & "] 中包含合并单元格，且当前筛选状态导致无法复制。" & vbCrLf & _
                                   "建议取消合并单元格或清除筛选后再试。"
                    Else
                        errorMsg = "工作表 [" & ws.Name & "] 复制时出现未知错误：" & Err.Description
                    End If
                    ' 写入错误日志
                    WriteToLog logPath, errorMsg
                    Err.Clear
                    Application.CutCopyMode = False
                    GoTo SkipPaste
                End If
                On Error GoTo ErrHandler   ' 恢复全局错误处理

                ' 粘贴到Word
                wdApp.Selection.Paste

SkipPaste:
                Application.CutCopyMode = False
                wdApp.Selection.TypeParagraph
            End If

            ' 添加分隔空行
            wdApp.Selection.TypeParagraph
            wdApp.Selection.TypeParagraph
        End If
    Next ws

    ' 保存Word文档
    On Error Resume Next
    wdDoc.SaveAs2 Filename:=savePath, FileFormat:=16
    wdDoc.Close

    ' 退出Word应用
    wdApp.Quit

Cleanup:
    On Error Resume Next
    Set dataRange = Nothing
    Set ws = Nothing
    Set wdDoc = Nothing
    Set wdApp = Nothing
    Exit Sub

ErrHandler:
    ' 全局错误写入日志
    errorMsg = "程序执行过程中发生严重错误：" & Err.Description
    WriteToLog logPath, errorMsg
    Resume Cleanup
End Sub

' ===== 写入错误日志的辅助过程（追加模式） =====
Private Sub WriteToLog(ByVal logFilePath As String, ByVal message As String)
    Dim fNum As Integer
    Dim timestamp As String
    timestamp = Format(Now, "yyyy-mm-dd hh:nn:ss")

    On Error Resume Next
    fNum = FreeFile
    Open logFilePath For Append As #fNum
    If Err.Number = 0 Then
        Print #fNum, "[" & timestamp & "] " & message
        Print #fNum, String(50, "-")   ' 分隔线
        Close #fNum
    Else
        ' 如果打开文件失败，静默忽略
        Err.Clear
    End If
    On Error GoTo 0
End Sub

' ===== 将 SharePoint 在线 URL 转换为本地 OneDrive 同步路径 =====
Private Function ConvertSharePointUrlToLocalPath(ByVal inputPath As String) As String
    ' 定义需要匹配的 SharePoint 前缀（注意末尾带斜杠）
    Const SP_URL_PREFIX As String = "https://bolingtech-my.sharepoint.com/personal/hailong_fu_bolingtech_onmicrosoft_com/Documents/"
    Const LOCAL_PREFIX As String = "D:\OneDrive - boling\"

    Dim result As String
    result = inputPath

    ' 如果输入路径以 http 开头，尝试转换
    If Left(LCase(inputPath), 4) = "http" Then
        ' 检查是否匹配特定的 SharePoint 前缀（不区分大小写）
        If StrComp(Left(inputPath, Len(SP_URL_PREFIX)), SP_URL_PREFIX, vbTextCompare) = 0 Then
            ' 截取前缀后的相对路径部分
            Dim relativePart As String
            relativePart = Mid(inputPath, Len(SP_URL_PREFIX) + 1)

            ' 将 URL 编码的空格（%20）还原为普通空格，将正斜杠转换为反斜杠
            relativePart = Replace(relativePart, "%20", " ")
            relativePart = Replace(relativePart, "/", "\")

            ' 组合成本地路径
            result = LOCAL_PREFIX & relativePart
        End If
    End If

    ConvertSharePointUrlToLocalPath = result
End Function
