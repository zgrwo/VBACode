Option Explicit

Private Const LOG_FILE_PATH As String = "D:\Workspace\zgrwo\VBA\Log\logfile.txt" ' 定义日志文件路径

Function RegexExtract(strSource As String, strPattern As String, Optional n As Long = 0, _
                      Optional blnIgnoreCase As Boolean = True, Optional blnGlobal As Boolean = True) As Variant
    ' strSource 为待匹配的字符串
    ' strPattern 为正则表达式模式
    ' n 为匹配的索引：
    '   默认为0，表示返回所有匹配
    '   n>0表示从1开始，例如1表示第一个匹配项
    '   n<0表示从后往前数，例如-1表示倒数第一个匹配项
    ' blnIgnoreCase 为是否忽略大小写，默认为True
    ' blnGlobal 为是否全局搜索，默认为True
    
    On Error GoTo ErrHandler

    ' 验证 pattern 是否为空
    If strPattern = "" Or strSource = "" Then
        RegexExtract = "错误：无效的正则表达式模式或无效的源字符串"
        GoTo Cleanup
    End If

    ' 创建或重用正则表达式对象
    Static objRegex As Object
    If objRegex Is Nothing Then
        Set objRegex = CreateObject("VBScript.RegExp")
    End If
    
    ' 设置正则表达式模式
    objRegex.Pattern = strPattern
    objRegex.IgnoreCase = blnIgnoreCase
    objRegex.Global = blnGlobal

    ' 执行正则表达式匹配
    Dim matches As Object
    Set matches = objRegex.Execute(strSource)

    ' 处理匹配结果
    If matches.Count > 0 Then
        If n = 0 Then
            ' 返回所有匹配
            RegexExtract = MatchesToArray(matches)
        ElseIf n > 0 And n <= matches.Count Then
            ' 返回第n个匹配
            RegexExtract = matches(n - 1).Value
        ElseIf n < 0 And Abs(n) <= matches.Count Then
            ' 返回倒数第n个匹配
            RegexExtract = matches(matches.Count + n).Value
        Else
            ' 如果n超出了匹配数量，则返回空字符串
            RegexExtract = ""
        End If
    Else
        ' 没有找到匹配
        RegexExtract = ""
    End If
    Call LogRecord("RegexExtract", "", Now)
    GoTo Cleanup

ErrHandler:
    ' 记录错误日志
    Call LogRecord("RegexExtract", Err.Description, Now)
    RegexExtract = "错误：" & Err.Description

Cleanup:
    ' 清理资源
    Set matches = Nothing
    Set objRegex = Nothing
End Function

Function RegexReplace(strSource As String, strPattern As String, strReplacement As String, _
                      Optional n As Long = 0, Optional blnIgnoreCase As Boolean = True, _
                      Optional blnGlobal As Boolean = True) As String
    ' strSource 为待替换的字符串
    ' strPattern 为正则表达式模式
    ' strReplacement 为替换字符串
    ' n 为匹配的索引：
    '   默认为0，表示全局替换
    '   n>0表示从1开始，例如1表示第一个匹配项
    '   n<0表示从后往前数，例如-1表示倒数第一个匹配项
    ' blnIgnoreCase 为是否忽略大小写，默认为True
    ' blnGlobal 为是否全局搜索，默认为True
    
    On Error GoTo ErrHandler

    ' 验证 pattern 是否为空
    If strPattern = "" Or strSource = "" Then
        RegexReplace = "错误：无效的正则表达式模式或无效的源字符串"
        GoTo Cleanup
    End If

    ' 创建或重用正则表达式对象
    Static objRegex As Object
    If objRegex Is Nothing Then
        Set objRegex = CreateObject("VBScript.RegExp")
    End If
    
    ' 设置正则表达式模式
    objRegex.Pattern = strPattern
    objRegex.IgnoreCase = blnIgnoreCase
    objRegex.Global = blnGlobal

    ' 执行正则表达式匹配
    Dim matches As Object
    Set matches = objRegex.Execute(strSource)

    ' 处理匹配结果
    If matches.Count > 0 Then
        If n = 0 Then
            ' 全局替换
            RegexReplace = objRegex.Replace(strSource, strReplacement)
        ElseIf n > 0 And n <= matches.Count Then
            ' 替换第n个匹配
            Dim match As Object
            Set match = matches(n - 1)
            strSource = Left$(strSource, match.FirstIndex) & strReplacement & Mid$(strSource, match.FirstIndex + match.Length + 1)
            RegexReplace = strSource
        ElseIf n < 0 And Abs(n) <= matches.Count Then
            ' 替换倒数第n个匹配
            Set match = matches(matches.Count + n)
            strSource = Left$(strSource, match.FirstIndex) & strReplacement & Mid$(strSource, match.FirstIndex + match.Length + 1)
            RegexReplace = strSource
        Else
            ' 如果n超出了匹配数量，则返回原字符串
            RegexReplace = strSource
        End If
    Else
        ' 没有找到匹配
        RegexReplace = strSource
    End If
    Call LogRecord("RegexReplace", "", Now)
    GoTo Cleanup

ErrHandler:
    ' 记录错误日志
    Call LogRecord("RegexReplace", Err.Description, Now)
    RegexReplace = "错误：" & Err.Description

Cleanup:
    ' 清理资源
    Set matches = Nothing
    Set objRegex = Nothing
End Function

' 将匹配结果转换成字符串数组
Function MatchesToArray(objMatches As Object) As String()
    Dim result() As String
    ReDim result(objMatches.Count - 1)
    Dim i As Long
    For i = 0 To objMatches.Count - 1
        result(i) = objMatches(i).Value
    Next i
    MatchesToArray = result
End Function

' 记录日志到指定文件
Sub LogRecord(FunctionName As String, strMessage As String, TimeStamp As Date)
    On Error GoTo ErrHandler

    If strMessage = "" Then
        strMessage = "已成功执行 " & FunctionName & "."
    End If

    ' 创建 FileSystemObject
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' 获取日志文件的目录
    Dim logDir As String
    logDir = fso.GetParentFolderName(LOG_FILE_PATH)

    ' 检查目录是否存在，如果不存在则创建
    If Not fso.FolderExists(logDir) Then
        fso.CreateFolder (logDir)
    End If

    ' 检查日志文件是否存在，如果不存在则创建
    If Not fso.FileExists(LOG_FILE_PATH) Then
        Dim LogFile As Object
        Set LogFile = fso.CreateTextFile(LOG_FILE_PATH, False) ' 不覆盖现有文件
    End If

    ' 打开日志文件并追加日志信息
    Open LOG_FILE_PATH For Append As #1
    Print #1, TimeStamp & " - " & FunctionName & " - " & strMessage
    Close #1

    Exit Sub

ErrHandler:
    ' 如果打开或写入日志文件时发生错误，显示消息框
    MsgBox "无法打开或写入日志文件: " & LOG_FILE_PATH, vbExclamation
    On Error GoTo 0
End Sub