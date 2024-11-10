Option Explicit

' 日志文件路径
Private Const LOG_FILE_PATH As String = "D:\Workspace\zgrwo\VBA\Log\logfile.txt" ' 定义日志文件路径

' 提取字符串中的匹配项
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

    If strPattern = "" Or strSource = "" Then
        RegexExtract = "错误：无效的正则表达式模式或无效的源字符串"
        GoTo Cleanup
    End If

    Dim objRegex As Object
    Set objRegex = InitializeRegex(strPattern, blnIgnoreCase, blnGlobal)

    Dim matches As Object
    Set matches = ExecuteRegex(objRegex, strSource)

    If Not matches Is Nothing Then
        If n = 0 Then
            RegexExtract = MatchesToArray(matches)
        ElseIf n > 0 And n <= matches.Count Then
            RegexExtract = matches(n - 1).Value
        ElseIf n < 0 And Abs(n) <= matches.Count Then
            RegexExtract = matches(matches.Count + n).Value
        Else
            RegexExtract = ""
        End If
    Else
        RegexExtract = ""
    End If
    Call LogRecord("RegexExtract", "", Now)
    GoTo Cleanup

ErrHandler:
    Call LogRecord("RegexExtract", Err.Description, Now)
    RegexExtract = "错误：" & Err.Description

Cleanup:
    Set matches = Nothing
    Set objRegex = Nothing
End Function

' 替换字符串中的匹配项
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

    If strPattern = "" Or strSource = "" Then
        RegexReplace = "错误：无效的正则表达式模式或无效的源字符串"
        GoTo Cleanup
    End If

    Dim objRegex As Object
    Set objRegex = InitializeRegex(strPattern, blnIgnoreCase, blnGlobal)

    Dim matches As Object
    Set matches = ExecuteRegex(objRegex, strSource)

    If Not matches Is Nothing Then
        If n = 0 Then
            RegexReplace = objRegex.Replace(strSource, strReplacement)
        ElseIf n > 0 And n <= matches.Count Then
            Dim match As Object
            Set match = matches(n - 1)
            strSource = Left$(strSource, match.FirstIndex) & strReplacement & Mid$(strSource, match.FirstIndex + match.Length + 1)
            RegexReplace = strSource
        ElseIf n < 0 And Abs(n) <= matches.Count Then
            Set match = matches(matches.Count + n)
            strSource = Left$(strSource, match.FirstIndex) & strReplacement & Mid$(strSource, match.FirstIndex + match.Length + 1)
            RegexReplace = strSource
        Else
            RegexReplace = strSource
        End If
    Else
        RegexReplace = strSource
    End If
    Call LogRecord("RegexReplace", "", Now)
    GoTo Cleanup

ErrHandler:
    Call LogRecord("RegexReplace", Err.Description, Now)
    RegexReplace = "错误：" & Err.Description

Cleanup:
    Set matches = Nothing
    Set objRegex = Nothing
End Function

' 初始化正则表达式对象
Private Function InitializeRegex(ByVal strPattern As String, ByVal blnIgnoreCase As Boolean, ByVal blnGlobal As Boolean) As Object
    Dim objRegex As Object
    Set objRegex = CreateObject("VBScript.RegExp")
    objRegex.Pattern = strPattern
    objRegex.IgnoreCase = blnIgnoreCase
    objRegex.Global = blnGlobal
    Set InitializeRegex = objRegex
End Function

' 执行正则表达式匹配
Private Function ExecuteRegex(objRegex As Object, strSource As String) As Object
    On Error GoTo ErrHandler
    Set ExecuteRegex = objRegex.Execute(strSource)
    Exit Function
ErrHandler:
    Set ExecuteRegex = Nothing
    Call LogRecord("ExecuteRegex", Err.Description, Now)
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

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim logDir As String
    logDir = fso.GetParentFolderName(LOG_FILE_PATH)

    If Not fso.FolderExists(logDir) Then
        fso.CreateFolder (logDir)
    End If

    If Not fso.FileExists(LOG_FILE_PATH) Then
        fso.CreateTextFile LOG_FILE_PATH, False
    End If

    Dim logFile As Object
    Set logFile = fso.OpenTextFile(LOG_FILE_PATH, 8, True) 'ForAppending = 8
    logFile.WriteLine TimeStamp & " - " & FunctionName & " - " & strMessage
    logFile.Close

    Exit Sub

ErrHandler:
    MsgBox "无法打开或写入日志文件: " & LOG_FILE_PATH, vbExclamation
    On Error GoTo 0
End Sub