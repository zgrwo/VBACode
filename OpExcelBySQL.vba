' SQLConnection
Sub SQLConnection()

    On Error GoTo ErrorHandler
    
    Dim Conn As Object, Rst As Object
    Dim PathStr As String, StrConn As String, StrSQL As String
    
    ' 创建 ADODB.Connection 和 ADODB.Recordset 对象
    Set Conn = CreateObject("ADODB.Connection")
    Set Rst = CreateObject("ADODB.Recordset")
    
    ' 获取当前工作簿的完整路径
    PathStr = ThisWorkbook.FullName
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("selected")
    On Error GoTo ErrorHandler
    
    ' 检查selected工作表是否存在，如果不存在则新建一个selected的工作表，如果工作表已存在，则清空其内容
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = "selected"
    Else
        ws.Cells.Clear
    End If
    
    ' 根据 Excel 版本选择连接字符串
    Select Case Application.Version * 1
        Case Is <= 11
            StrConn = "Provider=Microsoft.Jet.OLEDB.4.0;Extended Properties='Excel 8.0;HDR=YES';Data Source=" & PathStr
        Case Is >= 12
            StrConn = "Provider=Microsoft.ACE.OLEDB.12.0;Extended Properties='Excel 12.0;HDR=YES';Data Source=" & PathStr
        Case Else
            MsgBox "Excel version not supported", vbCritical
            Exit Sub
    End Select
    
    ' 编辑SQL查询语句
    StrSQL = "SELECT [Material], [Posting Date], [Movement Type], [Material Description], [Quantity], [Unit of Entry], [Amount in LC]" & _
            "FROM [rawdata$]"

    ' 打开连接，并执行' SQL 查询语句
    Conn.Open StrConn
    Set Rst = Conn.Execute(StrSQL)
    
    ' 将记录集复制到工作表
    With ws
        For i = 0 To Rst.Fields.Count - 1
            .Cells(1, i + 1).Value = Rst.Fields.Item(i).Name
        Next i
    End With
    
    If Not Rst.EOF Then
        ThisWorkbook.Sheets("selected").Range("A2").CopyFromRecordset Rst
    End If

    ' 关闭记录集和连接，并执行清理
Cleanup:
    If Not Rst Is Nothing Then Rst.Close
    If Not Conn Is Nothing Then Conn.Close
    Set Conn = Nothing
    Set Rst = Nothing
    Exit Sub

    ' 执行错误处理
ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical
    Resume Cleanup

End Sub

'SQL
Public Const strSQL As String = "SELECT * FROM [Equipment$]"
Public Const File_Path As String = "D:\Workspace\zgrwo\VBA\SQLSimulation\SAP_Rawdata_By_2023_Debug.xlsm"
Public Const strConn As String = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
                                "Data Source=" & File_Path & ";" & _
                                "Extended Properties=""Excel 12.0 Xml;HDR=YES;Mode=Share Deny Write"";"

Function CreateNewConnection(connectionName As String) As Object

    Dim wb As Workbook   
    Dim newConn As Object
    Set wb = Application.ActiveWorkbook

    On Error Resume Next
    Set newConn = wb.Connections.Add( _
        Name:=connectionName, _
        Description:="")
    On Error GoTo 0
    
    If newConn Is Nothing Then
        MsgBox "Connection already exists or could not be created."
        Set CreateNewConnection = Nothing
        Exit Function
    End If
    
    With newConn.OLEDBConnection
        .CommandText = strSQL
        .BackgroundQuery = True
        .RefreshOnFileOpen = False
        .SavePassword = False
        .AdjustColumnWidth = True
        .RefreshStyle = xlInsertDeleteCells
        .RefreshPeriod = 0
        .PreserveColumnInfo = True
        .Connection = strConn
    End With
    
    Set CreateNewConnection = newConn
    
End Function