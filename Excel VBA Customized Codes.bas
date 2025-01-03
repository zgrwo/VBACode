Option Explicit

Public Const DELIMITER_CODE As Integer = 1

Function CFILTER(ByVal dataSource As Variant, ByVal expIncludes As Variant, Optional byCol As Boolean = False, Optional ByRef valEmpty As Variant = Null) As Variant

    ' dataSource 数据源
    ' expIncludes 筛选表达式
    ' byCol 是否按列筛选，默认为：False，表示按行筛选
    ' valEmpty 未匹配到时返回的值，默认为：Null
    ' CFILTER 返回结果

    On Error GoTo ErrorHandler

    Dim arrSource As Variant, criIncludes As Variant
    Dim rowsSource As Long, colsSource As Long, rowsExp As Long, colsExp As Long
    Dim matchCount As Long, i As Long, j As Long, k As Long
    Dim results As Variant

    ' 检查输入参数是否缺失
    If IsMissing(dataSource) Or IsMissing(expIncludes) Then
        CFILTER = CVErr(xlErrValue) ' 参数不匹配
        Exit Function
    End If

    ' 处理输入数据
    arrSource = ProcessValue(dataSource)
    GetArrayDimensions arrSource, rowsSource, colsSource

    ' 处理条件表达式
    criIncludes = ProcessValue(expIncludes)
    GetArrayDimensions criIncludes, rowsExp, colsExp

    ' 检查维度是否匹配
    If byCol Then
        If colsSource <> colsExp Then
            CFILTER = CVErr(xlErrValue) ' 维度不匹配
            Exit Function
        End If
    Else
        If rowsSource <> rowsExp Then
            CFILTER = CVErr(xlErrValue) ' 维度不匹配
            Exit Function
        End If
    End If

    ' 计算匹配数量
    matchCount = 0
    If byCol Then
        For j = LBound(criIncludes, 2) To UBound(criIncludes, 2)
            matchCount = matchCount + criIncludes(1, j)
        Next j
    Else
        For i = LBound(criIncludes, 1) To UBound(criIncludes, 1)
            matchCount = matchCount + criIncludes(i, 1)
        Next i
    End If

    ' 筛选数据
    If byCol Then
        ReDim results(1 To rowsSource, 1 To matchCount)
        k = 0
        For j = LBound(arrSource, 2) To UBound(arrSource, 2)
            If criIncludes(1, j) Then
                k = k + 1
                For i = LBound(arrSource, 1) To UBound(arrSource, 1)
                    results(i, k) = arrSource(i, j)
                Next i
            End If
        Next j
    Else
        ReDim results(1 To matchCount, 1 To colsSource)
        k = 0
        For i = LBound(arrSource, 1) To UBound(arrSource, 1)
            If criIncludes(i, 1) Then
                k = k + 1
                For j = 1 To colsSource
                    results(k, j) = arrSource(i, j)
                Next j
            End If
        Next i
    End If

    CFILTER = results
    Exit Function

ErrorHandler:
    MsgBox "Error: Source - " & Err.Source & ", Number - " & Err.Number & ", Description - " & Err.Description, vbCritical, "Error"
    CFILTER = CVErr(xlErrValue) ' 发生错误时返回#VALUE!
End Function

Function CUNIQUE(arr As Variant, Optional by_col As Boolean = False, Optional exactly_once As Boolean = False) As Variant

    ' 功能：从输入数据中提取唯一值，并保持维度与原数据一致
    ' arr 数据源，（1个字符串|1个单元格区域|1-2维数组）
    ' by_col 是否按列提取唯一值，默认为：False，表示按行提取
    ' exactly_once 是否提取仅出现一次的值，默认为：False，表示将提取不重复值，如果出现多次，则保留一次
    ' 返回值：提取唯一值后的二维数组或错误值

    On Error GoTo ErrorHandler
    
    ' 输入验证
    If IsEmpty(arr) Or IsNull(arr) Then
        CUNIQUE = CVErr(xlErrValue)
        Exit Function
    End If
    
    ' 初始化字典
    Dim uniqueItems As Object
    Set uniqueItems = CreateObject("Scripting.Dictionary")
    
    ' 处理输入数据
    Dim data As Variant
    data = ProcessValue(arr)
    
    ' 检查数据维度
    Dim m As Long, n As Long
    m = UBound(data, 1) - LBound(data, 1) + 1
    n = UBound(data, 2) - LBound(data, 2) + 1
    
    ' 如果是单列且按列处理，直接返回原数据
    If by_col And n = 1 Then
        CUNIQUE = data
        Exit Function
    End If
    
    ' 按列处理时，转置数据
    If by_col Then
        data = Application.Transpose(data)
    End If
    
    ' 定义分隔符
    Dim delimiter As String
    delimiter = Chr$(DELIMITER_CODE)
    
    ' 统计每行或每列的出现次数
    Dim i As Long, j As Long
    Dim key As Variant
    For i = LBound(data, 1) To UBound(data, 1)
        key = Join(Application.index(data, i, 0), delimiter) ' 将行或列拼接为字符串作为 Key
        If Not uniqueItems.exists(key) Then
            uniqueItems.Add key, 1 ' 初始化出现次数为 1
        Else
            uniqueItems(key) = uniqueItems(key) + 1 ' 增加出现次数
        End If
    Next i
    
    ' 准备输出数组
    Dim output() As Variant
    Dim count As Long
    count = 0
    
    ' 遍历字典，统计符合条件的项数
    For Each key In uniqueItems.keys
        If Not exactly_once Or uniqueItems(key) = 1 Then
            count = count + 1
        End If
    Next key
    
    ' 如果没有任何符合条件的值，返回错误值
    If count = 0 Then
        ReDim output(1 To 1, 1 To 1)
        output(1, 1) = CVErr(xlErrNA)
        CUNIQUE = output
        Exit Function
    End If
    
    ' 重新定义输出数组的大小
    ReDim output(1 To count, 1 To UBound(data, 2) + 1) ' 最后一列用于存储出现次数
    
    ' 填充输出数组
    Dim rowIndex As Long
    rowIndex = 1
    For Each key In uniqueItems.keys
        If Not exactly_once Or uniqueItems(key) = 1 Then
            ' 拆分 Key 为数组，填充到输出数组
            Dim values As Variant
            values = Split(key, delimiter)
            For j = LBound(values) To UBound(values)
                output(rowIndex, j + 1) = values(j)
            Next j
            ' 最后一列存储出现次数
            output(rowIndex, UBound(data, 2) + 1) = uniqueItems(key)
            rowIndex = rowIndex + 1
        End If
    Next key
    
    ' 按列处理时，转置输出数组
    If by_col Then
        output = Application.Transpose(output)
    End If
    
    CUNIQUE = output
    Exit Function
    
ErrorHandler:
    MsgBox "Error: Source - " & Err.Source & ", Number - " & Err.Number & ", Description - " & Err.Description, vbCritical, "Error"
    CUNIQUE = CVErr(xlErrValue) ' 发生错误时返回#VALUE!
End Function

Function GetArrayDimension(varData As Variant) As Integer

    ' 用于计算 varData 的维度
    ' 单个值时，返回 0
    ' 一维向量，返回 1
    ' 二维数组，返回 2
    ' 其余维度，返回 -1
    
    On Error GoTo ErrorHandler
    
    If Not IsArray(varData) Then
        GetArrayDimension = 0
        Exit Function
    End If

    Dim dimension As Integer
    Dim currentDim As Integer
    Dim lb As Integer
    dimension = 0
    currentDim = 1
    
    ' 获取当前数据的维度
    On Error Resume Next
    Do
        lb = LBound(varData, currentDim)
        If Err.Number <> 0 Then
            Exit Do
        Else
            dimension = currentDim
            currentDim = currentDim + 1
            Err.Clear
        End If
    Loop
    On Error GoTo 0

    ' 根据维度数决定返回值
    Select Case dimension
        Case 0 To 2
            GetArrayDimension = dimension
        Case Else
            GetArrayDimension = -1
    End Select
    
    Exit Function

ErrorHandler:
    MsgBox "Error: Source - " & Err.Source & ", Number - " & Err.Number & ", Description - " & Err.Description, vbCritical, "Error"
End Function

Function ProcessValue(varInput As Variant) As Variant

    ' 用于将 varInput 转化为一个二维数组
    ' 单个值时，返回 1 x 1
    ' 一维向量，返回 1 x n
    ' 二维数组，返回 m x n
    ' 其余维度，返回 xlErrRef，不支持的数据类型
    
    On Error GoTo ErrorHandler

    Dim varTemp As Variant
    Dim varCount As Long

    ' 检查SourceData是否为连续对象, 如果时连续的Range范围，则返回其值，否则报错
    If IsObject(varInput) Then
        If TypeName(varInput) = "Range" Then
            If varInput.Areas.count > 1 Then
                ProcessValue = CVErr(xlErrRef)
                Exit Function
            End If
            varTemp = varInput.value
        Else
            ProcessValue = CVErr(xlErrRef)
            Exit Function
        End If
    Else
        varTemp = varInput
    End If

    ' 获取该数据的维度
    Dim dimension As Integer
    dimension = GetArrayDimension(varTemp)

    ' 将数据转化为二维数组
    Select Case dimension
        Case 0 ' 单个值，转化成1 x 1的二维数组
            ProcessValue = Array(Array(varTemp))
        Case 1                                  ' 一维向量，转成1 x n的二维数组
            Dim i As Long
            Dim lowerBound As Long
            Dim upperBound As Long
            lowerBound = LBound(varTemp)
            upperBound = UBound(varTemp)
            varCount = upperBound - lowerBound + 1
            Dim arr As Variant
            ReDim arr(1 To 1, 1 To varCount)
            For i = lowerBound To upperBound
                arr(1, i - lowerBound + 1) = varTemp(i)
            Next i
            ProcessValue = arr
        Case 2                                  ' 二维数组，保持不变
            ProcessValue = varTemp
        Case Else                               ' 其他，返回不支持的数据类型
            ProcessValue = CVErr(xlErrRef)
    End Select

    Exit Function

ErrorHandler:
    MsgBox "Error: Source - " & Err.Source & ", Number - " & Err.Number & ", Description - " & Err.Description, vbCritical, "Error"
End Function

Sub GetArrayDimensions(arrSource As Variant, ByRef m As Long, ByRef n As Long, Optional Standardization As Boolean = True)

    ' 获取数组的行数和列数
    ' 如果不是二维数组，则尝试将其转化为二维数组，再返回其行数和列数
    ' arrSource 输入数据
    ' m：该数组的行数
    ' n：该数组的列数
    ' Standardization，可选，默认表示输入数据已经是二维数组
    On Error GoTo ErrorHandler
    
    Dim result As Variant
    
    ' 如果不是二维数组，则将其转化为二维数组
    If Not Standardization Then
        result = ProcessValue(arrSource)
    Else
        result = arrSource
    End If
    
    ' 获取行数和列数
    m = UBound(result, 1) - LBound(result, 1) + 1
    n = UBound(result, 2) - LBound(result, 2) + 1
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error: Source - " & Err.Source & ", Number - " & Err.Number & ", Description - " & Err.Description, vbCritical, "Error"
End Sub