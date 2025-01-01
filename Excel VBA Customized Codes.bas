'''
    1. 部分自定义函数 / 过程，可一定程度实现新版 Excel 内置函数的功能
    2. 部分自定义函数 / 过程，具有相当程度的使用频率
    3. 作者：Samov Ran / zgrwo@163.com
    4. 不是经过严格测试过的代码，仅供学习和参考
'''
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
            varTemp = varInput.Value2
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