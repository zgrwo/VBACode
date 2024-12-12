'矩阵操作函数

Option Explicit
Sub SelectDataRange()

    Dim selectedRange As Range
    Dim varData As Variant, arrData As Variant
    Dim m As Long, n As Long
    
    On Error Resume Next
    Set selectedRange = Application.InputBox("请选中连续的多个单元格作为数据区域", "选择数据区域", Type:=8)
    On Error GoTo 0
    
    If selectedRange Is Nothing Then
        MsgBox "您未选择任何单元格。"
        Exit Sub
    End If
    
    varData = selectedRange.Value
    
    ' 计算行数和列数，并确定数组的维数
    Call ArrayDimensions(varData, m, n)
    
    '如果只有一个元素，则将其转换为二维数组'
    If m = 0 And n = 0 Then
        ReDim arrData(1 To 1, 1 To 1)
        arrData(1, 1) = varData
    Else
        arrData = varData
    End If

    ' 计算最大值，支持M, R, C，M表示遍历整个矩阵，R表示按求取，C表示按列求取
    maxValueOverall = MMax(arrData, "M")

    ' 显示结果
    Dim arrA As Variant
    Dim element As Variant
    arrA = maxValuePerRow
    For Each element In arrA
        Debug.Print element
    Next element
    
End Sub

Sub ArrayDimensions(arrA As Variant, ByRef lngRows As Long, ByRef lngCols As Long)

    If Not IsArray(arrA) Then
        lngRows = 0
        lngCols = 0
    Else
        lngRows = UBound(arrA, 1) - LBound(arrA, 1) + 1
        lngCols = UBound(arrA, 2) - LBound(arrA, 2) + 1
    End If
    
End Sub

Function MMax(arrA As Variant, Optional Method As String = "M") As Variant

    Dim i As Long, j As Long
    Dim maxVal As Variant
       
    Dim lRow As Long, uRow As Long
    Dim lCol As Long, uCol As Long
    lRow = LBound(arrA, 1)
    uRow = UBound(arrA, 1)
    lCol = LBound(arrA, 2)
    uCol = UBound(arrA, 2)
    
    ' 根据Method参数选择不同的计算方式
    Select Case Method
        Case "M" ' 计算整个数组的最大值
            ReDim maxVal(1 To 1, 1 To 1)
            maxVal(1, 1) = arrA(lRow, lCol)
            For i = lRow To uRow
                For j = lCol To uCol
                    If arrA(i, j) > maxVal(1, 1) Then
                        maxVal(1, 1) = arrA(i, j)
                    End If
                Next j
            Next i
        Case "R" ' 按行计算最大值
            ReDim maxVal(1 To uRow - lRow + 1, 1 To 1)
            For i = lRow To uRow
                maxVal(i - lRow + 1, 1) = arrA(i, lCol)
                For j = lCol To uCol
                    If arrA(i, j) > maxVal(i - lRow + 1, 1) Then
                        maxVal(i - lRow + 1, 1) = arrA(i, j)
                    End If
                Next j
            Next i
        Case "C" ' 按列计算最大值
            ReDim maxVal(1 To 1, 1 To uCol - lCol + 1)
            For j = lCol To uCol
                maxVal(1, j - lCol + 1) = arrA(lRow, j)
                For i = lRow To uRow
                    If arrA(i, j) > maxVal(1, j - lCol + 1) Then
                        maxVal(1, j - lCol + 1) = arrA(i, j)
                    End If
                Next i
            Next j
        Case Else
            Debug.Print "不支持的方法，仅支持 M、R 或 C。"
            Exit Function
    End Select
    
    MMax = maxVal
    
End Function

Function MMin(arrA As Variant, Optional Method As String = "M") As Variant

    Dim i As Long, j As Long
    Dim minVal As Variant
       
    Dim lRow As Long, uRow As Long
    Dim lCol As Long, uCol As Long
    lRow = LBound(arrA, 1)
    uRow = UBound(arrA, 1)
    lCol = LBound(arrA, 2)
    uCol = UBound(arrA, 2)
    
    ' 根据Method参数选择不同的计算方式
    Select Case Method
        Case "M" ' 计算整个数组的最小值
            ReDim minVal(1 To 1, 1 To 1)
            minVal(1, 1) = arrA(lRow, lCol)
            For i = lRow To uRow
                For j = lCol To uCol
                    If arrA(i, j) < minVal(1, 1) Then
                        minVal(1, 1) = arrA(i, j)
                    End If
                Next j
            Next i
        Case "R" ' 按行计算最小值
            ReDim minVal(1 To uRow - lRow + 1, 1 To 1)
            For i = lRow To uRow
                minVal(i - lRow + 1, 1) = arrA(i, lCol)
                For j = lCol To uCol
                    If arrA(i, j) < minVal(i - lRow + 1, 1) Then
                        minVal(i - lRow + 1, 1) = arrA(i, j)
                    End If
                Next j
            Next i
        Case "C" ' 按列计算最小值
            ReDim minVal(1 To 1, 1 To uCol - lCol + 1)
            For j = lCol To uCol
                minVal(1, j - lCol + 1) = arrA(lRow, j)
                For i = lRow To uRow
                    If arrA(i, j) < minVal(1, j - lCol + 1) Then
                        minVal(1, j - lCol + 1) = arrA(i, j)
                    End If
                Next i
            Next j
        Case Else
            Debug.Print "不支持的方法，仅支持 M、R 或 C。"
            Exit Function
    End Select
    
    MMin = minVal
    
End Function

Function MAvg(arrA As Variant, Optional Method As String = "M") As Variant

    Dim i As Long, j As Long
    Dim sumVal As Double
    Dim count As Long
    Dim avgVal As Variant
    
    Dim lRow As Long, uRow As Long
    Dim lCol As Long, uCol As Long
    lRow = LBound(arrA, 1)
    uRow = UBound(arrA, 1)
    lCol = LBound(arrA, 2)
    uCol = UBound(arrA, 2)
    
    ' 根据Method参数选择不同的计算方式
    Select Case Method
        Case "M" ' 计算整个数组的平均值
            ReDim avgVal(1 To 1, 1 To 1)
            sumVal = 0
            count = 0
            For i = lRow To uRow
                For j = lCol To uCol
                    sumVal = sumVal + arrA(i, j)
                    count = count + 1
                Next j
            Next i
            If count > 0 Then
                avgVal(1, 1) = sumVal / count
            Else
                avgVal(1, 1) = CVErr(xlErrDiv0)
            End If
        Case "R" ' 按行计算平均值
            ReDim avgVal(1 To uRow - lRow + 1, 1 To 1)
            For i = lRow To uRow
                sumVal = 0
                count = 0
                For j = lCol To uCol
                    sumVal = sumVal + arrA(i, j)
                    count = count + 1
                Next j
                If count > 0 Then
                    avgVal(i - lRow + 1, 1) = sumVal / count
                Else
                    avgVal(i - lRow + 1, 1) = CVErr(xlErrDiv0)
                End If
            Next i
        Case "C" ' 按列计算平均值
            ReDim avgVal(1 To 1, 1 To uCol - lCol + 1)
            For j = lCol To uCol
                sumVal = 0
                count = 0
                For i = lRow To uRow
                    sumVal = sumVal + arrA(i, j)
                    count = count + 1
                Next i
                If count > 0 Then
                    avgVal(1, j - lCol + 1) = sumVal / count
                Else
                    avgVal(1, j - lCol + 1) = CVErr(xlErrDiv0)
                End If
            Next j
        Case Else
            Debug.Print "不支持的方法，仅支持 M、R 或 C。"
            Exit Function
    End Select
    
    MAvg = avgVal
    
End Function

Function MAdd(arrA As Variant, arrB As Variant) As Variant
    Dim rowsA As Long, colsA As Long
    Dim rowsB As Long, colsB As Long
    Dim i As Long, j As Long
    Dim result As Variant
    
    ' 获取矩阵 A 和 B 的维度
    Call ArrayDimensions(arrA, rowsA, colsA)
    Call ArrayDimensions(arrB, rowsB, colsB)
    
    ' 检查矩阵 A 和 B 是否具有相同的维度
    If rowsA <> rowsB Or colsA <> colsB Then
        MSub = CVErr(xlErrValue)
        Exit Function
    End If
    
    ' 初始化结果矩阵
    ReDim result(1 To rowsA, 1 To colsA)
    
    ' 进行矩阵加法
    For i = 1 To rowsA
        For j = 1 To colsA
            result(i, j) = arrA(i, j) + arrB(i, j)
        Next j
    Next i
    
    MAdd = result
End Function

Function MSub(arrA As Variant, arrB As Variant) As Variant
    Dim rowsA As Long, colsA As Long
    Dim rowsB As Long, colsB As Long
    Dim i As Long, j As Long
    Dim result As Variant
    
    ' 获取矩阵 A 和 B 的维度
    Call ArrayDimensions(arrA, rowsA, colsA)
    Call ArrayDimensions(arrB, rowsB, colsB)
    
    ' 检查矩阵 A 和 B 是否具有相同的维度
    If rowsA <> rowsB Or colsA <> colsB Then
        MSub = CVErr(xlErrValue)
        Exit Function
    End If
    
    ' 初始化结果矩阵
    ReDim result(1 To rowsA, 1 To colsA)
    
    ' 进行矩阵减法
    For i = 1 To rowsA
        For j = 1 To colsA
            result(i, j) = arrA(i, j) - arrB(i, j)
        Next j
    Next i
    
    MSub = result
End Function

Function MMul(arrA As Variant, arrB As Variant) As Variant
    Dim rowsA As Long, colsA As Long
    Dim rowsB As Long, colsB As Long
    Dim i As Long, j As Long, k As Long
    Dim result As Variant
    
    ' 获取矩阵 A 和 B 的维度
    Call ArrayDimensions(arrA, rowsA, colsA)
    Call ArrayDimensions(arrB, rowsB, colsB)
    
    ' 检查矩阵 A 的列数是否等于矩阵 B 的行数
    If colsA <> rowsB Then
        MMul = CVErr(xlErrValue)
        Exit Function
    End If
    
    ' 初始化结果矩阵
    ReDim result(1 To rowsA, 1 To colsB)
    
    ' 进行矩阵乘法
    For i = 1 To rowsA
        For j = 1 To colsB
            result(i, j) = 0
            For k = 1 To colsA
                result(i, j) = result(i, j) + arrA(i, k) * arrB(k, j)
            Next k
        Next j
    Next i
    
    MMul = result
End Function

Function MDiv(arrA As Variant, arrB As Variant) As Variant
    Dim rowsA As Long, colsA As Long
    Dim rowsB As Long, colsB As Long
    Dim inverseB As Variant
    Dim result As Variant
    
    ' 获取矩阵 A 和 B 的维度
    Call ArrayDimensions(arrA, rowsA, colsA)
    Call ArrayDimensions(arrB, rowsB, colsB)
    
    ' 检查矩阵 B 是否为方阵
    If rowsB <> colsB Then
        MDiv = CVErr(xlErrValue)
        Exit Function
    End If
    
    ' 计算矩阵 B 的逆矩阵
    inverseB = MInv(arrB)
    
    ' 如果逆矩阵为空，则表示矩阵 B 不可逆
    If IsEmpty(inverseB) Then
        MDiv = CVErr(xlErrValue)
        Exit Function
    End If
    
    ' 进行矩阵乘法 A * B^-1
    result = MMul(arrA, inverseB)
    
    MDiv = result
End Function

Function MInv(matrix As Variant) As Variant
    Dim m as long, n As Long
    Dim det As Double
    Dim adjoint() As Double
    Dim inv() As Double
    Dim i As Long, j As Long
    
    ' 获取矩阵的维度
    Call ArrayDimensions(matrix, m, n)
    
    ' 检查矩阵是否为方阵
    If m <> n Then
        MInv = CVErr(xlErrValue)
        Exit Function
    End If
    
    ' 计算行列式
    det = MDet(matrix)
    
    ' 检查行列式是否为零
    If det = 0 Then
        MInv = CVErr(xlErrValue)
        Exit Function
    End If
    
    ' 计算伴随矩阵
    ReDim adjoint(1 To n, 1 To n)
    For i = 1 To n
        For j = 1 To n
            adjoint(i, j) = Cofactor(matrix, i, j) / det
        Next j
    Next i
    
    ' 转置伴随矩阵得到逆矩阵
    ReDim inv(1 To n, 1 To n)
    For i = 1 To n
        For j = 1 To n
            inv(j, i) = adjoint(i, j)
        Next j
    Next i
    
    MInv = inv
End Function

Function MDet(matrix As Variant) As Double
    Dim m as long, n As Long
    Dim det As Double
    Dim submatrix() As Double
    Dim i As Long, j As Long, sign As Integer
    
    ' 获取矩阵的维度
    Call ArrayDimensions(matrix, m, n)
    
    ' 检查矩阵是否为方阵
    If m <> n Then
        MDet = CVErr(xlErrValue)
        Exit Function
    End If
    
    ' 基本情况：1x1 矩阵
    If n = 1 Then
        MDet = matrix(1, 1)
        Exit Function
    End If
    
    ' 基本情况：2x2 矩阵
    If n = 2 Then
        MDet = matrix(1, 1) * matrix(2, 2) - matrix(1, 2) * matrix(2, 1)
        Exit Function
    End If
    
    ' 递归计算更高阶矩阵的行列式
    det = 0
    sign = 1
    For j = 1 To n
        submatrix = GetMinor(matrix, 1, j)
        det = det + sign * matrix(1, j) * MDet(submatrix)
        sign = -sign
    Next j
    
    MDet = det
End Function

Function Cofactor(matrix As Variant, row As Long, col As Long) As Double
    Dim minor() As Double
    Dim cofact As Double
    
    ' 获取余子矩阵
    minor = GetMinor(matrix, row, col)
    
    ' 计算代数余子式的符号
    If (row + col) Mod 2 = 0 Then
        cofact = MDet(minor)
    Else
        cofact = -MDet(minor)
    End If
    
    Cofactor = cofact
End Function

Function GetMinor(matrix As Variant, row As Long, col As Long) As Double()
    Dim m as long, n As Long
    Dim i As Long, j As Long
    Dim minor() As Double
    Dim mi As Long, mj As Long
    
    ' 获取矩阵的维度
    Call ArrayDimensions(matrix, m, n)
    
    ' 初始化余子矩阵
    ReDim minor(1 To n - 1, 1 To n - 1)
    
    ' 构建余子矩阵
    mi = 1
    For i = 1 To n
        If i = row Then GoTo SkipRow
        mj = 1
        For j = 1 To n
            If j = col Then GoTo SkipCol
            minor(mi, mj) = matrix(i, j)
            mj = mj + 1
SkipCol:
        Next j
        mi = mi + 1
SkipRow:
    Next i
    
    GetMinor = minor
End Function