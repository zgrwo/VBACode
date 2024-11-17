Sub invA()

    'Disable Screen Updating and Events for status bar application
    Application.EnableEvents = False
    Application.ScreenUpdating = False

    Dim sht As Worksheet
    Dim temp As Range
    Dim nRow, nCol, minD, maxD As Long

    'A array must be 1 based for both dimensions, i.e. elements in the 0 index will be omited and might cause errors
    Dim A
    Dim S, X, Sd, id, AX
    Dim SVD As Collection
    Dim cond, max, min, FNorm, tol As Double
    Dim mType As String

    'Constants
    Set sht = ThisWorkbook.Worksheets("Start")
    tol = sht.Range("tol").Value
    If tol = 0 Then
        MsgBox ("Please input a non-zero tolerance. 1e-16 recommended.")
        Exit Sub
    End If
        
    'Populate array from matrix A
    Set sht = ThisWorkbook.Worksheets("A")
    nRow = sht.Range("A10000").End(xlUp).Row
    nCol = sht.Range("XFD1").End(xlToLeft).Column
    If nRow < 1 And nCol < 1 Then
        MsgBox ("Please input matrix A. Must be at least 1x2 or 2x1")
        Exit Sub
    End If
    maxD = Application.max(nRow, nCol)
    minD = Application.min(nRow, nCol)
    Set temp = sht.Range("A1:" & sht.Cells(nRow, nCol).Address)
    A = temp.Value

    'Estimate condition number from matrix S
    Set SVD = SVDr(A, tol)
    S = SVD("S")
    Sd = DIAG(S)
    max = 0
    min = 1.79E+308
    For i = 1 To minD
        If Abs(Sd(i)) > max Then
            max = Abs(Sd(i))
        End If
        'Zero values excluded from min calculation to avoid div0 error
        If Abs(Sd(i)) < min And Abs(Sd(i)) > 0 Then
            min = Abs(Sd(i))
        End If
    Next i
    cond = max / min
    Debug.Print "Condition number estimate = " & Format(cond, "0.00E+00")
    Set sht = ThisWorkbook.Worksheets("Start")
    sht.Range("cond") = cond

    'Determine type of matrix A and print to spreadsheet
    If cond > 1 / tol And nRow = nCol Then
        mType = "Square with singular values"
    ElseIf cond > 1 / tol And nRow < nCol Then
        mType = "Underdetermined with singular values"
    ElseIf cond > 1 / tol And nRow > nCol Then
        mType = "Overdetermined with singular values"
    ElseIf nRow < nCol Then
        mType = "Underdetermined"
    ElseIf nRow > nCol Then
        mType = "Overdetermined"
    ElseIf nRow = nCol Then
        mType = "Square"
    End If
    sht.Range("mtype") = mType

    'Calculate pseudo-inverse matrix X
    X = PInv(A, tol)

    'A matrix multiplied by its inverse should equal the identity matrix
    AX = MultArrays(A, X)
    id = EYE(maxD)

    'Check closeness of AX and id with the Forbenius Norm
    FNorm = 0
    For j = 1 To minD
        For i = 1 To minD
            FNorm = FNorm + Abs(AX(i, j) - id(i, j)) ^ 2
        Next i
    Next j
    FNorm = Sqr(FNorm)
    Debug.Print "Forbenius Norm(AX,I) = " & Format(FNorm, "0.00E+00")
    Set sht = ThisWorkbook.Worksheets("Start")
    sht.Range("fnorm") = FNorm

    'Print matrix to sheet
    Set sht = ThisWorkbook.Worksheets("X")
    sht.Select
    sht.UsedRange.ClearContents
    For j = 1 To nRow
        For i = 1 To nCol
            sht.Cells(i, j).Value = X(i, j)
        Next i
    Next j

    'Enable Screen Updating and Events for status bar application
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.StatusBar = ""

End Sub

Public Function PInv(ByVal A As Variant, ByVal tol As Double) As Variant
    'Calculate the pseudoinverse of variance-covariance matrix using the Moore-Penrose method. The pseudoinverse is used _
        to compute a best fit (least squares) solution.
    'This function produces a matrix X of the same dimensions as A' so that A*X*A = A, X*A*X = X and A*X and X*A _
        are Hermitian. The computation is based on SVD(A) and any singular values less than a tolerance are treated as zero.
        
    Dim U, S, Sd, V, X
    Dim i, j
    Dim SVD As Collection
    Dim nCol, nRow, r1, minD, maxD As Long

    nRow = UBound(A, 1)
    nCol = UBound(A, 2)
    minD = Application.min(nRow, nCol)
    maxD = Application.max(nRow, nCol)

    Set SVD = SVDr(A, tol)
    U = SVD("U")
    S = SVD("S")
    V = SVD("V")
    Sd = DIAG(S)

    'Calculate number of singular values that are greater than the tolerance
    r1 = 0
    For i = 1 To minD
        If Sd(i) > tol Then
            r1 = r1 + 1
        End If
    Next i
    Debug.Print "Rank(Sigma) = " & r1
    Debug.Print minD - r1 & " singular value(s) excluded."
    Dim sht As Worksheet
    Set sht = ThisWorkbook.Worksheets("Start")
    sht.Range("rank") = r1
    sht.Range("sing") = minD - r1

    'Remove singular values from V, U matrices and Sd vector
    ReDim Preserve V(1 To nCol, 1 To r1)
    ReDim Preserve U(1 To nRow, 1 To r1)
    ReDim Preserve Sd(1 To r1)

    'Invert Sd vector elements
    For i = 1 To r1
        Sd(i) = 1 / Sd(i)
    Next i

    'Populate pseudo-inverse matrix X
    For j = 1 To r1
        For i = 1 To nCol
            V(i, j) = V(i, j) * Sd(j)   'Sd row vector in this case
        Next i
    Next j

    ReDim X(1 To nCol, 1 To nRow)
    X = MultArrays(V, TransposeArray(U))

    PInv = X

End Function

Public Function SVDr(ByVal A As Variant, ByVal tol As Double) As Object
    'Singular Value Decomposition (SVD) of matrix A Returns components U, S & V such that A = USVt
        
    Dim U, S, Sd, V, Q, R, errM, errV
    'U Orthonormal matrix
    'S Diagonal matrix
    'V Orthormal matrix
    'Q Orthonormal matrix from QR Factorisation
    'R Upper triangular matrix from QR Factorisation
    'errM Upper triangular error matrix
    'errV Error vector
    Dim QR As Collection
    Dim nRow, nCol, minD, maxD, i, j, ind As Long
    Dim loopCount, loopMax As Long
    Dim err As Double
    Dim E, F As Double

    'Square matrix assumed, i.e. numCols = numRows
    nRow = UBound(A, 1)
    nCol = UBound(A, 2)
    minD = Application.min(nRow, nCol)
    maxD = Application.max(nRow, nCol)

    ReDim errV(1 To minD ^ 2)
    loopMax = (Application.max(nCol, nRow)) * 100
    loopCount = 0
    err = 1.79E+308
    U = EYE(nRow)
    S = TransposeArray(A)
    V = EYE(nCol)

    Do While loopCount < loopMax And err > tol

        Set QR = QR_GS(TransposeArray(S))
        Q = QR("Q")
        S = QR("R")
        U = MultArrays(U, Q)

        Set QR = QR_GS(TransposeArray(S))
        Q = QR("Q")
        S = QR("R")
        V = MultArrays(V, Q)
        
        errM = TRIU(S, 1)   'S square matrix at this point
        'Convert errM to column vector
        For j = 1 To minD
            For i = 1 To minD
                ind = i + (j - 1) * minD
                errV(ind) = errM(i, j)
            Next i
        Next j
        E = NORM(errV)
        F = NORM(DIAG(S))
        If F = 0 Then: F = 1
        err = E / F
        
        loopCount = loopCount + 1
        Application.StatusBar = "SVD Iterations " & loopCount & " of " & loopMax
    Loop

    'Adjust signs in S matrix and remove near zero entries (i.e. entries not on diagonal)
    Sd = DIAG(S)
    S = ZEROS(nRow, nCol)
    For j = 1 To minD
        S(j, j) = Abs(Sd(j))
        If Sd(j) < 0 Then
            'Adjust signs of U matrix
            For i = 1 To maxD
                U(i, j) = -U(i, j)
            Next i
        End If
    Next j

    Set SVDr = New Collection
    SVDr.add Item:=U, Key:="U"
    SVDr.add Item:=S, Key:="S"
    SVDr.add Item:=V, Key:="V"

End Function

Public Function QR_GS(A As Variant) As Object
    'QR Factorisation via Gram-Schmidt orthoganolisation and least squares Method modified from Timothy Sauer - Numerical Analysis 2E _
        [This method suffers from subtractive cancellation]

    Dim Q, R, Aj, Y, RijQi
    'Q Orthonormal matrix
    'R Upper triangular matrix
    'Aj Column vectors of matrix A
    'Y Partial basis set
    'RijQi Intermediate dot product
    Dim i, j, k
    Dim sum As Double   'Running summation
    Dim nRow, nCol As Long
        
    nRow = UBound(A, 1)
    nCol = UBound(A, 2)

    ReDim Q(1 To nRow, 1 To nCol)
    ReDim R(1 To nCol, 1 To nCol)
    ReDim Y(1 To nRow, 1 To nCol)
    ReDim Aj(1 To nRow)
    ReDim RijQi(1 To nRow)

    Q = ZEROS(nRow, nCol)
    R = ZEROS(nCol, nCol)

    'i represents row number
    'j represents column number
    For j = 1 To nCol
        
        For k = 1 To nRow
            Aj(k) = A(k, j)
            Y(k, j) = A(k, j)
        Next k
        
        For i = 1 To j - 1
            'Rij = Qi . Aj
            R(i, j) = 0
            'For k = LB To UB: R(i, j) = R(i, j) + Q(k, i) * Aj(k): Next k   'Classical method
            For k = 1 To nRow: R(i, j) = R(i, j) + Q(k, i) * Y(k, j): Next k 'Modified method
            'RijQi = Rij * Qi
            For k = 1 To nRow: RijQi(k) = R(i, j) * Q(k, i): Next k
            'Yj = Yj - RijQi
            For k = 1 To nRow: Y(k, j) = Y(k, j) - RijQi(k): Next k
   
        Next i
        
        R(j, j) = 0
        For i = 1 To nRow: R(j, j) = R(j, j) + Y(i, j) ^ 2: Next i
        R(j, j) = Sqr(R(j, j))
        
        'Qj = Yj / Rjj
        For i = 1 To nRow
            If R(j, j) = 0 Then
                Q(i, j) = 0
            Else
                Q(i, j) = Y(i, j) / R(j, j)
            End If
        Next i

    Next j

    Set QR_GS = New Collection
    QR_GS.add Item:=Q, Key:="Q"
    QR_GS.add Item:=R, Key:="R"

End Function

Public Function EYE(ByVal n As Long) As Variant
    'Return nxn identity matrix

    If n < 1 Then
        MsgBox ("Identity must have at least 1 element")
        Exit Function
    End If

    Dim id() As Variant
    Dim j, k As Long

    ReDim id(1 To n, 1 To n)
    For j = 1 To n
        For k = 1 To n
            If j = k Then: id(j, k) = 1: Else: id(j, k) = 0
        Next k
    Next j

    EYE = id
    
End Function

Public Function ZEROS(ByVal nRow As Long, ByVal nCol As Long) As Variant
    'Return nRow x nCol matrix of zeros

    If nRow = 0 Or nCol = 0 Then
        MsgBox ("Matrix must have at least 1 element")
        Exit Function
    End If

    Dim Z() As Variant
    Dim j, k As Long

    ReDim Z(1 To nRow, 1 To nCol)
    For j = 1 To nCol
        For k = 1 To Row
            Z(j, k) = 0
        Next k
    Next j

    ZEROS = Z

End Function

Public Function NORM(ByVal yi As Variant) As Double
    'Returns the length of vector of yi

    Dim sum As Double   'Running summation
    Dim L As Double 'Vector length

    'Calculate vector length
    sum = 0
    For i = 1 To UBound(yi)
        sum = sum + yi(i) * yi(i)
    Next i
    L = Sqr(sum)

    NORM = L

End Function

Public Function TRIU(ByVal A As Variant, ByVal n As Integer) As Variant
    'Returns the upper triangular part of matrix A Elements on or above the nth diagonal of A are returned _
        where n=0 corresponds to the main diagonal
        
    Dim nRow, nCol As Long
    nRow = UBound(A, 1)
    nCol = UBound(A, 2)

    Dim i, j
    Dim tri() As Double
    ReDim tri(1 To nRow, 1 To nCol)

    For j = 1 To nCol
        For i = 1 To nRow
            If i <= j - n Then
                tri(i, j) = A(i, j)
            Else
                tri(i, j) = 0
            End If
        Next i
    Next j


    TRIU = tri

End Function

Public Function DIAG(ByVal A As Variant) As Variant
    'returns a column vector of the main diagonal elements of A

    Dim nRow, nCol, num As Long
    nRow = UBound(A, 1)
    nCol = UBound(A, 2)
    num = Application.min(nRow, nCol)

    Dim i, j
    Dim dia() As Double
    ReDim dia(1 To num)

    For j = 1 To nCol
        For i = 1 To nRow
            If i = j Then
                dia(j) = A(i, j)
            End If
        Next i
    Next j

    DIAG = dia

End Function

Public Function TransposeArray(ByVal myarray As Variant) As Variant
    'This function will transpose a 1-based, 2-dimensional variant array

    Dim X As Long
    Dim Y As Long
    Dim Xupper As Long
    Dim Yupper As Long
    Dim tempArray As Variant

    Xupper = UBound(myarray, 2)
    Yupper = UBound(myarray, 1)

    ReDim tempArray(1 To Xupper, 1 To Yupper)

    For X = 1 To Xupper
        For Y = 1 To Yupper
            tempArray(X, Y) = myarray(Y, X)
        Next Y
    Next X

    TransposeArray = tempArray
    
End Function

Public Function MultArrays(ByVal myarray1 As Variant, ByVal myarray2 As Variant) As Variant
    'This function will return the matrix multiplication of two 1-based, 2-dimensional arrays of compatible size

    Dim X As Long
    Dim Y As Long
    Dim Z As Long
    Dim X1upper, X2upper As Long
    Dim Y1upper, Y2upper As Long
    Dim tempArray As Variant
    Dim temp As Double

    Y1upper = UBound(myarray1, 1)   'num rows
    X1upper = UBound(myarray1, 2)   'num columns
    Y2upper = UBound(myarray2, 1)   'num rows
    X2upper = UBound(myarray2, 2)   'num columns

    If X1upper <> Y2upper Then
        MsgBox ("Incompatible array sizes.")
        Exit Function   'Incompatible matrix sizes
    End If

    ReDim tempArray(1 To Y1upper, 1 To X2upper)

    For Y = 1 To Y1upper
        For X = 1 To X2upper
            For Z = 1 To X1upper
                temp = temp + myarray1(Y, Z) * myarray2(Z, X)
            Next Z
            tempArray(Y, X) = temp
            temp = 0
        Next X
    Next Y

    MultArrays = tempArray
    
End Function