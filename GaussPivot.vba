//VBA code I wrote to find the row reduction of  an arbitrary matrix

//ran as a vba module in a spreadsheet (MS Excel 2007)

//known as the Gauss Pivot

Option Explicit
Option Base 1
Sub Gauss()
  Dim Amat() As Double, bvec() As Double
  Dim i As Integer, j As Integer, k As Integer
  Dim n As Integer
  Dim Eps As Single
  Eps = 1E-06
  n = Range("A").Rows.Count
  ReDim Amat(n, n) As Double
  ReDim bvec(n) As Double
  For i = 1 To n
    For j = 1 To n
      Amat(i, j) = Application.WorksheetFunction.Index(Range("A"), i, j)
    Next j
  bvec(i) = Application.WorksheetFunction.Index(Range("b"), i, 1)
  Next i
  For i = 1 To n
    Call Pivot(Amat, bvec, n, i)
    If Abs(Amat(i, i)) < Eps Then
        MsgBox "Singular Equation Set ==>Cannot be solved"
        Stop
    End If
    For j = i + 1 To n
      Amat(i, j) = Amat(i, j) / Amat(i, i)
    Next j
    bvec(i) = bvec(i) / Amat(i, i)
    For k = i + 1 To n
      For j = i + 1 To n
        Amat(k, j) = Amat(k, j) - Amat(i, j) * Amat(k, i)
      Next j
      bvec(k) = bvec(k) - bvec(i) * Amat(k, i)
    Next k
  Next i
  For i = n To 2 Step -1
    For j = i - 1 To 1 Step -1
      bvec(j) = bvec(j) - Amat(j, i) * bvec(i)
    Next j
  Next i
  Range("b").Select
  ActiveCell.Offset(0, 1).Select
  For i = 1 To n
    ActiveCell.Offset(i - 1, 0).Value = bvec(i)
  Next i
End Sub

Sub Pivot(CoeffMat, ConstVec, n, i)
    Dim ip As Integer, k As Integer, j As Integer
    Dim MaxPiv As Double
    MaxPiv = CoeffMat(i, i)
    ip = 1
    For k = i + 1 To n
        If Abs(CoeffMat(k, i)) > MaxPiv Then
            MaxPiv = Abs(CoeffMat(k, i))
            ip = k
        End If
    Next k
    If ip <> i Then
        For j = i To n
            Call Swap(CoeffMat(i, j), CoeffMat(ip, j))
        Next j
        Call Swap(ConstVec(i), ConstVec(ip))
    End If
End Sub

Sub Swap(a, b)
    Dim Temp As Variant
    Temp = a
    a = b
    b = Temp
End Sub
