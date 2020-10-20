Sub calc_temp()

Dim NN As Integer
Dim NE As Integer
Dim NW As Integer
Dim NT As Integer
Dim N1 As Integer
Dim N2 As Integer
Dim el As Single
Dim A() As Single
Dim B() As Single
Dim X() As Single

COLUMN1 = 4
COLUMN2 = 6

NN = Cells(COLUMN1, 1)
NE = Cells(COLUMN1, 2)
NW = Cells(COLUMN1, 3)
NT = Cells(COLUMN1, 4)

COLUMN3 = COLUMN2 + NE + 2

ReDim A(NN, NN)
ReDim B(NN)
ReDim X(NN)

For i = 1 To NE
    N1 = Cells(i + COLUMN2, 1)
    N2 = Cells(i + COLUMN2, 2)
    el = Cells(i + COLUMN2, 3)
    A(N1, N2) = -el
    A(N2, N1) = -el
    A(N1, N1) = A(N1, N1) + el
    A(N2, N2) = A(N2, N2) + el
Next i

For i = 1 To NT
    num = Cells(i + COLUMN2, 6)
    For j = 1 To NN
        A(num, j) = 0
    Next j
    A(num, num) = 1
    B(num) = Cells(i + COLUMN2, 7)
Next i

For i = 1 To NW
    num = Cells(i + COLUMN2, 4)
    B(num) = Cells(i + COLUMN2, 5)
Next i

For i = 1 To NN
 For j = 1 To NN
   Cells(i + COLUMN3, j) = A(i, j)
 Next j
 Cells(i + COLUMN3 + NN + 1, NN + 2) = B(i)
Next i


End Sub
