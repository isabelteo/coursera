Sub AddNumbersA()
'Create a subroutine called AddNumbersA that adds a number input into an input box to the value in cell D4.
'The result should be output in cell G12.
Range("G12") = Range("D4") + InputBox("Enter a number to add to the value of cell D4:")
End Sub


Sub AddNumbersB()
'Create a subroutine called AddNumbersB that asks the user for a number,
 'adds that number to the value in the active cell
 'then places the result in a cell that is 3 rows up and two columns right of the active cell.
 'Make sure that you select an appropriate active cell prior to running the sub!
 Dim x As Double
 x = InputBox("Input a number")
 ActiveCell.Offset(-3, 2) = x + ActiveCell
End Sub

Sub WherePutMe()
'Asks the user for a row number and column letter then places the 2,2 position of a selection into that cell
Dim x As Double, y As String

x = InputBox("Input a row number")
y = InputBox("Input a col letter")
Range(y & CStr(x)) = Selection.Cells(2, 2)
End Sub

Sub Swap()
'Swap the values in two adjacent cells (in the same row, anywhere on the spreadsheet)
Dim temp As Double
temp = Selection.Cells(1, 1)
Selection.Cells(1, 1) = Selection.Cells(1, 2)
Selection.Cells(1, 2) = temp
End Sub
