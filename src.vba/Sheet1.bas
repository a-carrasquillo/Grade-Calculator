'Description of the program: form to record the grades of countless students, allows entering up to 6 evaluations
'and calculates the average based on the percentages assigned to each grade individually.
'Transfer student personal information and grades to Excel.

'Sub-procedure to open the form
Sub openGradeForm()
     UserForm1.Show vbModeless
End Sub

'Sub-procedure to delete all data transferred to Excel
Sub ClearData()
     'Confirms if the user really wants to erase all
     Dim answer As Integer
     answer = MsgBox("Do you really want to erase all?", vbExclamation + vbYesNo + vbDefaultButton2, "Erase All")
          
     If answer = vbYes Then
          'Define the variable i to control the loop
          Dim i As Integer
          'Define the lastRow variable, used to determine the number of rows to be delete it
          Dim lastRow As Integer
          'Search for all the non-empty cells
          lastRow = Application.WorksheetFunction.CountA(Range("A:A"))
          'Repeating loop starts from row 2 to the last row containing data
          For i = 2 To lastRow
               'Clears all the data from one row
               For j = 1 To 25
                    Cells(i, j).ClearContents
               Next j
          Next i
     Else
          MsgBox ("The operation was canceled.")
     End If

     
End Sub