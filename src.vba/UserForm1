
'Sub-procedure that allows us to add the information for some elements at the beggining
Private Sub UserForm_Initialize()
     'Allows you to add the universities
     With universities
          .AddItem "University X"
          .AddItem "University Y"
          .AddItem "University Z"
     End With
     
     'Allows you to add the different faculties
     With faculty
          .AddItem "Engineering"
          .AddItem "Education"
          .AddItem "Health Sciences"
          .AddItem "Natural Sciences"
     End With
     
     'Allows you to add different classes
     With class
          .AddItem "General Chemistry"
          .AddItem "Physics 1"
          .AddItem "Precalculus 2"
          .AddItem "Calculus 1"
          .AddItem "Calculus 2"
          .AddItem "Social Sciences"
          .AddItem "Differential Equations"
          .AddItem "Discrete Mathematics"
          .AddItem "Introduction to Engineering"
          .AddItem "Introduction to Programming"
          .AddItem "Advanced Communicative English"
          .AddItem "Intermediate Programming"
          .AddItem "Research and Writing, English"
     End With
     
     'Allows us to add the semesters
     With semester
          .AddItem "2017-01"
          .AddItem "2017-02"
          .AddItem "2018-01"
          .AddItem "2018-02"
          .AddItem "2019-01"
          .AddItem "2019-02"
          .AddItem "2020-01"
          .AddItem "2020-02"
          .AddItem "2021-01"
          .AddItem "2021-02"
     End With
     
     'Allows us to add the number of evaluations
     With numberEvaluations
          .AddItem "4"
          .AddItem "5"
          .AddItem "6"
     End With
     
     'Allows us to add the absences
     With absences
          .AddItem "None"
          .AddItem "Few"
          .AddItem "Many"
     End With
     
End Sub
'Sub-procedure that allows us to transfer the information filled in the form to an Excel worksheet
Private Sub transfer_Click()
     'The message variable is defined which will save the message that we want to display in the MsgBox
     Dim message As String
     'Define the variable that will store the location of the next empty row to put the data
     Dim nextRow As Integer
     'A constant is set for the minimum number of characters for the student number
     Const minLength As Integer = 9

     'Activate worksheet1
     Sheets("Sheet1").Activate
     'Find the next empty row
     nextRow = Application.WorksheetFunction.CountA(Range("A:A")) + 1
     
     'Check if there are any unfilled boxes/fields
     If studentName <> "" And lastNames <> "" And studentNum <> "" And universities <> "" And faculty <> "" And concentration <> "" And class <> "" And semester <> "" And absences <> "" And weight1 <> "" And weight2 <> "" And weight3 <> "" And weight4 <> "" And grade1 <> "" And grade2 <> "" And grade3 <> "" And grade4 <> "" And finalGrade <> "" And finalGradeLetter <> "" Then
          'Evaluate if it does not meet the minimum number of characters for the student number
          If Len(studentNum.Text) < minLength Then
               'Displays a message to correct the error
               MsgBox ("The minimum number of characters in the student number must be 9." & vbNewLine & "Please correct the problem.")
               Exit Sub
          Else
               'Verify that when 5 or 6 grades are used, their corresponding fields are not empty
               Select Case numberEvaluations
                    Case Is = "5":
                         If grade5 = "" Or weight5 = "" Then
                              MsgBox ("Please verify that all fields have been completed")
                              Exit Sub
                         End If
                    Case Is = "6":
                         If grade5 = "" Or weight5 = "" Or grade6 = "" Or weight6 = "" Then
                              MsgBox ("Please verify that all fields have been completed")
                              Exit Sub
                         End If
                End Select
                
               'If all the fields were completed, it will transfer the following:
               'Transfer the student name to column A
               Cells(nextRow, 1) = studentName.Text
               'Transfer the last names to column B
               Cells(nextRow, 2) = lastNames.Text
               'Transfer the student number to column C
               Cells(nextRow, 3) = studentNum.Text
               'Transfer the selected university to column D
               Cells(nextRow, 4) = universities.Value
               'If the user selected the option of Female in Gender, "Female" will be transferred to column E
               If optionFemale.Value = True Then Cells(nextRow, 5) = "Female"
               'If the user selected the Male option in Gender, "Male" will be transferred to column E
               If optionMale.Value = True Then Cells(nextRow, 5) = "Male"
               'If the user selected the Other option in Gender, "Other" will be transferred to column E
               If optionOther.Value = True Then Cells(nextRow, 5) = "Other"
               'Transfer the selected faculty to column F
               Cells(nextRow, 6) = faculty.Text
               'Transfer the selected concentration to column G
               Cells(nextRow, 7) = concentration.Text
               'Transfer the selected class to column H
               Cells(nextRow, 8) = class.Value
               'Transfer the selected semester to column I
               Cells(nextRow, 9) = semester.Value
               'Transfer the selected number of evaluations to column J
               Cells(nextRow, 10) = numberEvaluations.Value
               'Transfer the selected absences to column K
               Cells(nextRow, 11) = absences.Value
               
               'Transfer the first evaluation to column L
               Cells(nextRow, 12) = grade1.Value
               'Transfer the weight/percentage of the first evaluation to column M
               Cells(nextRow, 13) = weight1.Value
               'Transfer the second evaluation to column N
               Cells(nextRow, 14) = grade2.Value
               'Transfer the weight/percentage of the second evaluation to column O
               Cells(nextRow, 15) = weight2.Value
               'Transfer the third evaluation to column P
               Cells(nextRow, 16) = grade3.Value
               'Transfer the weight/percentage of the third evaluation to column Q
               Cells(nextRow, 17) = weight3.Value
               'Transfer the fourth evaluation to column R
               Cells(nextRow, 18) = grade4.Value
               'Transfer the weight/percentage of the fourth evaluation to column S
               Cells(nextRow, 19) = weight4.Value
               
               'Determine the number of evaluations
               Select Case numberEvaluations
                    Case Is = "4":
                         'Since there are only four evaluations, we transfer N/A for the fifth and sixth evaluation and their corresponding weights/percentages
                         Cells(nextRow, 20) = "N/A"
                         Cells(nextRow, 21) = "N/A"
                         Cells(nextRow, 22) = "N/A"
                         Cells(nextRow, 23) = "N/A"
                    Case Is = "5":
                         'Transfer the fifth evaluation to column T
                         Cells(nextRow, 20) = grade5.Value
                         'Transfer the weight/percentage of the fifth evaluation to column U
                         Cells(nextRow, 21) = weight5.Value
                         'Since there are only five evaluations, we transfer N/A for the sixth evaluation and its corresponding weight/percentage
                         Cells(nextRow, 22) = "N/A"
                         Cells(nextRow, 23) = "N/A"
                    Case Is = "6":
                         'Transfer the fifth evaluation to column T
                         Cells(nextRow, 20) = grade5.Value
                         'Transfer the weight/percentage of the fifth evaluation to column U
                         Cells(nextRow, 21) = weight5.Value
                         'Transfer the sixth evaluation to column V
                         Cells(nextRow, 22) = grade6.Value
                         'Transfer the weight/percentage of the sixth evaluation to column W
                         Cells(nextRow, 23) = weight6.Value
               End Select
               'Transfer the final grade in percentage to column X
               Cells(nextRow, 24) = finalGrade.Value
               'Transfers the final grade in letter to column Y
               Cells(nextRow, 25) = finalGradeLetter.Text
               'The message is created
               message = "Student's name and lastname/s: " & studentName.Text & " " & lastNames.Text & vbNewLine
               message = message & "The student number is: " & studentNum.Text & vbNewLine
               message = message & "The grade in percent is: " & finalGrade.Value & "%" & vbNewLine
               message = message & "The letter of the grade is: " & finalGradeLetter.Text
               'The message is displayed in a message box
               MsgBox message
          End If
     Else
          'If any of the fields mentioned in the condition (decision) is not completed, it shows a MsgBox to indicate this problem
          MsgBox ("Please verify that all fields have been completed")
     End If
End Sub

'Sub-procedure used to close the form
Private Sub closeForm_Click()
     'Allows to close the form
     Unload UserForm1
End Sub

'Sub-procedure that helps us show and hide the fields related to the grades and their weight when the number of evaluations changes
Private Sub numberEvaluations_Change()
     'The True means that the element is being shown, the False, means that the element is hidden
     grade1WeightTag.Visible = True
     grade2WeightTag.Visible = True
     grade3WeightTag.Visible = True
     grade4WeightTag.Visible = True
     weight1.Visible = True
     weight2.Visible = True
     weight3.Visible = True
     weight4.Visible = True
     grade1PercTag.Visible = True
     grade2PercTag.Visible = True
     grade3PercTag.Visible = True
     grade4PercTag.Visible = True
     
     'Determine the number of evaluations
     Select Case numberEvaluations.Text
          Case Is = "4"
              grade5WeightTag.Visible = False
              weight5Tag.Visible = False
              grade6WeightTag.Visible = False
              weight6Tag.Visible = False
              weight5.Visible = False
              weight6.Visible = False
              grade5PercTag.Visible = False
              grade6PercTag.Visible = False
              grade5Tag.Visible = False
              grade6Tag.Visible = False
              grade5.Visible = False
              grade6.Visible = False
              SpinButton5.Visible = False
              SpinButton6.Visible = False
              partialFinalGrade5.Visible = False
              partialFinalGrade6.Visible = False
                            
          Case Is = "5"
              grade5WeightTag.Visible = True
              weight5Tag.Visible = True
              grade6WeightTag.Visible = False
              weight6Tag.Visible = False
              weight5.Visible = True
              weight6.Visible = False
              grade5PercTag.Visible = True
              grade6PercTag.Visible = False
              grade5Tag.Visible = True
              grade6Tag.Visible = False
              grade5.Visible = True
              grade6.Visible = False
              SpinButton5.Visible = True
              SpinButton6.Visible = False
              partialFinalGrade5.Visible = True
              partialFinalGrade6.Visible = False
                            
          Case Is = "6"
              grade5WeightTag.Visible = True
              weight5Tag.Visible = True
              grade6WeightTag.Visible = True
              weight6Tag.Visible = True
              weight5.Visible = True
              weight6.Visible = True
              grade5PercTag.Visible = True
              grade6PercTag.Visible = True
              grade5Tag.Visible = True
              grade6Tag.Visible = True
              grade5.Visible = True
              grade6.Visible = True
              SpinButton5.Visible = True
              SpinButton6.Visible = True
              partialFinalGrade5.Visible = True
              partialFinalGrade6.Visible = True
              
          End Select
End Sub

'Sub-procedure that allows us to relate the faculty to the concentration
Private Sub faculty_Change()
     'Define the index that will be used to store the selected index of the faculty
     Dim index As Integer
     'Determine the selected index
     index = faculty.ListIndex
     'Clears the content of the concentration ComboBox
     concentration.Clear
     
     'A case evaluation is started for the variable index
     Select Case index
          'Engineering Option
          Case Is = 0
               'Adds the engineering concentrations
               With concentration
                    .AddItem "Mechanics"
                    .AddItem "Electrical"
                    .AddItem "Industrial and Management"
                    .AddItem "Computers"
                    .AddItem "Civil"
               End With
          'Education Option
          Case Is = 1
               'Adds the education concentrations
               With concentration
                    .AddItem "Preschool"
                    .AddItem "Elementary-Primary Level (K -3)"
                    .AddItem "Elementary - English"
                    .AddItem "Secondary - Biology"
                    .AddItem "Secondary - General Science"
                    .AddItem "Secondary - History"
                    .AddItem "Secondary - English"
                    .AddItem "Secondary - Mathematics"
                    .AddItem "Secondary - Chemistry"
                    .AddItem "Sec. Vocational and Industrial Education"
                    .AddItem "K-12 Special Education"
                    .AddItem "Recreation"
               End With
          'Health Science Option
          Case Is = 2
               'Adds the Health Science concentrations
               With concentration
                    .AddItem "Nutrition and Diet"
                    .AddItem "Speech-Language Therapy"
                    .AddItem "Nursing"
                    .AddItem "Food and Nutrition Management"
               End With
          'Natural Science Option
          Case Is = 3
          'Adds the Natural Science concentrations
               With concentration
                    .AddItem "General Science"
                    .AddItem "Biology"
                    .AddItem "Biotechnology"
                    .AddItem "Chemistry"
                    .AddItem "Medical Technology"
               End With
     End Select
End Sub

'Sub-procedure that allows us to change the value of the first weight tag on the third page.
'It also validates the first weight value.
Private Sub weight1_Change()
     'Validates the weight value
     If weight1.Value > 40 Or weight1.Value < 1 Then
          'If the weight value is greater than 40 or less than 1, the entered value is cleared
          weight1.Value = ""
          'Shows a message box to fix the error
          MsgBox ("The entered value must be greater than 0 and less than 40.")
          'Exit the subprocedure
          Exit Sub
     End If
     'Change the tag caption
     weight1Tag.Caption = weight1.Value & "%"
End Sub

'Sub-procedure that allows us to change the value of the second weight tag on the third page.
'It also validates the second weight value.
Private Sub weight2_Change()
     'Validates the weight value
     If weight2.Value > 40 Or weight2.Value < 1 Then
          'If the weight value is greater than 40 or less than 1, the entered value is cleared
          weight2.Value = ""
          'Shows a message box to fix the error
          MsgBox ("The entered value must be greater than 0 and less than 40.")
          'Exit the subprocedure
          Exit Sub
     End If
     'Change the tag caption
     weight2Tag.Caption = weight2.Value & "%"
End Sub

'Sub-procedure that allows us to change the value of the third weight tag on the third page.
'It also validates the third weight value.
Private Sub weight3_Change()
     'Validates the weight value
     If weight3.Value > 40 Or weight3.Value < 1 Then
          'If the weight value is greater than 40 or less than 1, the entered value is cleared
          weight3.Value = ""
          'Shows a message box to fix the error
          MsgBox ("The entered value must be greater than 0 and less than 40.")
          'Exit the subprocedure
          Exit Sub
     End If
     'Change the tag caption
     weight3Tag.Caption = weight3.Value & "%"
End Sub

'Sub-procedure that allows us to change the value of the fourth weight tag on the third page.
'It also validates the fourth weight value.
Private Sub weight4_Change()
     'Validates the weight value
     If weight4.Value > 40 Or weight4.Value < 1 Then
          'If the weight value is greater than 40 or less than 1, the entered value is cleared
          weight4.Value = ""
          'Shows a message box to fix the error
          MsgBox ("The entered value must be greater than 0 and less than 40.")
          'Exit the subprocedure
          Exit Sub
     End If
     'Change the tag caption
     weight4Tag.Caption = weight4.Value & "%"
End Sub

'Sub-procedure that allows us to change the value of the fifth weight tag on the third page.
'It also validates the fifth weight value.
Private Sub weight5_Change()
     'Validates the weight value
     If weight5.Value > 40 Or weight5.Value < 1 Then
          'If the weight value is greater than 40 or less than 1, the entered value is cleared
          weight5.Value = ""
          'Shows a message box to fix the error
          MsgBox ("The entered value must be greater than 0 and less than 40.")
          'Exit the subprocedure
          Exit Sub
     End If
     'Change the tag caption
     weight5Tag.Caption = weight5.Value & "%"
End Sub

'Sub-procedure that allows us to change the value of the sixth weight tag on the third page.
'It also validates the sixth weight value.
Private Sub weight6_Change()
     'Validates the weight value
     If weight6.Value > 40 Or weight6.Value < 1 Then
          'If the weight value is greater than 40 or less than 1, the entered value is cleared
          weight6.Value = ""
          'Shows a message box to fix the error
          MsgBox ("The entered value must be greater than 0 and less than 40.")
          'Exit the subprocedure
          Exit Sub
     End If
     'Change the tag caption
     weight6Tag.Caption = weight6.Value & "%"
End Sub

'Sub-procedure that allows us to verify the sum of the weights
Private Sub verifyWeights_Click()
     'Define a variable to hold the sum of the weights
     Dim weightSum As Double
     'Check the number of evaluations and sum the weights properly
     Select Case numberEvaluations.Value
          Case Is = "4":
              weightSum = CDbl(weight1.Text) + CDbl(weight2.Text) + CDbl(weight3.Text) + CDbl(weight4.Text)
          Case Is = "5":
               weightSum = CDbl(weight1.Text) + CDbl(weight2.Text) + CDbl(weight3.Text) + CDbl(weight4.Text) + CDbl(weight5.Text)
          Case Is = "6":
               weightSum = CDbl(weight1.Text) + CDbl(weight2.Text) + CDbl(weight3.Text) + CDbl(weight4.Text) + CDbl(weight5.Text) + CDbl(weight6.Text)
    End Select
    'Check the weight sum and display the appropriate message
    Select Case weightSum
         Case Is > 100:
              MsgBox ("The sum of the weights cannot exceed 100%")
         Case Is < 100:
              MsgBox ("Verify the weights of the grades since they do not give 100%, they give: " & weightSum)
         Case Is = 100:
              MsgBox ("The assigned weights give a total of 100%")
              'Show the hidden page of the Grades
              UserForm1.MultiPage1.Pages(2).Visible = True
     End Select
End Sub

'Sub-procedure that allows us to change the value of the first grade using the first spin button
Private Sub SpinButton1_Change()
     'Shows the value of the spin button in the first grade
     grade1.Text = SpinButton1.Value
End Sub

'Sub-procedure that allows us to change the value of the second grade using the second spin button
Private Sub SpinButton2_Change()
     'Shows the value of the spin button in the second grade
     grade2.Text = SpinButton2.Value
End Sub

'Sub-procedure that allows us to change the value of the third grade using the third spin button
Private Sub SpinButton3_Change()
     'Shows the value of the spin button in the third grade
     grade3.Text = SpinButton3.Value
End Sub

'Sub-procedure that allows us to change the value of the fourth grade using the fourth spin button
Private Sub SpinButton4_Change()
     'Shows the value of the spin button in the fourth grade
     grade4.Text = SpinButton4.Value
End Sub

'Sub-procedure that allows us to change the value of the fifth grade using the fifth spin button
Private Sub SpinButton5_Change()
     'Shows the value of the spin button in the fifth grade
     grade5.Text = SpinButton5.Value
End Sub

'Sub-procedure that allows us to change the value of the sixth grade using the sixth spin button
Private Sub SpinButton6_Change()
     'Shows the value of the spin button in the sixth grade
     grade6.Text = SpinButton6.Value
End Sub

'Sub-procedure that allows us to verify that the first grade value is in the correct range
'and adjust the first spin button as appropriate
Private Sub grade1_Change()
     'Evaluate if the entered value is greater than 100 or less than 0
     If grade1.Value > 100 Or grade1.Value < 0 Then
          'If the value exceeds 100 or is inferior to 0, the entered value is cleared
          grade1.Value = ""
          MsgBox ("The entered value must be positive and less than 100.")
          'Exits the sub-procedure
          Exit Sub
     Else 'The entered value is between 0 and 100
          'The button starts to change from the new value
          SpinButton1.Value = grade1.Value
     End If
End Sub

'Sub-procedure that allows us to verify that the second grade value is in the correct range
'and adjust the second spin button as appropriate
Private Sub grade2_Change()
     'Evaluate if the entered value is greater than 100 or less than 0
     If grade2.Value > 100 Or grade2.Value < 0 Then
          'If the value exceeds 100 or is inferior to 0, the entered value is cleared
          grade2.Value = ""
          MsgBox ("The entered value must be positive and less than 100.")
          'Exits the sub-procedure
          Exit Sub
     Else 'The entered value is between 0 and 100
          'The button starts to change from the new value
          SpinButton2.Value = grade2.Value
     End If
End Sub

'Sub-procedure that allows us to verify that the third grade value is in the correct range
'and adjust the third spin button as appropriate
Private Sub grade3_Change()
     'Evaluate if the entered value is greater than 100 or less than 0
     If grade3.Value > 100 Or grade3.Value < 0 Then
          'If the value exceeds 100 or is inferior to 0, the entered value is cleared
          grade3.Value = ""
          MsgBox ("The entered value must be positive and less than 100.")
          'Exits the sub-procedure
          Exit Sub
     Else 'The entered value is between 0 and 100
          'The button starts to change from the new value
          SpinButton3.Value = grade3.Value
     End If
End Sub

'Sub-procedure that allows us to verify that the fourth grade value is in the correct range
'and adjust the fourth spin button as appropriate
Private Sub grade4_Change()
     'Evaluate if the entered value is greater than 100 or less than 0
     If grade4.Value > 100 Or grade4.Value < 0 Then
          'If the value exceeds 100 or is inferior to 0, the entered value is cleared
          grade4.Value = ""
          MsgBox ("The entered value must be positive and less than 100.")
          'Exits the sub-procedure
          Exit Sub
     Else 'The entered value is between 0 and 100
          'The button starts to change from the new value
          SpinButton4.Value = grade4.Value
     End If
End Sub

'Sub-procedure that allows us to verify that the fifth grade value is in the correct range
'and adjust the fifth spin button as appropriate
Private Sub grade5_Change()
     'Evaluate if the entered value is greater than 100 or less than 0
     If grade5.Value > 100 Or grade5.Value < 0 Then
          'If the value exceeds 100 or is inferior to 0, the entered value is cleared
          grade5.Value = ""
          MsgBox ("The entered value must be positive and less than 100.")
          'Exits the sub-procedure
          Exit Sub
     Else 'The entered value is between 0 and 100
          'The button starts to change from the new value
          SpinButton5.Value = grade5.Value
     End If
End Sub

'Sub-procedure that allows us to verify that the sixth grade value is in the correct range
'and adjust the sixth spin button as appropriate
Private Sub grade6_Change()
     'Evaluate if the entered value is greater than 100 or less than 0
     If grade6.Value > 100 Or grade6.Value < 0 Then
          'If the value exceeds 100 or is inferior to 0, the entered value is cleared
          grade6.Value = ""
          MsgBox ("The entered value must be positive and less than 100.")
          'Exits the sub-procedure
          Exit Sub
     Else 'The entered value is between 0 and 100
          'The button starts to change from the new value
          SpinButton6.Value = grade6.Value
     End If
End Sub

'Sub-procedure that will calculate the partial grades, then the final grade in percentage
'and its corresponding letter.
Private Sub determineFinalGrade_Click()
     'Define the variable that will hold the weight of each grade
     Dim weight As Double
     'Define the variable that will hold the sum of partial grades
     Dim sumPartials As Double
     
     'Search the first weight and convert it to decimal equivalent
     weight = CDbl(weight1.Text) / 100
     'Calculates the first partial grade percentage and give the appropriate format
     partialFinalGrade1.Text = Format(CDbl(grade1 * weight), "#,##0.00")
     
     'Search the second weight and convert it to decimal equivalent
     weight = CDbl(weight2.Text) / 100
     'Calculates the second partial grade percentage and give the appropriate format
     partialFinalGrade2.Text = Format(CDbl(grade2 * weight), "#,##0.00")
     
     'Search the third weight and convert it to decimal equivalent
     weight = CDbl(weight3.Text) / 100
     'Calculates the third partial grade percentage and give the appropriate format
     partialFinalGrade3.Text = Format(CDbl(grade3 * weight), "#,##0.00")
     
     'Search the fourth weight and convert it to decimal equivalent
     weight = CDbl(weight4.Text) / 100
     'Calculates the fourth partial grade percentage and give the appropriate format
     partialFinalGrade4.Text = Format(CDbl(grade4 * weight), "#,##0.00")
     
     Select Case numberEvaluations
          Case Is = "4"
               'Sum the four partial grades
               sumPartials = CDbl(partialFinalGrade1.Value) + CDbl(partialFinalGrade2.Value) + CDbl(partialFinalGrade3.Value) + CDbl(partialFinalGrade4.Value)
          Case Is = "5"
               'Search the fifth weight and convert it to decimal equivalent
               weight = CDbl(weight5.Text) / 100
               'Calculates the fifth partial grade percentage and give the appropriate format
               partialFinalGrade5.Text = Format(CDbl(grade5 * weight), "#,##0.00")
               
               'Sum the five partial grades
               sumPartials = CDbl(partialFinalGrade1.Value) + CDbl(partialFinalGrade2.Value) + CDbl(partialFinalGrade3.Value) + CDbl(partialFinalGrade4.Value) + CDbl(partialFinalGrade5.Value)
          Case Is = "6"
               'Search the fifth weight and convert it to decimal equivalent
               weight = CDbl(weight5.Text) / 100
               'Calculates the fifth partial grade percentage and give the appropriate format
               partialFinalGrade5.Text = Format(CDbl(grade5 * weight), "#,##0.00")
               
               'Search the fifth weight and convert it to decimal equivalent
               weight = CDbl(weight6.Text) / 100
               'Calculates the fifth partial grade percentage and give the appropriate format
               partialFinalGrade6.Text = Format(CDbl(grade6 * weight), "#,##0.00")
               
               'Sum the six partial grades
               sumPartials = CDbl(partialFinalGrade1.Value) + CDbl(partialFinalGrade2.Value) + CDbl(partialFinalGrade3.Value) + CDbl(partialFinalGrade4.Value) + CDbl(partialFinalGrade5.Value) + CDbl(partialFinalGrade6.Value)
     End Select
     
     'Set the final result in the final grade field
     finalGrade.Value = Format(sumPartials, "#,##0.00")
     
     'Determine the letter equivalent of the grade
     Select Case finalGrade.Value
          Case 90 To 100: finalGradeLetter.Text = "A"
          Case 80 To 89.99: finalGradeLetter.Text = "B"
          Case 70 To 79.99: finalGradeLetter.Text = "C"
          Case 60 To 69.99: finalGradeLetter.Text = "D"
          Case Is < 59.99: finalGradeLetter.Text = "F"
     End Select
End Sub