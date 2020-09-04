Attribute VB_Name = "Module1"
Option Explicit

Function selectFile() As String

    Dim fd As Office.FileDialog
    Dim file As Variant
    Dim notCancel As Boolean
    Set fd = Application.FileDialog(msoFileDialogOpen)
    With fd
        notCancel = .Show
        selectFile = .SelectedItems(1)
    End With
End Function


Sub convertGrade()
    Dim file As String
    file = selectFile()
    Dim wApp As New Word.Application
    Dim wDoc As Word.Document
    Dim tRow As Long
    
    wApp.Visible = False
    Set wDoc = wApp.Documents.Open(file, False, True)
    
    ' Extract Grades from Transcript
    With wDoc.Tables(1)
        For tRow = 2 To .Rows.Count
            On Error Resume Next
            Cells(tRow - 1, 1).Value = WorksheetFunction.Trim(WorksheetFunction.Clean(.Cell(tRow, 5).Range.Text))
            On Error GoTo 0
        Next tRow
    End With
    wDoc.Close False
    Set wApp = Nothing
    
    
    Dim i As Long
    Dim grades(12) As String
    Dim convertedGrade(12) As Double
    Dim grade As Integer
    Dim gradeSum As Double
    Dim courseNumber As Integer
    
    gradeSum = 0
    courseNumber = 0
    
    'Create an array of grades on a 12.0 Scale
    grades(0) = "F"
    grades(1) = "D-"
    grades(2) = "D"
    grades(3) = "D+"
    grades(4) = "C-"
    grades(5) = "C"
    grades(6) = "C+"
    grades(7) = "B-"
    grades(8) = "B"
    grades(9) = "B+"
    grades(10) = "A-"
    grades(11) = "A"
    grades(12) = "A+"
    'Create an array of grades converted to a 4.0 Scale
    convertedGrade(0) = 0
    convertedGrade(1) = 0.7
    convertedGrade(2) = 1
    convertedGrade(3) = 1.3
    convertedGrade(4) = 1.7
    convertedGrade(5) = 2
    convertedGrade(6) = 2.3
    convertedGrade(7) = 2.7
    convertedGrade(8) = 3
    convertedGrade(9) = 3.3
    convertedGrade(10) = 3.7
    convertedGrade(11) = 3.9
    convertedGrade(12) = 4
    
    'Loop through each grade extracted from the transcript
    For i = 1 To Rows.Count
        For grade = 0 To 12
        'Compare cell to the list of grades in the grades array
            If StrComp(Cells(i, 1).Value, grades(grade)) = 0 Then
            ' Add Converted Grade to gradeSum
                gradeSum = gradeSum + convertedGrade(grade)
                'Increment number of courses
                courseNumber = courseNumber + 1
                Exit For
            End If
        Next
    Next
    'Display GPA
    MsgBox "Your GPA on a 4.0 scale is " & Round((gradeSum / courseNumber), 2)
    
    
        
    
End Sub



