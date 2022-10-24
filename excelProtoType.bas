Attribute VB_Name = "Module1"
Sub template()
Attribute template.VB_ProcData.VB_Invoke_Func = " \n14"
'
'   Creating titles, start/end dates
    Dim Title As String
    Dim Prompt As Worksheet
    Dim Timeline As Worksheet
    Dim StartDate As Date
    Dim EndDate As Date
    Set Prompt = ActiveSheet
    StartDate = Prompt.Range("D3").Value
    EndDate = Prompt.Range("D4").Value
    Title = "Timeline for " + Range("D2").Value
    Sheets.Add After:=ActiveSheet
    ActiveSheet.Name = Title
    Set Timeline = ActiveSheet
    With Timeline
    .Range("A1").Value = Title
    .Range("A3").Value = "Start Date"
    .Range("A4").Value = "End Date"
    .Range("B3").Value = StartDate
    .Range("B4").Value = EndDate
    End With
    
'   Making headers for the table

    Dim HeaderCount As Integer
    Dim i As Integer
    Dim Headers As New Collection
    With Headers
    .Add "WHO"
    .Add "NEXT STEPS"
    .Add "DAYS"
    .Add "START"
    .Add "END"
    .Add "GERMS"
    .Add "CLIENT"
    End With
    HeaderCount = Headers.Count
    For i = 1 To HeaderCount
    Range("A8").Offset(0, i) = Headers(i)
    Next i
' Sample code
    Dim Days As Integer
    With Timeline
    .Range("E10").Value = StartDate
    End With
    

' Start the table by making a collection of each step
    Dim Steps As New Collection
    Dim StepCount As Integer
    With Steps
    .Add "Job Start"
    .Add "Internal Review and Revision#1"
    .Add "Client Presentation R1 (Wireframe)"
    .Add "Client Feedback"
    .Add "Creative Development"
    .Add "Internal Review and Revision#2", Key:="InsertR2"
    .Add "Client Presentation Final"
    .Add "Client Feedback"
    End With

' Add 'Who' in table
    Dim Department As New Collection
    With Department
    .Add "CREATIVE"
    .Add "SVC/CREATIVE"
    .Add "SVC/CLIENT"
    .Add "CLIENT"
    .Add "CREATIVE"
    .Add "SVC/CREATIVE", Key:="BefR2"
    .Add "SVC/CLIENT"
    .Add "CLIENT"
    
    End With
    
   ' Add Content framework if selected

    If Prompt.Range("E7").Value = True Then
    Steps.Add "Content Framework", After:=1
    Department.Add "CREATIVE", After:=1
    Else
    Steps.Add "Wireframe", After:=1
    Department.Add "CREATIVE", After:=1
    End If
   
' Add 2nd round of client feedback if necessary
    Dim Rounds As Integer
    Rounds = Prompt.Range("D5")
    If Rounds = 3 Then
    Steps.Add "Client Presentation R2 (Creative)", Key:="R2", After:="InsertR2"
    Steps.Add "Client Feedback", Key:="Feedback", After:="R2"
    Steps.Add "Revision based on client feedback", Key:="Revision", After:="Feedback"
    Steps.Add "Internal Review and Revision final", After:="Revision"
    Department.Add "SVC/CLIENT", Key:="R2SVC", After:="BefR2"
    Department.Add "CLIENT", Key:="ClientR2", After:="R2SVC"
    Department.Add "CREATIVE", Key:="R2Revision", After:="ClientR2"
    Department.Add "SVC/CREATIVE", After:="R2Revision"
    End If
  
    
    
' Put in table

    StepCount = Steps.Count
    For i = 1 To StepCount
    Range("B9").Offset(i, 0) = Department(i)
    Range("C9").Offset(i, 0) = Steps(i)

   


    
' Add color rows

    If Department(i) = "SVC" Or Department(i) = "SVC/CREATIVE" Or Department(i) = "CREATIVE" Then
    Range("B9:H9").Offset(i, 0).Interior.Color = RGB(255, 153, 255)
    ElseIf Department(i) = "SVC/CLIENT" Or Department(i) = "CLIENT" Then
    Range("B9:H9").Offset(i, 0).Interior.Color = RGB(153, 255, 153)
    
    End If


    Next i

    
' Input number of days

    Dim DaysAvailable, Counter, Offset, Value As Integer
    DaysAvailable = EndDate - StartDate
    Counter = 1
    Value = 1
    
    While Not (DaysAvailable = 0)
    If Range("C9").Offset(Counter, 0) = "Job Start" Or Range("B9").Offset(Counter, 0) = "SVC/CLIENT" Then
    Range("D9").Offset(Counter, 0) = 0
    Counter = Counter + 1
    Else
    Range("D9").Offset(Counter, 0) = Value
    DaysAvailable = DaysAvailable - 1
    Counter = Counter + 1
        If Counter = StepCount + 1 And DaysAvailable > 0 Then
        Counter = 1
        Value = Value + 1
        End If
    End If
    Wend

' Calculate end date including holidays
    For i = 1 To StepCount
    Range("E10").Offset(i, 0).FormulaR1C1 = "=R[-1]C[1]"
    Range("F9").Offset(i, 0).Formula = "=WORKDAY(Range("E9").Offset(i, 0),Range("D9").Offset(i, 0),HOLIDAYS.Range("A4:A7"))"
    Next i

End Sub

