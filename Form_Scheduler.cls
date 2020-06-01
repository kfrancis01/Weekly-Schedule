VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Scheduler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Form_Load()

'*Monday Design and Training Schedule*
Dim Mon

Mon = Me.Monday.Value

Me.MonDesign.Form.RecordSource = "SELECT Design.Design, Design.[Scheduled Date] FROM Design WHERE ((Design.[Scheduled Date]= #" & (Mon) & "#))"
Me.MonTraining.Form.RecordSource = "SELECT [Training Schedule].Technician, [Training Schedule].Task, [Training Schedule].Completed, [Training Schedule].[Assignment Date] FROM [Training Schedule] WHERE (([Training Schedule].[Assignment Date] = #" & (Mon) & "#))"

'Tuesday Design and Training Schedule
Dim Tues

Tues = Me.Tuesday.Value

Me.TuesDesign.Form.RecordSource = "SELECT Design.Design, Design.[Scheduled Date] FROM Design WHERE ((Design.[Scheduled Date]= #" & (Tues) & "#))"
Me.TuesTraining.Form.RecordSource = "SELECT [Training Schedule].Technician, [Training Schedule].Task, [Training Schedule].Completed, [Training Schedule].[Assignment Date] FROM [Training Schedule] WHERE (([Training Schedule].[Assignment Date] = #" & (Tues) & "#))"

'Wednesday Design and Training Schedule
Dim Wed

Wed = Me.Wednesday.Value

Me.WedDesign.Form.RecordSource = "SELECT Design.Design, Design.[Scheduled Date] FROM Design WHERE ((Design.[Scheduled Date]= #" & (Wed) & "#))"
Me.WedTraining.Form.RecordSource = "SELECT [Training Schedule].Technician, [Training Schedule].Task, [Training Schedule].Completed, [Training Schedule].[Assignment Date] FROM [Training Schedule] WHERE (([Training Schedule].[Assignment Date] = #" & (Wed) & "#))"

'Thursday Design and Training Schedule
Dim Thurs
 
Thurs = Me.Thursday.Value

Me.ThursDesign.Form.RecordSource = "SELECT Design.Design, Design.[Scheduled Date] FROM Design WHERE ((Design.[Scheduled Date]= #" & (Thurs) & "#))"
Me.ThursTraining.Form.RecordSource = "SELECT [Training Schedule].Technician, [Training Schedule].Task, [Training Schedule].Completed, [Training Schedule].[Assignment Date] FROM [Training Schedule] WHERE (([Training Schedule].[Assignment Date] = #" & (Thurs) & "#))"

'Friday Design and Training Schedule
Dim Fri

Fri = Me.Friday.Value

Me.FriDesign.Form.RecordSource = "SELECT Design.Design, Design.[Scheduled Date] FROM Design WHERE ((Design.[Scheduled Date]= #" & (Fri) & "#))"
Me.FriTraining.Form.RecordSource = "SELECT [Training Schedule].Technician, [Training Schedule].Task, [Training Schedule].Completed, [Training Schedule].[Assignment Date] FROM [Training Schedule] WHERE (([Training Schedule].[Assignment Date] = #" & (Fri) & "#))"

'Saturday Design and Training Schedule
Dim SaT

SaT = Me.Saturday.Value

Me.SatDesign.Form.RecordSource = "SELECT Design.Design, Design.[Scheduled Date] FROM Design WHERE ((Design.[Scheduled Date]= #" & (SaT) & "#))"
Me.SatTraining.Form.RecordSource = "SELECT [Training Schedule].Technician, [Training Schedule].Task, [Training Schedule].Completed, [Training Schedule].[Assignment Date] FROM [Training Schedule] WHERE (([Training Schedule].[Assignment Date] = #" & (SaT) & "#))"

'Sunday Design and Training
Dim Sun

Sun = Me.Sunday.Value

Me.SunDesign.Form.RecordSource = "SELECT Design.Design, Design.[Scheduled Date] FROM Design WHERE ((Design.[Scheduled Date]= #" & (Sun) & "#))"
Me.SunTraining.Form.RecordSource = "SELECT [Training Schedule].Technician, [Training Schedule].Task, [Training Schedule].Completed, [Training Schedule].[Assignment Date] FROM [Training Schedule] WHERE (([Training Schedule].[Assignment Date] = #" & (Sun) & "#))"


' Days of the Week and Technician Schedules '

Dim db As dao.Database
Set db = CurrentDb
Dim qdf As dao.QueryDef
    
''Monday Date
If Weekday(Date) = 2 Then
Monday.ControlSource = "= Date()"
Me.MonTitle.BorderStyle = 4
Me.MonTitle.BorderColor = RGB(255, 153, 0)

'Sunday Technician Schedule

Tuesday.ControlSource = "= Date() + 1"
Wednesday.ControlSource = "= Date() + 2"
Thursday.ControlSource = "= Date() + 3"
Friday.ControlSource = "= Date() + 4"
Saturday.ControlSource = "= Date() + 5"
Sunday.ControlSource = "= Date() + 6"

Set qdf = db.QueryDefs("Sunday")
'Me.SunSch.Form.RecordSource
qdf.sql = IIf(DLookup("[Sunday 1]", "Sunday") = True, _
    "SELECT [Technician Schedule].[Specialized Technician], [Technician Schedule].Supervisor FROM [Technician Schedule] WHERE ((([Technician Schedule].[Sunday 1])=False) AND (([Technician Schedule].[Sunday 2])=True))", _
    "SELECT [Technician Schedule].[Specialized Technician], [Technician Schedule].Supervisor FROM [Technician Schedule] WHERE ((([Technician Schedule].[Sunday 1])=True) AND (([Technician Schedule].[Sunday 2])=False))")


'Tuesday Date
ElseIf Weekday(Date) = 3 Then
Tuesday.ControlSource = "= Date()"
Me.TuesTitle.BorderStyle = 4
Me.TuesTitle.BorderColor = RGB(255, 153, 0)

Monday.ControlSource = "= Date() + 6"
Wednesday.ControlSource = "= Date() + 1"
Thursday.ControlSource = "= Date() + 2"
Friday.ControlSource = "= Date() + 3"
Saturday.ControlSource = "= Date() + 4"
Sunday.ControlSource = "= Date() + 5"

'Monday Technician schedule

Set qdf = db.QueryDefs("Monday")
'Me.MondaySch.Form.RecordSource
qdf.sql = IIf(DLookup("[Monday 1]", "Monday") = True, _
    "SELECT [Technician Schedule].[Specialized Technician], [Technician Schedule].Supervisor FROM [Technician Schedule] WHERE ((([Technician Schedule].[Monday 1])=False) AND (([Technician Schedule].[Monday 2])=True))", _
    "SELECT [Technician Schedule].[Specialized Technician], [Technician Schedule].Supervisor FROM [Technician Schedule] WHERE ((([Technician Schedule].[Monday 1])=True) AND (([Technician Schedule].[Monday 2])=False))")


'Wednesday Date
ElseIf Weekday(Date) = 4 Then
Wednesday.ControlSource = "= Date"
Me.WedTitle.BorderStyle = 4
Me.WedTitle.BorderColor = RGB(255, 153, 0)

Monday.ControlSource = "= Date() + 5"
Tuesday.ControlSource = "= Date() + 6"
Thursday.ControlSource = "= Date() + 1"
Friday.ControlSource = "= Date() + 2"
Saturday.ControlSource = "= Date() + 3"
Sunday.ControlSource = "= Date() + 4"

'Tuesday Technician Schedule
Set qdf = db.QueryDefs("Tuesday")
'Me.TuesdaySch.Form.RecordSource
qdf.sql = IIf(DLookup("[Tuesday 1]", "Tuesday") = True, _
    "SELECT [Technician Schedule].[Specialized Technician], [Technician Schedule].Supervisor FROM [Technician Schedule] WHERE ((([Technician Schedule].[Tuesday 1])=False) AND (([Technician Schedule].[Tuesday 2])=True))", _
    "SELECT [Technician Schedule].[Specialized Technician], [Technician Schedule].Supervisor FROM [Technician Schedule] WHERE ((([Technician Schedule].[Tuesday 1])=True) AND (([Technician Schedule].[Tuesday 2])=False))")


'Thursday Date
ElseIf Weekday(Date) = 5 Then
Thursday.ControlSource = "= Date()"
Me.ThursTitle.BorderStyle = 4
Me.ThursTitle.BorderColor = RGB(255, 153, 0)

Monday.ControlSource = "= Date() + 4"
Tuesday.ControlSource = "= Date() + 5"
Wednesday.ControlSource = "= Date() + 6"
Friday.ControlSource = "= Date() + 1"
Saturday.ControlSource = "= Date() + 2"
Sunday.ControlSource = "= Date() + 3"

'Wednesday Technician Schedule
Set qdf = db.QueryDefs("Wednesday")
'Me.WedSch.Form.RecordSource
qdf.sql = IIf(DLookup("[Wednesday 1]", "Wednesday") = True, _
    "SELECT [Technician Schedule].[Specialized Technician], [Technician Schedule].Supervisor FROM [Technician Schedule] WHERE ((([Technician Schedule].[Wednesday 1])=False) AND (([Technician Schedule].[Wednesday 2])=True))", _
    "SELECT [Technician Schedule].[Specialized Technician], [Technician Schedule].Supervisor FROM [Technician Schedule] WHERE ((([Technician Schedule].[Wednesday 1])=True) AND (([Technician Schedule].[Wednesday 2])=False))")


''Friday Date
ElseIf Weekday(Date) = 6 Then
Friday.ControlSource = "= Date()"
Me.FriTitle.BorderStyle = 4
Me.FriTitle.BorderColor = RGB(255, 153, 0)

Monday.ControlSource = "= Date() + 3"
Tuesday.ControlSource = "= Date() + 4"
Wednesday.ControlSource = "= Date() + 5"
Thursday.ControlSource = "= Date() + 6"
Saturday.ControlSource = "= Date() + 1"
Sunday.ControlSource = "= Date() + 2"

'Thursday Technician Schedule
'Set qdf = db.QueryDefs("Thursday")
'Me.ThursSch.Form.RecordSource
qdf.sql = IIf(DLookup("[Thursday 1]", "Thursday") = True, _
    "SELECT [Technician Schedule].[Specialized Technician], [Technician Schedule].Supervisor FROM [Technician Schedule] WHERE ((([Technician Schedule].[Thursday 1])=False) AND (([Technician Schedule].[Thursday 2])=True))", _
    "SELECT [Technician Schedule].[Specialized Technician], [Technician Schedule].Supervisor FROM [Technician Schedule] WHERE ((([Technician Schedule].[Thursday 1])=True) AND (([Technician Schedule].[Thursday 2])=False))")


''Saturday Date

ElseIf Weekday(Date) = 7 Then
Saturday.ControlSource = "= Date()"
Me.SatTitle.BorderStyle = 4
Me.SatTitle.BorderColor = RGB(255, 153, 0)

Monday.ControlSource = "= Date() + 2"
Tuesday.ControlSource = "= Date() + 3"
Wednesday.ControlSource = "= Date() + 4"
Thursday.ControlSource = "= Date() + 5"
Friday.ControlSource = "= Date() + 6"
Sunday.ControlSource = "= Date() + 1"

'Friday Technician Schedule
Set qdf = db.QueryDefs("Friday")
'Me.FriSch.Form.RecordSource
qdf.sql = IIf(DLookup("[Friday 1]", "Friday") = True, _
    "SELECT [Technician Schedule].[Specialized Technician], [Technician Schedule].Supervisor FROM [Technician Schedule] WHERE ((([Technician Schedule].[Friday 1])=False) AND (([Technician Schedule].[Friday 2])=True))", _
    "SELECT [Technician Schedule].[Specialized Technician], [Technician Schedule].Supervisor FROM [Technician Schedule] WHERE ((([Technician Schedule].[Friday 1])=True) AND (([Technician Schedule].[Friday 2])=False))")


''Sunday Date

ElseIf Weekday(Date) = 1 Then
Sunday.ControlSource = "= Date()"
'Me.SatTitle.BorderStyle = 0
Me.SunTitle.BorderStyle = 4
Me.SunTitle.BorderColor = RGB(255, 153, 0)

Monday.ControlSource = "= Date() + 1"
Tuesday.ControlSource = "= Date() + 2"
Wednesday.ControlSource = "= Date() + 3"
Thursday.ControlSource = "= Date() + 4"
Friday.ControlSource = "= Date() + 5"
Saturday.ControlSource = "= Date() + 6"

'Saturday Technician Schedule
Set qdf = db.QueryDefs("Saturday")
'Me.SatSch.Form.RecordSource
qdf.sql = IIf(DLookup("[Saturday 1]", "Saturday") = True, _
    "SELECT [Technician Schedule].[Specialized Technician], [Technician Schedule].Supervisor FROM [Technician Schedule] WHERE ((([Technician Schedule].[Saturday 1])=False) AND (([Technician Schedule].[Saturday 2])=True))", _
    "SELECT [Technician Schedule].[Specialized Technician], [Technician Schedule].Supervisor FROM [Technician Schedule] WHERE ((([Technician Schedule].[Saturday 1])=True) AND (([Technician Schedule].[Saturday 2])=False))")

End If

End Sub

Private Sub MT_Click()

'Dim db As dao.Database
'Set db = CurrentDb
'Dim qdf As dao.QueryDef

Sunday.ControlSource = "= Date()"
'Me.SatTitle.BorderStyle = 0
Me.SunTitle.BorderStyle = 4
Me.SunTitle.BorderColor = RGB(255, 153, 0)

If DLookup("[Saturday 1]", "Saturday") = True Then
'Saturday Technician Schedule
'Set qdf = db.QueryDefs("Saturday")
Me.SatSch.Form.RecordSource = "SELECT [Technician Schedule].[Specialized Technician], [Technician Schedule].Supervisor FROM [Technician Schedule] WHERE ((([Technician Schedule].[Saturday 1])=False) AND (([Technician Schedule].[Saturday 2])=True))"
'Me.SatSch.Form.RecordSource = IIf(DLookup("[Saturday 1]", "Saturday") = True, _
    "SELECT [Technician Schedule].[Specialized Technician], [Technician Schedule].Supervisor FROM [Technician Schedule] WHERE ((([Technician Schedule].[Saturday 1])=False) AND (([Technician Schedule].[Saturday 2])=True))", _
    "SELECT [Technician Schedule].[Specialized Technician], [Technician Schedule].Supervisor FROM [Technician Schedule] WHERE ((([Technician Schedule].[Saturday 1])=True) AND (([Technician Schedule].[Saturday 2])=False))")
'qdf.sql = IIf(DLookup("[Saturday 1]", "Saturday") = True, _
    "SELECT [Technician Schedule].[Specialized Technician], [Technician Schedule].Supervisor FROM [Technician Schedule] WHERE ((([Technician Schedule].[Saturday 1])=False) AND (([Technician Schedule].[Saturday 2])=True))", _
    "SELECT [Technician Schedule].[Specialized Technician], [Technician Schedule].Supervisor FROM [Technician Schedule] WHERE ((([Technician Schedule].[Saturday 1])=True) AND (([Technician Schedule].[Saturday 2])=False))")
    
Else
Dim Mon

Mon = Me.Monday.Value
DoCmd.OpenForm "Training Assignment", , , "[Training Schedule].[Assignment Date] = # " & (Mon) & " # ", , acDialog

Me.MonDesign.Form.RecordSource = "SELECT Design.Design, Design.[Scheduled Date] FROM Design WHERE ((Design.[Scheduled Date]= #" & (Mon) & "#))"
Me.MonTraining.Form.RecordSource = "SELECT [Training Schedule].Technician, [Training Schedule].Task, [Training Schedule].Completed, [Training Schedule].[Assignment Date] FROM [Training Schedule] WHERE (([Training Schedule].[Assignment Date] = #" & (Mon) & "#))"
End If
End Sub

Private Sub TT_Click()
Dim Tues

Tues = Me.Tuesday.Value
DoCmd.OpenForm "Training Assignment", , , "[Training Schedule].[Assignment Date] = # " & (Tues) & " # ", , acDialog

'Dim Tues

'Tues = Me.Tuesday.Value

Me.TuesDesign.Form.RecordSource = "SELECT Design.Design, Design.[Scheduled Date] FROM Design WHERE ((Design.[Scheduled Date]= #" & (Tues) & "#))"
Me.TuesTraining.Form.RecordSource = "SELECT [Training Schedule].Technician, [Training Schedule].Task, [Training Schedule].Completed, [Training Schedule].[Assignment Date] FROM [Training Schedule] WHERE (([Training Schedule].[Assignment Date] = #" & (Tues) & "#))"
End Sub

Private Sub WT_Click()
Dim Wed

Wed = Me.Wednesday.Value
DoCmd.OpenForm "Training Assignment", , , "[Training Schedule].[Assignment Date] = # " & (Wed) & " # ", , acDialog
End Sub

Private Sub ThT_Click()
Dim Thurs

Thurs = Me.Thursday.Value
DoCmd.OpenForm "Training Assignment", , , "[Training Schedule].[Assignment Date] = # " & (Thurs) & " # ", , acDialog
End Sub

Private Sub FT_Click()
Dim Fri

Fri = Me.Friday.Value
DoCmd.OpenForm "Training Assignment", , , "[Training Schedule].[Assignment Date] = # " & (Fri) & " # ", , acDialog
End Sub

Private Sub SaT_Click()
Dim SaT

SaT = Me.Saturday.Value
DoCmd.OpenForm "Training Assignment", , , "[Training Schedule].[Assignment Date] = # " & (SaT) & " # ", , acDialog
End Sub

Private Sub SuT_Click()
Dim Sun

Sun = Me.Sunday.Value
DoCmd.OpenForm "Training Assignment", , , "[Training Schedule].[Assignment Date] = # " & (Sun) & " # ", , acDialog
End Sub

Private Sub New_Click()
DoCmd.OpenForm "Schedule Update"
End Sub

Private Sub Refresh_Click()
Form.Refresh
End Sub

Private Sub Save_Click()
Me.MonDesign.Locked = True
Me.TuesDesign.Locked = True
Me.WedDesign.Locked = True
Me.ThursDesign.Locked = True
Me.FriDesign.Locked = True
Me.SatDesign.Locked = True
Me.SunDesign.Locked = True
End Sub

Private Sub Edit_Click()
Me.MonDesign.Locked = False
Me.TuesDesign.Locked = False
Me.WedDesign.Locked = False
Me.ThursDesign.Locked = False
Me.FriDesign.Locked = False
Me.SatDesign.Locked = False
Me.SunDesign.Locked = False
End Sub



