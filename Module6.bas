Attribute VB_Name = "Module6"
'Code in this module is used to control Foxy R2


Dim tbno, state As Integer
Dim dtmstartime, totaltime

Sub CommandButton1_Click()
'The buttong Start Frac activates this sub routine.
'The once activated the sheet is locked to prevent any changes until stopped manual or automatically
Dim endtime As Double
Dim fracint As Double
Dim fractime As Double
Application.Sheets("FoxyCol").Protect , userinterfaceonly:=True

If readtime = 1 Then
 Sheets("FoxyCol").Select
 Cells(4, 2).Value = 0 'setting state value
 Sheets("FoxyCol").Select
 dtmstartime = Now + CDate(Cells(20, 2).Value)
 Cells(2, 2).Value = dtmstartime
 Cells(25, 2).Value = 0
 Cells(26, 2).Value = 0
 Cells(5, 2).Value = 1
 
 'Uses windows schedule event to start and stop the fract collection.
 'This prevents lock up of the VBA while the instrument is running.
 Application.OnTime dtmstartime, "Module6.StopFrac_Click"
 Application.OnTime Now, "Module6.MoveFrac"

End If

End Sub
Sub startfrac()
'This sub routine starts the fraction collection by activating the valve for the set amount of time.
Dim socketId As Integer
Dim Statefxy As String * 5
Dim pno As String * 11


Dim ipAddress As String
Dim port, temp, x As Integer
Dim tbnos As String
Dim com, time1, time2 As String
Sheets("FoxyCol").Select
state = Cells(4, 2).Value
 Cells(25, 2).Value = Cells(25, 2).Value + 1
If state = 0 Then
 Call StartIt
 ipAddress = Cells(12, 2).Value
 port = 23

 socketId = OpenSocket(ipAddress, port)
 'Opens the valve to start collection of fractions
 com = "REMOTE;VALVE=1;RSVP"
 temp = SendCommand(com)
 x = RecvAscii(Statefxy, 5)
 If Statefxy = "READY" Then
   Sheets("FoxyCol").Select
   tbno = Cells(16, 2).Value
   'This creates a running log of when a sample is collected.
   
   If tbno < 145 Then
   Cells(7, tbno).Value = tbno
   Cells(8, tbno).Value = Now()
   Else
   Cells(9, tbno - 144).Value = tbno
   Cells(10, tbno - 144).Value = Now()
   End If
   Cells(16, 2).Value = tbno + 1
 Else
   MsgBox ("R2-Unavalible")
 End If
 Call CloseConnection
 Call EndIt
 dtmstartime = Now + CDate(Cells(22, 2).Value)
  Application.OnTime dtmstartime, "Module6.MoveFrac"
  Cells(3, 2).Value = dtmstartime
 Cells(5, 2).Value = 1
Else
Cells(3, 2).Value = "S:Cleared"
End If

End Sub

Sub MoveFrac()
'This sub Rountine moves the fraction collectors arm to the next tube and waits for fraction collection to start.
Dim socketId As Integer
Dim Statefxy As String * 5
Dim pno As String * 11

Sheets("FoxyCol").Select
state = Cells(4, 2).Value
 Cells(26, 2).Value = Cells(26, 2).Value + 1
 
If state = 0 Then
  
 Sheets("FoxyCol").Select 'Selects sheet donot change name of the sheet
 tbno = Cells(16, 2).Value
 
 'Modify the rack number based on which tube is selected or next.
 If tbno < 145 Then
 rtbno = 1000 + tbno
 Else
 rtbno = 2000 + tbno - 144
 End If
  If tbno < 288 Then
   Call StartIt
   ipAddress = Cells(12, 2).Value
   port = 23
     socketId = OpenSocket(ipAddress, port)
    tbnos = CStr(rtbno)
    com = "REMOTE;VALVE=0;RTUBE=" + tbnos + ";RSVP"
     temp = SendCommand(com)
    x = RecvAscii(Statefxy, 5)
    If Statefxy = "READY" Then
     Else
         'MsgBox ("/" + Statefxy + "/")
            MsgBox ("R2-Unavalible")
     End If

   Call CloseConnection
   Call EndIt
  If tbno = 1 Then
   Application.OnTime Now, "Module6.StartFrac"
  Else
   dtmstartime = Now + CDate(Cells(21, 2).Value)
   Application.OnTime dtmstartime, "Module6.StartFrac"
   Cells(3, 2).Value = dtmstartime
   Cells(5, 2).Value = 0
  
   End If
  Else
  Call stop_run_nomes
  Cells(4, 2).Value = 1
  Sheets("FoxyCol").Select
  tim = CDate(Cells(2, 2).Value)
  Application.OnTime tim, "Module6.StopFrac_Click", , False
  MsgBox ("Run stopped: Ran out of tubes :(")
  
 End If
 Else
  Cells(3, 2).Value = "S:Cleared"
 End If
End Sub
Sub StopFrac_UserClick()
Dim tim As Date

Cells(4, 2).Value = 1 ' Changing state to stop
Application.Sheets("FoxyCol").Unprotect

Call stop_run_nomes
Sheets("FoxyCol").Select


If IsDate(Cells(2, 2).Value) = True Then
 tim = CDate(Cells(2, 2).Value)
 Application.OnTime tim, "Module6.StopFrac_Click", , False
 Cells(2, 2).Value = "Cleared"
End If


If IsDate(Cells(3, 2).Value) = False Then
 If MsgBox("Stopped Successfully" + vbNewLine + "Reset tube Count to 1?", vbOKCancel) = vbOK Then
  Cells(16, 2).Value = 1
  End If
 Exit Sub
End If

nxtcall = Cells(5, 2).Value
tim = CDate(Cells(3, 2).Value)

If nxtcall = 0 Then
 If endApplicaton_ontime(tim, "Module6.startfrac") = True Then
  Cells(3, 2).Value = "Cleared"
 Else
   If endApplicaton_ontime(tim, "Module6.MoveFrac") = False Then
      MsgBox ("Fraction Collection Que Seems Empty, Please wait till" + CStr(tim) + " to Restart")
      Exit Sub
   Else
    Cells(3, 2).Value = "Cleared"
   End If
   
 End If
Else
 If endApplicaton_ontime(tim, "Module6.MoveFrac") = True Then
  Cells(3, 2).Value = "Cleared"
 Else
   If endApplicaton_ontime(tim, "Module6.startfrac") = False Then
      MsgBox ("Fraction Collection Que Seems Empty, Please wait till: " + CStr(tim) + " to Restart")
      Exit Sub
   Else
    Cells(3, 2).Value = "Cleared"
   End If
   
 End If
End If


If MsgBox("Stopped Successfully" + vbNewLine + "Reset tube Count to 1?", vbOKCancel) = vbOK Then
Cells(16, 2).Value = 1
End If

End Sub
Sub StopFrac_Click()
Dim tim As Date

Cells(4, 2).Value = 1 ' Changing state to stop
Application.Sheets("FoxyCol").Unprotect

Call stop_run_nomes
Sheets("FoxyCol").Select
nxtcall = Cells(5, 2).Value
If IsDate(Cells(3, 2).Value) = False Then
 Cells(2, 2).Value = "Done"
  If MsgBox("Stopped Successfully" + vbNewLine + "Reset tube Count to 1?", vbOKCancel) = vbOK Then
   Cells(16, 2).Value = 1
  End If
 Exit Sub
Else
 tim = CDate(Cells(3, 2).Value)
 Cells(2, 2).Value = "Done"
End If
If nxtcall = 0 Then
 If endApplicaton_ontime(tim, "Module6.startfrac") = True Then
  Cells(3, 2).Value = "Cleared"
 Else
   If endApplicaton_ontime(tim, "Module6.MoveFrac") = False Then
      MsgBox ("S:Fraction Collection Que Seems Empty, Please wait till" + CStr(tim) + " to Restart")
      Exit Sub
   Else
    Cells(3, 2).Value = "Cleared"
   End If
   
 End If
Else
 If endApplicaton_ontime(tim, "Module6.MoveFrac") = True Then
  Cells(3, 2).Value = "Cleared"
 Else
   If endApplicaton_ontime(tim, "Module6.startfrac") = False Then
      MsgBox ("M:Fraction Collection Que Seems Empty, Please wait till: " + CStr(tim) + " to Restart")
      Exit Sub
   Else
    Cells(3, 2).Value = "Cleared"
   End If
   
 End If
End If
If MsgBox("Stopped Successfully" + vbNewLine + "Reset tube Count to 1?", vbOKCancel) = vbOK Then
Cells(16, 2).Value = 1
End If
End Sub

Sub stop_run_nomes()
Dim socketId As Integer
Dim Statefxy As String * 5
Dim port, temp, x, options As Integer
Dim ipAddress, com As String



Call StartIt
ipAddress = Cells(12, 2).Value
port = 23
com = "STOP;Home;RSVP"
socketId = OpenSocket(ipAddress, port)
If socketId = 0 Then
Exit Sub
End If

temp = SendCommand(com)



x = RecvAscii(Statefxy, 5)
If Statefxy = "READY" Then
'MsgBox ("Stopped")
Else
MsgBox ("R2-Unavalible")
End If
 Call CloseConnection
 Call EndIt
End Sub

