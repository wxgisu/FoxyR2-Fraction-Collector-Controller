Attribute VB_Name = "Module7"
Public Function readtime() As Integer
Dim endtime As String
Dim fracint As String
Dim fractime As String
Dim timint() As String
Dim tottime As Long
Dim inttime As Long

endtime = Cells(13, 2).Value
fracint = Cells(14, 2).Value
fractime = Cells(15, 2).Value

If tmconv(fracint) Or tmconv(fractime) Or tmconv(endtime) = TimeValue("00:00:00") Then
 MsgBox ("time must be non zero value")
 readtime = 0
Else

 Cells(21, 2).Value = tmconv(fracint) - tmconv(fractime)
 Cells(22, 2).Value = tmconv(fractime)
 
 
 tottime = ttime(endtime)
 inttime = ttime(fracint)
 If (tottime Mod inttime) = 0 Then
  Cells(20, 2).Value = tmconv(endtime) + Cells(21, 2).Value / 2
 Else
  Cells(20, 2).Value = tmconv(endtime)
 End If
 
 readtime = 1
End If

End Function

Public Function tmconv(tmtemp As String) As Date

Dim day  As Integer
Dim hr As Integer
Dim min As Integer
Dim sec As Integer
Dim dt As String

Dim tm() As String
Dim tottime As Date

tm = Split(tmtemp, ":")


   hr = CInt(tm(0))
   day = hr / 24
   hr = hr Mod 24
   min = CInt(tm(1))
   sec = CInt(tm(2))
   
    tottime = TimeValue("00:00:00")
    While day > 0
        tottime = tottime + TimeValue("23:00:00") + TimeValue("1:00:00")
        day = day - 1
    Wend
    dt = CStr(hr) + ":" + CStr(min) + ":" + CStr(sec)
    tottime = tottime + TimeValue(dt)
    tmconv = tottime


End Function

Public Function ttime(tmval As String) As Long
  Dim tm() As String
  Dim day  As Long
Dim hr As Long
Dim min As Long
Dim sec As Long
  Dim ttot As Long
   tm = Split(tmval, ":")
   hr = CInt(tm(0))
   day = hr / 24
   hr = hr Mod 24
   min = CInt(tm(1))
   sec = CInt(tm(2))
   ttot = 0
 ttot = ttot + (day * 86400) + (hr * 3600) + (min * 60) + sec
 ttime = ttot
End Function

Public Function cleanoulet()
Dim socketId As Integer
Dim Statefxy As String * 5
Call StartIt
   ipAddress = Cells(12, 2).Value
   port = 23
     socketId = OpenSocket(ipAddress, port)
        com = "REMOTE;TUBE=1;VALVE=1;RSVP"
     temp = SendCommand(com)
    x = RecvAscii(Statefxy, 5)
    If Statefxy = "READY" Then
     Else
         'MsgBox ("/" + Statefxy + "/")
            MsgBox ("R2-Unavalible")
     End If

   Call CloseConnection
   Call EndIt
   dtmstartime = Now + TimeSerial(0, 10, 0)
   Application.OnTime dtmstartime, "Module6.stop_run_nomes"
   Cells(2, 7).Value = dtmstartime
End Function
Sub outletCleanSetup()
If (MsgBox("Make sure the run is done and place a beaker at Tube1, remove Rack1", vbOKCancel) = vbOK) Then
Call cleanoulet
End If

End Sub

Public Function endApplicaton_ontime(endtime As Date, funcname As String) As Boolean

On Error GoTo errHandler
Application.OnTime endtime, funcname, , False
endApplicaton_ontime = True

Exit Function
errHandler:
If Err = 1004 Then
'MsgBox ("Specified Application Que not found")
Else
MsgBox ("Sorry this error cannot be handled: " + Err.Description)
End If
endApplicaton_ontime = False
End Function

Sub Stop_Cleanout()
Dim tim As Date
tim = CDate(Cells(2, 7).Value)
If endApplicaton_ontime(tim, "Module6.stop_run_nomes") = True Then
Cells(2, 7).Value = "Cleared"
Call Module6.stop_run_nomes
Else
MsgBox ("Still in que")
End If
End Sub


