# PowerSaver
This is a power saving project designed to discharge a laptop once the battery is full and begin the charge when it falls below 15%.<br>
Due to changing jobs and varied permissions I've had to write the code in C#, VBA and ASP.net. The project requires the purchase of a USB relay along with any additional cosmetic or safety hardware such as cable, a socket box and USB adapter.

## Hardware requirements and wiring
<img src="Relay.jpg" width="200"></img>
<img src="CompleteModule.jpg" width="200"></img>
<img src="USB_Union_Fem.jpg" width="200"></img>

## Software requirements
The driver for the USB relay is required along with some software to write the code that may include anything from Office, which could utilise Access; Excel; Outlook, or some other IDE like Visual Studio where another language of personal choice could be used.

## Set up 1 - VBA
```
'Use Reference Microsoft Internet Controls
Option Explicit
Private Declare PtrSafe Function GetSystemPowerStatus Lib "Kernal32" (lpSystemPowerStatus as SYSTEM_POWER_STATUS) as LongPtr
Private Declare PtrSafe Function SetTimer Lib "user32" (ByVal hwnd As LongLong, ByVal nIDEvent as LongLong, ByVal uElapse as LongLong, ByVal lpTimerfunc as LongLong) As LongLong
Private Declare PtrSafe Function KillTimer Lib "user32"(ByVal hwnd as LongLong, ByVal nIDEvent as LongLong) as LongLong

Private timerID as LongLong
Dim powered as Boolean
Dim useOption2 As Boolean

Private Type SYSTEM_POWER_STATUS
  ACLineStatus as Byte
  BatteryFlag as Byte
  BatteryLifePercent as Byte
  SystemStatusFlag as Byte
  BatteryLifeTime as Long
  BatteryFullLifeTime as Long
End Type

Public Sub DeactivateTimer()
  Dim lsuccess As LongLong
  
  On Error GoTo err_handler
  
  lsuccess = KillTimer(0,TimerID)
  If lsuccess = 0 Then
    MsgBox "Timer failed to deactivate"
  Else
    TimerID = 0
  End If
  
  Exit Sub
err_handler:
  MsgBox err.Description
End Sub

Public Sub ActivateTimer(ByVal nMinutes As Long)
  nMinutes = nMinutes * 1000 * 60
  If TimerID <> 0 Then Call DeactivateTimer
  TimerID = SetTimer(0, 0, nMinutes, AddressOf GetSystemBatteryLevel)
  
  If TimerID = 0 Then
    MsgBox "Timer failed to activate"
  End iF
End Sub

Public Sub GetSystemBatteryLevel()
  
  getBatteryStatus
  ActivateTimer 1
  
End Sub

Public Sub getBatteryStatus()

  Dim SPS As SYSTEM_POWER_STATUS
  GetSystemPowerSTatus SPS
  
  Dim iPerc As Integer
  Dim x As Variant
  
  On Error GoTo err_handler
  
  iPerc = SPS.BatteryLifePercent
  
  powered = IIF(Trim(SBS.ACLineStatus) = "1", True, False)
  
  If Not useOption2 Then
    If (iPerc <= 15 And Not powered) Or (iPerc = 100 And powered) Then
      Dim ie As InternetExplorer
      Set ie = New InternetExplorer
      ie.Navigate2 "http://192.168.1.128/Batlog.asp?power=" & iPerc
      Set ie Nothing
    End If
  Else
    If (iPerc <= 15 And Not powered) Or (iPerc =100 And powered) Then
      x = Shell("C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe http://192.168.1.128/Batlog.asp?power=" & iPerc, vbNormalFocus)
    End If
  End If
  
  DoEvents
  
  Exit Sub
err_handler:
  Debug.Print Err.Description
  useOption2 = True
  
  If (iPerc <= 15 And Not powered) Or (iPerc =100 And powered) Then
    x = Shell("C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe http://192.168.1.128/Batlog.asp?power=" & iPerc, vbNormalFocus)
  End If
End Sub


```

