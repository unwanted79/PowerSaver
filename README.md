# PowerSaver
This is a power saving project designed to discharge a laptop once the battery is full and begin the charge when it falls below 15%.<br>
Due to changing jobs and varied permissions I've had to write the code in C#, VBA and ASP.net. The project requires the purchase of a USB relay along with any additional cosmetic or safety hardware such as cable, a socket box and USB adapter.

- ## <a href="hardware.md">Hardware requirements and wiring</a>

## Software requirements
The driver for the USB relay is required along with some software to write the code that may include anything from Office, which could utilise Access; Excel; Outlook, or some other IDE like Visual Studio where another language of personal choice could be used.

## Stage 1 - VBA on your work/company PC
```
'Use Reference Microsoft Internet Controls
Option Explicit
'Declare the API's for x64 bit (PtrSafe)
Private Declare PtrSafe Function GetSystemPowerStatus Lib "Kernal32" (lpSystemPowerStatus as SYSTEM_POWER_STATUS) as LongPtr
Private Declare PtrSafe Function SetTimer Lib "user32" (ByVal hwnd As LongLong, ByVal nIDEvent as LongLong, ByVal uElapse as LongLong, ByVal lpTimerfunc as LongLong) As LongLong
Private Declare PtrSafe Function KillTimer Lib "user32"(ByVal hwnd as LongLong, ByVal nIDEvent as LongLong) as LongLong

Private timerID as LongLong
Dim powered as Boolean 'flag represent the application of power
Dim useOption2 As Boolean 'Alternate internet call should the first one fail if flagged by I.T. limits/violations

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
  'This would be where you run the code from using F5 or you could call this sub from a startup routine
  getBatteryStatus
  ActivateTimer 1
  
End Sub

Public Sub getBatteryStatus()

  Dim SPS As SYSTEM_POWER_STATUS
  GetSystemPowerSTatus SPS
  
  Dim iPerc As Integer
  Dim x As Variant
  
  On Error GoTo err_handler
  
  'record the battery percentage
  iPerc = SPS.BatteryLifePercent
  
  'Pass the mains powered status to our variable
  powered = IIF(Trim(SBS.ACLineStatus) = "1", True, False)
  
  'If the first option hasn't failed yet then continue in the prefered way
  If Not useOption2 Then
    'Has the powered level dropped below 15% on battery OR reached 100% when powered
    If (iPerc <= 15 And Not powered) Or (iPerc = 100 And powered) Then
      Dim ie As InternetExplorer
      Set ie = New InternetExplorer
      'send the percent to our internal webserver. This then sends the data to a database to be monitored by another program
      'This assumes, as was the case for me, that I.T limitations prevent you from installing any drivers or software that we would need
      'to either interact with the Relay or the database (ODBC Driver etc)
      ie.Navigate2 "http://192.168.1.128/Batlog.asp?power=" & iPerc
      Set ie Nothing
    End If
  Else
    If (iPerc <= 15 And Not powered) Or (iPerc =100 And powered) Then
      'I had limited attempts at using the InternetExplorer class before I.T. raises alarm bells.
      'With that in mind the alternate solution is to use Edge but this a last resort as it's a bit clumsy having edge pop up on the screen
      'whilst you're working, but at least it prevents the battery from going flat if you're not keeping an eye on it.
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

## Stage 2 - Set Up SQL Table
This assumes you already have an Azure SQL database set up, if not you can download the SQL Server engine and host it locally.<br>
Set up a simple table to store the data

```
CREATE TABLE [dbo].[WORK_BATTERY](
	[Percent] [tinyint] NULL,
	[Charge] [bit] NULL
) ON [PRIMARY]
GO

```

## Stage 3 - ASP (Classic to keep things simple)
Assuming you have IIS installed on Windows and have enabled it to process ASP pages you can add the following code to the file named Batlog.asp<br>
This is required for me because I couldn't install the ODBC Driver on my work PC.

```
<% 

Dim power
power=Request.QueryString("power")
response.write(power)

    Dim con
	Set con = Server.CreateObject("ADODB.Connection")
    With con
    '   .CursorLocation = adUseClient
    '    .Mode = adModeRead
        .ConnectionString = "Driver={ODBC Driver 18 for SQL Server};Server=tcp:[your azure SQL Server].database.windows.net,1433;Database=[Database Name];Uid=[your username];Pwd=[your password];Encrypt=yes;TrustServerCertificate=no;Connection Timeout=30;"
        .Open
    End With
	
	con.Execute "UPDATE WORK_BATTERY SET [Percent]=" & power
	con.close
	set con = nothing
	
 %>


```

## Stage 4 - Set up the relay
You'll need to install the drivers for the relay. I've attached them to this project. See <a href="CH341SER.zip">CH341SER.zip</a><br>
Right click the start menu on your home PC and select Device Manager<br>
Expand the Ports (COM &amp; LPT) node<br>
Right click the USB-SERIAL CH340 item and select Properties. This assumes you have connected the USB Relay to your computer.<br>
Navigate to the Port Settings tab and check the settings are the same as illustrated below.<br> 
<img src="Device Settings.JPG" width=800></img>

## Stage 5 - Set up the VBA code to monitor the database and toggle the relay. In this example I am using Excel
```
'include reference Microsoft ActiveX Data Objects 6.1 Library
Public Type DCB
  DCBlength As Long
  BaudRate As Long
  fBitFields As Long
  wReserved As Integer
  XonLim As Integer
  XoffLim As Integer
  ByteSize As Byte
  Parity As Byte
  StopBits As Byte
  XonChar As Byte
  XoffChar As Byte
  ErrorChar As Byte
  EofChar As Byte
  EvtChar As Byte
  wReserved1 As Integer
End Type

' The structure of the fBitFields field.
' FieldName             Bit #     Description
' -----------------     -----     ------------------------------
' fBinary                 1       Windows does not support nonbinary mode transfers, so this member must be =1.
' fParity                 2       If =1, parity checking is performed and errors are reported
' fOutxCtsFlow            3       If =1 and CTS is turned off, output is suspended until CTS is sent again.
' fOutxDsrFlow            4       If =1 and DSR is turned off, output is suspended until DSR is sent again.
' fDtrControl             5,6     DTR flow control (2 bits)
' fDsrSensitivity         7       The driver ignores any bytes received, unless the DSR modem input line is high.
' fTXContinueOnXoff       8       XOFF continues Tx
' fOutX                   9       If =1, TX stops when the XoffChar character is received and starts again when the XonChar character is received.
' fInX                   10       Indicates whether XON/XOFF flow control is used during reception.
' fErrorChar             11       Indicates whether bytes received with parity errors are replaced with the character specified by the ErrorChar.
' fNull                  12       If =1, null bytes are discarded when received.
' fRtsControl            13,14    RTS flow control (2 bits)
' fAbortOnError          15       If =1, the driver terminates all I/O operations with an error status if an error occurs.
' fDummy2                16       reserved

'---------fBitFields-------------
Public Const F_BINARY = 1
Public Const F_PARITY = 2
Public Const F_OUTX_CTS_FLOW = 4
Public Const F_OUTX_DSR_FLOW = 8

'DTR Control Flow Values.
Public Const F_DTR_CONTROL_ENABLE = &H10
Public Const F_DTR_CONTROL_HANDSHAKE = &H20

Public Const F_DSR_SENSITIVITY = &H40
Public Const F_TX_CONTINUE_ON_XOFF = &H80
Public Const F_OUT_X = &H100
Public Const F_IN_X = &H200
Public Const F_ERROR_CHAR = &H400
Public Const F_NULL = &H800

'RTS Control Flow Values
Public Const F_RTS_CONTROL_ENABLE = &H1000
Public Const F_RTS_CONTROL_HANDSHAKE = &H2000
Public Const F_RTS_CONTROL_TOGGLE = &H3000

Public Const F_ABORT_ON_ERROR = &H4000

'---------Parity flags--------
Public Const EVENPARITY = 2
Public Const MARKPARITY = 3
Public Const NOPARITY = 0
Public Const ODDPARITY = 1
Public Const SPACEPARITY = 4

'---------StopBits-----------
Public Const ONESTOPBIT = 0
Public Const ONE5STOPBITS = 1
Public Const TWOSTOPBITS = 2

Public Type COMMTIMEOUTS
  ReadIntervalTimeout As Long
  ReadTotalTimeoutMultiplier As Long
  ReadTotalTimeoutConstant As Long
  WriteTotalTimeoutMultiplier As Long
  WriteTotalTimeoutConstant As Long
End Type

'Constants for the dwDesiredAccess parameter of the CreateFile() function
Public Const GENERIC_READ = &H80000000
Public Const GENERIC_WRITE = &H40000000

'Constants for the dwShareMode parameter of the CreateFile() function
Public Const FILE_SHARE_READ = &H1
Public Const FILE_SHARE_WRITE = &H2

'Constants for the dwCreationDisposition parameter of the CreateFile() function
Public Const CREATE_NEW = 1
Public Const CREATE_ALWAYS = 2
Public Const OPEN_EXISTING = 3

'Constants for the dwFlagsAndAttributes parameter of the CreateFile() function
Public Const FILE_ATTRIBUTE_NORMAL = &H80
Public Const FILE_FLAG_OVERLAPPED = &H40000000

'Error codes reported by the CreateFile().
Public Const ERROR_FILE_NOT_FOUND = 2
Public Const ERROR_ACCESS_DENIED = 5
Public Const ERROR_INVALID_HANDLE = 6


Public Declare PtrSafe Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, _
        ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, _
        ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, _
        ByVal hTemplateFile As Long) As Long

Public Declare PtrSafe Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare PtrSafe Function SetCommState Lib "kernel32" (ByVal hFile As Long, lpDCB As DCB) As Long
Public Declare PtrSafe Function GetCommState Lib "kernel32" (ByVal hFile As Long, lpDCB As DCB) As Long

Public Declare PtrSafe Function SetCommTimeouts Lib "kernel32" (ByVal hFile As Long, _
        lpCommTimeouts As COMMTIMEOUTS) As Long

Public Declare PtrSafe Function GetCommTimeouts Lib "kernel32" (ByVal hFile As Long, _
        lpCommTimeouts As COMMTIMEOUTS) As Long

Public Declare PtrSafe Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, _
         ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal lpOverlapped As Long) _
         As Long

Public Declare PtrSafe Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, _
         ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, _
         ByVal lpOverlapped As Long) As Long
	 

Public Sub Init_Com(bOpen As Boolean)
    Dim rcp As Long
    
    Dim p As Long
    p = CreateFile("\\.\COM3", GENERIC_READ Or GENERIC_WRITE, 0, 0, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0)
    'For serial port numbers higher than 9 see KB115831

    If p = -1 Then
        rcp = Err.LastDllError
        Select Case rcp 'Two typical error codes when trying to open a serial port:
         Case ERROR_ACCESS_DENIED  ' - The serial port opened by another application
           MsgBox "The serial port is used by another program"
         Case ERROR_FILE_NOT_FOUND ' - The serial port does not exist, check the port name specified in the CreateFile()
           MsgBox "The serial port does not exist"
         Case Else
           MsgBox "CreateFile failed, the error code is " & Str(rcp)
        End Select
        Exit Sub
    End If

    Dim d As DCB 'The DCB structure and the SetCommState() function allow to set the baud rate and the byte size of the serial port.
    rcp = GetCommState(p, d)
    d.ByteSize = 8
    d.BaudRate = 9600
    d.fBitFields = F_BINARY 'Windows does not support non-binary data transfers so the flag must always be set in the DCB structure.
    
    'Another example how to set some flags in the DCB.
    'd.fBitFields = F_BINARY Or F_PARITY Or F_RTS_CONTROL_ENABLE
    
    d.StopBits = ONESTOPBIT
    d.Parity = NOPARITY
    rcp = SetCommState(p, d)
    If rcp = 0 Then
      rcp = Err.LastDllError
      MsgBox "SetCommState failed, the error code is " & Str(rcp)
    End If
    
    
    Dim timeouts As COMMTIMEOUTS 'prevents VBA code from hanging
    rc = GetCommTimeouts(p, timeouts)  'specify the maximum time Windows will wait to receive data
    timeouts.ReadIntervalTimeout = 3  'The max time in ms between receiving two bytes of data
    timeouts.ReadTotalTimeoutConstant = 20 'The max wait time for data.
    timeouts.ReadTotalTimeoutMultiplier = 0
    rcp = SetCommTimeouts(p, timeouts)
    If rcp = 0 Then
      rcp = Err.LastDllError
      MsgBox "SetCommTimeouts failed, the error code is " & Str(rcp)
      GoTo close_and_exit
    End If

    Dim bytOpen(1 To 4) As Byte
    bytOpen(1) = &HA0
    bytOpen(2) = &H1
    bytOpen(3) = &H1
    bytOpen(4) = &HA2
    
    Dim bClose(1 To 4) As Byte
    bClose(1) = &HA0
    bClose(2) = &H1
    bClose(3) = &H0
    bClose(4) = &HA1
    
    If bOpen Then
        Dim wr As Long
        rcp = WriteFile(p, bytOpen(1), 4, wr, 0) 'wr indicates how many bytes went to the port.
        If rcp = 0 Then
          rcp = Err.LastDllError
          MsgBox "WriteFile failed, the error code is " & Str(rcp)
          GoTo close_and_exit
        End If
    Else
        rcp = WriteFile(p, bClose(1), 4, wr, 0) 'wr indicates how many bytes went to the port.
        If rcp = 0 Then
          rcp = Err.LastDllError
          MsgBox "WriteFile failed, the error code is " & Str(rcp)
          GoTo close_and_exit
        End If
    End If
        
close_and_exit:
    rcp = CloseHandle(p) 
    'In VBA, always execute this call. Or you will receive the ERROR_ACCESS_DENIED next time when opening the port
    'and you will need to reload Word/Excel/Access to free the port.
End Sub

Private Function checkCharge() As Boolean
    Dim con As New ADODB.Connection
    Dim rs As New ADODB.Recordset
    Dim iValue As Integer
    Dim bCharge As Boolean
        
    On Error GoTo err_handler
    
    With con
        .CursorLocation = adUseClient
        .Mode = adModeRead
        .ConnectionString = "Driver={ODBC Driver 18 for SQL Server};Server=tcp:[Your Azure Server].database.windows.net,1433;Database=[Your Database Name];Uid=[Your User name];Pwd=[Your password];Encrypt=yes;TrustServerCertificate=no;Connection Timeout=30;"
        .Open
    End With
    
    rs.Open "SELECT TOP 1 [Percent], Charge FROM WORK_BATTERY", con, adOpenStatic, adLockReadOnly
    
    iValue = rs.Fields(0)
    bCharge = rs.Fields(1)
    Debug.Print (iValue)
        
    If iValue <= 15 And Not bCharge Then
        con.Execute "UPDATE WORK_BATTERY SET Charge=1"
        Init_Com True
    ElseIf iValue = 100 And bCharge Then
        con.Execute "UPDATE WORK_BATTERY SET Charge=0"
        Init_Com False
    End If
    
    rs.Close
    Set rs = Nothing
    con.Close
    Set con = Nothing
    Exit Function

err_handler:
    MsgBox Err.Description
End Function

```

And that's it really. You just need to put the checkCharge() sub in a timer so you could activate it from a button using the example below:
```
Sub StartPowerMonitor_Click()
    Dim interval As Variant
    interval = Now + TimeValue("00:00:10")
    Application.OnTime interval, "checkCharge"
End Sub

```
You can adjust this code according to your circumstances. So you may be able to omit a massive chunck of the code if your laptop allows you to install drivers. You could just put the battery monitor on a timer and them call the Init_Com sub when the criteria is met and remove the whole need of a database or a web page etc.<br>

Sometimes the relay sticks. It's only happened once for me. In this case just give it a flick or a tap and it should respond.
