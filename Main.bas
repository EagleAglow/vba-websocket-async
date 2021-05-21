Attribute VB_Name = "Main"
Option Explicit

' ===================================================
' =  Echo server thanks to: https://websocket.org/  =
' ===================================================


Sub Main()
Dim LoopStart As Double
Dim msgBinary(1 To 18) As Byte, n As Long, strBinary As String
For n = 1 To 9
  msgBinary(n) = n
  msgBinary(n + 9) = (10 - n)
Next

Call Initialize

wsServer = "echo.websocket.org"

Call Connect
If wsState = 3 Then  ' connected, so continue

  OutBoxBinary = msgBinary()  ' put message in outbox
  Call SendBinary ' transmit message, set wsWriteComplete = False
  ' wait until wsWriteComplete = True
  LoopStart = wsTimer()
  Do Until ((wsServerDisconnect) Or (wsWriteComplete) Or (wsTimer() > (LoopStart + EmergencyStop)))
    DoEvents
  Loop
  If wsServerDisconnect Then ' server wants to quit
    Call Disconnect
  End If
  strBinary = ""
  For n = LBound(OutBoxBinary) To UBound(OutBoxBinary)
    strBinary = strBinary & OutBoxBinary(n)
  Next
  Debug.Print "Sent: " & strBinary
    
    
  Call ReceiveBinary   ' ask for message, set wsReadComplete = False
  ' wait until wsReadComplete
  LoopStart = wsTimer()
  Do Until ((wsServerDisconnect) Or (wsReadComplete) Or (wsTimer() > (LoopStart + EmergencyStop)))
    DoEvents
  Loop
  If wsServerDisconnect Then ' server wants to quit
    Call Disconnect
  End If
  strBinary = ""
  For n = LBound(InBoxBinary) To UBound(InBoxBinary)
    strBinary = strBinary & InBoxBinary(n)
  Next
  Debug.Print "Reply: " & strBinary
  

  OutBoxUTF8 = "Hello World!"  ' put message in outbox
  Call SendUTF8 ' transmit message, set wsWriteComplete = False
  ' wait until wsWriteComplete = True
  LoopStart = wsTimer()
  Do Until ((wsServerDisconnect) Or (wsWriteComplete) Or (wsTimer() > (LoopStart + EmergencyStop)))
    DoEvents
  Loop
  If wsServerDisconnect Then ' server wants to quit
    Call Disconnect
  End If
  Debug.Print "Sent: " & OutBoxUTF8
    
    
  Call ReceiveUTF8   ' ask for message, set wsReadComplete = False
  ' wait until wsReadComplete
  LoopStart = wsTimer()
  Do Until ((wsServerDisconnect) Or (wsReadComplete) Or (wsTimer() > (LoopStart + EmergencyStop)))
    DoEvents
  Loop
  If wsServerDisconnect Then ' server wants to quit
    Call Disconnect
  End If
  Debug.Print "Reply: " & InBoxUTF8
  

End If
Call Disconnect

Debug.Print "Program ends here..."
End Sub

