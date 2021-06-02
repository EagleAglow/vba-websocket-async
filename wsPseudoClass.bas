Attribute VB_Name = "wsPseudoClass"
Option Explicit
Option Base 0
' ----------------------------------------------------------------------------------------
' An asynchronous websocket "class"
' This can't actually be a VBA class, as setting callbacks with "AddressOf" won't work
' ----------------------------------------------------------------------------------------
'
' Global variables, in lieu of class properties
Global OutBoxUTF8 As String            ' To send this, call "SendUTF8"
Global OutBoxBinary() As Byte          ' To send this, call "SendBinary"
Global InBoxUTF8 As String             ' To fill this, call "ReceiveUTF8" (which uses one or more calls to "ReadUTF8")
Global InBoxBinary() As Byte           ' To fill this, call "ReceiveBinary" (which uses one or more calls to "ReadBinary")

Global wsWriteComplete As Boolean      ' callback received: WINHTTP_CALLBACK_STATUS_WRITE_COMPLETE
Global wsReadComplete As Boolean       ' callback received: WINHTTP_CALLBACK_STATUS_READ_COMPLETE
Global wsReceiveComplete As Boolean    ' callback received: WINHTTP_CALLBACK_STATUS_READ_COMPLETE and no more fragments!
Global wsCloseComplete As Boolean      ' callback received: WINHTTP_CALLBACK_STATUS_CLOSE_COMPLETE
Global wsServerDisconnect As Boolean   ' callback received: WINHTTP_WEB_SOCKET_CLOSE_BUFFER_TYPE
Global TerminationComplete As Boolean  ' handle closed? maybe WINHTTP_CALLBACK_STATUS_SHUTDOWN_COMPLETE
Global WaitComplete As Boolean         ' result of loops to wait for events

Global wsSessionHandle As LongPtr      ' equal to zero, if not set
Global wsConnectionHandle As LongPtr   ' equal to zero, if not set
Global wsRequestHandle As LongPtr      ' equal to zero, if not set
Global wsWebSocketHandle As LongPtr    ' equal to zero, if not set
Global wsContext As Long               ' placeholder, to set a "context" for callbacks
Global wsContextPointer As LongPtr
Global wsLastCallback As Boolean       ' result of closing websocket handle

Global wsBuffer(0 To 1023) As Byte     ' websocket receive buffer, content copied to InBox
Global wsBufferLength As LongPtr       ' websocket receive buffer length, set by call to "Initialize"
Global wsBufferIndex As Long           ' index into websocket receive buffer, adjusted for each API call to receive

Global wsServer As String          ' host name
Global wsPort As Long              ' internet port, usually 80 or 443
Global wsPath As String            ' at least "/"
Global wsProtocol As String        ' actually "sub-protocol" per RFC, e.g. WASM or echo-protocol
Global wsAgentHeader As String     ' CWebSocket

Global wsState As Long          ' For websocket - 1:Not Connected, 2:Connecting, 3:Connected, 4:Disconnecting, Else: Error/Unknown
Global httpState As Long        ' For http transport - 1:Not Connected, 2:Connecting, 3:Connected, 4:Disconnecting, Else: Error/Unknown
Global wsErrorText As String    ' set to Text e_plaining last error
Global wsReadError As Boolean   ' probably error 4317 in Sub ReadUTF8

Global NormalStop  As Double    ' how long to wait for async function (read/write loop timeout) should be a bit longer than worst ping to server
Global EmergencyStop As Double  ' ultimate loop time limit, seconds

Global debugPrint As Boolean    ' set this to True for errors/messages to Debug.Print


Sub Initialize()
debugPrint = True
wsSessionHandle = 0
wsConnectionHandle = 0
wsRequestHandle = 0
wsWebSocketHandle = 0
wsContext = 123 ' picked a random number
wsContextPointer = VarPtr(wsContext)
wsLastCallback = False
wsReadComplete = False
wsReceiveComplete = False
wsServerDisconnect = False
WaitComplete = False
NormalStop = 0.5 / 3600 / 24  ' normal read/write loop timeout - should be a bit longer than worst ping to server
EmergencyStop = 5 / 3600 / 24  ' ultimate loop time limit (5 seconds converted to days)

wsBufferLength = UBound(wsBuffer) + 1
wsBufferIndex = 0

wsServer = ""           ' default, produces error if not set before connection
wsPort = 80             ' default
wsPath = "/"            ' default
wsProtocol = ""
wsAgentHeader = "WebSocketVBA"  ' just to serve as an identifier

wsState = 1
httpState = 1
wsErrorText = "None"
End Sub

Sub Connect() ' setup http and convert to websocket
Dim result As Long ' return value from API calls
If (wsState <> 1) Or (httpState <> 1) Then ' only run this while not connected
  wsErrorText = "Must be disconnected before attempting to connect"
  If debugPrint Then Debug.Print wsErrorText
  Exit Sub
End If

httpState = 2 ' connecting
' check to see if server is set
If Len(wsServer) = 0 Then
  wsErrorText = "Missing Server Name"
  httpState = 1
  If debugPrint Then Debug.Print wsErrorText
  Exit Sub
End If

' check to see if path begins with "/"
If Left(wsPath, 1) <> "/" Then
  wsErrorText = "Invalid Path: " & wsPath
  httpState = 1
  If debugPrint Then Debug.Print wsErrorText
  Exit Sub
End If

' Create session handle
wsSessionHandle = WinHttpOpen(StrPtr(wsAgentHeader), _
      WINHTTP_ACCESS_TYPE_DEFAULT_PRO_Y, 0, 0, WINHTTP_FLAG_ASYNC)
If wsSessionHandle = 0 Then
  wsErrorText = "Could not create WinHttp session handle"
  httpState = 1
  If debugPrint Then Debug.Print wsErrorText
  Exit Sub
End If

' Create connection handle
wsConnectionHandle = WinHttpConnect(wsSessionHandle, StrPtr(wsServer), wsPort, 0)
If wsConnectionHandle = 0 Then
  wsErrorText = "Failed to reach server:port at: " & wsServer & ":" & wsPort
  httpState = 1
  If debugPrint Then Debug.Print wsErrorText
  Exit Sub
End If

' Create request handle - use 0 for null pointers to empty strings: Version, Referrer, AcceptTypes
Dim method As String
method = "GET" ' always
wsRequestHandle = WinHttpOpenRequest(wsConnectionHandle, StrPtr(method), StrPtr(wsPath), 0, 0, 0, 0)
If wsRequestHandle = 0 Then
  wsErrorText = "Connection request failed for path: " & wsPath
  httpState = 1
  If debugPrint Then Debug.Print wsErrorText
  Exit Sub
Else
  httpState = 3 ' http connected
  If debugPrint Then Debug.Print "httpConnected"
End If

' Prepare to request client protocol upgrade from http to websocket, returns true if success
wsState = 2
result = WinHttpSetOption(wsRequestHandle, WINHTTP_OPTION_UPGRADE_TO_WEB_SOCKET, 0, 0)
If result = 0 Then ' failed
  wsErrorText = "Upgrade from http to websocket failed (Step 1/6)"
  wsState = 1  ' note: httpState is still 3 (connected)
  If debugPrint Then Debug.Print wsErrorText
  Exit Sub
End If

' Perform websocket handshake by sending the upgrade request to server
' --------------------------------------------------------------------
' Application may specify additional headers if needed.
' --------------------------------------------------------------------
' Each header except the last must be terminated by a carriage return/line feed (vbCrLf).
' Uses an odd API feature: passes string length as -1, API figures out length
' Note: This is where websocket (internal, RFC "subprotocol") protocol is set
' --------------------------------------------------------------------
Dim HeaderText As String
Dim HeaderTextLength As Long
HeaderText = ""
HeaderText = HeaderText & "Host: " & wsServer & vbCrLf   ' may be redundant or unnecessary
HeaderText = HeaderText & "Sec-WebSocket-Version: 13" & vbCrLf  ' 8 or 13, may be redundant or unnecessary
HeaderText = HeaderText & "Sec-Websocket-Protocol: echo-protocol" & vbCrLf  ' subprotocol
' setup for API call, trim any trailing vbCrLf
If (Right(HeaderText, 2) = vbCrLf) Then
  HeaderText = Left(HeaderText, Len(HeaderText) - 2)
End If

If Len(HeaderText) > 0 Then ' let the API figure it out
  HeaderTextLength = -1
  result = WinHttpSendRequest(wsRequestHandle, StrPtr(HeaderText), _
               HeaderTextLength, WINHTTP_NO_REQUEST_DATA, 0, 0, 0)
Else  ' call without adding headers
  result = WinHttpSendRequest(wsRequestHandle, WINHTTP_NO_ADDITIONAL_HEADERS, _
               0, WINHTTP_NO_REQUEST_DATA, 0, 0, 0)
End If
If (result = 0) Then ' failed
  wsErrorText = "Upgrade from http to websocket failed (Step 2/6)"
  wsState = 1  ' note: httpState is still 3 (connected)
  If debugPrint Then Debug.Print wsErrorText
  Exit Sub
End If

' Receive server reply
result = WinHttpReceiveResponse(wsRequestHandle, 0)
If result = 0 Then ' failed
  wsErrorText = "Upgrade from http to websocket failed (Step 3/6)"
  wsState = 1  ' note: httpState is still 3 (connected)
  If debugPrint Then Debug.Print wsErrorText
  Exit Sub
End If

' See if the HTTP Response confirms the upgrade, with HTTP status code 101.
Dim dwStatusCode As Long
Dim sizeStatusCode As Long  ' for HTTP result request
sizeStatusCode = 4 ' four bytes for long HTTP result request
result = WinHttpQueryHeaders(wsRequestHandle, _
    (WINHTTP_QUERY_STATUS_CODE Or WINHTTP_QUERY_FLAG_NUMBER), _
    WINHTTP_HEADER_NAME_BY_INDE_, _
    dwStatusCode, sizeStatusCode, WINHTTP_NO_HEADER_INDE_)
If (result = 0) Then ' failed
  wsErrorText = "Upgrade from http to websocket failed (Step 4/6)"
  wsState = 1  ' note: httpState is still 3 (connected)
  If debugPrint Then Debug.Print wsErrorText
  Exit Sub
End If

If dwStatusCode <> 101 Then
  Debug.Print "Code needs to be 101, ending..."
  wsErrorText = "Upgrade from http to websocket failed (Step 5/6)"
  wsState = 1  ' note: httpState is still 3 (connected)
  If debugPrint Then Debug.Print wsErrorText
  Exit Sub
End If

' finally, get handle to websocket
wsWebSocketHandle = WinHttpWebSocketCompleteUpgrade(wsRequestHandle, 0)
If wsWebSocketHandle = 0 Then
  wsErrorText = "Upgrade from http to websocket failed (Step 6/6)"
  wsState = 1  ' note: httpState is still 3 (connected)
  If debugPrint Then Debug.Print wsErrorText
  Exit Sub
End If

' The request handle is not needed anymore.
WinHttpCloseHandle (wsRequestHandle)
wsRequestHandle = 0

' A callback routine is required for asynchronous send/receive. Also, we want to receive
' WINHTTP_CALLBACK_STATUS_HANDLE_CLOSING to confirm that the websocket is closing. And,
' to generate that particular callback, the websocket handle must have a non-Null context
' value. So, set wsWebSocketHandle "context", with a pointer to a pointer.
result = WinHttpSetOption(wsWebSocketHandle, WINHTTP_OPTION_CONTEXT_VALUE, VarPtr(wsContextPointer), 4) '4 bytes for pointer
If result = 0 Then ' failed
  wsState = 1  ' note: httpState is still 3 (connected)
  wsErrorText = "Setting websocket context failed. Error:" & dwError & ":" & GetLastError
  If debugPrint Then Debug.Print wsErrorText
  Exit Sub
End If

' the callback seems to get all handle events, even if not flagged on setup ???
' set callback function, might as well flag all events
' "WebSocketCallback" for normal use, "WebSocketCallbackDebug" for trouble shooting
Dim dwError As Long
dwError = WinHttpSetStatusCallback(wsWebSocketHandle, AddressOf WebSocketCallback, _
            WINHTTP_CALLBACK_FLAG_ALL_NOTIFICATIONS, 0)
If dwError < 0 Then
  wsState = 1  ' note: httpState is still 3 (connected)
  wsErrorText = "Could not set callback function. Error:" & dwError & ":" & GetLastError
  If debugPrint Then Debug.Print wsErrorText
  Exit Sub
End If

If debugPrint Then Debug.Print "Succesfully upgraded to websocket protocol at: " & wsServer & ":" & wsPort & wsPath
wsState = 3  ' connected!
End Sub

Sub SendBinary()  ' transmit the array in OutBoxBinary
Dim result As Long
Dim BinaryMessage() As Byte
Dim BinaryMessageLength As Long
Dim i As Long
wsErrorText = "None"
wsWriteComplete = False  ' until reset by callback
If wsState <> 3 Then
  wsErrorText = "Must be connected to send"
  If debugPrint Then Debug.Print wsErrorText
  Exit Sub
End If
BinaryMessageLength = (UBound(OutBoxBinary) - LBound(OutBoxBinary) + 1)
ReDim BinaryMessage(0 To (BinaryMessageLength - 1))
For i = 0 To (BinaryMessageLength - 1)
  BinaryMessage(i) = OutBoxBinary(LBound(OutBoxBinary) + i)
Next

result = WinHttpWebSocketSend(wsWebSocketHandle, _
             WINHTTP_WEB_SOCKET_BINARY_MESSAGE_BUFFER_TYPE, _
             VarPtr(BinaryMessage(0)), BinaryMessageLength)
If (result <> ERROR_SUCCESS) Then
  wsErrorText = "SendUTF8 Error: " & result & "/" & GetLastError
  If debugPrint Then Debug.Print wsErrorText
End If
End Sub

Sub ReceiveBinary()  ' ask for a message
Dim i As Long
Dim LoopStart As Double
' reset buffer pointer and available length
wsBufferLength = UBound(wsBuffer) + 1
wsBufferIndex = 0

wsReceiveComplete = False     ' until reset by callback for unfragmented frame (or error) with WINHTTP_CALLBACK_STATUS_READ_COMPLETE
wsReadError = False
'loop until wsReceiveComplete
LoopStart = wsTimer()
Do Until ((wsReadError) Or (wsReceiveComplete) Or (wsTimer() > (LoopStart + EmergencyStop)))
  DoEvents       ' let any outstanding callback resolve, maybe?
  Call ReadBinary  ' semi-synchronous, returns after callback from WinHttpWebSocketReceive
Loop

If wsReceiveComplete Then ' put result into InBox...
  If wsBufferIndex > 0 Then
    ReDim InBoxBinary(0 To (wsBufferIndex - 1))
    For i = 0 To wsBufferIndex - 1
      InBoxBinary(i) = wsBuffer(i)
    Next
  End If
  ' reset buffer pointer and available length
  wsBufferLength = UBound(wsBuffer) + 1
  wsBufferIndex = 0
End If

If wsReadError Then ' reset buffer
  ' reset buffer pointer and available length
  wsBufferLength = UBound(wsBuffer) + 1
  wsBufferIndex = 0
End If
End Sub

Sub ReadBinary()  ' call WinHttpWebSocketReceive, wait for it to complete
' ----------------------------------------------------------
' do not call this directly, use ReceiveBinary to get here!!!!
' ----------------------------------------------------------
Dim LoopStart As Double
Dim result As Long
Dim eBufferType As Long ' defined, but not usefull when asynchronous, returns zero
Dim eBufferBytesTransferred As Long ' likwise
wsErrorText = "None"
eBufferBytesTransferred = 0
wsReadComplete = False       ' until reset by callback for unfragmented frame (or error) with WINHTTP_CALLBACK_STATUS_READ_COMPLETE

result = WinHttpWebSocketReceive(wsWebSocketHandle, wsBuffer(wsBufferIndex), CLng(wsBufferLength), eBufferBytesTransferred, eBufferType)
If (result = 4317) Then  ' ERROR_INVALID_OPERATION
  wsReadError = True
  ' we're not getting anything, set wsReadComplete and ReceiveComplete and return
  ' this may terminate the connection, though...  future investigation...
  wsErrorText = "ReadBinary Error: INVALID OPERATION" & "/" & GetLastError
  If debugPrint Then Debug.Print wsErrorText
  wsReceiveComplete = True
  wsReadComplete = True
  Exit Sub
Else
  If (result <> ERROR_SUCCESS) Then
    wsErrorText = "ReadBinary Error: " & result & "/" & GetLastError
    If debugPrint Then Debug.Print wsErrorText
    wsReceiveComplete = True
    wsReadComplete = True
    Exit Sub
  End If
  ' wait until wsReadComplete
  LoopStart = wsTimer()
  Do Until ((wsReadComplete) Or (wsTimer() > (LoopStart + EmergencyStop)))
    DoEvents
  Loop
End If
End Sub

Sub SendUTF8()  ' transmit the string in OutBoxUTF8
Dim result As Long
Dim UTF8Message() As Byte
Dim UTF8MessageLength As Long
wsErrorText = "None"
wsWriteComplete = False  ' until reset by callback
If wsState <> 3 Then
  wsErrorText = "Must be connected to send"
  If debugPrint Then Debug.Print wsErrorText
  Exit Sub
End If

UTF8Message = Utf8BytesFromString(OutBoxUTF8)
UTF8MessageLength = BytesLength(UTF8Message)

result = WinHttpWebSocketSend(wsWebSocketHandle, _
             WINHTTP_WEB_SOCKET_UTF8_MESSAGE_BUFFER_TYPE, _
             VarPtr(UTF8Message(0)), UTF8MessageLength)
If (result <> ERROR_SUCCESS) Then
  wsErrorText = "SendUTF8 Error: " & result & "/" & GetLastError
  If debugPrint Then Debug.Print wsErrorText
End If
End Sub

Sub ReceiveUTF8()  ' ask for a message
Dim ByteBuffer() As Byte  ' so we can redim later, to size of actual message
Dim i As Long
Dim LoopStart As Double
' reset buffer pointer and available length
wsBufferLength = UBound(wsBuffer) + 1
wsBufferIndex = 0

InBoxUTF8 = ""  ' empty InBox
wsReceiveComplete = False     ' until reset by callback for unfragmented frame (or error) with WINHTTP_CALLBACK_STATUS_READ_COMPLETE
wsReadError = False
'loop until wsReceiveComplete
LoopStart = wsTimer()
Do Until ((wsReadError) Or (wsReceiveComplete) Or (wsTimer() > (LoopStart + EmergencyStop)))
  DoEvents       ' let any outstanding callback resolve, maybe?
  Call ReadUTF8  ' semi-synchronous, returns after callback from WinHttpWebSocketReceive
Loop

If wsReceiveComplete Then ' put result into InBox...
  If wsBufferIndex > 0 Then
    ReDim ByteBuffer(0 To (wsBufferIndex - 1))
    For i = 0 To wsBufferIndex - 1
      ByteBuffer(i) = wsBuffer(i)
    Next
    InBoxUTF8 = Utf8BytesToString(ByteBuffer())
  End If
  ' reset buffer pointer and available length
  wsBufferLength = UBound(wsBuffer) + 1
  wsBufferIndex = 0
End If

If wsReadError Then ' reset buffer and InBox
  ' reset buffer pointer and available length
  wsBufferLength = UBound(wsBuffer) + 1
  wsBufferIndex = 0
  InBoxUTF8 = ""
End If
End Sub

Sub ReadUTF8()  ' call WinHttpWebSocketReceive, wait for it to complete
' ----------------------------------------------------------
' do not call this directly, use ReceiveUTF8 to get here!!!!
' ----------------------------------------------------------
Dim LoopStart As Double
Dim result As Long
Dim eBufferType As Long ' defined, but not usefull when asynchronous, returns zero
Dim eBufferBytesTransferred As Long ' likwise
wsErrorText = "None"
eBufferBytesTransferred = 0
wsReadComplete = False       ' until reset by callback for unfragmented frame (or error) with WINHTTP_CALLBACK_STATUS_READ_COMPLETE

result = WinHttpWebSocketReceive(wsWebSocketHandle, wsBuffer(wsBufferIndex), CLng(wsBufferLength), eBufferBytesTransferred, eBufferType)
If (result = 4317) Then  ' ERROR_INVALID_OPERATION
  wsReadError = True
  ' we're not getting anything, set wsReadComplete and ReceiveComplete and return
  ' this may terminate the connection, though...  future investigation...
  wsErrorText = "ReadUTF8 Error: INVALID OPERATION" & "/" & GetLastError
  If debugPrint Then Debug.Print wsErrorText
  wsReceiveComplete = True
  wsReadComplete = True
  Exit Sub
Else
  If (result <> ERROR_SUCCESS) Then
    wsErrorText = "ReadUTF8 Error: " & result & "/" & GetLastError
    If debugPrint Then Debug.Print wsErrorText
    wsReceiveComplete = True
    wsReadComplete = True
    Exit Sub
  End If
  ' wait until wsReadComplete
  LoopStart = wsTimer()
  Do Until ((wsReadComplete) Or (wsTimer() > (LoopStart + EmergencyStop)))
    DoEvents
  Loop
End If
End Sub

Public Function wsTimer() As Double  ' returns a "mid-night-safe" value from Timer function
' https://stackoverflow.com/questions/27304842/how-to-programm-a-timer-on-vba-that-runs-over-midnight
Dim StartDate As Long
Dim StartTime As Double
Dim CurrentDate As Long
Dim CurrentTime As Double
Dim dblStartDateTime   As Double
Dim dblCurrentDateTime As Double
Dim dblElapsedDateTime As Double
StartDate = Date
StartTime = VBA.Timer
wsTimer = StartDate + (StartTime / 24 / 3600)  ' note: units are days
End Function

Sub Disconnect()
Dim result As Long ' return value from API calls
Dim uStatus As Integer
Dim LoopStart As Double
httpState = 4 ' disconnecting

' try to gracefully close websocket - tell the server goodbye... (leaves receive channel open)
' if the connection was successfully shut down via a call to WinHttpWebSocketShutdown,
' the function sends a close frame to WebSocket server, and the callback function eventually
' gets WINHTTP_CALLBACK_STATUS_SHUTDOWN_COMPLETE

wsCloseComplete = False
uStatus = WINHTTP_WEB_SOCKET_SUCCESS_CLOSE_STATUS   ' tell the server this is OK?
result = WinHttpWebSocketShutdown(wsWebSocketHandle, uStatus, 0, 0)
' any error other than ERROR_IO_PENDING (possibly &H000003E5) means underlying TCP connection has been aborted.
If (result <> ERROR_SUCCESS) Then
  wsErrorText = "Websocket shutdown failed with error: " & result
  If debugPrint Then Debug.Print wsErrorText
  wsState = 5
End If
'loop until wsCloseComplete
LoopStart = wsTimer()
Do Until ((wsCloseComplete) Or (wsTimer() > (LoopStart + EmergencyStop)))
  DoEvents
Loop
If debugPrint Then
  If (Not wsCloseComplete) Then Debug.Print "Websocket Shutdown Timed Out"
End If

' close the websocket - this does not close the handle
' the callback function eventually gets WINHTTP_CALLBACK_STATUS_CLOSE_COMPLETE
wsCloseComplete = False
result = WinHttpWebSocketClose(wsWebSocketHandle, uStatus, 0, 0)
If (result <> ERROR_SUCCESS) Then
  wsErrorText = "Websocket close failed with error: " & result
  If debugPrint Then Debug.Print wsErrorText
  wsState = 5
End If

'loop until wsCloseComplete
LoopStart = wsTimer()
Do Until ((wsCloseComplete) Or (wsTimer() > (LoopStart + EmergencyStop)))
  DoEvents
Loop
If debugPrint Then
  If (Not wsCloseComplete) Then Debug.Print "Websocket Close Timed Out"
End If

' finally, close the websocket handle with WinHttpCloseHandle...
' (that call is supposedly synchronous, but the results may not be... )
' callback function either gets WINHTTP_CALLBACK_STATUS_REQUEST_ERROR
' or WINHTTP_CALLBACK_STATUS_HANDLE_CLOSING (which will be the last call for that handle)
wsLastCallback = False
result = WinHttpCloseHandle(wsWebSocketHandle)

If result = 0 Then 'error
  wsErrorText = "Error Closing WebSocketHandle: " & result
  If debugPrint Then Debug.Print wsErrorText
End If
' wait for callback to get WINHTTP_CALLBACK_STATUS_HANDLE_CLOSING
LoopStart = wsTimer()
Do Until ((wsLastCallback) Or (wsTimer() > (LoopStart + EmergencyStop)))
  DoEvents
Loop
If debugPrint Then
  If (Not wsLastCallback) Then Debug.Print "Websocket Handle Close Timed Out"
End If

' disconnected!
If debugPrint Then Debug.Print "Succesfully disconnected"

End Sub



