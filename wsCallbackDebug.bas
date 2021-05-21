Attribute VB_Name = "wsCallbackDebug"
Option Explicit

' debugging version - slower, increases the odds of crash/hang!!

Public Sub WebSocketCallbackDebug(ByVal hInternet As LongPtr, ByVal dwContext As LongPtr, ByVal dwInternetStatus As Long, _
                                 ByVal lpvStatusInformation As LongPtr, ByVal dwStatusInformationLength As Long)
Dim Info(0 To 1) As Long ' used to get data from status structure
Dim Context As Long ' used to get data from context structure
Dim strHandle As String

Select Case hInternet
  Case wsSessionHandle
    strHandle = " - Session Handle"
  Case wsConnectionHandle
    strHandle = " - Connection Handle"
  Case wsRequestHandle
    strHandle = " - Request Handle"
  Case wsWebSocketHandle
    strHandle = " - WebSocket Handle"
  Case Else
    strHandle = " - Unknown Handle (" & hInternet & ")"
End Select

Call RtlMoveMemory(VarPtr(Context), dwContext, 4)
If Context <> wsContext Then
  strHandle = strHandle & " Unknown Context: " & Context
Else ' this call belongs to us, handle it
  
  Select Case dwInternetStatus
  '     the following flag was found on internet, does not seem to be valid
  '  Case WINHTTP_CALLBACK_STATUS_HANDLE_CLOSED ' indicate that the handle is completely closed.
  '    Debug.Print "WINHTTP_CALLBACK_STATUS_HANDLE_CLOSED" & strHandle
  '    TerminationComplete = True
    
    Case WINHTTP_CALLBACK_STATUS_HANDLE_CLOSING
      If debugPrint Then Debug.Print "WINHTTP_CALLBACK_STATUS_HANDLE_CLOSING" & strHandle
      If hInternet = wsWebSocketHandle Then wsLastCallback = True  ' also means websocket is closed?
      
    Case WINHTTP_CALLBACK_STATUS_CLOSE_COMPLETE ' successfully closed by call to WinHttpWebSocketClose
      wsCloseComplete = True
      wsState = 1
      If debugPrint Then Debug.Print "WINHTTP_CALLBACK_STATUS_CLOSE_COMPLETE" & strHandle
    
    Case WINHTTP_CALLBACK_STATUS_SHUTDOWN_COMPLETE ' successfully closed by call to WinHttpWebSocketShutdown
      wsCloseComplete = True
      wsState = 1
      If debugPrint Then Debug.Print "WINHTTP_CALLBACK_STATUS_SHUTDOWN_COMPLETE" & strHandle
    
    Case WINHTTP_CALLBACK_STATUS_REQUEST_ERROR
      If debugPrint Then Debug.Print "WINHTTP_CALLBACK_STATUS_REQUEST_ERROR" & strHandle
      
    Case WINHTTP_CALLBACK_STATUS_WRITE_COMPLETE
      If debugPrint Then Debug.Print "WINHTTP_CALLBACK_STATUS_WRITE_COMPLETE" & strHandle
      wsWriteComplete = True
      
    Case WINHTTP_CALLBACK_STATUS_READ_COMPLETE
      Call RtlMoveMemory(VarPtr(Info(0)), lpvStatusInformation, 8)
      If debugPrint Then Debug.Print "WINHTTP_CALLBACK_STATUS_READ_COMPLETE" & strHandle
      
      ' Info(1) is buffer type: 0 is Binary, 1 is Binary fragment,
      '    2 is UTF8, 3 is UTF8 fragment, 4 is close
      If debugPrint Then Debug.Print "Bytes: " & Info(0) & " Buffer type: " & Info(1)
      
      ' future - need to make sure we don't overrun buffer !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
      
      Select Case Info(1)
        Case WINHTTP_WEB_SOCKET_UTF8_FRAGMENT_BUFFER_TYPE ' 3: UTF8 fragment
          wsBufferIndex = wsBufferIndex + Info(0)
          wsBufferLength = wsBufferLength - Info(0)
          wsReadComplete = True
          If debugPrint Then Debug.Print "Buffer Index: " & wsBufferIndex & " Remaining Length: " & wsBufferLength
        
        Case WINHTTP_WEB_SOCKET_UTF8_MESSAGE_BUFFER_TYPE ' 2: end of UTF8 message
          wsBufferIndex = wsBufferIndex + Info(0)
          wsBufferLength = wsBufferLength - Info(0)
          wsReadComplete = True
          wsReceiveComplete = True
          If debugPrint Then Debug.Print "Buffer Index: " & wsBufferIndex & " Remaining Length: " & wsBufferLength
  
        Case WINHTTP_WEB_SOCKET_BINARY_FRAGMENT_BUFFER_TYPE  ' 1: binary fragment
          wsBufferIndex = wsBufferIndex + Info(0)
          wsBufferLength = wsBufferLength - Info(0)
          wsReadComplete = True
          If debugPrint Then Debug.Print "Buffer Index: " & wsBufferIndex & " Remaining Length: " & wsBufferLength
      
        Case WINHTTP_WEB_SOCKET_BINARY_MESSAGE_BUFFER_TYPE ' 0: end of binary message
          wsBufferIndex = wsBufferIndex + Info(0)
          wsBufferLength = wsBufferLength - Info(0)
          wsReadComplete = True
          wsReceiveComplete = True
          If debugPrint Then Debug.Print "Buffer Index: " & wsBufferIndex & " Remaining Length: " & wsBufferLength
      
        Case WINHTTP_WEB_SOCKET_CLOSE_BUFFER_TYPE ' 4: server wants to close connection
          wsReadComplete = True
          wsReceiveComplete = True
          wsServerDisconnect = True
          If debugPrint Then Debug.Print "Callback buffer=WINHTTP_WEB_SOCKET_CLOSE_BUFFER_TYPE"
          Call Disconnect
          
      End Select
      
    Case WINHTTP_CALLBACK_STATUS_DATA_AVAILABLE  ' http request data, not websocket data
      If debugPrint Then Debug.Print "Unexpected Callback: WINHTTP_CALLBACK_STATUS_DATA_AVAILABLE" & strHandle
      
    Case WINHTTP_CALLBACK_STATUS_SENDREQUEST_COMPLETE
      If debugPrint Then Debug.Print "Unexpected Callback: WINHTTP_CALLBACK_STATUS_SENDREQUEST_COMPLETE" & strHandle
      
    Case WINHTTP_CALLBACK_STATUS_CONNECTING_TO_SERVER
      If debugPrint Then Debug.Print "Unexpected Callback: WINHTTP_CALLBACK_STATUS_CONNECTING_TO_SERVER" & strHandle
      
    Case WINHTTP_CALLBACK_STATUS_CLOSING_CONNECTION
      If debugPrint Then Debug.Print "Unexpected Callback: WINHTTP_CALLBACK_STATUS_CLOSING_CONNECTION" & strHandle
      
    Case WINHTTP_CALLBACK_STATUS_CONNECTION_CLOSED
      If debugPrint Then Debug.Print "Unexpected Callback: WINHTTP_CALLBACK_STATUS_CONNECTION_CLOSED" & strHandle
      
    Case WINHTTP_CALLBACK_STATUS_HEADERS_AVAILABLE
      If debugPrint Then Debug.Print "Unexpected Callback: WINHTTP_CALLBACK_STATUS_HEADERS_AVAILABLE" & strHandle
      
    Case WINHTTP_CALLBACK_STATUS_REQUEST_ERROR
      If debugPrint Then Debug.Print "Unexpected Callback: WINHTTP_CALLBACK_STATUS_REQUEST_ERROR" & strHandle
      
    Case WINHTTP_CALLBACK_STATUS_HEADERS_AVAILABLE
      If debugPrint Then Debug.Print "Unexpected Callback: WINHTTP_CALLBACK_STATUS_HEADERS_AVAILABLE" & strHandle
      
    Case WINHTTP_CALLBACK_STATUS_SECURE_FAILURE
      If debugPrint Then Debug.Print "Unexpected Callback: WINHTTP_CALLBACK_STATUS_SECURE_FAILURE" & strHandle
      
    Case Else
      If debugPrint Then Debug.Print "Unexpected Callback: " & dwContext & ":" & Hex(dwInternetStatus) & ":" & Hex(lpvStatusInformation) & ":" & dwStatusInformationLength & strHandle
  End Select
End If
End Sub


