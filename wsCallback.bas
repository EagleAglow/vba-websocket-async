Attribute VB_Name = "wsCallback"
Option Explicit  ' shorter/faster version for normal use

Public Sub WebSocketCallback(ByVal hInternet As LongPtr, ByVal dwContext As LongPtr, ByVal dwInternetStatus As Long, _
                                 ByVal lpvStatusInformation As LongPtr, ByVal dwStatusInformationLength As Long)
Dim Info(0 To 1) As Long ' used to get data from status structure
Dim Context As Long ' used to get data from context structure

Call RtlMoveMemory(VarPtr(Context), dwContext, 4)
If Context = wsContext Then ' this is our call - handle it

  Select Case dwInternetStatus
    
    Case WINHTTP_CALLBACK_STATUS_HANDLE_CLOSING
      If hInternet = wsWebSocketHandle Then wsLastCallback = True  ' also means websocket is closed?
      
    Case WINHTTP_CALLBACK_STATUS_CLOSE_COMPLETE ' successfully closed by call to WinHttpWebSocketClose
      wsCloseComplete = True
      wsState = 1
    
    Case WINHTTP_CALLBACK_STATUS_SHUTDOWN_COMPLETE ' successfully closed by call to WinHttpWebSocketShutdown
      wsCloseComplete = True
      wsState = 1
    
  '  Case WINHTTP_CALLBACK_STATUS_REQUEST_ERROR
      
    Case WINHTTP_CALLBACK_STATUS_WRITE_COMPLETE
      wsWriteComplete = True
      
    Case WINHTTP_CALLBACK_STATUS_READ_COMPLETE
      Call RtlMoveMemory(VarPtr(Info(0)), lpvStatusInformation, 8)
      
      ' future - need to make sure we don't overrun buffer !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
      
      Select Case Info(1)
        Case WINHTTP_WEB_SOCKET_UTF8_FRAGMENT_BUFFER_TYPE ' 3: UTF8 fragment
          wsBufferIndex = wsBufferIndex + Info(0)
          wsBufferLength = wsBufferLength - Info(0)
          wsReadComplete = True
        
        Case WINHTTP_WEB_SOCKET_UTF8_MESSAGE_BUFFER_TYPE ' 2: end of UTF8 message
          wsBufferIndex = wsBufferIndex + Info(0)
          wsBufferLength = wsBufferLength - Info(0)
          wsReadComplete = True
          wsReceiveComplete = True
  
        Case WINHTTP_WEB_SOCKET_BINARY_FRAGMENT_BUFFER_TYPE  ' 1: binary fragment
          wsBufferIndex = wsBufferIndex + Info(0)
          wsBufferLength = wsBufferLength - Info(0)
          wsReadComplete = True
      
        Case WINHTTP_WEB_SOCKET_BINARY_MESSAGE_BUFFER_TYPE ' 0: end of binary message
          wsBufferIndex = wsBufferIndex + Info(0)
          wsBufferLength = wsBufferLength - Info(0)
          wsReadComplete = True
          wsReceiveComplete = True
      
        Case WINHTTP_WEB_SOCKET_CLOSE_BUFFER_TYPE ' 4: server wants to close connection
          wsReadComplete = True
          wsReceiveComplete = True
          wsServerDisconnect = True
          Call Disconnect
          
      End Select
  End Select
End If
End Sub


