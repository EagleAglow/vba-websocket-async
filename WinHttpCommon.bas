Attribute VB_Name = "WinHttpCommon"
Option Explicit
' note constant names from C header file had underscore character
' these have been replaced with lowercase "_" to avoid possibel conflicts with object scripting

' flags for WinHttpOpen():
Public Const WINHTTP_FLAG_SYNC = &H0          ' session is synchronous - use this
Public Const WINHTTP_FLAG_ASYNC = &H10000000  ' session is asynchronous - not implemented

' codes from GetLastError
' see: https://docs.microsoft.com/en-us/windows/win32/debug/system-error-codes--0-499-
Public Const ERROR_SUCCESS = 0
Public Const ERROR_INVALID_FUNCTION = 1
Public Const ERROR_INVALID_HANDLE = 6
Public Const ERROR_NOT_ENOUGH_MEMORY = 8

' flags for status check
Public Const WINHTTP_QUERY_STATUS_CODE = 19  ' special: part of status line
Public Const WINHTTP_QUERY_FLAG_NUMBER = &H20000000   ' bit flag to get result as number

' WinHttpQueryHeaders constants for code readability
Public Const WINHTTP_HEADER_NAME_BY_INDE_ = 0
Public Const WINHTTP_NO_OUTPUT_BUFFER = 0
Public Const WINHTTP_NO_HEADER_INDE_ = 0
Public Const WINHTTP_NO_REQUEST_DATA = 0

' a few HTTP Response Status Codes
Public Const HTTP_STATUS_CONTINUE = 100           ' OK to continue with request
Public Const HTTP_STATUS_SWITCH_PROTOCOLS = 101   ' server has switched protocols in upgrade header
Public Const HTTP_STATUS_OK = 200                 ' request completed

Public Const WINHTTP_WEB_SOCKET_BUFFER_TYPE = 0  ' Text, not specified in original sample code

Public Const INTERNET_DEFAULT_PORT = 0           ' use the protocol-specific default
Public Const INTERNET_DEFAULT_HTTP_PORT = 80     ' use the HTTP default
Public Const INTERNET_DEFAULT_HTTPS_PORT = 443   ' use the HTTPS default

Public Const WINHTTP_ACCESS_TYPE_DEFAULT_PRO_Y = 0
Public Const WINHTTP_ACCESS_TYPE_NO_PRO_Y = 1
Public Const WINHTTP_ACCESS_TYPE_NAMED_PRO_Y = 3
Public Const WINHTTP_ACCESS_TYPE_AUTOMATIC_PRO_Y = 4

Public Const WINHTTP_OPTION_UPGRADE_TO_WEB_SOCKET = 114
Public Const WINHTTP_OPTION_CONTEXT_VALUE = 45
Public Const WINHTTP_NO_ADDITIONAL_HEADERS = 0

' buffer types - old?
Public Const WINHTTP_WEB_SOCKET_BINARY_MESSAGE_BUFFER_TYPE = 0
Public Const WINHTTP_WEB_SOCKET_BINARY_FRAGMENT_BUFFER_TYPE = 1
Public Const WINHTTP_WEB_SOCKET_UTF8_MESSAGE_BUFFER_TYPE = 2
Public Const WINHTTP_WEB_SOCKET_UTF8_FRAGMENT_BUFFER_TYPE = 3
Public Const WINHTTP_WEB_SOCKET_CLOSE_BUFFER_TYPE = 4
' new? where did I see these?
'Public Const WEB_SOCKET_UTF8_MESSAGE_BUFFER_TYPE = &H80000000
'Public Const WEB_SOCKET_UTF8_FRAGMENT_BUFFER_TYPE = &H80000001
'Public Const WEB_SOCKET_BINARY_MESSAGE_BUFFER_TYPE = &H80000002
'Public Const WEB_SOCKET_BINARY_FRAGMENT_BUFFER_TYPE = &H80000003
'Public Const WEB_SOCKET_CLOSE_BUFFER_TYPE = &H80000004
'Public Const WEB_SOCKET_PING_PONG_BUFFER_TYPE = &H80000005
'Public Const WEB_SOCKET_UNSOLICITED_PONG_BUFFER_TYPE = &H80000006


' status codes
Public Const WINHTTP_WEB_SOCKET_SUCCESS_CLOSE_STATUS = 1000
Public Const WINHTTP_WEB_SOCKET_ENDPOINT_TERMINATED_CLOSE_STATUS = 1001
Public Const WINHTTP_WEB_SOCKET_PROTOCOL_ERROR_CLOSE_STATUS = 1002
Public Const WINHTTP_WEB_SOCKET_INVALID_DATA_TYPE_CLOSE_STATUS = 1003
Public Const WINHTTP_WEB_SOCKET_EMPTY_CLOSE_STATUS = 1005
Public Const WINHTTP_WEB_SOCKET_ABORTED_CLOSE_STATUS = 1006
Public Const WINHTTP_WEB_SOCKET_INVALID_PAYLOAD_CLOSE_STATUS = 1007
Public Const WINHTTP_WEB_SOCKET_POLICY_VIOLATION_CLOSE_STATUS = 1008
Public Const WINHTTP_WEB_SOCKET_MESSAGE_TOO_BIG_CLOSE_STATUS = 1009
Public Const WINHTTP_WEB_SOCKET_UNSUPPORTED_E_TENSIONS_CLOSE_STATUS = 1010
Public Const WINHTTP_WEB_SOCKET_SERVER_ERROR_CLOSE_STATUS = 1011
Public Const WINHTTP_WEB_SOCKET_SECURE_HANDSHAKE_ERROR_CLOSE_STATUS = 1015


'//
'// status manifests for WinHttp status callback
'//

Public Const WINHTTP_CALLBACK_STATUS_RESOLVING_NAME = &H1
Public Const WINHTTP_CALLBACK_STATUS_NAME_RESOLVED = &H2
Public Const WINHTTP_CALLBACK_STATUS_CONNECTING_TO_SERVER = &H4
Public Const WINHTTP_CALLBACK_STATUS_CONNECTED_TO_SERVER = &H8
Public Const WINHTTP_CALLBACK_STATUS_SENDING_REQUEST = &H10
Public Const WINHTTP_CALLBACK_STATUS_REQUEST_SENT = &H20
Public Const WINHTTP_CALLBACK_STATUS_RECEIVING_RESPONSE = &H40
Public Const WINHTTP_CALLBACK_STATUS_RESPONSE_RECEIVED = &H80
Public Const WINHTTP_CALLBACK_STATUS_CLOSING_CONNECTION = &H100
Public Const WINHTTP_CALLBACK_STATUS_CONNECTION_CLOSED = &H200
Public Const WINHTTP_CALLBACK_STATUS_HANDLE_CREATED = &H400
Public Const WINHTTP_CALLBACK_STATUS_HANDLE_CLOSING = &H800
Public Const WINHTTP_CALLBACK_STATUS_DETECTING_PROXY = &H1000

Public Const WINHTTP_CALLBACK_STATUS_REDIRECT = &H4000
Public Const WINHTTP_CALLBACK_STATUS_INTERMEDIATE_RESPONSE = &H8000
Public Const WINHTTP_CALLBACK_STATUS_SECURE_FAILURE = &H10000
Public Const WINHTTP_CALLBACK_STATUS_HEADERS_AVAILABLE = &H20000
Public Const WINHTTP_CALLBACK_STATUS_DATA_AVAILABLE = &H40000
Public Const WINHTTP_CALLBACK_STATUS_READ_COMPLETE = &H80000
Public Const WINHTTP_CALLBACK_STATUS_WRITE_COMPLETE = &H100000
Public Const WINHTTP_CALLBACK_STATUS_REQUEST_ERROR = &H200000
Public Const WINHTTP_CALLBACK_STATUS_SENDREQUEST_COMPLETE = &H400000


Public Const WINHTTP_CALLBACK_STATUS_GETPRO_YFORURL_COMPLETE = &H1000000
Public Const WINHTTP_CALLBACK_STATUS_CLOSE_COMPLETE = &H2000000
Public Const WINHTTP_CALLBACK_STATUS_SHUTDOWN_COMPLETE = &H4000000

Public Const WINHTTP_CALLBACK_STATUS_SETTINGS_WRITE_COMPLETE = &H10000000
Public Const WINHTTP_CALLBACK_STATUS_SETTINGS_READ_COMPLETE = &H20000000

'// API Enums for WINHTTP_CALLBACK_STATUS_REQUEST_ERROR:
Public Const API_RECEIVE_RESPONSE = 1
Public Const API_QUERY_DATA_AVAILABLE = 2
Public Const API_READ_DATA = 3
Public Const API_WRITE_DATA = 4
Public Const API_SEND_REQUEST = 5
Public Const API_GET_PRO_Y_FOR_URL = 6


Public Const WINHTTP_CALLBACK_FLAG_RESOLVE_NAME = (WINHTTP_CALLBACK_STATUS_RESOLVING_NAME Or WINHTTP_CALLBACK_STATUS_NAME_RESOLVED)
Public Const WINHTTP_CALLBACK_FLAG_CONNECT_TO_SERVER = (WINHTTP_CALLBACK_STATUS_CONNECTING_TO_SERVER Or WINHTTP_CALLBACK_STATUS_CONNECTED_TO_SERVER)
Public Const WINHTTP_CALLBACK_FLAG_SEND_REQUEST = (WINHTTP_CALLBACK_STATUS_SENDING_REQUEST Or WINHTTP_CALLBACK_STATUS_REQUEST_SENT)
Public Const WINHTTP_CALLBACK_FLAG_RECEIVE_RESPONSE = (WINHTTP_CALLBACK_STATUS_RECEIVING_RESPONSE Or WINHTTP_CALLBACK_STATUS_RESPONSE_RECEIVED)
Public Const WINHTTP_CALLBACK_FLAG_CLOSE_CONNECTION = (WINHTTP_CALLBACK_STATUS_CLOSING_CONNECTION Or WINHTTP_CALLBACK_STATUS_CONNECTION_CLOSED)
Public Const WINHTTP_CALLBACK_FLAG_HANDLES = (WINHTTP_CALLBACK_STATUS_HANDLE_CREATED Or WINHTTP_CALLBACK_STATUS_HANDLE_CLOSING)
Public Const WINHTTP_CALLBACK_FLAG_DETECTING_PRO_Y = WINHTTP_CALLBACK_STATUS_DETECTING_PROXY
Public Const WINHTTP_CALLBACK_FLAG_REDIRECT = WINHTTP_CALLBACK_STATUS_REDIRECT
Public Const WINHTTP_CALLBACK_FLAG_INTERMEDIATE_RESPONSE = WINHTTP_CALLBACK_STATUS_INTERMEDIATE_RESPONSE
Public Const WINHTTP_CALLBACK_FLAG_SECURE_FAILURE = WINHTTP_CALLBACK_STATUS_SECURE_FAILURE
Public Const WINHTTP_CALLBACK_FLAG_SENDREQUEST_COMPLETE = WINHTTP_CALLBACK_STATUS_SENDREQUEST_COMPLETE
Public Const WINHTTP_CALLBACK_FLAG_HEADERS_AVAILABLE = WINHTTP_CALLBACK_STATUS_HEADERS_AVAILABLE
Public Const WINHTTP_CALLBACK_FLAG_DATA_AVAILABLE = WINHTTP_CALLBACK_STATUS_DATA_AVAILABLE
Public Const WINHTTP_CALLBACK_FLAG_READ_COMPLETE = WINHTTP_CALLBACK_STATUS_READ_COMPLETE
Public Const WINHTTP_CALLBACK_FLAG_WRITE_COMPLETE = WINHTTP_CALLBACK_STATUS_WRITE_COMPLETE
Public Const WINHTTP_CALLBACK_FLAG_REQUEST_ERROR = WINHTTP_CALLBACK_STATUS_REQUEST_ERROR


Public Const WINHTTP_CALLBACK_FLAG_GETPRO_YFORURL_COMPLETE = WINHTTP_CALLBACK_STATUS_GETPRO_YFORURL_COMPLETE

Public Const WINHTTP_CALLBACK_FLAG_ALL_COMPLETIONS = (WINHTTP_CALLBACK_STATUS_SENDREQUEST_COMPLETE Or _
                 WINHTTP_CALLBACK_STATUS_HEADERS_AVAILABLE Or WINHTTP_CALLBACK_STATUS_DATA_AVAILABLE Or _
                 WINHTTP_CALLBACK_STATUS_READ_COMPLETE Or WINHTTP_CALLBACK_STATUS_WRITE_COMPLETE Or _
                 WINHTTP_CALLBACK_STATUS_REQUEST_ERROR Or WINHTTP_CALLBACK_STATUS_GETPRO_YFORURL_COMPLETE)
Public Const WINHTTP_CALLBACK_FLAG_ALL_NOTIFICATIONS = &HFFFFFFFF

'//
'// if the following value is returned by WinHttpSetStatusCallback, then
'// probably an invalid (non-code) address was supplied for the callback
'//
'Public Const WINHTTP_INVALID_STATUS_CALLBACK      =  ((WINHTTP_STATUS_CALLBACK)(-1L))
Public Const WINHTTP_INVALID_STATUS_CALLBACK As Long = -1

' ====================================================
' API functions
' ====================================================
' DO Use StrPtr to pass strings,
' DO NOT append Chr(0) to strings before passing them
' ====================================================

Declare PtrSafe Function GetLastError Lib "kernel32" () As Long

Public Declare PtrSafe Function WinHttpOpen Lib "winhttp" ( _
   ByVal pszAgentW As LongPtr, _
   ByVal dwAccessType As Long, _
   ByVal pszPro_yW As LongPtr, _
   ByVal pszPro_yBypassW As LongPtr, _
   ByVal dwFlags As Long _
   ) As LongPtr

Public Declare PtrSafe Function WinHttpConnect Lib "winhttp" ( _
   ByVal hSession As LongPtr, _
   ByVal pswzServerName As LongPtr, _
   ByVal nServerPort As Long, _
   ByVal dwReserved As Long _
   ) As LongPtr

Public Declare PtrSafe Function WinHttpOpenRequest Lib "winhttp" ( _
   ByVal hConnect As LongPtr, _
   ByVal pwszVerb As LongPtr, _
   ByVal pwszObjectName As LongPtr, _
   ByVal pwszVersion As LongPtr, _
   ByVal pwszReferrer As LongPtr, _
   ByVal ppwszAcceptTypes As LongPtr, _
   ByVal dwFlags As Long _
   ) As LongPtr

Public Declare PtrSafe Function WinHttpSetOption Lib "winhttp" ( _
   ByVal hInternet As LongPtr, _
   ByVal dwOption As Long, _
   ByVal lpBuffer As LongPtr, _
   ByVal dwBufferLength As Long _
   ) As Long

Public Declare PtrSafe Function WinHttpSendRequest Lib "winhttp" ( _
   ByVal hRequest As LongPtr, _
   ByVal lpszHeaders As LongPtr, _
   ByVal dwHeadersLength As Long, _
   ByVal lpOptional As LongPtr, _
   ByVal dwOptionalLength As Long, _
   ByVal dwTotalLength As Long, _
   ByVal dwContext As Long _
   ) As Long

Public Declare PtrSafe Function WinHttpReceiveResponse Lib "winhttp" ( _
   ByVal hRequest As LongPtr, _
   ByVal lpReserved As LongPtr _
   ) As Long

Public Declare PtrSafe Function WinHttpWebSocketCompleteUpgrade Lib "winhttp" ( _
   ByVal hRequest As LongPtr, _
   ByVal pConText As LongPtr _
   ) As LongPtr

Public Declare PtrSafe Function WinHttpCloseHandle Lib "winhttp" ( _
   ByVal hRequest As LongPtr _
   ) As Long

Public Declare PtrSafe Function WinHttpWebSocketSend Lib "winhttp" ( _
   ByVal hWebSocket As LongPtr, _
   ByVal eBufferType As Long, _
   ByVal pvBuffer As LongPtr, _
   ByVal dwBufferLength As Long _
   ) As Long

Public Declare PtrSafe Function WinHttpWebSocketReceive Lib "winhttp" ( _
   ByVal hWebSocket As LongPtr, _
   ByRef pvBuffer As Any, _
   ByVal dwBufferLength As Long, _
   ByRef pdwBytesRead As LongPtr, _
   ByRef peBufferType As LongPtr _
   ) As Long

Public Declare PtrSafe Function WinHttpWebSocketClose Lib "winhttp" ( _
   ByVal hWebSocket As LongPtr, _
   ByVal usStatus As Integer, _
   ByVal pvReason As LongPtr, _
   ByVal dwReasonLength As Long _
   ) As Long

Public Declare PtrSafe Function WinHttpWebSocketShutdown Lib "winhttp" ( _
   ByVal hWebSocket As LongPtr, _
   ByVal usStatus As Integer, _
   ByVal pvReason As LongPtr, _
   ByVal dwReasonLength As Long _
   ) As Long

Public Declare PtrSafe Function WinHttpWebSocketQueryCloseStatus Lib "winhttp" ( _
   ByVal hWebSocket As LongPtr, _
   ByRef usStatus As Integer, _
   ByRef pvReason As Any, _
   ByVal dwReasonLength As Long, _
   ByRef pdwReasonLengthConsumed As LongPtr _
   ) As Long

Public Declare PtrSafe Function WinHttpQueryHeaders Lib "winhttp" ( _
  ByVal hRequest As LongPtr, _
  ByVal dwInfoLevel As Long, _
  ByVal pwszName As LongPtr, _
  ByRef lpBuffer As Long, _
  ByRef lpdwBufferLength As Long, _
  ByRef lpdwInde_ As Long _
   ) As Long

' memcopy
Public Declare PtrSafe Sub RtlMoveMemory Lib "kernel32" ( _
                              ByVal dest As LongPtr, _
                              ByVal src As LongPtr, _
                              ByVal size As LongPtr)


' Pointer to an array of WEB_SOCKET_BUFFER structures that contain WebSocket buffer data.
' not using the following - delete later
Public Declare PtrSafe Function WebSocketGetAction Lib "websocket" ( _
  ByVal hWebSocket As LongPtr, _
  ByVal eActionQueue As Long, _
  ByVal pDataBuffers As LongPtr, _
  ByRef pulDataBufferCount As Long, _
  ByRef pAction As Long, _
  ByRef pBufferType As Long, _
  ByRef pvApplicationConText As LongPtr, _
  ByRef pvActionConText As LongPtr _
   ) As Long


' WinHttpSetStatusCallback function...

Public Declare PtrSafe Function WinHttpSetStatusCallback Lib "winhttp" ( _
  ByVal hWebSocket As LongPtr, _
  ByVal lpfnInternetCallback As LongPtr, _
  ByVal dwNotificationFlags As Long, _
  ByVal dwReserved As LongPtr _
   ) As Long


' ................................
' WinHttpStatusCallback function...

'SUB CALLBACK WINHTTP_STATUS_CALLBACK (
Public Sub WINHTTP_STATUS_CALLBACK( _
              ByVal hInternet As LongPtr, _
              ByVal dwContext As Long, _
              ByVal dwInternetStatus As Long, _
              ByVal lpvStatusInformation As LongPtr, _
              ByVal dwStatusInformationLength As Long _
              )
' do things with this information
End Sub
