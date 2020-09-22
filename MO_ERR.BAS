Attribute VB_Name = "MO_Error"
Option Explicit


' ERG2: Modulo que contiene la mayoría de los
' Errores atrapables de Visual Basic


Sub ME_Muestra_Error()
' Función que muestra los errores que genera Visual Basic
  
  MG_Mensaje Error(Err) + "=> : Codigo :" + Str(Err)
  
End Sub


Function ME_WsockError(LP_NUM%) As String

' Función que se encarga de mostrar todos los errores que se
' puedan presentar utilizando el API de Winsock

Dim L_Mensaje$
    
 
    L_Mensaje = "ERROR " & LP_NUM & ": "
    Select Case LP_NUM
        Case WSASYSNOTREADY
            L_Mensaje = L_Mensaje & "WSASYSNOTREADY" & Chr(13)
            L_Mensaje = L_Mensaje & _
            "The underlying network subsystem is not " + _
            " ready for network communication."
            
        Case WSAVERNOTSUPPORTED
            L_Mensaje = L_Mensaje & "WSAVERNOTSUPPORTED" & Chr(13)
            L_Mensaje = L_Mensaje & _
            "The version of Windows Sockets API support " + _
            "requested is not provided by this particular " + _
            " Windows Sockets implementation."
            
        Case WSAEINVAL
            L_Mensaje = L_Mensaje & "WSAEINVAL" & Chr(13)
            L_Mensaje = L_Mensaje & _
            "The Windows Sockets version specified by the " + _
            "application is not supported by this DLL."
            
        Case WSANOTINITIALISED
            L_Mensaje = L_Mensaje & "WSANOTINITIALISED" & Chr(13)
            L_Mensaje = L_Mensaje & _
            "A successful WSAStartup() must occur before using this API."
            
        Case WSAENETDOWN
            L_Mensaje = L_Mensaje & "WSAENETDOWN" & Chr(13)
            L_Mensaje = L_Mensaje & _
            "The Windows Sockets implementation has detected " + _
            "that the network subsystem has failed."
            
        Case WSAEINPROGRESS
            L_Mensaje = L_Mensaje & "WSAEINPROGRESS" & Chr(13)
            L_Mensaje = L_Mensaje & _
            "A blocking Windows Sockets operation is in progress."
            
        Case WSAEFAULT
            L_Mensaje = L_Mensaje & "WSAEFAULT" & Chr(13)
            L_Mensaje = L_Mensaje & _
            "The namelen parameter is too small"
            
        Case WSAEADDRINUSE
            L_Mensaje = L_Mensaje & "WSAEADDRINUSE" & Chr(13)
            L_Mensaje = L_Mensaje & _
            "The specified address is already in use."
            
        Case WSAEINTR
            L_Mensaje = L_Mensaje & "WSAEINTR" & Chr(13)
            L_Mensaje = L_Mensaje & _
            "The (blocking) call was canceled via " + _
            " WSACancelBlockingCall()."
            
        Case WSAEADDRNOTAVAIL
            L_Mensaje = L_Mensaje & "WSAEADDRNOTAVAIL" & Chr(13)
            L_Mensaje = L_Mensaje & _
            "The specified address is not available from " + _
            "the local machine."
            
        Case WSAEAFNOSUPPORT
            L_Mensaje = L_Mensaje & "WSAEAFNOSUPPORT" & Chr(13)
            L_Mensaje = L_Mensaje & _
            "Addresses in the specified family cannot be " + _
            "used with this socket."
            
        Case WSAECONNREFUSED
            L_Mensaje = L_Mensaje & "WSAECONNREFUSED" & Chr(13)
            L_Mensaje = L_Mensaje & _
            "The attempt to connect was forcefully rejected."
        
        Case WSAEFAULT
            L_Mensaje = L_Mensaje & "WSAEFAULT" & Chr(13)
            L_Mensaje = L_Mensaje & _
            "The namelen argument is incorrect."
        
        Case WSAEINVAL
            L_Mensaje = L_Mensaje & "WSAEINVAL" & Chr(13)
            L_Mensaje = L_Mensaje & _
            "The socket is not already bound to an address."
            
        Case WSAEISCONN
            L_Mensaje = L_Mensaje & "WSAEISCONN" & Chr(13)
            L_Mensaje = L_Mensaje & _
            "The socket is already connected."
            
        Case WSAEMFILE
            L_Mensaje = L_Mensaje & "WSAEMFILE" & Chr(13)
            L_Mensaje = L_Mensaje & _
            "No more file descriptors are available."
            
        Case WSAENETUNREACH
            L_Mensaje = L_Mensaje & "WSAENETUNREACH" & Chr(13)
            L_Mensaje = L_Mensaje & _
            "The network can't be reached from this host at" + _
            "this time."
            
        Case WSAENOBUFS
            L_Mensaje = L_Mensaje & "WSAENOBUFS" & Chr(13)
            L_Mensaje = L_Mensaje & _
            "No buffer space is available.  The socket " + _
            "cannot be connected."
            
        Case WSAENOTSOCK
            L_Mensaje = L_Mensaje & "WSAENOTSOCK" & Chr(13)
            L_Mensaje = L_Mensaje & _
            "The descriptor is not a socket."
            
        Case WSAETIMEDOUT
            L_Mensaje = L_Mensaje & "WSAETIMEDOUT" & Chr(13)
            L_Mensaje = L_Mensaje & _
            "Attempt to connect timed out without " + _
            "establishing a connection"
            
        Case WSAEWOULDBLOCK
            L_Mensaje = L_Mensaje & "WSAEWOULDBLOCK" & Chr(13)
            L_Mensaje = L_Mensaje & _
            "The socket is marked as non-blocking and the " + _
            "connection cannot be completed immediately.  " + _
            " It is possible to select() the socket while " + _
            "it is connecting by select()ing it for writing."
            
        Case WSAEOPNOTSUPP
            L_Mensaje = L_Mensaje & "WSAEOPNOTSUPP" & Chr(13)
            L_Mensaje = L_Mensaje & _
            "The referenced socket is not of a type that " + _
            "supports the listen() operation."
            
        Case Else
            L_Mensaje = L_Mensaje & "Unknown error."
    End Select
    
    ME_WsockError = L_Mensaje
    

End Function


