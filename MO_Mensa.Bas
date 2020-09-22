Attribute VB_Name = "MO_Mensajes"
Option Explicit
' Máximo número de Sockets soportados por nuestro sistema
    
    Global Const GC_MAXSOCKETS = 5
    

' Declaración de la estructura de Sockets

Type ES_SOCKET    'Estructura de Sockets
' Dirección del servidor a conectarse con...
' ver(MO_Winsock=>sockaddr_in )
Addr As sockaddr_in

' Socket creado para realizar la conexión
socket As Long

' Nick con el cual se conoce el usuario en esta conexión
Nick As String

' Nombre real del Usuario
Usuario As String

' Nombre del Servidor al cual se conectará con este socket
Nombre As String

' Dirección del Servidor formato Internet i.e: unitec.hn
' al cual se conectará con este socket
Direcc As String

' Puerto utilizado  para la conexion
Puerto As Integer

' Tipo de Servidor al que representa
Tipo As Long

' Nos indica si este item del arreglo esta libre para ser
' utilizado
Libre As Boolean

' Indice en el arreglo de Ventanas, la cual representa
' la ventana de estatus de este servidor
Ventana As Integer

' Bandera que nos indica si existe un join pendiente
' en el servidor que representa el socket
JOIN_PENDIENTE As Boolean

End Type

' Arreglo global de Sockets , solo 5 en este caso de
' la estructura ante mencionada
Global GV_Sockets(GC_MAXSOCKETS) As ES_SOCKET

' Declaración de la variable global que contiene
' información de el API de WINSOCK que se esta utilizando
Global GV_WSAdata As WSAData

' Declaración de la variable global nos indica si WINSOCK
' ya ha sido levantado...

Global GV_WINSOCK_ARRIBA As Boolean

Sub MM_Cargar_Usuarios_A_Canal(LP_Lista_Parametros() As String, LP_Forma_Canal As OF_Hablar_Canal)
' /************************************************************/
' Procedimiento que carga la lista de usuarios que está en un
' canal. La lista de usuarios de un canal viene en el parámetro
' LP_Lista_Parametros(). LP_Forma_Canal es la ventana a la cual
' se cargará la lista de usuarios.  Esta ventana es una instancia
' de la forma OF_Hablar_Canal.
' /************************************************************/

On Error GoTo Etiqueta_Error:
Dim L_i As Long
Dim L_Largo As Integer
L_Largo = UBound(LP_Lista_Parametros) - 1
For L_i = 1 To UBound(LP_Lista_Parametros) - 1
  LP_Forma_Canal.OL_Usuarios.AddItem LP_Lista_Parametros(L_i)
  
Next L_i

Exit Sub
Etiqueta_Error:
ME_Muestra_Error
End Sub

Sub MM_Recibir_Archivo(Lp_Minick$, Lp_Emisor$, Lp_info$)
' /************************************************************/
' Procedimiento que se encarga de parsear la información de un
' nuevo archivo que se nos esta enviando y es colocada en una
' nueva ventana de donde se recibirá el archivo.
' LP_Lp_Minick es el alias que representa al usuario en un
' servidor.
' Lp_Emisor es el alias de la persona que envía el archivo.
' LP_info contiene toda la información que se necesita para
' recibir el archivo
' (nombre, dirección IP, puerto, tamaño).
' /************************************************************/
On Error GoTo Etiqueta_Error:

'Nombre del archivo que recibiremos
Dim L_Archivo$

'Dirección IP de la cual recibiremos el archivo
Dim L_IP$

'Puerto en la dirección IP de la cual recibiremos el archivo
Dim L_Puerto%

'Tamaño del archivo que se recibirá
Dim L_Tamano&

'Comando que especifica que opción del protocolo se utilizará.
'Ej: send y chat.
Dim L_Comando$
                
Dim L_Protocol$ 'DCC
Dim L_pos%, L_pos2%

L_pos = InStr(1, Lp_info, " ")
L_Protocol = Trim(Mid(Lp_info, 2, L_pos - 1))
If UCase(L_Protocol) = "DCC" Then
    L_pos2 = InStr(L_pos + 1, Lp_info, " ")
    L_Comando = Trim(Mid(Lp_info, L_pos, L_pos2 - L_pos))
    L_pos = InStr(L_pos2 + 1, Lp_info, " ")
    L_Archivo = Trim(Mid(Lp_info, L_pos2, L_pos - L_pos2))
    L_pos2 = InStr(L_pos + 1, Lp_info, " ")
    L_IP = (Trim(Mid(Lp_info, L_pos, L_pos2 - L_pos)))
    
End If
Select Case UCase(L_Protocol)
 Case "DCC"
   Select Case UCase(L_Comando)
    Case "SEND"
     L_pos = InStr(L_pos2 + 1, Lp_info, " ")
     L_Puerto = CInt(Trim(Mid(Lp_info, L_pos2, L_pos - L_pos2)))
     L_pos2 = InStr(L_pos + 1, Lp_info, Chr(1))
     L_Tamano = CLng(Trim(Mid(Lp_info, L_pos, L_pos2 - L_pos)))
     Dim L_Ventana As New OF_Archivo
     Load L_Ventana
     L_Ventana.OT_Nombre = L_Archivo
     L_Ventana.OT_SALVAR = L_Archivo
     L_Ventana.OT_IP = L_IP
     L_Ventana.OT_Puerto = L_Puerto
     L_Ventana.OT_Nick = Lp_Emisor
     L_Ventana.OT_Size = L_Tamano
     L_Ventana.Show
         
   Case "CHAT"
     L_pos = InStr(L_pos2 + 1, Lp_info, Chr(1))
     L_Puerto = CInt(Trim(Mid(Lp_info, L_pos2, L_pos - L_pos2)))
     MG_Mensaje "Prot:" + L_Protocol + " Comando: " + L_Comando + _
    " Archivo: " + L_Archivo + " IP: " + CStr(L_IP) + " Puerto: " + _
     CStr(L_Puerto)
       
   Case Else
     ' No soportado
  End Select
 
 Case Else ' Del L_Protocol
     ' No soportado

End Select

Exit Sub
Etiqueta_Error:
ME_Muestra_Error

End Sub

Function MM_Verificar_Comando(ByVal L_Mensaje$, L_Nick$) As String
' /**********************************************************/
' Diferencia si un mensaje a enviar es un comando o un simple
' mensaje texto a un canal o usuario.
' Todos los comando deben ser antecedidos por una pleca (/),
' a excepción de las ventanas de estátus en las cuales no
' es permitido.
' L_Mensaje contiene el mensaje o comando a enviar.
' L_Nick contiene el alias del usuario o el nombre del canal
' al que se envía el
' mensaje en caso de que este no sea un comando.
' /************************************************************/

On Error GoTo Etiqueta_Error:

Dim L_Texto$
Dim L_res%
If Mid(L_Mensaje, 1, 1) <> "/" Then 'Si es mensaje texto
  MM_Verificar_Comando = "PRIVMSG " + L_Nick + " : " + L_Mensaje
Else 'Si es comando
  L_res = InStr(1, L_Mensaje, " ")
  If L_res = 0 Then
   L_Texto = Trim(L_Mensaje)
  Else
   L_Texto = MM_Obtener_Mensaje_Parametros(1, L_Mensaje, " ")
  End If
    L_Texto = MM_Obtener_Mensaje_Parametros(2, L_Texto, "/")
    If L_res <> 0 Then
        ' En caso de ser un comando en español, se obtiene el
        ' comando equivalente válido para el servidor.
        MM_Verificar_Comando = _
        MD_Obtener_Comando(L_Texto) + " " + _
        MM_Obtener_Mensaje_Parametros(2, L_Mensaje, " ")
    Else
        MM_Verificar_Comando = Trim(MD_Obtener_Comando(L_Texto))
    End If
End If

Exit Function
Etiqueta_Error:
ME_Muestra_Error
End Function

Sub MM_Cerrar_Conexion(LP_Cual%)
' /************************************************************/
' Procedimiento que se encarga de cerrar una conexión.  De
' parametro recibe el índice en el arreglo de sockets, la
' conexión se cerrara para el socket que se encuentre
' en ese indice en el arreglo.
' /************************************************************/

On Error GoTo Etiqueta_Error:

Dim L_Status%
 ' Revisar todas las ventanas que tengan asociado este socket
 ' y a cada una setearles el socket asociado como cero
 MV_Setear_Socket LP_Cual
 
 ' Revisar todas las entradas en la estructura de retransmisión
 ' con el socket asociado
 ' y cada una setearlas como borradas
 MM_Borrar_Retransmisiones LP_Cual
 
 ' Despues enviar al servidor el comando quit
 MV_Pone_Mensaje False, _
 GV_VENTANAS_Estatus(GV_Sockets(LP_Cual).Ventana), _
 "Enviando el comando Quit...." & GV_EOD, vbRed
 
 If send(GV_Sockets(LP_Cual).socket, "QUIT" + GV_EOD, 6, 0) _
            <> SOCKET_ERROR Then
   MV_Pone_Mensaje False, _
   GV_VENTANAS_Estatus(GV_Sockets(LP_Cual).Ventana), _
   "Quit enviado..." & GV_EOD, vbRed
 End If
 
 ' Cerrar el socket
 MV_Pone_Mensaje False, _
 GV_VENTANAS_Estatus(GV_Sockets(LP_Cual).Ventana), _
 "Cerrando el socket del servidor..." & GV_EOD, vbRed
 
 L_Status = closesocket(GV_Sockets(LP_Cual).socket)
 If L_Status <> SOCKET_ERROR Then
     MV_Pone_Mensaje False, _
     GV_VENTANAS_Estatus(GV_Sockets(LP_Cual).Ventana), _
     "Socket Cerrado" & GV_EOD, vbRed
     GV_Sockets(LP_Cual).socket = INVALID_SOCKET
 End If
 ' Setear como libre el socket
  GV_Sockets(LP_Cual).Libre = True
 
 ' Ya no tiene socket asociado la ventana
  GV_VENTANAS_Estatus( _
  GV_Sockets(LP_Cual).Ventana).OT_SocketAsociado = 0
  
' Cambiar el Icono al servidor en el explorador
If GV_Explorador Then
    OF_Explorador.OA_Explorador.Nodes.Item( _
    "S" + CStr(GV_Sockets(LP_Cual).Ventana)).Image = 5
    
    OF_Explorador.OA_Explorador.Nodes.Item( _
    "S" + CStr(GV_Sockets(LP_Cual).Ventana)).Text = _
    GV_Sockets(LP_Cual).Direcc + " Estado [No Conectado]"
    
End If
 
 ' Actualizar el titulo de la ventana que representaba el socket
  GV_VENTANAS_Estatus(GV_Sockets(LP_Cual).Ventana).Caption = _
  GV_Sockets(LP_Cual).Direcc + " Estado [No Conectado]"
  
Exit Sub
Etiqueta_Error:
ME_Muestra_Error
End Sub

Function MM_Busca_Libre() As Integer
' /**********************************************************/
' Función que se encarga de buscar un item libre en el
' arreglo de sockets para ser utilizado para una nueva conexión.
' /**********************************************************/

On Error GoTo Etiqueta_Error:

Dim L_i%

' Buscar hasta el maximo # de Sockets, esta version=5
For L_i = 1 To GC_MAXSOCKETS
   ' Si esta marcado como libre
   If GV_Sockets(L_i).Libre = True Then
        MM_Busca_Libre = L_i   ' Retorne el indice libre
        Exit Function
   End If
Next L_i
MM_Busca_Libre = -1 ' No hay disponibles

Exit Function
Etiqueta_Error:
ME_Muestra_Error
End Function

Sub MM_About_WINSOCK(LP_WSA As WSAData)
' /**********************************************************/
' Procedimiento que se encarga de mostrar información del API
' de Winsock, Con el cual se está trabajando, esta
' información fue cargada cuando se inicializó la interfase
' de WINSOCK con el procedimiento MM_Inicializa_WINSOCK
' específicamente con la función WSAStartup()
' LP_WSA : Estructura utilizada por WINSOCK para registrar
' datos de él (Ver, MO_Winsock).
' /**********************************************************/

On Error GoTo Etiqueta_Error:

Dim L_Mensaje$
  With LP_WSA   ' Con el registro
   L_Mensaje = "Version: " & _
   Hex(.wVersion) & GV_EOD
   
   L_Mensaje = L_Mensaje & "High Version: " & _
   Hex(.wHighVersion) & GV_EOD
   
   L_Mensaje = L_Mensaje & "Description: " & _
   .szDescription & GV_EOD
   
   L_Mensaje = L_Mensaje & "System_Mensaje Status: " & _
   .szSystemStatus & GV_EOD
   
   L_Mensaje = L_Mensaje & "Max Sockets: " & _
   .iMaxSockets & GV_EOD
   
   L_Mensaje = L_Mensaje & "Max UdpDg: " & _
   .iMaxUdpDg & GV_EOD
   
   L_Mensaje = L_Mensaje & "*Vendor: " & .lpVendorInfo
   
  End With
  MG_Mensaje "Información de Windows Socket " + _
  GV_EOD + L_Mensaje ' Presenta la información

Exit Sub
Etiqueta_Error:
ME_Muestra_Error
End Sub

Function MM_Cierra_WINSOCK() As Boolean
' /**********************************************************/
' Función que se encarga de cerrar la interfase con el API
' de Winsock
' /**********************************************************/

On Error GoTo Etiqueta_Error:

Dim L_Status As Integer

L_Status = WSACleanUp() ' Cierra Winsock
If L_Status = SOCKET_ERROR Then 'Si existe error
    L_Status = WSAGetLastError()
    MG_Mensaje ME_WsockError(L_Status)
    MM_Cierra_WINSOCK = False
    Exit Function
End If
     
MM_Cierra_WINSOCK = True ' Winsock cerrado ...

Exit Function
Etiqueta_Error:
ME_Muestra_Error
End Function

Function MM_Obtener_Puerto_Libre() As Integer
' /**********************************************************/
' Función que obtiene el primer índice libre en el arreglo
' de puertos de la aplicación.  De no encontrar un puerto
' libre, esta función retorna 0.
' /**********************************************************/
 
 Dim L_i%
 MM_Obtener_Puerto_Libre = 0
 For L_i = 1 To 5
   If GV_Puertos(L_i).E_Libre = True Then
      GV_Puertos(L_i).E_Libre = False
      MM_Obtener_Puerto_Libre = L_i
      Exit Function
   End If
   
 Next L_i
  
 End Function
Function MM_Conectarse_Servidor(LP_Cual%, LP_Dirr$, LP_Puerto%, Lp_Tipo&) As Boolean
' /******************************************************/
' Función que se encarga de conectarse a un servidor
' específico.  Retorna True si se lleva a cabo la
' conexión, de lo contrario se retorna false.
' LP_Cual: índice en el arreglo de sockets que representa
' la conexión con el servidor.
' LP_Dirr: Dirección del Servidor al que se conectará
' la aplicación
' LP_Puerto: Puerto que se utilizará para la conexión.
' LP_Tipo: Tipo de servidor con el que se hará la conexión.
' Este es un código
' /******************************************************/

On Error GoTo Etiqueta_Error:

Dim L_Status As Integer
Dim L_direccion As Long

MM_Conectarse_Servidor = False

' Dirección de Internet
GV_Sockets(LP_Cual).Addr.sin_family = PF_INET

' Convertir el formato del puerto _
de host order a network order
GV_Sockets(LP_Cual).Addr.sin_port = htons(LP_Puerto)

' Formato IP i.e 130.1.1.1
If MM_Tipo_Direccion(LP_Dirr) = True Then
     ' Convertir la dirección a formato de red (Long)
    L_direccion = inet_addr(LP_Dirr)
Else  ' Formato toronto.undernet.org
  ' función del DLL que convierte la dirección
  L_direccion = DireccionRed(Trim(LP_Dirr))
  If L_direccion = 0 Then
    L_Status = closesocket(GV_Sockets(LP_Cual).socket)
    GV_Sockets(LP_Cual).socket = INVALID_SOCKET
    MV_Pone_Mensaje False, _
    GV_VENTANAS_Estatus(GV_Sockets(LP_Cual).Ventana), _
    "No se obtuvo la direccion IP del servidor.." & _
    GV_EOD, vbRed
    Exit Function
  End If
End If
 
' Asociar al socket la dirección a la cual se desea conectar
GV_Sockets(LP_Cual).Addr.sin_addr = L_direccion

' Realizar la conexión utilizando la función ConnectIn del
' APi de Winsock

L_Status = connectIn(GV_Sockets(LP_Cual).socket, _
            GV_Sockets(LP_Cual).Addr, 16)
If L_Status = SOCKET_ERROR Then ' Si ocurrió un error
  L_Status = WSAGetLastError()
  ' En caso de error se cierra el socket asociado en el
  ' arreglo de sockets
   ' Error , Time Out de Winsock, el servidor no responde
  If L_Status = 10060 Then
    L_Status = closesocket(GV_Sockets(LP_Cual).socket)
    GV_Sockets(LP_Cual).socket = INVALID_SOCKET
    Exit Function
  Else
    L_Status = closesocket(GV_Sockets(LP_Cual).socket)
    GV_Sockets(LP_Cual).socket = INVALID_SOCKET
    Exit Function
  End If

End If
' Si llega aqui es que logro conectarse, pero falta
' realizar la notificación de eventos
' Todavía no es del todo un exito ....
 
' Fijar el modo Asincrono para que seamos notificados de
' lecturas y de cierre de conexión
' Utilizando la función asíncrona del API de
' Winsock <WSAAsyncSelect>


L_Status = WSAAsyncSelect(GV_Sockets(LP_Cual).socket, _
    GV_VENTANAS_Estatus(GV_Sockets( _
    LP_Cual).Ventana).OT_Asynccontrol.hwnd, _
    WM_SIZE, FD_READ + FD_CLOSE)
          
If L_Status = SOCKET_ERROR Then ' Si hubo error ....
   L_Status = WSAGetLastError()
   MG_Mensaje ME_WsockError(L_Status)
   L_Status = closesocket(GV_Sockets(LP_Cual).socket)
   GV_Sockets(LP_Cual).socket = INVALID_SOCKET
   Exit Function

End If
    
' Si llega aqui entonces lo logro...

' Retorne True, ya que la conexión se realizó
MM_Conectarse_Servidor = True

Exit Function

Etiqueta_Error:

ME_Muestra_Error
End Function

Function MM_Connect(LP_Servidor As String, LP_Puerto%, Lp_Tipo&, LP_Ventana%) As Boolean
' /**********************************************************/
' Función que se encarga de:
' - Buscar una casilla libre en el arreglo se sockets
' - Crear el socket para la casilla libre
' - Crear la ventana de estatus que representará la conexión
' - Realizar la conexión con el servidor deseado. Parametros
'   LP_Servidor, Lp_Puerto
' - Registrar la conexión con la información del usuario
'
' /**********************************************************/
On Error GoTo Etiqueta_Error:

Dim L_libre As Integer
Dim L_Ventana As Integer
Dim L_direccion As Long
Dim L_Usuario As ES_USUARIO
Dim L_nodo As Node

MM_Connect = False

' Busca un item libre en el arreglo de sockets
L_libre = MM_Busca_Libre()

' Revisar si encontro una casilla libre
If L_libre <> -1 Then

  ' Asegurarse que esta ya no esta disponible
  GV_Sockets(L_libre).Libre = False
   
  'Busca un indice de ventana disponible
  If LP_Ventana = 0 Then
    L_Ventana = MV_Indice_Libre_Estatus()
  Else
    L_Ventana = LP_Ventana
  End If
  
  GV_VENTANAS_Estatus(L_Ventana).OT_Ventana_Estatus = L_Ventana
  
  'Asociar al Tag de la ventana el indice en el arreglo
  GV_VENTANAS_Estatus(L_Ventana).Tag = L_Ventana
  GV_VENTANAS_Estatus(L_Ventana).OT_Direccion = LP_Servidor
  GV_VENTANAS_Estatus(L_Ventana).OT_Puerto = LP_Puerto
  GV_VENTANAS_Estatus(L_Ventana).OT_Tipo = Lp_Tipo
  ' Asociar al arreglo de sockets en la casilla _
  disponible que ventana lo representara
  GV_Sockets(L_libre).Ventana = L_Ventana
  
  ' Asociar el tipo de servidor
  GV_Sockets(L_libre).Tipo = Lp_Tipo
  
  ' Asociar la dirección del servidor al socket que _
  lo representara
  GV_Sockets(L_libre).Direcc = LP_Servidor
  GV_Sockets(L_libre).Puerto = LP_Puerto
  GV_VENTANAS_Estatus(L_Ventana).Show
  GV_VENTANAS_Estatus(L_Ventana).Caption = _
  "Conectandose .....a " + LP_Servidor
  If GV_Explorador Then
     If LP_Ventana = 0 Then
        Set L_nodo = OF_Explorador.OA_Explorador _
        .Nodes.Add("EX", tvwChild, "S" + _
        CStr(L_Ventana), _
        GV_VENTANAS_Estatus(L_Ventana).Caption, 3)
        L_nodo.EnsureVisible
      Else
        OF_Explorador.OA_Explorador.Nodes.Item( _
        "S" + CStr(L_Ventana)).Image = 3
      
      End If
  End If
  
  MV_Pone_Mensaje False, _
  GV_VENTANAS_Estatus(L_Ventana), _
  "Ventana Creada para servidor:" + LP_Servidor & GV_EOD, _
  GV_Azul
    
' No asociar ningun socket al inicio
  GV_VENTANAS_Estatus(L_Ventana).OT_SocketAsociado = 0
  Screen.MousePointer = vbHourglass
  
  ' Trata de Obtener un handle para un Socket
  If MM_Obtener_Socket(L_libre) Then
     
     ' Convertir LP_Servidor a L_direccion a Formato de Internet
     ' Asociar el socket con la ventana, es decir el indice
     ' del arreglo de sockets
     
     GV_VENTANAS_Estatus(L_Ventana).OT_SocketAsociado = L_libre
     

     If MM_Conectarse_Servidor(L_libre, LP_Servidor, LP_Puerto, _
                               Lp_Tipo) Then
        ' Si llega a este punto entonces, registrarse como
        ' un usuario
         MV_Pone_Mensaje False, _
         GV_VENTANAS_Estatus(L_Ventana), _
         "Registrandose como usuario...." & GV_EOD, GV_Azul
         
         ' Recuperar la info del usuario
         If MD_Recupera_Infousuario(L_Usuario) = 0 Then

            MM_Registrar_Usuario L_Usuario.E_Alias, _
            L_Usuario.E_Nombre + "  " + L_Usuario.E_Host + _
            " " + L_Usuario.E_Host + " : " + L_Usuario.E_Nombre, _
            L_libre
            
            GV_Sockets(L_libre).Nick = Trim(L_Usuario.E_Alias)
            GV_Sockets(L_libre).Usuario = Trim(L_Usuario.E_Nombre)
            
         Else
            MV_Pone_Mensaje False, _
            GV_VENTANAS_Estatus(L_Ventana), _
            "No se pudo obtener la información del Usuario.." & _
            GV_EOD, vbBlue
            
            GV_VENTANAS_Estatus(L_Ventana).Caption = _
            LP_Servidor + "(" + CStr(LP_Puerto) + ") Estado : [No Conectado]"
            
            Screen.MousePointer = vbDefault
            
            Exit Function
         End If
     
         
     Else
        MV_Pone_Mensaje False, _
        GV_VENTANAS_Estatus(L_Ventana), _
        "No se pudo realizar la conexión con " _
        & " el servidor : " + LP_Servidor + GV_EOD, vbBlue
        
        'Desasociar el socket de la ventana
        GV_VENTANAS_Estatus(L_Ventana).OT_SocketAsociado = 0
        
        'Ahora el socket esta libre
        GV_Sockets(L_libre).Libre = True
        GV_VENTANAS_Estatus(L_Ventana).Caption = _
        LP_Servidor + "(" + CStr(LP_Puerto) + ") Estado : [No Conectado]"
        
        If GV_Explorador Then
           OF_Explorador.OA_Explorador.Nodes.Item( _
           "S" + CStr(L_Ventana)).Text = _
           GV_VENTANAS_Estatus(L_Ventana).Caption
           
           OF_Explorador.OA_Explorador.Nodes.Item( _
           "S" + CStr(L_Ventana)).Image = 5
    
        End If
        Screen.MousePointer = vbDefault
        Exit Function
     End If
  Else
        MV_Pone_Mensaje False, _
        GV_VENTANAS_Estatus(L_Ventana), _
        "No se pudo obtener el socket para la conexión" & _
        GV_EOD, vbBlue
        
        GV_VENTANAS_Estatus(L_Ventana).Caption = _
        LP_Servidor + "(" + CStr(LP_Puerto) + ") Estado : [No Conectado]"
        
        If GV_Explorador Then
           OF_Explorador.OA_Explorador.Nodes.Item( _
           "S" + CStr(L_Ventana)).Text = _
           GV_VENTANAS_Estatus(L_Ventana).Caption
    
        End If

        Screen.MousePointer = vbDefault
        Exit Function
  End If
Else
  MG_Mensaje "< No existen conexiones disponibles... >"
  Screen.MousePointer = vbDefault
  Exit Function
End If

' Si llega a este punto, entonces se pudo realizar una
' conexión ...
GV_VENTANAS_Estatus(L_Ventana).Caption = _
LP_Servidor + "(" + CStr(LP_Puerto) + ") Estado : [CONECTADO]"

GV_VENTANAS_Estatus(L_Ventana).OT_Comando.SetFocus
    
If GV_Explorador Then
    OF_Explorador.OA_Explorador.Nodes.Item( _
    "S" + CStr(L_Ventana)).Text = _
    GV_VENTANAS_Estatus(L_Ventana).Caption
    
End If
Screen.MousePointer = vbDefault
MM_Connect = True

Exit Function
Etiqueta_Error:
Screen.MousePointer = vbDefault
ME_Muestra_Error
End Function

Sub MM_Descargar_Usuarios_De_Canal(LP_Usuario As String, LP_Forma_Canal As OF_Hablar_Canal)
' /*********************************************************/
' Elimina un usuario de la lista de un canal considerando
' el caso especial de que este sea un operador(@).
' LP_Usuario: Usuario a eliminar.
' LP_Forma_Canal: Ventana del canal en la cual se eliminará
' al usuario.  Esta es una instancia de la forma
' OF_Hablar_Canal.
' /*********************************************************/

On Error GoTo Etiqueta_Error:

Dim L_i As Long
Dim L_Item As String

For L_i = 0 To (LP_Forma_Canal.OL_Usuarios.ListCount - 1)
    L_Item = Trim(LP_Forma_Canal.OL_Usuarios.List(L_i))
    If Left(L_Item, 1) = _
     "@" Then L_Item = Right(L_Item, Len(L_Item) - 1)
     
    If UCase(L_Item) = UCase(Trim(LP_Usuario)) Then
        LP_Forma_Canal.OL_Usuarios.RemoveItem (L_i)
        Exit For
    End If
    If GV_DoEVENTS Then DoEvents
Next L_i
Exit Sub
Etiqueta_Error:
ME_Muestra_Error
End Sub

Sub MM_Indicar_Operador(LP_Usuario As String, LP_Forma_Canal As OF_Hablar_Canal, LP_Signo As String)
' /*********************************************************/
' Este procedimiento muestra en una ventana que un usuario
' es o no operador de un canal.  En otras palabras, coloca
' el signo @ precediendo el alias de los operadores.
' LP_Usuario: es el usuario al que se le asignarán o
' quitarán los derechos de operador.
' LP_Forma_Canal: Es la ventana que representa el canal
' donde se definirán los cambios.
' LP_Signo: Indica si se asignarán (+) o revocarán (-)
' los derechos de operador.
' /*********************************************************/

On Error GoTo Etiqueta_Error:

Dim L_i As Long
Dim L_Item As String

For L_i = 0 To (LP_Forma_Canal.OL_Usuarios.ListCount - 1)
 L_Item = Trim(LP_Forma_Canal.OL_Usuarios.List(L_i))
 If Left(L_Item, 1) = "@" Then
   L_Item = Right(L_Item, Len(L_Item) - 1)
 End If
 If UCase(L_Item) = UCase(Trim(LP_Usuario)) Then
   If LP_Signo = "+" Then
    LP_Forma_Canal.OL_Usuarios.List(L_i) = "@" & L_Item
   Else
    LP_Forma_Canal.OL_Usuarios.List(L_i) = L_Item
   End If
   Exit For
 End If
 If GV_DoEVENTS Then DoEvents
Next L_i
Exit Sub
Etiqueta_Error:
ME_Muestra_Error
End Sub

Function MM_Enviar_Mensaje(LP_Mensaje$, LP_Cual) As Boolean
' /*********************************************************/
' Procedimiento para enviar un mensaje a un socket
' especifico (A un servidor específico representado por el
' socket).
' LP_Mensaje: Mensaje a enviar.
' LP_Cual: Indice en el arreglo de sockets que representa
' el servidor al que enviaremos el mensaje.
' /*********************************************************/

On Error GoTo Etiqueta_Error:

Dim L_res%
Dim L_Texto$

MM_Enviar_Mensaje = False
L_Texto = Trim(LP_Mensaje) & GV_EOD

L_res = _
send(GV_Sockets(LP_Cual).socket, L_Texto, Len(L_Texto), 0)

If L_res = SOCKET_ERROR Then
    L_res = WSAGetLastError()
    MG_Mensaje ME_WsockError(L_res)
    Exit Function
 End If
MM_Enviar_Mensaje = True
Exit Function
Etiqueta_Error:
ME_Muestra_Error
 
End Function

Function MM_EsMiNick(Lp_Nick, LP_Socket) As Boolean
' /*********************************************************/
' Función que determina si un alias es igual al registrado
' en la conexión.
' /*********************************************************/
On Error GoTo Etiqueta_Error:

If UCase(Trim(Lp_Nick)) = _
  UCase(Trim(GV_Sockets(LP_Socket).Nick)) Then
  MM_EsMiNick = True
Else
   MM_EsMiNick = False

End If

Exit Function
Etiqueta_Error:
ME_Muestra_Error

End Function

Function MM_Obtener_Canal(ByVal LP_Params$, LP_Simbolo$, LP_Canal$) As String
' /*********************************************************/
' Obtiene el nombre de un canal de un string.
'
' LP_Params: String en el que está contenido el canal.
' LP_Simbolo: Simbolo delimitador del string.
' /*********************************************************/

On Error GoTo Etiqueta_Error:

LP_Canal = MM_Obtener_Mensaje_Parametros(1, LP_Params, ":")
 MM_Obtener_Canal = _
 Trim(Right(LP_Canal, Len(LP_Canal) - InStr(LP_Canal, _
 LP_Simbolo)))
 
Exit Function

Etiqueta_Error:
ME_Muestra_Error

End Function

Function MM_Obtener_Mensaje_Parametros(LP_Caso%, ByVal LP_Params$, LP_Simbolo$) As String
' /*********************************************************/
' Obtiene un substring de un string de acuerdo a una serie
' de criteros utilizados.
' Ej: si se toma un substring desde la izquierda hasta que
' encuentra un símbolo.
' si se toma un substring desde la derecha hasta que encuentra
' un símbolo, etc.
'
' LP_Caso: Indica el caso a aplicar (Izquierda, derecha,
' centro, etc).
' LP_Params: String en el que se encuentra el substring a
' encontrar.
' LP_Simbolo: Simbolo delimitador.
' /*********************************************************/

On Error GoTo Etiqueta_Error:

Dim L_Posicion%
Dim L_Parametro$
L_Posicion = InStr(LP_Params, LP_Simbolo)
Select Case LP_Caso
 Case 1
    MM_Obtener_Mensaje_Parametros = _
    Trim(Left(LP_Params, L_Posicion - 1))
 Case 2
    MM_Obtener_Mensaje_Parametros = _
    Trim(Right(LP_Params, Len(LP_Params) - L_Posicion))
 Case 3
    MM_Obtener_Mensaje_Parametros = _
    Trim(Mid(LP_Params, L_Posicion + 1, _
    Len(LP_Params) - L_Posicion + 1))
 Case 4
    L_Parametro = _
    Trim(Right(LP_Params, Len(LP_Params) - L_Posicion))
    MM_Obtener_Mensaje_Parametros = _
    Trim(Left(L_Parametro, InStr(L_Parametro, LP_Simbolo) - 1))
End Select

Exit Function
Etiqueta_Error:
ME_Muestra_Error

End Function

Function MM_Obtener_Socket(LP_Cual) As Boolean
' /*********************************************************/
' Función que obtiene un handle para un socket y lo asigna
' a la casilla de arreglo de sockets indexado por el
' parametro LP_Cual
' /*********************************************************/

On Error GoTo Etiqueta_Error:

Dim L_socket As Long
Dim L_Status As Integer
Dim L_Texto As String
Dim L_Sockaddr_in As sockaddr_in

MM_Obtener_Socket = False
' Obtener un Socket de Winsock

' Utilizar la función <Socket> del API de Winsock, para
' crear un socket Para la familia Internet, de tipo Stream,
' utilizando el protocolo TCP

GV_Sockets(LP_Cual).socket = _
  socket(PF_INET, SOCK_STREAM, IPPROTO_TCP)
If GV_Sockets(LP_Cual).socket = INVALID_SOCKET Then
    MG_Mensaje ME_WsockError(WSAGetLastError())
    Exit Function
End If
  
' Si llega a este punto es que pudo crear el Socket
' exitosamente

MM_Obtener_Socket = True

Exit Function
Etiqueta_Error:
ME_Muestra_Error

End Function

Function MM_Parsear_Mensaje(LP_Men$, LP_Prefijo$, LP_Comando$, LP_Params$) _
      As Byte

' /*********************************************************/
' Función que parsea un mensaje y retorna que tipo de comando
' es, asi como las 3 diferentes partes de un mensaje:
' El prefijo , el comando y los parametros del comando.
' Si la función retorna 1 significa que es un codigo de error
' o reply de 3 digitos Si la función retorna 0 entonces es un
' comando(generalmente definido en el servidor)

' Gramatica del Mensaje en BNF
'------------------------------------------------------------
'<mensaje>::=[':'<prefijo><espacio>]<comando><parametros> CR LF
'<prefijo>::=<Nombre Servidor>|<nick>['!'<user>]['@'<host>]
'<comando>::=<letra>{<letra>}|<numero> <numero> <numero>
'<espacio>::=' ' {' '}
'<parametros>::=<espacio>[':'<cola>|<contenido> <parametros> ]
'<contenido>::=<Cualquier secuencia no vacia de carateres,esto es _
               que no se debe incluir espacios CR, LF, Null y el _
               primero no sea ':' >
' <cola> ::= <Cualquier secuencia vacia de caracteres, _
              exceptuando Null,CR yLF>
'------------------------------------------------------------

' /*********************************************************/

On Error GoTo Etiqueta_Error:

Dim L_pos As Integer
Dim L_Num As Integer

LP_Prefijo = "": LP_Comando = "": LP_Params = ""
L_pos = 1

' Verificar si vienen dos puntos
If Mid(Trim(LP_Men), 1, 1) = ":" Then

   ' Si estan los dos puntos, por fuerza tiene que venir un
   ' prefijo por lo tanto el prefijo esta a partir de la
   ' posición 2 hasta que vuelva a encontrarse un espacio.
   ' "Segun la gramatica....:)
   LP_Prefijo = Trim(Mid(LP_Men, 2, InStr(2, LP_Men, " ") - 2))
   L_pos = InStr(2, LP_Men, " ") + 1 ' Ultima posicion

End If
L_Num = InStr(L_pos, LP_Men, " ")
If L_Num <> 0 Then
  LP_Comando = _
  Trim(Mid(LP_Men, L_pos, InStr(L_pos, LP_Men, " ") - L_pos))
  L_pos = InStr(L_pos, LP_Men, " ")
  
    If L_pos < Len(LP_Men) Then
       LP_Params = Mid(LP_Men, L_pos, Len(LP_Men) - L_pos - 1)
       'LP_Params = Mid(LP_Men, L_pos, Len(LP_Men) - L_pos)
    End If
Else
   LP_Comando = ""
   LP_Params = ""
End If
If IsNumeric(LP_Comando) Then
    MM_Parsear_Mensaje = 1
Else
    MM_Parsear_Mensaje = 0
End If
LP_Params = Trim(LP_Params)

Exit Function
Etiqueta_Error:
ME_Muestra_Error

End Function

Sub MM_Registrar_Usuario(Lp_Nick$, LP_USER$, LP_Cual%)

' /*********************************************************/
' Procedimiento para registrar un usuario en un servidor.
' Envía el comando user seguido del comando nick.
' ej: user rav : david.intertel.hn : rogger
'     nick ravm
'
' Lp_Nick: Alias de usuario a registrar.
' LP_USER: Nombre real de usuario en el servidor local.
' LP_Cual: Socket que representa el servidor.
' /*********************************************************/

On Error GoTo Etiqueta_Error:

Dim L_Mensaje As String
Dim L_Status As Integer
        
' Primero armamos el Alias
L_Mensaje = "NICK " & Trim(Lp_Nick) & vbCrLf

' Luego se arma el User
L_Status = _
    send(GV_Sockets(LP_Cual).socket, L_Mensaje, Len(L_Mensaje), 0)
   If L_Status = SOCKET_ERROR Then
      L_Status = WSAGetLastError()
      MV_Pone_Mensaje False, _
      GV_VENTANAS_Estatus(GV_Sockets(LP_Cual).Ventana), _
      "Envio de información de registro no se pudo efectuar " + _
      ME_WsockError(L_Status) & GV_EOD, vbBlue
                        
   
   End If

L_Mensaje = "USER " + LP_USER & vbCrLf
   ' ejemplo :
   ' Nick erg
   ' User erg unitec.hn unitec.hn : erg
   ' Para estos dos comandos ver el documento de
   ' principal de ERG2
   ' para mayor explicación
   ' enviar el mensaje
   
   L_Status = _
    send(GV_Sockets(LP_Cual).socket, L_Mensaje, Len(L_Mensaje), 0)
   If L_Status = SOCKET_ERROR Then
      L_Status = WSAGetLastError()
      MV_Pone_Mensaje False, _
      GV_VENTANAS_Estatus(GV_Sockets(LP_Cual).Ventana), _
      "Envio de información de registro no se pudo efectuar " + _
      ME_WsockError(L_Status) & GV_EOD, vbBlue
                        
   
   End If

Exit Sub
Etiqueta_Error:
ME_Muestra_Error

End Sub

Sub MM_Inicializa_WINSOCK()
' /******************************************************/
' Procedimiento que se encarga de Inicializar la interfase
' con Winsock.
' Primero verifica que la interfase con WINSOCK, no se haya
' inicializado previamente, si esta no ha sido
' inicializada entonces se inicializa
' con la funcion  WSAStartup() del API de WINSOCK
' en caso de existir un error este es presentado al usuario
' caso contrario se dada por inicializada la interfase con
' WINSOCK
' /******************************************************/

Dim L_Resultado%
Dim L_Texto As String

 If Not GV_WINSOCK_ARRIBA Then 'WINSOCK ARRIBA ?
  'No, entonces levantarlo
  L_Resultado = WSAStartup(WSA_VERSION(1, 1), GV_WSAdata)
  If L_Resultado <> 0 Then 'Hubo Error ?
    'Si, Muestrelo al usuario       True
    MG_Mensaje ME_WsockError(L_Resultado)
   Else 'No , WINSOCK Inicializado
    GV_WINSOCK_ARRIBA = True
   End If
     
 End If
End Sub

Function MM_Tipo_Direccion(LP_Dirr) As Boolean
' /******************************************************/
' Retorna true si la dirección que viene en LP_Dirr está
' en formato numérico.  De
' lo contrario retorna false, significando que la
' dirección se encuentra en formato
' de caracteres.
' /*******************************************************/

On Error GoTo Etiqueta_Error:

If IsNumeric(Mid(LP_Dirr, 1, 3)) Then
  MM_Tipo_Direccion = True
Else
 MM_Tipo_Direccion = False
End If
  

Exit Function
Etiqueta_Error:
ME_Muestra_Error

End Function
Sub MM_Procesa_Comandos_Numericos(LP_Socket%, LP_Prefijo$, LP_Comando$, LP_Params$, LP_Status%)
' /*********************************************************/
' Este procedimiento se encarga de procesar los comando
' numéricos provenientes de un servidor y enviarlos al
' lugar correspondiente de la aplicación.
'
' LP_Socket: Indice que representa el servidor del cual
' provino el mensaje.
' LP_Prefijo: Prefijo del mensaje(Ver gramática de mensajes).
' LP_Comando: Comando del mensaje(Ver gramática de mensajes).
' LP_Params: Parámetros del mensaje(Ver gramática de mensajes).
' /**********************************************************/

On Error GoTo Etiqueta_Error:

Dim L_vent%, L_i%, L_j%
Dim L_Usuario As ES_USUARIO
Dim L_Datos(10) As String, L_Botado$
Dim L_Lista_Parametros() As String

'l_datos(1) = canal l_datos(2) = nick  L_Datos(3)=mensaje

Select Case CInt(LP_Comando)
    
    'Respuestas a comandos
   
Case 1 'Aceptar Alias ****
   MV_Pone_Mensaje False, _
   GV_VENTANAS_Estatus(GV_Sockets(LP_Socket).Ventana), _
   MM_Obtener_Mensaje_Parametros(2, LP_Params, ":") + _
   GV_EOD, GV_Morado
        
   GV_Sockets(LP_Socket).Nick = _
   MM_Obtener_Mensaje_Parametros(1, LP_Params, ":")
   GV_VENTANAS_Estatus(GV_Sockets(LP_Socket).Ventana).Caption _
   = GV_Sockets(LP_Socket).Direcc + "(" + _
   CStr(GV_Sockets(LP_Socket).Puerto) + ") Estado : [CONECTADO]" + _
   " :¬) = " + GV_Sockets(LP_Socket).Nick
   
   If GV_Explorador Then
    
    OF_Explorador.OA_Explorador.Nodes.Item( _
    "S" + CStr(GV_VENTANAS_Estatus( _
    GV_Sockets(LP_Socket).Ventana).Tag)).Text = _
    GV_Sockets(LP_Socket).Direcc + " Estado : [CONECTADO]" + _
    " :¬) = " + GV_Sockets(LP_Socket).Nick
   End If

Case 2, 3 '****
   MV_Pone_Mensaje False, _
   GV_VENTANAS_Estatus(GV_Sockets(LP_Socket).Ventana), _
   MM_Obtener_Mensaje_Parametros(2, LP_Params, ":") + _
   GV_EOD, GV_Morado
      
Case 4 '****
  MV_Pone_Mensaje False, _
  GV_VENTANAS_Estatus(GV_Sockets(LP_Socket).Ventana), _
  MM_Obtener_Mensaje_Parametros(2, LP_Params, " ") + _
  GV_EOD + GV_EOD, GV_Morado
      
Case 301    'RPL_AWAY ****
  MV_Pone_Mensaje False, _
  GV_VENTANAS_Estatus(GV_Sockets(LP_Socket).Ventana), _
  "User Is Marked As Away/Usuario No Se Encuentra Disponible ", _
  GV_Morado
        
  MV_Pone_Mensaje False, _
  GV_VENTANAS_Estatus(GV_Sockets(LP_Socket).Ventana), _
  MM_Obtener_Mensaje_Parametros(2, LP_Params, " ") + _
  GV_EOD + GV_EOD, GV_Verde
         
Case 302    'RPL_USERHOST ****
  MV_Pone_Mensaje False, _
  GV_VENTANAS_Estatus(GV_Sockets(LP_Socket).Ventana), _
  LP_Params + GV_EOD + GV_EOD, GV_Morado
                         
Case 303    'RPL_ISON ****
  MV_Pone_Mensaje False, _
  GV_VENTANAS_Estatus(GV_Sockets(LP_Socket).Ventana), _
  "Found Nicks/Alias Encontrados " + LP_Params + _
  GV_EOD + GV_EOD, GV_Morado
              
Case 305    'RPL_UNAWAY ****
  MV_Pone_Mensaje False, _
  GV_VENTANAS_Estatus(GV_Sockets(LP_Socket).Ventana), _
  LP_Params + "/Ha Sido Marcado Disponible" + _
  GV_EOD + GV_EOD, GV_Morado
        
Case 306    'RPL_NOWAWAY ****
  MV_Pone_Mensaje False, _
  GV_VENTANAS_Estatus(GV_Sockets(LP_Socket).Ventana), _
  LP_Params + "/Ha Sido Marcado No Disponible" + _
  GV_EOD + GV_EOD, GV_Morado
        
Case 311    'RPL_WHOISUSER ****
  MV_Pone_Mensaje False, _
  GV_VENTANAS_Estatus(GV_Sockets(LP_Socket).Ventana), _
  "Internet User Address/Dirección De Usuario En Internet " + _
  GV_EOD, GV_Azul
  L_Datos(1) = MM_Obtener_Mensaje_Parametros(2, LP_Params, " ")
  MV_Pone_Mensaje False, _
  GV_VENTANAS_Estatus(GV_Sockets(LP_Socket).Ventana), _
  "   " + L_Datos(1) + GV_EOD, GV_Morado
 
  L_Datos(2) = MM_Obtener_Mensaje_Parametros(1, L_Datos(1), " ")
  L_vent = MV_Busca_Ventana_Usuario(L_Datos(2), LP_Socket)
  If L_vent <> 0 Then
   MV_Pone_Mensaje False, _
   GV_VENTANAS_Usuario(L_vent), _
   "Internet User Address/Dirección De Usuario En Internet " + _
   GV_EOD, GV_Azul
  
   MV_Pone_Mensaje False, _
   GV_VENTANAS_Usuario(L_vent), "   " + _
   L_Datos(1) + GV_EOD, GV_Morado
          
  End If
      
Case 312    'RPL_WHOISSERVER ****
  MV_Pone_Mensaje False, _
  GV_VENTANAS_Estatus(GV_Sockets(LP_Socket).Ventana), _
  "IRC Server Name/Nombre Servidor IRC " + GV_EOD, GV_Azul
  L_Datos(1) = MM_Obtener_Mensaje_Parametros(2, LP_Params, " ")
  MV_Pone_Mensaje False, _
  GV_VENTANAS_Estatus(GV_Sockets(LP_Socket).Ventana), _
  "   " + L_Datos(1) + GV_EOD, GV_Morado
  L_Datos(2) = MM_Obtener_Mensaje_Parametros(1, L_Datos(1), " ")
  L_vent = MV_Busca_Ventana_Usuario(L_Datos(2), LP_Socket)
  If L_vent <> 0 Then
   MV_Pone_Mensaje False, _
   GV_VENTANAS_Usuario(L_vent), _
   "IRC Server Name/Nombre Servidor IRC " + GV_EOD, GV_Azul
  
   MV_Pone_Mensaje False, _
   GV_VENTANAS_Usuario(L_vent), _
   "   " + L_Datos(1) + GV_EOD, GV_Morado
          
  End If
      
Case 313    'RPL_WHOISOPERATOR ****
  MV_Pone_Mensaje False, _
  GV_VENTANAS_Estatus(GV_Sockets(LP_Socket).Ventana), _
  "   " + MM_Obtener_Mensaje_Parametros(2, LP_Params, " ") + _
  "/Es Operador De IRC" + GV_EOD + GV_EOD, GV_Morado
        
Case 314  'RPL_WHOWASUSER probar en otros servers internet
  MV_Pone_Mensaje False, _
  GV_VENTANAS_Estatus(GV_Sockets(LP_Socket).Ventana), _
  LP_Prefijo + " " + LP_Comando + " " + LP_Params + _
  GV_EOD + GV_EOD, GV_Morado
        
Case 317    'RPL_WHOISIDLE ****
  L_Datos(1) = MM_Obtener_Mensaje_Parametros(2, LP_Params, " ")
  MV_Pone_Mensaje False, _
  GV_VENTANAS_Estatus(GV_Sockets(LP_Socket).Ventana), _
  "   " + L_Datos(1) + "/Segundos Inactivo" + GV_EOD, GV_Morado
  L_Datos(2) = MM_Obtener_Mensaje_Parametros(1, L_Datos(1), " ")
  L_vent = MV_Busca_Ventana_Usuario(L_Datos(2), LP_Socket)
  If L_vent <> 0 Then
    MV_Pone_Mensaje False, _
    GV_VENTANAS_Usuario(L_vent), "   " + _
    L_Datos(1) + "/Segundos Inactivo" + GV_EOD, GV_Morado
  End If

Case 318    'RPL_ENDOFWHOIS ****
  L_Datos(1) = MM_Obtener_Mensaje_Parametros(2, LP_Params, " ")
  MV_Pone_Mensaje False, _
  GV_VENTANAS_Estatus(GV_Sockets(LP_Socket).Ventana), _
  L_Datos(1) + "/Fin De Lista" + GV_EOD + GV_EOD, GV_Azul
  L_Datos(2) = MM_Obtener_Mensaje_Parametros(1, L_Datos(1), " ")
  L_vent = MV_Busca_Ventana_Usuario(L_Datos(2), LP_Socket)
  If L_vent <> 0 Then
    MV_Pone_Mensaje False, _
    GV_VENTANAS_Usuario(L_vent), L_Datos(1) + _
    "/Fin De Lista" + GV_EOD + GV_EOD, GV_Azul
  End If
    
Case 319    'RPL_WHOISCHANNELS ****
  MV_Pone_Mensaje False, _
  GV_VENTANAS_Estatus(GV_Sockets(LP_Socket).Ventana), _
  "Channels Where User Is In/Canales En Los Que Se " + _
  " Encuentra El Usuario " + GV_EOD, GV_Azul
  L_Datos(1) = MM_Obtener_Mensaje_Parametros(2, LP_Params, " ")
  MV_Pone_Mensaje False, _
  GV_VENTANAS_Estatus(GV_Sockets(LP_Socket).Ventana), _
  "   " + L_Datos(1) + GV_EOD + GV_EOD, GV_Morado
  L_Datos(2) = MM_Obtener_Mensaje_Parametros(1, L_Datos(1), " ")
  L_vent = MV_Busca_Ventana_Usuario(L_Datos(2), LP_Socket)
  If L_vent <> 0 Then
    MV_Pone_Mensaje False, _
    GV_VENTANAS_Usuario(L_vent), _
    "Channels Where User Is In/Canales En Los Que " + _
    " Se Encuentra El Usuario " + GV_EOD, GV_Azul
    
    MV_Pone_Mensaje False, _
    GV_VENTANAS_Usuario(L_vent), "   " + L_Datos(1) + _
    GV_EOD + GV_EOD, GV_Morado
  End If
 
      
Case 369    'RPL_ENDOFWHOWAS probar en otros servers internet
  MV_Pone_Mensaje False, _
  GV_VENTANAS_Estatus(GV_Sockets(LP_Socket).Ventana), _
  LP_Prefijo + " " + LP_Comando + " " + LP_Params + _
  GV_EOD + GV_EOD, GV_Morado
      
Case 321    'RPL_LISTSTART ****
  'Empieza un listado de Canales
   MV_Pone_Mensaje False, _
   GV_VENTANAS_Estatus(GV_Sockets(LP_Socket).Ventana), _
   MM_Obtener_Mensaje_Parametros(2, LP_Params, " ") + _
   "/Inicio Lista Canales" + GV_EOD + GV_EOD, GV_Azul
   
   L_vent = MV_Busca_Ventana_Lista_Canales(LP_Socket)
   If L_vent = 0 Then
      L_vent = MV_CreaVentana_Lista_Canales(LP_Socket)
      GV_VENTANAS_Lista_Canales(L_vent).Caption = _
      "Lista de Canales : [" + _
      GV_Sockets(LP_Socket).Direcc + "]"
   End If
   GV_VENTANAS_Lista_Canales(L_vent).Show
   GV_VENTANAS_Lista_Canales(L_vent).OL_Estatus.Clear
                          
Case 322    'RPL_LIST ****
  'Lista de Canales "NIck Canal 1 : Hola "
  L_Datos(1) = MM_Obtener_Mensaje_Parametros(3, LP_Params, " ")
  L_vent = MV_Busca_Ventana_Lista_Canales(LP_Socket)
  If L_vent <> 0 Then
   GV_VENTANAS_Lista_Canales( _
    L_vent).OL_Estatus.AddItem L_Datos(1)
  End If

Case 323    'RPL_LISTEND ****
  MV_Pone_Mensaje False, _
  GV_VENTANAS_Estatus(GV_Sockets(LP_Socket).Ventana), _
  MM_Obtener_Mensaje_Parametros(2, LP_Params, " ") + _
  "/Fin Lista de Canales" + GV_EOD + GV_EOD, GV_Azul
       
Case 324    'RPL_CHANNELMODEIS
  MV_Pone_Mensaje False, _
  GV_VENTANAS_Estatus(GV_Sockets(LP_Socket).Ventana), _
  LP_Prefijo + " " + LP_Comando + " " + LP_Params + _
  GV_EOD + GV_EOD, GV_Morado
       
Case 331    'RPL_NOTOPIC ****
  MV_Pone_Mensaje False, _
  GV_VENTANAS_Estatus(GV_Sockets(LP_Socket).Ventana), _
  LP_Params + "/Tópico No Establecido" + GV_EOD + _
  GV_EOD, GV_Morado
       
Case 332    'RPL_TOPIC ****
  'Establecer topico de un canal
  L_Datos(1) = MM_Obtener_Canal(LP_Params, " ", L_Datos(1))
  L_vent = MV_Busca_Ventana_Canal(L_Datos(1), LP_Socket)
  If L_vent <> 0 Then
    GV_VENTANAS_Canal(L_vent).Caption = L_Datos(1) + " [" + _
    GV_Sockets(LP_Socket).Direcc + "] : " + _
    MM_Obtener_Mensaje_Parametros(2, LP_Params, ":")
  End If
  
Case 333    'RPL_TOPIC II ****
        
Case 341    'RPL_INVITING ****
  MV_Pone_Mensaje False, _
  GV_VENTANAS_Estatus(GV_Sockets(LP_Socket).Ventana), _
  "Inviting/Invitando a ", GV_Morado
  
  MV_Pone_Mensaje False, _
  GV_VENTANAS_Estatus(GV_Sockets(LP_Socket).Ventana), _
  MM_Obtener_Mensaje_Parametros(2, LP_Params, " ") + _
  GV_EOD + GV_EOD, GV_Verde
       
Case 342    'RPL_SUMMONING probar con net
  MV_Pone_Mensaje False, _
  GV_VENTANAS_Estatus(GV_Sockets(LP_Socket).Ventana), _
  LP_Prefijo + " " + LP_Comando + " " + LP_Params + _
  GV_EOD + GV_EOD, GV_Morado
        
Case 351    'RPL_VERSION ****
  MV_Pone_Mensaje False, _
  GV_VENTANAS_Estatus(GV_Sockets(LP_Socket).Ventana), _
  "Version/Versión " + _
  MM_Obtener_Mensaje_Parametros(2, LP_Params, " ") + _
  GV_EOD + GV_EOD, GV_Morado
        
Case 352    'RPL_WHOREPLY ****
  MV_Pone_Mensaje False, _
  GV_VENTANAS_Estatus(GV_Sockets(LP_Socket).Ventana), _
  "    " + MM_Obtener_Mensaje_Parametros(2, LP_Params, " ") _
  + GV_EOD + GV_EOD, GV_Morado
        
Case 315    'RPL_ENDOFWHO ****
  MV_Pone_Mensaje False, _
  GV_VENTANAS_Estatus(GV_Sockets(LP_Socket).Ventana), _
  LP_Params + "/Fin De Lista Who" + GV_EOD + GV_EOD, GV_Azul
        
Case 353    'RPL_NAMREPLY ****
  'Retorna lista de alias conectados a un canal
   ReDim Preserve L_Lista_Parametros(1)
  L_Datos(1) = MM_Obtener_Canal(LP_Params, "=", L_Datos(1))
  MM_Obtener_Lista_Parametros _
   MM_Obtener_Mensaje_Parametros(2, LP_Params, ":"), _
   L_Lista_Parametros()
  
  L_vent = MV_Busca_Ventana_Canal(L_Datos(1), LP_Socket)
  If L_vent <> 0 Then
   MV_Pone_Mensaje False, _
   GV_VENTANAS_Estatus(GV_Sockets(LP_Socket).Ventana), _
   "    " + _
   MM_Obtener_Mensaje_Parametros(2, LP_Params, "=") + _
   GV_EOD + GV_EOD, GV_Morado
   If GV_Sockets(LP_Socket).JOIN_PENDIENTE Then
     MM_Cargar_Usuarios_A_Canal L_Lista_Parametros(), _
     GV_VENTANAS_Canal(L_vent)
   End If
  End If
      
Case 366    'RPL_ENDOFNAMES ****
  MV_Pone_Mensaje False, _
  GV_VENTANAS_Estatus(GV_Sockets(LP_Socket).Ventana), _
  MM_Obtener_Mensaje_Parametros(2, LP_Params, " ") + _
  "/Fin De Lista De Usuarios En Canal" + GV_EOD + GV_EOD, _
  GV_Azul
  
  GV_Sockets(LP_Socket).JOIN_PENDIENTE = False
     
Case 364    'RPL_LINKS ****
  MV_Pone_Mensaje False, _
  GV_VENTANAS_Estatus(GV_Sockets(LP_Socket).Ventana), _
  "    " + MM_Obtener_Mensaje_Parametros(2, LP_Params, " ") _
  + GV_EOD + GV_EOD, GV_Morado
        
Case 365    'RPL_ENDOFLINKS ****
  MV_Pone_Mensaje False, _
  GV_VENTANAS_Estatus(GV_Sockets(LP_Socket).Ventana), _
  MM_Obtener_Mensaje_Parametros(2, LP_Params, ":") + _
  "/Fin De Lista De Links" + GV_EOD + GV_EOD, GV_Azul
        
Case 367    'RPL_BANLIST
  MV_Pone_Mensaje False, _
  GV_VENTANAS_Estatus(GV_Sockets(LP_Socket).Ventana), _
  MM_Obtener_Mensaje_Parametros(2, LP_Params, " ") + _
  GV_EOD + GV_EOD, GV_Morado
        
Case 368    'RPL_ENDOFBANLIST
  MV_Pone_Mensaje False, _
  GV_VENTANAS_Estatus(GV_Sockets(LP_Socket).Ventana), _
  LP_Params + GV_EOD + GV_EOD, GV_Morado
        
Case 371    'RPL_INFO ****
  MV_Pone_Mensaje False, _
  GV_VENTANAS_Estatus(GV_Sockets(LP_Socket).Ventana), _
  "    " + MM_Obtener_Mensaje_Parametros(2, LP_Params, ":") _
  + GV_EOD + GV_EOD, GV_Morado
        
Case 374    'RPL_ENDOFINFO ****
  MV_Pone_Mensaje False, _
  GV_VENTANAS_Estatus(GV_Sockets(LP_Socket).Ventana), _
  MM_Obtener_Mensaje_Parametros(2, LP_Params, ":") + _
  "/Fin De Información " + GV_EOD + GV_EOD, GV_Azul
        
Case 375    'RPL_MOTDSTART ****
  MV_Pone_Mensaje False, _
  GV_VENTANAS_Estatus(GV_Sockets(LP_Socket).Ventana), _
  MM_Obtener_Mensaje_Parametros(2, LP_Params, " ") + _
  "/Inicio de Mensaje" + GV_EOD + GV_EOD, GV_Azul
        
Case 372    'RPL_MOTD ****
  MV_Pone_Mensaje False, _
  GV_VENTANAS_Estatus(GV_Sockets(LP_Socket).Ventana), _
  "    " + MM_Obtener_Mensaje_Parametros(2, LP_Params, ":") _
  + GV_EOD + GV_EOD, GV_Morado
        
Case 376    'RPL_ENDOFMOTD ****
  MV_Pone_Mensaje False, _
  GV_VENTANAS_Estatus(GV_Sockets(LP_Socket).Ventana), _
  MM_Obtener_Mensaje_Parametros(2, LP_Params, ":") + _
  "/Fin De Mensaje" + GV_EOD + GV_EOD, GV_Azul
        
Case 381    'RPL_YOUREOPER ****
  MV_Pone_Mensaje False, _
  GV_VENTANAS_Estatus(GV_Sockets(LP_Socket).Ventana), _
  MM_Obtener_Mensaje_Parametros(2, LP_Params, ":") + _
  "/Usted Se Ha Convertido En Operador IRC" + _
  GV_EOD + GV_EOD, GV_Morado
        
Case 382    'RPL_REHASHING ****
  MV_Pone_Mensaje False, _
  GV_VENTANAS_Estatus(GV_Sockets(LP_Socket).Ventana), _
  MM_Obtener_Mensaje_Parametros(2, LP_Params, " ") + _
  "/Reprocesando INI" + GV_EOD + GV_EOD, GV_Morado
        
Case 391    'RPL_TIME ****
  MV_Pone_Mensaje False, _
  GV_VENTANAS_Estatus(GV_Sockets(LP_Socket).Ventana), _
  MM_Obtener_Mensaje_Parametros(2, LP_Params, ":") + _
  GV_EOD + GV_EOD, GV_Morado
        
Case 392    'RPL_USERSSTART ****
  MV_Pone_Mensaje False, _
  GV_VENTANAS_Estatus(GV_Sockets(LP_Socket).Ventana), _
  MM_Obtener_Mensaje_Parametros(2, LP_Params, ":") + _
  "/Lista De Usuarios En Servidor" + GV_EOD + GV_EOD, GV_Azul
        
Case 393    'RPL_USERS ****
  MV_Pone_Mensaje False, _
  GV_VENTANAS_Estatus(GV_Sockets(LP_Socket).Ventana), _
  "    " + MM_Obtener_Mensaje_Parametros(2, LP_Params, ":") _
  + GV_EOD + GV_EOD, GV_Morado
        
Case 394    'RPL_ENDOFUSERS ****
  MV_Pone_Mensaje False, _
  GV_VENTANAS_Estatus(GV_Sockets(LP_Socket).Ventana), _
  MM_Obtener_Mensaje_Parametros(2, LP_Params, ":") + _
  "/Fin De Lista De Usuarios En Servidor" + GV_EOD + GV_EOD, GV_Azul
        
Case 395    'RPL_NOUSERS ****
  MV_Pone_Mensaje False, _
  GV_VENTANAS_Estatus(GV_Sockets(LP_Socket).Ventana), _
  MM_Obtener_Mensaje_Parametros(2, LP_Params, ":") + _
  "/Nadie Ha Ingresado A Este Servidor" + GV_EOD + _
  GV_EOD, GV_Morado
    
Case 200    'RPL_TRACELINK ****
  MV_Pone_Mensaje False, _
  GV_VENTANAS_Estatus(GV_Sockets(LP_Socket).Ventana), _
  "/Usuario" + MM_Obtener_Mensaje_Parametros(2, LP_Params, " ") _
  + GV_EOD + GV_EOD, GV_Morado
        
Case 201    'RPL_TRACECONNECTING ****
  MV_Pone_Mensaje False, _
  GV_VENTANAS_Estatus(GV_Sockets(LP_Socket).Ventana), _
  MM_Obtener_Mensaje_Parametros(2, LP_Params, " ") + _
  GV_EOD + GV_EOD, GV_Morado
        
Case 202    'RPL_TRACEHANDSHAKE ****
  MV_Pone_Mensaje False, _
  GV_VENTANAS_Estatus(GV_Sockets(LP_Socket).Ventana), _
  MM_Obtener_Mensaje_Parametros(2, LP_Params, " ") + _
  GV_EOD + GV_EOD, GV_Morado
        
Case 203    'RPL_TRACEUNKNOWN ****
  MV_Pone_Mensaje False, _
  GV_VENTANAS_Estatus(GV_Sockets(LP_Socket).Ventana), _
  MM_Obtener_Mensaje_Parametros(2, LP_Params, " ") + _
  GV_EOD + GV_EOD, GV_Morado
        
Case 204    'RPL_TRACEOPERATOR ****
  MV_Pone_Mensaje False, _
  GV_VENTANAS_Estatus(GV_Sockets(LP_Socket).Ventana), _
  MM_Obtener_Mensaje_Parametros(2, LP_Params, " ") + _
  GV_EOD + GV_EOD, GV_Morado
       
Case 205    'RPL_TRACEUSER ****
  MV_Pone_Mensaje False, _
  GV_VENTANAS_Estatus(GV_Sockets(LP_Socket).Ventana), _
  "Usuario/ " + MM_Obtener_Mensaje_Parametros(2, LP_Params, " ") _
  + GV_EOD + GV_EOD, GV_Morado
       
Case 206    'RPL_TRACESERVER ****
  MV_Pone_Mensaje False, _
  GV_VENTANAS_Estatus(GV_Sockets(LP_Socket).Ventana), _
  MM_Obtener_Mensaje_Parametros(2, LP_Params, " ") + _
  GV_EOD + GV_EOD, GV_Morado
        
Case 208    'RPL_TRACENEWTYPE ****
  MV_Pone_Mensaje False, _
  GV_VENTANAS_Estatus(GV_Sockets(LP_Socket).Ventana), _
  MM_Obtener_Mensaje_Parametros(2, LP_Params, " ") + _
  GV_EOD + GV_EOD, GV_Morado
        
Case 261    'RPL_TRACELOG ****
  MV_Pone_Mensaje False, _
  GV_VENTANAS_Estatus(GV_Sockets(LP_Socket).Ventana), _
  MM_Obtener_Mensaje_Parametros(2, LP_Params, " ") + _
  GV_EOD + GV_EOD, GV_Morado
        
Case 211    'RPL_STATSLINKINFO ****
  MV_Pone_Mensaje False, _
  GV_VENTANAS_Estatus(GV_Sockets(LP_Socket).Ventana), _
  "    " + MM_Obtener_Mensaje_Parametros(2, LP_Params, ":") _
  + GV_EOD + GV_EOD, GV_Morado
        
Case 212    'RPL_STATSCOMMANDS ****
  MV_Pone_Mensaje False, _
  GV_VENTANAS_Estatus(GV_Sockets(LP_Socket).Ventana), _
  "    " + MM_Obtener_Mensaje_Parametros(2, LP_Params, ":") _
  + GV_EOD + GV_EOD, GV_Morado
        
Case 213    'RPL_STATSCLINE ****
  MV_Pone_Mensaje False, _
  GV_VENTANAS_Estatus(GV_Sockets(LP_Socket).Ventana), _
  "    " + MM_Obtener_Mensaje_Parametros(2, LP_Params, ":") _
  + GV_EOD + GV_EOD, GV_Morado
        
Case 214    'RPL_STATSNLINE ****
  MV_Pone_Mensaje False, _
  GV_VENTANAS_Estatus(GV_Sockets(LP_Socket).Ventana), _
  "    " + MM_Obtener_Mensaje_Parametros(2, LP_Params, ":") _
  + GV_EOD + GV_EOD, GV_Morado
        
Case 215    'RPL_STATSILINE ****
  MV_Pone_Mensaje False, _
  GV_VENTANAS_Estatus(GV_Sockets(LP_Socket).Ventana), _
  "    " + MM_Obtener_Mensaje_Parametros(2, LP_Params, ":") _
  + GV_EOD + GV_EOD, GV_Morado
        
Case 216    'RPL_STATSKLINE ****
  MV_Pone_Mensaje False, _
  GV_VENTANAS_Estatus(GV_Sockets(LP_Socket).Ventana), _
  "    " + MM_Obtener_Mensaje_Parametros(2, LP_Params, ":") _
  + GV_EOD + GV_EOD, GV_Morado
        
Case 218    'RPL_STATSYLINE ****
  MV_Pone_Mensaje False, _
  GV_VENTANAS_Estatus(GV_Sockets(LP_Socket).Ventana), _
  "    " + MM_Obtener_Mensaje_Parametros(2, LP_Params, ":") _
  + GV_EOD + GV_EOD, GV_Morado
        
Case 219    'RPL_ENDOFSTATS ****
  MV_Pone_Mensaje False, _
  GV_VENTANAS_Estatus(GV_Sockets(LP_Socket).Ventana), _
  MM_Obtener_Mensaje_Parametros(2, LP_Params, ":") + _
  "/Fin De Reporte De Estatus" + GV_EOD + GV_EOD, GV_Azul
        
Case 241    'RPL_STATSLLINE ****
  MV_Pone_Mensaje False, _
  GV_VENTANAS_Estatus(GV_Sockets(LP_Socket).Ventana), _
  "    " + MM_Obtener_Mensaje_Parametros(2, LP_Params, ":") _
  + GV_EOD + GV_EOD, GV_Morado
        
Case 242    'RPL_STATSUPTIME ****
  MV_Pone_Mensaje False, _
  GV_VENTANAS_Estatus(GV_Sockets(LP_Socket).Ventana), _
  "    " + MM_Obtener_Mensaje_Parametros(2, LP_Params, ":") _
  + GV_EOD + GV_EOD, GV_Morado
        
Case 243    'RPL_STATSOLINE ****
  MV_Pone_Mensaje False, _
  GV_VENTANAS_Estatus(GV_Sockets(LP_Socket).Ventana), _
  "    " + MM_Obtener_Mensaje_Parametros(2, LP_Params, ":") _
  + GV_EOD + GV_EOD, GV_Morado
        
Case 244    'RPL_STATSHLINE ****
  MV_Pone_Mensaje False, _
  GV_VENTANAS_Estatus(GV_Sockets(LP_Socket).Ventana), _
  "    " + MM_Obtener_Mensaje_Parametros(2, LP_Params, ":") _
  + GV_EOD + GV_EOD, GV_Morado
        
Case 221    'RPL_UMODEIS
 MV_Pone_Mensaje False, _
 GV_VENTANAS_Estatus(GV_Sockets(LP_Socket).Ventana), _
 LP_Params + GV_EOD + GV_EOD, GV_Morado
                      
Case 250 '****
  MV_Pone_Mensaje False, _
  GV_VENTANAS_Estatus(GV_Sockets(LP_Socket).Ventana), _
  MM_Obtener_Mensaje_Parametros(2, LP_Params, ":") + _
  GV_EOD + GV_EOD, GV_Morado
        
Case 251    'RPL_LUSERCLIENT  ****
  MV_Pone_Mensaje False, _
  GV_VENTANAS_Estatus(GV_Sockets(LP_Socket).Ventana), _
  MM_Obtener_Mensaje_Parametros(2, LP_Params, ":") + _
  GV_EOD + GV_EOD, GV_Morado
        
Case 252    'RPL_LUSEROP ****
  MV_Pone_Mensaje False, _
  GV_VENTANAS_Estatus(GV_Sockets(LP_Socket).Ventana), _
  MM_Obtener_Mensaje_Parametros(2, LP_Params, " ") + _
  GV_EOD + GV_EOD, GV_Morado
        
Case 253    'RPL_LUSERUNKNOWN ****
  MV_Pone_Mensaje False, _
  GV_VENTANAS_Estatus(GV_Sockets(LP_Socket).Ventana), _
  MM_Obtener_Mensaje_Parametros(2, LP_Params, " ") + _
  GV_EOD + GV_EOD, GV_Morado
              
Case 254    'RPL_LUSERCHANNELS ****
  MV_Pone_Mensaje False, _
  GV_VENTANAS_Estatus(GV_Sockets(LP_Socket).Ventana), _
  MM_Obtener_Mensaje_Parametros(2, LP_Params, " ") + _
  GV_EOD + GV_EOD, GV_Morado
        
Case 255    'RPL_LUSERME ****
  MV_Pone_Mensaje False, _
  GV_VENTANAS_Estatus(GV_Sockets(LP_Socket).Ventana), _
  MM_Obtener_Mensaje_Parametros(2, LP_Params, ":") + _
  GV_EOD + GV_EOD, GV_Morado
        
Case 256    'RPL_ADMINME ****
  MV_Pone_Mensaje False, _
  GV_VENTANAS_Estatus(GV_Sockets(LP_Socket).Ventana), _
  MM_Obtener_Mensaje_Parametros(2, LP_Params, ":") + _
  GV_EOD + GV_EOD, GV_Morado
        
Case 257    'RPL_ADMINLOC1 ****
  MV_Pone_Mensaje False, _
  GV_VENTANAS_Estatus(GV_Sockets(LP_Socket).Ventana), _
  MM_Obtener_Mensaje_Parametros(2, LP_Params, ":") + _
  GV_EOD + GV_EOD, GV_Morado
              
Case 258    'RPL_ADMINLOC2 ****
  MV_Pone_Mensaje False, _
  GV_VENTANAS_Estatus(GV_Sockets(LP_Socket).Ventana), _
  MM_Obtener_Mensaje_Parametros(2, LP_Params, ":") + _
  GV_EOD + GV_EOD, GV_Morado
            
Case 259    'RPL_ADMINEMAIL ****
  MV_Pone_Mensaje False, _
  GV_VENTANAS_Estatus(GV_Sockets(LP_Socket).Ventana), _
  MM_Obtener_Mensaje_Parametros(2, LP_Params, ":") + _
  GV_EOD + GV_EOD, GV_Morado
   
 'errores errores errores errores
        
Case 401    'ERR_NOSUCHNICK ****
  MV_Pone_Mensaje False, _
  GV_VENTANAS_Estatus(GV_Sockets(LP_Socket).Ventana), _
  MM_Obtener_Mensaje_Parametros(2, LP_Params, " ") + _
  "/No Existe Usuario o Canal" + GV_EOD + GV_EOD, GV_Rojo
        
Case 402    'ERR_NOSUCHSERVER ****
  MV_Pone_Mensaje False, _
  GV_VENTANAS_Estatus(GV_Sockets(LP_Socket).Ventana), _
  MM_Obtener_Mensaje_Parametros(2, LP_Params, " ") + _
  "/No Existe Este Servidor" + GV_EOD + GV_EOD, GV_Rojo
        
Case 403    'ERR_NOSUCHCHANNEL ****
  MV_Pone_Mensaje False, _
  GV_VENTANAS_Estatus(GV_Sockets(LP_Socket).Ventana), _
  MM_Obtener_Mensaje_Parametros(2, LP_Params, ":") + _
  "/Canal No Válido" + GV_EOD + GV_EOD, GV_Rojo
        
Case 404    'ERR_CANNOTSENDTOCHAN ****
  MV_Pone_Mensaje False, _
  GV_VENTANAS_Estatus(GV_Sockets(LP_Socket).Ventana), _
  MM_Obtener_Mensaje_Parametros(2, LP_Params, ":") + _
  "/No Puede Enviar Mensaje" + GV_EOD + GV_EOD, GV_Rojo
        
Case 405    'ERR_TOOMANYCHANNELS
  MV_Pone_Mensaje False, _
  GV_VENTANAS_Estatus(GV_Sockets(LP_Socket).Ventana), _
   LP_Params + _
  "/Demasiados Canales" + GV_EOD + GV_EOD, GV_Rojo

              
Case 406    'ERR_WASNOSUCHNICK
  MV_Pone_Mensaje False, _
  GV_VENTANAS_Estatus(GV_Sockets(LP_Socket).Ventana), _
   LP_Params + _
  "/No existía Alias" + GV_EOD + GV_EOD, GV_Rojo
                 
Case 407    'ERR_TOOMANYTARGETS
  MV_Pone_Mensaje False, _
  GV_VENTANAS_Estatus(GV_Sockets(LP_Socket).Ventana), _
   LP_Params + _
  "/Muchos Destinos" + GV_EOD + GV_EOD, GV_Rojo
                 
Case 409    'ERR_NOORIGIN
  MV_Pone_Mensaje False, _
  GV_VENTANAS_Estatus(GV_Sockets(LP_Socket).Ventana), _
   LP_Params + _
  "/No Origen" + GV_EOD + GV_EOD, GV_Rojo
                 
Case 411    'ERR_NORECIPIENT
  MV_Pone_Mensaje False, _
  GV_VENTANAS_Estatus(GV_Sockets(LP_Socket).Ventana), _
   LP_Params + _
  "/No Destino" + GV_EOD + GV_EOD, GV_Rojo
                            
Case 412    'ERR_NOTEXTTOSEND
  MV_Pone_Mensaje False, _
  GV_VENTANAS_Estatus(GV_Sockets(LP_Socket).Ventana), _
   LP_Params + _
  "/No texto que enviar" + GV_EOD + GV_EOD, GV_Rojo
                 
Case 413    'ERR_NOTOPLEVEL
  MV_Pone_Mensaje False, _
  GV_VENTANAS_Estatus(GV_Sockets(LP_Socket).Ventana), _
   LP_Params + _
  "/No nivel superior" + GV_EOD + GV_EOD, GV_Rojo
                 
Case 414    'ERR_WILDTOPLEVEL
  MV_Pone_Mensaje False, _
  GV_VENTANAS_Estatus(GV_Sockets(LP_Socket).Ventana), _
   LP_Params + _
  "/No nivel superior" + GV_EOD + GV_EOD, GV_Rojo
                 
Case 421    'ERR_UNKNOWNCOMMAND
  MV_Pone_Mensaje True, _
  GV_VENTANAS_Estatus(GV_Sockets(LP_Socket).Ventana), _
  MM_Obtener_Mensaje_Parametros(2, LP_Params, ":") + _
  "/Error Comando No Encontrado" + GV_EOD + GV_EOD, GV_Rojo
                 
Case 422    'ERR_NOMOTD
  MV_Pone_Mensaje False, _
  GV_VENTANAS_Estatus(GV_Sockets(LP_Socket).Ventana), _
   LP_Params + _
  "/No hay mensajes del dia" + GV_EOD + GV_EOD, GV_Rojo
                 
Case 423    'ERR_NOADMININFO
  MV_Pone_Mensaje False, _
  GV_VENTANAS_Estatus(GV_Sockets(LP_Socket).Ventana), _
   LP_Params + _
  "/No existe info. Administrativa" + GV_EOD + GV_EOD, GV_Rojo
        
Case 424    'ERR_FILEERROR
  MV_Pone_Mensaje False, _
  GV_VENTANAS_Estatus(GV_Sockets(LP_Socket).Ventana), _
   LP_Params + _
  "/Error de Archivo" + GV_EOD + GV_EOD, GV_Rojo
                 
Case 431    'ERR_NONICKNAMEGIVEN ****
   MV_Pone_Mensaje False, _
   GV_VENTANAS_Estatus(GV_Sockets(LP_Socket).Ventana), _
   MM_Obtener_Mensaje_Parametros(2, LP_Params, ":") + _
   "/No Se Ha Proporcionado Alias" + GV_EOD + GV_EOD, GV_Rojo
        
Case 432    'ERR_ERRONEUSNICKNAME ****
   MV_Pone_Mensaje False, _
   GV_VENTANAS_Estatus(GV_Sockets(LP_Socket).Ventana), _
   MM_Obtener_Mensaje_Parametros(2, LP_Params, ":") + _
   "/Alias No Válido" + GV_EOD + GV_EOD, GV_Rojo
                 
Case 433    'ERR_NICKNAMEINUSE ****
   'Alias ya se encuentra en uso
   If MD_Recupera_Infousuario(L_Usuario) = 0 Then
     If L_Usuario.E_nombre_alterno <> _
      MM_Obtener_Mensaje_Parametros(1, LP_Params, ":") Then
        MM_Enviar_Mensaje "Nick " + _
          L_Usuario.E_nombre_alterno, LP_Socket
     Else
       GV_VENTANAS_Estatus( _
       GV_Sockets(LP_Socket).Ventana).OT_Comando = "NICK "
       GV_VENTANAS_Estatus( _
       GV_Sockets(LP_Socket).Ventana).OT_Comando.SelStart = 6
     End If
     MV_Pone_Mensaje False, _
     GV_VENTANAS_Estatus(GV_Sockets(LP_Socket).Ventana), _
     MM_Obtener_Mensaje_Parametros(2, LP_Params, ":") + _
     "/Alias Ya Se Encuentra En Uso" + GV_EOD + GV_EOD, GV_Rojo
   End If
        
 Case 436    'ERR_NICKCOLLISION
   MV_Pone_Mensaje False, _
   GV_VENTANAS_Estatus(GV_Sockets(LP_Socket).Ventana), _
   LP_Params + _
   "/Colisión de Alias" + GV_EOD + GV_EOD, GV_Rojo
    
 Case 441    'ERR_USERNOTINCHANNEL
   MV_Pone_Mensaje False, _
   GV_VENTANAS_Estatus(GV_Sockets(LP_Socket).Ventana), _
   LP_Params + _
   "/Usuario no esta en canal" + GV_EOD + GV_EOD, GV_Rojo
    
 Case 442    'ERR_NOTONCHANNEL
   MV_Pone_Mensaje False, _
   GV_VENTANAS_Estatus(GV_Sockets(LP_Socket).Ventana), _
   LP_Params + _
   "/No se encuentra en un canal" + GV_EOD + GV_EOD, GV_Rojo
                 
 Case 443    'ERR_USERONCHANNEL
   MV_Pone_Mensaje False, _
   GV_VENTANAS_Estatus(GV_Sockets(LP_Socket).Ventana), _
   LP_Params + _
   "/Usuario esta en canal" + GV_EOD + GV_EOD, GV_Rojo
                 
 Case 444    'ERR_NOLOGIN
   MV_Pone_Mensaje False, _
   GV_VENTANAS_Estatus(GV_Sockets(LP_Socket).Ventana), _
   LP_Params + _
   "/No ha efectuado un login" + GV_EOD + GV_EOD, GV_Rojo
                 
 Case 445    'ERR_SUMMONDISABLED
   MV_Pone_Mensaje False, _
   GV_VENTANAS_Estatus(GV_Sockets(LP_Socket).Ventana), _
   LP_Params + _
   "/SUMMON Deshabilitado" + GV_EOD + GV_EOD, GV_Rojo
                 
 Case 446    'ERR_USERSDISABLED
   MV_Pone_Mensaje False, _
   GV_VENTANAS_Estatus(GV_Sockets(LP_Socket).Ventana), _
   LP_Params + _
   "/Usuario deshabilitado" + GV_EOD + GV_EOD, GV_Rojo
        
 Case 451    'ERR_NOTREGISTERED
   MV_Pone_Mensaje False, _
   GV_VENTANAS_Estatus(GV_Sockets(LP_Socket).Ventana), _
   MM_Obtener_Mensaje_Parametros(2, LP_Params, " ") + _
   "/Primero debe registrarse" + GV_EOD + GV_EOD, GV_Rojo
    
 Case 461    'ERR_NEEDMOREPARAMS ****
   MV_Pone_Mensaje False, _
   GV_VENTANAS_Estatus(GV_Sockets(LP_Socket).Ventana), _
   MM_Obtener_Mensaje_Parametros(2, LP_Params, " ") + _
   "/No Existen Suficientes Parámetros" + GV_EOD + GV_EOD, GV_Rojo
        
 Case 462    'ERR_ALREADYREGISTRED ****
   MV_Pone_Mensaje False, _
   GV_VENTANAS_Estatus(GV_Sockets(LP_Socket).Ventana), _
   MM_Obtener_Mensaje_Parametros(2, LP_Params, ":") + _
   "/Usted No Puede Registrarse Nuevamente" + _
   GV_EOD + GV_EOD, GV_Rojo
        
 Case 463    'ERR_NOPERMFORHOST
   MV_Pone_Mensaje False, _
  GV_VENTANAS_Estatus(GV_Sockets(LP_Socket).Ventana), _
   LP_Params + _
  "/Demasiados Canales" + GV_EOD + GV_EOD, GV_Rojo
                 
 Case 464    'ERR_PASSWDMISMATCH
   MV_Pone_Mensaje False, _
   GV_VENTANAS_Estatus(GV_Sockets(LP_Socket).Ventana), _
   LP_Params + _
   "/Password erroneo" + GV_EOD + GV_EOD, GV_Rojo
                 
 Case 465    'ERR_YOUREBANNEDCREEP
   MV_Pone_Mensaje False, _
   GV_VENTANAS_Estatus(GV_Sockets(LP_Socket).Ventana), _
   LP_Params + _
   "/Usted esta fuera ..." + GV_EOD + GV_EOD, GV_Rojo
                 
 Case 467    'ERR_KEYSET
   MV_Pone_Mensaje False, _
   GV_VENTANAS_Estatus(GV_Sockets(LP_Socket).Ventana), _
   LP_Params + _
   "/Llave establecida" + GV_EOD + GV_EOD, GV_Rojo
                 
 Case 471    'ERR_CHANNELISFULL ****
   MV_Pone_Mensaje False, _
   GV_VENTANAS_Estatus(GV_Sockets(LP_Socket).Ventana), _
   "Channel Is Full/Canal Se Encuentra Lleno" + _
   GV_EOD + GV_EOD, GV_Rojo
    
 Case 472    'ERR_UNKNOWNMODE ****
   MV_Pone_Mensaje False, _
   GV_VENTANAS_Estatus(GV_Sockets(LP_Socket).Ventana), _
   MM_Obtener_Mensaje_Parametros(2, LP_Params, ":") + _
   "/Modo Desconocido" + GV_EOD + GV_EOD, GV_Rojo
        
 Case 473    'ERR_INVITEONLYCHAN ****
   MV_Pone_Mensaje False, _
   GV_VENTANAS_Estatus(GV_Sockets(LP_Socket).Ventana), _
   MM_Obtener_Mensaje_Parametros(2, LP_Params, ":") + _
   "/No Puede Ingresar A Canal(Es Solamente Por Invitación)." _
   + GV_EOD + GV_EOD, GV_Rojo
        
 Case 474    'ERR_BANNEDFROMCHAN ****
   MV_Pone_Mensaje False, _
   GV_VENTANAS_Estatus(GV_Sockets(LP_Socket).Ventana), _
   MM_Obtener_Mensaje_Parametros(2, LP_Params, ":") + _
   "/No Puede Ingresar A Canal(Se Encuentra Restringido)." + _
   GV_EOD + GV_EOD, GV_Rojo
        
 Case 475    'ERR_BADCHANNELKEY
   MV_Pone_Mensaje False, _
   GV_VENTANAS_Estatus(GV_Sockets(LP_Socket).Ventana), _
   LP_Params + _
   "/Llave de Canal erronea" + GV_EOD + GV_EOD, GV_Rojo
   
                 
 Case 481    'ERR_NOPRIVILEGES ****
   MV_Pone_Mensaje False, _
   GV_VENTANAS_Estatus(GV_Sockets(LP_Socket).Ventana), _
   MM_Obtener_Mensaje_Parametros(2, LP_Params, ":") + _
   "/Permiso Denegado- Usted No Es Un Operador IRC" + _
   GV_EOD + GV_EOD, GV_Rojo
           
 Case 482    'ERR_CHANOPRIVSNEEDED ****
   MV_Pone_Mensaje False, _
   GV_VENTANAS_Estatus(GV_Sockets(LP_Socket).Ventana), _
   MM_Obtener_Mensaje_Parametros(2, LP_Params, " ") + _
   "/Usted No Es Operador Del Canal" + GV_EOD + GV_EOD, GV_Rojo
   
 Case 483    'ERR_CANTKILLSERVER
   MV_Pone_Mensaje False, _
   GV_VENTANAS_Estatus(GV_Sockets(LP_Socket).Ventana), _
   LP_Params + _
   "/No puede matar servidor" + GV_EOD + GV_EOD, GV_Rojo
    
 Case 491    ' ERR_NOOPERHOST
   MV_Pone_Mensaje False, _
   GV_VENTANAS_Estatus(GV_Sockets(LP_Socket).Ventana), _
   LP_Params + _
   "/No existe operador de host" + GV_EOD + GV_EOD, GV_Rojo
                 
 Case 501    'ERR_UMODEUNKNOWNFLAG
   MV_Pone_Mensaje False, _
   GV_VENTANAS_Estatus(GV_Sockets(LP_Socket).Ventana), _
   LP_Params + _
   GV_EOD + GV_EOD, GV_Rojo
                 
 Case 502    'ERR_USERSDONTMATCH
   MV_Pone_Mensaje False, _
   GV_VENTANAS_Estatus(GV_Sockets(LP_Socket).Ventana), _
   LP_Params + _
  "/Usuario no cumple criterio" + GV_EOD + GV_EOD, GV_Rojo
 
 
 Case 999 '****
   MV_Pone_Mensaje True, _
   GV_VENTANAS_Estatus(GV_Sockets(LP_Socket).Ventana), _
   MM_Obtener_Mensaje_Parametros(2, LP_Params, ":") + _
   "/Error Comando No Encontrado" + GV_EOD + GV_EOD, GV_Rojo
        
 Case Else
   MV_Pone_Mensaje True, _
   GV_VENTANAS_Estatus(GV_Sockets(LP_Socket).Ventana), _
   LP_Prefijo + " " + LP_Comando + " " + LP_Params + _
   GV_EOD + GV_EOD, GV_Azul
            
  End Select
Exit Sub
Etiqueta_Error:
ME_Muestra_Error

End Sub

Sub MM_Borrar_Retransmisiones(LP_Socket%)
' /***************************************************/
' Toma los registros asociados a LP_Socket en la
' estructura de retransmisiones y los
' marca como borrados.
' /***************************************************/
Dim L_i%

For L_i = 1 To UBound(GV_Estructura_Retransmision)
  If GV_Estructura_Retransmision(L_i).E_Borrado = False _
   And GV_Estructura_Retransmision(L_i).E_Servidor_Emisor = _
     LP_Socket Then
       GV_Estructura_Retransmision(L_i).E_Borrado = True
  End If
  If GV_Estructura_Retransmision(L_i).E_Borrado = False _
   And GV_Estructura_Retransmision(L_i).E_Servidor_Receptor = _
   LP_Socket Then
       GV_Estructura_Retransmision(L_i).E_Borrado = True
  End If
  DoEvents
Next L_i
End Sub

Sub MM_Procesar_Retransmision(Lp_Quien$, LP_Prefijo$, LP_Comando$, LP_Params$, LP_Socket%)
' /***************************************************/
' Procedimiento que se encarga de retransmitir los
' mensajes que nos llegan de los
' respectivos emisores.
' Lp_Quien: Usuario de quien proviene el mensaje.
' LP_Prefijo,LP_Comando, LP_Params: mensaje separado
' en tres componentes (Prefijo,
' comando, parámetros.  Ver gramática).
' LP_Socket : Indice que representa el servidor del
' cual provino el mensaje.
' /***************************************************/

On Error GoTo Etiqueta_Error:
Dim L_i%
Dim L_Mensaje$


For L_i = 1 To UBound(GV_Estructura_Retransmision)
  If GV_Estructura_Retransmision(L_i).E_Borrado = False And _
   GV_Estructura_Retransmision(L_i).E_Servidor_Emisor = _
   LP_Socket Then
     If UCase( _
      GV_Estructura_Retransmision(L_i).E_Alias_Emisor) = _
        UCase(Lp_Quien) Then
          If GV_Estructura_Retransmision(L_i).E_Omitir_Mensaje _
           Then
            L_Mensaje = _
            GV_Estructura_Retransmision(L_i).E_Prefijo + _
            GV_Estructura_Retransmision(L_i).E_Sufijo
         Else
           If GV_Estructura_Retransmision(L_i).E_SoloMensaje Then
            L_Mensaje = _
            GV_Estructura_Retransmision(L_i).E_Prefijo + _
            LP_Params + _
            GV_Estructura_Retransmision(L_i).E_Sufijo
          Else
            L_Mensaje = _
            GV_Estructura_Retransmision(L_i).E_Prefijo + _
            LP_Prefijo + LP_Comando + LP_Params + _
            GV_Estructura_Retransmision(L_i).E_Sufijo
         End If
       End If
       L_Mensaje = Trim(L_Mensaje)
       MV_Pone_Mensaje True, _
       GV_VENTANAS_Estatus(GV_Sockets(LP_Socket).Ventana), _
       GV_EOD + "~~~~~~~~~~~~~~~~~~~~~~~~~" + GV_EOD, GV_Negro
       MV_Pone_Mensaje True, _
       GV_VENTANAS_Estatus(GV_Sockets(LP_Socket).Ventana), _
       "*** Iniciando Retranmisión:" + _
       " Alias Emisor :", GV_RojoAlgo
       MV_Pone_Mensaje True, _
       GV_VENTANAS_Estatus(GV_Sockets(LP_Socket).Ventana), _
       " < " + Lp_Quien + ">" + GV_EOD, GV_Azul
            
       MV_Pone_Mensaje True, _
       GV_VENTANAS_Estatus(GV_Sockets(LP_Socket).Ventana), _
       "       Mensaje =>  " + _
       L_Mensaje + GV_EOD, GV_RojoAlgo
       
       MV_Pone_Mensaje True, _
       GV_VENTANAS_Estatus(GV_Sockets(LP_Socket).Ventana), _
       "*** Servidor Destino : ", GV_RojoAlgo
       MV_Pone_Mensaje True, _
       GV_VENTANAS_Estatus(GV_Sockets(LP_Socket).Ventana), _
         "< " + GV_Sockets( _
         GV_Estructura_Retransmision( _
         L_i).E_Servidor_Receptor).Direcc + " >" + _
         GV_EOD + GV_EOD + GV_EOD, GV_Azul
       
       'Enviar mensaje al receptor especificado en la
       ' retransmisión.
       
       MM_Enviar_Mensaje L_Mensaje, _
       GV_Estructura_Retransmision(L_i).E_Servidor_Receptor
    End If
       
  
  End If

Next L_i

Exit Sub
Etiqueta_Error:
ME_Muestra_Error

End Sub

Sub MM_Entrega_Mensaje(LP_Socket%, LP_Prefijo$, LP_Comando$, LP_Params$, LP_Status%)
' /*******************************************************/
' Este procedimiento se encarga de procesar los comando
' provenientes de un servidor y enviarlos al lugar
' correspondiente de la aplicación.
' LP_Socket: Indice que representa el servidor del cual
' provino el mensaje.
' LP_Prefijo:Prefijo del mensaje(Ver gramática de mensajes).
' LP_Comando:Comando del mensaje(Ver gramática de mensajes).
' LP_Params:Parámetros del mensaje(Ver gramática de mensajes).
' /********************************************************/

On Error GoTo Etiqueta_Error:

Dim L_vent%, L_i%, L_j%
Dim L_Usuario As ES_USUARIO
Dim L_Datos(10) As String, L_Botado$, L_Lista_Parametros() As String

'l_datos(1) = canal l_datos(2) = nick  L_Datos(3)=mensaje

If LP_Status <> 0 Then
' Codigos de error y de replys
 MM_Procesa_Comandos_Numericos LP_Socket, _
 LP_Prefijo, LP_Comando, LP_Params, LP_Status
 
Else
 Select Case UCase(Trim(LP_Comando))
       
  Case "JOIN"
   ' En caso que alguien halla entrado a un canal _ ****
   ' en el que se esta, o el usuario acaba de unirse a
   ' un canal
   L_Datos(1) = MM_Obtener_Mensaje_Parametros(2, LP_Params, ":")
   L_Datos(2) = MM_Obtener_Mensaje_Parametros(1, LP_Prefijo, "!")
   L_vent = MV_Busca_Ventana_Canal(L_Datos(1), LP_Socket)
   If MM_EsMiNick(L_Datos(2), LP_Socket) Then
    If L_vent = 0 Then
      MV_CreaVentana_Canal LP_Socket, L_Datos(1)
      GV_Sockets(LP_Socket).JOIN_PENDIENTE = True
      MV_Pone_Mensaje True, _
      GV_VENTANAS_Estatus(GV_Sockets(LP_Socket).Ventana), _
      "Joining Channel/Ingresando a Canal ", GV_Azul
      MV_Pone_Mensaje True, _
      GV_VENTANAS_Estatus(GV_Sockets(LP_Socket).Ventana), _
      L_Datos(1) + GV_EOD, GV_Verde
    End If
   Else
    If L_vent <> 0 Then
     ReDim Preserve L_Lista_Parametros(2)
     L_Lista_Parametros(1) = L_Datos(2)
     MM_Cargar_Usuarios_A_Canal L_Lista_Parametros(), _
     GV_VENTANAS_Canal(L_vent)
     MV_Pone_Mensaje True, _
     GV_VENTANAS_Canal(L_vent), "Usuario <" + L_Datos(2) _
     + "> have entered a Channel/Ha Ingresado a Canal " + _
     GV_EOD + GV_EOD, GV_Verde
    End If
   End If
        
  Case "PART"
   ' Aparte del Part se debe revisar cuando a un usuario le es
   ' Cerrada la conexión, para que tambien sea removido de la
   ' lista de canales
   L_Datos(1) = MM_Obtener_Mensaje_Parametros(2, LP_Params, ":")
   L_Datos(2) = MM_Obtener_Mensaje_Parametros(1, LP_Prefijo, "!")
   L_vent = MV_Busca_Ventana_Canal(L_Datos(1), LP_Socket)
    If L_vent <> 0 Then
      If MM_EsMiNick(L_Datos(2), LP_Socket) Then
         GV_VENTANAS_Canal(L_vent).OT_SocketAsociado = 0
         If L_vent <> 0 Then Unload GV_VENTANAS_Canal(L_vent)
         Else
           MM_Descargar_Usuarios_De_Canal L_Datos(2), _
           GV_VENTANAS_Canal(L_vent)
         End If
      End If
        
  Case "PRIVMSG"
   'Mensajes privados provenientes de usuarios o canales
   L_Datos(1) = MM_Obtener_Mensaje_Parametros(1, LP_Params, ":")
   L_Datos(3) = MM_Obtener_Mensaje_Parametros(2, LP_Params, ":")
   L_Datos(2) = MM_Obtener_Mensaje_Parametros(1, LP_Prefijo, "!")
   If Left(L_Datos(1), 1) = "#" Or Left(L_Datos(1), 1) = "&" Then
     L_vent = MV_Busca_Ventana_Canal(L_Datos(1), LP_Socket)
     If L_vent <> 0 Then
       MV_Pone_Mensaje True, _
       GV_VENTANAS_Canal(L_vent), _
       "<" + L_Datos(2) + ">: ", GV_Rojo
       MV_Pone_Mensaje True, _
       GV_VENTANAS_Canal(L_vent), L_Datos(3) + _
       GV_EOD, GV_Negro
     End If
     MM_Procesar_Retransmision L_Datos(2), _
     LP_Prefijo, LP_Comando, LP_Params, _
     LP_Socket ' Procesar la Retransmision
   Else
    'faltan mensajes privados
    'Codigo para recibir archivos
     If Asc(Left(L_Datos(3), 1)) = 1 Then
       ' Hay que verificar que sea un DCC
       If Mid(L_Datos(3), 2, 3) = "DCC" Then
         MM_Recibir_Archivo L_Datos(1), L_Datos(2), L_Datos(3)
       Else
         MV_Pone_Mensaje True, _
         GV_VENTANAS_Estatus(GV_Sockets(LP_Socket).Ventana), _
         "<" + L_Datos(2) + ">: ", _
         GV_Rojo
         MV_Pone_Mensaje True, _
         GV_VENTANAS_Estatus(GV_Sockets(LP_Socket).Ventana), _
         L_Datos(3) + GV_EOD, GV_Negro
       End If
            
    Else
      L_vent = MV_Busca_Ventana_Usuario(L_Datos(2), LP_Socket)
      If L_vent = 0 Then
        L_vent = MV_CreaVentana_Usuario(LP_Socket, L_Datos(2))
        GV_VENTANAS_Usuario(L_vent).Caption = _
         L_Datos(2) + " [ " + _
         GV_Sockets(LP_Socket).Direcc + " ]"
      End If
      Beep
      GV_VENTANAS_Usuario(L_vent).Icon = _
       LoadPicture(App.Path + "\men.ico")
       MV_Pone_Mensaje True, _
       GV_VENTANAS_Usuario(L_vent), _
       "<" + L_Datos(2) + ">: ", GV_Rojo
       MV_Pone_Mensaje True, _
       GV_VENTANAS_Usuario(L_vent), L_Datos(3) + GV_EOD, GV_Negro
       If GV_Explorador Then
         OF_Explorador.OA_Explorador.Nodes.Item( _
         "U" + CStr(GV_VENTANAS_Usuario(L_vent).Tag)).Image = 9
       End If
       ' Procesar la Retransmision
       MM_Procesar_Retransmision L_Datos(2), _
       LP_Prefijo, LP_Comando, LP_Params, LP_Socket
    End If
   End If
        
  Case "PING"
    L_Datos(1) = MM_Obtener_Mensaje_Parametros(2, LP_Params, ":")
    MM_Enviar_Mensaje "Pong " + L_Datos(1) + GV_EOD, LP_Socket
    MV_Pone_Mensaje False, _
    GV_VENTANAS_Estatus(GV_Sockets(LP_Socket).Ventana), _
    "Ping Contestado a :" + L_Datos(1) + GV_EOD, GV_Verde
        
  Case "PONG"
    MV_Pone_Mensaje False, _
    GV_VENTANAS_Estatus(GV_Sockets(LP_Socket).Ventana), _
    "Pong " + LP_Params + GV_EOD, GV_Verde
        
  Case "KICK"
    L_Datos(1) = MM_Obtener_Mensaje_Parametros(1, LP_Params, " ")
    L_Botado = MM_Obtener_Canal(LP_Params, " ", L_Botado)
    L_Datos(2) = MM_Obtener_Mensaje_Parametros(1, LP_Prefijo, "!")
    L_Datos(3) = MM_Obtener_Mensaje_Parametros(2, LP_Params, ":")
    L_vent = MV_Busca_Ventana_Canal(L_Datos(1), LP_Socket)
    If L_vent <> 0 Then
       If MM_EsMiNick(L_Botado, LP_Socket) Then
           GV_VENTANAS_Canal(L_vent).OT_SocketAsociado = 0
           Unload GV_VENTANAS_Canal(L_vent)
           MV_Pone_Mensaje False, _
           GV_VENTANAS_Estatus(GV_Sockets(LP_Socket).Ventana), _
           "Ha Sido Expulsado De " + _
           L_Datos(1) + ".  " + L_Datos(3) + GV_EOD, vbBlack
       Else
          MM_Descargar_Usuarios_De_Canal L_Botado, _
          GV_VENTANAS_Canal(L_vent)
          MV_Pone_Mensaje True, _
          GV_VENTANAS_Canal(L_vent), "<" + L_Botado + _
          ">: Ha Sido Expulsado.  " + L_Datos(3) + GV_EOD, vbBlack
       End If
    End If
        
  Case "NICK"
    L_Datos(2) = MM_Obtener_Mensaje_Parametros(1, LP_Prefijo, "!")
    L_Datos(3) = MM_Obtener_Mensaje_Parametros(2, LP_Params, ":")
    MV_Reemplazar_Alias L_Datos(2), L_Datos(3), LP_Socket
    If MM_EsMiNick(L_Datos(2), LP_Socket) Then
     GV_Sockets(LP_Socket).Nick = Trim(L_Datos(3))
     GV_VENTANAS_Estatus(GV_Sockets(LP_Socket).Ventana).Caption = _
     GV_Sockets(LP_Socket).Direcc + "(" + _
     CStr(GV_Sockets(LP_Socket).Puerto) + _
     ") Estado : [CONECTADO]" + " :¬) = " + _
     GV_Sockets(LP_Socket).Nick
    End If
        
  Case "TOPIC"
    L_Datos(1) = MM_Obtener_Mensaje_Parametros(1, LP_Params, ":")
    L_Datos(3) = MM_Obtener_Mensaje_Parametros(2, LP_Params, ":")
    L_vent = MV_Busca_Ventana_Canal(L_Datos(1), LP_Socket)
    If L_vent <> 0 Then
      GV_VENTANAS_Canal(L_vent).Caption = L_Datos(1) + " [" + _
      GV_Sockets(LP_Socket).Direcc + "] : " + L_Datos(3)
    End If
        
  Case "QUIT"
    ' Cuando un usuario es desconectado de un servidor, se tiene
    ' Que tener cuidado de removerlo de la lista de Canales si se
    ' Encuentra en uno de ellos, las ventanas privadas no serán
    ' revisadas
    L_Datos(1) = MM_Obtener_Mensaje_Parametros(1, LP_Prefijo, "!")
     If Not MM_EsMiNick(L_Datos(1), LP_Socket) Then
       MV_Descarga_Nick_de_Canales L_Datos(1), LP_Socket
       MV_Notifica_Salida L_Datos(1), LP_Socket
     End If
    MV_Pone_Mensaje False, _
    GV_VENTANAS_Estatus(GV_Sockets(LP_Socket).Ventana), _
    "<Usuario :", vbRed
       
    MV_Pone_Mensaje False, _
    GV_VENTANAS_Estatus(GV_Sockets(LP_Socket).Ventana), _
    " " + L_Datos(1), vbBlack
       
    MV_Pone_Mensaje False, _
    GV_VENTANAS_Estatus(GV_Sockets(LP_Socket).Ventana), _
    " >: Ha sido sacado del servidor ..." + GV_EOD, vbRed
    
             
  Case "MODE"  '****
    'quien ejecuto el comando
    L_i = InStr(1, LP_Params, "!")
    If L_i <> 0 Then
      L_Datos(1) = _
      MM_Obtener_Mensaje_Parametros(1, LP_Prefijo, "!")
    Else
       L_Datos(1) = Trim(LP_Prefijo)
    End If
    'canal o nick
    L_Datos(2) = MM_Obtener_Mensaje_Parametros(1, LP_Params, " ")
    
    'switches
    ' Temporal por Rogger
    L_Datos(3) = MM_Obtener_Mensaje_Parametros(4, LP_Params, " ")
    'L_Datos(3) = MM_Obtener_Mensaje_Parametros(4, LP_Params, ":")
    'Lista de usuarios
    L_Datos(5) = _
     MM_Obtener_Mensaje_Parametros(2, LP_Params, L_Datos(3))
      
    'Lista de usuarios
    L_Datos(5) = _
     MM_Obtener_Mensaje_Parametros(3, L_Datos(5), " ")
     
    If Left(L_Datos(2), 1) = "#" Or Left(L_Datos(2), 1) = "&" _
     Then
      L_vent = MV_Busca_Ventana_Canal(L_Datos(2), LP_Socket)
      L_Datos(4) = "+"
      If Left(L_Datos(3), 1) = "-" Or _
       Left(L_Datos(3), 1) = "+" Then
         L_Datos(4) = Left(L_Datos(3), 1)
         L_Datos(3) = Right(L_Datos(3), Len(L_Datos(3)) - 1)
      End If
            
      While InStr(L_Datos(3), "o") > 0
        L_Datos(3) = _
         Right(L_Datos(3), _
         Len(L_Datos(3)) - InStr(L_Datos(3), "o"))
        L_i = L_i + 1
      Wend
            
      ReDim Preserve L_Lista_Parametros(1)
      MM_Obtener_Lista_Parametros L_Datos(5), _
      L_Lista_Parametros()
                        
      For L_j = 1 To L_i
        MM_Indicar_Operador L_Lista_Parametros(L_j), _
        GV_VENTANAS_Canal(L_vent), L_Datos(4)
        If L_Datos(4) = "+" Then
          MV_Pone_Mensaje True, _
          GV_VENTANAS_Estatus(GV_Sockets(LP_Socket).Ventana), _
          L_Lista_Parametros(L_j) + _
          " Has Become An Operator Of Channel:/ " + _
          " Se Ha Convertido En Operador Del Canal: " + _
          L_Datos(2) + GV_EOD, GV_Morado
        Else
          MV_Pone_Mensaje True, _
          GV_VENTANAS_Estatus(GV_Sockets(LP_Socket).Ventana), _
          L_Lista_Parametros(L_j) + _
          " Has Lost Operator rights In Channel:/Ha Perdido " + _
          "Derechos De Operador Del Canal: " + L_Datos(2) + _
          GV_EOD, GV_Morado
        End If
     Next
            
   Else
     MV_Pone_Mensaje True, _
     GV_VENTANAS_Estatus(GV_Sockets(LP_Socket).Ventana), _
     LP_Comando + "/Modo: " + LP_Params + GV_EOD, GV_Morado
               
   End If
    
  Case "INVITE" ' ****
    MV_Pone_Mensaje True, _
    GV_VENTANAS_Estatus(GV_Sockets(LP_Socket).Ventana), _
    "You Have been Invited To Channel/ " + _
    " Ha Sido Invitación A Canal: ", GV_Morado
    
    MV_Pone_Mensaje True, _
    GV_VENTANAS_Estatus(GV_Sockets(LP_Socket).Ventana), _
    MM_Obtener_Mensaje_Parametros(2, LP_Params, ":") + _
    GV_EOD, GV_Verde
    
  Case "ERROR" ' ****
    MV_Pone_Mensaje True, _
    GV_VENTANAS_Estatus(GV_Sockets(LP_Socket).Ventana), _
    LP_Params + GV_EOD, GV_Rojo
  
  Case Else
    MV_Pone_Mensaje True, _
    GV_VENTANAS_Estatus(GV_Sockets(LP_Socket).Ventana), _
    LP_Prefijo + LP_Comando + LP_Params + GV_EOD, GV_Azul
  
  End Select
  
End If
Exit Sub

Etiqueta_Error:

ME_Muestra_Error

End Sub
Sub MM_Obtener_Lista_Parametros(ByVal LP_Params$, LP_Lista_Parametros() As String)
' /*******************************************************/
' Toma un string que viene en LP_Params y lo secciona en
' sub-strings, tomando como delimitador el espacio en
' blanco. Luego retorna los sub-strings en un arreglo
' llamando LP_Lista_Parametros.
' /*******************************************************/

On Error GoTo Etiqueta_Error:

Dim L_Ultimo%
Dim L_Posicion&, L_No_Parametro&, L_Conteo_Arreglo&

L_Conteo_Arreglo = UBound(LP_Lista_Parametros)
L_No_Parametro = 1

LP_Params = Trim(LP_Params)
While Len(LP_Params) > 0
    L_Posicion = InStr(LP_Params, " ")
    ReDim Preserve LP_Lista_Parametros(L_Conteo_Arreglo + 1)
    L_Conteo_Arreglo = L_Conteo_Arreglo + 1
    If L_Posicion = 0 Then
      L_Posicion = Len(LP_Params): L_Ultimo = 1
    End If
    LP_Lista_Parametros(L_No_Parametro) = _
    Left(LP_Params, L_Posicion + L_Ultimo - 1)
    LP_Params = _
    Trim(Right(LP_Params, Len(LP_Params) - L_Posicion))
    L_No_Parametro = L_No_Parametro + 1
Wend
Exit Sub
Etiqueta_Error:
ME_Muestra_Error

End Sub


