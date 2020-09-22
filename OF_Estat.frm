VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.1#0"; "comctl32.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.1#0"; "richtx32.ocx"
Begin VB.Form OF_Estatus 
   AutoRedraw      =   -1  'True
   ClientHeight    =   5655
   ClientLeft      =   1140
   ClientTop       =   1230
   ClientWidth     =   7995
   HelpContextID   =   6
   Icon            =   "OF_Estat.frx":0000
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5655
   ScaleWidth      =   7995
   Begin ComctlLib.Toolbar OTB_Estatus 
      Align           =   1  'Align Top
      Height          =   450
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   7995
      _ExtentX        =   14102
      _ExtentY        =   794
      ButtonWidth     =   714
      ButtonHeight    =   688
      AllowCustomize  =   0   'False
      HelpContextID   =   6
      ImageList       =   "ImageList1"
      _Version        =   327680
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   5
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Enviar Archivo"
            Object.Tag             =   ""
            ImageIndex      =   5
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Retransmisión de Mensajes"
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Canales Favoritos"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Desconectarse de Servidor"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Conectarse a Servidor"
            Object.Tag             =   ""
            ImageIndex      =   7
         EndProperty
      EndProperty
   End
   Begin VB.TextBox OT_Tipo 
      Height          =   285
      Left            =   2400
      TabIndex        =   8
      Top             =   5340
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.TextBox OT_Puerto 
      Height          =   285
      Left            =   1365
      TabIndex        =   7
      Top             =   5355
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.TextBox OT_Direccion 
      Height          =   285
      Left            =   195
      TabIndex        =   6
      Top             =   5340
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.TextBox OT_Ventana_Estatus 
      Height          =   288
      Left            =   7572
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   5256
      Visible         =   0   'False
      Width           =   768
   End
   Begin VB.PictureBox OT_Asynccontrol 
      Height          =   216
      Left            =   6516
      ScaleHeight     =   150
      ScaleWidth      =   930
      TabIndex        =   2
      Top             =   5265
      Visible         =   0   'False
      Width           =   996
   End
   Begin VB.TextBox OT_SocketAsociado 
      Height          =   285
      Left            =   4212
      TabIndex        =   1
      Text            =   "2"
      Top             =   5304
      Visible         =   0   'False
      Width           =   2070
   End
   Begin VB.TextBox OT_Comando 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      HelpContextID   =   6
      Left            =   165
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   4680
      Width           =   7845
   End
   Begin RichTextLib.RichTextBox OL_Estatus 
      Height          =   4080
      HelpContextID   =   6
      Left            =   120
      TabIndex        =   3
      Top             =   450
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   7197
      _Version        =   327680
      HideSelection   =   0   'False
      ScrollBars      =   2
      DisableNoScroll =   -1  'True
      RightMargin     =   5000
      TextRTF         =   $"OF_Estat.frx":0442
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   7512
      Top             =   384
      _ExtentX        =   794
      _ExtentY        =   794
      BackColor       =   -2147483643
      ImageWidth      =   20
      ImageHeight     =   20
      _Version        =   327680
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   7
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "OF_Estat.frx":050B
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "OF_Estat.frx":0825
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "OF_Estat.frx":0B3F
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "OF_Estat.frx":0E59
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "OF_Estat.frx":1173
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "OF_Estat.frx":148D
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "OF_Estat.frx":17A7
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "OF_Estatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' Variables locales a la forma
' LF_Historial es un arreglo de strings que lleva un
' historial de los ultimos 20 mensajes digitados por
' el usuario
Dim LF_Historial(20) As String
' LF_Ultimo lleva el indice del ultimo indice utilizado
' en el Historial Y LF_actual lleva la posición en que
' se encuentra en el historial si un usuario esta
' navegando  en el
Dim LF_Ultimo%, LF_Actual%

Private Sub Form_Load()
'/********************************************************/
' Al cargarse la forma se inicializan las variables
' utilizadas en el historial, el textbox de mensajes es
' blanqueado
'/********************************************************/
LF_Ultimo = 0
LF_Actual = 0

OL_Estatus.Text = ""

End Sub

Private Sub Form_Resize()
'/********************************************************/
' Cuando se cambia el tamaño de la forma entonces ajustamos
' los controles  de la forma a otros tamaños de acuerdo al
' nuevo tamaño de la forma
'/********************************************************/
On Error GoTo Etiqueta_Error:
If ScaleWidth <> 0 And ScaleHeight <> 0 And _
  ScaleHeight > 500 And ScaleWidth > 500 Then
    OL_Estatus.Move 0, 450, ScaleWidth, ScaleHeight - 1050
    OT_Comando.Move 0, ScaleHeight - 600, ScaleWidth, 600
    OL_Estatus.SelStart = Len(OL_Estatus.Text)
    OL_Estatus.RightMargin = ScaleWidth - 500
End If

Exit Sub
Etiqueta_Error:
ME_Muestra_Error
End Sub

Private Sub Form_Unload(Cancel As Integer)
'/*********************************************************/
' Al descargar una forma de servidores se tiene que
' realizar primero la  verificación de que si la forma
' todavía tiene activa su conexión. Esto se hace
' verificando el TextBox OT_SocketAsociado el cual
' permanece invisible  al usuario, si el textbox tiene
' un valor diferente de cero, significa que
' la conexión todavía se encuentra activa. Seguidamente
' se llama a la función de cerrar conexión enviando
' de parametro el socket asociado
' que representaba la conexión que se acaba de cerrar.
' Seguidamente cierra todas las ventanas asociadas a el.
' Si el explorador se encuentra activo entonces
' Procede a eliminarlo del arbol del explorador
'/*******************************************************/
On Error GoTo Etiqueta_Error:

Dim L_Status As Integer


If OT_SocketAsociado <> 0 Then
      MM_Cerrar_Conexion OT_SocketAsociado
End If
If Trim(Me.Tag) = "" Then Exit Sub
 ' Setea la ventana como borrada
MV_Cerrar_El_Resto_De_Ventanas Me.Tag

GV_Estado_Estatus(Me.Tag).Deleted = True
If GV_Explorador Then ' Si esta activo el explorador
  OF_Explorador.OA_Explorador.Nodes.Remove ("S" + CStr(Me.Tag))
End If
Exit Sub
Etiqueta_Error:
ME_Muestra_Error

End Sub

Private Sub OL_Estatus_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'/*****************************************************/
' Si se hace click con el botón derecho sobre el Textbox
' donde se reciben mensajes y se ha seleccionado una
' porción de texto se procede  ha habilitar el menu que
' presenta la opción de abrir una pagina WEB
'/*****************************************************/
If Button = 2 Then
  If Trim(OL_Estatus.SelText) = "" Then Exit Sub
  GV_Nombre = OL_Estatus.SelText
  PopupMenu OF_principal.M_ficticio
  
End If
End Sub

Private Sub OT_Asynccontrol_Resize()
' /******************************************************/
' Este control es utilizado para recibir los eventos
' asíncronos de Winsock  para el socket Asociado
' (OT_SocketAsociado), este control es invisible
' al usuario, y los eventos son recibidos en el evento
' RESIZE del control
'
' Recepción de datos de las ventanas de estatus
'El manejador para el read de IRC
' /******************************************************/
On Error GoTo Etiqueta_Error:
Dim L_Msg$
Dim L_Texto$
Dim L_Largo$
Dim L_Status As Integer
Dim L_prefijo$, L_Comando$, L_params$
Dim L_Tam As Long
    
    ' Variable en donde se recuperan los datos que
    ' Provienen del servidor por el socket especificado
    
                        
    
    
    L_Largo = ""
    DoEvents

    ' Función de Winsock que nos indica cuantos bytes se han
    ' llegado  a el socket
    
    L_Status = ioctlsocket(GV_Sockets(Me.OT_SocketAsociado).socket, _
                        FIONREAD, L_Tam)
                        
    L_Msg = Space$(L_Tam)
    ' Función de Winsock que lee datos del Buffer
    L_Status = recv(GV_Sockets(Me.OT_SocketAsociado).socket, _
               L_Msg, Len(L_Msg), 0) = SOCKET_ERROR
               
    If L_Status = SOCKET_ERROR Then
          L_Status = WSAGetLastError()
          ' si ocurrió un error, lo mas seguro es que
          ' hemos sido desconectados
          ' del servidor
          
          Beep

          MV_Pone_Mensaje True, Me, _
           "Connection Finished/La conexión ha " + _
           "sido terminada" + GV_EOD, vbRed
           
          MV_Setear_Socket OT_SocketAsociado
          ' Borrar todas las entradas en la estructura de
          ' Retransmión del socket asociado
          MM_Borrar_Retransmisiones OT_SocketAsociado
          ' Ir a todas las ventanas que tienen este socket
          ' asociado y setearlos en cero
          ' Luego cerrar el socket
          L_Status = closesocket( _
                     GV_Sockets(OT_SocketAsociado).socket)
          
          Me.Caption = GV_Sockets( _
              OT_SocketAsociado).Direcc + _
              "  Estado :[No Conectado] " + _
              " :¬( = " + GV_Sockets(OT_SocketAsociado).Nick
              
          GV_Sockets(OT_SocketAsociado).socket = _
            INVALID_SOCKET
         ' Setear como libre el socket
          GV_Sockets(OT_SocketAsociado).Libre = True
        ' Ya no tiene socket asociado la ventana
          Me.OT_SocketAsociado = 0
           ' Cambiar el Icono al servidor en el explorador
           If GV_Explorador Then
              OF_Explorador.OA_Explorador.Nodes.Item( _
              "S" + CStr(Me.Tag)).Image = 5
              OF_Explorador.OA_Explorador.Nodes.Item( _
              "S" + CStr(Me.Tag)).Text = Me.Caption
              
           End If
      Else
         ' Envie el mensaje al procesador de mensajes para que
         ' este lo distribuya  a su respectivo destino
         MV_Despacha_mensaje Trim(L_Msg), _
         GV_Sockets(OT_SocketAsociado).Ventana, OT_SocketAsociado
    End If
    
Exit Sub
Etiqueta_Error:
ME_Muestra_Error

End Sub

Private Sub OT_Comando_GotFocus()
OL_Estatus.SelStart = Len(OL_Estatus.Text)
End Sub

Private Sub OT_comando_KeyDown(KeyCode As Integer, Shift As Integer)
'/********************************************************/
' Si se presionan las teclas Ctrl-Up o Ctrl-down, entonces
' se recorre el historial de comandos
'/********************************************************/
On Error GoTo Etiqueta_Error:
If LF_Actual = 0 Then Exit Sub

If Shift = 2 Then ' Ctrl key
   Select Case KeyCode
    
    Case 38 ' Up key
       If LF_Ultimo <> 1 Then
          LF_Ultimo = LF_Ultimo - 1
          OT_Comando = LF_Historial(LF_Ultimo)
    End If
    
    
    Case 40 ' Downkey
       
       If LF_Ultimo <> 20 And LF_Ultimo <= LF_Actual Then
           LF_Ultimo = LF_Ultimo + 1
           OT_Comando = LF_Historial(LF_Ultimo)
       End If
       

   End Select
      
End If

Exit Sub
Etiqueta_Error:
ME_Muestra_Error

End Sub

Private Sub OT_comando_KeyPress(KeyAscii As Integer)
'/*******************************************************/
' Cuando se presiona <ENTER> en el textbox de comandos se
' procede a verificar si la ventana todavia posee una
' conexión activo(OT_SocketAsociado<>0).
' Si la conexión todavía se encuentra activa, se procede
' a procesar el  mensaje digitado.
' Una vez procesado(haber buscado su equivalente en la
' tabla de comandos) se procede a enviarlo por
' el socket Asociado. A la vez registra el mensaje en el
' historial de mensajes  de la ventana.
'/*******************************************************/
On Error GoTo Etiqueta_Error:
Dim L_res%
Dim L_Texto$

DoEvents

If KeyAscii = 13 And Len(OT_Comando) > 0 Then
    If Me.OT_SocketAsociado <> 0 Then
        L_res = InStr(1, Trim(OT_Comando), " ")
        If L_res = 0 Then
            L_Texto = Trim(OT_Comando)
        Else
            L_Texto = _
            MM_Obtener_Mensaje_Parametros(1, OT_Comando, " ")
        End If
        L_Texto = MD_Obtener_Comando(L_Texto)
        If L_res <> 0 Then
          OT_Comando = L_Texto + " " + _
          MM_Obtener_Mensaje_Parametros(2, OT_Comando, " ")
        Else
          OT_Comando = L_Texto
        End If
        L_Texto = OT_Comando & GV_EOD
        L_res = send(GV_Sockets(Me.OT_SocketAsociado).socket, _
                L_Texto, Len(L_Texto), 0)
        
        If L_res = SOCKET_ERROR Then
            L_res = WSAGetLastError()
            MV_Pone_Mensaje False, Me, _
            ME_WsockError(L_res), vbRed
         End If
        
        If LF_Actual <> 20 Then
            LF_Ultimo = LF_Actual + 1
            LF_Historial(LF_Ultimo) = OT_Comando
            LF_Actual = LF_Ultimo
            LF_Ultimo = LF_Ultimo + 1
        Else
           MG_Borra_Primero LF_Historial
           LF_Historial(20) = OT_Comando
        End If
        OT_Comando = ""
    Else
        MV_Pone_Mensaje True, Me, _
        "Usted no se encuentra conectado al servidor ..." _
        & GV_EOD, vbRed

      
    End If
    OT_Comando = ""
    KeyAscii = 8
ElseIf KeyAscii = 13 Then KeyAscii = 8
End If

Exit Sub
Etiqueta_Error:
ME_Muestra_Error
End Sub

Private Sub OT_Comando_KeyUp(KeyCode As Integer, Shift As Integer)
'/***************************************************/
' Procedimiento que nos permite movernos entre las
' ventanas de estatus  si se presionan las teclas de
' Shift-Up o Shift_Down.
'/****************************************************/
On Error GoTo Etiqueta_Error:
Dim L_cual%
If Shift = 1 Then ' shift key
   Select Case KeyCode
    
        Case 38 ' Up key
          L_cual = MV_BuscaSigAnt_Ventana_Estatus( _
          Me.Tag, "ANTERIOR", False)
                      
          If GV_VENTANAS_Estatus(L_cual).Visible Then
           GV_VENTANAS_Estatus(L_cual).SetFocus
           If GV_VENTANAS_Estatus(L_cual).WindowState = 1 _
            Then GV_VENTANAS_Estatus(L_cual).WindowState = 0
           End If
        
        Case 40 ' Downkey
          L_cual = MV_BuscaSigAnt_Ventana_Estatus( _
          Me.Tag, "SIGUIENTE", False)
          
          If GV_VENTANAS_Estatus(L_cual).Visible Then
            GV_VENTANAS_Estatus(L_cual).SetFocus
            If GV_VENTANAS_Estatus(L_cual).WindowState = 1 _
            Then GV_VENTANAS_Estatus(L_cual).WindowState = 0
          End If
  End Select
End If

Exit Sub
Etiqueta_Error:
ME_Muestra_Error

End Sub

Private Sub OTB_Estatus_ButtonClick(ByVal Button As Button)
'/********************************************************/
' Este es el código utilizado para determinar que opción
' del toolbar de la  ventana fue seleccionada

'/********************************************************/
On Error GoTo Etiqueta_Error:
Dim L_cual%
Dim L_Ventana As New OF_Enviar
Dim L_Ventana1 As New OF_Retransmision
Select Case Button.Index
Case 1 ' Opción de Enviar Archivos
     Load L_Ventana
    If Trim(OL_Estatus.SelText) <> "" Then
      L_Ventana.OT_Nick = OL_Estatus.SelText
    End If
    L_Ventana.Show
   
Case 2 ' Opción de Retransmisión de MEnsajes
   
    Load L_Ventana1
    L_Ventana1.Show
    
Case 3 ' Opción de Canales Preferidos
    GV_Seleccion = Me.OT_SocketAsociado
    Load OF_Canales_Preferidos
    OF_Canales_Preferidos.Show
   
Case 4 ' Opción para cerrar la ventana
       ' (Deconectarse del Servidor)
   If Me.OT_SocketAsociado <> 0 Then
     MM_Cerrar_Conexion OT_SocketAsociado
   End If
Case 5 ' Opción para reconectarse al servidor
       
   If Me.OT_SocketAsociado = 0 Then
     MV_Cerrar_El_Resto_De_Ventanas Me.Tag
     Me.OL_Estatus = ""
     MM_Connect Trim(OT_Direccion), CInt(OT_Puerto), _
     1, CInt(Me.Tag)
     
   End If
End Select

Exit Sub
Etiqueta_Error:
ME_Muestra_Error
End Sub



Private Sub OTB_Estatus_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then
   DoEvents
   PopupMenu OF_principal.M_Ventanas
End If
End Sub
