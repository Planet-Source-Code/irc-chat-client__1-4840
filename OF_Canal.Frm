VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.1#0"; "comctl32.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.1#0"; "richtx32.ocx"
Begin VB.Form OF_Hablar_Canal 
   AutoRedraw      =   -1  'True
   Caption         =   "Hablar Con Canal"
   ClientHeight    =   4860
   ClientLeft      =   765
   ClientTop       =   840
   ClientWidth     =   8625
   HelpContextID   =   27
   Icon            =   "OF_Canal.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4860
   ScaleWidth      =   8625
   Begin ComctlLib.Toolbar OTB_Estatus 
      Align           =   1  'Align Top
      Height          =   450
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   8625
      _ExtentX        =   15214
      _ExtentY        =   794
      ButtonWidth     =   714
      ButtonHeight    =   688
      AllowCustomize  =   0   'False
      HelpContextID   =   27
      ImageList       =   "ImageList1"
      _Version        =   327680
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   5
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Enviar Archivo"
            Object.Tag             =   ""
            ImageIndex      =   5
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Retransmisión de Mensajes"
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Canales Favoritos"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Notificación de Mensajes"
            Object.Tag             =   ""
            ImageIndex      =   7
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Salir del Canal"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
      EndProperty
      MouseIcon       =   "OF_Canal.frx":0442
   End
   Begin VB.TextBox OT_Ventana_Estatus 
      Height          =   288
      Left            =   5832
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   4776
      Visible         =   0   'False
      Width           =   768
   End
   Begin VB.ListBox OL_Usuarios 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   3180
      HelpContextID   =   27
      ItemData        =   "OF_Canal.frx":045E
      Left            =   120
      List            =   "OF_Canal.frx":0460
      Sorted          =   -1  'True
      TabIndex        =   3
      Top             =   585
      Width           =   2148
   End
   Begin VB.TextBox OT_SocketAsociado 
      Height          =   285
      Left            =   2460
      TabIndex        =   2
      Top             =   4815
      Visible         =   0   'False
      Width           =   1470
   End
   Begin VB.TextBox OT_Canal 
      Height          =   375
      Left            =   3990
      TabIndex        =   1
      Top             =   4800
      Visible         =   0   'False
      Width           =   1470
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
      Height          =   1212
      HelpContextID   =   27
      Left            =   2388
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   3480
      Width           =   6108
   End
   Begin RichTextLib.RichTextBox OL_Estatus 
      Height          =   2976
      HelpContextID   =   27
      Left            =   2352
      TabIndex        =   4
      Top             =   468
      Width           =   6024
      _ExtentX        =   10610
      _ExtentY        =   5265
      _Version        =   327680
      BackColor       =   16777215
      HideSelection   =   0   'False
      ScrollBars      =   2
      DisableNoScroll =   -1  'True
      TextRTF         =   $"OF_Canal.frx":0462
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
      Left            =   7884
      Top             =   468
      _ExtentX        =   794
      _ExtentY        =   794
      BackColor       =   -2147483643
      ImageWidth      =   20
      ImageHeight     =   20
      _Version        =   327680
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   8
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "OF_Canal.frx":052B
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "OF_Canal.frx":0845
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "OF_Canal.frx":0B5F
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "OF_Canal.frx":0E79
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "OF_Canal.frx":1193
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "OF_Canal.frx":14AD
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "OF_Canal.frx":17C7
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "OF_Canal.frx":1AE1
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "OF_Hablar_Canal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' Variables locales a la forma
' LF_Historial es un arreglo de strings que lleva un historial
' de los ultimos 20 mensajes digitados por el usuario
Dim LF_Historial(20) As String
' LF_Ultimo lleva el indice del ultimo indice utilizado en el
' Historial
' Y LF_actual lleva la posición en que se encuentra en el historial
' si un usuario esta navegando  en el
Dim LF_Notificacion As Boolean
Dim LF_Ultimo%, LF_Actual%

Private Sub Form_Activate()
If LF_Notificacion Then
    Me.Icon = LoadPicture(App.Path + "\hands.ico")
    If Me.Tag <> "" Then
      If GV_Explorador And Trim(Me.Tag) <> "" Then
          OF_Explorador.OA_Explorador.Nodes.Item( _
          "C" + CStr(Me.Tag)).Image = 2
      End If
    End If
End If
OL_Estatus.SelStart = Len(OL_Estatus.Text)

End Sub

Private Sub Form_Load()
'/*******************************************************/
' Al cargarse la forma se inicializan las variables
' utilizadas en el historial, el textbox de mensajes es
' blanqueado
'/*******************************************************/
' Variable utilizada para mensajes en el canal
LF_Notificacion = False
OL_Estatus = ""
LF_Ultimo = 0
LF_Actual = 0

End Sub
Private Sub Form_Resize()
'/*******************************************************/
' Cuando se cambia el tamaño de la forma entonces ajustamos
' los controles de la forma a otros tamaños de acuerdo al
' nuevo tamaño de la forma
'/********************************************************/

On Error GoTo Etiqueta_Error:

If ScaleWidth <> 0 And ScaleHeight <> 0 And _
ScaleHeight > 500 And ScaleWidth > 500 Then
  OL_Usuarios.Move 0, 450, ScaleWidth / 4, ScaleHeight - 400
  OL_Estatus.Move (ScaleWidth / 4), 450, _
   ScaleWidth - ScaleWidth / 4, ScaleHeight - 1050
  OT_Comando.Move (ScaleWidth / 4), _
  ScaleHeight - 600, ScaleWidth - ScaleWidth / 4, 600
    
  OL_Estatus.RightMargin = ScaleWidth - 500

End If

Exit Sub
Etiqueta_Error:
 
End Sub

Private Sub Form_Unload(Cancel As Integer)
'/*********************************************************/
' Al descargar una forma de Canales se envia el mensaje al
' servidor especificando que el usuario esta saliendo del
' Canal. El mensaje es  PART #NombreCanal
'/*********************************************************/
On Error GoTo Etiqueta_Error:
If OT_SocketAsociado <> 0 Then
    MM_Enviar_Mensaje "Part " + OT_Canal + _
    GV_EOD, OT_SocketAsociado
End If
If Trim(Me.Tag) = "" Then Exit Sub
GV_Estado_Canal(Me.Tag).Deleted = True
If GV_Explorador Then ' Si el explorador esta arriba, borra
                      ' el Item que representaba el canal
                      ' en el explorador
  OF_Explorador.OA_Explorador.Nodes.Remove ("C" + CStr(Me.Tag))
 
End If

Exit Sub
Etiqueta_Error:
ME_Muestra_Error
End Sub

Private Sub OL_Estatus_Change()
' /****************************************************/
' Evento que nos indica de nuevos mensajes a la ventana
' /****************************************************/

On Error GoTo Etiqueta_Error:

If Me.OT_SocketAsociado <> 0 Then
 ' Si la notificación de Mensajes esta activa
 If LF_Notificacion Then
    If Screen.ActiveForm.name <> "OF_Hablar_Canal" Then
        Beep
        Me.Icon = LoadPicture(App.Path + "\Light.ico")
        If Me.Tag <> "" Then
            If GV_Explorador Then
             OF_Explorador.OA_Explorador.Nodes.Item( _
             "C" + CStr(Me.Tag)).Image = 10
            End If
        End If
    ElseIf Screen.ActiveForm.Tag <> Me.Tag Then
        Beep
        Me.Icon = LoadPicture(App.Path + "\Light.ico")
        If Me.Tag <> "" Then
            If GV_Explorador Then
             OF_Explorador.OA_Explorador.Nodes.Item( _
             "C" + CStr(Me.Tag)).Image = 10
            End If
        End If
    End If
    
  End If
End If

Exit Sub
Etiqueta_Error:
ME_Muestra_Error
End Sub

Private Sub OL_Estatus_GotFocus()
If LF_Notificacion Then
    Me.Icon = LoadPicture(App.Path + "\hands.ico")
    If Me.Tag = "" Then
        If GV_Explorador Then
             OF_Explorador.OA_Explorador.Nodes.Item( _
             "C" + CStr(Me.Tag)).Image = 2
        End If
    End If
End If
    OL_Estatus.SelStart = Len(OL_Estatus.Text)
End Sub

Private Sub OL_Estatus_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'/***************************************************/
' Si se hace click con el botón derecho sobre el
' Textbox donde se reciben mensajes y se ha seleccionado
' una porción de texto se procede
' ha habilitar el menu que presenta la opción de abrir
' una pagina WEB
'/****************************************************/

If Button = 2 Then
  If Trim(OL_Estatus.SelText) = "" Then Exit Sub
  GV_Nombre = OL_Estatus.SelText
  PopupMenu OF_principal.M_ficticio
  
End If
End Sub

Private Sub OL_usuarios_DblClick()
'/******************************************************/
' Si se hace Doble Click en un usuario de la lista de
' canales, entonces se procede a abrir una ventana de
' platica privada con ese usuario
'/******************************************************/

On Error GoTo Etiqueta_Error:
Dim L_vent%, L_Alias$
If Me.OT_SocketAsociado <> 0 Then
    L_Alias = OL_Usuarios.List(OL_Usuarios.ListIndex)
    If Left(L_Alias, 1) = "@" Then _
     L_Alias = Trim(Right(L_Alias, Len(L_Alias) - 1))
    L_vent = _
    MV_Busca_Ventana_Usuario(L_Alias, Me.OT_SocketAsociado)
    If L_vent = 0 Then
       L_vent = MV_CreaVentana_Usuario( _
        Me.OT_SocketAsociado, L_Alias)
       GV_VENTANAS_Usuario(L_vent).Caption = _
       L_Alias + ":[ " + _
       GV_Sockets(OT_SocketAsociado).Direcc + "]"
    End If
    If GV_VENTANAS_Usuario(L_vent).Visible Then
       GV_VENTANAS_Usuario(L_vent).SetFocus
     End If
End If

Exit Sub
Etiqueta_Error:
ME_Muestra_Error
End Sub

Private Sub OT_Comando_GotFocus()
OL_Estatus.SelStart = Len(OL_Estatus.Text)
End Sub

Private Sub OT_comando_KeyDown(KeyCode As Integer, Shift As Integer)
'/*********************************************************/
' Si se presionan las teclas Ctrl-Up o Ctrl-down, entonces
' se recorre el historial de comandos
'/*********************************************************/

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
'/*********************************************************/
' Cuando se presiona <ENTER> en el textbox de comandos
' se procede a verificar si la ventana todavia posee una
' conexión activo(OT_SocketAsociado<>0).
' Si la conexión todavía se encuentra activa, se
'  procede a procesar el  mensaje digitado.
' Una vez procesado(haber buscado su equivalente en la
' tabla de comandos )
' se procede a enviarlo por el socket Asociado.
' A la vez registra el mensaje en el historial de
' mensajes de la ventana.
'/**********************************************************/
On Error GoTo Etiqueta_Error:
Dim L_res%
Dim L_Texto$, L_Retorno$

If KeyAscii = 13 And Len(Trim(OT_Comando)) > 0 Then
    If Me.OT_SocketAsociado <> 0 Then
       L_Texto = MM_Verificar_Comando(OT_Comando, OT_Canal)
       If Left(Trim(OT_Comando), 1) <> "/" Then
          MV_Pone_Mensaje True, Me, _
          "<" + _
          GV_Sockets(Me.OT_SocketAsociado).Nick + ">: ", _
          GV_RojoAlgo
          
          MV_Pone_Mensaje True, Me, _
          Trim(OT_Comando) + GV_EOD, GV_Azul
        End If
        L_Texto = L_Texto & GV_EOD
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
    Else
        MV_Pone_Mensaje True, Me, _
        "Usted no se encuentra conectado al " + _
        "servidor ..." & GV_EOD, vbRed
    End If
    OT_Comando = ""
    KeyAscii = 8
ElseIf KeyAscii = 13 Then KeyAscii = 8
Exit Sub
Etiqueta_Error:
ME_Muestra_Error
End If
End Sub

Private Sub OT_Comando_KeyUp(KeyCode As Integer, Shift As Integer)
'/***************************************************/
' Procedimiento que nos permite movernos entre las
' ventanas de Canales si se presionan las teclas de
' Shift-Up o Shift_Down.
'/***************************************************/
On Error GoTo Etiqueta_Error:
Dim L_cual%
If Shift = 1 Then ' shift key
   Select Case KeyCode
    
        Case 38 ' Up key
          L_cual = _
           MV_BuscaSigAnt_Ventana_Canal( _
           Me.Tag, "ANTERIOR", False)
          
           If GV_VENTANAS_Canal(L_cual).Visible Then
            GV_VENTANAS_Canal(L_cual).OT_Comando.SetFocus
              If GV_VENTANAS_Canal(L_cual).WindowState = 1 _
               Then GV_VENTANAS_Canal(L_cual).WindowState = 0
           End If

          
        Case 40 ' Downkey
          L_cual = _
          MV_BuscaSigAnt_Ventana_Canal( _
          Me.Tag, "SIGUIENTE", False)
          
          If GV_VENTANAS_Canal(L_cual).Visible Then
           GV_VENTANAS_Canal(L_cual).OT_Comando.SetFocus
            If GV_VENTANAS_Canal(L_cual).WindowState = 1 _
            Then GV_VENTANAS_Canal(L_cual).WindowState = 0
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
' del toolbar de la ventana fue seleccionada
'/********************************************************/

On Error GoTo Etiqueta_Error:
Dim L_cual%
Dim L_Ventana As New OF_Enviar
Dim L_Ventana1 As New OF_Retransmision

Select Case Button.Index
Case 1 ' Opción de Enviar Archivos
    If OT_SocketAsociado = 0 Then Exit Sub
    Load L_Ventana
    If Trim(OL_Estatus.SelText) <> "" Then
      L_Ventana.OT_Nick = OL_Estatus.SelText
    End If
    L_Ventana.Show
   
Case 2 ' Opción de Retransmisión de Mensajes
    If OT_SocketAsociado = 0 Then Exit Sub
   Load L_Ventana1
   L_Ventana1.Show
   
Case 3 ' Opción de Canales Preferidos
    GV_Seleccion = Me.OT_SocketAsociado
    Load OF_Canales_Preferidos
    OF_Canales_Preferidos.Show
Case 4 ' Para activar o desactivar la notificación
       ' de mensajes
    If LF_Notificacion Then
      LF_Notificacion = False
      Button.Image = 7
    Else
      LF_Notificacion = True
      Button.Image = 8
    End If

Case 5 ' Opción para cerrar la ventana (Salir del Canal)
    Unload Me
   
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
