VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "COMCTL32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.1#0"; "RICHTX32.OCX"
Begin VB.Form OF_Hablar_Usuario 
   Caption         =   "Hablar Con Usuario"
   ClientHeight    =   4680
   ClientLeft      =   630
   ClientTop       =   1650
   ClientWidth     =   7110
   HelpContextID   =   21
   Icon            =   "OF_HablU.frx":0000
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4680
   ScaleWidth      =   7110
   Begin ComctlLib.Toolbar OTB_Estatus 
      Align           =   1  'Align Top
      Height          =   450
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   7110
      _ExtentX        =   12541
      _ExtentY        =   794
      ButtonWidth     =   714
      ButtonHeight    =   688
      AllowCustomize  =   0   'False
      HelpContextID   =   21
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   4
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
            Object.ToolTipText     =   "Cerrar Ventana"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
      EndProperty
   End
   Begin VB.TextBox OT_Ventana_Estatus 
      Height          =   288
      Left            =   1728
      TabIndex        =   5
      Top             =   4404
      Visible         =   0   'False
      Width           =   768
   End
   Begin VB.TextBox OT_Nick 
      Height          =   375
      Left            =   4215
      TabIndex        =   2
      Top             =   4470
      Visible         =   0   'False
      Width           =   1470
   End
   Begin VB.TextBox OT_SocketAsociado 
      Height          =   285
      Left            =   2550
      TabIndex        =   1
      Top             =   4455
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
      Height          =   840
      HelpContextID   =   21
      HideSelection   =   0   'False
      Left            =   96
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   3804
      Width           =   6930
   End
   Begin RichTextLib.RichTextBox OL_Estatus 
      Height          =   3255
      HelpContextID   =   21
      Left            =   90
      TabIndex        =   3
      Top             =   495
      Width           =   6840
      _ExtentX        =   12065
      _ExtentY        =   5741
      _Version        =   327680
      Enabled         =   -1  'True
      HideSelection   =   0   'False
      ScrollBars      =   2
      DisableNoScroll =   -1  'True
      TextRTF         =   $"OF_HablU.frx":0442
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
      Left            =   5928
      Top             =   564
      _ExtentX        =   794
      _ExtentY        =   794
      BackColor       =   -2147483643
      ImageWidth      =   20
      ImageHeight     =   20
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   6
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "OF_HablU.frx":050B
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "OF_HablU.frx":0825
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "OF_HablU.frx":0B3F
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "OF_HablU.frx":0E59
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "OF_HablU.frx":1173
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "OF_HablU.frx":148D
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "OF_Hablar_Usuario"
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
' Historial Y LF_actual lleva la posición en que se encuentra
' en el historial
' si un usuario esta navegando  en el
Dim LF_Ultimo%, LF_Actual%

Private Sub Form_Activate()
'/*********************************************************/
' Procedimiento que cambia el icono de la ventana cuando
' esta es Activada, es utilizado para diferencias a que
' ventanas les llego un nuevo mensaje y este no ha sido
' leido
'/*********************************************************/
Me.Icon = LoadPicture(App.Path + "\Face.ico")
If Trim(Me.Tag) = "" Then Exit Sub
If GV_Explorador Then
 OF_Explorador.OA_Explorador.Nodes.Item( _
 "U" + CStr(Me.Tag)).Image = 4
End If
End Sub

Private Sub Form_Load()
'/********************************************************/
' Al cargarse la forma se inicializan las variables
' utilizadas en el historial, el textbox de mensajes es
' blanqueado
'/********************************************************/
OL_Estatus = ""
LF_Ultimo = 0
LF_Actual = 0

End Sub

Private Sub Form_Resize()
'/********************************************************/
' Cuando se cambia el tamaño de la forma entonces ajustamos
' los controles de la forma a otros tamaños de acuerdo al
' nuevo tamaño de la forma
'/********************************************************/
On Error GoTo Etiqueta_Error:
If ScaleWidth <> 0 And ScaleHeight <> 0 And _
ScaleWidth > 500 And ScaleHeight > 500 Then
  OL_Estatus.Move 0, 450, ScaleWidth, ScaleHeight - 1050
  OT_Comando.Move 0, ScaleHeight - 600, ScaleWidth, 600
  OL_Estatus.RightMargin = ScaleWidth - 500

End If
Exit Sub

Etiqueta_Error:

End Sub

Private Sub Form_Unload(Cancel As Integer)
'/*********************************************************/
' Al descargar una forma de usuario no  se envia ningún
' mensaje al servidor simplemente se cierra la ventana y
' se marca como borrada en el arreglo
' de estados de ventanas.
'/**********************************************************/
If Me.Tag = "" Then Exit Sub

If GV_Explorador Then '  Si el explorador esta arriba,
                      ' borra el Item que
                      ' representaba el usuario en el
                      ' explorador
 
 OF_Explorador.OA_Explorador.Nodes.Remove ("U" + CStr(Me.Tag))
 
End If
GV_Estado_Usuario(Me.Tag).Deleted = True
End Sub

Private Sub OL_Estatus_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'/********************************************************/
' Si se hace click con el botón derecho sobre el Textbox
' donde se reciben  mensajes y se ha seleccionado una
' porción de texto se procede ha habilitar el menu que
' presenta la opción de abrir una pagina WEB
'/********************************************************/

If Button = 2 Then
  If Trim(OL_Estatus.SelText) = "" Then Exit Sub
  GV_Nombre = OL_Estatus.SelText
  PopupMenu OF_principal.M_ficticio
  
End If
End Sub

Private Sub OT_Comando_GotFocus()
'/********************************************************/
' Si el textbox de comandos en la ventana de usuario
' recibe el focus y el explorador esta arriba entonces se
' procede a cambiar el icono en el explorador. El icono
' rojo representa que no se ha leido el  mensaje que llego
' a la ventana.
'/********************************************************/
Me.Icon = LoadPicture(App.Path + "\Face.ico")
If GV_Explorador Then
 OF_Explorador.OA_Explorador.Nodes.Item( _
   "U" + CStr(Me.Tag)).Image = 4
End If
OL_Estatus.SelStart = Len(OL_Estatus.Text)
End Sub

Private Sub OT_comando_KeyDown(KeyCode As Integer, Shift As Integer)
'/*********************************************************/
' Si se presionan las teclas Ctrl-Up o Ctrl-down, entonces
' se recorre el historial de comandos
'/*********************************************************/

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
End Sub

Private Sub OT_comando_KeyPress(KeyAscii As Integer)
'/*******************************************************/
' Cuando se presiona <ENTER> en el textbox de comandos
' se procede a verificar si la ventana todavia posee una
' conexión activo(OT_SocketAsociado<>0).
' Si la conexión todavía se encuentra activa, se
' procede a procesar el  mensaje digitado.
' Una vez procesado(haber buscado su equivalente en la
' tabla de comandos )
' se procede a enviarlo por el socket Asociado.
'  A la vez registra el mensaje en el historial de
' mensajes de la ventana.
'/******************************************************/
Dim L_res%
Dim L_Texto$, L_Retorno$

If KeyAscii = 13 And Len(OT_Comando) > 0 Then
    If Me.OT_SocketAsociado <> 0 Then
        L_Texto = MM_Verificar_Comando(OT_Comando, OT_Nick)
        If Left(Trim(OT_Comando), 1) <> "/" Then
            MV_Pone_Mensaje True, Me, _
            "<" + GV_Sockets(Me.OT_SocketAsociado).Nick + _
            ">: ", GV_RojoAlgo
            MV_Pone_Mensaje True, Me, _
            Trim(OT_Comando) + GV_EOD, GV_Azul
        End If
        
        L_Texto = L_Texto & GV_EOD
        L_res = send(GV_Sockets(Me.OT_SocketAsociado).socket, _
        L_Texto, Len(L_Texto), 0)
    
        If L_res = SOCKET_ERROR Then
         L_res = WSAGetLastError()
         MV_Pone_Mensaje False, Me, ME_WsockError(L_res), vbRed
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
        "Usted no se encuentra conectado al " + _
        "servidor ..." & GV_EOD, vbRed
    End If
    OT_Comando = ""
    KeyAscii = 8
ElseIf KeyAscii = 13 Then KeyAscii = 8
    
End If
End Sub

Private Sub OT_Comando_KeyUp(KeyCode As Integer, Shift As Integer)
'/***************************************************/
' Procedimiento que nos permite movernos entre las
' ventanas de Canales si se presionan las teclas de
' Shift-Up o Shift_Down.
'/****************************************************/
Dim L_cual%
If Shift = 1 Then ' shift key
   Select Case KeyCode
    
        Case 38 ' Up key
          L_cual = _
          MV_BuscaSigAnt_Ventana_Usuario( _
          Me.Tag, "ANTERIOR", False)
         
          If GV_VENTANAS_Usuario(L_cual).Visible Then
             GV_VENTANAS_Usuario(L_cual).SetFocus
             If GV_VENTANAS_Usuario( _
             L_cual).WindowState = 1 _
             Then GV_VENTANAS_Usuario(L_cual).WindowState = 0
            End If

         
        Case 40 ' Downkey
          L_cual = _
          MV_BuscaSigAnt_Ventana_Usuario( _
          Me.Tag, "SIGUIENTE", False)
         
            If GV_VENTANAS_Usuario(L_cual).Visible Then
             GV_VENTANAS_Usuario(L_cual).SetFocus
             If GV_VENTANAS_Usuario( _
             L_cual).WindowState = 1 Then _
             GV_VENTANAS_Usuario(L_cual).WindowState = 0
            End If
         
  End Select
End If
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
    If OT_SocketAsociado = 0 Then Exit Sub
    Load L_Ventana
    If Trim(OL_Estatus.SelText) <> "" Then
      L_Ventana.OT_Nick = OL_Estatus.SelText
    Else
     L_Ventana.OT_Nick = OT_Nick
    End If
    L_Ventana.Show
   
Case 2 ' Opción de Retransmisión de MEnsajes
    If OT_SocketAsociado = 0 Then Exit Sub
    Load L_Ventana1
    L_Ventana1.OT_Alias_Emisor = OT_Nick
    L_Ventana1.Show

Case 3 ' Opción de Canales Preferidos
    
    GV_Seleccion = Me.OT_SocketAsociado
    Load OF_Canales_Preferidos
    OF_Canales_Preferidos.Show

Case 4 ' Opción para cerrar la ventana
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
