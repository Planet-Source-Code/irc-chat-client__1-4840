VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.1#0"; "comctl32.ocx"
Begin VB.Form OF_Lista_Canales 
   Caption         =   "Listado de Canales"
   ClientHeight    =   4455
   ClientLeft      =   1605
   ClientTop       =   1200
   ClientWidth     =   5100
   HelpContextID   =   19
   Icon            =   "OF_ListCan.frx":0000
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4455
   ScaleWidth      =   5100
   Begin ComctlLib.Toolbar OTB_Estatus 
      Align           =   1  'Align Top
      Height          =   450
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   5100
      _ExtentX        =   8996
      _ExtentY        =   794
      ButtonWidth     =   714
      ButtonHeight    =   688
      AllowCustomize  =   0   'False
      HelpContextID   =   19
      ImageList       =   "ImageList1"
      _Version        =   327680
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   4
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Enviar Archivo"
            Object.Tag             =   ""
            ImageIndex      =   5
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Retransmisión de Canales"
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
      Height          =   300
      HelpContextID   =   19
      Left            =   60
      TabIndex        =   4
      Top             =   4008
      Width           =   4896
   End
   Begin VB.TextBox OT_Ventana_Estatus 
      Height          =   288
      Left            =   2880
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   4056
      Visible         =   0   'False
      Width           =   768
   End
   Begin VB.TextBox OT_SocketAsociado 
      Height          =   288
      Left            =   3768
      TabIndex        =   1
      Top             =   4056
      Visible         =   0   'False
      Width           =   1164
   End
   Begin VB.ListBox OL_Estatus 
      Height          =   3375
      HelpContextID   =   19
      Left            =   60
      TabIndex        =   0
      Top             =   420
      Width           =   4860
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   5028
      Top             =   492
      _ExtentX        =   794
      _ExtentY        =   794
      BackColor       =   -2147483643
      ImageWidth      =   20
      ImageHeight     =   20
      _Version        =   327680
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   6
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "OF_ListCan.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "OF_ListCan.frx":075C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "OF_ListCan.frx":0A76
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "OF_ListCan.frx":0D90
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "OF_ListCan.frx":10AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "OF_ListCan.frx":13C4
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "OF_Lista_Canales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' Variables locales a la forma
' LF_Historial es un arreglo de strings que lleva un
' historial de los ultimos 20 mensajes digitados por el
' usuario
Dim LF_Historial(20) As String
' LF_Ultimo lleva el indice del ultimo indice utilizado
' en el Historial  Y LF_actual lleva la posición en que
' se encuentra en el historial si un usuario esta
' navegando  en el

Dim LF_Ultimo%, LF_Actual%

Private Sub Form_Load()
'/**********************************************************/
' Al cargarse la forma se inicializan las variables
' utilizadas en el historial, el textbox de mensajes es
' blanqueado
'/**********************************************************/
LF_Ultimo = 0
LF_Actual = 0
End Sub

Private Sub Form_Resize()
'/*********************************************************/
' Cuando se cambia el tamaño de la forma entonces ajustamos
' los controles de la forma a otros tamaños de acuerdo al
' nuevo tamaño de la forma
'/********************************************************/

On Error GoTo Etiqueta_Error:

If ScaleWidth <> 0 And ScaleHeight <> 0 And _
ScaleHeight > 400 And ScaleWidth > 400 Then
  OL_Estatus.Move 0, 450, ScaleWidth, ScaleHeight - 950
  OT_Comando.Move 0, ScaleHeight - 500, ScaleWidth, 500

End If
Exit Sub

Etiqueta_Error:
End Sub

Private Sub Form_Unload(Cancel As Integer)
'/**********************************************************/
' Al descargar una forma de Lista de Canales se envia ningún
' mensaje al servidor simplemente se cierra la ventana y se
' marca como borrada en el arreglo de estados de ventanas.
'/**********************************************************/


If GV_Explorador Then
 OF_Explorador.OA_Explorador.Nodes.Remove ("Z" + CStr(Me.Tag))
End If
GV_Estado_Lista_Canales(Me.Tag).Deleted = True
End Sub

Private Sub OL_Estatus_DblClick()
'/**********************************************************/
' Si se hace doble click en un item de la lista de canales,
' se procede a enviar un mensaje de JOIN al canal seleccionado.
'/**********************************************************/
Dim L_Canal$
If OL_Estatus.ListIndex <> -1 And Me.OT_SocketAsociado <> 0 _
  Then
    L_Canal = _
     Left(OL_Estatus.List(OL_Estatus.ListIndex), _
     InStr(1, _
     OL_Estatus.List(OL_Estatus.ListIndex), " ") - 1)
    
    L_Canal = Trim(L_Canal)
    MM_Enviar_Mensaje "JOIN " + L_Canal, Me.OT_SocketAsociado
End If

End Sub

Private Sub OT_comando_KeyDown(KeyCode As Integer, Shift As Integer)
'/********************************************************/
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
'/********************************************************/
' Cuando se presiona <ENTER> en el textbox de comandos se
' procede a verificar si la ventana todavia posee una
' conexión activo(OT_SocketAsociado<>0).
' Si la conexión todavía se encuentra activa, se procede a
' procesar el mensaje digitado.
' Una vez procesado(haber buscado su equivalente en la
' tabla de comandos )
' se procede a enviarlo por el socket Asociado.
'  A la vez registra el mensaje en el historial de mensajes
' de la ventana.
'/**********************************************************/

Dim L_res%
Dim L_Texto$, L_Retorno$

If KeyAscii = 13 And Len(OT_Comando) > 0 Then
    If Me.OT_SocketAsociado <> 0 Then
        L_Texto = MM_Verificar_Comando(OT_Comando, "")
        If Left(Trim(OT_Comando), 1) <> "/" Then
          Exit Sub
        End If
        L_Texto = L_Texto & GV_EOD
        L_res = _
          send(GV_Sockets(Me.OT_SocketAsociado).socket, _
          L_Texto, Len(L_Texto), 0)
       
        If L_res = SOCKET_ERROR Then
            L_res = WSAGetLastError()
            MV_Pone_Mensaje False, Me, _
             ME_WsockError(L_res), vbBlack
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
        MV_Pone_Mensaje False, Me, _
        "No esta conectado a ningún servidor ..." & _
        GV_EOD, vbRed
    End If
    OT_Comando = ""
    KeyAscii = 8
ElseIf KeyAscii = 13 Then KeyAscii = 8

End If

End Sub


Private Sub OTB_Estatus_ButtonClick(ByVal Button As Button)
'/*******************************************************/
' Este es el código utilizado para determinar que opción
' del toolbar de la  ventana fue seleccionada
'/*******************************************************/

On Error GoTo Etiqueta_Error:
Dim L_cual%
Dim L_Ventana As New OF_Enviar
Dim L_Ventana1 As New OF_Retransmision

Select Case Button.Index
Case 1 ' Opción de Enviar Archivos
    If OT_SocketAsociado = 0 Then Exit Sub
    Load L_Ventana
    L_Ventana.Show
   
Case 2 ' Opción de Retransmisión de MEnsajes
   If OT_SocketAsociado = 0 Then Exit Sub
   Load L_Ventana1
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
