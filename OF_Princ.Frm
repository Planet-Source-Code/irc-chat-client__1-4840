VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.MDIForm OF_principal 
   BackColor       =   &H8000000C&
   Caption         =   "Sistema ERG2 v1.0"
   ClientHeight    =   4755
   ClientLeft      =   75
   ClientTop       =   1005
   ClientWidth     =   9480
   HelpContextID   =   1
   Icon            =   "OF_Princ.frx":0000
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  'Maximized
   Begin ComctlLib.Toolbar OTB_Toolbar 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   9480
      _ExtentX        =   16722
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   13
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Usuario"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Servidores"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Conectarse a Ultimo Servidor"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Desconectarse de Servidores"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Comandos"
            Object.Tag             =   ""
            ImageIndex      =   12
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Tipos Servidores"
            Object.Tag             =   ""
            ImageIndex      =   13
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Servidores Favoritos"
            Object.Tag             =   ""
            ImageIndex      =   17
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Canales Favoritos"
            Object.Tag             =   ""
            ImageIndex      =   7
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Browser de Internet"
            Object.Tag             =   ""
            ImageIndex      =   14
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Enviar Archivos"
            Object.Tag             =   ""
            ImageIndex      =   20
         EndProperty
         BeginProperty Button11 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Explorador de ERG2"
            Object.Tag             =   ""
            ImageIndex      =   11
         EndProperty
         BeginProperty Button12 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Retransmisión de Mensajes"
            Object.Tag             =   ""
            ImageIndex      =   18
         EndProperty
         BeginProperty Button13 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Salir del Sistema"
            Object.Tag             =   ""
            ImageIndex      =   21
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin MSComDlg.CommonDialog CDialog 
      Left            =   456
      Top             =   492
      _ExtentX        =   688
      _ExtentY        =   688
      _Version        =   393216
      HelpContext     =   1
      HelpFile        =   "erg2.hlp"
   End
   Begin ComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   0
      Top             =   4455
      Width           =   9480
      _ExtentX        =   16722
      _ExtentY        =   529
      Style           =   1
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   9960
      Top             =   720
      _ExtentX        =   794
      _ExtentY        =   794
      BackColor       =   -2147483636
      ImageWidth      =   16
      ImageHeight     =   16
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   21
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "OF_Princ.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "OF_Princ.frx":075C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "OF_Princ.frx":0A76
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "OF_Princ.frx":0D90
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "OF_Princ.frx":10AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "OF_Princ.frx":13C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "OF_Princ.frx":16DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "OF_Princ.frx":19F8
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "OF_Princ.frx":1D12
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "OF_Princ.frx":202C
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "OF_Princ.frx":2346
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "OF_Princ.frx":2660
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "OF_Princ.frx":297A
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "OF_Princ.frx":2C94
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "OF_Princ.frx":2FAE
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "OF_Princ.frx":32C8
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "OF_Princ.frx":35E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "OF_Princ.frx":38FC
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "OF_Princ.frx":3C16
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "OF_Princ.frx":3F30
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "OF_Princ.frx":424A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu M_Archivo 
      Caption         =   "&Opciones"
      HelpContextID   =   28
      Begin VB.Menu M_Usuario 
         Caption         =   "&Usuario "
         HelpContextID   =   2
      End
      Begin VB.Menu M_Servidores 
         Caption         =   "S&ervidores"
         HelpContextID   =   4
      End
      Begin VB.Menu barra_02 
         Caption         =   "-"
      End
      Begin VB.Menu M_Conexion 
         Caption         =   "&Conectarse a Ultimo Servidor"
         HelpContextID   =   6
      End
      Begin VB.Menu M_Desconexion 
         Caption         =   "D&esconectarse de Servidores"
         HelpContextID   =   7
      End
      Begin VB.Menu barra_01 
         Caption         =   "-"
      End
      Begin VB.Menu M_Config 
         Caption         =   "C&onfiguración"
         HelpContextID   =   28
         Begin VB.Menu M_Comandos 
            Caption         =   "&Comandos"
            HelpContextID   =   11
         End
         Begin VB.Menu M_Tipos 
            Caption         =   "&Tipos de Servidores"
            HelpContextID   =   3
         End
         Begin VB.Menu M_PATH 
            Caption         =   "&Registrar PATH del Browser"
            HelpContextID   =   25
         End
      End
      Begin VB.Menu M_Favoritos 
         Caption         =   "&Favoritos"
         HelpContextID   =   28
         Begin VB.Menu M_F_Servidores 
            Caption         =   "&Servidores"
            HelpContextID   =   5
         End
         Begin VB.Menu M_F_Canales 
            Caption         =   "&Canales"
            HelpContextID   =   27
         End
      End
      Begin VB.Menu barra 
         Caption         =   "-"
      End
      Begin VB.Menu M_Internet 
         Caption         =   "&Browser de Internet"
         HelpContextID   =   25
      End
      Begin VB.Menu M_Enviar 
         Caption         =   "E&nviar Archivos"
         HelpContextID   =   23
      End
      Begin VB.Menu M_Explorador 
         Caption         =   "E&xplorador de ERG2"
         HelpContextID   =   22
      End
      Begin VB.Menu M_Retransmisión 
         Caption         =   "Re&transmisión de Mensajes"
         HelpContextID   =   26
      End
      Begin VB.Menu barra2 
         Caption         =   "-"
      End
      Begin VB.Menu M_Salir 
         Caption         =   "&Salir"
      End
   End
   Begin VB.Menu M_Ventanas 
      Caption         =   "&Ventanas"
      HelpContextID   =   22
      WindowList      =   -1  'True
      Begin VB.Menu M_Vent 
         Caption         =   "Organizar en &Cascada"
         HelpContextID   =   22
         Index           =   0
      End
      Begin VB.Menu M_Vent 
         Caption         =   "Organizar &Verticalmente"
         HelpContextID   =   22
         Index           =   1
      End
      Begin VB.Menu M_Vent 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu M_Vent 
         Caption         =   "Organizar &Horizontalmente"
         HelpContextID   =   22
         Index           =   3
      End
      Begin VB.Menu M_Vent 
         Caption         =   "Organizar &Iconos"
         HelpContextID   =   22
         Index           =   4
      End
      Begin VB.Menu M_Vent 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu M_Vent 
         Caption         =   "&Minimizar Ventanas"
         HelpContextID   =   22
         Index           =   6
      End
      Begin VB.Menu M_Vent 
         Caption         =   "&Restaurar Ventanas"
         HelpContextID   =   22
         Index           =   7
      End
      Begin VB.Menu M_Vent 
         Caption         =   "M&aximizar Ventanas"
         HelpContextID   =   22
         Index           =   8
      End
   End
   Begin VB.Menu M_Ayuda 
      Caption         =   "A&yuda"
      HelpContextID   =   1
      Begin VB.Menu M_Contenido 
         Caption         =   "&Contenido"
         HelpContextID   =   1
      End
      Begin VB.Menu M_About 
         Caption         =   "&Acerca de ERG2"
         HelpContextID   =   30
      End
      Begin VB.Menu M_AWinsock 
         Caption         =   "A&cerca de Winsock"
         HelpContextID   =   30
      End
   End
   Begin VB.Menu M_ficticio 
      Caption         =   "Ficticio"
      Visible         =   0   'False
      Begin VB.Menu M_Pag_Internet 
         Caption         =   "Abrir página de Internet"
      End
   End
End
Attribute VB_Name = "OF_principal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub M_About_Click()
' Muestra la ventana de About
OF_About.Show
End Sub

Private Sub M_AWinsock_Click()
' Muestra la información de Winsock
MM_About_WINSOCK GV_WSAdata
End Sub

Private Sub M_Comandos_Click()
' Muestra la ventana de Comandos
OF_Comandos.Show
End Sub

Private Sub M_Conexion_Click()
'/*********************************************************/
' Procedimiento que obtiene el ultimo servidor al cual se
' conecto o intento conectarse el usuario , y se procede a
' realizar la conexión con la información de ese servidor
'/**********************************************************/

On Error GoTo Etiqueta_Error:
Dim L_Servidor As Recordset

Set L_Servidor = GV_Base_De_Datos.OpenRecordset( _
         "select Direccion, U_Puerto, Tipo_Servidor " + _
         " from Servidores, Ini where Codigo = Ultimo_Servidor")

MM_Connect L_Servidor("direccion"), L_Servidor("u_puerto"), _
L_Servidor("tipo_servidor"), 0

L_Servidor.Close

Exit Sub

Etiqueta_Error:
ME_Muestra_Error


End Sub

Private Sub M_Contenido_Click()
' Muestra la ayuda de la aplicación
On Error GoTo Etiqueta_Error:
  
Dim L_Ret As Integer
'Llamar al archivo de Ayuda en la sección de Contenido
L_Ret = OSWinHelp(OF_principal.hwnd, App.HelpFile, _
HelpConstants.cdlHelpContents, 0)
  
Exit Sub
Etiqueta_Error:
ME_Muestra_Error
End Sub

Private Sub M_Desconexion_Click()
' Muestra la ventana Desconexión de Servidores
OF_Cerrar.Show
End Sub

Private Sub M_Enviar_Click()
' Muestra la ventana de Enviar Archivos
Dim L_Ventana As New OF_Enviar

Load L_Ventana
L_Ventana.Show

End Sub

Private Sub M_Explorador_Click()
'/*****************************************************/
' Procedimiento que muestra el explorador si esta ya se
' encuentra activo o carga el explorador, y arregla el
' MDI y el explorador en toda la pantalla.
'/*****************************************************/
 If Not GV_Explorador Then
  OF_Explorador.Show
  
 Else
  OF_Explorador.WindowState = 0
  
  OF_Explorador.Move 0, 0, 2500, Screen.Height - 300
  OF_principal.WindowState = 0
  OF_principal.Move 0 + OF_Explorador.Width, 0, _
  Screen.Width - OF_Explorador.Width, Screen.Height - 300
  
 End If

End Sub

Private Sub M_F_Canales_Click()
' Muestra la ventana de Canales preferidos
Load OF_Canales_Preferidos
GV_Seleccion = 0
OF_Canales_Preferidos.Show
End Sub

Private Sub M_F_Servidores_Click()
' Muestra la ventana de servidores favoritos
OF_Servidores_Favoritos.Show

End Sub

Private Sub M_Internet_Click()
' /******************************************************/
' Muestra el browser de Internet
' Primero buscamos en la tabla INI el PATH del browser
' especificado
' por el usuario, despues llamamos al browser
' /******************************************************/

On Error GoTo Etiqueta_Error:
Dim L_Registro As Recordset
Dim L_PATH$, L_Retorno

Set L_Registro = GV_Base_De_Datos.OpenRecordset("Ini")
If Not L_Registro.EOF Then
  If IsNull(L_Registro!Browser) Then
   L_PATH = ""
  Else
    L_PATH = L_Registro!Browser
  End If
End If
L_Registro.Close
If Trim(L_PATH) = "" Then
  MG_Mensaje "Debe especificar el PATH del Browser"
  Exit Sub
End If
L_Retorno = Shell(L_PATH, 1)
Exit Sub

Etiqueta_Error:
ME_Muestra_Error


End Sub

Private Sub M_Pag_Internet_Click()
' /*****************************************************/
' Abre una pagina de Web, la cual esta definida en la
' variable global  GV_Nombre
' /*****************************************************/

On Error GoTo Etiqueta_Error:
Dim L_Registro As Recordset
Dim L_PATH$, L_Retorno

Set L_Registro = GV_Base_De_Datos.OpenRecordset("Ini")
If Not L_Registro.EOF Then
  If IsNull(L_Registro!Browser) Then
   L_PATH = ""
  Else
    L_PATH = L_Registro!Browser
  End If
End If
L_Registro.Close
If Trim(L_PATH) = "" Then
  MG_Mensaje "Debe especificar el PATH del Browser"
  Exit Sub
End If
L_PATH = L_PATH + " " + GV_Nombre
L_Retorno = Shell(L_PATH, 1)
Exit Sub

Etiqueta_Error:
ME_Muestra_Error
  
End Sub

Private Sub M_PATH_Click()
OF_Registrar_Browser.Show
End Sub

Private Sub M_Retransmisión_Click()
' Muestra la ventana de Retransmisión de Mensajes
OF_Retransmision.Show
End Sub

Private Sub M_Salir_Click()
Unload Me
End Sub

Private Sub M_Servidores_Click()
' Muestra la ventana de Mantenimiento de Servidores
OF_Servidores.Show
End Sub


Private Sub M_Tipos_Click()
' Muestra la ventana de Tipos de Servidores
OF_Tipos_Servidores.Show
End Sub

Private Sub M_Usuario_Click()
' Muestra la ventana de Datos de Usuario
OF_Usuarios.Show
End Sub

Private Sub m_vent_Click(Index As Integer)
' /***************************************************/
' Este procedimiento arregla las ventanas del MDI de
' acuerdo a la opción del menu seleccionada
' /***************************************************/
Select Case Index
  Case 0
      OF_principal.Arrange 0
  Case 1
      OF_principal.Arrange 2
  Case 3
      OF_principal.Arrange 1
  Case 4
      OF_principal.Arrange 3
  
  Case 6
      MV_Minimizar_Ventanas "TODAS", 0, 1
  Case 7
      MV_Minimizar_Ventanas "TODAS", 0, 0
  
  Case 8
     MV_Minimizar_Ventanas "TODAS", 0, 3
    
      
End Select
End Sub

Private Sub M_Ventanas_Click()
 DoEvents
End Sub

Private Sub MDIForm_Load()


Dim L_Result As Integer
Dim L_HostName As String * 25

' Inicializa la variable del Explorador como Falso
GV_Explorador = False
GV_DoEVENTS = False

'Inicializacion de la Base de Datos
App.TaskVisible = False
App.HelpFile = App.Path & "\ERG2.HLP"


L_Result = MD_Abrir_Base_Datos(App.Path + "\ERG.mdb", _
            GV_Base_De_Datos)
If L_Result = 3024 Then
   CDialog.Filter = "Base de Datos (*.mdb)| *.mdb"
   CDialog.Action = 1
   If CDialog.FileName = "" Then
     MG_Mensaje _
     "La aplicación no podra iniciarse sin " + _
     "la Base de Datos ..."
     End
   Else
     If MD_Abrir_Base_Datos(CDialog.FileName, _
      GV_Base_De_Datos) <> 0 Then
       MG_Mensaje _
       "La aplicación no podra iniciarse sin " + _
       "la Base de Datos.. "
       End
     Else
       ' MG_Mensaje "< Base de Datos Inicializada >"
     End If
   End If
Else
  If L_Result <> 0 Then
      End
  Else
   'MG_Mensaje "< Base de Datos Inicializada >"
    
  End If
 End If
 

' /********************************/
' Comentado para no registrar mas a log

'L_Result = MD_Abrir_Base_Datos( _
            App.Path + "\LOG.mdb", GV_LOG)



Me.Caption = GC_Sistema
MG_Barra GC_Sistema + " Por : " + GC_Autores



' Inicialización de Variables
' Fin del Linea y Nueva Linea y los diferentes
' colores utilizados
 
GV_Verde = RGB(31, 155, 50)
GV_Azul = RGB(35, 61, 170)
GV_Rojo = vbRed
GV_RojoAlgo = RGB(72, 59, 57)
GV_Morado = RGB(90, 59, 170)
GV_Negro = vbBlack
GV_EOD = Chr(13) & Chr(10)

' Inicializa el Arreglo de Sockets, todos como sockets
' Invalidos

   GV_Sockets(1).socket = INVALID_SOCKET
   GV_Sockets(2).socket = INVALID_SOCKET
   GV_Sockets(3).socket = INVALID_SOCKET
   GV_Sockets(4).socket = INVALID_SOCKET
   GV_Sockets(5).socket = INVALID_SOCKET


' Inicialmente todas las casillas del arreglo de sockets
' estan libres

    GV_Sockets(1).Libre = True
    GV_Sockets(2).Libre = True
    GV_Sockets(3).Libre = True
    GV_Sockets(4).Libre = True
    GV_Sockets(5).Libre = True
    
' Inicializar el arreglo de puertos para enviar archivos
    GV_Puertos(1).E_Libre = True
    GV_Puertos(1).E_Puerto = 1040
    
    GV_Puertos(2).E_Libre = True
    GV_Puertos(2).E_Puerto = 1041
    
    GV_Puertos(3).E_Libre = True
    GV_Puertos(3).E_Puerto = 1042
    
    GV_Puertos(4).E_Libre = True
    GV_Puertos(4).E_Puerto = 1043
    
    GV_Puertos(5).E_Libre = True
    GV_Puertos(5).E_Puerto = 1044

' Bandera que Nos indica que la interfase con Winsock
' todavia no se ha  efectuado
  
GV_WINSOCK_ARRIBA = False


' Inicio
' Procedimiento Encargado de Inicializar  la
' interfase con Winsock

    
MM_Inicializa_WINSOCK
If Not GV_WINSOCK_ARRIBA Then
  MG_Mensaje "< No se pudo realizar la conexión de Winsock >"
  
  Exit Sub
Else
  L_Result = GetHostName(L_HostName, 25)
  If L_Result = SOCKET_ERROR Then
    L_HostName = "default" + Chr(0)
  End If
  L_Result = InStr(1, L_HostName, Chr(0))
  L_HostName = Trim(Left(L_HostName, L_Result - 1))
  If Trim(L_HostName) <> "" Then
      MD_Registra_Hostname Trim(L_HostName)
    End If
  
  

End If

' Crea la ventana de Usuario principal
' Solo para definir el arreglo con una dimensión
' En este caso de dimensión 1
' Lo mismo haremos con los otros arreglos de Ventanas
    
' Puesto que el arreglo es dinámico, lo redimensionamos

 ReDim GV_VENTANAS_Usuario(1)
 ReDim GV_Estado_Usuario(1)
 ' Inicializa la Ventana como borrada
 GV_Estado_Usuario(1).Deleted = True
    
 ' Puesto que el arreglo es dinámico, lo redimensionamos
 ReDim GV_VENTANAS_Canal(1)
 ReDim GV_Estado_Canal(1)
 
 ReDim GV_VENTANAS_Lista_Canales(1)
 ReDim GV_Estado_Lista_Canales(1)
 GV_Estado_Lista_Canales(1).Deleted = True
 ' Inicializa la Ventana como borrada
 GV_Estado_Canal(1).Deleted = True
    
' Crea la ventana de Estatus principal
   ReDim GV_VENTANAS_Estatus(1)
   ReDim GV_Estado_Estatus(1)
   ' Inicializa la Ventana como borrada
   GV_Estado_Estatus(1).Deleted = True
   
   
   
   ReDim GV_Estructura_Retransmision(1)
   GV_Estructura_Retransmision(1).E_Borrado = True
   
 
End Sub

Private Sub MDIForm_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
' /**************************************************/
' Si se hace click con el botón derecho en el MDI, se
' procede a mostrar el menu M_Ventanas, el cual es
' utilizado para arreglar las ventanas de
' la aplicación de la forma que el usuario desee.
' /***************************************************/
If Button = 2 Then
   DoEvents
   PopupMenu M_Ventanas
End If
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
' /**************************************************/
'   Cerrar todas las conexiones pendientes
'   y cerrar la interfase con el API de WINSOCK
' /**************************************************/

 Dim L_Texto$, i%

For i = 1 To GC_MAXSOCKETS
 If GV_Sockets(1).socket <> INVALID_SOCKET Then
  If closesocket(GV_Sockets(i).socket) = SOCKET_ERROR Then
  End If
 End If
Next i
 
 
' Cierra la interfase con Winsock
MM_Cierra_WINSOCK
 
' Cierra la Base de Datos

MD_Cerrar_Base_Datos GV_Base_De_Datos



' Cerrar la Base de Datos de LOG
' MD_Cerrar_Base_Datos GV_LOG

 
End 'Termina el programa
End Sub

Private Sub OTB_Toolbar_ButtonClick(ByVal Button As Button)
'/*******************************************************/
' Este es el código utilizado para determinar que opción
' del toolbar del MDI
' fue seleccionada
'/*******************************************************/

Select Case Button.Index
Case 1 ' Datos de Usuario
  M_Usuario_Click
Case 2 ' Datos de Servidores
  M_Servidores_Click
Case 3 ' Conectarse al ultimo servidor
  M_Conexion_Click
Case 4 ' Desconectarse de uno o varios servidores
  M_Desconexion_Click
Case 5 ' Mantenimiento de Comandos
  M_Comandos_Click
Case 6 ' Mantenimiento de Tipos de Servidores
  M_Tipos_Click
Case 7 ' Servidores Favoritos
  M_F_Servidores_Click
Case 8 ' Canales Favoritos
  M_F_Canales_Click
Case 9 ' browser de Internet
  M_Internet_Click
Case 10 ' Enviar Archivos
  M_Enviar_Click
Case 11 ' Explorador de ERG2
  M_Explorador_Click
Case 12 ' Ventana de Retransmisión de Mensajes
  M_Retransmisión_Click
Case 13 ' Salir de la Aplicación
  M_Salir_Click
Case Else
   ' nada
End Select
End Sub


