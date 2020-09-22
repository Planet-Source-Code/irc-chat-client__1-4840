VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.1#0"; "comctl32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "comdlg32.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form OF_Enviar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Enviar Archivo"
   ClientHeight    =   5895
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7440
   HelpContextID   =   23
   Icon            =   "OF_Enviar.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   5895
   ScaleWidth      =   7440
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CommandButton OB_Salir 
      Caption         =   "&Salir"
      Height          =   375
      HelpContextID   =   23
      Left            =   6195
      TabIndex        =   18
      Top             =   5415
      Width           =   1035
   End
   Begin VB.TextBox OT_Mensajes 
      Height          =   1440
      HelpContextID   =   23
      Left            =   225
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   16
      Top             =   3795
      Width           =   7020
   End
   Begin VB.TextBox OT_SocketAsociado 
      Height          =   375
      Left            =   2055
      TabIndex        =   15
      Text            =   "1"
      Top             =   5565
      Visible         =   0   'False
      Width           =   795
   End
   Begin MSWinsockLib.Winsock Winsock2 
      Left            =   1335
      Top             =   5565
      _ExtentX        =   741
      _ExtentY        =   741
   End
   Begin VB.Frame OM_Progreso 
      Caption         =   "Progreso"
      Height          =   915
      HelpContextID   =   23
      Left            =   180
      TabIndex        =   5
      Top             =   2565
      Width           =   7020
      Begin ComctlLib.ProgressBar ProgressBar 
         Height          =   330
         Left            =   135
         TabIndex        =   6
         Top             =   375
         Width           =   6675
         _ExtentX        =   11774
         _ExtentY        =   582
         _Version        =   327680
         Appearance      =   1
         MouseIcon       =   "OF_Enviar.frx":0442
      End
   End
   Begin MSComDlg.CommonDialog Dialog 
      Left            =   495
      Top             =   5535
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327680
   End
   Begin VB.Frame OM_Enviar 
      Caption         =   "Enviar Archivo"
      Height          =   2235
      HelpContextID   =   23
      Left            =   165
      TabIndex        =   2
      Top             =   225
      Width           =   7020
      Begin VB.ComboBox OC_Servidor 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         HelpContextID   =   23
         Left            =   1770
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   345
         Width           =   5070
      End
      Begin VB.TextBox OT_Size 
         Enabled         =   0   'False
         Height          =   300
         HelpContextID   =   23
         Left            =   1770
         TabIndex        =   12
         Top             =   1710
         Width           =   900
      End
      Begin VB.TextBox OT_Puerto 
         Height          =   375
         Left            =   4320
         TabIndex        =   11
         Top             =   2100
         Visible         =   0   'False
         Width           =   1230
      End
      Begin VB.TextBox OT_IP 
         Height          =   375
         Left            =   5565
         TabIndex        =   10
         Top             =   2085
         Visible         =   0   'False
         Width           =   1350
      End
      Begin VB.CommandButton OB_Directorio 
         Caption         =   "..."
         Height          =   300
         HelpContextID   =   23
         Left            =   6495
         TabIndex        =   8
         Top             =   1245
         Width           =   285
      End
      Begin VB.TextBox OT_SALVAR 
         Height          =   300
         HelpContextID   =   23
         Left            =   1770
         TabIndex        =   4
         Top             =   1245
         Width           =   4695
      End
      Begin VB.TextBox OT_Nick 
         Height          =   300
         HelpContextID   =   23
         Left            =   1770
         TabIndex        =   3
         Top             =   855
         Width           =   1290
      End
      Begin VB.Label OE_Servidor 
         AutoSize        =   -1  'True
         Caption         =   "Servidor "
         Height          =   195
         Left            =   975
         TabIndex        =   20
         Top             =   420
         Width           =   630
      End
      Begin VB.Label OE_Bytes 
         AutoSize        =   -1  'True
         Caption         =   "Bytes"
         Height          =   195
         Left            =   2820
         TabIndex        =   14
         Top             =   1815
         Width           =   390
      End
      Begin VB.Label OE_Tamano 
         AutoSize        =   -1  'True
         Caption         =   "Tamaño del Archivo"
         Height          =   195
         Left            =   180
         TabIndex        =   13
         Top             =   1815
         Width           =   1425
      End
      Begin VB.Label OE_Alias 
         AutoSize        =   -1  'True
         Caption         =   "Enviar a :"
         Height          =   195
         Left            =   930
         TabIndex        =   9
         Top             =   870
         Width           =   675
      End
      Begin VB.Label OE_Archivo 
         AutoSize        =   -1  'True
         Caption         =   "Archivo a Enviar"
         Height          =   195
         Left            =   435
         TabIndex        =   7
         Top             =   1275
         Width           =   1170
      End
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   45
      Top             =   5535
      _ExtentX        =   741
      _ExtentY        =   741
   End
   Begin VB.CommandButton OB_Cancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      HelpContextID   =   23
      Left            =   5085
      TabIndex        =   1
      Top             =   5415
      Width           =   1035
   End
   Begin VB.CommandButton OB_Aceptar 
      Caption         =   "&Enviar"
      Height          =   375
      HelpContextID   =   23
      Left            =   3915
      TabIndex        =   0
      Top             =   5415
      Width           =   1035
   End
   Begin VB.Label OE_Mensajes 
      AutoSize        =   -1  'True
      Caption         =   "Mensajes"
      Height          =   195
      Left            =   285
      TabIndex        =   17
      Top             =   3570
      Width           =   675
   End
End
Attribute VB_Name = "OF_Enviar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' LF_Contador lleva un registro de los bytes recibidos del
' archivo que aceptamos LF_NUM es utilizado para
' referenciar en el arreglo de puertos el puerto utilizado
' para enviar archivos LF_Filenumber es el identificador
' del archivo donde salvaremos el que estamos
' Aceptando

Dim LF_Contador&
Dim LF_NUM%
Dim LF_FileNumber
    
Private Sub Form_Load()
'/******************************************************/
' Habilitar y deshabilitar los botones correspondientes
' y se procede a cargar los servidores activos
'/******************************************************/
OB_Aceptar.Enabled = True
OB_Cancelar.Enabled = False
OB_salir.Enabled = True


PL_Cargar_Servidores_Activos
End Sub

Sub PL_Cargar_Servidores_Activos()
' /******************************************************/
' Este procedimiento carga todos los servidores activos
' en el momento. Los servidores son cargados, para que el
' usuario tenga opción a enviar a cualquiera de los
' diferentes servidores activos, a un usuario especifico
' Los servidores activos son cargados del arreglo global
' de Sockets
' /******************************************************/

Dim L_i As Integer
Dim L_libre As Integer

L_libre = 1

For L_i = 1 To 5
  If GV_Sockets(L_i).socket <> INVALID_SOCKET Then
     OC_Servidor.AddItem GV_Sockets(L_i).Direcc + "(" + _
     CStr(GV_Sockets(L_i).Puerto) + _
     ") NICK => " + GV_Sockets(L_i).Nick
     OC_Servidor.ItemData(L_libre - 1) = L_i
     
     L_libre = L_libre + 1
  
  End If

Next L_i
If OC_Servidor.ListCount > 0 Then _
OC_Servidor.ListIndex = 0: OT_SocketAsociado = _
OC_Servidor.ItemData(OC_Servidor.ListIndex)
End Sub


Private Sub Form_Unload(Cancel As Integer)
' Si no esta habilitado el botón de Salir
If Not OB_salir.Enabled Then Cancel = 1
End Sub

Private Sub Label1_Click()

End Sub

Private Sub OB_Aceptar_Click()
'/******************************************************/
' En el Click, de este botón hace la confirmación para
' enviar un archivo Primero validar que se haya
' seleccionado un servidor, luego se valida que se
' haya especificado el usuario al que se le desee enviar
' el archivo, luego se valida que se haya seleccionado
' el archivo a enviar, una vez validados estos
' Puntos se asigna un número de archivo libre para enviar,
' luego se busca un puerto libre por donde se pueda
' enviar el archivo
'/*******************************************************/
On Error GoTo Etiqueta_Error:
Dim L_Num%
Dim L_Archivo$

If OT_SocketAsociado = 0 Then
  MG_Mensaje _
  "No se ha especificado al servidor Activo " + _
    "del Usuario al que se desea enviar el Archivo"
  Exit Sub
End If

If Trim(OT_Nick) = "" Then
   MG_Mensaje "No se ha especificado el " + _
   "Usuario al que se desea enviar el Archivo"
  Exit Sub
End If

If Trim(OT_SALVAR) = "" Then
  MG_Mensaje "Debe especificar el Archivo a enviar..."
  Exit Sub
End If

OB_Aceptar.Enabled = False
OB_Cancelar.Enabled = True
OB_salir.Enabled = False

LF_FileNumber = FreeFile
If Trim(OT_SALVAR) <> "" Then
    'Abrir archivo
    Open OT_SALVAR For Binary As #LF_FileNumber
      
    LF_NUM = MM_Obtener_Puerto_Libre()
    If LF_NUM = 0 Then _
     MG_Mensaje _
      "No existen mas puertos para enviar archivos": _
      Exit Sub
    
    OT_Puerto = GV_Puertos(LF_NUM).E_Puerto
    
    ' Obtiene el archivo sin el PATH
    L_Archivo = MV_Obtener_Archivo(OT_SALVAR)
   ' el mensaje que se envia tiene el siguiente formato
   ' Primero el archivo luego la dirección
   ' IP luego el puerto , luego el tamño del
   ' archivo
   ' PRIVMSG rav: DCC SEND archivo.txt 130.111.111.111 1040 400
    ' Si no se pudo enviar el archivo entonces cerrar el archivo
    If Not MM_Enviar_Mensaje("PRIVMSG " + _
    Trim(OT_Nick) + " : " + _
    "" + "DCC SEND " + _
    L_Archivo + " " + Winsock1.LocalIP _
    + " " + OT_Puerto + " " _
    + CStr(OT_Size) + "", OT_SocketAsociado) Then
       Close #LF_FileNumber ' Cerrar el archivo
       
       Exit Sub
    End If
    ' Asignar el puerto libre al objeto de Winsock
    Winsock1.LocalPort = OT_Puerto
    ' Pone a escuhar al socket
    Winsock1.listen
    OT_Mensajes = _
    "Waiting for Response/Esperando Contestación ....." _
    + GV_EOD
End If
Exit Sub
Etiqueta_Error:
ME_Muestra_Error
End Sub

Private Sub OB_Cancelar_Click()
' /******************************************************/
' En el botón de Cancelar primero se debe poner en libre
' el puerto usado luego se cierran los sockets
' /******************************************************/
On Error GoTo Etiqueta_Error:
If LF_NUM <> 0 Then
    GV_Puertos(LF_NUM).E_Libre = True
End If

If LF_Contador < CLng(OT_Size) Then
  Close #LF_FileNumber
  Winsock1.Close
  Winsock2.Close
  
  OT_Mensajes = "Transmisión Interrumpida ..." + GV_EOD
  
End If
OB_Aceptar.Enabled = False
OB_Cancelar.Enabled = False
OB_salir.Enabled = True

Exit Sub
Etiqueta_Error:
ME_Muestra_Error

End Sub

Private Sub OB_Directorio_Click()
Dialog.filename = ""
Dialog.ShowOpen

OT_SALVAR = Dialog.filename
If Trim(OT_SALVAR) <> "" Then OT_Size = FileLen(OT_SALVAR)

End Sub

Private Sub OB_Salir_Click()
' Descargar la forma
Unload Me
End Sub

Private Sub OC_Servidor_Click()
' Asignar el socket del servidor seleccionado
OT_SocketAsociado = OC_Servidor.ItemData(OC_Servidor.ListIndex)
End Sub

Private Sub OT_SALVAR_KeyPress(KeyAscii As Integer)
' Deshabilitar para no escribir
KeyAscii = 0
End Sub

Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)
' /*****************************************************/
' Este evento de Winsock es activado cuando nos llega
' la confirmación del usuario al cual le estamos enviando
' el archivo  El otro usuario realizó un CONNECT la
' dirección IP y Puerto que enviamos
' /*****************************************************/

   OT_Mensajes = _
   "Usuario ha aceptado la Transferencia del Archivo...." _
   + GV_EOD
   Winsock2.accept requestID ' Aceptar la requisición
   LF_Contador = 0
   OT_Mensajes = "Iniciando Transmisión...." + GV_EOD
End Sub

Private Sub Winsock1_Error(ByVal number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
'/****************************************************/
' Si hubo algún error en el socket entonces mostrar
' el error  y cerrar el socket
'/****************************************************/

CancelDisplay = True
Winsock1.Close


End Sub

Private Sub Winsock2_Close()
' Cerrar el socket
Close #LF_FileNumber
End Sub

Private Sub Winsock2_DataArrival(ByVal bytesTotal As Long)
'/******************************************************/
' Evento READ en el Socket, nos indica que nos acaba de
' llegar datos  de Tamaño = [BytesTotal]
' El socket recibe solo dos mensajes :
' OK_INICIE, FIN_TRANSMISION
' cada vez que se recibe OK_INICIE, se leen los
' siguientes 1024 bytes
' del archivo para ser enviados
'/*******************************************************/

On Error GoTo Etiqueta_Error:
Dim L_Data$
Dim L_Mychar() As Byte
Dim L_cont&

Winsock2.GetData L_Data ' Recibe los datos del otro extremo
If Trim(L_Data) = "OK_INICIE" Then ' Si todavia es OK_INICIE
      
     If LF_Contador = _
      CLng(OT_Size) Then ProgressBar.Value = 100: Exit Sub
     LF_Contador = LF_Contador + 1024
     If LF_Contador > CLng(OT_Size) Then
          L_cont = 1024 - (LF_Contador - CLng(OT_Size))
          LF_Contador = CLng(OT_Size)
          ReDim L_Mychar(L_cont) As Byte
          Get #LF_FileNumber, , L_Mychar
          ProgressBar.Value = Int((LF_Contador * 100) / OT_Size)
          OT_Mensajes = "Transmisión Terminada...." + GV_EOD
          Winsock2.SendData L_Mychar
          OB_Aceptar.Enabled = False
          OB_Cancelar.Enabled = False
          OB_salir.Enabled = True

          
          Beep
          
  
     ElseIf LF_Contador = CLng(OT_Size) Then
         ReDim L_Mychar(1024) As Byte
         Get #LF_FileNumber, , L_Mychar
         ProgressBar.Value = Int((LF_Contador * 100) / OT_Size)
         OT_Mensajes = "Transmisión Terminada...." + GV_EOD
         Winsock2.SendData L_Mychar
         OB_Aceptar.Enabled = False
         OB_Cancelar.Enabled = False
         OB_salir.Enabled = True
         
         Beep
       
     ElseIf LF_Contador < CLng(OT_Size) Then
         ReDim L_Mychar(1024) As Byte
         Get #LF_FileNumber, , L_Mychar
         ProgressBar.Value = Int((LF_Contador * 100) / OT_Size)
         Winsock2.SendData L_Mychar
               
     End If
     
   
   
ElseIf Trim(L_Data) = "FIN TRANSMISION" Then
   
   Winsock2.Close
   OB_Aceptar.Enabled = False
   OB_Cancelar.Enabled = False
   OB_salir.Enabled = True
   
End If

Exit Sub
Etiqueta_Error:
ME_Muestra_Error
End Sub

Private Sub Winsock2_Error(ByVal number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
'/******************************************************/
' Si hubo algún error en el socket entonces mostrar el
' error y cerrar el socket
'/******************************************************/

CancelDisplay = True
Winsock2.Close
Winsock1.Close

End Sub
