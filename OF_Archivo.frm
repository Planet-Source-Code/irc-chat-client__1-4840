VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.1#0"; "comctl32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "comdlg32.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form OF_Archivo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Aceptar Archivo"
   ClientHeight    =   5850
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7395
   HelpContextID   =   24
   Icon            =   "OF_Archivo.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   5850
   ScaleWidth      =   7395
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CommandButton OB_Salir 
      Caption         =   "&Salir"
      Height          =   375
      HelpContextID   =   24
      Left            =   6255
      TabIndex        =   19
      Top             =   5355
      Width           =   1035
   End
   Begin VB.TextBox OT_Mensajes 
      Height          =   1440
      HelpContextID   =   24
      Left            =   195
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   16
      Top             =   3765
      Width           =   7020
   End
   Begin VB.Frame OM_Progreso 
      Caption         =   "Progreso"
      Height          =   915
      Left            =   195
      TabIndex        =   13
      Top             =   2550
      Width           =   7020
      Begin ComctlLib.ProgressBar ProgressBar 
         Height          =   330
         Left            =   150
         TabIndex        =   14
         Top             =   375
         Width           =   6690
         _ExtentX        =   11800
         _ExtentY        =   582
         _Version        =   327680
         Appearance      =   1
      End
   End
   Begin MSComDlg.CommonDialog Dialog 
      Left            =   1635
      Top             =   5325
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327680
   End
   Begin VB.Frame OM_Datos 
      Caption         =   "Recibir Archivo"
      Height          =   2235
      HelpContextID   =   24
      Left            =   165
      TabIndex        =   0
      Top             =   225
      Width           =   7020
      Begin VB.TextBox OT_Size 
         Enabled         =   0   'False
         Height          =   300
         HelpContextID   =   24
         Left            =   1770
         TabIndex        =   9
         Top             =   1545
         Width           =   900
      End
      Begin VB.TextBox OT_Puerto 
         Height          =   375
         HelpContextID   =   24
         Left            =   4170
         TabIndex        =   11
         Top             =   1560
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.TextBox OT_IP 
         Height          =   375
         HelpContextID   =   24
         Left            =   5430
         TabIndex        =   12
         Top             =   1575
         Visible         =   0   'False
         Width           =   1170
      End
      Begin VB.CommandButton OB_Directorio 
         Caption         =   "..."
         Height          =   300
         HelpContextID   =   24
         Left            =   6465
         TabIndex        =   7
         Top             =   1050
         Width           =   285
      End
      Begin VB.TextBox OT_Nombre 
         Enabled         =   0   'False
         Height          =   300
         HelpContextID   =   24
         Left            =   1770
         TabIndex        =   2
         Top             =   540
         Width           =   1680
      End
      Begin VB.TextBox OT_SALVAR 
         Height          =   300
         HelpContextID   =   24
         Left            =   1770
         TabIndex        =   6
         Top             =   1035
         Width           =   4695
      End
      Begin VB.TextBox OT_Nick 
         Enabled         =   0   'False
         Height          =   300
         HelpContextID   =   24
         Left            =   5190
         TabIndex        =   4
         Top             =   555
         Width           =   1290
      End
      Begin VB.Label OE_Bytes 
         AutoSize        =   -1  'True
         Caption         =   "Bytes"
         Height          =   195
         Left            =   2835
         TabIndex        =   10
         Top             =   1590
         Width           =   390
      End
      Begin VB.Label OE_Tamano 
         AutoSize        =   -1  'True
         Caption         =   "Tamaño del Archivo"
         Height          =   195
         Left            =   195
         TabIndex        =   8
         Top             =   1605
         Width           =   1425
      End
      Begin VB.Label OE_Usuario 
         AutoSize        =   -1  'True
         Caption         =   "Usuario que Envia"
         Height          =   195
         Left            =   3705
         TabIndex        =   3
         Top             =   585
         Width           =   1305
      End
      Begin VB.Label OE_SalvarComo 
         AutoSize        =   -1  'True
         Caption         =   "Salvar Archivo Como"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   1110
         Width           =   1485
      End
      Begin VB.Label OE_NombreArchivo 
         AutoSize        =   -1  'True
         Caption         =   "Nombre del Archivo"
         Height          =   195
         Left            =   135
         TabIndex        =   1
         Top             =   540
         Width           =   1395
      End
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   1185
      Top             =   5340
      _ExtentX        =   741
      _ExtentY        =   741
   End
   Begin VB.CommandButton OB_Cancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      HelpContextID   =   24
      Left            =   5145
      TabIndex        =   18
      Top             =   5355
      Width           =   1035
   End
   Begin VB.CommandButton OB_Aceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      HelpContextID   =   24
      Left            =   4035
      TabIndex        =   17
      Top             =   5355
      Width           =   1035
   End
   Begin VB.Label OE_Mensajes 
      AutoSize        =   -1  'True
      Caption         =   "Mensajes"
      Height          =   195
      Left            =   225
      TabIndex        =   15
      Top             =   3525
      Width           =   675
   End
End
Attribute VB_Name = "OF_Archivo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' LF_Contador lleva un registro de los bytes recibidos del archivo
' que aceptamos
Dim LF_Contador As Long
' LF_Filenumber es el identificador del archivo donde salvaremos
' el que estamos Aceptando
Dim LF_FileNumber

Private Sub Form_Load()
'/*******************************************************/
' Al cargarse la forma se deshabilita el botón de Salir,
' para activarse
' primero se debe hacer click en el botón de Cancelar
'/*******************************************************/

OB_Salir.Enabled = False
Beep
End Sub

Private Sub OB_Aceptar_Click()
'/***************************************************************/
' Cuando se hace Click en el botón de Aceptar, primero se
' deshabilita el botón de Aceptar, seguidamente a la variable
' local a la forma  LF_FileNumber le asignamos un numero de
' archivo libre (Visual Basic maneja los archivos de este modo)
' , esto se realiza  con la función de Visual Basic [FreeFile].
' Una vez obtenido un número de archivo libre, procedemos a abrir
' el archivo en formato binario, el TextBox OT_Salvar, contiene el
' PATH y nombre del archivo  donde se grabará el archivo que
' recibiremos.
' La variable local a la forma LF_Contador es inicializada en Cero
' (Cero bytes Recibidos).
' Seguidamente utilizamos el control de Winsock que provee
' Visual Basic 5.0. Para conectarnos a la dirección IP y el puerto
' que nos envió el usuario
' del cual recibiremos un archivo
'/****************************************************************/

OB_Aceptar.Enabled = False
OM_Datos.Enabled = False
LF_FileNumber = FreeFile
'Crear archivo binario
Open OT_SALVAR For Binary As #LF_FileNumber
LF_Contador = 0

Winsock1.Connect OT_IP, OT_Puerto ' Conectarse
End Sub

Private Sub OB_Cancelar_Click()
'/****************************************************/
' Si el usuario presiona el botón de Cancelar, entonces
' se cierra el socket, esto se realiza con el control
' que provee Visual Basic 5.0 para el manejo de Sockets.
' A la vez se habilita el botón de Salir y se
' deshabilita el botón de Cancelar
'/****************************************************/
On Error GoTo Etiqueta_Error:
Winsock1.Close ' Cerrar el Socket
Close #LF_FileNumber
OB_Salir.Enabled = True
OB_Cancelar.Enabled = False
Exit Sub

Etiqueta_Error:


End Sub

Private Sub OB_Directorio_Click()
'/*****************************************************/
' Si se presiona el botón que tiene los tres puntos,
' entonces se especifica el nombre del archivo del
' Dialogo y se abre el dialogo para salvar un Archivo.
' El Dialogo de Archivos
' es un objeto propio de Visual Basic.
' El objeto se llama Dialog, una vez que se selecciona el
' nombre y el directorio donde se guardará el archivo que
' se  recibirá este es asignado al TEXTBOX OT_SALVAR
'/****************************************************/

Dialog.filename = OT_Nombre
Dialog.ShowSave ' Mostrar la pantalla de Dialogo

OT_SALVAR = Dialog.filename 'Dialog.Filename posee el
                        ' nombre del archivo a Guardar

End Sub

Private Sub OB_Salir_Click()
Unload Me ' Descargar la forma
End Sub

Private Sub Winsock1_Close()
'/********************************************************/
' Evento CLOSE en el Socket
'/********************************************************/
On Error GoTo Etiqueta_Error:
Close #LF_FileNumber   ' Cerrar el archivo
    
Exit Sub
Etiqueta_Error:
ME_Muestra_Error
End Sub

Private Sub Winsock1_Connect()
'/*********************************************************/
' Evento Connect en el Socket
' Mostrar el Mensaje de conexión
' Enviar al otro socket al que se conectó un mensaje que le
' indica que inicie la transmisión del Archivo
'/*********************************************************/

 OT_Mensajes = "Conectado al otro Usuario " + GV_EOD
 
 Winsock1.SendData "OK_INICIE" ' Mensaje de Inicio
  
 OT_Mensajes = "Empezando Transmisión" + GV_EOD
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
'/**********************************************************/
' Evento READ en el Socket, nos indica que nos acaba de
' llegar datos de Tamaño = [BytesTotal]
' Cada vez que nos llegan datos, redimensionamos el arreglo
' de bytes de acuerdo a la cantidad de bytes que nos llego.
' Con la función GetDATA del Control de Winsock, recuperamos
' los datos.
' Verificamos el contador de Bytes recibidos para que nos
' indique si ya recibimos todo el archivo, seguidamente se
' redibuja la barra de progreso y grabamos al archivo los
' datos que nos acaban de llegar.
'/************************************************************/
On Error GoTo Etiqueta_Error:
Dim L_Data() As Byte

' Arreglo de Bytes para recibir los datos
ReDim L_Data(bytesTotal) As Byte

' Obtener los datos
Winsock1.GetData L_Data, , bytesTotal

' Quitarle el ultimo byte el cual trae el CR-LF
ReDim Preserve L_Data(bytesTotal - 1)

If LF_Contador = CLng(OT_Size) Then _
  ProgressBar.Value = 100: Exit Sub

LF_Contador = LF_Contador + bytesTotal - 1
If LF_Contador > CLng(OT_Size) Then
    LF_Contador = LF_Contador = CLng(OT_Size)
End If

' Si se obtuvo los bytes esperados
If LF_Contador = CLng(OT_Size) Then
    Put #LF_FileNumber, , L_Data ' Grabar en el Archivo.
    Close #LF_FileNumber ' Cerrar el archivo si ya se llego
                         ' a la cantidad de byrtes esperados
                         
    ProgressBar.Value = Int((LF_Contador * 100) / OT_Size)
    Beep ' Sonido de Terminación
    OT_Mensajes = _
       "Transmition Completed/Transmisión Terminada" + GV_EOD
    ' Notificar Fin al otro extremo
    Winsock1.SendData "FIN TRANSMISION"
    Winsock1.Close ' Cerrar el Socket
    
' Si todavia nos faltan bytes ...
ElseIf LF_Contador < CLng(OT_Size) Then
    Put #LF_FileNumber, , L_Data ' Grabar en el Archivo
    ProgressBar.Value = Int((LF_Contador * 100) / OT_Size)
    ' Indicador de continuación de Transmisión
    Winsock1.SendData "OK_INICIE"
               
End If

Exit Sub
Etiqueta_Error:
ME_Muestra_Error
End Sub

Private Sub Winsock1_Error(ByVal number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
'/**********************************************************/
' Si hubo algún error en el socket entonces mostrar el error
' y cerrar el socket
'/*********************************************************/
MG_Mensaje CStr(number) + " " + Description
CancelDisplay = True
Winsock1.Close

End Sub
