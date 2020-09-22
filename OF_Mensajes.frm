VERSION 5.00
Begin VB.Form OF_Mensajes 
   Caption         =   "Mensajes"
   ClientHeight    =   1920
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5835
   HelpContextID   =   1
   Icon            =   "OF_Mensajes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1920
   ScaleWidth      =   5835
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox OT_Mensaje 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1125
      HelpContextID   =   1
      Left            =   165
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Text            =   "OF_Mensajes.frx":0442
      Top             =   150
      Width           =   5475
   End
   Begin VB.CommandButton OB_OK 
      Caption         =   "&Aceptar"
      Height          =   345
      HelpContextID   =   1
      Left            =   2280
      TabIndex        =   0
      Top             =   1530
      Width           =   1275
   End
End
Attribute VB_Name = "OF_Mensajes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
'Esta forma es utilizada en lugar del msgbox para no
'interrumpir los eventos
'asíncronos de nuestra aplicación.
'Esta forma es cargada  modalmente.
End Sub

Private Sub OB_OK_Click()
' Descargar forma
 Unload Me
End Sub
