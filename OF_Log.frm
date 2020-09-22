VERSION 5.00
Begin VB.Form OF_Log 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registrar Errores y Sugerencias"
   ClientHeight    =   5850
   ClientLeft      =   1350
   ClientTop       =   645
   ClientWidth     =   6720
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5850
   ScaleWidth      =   6720
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame OM_Marco 
      Caption         =   "Sugerencias y errores"
      Height          =   4995
      HelpContextID   =   1
      Left            =   180
      TabIndex        =   0
      Top             =   225
      Width           =   6345
      Begin VB.TextBox OT_Persona 
         Height          =   315
         HelpContextID   =   1
         Left            =   210
         MaxLength       =   100
         TabIndex        =   2
         Top             =   675
         Width           =   5925
      End
      Begin VB.TextBox OT_Comentario 
         Height          =   3285
         HelpContextID   =   1
         Left            =   210
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   1530
         Width           =   5925
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Persona que registra sugerencia o error (este campo puede quedar vacio)"
         Height          =   195
         Left            =   210
         TabIndex        =   1
         Top             =   390
         Width           =   5205
      End
      Begin VB.Label Label1 
         Caption         =   "Sugerencia o error"
         Height          =   192
         Left            =   216
         TabIndex        =   3
         Top             =   1272
         Width           =   1404
      End
   End
   Begin VB.CommandButton OB_Salir 
      Caption         =   "&Salir"
      Height          =   390
      HelpContextID   =   1
      Left            =   5475
      TabIndex        =   6
      Top             =   5355
      Width           =   1035
   End
   Begin VB.CommandButton OB_Ok 
      Caption         =   "&Aceptar"
      Height          =   390
      HelpContextID   =   1
      Left            =   4335
      TabIndex        =   5
      Top             =   5370
      Width           =   1035
   End
End
Attribute VB_Name = "OF_Log"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Load()

OT_Persona = ""
OT_Comentario = ""

End Sub


Private Sub OB_OK_Click()
On Error GoTo Etiqueta_Error:
Dim L_Registro As Recordset
If Trim(OT_Comentario) = "" Then Exit Sub

Set L_Registro = GV_LOG.OpenRecordset("TIPS")
L_Registro.AddNew
L_Registro!Fecha = Date
L_Registro!Comentario = OT_Comentario + Chr(13)
L_Registro!Persona = OT_Persona + Chr(13)
L_Registro.Update
L_Registro.Close

OT_Persona = ""
OT_Comentario = ""

Exit Sub
Etiqueta_Error:
ME_Muestra_Error
End Sub

Private Sub OB_Salir_Click()
Unload Me
End Sub


