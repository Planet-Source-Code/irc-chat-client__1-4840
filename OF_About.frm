VERSION 5.00
Begin VB.Form OF_About 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Acerca de ERG2"
   ClientHeight    =   5220
   ClientLeft      =   1590
   ClientTop       =   825
   ClientWidth     =   7695
   HelpContextID   =   30
   Icon            =   "OF_About.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5220
   ScaleWidth      =   7695
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton OB_Ok 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Height          =   372
      HelpContextID   =   30
      Left            =   5955
      TabIndex        =   0
      Top             =   4560
      Width           =   1308
   End
   Begin VB.Frame OM_Marco 
      Height          =   5115
      HelpContextID   =   1
      Left            =   45
      TabIndex        =   1
      Top             =   15
      Width           =   7560
      Begin VB.Frame OM_Modo 
         Caption         =   "Modo de ERG2"
         Height          =   900
         HelpContextID   =   1
         Left            =   4905
         TabIndex        =   14
         Top             =   3180
         Visible         =   0   'False
         Width           =   2310
         Begin VB.OptionButton OO_Normal 
            Caption         =   "Normal"
            Height          =   195
            Left            =   195
            TabIndex        =   16
            Top             =   420
            Width           =   870
         End
         Begin VB.OptionButton OO_Exten 
            Caption         =   "Extendido"
            Height          =   195
            Left            =   1230
            TabIndex        =   15
            Top             =   435
            Width           =   1005
         End
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         X1              =   120
         X2              =   7410
         Y1              =   3015
         Y2              =   3000
      End
      Begin VB.Label OE_Ingeniería_En_Sistemas 
         AutoSize        =   -1  'True
         Caption         =   "Ingeniería en Sistemas Computacionales"
         Height          =   195
         Left            =   1620
         TabIndex        =   13
         Top             =   3615
         Width           =   2925
      End
      Begin VB.Label OE_Honduras 
         AutoSize        =   -1  'True
         Caption         =   "Honduras, C.A"
         Height          =   195
         Left            =   1455
         TabIndex        =   12
         Top             =   4680
         Width           =   1035
      End
      Begin VB.Label OE_Unitec 
         AutoSize        =   -1  'True
         Caption         =   "UNITEC"
         Height          =   195
         Left            =   1680
         TabIndex        =   11
         Top             =   4425
         Width           =   600
      End
      Begin VB.Label OE_Universidad 
         AutoSize        =   -1  'True
         Caption         =   "Universidad Tecnológica Centroamericana"
         Height          =   195
         Left            =   630
         TabIndex        =   10
         Top             =   4185
         Width           =   3090
      End
      Begin VB.Label OE_Titulo 
         AutoSize        =   -1  'True
         Caption         =   "optar al titulo de  : "
         Height          =   195
         Left            =   270
         TabIndex        =   9
         Top             =   3600
         Width           =   1275
      End
      Begin VB.Label OE_Graduación 
         AutoSize        =   -1  'True
         Caption         =   "Fue realizado como proyecto de graduación para "
         Height          =   195
         Left            =   270
         TabIndex        =   8
         Top             =   3360
         Width           =   3585
      End
      Begin VB.Label OE_Visual50 
         AutoSize        =   -1  'True
         Caption         =   "Este sistema fue elaborado en Visual Basic 5.0"
         Height          =   195
         Left            =   270
         TabIndex        =   7
         Top             =   3105
         Width           =   3300
      End
      Begin VB.Label OE_Gustavo 
         AutoSize        =   -1  'True
         Caption         =   "Gustavo Enrique Oviedo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3990
         TabIndex        =   6
         Top             =   1710
         Width           =   2580
      End
      Begin VB.Label OE_Autores 
         AutoSize        =   -1  'True
         Caption         =   "Autores :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3990
         TabIndex        =   5
         Top             =   975
         Width           =   1110
      End
      Begin VB.Label OE_Rogger 
         AutoSize        =   -1  'True
         Caption         =   "Rogger Alexis Vásquez"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3990
         TabIndex        =   4
         Top             =   1335
         Width           =   2475
      End
      Begin VB.Label OE_Derechos 
         AutoSize        =   -1  'True
         Caption         =   "Todos los Derechos Reservados. 1997"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3990
         TabIndex        =   3
         Top             =   2520
         Width           =   3315
      End
      Begin VB.Label OE_Version 
         AutoSize        =   -1  'True
         Caption         =   "Sistema ERG2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   3990
         TabIndex        =   2
         Top             =   315
         Width           =   2355
      End
      Begin VB.Image Image2 
         BorderStyle     =   1  'Fixed Single
         Height          =   2640
         Left            =   135
         Picture         =   "OF_About.frx":0442
         Stretch         =   -1  'True
         Top             =   240
         Width           =   3645
      End
   End
   Begin VB.Image Image1 
      Height          =   7470
      Left            =   0
      Picture         =   "OF_About.frx":17104
      Top             =   -15
      Width           =   10035
   End
End
Attribute VB_Name = "OF_About"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
Image1.Visible = False
If GV_DoEVENTS Then
  OO_Exten.Value = True
Else
 OO_Normal.Value = True
End If

End Sub

Private Sub Image1_DblClick()
Image1.Visible = False
OM_Marco.Visible = True

End Sub

Private Sub Image2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 And Shift = 1 Then OM_Modo.Visible = True
End Sub

Private Sub OB_OK_Click()
Unload Me
End Sub


Private Sub OM_Marco_DblClick()
Image1.Visible = True
OM_Marco.Visible = False
End Sub

Private Sub OO_Exten_Click()
GV_DoEVENTS = True
End Sub

Private Sub OO_Normal_Click()
GV_DoEVENTS = False
End Sub
