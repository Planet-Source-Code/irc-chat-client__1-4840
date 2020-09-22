VERSION 5.00
Begin VB.Form OF_Usuarios 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Usuario"
   ClientHeight    =   2865
   ClientLeft      =   2040
   ClientTop       =   1815
   ClientWidth     =   5730
   HelpContextID   =   2
   Icon            =   "OF_Usuar.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2865
   ScaleWidth      =   5730
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame OM_Marco 
      Caption         =   "Usuario"
      Height          =   2304
      HelpContextID   =   2
      Left            =   96
      TabIndex        =   0
      Top             =   30
      Width           =   5520
      Begin VB.TextBox OT_Host 
         Height          =   300
         HelpContextID   =   2
         Left            =   156
         TabIndex        =   2
         Top             =   480
         Width           =   2448
      End
      Begin VB.TextBox OT_Direccion 
         Height          =   285
         HelpContextID   =   2
         Left            =   2832
         TabIndex        =   4
         Top             =   480
         Width           =   2496
      End
      Begin VB.TextBox OT_Nombre 
         Height          =   300
         HelpContextID   =   2
         Left            =   156
         TabIndex        =   6
         Top             =   1140
         Width           =   2450
      End
      Begin VB.TextBox OT_Alias 
         Height          =   300
         HelpContextID   =   2
         Left            =   2832
         TabIndex        =   8
         Top             =   1140
         Width           =   2484
      End
      Begin VB.TextBox OT_Alterno 
         Height          =   285
         HelpContextID   =   2
         Left            =   156
         TabIndex        =   10
         Top             =   1850
         Width           =   2450
      End
      Begin VB.TextBox OT_Email 
         Height          =   288
         HelpContextID   =   2
         Left            =   2832
         TabIndex        =   12
         Top             =   1850
         Width           =   2496
      End
      Begin VB.Label OE_Host 
         AutoSize        =   -1  'True
         Caption         =   "&Host Local"
         Height          =   192
         Left            =   156
         TabIndex        =   1
         Top             =   264
         Width           =   768
      End
      Begin VB.Label OE_Direccion 
         AutoSize        =   -1  'True
         Caption         =   "&Dirección IP"
         Height          =   192
         Left            =   2832
         TabIndex        =   3
         Top             =   264
         Width           =   864
      End
      Begin VB.Label OE_Nombre 
         AutoSize        =   -1  'True
         Caption         =   "&Nombre Real"
         Height          =   192
         Left            =   156
         TabIndex        =   5
         Top             =   930
         Width           =   972
      End
      Begin VB.Label OE_Alias 
         AutoSize        =   -1  'True
         Caption         =   "A&lias"
         Height          =   192
         Left            =   2832
         TabIndex        =   7
         Top             =   930
         Width           =   360
      End
      Begin VB.Label OE_Alterno 
         AutoSize        =   -1  'True
         Caption         =   "Alias Alterno"
         Height          =   192
         Left            =   204
         TabIndex        =   9
         Top             =   1620
         Width           =   900
      End
      Begin VB.Label OE_Email 
         AutoSize        =   -1  'True
         Caption         =   "&E-mail"
         Height          =   192
         Left            =   2832
         TabIndex        =   11
         Top             =   1620
         Width           =   456
      End
   End
   Begin VB.CommandButton OB_Salir 
      Caption         =   "&Salir"
      Height          =   312
      HelpContextID   =   2
      Left            =   4632
      TabIndex        =   14
      Top             =   2460
      Width           =   972
   End
   Begin VB.CommandButton OB_Aceptar 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Height          =   312
      HelpContextID   =   2
      Left            =   3552
      TabIndex        =   13
      Top             =   2460
      Width           =   972
   End
End
Attribute VB_Name = "OF_Usuarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
' /*****************************************************/
' Llamar al procedimiento local PL_Carga_Datos, el cual
' recupera la información de la tabla de usuario
' /*****************************************************/
PL_Carga_Datos

End Sub


Sub PL_Carga_Datos()
' /******************************************************/
' Procedimiento local PL_Carga_Datos, el cual recupera la
' información de la tabla de usuario
' /******************************************************/

Dim Usuario As ES_USUARIO
Dim L_Result As Long

L_Result = MD_Recupera_Infousuario(Usuario)
If L_Result = 0 Then
    OT_Host = Usuario.E_Host
    OT_Direccion = Usuario.E_IP
    OT_Nombre = Usuario.E_Nombre
    OT_Alias = Usuario.E_Alias
    OT_Alterno = Usuario.E_nombre_alterno
    OT_Email = Usuario.E_EMAIL
Else
   MG_Mensaje _
    "< Información del Usuario no pudo ser recuperada... >"
   
End If
End Sub

Private Sub OB_Aceptar_Click()
' /*******************************************************/
' Evento click en le botón de aceptar lo cual indica que
' el usuario desea guardar la información modificada
' /*******************************************************/

Dim Usuario As ES_USUARIO
Dim L_Result As Long

If Trim(OT_Alias) = "" Then
  MG_Mensaje "Alias no puede quedar Vacio"
  OT_Alias.SetFocus
  Exit Sub
End If

Usuario.E_Host = OT_Host
Usuario.E_IP = OT_Direccion
Usuario.E_Nombre = OT_Nombre
Usuario.E_Alias = OT_Alias
Usuario.E_nombre_alterno = OT_Alterno
Usuario.E_EMAIL = OT_Email

L_Result = MD_Actualiza_Infousuario(Usuario)
If L_Result = 0 Then
  Unload Me
End If

End Sub

Private Sub OB_Salir_Click()
' Descargar la forma
Unload Me
End Sub


