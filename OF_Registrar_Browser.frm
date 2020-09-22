VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "comdlg32.ocx"
Begin VB.Form OF_Registrar_Browser 
   Caption         =   "Registrar el PATH del Browser de Internet"
   ClientHeight    =   1725
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7800
   HelpContextID   =   25
   Icon            =   "OF_Registrar_Browser.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1725
   ScaleWidth      =   7800
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton OB_Salir 
      Caption         =   "&Salir"
      Height          =   360
      HelpContextID   =   25
      Left            =   4005
      TabIndex        =   4
      Top             =   1125
      Width           =   1560
   End
   Begin VB.CommandButton OT_Aceptar 
      Caption         =   "&Aceptar"
      Height          =   360
      HelpContextID   =   25
      Left            =   2280
      TabIndex        =   3
      Top             =   1125
      Width           =   1560
   End
   Begin VB.CommandButton OB_Path 
      Caption         =   "..."
      Height          =   330
      HelpContextID   =   25
      Left            =   7215
      TabIndex        =   2
      Top             =   540
      Width           =   360
   End
   Begin VB.TextBox OT_Path 
      Height          =   315
      HelpContextID   =   25
      Left            =   240
      TabIndex        =   0
      Top             =   540
      Width           =   6930
   End
   Begin MSComDlg.CommonDialog Dialog 
      Left            =   855
      Top             =   1065
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327680
   End
   Begin VB.Label Label1 
      Caption         =   "Path del Browser de Internet"
      Height          =   255
      Left            =   255
      TabIndex        =   1
      Top             =   225
      Width           =   2955
   End
End
Attribute VB_Name = "OF_Registrar_Browser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
'/*******************************************************/
' Cargar el browser que se tiene registrado
' /******************************************************/
Dim L_Registro As Recordset
Set L_Registro = GV_Base_De_Datos.OpenRecordset("Ini")
If Not L_Registro.EOF Then
 If IsNull(L_Registro!Browser) Then
    OT_Path = ""
 Else
    OT_Path = L_Registro!Browser
 End If

End If
L_Registro.Close
End Sub

Private Sub OB_Path_Click()
Dialog.filename = ""
Dialog.ShowOpen

OT_Path = Dialog.filename


End Sub


Private Sub OB_Salir_Click()
Unload Me
End Sub

Private Sub OT_Aceptar_Click()
Dim L_Registro As Recordset

Set L_Registro = GV_Base_De_Datos.OpenRecordset("Ini")
If Not L_Registro.EOF Then
   L_Registro.Edit
   L_Registro!Browser = OT_Path
   
   L_Registro.Update
End If

L_Registro.Close
Unload Me
End Sub

