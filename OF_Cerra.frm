VERSION 5.00
Begin VB.Form OF_Cerrar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cerrar Conexiones"
   ClientHeight    =   3060
   ClientLeft      =   1710
   ClientTop       =   1635
   ClientWidth     =   4860
   HelpContextID   =   7
   Icon            =   "OF_Cerra.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3060
   ScaleWidth      =   4860
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame OM_Marco 
      Caption         =   "Servidores Activos"
      Height          =   2388
      HelpContextID   =   7
      Left            =   96
      TabIndex        =   2
      Top             =   132
      Width           =   4596
      Begin VB.TextBox OT_Nombre 
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
         Height          =   285
         Index           =   1
         Left            =   348
         TabIndex        =   17
         Top             =   384
         Visible         =   0   'False
         Width           =   3700
      End
      Begin VB.TextBox OT_Nombre 
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
         Height          =   285
         Index           =   2
         Left            =   348
         TabIndex        =   16
         Top             =   756
         Visible         =   0   'False
         Width           =   3700
      End
      Begin VB.TextBox OT_Nombre 
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
         Height          =   285
         Index           =   3
         Left            =   348
         TabIndex        =   15
         Top             =   1104
         Visible         =   0   'False
         Width           =   3700
      End
      Begin VB.TextBox OT_Nombre 
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
         Height          =   285
         Index           =   4
         Left            =   348
         TabIndex        =   14
         Top             =   1476
         Visible         =   0   'False
         Width           =   3700
      End
      Begin VB.TextBox OT_Nombre 
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
         Height          =   285
         Index           =   5
         Left            =   348
         TabIndex        =   13
         Top             =   1836
         Visible         =   0   'False
         Width           =   3700
      End
      Begin VB.CheckBox OCH_Cerrar 
         Height          =   195
         HelpContextID   =   7
         Index           =   1
         Left            =   4212
         TabIndex        =   12
         Top             =   432
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.CheckBox OCH_Cerrar 
         Height          =   195
         HelpContextID   =   7
         Index           =   2
         Left            =   4212
         TabIndex        =   11
         Top             =   804
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.CheckBox OCH_Cerrar 
         Height          =   195
         HelpContextID   =   7
         Index           =   3
         Left            =   4212
         TabIndex        =   10
         Top             =   1152
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.CheckBox OCH_Cerrar 
         Height          =   195
         HelpContextID   =   7
         Index           =   4
         Left            =   4230
         TabIndex        =   9
         Top             =   1515
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.CheckBox OCH_Cerrar 
         Height          =   195
         HelpContextID   =   7
         Index           =   5
         Left            =   4212
         TabIndex        =   8
         Top             =   1884
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox OT_Indice 
         Height          =   285
         Index           =   1
         Left            =   192
         TabIndex        =   7
         Top             =   396
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.TextBox OT_Indice 
         Height          =   285
         Index           =   2
         Left            =   192
         TabIndex        =   6
         Top             =   744
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.TextBox OT_Indice 
         Height          =   285
         Index           =   3
         Left            =   192
         TabIndex        =   5
         Top             =   1092
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.TextBox OT_Indice 
         Height          =   285
         Index           =   4
         Left            =   204
         TabIndex        =   4
         Top             =   1476
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.TextBox OT_Indice 
         Height          =   285
         Index           =   5
         Left            =   204
         TabIndex        =   3
         Top             =   1872
         Visible         =   0   'False
         Width           =   195
      End
   End
   Begin VB.CommandButton OB_Salir 
      Caption         =   "&Salir"
      Height          =   312
      HelpContextID   =   7
      Left            =   3756
      TabIndex        =   1
      Top             =   2628
      Width           =   972
   End
   Begin VB.CommandButton OB_Aceptar 
      Caption         =   "&Aceptar"
      Height          =   312
      HelpContextID   =   7
      Left            =   2712
      TabIndex        =   0
      Top             =   2628
      Width           =   972
   End
End
Attribute VB_Name = "OF_Cerrar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub PL_Cerrar_Conexiones()
' /*********************************************************/
' Este procedimiento recorre los servidores y para todos los
' servidores marcados se procede a ejecutar el procedimiento
' que cierra la conexión con un servidor.
' /*********************************************************/
Dim L_i As Integer
Dim L_libre As Integer

L_libre = 1

For L_i = 1 To 5
     If OCH_Cerrar(L_i).Value = 1 Then
       'llamar al procedimiento de cerrar
       MM_Cerrar_Conexion OT_Indice(L_i)
     End If

Next L_i

Unload Me

End Sub

Sub PL_Cargar_Servidores_Activos()
' /*********************************************************/
' Este procedimiento carga todos los servidores activos en
' el momento Los servidores son cargados, para que el usuario
' tenga opción a cerrar varias conexiones  en diferentes
' servidores. Los servidores activos son cargados del arreglo
' global de Sockets
' /*********************************************************/

Dim L_i As Integer
Dim L_libre As Integer

L_libre = 1

For L_i = 1 To 5
  If GV_Sockets(L_i).socket <> INVALID_SOCKET Then
     OT_Indice(L_libre) = L_i
   
     OT_Nombre(L_libre).Visible = True
     OCH_Cerrar(L_libre).Visible = True
     OT_Nombre(L_libre) = GV_Sockets(L_i).Direcc + "(" + _
                       CStr(GV_Sockets(L_i).Puerto) + _
                       ") NICK => " + GV_Sockets(L_i).Nick
     L_libre = L_libre + 1
  
  End If

Next L_i

End Sub

Private Sub Form_Load()
'/* Procedimiento para cargar los servidores activos
PL_Cargar_Servidores_Activos
End Sub

Private Sub OB_Aceptar_Click()
' /*********************************************************/
' Revisar si hay servidores activos, si los hay entonces,
' llamar al procedimiento que cierra las conexiones de los
' servidores
' /*********************************************************/
If Trim(OT_Nombre(1)) = "" Then Unload Me: Exit Sub
               
   PL_Cerrar_Conexiones
               

               
End Sub

Private Sub OB_Salir_Click()
' Descargar la forma
Unload Me
End Sub
