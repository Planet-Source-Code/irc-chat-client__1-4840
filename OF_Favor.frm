VERSION 5.00
Begin VB.Form OF_Servidores_Favoritos 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Servidores Favoritos"
   ClientHeight    =   4845
   ClientLeft      =   330
   ClientTop       =   1110
   ClientWidth     =   9780
   HelpContextID   =   5
   Icon            =   "OF_Favor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4845
   ScaleWidth      =   9780
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame OM_Marco 
      Caption         =   "Servidores Favoritos"
      Height          =   4155
      HelpContextID   =   5
      Left            =   135
      TabIndex        =   0
      Top             =   105
      Width           =   9525
      Begin VB.ListBox OL_Servidores 
         Height          =   2595
         HelpContextID   =   5
         Left            =   180
         MultiSelect     =   2  'Extended
         TabIndex        =   4
         Top             =   1215
         Width           =   4185
      End
      Begin VB.ComboBox OC_Tipos_Servidores 
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
         HelpContextID   =   5
         Left            =   180
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   480
         Width           =   4185
      End
      Begin VB.ListBox OL_Favoritos 
         Height          =   2595
         HelpContextID   =   5
         Left            =   5145
         MultiSelect     =   2  'Extended
         TabIndex        =   8
         Top             =   1215
         Width           =   4185
      End
      Begin VB.CommandButton OB_agregar 
         BackColor       =   &H00C0C0C0&
         Height          =   585
         HelpContextID   =   5
         Left            =   4485
         MaskColor       =   &H00FFFFFF&
         Picture         =   "OF_Favor.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   1845
         Width           =   585
      End
      Begin VB.CommandButton OB_Modificar 
         BackColor       =   &H00C0C0C0&
         Height          =   585
         HelpContextID   =   5
         Left            =   4470
         Picture         =   "OF_Favor.frx":0884
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   2625
         Width           =   585
      End
      Begin VB.Label OE_Tipos_de_Servidores 
         Caption         =   "&Tipos de Servidores"
         Height          =   195
         Left            =   180
         TabIndex        =   1
         Top             =   255
         Width           =   2250
      End
      Begin VB.Label OE_Servidores 
         Caption         =   "Se&rvidores"
         Height          =   210
         Left            =   180
         TabIndex        =   3
         Top             =   990
         Width           =   2610
      End
      Begin VB.Label OE_Servidores_Favoritos 
         Caption         =   "Servidores &Favoritos"
         Height          =   195
         Left            =   5175
         TabIndex        =   7
         Top             =   990
         Width           =   1920
      End
   End
   Begin VB.CommandButton OB_salir 
      Caption         =   "&Salir"
      Height          =   348
      HelpContextID   =   5
      Left            =   8565
      TabIndex        =   9
      Top             =   4395
      Width           =   1104
   End
End
Attribute VB_Name = "OF_Servidores_Favoritos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Recordset para manejar los servidores favoritos de nuestra
' aplicación
Dim LF_Favoritos As Recordset

Private Sub Form_Load()
' /********************************************************/
' Se llama a la función MD_Cargar_Lista, la cual cargar la
' lista de tipos de
' servidores.
' /********************************************************/
Dim L_Tipos_Servidores As Recordset

Set L_Tipos_Servidores = GV_Base_De_Datos.OpenRecordset( _
"select descripcion, tipo from tipos", dbOpenSnapshot)
MD_Cargar_Lista L_Tipos_Servidores, OC_Tipos_Servidores
End Sub

Private Sub OB_agregar_Click()
' /********************************************************/
' Este evento se utiliza para agregar un nuevo registro a
' la tabla de servidores  favoritos.
' /********************************************************/

Dim L_i, L_j As Integer
Dim L_Existe As Boolean
Dim L_Favoritos As Recordset

If OL_Servidores.ListCount > 0 Then
    For L_i = 0 To (OL_Servidores.ListCount - 1)
        If OL_Servidores.Selected(L_i) = True Then
            For L_j = 0 To (OL_Favoritos.ListCount - 1)
                If OL_Favoritos.ItemData(L_j) = _
                OL_Servidores.ItemData(L_i) Then
                   L_Existe = True
                   Exit For
                End If
            Next L_j
            If Not L_Existe Then
                OL_Favoritos.AddItem OL_Servidores.List(L_i)
                OL_Favoritos.ItemData(OL_Favoritos.NewIndex) _
                = OL_Servidores.ItemData(L_i)
                'Grabar nuevo registro a la base de datos
                Set L_Favoritos = _
                GV_Base_De_Datos.OpenRecordset( _
                "select * from Servidores_Favoritos", _
                dbOpenDynaset)
                L_Favoritos.AddNew
                L_Favoritos("Codigo") = _
                OL_Servidores.ItemData(L_i)
                L_Favoritos("Tipo_Servidor") = _
                OC_Tipos_Servidores.ItemData( _
                OC_Tipos_Servidores.ListIndex)
                L_Favoritos.Update
            Else
                L_Existe = False
            End If
        End If
    Next L_i
End If
End Sub

Private Sub OB_Modificar_Click()
' /*****************************************************/
' Este evento se utiliza para eliminar un registro de la
' tabla de servidores
' favoritos.
' /*****************************************************/

Dim L_i As Integer
Dim L_Favoritos As Recordset

If OL_Favoritos.ListCount > 0 Then
    For L_i = 0 To (OL_Favoritos.ListCount - 1)
        If OL_Favoritos.Selected(L_i) = True Then
            Set L_Favoritos = GV_Base_De_Datos.OpenRecordset( _
            "select * from Servidores_Favoritos", dbOpenDynaset)
            GV_Base_De_Datos.Execute _
            "delete from Servidores_Favoritos " + _
            "where Codigo = " & _
            OL_Favoritos.ItemData(L_i) & " and " + _
            "Tipo_Servidor = " & _
            OC_Tipos_Servidores.ItemData( _
            OC_Tipos_Servidores.ListIndex)
        End If
    Next L_i
    Set LF_Favoritos = GV_Base_De_Datos.OpenRecordset( _
       "select descripcion + ' ( ' + Direccion + ' ) ', " + _
       "Servidores_Favoritos.codigo, " + _
       "Servidores_Favoritos.tipo_servidor  from " + _
       "Servidores_Favoritos, Servidores " + _
       "where Servidores_Favoritos.tipo_servidor = " _
       & OC_Tipos_Servidores.ItemData( _
       OC_Tipos_Servidores.ListIndex) & _
       " and Servidores.codigo = Servidores_Favoritos.codigo", _
       dbOpenSnapshot)
      MD_Cargar_Lista LF_Favoritos, OL_Favoritos
End If
End Sub

Private Sub OB_Salir_Click()
' /******************************************************/
' Descarga la forma.
' /******************************************************/

Unload Me
End Sub

Private Sub OC_Tipos_Servidores_Click()
' /******************************************************/
' Este evento se utiliza para llamar a la función
' MD_Cargar_Lista la cual carga las listas de servidores
' conocidos y servidores favoritos.
' /******************************************************/
Dim L_Servidores As Recordset

If OC_Tipos_Servidores.ListCount > 0 Then
    Set L_Servidores = _
    GV_Base_De_Datos.OpenRecordset( _
    "select descripcion + '( ' + Direccion + ' )' , " + _
    "codigo from Servidores where tipo_servidor = " & _
    OC_Tipos_Servidores.ItemData( _
    OC_Tipos_Servidores.ListIndex), dbOpenSnapshot)
    If MD_Cargar_Lista(L_Servidores, OL_Servidores) > 0 Then
        Set LF_Favoritos = GV_Base_De_Datos.OpenRecordset( _
        "select descripcion+ ' ( ' + Direccion + ' ) ', " + _
        "Servidores_Favoritos.codigo,  " + _
        "Servidores_Favoritos.tipo_servidor  " + _
        "from Servidores_Favoritos, Servidores " + _
        "where Servidores_Favoritos.tipo_servidor = " _
        & OC_Tipos_Servidores.ItemData( _
        OC_Tipos_Servidores.ListIndex) & _
        " and Servidores.codigo = Servidores_Favoritos.codigo", _
        dbOpenSnapshot)
        MD_Cargar_Lista LF_Favoritos, OL_Favoritos
    End If
End If
End Sub

Private Sub OL_Favoritos_DblClick()
' /*******************************************************/
' Este evento se utiliza llamar a la función MM_Connect.
' Esta nos permite conectarnos a uno de los servidores
' favoritos al cual le hemos dado un dobleclick.
' /*******************************************************/

Dim L_Favoritos As Recordset

Set L_Favoritos = GV_Base_De_Datos.OpenRecordset( _
"select Direccion,U_Puerto, Servidores.Tipo_Servidor, " + _
" Servidores.Codigo" & _
" from Servidores, Servidores_Favoritos where " + _
"Servidores.Codigo= " & _
OL_Favoritos.ItemData(OL_Favoritos.ListIndex), dbOpenDynaset)

MM_Connect L_Favoritos("direccion"), L_Favoritos("u_puerto"), _
L_Favoritos("tipo_servidor"), 0
MD_Actualiza_Ultimo_Servidor L_Favoritos("Codigo")
L_Favoritos.Close
Unload Me
End Sub

Private Sub OL_Servidores_DblClick()
' /*****************************************************/
' Este evento se utiliza llamar a la función MM_Connect.
' Esta nos permite conectarnos al servidor conocido al
' cual le hemos dado un dobleclick.
' /*****************************************************/

Dim L_Servidores As Recordset

Set L_Servidores = GV_Base_De_Datos.OpenRecordset( _
"select Direccion,U_Puerto,Tipo_Servidor, Codigo from " + _
" Servidores where Codigo= " & _
OL_Servidores.ItemData(OL_Servidores.ListIndex), _
dbOpenDynaset)
MM_Connect L_Servidores("direccion"), L_Servidores("u_puerto"), _
L_Servidores("tipo_servidor"), 0
MD_Actualiza_Ultimo_Servidor L_Servidores("Codigo")
L_Servidores.Close
Unload Me
End Sub


