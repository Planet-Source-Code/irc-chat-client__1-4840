VERSION 5.00
Begin VB.Form OF_Servidores 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Servidores"
   ClientHeight    =   5265
   ClientLeft      =   2160
   ClientTop       =   1230
   ClientWidth     =   6675
   HelpContextID   =   4
   Icon            =   "OF_Servi.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5265
   ScaleWidth      =   6675
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox OT_Tipo_Servidor 
      Height          =   285
      Left            =   4035
      TabIndex        =   21
      Text            =   "Text1"
      Top             =   4815
      Visible         =   0   'False
      Width           =   168
   End
   Begin VB.CommandButton OB_Eliminar 
      Caption         =   "&Eliminar"
      Height          =   350
      HelpContextID   =   4
      Left            =   2370
      TabIndex        =   7
      Top             =   4815
      Width           =   975
   End
   Begin VB.CommandButton OB_Modificar 
      Caption         =   "&Modificar"
      Height          =   350
      HelpContextID   =   4
      Left            =   1290
      TabIndex        =   6
      Top             =   4815
      Width           =   975
   End
   Begin VB.CommandButton OB_agregar 
      Caption         =   "&Agregar"
      Height          =   350
      HelpContextID   =   4
      Left            =   180
      TabIndex        =   5
      Top             =   4815
      Width           =   975
   End
   Begin VB.CommandButton OB_salir 
      Caption         =   "&Salir"
      Height          =   350
      HelpContextID   =   4
      Left            =   5535
      TabIndex        =   8
      Top             =   4815
      Width           =   975
   End
   Begin VB.Frame OM_Servidores 
      Caption         =   "Servidores"
      Height          =   4560
      HelpContextID   =   4
      Left            =   210
      TabIndex        =   0
      Top             =   150
      Width           =   6255
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
         HelpContextID   =   4
         Left            =   300
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   555
         Width           =   5655
      End
      Begin VB.ListBox OL_Servidores 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2940
         HelpContextID   =   4
         Left            =   300
         Sorted          =   -1  'True
         TabIndex        =   4
         Top             =   1275
         Width           =   5640
      End
      Begin VB.Label OE_Servidores 
         Caption         =   "Se&rvidores"
         Height          =   210
         Left            =   300
         TabIndex        =   3
         Top             =   1050
         Width           =   2610
      End
      Begin VB.Label OE_Tipos_de_Servidores 
         Caption         =   "&Tipos de Servidores"
         Height          =   195
         Left            =   285
         TabIndex        =   1
         Top             =   300
         Width           =   2250
      End
   End
   Begin VB.Frame OM_Servidor 
      Caption         =   "Datos de Servidores"
      Height          =   4515
      Left            =   216
      TabIndex        =   22
      Top             =   180
      Width           =   6240
      Begin VB.TextBox OT_Codigo 
         Enabled         =   0   'False
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
         HelpContextID   =   4
         Left            =   330
         TabIndex        =   10
         Top             =   600
         Width           =   1380
      End
      Begin VB.TextBox OT_Descripcion 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         HelpContextID   =   4
         Left            =   330
         TabIndex        =   12
         Top             =   1290
         Width           =   5580
      End
      Begin VB.TextBox OT_Direccion 
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
         HelpContextID   =   4
         Left            =   330
         TabIndex        =   14
         Top             =   2040
         Width           =   5580
      End
      Begin VB.TextBox OT_Puertos 
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
         HelpContextID   =   4
         Left            =   330
         TabIndex        =   16
         Top             =   2910
         Width           =   5580
      End
      Begin VB.TextBox OT_Ultimo_Puerto 
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
         HelpContextID   =   4
         Left            =   345
         TabIndex        =   18
         Top             =   3855
         Width           =   2100
      End
      Begin VB.TextBox OT_Contraseña 
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
         HelpContextID   =   4
         Left            =   3795
         TabIndex        =   20
         Top             =   3855
         Width           =   2052
      End
      Begin VB.Label OE_Codigo 
         AutoSize        =   -1  'True
         Caption         =   "Código"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   330
         TabIndex        =   9
         Top             =   345
         Width           =   660
      End
      Begin VB.Label OE_Descripcion 
         Caption         =   "&Descripción"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   330
         TabIndex        =   11
         Top             =   1005
         Width           =   1110
      End
      Begin VB.Label OE_Direccion 
         Caption         =   "Di&rección"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   330
         TabIndex        =   13
         Top             =   1785
         Width           =   990
      End
      Begin VB.Label OE_Puertos 
         Caption         =   "&Puertos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   330
         TabIndex        =   15
         Top             =   2640
         Width           =   990
      End
      Begin VB.Label OE_Ultimo_Puerto 
         AutoSize        =   -1  'True
         Caption         =   "&Ultimo Puerto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   330
         TabIndex        =   17
         Top             =   3615
         Width           =   1200
      End
      Begin VB.Label OE_Contraseña 
         AutoSize        =   -1  'True
         Caption         =   "Con&traseña"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3705
         TabIndex        =   19
         Top             =   3615
         Width           =   1035
      End
   End
End
Attribute VB_Name = "OF_Servidores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Lista de arreglo que contiene los datos para actualizar un registro
Dim LF_Datos(0 To 7) As Control
'Recordset en el que se manejan los registro a actualizar
Dim LF_Tabla As Recordset

Private Sub PL_Habilitar(LP_habilitar1%, LP_habilitar2%, LP_habilitar3%, LP_habilitar4%, LP_habilitar5%)
' /*****************************************************/
' Procedimiento encargado de hacer visible o invisibles
' y activos o inactivos ciertos controles de acuerdo a
' la operación que se esté realizando.
' Ej:  Cuando se agrega un registro, se deshabilitan
' los botones de modificar y eliminar.
' Además se define el botón OB_Salir como de Salir o de
' Cancelar
' /*****************************************************/

OM_Servidor.Visible = LP_habilitar1
OM_Servidores.Visible = LP_habilitar2
OB_Agregar.Enabled = LP_habilitar3
OB_Modificar.Enabled = LP_habilitar4
OB_Eliminar.Enabled = LP_habilitar5
If OB_Salir.Caption = "&Salir" Then
    OB_Salir.Caption = "&Cancelar"
Else
    OB_Salir.Caption = "&Salir"
End If
End Sub

Private Sub Form_Load()
' /*******************************************************/
' En este evento se carga la lista de tipos de servidores.
' Se colocan además los campos a actualizar en un arreglo
' que se utiliza para tal efecto.  Se definen las
' propiedades de cada campo, como ser si son modificables,
' si son obligatorios, o el mensaje que se desea se
' despliegue cuando un campo obligatorio es dejado
' en blanco.
' /*******************************************************/

Dim L_Tipos_Servidores As Recordset

' Cargar la lista de tipos de servidores
Set L_Tipos_Servidores = GV_Base_De_Datos.OpenRecordset( _
      "select descripcion, tipo from tipos", dbOpenSnapshot)
      
If MD_Cargar_Lista(L_Tipos_Servidores, OC_Tipos_Servidores) _
  = 0 Then
  PL_Habilitar True, False, False, False, False
  OB_Salir.Caption = "&Salir"
  Exit Sub
Else
  ' Se colocan los campos de un registro a agregar o modificar _
  en el arreglo de control LF_Datos
  OT_Codigo.Tag = " %@[Codigo]": Set LF_Datos(0) = OT_Codigo
  OT_Descripcion.Tag = _
  " @[Descripcion]^Debe Ingresar Descripción^": _
  Set LF_Datos(1) = OT_Descripcion
  
  OT_Direccion.Tag = _
  " @[Direccion]^Debe Ingresar Dirección^": _
  Set LF_Datos(2) = OT_Direccion
  
  OT_Puertos.Tag = _
  " @[Puertos]^Debe Ingresar Puertos^": _
  Set LF_Datos(3) = OT_Puertos
   
  OT_Ultimo_Puerto.Tag = _
  " @[U_Puerto]^Debe Ingresar Ultimo Puerto^": _
  Set LF_Datos(4) = OT_Ultimo_Puerto
    
  OT_Contraseña.Tag = _
  " @[Contrasena]": Set LF_Datos(5) = OT_Contraseña
  
  OT_Tipo_Servidor.Tag = _
  " &[Tipo_Servidor]": Set LF_Datos(6) = OT_Tipo_Servidor
 End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
' /****************************************************/
' Se evita descargar la forma solo si el caption del
' botón OB_salir se encuentra
' en modo Cancelar.
' /****************************************************/

If OB_Salir.Caption = "&Cancelar" Then Cancel = 1
End Sub

Private Sub OB_agregar_Click()
' /****************************************************/
' Evento utilizado para agregar un registro a la tabla
' de servidores.  Esta opción de agregar se realiza en
' dos pasos.  Primero se coloca el espacio en blanco para
' que el usuario ingrese los datos deseados . El segundo
' paso graba a la tabla los nuevos datos.
' /****************************************************/

On Error GoTo Etiqueta_Error

If OB_Salir.Caption = "&Salir" Then
' Primer paso, seteo de controles y preparación de campos
' para nuevo registro.
    If OC_Tipos_Servidores.ListCount > 0 Then
     PL_Habilitar True, False, True, False, False
     OT_Descripcion.SetFocus
     Set LF_Tabla = GV_Base_De_Datos.OpenRecordset( _
     "select * from Servidores where Tipo_Servidor = " _
     & OT_Tipo_Servidor, dbOpenDynaset)
    Else
    End If
Else
' Segundo paso, verificación de datos y grabado de
' nuevo registro.
    If Not PL_Validacion Then Exit Sub
    MD_Actualizar_Tabla LF_Datos(), LF_Tabla, 4
    OC_Tipos_Servidores_Click
    MD_Limpiar_Datos LF_Datos()
    PL_Habilitar False, True, True, True, True
End If
Exit Sub

Etiqueta_Error:
    ME_Muestra_Error
    
End Sub

Private Sub OB_Eliminar_Click()
' /***************************************************/
' Evento utilizado para eliminar un registro de la
' tabla de servidores.  Esta  opción se realiza de la
' siguiente forma:  El usuario indica el registro a
' eliminar, luego se pregunta al mismo si desea
' realmente eliminarlo.  De contestar que si se elimina
' el registro de la base de datos.
' /***************************************************/

On Error GoTo Etiqueta_Error

If OC_Tipos_Servidores.ListCount > 0 Then
  If OL_Servidores.ListIndex > -1 Then
    If MG_Pregunta("¿Desea Eliminar Registro?") = vbYes Then
      Set LF_Tabla = GV_Base_De_Datos.OpenRecordset( _
      "select * from Servidores where Codigo= " _
      & OL_Servidores.ItemData(OL_Servidores.ListIndex), _
      dbOpenDynaset)
      LF_Tabla.Delete
      OL_Servidores.RemoveItem (OL_Servidores.ListIndex)
        End If
     Else
        MG_Mensaje "Debe Seleccionar Un Servidor"
    End If
Else
    MG_Mensaje "No Existen tipos de Servidores Definidos"
End If
Exit Sub

Etiqueta_Error:
    ME_Muestra_Error
    
End Sub

Private Function PL_Validacion()
' /*******************************************************/
' Se revisan los campos de un registro con el propósito de
' encontrar aqellos que deberían contener datos y por el
' contrario se encuentran vacios, caso en el que se
' despliega un mensaje previamente definido en el evento
' load de la forma.
' /*******************************************************/

Dim L_i%

L_i = MD_Validar_Nulos(LF_Datos())

If L_i >= 0 Then
 LF_Datos(L_i).SetFocus
 MG_Mensaje MD_Obtener_String("^", (LF_Datos(L_i).Tag), "^")
 Exit Function
End If
PL_Validacion = -1
End Function

Private Sub OB_Modificar_Click()
' /********************************************************/
' evento utilizado para modificar un registro de la tabla
' de servidores.  Esta opción se realiza en dos pasos.
' Primero se recupera el registro a modificar y se presentan
' los datos al usuario para que este realize las
' modificaciones deseadas. El segundo
' paso registra las modificaciones en la base de datos.
' /********************************************************/

On Error GoTo Etiqueta_Error

If OB_Salir.Caption = "&Salir" Then
' Primer paso, recuperación de registro.
    If OC_Tipos_Servidores.ListCount > 0 Then
        If OL_Servidores.ListIndex > -1 Then
            Set LF_Tabla = GV_Base_De_Datos.OpenRecordset( _
            "select * from Servidores where Codigo= " _
            & OL_Servidores.ItemData(OL_Servidores.ListIndex), _
            dbOpenDynaset)
            MD_Cargar_Datos LF_Datos, LF_Tabla, 4
            PL_Habilitar True, False, False, True, False
            OT_Descripcion.SetFocus
        Else
            MG_Mensaje "Debe Seleccionar Un Servidor"
        End If
    Else
        MG_Mensaje "No Existen tipos de Servidores Definidos"
    End If
Else
' Segundo paso, modificación de registro.
    If Not PL_Validacion Then Exit Sub
    MD_Actualizar_Tabla LF_Datos(), LF_Tabla, 5
    OC_Tipos_Servidores_Click
    MD_Limpiar_Datos LF_Datos()
    PL_Habilitar False, True, True, True, True
End If
Exit Sub

Etiqueta_Error:
    ME_Muestra_Error
    
End Sub

Private Sub OB_Salir_Click()
' /********************************************************/
' Este evento se utiliza para cerrar la forma o para
' cancelar la operación (agregar,modificar, eliminar)que
' se estaba realizando.
' /********************************************************/

If OB_Salir.Caption = "&Salir" Then
    Unload Me
Else
    MD_Limpiar_Datos LF_Datos()
    PL_Habilitar False, True, True, True, True
End If
End Sub

Private Sub OC_Tipos_Servidores_Click()
' /********************************************************/
' Este evento se utiliza para cargar los servidores
' existentes de acuerdo al tipo de servidor seleccionado.
' /********************************************************/

Dim L_Servidores As Recordset

If OC_Tipos_Servidores.ListCount > 0 Then
    Set L_Servidores = GV_Base_De_Datos.OpenRecordset( _
    "select descripcion + ' ('+ Direccion + ' )', codigo " + _
    "from Servidores where tipo_servidor = " _
    & OC_Tipos_Servidores.ItemData( _
    OC_Tipos_Servidores.ListIndex), dbOpenSnapshot)
    MD_Cargar_Lista L_Servidores, OL_Servidores
    OT_Tipo_Servidor = OC_Tipos_Servidores.ItemData( _
    OC_Tipos_Servidores.ListIndex)
Else
    OT_Tipo_Servidor = 0
End If
End Sub

Private Sub OL_Servidores_DblClick()
' /********************************************************/
' Este evento se utiliza llamar a la función MM_Connect.
' Esta nos permite conectarnos al servidor al cual le hemos
' dado un dobleclick.
' /********************************************************/

Set LF_Tabla = GV_Base_De_Datos.OpenRecordset( _
"select Direccion,U_Puerto,Tipo_Servidor, Codigo from " + _
"Servidores where Codigo= " & OL_Servidores.ItemData( _
OL_Servidores.ListIndex), dbOpenDynaset)

MM_Connect LF_Tabla("direccion"), LF_Tabla("u_puerto"), _
LF_Tabla("tipo_servidor"), 0

MD_Actualiza_Ultimo_Servidor LF_Tabla("Codigo")
Unload Me
End Sub


