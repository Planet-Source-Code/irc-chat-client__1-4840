VERSION 5.00
Begin VB.Form OF_Tipos_Servidores 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tipos de Servidores"
   ClientHeight    =   2790
   ClientLeft      =   1335
   ClientTop       =   1440
   ClientWidth     =   6105
   HelpContextID   =   3
   Icon            =   "OF_Tipos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2790
   ScaleWidth      =   6105
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton OB_Eliminar 
      Caption         =   "&Eliminar"
      Height          =   350
      HelpContextID   =   3
      Left            =   2400
      TabIndex        =   2
      Top             =   2325
      Width           =   975
   End
   Begin VB.CommandButton OB_Modificar 
      Caption         =   "&Modificar"
      Height          =   350
      HelpContextID   =   3
      Left            =   1320
      TabIndex        =   1
      Top             =   2325
      Width           =   975
   End
   Begin VB.CommandButton OB_agregar 
      Caption         =   "&Agregar"
      Height          =   350
      HelpContextID   =   3
      Left            =   240
      TabIndex        =   0
      Top             =   2325
      Width           =   975
   End
   Begin VB.CommandButton OB_salir 
      Caption         =   "&Salir"
      Height          =   350
      HelpContextID   =   3
      Left            =   4860
      TabIndex        =   3
      Top             =   2310
      Width           =   975
   End
   Begin VB.Frame OM_Tipos_Servidores 
      Caption         =   "Tipos"
      Height          =   2040
      HelpContextID   =   3
      Left            =   240
      TabIndex        =   4
      Top             =   90
      Width           =   5595
      Begin VB.ListBox OL_Tipos_Servidores 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1260
         HelpContextID   =   3
         Left            =   300
         TabIndex        =   5
         Top             =   480
         Width           =   4995
      End
      Begin VB.Label OE_Tipos_de_Servidores 
         Caption         =   "&Tipos de Servidores"
         Height          =   195
         Left            =   285
         TabIndex        =   6
         Top             =   270
         Width           =   2250
      End
   End
   Begin VB.Frame OM_Tipo_Servidor 
      Caption         =   "Datos"
      Height          =   2040
      Left            =   255
      TabIndex        =   7
      Top             =   90
      Width           =   5595
      Begin VB.TextBox OT_Tipo 
         Enabled         =   0   'False
         Height          =   285
         HelpContextID   =   3
         Left            =   270
         TabIndex        =   9
         Top             =   516
         Width           =   1380
      End
      Begin VB.TextBox OT_Descripcion 
         Height          =   315
         HelpContextID   =   3
         Left            =   270
         TabIndex        =   8
         Top             =   1236
         Width           =   5016
      End
      Begin VB.Label OE_Tipo 
         AutoSize        =   -1  'True
         Caption         =   "&Tipo"
         Height          =   195
         Left            =   270
         TabIndex        =   11
         Top             =   285
         Width           =   330
      End
      Begin VB.Label OE_Descripcion 
         Caption         =   "&Descripción"
         Height          =   195
         Left            =   270
         TabIndex        =   10
         Top             =   1020
         Width           =   1110
      End
   End
End
Attribute VB_Name = "OF_Tipos_Servidores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Lista de arreglo que contiene los datos para actualizar un
' registro
Dim LF_Datos(0 To 2) As Control
'Recordset en el que se manejan los registro a actualizar
Dim LF_Tabla As Recordset

Private Sub Cargar_Tipos()
' /********************************************************/
' Procedimiento utilizado para cargar la lista de tipos de
' servidores
' /********************************************************/
Dim L_Tipos_Servidores As Recordset


Set L_Tipos_Servidores = GV_Base_De_Datos.OpenRecordset( _
"select descripcion, tipo from tipos", dbOpenSnapshot)
MD_Cargar_Lista L_Tipos_Servidores, OL_Tipos_Servidores

End Sub

Private Sub PL_Habilitar(LP_habilitar1%, LP_habilitar2%, LP_habilitar3%, LP_habilitar4%, LP_habilitar5%)
' /********************************************************/
' Procedimiento encargado de hacer visible o invisibles y
' activos o inactivos  ciertos controles de acuerdo a la
' operación que se esté realizando.
' Ej:  Cuando se agrega un registro, se deshabilitan los
' botones de modificar y  eliminar.
' Además se define el botón OB_Salir como de Salir o de
' Cancelar
' /********************************************************/

OM_Tipo_Servidor.Visible = LP_habilitar1
OM_Tipos_Servidores.Visible = LP_habilitar2
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
' Se colocan los campos a actualizar de la tabla tipos en
' un arreglo que se utiliza para tal efecto.  Se definen
' las propiedades de cada campo, como
' ser si son modificables, si son obligatorios, o el
' mensaje que se desea se
' despliegue cuando un campo obligatorio es dejado en blanco.
' /********************************************************/

Cargar_Tipos

' Se colocan los campos de un registro a agregar o modificar
' en el arreglo de control LF_Datos

OT_Tipo.Tag = " %@[Tipo]": Set LF_Datos(0) = OT_Tipo
OT_Descripcion.Tag = _
" @[Descripcion]^Debe Ingresar Descripción^": _
Set LF_Datos(1) = OT_Descripcion
End Sub

Private Function PL_Validacion()
' /********************************************************/
' Se revisan los campos de un registro con el propósito de
' encontrar aqellos que
' deberían contener datos y por el contrario se encuentran
' vacios, caso en el que se despliega un mensaje previamente
' definido en el evento load de la forma.
' /*********************************************************/

Dim L_i%

L_i = MD_Validar_Nulos(LF_Datos())

If L_i >= 0 Then
    LF_Datos(L_i).SetFocus
    MG_Mensaje MD_Obtener_String("^", (LF_Datos(L_i).Tag), "^")
    Exit Function
End If
PL_Validacion = -1
End Function

Private Sub Form_Unload(Cancel As Integer)
' /************************************************************/
' Se evita descargar la forma solo si el caption del botón
' OB_salir se encuentra en modo Cancelar.
' /************************************************************/

If OB_Salir.Caption = "&Cancelar" Then Cancel = 1
End Sub

Private Sub OB_agregar_Click()
' /************************************************************/
' Evento utilizado para agregar un registro a la tabla tipos.
' Esta opción de agregar se realiza en dos pasos.  Primero se
' coloca el espacio en blanco para que el usuario ingrese los
' datos deseados. El segundo paso graba a la tabla
' los nuevos datos.
' /************************************************************/

On Error GoTo Etiqueta_Error
If OB_Salir.Caption = "&Salir" Then
' Primer paso, seteo de controles y preparación de campos para
' nuevo registro.
    PL_Habilitar True, False, True, False, False
    OT_Descripcion.SetFocus
    Set LF_Tabla = GV_Base_De_Datos.OpenRecordset( _
    "select * from Tipos", dbOpenDynaset)
Else
' Segundo paso, verificación de datos y grabado de nuevo registro.
    If Not PL_Validacion Then Exit Sub
    MD_Actualizar_Tabla LF_Datos(), LF_Tabla, 4
    MG_Mensaje "Registro Ha Sido Agregado"
    MD_Limpiar_Datos LF_Datos()
    PL_Habilitar False, True, True, True, True
    Cargar_Tipos
End If
Exit Sub

Etiqueta_Error:
    ME_Muestra_Error
    
End Sub
Private Sub OB_Eliminar_Click()
' /*******************************************************************************\
' Evento utilizado para eliminar un registro de la tabla tipos.  Esta opción
' se realiza de la siguiente forma:  El usuario indica el registro a eliminar,
' luego se pregunta al mismo si desea realmente eliminarlo.  De contestar que si
' se elimina el registro de la base de datos.
' /*******************************************************************************\

On Error GoTo Etiqueta_Error

If OL_Tipos_Servidores.ListIndex > -1 Then
    If MG_Pregunta("¿Desea Eliminar Registro?") = vbYes Then
        Set LF_Tabla = GV_Base_De_Datos.OpenRecordset("select * from Tipos where Tipo= " & OL_Tipos_Servidores.ItemData(OL_Tipos_Servidores.ListIndex), dbOpenDynaset)
        LF_Tabla.Delete
        OL_Tipos_Servidores.RemoveItem (OL_Tipos_Servidores.ListIndex)
    End If
Else
    MG_Mensaje "No Existen tipos de Servidores Definidos"
End If
Exit Sub

Etiqueta_Error:
    ME_Muestra_Error
    
End Sub

Private Sub OB_Modificar_Click()
' /************************************************************/
' Evento utilizado para modificar un registro de la tabla Tipos.
' Esta opción se realiza en dos pasos.  Primero se recupera el
' registro a modificar y se presentan
' los datos al usuario para que este realize las modificaciones
' deseadas. El segundo paso registra las modificaciones en la
' base de datos.
' /***********************************************************/


On Error GoTo Etiqueta_Error

If OB_Salir.Caption = "&Salir" Then
' Primer paso, recuperación de registro.
    If OL_Tipos_Servidores.ListIndex > -1 Then
        Set LF_Tabla = GV_Base_De_Datos.OpenRecordset( _
        "select * from Tipos where Tipo= " & _
        OL_Tipos_Servidores.ItemData( _
        OL_Tipos_Servidores.ListIndex), dbOpenDynaset)
        MD_Cargar_Datos LF_Datos, LF_Tabla, 4
        PL_Habilitar True, False, False, True, False
        OT_Descripcion.SetFocus
    Else
        MG_Mensaje "Debe Seleccionar Un Tipo de Servidor"
    End If
Else
' Segundo paso, modificación de registro.
    If Not PL_Validacion Then Exit Sub
    MD_Actualizar_Tabla LF_Datos(), LF_Tabla, 5
    MG_Mensaje "Registro Ha Sido Modificado"
    MD_Limpiar_Datos LF_Datos()
    PL_Habilitar False, True, True, True, True
    Cargar_Tipos
End If
Exit Sub

Etiqueta_Error:
    ME_Muestra_Error
    
End Sub

Private Sub OB_Salir_Click()
' /******************************************************/
' Este evento se utiliza para cerrar la forma o para
' cancelar la operación (agregar,modificar, eliminar)que
' se estaba realizando.
' /******************************************************/

If OB_Salir.Caption = "&Salir" Then
    Unload Me
Else
    MD_Limpiar_Datos LF_Datos()
    PL_Habilitar False, True, True, True, True
End If
End Sub





