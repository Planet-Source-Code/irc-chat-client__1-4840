VERSION 5.00
Begin VB.Form OF_Comandos 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Comandos"
   ClientHeight    =   5280
   ClientLeft      =   960
   ClientTop       =   1440
   ClientWidth     =   5655
   HelpContextID   =   11
   Icon            =   "OF_Coman.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5280
   ScaleWidth      =   5655
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox OT_Tipo_Servidor 
      Height          =   285
      Left            =   3456
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   4884
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton OB_Eliminar 
      Caption         =   "&Eliminar"
      Height          =   348
      HelpContextID   =   11
      Left            =   2265
      TabIndex        =   7
      Top             =   4836
      Width           =   975
   End
   Begin VB.CommandButton OB_Modificar 
      Caption         =   "&Modificar"
      Height          =   348
      HelpContextID   =   11
      Left            =   1188
      TabIndex        =   6
      Top             =   4836
      Width           =   975
   End
   Begin VB.CommandButton OB_Agregar 
      Caption         =   "&Agregar"
      Height          =   348
      HelpContextID   =   11
      Left            =   120
      TabIndex        =   5
      Top             =   4836
      Width           =   972
   End
   Begin VB.CommandButton OB_salir 
      Caption         =   "&Salir"
      Height          =   348
      HelpContextID   =   11
      Left            =   4560
      TabIndex        =   8
      Top             =   4815
      Width           =   975
   End
   Begin VB.Frame OM_Servidores 
      Caption         =   "Comandos"
      Height          =   4560
      HelpContextID   =   11
      Left            =   120
      TabIndex        =   0
      Top             =   84
      Width           =   5388
      Begin VB.ListBox OL_Comandos 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3180
         HelpContextID   =   11
         Left            =   270
         TabIndex        =   4
         Top             =   1185
         Width           =   4830
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
         HelpContextID   =   11
         Left            =   204
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   465
         Width           =   4965
      End
      Begin VB.Label OE_Comandos 
         Caption         =   "&Comandos"
         Height          =   210
         Left            =   255
         TabIndex        =   3
         Top             =   930
         Width           =   2610
      End
      Begin VB.Label OE_Tipos_de_Servidores 
         Caption         =   "&Tipos de Servidores"
         Height          =   192
         Left            =   204
         TabIndex        =   1
         Top             =   240
         Width           =   2256
      End
   End
   Begin VB.Frame OM_Comando 
      Caption         =   "Datos de Comandos"
      Height          =   4530
      HelpContextID   =   11
      Left            =   132
      TabIndex        =   10
      Top             =   108
      Width           =   5376
      Begin VB.TextBox OT_Sintáxis_Inglés 
         Height          =   285
         HelpContextID   =   11
         Left            =   195
         TabIndex        =   20
         Top             =   3315
         Width           =   4950
      End
      Begin VB.TextBox OT_Ejemplo_Inglés 
         Height          =   285
         HelpContextID   =   11
         Left            =   210
         TabIndex        =   23
         Top             =   3990
         Width           =   4950
      End
      Begin VB.TextBox OT_Ejemplo_Español 
         Height          =   285
         HelpContextID   =   11
         Left            =   225
         MultiLine       =   -1  'True
         TabIndex        =   18
         Top             =   2610
         Width           =   4950
      End
      Begin VB.TextBox OT_Sintáxis_Español 
         Height          =   285
         HelpContextID   =   11
         Left            =   225
         TabIndex        =   16
         Top             =   1935
         Width           =   4950
      End
      Begin VB.TextBox OT_Comando_Español 
         Height          =   285
         HelpContextID   =   11
         Left            =   210
         TabIndex        =   12
         Top             =   540
         Width           =   4950
      End
      Begin VB.TextBox OT_Comando_Inglés 
         Height          =   285
         HelpContextID   =   11
         Left            =   210
         TabIndex        =   14
         Top             =   1215
         Width           =   4950
      End
      Begin VB.TextBox OT_codigo 
         Height          =   285
         Left            =   2328
         TabIndex        =   11
         Text            =   "Text1"
         Top             =   180
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label OE_Sintáxis_Inglés 
         AutoSize        =   -1  'True
         Caption         =   "Sintáxis I&nglés"
         Height          =   195
         Left            =   210
         TabIndex        =   19
         Top             =   3060
         Width           =   1020
      End
      Begin VB.Label OE_Ejemplo_Inglés 
         AutoSize        =   -1  'True
         Caption         =   "Ejemplo In&glés"
         Height          =   195
         Left            =   210
         TabIndex        =   21
         Top             =   3750
         Width           =   1065
      End
      Begin VB.Label OE_Ejemplo_Español 
         AutoSize        =   -1  'True
         Caption         =   "Ejemplo Es&pañol"
         Height          =   195
         Left            =   225
         TabIndex        =   17
         Top             =   2355
         Width           =   1230
      End
      Begin VB.Label OE_Sintáxis_Español 
         AutoSize        =   -1  'True
         Caption         =   "Sintáxis E&spañol"
         Height          =   195
         Left            =   225
         TabIndex        =   15
         Top             =   1710
         Width           =   1185
      End
      Begin VB.Label OE_Comando_Español 
         AutoSize        =   -1  'True
         Caption         =   "Equivalencia/Shortcut"
         Height          =   195
         Left            =   210
         TabIndex        =   22
         Top             =   330
         Width           =   1590
      End
      Begin VB.Label OE_Comando_I 
         AutoSize        =   -1  'True
         Caption         =   "Comando del Servidor"
         Height          =   195
         Left            =   210
         TabIndex        =   13
         Top             =   1020
         Width           =   1560
      End
   End
End
Attribute VB_Name = "OF_Comandos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Lista de arreglo que contiene los datos para actualizar
' un registro
Dim LF_Datos(0 To 8) As Control
'Recordset en el que se manejan los registro a actualizar
Dim LF_Tabla As Recordset

Private Sub PL_Habilitar(LP_habilitar1%, LP_habilitar2%, LP_habilitar3%, LP_habilitar4%, LP_habilitar5%)
' /****************************************************/
' Procedimiento encargado de hacer visible o invisibles
' y activos o inactivos ciertos controles de acuerdo a
' la operación que se esté realizando.
' Ej:  Cuando se agrega un registro, se deshabilitan
' los botones de modificar y  eliminar.
' Además se define el botón OB_Salir como de Salir o
' de Cancelar
' /*****************************************************/

OM_Comando.Visible = LP_habilitar1
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
' /*********************************************************/
' En este evento se carga la lista de tipos de servidores
' actualmente conocidos. Se colocan además los campos a
' actualizar de la tabla comandos en un arreglo que se
' utiliza para tal efecto.  Se definen las propiedades de
' cada campo, como ser si son modificables, si son
' obligatorios, o el mensaje que se desea se despliegue
' cuando un campo obligatorio es dejado en blanco.
' /*********************************************************/

Dim L_Tipos_Servidores As Recordset

' Cargar la lista de tipos de servidores
Set L_Tipos_Servidores = GV_Base_De_Datos.OpenRecordset( _
                        "select descripcion, tipo from tipos", _
                        dbOpenSnapshot)
                         
If MD_Cargar_Lista(L_Tipos_Servidores, OC_Tipos_Servidores) = 0 _
    Then
    PL_Habilitar True, False, False, False, False
    OB_Salir.Caption = "&Salir"
    Exit Sub
Else
    ' Se colocan los campos de un registro a agregar o modificar
    ' en el arreglo de control LF_Datos
    OT_Codigo.Tag = " %@[Codigo]": Set LF_Datos(0) = OT_Codigo
    
    OT_Comando_Español.Tag = _
     " @[Comando_E]^Debe Ingresar Comando en Español^"
     Set LF_Datos(1) = OT_Comando_Español
     
    OT_Comando_Inglés.Tag = _
     " @[Comando_I]^Debe Ingresar Comando en Inglés^"
     Set LF_Datos(2) = OT_Comando_Inglés
    
    OT_Sintáxis_Español.Tag = _
      " @[Sintaxis_E]": Set LF_Datos(3) = OT_Sintáxis_Español
      
    OT_Sintáxis_Inglés.Tag = _
      " @[Sintaxis_I]": Set LF_Datos(4) = OT_Sintáxis_Inglés
      
    OT_Ejemplo_Español.Tag = _
      " @[Ejemplo_E]": Set LF_Datos(5) = OT_Ejemplo_Español
    
    OT_Ejemplo_Inglés.Tag = _
      " @[Ejemplo_I]": Set LF_Datos(6) = OT_Ejemplo_Inglés
      
    OT_Tipo_Servidor.Tag = _
      " &[Tipo_Servidor]": Set LF_Datos(7) = OT_Tipo_Servidor
End If
End Sub
Private Function PL_Validacion()
' /*********************************************************/
' Se revisan los campos de un registro con el propósito de
' encontrar aqellos que deberían contener datos y por el
' contrario se encuentran vacios, caso en el que se despliega
' un mensaje previamente definido en el evento load de la forma.
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
' /*********************************************************/
' Se evita descargar la forma solo si el caption del botón
' OB_salir se encuentra en modo Cancelar.
' /*********************************************************/

If OB_Salir.Caption = "&Cancelar" Then Cancel = 1
End Sub

Private Sub OB_agregar_Click()
' /*********************************************************/
' Evento utilizado para agregar un registro a la tabla de
' comandos.  Esta opción de agregar se realiza en dos pasos.
' Primero se coloca el espacio en blanco para que el usuario
' ingrese los datos deseados . El segundo paso graba a la
' tabla los nuevos datos.
' /*********************************************************/

On Error GoTo Etiqueta_Error:

If OB_Salir.Caption = "&Salir" Then
' Primer paso, seteo de controles y preparación de campos
' para nuevo registro.
    If OC_Tipos_Servidores.ListCount > 0 Then
        PL_Habilitar True, False, True, False, False
        OT_Comando_Español.SetFocus
        Set LF_Tabla = GV_Base_De_Datos.OpenRecordset( _
                        "select * from Comandos " + _
                        " Where Tipo_Servidor = " & _
                        OT_Tipo_Servidor, dbOpenDynaset)
    Else
    End If
Else
' Segundo paso,verificación de datos y grabado de nuevo registro.
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
' /********************************************************/
' Evento utilizado para eliminar un registro de la tabla de
' comandos.  Esta opción se realiza de la siguiente forma:
' El usuario indica el registro a eliminar, luego se
' pregunta al mismo si desea realmente eliminarlo.  De
' contestar que si se elimina el registro de la base de
' datos.
' /********************************************************/

On Error GoTo Etiqueta_Error:

If OC_Tipos_Servidores.ListCount > 0 Then
    If OL_Comandos.ListIndex > -1 Then
      If MG_Pregunta("¿Desea Eliminar Registro?") = vbYes Then
       Set LF_Tabla = GV_Base_De_Datos.OpenRecordset( _
                  "select * from Comandos where " + _
                  " Codigo= " & _
                  OL_Comandos.ItemData(OL_Comandos.ListIndex), _
                   dbOpenDynaset)
       LF_Tabla.Delete
       OL_Comandos.RemoveItem (OL_Comandos.ListIndex)
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

Private Sub OB_Modificar_Click()
' /********************************************************/
' evento utilizado para modificar un registro de la tabla
' de comandos.  Esta opción se realiza en dos pasos.
' Primero se recupera el registro a modificar y se presentan
' los datos al usuario para que este realize las
' modificaciones deseadas. El segundo paso registra las
' modificaciones en la base de datos.
' /********************************************************/

On Error GoTo Etiqueta_Error:

If OB_Salir.Caption = "&Salir" Then
' Primer paso, recuperación de registro.
    If OC_Tipos_Servidores.ListCount > 0 Then
        If OL_Comandos.ListIndex > -1 Then
            
            Set LF_Tabla = GV_Base_De_Datos.OpenRecordset( _
            "select * from Comandos where Codigo= " & _
            OL_Comandos.ItemData(OL_Comandos.ListIndex), _
            dbOpenDynaset)
             
            MD_Cargar_Datos LF_Datos, LF_Tabla, 4
            PL_Habilitar True, False, False, True, False
            OT_Comando_Español.SetFocus
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
' Este evento se utiliza para cargar los comandos
' existentes de acuerdo al tipo de servidor seleccionado.
' /********************************************************/
Dim L_Comandos As Recordset

If OC_Tipos_Servidores.ListCount > 0 Then
    
    Set L_Comandos = GV_Base_De_Datos.OpenRecordset( _
    "select Comando_E + '/' + Comando_I as Comando," _
    + " codigo from Comandos where tipo_servidor = " _
    & OC_Tipos_Servidores.ItemData( _
    OC_Tipos_Servidores.ListIndex), _
    dbOpenSnapshot)
    
    MD_Cargar_Lista L_Comandos, OL_Comandos
    
    OT_Tipo_Servidor = OC_Tipos_Servidores.ItemData( _
                       OC_Tipos_Servidores.ListIndex)
Else
    OT_Tipo_Servidor = 0
End If
End Sub
