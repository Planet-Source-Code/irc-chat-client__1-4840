VERSION 5.00
Begin VB.Form OF_Canales_Preferidos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Canales Favoritos"
   ClientHeight    =   5850
   ClientLeft      =   135
   ClientTop       =   1845
   ClientWidth     =   4920
   HelpContextID   =   27
   Icon            =   "OF_CaPre.frx":0000
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5850
   ScaleWidth      =   4920
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame OM_Marco2 
      Caption         =   "Canales"
      Height          =   2784
      HelpContextID   =   27
      Left            =   132
      TabIndex        =   16
      Top             =   2472
      Width           =   4692
      Begin VB.TextBox OT_Canal 
         Height          =   300
         HelpContextID   =   27
         Left            =   264
         TabIndex        =   19
         Top             =   768
         Visible         =   0   'False
         Width           =   1524
      End
      Begin VB.TextBox OT_Descripcion 
         Height          =   300
         HelpContextID   =   27
         Left            =   264
         TabIndex        =   21
         Top             =   1428
         Visible         =   0   'False
         Width           =   4248
      End
      Begin VB.CommandButton OB_Ok 
         Caption         =   "&Aceptar"
         Height          =   324
         HelpContextID   =   27
         Left            =   1272
         TabIndex        =   22
         Top             =   1980
         Visible         =   0   'False
         Width           =   876
      End
      Begin VB.CommandButton OB_Cancelar 
         Caption         =   "&Cancelar"
         Height          =   324
         HelpContextID   =   27
         Left            =   2160
         TabIndex        =   23
         Top             =   1980
         Visible         =   0   'False
         Width           =   876
      End
      Begin VB.ListBox OL_Canales 
         Height          =   2205
         Left            =   135
         MultiSelect     =   2  'Extended
         Sorted          =   -1  'True
         TabIndex        =   17
         Top             =   345
         Width           =   4392
      End
      Begin VB.Label OE_Descripcion 
         AutoSize        =   -1  'True
         Caption         =   "Descripción"
         Height          =   192
         Left            =   276
         TabIndex        =   20
         Top             =   1188
         Visible         =   0   'False
         Width           =   864
      End
      Begin VB.Label OE_Canal 
         AutoSize        =   -1  'True
         Caption         =   "Canal"
         Height          =   228
         Left            =   264
         TabIndex        =   18
         Top             =   468
         Visible         =   0   'False
         Width           =   420
      End
   End
   Begin VB.Frame OM_Marco1 
      Caption         =   "Servidores"
      Height          =   2310
      HelpContextID   =   27
      Left            =   150
      TabIndex        =   0
      Top             =   90
      Width           =   4692
      Begin VB.TextBox OT_Indice 
         Height          =   285
         Index           =   5
         Left            =   60
         TabIndex        =   15
         Top             =   1800
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.TextBox OT_Indice 
         Height          =   285
         Index           =   4
         Left            =   84
         TabIndex        =   14
         Top             =   1380
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.TextBox OT_Indice 
         Height          =   285
         Index           =   3
         Left            =   36
         TabIndex        =   13
         Top             =   1032
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.TextBox OT_Indice 
         Height          =   285
         Index           =   2
         Left            =   60
         TabIndex        =   12
         Top             =   648
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.TextBox OT_Indice 
         Height          =   285
         Index           =   1
         Left            =   72
         TabIndex        =   11
         Top             =   264
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.CheckBox OCH_Cerrar 
         Height          =   195
         HelpContextID   =   27
         Index           =   5
         Left            =   4140
         TabIndex        =   10
         Top             =   1935
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.CheckBox OCH_Cerrar 
         Height          =   195
         HelpContextID   =   27
         Index           =   4
         Left            =   4140
         TabIndex        =   9
         Top             =   1500
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.CheckBox OCH_Cerrar 
         Height          =   195
         HelpContextID   =   27
         Index           =   3
         Left            =   4140
         TabIndex        =   8
         Top             =   1080
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.CheckBox OCH_Cerrar 
         Height          =   195
         HelpContextID   =   27
         Index           =   2
         Left            =   4140
         TabIndex        =   7
         Top             =   713
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.CheckBox OCH_Cerrar 
         Height          =   195
         HelpContextID   =   27
         Index           =   1
         Left            =   4140
         TabIndex        =   6
         Top             =   312
         Visible         =   0   'False
         Width           =   255
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
         Height          =   360
         Index           =   5
         Left            =   276
         TabIndex        =   5
         Top             =   1845
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
         Height          =   360
         Index           =   4
         Left            =   276
         TabIndex        =   4
         Top             =   1440
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
         Height          =   360
         Index           =   3
         Left            =   276
         TabIndex        =   3
         Top             =   1020
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
         Height          =   360
         Index           =   2
         Left            =   276
         TabIndex        =   2
         Top             =   630
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
         Height          =   315
         Index           =   1
         Left            =   276
         TabIndex        =   1
         Top             =   252
         Visible         =   0   'False
         Width           =   3700
      End
   End
   Begin VB.CommandButton OB_Borrar 
      Caption         =   "&Borrar"
      Height          =   336
      HelpContextID   =   27
      Left            =   1020
      TabIndex        =   25
      Top             =   5388
      Width           =   864
   End
   Begin VB.CommandButton OB_Agregar 
      Caption         =   "&Agregar"
      Height          =   336
      HelpContextID   =   27
      Left            =   120
      TabIndex        =   24
      Top             =   5388
      Width           =   864
   End
   Begin VB.CommandButton OB_Join 
      Caption         =   "&Entrar/Join"
      Height          =   336
      HelpContextID   =   27
      Left            =   2856
      TabIndex        =   26
      Top             =   5388
      Width           =   1008
   End
   Begin VB.CommandButton OB_Salir 
      Caption         =   "&Salir"
      Height          =   336
      HelpContextID   =   27
      Left            =   3888
      TabIndex        =   27
      Top             =   5388
      Width           =   864
   End
End
Attribute VB_Name = "OF_Canales_Preferidos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' Variable local a la forma que nos indica la acción en que se
' encuentra la ventana cuando el usuario hace click ya sea en
' cualquiera de los botones de AGREGAR, BORRAR
Dim LF_Accion$

Sub PL_Cargar_Favoritos()
' /*********************************************************/
' Procedimiento local a la forma que carga los canales
' favoritos del usuario a la Lista OL_CANALES de la forma.
' Los canales favoritos estan registrados en la tabla de
' CANALES_FAVORITOS.
' /*********************************************************/
On Error GoTo Etiqueta_Error:
Dim L_Registro As Recordset
Dim L_Canal$
OL_Canales.Clear ' Limpiar la lista
' Cargar la tabla
Set L_Registro = GV_Base_De_Datos.OpenRecordset("Select " + _
                    "* from  CANALES_FAVORITOS")

While Not L_Registro.EOF ' Mientras haya registros
  L_Canal = Trim(L_Registro!Nombre)
  MG_Rellena_Espacios L_Canal, 35
  ' Agregar a la lista
  OL_Canales.AddItem L_Canal + Chr(9) + ":" + _
  L_Registro!Descripcion
  L_Registro.MoveNext ' Moverse al siguiente registro
                      ' en la tabla
Wend

L_Registro.Close ' Cerrar la tabla

Exit Sub
Etiqueta_Error:
ME_Muestra_Error
End Sub

Sub PL_Cargar_Servidores_Activos(LP_Cual%)
' /*********************************************************/
' Este procedimiento carga todos los servidores activos en
' el momento Los servidores son cargardos, para que el usuario
' tenga opción a ingresar a varios canales en diferentes
' servidores. Los servidores activos son cargados del arreglo
' global de Sockets
' /*********************************************************/

Dim L_i As Integer
Dim L_libre As Integer

' Maneja el numero asociado al indice en el arreglo de Sockets
L_libre = 1

For L_i = 1 To 5 ' Recorrer el arreglo de Sockets
  If GV_Sockets(L_i).socket <> INVALID_SOCKET Then
     OT_Indice(L_libre) = L_i
     
     ' Mostrar el servidor en la forma
     OT_Nombre(L_libre).Visible = True
     
     ' Mostrar el Check Box
     OCH_Cerrar(L_libre).Visible = True
     
     OT_Nombre(L_libre) = GV_Sockets(L_i).Direcc + "(" _
                        + CStr(GV_Sockets(L_i).Puerto) + _
                        ") NICK => " + GV_Sockets(L_i).Nick
                          
     If L_i = LP_Cual Then OCH_Cerrar(L_libre).Value = 1
     L_libre = L_libre + 1
  
  End If

Next L_i

End Sub

Sub PL_Deshabilitar_Habilitar(LP_Estado As Boolean)
' /********************************************************/
' Este procedimiento habilita o deshabilita objetos para
' cuando el usuario agrega un canal favorito
' /********************************************************/

If LP_Estado Then
    OB_Agregar.Enabled = False
    OB_Borrar.Enabled = False
    OB_Join.Enabled = False
    OB_Salir.Enabled = False
    OM_Marco1.Enabled = False
    OL_Canales.Visible = False
    OE_Canal.Visible = True
    OE_Descripcion.Visible = True
    OT_Canal.Visible = True
    OT_Canal.Enabled = True
    OT_Canal.SetFocus
    OT_Descripcion.Visible = True
    OT_Descripcion.Enabled = True
    OT_Canal = "": OT_Descripcion = ""
    OB_Ok.Visible = True
    OB_Ok.Enabled = True
    OB_Cancelar.Visible = True
    OB_Cancelar.Enabled = True
    
    OB_Ok.Default = True
    
Else
    OB_Agregar.Enabled = True
    OB_Borrar.Enabled = True
    OB_Join.Enabled = True
    OB_Salir.Enabled = True
    OM_Marco1.Enabled = True
    OM_Marco2.Enabled = True
    OL_Canales.Visible = True
    OL_Canales.Enabled = True
    OE_Canal.Visible = False
    OE_Descripcion.Visible = False
    OT_Canal.Visible = False
    OT_Descripcion.Visible = False
    OB_Ok.Visible = False
    OB_Cancelar.Visible = False
    

End If
End Sub

Sub PL_Ejecutar_Join(LP_Cual%)
' /********************************************************/
' Este procedimiento envia el comando JOIN al socket que
' viene de parametro en LP_CUAL, LP_CUAL no es en si el
' socket sino: GV_Sockets(LP_CUAL).socket.
' LP_CUAL solo representa el indice de donde se encuentra
' el socket.
' El ccomando es enviado al socket para todos los canales
' seleccionados en la lista de CANALES PREFERIDOS
' /********************************************************/
Dim L_i As Integer
Dim L_Canal$
For L_i = 0 To OL_Canales.ListCount - 1
  
  ' Si el canal esta seleccionado
  If OL_Canales.Selected(L_i) Then
    L_Canal = Left(OL_Canales.List(L_i), _
              InStr(1, OL_Canales.List(L_i), " ") - 1)
              
    L_Canal = Trim(L_Canal)
    ' Enviar Mensaje al Servidor , Socket en la Posición:
    ' OT_Indice(LP_Cual)del arreglo de Sockets
    MM_Enviar_Mensaje "JOIN " + L_Canal, OT_Indice(LP_Cual)
    
    
  End If
  DoEvents
Next L_i

End Sub

Sub PL_Unirse_a_Canales()
' /*********************************************************/
' Este procedimiento recorre los servidores marcados en el
' checboxpara los cuales para cada uno se llama al
' PROCEDIMIENTO PL_Ejecutar_JOIN con el socket que representa
' el servidor marcado en el CHECKBOX
' /*********************************************************/
Dim L_i As Integer
Dim L_libre As Integer

L_libre = 1

For L_i = 1 To 5
     ' Si el CHECKBOX esta marcado
     If OCH_Cerrar(L_i).Value = 1 Then
       PL_Ejecutar_Join L_i ' Ejecutar el JOIN para ese socket
     End If

Next L_i
Unload Me
End Sub

Private Sub Form_Activate()
' /*********************************************************/
' Cada vez que la forma se active, que se llame nuevamente
' al procedimiento de cargar servidores activos, esto se hace
' porque la ventana puede permanecer abierta, y mas de algún
' servidor se puede cerrar. La ventana entonces no reflejará
' correctamente los servidores Activos. No garantiza que
' funcionará al 100% pero si minimizará el error.
' La variable Global GV_Seleccion, posee el indice de donde
' se llamo la ventana de CANALES PREFERIDOS, esto le sirve al
' procedimiento para marcar el CHECKBOX de ese servidor
' (GV_Seleccion) por omisión
' /*********************************************************/
PL_Cargar_Servidores_Activos GV_Seleccion

End Sub

Private Sub Form_Load()
' /*********************************************************/
' Cada vez que la forma se cargue , se llama a el
' procedimiento que carga los servidores activos, marcando
' el CHECKBOX del servidor por omisión el cual esta
' determinado por la variable global GV_Seleccion.
' Ademas se deben cargar los canales favoritos del usuario
' ya registrados
' /*********************************************************/

PL_Cargar_Servidores_Activos GV_Seleccion
PL_Cargar_Favoritos

End Sub

Private Sub OB_agregar_Click()
' /*********************************************************/
' Cuando se hace click en el boton de agregar entonces la
' acción se registra como agregar y se deshabilitan todos
' los objetos no relacionados con esta acción.
' /*********************************************************/
 
 LF_Accion = "AGREGAR"
 PL_Deshabilitar_Habilitar True
End Sub

Private Sub OB_Borrar_Click()
' /*********************************************************/
' Cuando se hace click en el boton de borrar se llama al
' procedimiento de Borrar Canales Favoritos
' /*********************************************************/

PL_Borrar_Canales
End Sub

Private Sub OB_Cancelar_Click()
' /*********************************************************/
' Cuando se hace click en el boton de Cancelar.. Este botón
' solo esta activo cuando se esta agregando un canal
' /*********************************************************/

LF_Accion = ""
PL_Deshabilitar_Habilitar False
End Sub

Private Sub OB_Join_Click()
' /*********************************************************/
' Cuando se hace click en el boton Join. Este botón es
' utilizado para especificar al programa que envie los
' mensajes necesarios a los diferentes servidores
' seleccionados para unirse a varios canales
' /*********************************************************/

PL_Unirse_a_Canales
End Sub

Private Sub OB_OK_Click()
' /*********************************************************/
' Cuando se hace click en el boton Ok. Este botón es
' utilizado para confirmar que se agregará un nuevo canal
' a la Tabla de CANALES FAVORITOS
' /*********************************************************/

On Error GoTo Etiqueta_Error:

Dim L_Registro As Recordset
Dim L_Canal$
If LF_Accion = "AGREGAR" Then ' Si la acción es agregar
  ' Abre la tabla
  Set L_Registro = GV_Base_De_Datos.OpenRecordset( _
                                    "CANALES_Favoritos")
  L_Registro.AddNew
  L_Registro!Nombre = OT_Canal
  L_Registro!Descripcion = OT_Descripcion
  L_Registro.Update ' Agrega el Registro
  L_Registro.Close
  L_Canal = Trim(OT_Canal)
  MG_Rellena_Espacios L_Canal, 35
  'Agrega el nuevo Canal a la lista de Canales
  OL_Canales.AddItem L_Canal + Chr(9) + ":" + OT_Descripcion
  PL_Deshabilitar_Habilitar False

End If

Exit Sub
Etiqueta_Error:
ME_Muestra_Error
End Sub

Private Sub OB_Salir_Click()
' Descarga la ventana
Unload Me
End Sub


Private Sub OL_Canales_DblClick()
' /*********************************************************/
' Cuando se hace doble click en la lista de Canales, se
' llama a la función de unirse a canales con el canal
' seleccionado de la lista
' /*********************************************************/

PL_Unirse_a_Canales
Unload Me
End Sub

Private Sub OL_Canales_KeyDown(KeyCode As Integer, Shift As Integer)
' /*********************************************************/
' Cuando se presiona la tecla DELETE estando en la lista
' Se llama al procedimiento que borra canales de la Tabla de
' CANALES FAVORITOS, Si se presiona la tecla ENTER se llama
' al procedimiento que ejecuta los JOINS a los diferentes
' canales
' /*********************************************************/

If KeyCode = vbKeyDelete Then
  PL_Borrar_Canales
ElseIf KeyCode = vbKeyReturn Then
 PL_Unirse_a_Canales
End If
End Sub

Sub PL_Borrar_Canales()
' /*********************************************************/
' Procedimiento que se encarga de borrar de la tabla de
' CANALES FAVORITOS los canales seleccionados de la lista
' /*********************************************************/

Dim L_i As Integer
Dim L_Registro As Recordset
Dim L_Canal$
' Abrir la tabla
Set L_Registro = GV_Base_De_Datos.OpenRecordset( _
                          "CANALES_FAVORITOS", dbOpenDynaset)
If Not L_Registro.EOF Then
    For L_i = 0 To (OL_Canales.ListCount - 1)
      If OL_Canales.Selected(L_i) Then ' Si esta seleccionado
        L_Canal = Left(OL_Canales.List(L_i), _
                  InStr(1, OL_Canales.List(L_i), " ") - 1)
        L_Canal = Trim(L_Canal)
        ' Busquelo en la tabla
        L_Registro.FindFirst "Nombre='" + L_Canal + "'"
        If Not L_Registro.NoMatch Then
          L_Registro.Delete ' Borre el registro
          
        End If
        
      End If
    DoEvents
    Next L_i

End If
L_Registro.Close ' Cerrar la tabla
PL_Cargar_Favoritos ' Cargar nuevamente la lista de Canales
OL_Canales.Refresh

End Sub

