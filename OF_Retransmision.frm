VERSION 5.00
Begin VB.Form OF_Retransmision 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Retransmisión de Mensajes"
   ClientHeight    =   4905
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9390
   HelpContextID   =   26
   Icon            =   "OF_Retransmision.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4905
   ScaleWidth      =   9390
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton OB_salir 
      Caption         =   "&Salir"
      Height          =   348
      HelpContextID   =   26
      Left            =   8205
      TabIndex        =   17
      Top             =   4395
      Width           =   975
   End
   Begin VB.CommandButton OB_Agregar 
      Caption         =   "&Agregar"
      Height          =   348
      HelpContextID   =   26
      Left            =   255
      TabIndex        =   14
      Top             =   4425
      Width           =   972
   End
   Begin VB.CommandButton OB_Modificar 
      Caption         =   "&Modificar"
      Height          =   348
      HelpContextID   =   26
      Left            =   1335
      TabIndex        =   15
      Top             =   4425
      Width           =   975
   End
   Begin VB.CommandButton OB_Eliminar 
      Caption         =   "&Eliminar"
      Height          =   348
      HelpContextID   =   26
      Left            =   2430
      TabIndex        =   16
      Top             =   4425
      Width           =   975
   End
   Begin VB.Frame OM_Servidores_Emisores 
      Caption         =   "Retransmisiones"
      Height          =   4170
      HelpContextID   =   26
      Left            =   240
      TabIndex        =   10
      Top             =   90
      Width           =   8940
      Begin VB.ComboBox OC_Servidor_Emisor 
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
         HelpContextID   =   26
         Left            =   1965
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   450
         Width           =   5160
      End
      Begin VB.ListBox OL_Estructuras_Registradas 
         Height          =   2790
         HelpContextID   =   26
         ItemData        =   "OF_Retransmision.frx":0442
         Left            =   120
         List            =   "OF_Retransmision.frx":0444
         MultiSelect     =   1  'Simple
         TabIndex        =   13
         Top             =   1050
         Width           =   8685
      End
      Begin VB.Label OE_Servidores_Emisor 
         AutoSize        =   -1  'True
         Caption         =   "Servidor &Emisor"
         Height          =   195
         Index           =   0
         Left            =   390
         TabIndex        =   11
         Top             =   540
         Width           =   1095
      End
   End
   Begin VB.Frame OM_Datos_Transmision 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4155
      Left            =   255
      TabIndex        =   0
      Top             =   105
      Visible         =   0   'False
      Width           =   8910
      Begin VB.CheckBox OCH_Omitir 
         Caption         =   "Omitir Mensaje Original"
         Height          =   285
         HelpContextID   =   26
         Left            =   5445
         TabIndex        =   18
         Top             =   3435
         Width           =   2505
      End
      Begin VB.TextBox OT_Alias_Emisor 
         Height          =   315
         HelpContextID   =   26
         Left            =   345
         TabIndex        =   2
         Top             =   540
         Width           =   1410
      End
      Begin VB.CheckBox OCH_Solo_Mensaje 
         Caption         =   "Transmitir Solamente Mensaje De Usuario"
         Height          =   285
         HelpContextID   =   26
         Left            =   1230
         TabIndex        =   9
         Top             =   3450
         Width           =   3495
      End
      Begin VB.TextBox OT_Prefijo 
         Height          =   570
         HelpContextID   =   26
         Left            =   345
         MultiLine       =   -1  'True
         TabIndex        =   6
         Top             =   1395
         Width           =   8370
      End
      Begin VB.TextBox OT_Sufijo 
         Height          =   585
         HelpContextID   =   26
         Left            =   345
         MultiLine       =   -1  'True
         TabIndex        =   8
         Top             =   2415
         Width           =   8340
      End
      Begin VB.ComboBox OC_Servidor_Receptor 
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
         HelpContextID   =   26
         Left            =   2355
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   540
         Width           =   4815
      End
      Begin VB.Label OE_Alias_Emisor 
         AutoSize        =   -1  'True
         Caption         =   "&Alias De Emisor"
         Height          =   195
         Left            =   345
         TabIndex        =   1
         Top             =   330
         Width           =   1095
      End
      Begin VB.Label OE_Servidor_Receptor 
         AutoSize        =   -1  'True
         Caption         =   "Servidor &Receptor"
         Height          =   195
         Index           =   1
         Left            =   2415
         TabIndex        =   3
         Top             =   330
         Width           =   1290
      End
      Begin VB.Label OE_Prefijo 
         AutoSize        =   -1  'True
         Caption         =   "&Prefijo"
         Height          =   195
         Left            =   345
         TabIndex        =   5
         Top             =   1155
         Width           =   435
      End
      Begin VB.Label OE_Sufijo 
         AutoSize        =   -1  'True
         Caption         =   "&Sufijo"
         Height          =   195
         Left            =   345
         TabIndex        =   7
         Top             =   2175
         Width           =   390
      End
   End
End
Attribute VB_Name = "OF_Retransmision"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Function PL_Buscar_Servidor(LP_Cual%) As Integer
' /***********************************************/
' Esta función se encarga de buscar el índice en
' el combo del servidor asociado al socket LP_cual.
' /***********************************************/
Dim L_i As Integer
Dim L_libre As Integer

L_libre = 1

For L_i = 0 To OC_Servidor_Emisor.ListCount - 1
  If OC_Servidor_Emisor.ItemData(L_i) = LP_Cual Then
     PL_Buscar_Servidor = L_i
  End If
Next L_i



End Function

Private Sub Form_Load()
' /*************************************************/
' En este evento se cargan los servidores activos
' que pueden ser emisores y receptores.
' /*************************************************/

PL_Cargar_Servidores_Activos OC_Servidor_Emisor
PL_Cargar_Servidores_Activos OC_Servidor_Receptor
End Sub

Private Function PL_Validacion() As Boolean
' /************************************************/
' Se revisan los campos de de la estructura volátil
' gv_estructura_retransmisión
' con el propósito de encontrar aqellos campos que
' deberían contener datos y por el
' contrario se encuentran vacios, caso en el que
' se despliega un mensaje definido
' en este evento.
' /************************************************/

Dim L_i%

If Trim(OT_Alias_Emisor.Text) = "" Then
    MG_Mensaje "Debe Ingresar Alias"
    OT_Alias_Emisor.SetFocus
    Exit Function
End If

If OC_Servidor_Receptor.ListIndex = -1 Then
    MG_Mensaje "Debe Seleccionar Servidor Receptor"
    OC_Servidor_Receptor.SetFocus
    Exit Function
End If
PL_Validacion = -1
End Function

Sub PL_Cargar_Servidores_Activos(LP_Combo As ComboBox)
' /***************************************************/
' Este procedimiento carga todos los servidores
' activos en el momento.
' Los servidores son cargados, para que el usuario
' tenga opción de retransmitir a cualquiera de los
' diferentes servidores activos o a un usuario especifico.
' Los servidores activos son cargados del arreglo
' global de Sockets
' /****************************************************/

Dim L_i As Integer
Dim L_libre As Integer

L_libre = 1

For L_i = 1 To 5
  If GV_Sockets(L_i).socket <> INVALID_SOCKET Then
     LP_Combo.AddItem GV_Sockets(L_i).Direcc + "(" + _
     CStr(GV_Sockets(L_i).Puerto) + _
     ") NICK => " + GV_Sockets(L_i).Nick
     LP_Combo.ItemData(L_libre - 1) = L_i
     L_libre = L_libre + 1
  End If
Next L_i
If LP_Combo.ListCount > 0 Then LP_Combo.ListIndex = 0
End Sub

Sub PL_Limpiar()
' /****************************************************/
' Utilizado para limpiar los campos utilizados para
' manejar los datos de la forma.
' /****************************************************/

OT_Prefijo = ""
OT_Sufijo = ""
OT_Alias_Emisor = ""
OC_Servidor_Receptor.ListIndex = -1
OCH_Solo_Mensaje.Value = 0
OCH_Omitir.Value = 0
End Sub

Sub PL_Cargar_Lista_Retransmision(LP_Servidor As Integer)
' /******************************************************/
' Este procedimiento se utiliza para cargar la lista de
' posibles servidores
' emisores.
' /*****************************************************/

Dim L_i%, L_Cuantos%
Dim L_Linea$

' No de Estructuras
L_Cuantos = UBound(GV_Estructura_Retransmision)

OL_Estructuras_Registradas.Clear
For L_i = 1 To L_Cuantos
    If GV_Estructura_Retransmision(L_i).E_Servidor_Emisor _
    = LP_Servidor And _
    GV_Estructura_Retransmision(L_i).E_Borrado = False Then
       
       L_Linea = _
       "Usuario Emisor :[" + _
       GV_Estructura_Retransmision(L_i).E_Alias_Emisor _
       + "]   Servidor Receptor:[" + _
       GV_Sockets(GV_Estructura_Retransmision( _
       L_i).E_Servidor_Receptor).Direcc + _
       "]    Prefijo:[" + GV_Estructura_Retransmision( _
       L_i).E_Prefijo + "]    Sufijo:[" + _
       GV_Estructura_Retransmision(L_i).E_Sufijo + " ]"
           
           
       OL_Estructuras_Registradas.AddItem L_Linea
       
     
       OL_Estructuras_Registradas.ItemData( _
       OL_Estructuras_Registradas.NewIndex) = L_i
    End If
Next L_i

End Sub

Private Sub Form_Unload(Cancel As Integer)
' /***********************************************/
' Se evita descargar la forma solo si el caption
' del botón OB_salir se encuentra
' en modo Cancelar.
' /***********************************************/

If OB_salir.Caption = "&Cancelar" Then Cancel = 1
End Sub

Public Sub OB_agregar_Click()
' /************************************************/
' Evento utilizado para modificar un registro de la
' estructura volátil gv_estructura_retransmisión.
' Esta opción se realiza en dos pasos.  Primero se
' coloca el espacio en blanco para que el usuario
' ingrese los datos deseados . El segundo paso
' graba los nuevos datos en la estructura.
' /***********************************************/

Dim L_Cuantos%

If OB_salir.Caption = "&Salir" Then
' Primer paso, seteo de controles y preparación
' de campos para nuevo registro.
    PL_Limpiar
    If OC_Servidor_Emisor.ListCount = 0 Then Exit Sub
    PL_Habilitar True, False, True, False, False
  
Else
' Segundo paso, verificación de datos y grabado de
' nuevo registro en estructura.
    If Not PL_Validacion() Then Exit Sub
    L_Cuantos = PL_Indice_Libre_Retransmision
    ReDim Preserve GV_Estructura_Retransmision(L_Cuantos)
    GV_Estructura_Retransmision( _
    L_Cuantos).E_Alias_Emisor = Trim(OT_Alias_Emisor)
    GV_Estructura_Retransmision( _
    L_Cuantos).E_Servidor_Emisor = _
    OC_Servidor_Emisor.ItemData(OC_Servidor_Emisor.ListIndex)
    GV_Estructura_Retransmision( _
    L_Cuantos).E_Servidor_Receptor = _
    OC_Servidor_Receptor.ItemData(OC_Servidor_Receptor.ListIndex)
    If OCH_Solo_Mensaje.Value = 0 Then
     GV_Estructura_Retransmision(L_Cuantos).E_SoloMensaje = False
    Else
     GV_Estructura_Retransmision(L_Cuantos).E_SoloMensaje = True
    End If
    
    If OCH_Omitir.Value = 0 Then
     GV_Estructura_Retransmision(L_Cuantos).E_Omitir_Mensaje = _
      False
    Else
     GV_Estructura_Retransmision(L_Cuantos).E_Omitir_Mensaje = _
     True
    End If


    GV_Estructura_Retransmision(L_Cuantos).E_Prefijo = _
    OT_Prefijo
    GV_Estructura_Retransmision(L_Cuantos).E_Sufijo = _
    OT_Sufijo
    OC_Servidor_Emisor_Click
    PL_Habilitar False, True, True, True, True
End If

End Sub

Private Sub OB_Eliminar_Click()
' /****************************************************/
' Evento utilizado para eliminar un registro de de la
' estructura volátil gv_estructura_retransmisión.
' Esta opción se realiza de la siguiente forma:  El
' usuario indica los registro a eliminar, luego presiona
' el botón OB_eliminar para borrar los registro de la
' estructura.
' /****************************************************/

Dim L_i%
If OL_Estructuras_Registradas.ListCount = 0 Then Exit Sub
For L_i = 0 To OL_Estructuras_Registradas.ListCount - 1
    If OL_Estructuras_Registradas.Selected(L_i) Then
        GV_Estructura_Retransmision( _
        OL_Estructuras_Registradas.ItemData(L_i)).E_Borrado _
        = True
    End If
Next L_i
OC_Servidor_Emisor_Click
End Sub
Private Sub OB_Modificar_Click()
' /******************************************************/
' Evento utilizado para modificar un registro de la
' estructura volátil gv_estructura_retransmisión.
' Esta opción se realiza en dos pasos.  Primero se
' recupera el registro a modificar y  se presentan los
' datos al usuario para que este realize las
' modificaciones deseadas. El segundo paso registra
' las modificaciones en la estructura.
' /******************************************************/

Dim L_Cuantos%

If OB_salir.Caption = "&Salir" Then
' Primer paso, recuperación de registro.
    If OC_Servidor_Emisor.ListCount = 0 Then Exit Sub
    If OL_Estructuras_Registradas.ListCount < 1 Then Exit Sub
    PL_Habilitar True, False, False, True, False
    L_Cuantos = OL_Estructuras_Registradas.ItemData( _
    OL_Estructuras_Registradas.ListIndex)
    
    OT_Alias_Emisor = GV_Estructura_Retransmision( _
     L_Cuantos).E_Alias_Emisor
    
    OC_Servidor_Receptor.ListIndex = _
    PL_Buscar_Servidor(GV_Estructura_Retransmision( _
    L_Cuantos).E_Servidor_Receptor)
    If GV_Estructura_Retransmision( _
    L_Cuantos).E_SoloMensaje = False Then
      OCH_Solo_Mensaje.Value = 0
    Else
        OCH_Solo_Mensaje.Value = 1
    End If
    
    If GV_Estructura_Retransmision( _
     L_Cuantos).E_Omitir_Mensaje = False Then
         OCH_Omitir.Value = 0
    Else
        OCH_Omitir.Value = 1
    End If
    
    OT_Prefijo = GV_Estructura_Retransmision( _
      L_Cuantos).E_Prefijo
    OT_Sufijo = GV_Estructura_Retransmision( _
     L_Cuantos).E_Sufijo
    OT_Alias_Emisor.SetFocus
Else
' Segundo paso, modificación de registro.
    If Not PL_Validacion() Then Exit Sub
    L_Cuantos = OL_Estructuras_Registradas.ItemData( _
     OL_Estructuras_Registradas.ListIndex)
    GV_Estructura_Retransmision(L_Cuantos).E_Alias_Emisor _
     = Trim(OT_Alias_Emisor)
    GV_Estructura_Retransmision(L_Cuantos).E_Servidor_Emisor _
    = OC_Servidor_Emisor.ItemData(OC_Servidor_Emisor.ListIndex)
    GV_Estructura_Retransmision(L_Cuantos).E_Servidor_Receptor _
     = OC_Servidor_Receptor.ItemData( _
     OC_Servidor_Receptor.ListIndex)
    If OCH_Solo_Mensaje.Value = 0 Then
      GV_Estructura_Retransmision(L_Cuantos).E_SoloMensaje _
       = False
    Else
      GV_Estructura_Retransmision(L_Cuantos).E_SoloMensaje _
       = True
    End If
    If OCH_Omitir.Value = 0 Then
     GV_Estructura_Retransmision(L_Cuantos).E_Omitir_Mensaje _
      = False
    Else
     GV_Estructura_Retransmision(L_Cuantos).E_Omitir_Mensaje _
     = True
    End If
    GV_Estructura_Retransmision(L_Cuantos).E_Prefijo = _
     Trim(OT_Prefijo)
    GV_Estructura_Retransmision(L_Cuantos).E_Sufijo = _
     Trim(OT_Sufijo)
    OC_Servidor_Emisor_Click
    PL_Habilitar False, True, True, True, True
    PL_Limpiar
End If
End Sub

Private Sub OB_Salir_Click()
' /*****************************************************/
' Este evento se utiliza para cerrar la forma o para
' cancelar la operación (agregar, modificar, eliminar)
' que se estaba realizando.
' /*****************************************************/

If OB_salir.Caption = "&Salir" Then
    Unload Me
Else
    PL_Limpiar
    PL_Habilitar False, True, True, True, True
End If
End Sub

Private Sub OC_Servidor_Emisor_Click()
' /*****************************************************/
' Este evento llama al procedimiento
' PL_Cargar_Lista_Retransmision.
' /****************************************************/

PL_Cargar_Lista_Retransmision ( _
OC_Servidor_Emisor.ItemData(OC_Servidor_Emisor.ListIndex))
End Sub

Private Sub PL_Habilitar(LP_habilitar1%, LP_habilitar2%, LP_habilitar3%, LP_habilitar4%, LP_habilitar5%)
' /******************************************************/
' Procedimiento encargado de hacer visible o invisibles
' y activos o inactivos ciertos controles de acuerdo a
' la operación que se esté realizando.
' Ej:  Cuando se agrega un registro, se deshabilitan los
' botones de modificar y  eliminar.
' Además se define el botón OB_Salir como de Salir o de
' Cancelar
' /******************************************************/

OM_Datos_Transmision.Visible = LP_habilitar1
OM_Servidores_Emisores.Visible = LP_habilitar2
OB_Agregar.Enabled = LP_habilitar3
OB_Modificar.Enabled = LP_habilitar4
OB_Eliminar.Enabled = LP_habilitar5

If OB_salir.Caption = "&Salir" Then
    OB_salir.Caption = "&Cancelar"
Else
    OB_salir.Caption = "&Salir"
End If
End Sub

Function PL_Indice_Libre_Retransmision() As Integer
' /********************************************************/
' Función que se encarga de buscar una casilla libre en el
' arreglo de  Ventanas de canales
' /********************************************************/
   
   Dim L_i As Integer
   Dim L_ArrayCount As Integer

    L_ArrayCount = UBound(GV_Estructura_Retransmision)

    ' Recorrer el arreglo de ventanas. Si la ventana ha sido
    ' borrada  entonces, retorna ese indice.
    For L_i = 1 To L_ArrayCount
        If GV_Estructura_Retransmision(L_i).E_Borrado Then
          PL_Indice_Libre_Retransmision = L_i
          GV_Estructura_Retransmision(L_i).E_Borrado = False
         Exit Function
        End If
    Next

    ' Si ninguno de los elementos del arreglo han sido
    ' borrados entonces se crea una nueva casilla en el
    ' arreglo, redimensionandolo
    ' y retorna el nuevo indice.
    ReDim Preserve GV_Estructura_Retransmision( _
    L_ArrayCount + 1)
    PL_Indice_Libre_Retransmision = L_ArrayCount + 1
End Function

