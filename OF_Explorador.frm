VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.1#0"; "comctl32.ocx"
Begin VB.Form OF_Explorador 
   Caption         =   "@ERG2 -Explorador"
   ClientHeight    =   4545
   ClientLeft      =   1410
   ClientTop       =   1425
   ClientWidth     =   3690
   HelpContextID   =   22
   Icon            =   "OF_Explorador.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4545
   ScaleWidth      =   3690
   Begin ComctlLib.TreeView OA_Explorador 
      Height          =   3984
      HelpContextID   =   22
      Left            =   72
      TabIndex        =   0
      Top             =   384
      Width           =   3132
      _ExtentX        =   5530
      _ExtentY        =   7011
      _Version        =   327680
      Indentation     =   176
      LabelEdit       =   1
      Style           =   7
      ImageList       =   "ImageList1"
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   3120
      Top             =   1860
      _ExtentX        =   794
      _ExtentY        =   794
      BackColor       =   -2147483643
      ImageWidth      =   15
      ImageHeight     =   15
      _Version        =   327680
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   10
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "OF_Explorador.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "OF_Explorador.frx":075C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "OF_Explorador.frx":0A76
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "OF_Explorador.frx":0D90
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "OF_Explorador.frx":10AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "OF_Explorador.frx":13C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "OF_Explorador.frx":16DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "OF_Explorador.frx":19F8
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "OF_Explorador.frx":1D12
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "OF_Explorador.frx":202C
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "OF_Explorador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub PL_Cargar_Canales(LP_Cual%)
' /***********************************************************/
' Función que carga todos las ventanas de canales al Treeview
' /**********************************************************/
On Error GoTo Etiqueta_Error:

Dim L_nodX As Node
Dim L_i As Integer
Dim L_ArrayCount As Integer

    L_ArrayCount = UBound(GV_VENTANAS_Canal)

    
    For L_i = 1 To L_ArrayCount
        If Not GV_Estado_Canal(L_i).Deleted Then
          If GV_VENTANAS_Canal(L_i).OT_Ventana_Estatus = _
            LP_Cual Then
            Set nodX = OA_Explorador.Nodes.Add( _
            "S" + CStr(LP_Cual), tvwChild, "C" + _
            CStr(L_i), GV_VENTANAS_Canal(L_i).OT_Canal, 2)
            nodX.Tag = CStr(L_i)
            nodX.EnsureVisible
          End If
        End If
    Next L_i

    
    'nodX.EnsureVisible  ' Show all nodes.

Exit Sub
Etiqueta_Error:

End Sub

Sub PL_Cargar_Listas_Canales(LP_Cual%)
' /***********************************************************/
' Función que carga todos las ventanas de Listas de canales al
' Treeview
' /**********************************************************/

On Error GoTo Etiqueta_Error:
Dim L_nodX As Node
Dim L_i As Integer
Dim L_ArrayCount As Integer

    L_ArrayCount = UBound(GV_VENTANAS_Lista_Canales)

    
    For L_i = 1 To L_ArrayCount
     If Not GV_Estado_Lista_Canales(L_i).Deleted Then
      If GV_VENTANAS_Lista_Canales(L_i).OT_Ventana_Estatus _
       = LP_Cual Then
        
       Set nodX = OA_Explorador.Nodes.Add( _
       "S" + CStr(LP_Cual), tvwChild, "Z" + _
       CStr(L_i), GV_VENTANAS_Lista_Canales(L_i).Caption, 6)
       nodX.Tag = CStr(L_i)
       nodX.EnsureVisible
     End If
    End If
   Next L_i

Exit Sub
Etiqueta_Error:
End Sub

Sub PL_Cargar_Privados(LP_Cual%)
' /***********************************************************/
' Función que carga todos las ventanas de usuarios al
' Treeview
' /**********************************************************/
On Error GoTo Etiqueta_Error:
Dim L_nodX As Node
Dim L_i As Integer
Dim L_ArrayCount As Integer

    L_ArrayCount = UBound(GV_VENTANAS_Usuario)

    
    For L_i = 1 To L_ArrayCount
        If Not GV_Estado_Usuario(L_i).Deleted Then
          If GV_VENTANAS_Usuario(L_i).OT_Ventana_Estatus _
          = LP_Cual Then
          
          Set nodX = OA_Explorador.Nodes.Add( _
          "S" + CStr(LP_Cual), tvwChild, "U" + _
          CStr(L_i), GV_VENTANAS_Usuario(L_i).OT_Nick, 4)
          nodX.Tag = CStr(L_i)
          nodX.EnsureVisible
         End If
      End If
    Next L_i

Etiqueta_Error:
End Sub


Sub PL_Cargar_Treeview()
' /***********************************************************/
' Función que carga llama a las funciones de carga
' de las diferentes tipos de ventanas que maneja la
' aplicación
' /**********************************************************/
On Error GoTo Etiqueta_Error:

Dim L_nodX As Node
Dim L_i As Integer
Dim L_ArrayCount As Integer

    L_ArrayCount = UBound(GV_VENTANAS_Estatus)

    Set nodX = OA_Explorador.Nodes.Add(, , "EX", "ERG2", 8)
    nodX.Tag = "EXPLORADOR"
    nodX.EnsureVisible
    For L_i = 1 To L_ArrayCount
        If Not GV_Estado_Estatus(L_i).Deleted Then
          If GV_VENTANAS_Estatus(L_i).OT_SocketAsociado _
           <> 0 Then
            Set nodX = OA_Explorador.Nodes.Add( _
            "EX", tvwChild, "S" + CStr(L_i), _
            GV_VENTANAS_Estatus(L_i).Caption, 3)
          Else
            Set nodX = OA_Explorador.Nodes.Add( _
            "EX", tvwChild, "S" + CStr(L_i), _
            GV_VENTANAS_Estatus(L_i).Caption, 5)
          End If
          nodX.Tag = CStr(L_i)
           ' Procesar
          nodX.EnsureVisible
          PL_Cargar_Canales L_i
          PL_Cargar_Privados L_i
          PL_Cargar_Listas_Canales L_i
        End If
    Next L_i

Exit Sub
Etiqueta_Error:

End Sub


Private Sub Form_Activate()
OF_principal.WindowState = 0

End Sub

Private Sub Form_Load()
' /***********************************************************/
' Procedimiento que carga al explorador y mueve el MDI al
' lado de él
' /**********************************************************/

GV_Explorador = True
Me.Move 0, 0, 2500, Screen.Height - 300
OF_principal.WindowState = 0
OF_principal.Move 0 + Me.Width, 0, _
Screen.Width - Me.Width, Screen.Height - 300
MV_Minimizar_Ventanas "TODAS", 0, 3

PL_Cargar_Treeview



End Sub

Private Sub Form_Resize()

On Error GoTo Etiqueta_Error:
 OA_Explorador.Move 0, 0, ScaleWidth, ScaleHeight - 400

Exit Sub
Etiqueta_Error:

End Sub


Private Sub Form_Unload(Cancel As Integer)
' /***********************************************************/
' Procedimiento que descarga al explorador y restaura el MDI al
' estado de maximizado
' /**********************************************************/


MV_Esconde_Servidores 0, False
MV_Esconde_Canales 0, False
MV_Esconde_Usuarios 0, False
MV_Esconde_Lista_Canales 0, False
OF_principal.WindowState = 2

GV_Explorador = False

End Sub


Private Sub OA_Explorador_Collapse(ByVal Node As ComctlLib.Node)
' /***********************************************************/
' Procedimiento que esconde las ventanas asociadas a un
' servidor cuando se hace click en el botón de "+"
' en el explorador
' /**********************************************************/

Dim L_indice%
If Node.Key = "EX" Then Exit Sub
L_indice = CInt(Mid(Node.Key, 2, Len(Node.Key) - 1))
MV_Esconde_Servidores L_indice, True
MV_Esconde_Canales L_indice, True
MV_Esconde_Usuarios L_indice, True
MV_Esconde_Lista_Canales L_indice, True

End Sub

Private Sub OA_Explorador_Expand(ByVal Node As ComctlLib.Node)
' /***********************************************************/
' Procedimiento que muestra las ventanas asociadas a un
' servidor cuando se hace click en el botón de "-"
' en el explorador
' /**********************************************************/

Dim L_indice%
If Node.Key = "EX" Then Exit Sub
L_indice = CInt(Mid(Node.Key, 2, Len(Node.Key) - 1))
MV_Esconde_Servidores L_indice, False
MV_Esconde_Canales L_indice, False
MV_Esconde_Usuarios L_indice, False
MV_Esconde_Lista_Canales L_indice, False

End Sub

Private Sub OA_Explorador_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then
   DoEvents
   PopupMenu OF_principal.M_Ventanas
End If
End Sub


Private Sub OA_Explorador_NodeClick(ByVal Node As ComctlLib.Node)
' /***********************************************************/
' Procedimiento que se posiciona y muestra la ventana asociada
' al nodo al que se hizo CLICK en el explorador
' /**********************************************************/


Dim L_cual$
Dim L_indice%
Dim L_estado%

If Node.Key = "EX" Then Exit Sub

L_cual = Left(Node.Key, 1)
L_indice = CInt(Mid(Node.Key, 2, Len(Node.Key) - 1))
Select Case L_cual

  Case "S"
    
    If GV_VENTANAS_Estatus(L_indice).Visible Then
      GV_VENTANAS_Estatus(L_indice).OT_Comando.SetFocus
      
    End If
    
  Case "C"
    If GV_VENTANAS_Canal(L_indice).Visible Then
      GV_VENTANAS_Canal(L_indice).OT_Comando.SetFocus
      
    End If
  Case "U"
    If GV_VENTANAS_Usuario(L_indice).Visible Then
       GV_VENTANAS_Usuario(L_indice).OT_Comando.SetFocus
       
    End If
  Case "Z"
    If GV_VENTANAS_Lista_Canales(L_indice).Visible Then
      GV_VENTANAS_Lista_Canales(L_indice).OT_Comando.SetFocus
    End If
End Select
Me.SetFocus

End Sub


