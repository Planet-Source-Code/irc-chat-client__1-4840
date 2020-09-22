Attribute VB_Name = "MO_Ventanas"
Option Explicit

' Constante que nos indica la cantidad de bytes que tomara el
' texto donde se reflejan los mensajes
Global Const GC_Maximo_Texto = 20000

' Variable booleana que nos indica si el explorador esta activo
Global GV_Explorador As Boolean

' Arreglo de Estados de las Ventanas
    
    
Type ES_Estados
  Deleted As Integer ' Si la ventana tiene estado de borrado
End Type
    

' Definición de arreglos que nos
' permitan rastrear las ventanas activas en el sistema

    Global GV_Estado_Estatus()  As ES_Estados
    Global GV_Estado_Canal()  As ES_Estados
    Global GV_Estado_Usuario()  As ES_Estados
    Global GV_Estado_Lista_Canales() As ES_Estados
        

' Declaración de arreglos de las
' Diferentes ventanas del sisema
' Ventanas de servidores
 Global GV_VENTANAS_Estatus() As New OF_Estatus
 
 ' Ventanas de Canales
 Global GV_VENTANAS_Usuario() As New OF_Hablar_Usuario
 
 ' Ventanas de Lista de Canales
 Global GV_VENTANAS_Lista_Canales() As New OF_Lista_Canales
 
 ' Ventanas de Platicas con Usuarios
 Global GV_VENTANAS_Canal() As New OF_Hablar_Canal
 
Function MV_Busca_Ventana_Lista_Canales(Lp_Cualsocket%) As Integer
' /********************************************************/
' Procedimiento que se encarga de buscar la ventana de
' Lista de Canales
' respectiva a un socket asociado.
' /********************************************************/

On Error GoTo Etiqueta_Error:

Dim L_i%, L_Cuantos%

' Cantidad de Ventanas de Listas de Canales
L_Cuantos = UBound(GV_VENTANAS_Lista_Canales)

For L_i = 1 To L_Cuantos
  If Not GV_Estado_Lista_Canales(L_i).Deleted Then
   If GV_VENTANAS_Lista_Canales(L_i).OT_SocketAsociado = _
     Lp_Cualsocket Then
        MV_Busca_Ventana_Lista_Canales = L_i
        Exit Function
    End If
   End If
Next L_i

MV_Busca_Ventana_Lista_Canales = 0

Exit Function
Etiqueta_Error:
ME_Muestra_Error
End Function

Function MV_BuscaSigAnt_Ventana_Canal(LP_Cual, _
Lp_Direcc$, LP_Activos As Boolean) As Integer
' /*****************************************************/
' Busca ya sea la ventana anterior o siguiente entre las
' dependiendo de la dirección especificada.  Además
' indicando si se desea que la
' Búsqueda sea sobre ventanas activas.
' /*****************************************************/

On Error GoTo Etiqueta_Error:

Dim L_i%

If Lp_Direcc = "ANTERIOR" Then
 If LP_Cual = 1 Then _
  L_i = UBound(GV_VENTANAS_Canal) Else L_i = LP_Cual - 1
 While True
      If L_i = 0 Then L_i = UBound(GV_VENTANAS_Canal)
      If (Not GV_Estado_Canal(L_i).Deleted) Then
         If GV_VENTANAS_Canal(L_i).Visible = True Then
            ' Puede dar un circulo completo
            MV_BuscaSigAnt_Ventana_Canal = L_i
            Exit Function
          End If
      End If
      If L_i = 0 Then
        L_i = UBound(GV_VENTANAS_Canal)
      Else
        L_i = L_i - 1
      End If
      
      DoEvents
    Wend

Else
   
    If LP_Cual = UBound(GV_VENTANAS_Canal) Then _
      L_i = 1 Else L_i = LP_Cual + 1
    While True
      If L_i = UBound(GV_VENTANAS_Canal) + 1 Then L_i = 1
      If (Not GV_Estado_Canal(L_i).Deleted) Then
          If GV_VENTANAS_Canal(L_i).Visible = True Then
            ' Puede dar un circulo completo
            MV_BuscaSigAnt_Ventana_Canal = L_i
            Exit Function
          End If
      End If
      If L_i = UBound(GV_VENTANAS_Canal) Then
        L_i = 1
      Else
        L_i = L_i + 1
      End If
      
      DoEvents
    Wend

End If

Exit Function
Etiqueta_Error:
ME_Muestra_Error

End Function

Function MV_BuscaSigAnt_Ventana_Lista_Canales(LP_Cual, _
Lp_Direcc$, LP_Activos As Boolean) As Integer
' /******************************************************/
' Busca ya sea la ventana anterior o siguiente entre las
' ventanas de lista de canales,
' dependiendo de la dirección especificada.  Además
' indicando si se desea que la
' Búsqueda sea sobre ventanas activas.
' /******************************************************/

On Error GoTo Etiqueta_Error:

Dim L_i%

If Lp_Direcc = "ANTERIOR" Then
    If LP_Cual = 1 Then _
    L_i = UBound(GV_VENTANAS_Lista_Canales) _
    Else L_i = LP_Cual - 1
    While True
       If L_i = 0 Then L_i = UBound(GV_VENTANAS_Lista_Canales)
      If (Not GV_Estado_Lista_Canales(L_i).Deleted) Then
        If GV_VENTANAS_Lista_Canales(L_i).Visible = True Then
         ' Puede dar un circulo completo
         MV_BuscaSigAnt_Ventana_Lista_Canales = L_i
         Exit Function
        End If
     End If
      If L_i = 0 Then
        L_i = UBound(GV_VENTANAS_Lista_Canales)
      Else
        L_i = L_i - 1
      End If
      DoEvents
    Wend

Else
   
    If LP_Cual = UBound(GV_VENTANAS_Lista_Canales) Then _
     L_i = 1 Else L_i = LP_Cual + 1
    While True
      If L_i = UBound(GV_VENTANAS_Lista_Canales) + 1 Then _
       L_i = 1
      If (Not GV_Estado_Lista_Canales(L_i).Deleted) Then
        If GV_VENTANAS_Lista_Canales(L_i).Visible = True Then
            ' Puede dar un circulo completo
            MV_BuscaSigAnt_Ventana_Lista_Canales = L_i
            Exit Function
         End If
      End If
      If L_i = UBound(GV_VENTANAS_Lista_Canales) Then
        L_i = 1
      Else
        L_i = L_i + 1
      End If
      DoEvents
      
    Wend

End If

Exit Function
Etiqueta_Error:
ME_Muestra_Error

End Function

Function MV_BuscaSigAnt_Ventana_Estatus(LP_Cual, _
Lp_Direcc$, LP_Activos As Boolean) As Integer
' /*******************************************************/
' Busca ya sea la ventana anterior o siguiente entre las
' ventanas de estatus, dependiendo de la dirección
' especificada.  Además indicando si se desea que la
' Búsqueda sea sobre ventanas activas.
' /*******************************************************/

On Error GoTo Etiqueta_Error:

Dim L_i%

If Lp_Direcc = "ANTERIOR" Then
    If LP_Cual = 1 Then _
     L_i = UBound(GV_VENTANAS_Estatus) Else L_i = LP_Cual - 1
    While True
       If L_i = 0 Then L_i = UBound(GV_VENTANAS_Estatus)
      If (Not GV_Estado_Estatus(L_i).Deleted) Then
          If GV_VENTANAS_Estatus(L_i).Visible = True Then
            ' Puede dar un circulo completo
            MV_BuscaSigAnt_Ventana_Estatus = L_i
            Exit Function
          End If
      End If
      If L_i = 0 Then
        L_i = UBound(GV_VENTANAS_Estatus)
      Else
        L_i = L_i - 1
      End If
      
      DoEvents
    Wend

Else
   
    If LP_Cual = UBound(GV_VENTANAS_Estatus) + 1 _
      Then L_i = 1 Else L_i = LP_Cual + 1
    While True
      If L_i = UBound(GV_VENTANAS_Estatus) + 1 Then _
        L_i = 1
      If (Not GV_Estado_Estatus(L_i).Deleted) Then
         If GV_VENTANAS_Estatus(L_i).Visible = True Then
            ' Puede dar un circulo completo
            MV_BuscaSigAnt_Ventana_Estatus = L_i
            Exit Function
         End If
      End If
      If L_i = UBound(GV_VENTANAS_Estatus) Then
         L_i = 1
      Else
        L_i = L_i + 1
      End If
      DoEvents
    Wend

End If

Exit Function
Etiqueta_Error:
ME_Muestra_Error

End Function

Function MV_BuscaSigAnt_Ventana_Usuario(LP_Cual, _
Lp_Direcc$, LP_Activos As Boolean) As Integer

' /******************************************************/
' Busca ya sea la ventana anterior o siguiente entre las
' ventanas de usuarios, dependiendo de la dirección
' especificada.  Además indicando si se desea que la
' Búsqueda sea sobre ventanas activas.
' /******************************************************/

On Error GoTo Etiqueta_Error:

Dim L_i%

If Lp_Direcc = "ANTERIOR" Then
    If LP_Cual = 1 Then L_i = _
     UBound(GV_VENTANAS_Usuario) Else L_i = LP_Cual - 1
    While True
      If L_i = 0 Then L_i = UBound(GV_VENTANAS_Usuario)
      If (Not GV_Estado_Usuario(L_i).Deleted) Then
         If GV_VENTANAS_Usuario(L_i).Visible = True Then
            ' Puede dar un circulo completo
            MV_BuscaSigAnt_Ventana_Usuario = L_i
            Exit Function
          End If
      End If
      If L_i = 0 Then
        L_i = UBound(GV_VENTANAS_Usuario)
      Else
        L_i = L_i - 1
      End If

      DoEvents
    Wend

Else
   
    If LP_Cual = UBound(GV_VENTANAS_Usuario) _
     Then L_i = 1 Else L_i = LP_Cual + 1
    While True
      If L_i = UBound(GV_VENTANAS_Usuario) + 1 Then L_i = 1
      If (Not GV_Estado_Usuario(L_i).Deleted) Then
        If GV_VENTANAS_Usuario(L_i).Visible = True Then
            ' Puede dar un circulo completo
            MV_BuscaSigAnt_Ventana_Usuario = L_i
            Exit Function
        End If
      End If
      If L_i = UBound(GV_VENTANAS_Usuario) Then
        L_i = 1
      Else
        L_i = L_i + 1
      End If
      DoEvents
    Wend

End If

Exit Function
Etiqueta_Error:
ME_Muestra_Error

End Function

Sub MV_Cerrar_El_Resto_De_Ventanas(LP_Cual)
Dim L_i%, L_Cuantos%

' /****************************************************/
' Cierra todas las ventanas sin importar el tipo, que
' estén asociadas con el socket
' LP_Cual.
' /****************************************************/

' Cantidad de Ventanas de Canal
L_Cuantos = UBound(GV_VENTANAS_Canal)

'Cierra ventanas de canal
For L_i = 1 To L_Cuantos
    If Not GV_Estado_Canal(L_i).Deleted Then
        If GV_VENTANAS_Canal(L_i).OT_Ventana_Estatus = _
         LP_Cual Then
            Unload GV_VENTANAS_Canal(L_i)
                       
        End If
    End If
    DoEvents
Next L_i

' Cantidad de Ventanas de Canal
L_Cuantos = UBound(GV_VENTANAS_Usuario)

'Cierra ventanas de usuario
For L_i = 1 To L_Cuantos
    If Not GV_Estado_Usuario(L_i).Deleted Then
        If GV_VENTANAS_Usuario(L_i).OT_Ventana_Estatus = _
         LP_Cual Then
            Unload GV_VENTANAS_Usuario(L_i)
                       
        End If
    End If
    DoEvents
Next L_i

' Cantidad de Ventanas de Canal
L_Cuantos = UBound(GV_VENTANAS_Lista_Canales)

'Cierra ventanas de lista de canales
For L_i = 1 To L_Cuantos
    If Not GV_Estado_Lista_Canales(L_i).Deleted Then
        If GV_VENTANAS_Lista_Canales(L_i).OT_Ventana_Estatus = _
        LP_Cual Then
            Unload GV_VENTANAS_Lista_Canales(L_i)
                       
        End If
    End If
    DoEvents
Next L_i
End Sub

Function MV_CreaVentana_Lista_Canales(LP_Socket%) As Integer
' /**************************************************/
' LP_socket : el socket asociado a la ventana a crear
' Una vez creada la ventana se retorna el indice del
' arreglo donde
' Fue creada la ventana.
' Descripción :
' -------------
' Procedimiento que se encarga de una ventana de Lista
' de canales con el respectivo socket asociado
' /****************************************************/

On Error GoTo Etiqueta_Error:

Dim L_libre%
Dim L_nodo As Node

' Busca en el arreglo de ventanas una casilla libre
L_libre = MV_Indice_Libre_Lista_Canales

GV_VENTANAS_Lista_Canales(L_libre).Tag = L_libre
GV_VENTANAS_Lista_Canales(L_libre).OT_SocketAsociado = LP_Socket
GV_VENTANAS_Lista_Canales(L_libre).OT_Ventana_Estatus = _
GV_Sockets(LP_Socket).Ventana

If GV_Explorador Then
   Set L_nodo = OF_Explorador.OA_Explorador _
    .Nodes.Add( _
    "S" + CStr( _
    GV_VENTANAS_Lista_Canales(L_libre).OT_Ventana_Estatus) _
    , tvwChild, "Z" + CStr(L_libre), _
    GV_VENTANAS_Lista_Canales(L_libre).Caption, 6)
    
    L_nodo.EnsureVisible
    
End If

MV_CreaVentana_Lista_Canales = L_libre

Exit Function

Etiqueta_Error:
ME_Muestra_Error

End Function

Sub MV_Notifica_Salida(Lp_Nick$, Lp_Cualsocket)
' /*******************************************************/
' Procedimiento que se encarga de buscar en todas las
' ventanas de platicas privadas asociadas
' al socket LP_Cualsocket el alias LP_Nick y notifica
' que este usuario ha sido sacado del servidor
' /*******************************************************/

On Error GoTo Etiqueta_Error:

Dim L_i%, L_Cuantos%
Dim L_vent%

L_Cuantos = UBound(GV_VENTANAS_Usuario)

For L_i = 1 To L_Cuantos
    If Not GV_Estado_Usuario(L_i).Deleted Then
        If GV_VENTANAS_Usuario(L_i).OT_SocketAsociado = _
          Lp_Cualsocket Then
           If GV_VENTANAS_Usuario(L_i).OT_Nick = Lp_Nick Then
             MV_Pone_Mensaje True, GV_VENTANAS_Usuario(L_i), _
            "Usuario < " + Lp_Nick + " > ha sido sacado " + _
               " del servidor", GV_Rojo
            Exit For
           End If
            
        End If
    End If
If GV_DoEVENTS Then DoEvents
Next L_i

Exit Sub
Etiqueta_Error:
ME_Muestra_Error



End Sub
Sub MV_Descarga_Nick_de_Canales(Lp_Nick$, Lp_Cualsocket)
' /*******************************************************/
' Procedimiento que se encarga de buscar en todas las
' ventanas de canal asociadas
' al socket LP_Cualsocket el alias LP_Nick y lo remueve
' de la lista del canal.
' /*******************************************************/

On Error GoTo Etiqueta_Error:

Dim L_i%, L_Cuantos%

L_Cuantos = UBound(GV_VENTANAS_Canal)

For L_i = 1 To L_Cuantos
    If Not GV_Estado_Canal(L_i).Deleted Then
        If GV_VENTANAS_Canal(L_i).OT_SocketAsociado = _
          Lp_Cualsocket Then
           MM_Descargar_Usuarios_De_Canal Lp_Nick, _
           GV_VENTANAS_Canal(L_i)
            
        End If
    End If
If GV_DoEVENTS Then DoEvents
Next L_i

Exit Sub
Etiqueta_Error:
ME_Muestra_Error


End Sub

Function MV_Indice_Libre_Lista_Canales()
' /****************************************************/
' Función que se encarga de buscar una casilla libre en
' el arreglo de Ventanas de Lista de Canales.
' /****************************************************/

   Dim L_i As Integer
   Dim L_ArrayCount As Integer

    L_ArrayCount = UBound(GV_VENTANAS_Lista_Canales)

    ' Recorrer el arreglo de ventanas. Si la ventana ha
    ' sido borrada
    ' entonces, retorna ese indice.
    For L_i = 1 To L_ArrayCount
        If GV_Estado_Lista_Canales(L_i).Deleted Then
            MV_Indice_Libre_Lista_Canales = L_i
            GV_Estado_Lista_Canales(L_i).Deleted = False
            Exit Function
        End If
    Next

    ' Si ninguno de los elementos del arreglo han sido borrados
    ' entonces se crea una nueva casilla en el arreglo,
    ' redimensionandolo y retorna el nuevo indice.
    ReDim Preserve GV_VENTANAS_Lista_Canales(L_ArrayCount + 1)
    ReDim Preserve GV_Estado_Lista_Canales(L_ArrayCount + 1)
    MV_Indice_Libre_Lista_Canales = L_ArrayCount + 1
End Function

Sub MV_Esconde_Lista_Canales(LP_Cual%, LP_Esconder As Boolean)
' /**********************************************************/
' Procedimiento que esconde o muestra al usuario las ventanas
' de lista de canales asociadas con el valor de LP_Cual.
'Si Lp_Esconder es True entonces se esconden las ventanas,
' de lo contrario estas son mostradas en la aplicación.
' /**********************************************************/

On Error GoTo Etiqueta_Error:

Dim L_i%, L_Cuantos%

' Cantidad de Ventanas de Canal
L_Cuantos = UBound(GV_VENTANAS_Lista_Canales)

For L_i = 1 To L_Cuantos
    If Not GV_Estado_Lista_Canales(L_i).Deleted Then
        If LP_Cual <> 0 Then
            If GV_VENTANAS_Lista_Canales( _
             L_i).OT_Ventana_Estatus = LP_Cual Then
              If LP_Esconder Then
               GV_VENTANAS_Lista_Canales(L_i).Hide
              Else
               GV_VENTANAS_Lista_Canales(L_i).Show
              End If
            End If
        Else
             If LP_Esconder Then
               GV_VENTANAS_Lista_Canales(L_i).Hide
             Else
               GV_VENTANAS_Lista_Canales(L_i).Show
             End If
        End If
    End If
    DoEvents
Next L_i

Exit Sub
Etiqueta_Error:
ME_Muestra_Error
End Sub

Sub MV_Esconde_Canales(LP_Cual%, LP_Esconder As Boolean)
' /******************************************************/
' Procedimiento que esconde o muestra al usuario las
' ventanas de canales asociadas con el valor de LP_Cual.
'
' Si Lp_Esconder es True entonces se esconden las
' ventanas, de lo contrario estas
' son mostradas en la aplicación.
' /******************************************************/

On Error GoTo Etiqueta_Error:

Dim L_i%, L_Cuantos%

' Cantidad de Ventanas de Canal
L_Cuantos = UBound(GV_VENTANAS_Canal)

For L_i = 1 To L_Cuantos
    If Not GV_Estado_Canal(L_i).Deleted Then
        If LP_Cual <> 0 Then
            If GV_VENTANAS_Canal(L_i).OT_Ventana_Estatus = _
             LP_Cual Then
              If LP_Esconder Then
               GV_VENTANAS_Canal(L_i).Hide
              Else
               GV_VENTANAS_Canal(L_i).Show
              End If
            End If
        Else
             If LP_Esconder Then
               GV_VENTANAS_Canal(L_i).Hide
             Else
               GV_VENTANAS_Canal(L_i).Show
             End If
        End If
    End If
    DoEvents
Next L_i

Exit Sub
Etiqueta_Error:
ME_Muestra_Error
End Sub

Sub MV_Min_Lista_Canales(LP_Cual%, LP_Estado%)
' /**************************************************/
' Procedimiento que aplica Lp_Estado a todas las
' ventanas de lista de canales cuya ventana de
' estatus es igual a LP_Cual (si este es diferente
' de 0), o a todas las  ventanas de canales si
' LP_Cual es 0.
' /**************************************************/

On Error GoTo Etiqueta_Error:

Dim L_i%, L_Cuantos%

' Cantidad de Ventanas de Canal
L_Cuantos = UBound(GV_VENTANAS_Lista_Canales)

For L_i = 1 To L_Cuantos
    If Not GV_Estado_Lista_Canales(L_i).Deleted Then
        If LP_Cual <> 0 Then
            If GV_VENTANAS_Lista_Canales( _
              L_i).OT_Ventana_Estatus = LP_Cual Then
               GV_VENTANAS_Lista_Canales(L_i).WindowState = _
                LP_Estado
            End If
        Else
            If LP_Estado = 3 Then
               GV_VENTANAS_Lista_Canales(L_i).WindowState = 0
               GV_VENTANAS_Lista_Canales(L_i).Move 0, 0, _
               OF_principal.ScaleWidth, _
               OF_principal.ScaleHeight
            Else
               GV_VENTANAS_Lista_Canales(L_i).WindowState = _
               LP_Estado
            End If
        End If
    End If
    DoEvents
Next L_i

Exit Sub
Etiqueta_Error:
ME_Muestra_Error


End Sub

Sub MV_Min_Canales(LP_Cual%, LP_Estado%)
' /***************************************************/
' Procedimiento que aplica Lp_Estado a todas las
' ventanas de canales cuya  ventana de estatus es
' igual a LP_Cual (si este es diferente de 0), o a todas
' las ventanas de canales si LP_Cual es 0.
' /***************************************************/

On Error GoTo Etiqueta_Error:

Dim L_i%, L_Cuantos%

' Cantidad de Ventanas de Canal
L_Cuantos = UBound(GV_VENTANAS_Canal)

For L_i = 1 To L_Cuantos
    If Not GV_Estado_Canal(L_i).Deleted Then
        If LP_Cual <> 0 Then
            If GV_VENTANAS_Canal(L_i).OT_Ventana_Estatus = _
             LP_Cual Then
               GV_VENTANAS_Canal(L_i).WindowState = LP_Estado
            End If
        Else
            If LP_Estado = 3 Then
               GV_VENTANAS_Canal(L_i).WindowState = 0
               GV_VENTANAS_Canal(L_i).Move 0, 0, _
               OF_principal.ScaleWidth, _
               OF_principal.ScaleHeight
            Else
               GV_VENTANAS_Canal(L_i).WindowState = LP_Estado
            End If
        End If
    End If
    DoEvents
Next L_i

Exit Sub
Etiqueta_Error:
ME_Muestra_Error

End Sub

Sub MV_Esconde_Servidores(LP_Cual%, LP_Esconder As Boolean)
' /*******************************************************/
' Procedimiento que esconde o muestra al usuario las
' ventanas de estatus asociadas con el valor de LP_Cual.
' Si Lp_Esconder es True entonces se esconden las ventanas,
' de lo contrario estas son mostradas en la aplicación.
' /*******************************************************/

On Error GoTo Etiqueta_Error:
Dim L_i%, L_Cuantos%

' Cantidad de Ventanas de Estatus
L_Cuantos = UBound(GV_VENTANAS_Estatus)

For L_i = 1 To L_Cuantos
    If Not GV_Estado_Estatus(L_i).Deleted Then
        If LP_Cual <> 0 Then
            If GV_VENTANAS_Estatus(L_i).OT_Ventana_Estatus = _
             LP_Cual Then
               If LP_Esconder Then
                GV_VENTANAS_Estatus(L_i).Hide
               Else
                 GV_VENTANAS_Estatus(L_i).Show
               End If
               
            End If
        Else
               If LP_Esconder Then
                GV_VENTANAS_Estatus(L_i).Hide
               Else
                 GV_VENTANAS_Estatus(L_i).Show
               End If

         End If
    End If
    DoEvents
Next L_i

Exit Sub
Etiqueta_Error:
ME_Muestra_Error

End Sub

Sub MV_Min_Servidores(LP_Cual%, LP_Estado)
' /*****************************************************/
' Procedimiento que aplica Lp_Estado a todas las ventanas
' de estatus cuya ventana de estatus es igual a LP_Cual
' (si este es diferente de 0), o a todas las
' ventanas de canales si LP_Cual es 0.
' /*****************************************************/

On Error GoTo Etiqueta_Error:

Dim L_i%, L_Cuantos%

' Cantidad de Ventanas de Estatus
L_Cuantos = UBound(GV_VENTANAS_Estatus)

For L_i = 1 To L_Cuantos
    If Not GV_Estado_Estatus(L_i).Deleted Then
        If LP_Cual <> 0 Then
            If GV_VENTANAS_Estatus(L_i).OT_Ventana_Estatus = _
             LP_Cual Then
               GV_VENTANAS_Estatus(L_i).WindowState = LP_Estado
            End If
        Else
            If LP_Estado = 3 Then
               GV_VENTANAS_Estatus(L_i).WindowState = 0
               GV_VENTANAS_Estatus(L_i).Move 0, 0, _
               OF_principal.ScaleWidth, _
               OF_principal.ScaleHeight
            Else
               GV_VENTANAS_Estatus(L_i).WindowState = LP_Estado
            End If
         End If
    End If
    DoEvents
Next L_i

Exit Sub
Etiqueta_Error:
ME_Muestra_Error

End Sub

Sub MV_Esconde_Usuarios(LP_Cual%, LP_Esconder As Boolean)
' /*******************************************************/
' Procedimiento que esconde o muestra al usuario las
' ventanas de usuarios asociadas con el valor de LP_Cual.
'
' Si Lp_Esconder es True entonces se esconden las ventanas,
' de lo contrario estas son mostradas en la aplicación.
' /*******************************************************/

On Error GoTo Etiqueta_Error:

Dim L_i%, L_Cuantos%

' Cantidad de Ventanas de Usuario
L_Cuantos = UBound(GV_VENTANAS_Usuario)

For L_i = 1 To L_Cuantos
    If Not GV_Estado_Usuario(L_i).Deleted Then
        If LP_Cual <> 0 Then
            If GV_VENTANAS_Usuario(L_i).OT_Ventana_Estatus = _
            LP_Cual Then
               If LP_Esconder Then
                   GV_VENTANAS_Usuario(L_i).Hide
               Else
                   GV_VENTANAS_Usuario(L_i).Show
               End If
            End If
        Else
               If LP_Esconder Then
                   GV_VENTANAS_Usuario(L_i).Hide
               Else
                   GV_VENTANAS_Usuario(L_i).Show
               End If
        End If
    End If
    DoEvents
Next L_i

Exit Sub
Etiqueta_Error:
ME_Muestra_Error

End Sub

Sub MV_Min_Usuarios(LP_Cual%, LP_Estado%)
' /************************************************/
' Procedimiento que aplica Lp_Estado a todas las
' ventanas de usuarios cuya ventana de estatus es
' igual a LP_Cual (si este es diferente de 0), o a
' todas las ventanas de canales si LP_Cual es 0.
' /************************************************/

On Error GoTo Etiqueta_Error:

Dim L_i%, L_Cuantos%

' Cantidad de Ventanas de Usuario
L_Cuantos = UBound(GV_VENTANAS_Usuario)

For L_i = 1 To L_Cuantos
    If Not GV_Estado_Usuario(L_i).Deleted Then
        If LP_Cual <> 0 Then
            If GV_VENTANAS_Usuario(L_i).OT_Ventana_Estatus = _
            LP_Cual Then
              GV_VENTANAS_Usuario(L_i).WindowState = LP_Estado
            End If
        Else
           If LP_Estado = 3 Then
               GV_VENTANAS_Usuario(L_i).WindowState = 0
               GV_VENTANAS_Usuario(L_i).Move 0, 0, _
               OF_principal.ScaleWidth, _
               OF_principal.ScaleHeight
            Else
             GV_VENTANAS_Usuario(L_i).WindowState = LP_Estado
            End If
        End If
    End If
    DoEvents
Next L_i

Exit Sub
Etiqueta_Error:
ME_Muestra_Error


End Sub

Sub MV_Minimizar_Ventanas(LP_Cuales$, LP_Cual%, LP_Estado%)
' /*****************************************************/
' Aplica LP_estado a ciertas ventanas de acuerdo al
' criterio establecido en  LP_Cuales.
' LP_Cual: si se aplica LP_Estado a uno o varios
' servidores.
' /*****************************************************/

On Error GoTo Etiqueta_Error:

Select Case LP_Cuales

    Case "TODAS"
          MV_Min_Canales LP_Cual, LP_Estado
          MV_Min_Lista_Canales LP_Cual, LP_Estado
          MV_Min_Servidores LP_Cual, LP_Estado
          MV_Min_Usuarios LP_Cual, LP_Estado

    Case "CANALES"
          MV_Min_Canales LP_Cual, LP_Estado
    Case "SERVIDORES"
          MV_Min_Servidores LP_Cual, LP_Estado
    Case "USUARIOS"
          MV_Min_Usuarios LP_Cual, LP_Estado

End Select

Exit Sub
Etiqueta_Error:
ME_Muestra_Error
End Sub

Sub MV_Pone_Mensaje(LP_End As Boolean, LP_Forma As Form, LP_Mensaje$, LP_Color&)
' /*****************************************************/
' Coloca un mensaje en el textbox de mensajes de una
' forma.
' LP_Forma: Forma en la que se coloca el mensaje.
' LP_Color: Color del mensaje.
' /*****************************************************/

On Error GoTo Etiqueta_Error:

Dim L_temp$
Dim L_Largo%, L_largo2%
' Este procedimiento se encarga de mostrar un mensaje en una
' ventana
L_temp = LP_Forma.OL_Estatus.Text
L_largo2 = Len(L_temp)
L_Largo = Len(LP_Mensaje)

' Si el textbox esta demasiado lleno, esto se puede
'hacer una constante
If (L_largo2 + L_Largo) > GC_Maximo_Texto Then

  
  LP_Forma.OL_Estatus.SelStart = 0
  LP_Forma.OL_Estatus.SelLength = Len(LP_Mensaje)
  LP_Forma.OL_Estatus.SelText = ""

  
End If
L_largo2 = Len(L_temp)
LP_Forma.OL_Estatus.SelStart = L_largo2

LP_Forma.OL_Estatus.SelColor = LP_Color
LP_Forma.OL_Estatus.SelText = LP_Mensaje


Exit Sub
Etiqueta_Error:
ME_Muestra_Error

End Sub

Function MV_Busca_Ventana_Canal(Lp_CanalBuscado$, Lp_Cualsocket) As Integer
' /***************************************************/
' Procedimiento que se encarga de buscar la ventana
' de canal respectiva a un socket asociado utilizando
' el canal  buscado.  Si encuentra la ventana entonces
' retorna el indice donde se encontro la ventana de '
' canal en el arreglo de ventanas de canales. De lo
' contrario retorna 0.
'
' LP_canalbuscado : el canal con el que se desea buscar
' la ventana.
' LP_cualsocket : el inidce del arreglo del socket con
' el que se desea buscar la ventana.
' /****************************************************/

On Error GoTo Etiqueta_Error:

Dim L_i%, L_Cuantos%

' Cantidad de Ventanas de Canal
L_Cuantos = UBound(GV_VENTANAS_Canal)

For L_i = 1 To L_Cuantos
    If Not GV_Estado_Canal(L_i).Deleted Then
        If GV_VENTANAS_Canal(L_i).OT_SocketAsociado = _
         Lp_Cualsocket And _
         (UCase(GV_VENTANAS_Canal(L_i).OT_Canal)) = _
          UCase(Lp_CanalBuscado) Then
            MV_Busca_Ventana_Canal = L_i
            Exit Function
        End If
    End If

Next L_i

MV_Busca_Ventana_Canal = 0
Exit Function
Etiqueta_Error:
ME_Muestra_Error
End Function

Function MV_CreaVentana_Canal(LP_Socket%, LP_Canal$) As Integer
' /********************************************************/
' Procedimiento que se encarga de una ventana de Canal con
' el respectivo socket asociado, utilizando el nombre del
' canal de parametro.  Una vez creada la ventana
' se retorna el indice del arreglo donde fue creada la
' ventana.
' LP_socket : el socket (indice) asociado a la ventana a
' crear  LP_Canal: el Canal asociado a la ventana a crear
' /********************************************************/

On Error GoTo Etiqueta_Error:

Dim L_libre%
Dim L_nodo As Node


' Busca en el arreglo de ventanas una casilla libre
L_libre = MV_Indice_Libre_Canal

GV_VENTANAS_Canal(L_libre).Tag = L_libre
GV_VENTANAS_Canal(L_libre).OT_SocketAsociado = LP_Socket
GV_VENTANAS_Canal(L_libre).OT_Ventana_Estatus = _
GV_Sockets(LP_Socket).Ventana
GV_VENTANAS_Canal(L_libre).OT_Canal = LP_Canal
GV_VENTANAS_Canal(L_libre).Caption = LP_Canal

MV_CreaVentana_Canal = L_libre
If GV_Explorador Then
   Set L_nodo = OF_Explorador.OA_Explorador _
    .Nodes.Add( _
    "S" + CStr(GV_VENTANAS_Canal(L_libre).OT_Ventana_Estatus) _
    , tvwChild, "C" + CStr(L_libre), LP_Canal, 2)
    L_nodo.EnsureVisible
End If

Exit Function
Etiqueta_Error:
ME_Muestra_Error
End Function

Function MV_CreaVentana_Usuario(LP_Socket%, Lp_Nick$) As Integer
' /****************************************************/
' Procedimiento que se encarga de una ventana de usuario
' con el respectivo socket asociado, utilizando el alias
' de parametro una vez creada la ventana se retorna
' el indice del arreglo donde fue creada la ventana.
'
' LP_socket : el socket asociado a la ventana a crear.
' LP_Nick  : el alias asociado a la ventana a crear.
' /****************************************************/

On Error GoTo Etiqueta_Error:

Dim L_libre%
Dim L_nodo As Node

' Busca en el arreglo de ventanas una casilla libre
L_libre = MV_Indice_Libre_Usuario

GV_VENTANAS_Usuario(L_libre).Tag = L_libre
GV_VENTANAS_Usuario(L_libre).OT_SocketAsociado = LP_Socket
GV_VENTANAS_Usuario(L_libre).OT_Ventana_Estatus = _
GV_Sockets(LP_Socket).Ventana

GV_VENTANAS_Usuario(L_libre).OT_Nick = Lp_Nick

If GV_Explorador Then
   Set L_nodo = OF_Explorador.OA_Explorador _
    .Nodes.Add( _
    "S" + CStr(GV_VENTANAS_Usuario(L_libre).OT_Ventana_Estatus) _
    , tvwChild, "U" + CStr(L_libre), Lp_Nick, 4)
   L_nodo.EnsureVisible
End If


MV_CreaVentana_Usuario = L_libre

Exit Function
Etiqueta_Error:
ME_Muestra_Error
End Function

Sub MV_Despacha_mensaje(LP_Mensaje As String, LP_Ventana%, LP_Socket%)
' /*******************************************************/
' Procedimiento que secciona un mensaje proveniente de un
' servidor, usando como delimitador <CR-LF>.  Cada sección
' es luego enviada al parser de mensajes.
'
' LP_Mensaje : el mensaje a enviar sin formatear
' LP_ventana  : La ventana que originó el mensaje
' LP_socket : el indice del arrreglo de sockets donde
' será enviado el mensaje
' /*******************************************************/

On Error GoTo Etiqueta_Error:

Dim L_vent%, L_pos1%, L_pos2%
Dim L_Men$, L_Datos As Boolean
Dim L_prefijo$, L_Comando$, L_params$
Dim L_Status%

Dim L_Temp1%
L_Datos = True
L_pos1 = 1
L_pos2 = InStr(1, LP_Mensaje, Chr(10))
L_Temp1 = InStr(1, LP_Mensaje, Chr(13))
' Parche por Rogger Vasquez 19 de enero de 1998
If L_Temp1 = 0 Then
  LP_Mensaje = LP_Mensaje + Chr(13)
End If
DoEvents
While L_Datos
    
    L_Men = Mid(LP_Mensaje, L_pos1, L_pos2 - L_pos1 + 1)
    ' PArche por Rogger 19 de Enero de 1998
    L_Temp1 = InStr(1, L_Men, vbCrLf)
    If L_Temp1 = 0 Then L_Men = L_Men + Chr(13)
    
    L_Status = _
    MM_Parsear_Mensaje(L_Men, L_prefijo, L_Comando, L_params)

    ' Tomar las acciones necesarias de acuerdo al mensaje
    
    '@30@ MV_Pone_Mensaje True, GV_VENTANAS_Estatus(Lp_ventana), L_Men, GV_Rojo
    
    MM_Entrega_Mensaje LP_Socket, L_prefijo, L_Comando, _
     L_params, L_Status
     
    If L_pos2 = Len(LP_Mensaje) Then
       L_Datos = False
    Else
       L_pos1 = L_pos2 + 1
       L_pos2 = InStr(L_pos1, LP_Mensaje, Chr(10))
       
       If L_pos2 = 0 Then
            L_Datos = False
       End If
      
    End If
    If GV_DoEVENTS Then DoEvents

Wend

Exit Sub
Etiqueta_Error:
ME_Muestra_Error

End Sub

Function MV_Indice_Libre_Canal()
' /******************************************************/
' Función que se encarga de buscar una casilla libre en
' el arreglo de ventanas de canales.
' /******************************************************/

   Dim L_i As Integer
   Dim L_ArrayCount As Integer

    L_ArrayCount = UBound(GV_VENTANAS_Canal)

  ' Recorrer el arreglo de ventanas. Si la ventana ha sido
  ' borrada entonces, retorna ese indice.
    For L_i = 1 To L_ArrayCount
        If GV_Estado_Canal(L_i).Deleted Then
            MV_Indice_Libre_Canal = L_i
            GV_Estado_Canal(L_i).Deleted = False
            Exit Function
        End If
    Next

    ' Si ninguno de los elementos del arreglo han sido
    ' borrados  entonces se crea una nueva casilla en
    ' el arreglo, redimensionandolo
    ' y retorna el nuevo indice.
    ReDim Preserve GV_VENTANAS_Canal(L_ArrayCount + 1)
    ReDim Preserve GV_Estado_Canal(L_ArrayCount + 1)
    MV_Indice_Libre_Canal = UBound(GV_VENTANAS_Canal)
End Function

Sub MV_Reemplazar_Alias(LP_Alias_Anterior$, LP_Alias_Nuevo$, LP_Socket%)
' /*******************************************************/
' Busca un alias y lo remplaza por otro en todas las
' ventanas de canales asociadas  a LP_Socket.
'
' LP_Alias_Anterior: Alias a reemplazar.
' LP_Alias_Nuevo: Alias con que se reemplazará el
' alias anterior.
' LP_Socket: Indice que representa el servidor en
' el arreglo de sockets.
' /*******************************************************/

On Error GoTo Etiqueta_Error:

Dim L_i%, L_j%, L_Cuantos%, L_Simbolo$

L_Simbolo = ""
'Ventanas de Canales Existentes
L_Cuantos = UBound(GV_VENTANAS_Canal)

For L_i = 1 To L_Cuantos
 'DoEvents
 If Not GV_Estado_Canal(L_i).Deleted Then
  If GV_VENTANAS_Canal(L_i).OT_SocketAsociado = _
    LP_Socket Then
       For L_j = 0 To _
        (GV_VENTANAS_Canal(L_i).OL_Usuarios.ListCount - 1)
          If Left(Trim(GV_VENTANAS_Canal( _
           L_i).OL_Usuarios.List(L_j)), 1) = "@" _
            Then L_Simbolo = "@"
             If UCase(Trim(GV_VENTANAS_Canal( _
               L_i).OL_Usuarios.List(L_j))) = _
                Trim(L_Simbolo + _
                 UCase(Trim(LP_Alias_Anterior))) Then
                    GV_VENTANAS_Canal( _
                    L_i).OL_Usuarios.List(L_j) = _
                    Trim(L_Simbolo + Trim(LP_Alias_Nuevo))
                     Exit For
           End If
             L_Simbolo = ""
       Next L_j
     End If
  End If
Next L_i
' Cantidad de Ventanas de Usuario
L_Cuantos = UBound(GV_VENTANAS_Usuario)
For L_i = 1 To L_Cuantos
 If Not GV_Estado_Usuario(L_i).Deleted Then
   If GV_VENTANAS_Usuario(L_i).OT_SocketAsociado = _
    LP_Socket Then
     If UCase(Trim(GV_VENTANAS_Usuario(L_i).OT_Nick)) = _
      UCase(Trim(LP_Alias_Anterior)) Then
       GV_VENTANAS_Usuario(L_i).OT_Nick = Trim(LP_Alias_Nuevo)
       GV_VENTANAS_Usuario(L_i).Caption = Trim(LP_Alias_Nuevo)
     End If
   End If
 End If
Next L_i

Exit Sub
Etiqueta_Error:
ME_Muestra_Error
End Sub

Sub MV_Setear_Socket(LP_Cual%)
' /********************************************************/
' Se setea el socket de todas las ventanas asociadas a
' LP_Cual con el valor 0.
' Esto se hace cuando se ha desconectado la ventana de
' un sevidor.
' /********************************************************/

On Error GoTo Etiqueta_Error:

Dim L_i%, L_Cuantos%

'Setear ventanas de canal
' Cantidad de Ventanas de Canal
L_Cuantos = UBound(GV_VENTANAS_Canal)

For L_i = 1 To L_Cuantos
        If Not GV_Estado_Canal(L_i).Deleted Then
            If GV_VENTANAS_Canal(L_i).OT_SocketAsociado = _
             LP_Cual Then
              GV_VENTANAS_Canal(L_i).OT_SocketAsociado = 0
                GV_VENTANAS_Canal(L_i).Caption = _
                GV_VENTANAS_Canal(L_i).OT_Canal + _
                ":Connection Finished/La conexión ha " + _
                "sido terminada"
                MV_Pone_Mensaje True, _
                GV_VENTANAS_Canal(L_i), _
                "Connection Finished/La conexión " + _
                "ha sido terminada" + GV_EOD, vbRed
            End If
        End If
        DoEvents
Next L_i

'Setear ventanas de usuarios
' Cantidad de Ventanas de Usuario
L_Cuantos = UBound(GV_VENTANAS_Usuario)

For L_i = 1 To L_Cuantos
        If Not GV_Estado_Usuario(L_i).Deleted Then
            If GV_VENTANAS_Usuario(L_i).OT_SocketAsociado = _
             LP_Cual Then
               GV_VENTANAS_Usuario(L_i).OT_SocketAsociado = 0
               GV_VENTANAS_Usuario(L_i).Caption = _
               GV_VENTANAS_Usuario(L_i).OT_Nick + _
               ":Connection Finished/La conexión ha " + _
               "sido terminada"
               MV_Pone_Mensaje True, _
               GV_VENTANAS_Usuario(L_i), _
               "Connection Finished/La conexión " + _
               "ha sido terminada" + GV_EOD, vbRed
            End If
        End If
        DoEvents
Next L_i

' Cantidad de Ventanas de Usuario
L_Cuantos = UBound(GV_VENTANAS_Lista_Canales)

For L_i = 1 To L_Cuantos
 If Not GV_Estado_Lista_Canales(L_i).Deleted Then
  If GV_VENTANAS_Lista_Canales(L_i).OT_SocketAsociado = _
   LP_Cual Then
     GV_VENTANAS_Lista_Canales(L_i).OT_SocketAsociado = 0
     GV_VENTANAS_Lista_Canales(L_i).Caption = _
     "Lista Canales:Connection Finished/La conexión " + _
     "ha sido terminada"
  End If
 End If
 DoEvents
Next L_i


Exit Sub
Etiqueta_Error:
ME_Muestra_Error
End Sub

Function MV_Indice_Libre_Estatus() As Integer
' /*************************************************/
' Función que se encarga de buscar una casilla libre
' en el arreglo de Ventanas de
' Estatus.
' /*************************************************/
    
    Dim i As Integer
    Dim ArrayCount As Integer

    ArrayCount = UBound(GV_VENTANAS_Estatus)

    For i = 1 To ArrayCount
        If GV_Estado_Estatus(i).Deleted Then
            MV_Indice_Libre_Estatus = i
            GV_Estado_Estatus(i).Deleted = False
            Exit Function
        End If
    Next
    
    ReDim Preserve GV_VENTANAS_Estatus(ArrayCount + 1)
    ReDim Preserve GV_Estado_Estatus(ArrayCount + 1)
    MV_Indice_Libre_Estatus = UBound(GV_VENTANAS_Estatus)
End Function

Function MV_Busca_Ventana_Usuario(Lp_nickbuscado$, Lp_Cualsocket%) As Integer
' /***************************************************/
' Procedimiento que se encarga de buscar la ventana
' de usuario que corresponde a un socket asociado,
' utilizando el alias del usuario buscado. Si
' encuentra la ventana  entonces retorna el indice
' donde se encontro la ventana de usuario en el
' arreglo de ventanas.  De lo contrario retorna 0.
'
' LP_nickbuscado : El alias con el que se desea buscar
' la ventana
' LP_cualsocket : El alias con el que se desea buscar
' la ventana
' /***************************************************/

On Error GoTo Etiqueta_Error:

Dim L_i%, L_Cuantos%

' Cantidad de Ventanas de Usuario
L_Cuantos = UBound(GV_VENTANAS_Usuario)

For L_i = 1 To L_Cuantos
  If Not GV_Estado_Usuario(L_i).Deleted Then
    If GV_VENTANAS_Usuario(L_i).OT_SocketAsociado = _
     Lp_Cualsocket And _
     UCase(GV_VENTANAS_Usuario(L_i).OT_Nick) = _
     UCase(Lp_nickbuscado) Then
         MV_Busca_Ventana_Usuario = L_i
         Exit Function
   End If
 End If
Next L_i

MV_Busca_Ventana_Usuario = 0

Exit Function
Etiqueta_Error:
ME_Muestra_Error
End Function

Function MV_Indice_Libre_Usuario()
' /***********************************************/
' Función que se encarga de buscar una casilla
' libre en el arreglo de
' ventanas de usuarios.
' /***********************************************/
   
   Dim L_i As Integer
   Dim L_ArrayCount As Integer

    L_ArrayCount = UBound(GV_VENTANAS_Usuario)

  ' Recorrer el arreglo de ventanas. Si la ventana ha sido
  ' borrada entonces, retorna ese indice.
    For L_i = 1 To L_ArrayCount
        If GV_Estado_Usuario(L_i).Deleted Then
            MV_Indice_Libre_Usuario = L_i
            GV_Estado_Usuario(L_i).Deleted = False
            Exit Function
        End If
    Next

    ' Si ninguno de los elementos del arreglo han sido
    ' borrados entonces se crea una nueva casilla en el
    ' arreglo, redimensionandolo
    ' y retorna el nuevo indice.
    ReDim Preserve GV_VENTANAS_Usuario(L_ArrayCount + 1)
    ReDim Preserve GV_Estado_Usuario(L_ArrayCount + 1)
    MV_Indice_Libre_Usuario = UBound(GV_VENTANAS_Usuario)
End Function

Function MV_Obtener_Archivo(LP_Archivo$) As String
' /****************************************************/
' Retorna el nombre de un archivo sin su respectivo path.
' Ej: C:\windows\system\archivo.dll es retornado como
' archivo.dll
'
' LP_archivo: contiene el archivo y su path.
' /***************************************************/

Dim L_pos%, L_Largo%, L_i%

L_pos = InStr(1, LP_Archivo, "\")
If L_pos = 0 Then _
 MV_Obtener_Archivo = LP_Archivo: Exit Function
L_Largo = Len(LP_Archivo)
For L_i = 1 To L_Largo
  If Trim(Mid(Right(LP_Archivo, L_i), 1, 1)) = "\" Then
    MV_Obtener_Archivo = Right(LP_Archivo, L_i - 1)
    Exit Function
  End If

Next L_i
End Function


