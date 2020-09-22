Attribute VB_Name = "MO_Datos"
Option Explicit
'ERG2: Declaración de la Base de Datos

Global GV_Base_De_Datos As Database
Global GV_LOG As Database

' Declaración de la estructura de Datos de los usuarios

Type ES_USUARIO

E_Host As String
E_IP As String
E_Nombre As String
E_Alias As String
E_nombre_alterno As String
E_EMAIL As String


End Type
Sub MD_Registra_Hostname(Lp_HostName$)
' /********************************************************/
' Procedimiento que registra el hostname en los datos
' del usuario
' /********************************************************/
On Error GoTo Etiqueta_Error:

Dim L_Registro As Recordset

Set L_Registro = GV_Base_De_Datos.OpenRecordset( _
                 "Select * from Usuario", dbOpenDynaset)

L_Registro.LockEdits = False

' Pasa la información de del Host
If Not L_Registro.EOF Then
   L_Registro.Edit
   L_Registro!Host = Lp_HostName
   L_Registro.Update  ' Actualiza el registro
End If

L_Registro.Close

Exit Sub
Etiqueta_Error:
ME_Muestra_Error

End Sub

Function MD_Seleccion_Item(LP_CL As Control, LP_Cual&) As Integer
' /********************************************************/
' Seleciona un item de una lista o combo que tenga itemdata
' el valor en el parámetro LP_Cual.
' /********************************************************/

On Error GoTo Etiqueta_Error:
 Dim L_i%
 
 Do While L_i < LP_CL.ListCount
   If LP_CL.ItemData(L_i) = LP_Cual Then Exit Do
   L_i = L_i + 1
 Loop
 If L_i < LP_CL.ListCount Then
    MD_Seleccion_Item = L_i
 Else
    MD_Seleccion_Item = -1
 End If
Exit Function
Etiqueta_Error:
ME_Muestra_Error

End Function

Function MD_Seleccion_String(LP_CL As Control, LP_Cual$, LP_NC%) As Integer
On Error GoTo Etiqueta_Error:
' /********************************************************/
' Seleciona un item de una lista o combo que tenga de
' string en los primeros NC  caracteres el parámetro LP_Cual.
' /********************************************************/

 Dim L_i%
 
 Do While L_i < LP_CL.ListCount
   If Left(LP_CL.List(L_i), LP_NC) = LP_Cual Then Exit Do
   L_i = L_i + 1
 Loop
 If L_i < LP_CL.ListCount Then
    MD_Seleccion_String = L_i
 Else
    MD_Seleccion_String = -1
 End If

Exit Function
Etiqueta_Error:
ME_Muestra_Error
End Function

Sub MD_Actualiza_Ultimo_Servidor(LP_Codigo_Ultimo_Servidor As Long)
' /********************************************************/
' Guarda en la tabla Ini el último servidor al que un
' usuario intentó conectarse por última vez.  Esto con el
' fín de usarlo como servidor por omisión al momento
' de intentar una nueva conexión.
' /********************************************************/
On Error GoTo Etiqueta_Error:

Dim L_Ini As Recordset

Set L_Ini = GV_Base_De_Datos.OpenRecordset("Ini", dbOpenDynaset)
L_Ini.Edit
L_Ini("Ultimo_Servidor") = LP_Codigo_Ultimo_Servidor
L_Ini.Update
L_Ini.Close

Exit Sub
Etiqueta_Error:
ME_Muestra_Error
End Sub

Function MD_Obtener_Comando(LP_Texto$) As String
' /********************************************************/
' Función que busca la traducción en inglés de un comando
' enviado en español.
' /********************************************************/

On Error GoTo Etiqueta_Error:
Dim L_Comando As Recordset

Set L_Comando = GV_Base_De_Datos.OpenRecordset( _
                "select comando_I from comandos " + _
                "where ucase(comando_i)= '" + _
                UCase(LP_Texto) + "' or ucase(comando_E)= '" + _
                UCase(LP_Texto) + "'", dbOpenSnapshot)
                
If L_Comando.EOF Then
    MD_Obtener_Comando = LP_Texto
Else
    MD_Obtener_Comando = Trim(L_Comando("comando_I"))
End If
L_Comando.Close

Exit Function
Etiqueta_Error:
ME_Muestra_Error
End Function
Sub MD_Registrar_Log(LP_Accion$)
' /********************************************************/
' Guarda todos los errores en una tabla de bitácoras.
' /********************************************************/

On Error GoTo Etiqueta_Error:

Dim L_Registro As Recordset

Set L_Registro = GV_LOG.OpenRecordset("LOG")
L_Registro.AddNew
L_Registro!Fecha = Date
L_Registro!Hora = Time
L_Registro!Accion = LP_Accion

L_Registro.Update

L_Registro.Close

Exit Sub
Etiqueta_Error:
ME_Muestra_Error

End Sub

Function MD_Abrir_Base_Datos(LP_Base$, LP_Objeto As Database) As Long
' /********************************************************/
' Función que abre la base de datos
' Recibe de parametro el string donde tiene que buscar
' el archivo de la base de datos.
' Y ademas para hacerlo generico recibe de parametro el
' objeto de Visual Basic que representa la BD.
' /********************************************************/

On Error GoTo Etiqueta_Error:

Set LP_Objeto = OpenDatabase(LP_Base)

MD_Abrir_Base_Datos = 0  ' La base de datos fue abierta con exito


Exit Function


Etiqueta_Error:

ME_Muestra_Error
' La Base de datos no se pudo abrir, codigo de error Err
MD_Abrir_Base_Datos = Err

End Function

Function MD_Actualiza_Infousuario(LP_Usuario As ES_USUARIO) As Long
' /**********************************************************/
' Función que actualiza el registro de la información del
' usuario . Actualmente la tabla solo posee un registro
' /**********************************************************/

On Error GoTo Etiqueta_Error:

Dim L_Registro As Recordset

Set L_Registro = GV_Base_De_Datos.OpenRecordset( _
                 "Select * from Usuario", dbOpenDynaset)

L_Registro.LockEdits = False

' Pasa la información de la estructura temporal al archivo
If Not L_Registro.EOF Then
   L_Registro.Edit
   L_Registro!Host = LP_Usuario.E_Host
   L_Registro!IP = LP_Usuario.E_IP
   L_Registro!Nombre = LP_Usuario.E_Nombre
   L_Registro!Alias = LP_Usuario.E_Alias
   L_Registro!nombre_alterno = LP_Usuario.E_nombre_alterno
   L_Registro!EMail = LP_Usuario.E_EMAIL
     
   L_Registro.Update  ' Actualiza el registro
End If

L_Registro.Close
MD_Actualiza_Infousuario = 0 ' Si no hubo error se retorna 0

Exit Function

Etiqueta_Error:

ME_Muestra_Error
MD_Actualiza_Infousuario = Err ' Hubo error Err

End Function

Function MD_Cerrar_Base_Datos(LP_BASE_DATOS As Database) As Long
' /**********************************************************/
' Función que se encarga de cerrar la Base de Datos
' /**********************************************************/

On Error GoTo Etiqueta_Error:
  
LP_BASE_DATOS.Close ' Cierra la Base de Datos

MD_Cerrar_Base_Datos = 0 ' Base de Datos cerrada


Exit Function
Etiqueta_Error:

ME_Muestra_Error
MD_Cerrar_Base_Datos = Err ' Ocurrió un error

  
End Function

Function MD_Recupera_Infousuario(LP_Usuario As ES_USUARIO) As Long
' /**********************************************************/
' Esta Función recupera la información del usuario
' Devuelve el codigo de error si este existe o sino devuelve 0
' /**********************************************************/

On Error GoTo Etiqueta_Error:

Dim L_Registro As Recordset

Set L_Registro = GV_Base_De_Datos.OpenRecordset( _
                 "Select * from Usuario") ' Busca el registro
                 
                 
If Not L_Registro.EOF Then
 'Carga el registro a la estructura de Usuario
 LP_Usuario.E_Host = L_Registro!Host
 LP_Usuario.E_IP = L_Registro!IP
 LP_Usuario.E_Nombre = L_Registro!Nombre
 LP_Usuario.E_Alias = L_Registro!Alias
 LP_Usuario.E_nombre_alterno = L_Registro!nombre_alterno
 LP_Usuario.E_EMAIL = L_Registro!EMail


End If
L_Registro.Close

' Recuperó la  información sin nungún error
MD_Recupera_Infousuario = 0
Exit Function

Etiqueta_Error:

ME_Muestra_Error
MD_Recupera_Infousuario = Err ' Ocurrió un error


End Function

Function MD_Actualizar_Tabla(LP_Datos() As Control, LP_Tabla As Recordset, LP_Es_Nuevo As Integer)
' /***********************************************************/
'   LP_Datos: Arreglo de objetos que contienen los datos que
'   se desean agregar a un nuevo registro o modificar en uno
'   ya existente.
'   LP_Tabla: Tabla a la que se agregará o modificará un registro.
'   LP_Es_Nuevo : Variable utilizada para indicar si se creará
'   o modificará un registro
'   Esta función es utilizada para Agregar o Modificar datos a
'   un registro de una tabla.
' /************************************************************/

On Error GoTo Etiqueta_Error:

Dim L_Limite_Campos%, L_i%, L_S$, L_Numero%

L_Limite_Campos = UBound(LP_Datos, 1)
If LP_Es_Nuevo = 4 Then LP_Tabla.AddNew
For L_i = 0 To L_Limite_Campos - 1
  If Not IsNull(LP_Datos(L_i)) Then
   If 0 = InStr(LP_Datos(L_i).Tag, "%") Then
      If Not (InStr(LP_Datos(L_i).Tag, "&") <> 0 And _
        LP_Es_Nuevo <> 4) Then
          L_S = MD_Obtener_String("[", (LP_Datos(L_i).Tag), "]")
          L_Numero = _
            Val(MD_Obtener_String("$", (LP_Datos(L_i).Tag), "$"))
          If TypeOf LP_Datos(L_i) Is TextBox Then
            If L_Numero Then
              LP_Tabla(L_S) = Left$(LP_Datos(L_i), L_Numero)
            Else
              LP_Tabla(L_S) = LP_Datos(L_i)
            End If
          
          ElseIf TypeOf LP_Datos(L_i) Is ComboBox Then
            If LP_Datos(L_i).ListIndex >= 0 Then
              If L_Numero Then
                LP_Tabla(L_S) = Left$(LP_Datos(L_i), L_Numero)
              Else
                LP_Tabla(L_S) = _
                LP_Datos(L_i).ItemData(LP_Datos(L_i).ListIndex)
              End If
            End If
           ElseIf TypeOf LP_Datos(L_i) Is ListBox Then
            If LP_Datos(L_i).ListIndex >= 0 Then
              If L_Numero Then
                LP_Tabla(L_S) = _
                Left$(LP_Datos(L_i).ItemData(LP_Datos(L_i).ListIndex), _
                L_Numero)
              Else
                LP_Tabla(L_S) = LP_Datos(L_i)
              End If
            End If
           ElseIf TypeOf LP_Datos(L_i) Is CheckBox Then
             LP_Tabla(L_S) = LP_Datos(L_i).Value
           ElseIf TypeOf LP_Datos(L_i) Is Menu Then
             LP_Tabla(L_S) = LP_Datos(L_i).Checked
           
            End If
            End If
        End If
    End If
Next L_i
LP_Tabla.Update


Exit Function
Etiqueta_Error:
ME_Muestra_Error
End Function

Function MD_Cargar_Datos(LP_Datos() As Control, LP_Tabla As Recordset, MAKE_EDIT%)
' /************************************************************/
' LP_Datos: Arreglo de objetos que contienen los campos cuyos
' datos se desean cargar a los controles de una forma.
' LP_Tabla: Tabla de la que se cargarán los datos.
' Retorna el número de error de visual basic en caso de que haya
' error, o 0 si no existe error.
'  Descripción :
' -------------
'  Esta función es utilizada para cargar datos a los controles
'  definidos en una  forma.
' /************************************************************/

On Error GoTo Etiqueta_Error:

Dim L_Limite_Campos%, L_i%, NUM_ERR, L_S$, L_Numero%
Dim L_Fecha As String

'If MAKE_EDIT% Then LP_Tabla.LockEdits = LOCK_EDITS:
LP_Tabla.Edit
L_Limite_Campos = UBound(LP_Datos, 1)
For L_i = 0 To L_Limite_Campos - 1
 If Not IsNull(LP_Datos(L_i)) Then
   If InStr(LP_Datos(L_i).Tag, "@") Then
     L_S = MD_Obtener_String("[", (LP_Datos(L_i).Tag), "]")
     L_Numero = Val(MD_Obtener_String("$", (LP_Datos(L_i).Tag), "$"))
      If TypeOf LP_Datos(L_i) Is TextBox Then
        LP_Datos(L_i) = "" & LP_Tabla(L_S)
      
      ElseIf TypeOf LP_Datos(L_i) Is ComboBox Then
       If L_Numero Then _
        LP_Datos(L_i).ListIndex = _
        MD_Seleccion_String(LP_Datos(L_i), (LP_Tabla(L_S)), L_Numero)
        If LP_Datos(L_i).ListIndex = -1 Then NUM_ERR = NUM_ERR + 1
        Else
          LP_Datos(L_i).ListIndex = _
          MD_Seleccion_Item(LP_Datos(L_i), (LP_Tabla(L_S)))
          If LP_Datos(L_i).ListIndex = -1 Then NUM_ERR = NUM_ERR + 1
        End If
     ElseIf TypeOf LP_Datos(L_i) Is ListBox Then
         If L_Numero Then
          LP_Datos(L_i).ListIndex = _
          MD_Seleccion_String(LP_Datos(L_i), (LP_Tabla(L_S)), L_Numero)
          If LP_Datos(L_i).ListIndex = -1 Then NUM_ERR = NUM_ERR + 1
          Else
             LP_Datos(L_i).ListIndex = _
             MD_Seleccion_Item(LP_Datos(L_i), (LP_Tabla(L_S)))
             If LP_Datos(L_i).ListIndex = -1 Then NUM_ERR = NUM_ERR + 1
          End If
         End If
     ElseIf TypeOf LP_Datos(L_i) Is CheckBox Then
          LP_Datos(L_i).Value = LP_Tabla(L_S)
     ElseIf TypeOf LP_Datos(L_i) Is Menu Then
          LP_Datos(L_i).Checked = LP_Tabla(L_S)
     
     End If
    
  
Next L_i
MD_Cargar_Datos = NUM_ERR
Exit Function
Etiqueta_Error:
ME_Muestra_Error
End Function

Function MD_Convertir_Fecha(LP_Fecha As String) As String
' /************************************************************/
'  LP_Fecha    : Fecha a convertir
'  Retorna la fecha en formato dd/mm/yyyy.
'  Convierte fecha a formato día/mes/año (dd/mm/yyyy).
' /************************************************************/

On Error GoTo Etiqueta_Error:

If Mid(LP_Fecha, 3, 1) <> "/" Then LP_Fecha = 0 & LP_Fecha
If Mid(LP_Fecha, 6, 1) <> "/" Then LP_Fecha = Left(LP_Fecha, 3) & _
0 & Right(LP_Fecha, 4)

MD_Convertir_Fecha = LP_Fecha

Exit Function
Etiqueta_Error:
ME_Muestra_Error
End Function

Sub MD_Limpiar_Datos(LP_Datos() As Control)
' /************************************************************/
' LP_Datos: Arreglo de objetos que contienen los controles
' cuyos datos se desean limpiar.
' Esta función es utilizada limpiar (poner en blanco)los
' controles enviados en un arreglo.
' /************************************************************/

On Error GoTo Etiqueta_Error:

Dim L_Limite_Campos%, L_i%
L_Limite_Campos% = UBound(LP_Datos, 1)
For L_i = 0 To L_Limite_Campos% - 1
    If Not IsNull(LP_Datos(L_i)) Then
        If InStr(LP_Datos(L_i).Tag, "@") Then
            If TypeOf LP_Datos(L_i) Is TextBox Then
                LP_Datos(L_i) = ""
            
            ElseIf TypeOf LP_Datos(L_i) Is ComboBox Then
                LP_Datos(L_i).ListIndex = -1
            ElseIf TypeOf LP_Datos(L_i) Is ListBox Then
                LP_Datos(L_i).ListIndex = -1
            ElseIf TypeOf LP_Datos(L_i) Is CheckBox Then
                LP_Datos(L_i).Value = False
            ElseIf TypeOf LP_Datos(L_i) Is Menu Then
                LP_Datos(L_i).Checked = False
            'ElseIf TypeOf LP_Datos(L_I) Is SSCheck Then
                LP_Datos(L_i).Value = False
            ElseIf TypeOf LP_Datos(L_i) Is OptionButton Then
                LP_Datos(L_i).Value = False ' TODAVIA FALTA
            'ElseIf TypeOf LP_Datos(L_I) Is SSOption Then
                LP_Datos(L_i).Value = False ' TODAVIA FALTA
            End If
        End If
    End If
Next L_i

Exit Sub
Etiqueta_Error:
ME_Muestra_Error
End Sub
Function MD_Obtener_String(LP_String_Encontrar1$, LP_String_Buscar$, LP_String_Encontrar2$) As String
' /************************************************************/
' Función que retorna un substring de LP_String_Buscar limitado
' por LP_String_Encontrar1 y  LP_String_Encontrar2
' /************************************************************/

On Error GoTo Etiqueta_Error:

Dim L_Inicio%, L_Largo%
L_Inicio = InStr(LP_String_Buscar, LP_String_Encontrar1) + 1
L_Largo = InStr(L_Inicio, _
           LP_String_Buscar, LP_String_Encontrar2) - L_Inicio
           
If L_Largo < 1 Then Exit Function
MD_Obtener_String = Mid(LP_String_Buscar, L_Inicio, L_Largo)

Exit Function
Etiqueta_Error:
ME_Muestra_Error
End Function

Function MD_Cargar_Lista(LP_registros, LP_Lista)
' /************************************************************/
' Este procedimiento se utiliza para cargar datos del recordset
' LP_Tabla en una lista o combo cualquiera.  La lista o combo
' viene en LP_Lista
' /************************************************************/

On Error GoTo Etiqueta_Error:

LP_Lista.Clear
If LP_registros.EOF Then
    MD_Cargar_Lista = 0
Else
    While Not LP_registros.EOF
        LP_Lista.AddItem LP_registros(0)
        LP_Lista.ItemData(LP_Lista.NewIndex) = LP_registros(1)
        LP_registros.MoveNext
        DoEvents
    Wend
    LP_Lista.ListIndex = 0
    MD_Cargar_Lista = 1
End If

Exit Function
Etiqueta_Error:
ME_Muestra_Error
End Function

Function MD_Validar_Nulos(LP_Datos() As Control) As Integer
' /************************************************************/
' Función que se encarga de encontrar los campos de los
' controles que se encuentran nulos o vacíos pero que no
' deberían serlo.
' LP_Datos contiene los controles en los que se deben buscar
' los espacios vacíos o nulos.
' /************************************************************/

Dim L_Limite_Campos%, L_i%, s$, L_Num%

L_Limite_Campos = UBound(LP_Datos, 1)
For L_i = 0 To L_Limite_Campos - 1
 MD_Validar_Nulos = L_i
 If Not IsNull(LP_Datos(L_i)) Then
    s = MD_Obtener_String("^", (LP_Datos(L_i).Tag), "^")
    If s <> "" Then
      L_Num = _
      Val(MD_Obtener_String("$", (LP_Datos(L_i).Tag), "$"))
      MD_Validar_Nulos = L_i
       If TypeOf LP_Datos(L_i) Is TextBox Then
            If LP_Datos(L_i) = "" Then Exit Function
        
        ElseIf TypeOf LP_Datos(L_i) Is ComboBox Then
            If L_Num Then
                If LTrim(LP_Datos(L_i)) = "" Then Exit Function
            Else
                If LP_Datos(L_i).ListIndex < 0 Then Exit Function
            End If
        ElseIf TypeOf LP_Datos(L_i) Is ListBox Then
            If L_Num Then
                If LTrim(LP_Datos(L_i)) = "" Then Exit Function
            Else
                If LP_Datos(L_i).ListIndex < 0 Then Exit Function
            End If
        ElseIf TypeOf LP_Datos(L_i) Is CheckBox Then
            ' NO SE PUEDE VALIDAR
        ElseIf TypeOf LP_Datos(L_i) Is Menu Then
            ' NO SE PUEDE VALIDAR
        ElseIf TypeOf LP_Datos(L_i) Is OptionButton Then
            'NO SE PUEDE VALIDAR
        End If
    End If
 End If
Next L_i
 MD_Validar_Nulos = -1
End Function

