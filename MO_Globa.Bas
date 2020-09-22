Attribute VB_Name = "MO_Global"
Option Explicit
'ERG2:   Constantes para el programa
' Modulo GLobal que posee declaraciones y funciones generales
' al programa

Global Const GC_Sistema = "Sistema ERG2 Version 1.03 "
Global Const GC_Autores = "Rogger Vasquez, Gustavo Oviedo"
Global Const GC_Licencia = "UNITEC"
Global Const GC_Institucion = "UNITEC"
Global Const GC_Serie = "911117-921153"

' Variable que nos indica el ultimo servidor de donde se hizo
' una llamada a otra ventana
Global GV_Seleccion As Integer
Global GV_Nombre As String
Global GV_DoEVENTS As Boolean

' Definición de Colores
Global GV_Verde As Long
Global GV_Azul As Long
Global GV_RojoAlgo As Long
Global GV_Rojo As Long
Global GV_Morado As Long
Global GV_Negro As Long

'  Constantes de Eventos en un una Ventana
'  Tomadas del API de Windows

Public Const WM_SIZE = &H5
Public Const WM_CTLCOLOREDIT = &H133
Public Const WM_CUT = &H300
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_MBUTTONDBLCLK = &H209
Public Const WM_MBUTTONDOWN = &H207
Public Const WM_MBUTTONUP = &H208
Public Const WM_MDIACTIVATE = &H222
Public Const WM_MDICASCADE = &H227
Public Const WM_MDICREATE = &H220
Public Const WM_MDIDESTROY = &H221
Public Const WM_MDIGETACTIVE = &H229
Public Const WM_USUARIO = &H2000
Public Const WM_USER = &H400


' Estructura de Retransmisión la cual nos indica cuales
' mensajes
' tienen que ser retransmitidos y hacia donde
Type ES_RETRANSMISION
    E_Servidor_Emisor As Integer
    E_Alias_Emisor As String
    E_Servidor_Receptor As Integer
    E_Prefijo As String
    E_Sufijo As String
    E_SoloMensaje As Boolean
    E_Omitir_Mensaje As Boolean
    E_Borrado As Boolean 'true = borrado
End Type

' Estructura de Puertos disponibles para enviar Archivos
Type Es_Puertos
  E_Puerto As Integer
  E_Libre As Boolean ' True = Libre , False = Ocupado
End Type


Global GV_Puertos(5) As Es_Puertos

' Variable utilizada para determinar End of Data <CR LF>
Global GV_Estructura_Retransmision() As ES_RETRANSMISION
' Variable utilizada para determinar End of Data <CR LF>
Global GV_EOD$


Sub MG_Barra(LP_Texto As String)
' /******************************************************/
' Procedimiento que muestra un mensaje en la barra de
' estatus del programa
'/ ******************************************************/
  
  OF_principal.StatusBar.SimpleText = LP_Texto
End Sub

Sub MG_Mensaje(mensaje As String)
' /******************************************************/
' Procedimiento que muestra un mensaje de "OK" con el
' nombre del sistema como titulo
' /******************************************************/

 Load OF_Mensajes
 OF_Mensajes.OT_Mensaje = mensaje
 OF_Mensajes.Caption = "<Mensajes> : " + GC_Sistema
 OF_Mensajes.Show 1

 End Sub

Sub MG_Borra_Primero(ByRef LP_Historial() As String)

' /******************************************************/
' Borra el primer elemento del arreglo moviendo todos los
' items uno a a la izquierda, este procedimiento es
' utilizado para borrar el primer
' Elemento del arreglo de historial de comandos
' /******************************************************/

Dim L_i%

For L_i = 1 To 19
  
  LP_Historial(L_i) = LP_Historial(L_i + 1)
  DoEvents
Next L_i

End Sub

Function MG_Pregunta(mensaje As String) As Integer
' /*****************************************************/
' Función utilizada para realizar una pregunta dentro del
' programa
' Ej.  Desea Continuar ? '
' /*****************************************************/

MG_Pregunta = MsgBox(mensaje, vbYesNo + vbQuestion + _
vbDefaultButton1, GC_Sistema)
End Function

Sub MG_Rellena_Espacios(ByRef LP_String$, LP_Hasta%)
' /*****************************************************/
' Procedimiento que rellena de espacios en blanco una
' variable hasta un numeroespecificado por el parametro
' LP_Hasta
' /******************************************************/

Dim L_i%, L_j%
Dim L_Largo%

L_Largo = Len(LP_String)
If LP_Hasta < L_Largo Then Exit Sub

Select Case L_Largo
Case 1 To 10
   L_j = 0
Case 11 To 25
   L_j = 1
Case Else
  L_j = 0
End Select
For L_i = 1 To LP_Hasta - L_Largo - (L_j * 10)
  LP_String = LP_String + " "
Next L_i
End Sub
