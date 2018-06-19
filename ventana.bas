Attribute VB_Name = "ventana"
Const HWND_TOPMOST = -1 'Situa el form encima de todos
Const HWND_NOTOPMOST = -2 'Vuelve al estado normal
Const SWP_NOSIZE = &H1 'SWP_NOSIZE y SWP_NOMOVE se usan juntas en la constante SWP_FLAGS para que el form no cambie de tamaño ni de lugar
Const SWP_NOMOVE = &H2
Const SWP_FLAGS = SWP_NOMOVE Or SWP_NOSIZE 'Para que no cambie el tamaño ni la posición

Public menu As Integer
Public clientefacctacte As String
Public query As String
Public xclaveprimariaremito As Double
Public xfila As Integer
Public remdev As Integer
Public tipodeventa As Integer
Private Declare Function SetWindowPos Lib "User32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Public Declare Function GetSystemMenu Lib "User32" _
(ByVal hWnd As Long, ByVal bRevert As Long) As Long

Public Declare Function RemoveMenu Lib "User32" _
(ByVal hMenu As Long, ByVal nPosition As Long, _
ByVal wFlags As Long) As Long


'A esta función le enviamos la ventana o formulario con un valor boolean
'indicando si queremos o no dejar siempre visible la ventana
Public Sub MiFuncionDeAjuste(Formulario As Form, Estado As Boolean)

'Variable de retorno de la función
Dim retorno As Long

If Estado = True Then
'Siempre Visible - always On Top
retorno = SetWindowPos(Formulario.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_FLAGS)
Else
'Ventana normal
retorno = SetWindowPos(Formulario.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_FLAGS)
End If

End Sub

