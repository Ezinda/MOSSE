VERSION 5.00
Begin VB.Form lista_condventa 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cond.Venta"
   ClientHeight    =   4830
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3300
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4830
   ScaleWidth      =   3300
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List1 
      Height          =   4155
      ItemData        =   "lista_condventa.frx":0000
      Left            =   240
      List            =   "lista_condventa.frx":0002
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   240
      Width           =   2775
   End
End
Attribute VB_Name = "lista_condventa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public lista As String
Public posicion As Integer


Private Sub Form_Load()
Aplicar_skin Me

MiFuncionDeAjuste Me, True

bases.datbasemenu1.ConnectionString = login.conexiontotal
bases.datbasemenu1.RecordSource = "select condventas.* from condventas order by codigo"
bases.datbasemenu1.Refresh
If bases.datbasemenu1.Recordset.EOF = True Then
    List1.AddItem "Tabla de datos Condición de Ventas vacia"
    Exit Sub
End If

posicion = 0
bases.datbasemenu1.Recordset.MoveFirst
Do While Not bases.datbasemenu1.Recordset.EOF
     If IsNull(bases.datbasemenu1.Recordset.Fields("codigo")) = False Then
        List1.AddItem Str(bases.datbasemenu1.Recordset.Fields("codigo")) + " - " + bases.datbasemenu1.Recordset.Fields("descripcion")
    End If
    bases.datbasemenu1.Recordset.MoveNext
Loop
SendKeys "{tab}", False
 
End Sub

Private Sub List1_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        If ventana.menu = 1 Then
            lista = Left(List1.Text, 3)
            frmclientes.SetFocus
        End If
        Unload Me
    End If
    

End Sub
