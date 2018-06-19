VERSION 5.00
Begin VB.Form lista_condtrib 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cond.Tributaria"
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
      ItemData        =   "lista_condtrib.frx":0000
      Left            =   240
      List            =   "lista_condtrib.frx":0002
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   240
      Width           =   2775
   End
End
Attribute VB_Name = "lista_condtrib"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public lista As String
Public posicion As Integer


Private Sub Form_Load()
Aplicar_skin Me

MiFuncionDeAjuste Me, True

bases.datbasemenu1.ConnectionString = login.conexionps
bases.datbasemenu1.RecordSource = "select condtrib.* from condtrib order by categ"
bases.datbasemenu1.Refresh
If bases.datbasemenu1.Recordset.EOF = True Then
    List1.AddItem "Tabla de datos Condición Tributaria está Vacia"
    Exit Sub
End If

posicion = 0
bases.datbasemenu1.Recordset.MoveFirst
Do While Not bases.datbasemenu1.Recordset.EOF
     If IsNull(bases.datbasemenu1.Recordset.Fields("categ")) = False Then
        List1.AddItem bases.datbasemenu1.Recordset.Fields("categ") + " - " + bases.datbasemenu1.Recordset.Fields("descripcion")
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
            frmconftalcompras.SetFocus
        End If
        Unload Me
    End If
    

End Sub
