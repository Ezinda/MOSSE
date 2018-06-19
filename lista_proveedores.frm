VERSION 5.00
Begin VB.Form lista_proveedores 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Proveedores"
   ClientHeight    =   4830
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6315
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4830
   ScaleWidth      =   6315
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List1 
      Height          =   4155
      ItemData        =   "lista_proveedores.frx":0000
      Left            =   240
      List            =   "lista_proveedores.frx":0002
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   240
      Width           =   5775
   End
End
Attribute VB_Name = "lista_proveedores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public lista As String
Public posicion As Integer


Private Sub Form_Load()
Aplicar_skin Me

MiFuncionDeAjuste Me, True

If login.empresaact1 > 0 Then
    empresareal = login.empresaact1
Else
    empresareal = login.empresaact
End If

bases.datbasemenu.ConnectionString = login.conexiontotal
If ventana.menu = 1 Or ventana.menu = 3 Then
    lista_proveedores.Caption = "Proveedores"
    bases.datbasemenu.RecordSource = "select proveedores.* from proveedores where empresa = " & empresareal & " order by razonsocial"
    bases.datbasemenu.Refresh
    If bases.datbasemenu.Recordset.EOF = True Then
        List1.AddItem "Base de datos Proveedores Vacia"
        Exit Sub
    End If

    posicion = 0
    bases.datbasemenu.Recordset.MoveFirst
    Do While Not bases.datbasemenu.Recordset.EOF
         If IsNull(bases.datbasemenu.Recordset.Fields("razonsocial")) = False Then
            List1.AddItem bases.datbasemenu.Recordset.Fields("razonsocial")
        End If
        bases.datbasemenu.Recordset.MoveNext
    Loop
End If

If ventana.menu = 2 Then
    lista_proveedores.Caption = "Clientes"
    bases.datbasemenu.RecordSource = "select clientes.* from clientes where empresa = " & empresareal & " order by razonsocial"
    bases.datbasemenu.Refresh
    If bases.datbasemenu.Recordset.EOF = True Then
        List1.AddItem "Base de datos Clientes Vacia"
        Exit Sub
    End If

    posicion = 0
    bases.datbasemenu.Recordset.MoveFirst
    Do While Not bases.datbasemenu.Recordset.EOF
         If IsNull(bases.datbasemenu.Recordset.Fields("razonsocial")) = False Then
            List1.AddItem bases.datbasemenu.Recordset.Fields("razonsocial")
        End If
        bases.datbasemenu.Recordset.MoveNext
    Loop
End If


SendKeys "{tab}", False
 
End Sub

Private Sub List1_DblClick()

        If ventana.menu = 1 Then
            lista = List1.Text
            posicion = List1.ListIndex
            frmproveedores.flagnuevo = 1
            frmproveedores.SetFocus
        End If
        If ventana.menu = 2 Then
            lista = List1.Text
            posicion = List1.ListIndex
            frmclientes.flagnuevo = 1
            frmclientes.SetFocus
        End If
        If ventana.menu = 3 Then
            lista = List1.Text
            frmctacte.SetFocus
        End If
        Unload Me
        
End Sub

Private Sub List1_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        If ventana.menu = 1 Then
            lista = List1.Text
            posicion = List1.ListIndex
            frmproveedores.flagnuevo = 1
            frmproveedores.SetFocus
        End If
        If ventana.menu = 2 Then
            lista = List1.Text
            posicion = List1.ListIndex
            frmclientes.flagnuevo = 1
            frmclientes.SetFocus
        End If
        If ventana.menu = 3 Then
            lista = List1.Text
            frmctacte.SetFocus
        End If
        Unload Me
    End If
    

End Sub
