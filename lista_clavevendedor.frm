VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form lista_clavevendedor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "        Clave Vendedor"
   ClientHeight    =   870
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   2805
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   870
   ScaleWidth      =   2805
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   2295
   End
   Begin MSAdodcLib.Adodc datvendedor 
      Height          =   330
      Left            =   240
      Top             =   600
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   582
      ConnectMode     =   3
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "contable"
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
End
Attribute VB_Name = "lista_clavevendedor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public cuentacont As String
Dim cuenta(99999) As Integer



Private Sub DataGrid1_DblClick()

        If menu = 1 Then
                frmnota_venta.Text1(0).Text = DataGrid1.Columns(2).Text
                SendKeys "{ENTER}", False
                Unload Me
        End If
        If menu = 2 Then
                frmpresupuesto.Text1(0).Text = DataGrid1.Columns(2).Text
                SendKeys "{ENTER}", False
                Unload Me
        End If
        If menu = 3 Then
                frmalquiler.Text1(0).Text = DataGrid1.Columns(2).Text
                SendKeys "{ENTER}", False
                Unload Me
        End If
        If menu = 4 Then
                frmfacctacte_alquiler.Text1(0).Text = DataGrid1.Columns(2).Text
                SendKeys "{ENTER}", False
                Unload Me
        End If
        If menu = 5 Then
                frmfacctacte_venta.Text1(0).Text = DataGrid1.Columns(2).Text
                SendKeys "{ENTER}", False
                Unload Me
        End If
        If menu = 6 Then
                frmnota_credito.Text1(0).Text = DataGrid1.Columns(2).Text
                SendKeys "{ENTER}", False
                Unload Me
        End If
        If menu = 7 Then
                frmnota_debito.Text1(0).Text = DataGrid1.Columns(2).Text
                SendKeys "{ENTER}", False
                Unload Me
        End If


End Sub

Private Sub DataGrid1_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
            If menu = 1 Then
                frmnota_venta.Text1(0).Text = DataGrid1.Columns(2).Text
                SendKeys "{ENTER}", False
            End If
            If menu = 2 Then
                frmpresupuesto.Text1(0).Text = DataGrid1.Columns(2).Text
                SendKeys "{ENTER}", False
            End If
            If menu = 3 Then
                frmalquiler.Text1(0).Text = DataGrid1.Columns(2).Text
                SendKeys "{ENTER}", False
            End If
            If menu = 4 Then
                frmfacctacte_alquiler.Text1(0).Text = DataGrid1.Columns(2).Text
                SendKeys "{ENTER}", False
            End If
            If menu = 5 Then
                frmfacctacte_venta.Text1(0).Text = DataGrid1.Columns(2).Text
                SendKeys "{ENTER}", False
            End If
            If menu = 6 Then
                frmnota_credito.Text1(0).Text = DataGrid1.Columns(2).Text
                SendKeys "{ENTER}", False
            End If
            If menu = 7 Then
                frmnota_debito.Text1(0).Text = DataGrid1.Columns(2).Text
                SendKeys "{ENTER}", False
            End If
            
        Unload Me
    End If

End Sub



Private Sub Form_Activate()
On Error Resume Next
DoEvents

'Text1.SetFocus



End Sub

Private Sub Form_Load()

'If menu = 2 Then
'    Aplicar_skin2 Me
    Aplicar_skin Me
'Else
'    Aplicar_skin Me
'End If

MiFuncionDeAjuste Me, True

datvendedor.ConnectionString = login.conexiontotal


If menu = 1 Then
    datvendedor.RecordSource = "select * from ud_ezi_empleado where id = '" & frmnota_venta.DataGrid1.Columns("id") & "'"
    datvendedor.Refresh
End If
If menu = 2 Then
    datvendedor.RecordSource = "select * from ud_ezi_empleado where id = '" & frmpresupuesto.DataGrid1.Columns("id") & "'"
    datvendedor.Refresh
End If
If menu = 3 Then
    datvendedor.RecordSource = "select * from ud_ezi_empleado where id = '" & frmalquiler.DataGrid1.Columns("id") & "'"
    datvendedor.Refresh
End If
If menu = 4 Then
    datvendedor.RecordSource = "select * from ud_ezi_empleado where id = '" & frmfacctacte_alquiler.DataGrid1.Columns("id") & "'"
    datvendedor.Refresh
End If
If menu = 5 Then
    datvendedor.RecordSource = "select * from ud_ezi_empleado where id = '" & frmfacctacte_venta.DataGrid1.Columns("id") & "'"
    datvendedor.Refresh
End If
If menu = 6 Then
    datvendedor.RecordSource = "select * from ud_ezi_empleado where id = '" & frmnota_credito.DataGrid1.Columns("id") & "'"
    datvendedor.Refresh
End If
If menu = 7 Then
    datvendedor.RecordSource = "select * from ud_ezi_empleado where id = '" & frmnota_debito.DataGrid1.Columns("id") & "'"
    datvendedor.Refresh
End If

If datvendedor.Recordset.EOF = True Then
    mensa = MsgBox("El vendedor no se encuentra dado de alta para operar en este sistema", vbCritical, "Error")
    If menu = 1 Then
        frmnota_venta.Text1(0).Text = ""
        frmnota_venta.Text1(0).SetFocus
    End If
    If menu = 2 Then
        frmpresupuesto.Text1(0).Text = ""
        frmpresupuesto.Text1(0).SetFocus
    End If
    If menu = 3 Then
        frmalquiler.Text1(0).Text = ""
        frmalquiler.Text1(0).SetFocus
    End If
    If menu = 4 Then
        frmfacctacte_alquiler.Text1(0).Text = ""
        frmfacctacte_alquiler.Text1(0).SetFocus
    End If
    If menu = 5 Then
        frmfacctacte_venta.Text1(0).Text = ""
        frmfacctacte_venta.Text1(0).SetFocus
    End If
    If menu = 6 Then
        frmnota_credito.Text1(0).Text = ""
        frmnota_credito.Text1(0).SetFocus
    End If
    If menu = 7 Then
        frmnota_debito.Text1(0).Text = ""
        frmnota_debito.Text1(0).SetFocus
    End If
    
    Unload Me
    Exit Sub
End If

'**** Control de vendedor habilitado
xhabilitado = datvendedor.Recordset.Fields("habilitado")
If xhabilitado <> 1 Then
    mensa = MsgBox("El vendedor no se encuentra HABILITADO para operar en este sistema", vbCritical, "Error")
    If menu = 1 Then
        frmnota_venta.Text1(0).Text = ""
        frmnota_venta.Text1(0).SetFocus
    End If
    If menu = 2 Then
        frmpresupuesto.Text1(0).Text = ""
        frmpresupuesto.Text1(0).SetFocus
    End If
    If menu = 3 Then
        frmalquiler.Text1(0).Text = ""
        frmalquiler.Text1(0).SetFocus
    End If
    If menu = 4 Then
        frmfacctacte_alquiler.Text1(0).Text = ""
        frmfacctacte_alquiler.Text1(0).SetFocus
    End If
    If menu = 5 Then
        frmfacctacte_venta.Text1(0).Text = ""
        frmfacctacte_venta.Text1(0).SetFocus
    End If
    If menu = 6 Then
        frmnota_credito.Text1(0).Text = ""
        frmnota_credito.Text1(0).SetFocus
    End If
    If menu = 7 Then
        frmnota_debito.Text1(0).Text = ""
        frmnota_debito.Text1(0).SetFocus
    End If
    
    Unload Me
    Exit Sub
End If

'*** control de permiso de acceso
If menu = 1 Then
    If datvendedor.Recordset.Fields("ventaingresar") = "N" Then
        mensa = MsgBox("El vendedor no tiene permisos para Realizar Ventas", vbCritical, "Error")
        frmnota_venta.Text1(0).Text = ""
        frmnota_venta.Text1(0).SetFocus
        Unload Me
        Exit Sub
    End If
End If
If menu = 2 Then
    If datvendedor.Recordset.Fields("presupingresar") = "N" Then
        mensa = MsgBox("El vendedor no tiene permisos para Realizar Ventas", vbCritical, "Error")
        frmpresupuesto.Text1(0).Text = ""
        frmpresupuesto.Text1(0).SetFocus
        Unload Me
        Exit Sub
    End If
End If
If menu = 3 Then
    If datvendedor.Recordset.Fields("alquileringresar") = "N" Then
        mensa = MsgBox("El vendedor no tiene permisos para Realizar Ventas", vbCritical, "Error")
        frmalquiler.Text1(0).Text = ""
        frmalquiler.Text1(0).SetFocus
        Unload Me
        Exit Sub
    End If
End If
If menu = 4 Then
    If datvendedor.Recordset.Fields("ventactacte") = "N" Then
        mensa = MsgBox("El vendedor no tiene permisos para Realizar Ventas", vbCritical, "Error")
        frmfacctacte_alquiler.Text1(0).Text = ""
        frmfacctacte_alquiler.Text1(0).SetFocus
        Unload Me
        Exit Sub
    End If
End If
If menu = 5 Then
    If datvendedor.Recordset.Fields("ventactacte") = "N" Then
        mensa = MsgBox("El vendedor no tiene permisos para Realizar Ventas", vbCritical, "Error")
        frmfacctacte_venta.Text1(0).Text = ""
        frmfacctacte_venta.Text1(0).SetFocus
        Unload Me
        Exit Sub
    End If
End If
If menu = 6 Then
    If datvendedor.Recordset.Fields("ventactacte") = "N" Then
        mensa = MsgBox("El vendedor no tiene permisos para Realizar Notas de Credito", vbCritical, "Error")
        frmnota_credito.Text1(0).Text = ""
        frmnota_credito.Text1(0).SetFocus
        Unload Me
        Exit Sub
    End If
End If
If menu = 7 Then
    If datvendedor.Recordset.Fields("ventactacte") = "N" Then
        mensa = MsgBox("El vendedor no tiene permisos para Realizar Notas de Debito", vbCritical, "Error")
        frmnota_debito.Text1(0).Text = ""
        frmnota_debito.Text1(0).SetFocus
        Unload Me
        Exit Sub
    End If
End If
 
 
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

    If datvendedor.Recordset.EOF = False Then
        If Text1.Text = datvendedor.Recordset.Fields("clave") Then
             If menu = 1 Then
                frmnota_venta.Text1(1).SetFocus
             End If
             If menu = 2 Then
                frmpresupuesto.Text1(1).SetFocus
             End If
             If menu = 3 Then
                frmalquiler.Text1(1).SetFocus
             End If
             If menu = 4 Then
                frmfacctacte_alquiler.Text1(1).SetFocus
             End If
             If menu = 5 Then
                frmfacctacte_venta.Text1(1).SetFocus
             End If
             If menu = 6 Then
                frmnota_credito.Text1(1).SetFocus
             End If
             If menu = 7 Then
                frmnota_debito.Text1(1).SetFocus
             End If
             
        Else
            If menu = 1 Then
                frmnota_venta.Text1(0).Text = ""
                frmnota_venta.Text1(0).SetFocus
            End If
            If menu = 2 Then
                frmpresupuesto.Text1(0).Text = ""
                frmpresupuesto.Text1(0).SetFocus
            End If
            If menu = 3 Then
                frmalquiler.Text1(0).Text = ""
                frmalquiler.Text1(0).SetFocus
            End If
            If menu = 4 Then
                frmfacctacte_alquiler.Text1(0).Text = ""
                frmfacctacte_alquiler.Text1(0).SetFocus
            End If
            If menu = 5 Then
                frmfacctacte_venta.Text1(0).Text = ""
                frmfacctacte_venta.Text1(0).SetFocus
            End If
            If menu = 6 Then
                frmnota_credito.Text1(0).Text = ""
                frmnota_credito.Text1(0).SetFocus
            End If
            If menu = 7 Then
                frmnota_debito.Text1(0).Text = ""
                frmnota_debito.Text1(0).SetFocus
            End If
        End If
    End If

End Sub

Private Sub Text1_Change()

 Text1.PasswordChar = "*"
 

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)

On Error Resume Next


    If KeyAscii = 13 Then
        KeyAscii = 0
        If Text1.Text = datvendedor.Recordset.Fields("clave") Then
             If menu = 1 Then
                frmnota_venta.Text1(1).SetFocus
             End If
             If menu = 2 Then
                frmpresupuesto.Text1(1).SetFocus
             End If
             If menu = 3 Then
                frmalquiler.Text1(1).SetFocus
             End If
             If menu = 4 Then
                frmfacctacte_alquiler.Text1(1).SetFocus
             End If
             If menu = 5 Then
                frmfacctacte_venta.Text1(1).SetFocus
             End If
             If menu = 6 Then
                frmnota_credito.Text1(1).SetFocus
             End If
             If menu = 7 Then
                frmnota_debito.Text1(1).SetFocus
             End If
             
        Else
            mensa = MsgBox("Clave Incorrecta!!", vbCritical, "Error")
            If menu = 1 Then
                frmnota_venta.Text1(0).Text = ""
                frmnota_venta.Text1(0).SetFocus
            End If
            If menu = 2 Then
                frmpresupuesto.Text1(0).Text = ""
                frmpresupuesto.Text1(0).SetFocus
            End If
            If menu = 3 Then
                frmalquiler.Text1(0).Text = ""
                frmalquiler.Text1(0).SetFocus
            End If
            If menu = 4 Then
                frmfacctacte_alquiler.Text1(0).Text = ""
                frmfacctacte_alquiler.Text1(0).SetFocus
            End If
            If menu = 5 Then
                frmfacctacte_venta.Text1(0).Text = ""
                frmfacctacte_venta.Text1(0).SetFocus
            End If
            If menu = 6 Then
                frmnota_credito.Text1(0).Text = ""
                frmnota_credito.Text1(0).SetFocus
            End If
            If menu = 7 Then
                frmnota_debito.Text1(0).Text = ""
                frmnota_debito.Text1(0).SetFocus
            End If

        End If
        Unload Me
    End If
            

End Sub
