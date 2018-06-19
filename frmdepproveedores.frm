VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmdepproveedores 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Depuracion Proveedores"
   ClientHeight    =   5565
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8610
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5565
   ScaleWidth      =   8610
   Begin MSDataListLib.DataList DataList2 
      Bindings        =   "frmdepproveedores.frx":0000
      Height          =   1425
      Left            =   2400
      TabIndex        =   4
      Top             =   3480
      Visible         =   0   'False
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   2514
      _Version        =   393216
      MatchEntry      =   -1  'True
      BackColor       =   -2147483626
      ListField       =   "razonsocial"
   End
   Begin MSDataListLib.DataList DataList1 
      Bindings        =   "frmdepproveedores.frx":0016
      Height          =   1425
      Left            =   3240
      TabIndex        =   0
      Top             =   720
      Visible         =   0   'False
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   2514
      _Version        =   393216
      MatchEntry      =   -1  'True
      BackColor       =   -2147483626
      ListField       =   "razonsocial"
   End
   Begin MSComctlLib.ProgressBar bar1 
      Height          =   255
      Left            =   3000
      TabIndex        =   6
      Top             =   2760
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Depurar"
      Height          =   375
      Left            =   6480
      TabIndex        =   5
      Top             =   2640
      Width           =   1695
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grilla1 
      Height          =   2055
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   3625
      _Version        =   393216
      FixedCols       =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSAdodcLib.Adodc datproveedor 
      Height          =   330
      Left            =   2040
      Top             =   5160
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
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
      DataSourceName  =   ""
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
   Begin MSAdodcLib.Adodc datprov 
      Height          =   330
      Left            =   840
      Top             =   5160
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
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
      DataSourceName  =   ""
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
   Begin MSAdodcLib.Adodc datproveedores1 
      Height          =   330
      Left            =   3360
      Top             =   5160
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
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
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "datproveedores1"
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
   Begin VB.TextBox Text1 
      BackColor       =   &H80000016&
      Height          =   375
      Left            =   3000
      TabIndex        =   2
      Top             =   120
      Width           =   5175
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grilla2 
      Height          =   2055
      Left            =   120
      TabIndex        =   3
      Top             =   3120
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   3625
      _Version        =   393216
      FixedCols       =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSAdodcLib.Adodc datlibrocompras 
      Height          =   330
      Left            =   7320
      Top             =   0
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
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
      DataSourceName  =   ""
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
   Begin MSAdodcLib.Adodc datordenes 
      Height          =   330
      Left            =   2640
      Top             =   0
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
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
      DataSourceName  =   ""
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
   Begin VB.Frame Frame1 
      Caption         =   "Depuración Libro Compras:"
      Height          =   1095
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   2775
   End
   Begin VB.Frame Frame2 
      Caption         =   "Depuracion Ordenes de Pago:"
      Height          =   1335
      Left            =   120
      TabIndex        =   8
      Top             =   2760
      Width           =   2775
   End
End
Attribute VB_Name = "frmdepproveedores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error Resume Next
bar1.Min = 0
For x = 1 To grilla1.Rows - 1

    grilla1.Col = 0
    grilla1.Row = x
    viejovalor = grilla1.Text
    datlibrocompras.RecordSource = "select librocompras.* from librocompras WHERE librocompras.empresa = " & login.empresaact & " and proveedor = '" & grilla1.Text & "'"
    datlibrocompras.Refresh
    If datlibrocompras.Recordset.EOF = True Then GoTo sale
    bar1.max = datlibrocompras.Recordset.RecordCount
    i = 0
    Do While Not datlibrocompras.Recordset.EOF
            grilla1.Col = 1
            i = i + 1
            bar1.Value = i
            If grilla1.Text = "" Then grilla1.Text = viejovalor
            datlibrocompras.Recordset.Fields("proveedor") = grilla1.Text
            datlibrocompras.Recordset.UpdateBatch adAffectCurrent
Rem            Debug.Print datlibrocompras.Recordset.Fields("proveedor"), grilla1.Text
            datlibrocompras.Recordset.MoveNext

    Loop
Next x
bar1.Value = 0
sale:
bar1.Min = 0
For x = 1 To grilla2.Rows - 1
    grilla2.Col = 0
    grilla2.Row = x
    viejovalor = grilla2.Text
    datordenes.RecordSource = "select ordendepagoabonan.* from ordendepagoabonan WHERE ordendepagoabonan.empresa = " & login.empresaact & " and nomproveedor = '" & grilla2.Text & "'"
    datordenes.Refresh
    If datordenes.Recordset.EOF = True Then Exit Sub
    bar1.max = datordenes.Recordset.RecordCount
    i = 0
    Do While Not datordenes.Recordset.EOF
            grilla2.Col = 1
            i = i + 1
            bar1.Value = i
            If grilla2.Text = "" Then grilla2.Text = viejovalor
            datordenes.Recordset.Fields("nomproveedor") = grilla2.Text
            datordenes.Recordset.UpdateBatch adAffectCurrent
            datordenes.Recordset.MoveNext

    Loop
Next x
bar1.Value = 0
Unload Me


End Sub

Private Sub DataList1_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        grilla1.Text = DataList1.Text
        DataList1.Visible = False
        grilla1.SetFocus
     End If

End Sub

Private Sub DataList1_LostFocus()

    DataList1.Visible = False
    

End Sub

Private Sub DataList2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        grilla2.Text = DataList2.Text
        DataList2.Visible = False
        grilla2.SetFocus
     End If
End Sub

Private Sub Form_Load()
Aplicar_skin Me

frmdepproveedores.Top = 0
frmdepproveedores.Left = 0

If login.provaltas = "N" Or login.provmodi = "N" Then
    Command1.Enabled = False
Else
    Command1.Enabled = True
End If

datprov.ConnectionString = login.conexiontotal
datproveedor.ConnectionString = login.conexiontotal
datproveedores1.ConnectionString = login.conexiontotal
datlibrocompras.ConnectionString = login.conexiontotal
datordenes.ConnectionString = login.conexiontotal
    
If login.empresaact1 > 0 Then
    empresareal = login.empresaact1
Else
    empresareal = login.empresaact
End If
    
    datprov.RecordSource = "select proveedores.* from proveedores where empresa = " & empresareal & " ORDER BY razonsocial"
    datprov.Refresh
    
    datproveedor.RecordSource = "select PROVEEDORESPERDIDOS.* from PROVEEDORESPERDIDOS where empresa = " & login.empresaact & ""
    datproveedor.Refresh
    
    datproveedores1.RecordSource = "select PROVEEDORESPERDIDOS1.* from PROVEEDORESPERDIDOS1 where empresa = " & login.empresaact & ""
    datproveedores1.Refresh
    
    grilla1.Col = 0
    grilla1.Row = 0
    grilla1.Text = "Razon Social Inexistente"
    grilla1.Col = 1
    grilla1.Row = 0
    grilla1.Text = "Razon Social Provable"
    grilla1.ColWidth(0) = 3500
    grilla1.ColWidth(1) = 3500
    grilla2.Col = 0
    grilla2.Row = 0
    grilla2.Text = "Razon Social Inexistente"
    grilla2.Col = 1
    grilla2.Row = 0
    grilla2.Text = "Razon Social Provable"
    grilla2.ColWidth(0) = 3500
    grilla2.ColWidth(1) = 3500

    grilla1.Rows = datproveedor.Recordset.RecordCount + 1
    grilla2.Rows = datproveedores1.Recordset.RecordCount + 1

    For x = 1 To grilla2.Rows - 1
         grilla2.Col = 1
         grilla2.Row = x
         grilla2.CellBackColor = QBColor(11)
    Next x

    For x = 1 To grilla1.Rows - 1
         grilla1.Col = 1
         grilla1.Row = x
         grilla1.CellBackColor = QBColor(11)
    Next x
        

If datproveedor.Recordset.EOF = True Then GoTo ordenes
i = 1
datproveedor.Recordset.MoveFirst
Do While Not datproveedor.Recordset.EOF
    grilla1.Col = 0
    grilla1.Row = i
    grilla1.Text = datproveedor.Recordset.Fields("proveedor")
    grilla1.Col = 1
    If IsNull(datproveedor.Recordset.Fields("provprobable")) = True Then
        grilla1.Text = ""
    Else
        grilla1.Text = datproveedor.Recordset.Fields("provprobable")
    End If
    datproveedor.Recordset.MoveNext
    i = i + 1
Loop

ordenes:
If datproveedores1.Recordset.EOF = True Then Exit Sub
i = 1
datproveedores1.Recordset.MoveFirst
Do While Not datproveedores1.Recordset.EOF
    grilla2.Col = 0
    grilla2.Row = i
    grilla2.Text = datproveedores1.Recordset.Fields("nomproveedor")
    grilla2.Col = 1
    If IsNull(datproveedores1.Recordset.Fields("provprobable")) = True Then
        grilla2.Text = ""
    Else
        grilla2.Text = datproveedores1.Recordset.Fields("provprobable")
    End If
    datproveedores1.Recordset.MoveNext
    i = i + 1
Loop

End Sub

Private Sub grilla1_Click()

       Text1.Text = grilla1.Text

End Sub

Private Sub grilla1_EnterCell()

           Text1.Text = grilla1.Text

End Sub

Private Sub grilla1_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 And grilla1.Col = 1 Then
        KeyAscii = 0
        DataList1.Visible = True
        DataList1.Left = grilla1.Left + 3500
        DataList1.Width = grilla1.ColWidth(1)
        DataList1.Top = grilla1.Top + grilla1.RowHeight(0) * grilla1.Row + grilla1.RowHeight(0)
        If grilla1.Text <> "" Then DataList1.Text = grilla1.Text
        DataList1.SetFocus
    End If

End Sub

Private Sub grilla1_KeyUp(KeyCode As Integer, Shift As Integer)

If KeyCode = 46 Then grilla1.Text = ""


End Sub


Private Sub grilla2_Click()

       Text1.Text = grilla2.Text

End Sub

Private Sub grilla2_EnterCell()

           Text1.Text = grilla2.Text

End Sub

Private Sub grilla2_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 And grilla2.Col = 1 Then
        KeyAscii = 0
        DataList2.Visible = True
        DataList2.Left = grilla2.Left + 3500
        DataList2.Width = grilla2.ColWidth(1)
        DataList2.Top = grilla2.Top + grilla2.RowHeight(0) * grilla2.Row + grilla2.RowHeight(0)
        If grilla2.Text <> "" Then DataList2.Text = grilla2.Text
        DataList2.SetFocus
    End If

End Sub

Private Sub grilla2_KeyUp(KeyCode As Integer, Shift As Integer)

If KeyCode = 46 Then grilla2.Text = ""


End Sub

Private Sub Label1_Click()

End Sub
