VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form lista_cuentas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Plan de Cuentas"
   ClientHeight    =   5057
   ClientLeft      =   39
   ClientTop       =   429
   ClientWidth     =   6318
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5057
   ScaleWidth      =   6318
   StartUpPosition =   2  'CenterScreen
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "lista_cuentas.frx":0000
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   4320
      Visible         =   0   'False
      Width           =   5175
      _ExtentX        =   9561
      _ExtentY        =   695
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.4717
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.4717
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSDataListLib.DataList DataList1 
      Bindings        =   "lista_cuentas.frx":0019
      DataSource      =   "datcuentas"
      Height          =   3960
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   6015
      _ExtentX        =   11095
      _ExtentY        =   7261
      _Version        =   393216
      MatchEntry      =   -1  'True
      ListField       =   "codigo"
      BoundColumn     =   "Cod Contable"
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   533
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   6318
      _ExtentX        =   11142
      _ExtentY        =   935
      ButtonWidth     =   551
      ButtonHeight    =   887
      Appearance      =   1
      Style           =   1
      _Version        =   393216
      Begin VB.CommandButton Command2 
         Caption         =   "Ordena por &Nombre"
         Height          =   375
         Left            =   3360
         TabIndex        =   3
         Top             =   120
         Width           =   1815
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Ordena por &Codigo"
         Height          =   375
         Left            =   1080
         TabIndex        =   2
         Top             =   120
         Width           =   1815
      End
   End
   Begin MSAdodcLib.Adodc datcuentas 
      Height          =   330
      Left            =   120
      Top             =   4680
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2228
      _ExtentY        =   599
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
         Size            =   7.47
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
End
Attribute VB_Name = "lista_cuentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public cuentacont As String
Dim Cuenta(99999) As Integer


Private Sub Command1_Click()
On Error GoTo fuera
datcuentas.RecordSource = "select listacuentas.* from listacuentas where empre = " & login.empresaact & " and imp = 'S' and inicioper = '" & login.iper & "'   order by idcuenta"
datcuentas.Refresh
DataList1.ListField = "codigo"

If datcuentas.Recordset.EOF = True Then
    MsgBox "Plan de Cuentas Vacio", vbCritical, "Error"
    Exit Sub
End If

DataList1.SetFocus
 
SendKeys "{down}", False
Exit Sub
fuera:
SendKeys "{tab}", False
SendKeys "{down}", False

End Sub

Private Sub Command2_Click()

datcuentas.RecordSource = "select listacuentas.* from listacuentas where empre = " & login.empresaact & " and imp = 'S' and inicioper = '" & login.iper & "'   order by nombre"
datcuentas.Refresh
DataList1.ListField = "nombre"

If datcuentas.Recordset.EOF = True Then
    MsgBox "Plan de Cuentas Vacio", vbCritical, "Error"
    Exit Sub
End If


DataList1.SetFocus

SendKeys "{down}", False


End Sub

Private Sub DataList1_Click()

    DataGrid1.Bookmark = DataList1.SelectedItem

End Sub

Private Sub DataList1_DblClick()
    
        If ventana.menu = 2 Then
            cuentacont = DataList1.BoundText
            frmchequescancelados.SetFocus
            SendKeys "{ENTER}", False
        End If
        If ventana.menu = 3 Then
            cuentacont = DataList1.BoundText
            frmCuentas.SetFocus
            SendKeys "{ENTER}", False
        End If
        If ventana.menu = 4 Then
            cuentacont = DataList1.BoundText
            frmclientes.SetFocus
        End If
        If ventana.menu = 5 Then
            cuentacont = DataList1.BoundText
            frmproveedores.SetFocus
        End If
        
        If ventana.menu = 6 Then
            cuentacont = DataGrid1.Columns(2).Text
            impmayoranalitico.SetFocus
            SendKeys "{ENTER}", False
        End If
        If ventana.menu = 7 Then
            cuentacont = DataGrid1.Columns(2).Text
            impsumasysaldos.SetFocus
            SendKeys "{ENTER}", False
        End If
        
        If ventana.menu = 8 Then
            cuentacont = DataList1.BoundText
            asienmodelo.SetFocus
        End If
        
        
        
        Unload Me
End Sub

Private Sub DataList1_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        If ventana.menu = 2 Then
            cuentacont = DataList1.BoundText
            frmchequescancelados.SetFocus
            SendKeys "{ENTER}", False
        End If
        If ventana.menu = 3 Then
            cuentacont = DataList1.BoundText
            frmCuentas.SetFocus
            SendKeys "{ENTER}", False
        End If
        If ventana.menu = 4 Then
            cuentacont = DataList1.BoundText
            frmclientes.SetFocus
        End If
        If ventana.menu = 5 Then
            cuentacont = DataList1.BoundText
            frmproveedores.SetFocus
        End If
        
        If ventana.menu = 6 Then
            cuentacont = DataGrid1.Columns(2).Text
            impmayoranalitico.SetFocus
            SendKeys "{ENTER}", False
        End If
        If ventana.menu = 7 Then
            cuentacont = DataGrid1.Columns(2).Text
            impsumasysaldos.SetFocus
            SendKeys "{ENTER}", False
        End If
        
        If ventana.menu = 8 Then
            cuentacont = DataList1.BoundText
            asienmodelo.SetFocus
        End If
        
        Unload Me
    End If

End Sub

Private Sub Form_Load()
Aplicar_skin Me

MiFuncionDeAjuste Me, True

datcuentas.ConnectionString = login.conexiontotal

Call Command1_Click

 
End Sub

Private Sub List1_KeyPress(KeyAscii As Integer)


    

End Sub
