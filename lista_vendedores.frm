VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form lista_vendedores 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de Vendedores"
   ClientHeight    =   4710
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6315
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4710
   ScaleWidth      =   6315
   StartUpPosition =   2  'CenterScreen
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "lista_vendedores.frx":0000
      Height          =   3855
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   6800
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
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
            LCID            =   11274
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
            LCID            =   11274
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
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   540
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6315
      _ExtentX        =   11139
      _ExtentY        =   953
      ButtonWidth     =   609
      ButtonHeight    =   900
      Appearance      =   1
      Style           =   1
      _Version        =   393216
      Begin VB.TextBox Text1 
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
         Left            =   1080
         TabIndex        =   3
         Top             =   120
         Width           =   5055
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Buscar:"
         Height          =   375
         Left            =   120
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   120
         Width           =   855
      End
   End
   Begin MSAdodcLib.Adodc datvendedor 
      Height          =   330
      Left            =   120
      Top             =   4440
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   582
      ConnectMode     =   1
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   1
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
Attribute VB_Name = "lista_vendedores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public cuentacont As String
Dim cuenta(99999) As Integer



Private Sub DataGrid1_DblClick()
On Error Resume Next


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
On Error Resume Next
    If KeyAscii = 13 Then
        KeyAscii = 0
        If datvendedor.Recordset.RecordCount = 0 Then
            Unload Me
            Exit Sub
        End If
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

DataGrid1.SetFocus

End Sub

Private Sub Form_Load()

If menu = 2 Then
'    Aplicar_skin2 Me
    Aplicar_skin Me
Else
    Aplicar_skin Me
End If

MiFuncionDeAjuste Me, True

datvendedor.ConnectionString = login.conexiontotal

datvendedor.RecordSource = query
datvendedor.Refresh

DataGrid1.Columns(0).Visible = False
DataGrid1.Columns(3).Visible = False
DataGrid1.Columns(1).Width = 1000
DataGrid1.Columns(2).Width = 3000


 
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)

On Error Resume Next


    If KeyAscii = 13 Then
        KeyAscii = 0
        If Text1.Text <> "" Then
            xbusqueda = "%" + Text1.Text + "%"
            xquery1 = "SELECT    V_VENDEDOR_.ID, V_VENDEDOR_.CODIGO, V_PERSONA_.NOMBRE, V_VENDEDOR_.CODIGO  + ' ' + V_PERSONA_.NOMBRE as cancatena " & _
                      "FROM V_VENDEDOR_ INNER JOIN V_PERSONA_ ON V_VENDEDOR_.ENTEASOCIADO_ID = V_PERSONA_.ID " & _
                      "WHERE (V_VENDEDOR_.ACTIVESTATUS = 0)  and V_VENDEDOR_.CODIGO  + ' ' + V_PERSONA_.NOMBRE like '" & xbusqueda & "' order by V_PERSONA_.NOMBRE"
                      
            datvendedor.RecordSource = xquery1
            datvendedor.Refresh
            DataGrid1.Columns(0).Visible = False
            DataGrid1.Columns(3).Visible = False
            DataGrid1.Columns(1).Width = 1000
            DataGrid1.Columns(2).Width = 3000

        End If
        DataGrid1.SetFocus
        
        
    End If

End Sub
