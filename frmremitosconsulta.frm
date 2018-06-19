VERSION 5.00
Object = "{F6125AB1-8AB1-11CE-A77F-08002B2F4E98}#2.0#0"; "MSRDC20.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{D18BBD1F-82BB-4385-BED3-E9D31A3E361E}#1.0#0"; "KewlButtonz.ocx"
Begin VB.Form frmremitosconsulta 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta de Remitos Emitidos"
   ClientHeight    =   7440
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15195
   Icon            =   "frmremitosconsulta.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7440
   ScaleWidth      =   15195
   Begin MSAdodcLib.Adodc datsaldar2 
      Height          =   330
      Left            =   9720
      Top             =   1680
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
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
      Caption         =   ""
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
   Begin MSAdodcLib.Adodc datsaldar 
      Height          =   330
      Left            =   9600
      Top             =   1320
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
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
      Caption         =   ""
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
   Begin VB.Frame Frame3 
      Caption         =   "Filtro"
      Height          =   855
      Left            =   8880
      TabIndex        =   9
      Top             =   120
      Width           =   1695
      Begin VB.OptionButton Option5 
         Caption         =   "Todos"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   480
         Width           =   1215
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Pend.de Fact"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   1455
      End
   End
   Begin MSAdodcLib.Adodc datitemremito 
      Height          =   330
      Left            =   11160
      Top             =   2040
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
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
      Caption         =   ""
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
   Begin VB.CommandButton calcula 
      Caption         =   "calcula"
      Height          =   255
      Left            =   11160
      TabIndex        =   6
      Top             =   2760
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc datremitos 
      Height          =   330
      Left            =   11160
      Top             =   1320
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
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
      Caption         =   ""
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
   Begin Crystal.CrystalReport CrystalReporte 
      Left            =   12000
      Top             =   2280
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTitle     =   "Orden de Pago"
      PrintFileLinesPerPage=   60
   End
   Begin MSRDC.MSRDC reporte 
      Height          =   375
      Left            =   11160
      Top             =   1680
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   661
      _Version        =   393216
      Options         =   0
      CursorDriver    =   0
      BOFAction       =   0
      EOFAction       =   0
      RecordsetType   =   1
      LockType        =   3
      QueryType       =   0
      Prompt          =   3
      Appearance      =   1
      QueryTimeout    =   30
      RowsetSize      =   100
      LoginTimeout    =   15
      KeysetSize      =   0
      MaxRows         =   0
      ErrorThreshold  =   -1
      BatchSize       =   15
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   -1  'True
      ReadOnly        =   0   'False
      Appearance      =   -1  'True
      DataSourceName  =   "contable"
      RecordSource    =   ""
      UserName        =   "sa"
      Password        =   ""
      Connect         =   ""
      LogMessages     =   ""
      Caption         =   "MSRDC1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmremitosconsulta.frx":0442
      Height          =   2775
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   14895
      _ExtentX        =   26273
      _ExtentY        =   4895
      _Version        =   393216
      AllowUpdate     =   -1  'True
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
            Format          =   "0"
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
   Begin VB.Frame Frame1 
      Caption         =   "Busqueda"
      Height          =   855
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   8775
      Begin VB.CommandButton Command1 
         Caption         =   "N.V. Nro:"
         Enabled         =   0   'False
         Height          =   375
         Left            =   5880
         TabIndex        =   17
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   6840
         TabIndex        =   2
         Top             =   240
         Width           =   1695
      End
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
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   5535
      End
   End
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   10560
      TabIndex        =   7
      Top             =   120
      Width           =   4455
      Begin VB.OptionButton Option1 
         Caption         =   "Original"
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   120
         Width           =   1335
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Duplicado"
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   360
         Width           =   1335
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Triplicado"
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   600
         Width           =   1335
      End
      Begin KewlButtonz.KewlButtons Command4 
         Height          =   495
         Left            =   1560
         TabIndex        =   3
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   873
         BTYPE           =   14
         TX              =   "&Previsualizar"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   -2147483645
         BCOLO           =   -2147483645
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmremitosconsulta.frx":045B
         PICN            =   "frmremitosconsulta.frx":0477
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin KewlButtonz.KewlButtons salir 
         Height          =   495
         Left            =   3120
         TabIndex        =   4
         Top             =   240
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   873
         BTYPE           =   14
         TX              =   "&Salir"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   -2147483645
         BCOLO           =   -2147483645
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmremitosconsulta.frx":3869
         PICN            =   "frmremitosconsulta.frx":3885
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid2 
      Bindings        =   "frmremitosconsulta.frx":43CF
      Height          =   2775
      Left            =   120
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   3960
      Width           =   14895
      _ExtentX        =   26273
      _ExtentY        =   4895
      _Version        =   393216
      AllowUpdate     =   0   'False
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
   Begin KewlButtonz.KewlButtons facturar 
      Height          =   495
      Left            =   6240
      TabIndex        =   12
      Top             =   6840
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
      BTYPE           =   14
      TX              =   "&Facturar"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   -2147483645
      BCOLO           =   -2147483645
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmremitosconsulta.frx":43EB
      PICN            =   "frmremitosconsulta.frx":4407
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin KewlButtonz.KewlButtons saldaremito 
      Height          =   495
      Left            =   9120
      TabIndex        =   13
      Top             =   6840
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
      BTYPE           =   14
      TX              =   "Saldar Remitos"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   -2147483645
      BCOLO           =   -2147483645
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmremitosconsulta.frx":49A1
      PICN            =   "frmremitosconsulta.frx":49BD
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
End
Attribute VB_Name = "frmremitosconsulta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim importeapagar As Double
Dim totalab As Currency
Dim totalinst(50) As Currency
Dim detalleint(50) As String
Dim totalconc(50) As Currency
Dim nrocompro(50) As String
Dim cuentaint(50) As Integer
Dim nomprov(50) As String
Dim saldoactual As Currency
Dim cuenta As Integer
Dim codprove As Integer
Dim idlibrogrid(50) As Integer
Dim saldolibro(50) As Currency
Public numorden As String
Dim xcon As Integer




Private Sub Combo1_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        List1.Clear
        datbuscaorden.RecordSource = "select libroventas.* from libroventas WHERE empresa = " & login.empresaact & " and tipocompr = '" & Combo1.Text & "' order by numcompr"
        datbuscaorden.Refresh
        datbuscaorden.Recordset.MoveFirst
        Do While Not datbuscaorden.Recordset.EOF
            List1.AddItem (datbuscaorden.Recordset.Fields("numcompr"))
            datbuscaorden.Recordset.MoveNext
        Loop
        DataCombo1.Text = ""
        DataCombo1.SetFocus
    End If

End Sub

Private Sub calcula_Click()
On Error Resume Next
Dim varBmk As Variant


    
    xquery = "select id_remito as id, referenciaproducto as Codigo, nombre_producto as Descripcion, cantidadoriginal as Cant_Orig, cantidadremitida as Cant_Remitida, " & _
             "cantfac as Cant_Facturada, pendfacturar as PendFacturar,unidaddemedida as Um, null as numeradorinterno, item as iditem " & _
             "from v_ezi_pos_traza_remito_factura as T " & _
             "where      (t.id_remito= " & DataGrid1.Columns(0).Text & ") " & _
             "ORDER BY iditem"
             
             
If datremitos.Recordset.EOF = False Then
        datitemremito.RecordSource = xquery
        datitemremito.Refresh
        DataGrid2.Visible = True
Else
        DataGrid2.Visible = False
End If


            For X = 2 To 14
                DataGrid1.Columns(X).Locked = True
            Next X
            If Text2.Text = "" Then
                DataGrid1.Columns(1).Locked = True
            Else
                DataGrid1.Columns(1).Locked = False
            End If
            DataGrid1.Columns(1).Width = 300
            DataGrid1.Columns(4).Width = 1000
            DataGrid1.Columns(5).Width = 1000
            DataGrid1.Columns(6).Width = 3500
            
            DataGrid1.Columns(0).Visible = False
            DataGrid1.Columns(9).Visible = False
            DataGrid1.Columns(10).Visible = False
            DataGrid1.Columns(12).Visible = False
            DataGrid1.Columns(13).Visible = False
            DataGrid1.Columns(14).Visible = False
            DataGrid1.Columns(15).Visible = False
            
            DataGrid2.Columns(0).Visible = False
            DataGrid2.Columns(2).Width = 5500
            DataGrid2.Columns(3).Alignment = dbgCenter
            DataGrid2.Columns(4).Alignment = dbgCenter
            DataGrid2.Columns(5).Alignment = dbgCenter


If Option4.Value = True Then
    facturar.Enabled = True
Else
    facturar.Enabled = True
    
End If

End Sub

Private Sub Command4_Click()
On Error Resume Next
Dim crApp As New CRAXDRT.Application
Dim crReport As New CRAXDRT.Report

Dim tabla As String
Dim ruta As String


reporte.SQL = "SELECT v_ezi_pos_remito.id, v_ezi_pos_remito.NUMERODOCUMENTO, v_ezi_pos_remito.FECHAEMISION, v_ezi_pos_remito.cod_cliente, v_ezi_pos_remito.cliente, v_ezi_pos_remito.CALLE, v_ezi_pos_remito.CODPOS, v_ezi_pos_remito.provincia, v_ezi_pos_remito.detalle, v_ezi_pos_remito.tipopago, v_ezi_pos_remito.referenciaproducto, v_ezi_pos_remito.nombre_producto, v_ezi_pos_remito.cantidadremitida, v_ezi_pos_remito.nota, v_ezi_pos_remito.condiva, v_ezi_pos_remito.ciudad, v_ezi_pos_remito.TIPOVENTA, v_ezi_pos_remito.SIMBOLO, v_ezi_pos_remito.iditem FROM MMOSSE.dbo.v_ezi_pos_remito v_ezi_pos_remito where v_ezi_pos_remito.id = " & DataGrid1.Columns(0).Text & " order by v_ezi_pos_remito.iditem"
tabla = reporte.SQL

With CrystalReporte
    .PrinterCollation = crptCollated
    .ReportFileName = App.Path & "\RemitoVta.rpt"
    If Option1.Value = True Then xtipo = "ORIGINAL"
    If Option2.Value = True Then xtipo = "DUPLICADO"
    If Option3.Value = True Then xtipo = "TRIPLICADO"
    .Formulas(0) = "copia=""" & xtipo & """"
    .WindowTitle = "Remito " + DataGrid1.Columns(1).Text + " " + xtipo
    .Connect = login.conexionreporte
    .DiscardSavedData = True
    .RetrieveDataFiles
    .ReportSource = 0
    .SQLQuery = tabla
 Rem   .Destination = crptToWindow
    .Destination = crptToFile
    .PrintFileType = crptCrystal
'    .WindowState = crptMaximized
    .PrintFileName = App.Path & "\remconsulta.rpt"
    .Action = 1

End With

impresos.Show
Set crReport = crApp.OpenReport(App.Path & "\remconsulta.rpt", 1)
impresos.Caption = "Remito " + DataGrid1.Columns(1).Text + " " + xtipo
impresos.CRViewer1.ReportSource = crReport
impresos.CRViewer1.ViewReport


End Sub



Private Sub DataCombo1_KeyPress(KeyAscii As Integer)
On Error GoTo fueraderango
    If KeyAscii = 13 Then
        KeyAscii = 0
        List1.ListIndex = DataCombo1.SelectedItem - 1
        Call Command4_Click
    End If
fueraderango:
End Sub

Private Sub DataGrid1_Click()
    
    Call calcula_Click
    
End Sub

Private Sub DataGrid1_GotFocus()
    
    xcon = 0

End Sub

Private Sub DataGrid1_KeyDown(KeyCode As Integer, Shift As Integer)

    Call calcula_Click

End Sub

Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
On Error Resume Next

    If KeyAscii = 13 Then
        datremitos.Recordset.MoveNext
'        Call Command4_Click
    End If
    
    
End Sub

Private Sub DataGrid1_KeyUp(KeyCode As Integer, Shift As Integer)

    Call calcula_Click

End Sub

Private Sub facturar_Click()
On Error Resume Next

xquery = "SELECT   distinct  N.id, N.numeradorinterno, N.clienteid, N.cliente, N.vendedorid, N.vendedor, N.detalle, N.nota, N.mediodetransporteid, N.monedaid, N.cotizacion, N.listadeprecioid, " & _
         "N.tipodepagoid, N.tipodefacturacionid, N.tipodeentregaid, N.fechadeentrega, N.recargo, N.tiporecargo, N.bonificacion, N.tipobonificacion, N.importeglobal, " & _
         "N.numerodefactura, N.tipodefactura, N.numerador, N.domicilioid, N.domiciliodeentregaid, N.puntodeventa, N.descuento, N.subtotalsiniva, N.totaliva, N.generada, " & _
         "N.presupuestobase, N.importado, N.comprobanteorigen, N.totaltr, N.percepiibb, N.formapago, N.target, N.nombrepc, N.nrofactinicial, N.nrofactfinal, N.responsabilidad, " & _
         "N.nrofactoriginante, N.factoriginanteid, N.estado, N.transferido, N.calipsoid, N.osmpmutualid, N.recetaid, N.medicoid, N.adicionalid, N.reconocido, N.facturar, N.senia, " & _
         "N.nroorden, N.fechaorden, N.estadoimpresion, N.imprimir, N.perceptem, N.perceppyp, R.id AS idremito, R.NUMERODOCUMENTO, N.claveprimaria, " & _
         "N.fechadelcomprobante , N.sucursal, N.obraid FROM  ud_ezi_puntodeventa_encabezado AS N LEFT OUTER JOIN v_ezi_pos_remito AS R ON N.id = R.trazabilidad_id " & _
         "WHERE (N.numeradorinterno = 'Nota de Venta') and  r.id ='" & DataGrid1.Columns(0).Text & "' "

X = 0
C = 1
xcuenta = datremitos.Recordset.RecordCount
xclaveprimariaremito = DataGrid1.Columns(0).Text
datremitos.Recordset.MoveFirst
Do While xcuenta >= C
  If UCase(DataGrid1.Columns(1).Text) = "S" Then
    
    
    xquery1 = "SELECT   distinct  N.id, N.numeradorinterno, N.clienteid, N.cliente, N.vendedorid, N.vendedor, N.detalle, N.nota, N.mediodetransporteid, N.monedaid, N.cotizacion, N.listadeprecioid, " & _
              "N.tipodepagoid, N.tipodefacturacionid, N.tipodeentregaid, N.fechadeentrega, N.recargo, N.tiporecargo, N.bonificacion, N.tipobonificacion, N.importeglobal, " & _
              "N.numerodefactura, N.tipodefactura, N.numerador, N.domicilioid, N.domiciliodeentregaid, N.puntodeventa, N.descuento, N.subtotalsiniva, N.totaliva, N.generada, " & _
              "N.presupuestobase, N.importado, N.comprobanteorigen, N.totaltr, N.percepiibb, N.formapago, N.target, N.nombrepc, N.nrofactinicial, N.nrofactfinal, N.responsabilidad, " & _
              "N.nrofactoriginante, N.factoriginanteid, N.estado, N.transferido, N.calipsoid, N.osmpmutualid, N.recetaid, N.medicoid, N.adicionalid, N.reconocido, N.facturar, N.senia, " & _
              "N.nroorden, N.fechaorden, N.estadoimpresion, N.imprimir, N.perceptem, N.perceppyp, R.id AS idremito, R.NUMERODOCUMENTO, N.claveprimaria, " & _
              "N.fechadelcomprobante , N.sucursal, N.obraid FROM  ud_ezi_puntodeventa_encabezado AS N LEFT OUTER JOIN v_ezi_pos_remito AS R ON N.id = R.trazabilidad_id " & _
              "WHERE (N.numeradorinterno = 'Nota de Venta') and  r.id ='" & DataGrid1.Columns(0).Text & "' "
   
    If X >= 1 Then
        xquery2 = " Union all "
        xquery = xquery1 + xquery2 + xquery
        xquery1 = ""
    Else
        xquery = xquery1
    End If
    
    X = X + 1
  End If
    datremitos.Recordset.MoveNext
    C = C + 1
 
Loop


    query = xquery
    If query = "" Then
        MsgBox "Seleccione nuevamente el remito a Facturar", vbInformation, "Atención"
        Exit Sub
    End If
    
    remdev = 0
    frmfacctacte_venta.Show
    If X = 0 Then
    '    frmfacctacte_venta.Text18.Text = DataGrid1.Columns(12).Text
'        frmfacctacte_venta.Text17.Text = DataGrid1.Columns(2).Text
    End If
    frmfacctacte_venta.Text17.SetFocus
    SendKeys "{ENTER}", False
    
    

End Sub

Private Sub Form_Activate()

    DataGrid1.SetFocus
    If menu = 1 Then
        If Text1.Text = "" Then Text1.Text = " "
        Text1.SetFocus
        SendKeys "{ENTER}", False
    End If

    If menu = 10 Then
        Text1.Text = lista_trazabilidad.xtrazabilidad
        Text1.SetFocus
        SendKeys "{ENTER}", False
    End If
    


End Sub

Private Sub Form_Load()
On Error Resume Next
Aplicar_skin Me

yventana = Inicio.Height / 2 - 1000
xventana = Inicio.Width / 2

frmremitosconsulta.Top = yventana - frmremitosconsulta.Height / 2
frmremitosconsulta.Left = xventana - frmremitosconsulta.Width / 2


datremitos.ConnectionString = login.conexiontotal
datitemremito.ConnectionString = login.conexiontotal
datsaldar.ConnectionString = login.conexiontotal
datsaldar2.ConnectionString = login.conexiontotal

xcon = 1
facturar.Visible = False

If UCase(login.usuarioactivo) = "ADMIN" Then
    saldaremito.Visible = False
Else
    saldaremito.Visible = False
End If


xquery1 = "SELECT     id, Sel, NroRemito, Fecha, presupuestobase as NV, CodCliente, Cliente, CUIT, MAX(NroFactura) AS NroFactura, Tipopago, TipodeVenta,  Vendedor, MAX(cantremitida) " & _
          "AS cantremitida, SUM(cantfacturada) AS cantfacturada, MAX(concatenado) AS concatenado, saldado " & _
          "FROM         (SELECT     R.id, R.Sel, R.NUMERODOCUMENTO AS NroRemito, R.FECHAEMISION AS Fecha, R.cod_cliente AS CodCliente, R.cliente AS Cliente, R.CUIT,  " & _
          "R.vendedor AS Vendedor, R.tipopago AS Tipopago, R.TIPOVENTA AS TipodeVenta, R.presupuestobase, E.tipodefactura + ' ' + E.puntodeventa + RIGHT('0000000' + E.numerodefactura, 8) AS NroFactura, SUM(R.cantidadremitida) AS cantremitida,  " & _
          "SUM(ISNULL(IFAC.cantidadproducto, 0)) + ISNULL(FA.cantidadfacturada2, 0) AS cantfacturada, R.NUMERODOCUMENTO + 'r  ' + R.cliente + ' ' + R.CUIT + ' ' + ISNULL(E.tipodefactura + ' ' + E.puntodeventa + RIGHT('0000000' + E.numerodefactura, 8), '')  " & _
          "+ 'f' AS concatenado, R.saldado, E.numeradorinterno, FA.cantidadfacturada2 FROM (SELECT     F.trazabilidad_id, SUM(ITF.cantidadproducto) AS cantidadfacturada2 " & _
          "FROM ud_ezi_puntodeventa_encabezado AS F WITH (readpast) INNER JOIN ud_ezi_puntodeventa_detalle_factm AS ITF WITH (readpast) ON F.id = ITF.claveprimaria " & _
          "WHERE      (F.numeradorinterno = 'Factura de Venta') GROUP BY F.trazabilidad_id  HAVING      (NOT (F.trazabilidad_id IS NULL))) AS FA RIGHT OUTER JOIN " & _
          "v_ezi_pos_remito AS R ON FA.trazabilidad_id = R.id LEFT OUTER JOIN ud_ezi_puntodeventa_detalle_factm AS IFAC WITH (readpast) RIGHT OUTER JOIN " & _
          "ud_ezi_puntodeventa_encabezado AS E WITH (readpast) ON IFAC.claveprimaria = E.id ON R.trazabilidad_id = E.trazabilidad_id GROUP BY R.id, R.Sel, R.NUMERODOCUMENTO, R.FECHAEMISION, R.cod_cliente, R.cliente, R.CUIT, R.vendedor, R.tipopago, R.TIPOVENTA,  " & _
          "E.tipodefactura + ' ' + E.puntodeventa + RIGHT('0000000' + E.numerodefactura, 8), R.presupuestobase, R.alquiler, R.saldado, E.numeradorinterno, FA.cantidadfacturada2,R.estado " & _
          "HAVING    (E.numeradorinterno LIKE '%Factura%' OR " & _
          "E.numeradorinterno LIKE '%Nota%') AND (R.estado = 'Remitido')) AS RC GROUP BY id, Sel, NroRemito, Fecha, CodCliente, Cliente, CUIT, Vendedor, Tipopago, TipodeVenta, presupuestobase, saldado " & _
          "HAVING      (saldado IS NULL) AND (SUM(cantfacturada) < sum(cantremitida)) ORDER BY id DESC"

datremitos.RecordSource = xquery1
datremitos.Refresh

If datremitos.Recordset.EOF = False Then
        datremitos.Recordset.MoveFirst
        datitemremito.RecordSource = "select id_remito as id, referenciaproducto as Codigo, nombre_producto as Descripcion, cantidadoriginal as Cant_Orig, cantidadremitida as Cant_Remitida, " & _
                                     "cantfac as Cant_Facturada, pendfacturar as PendFacturar,unidaddemedida as Um, null as numeradorinterno, item as iditem " & _
                                     "from v_ezi_pos_traza_remito_factura as T " & _
                                     "where      (t.id_remito= " & DataGrid1.Columns(0).Text & ") " & _
                                     "ORDER BY iditem"
        datitemremito.Refresh
        DataGrid2.Visible = True
Else
        DataGrid2.Visible = False
End If

            For X = 2 To 14
                DataGrid1.Columns(X).Locked = True
            Next X
            DataGrid1.Columns(1).Width = 300
            DataGrid1.Columns(4).Width = 1000
            DataGrid1.Columns(5).Width = 1000
            DataGrid1.Columns(6).Width = 3500
            
            DataGrid1.Columns(0).Visible = False
            DataGrid1.Columns(9).Visible = False
            DataGrid1.Columns(10).Visible = False
            DataGrid1.Columns(12).Visible = False
            DataGrid1.Columns(13).Visible = False
            DataGrid1.Columns(14).Visible = False
            DataGrid1.Columns(15).Visible = False
            
            
            DataGrid2.Columns(0).Visible = False
            DataGrid2.Columns(2).Width = 5500
            DataGrid2.Columns(3).Alignment = dbgCenter
            DataGrid2.Columns(4).Alignment = dbgCenter
            DataGrid2.Columns(5).Alignment = dbgCenter

Option1.Value = True
Option5.Value = True

If menu = 1 Then
    'Option5.Value = False
    'Option5.Enabled = False
    Option5.Enabled = True
    Option4.Value = True
    facturar.Visible = True
End If


End Sub

Private Sub Option4_Click()

    If xcon = 0 Then
        If Text1.Text = "" Then Text1.Text = " "
        Text1.SetFocus
        SendKeys "{ENTER}", False
    End If


End Sub

Private Sub Option5_Click()

    If xcon = 0 Then
        If Text1.Text = "" Then Text1.Text = " "
        Text1.SetFocus
        SendKeys "{ENTER}", False
    End If

End Sub

Private Sub saldaremito_Click()
On Error Resume Next

mensa = MsgBox("Esta por saldar los pendientes de los remitos seleccionados, Esta Seguro ?", vbYesNo, "Atención")
If mensa = vbYes Then
    xquery = "SELECT   distinct  N.id, N.numeradorinterno, N.clienteid, N.cliente, N.vendedorid, N.vendedor, N.detalle, N.nota, N.mediodetransporteid, N.monedaid, N.cotizacion, N.listadeprecioid, " & _
         "N.tipodepagoid, N.tipodefacturacionid, N.tipodeentregaid, N.fechadeentrega, N.recargo, N.tiporecargo, N.bonificacion, N.tipobonificacion, N.importeglobal, " & _
         "N.numerodefactura, N.tipodefactura, N.numerador, N.domicilioid, N.domiciliodeentregaid, N.puntodeventa, N.descuento, N.subtotalsiniva, N.totaliva, N.generada, " & _
         "N.presupuestobase, N.importado, N.comprobanteorigen, N.totaltr, N.percepiibb, N.formapago, N.target, N.nombrepc, N.nrofactinicial, N.nrofactfinal, N.responsabilidad, " & _
         "N.nrofactoriginante, N.factoriginanteid, N.estado, N.transferido, N.calipsoid, N.osmpmutualid, N.recetaid, N.medicoid, N.adicionalid, N.reconocido, N.facturar, N.senia, " & _
         "N.nroorden, N.fechaorden, N.estadoimpresion, N.imprimir, N.perceptem, N.perceppyp, R.id AS idremito, R.NUMERODOCUMENTO, N.claveprimaria, " & _
         "N.fechadelcomprobante , N.sucursal, N.obraid FROM  ud_ezi_puntodeventa_encabezado AS N LEFT OUTER JOIN v_ezi_pos_remito AS R ON N.id = R.presupuestobase " & _
         "WHERE (N.numeradorinterno = 'Nota de Venta') and  N.id ='" & DataGrid1.Columns(10).Text & "' "

X = 0
C = 1
xcuenta = datremitos.Recordset.RecordCount
datremitos.Recordset.MoveFirst
Do While xcuenta >= C
  If UCase(DataGrid1.Columns(1).Text) = "S" Then
    
    
    xquery1 = "SELECT  distinct   N.id, N.numeradorinterno, N.clienteid, N.cliente, N.vendedorid, N.vendedor, N.detalle, N.nota, N.mediodetransporteid, N.monedaid, N.cotizacion, N.listadeprecioid, " & _
              "N.tipodepagoid, N.tipodefacturacionid, N.tipodeentregaid, N.fechadeentrega, N.recargo, N.tiporecargo, N.bonificacion, N.tipobonificacion, N.importeglobal, " & _
              "N.numerodefactura, N.tipodefactura, N.numerador, N.domicilioid, N.domiciliodeentregaid, N.puntodeventa, N.descuento, N.subtotalsiniva, N.totaliva, N.generada, " & _
              "N.presupuestobase, N.importado, N.comprobanteorigen, N.totaltr, N.percepiibb, N.formapago, N.target, N.nombrepc, N.nrofactinicial, N.nrofactfinal, N.responsabilidad, " & _
              "N.nrofactoriginante, N.factoriginanteid, N.estado, N.transferido, N.calipsoid, N.osmpmutualid, N.recetaid, N.medicoid, N.adicionalid, N.reconocido, N.facturar, N.senia, " & _
              "N.nroorden, N.fechaorden, N.estadoimpresion, N.imprimir, N.perceptem, N.perceppyp, R.id AS idremito, R.NUMERODOCUMENTO, N.claveprimaria,  " & _
              "N.fechadelcomprobante , N.sucursal, N.obraid FROM  ud_ezi_puntodeventa_encabezado AS N LEFT OUTER JOIN v_ezi_pos_remito AS R ON N.id = R.presupuestobase " & _
              "WHERE (N.numeradorinterno = 'Nota de Venta') and  N.id ='" & DataGrid1.Columns(10).Text & "' "
   
    If X >= 1 Then
        xquery2 = " Union all "
        xquery = xquery1 + xquery2 + xquery
        xquery1 = ""
    Else
        xquery = xquery1
    End If
    
    X = X + 1
  End If
    datremitos.Recordset.MoveNext
    C = C + 1
 
Loop
    
    datsaldar.RecordSource = xquery
    datsaldar.Refresh
    
    If datsaldar.Recordset.EOF = False Then
        Do While Not datsaldar.Recordset.EOF
            datsaldar2.RecordSource = "select claveprimaria, saldado from ud_ezi_puntodeventa_detalle_rem where claveprimaria = '" & datsaldar.Recordset.Fields("idremito") & "'"
            datsaldar2.Refresh
            If datsaldar2.Recordset.EOF = False Then
                datsaldar2.Recordset.Fields("saldado") = "1"
                datsaldar2.Recordset.UpdateBatch adAffectCurrent
            End If
            datsaldar.Recordset.MoveNext
        Loop
    End If
    
    
     Text1.Text = ""
     Text1.SetFocus
     SendKeys "{ENTER}", False

End If

    


End Sub

Private Sub salir_Click()

    Unload Me

End Sub


Private Sub Text1_KeyPress(KeyAscii As Integer)
On Error Resume Next


    If KeyAscii = 13 Then
        KeyAscii = 0
        If Text1.Text = "" Then Text1.Text = " "
        If Text1.Text <> "" Then
            xbusqueda = "%" + Text1.Text + "%"
            If Option4.Value = True Then
              If Text2.Text = "" Then
                xquery1 = "SELECT     id, Sel, NroRemito, Fecha, presupuestobase AS NV, CodCliente, Cliente, CUIT, case when LEFT( MIN(ISNULL(NroFactura, nfactura)),1) = 'R' then null else MIN(ISNULL(NroFactura, nfactura)) end  AS NroFactura, Tipopago, TipodeVenta, Vendedor, MAX(cantremitida) " & _
                          "AS cantremitida, SUM(cantfacturada) AS cantfacturada, MAX(concatenado) AS concatenado, saldado " & _
                          "FROM         (SELECT     R.id, R.Sel, R.NUMERODOCUMENTO AS NroRemito, R.FECHAEMISION AS Fecha, R.cod_cliente AS CodCliente, R.cliente AS Cliente, R.CUIT, " & _
                          "R.vendedor AS Vendedor, R.tipopago AS Tipopago, R.TIPOVENTA AS TipodeVenta, R.presupuestobase, " & _
                          "E.tipodefactura + ' ' + E.puntodeventa + RIGHT('0000000' + E.numerodefactura, 8) AS NroFactura, SUM(R.cantidadremitida) AS cantremitida, " & _
                          "SUM(ISNULL(IFAC.cantidadproducto, 0)) + ISNULL(FA.cantidadfacturada2, 0) AS cantfacturada, " & _
                          "R.NUMERODOCUMENTO + 'r  ' + R.cliente + ' ' + R.CUIT + ' ' + ISNULL(E.tipodefactura + ' ' + E.puntodeventa + RIGHT('0000000' + E.numerodefactura, 8), '') " & _
                          "+ 'f' AS concatenado, R.saldado, E.numeradorinterno, FA.cantidadfacturada2, FA.NroFactura as nfactura " & _
                          "FROM          (SELECT    case when ITF.idclaveprimariaremito = '' then F.trazabilidad_id else ITF.idclaveprimariaremito end as  trazabilidad_id, SUM(ITF.cantidadproducto) AS cantidadfacturada2, " & _
                          "F.tipodefactura + ' ' + F.puntodeventa + RIGHT('0000000' +F.numerodefactura, 8) AS NroFactura " & _
                          "FROM         ud_ezi_puntodeventa_encabezado AS F WITH (readpast) INNER JOIN " & _
                          "ud_ezi_puntodeventa_detalle_factm AS ITF WITH (readpast) ON F.id = ITF.claveprimaria " & _
                          "WHERE     (F.numeradorinterno = 'Factura de Venta') " & _
                          "GROUP BY F.trazabilidad_id, ITF.idclaveprimariaremito, tipodefactura, puntodeventa, numerodefactura " & _
                          "HAVING      (NOT (F.trazabilidad_id IS NULL))) AS FA RIGHT OUTER JOIN " & _
                          "v_ezi_pos_remito AS R ON FA.trazabilidad_id = R.id LEFT OUTER JOIN " & _
                          "ud_ezi_puntodeventa_detalle_factm AS IFAC WITH (readpast) RIGHT OUTER JOIN " & _
                          "ud_ezi_puntodeventa_encabezado AS E WITH (readpast) ON IFAC.claveprimaria = E.id ON R.trazabilidad_id = E.trazabilidad_id " & _
                          "GROUP BY R.id, R.Sel, R.NUMERODOCUMENTO, R.FECHAEMISION, R.cod_cliente, R.cliente, R.CUIT, R.vendedor, R.tipopago, R.TIPOVENTA, " & _
                          "E.tipodefactura + ' ' + E.puntodeventa + RIGHT('0000000' + E.numerodefactura, 8), R.presupuestobase, R.alquiler, R.saldado, E.numeradorinterno, " & _
                          "FA.cantidadfacturada2 , R.Estado, FA.NroFactura " & _
                          "HAVING      (R.alquiler = 'N') AND (R.NUMERODOCUMENTO + 'r  ' + R.cliente + ' ' + R.CUIT + ' ' + ISNULL(E.tipodefactura + ' ' + E.puntodeventa + RIGHT('0000000' + E.numerodefactura, 8), '') + 'f' LIKE '" & xbusqueda & "')  " & _
                          "AND (R.estado = 'Remitido')) AS RC " & _
                          "GROUP BY id, Sel, NroRemito, Fecha, CodCliente, Cliente, CUIT, Vendedor, Tipopago, TipodeVenta, presupuestobase, saldado " & _
                          "Having (saldado Is Null) And (Sum(cantfacturada) < Max(cantremitida)) ORDER BY id DESC"
              Else
                xquery1 = "SELECT     id, Sel, NroRemito, Fecha, presupuestobase AS NV, CodCliente, Cliente, CUIT, case when LEFT( MIN(ISNULL(NroFactura, nfactura)),1) = 'R' then null else MIN(ISNULL(NroFactura, nfactura)) end  AS NroFactura, Tipopago, TipodeVenta, Vendedor, MAX(cantremitida) " & _
                          "AS cantremitida, SUM(cantfacturada) AS cantfacturada, MAX(concatenado) AS concatenado, saldado " & _
                          "FROM         (SELECT     R.id, R.Sel, R.NUMERODOCUMENTO AS NroRemito, R.FECHAEMISION AS Fecha, R.cod_cliente AS CodCliente, R.cliente AS Cliente, R.CUIT, " & _
                          "R.vendedor AS Vendedor, R.tipopago AS Tipopago, R.TIPOVENTA AS TipodeVenta, R.presupuestobase, " & _
                          "E.tipodefactura + ' ' + E.puntodeventa + RIGHT('0000000' + E.numerodefactura, 8) AS NroFactura, SUM(R.cantidadremitida) AS cantremitida, " & _
                          "SUM(ISNULL(IFAC.cantidadproducto, 0)) + ISNULL(FA.cantidadfacturada2, 0) AS cantfacturada, " & _
                          "R.NUMERODOCUMENTO + 'r  ' + R.cliente + ' ' + R.CUIT + ' ' + ISNULL(E.tipodefactura + ' ' + E.puntodeventa + RIGHT('0000000' + E.numerodefactura, 8), '') " & _
                          "+ 'f' AS concatenado, R.saldado, E.numeradorinterno, FA.cantidadfacturada2, FA.NroFactura as nfactura " & _
                          "FROM          (SELECT    case when ITF.idclaveprimariaremito = '' then F.trazabilidad_id else ITF.idclaveprimariaremito end as  trazabilidad_id, SUM(ITF.cantidadproducto) AS cantidadfacturada2, " & _
                          "F.tipodefactura + ' ' + F.puntodeventa + RIGHT('0000000' +F.numerodefactura, 8) AS NroFactura " & _
                          "FROM         ud_ezi_puntodeventa_encabezado AS F WITH (readpast) INNER JOIN " & _
                          "ud_ezi_puntodeventa_detalle_factm AS ITF WITH (readpast) ON F.id = ITF.claveprimaria " & _
                          "WHERE     (F.numeradorinterno = 'Factura de Venta') " & _
                          "GROUP BY F.trazabilidad_id, ITF.idclaveprimariaremito, tipodefactura, puntodeventa, numerodefactura " & _
                          "HAVING      (NOT (F.trazabilidad_id IS NULL))) AS FA RIGHT OUTER JOIN " & _
                          "v_ezi_pos_remito AS R ON FA.trazabilidad_id = R.id LEFT OUTER JOIN " & _
                          "ud_ezi_puntodeventa_detalle_factm AS IFAC WITH (readpast) RIGHT OUTER JOIN " & _
                          "ud_ezi_puntodeventa_encabezado AS E WITH (readpast) ON IFAC.claveprimaria = E.id ON R.trazabilidad_id = E.trazabilidad_id " & _
                          "GROUP BY R.id, R.Sel, R.NUMERODOCUMENTO, R.FECHAEMISION, R.cod_cliente, R.cliente, R.CUIT, R.vendedor, R.tipopago, R.TIPOVENTA, " & _
                          "E.tipodefactura + ' ' + E.puntodeventa + RIGHT('0000000' + E.numerodefactura, 8), R.presupuestobase, R.alquiler, R.saldado, E.numeradorinterno, " & _
                          "FA.cantidadfacturada2 , R.Estado, FA.NroFactura " & _
                          "HAVING      (R.alquiler = 'N') AND (R.NUMERODOCUMENTO + 'r  ' + R.cliente + ' ' + R.CUIT + ' ' + ISNULL(E.tipodefactura + ' ' + E.puntodeventa + RIGHT('0000000' + E.numerodefactura, 8), '') + 'f' LIKE '" & xbusqueda & "') " & _
                          "AND (R.estado = 'Remitido')) AS RC " & _
                          "GROUP BY id, Sel, NroRemito, Fecha, CodCliente, Cliente, CUIT, Vendedor, Tipopago, TipodeVenta, presupuestobase, saldado " & _
                          "Having (saldado Is Null) And (Sum(cantfacturada) < Max(cantremitida)) and presupuestobase = '" & Text2.Text & "' ORDER BY id DESC"
                End If
            Else
             If Text2.Text = "" Then
                xquery1 = "SELECT     id, Sel, NroRemito, Fecha, presupuestobase AS NV, CodCliente, Cliente, CUIT, case when LEFT( MIN(ISNULL(NroFactura, nfactura)),1) = 'R' then null else MIN(ISNULL(NroFactura, nfactura)) end  AS NroFactura, Tipopago, TipodeVenta, Vendedor, MAX(cantremitida) " & _
                          "AS cantremitida, SUM(cantfacturada) AS cantfacturada, MAX(concatenado) AS concatenado, saldado " & _
                          "FROM         (SELECT     R.id, R.Sel, R.NUMERODOCUMENTO AS NroRemito, R.FECHAEMISION AS Fecha, R.cod_cliente AS CodCliente, R.cliente AS Cliente, R.CUIT, " & _
                          "R.vendedor AS Vendedor, R.tipopago AS Tipopago, R.TIPOVENTA AS TipodeVenta, R.presupuestobase, " & _
                          "E.tipodefactura + ' ' + E.puntodeventa + RIGHT('0000000' + E.numerodefactura, 8) AS NroFactura, SUM(R.cantidadremitida) AS cantremitida, " & _
                          "SUM(ISNULL(IFAC.cantidadproducto, 0)) + ISNULL(FA.cantidadfacturada2, 0) AS cantfacturada, " & _
                          "R.NUMERODOCUMENTO + 'r  ' + R.cliente + ' ' + R.CUIT + ' ' + ISNULL(E.tipodefactura + ' ' + E.puntodeventa + RIGHT('0000000' + E.numerodefactura, 8), '') " & _
                          "+ 'f' AS concatenado, R.saldado, E.numeradorinterno, FA.cantidadfacturada2, FA.NroFactura as nfactura " & _
                          "FROM          (SELECT    case when ITF.idclaveprimariaremito = '' then F.trazabilidad_id else ITF.idclaveprimariaremito end as  trazabilidad_id, SUM(ITF.cantidadproducto) AS cantidadfacturada2, " & _
                          "F.tipodefactura + ' ' + F.puntodeventa + RIGHT('0000000' +F.numerodefactura, 8) AS NroFactura " & _
                          "FROM         ud_ezi_puntodeventa_encabezado AS F WITH (readpast) INNER JOIN " & _
                          "ud_ezi_puntodeventa_detalle_factm AS ITF WITH (readpast) ON F.id = ITF.claveprimaria " & _
                          "WHERE     (F.numeradorinterno = 'Factura de Venta') " & _
                          "GROUP BY F.trazabilidad_id, ITF.idclaveprimariaremito, tipodefactura, puntodeventa, numerodefactura " & _
                          "HAVING      (NOT (F.trazabilidad_id IS NULL))) AS FA RIGHT OUTER JOIN " & _
                          "v_ezi_pos_remito AS R ON FA.trazabilidad_id = R.id LEFT OUTER JOIN " & _
                          "ud_ezi_puntodeventa_detalle_factm AS IFAC WITH (readpast) RIGHT OUTER JOIN " & _
                          "ud_ezi_puntodeventa_encabezado AS E WITH (readpast) ON IFAC.claveprimaria = E.id ON R.trazabilidad_id = E.trazabilidad_id " & _
                          "GROUP BY R.id, R.Sel, R.NUMERODOCUMENTO, R.FECHAEMISION, R.cod_cliente, R.cliente, R.CUIT, R.vendedor, R.tipopago, R.TIPOVENTA, " & _
                          "E.tipodefactura + ' ' + E.puntodeventa + RIGHT('0000000' + E.numerodefactura, 8), R.presupuestobase, R.alquiler, R.saldado, E.numeradorinterno, " & _
                          "FA.cantidadfacturada2 , R.Estado, FA.NroFactura " & _
                          "HAVING      (R.alquiler = 'N') AND (R.NUMERODOCUMENTO + 'r  ' + R.cliente + ' ' + R.CUIT + ' ' + ISNULL(E.tipodefactura + ' ' + E.puntodeventa + RIGHT('0000000' + E.numerodefactura, 8), '') + 'f' LIKE '" & xbusqueda & "') " & _
                          "AND (R.estado = 'Remitido')) AS RC " & _
                          "GROUP BY id, Sel, NroRemito, Fecha, CodCliente, Cliente, CUIT, Vendedor, Tipopago, TipodeVenta, presupuestobase, saldado " & _
                          "ORDER BY id DESC"
              Else
                xquery1 = "SELECT     id, Sel, NroRemito, Fecha, presupuestobase AS NV, CodCliente, Cliente, CUIT, case when LEFT( MIN(ISNULL(NroFactura, nfactura)),1) = 'R' then null else MIN(ISNULL(NroFactura, nfactura)) end  AS NroFactura, Tipopago, TipodeVenta, Vendedor, MAX(cantremitida) " & _
                          "AS cantremitida, SUM(cantfacturada) AS cantfacturada, MAX(concatenado) AS concatenado, saldado " & _
                          "FROM         (SELECT     R.id, R.Sel, R.NUMERODOCUMENTO AS NroRemito, R.FECHAEMISION AS Fecha, R.cod_cliente AS CodCliente, R.cliente AS Cliente, R.CUIT, " & _
                          "R.vendedor AS Vendedor, R.tipopago AS Tipopago, R.TIPOVENTA AS TipodeVenta, R.presupuestobase, " & _
                          "E.tipodefactura + ' ' + E.puntodeventa + RIGHT('0000000' + E.numerodefactura, 8) AS NroFactura, SUM(R.cantidadremitida) AS cantremitida, " & _
                          "SUM(ISNULL(IFAC.cantidadproducto, 0)) + ISNULL(FA.cantidadfacturada2, 0) AS cantfacturada, " & _
                          "R.NUMERODOCUMENTO + 'r  ' + R.cliente + ' ' + R.CUIT + ' ' + ISNULL(E.tipodefactura + ' ' + E.puntodeventa + RIGHT('0000000' + E.numerodefactura, 8), '') " & _
                          "+ 'f' AS concatenado, R.saldado, E.numeradorinterno, FA.cantidadfacturada2, FA.NroFactura as nfactura " & _
                          "FROM          (SELECT    case when ITF.idclaveprimariaremito = '' then F.trazabilidad_id else ITF.idclaveprimariaremito end as  trazabilidad_id, SUM(ITF.cantidadproducto) AS cantidadfacturada2, " & _
                          "F.tipodefactura + ' ' + F.puntodeventa + RIGHT('0000000' +F.numerodefactura, 8) AS NroFactura " & _
                          "FROM         ud_ezi_puntodeventa_encabezado AS F WITH (readpast) INNER JOIN " & _
                          "ud_ezi_puntodeventa_detalle_factm AS ITF WITH (readpast) ON F.id = ITF.claveprimaria " & _
                          "WHERE     (F.numeradorinterno = 'Factura de Venta') " & _
                          "GROUP BY F.trazabilidad_id, ITF.idclaveprimariaremito, tipodefactura, puntodeventa, numerodefactura " & _
                          "HAVING      (NOT (F.trazabilidad_id IS NULL))) AS FA RIGHT OUTER JOIN " & _
                          "v_ezi_pos_remito AS R ON FA.trazabilidad_id = R.id LEFT OUTER JOIN " & _
                          "ud_ezi_puntodeventa_detalle_factm AS IFAC WITH (readpast) RIGHT OUTER JOIN " & _
                          "ud_ezi_puntodeventa_encabezado AS E WITH (readpast) ON IFAC.claveprimaria = E.id ON R.trazabilidad_id = E.trazabilidad_id " & _
                          "GROUP BY R.id, R.Sel, R.NUMERODOCUMENTO, R.FECHAEMISION, R.cod_cliente, R.cliente, R.CUIT, R.vendedor, R.tipopago, R.TIPOVENTA, " & _
                          "E.tipodefactura + ' ' + E.puntodeventa + RIGHT('0000000' + E.numerodefactura, 8), R.presupuestobase, R.alquiler, R.saldado, E.numeradorinterno, " & _
                          "FA.cantidadfacturada2 , R.Estado, FA.NroFactura " & _
                          "HAVING      (R.alquiler = 'N') AND (R.NUMERODOCUMENTO + 'r  ' + R.cliente + ' ' + R.CUIT + ' ' + ISNULL(E.tipodefactura + ' ' + E.puntodeventa + RIGHT('0000000' + E.numerodefactura, 8), '') + 'f' LIKE '" & xbusqueda & "') " & _
                          "AND (R.estado = 'Remitido')) AS RC " & _
                          "GROUP BY id, Sel, NroRemito, Fecha, CodCliente, Cliente, CUIT, Vendedor, Tipopago, TipodeVenta, presupuestobase, saldado " & _
                          "Having presupuestobase = '" & Text2.Text & "' ORDER BY id DESC"
              End If
            End If
                    
            datremitos.RecordSource = xquery1
            datremitos.Refresh
            If Text1.Text = " " Then Text1.Text = ""
            For X = 2 To 14
                DataGrid1.Columns(X).Locked = True
            Next X
            If Text2.Text = "" Then
                DataGrid1.Columns(1).Locked = True
            Else
                DataGrid1.Columns(1).Locked = False
            End If
            DataGrid1.Columns(1).Width = 300
            DataGrid1.Columns(4).Width = 1000
            DataGrid1.Columns(5).Width = 1000
            DataGrid1.Columns(6).Width = 3500
            
            DataGrid1.Columns(0).Visible = False
            DataGrid1.Columns(9).Visible = False
            DataGrid1.Columns(10).Visible = False
            DataGrid1.Columns(12).Visible = False
            DataGrid1.Columns(13).Visible = False
            DataGrid1.Columns(14).Visible = False
            DataGrid1.Columns(15).Visible = False
            
            Call DataGrid1_Click
            
        End If
        DataGrid1.SetFocus
        
        
    End If


End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
On Error Resume Next

    If KeyAscii = 13 Then
        KeyAscii = 0
        Text1.SetFocus
        SendKeys "{ENTER}", False
    End If

End Sub
