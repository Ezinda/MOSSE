VERSION 5.00
Object = "{F6125AB1-8AB1-11CE-A77F-08002B2F4E98}#2.0#0"; "MSRDC20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{D18BBD1F-82BB-4385-BED3-E9D31A3E361E}#1.0#0"; "KewlButtonz.ocx"
Begin VB.Form lista_presupuestos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta de Cotizaciones"
   ClientHeight    =   7590
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15525
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7590
   ScaleWidth      =   15525
   Begin VB.TextBox Text5 
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
      TabIndex        =   14
      Top             =   600
      Width           =   2295
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Nro.Exp"
      Height          =   375
      Left            =   120
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   600
      Width           =   855
   End
   Begin VB.TextBox Text3 
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
      Left            =   4680
      TabIndex        =   12
      Top             =   600
      Width           =   2295
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Nro.Lic."
      Height          =   375
      Left            =   3720
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   600
      Width           =   855
   End
   Begin VB.TextBox Text4 
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
      Left            =   9240
      TabIndex        =   10
      Top             =   600
      Width           =   2295
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Nro.Compra Directa"
      Height          =   375
      Left            =   7320
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   600
      Width           =   1815
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "lista_presupuestos.frx":0000
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   15255
      _ExtentX        =   26908
      _ExtentY        =   4895
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      BackColor       =   16777215
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
      TabIndex        =   5
      Top             =   0
      Width           =   15525
      _ExtentX        =   27384
      _ExtentY        =   953
      ButtonWidth     =   609
      ButtonHeight    =   900
      Appearance      =   1
      Style           =   1
      _Version        =   393216
      Begin Crystal.CrystalReport CrystalReport1 
         Left            =   9120
         Top             =   120
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
         WindowShowPrintSetupBtn=   -1  'True
      End
      Begin VB.TextBox Text2 
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
         Left            =   7680
         TabIndex        =   2
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton Command2 
         Caption         =   "OC:"
         Height          =   375
         Left            =   6720
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   120
         Width           =   855
      End
      Begin MSRDC.MSRDC reporte 
         Height          =   375
         Left            =   5280
         Top             =   5880
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
      Begin KewlButtonz.KewlButtons salir 
         Height          =   495
         Left            =   11520
         TabIndex        =   4
         Top             =   0
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
         MICON           =   "lista_presupuestos.frx":001D
         PICN            =   "lista_presupuestos.frx":0039
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin KewlButtonz.KewlButtons Command4 
         Height          =   495
         Left            =   9720
         TabIndex        =   3
         Top             =   0
         Width           =   1575
         _ExtentX        =   2778
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
         MICON           =   "lista_presupuestos.frx":0B83
         PICN            =   "lista_presupuestos.frx":0B9F
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
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
         Left            =   1080
         TabIndex        =   1
         Top             =   120
         Width           =   5415
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Buscar:"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   120
         Width           =   855
      End
   End
   Begin MSAdodcLib.Adodc datpresupuesto 
      Height          =   330
      Left            =   0
      Top             =   6840
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
   Begin MSDataGridLib.DataGrid DataGrid2 
      Bindings        =   "lista_presupuestos.frx":3F91
      Height          =   3255
      Left            =   120
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   4200
      Width           =   15255
      _ExtentX        =   26908
      _ExtentY        =   5741
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      BackColor       =   16777215
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
   Begin MSAdodcLib.Adodc datitems 
      Height          =   330
      Left            =   1200
      Top             =   6840
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
      Caption         =   "datitems"
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
Attribute VB_Name = "lista_presupuestos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public cuentacont As String
Dim cuenta(99999) As Integer
Private Const CAMPO_expediente As String = "Nro_Exp"
Private Const CAMPO_licitacion As String = "Nro_Lic"
Private Const CAMPO_compradirecta As String = "Nro_Compradirecta"

Private Sub Command4_Click()
On Error GoTo fuera
Dim crApp As New CRAXDRT.Application
Dim crReport As New CRAXDRT.Report
Dim tabla As String
Dim ruta As String

reporte.SQL = "SELECT v_ezi_pos_presupuesto.NUMERODOCUMENTO, v_ezi_pos_presupuesto.FECHAEMISION, v_ezi_pos_presupuesto.cod_cliente, v_ezi_pos_presupuesto.cliente, v_ezi_pos_presupuesto.CUIT, v_ezi_pos_presupuesto.CODPOS, v_ezi_pos_presupuesto.provincia, v_ezi_pos_presupuesto.vendedor, v_ezi_pos_presupuesto.detalle, v_ezi_pos_presupuesto.tipopago, v_ezi_pos_presupuesto.codigoproducto, v_ezi_pos_presupuesto.nombre_producto, v_ezi_pos_presupuesto.cantidadproducto, v_ezi_pos_presupuesto.nota, v_ezi_pos_presupuesto.condiva, v_ezi_pos_presupuesto.ciudad, v_ezi_pos_presupuesto.TIPOVENTA, v_ezi_pos_presupuesto.SIMBOLO, v_ezi_pos_presupuesto.CODVENDEDOR, v_ezi_pos_presupuesto.preciusiniva, v_ezi_pos_presupuesto.subtotalsiniva, v_ezi_pos_presupuesto.impbonifsiniva, v_ezi_pos_presupuesto.percepiibb, v_ezi_pos_presupuesto.perceptem, v_ezi_pos_presupuesto.totaltr, v_ezi_pos_presupuesto.importeiva21, v_ezi_pos_presupuesto.importeiva105 FROM MMOSSE.dbo.v_ezi_pos_presupuesto v_ezi_pos_presupuesto " & _
              " where v_ezi_pos_presupuesto.id = " & DataGrid1.Columns(7).Text & " ORDER BY case when ascii(SUBSTRING(( case when len(v_ezi_pos_presupuesto.item)= 1 then '0'+v_ezi_pos_presupuesto.item else v_ezi_pos_presupuesto.item end),2,1)) >= 48 and ascii(SUBSTRING(( case when len(v_ezi_pos_presupuesto.item)= 1 then '0'+v_ezi_pos_presupuesto.item else v_ezi_pos_presupuesto.item end),2,1)) <= 57 then  case when len(v_ezi_pos_presupuesto.item)= 1 then '0'+v_ezi_pos_presupuesto.item else v_ezi_pos_presupuesto.item end else '0' +  case when len(v_ezi_pos_presupuesto.item)= 1 then '0'+v_ezi_pos_presupuesto.item else v_ezi_pos_presupuesto.item end end "


'reporte.SQL = "SELECT v_ezi_pos_presupuesto.id, v_ezi_pos_presupuesto.NUMERODOCUMENTO, v_ezi_pos_presupuesto.FECHAEMISION, v_ezi_pos_presupuesto.cod_cliente, v_ezi_pos_presupuesto.cliente, v_ezi_pos_presupuesto.CUIT, v_ezi_pos_presupuesto.CALLE, v_ezi_pos_presupuesto.CODPOS, v_ezi_pos_presupuesto.provincia, v_ezi_pos_presupuesto.detalle, v_ezi_pos_presupuesto.tipopago, v_ezi_pos_presupuesto.codigoproducto, v_ezi_pos_presupuesto.nombre_producto, v_ezi_pos_presupuesto.cantidadproducto, v_ezi_pos_presupuesto.nota, v_ezi_pos_presupuesto.condiva, v_ezi_pos_presupuesto.ciudad, v_ezi_pos_presupuesto.TIPOVENTA, v_ezi_pos_presupuesto.SIMBOLO, v_ezi_pos_presupuesto.CODVENDEDOR, v_ezi_pos_presupuesto.preciusiniva, v_ezi_pos_presupuesto.subtotalsiniva, v_ezi_pos_presupuesto.impbonifsiniva, v_ezi_pos_presupuesto.nroremito, v_ezi_pos_presupuesto.percepiibb, v_ezi_pos_presupuesto.perceptem, v_ezi_pos_presupuesto.totaltr, " & _
'              "v_ezi_pos_presupuesto.importeiva21, v_ezi_pos_presupuesto.importeiva105, v_ezi_pos_presupuesto.iditem " & _
'              "FROM dbo.v_ezi_pos_presupuesto v_ezi_pos_presupuesto " & _
'              "where v_ezi_pos_presupuesto.id = " & DataGrid1.Columns(7).Text & " order by v_ezi_pos_presupuesto.iditem"

tabla = reporte.SQL

With CrystalReporte
    .PrinterCollation = crptCollated
    If DataGrid1.Columns(9).Text = "A" Then
        .ReportFileName = App.Path & "\CotizacionA.rpt"
    Else
        .ReportFileName = App.Path & "\CotizacionB.rpt"
    End If
   
    .Connect = login.conexionreporte
    .DiscardSavedData = True
    .RetrieveDataFiles
    .ReportSource = 0
    .SQLQuery = tabla
    .Destination = crptToWindow
'    .Destination = crptToPrinter
    .PrintFileType = crptCrystal
    .WindowState = crptMaximized
    
    .Action = 1
    
End With

Exit Sub

fuera:
    
    MsgBox "Reporte de Presupuesto no Encontado, o error de configuracion de reporte", vbCritical, "Error"



End Sub

Private Sub DataGrid1_Click()

    xidencabezado = DataGrid1.Columns(7).Text
    datitems.RecordSource = "select codigoproducto as Codigo, nombre_producto as Descripcion, cantidadproducto as Cantidad, unidaddemedidaid as Um, preciou as Precio, subtotal as Subtotal from ud_ezi_puntodeventa_detalle_presu with (readpast) where claveprimaria = " & xidencabezado & ""
    datitems.Refresh
            DataGrid2.Columns(1).Width = 3500
            DataGrid2.Columns(2).Alignment = dbgRight
            DataGrid2.Columns(4).Alignment = dbgRight
            DataGrid2.Columns(5).Alignment = dbgRight


End Sub

Private Sub DataGrid1_DblClick()
    
If menu = 1 Then
            frmnota_venta.Text17.Text = DataGrid1.Columns(1).Text
            frmnota_venta.Text18.Text = DataGrid1.Columns(7).Text
            frmnota_venta.Text17.SetFocus
            SendKeys "{ENTER}", False
            Unload Me
End If

If menu = 2 Then
            frmpresupuesto.Text17.Text = DataGrid1.Columns(1).Text
            frmpresupuesto.Text18.Text = DataGrid1.Columns(7).Text
            frmpresupuesto.Text17.SetFocus
            SendKeys "{ENTER}", False
            Unload Me
End If

If menu = 5 Then
            frmcomparativa.Text17.Text = DataGrid1.Columns(1).Text
            frmcomparativa.Text18.Text = DataGrid1.Columns(7).Text
            frmcomparativa.Text17.SetFocus
            SendKeys "{ENTER}", False
            Unload Me
End If
        
If menu = 6 Then
            frmcomparativa.Show
            frmcomparativa.Text17.Text = DataGrid1.Columns(1).Text
            frmcomparativa.Text18.Text = DataGrid1.Columns(7).Text
            frmcomparativa.Text17.SetFocus
            SendKeys "{ENTER}", False
            Unload Me
End If
        
        

End Sub

Private Sub DataGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    xidencabezado = DataGrid1.Columns(7).Text
    datitems.RecordSource = "select codigoproducto as Codigo, nombre_producto as Descripcion, cantidadproducto as Cantidad, unidaddemedidaid as Um, preciou as Precio, subtotal as Subtotal from ud_ezi_puntodeventa_detalle_presu with (readpast) where claveprimaria = " & xidencabezado & ""
    datitems.Refresh
            DataGrid2.Columns(1).Width = 3500
            DataGrid2.Columns(2).Alignment = dbgRight
            DataGrid2.Columns(4).Alignment = dbgRight
            DataGrid2.Columns(5).Alignment = dbgRight



End Sub

Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = 13 Then
        KeyAscii = 0
        If menu = 1 Then
                frmnota_venta.Text17.Text = DataGrid1.Columns(1).Text
                frmnota_venta.Text18.Text = DataGrid1.Columns(7).Text
                frmnota_venta.Text17.SetFocus
                SendKeys "{ENTER}", False
                Unload Me
        End If
        If menu = 2 Then
                frmpresupuesto.Text17.Text = DataGrid1.Columns(1).Text
                frmpresupuesto.Text18.Text = DataGrid1.Columns(7).Text
                frmpresupuesto.Text17.SetFocus
                SendKeys "{ENTER}", False
                Unload Me
        End If
        If menu = 5 Then
                frmcomparativa.Text17.Text = DataGrid1.Columns(1).Text
                frmcomparativa.Text18.Text = DataGrid1.Columns(7).Text
                frmcomparativa.Text17.SetFocus
                SendKeys "{ENTER}", False
                Unload Me
        End If
        
        If menu = 6 Then
            frmcomparativa.Show
            frmcomparativa.Text17.Text = DataGrid1.Columns(1).Text
            frmcomparativa.Text18.Text = DataGrid1.Columns(7).Text
            frmcomparativa.Text17.SetFocus
            SendKeys "{ENTER}", False
            Unload Me
        End If

        
    End If

End Sub


Private Sub DataGrid1_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next
   If KeyCode <> 13 Then
    xidencabezado = DataGrid1.Columns(7).Text
    datitems.RecordSource = "select codigoproducto as Codigo, nombre_producto as Descripcion, cantidadproducto as Cantidad, unidaddemedidaid as Um, preciou as Precio, subtotal as Subtotal from ud_ezi_puntodeventa_detalle_presu with (readpast) where claveprimaria = " & xidencabezado & ""
    datitems.Refresh
            DataGrid2.Columns(1).Width = 3500
            DataGrid2.Columns(2).Alignment = dbgRight
            DataGrid2.Columns(4).Alignment = dbgRight
            DataGrid2.Columns(5).Alignment = dbgRight
   End If

End Sub

Private Sub Form_Activate()

DataGrid1.SetFocus

End Sub

Private Sub Form_Load()
If menu = 2 Then
'       Aplicar_skin2 Me
    Aplicar_skin Me
Else
    Aplicar_skin Me
End If

MiFuncionDeAjuste Me, True

yventana = Inicio.Height / 2 - 1000
xventana = Inicio.Width / 2

lista_presupuestos.Top = yventana - lista_presupuestos.Height / 2
lista_presupuestos.Left = xventana - lista_presupuestos.Width / 2


datpresupuesto.ConnectionString = login.conexiontotal
datitems.ConnectionString = login.conexiontotal

If menu = 5 Then
    xquery1 = "SELECT     top 15 ud_ezi_puntodeventa_encabezado.claveprimaria AS Numero, ud_ezi_puntodeventa_encabezado.numerodefactura as Nro, ud_ezi_puntodeventa_encabezado.fechadelcomprobante AS Fecha, " & _
                      "ud_ezi_puntodeventa_encabezado.cliente AS Cliente, V_PERSONA_.CUIT, ud_ezi_puntodeventa_encabezado.vendedor AS Vendedor, " & _
                      "ud_ezi_puntodeventa_encabezado.importeglobal AS Importe, ud_ezi_puntodeventa_encabezado.id, ud_ezi_puntodeventa_encabezado.generada, ud_ezi_puntodeventa_encabezado.tipodefactura, " & _
                      "ud_ezi_puntodeventa_encabezado.numerodefactura + 'n ' + ud_ezi_puntodeventa_encabezado.cliente + ' ' + V_PERSONA_.CUIT + ' ' + ud_ezi_puntodeventa_encabezado.vendedor AS concatenado, ud_ezi_puntodeventa_encabezado.adicionalid as OC,case when ud_ezi_puntodeventa_encabezado.leslicitacion = 1 then 'S' else 'N' end as Licitacion, " & _
                      "ud_ezi_puntodeventa_encabezado.lexpediente as Nro_Exp,ud_ezi_puntodeventa_encabezado.llicitacionnro as Nro_Lic, ud_ezi_puntodeventa_encabezado.lcompradirectanro as Nro_CompraDirecta " & _
                      "FROM V_PERSONA_ RIGHT OUTER JOIN V_CLIENTE_ ON V_PERSONA_.ID = V_CLIENTE_.ENTEASOCIADO_ID RIGHT OUTER JOIN " & _
                      "ud_ezi_puntodeventa_encabezado WITH (readpast) ON V_CLIENTE_.ID = ud_ezi_puntodeventa_encabezado.clienteid " & _
                      "WHERE     (ud_ezi_puntodeventa_encabezado.numeradorinterno = 'Presupuesto de Venta') AND (ud_ezi_puntodeventa_encabezado.generada = 0) and ud_ezi_puntodeventa_encabezado.importado = 'False' and ud_ezi_puntodeventa_encabezado.flag = 0" & _
                      "ORDER BY Fecha DESC"
Else
  If menu = 6 Then
    xquery1 = "SELECT     top 15 ud_ezi_puntodeventa_encabezado.claveprimaria AS Numero, ud_ezi_puntodeventa_encabezado.numerodefactura as Nro, ud_ezi_puntodeventa_encabezado.fechadelcomprobante AS Fecha, " & _
                      "ud_ezi_puntodeventa_encabezado.cliente AS Cliente, V_PERSONA_.CUIT, ud_ezi_puntodeventa_encabezado.vendedor AS Vendedor, " & _
                      "ud_ezi_puntodeventa_encabezado.importeglobal AS Importe, ud_ezi_puntodeventa_encabezado.id, ud_ezi_puntodeventa_encabezado.generada, ud_ezi_puntodeventa_encabezado.tipodefactura, " & _
                      "ud_ezi_puntodeventa_encabezado.numerodefactura + 'n ' + ud_ezi_puntodeventa_encabezado.cliente + ' ' + V_PERSONA_.CUIT + ' ' + ud_ezi_puntodeventa_encabezado.vendedor AS concatenado, " & _
                      "ud_ezi_puntodeventa_encabezado.cprov1 as Prov1, ud_ezi_puntodeventa_encabezado.cprov2 as Prov2, ud_ezi_puntodeventa_encabezado.cprov3 as Prov3, " & _
                      "ud_ezi_puntodeventa_encabezado.cprov4 as Prov4, ud_ezi_puntodeventa_encabezado.cprov5 as Prov5, ud_ezi_puntodeventa_encabezado.cprov6 as Prov6, " & _
                      "ud_ezi_puntodeventa_encabezado.lexpediente as Nro_Exp,ud_ezi_puntodeventa_encabezado.llicitacionnro as Nro_Lic, ud_ezi_puntodeventa_encabezado.lcompradirectanro as Nro_CompraDirecta " & _
                      "FROM V_PERSONA_ RIGHT OUTER JOIN V_CLIENTE_ ON V_PERSONA_.ID = V_CLIENTE_.ENTEASOCIADO_ID RIGHT OUTER JOIN " & _
                      "ud_ezi_puntodeventa_encabezado WITH (readpast) ON V_CLIENTE_.ID = ud_ezi_puntodeventa_encabezado.clienteid " & _
                      "WHERE     (ud_ezi_puntodeventa_encabezado.numeradorinterno = 'Comparativa de Precios') AND (ud_ezi_puntodeventa_encabezado.generada = 0) " & _
                      "ORDER BY Fecha DESC"
                      
    lista_presupuestos.Caption = "Consulta de Comparativas"
  Else
    xquery1 = "SELECT     top 15 ud_ezi_puntodeventa_encabezado.claveprimaria AS Numero, ud_ezi_puntodeventa_encabezado.numerodefactura as Nro, ud_ezi_puntodeventa_encabezado.fechadelcomprobante AS Fecha, " & _
                      "ud_ezi_puntodeventa_encabezado.cliente AS Cliente, V_PERSONA_.CUIT, ud_ezi_puntodeventa_encabezado.vendedor AS Vendedor, " & _
                      "ud_ezi_puntodeventa_encabezado.importeglobal AS Importe, ud_ezi_puntodeventa_encabezado.id, ud_ezi_puntodeventa_encabezado.generada, ud_ezi_puntodeventa_encabezado.tipodefactura, " & _
                      "ud_ezi_puntodeventa_encabezado.numerodefactura + 'n ' + ud_ezi_puntodeventa_encabezado.cliente + ' ' + V_PERSONA_.CUIT + ' ' + ud_ezi_puntodeventa_encabezado.vendedor AS concatenado, ud_ezi_puntodeventa_encabezado.adicionalid as OC,case when ud_ezi_puntodeventa_encabezado.leslicitacion = 1 then 'S' else 'N' end as Licitacion, " & _
                      "ud_ezi_puntodeventa_encabezado.lexpediente as Nro_Exp,ud_ezi_puntodeventa_encabezado.llicitacionnro as Nro_Lic, ud_ezi_puntodeventa_encabezado.lcompradirectanro as Nro_CompraDirecta " & _
                      "FROM V_PERSONA_ RIGHT OUTER JOIN V_CLIENTE_ ON V_PERSONA_.ID = V_CLIENTE_.ENTEASOCIADO_ID RIGHT OUTER JOIN " & _
                      "ud_ezi_puntodeventa_encabezado WITH (readpast) ON V_CLIENTE_.ID = ud_ezi_puntodeventa_encabezado.clienteid " & _
                      "WHERE     (ud_ezi_puntodeventa_encabezado.numeradorinterno = 'Presupuesto de Venta') AND (ud_ezi_puntodeventa_encabezado.generada = 0) and ud_ezi_puntodeventa_encabezado.flag = 0 " & _
                      "ORDER BY Fecha DESC"
  End If
End If

datpresupuesto.RecordSource = xquery1
datpresupuesto.Refresh

            
            DataGrid1.Columns(0).Visible = False
            DataGrid1.Columns(7).Visible = False
            DataGrid1.Columns(8).Visible = False
            DataGrid1.Columns(9).Visible = False
            DataGrid1.Columns(10).Visible = False
            DataGrid1.Columns(11).Width = 800
            DataGrid1.Columns(1).Width = 1000
            DataGrid1.Columns(3).Width = 3500
            If menu = 6 Then
                DataGrid1.Columns(6).Visible = False
            End If
            DataGrid1.Columns(6).Alignment = dbgRight
            DataGrid1.Columns(6).NumberFormat = "Currency"
 
End Sub

Private Sub salir_Click()
    
    Unload Me

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)

On Error Resume Next


    If KeyAscii = 13 Then
        KeyAscii = 0
        Text5.Text = ""
        Text3.Text = ""
        Text4.Text = ""
        If Text1.Text <> "" Then
            Text1.Text = Replace(Text1.Text, " ", "%%")
            xbusqueda = "%" + Text1.Text + "%"
            If menu = 5 Then
              If Text2.Text = "" Then
                xquery1 = "SELECT   ud_ezi_puntodeventa_encabezado.claveprimaria AS Numero, ud_ezi_puntodeventa_encabezado.numerodefactura as Nro, ud_ezi_puntodeventa_encabezado.fechadelcomprobante AS Fecha, " & _
                      "ud_ezi_puntodeventa_encabezado.cliente AS Cliente, V_PERSONA_.CUIT, ud_ezi_puntodeventa_encabezado.vendedor AS Vendedor, " & _
                      "ud_ezi_puntodeventa_encabezado.importeglobal AS Importe, ud_ezi_puntodeventa_encabezado.id, ud_ezi_puntodeventa_encabezado.generada, ud_ezi_puntodeventa_encabezado.tipodefactura, " & _
                      "ud_ezi_puntodeventa_encabezado.numerodefactura + 'n ' + ud_ezi_puntodeventa_encabezado.cliente + ' ' + V_PERSONA_.CUIT + ' ' + ud_ezi_puntodeventa_encabezado.vendedor AS concatenado, ud_ezi_puntodeventa_encabezado.adicionalid as OC,case when ud_ezi_puntodeventa_encabezado.leslicitacion = 1 then 'S' else 'N' end as Licitacion, " & _
                      "ud_ezi_puntodeventa_encabezado.lexpediente as Nro_Exp,ud_ezi_puntodeventa_encabezado.llicitacionnro as Nro_Lic, ud_ezi_puntodeventa_encabezado.lcompradirectanro as Nro_CompraDirecta " & _
                      "FROM V_PERSONA_ RIGHT OUTER JOIN V_CLIENTE_ ON V_PERSONA_.ID = V_CLIENTE_.ENTEASOCIADO_ID RIGHT OUTER JOIN " & _
                      "ud_ezi_puntodeventa_encabezado WITH (readpast) ON V_CLIENTE_.ID = ud_ezi_puntodeventa_encabezado.clienteid " & _
                      "WHERE     (ud_ezi_puntodeventa_encabezado.numeradorinterno = 'Presupuesto de Venta') AND (ud_ezi_puntodeventa_encabezado.generada = 0) AND " & _
                      "(ud_ezi_puntodeventa_encabezado.numerodefactura + 'n ' + ud_ezi_puntodeventa_encabezado.cliente + ' ' + V_PERSONA_.CUIT + ' ' + ud_ezi_puntodeventa_encabezado.vendedor like '" & xbusqueda & "') and ud_ezi_puntodeventa_encabezado.importado = 'False' And ud_ezi_puntodeventa_encabezado.Flag = 0 " & _
                      "ORDER BY Fecha DESC"
              Else
                 xquery1 = "SELECT     ud_ezi_puntodeventa_encabezado.claveprimaria AS Numero, ud_ezi_puntodeventa_encabezado.numerodefactura as Nro, ud_ezi_puntodeventa_encabezado.fechadelcomprobante AS Fecha, " & _
                      "ud_ezi_puntodeventa_encabezado.cliente AS Cliente, V_PERSONA_.CUIT, ud_ezi_puntodeventa_encabezado.vendedor AS Vendedor, " & _
                      "ud_ezi_puntodeventa_encabezado.importeglobal AS Importe, ud_ezi_puntodeventa_encabezado.id, ud_ezi_puntodeventa_encabezado.generada, ud_ezi_puntodeventa_encabezado.tipodefactura, " & _
                      "ud_ezi_puntodeventa_encabezado.numerodefactura + 'n ' + ud_ezi_puntodeventa_encabezado.cliente + ' ' + V_PERSONA_.CUIT + ' ' + ud_ezi_puntodeventa_encabezado.vendedor AS concatenado, ud_ezi_puntodeventa_encabezado.adicionalid as OC,case when ud_ezi_puntodeventa_encabezado.leslicitacion = 1 then 'S' else 'N' end as Licitacion, " & _
                      "ud_ezi_puntodeventa_encabezado.lexpediente as Nro_Exp,ud_ezi_puntodeventa_encabezado.llicitacionnro as Nro_Lic, ud_ezi_puntodeventa_encabezado.lcompradirectanro as Nro_CompraDirecta " & _
                      "FROM V_PERSONA_ RIGHT OUTER JOIN V_CLIENTE_ ON V_PERSONA_.ID = V_CLIENTE_.ENTEASOCIADO_ID RIGHT OUTER JOIN " & _
                      "ud_ezi_puntodeventa_encabezado WITH (readpast) ON V_CLIENTE_.ID = ud_ezi_puntodeventa_encabezado.clienteid " & _
                      "WHERE     (ud_ezi_puntodeventa_encabezado.numeradorinterno = 'Presupuesto de Venta') AND (ud_ezi_puntodeventa_encabezado.generada = 0) AND " & _
                      "(ud_ezi_puntodeventa_encabezado.numerodefactura + 'n ' + ud_ezi_puntodeventa_encabezado.cliente + ' ' + V_PERSONA_.CUIT + ' ' + ud_ezi_puntodeventa_encabezado.vendedor like '" & xbusqueda & "') and ud_ezi_puntodeventa_encabezado.importado = 'False' And ud_ezi_puntodeventa_encabezado.Flag = 0  and ud_ezi_puntodeventa_encabezado.adicionalid = '" & Text2.Text & "'  " & _
                      "ORDER BY Fecha DESC"
              End If
            Else
              If menu = 6 Then
                 xquery1 = "SELECT     ud_ezi_puntodeventa_encabezado.claveprimaria AS Numero, ud_ezi_puntodeventa_encabezado.numerodefactura as Nro, ud_ezi_puntodeventa_encabezado.fechadelcomprobante AS Fecha, " & _
                      "ud_ezi_puntodeventa_encabezado.cliente AS Cliente, V_PERSONA_.CUIT, ud_ezi_puntodeventa_encabezado.vendedor AS Vendedor, " & _
                      "ud_ezi_puntodeventa_encabezado.importeglobal AS Importe, ud_ezi_puntodeventa_encabezado.id, ud_ezi_puntodeventa_encabezado.generada, ud_ezi_puntodeventa_encabezado.tipodefactura, " & _
                      "ud_ezi_puntodeventa_encabezado.numerodefactura + 'n ' + ud_ezi_puntodeventa_encabezado.cliente + ' ' + V_PERSONA_.CUIT + ' ' + ud_ezi_puntodeventa_encabezado.vendedor AS concatenado, " & _
                      "ud_ezi_puntodeventa_encabezado.cprov1 as Prov1, ud_ezi_puntodeventa_encabezado.cprov2 as Prov2, ud_ezi_puntodeventa_encabezado.cprov3 as Prov3, " & _
                      "ud_ezi_puntodeventa_encabezado.cprov4 as Prov4, ud_ezi_puntodeventa_encabezado.cprov5 as Prov5, ud_ezi_puntodeventa_encabezado.cprov6 as Prov6, " & _
                      "ud_ezi_puntodeventa_encabezado.lexpediente as Nro_Exp,ud_ezi_puntodeventa_encabezado.llicitacionnro as Nro_Lic, ud_ezi_puntodeventa_encabezado.lcompradirectanro as Nro_CompraDirecta " & _
                      "FROM V_PERSONA_ RIGHT OUTER JOIN V_CLIENTE_ ON V_PERSONA_.ID = V_CLIENTE_.ENTEASOCIADO_ID RIGHT OUTER JOIN " & _
                      "ud_ezi_puntodeventa_encabezado WITH (readpast) ON V_CLIENTE_.ID = ud_ezi_puntodeventa_encabezado.clienteid " & _
                      "WHERE     (ud_ezi_puntodeventa_encabezado.numeradorinterno = 'Comparativa de Precios') AND (ud_ezi_puntodeventa_encabezado.generada = 0) AND " & _
                      "(ud_ezi_puntodeventa_encabezado.numerodefactura + 'n ' + ud_ezi_puntodeventa_encabezado.cliente + ' ' + V_PERSONA_.CUIT + ' ' + ud_ezi_puntodeventa_encabezado.vendedor like '" & xbusqueda & "') And ud_ezi_puntodeventa_encabezado.Flag = 0 " & _
                      "ORDER BY Fecha DESC"
              Else
               If Text2.Text = "" Then
                 xquery1 = "SELECT     ud_ezi_puntodeventa_encabezado.claveprimaria AS Numero, ud_ezi_puntodeventa_encabezado.numerodefactura as Nro, ud_ezi_puntodeventa_encabezado.fechadelcomprobante AS Fecha, " & _
                      "ud_ezi_puntodeventa_encabezado.cliente AS Cliente, V_PERSONA_.CUIT, ud_ezi_puntodeventa_encabezado.vendedor AS Vendedor, " & _
                      "ud_ezi_puntodeventa_encabezado.importeglobal AS Importe, ud_ezi_puntodeventa_encabezado.id, ud_ezi_puntodeventa_encabezado.generada, ud_ezi_puntodeventa_encabezado.tipodefactura, " & _
                      "ud_ezi_puntodeventa_encabezado.numerodefactura + 'n ' + ud_ezi_puntodeventa_encabezado.cliente + ' ' + V_PERSONA_.CUIT + ' ' + ud_ezi_puntodeventa_encabezado.vendedor AS concatenado, ud_ezi_puntodeventa_encabezado.adicionalid as OC,case when ud_ezi_puntodeventa_encabezado.leslicitacion = 1 then 'S' else 'N' end as Licitacion, " & _
                      "ud_ezi_puntodeventa_encabezado.lexpediente as Nro_Exp,ud_ezi_puntodeventa_encabezado.llicitacionnro as Nro_Lic, ud_ezi_puntodeventa_encabezado.lcompradirectanro as Nro_CompraDirecta " & _
                      "FROM V_PERSONA_ RIGHT OUTER JOIN V_CLIENTE_ ON V_PERSONA_.ID = V_CLIENTE_.ENTEASOCIADO_ID RIGHT OUTER JOIN " & _
                      "ud_ezi_puntodeventa_encabezado WITH (readpast) ON V_CLIENTE_.ID = ud_ezi_puntodeventa_encabezado.clienteid " & _
                      "WHERE     (ud_ezi_puntodeventa_encabezado.numeradorinterno = 'Presupuesto de Venta') AND (ud_ezi_puntodeventa_encabezado.generada = 0) AND " & _
                      "(ud_ezi_puntodeventa_encabezado.numerodefactura + 'n ' + ud_ezi_puntodeventa_encabezado.cliente + ' ' + V_PERSONA_.CUIT + ' ' + ud_ezi_puntodeventa_encabezado.vendedor like '" & xbusqueda & "')  and ud_ezi_puntodeventa_encabezado.flag = 0 " & _
                      "ORDER BY Fecha DESC"
               Else
                  xquery1 = "SELECT     ud_ezi_puntodeventa_encabezado.claveprimaria AS Numero, ud_ezi_puntodeventa_encabezado.numerodefactura as Nro, ud_ezi_puntodeventa_encabezado.fechadelcomprobante AS Fecha, " & _
                      "ud_ezi_puntodeventa_encabezado.cliente AS Cliente, V_PERSONA_.CUIT, ud_ezi_puntodeventa_encabezado.vendedor AS Vendedor, " & _
                      "ud_ezi_puntodeventa_encabezado.importeglobal AS Importe, ud_ezi_puntodeventa_encabezado.id, ud_ezi_puntodeventa_encabezado.generada, ud_ezi_puntodeventa_encabezado.tipodefactura, " & _
                      "ud_ezi_puntodeventa_encabezado.numerodefactura + 'n ' + ud_ezi_puntodeventa_encabezado.cliente + ' ' + V_PERSONA_.CUIT + ' ' + ud_ezi_puntodeventa_encabezado.vendedor AS concatenado, ud_ezi_puntodeventa_encabezado.adicionalid as OC,case when ud_ezi_puntodeventa_encabezado.leslicitacion = 1 then 'S' else 'N' end as Licitacion, " & _
                      "ud_ezi_puntodeventa_encabezado.lexpediente as Nro_Exp,ud_ezi_puntodeventa_encabezado.llicitacionnro as Nro_Lic, ud_ezi_puntodeventa_encabezado.lcompradirectanro as Nro_CompraDirecta " & _
                      "FROM V_PERSONA_ RIGHT OUTER JOIN V_CLIENTE_ ON V_PERSONA_.ID = V_CLIENTE_.ENTEASOCIADO_ID RIGHT OUTER JOIN " & _
                      "ud_ezi_puntodeventa_encabezado WITH (readpast) ON V_CLIENTE_.ID = ud_ezi_puntodeventa_encabezado.clienteid " & _
                      "WHERE     (ud_ezi_puntodeventa_encabezado.numeradorinterno = 'Presupuesto de Venta') AND (ud_ezi_puntodeventa_encabezado.generada = 0) AND " & _
                      "(ud_ezi_puntodeventa_encabezado.numerodefactura + 'n ' + ud_ezi_puntodeventa_encabezado.cliente + ' ' + V_PERSONA_.CUIT + ' ' + ud_ezi_puntodeventa_encabezado.vendedor like '" & xbusqueda & "')  and ud_ezi_puntodeventa_encabezado.flag = 0   and ud_ezi_puntodeventa_encabezado.adicionalid = '" & Text2.Text & "'   " & _
                      "ORDER BY Fecha DESC"
                End If
              End If
            End If
            
            datpresupuesto.RecordSource = xquery1
            datpresupuesto.Refresh
            DataGrid1.Columns(0).Visible = False
            DataGrid1.Columns(7).Visible = False
            DataGrid1.Columns(8).Visible = False
            DataGrid1.Columns(9).Visible = False
            DataGrid1.Columns(10).Visible = False
            DataGrid1.Columns(1).Width = 1000
            DataGrid1.Columns(3).Width = 3500
            DataGrid1.Columns(6).Alignment = dbgRight
            DataGrid1.Columns(6).NumberFormat = "Currency"
            
            Call DataGrid1_Click
            
            DataGrid2.Columns(1).Width = 3500
            DataGrid2.Columns(3).Alignment = dbgRight
            DataGrid2.Columns(5).Alignment = dbgRight
            DataGrid2.Columns(6).Alignment = dbgRight


        End If
        DataGrid1.SetFocus
        
        
    End If

End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
On Error Resume Next

    If KeyAscii = 13 Then
        KeyAscii = 0
        Text1.SetFocus
        If Text1.Text = "" Then Text1.Text = " "
        SendKeys "{ENTER}", False
    End If


End Sub
Private Sub Text5_Change()
On Error Resume Next

        datpresupuesto.Recordset.Filter = CAMPO_expediente & " LIKE '*" + Text5.Text + "*'"
   

End Sub

Private Sub Text3_Change()
On Error Resume Next

        datpresupuesto.Recordset.Filter = CAMPO_licitacion & " LIKE '*" + Text3.Text + "*'"
   

End Sub

Private Sub Text4_Change()
On Error Resume Next

        datpresupuesto.Recordset.Filter = CAMPO_compradirecta & " LIKE '*" + Text4.Text + "*'"
   

End Sub

