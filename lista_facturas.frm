VERSION 5.00
Object = "{F6125AB1-8AB1-11CE-A77F-08002B2F4E98}#2.0#0"; "MSRDC20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{D18BBD1F-82BB-4385-BED3-E9D31A3E361E}#1.0#0"; "KewlButtonz.ocx"
Begin VB.Form lista_facturas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta de Facturas"
   ClientHeight    =   7320
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14340
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7320
   ScaleWidth      =   14340
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "lista_facturas.frx":0000
      Height          =   6255
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   13935
      _ExtentX        =   24580
      _ExtentY        =   11033
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
      TabIndex        =   4
      Top             =   0
      Width           =   14340
      _ExtentX        =   25294
      _ExtentY        =   953
      ButtonWidth     =   609
      ButtonHeight    =   900
      Appearance      =   1
      Style           =   1
      _Version        =   393216
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
      Begin Crystal.CrystalReport CrystalReporte 
         Left            =   7800
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowTitle     =   "Presupusto de Venta"
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         PrinterCollation=   0
         PrintFileLinesPerPage=   60
      End
      Begin KewlButtonz.KewlButtons salir 
         Height          =   495
         Left            =   11520
         TabIndex        =   3
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
         MICON           =   "lista_facturas.frx":001D
         PICN            =   "lista_facturas.frx":0039
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
         TabIndex        =   2
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
         MICON           =   "lista_facturas.frx":0B83
         PICN            =   "lista_facturas.frx":0B9F
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
         TabIndex        =   5
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
      Bindings        =   "lista_facturas.frx":3F91
      Height          =   1215
      Left            =   120
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   6000
      Visible         =   0   'False
      Width           =   13095
      _ExtentX        =   23098
      _ExtentY        =   2143
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
Attribute VB_Name = "lista_facturas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public cuentacont As String
Dim cuenta(99999) As Integer



Private Sub Command4_Click()
On Error GoTo fuera
Dim crApp As New CRAXDRT.Application
Dim crReport As New CRAXDRT.Report
Dim tabla As String
Dim ruta As String

xbase = Left(login.nombrebd, 17)
                     
If xbase = "MMOSSEPRUEBA" Then
    reporte.SQL = "SELECT v_ezi_pos_factctacte.id, v_ezi_pos_factctacte.NUMERODOCUMENTO, v_ezi_pos_factctacte.FECHAEMISION, v_ezi_pos_factctacte.cod_cliente, v_ezi_pos_factctacte.cliente, v_ezi_pos_factctacte.CUIT, v_ezi_pos_factctacte.CALLE, v_ezi_pos_factctacte.CODPOS, v_ezi_pos_factctacte.provincia, v_ezi_pos_factctacte.detalle, v_ezi_pos_factctacte.tipopago, v_ezi_pos_factctacte.codigoproducto, v_ezi_pos_factctacte.nombre_producto, v_ezi_pos_factctacte.cantidadproducto, v_ezi_pos_factctacte.nota, v_ezi_pos_factctacte.condiva, v_ezi_pos_factctacte.ciudad, v_ezi_pos_factctacte.TIPOVENTA, v_ezi_pos_factctacte.SIMBOLO, v_ezi_pos_factctacte.CODVENDEDOR, v_ezi_pos_factctacte.preciusiniva, v_ezi_pos_factctacte.subtotalsiniva, v_ezi_pos_factctacte.impbonifsiniva, v_ezi_pos_factctacte.nroremito, v_ezi_pos_factctacte.percepiibb, v_ezi_pos_factctacte.perceptem, v_ezi_pos_factctacte.totaltr, v_ezi_pos_factctacte.importeiva21, v_ezi_pos_factctacte.importeiva105, v_ezi_pos_factctacte.iditem " & _
              "FROM  MMOSSEPRUEBA.dbo.v_ezi_pos_factctacte v_ezi_pos_factctacte " & _
              "where v_ezi_pos_factctacte.id = '" & DataGrid1.Columns(7).Text & "' order by v_ezi_pos_factctacte.iditem"
Else
    reporte.SQL = "SELECT v_ezi_pos_factctacte.id, v_ezi_pos_factctacte.NUMERODOCUMENTO, v_ezi_pos_factctacte.FECHAEMISION, v_ezi_pos_factctacte.cod_cliente, v_ezi_pos_factctacte.cliente, v_ezi_pos_factctacte.CUIT, v_ezi_pos_factctacte.CALLE, v_ezi_pos_factctacte.CODPOS, v_ezi_pos_factctacte.provincia, v_ezi_pos_factctacte.detalle, v_ezi_pos_factctacte.tipopago, v_ezi_pos_factctacte.codigoproducto, v_ezi_pos_factctacte.nombre_producto, v_ezi_pos_factctacte.cantidadproducto, v_ezi_pos_factctacte.nota, v_ezi_pos_factctacte.condiva, v_ezi_pos_factctacte.ciudad, v_ezi_pos_factctacte.TIPOVENTA, v_ezi_pos_factctacte.SIMBOLO, v_ezi_pos_factctacte.CODVENDEDOR, v_ezi_pos_factctacte.preciusiniva, v_ezi_pos_factctacte.subtotalsiniva, v_ezi_pos_factctacte.impbonifsiniva, v_ezi_pos_factctacte.nroremito, v_ezi_pos_factctacte.percepiibb, v_ezi_pos_factctacte.perceptem, v_ezi_pos_factctacte.totaltr, v_ezi_pos_factctacte.importeiva21, v_ezi_pos_factctacte.importeiva105, v_ezi_pos_factctacte.iditem " & _
              "FROM  MMOSSE.dbo.v_ezi_pos_factctacte v_ezi_pos_factctacte " & _
              "where v_ezi_pos_factctacte.id = '" & DataGrid1.Columns(7).Text & "' order by v_ezi_pos_factctacte.iditem"
End If


tabla = reporte.SQL
'Debug.Print reporte.SQL

With CrystalReporte
    .PrinterCollation = crptCollated
    If tipofac <> "NN" Then
        .Formulas(0) = "copia="" ORIGINAL """
    End If
    If DataGrid1.Columns(9).Text = "A" Then
'      If tipofac <> "NN" Then
'       If DataGrid1.Columns(11).Text = "N" Then
'        .ReportFileName = App.Path & "\FacturaCtaCteA.rpt"
'       Else
'        .ReportFileName = App.Path & "\FacturaCtaCteA_alquiler.rpt"
'       End If
'      Else
        .ReportFileName = App.Path & "\PresupuestoA.rpt"
'      End If
    Else
'      If tipofac <> "NN" Then
'       If DataGrid1.Columns(11).Text = "N" Then
'        .ReportFileName = App.Path & "\FacturaCtaCteB.rpt"
'       Else
'        .ReportFileName = App.Path & "\FacturaCtaCteB_alquiler.rpt"
'       End If
'      Else
        .ReportFileName = App.Path & "\PresupuestoB.rpt"
'      End If
    End If
    .WindowTitle = "Factura Vta Orig"
    .Connect = login.conexionreporte
    .DiscardSavedData = True
    .RetrieveDataFiles
    .ReportSource = 0
    .SQLQuery = tabla
    .Destination = crptToWindow
 '   .Destination = crptToPrinter
    .PrintFileType = crptCrystal
    .WindowState = crptMaximized
    .Action = 1
    .WindowTitle = "Factura Vta Dupl"
    .Formulas(0) = "copia="" DUPLICADO """
    .Action = 1
 '   If tipofac <> "NN" Then
 '   .WindowTitle = "Factura Vta Trip"
 '   .Formulas(0) = "copia="" TRIPLICADO """
 '   .Action = 1
 '   End If
End With
    
Exit Sub

fuera:
    
    MsgBox "Reporte de Factura no Encontado, o error de configuracion de reporte", vbCritical, "Error"


Exit Sub




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
    

            frmnota_credito.xsaldofactura = Val(DataGrid1.Columns(12).Text)
            frmnota_credito.Text17.Text = DataGrid1.Columns(1).Text
            frmnota_credito.Text18.Text = DataGrid1.Columns(7).Text
            frmnota_credito.Text17.SetFocus
            SendKeys "{ENTER}", False
            Unload Me

        

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

    If KeyAscii = 13 Then
        KeyAscii = 0
                If Val(DataGrid1.Columns(12).Text) = 0 Then
                  MsgBox "No hay saldo suficiente para hacer Nota de Credito sobre este comprobante", vbCritical, "Error"
                  Exit Sub
                End If
                frmnota_credito.xsaldofactura = Val(DataGrid1.Columns(12).Text)
                frmnota_credito.Text17.Text = DataGrid1.Columns(1).Text
                frmnota_credito.Text18.Text = DataGrid1.Columns(7).Text
                frmnota_credito.Text17.SetFocus
                SendKeys "{ENTER}", False
        Unload Me
    End If

End Sub


Private Sub DataGrid1_KeyUp(KeyCode As Integer, Shift As Integer)
    
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
Aplicar_skin Me

MiFuncionDeAjuste Me, True

yventana = Inicio.Height / 2 - 1000
xventana = Inicio.Width / 2

lista_facturas.Top = yventana - lista_facturas.Height / 2
lista_facturas.Left = xventana - lista_facturas.Width / 2


datpresupuesto.ConnectionString = login.conexiontotal
datitems.ConnectionString = login.conexiontotal

xsuc = login.nomsucursal
'xquery1 = "SELECT     top 15 ud_ezi_puntodeventa_encabezado.claveprimaria AS Numero, ud_ezi_puntodeventa_encabezado.numerodefactura as Nro, ud_ezi_puntodeventa_encabezado.fechadelcomprobante AS Fecha, " & _
'                      "ud_ezi_puntodeventa_encabezado.cliente AS Cliente, V_PERSONA_.CUIT, ud_ezi_puntodeventa_encabezado.vendedor AS Vendedor, " & _
'                      "ud_ezi_puntodeventa_encabezado.importeglobal AS Importe, ud_ezi_puntodeventa_encabezado.id, ud_ezi_puntodeventa_encabezado.generada, ud_ezi_puntodeventa_encabezado.tipodefactura, " & _
'                      "ud_ezi_puntodeventa_encabezado.numerodefactura + 'n ' + ud_ezi_puntodeventa_encabezado.cliente + ' ' + V_PERSONA_.CUIT + ' ' + ud_ezi_puntodeventa_encabezado.vendedor AS concatenado, ud_ezi_puntodeventa_encabezado.alquiler " & _
'                      "FROM V_PERSONA_ RIGHT OUTER JOIN V_CLIENTE_ ON V_PERSONA_.ID = V_CLIENTE_.ENTEASOCIADO_ID RIGHT OUTER JOIN " & _
'                      "ud_ezi_puntodeventa_encabezado WITH (readpast) ON V_CLIENTE_.ID = ud_ezi_puntodeventa_encabezado.clienteid " & _
'                      "WHERE     (ud_ezi_puntodeventa_encabezado.numeradorinterno like '%Factura de Venta%') and (ud_ezi_puntodeventa_encabezado.tipodefacturacionid <> 'NN') AND (ud_ezi_puntodeventa_encabezado.sucursal = '" & xsuc & "') " & _
'                      "ORDER BY Fecha DESC"
                      
'  NUevo 13/01/2017
xquery1 = "SELECT   TOP (30)  PV.claveprimaria AS Numero, PV.numerodefactura AS Nro, PV.fechadelcomprobante AS Fecha, PV.cliente AS Cliente, P.CUIT, PV.vendedor AS Vendedor, " & _
          "PV.importeglobal AS Importe, PV.id, PV.generada, PV.tipodefactura, PV.numerodefactura + 'n ' + PV.cliente + ' ' + P.CUIT + ' ' + PV.vendedor AS concatenado, " & _
          "PV.alquiler, CASE WHEN RIGHT(PV.numeradorinterno, 9) = 'Mostrador' THEN ISNULL(CNCC.saldo, PV.importeglobal) ELSE ISNULL(CP.SALDO2_IMPORTE, 0) " & _
          "END AS Saldo, CASE WHEN RIGHT(PV.numeradorinterno, 9) = 'Mostrador' THEN 'Contado' ELSE 'Cta.Cte.' END AS tipo " & _
          "FROM         v_ud_ezi_controlncc_pos AS CNCC RIGHT OUTER JOIN " & _
          "ud_ezi_puntodeventa_encabezado AS PV WITH (readpast) ON CNCC.idfactura = PV.id LEFT OUTER JOIN " & _
          "V_COMPROMISOPAGO_ AS CP ON PV.calipsoid = CP.TRORIGINANTE_ID LEFT OUTER JOIN " & _
          "V_PERSONA_ AS P RIGHT OUTER JOIN " & _
          "V_CLIENTE_ AS C ON P.ID = C.ENTEASOCIADO_ID ON PV.clienteid = C.ID " & _
          "WHERE     (PV.numeradorinterno LIKE '%Factura de Venta%') AND (PV.tipodefacturacionid <> 'NN') AND (pv.sucursal = '" & xsuc & "') AND (CP.NIVEL = 1 OR CP.NIVEL IS NULL) " & _
          "ORDER BY Fecha DESC "

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
            

 
End Sub

Private Sub salir_Click()
    
    Unload Me

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)

On Error Resume Next


    If KeyAscii = 13 Then
        KeyAscii = 0
        If Text1.Text <> "" Then
            xbusqueda = "%" + Text1.Text + "%"
            xsuc = login.nomsucursal
          '  xquery1 = "SELECT     ud_ezi_puntodeventa_encabezado.claveprimaria AS Numero, ud_ezi_puntodeventa_encabezado.numerodefactura as Nro, ud_ezi_puntodeventa_encabezado.fechadelcomprobante AS Fecha, " & _
          '            "ud_ezi_puntodeventa_encabezado.cliente AS Cliente, V_PERSONA_.CUIT, ud_ezi_puntodeventa_encabezado.vendedor AS Vendedor, " & _
          '            "ud_ezi_puntodeventa_encabezado.importeglobal AS Importe, ud_ezi_puntodeventa_encabezado.id, ud_ezi_puntodeventa_encabezado.generada, ud_ezi_puntodeventa_encabezado.tipodefactura, " & _
          '            "ud_ezi_puntodeventa_encabezado.numerodefactura + 'n ' + ud_ezi_puntodeventa_encabezado.cliente + ' ' + V_PERSONA_.CUIT + ' ' + ud_ezi_puntodeventa_encabezado.vendedor AS concatenado, ud_ezi_puntodeventa_encabezado.alquiler " & _
          '            "FROM V_PERSONA_ RIGHT OUTER JOIN V_CLIENTE_ ON V_PERSONA_.ID = V_CLIENTE_.ENTEASOCIADO_ID RIGHT OUTER JOIN " & _
          '            "ud_ezi_puntodeventa_encabezado WITH (readpast) ON V_CLIENTE_.ID = ud_ezi_puntodeventa_encabezado.clienteid " & _
          '            "WHERE     (ud_ezi_puntodeventa_encabezado.numeradorinterno like '%Factura de Venta%') and (ud_ezi_puntodeventa_encabezado.tipodefacturacionid <> 'NN') AND " & _
          '            "(ud_ezi_puntodeventa_encabezado.numerodefactura + 'n ' + ud_ezi_puntodeventa_encabezado.cliente + ' ' + V_PERSONA_.CUIT + ' ' + ud_ezi_puntodeventa_encabezado.vendedor like '" & xbusqueda & "') AND (ud_ezi_puntodeventa_encabezado.sucursal = '" & xsuc & "') " & _
          '            "ORDER BY Fecha DESC"
                      
'  Nuevo 13/01/2017
             xquery1 = "SELECT   TOP (30)  PV.claveprimaria AS Numero, PV.numerodefactura AS Nro, PV.fechadelcomprobante AS Fecha, PV.cliente AS Cliente, P.CUIT, PV.vendedor AS Vendedor, " & _
                      "PV.importeglobal AS Importe, PV.id, PV.generada, PV.tipodefactura, PV.numerodefactura + 'n ' + PV.cliente + ' ' + P.CUIT + ' ' + PV.vendedor AS concatenado, " & _
                      "PV.alquiler, CASE WHEN RIGHT(PV.numeradorinterno, 9) = 'Mostrador' THEN ISNULL(CNCC.saldo, PV.importeglobal) ELSE ISNULL(CP.SALDO2_IMPORTE, 0) " & _
                      "END AS Saldo, CASE WHEN RIGHT(PV.numeradorinterno, 9) = 'Mostrador' THEN 'Contado' ELSE 'Cta.Cte.' END AS tipo " & _
                      "FROM         v_ud_ezi_controlncc_pos AS CNCC RIGHT OUTER JOIN " & _
                      "ud_ezi_puntodeventa_encabezado AS PV WITH (readpast) ON CNCC.idfactura = PV.id LEFT OUTER JOIN " & _
                      "V_COMPROMISOPAGO_ AS CP ON PV.calipsoid = CP.TRORIGINANTE_ID LEFT OUTER JOIN " & _
                      "V_PERSONA_ AS P RIGHT OUTER JOIN " & _
                      "V_CLIENTE_ AS C ON P.ID = C.ENTEASOCIADO_ID ON PV.clienteid = C.ID " & _
                      "WHERE     (PV.numeradorinterno LIKE '%Factura de Venta%') AND (PV.tipodefacturacionid <> 'NN') AND (CP.NIVEL = 1 OR CP.NIVEL IS NULL) and " & _
                      "(pv.numerodefactura + 'n ' + pv.cliente + ' ' + P.CUIT + ' ' + pv.vendedor like '" & xbusqueda & "') AND (pv.sucursal = '" & xsuc & "') " & _
                      "ORDER BY Fecha DESC "
                      
                    
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
