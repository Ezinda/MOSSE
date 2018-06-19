VERSION 5.00
Object = "{F6125AB1-8AB1-11CE-A77F-08002B2F4E98}#2.0#0"; "MSRDC20.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{D18BBD1F-82BB-4385-BED3-E9D31A3E361E}#1.0#0"; "KewlButtonz.ocx"
Begin VB.Form frmremitosdevolucion_alquiler 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Remitos de Alquiler pendientes para Carga de Devolucion del bien"
   ClientHeight    =   7440
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15195
   Icon            =   "frmremitosdevolucion_alquiler.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7440
   ScaleWidth      =   15195
   Begin VB.Frame Frame3 
      Caption         =   "Filtro"
      Height          =   855
      Left            =   8160
      TabIndex        =   11
      Top             =   120
      Width           =   1695
      Begin VB.OptionButton Option5 
         Caption         =   "Todos"
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   480
         Width           =   1215
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Pend.de dev."
         Height          =   195
         Left            =   120
         TabIndex        =   12
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
      TabIndex        =   8
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
      Bindings        =   "frmremitosdevolucion_alquiler.frx":0442
      Height          =   2775
      Left            =   120
      TabIndex        =   4
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
      TabIndex        =   7
      Top             =   120
      Width           =   7935
      Begin VB.OptionButton Option3 
         Caption         =   "Triplicado"
         Height          =   195
         Left            =   6360
         TabIndex        =   3
         Top             =   600
         Width           =   1335
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Duplicado"
         Height          =   195
         Left            =   6360
         TabIndex        =   2
         Top             =   360
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Original"
         Height          =   195
         Left            =   6360
         TabIndex        =   1
         Top             =   120
         Width           =   1335
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
         Width           =   5895
      End
   End
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   9960
      TabIndex        =   9
      Top             =   120
      Width           =   5055
      Begin KewlButtonz.KewlButtons Command4 
         Height          =   495
         Left            =   240
         TabIndex        =   5
         Top             =   240
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
         MICON           =   "frmremitosdevolucion_alquiler.frx":045B
         PICN            =   "frmremitosdevolucion_alquiler.frx":0477
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
         Left            =   2160
         TabIndex        =   6
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
         MICON           =   "frmremitosdevolucion_alquiler.frx":3869
         PICN            =   "frmremitosdevolucion_alquiler.frx":3885
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
      Bindings        =   "frmremitosdevolucion_alquiler.frx":43CF
      Height          =   2775
      Left            =   120
      TabIndex        =   10
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
      Left            =   2640
      TabIndex        =   14
      Top             =   6840
      Visible         =   0   'False
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
      MICON           =   "frmremitosdevolucion_alquiler.frx":43EB
      PICN            =   "frmremitosdevolucion_alquiler.frx":4407
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin KewlButtonz.KewlButtons KewlButtons1 
      Height          =   495
      Left            =   6240
      TabIndex        =   15
      Top             =   6840
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   873
      BTYPE           =   14
      TX              =   "&Devolucion de Bien"
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
      MICON           =   "frmremitosdevolucion_alquiler.frx":49A1
      PICN            =   "frmremitosdevolucion_alquiler.frx":49BD
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
Attribute VB_Name = "frmremitosdevolucion_alquiler"
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

    xquery = "SELECT  R.id, R.referenciaproducto AS Codigo, R.nombre_producto AS Descipcion, R.cantidadoriginal AS Cant_Orig, R.cantidadremitida AS Cant_Remitida, " & _
             "SUM(ISNULL(IFAC.cantidadproducto, 0)) AS Cant_Facturada, R.unidaddemedida AS Um " & _
             "FROM ud_ezi_puntodeventa_detalle_factm AS IFAC WITH (readpast) LEFT OUTER JOIN " & _
             "ud_ezi_puntodeventa_encabezado AS E WITH (readpast) ON IFAC.claveprimaria = E.id RIGHT OUTER JOIN " & _
             "v_ezi_pos_remito AS R ON IFAC.iditemremito = R.iditem " & _
             "GROUP BY R.id, R.referenciaproducto, R.nombre_producto, R.cantidadoriginal, R.cantidadremitida, R.unidaddemedida, R.iditem " & _
             "Having R.id = " & DataGrid1.Columns(0).Text & " ORDER BY R.iditem"


If datremitos.Recordset.EOF = False Then
        datitemremito.RecordSource = xquery
        datitemremito.Refresh
End If


            For X = 2 To 14
                DataGrid1.Columns(X).Locked = True
            Next X
            DataGrid1.Columns(1).Width = 300
            DataGrid1.Columns(4).Width = 1000
            DataGrid1.Columns(5).Width = 3500
            
            DataGrid1.Columns(0).Visible = False
            DataGrid1.Columns(9).Visible = False
            DataGrid1.Columns(10).Visible = False
            DataGrid1.Columns(12).Visible = False
            DataGrid1.Columns(13).Visible = False
            DataGrid1.Columns(14).Visible = False
            
            DataGrid2.Columns(0).Visible = False
            DataGrid2.Columns(2).Width = 6500
            DataGrid2.Columns(3).Alignment = dbgCenter
            DataGrid2.Columns(4).Visible = False
            DataGrid2.Columns(5).Alignment = dbgCenter


End Sub

Private Sub Command4_Click()
On Error Resume Next
Dim crApp As New CRAXDRT.Application
Dim crReport As New CRAXDRT.Report

Dim tabla As String
Dim ruta As String


reporte.SQL = "SELECT v_ezi_pos_remito.id, v_ezi_pos_remito.NUMERODOCUMENTO, v_ezi_pos_remito.FECHAEMISION, v_ezi_pos_remito.cod_cliente, v_ezi_pos_remito.cliente, v_ezi_pos_remito.CALLE, v_ezi_pos_remito.CODPOS, v_ezi_pos_remito.provincia, v_ezi_pos_remito.detalle, v_ezi_pos_remito.tipopago, v_ezi_pos_remito.referenciaproducto, v_ezi_pos_remito.nombre_producto, v_ezi_pos_remito.cantidadremitida, v_ezi_pos_remito.nota, v_ezi_pos_remito.condiva, v_ezi_pos_remito.ciudad, v_ezi_pos_remito.TIPOVENTA, v_ezi_pos_remito.SIMBOLO, v_ezi_pos_remito.iditem FROM COMERCIALCOLON.dbo.v_ezi_pos_remito v_ezi_pos_remito where v_ezi_pos_remito.id = " & DataGrid1.Columns(0).Text & " order by v_ezi_pos_remito.iditem"
tabla = reporte.SQL

With CrystalReporte
    .PrinterCollation = crptCollated
    .ReportFileName = App.Path & "\RemitoVtaAlquiler.rpt"
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
         "N.fechadelcomprobante , N.sucursal, N.obraid,N.estadoretira FROM  ud_ezi_puntodeventa_encabezado AS N LEFT OUTER JOIN v_ezi_pos_remito AS R ON N.id = R.presupuestobase " & _
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
              "N.fechadelcomprobante , N.sucursal, N.obraid, N.estadoretira FROM  ud_ezi_puntodeventa_encabezado AS N LEFT OUTER JOIN v_ezi_pos_remito AS R ON N.id = R.presupuestobase " & _
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

Debug.Print xquery

    query = xquery
    frmfacctacte_alquiler.Show
    If X = 0 Then
    '    frmfacctacte_venta.Text18.Text = DataGrid1.Columns(12).Text
'        frmfacctacte_venta.Text17.Text = DataGrid1.Columns(2).Text
    End If
    frmfacctacte_alquiler.Text17.SetFocus
    SendKeys "{ENTER}", False
    
    

End Sub

Private Sub Form_Activate()

    DataGrid1.SetFocus
    If menu = 1 Then
        If Text1.Text = "" Then Text1.Text = " "
        Text1.SetFocus
        SendKeys "{ENTER}", False
    End If


End Sub

Private Sub Form_Load()
On Error Resume Next
Aplicar_skin Me

yventana = Inicio.Height / 2 - 1000
xventana = Inicio.Width / 2

frmremitosdevolucion_alquiler.Top = yventana - frmremitosdevolucion_alquiler.Height / 2
frmremitosdevolucion_alquiler.Left = xventana - frmremitosdevolucion_alquiler.Width / 2


datremitos.ConnectionString = login.conexiontotal
datitemremito.ConnectionString = login.conexiontotal
xcon = 1
facturar.Visible = False

xquery1 = "SELECT     id, Sel, NroRemito, Fecha, CodCliente, Cliente, CUIT, Vendedor, Tipopago, TipodeVenta, presupuestobase, MAX(NroFactura) AS NroFactura, MAX(cantremitida) " & _
          "AS cantremitida, SUM(cantfacturada) AS cantfacturada, MIN(concatenado) AS concatenado " & _
          "FROM (SELECT     TOP (50) R.id, R.Sel, R.NUMERODOCUMENTO AS NroRemito, R.FECHAEMISION AS Fecha, R.cod_cliente AS CodCliente, R.cliente AS Cliente, R.CUIT, " & _
          "R.vendedor AS Vendedor, R.tipopago AS Tipopago, R.TIPOVENTA AS TipodeVenta, R.presupuestobase, " & _
          "E.tipodefactura + ' ' + E.puntodeventa + RIGHT('0000000' + E.numerodefactura, 8) AS NroFactura, max(R.cantidadremitida) AS cantremitida, " & _
          "SUM(ISNULL(IFAC.cantidadproducto, 0)) AS cantfacturada, R.NUMERODOCUMENTO + 'r  ' + R.cliente + ' ' + R.CUIT + ' ' + ISNULL(E.tipodefactura + ' ' + E.puntodeventa + RIGHT('0000000' + E.numerodefactura, 8), '') + 'f' AS concatenado " & _
          "FROM ud_ezi_puntodeventa_detalle_factm AS IFAC WITH (readpast) LEFT OUTER JOIN ud_ezi_puntodeventa_encabezado AS E WITH (readpast) ON IFAC.claveprimaria = E.id RIGHT OUTER JOIN " & _
          "v_ezi_pos_remito AS R ON IFAC.iditemremito = R.iditem " & _
          "GROUP BY R.id, R.Sel, R.NUMERODOCUMENTO, R.FECHAEMISION, R.cod_cliente, R.cliente, R.CUIT, R.vendedor, R.tipopago, R.TIPOVENTA, " & _
          "E.tipodefactura + ' ' + E.puntodeventa + RIGHT('0000000' + E.numerodefactura, 8), R.presupuestobase, R.alquiler HAVING      (R.alquiler ='S')" & _
          "ORDER BY R.id DESC) AS rem GROUP BY id, Sel, NroRemito, Fecha, CodCliente, Cliente, CUIT, Vendedor, Tipopago, TipodeVenta, presupuestobase ORDER BY id DESC"

datremitos.RecordSource = xquery1
datremitos.Refresh

If datremitos.Recordset.EOF = False Then
        datremitos.Recordset.MoveFirst
        If Option4.Value = False Then
             datitemremito.RecordSource = "SELECT R.id, R.referenciaproducto AS Codigo, R.nombre_producto AS Descipcion, R.cantidadoriginal AS Cant_Orig, R.cantidadremitida AS Cant_Remitida, " & _
                      "ISNULL(IFAC.cantidadproducto, 0) AS Cant_Facturada, R.unidaddemedida AS Um " & _
                      "FROM ud_ezi_puntodeventa_detalle_factm AS IFAC with (readpast) LEFT OUTER JOIN " & _
                      "ud_ezi_puntodeventa_encabezado AS E with (readpast) ON IFAC.claveprimaria = E.id RIGHT OUTER JOIN " & _
                      "v_ezi_pos_remito AS R ON IFAC.iditemremito = R.iditem " & _
                      "where R.id = " & DataGrid1.Columns(0).Text & " order by r.iditem "
        Else
             datitemremito.RecordSource = "SELECT R.id, R.referenciaproducto AS Codigo, R.nombre_producto AS Descipcion, R.cantidadoriginal AS Cant_Orig, R.cantidadremitida AS Cant_Remitida, " & _
                      "ISNULL(IFAC.cantidadproducto, 0) AS Cant_Facturada, R.unidaddemedida AS Um " & _
                      "FROM ud_ezi_puntodeventa_detalle_factm AS IFAC with (readpast) LEFT OUTER JOIN " & _
                      "ud_ezi_puntodeventa_encabezado AS E with (readpast) ON IFAC.claveprimaria = E.id RIGHT OUTER JOIN " & _
                      "v_ezi_pos_remito AS R ON IFAC.iditemremito = R.iditem " & _
                      "where R.id = " & DataGrid1.Columns(0).Text & " and (R.cantidadoriginal > ISNULL(IFAC.cantidadproducto, 0)) order by r.iditem "
        End If
        datitemremito.Refresh
End If

            For X = 2 To 14
                DataGrid1.Columns(X).Locked = True
            Next X
            DataGrid1.Columns(1).Width = 300
            DataGrid1.Columns(4).Width = 1000
            DataGrid1.Columns(5).Width = 3500
            
            DataGrid1.Columns(0).Visible = False
            DataGrid1.Columns(9).Visible = False
            DataGrid1.Columns(10).Visible = False
            DataGrid1.Columns(12).Visible = False
            DataGrid1.Columns(13).Visible = False
            DataGrid1.Columns(14).Visible = False
            
            DataGrid2.Columns(0).Visible = False
            DataGrid2.Columns(2).Width = 6800
            DataGrid2.Columns(4).Visible = False
            DataGrid2.Columns(3).Alignment = dbgCenter
'            DataGrid2.Columns(4).Alignment = dbgCenter
            DataGrid2.Columns(5).Alignment = dbgCenter


Option1.Value = True
Option5.Value = True

If menu = 1 Then
    Option5.Value = False
    Option4.Value = True
End If


End Sub

Private Sub KewlButtons1_Click()
On Error Resume Next

xquery = "SELECT   distinct  N.id, N.numeradorinterno, N.clienteid, N.cliente, N.vendedorid, N.vendedor, N.detalle, N.nota, N.mediodetransporteid, N.monedaid, N.cotizacion, N.listadeprecioid, " & _
         "N.tipodepagoid, N.tipodefacturacionid, N.tipodeentregaid, N.fechadeentrega, N.recargo, N.tiporecargo, N.bonificacion, N.tipobonificacion, N.importeglobal, " & _
         "N.numerodefactura, N.tipodefactura, N.numerador, N.domicilioid, N.domiciliodeentregaid, N.puntodeventa, N.descuento, N.subtotalsiniva, N.totaliva, N.generada, " & _
         "N.presupuestobase, N.importado, N.comprobanteorigen, N.totaltr, N.percepiibb, N.formapago, N.target, N.nombrepc, N.nrofactinicial, N.nrofactfinal, N.responsabilidad, " & _
         "N.nrofactoriginante, N.factoriginanteid, N.estado, N.transferido, N.calipsoid, N.osmpmutualid, N.recetaid, N.medicoid, N.adicionalid, N.reconocido, N.facturar, N.senia, " & _
         "N.nroorden, N.fechaorden, N.estadoimpresion, N.imprimir, N.perceptem, N.perceppyp, R.id AS idremito, R.NUMERODOCUMENTO, N.claveprimaria, " & _
         "N.fechadelcomprobante , N.sucursal, N.obraid, N.horaretiro, N.retira, N.estadoretira FROM  ud_ezi_puntodeventa_encabezado AS N LEFT OUTER JOIN v_ezi_pos_remito AS R ON N.id = R.presupuestobase " & _
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
              "N.fechadelcomprobante , N.sucursal, N.obraid, N.horaretiro, N.retira,N.estadoretira FROM  ud_ezi_puntodeventa_encabezado AS N LEFT OUTER JOIN v_ezi_pos_remito AS R ON N.id = R.presupuestobase " & _
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

Debug.Print xquery

    query = xquery
    frmalquiler_devolucion.Show
    If X = 0 Then
    '    frmfacctacte_venta.Text18.Text = DataGrid1.Columns(12).Text
'        frmfacctacte_venta.Text17.Text = DataGrid1.Columns(2).Text
    End If
    frmalquiler_devolucion.Text17.SetFocus
    SendKeys "{ENTER}", False
    
    


    

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
            
                xquery1 = "SELECT id, Sel, NroRemito, Fecha, CodCliente, Cliente, CUIT,MAX(NroFactura) AS NroFactura , Tipopago, TipodeVenta, presupuestobase, Vendedor, MAX(cantremitida) " & _
                          "AS cantremitida, SUM(cantfacturada) AS cantfacturada, MAX(concatenado) AS concatenado  " & _
                          "FROM (SELECT     TOP (50) R.id, R.Sel, R.NUMERODOCUMENTO AS NroRemito, R.FECHAEMISION AS Fecha, R.cod_cliente AS CodCliente, R.cliente AS Cliente, R.CUIT,  " & _
                          "R.vendedor AS Vendedor, R.tipopago AS Tipopago, R.TIPOVENTA AS TipodeVenta, R.presupuestobase,  " & _
                          "E.tipodefactura + ' ' + E.puntodeventa + RIGHT('0000000' + E.numerodefactura, 8) AS NroFactura, MAX(R.cantidadremitida) AS cantremitida,  " & _
                          "SUM(ISNULL(IFAC.cantidadproducto, 0)) AS cantfacturada,  " & _
                          "R.NUMERODOCUMENTO + 'r  ' + R.cliente + ' ' + R.CUIT + ' ' + ISNULL(E.tipodefactura + ' ' + E.puntodeventa + RIGHT('0000000' + E.numerodefactura, 8), '') + 'f' AS concatenado  " & _
                          "FROM         ud_ezi_puntodeventa_detalle_factm AS IFAC WITH (readpast) LEFT OUTER JOIN  " & _
                          "ud_ezi_puntodeventa_encabezado AS E WITH (readpast) ON IFAC.claveprimaria = E.id RIGHT OUTER JOIN  " & _
                          "v_ezi_pos_remito AS R ON IFAC.iditemremito = R.iditem  " & _
                          "GROUP BY R.id, R.Sel, R.NUMERODOCUMENTO, R.FECHAEMISION, R.cod_cliente, R.cliente, R.CUIT, R.vendedor, R.tipopago, R.TIPOVENTA, " & _
                          "E.tipodefactura + ' ' + E.puntodeventa + RIGHT('0000000' + E.numerodefactura, 8), R.presupuestobase, R.alquiler " & _
                          "HAVING      (R.alquiler = 'S') AND " & _
                          "(R.NUMERODOCUMENTO + 'r  ' + R.cliente + ' ' + R.CUIT + ' ' + ISNULL(E.tipodefactura + ' ' + E.puntodeventa + RIGHT('0000000' + E.numerodefactura, 8), N'') " & _
                          "+ 'f' LIKE '" & xbusqueda & "') AND (NOT EXISTS ( " & _
                          "(SELECT     comprobanteorigen " & _
                          "From ud_ezi_puntodeventa_encabezado " & _
                          "WHERE      (numeradorinterno = 'Remito de Devolucion') AND R.id = comprobanteorigen)))) AS RC " & _
                          "GROUP BY id, Sel, NroRemito, Fecha, CodCliente, Cliente, CUIT, Vendedor, Tipopago, TipodeVenta, presupuestobase " & _
                          "ORDER BY id DESC "

            Else
                xquery1 = "SELECT id, Sel, NroRemito, Fecha, CodCliente, Cliente, CUIT,MAX(NroFactura) AS NroFactura , Tipopago, TipodeVenta, presupuestobase,Vendedor , MAX(cantremitida) " & _
                          "AS cantremitida, SUM(cantfacturada) AS cantfacturada, MAX(concatenado) AS concatenado  " & _
                          "FROM (SELECT     TOP (50) R.id, R.Sel, R.NUMERODOCUMENTO AS NroRemito, R.FECHAEMISION AS Fecha, R.cod_cliente AS CodCliente, R.cliente AS Cliente, R.CUIT,  " & _
                          "R.vendedor AS Vendedor, R.tipopago AS Tipopago, R.TIPOVENTA AS TipodeVenta, R.presupuestobase,  " & _
                          "E.tipodefactura + ' ' + E.puntodeventa + RIGHT('0000000' + E.numerodefactura, 8) AS NroFactura, MAX(R.cantidadremitida) AS cantremitida,  " & _
                          "SUM(ISNULL(IFAC.cantidadproducto, 0)) AS cantfacturada,  " & _
                          "R.NUMERODOCUMENTO + 'r  ' + R.cliente + ' ' + R.CUIT + ' ' + ISNULL(E.tipodefactura + ' ' + E.puntodeventa + RIGHT('0000000' + E.numerodefactura, 8), '') + 'f' AS concatenado  " & _
                          "FROM         ud_ezi_puntodeventa_detalle_factm AS IFAC WITH (readpast) LEFT OUTER JOIN  " & _
                          "ud_ezi_puntodeventa_encabezado AS E WITH (readpast) ON IFAC.claveprimaria = E.id RIGHT OUTER JOIN  " & _
                          "v_ezi_pos_remito AS R ON IFAC.iditemremito = R.iditem  " & _
                          "GROUP BY R.id, R.Sel, R.NUMERODOCUMENTO, R.FECHAEMISION, R.cod_cliente, R.cliente, R.CUIT, R.vendedor, R.tipopago, R.TIPOVENTA, " & _
                          "E.tipodefactura + ' ' + E.puntodeventa + RIGHT('0000000' + E.numerodefactura, 8), R.presupuestobase, R.alquiler " & _
                          "HAVING      (R.alquiler = 'S') AND " & _
                          "(R.NUMERODOCUMENTO + 'r  ' + R.cliente + ' ' + R.CUIT + ' ' + ISNULL(E.tipodefactura + ' ' + E.puntodeventa + RIGHT('0000000' + E.numerodefactura, 8), N'') " & _
                          "+ 'f' LIKE '" & xbusqueda & "')) AS RC " & _
                          "GROUP BY id, Sel, NroRemito, Fecha, CodCliente, Cliente, CUIT, Vendedor, Tipopago, TipodeVenta, presupuestobase " & _
                          "ORDER BY id DESC "

            End If
                    
            datremitos.RecordSource = xquery1
            datremitos.Refresh
            If datremitos.Recordset.RecordCount = 0 Then
                DataGrid2.Visible = False
            Else
                DataGrid2.Visible = True
            End If
            
            For X = 2 To 14
                DataGrid1.Columns(X).Locked = True
            Next X
            DataGrid1.Columns(1).Width = 300
            DataGrid1.Columns(4).Width = 1000
            DataGrid1.Columns(5).Width = 3500
            
            DataGrid1.Columns(0).Visible = False
            DataGrid1.Columns(9).Visible = False
            DataGrid1.Columns(10).Visible = False
            DataGrid1.Columns(12).Visible = False
            DataGrid1.Columns(13).Visible = False
            DataGrid1.Columns(14).Visible = False
            
            Call DataGrid1_Click
            
        End If
        DataGrid1.SetFocus
        
        
    End If


End Sub
