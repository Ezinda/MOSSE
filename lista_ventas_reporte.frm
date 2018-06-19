VERSION 5.00
Object = "{F6125AB1-8AB1-11CE-A77F-08002B2F4E98}#2.0#0"; "MSRDC20.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{D18BBD1F-82BB-4385-BED3-E9D31A3E361E}#1.0#0"; "KewlButtonz.ocx"
Begin VB.Form lista_ventas_reporte 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reporte de Ventas"
   ClientHeight    =   3210
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5235
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3210
   ScaleWidth      =   5235
   Begin VB.Frame Frame1 
      Caption         =   "Filtro de Fechas"
      Height          =   3135
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4935
      Begin VB.CommandButton Command3 
         Caption         =   "Hasta Fecha:"
         Height          =   375
         Left            =   480
         TabIndex        =   4
         Top             =   1320
         Width           =   1335
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Desde Fecha:"
         Height          =   375
         Left            =   480
         TabIndex        =   3
         Top             =   720
         Width           =   1335
      End
      Begin MSComCtl2.DTPicker DesdeFecha 
         Height          =   375
         Left            =   1920
         TabIndex        =   1
         Top             =   720
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   97517569
         CurrentDate     =   42198
      End
      Begin MSComCtl2.DTPicker HastaFecha 
         Height          =   375
         Left            =   1920
         TabIndex        =   2
         Top             =   1320
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   97517569
         CurrentDate     =   42198
      End
      Begin KewlButtonz.KewlButtons salir 
         Height          =   495
         Left            =   3600
         TabIndex        =   5
         Top             =   2160
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
         MICON           =   "lista_ventas_reporte.frx":0000
         PICN            =   "lista_ventas_reporte.frx":001C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Crystal.CrystalReport CrystalReporte 
         Left            =   4320
         Top             =   240
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
      Begin MSRDC.MSRDC reporte 
         Height          =   375
         Left            =   2280
         Top             =   120
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
      Begin KewlButtonz.KewlButtons aceptar 
         Height          =   495
         Left            =   1560
         TabIndex        =   6
         Top             =   2160
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   873
         BTYPE           =   14
         TX              =   "&Aceptar"
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
         MICON           =   "lista_ventas_reporte.frx":0B66
         PICN            =   "lista_ventas_reporte.frx":0B82
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
   Begin MSAdodcLib.Adodc datcalipso 
      Height          =   330
      Left            =   10680
      Top             =   7560
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
   Begin MSAdodcLib.Adodc datparametros 
      Height          =   330
      Left            =   10680
      Top             =   6840
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
   Begin MSAdodcLib.Adodc datencabezado 
      Height          =   330
      Left            =   9120
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
   Begin MSAdodcLib.Adodc datcomp 
      Height          =   330
      Left            =   12000
      Top             =   7560
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
   Begin MSAdodcLib.Adodc datpresupuesto 
      Height          =   330
      Left            =   11880
      Top             =   6960
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
   Begin MSAdodcLib.Adodc datitems 
      Height          =   330
      Left            =   9000
      Top             =   7440
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
Attribute VB_Name = "lista_ventas_reporte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public cuentacont As String
Dim cuenta(99999) As Integer
Public xbusqueda As String



Private Sub Command4_Click()
On Error GoTo fuera
Dim crApp As New CRAXDRT.Application
Dim crReport As New CRAXDRT.Report
Dim tabla As String
Dim ruta As String


reporte.SQL = "SELECT v_ezi_pos_factctacte.id, v_ezi_pos_factctacte.NUMERODOCUMENTO, v_ezi_pos_factctacte.FECHAEMISION, v_ezi_pos_factctacte.cod_cliente, v_ezi_pos_factctacte.cliente, v_ezi_pos_factctacte.CUIT, v_ezi_pos_factctacte.CALLE, v_ezi_pos_factctacte.CODPOS, v_ezi_pos_factctacte.provincia, v_ezi_pos_factctacte.detalle, v_ezi_pos_factctacte.tipopago, v_ezi_pos_factctacte.codigoproducto, v_ezi_pos_factctacte.nombre_producto, v_ezi_pos_factctacte.cantidadproducto, v_ezi_pos_factctacte.nota, v_ezi_pos_factctacte.condiva, v_ezi_pos_factctacte.ciudad, v_ezi_pos_factctacte.TIPOVENTA, v_ezi_pos_factctacte.SIMBOLO, v_ezi_pos_factctacte.CODVENDEDOR, v_ezi_pos_factctacte.preciusiniva, v_ezi_pos_factctacte.subtotalsiniva, v_ezi_pos_factctacte.impbonifsiniva, v_ezi_pos_factctacte.nroremito, v_ezi_pos_factctacte.percepiibb, v_ezi_pos_factctacte.perceptem, v_ezi_pos_factctacte.totaltr, v_ezi_pos_factctacte.importeiva21, v_ezi_pos_factctacte.importeiva105, v_ezi_pos_factctacte.iditem,  " & _
              "v_ezi_pos_factctacte.cae, v_ezi_pos_factctacte.vto " & _
              "FROM  MMOSSE.DBO.v_ezi_pos_factctacte v_ezi_pos_factctacte " & _
              "where v_ezi_pos_factctacte.id = '" & DataGrid1.Columns(7).Text & "' order by v_ezi_pos_factctacte.iditem"

tabla = reporte.SQL

With CrystalReporte
    .PrinterCollation = crptCollated
   
    If DataGrid1.Columns(9).Text = "A" Then
      If tipofac <> "NN" Then
       If DataGrid1.Columns(11).Text = "N" Then
        .ReportFileName = App.Path & "\FacturaCtaCteA.rpt"
'         .ReportFileName = App.Path & "\PresupuestoA.rpt"
       Else
        .ReportFileName = App.Path & "\FacturaCtaCteA_alquiler.rpt"
'        .ReportFileName = App.Path & "\PresupuestoA.rpt"
       End If
      Else
        .ReportFileName = App.Path & "\PresupuestoA.rpt"
      End If
    Else
      If tipofac <> "NN" Then
       If DataGrid1.Columns(11).Text = "N" Then
        .ReportFileName = App.Path & "\FacturaCtaCteB.rpt"
'        .ReportFileName = App.Path & "\PresupuestoB.rpt"
       Else
        .ReportFileName = App.Path & "\FacturaCtaCteB_alquiler.rpt"
'        .ReportFileName = App.Path & "\PresupuestoB.rpt"
       End If
      Else
        .ReportFileName = App.Path & "\PresupuestoB.rpt"
      End If
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
End With
    
Exit Sub

fuera:
    
    MsgBox "Reporte de Factura no Encontado, o error de configuracion de reporte", vbCritical, "Error"


Exit Sub




End Sub


Private Sub aceptar_Click()
On Error GoTo fuera
Dim crApp As New CRAXDRT.Application
Dim crReport As New CRAXDRT.Report
Dim tabla As String
Dim ruta As String

 reporte.SQL = "SELECT     Numero, tipodefactura, fechadelcomprobante, numeradorinterno, codcliente, cliente, totaltr, NroRemito " & _
              "FROM         MMOSSE.dbo.v_ezi_pos_reporte_ventas AS v_ezi_pos_listadofacturas " & _
              "WHERE     convert(date,v_ezi_pos_listadofacturas.fechadelcomprobante) >= '" & DesdeFecha.Value & "' and " & _
              "convert(date,v_ezi_pos_listadofacturas.fechadelcomprobante) <= '" & HastaFecha.Value & "' and " & _
              "v_ezi_pos_listadofacturas.sucursal = '" & login.nomsucursal & "' " & _
              "order by fechadelcomprobante desc"

tabla = reporte.SQL

With CrystalReporte
    .PrinterCollation = crptCollated
    .ReportFileName = App.Path & "\Reporte_listado_ventas.rpt"
    .WindowTitle = "Listado de Facturas"
    .Formulas(0) = "desdefecha=""" & DesdeFecha.Value & """"
    .Formulas(1) = "hastafecha=""" & HastaFecha.Value & """"
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
End With
    
Exit Sub

fuera:
    
    MsgBox "Reporte no Encontado, o error de configuracion de reporte", vbCritical, "Error"



End Sub

Private Sub Command5_Click()
On Error Resume Next

xsuc = login.nomsucursal

If Text1.Text = "" Then
    xquery1 = "SELECT distinct ud_ezi_puntodeventa_encabezado.claveprimaria AS Numero, ud_ezi_puntodeventa_encabezado.numerodefactura AS Nro, " & _
                      "ud_ezi_puntodeventa_encabezado.fechadelcomprobante AS Fecha, ud_ezi_puntodeventa_encabezado.cliente AS Cliente, V_PERSONA_.CUIT, " & _
                      "ud_ezi_puntodeventa_encabezado.vendedor AS Vendedor, ud_ezi_puntodeventa_encabezado.importeglobal AS Importe, ud_ezi_puntodeventa_encabezado.id, " & _
                      "ud_ezi_puntodeventa_encabezado.generada, ud_ezi_puntodeventa_encabezado.tipodefactura, " & _
                      "ud_ezi_puntodeventa_encabezado.numerodefactura + 'n ' + ud_ezi_puntodeventa_encabezado.cliente + ' ' + V_PERSONA_.CUIT + ' ' + ud_ezi_puntodeventa_encabezado.vendedor " & _
                      "AS concatenado, ud_ezi_puntodeventa_encabezado.alquiler, ISNULL(V_COMPROMISOPAGO_.SALDO2_IMPORTE, ud_ezi_puntodeventa_encabezado.importeglobal) " & _
                      "AS SALDO, 0 AS NCIMPORTE, ud_ezi_puntodeventa_encabezado.nroorden as CAE, ud_ezi_puntodeventa_encabezado.recetaid as Autoriza, " & _
                      "case when ud_ezi_puntodeventa_encabezado.numeradorinterno like '%Mostrador%' then 'Contado' ELse 'Cta.Cte' end as Tipo " & _
                      "FROM         V_TRFACTURAVENTA_ RIGHT OUTER JOIN " & _
                      "ud_ezi_puntodeventa_encabezado WITH (readpast) ON V_TRFACTURAVENTA_.ID = ud_ezi_puntodeventa_encabezado.calipsoid LEFT OUTER JOIN " & _
                      "V_TRCREDITOVENTA_ ON ud_ezi_puntodeventa_encabezado.calipsoid = V_TRCREDITOVENTA_.VINCULOTR_ID LEFT OUTER JOIN " & _
                      "V_COMPROMISOPAGO_ ON ud_ezi_puntodeventa_encabezado.calipsoid = V_COMPROMISOPAGO_.TRORIGINANTE_ID LEFT OUTER JOIN " & _
                      "V_PERSONA_ RIGHT OUTER JOIN " & _
                      "V_CLIENTE_ ON V_PERSONA_.ID = V_CLIENTE_.ENTEASOCIADO_ID ON ud_ezi_puntodeventa_encabezado.clienteid = V_CLIENTE_.ID " & _
                      "WHERE     (ud_ezi_puntodeventa_encabezado.numeradorinterno LIKE '%Factura de Venta%') and ud_ezi_puntodeventa_encabezado.sucursal = '" & xsuc & "' AND (ud_ezi_puntodeventa_encabezado.tipodefacturacionid <> 'NN') AND (V_TRFACTURAVENTA_.FLAG_ID IS NULL) AND (V_COMPROMISOPAGO_.NIVEL = 1 OR " & _
                      "V_COMPROMISOPAGO_.NIVEL IS NULL) and " & _
                      "(ud_ezi_puntodeventa_encabezado.fechadelcomprobante >= '" & DateValue(DesdeFecha.Value) & "') and " & _
                      "(ud_ezi_puntodeventa_encabezado.fechadelcomprobante < '" & DateValue(HastaFecha.Value) + 1 & "') " & _
                      "ORDER BY Fecha DESC"
Else
            xquery1 = "SELECT   distinct  ud_ezi_puntodeventa_encabezado.claveprimaria AS Numero, ud_ezi_puntodeventa_encabezado.numerodefactura AS Nro, " & _
                      "ud_ezi_puntodeventa_encabezado.fechadelcomprobante AS Fecha, ud_ezi_puntodeventa_encabezado.cliente AS Cliente, V_PERSONA_.CUIT, " & _
                      "ud_ezi_puntodeventa_encabezado.vendedor AS Vendedor, ud_ezi_puntodeventa_encabezado.importeglobal AS Importe, ud_ezi_puntodeventa_encabezado.id, " & _
                      "ud_ezi_puntodeventa_encabezado.generada, ud_ezi_puntodeventa_encabezado.tipodefactura, " & _
                      "ud_ezi_puntodeventa_encabezado.numerodefactura + 'n ' + ud_ezi_puntodeventa_encabezado.cliente + ' ' + V_PERSONA_.CUIT + ' ' + ud_ezi_puntodeventa_encabezado.vendedor " & _
                      "AS concatenado, ud_ezi_puntodeventa_encabezado.alquiler, ISNULL(V_COMPROMISOPAGO_.SALDO2_IMPORTE, ud_ezi_puntodeventa_encabezado.importeglobal) " & _
                      "AS SALDO, 0 AS NCIMPORTE, ud_ezi_puntodeventa_encabezado.nroorden as CAE, ud_ezi_puntodeventa_encabezado.recetaid as Autoriza, " & _
                      "case when ud_ezi_puntodeventa_encabezado.numeradorinterno like '%Mostrador%' then 'Contado' ELse 'Cta.Cte' end as Tipo " & _
                      "FROM         V_TRFACTURAVENTA_ RIGHT OUTER JOIN " & _
                      "ud_ezi_puntodeventa_encabezado WITH (readpast) ON V_TRFACTURAVENTA_.ID = ud_ezi_puntodeventa_encabezado.calipsoid LEFT OUTER JOIN " & _
                      "V_TRCREDITOVENTA_ ON ud_ezi_puntodeventa_encabezado.calipsoid = V_TRCREDITOVENTA_.VINCULOTR_ID LEFT OUTER JOIN " & _
                      "V_COMPROMISOPAGO_ ON ud_ezi_puntodeventa_encabezado.calipsoid = V_COMPROMISOPAGO_.TRORIGINANTE_ID LEFT OUTER JOIN " & _
                      "V_PERSONA_ RIGHT OUTER JOIN " & _
                      "V_CLIENTE_ ON V_PERSONA_.ID = V_CLIENTE_.ENTEASOCIADO_ID ON ud_ezi_puntodeventa_encabezado.clienteid = V_CLIENTE_.ID " & _
                      "WHERE     (ud_ezi_puntodeventa_encabezado.numeradorinterno LIKE '%Factura de Venta%') AND (ud_ezi_puntodeventa_encabezado.tipodefacturacionid <> 'NN') AND ud_ezi_puntodeventa_encabezado.sucursal = '" & xsuc & "' and " & _
                      "(ud_ezi_puntodeventa_encabezado.numerodefactura + 'n ' + ud_ezi_puntodeventa_encabezado.cliente + ' ' + V_PERSONA_.CUIT + ' ' + ud_ezi_puntodeventa_encabezado.vendedor " & _
                      "LIKE '" & xbusqueda & "') AND (V_COMPROMISOPAGO_.NIVEL = 1 OR " & _
                      "V_COMPROMISOPAGO_.NIVEL IS NULL) AND (V_TRCREDITOVENTA_.FLAG_ID IS NULL) AND (V_TRFACTURAVENTA_.FLAG_ID IS NULL) and " & _
                      "(ud_ezi_puntodeventa_encabezado.fechadelcomprobante >= '" & DateValue(DesdeFecha.Value) & "') and " & _
                      "(ud_ezi_puntodeventa_encabezado.fechadelcomprobante < '" & DateValue(HastaFecha.Value) + 1 & "') " & _
                      "ORDER BY Fecha DESC"
End If

datpresupuesto.RecordSource = xquery1
datpresupuesto.Refresh

            DataGrid1.Columns(0).Visible = False
            DataGrid1.Columns(7).Visible = False
            DataGrid1.Columns(8).Visible = False
            DataGrid1.Columns(9).Visible = False
            DataGrid1.Columns(10).Visible = False
            DataGrid1.Columns(12).Visible = False
            DataGrid1.Columns(13).Visible = False
            DataGrid1.Columns("alquiler").Visible = False
            DataGrid1.Columns("cae").Visible = False
            DataGrid1.Columns(1).Width = 1000
            DataGrid1.Columns(3).Width = 3500
            DataGrid1.Columns(6).Alignment = dbgRight
            DataGrid1.Columns(6).NumberFormat = "Currency"
            
        
        

End Sub


Private Sub Command7_Click()

End Sub

Private Sub Form_Load()
Aplicar_skin Me

MiFuncionDeAjuste Me, True

yventana = Inicio.Height / 2 - 1000
xventana = Inicio.Width / 2

lista_ventas_reporte.Top = yventana - lista_ventas_reporte.Height / 2
lista_ventas_reporte.Left = xventana - lista_ventas_reporte.Width / 2

DesdeFecha.Value = Date - Day(Date) + 1
HastaFecha.Value = Date

datpresupuesto.ConnectionString = login.conexiontotal
datitems.ConnectionString = login.conexiontotal

xsuc = login.nomsucursal

                     

 
End Sub

Private Sub grabar_Click()

End Sub

Private Sub salir_Click()
    
    Unload Me

End Sub

