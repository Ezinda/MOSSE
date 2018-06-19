VERSION 5.00
Object = "{F6125AB1-8AB1-11CE-A77F-08002B2F4E98}#2.0#0"; "MSRDC20.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form imprecibos 
   Caption         =   "Impresion Recibos"
   ClientHeight    =   735
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3210
   LinkTopic       =   "Form1"
   ScaleHeight     =   735
   ScaleWidth      =   3210
   StartUpPosition =   3  'Windows Default
   WindowState     =   1  'Minimized
   Begin Crystal.CrystalReport CrystalReporte 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTitle     =   "Orden de Pago"
      PrintFileLinesPerPage=   60
   End
   Begin MSRDC.MSRDC reporte 
      Height          =   375
      Left            =   360
      Top             =   0
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
   Begin MSAdodcLib.Adodc criterio 
      Height          =   330
      Left            =   1560
      Top             =   0
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
End
Attribute VB_Name = "imprecibos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

End Sub

Private Sub Form_Load()

On Error GoTo errorimp
Dim crApp As New CRAXDRT.Application
Dim crReport As New CRAXDRT.Report
Dim tabla As String
Dim tabla1 As String
Dim ruta As String

parametros = Clipboard.GetText
empresaact = Val(Mid(parametros, 1, 10))
conexiontotal = Mid(parametros, 11, 255)
conexionreporte = Mid(parametros, 266, 255)
numorden = Right(parametros, Len(parametros) - 520)

criterio.ConnectionString = conexiontotal

criterio.RecordSource = "select empreactiva.* from empreactiva"
criterio.Refresh

criterio.Recordset.Fields(0) = empresaact
criterio.Recordset.UpdateBatch adAffectCurrent

ruta = "\Empresa" + Right(Str(empresaact), Len(Str(empresaact)) - 1)

reporte.SQL = "consultarecibocobro.nrorden, consultarecibocobro.empresa, consultarecibocobro.nomproveedor, consultarecibocobro.comprobante, consultarecibocobro.fechacompro, consultarecibocobro.importe, consultarecibocobro.id, consultarecibocobro.razonsocial, consultarecibocobro.cuit, consultarecibocobro.domicilio, consultarecibocobro.localidad, consultarecibocobro.fecha, consultarecibocobro.domprov, consultarecibocobro.locprov, consultarecibocobro.cuitprov, consultarecibocobro.saldofactura FROM contablesql.dbo.consultarecibocobro consultarecibocobro WHERE consultarecibocobro.nrorden= '" & numorden & "' and consultarecibocobro.empresa = " & empresaact & " ORDER BY consultarecibocobro.razonsocial ASC, consultarecibocobro.id ASC"
tabla = reporte.SQL

With CrystalReporte
  
    .ReportFileName = App.Path & ruta + "\Recibocliente.rpt"
    .Connect = conexionreporte
 For X = 0 To 1
    .SubreportToChange = .GetNthSubreportName(X)
    .Connect = conexionreporte
    .SubreportToChange = ""
    .Connect = conexionreporte
 Next X
    .DiscardSavedData = True
    .RetrieveDataFiles
    .ReportSource = 0
    .SQLQuery = tabla
    .Destination = crptToPrinter
    .PrintFileType = crptCrystal
    .Action = 1
                      
End With

errorimp:
Unload Me

End Sub
