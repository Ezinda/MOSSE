VERSION 5.00
Object = "{F6125AB1-8AB1-11CE-A77F-08002B2F4E98}#2.0#0"; "MSRDC20.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form impordeneslistado 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de Ordenes de Pago"
   ClientHeight    =   2310
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   6210
   Icon            =   "impclientes.frx":0000
   LinkTopic       =   "From1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2310
   ScaleWidth      =   6210
   Begin VB.TextBox Text1 
      DataField       =   "cerrado"
      DataSource      =   "criterio"
      Height          =   285
      Index           =   8
      Left            =   600
      TabIndex        =   10
      Text            =   "Text1"
      Top             =   1560
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton listar 
      Caption         =   "&Listar"
      Height          =   375
      Left            =   2400
      TabIndex        =   9
      Top             =   1680
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      DataField       =   "hasta"
      DataSource      =   "criterio"
      Height          =   285
      Index           =   2
      Left            =   0
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   1560
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Text1 
      DataField       =   "desde"
      DataSource      =   "criterio"
      Height          =   285
      Index           =   1
      Left            =   600
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   1320
      Visible         =   0   'False
      Width           =   855
   End
   Begin MSComCtl2.DTPicker cargahasta 
      Height          =   375
      Left            =   3840
      TabIndex        =   3
      Top             =   360
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   12648447
      Format          =   65732609
      CurrentDate     =   38415
   End
   Begin MSComCtl2.DTPicker cargadesde 
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Top             =   360
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   12648447
      Format          =   65732609
      CurrentDate     =   38415
   End
   Begin VB.TextBox Text1 
      DataField       =   "empresa"
      DataSource      =   "criterio"
      Height          =   285
      Index           =   0
      Left            =   0
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   1320
      Visible         =   0   'False
      Width           =   615
   End
   Begin Crystal.CrystalReport CrystalReporte 
      Left            =   120
      Top             =   840
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTitle     =   "Listado Ordenes Emitidas"
      PrintFileLinesPerPage=   60
      WindowShowPrintSetupBtn=   -1  'True
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   255
      Left            =   480
      TabIndex        =   0
      Top             =   1080
      Visible         =   0   'False
      Width           =   975
   End
   Begin MSRDC.MSRDC reporte 
      Height          =   375
      Left            =   240
      Top             =   840
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
      Connect         =   "PROVIDER=MSDASQL;dsn=contable;uid=sa;pwd=;database=contablesql;"
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
   Begin VB.TextBox Text1 
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   3
      Left            =   3240
      TabIndex        =   6
      Text            =   "Hasta"
      Top             =   450
      Width           =   495
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   4
      Left            =   720
      TabIndex        =   7
      Text            =   "Desde"
      Top             =   450
      Width           =   615
   End
   Begin VB.Frame Frame1 
      Caption         =   "Periodo a Listar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1455
      Left            =   360
      TabIndex        =   8
      Top             =   0
      Width           =   5535
   End
End
Attribute VB_Name = "impordeneslistado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private crApp As New CRAXDRT.Application
Private crReport As New CRAXDRT.Report


Private Sub Command1_Click()
Dim tabla As String
Dim tabla1 As String

reporte.SQL = "consultaordesnpago.nrorden, consultaordesnpago.empresa, consultaordesnpago.nomproveedor, consultaordesnpago.comprobante, consultaordesnpago.fechacompro, consultaordesnpago.importe, consultaordesnpago.id, consultaordesnpago.razonsocial, consultaordesnpago.cuit, consultaordesnpago.domicilio, consultaordesnpago.localidad, consultaordesnpago.fecha, consultaordesnpago.domprov, consultaordesnpago.locprov, consultaordesnpago.cuitprov, consultaordesnpago.saldofactura FROM contablesql.dbo.consultaordesnpago consultaordesnpago WHERE consultaordesnpago.empresa = " & login.empresaact & "  and consultaordesnpago.fecha >= '" & cargadesde & "' and consultaordesnpago.fecha <= '" & cargahasta & "' "
tabla = reporte.SQL

With CrystalReporte
    .ReportFileName = App.Path & "\ListadoOrdenespago.rpt"
    .Connect = "PROVIDER=MSDASQL;dsn=contable;uid=lucva;pwd=25072004;database=contablesql;"
    .DiscardSavedData = True
    .RetrieveDataFiles
    .ReportSource = 0
    .SQLQuery = tabla
    .Destination = crptToWindow
 Rem   .Destination = crptToPrinter
    .PrintFileType = crptCrystal
    .WindowState = crptMaximized
    .Action = 1
     
End With
End Sub


Private Sub Form_Load()

Text1(0).Text = login.empresaact
cargadesde = Date - Day(Date) + 1
cargahasta = Date
Text1(1).Text = cargadesde
Text1(2).Text = cargahasta

End Sub


Private Sub listar_Click()

    Text1(1).Text = cargadesde.Value
    Text1(2).Text = cargahasta.Value

    Call Command1_Click

End Sub

