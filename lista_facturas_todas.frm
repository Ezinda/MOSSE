VERSION 5.00
Object = "{F6125AB1-8AB1-11CE-A77F-08002B2F4E98}#2.0#0"; "MSRDC20.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{D18BBD1F-82BB-4385-BED3-E9D31A3E361E}#1.0#0"; "KewlButtonz.ocx"
Begin VB.Form lista_facturas_todas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta de Facturas"
   ClientHeight    =   8760
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15195
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8760
   ScaleWidth      =   15195
   Begin VB.Frame Frame1 
      Caption         =   "Filtro de Comprobantes"
      Height          =   1455
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   15015
      Begin VB.CommandButton FACTURAELECTRONICA 
         Caption         =   "Fac.Electronica"
         Height          =   375
         Left            =   12120
         TabIndex        =   16
         Top             =   840
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Original"
         Height          =   255
         Left            =   5640
         TabIndex        =   15
         Top             =   1080
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Duplicado"
         Height          =   255
         Left            =   6480
         TabIndex        =   14
         Top             =   1080
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Buscar:"
         Height          =   375
         Left            =   240
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   720
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
         Left            =   1560
         TabIndex        =   12
         Top             =   720
         Width           =   6015
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Filtrar"
         Height          =   375
         Left            =   6240
         TabIndex        =   9
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Hasta Fecha:"
         Height          =   375
         Left            =   3240
         TabIndex        =   8
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Imprimir Listado"
         Height          =   375
         Left            =   7680
         TabIndex        =   7
         Top             =   720
         Width           =   1335
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Generar PDF"
         Height          =   375
         Left            =   9120
         TabIndex        =   6
         Top             =   240
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Desde Fecha:"
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   240
         Width           =   1335
      End
      Begin MSComCtl2.DTPicker DesdeFecha 
         Height          =   375
         Left            =   1560
         TabIndex        =   3
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   49610753
         CurrentDate     =   42198
      End
      Begin MSComCtl2.DTPicker HastaFecha 
         Height          =   375
         Left            =   4560
         TabIndex        =   4
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   49610753
         CurrentDate     =   42198
      End
      Begin KewlButtonz.KewlButtons salir 
         Height          =   495
         Left            =   12360
         TabIndex        =   10
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
         MICON           =   "lista_facturas_todas.frx":0000
         PICN            =   "lista_facturas_todas.frx":001C
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
         Left            =   10320
         TabIndex        =   11
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   873
         BTYPE           =   14
         TX              =   "&Previsualizar Comprobante"
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
         MICON           =   "lista_facturas_todas.frx":0B66
         PICN            =   "lista_facturas_todas.frx":0B82
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
         Left            =   8040
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
         WindowShowPrintSetupBtn=   -1  'True
      End
      Begin MSRDC.MSRDC reporte 
         Height          =   375
         Left            =   0
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
      Begin KewlButtonz.KewlButtons KewlButtons1 
         Height          =   495
         Left            =   10320
         TabIndex        =   17
         Top             =   840
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   873
         BTYPE           =   14
         TX              =   "Enviar por Mail"
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
         MICON           =   "lista_facturas_todas.frx":3F74
         PICN            =   "lista_facturas_todas.frx":3F90
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "lista_facturas_todas.frx":43E2
      Height          =   4695
      Left            =   120
      TabIndex        =   0
      Top             =   1560
      Width           =   15015
      _ExtentX        =   26485
      _ExtentY        =   8281
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
   Begin MSDataGridLib.DataGrid DataGrid2 
      Bindings        =   "lista_facturas_todas.frx":43FF
      Height          =   2175
      Left            =   120
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   6360
      Width           =   15015
      _ExtentX        =   26485
      _ExtentY        =   3836
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
Attribute VB_Name = "lista_facturas_todas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public cuentacont As String
Dim cuenta(99999) As Integer
Public xbusqueda As String

Dim oApp As Outlook.Application
Dim myItem As Outlook.MailItem
Dim myAttachments As Outlook.Attachments



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
    .Formulas(0) = "copia="" ORIGINAL """
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
            DataGrid1.Columns("cae").Visible = True
            DataGrid1.Columns(1).Width = 1000
            DataGrid1.Columns(3).Width = 3500
            DataGrid1.Columns(6).Alignment = dbgRight
            DataGrid1.Columns(6).NumberFormat = "Currency"
            
        
        

End Sub

Private Sub Command6_Click()
On Error GoTo fuera
Dim crApp As New CRAXDRT.Application
Dim crReport As New CRAXDRT.Report
Dim tabla As String
Dim ruta As String

datpresupuesto.Recordset.MoveFirst

Do While Not datpresupuesto.Recordset.EOF

reporte.SQL = "SELECT v_ezi_pos_factctacte.id, v_ezi_pos_factctacte.NUMERODOCUMENTO, v_ezi_pos_factctacte.FECHAEMISION, v_ezi_pos_factctacte.cod_cliente, v_ezi_pos_factctacte.cliente, v_ezi_pos_factctacte.CUIT, v_ezi_pos_factctacte.CALLE, v_ezi_pos_factctacte.CODPOS, v_ezi_pos_factctacte.provincia, v_ezi_pos_factctacte.detalle, v_ezi_pos_factctacte.tipopago, v_ezi_pos_factctacte.codigoproducto, v_ezi_pos_factctacte.nombre_producto, v_ezi_pos_factctacte.cantidadproducto, v_ezi_pos_factctacte.nota, v_ezi_pos_factctacte.condiva, v_ezi_pos_factctacte.ciudad, v_ezi_pos_factctacte.TIPOVENTA, v_ezi_pos_factctacte.SIMBOLO, v_ezi_pos_factctacte.CODVENDEDOR, v_ezi_pos_factctacte.preciusiniva, v_ezi_pos_factctacte.subtotalsiniva, v_ezi_pos_factctacte.impbonifsiniva, v_ezi_pos_factctacte.nroremito, v_ezi_pos_factctacte.percepiibb, v_ezi_pos_factctacte.perceptem, v_ezi_pos_factctacte.totaltr, v_ezi_pos_factctacte.importeiva21, v_ezi_pos_factctacte.importeiva105, v_ezi_pos_factctacte.iditem,  " & _
              "v_ezi_pos_factctacte.cae, v_ezi_pos_factctacte.vto " & _
              "FROM  MMOSSE.DBO.v_ezi_pos_factctacte v_ezi_pos_factctacte " & _
              "where v_ezi_pos_factctacte.id = '" & DataGrid1.Columns(7).Text & "' order by v_ezi_pos_factctacte.iditem"

tabla = reporte.SQL

xnumerofac = "Fac " + datpresupuesto.Recordset.Fields("tipodefactura") + " 0006" + "-" + Right("00000000" + datpresupuesto.Recordset.Fields("Nro"), 8) + ".pdf"


With CrystalReporte
    .PrinterCollation = crptCollated
   
    If tipofac <> "NN" Then
        If Option1.Value = True Then
            .Formulas(0) = "copia="" ORIGINAL """
        Else
            .Formulas(0) = "copia="" DUPLICADO """
        End If
        
    End If
    If DataGrid1.Columns(9).Text = "A" Then
      If tipofac <> "NN" Then
       If DataGrid1.Columns(11).Text = "N" Then
        .ReportFileName = App.Path & "\FacturaCtaCteA.rpt"
       Else
        .ReportFileName = App.Path & "\FacturaCtaCteA_alquiler.rpt"
       End If
      Else
        .ReportFileName = App.Path & "\PresupuestoA.rpt"
      End If
    Else
      If tipofac <> "NN" Then
       If DataGrid1.Columns(11).Text = "N" Then
        .ReportFileName = App.Path & "\FacturaCtaCteB.rpt"
       Else
        .ReportFileName = App.Path & "\FacturaCtaCteB_alquiler.rpt"
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
 '   .Destination = crptToWindow
 '   .Destination = crptToPrinter
    .Destination = crptToFile
    
    .PrintFileType = crptRTF
    
    
    .PrintFileName = "c:\fe\" + xnumerofac
    
    .WindowState = crptMaximized
    .Action = 1
End With
    
datpresupuesto.Recordset.MoveNext

Loop

    
Exit Sub

fuera:
    
    MsgBox "Reporte de Factura no Encontado, o error de configuracion de reporte", vbCritical, "Error"


Exit Sub


End Sub

Private Sub Command7_Click()
On Error GoTo fuera
Dim crApp As New CRAXDRT.Application
Dim crReport As New CRAXDRT.Report
Dim tabla As String
Dim ruta As String

If Text1.Text = "" Then
 reporte.SQL = "SELECT     Numero, tipodefactura, fechadelcomprobante, numeradorinterno, codcliente, cliente, totaltr, NroRemito " & _
              "FROM         MMOSSE.dbo.v_ezi_pos_listadofacturas AS v_ezi_pos_listadofacturas " & _
              "WHERE     convert(date,v_ezi_pos_listadofacturas.fechadelcomprobante) >= '" & DesdeFecha.Value & "' and " & _
              "convert(date,v_ezi_pos_listadofacturas.fechadelcomprobante) <= '" & HastaFecha.Value & "' and " & _
              "v_ezi_pos_listadofacturas.sucursal = '" & login.nomsucursal & "' " & _
              "order by fechadelcomprobante desc"
Else
 reporte.SQL = "SELECT     Numero, tipodefactura, fechadelcomprobante, numeradorinterno, codcliente, cliente, totaltr, NroRemito " & _
              "FROM         MMOSSE.dbo.v_ezi_pos_listadofacturas AS v_ezi_pos_listadofacturas " & _
              "WHERE     convert(date,v_ezi_pos_listadofacturas.fechadelcomprobante) >= '" & DesdeFecha.Value & "' and " & _
              "convert(date,v_ezi_pos_listadofacturas.fechadelcomprobante) <= '" & HastaFecha.Value & "' and " & _
              "v_ezi_pos_listadofacturas.sucursal = '" & login.nomsucursal & "' and " & _
              "v_ezi_pos_listadofacturas.cliente = '" & DataGrid1.Columns("cliente").Text & "' " & _
              "order by fechadelcomprobante desc"
End If

tabla = reporte.SQL

With CrystalReporte
    .PrinterCollation = crptCollated
    .ReportFileName = App.Path & "\Reporte_listado_facturas.rpt"
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


Private Sub DataGrid1_Click()

    xidencabezado = DataGrid1.Columns(7).Text
    datitems.RecordSource = "select codigoproducto as Codigo, nombre_producto as Descripcion, cantidadproducto as Cantidad, unidaddemedidaid as Um, preciou as Precio, subtotal as Subtotal from ud_ezi_puntodeventa_detalle_factm with (readpast) where claveprimaria = " & xidencabezado & ""
    datitems.Refresh
            DataGrid2.Columns(1).Width = 5500
            DataGrid2.Columns(2).Alignment = dbgCenter
            DataGrid2.Columns(3).Alignment = dbgCenter
            DataGrid2.Columns(4).Alignment = dbgRight
            DataGrid2.Columns(5).Alignment = dbgRight
            DataGrid2.Columns(4).NumberFormat = "Currency"
            DataGrid2.Columns(5).NumberFormat = "Currency"
            


End Sub



Private Sub DataGrid1_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        Call Command4_Click
            
    End If

End Sub




Private Sub DataGrid1_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next

    xidencabezado = DataGrid1.Columns(7).Text
    datitems.RecordSource = "select codigoproducto as Codigo, nombre_producto as Descripcion, cantidadproducto as Cantidad, unidaddemedidaid as Um, preciou as Precio, subtotal as Subtotal from ud_ezi_puntodeventa_detalle_factm with (readpast) where claveprimaria = " & xidencabezado & ""
    datitems.Refresh
            DataGrid2.Columns(1).Width = 5500
            DataGrid2.Columns(2).Alignment = dbgCenter
            DataGrid2.Columns(3).Alignment = dbgCenter
            DataGrid2.Columns(4).Alignment = dbgRight
            DataGrid2.Columns(5).Alignment = dbgRight
            DataGrid2.Columns(4).NumberFormat = "Currency"
            DataGrid2.Columns(5).NumberFormat = "Currency"
            
End Sub

Private Sub FACTURAELECTRONICA_Click()
Dim fe As New WSAFIPFE.factura
On Error Resume Next

datcomp.ConnectionString = login.conexiontotal
datencabezado.ConnectionString = login.conexiontotal
datparametros.ConnectionString = login.conexiontotal

datparametros.RecordSource = "select * from ud_ezi_parametros_pos where sucursal = '" & login.nomsucursal & "' "
datparametros.Refresh


datcomp.RecordSource = "SELECT  ud_ezi_puntodeventa_encabezado.claveprimaria AS Numero, ud_ezi_puntodeventa_encabezado.numerodefactura AS Nro, " & _
                       "ud_ezi_puntodeventa_encabezado.fechadelcomprobante AS Fecha, ud_ezi_puntodeventa_encabezado.cliente AS Cliente, V_PERSONA_.CUIT,  " & _
                       "ud_ezi_puntodeventa_encabezado.vendedor AS Vendedor, ud_ezi_puntodeventa_encabezado.importeglobal AS Importe, ud_ezi_puntodeventa_encabezado.id,  " & _
                       "ud_ezi_puntodeventa_encabezado.generada, ud_ezi_puntodeventa_encabezado.tipodefactura, " & _
                       "ud_ezi_puntodeventa_encabezado.numerodefactura + 'n ' + ud_ezi_puntodeventa_encabezado.cliente + ' ' + V_PERSONA_.CUIT + ' ' + ud_ezi_puntodeventa_encabezado.vendedor " & _
                       "AS concatenado, ud_ezi_puntodeventa_encabezado.alquiler, ISNULL(V_COMPROMISOPAGO_.SALDO2_IMPORTE, ud_ezi_puntodeventa_encabezado.importeglobal) " & _
                       "AS SALDO, ISNULL(V_TRCREDITOVENTA_.VALORTOTAL, 0) AS NCIMPORTE, ud_ezi_puntodeventa_encabezado.subtotalsiniva, " & _
                       "ud_ezi_puntodeventa_encabezado.totaliva, ud_ezi_puntodeventa_encabezado.responsabilidad, round(v_ezi_pos_iva_facctacte.importeiva21,2) as  importeiva21, " & _
                       "round(v_ezi_pos_iva_facctacte.importeiva105,2) as importeiva105  , isnull(ud_ezi_puntodeventa_encabezado.nroorden,'') as nroorden, ud_ezi_puntodeventa_encabezado.fechaorden, " & _
                       "ud_ezi_puntodeventa_encabezado.percepiibb , ud_ezi_puntodeventa_encabezado.perceptem " & _
                       "FROM         v_ezi_pos_iva_facctacte RIGHT OUTER JOIN " & _
                       "ud_ezi_puntodeventa_encabezado WITH (readpast) ON v_ezi_pos_iva_facctacte.id = ud_ezi_puntodeventa_encabezado.id LEFT OUTER JOIN " & _
                       "V_TRCREDITOVENTA_ ON ud_ezi_puntodeventa_encabezado.calipsoid = V_TRCREDITOVENTA_.VINCULOTR_ID LEFT OUTER JOIN " & _
                       "V_COMPROMISOPAGO_ ON ud_ezi_puntodeventa_encabezado.calipsoid = V_COMPROMISOPAGO_.TRORIGINANTE_ID LEFT OUTER JOIN " & _
                       "V_PERSONA_ RIGHT OUTER JOIN " & _
                       "V_CLIENTE_ ON V_PERSONA_.ID = V_CLIENTE_.ENTEASOCIADO_ID ON ud_ezi_puntodeventa_encabezado.clienteid = V_CLIENTE_.ID " & _
                       "WHERE     (ud_ezi_puntodeventa_encabezado.numeradorinterno LIKE '%Factura de Venta%') AND (ud_ezi_puntodeventa_encabezado.tipodefacturacionid <> 'NN') AND " & _
                       "(ud_ezi_puntodeventa_encabezado.claveprimaria = '" & DataGrid1.Columns(7).Text & "') " & _
                       "ORDER BY Fecha DESC"
datcomp.Refresh

If datcomp.Recordset.Fields("nroorden") <> "" Then
    MsgBox "Esta factura ya se encuentra Validada por AFIP", vbCritical, "Error"
    Exit Sub
End If

datencabezado.RecordSource = "select * from ud_ezi_puntodeventa_encabezado where id ='" & DataGrid1.Columns(7).Text & "' "
datencabezado.Refresh

xactivar = fe.ActivarLicencia("20102028245", "WSAFIPFE.lic", "servcomsrl@gmail.com", "")


If fe.iniciar(modoFiscal_Fiscal, "20102028245", "mmosse.pfx", "WSAFIPFE.lic") Then
'If fe.iniciar(modoFiscal_Test, "20102028245", "mmosse_test.pfx", "") Then
   If fe.f1ObtenerTicketAcceso() Then
   
        PtoVta = datparametros.Recordset.Fields("ptovtaFE")
        
        If datcomp.Recordset.Fields("tipodefactura") = "A" Then
            TipoComp = 1 ' Factura A(Ver excel referencias codigos AFIP)
        Else
            TipoComp = 6 ' Factura B(Ver excel referencias codigos AFIP)
        End If
  
        xitemiva = 0
        xitemper = 0
        xcuit = datcomp.Recordset.Fields("cuit")
        xtotal = datcomp.Recordset.Fields("importe")
        xneto = datcomp.Recordset.Fields("subtotalsiniva")
        xtotaliva = datcomp.Recordset.Fields("totaliva")
        xtotaltrib = datcomp.Recordset.Fields("percepiibb") + datcomp.Recordset.Fields("perceptem")
        If Round(datcomp.Recordset.Fields("importeiva21"), 2) <> 0 Then xitemiva = xitemiva + 1
        If Round(datcomp.Recordset.Fields("importeiva105"), 2) <> 0 Then xitemiva = xitemiva + 1
        If datcomp.Recordset.Fields("percepiibb") <> 0 Then xitemper = xitemper + 1
        If datcomp.Recordset.Fields("perceptem") <> 0 Then xitemper = xitemper + 1
                
                
        If UCase(datcomp.Recordset.Fields("responsabilidad")) = "CONSUMIDOR FINAL" Then
                xdoctipo = 96
                xcuit = "11111111111"
        Else
                xdoctipo = 80
        End If
      FechaComp = Format(Now(), "yyyymmdd")
'      FechaComp = Format("30/01/2016", "yyyymmdd")

   
      fe.F1CabeceraCantReg = 1
      fe.F1CabeceraPtoVta = PtoVta
      fe.F1CabeceraCbteTipo = TipoComp

      fe.f1Indice = 0
      If datcomp.Recordset.Fields("alquiler") = "N" Then
        fe.F1DetalleConcepto = 1  '1 = producto , 2 = serviciop
      Else
        fe.F1DetalleConcepto = 2  '1 = producto , 2 = serviciop
      End If
      
      fe.F1DetalleDocTipo = xdoctipo
      
      nro = fe.F1CompUltimoAutorizado(PtoVta, TipoComp) + 1
      
      fe.F1DetalleDocNro = xcuit
      fe.F1DetalleCbteDesde = nro
      fe.F1DetalleCbteHasta = nro
      fe.F1DetalleCbteFch = FechaComp
      fe.F1DetalleImpTotal = xtotal
      fe.F1DetalleImpTotalConc = 0
      fe.F1DetalleImpNeto = xneto
      fe.F1DetalleImpOpEx = 0
      fe.F1DetalleImpTrib = Round(xtotaltrib, 2)
      fe.F1DetalleImpIva = Round(xtotaliva, 2)
      If datcomp.Recordset.Fields("alquiler") = "S" Then
         fe.F1DetalleFchServDesde = FechaComp
         fe.F1DetalleFchServHasta = FechaComp
         fe.F1DetalleFchVtoPago = FechaComp
      End If
      fe.F1DetalleMonId = "PES"
      fe.F1DetalleMonCotiz = 1

      fe.F1DetalleTributoItemCantidad = xitemper
'TEM
    xp = 0
    If Round(datcomp.Recordset.Fields("perceptem"), 2) <> 0 Then
      fe.f1IndiceItem = xp
      fe.F1DetalleTributoId = 99
      fe.F1DetalleTributoDesc = "TEM/PYP"
      fe.F1DetalleTributoBaseImp = Round(xneto, 2)
      fe.F1DetalleTributoAlic = 1.38
      fe.F1DetalleTributoImporte = Round(datcomp.Recordset.Fields("perceptem"), 2)
      xp = xp + 1
    End If
' IIBB
    If Round(datcomp.Recordset.Fields("percepiibb"), 2) <> 0 Then
      fe.f1IndiceItem = xp
      fe.F1DetalleTributoId = 2
      fe.F1DetalleTributoDesc = "IIBB"
      fe.F1DetalleTributoBaseImp = Round(xneto, 2)
      fe.F1DetalleTributoAlic = Round((datcomp.Recordset.Fields("percepiibb") / xneto) * 100, 2)
      fe.F1DetalleTributoImporte = Round(datcomp.Recordset.Fields("percepiibb"), 2)
      xp = xp + 1
    End If


      fe.F1DetalleIvaItemCantidad = xitemiva
' Iva 21
   xi = 0
   xbaseiva21 = 0
     If Round(datcomp.Recordset.Fields("importeiva21"), 2) <> 0 Then
      fe.f1IndiceItem = xi
      fe.F1DetalleIvaId = 5
      
     If Round(datcomp.Recordset.Fields("importeiva105"), 2) <> 0 Then
      fe.F1DetalleIvaBaseImp = Round(datcomp.Recordset.Fields("importeiva21") / 0.21, 2)
      xbaseiva21 = Round(datcomp.Recordset.Fields("importeiva21") / 0.21, 2)
     Else
     If datcomp.Recordset.Fields("importeiva21") <> xtotaliva And datcomp.Recordset.Fields("importeiva105") = 0 Then
      fe.F1DetalleImpIva = xtotaliva
     Else
      fe.F1DetalleImpIva = Round(datcomp.Recordset.Fields("importeiva21"), 2)
     End If
      fe.F1DetalleIvaBaseImp = Round(xneto, 2)
      xbaseiva21 = Round(xneto, 2)
     End If

     If datcomp.Recordset.Fields("importeiva21") <> xtotaliva And datcomp.Recordset.Fields("importeiva105") = 0 Then
        fe.F1DetalleIvaImporte = xtotaliva
     Else
        fe.F1DetalleIvaImporte = Round(datcomp.Recordset.Fields("importeiva21"), 2)
     End If

      xi = xi + 1
   End If
      
 'Iva 105
    If Round(datcomp.Recordset.Fields("importeiva105"), 2) <> 0 Then
      fe.f1IndiceItem = xi
      fe.F1DetalleIvaId = 4
      fe.F1DetalleIvaBaseImp = xneto - xbaseiva21
      fe.F1DetalleIvaImporte = Round(datcomp.Recordset.Fields("importeiva105"), 2)
    End If

      fe.F1DetalleCbtesAsocItemCantidad = 0
      fe.F1DetalleOpcionalItemCantidad = 0

      fe.ArchivoXMLRecibido = "c:\recibido.xml"
      fe.ArchivoXMLEnviado = "c:\enviado.xml"

      lResultado = fe.F1CAESolicitar()
      
      If lResultado Then
         MsgBox "Nro de CAE: " + fe.F1RespuestaDetalleCae + " -- Nro de Factura: " + Str(nro)
                        datencabezado.Recordset.Fields("numerodefactura") = nro
                        datencabezado.Recordset.Fields("puntodeventa") = Right("0000" + Replace(Str(PtoVta), " ", ""), 4)
                        datencabezado.Recordset.Fields("fechadelcomprobante") = Now()
                        datencabezado.Recordset.Fields("nroorden") = fe.F1RespuestaDetalleCae
                        datencabezado.Recordset.Fields("estadoimpresion") = fe.F1RespuestaDetalleCAEFchVto
                        datencabezado.Recordset.UpdateBatch adAffectCurrent
                        
'------ Actualiza Calipso
    xtransaccion = datencabezado.Recordset.Fields("calipsoid")
    datcalipso.ConnectionString = login.conexiontotal
    
If IsNull(xtransaccion) = False Then
xquery1 = "select * from  TRFACTURAVENTA where id = '" & xtransaccion & "' "
datcalipso.RecordSource = xquery1
datcalipso.Refresh
If datcalipso.Recordset.EOF = False Then
    datcalipso.Recordset.Fields("fechaactual") = FechaComp + "000000000"
    xfechacom = Right(FechaComp, 2) + "-" + Mid(FechaComp, 5, 2) + "-" + Left(FechaComp, 4)
    datcalipso.Recordset.Fields("numerodocumento") = Right("0000" + Replace(Str(PtoVta), " ", ""), 4) + Right("00000000" + Replace(Str(nro), " ", ""), 8)
    xnombre = Left(datcalipso.Recordset.Fields("nombre"), 16) + datcalipso.Recordset.Fields("numerodocumento") + Right(datcalipso.Recordset.Fields("nombre"), 19)
    xnombre1 = Left(xnombre, Len(xnombre) - 10) + xfechacom
    datcalipso.Recordset.Fields("nombre") = xnombre1
    datcalipso.Recordset.UpdateBatch adAffectCurrent
End If
End If
         
      Else
         
          MsgBox ("error detallado comprobante: " + fe.F1RespuestaDetalleObservacionMsg1)
         
      End If
   Else
      MsgBox ("fallo acceso " + fe.UltimoMensajeError)
   End If
Else
   MsgBox ("fallo iniciar " + fe.UltimoMensajeError)
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

Option1.Value = True
lista_facturas_todas.Top = yventana - lista_facturas_todas.Height / 2
lista_facturas_todas.Left = xventana - lista_facturas_todas.Width / 2

DesdeFecha.Value = Date - Day(Date) + 1
HastaFecha.Value = Date

datpresupuesto.ConnectionString = login.conexiontotal
datitems.ConnectionString = login.conexiontotal

xsuc = login.nomsucursal

                     
xquery1 = "SELECT  distinct ud_ezi_puntodeventa_encabezado.claveprimaria AS Numero, ud_ezi_puntodeventa_encabezado.numerodefactura AS Nro, " & _
                      "ud_ezi_puntodeventa_encabezado.fechadelcomprobante AS Fecha, ud_ezi_puntodeventa_encabezado.cliente AS Cliente, V_PERSONA_.CUIT, " & _
                      "ud_ezi_puntodeventa_encabezado.vendedor AS Vendedor, ud_ezi_puntodeventa_encabezado.importeglobal AS Importe, ud_ezi_puntodeventa_encabezado.id, " & _
                      "ud_ezi_puntodeventa_encabezado.generada, ud_ezi_puntodeventa_encabezado.tipodefactura, " & _
                      "ud_ezi_puntodeventa_encabezado.numerodefactura + 'n ' + ud_ezi_puntodeventa_encabezado.cliente + ' ' + V_PERSONA_.CUIT + ' ' + ud_ezi_puntodeventa_encabezado.vendedor " & _
                      "AS concatenado, ud_ezi_puntodeventa_encabezado.alquiler, ISNULL(V_COMPROMISOPAGO_.SALDO2_IMPORTE, ud_ezi_puntodeventa_encabezado.importeglobal) " & _
                      "AS SALDO, 0 AS NCIMPORTE, ud_ezi_puntodeventa_encabezado.nroorden as CAE, ud_ezi_puntodeventa_encabezado.recetaid as Autoriza, " & _
                      "case when ud_ezi_puntodeventa_encabezado.numeradorinterno like '%Mostrador%' then 'Contado' ELse 'Cta.Cte' end as Tipo, ud_ezi_puntodeventa_encabezado.nroorden as CAE " & _
                      "FROM         V_TRFACTURAVENTA_ RIGHT OUTER JOIN " & _
                      "ud_ezi_puntodeventa_encabezado WITH (readpast) ON V_TRFACTURAVENTA_.ID = ud_ezi_puntodeventa_encabezado.calipsoid LEFT OUTER JOIN " & _
                      "V_TRCREDITOVENTA_ ON ud_ezi_puntodeventa_encabezado.calipsoid = V_TRCREDITOVENTA_.VINCULOTR_ID LEFT OUTER JOIN " & _
                      "V_COMPROMISOPAGO_ ON ud_ezi_puntodeventa_encabezado.calipsoid = V_COMPROMISOPAGO_.TRORIGINANTE_ID LEFT OUTER JOIN " & _
                      "V_PERSONA_ RIGHT OUTER JOIN " & _
                      "V_CLIENTE_ ON V_PERSONA_.ID = V_CLIENTE_.ENTEASOCIADO_ID ON ud_ezi_puntodeventa_encabezado.clienteid = V_CLIENTE_.ID " & _
                      "WHERE     (ud_ezi_puntodeventa_encabezado.numeradorinterno LIKE '%Factura de Venta%') and ud_ezi_puntodeventa_encabezado.sucursal = '" & xsuc & "' AND (ud_ezi_puntodeventa_encabezado.tipodefacturacionid <> 'NN') AND (V_TRFACTURAVENTA_.FLAG_ID IS NULL) AND (V_COMPROMISOPAGO_.NIVEL = 1 OR " & _
                      "V_COMPROMISOPAGO_.NIVEL IS NULL) and " & _
                      "(ud_ezi_puntodeventa_encabezado.fechadelcomprobante >= '" & DesdeFecha.Value & "') and " & _
                      "(ud_ezi_puntodeventa_encabezado.fechadelcomprobante < '" & HastaFecha.Value + 1 & "') " & _
                      "ORDER BY Fecha DESC"

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

Private Sub KewlButtons1_Click()
On Error Resume Next
Dim crApp As New CRAXDRT.Application
Dim crReport As New CRAXDRT.Report
Dim tabla As String
Dim ruta As String


Kill ("c:\util\*.pdf")
Kill ("c:\util\*.rtf")

reporte.SQL = "SELECT v_ezi_pos_factctacte.id, v_ezi_pos_factctacte.NUMERODOCUMENTO, v_ezi_pos_factctacte.FECHAEMISION, v_ezi_pos_factctacte.cod_cliente, v_ezi_pos_factctacte.cliente, v_ezi_pos_factctacte.CUIT, v_ezi_pos_factctacte.CALLE, v_ezi_pos_factctacte.CODPOS, v_ezi_pos_factctacte.provincia, v_ezi_pos_factctacte.detalle, v_ezi_pos_factctacte.tipopago, v_ezi_pos_factctacte.codigoproducto, v_ezi_pos_factctacte.nombre_producto, v_ezi_pos_factctacte.cantidadproducto, v_ezi_pos_factctacte.nota, v_ezi_pos_factctacte.condiva, v_ezi_pos_factctacte.ciudad, v_ezi_pos_factctacte.TIPOVENTA, v_ezi_pos_factctacte.SIMBOLO, v_ezi_pos_factctacte.CODVENDEDOR, v_ezi_pos_factctacte.preciusiniva, v_ezi_pos_factctacte.subtotalsiniva, v_ezi_pos_factctacte.impbonifsiniva, v_ezi_pos_factctacte.nroremito, v_ezi_pos_factctacte.percepiibb, v_ezi_pos_factctacte.perceptem, v_ezi_pos_factctacte.totaltr, v_ezi_pos_factctacte.importeiva21, v_ezi_pos_factctacte.importeiva105, v_ezi_pos_factctacte.iditem,  " & _
              "v_ezi_pos_factctacte.cae, v_ezi_pos_factctacte.vto " & _
              "FROM  MMOSSE.DBO.v_ezi_pos_factctacte v_ezi_pos_factctacte " & _
              "where v_ezi_pos_factctacte.id = '" & DataGrid1.Columns(7).Text & "' order by v_ezi_pos_factctacte.iditem"

tabla = reporte.SQL

With CrystalReporte
    .PrinterCollation = crptCollated
    .Formulas(0) = "copia="" ORIGINAL """
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
    .Destination = crptToFile
    .PrintFileType = crptRTF
    .PrintFileName = "c:\util\Factura " + Replace(DataGrid1.Columns("cliente"), "*", "") + ".pdf"
'    .WindowState = crptMaximized
    .Action = 1
    
'     PDFCreator_CreatePDF Me, "c:\util\Factura " + Replace(DataGrid1.Columns("cliente"), "*", "") + ".rtf", "c:\util\Factura " + Replace(DataGrid1.Columns("cliente"), "*", "") + ".pdf"
    
End With
    
 Set oApp = New Outlook.Application
 Set myItem = oApp.CreateItem(Outlook.OlItemType.olMailItem)
 Set myAttachments = myItem.Attachments
' myAttachments.Add "c:\util\Factura " + Replace(DataGrid1.Columns("cliente"), "*", "") + ".pdf", 4, 2, " "
 myAttachments.Add "c:\util\Factura " + Replace(DataGrid1.Columns("cliente"), "*", "") + ".pdf", 4, 2, " "
 myItem.Display

 Set oApp = Nothing
 Set myItem = Nothing
 Set myAttachments = Nothing
    
Exit Sub

fuera:
    
    MsgBox "Reporte de Factura no Encontado, o error de configuracion de reporte", vbCritical, "Error"


Exit Sub


End Sub


Private Sub salir_Click()
    
    Unload Me

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)

On Error Resume Next


    If KeyAscii = 13 Then
        KeyAscii = 0
        xsuc = login.nomsucursal
        If Text1.Text <> "" Then
            xbusqueda = "%" + Text1.Text + "%"
'            xquery1 = "SELECT     ud_ezi_puntodeventa_encabezado.claveprimaria AS Numero, ud_ezi_puntodeventa_encabezado.numerodefactura as Nro, ud_ezi_puntodeventa_encabezado.fechadelcomprobante AS Fecha, " & _
'                      "ud_ezi_puntodeventa_encabezado.cliente AS Cliente, V_PERSONA_.CUIT, ud_ezi_puntodeventa_encabezado.vendedor AS Vendedor, " & _
'                      "ud_ezi_puntodeventa_encabezado.importeglobal AS Importe, ud_ezi_puntodeventa_encabezado.id, ud_ezi_puntodeventa_encabezado.generada, ud_ezi_puntodeventa_encabezado.tipodefactura, " & _
'                      "ud_ezi_puntodeventa_encabezado.numerodefactura + 'n ' + ud_ezi_puntodeventa_encabezado.cliente + ' ' + V_PERSONA_.CUIT + ' ' + ud_ezi_puntodeventa_encabezado.vendedor AS concatenado, ud_ezi_puntodeventa_encabezado.alquiler " & _
'                      "FROM V_PERSONA_ RIGHT OUTER JOIN V_CLIENTE_ ON V_PERSONA_.ID = V_CLIENTE_.ENTEASOCIADO_ID RIGHT OUTER JOIN " & _
'                      "ud_ezi_puntodeventa_encabezado WITH (readpast) ON V_CLIENTE_.ID = ud_ezi_puntodeventa_encabezado.clienteid " & _
'                      "WHERE     (ud_ezi_puntodeventa_encabezado.numeradorinterno like '%Factura de Venta%') and (ud_ezi_puntodeventa_encabezado.tipodefacturacionid <> 'NN') AND " & _
'                      "(ud_ezi_puntodeventa_encabezado.numerodefactura + 'n ' + ud_ezi_puntodeventa_encabezado.cliente + ' ' + V_PERSONA_.CUIT + ' ' + ud_ezi_puntodeventa_encabezado.vendedor like '" & xbusqueda & "') " & _
'                      "ORDER BY Fecha DESC"
                      
            xquery1 = "SELECT distinct     ud_ezi_puntodeventa_encabezado.claveprimaria AS Numero, ud_ezi_puntodeventa_encabezado.numerodefactura AS Nro, " & _
                      "ud_ezi_puntodeventa_encabezado.fechadelcomprobante AS Fecha, ud_ezi_puntodeventa_encabezado.cliente AS Cliente, V_PERSONA_.CUIT, " & _
                      "ud_ezi_puntodeventa_encabezado.vendedor AS Vendedor, ud_ezi_puntodeventa_encabezado.importeglobal AS Importe, ud_ezi_puntodeventa_encabezado.id, " & _
                      "ud_ezi_puntodeventa_encabezado.generada, ud_ezi_puntodeventa_encabezado.tipodefactura, " & _
                      "ud_ezi_puntodeventa_encabezado.numerodefactura + 'n ' + ud_ezi_puntodeventa_encabezado.cliente + ' ' + V_PERSONA_.CUIT + ' ' + ud_ezi_puntodeventa_encabezado.vendedor " & _
                      "AS concatenado, ud_ezi_puntodeventa_encabezado.alquiler, ISNULL(V_COMPROMISOPAGO_.SALDO2_IMPORTE, ud_ezi_puntodeventa_encabezado.importeglobal) " & _
                      "AS SALDO, 0 AS NCIMPORTE, ud_ezi_puntodeventa_encabezado.recetaid as Autoriza, ud_ezi_puntodeventa_encabezado.nroorden as CAE " & _
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
                    
            datpresupuesto.RecordSource = xquery1
            datpresupuesto.Refresh
            DataGrid1.Columns(0).Visible = False
            DataGrid1.Columns(7).Visible = False
            DataGrid1.Columns(8).Visible = False
            DataGrid1.Columns(9).Visible = False
            DataGrid1.Columns(10).Visible = False
            DataGrid1.Columns(12).Visible = False
            DataGrid1.Columns(13).Visible = False
            DataGrid1.Columns(15).Visible = True
            DataGrid1.Columns("alquiler").Visible = False
            DataGrid1.Columns("cae").Visible = False
            DataGrid1.Columns(1).Width = 1000
            DataGrid1.Columns(3).Width = 3500
            DataGrid1.Columns(6).Alignment = dbgRight
            DataGrid1.Columns(6).NumberFormat = "Currency"
            
                        
            DataGrid2.Columns(1).Width = 3500
            DataGrid2.Columns(3).Alignment = dbgRight
            DataGrid2.Columns(5).Alignment = dbgRight
            DataGrid2.Columns(6).Alignment = dbgRight


        End If
        DataGrid1.SetFocus
        
        
    End If

End Sub
