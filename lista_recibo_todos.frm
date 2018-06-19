VERSION 5.00
Object = "{F6125AB1-8AB1-11CE-A77F-08002B2F4E98}#2.0#0"; "MSRDC20.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{D18BBD1F-82BB-4385-BED3-E9D31A3E361E}#1.0#0"; "KewlButtonz.ocx"
Begin VB.Form lista_recibo_todos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta de Recibos"
   ClientHeight    =   8760
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13890
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8760
   ScaleWidth      =   13890
   Begin VB.Frame Frame1 
      Caption         =   "Filtro de Comprobantes"
      Height          =   1455
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   13695
      Begin VB.CommandButton FACTURAELECTRONICA 
         Caption         =   "Fac.Electronica"
         Height          =   375
         Left            =   12360
         TabIndex        =   16
         Top             =   840
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Original"
         Height          =   255
         Left            =   9240
         TabIndex        =   15
         Top             =   960
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Duplicado"
         Height          =   255
         Left            =   9240
         TabIndex        =   14
         Top             =   480
         Visible         =   0   'False
         Width           =   1095
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
         Caption         =   "Imprimir Reporte"
         Height          =   375
         Left            =   7680
         TabIndex        =   7
         Top             =   720
         Width           =   1335
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Generar PDF"
         Height          =   375
         Left            =   8520
         TabIndex        =   6
         Top             =   120
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
         Format          =   49676289
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
         Format          =   49676289
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
         MICON           =   "lista_recibo_todos.frx":0000
         PICN            =   "lista_recibo_todos.frx":001C
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
         Left            =   10560
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
         MICON           =   "lista_recibo_todos.frx":0B66
         PICN            =   "lista_recibo_todos.frx":0B82
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
         WindowShowSearchBtn=   -1  'True
         WindowShowPrintSetupBtn=   -1  'True
         WindowShowRefreshBtn=   -1  'True
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
         Left            =   10560
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
         MICON           =   "lista_recibo_todos.frx":3F74
         PICN            =   "lista_recibo_todos.frx":3F90
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
      Bindings        =   "lista_recibo_todos.frx":43E2
      Height          =   6855
      Left            =   120
      TabIndex        =   0
      Top             =   1560
      Width           =   13695
      _ExtentX        =   24156
      _ExtentY        =   12091
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
      Bindings        =   "lista_recibo_todos.frx":43FF
      Height          =   1215
      Left            =   120
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   5400
      Visible         =   0   'False
      Width           =   13695
      _ExtentX        =   24156
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
Attribute VB_Name = "lista_recibo_todos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public cuentacont As String
Dim cuenta(99999) As Integer
Public xbusqueda As String

Dim fldName As String
Dim fName As String
Dim sAttName As String
Dim strName As String
Dim oApp As Outlook.Application
Dim myItem As Outlook.MailItem
Dim myAttachments As Outlook.Attachments






Private Sub Command4_Click()
On Error Resume Next
Dim crApp As New CRAXDRT.Application
Dim crReport As New CRAXDRT.Report
Dim tabla As String
Dim ruta As String


reporte.SQL = "SELECT v_ezi_pos_recibo.GRUPO,v_ezi_pos_recibo.id, v_ezi_pos_recibo.nrorecibo, v_ezi_pos_recibo.fechadelcomprobante, v_ezi_pos_recibo.cliente, v_ezi_pos_recibo.CUIT, v_ezi_pos_recibo.CODPOS, v_ezi_pos_recibo.CALLE, v_ezi_pos_recibo.LOCALIDAD, v_ezi_pos_recibo.CONDIVA, v_ezi_pos_recibo.Comprobante, v_ezi_pos_recibo.Cancela, v_ezi_pos_recibo.totalfactura, v_ezi_pos_recibo.formadepago, v_ezi_pos_recibo.banco, v_ezi_pos_recibo.tarjeta, v_ezi_pos_recibo.numerocheque, v_ezi_pos_recibo.monto, v_ezi_pos_recibo.fechaemision, v_ezi_pos_recibo.fechavencimiento FROM MMOSSE.dbo.v_ezi_pos_recibo v_ezi_pos_recibo where v_ezi_pos_recibo.id = " & datagrid1.Columns("id") & " ORDER BY v_ezi_pos_recibo.GRUPO ASC "
tabla = reporte.SQL

With CrystalReporte
    .PrinterCollation = crptCollated
    .ReportFileName = App.Path & "\Recibos.rpt"
    .WindowTitle = "Remito Vta Orig"
    '.Connect = "PROVIDER=MSDASQL;dsn=facturacion;uid=lucva;pwd=25072004;database=facturacionsql;"
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
            xquery1 = "select fechadelcomprobante, puntodeventa+RIGHT('00000000'+numerodefactura,8) AS nrorecibo, cliente, totaltr, sucursal, id from ud_ezi_puntodeventa_encabezado " & _
                      "where numeradorinterno = 'Recibo de Cobrana' and ud_ezi_puntodeventa_encabezado.sucursal = '" & xsuc & "' and " & _
                      "(ud_ezi_puntodeventa_encabezado.fechadelcomprobante >= '" & DesdeFecha.Value & "') and " & _
                      "(ud_ezi_puntodeventa_encabezado.fechadelcomprobante < '" & HastaFecha.Value + 1 & "') " & _
                      "ORDER BY ud_ezi_puntodeventa_encabezado.fechadelcomprobante DESC"

Else
            xquery1 = xquery1 = "select fechadelcomprobante, puntodeventa+RIGHT('00000000'+numerodefactura,8) AS nrorecibo, cliente, totaltr, sucursal, id from ud_ezi_puntodeventa_encabezado " & _
                      "where numeradorinterno = 'Recibo de Cobrana' and ud_ezi_puntodeventa_encabezado.sucursal = '" & xsuc & "' and " & _
                      "(ud_ezi_puntodeventa_encabezado.fechadelcomprobante >= '" & DesdeFecha.Value & "') and " & _
                      "(ud_ezi_puntodeventa_encabezado.fechadelcomprobante < '" & HastaFecha.Value + 1 & "') and " & _
                      "(ud_ezi_puntodeventa_encabezado.numerodefactura + 'n ' + ud_ezi_puntodeventa_encabezado.cliente  " & _
                      "LIKE '" & xbusqueda & "') " & _
                      "ORDER BY ud_ezi_puntodeventa_encabezado.fechadelcomprobante DESC"
End If

datpresupuesto.RecordSource = xquery1
datpresupuesto.Refresh

            datagrid1.Columns(4).Visible = False
            datagrid1.Columns(5).Visible = False
            datagrid1.Columns(1).Width = 2000
            datagrid1.Columns(2).Width = 3500
            datagrid1.Columns(3).Alignment = dbgRight
            datagrid1.Columns(3).NumberFormat = "Currency"
        
        

End Sub

Private Sub Command6_Click()
'On Error GoTo fuera
Dim crApp As New CRAXDRT.Application
Dim crReport As New CRAXDRT.Report
Dim tabla As String
Dim ruta As String

datpresupuesto.Recordset.MoveFirst

Do While Not datpresupuesto.Recordset.EOF

reporte.SQL = "SELECT v_ezi_pos_factctacte.id, v_ezi_pos_factctacte.NUMERODOCUMENTO, v_ezi_pos_factctacte.FECHAEMISION, v_ezi_pos_factctacte.cod_cliente, v_ezi_pos_factctacte.cliente, v_ezi_pos_factctacte.CUIT, v_ezi_pos_factctacte.CALLE, v_ezi_pos_factctacte.CODPOS, v_ezi_pos_factctacte.provincia, v_ezi_pos_factctacte.detalle, v_ezi_pos_factctacte.tipopago, v_ezi_pos_factctacte.codigoproducto, v_ezi_pos_factctacte.nombre_producto, v_ezi_pos_factctacte.cantidadproducto, v_ezi_pos_factctacte.nota, v_ezi_pos_factctacte.condiva, v_ezi_pos_factctacte.ciudad, v_ezi_pos_factctacte.TIPOVENTA, v_ezi_pos_factctacte.SIMBOLO, v_ezi_pos_factctacte.CODVENDEDOR, v_ezi_pos_factctacte.preciusiniva, v_ezi_pos_factctacte.subtotalsiniva, v_ezi_pos_factctacte.impbonifsiniva, v_ezi_pos_factctacte.nroremito, v_ezi_pos_factctacte.percepiibb, v_ezi_pos_factctacte.perceptem, v_ezi_pos_factctacte.totaltr, v_ezi_pos_factctacte.importeiva21, v_ezi_pos_factctacte.importeiva105, v_ezi_pos_factctacte.iditem,  " & _
              "v_ezi_pos_factctacte.cae, v_ezi_pos_factctacte.vto " & _
              "FROM  MMOSSE.DBO.v_ezi_pos_factctacte v_ezi_pos_factctacte " & _
              "where v_ezi_pos_factctacte.id = '" & datagrid1.Columns(7).Text & "' order by v_ezi_pos_factctacte.iditem"

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
    If datagrid1.Columns(9).Text = "A" Then
      If tipofac <> "NN" Then
       If datagrid1.Columns(11).Text = "N" Then
        .ReportFileName = App.Path & "\FacturaCtaCteA.rpt"
       Else
        .ReportFileName = App.Path & "\FacturaCtaCteA_alquiler.rpt"
       End If
      Else
        .ReportFileName = App.Path & "\PresupuestoA.rpt"
      End If
    Else
      If tipofac <> "NN" Then
       If datagrid1.Columns(11).Text = "N" Then
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


 reporte.SQL = "SELECT v_ezi_pos_reporte_cobranzas.Numero, v_ezi_pos_reporte_cobranzas.tipodefactura, v_ezi_pos_reporte_cobranzas.fechadelcomprobante, v_ezi_pos_reporte_cobranzas.codcliente, v_ezi_pos_reporte_cobranzas.cliente, v_ezi_pos_reporte_cobranzas.totaltr, v_ezi_pos_reporte_cobranzas.subtotalsiniva, v_ezi_pos_reporte_cobranzas.totaliva, v_ezi_pos_reporte_cobranzas.percepiibb, v_ezi_pos_reporte_cobranzas.tpago, v_ezi_pos_reporte_cobranzas.formadepago, v_ezi_pos_reporte_cobranzas.fechadeemision, v_ezi_pos_reporte_cobranzas.fechadevencimiento, v_ezi_pos_reporte_cobranzas.monto, v_ezi_pos_reporte_cobranzas.id FROM MMOSSE.dbo.v_ezi_pos_reporte_cobranzas v_ezi_pos_reporte_cobranzas " & _
               " WHERE     convert(date,v_ezi_pos_reporte_cobranzas.fechadelcomprobante) >= '" & DesdeFecha.Value & "' and " & _
               "convert(date,v_ezi_pos_reporte_cobranzas.fechadelcomprobante) <= '" & HastaFecha.Value & "' ORDER BY v_ezi_pos_reporte_cobranzas.id ASC "

 tabla = reporte.SQL

With CrystalReporte
    .PrinterCollation = crptCollated
    .ReportFileName = App.Path & "\Reporte_listado_cobranzas.rpt"
    .WindowTitle = "Reporte de Cobranzas"
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

Private Sub Command8_Click()

End Sub

Private Sub DataGrid1_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        Call Command4_Click
            
    End If

End Sub




Private Sub Form_Activate()

datagrid1.SetFocus

End Sub

Private Sub Form_Load()
Aplicar_skin Me

MiFuncionDeAjuste Me, True

yventana = Inicio.Height / 2 - 1000
xventana = Inicio.Width / 2

Option1.Value = True
lista_recibo_todos.Top = yventana - lista_recibo_todos.Height / 2
lista_recibo_todos.Left = xventana - lista_recibo_todos.Width / 2

DesdeFecha.Value = Date - Day(Date) + 1
HastaFecha.Value = Date

datpresupuesto.ConnectionString = login.conexiontotal
datitems.ConnectionString = login.conexiontotal

xsuc = login.nomsucursal


xquery1 = "select fechadelcomprobante, puntodeventa+RIGHT('00000000'+numerodefactura,8) AS nrorecibo, cliente, totaltr, sucursal, id from ud_ezi_puntodeventa_encabezado " & _
          "where numeradorinterno = 'Recibo de Cobrana' and ud_ezi_puntodeventa_encabezado.sucursal = '" & xsuc & "' and " & _
          "(ud_ezi_puntodeventa_encabezado.fechadelcomprobante >= '" & DesdeFecha.Value & "') and " & _
          "(ud_ezi_puntodeventa_encabezado.fechadelcomprobante < '" & HastaFecha.Value + 1 & "') " & _
          "ORDER BY ud_ezi_puntodeventa_encabezado.fechadelcomprobante DESC"
                      

datpresupuesto.RecordSource = xquery1
datpresupuesto.Refresh

            datagrid1.Columns(4).Visible = False
            datagrid1.Columns(5).Visible = False
            datagrid1.Columns(1).Width = 2000
            datagrid1.Columns(2).Width = 3500
            datagrid1.Columns(3).Alignment = dbgRight
            datagrid1.Columns(3).NumberFormat = "Currency"
            
            

 
End Sub

Private Sub KewlButtons1_Click()
On Error Resume Next
Dim crApp As New CRAXDRT.Application
Dim crReport As New CRAXDRT.Report
Dim tabla As String
Dim ruta As String

Kill ("c:\util\*.pdf")
Kill ("c:\util\*.rtf")

reporte.SQL = "SELECT v_ezi_pos_recibo.GRUPO,v_ezi_pos_recibo.id, v_ezi_pos_recibo.nrorecibo, v_ezi_pos_recibo.fechadelcomprobante, v_ezi_pos_recibo.cliente, v_ezi_pos_recibo.CUIT, v_ezi_pos_recibo.CODPOS, v_ezi_pos_recibo.CALLE, v_ezi_pos_recibo.LOCALIDAD, v_ezi_pos_recibo.CONDIVA, v_ezi_pos_recibo.Comprobante, v_ezi_pos_recibo.Cancela, v_ezi_pos_recibo.totalfactura, v_ezi_pos_recibo.formadepago, v_ezi_pos_recibo.banco, v_ezi_pos_recibo.tarjeta, v_ezi_pos_recibo.numerocheque, v_ezi_pos_recibo.monto, v_ezi_pos_recibo.fechaemision, v_ezi_pos_recibo.fechavencimiento FROM MMOSSE.dbo.v_ezi_pos_recibo v_ezi_pos_recibo where v_ezi_pos_recibo.id = " & datagrid1.Columns("id") & " ORDER BY v_ezi_pos_recibo.GRUPO ASC "
tabla = reporte.SQL

With CrystalReporte
    .PrinterCollation = crptCollated
    .ReportFileName = App.Path & "\Recibos.rpt"
    .WindowTitle = "Remito Vta Orig"
    '.Connect = "PROVIDER=MSDASQL;dsn=facturacion;uid=lucva;pwd=25072004;database=facturacionsql;"
    .Connect = login.conexionreporte
    .DiscardSavedData = Truertf
    .RetrieveDataFiles
    .ReportSource = 0
    .SQLQuery = tabla
   .Destination = crptToFile
    .PrintFileType = crptRTF

    .PrintFileName = "c:\util\Recibo " + Replace(datagrid1.Columns("cliente"), "*", "") + ".pdf"

    .WindowState = crptMaximized
    .Action = 1
    
'    PDFCreator_CreatePDF Me, "c:\util\Recibo " + Replace(DataGrid1.Columns("cliente"), "*", "") + ".rtf", "c:\util\Recibo " + Replace(DataGrid1.Columns("cliente"), "*", "") + ".pdf"
    
End With

 Set oApp = New Outlook.Application
 Set myItem = oApp.CreateItem(Outlook.OlItemType.olMailItem)
 Set myAttachments = myItem.Attachments
 myAttachments.Add "c:\util\Recibo " + Replace(datagrid1.Columns("cliente"), "*", "") + ".pdf", 4, 2, " "
 myItem.Display

 Set oApp = Nothing
 Set myItem = Nothing
 Set myAttachments = Nothing
 
  
 

Exit Sub

fuera:
    
    MsgBox "Reporte no Encontado, o error de configuracion de reporte", vbCritical, "Error"



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
                     
            xquery1 = "select fechadelcomprobante, puntodeventa+RIGHT('00000000'+numerodefactura,8) AS nrorecibo, cliente, totaltr, sucursal, id from ud_ezi_puntodeventa_encabezado " & _
                      "where numeradorinterno = 'Recibo de Cobrana' and ud_ezi_puntodeventa_encabezado.sucursal = '" & xsuc & "' and " & _
                      "(ud_ezi_puntodeventa_encabezado.fechadelcomprobante >= '" & DesdeFecha.Value & "') and " & _
                      "(ud_ezi_puntodeventa_encabezado.fechadelcomprobante < '" & HastaFecha.Value + 1 & "') and " & _
                      "(ud_ezi_puntodeventa_encabezado.numerodefactura + 'n ' + ud_ezi_puntodeventa_encabezado.cliente  " & _
                      "LIKE '" & xbusqueda & "') " & _
                      "ORDER BY ud_ezi_puntodeventa_encabezado.fechadelcomprobante DESC"
                      

            datpresupuesto.RecordSource = xquery1
            datpresupuesto.Refresh

            datagrid1.Columns(4).Visible = False
            datagrid1.Columns(5).Visible = False
            datagrid1.Columns(1).Width = 2000
            datagrid1.Columns(2).Width = 3500
            datagrid1.Columns(3).Alignment = dbgRight
            datagrid1.Columns(3).NumberFormat = "Currency"
            


        End If
        datagrid1.SetFocus
        
        
    End If

End Sub
