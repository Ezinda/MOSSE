VERSION 5.00
Object = "{F6125AB1-8AB1-11CE-A77F-08002B2F4E98}#2.0#0"; "MSRDC20.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{D18BBD1F-82BB-4385-BED3-E9D31A3E361E}#1.0#0"; "KewlButtonz.ocx"
Begin VB.Form frmcentrodefacturacion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta de Documentos"
   ClientHeight    =   7635
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15240
   Icon            =   "frmcentrodefacturacion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7635
   ScaleWidth      =   15240
   Begin VB.Frame Frame5 
      Caption         =   "Cta.Cte."
      Height          =   5895
      Left            =   120
      TabIndex        =   8
      Top             =   1080
      Width           =   14895
      Begin MSDataGridLib.DataGrid DataGrid3 
         Bindings        =   "frmcentrodefacturacion.frx":0442
         Height          =   4935
         Left            =   240
         TabIndex        =   9
         Top             =   360
         Width           =   12615
         _ExtentX        =   22251
         _ExtentY        =   8705
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
      Begin KewlButtonz.KewlButtons facturar 
         Height          =   495
         Left            =   13080
         TabIndex        =   12
         Top             =   720
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
         MICON           =   "frmcentrodefacturacion.frx":045A
         PICN            =   "frmcentrodefacturacion.frx":0476
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin KewlButtonz.KewlButtons KewlButtons3 
         Height          =   495
         Left            =   13080
         TabIndex        =   13
         Top             =   1680
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   873
         BTYPE           =   14
         TX              =   "Anular"
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
         MICON           =   "frmcentrodefacturacion.frx":0A10
         PICN            =   "frmcentrodefacturacion.frx":0A2C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSDataGridLib.DataGrid DataGrid2 
         Bindings        =   "frmcentrodefacturacion.frx":0FC6
         Height          =   2295
         Left            =   240
         TabIndex        =   14
         Top             =   3120
         Visible         =   0   'False
         Width           =   12615
         _ExtentX        =   22251
         _ExtentY        =   4048
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
   End
   Begin VB.Frame Frame4 
      Caption         =   "Contado"
      Height          =   2895
      Left            =   120
      TabIndex        =   6
      Top             =   1080
      Visible         =   0   'False
      Width           =   14895
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "frmcentrodefacturacion.frx":0FDE
         Height          =   2295
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   11895
         _ExtentX        =   20981
         _ExtentY        =   4048
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
      Begin KewlButtonz.KewlButtons KewlButtons1 
         Height          =   495
         Left            =   12600
         TabIndex        =   10
         Top             =   600
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   873
         BTYPE           =   14
         TX              =   "Facturar/&Cobrar"
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
         MICON           =   "frmcentrodefacturacion.frx":0FF7
         PICN            =   "frmcentrodefacturacion.frx":1013
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin KewlButtonz.KewlButtons KewlButtons2 
         Height          =   495
         Left            =   12600
         TabIndex        =   11
         Top             =   1560
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   873
         BTYPE           =   14
         TX              =   "Anular"
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
         MICON           =   "frmcentrodefacturacion.frx":15AD
         PICN            =   "frmcentrodefacturacion.frx":15C9
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
   Begin MSAdodcLib.Adodc datsaldar2 
      Height          =   330
      Left            =   5880
      Top             =   7200
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
      Left            =   240
      Top             =   7200
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
   Begin MSAdodcLib.Adodc datctacte 
      Height          =   330
      Left            =   7440
      Top             =   7200
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
      Left            =   4080
      TabIndex        =   4
      Top             =   7200
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc datremitos 
      Height          =   330
      Left            =   2880
      Top             =   7200
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
      Left            =   5400
      Top             =   7200
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTitle     =   "Orden de Pago"
      PrintFileLinesPerPage=   60
   End
   Begin MSRDC.MSRDC reporte 
      Height          =   330
      Left            =   1440
      Top             =   7200
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
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
   Begin VB.Frame Frame1 
      Caption         =   "Busqueda"
      Height          =   855
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   9615
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
         Width           =   9255
      End
   End
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   9960
      TabIndex        =   5
      Top             =   120
      Width           =   5055
      Begin KewlButtonz.KewlButtons Command4 
         Height          =   495
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   873
         BTYPE           =   14
         TX              =   "&Actualizar"
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
         MICON           =   "frmcentrodefacturacion.frx":1B63
         PICN            =   "frmcentrodefacturacion.frx":1B7F
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
         Left            =   2760
         TabIndex        =   2
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
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
         MICON           =   "frmcentrodefacturacion.frx":32F1
         PICN            =   "frmcentrodefacturacion.frx":330D
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
   Begin MSAdodcLib.Adodc datcontrol 
      Height          =   330
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   1
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   2
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
Attribute VB_Name = "frmcentrodefacturacion"
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

xsuc = login.nomsucursal
'Contado
xquery1 = "SELECT     ud_ezi_puntodeventa_encabezado.fechadelcomprobante AS Fecha, ud_ezi_puntodeventa_encabezado.numerodefactura AS Numero, " & _
          "ud_ezi_puntodeventa_encabezado.tipodefactura, ud_ezi_puntodeventa_encabezado.cliente AS Cliente, V_PERSONA_.CUIT, " & _
          "ud_ezi_puntodeventa_encabezado.vendedor AS Vendedor, ud_ezi_puntodeventa_encabezado.importeglobal AS Importe, ud_ezi_puntodeventa_encabezado.id, " & _
          "ud_ezi_puntodeventa_encabezado.generada, " & _
          "ud_ezi_puntodeventa_encabezado.numerodefactura + 'n ' + ud_ezi_puntodeventa_encabezado.cliente + ' ' + V_PERSONA_.CUIT + ' ' + ud_ezi_puntodeventa_encabezado.vendedor " & _
          "AS concatenado " & _
          "FROM V_PERSONA_ RIGHT OUTER JOIN " & _
          "V_CLIENTE_ ON V_PERSONA_.ID = V_CLIENTE_.ENTEASOCIADO_ID RIGHT OUTER JOIN " & _
          "ud_ezi_puntodeventa_encabezado WITH (readpast) ON V_CLIENTE_.ID = ud_ezi_puntodeventa_encabezado.clienteid " & _
          "WHERE     (ud_ezi_puntodeventa_encabezado.numeradorinterno = 'Nota de Venta') AND (ud_ezi_puntodeventa_encabezado.generada = 0) AND " & _
          "(ud_ezi_puntodeventa_encabezado.tipodepagoid = '{E88275FE-0A80-11D6-B0D1-004854841C8A}') AND (ud_ezi_puntodeventa_encabezado.importeglobal <> 0) AND (ud_ezi_puntodeventa_encabezado.sucursal = '" & xsuc & "') " & _
          "ORDER BY Fecha DESC "
'Cta.Cte
xquery2 = "SELECT     ud_ezi_puntodeventa_encabezado.fechadelcomprobante AS Fecha, ud_ezi_puntodeventa_encabezado.numerodefactura AS Numero, " & _
          "ud_ezi_puntodeventa_encabezado.tipodefactura, ud_ezi_puntodeventa_encabezado.cliente AS Cliente, V_PERSONA_.CUIT, " & _
          "ud_ezi_puntodeventa_encabezado.vendedor AS Vendedor, ud_ezi_puntodeventa_encabezado.importeglobal AS Importe, ud_ezi_puntodeventa_encabezado.id, " & _
          "ud_ezi_puntodeventa_encabezado.generada, " & _
          "ud_ezi_puntodeventa_encabezado.numerodefactura + 'n ' + ud_ezi_puntodeventa_encabezado.cliente + ' ' + V_PERSONA_.CUIT + ' ' + ud_ezi_puntodeventa_encabezado.vendedor " & _
          "AS concatenado " & _
          "FROM V_PERSONA_ RIGHT OUTER JOIN " & _
          "V_CLIENTE_ ON V_PERSONA_.ID = V_CLIENTE_.ENTEASOCIADO_ID RIGHT OUTER JOIN " & _
          "ud_ezi_puntodeventa_encabezado WITH (readpast) ON V_CLIENTE_.ID = ud_ezi_puntodeventa_encabezado.clienteid " & _
          "WHERE     (ud_ezi_puntodeventa_encabezado.numeradorinterno = 'Nota de Venta') AND (ud_ezi_puntodeventa_encabezado.generada = 0) AND " & _
          "(ud_ezi_puntodeventa_encabezado.tipodepagoid <> '{E88275FE-0A80-11D6-B0D1-004854841C8A}') AND (ud_ezi_puntodeventa_encabezado.importeglobal <> 0) AND (ud_ezi_puntodeventa_encabezado.sucursal = '" & xsuc & "') " & _
          "ORDER BY Fecha DESC "


datremitos.RecordSource = xquery1
datremitos.Refresh
datctacte.RecordSource = xquery2
datctacte.Refresh

            DataGrid1.Columns(1).Width = 1000
            DataGrid1.Columns(2).Width = 500
            DataGrid1.Columns(2).Caption = "T.Fac"
            DataGrid1.Columns(3).Width = 3500
            DataGrid1.Columns(4).Width = 1200
            DataGrid1.Columns(6).Alignment = dbgRight
            DataGrid1.Columns(6).NumberFormat = "Currency"
            DataGrid1.Columns(7).Visible = False
            DataGrid1.Columns(8).Visible = False
            DataGrid1.Columns(9).Visible = False

            DataGrid3.Columns(1).Width = 1000
            DataGrid3.Columns(2).Width = 500
            DataGrid3.Columns(2).Caption = "T.Fac"
            DataGrid3.Columns(3).Width = 3500
            DataGrid3.Columns(4).Width = 1200
            DataGrid3.Columns(6).Alignment = dbgRight
            DataGrid3.Columns(6).NumberFormat = "Currency"
            DataGrid3.Columns(7).Visible = False
            DataGrid3.Columns(8).Visible = False
            DataGrid3.Columns(9).Visible = False




End Sub




Private Sub Command4_Click()

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
         "N.fechadelcomprobante , N.sucursal, N.obraid, N.retira FROM  ud_ezi_puntodeventa_encabezado AS N LEFT OUTER JOIN v_ezi_pos_remito AS R ON N.id = R.presupuestobase " & _
         "WHERE (N.numeradorinterno = 'Nota de Venta') and  N.id ='" & DataGrid3.Columns(7).Text & "' "

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
              "N.fechadelcomprobante , N.sucursal, N.obraid, N.retira FROM  ud_ezi_puntodeventa_encabezado AS N LEFT OUTER JOIN v_ezi_pos_remito AS R ON N.id = R.presupuestobase " & _
              "WHERE (N.numeradorinterno = 'Nota de Venta') and  N.id ='" & DataGrid3.Columns(7).Text & "' "
   
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
    frmnota_venta_remito_factura.Show
    If X = 0 Then
    '    frmfacctacte_venta.Text18.Text = DataGrid1.Columns(12).Text
'        frmfacctacte_venta.Text17.Text = DataGrid1.Columns(2).Text
    End If
    frmnota_venta_remito_factura.Text17.SetFocus
    SendKeys "{ENTER}", False
    
    

End Sub

Private Sub Form_Activate()

'    DataGrid1.SetFocus
'    If menu = 1 Then
'        If Text1.Text = "" Then Text1.Text = " "
'        Text1.SetFocus
'        SendKeys "{ENTER}", False
'    End If


End Sub

Private Sub Form_Load()
On Error Resume Next
Aplicar_skin Me

yventana = Inicio.Height / 2 - 1000
xventana = Inicio.Width / 2

frmcentrodefacturacion.Top = yventana - frmcentrodefacturacion.Height / 2
frmcentrodefacturacion.Left = xventana - frmcentrodefacturacion.Width / 2


datremitos.ConnectionString = login.conexiontotal
datctacte.ConnectionString = login.conexiontotal
datsaldar.ConnectionString = login.conexiontotal
datsaldar2.ConnectionString = login.conexiontotal

xcon = 1


If UCase(login.usuarioactivo) = "ADMIN" Then
    saldaremito.Visible = True
Else
    saldaremito.Visible = False
End If

xsuc = login.nomsucursal
'Contado
xquery1 = "SELECT     ud_ezi_puntodeventa_encabezado.fechadelcomprobante AS Fecha, ud_ezi_puntodeventa_encabezado.numerodefactura AS Numero, " & _
          "ud_ezi_puntodeventa_encabezado.tipodefactura, ud_ezi_puntodeventa_encabezado.cliente AS Cliente, V_PERSONA_.CUIT, " & _
          "ud_ezi_puntodeventa_encabezado.vendedor AS Vendedor, ud_ezi_puntodeventa_encabezado.importeglobal AS Importe, ud_ezi_puntodeventa_encabezado.id, " & _
          "ud_ezi_puntodeventa_encabezado.generada, " & _
          "ud_ezi_puntodeventa_encabezado.numerodefactura + 'n ' + ud_ezi_puntodeventa_encabezado.cliente + ' ' + V_PERSONA_.CUIT + ' ' + ud_ezi_puntodeventa_encabezado.vendedor " & _
          "AS concatenado " & _
          "FROM V_PERSONA_ RIGHT OUTER JOIN " & _
          "V_CLIENTE_ ON V_PERSONA_.ID = V_CLIENTE_.ENTEASOCIADO_ID RIGHT OUTER JOIN " & _
          "ud_ezi_puntodeventa_encabezado WITH (readpast) ON V_CLIENTE_.ID = ud_ezi_puntodeventa_encabezado.clienteid " & _
          "WHERE     (ud_ezi_puntodeventa_encabezado.numeradorinterno = 'Nota de Venta') AND (ud_ezi_puntodeventa_encabezado.generada = 0) AND " & _
          "(ud_ezi_puntodeventa_encabezado.tipodepagoid = '{E88275FE-0A80-11D6-B0D1-004854841C8A}') AND (ud_ezi_puntodeventa_encabezado.importeglobal <> 0) AND (ud_ezi_puntodeventa_encabezado.sucursal = '" & xsuc & "')  " & _
          "ORDER BY Fecha DESC "
'Cta.Cte
xquery2 = "SELECT     ud_ezi_puntodeventa_encabezado.fechadelcomprobante AS Fecha, ud_ezi_puntodeventa_encabezado.numerodefactura AS Numero, " & _
          "ud_ezi_puntodeventa_encabezado.tipodefactura, ud_ezi_puntodeventa_encabezado.cliente AS Cliente, V_PERSONA_.CUIT, " & _
          "ud_ezi_puntodeventa_encabezado.vendedor AS Vendedor, ud_ezi_puntodeventa_encabezado.importeglobal AS Importe, ud_ezi_puntodeventa_encabezado.id, " & _
          "ud_ezi_puntodeventa_encabezado.generada, " & _
          "ud_ezi_puntodeventa_encabezado.numerodefactura + 'n ' + ud_ezi_puntodeventa_encabezado.cliente + ' ' + V_PERSONA_.CUIT + ' ' + ud_ezi_puntodeventa_encabezado.vendedor " & _
          "AS concatenado " & _
          "FROM V_PERSONA_ RIGHT OUTER JOIN " & _
          "V_CLIENTE_ ON V_PERSONA_.ID = V_CLIENTE_.ENTEASOCIADO_ID RIGHT OUTER JOIN " & _
          "ud_ezi_puntodeventa_encabezado WITH (readpast) ON V_CLIENTE_.ID = ud_ezi_puntodeventa_encabezado.clienteid " & _
          "WHERE     (ud_ezi_puntodeventa_encabezado.numeradorinterno = 'Nota de Venta') AND (ud_ezi_puntodeventa_encabezado.generada = 0) AND " & _
          "(ud_ezi_puntodeventa_encabezado.tipodepagoid <> '{E88275FE-0A80-11D6-B0D1-004854841C8A}') AND (ud_ezi_puntodeventa_encabezado.importeglobal <> 0) AND (ud_ezi_puntodeventa_encabezado.sucursal = '" & xsuc & "') " & _
          "ORDER BY Fecha DESC "


datremitos.RecordSource = xquery1
datremitos.Refresh
datctacte.RecordSource = xquery2
datctacte.Refresh

            DataGrid1.Columns(1).Width = 1000
            DataGrid1.Columns(2).Width = 500
            DataGrid1.Columns(2).Caption = "T.Fac"
            DataGrid1.Columns(3).Width = 3500
            DataGrid1.Columns(4).Width = 1200
            DataGrid1.Columns(6).Alignment = dbgRight
            DataGrid1.Columns(6).NumberFormat = "Currency"
            DataGrid1.Columns(7).Visible = False
            DataGrid1.Columns(8).Visible = False
            DataGrid1.Columns(9).Visible = False

            DataGrid3.Columns(1).Width = 1000
            DataGrid3.Columns(2).Width = 500
            DataGrid3.Columns(2).Caption = "T.Fac"
            DataGrid3.Columns(3).Width = 3500
            DataGrid3.Columns(4).Width = 1200
            DataGrid3.Columns(6).Alignment = dbgRight
            DataGrid3.Columns(6).NumberFormat = "Currency"
            DataGrid3.Columns(7).Visible = False
            DataGrid3.Columns(8).Visible = False
            DataGrid3.Columns(9).Visible = False





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

Private Sub KewlButtons1_Click()
On Error Resume Next

xquery = "SELECT   distinct  N.id, N.numeradorinterno, N.clienteid, N.cliente, N.vendedorid, N.vendedor, N.detalle, N.nota, N.mediodetransporteid, N.monedaid, N.cotizacion, N.listadeprecioid, " & _
         "N.tipodepagoid, N.tipodefacturacionid, N.tipodeentregaid, N.fechadeentrega, N.recargo, N.tiporecargo, N.bonificacion, N.tipobonificacion, N.importeglobal, " & _
         "N.numerodefactura, N.tipodefactura, N.numerador, N.domicilioid, N.domiciliodeentregaid, N.puntodeventa, N.descuento, N.subtotalsiniva, N.totaliva, N.generada, " & _
         "N.presupuestobase, N.importado, N.comprobanteorigen, N.totaltr, N.percepiibb, N.formapago, N.target, N.nombrepc, N.nrofactinicial, N.nrofactfinal, N.responsabilidad, " & _
         "N.nrofactoriginante, N.factoriginanteid, N.estado, N.transferido, N.calipsoid, N.osmpmutualid, N.recetaid, N.medicoid, N.adicionalid, N.reconocido, N.facturar, N.senia, " & _
         "N.nroorden, N.fechaorden, N.estadoimpresion, N.imprimir, N.perceptem, N.perceppyp, R.id AS idremito, R.NUMERODOCUMENTO, N.claveprimaria, " & _
         "N.fechadelcomprobante , N.sucursal, N.obraid, N.retira FROM  ud_ezi_puntodeventa_encabezado AS N LEFT OUTER JOIN v_ezi_pos_remito AS R ON N.id = R.presupuestobase " & _
         "WHERE (N.numeradorinterno = 'Nota de Venta') and  N.id ='" & DataGrid1.Columns(7).Text & "' "

    query = xquery
    If query = "" Then
        MsgBox "Seleccione nuevamente la Nota de Venta a Facturar", vbInformation, "Atención"
        Exit Sub
    End If
    
    remdev = 0
    frmnota_venta_cobrar.Show
    frmnota_venta_cobrar.Text17.SetFocus
    SendKeys "{ENTER}", False
    
    

End Sub





Private Sub KewlButtons2_Click()

    mensa = MsgBox("Desea anular esta Nota de Venta ? ", vbYesNo, "Atención !!")
    If mensa = vbYes Then
        datcontrol.ConnectionString = login.conexiontotal
        datcontrol.RecordSource = "select * from ud_ezi_puntodeventa_encabezado where id = '" & DataGrid1.Columns("id").Text & "'"
        datcontrol.Refresh
        If datcontrol.Recordset.EOF = False Then
            datcontrol.Recordset.Fields("generada") = 1
            datcontrol.Recordset.UpdateBatch adAffectCurrent
        End If
        Call Command4_Click
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

Private Sub KewlButtons3_Click()
    mensa = MsgBox("Desea anular esta Nota de Venta ? ", vbYesNo, "Atención !!")
    If mensa = vbYes Then
        datcontrol.ConnectionString = login.conexiontotal
        datcontrol.RecordSource = "select * from ud_ezi_puntodeventa_encabezado where id = '" & DataGrid3.Columns("id").Text & "'"
        datcontrol.Refresh
        If datcontrol.Recordset.EOF = False Then
            datcontrol.Recordset.Fields("generada") = 1
            datcontrol.Recordset.UpdateBatch adAffectCurrent
        End If
        Call Command4_Click
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
              '  xquery1 = "SELECT id, Sel, NroRemito, Fecha, CodCliente, Cliente, CUIT,MAX(NroFactura) AS NroFactura, Tipopago, TipodeVenta, presupuestobase,  Vendedor, MAX(cantremitida) " & _
              '            "AS cantremitida, SUM(cantfacturada) AS cantfacturada, MAX(concatenado) AS concatenado, saldado " & _
              '            "FROM (SELECT     R.id, R.Sel, R.NUMERODOCUMENTO AS NroRemito, R.FECHAEMISION AS Fecha, R.cod_cliente AS CodCliente, R.cliente AS Cliente, R.CUIT, " & _
              '            "R.vendedor AS Vendedor, R.tipopago AS Tipopago, R.TIPOVENTA AS TipodeVenta, R.presupuestobase, " & _
              '            "E.tipodefactura + ' ' + E.puntodeventa + RIGHT('0000000' + E.numerodefactura, 8) AS NroFactura, SUM(R.cantidadremitida) AS cantremitida, " & _
              '            "SUM(ISNULL(IFAC.cantidadproducto, 0)) AS cantfacturada, " & _
              '            "R.NUMERODOCUMENTO + 'r  ' + R.cliente + ' ' + R.CUIT + ' ' + ISNULL(E.tipodefactura + ' ' + E.puntodeventa + RIGHT('0000000' + E.numerodefactura, 8), '') + 'f' AS concatenado, r.saldado " & _
              '            "FROM ud_ezi_puntodeventa_detalle_factm AS IFAC WITH (readpast) LEFT OUTER JOIN ud_ezi_puntodeventa_encabezado AS E WITH (readpast) ON IFAC.claveprimaria = E.id RIGHT OUTER JOIN " & _
              '            "v_ezi_pos_remito AS R ON IFAC.iditemremito = R.iditem " & _
              '            "GROUP BY R.id, R.Sel, R.NUMERODOCUMENTO, R.FECHAEMISION, R.cod_cliente, R.cliente, R.CUIT, R.vendedor, R.tipopago, R.TIPOVENTA, " & _
              '            "E.tipodefactura + ' ' + E.puntodeventa + RIGHT('0000000' + E.numerodefactura, 8), R.presupuestobase, R.alquiler, r.saldado " & _
              '            "HAVING (R.alquiler ='N') and (R.NUMERODOCUMENTO + 'r  ' + R.cliente + ' ' + R.CUIT + ' ' + ISNULL(E.tipodefactura + ' ' + E.puntodeventa + RIGHT('0000000' + E.numerodefactura, 8), '') + 'f' LIKE '" & xbusqueda & "')) AS RC " & _
              '            "GROUP BY id, Sel, NroRemito, Fecha, CodCliente, Cliente, CUIT, Vendedor, Tipopago, TipodeVenta, presupuestobase, saldado " & _
              '            "Having (Max(cantremitida) > SUM(cantfacturada)) and (saldado IS NULL) ORDER BY id DESC"
                          
                 xquery1 = "SELECT     ud_ezi_puntodeventa_encabezado.fechadelcomprobante AS Fecha, ud_ezi_puntodeventa_encabezado.numerodefactura AS Numero, " & _
                          "ud_ezi_puntodeventa_encabezado.tipodefactura, ud_ezi_puntodeventa_encabezado.cliente AS Cliente, V_PERSONA_.CUIT, " & _
                          "ud_ezi_puntodeventa_encabezado.vendedor AS Vendedor, ud_ezi_puntodeventa_encabezado.importeglobal AS Importe, ud_ezi_puntodeventa_encabezado.id, " & _
                          "ud_ezi_puntodeventa_encabezado.generada, " & _
                          "ud_ezi_puntodeventa_encabezado.numerodefactura + 'n ' + ud_ezi_puntodeventa_encabezado.cliente + ' ' + V_PERSONA_.CUIT + ' ' + ud_ezi_puntodeventa_encabezado.vendedor " & _
                          "AS concatenado " & _
                          "FROM V_PERSONA_ RIGHT OUTER JOIN " & _
                          "V_CLIENTE_ ON V_PERSONA_.ID = V_CLIENTE_.ENTEASOCIADO_ID RIGHT OUTER JOIN " & _
                          "ud_ezi_puntodeventa_encabezado WITH (readpast) ON V_CLIENTE_.ID = ud_ezi_puntodeventa_encabezado.clienteid " & _
                          "WHERE     (ud_ezi_puntodeventa_encabezado.numeradorinterno = 'Nota de Venta') AND (ud_ezi_puntodeventa_encabezado.generada = 0) AND " & _
                          "(ud_ezi_puntodeventa_encabezado.tipodepagoid <> '{E88275FE-0A80-11D6-B0D1-004854841C8A}') AND (ud_ezi_puntodeventa_encabezado.importeglobal <> 0)  " & _
                          "and (ud_ezi_puntodeventa_encabezado.cliente + ' ' + ISNULL(ud_ezi_puntodeventa_encabezado.tipodefactura + ' ' + isnull(ud_ezi_puntodeventa_encabezado.puntodeventa,'') + RIGHT('0000000' + ud_ezi_puntodeventa_encabezado.numerodefactura, 8), '') + 'f' LIKE '" & xbusqueda & "')" & _
                          "ORDER BY Fecha DESC "
                         
            Else
              '  xquery1 = "SELECT id, Sel, NroRemito, Fecha, CodCliente, Cliente, CUIT,MAX(NroFactura) AS NroFactura , Tipopago, TipodeVenta, presupuestobase, Vendedor, MAX(cantremitida) " & _
              '            "AS cantremitida, SUM(cantfacturada) AS cantfacturada, MAX(concatenado) AS concatenado " & _
              '            "FROM (SELECT     R.id, R.Sel, R.NUMERODOCUMENTO AS NroRemito, R.FECHAEMISION AS Fecha, R.cod_cliente AS CodCliente, R.cliente AS Cliente, R.CUIT, " & _
              '            "R.vendedor AS Vendedor, R.tipopago AS Tipopago, R.TIPOVENTA AS TipodeVenta, R.presupuestobase, " & _
              '            "E.tipodefactura + ' ' + E.puntodeventa + RIGHT('0000000' + E.numerodefactura, 8) AS NroFactura, SUM(R.cantidadremitida) AS cantremitida, " & _
              '            "SUM(ISNULL(IFAC.cantidadproducto, 0)) AS cantfacturada, " & _
              '            "R.NUMERODOCUMENTO + 'r  ' + R.cliente + ' ' + R.CUIT + ' ' + ISNULL(E.tipodefactura + ' ' + E.puntodeventa + RIGHT('0000000' + E.numerodefactura, 8), '') + 'f' AS concatenado " & _
              '            "FROM ud_ezi_puntodeventa_detalle_factm AS IFAC WITH (readpast) LEFT OUTER JOIN ud_ezi_puntodeventa_encabezado AS E WITH (readpast) ON IFAC.claveprimaria = E.id RIGHT OUTER JOIN " & _
              '            "v_ezi_pos_remito AS R ON IFAC.iditemremito = R.iditem " & _
              '            "GROUP BY R.id, R.Sel, R.NUMERODOCUMENTO, R.FECHAEMISION, R.cod_cliente, R.cliente, R.CUIT, R.vendedor, R.tipopago, R.TIPOVENTA, " & _
              '            "E.tipodefactura + ' ' + E.puntodeventa + RIGHT('0000000' + E.numerodefactura, 8), R.presupuestobase, R.alquiler " & _
              '            "HAVING (R.alquiler = 'N') and (R.NUMERODOCUMENTO + 'r  ' + R.cliente + ' ' + R.CUIT + ' ' + ISNULL(E.tipodefactura + ' ' + E.puntodeventa + RIGHT('0000000' + E.numerodefactura, 8), '') + 'f' LIKE '" & xbusqueda & "')) AS RC " & _
              '            "GROUP BY id, Sel, NroRemito, Fecha, CodCliente, Cliente, CUIT, Vendedor, Tipopago, TipodeVenta, presupuestobase " & _
              '            "ORDER BY id DESC"

                xquery1 = "SELECT     ud_ezi_puntodeventa_encabezado.fechadelcomprobante AS Fecha, ud_ezi_puntodeventa_encabezado.numerodefactura AS Numero, " & _
                          "ud_ezi_puntodeventa_encabezado.tipodefactura, ud_ezi_puntodeventa_encabezado.cliente AS Cliente, V_PERSONA_.CUIT, " & _
                          "ud_ezi_puntodeventa_encabezado.vendedor AS Vendedor, ud_ezi_puntodeventa_encabezado.importeglobal AS Importe, ud_ezi_puntodeventa_encabezado.id, " & _
                          "ud_ezi_puntodeventa_encabezado.generada, " & _
                          "ud_ezi_puntodeventa_encabezado.numerodefactura + 'n ' + ud_ezi_puntodeventa_encabezado.cliente + ' ' + V_PERSONA_.CUIT + ' ' + ud_ezi_puntodeventa_encabezado.vendedor " & _
                          "AS concatenado " & _
                          "FROM V_PERSONA_ RIGHT OUTER JOIN " & _
                          "V_CLIENTE_ ON V_PERSONA_.ID = V_CLIENTE_.ENTEASOCIADO_ID RIGHT OUTER JOIN " & _
                          "ud_ezi_puntodeventa_encabezado WITH (readpast) ON V_CLIENTE_.ID = ud_ezi_puntodeventa_encabezado.clienteid " & _
                          "WHERE     (ud_ezi_puntodeventa_encabezado.numeradorinterno = 'Nota de Venta') AND (ud_ezi_puntodeventa_encabezado.generada = 0) AND " & _
                          "(ud_ezi_puntodeventa_encabezado.tipodepagoid <> '{E88275FE-0A80-11D6-B0D1-004854841C8A}') AND (ud_ezi_puntodeventa_encabezado.importeglobal <> 0) AND (ud_ezi_puntodeventa_encabezado.sucursal = '" & xsuc & "') " & _
                          "and (ud_ezi_puntodeventa_encabezado.cliente + ' ' + ISNULL(ud_ezi_puntodeventa_encabezado.tipodefactura + ' ' + isnull(ud_ezi_puntodeventa_encabezado.puntodeventa,'') + RIGHT('0000000' + ud_ezi_puntodeventa_encabezado.numerodefactura, 8), '') + 'f' LIKE '" & xbusqueda & "')" & _
                          "ORDER BY Fecha DESC "
            
            End If
                    
'            datremitos.RecordSource = xquery1
'            datremitos.Refresh
            datctacte.RecordSource = xquery1
            datctacte.Refresh
            If Text1.Text = " " Then Text1.Text = ""
            DataGrid1.Columns(1).Width = 1000
            DataGrid1.Columns(2).Width = 500
            DataGrid1.Columns(2).Caption = "T.Fac"
            DataGrid1.Columns(3).Width = 3500
            DataGrid1.Columns(4).Width = 1200
            DataGrid1.Columns(6).Alignment = dbgRight
            DataGrid1.Columns(6).NumberFormat = "Currency"
            DataGrid1.Columns(7).Visible = False
            DataGrid1.Columns(8).Visible = False
            DataGrid1.Columns(9).Visible = False

            DataGrid3.Columns(1).Width = 1000
            DataGrid3.Columns(2).Width = 500
            DataGrid3.Columns(2).Caption = "T.Fac"
            DataGrid3.Columns(3).Width = 3500
            DataGrid3.Columns(4).Width = 1200
            DataGrid3.Columns(6).Alignment = dbgRight
            DataGrid3.Columns(6).NumberFormat = "Currency"
            DataGrid3.Columns(7).Visible = False
            DataGrid3.Columns(8).Visible = False
            DataGrid3.Columns(9).Visible = False

            
            
        End If
        DataGrid1.SetFocus
        
        
    End If


End Sub
