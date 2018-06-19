VERSION 5.00
Object = "{F6125AB1-8AB1-11CE-A77F-08002B2F4E98}#2.0#0"; "MSRDC20.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{D18BBD1F-82BB-4385-BED3-E9D31A3E361E}#1.0#0"; "KewlButtonz.ocx"
Begin VB.Form frmremitosconsulta_remito 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta de Remitos Emitidos"
   ClientHeight    =   8595
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15195
   Icon            =   "frmremitosconsulta_remito.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8595
   ScaleWidth      =   15195
   Begin MSAdodcLib.Adodc datcolaimportar 
      Height          =   330
      Left            =   9480
      Top             =   2160
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
      Left            =   9000
      Top             =   1800
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
      Height          =   1455
      Left            =   8160
      TabIndex        =   11
      Top             =   120
      Width           =   1695
      Begin VB.OptionButton Option5 
         Caption         =   "Todos"
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   840
         Width           =   1215
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Pend.de Fact"
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   360
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
      Top             =   1560
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
      Bindings        =   "frmremitosconsulta_remito.frx":0442
      Height          =   3855
      Left            =   120
      TabIndex        =   4
      Top             =   1680
      Width           =   14895
      _ExtentX        =   26273
      _ExtentY        =   6800
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
      Height          =   1455
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   9735
      Begin VB.CommandButton Command2 
         Caption         =   "Desde Fecha:"
         Height          =   375
         Left            =   120
         TabIndex        =   18
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Hasta Fecha:"
         Height          =   375
         Left            =   3120
         TabIndex        =   17
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Filtrar"
         Height          =   375
         Left            =   6120
         TabIndex        =   16
         Top             =   360
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
         Top             =   840
         Width           =   7575
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Triplicado"
         Height          =   195
         Left            =   6360
         TabIndex        =   3
         Top             =   600
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Duplicado"
         Height          =   195
         Left            =   6360
         TabIndex        =   2
         Top             =   360
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Original"
         Height          =   195
         Left            =   6360
         TabIndex        =   1
         Top             =   120
         Visible         =   0   'False
         Width           =   1335
      End
      Begin MSComCtl2.DTPicker DesdeFecha 
         Height          =   375
         Left            =   1440
         TabIndex        =   19
         Top             =   360
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   98566145
         CurrentDate     =   42198
      End
      Begin MSComCtl2.DTPicker HastaFecha 
         Height          =   375
         Left            =   4440
         TabIndex        =   20
         Top             =   360
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   98566145
         CurrentDate     =   42198
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1455
      Left            =   9960
      TabIndex        =   9
      Top             =   120
      Width           =   5055
      Begin KewlButtonz.KewlButtons Command4 
         Height          =   495
         Left            =   2760
         TabIndex        =   5
         Top             =   840
         Visible         =   0   'False
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   873
         BTYPE           =   14
         TX              =   "&Listar"
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
         MICON           =   "frmremitosconsulta_remito.frx":045B
         PICN            =   "frmremitosconsulta_remito.frx":0477
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
         Left            =   3720
         TabIndex        =   6
         Top             =   120
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
         MICON           =   "frmremitosconsulta_remito.frx":3869
         PICN            =   "frmremitosconsulta_remito.frx":3885
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin KewlButtonz.KewlButtons cancelar 
         Height          =   495
         Left            =   240
         TabIndex        =   21
         Top             =   240
         Width           =   2295
         _ExtentX        =   4048
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
         MICON           =   "frmremitosconsulta_remito.frx":43CF
         PICN            =   "frmremitosconsulta_remito.frx":43EB
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin KewlButtonz.KewlButtons reimprimir 
         Height          =   495
         Left            =   240
         TabIndex        =   22
         Top             =   840
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   873
         BTYPE           =   14
         TX              =   "Reimprimir con Nuevo Numero"
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
         MICON           =   "frmremitosconsulta_remito.frx":4DFD
         PICN            =   "frmremitosconsulta_remito.frx":4E19
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
      Bindings        =   "frmremitosconsulta_remito.frx":689B
      Height          =   2775
      Left            =   120
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   5640
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
      TabIndex        =   14
      Top             =   8040
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
      MICON           =   "frmremitosconsulta_remito.frx":68B7
      PICN            =   "frmremitosconsulta_remito.frx":68D3
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
      TabIndex        =   15
      Top             =   8040
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
      MICON           =   "frmremitosconsulta_remito.frx":6E6D
      PICN            =   "frmremitosconsulta_remito.frx":6E89
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
Attribute VB_Name = "frmremitosconsulta_remito"
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


End Sub

Private Sub Cancelar_Click()
On Error Resume Next

If DataGrid1.Columns("nrofactura") <> "" Then
    If UCase(login.usuarioactivo) <> "ADMIN" Then
        MsgBox "Este remito ya tiene una factura Generada, debe realizar la NOTA DE CREDITO sobre la factura para liberar el remito", vbCritical, "Error"
        Exit Sub
    End If
End If

    mensa = MsgBox("Esta Seguro de Anular el Remito: " + DataGrid1.Columns("NroRemito").Text, vbYesNo, "Anulación")
    
    If mensa = vbYes Then
        xidremitoanular = DataGrid1.Columns("id")
        xidcalipso = DataGrid1.Columns("calipsoid")
        
        mensa2 = MsgBox("Mantiene el armado Realizado", vbYesNo, "Atención")

''' Elimina el remito generado y genera en cola anulacion en calipco
        If mensa2 = vbNo Then
            datsaldar.RecordSource = "select * from ud_ezi_puntodeventa_encabezado where id = " & xidremitoanular & ""
            datsaldar.Refresh
            If datsaldar.Recordset.EOF = False Then
                xclaveprimaria = datsaldar.Recordset.Fields("claveprimaria")
                
                datsaldar2.RecordSource = "select * from ud_ezi_puntodeventa_detalle_rem where claveprimaria = " & xclaveprimaria & " "
                datsaldar2.Refresh
                If datsaldar2.Recordset.EOF = False Then
                    datsaldar2.Recordset.MoveFirst
                    Do While Not datsaldar2.Recordset.EOF
                        datsaldar2.Recordset.Delete adAffectCurrent
                        datsaldar2.Recordset.MoveNext
                    Loop
                End If
            
                datsaldar.Recordset.Delete adAffectCurrent
            End If
            
                datcolaimportar.RecordSource = "Select * from ud_ezi_cola_importar where id = 0"
                datcolaimportar.Refresh
        
                datcolaimportar.Recordset.AddNew
                datcolaimportar.Recordset.Fields("id_encabezado") = xidremitoanular
                datcolaimportar.Recordset.Fields("tipodedocumentoid") = "AnulacionRemito"
                datcolaimportar.Recordset.Fields("unidadoperativaid") = xidcalipso
                datcolaimportar.Recordset.Fields("fecha_hora") = DateValue(Text3.Text) + TimeValue(Str(Time))
        
                datcolaimportar.Recordset.UpdateBatch adAffectCurrent
    
            
                MsgBox "Remito Anulado Correctamente"
                
                
        End If
                
        If mensa2 = vbYes Then
        
            datsaldar.RecordSource = "select * from ud_ezi_puntodeventa_encabezado where id = " & xidremitoanular & ""
            datsaldar.Refresh
            If datsaldar.Recordset.EOF = False Then
                xclaveprimaria = datsaldar.Recordset.Fields("claveprimaria")
                datsaldar.Recordset.Fields("numerodefactura") = xclaveprimaria
                datsaldar.Recordset.Fields("puntodeventa") = "99"
                datsaldar.Recordset.Fields("estado") = "Preparado"
                datsaldar.Recordset.Fields("calipsoid") = Null
                
                datsaldar2.RecordSource = "select * from ud_ezi_puntodeventa_detalle_rem where claveprimaria = " & xclaveprimaria & " "
                datsaldar2.Refresh
                If datsaldar2.Recordset.EOF = False Then
                    datsaldar2.Recordset.MoveFirst
                    Do While Not datsaldar2.Recordset.EOF
                        datsaldar2.Recordset.Fields("cantidadremitida") = 0
                        datsaldar2.Recordset.UpdateBatch adAffectCurrent
                        datsaldar2.Recordset.MoveNext
                    Loop
                End If
            
                datsaldar.Recordset.UpdateBatch adAffectCurrent
            End If
      
                datcolaimportar.RecordSource = "Select * from ud_ezi_cola_importar where id = 0"
                datcolaimportar.Refresh
        
                datcolaimportar.Recordset.AddNew
                datcolaimportar.Recordset.Fields("id_encabezado") = xidremitoanular
                datcolaimportar.Recordset.Fields("tipodedocumentoid") = "AnulacionRemito"
                datcolaimportar.Recordset.Fields("unidadoperativaid") = xidcalipso
                datcolaimportar.Recordset.Fields("fecha_hora") = DateValue(Text3.Text) + TimeValue(Str(Time))
        
                datcolaimportar.Recordset.UpdateBatch adAffectCurrent
                
                MsgBox "Remito Anulado Correctamente, Armado Restaurado"
        
        End If
        
    End If
        


End Sub

Private Sub Command4_Click()
On Error GoTo fuera
Dim crApp As New CRAXDRT.Application
Dim crReport As New CRAXDRT.Report
Dim tabla As String
Dim ruta As String

If Text1.Text = "" Then
 reporte.SQL = "SELECT     Numero, tipodefactura, fechadelcomprobante, numeradorinterno, codcliente, cliente, totaltr, NroRemito " & _
              "FROM         MMOSSE.dbo.v_ezi_pos_listadoremitos AS v_ezi_pos_listadoremitos " & _
              "WHERE     convert(date,v_ezi_pos_listadoremitos.fechadelcomprobante) >= '" & DesdeFecha.Value & "' and " & _
              "convert(date,v_ezi_pos_listadoremitos.fechadelcomprobante) <= '" & HastaFecha.Value & "' and " & _
              "v_ezi_pos_listadoremitos.sucursal = '" & login.nomsucursal & "' " & _
              "order by fechadelcomprobante desc"
Else
 reporte.SQL = "SELECT     Numero, tipodefactura, fechadelcomprobante, numeradorinterno, codcliente, cliente, totaltr, NroRemito " & _
              "FROM         MMOSSE.dbo.v_ezi_pos_listadoremitos AS v_ezi_pos_listadoremitos " & _
              "WHERE     convert(date,v_ezi_pos_listadoremitos.fechadelcomprobante) >= '" & DesdeFecha.Value & "' and " & _
              "convert(date,v_ezi_pos_listadoremitos.fechadelcomprobante) <= '" & HastaFecha.Value & "' and " & _
              "v_ezi_pos_listadoremitos.sucursal = '" & login.nomsucursal & "' and " & _
              "v_ezi_pos_listadoremitos.cliente = '" & DataGrid1.Columns("cliente").Text & "' " & _
              "order by fechadelcomprobante desc"
End If

tabla = reporte.SQL

With CrystalReporte
    .PrinterCollation = crptCollated
    .ReportFileName = App.Path & "\Reporte_listado_remitos.rpt"
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



Private Sub DataCombo1_KeyPress(KeyAscii As Integer)
On Error GoTo fueraderango
    If KeyAscii = 13 Then
        KeyAscii = 0
        List1.ListIndex = DataCombo1.SelectedItem - 1
        Call Command4_Click
    End If
fueraderango:
End Sub

Private Sub Command5_Click()
On Error Resume Next
xsuc = login.nomsucursal
xfechadesde = Replace(Str(Year(DesdeFecha.Value)), " ", "") + "-" + Right("0" + Replace(Str(Month(DesdeFecha.Value)), " ", ""), 2) + "-" + Right("0" + Replace(Str(Day(DesdeFecha.Value)), " ", ""), 2)
xfechahasta = Replace(Str(Year(HastaFecha.Value)), " ", "") + "-" + Right("0" + Replace(Str(Month(HastaFecha.Value)), " ", ""), 2) + "-" + Right("0" + Replace(Str(Day(HastaFecha.Value)), " ", ""), 2)


        If Text1.Text <> "" Then
            Text1.Text = Replace(Text1.Text, " ", "%%")
            xbusqueda = "%" + Text1.Text + "%"
         If Option4.Value = True Then
                xquery1 = "SELECT     id, Sel, NroRemito, Fecha, presupuestobase AS NV, CodCliente, Cliente, CUIT, MAX(isnull(NroFactura, nfactura)) AS NroFactura, Tipopago, TipodeVenta, Vendedor, MAX(cantremitida) " & _
                          "AS cantremitida, SUM(cantfacturada) AS cantfacturada, MAX(concatenado) AS concatenado, saldado,calipsoid " & _
                          "FROM         (SELECT     R.id, R.Sel, R.NUMERODOCUMENTO AS NroRemito, R.FECHAEMISION AS Fecha, R.cod_cliente AS CodCliente, R.cliente AS Cliente, R.CUIT, " & _
                          "R.vendedor AS Vendedor, R.tipopago AS Tipopago, R.TIPOVENTA AS TipodeVenta, R.presupuestobase, " & _
                          "E.tipodefactura + ' ' + E.puntodeventa + RIGHT('0000000' + E.numerodefactura, 8) AS NroFactura, SUM(R.cantidadremitida) AS cantremitida, " & _
                          "SUM(ISNULL(IFAC.cantidadproducto, 0)) + ISNULL(FA.cantidadfacturada2, 0) AS cantfacturada, " & _
                          "R.NUMERODOCUMENTO + 'r  ' + R.cliente + ' ' + R.CUIT + ' ' + ISNULL(E.tipodefactura + ' ' + E.puntodeventa + RIGHT('0000000' + E.numerodefactura, 8), '') " & _
                          "+ 'f' AS concatenado, R.saldado, E.numeradorinterno, FA.cantidadfacturada2, FA.NroFactura as nfactura,r.calipsoid " & _
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
                          "FA.cantidadfacturada2 , R.Estado, FA.NroFactura,r.calipsoid " & _
                          "HAVING      (R.alquiler = 'N') AND (R.NUMERODOCUMENTO + 'r  ' + R.cliente + ' ' + R.CUIT + ' ' + ISNULL(E.tipodefactura + ' ' + E.puntodeventa + RIGHT('0000000' + E.numerodefactura, 8), '') + 'f' LIKE '" & xbusqueda & "') AND (E.numeradorinterno LIKE '%Factura%' OR  " & _
                          "E.numeradorinterno LIKE '%Nota%') AND (R.estado = 'Remitido') and (R.FECHAEMISION >= CONVERT(date, '" & xfechadesde & "', 102)) AND (R.FECHAEMISION <= CONVERT(date, '" & xfechahasta & "', 102)) ) AS RC " & _
                          "GROUP BY id, Sel, NroRemito, Fecha, CodCliente, Cliente, CUIT, Vendedor, Tipopago, TipodeVenta, presupuestobase, saldado, calipsoid " & _
                          "Having (saldado Is Null) And (Sum(cantfacturada) < Max(cantremitida)) ORDER BY id DESC"

         Else
                         
                xquery1 = "SELECT     id, Sel, NroRemito, Fecha, presupuestobase AS NV, CodCliente, Cliente, CUIT, case when LEFT( MIN(ISNULL(NroFactura, nfactura)),1) = 'R' then null else MIN(ISNULL(NroFactura, nfactura)) end  AS NroFactura, Tipopago, TipodeVenta, Vendedor, MAX(cantremitida) " & _
                          "AS cantremitida, SUM(cantfacturada) AS cantfacturada, MAX(concatenado) AS concatenado, saldado, calipsoid  " & _
                          "FROM         (SELECT     R.id, R.Sel, R.NUMERODOCUMENTO AS NroRemito, R.FECHAEMISION AS Fecha, R.cod_cliente AS CodCliente, R.cliente AS Cliente, R.CUIT, " & _
                          "R.vendedor AS Vendedor, R.tipopago AS Tipopago, R.TIPOVENTA AS TipodeVenta, R.presupuestobase, " & _
                          "E.tipodefactura + ' ' + E.puntodeventa + RIGHT('0000000' + E.numerodefactura, 8) AS NroFactura, SUM(R.cantidadremitida) AS cantremitida, " & _
                          "SUM(ISNULL(IFAC.cantidadproducto, 0)) + ISNULL(FA.cantidadfacturada2, 0) AS cantfacturada, " & _
                          "R.NUMERODOCUMENTO + 'r  ' + R.cliente + ' ' + R.CUIT + ' ' + ISNULL(E.tipodefactura + ' ' + E.puntodeventa + RIGHT('0000000' + E.numerodefactura, 8), '') " & _
                          "+ 'f' AS concatenado, R.saldado, E.numeradorinterno, FA.cantidadfacturada2, FA.NroFactura as nfactura,r.calipsoid " & _
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
                          "FA.cantidadfacturada2 , R.Estado, FA.NroFactura, r.calipsoid " & _
                          "HAVING      (R.alquiler = 'N') AND (R.NUMERODOCUMENTO + 'r  ' + R.cliente + ' ' + R.CUIT + ' ' + ISNULL(E.tipodefactura + ' ' + E.puntodeventa + RIGHT('0000000' + E.numerodefactura, 8), '') + 'f' LIKE '" & xbusqueda & "') " & _
                          "AND (R.estado = 'Remitido') and (R.FECHAEMISION >= CONVERT(date, '" & xfechadesde & "', 102)) AND (R.FECHAEMISION <= CONVERT(date, '" & xfechahasta & "', 102)) ) AS RC " & _
                          "GROUP BY id, Sel, NroRemito, Fecha, CodCliente, Cliente, CUIT, Vendedor, Tipopago, TipodeVenta, presupuestobase, saldado, calipsoid " & _
                          "ORDER BY id DESC"
         End If
        Else
         If Option4.Value = True Then
            xquery1 = "SELECT     id, Sel, NroRemito, Fecha, presupuestobase AS NV, CodCliente, Cliente, CUIT, MAX(isnull(NroFactura, nfactura)) AS NroFactura, Tipopago, TipodeVenta, Vendedor, MAX(cantremitida) " & _
                          "AS cantremitida, SUM(cantfacturada) AS cantfacturada, MAX(concatenado) AS concatenado, saldado, calipsoid " & _
                          "FROM         (SELECT     R.id, R.Sel, R.NUMERODOCUMENTO AS NroRemito, R.FECHAEMISION AS Fecha, R.cod_cliente AS CodCliente, R.cliente AS Cliente, R.CUIT, " & _
                          "R.vendedor AS Vendedor, R.tipopago AS Tipopago, R.TIPOVENTA AS TipodeVenta, R.presupuestobase, " & _
                          "E.tipodefactura + ' ' + E.puntodeventa + RIGHT('0000000' + E.numerodefactura, 8) AS NroFactura, SUM(R.cantidadremitida) AS cantremitida, " & _
                          "SUM(ISNULL(IFAC.cantidadproducto, 0)) + ISNULL(FA.cantidadfacturada2, 0) AS cantfacturada, " & _
                          "R.NUMERODOCUMENTO + 'r  ' + R.cliente + ' ' + R.CUIT + ' ' + ISNULL(E.tipodefactura + ' ' + E.puntodeventa + RIGHT('0000000' + E.numerodefactura, 8), '') " & _
                          "+ 'f' AS concatenado, R.saldado, E.numeradorinterno, FA.cantidadfacturada2, FA.NroFactura as nfactura,r.calipsoid " & _
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
                          "FA.cantidadfacturada2 , R.Estado, FA.NroFactura,r.calipsoid " & _
                          "HAVING      (R.alquiler = 'N')  AND (E.numeradorinterno LIKE '%Factura%' OR  " & _
                          "E.numeradorinterno LIKE '%Nota%') AND (R.estado = 'Remitido') and (R.FECHAEMISION >= CONVERT(date, '" & xfechadesde & "', 102)) AND (R.FECHAEMISION <= CONVERT(date, '" & xfechahasta & "', 102)) ) AS RC " & _
                          "GROUP BY id, Sel, NroRemito, Fecha, CodCliente, Cliente, CUIT, Vendedor, Tipopago, TipodeVenta, presupuestobase, saldado, calipsoid " & _
                          "Having (saldado Is Null) And (Sum(cantfacturada) < Max(cantremitida)) ORDER BY id DESC"

         Else
                          
                xquery1 = "SELECT     id, Sel, NroRemito, Fecha, presupuestobase AS NV, CodCliente, Cliente, CUIT, case when LEFT( MIN(ISNULL(NroFactura, nfactura)),1) = 'R' then null else MIN(ISNULL(NroFactura, nfactura)) end  AS NroFactura, Tipopago, TipodeVenta, Vendedor, MAX(cantremitida) " & _
                          "AS cantremitida, SUM(cantfacturada) AS cantfacturada, MAX(concatenado) AS concatenado, saldado, calipsoid  " & _
                          "FROM         (SELECT     R.id, R.Sel, R.NUMERODOCUMENTO AS NroRemito, R.FECHAEMISION AS Fecha, R.cod_cliente AS CodCliente, R.cliente AS Cliente, R.CUIT, " & _
                          "R.vendedor AS Vendedor, R.tipopago AS Tipopago, R.TIPOVENTA AS TipodeVenta, R.presupuestobase, " & _
                          "E.tipodefactura + ' ' + E.puntodeventa + RIGHT('0000000' + E.numerodefactura, 8) AS NroFactura, SUM(R.cantidadremitida) AS cantremitida, " & _
                          "SUM(ISNULL(IFAC.cantidadproducto, 0)) + ISNULL(FA.cantidadfacturada2, 0) AS cantfacturada, " & _
                          "R.NUMERODOCUMENTO + 'r  ' + R.cliente + ' ' + R.CUIT + ' ' + ISNULL(E.tipodefactura + ' ' + E.puntodeventa + RIGHT('0000000' + E.numerodefactura, 8), '') " & _
                          "+ 'f' AS concatenado, R.saldado, E.numeradorinterno, FA.cantidadfacturada2, FA.NroFactura as nfactura,r.calipsoid " & _
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
                          "FA.cantidadfacturada2 , R.Estado, FA.NroFactura, r.calipsoid " & _
                          "HAVING      (R.alquiler = 'N') " & _
                          "AND (R.estado = 'Remitido') and (R.FECHAEMISION >= CONVERT(date, '" & xfechadesde & "', 102)) AND (R.FECHAEMISION <= CONVERT(date, '" & xfechahasta & "', 102)) ) AS RC " & _
                          "GROUP BY id, Sel, NroRemito, Fecha, CodCliente, Cliente, CUIT, Vendedor, Tipopago, TipodeVenta, presupuestobase, saldado, calipsoid " & _
                          "ORDER BY id DESC"
                          
            End If
        End If

                    
            datremitos.RecordSource = xquery1
            datremitos.Refresh
            If Text1.Text = " " Then Text1.Text = ""
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
            
            Call DataGrid1_Click
            
            DataGrid1.SetFocus
        
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


End Sub

Private Sub Form_Load()
On Error Resume Next
Aplicar_skin Me

yventana = Inicio.Height / 2 - 1000
xventana = Inicio.Width / 2

frmremitosconsulta_remito.Top = yventana - frmremitosconsulta_remito.Height / 2
frmremitosconsulta_remito.Left = xventana - frmremitosconsulta_remito.Width / 2


datremitos.ConnectionString = login.conexiontotal
datitemremito.ConnectionString = login.conexiontotal
datsaldar.ConnectionString = login.conexiontotal
datsaldar2.ConnectionString = login.conexiontotal
datcolaimportar.ConnectionString = login.conexiontotal

DesdeFecha.Value = Date - Day(Date) + 1
HastaFecha.Value = Date


xcon = 1
facturar.Visible = False

If UCase(login.usuarioactivo) = "ADMIN" Then
    saldaremito.Visible = True
Else
    saldaremito.Visible = False
End If

xsuc = login.nomsucursal
xfechadesde = Replace(Str(Year(DesdeFecha.Value)), " ", "") + "-" + Right("0" + Replace(Str(Month(DesdeFecha.Value)), " ", ""), 2) + "-" + Right("0" + Replace(Str(Day(DesdeFecha.Value)), " ", ""), 2)
xfechahasta = Replace(Str(Year(HastaFecha.Value)), " ", "") + "-" + Right("0" + Replace(Str(Month(HastaFecha.Value)), " ", ""), 2) + "-" + Right("0" + Replace(Str(Day(HastaFecha.Value)), " ", ""), 2)
xquery1 = "SELECT     id, Sel, NroRemito, Fecha, CodCliente, Cliente, CUIT, Vendedor, Tipopago, TipodeVenta, presupuestobase, MAX(NroFactura) AS NroFactura, MAX(cantremitida) " & _
          "AS cantremitida, SUM(cantfacturada) AS cantfacturada, MIN(concatenado) AS concatenado " & _
          "FROM (SELECT     R.id, R.Sel, R.NUMERODOCUMENTO AS NroRemito, R.FECHAEMISION AS Fecha, R.cod_cliente AS CodCliente, R.cliente AS Cliente, R.CUIT, " & _
          "R.vendedor AS Vendedor, R.tipopago AS Tipopago, R.TIPOVENTA AS TipodeVenta, R.presupuestobase, " & _
          "E.tipodefactura + ' ' + E.puntodeventa + RIGHT('0000000' + E.numerodefactura, 8) AS NroFactura, SUM(R.cantidadremitida) AS cantremitida, " & _
          "SUM(ISNULL(IFAC.cantidadproducto, 0)) AS cantfacturada, R.NUMERODOCUMENTO + 'r  ' + R.cliente + ' ' + R.CUIT + ' ' + ISNULL(E.tipodefactura + ' ' + E.puntodeventa + RIGHT('0000000' + E.numerodefactura, 8), '') + 'f' AS concatenado " & _
          "FROM ud_ezi_puntodeventa_detalle_factm AS IFAC WITH (readpast) LEFT OUTER JOIN ud_ezi_puntodeventa_encabezado AS E WITH (readpast) ON IFAC.claveprimaria = E.id RIGHT OUTER JOIN " & _
          "v_ezi_pos_remito AS R ON IFAC.iditemremito = R.iditem " & _
          "where E.sucursal = '" & xsuc & "' and (E.fechadelcomprobante >= convert(date, '" & xfechadesde & "', 102)) and " & _
          "(E.fechadelcomprobante <= convert(date,'" & xfechahasta & "', 102)) " & _
          "GROUP BY R.id, R.Sel, R.NUMERODOCUMENTO, R.FECHAEMISION, R.cod_cliente, R.cliente, R.CUIT, R.vendedor, R.tipopago, R.TIPOVENTA, " & _
          "E.tipodefactura + ' ' + E.puntodeventa + RIGHT('0000000' + E.numerodefactura, 8), R.presupuestobase, R.alquiler HAVING      (R.alquiler = 'N')" & _
          ") AS rem GROUP BY id, Sel, NroRemito, Fecha, CodCliente, Cliente, CUIT, Vendedor, Tipopago, TipodeVenta, presupuestobase ORDER BY id DESC"

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

Private Sub reimprimir_Click()
On Error Resume Next
Dim crApp As New CRAXDRT.Application
Dim crReport As New CRAXDRT.Report

Dim tabla As String
Dim ruta As String


    mensa = MsgBox("Esta por reimprimir un remito actual corrigiendo su numeración al proximo remito disponible, Esta Seguro ?", vbYesNo, "Atención")
    If mensa = vbYes Then
        xidremitoanular = DataGrid1.Columns("id")
        xidcalipso = DataGrid1.Columns("calipsoid")
        
        datsaldar.RecordSource = "select * from ud_ezi_puntodeventa_encabezado where id = " & xidremitoanular & ""
        datsaldar.Refresh
        If datsaldar.Recordset.EOF = False Then
    
    '** Establene numero de Remitos Manuales, y no Fiscales
            xnumerador = "Remito A (Vtas) TUCUMAN"
 
    
            datcolaimportar.RecordSource = "SELECT p.ID AS ID, p.NOMBRE AS puntero, n.NUMERO AS NUMERO, p.CARACTERESPREFIJO AS puntoventa, p.NUMERO_ID " & _
                                           "FROM V_NUMERADOR_ AS p LEFT OUTER JOIN V_NUMERO_ AS n ON p.NUMERO_ID = n.ID " & _
                                           "WHERE     (p.ACTIVESTATUS <> 2) AND (p.NOMBRE = '" & xnumerador & "') "
            datcolaimportar.Refresh
      
            datsaldar.Recordset.Fields("numerodefactura") = datcolaimportar.Recordset.Fields("numero")
            xnumero = datcolaimportar.Recordset.Fields("numero")
            xidnumero = datcolaimportar.Recordset.Fields("numero_id")
            datsaldar.Recordset.Fields("puntodeventa") = datcolaimportar.Recordset.Fields("puntoventa")
            datsaldar.Recordset.UpdateBatch adAffectCurrent
            
            datcolaimportar.RecordSource = "Select * from numero with(readpast) where id = '" & xidnumero & "'"
            datcolaimportar.Refresh
            datcolaimportar.Recordset.Fields("numero") = xnumero + 1
            datcolaimportar.Recordset.UpdateBatch adAffectCurrent
    '** Fin de asignacion de numero a Remtio
    '** Cambia Numero en Calipso
            xnumerocalipso = datsaldar.Recordset.Fields("puntodeventa") + Right("00000000" + datsaldar.Recordset.Fields("numerodefactura"), 8)
        
            datcolaimportar.RecordSource = "select * from tregresoinventario with (readpast) WHERE ID = '" & xidcalipso & "'"
            datcolaimportar.Refresh
            If datcolaimportar.Recordset.EOF = False Then
                    datcolaimportar.Recordset.Fields("numerodocumento") = xnumerocalipso
                    datcolaimportar.Recordset.Fields("nombre") = Left(datcolaimportar.Recordset.Fields("nombre"), 15) + xnumerocalipso + Right(datcolaimportar.Recordset.Fields("nombre"), 19)
                    datcolaimportar.Recordset.UpdateBatch adAffectCurrent
            End If
        
        End If
    '** Fin Cambia Numero en Calipso
        MsgBox "Numeración Cambiada Correctamente"
    
    ''' imprime remito
    
    reporte.SQL = "SELECT v_ezi_pos_remito.id, v_ezi_pos_remito.NUMERODOCUMENTO, v_ezi_pos_remito.FECHAEMISION, v_ezi_pos_remito.cod_cliente, v_ezi_pos_remito.cliente, v_ezi_pos_remito.CALLE, v_ezi_pos_remito.CODPOS, v_ezi_pos_remito.provincia, v_ezi_pos_remito.detalle, v_ezi_pos_remito.tipopago, v_ezi_pos_remito.referenciaproducto, v_ezi_pos_remito.nombre_producto, v_ezi_pos_remito.cantidadremitida, v_ezi_pos_remito.nota, v_ezi_pos_remito.condiva, v_ezi_pos_remito.ciudad, v_ezi_pos_remito.TIPOVENTA, v_ezi_pos_remito.SIMBOLO, v_ezi_pos_remito.iditem FROM MMOSSE.dbo.v_ezi_pos_remito v_ezi_pos_remito where v_ezi_pos_remito.id = " & xidremitoanular & " order by v_ezi_pos_remito.iditem"
    tabla = reporte.SQL

    With CrystalReporte
    .PrinterCollation = crptCollated
    .Formulas(0) = "copia="" ORIGINAL """
    .ReportFileName = App.Path & "\RemitoVta.rpt"
    .WindowTitle = "Remito Vta Orig"
    .Connect = login.conexionreporte
    .DiscardSavedData = True
    .RetrieveDataFiles
    .ReportSource = 0
    .SQLQuery = tabla
 Rem   .Destination = crptToWindow
    .Destination = crptToPrinter
    .PrintFileType = crptCrystal
    .WindowState = crptMaximized
    .Action = 1
    .WindowTitle = "Remito Vta Dupl"
    .Formulas(0) = "copia="" DUPLICADO """
    .Action = 1
    .WindowTitle = "Remito Vta Trip"
    .Formulas(0) = "copia="" TRIPLICADO """
    .Action = 1
    End With

    
    
    
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

xsuc = login.nomsucursal
xfechadesde = Replace(Str(Year(DesdeFecha.Value)), " ", "") + "-" + Right("0" + Replace(Str(Month(DesdeFecha.Value)), " ", ""), 2) + "-" + Right("0" + Replace(Str(Day(DesdeFecha.Value)), " ", ""), 2)
xfechahasta = Replace(Str(Year(HastaFecha.Value)), " ", "") + "-" + Right("0" + Replace(Str(Month(HastaFecha.Value)), " ", ""), 2) + "-" + Right("0" + Replace(Str(Day(HastaFecha.Value)), " ", ""), 2)


    If KeyAscii = 13 Then
        KeyAscii = 0
        Call Command5_Click
        
    End If


End Sub
