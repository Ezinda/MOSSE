VERSION 5.00
Object = "{F6125AB1-8AB1-11CE-A77F-08002B2F4E98}#2.0#0"; "MSRDC20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{D18BBD1F-82BB-4385-BED3-E9D31A3E361E}#1.0#0"; "KewlButtonz.ocx"
Object = "{50BF2256-701F-46F2-8ADB-2202CE6922BC}#1.0#0"; "KlexGrid.ocx"
Begin VB.Form lista_pendientesremitir 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pendientes de Remitir"
   ClientHeight    =   8745
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   18945
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8745
   ScaleWidth      =   18945
   WindowState     =   2  'Maximized
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid msgridlote 
      Height          =   3255
      Left            =   15720
      TabIndex        =   10
      Top             =   5160
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   5741
      _Version        =   393216
      Rows            =   21
      Cols            =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   3
   End
   Begin Grid.KlexGrid KlexGrid1 
      Height          =   3255
      Left            =   120
      TabIndex        =   8
      Top             =   5160
      Width           =   15495
      _ExtentX        =   27331
      _ExtentY        =   5741
      EnterKeyBehaviour=   0
      BackColorAlternate=   0
      GridLinesFixed  =   2
      BackColorFixed  =   -2147483626
      Cols            =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GridColorFixed  =   8421504
      MouseIcon       =   "lista_pendientesremitir.frx":0000
      Rows            =   10
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Lotes"
      Height          =   255
      Left            =   1080
      TabIndex        =   7
      Top             =   0
      Visible         =   0   'False
      Width           =   1695
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid msgrid1 
      Bindings        =   "lista_pendientesremitir.frx":001C
      Height          =   4455
      Left            =   120
      TabIndex        =   6
      Top             =   600
      Width           =   15855
      _ExtentX        =   27966
      _ExtentY        =   7858
      _Version        =   393216
      SelectionMode   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "lista_pendientesremitir.frx":0039
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Visible         =   0   'False
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   1720
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      BackColor       =   16777215
      HeadLines       =   1
      RowHeight       =   19
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
         Size            =   9.75
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
      Height          =   600
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   18945
      _ExtentX        =   33417
      _ExtentY        =   1058
      ButtonWidth     =   609
      ButtonHeight    =   1005
      Appearance      =   1
      Style           =   1
      _Version        =   393216
      Begin VB.CommandButton Command2 
         Caption         =   "O.C.:"
         Height          =   375
         Left            =   6600
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   120
         Width           =   855
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
         Left            =   7560
         TabIndex        =   16
         Top             =   120
         Width           =   1335
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Ambos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   13800
         TabIndex        =   14
         Top             =   120
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Preparados"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   12120
         TabIndex        =   13
         Top             =   120
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.OptionButton Option1 
         Caption         =   "No Preparadas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   10200
         TabIndex        =   12
         Top             =   120
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.CommandButton filtra 
         Caption         =   "filtra"
         Height          =   375
         Left            =   17160
         TabIndex        =   11
         Top             =   0
         Visible         =   0   'False
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
      Begin Crystal.CrystalReport CrystalReporte 
         Left            =   3000
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
         Left            =   15960
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
         MICON           =   "lista_pendientesremitir.frx":0056
         PICN            =   "lista_pendientesremitir.frx":0072
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
         Left            =   14280
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
         MICON           =   "lista_pendientesremitir.frx":0BBC
         PICN            =   "lista_pendientesremitir.frx":0BD8
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
      Top             =   8520
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
      Left            =   1200
      Top             =   8520
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
   Begin KewlButtonz.KewlButtons grabar 
      Height          =   855
      Left            =   16080
      TabIndex        =   9
      Top             =   1800
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   1508
      BTYPE           =   14
      TX              =   "&Remitir y Facturar"
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
      MICON           =   "lista_pendientesremitir.frx":3FCA
      PICN            =   "lista_pendientesremitir.frx":3FE6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSAdodcLib.Adodc datparametros 
      Height          =   450
      Left            =   2520
      Top             =   8520
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   794
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
   Begin MSAdodcLib.Adodc datencabezado 
      Height          =   330
      Left            =   3720
      Top             =   8520
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
   Begin MSAdodcLib.Adodc datencabezado2 
      Height          =   330
      Left            =   5280
      Top             =   8520
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
   Begin MSAdodcLib.Adodc datpreparados 
      Height          =   450
      Left            =   6600
      Top             =   8520
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   794
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
   Begin KewlButtonz.KewlButtons remitir 
      Height          =   855
      Left            =   16080
      TabIndex        =   15
      Top             =   720
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   1508
      BTYPE           =   14
      TX              =   "&Remitir"
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
      MICON           =   "lista_pendientesremitir.frx":5A68
      PICN            =   "lista_pendientesremitir.frx":5A84
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
Attribute VB_Name = "lista_pendientesremitir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public cuentacont As String
Public controlsalto As Integer
Public xclaveremito As Double
Dim cuenta(99999) As Integer



Private Sub Check1_Click()

Call filtra_Click

End Sub

Private Sub Check2_Click()

Call filtra_Click

End Sub

Private Sub Command1_Click()
On Error Resume Next
'' Actiualiza msflexgid
msgrid1.WordWrap = True

msgrid1.ColWidth(0) = 200
msgrid1.ColWidth(1) = 0
msgrid1.ColWidth(4) = 3500
msgrid1.ColAlignment(7) = 7
msgrid1.ColWidth(8) = 0
msgrid1.ColWidth(9) = 0
msgrid1.ColWidth(10) = 0
msgrid1.ColWidth(11) = 4500
msgrid1.ColWidth(12) = 0
msgrid1.ColWidth(13) = 0
msgrid1.ColWidth(14) = 0
msgrid1.ColWidth(15) = 1500
msgrid1.TextMatrix(0, 2) = "N.V. Nro"
    
For X = 1 To datpresupuesto.Recordset.RecordCount
    msgrid1.RowHeight(X) = 500
    msgrid1.TextMatrix(X, 3) = Format(msgrid1.TextMatrix(X, 3), "dd/mm/yyyy")
    msgrid1.TextMatrix(X, 7) = Format(msgrid1.TextMatrix(X, 7), "#,##0.00")


'    If Lm = 0 Then
'      For Y = 1 To 15
'         msgrid1.Col = Y
'         msgrid1.Row = X
'         msgrid1.CellBackColor = QBColor(11)
'      Next Y
'      Lm = 1
'    Else
'      Lm = 0
'    End If
Next X


End Sub

Private Sub Command4_Click()
On Error GoTo fuera
Dim crApp As New CRAXDRT.Application
Dim crReport As New CRAXDRT.Report
Dim tabla As String
Dim ruta As String

reporte.SQL = "SELECT v_ezi_pos_presupuesto.NUMERODOCUMENTO, v_ezi_pos_presupuesto.FECHAEMISION, v_ezi_pos_presupuesto.cod_cliente, v_ezi_pos_presupuesto.cliente, v_ezi_pos_presupuesto.CUIT, v_ezi_pos_presupuesto.CODPOS, v_ezi_pos_presupuesto.provincia, v_ezi_pos_presupuesto.vendedor, v_ezi_pos_presupuesto.detalle, v_ezi_pos_presupuesto.tipopago, v_ezi_pos_presupuesto.codigoproducto, v_ezi_pos_presupuesto.nombre_producto, v_ezi_pos_presupuesto.cantidadproducto, v_ezi_pos_presupuesto.nota, v_ezi_pos_presupuesto.condiva, v_ezi_pos_presupuesto.ciudad, v_ezi_pos_presupuesto.TIPOVENTA, v_ezi_pos_presupuesto.SIMBOLO, v_ezi_pos_presupuesto.CODVENDEDOR, v_ezi_pos_presupuesto.preciusiniva, v_ezi_pos_presupuesto.subtotalsiniva, v_ezi_pos_presupuesto.impbonifsiniva, v_ezi_pos_presupuesto.percepiibb, v_ezi_pos_presupuesto.perceptem, v_ezi_pos_presupuesto.totaltr, v_ezi_pos_presupuesto.importeiva21, v_ezi_pos_presupuesto.importeiva105 FROM MMOSSE.dbo.v_ezi_pos_presupuesto v_ezi_pos_presupuesto " & _
              " where v_ezi_pos_presupuesto.id = " & msgrid1.TextMatrix(msgrid1.Row, 1) & " order by v_ezi_pos_presupuesto.iditem"


tabla = reporte.SQL

With CrystalReporte
    .PrinterCollation = crptCollated
    If msgrid1.TextMatrix(msgrid1.Row, 10) = "A" Then
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

Private Sub Command8_Click()
On Error Resume Next
    menu = 2
'    xfila = KlexGrid1.Row
      query = "SELECT  * from  v_ezi_pos_stock_lotes " & _
            "where REFERENCIATIPO_ID = '" & KlexGrid1.TextMatrix(KlexGrid1.Row, 10) & "' " & _
            "ORDER BY FECHAVENCIMIENTO, CODIGO"
    lista_lotes.Show
'    lista_lotes.DataGrid1.SetFocus
    
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
    
'If menu = 1 Then
'            frmnota_venta.Text17.Text = DataGrid1.Columns(1).Text
'            frmnota_venta.Text18.Text = DataGrid1.Columns(7).Text
'            frmnota_venta.Text17.SetFocus
'            SendKeys "{ENTER}", False
'            Unload Me
'End If

'If menu = 2 Then
'            frmpresupuesto.Text17.Text = DataGrid1.Columns(1).Text
'            frmpresupuesto.Text18.Text = DataGrid1.Columns(7).Text
'            frmpresupuesto.Text17.SetFocus
'            SendKeys "{ENTER}", False
'            Unload Me
'End If
        

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

Private Sub facturar_Click()

End Sub

Private Sub filtra_Click()
On Error Resume Next

If Text1.Text = "" Then
        
      
If Option1.Value = True Then
  xquery1 = "SELECT     ud_ezi_puntodeventa_encabezado.claveprimaria AS Numero, ud_ezi_puntodeventa_encabezado.numerodefactura AS Nro, " & _
            "ud_ezi_puntodeventa_encabezado.fechadelcomprobante AS Fecha, ud_ezi_puntodeventa_encabezado.cliente AS Cliente, V_PERSONA_.CUIT, " & _
            "ud_ezi_puntodeventa_encabezado.vendedor AS Vendedor, ud_ezi_puntodeventa_encabezado.importeglobal AS Importe, ud_ezi_puntodeventa_encabezado.id, " & _
            "ud_ezi_puntodeventa_encabezado.generada, ud_ezi_puntodeventa_encabezado.tipodefactura, ud_ezi_puntodeventa_encabezado.nota, " & _
            "ud_ezi_puntodeventa_encabezado.numerodefactura + 'n ' + ud_ezi_puntodeventa_encabezado.cliente + ' ' + V_PERSONA_.CUIT + ' ' + ud_ezi_puntodeventa_encabezado.vendedor " & _
            "AS concatenado, prep.claveprimaria AS Preparadonro,ud_ezi_puntodeventa_encabezado.id, ud_ezi_puntodeventa_encabezado.adicionalid as OC FROM         ud_ezi_puntodeventa_encabezado AS prep WITH (readpast) RIGHT OUTER JOIN  " & _
            "ud_ezi_puntodeventa_encabezado WITH (readpast) ON prep.presupuestobase = ud_ezi_puntodeventa_encabezado.claveprimaria LEFT OUTER JOIN " & _
            "V_PERSONA_ RIGHT OUTER JOIN V_CLIENTE_ ON V_PERSONA_.ID = V_CLIENTE_.ENTEASOCIADO_ID ON ud_ezi_puntodeventa_encabezado.clienteid = V_CLIENTE_.ID " & _
            "WHERE     (ud_ezi_puntodeventa_encabezado.numeradorinterno = 'Nota de Venta') AND (prep.estado IS NULL) ORDER BY Fecha DESC"
End If
If Option3.Value = True Then
  xquery1 = "SELECT     ud_ezi_puntodeventa_encabezado.claveprimaria AS Numero, ud_ezi_puntodeventa_encabezado.numerodefactura AS Nro, " & _
            "ud_ezi_puntodeventa_encabezado.fechadelcomprobante AS Fecha, ud_ezi_puntodeventa_encabezado.cliente AS Cliente, V_PERSONA_.CUIT, " & _
            "ud_ezi_puntodeventa_encabezado.vendedor AS Vendedor, ud_ezi_puntodeventa_encabezado.importeglobal AS Importe, ud_ezi_puntodeventa_encabezado.id, " & _
            "ud_ezi_puntodeventa_encabezado.generada, ud_ezi_puntodeventa_encabezado.tipodefactura, ud_ezi_puntodeventa_encabezado.nota, " & _
            "ud_ezi_puntodeventa_encabezado.numerodefactura + 'n ' + ud_ezi_puntodeventa_encabezado.cliente + ' ' + V_PERSONA_.CUIT + ' ' + ud_ezi_puntodeventa_encabezado.vendedor " & _
            "AS concatenado, prep.claveprimaria AS Preparadonro,ud_ezi_puntodeventa_encabezado.id, ud_ezi_puntodeventa_encabezado.adicionalid as OC FROM         ud_ezi_puntodeventa_encabezado AS prep WITH (readpast) RIGHT OUTER JOIN  " & _
            "ud_ezi_puntodeventa_encabezado WITH (readpast) ON prep.presupuestobase = ud_ezi_puntodeventa_encabezado.claveprimaria LEFT OUTER JOIN " & _
            "V_PERSONA_ RIGHT OUTER JOIN V_CLIENTE_ ON V_PERSONA_.ID = V_CLIENTE_.ENTEASOCIADO_ID ON ud_ezi_puntodeventa_encabezado.clienteid = V_CLIENTE_.ID " & _
            "WHERE     (ud_ezi_puntodeventa_encabezado.numeradorinterno = 'Nota de Venta') AND (prep.estado IS NULL OR prep.estado = 'Preparado') ORDER BY Fecha DESC"
End If

If Option2.Value = True Then
  xquery1 = "SELECT     ud_ezi_puntodeventa_encabezado.claveprimaria AS Numero, ud_ezi_puntodeventa_encabezado.numerodefactura AS Nro, " & _
            "ud_ezi_puntodeventa_encabezado.fechadelcomprobante AS Fecha, ud_ezi_puntodeventa_encabezado.cliente AS Cliente, V_PERSONA_.CUIT, " & _
            "ud_ezi_puntodeventa_encabezado.vendedor AS Vendedor, ud_ezi_puntodeventa_encabezado.importeglobal AS Importe, ud_ezi_puntodeventa_encabezado.id, " & _
            "ud_ezi_puntodeventa_encabezado.generada, ud_ezi_puntodeventa_encabezado.tipodefactura, ud_ezi_puntodeventa_encabezado.nota, " & _
            "ud_ezi_puntodeventa_encabezado.numerodefactura + 'n ' + ud_ezi_puntodeventa_encabezado.cliente + ' ' + V_PERSONA_.CUIT + ' ' + ud_ezi_puntodeventa_encabezado.vendedor " & _
            "AS concatenado, prep.claveprimaria AS Preparadonro,ud_ezi_puntodeventa_encabezado.id,  ud_ezi_puntodeventa_encabezado.adicionalid as OC FROM         ud_ezi_puntodeventa_encabezado AS prep WITH (readpast) RIGHT OUTER JOIN  " & _
            "ud_ezi_puntodeventa_encabezado WITH (readpast) ON prep.presupuestobase = ud_ezi_puntodeventa_encabezado.claveprimaria LEFT OUTER JOIN " & _
            "V_PERSONA_ RIGHT OUTER JOIN V_CLIENTE_ ON V_PERSONA_.ID = V_CLIENTE_.ENTEASOCIADO_ID ON ud_ezi_puntodeventa_encabezado.clienteid = V_CLIENTE_.ID " & _
            "WHERE     (ud_ezi_puntodeventa_encabezado.numeradorinterno = 'Nota de Venta') AND (prep.estado = 'Preparado') ORDER BY Fecha DESC"
End If
Else
        xbusqueda = "%" + Text1.Text + "%"
        If Option1.Value = True Then
            xquery1 = "SELECT     ud_ezi_puntodeventa_encabezado.claveprimaria AS Numero, ud_ezi_puntodeventa_encabezado.numerodefactura AS Nro, " & _
                      "ud_ezi_puntodeventa_encabezado.fechadelcomprobante AS Fecha, ud_ezi_puntodeventa_encabezado.cliente AS Cliente, V_PERSONA_.CUIT, " & _
                      "ud_ezi_puntodeventa_encabezado.vendedor AS Vendedor, ud_ezi_puntodeventa_encabezado.importeglobal AS Importe, ud_ezi_puntodeventa_encabezado.id, " & _
                      "ud_ezi_puntodeventa_encabezado.generada, ud_ezi_puntodeventa_encabezado.tipodefactura, ud_ezi_puntodeventa_encabezado.nota, " & _
                      "ud_ezi_puntodeventa_encabezado.numerodefactura + 'n ' + ud_ezi_puntodeventa_encabezado.cliente + ' ' + V_PERSONA_.CUIT + ' ' + ud_ezi_puntodeventa_encabezado.vendedor " & _
                      "AS concatenado, prep.claveprimaria AS Preparadonro,ud_ezi_puntodeventa_encabezado.id, ud_ezi_puntodeventa_encabezado.adicionalid as OC FROM         ud_ezi_puntodeventa_encabezado AS prep WITH (readpast) RIGHT OUTER JOIN  " & _
                      "ud_ezi_puntodeventa_encabezado WITH (readpast) ON prep.presupuestobase = ud_ezi_puntodeventa_encabezado.claveprimaria LEFT OUTER JOIN " & _
                      "V_PERSONA_ RIGHT OUTER JOIN V_CLIENTE_ ON V_PERSONA_.ID = V_CLIENTE_.ENTEASOCIADO_ID ON ud_ezi_puntodeventa_encabezado.clienteid = V_CLIENTE_.ID " & _
                      "WHERE     (ud_ezi_puntodeventa_encabezado.numeradorinterno = 'Nota de Venta') AND (prep.estado IS NULL) AND " & _
                      "(ud_ezi_puntodeventa_encabezado.numerodefactura + 'n ' + ud_ezi_puntodeventa_encabezado.cliente + ' ' + V_PERSONA_.CUIT + ' ' + ud_ezi_puntodeventa_encabezado.vendedor like '" & xbusqueda & "') ORDER BY Fecha DESC"
        End If
        If Option3.Value = True Then
            xquery1 = "SELECT     ud_ezi_puntodeventa_encabezado.claveprimaria AS Numero, ud_ezi_puntodeventa_encabezado.numerodefactura AS Nro, " & _
                      "ud_ezi_puntodeventa_encabezado.fechadelcomprobante AS Fecha, ud_ezi_puntodeventa_encabezado.cliente AS Cliente, V_PERSONA_.CUIT, " & _
                      "ud_ezi_puntodeventa_encabezado.vendedor AS Vendedor, ud_ezi_puntodeventa_encabezado.importeglobal AS Importe, ud_ezi_puntodeventa_encabezado.id, " & _
                      "ud_ezi_puntodeventa_encabezado.generada, ud_ezi_puntodeventa_encabezado.tipodefactura, ud_ezi_puntodeventa_encabezado.nota, " & _
                      "ud_ezi_puntodeventa_encabezado.numerodefactura + 'n ' + ud_ezi_puntodeventa_encabezado.cliente + ' ' + V_PERSONA_.CUIT + ' ' + ud_ezi_puntodeventa_encabezado.vendedor " & _
                      "AS concatenado, prep.claveprimaria AS Preparadonro,ud_ezi_puntodeventa_encabezado.id,  ud_ezi_puntodeventa_encabezado.adicionalid as OC FROM         ud_ezi_puntodeventa_encabezado AS prep WITH (readpast) RIGHT OUTER JOIN  " & _
                      "ud_ezi_puntodeventa_encabezado WITH (readpast) ON prep.presupuestobase = ud_ezi_puntodeventa_encabezado.claveprimaria LEFT OUTER JOIN " & _
                      "V_PERSONA_ RIGHT OUTER JOIN V_CLIENTE_ ON V_PERSONA_.ID = V_CLIENTE_.ENTEASOCIADO_ID ON ud_ezi_puntodeventa_encabezado.clienteid = V_CLIENTE_.ID " & _
                      "WHERE     (ud_ezi_puntodeventa_encabezado.numeradorinterno = 'Nota de Venta') AND (prep.estado IS NULL OR prep.estado = 'Preparado') AND " & _
                      "(ud_ezi_puntodeventa_encabezado.numerodefactura + 'n ' + ud_ezi_puntodeventa_encabezado.cliente + ' ' + V_PERSONA_.CUIT + ' ' + ud_ezi_puntodeventa_encabezado.vendedor like '" & xbusqueda & "') ORDER BY Fecha DESC"
        End If

        If Option2.Value = True Then
            xquery1 = "SELECT     ud_ezi_puntodeventa_encabezado.claveprimaria AS Numero, ud_ezi_puntodeventa_encabezado.numerodefactura AS Nro, " & _
                      "ud_ezi_puntodeventa_encabezado.fechadelcomprobante AS Fecha, ud_ezi_puntodeventa_encabezado.cliente AS Cliente, V_PERSONA_.CUIT, " & _
                      "ud_ezi_puntodeventa_encabezado.vendedor AS Vendedor, ud_ezi_puntodeventa_encabezado.importeglobal AS Importe, ud_ezi_puntodeventa_encabezado.id, " & _
                      "ud_ezi_puntodeventa_encabezado.generada, ud_ezi_puntodeventa_encabezado.tipodefactura, ud_ezi_puntodeventa_encabezado.nota, " & _
                      "ud_ezi_puntodeventa_encabezado.numerodefactura + 'n ' + ud_ezi_puntodeventa_encabezado.cliente + ' ' + V_PERSONA_.CUIT + ' ' + ud_ezi_puntodeventa_encabezado.vendedor " & _
                      "AS concatenado, prep.claveprimaria AS Preparadonro,ud_ezi_puntodeventa_encabezado.id,  ud_ezi_puntodeventa_encabezado.adicionalid as OC FROM         ud_ezi_puntodeventa_encabezado AS prep WITH (readpast) RIGHT OUTER JOIN  " & _
                      "ud_ezi_puntodeventa_encabezado WITH (readpast) ON prep.presupuestobase = ud_ezi_puntodeventa_encabezado.claveprimaria LEFT OUTER JOIN " & _
                      "V_PERSONA_ RIGHT OUTER JOIN V_CLIENTE_ ON V_PERSONA_.ID = V_CLIENTE_.ENTEASOCIADO_ID ON ud_ezi_puntodeventa_encabezado.clienteid = V_CLIENTE_.ID " & _
                      "WHERE     (ud_ezi_puntodeventa_encabezado.numeradorinterno = 'Nota de Venta') AND (prep.estado = 'Preparado') AND " & _
                      "(ud_ezi_puntodeventa_encabezado.numerodefactura + 'n ' + ud_ezi_puntodeventa_encabezado.cliente + ' ' + V_PERSONA_.CUIT + ' ' + ud_ezi_puntodeventa_encabezado.vendedor like '" & xbusqueda & "') ORDER BY Fecha DESC"
        End If
End If





datpresupuesto.RecordSource = xquery1
datpresupuesto.Refresh
If datpresupuesto.Recordset.EOF = False Then datpresupuesto.Recordset.MoveFirst

Call Command1_Click
If msgrid1.Rows >= 1 Then
    msgrid1.Row = 1
    msgrid1.Col = 1
End If

Call msgrid1_Click

End Sub

Private Sub Form_Activate()

'DataGrid1.SetFocus
msgrid1.SetFocus

End Sub

Private Sub Form_Load()
'If menu = 2 Then
'       Aplicar_skin2 Me
'Else
    Aplicar_skin Me
'End If

MiFuncionDeAjuste Me, True

yventana = Inicio.Height / 2 - 1000
xventana = Inicio.Width / 2

lista_pendientesremitir.Top = yventana - lista_pendientesremitir.Height / 2
lista_pendientesremitir.Left = xventana - lista_pendientesremitir.Width / 2


datpresupuesto.ConnectionString = login.conexiontotal
datitems.ConnectionString = login.conexiontotal
datparametros.ConnectionString = login.conexiontotal
datencabezado.ConnectionString = login.conexiontotal
datencabezado2.ConnectionString = login.conexiontotal
datpreparados.ConnectionString = login.conexiontotal

Option2.Value = True


    datparametros.RecordSource = "select * from ud_ezi_parametros_pos where sucursal = '" & login.nomsucursal & "' "
    datparametros.Refresh


Call filtra_Click
 
End Sub

Private Sub grabar_Click()
query = "Select * from ud_ezi_puntodeventa_encabezado with (nolock) where claveprimaria = " & xclaveremito & ""


    remdev = 0
    frmnota_venta_remito.Text17 = xclaveremito
    frmnota_venta_remito.Show
    frmnota_venta_remito.Text17.SetFocus
    SendKeys "{ENTER}", False
    
    
    
    Unload Me

End Sub

Private Sub KlexGrid1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
On Error Resume Next
 
 If Val(KlexGrid1.TextMatrix(KlexGrid1.Row, KlexGrid1.Col)) > Val(KlexGrid1.TextMatrix(KlexGrid1.Row, KlexGrid1.Col - 2)) Then
  '  MsgBox "El valor no puede ser mayor a la cantidad establecida", vbCritical, "Error !!"
    KlexGrid1.TextMatrix(KlexGrid1.Row, KlexGrid1.Col) = 0
    controlsalto = 1
 Else
    controlsalto = 0
'    Call Command8_Click
 End If

End Sub



Private Sub KlexGrid1_BeforeEdit(Cancel As Boolean)

    xfila = KlexGrid1.Row

End Sub

Private Sub KlexGrid1_EnterCell()

On Error Resume Next



If (KlexGrid1.Col >= 0 And KlexGrid1.Col <= 6) Or (KlexGrid1.Col >= 8) Then
    KlexGrid1.Editable = False
    Exit Sub
Else
    If controlsalto = 1 And KlexGrid1.Row > 1 Then KlexGrid1.Row = KlexGrid1.Row - 1
    
    If Val(KlexGrid1.TextMatrix(KlexGrid1.Row, KlexGrid1.Col)) > Val(KlexGrid1.TextMatrix(KlexGrid1.Row, KlexGrid1.Col - 2)) Then
         KlexGrid1.TextMatrix(KlexGrid1.Row, KlexGrid1.Col) = 0
    End If
    
    
    KlexGrid1.Editable = False
    
    msgridlote.Clear
    msgridlote.TextMatrix(0, 0) = "Lote"
    msgridlote.TextMatrix(0, 1) = "Cantidad"
    msgridlote.ColWidth(2) = 0
    xlin = 1
    For X = 1 To 20
            If lotecantidad(KlexGrid1.Row, X) <> 0 Then
                msgridlote.TextMatrix(xlin, 0) = lotecodigo(KlexGrid1.Row, X)
                msgridlote.TextMatrix(xlin, 1) = lotecantidad(KlexGrid1.Row, X)
                msgridlote.TextMatrix(xlin, 2) = loteid(KlexGrid1.Row, X)
                xlin = xlin + 1
            End If
    Next X

    
End If


End Sub

Private Sub msgrid1_Click()
On Error Resume Next
    xidencabezado = msgrid1.TextMatrix(msgrid1.Row, 14)
    If xidencabezado = "id" Then
        KlexGrid1.Visible = False
        Exit Sub
    Else
        KlexGrid1.Visible = True
    End If
    
    
      datitems.RecordSource = "SELECT     ud_ezi_puntodeventa_detalle_notav.id, ud_ezi_puntodeventa_detalle_notav.codigoproducto AS Codigo, " & _
                            "ISNULL(ud_ezi_puntodeventa_detalle_notav.nombre_producto, '') AS Descripcion, " & _
                            "ISNULL(M.CODIGO, '') AS Marca, ISNULL(EP.PRESENTACION, '') AS Presentacion, ISNULL(ud_ezi_puntodeventa_detalle_notav.cantidadproducto, 0) AS Cantidad, " & _
                            "ISNULL(ud_ezi_puntodeventa_detalle_notav.unidaddemedidaid, '') AS Um, ISNULL(ud_ezi_puntodeventa_detalle_notav.preparado, 0) AS Preparado, " & _
                            "ISNULL(xrem.cantidadremitida, 0) AS Remitido, ISNULL(ud_ezi_puntodeventa_detalle_notav.lote, 0) AS Lote, P.ID AS prod_id, " & _
                            "ISNULL(ud_ezi_puntodeventa_detalle_notav.lote_id, '') AS lote_id, ud_ezi_puntodeventa_detalle_notav.claveprimaria, NV.numerodefactura, xrem.cantidadoriginal, xrem.cantidadremitida,ud_ezi_puntodeventa_detalle_notav.item as itemnv " & _
                            "FROM         ud_ezi_puntodeventa_encabezado AS NV WITH (readpast) INNER JOIN " & _
                            "ud_ezi_puntodeventa_detalle_notav WITH (nolock) ON NV.id = ud_ezi_puntodeventa_detalle_notav.claveprimaria LEFT OUTER JOIN " & _
                            "(SELECT     idproducto, facturaorigen, cantidadoriginal, SUM(cantidadremitida) AS cantidadremitida, cantidadoriginal - SUM(cantidadremitida) AS dif, item " & _
                            "FROM          ud_ezi_puntodeventa_detalle_rem WITH (readpast) GROUP BY idproducto, facturaorigen, cantidadoriginal, item HAVING      (SUM(cantidadremitida) <> 0)) AS xrem ON ud_ezi_puntodeventa_detalle_notav.item = xrem.item AND " & _
                            "ud_ezi_puntodeventa_detalle_notav.idproducto = xrem.idproducto AND NV.numerodefactura = xrem.facturaorigen LEFT OUTER JOIN " & _
                            "V_ITEMTIPOCLASIFICADOR AS M WITH (readpast) RIGHT OUTER JOIN " & _
                            "V_UD_EZI_PRODUCTOS AS EP WITH (readpast) ON M.ID = EP.MARCA_ID RIGHT OUTER JOIN " & _
                            "V_PRODUCTO AS P WITH (readpast) ON EP.ID = P.BOEXTENSION_ID ON ud_ezi_puntodeventa_detalle_notav.idproducto = P.ID " & _
                            "where ud_ezi_puntodeventa_detalle_notav.claveprimaria = " & xidencabezado & " and ISNULL(ud_ezi_puntodeventa_detalle_notav.cantidadproducto, 0) > ISNULL(xrem.cantidadremitida, 0)  ORDER BY ud_ezi_puntodeventa_detalle_notav.item"

    datitems.Refresh

    If datitems.Recordset.EOF = False Then
        datencabezado2.RecordSource = "Select * from ud_ezi_puntodeventa_encabezado with (nolock) where id = '" & datitems.Recordset.Fields("claveprimaria") & "'"
        datencabezado2.Refresh
    End If

    KlexGrid1.Rows = datitems.Recordset.RecordCount + 1
    KlexGrid1.Cols = 12
    KlexGrid1.ColWidth(0) = 50
    KlexGrid1.ColWidth(1) = 1500
    KlexGrid1.ColWidth(2) = 5000
    KlexGrid1.ColWidth(3) = 2000
    KlexGrid1.ColWidth(4) = 2000
    KlexGrid1.TextMatrix(0, 1) = "Código"
    KlexGrid1.TextMatrix(0, 2) = "Descripción"
    KlexGrid1.TextMatrix(0, 3) = "Marca"
    KlexGrid1.TextMatrix(0, 4) = "Presentación"
    KlexGrid1.TextMatrix(0, 5) = "Cantidad"
    KlexGrid1.TextMatrix(0, 6) = "U.M."
    KlexGrid1.TextMatrix(0, 7) = "Preparar"
    KlexGrid1.TextMatrix(0, 8) = "Remitido"
    KlexGrid1.ColWidth(9) = 0
    KlexGrid1.ColWidth(10) = 0
    KlexGrid1.ColWidth(11) = 0

    
    For X = 1 To 50
        For Y = 1 To 20
         lotecodigo(X, Y) = ""
         lotecantidad(X, Y) = 0
         loteid(X, Y) = ""
        Next Y
    Next X
    
    
    lin = 1
If datitems.Recordset.EOF = False Then
    datitems.Recordset.MoveFirst
    Do While Not datitems.Recordset.EOF
        For X = 0 To KlexGrid1.Cols - 1
            If X >= 1 And X <= 4 Then
                KlexGrid1.TextMatrix(lin, X) = CStr(datitems.Recordset.Fields(X)) + "."
            Else
                KlexGrid1.TextMatrix(lin, X) = datitems.Recordset.Fields(X)
            End If
        Next X
'            datpreparados.RecordSource = "select * from ud_ezi_puntodeventa_detalle_rem with (nolock) where facturaorigen = '" & msgrid1.TextMatrix(msgrid1.Row, 1) & "' and item = " & lin & " and cantidadremitida = 0   "
'            datpreparados.Refresh
            datpreparados.RecordSource = "select * from ud_ezi_puntodeventa_detalle_rem with (nolock) where facturaorigen = '" & msgrid1.TextMatrix(msgrid1.Row, 1) & "' and item = " & datitems.Recordset.Fields("itemnv") & " and cantidadremitida = 0   "
            datpreparados.Refresh

            plin = 1
            If datpreparados.Recordset.EOF = False Then
                xclaveremito = datpreparados.Recordset.Fields("claveprimaria")
                xprep = 0
                xlin = 1
                Do While Not datpreparados.Recordset.EOF
                    xprep = xprep + Val(datpreparados.Recordset.Fields("cantidadaremitir"))
                    
                    msgridlote.TextMatrix(xlin, 0) = datpreparados.Recordset.Fields("lote")
                    msgridlote.TextMatrix(xlin, 1) = datpreparados.Recordset.Fields("cantidadaremitir")
                    msgridlote.TextMatrix(xlin, 2) = datpreparados.Recordset.Fields("lote_id")
                    
                    lotecodigo(lin, xlin) = datpreparados.Recordset.Fields("lote")
                    lotecantidad(lin, xlin) = datpreparados.Recordset.Fields("cantidadaremitir")
                    loteid(lin, xlin) = datpreparados.Recordset.Fields("lote_id")
                    
                    xlin = xlin + 1

                    datpreparados.Recordset.MoveNext
                Loop
               KlexGrid1.TextMatrix(lin, 7) = xprep
            End If

        lin = lin + 1
        datitems.Recordset.MoveNext
        
    Loop
End If
    
If 1 = 2 Then
   xrow = msgrid1.Row
   For X = 1 To msgrid1.Rows - 1
    If Lm = 0 Then
      For Y = 1 To 15
         msgrid1.Col = Y
         msgrid1.Row = X
         msgrid1.CellBackColor = QBColor(11)
      Next Y
      Lm = 1
    Else
      Lm = 0
      For Y = 1 To 15
         msgrid1.Col = Y
         msgrid1.Row = X
         msgrid1.CellBackColor = QBColor(15)
      Next Y
    End If
   Next X
    
    For Y = 1 To 15
         msgrid1.Col = Y
         msgrid1.Row = xrow
         msgrid1.CellBackColor = QBColor(10)
    Next Y
End If

    msgridlote.Clear

End Sub

Private Sub msgrid1_KeyDown(KeyCode As Integer, Shift As Integer)

    Call msgrid1_Click


End Sub

Private Sub Option1_Click()

    Call filtra_Click

End Sub

Private Sub Option2_Click()

Call filtra_Click

End Sub

Private Sub Option3_Click()

Call filtra_Click

End Sub

Private Sub remitir_Click()
On Error Resume Next


query = "Select * from ud_ezi_puntodeventa_encabezado with (nolock) where claveprimaria = " & xclaveremito & ""


    remdev = 0
    frmnota_venta_remito.Text17 = xclaveremito
    frmnota_venta_remito.Show
    frmnota_venta_remito.Text17.SetFocus
    SendKeys "{ENTER}", False
    Unload Me
    

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
           
                    
        If Option1.Value = True Then
          If Text2.Text = "" Then
            xquery1 = "SELECT     ud_ezi_puntodeventa_encabezado.claveprimaria AS Numero, ud_ezi_puntodeventa_encabezado.numerodefactura AS Nro, " & _
                      "ud_ezi_puntodeventa_encabezado.fechadelcomprobante AS Fecha, ud_ezi_puntodeventa_encabezado.cliente AS Cliente, V_PERSONA_.CUIT, " & _
                      "ud_ezi_puntodeventa_encabezado.vendedor AS Vendedor, ud_ezi_puntodeventa_encabezado.importeglobal AS Importe, ud_ezi_puntodeventa_encabezado.id, " & _
                      "ud_ezi_puntodeventa_encabezado.generada, ud_ezi_puntodeventa_encabezado.tipodefactura, ud_ezi_puntodeventa_encabezado.nota, " & _
                      "ud_ezi_puntodeventa_encabezado.numerodefactura + 'n ' + ud_ezi_puntodeventa_encabezado.cliente + ' ' + V_PERSONA_.CUIT + ' ' + ud_ezi_puntodeventa_encabezado.vendedor " & _
                      "AS concatenado, prep.claveprimaria AS Preparadonro,ud_ezi_puntodeventa_encabezado.id,  ud_ezi_puntodeventa_encabezado.adicionalid as OC FROM         ud_ezi_puntodeventa_encabezado AS prep WITH (readpast) RIGHT OUTER JOIN  " & _
                      "ud_ezi_puntodeventa_encabezado WITH (readpast) ON prep.presupuestobase = ud_ezi_puntodeventa_encabezado.claveprimaria LEFT OUTER JOIN " & _
                      "V_PERSONA_ RIGHT OUTER JOIN V_CLIENTE_ ON V_PERSONA_.ID = V_CLIENTE_.ENTEASOCIADO_ID ON ud_ezi_puntodeventa_encabezado.clienteid = V_CLIENTE_.ID " & _
                      "WHERE     (ud_ezi_puntodeventa_encabezado.numeradorinterno = 'Nota de Venta') AND (prep.estado IS NULL) AND " & _
                      "(ud_ezi_puntodeventa_encabezado.numerodefactura + 'n ' + ud_ezi_puntodeventa_encabezado.cliente + ' ' + V_PERSONA_.CUIT + ' ' + ud_ezi_puntodeventa_encabezado.vendedor like '" & xbusqueda & "') ORDER BY Fecha DESC"
          Else
            xquery1 = "SELECT     ud_ezi_puntodeventa_encabezado.claveprimaria AS Numero, ud_ezi_puntodeventa_encabezado.numerodefactura AS Nro, " & _
                      "ud_ezi_puntodeventa_encabezado.fechadelcomprobante AS Fecha, ud_ezi_puntodeventa_encabezado.cliente AS Cliente, V_PERSONA_.CUIT, " & _
                      "ud_ezi_puntodeventa_encabezado.vendedor AS Vendedor, ud_ezi_puntodeventa_encabezado.importeglobal AS Importe, ud_ezi_puntodeventa_encabezado.id, " & _
                      "ud_ezi_puntodeventa_encabezado.generada, ud_ezi_puntodeventa_encabezado.tipodefactura, ud_ezi_puntodeventa_encabezado.nota, " & _
                      "ud_ezi_puntodeventa_encabezado.numerodefactura + 'n ' + ud_ezi_puntodeventa_encabezado.cliente + ' ' + V_PERSONA_.CUIT + ' ' + ud_ezi_puntodeventa_encabezado.vendedor " & _
                      "AS concatenado, prep.claveprimaria AS Preparadonro,ud_ezi_puntodeventa_encabezado.id,  ud_ezi_puntodeventa_encabezado.adicionalid as OC FROM         ud_ezi_puntodeventa_encabezado AS prep WITH (readpast) RIGHT OUTER JOIN  " & _
                      "ud_ezi_puntodeventa_encabezado WITH (readpast) ON prep.presupuestobase = ud_ezi_puntodeventa_encabezado.claveprimaria LEFT OUTER JOIN " & _
                      "V_PERSONA_ RIGHT OUTER JOIN V_CLIENTE_ ON V_PERSONA_.ID = V_CLIENTE_.ENTEASOCIADO_ID ON ud_ezi_puntodeventa_encabezado.clienteid = V_CLIENTE_.ID " & _
                      "WHERE     (ud_ezi_puntodeventa_encabezado.numeradorinterno = 'Nota de Venta') AND (prep.estado IS NULL) AND " & _
                      "(ud_ezi_puntodeventa_encabezado.numerodefactura + 'n ' + ud_ezi_puntodeventa_encabezado.cliente + ' ' + V_PERSONA_.CUIT + ' ' + ud_ezi_puntodeventa_encabezado.vendedor like '" & xbusqueda & "') and ud_ezi_puntodeventa_encabezado.adicionalid = '" & Text2.Text & "' ORDER BY Fecha DESC"
           End If
        End If
        If Option3.Value = True Then
          If Text2.Text = "" Then
            xquery1 = "SELECT     ud_ezi_puntodeventa_encabezado.claveprimaria AS Numero, ud_ezi_puntodeventa_encabezado.numerodefactura AS Nro, " & _
                      "ud_ezi_puntodeventa_encabezado.fechadelcomprobante AS Fecha, ud_ezi_puntodeventa_encabezado.cliente AS Cliente, V_PERSONA_.CUIT, " & _
                      "ud_ezi_puntodeventa_encabezado.vendedor AS Vendedor, ud_ezi_puntodeventa_encabezado.importeglobal AS Importe, ud_ezi_puntodeventa_encabezado.id, " & _
                      "ud_ezi_puntodeventa_encabezado.generada, ud_ezi_puntodeventa_encabezado.tipodefactura, ud_ezi_puntodeventa_encabezado.nota, " & _
                      "ud_ezi_puntodeventa_encabezado.numerodefactura + 'n ' + ud_ezi_puntodeventa_encabezado.cliente + ' ' + V_PERSONA_.CUIT + ' ' + ud_ezi_puntodeventa_encabezado.vendedor " & _
                      "AS concatenado, prep.claveprimaria AS Preparadonro,ud_ezi_puntodeventa_encabezado.id,  ud_ezi_puntodeventa_encabezado.adicionalid as OC FROM         ud_ezi_puntodeventa_encabezado AS prep WITH (readpast) RIGHT OUTER JOIN  " & _
                      "ud_ezi_puntodeventa_encabezado WITH (readpast) ON prep.presupuestobase = ud_ezi_puntodeventa_encabezado.claveprimaria LEFT OUTER JOIN " & _
                      "V_PERSONA_ RIGHT OUTER JOIN V_CLIENTE_ ON V_PERSONA_.ID = V_CLIENTE_.ENTEASOCIADO_ID ON ud_ezi_puntodeventa_encabezado.clienteid = V_CLIENTE_.ID " & _
                      "WHERE     (ud_ezi_puntodeventa_encabezado.numeradorinterno = 'Nota de Venta') AND (prep.estado IS NULL OR prep.estado = 'Preparado') AND " & _
                      "(ud_ezi_puntodeventa_encabezado.numerodefactura + 'n ' + ud_ezi_puntodeventa_encabezado.cliente + ' ' + V_PERSONA_.CUIT + ' ' + ud_ezi_puntodeventa_encabezado.vendedor like '" & xbusqueda & "') ORDER BY Fecha DESC"
          Else
            xquery1 = "SELECT     ud_ezi_puntodeventa_encabezado.claveprimaria AS Numero, ud_ezi_puntodeventa_encabezado.numerodefactura AS Nro, " & _
                      "ud_ezi_puntodeventa_encabezado.fechadelcomprobante AS Fecha, ud_ezi_puntodeventa_encabezado.cliente AS Cliente, V_PERSONA_.CUIT, " & _
                      "ud_ezi_puntodeventa_encabezado.vendedor AS Vendedor, ud_ezi_puntodeventa_encabezado.importeglobal AS Importe, ud_ezi_puntodeventa_encabezado.id, " & _
                      "ud_ezi_puntodeventa_encabezado.generada, ud_ezi_puntodeventa_encabezado.tipodefactura, ud_ezi_puntodeventa_encabezado.nota, " & _
                      "ud_ezi_puntodeventa_encabezado.numerodefactura + 'n ' + ud_ezi_puntodeventa_encabezado.cliente + ' ' + V_PERSONA_.CUIT + ' ' + ud_ezi_puntodeventa_encabezado.vendedor " & _
                      "AS concatenado, prep.claveprimaria AS Preparadonro,ud_ezi_puntodeventa_encabezado.id,  ud_ezi_puntodeventa_encabezado.adicionalid as OC FROM         ud_ezi_puntodeventa_encabezado AS prep WITH (readpast) RIGHT OUTER JOIN  " & _
                      "ud_ezi_puntodeventa_encabezado WITH (readpast) ON prep.presupuestobase = ud_ezi_puntodeventa_encabezado.claveprimaria LEFT OUTER JOIN " & _
                      "V_PERSONA_ RIGHT OUTER JOIN V_CLIENTE_ ON V_PERSONA_.ID = V_CLIENTE_.ENTEASOCIADO_ID ON ud_ezi_puntodeventa_encabezado.clienteid = V_CLIENTE_.ID " & _
                      "WHERE     (ud_ezi_puntodeventa_encabezado.numeradorinterno = 'Nota de Venta') AND (prep.estado IS NULL OR prep.estado = 'Preparado') AND " & _
                      "(ud_ezi_puntodeventa_encabezado.numerodefactura + 'n ' + ud_ezi_puntodeventa_encabezado.cliente + ' ' + V_PERSONA_.CUIT + ' ' + ud_ezi_puntodeventa_encabezado.vendedor like '" & xbusqueda & "') and ud_ezi_puntodeventa_encabezado.adicionalid = '" & Text2.Text & "' ORDER BY Fecha DESC"
           End If
        End If

        If Option2.Value = True Then
          If Text2.Text = "" Then
            xquery1 = "SELECT     ud_ezi_puntodeventa_encabezado.claveprimaria AS Numero, ud_ezi_puntodeventa_encabezado.numerodefactura AS Nro, " & _
                      "ud_ezi_puntodeventa_encabezado.fechadelcomprobante AS Fecha, ud_ezi_puntodeventa_encabezado.cliente AS Cliente, V_PERSONA_.CUIT, " & _
                      "ud_ezi_puntodeventa_encabezado.vendedor AS Vendedor, ud_ezi_puntodeventa_encabezado.importeglobal AS Importe, ud_ezi_puntodeventa_encabezado.id, " & _
                      "ud_ezi_puntodeventa_encabezado.generada, ud_ezi_puntodeventa_encabezado.tipodefactura, ud_ezi_puntodeventa_encabezado.nota, " & _
                      "ud_ezi_puntodeventa_encabezado.numerodefactura + 'n ' + ud_ezi_puntodeventa_encabezado.cliente + ' ' + V_PERSONA_.CUIT + ' ' + ud_ezi_puntodeventa_encabezado.vendedor " & _
                      "AS concatenado, prep.claveprimaria AS Preparadonro,ud_ezi_puntodeventa_encabezado.id,  ud_ezi_puntodeventa_encabezado.adicionalid as OC FROM         ud_ezi_puntodeventa_encabezado AS prep WITH (readpast) RIGHT OUTER JOIN  " & _
                      "ud_ezi_puntodeventa_encabezado WITH (readpast) ON prep.presupuestobase = ud_ezi_puntodeventa_encabezado.claveprimaria LEFT OUTER JOIN " & _
                      "V_PERSONA_ RIGHT OUTER JOIN V_CLIENTE_ ON V_PERSONA_.ID = V_CLIENTE_.ENTEASOCIADO_ID ON ud_ezi_puntodeventa_encabezado.clienteid = V_CLIENTE_.ID " & _
                      "WHERE     (ud_ezi_puntodeventa_encabezado.numeradorinterno = 'Nota de Venta') AND (prep.estado = 'Preparado') AND " & _
                      "(ud_ezi_puntodeventa_encabezado.numerodefactura + 'n ' + ud_ezi_puntodeventa_encabezado.cliente + ' ' + V_PERSONA_.CUIT + ' ' + ud_ezi_puntodeventa_encabezado.vendedor like '" & xbusqueda & "') ORDER BY Fecha DESC"
           Else
             xquery1 = "SELECT     ud_ezi_puntodeventa_encabezado.claveprimaria AS Numero, ud_ezi_puntodeventa_encabezado.numerodefactura AS Nro, " & _
                      "ud_ezi_puntodeventa_encabezado.fechadelcomprobante AS Fecha, ud_ezi_puntodeventa_encabezado.cliente AS Cliente, V_PERSONA_.CUIT, " & _
                      "ud_ezi_puntodeventa_encabezado.vendedor AS Vendedor, ud_ezi_puntodeventa_encabezado.importeglobal AS Importe, ud_ezi_puntodeventa_encabezado.id, " & _
                      "ud_ezi_puntodeventa_encabezado.generada, ud_ezi_puntodeventa_encabezado.tipodefactura, ud_ezi_puntodeventa_encabezado.nota, " & _
                      "ud_ezi_puntodeventa_encabezado.numerodefactura + 'n ' + ud_ezi_puntodeventa_encabezado.cliente + ' ' + V_PERSONA_.CUIT + ' ' + ud_ezi_puntodeventa_encabezado.vendedor " & _
                      "AS concatenado, prep.claveprimaria AS Preparadonro,ud_ezi_puntodeventa_encabezado.id,  ud_ezi_puntodeventa_encabezado.adicionalid as OC FROM         ud_ezi_puntodeventa_encabezado AS prep WITH (readpast) RIGHT OUTER JOIN  " & _
                      "ud_ezi_puntodeventa_encabezado WITH (readpast) ON prep.presupuestobase = ud_ezi_puntodeventa_encabezado.claveprimaria LEFT OUTER JOIN " & _
                      "V_PERSONA_ RIGHT OUTER JOIN V_CLIENTE_ ON V_PERSONA_.ID = V_CLIENTE_.ENTEASOCIADO_ID ON ud_ezi_puntodeventa_encabezado.clienteid = V_CLIENTE_.ID " & _
                      "WHERE     (ud_ezi_puntodeventa_encabezado.numeradorinterno = 'Nota de Venta') AND (prep.estado = 'Preparado') AND " & _
                      "(ud_ezi_puntodeventa_encabezado.numerodefactura + 'n ' + ud_ezi_puntodeventa_encabezado.cliente + ' ' + V_PERSONA_.CUIT + ' ' + ud_ezi_puntodeventa_encabezado.vendedor like '" & xbusqueda & "')  and ud_ezi_puntodeventa_encabezado.adicionalid = '" & Text2.Text & "' ORDER BY Fecha DESC"
           End If
        End If


        datpresupuesto.RecordSource = xquery1
        datpresupuesto.Refresh
        If datpresupuesto.Recordset.EOF = False Then datpresupuesto.Recordset.MoveFirst

        Call Command1_Click
            If msgrid1.Rows >= 1 Then
            msgrid1.Row = 1
            msgrid1.Col = 1
        End If

        Call msgrid1_Click
        
      End If
      Call Command1_Click
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
