VERSION 5.00
Object = "{F6125AB1-8AB1-11CE-A77F-08002B2F4E98}#2.0#0"; "MSRDC20.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{D18BBD1F-82BB-4385-BED3-E9D31A3E361E}#1.0#0"; "KewlButtonz.ocx"
Begin VB.Form frmnota_venta 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NOTA DE VENTA"
   ClientHeight    =   8685
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   16095
   Icon            =   "frmnota_venta.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8685
   ScaleWidth      =   16095
   Begin VB.CommandButton imprimefactura 
      Caption         =   "imprimefactura"
      Height          =   255
      Left            =   3480
      TabIndex        =   75
      Top             =   2160
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00C0FFFF&
      Height          =   315
      Left            =   4080
      OLEDragMode     =   1  'Automatic
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   2160
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton UM 
      Caption         =   "UM"
      Height          =   315
      Left            =   4440
      TabIndex        =   43
      Top             =   1920
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton calcula 
      Caption         =   "calcula"
      Height          =   315
      Left            =   2400
      TabIndex        =   42
      Top             =   2040
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton ubicatextogrilla 
      Caption         =   "ubicatextogrilla"
      Height          =   315
      Left            =   4440
      TabIndex        =   41
      Top             =   2400
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5040
      TabIndex        =   11
      Text            =   "Text2"
      Top             =   2160
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Frame Frame2 
      Caption         =   "Nro. Cotización"
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
      Height          =   975
      Left            =   240
      TabIndex        =   31
      Top             =   7440
      Width           =   4095
      Begin VB.TextBox Text18 
         Height          =   285
         Left            =   1800
         TabIndex        =   70
         Text            =   "Text18"
         Top             =   600
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton cargapresupuesto 
         Caption         =   "cargapresupuesto"
         Height          =   315
         Left            =   2160
         TabIndex        =   69
         Top             =   120
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.TextBox Text17 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   360
         TabIndex        =   68
         Top             =   360
         Width           =   1335
      End
      Begin KewlButtonz.KewlButtons buscar 
         Height          =   495
         Left            =   2400
         TabIndex        =   32
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   873
         BTYPE           =   14
         TX              =   "&Buscar"
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
         MICON           =   "frmnota_venta.frx":0442
         PICN            =   "frmnota_venta.frx":045E
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
   Begin MSAdodcLib.Adodc datvendedor 
      Height          =   330
      Left            =   9720
      Top             =   6720
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
   Begin MSAdodcLib.Adodc datcliente 
      Height          =   330
      Left            =   9720
      Top             =   6360
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
   Begin MSAdodcLib.Adodc datmovimientos 
      Height          =   330
      Left            =   9720
      Top             =   6120
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
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
   Begin MSAdodcLib.Adodc datproductos 
      Height          =   330
      Left            =   9720
      Top             =   5880
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
   Begin VB.Frame Frame3 
      Caption         =   "Opciones"
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
      Height          =   975
      Left            =   4560
      TabIndex        =   22
      Top             =   7440
      Width           =   4815
      Begin KewlButtonz.KewlButtons salir 
         Height          =   495
         Left            =   3480
         TabIndex        =   21
         Top             =   360
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
         MICON           =   "frmnota_venta.frx":09F0
         PICN            =   "frmnota_venta.frx":0A0C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin KewlButtonz.KewlButtons grabar 
         Height          =   495
         Left            =   240
         TabIndex        =   19
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   873
         BTYPE           =   14
         TX              =   "&Grabar (F10)"
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
         MICON           =   "frmnota_venta.frx":1556
         PICN            =   "frmnota_venta.frx":1572
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
         Left            =   2160
         TabIndex        =   20
         Top             =   360
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   873
         BTYPE           =   14
         TX              =   "&Cancelar"
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
         MICON           =   "frmnota_venta.frx":2FF4
         PICN            =   "frmnota_venta.frx":3010
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
   Begin VB.PictureBox Picture1 
      Height          =   8535
      Left            =   120
      ScaleHeight     =   8475
      ScaleWidth      =   15795
      TabIndex        =   24
      Top             =   120
      Width           =   15855
      Begin VB.Frame Frame4 
         Caption         =   "Anotaciones de Clientes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4815
         Left            =   5880
         TabIndex        =   93
         Top             =   2040
         Visible         =   0   'False
         Width           =   9735
         Begin VB.TextBox Text20 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3975
            Left            =   240
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   94
            Top             =   600
            Width           =   9375
         End
         Begin VB.CommandButton Command9 
            Caption         =   "X"
            Height          =   375
            Left            =   9120
            TabIndex        =   95
            Top             =   120
            Width           =   495
         End
      End
      Begin VB.CommandButton acomodaitems 
         Caption         =   "acomodaitems"
         Height          =   255
         Left            =   9720
         TabIndex        =   91
         Top             =   7800
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox Text22 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   10800
         Locked          =   -1  'True
         TabIndex        =   89
         Top             =   7440
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker DFECHA 
         Height          =   375
         Left            =   1440
         TabIndex        =   88
         Top             =   120
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   99614721
         CurrentDate     =   43157
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1440
         TabIndex        =   33
         Top             =   120
         Visible         =   0   'False
         Width           =   1815
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
         Height          =   435
         Index           =   9
         Left            =   9720
         MaxLength       =   300
         TabIndex        =   87
         Top             =   1560
         Width           =   1575
      End
      Begin VB.CommandButton verificalotes 
         Caption         =   "verificalotes"
         Height          =   255
         Left            =   120
         TabIndex        =   84
         Top             =   1440
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Lotes"
         Height          =   255
         Left            =   5520
         TabIndex        =   83
         Top             =   2640
         Width           =   1695
      End
      Begin VB.CommandButton limitecredito 
         Caption         =   "limitecredito"
         Height          =   255
         Left            =   9480
         TabIndex        =   82
         Top             =   1440
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CommandButton Command7 
         Caption         =   "&Historial de Venta"
         Height          =   255
         Left            =   3720
         TabIndex        =   81
         Top             =   2640
         Width           =   1695
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "frmnota_venta.frx":3A22
         Height          =   495
         Left            =   1680
         TabIndex        =   35
         Top             =   -120
         Visible         =   0   'False
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   873
         _Version        =   393216
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
         Height          =   435
         Index           =   8
         Left            =   13800
         MaxLength       =   8
         TabIndex        =   79
         Top             =   2040
         Width           =   1815
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
         Height          =   435
         Index           =   7
         Left            =   13800
         MaxLength       =   8
         TabIndex        =   77
         Top             =   1560
         Width           =   1815
      End
      Begin VB.CommandButton imprimeremito 
         Caption         =   "imprimeremito"
         Height          =   255
         Left            =   3240
         TabIndex        =   74
         Top             =   2280
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.TextBox Text19 
         Height          =   285
         Left            =   240
         TabIndex        =   73
         Top             =   1800
         Visible         =   0   'False
         Width           =   1215
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
         Height          =   375
         Index           =   6
         Left            =   9240
         MaxLength       =   300
         TabIndex        =   72
         Top             =   600
         Visible         =   0   'False
         Width           =   4935
      End
      Begin VB.CommandButton grremito 
         Caption         =   "grremito"
         Height          =   315
         Left            =   10080
         TabIndex        =   67
         Top             =   2280
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.CommandButton grfacturactacte 
         Caption         =   "grfacturactacte"
         Height          =   315
         Left            =   11640
         TabIndex        =   66
         Top             =   2400
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.TextBox Text16 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   13680
         Locked          =   -1  'True
         TabIndex        =   63
         TabStop         =   0   'False
         Top             =   6720
         Width           =   1815
      End
      Begin VB.Frame Frame1 
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
         Height          =   1575
         Left            =   120
         TabIndex        =   55
         Top             =   5520
         Width           =   9975
         Begin VB.TextBox Text15 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1080
            Left            =   4800
            MaxLength       =   300
            MultiLine       =   -1  'True
            TabIndex        =   18
            Top             =   360
            Width           =   5055
         End
         Begin VB.CommandButton Command6 
            Caption         =   "Observaciones al Pie"
            Height          =   255
            Left            =   4800
            TabIndex        =   61
            TabStop         =   0   'False
            Top             =   120
            Width           =   1935
         End
         Begin VB.TextBox Text14 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   3240
            TabIndex        =   17
            Top             =   720
            Width           =   1335
         End
         Begin VB.CommandButton Command5 
            Caption         =   "Importe Global $"
            Height          =   255
            Left            =   3240
            TabIndex        =   60
            TabStop         =   0   'False
            Top             =   480
            Width           =   1335
         End
         Begin VB.TextBox Text13 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1680
            TabIndex        =   16
            Top             =   1080
            Width           =   1335
         End
         Begin VB.TextBox Text12 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   240
            TabIndex        =   15
            Top             =   1080
            Width           =   1335
         End
         Begin VB.TextBox Text11 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1680
            TabIndex        =   14
            Top             =   480
            Width           =   1335
         End
         Begin VB.TextBox Text10 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   240
            TabIndex        =   13
            Top             =   480
            Width           =   1335
         End
         Begin VB.CommandButton Command4 
            Caption         =   "Recargo $"
            Height          =   255
            Left            =   1680
            TabIndex        =   59
            TabStop         =   0   'False
            Top             =   840
            Width           =   1335
         End
         Begin VB.CommandButton Command3 
            Caption         =   "Bonif. Global $"
            Height          =   255
            Left            =   1680
            TabIndex        =   58
            TabStop         =   0   'False
            Top             =   240
            Width           =   1335
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Recargo %"
            Height          =   255
            Left            =   240
            TabIndex        =   57
            TabStop         =   0   'False
            Top             =   840
            Width           =   1335
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Bonif. Global %"
            Height          =   255
            Left            =   240
            TabIndex        =   56
            TabStop         =   0   'False
            Top             =   240
            Width           =   1335
         End
      End
      Begin VB.TextBox Text9 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   13680
         Locked          =   -1  'True
         TabIndex        =   54
         TabStop         =   0   'False
         Top             =   7440
         Width           =   1815
      End
      Begin VB.TextBox Text8 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   13680
         Locked          =   -1  'True
         TabIndex        =   53
         TabStop         =   0   'False
         Top             =   7080
         Width           =   1815
      End
      Begin VB.TextBox Text7 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   13680
         Locked          =   -1  'True
         TabIndex        =   50
         TabStop         =   0   'False
         Top             =   6360
         Width           =   1815
      End
      Begin VB.TextBox Text6 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   13680
         Locked          =   -1  'True
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   6000
         Width           =   1815
      End
      Begin VB.TextBox Text5 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   13680
         Locked          =   -1  'True
         TabIndex        =   46
         TabStop         =   0   'False
         Top             =   5640
         Width           =   1815
      End
      Begin VB.TextBox Text4 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   13680
         Locked          =   -1  'True
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   7920
         Width           =   1815
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grilla 
         Height          =   2895
         Left            =   120
         TabIndex        =   10
         Top             =   2640
         Width           =   15615
         _ExtentX        =   27543
         _ExtentY        =   5106
         _Version        =   393216
         BackColor       =   16777215
         Rows            =   300
         Cols            =   30
         ScrollBars      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   30
         _Band(0).GridLinesBand=   2
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
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
         Height          =   435
         Index           =   5
         Left            =   1440
         MaxLength       =   300
         TabIndex        =   7
         Top             =   1560
         Width           =   6495
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   3
         Left            =   1440
         MaxLength       =   300
         MultiLine       =   -1  'True
         TabIndex        =   5
         Top             =   1080
         Width           =   7695
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
         Height          =   360
         Index           =   1
         Left            =   1440
         MaxLength       =   50
         TabIndex        =   2
         Top             =   600
         Width           =   5295
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   4
         Left            =   12240
         TabIndex        =   37
         Top             =   120
         Width           =   1935
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   10320
         TabIndex        =   23
         Top             =   120
         Width           =   855
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
         Height          =   360
         Index           =   0
         Left            =   4680
         MaxLength       =   50
         TabIndex        =   0
         Top             =   120
         Width           =   3375
      End
      Begin MSDataListLib.DataCombo DataCombo3 
         Bindings        =   "frmnota_venta.frx":3A3C
         Height          =   360
         Left            =   9120
         TabIndex        =   4
         Top             =   600
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   741
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         ListField       =   "tipopago"
         BoundColumn     =   "id"
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin KewlButtonz.KewlButtons bclientes 
         Height          =   375
         Left            =   6720
         TabIndex        =   3
         ToolTipText     =   "F3-Busca Clientes"
         Top             =   600
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   661
         BTYPE           =   14
         TX              =   ""
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
         MICON           =   "frmnota_venta.frx":3A56
         PICN            =   "frmnota_venta.frx":3A72
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin KewlButtonz.KewlButtons bvendedor 
         Height          =   375
         Left            =   8040
         TabIndex        =   1
         ToolTipText     =   "F3-Busca Clientes"
         Top             =   120
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   661
         BTYPE           =   14
         TX              =   ""
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
         MICON           =   "frmnota_venta.frx":400C
         PICN            =   "frmnota_venta.frx":4028
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "frmnota_venta.frx":45C2
         Height          =   360
         Left            =   11160
         TabIndex        =   6
         Top             =   1080
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   741
         _Version        =   393216
         Style           =   2
         ListField       =   "nombre"
         BoundColumn     =   "id"
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin KewlButtonz.KewlButtons agregaproducto 
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   2160
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         BTYPE           =   14
         TX              =   "&Agregar Item (F5)"
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
         MICON           =   "frmnota_venta.frx":45E0
         PICN            =   "frmnota_venta.frx":45FC
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
         Height          =   375
         Left            =   2280
         TabIndex        =   9
         Top             =   2160
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         BTYPE           =   14
         TX              =   "&Eliminar Item"
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
         MICON           =   "frmnota_venta.frx":4B96
         PICN            =   "frmnota_venta.frx":4BB2
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSAdodcLib.Adodc datiibb 
         Height          =   330
         Left            =   240
         Top             =   7800
         Visible         =   0   'False
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   582
         ConnectMode     =   0
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
         Left            =   11160
         Top             =   6240
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
      Begin KewlButtonz.KewlButtons blanco 
         Height          =   375
         Left            =   7920
         TabIndex        =   64
         TabStop         =   0   'False
         ToolTipText     =   "F3-Busca Clientes"
         Top             =   1560
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   661
         BTYPE           =   14
         TX              =   ""
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
         BCOL            =   49152
         BCOLO           =   49152
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmnota_venta.frx":514C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin KewlButtonz.KewlButtons negro 
         Height          =   375
         Left            =   7920
         TabIndex        =   65
         TabStop         =   0   'False
         ToolTipText     =   "F3-Busca Clientes"
         Top             =   1560
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   661
         BTYPE           =   14
         TX              =   ""
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
         BCOL            =   192
         BCOLO           =   192
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmnota_venta.frx":5168
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSAdodcLib.Adodc datitems 
         Height          =   330
         Left            =   11160
         Top             =   6120
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
      Begin MSAdodcLib.Adodc datitempresup 
         Height          =   330
         Left            =   9480
         Top             =   7080
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
         Bindings        =   "frmnota_venta.frx":5184
         Height          =   495
         Left            =   2880
         TabIndex        =   36
         Top             =   600
         Visible         =   0   'False
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   873
         _Version        =   393216
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
      Begin MSAdodcLib.Adodc datcontrol 
         Height          =   330
         Left            =   11040
         Top             =   5760
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
      Begin MSAdodcLib.Adodc datcola 
         Height          =   330
         Left            =   11640
         Top             =   6600
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
      Begin MSAdodcLib.Adodc datcolaimportar 
         Height          =   330
         Left            =   11040
         Top             =   6000
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
      Begin MSAdodcLib.Adodc datpago 
         Height          =   330
         Left            =   8760
         Top             =   5880
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
      Begin MSRDC.MSRDC reporte 
         Height          =   375
         Left            =   10920
         Top             =   5760
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
         Left            =   0
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowTitle     =   "Libro IVA Compras"
         PrintFileLinesPerPage=   60
      End
      Begin MSAdodcLib.Adodc datitemsnv 
         Height          =   330
         Left            =   10320
         Top             =   5640
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
         Caption         =   "datitemsnv"
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
      Begin MSAdodcLib.Adodc datcredito 
         Height          =   330
         Left            =   11640
         Top             =   6120
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
      Begin KewlButtonz.KewlButtons cancelar2 
         Height          =   495
         Left            =   240
         TabIndex        =   85
         Top             =   2640
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   873
         BTYPE           =   14
         TX              =   "Cancelar2"
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
         MICON           =   "frmnota_venta.frx":519D
         PICN            =   "frmnota_venta.frx":51B9
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
         Left            =   14280
         TabIndex        =   92
         Top             =   960
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   873
         BTYPE           =   14
         TX              =   "Info Cliente"
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
         MICON           =   "frmnota_venta.frx":5BCB
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Cant.Items:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   29
         Left            =   9360
         TabIndex        =   90
         Top             =   7440
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "O.C. Nro:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   20
         Left            =   8520
         TabIndex        =   86
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente Precio Especial"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Index           =   19
         Left            =   6720
         TabIndex        =   80
         Top             =   2160
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Nro.Rem.Manual:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   18
         Left            =   11640
         TabIndex        =   78
         Top             =   2160
         Width           =   2055
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Nro.Fac.Manual:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   17
         Left            =   11760
         TabIndex        =   76
         Top             =   1560
         Width           =   1935
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   16
         Left            =   7920
         TabIndex        =   71
         Top             =   600
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Subtotal 2:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   15
         Left            =   11640
         TabIndex        =   62
         Top             =   6720
         Width           =   1935
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Percep IIBB:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   14
         Left            =   10680
         TabIndex        =   52
         Top             =   7440
         Width           =   2895
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Tem/Pyp:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   13
         Left            =   11880
         TabIndex        =   51
         Top             =   7080
         Width           =   1695
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Iva 21%:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   12
         Left            =   11880
         TabIndex        =   48
         Top             =   6360
         Width           =   1695
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Iva 10.5%:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   11
         Left            =   11880
         TabIndex        =   47
         Top             =   6000
         Width           =   1695
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Subtotal:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   10
         Left            =   11640
         TabIndex        =   45
         Top             =   5640
         Width           =   1935
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Nota:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   9
         Left            =   120
         TabIndex        =   40
         Top             =   1560
         Width           =   3015
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Lista de Precio:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   8
         Left            =   9240
         TabIndex        =   39
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Domicilio:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   38
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   120
         TabIndex        =   34
         Top             =   120
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Importe Total:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   11640
         TabIndex        =   30
         Top             =   7920
         Width           =   1935
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "C.U.I.T:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   11280
         TabIndex        =   29
         Top             =   120
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo Compr.:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   7
         Left            =   8760
         TabIndex        =   28
         Top             =   120
         Width           =   1695
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   14760
         Y1              =   2040
         Y2              =   2040
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   120
         TabIndex        =   27
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo de Pago:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   7440
         TabIndex        =   26
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Vendedor:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   6
         Left            =   3360
         TabIndex        =   25
         Top             =   120
         Width           =   1335
      End
   End
   Begin MSAdodcLib.Adodc dattipopago 
      Height          =   330
      Left            =   9960
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
   Begin MSAdodcLib.Adodc datparametros 
      Height          =   330
      Left            =   12600
      Top             =   8280
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
   Begin MSAdodcLib.Adodc datlistaprecios 
      Height          =   330
      Left            =   11760
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
   Begin MSAdodcLib.Adodc datum 
      Height          =   330
      Left            =   13800
      Top             =   8160
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
End
Attribute VB_Name = "frmnota_venta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public modo As String
Public usuariomanual As String
Dim xfaconv(10) As Double
Public xumvta As String
Public xalicuotaiibb As Double
Public xexentoiibb As Double
Public xcalculaiibb As String
Public xcalculatempyp As String
Public xalicuptatempip As String
Public xciudadtem As String
Public xciudadcliente As String
Public tipofac As String
Public presupuestobase As Double
Public xlineasmax As Integer
Public xremito As Double
Public xid As Double
Public tpago As String
Public xcontroltem As Double
Public xlimitebonif As Double
Public xdisponible As Double
Public xcreditomaximo As Double
Public xlimitecredito As Double
Public xvendedorautoriza As String
Public xidpre As Double
Public xmodifica As Integer
Dim xclienteinfo As String
Dim xcontrol As Integer



Private Sub acomodaitems_Click()
On Error Resume Next
Dim oCmd As Command
Set oCmd = New Command

oCmd.ActiveConnection = login.conexiontotal

oCmd.CommandText = "update ud_ezi_puntodeventa_detalle_rem set item =  ud_ezi_puntodeventa_detalle_notav.item " & _
                           "FROM         ud_ezi_puntodeventa_encabezado INNER JOIN  " & _
                           "ud_ezi_puntodeventa_detalle_notav ON ud_ezi_puntodeventa_encabezado.id = ud_ezi_puntodeventa_detalle_notav.claveprimaria LEFT OUTER JOIN " & _
                           "ud_ezi_puntodeventa_detalle_rem ON ud_ezi_puntodeventa_detalle_notav.item <> ud_ezi_puntodeventa_detalle_rem.item AND " & _
                           "ud_ezi_puntodeventa_detalle_notav.cantidadproducto = ud_ezi_puntodeventa_detalle_rem.cantidadoriginal AND " & _
                           "ud_ezi_puntodeventa_encabezado.claveprimaria = ud_ezi_puntodeventa_detalle_rem.facturaorigen AND " & _
                           "ud_ezi_puntodeventa_detalle_notav.idproducto = ud_ezi_puntodeventa_detalle_rem.idproducto " & _
                           "WHERE    (ud_ezi_puntodeventa_encabezado.numeradorinterno = 'Nota de Venta') AND " & _
                           "(NOT (ud_ezi_puntodeventa_detalle_rem.item IS NULL))"
oCmd.Execute



End Sub

Private Sub agregaproducto_Click()
On Error Resume Next
    
        If Text1(0).Text = "" Then
            mensa = MsgBox("Debe ingresar un vendedor", vbCritical, "Error")
            Exit Sub
        End If
        If Text1(1).Text = "" Then
            mensa = MsgBox("Debe ingresar un Cliente", vbCritical, "Error")
            Exit Sub
        End If
    


     menu = 1
     query = "SELECT top 10 p.ID, p.CODIGO, p.DESCRIPCION, v.DENOMINACION AS proveedor, " & _
                      "CASE i.CODIGO WHEN 'I' THEN 'Internacional' WHEN 'N' THEN 'Nacional' ELSE '' END AS Nacionalidad, u.DETALLE AS rubro, " & _
                      "ROUND(CAST(r.PRECIOCTA AS decimal(14, 3)), 3) AS precio, r.CODPROVEEDOR, t.CODIGO AS marca, " & _
                      "p.CODIGO + p.DESCRIPCION+isnull(v.DENOMINACION,'')+CASE i.CODIGO WHEN 'I' THEN 'Internacional' WHEN 'N' THEN 'Nacional' ELSE '' END + isnull(u.DETALLE,'')+isnull(r.CODPROVEEDOR,'')+isnull(t.CODIGO,'') as concatenado, V_UNIDADMEDIDA_.NOMBRE AS um " & _
                      "FROM V_PRODUCTO_ AS p LEFT OUTER JOIN V_UNIDADMEDIDA_ ON p.UNIDADMEDIDANOLINEAL_ID = V_UNIDADMEDIDA_.ID LEFT OUTER JOIN " & _
                      "V_UD_EZI_PRODUCTOS_ AS r ON p.BOEXTENSION_ID = r.ID LEFT OUTER JOIN " & _
                      "V_PROVEEDOR_ AS v ON r.PROVEEDOR_ID = v.ID LEFT OUTER JOIN " & _
                      "V_ITEMTIPOCLASIFICADOR_ AS i ON r.NACIONALIDAD_ID = i.ID LEFT OUTER JOIN " & _
                      "V_ITEMTIPOCLASIFICADOR_ AS t ON r.MARCA_ID = t.ID LEFT OUTER JOIN " & _
                      "V_RUBRO_ AS u ON p.RUBRO_ID = u.ID " & _
                      "Where (p.ACTIVESTATUS = 0) And (p.TIPOOBJETOESTATICO_ID Is Null) " & _
                      "ORDER BY p.DESCRIPCION"
                      
    
    For X = 1 To xlineasmax
        grilla.Col = 1
        grilla.Row = X
        If grilla.Text = "" Then
           xfila = X
           lista_productos_colon.Show
           Exit For
        End If
    Next X
    
    

End Sub

Private Sub bclientes_Click()
On Error Resume Next

   xcodclientefiltra = datparametros.Recordset.Fields("codclientefiltra")
   If xcodclientefiltra = "01" Then xcontrocliente = "88447B8E-14FE-4D60-9622-B22F6C735701"  ' tucuman
   If xcodclientefiltra = "04" Then xcontrocliente = "4234CA46-B2BE-4690-AC6A-F0DE206F94A9"  ' salta
   If xcodclientefiltra = "03" Then xcontrocliente = "AEC7FBAC-63F7-4404-9512-033D0961D9BC"  ' jujuy
    
  If Text1(1).Text <> "" Then
   If Text19.Text = "" Then
    Text1(1).Text = Replace(Text1(1).Text, " ", "%%")
    xbusqueda = "%" + Text1(1).Text + "%"
    xquery1 = "SELECT     ALIAS_0.ID, ALIAS_0.CODIGO, ALIAS_3.NOMBRE AS RAZONSOCIAL, ALIAS_3.CUIT, ALIAS_6.CALLE, ISNULL(ALIAS_6.CALLE, '') + '-' + ISNULL(V_CIUDAD_.NOMBRE, '') " & _
              "+ '-' + ISNULL(V_PROVINCIA_.NOMBRE, '') AS DOMICILIO, ALIAS_0.DENOMINACION, ALIAS_5.NOMBRE AS ZONA, ALIAS_7.NUMERO AS TELEFONO, " & _
              "ALIAS_8.DIRECCIONELECTRONICA AS MAIL, V_TIPOPAGO_.NOMBRE AS TP, V_TIPOPAGO_.OBSERVACION AS TipoPago, " & _
              "V_TIPOPAGO_.ACTIVESTATUS AS ACTIVESTATUSTP, V_PROVINCIA_.NOMBRE AS Provincia, V_CIUDAD_.NOMBRE AS Ciudad, " & _
              "ALIAS_0.CODIGO + ' ' + ALIAS_3.CUIT + ' ' + ALIAS_3.NOMBRE AS CONCATENADO, V_EZI_POCICION_IVA_CLIENTES.CODIGO AS IVA, V_TIPOPAGO_.ID AS IDPAGO, " & _
              "ALIAS_0.DOMICILIOFACTURACION_ID AS domicilio_id, ALIAS_0.LISTAPRECIO_ID AS listaprecio, V_UD_CLIENTE.observacion, ALIAS_0.creditomaximo, alias_0.diasplazo, V_UD_CLIENTE.Anotaciones  " & _
              "FROM         V_TIPOPAGO_ RIGHT OUTER JOIN " & _
              "V_CLIENTE AS ALIAS_0 WITH (NOLOCK) LEFT OUTER JOIN " & _
              "V_UD_CLIENTE with (nolock) ON ALIAS_0.BOEXTENSION_ID = V_UD_CLIENTE.ID LEFT OUTER JOIN " & _
              "V_EZI_POCICION_IVA_CLIENTES ON ALIAS_0.ID = V_EZI_POCICION_IVA_CLIENTES.idcliente ON V_TIPOPAGO_.ID = ALIAS_0.TIPOPAGO_ID LEFT OUTER JOIN " & _
              "V_PERSONA AS ALIAS_3 WITH (nolock) ON ALIAS_0.ENTEASOCIADO_ID = ALIAS_3.ID LEFT OUTER JOIN " & _
              "V_CIUDAD_ RIGHT OUTER JOIN " & _
              "V_DOMICILIO_ AS ALIAS_6 ON V_CIUDAD_.ID = ALIAS_6.CIUDAD_ID LEFT OUTER JOIN " & _
              "V_PROVINCIA_ ON ALIAS_6.PROVINCIA_ID = V_PROVINCIA_.ID ON ALIAS_0.DOMICILIOFACTURACION_ID = ALIAS_6.ID LEFT OUTER JOIN " & _
              "V_TELEFONO_ AS ALIAS_7 ON ALIAS_3.TELEFONOPRINCIPAL_ID = ALIAS_7.ID LEFT OUTER JOIN " & _
              "V_DIRECCIONELECTRONICA_ AS ALIAS_8 ON ALIAS_3.DIRECELECTRONICAPRINCIPAL_ID = ALIAS_8.ID LEFT OUTER JOIN " & _
              "V_ZONA_ AS ALIAS_5 ON ALIAS_0.ZONA_ID = ALIAS_5.ID " & _
              "WHERE     (ALIAS_0.ACTIVESTATUS = 0) AND ALIAS_0.CODIGO + ' ' + ALIAS_3.CUIT + ' ' + ALIAS_3.NOMBRE + ' ' + ALIAS_0.DENOMINACION  like '" & xbusqueda & "'  " & _
              "order by ALIAS_3.NOMBRE "
   Else
    xquery1 = "SELECT     ALIAS_0.ID, ALIAS_0.CODIGO, ALIAS_3.NOMBRE AS RAZONSOCIAL, ALIAS_3.CUIT, ALIAS_6.CALLE, ISNULL(ALIAS_6.CALLE, '') + '-' + ISNULL(V_CIUDAD_.NOMBRE, '') " & _
              "+ '-' + ISNULL(V_PROVINCIA_.NOMBRE, '') AS DOMICILIO, ALIAS_0.DENOMINACION, ALIAS_5.NOMBRE AS ZONA, ALIAS_7.NUMERO AS TELEFONO, " & _
              "ALIAS_8.DIRECCIONELECTRONICA AS MAIL, V_TIPOPAGO_.NOMBRE AS TP, V_TIPOPAGO_.OBSERVACION AS TipoPago, " & _
              "V_TIPOPAGO_.ACTIVESTATUS AS ACTIVESTATUSTP, V_PROVINCIA_.NOMBRE AS Provincia, V_CIUDAD_.NOMBRE AS Ciudad, " & _
              "ALIAS_0.CODIGO + ' ' + ALIAS_3.CUIT + ' ' + ALIAS_3.NOMBRE AS CONCATENADO, V_EZI_POCICION_IVA_CLIENTES.CODIGO AS IVA, V_TIPOPAGO_.ID AS IDPAGO, " & _
              "ALIAS_0.DOMICILIOFACTURACION_ID AS domicilio_id, ALIAS_0.LISTAPRECIO_ID AS listaprecio, V_UD_CLIENTE.observacion, ALIAS_0.creditomaximo, alias_0.diasplazo, V_UD_CLIENTE.Anotaciones  " & _
              "FROM         V_TIPOPAGO_ RIGHT OUTER JOIN " & _
              "V_CLIENTE AS ALIAS_0 WITH (NOLOCK) LEFT OUTER JOIN " & _
              "V_UD_CLIENTE with (nolock) ON ALIAS_0.BOEXTENSION_ID = V_UD_CLIENTE.ID LEFT OUTER JOIN " & _
              "V_EZI_POCICION_IVA_CLIENTES ON ALIAS_0.ID = V_EZI_POCICION_IVA_CLIENTES.idcliente ON V_TIPOPAGO_.ID = ALIAS_0.TIPOPAGO_ID LEFT OUTER JOIN " & _
              "V_PERSONA AS ALIAS_3 WITH (nolock) ON ALIAS_0.ENTEASOCIADO_ID = ALIAS_3.ID LEFT OUTER JOIN " & _
              "V_CIUDAD_ RIGHT OUTER JOIN " & _
              "V_DOMICILIO_ AS ALIAS_6 ON V_CIUDAD_.ID = ALIAS_6.CIUDAD_ID LEFT OUTER JOIN " & _
              "V_PROVINCIA_ ON ALIAS_6.PROVINCIA_ID = V_PROVINCIA_.ID ON ALIAS_0.DOMICILIOFACTURACION_ID = ALIAS_6.ID LEFT OUTER JOIN " & _
              "V_TELEFONO_ AS ALIAS_7 ON ALIAS_3.TELEFONOPRINCIPAL_ID = ALIAS_7.ID LEFT OUTER JOIN " & _
              "V_DIRECCIONELECTRONICA_ AS ALIAS_8 ON ALIAS_3.DIRECELECTRONICAPRINCIPAL_ID = ALIAS_8.ID LEFT OUTER JOIN " & _
              "V_ZONA_ AS ALIAS_5 ON ALIAS_0.ZONA_ID = ALIAS_5.ID " & _
              "WHERE   (ALIAS_0.ACTIVESTATUS = 0) AND  ALIAS_0.ID = '" & Text19.Text & "' "
   End If
  Else
    xquery1 = "SELECT     ALIAS_0.ID, ALIAS_0.CODIGO, ALIAS_3.NOMBRE AS RAZONSOCIAL, ALIAS_3.CUIT, ALIAS_6.CALLE, ISNULL(ALIAS_6.CALLE, '') + '-' + ISNULL(V_CIUDAD_.NOMBRE, '') " & _
              "+ '-' + ISNULL(V_PROVINCIA_.NOMBRE, '') AS DOMICILIO, ALIAS_0.DENOMINACION, ALIAS_5.NOMBRE AS ZONA, ALIAS_7.NUMERO AS TELEFONO, " & _
              "ALIAS_8.DIRECCIONELECTRONICA AS MAIL, V_TIPOPAGO_.NOMBRE AS TP, V_TIPOPAGO_.OBSERVACION AS TipoPago, " & _
              "V_TIPOPAGO_.ACTIVESTATUS AS ACTIVESTATUSTP, V_PROVINCIA_.NOMBRE AS Provincia, V_CIUDAD_.NOMBRE AS Ciudad, " & _
              "ALIAS_0.CODIGO + ' ' + ALIAS_3.CUIT + ' ' + ALIAS_3.NOMBRE AS CONCATENADO, V_EZI_POCICION_IVA_CLIENTES.CODIGO AS IVA, V_TIPOPAGO_.ID AS IDPAGO, " & _
              "ALIAS_0.DOMICILIOFACTURACION_ID AS domicilio_id, ALIAS_0.LISTAPRECIO_ID AS listaprecio, V_UD_CLIENTE.observacion, ALIAS_0.creditomaximo, alias_0.diasplazo, V_UD_CLIENTE.Anotaciones " & _
              "FROM         V_TIPOPAGO_ RIGHT OUTER JOIN " & _
              "V_CLIENTE AS ALIAS_0 WITH (NOLOCK) LEFT OUTER JOIN " & _
              "V_UD_CLIENTE with (nolock) ON ALIAS_0.BOEXTENSION_ID = V_UD_CLIENTE.ID LEFT OUTER JOIN " & _
              "V_EZI_POCICION_IVA_CLIENTES ON ALIAS_0.ID = V_EZI_POCICION_IVA_CLIENTES.idcliente ON V_TIPOPAGO_.ID = ALIAS_0.TIPOPAGO_ID LEFT OUTER JOIN " & _
              "V_PERSONA AS ALIAS_3 WITH (nolock) ON ALIAS_0.ENTEASOCIADO_ID = ALIAS_3.ID LEFT OUTER JOIN " & _
              "V_CIUDAD_ RIGHT OUTER JOIN " & _
              "V_DOMICILIO_ AS ALIAS_6 ON V_CIUDAD_.ID = ALIAS_6.CIUDAD_ID LEFT OUTER JOIN " & _
              "V_PROVINCIA_ ON ALIAS_6.PROVINCIA_ID = V_PROVINCIA_.ID ON ALIAS_0.DOMICILIOFACTURACION_ID = ALIAS_6.ID LEFT OUTER JOIN " & _
              "V_TELEFONO_ AS ALIAS_7 ON ALIAS_3.TELEFONOPRINCIPAL_ID = ALIAS_7.ID LEFT OUTER JOIN " & _
              "V_DIRECCIONELECTRONICA_ AS ALIAS_8 ON ALIAS_3.DIRECELECTRONICAPRINCIPAL_ID = ALIAS_8.ID LEFT OUTER JOIN " & _
              "V_ZONA_ AS ALIAS_5 ON ALIAS_0.ZONA_ID = ALIAS_5.ID " & _
              "WHERE     (ALIAS_0.ACTIVESTATUS = 0) order by ALIAS_3.NOMBRE "
  End If

    Text19.Text = ""
    datcliente.RecordSource = xquery1
    datcliente.Refresh
    If datcliente.Recordset.EOF = True Then
        mensa = MsgBox("No existe Cliente", vbInformation, "!! Atencion !!")
        Text1(1).Text = ""
        Text1(1).SetFocus
    End If
    
    xclienteinfo = DataGrid2.Columns("anotaciones").Text
    
    If datcliente.Recordset.Fields("observacion") <> "" Or IsNull(datcliente.Recordset.Fields("observacion")) = False Then
      If datcliente.Recordset.Fields("observacion") <> "" Then
        MsgBox datcliente.Recordset.Fields("observacion"), vbInformation, "Mensaje de Gestión"
      End If
    End If
    
    datcliente.Recordset.MoveFirst
    If datcliente.Recordset.RecordCount = 1 Then
        Text1(1).Text = DataGrid2.Columns(2).Text
        Text1(5).Text = DataGrid2.Columns(6).Text
        If DataGrid2.Columns(16).Text = "RI" Then
            Text1(2).Text = "A"
        Else
            Text1(2).Text = "B"
        End If
        If DataGrid2.Columns(17).Text <> "" Then
            DataCombo3.BoundText = DataGrid2.Columns(17).Text
            testpos = InStr(1, DataCombo3.Text, "- ", 1)
            tpago = Right(DataCombo3.Text, Len(DataCombo3.Text) - testpos - 1)
        Else
            tpago = "CONTADO"
            DataCombo3.Text = "01- CONTADO"
        End If
        Text1(4).Text = DataGrid2.Columns(3).Text
        Text1(3).Text = DataGrid2.Columns(5).Text

'    If login.nomsucursal = "TUCUMAN" Then
'        If tpago = "CONTADO" Then
'            MsgBox "El sistema no se encuentra habilitado para facturar de Contado, se cambiara a cta.cte la condicion de este cliente solo para esta factura, si quiere hacerlo en forma permanente cambie en Calipso la condicion", vbCritical, "Error"
'            DataCombo3.Text = "02- CTA.CTE."
'            DataCombo3.BoundText = "{8982f2d7-d17c-44e3-a446-797a98d974fb}"
'            DataCombo3.Enabled = False
'        Else
'            DataCombo3.Enabled = True
'        End If
'    Else
        If UCase(tpago) = "CONTADO" Then
         If UCase(login.usuarioactivo) <> "ADMIN" Then
            DataCombo3.Enabled = False
         Else
            DataCombo3.Enabled = True
         End If
        Else
            DataCombo3.Enabled = True
        End If
'    End If
        
        If DataGrid2.Columns(16).Text = "CF" And Text1(1).Text = "CONSUMIDOR FINAL" Then
            Label1(5).Visible = False
            DataCombo3.Visible = False
            Label1(16).Visible = True
            Text1(6).Visible = True
            Text1(3).Enabled = True
            
            Label1(0).Caption = "D.N.I:"
            Text1(4).Enabled = True
            Text1(4).MaxLength = 8
            
            Text1(6).SetFocus
        Else
            Label1(5).Visible = True
            DataCombo3.Visible = True
            Label1(16).Visible = False
            Text1(6).Visible = False
            Text1(3).Enabled = False
            
            Label1(0).Caption = "C.U.I.T:"
            Text1(4).Enabled = False
            Text1(4).MaxLength = 20

            Text1(5).SetFocus
        End If
                
        xcontroltem = 1
        datiibb.RecordSource = "SELECT * from v_ezi_pos_impuestos " & _
                               "WHERE idcliente = '" & DataGrid2.Columns(0).Text & "'"
        datiibb.Refresh

        datiibb.Refresh
        
        
        
        If datiibb.Recordset.EOF = False Then
            xalicuotaiibb = datiibb.Recordset.Fields("COEFICIENTE")
'            If UCase(login.nomsucursal) <> "TUCUMAN" Then xalicuotaiibb = 0
            xexentoiibb = datiibb.Recordset.Fields("exencion")
            If xexentoiibb = 0 Then
                Label1(14).Caption = "Percep IIBB " + Str(xalicuotaiibb) + " %"
            Else
                Label1(14).Caption = "Percep IIBB " + Str(xalicuotaiibb) + " % Cert.Ex"
            End If
            If DataGrid2.Columns(16).Text = "CF" Then Label1(14).Caption = "Percep IIBB "
            
        End If
        xcontroltem = datiibb.Recordset.Fields("tem")
        xciudadcliente = DataGrid2.Columns("Ciudad").Text
        grilla.Row = 1
        Call calcula_Click
        
    Else
        menu = 1
        query = xquery1
        lista_clientes.Show
    End If

'' Evalua limite de credito *************
  blanco.Visible = False
  negro.Visible = False
  If UCase(DataCombo3.Text) <> "01- CONTADO" Then
    Call limitecredito_Click
  End If


End Sub

Private Sub blanco_Click()
On Error Resume Next

    negro.Visible = True
    blanco.Visible = False
    tipofac = "NN"
    Call calcula_Click

End Sub

Private Sub buscar_Click()
On Error Resume Next
    
    menu = 1
    lista_presupuestos.Show

End Sub

Private Sub bvendedor_Click()
On Error Resume Next
    
  If Text1(0).Text <> "" Then
    xbusqueda = "%" + Text1(0).Text + "%"
    xquery1 = "SELECT V_VENDEDOR_.ID, V_VENDEDOR_.CODIGO, V_PERSONA_.NOMBRE, V_VENDEDOR_.CODIGO + ' ' + V_PERSONA_.NOMBRE AS cancatena,ud_ezi_empleado.limite " & _
              "FROM V_VENDEDOR_ INNER JOIN V_PERSONA_ ON V_VENDEDOR_.ENTEASOCIADO_ID = V_PERSONA_.ID LEFT OUTER JOIN ud_ezi_empleado WITH (nolock) ON V_VENDEDOR_.ID = ud_ezi_empleado.id " & _
              "WHERE (V_VENDEDOR_.ACTIVESTATUS = 0)  and V_VENDEDOR_.CODIGO  + ' ' + V_PERSONA_.NOMBRE like '" & xbusqueda & "' order by V_PERSONA_.NOMBRE"

'    xquery1 = "SELECT    V_VENDEDOR_.ID, V_VENDEDOR_.CODIGO, V_PERSONA_.NOMBRE, V_VENDEDOR_.CODIGO  + ' ' + V_PERSONA_.NOMBRE as cancatena " & _
'              "FROM V_VENDEDOR_ INNER JOIN V_PERSONA_ ON V_VENDEDOR_.ENTEASOCIADO_ID = V_PERSONA_.ID " & _
'              "WHERE (V_VENDEDOR_.ACTIVESTATUS = 0)  and V_VENDEDOR_.CODIGO  + ' ' + V_PERSONA_.NOMBRE like '" & xbusqueda & "' order by V_PERSONA_.NOMBRE"
  Else
'    xquery1 = "SELECT    V_VENDEDOR_.ID, V_VENDEDOR_.CODIGO, V_PERSONA_.NOMBRE, V_VENDEDOR_.CODIGO  + ' ' + V_PERSONA_.NOMBRE as cancatena " & _
'              "FROM V_VENDEDOR_ INNER JOIN V_PERSONA_ ON V_VENDEDOR_.ENTEASOCIADO_ID = V_PERSONA_.ID " & _
'              "WHERE (V_VENDEDOR_.ACTIVESTATUS = 0)  order by V_PERSONA_.NOMBRE"

     xqueri1 = "SELECT V_VENDEDOR_.ID, V_VENDEDOR_.CODIGO, V_PERSONA_.NOMBRE, V_VENDEDOR_.CODIGO + ' ' + V_PERSONA_.NOMBRE AS cancatena,ud_ezi_empleado.limite " & _
               "FROM V_VENDEDOR_ INNER JOIN V_PERSONA_ ON V_VENDEDOR_.ENTEASOCIADO_ID = V_PERSONA_.ID LEFT OUTER JOIN ud_ezi_empleado WITH (nolock) ON V_VENDEDOR_.ID = ud_ezi_empleado.id " & _
               "WHERE (V_VENDEDOR_.ACTIVESTATUS = 0)  order by V_PERSONA_.NOMBRE"
  End If


    datvendedor.RecordSource = xquery1
    datvendedor.Refresh
    If datvendedor.Recordset.EOF = True Then
        mensa = MsgBox("No existe Vendedor", vbInformation, "!! Atencion !!")
        Text1(0).Text = ""
        Text1(0).SetFocus
    End If
 
    If datvendedor.Recordset.RecordCount = 1 Then
        Text1(0).Text = DataGrid1.Columns(2).Text
        Text1(1).SetFocus
        xlimitebonif = DataGrid1.Columns(4).Text
        If xlimitebonif = 100 Then xvendedorautoriza = Text1(0).Text
    Else
        menu = 1
        query = xquery1
        lista_vendedores.Show
        lista_vendedores.DataGrid1.SetFocus
    End If
    
    



End Sub

Private Sub calcula_Click()
On Error Resume Next

    If grilla.Row > 10 Then grilla.TopRow = grilla.Row - 9
    xcol = grilla.Col
    grilla.Col = 3
    xcant = Val(grilla.Text)
    
    grilla.Col = 6
    ximporte = grilla.Text
    grilla.Text = Format(ximporte, "###,##0.00")
    If grilla.Text = "0.00" And grilla.TextMatrix(grilla.Row, 0) = "" Then
        grilla.Text = ""
    End If
   
    ximportelista = Val(Format(grilla.TextMatrix(grilla.Row, 17), "0.00"))

    If ximporte < Val(Format(grilla.TextMatrix(grilla.Row, 17), "0.00")) And xlimitebonif <> 100 Then
     If Abs(100 - (ximporte / (Val(Format(grilla.TextMatrix(grilla.Row, 17), "0.00")))) * 100) > xlimitebonif Then
        MsgBox "No puede ingresar un importe menor al precio de lista", vbCritical, "Error"
        grilla.Text = Format(grilla.TextMatrix(grilla.Row, 17), "###,##0.00")
        Text2.Text = grilla.Text
        grilla.TextMatrix(grilla.Row, 5) = Format(Round(Val(Format(grilla.TextMatrix(grilla.Row, 17), "0.00")) / Val(Format(grilla.TextMatrix(grilla.Row, 12), "0.000")), 2), "###,##0.00")
        grilla.Col = xcol
        Call calcula_Click
        Exit Sub
      End If
    End If
        
    If ximporte = "" Then ximporte = 0
    grilla.Text = Format(ximporte, "###,##0.0000")
    If grilla.Text = "0.00" And grilla.TextMatrix(grilla.Row, 0) = "" Then
        grilla.Text = ""
    End If

    grilla.Col = 12
    xiva = grilla.Text
    If xiva = "" Then xiva = 1.21

'    ximportesiva2 = Round(ximporte / xiva, 5)
    ximportesiva2 = ximporte / xiva
    grilla.Col = 5
    grilla.Text = Format(ximportesiva2, "###,##0.0000")
    If grilla.Text = "0.000" And grilla.TextMatrix(grilla.Row, 0) = "" Then
        grilla.Text = ""
    End If

'    ximportesiva = Round(ximportesiva2 * xcant, 5)
    ximportesiva = ximporte / xiva * xcant

    grilla.Col = 7
    grilla.Text = Format(ximportesiva, "###,##0.000")
    If grilla.Text = "0.00" And grilla.TextMatrix(grilla.Row, 0) = "" Then
        grilla.Text = ""
    End If
    
    xbonifporcent = grilla.TextMatrix(grilla.Row, 9)
    xbonifimporte = grilla.TextMatrix(grilla.Row, 8)
    
    
    If xcol = 8 Or Val(Text11.Text) <> 0 Then
        grilla.Col = 8
        xbonifimporte = grilla.Text
        grilla.Text = Format(xbonifimporte, "###,##0.00")
        If grilla.Text = "0.00" And grilla.TextMatrix(grilla.Row, 0) = "" Then
            grilla.Text = ""
        End If
        grilla.Col = 9
  '      xbonifporcent = Round((xbonifimporte / (xcant * ximporte)) * 100, 2)
  '      xbonifporcent = Round((xbonifimporte / (xcant * (ximporte / xiva))) * 100, 2)
        xbonifporcent = Round((xbonifimporte / (xcant * (ximportelista / xiva))) * 100, 5)
        grilla.Text = Format(xbonifporcent, "###,##0.00")
        If grilla.Text = "0.00" And grilla.TextMatrix(grilla.Row, 0) = "" Then
          grilla.Text = ""
        End If

    End If

    

    If xcol = 9 Or xcol = 3 Or xcol = 5 Or xcol = 6 Or Val(Text10.Text) <> 0 Then
        grilla.Col = 9
        
        xbonifporcent = grilla.Text
        If xbonifporcent = "" Then xbonifporcent = 0

        
        If xlimitebonif <> 100 Then
            If xbonifporcent > xlimitebonif Then
                MsgBox "No puede ingresar un porcentaje de bonificacion mayor a al " + Str(xlimitebonif) + " %"
                xbonifporcent = xlimitebonif
            End If
        End If
        
        grilla.Text = Format(xbonifporcent, "###,##0.00")
        If grilla.Text = "0.00" And grilla.TextMatrix(grilla.Row, 0) = "" Then
            grilla.Text = ""
        End If
        grilla.Col = 8
'        xbonifimporte = Round(xbonifporcent * xcant * ximporte / 100)
        xbonifimporte = Round(xbonifporcent * xcant * (ximporte / xiva) / 100, 5)
        grilla.Text = Format(xbonifimporte, "###,##0.00")
        If grilla.Text = "0.00" And grilla.TextMatrix(grilla.Row, 0) = "" Then
            grilla.Text = ""
        End If

    End If
    
    grilla.Col = 10
    If xbonifimporte = "" Then xbonifimporte = 0
    If ximporte = "" Then ximporte = 0
    
    If grilla.TextMatrix(grilla.Row, 12) = "" Then
        grilla.TextMatrix(grilla.Row, 12) = 1.21
    End If
    
'    xtotal = Round(Round(ximportesiva * grilla.TextMatrix(grilla.Row, 12), 3) - Round(xbonifimporte, 10), 2)
    xtotal = Round((ximportesiva - xbonifimporte) * grilla.TextMatrix(grilla.Row, 12), 5) ' corregido 04/01/2016
    
    
 '   xtotal = Round(xcant * Round(ximporte, 10) - Round(xbonifimporte, 10), 2)
    grilla.Text = Format(xtotal, "###,##0.00")
    If grilla.Text = "0.00" And grilla.TextMatrix(grilla.Row, 0) = "" Then
          grilla.Text = ""
    End If


    grilla.Col = 11
    grilla.Text = xcant
        If grilla.Text = "0" And grilla.TextMatrix(grilla.Row, 0) = "" Then
            grilla.Text = ""
        End If


    xtotalgral = 0
    xsubtotalgral = 0
    xiva10 = 0
    xiva21 = 0
    xcuentaitems = 0
    For X = 1 To xlineasmax
            If grilla.TextMatrix(X, 10) = "" Then
                xgrilla = 0
                xsubtotal = 0
            Else
                xcuentaitems = xcuentaitems + 1
                xgrilla = grilla.TextMatrix(X, 10)
                If grilla.TextMatrix(X, 11) = "" Then
                 '   xsubtotal = grilla.TextMatrix(X, 10) / 1.21
                     xsubtotal = grilla.TextMatrix(X, 7) - grilla.TextMatrix(X, 8) ' corregido 04/01/2016
'                    xsubtotal = grilla.TextMatrix(X, 7)
                Else
'                    xsubtotal = grilla.TextMatrix(X, 10) / grilla.TextMatrix(X, 12)
                     xsubtotal = grilla.TextMatrix(X, 7) - grilla.TextMatrix(X, 8) ' corregido 2018
'                     xsubtotal = grilla.TextMatrix(X, 7)
                End If
            End If
            
            If grilla.TextMatrix(X, 12) = "1.105" Then
                xiva10 = xiva10 + (xgrilla - xsubtotal)
            Else
                xiva21 = xiva21 + (xgrilla - xsubtotal)
            End If
            
            xtotalgral = xtotalgral + xgrilla
            xsubtotalgral = xsubtotalgral + xsubtotal
            
    Next X
    
'--- Calculo de Tem
    xcalculotem = 0
    If UCase(xcalculatempyp) = "S" And xciudadtem = xciudadcliente And DataGrid2.Columns(16).Text <> "CF" And DataGrid2.Columns(16).Text <> "EX" And tipofac <> "NN" And xcontroltem <> 0 Then
        xcalculotem = xsubtotalgral * xalicuptatempip / 100
    End If
'--- Fin Calculo Tem

'--- Calculo de IIBB
    xcalculoIIBB = 0
    If UCase(xcalculaiibb) = "S" And DataGrid2.Columns(16).Text <> "CF" And DataGrid2.Columns(16).Text <> "EX" And tipofac <> "NN" Then
        xcalculoIIBB = (xsubtotalgral * xalicuotaiibb / 100) * ((100 - xexentoiibb) / 100)
'        If xcalculoIIBB <= 50 Then xcalculoIIBB = 0  ' Limite inferior para calculo de iibb
    End If
'--- Fin Calculo IIBB
    xtotalgral = xtotalgral + xcalculotem + xcalculoIIBB
    xsubtotal2 = xsubtotalgral + xiva10 + xiva21
    
    Text5.Text = Format(Round(xsubtotalgral, 2), "$ ###,##0.00")
    Text6.Text = Format(Round(xiva10, 3), "$ ###,##0.00")
    Text7.Text = Format(Round(xiva21, 3), "$ ###,##0.00")
    Text8.Text = Format(Round(xcalculotem, 2), "$ ###,##0.00")
    Text9.Text = Format(Round(xcalculoIIBB, 2), "$ ###,##0.00")
    Text16.Text = Format(Round(xsubtotal2, 2), "$ ###,##0.00")
    
    Text4.Text = Format(Round(xtotalgral, 2), "$ ###,##0.00")
    Text22.Text = xcuentaitems
    
'---- Limite de credito
   If xcreditomaximo <> 0 Then
    xlimitecredito = xdisponible - xtotalgral
    If xlimitecredito < 0 And UCase(DataCombo3.Text) <> "01- CONTADO" Then
        negro.Visible = True
        blanco.Visible = False
    Else
        negro.Visible = False
        blanco.Visible = True
    End If
   End If

    grilla.Col = xcol



End Sub

Private Sub Cancelar_Click()
On Error Resume Next
    mensa = MsgBox("Desea Cancelar la Operación", vbYesNo, "Atención !!")
    If mensa = vbYes Then
            Unload Me
            frmnota_venta.Show
    End If

End Sub

Private Sub DataList1_KeyPress(KeyAscii As Integer)
On Error GoTo fuera

    If KeyAscii = 13 Then
        KeyAscii = 0
        Text1(5).Text = DataList1.BoundText
        Text1(6).SetFocus
    End If

fuera:

End Sub

Private Sub DataList1_LostFocus()
On Error Resume Next

    DataList1.Visible = False

End Sub

Private Sub cancelar2_Click()
On Error Resume Next
            
            Unload Me
            frmnota_venta.Show
            
End Sub

Private Sub cargapresupuesto_Click()
On Error Resume Next

If menu <> 6 Then
    datencabezado.RecordSource = "SELECT * From ud_ezi_puntodeventa_encabezado where  numeradorinterno = 'Presupuesto de Venta' and  id ='" & Text18.Text & "' and generada = 'False'"
    datencabezado.Refresh
Else
    Frame2.Caption = "Nota de Venta"
    grabar.Caption = "&Modificar"
    xmodifica = 1
    datencabezado.RecordSource = "SELECT * From ud_ezi_puntodeventa_encabezado where  numeradorinterno = 'Nota de Venta' and  claveprimaria ='" & Text17.Text & "'"
    datencabezado.Refresh
End If
    
    If datencabezado.Recordset.EOF = False Then
        xidpre = datencabezado.Recordset.Fields("id")
    '**** Carga cliente
        xquericliente = "SELECT  ALIAS_0.ID AS ID,  ALIAS_0.CODIGO, ALIAS_3.NOMBRE AS RAZONSOCIAL, ALIAS_3.CUIT,ALIAS_6.CALLE, ISNULL(ALIAS_6.CALLE,'') + '-'+ ISNULL(V_CIUDAD_.NOMBRE,'') + '-'+ ISNULL(V_PROVINCIA_.NOMBRE,'') AS DOMICILIO, ALIAS_0.DENOMINACION,ALIAS_5.NOMBRE AS ZONA, " & _
              "ALIAS_7.NUMERO AS TELEFONO, ALIAS_8.DIRECCIONELECTRONICA AS MAIL, " & _
              "V_TIPOPAGO_.NOMBRE AS TP, V_TIPOPAGO_.OBSERVACION as TipoPago, V_TIPOPAGO_.ACTIVESTATUS AS ACTIVESTATUSTP, V_PROVINCIA_.NOMBRE AS Provincia, " & _
              "V_CIUDAD_.NOMBRE AS Ciudad, ALIAS_0.CODIGO + ' ' + ALIAS_3.CUIT + ' ' + ALIAS_3.NOMBRE AS CONCATENADO, V_EZI_POCICION_IVA_CLIENTES.CODIGO AS IVA, V_TIPOPAGO_.ID AS IDPAGO, ALIAS_0.DOMICILIOFACTURACION_ID as domicilio_id, ALIAS_0.creditomaximo    " & _
              "FROM V_PERSONA_ AS ALIAS_3 RIGHT OUTER JOIN V_CLIENTE AS ALIAS_0 WITH (READPAST) LEFT OUTER JOIN " & _
              "V_EZI_POCICION_IVA_CLIENTES ON ALIAS_0.ID = V_EZI_POCICION_IVA_CLIENTES.idcliente LEFT OUTER JOIN " & _
              "V_TIPOPAGO_ ON ALIAS_0.TIPOPAGO_ID = V_TIPOPAGO_.ID ON ALIAS_3.ID = ALIAS_0.ENTEASOCIADO_ID LEFT OUTER JOIN " & _
              "V_CIUDAD_ RIGHT OUTER JOIN V_DOMICILIO_ AS ALIAS_6 ON V_CIUDAD_.ID = ALIAS_6.CIUDAD_ID LEFT OUTER JOIN " & _
              "V_PROVINCIA_ ON ALIAS_6.PROVINCIA_ID = V_PROVINCIA_.ID ON ALIAS_0.DOMICILIOFACTURACION_ID = ALIAS_6.ID LEFT OUTER JOIN " & _
              "V_TELEFONO_ AS ALIAS_7 ON ALIAS_3.TELEFONOPRINCIPAL_ID = ALIAS_7.ID LEFT OUTER JOIN " & _
              "V_DIRECCIONELECTRONICA_ AS ALIAS_8 ON ALIAS_3.DIRECELECTRONICAPRINCIPAL_ID = ALIAS_8.ID LEFT OUTER JOIN " & _
              "V_ZONA_ AS ALIAS_5 ON ALIAS_0.ZONA_ID = ALIAS_5.ID " & _
              "WHERE     (ALIAS_0.ACTIVESTATUS <> 2) and ALIAS_0.ID = '" & datencabezado.Recordset.Fields("clienteid") & "'"
              
        datcliente.RecordSource = xquericliente
        datcliente.Refresh
        
        Text1(1).Text = DataGrid2.Columns(2).Text
        If DataGrid2.Columns(16).Text = "RI" Then
            Text1(2).Text = "A"
        Else
            Text1(2).Text = "B"
        End If
        If DataGrid2.Columns(17).Text <> "" Then
           If menu <> 6 Then
            DataCombo3.BoundText = DataGrid2.Columns(17).Text
           Else
            DataCombo3.BoundText = datencabezado.Recordset.Fields("tipodepagoid")
           End If
           
            testpos = InStr(1, DataCombo3.Text, "- ", 1)
            tpago = Right(DataCombo3.Text, Len(DataCombo3.Text) - testpos - 1)
        Else
            If menu <> 6 Then
                tpago = "CONTADO"
                DataCombo3.Text = "01- CONTADO"
            Else
                DataCombo3.BoundText = datencabezado.Recordset.Fields("tipodepagoid")
            End If
        End If
        Text1(4).Text = DataGrid2.Columns(3).Text
        Text1(3).Text = DataGrid2.Columns(5).Text
        If tpago = "CONTADO" Then
            'DataCombo3.Enabled = False
            DataCombo3.Enabled = True
        Else
            DataCombo3.Enabled = True
        End If
        Text1(5).SetFocus
        
        xcontroltem = 1
        datiibb.RecordSource = "SELECT * from v_ezi_pos_impuestos " & _
                               "WHERE idcliente = '" & DataGrid2.Columns(0).Text & "'"
        datiibb.Refresh
        If datiibb.Recordset.EOF = False Then
            xalicuotaiibb = datiibb.Recordset.Fields("COEFICIENTE")
            xexentoiibb = datiibb.Recordset.Fields("exencion")
            If xexentoiibb = 0 Then
                Label1(14).Caption = "Percep IIBB " + Str(xalicuotaiibb) + " %"
            Else
                Label1(14).Caption = "Percep IIBB " + Str(xalicuotaiibb) + " % Cert.Ex"
            End If
            If DataGrid2.Columns(16).Text = "CF" Then Label1(14).Caption = "Percep IIBB "
            xcontroltem = datiibb.Recordset.Fields("tem")
        End If
        
        xciudadcliente = DataGrid2.Columns("Ciudad").Text
        Call limitecredito_Click
'*** fin Carga Clinte
'*** Carga Vendedor
        xqueryvende = "SELECT    V_VENDEDOR_.ID, V_VENDEDOR_.CODIGO, V_PERSONA_.NOMBRE, V_VENDEDOR_.CODIGO  + ' ' + V_PERSONA_.NOMBRE as cancatena, ud_ezi_empleado.limite  " & _
              "FROM V_VENDEDOR_ INNER JOIN V_PERSONA_ ON V_VENDEDOR_.ENTEASOCIADO_ID = V_PERSONA_.ID LEFT OUTER JOIN ud_ezi_empleado WITH (nolock) ON V_VENDEDOR_.ID = ud_ezi_empleado.id " & _
              "WHERE (V_VENDEDOR_.ACTIVESTATUS = 0) and V_VENDEDOR_.ID = '" & datencabezado.Recordset.Fields("vendedorid") & "'"
        
        datvendedor.RecordSource = xqueryvende
        datvendedor.Refresh
        Text1(0).Text = DataGrid1.Columns(2).Text
        Text1(1).SetFocus
        xlimitebonif = DataGrid1.Columns(4).Text
'*** Fin carga Vendedor
        Text1(5).Text = datencabezado.Recordset.Fields("detalle")
        Text15.Text = datencabezado.Recordset.Fields("nota")
        Text1(9).Text = datencabezado.Recordset.Fields("adicionalid")
        If Text1(9).Text <> "" Then
            Text1(9).Locked = False
        Else
            Text1(9).Locked = False
        End If

       If menu <> 6 Then
        datitems.RecordSource = "Select * from ud_ezi_puntodeventa_detalle_presu where claveprimaria = " & xidpre & " order by id"
        datitems.Refresh
       Else
        datitems.RecordSource = "Select * from ud_ezi_puntodeventa_detalle_notav where claveprimaria = " & xidpre & " order by id"
        datitems.Refresh
       End If
    
        datitems.Recordset.MoveFirst
        For X = 1 To datitems.Recordset.RecordCount
            grilla.TextMatrix(X, 0) = datitems.Recordset.Fields("idproducto")
            grilla.TextMatrix(X, 1) = datitems.Recordset.Fields("codigoproducto")
            grilla.TextMatrix(X, 2) = datitems.Recordset.Fields("nombre_producto")
            grilla.TextMatrix(X, 3) = datitems.Recordset.Fields("cantidadproducto")
            grilla.TextMatrix(X, 4) = datitems.Recordset.Fields("unidaddemedidaid")
'            grilla.TextMatrix(X, 5) = Round(datitems.Recordset.Fields("preciou") / ((datitems.Recordset.Fields("iva") + 100) / 100), 5)
            grilla.TextMatrix(X, 5) = Round(((datitems.Recordset.Fields("subtotal") / ((datitems.Recordset.Fields("iva") + 100) / 100)) / datitems.Recordset.Fields("cantidadproducto")), 5)

'            grilla.TextMatrix(X, 6) = Round(datitems.Recordset.Fields("preciou") * ((datitems.Recordset.Fields("bonificacionitem") + 100) / 100), 5)
            grilla.TextMatrix(X, 6) = Round((datitems.Recordset.Fields("subtotal") / datitems.Recordset.Fields("cantidadproducto")), 5)

            
            
            grilla.TextMatrix(X, 12) = (datitems.Recordset.Fields("iva") / 100) + 1
'            grilla.TextMatrix(X, 7) = Round((datitems.Recordset.Fields("preciou") / ((datitems.Recordset.Fields("iva") + 100) / 100)) * datitems.Recordset.Fields("cantidadproducto"), 5)

            grilla.TextMatrix(X, 7) = Round((datitems.Recordset.Fields("subtotal") / ((datitems.Recordset.Fields("iva") + 100) / 100)), 5)
            
'            grilla.TextMatrix(X, 9) = Round(datitems.Recordset.Fields("bonificacionitem"), 10)
            grilla.TextMatrix(X, 10) = Round(datitems.Recordset.Fields("subtotal"), 10)
            grilla.TextMatrix(X, 11) = datitems.Recordset.Fields("cantidadproducto")
            grilla.TextMatrix(X, 14) = datitems.Recordset.Fields("preciou")
            grilla.TextMatrix(X, 17) = datitems.Recordset.Fields("preciou")
            grilla.Col = 3
            grilla.Row = X
            Call calcula_Click
            datitems.Recordset.MoveNext
            
            
        Next X
    
 '       If datencabezado.Recordset.Fields("bonificacion") <> 0 Then
 '           Text11.SetFocus
 '           Text11.Text = Format(datencabezado.Recordset.Fields("bonificacion"), "###,##0.00")
 '           SendKeys "{ENTER}", False
 '       End If
        
 '       If datencabezado.Recordset.Fields("recargo") <> 0 Then
 '           Text13.SetFocus
 '           Text13.Text = Format(datencabezado.Recordset.Fields("recargo"), "###,##0.00")
 '           SendKeys "{ENTER}", False
 '       End If
    
    
    Else
        mensa = MsgBox("No Existe el Nro de cotizacion seleccionado", vbInformation, "!! Sin Coincidencias !!")
    End If
    

    

End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
On Error Resume Next

If KeyAscii = 13 Then
        KeyAscii = 0
        
        If Combo1.Text <> xumvta And grilla.Text <> Combo1.Text Then
            grilla.Text = Combo1.Text
            grilla.Col = 6
            grilla.Text = Round(grilla.Text / xfaconv(Combo1.ListIndex), 2)
            grilla.TextMatrix(grilla.Row, 17) = grilla.Text
            grilla.TextMatrix(grilla.Row, 13) = xfaconv(Combo1.ListIndex)
            Call calcula_Click
            Call ubicatextogrilla_Click
            Exit Sub
        End If
        If Combo1.Text = xumvta And grilla.Text <> Combo1.Text Then
            grilla.Text = Combo1.Text
            grilla.Col = 6
            grilla.TextMatrix(grilla.Row, 17) = grilla.Text
            grilla.Text = Round(grilla.Text * grilla.TextMatrix(grilla.Row, 13), 2)
            Call calcula_Click
            Call ubicatextogrilla_Click
            Exit Sub
        End If
        grilla.Col = 5
        Call ubicatextogrilla_Click

End If


End Sub

Private Sub Combo1_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next

If KeyCode = 39 Then
    grilla.Col = 5
    Call ubicatextogrilla_Click
End If

If KeyCode = 37 Then
    grilla.Col = 3
    Call ubicatextogrilla_Click
End If

If KeyCode = 121 Then
       KeyCode = 0
       Call grabar_Click
End If


If KeyCode = 116 Then
       Call agregaproducto_Click
End If


End Sub

Private Sub Command7_Click()
On Error Resume Next
    menu = 1
      query = "SELECT     R.claveprimaria, R.clienteid, R.cliente, R.fechadelcomprobante, RD.referenciaproducto, RD.nombre_producto, RD.cantidadoriginal, RD.unidaddemedida, R.presupuestobase , NVD.preciou " & _
              "FROM         ud_ezi_puntodeventa_encabezado AS R WITH (nolock) INNER JOIN  " & _
              "ud_ezi_puntodeventa_detalle_rem AS RD WITH (nolock) ON R.claveprimaria = RD.claveprimaria INNER JOIN " & _
              "ud_ezi_puntodeventa_encabezado AS NV WITH (nolock) ON R.presupuestobase = NV.claveprimaria INNER JOIN " & _
              "ud_ezi_puntodeventa_detalle_notav AS NVD WITH (nolock) ON RD.idproducto = NVD.idproducto AND RD.item = NVD.item AND NV.id = NVD.claveprimaria " & _
              "WHERE     (R.numeradorinterno = 'Remito de Venta') and R.clienteid =  '" & DataGrid2.Columns("id").Text & "'" & _
              "ORDER BY R.fechadelcomprobante DESC  "
    lista_historial.Show
    Text2.SetFocus

    



End Sub

Private Sub Command8_Click()
On Error Resume Next
    
If DataCombo3.Text = "01- CONTADO" Then
    menu = 1
    xfila = grilla.Row
      query = "SELECT  * from  v_ezi_pos_stock_lotes " & _
            "where REFERENCIATIPO_ID = '" & grilla.TextMatrix(grilla.Row, 0) & "' " & _
            "ORDER BY FECHAVENCIMIENTO, CODIGO"
    lista_lotes.Show
'    lista_lotes.DataGrid1.SetFocus
End If

End Sub

Private Sub Command9_Click()

grilla.SetFocus
Frame4.Visible = False


End Sub

Private Sub DataCombo1_KeyPress(KeyAscii As Integer)
On Error Resume Next

    If KeyAscii = 13 Then
        KeyAscii = 0
        DataCombo2.SetFocus
    End If

End Sub

Private Sub DataList2_KeyPress(KeyAscii As Integer)
On Error GoTo fuera

    If KeyAscii = 13 Then
        KeyAscii = 0
        Text1(6).Text = DataList2.BoundText
        grabar.SetFocus
    End If

fuera:
End Sub

Private Sub DataList2_LostFocus()
On Error Resume Next

    DataList2.Visible = False

End Sub

Private Sub DataList3_KeyPress(KeyAscii As Integer)
On Error Resume Next

    If KeyAscii = 13 Then
        KeyAscii = 0
        Text1(4).Text = DataList3.Text
        Text1(5).SetFocus
        DataList3.Visible = False
    End If


End Sub

Private Sub DataCombo2_KeyPress(KeyAscii As Integer)

On Error Resume Next

    If KeyAscii = 13 Then
        KeyAscii = 0
        DataCombo3.SetFocus
    End If


End Sub

Private Sub DataCombo3_Click(Area As Integer)
On Error Resume Next
testpos = InStr(1, DataCombo3.Text, "- ", 1)
tpago = Right(DataCombo3.Text, Len(DataCombo3.Text) - testpos - 1)

End Sub

Private Sub DataCombo3_KeyPress(KeyAscii As Integer)

On Error Resume Next

    If KeyAscii = 13 Then
        KeyAscii = 0
        testpos = InStr(1, DataCombo3.Text, "- ", 1)
        tpago = Right(DataCombo3.Text, Len(DataCombo3.Text) - testpos - 1)
        Text1(5).SetFocus
    End If

End Sub

Private Sub DataCombo3_LostFocus()
On Error Resume Next
    
    Call limitecredito_Click
    
    If DataCombo3.Text = "01- CONTADO" Then Call verificalotes_Click
    

End Sub

Private Sub Form_Load()
On Error Resume Next

yventana = Inicio.Height / 2 - 1000
xventana = Inicio.Width / 2

frmnota_venta.Top = yventana - frmnota_venta.Height / 2
frmnota_venta.Left = xventana - frmnota_venta.Width / 2

If login.nomsucursal = "EMPORIOZIP" Or login.nomsucursal = "TUCUMANZIP" Then
'    Aplicar_skin2 Me
    Aplicar_skin Me
Else
    Aplicar_skin Me
End If

xmodifica = 0

xvendedorautoriza = ""
datvendedor.ConnectionString = login.conexiontotal
datcliente.ConnectionString = login.conexiontotal
datproductos.ConnectionString = login.conexiontotal
datmovimientos.ConnectionString = login.conexiontotal
dattipopago.ConnectionString = login.conexiontotal
datparametros.ConnectionString = login.conexiontotal
datlistaprecios.ConnectionString = login.conexiontotal
datum.ConnectionString = login.conexiontotal
datiibb.ConnectionString = login.conexiontotal
datencabezado.ConnectionString = login.conexiontotal
datitems.ConnectionString = login.conexiontotal
datitemsnv.ConnectionString = login.conexiontotal
datitempresup.ConnectionString = login.conexiontotal
datcontrol.ConnectionString = login.conexiontotal
datcola.ConnectionString = login.conexiontotal
datcolaimportar.ConnectionString = login.conexiontotal
datpago.ConnectionString = login.conexiontotal
datcredito.ConnectionString = login.conexiontotal

negro.Visible = False
blanco.Visible = False

If tipodeventa = 0 Then
    tipofac = "CF"
Else
    tipofac = "NN"
End If

Text18.Text = ""
Text17.Text = ""

DFECHA.Value = Date
Text19.Text = ""


    dattipopago.RecordSource = "SELECT ID, NOMBRE AS CODIGO, nombre +'- '+ OBSERVACION AS TipoPago From V_TIPOPAGO_ WHERE (ACTIVESTATUS = 0) order by NOMBRE"
    dattipopago.Refresh
    
    datvendedor.RecordSource = "SELECT     V_VENDEDOR_.CODIGO, V_PERSONA_.NOMBRE FROM V_VENDEDOR_ INNER JOIN " & _
                               "V_PERSONA_ ON V_VENDEDOR_.ENTEASOCIADO_ID = V_PERSONA_.ID Where (V_VENDEDOR_.ACTIVESTATUS = 0)"
    datvendedor.Refresh

    datparametros.RecordSource = "select * from ud_ezi_parametros_pos where sucursal = '" & login.nomsucursal & "' "
    datparametros.Refresh
    
    datlistaprecios.RecordSource = "select id, nombre from v_listaprecio_"
    datlistaprecios.Refresh
    
    DataCombo1.BoundText = datparametros.Recordset.Fields("listaprecio")
    DataCombo1.Enabled = False
    
    xcalculaiibb = datparametros.Recordset.Fields("calculaiibb")
    xcalculatempyp = datparametros.Recordset.Fields("calculatempyp")
    xalicuptatempip = datparametros.Recordset.Fields("alicuotatempip")
    xciudadtem = datparametros.Recordset.Fields("ciudadtem")
    
    
grilla.Row = 0
grilla.Col = 0
grilla.ColWidth(0) = 100
grilla.Col = 1
grilla.Text = "Codigo"
grilla.ColWidth(1) = 2000
grilla.Col = 2
grilla.Text = "Descipcion"
grilla.ColWidth(2) = 5000
grilla.Col = 3
grilla.Text = "Cant."
grilla.ColWidth(3) = 800
grilla.Col = 4
grilla.Text = "U.M."
grilla.ColWidth(4) = 1000

grilla.Col = 5
grilla.Text = "$ Unit.S/Iva."
grilla.ColWidth(5) = 1200

grilla.Col = 6
grilla.Text = "$ Unit.C/Iva."
grilla.ColWidth(6) = 1200
grilla.Col = 7
grilla.Text = "$Total S/Iva."
grilla.ColWidth(7) = 1200
grilla.Col = 8
grilla.Text = "$ Bonif."
'grilla.ColWidth(7) = 1000
grilla.ColWidth(8) = 0
grilla.Col = 9
grilla.Text = "% Bonif."
grilla.ColWidth(9) = 1200
grilla.Col = 10
grilla.Text = "$ Total"
grilla.ColWidth(10) = 1200
grilla.Col = 11
grilla.Text = "Cant.Ent"
grilla.ColWidth(11) = 0
grilla.Col = 12
grilla.Text = "Iva"
grilla.ColWidth(12) = 0
grilla.Col = 13
grilla.Text = ".."
grilla.ColWidth(13) = 0
grilla.Col = 15
grilla.Text = "Remito"
grilla.ColWidth(15) = 0
grilla.Col = 16
grilla.Text = "IdItemRemito"
grilla.ColWidth(16) = 800
grilla.Col = 17
grilla.ColWidth(17) = 0

grilla.ColWidth(18) = 0 ' Lote1
grilla.ColWidth(19) = 0 'cant lote1

grilla.ColWidth(20) = 0 'lote2
grilla.ColWidth(21) = 0 'cant lote2

grilla.ColWidth(22) = 0 'lote3
grilla.ColWidth(23) = 0 'cant lote3

grilla.ColWidth(24) = 0 'lote4
grilla.ColWidth(25) = 0 'cant lote4

grilla.ColWidth(26) = 0 'lote4
grilla.ColWidth(27) = 0 'cant lote4

grilla.ColWidth(28) = 0 'lote5
grilla.ColWidth(29) = 0 'cant lote5



xlineasmax = datparametros.Recordset.Fields("limiteitemsnotaventa")
grilla.Rows = xlineasmax + 1

For X = 2 To xlineasmax Step 2
  For Y = 1 To 11
    grilla.Col = Y
    grilla.Row = X
    grilla.CellBackColor = RGB(231, 235, 218)
  Next Y
Next X




   
End Sub

Private Sub grabar_Click()
On Error Resume Next

  
    If Text1(7).Text <> "" Then
        mensa = MsgBox("Atencion esta ingresando una factura de talonario Manual, este comprobante no sera impreso", vbInformation, "Atención !!")
    End If
    If Text1(8).Text <> "" Then
        mensa = MsgBox("Atencion esta ingresando un Remito de talonario Manual, este comprobante no sera impreso", vbInformation, "Atención !!")
    End If



    If UCase(datparametros.Recordset.Fields("cobroautomatico")) = "S" And tpago = "CONTADO" Then
        grabar.Enabled = False
        frmcobranzacdo.Show
        frmcobranzacdo.Text1(0).Text = Format(Text4.Text, "#,###,##0.00")
        Exit Sub
    End If
    
 If xlimitecredito < 0 Then
    mensa = MsgBox("El Limite de Credito será superado con esta Factura, no Podrá grabar Esta Nota de venta", vbInformation, "LIMITE DE CRÉDITO SUPERADO")
    Exit Sub
 End If
    
If xcontrol = 1 Then
    mensa = MsgBox("El Cliente tiene Facturas Vencidas, no Podrá grabar Esta Nota de venta", vbInformation, "Control Limite de Crédito")
    Exit Sub
 End If
    
 mensa = MsgBox("Desea Grabar esta Venta ?", vbYesNo, "!! Atención !!")
 If mensa = vbNo Then Exit Sub
    
    For j = 1 To xlineasmax
        If grilla.TextMatrix(j, 0) <> "" Then
            xcontrolprecio = Round(grilla.TextMatrix(j, 6), 4)
            xcontrolcantidad = Round(grilla.TextMatrix(j, 3), 4)
'            If xcontrolprecio = 0 Then
'                MsgBox "El Comprobante no se puede emitir porque un tiem no tiene precio", vbCritical, "Error"
'                Exit Sub
'            End If
            If xcontrolcantidad = 0 Then
                MsgBox "El Comprobante no se puede emitir porque un tiem no tiene cantidad", vbCritical, "Error"
                Exit Sub
            End If
        End If
    Next j
    
If xmodifica = 1 Then
    datencabezado.RecordSource = "SELECT * From ud_ezi_puntodeventa_encabezado where  numeradorinterno = 'Nota de Venta' and  claveprimaria ='" & Text17.Text & "'"
    datencabezado.Refresh
    
    datitems.RecordSource = "Select * from ud_ezi_puntodeventa_detalle_notav where claveprimaria = " & datencabezado.Recordset.Fields("id") & " "
    datitems.Refresh
    If datitems.Recordset.EOF = False Then
        datitems.Recordset.MoveFirst
        
        Do While Not datitems.Recordset.EOF
            datitems.Recordset.Delete adAffectCurrent
            datitems.Recordset.MoveNext
        Loop
    End If
Else
    datencabezado.RecordSource = "SELECT MAX(CONVERT(decimal, isnull(claveprimaria,0))) AS claveprimaria FROM ud_ezi_puntodeventa_encabezado with(readpast) where ud_ezi_puntodeventa_encabezado.numeradorinterno = 'Nota de Venta'  "
    datencabezado.Refresh
    
    If datencabezado.Recordset.EOF = True Then
        xclaveprimaria = 1
    Else
        xclaveprimaria = datencabezado.Recordset.Fields("claveprimaria") + 1
    End If
    If IsNull(xclaveprimaria) = True Then xclaveprimaria = 1
    

    datencabezado.RecordSource = "SELECT * From ud_ezi_puntodeventa_encabezado where id =0 "
    datencabezado.Refresh
    datitems.RecordSource = "Select * from ud_ezi_puntodeventa_detalle_notav where id = 0"
    datitems.Refresh
    
    datencabezado.Recordset.AddNew
    datencabezado.Recordset.Fields("numeradorinterno") = "Nota de Venta"
 End If ' Modifica
    datencabezado.Recordset.Fields("fechadelcomprobante") = DateValue(DFECHA.Value) + TimeValue(Str(Time))
    datencabezado.Recordset.Fields("sucursal") = datparametros.Recordset.Fields("sucursal")
    datencabezado.Recordset.Fields("clienteid") = DataGrid2.Columns(0).Text
    datencabezado.Recordset.Fields("cliente") = DataGrid2.Columns(2).Text
    If Text1(6).Text <> "" Then
        datencabezado.Recordset.Fields("cliente") = Text1(6).Text
        datencabezado.Recordset.Fields("recetaid") = Text1(3).Text
    End If
    datencabezado.Recordset.Fields("vendedorid") = DataGrid1.Columns(0).Text
    datencabezado.Recordset.Fields("vendedor") = DataGrid1.Columns(2).Text
    datencabezado.Recordset.Fields("detalle") = Text1(5).Text
    datencabezado.Recordset.Fields("nota") = Text15.Text
    datencabezado.Recordset.Fields("cotizacion") = 1
    datencabezado.Recordset.Fields("listadeprecioid") = DataCombo1.BoundText
    datencabezado.Recordset.Fields("tipodepagoid") = DataCombo3.BoundText
    datencabezado.Recordset.Fields("tipodefactura") = Text1(2).Text
    datencabezado.Recordset.Fields("alquiler") = "N"
    datencabezado.Recordset.Fields("nroorden") = Text1(4).Text
    datencabezado.Recordset.Fields("adicionalid") = Text1(9).Text
    If xvendedorautoriza <> "" Then
        datencabezado.Recordset.Fields("recetaid") = xvendedorautoriza
    End If
    If Text1(7).Text <> "" Then
        datencabezado.Recordset.Fields("retira") = Text1(7).Text
    End If
    
    If tipofac <> "NN" Then
        datencabezado.Recordset.Fields("tipodefacturacionid") = datparametros.Recordset.Fields("facdefecto")
    Else
        datencabezado.Recordset.Fields("tipodefacturacionid") = tipofac
    End If
    
    If Left(login.nombrebd, 14) = "MMOSSE" And tipofac <> "NN" Then
        If tpago = "CONTADO" Then
            datencabezado.Recordset.Fields("tipodefacturacionid") = datparametros.Recordset.Fields("facdefecto")
        Else
            datencabezado.Recordset.Fields("tipodefacturacionid") = datparametros.Recordset.Fields("facdefectocc")
        End If
    End If
If xmodifica = 0 Then
    datencabezado.Recordset.Fields("claveprimaria") = xclaveprimaria
Else
    xclaveprimaria = datencabezado.Recordset.Fields("claveprimaria")
End If
    datencabezado.Recordset.Fields("fechadeentrega") = DateValue(DFECHA.Value) + TimeValue(Str(Time))
    If Text13.Text = "" Then Text13.Text = 0
    datencabezado.Recordset.Fields("recargo") = Round(Text13.Text, 2)
    datencabezado.Recordset.Fields("tiporecargo") = "$"
    If Text11.Text = "" Then Text11.Text = 0
    datencabezado.Recordset.Fields("bonificacion") = Round(Text11.Text, 2)
    datencabezado.Recordset.Fields("tipobonificacion") = "$"
    datencabezado.Recordset.Fields("importeglobal") = Round(Text4.Text, 2)
    datencabezado.Recordset.Fields("numerodefactura") = xclaveprimaria
    datencabezado.Recordset.Fields("domicilioid") = Text1(3).Text
    If DataGrid2.Columns("domicilio_id").Text = "" Then
       If login.nomsucursal <> "EMPORIOZIP" And login.nomsucursal <> "TUCUMANZIP" Then
        MsgBox "Debe ingresar un domicilio de Facturacion Valido en el cliente", vbCritical, "Error"
        Exit Sub
       End If
    End If
    datencabezado.Recordset.Fields("domicilio_id") = DataGrid2.Columns("domicilio_id").Text
    datencabezado.Recordset.Fields("domiciliodeentregaid") = DataGrid2.Columns("domicilio_id").Text
    datencabezado.Recordset.Fields("subtotalsiniva") = Round(Text5.Text, 2)
    datencabezado.Recordset.Fields("totaliva") = Round(Text6.Text, 2) + Round(Text7.Text, 2)
If xmodifica = 0 Then
    datencabezado.Recordset.Fields("generada") = "False"
    datencabezado.Recordset.Fields("importado") = "False"
    datencabezado.Recordset.Fields("nombrepc") = Environ("computername")
    datencabezado.Recordset.Fields("responsabilidad") = DataGrid2.Columns(16).Text
    datencabezado.Recordset.Fields("transferido") = "False"
End If

    datencabezado.Recordset.Fields("percepiibb") = Round(Text9.Text, 2)
    datencabezado.Recordset.Fields("perceptem") = Round(Text8.Text, 2)
    datencabezado.Recordset.Fields("totaltr") = Round(Text4.Text, 2)
    
If xmodifica = 0 Then
    datencabezado.Recordset.Fields("presupuestobase") = xidpre
    datencabezado.Recordset.UpdateBatch adAffectCurrent
    presupuestobase = datencabezado.Recordset.Fields("id")
    If Text17.Text <> "" Then
        datencabezado.Recordset.Fields("trazabilidad_id") = xidpre
    Else
        datencabezado.Recordset.Fields("trazabilidad_id") = datencabezado.Recordset.Fields("id")
    End If
    If xclaveprimaria = 0 Then
        datencabezado.Recordset.Fields("claveprimaria") = datencabezado.Recordset.Fields("id")
    End If
End If
    datencabezado.Recordset.Fields("FILTRAARMADO") = Null
    datencabezado.Recordset.UpdateBatch adAffectCurrent
    
    
    
'--- Graba Items
    
    For X = 1 To xlineasmax
        If grilla.TextMatrix(1, 0) = "" Then
          mensa = MsgBox("No puede Grabar esta Venta sin Items", vbCritical, "Error")
          Exit Sub
        End If
        If grilla.TextMatrix(X, 0) = "" Then Exit For
        
        datitems.Recordset.AddNew
    If xmodifica = 0 Then
        datitems.Recordset.Fields("claveprimaria") = presupuestobase
    Else
        datitems.Recordset.Fields("claveprimaria") = datencabezado.Recordset.Fields("id")
    End If
        datitems.Recordset.Fields("idproducto") = grilla.TextMatrix(X, 0)
        datitems.Recordset.Fields("codigoproducto") = grilla.TextMatrix(X, 1)
        datitems.Recordset.Fields("nombre_producto") = grilla.TextMatrix(X, 2)
        datitems.Recordset.Fields("cantidadproducto") = grilla.TextMatrix(X, 3)
        datitems.Recordset.Fields("unidaddemedidaid") = grilla.TextMatrix(X, 4)
        datitems.Recordset.Fields("preciou") = Round(grilla.TextMatrix(X, 6), 4)
        datitems.Recordset.Fields("preciousiva") = Round(grilla.TextMatrix(X, 5), 4)
        datitems.Recordset.Fields("bonificacionitem") = grilla.TextMatrix(X, 9)
        If grilla.TextMatrix(X, 8) = "" Then grilla.TextMatrix(X, 8) = 0
        datitems.Recordset.Fields("importedebonificacion") = Round(grilla.TextMatrix(X, 8), 4)
        datitems.Recordset.Fields("subtotal") = Round(grilla.TextMatrix(X, 10), 3)
        datitems.Recordset.Fields("entregar") = Round(grilla.TextMatrix(X, 11), 4)
        datitems.Recordset.Fields("iva") = (Round(grilla.TextMatrix(X, 12), 4) - 1) * 100
        datitems.Recordset.Fields("listaid") = grilla.TextMatrix(X, 18)  ' Lote
        datitems.Recordset.Fields("item") = X
        datitems.Recordset.UpdateBatch adAffectCurrent
    Next
    
    
'    frmnota_venta.Enabled = False
If xmodifica = 0 Then
    If Text18.Text <> "" Then
        datitempresup.RecordSource = "Select id, generada from ud_ezi_puntodeventa_encabezado where id = " & Val(Text18.Text) & "    "
        datitempresup.Refresh
        If datitempresup.Recordset.EOF = False Then
            datitempresup.Recordset.Fields("generada") = "True"
            datitempresup.Recordset.UpdateBatch adAffectCurrent
        End If
    End If

'******* Graba Cola importar
        
        datcolaimportar.RecordSource = "Select * from ud_ezi_cola_importar where id = 0"
        datcolaimportar.Refresh
        
        datcolaimportar.Recordset.AddNew
        datcolaimportar.Recordset.Fields("id_encabezado") = presupuestobase
        datcolaimportar.Recordset.Fields("tipodedocumentoid") = datparametros.Recordset.Fields("idnotaventa")
        datcolaimportar.Recordset.Fields("unidadoperativaid") = datparametros.Recordset.Fields("target")
        datcolaimportar.Recordset.Fields("fecha_hora") = DateValue(DFECHA.Value) + TimeValue(Str(Time))
        
        datcolaimportar.Recordset.UpdateBatch adAffectCurrent
End If

If xmodifica = 1 Then
    Call acomodaitems_Click
End If

''' Graba Remito y/o Factura
  If datparametros.Recordset.Fields("remitofacturaautomatico") = "S" Then
        
    Call grremito_Click
    
    
    If UCase(datparametros.Recordset.Fields("preguntafacturaimprimir")) = "S" And tpago <> "CONTADO" Then
'        mensa = MsgBox("Desea Imprimir esta Factura de Cta.Cte (s/n) ?", vbYesNo, "!! Atención !!")
'        If mensa = vbYes Then
            Call grfacturactacte_Click
'        End If
        Call cancelar2_Click
    End If
    
    If Text1(7).Text <> "" Then
                Call grfacturactacte_Click
                Call cancelar2_Click
    End If
    
  End If
  
  If UCase(datparametros.Recordset.Fields("cobroautomatico")) = "N" Then
    mensa = MsgBox("Nota de Venta: " + Str(xclaveprimaria), vbInformation, "Grabado Correctamente")
    Call cancelar2_Click
    Exit Sub
  Else
    mensa = MsgBox("Nota de Venta: " + Str(xclaveprimaria), vbInformation, "Grabado Correctamente")
    Call cancelar2_Click
    Exit Sub
  End If
    
    
Exit Sub
errorgrabar:
    mensa = MsgBox("No se pudo registrar la información", vbCritical, "Error !!")






End Sub

Private Sub List1_KeyPress(KeyAscii As Integer)
On Error GoTo fuera

    If KeyAscii = 13 Then
        KeyAscii = 0
        Text1(4).Text = List1.ListIndex + 1
        Text1(5).SetFocus
    End If

fuera:
End Sub

Private Sub List1_LostFocus()

    List1.Visible = False

End Sub


Private Sub manual_Click()


    For X = frmpesada_cania.Width To 13000 Step 100
            frmpesada_cania.Width = X
    Next X
    Text4.SetFocus

End Sub

Private Sub grabaregistros_Click()
'On Error GoTo errorgrabar


End Sub

Private Sub grfacturactacte_Click()
On Error GoTo errorgrabar
    
        
    For j = 1 To xlineasmax
        If grilla.TextMatrix(j, 0) <> "" Then
            xcontrolprecio = Round(grilla.TextMatrix(j, 6), 4)
'            If xcontrolprecio = 0 Then
'                MsgBox "La factura no se puede emitir porque un tiem no tiene precio", vbCritical, "Error"
'                Exit Sub
'            End If
        End If
    Next j
    
    datencabezado.RecordSource = "SELECT MAX(CONVERT(decimal, isnull(claveprimaria,0))) AS claveprimaria FROM ud_ezi_puntodeventa_encabezado with(readpast) "
    datencabezado.Refresh
    
    If datencabezado.Recordset.EOF = True Then
        xclaveprimaria = 1
    Else
        xclaveprimaria = datencabezado.Recordset.Fields("claveprimaria") + 1
    End If
    
    datencabezado.RecordSource = "SELECT * From ud_ezi_puntodeventa_encabezado with(readpast) where id =0 "
    datencabezado.Refresh
    datitems.RecordSource = "Select * from ud_ezi_puntodeventa_detalle_factm with(readpast) where id = 0"
    datitems.Refresh
    
    datencabezado.Recordset.AddNew
    datencabezado.Recordset.Fields("claveprimaria") = xclaveprimaria
    datencabezado.Recordset.Fields("numeradorinterno") = "Factura de Venta"
    datencabezado.Recordset.Fields("fechadelcomprobante") = DateValue(frmnota_venta.DFECHA.Value) + TimeValue(Str(Time))
    datencabezado.Recordset.Fields("sucursal") = frmnota_venta.datparametros.Recordset.Fields("sucursal")
    datencabezado.Recordset.Fields("clienteid") = frmnota_venta.DataGrid2.Columns(0).Text
    datencabezado.Recordset.Fields("cliente") = frmnota_venta.DataGrid2.Columns(2).Text
    If Text1(6).Text <> "" Then
        datencabezado.Recordset.Fields("cliente") = Text1(6).Text
        datencabezado.Recordset.Fields("recetaid") = Text1(3).Text
    End If
    datencabezado.Recordset.Fields("vendedorid") = frmnota_venta.DataGrid1.Columns(0).Text
    datencabezado.Recordset.Fields("vendedor") = frmnota_venta.DataGrid1.Columns(2).Text
    datencabezado.Recordset.Fields("detalle") = frmnota_venta.Text1(5).Text
    datencabezado.Recordset.Fields("nota") = frmnota_venta.Text15.Text
    datencabezado.Recordset.Fields("cotizacion") = 1
    datencabezado.Recordset.Fields("listadeprecioid") = frmnota_venta.DataCombo1.BoundText
    datencabezado.Recordset.Fields("tipodepagoid") = frmnota_venta.DataCombo3.BoundText
    datencabezado.Recordset.Fields("alquiler") = "N"
    
    If frmnota_venta.tipofac <> "NN" Then
        datencabezado.Recordset.Fields("tipodefacturacionid") = frmnota_venta.datparametros.Recordset.Fields("facdefecto")
    Else
        datencabezado.Recordset.Fields("tipodefacturacionid") = frmnota_venta.tipofac
    End If
    
    If Left(login.nombrebd, 14) = "MMOSSE" And frmnota_venta.tipofac <> "NN" Then
        If frmnota_venta.tpago = "CONTADO" Then
            datencabezado.Recordset.Fields("tipodefacturacionid") = frmnota_venta.datparametros.Recordset.Fields("facdefecto")
        Else
            datencabezado.Recordset.Fields("tipodefacturacionid") = frmnota_venta.datparametros.Recordset.Fields("facdefectocc")
        End If
    End If
    
    datencabezado.Recordset.Fields("fechadeentrega") = DateValue(frmnota_venta.DFECHA.Value) + TimeValue(Str(Time))
    If frmnota_venta.Text13.Text = "" Then frmnota_venta.Text13.Text = 0
    datencabezado.Recordset.Fields("recargo") = Round(frmnota_venta.Text13.Text, 2)
    datencabezado.Recordset.Fields("tiporecargo") = "$"
    If frmnota_venta.Text11.Text = "" Then frmnota_venta.Text11.Text = 0
    datencabezado.Recordset.Fields("bonificacion") = Round(frmnota_venta.Text11.Text, 2)
    datencabezado.Recordset.Fields("tipobonificacion") = "$"
    datencabezado.Recordset.Fields("importeglobal") = Round(frmnota_venta.Text4.Text, 2)
    datencabezado.Recordset.Fields("domicilioid") = frmnota_venta.Text1(3).Text
    datencabezado.Recordset.Fields("domicilio_id") = frmnota_venta.DataGrid2.Columns("domicilio_id").Text
    datencabezado.Recordset.Fields("domiciliodeentregaid") = frmnota_venta.DataGrid2.Columns("domicilio_id").Text
    datencabezado.Recordset.Fields("subtotalsiniva") = Round(frmnota_venta.Text5.Text, 2)
    datencabezado.Recordset.Fields("totaliva") = Round(frmnota_venta.Text6.Text, 2) + Round(frmnota_venta.Text7.Text, 2)
    datencabezado.Recordset.Fields("generada") = "True"
    datencabezado.Recordset.Fields("importado") = "False"
    datencabezado.Recordset.Fields("nombrepc") = Environ("computername")
    datencabezado.Recordset.Fields("target") = frmnota_venta.datparametros.Recordset.Fields("sucursal")
    If DataGrid2.Columns(16).Text = "RI" Then xresponsabilidad = "Resp. Inscripto"
    If DataGrid2.Columns(16).Text = "EX" Then xresponsabilidad = "Exento"
    If DataGrid2.Columns(16).Text = "MT" Then xresponsabilidad = "Monotributista"
    If DataGrid2.Columns(16).Text = "CF" Then xresponsabilidad = "Consumidor Final"
    datencabezado.Recordset.Fields("responsabilidad") = xresponsabilidad
    datencabezado.Recordset.Fields("transferido") = "False"
    datencabezado.Recordset.Fields("comprobanteorigen") = xremito  ' *** Aqui va el Id del Remito
    datencabezado.Recordset.Fields("tipodefactura") = frmnota_venta.Text1(2).Text
    datencabezado.Recordset.Fields("percepiibb") = Round(frmnota_venta.Text9.Text, 2)
    datencabezado.Recordset.Fields("perceptem") = Round(frmnota_venta.Text8.Text, 2)
    datencabezado.Recordset.Fields("totaltr") = Round(frmnota_venta.Text4.Text, 2)
    datencabezado.Recordset.Fields("presupuestobase") = presupuestobase
    nromanual = Text1(7).Text
    puntomanual = datparametros.Recordset.Fields("puntovtamanual")
    
    datencabezado.Recordset.UpdateBatch adAffectCurrent
    xid = datencabezado.Recordset.Fields("id")
    
    '** Establene numero de Facturas Manuales, y no Fiscales
    If datencabezado.Recordset.Fields("tipodefacturacionid") = "MA" Then
      If Text1(2).Text = "A" Then
            xnumerador = "Factura A (Vtas) " + datparametros.Recordset.Fields("sucursal")
      Else
            xnumerador = "Factura B (Vtas) " + datparametros.Recordset.Fields("sucursal")
      End If
    End If
    If datencabezado.Recordset.Fields("tipodefacturacionid") = "NN" Then
            xnumerador = "Factura de Venta Val " + datparametros.Recordset.Fields("sucursal")
    End If
    datcola.RecordSource = "SELECT p.ID AS ID, p.NOMBRE AS puntero, n.NUMERO AS NUMERO, p.CARACTERESPREFIJO AS puntoventa, p.NUMERO_ID " & _
                           "FROM V_NUMERADOR_ AS p LEFT OUTER JOIN V_NUMERO_ AS n ON p.NUMERO_ID = n.ID " & _
                           "WHERE     (p.ACTIVESTATUS <> 2) AND (p.NOMBRE = '" & xnumerador & "') "
    datcola.Refresh
    If datencabezado.Recordset.Fields("tipodefacturacionid") = "CF" Then
        datencabezado.Recordset.Fields("numerodefactura") = Replace(datencabezado.Recordset.Fields("tipodefacturacionid") + Str(xid) + "-" + Str(xclaveprimaria), " ", "")
    Else
       If nromanual = "" Then
        datencabezado.Recordset.Fields("numerodefactura") = datcola.Recordset.Fields("numero")
        datencabezado.Recordset.Fields("puntodeventa") = datcola.Recordset.Fields("puntoventa")
        xnumero = datcola.Recordset.Fields("numero")
        xidnumero = datcola.Recordset.Fields("numero_id")
        datcola.RecordSource = "Select * from numero with(readpast) where id = '" & xidnumero & "'"
        datcola.Refresh
        datcola.Recordset.Fields("numero") = xnumero + 1
        datcola.Recordset.UpdateBatch adAffectCurrent
       Else
        For X = 1 To Len(nromanual)
            car = Mid(nromanual, X, 1)
            If car = "-" Then
                datencabezado.Recordset.Fields("puntodeventa") = Left(nromanual, X - 1)
                datencabezado.Recordset.Fields("numerodefactura") = Mid(nromanual, X + 1, Len(nromanual) - X)
                Exit For
            Else
                datencabezado.Recordset.Fields("puntodeventa") = puntomanual
                datencabezado.Recordset.Fields("numerodefactura") = nromanual
            End If
        Next
                
'        datencabezado.Recordset.Fields("puntodeventa") = puntomanual
'        datencabezado.Recordset.Fields("numerodefactura") = nromanual
       End If

    End If
    '** Fin de asignacion de numero a Factura
    
    datencabezado.Recordset.Fields("claveprimaria") = xid
    datencabezado.Recordset.UpdateBatch adAffectCurrent
    
    
'--- Graba Items
    
    For X = 1 To xlineasmax
        If frmnota_venta.grilla.TextMatrix(1, 0) = "" Then
          mensa = MsgBox("No puede Grabar esta Venta sin Items", vbCritical, "Error")
          Exit Sub
        End If
        If frmnota_venta.grilla.TextMatrix(X, 0) = "" Then Exit For
        
        datitems.Recordset.AddNew
        datitems.Recordset.Fields("claveprimaria") = xid
        datitems.Recordset.Fields("idproducto") = frmnota_venta.grilla.TextMatrix(X, 0)
        datitems.Recordset.Fields("codigoproducto") = frmnota_venta.grilla.TextMatrix(X, 1)
        datitems.Recordset.Fields("nombre_producto") = frmnota_venta.grilla.TextMatrix(X, 2)
        datitems.Recordset.Fields("cantidadproducto") = frmnota_venta.grilla.TextMatrix(X, 3)
        datitems.Recordset.Fields("unidaddemedidaid") = frmnota_venta.grilla.TextMatrix(X, 4)
        datitems.Recordset.Fields("preciou") = Round(frmnota_venta.grilla.TextMatrix(X, 6), 4)
        datitems.Recordset.Fields("preciousiva") = Round(frmnota_venta.grilla.TextMatrix(X, 5), 4)
        datitems.Recordset.Fields("bonificacionitem") = frmnota_venta.grilla.TextMatrix(X, 9)
        datitems.Recordset.Fields("importedebonificacion") = Round(frmnota_venta.grilla.TextMatrix(X, 8), 4)
        datitems.Recordset.Fields("subtotal") = Round(frmnota_venta.grilla.TextMatrix(X, 10), 4)
        datitems.Recordset.Fields("iva") = (Round(frmnota_venta.grilla.TextMatrix(X, 12), 4) - 1) * 100
        datitems.Recordset.Fields("idclaveprimariaremito") = grilla.TextMatrix(X, 15)
        datitems.Recordset.Fields("iditemremito") = grilla.TextMatrix(X, 16)
        datitems.Recordset.UpdateBatch adAffectCurrent
    Next X
    


    datcontrol.RecordSource = "Select id, generada from ud_ezi_puntodeventa_encabezado with (readpast) where id = '" & presupuestobase & "'"
    datcontrol.Refresh
    datcontrol.Recordset.Fields("generada") = "True"
    datcontrol.Recordset.UpdateBatch adAffectCurrent

'******* Graba Pago
    datpago.RecordSource = "Select * from ud_ezi_pago where claveprimaria = ''"
    datpago.Refresh
        datpago.Recordset.AddNew
        datpago.Recordset.Fields("claveprimaria") = xid
        datpago.Recordset.Fields("tipovalor") = "True"
        datpago.Recordset.Fields("valoroseniaid") = ""
        datpago.Recordset.Fields("destinoid") = ""
        datpago.Recordset.Fields("formadepago") = "Debito en Cuenta Corriente"
        datpago.Recordset.Fields("monto") = Round(Text4.Text, 2)
        datpago.Recordset.Fields("sucursal") = login.nomsucursal

        datpago.Recordset.UpdateBatch adAffectCurrent
        
'******* Graba ud_ezi_cola  y/o Cola importar
    
    If datencabezado.Recordset.Fields("tipodefacturacionid") = "CF" Then
        datcola.RecordSource = "select * from ud_ezi_cola where nombrepc = '1'"
        datcola.Refresh
        
        datcola.Recordset.AddNew
        datcola.Recordset.Fields("nombrepc") = Environ("computername")
        datcola.Recordset.Fields("numero") = datencabezado.Recordset.Fields("numerodefactura")
        datcola.Recordset.Fields("accion") = datencabezado.Recordset.Fields("tipodefactura")
        datcola.Recordset.Fields("target") = datparametros.Recordset.Fields("target")
        datcola.Recordset.Fields("claveprimaria") = xid
    
        datcola.Recordset.UpdateBatch adAffectCurrent
    Else
        datcolaimportar.RecordSource = "Select * from ud_ezi_cola_importar where id = 0"
        datcolaimportar.Refresh
        
        datcolaimportar.Recordset.AddNew
        datcolaimportar.Recordset.Fields("id_encabezado") = xid
        datcolaimportar.Recordset.Fields("tipodedocumentoid") = datparametros.Recordset.Fields("idfacctacte")
        datcolaimportar.Recordset.Fields("unidadoperativaid") = datparametros.Recordset.Fields("target")
        datcolaimportar.Recordset.Fields("fecha_hora") = DateValue(DFECHA.Value) + TimeValue(Str(Time))
        
        datcolaimportar.Recordset.UpdateBatch adAffectCurrent
                
    End If

    If datencabezado.Recordset.Fields("tipodefacturacionid") <> "CF" Then
        Call imprimefactura_Click
    End If
    
'    mensa = MsgBox("Factura de Cta.Cte. Grabada Correctamente", vbInformation, "Registro Correcto !!")

    Call cancelar2_Click
    Unload Me
    
Exit Sub
errorgrabar:
    mensa = MsgBox("No se pudo registrar la información", vbCritical, "Error !!")




End Sub

Private Sub grilla_Click()

       If grilla.Col = 4 Then
          Call UM_Click
          Exit Sub
       End If

Text2.Visible = True

Call ubicatextogrilla_Click


End Sub

Private Sub grilla_GotFocus()

Call calcula_Click



End Sub

Private Sub grilla_KeyDown(KeyCode As Integer, Shift As Integer)

       If grilla.Col = 4 Then
          Call UM_Click
          Exit Sub
       End If

Call ubicatextogrilla_Click

End Sub



Private Sub grilla_KeyUp(KeyCode As Integer, Shift As Integer)

       If grilla.Col = 4 Then
          Call UM_Click
          Exit Sub
       End If

Call ubicatextogrilla_Click

End Sub

Private Sub grilla_LostFocus()

    Call calcula_Click

End Sub

Private Sub grilla_Scroll()

Text2.Visible = False

End Sub

Private Sub grremitoctacte_Click()

End Sub

Private Sub grremito_Click()
'On Error GoTo errorgrabar
    
    datencabezado.RecordSource = "SELECT MAX(CONVERT(decimal, isnull(claveprimaria,0))) AS claveprimaria FROM ud_ezi_puntodeventa_encabezado with(readpast) "
    datencabezado.Refresh
    
    If datencabezado.Recordset.EOF = True Then
        xclaveprimaria = 1
    Else
        xclaveprimaria = datencabezado.Recordset.Fields("claveprimaria") + 1
    End If
    
    datencabezado.RecordSource = "SELECT * From ud_ezi_puntodeventa_encabezado with(readpast) where id =0 "
    datencabezado.Refresh
    datitems.RecordSource = "Select * from ud_ezi_puntodeventa_detalle_rem with(readpast) where id = 0"
    datitems.Refresh
    
    datencabezado.Recordset.AddNew
    datencabezado.Recordset.Fields("claveprimaria") = xclaveprimaria
    If tpago <> "CONTADO" Then
        datencabezado.Recordset.Fields("numeradorinterno") = "Remito de Venta"
    Else
        datencabezado.Recordset.Fields("numeradorinterno") = "Remito de Venta Contado"
    End If
    datencabezado.Recordset.Fields("fechadelcomprobante") = DateValue(frmnota_venta.DFECHA.Value) + TimeValue(Str(Time))
    datencabezado.Recordset.Fields("sucursal") = frmnota_venta.datparametros.Recordset.Fields("sucursal")
    datencabezado.Recordset.Fields("clienteid") = frmnota_venta.DataGrid2.Columns(0).Text
    datencabezado.Recordset.Fields("cliente") = frmnota_venta.DataGrid2.Columns(2).Text
    If Text1(6).Text <> "" Then
        datencabezado.Recordset.Fields("cliente") = Text1(6).Text
        datencabezado.Recordset.Fields("recetaid") = Text1(3).Text
    End If
    datencabezado.Recordset.Fields("vendedorid") = frmnota_venta.DataGrid1.Columns(0).Text
    datencabezado.Recordset.Fields("vendedor") = frmnota_venta.DataGrid1.Columns(2).Text
    datencabezado.Recordset.Fields("detalle") = frmnota_venta.Text1(5).Text
    datencabezado.Recordset.Fields("nota") = frmnota_venta.Text15.Text
    datencabezado.Recordset.Fields("cotizacion") = 1
    datencabezado.Recordset.Fields("listadeprecioid") = frmnota_venta.DataCombo1.BoundText
    datencabezado.Recordset.Fields("tipodepagoid") = frmnota_venta.DataCombo3.BoundText
    datencabezado.Recordset.Fields("alquiler") = "N"
    
    If frmnota_venta.tipofac <> "NN" Then
        datencabezado.Recordset.Fields("tipodefacturacionid") = frmnota_venta.datparametros.Recordset.Fields("remitodefecto")
    Else
        datencabezado.Recordset.Fields("tipodefacturacionid") = frmnota_venta.tipofac
    End If
    
    If Left(login.nombrebd, 14) = "MMOSSE" And frmnota_venta.tipofac <> "NN" Then
        If frmnota_venta.tpago = "CONTADO" Then
            datencabezado.Recordset.Fields("tipodefacturacionid") = frmnota_venta.datparametros.Recordset.Fields("facdefecto")
        Else
            datencabezado.Recordset.Fields("tipodefacturacionid") = frmnota_venta.datparametros.Recordset.Fields("remitodefecto")
        End If
    End If
    
    If tpago = "CONTADO" Then
        datencabezado.Recordset.Fields("tipodefacturacionid") = "MA"
    End If
    
    datencabezado.Recordset.Fields("fechadeentrega") = DateValue(frmnota_venta.DFECHA.Value) + TimeValue(Str(Time))
    If frmnota_venta.Text13.Text = "" Then frmnota_venta.Text13.Text = 0
    datencabezado.Recordset.Fields("recargo") = Round(frmnota_venta.Text13.Text, 2)
    datencabezado.Recordset.Fields("tiporecargo") = "$"
    If frmnota_venta.Text11.Text = "" Then frmnota_venta.Text11.Text = 0
    datencabezado.Recordset.Fields("bonificacion") = Round(frmnota_venta.Text11.Text, 2)
    datencabezado.Recordset.Fields("tipobonificacion") = "$"
    datencabezado.Recordset.Fields("importeglobal") = Round(frmnota_venta.Text4.Text, 2)
    datencabezado.Recordset.Fields("domicilioid") = frmnota_venta.Text1(3).Text
    datencabezado.Recordset.Fields("domicilio_id") = frmnota_venta.DataGrid2.Columns("domicilio_id").Text
    datencabezado.Recordset.Fields("domiciliodeentregaid") = frmnota_venta.DataGrid2.Columns("domicilio_id").Text
    datencabezado.Recordset.Fields("subtotalsiniva") = Round(frmnota_venta.Text5.Text, 2)
    datencabezado.Recordset.Fields("totaliva") = Round(frmnota_venta.Text6.Text, 2) + Round(frmnota_venta.Text7.Text, 2)
    datencabezado.Recordset.Fields("generada") = "True"
    datencabezado.Recordset.Fields("importado") = "False"
    datencabezado.Recordset.Fields("nombrepc") = Environ("computername")
    datencabezado.Recordset.Fields("target") = datparametros.Recordset.Fields("sucursal")
    If DataGrid2.Columns(16).Text = "RI" Then xresponsabilidad = "Resp. Inscripto"
    If DataGrid2.Columns(16).Text = "EX" Then xresponsabilidad = "Exento"
    If DataGrid2.Columns(16).Text = "MT" Then xresponsabilidad = "Monotributista"
    If DataGrid2.Columns(16).Text = "CF" Then xresponsabilidad = "Consumidor Final"
    datencabezado.Recordset.Fields("responsabilidad") = xresponsabilidad
    datencabezado.Recordset.Fields("transferido") = "False"
    datencabezado.Recordset.Fields("tipodefactura") = "R"
    datencabezado.Recordset.Fields("nota") = "A"
    datencabezado.Recordset.Fields("percepiibb") = Round(frmnota_venta.Text9.Text, 2)
    datencabezado.Recordset.Fields("perceptem") = Round(frmnota_venta.Text8.Text, 2)
    datencabezado.Recordset.Fields("totaltr") = Round(frmnota_venta.Text4.Text, 2)
    datencabezado.Recordset.Fields("presupuestobase") = presupuestobase
    
    datencabezado.Recordset.UpdateBatch adAffectCurrent
    xid = datencabezado.Recordset.Fields("id")
    xremito = xid
    
    
    '** Establene numero de Remitos Manuales, y no Fiscales
    If datencabezado.Recordset.Fields("tipodefacturacionid") = "MA" Then
       If tpago <> "CONTADO" Then
            xnumerador = "Remito A (Vtas) " + datparametros.Recordset.Fields("sucursal")
       Else
            xnumerador = "Remito de Venta Mostrador " + datparametros.Recordset.Fields("sucursal")
       End If
    End If
    If datencabezado.Recordset.Fields("tipodefacturacionid") = "NN" Then
        If tpago <> "CONTADO" Then
            xnumerador = "Remito Val " + datparametros.Recordset.Fields("sucursal")
        Else
            xnumerador = "Remito de Venta Mostrador Val" + datparametros.Recordset.Fields("sucursal")
       End If
    End If
    datcola.RecordSource = "SELECT p.ID AS ID, p.NOMBRE AS puntero, n.NUMERO AS NUMERO, p.CARACTERESPREFIJO AS puntoventa, p.NUMERO_ID " & _
                           "FROM V_NUMERADOR_ AS p LEFT OUTER JOIN V_NUMERO_ AS n ON p.NUMERO_ID = n.ID " & _
                           "WHERE     (p.ACTIVESTATUS <> 2) AND (p.NOMBRE = '" & xnumerador & "') "
    datcola.Refresh
    If datencabezado.Recordset.Fields("tipodefacturacionid") = "CF" Then
        datencabezado.Recordset.Fields("numerodefactura") = Replace(datencabezado.Recordset.Fields("tipodefacturacionid") + Str(xid) + "-" + Str(xclaveprimaria), " ", "")
    Else
      If Text1(8).Text = "" Then
        datencabezado.Recordset.Fields("numerodefactura") = datcola.Recordset.Fields("numero")
        xnumero = datcola.Recordset.Fields("numero")
        xidnumero = datcola.Recordset.Fields("numero_id")
        datencabezado.Recordset.Fields("puntodeventa") = datcola.Recordset.Fields("puntoventa")
        datcola.RecordSource = "Select * from numero with(readpast) where id = '" & xidnumero & "'"
        datcola.Refresh
        datcola.Recordset.Fields("numero") = xnumero + 1
        datcola.Recordset.UpdateBatch adAffectCurrent
       Else
         datencabezado.Recordset.Fields("numerodefactura") = Text1(8).Text
         datencabezado.Recordset.Fields("puntodeventa") = "0002"
       End If
    End If
    '** Fin de asignacion de numero a Remtio
    
    datencabezado.Recordset.Fields("claveprimaria") = xid
    datencabezado.Recordset.UpdateBatch adAffectCurrent
    
    
    
'--- Graba Items
    
    For X = 1 To xlineasmax
        If frmnota_venta.grilla.TextMatrix(1, 0) = "" Then
          mensa = MsgBox("No puede Grabar esta Venta sin Items", vbCritical, "Error")
          Exit Sub
        End If
        If frmnota_venta.grilla.TextMatrix(X, 0) = "" Then Exit For
        
        datitems.Recordset.AddNew
        datitems.Recordset.Fields("claveprimaria") = xid
        datitems.Recordset.Fields("idproducto") = frmnota_venta.grilla.TextMatrix(X, 0)
        datitems.Recordset.Fields("referenciaproducto") = frmnota_venta.grilla.TextMatrix(X, 1)
        datitems.Recordset.Fields("nombre_producto") = frmnota_venta.grilla.TextMatrix(X, 2)
        datitems.Recordset.Fields("cantidadoriginal") = frmnota_venta.grilla.TextMatrix(X, 3)
        datitems.Recordset.Fields("cantidadremitida") = frmnota_venta.grilla.TextMatrix(X, 3)
        datitems.Recordset.Fields("cantidadaremitir") = frmnota_venta.grilla.TextMatrix(X, 3)
        
        datitems.Recordset.Fields("unidaddemedida") = frmnota_venta.grilla.TextMatrix(X, 4)
        
        datitems.Recordset.Fields("facturaorigen") = xid
        datitems.Recordset.UpdateBatch adAffectCurrent
        frmnota_venta.grilla.TextMatrix(X, 15) = xid
        frmnota_venta.grilla.TextMatrix(X, 16) = datitems.Recordset.Fields("id")
    Next X
'*** Graba iditem de remito en Items de Nota de Venta
    datitemsnv.RecordSource = "Select * from ud_ezi_puntodeventa_detalle_notav where claveprimaria = " & presupuestobase & " order by id"
    datitemsnv.Refresh
    datitemsnv.Recordset.MoveFirst
    For X = 1 To xlineasmax
        If frmnota_venta.grilla.TextMatrix(X, 0) = "" Then Exit For
        datitemsnv.Recordset.Fields("idclaveprimariaremito") = grilla.TextMatrix(X, 15)
        datitemsnv.Recordset.Fields("iditemremito") = grilla.TextMatrix(X, 16)
        datitemsnv.Recordset.UpdateBatch adAffectCurrent
        datitemsnv.Recordset.MoveNext
    Next X
'*** Fin Graba iditem de remito en Items de Nota de Venta
    
    
    datcontrol.RecordSource = "Select id, generada from ud_ezi_puntodeventa_encabezado with (readpast) where id = '" & presupuestobase & "'"
    datcontrol.Refresh
    datcontrol.Recordset.Fields("generada") = "True"
    datcontrol.Recordset.UpdateBatch adAffectCurrent

'******* Graba ud_ezi_cola  y/o Cola importar
    
    If datencabezado.Recordset.Fields("tipodefacturacionid") = "CF" Then
        datcola.RecordSource = "select * from ud_ezi_cola where nombrepc = '1'"
        datcola.Refresh
    
        datcola.Recordset.AddNew
        datcola.Recordset.Fields("nombrepc") = Environ("computername")
        datcola.Recordset.Fields("numero") = datencabezado.Recordset.Fields("numerodefactura")
        datcola.Recordset.Fields("accion") = "R"
        datcola.Recordset.Fields("target") = datparametros.Recordset.Fields("target")
        datcola.Recordset.Fields("claveprimaria") = xid
    
        datcola.Recordset.UpdateBatch adAffectCurrent
    Else
      If tpago <> "CONTADO" Then
        datcolaimportar.RecordSource = "Select * from ud_ezi_cola_importar where id = 0"
        datcolaimportar.Refresh
        
        datcolaimportar.Recordset.AddNew
        datcolaimportar.Recordset.Fields("id_encabezado") = xid
        datcolaimportar.Recordset.Fields("tipodedocumentoid") = datparametros.Recordset.Fields("idremito")
        datcolaimportar.Recordset.Fields("unidadoperativaid") = datparametros.Recordset.Fields("target")
        datcolaimportar.Recordset.Fields("fecha_hora") = DateValue(DFECHA.Value) + TimeValue(Str(Time))
        
        datcolaimportar.Recordset.UpdateBatch adAffectCurrent
      End If
    End If

''**** Imprime remito
    
    If tpago <> "CONTADO" And datencabezado.Recordset.Fields("tipodefacturacionid") = "MA" Then
        Call imprimeremito_Click
    End If


    'mensa = MsgBox("Remito Grabado Correctamente", vbInformation, "Registro Correcto !!")
   
Exit Sub
errorgrabar:
    mensa = MsgBox("No se pudo registrar la información", vbCritical, "Error !!")


End Sub

Private Sub imprimefactura_Click()
On Error GoTo fuera
Dim crApp As New CRAXDRT.Application
Dim crReport As New CRAXDRT.Report
Dim tabla As String
Dim ruta As String

If datparametros.Recordset.Fields("imprimemanual") = "N" Or Text1(7).Text <> "" Then Exit Sub

reporte.SQL = "SELECT v_ezi_pos_factctacte.id, v_ezi_pos_factctacte.NUMERODOCUMENTO, v_ezi_pos_factctacte.FECHAEMISION, v_ezi_pos_factctacte.cod_cliente, v_ezi_pos_factctacte.cliente, v_ezi_pos_factctacte.CUIT, v_ezi_pos_factctacte.CALLE, v_ezi_pos_factctacte.CODPOS, v_ezi_pos_factctacte.provincia, v_ezi_pos_factctacte.detalle, v_ezi_pos_factctacte.tipopago, v_ezi_pos_factctacte.codigoproducto, v_ezi_pos_factctacte.nombre_producto, v_ezi_pos_factctacte.cantidadproducto, v_ezi_pos_factctacte.nota, v_ezi_pos_factctacte.condiva, v_ezi_pos_factctacte.ciudad, v_ezi_pos_factctacte.TIPOVENTA, v_ezi_pos_factctacte.SIMBOLO, v_ezi_pos_factctacte.CODVENDEDOR, v_ezi_pos_factctacte.preciusiniva, v_ezi_pos_factctacte.subtotalsiniva, v_ezi_pos_factctacte.impbonifsiniva, v_ezi_pos_factctacte.nroremito, v_ezi_pos_factctacte.percepiibb, v_ezi_pos_factctacte.perceptem, v_ezi_pos_factctacte.totaltr, v_ezi_pos_factctacte.importeiva21, v_ezi_pos_factctacte.importeiva105, v_ezi_pos_factctacte.iditem " & _
              "FROM  MMOSSE.dbo.v_ezi_pos_factctacte v_ezi_pos_factctacte " & _
              "where v_ezi_pos_factctacte.id = " & xid & " order by v_ezi_pos_factctacte.iditem"

tabla = reporte.SQL

With CrystalReporte
    .PrinterCollation = crptCollated
    If tipofac <> "NN" Then
        .Formulas(0) = "copia="" ORIGINAL """
    End If
    If Text1(2).Text = "A" Then
      If tipofac <> "NN" Then
        .ReportFileName = App.Path & "\FacturaCtaCteA.rpt"
      Else
        .ReportFileName = App.Path & "\PresupuestoA.rpt"
      End If
    Else
      If tipofac <> "NN" Then
        .ReportFileName = App.Path & "\FacturaCtaCteB.rpt"
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
 Rem   .Destination = crptToWindowfrmnota_venta.Text1(2).Text
    .Destination = crptToPrinter
    .PrintFileType = crptCrystal
    .WindowState = crptMaximized
    .Action = 1
    If tipofac <> "NN" Then
    .WindowTitle = "Factura Vta Dupl"
    .Formulas(0) = "copia="" DUPLICADO """
    .Action = 1
     If Text1(2).Text = "A" Then
      .WindowTitle = "Factura Vta Trip"
      .Formulas(0) = "copia="" TRIPLICADO """
      .Action = 1
     End If
    End If
    
End With
Exit Sub

fuera:
    
    MsgBox "Reporte de Factura no Encontado, o error de configuracion de reporte", vbCritical, "Error"


End Sub

Private Sub imprimeremito_Click()
'On Error GoTo fuera
Dim crApp As New CRAXDRT.Application
Dim crReport As New CRAXDRT.Report
Dim tabla As String
Dim ruta As String

If datparametros.Recordset.Fields("imprimemanual") = "N" Or Text1(8).Text <> "" Then Exit Sub

reporte.SQL = "SELECT v_ezi_pos_remito.id, v_ezi_pos_remito.NUMERODOCUMENTO, v_ezi_pos_remito.FECHAEMISION, v_ezi_pos_remito.cod_cliente, v_ezi_pos_remito.cliente, v_ezi_pos_remito.CALLE, v_ezi_pos_remito.CODPOS, v_ezi_pos_remito.provincia, v_ezi_pos_remito.detalle, v_ezi_pos_remito.tipopago, v_ezi_pos_remito.referenciaproducto, v_ezi_pos_remito.nombre_producto, v_ezi_pos_remito.cantidadremitida, v_ezi_pos_remito.nota, v_ezi_pos_remito.condiva, v_ezi_pos_remito.ciudad, v_ezi_pos_remito.TIPOVENTA, v_ezi_pos_remito.SIMBOLO, v_ezi_pos_remito.iditem FROM MMOSSE.dbo.v_ezi_pos_remito v_ezi_pos_remito where v_ezi_pos_remito.id = " & xremito & " order by v_ezi_pos_remito.iditem"
tabla = reporte.SQL

With CrystalReporte
    .PrinterCollation = crptCollated
    .Formulas(0) = "copia="" ORIGINAL """
    .ReportFileName = App.Path & "\RemitoVta.rpt"
    .WindowTitle = "Remito Vta Orig"
    '.Connect = "PROVIDER=MSDASQL;dsn=facturacion;uid=lucva;pwd=25072004;database=facturacionsql;"
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
Exit Sub

fuera:
    
    MsgBox "Reporte de Factura no Encontado, o error de configuracion de reporte", vbCritical, "Error"




End Sub

Private Sub KewlButtons1_Click()
On Error Resume Next
        For X = 0 To 10
            grilla.Col = X
            grilla.Text = ""
        Next X
        
        xmemoriarow = grilla.Row
        For X = grilla.Row + 1 To 48
            For Y = 0 To 13
                grilla.Col = Y
                grilla.Row = X
                xcampo = grilla.Text
                grilla.Row = X - 1
                grilla.Text = xcampo
                Text2.Text = ""
                Combo1.Clear
            Next Y
        Next X
        grilla.Row = xmemoriarow
    Call calcula_Click
    Call ubicatextogrilla_Click

End Sub

Private Sub KewlButtons2_Click()
On Error Resume Next
    Frame4.Visible = True
    Text2.Visible = False
    Frame4.Caption = Text1(1).Text
    Text20.Text = xclienteinfo

    
End Sub

Private Sub limitecredito_Click()
On Error Resume Next


  xcreditomaximo = DataGrid2.Columns("creditomaximo")
  xdiasplazo = DataGrid2.Columns("diasplazo")
  xcontrol = 0
  If xcreditomaximo = 0 And xdiasplazo = 0 Then Exit Sub
    If login.nomsucursal = "EMPORIO" Or login.nomsucursal = "EMPORIOZIP" Then
        xcp = "EL EMP.TUCUMAN"
    Else
        xcp = "DIM TOLEDO"
    End If
    
    'datcredito.RecordSource = "select * from v_ezi_pos_ctacte_control where id = '" & DataGrid2.Columns(0).Text & "' and nomclasificador = '" & xcp & "'"
    datcredito.RecordSource = "select * from v_ezi_pos_ctacte_control where id = '" & DataGrid2.Columns(0).Text & "'"
    datcredito.Refresh
    
    xcontrol = 0
    If datcredito.Recordset.EOF = True Then
        xdisponible = xcreditomaximo
    Else
        xsaldo = datcredito.Recordset.Fields("saldo")
        xdisponible = Round(xcreditomaximo - xsaldo, 2)
        xcontrol = datcredito.Recordset.Fields("control")
    End If
    If xdisponible <> 0 And xcontrol = 0 Then
        Text1(5).Text = DataGrid2.Columns(6).Text + " -- C.Disp.:$ " + Str(xdisponible)
    End If
    
    If xdisponible >= 0 And xcontrol = 0 Then
        blanco.Visible = True
        negro.Visible = False
    Else
        blanco.Visible = False
        negro.Visible = True
        If xcontrol = 1 Then
            MsgBox "Este cliente tiene Facturas Vencidas, no podra grabar esta NV", vbCritical, "Control de Crédito"
            Exit Sub
        End If
    End If


End Sub

Private Sub negro_Click()


    negro.Visible = False
    blanco.Visible = True
    tipofac = "CF"
    Call calcula_Click
    

End Sub

Private Sub salir_Click()

    mensa = MsgBox("Desea Salir de este documento", vbYesNo, "Atención !!")
    If mensa = vbYes Then
        Unload Me
    End If

End Sub


Private Sub Text1_Change(Index As Integer)

    If Index = 7 Then
        If Not IsNumeric(Text1(Index).Text) Then
            Text1(Index).Text = ""
            mensa = MsgBox("Valor Numerico no Valido ", vbCritical, "Error !!")
        End If
    End If

End Sub

Private Sub Text1_GotFocus(Index As Integer)

    If Index = 1 Then
        If Text1(0).Text = "" Then mensa = MsgBox("Debe ingresar un vendedor", vbCritical, "Error")
    End If
    If Index = 5 Then
        If Text1(0).Text = "" Then mensa = MsgBox("Debe ingresar un vendedor", vbCritical, "Error")
        If Text1(1).Text = "" Then mensa = MsgBox("Debe ingresar un Cliente", vbCritical, "Error")
    End If

End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
On Error Resume Next


    If KeyAscii = 13 Then
        KeyAscii = 0
        Text1(Index).Text = UCase(Text1(Index).Text)
        If Index = 0 Then
            Call bvendedor_Click
            If Text1(0).Text <> "" Then
                menu = 1
                If datvendedor.Recordset.RecordCount = 1 Then
                    lista_clavevendedor.Show
                End If
            End If
        End If
        If Index = 1 Then
            If Text1(1).Text = "" Then Text1(1).Text = "CONSUMIDOR FINAL"
            Call bclientes_Click
        End If
        
        If Index = 6 Then
            Text1(3).Text = ""
            Text1(3).SetFocus
        End If
        If Index = 3 Then
            Text1(5).SetFocus
        End If
        If Index = 5 Then
            Text1(9).SetFocus
        End If
        
        If Index = 9 Then
            agregaproducto.SetFocus
            Call agregaproducto_Click
        End If
        If Index = 7 Then
            xnumerocontrol = datparametros.Recordset.Fields("puntovtamanual") + Right("00000000" + Text1(7).Text, 8)
            login.datcontrol.RecordSource = "SELECT     NUMERODOCUMENTO From V_TRFACTURAVENTA_ Where (FLAG_ID Is Null) and numerodocumento = '" & xnumerocontrol & "'"
            login.datcontrol.Refresh
            If login.datcontrol.Recordset.EOF = False Then
                mensa = MsgBox("Nro de Factura Manual " + xnumerocontrol + " Existente, verifique Nro", vbCritical, "Error")
                Text1(7).Text = ""
            End If
            mensa = MsgBox("Si la factura corresponde a una factura electronica ingrese el cae en el campo NOTA", vbInformation)
            Text1(5).Text = ""
            Text1(5).SetFocus
        End If
        
        
    End If
    
End Sub

Private Sub Text1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)

If KeyCode = 114 Then
       Call bvendedor_Click
End If

If KeyCode = 116 Then
       Call agregaproducto_Click
End If

If KeyCode = 38 Then
    If Index = 5 Then Text1(1).SetFocus
    If Index = 1 Then Text1(0).SetFocus
End If

If KeyCode = 121 Then
    KeyCode = 0
       Call grabar_Click
End If


End Sub

Private Sub Text1_LostFocus(Index As Integer)
On Error Resume Next
        If Index = 2 Then
            For X = 1 To Len(Text1(2).Text)
               car = Mid(Text1(2).Text, X, 1)
               If car = "-" Then
                  PVta = Right("0000" + Left(Text1(2).Text, X - 1), 4)
                  nu = Right("00000000" + Right(Text1(2).Text, Len(Text1(2).Text) - X), 8)
                  Text1(2).Text = PVta + nu
                  Exit For
               End If
               Text1(2).Text = Right("00000000" + Text1(2).Text, 8)
            Next X
        End If
        
        If Index = 1 And Text1(1).Text <> "" Then
            Text1(1).Text = DataGrid2.Columns(2).Text
        End If
        
        If Index = 0 And Text1(0).Text <> "" Then
            Call bvendedor_Click
            If Text1(0).Text <> "" Then
                menu = 1
                If datvendedor.Recordset.RecordCount = 1 Then
                    lista_clavevendedor.Show
                End If
            End If
        End If

End Sub

Private Sub Text10_GotFocus()

    Text10.SelStart = 0
    Text10.SelLength = Len(Text10.Text)

End Sub

Private Sub Text10_KeyPress(KeyAscii As Integer)
On Error Resume Next

If KeyAscii = 13 Then
        KeyAscii = 0
        Text11.Text = Format(0, "##0.00")
        Text14.Text = Format(0, "##0.00")
        Text10.Text = Format(Text10.Text, "##0.00")
        ximporteboniftotal = 0
        For X = 1 To xlineasmax
              If grilla.TextMatrix(X, 1) = "" Then Exit For
              grilla.Row = X
              grilla.TextMatrix(X, 9) = Text10.Text
              If Val(Text10.Text) = 0 Then
                 grilla.TextMatrix(X, 8) = 0
              End If
              Call calcula_Click
              ximporteboniftotal = ximporteboniftotal + grilla.TextMatrix(X, 8)
        Next X
        Text11.Text = Format(ximporteboniftotal, "##0.00")
        Text15.SetFocus
End If


End Sub

Private Sub Text10_KeyUp(KeyCode As Integer, Shift As Integer)


If KeyCode = 116 Then
       Call agregaproducto_Click
End If

If KeyCode = 121 Then
       KeyCode = 0
       Call grabar_Click
End If



End Sub

Private Sub Text11_GotFocus()


    Text11.SelStart = 0
    Text11.SelLength = Len(Text11.Text)

End Sub

Private Sub Text11_KeyPress(KeyAscii As Integer)
On Error Resume Next

If KeyAscii = 13 Then
        KeyAscii = 0

        xvalorsubtotal = Round(Text5.Text, 10) + Round(Text6.Text, 10) + Round(Text7.Text, 10)
        xporcenboniftotal = 0
        For X = 1 To xlineasmax
              If grilla.TextMatrix(X, 1) = "" Then Exit For
              grilla.Row = X
              grilla.TextMatrix(X, 8) = Round((grilla.TextMatrix(X, 10) / xvalorsubtotal) * Text11.Text, 5)
              If Val(Text11.Text) = 0 Then
                 grilla.TextMatrix(X, 9) = 0
              End If
              Call calcula_Click
       
        Next X
        Text11.Text = Format(Text11.Text, "##0.00")
        Text14.Text = Format(0, "##0.00")
        Text10.Text = Format(0, "##0.00")
        Text10.Text = Format(grilla.TextMatrix(1, 9), "##0.00")
        Text15.SetFocus
End If

End Sub

Private Sub Text11_KeyUp(KeyCode As Integer, Shift As Integer)

If KeyCode = 116 Then
       Call agregaproducto_Click
End If

If KeyCode = 121 Then
       KeyCode = 0
       Call grabar_Click
End If



End Sub

Private Sub Text12_GotFocus()
    Text12.SelStart = 0
    Text12.SelLength = Len(Text12.Text)

End Sub

Private Sub Text12_KeyPress(KeyAscii As Integer)

On Error Resume Next

If KeyAscii = 13 Then
        KeyAscii = 0
        
'        grilla.TextMatrix(x, 5) = grilla.TextMatrix(X, 13)
        Text13.Text = 0
        Call calcula_Click
        
        Text13.Text = Format(0, "##0.00")
        Text14.Text = Format(0, "##0.00")
        ximporteboniftotal = Round(Text16.Text, 20)
        For X = 1 To xlineasmax
              If grilla.TextMatrix(X, 1) = "" Then Exit For
              grilla.Row = X
              
              grilla.TextMatrix(X, 6) = Round(grilla.TextMatrix(X, 6), 20) * ((Round(Text12.Text, 20) / 100) + 1)
              'grilla.TextMatrix(X, 6) = Round(grilla.TextMatrix(X, 14), 20) * ((Round(Text12.Text, 20) / 100) + 1)
              If Val(Text12.Text) = 0 Then
                 grilla.TextMatrix(X, 6) = grilla.TextMatrix(X, 14)
              End If
              Call calcula_Click
        Next X
        ximporteboniftotal = Round(Text16.Text, 20) - ximporteboniftotal
        If Val(Text12.Text) = 0 Then
            Text13.Text = Format(0, "##0.00")
        Else
            Text13.Text = Format(ximporteboniftotal, "###,##0.00")
        End If
        Text12.Text = Format(Text12.Text, "##0.00")
        Text15.SetFocus
End If


End Sub

Private Sub Text12_KeyUp(KeyCode As Integer, Shift As Integer)

If KeyCode = 116 Then
       Call agregaproducto_Click
End If

If KeyCode = 121 Then
       KeyCode = 0
       Call grabar_Click
End If


End Sub

Private Sub Text13_GotFocus()


        Text13.SelStart = 0
        Text13.SelLength = Len(Text13.Text)



End Sub

Private Sub Text13_KeyPress(KeyAscii As Integer)

On Error Resume Next

If KeyAscii = 13 Then
        KeyAscii = 0
        Text14.Text = Format(0, "##0.00")
        Text13.Text = Format(Text13.Text, "##0.00")
        ximporteboniftotal = (Round(Text13.Text, 20) / Round(Text16.Text, 20)) * 100
        Text12.Text = ximporteboniftotal
        Text12.SetFocus
        SendKeys "{ENTER}", False
        Text2.Text = Format(Text2.Text, "##0.00")
'        Text15.SetFocus
End If


End Sub

Private Sub Text13_KeyUp(KeyCode As Integer, Shift As Integer)


If KeyCode = 116 Then
       Call agregaproducto_Click
End If

If KeyCode = 121 Then
       KeyCode = 0
       Call grabar_Click
End If



End Sub

Private Sub Text14_GotFocus()


    Text14.SelStart = 0
    Text14.SelLength = Len(Text10.Text)

End Sub

Private Sub Text14_KeyPress(KeyAscii As Integer)

On Error Resume Next

If KeyAscii = 13 Then
        KeyAscii = 0
              
        xxiibb = Round(Text14.Text, 2) * Round(Text9.Text, 2) / Round(Text4.Text, 2)
        xxtem = Round(Text14.Text, 2) * Round(Text8.Text, 2) / Round(Text4.Text, 2)
        xxiva21 = Round(Text14.Text, 2) * Round(Text7.Text, 2) / Round(Text4.Text, 2)
        xxiva10 = Round(Text14.Text, 2) * Round(Text6.Text, 2) / Round(Text4.Text, 2)
        xxsubtotal1 = Round(Text14.Text, 2) - xxiibb - xxtem - xxiva21 - xxiva10
        xxsubtotal2 = xxsubtotal1 + xxiva21 + xxiva10
        xsigno = xxsubtotal2 - Round(Text16.Text, 2)
        'xdiferencia = Abs(xxsubtotal2 - Round(Text16.Text, 2))
        xdiferencia = Abs(xxsubtotal1 - Round(Text5.Text, 2))
        
        
        
        If xsigno > 0 Then
                Text13.Text = xdiferencia
                Text13.SetFocus
                SendKeys "{ENTER}", False
'                Text13.Text = Format(Text13.Text, "##0.00")

        Else
        
                Text11.Text = xdiferencia
                Text11.SetFocus
                SendKeys "{ENTER}", False
'                Text11.Text = Format(Text11.Text, "##0.00")
                Text11.SetFocus
                
        End If

End If



End Sub

Private Sub Text14_KeyUp(KeyCode As Integer, Shift As Integer)

If KeyCode = 116 Then
       Call agregaproducto_Click
End If

If KeyCode = 121 Then
        KeyCode = 0
       Call grabar_Click
End If

End Sub

Private Sub Text15_KeyUp(KeyCode As Integer, Shift As Integer)

If KeyCode = 121 Then
       KeyCode = 0
       Call grabar_Click
End If

If KeyCode = 116 Then
       Call agregaproducto_Click
End If


End Sub

Private Sub Text17_KeyPress(KeyAscii As Integer)
On Error Resume Next

If KeyAscii = 13 Then
    If Text18.Text = "" And menu <> 6 Then
        KeyAscii = 0
        lista_presupuestos.Show
        lista_presupuestos.Text1.Text = Text17.Text
        lista_presupuestos.Text1.SetFocus
        SendKeys "{ENTER}", False
    Else
        Call cargapresupuesto_Click
    End If
    
End If

End Sub

Private Sub Text2_GotFocus()
On Error Resume Next
    
    If grilla.Col <> 2 Then
        Text2.SelStart = 0
        Text2.SelLength = Len(Text2.Text)
    End If

    

End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
On Error Resume Next

If KeyAscii = 13 Then
        Text2.Text = UCase(Text2.Text)
        KeyAscii = 0
        grilla.Text = Text2.Text
        If grilla.Col = 3 Then Command8_Click
        If grilla.Col = 5 Then
            grilla.TextMatrix(grilla.Row, 13) = Text2.Text
        End If
        If grilla.Col = 5 Then
            grilla.TextMatrix(grilla.Row, 6) = grilla.Text * grilla.TextMatrix(grilla.Row, 12)
        End If
            
        If grilla.Col = 5 Or grilla.Col = 6 Then
            grilla.TextMatrix(grilla.Row, 9) = "0.00"
        End If
            
        If grilla.Col = 9 And grilla.Text <> "0.00" Then
            grilla.TextMatrix(grilla.Row, 6) = grilla.TextMatrix(grilla.Row, 17)
            Call calcula_Click
        End If
        
        If grilla.Col = 3 Or grilla.Col = 5 Or grilla.Col = 6 Or grilla.Col = 7 Or grilla.Col = 8 Or grilla.Col = 9 Then Call calcula_Click
        
        If grilla.Col > 10 Then
                Call agregaproducto_Click
                Exit Sub
        End If
        grilla.Col = grilla.Col + 1
        If grilla.Col = 4 Then
           Call UM_Click
           Exit Sub
        End If
        If grilla.Col = 7 Then grilla.Col = 9
        If grilla.Col = 10 Then grilla.Col = 11
        
       
        Call ubicatextogrilla_Click
End If

    

End Sub


Private Sub Text4_KeyPress(KeyAscii As Integer)
On Error Resume Next

    If KeyAscii = 13 Then
        KeyAscii = 0
        Text5.SetFocus
    End If

End Sub


Private Sub Text5_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = 13 Then
        KeyAscii = 0
        autorizar.SetFocus
    End If
    

End Sub

Private Sub Text2_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next

If KeyCode = 39 And grilla.Col < 10 Then   'And Text2.SelStart = Len(Text2.Text)
  If grilla.Col <> 2 Then
    grilla.Col = grilla.Col + 1
       If grilla.Col = 4 Then
          Call UM_Click
          Exit Sub
       End If
    If grilla.Col = 7 Then grilla.Col = 9
    If grilla.Col = 10 Then grilla.Col = 11
    Call ubicatextogrilla_Click
  Else
    If Text2.SelStart = Len(Text2.Text) Then
        grilla.Col = grilla.Col + 1
        Call ubicatextogrilla_Click
    End If
  End If
End If

If KeyCode = 38 And grilla.Row > 0 Then
   If grilla.Row = 1 Then
     Text1(5).SetFocus
     Text2.Visible = False
     Exit Sub
   End If
    grilla.Row = grilla.Row - 1
    If grilla.RowIsVisible(grilla.Row) = False Then
        grilla.TopRow = grilla.TopRow - 1
    End If
    Call ubicatextogrilla_Click

End If

If KeyCode = 40 And grilla.Row < xlineasmax Then
    grilla.Row = grilla.Row + 1
    If grilla.RowIsVisible(grilla.Row) = False Then
        grilla.TopRow = grilla.TopRow + 1
    End If
    Call ubicatextogrilla_Click
End If



If KeyCode = 37 And grilla.Col > 2 Then ' And Text2.SelStart = 0
    grilla.Col = grilla.Col - 1
       If grilla.Col = 4 Then
          Call UM_Click
          Exit Sub
       End If
    If grilla.Col = 8 Then grilla.Col = 6
    If grilla.Col = 10 Then grilla.Col = 9
    Call ubicatextogrilla_Click
End If

If KeyCode = 116 Then
       Call agregaproducto_Click
End If

    
If KeyCode = 121 Then
       KeyCode = 0
       Call grabar_Click
End If

End Sub

Private Sub ubicatextogrilla_Click()
On Error Resume Next

Combo1.Visible = False
If grilla.Col = 7 Then
     Exit Sub
End If

Text2.Visible = True

If grilla.Col = 3 Or grilla.Col = 5 Or grilla.Col = 6 Or grilla.Col = 7 Or grilla.Col = 8 Or grilla.Col = 9 Then
    Text2.Alignment = 1
Else
    Text2.Alignment = 0
End If

If grilla.Col = 10 Then
    Text2.Enabled = False
Else
    Text2.Enabled = True
End If
    
Text2.Width = grilla.ColWidth(grilla.Col)
Text2.Text = grilla.Text

xfila = grilla.RowPos(grilla.Row) + grilla.Top + Picture1.Top + 50
xcolu = grilla.ColPos(grilla.Col) + grilla.Left + Picture1.Left + 50
Text2.Left = xcolu
Text2.Top = xfila

If grilla.Col <> 2 Then
    Text2.SelStart = 0
    Text2.SelLength = Len(Text2.Text)
End If

Text2.SetFocus

End Sub

Private Sub UM_Click()
On Error Resume Next
    
    xid2 = grilla.TextMatrix(grilla.Row, 0)
    If xid2 = "" Then Exit Sub
    
    Combo1.Visible = True
    Combo1.Clear
    
    
    
    
    datum.RecordSource = "SELECT     P.ID, P.CODIGO, P.UNIDADMEDIDA_ID, P.UNIDADMEDIDANOLINEAL_ID, V_UNIDADMEDIDA_.NOMBRE AS UMVTA, V_UNIDADMEDIDA__1.NOMBRE AS UMSTK, " & _
                         "ISNULL(V_ITEMCONVUNIDADESDEMEDIDA_.RELACIONCONVERSION, 1) AS RELACIONCONVERSION, V_UNIDADMEDIDA__2.NOMBRE AS UMDESTINO " & _
                         "FROM         V_UNIDADMEDIDA_ AS V_UNIDADMEDIDA__2 RIGHT OUTER JOIN " & _
                         "V_ITEMCONVUNIDADESDEMEDIDA_ ON V_UNIDADMEDIDA__2.ID = V_ITEMCONVUNIDADESDEMEDIDA_.UNIDADDESTINO_ID RIGHT OUTER JOIN " & _
                         "V_PRODUCTO_ AS P ON V_ITEMCONVUNIDADESDEMEDIDA_.UNIDADORIGEN_ID = P.UNIDADMEDIDANOLINEAL_ID AND " & _
                         "V_ITEMCONVUNIDADESDEMEDIDA_.BO_PLACE_ID = P.EQUIVALENCIASPARTICULARES_ID LEFT OUTER JOIN " & _
                         "V_UNIDADMEDIDA_ AS V_UNIDADMEDIDA__1 ON P.UNIDADMEDIDA_ID = V_UNIDADMEDIDA__1.ID LEFT OUTER JOIN " & _
                         "V_UNIDADMEDIDA_ ON P.UNIDADMEDIDANOLINEAL_ID = V_UNIDADMEDIDA_.ID " & _
                         "WHERE  P.ID = '" & xid2 & "'"
    datum.Refresh
    If datum.Recordset.EOF = False Then
      Combo1.AddItem datum.Recordset.Fields("UMVTA")
      xumvta = datum.Recordset.Fields("UMVTA")
      xlist = 1
      Do While Not datum.Recordset.EOF
        If datum.Recordset.Fields("UMSTK") <> datum.Recordset.Fields("UMVTA") Then
            Combo1.AddItem datum.Recordset.Fields("UMSTK")
            xfaconv(xlist) = datum.Recordset.Fields("relacionconversion")
            xlist = xlist + 1
        End If
        If datum.Recordset.Fields("UMDESTINO") <> datum.Recordset.Fields("UMVTA") And datum.Recordset.Fields("UMDESTINO") <> datum.Recordset.Fields("UMSTK") Then
            Combo1.AddItem datum.Recordset.Fields("UMDESTINO")
            xfaconv(xlist) = datum.Recordset.Fields("relacionconversion")
            xlist = xlist + 1
        End If
        
        datum.Recordset.MoveNext
       Loop
        
    Else
        Combo1.AddItem "Unidad"
    End If

Text2.Visible = False
    
Combo1.Width = grilla.ColWidth(grilla.Col) + 200
If grilla.Text = "" Then
    Combo1.ListIndex = 0
Else
    Combo1.Text = grilla.Text
End If


xfila = grilla.RowPos(grilla.Row) + grilla.Top + Picture1.Top + 30
xcolu = grilla.ColPos(grilla.Col) + grilla.Left + Picture1.Left + 30
Combo1.Left = xcolu
Combo1.Top = xfila

xfila = grilla.Row
Combo1.SetFocus
    
    

End Sub

Private Sub verificalotes_Click()
On Error Resume Next

    For j = 1 To xlineasmax
        If grilla.TextMatrix(j, 0) <> "" Then
            xcontrollote = grilla.TextMatrix(j, 18)
            If xcontrollote = "" Then
                MsgBox grilla.TextMatrix(j, 1) + " " + grilla.TextMatrix(j, 2), vbInformation, "Ingrese Lote o Valide Stock"
            End If
        Else
           Exit For
        End If
    Next j

End Sub
