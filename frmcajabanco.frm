VERSION 5.00
Object = "{F6125AB1-8AB1-11CE-A77F-08002B2F4E98}#2.0#0"; "MSRDC20.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{5734474E-78D3-4254-99B9-C35F31BDF509}#62.0#0"; "vbskpro2.ocx"
Begin VB.Form frmcajabanco 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Caja / Banco"
   ClientHeight    =   7320
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7320
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.CommandButton grabaasiento 
      Caption         =   "grabaasiento"
      Height          =   255
      Left            =   3480
      TabIndex        =   33
      Top             =   7080
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton limpia 
      Caption         =   "limpia"
      Height          =   255
      Left            =   1800
      TabIndex        =   28
      Top             =   7080
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton grilla 
      Caption         =   "grilla"
      Height          =   255
      Left            =   360
      TabIndex        =   27
      Top             =   7080
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton eliminar 
      Caption         =   "&Eliminar"
      Height          =   375
      Left            =   9000
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   2880
      UseMaskColor    =   -1  'True
      Width           =   975
   End
   Begin VB.CommandButton aceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   10320
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   2880
      UseMaskColor    =   -1  'True
      Width           =   975
   End
   Begin MSMask.MaskEdBox egreso 
      Height          =   375
      Left            =   9960
      TabIndex        =   15
      Top             =   2280
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      Format          =   "$   #,##0.00;($   #,##0.00)"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox ingreso 
      Height          =   375
      Left            =   9960
      TabIndex        =   14
      Top             =   1800
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      Format          =   "$   #,##0.00;($   #,##0.00)"
      PromptChar      =   "_"
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
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
      Left            =   9000
      TabIndex        =   23
      TabStop         =   0   'False
      Text            =   "Egreso:"
      Top             =   2400
      Width           =   855
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
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
      Left            =   9000
      TabIndex        =   22
      TabStop         =   0   'False
      Text            =   "Ingreso:"
      Top             =   1920
      Width           =   855
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
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
      Index           =   2
      Left            =   7200
      TabIndex        =   21
      TabStop         =   0   'False
      Text            =   "C.Costo:"
      Top             =   1440
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
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
      Index           =   1
      Left            =   360
      TabIndex        =   20
      TabStop         =   0   'False
      Text            =   "Detalle:"
      Top             =   1800
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
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
      Index           =   0
      Left            =   360
      TabIndex        =   19
      TabStop         =   0   'False
      Text            =   "Concepto:"
      Top             =   1440
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   0
      Left            =   1440
      MaxLength       =   50
      TabIndex        =   7
      Top             =   1800
      Width           =   5655
   End
   Begin MSDataListLib.DataCombo DataCombo3 
      Bindings        =   "frmcajabanco.frx":0000
      Height          =   315
      Left            =   8280
      TabIndex        =   10
      Top             =   1440
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   556
      _Version        =   393216
      MatchEntry      =   -1  'True
      ListField       =   "descripcion"
      BoundColumn     =   "cc"
      Text            =   ""
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmcajabanco.frx":0017
      Height          =   3495
      Left            =   240
      TabIndex        =   4
      Top             =   3360
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   6165
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
            LCID            =   3082
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
            LCID            =   3082
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
      Height          =   855
      Left            =   240
      TabIndex        =   1
      Top             =   0
      Width           =   11535
      Begin VB.CommandButton Cuenta 
         Caption         =   "Empresa"
         Height          =   255
         Index           =   0
         Left            =   3360
         Picture         =   "frmcajabanco.frx":0034
         TabIndex        =   57
         TabStop         =   0   'False
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   855
      End
      Begin VB.CommandButton Cuenta 
         Caption         =   "Fecha:"
         Height          =   255
         Index           =   1
         Left            =   240
         Picture         =   "frmcajabanco.frx":0566
         TabIndex        =   56
         TabStop         =   0   'False
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Cerrar"
         Height          =   495
         Left            =   10560
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   855
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "frmcajabanco.frx":0A98
         Height          =   315
         Left            =   4200
         TabIndex        =   3
         Top             =   240
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "razonsocial"
         BoundColumn     =   "empresa"
         Text            =   "DataCombo1"
      End
      Begin MSComCtl2.DTPicker fecha 
         Height          =   375
         Left            =   1080
         TabIndex        =   2
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   62390273
         CurrentDate     =   39377
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6135
      Left            =   240
      TabIndex        =   0
      Top             =   960
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   10821
      _Version        =   393216
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "Libro &CAJA"
      TabPicture(0)   =   "frmcajabanco.frx":0AB1
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "DataCombo2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Text2(10)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "DataCombo4"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "DataGrid2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Libro &BANCO"
      TabPicture(1)   =   "frmcajabanco.frx":0ACD
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "DataGrid3"
      Tab(1).Control(1)=   "Text2(11)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Text2(7)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Text2(6)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Text1(1)"
      Tab(1).Control(5)=   "Text2(5)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "femision"
      Tab(1).Control(7)=   "fvenci"
      Tab(1).Control(8)=   "DataCombo5"
      Tab(1).Control(9)=   "DataCombo6"
      Tab(1).ControlCount=   10
      TabCaption(2)   =   "Ci&erre"
      TabPicture(2)   =   "frmcajabanco.frx":0AE9
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame3"
      Tab(2).Control(1)=   "Frame2"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "&Reportes"
      TabPicture(3)   =   "frmcajabanco.frx":0B05
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "reporte"
      Tab(3).Control(1)=   "CrystalReporte"
      Tab(3).Control(2)=   "Frame5"
      Tab(3).Control(3)=   "imprimir"
      Tab(3).Control(4)=   "banco"
      Tab(3).Control(5)=   "Frame4"
      Tab(3).ControlCount=   6
      TabCaption(4)   =   "&Parametros"
      TabPicture(4)   =   "frmcajabanco.frx":0B21
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Combo2"
      Tab(4).Control(1)=   "Combo1"
      Tab(4).Control(2)=   "Text1(3)"
      Tab(4).Control(3)=   "Text1(2)"
      Tab(4).Control(4)=   "Text2(9)"
      Tab(4).Control(4).Enabled=   0   'False
      Tab(4).Control(5)=   "Text2(8)"
      Tab(4).Control(5).Enabled=   0   'False
      Tab(4).ControlCount=   6
      Begin MSDataGridLib.DataGrid DataGrid3 
         Bindings        =   "frmcajabanco.frx":0B3D
         Height          =   495
         Left            =   -74400
         TabIndex        =   55
         Top             =   2160
         Visible         =   0   'False
         Width           =   4335
         _ExtentX        =   7646
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
               LCID            =   1034
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
               LCID            =   1034
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
      Begin MSDataGridLib.DataGrid DataGrid2 
         Bindings        =   "frmcajabanco.frx":0B59
         Height          =   495
         Left            =   480
         TabIndex        =   54
         Top             =   1680
         Visible         =   0   'False
         Width           =   4335
         _ExtentX        =   7646
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
               LCID            =   1034
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
               LCID            =   1034
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
      Begin VB.Frame Frame4 
         Caption         =   "Periodo"
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
         Height          =   1815
         Left            =   -72960
         TabIndex        =   40
         Top             =   960
         Width           =   5655
         Begin VB.CheckBox Check2 
            Caption         =   "Consolida cuentas de Caja"
            Height          =   255
            Left            =   960
            TabIndex        =   58
            Top             =   1440
            Width           =   4095
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Ver Solo Comprobantes cargados por este módulo"
            Height          =   255
            Left            =   960
            TabIndex        =   53
            Top             =   1080
            Width           =   4095
         End
         Begin VB.TextBox Text1 
            BackColor       =   &H8000000F&
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
            Index           =   5
            Left            =   3120
            TabIndex        =   42
            Text            =   "Hasta"
            Top             =   570
            Width           =   495
         End
         Begin VB.TextBox Text1 
            BackColor       =   &H8000000F&
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
            Index           =   6
            Left            =   600
            TabIndex        =   41
            Text            =   "Desde"
            Top             =   570
            Width           =   615
         End
         Begin MSComCtl2.DTPicker cargahasta 
            Height          =   375
            Left            =   3720
            TabIndex        =   43
            Top             =   480
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            _Version        =   393216
            CalendarBackColor=   12648447
            Format          =   62390273
            CurrentDate     =   38415
         End
         Begin MSComCtl2.DTPicker cargadesde 
            Height          =   375
            Left            =   1440
            TabIndex        =   44
            Top             =   480
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            _Version        =   393216
            CalendarBackColor=   12648447
            Format          =   62390273
            CurrentDate     =   38415
         End
      End
      Begin MSDataListLib.DataCombo DataCombo4 
         Bindings        =   "frmcajabanco.frx":0B74
         Height          =   315
         Left            =   1200
         TabIndex        =   8
         Top             =   1200
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "Nombre Cuenta"
         BoundColumn     =   "cuenta"
         Text            =   "DataCombo4"
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
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
         Index           =   11
         Left            =   -74880
         TabIndex        =   52
         TabStop         =   0   'False
         Text            =   "Cuenta:"
         Top             =   1200
         Width           =   975
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
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
         Index           =   10
         Left            =   120
         TabIndex        =   51
         TabStop         =   0   'False
         Text            =   "Cuenta:"
         Top             =   1200
         Width           =   975
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   -69600
         TabIndex        =   50
         Top             =   2040
         Width           =   2175
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   -69600
         TabIndex        =   49
         Top             =   1320
         Width           =   2175
      End
      Begin VB.CommandButton banco 
         Caption         =   "&Ver Libro Banco"
         Height          =   615
         Left            =   -67200
         TabIndex        =   48
         Top             =   1800
         Width           =   1455
      End
      Begin VB.CommandButton imprimir 
         Caption         =   "Ver Libro &Caja"
         Height          =   615
         Left            =   -67200
         TabIndex        =   47
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Frame Frame5 
         Caption         =   "Empresas"
         Height          =   1695
         Left            =   -72960
         TabIndex        =   45
         Top             =   2400
         Visible         =   0   'False
         Width           =   5655
         Begin VB.ListBox List1 
            Height          =   960
            ItemData        =   "frmcajabanco.frx":0B8F
            Left            =   480
            List            =   "frmcajabanco.frx":0B91
            Style           =   1  'Checkbox
            TabIndex        =   46
            Top             =   360
            Width           =   4695
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Banco"
         Height          =   1455
         Left            =   -74760
         TabIndex        =   37
         Top             =   2280
         Width           =   3495
         Begin MSComctlLib.ProgressBar Bar2 
            Height          =   255
            Left            =   240
            TabIndex        =   38
            Top             =   960
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   450
            _Version        =   393216
            Appearance      =   1
            Scrolling       =   1
         End
         Begin VB.CommandButton cerrarbanco 
            Caption         =   "Cerrar Banco"
            Height          =   375
            Left            =   1080
            Style           =   1  'Graphical
            TabIndex        =   39
            Top             =   360
            UseMaskColor    =   -1  'True
            Width           =   1095
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Caja"
         Height          =   1455
         Left            =   -74760
         TabIndex        =   34
         Top             =   600
         Width           =   3495
         Begin MSComctlLib.ProgressBar Bar1 
            Height          =   255
            Left            =   240
            TabIndex        =   36
            Top             =   960
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   450
            _Version        =   393216
            Appearance      =   1
            Scrolling       =   1
         End
         Begin VB.CommandButton cerrarcaja 
            Caption         =   "Cerrar Caja"
            Height          =   375
            Left            =   1080
            Style           =   1  'Graphical
            TabIndex        =   35
            Top             =   360
            UseMaskColor    =   -1  'True
            Width           =   1095
         End
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         DataField       =   "cuentalibrobanco"
         DataSource      =   "datempresa"
         Height          =   285
         Index           =   3
         Left            =   -70560
         MaxLength       =   50
         TabIndex        =   32
         Top             =   2040
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         DataField       =   "cuentalibrocaja"
         DataSource      =   "datempresa"
         Height          =   285
         Index           =   2
         Left            =   -70560
         MaxLength       =   50
         TabIndex        =   31
         Top             =   1320
         Width           =   855
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
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
         Index           =   9
         Left            =   -73200
         TabIndex        =   30
         TabStop         =   0   'False
         Text            =   "Fondo contable libro Banco:"
         Top             =   2040
         Width           =   2535
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
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
         Index           =   8
         Left            =   -73200
         TabIndex        =   29
         TabStop         =   0   'False
         Text            =   "Fondo contable libro caja:"
         Top             =   1320
         Width           =   2535
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
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
         Index           =   7
         Left            =   -68880
         TabIndex        =   26
         TabStop         =   0   'False
         Text            =   "Fecha Venc.:"
         Top             =   1740
         Width           =   1095
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
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
         Index           =   6
         Left            =   -71640
         TabIndex        =   25
         TabStop         =   0   'False
         Text            =   "Fecha Emi.:"
         Top             =   1740
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   1
         Left            =   -73680
         MaxLength       =   30
         TabIndex        =   11
         Top             =   1740
         Width           =   1935
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
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
         Index           =   5
         Left            =   -74880
         TabIndex        =   24
         TabStop         =   0   'False
         Text            =   "Nº Cheque:"
         Top             =   1740
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker femision 
         Height          =   375
         Left            =   -70560
         TabIndex        =   12
         Top             =   1650
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   62390273
         CurrentDate     =   39377
      End
      Begin MSComCtl2.DTPicker fvenci 
         Height          =   375
         Left            =   -67800
         TabIndex        =   13
         Top             =   1650
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   62390273
         CurrentDate     =   39377
      End
      Begin Crystal.CrystalReport CrystalReporte 
         Left            =   -66960
         Top             =   1800
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowTitle     =   "Mayor Analitico"
         PrintFileLinesPerPage=   60
         WindowShowGroupTree=   -1  'True
         WindowShowCloseBtn=   -1  'True
         WindowShowSearchBtn=   -1  'True
         WindowShowPrintSetupBtn=   -1  'True
         WindowShowRefreshBtn=   -1  'True
      End
      Begin MSRDC.MSRDC reporte 
         Height          =   375
         Left            =   -72120
         Top             =   2040
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
      Begin MSDataListLib.DataCombo DataCombo5 
         Bindings        =   "frmcajabanco.frx":0B93
         Height          =   315
         Left            =   -73800
         TabIndex        =   9
         Top             =   1200
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "Nombre Cuenta"
         BoundColumn     =   "cuenta"
         Text            =   "DataCombo4"
      End
      Begin MSDataListLib.DataCombo DataCombo2 
         Bindings        =   "frmcajabanco.frx":0BAF
         Height          =   315
         Left            =   1200
         TabIndex        =   6
         Top             =   480
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         ListField       =   "descripcion"
         BoundColumn     =   "codcontable"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo DataCombo6 
         Bindings        =   "frmcajabanco.frx":0BCA
         Height          =   315
         Left            =   -73800
         TabIndex        =   5
         Top             =   480
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         ListField       =   "descripcion"
         BoundColumn     =   "codcontablebanco"
         Text            =   ""
      End
   End
   Begin vbskpro.Skinner Skinner1 
      Left            =   6600
      Top             =   0
      _ExtentX        =   1270
      _ExtentY        =   1270
      CloseButtonToolTipText=   "Cerrar"
      MinButtonToolTipText=   "Minimizar"
      MaxButtonToolTipText=   "Maximizar"
      OldForeColor    =   0
      RestoreButtonToolTipText=   "Restaurar"
      Enabled         =   0   'False
      ChangeSkinButton=   0   'False
      MinToBarButtonToolTipText=   "Minimizar a la barra de títulos"
      RestoreFromBarButtonToolTipText=   "Restaurar ventana"
      AlwaysOnTopButtonToolTipText=   "Hacer siempre visible"
      AlwaysOnTopDownButtonToolTipText=   "Hacer no siempre visible"
      ChangeSkinButtonToolTipText=   "Cambiar skin"
      HelpButtonToolTipText=   "Ayuda"
      SysEnableSkinCaption=   "Habilitar &Skin"
      SysDisableSkinCaption=   "Deshabilitar &Skin"
      LcK1            =   "3.66*/4/0*/1-5*210/."
      LcK2            =   $"frmcajabanco.frx":0BE5
      AmbientB        =   ";<=>?7B:><7=<A<7CC;@"
      ChSD_FormCaption=   "Seleccione Skin"
      ChSD_ManualSetFrameCaption=   "S&elección manual "
      ChSD_TitleBarSkinComboBoxCaption=   "Skin &barra de Tít."
      ChSD_TitleBarForeColorSetCaption=   "T&exto barra de Tít."
      ChSD_BodySkinComboBoxCaption=   "Skin del cuer&po"
      ChSD_BodyForeColorSetCaption=   "Te&xto del cuerpo"
      ChSD_ChangeForeColorCaption=   "Cambia&r"
      ChSD_SaveToFileCaption=   "&Guardar en un archivo"
      ChSD_LoadFromFileCaption=   "Cargar desde arc&hivo"
      ChSD_UseSkinFileCaption=   "&Usar archivo de skin"
      ChSD_OkCommandButtonCaption=   "&Aceptar"
      ChSD_CancelCommandButtonCaption=   "&Cancelar"
   End
   Begin MSAdodcLib.Adodc datempresa 
      Height          =   330
      Left            =   0
      Top             =   480
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
      DataSourceName  =   ""
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
   Begin MSAdodcLib.Adodc datconceptos 
      Height          =   330
      Left            =   1200
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   4
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
      Left            =   2400
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
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
   Begin MSAdodcLib.Adodc datcosto 
      Height          =   450
      Left            =   0
      Top             =   240
      Visible         =   0   'False
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   794
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
      DataSourceName  =   ""
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
   Begin MSAdodcLib.Adodc datmaestro 
      Height          =   330
      Left            =   5040
      Top             =   6960
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
   Begin MSAdodcLib.Adodc datasiento 
      Height          =   330
      Left            =   6720
      Top             =   6360
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
      UserName        =   "lucva"
      Password        =   "25072004"
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
   Begin MSAdodcLib.Adodc datfiltro 
      Height          =   330
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
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
   Begin MSAdodcLib.Adodc datfondo 
      Height          =   330
      Left            =   6360
      Top             =   6960
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
   Begin MSAdodcLib.Adodc datfondocaja 
      Height          =   330
      Left            =   7560
      Top             =   6960
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
   Begin MSAdodcLib.Adodc datfondobanco 
      Height          =   330
      Left            =   8760
      Top             =   6960
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
   Begin MSAdodcLib.Adodc criterio 
      Height          =   330
      Left            =   0
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
   Begin MSAdodcLib.Adodc datlibrodiario 
      Height          =   330
      Left            =   0
      Top             =   0
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
Attribute VB_Name = "frmcajabanco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cc As Integer
Dim empre(100) As Integer
Dim fondo(100) As Integer
Dim cuentaimputa As Integer


Private Sub aceptar_Click()
On Error GoTo fuera

    datmovimientos.Recordset.AddNew
    
    If DataCombo2.Text = "" And SSTab1.Tab = 0 Then
         MsgBox "Debe ingresar un concepto", vbCritical, "Error"
         DataCombo2.SetFocus
    End If
    If DataCombo6.Text = "" And SSTab1.Tab = 1 Then
         MsgBox "Debe ingresar un concepto", vbCritical, "Error"
         DataCombo6.SetFocus
    End If
    If Text1(0).Text = "" Then
         MsgBox "Debe ingresar el Detalle del movimiento", vbCritical, "Error"
         Text1(0).SetFocus
    End If
    If DataCombo3.Text = "" Then
        cc = 0
    Else
        cc = Val(DataCombo3.BoundText)
    End If
    If Val(ingreso.Text) < 0 Then
         MsgBox "Valor incorrecto en Ingreso", vbCritical, "Error"
         ingreso.SetFocus
    End If
    If Val(egreso.Text) < 0 Then
         MsgBox "Valor incorrecto en Egreso", vbCritical, "Error"
         ingreso.SetFocus
    End If
    If Val(ingreso.Text) = 0 And Val(egreso.Text) = 0 Then
         MsgBox "Debe ingresar un importe en Ingreso o Egreso", vbCritical, "Error"
         ingreso.SetFocus
    End If
    
    If SSTab1.Tab = 1 Then
        If Text1(1).Text = "" Then
         MsgBox "Debe ingresar Nº de cheque", vbCritical, "Error"
         Text1(1).SetFocus
        End If
    End If



    datmovimientos.Recordset.Fields("fecha") = fecha.Value
    If SSTab1.Tab = 0 Then
        datmovimientos.Recordset.Fields("concepto") = DataCombo2.Text
    Else
        datmovimientos.Recordset.Fields("concepto") = DataCombo6.Text
    End If
    datmovimientos.Recordset.Fields("detalle") = Text1(0).Text
    datmovimientos.Recordset.Fields("ccosto") = cc
    datmovimientos.Recordset.Fields("ingreso") = Val(ingreso.Text)
    datmovimientos.Recordset.Fields("egreso") = Val(egreso.Text)
    datmovimientos.Recordset.Fields("ingreso") = Val(ingreso.Text)
    datmovimientos.Recordset.Fields("cerrado") = "N"
    datmovimientos.Recordset.Fields("empresa") = DataCombo1.BoundText
    If SSTab1.Tab = 0 Then
        datmovimientos.Recordset.Fields("cajabanco") = "C"
        datmovimientos.Recordset.Fields("detallecuenta") = DataCombo4.Text
        If DataCombo4.BoundText = "" Then
            MsgBox "No ingreso la cuenta de Caja", vbCritical, "Error"
            DataCombo4.SetFocus
            Exit Sub
        End If
        cuentaimputa = DataCombo4.BoundText
        
    Else
        datmovimientos.Recordset.Fields("nrocheque") = Text1(1).Text
        datmovimientos.Recordset.Fields("fechacheque") = femision.Value
        datmovimientos.Recordset.Fields("fechavencimiento") = fvenci.Value
        datmovimientos.Recordset.Fields("detalle") = Text1(0).Text
        datmovimientos.Recordset.Fields("cajabanco") = "B"
        datmovimientos.Recordset.Fields("detallecuenta") = DataCombo5.Text
        If DataCombo5.BoundText = "" Then
            MsgBox "No ingreso la cuenta de Banco", vbCritical, "Error"
            DataCombo5.SetFocus
            Exit Sub
        End If
        cuentaimputa = DataCombo5.BoundText
    End If
    If DataCombo2.BoundText = "" And SSTab1.Tab = 0 Then
         MsgBox "El concepto no posee codigo contable asociado", vbCritical, "Error"
         DataCombo2.SetFocus
         Exit Sub
    End If
    If DataCombo6.BoundText = "" And SSTab1.Tab = 1 Then
         MsgBox "El concepto no posee codigo contable asociado", vbCritical, "Error"
         DataCombo6.SetFocus
         Exit Sub
    End If
    If SSTab1.Tab = 0 Then datmovimientos.Recordset.Fields("codcontable") = DataCombo2.BoundText
    If SSTab1.Tab = 1 Then datmovimientos.Recordset.Fields("codcontable") = DataCombo6.BoundText
    datmovimientos.Recordset.Fields("cuenta") = cuentaimputa
    datmovimientos.Recordset.UpdateBatch adAffectCurrent

    Call grabaasiento_Click
    Call limpia_Click
    
    Exit Sub
fuera:
    MsgBox "Conceptos mal ingresados, no se pudo grabar", vbCritical, "Error"
    

End Sub

Private Sub banco_Click()
Dim tabla As String
Dim ruta As String


ruta = "\Empresa" + Right(Str(login.empresaact), Len(Str(login.empresaact)) - 1)
    
If Check1.Value = 1 Then

 datfiltro.RecordSource = "select filtracaja.* from filtracaja"
datfiltro.Refresh
If datfiltro.Recordset.EOF = False Then
    datfiltro.Recordset.MoveFirst
    Do While Not datfiltro.Recordset.EOF
        datfiltro.Recordset.Delete adAffectCurrent
        datfiltro.Recordset.MoveNext
    Loop
End If

For x = 0 To List1.ListCount - 1
    If List1.Selected(x) = True Then
        datfiltro.Recordset.AddNew
        datfiltro.Recordset.Fields(0) = empre(x)
        datfiltro.Recordset.Fields(1) = cargadesde.Value
        datfiltro.Recordset.UpdateBatch adAffectCurrent
    End If
Next x
    

    reporte.SQL = "SELECT librocajabanco_c.id, librocajabanco_c.fecha, librocajabanco_c.concepto, librocajabanco_c.detalle, librocajabanco_c.ccosto, librocajabanco_c.ingreso, librocajabanco_c.egreso, librocajabanco_c.empresa, librocajabanco_c.ingresoant, librocajabanco_c.egresoant FROM contablesql.dbo.librocajabanco_c librocajabanco_c where cuenta = '" & DataCombo5.BoundText & "' and fecha >= '" & cargadesde.Value & "' and fecha <= '" & cargahasta.Value & "' and cajabanco = 'B'  order by fecha,id "

tabla = reporte.SQL

With CrystalReporte
    .ReportFileName = App.Path & ruta + "\librobanco.rpt"
 Rem   .Formulas(0) = "fecha1=""" & cargadesde.Value & """"
 Rem   .Formulas(1) = "fecha2=""" & cargahasta.Value & """"
 Rem   .Formulas(2) = "fondo=""" & DataCombo1.Text & """"
 Rem   .Formulas(3) = "arqueo=""" & arqueo.Text & """"
    .Connect = login.conexionreporte
    .DiscardSavedData = True
    .RetrieveDataFiles
    .ReportSource = 0
    .SQLQuery = tabla
    .Destination = crptToWindow
    .PrintFileType = crptCrystal
    .WindowState = crptMaximized
    .Action = 1
    
End With

Else

    criterio.ConnectionString = login.conexiontotal
    criterio.RecordSource = "select empreactiva.* from empreactiva"
    criterio.Refresh

    criterio.Recordset.Fields(0) = login.empresaact
    criterio.Recordset.Fields(1) = cargadesde.Value
    criterio.Recordset.Fields(4) = DataGrid3.Columns(6).Text
    criterio.Recordset.Fields(5) = DataGrid3.Columns(6).Text
    criterio.Recordset.Fields(6) = login.iper
    criterio.Recordset.UpdateBatch adAffectCurrent
    criterio.Refresh

    datlibrodiario.RecordSource = "select librodiario1.* from librodiario1 where librodiario1.idcuenta = '" & DataCombo5.BoundText & "' and librodiario1.empresa = " & login.empresaact & ""
    datlibrodiario.Refresh
        
If datlibrodiario.Recordset.EOF = False Then
    reporte.SQL = "SELECT  libroca_caja.Fecha, libroca_caja.concepto, libroca_caja.detalle, libroca_caja.ccosto, libroca_caja.ingreso, libroca_caja.egreso, libroca_caja.empresa, librodiario1.debe, librodiario1.haber FROM { oj contablesql.dbo.libro_caja libroca_caja INNER JOIN contablesql.dbo.librodiario1 librodiario1 ON libroca_caja.cuenta = librodiario1.idcuenta AND libroca_caja.empresa = librodiario1.empresa} WHERE libroca_caja.cuenta = '" & DataCombo5.BoundText & "' and libroca_caja.empresa = " & login.empresaact & " and libroca_caja.fecha >= '" & cargadesde.Value & "' and libroca_caja.fecha <= '" & cargahasta.Value & "' ORDER BY libroca_caja.fecha ASC, libroca_caja.idasiento ASC"
Else
    reporte.SQL = "SELECT libroca_caja.Fecha, libroca_caja.concepto, libroca_caja.detalle, libroca_caja.ccosto, libroca_caja.ingreso, libroca_caja.egreso, libroca_caja.empresa FROM contablesql.dbo.libro_caja libroca_caja WHERE libroca_caja.cuenta = '" & DataCombo5.BoundText & "' and libroca_caja.empresa = " & login.empresaact & " and libroca_caja.fecha >= '" & cargadesde.Value & "' and libroca_caja.fecha <= '" & cargahasta.Value & "' ORDER BY libroca_caja.fecha ASC, libroca_caja.idasiento ASC"
End If
tabla = reporte.SQL

With CrystalReporte
If datlibrodiario.Recordset.EOF = False Then
    .ReportFileName = App.Path & ruta + "\librobanco_mayor.rpt"
Else
    .ReportFileName = App.Path & ruta + "\librobanco_mayor2.rpt"
End If
    .Connect = login.conexionreporte
    .DiscardSavedData = True
    .RetrieveDataFiles
    .ReportSource = 0
    .SQLQuery = tabla
    .Destination = crptToWindow
    .PrintFileType = crptCrystal
    .WindowState = crptMaximized
    .Action = 1
    
End With


End If

End Sub

Private Sub cargadesde_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = 69 And Frame5.Visible = False Then
        Frame5.Visible = True
    End If

End Sub

Private Sub cerrarbanco_Click()

        datmovimientos.RecordSource = "select librocajabanco.* from librocajabanco where fecha = '" & fecha.Value & "' and empresa = " & DataCombo1.BoundText & " AND cajabanco = 'B'"
        datmovimientos.Refresh
        bar2.Min = 0
        
        If datmovimientos.Recordset.EOF = True Then Exit Sub
        bar2.max = datmovimientos.Recordset.RecordCount
        datmovimientos.Recordset.MoveFirst
        
        Do While Not datmovimientos.Recordset.EOF
            datmovimientos.Recordset.Fields("cerrado") = "S"
            bar2.Value = datmovimientos.Recordset.AbsolutePosition
            datmovimientos.Recordset.MoveNext
        Loop
        bar2.Value = 0
        MsgBox "Banco dia:" + Str(fecha.Value) + " Cerrado", vbInformation, "Proceso"
        

End Sub

Private Sub cerrarcaja_Click()

        datmovimientos.RecordSource = "select librocajabanco.* from librocajabanco where fecha = '" & fecha.Value & "' and empresa = " & DataCombo1.BoundText & " AND cajabanco = 'C'"
        datmovimientos.Refresh
        bar1.Min = 0
        
        If datmovimientos.Recordset.EOF = True Then Exit Sub
        bar1.max = datmovimientos.Recordset.RecordCount
        datmovimientos.Recordset.MoveFirst
        
        Do While Not datmovimientos.Recordset.EOF
            datmovimientos.Recordset.Fields("cerrado") = "S"
            bar1.Value = datmovimientos.Recordset.AbsolutePosition
            datmovimientos.Recordset.MoveNext
        Loop
        bar1.Value = 0
        MsgBox "Caja dia:" + Str(fecha.Value) + " Cerrada", vbInformation, "Proceso"
        
        

End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        Text1(2).Text = fondo(Combo1.ListIndex)
        datempresa.Recordset.UpdateBatch adAffectCurrent
    End If

End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        Text1(3).Text = fondo(Combo2.ListIndex)
        datempresa.Recordset.UpdateBatch adAffectCurrent
    End If

End Sub

Private Sub Command1_Click()

    Unload Me

End Sub

Private Sub DataCombo1_KeyPress(KeyAscii As Integer)

On Error GoTo fuera

    If KeyAscii = 13 Then
        KeyAscii = 0
        If IsNull(DataCombo1.Text) = True Then Exit Sub

        login.datusyempresas.RecordSource = "select usuarioyempresa.* from usuarioyempresa WHERE usuarioyempresa.nomusuario = '" & login.usuarioactivo & "' and empresa = " & DataCombo1.BoundText & ""
        login.datusyempresas.Refresh
        If login.datusyempresas.Recordset.EOF = True Then
            mensa = MsgBox("Permiso denegado a esta empresa", vbCritical, "Error")
            DataCombo1.Text = login.nomempresa
            Exit Sub
        End If
        login.empresaact = DataCombo1.BoundText
        login.nomempresa = DataCombo1.Text
        
        
        datempresa.RecordSource = "select empresa.* from empresa"
        datempresa.Refresh

        datempresa.RecordSource = "select empresa.* from empresa where empresa = " & login.empresaact & " "
        datempresa.Refresh
        login.iper = datempresa.Recordset.Fields("inicioperiodo")
        login.fper = datempresa.Recordset.Fields("finperiodo")
    
        Unload Me
        frmcajabanco.Show
    End If
fuera:

End Sub

Private Sub DataCombo2_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{tab}", False
    End If
    

End Sub

Private Sub DataCombo3_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        If SSTab1.Tab = 1 Then
            SendKeys "{tab}", False
        Else
            ingreso.SetFocus
        End If
    End If

End Sub

Private Sub DataCombo4_Click(Area As Integer)
On Error Resume Next

    datmovimientos.RecordSource = "select librocajabanco.* from librocajabanco where fecha = '" & fecha.Value & "' and empresa = " & DataCombo1.BoundText & " AND cajabanco = 'C' and cuenta = " & DataCombo4.BoundText & " "
    datmovimientos.Refresh
    DataGrid2.Bookmark = DataCombo4.SelectedItem
    Call grilla_Click
    
End Sub

Private Sub DataCombo4_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = 13 Then
        KeyAscii = 0
        DataGrid2.Bookmark = DataCombo4.SelectedItem
        DataCombo3.SetFocus
    End If

End Sub

Private Sub DataCombo4_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next

    datmovimientos.RecordSource = "select librocajabanco.* from librocajabanco where fecha = '" & fecha.Value & "' and empresa = " & DataCombo1.BoundText & " AND cajabanco = 'C' and cuenta = " & DataCombo4.BoundText & " "
    datmovimientos.Refresh
    DataGrid2.Bookmark = DataCombo4.SelectedItem
    Call grilla_Click
    
    
End Sub

Private Sub DataCombo5_Click(Area As Integer)
On Error Resume Next

        datmovimientos.RecordSource = "select librocajabanco.* from librocajabanco where fecha = '" & fecha.Value & "' and empresa = " & DataCombo1.BoundText & " AND cajabanco = 'B' and cuenta = " & DataCombo5.BoundText & " "
        datmovimientos.Refresh
        cuentaimputa = DataCombo5.BoundText
        DataGrid3.Bookmark = DataCombo5.SelectedItem
        Call grilla_Click
End Sub

Private Sub DataCombo5_KeyPress(KeyAscii As Integer)
On Error Resume Next

    If KeyAscii = 13 Then
        KeyAscii = 0
        DataGrid3.Bookmark = DataCombo5.SelectedItem
        DataCombo3.SetFocus
    End If

End Sub

Private Sub DataCombo5_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next

        datmovimientos.RecordSource = "select librocajabanco.* from librocajabanco where fecha = '" & fecha.Value & "' and empresa = " & DataCombo1.BoundText & " AND cajabanco = 'B' and cuenta = " & DataCombo5.BoundText & " "
        datmovimientos.Refresh
        cuentaimputa = DataCombo5.BoundText
        DataGrid3.Bookmark = DataCombo5.SelectedItem
        Call grilla_Click
        
End Sub

Private Sub DataCombo6_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        Text1(0).SetFocus
    End If

End Sub

Private Sub egreso_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        If Val(egreso.Text) > 0 Then ingreso.Text = 0
           
        SendKeys "{tab}", False
    End If


End Sub

Private Sub egreso_LostFocus()

        If Val(egreso.Text) > 0 Then ingreso.Text = 0

End Sub

Private Sub eliminar_Click()
On Error Resume Next
    mensa = MsgBox("Esta seguro de eliminar este movimiento", vbYesNo, "Atención")
    If mensa = vbYes Then
        Rem ************* borra asiento ****************
        filtroasiento = datmovimientos.Recordset.Fields("idmasterasientos")
        If IsNull(filtroasiento) = True Then GoTo paso2
        datmaestro.RecordSource = "select [Maestro Asientos].* from [Maestro Asientos] where empresa = " & login.empresaact & " and idmasterasientos = " & filtroasiento & ""
        datmaestro.Refresh
        datasiento.RecordSource = "select [Detalle Asientos].* from [Detalle Asientos] where empresa = " & login.empresaact & " and idmasterasientos = " & filtroasiento & ""
        datasiento.Refresh
        If datasiento.Recordset.EOF = True Then GoTo paso1
        datasiento.Recordset.MoveFirst
        
        Do While Not datasiento.Recordset.EOF
                datasiento.Recordset.Delete adAffectCurrent
                datasiento.Recordset.MoveNext
        Loop
paso1:
        datmaestro.Recordset.Delete adAffectCurrent
paso2:
        Rem   fin borrado de asiento  ******************
                
    
        datmovimientos.Recordset.Delete adAffectCurrent
        datmovimientos.Refresh
        
        
        
        
        
        Call SSTab1_Click(SSTab1.Tab)
    End If


End Sub

Private Sub fecha_Change()

        If fecha.Value > Date Then
            MsgBox "No puede ingresar movimientos con esta fecha", vbCritical, "Error"
            fecha.Value = Date
            Exit Sub
        End If
        If fecha.Value < login.iper Then
            mensa = MsgBox("La Fecha es erronea o no pertenecia al periodo en ejercicio", vbCritical, "!! Atención !!")
            fecha.Value = Date
            Exit Sub
        End If
            
        cargadesde.Value = fecha.Value
        cargahasta.Value = fecha.Value
        datmovimientos.RecordSource = "select librocajabanco.* from librocajabanco where fecha = '" & fecha.Value & "' and empresa = " & DataCombo1.BoundText & " AND cajabanco = 'C'"
        datmovimientos.Refresh
        If login.librocajamodif = "N" Then
            eliminar.Enabled = False
            aceptar.Enabled = False
        Else
            eliminar.Enabled = True
            aceptar.Enabled = True
        End If
        If datmovimientos.Recordset.EOF = False Then
            datmovimientos.Recordset.MoveFirst
            If datmovimientos.Recordset.Fields("cerrado") = "S" Then
                eliminar.Enabled = False
                aceptar.Enabled = False
            End If
        End If
        Call grilla_Click
        SSTab1.Tab = 0
        
        
End Sub

Private Sub Form_Load()
On Error Resume Next
Aplicar_skin Me

fecha.Value = Date
cargadesde.Value = Date
cargahasta.Value = Date
fvenci.Value = Date
femision.Value = Date
Check1.Value = 1
Check2.Value = 0

If login.librocajamodif = "N" Then
    eliminar.Enabled = False
    aceptar.Enabled = False
Else
    eliminar.Enabled = True
    aceptar.Enabled = True
End If

If login.librocajalistar = "N" Then
    imprimir.Enabled = False
    banco.Enabled = False
Else
    imprimir.Enabled = True
    banco.Enabled = True
End If

datempresa.ConnectionString = login.conexiontotal
datconceptos.ConnectionString = login.conexiontotal
datmovimientos.ConnectionString = login.conexiontotal
datcosto.ConnectionString = login.conexiontotal
datasiento.ConnectionString = login.conexiontotal
datmaestro.ConnectionString = login.conexiontotal
datfiltro.ConnectionString = login.conexiontotal
datfondo.ConnectionString = login.conexiontotal
datfondocaja.ConnectionString = login.conexiontotal
datfondobanco.ConnectionString = login.conexiontotal
datlibrodiario.ConnectionString = login.conexiontotal
    
    If login.administrador = "N" Then
        SSTab1.TabEnabled(4) = False
    Else
       SSTab1.TabEnabled(4) = True
    End If
    
Frame5.Visible = False
datempresa.RecordSource = "select usuarioyempresa.* from usuarioyempresa where nomusuario = '" & login.usuarioactivo & "'"
datempresa.Refresh

datempresa.Recordset.MoveFirst
i = 0
Do While Not datempresa.Recordset.EOF
    List1.AddItem datempresa.Recordset.Fields("razonsocial")
    empre(i) = datempresa.Recordset.Fields("empresa")
    datempresa.Recordset.MoveNext
    i = i + 1
Loop


datempresa.RecordSource = "select empresa.* from empresa order by empresa"
datempresa.Refresh
datconceptos.RecordSource = "select conceptoscaja.* from conceptoscaja order by descripcion"
datconceptos.Refresh

datfondo.RecordSource = "select fondos.* from fondos where empresa = " & login.empresaact & " order by id"
datfondo.Refresh

i = 0
cod = 0
If datfondo.Recordset.EOF = False Then
Do While Not datfondo.Recordset.EOF
    If cod <> datfondo.Recordset.Fields("id") Then
        Combo1.AddItem datfondo.Recordset.Fields("nombrefondo")
        Combo2.AddItem datfondo.Recordset.Fields("nombrefondo")
        fondo(i) = datfondo.Recordset.Fields("id")
        i = i + 1
    End If
    cod = datfondo.Recordset.Fields("id")
    datfondo.Recordset.MoveNext
Loop
    
End If

datfondocaja.RecordSource = "select fondos.* from fondos where empresa = " & login.empresaact & " and id = " & datempresa.Recordset.Fields("cuentalibrocaja") & " and prorrateo > 0 "
datfondocaja.Refresh

If datfondocaja.Recordset.EOF = False Then
    datfondocaja.Recordset.MoveFirst
    DataCombo4.Text = datfondocaja.Recordset.Fields(3)
    DataCombo4.BoundText = datfondocaja.Recordset.Fields(1)
End If
    
datfondobanco.RecordSource = "select fondos.* from fondos where empresa = " & login.empresaact & " and id = " & datempresa.Recordset.Fields("cuentalibrobanco") & " and prorrateo > 0 "
datfondobanco.Refresh
If datfondobanco.Recordset.EOF = False Then
    datfondobanco.Recordset.MoveFirst
    DataCombo5.Text = datfondobanco.Recordset.Fields(3)
    DataCombo5.BoundText = datfondobanco.Recordset.Fields(1)
End If



DataCombo1.Text = login.razonsoc

For Y = 0 To i - 1
    If empre(Y) = DataCombo1.BoundText Then List1.Selected(Y) = True
Next Y

datmovimientos.RecordSource = "select librocajabanco.* from librocajabanco where fecha = '" & fecha.Value & "' and empresa = " & DataCombo1.BoundText & " AND cajabanco = 'C'"
datmovimientos.Refresh

datcosto.RecordSource = "select ccostos.* from ccostos where empresa = " & DataCombo1.BoundText & " "
datcosto.Refresh

    datmaestro.RecordSource = "select [Maestro Asientos].* from [Maestro Asientos] where empresa = " & login.empresaact & " and perinicial = '" & login.iper & "' order by nroasiento"
    datmaestro.Refresh
    datasiento.RecordSource = "select [Detalle Asientos].* from [Detalle Asientos] where empresa = " & login.empresaact & " order by idmasterasientos"
    datasiento.Refresh

Call grilla_Click

SSTab1.Tab = 0


End Sub

Private Sub grabaasiento_Click()
            
     datmaestro.RecordSource = "select [Maestro Asientos].* from [Maestro Asientos] where empresa = " & DataCombo1.BoundText & " and perinicial = '" & login.iper & "' order by nroasiento"
     datmaestro.Refresh
    
    If datmaestro.Recordset.EOF = False Then

            datmaestro.Recordset.MoveLast
            nroasie = datmaestro.Recordset.Fields(3) + 1
    Else
            nroasie = 1
    End If
         
    datmaestro.Recordset.AddNew
    datmaestro.Recordset.Fields(0) = fecha.Value
    datmaestro.Recordset.Fields(1) = Date
    datmaestro.Recordset.Fields(3) = nroasie
    If SSTab1.Tab = 0 Then
         datmaestro.Recordset.Fields(4) = DataCombo2.Text
    Else
         datmaestro.Recordset.Fields(4) = DataCombo6.Text
    End If
    datmaestro.Recordset.Fields(5) = login.iper
    datmaestro.Recordset.Fields(6) = login.fper
    datmaestro.Recordset.Fields(7) = DataCombo1.BoundText
    datmaestro.Recordset.Fields(8) = "N"
    datmaestro.Recordset.Fields(10) = "L"
    datmaestro.Recordset.Fields(11) = "S"
    datmaestro.Recordset.UpdateBatch adAffectCurrent

            datasiento.Recordset.AddNew
            datasiento.Recordset.Fields(1) = datmaestro.Recordset.Fields(0).Value
            If SSTab1.Tab = 0 Then
                datasiento.Recordset.Fields(2) = DataCombo2.BoundText
            Else
                datasiento.Recordset.Fields(2) = DataCombo6.BoundText
            End If
            datasiento.Recordset.Fields(3) = Val(egreso.Text)
            datasiento.Recordset.Fields(4) = Val(ingreso.Text)
            datasiento.Recordset.Fields(5) = Val(datmaestro.Recordset.Fields(2))
            datasiento.Recordset.Fields(6) = Text1(0).Text
            datasiento.Recordset.Fields(7) = DataCombo1.BoundText
            datasiento.Recordset.Fields(8) = cc
            datasiento.Recordset.UpdateBatch adAffectCurrent
            
            datasiento.Recordset.AddNew
            datasiento.Recordset.Fields(1) = datmaestro.Recordset.Fields(0).Value
            datasiento.Recordset.Fields(2) = cuentaimputa
            If Val(egreso.Text) = 0 Then datasiento.Recordset.Fields(3) = Val(ingreso.Text)
            If Val(ingreso.Text) = 0 Then datasiento.Recordset.Fields(4) = Val(egreso.Text)
            datasiento.Recordset.Fields(5) = Val(datmaestro.Recordset.Fields(2))
            datasiento.Recordset.Fields(6) = Text1(0).Text
            datasiento.Recordset.Fields(7) = DataCombo1.BoundText
            datasiento.Recordset.Fields(8) = cc
            datasiento.Recordset.UpdateBatch adAffectCurrent
     datmovimientos.Recordset.Fields("idmasterasientos") = datmaestro.Recordset.Fields("idmasterasientos")
     datmovimientos.Recordset.UpdateBatch adAffectCurrent

End Sub

Private Sub grilla_Click()
DataGrid1.Columns(0).Visible = False
DataGrid1.Columns(0).Width = 0
DataGrid1.Columns(1).Visible = False
DataGrid1.Columns(4).Visible = False
DataGrid1.Columns(5).Visible = False
DataGrid1.Columns(6).Visible = False
DataGrid1.Columns(10).Visible = False
DataGrid1.Columns(11).Visible = False
DataGrid1.Columns(12).Visible = False
DataGrid1.Columns(13).Visible = False
DataGrid1.Columns(14).Visible = False
DataGrid1.Columns(8).Alignment = dbgRight
DataGrid1.Columns(9).Alignment = dbgRight
DataGrid1.Columns(8).NumberFormat = "#,##0.00"
DataGrid1.Columns(9).NumberFormat = "#,##0.00"
DataGrid1.Columns(2).Width = 2500
DataGrid1.Columns(3).Width = 2500
DataGrid1.Columns(4).Width = 1000
DataGrid1.Columns(5).Width = 1000
DataGrid1.Columns(6).Width = 1000
DataGrid1.Columns(7).Width = 400
DataGrid1.Columns(8).Width = 1000
DataGrid1.Columns(9).Width = 1000
End Sub

Private Sub imprimir_Click()
On Error Resume Next
Dim tabla As String
Dim ruta As String


ruta = "\Empresa" + Right(Str(login.empresaact), Len(Str(login.empresaact)) - 1)

If Check1.Value = 1 Then


 datfiltro.RecordSource = "select filtracaja.* from filtracaja"
datfiltro.Refresh
If datfiltro.Recordset.EOF = False Then
    datfiltro.Recordset.MoveFirst
    Do While Not datfiltro.Recordset.EOF
        datfiltro.Recordset.Delete adAffectCurrent
        datfiltro.Recordset.MoveNext
    Loop
End If

For x = 0 To List1.ListCount - 1
    If List1.Selected(x) = True Then
        datfiltro.Recordset.AddNew
        datfiltro.Recordset.Fields(0) = empre(x)
        datfiltro.Recordset.Fields(1) = cargadesde.Value
        datfiltro.Recordset.UpdateBatch adAffectCurrent
    End If
Next x
    
If Check2.Value = 0 Then
    reporte.SQL = "SELECT librocajabanco_c.id, librocajabanco_c.fecha, librocajabanco_c.concepto, librocajabanco_c.detalle, librocajabanco_c.ccosto, librocajabanco_c.ingreso, librocajabanco_c.egreso, librocajabanco_c.empresa, librocajabanco_c.ingresoant, librocajabanco_c.egresoant FROM contablesql.dbo.librocajabanco_c librocajabanco_c where cuenta = '" & DataCombo4.BoundText & "' and  fecha >= '" & cargadesde.Value & "' and fecha <= '" & cargahasta.Value & "' and cajabanco = 'C' order by fecha,id "
Else
    reporte.SQL = "SELECT librocajabanco_c.id, librocajabanco_c.fecha, librocajabanco_c.concepto, librocajabanco_c.detalle, librocajabanco_c.ccosto, librocajabanco_c.ingreso, librocajabanco_c.egreso, librocajabanco_c.empresa, librocajabanco_c.ingresoant, librocajabanco_c.egresoant FROM contablesql.dbo.librocajabanco_c librocajabanco_c where fecha >= '" & cargadesde.Value & "' and fecha <= '" & cargahasta.Value & "' and cajabanco = 'C' order by fecha,id "
End If
tabla = reporte.SQL

With CrystalReporte
If Check2.Value = 0 Then
    .ReportFileName = App.Path & ruta + "\librocaja.rpt"
Else
    .ReportFileName = App.Path & ruta + "\librocaja_c.rpt"
End If
 Rem   .Formulas(0) = "fecha1=""" & cargadesde.Value & """"
 Rem   .Formulas(1) = "fecha2=""" & cargahasta.Value & """"
 Rem   .Formulas(2) = "fondo=""" & DataCombo1.Text & """"
 Rem   .Formulas(3) = "arqueo=""" & arqueo.Text & """"
    .Connect = login.conexionreporte
    .DiscardSavedData = True
    .RetrieveDataFiles
    .ReportSource = 0
    .SQLQuery = tabla
    .Destination = crptToWindow
    .PrintFileType = crptCrystal
    .WindowState = crptMaximized
    .Action = 1
    
End With

Else

    criterio.ConnectionString = login.conexiontotal
    criterio.RecordSource = "select empreactiva.* from empreactiva"
    criterio.Refresh

    

    criterio.Recordset.Fields(0) = login.empresaact
    criterio.Recordset.Fields(1) = cargadesde.Value
    criterio.Recordset.Fields(4) = DataGrid2.Columns(6).Text
    criterio.Recordset.Fields(5) = DataGrid2.Columns(6).Text
    criterio.Recordset.Fields(6) = login.iper
    criterio.Recordset.UpdateBatch adAffectCurrent
    criterio.Refresh

If Check2.Value = 0 Then
    datlibrodiario.RecordSource = "select librodiario1.* from librodiario1 where librodiario1.idcuenta = '" & DataCombo4.BoundText & "' and librodiario1.empresa = " & login.empresaact & ""
    datlibrodiario.Refresh
Else
    datlibrodiario.RecordSource = "select librodiario1.* from librodiario1 where librodiario1.empresa = " & login.empresaact & ""
    datlibrodiario.Refresh
End If
        
If datlibrodiario.Recordset.EOF = False Then
 If Check2.Value = 0 Then
    reporte.SQL = "SELECT  libroca_caja.Fecha, libroca_caja.concepto, libroca_caja.detalle, libroca_caja.ccosto, libroca_caja.ingreso, libroca_caja.egreso, libroca_caja.empresa, librodiario1.debe, librodiario1.haber FROM { oj contablesql.dbo.libro_caja libroca_caja INNER JOIN contablesql.dbo.librodiario1 librodiario1 ON libroca_caja.cuenta = librodiario1.idcuenta AND libroca_caja.empresa = librodiario1.empresa} WHERE libroca_caja.cuenta = '" & DataCombo4.BoundText & "' and libroca_caja.empresa = " & login.empresaact & " and libroca_caja.fecha >= '" & cargadesde.Value & "' and libroca_caja.fecha <= '" & cargahasta.Value & "' ORDER BY libroca_caja.fecha ASC, libroca_caja.idasiento ASC"
 Else
    reporte.SQL = "SELECT  libroca_caja.Fecha, libroca_caja.concepto, libroca_caja.detalle, libroca_caja.ccosto, libroca_caja.ingreso, libroca_caja.egreso, libroca_caja.empresa, librodiario1.debe, librodiario1.haber FROM { oj contablesql.dbo.libro_caja libroca_caja INNER JOIN contablesql.dbo.librodiario1 librodiario1 ON libroca_caja.cuenta = librodiario1.idcuenta AND libroca_caja.empresa = librodiario1.empresa} WHERE libroca_caja.empresa = " & login.empresaact & " and libroca_caja.fecha >= '" & cargadesde.Value & "' and libroca_caja.fecha <= '" & cargahasta.Value & "' ORDER BY libroca_caja.fecha ASC, libroca_caja.idasiento ASC"
 End If
Else
 If Check2.Value = 0 Then
    reporte.SQL = "SELECT libroca_caja.Fecha, libroca_caja.concepto, libroca_caja.detalle, libroca_caja.ccosto, libroca_caja.ingreso, libroca_caja.egreso, libroca_caja.empresa FROM contablesql.dbo.libro_caja libroca_caja WHERE libroca_caja.cuenta = '" & DataCombo4.BoundText & "' and libroca_caja.empresa = " & login.empresaact & " and libroca_caja.fecha >= '" & cargadesde.Value & "' and libroca_caja.fecha <= '" & cargahasta.Value & "' ORDER BY libroca_caja.fecha ASC, libroca_caja.idasiento ASC"
 Else
    reporte.SQL = "SELECT libroca_caja.Fecha, libroca_caja.concepto, libroca_caja.detalle, libroca_caja.ccosto, libroca_caja.ingreso, libroca_caja.egreso, libroca_caja.empresa FROM contablesql.dbo.libro_caja libroca_caja WHERE libroca_caja.empresa = " & login.empresaact & " and libroca_caja.fecha >= '" & cargadesde.Value & "' and libroca_caja.fecha <= '" & cargahasta.Value & "' ORDER BY libroca_caja.fecha ASC, libroca_caja.idasiento ASC"
 End If
End If
tabla = reporte.SQL

With CrystalReporte
If datlibrodiario.Recordset.EOF = False Then
  If Check2.Value = 0 Then
    .ReportFileName = App.Path & ruta + "\librocaja_mayor.rpt"
  Else
    .ReportFileName = App.Path & ruta + "\librocaja_mayor_c.rpt"
  End If
Else
  If Check2.Value = 0 Then
    .ReportFileName = App.Path & ruta + "\librocaja_mayor2.rpt"
  Else
    .ReportFileName = App.Path & ruta + "\librocaja_mayor2_c.rpt"
  End If
End If
    .Connect = login.conexionreporte
    .DiscardSavedData = True
    .RetrieveDataFiles
    .ReportSource = 0
    .SQLQuery = tabla
    .Destination = crptToWindow
    .PrintFileType = crptCrystal
    .WindowState = crptMaximized
    .Action = 1
    
End With

End If
End Sub

Private Sub ingreso_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        If Val(ingreso.Text) > 0 Then
            egreso.Text = 0
            aceptar.SetFocus
            Exit Sub
        End If
        SendKeys "{tab}", False
    End If

End Sub

Private Sub ingreso_LostFocus()

        If Val(ingreso.Text) > 0 Then egreso.Text = 0

End Sub

Private Sub limpia_Click()

    DataCombo2.Text = ""
    DataCombo3.Text = ""
    Text1(0).Text = ""
    Text1(1).Text = ""
    ingreso.Text = 0
    egreso.Text = 0
    femision.Value = Date
    fvenci.Value = Date
    DataCombo2.SetFocus

End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)

    If SSTab1.Tab = 0 Then
        datmovimientos.RecordSource = "select librocajabanco.* from librocajabanco where fecha = '" & fecha.Value & "' and empresa = " & DataCombo1.BoundText & " AND cajabanco = 'C' and cuenta = " & DataCombo4.BoundText & " "
        datmovimientos.Refresh
        Call grilla_Click
        If login.librocajamodif = "N" Then
            eliminar.Enabled = False
            aceptar.Enabled = False
        Else
            eliminar.Enabled = True
            aceptar.Enabled = True
        End If
        If datmovimientos.Recordset.EOF = False Then
            datmovimientos.Recordset.MoveFirst
            If datmovimientos.Recordset.Fields("cerrado") = "S" Then
                eliminar.Enabled = False
                aceptar.Enabled = False
            End If
        End If
        DataGrid1.Columns(4).Visible = False
        DataGrid1.Columns(5).Visible = False
        DataGrid1.Columns(6).Visible = False
    End If
    If SSTab1.Tab = 1 Then
        datmovimientos.RecordSource = "select librocajabanco.* from librocajabanco where fecha = '" & fecha.Value & "' and empresa = " & DataCombo1.BoundText & " AND cajabanco = 'B' and cuenta = " & DataCombo5.BoundText & " "
        datmovimientos.Refresh
        Call grilla_Click
        If login.librocajamodif = "N" Then
            eliminar.Enabled = False
            aceptar.Enabled = False
        Else
            eliminar.Enabled = True
            aceptar.Enabled = True
        End If
        If datmovimientos.Recordset.EOF = False Then
            datmovimientos.Recordset.MoveFirst
            If datmovimientos.Recordset.Fields("cerrado") = "S" Then
                eliminar.Enabled = False
                aceptar.Enabled = False
            End If
        End If
        DataGrid1.Columns(4).Visible = True
        DataGrid1.Columns(5).Visible = True
        DataGrid1.Columns(6).Visible = True
        cuentaimputa = DataCombo5.BoundText
    End If
    

    
    If SSTab1.Tab > 1 Then
         DataGrid1.Visible = False
         For x = 0 To 7
            Text2(x).Visible = False
         Next x
         DataCombo2.Visible = False
         DataCombo3.Visible = False
         Text1(0).Visible = False
         Text1(1).Visible = False
         ingreso.Visible = False
         egreso.Visible = False
         eliminar.Visible = False
         aceptar.Visible = False
     Else
         DataGrid1.Visible = True
         For x = 0 To 7
            Text2(x).Visible = True
         Next x
         DataCombo2.Visible = True
         DataCombo3.Visible = True
         Text1(0).Visible = True
         Text1(1).Visible = True
         ingreso.Visible = True
         egreso.Visible = True
         eliminar.Visible = True
         aceptar.Visible = True
     End If

End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        If SSTab1.Tab = 0 Then
            SendKeys "{tab}", False
        Else
            If Index = 1 Then
                femision.SetFocus
            Else
                DataCombo5.SetFocus
            End If
        End If
    End If

End Sub

