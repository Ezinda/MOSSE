VERSION 5.00
Object = "{F6125AB1-8AB1-11CE-A77F-08002B2F4E98}#2.0#0"; "MSRDC20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{5734474E-78D3-4254-99B9-C35F31BDF509}#62.0#0"; "vbskpro2.ocx"
Begin VB.Form frmCuentas_viejo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Plan de Cuentas "
   ClientHeight    =   7245
   ClientLeft      =   1095
   ClientTop       =   330
   ClientWidth     =   11595
   HasDC           =   0   'False
   Icon            =   "frmCuentas_viejo.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7245
   ScaleWidth      =   11595
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text7 
      BackColor       =   &H80000004&
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
      Height          =   195
      Left            =   360
      TabIndex        =   44
      Text            =   "Nro.Cuenta:"
      Top             =   1320
      Width           =   1455
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H80000004&
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
      Height          =   195
      Left            =   360
      TabIndex        =   43
      Text            =   "Imputable s/n:"
      Top             =   960
      Width           =   1455
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H80000004&
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
      Height          =   195
      Left            =   360
      TabIndex        =   42
      Text            =   "Nombre Cuenta:"
      Top             =   600
      Width           =   1455
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H80000004&
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
      Height          =   195
      Left            =   360
      TabIndex        =   41
      Text            =   "Cod. Cuenta:"
      Top             =   240
      Width           =   1455
   End
   Begin VB.TextBox Text3 
      DataField       =   "empresa"
      DataSource      =   "criterio"
      Height          =   285
      Left            =   4080
      TabIndex        =   38
      Text            =   "Text3"
      Top             =   6120
      Visible         =   0   'False
      Width           =   855
   End
   Begin MSAdodcLib.Adodc datempresa 
      Height          =   330
      Left            =   2640
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
   Begin VB.TextBox detalle 
      BackColor       =   &H00E0E0E0&
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
      Left            =   120
      TabIndex        =   36
      Top             =   5760
      Width           =   11415
   End
   Begin VB.TextBox max6 
      DataField       =   "maximonivel16"
      DataSource      =   "maximonivel16"
      Height          =   285
      Left            =   5160
      TabIndex        =   27
      Text            =   "Text2"
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Frame Frame4 
      Height          =   15
      Left            =   120
      TabIndex        =   21
      Top             =   0
      Width           =   11655
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   8520
      ScaleHeight     =   615
      ScaleWidth      =   615
      TabIndex        =   20
      Top             =   2880
      Width           =   615
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   5655
      Left            =   6000
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   120
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   9975
      _Version        =   393217
      LineStyle       =   1
      Style           =   7
      Checkboxes      =   -1  'True
      Appearance      =   1
      MousePointer    =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton agregamismonivel 
      Caption         =   "&Agregar"
      Height          =   735
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   6360
      UseMaskColor    =   -1  'True
      Width           =   1095
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmCuentas_viejo.frx":0442
      Height          =   3975
      Left            =   120
      TabIndex        =   9
      Top             =   1800
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   7011
      _Version        =   393216
      AllowArrows     =   -1  'True
      BackColor       =   14737632
      ForeColor       =   -2147483642
      HeadLines       =   1
      RowHeight       =   15
      TabAction       =   2
      FormatLocked    =   -1  'True
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
      ColumnCount     =   5
      BeginProperty Column00 
         DataField       =   "idcuenta"
         Caption         =   "Cod Contable"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "Cod Contable"
         Caption         =   "Cod Abrev"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "Nombre Cuenta"
         Caption         =   "Nombre Cuenta"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "imp"
         Caption         =   "Imputable"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "Id Cuenta"
         Caption         =   "Cod Puro"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   4
         BeginProperty Column00 
            Alignment       =   2
            Locked          =   -1  'True
         EndProperty
         BeginProperty Column01 
            Alignment       =   2
            Locked          =   -1  'True
            ColumnWidth     =   720
         EndProperty
         BeginProperty Column02 
            Locked          =   -1  'True
         EndProperty
         BeginProperty Column03 
            Locked          =   -1  'True
         EndProperty
         BeginProperty Column04 
            Locked          =   -1  'True
            Object.Visible         =   0   'False
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cerrar 
      Caption         =   "&Cerrar"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   6480
      UseMaskColor    =   -1  'True
      Width           =   975
   End
   Begin VB.CommandButton Borrar 
      Caption         =   "&Borrrar"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6480
      UseMaskColor    =   -1  'True
      Width           =   855
   End
   Begin VB.CommandButton Nuevo 
      Caption         =   "&Nuevo"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6480
      UseMaskColor    =   -1  'True
      Width           =   855
   End
   Begin VB.CommandButton Grabar 
      Caption         =   "&Grabar"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6480
      UseMaskColor    =   -1  'True
      Width           =   855
   End
   Begin VB.TextBox codabrev 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      DataField       =   "Cod Contable"
      DataSource      =   "datPrimaryRS"
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
      ForeColor       =   &H00C0C0C0&
      Height          =   480
      Left            =   3480
      TabIndex        =   4
      Top             =   1155
      Width           =   1095
   End
   Begin VB.TextBox imputable 
      Alignment       =   2  'Center
      DataField       =   "imp"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Left            =   1920
      TabIndex        =   2
      Top             =   960
      Width           =   735
   End
   Begin VB.TextBox codcuenta 
      Alignment       =   2  'Center
      DataField       =   "Id Cuenta"
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "0.0.0.0.0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   11274
         SubFormatType   =   0
      EndProperty
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Left            =   1920
      TabIndex        =   0
      Top             =   240
      Width           =   1455
   End
   Begin VB.TextBox nombrecuenta 
      DataField       =   "Nombre Cuenta"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Left            =   1920
      TabIndex        =   1
      Top             =   600
      Width           =   3375
   End
   Begin VB.TextBox idcuenta0 
      Alignment       =   2  'Center
      DataField       =   "idcuenta"
      DataSource      =   "datPrimaryRS"
      Enabled         =   0   'False
      Height          =   285
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Frame Frame1 
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
      ForeColor       =   &H00C00000&
      Height          =   975
      Left            =   120
      TabIndex        =   11
      Top             =   6240
      Width           =   5775
      Begin VB.CommandButton Command1 
         Cancel          =   -1  'True
         Caption         =   "Cance&lar"
         Height          =   615
         Left            =   2160
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   855
      End
      Begin Crystal.CrystalReport crystalreporte 
         Left            =   5040
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
         WindowShowGroupTree=   -1  'True
         WindowShowCloseBtn=   -1  'True
         WindowShowSearchBtn=   -1  'True
         WindowShowPrintSetupBtn=   -1  'True
         WindowShowRefreshBtn=   -1  'True
      End
      Begin MSRDC.MSRDC reporte 
         Height          =   375
         Left            =   1080
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
   End
   Begin VB.Frame Frame2 
      Height          =   1695
      Left            =   120
      TabIndex        =   12
      Top             =   0
      Width           =   5775
      Begin VB.TextBox Text1 
         BackColor       =   &H80000004&
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
         Height          =   195
         Left            =   3480
         TabIndex        =   16
         Text            =   "Cod.Abrev."
         Top             =   960
         Width           =   975
      End
      Begin MSDataListLib.DataCombo idcuentacombo 
         Bindings        =   "frmCuentas_viejo.frx":045D
         DataSource      =   "datPrimaryRS"
         Height          =   315
         Left            =   3240
         TabIndex        =   28
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         ListField       =   "idcuenta"
         BoundColumn     =   ""
         Text            =   ""
      End
      Begin VB.TextBox max5 
         DataField       =   "maximonivel15"
         DataSource      =   "maximonivel15"
         Height          =   285
         Left            =   3120
         TabIndex        =   26
         Text            =   "Text2"
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox max4 
         DataField       =   "maximonivel14"
         DataSource      =   "maximonivel14"
         Height          =   285
         Left            =   2760
         TabIndex        =   25
         Text            =   "Text2"
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox max3 
         DataField       =   "maximonivel13"
         DataSource      =   "maximonivel13"
         Height          =   285
         Left            =   2400
         TabIndex        =   24
         Text            =   "Text2"
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox max2 
         DataField       =   "maximonivel12"
         DataSource      =   "maximonivel12"
         Height          =   285
         Left            =   2040
         TabIndex        =   23
         Text            =   "Text2"
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox max1 
         DataField       =   "maximonivel11"
         DataSource      =   "maximonivel11"
         Height          =   285
         Left            =   1680
         TabIndex        =   22
         Text            =   "Text2"
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin MSDataListLib.DataCombo combocuenta 
         Bindings        =   "frmCuentas_viejo.frx":0478
         DataSource      =   "datPrimaryRS"
         Height          =   360
         Left            =   4560
         TabIndex        =   15
         Top             =   1200
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   741
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         ListField       =   "Cod Contable"
         BoundColumn     =   ""
         Text            =   ""
         Object.DataMember      =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.CommandButton acttree 
         Height          =   735
         Left            =   3120
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   840
         UseMaskColor    =   -1  'True
         Visible         =   0   'False
         Width           =   855
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Frame3"
      Height          =   1335
      Left            =   6000
      TabIndex        =   17
      Top             =   5880
      Width           =   5535
      Begin VB.CommandButton reordenar 
         Caption         =   "Reordenar Codigos"
         Height          =   735
         Left            =   4440
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   480
         UseMaskColor    =   -1  'True
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CommandButton imprimir 
         Caption         =   "&Imprimir"
         Height          =   735
         Left            =   3240
         Picture         =   "frmCuentas_viejo.frx":0493
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   480
         UseMaskColor    =   -1  'True
         Width           =   975
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Desactivar"
         Height          =   735
         Left            =   1800
         Picture         =   "frmCuentas_viejo.frx":09C5
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   480
         UseMaskColor    =   -1  'True
         Width           =   975
      End
   End
   Begin MSAdodcLib.Adodc niveles 
      Height          =   330
      Left            =   6240
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
   Begin MSAdodcLib.Adodc datPrimaryRS 
      Height          =   330
      Left            =   1200
      Top             =   6840
      Visible         =   0   'False
      Width           =   3135
      _ExtentX        =   5530
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
      CacheSize       =   2000
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   1
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
   Begin MSAdodcLib.Adodc maximonivel11 
      Height          =   330
      Left            =   600
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
   Begin MSAdodcLib.Adodc maximonivel12 
      Height          =   330
      Left            =   1200
      Top             =   6360
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
   Begin MSAdodcLib.Adodc maximonivel13 
      Height          =   330
      Left            =   1920
      Top             =   6360
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
   Begin MSAdodcLib.Adodc maximonivel14 
      Height          =   330
      Left            =   2640
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
   Begin MSAdodcLib.Adodc maximonivel15 
      Height          =   330
      Left            =   3240
      Top             =   6360
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
   Begin MSAdodcLib.Adodc maximonivel16 
      Height          =   330
      Left            =   3840
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
   Begin VB.TextBox car5 
      DataField       =   "niv5"
      DataSource      =   "niveles"
      Height          =   285
      Left            =   7680
      TabIndex        =   33
      Text            =   "Text6"
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox car4 
      DataField       =   "niv4"
      DataSource      =   "niveles"
      Height          =   285
      Left            =   7320
      TabIndex        =   32
      Text            =   "Text5"
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox car3 
      DataField       =   "niv3"
      DataSource      =   "niveles"
      Height          =   285
      Left            =   6960
      TabIndex        =   31
      Text            =   "Text4"
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox car2 
      DataField       =   "niv2"
      DataSource      =   "niveles"
      Height          =   285
      Left            =   6600
      TabIndex        =   30
      Text            =   "Text3"
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox car1 
      DataField       =   "niv1"
      DataSource      =   "niveles"
      Height          =   285
      Left            =   6240
      TabIndex        =   29
      Text            =   "Text2"
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox codempresa 
      DataField       =   "empre"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Left            =   8880
      TabIndex        =   34
      Text            =   "Text2"
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox empresacod 
      DataField       =   "empre"
      DataSource      =   "niveles"
      Height          =   285
      Left            =   8040
      TabIndex        =   35
      Top             =   120
      Width           =   495
   End
   Begin VB.TextBox Text2 
      DataField       =   "empresaactiva"
      DataSource      =   "refrezcaempresa"
      Height          =   285
      Left            =   2280
      TabIndex        =   37
      Text            =   "Text2"
      Top             =   2520
      Visible         =   0   'False
      Width           =   495
   End
   Begin MSAdodcLib.Adodc criterio 
      Height          =   375
      Left            =   5520
      Top             =   6000
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
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
   Begin vbskpro.Skinner Skinner1 
      Left            =   6000
      Top             =   0
      _ExtentX        =   1270
      _ExtentY        =   1270
      CloseButtonToolTipText=   "Cerrar"
      MinButtonToolTipText=   "Minimizar"
      MaxButtonToolTipText=   "Maximizar"
      RestoreButtonToolTipText=   "Restaurar"
      ChangeSkinButton=   0   'False
      MinToBarButtonToolTipText=   "Minimizar a la barra de títulos"
      RestoreFromBarButtonToolTipText=   "Restaurar ventana"
      AlwaysOnTopButtonToolTipText=   "Hacer siempre visible"
      AlwaysOnTopDownButtonToolTipText=   "Hacer no siempre visible"
      ChangeSkinButtonToolTipText=   "Cambiar skin"
      HelpButtonToolTipText=   "Ayuda"
      SysEnableSkinCaption=   "Habilitar &Skin"
      SysDisableSkinCaption=   "Deshabilitar &Skin"
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
End
Attribute VB_Name = "frmCuentas_viejo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cantnodos As Integer
Dim cuenta As String
Dim posicion(10000) As Integer
Dim hijos(10000) As Integer
Dim ultimonodo2 As Integer
Dim ultimonodo3 As Integer
Dim ultimonodo5 As Integer
Dim nivelanterior As String
Dim ruta As String
Dim canthijos As Integer
Dim nodogeneral As Integer
Dim bandera As Integer
Dim ni1, ni2, ni3, ni4, ni5 As String
Dim n1, n2, n3, n4, n5, c1, c2, c3, c4, c5 As Integer



Private Sub acttree_Click()
On Error Resume Next

Dim nodx As Node
Dim imputable1 As String

maximonivel11.RecordSource = "SELECT MAX([Cod Contable]) AS maximonivel11, empre From Cuentas WHERE (LEFT(idcuenta, 1) = '1') and empre = " & login.empresaact & " and inicioper = '" & login.iper & "' and finper = '" & login.fper & "' GROUP BY empre"
maximonivel12.RecordSource = "SELECT MAX([Cod Contable]) AS maximonivel12, empre From Cuentas WHERE (LEFT(idcuenta, 1) = '2') and empre = " & login.empresaact & " and inicioper = '" & login.iper & "' and finper = '" & login.fper & "' GROUP BY empre"
maximonivel13.RecordSource = "SELECT MAX([Cod Contable]) AS maximonivel13, empre From Cuentas WHERE (LEFT(idcuenta, 1) = '3') and empre = " & login.empresaact & " and inicioper = '" & login.iper & "' and finper = '" & login.fper & "' GROUP BY empre"
maximonivel14.RecordSource = "SELECT MAX([Cod Contable]) AS maximonivel14, empre From Cuentas WHERE (LEFT(idcuenta, 1) = '4') and empre = " & login.empresaact & " and inicioper = '" & login.iper & "' and finper = '" & login.fper & "' GROUP BY empre"
maximonivel15.RecordSource = "SELECT MAX([Cod Contable]) AS maximonivel15, empre From Cuentas WHERE (LEFT(idcuenta, 1) = '5') and empre = " & login.empresaact & " and inicioper = '" & login.iper & "' and finper = '" & login.fper & "' GROUP BY empre"
maximonivel16.RecordSource = "SELECT MAX([Cod Contable]) AS maximonivel16, empre From Cuentas WHERE (LEFT(idcuenta, 1) = '6') and empre = " & login.empresaact & " and inicioper = '" & login.iper & "' and finper = '" & login.fper & "' GROUP BY empre"

  maximonivel11.Refresh
  maximonivel12.Refresh
  maximonivel13.Refresh
  maximonivel14.Refresh
  maximonivel15.Refresh
  maximonivel16.Refresh

If Command2.Caption = "Acti&var" Then
    Picture1.Visible = True
    Exit Sub
Else
    Picture1.Visible = False
End If

TreeView1.Nodes.Clear  'Limpia el Treeview
  
 If datPrimaryRS.Recordset.RecordCount = 0 Then Exit Sub
 datPrimaryRS.Recordset.MoveLast
 ultimo = datPrimaryRS.Recordset.AbsolutePosition
 datPrimaryRS.Recordset.MoveFirst

pru:

registro = datPrimaryRS.Recordset.AbsolutePosition

cuenta = Str(codcuenta.Text)

n1 = 2: c1 = Val(car1)
n2 = n1 + c1: c2 = Val(car2)
n3 = n2 + c2: c3 = Val(car3)
n4 = n3 + c3: c4 = Val(car4)
n5 = n4 + c4: c5 = Val(car5)

ni1 = Mid(cuenta, n1, c1)
ni2 = Mid(cuenta, n2, c2)
ni3 = Mid(cuenta, n3, c3)
ni4 = Mid(cuenta, n4, c4)
ni5 = Mid(cuenta, n5, c5)

nombrenodo = nombrecuenta.Text
imputable1 = imputable.Text

If Val(ni2) = 0 Then
      nodouno = nodouno + 1
      Set nodx = TreeView1.Nodes.Add(, , "t" + ni1, nombrenodo)
      GoTo salta
End If
If Val(ni2) <> 0 And Val(ni3) = 0 Then
      Set nodx = TreeView1.Nodes.Add("t" + ni1, tvwChild, "a" + ni1 + ni2, nombrenodo)
      GoTo salta
End If
If Val(ni3) <> 0 And Val(ni4) = 0 Then
      Set nodx = TreeView1.Nodes.Add("a" + ni1 + ni2, tvwChild, "p" + ni1 + ni2 + ni3, nombrenodo)
      GoTo salta
End If
If Val(ni4) <> 0 And Val(ni5) = 0 Then
      Set nodx = TreeView1.Nodes.Add("p" + ni1 + ni2 + ni3, tvwChild, "h" + ni1 + ni2 + ni3 + ni4, nombrenodo)
      GoTo salta
End If
If Val(ni4) <> 0 And Val(ni5) <> 0 Then
      Set nodx = TreeView1.Nodes.Add("h" + ni1 + ni2 + ni3 + ni4, tvwChild, "n" + ni1 + ni2 + ni3 + ni4 + ni5, nombrenodo)
      GoTo salta
End If

salta:
posicion(nodx.Index) = registro
If imputable1 = "S" Then nodx.Checked = True
If registro = ultimo Then GoTo fin
datPrimaryRS.Recordset.MoveNext
GoTo pru

fin:
For x = 1 To ultimo
     hijos(x) = TreeView1.Nodes.Item(x).Children
Next x


fuera:
End Sub



Private Sub cancelar_Click()
  On Error GoTo cancelaErr



  datPrimaryRS.Recordset.Cancel
Exit Sub
cancelaErr:
  MsgBox Err.Description
End Sub

Private Sub actualiza_Click()
  Call acttree_Click
End Sub

Private Sub agregamismonivel_Click()
On Error GoTo errortagrega
'agrega nodo

    If login.plancuentasaltas = "N" Then
        mensa = MsgBox("Acceso Denegado", , "Sistema")
        Exit Sub
    End If
       
        nodogeneral = nodogeneral + 1
        nivel = 0
        For x = 1 To Len(ruta)
               letra = Mid(ruta, x, 1)
               If letra = "\" Then nivel = nivel + 1
        Next x
        If nivelanterior = "" Then
                Set nodx = TreeView1.Nodes.Add(Val(ni1), tvwChild, "a" + Str(nodogeneral), "Nuevo")
        Else
                Set nodx = TreeView1.Nodes.Add(nivelanterior, tvwChild, "a" + Str(nodogeneral), "Nuevo")
        End If
        
        Select Case nivel
        Case 0
            ni1 = Mid(cuenta, n1, c1)
            prev = Right(Str(canthijos + 1), Len(Str(canthijos + 1)) - 1)
            ni2 = Mid("00000000000000", n2, c2 - Len(prev)) + prev
            ni3 = Mid("00000000000000", n3, c3)
            ni4 = Mid("00000000000000", n4, c4)
            ni5 = Mid("00000000000000", n5, c5)
        Case 1
            ni1 = Mid(cuenta, n1, c1)
            ni2 = Mid(cuenta, n2, c2)
            prev = Right(Str(canthijos + 1), Len(Str(canthijos + 1)) - 1)
            ni3 = Mid("00000000000000", n3, c3 - Len(prev)) + prev
            ni4 = Mid("00000000000000", n4, c4)
            ni5 = Mid("00000000000000", n5, c5)
        Case 2
            ni1 = Mid(cuenta, n1, c1)
            ni2 = Mid(cuenta, n2, c2)
            ni3 = Mid(cuenta, n3, c3)
            prev = Right(Str(canthijos + 1), Len(Str(canthijos + 1)) - 1)
            ni4 = Mid("00000000000000", n4, c4 - Len(prev)) + prev
            ni5 = Mid("00000000000000", n5, c5)
        Case 3
            ni1 = Mid(cuenta, n1, c1)
            ni2 = Mid(cuenta, n2, c2)
            ni3 = Mid(cuenta, n3, c3)
            ni4 = Mid(cuenta, n4, c4)
            If canthijos > 100 Then canthijos = canthijos + 1
            prev = Right(Str(canthijos + 1), Len(Str(canthijos + 1)) - 1)
            ni5 = Mid("00000000000000", n5, c5 - Len(prev)) + prev
        End Select
            
        codprevio = ni1 + ni2 + ni3 + ni4 + ni5
        datPrimaryRS.Recordset.AddNew
        codempresa = empresacod
        codcuenta = Str(codprevio)
        imputable = ""
        datPrimaryRS.Recordset.Fields(6) = login.iper
        datPrimaryRS.Recordset.Fields(7) = login.fper
        posicion(nodx.Index) = datPrimaryRS.Recordset.AbsolutePosition
        
 
 Exit Sub
errortagrega:
    mensa = MsgBox("Error al dar de alta este nivel", vbCritical, "Error")
    Call Command1_Click

End Sub


Private Sub agregamismonivel_LostFocus()
 agregamismonivel.Enabled = False
 borrar.Enabled = True
End Sub

Private Sub borrar_Click()
  On Error GoTo DeleteErr
  
  
    If login.plancuentasbajas = "N" Then
        mensa = MsgBox("Acceso Denegado", , "Sistema")
        Exit Sub
    End If
  
  DataGrid1.SetFocus
  DataGrid1.Bookmark = datPrimaryRS.Recordset.AbsolutePosition
  KeyAscii = 13
  Respuesta = MsgBox("ESTA POR BORRAR UNA CUENTA, ESTA SEGURO?", vbYesNo, "Atención")
If Respuesta = vbYes Then
    If hijos(datPrimaryRS.Recordset.AbsolutePosition) > 0 Then
        mensa = MsgBox("No puede eliminar esta cuenta, elimine primero los subrrubros", vbCritical, "Error")
        Exit Sub
    End If
    datPrimaryRS.Recordset.Delete
    Call acttree_Click
Rem    datPrimaryRS.Refresh
Rem    Call Grabar_Click
Else
    Exit Sub
End If

 
 Exit Sub
DeleteErr:
  MsgBox "No se pudo borrar, pulse el boton ´Grabar´ e intente eliminar el registro nuevamente"
  Call Command1_Click
End Sub

Private Sub cerrar_Click()

    Unload Me

End Sub

Private Sub codcuenta_Change()
On Error GoTo fuera

If codcuenta <> "" Then cuenta = Str(codcuenta.Text)
n1 = 2: c1 = Val(car1)
n2 = n1 + c1: c2 = Val(car2)
n3 = n2 + c2: c3 = Val(car3)
n4 = n3 + c3: c4 = Val(car4)
n5 = n4 + c4: c5 = Val(car5)

ni1 = Mid(cuenta, n1, c1)
ni2 = Mid(cuenta, n2, c2)
ni3 = Mid(cuenta, n3, c3)
ni4 = Mid(cuenta, n4, c4)
ni5 = Mid(cuenta, n5, c5)


If codabrev = "" Then
    If max1.Text <> "" Then
        maxim1 = Right(max1.Text, Len(max1.Text) - 1)
    End If
    If max2.Text <> "" Then
        maxim2 = Right(max2.Text, Len(max2.Text) - 1)
    End If
    If max3.Text <> "" Then
        maxim3 = Right(max3.Text, Len(max3.Text) - 1)
    End If
    If max4.Text <> "" Then
        maxim4 = Right(max4.Text, Len(max4.Text) - 1)
    End If
    If max5.Text <> "" Then
        maxim5 = Right(max5.Text, Len(max5.Text) - 1)
    End If
    If max6.Text <> "" Then
        maxim6 = Right(max6.Text, Len(max6.Text) - 1)
    End If
    If max1.Text <> "" And ni1 = "1" Then codabrev = Val(maxim1) + 1
    If max2.Text <> "" And ni1 = "2" Then codabrev = Val(maxim2) + 1
    If max3.Text <> "" And ni1 = "3" Then codabrev = Val(maxim3) + 1
    If max4.Text <> "" And ni1 = "4" Then codabrev = Val(maxim4) + 1
    If max5.Text <> "" And ni1 = "5" Then codabrev = Val(maxim5) + 1
    If max6.Text <> "" And ni1 = "6" Then codabrev = Val(maxim6) + 1
    If codabrev <> "" Then
        codigoabreviado = ni1 + Right(Str(codabrev), Len(Str(codabrev)) - 1)
        codabrev = Val(codigoabreviado)
    End If
End If
    
  If max1.Text = "" And ni1 = "1" Then codabrev = 10
  If max2.Text = "" And ni1 = "2" Then codabrev = 20
  If max3.Text = "" And ni1 = "3" Then codabrev = 30
  If max4.Text = "" And ni1 = "4" Then codabrev = 40
  If max5.Text = "" And ni1 = "5" Then codabrev = 50
  If max6.Text = "" And ni1 = "6" Then codabrev = 60

  If bandera = 0 Then idcuenta0.Text = ni1 + "." + ni2 + "." + ni3 + "." + ni4 + "." + ni5
 
fuera:
End Sub

Private Sub codcuenta_KeyPress(KeyAscii As Integer)
On Error GoTo fuera

    If KeyAscii = 15 Then reordenar.Visible = True
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        If Val(ni2) <> 0 Or Val(ni3) <> 0 Or Val(ni4) <> 0 Or Val(ni5) <> 0 Then
            mensa = MsgBox("Solo puede ingresar en el primer nivel", vbCritical, "Error")
            codcuenta.SetFocus
            Exit Sub
        End If
        If Len(codcuenta.Text) <> c1 + c2 + c3 + c4 + c5 Then
            mensa = MsgBox("Codigo fuera de rango", vbCritical, "Error")
            codcuenta.SetFocus
            Exit Sub
        End If
        nombrecuenta.SetFocus
    End If

fuera:
End Sub

Private Sub codcuenta_LostFocus()
idcuentacombo = codcuenta
combocuenta = codabrev
End Sub


Private Sub combocuenta_Change()
On Error GoTo fuera

If combocuenta.SelectedItem <> 0 Then datPrimaryRS.Recordset.Bookmark = combocuenta.SelectedItem

fuera:
End Sub


Private Sub combocuenta_Validate(Cancel As Boolean)
combocuenta.Refresh
End Sub

Private Sub Command1_Click()
    bandera = 1
    datPrimaryRS.Refresh
    Call grabar_Click

End Sub



Private Sub Command2_Click()
On Error GoTo fin

        If Command2.Caption = "&Desactivar" Then
            TreeView1.Enabled = False
            Command2.Caption = "Acti&var"
            Picture1.Visible = True
            GoTo fin
        End If
        If Command2.Caption = "Acti&var" Then
            TreeView1.Enabled = True
            Command2.Caption = "&Desactivar"
             Picture1.Visible = False
        End If

fin:

End Sub

Private Sub DataCombo1_Change()
On Error GoTo fuera

        niveles.Recordset.Bookmark = DataCombo1.SelectedItem
    
fuera:
End Sub



Private Sub DataGrid1_Click()
 borrar.Enabled = True
End Sub


Private Sub Form_Load()

  bandera = 1
  
  datempresa.ConnectionString = login.conexiontotal
  datPrimaryRS.ConnectionString = login.conexiontotal
  criterio.ConnectionString = login.conexiontotal
  niveles.ConnectionString = login.conexiontotal
  maximonivel11.ConnectionString = login.conexiontotal
  maximonivel12.ConnectionString = login.conexiontotal
  maximonivel13.ConnectionString = login.conexiontotal
  maximonivel14.ConnectionString = login.conexiontotal
  maximonivel15.ConnectionString = login.conexiontotal
  maximonivel16.ConnectionString = login.conexiontotal
  
  
  
  criterio.RecordSource = "select empreactiva.* from empreactiva"
  criterio.Refresh
  
  datPrimaryRS.RecordSource = "select cuentas.* from cuentas WHERE inicioper = '" & login.iper & "' and cuentas.empre = " & login.empresaact & " ORDER BY IDCUENTA"
  datPrimaryRS.Refresh
  
  niveles.RecordSource = "select niveles.* from niveles Where empre = " & login.empresaact & " and inicioper = '" & login.iper & "' "
  niveles.Refresh
  
  datempresa.RecordSource = "select empresa.* from empresa Where empresa = " & login.empresaact & ""
  datempresa.Refresh
  
  
  If login.plancuentasmodi = "N" Then
    TreeView1.LabelEdit = tvwManual
    TreeView1.Checkboxes = False
    codcuenta.Locked = True
    nombrecuenta.Locked = True
    imputable.Locked = True
  Else
    TreeView1.LabelEdit = tvwAutomatic
    TreeView1.Checkboxes = True
    codcuenta.Locked = False
    nombrecuenta.Locked = False
    imputable.Locked = False
  End If
      
  agregamismonivel.Enabled = False
  
  borrar.Enabled = True
  Call acttree_Click
  nodogeneral = 5000
  

End Sub

Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault
  
End Sub

Private Sub datPrimaryRS_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)
  'Aquí es donde puede colocar el código de control de errores
  'Si desea pasar por alto los errores, marque como comentario la siguiente línea
  'Si desea detectarlos, agregue código aquí para controlarlos
  MsgBox "Data error event hit err:" & Description
End Sub

Private Sub datPrimaryRS_Movelete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  'Esto mostrará la posición de registro actual para este Recordset
  datPrimaryRS.Caption = "Registro: " & CStr(datPrimaryRS.Recordset.AbsolutePosition)
End Sub

Private Sub data1_WillChangeRecord(ByVal adReason As ADODB.EventReasonEnum, ByVal cRecords As Long, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  'Aquí se coloca el código de validación
  'Se llama a este evento cuando ocurre la siguiente acción
  Dim bCancel As Boolean

  Select Case adReason
  Case adRsnAddNew
  Case adRsnClose
  Case adRsnDelete
  Case adRsnFirstChange
  Case adRsnMove
  Case adRsnRequery
  Case adRsnResynch
  Case adRsnUndoAddNew
  Case adRsnUndoDelete
  Case adRsnUndoUpdate
  Case adRsnUpdate
  End Select

  If bCancel Then adStatus = adStatusCancel
End Sub

Private Sub cmdAdd_Click()
  On Error GoTo AddErr
  datPrimaryRS.Recordset.AddNew



  Exit Sub
AddErr:
  MsgBox Err.Description
End Sub

Private Sub cmdDelete_Click()
  On Error GoTo DeleteErr
  With datPrimaryRS.Recordset
    .Delete
    .MoveNext
    If .EOF Then .MoveLast
  End With
  Exit Sub
DeleteErr:
  MsgBox Err.Description
End Sub


Private Sub cmdUpdate_Click()
  On Error GoTo UpdateErr



  datPrimaryRS.Recordset.UpdateBatch adAffectAll
  Exit Sub
UpdateErr:
  MsgBox Err.Description
End Sub

Private Sub cmdClose_Click()
  Unload Me
End Sub


Private Sub grabar_Click()
On Error GoTo grabErr

  codempresa = empresacod
  datPrimaryRS.Recordset.Fields(6) = login.iper
  datPrimaryRS.Recordset.Fields(7) = login.fper
  datPrimaryRS.Recordset.UpdateBatch adAffectCurrent
  
  
    Inicio.datauditoria.ConnectionString = login.conexiontotal
    
    Inicio.datauditoria.RecordSource = "select auditoria.* from auditoria"
    Inicio.datauditoria.Refresh
    
    Inicio.datauditoria.Recordset.AddNew
    Inicio.datauditoria.Recordset.Fields("fecha") = Date
    Inicio.datauditoria.Recordset.Fields("hora") = Str(Time)
    Inicio.datauditoria.Recordset.Fields("ventana") = "Plan de Cuentas"
    Inicio.datauditoria.Recordset.Fields("accion") = "Alta/Modificacion cuenta:" + Str(codabrev)
    Inicio.datauditoria.Recordset.Fields("usuario") = login.usuarioactivo
    Inicio.datauditoria.Recordset.Fields("empresa") = login.empresaact
    Inicio.datauditoria.Recordset.UpdateBatch adAffectCurrent
  

  Call acttree_Click
  Exit Sub
grabErr:
  MsgBox Err.Description
End Sub



Private Sub idcuentacombo_Change()
On Error GoTo fuera

If idcuentacombo.SelectedItem <> 0 Then datPrimaryRS.Recordset.Bookmark = idcuentacombo.SelectedItem

fuera:
End Sub



Private Sub imprimir_Click()
On Error GoTo fuera

Dim ruta As String

ruta = "\Empresa" + Right(Str(login.empresaact), Len(Str(login.empresaact)) - 1)

    Text3.Text = login.empresaact
    criterio.Recordset.Fields(1) = login.iper
    criterio.Recordset.Fields(2) = login.fper
    criterio.Recordset.UpdateBatch adAffectCurrent
    criterio.Refresh

reporte.SQL = "SELECT plancuentas.empre, plancuentas.Nombrecuenta, plancuentas.imp, plancuentas.niv1, plancuentas.niv2, plancuentas.razonsocial, plancuentas.idcuenta, plancuentas.codcontable, plancuentas.inicioper, plancuentas.finper FROM contablesql.dbo.plancuentas plancuentas WHERE plancuentas.inicioper = '" & login.iper & "' and plancuentas.empre = " & login.empresaact & " ORDER BY plancuentas.idcuenta ASC"
tabla = reporte.SQL

With CrystalReporte
    .ReportFileName = App.Path & ruta + "\Plancuentas.rpt"
    Debug.Print login.conexionreporte
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
Exit Sub
fuera:
    mensa = MsgBox("No se puede Imprimir", vbCritical, "Error")

End Sub

Private Sub imputable_KeyPress(KeyAscii As Integer)
On Error GoTo fuera

    If KeyAscii = 13 Then
        KeyAscii = 0
        grabar.SetFocus
    End If
    
fuera:
End Sub

Private Sub imputable_LostFocus()
On Error GoTo fin

    If imputable = "s" Or imputable = "S" Then
            imputable.Text = "S"
            GoTo fin
    End If
    If imputable = "n" Or imputable = "N" Then
            imputable.Text = "N"
            GoTo fin
    End If
     
     Mensaje = MsgBox("Tecla no corresponde", vbCritical, "Error")
     imputable.SetFocus
fin:
End Sub



Private Sub menureenumerar_Click(Index As Integer)
On Error GoTo fuera

  KeyAscii = 13
Respuesta = MsgBox("ESTO SOLO PUEDE REALIZARCE SI NO TIENE MOVIMIENTOS CARGADOS, DEBE ESTAR TOTALMENTE SEGURO PARA REALIZAR ESTA TAREA, REALIZA LA ACCION ? ", vbYesNo, "!! Atención !!")
If Respuesta = vbYes Then
    datPrimaryRS.Recordset.MoveFirst
    orden = 0
paso1:
    codabrev0 = Mid(Str(codcuenta.Text), 2, 1)
    codabrev1 = Str(orden)
    codabrev2 = codabrev0 + Right(codabrev1, Len(codabrev1) - 1)
    codabrev = Str(codabrev2)
    orden = orden + 1
    datPrimaryRS.Recordset.MoveNext
    If datPrimaryRS.Recordset.EOF = True Then GoTo paso2
    If Mid(Str(codcuenta.Text), 2, 1) <> codabrev0 Then orden = 0
    GoTo paso1
paso2:
    mensa = MsgBox("Codigo Abreviado Reenumerado", vbDefaultButton1)
    datPrimaryRS.Recordset.MoveLast
    Exit Sub
Else
    Exit Sub
End If
    
fuera:
End Sub



Private Sub nombrecuenta_KeyPress(KeyAscii As Integer)
On Error GoTo fuera

    If KeyAscii = 13 Then
        KeyAscii = 0
        imputable.SetFocus
    End If
    
fuera:
End Sub

Private Sub nuevo_Click()
  On Error GoTo AddErr
  
    If login.plancuentasaltas = "N" Then
        mensa = MsgBox("Acceso Denegado", , "Sistema")
        Exit Sub
    End If
  
  datPrimaryRS.Recordset.AddNew
  bandera = 0
  codabrev = ""
  codcuenta.SetFocus

    
  Exit Sub
AddErr:
  MsgBox Err.Description
End Sub

Private Sub Nuevo_KeyPress(KeyAscii As Integer)
On Error GoTo fuera

    If KeyAscii = 13 Then
        KeyAscii = 0
        codcuenta.SetFocus
    End If

fuera:
End Sub



Private Sub reordenar_Click()
On Error GoTo fin

contador = 0
    datPrimaryRS.Recordset.MoveFirst
paso1:
    If datPrimaryRS.Recordset.EOF = True Then GoTo fin
    digito1 = Left(codcuenta, 1)
    contador = contador + 1
    If digito0 <> digito1 Then contador = 0
    Numero = Str(digito1) + Right(Str(contador), Len(Str(contador)) - 1)
    codabrev = Val(Numero)
    datPrimaryRS.Recordset.UpdateBatch adAffectCurrent
    If datPrimaryRS.Recordset.EOF = False Then
        digito0 = Left(codcuenta, 1)
        datPrimaryRS.Recordset.MoveNext
        GoTo paso1
    Else
        GoTo fin
    End If
    
fin:
End Sub

Private Sub reordenar_LostFocus()

        reordenar.Visible = False
    
End Sub

Private Sub TreeView1_AfterLabelEdit(Cancel As Integer, NewString As String)
On Error GoTo fuera

   Dim pos As Integer
            
            bandera = 0
            nombrecuenta = NewString
            codempresa = empresacod
            datPrimaryRS.Recordset.UpdateBatch adAffectCurrent
fuera:

End Sub

Private Sub TreeView1_NodeCheck(ByVal Node As MSComctlLib.Node)
On Error GoTo fuera

    bandera = 0
   pos = posicion(Node.Index)
   datPrimaryRS.Recordset.AbsolutePosition = pos
   If Node.Checked = True Then
            imputable.Text = "S"
   Else
            imputable.Text = "N"
   End If
        
fuera:
End Sub

Private Sub TreeView1_NodeClick(ByVal Node As Node)
On Error GoTo errortree
   Dim pos As Integer

    bandera = 0
   agregamismonivel.Enabled = True
   Rem borrar.Enabled = False
    
   pos = posicion(Node.Index)
   datPrimaryRS.Recordset.AbsolutePosition = pos
          
   nombrecuenta = Node.Text
   If Node.Checked = True Then
            imputable.Text = "S"
   Else
            imputable.Text = "N"
   End If
      
   nivelanterior = Node.Key
   canthijos = Node.Children
   ruta = Node.FullPath
   hijos(pos) = canthijos
   detalle.Text = idcuenta0.Text + "=>" + ruta
   
   nivel = 0
   For x = 1 To Len(ruta)
          letra = Mid(ruta, x, 1)
          If letra = "\" Then nivel = nivel + 1
   Next x
   
   If nivel = 4 Then
     agregamismonivel.Enabled = False
   Else
     agregamismonivel.Enabled = True
   End If


ni1 = Mid(cuenta, n1, c1)
ni2 = Mid(cuenta, n2, c2)
ni3 = Mid(cuenta, n3, c3)
ni4 = Mid(cuenta, n4, c4)
ni5 = Mid(cuenta, n5, c5)

maximonivel11.RecordSource = "SELECT MAX([Cod Contable]) AS maximonivel11, empre From Cuentas WHERE (LEFT(idcuenta, 1) = '1') and empre = " & login.empresaact & " and inicioper = '" & login.iper & "' and finper = '" & login.fper & "' GROUP BY empre"
maximonivel12.RecordSource = "SELECT MAX([Cod Contable]) AS maximonivel12, empre From Cuentas WHERE (LEFT(idcuenta, 1) = '2') and empre = " & login.empresaact & " and inicioper = '" & login.iper & "' and finper = '" & login.fper & "' GROUP BY empre"
maximonivel13.RecordSource = "SELECT MAX([Cod Contable]) AS maximonivel13, empre From Cuentas WHERE (LEFT(idcuenta, 1) = '3') and empre = " & login.empresaact & " and inicioper = '" & login.iper & "' and finper = '" & login.fper & "' GROUP BY empre"
maximonivel14.RecordSource = "SELECT MAX([Cod Contable]) AS maximonivel14, empre From Cuentas WHERE (LEFT(idcuenta, 1) = '4') and empre = " & login.empresaact & " and inicioper = '" & login.iper & "' and finper = '" & login.fper & "' GROUP BY empre"
maximonivel15.RecordSource = "SELECT MAX([Cod Contable]) AS maximonivel15, empre From Cuentas WHERE (LEFT(idcuenta, 1) = '5') and empre = " & login.empresaact & " and inicioper = '" & login.iper & "' and finper = '" & login.fper & "' GROUP BY empre"
maximonivel16.RecordSource = "SELECT MAX([Cod Contable]) AS maximonivel16, empre From Cuentas WHERE (LEFT(idcuenta, 1) = '6') and empre = " & login.empresaact & " and inicioper = '" & login.iper & "' and finper = '" & login.fper & "' GROUP BY empre"

            maximonivel11.Refresh
            maximonivel12.Refresh
            maximonivel13.Refresh
            maximonivel14.Refresh
            maximonivel15.Refresh
            maximonivel16.Refresh
    
 Exit Sub
errortree:
    mensa = MsgBox("Error al dar de alta este nivel", vbCritical, "Error")
          
End Sub
