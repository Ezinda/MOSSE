VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{5734474E-78D3-4254-99B9-C35F31BDF509}#62.0#0"; "vbskpro2.ocx"
Object = "{D18BBD1F-82BB-4385-BED3-E9D31A3E361E}#1.0#0"; "KewlButtonz.ocx"
Begin VB.Form frmasientoscancelacion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Carga de Minutas contables"
   ClientHeight    =   6270
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10200
   ControlBox      =   0   'False
   Icon            =   "frmasientoscancelacion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6270
   ScaleWidth      =   10200
   Begin MSDataListLib.DataList DataList2 
      Bindings        =   "frmasientoscancelacion.frx":0442
      Height          =   2205
      Left            =   600
      TabIndex        =   2
      Top             =   2760
      Visible         =   0   'False
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   3889
      _Version        =   393216
      MatchEntry      =   -1  'True
      OLEDragMode     =   1
      OLEDropMode     =   1
      BackColor       =   16777215
      ListField       =   "codigo"
      BoundColumn     =   "Cod Contable"
      Object.DataMember      =   ""
   End
   Begin MSDataListLib.DataList DataList3 
      Bindings        =   "frmasientoscancelacion.frx":045B
      Height          =   1620
      Left            =   2880
      TabIndex        =   15
      Top             =   3720
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   2752
      _Version        =   393216
      MatchEntry      =   -1  'True
      BackColor       =   -2147483626
      ForeColor       =   -2147483647
      ListField       =   "ccostoslista"
      BoundColumn     =   "cc"
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
   Begin VB.TextBox Text9 
      Alignment       =   2  'Center
      DataField       =   "ccosto"
      DataSource      =   "datasiento"
      Height          =   285
      Left            =   2880
      TabIndex        =   14
      Top             =   3360
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000016&
      Caption         =   "Centros de Costo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   2640
      TabIndex        =   16
      Top             =   3000
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Frame Frame4 
      Caption         =   "Movimeintos"
      Height          =   1935
      Left            =   240
      TabIndex        =   39
      Top             =   1800
      Width           =   8055
      Begin VB.CommandButton Cuenta 
         Caption         =   "Nº Cuenta"
         Height          =   255
         Index           =   0
         Left            =   360
         Picture         =   "frmasientoscancelacion.frx":0478
         TabIndex        =   50
         TabStop         =   0   'False
         Top             =   360
         UseMaskColor    =   -1  'True
         Width           =   855
      End
      Begin VB.CommandButton Cuenta 
         Caption         =   "Debe"
         Height          =   255
         Index           =   1
         Left            =   1680
         Picture         =   "frmasientoscancelacion.frx":09AA
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   360
         UseMaskColor    =   -1  'True
         Width           =   855
      End
      Begin VB.CommandButton Cuenta 
         Caption         =   "Haber"
         Height          =   255
         Index           =   2
         Left            =   3120
         Picture         =   "frmasientoscancelacion.frx":0EDC
         TabIndex        =   48
         TabStop         =   0   'False
         Top             =   360
         UseMaskColor    =   -1  'True
         Width           =   855
      End
      Begin VB.CommandButton Cuenta 
         Caption         =   "Detalle"
         Height          =   255
         Index           =   3
         Left            =   4680
         Picture         =   "frmasientoscancelacion.frx":140E
         TabIndex        =   47
         TabStop         =   0   'False
         Top             =   360
         UseMaskColor    =   -1  'True
         Width           =   1815
      End
      Begin VB.CommandButton eliminarmovimiento 
         Caption         =   "&Eliminar Movimiento"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4680
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   1200
         Width           =   1935
      End
      Begin VB.CommandButton nuevomovimiento 
         Caption         =   "&Nuevo Movimiento"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2520
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   1200
         Width           =   1815
      End
      Begin VB.CommandButton grabamovimiento 
         Caption         =   "Gra&bar Movimiento"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   1200
         Width           =   1935
      End
      Begin VB.TextBox Text3 
         DataField       =   "detallefila"
         DataSource      =   "datasiento"
         Height          =   285
         Index           =   6
         Left            =   4320
         MaxLength       =   50
         TabIndex        =   43
         Top             =   600
         Width           =   3375
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         DataField       =   "Haber"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         DataSource      =   "datasiento"
         Height          =   285
         Index           =   5
         Left            =   3000
         TabIndex        =   42
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         DataField       =   "Debe"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         DataSource      =   "datasiento"
         Height          =   285
         Index           =   4
         Left            =   1560
         TabIndex        =   41
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox Text4 
         Alignment       =   2  'Center
         DataField       =   "idcuenta"
         DataSource      =   "datasiento"
         Height          =   285
         Left            =   360
         TabIndex        =   40
         Top             =   600
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Cabecera de Asiento"
      Height          =   1695
      Left            =   240
      TabIndex        =   20
      Top             =   120
      Width           =   8055
      Begin VB.CommandButton Cuenta 
         Caption         =   "Fecha Asiento:"
         Height          =   255
         Index           =   4
         Left            =   240
         Picture         =   "frmasientoscancelacion.frx":1940
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   480
         UseMaskColor    =   -1  'True
         Width           =   1455
      End
      Begin VB.CommandButton Cuenta 
         Caption         =   "Fecha Registro:"
         Height          =   255
         Index           =   9
         Left            =   3360
         Picture         =   "frmasientoscancelacion.frx":1E72
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   480
         UseMaskColor    =   -1  'True
         Width           =   1455
      End
      Begin VB.CommandButton Cuenta 
         Caption         =   "Nº de Asiento:"
         Height          =   255
         Index           =   5
         Left            =   240
         Picture         =   "frmasientoscancelacion.frx":23A4
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   840
         UseMaskColor    =   -1  'True
         Width           =   1455
      End
      Begin VB.CommandButton Cuenta 
         Caption         =   "Referencia:"
         Height          =   255
         Index           =   6
         Left            =   240
         Picture         =   "frmasientoscancelacion.frx":28D6
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   1200
         UseMaskColor    =   -1  'True
         Width           =   1455
      End
      Begin VB.CommandButton Cuenta 
         Caption         =   "Periodo:"
         Height          =   255
         Index           =   10
         Left            =   2760
         Picture         =   "frmasientoscancelacion.frx":2E08
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   840
         UseMaskColor    =   -1  'True
         Width           =   735
      End
      Begin VB.TextBox Text5 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   285
         Left            =   4800
         Locked          =   -1  'True
         TabIndex        =   32
         Top             =   480
         Width           =   1455
      End
      Begin VB.TextBox Text3 
         DataField       =   "idmasterasientos"
         DataSource      =   "datasiento"
         Height          =   285
         Index           =   8
         Left            =   7320
         TabIndex        =   31
         Top             =   960
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox Text3 
         DataField       =   "Fecha"
         DataSource      =   "datasiento"
         Height          =   285
         Index           =   2
         Left            =   6960
         TabIndex        =   30
         Top             =   960
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox Text3 
         DataField       =   "empresa"
         DataSource      =   "datasiento"
         Height          =   285
         Index           =   1
         Left            =   6600
         TabIndex        =   29
         Top             =   960
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox Text3 
         DataField       =   "idasiento"
         DataSource      =   "datasiento"
         Height          =   285
         Index           =   0
         Left            =   6240
         TabIndex        =   28
         Top             =   960
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   6480
         Locked          =   -1  'True
         TabIndex        =   27
         Text            =   "Text2"
         Top             =   240
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         DataField       =   "cerrado"
         DataSource      =   "datmaestro"
         Height          =   285
         Index           =   6
         Left            =   7200
         Locked          =   -1  'True
         TabIndex        =   26
         Top             =   600
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox Text1 
         DataField       =   "empresa"
         DataSource      =   "datmaestro"
         Height          =   285
         Index           =   5
         Left            =   6600
         Locked          =   -1  'True
         TabIndex        =   25
         Top             =   600
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         DataField       =   "perfinal"
         DataSource      =   "datmaestro"
         Height          =   285
         Index           =   4
         Left            =   4800
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   840
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         DataField       =   "perinicial"
         DataSource      =   "datmaestro"
         Height          =   285
         Index           =   3
         Left            =   3480
         Locked          =   -1  'True
         TabIndex        =   23
         Top             =   840
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   2
         Left            =   2040
         MaxLength       =   50
         TabIndex        =   22
         Top             =   1200
         Width           =   4215
      End
      Begin VB.TextBox asiento 
         Alignment       =   2  'Center
         DataField       =   "nroasiento"
         DataSource      =   "datmaestro"
         Height          =   285
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   840
         Width           =   735
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         DataField       =   "fecha"
         Height          =   285
         Left            =   2040
         TabIndex        =   0
         Top             =   480
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         AllowPrompt     =   -1  'True
         Format          =   "dd/mm/yyyy"
         PromptChar      =   "_"
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "frmasientoscancelacion.frx":333A
         Height          =   315
         Left            =   2040
         TabIndex        =   33
         Top             =   120
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         BackColor       =   16777215
         ListField       =   "razonsocial"
         BoundColumn     =   "empresa"
         Text            =   ""
      End
   End
   Begin VB.CommandButton llena 
      Caption         =   "llena"
      Height          =   255
      Left            =   7200
      TabIndex        =   19
      Top             =   5880
      Visible         =   0   'False
      Width           =   855
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grilla 
      Height          =   1935
      Left            =   240
      TabIndex        =   18
      Top             =   3840
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   3413
      _Version        =   393216
      Rows            =   2000
      Cols            =   6
      SelectionMode   =   1
      AllowUserResizing=   2
      _NumberOfBands  =   1
      _Band(0).Cols   =   6
   End
   Begin VB.CommandButton nuevo0 
      Caption         =   "Command1"
      Height          =   255
      Left            =   8640
      TabIndex        =   17
      Top             =   3600
      Visible         =   0   'False
      Width           =   855
   End
   Begin MSAdodcLib.Adodc datccostos 
      Height          =   330
      Left            =   7080
      Top             =   5520
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
   Begin MSMask.MaskEdBox Masksaldo 
      CausesValidation=   0   'False
      Height          =   255
      Left            =   5400
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   6000
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   0
      Appearance      =   0
      BackColor       =   -2147483633
      ForeColor       =   4210752
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "#,##0.00;-#,##0.00"
      PromptChar      =   "_"
   End
   Begin VB.TextBox Text6 
      Alignment       =   1  'Right Justify
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """$"" #,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   11274
         SubFormatType   =   2
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   8520
      Locked          =   -1  'True
      TabIndex        =   7
      TabStop         =   0   'False
      Text            =   " "
      Top             =   5880
      Visible         =   0   'False
      Width           =   855
   End
   Begin MSMask.MaskEdBox Maskhaber 
      CausesValidation=   0   'False
      Height          =   255
      Left            =   3600
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   6000
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   0
      Appearance      =   0
      BackColor       =   -2147483633
      ForeColor       =   0
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "#,##0.00;-#,##0.00"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox Maskdebe 
      CausesValidation=   0   'False
      Height          =   255
      Left            =   1800
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   6000
      Width           =   1710
      _ExtentX        =   3016
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   0
      Appearance      =   0
      BackColor       =   -2147483633
      ForeColor       =   0
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "#,##0.00;-#,##0.00"
      PromptChar      =   "_"
   End
   Begin VB.TextBox Text6 
      Alignment       =   1  'Right Justify
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """$"" #,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   11274
         SubFormatType   =   2
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   7680
      Locked          =   -1  'True
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   5880
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox Text6 
      Alignment       =   1  'Right Justify
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """$"" #,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   11274
         SubFormatType   =   2
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   6960
      Locked          =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   5880
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSDataGridLib.DataGrid DataGrid2 
      Bindings        =   "frmasientoscancelacion.frx":3353
      Height          =   255
      Left            =   1200
      TabIndex        =   3
      Top             =   5880
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   450
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
   Begin VB.CommandButton grabar 
      Caption         =   "grabar"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   5760
      Visible         =   0   'False
      Width           =   1095
   End
   Begin MSAdodcLib.Adodc datmaestro 
      Height          =   330
      Left            =   960
      Top             =   6240
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
   Begin MSAdodcLib.Adodc datperiodo 
      Height          =   330
      Left            =   1440
      Top             =   5640
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
   Begin MSAdodcLib.Adodc datasiento 
      Height          =   330
      Left            =   2640
      Top             =   5640
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
   Begin MSAdodcLib.Adodc datcuentas 
      Height          =   330
      Left            =   6840
      Top             =   5640
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
      Height          =   3615
      Left            =   8400
      TabIndex        =   4
      Top             =   120
      Width           =   1575
      Begin KewlButtonz.KewlButtons nuevo 
         Height          =   615
         Left            =   240
         TabIndex        =   52
         Top             =   360
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   1085
         BTYPE           =   14
         TX              =   "&Grabar Asiento"
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
         BCOL            =   -2147483629
         BCOLO           =   -2147483629
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmasientoscancelacion.frx":336C
         PICN            =   "frmasientoscancelacion.frx":3388
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
         Height          =   615
         Left            =   240
         TabIndex        =   53
         Top             =   2640
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   1085
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
         BCOL            =   -2147483629
         BCOLO           =   -2147483629
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmasientoscancelacion.frx":4E0A
         PICN            =   "frmasientoscancelacion.frx":4E26
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin KewlButtonz.KewlButtons ver 
         Height          =   615
         Left            =   240
         TabIndex        =   51
         Top             =   1440
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   1085
         BTYPE           =   14
         TX              =   "&Ver o Modif."
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
         BCOL            =   -2147483629
         BCOLO           =   -2147483629
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmasientoscancelacion.frx":5970
         PICN            =   "frmasientoscancelacion.frx":598C
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
   Begin MSAdodcLib.Adodc datlistacostos 
      Height          =   330
      Left            =   6840
      Top             =   6240
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
   Begin vbskpro.Skinner Skinner1 
      Left            =   0
      Top             =   6120
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
      LcK2            =   $"frmasientoscancelacion.frx":8D7E
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
   Begin MSAdodcLib.Adodc datmaestro1 
      Height          =   330
      Left            =   4080
      Top             =   6240
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
   Begin MSAdodcLib.Adodc datempresa 
      Height          =   330
      Left            =   8160
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
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Saldo"
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
      Left            =   5400
      TabIndex        =   13
      Top             =   6240
      Width           =   615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Haber"
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
      Left            =   3600
      TabIndex        =   12
      Top             =   6240
      Width           =   615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Debe"
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
      Left            =   1800
      TabIndex        =   11
      Top             =   6240
      Width           =   615
   End
End
Attribute VB_Name = "frmasientoscancelacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public movimientohijo As Double
Dim detalle As String
Dim posicion As Double
Dim numerodisco As Double
Dim registro As Integer
Dim registrobloqueado As Integer
Dim sumadebe As Currency
Dim sumahaber As Currency
Dim modi As Integer
Dim nograbar As Integer


Private Sub asiento_Change()

Rem If Text2.Text <> "" Then
Rem    movimientohijo = Text2.Text
Rem Else
Rem   movimientohijo = 0
Rem End If

    
Rem    datasiento.RecordSource = "select [Detalle Asientos].* from [Detalle Asientos] where idmasterasientos = " & frmasientos.movimientohijo & " "
Rem    datasiento.Refresh

End Sub



Private Sub Command1_Click()



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
        frmasientos.Show
    End If
    
fuera:

End Sub

Private Sub DataList2_Click()
On Error GoTo fuera

    Text4.Text = DataList2.BoundText
    
fuera:
End Sub

Private Sub DataList2_GotFocus()
On Error GoTo fuera

    If Inicio.opcion1 = True Then
        datcuentas.RecordSource = "select listacuentas.* from listacuentas WHERE empre = " & login.empresaact & " and inicioper = '" & login.iper & "' and finper = '" & login.fper & "' ORDER BY IDCUENTA"
        datcuentas.Refresh
        DataList2.ListField = "codigo"
    End If
    If Inicio.opcion2 = True Then
        datcuentas.RecordSource = "select listacuentas.* from listacuentas WHERE empre = " & login.empresaact & " and inicioper = '" & login.iper & "' and finper = '" & login.fper & "' ORDER BY nombre"
        datcuentas.Refresh
        DataList2.ListField = "nombre"
    End If
    
fuera:
End Sub

Private Sub DataList2_KeyPress(KeyAscii As Integer)
On Error GoTo fuera

    If KeyAscii = 13 Then
        KeyAscii = 0
        Text4.Text = DataList2.BoundText
            If datccostos.Recordset.EOF = True Then GoTo sigue
              datccostos.Recordset.MoveFirst
              digito = Val(datccostos.Recordset.Fields(3))
              digito1 = Val(datccostos.Recordset.Fields(4))
              digcue = Val(Mid(Text4.Text, 1, 1))
              If digcue = digito Or digcue = digito1 And login.habcc = True Then
                DataList3.Visible = True
                Text9.Visible = True
                Frame1.Visible = True
                DataList3.SetFocus
                Exit Sub
            End If
sigue:
        Text3(4).SetFocus
    End If
    
fuera:
End Sub

Private Sub DataList2_LostFocus()

    DataList2.Visible = False
                     
End Sub

Private Sub DataList3_Click()
On Error GoTo fuera

    Text9.Text = DataList3.BoundText
    
fuera:
End Sub

Private Sub DataList3_KeyPress(KeyAscii As Integer)
On Error GoTo fuera

    If KeyAscii = 13 Then
        KeyAscii = 0
        Text9.Text = DataList3.BoundText
        Text3(4).SetFocus
    End If
    
fuera:
End Sub

Private Sub DataList3_LostFocus()
On Error GoTo fuera

    If Text9.Text = "" Then
        mensa = MsgBox("Debe ingresa un Centro de Costo", vbCritical, "!Error¡")
        DataList3.SetFocus
        Exit Sub
    End If
Frame1.Visible = False
Text9.Visible = False
DataList3.Visible = False

fuera:
End Sub

Private Sub eliminarmovimiento_Click()
On Error GoTo erroreliminar
Dim Cuenta(2000) As Integer
Dim debe(2000) As Currency
Dim haber(2000) As Currency
Dim detalle(2000) As String
Dim ccosto(2000) As Integer

    mensa = MsgBox("Esta por eliminar un movimiento de este asiento, esta seguro", vbYesNo, "!! Atención !!")
    If mensa = vbYes Then
        grilla.Col = 1
        grilla.Text = "x"
        grilla.Col = 2
        grilla.Text = ""
        grilla.Col = 3
        grilla.Text = ""
        grilla.Col = 4
        grilla.Text = ""
        grilla.Col = 5
        grilla.Text = ""
        
        Text9.Text = ""
        Text4.Text = ""
        Text3(4).Text = ""
        Text3(5).Text = ""
        Text3(6).Text = ""
        Y = grilla.Row
        For X = Y To 2000
            grilla.Row = X + 1
            grilla.Col = 1
            If grilla.Text = "" Then Exit For
            Cuenta(X) = grilla.Text
            grilla.Col = 2
            debe(X) = grilla.Text
            grilla.Col = 3
            haber(X) = grilla.Text
            grilla.Col = 4
            detalle(X) = grilla.Text
            grilla.Col = 5
            ccosto(X) = grilla.Text
        Next X
        For Z = Y To X - 1
            grilla.Row = Z
            grilla.Col = 1
            grilla.Text = Cuenta(Z)
            grilla.Col = 2
            grilla.Text = debe(Z)
            grilla.Col = 3
            grilla.Text = haber(Z)
            grilla.Col = 4
            grilla.Text = detalle(Z)
            grilla.Col = 5
            grilla.Text = ccosto(Z)
        Next Z

        sumadebe = 0
        sumahaber = 0

        For Y = 1 To 2000
            grilla.Row = Y
            grilla.Col = 1
            If grilla.Text = "" Then Exit For
        Next Y
            grilla.Row = Y - 1
            grilla.Col = 1
            grilla.Text = ""
            grilla.Col = 2
            grilla.Text = ""
            grilla.Col = 3
            grilla.Text = ""
            grilla.Col = 4
            grilla.Text = ""
            grilla.Col = 5
            grilla.Text = ""

        For X = 1 To Y
            grilla.Row = X
            grilla.Col = 2
            Text6(0).Text = Replace(grilla.Text, ",", "")
            sumadebe = Val(Text6(0).Text) + sumadebe
            grilla.Text = Format(grilla.Text, "#,###,##0.00")
            grilla.Col = 3
            Text6(1).Text = Replace(grilla.Text, ",", "")
            grilla.Text = Format(grilla.Text, "#,###,##0.00")
            sumahaber = Val(Text6(1).Text) + sumahaber
        Next X


        Maskdebe.Text = sumadebe
        Maskhaber.Text = sumahaber
        Masksaldo.Text = sumadebe - sumahaber

        Text4.Text = ""
        Text9.Text = ""
        Text3(4).Text = ""
        Text3(5).Text = ""
        Text3(6).Text = ""


    End If

erroreliminar:


End Sub

Private Sub Form_Load()
Aplicar_skin Me

datasiento.ConnectionString = login.conexiontotal
datccostos.ConnectionString = login.conexiontotal
datcuentas.ConnectionString = login.conexiontotal
datlistacostos.ConnectionString = login.conexiontotal
datmaestro.ConnectionString = login.conexiontotal
datmaestro1.ConnectionString = login.conexiontotal
datperiodo.ConnectionString = login.conexiontotal
datempresa.ConnectionString = login.conexiontotal


  DataCombo1.Text = login.nomempresa

  Inicio.Caption = login.nomempresa + "-Periodo Contable: " + Str(login.iper) + " -" + Str(login.fper)
 
  datempresa.RecordSource = "select empresa.* from empresa"
  datempresa.Refresh

  frmasientoscancelacion.Left = 0
  frmasientoscancelacion.Top = 0
  
  Inicio.Toolbar1.Visible = True
  
  grilla.ColWidth(0) = 300
  grilla.ColWidth(1) = 800
  grilla.ColAlignment(1) = 4
  grilla.ColWidth(2) = 1500
  grilla.ColAlignment(2) = 8
  grilla.ColWidth(3) = 1500
  grilla.ColAlignment(3) = 8
  grilla.ColWidth(4) = 3500
  grilla.ColAlignment(4) = 2
  grilla.ColWidth(5) = 800
  grilla.ColAlignment(5) = 2
  
  
  grilla.Col = 1
  grilla.Row = 0
  grilla.Text = "Cuenta"
  grilla.Col = 2
  grilla.Text = "Debe"
  grilla.Col = 3
  grilla.Text = "Haber"
  grilla.Col = 4
  grilla.Text = "Detalle"
  grilla.Col = 5
  grilla.Text = "C.Cost."
  
  MaskEdBox1.Text = Date
  Text1(2).Text = ""
  
For X = 1 To 99 Step 2
    For Y = 0 To 5
        grilla.Col = Y
        grilla.Row = X
        grilla.CellBackColor = QBColor(11)
    Next Y
Next X
  
    datlistacostos.RecordSource = "select listaccostos.* from listaccostos WHERE empresa = " & login.empresaact & ""
    datlistacostos.Refresh
    
    
    datcuentas.RecordSource = "select listacuentas.* from listacuentas WHERE empre = " & login.empresaact & " and inicioper = '" & login.iper & "' and finper = '" & login.fper & "' ORDER BY IDCUENTA"
    datcuentas.Refresh
  
    datmaestro.RecordSource = "select [Maestro Asientos].* from [Maestro Asientos] where empresa = " & login.empresaact & " and perinicial = '" & login.iper & "' order by nroasiento"
    datmaestro.Refresh

    datperiodo.RecordSource = "select EMPRESA.* from EMPRESA where empresa = " & login.empresaact & ""
    datperiodo.Refresh

    datccostos.RecordSource = "SELECT ccostos.* FROM ccostos WHERE empresa = " & login.empresaact & ""
    datccostos.Refresh

    MaskEdBox1.Mask = ""
    Text6(2).Text = 0
    nograbar = 0
    
    Call nuevo0_Click
        
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Inicio.Toolbar1.Visible = False
    
End Sub

Private Sub grabamovimiento_Click()
On Error GoTo errorasiento

    detalle = Text3(6).Text
       
    Text6(0).Text = 0
    Text6(1).Text = 0
    Text6(2).Text = 0
If Text4.Text <> "" Then
    registro = registro + 1
    
    grilla.Row = registro
    grilla.Col = 1
    grilla.Text = Text4.Text
    grilla.Col = 2
    grilla.Text = Text3(4).Text
    grilla.Col = 3
    grilla.Text = Text3(5).Text
    grilla.Col = 4
    grilla.Text = Text3(6).Text
    grilla.Col = 5
    grilla.Text = Text9.Text
End If

sumadebe = 0
sumahaber = 0

For Y = 1 To 2000
    grilla.Row = Y
    grilla.Col = 1
    If grilla.Text = "" Then Exit For
Next Y

For X = 1 To Y
    grilla.Row = X
    grilla.Col = 2
    Text6(0).Text = Replace(grilla.Text, ",", "")
    sumadebe = Val(Text6(0).Text) + sumadebe
    grilla.Col = 3
    Text6(1).Text = Replace(grilla.Text, ",", "")
    sumahaber = Val(Text6(1).Text) + sumahaber
 Next X


    Maskdebe.Text = sumadebe
    Maskhaber.Text = sumahaber
    Masksaldo.Text = sumadebe - sumahaber

   Text4.Text = ""
   Text3(4).Text = ""
   Text3(5).Text = ""
   Text3(6).Text = ""
   Text9.Text = ""

    nuevomovimiento.SetFocus
Exit Sub
errorasiento:
    mensa = MsgBox("Debe ingresar una referencia y presionar enter para grabar los movimientos", vbCritical, "Atencion !!")
    Text1(2).SetFocus
End Sub

Private Sub grabar_Click()
On Error GoTo fuera

    campoaño = Right(MaskEdBox1.Text, 4)
    campomes = Mid(MaskEdBox1.Text, 4, 2)
    campodia = Left(MaskEdBox1.Text, 2)
    campofecha = campoaño + "/" + campomes + "/" + campodia
    
    campoaño1 = Right(Text1(3).Text, 4)
    campomes1 = Mid(Text1(3).Text, 4, 2)
    campodia1 = Left(Text1(3).Text, 2)
    campofecha1 = campoaño1 + "/" + campomes1 + "/" + campodia1
    
    campoaño2 = Right(Text1(4).Text, 4)
    campomes2 = Mid(Text1(4).Text, 4, 2)
    campodia2 = Left(Text1(4).Text, 2)
    campofecha2 = campoaño2 + "/" + campomes2 + "/" + campodia2

    If campofecha < campofecha1 Or campofecha > campofecha2 Then
            mensa = MsgBox("La Fecha es erronea o no pertenecia al periodo en ejercicio", vbCritical, "!! Atención !!")
            MaskEdBox1.SelLength = 10
            MaskEdBox1.SetFocus
            nograbar = 1
            Exit Sub
    End If
    
   Text3(1).Text = login.empresaact
   Text3(2).Text = MaskEdBox1.Text
   Text3(8).Text = Text2.Text
   Text5.Text = Date
   Text1(2).SetFocus
   nograbar = 0
    
fuera:

End Sub




Private Sub movimientos_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub grilla_Click()

    Call llena_Click

End Sub

Private Sub grilla_KeyDown(KeyCode As Integer, Shift As Integer)

    Call llena_Click

End Sub

Private Sub llena_Click()

    If grilla.Row <= registrobloqueado Then
        MsgBox "No se puede modificar este movimiento", vbCritical, "Error"
        Exit Sub
    End If

    modi = 1
    grilla.Col = 1
    Text4.Text = grilla.Text
    grilla.Col = 2
    Text3(4).Text = grilla.Text
    grilla.Col = 3
    Text3(5).Text = grilla.Text
    grilla.Col = 4
    Text3(6).Text = grilla.Text
    registro = grilla.Row - 1
    


End Sub

Private Sub MaskEdBox1_GotFocus()

    For X = 1 To 2000
        grilla.Row = X
        grilla.Col = 1
        If grilla.Text = "" Then Exit For
    Next X
    
    grilla.Row = X - 1
    registrobloqueado = X - 1
    registro = X - 1


End Sub

Private Sub MaskEdBox1_KeyPress(KeyAscii As Integer)
On Error GoTo fuera

    If KeyAscii = 13 Or KeyAscii = 9 Then
        KeyAscii = 0
        Call grabar_Click
    End If
    
fuera:
End Sub



Private Sub nuevo1_Click()

End Sub

Private Sub nuevo_Click()
On Error GoTo fuera
Dim ultimoasiento As Double

    Call grabar_Click
    If nograbar = 1 Then
        grilla.Clear
        Call Form_Load
        Exit Sub
    End If

     If Masksaldo.Text <> 0 Then
          mensa = MsgBox("EL Asiento está Desvalanceado, no puede grabar", vbCritical, "!! Error !!")
          grilla.SetFocus
          Exit Sub
     End If
     If Text1(2).Text = "" Then
          mensa = MsgBox("Debe ingresar un texto de referencia", vbCritical, "!! Error !!")
          Text1(2).SetFocus
          Exit Sub
     End If
        datmaestro1.RecordSource = "select [Maestro Asientos].* from [Maestro Asientos] where empresa = " & login.empresaact & " and perinicial = '" & login.iper & "' order by nroasiento"
        datmaestro1.Refresh
        If datmaestro1.Recordset.EOF = True Then
            ultimoasiento = 1
        Else
            datmaestro1.Recordset.MoveLast
            ultimoasiento = datmaestro1.Recordset.Fields(3) + 1
            asiento.Text = ultimoasiento
        End If
        datmaestro1.Recordset.AddNew
        datmaestro1.Recordset.Fields("fecha") = MaskEdBox1.Text
        datmaestro1.Recordset.Fields("fecharegistro") = Text5.Text
        datmaestro1.Recordset.Fields("nroasiento") = ultimoasiento
        datmaestro1.Recordset.Fields("concepto") = Text1(2).Text
        datmaestro1.Recordset.Fields("perinicial") = login.iper
        datmaestro1.Recordset.Fields("perfinal") = login.fper
        datmaestro1.Recordset.Fields("empresa") = login.empresaact
        datmaestro1.Recordset.Fields(11) = "S"
        datmaestro1.Recordset.UpdateBatch adAffectCurrent
        
        Inicio.datauditoria.ConnectionString = login.conexiontotal
    
        Inicio.datauditoria.RecordSource = "select auditoria.* from auditoria"
        Inicio.datauditoria.Refresh
    
        Inicio.datauditoria.Recordset.AddNew
        Inicio.datauditoria.Recordset.Fields("fecha") = Date
        Inicio.datauditoria.Recordset.Fields("hora") = Str(Time)
        Inicio.datauditoria.Recordset.Fields("ventana") = "Carga de Minutas Contables"
        Inicio.datauditoria.Recordset.Fields("accion") = "Alta Asiento:" + Str(ultimoasiento - 1) + " Periodo:" + Str(login.iper) + "-" + Str(login.fper)
        Inicio.datauditoria.Recordset.Fields("usuario") = login.usuarioactivo
        Inicio.datauditoria.Recordset.Fields("empresa") = login.empresaact
        Inicio.datauditoria.Recordset.UpdateBatch adAffectCurrent
        
             
        Text2.Text = datmaestro1.Recordset.Fields("idmasterasientos")
        datasiento.RecordSource = "select [Detalle Asientos].* from [Detalle Asientos] where empresa = " & login.empresaact & " and idmasterasientos = 0"
        datasiento.Refresh
        
        contarerror = 0
        For X = 1 To 2000
            datasiento.Recordset.AddNew
            datasiento.Recordset.Fields("idmasterasientos") = datmaestro1.Recordset.Fields("idmasterasientos")
            grilla.Row = X
            grilla.Col = 1
            If grilla.Text = "" Then Exit For
            contarerror = 1
            datasiento.Recordset.Fields("fecha") = MaskEdBox1.Text
            contarerror = 2
            datasiento.Recordset.Fields("idmasterasientos") = datmaestro1.Recordset.Fields("idmasterasientos")
            contarerror = 3
            datasiento.Recordset.Fields("empresa") = login.empresaact
            contarerror = 4
            datasiento.Recordset.Fields("idcuenta") = grilla.Text
            contarerror = 5
            grilla.Col = 2
            datasiento.Recordset.Fields("debe") = grilla.Text
            contarerror = 6
            grilla.Col = 3
            datasiento.Recordset.Fields("haber") = grilla.Text
            contarerror = 7
            grilla.Col = 4
            datasiento.Recordset.Fields("detallefila") = grilla.Text
            contarerror = 8
            grilla.Col = 5
            If grilla.Text <> "" Then
                datasiento.Recordset.Fields("ccosto") = grilla.Text
            End If
            datasiento.Recordset.UpdateBatch adAffectCurrent
        Next X
        
        grilla.Clear
        datasiento.RecordSource = "select [Detalle Asientos].* from [Detalle Asientos] where empresa = " & login.empresaact & " and idmasterasientos = " & Text2.Text & " order by idasiento"
        datasiento.Refresh
        datasiento.Recordset.MoveFirst
        frmconciliacion.Text2.Text = datasiento.Recordset.Fields("idasiento")
        frmconciliacion.Text6.Text = datmaestro1.Recordset.Fields("nroasiento")
        frmconciliacion.lineagrilla = 1
        Unload Me
               
      
Exit Sub
fuera:

mensa = MsgBox("No se pudo grabar el asiento, error" + Str(contarerror), vbCritical, "Error")


End Sub

Private Sub nuevo0_Click()
On Error Resume Next
            
        
    modi = 0
    If datmaestro.Recordset.EOF = True Then
        ultimoasiento = 1
    Else
        datmaestro.Recordset.MoveLast
        ultimoasiento = datmaestro.Recordset.Fields(3) + 1
    End If
    datmaestro.Recordset.Filter = "empresa = 0"

    MaskEdBox1.Mask = "##/##/####"
    MaskEdBox1.SelLength = 10
    MaskEdBox1.SelText = ""
    asiento.Text = ultimoasiento
    Text1(5).Text = login.empresaact
    Text1(6).Text = "N"
    Text1(3).Text = login.iper
    Text1(4).Text = login.fper
    Text5.Text = Date
    registro = frmconciliacion.lineagrilla
    MaskEdBox1.SetFocus
    

fuera:
End Sub

Private Sub nuevomovimiento_Click()
On Error GoTo fuera
    
   Text3(1).Text = login.empresaact
   Text3(2).Text = MaskEdBox1.Text
   Text3(8).Text = Text2.Text
   Text3(6).Text = detalle

   Text4.SetFocus
   
fuera:
End Sub

Private Sub salir_Click()
On Error GoTo errorsalir

     If Text6(2).Text <> 0 Then
          mensa = MsgBox("EL Asiento está Desvalanceado, no puede grabar, desea realmente salir", vbYesNo, "!! Error !!")
          If mensa = vbYes Then GoTo sale
          datasiento.Recordset.MoveLast
          Text3(4).SetFocus
          Exit Sub
     End If
sale:
         frmconciliacion.lineagrilla = 0
         Unload Me
    Exit Sub
    
errorsalir:

End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
On Error GoTo fuera

    If KeyAscii = 13 Or KeyAscii = 9 Then
        KeyAscii = 0
        Text4.SetFocus
    End If

fuera:
End Sub

Private Sub Text1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)

     If KeyCode = 38 Then MaskEdBox1.SetFocus

End Sub

Private Sub Text3_GotFocus(Index As Integer)
On Error GoTo fuera

    If Text4.Text = "" Then
        mensa = MsgBox("Cuenta no existente", vbCritical, "!! Atencion !!")
        Text4.SetFocus
        Exit Sub
    End If
    

        Text3(4).Text = Format(Text3(4).Text, "#,###,##0.00")
        Text3(5).Text = Format(Text3(5).Text, "#,###,##0.00")


        Text3(Index).SelLength = Len(Text3(Index).Text)

        DataList2.Visible = False

fuera:
End Sub



Private Sub Text3_KeyPress(Index As Integer, KeyAscii As Integer)
On Error GoTo fuera

    If KeyAscii = 13 Or KeyAscii = 9 Then
        KeyAscii = 0
        If Index = 6 Then
            grabamovimiento.SetFocus
            Exit Sub
        End If
        Text3(Index + 1).SetFocus
    End If
  
fuera:
End Sub

Private Sub Text3_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
     
     If KeyCode = 38 And Index > 4 Then Text3(Index - 1).SetFocus
     If KeyCode = 38 And Index = 4 Then Text4.SetFocus

End Sub

Private Sub Text3_LostFocus(Index As Integer)
On Error GoTo fuera

        If Index = 4 Then
            If Text3(Index).Text = "" Then Text3(Index).Text = 0
            If Val(Text3(Index).Text) <> 0 Then
                Text3(Index + 1).Text = 0
            End If
        End If
        If Index = 5 Then
            If Val(Text3(Index).Text) <> 0 Then
                Text3(Index - 1).Text = 0
            End If
        End If
fuera:
End Sub


Private Sub Text4_GotFocus()
On Error GoTo fuera


    Text4.SelLength = Len(Text4)
    DataList2.BoundText = Text4.Text
    DataList2.Visible = True
    DataList2.Left = Text4.Left
Rem    DataList2.Top = Text4.Top + Text4.Height + 600
    DataList2.SetFocus
                  
fuera:
End Sub


Private Sub Ver_Click()
On Error GoTo fuera

     If Text6(2).Text <> 0 Then
          mensa = MsgBox("EL Asiento está Desvalanceado, no puede grabar", vbCritical, "!! Error !!")
          datasiento.Recordset.MoveLast
          Text3(4).SetFocus
          Exit Sub
     End If
    
    frmasientosbusca.Show
    Unload Me
    
fuera:
End Sub
