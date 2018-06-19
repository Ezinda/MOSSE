VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{5734474E-78D3-4254-99B9-C35F31BDF509}#62.0#0"; "vbskpro2.ocx"
Begin VB.Form frmcarteraasientos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Carga de Minutas contables"
   ClientHeight    =   6525
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10200
   ControlBox      =   0   'False
   Icon            =   "frmcarteraasiento.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6525
   ScaleWidth      =   10200
   Begin VB.CommandButton nuevo0 
      Caption         =   "Command1"
      Height          =   255
      Left            =   6000
      TabIndex        =   54
      Top             =   0
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
      Connect         =   "Provider=MSDASQL.1;Password=25072004;Persist Security Info=True;User ID=lucva;Data Source=contable;Initial Catalog=contablesql"
      OLEDBString     =   "Provider=MSDASQL.1;Password=25072004;Persist Security Info=True;User ID=lucva;Data Source=contable;Initial Catalog=contablesql"
      OLEDBFile       =   ""
      DataSourceName  =   "contable"
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select ccostos.* from ccostos"
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
   Begin MSDataListLib.DataList DataList3 
      Bindings        =   "frmcarteraasiento.frx":0442
      Height          =   1620
      Left            =   840
      TabIndex        =   52
      Top             =   3480
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   2752
      _Version        =   393216
      MatchEntry      =   -1  'True
      BackColor       =   12640511
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
      Left            =   840
      TabIndex        =   51
      Top             =   3120
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
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
      Left            =   720
      TabIndex        =   53
      Top             =   2880
      Visible         =   0   'False
      Width           =   3015
   End
   Begin MSMask.MaskEdBox Masksaldo 
      CausesValidation=   0   'False
      Height          =   255
      Left            =   5400
      TabIndex        =   47
      TabStop         =   0   'False
      Top             =   6000
      Width           =   1575
      _ExtentX        =   2778
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
      TabIndex        =   44
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
      TabIndex        =   46
      TabStop         =   0   'False
      Top             =   6000
      Width           =   1815
      _ExtentX        =   3201
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
      TabIndex        =   45
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
      TabIndex        =   43
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
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   5880
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      DataField       =   "fecharegistro"
      DataSource      =   "datmaestro"
      Enabled         =   0   'False
      Height          =   285
      Left            =   5160
      Locked          =   -1  'True
      TabIndex        =   37
      Top             =   480
      Width           =   1455
   End
   Begin MSDataListLib.DataList DataList2 
      Bindings        =   "frmcarteraasiento.frx":045F
      Height          =   2205
      Left            =   1080
      TabIndex        =   28
      Top             =   3480
      Visible         =   0   'False
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   3889
      _Version        =   393216
      MatchEntry      =   -1  'True
      OLEDragMode     =   1
      OLEDropMode     =   1
      BackColor       =   12640511
      ListField       =   "codigo"
      BoundColumn     =   "Cod Contable"
   End
   Begin VB.CommandButton eliminarmovimiento 
      Caption         =   "Eliminar Movimiento"
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
      Left            =   3720
      TabIndex        =   34
      Top             =   3120
      Width           =   1455
   End
   Begin MSMask.MaskEdBox visual 
      CausesValidation=   0   'False
      Height          =   495
      Left            =   4080
      TabIndex        =   30
      Top             =   1920
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   873
      _Version        =   393216
      BackColor       =   14737632
      ForeColor       =   16711680
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
      Format          =   "#,##0.00;(#,##0.00)"
      PromptChar      =   "_"
   End
   Begin MSDataGridLib.DataGrid DataGrid2 
      Bindings        =   "frmcarteraasiento.frx":0478
      Height          =   255
      Left            =   1200
      TabIndex        =   29
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
   Begin VB.CommandButton nuevomovimiento 
      Caption         =   "Nuevo Movimiento"
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
      Left            =   2040
      TabIndex        =   27
      Top             =   3120
      Width           =   1695
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmcarteraasiento.frx":0491
      Height          =   2175
      Left            =   240
      TabIndex        =   22
      Top             =   3720
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   3836
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
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
      ColumnCount     =   8
      BeginProperty Column00 
         DataField       =   "idasiento"
         Caption         =   "idasiento"
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
         DataField       =   "Fecha"
         Caption         =   "Fecha"
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
      BeginProperty Column02 
         DataField       =   "idcuenta"
         Caption         =   "Cod.Cuenta"
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
      BeginProperty Column03 
         DataField       =   "Debe"
         Caption         =   "Debe"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "Haber"
         Caption         =   "Haber"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "idmasterasientos"
         Caption         =   "idmasterasientos"
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
      BeginProperty Column06 
         DataField       =   "detallefila"
         Caption         =   "Detalle"
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
      BeginProperty Column07 
         DataField       =   "empresa"
         Caption         =   "empresa"
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
            Locked          =   -1  'True
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column01 
            Locked          =   -1  'True
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column02 
            Alignment       =   2
            Locked          =   -1  'True
         EndProperty
         BeginProperty Column03 
            Alignment       =   1
            Locked          =   -1  'True
         EndProperty
         BeginProperty Column04 
            Alignment       =   1
            Locked          =   -1  'True
         EndProperty
         BeginProperty Column05 
            Locked          =   -1  'True
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column06 
            Locked          =   -1  'True
            Object.Visible         =   -1  'True
         EndProperty
         BeginProperty Column07 
            Locked          =   -1  'True
            Object.Visible         =   0   'False
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton grabamovimiento 
      Caption         =   "Grabar Movimiento"
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
      Left            =   600
      TabIndex        =   21
      Top             =   3120
      Width           =   1455
   End
   Begin VB.TextBox Text3 
      DataField       =   "detallefila"
      DataSource      =   "datasiento"
      Height          =   285
      Index           =   6
      Left            =   4560
      TabIndex        =   20
      Top             =   2520
      Width           =   3375
   End
   Begin VB.TextBox Text3 
      DataField       =   "idmasterasientos"
      DataSource      =   "datasiento"
      Height          =   285
      Index           =   8
      Left            =   7680
      TabIndex        =   19
      Top             =   960
      Visible         =   0   'False
      Width           =   375
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
      Left            =   3240
      TabIndex        =   18
      Top             =   2520
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
      Left            =   1800
      TabIndex        =   17
      Top             =   2520
      Width           =   1095
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      DataField       =   "idcuenta"
      DataSource      =   "datasiento"
      Height          =   285
      Left            =   600
      TabIndex        =   16
      Top             =   2520
      Width           =   855
   End
   Begin VB.TextBox Text3 
      DataField       =   "Fecha"
      DataSource      =   "datasiento"
      Height          =   285
      Index           =   2
      Left            =   7320
      TabIndex        =   15
      Top             =   960
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text3 
      DataField       =   "empresa"
      DataSource      =   "datasiento"
      Height          =   285
      Index           =   1
      Left            =   6960
      TabIndex        =   14
      Top             =   960
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text3 
      DataField       =   "idasiento"
      DataSource      =   "datasiento"
      Height          =   285
      Index           =   0
      Left            =   6600
      TabIndex        =   13
      Top             =   960
      Visible         =   0   'False
      Width           =   375
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      DataField       =   "fecha"
      DataSource      =   "datmaestro"
      Height          =   285
      Left            =   2400
      TabIndex        =   1
      Top             =   480
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   503
      _Version        =   393216
      AllowPrompt     =   -1  'True
      Format          =   "dd/mm/yyyy"
      PromptChar      =   "_"
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   6840
      Locked          =   -1  'True
      TabIndex        =   12
      Text            =   "Text2"
      Top             =   240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton grabar 
      Caption         =   "grabar"
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   6000
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      DataField       =   "cerrado"
      DataSource      =   "datmaestro"
      Height          =   285
      Index           =   6
      Left            =   7560
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   600
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text1 
      DataField       =   "empresa"
      DataSource      =   "datmaestro"
      Height          =   285
      Index           =   5
      Left            =   6960
      Locked          =   -1  'True
      TabIndex        =   5
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
      Left            =   5160
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   840
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      DataField       =   "perinicial"
      DataSource      =   "datmaestro"
      Height          =   285
      Index           =   3
      Left            =   3840
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   840
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      DataField       =   "concepto"
      DataSource      =   "datmaestro"
      Height          =   285
      Index           =   2
      Left            =   2400
      TabIndex        =   0
      Top             =   1200
      Width           =   4215
   End
   Begin VB.TextBox asiento 
      Alignment       =   2  'Center
      DataField       =   "nroasiento"
      DataSource      =   "datmaestro"
      Height          =   285
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   840
      Width           =   735
   End
   Begin MSAdodcLib.Adodc datmaestro 
      Height          =   330
      Left            =   240
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
      Connect         =   "Provider=MSDASQL.1;Password=25072004;Persist Security Info=True;User ID=lucva;Data Source=contable;Initial Catalog=contablesql"
      OLEDBString     =   "Provider=MSDASQL.1;Password=25072004;Persist Security Info=True;User ID=lucva;Data Source=contable;Initial Catalog=contablesql"
      OLEDBFile       =   ""
      DataSourceName  =   "contable"
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select [Maestro Asientos].* from [Maestro Asientos]"
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
      Connect         =   "Provider=MSDASQL.1;Password=25072004;Persist Security Info=True;User ID=lucva;Data Source=contable;Initial Catalog=contablesql"
      OLEDBString     =   "Provider=MSDASQL.1;Password=25072004;Persist Security Info=True;User ID=lucva;Data Source=contable;Initial Catalog=contablesql"
      OLEDBFile       =   ""
      DataSourceName  =   "contable"
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select EMPRESA.* from EMPRESA"
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
      Connect         =   "Provider=MSDASQL.1;Password=25072004;Persist Security Info=True;User ID=lucva;Data Source=contable;Initial Catalog=contablesql"
      OLEDBString     =   "Provider=MSDASQL.1;Password=25072004;Persist Security Info=True;User ID=lucva;Data Source=contable;Initial Catalog=contablesql"
      OLEDBFile       =   ""
      DataSourceName  =   "contable"
      OtherAttributes =   ""
      UserName        =   "lucva"
      Password        =   "25072004"
      RecordSource    =   "select [Detalle Asientos].* from [Detalle Asientos]"
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
      Connect         =   "Provider=MSDASQL.1;Password=25072004;Persist Security Info=True;User ID=lucva;Data Source=contable;Initial Catalog=contablesql"
      OLEDBString     =   "Provider=MSDASQL.1;Password=25072004;Persist Security Info=True;User ID=lucva;Data Source=contable;Initial Catalog=contablesql"
      OLEDBFile       =   ""
      DataSourceName  =   "contable"
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select listacuentas.* from listacuentas ORDER BY IDCUENTA"
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
   Begin VB.CommandButton nuevo 
      Caption         =   "&Grabar Asiento"
      Height          =   735
      Left            =   8760
      Picture         =   "frmcarteraasiento.frx":04AA
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   480
      UseMaskColor    =   -1  'True
      Width           =   855
   End
   Begin VB.CommandButton salir 
      Caption         =   "&Cerrar"
      Height          =   735
      Left            =   8760
      Picture         =   "frmcarteraasiento.frx":09DC
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   2640
      UseMaskColor    =   -1  'True
      Width           =   855
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
      Height          =   3495
      Left            =   8400
      TabIndex        =   40
      Top             =   120
      Width           =   1575
      Begin VB.CommandButton Ver 
         Caption         =   "&Ver y/o Modificar"
         Height          =   735
         Left            =   360
         Picture         =   "frmcarteraasiento.frx":0E1E
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   1440
         UseMaskColor    =   -1  'True
         Width           =   855
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
      Connect         =   "Provider=MSDASQL.1;Password=25072004;Persist Security Info=True;User ID=lucva;Data Source=contable;Initial Catalog=contablesql"
      OLEDBString     =   "Provider=MSDASQL.1;Password=25072004;Persist Security Info=True;User ID=lucva;Data Source=contable;Initial Catalog=contablesql"
      OLEDBFile       =   ""
      DataSourceName  =   "contable"
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select listaccostos.* from listaccostos"
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
      Top             =   0
      _ExtentX        =   1270
      _ExtentY        =   1270
      CloseButtonToolTipText=   "Cerrar"
      MinButtonToolTipText=   "Minimizar"
      MaxButtonToolTipText=   "Maximizar"
      OldForeColor    =   0
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
      LcK1            =   "3.66*/4/0*/1-5*210/."
      LcK2            =   $"frmcarteraasiento.frx":1350
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
      Connect         =   "Provider=MSDASQL.1;Password=25072004;Persist Security Info=True;User ID=lucva;Data Source=contable;Initial Catalog=contablesql"
      OLEDBString     =   "Provider=MSDASQL.1;Password=25072004;Persist Security Info=True;User ID=lucva;Data Source=contable;Initial Catalog=contablesql"
      OLEDBFile       =   ""
      DataSourceName  =   "contable"
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select [Maestro Asientos].* from [Maestro Asientos]"
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
      TabIndex        =   50
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
      TabIndex        =   49
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
      TabIndex        =   48
      Top             =   6240
      Width           =   615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha de registro:"
      Height          =   255
      Index           =   12
      Left            =   3840
      TabIndex        =   36
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cabecera de Asiento"
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
      Height          =   255
      Index           =   11
      Left            =   360
      TabIndex        =   35
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "Egresos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   10
      Left            =   3240
      TabIndex        =   33
      Top             =   2805
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "Ingresos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   9
      Left            =   1800
      TabIndex        =   32
      Top             =   2805
      Width           =   1095
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      Visible         =   0   'False
      X1              =   1440
      X2              =   1920
      Y1              =   2160
      Y2              =   2520
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      Visible         =   0   'False
      X1              =   3000
      X2              =   3480
      Y1              =   2520
      Y2              =   2160
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Movimientos"
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
      Height          =   255
      Index           =   8
      Left            =   360
      TabIndex        =   31
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "DETALLE"
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
      Index           =   7
      Left            =   5520
      TabIndex        =   26
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "HABER"
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
      Index           =   6
      Left            =   3240
      TabIndex        =   25
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "DEBE"
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
      Index           =   5
      Left            =   1800
      TabIndex        =   24
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Nº Cuenta"
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
      Index           =   4
      Left            =   600
      TabIndex        =   23
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Periodo:"
      Height          =   255
      Index           =   3
      Left            =   3240
      TabIndex        =   10
      Top             =   840
      Width           =   615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Referencia:"
      Height          =   255
      Index           =   2
      Left            =   1080
      TabIndex        =   9
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Nro. Asiento:"
      Height          =   255
      Index           =   1
      Left            =   1080
      TabIndex        =   8
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha del Hecho:"
      Height          =   255
      Index           =   0
      Left            =   1080
      TabIndex        =   7
      Top             =   480
      Width           =   1335
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      Height          =   1695
      Left            =   240
      Top             =   1920
      Width           =   7935
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      Height          =   1455
      Left            =   240
      Top             =   240
      Width           =   7935
   End
End
Attribute VB_Name = "frmcarteraasientos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public movimientohijo As Double
Dim detalle As String
Dim posicion As Double
Dim numerodisco As Double


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

Private Sub DataList2_Click()
    Text4.Text = DataList2.BoundText
End Sub

Private Sub DataList2_GotFocus()
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
End Sub

Private Sub DataList2_KeyPress(KeyAscii As Integer)

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
    

End Sub

Private Sub DataList2_LostFocus()

    DataList2.Visible = False
                     
End Sub

Private Sub DataList3_Click()
    Text9.Text = DataList3.BoundText
End Sub

Private Sub DataList3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Text9.Text = DataList3.BoundText
        Text3(4).SetFocus
    End If
End Sub

Private Sub DataList3_LostFocus()
    If Text9.Text = "" Then
        mensa = MsgBox("Debe ingresa un Centro de Costo", vbCritical, "!Error¡")
        DataList3.SetFocus
        Exit Sub
    End If
Frame1.Visible = False
Text9.Visible = False
DataList3.Visible = False
End Sub

Private Sub eliminarmovimiento_Click()
On Error GoTo erroreliminar


    mensa = MsgBox("Esta por eliminar un movimiento de este asiento, esta seguro", vbYesNo, "!! Atención !!")
    If mensa = vbYes Then
            datasiento.Recordset.Delete
    End If

erroreliminar:


End Sub

Private Sub Form_Load()

    Dim fs, d, t
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set d = fs.GetDrive(fs.GetDriveName(fs.GetAbsolutePathName(drvpath)))
    s = d.serialnumber
    numerodisco = s

    Inicio.Toolbar1.Visible = True
    DataGrid1.Columns(4).NumberFormat = "#,##0.00"
    DataGrid1.Columns(3).NumberFormat = "#,##0.00"

    datcuentas.RecordSource = "select listacuentas.* from listacuentas WHERE empre = " & login.empresaact & " and inicioper = '" & login.iper & "' and finper = '" & login.fper & "' ORDER BY IDCUENTA"
    datcuentas.Refresh
    datmaestro.RecordSource = "select [Maestro Asientos].* from [Maestro Asientos] where empresa = " & login.empresaact & " and perinicial = '" & login.iper & "' order by nroasiento"
    datmaestro.Refresh
    datmaestro1.RecordSource = "select [Maestro Asientos].* from [Maestro Asientos] where empresa = " & login.empresaact & " order by nroasiento"
    datmaestro1.Refresh
    datperiodo.RecordSource = "select EMPRESA.* from EMPRESA where empresa = " & login.empresaact & ""
    datperiodo.Refresh
    datasiento.RecordSource = "select [Detalle Asientos].* from [Detalle Asientos] where empresa = " & login.empresaact & " order by idmasterasientos"
    datasiento.Refresh
    datccostos.RecordSource = "SELECT ccostos.* FROM ccostos WHERE empresa = " & login.empresaact & ""
    datccostos.Refresh
    MaskEdBox1.Mask = ""
    Text6(2).Text = 0
    Call nuevo0_Click
           
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Inicio.Toolbar1.Visible = False
End Sub

Private Sub grabamovimiento_Click()
On Error GoTo errorasiento

    detalle = Text3(6).Text
    datasiento.Recordset.UpdateBatch adAffectCurrent
    
    If Text2.Text <> "" Then
    movimientohijo = Text2.Text
Else
    movimientohijo = 0
End If
   
    datasiento.RecordSource = "select [Detalle Asientos].* from [Detalle Asientos] where idmasterasientos = " & movimientohijo & " "
    datasiento.Refresh
    
    Text6(0).Text = 0
    Text6(1).Text = 0
    Text6(2).Text = 0
    datasiento.Recordset.MoveFirst
paso1:
    Text6(0).Text = datasiento.Recordset.Fields(3) + Text6(0)
    Text6(1).Text = datasiento.Recordset.Fields(4) + Text6(1)
    Maskdebe = Text6(0)
    Maskhaber = Text6(1)
    datasiento.Recordset.MoveNext
    If datasiento.Recordset.EOF = False Then
        GoTo paso1
    Else
        GoTo paso2
    End If
paso2:
    Text6(2).Text = Text6(0) - Text6(1)
    Masksaldo = Text6(2)

    nuevomovimiento.SetFocus
Exit Sub
errorasiento:
    mensa = MsgBox("Debe ingresar una referencia y presionar enter para grabar los movimientos", vbCritical, "Atencion !!")
    Text1(2).SetFocus
End Sub

Private Sub grabar_Click()

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
            Exit Sub
    End If
    
    posicion = datmaestro.Recordset.AbsolutePosition
 Rem   datmaestro.Recordset.Fields(11) = "N"
 Rem   datmaestro.Recordset.UpdateBatch adAffectCurrent
 Rem   datmaestro.Refresh
 Rem   datmaestro.Recordset.AbsolutePosition = posicion
    
   datasiento.Recordset.AddNew

   Text3(1).Text = login.empresaact
   Text3(2).Text = MaskEdBox1.Text
   Text3(8).Text = Text2.Text
   Text5.Text = Date
   
    Text4.SetFocus

End Sub




Private Sub movimientos_DragDrop(Source As Control, x As Single, Y As Single)

End Sub

Private Sub MaskEdBox1_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Or KeyAscii = 9 Then
        KeyAscii = 0
        Text1(2).SetFocus
    End If
    
    
End Sub



Private Sub nuevo1_Click()

End Sub

Private Sub nuevo_Click()
    Dim ultimoasiento As Double

     If Text6(2).Text <> 0 Then
          mensa = MsgBox("EL Asiento está Desvalanceado, no puede grabar", vbCritical, "!! Error !!")
          datasiento.Recordset.MoveLast
          Text3(4).SetFocus
          Exit Sub
     Else
        frmcartera.datentrada.Recordset.Fields("cartera") = "N"
        frmcartera.datentrada.Recordset.UpdateBatch adAffectCurrent
        datmaestro1.RecordSource = "select [Maestro Asientos].* from [Maestro Asientos] where empresa = " & login.empresaact & " order by nroasiento"
        datmaestro1.Refresh
        If datmaestro1.Recordset.EOF = True Then
            ultimoasiento = 1
        Else
            datmaestro1.Recordset.MoveLast
            ultimoasiento = datmaestro1.Recordset.Fields(3) + 1
            asiento.Text = ultimoasiento
        End If
        datmaestro.Recordset.Fields(11) = "S"
        datmaestro.Recordset.UpdateBatch adAffectCurrent
        datasiento.RecordSource = "select [Detalle Asientos].* from [Detalle Asientos] where empresa = " & login.empresaact & " and idmasterasientos = " & Text2.Text & ""
        datasiento.Refresh
        datasiento.Recordset.MoveFirst
pas00:
        datasiento.Recordset.Fields("idmasterasientos") = datmaestro.Recordset.Fields("idmasterasientos")
        datasiento.Recordset.MoveNext
        If datasiento.Recordset.EOF = True Then
            nuevo0_Click
            MaskEdBox1.SetFocus
            Exit Sub
        End If
        GoTo pas00
     End If
     

End Sub

Private Sub nuevo0_Click()
            
    If datmaestro.Recordset.EOF = True Then Exit Sub
    If datmaestro1.Recordset.EOF = True Then
        ultimoasiento = 1
    Else
        datmaestro1.Recordset.MoveLast
        ultimoasiento = datmaestro1.Recordset.Fields(3) + 1
    End If

    datmaestro.Recordset.AddNew
    If datasiento.Recordset.EOF = True Then
        Text2.Text = numerodisco
    Else
        datasiento.Recordset.MoveLast
        Text2.Text = numerodisco + datasiento.Recordset.Fields("idmasterasientos")
    End If
    datasiento.RecordSource = "select [Detalle Asientos].* from [Detalle Asientos] where empresa = " & login.empresaact & " and idmasterasientos = " & Text2.Text & " order by idmasterasientos"
    datasiento.Refresh
    MaskEdBox1.Mask = "##/##/####"
    MaskEdBox1.Text = Date
    asiento.Text = ultimoasiento
    Text1(5).Text = login.empresaact
    Text1(6).Text = "N"
    Text1(3).Text = login.iper
    Text1(4).Text = login.fper
    Text5.Text = Date
    Text1(2).Text = "Baja de Cartera cheque:" + frmcartera.DataGrid1.Columns(8).Text
    
End Sub

Private Sub nuevomovimiento_Click()
    
    
   datasiento.Recordset.AddNew

   Text3(1).Text = login.empresaact
   Text3(2).Text = MaskEdBox1.Text
   Text3(8).Text = Text2.Text
   Text3(6).Text = detalle
   Text4.SetFocus
   
End Sub

Private Sub salir_Click()
On Error GoTo errorsalir


     If Text6(2).Text <> 0 Then
          mensa = MsgBox("EL Asiento está Desvalanceado, no puede grabar", vbCritical, "!! Error !!")
          datasiento.Recordset.MoveLast
          Text3(4).SetFocus
          Exit Sub
     End If
     
         Unload Me
    Exit Sub
    
errorsalir:

End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)

    If KeyAscii = 13 Or KeyAscii = 9 Then
        KeyAscii = 0
        Call grabar_Click
            Text4.Text = frmcartera.DataGrid1.Columns(12).Text
            Text3(4).Text = 0
            Text3(5).Text = frmcartera.DataGrid1.Columns(11).Value
            Text3(6).Text = "Cheque"
            Call grabamovimiento_Click
            Call nuevomovimiento_Click
    End If

End Sub

Private Sub Text1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)

     If KeyCode = 38 Then MaskEdBox1.SetFocus

End Sub

Private Sub Text3_Change(Index As Integer)
    visual.Text = Text3(Index).Text
End Sub

Private Sub Text3_GotFocus(Index As Integer)

    If Text4.Text = "" Then
        mensa = MsgBox("Cuenta no existente", vbCritical, "!! Atencion !!")
        Text4.SetFocus
        Exit Sub
    End If

        If Index = 4 Or Index = 5 Then
              visual.Text = Text3(Index).Text
              visual.Visible = True
              Line1.Visible = True
              Line2.Visible = True
              visual.Left = Text3(Index).Left - Text3(Index).Width / 2
              Line1.X1 = visual.Left
              Line1.X2 = Text3(Index).Left
              Line1.Y1 = visual.Top + visual.Height
              Line1.Y2 = Text3(Index).Top
              Line2.X1 = visual.Left + visual.Width
              Line2.X2 = Text3(Index).Left + Text3(Index).Width
              Line2.Y1 = visual.Top + visual.Height
              Line2.Y2 = Text3(Index).Top
        Else
             visual.Visible = False
             Line1.Visible = False
             Line2.Visible = False
        End If
                     
        Text3(Index).SelLength = Len(Text3(Index).Text)

        DataList2.Visible = False

End Sub



Private Sub Text3_KeyPress(Index As Integer, KeyAscii As Integer)
 

    If KeyAscii = 13 Or KeyAscii = 9 Then
        KeyAscii = 0
        If Index = 6 Then
            grabamovimiento.SetFocus
            Exit Sub
        End If
        Text3(Index + 1).SetFocus
    End If
  
End Sub

Private Sub Text3_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
     
     If KeyCode = 38 And Index > 4 Then Text3(Index - 1).SetFocus
     If KeyCode = 38 And Index = 4 Then Text4.SetFocus

End Sub

Private Sub Text3_LostFocus(Index As Integer)
        If Index = 4 Then
            If Text3(Index).Text = "" Then Text3(Index).Text = 0
            If Text3(Index).Text <> "0" Then
                Text3(Index + 1).Text = "0"
            End If
        End If
        If Index = 5 Then
            If Text3(Index).Text <> "0" Then
                Text3(Index - 1).Text = "0"
            End If
        End If
End Sub


Private Sub Text4_GotFocus()

    Text4.SelLength = Len(Text4)
    DataList2.BoundText = Text4.Text
    DataList2.Visible = True
    DataList2.Left = Text4.Left
    DataList2.Top = Text4.Top + Text4.Height
    DataList2.SetFocus
                  
End Sub


Private Sub Ver_Click()

     If Text6(2).Text <> 0 Then
          mensa = MsgBox("EL Asiento está Desvalanceado, no puede grabar", vbCritical, "!! Error !!")
          datasiento.Recordset.MoveLast
          Text3(4).SetFocus
          Exit Sub
     End If
    
    frmasientosbusca.Show
    Unload Me
    
End Sub
