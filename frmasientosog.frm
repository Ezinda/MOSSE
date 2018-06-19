VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{5734474E-78D3-4254-99B9-C35F31BDF509}#62.0#0"; "vbskpro2.ocx"
Begin VB.Form frmasientosog 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Carga de Asientos"
   ClientHeight    =   4770
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7365
   ControlBox      =   0   'False
   Icon            =   "frmasientosog.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4770
   ScaleWidth      =   7365
   Begin VB.CommandButton aceptar 
      Caption         =   "&Aceptar"
      Height          =   615
      Left            =   5760
      Picture         =   "frmasientosog.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   1080
      UseMaskColor    =   -1  'True
      Width           =   1215
   End
   Begin VB.TextBox Text9 
      Alignment       =   2  'Center
      DataField       =   "ccosto"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Left            =   480
      TabIndex        =   33
      Top             =   2040
      Visible         =   0   'False
      Width           =   615
   End
   Begin MSDataListLib.DataList DataList3 
      Bindings        =   "frmasientosog.frx":0884
      Height          =   1620
      Left            =   480
      TabIndex        =   32
      Top             =   2400
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
   Begin VB.Frame Frame2 
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
      Left            =   240
      TabIndex        =   34
      Top             =   1800
      Visible         =   0   'False
      Width           =   3015
   End
   Begin MSDataListLib.DataList DataList2 
      Bindings        =   "frmasientosog.frx":08A1
      Height          =   2205
      Left            =   1320
      TabIndex        =   31
      Top             =   1800
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
      Left            =   6000
      Locked          =   -1  'True
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   4080
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
      Index           =   1
      Left            =   6120
      Locked          =   -1  'True
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   4080
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
      Index           =   2
      Left            =   6360
      Locked          =   -1  'True
      TabIndex        =   26
      TabStop         =   0   'False
      Text            =   " "
      Top             =   4080
      Visible         =   0   'False
      Width           =   855
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
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   1320
      UseMaskColor    =   -1  'True
      Width           =   1455
   End
   Begin MSMask.MaskEdBox visual 
      CausesValidation=   0   'False
      Height          =   495
      Left            =   4080
      TabIndex        =   20
      Top             =   120
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
      Bindings        =   "frmasientosog.frx":08BA
      Height          =   255
      Left            =   1200
      TabIndex        =   19
      Top             =   4080
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
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   1320
      UseMaskColor    =   -1  'True
      Width           =   1695
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmasientosog.frx":08D3
      Height          =   2175
      Left            =   240
      TabIndex        =   13
      Top             =   1920
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   3836
      _Version        =   393216
      AllowUpdate     =   -1  'True
      AllowArrows     =   -1  'True
      Enabled         =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      TabAction       =   1
      WrapCellPointer =   -1  'True
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
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   1320
      UseMaskColor    =   -1  'True
      Width           =   1455
   End
   Begin VB.TextBox Text3 
      DataField       =   "detallefila"
      DataSource      =   "datasiento"
      Height          =   285
      Index           =   6
      Left            =   4560
      MaxLength       =   50
      TabIndex        =   6
      Top             =   720
      Width           =   2415
   End
   Begin VB.TextBox Text3 
      DataField       =   "idmasterasientos"
      DataSource      =   "datasiento"
      Height          =   285
      Index           =   8
      Left            =   1320
      TabIndex        =   11
      Top             =   360
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
      TabIndex        =   5
      Top             =   720
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
      Enabled         =   0   'False
      Height          =   285
      Index           =   4
      Left            =   1800
      TabIndex        =   10
      Top             =   720
      Width           =   1095
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      DataField       =   "idcuenta"
      DataSource      =   "datasiento"
      Height          =   285
      Left            =   600
      TabIndex        =   9
      Top             =   720
      Width           =   855
   End
   Begin VB.TextBox Text3 
      DataField       =   "Fecha"
      DataSource      =   "datasiento"
      Height          =   285
      Index           =   2
      Left            =   960
      TabIndex        =   8
      Top             =   360
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text3 
      DataField       =   "empresa"
      DataSource      =   "datasiento"
      Height          =   285
      Index           =   1
      Left            =   600
      TabIndex        =   7
      Top             =   360
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text3 
      DataField       =   "idasiento"
      DataSource      =   "datasiento"
      Height          =   285
      Index           =   0
      Left            =   240
      TabIndex        =   4
      Top             =   360
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text2 
      DataField       =   "idmasterasientos"
      DataSource      =   "datmaestro"
      Height          =   285
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   3
      Text            =   "Text2"
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton grabar 
      Caption         =   "grabar"
      Height          =   375
      Left            =   4320
      TabIndex        =   2
      Top             =   3840
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      DataField       =   "cerrado"
      DataSource      =   "datmaestro"
      Height          =   285
      Index           =   6
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text1 
      DataField       =   "empresa"
      DataSource      =   "datmaestro"
      Height          =   285
      Index           =   5
      Left            =   600
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
   Begin MSAdodcLib.Adodc datmaestro 
      Height          =   330
      Left            =   240
      Top             =   3840
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
      Top             =   3840
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
      Top             =   3840
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
   Begin MSAdodcLib.Adodc datcuentas 
      Height          =   330
      Left            =   6000
      Top             =   3960
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
      Left            =   5640
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   4200
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
   Begin MSMask.MaskEdBox Maskhaber 
      CausesValidation=   0   'False
      Height          =   255
      Left            =   3840
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   4200
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
      Left            =   2040
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   4200
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
   Begin MSAdodcLib.Adodc datccostos 
      Height          =   330
      Left            =   6000
      Top             =   3960
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
      Left            =   240
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
      LcK2            =   $"frmasientosog.frx":08EC
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
   Begin MSAdodcLib.Adodc datlistacostos 
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
      Left            =   5880
      TabIndex        =   37
      Top             =   4440
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
      Left            =   4080
      TabIndex        =   36
      Top             =   4440
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
      Left            =   2280
      TabIndex        =   35
      Top             =   4440
      Width           =   615
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
      TabIndex        =   23
      Top             =   1005
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
      TabIndex        =   22
      Top             =   1005
      Width           =   1095
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      Visible         =   0   'False
      X1              =   1440
      X2              =   1920
      Y1              =   360
      Y2              =   720
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      Visible         =   0   'False
      X1              =   3000
      X2              =   3480
      Y1              =   720
      Y2              =   360
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
      TabIndex        =   21
      Top             =   0
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
      TabIndex        =   17
      Top             =   480
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
      TabIndex        =   16
      Top             =   480
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
      TabIndex        =   15
      Top             =   480
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
      TabIndex        =   14
      Top             =   480
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000F&
      BackStyle       =   1  'Opaque
      Height          =   1695
      Left            =   240
      Top             =   120
      Width           =   6975
   End
End
Attribute VB_Name = "frmasientosog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public movimientohijo As Double
Public inicioper As Date
Public finper As Date
Dim posicion As Double


Private Sub asiento_Change()
On Error GoTo fuera

If Text2.Text <> "" Then
    movimientohijo = Text2.Text
Else
    movimientohijo = 0
End If

    
    datasiento.RecordSource = "select [Detalle Asientos].* from [Detalle Asientos] where idmasterasientos = " & movimientohijo & " and haber <> 0 order by idasiento "
    datasiento.Refresh

    Text6(0).Text = 0
    Text6(1).Text = 0
    Text6(2).Text = 0
    If datasiento.Recordset.EOF = False Then
            datasiento.Recordset.MoveFirst
    Else
        GoTo paso2
    End If
paso1:
    If datasiento.Recordset.EOF = True Then GoTo paso2
    If IsNull(datasiento.Recordset.Fields(3)) = True Then datasiento.Recordset.Fields(3) = 0
    If IsNull(datasiento.Recordset.Fields(4)) = True Then datasiento.Recordset.Fields(4) = 0
    Text6(0).Text = datasiento.Recordset.Fields(3) + Text6(0).Text
    Text6(1).Text = datasiento.Recordset.Fields(4) + Text6(1).Text
    If datasiento.Recordset.EOF = False Then
        datasiento.Recordset.MoveNext
        GoTo paso1
    End If
paso2:
  Rem  Maskdebe.Text = Text6(0).Text
    Maskhaber.Text = Text6(1).Text
    Masksaldo.Text = Text6(1).Text - Text6(0).Text
  
fuera:
End Sub



Private Sub Command1_Click()



End Sub


Private Sub Cancelar_Click()

  Rem  datmaestro.Recordset.Delete
    datmaestro.Refresh

End Sub


Private Sub aceptar_Click()
On Error GoTo fuera

    If Masksaldo.Text <> "0" Then
        mensa = MsgBox("No es posible grabar, el asiento esta desbalanceado", vbCritical, "Atención")
        DataGrid1.SetFocus
        Exit Sub
    End If
    Unload Me

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
        Text3(5).SetFocus
    End If
    
fuera:
End Sub

Private Sub DataList2_LostFocus()

    DataList2.Visible = False

End Sub


Private Sub desde_Change()
On Error GoTo fuera
    
    If desde < login.iper Then desde = login.iper
    Call Ver_Click
    
fuera:
End Sub

Private Sub eliminarmovimiento_Click()
On Error GoTo erroreliminar
    mensa = MsgBox("Esta por eliminar un movimiento de este asiento, esta seguro", vbYesNo, "!! Atención !!")
    If mensa = vbYes Then
            If datasiento.Recordset.EOF = False Then datasiento.Recordset.Delete
    End If
erroreliminar:
  mensa = MsgBox("No se puede eliminar", vbCritical, "Atención")
  Unload Me
  
End Sub

Private Sub Form_Load()
    
    Inicio.Toolbar1.Visible = True
    
datasiento.ConnectionString = login.conexiontotal
datccostos.ConnectionString = login.conexiontotal
datcuentas.ConnectionString = login.conexiontotal
datmaestro.ConnectionString = login.conexiontotal
datperiodo.ConnectionString = login.conexiontotal
datlistacostos.ConnectionString = login.conexiontotal
          
    datlistacostos.RecordSource = "select listaccostos.* from listaccostos WHERE empresa = " & login.empresaact & ""
    datlistacostos.Refresh
    datcuentas.RecordSource = "select listacuentas.* from listacuentas WHERE empre = " & login.empresaact & " and inicioper = '" & login.iper & "' and finper = '" & login.fper & "' ORDER BY IDCUENTA"
    datcuentas.Refresh
    
    datmaestro.RecordSource = "select [Maestro Asientos].* from [Maestro Asientos] where empresa = " & login.empresaact & " and perinicial = '" & login.iper & "' and nroasiento = " & frmotrosgastos.asientominuta & " "
    datmaestro.Refresh
    movimientohijo = datmaestro.Recordset.Fields("idmasterasientos")
    datasiento.RecordSource = "select [Detalle Asientos].* from [Detalle Asientos] where idmasterasientos = " & movimientohijo & " and haber <> 0 order by idasiento"
    datasiento.Refresh
    Maskdebe.Text = frmotrosgastos.debeminuta

            
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Inicio.Toolbar1.Visible = False

End Sub

Private Sub grabamovimiento_Click()
On Error GoTo fuera

    datasiento.Recordset.UpdateBatch adAffectCurrent
    
    If Text2.Text <> "" Then
    movimientohijo = datmaestro.Recordset.Fields("idmasterasientos")
Else
    movimientohijo = 0
End If
   
    datasiento.RecordSource = "select [Detalle Asientos].* from [Detalle Asientos] where idmasterasientos = " & movimientohijo & " and haber <> 0 order by idasiento "
    datasiento.Refresh
    
    Text6(0).Text = 0
    Text6(1).Text = 0
    Text6(2).Text = 0
    datasiento.Recordset.MoveFirst
paso1:
    Text6(1).Text = datasiento.Recordset.Fields(4) + Text6(1)
    Maskhaber = Text6(1)
    datasiento.Recordset.MoveNext
    If datasiento.Recordset.EOF = False Then
        GoTo paso1
    Else
        GoTo paso2
    End If
paso2:
    Text6(2).Text = Maskdebe - Text6(1)
    Masksaldo = Text6(2)

    nuevomovimiento.SetFocus
    nuevomovimiento.SetFocus

fuera:
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
            Exit Sub
    End If
    
    posicion = datmaestro.Recordset.AbsolutePosition
       
    datmaestro.Recordset.UpdateBatch adAffectCurrent
    datmaestro.Refresh
    datmaestro.Recordset.AbsolutePosition = posicion
    
   datasiento.Recordset.AddNew

   Text3(1).Text = login.empresaact
   Text3(2).Text = MaskEdBox1.Text
   Text3(8).Text = Text2.Text
   Text5.Text = Date
   Text4.SetFocus
    
fuera:
End Sub

Private Sub hasta_Change()
On Error GoTo fuera

    If hasta > login.fper Then hasta = login.fper
    Call Ver_Click
    
fuera:
End Sub
Private Sub movimientos_DragDrop(Source As Control, x As Single, Y As Single)

End Sub

Private Sub MaskEdBox1_KeyPress(KeyAscii As Integer)
On Error GoTo fuera

    If KeyAscii = 13 Or KeyAscii = 9 Then
        KeyAscii = 0
        Text1(2).SetFocus
    End If
    
fuera:
End Sub



Private Sub nuevo1_Click()

End Sub

Private Sub nuevo_Click()
On Error GoTo fuera

    Dim ultimoasiento As Double

    If datmaestro.Recordset.EOF = True Then
        ultimoasiento = 1
    Else
        datmaestro.Recordset.MoveLast
        ultimoasiento = datmaestro.Recordset.Fields(3) + 1
    End If

    datmaestro.Recordset.AddNew
    MaskEdBox1.Mask = "##/##/####"
    MaskEdBox1.SelLength = 10
    MaskEdBox1.SelText = ""
    asiento.Text = ultimoasiento
    Text1(5).Text = login.empresaact
    Text1(6).Text = "N"
    Text1(3).Text = datperiodo.Recordset.Fields(8)
    Text1(4).Text = datperiodo.Recordset.Fields(9)
    Text5.Text = Date
    
fuera:
End Sub


Private Sub nuevomovimiento_Click()
On Error GoTo fuera
    
   detalle = Text3(6).Text
   datasiento.Recordset.AddNew

   Text3(1).Text = login.empresaact
   Text3(8).Text = Text2.Text
   Text3(6).Text = detalle
   Text4.SetFocus
   
fuera:
End Sub

Private Sub ordenarfecha_Click()
On Error GoTo fuera

    datmaestro.RecordSource = "select [Maestro Asientos].* from [Maestro Asientos] where (fecha >= '" & frmasientosbusca.inicioper & "') and (fecha <= '" & frmasientosbusca.finper & "') and (empresa = " & login.empresaact & ") order by fecha"
    datmaestro.Refresh

fuera:
End Sub

Private Sub ordenasiento_Click()
On Error GoTo fuera

    datmaestro.RecordSource = "select [Maestro Asientos].* from [Maestro Asientos] where (fecha >= '" & frmasientosbusca.inicioper & "') and (fecha <= '" & frmasientosbusca.finper & "') and (empresa = " & login.empresaact & ") order by nroasiento"
    datmaestro.Refresh

fuera:
End Sub

Private Sub salir_Click()

    Unload Me

End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
On Error GoTo fuera

    If KeyAscii = 13 Or KeyAscii = 9 Then
        KeyAscii = 0
        Call grabar_Click
    End If

fuera:
End Sub

Private Sub Text1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)

     If KeyCode = 38 Then MaskEdBox1.SetFocus

End Sub

Private Sub Text3_Change(Index As Integer)
On Error GoTo fuera

    visual.Text = Text3(Index).Text
    
fuera:
End Sub

Private Sub Text3_GotFocus(Index As Integer)
On Error GoTo fuera

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
On Error GoTo fuera
     
     If KeyCode = 38 And Index > 4 Then Text3(Index - 1).SetFocus
     If KeyCode = 38 And Index = 4 Then Text4.SetFocus

fuera:
End Sub

Private Sub Text3_LostFocus(Index As Integer)
On Error GoTo errormod
        If Index = 4 Then
            If Text3(Index).Text = "" Then Text3(Index).Text = 0
            If Text3(Index).Text <> 0 Then
                Text3(Index + 1).Text = 0
            End If
        End If
        If Index = 5 Then
            If Text3(Index).Text <> 0 Then
                Text3(Index - 1).Text = 0
            End If
        End If
errormod:
End Sub


Private Sub Text4_GotFocus()
On Error GoTo fuera

    Text4.SelLength = Len(Text4)
    DataList2.BoundText = Text4.Text
    DataList2.Visible = True
    DataList2.Left = Text4.Left
    DataList2.Top = Text4.Top + Text4.Height
    DataList2.SetFocus
    
fuera:
End Sub

Private Sub Ver_Click()
On Error GoTo fuera

    inicioper = desde.Value
    finper = hasta.Value

    datmaestro.RecordSource = "select [Maestro Asientos].* from [Maestro Asientos] where (fecha >= '" & frmasientosbusca.inicioper & "') and (fecha <= '" & frmasientosbusca.finper & "') and (empresa = " & login.empresaact & ") order by fecha"
    datmaestro.Refresh
    movimientohijo = datmaestro.Recordset.Fields("idmasterasientos")
    datasiento.RecordSource = "select [Detalle Asientos].* from [Detalle Asientos] where idmasterasientos = " & movimientohijo & " and haber <> 0 order by idasiento "
    datasiento.Refresh

fuera:
End Sub

