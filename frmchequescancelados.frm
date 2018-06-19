VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{5734474E-78D3-4254-99B9-C35F31BDF509}#62.0#0"; "vbskpro2.ocx"
Object = "{D18BBD1F-82BB-4385-BED3-E9D31A3E361E}#1.0#0"; "KewlButtonz.ocx"
Begin VB.Form frmchequescancelados 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cheques Cancelados"
   ClientHeight    =   7755
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13200
   Icon            =   "frmchequescancelados.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7755
   ScaleWidth      =   13200
   Begin VB.CommandButton llena 
      Caption         =   "llena"
      Height          =   255
      Left            =   11880
      TabIndex        =   64
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton grillaref 
      Caption         =   "grillaref"
      Height          =   375
      Left            =   11760
      TabIndex        =   63
      Top             =   480
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Ordenar x Importe"
      Height          =   255
      Left            =   10680
      TabIndex        =   62
      Top             =   1680
      Width           =   1815
   End
   Begin VB.CommandButton Cuenta 
      Caption         =   "Fecha Registro:"
      Height          =   255
      Index           =   9
      Left            =   3720
      Picture         =   "frmchequescancelados.frx":0442
      TabIndex        =   55
      TabStop         =   0   'False
      Top             =   600
      UseMaskColor    =   -1  'True
      Width           =   1455
   End
   Begin MSComctlLib.ProgressBar bar1 
      Height          =   255
      Left            =   240
      TabIndex        =   43
      Top             =   6240
      Visible         =   0   'False
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.TextBox Text9 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   2520
      TabIndex        =   40
      Top             =   3600
      Visible         =   0   'False
      Width           =   615
   End
   Begin MSDataListLib.DataList DataList3 
      Bindings        =   "frmchequescancelados.frx":0974
      Height          =   1620
      Left            =   2520
      TabIndex        =   39
      Top             =   3960
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   2752
      _Version        =   393216
      MatchEntry      =   -1  'True
      BackColor       =   -2147483629
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
      BackColor       =   &H80000003&
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
      Left            =   2280
      TabIndex        =   41
      Top             =   3360
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.CommandButton ordenarfecha 
      Caption         =   "Ordenar x Fecha"
      Height          =   255
      Left            =   7320
      TabIndex        =   38
      Top             =   1680
      Width           =   1575
   End
   Begin VB.CommandButton ordenasiento 
      Caption         =   "Ordenar x NºCheque"
      Height          =   255
      Left            =   8880
      TabIndex        =   37
      Top             =   1680
      Width           =   1815
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
      TabIndex        =   33
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
      Index           =   1
      Left            =   6000
      Locked          =   -1  'True
      TabIndex        =   32
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
      Index           =   2
      Left            =   6840
      Locked          =   -1  'True
      TabIndex        =   29
      TabStop         =   0   'False
      Text            =   " "
      Top             =   5880
      Visible         =   0   'False
      Width           =   855
   End
   Begin MSDataGridLib.DataGrid DataGrid3 
      Bindings        =   "frmchequescancelados.frx":0991
      Height          =   5655
      Left            =   7320
      TabIndex        =   1
      Top             =   1920
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   9975
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   -2147483634
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
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      DataField       =   "fecharegistro"
      DataSource      =   "datmaestro"
      Enabled         =   0   'False
      Height          =   285
      Left            =   5280
      Locked          =   -1  'True
      TabIndex        =   23
      Top             =   600
      Width           =   1455
   End
   Begin VB.CommandButton eliminarmovimiento 
      Caption         =   "Eliminar Movimiento"
      Enabled         =   0   'False
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
      Left            =   4920
      TabIndex        =   22
      Top             =   2880
      Width           =   2055
   End
   Begin MSDataGridLib.DataGrid DataGrid2 
      Bindings        =   "frmchequescancelados.frx":09AE
      Height          =   255
      Left            =   240
      TabIndex        =   19
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
      Enabled         =   0   'False
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
      Left            =   2880
      TabIndex        =   18
      Top             =   2880
      Width           =   1695
   End
   Begin VB.CommandButton grabamovimiento 
      Caption         =   "Grabar Movimiento"
      Enabled         =   0   'False
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
      TabIndex        =   17
      Top             =   2880
      Width           =   1935
   End
   Begin VB.TextBox Text3 
      Enabled         =   0   'False
      Height          =   285
      Index           =   6
      Left            =   4560
      MaxLength       =   50
      TabIndex        =   16
      Top             =   2400
      Width           =   2415
   End
   Begin VB.TextBox Text3 
      DataField       =   "idmasterasientos"
      DataSource      =   "datasiento"
      Height          =   285
      Index           =   8
      Left            =   7680
      TabIndex        =   15
      Top             =   960
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#.##0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   3082
         SubFormatType   =   1
      EndProperty
      Enabled         =   0   'False
      Height          =   285
      Index           =   5
      Left            =   3240
      TabIndex        =   14
      Top             =   2400
      Width           =   1095
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#.##0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   3082
         SubFormatType   =   1
      EndProperty
      Enabled         =   0   'False
      Height          =   285
      Index           =   4
      Left            =   1800
      TabIndex        =   13
      Top             =   2400
      Width           =   1095
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   285
      Left            =   600
      TabIndex        =   12
      Top             =   2400
      Width           =   855
   End
   Begin VB.TextBox Text3 
      DataField       =   "Fecha"
      DataSource      =   "datasiento"
      Height          =   285
      Index           =   2
      Left            =   7320
      TabIndex        =   11
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
      TabIndex        =   10
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
      TabIndex        =   9
      Top             =   960
      Visible         =   0   'False
      Width           =   375
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      DataField       =   "fecha"
      DataSource      =   "datmaestro"
      Height          =   285
      Left            =   2400
      TabIndex        =   0
      Top             =   600
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   503
      _Version        =   393216
      AllowPrompt     =   -1  'True
      Enabled         =   0   'False
      Format          =   "dd/mm/yyyy"
      PromptChar      =   "_"
   End
   Begin VB.TextBox Text2 
      DataField       =   "idmasterasientos"
      DataSource      =   "datmaestro"
      Height          =   285
      Left            =   6600
      Locked          =   -1  'True
      TabIndex        =   8
      Text            =   "Text2"
      Top             =   600
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text1 
      DataField       =   "cerrado"
      DataSource      =   "datmaestro"
      Height          =   285
      Index           =   6
      Left            =   7560
      Locked          =   -1  'True
      TabIndex        =   7
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
      TabIndex        =   6
      Top             =   600
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      DataField       =   "perfinal"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd/MM/yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   11274
         SubFormatType   =   3
      EndProperty
      DataSource      =   "datmaestro"
      Height          =   285
      Index           =   4
      Left            =   5400
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   960
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      DataField       =   "perinicial"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd/MM/yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   11274
         SubFormatType   =   3
      EndProperty
      DataSource      =   "datmaestro"
      Height          =   285
      Index           =   3
      Left            =   3960
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   960
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      DataField       =   "concepto"
      DataSource      =   "datmaestro"
      Enabled         =   0   'False
      Height          =   285
      Index           =   2
      Left            =   2400
      MaxLength       =   50
      TabIndex        =   3
      Top             =   1320
      Width           =   4455
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
      Top             =   960
      Width           =   735
   End
   Begin MSAdodcLib.Adodc datmaestro 
      Height          =   330
      Left            =   120
      Top             =   5880
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
      CacheSize       =   2000
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
      Left            =   240
      Top             =   5760
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
      Left            =   0
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
      CacheSize       =   2000
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
      Left            =   6720
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
      Height          =   1095
      Left            =   240
      TabIndex        =   24
      Top             =   6480
      Width           =   6975
      Begin KewlButtonz.KewlButtons cancelar 
         Cancel          =   -1  'True
         Height          =   735
         Left            =   4560
         TabIndex        =   58
         Top             =   240
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   1296
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
         BCOL            =   -2147483629
         BCOLO           =   -2147483629
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmchequescancelados.frx":09C7
         PICN            =   "frmchequescancelados.frx":09E3
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin KewlButtonz.KewlButtons modificar 
         Height          =   735
         Left            =   360
         TabIndex        =   59
         Top             =   240
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   1296
         BTYPE           =   14
         TX              =   "&Modif."
         ENAB            =   0   'False
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
         MICON           =   "frmchequescancelados.frx":13F5
         PICN            =   "frmchequescancelados.frx":1411
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin KewlButtonz.KewlButtons borrar 
         Height          =   735
         Left            =   3120
         TabIndex        =   60
         Top             =   240
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   1296
         BTYPE           =   14
         TX              =   "&Borrar"
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
         MICON           =   "frmchequescancelados.frx":4803
         PICN            =   "frmchequescancelados.frx":481F
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
         Height          =   735
         Left            =   5760
         TabIndex        =   61
         Top             =   240
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   1296
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
         MICON           =   "frmchequescancelados.frx":7C11
         PICN            =   "frmchequescancelados.frx":7C2D
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
         Height          =   735
         Left            =   1680
         TabIndex        =   66
         Top             =   240
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   1296
         BTYPE           =   14
         TX              =   "&Grabar"
         ENAB            =   0   'False
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
         MICON           =   "frmchequescancelados.frx":8777
         PICN            =   "frmchequescancelados.frx":8793
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
   Begin MSComCtl2.DTPicker desde 
      Height          =   375
      Left            =   8520
      TabIndex        =   26
      Top             =   360
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarBackColor=   16777215
      Format          =   63176705
      CurrentDate     =   38410
   End
   Begin MSComCtl2.DTPicker hasta 
      Height          =   375
      Left            =   8520
      TabIndex        =   25
      Top             =   840
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarBackColor=   16777215
      Format          =   63176705
      CurrentDate     =   38410
   End
   Begin VB.Frame Frame1 
      Caption         =   "Fecha de Cancelación"
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
      Left            =   7320
      TabIndex        =   27
      Top             =   120
      Width           =   4215
      Begin VB.CommandButton Cuenta 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   8
         Left            =   360
         Picture         =   "frmchequescancelados.frx":A215
         TabIndex        =   54
         TabStop         =   0   'False
         Top             =   840
         UseMaskColor    =   -1  'True
         Width           =   735
      End
      Begin VB.CommandButton Cuenta 
         Caption         =   "Desde"
         Height          =   255
         Index           =   7
         Left            =   360
         Picture         =   "frmchequescancelados.frx":A747
         TabIndex        =   53
         TabStop         =   0   'False
         Top             =   360
         UseMaskColor    =   -1  'True
         Width           =   735
      End
      Begin KewlButtonz.KewlButtons ver 
         Height          =   735
         Left            =   3000
         TabIndex        =   57
         Top             =   240
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   1296
         BTYPE           =   14
         TX              =   "&Ver"
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
         MICON           =   "frmchequescancelados.frx":AC79
         PICN            =   "frmchequescancelados.frx":AC95
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
   Begin MSMask.MaskEdBox Masksaldo 
      CausesValidation=   0   'False
      Height          =   255
      Left            =   5640
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   6000
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   0
      Appearance      =   0
      BackColor       =   -2147483629
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
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   6000
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   0
      Appearance      =   0
      BackColor       =   -2147483629
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
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   6000
      Width           =   1710
      _ExtentX        =   3016
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   0
      Appearance      =   0
      BackColor       =   -2147483629
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
      Left            =   6720
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
      LcK2            =   $"frmchequescancelados.frx":E087
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
   Begin MSDataListLib.DataCombo DataCombo1 
      Bindings        =   "frmchequescancelados.frx":E096
      Height          =   315
      Left            =   2400
      TabIndex        =   42
      Top             =   240
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   556
      _Version        =   393216
      MatchEntry      =   -1  'True
      BackColor       =   -2147483629
      ListField       =   "razonsocial"
      BoundColumn     =   "empresa"
      Text            =   ""
   End
   Begin MSAdodcLib.Adodc datempresa 
      Height          =   330
      Left            =   8040
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
   Begin VB.Frame Frame4 
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
      Height          =   1575
      Left            =   240
      TabIndex        =   44
      Top             =   1800
      Width           =   6975
      Begin VB.CommandButton Cuenta 
         Caption         =   "Detalle"
         Height          =   255
         Index           =   3
         Left            =   4680
         Picture         =   "frmchequescancelados.frx":E0AF
         TabIndex        =   48
         TabStop         =   0   'False
         Top             =   360
         UseMaskColor    =   -1  'True
         Width           =   1815
      End
      Begin VB.CommandButton Cuenta 
         Caption         =   "Haber"
         Height          =   255
         Index           =   2
         Left            =   3120
         Picture         =   "frmchequescancelados.frx":E5E1
         TabIndex        =   47
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
         Picture         =   "frmchequescancelados.frx":EB13
         TabIndex        =   46
         TabStop         =   0   'False
         Top             =   360
         UseMaskColor    =   -1  'True
         Width           =   855
      End
      Begin VB.CommandButton Cuenta 
         Caption         =   "Nº Cuenta"
         Height          =   255
         Index           =   0
         Left            =   360
         Picture         =   "frmchequescancelados.frx":F045
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   360
         UseMaskColor    =   -1  'True
         Width           =   855
      End
   End
   Begin VB.Frame Frame5 
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
      Height          =   1695
      Left            =   240
      TabIndex        =   49
      Top             =   120
      Width           =   6975
      Begin VB.CommandButton Cuenta 
         Caption         =   "Periodo:"
         Height          =   255
         Index           =   10
         Left            =   3000
         Picture         =   "frmchequescancelados.frx":F577
         TabIndex        =   56
         TabStop         =   0   'False
         Top             =   840
         UseMaskColor    =   -1  'True
         Width           =   735
      End
      Begin VB.CommandButton Cuenta 
         Caption         =   "Referencia:"
         Height          =   255
         Index           =   6
         Left            =   480
         Picture         =   "frmchequescancelados.frx":FAA9
         TabIndex        =   52
         TabStop         =   0   'False
         Top             =   1200
         UseMaskColor    =   -1  'True
         Width           =   1455
      End
      Begin VB.CommandButton Cuenta 
         Caption         =   "Nº de Asiento:"
         Height          =   255
         Index           =   5
         Left            =   480
         Picture         =   "frmchequescancelados.frx":FFDB
         TabIndex        =   51
         TabStop         =   0   'False
         Top             =   840
         UseMaskColor    =   -1  'True
         Width           =   1455
      End
      Begin VB.CommandButton Cuenta 
         Caption         =   "Fecha Asiento:"
         Height          =   255
         Index           =   4
         Left            =   480
         Picture         =   "frmchequescancelados.frx":1050D
         TabIndex        =   50
         TabStop         =   0   'False
         Top             =   480
         UseMaskColor    =   -1  'True
         Width           =   1455
      End
   End
   Begin MSAdodcLib.Adodc datcancelacion 
      Height          =   330
      Left            =   1320
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
      CacheSize       =   2000
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grilla 
      Height          =   2415
      Left            =   240
      TabIndex        =   65
      Top             =   3480
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   4260
      _Version        =   393216
      Rows            =   2000
      Cols            =   6
      BackColorFixed  =   -2147483645
      FocusRect       =   2
      SelectionMode   =   1
      AllowUserResizing=   2
      _NumberOfBands  =   1
      _Band(0).Cols   =   6
   End
   Begin MSAdodcLib.Adodc datinstrumento 
      Height          =   330
      Left            =   2520
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
   Begin MSAdodcLib.Adodc datverifica 
      Height          =   330
      Left            =   3720
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
      Left            =   2040
      TabIndex        =   36
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
      Left            =   3840
      TabIndex        =   35
      Top             =   6240
      Width           =   615
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
      Left            =   5640
      TabIndex        =   34
      Top             =   6240
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
      TabIndex        =   21
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
      TabIndex        =   20
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
End
Attribute VB_Name = "frmchequescancelados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public movimientohijo As Double
Public inicioper As Date
Public finper As Date
Dim posicion As Double
Dim registro(9999) As Double
Dim idmaster As Double
Dim bandera As Integer
Dim bandeeli As Integer


Private Sub asiento_Change()
On Error GoTo fuera

If Text2.Text <> "" Then
    movimientohijo = Text2.Text
Else
    movimientohijo = 0
End If

    
    datasiento.RecordSource = "select [Detalle Asientos].* from [Detalle Asientos] where idmasterasientos = " & frmasientosbusca.movimientohijo & " order by idasiento "
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
    Maskdebe.Text = Text6(0).Text
    Maskhaber.Text = Text6(1).Text
    Masksaldo.Text = Text6(1).Text - Text6(0).Text
  
fuera:
End Sub



Private Sub Command1_Click()

bar1.Visible = True

   datasiento.RecordSource = "select [Detalle Asientos].* from [Detalle Asientos] where debe = 0 and haber = 0 order by idasiento "
   datasiento.Refresh

If datasiento.Recordset.EOF = True Then GoTo fin
maxcont = datasiento.Recordset.RecordCount
bar1.max = maxcont + 1
cont = 1
datasiento.Recordset.MoveFirst
Do While Not datasiento.Recordset.EOF
    
    datasiento.Recordset.Delete adAffectCurrent
    datasiento.Recordset.MoveNext
    cont = cont + 1
    bar1.Value = cont
Loop

fin:
    datmaestro.RecordSource = "select [Maestro Asientos].* from [Maestro Asientos] where (fecha >= '" & frmasientosbusca.inicioper & "') and (fecha <= '" & frmasientosbusca.finper & "') and (empresa = " & login.empresaact & ") order by fecha"
    datmaestro.Refresh
    bar1.Visible = False
    mensa = MsgBox("Proceso terminado", vbInformation, "Depuración")

End Sub


Private Sub borrar_Click()
On Error Resume Next

    If login.minutasbajas = "N" Then
        mensa = MsgBox("Acceso Denegado", , "Sistema")
        Exit Sub
    End If


         mensa = MsgBox("Esta por eliminar este movimiento,ESTA SEGURO", vbYesNo, "!! Atención !!")
         If mensa = vbNo Then Exit Sub

        datverifica.RecordSource = "SELECT nrorden, empresa, anulado, idasiento, inicioper From recibocobro WHERE anulado <> 'S' AND inicioper = '" & login.iper & "' AND empresa = " & login.empresaact & " and idasiento = " & asiento.Text & " "
         datverifica.Refresh
         If datverifica.Recordset.EOF = False Then
            mensa = MsgBox("Este asiento esta relacionado a un Recibo de Cliente, Elimine el Recibo correspondiente", vbCritical, "!! Error !!")
            Exit Sub
         End If
         
         datverifica.RecordSource = "SELECT nrorden, empresa, anulado, idasiento, inicioper From ordendepago WHERE anulado <> 'S' AND inicioper = '" & login.iper & "' AND empresa = " & login.empresaact & " and idasiento = " & asiento.Text & " "
         datverifica.Refresh
         If datverifica.Recordset.EOF = False Then
            mensa = MsgBox("Este asiento esta relacionado a una Orden de Pago, Elimine la Orden correspondiente", vbCritical, "!! Error !!")
            Exit Sub
         End If
            
         datverifica.RecordSource = "SELECT empresa, inicioper, asiento From dbo.libroventas WHERE inicioper = '" & login.iper & "' AND empresa = " & login.empresaact & " and asiento = " & asiento.Text & " "
         datverifica.Refresh
         If datverifica.Recordset.EOF = False Then
            mensa = MsgBox("Este asiento esta relacionado a una Factura de Venta, Elimine la Factura correspondiente", vbCritical, "!! Error !!")
            Exit Sub
         End If
    
         datverifica.RecordSource = "SELECT empresa, inicioper, asiento From dbo.librocompras WHERE inicioper = '" & login.iper & "' AND empresa = " & login.empresaact & " and asiento = " & asiento.Text & " "
         datverifica.Refresh
         If datverifica.Recordset.EOF = False Then
            mensa = MsgBox("Este asiento esta relacionado a una Factura de Compras, Elimine la Factura correspondiente", vbCritical, "!! Error !!")
            Exit Sub
         End If
            
         datcancelacion.Recordset.Filter = "asiento = " & DataGrid3.Columns(13).Text & " and inicioper = '" & DataGrid3.Columns(14).Text & "'"
         datcancelacion.Recordset.MoveFirst
          If DataGrid3.Columns(0).Text = "R" Then
            Do While Not datcancelacion.Recordset.EOF
                      datinstrumento.RecordSource = "select recibocobroinstrumento.* from recibocobroinstrumento where empresa = " & login.empresaact & " and id = " & DataGrid3.Columns(1).Text & " "
                      datinstrumento.Refresh
                      If datinstrumento.Recordset.EOF = True Then GoTo sigue0
                      datinstrumento.Recordset.Fields("conciliado") = False
                      datinstrumento.Recordset.UpdateBatch adAffectCurrent
                      datcancelacion.Recordset.MoveNext
            Loop
            GoTo sigue0
          End If
          If DataGrid3.Columns(0).Text = "O" Then
            Do While Not datcancelacion.Recordset.EOF
                    datinstrumento.RecordSource = "select ordendepagoinstrumento.* from ordendepagoinstrumento where empresa = " & login.empresaact & " and id = " & DataGrid3.Columns(1).Text & " "
                    datinstrumento.Refresh
                    If datinstrumento.Recordset.EOF = True Then GoTo sigue0
                    datinstrumento.Recordset.Fields("conciliado") = False
                    datinstrumento.Recordset.UpdateBatch adAffectCurrent
                    datcancelacion.Recordset.MoveNext
            Loop
            GoTo sigue0
          End If
sigue0:
    
        Inicio.datauditoria.ConnectionString = login.conexiontotal
    
        Inicio.datauditoria.RecordSource = "select auditoria.* from auditoria"
        Inicio.datauditoria.Refresh
    
        Inicio.datauditoria.Recordset.AddNew
        Inicio.datauditoria.Recordset.Fields("fecha") = Date
        Inicio.datauditoria.Recordset.Fields("hora") = Str(Time)
        Inicio.datauditoria.Recordset.Fields("ventana") = "Cheques Cancelados"
        Inicio.datauditoria.Recordset.Fields("accion") = "Borrado Asiento:" + asiento.Text + " Periodo:" + Str(login.iper) + "-" + Str(login.fper)
        Inicio.datauditoria.Recordset.Fields("usuario") = login.usuarioactivo
        Inicio.datauditoria.Recordset.Fields("empresa") = login.empresaact
        Inicio.datauditoria.Recordset.UpdateBatch adAffectCurrent
    
    datmaestro.Recordset.Delete adAffectCurrent
    MsgBox "Registro Eliminado", vbInformation, "!! Procesado !!"
    Call Ver_Click
    

fuera:

End Sub

Private Sub buscar_Click()

End Sub

Private Sub Cancelar_Click()

  Rem  datmaestro.Recordset.Delete
    datmaestro.Refresh

End Sub


Private Sub Command2_Click()
On Error GoTo fuera

    datcancelacion.RecordSource = "select cancelacion.* from cancelacion where fechavencim >= '" & desde.Value & "' and fechavencim <= '" & hasta.Value & "' and empresa = " & login.empresaact & " order by importe"
    datcancelacion.Refresh

    Call grillaref_Click

fuera:
End Sub

Private Sub Command3_Click()

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
        frmasientosbusca.Show
    End If
    
fuera:

End Sub

Private Sub DataCombo2_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        
        inicioper = desde.Value
        finper = hasta.Value

        datmaestro.RecordSource = "select [Maestro Asientos].* from [Maestro Asientos] where (fecha >= '" & frmasientosbusca.inicioper & "') and (fecha <= '" & frmasientosbusca.finper & "') and (empresa = " & login.empresaact & ") and nroasiento = " & DataCombo2.Text & " order by fecha"
        datmaestro.Refresh
        
    End If


End Sub

Private Sub DataGrid3_Click()

Call llena_Click

End Sub

Private Sub DataGrid3_KeyUp(KeyCode As Integer, Shift As Integer)

Call llena_Click

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
Frame2.Visible = False
Text9.Visible = False
DataList3.Visible = False

fuera:
End Sub

Private Sub desde_Change()
On Error GoTo fuera

    Call Ver_Click
    
fuera:
End Sub

Private Sub eliminarmovimiento_Click()

      Text4.Text = ""
   Text3(4).Text = ""
   Text3(5).Text = ""
   Text3(6).Text = ""
   Text9.Text = ""
   bandeeli = 1
   Call grabamovimiento_Click

End Sub

Private Sub Form_Load()
Aplicar_skin Me
frmchequescancelados.Top = 0
frmchequescancelados.Left = 0

    Inicio.Toolbar1.Visible = True
    
    
datasiento.ConnectionString = login.conexiontotal
datccostos.ConnectionString = login.conexiontotal
datcuentas.ConnectionString = login.conexiontotal
datmaestro.ConnectionString = login.conexiontotal
datperiodo.ConnectionString = login.conexiontotal
datlistacostos.ConnectionString = login.conexiontotal
datcancelacion.ConnectionString = login.conexiontotal
datinstrumento.ConnectionString = login.conexiontotal
datverifica.ConnectionString = login.conexiontotal

datempresa.ConnectionString = login.conexiontotal

  DataCombo1.Text = login.nomempresa
  

  Inicio.Caption = login.nomempresa + "-Periodo Contable: " + Str(login.iper) + " -" + Str(login.fper)
 
  datempresa.RecordSource = "select empresa.* from empresa"
  datempresa.Refresh

      
    
    
    datlistacostos.RecordSource = "select listaccostos.* from listaccostos WHERE empresa = " & login.empresaact & ""
    datlistacostos.Refresh

    datperiodo.RecordSource = "select EMPRESA.* from EMPRESA where empresa = " & login.empresaact & ""
    datperiodo.Refresh
    
    datccostos.RecordSource = "SELECT ccostos.* FROM ccostos WHERE empresa = " & login.empresaact & ""
    datccostos.Refresh
    
    inicioper = login.iper
    finper = login.fper
    desde.Value = login.iper
    hasta.Value = login.fper

    datmaestro.RecordSource = "select [Maestro Asientos].* from [Maestro Asientos] where (fecha >= '" & frmasientosbusca.inicioper & "') and (fecha <= '" & frmasientosbusca.finper & "') and (empresa = " & login.empresaact & ") order by fecha"
    datmaestro.Refresh
    If Text2.Text <> "" Then movimientohijo = Text2.Text
    datasiento.RecordSource = "select [Detalle Asientos].* from [Detalle Asientos] where idmasterasientos = " & frmasientosbusca.movimientohijo & " order by idasiento"
    datasiento.Refresh
    MaskEdBox1.Mask = ""
    
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
  
  bandera = 0
  bandeeli = 0
             
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Inicio.Toolbar1.Visible = False

End Sub

Private Sub grabamovimiento_Click()
On Error GoTo errorgravar

    If Text4.Text = "" And bandeeli = 0 Then Exit Sub
    If bandeeli = 1 Then bandeeli = 0

    linea = grilla.Row
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

debe = 0
haber = 0
    For x = 1 To grilla.Rows - 1
        grilla.Row = x
        grilla.Col = 1
        If grilla.Text = "" Then grilla.RowHeight(x) = 0
        grilla.Col = 2
        grilla.Text = Format(grilla.Text, "0.00")
        debe = Val(grilla.Text) + debe
        grilla.Text = Format(grilla.Text, "###,##0.00")
        grilla.Col = 3
        grilla.Text = Format(grilla.Text, "0.00")
        haber = Val(grilla.Text) + haber
        grilla.Text = Format(grilla.Text, "###,##0.00")
    Next x

Maskdebe.Text = debe
Maskhaber.Text = haber
Masksaldo.Text = debe - haber

   Text4.Text = ""
   Text3(4).Text = ""
   Text3(5).Text = ""
   Text3(6).Text = ""
   Text9.Text = ""

grilla.Row = linea


Exit Sub
   datasiento.RecordSource = "select [Detalle Asientos].* from [Detalle Asientos] where idmasterasientos = " & idmaster & " and idasiento = " & registro(grilla.Row) & "  "
   datasiento.Refresh

   datasiento.Recordset.Fields("fecha") = MaskEdBox1.Text
   datasiento.Recordset.Fields("idcuenta") = Text4.Text
   datasiento.Recordset.Fields("debe") = Text3(4).Text
   datasiento.Recordset.Fields("haber") = Text3(5).Text
   datasiento.Recordset.Fields("detallefila") = Text3(6).Text
   datasiento.Recordset.Fields("ccosto") = Val(Text9.Text)
   datasiento.Recordset.UpdateBatch adAffectCurrent
   

        Inicio.datauditoria.ConnectionString = login.conexiontotal
    
        Inicio.datauditoria.RecordSource = "select auditoria.* from auditoria"
        Inicio.datauditoria.Refresh
    
        Inicio.datauditoria.Recordset.AddNew
        Inicio.datauditoria.Recordset.Fields("fecha") = Date
        Inicio.datauditoria.Recordset.Fields("hora") = Str(Time)
        Inicio.datauditoria.Recordset.Fields("ventana") = "Cancelacion Cheques"
        Inicio.datauditoria.Recordset.Fields("accion") = "Modificacion Asiento:" + asiento.Text + " Periodo:" + Str(login.iper) + "-" + Str(login.fper)
        Inicio.datauditoria.Recordset.Fields("usuario") = login.usuarioactivo
        Inicio.datauditoria.Recordset.Fields("empresa") = login.empresaact
        Inicio.datauditoria.Recordset.UpdateBatch adAffectCurrent
        
bandera = 1
Call llena_Click
grilla.SetFocus

Exit Sub
errorgravar:
    mensa = MsgBox("No esta modificando ningun movimiento, haga click en nuevo movimiento para poder modificar el asiento", vbInformation, "Atencion")

End Sub



Private Sub grabar_Click()

   datmaestro.RecordSource = "SELECT [Maestro Asientos].idmasterasientos, [Detalle Asientos].idasiento FROM [Maestro Asientos] INNER JOIN" _
                    & "[Detalle Asientos] ON [Maestro Asientos].idmasterasientos = [Detalle Asientos].idmasterasientos AND [Maestro Asientos].empresa = [Detalle Asientos].empresa where [Detalle Asientos].idasiento = " & DataGrid3.Columns(2).Text & " "
   datmaestro.Refresh
   If datmaestro.Recordset.EOF = True Then Exit Sub
   
   idmaster = datmaestro.Recordset.Fields("idmasterasientos")
   datmaestro.RecordSource = "select [Maestro Asientos].* from [Maestro Asientos] where idmasterasientos = " & idmaster & " "
   datmaestro.Refresh

   datmaestro.Recordset.Fields("fecha") = MaskEdBox1.Text
   datmaestro.Recordset.Fields("concepto") = Text1(2).Text
   datmaestro.Recordset.UpdateBatch adAffectCurrent
    
   datasiento.RecordSource = "select [Detalle Asientos].* from [Detalle Asientos] where idmasterasientos = " & idmaster & " "
   datasiento.Refresh
   

End Sub

Private Sub grilla_Click()

    grilla.Col = 1
    Text4.Text = grilla.Text
    grilla.Col = 2
    Text3(4).Text = grilla.Text
    grilla.Col = 3
    Text3(5).Text = grilla.Text
    grilla.Col = 4
    Text3(6).Text = grilla.Text

End Sub

Private Sub grilla_KeyDown(KeyCode As Integer, Shift As Integer)

Call grilla_Click

End Sub

Private Sub grillaref_Click()
    DataGrid3.Columns(1).Visible = False
    DataGrid3.Columns(2).Visible = False
    DataGrid3.Columns(3).Visible = False
    DataGrid3.Columns(6).Visible = False
    DataGrid3.Columns(9).Visible = False
    DataGrid3.Columns(10).Visible = False
    DataGrid3.Columns(11).Visible = False
    DataGrid3.Columns(0).Width = 600
    DataGrid3.Columns(4).Width = 1500
    DataGrid3.Columns(5).Width = 1000
    DataGrid3.Columns(7).Width = 1000
    DataGrid3.Columns(7).Alignment = dbgRight
    DataGrid3.Columns(8).Width = 1000
    DataGrid3.Columns(8).Alignment = dbgRight
    DataGrid3.Columns(8).NumberFormat = "##0.00"
    Call llena_Click
    
End Sub

Private Sub hasta_Change()
On Error GoTo fuera

    Call Ver_Click
    
fuera:
End Sub
Private Sub movimientos_DragDrop(Source As Control, x As Single, Y As Single)

End Sub

Private Sub llena_Click()
On Error Resume Next

If bandera = 1 Then GoTo sigue
   datmaestro.RecordSource = "SELECT [Maestro Asientos].idmasterasientos, [Detalle Asientos].idasiento FROM [Maestro Asientos] INNER JOIN" _
                    & "[Detalle Asientos] ON [Maestro Asientos].idmasterasientos = [Detalle Asientos].idmasterasientos AND [Maestro Asientos].empresa = [Detalle Asientos].empresa where [Detalle Asientos].idasiento = " & DataGrid3.Columns(2).Text & " "
   datmaestro.Refresh
   If datmaestro.Recordset.EOF = True Then Exit Sub
   
   idmaster = datmaestro.Recordset.Fields("idmasterasientos")
   datmaestro.RecordSource = "select [Maestro Asientos].* from [Maestro Asientos] where idmasterasientos = " & idmaster & " "
   datmaestro.Refresh

sigue:
   bandera = 0
   datasiento.RecordSource = "select [Detalle Asientos].* from [Detalle Asientos] where idmasterasientos = " & idmaster & " "
   datasiento.Refresh
    
        Text3(4).Text = Format(Text3(4).Text, "#,###,##0.00")
        Text3(5).Text = Format(Text3(5).Text, "#,###,##0.00")
        
grilla.Clear
If datasiento.Recordset.EOF = False Then
  grilla.Rows = datasiento.Recordset.RecordCount + 1
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
  
  grilla.Col = 1
  
datasiento.Recordset.MoveFirst
i = 1
debe = 0
haber = 0
Do While Not datasiento.Recordset.EOF
 
    grilla.Row = i
    registro(i) = datasiento.Recordset.Fields("idasiento")
    grilla.Col = 1
    grilla.Text = datasiento.Recordset.Fields("idcuenta")
    grilla.Col = 2
    grilla.Text = datasiento.Recordset.Fields("debe")
    debe = Val(grilla.Text) + debe
    grilla.Text = Format(grilla.Text, "#,###,##0.00")
    grilla.Col = 3
    grilla.Text = datasiento.Recordset.Fields("haber")
    haber = Val(grilla.Text) + haber
    grilla.Text = Format(grilla.Text, "#,###,##0.00")
    grilla.Col = 4
    If IsNull(datasiento.Recordset.Fields("detallefila")) = False Then
        grilla.Text = datasiento.Recordset.Fields("detallefila")
    Else
        grilla.Text = ""
    End If
    grilla.Col = 5
    If IsNull(datasiento.Recordset.Fields("ccosto")) = False Then
        grilla.Text = datasiento.Recordset.Fields("ccosto")
    Else
        grilla.Text = ""
    End If
    datasiento.Recordset.MoveNext
    i = i + 1
Loop
Maskdebe.Text = debe
Maskhaber.Text = haber
Masksaldo.Text = debe - haber

End If


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

Private Sub modificar_Click()
On Error GoTo fuera

    If login.minutasmodi = "N" Then
        mensa = MsgBox("Acceso Denegado", , "Sistema")
        Exit Sub
    End If


    Text3(4).Enabled = True
    Text3(5).Enabled = True
    Text3(6).Enabled = True
    Text4.Enabled = True
    Text1(2).Enabled = True
    grabamovimiento.Enabled = True
    nuevomovimiento.Enabled = True
    eliminarmovimiento.Enabled = True
    borrar.Enabled = True
    DataGrid1.Enabled = True
    MaskEdBox1.Enabled = True
    Text4.SetFocus
    

    
fuera:
End Sub

Private Sub nuevomovimiento_Click()
On Error GoTo fuera
    
   detalle = Text3(6).Text
    
   grilla.Rows = grilla.Rows + 1
   Text3(1).Text = login.empresaact
   Text3(2).Text = MaskEdBox1.Text
   Text3(8).Text = Text2.Text
   Text3(6).Text = detalle
   Text4.Text = ""
   Text3(4).Text = 0
   Text3(5).Text = 0
   Text4.SetFocus
   grilla.Row = grilla.Row + 1
   
   
fuera:
End Sub

Private Sub ordenarfecha_Click()
On Error GoTo fuera

    datcancelacion.RecordSource = "select cancelacion.* from cancelacion where fechavencim >= '" & desde.Value & "' and fechavencim <= '" & hasta.Value & "' and empresa = " & login.empresaact & " order by fechavencim"
    datcancelacion.Refresh
    
    Call grillaref_Click

fuera:
End Sub

Private Sub ordenasiento_Click()
On Error GoTo fuera

    datcancelacion.RecordSource = "select cancelacion.* from cancelacion where fechavencim >= '" & desde.Value & "' and fechavencim <= '" & hasta.Value & "' and empresa = " & login.empresaact & " order by comprobante"
    datcancelacion.Refresh
    
    Call grillaref_Click

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

Private Sub Text3_GotFocus(Index As Integer)
On Error GoTo fuera

        Text3(4).Text = Format(Text3(4).Text, "#,###,##0.00")
        Text3(5).Text = Format(Text3(5).Text, "#,###,##0.00")

    If Text4.Text = "" Then
        Text4.SetFocus
        Exit Sub
    End If

        If Index = 4 Or Index = 5 Then
              Line1.Visible = True
              Line2.Visible = True
              Line1.X2 = Text3(Index).Left
              Line1.Y2 = Text3(Index).Top
              Line2.X2 = Text3(Index).Left + Text3(Index).Width
              Line2.Y2 = Text3(Index).Top
        Else
             Line1.Visible = False
             Line2.Visible = False
        End If
                     
        Text3(Index).SelLength = Len(Text3(Index).Text)


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

    If ventana.menu = 2 Then
        ventana.menu = 0
        Text4.Text = lista_cuentas.cuentacont
    End If

fuera:
End Sub

Private Sub text4_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
            
            Call verificacuenta_Click
        
            If datccostos.Recordset.EOF = True Then GoTo sigue
              datccostos.Recordset.MoveFirst
              digito = Val(datccostos.Recordset.Fields(3))
              digito1 = Val(datccostos.Recordset.Fields(4))
              digcue = Val(Mid(Text4.Text, 1, 1))
              If digcue = digito Or digcue = digito1 And login.habcc = True Then
                DataList3.Visible = True
                Text9.Visible = True
                Frame2.Visible = True
                DataList3.SetFocus
                Exit Sub
            End If
sigue:
        Text3(4).SetFocus
    End If

End Sub

Private Sub Text4_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = 114 And Index = 0 Then
        ventana.menu = 2
        lista_cuentas.Show
    End If

End Sub

Private Sub Ver_Click()
On Error GoTo fuera

    datcancelacion.RecordSource = "select cancelacion.* from cancelacion where fechacancel >= '" & desde.Value & "' and fechacancel <= '" & hasta.Value & "' and empresa = " & login.empresaact & " order by fechacancel"
    datcancelacion.Refresh
    
    Call grillaref_Click



fuera:
End Sub

Private Sub verificacuenta_Click()
    If Text4.Text = "" Then Exit Sub

    datcuentas.RecordSource = "SELECT [Cod Contable], inicioper, empre, imp From dbo.Cuentas WHERE inicioper = '" & login.iper & "' AND empre = " & login.empresaact & " AND imp = 'S' and [Cod Contable] = " & Text4.Text & ""
    datcuentas.Refresh
    
    If datcuentas.Recordset.EOF = True Then
        MsgBox "No Existe esta cuenta contable", vbCritical, "Verificar"
        Text4.Text = ""
        Text4.SetFocus
    End If
    
End Sub
