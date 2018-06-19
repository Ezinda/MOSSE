VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{5734474E-78D3-4254-99B9-C35F31BDF509}#62.0#0"; "vbskpro2.ocx"
Object = "{D18BBD1F-82BB-4385-BED3-E9D31A3E361E}#1.0#0"; "KewlButtonz.ocx"
Begin VB.Form frmasientosbusca 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Carga de Asientos"
   ClientHeight    =   7755
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11700
   Icon            =   "frmasientosbusca.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7755
   ScaleWidth      =   11700
   Begin VB.CommandButton Cuenta 
      Caption         =   "Fecha Registro:"
      Height          =   255
      Index           =   9
      Left            =   3720
      Picture         =   "frmasientosbusca.frx":0442
      TabIndex        =   60
      TabStop         =   0   'False
      Top             =   600
      UseMaskColor    =   -1  'True
      Width           =   1455
   End
   Begin MSComctlLib.ProgressBar bar1 
      Height          =   255
      Left            =   240
      TabIndex        =   47
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
      DataField       =   "ccosto"
      DataSource      =   "datasiento"
      Height          =   285
      Left            =   2640
      TabIndex        =   44
      Top             =   3120
      Visible         =   0   'False
      Width           =   615
   End
   Begin MSDataListLib.DataList DataList3 
      Bindings        =   "frmasientosbusca.frx":0974
      Height          =   1620
      Left            =   2640
      TabIndex        =   43
      Top             =   3480
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
      BackColor       =   &H00E0E0E0&
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
      Left            =   2400
      TabIndex        =   45
      Top             =   2880
      Visible         =   0   'False
      Width           =   3015
   End
   Begin MSDataListLib.DataList DataList2 
      Bindings        =   "frmasientosbusca.frx":0991
      Height          =   2205
      Left            =   2520
      TabIndex        =   42
      Top             =   3480
      Visible         =   0   'False
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   3889
      _Version        =   393216
      MatchEntry      =   -1  'True
      OLEDragMode     =   1
      OLEDropMode     =   1
      BackColor       =   -2147483629
      ListField       =   "codigo"
      BoundColumn     =   "Cod Contable"
   End
   Begin VB.CommandButton ordenarfecha 
      Caption         =   "Ordenar x Fecha"
      Height          =   255
      Left            =   7320
      TabIndex        =   41
      Top             =   1680
      Width           =   1575
   End
   Begin VB.CommandButton ordenasiento 
      Caption         =   "Ordenar x NºAsiento"
      Height          =   255
      Left            =   8880
      TabIndex        =   40
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
      TabIndex        =   36
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
      TabIndex        =   35
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
      TabIndex        =   32
      TabStop         =   0   'False
      Text            =   " "
      Top             =   5880
      Visible         =   0   'False
      Width           =   855
   End
   Begin MSDataGridLib.DataGrid DataGrid3 
      Bindings        =   "frmasientosbusca.frx":09AA
      Height          =   5655
      Left            =   7320
      TabIndex        =   1
      Top             =   1920
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   9975
      _Version        =   393216
      BackColor       =   -2147483629
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
      ColumnCount     =   9
      BeginProperty Column00 
         DataField       =   "fecha"
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
      BeginProperty Column01 
         DataField       =   "fecharegistro"
         Caption         =   "fecharegistro"
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
      BeginProperty Column03 
         DataField       =   "nroasiento"
         Caption         =   "Nº Asiento"
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
      BeginProperty Column04 
         DataField       =   "concepto"
         Caption         =   "Referencia"
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
      BeginProperty Column05 
         DataField       =   "perinicial"
         Caption         =   "perinicial"
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
         DataField       =   "perfinal"
         Caption         =   "perfinal"
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
      BeginProperty Column08 
         DataField       =   "cerrado"
         Caption         =   "cerrado"
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
            Alignment       =   2
            Locked          =   -1  'True
         EndProperty
         BeginProperty Column01 
            Locked          =   -1  'True
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column02 
            Locked          =   -1  'True
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column03 
            Alignment       =   2
            Locked          =   -1  'True
         EndProperty
         BeginProperty Column04 
            Locked          =   -1  'True
         EndProperty
         BeginProperty Column05 
            Locked          =   -1  'True
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column06 
            Locked          =   -1  'True
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column07 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column08 
            Locked          =   -1  'True
            Object.Visible         =   0   'False
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
      TabIndex        =   26
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
      TabIndex        =   25
      Top             =   3120
      Width           =   2055
   End
   Begin MSMask.MaskEdBox visual 
      CausesValidation=   0   'False
      Height          =   495
      Left            =   4080
      TabIndex        =   22
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
      Bindings        =   "frmasientosbusca.frx":09C3
      Height          =   255
      Left            =   1200
      TabIndex        =   21
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
      TabIndex        =   20
      Top             =   3120
      Width           =   1695
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmasientosbusca.frx":09DC
      Height          =   2175
      Left            =   240
      TabIndex        =   19
      Top             =   3720
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   3836
      _Version        =   393216
      BackColor       =   16777215
      Enabled         =   0   'False
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
      TabIndex        =   18
      Top             =   3120
      Width           =   1935
   End
   Begin VB.TextBox Text3 
      DataField       =   "detallefila"
      DataSource      =   "datasiento"
      Enabled         =   0   'False
      Height          =   285
      Index           =   6
      Left            =   4560
      MaxLength       =   50
      TabIndex        =   17
      Top             =   2520
      Width           =   2415
   End
   Begin VB.TextBox Text3 
      DataField       =   "idmasterasientos"
      DataSource      =   "datasiento"
      Height          =   285
      Index           =   8
      Left            =   7680
      TabIndex        =   16
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
      Enabled         =   0   'False
      Height          =   285
      Index           =   5
      Left            =   3240
      TabIndex        =   15
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
      Enabled         =   0   'False
      Height          =   285
      Index           =   4
      Left            =   1800
      TabIndex        =   14
      Top             =   2520
      Width           =   1095
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      DataField       =   "idcuenta"
      DataSource      =   "datasiento"
      Enabled         =   0   'False
      Height          =   285
      Left            =   600
      TabIndex        =   13
      Top             =   2520
      Width           =   855
   End
   Begin VB.TextBox Text3 
      DataField       =   "Fecha"
      DataSource      =   "datasiento"
      Height          =   285
      Index           =   2
      Left            =   7320
      TabIndex        =   12
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
      TabIndex        =   11
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
      TabIndex        =   10
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
      TabIndex        =   9
      Text            =   "Text2"
      Top             =   600
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton grabar 
      Caption         =   "grabar"
      Height          =   375
      Left            =   4560
      TabIndex        =   8
      Top             =   5520
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
      Height          =   1095
      Left            =   240
      TabIndex        =   27
      Top             =   6480
      Width           =   6975
      Begin KewlButtonz.KewlButtons cancelar 
         Cancel          =   -1  'True
         Height          =   735
         Left            =   360
         TabIndex        =   64
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
         MICON           =   "frmasientosbusca.frx":09F5
         PICN            =   "frmasientosbusca.frx":0A11
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
         Left            =   1440
         TabIndex        =   65
         Top             =   240
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   1296
         BTYPE           =   14
         TX              =   "&Modif."
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
         MICON           =   "frmasientosbusca.frx":1423
         PICN            =   "frmasientosbusca.frx":143F
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
         Left            =   2520
         TabIndex        =   66
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
         MICON           =   "frmasientosbusca.frx":4831
         PICN            =   "frmasientosbusca.frx":484D
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin KewlButtonz.KewlButtons buscar 
         Height          =   735
         Left            =   3600
         TabIndex        =   67
         Top             =   240
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   1296
         BTYPE           =   14
         TX              =   "Buscar &Errores"
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
         MICON           =   "frmasientosbusca.frx":7C3F
         PICN            =   "frmasientosbusca.frx":7C5B
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin KewlButtonz.KewlButtons Command1 
         Height          =   735
         Left            =   4680
         TabIndex        =   68
         Top             =   240
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   1296
         BTYPE           =   14
         TX              =   "&Depurar 0"
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
         MICON           =   "frmasientosbusca.frx":866D
         PICN            =   "frmasientosbusca.frx":8689
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
         TabIndex        =   69
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
         MICON           =   "frmasientosbusca.frx":909B
         PICN            =   "frmasientosbusca.frx":90B7
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
      TabIndex        =   29
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
      Format          =   159186945
      CurrentDate     =   38410
   End
   Begin MSComCtl2.DTPicker hasta 
      Height          =   375
      Left            =   8520
      TabIndex        =   28
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
      Format          =   159186945
      CurrentDate     =   38410
   End
   Begin VB.Frame Frame1 
      Caption         =   "Periodo de busqueda"
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
      TabIndex        =   30
      Top             =   120
      Width           =   4215
      Begin VB.CommandButton Cuenta 
         Caption         =   "Registro:"
         Height          =   255
         Index           =   11
         Left            =   2040
         Picture         =   "frmasientosbusca.frx":9C01
         TabIndex        =   62
         TabStop         =   0   'False
         Top             =   1200
         UseMaskColor    =   -1  'True
         Width           =   975
      End
      Begin VB.CommandButton Cuenta 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   8
         Left            =   360
         Picture         =   "frmasientosbusca.frx":A133
         TabIndex        =   59
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
         Picture         =   "frmasientosbusca.frx":A665
         TabIndex        =   58
         TabStop         =   0   'False
         Top             =   360
         UseMaskColor    =   -1  'True
         Width           =   735
      End
      Begin MSDataListLib.DataCombo DataCombo2 
         Bindings        =   "frmasientosbusca.frx":AB97
         Height          =   315
         Left            =   3000
         TabIndex        =   48
         Top             =   1200
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         ListField       =   "nroasiento"
         Text            =   ""
      End
      Begin KewlButtonz.KewlButtons ver 
         Height          =   735
         Left            =   3000
         TabIndex        =   63
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
         MICON           =   "frmasientosbusca.frx":ABB0
         PICN            =   "frmasientosbusca.frx":ABCC
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
      TabIndex        =   31
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
      TabIndex        =   33
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
      TabIndex        =   34
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
   Begin vbskpro.Skinner Skinner1 
      Left            =   0
      Top             =   0
      _ExtentX        =   1270
      _ExtentY        =   1270
      CloseButtonToolTipText=   "Cerrar"
      MinButtonToolTipText=   "Minimizar"
      MaxButtonToolTipText=   "Maximizar"
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
      LcK2            =   $"frmasientosbusca.frx":DFBE
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
      Bindings        =   "frmasientosbusca.frx":DFCD
      Height          =   315
      Left            =   2400
      TabIndex        =   46
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
      Height          =   1815
      Left            =   240
      TabIndex        =   49
      Top             =   1800
      Width           =   6975
      Begin VB.CommandButton Cuenta 
         Caption         =   "Detalle"
         Height          =   255
         Index           =   3
         Left            =   4680
         Picture         =   "frmasientosbusca.frx":DFE6
         TabIndex        =   53
         TabStop         =   0   'False
         Top             =   480
         UseMaskColor    =   -1  'True
         Width           =   1815
      End
      Begin VB.CommandButton Cuenta 
         Caption         =   "Haber"
         Height          =   255
         Index           =   2
         Left            =   3120
         Picture         =   "frmasientosbusca.frx":E518
         TabIndex        =   52
         TabStop         =   0   'False
         Top             =   480
         UseMaskColor    =   -1  'True
         Width           =   855
      End
      Begin VB.CommandButton Cuenta 
         Caption         =   "Debe"
         Height          =   255
         Index           =   1
         Left            =   1680
         Picture         =   "frmasientosbusca.frx":EA4A
         TabIndex        =   51
         TabStop         =   0   'False
         Top             =   480
         UseMaskColor    =   -1  'True
         Width           =   855
      End
      Begin VB.CommandButton Cuenta 
         Caption         =   "Nº Cuenta"
         Height          =   255
         Index           =   0
         Left            =   360
         Picture         =   "frmasientosbusca.frx":EF7C
         TabIndex        =   50
         TabStop         =   0   'False
         Top             =   480
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
      TabIndex        =   54
      Top             =   120
      Width           =   6975
      Begin VB.CommandButton Cuenta 
         Caption         =   "Periodo:"
         Height          =   255
         Index           =   10
         Left            =   3000
         Picture         =   "frmasientosbusca.frx":F4AE
         TabIndex        =   61
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
         Picture         =   "frmasientosbusca.frx":F9E0
         TabIndex        =   57
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
         TabIndex        =   56
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
         TabIndex        =   55
         TabStop         =   0   'False
         Top             =   480
         UseMaskColor    =   -1  'True
         Width           =   1455
      End
   End
   Begin MSAdodcLib.Adodc datverifica 
      Height          =   330
      Left            =   1680
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
      TabIndex        =   39
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
      TabIndex        =   38
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
      TabIndex        =   37
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
      TabIndex        =   24
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
      TabIndex        =   23
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
Attribute VB_Name = "frmasientosbusca"
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

         mensa = MsgBox("Esta por eliminar este Asiento,ESTA SEGURO", vbYesNo, "!! Atención !!")
         If mensa = vbNo Then Exit Sub
    
         datverifica.RecordSource = "SELECT nrorden, empresa, anulado, idasiento, inicioper From recibocobro WHERE anulado <> 'S' AND inicioper = '" & login.iper & "' AND empresa = " & login.empresaact & " and idasiento = " & asiento.Text & " "
         datverifica.Refresh
         If datverifica.Recordset.EOF = False Then
            mensa = MsgBox("Este asiento esta relacionado a un Recibo de Cliente, Elimine el Recibo correspondiente", vbCritical, "!! Error !!")
            Exit Sub
         End If
        
        
         datasiento.Recordset.MoveFirst
         Do While Not datasiento.Recordset.EOF
            idasiento = datasiento.Recordset.Fields("idasiento")
            
            datverifica.RecordSource = "select cancelacion.* from cancelacion where empresa = " & login.empresaact & " and idasiento = " & idasiento & ""
            datverifica.Refresh
            If datverifica.Recordset.EOF = False Then
                    mensa = MsgBox("Este asiento esta relacionado a un Movimiento Bancario, Elimine este asiento desde Cheques-Cancelacion", vbCritical, "!! Error !!")
                    Exit Sub
            End If
           
            datasiento.Recordset.MoveNext
         Loop
               
         
         If datverifica.Recordset.EOF = False Then
            mensa = MsgBox("Este asiento esta relacionado a un movimiento Bancario, Elimine la Cancelacion Bancaria desde el menu Cheques-Cancelacion", vbCritical, "!! Error !!")
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
         

         
         
    
        Inicio.datauditoria.ConnectionString = login.conexiontotal
    
        Inicio.datauditoria.RecordSource = "select auditoria.* from auditoria"
        Inicio.datauditoria.Refresh
    
        Inicio.datauditoria.Recordset.AddNew
        Inicio.datauditoria.Recordset.Fields("fecha") = Date
        Inicio.datauditoria.Recordset.Fields("hora") = Str(Time)
        Inicio.datauditoria.Recordset.Fields("ventana") = "Carga de Asientos"
        Inicio.datauditoria.Recordset.Fields("accion") = "Borrado Asiento:" + asiento.Text + " Periodo:" + Str(login.iper) + "-" + Str(login.fper)
        Inicio.datauditoria.Recordset.Fields("usuario") = login.usuarioactivo
        Inicio.datauditoria.Recordset.Fields("empresa") = login.empresaact
        Inicio.datauditoria.Recordset.UpdateBatch adAffectCurrent
    
    datmaestro.Recordset.Delete adAffectCurrent
    

fuera:

End Sub

Private Sub buscar_Click()
On Error GoTo fuera

    datmaestro.RecordSource = "select [Maestro Asientos].* from [Maestro Asientos] where (fecha >= '" & frmasientosbusca.inicioper & "') and (fecha <= '" & frmasientosbusca.finper & "') and (empresa = " & login.empresaact & ") order by fecha"
    datmaestro.Refresh

    datmaestro.Recordset.MoveFirst

Do While Not datmaestro.Recordset.EOF
   DataGrid3.Bookmark = datmaestro.Recordset.AbsolutePosition
   datasiento.RecordSource = "select [Detalle Asientos].* from [Detalle Asientos] where idmasterasientos = " & datmaestro.Recordset.Fields("idmasterasientos") & " order by idasiento "
   datasiento.Refresh
   frmasientosbusca.Refresh
   
    Text6(0).Text = 0
    Text6(1).Text = 0
    Text6(2).Text = 0
    If datasiento.Recordset.EOF = True Then GoTo paso2
    datasiento.Recordset.MoveFirst
paso1:
    If datasiento.Recordset.Fields(2) = 0 And (datasiento.Recordset.Fields(3) <> 0 Or datasiento.Recordset.Fields(4) <> 0) Then
        mensa = MsgBox("Cuenta 0 imputada " + Str(datmaestro.Recordset.Fields("nroasiento")), vbYesNo, "Continua ?")
        If mensa = vbNo Then Exit Sub
    End If
        
        
    Text6(0).Text = datasiento.Recordset.Fields(3) + Text6(0)
    Text6(1).Text = datasiento.Recordset.Fields(4) + Text6(1)
    Maskdebe.Text = Text6(0).Text
    Maskhaber.Text = Text6(1).Text
    datasiento.Recordset.MoveNext
    If datasiento.Recordset.EOF = False Then
        GoTo paso1
    Else
        GoTo paso2
    End If
paso2:
    Text6(2).Text = Text6(0) - Text6(1)
    Masksaldo = Text6(2)
   If Masksaldo.Text <> 0 Then
        mensa = MsgBox("Asiento " + Str(datmaestro.Recordset.Fields("nroasiento")) + " Desvalanceado", vbYesNo, "Continua ?")
        If mensa = vbNo Then Exit Sub
   End If
   datmaestro.Recordset.MoveNext
Loop

fuera:
End Sub

Private Sub Cancelar_Click()

  Rem  datmaestro.Recordset.Delete
    datmaestro.Refresh

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

Rem    If KeyAscii = 13 Then
Rem        KeyAscii = 0
Rem        Text4.Text = DataList2.BoundText
Rem        Text3(4).SetFocus
Rem    End If
    
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
                Frame2.Visible = True
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
Frame2.Visible = False
Text9.Visible = False
DataList3.Visible = False

fuera:
End Sub

Private Sub desde_Change()
On Error GoTo fuera
    
    If desde < login.iper Then desde = login.iper
    Call Ver_Click
    
fuera:
End Sub

Private Sub eliminarmovimiento_Click()
On Error GoTo fuera

    mensa = MsgBox("Esta por eliminar un movimiento de este asiento, esta seguro", vbYesNo, "!! Atención !!")
    If mensa = vbYes Then
            If datasiento.Recordset.EOF = False Then datasiento.Recordset.Delete
    End If

fuera:
End Sub

Private Sub Form_Load()
Aplicar_skin Me
frmasientosbusca.Top = 0
frmasientosbusca.Left = 0

    Inicio.Toolbar1.Visible = True
    
    
datasiento.ConnectionString = login.conexiontotal
datccostos.ConnectionString = login.conexiontotal
datcuentas.ConnectionString = login.conexiontotal
datmaestro.ConnectionString = login.conexiontotal
datperiodo.ConnectionString = login.conexiontotal
datlistacostos.ConnectionString = login.conexiontotal
datempresa.ConnectionString = login.conexiontotal
datverifica.ConnectionString = login.conexiontotal

  DataCombo1.Text = login.nomempresa
  

  Inicio.Caption = login.nomempresa + "-Periodo Contable: " + Str(login.iper) + " -" + Str(login.fper)
 
  datempresa.RecordSource = "select empresa.* from empresa"
  datempresa.Refresh

     
    
    datlistacostos.RecordSource = "select listaccostos.* from listaccostos WHERE empresa = " & login.empresaact & " order by cc"
    datlistacostos.Refresh
    datcuentas.RecordSource = "select listacuentas.* from listacuentas WHERE empre = " & login.empresaact & " and inicioper = '" & login.iper & "' and finper = '" & login.fper & "' ORDER BY IDCUENTA"
    datcuentas.Refresh
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
    
             
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Inicio.Toolbar1.Visible = False

End Sub

Private Sub grabamovimiento_Click()
On Error GoTo errorgravar

    Text3(2).Text = MaskEdBox1.Text
    datasiento.Recordset.UpdateBatch adAffectCurrent
    
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
    datasiento.Recordset.MoveFirst
paso1:
    Text6(0).Text = datasiento.Recordset.Fields(3) + Text6(0).Text
    Text6(1).Text = datasiento.Recordset.Fields(4) + Text6(1).Text
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

        Inicio.datauditoria.ConnectionString = login.conexiontotal
    
        Inicio.datauditoria.RecordSource = "select auditoria.* from auditoria"
        Inicio.datauditoria.Refresh
    
        Inicio.datauditoria.Recordset.AddNew
        Inicio.datauditoria.Recordset.Fields("fecha") = Date
        Inicio.datauditoria.Recordset.Fields("hora") = Str(Time)
        Inicio.datauditoria.Recordset.Fields("ventana") = "Carga de Asientos"
        Inicio.datauditoria.Recordset.Fields("accion") = "Modificacion Asiento:" + asiento.Text + " Periodo:" + Str(login.iper) + "-" + Str(login.fper)
        Inicio.datauditoria.Recordset.Fields("usuario") = login.usuarioactivo
        Inicio.datauditoria.Recordset.Fields("empresa") = login.empresaact
        Inicio.datauditoria.Recordset.UpdateBatch adAffectCurrent

    nuevomovimiento.SetFocus
    nuevomovimiento.SetFocus
Exit Sub
errorgravar:
    mensa = MsgBox("No esta modificando ningun movimiento, haga click en nuevo movimiento para poder modificar el asiento", vbInformation, "Atencion")

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
    DataList2.Enabled = True
    MaskEdBox1.Enabled = True
    

    
fuera:
End Sub

Private Sub nuevomovimiento_Click()
On Error GoTo fuera
    
   detalle = Text3(6).Text
   datasiento.Recordset.AddNew

   Text3(1).Text = login.empresaact
   Text3(2).Text = MaskEdBox1.Text
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


        Text3(4).Text = Format(Text3(4).Text, "#,###,##0.00")
        Text3(5).Text = Format(Text3(5).Text, "#,###,##0.00")

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
    If Text2.Text <> "" Then movimientohijo = Text2.Text
    datasiento.RecordSource = "select [Detalle Asientos].* from [Detalle Asientos] where idmasterasientos = " & frmasientosbusca.movimientohijo & " order by idasiento "
    datasiento.Refresh

fuera:
End Sub

