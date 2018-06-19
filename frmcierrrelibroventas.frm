VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{5734474E-78D3-4254-99B9-C35F31BDF509}#62.0#0"; "vbskpro2.ocx"
Begin VB.Form frmclventas 
   Caption         =   "Cierre de Libro / Ver Libro Ventas"
   ClientHeight    =   6690
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11850
   Icon            =   "frmcierrrelibroventas.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6690
   ScaleWidth      =   11850
   WindowState     =   2  'Maximized
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
      Index           =   2
      Left            =   1440
      TabIndex        =   28
      Text            =   "Mes a Cerrar:"
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton totales 
      Caption         =   "totales"
      Height          =   255
      Left            =   600
      TabIndex        =   26
      Top             =   5880
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton salir 
      Caption         =   "&Salir"
      Height          =   855
      Left            =   8760
      Picture         =   "frmcierrrelibroventas.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   240
      UseMaskColor    =   -1  'True
      Width           =   1095
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmcierrrelibroventas.frx":0884
      Height          =   4215
      Left            =   120
      TabIndex        =   7
      Top             =   1560
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   7435
      _Version        =   393216
      AllowUpdate     =   -1  'True
      AllowArrows     =   -1  'True
      BackColor       =   14737632
      ForeColor       =   -2147483641
      HeadLines       =   4
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   25
      BeginProperty Column00 
         DataField       =   "empresa"
         Caption         =   "empresa"
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
         DataField       =   "fecha"
         Caption         =   "Fecha"
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
         DataField       =   "cliente"
         Caption         =   "Cliente"
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
         DataField       =   "tipoiva"
         Caption         =   "Tipo  IVA"
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
         DataField       =   "cuit"
         Caption         =   "C.U.I.T."
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
      BeginProperty Column05 
         DataField       =   "tipocompr"
         Caption         =   "Tipo Comrp."
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
      BeginProperty Column06 
         DataField       =   "numcompr"
         Caption         =   "Nº Compr."
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
      BeginProperty Column07 
         DataField       =   "col1"
         Caption         =   "col1"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column08 
         DataField       =   "col2"
         Caption         =   "col2"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column09 
         DataField       =   "col3"
         Caption         =   "col3"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column10 
         DataField       =   "col4"
         Caption         =   "col4"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column11 
         DataField       =   "col5"
         Caption         =   "col5"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column12 
         DataField       =   "col6"
         Caption         =   "col6"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column13 
         DataField       =   "col7"
         Caption         =   "col7"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column14 
         DataField       =   "col8"
         Caption         =   "col8"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column15 
         DataField       =   "col9"
         Caption         =   "col9"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column16 
         DataField       =   "col10"
         Caption         =   "col10"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column17 
         DataField       =   "col11"
         Caption         =   "col11"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column18 
         DataField       =   "col12"
         Caption         =   "col12"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column19 
         DataField       =   "col13"
         Caption         =   "col13"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column20 
         DataField       =   "col14"
         Caption         =   "col14"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column21 
         DataField       =   "col15"
         Caption         =   "col15"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column22 
         DataField       =   "total"
         Caption         =   "Total"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column23 
         DataField       =   "cuenta"
         Caption         =   "Nº Cuenta"
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
      BeginProperty Column24 
         DataField       =   "cerrado"
         Caption         =   "cerrado"
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
         MarqueeStyle    =   2
         SizeMode        =   1
         BeginProperty Column00 
            ColumnAllowSizing=   0   'False
            Locked          =   -1  'True
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column01 
            Locked          =   -1  'True
         EndProperty
         BeginProperty Column02 
            Locked          =   -1  'True
         EndProperty
         BeginProperty Column03 
            Locked          =   -1  'True
         EndProperty
         BeginProperty Column04 
            Locked          =   -1  'True
         EndProperty
         BeginProperty Column05 
            Alignment       =   2
            Locked          =   -1  'True
         EndProperty
         BeginProperty Column06 
            Locked          =   -1  'True
         EndProperty
         BeginProperty Column07 
            Alignment       =   1
            Locked          =   -1  'True
         EndProperty
         BeginProperty Column08 
            Alignment       =   1
            Locked          =   -1  'True
         EndProperty
         BeginProperty Column09 
            Alignment       =   1
            Locked          =   -1  'True
         EndProperty
         BeginProperty Column10 
            Alignment       =   1
            Locked          =   -1  'True
         EndProperty
         BeginProperty Column11 
            Alignment       =   1
            Locked          =   -1  'True
         EndProperty
         BeginProperty Column12 
            Alignment       =   1
            Locked          =   -1  'True
         EndProperty
         BeginProperty Column13 
            Alignment       =   1
            Locked          =   -1  'True
         EndProperty
         BeginProperty Column14 
            Alignment       =   1
            Locked          =   -1  'True
         EndProperty
         BeginProperty Column15 
            Alignment       =   1
            Locked          =   -1  'True
         EndProperty
         BeginProperty Column16 
            Alignment       =   1
            Locked          =   -1  'True
         EndProperty
         BeginProperty Column17 
            Alignment       =   1
            Locked          =   -1  'True
         EndProperty
         BeginProperty Column18 
            Alignment       =   1
            Locked          =   -1  'True
         EndProperty
         BeginProperty Column19 
            Alignment       =   1
            Locked          =   -1  'True
         EndProperty
         BeginProperty Column20 
            Alignment       =   1
            Locked          =   -1  'True
         EndProperty
         BeginProperty Column21 
            Alignment       =   1
            Locked          =   -1  'True
         EndProperty
         BeginProperty Column22 
            Alignment       =   1
            Locked          =   -1  'True
         EndProperty
         BeginProperty Column23 
            Alignment       =   2
            Locked          =   -1  'True
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column24 
            Locked          =   -1  'True
            Object.Visible         =   -1  'True
         EndProperty
      EndProperty
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
      Index           =   1
      Left            =   3000
      TabIndex        =   5
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
      Index           =   0
      Left            =   480
      TabIndex        =   4
      Text            =   "Desde"
      Top             =   570
      Width           =   615
   End
   Begin VB.CommandButton Aceptar 
      Caption         =   "Cerrar &Libro"
      Height          =   855
      Left            =   10440
      Picture         =   "frmcierrrelibroventas.frx":089F
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   240
      UseMaskColor    =   -1  'True
      Width           =   1095
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   3600
      TabIndex        =   1
      Top             =   480
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
      Format          =   62062593
      CurrentDate     =   38410
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   1200
      TabIndex        =   0
      Top             =   480
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
      Format          =   62062593
      CurrentDate     =   38410
   End
   Begin MSAdodcLib.Adodc datprimaryrs 
      Height          =   330
      Left            =   120
      Top             =   6120
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
   Begin VB.Frame Frame1 
      Caption         =   "Periodo a Cerrar"
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
      Height          =   1335
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   5535
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   2520
         TabIndex        =   27
         Top             =   840
         Width           =   1695
      End
   End
   Begin MSAdodcLib.Adodc datcolumnas 
      Height          =   330
      Left            =   1440
      Top             =   6120
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
   Begin MSDataGridLib.DataGrid DataGrid2 
      Bindings        =   "frmcierrrelibroventas.frx":0CE1
      Height          =   1335
      Left            =   840
      TabIndex        =   8
      Top             =   2880
      Visible         =   0   'False
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   2355
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
   Begin VB.CommandButton ver 
      Caption         =   "&Ver"
      Height          =   735
      Left            =   7680
      Picture         =   "frmcierrrelibroventas.frx":0CFB
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   240
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   1095
   End
   Begin MSMask.MaskEdBox masc 
      Height          =   375
      Index           =   0
      Left            =   3480
      TabIndex        =   10
      Top             =   5880
      Visible         =   0   'False
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      _Version        =   393216
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "$     #,##0.00;($      #,##0.00)"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox masc 
      Height          =   375
      Index           =   1
      Left            =   3840
      TabIndex        =   11
      Top             =   5880
      Visible         =   0   'False
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      _Version        =   393216
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "$     #,##0.00;($      #,##0.00)"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox masc 
      Height          =   375
      Index           =   2
      Left            =   4200
      TabIndex        =   12
      Top             =   5880
      Visible         =   0   'False
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      _Version        =   393216
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "$     #,##0.00;($      #,##0.00)"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox masc 
      Height          =   375
      Index           =   3
      Left            =   4560
      TabIndex        =   13
      Top             =   5880
      Visible         =   0   'False
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      _Version        =   393216
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "$     #,##0.00;($      #,##0.00)"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox masc 
      Height          =   375
      Index           =   4
      Left            =   4920
      TabIndex        =   14
      Top             =   5880
      Visible         =   0   'False
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      _Version        =   393216
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "$     #,##0.00;($      #,##0.00)"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox masc 
      Height          =   375
      Index           =   5
      Left            =   5280
      TabIndex        =   15
      Top             =   5880
      Visible         =   0   'False
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      _Version        =   393216
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "$     #,##0.00;($      #,##0.00)"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox masc 
      Height          =   375
      Index           =   6
      Left            =   5640
      TabIndex        =   16
      Top             =   5880
      Visible         =   0   'False
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      _Version        =   393216
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "$     #,##0.00;($      #,##0.00)"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox masc 
      Height          =   375
      Index           =   7
      Left            =   6000
      TabIndex        =   17
      Top             =   5880
      Visible         =   0   'False
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      _Version        =   393216
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "$     #,##0.00;($      #,##0.00)"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox masc 
      Height          =   375
      Index           =   8
      Left            =   6360
      TabIndex        =   18
      Top             =   5760
      Visible         =   0   'False
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   661
      _Version        =   393216
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "$     #,##0.00;($      #,##0.00)"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox masc 
      Height          =   375
      Index           =   9
      Left            =   3480
      TabIndex        =   19
      Top             =   6240
      Visible         =   0   'False
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      _Version        =   393216
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "$     #,##0.00;($      #,##0.00)"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox masc 
      Height          =   375
      Index           =   10
      Left            =   3840
      TabIndex        =   20
      Top             =   6240
      Visible         =   0   'False
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      _Version        =   393216
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "$     #,##0.00;($      #,##0.00)"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox masc 
      Height          =   375
      Index           =   11
      Left            =   4200
      TabIndex        =   21
      Top             =   6240
      Visible         =   0   'False
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      _Version        =   393216
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "$     #,##0.00;($      #,##0.00)"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox masc 
      Height          =   375
      Index           =   12
      Left            =   4560
      TabIndex        =   22
      Top             =   6240
      Visible         =   0   'False
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      _Version        =   393216
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "$     #,##0.00;($      #,##0.00)"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox masc 
      Height          =   375
      Index           =   13
      Left            =   4920
      TabIndex        =   23
      Top             =   6240
      Visible         =   0   'False
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      _Version        =   393216
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "$     #,##0.00;($      #,##0.00)"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox masc 
      Height          =   375
      Index           =   14
      Left            =   5280
      TabIndex        =   24
      Top             =   6240
      Visible         =   0   'False
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      _Version        =   393216
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "$     #,##0.00;($      #,##0.00)"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox masc 
      Height          =   375
      Index           =   15
      Left            =   5640
      TabIndex        =   25
      Top             =   6240
      Visible         =   0   'False
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      _Version        =   393216
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "$     #,##0.00;($      #,##0.00)"
      PromptChar      =   "_"
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
      LcK2            =   $"frmcierrrelibroventas.frx":1005
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
End
Attribute VB_Name = "frmclventas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public desde, hasta As String
Dim mescierre As String


Private Sub aceptar_Click()
On Error GoTo fuera

Dim perdesde As String
Dim perhasta As String

    If login.livaventascerrar = "N" Then
        mensa = MsgBox("Acceso Denegado", , "Sistema")
        Exit Sub
    End If


  KeyAscii = 13
  perdesde = desde
  perhasta = hasta
  Respuesta = MsgBox("ESTA POR CERRAR EL LIBRO VENTAS del mes: " + mescierre + ", correspondiente al periodo " + perdesde + "-" + perhasta + " , ESTA SEGURO?", vbYesNo, "Atención")
If Respuesta = vbYes Then
 datPrimaryRS.RecordSource = "select libroventas.* from libroventas WHERE fecha >= '" & frmclventas.desde & "' and fecha <= '" & frmclventas.hasta & "' and empresa = " & login.empresaact & " and inicioper = '" & login.iper & "' and libroventas.cerrado = 'N'"
 datPrimaryRS.Refresh
    If datPrimaryRS.Recordset.EOF = True Then Exit Sub

    datPrimaryRS.Recordset.MoveFirst
Do While Not datPrimaryRS.Recordset.EOF

    datPrimaryRS.Recordset.Fields("cerrado") = mescierre
    datPrimaryRS.Recordset.UpdateBatch adAffectCurrent
    datPrimaryRS.Recordset.MoveNext
Loop

    Inicio.datauditoria.ConnectionString = login.conexiontotal
    
    Inicio.datauditoria.RecordSource = "select auditoria.* from auditoria"
    Inicio.datauditoria.Refresh
    
    Inicio.datauditoria.Recordset.AddNew
    Inicio.datauditoria.Recordset.Fields("fecha") = Date
    Inicio.datauditoria.Recordset.Fields("hora") = Str(Time)
    Inicio.datauditoria.Recordset.Fields("ventana") = "Cierre Libro Ventas"
    Inicio.datauditoria.Recordset.Fields("accion") = "Cierre libro:" + mescierre + " Periodo:" + Str(login.iper) + "-" + Str(login.fper)
    Inicio.datauditoria.Recordset.Fields("usuario") = login.usuarioactivo
    Inicio.datauditoria.Recordset.Fields("empresa") = login.empresaact
    Inicio.datauditoria.Recordset.UpdateBatch adAffectCurrent

End If


Call Form_Load

fuera:
End Sub



Private Sub Command1_Click()
        libroventas.Show
End Sub


Private Sub Combo1_Change()
On Error GoTo fuera

    mescierre = Combo1.ListIndex + 1
    If Len(mescierre) = 1 Then mescierre = "0" + mescierre

fuera:
End Sub

Private Sub Combo1_Click()
On Error GoTo fuera

    mescierre = Combo1.ListIndex + 1
    If Len(mescierre) = 1 Then mescierre = "0" + mescierre
    
fuera:
End Sub

Private Sub DataGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
    Call totales_Click
End Sub

Private Sub DataGrid1_KeyUp(KeyCode As Integer, Shift As Integer)
    
  Call totales_Click
End Sub

Private Sub DataGrid1_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Call totales_Click
End Sub

Private Sub DataGrid1_Scroll(Cancel As Integer)

    Call totales_Click

End Sub

Private Sub DTPicker1_Change()
On Error GoTo fuera
    
      frmclventas.desde = DTPicker1.Value
         
     Call Ver_Click

fuera:
End Sub

Private Sub DTPicker2_Change()
On Error GoTo fuera

    frmclventas.hasta = DTPicker2.Value
    Call Ver_Click

fuera:
End Sub



Private Sub Form_Load()

datcolumnas.ConnectionString = login.conexiontotal
datPrimaryRS.ConnectionString = login.conexiontotal

 Combo1.AddItem "ENERO"
 Combo1.AddItem "FEBRERO"
 Combo1.AddItem "MARZO"
 Combo1.AddItem "ABRIL"
 Combo1.AddItem "MAYO"
 Combo1.AddItem "JUNIO"
 Combo1.AddItem "JULIO"
 Combo1.AddItem "AGOSTO"
 Combo1.AddItem "SEPTIEMBRE"
 Combo1.AddItem "OCTUBRE"
 Combo1.AddItem "NOVIEMBRE"
 Combo1.AddItem "DICIEMBRE"

 mesi = Month(Date - 20)
 Combo1.Text = Combo1.List(mesi - 1)
 mescierre = mesi
 If Len(mescierre) = 1 Then mescierre = "0" + mescierre

 datPrimaryRS.RecordSource = "select libroventas.* from libroventas WHERE libroventas.empresa = " & login.empresaact & " and inicioper = '" & login.iper & "' and finper = '" & login.fper & "' and cerrado <> 'N' Order by cerrado"
 datPrimaryRS.Refresh
  
 datPrimaryRS.RecordSource = "select libroventas.* from libroventas where empresa = " & login.empresaact & " and inicioper = '" & login.iper & "' and finper = '" & login.fper & "' and cerrado = 'N' Order by fecha"
 datPrimaryRS.Refresh
 
 datcolumnas.RecordSource = "select columnasventa.* from columnasventa where empresa = " & login.empresaact & " and inicioper = '" & login.iper & "' and finper = '" & login.fper & "'"
 datcolumnas.Refresh
 
     For x = 0 To 14
        valida = IsNull(datcolumnas.Recordset.Fields(x * 2 + 1))
        If datcolumnas.Recordset.Fields(x * 2 + 1) = "" Then valida = True
        If valida = False Then
            DataGrid1.Columns(x + 7).Caption = datcolumnas.Recordset.Fields(x * 2 + 1)
        Else
            DataGrid1.Columns(x + 7).Visible = False
        End If
    Next x
 
 desde = Date - Day(Date) + 1
 hasta = Date

 DTPicker1.Value = desde
 DTPicker2.Value = hasta

 Call Ver_Click

 
End Sub

Private Sub Frame2_DragDrop(Source As Control, x As Single, Y As Single)

End Sub

Private Sub salir_Click()
  Unload Me
End Sub

Private Sub totales_Click()
On Error GoTo fuera

For x = 0 To 15
    If DataGrid1.Columns(x + 7).Left = 0 Or DataGrid1.Columns(x + 7).Left > 10000 Then
            masc(x).Visible = False
    Else
            masc(x).Visible = True
    End If
Next x
For x = 0 To 15
    If masc(x).Visible = False Then GoTo fin
    masc(x).Left = DataGrid1.Columns(x + 7).Left + 150
    masc(x).Width = DataGrid1.Columns(x + 7).Width
fin:
Next x

fuera:
End Sub

Private Sub Ver_Click()
On Error GoTo fuera

Dim suma(16) As Currency

    datPrimaryRS.RecordSource = "select libroventas.* from libroventas where (empresa = " & login.empresaact & " and cerrado = 'N') and inicioper = '" & login.iper & "' and finper = '" & login.fper & "' Order by fecha"
    datPrimaryRS.Refresh
 Rem   If datPrimaryRS.Recordset.EOF = False Then
 Rem       datPrimaryRS.Recordset.MoveFirst
 Rem       desde = datPrimaryRS.Recordset.Fields("fecha")
 Rem       DTPicker1.Value = desde
 Rem   End If
    datPrimaryRS.RecordSource = "select libroventas.* from libroventas where (empresa = " & login.empresaact & " and fecha >= '" & frmclventas.desde & "' and fecha <= '" & frmclventas.hasta & "' and cerrado = 'N') and inicioper = '" & login.iper & "' and finper = '" & login.fper & "' Order by fecha"
    datPrimaryRS.Refresh

For x = 1 To 16
    suma(x) = 0
Next x
 
If datPrimaryRS.Recordset.EOF = True Then GoTo paso2
datPrimaryRS.Recordset.MoveFirst
paso1:
If datPrimaryRS.Recordset.EOF = True Then GoTo paso2
For x = 8 To 22
    If IsNull(datPrimaryRS.Recordset.Fields(x)) = True Then
            datPrimaryRS.Recordset.Fields(x) = 0
    End If
Next x

If IsNull(datPrimaryRS.Recordset.Fields(24)) = True Then datPrimaryRS.Recordset.Fields(24) = 0
        
suma(1) = suma(1) + datPrimaryRS.Recordset.Fields(8).Value
suma(2) = suma(2) + datPrimaryRS.Recordset.Fields(9).Value
suma(3) = suma(3) + datPrimaryRS.Recordset.Fields(10).Value
suma(4) = suma(4) + datPrimaryRS.Recordset.Fields(11).Value
suma(5) = suma(5) + datPrimaryRS.Recordset.Fields(12).Value
suma(6) = suma(6) + datPrimaryRS.Recordset.Fields(13).Value
suma(7) = suma(7) + datPrimaryRS.Recordset.Fields(14).Value
suma(8) = suma(8) + datPrimaryRS.Recordset.Fields(15).Value
suma(9) = suma(9) + datPrimaryRS.Recordset.Fields(16).Value
suma(10) = suma(10) + datPrimaryRS.Recordset.Fields(17).Value
suma(11) = suma(11) + datPrimaryRS.Recordset.Fields(18).Value
suma(12) = suma(12) + datPrimaryRS.Recordset.Fields(19).Value
suma(13) = suma(13) + datPrimaryRS.Recordset.Fields(20).Value
suma(14) = suma(14) + datPrimaryRS.Recordset.Fields(21).Value
suma(15) = suma(15) + datPrimaryRS.Recordset.Fields(22).Value
suma(16) = suma(16) + datPrimaryRS.Recordset.Fields(24).Value
datPrimaryRS.Recordset.MoveNext
GoTo paso1
paso2:
For x = 0 To 15
    masc(x) = suma(x + 1)
    masc(x).Top = 5760
Next x

fuera:
End Sub
