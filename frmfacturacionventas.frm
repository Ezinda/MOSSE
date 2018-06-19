VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmfacturacionventas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Libro IVA Ventas"
   ClientHeight    =   6765
   ClientLeft      =   1095
   ClientTop       =   330
   ClientWidth     =   14610
   ControlBox      =   0   'False
   Icon            =   "frmfacturacionventas.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   119.327
   ScaleMode       =   0  'User
   ScaleWidth      =   382.711
   Begin VB.TextBox Text9 
      Alignment       =   2  'Center
      DataField       =   "ccosto"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Left            =   360
      TabIndex        =   99
      Top             =   2760
      Visible         =   0   'False
      Width           =   615
   End
   Begin MSDataListLib.DataList DataList3 
      Bindings        =   "frmfacturacionventas.frx":0442
      Height          =   1620
      Left            =   360
      TabIndex        =   100
      Top             =   3120
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
   Begin MSAdodcLib.Adodc datlistacostos 
      Height          =   330
      Left            =   6600
      Top             =   4920
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
   Begin MSAdodcLib.Adodc datccostos 
      Height          =   330
      Left            =   6600
      Top             =   5160
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
   Begin VB.TextBox Text10 
      DataField       =   "asiento"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Left            =   5280
      TabIndex        =   98
      Text            =   "Text10"
      Top             =   360
      Visible         =   0   'False
      Width           =   615
   End
   Begin MSDataListLib.DataList DataList2 
      Bindings        =   "frmfacturacionventas.frx":045F
      Height          =   1620
      Left            =   240
      TabIndex        =   97
      Top             =   4320
      Visible         =   0   'False
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   2752
      _Version        =   393216
      MatchEntry      =   -1  'True
      BackColor       =   12640511
      ForeColor       =   -2147483647
      ListField       =   "codigo"
      BoundColumn     =   "Cod Contable"
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
   Begin VB.TextBox Text5 
      DataField       =   "inicioperiodo"
      DataSource      =   "datperiodo"
      Height          =   285
      Index           =   0
      Left            =   6120
      TabIndex        =   96
      Text            =   "Text5"
      Top             =   4920
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text5 
      DataField       =   "finperiodo"
      DataSource      =   "datperiodo"
      Height          =   285
      Index           =   1
      Left            =   6120
      TabIndex        =   95
      Text            =   "Text5"
      Top             =   5160
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text8 
      DataField       =   "asentado"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Left            =   5400
      TabIndex        =   94
      Text            =   "Text8"
      Top             =   120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton buscafact 
      Caption         =   "buscafact"
      Height          =   255
      Left            =   1560
      TabIndex        =   93
      Top             =   6120
      Visible         =   0   'False
      Width           =   975
   End
   Begin MSDataListLib.DataList DataList1 
      Bindings        =   "frmfacturacionventas.frx":0478
      Height          =   1605
      Left            =   2760
      TabIndex        =   5
      Top             =   600
      Visible         =   0   'False
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   2831
      _Version        =   393216
      IntegralHeight  =   0   'False
      MatchEntry      =   -1  'True
      BackColor       =   16777215
      ListField       =   "razonsocial"
   End
   Begin VB.CommandButton anulada 
      Caption         =   "Fact.Anulada"
      Height          =   255
      Left            =   2400
      TabIndex        =   92
      TabStop         =   0   'False
      Top             =   1680
      Width           =   1095
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      DataField       =   "cht"
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
      Height          =   285
      Index           =   31
      Left            =   6960
      TabIndex        =   89
      Text            =   " "
      Top             =   4560
      Width           =   615
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      DataField       =   "cdt"
      DataSource      =   "datPrimaryRS"
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
      Index           =   30
      Left            =   6240
      TabIndex        =   88
      Text            =   " "
      Top             =   4560
      Width           =   615
   End
   Begin VB.CommandButton calcular 
      Caption         =   "calcular"
      Height          =   255
      Left            =   240
      TabIndex        =   53
      Top             =   6120
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      DataField       =   "col15"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """$"" #,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   11274
         SubFormatType   =   2
      EndProperty
      DataSource      =   "datPrimaryRS"
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
      Index           =   14
      Left            =   3360
      TabIndex        =   22
      Top             =   5880
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      DataField       =   "col14"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """$"" #,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   11274
         SubFormatType   =   2
      EndProperty
      DataSource      =   "datPrimaryRS"
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
      Index           =   13
      Left            =   3360
      TabIndex        =   21
      Top             =   5640
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      DataField       =   "col13"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """$"" #,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2058
         SubFormatType   =   2
      EndProperty
      DataSource      =   "datPrimaryRS"
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
      Index           =   12
      Left            =   3360
      TabIndex        =   20
      Top             =   5400
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      DataField       =   "col12"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """$"" #,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2058
         SubFormatType   =   2
      EndProperty
      DataSource      =   "datPrimaryRS"
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
      Left            =   3360
      TabIndex        =   19
      Top             =   5160
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      DataField       =   "col11"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """$"" #,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2058
         SubFormatType   =   2
      EndProperty
      DataSource      =   "datPrimaryRS"
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
      Left            =   3360
      TabIndex        =   18
      Top             =   4920
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      DataField       =   "col10"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """$"" #,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2058
         SubFormatType   =   2
      EndProperty
      DataSource      =   "datPrimaryRS"
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
      Left            =   3360
      TabIndex        =   17
      Top             =   4680
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      DataField       =   "col9"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """$"" #,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2058
         SubFormatType   =   2
      EndProperty
      DataSource      =   "datPrimaryRS"
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
      Left            =   3360
      TabIndex        =   16
      Top             =   4440
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      DataField       =   "col8"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """$"" #,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2058
         SubFormatType   =   2
      EndProperty
      DataSource      =   "datPrimaryRS"
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
      Left            =   3360
      TabIndex        =   15
      Top             =   4200
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      DataField       =   "col7"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """$"" #,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2058
         SubFormatType   =   2
      EndProperty
      DataSource      =   "datPrimaryRS"
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
      Left            =   3360
      TabIndex        =   14
      Top             =   3960
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      DataField       =   "col6"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """$"" #,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2058
         SubFormatType   =   2
      EndProperty
      DataSource      =   "datPrimaryRS"
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
      Left            =   3360
      TabIndex        =   13
      Top             =   3720
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      DataField       =   "col5"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """$"" #,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2058
         SubFormatType   =   2
      EndProperty
      DataSource      =   "datPrimaryRS"
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
      Left            =   3360
      TabIndex        =   12
      Top             =   3480
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      DataField       =   "col4"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """$"" #,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2058
         SubFormatType   =   2
      EndProperty
      DataSource      =   "datPrimaryRS"
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
      Left            =   3360
      TabIndex        =   11
      Top             =   3240
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      DataField       =   "col3"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """$"" #,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2058
         SubFormatType   =   2
      EndProperty
      DataSource      =   "datPrimaryRS"
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
      Left            =   3360
      TabIndex        =   10
      Top             =   3000
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      DataField       =   "col2"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """$"" #,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2058
         SubFormatType   =   2
      EndProperty
      DataSource      =   "datPrimaryRS"
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
      Left            =   3360
      TabIndex        =   9
      Top             =   2760
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      DataField       =   "col1"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """$"" #,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2058
         SubFormatType   =   2
      EndProperty
      DataSource      =   "datPrimaryRS"
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
      Left            =   3360
      TabIndex        =   8
      Top             =   2490
      Width           =   1215
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      DataField       =   "ch15"
      DataSource      =   "datPrimaryRS"
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
      Index           =   29
      Left            =   5280
      TabIndex        =   85
      Text            =   " "
      Top             =   5880
      Width           =   615
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      DataField       =   "cd15"
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
      Height          =   285
      Index           =   28
      Left            =   4680
      TabIndex        =   84
      Text            =   " "
      Top             =   5880
      Width           =   615
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      DataField       =   "ch14"
      DataSource      =   "datPrimaryRS"
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
      Index           =   27
      Left            =   5280
      TabIndex        =   83
      Text            =   " "
      Top             =   5640
      Width           =   615
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      DataField       =   "cd14"
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
      Height          =   285
      Index           =   26
      Left            =   4680
      TabIndex        =   82
      Text            =   " "
      Top             =   5640
      Width           =   615
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      DataField       =   "ch13"
      DataSource      =   "datPrimaryRS"
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
      Index           =   25
      Left            =   5280
      TabIndex        =   81
      Text            =   " "
      Top             =   5400
      Width           =   615
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      DataField       =   "cd13"
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
      Height          =   285
      Index           =   24
      Left            =   4680
      TabIndex        =   80
      Text            =   " "
      Top             =   5400
      Width           =   615
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      DataField       =   "ch12"
      DataSource      =   "datPrimaryRS"
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
      Index           =   23
      Left            =   5280
      TabIndex        =   79
      Text            =   " "
      Top             =   5160
      Width           =   615
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      DataField       =   "cd12"
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
      Height          =   285
      Index           =   22
      Left            =   4680
      TabIndex        =   78
      Text            =   " "
      Top             =   5160
      Width           =   615
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      DataField       =   "ch11"
      DataSource      =   "datPrimaryRS"
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
      Index           =   21
      Left            =   5280
      TabIndex        =   77
      Text            =   " "
      Top             =   4920
      Width           =   615
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      DataField       =   "cd11"
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
      Height          =   285
      Index           =   20
      Left            =   4680
      TabIndex        =   76
      Text            =   " "
      Top             =   4920
      Width           =   615
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      DataField       =   "ch10"
      DataSource      =   "datPrimaryRS"
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
      Index           =   19
      Left            =   5280
      TabIndex        =   75
      Text            =   " "
      Top             =   4680
      Width           =   615
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      DataField       =   "cd10"
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
      Height          =   285
      Index           =   18
      Left            =   4680
      TabIndex        =   74
      Text            =   " "
      Top             =   4680
      Width           =   615
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      DataField       =   "ch9"
      DataSource      =   "datPrimaryRS"
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
      Index           =   17
      Left            =   5280
      TabIndex        =   73
      Text            =   " "
      Top             =   4440
      Width           =   615
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      DataField       =   "cd9"
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
      Height          =   285
      Index           =   16
      Left            =   4680
      TabIndex        =   72
      Text            =   " "
      Top             =   4440
      Width           =   615
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      DataField       =   "ch8"
      DataSource      =   "datPrimaryRS"
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
      Index           =   15
      Left            =   5280
      TabIndex        =   71
      Text            =   " "
      Top             =   4200
      Width           =   615
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      DataField       =   "cd8"
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
      Height          =   285
      Index           =   14
      Left            =   4680
      TabIndex        =   70
      Text            =   " "
      Top             =   4200
      Width           =   615
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      DataField       =   "ch7"
      DataSource      =   "datPrimaryRS"
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
      Index           =   13
      Left            =   5280
      TabIndex        =   69
      Text            =   " "
      Top             =   3960
      Width           =   615
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      DataField       =   "cd7"
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
      Height          =   285
      Index           =   12
      Left            =   4680
      TabIndex        =   68
      Text            =   " "
      Top             =   3960
      Width           =   615
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      DataField       =   "ch6"
      DataSource      =   "datPrimaryRS"
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
      Left            =   5280
      TabIndex        =   67
      Text            =   " "
      Top             =   3720
      Width           =   615
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      DataField       =   "cd6"
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
      Height          =   285
      Index           =   10
      Left            =   4680
      TabIndex        =   66
      Text            =   " "
      Top             =   3720
      Width           =   615
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      DataField       =   "ch5"
      DataSource      =   "datPrimaryRS"
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
      Left            =   5280
      TabIndex        =   65
      Text            =   " "
      Top             =   3480
      Width           =   615
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      DataField       =   "cd5"
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
      Height          =   285
      Index           =   8
      Left            =   4680
      TabIndex        =   64
      Text            =   " "
      Top             =   3480
      Width           =   615
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      DataField       =   "ch4"
      DataSource      =   "datPrimaryRS"
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
      Left            =   5280
      TabIndex        =   63
      Text            =   " "
      Top             =   3240
      Width           =   615
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      DataField       =   "cd4"
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
      Height          =   285
      Index           =   6
      Left            =   4680
      TabIndex        =   62
      Text            =   " "
      Top             =   3240
      Width           =   615
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      DataField       =   "ch3"
      DataSource      =   "datPrimaryRS"
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
      Left            =   5280
      TabIndex        =   61
      Text            =   " "
      Top             =   3000
      Width           =   615
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      DataField       =   "cd3"
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
      Height          =   285
      Index           =   4
      Left            =   4680
      TabIndex        =   60
      Text            =   " "
      Top             =   3000
      Width           =   615
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      DataField       =   "ch2"
      DataSource      =   "datPrimaryRS"
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
      Left            =   5280
      TabIndex        =   59
      Text            =   " "
      Top             =   2760
      Width           =   615
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      DataField       =   "cd2"
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
      Height          =   285
      Index           =   2
      Left            =   4680
      TabIndex        =   58
      Text            =   " "
      Top             =   2760
      Width           =   615
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      DataField       =   "ch1"
      DataSource      =   "datPrimaryRS"
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
      Left            =   5280
      TabIndex        =   57
      Text            =   " "
      Top             =   2490
      Width           =   615
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      DataField       =   "cd1"
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
      Height          =   285
      Index           =   0
      Left            =   4680
      TabIndex        =   56
      Top             =   2490
      Width           =   615
   End
   Begin VB.TextBox Text1 
      DataField       =   "empresa"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Left            =   2040
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CheckBox automatico 
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   2640
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   2160
      Width           =   255
   End
   Begin MSDataGridLib.DataGrid DataGrid2 
      Bindings        =   "frmfacturacionventas.frx":0492
      Height          =   1335
      Left            =   8040
      TabIndex        =   47
      Top             =   2280
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
            LCID            =   2058
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
            LCID            =   2058
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
   Begin VB.CommandButton salir 
      Caption         =   "&Cerrar"
      Height          =   615
      Left            =   6480
      Picture         =   "frmfacturacionventas.frx":04AC
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   2880
      UseMaskColor    =   -1  'True
      Width           =   1095
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      DataField       =   "total"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """$"" #,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2058
         SubFormatType   =   2
      EndProperty
      DataSource      =   "datPrimaryRS"
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
      Index           =   15
      Left            =   6000
      Locked          =   -1  'True
      TabIndex        =   50
      TabStop         =   0   'False
      Top             =   3960
      Width           =   1935
   End
   Begin VB.CommandButton borrar 
      Caption         =   "&Borrar"
      Enabled         =   0   'False
      Height          =   615
      Left            =   6480
      Picture         =   "frmfacturacionventas.frx":08EE
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   1560
      UseMaskColor    =   -1  'True
      Width           =   1095
   End
   Begin VB.CommandButton Cancelar 
      Cancel          =   -1  'True
      Caption         =   "Cance&lar"
      Height          =   615
      Left            =   6480
      Picture         =   "frmfacturacionventas.frx":09F0
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   960
      UseMaskColor    =   -1  'True
      Width           =   1095
   End
   Begin VB.CommandButton nuevo 
      Caption         =   "&Nuevo"
      Height          =   615
      Left            =   6480
      Picture         =   "frmfacturacionventas.frx":0F22
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   360
      UseMaskColor    =   -1  'True
      Width           =   1095
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00FFFFFF&
      Height          =   840
      ItemData        =   "frmfacturacionventas.frx":1454
      Left            =   1200
      List            =   "frmfacturacionventas.frx":1456
      Sorted          =   -1  'True
      TabIndex        =   6
      Top             =   1320
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox tipocomp 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      DataField       =   "tipocompr"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Left            =   480
      Locked          =   -1  'True
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   1320
      Width           =   615
   End
   Begin VB.TextBox cuit 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      DataField       =   "cuit"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   960
      Width           =   1575
   End
   Begin VB.TextBox tipoiva 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      DataField       =   "tipoiva"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Left            =   480
      Locked          =   -1  'True
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   960
      Width           =   495
   End
   Begin VB.TextBox denominacion 
      BackColor       =   &H00E0E0E0&
      DataField       =   "cliente"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Left            =   480
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   600
      Width           =   3735
   End
   Begin MSMask.MaskEdBox Maskfecha 
      DataField       =   "fecha"
      DataSource      =   "datPrimaryRS"
      Height          =   255
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   450
      _Version        =   393216
      AllowPrompt     =   -1  'True
      Format          =   "dd/mm/yyyy"
      PromptChar      =   "_"
   End
   Begin MSAdodcLib.Adodc datPrimaryRS 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      Top             =   6435
      Visible         =   0   'False
      Width           =   14610
      _ExtentX        =   25770
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
      Connect         =   "Provider=MSDASQL.1;Password=25072004;Persist Security Info=True;User ID=lucva;Data Source=contable;Initial Catalog=contablesql"
      OLEDBString     =   "Provider=MSDASQL.1;Password=25072004;Persist Security Info=True;User ID=lucva;Data Source=contable;Initial Catalog=contablesql"
      OLEDBFile       =   ""
      DataSourceName  =   "contable"
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select libroventas.* from libroventas Order by fecha"
      Caption         =   " "
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
   Begin MSMask.MaskEdBox Maskcomprobante 
      DataField       =   "numcompr"
      DataSource      =   "datPrimaryRS"
      Height          =   255
      Left            =   480
      TabIndex        =   7
      Top             =   1680
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   450
      _Version        =   393216
      AllowPrompt     =   -1  'True
      PromptChar      =   "_"
   End
   Begin VB.CommandButton confcolumnas 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Conf.Libro"
      Height          =   735
      Left            =   6480
      Picture         =   "frmfacturacionventas.frx":1458
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   2160
      UseMaskColor    =   -1  'True
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H8000000A&
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
      Height          =   3615
      Left            =   6000
      TabIndex        =   48
      Top             =   0
      Width           =   1935
      Begin MSAdodcLib.Adodc datcuentas 
         Height          =   330
         Left            =   240
         Top             =   2880
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
         Connect         =   "Provider=MSDASQL.1;Password=25072004;Persist Security Info=True;User ID=lucva;Data Source=contable;Initial Catalog=contablesql"
         OLEDBString     =   "Provider=MSDASQL.1;Password=25072004;Persist Security Info=True;User ID=lucva;Data Source=contable;Initial Catalog=contablesql"
         OLEDBFile       =   ""
         DataSourceName  =   "contable"
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select listacuentas.* from listacuentas ORDER BY IDCUENTA"
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
   Begin MSAdodcLib.Adodc datcolumnas 
      Height          =   330
      Left            =   0
      Top             =   480
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
      Connect         =   "Provider=MSDASQL.1;Password=25072004;Persist Security Info=True;User ID=lucva;Data Source=contable;Initial Catalog=contablesql"
      OLEDBString     =   "Provider=MSDASQL.1;Password=25072004;Persist Security Info=True;User ID=lucva;Data Source=contable;Initial Catalog=contablesql"
      OLEDBFile       =   ""
      DataSourceName  =   "contable"
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   $"frmfacturacionventas.frx":189A
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
   Begin VB.TextBox Text2 
      DataField       =   "fecha"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Left            =   120
      TabIndex        =   46
      Text            =   "Text2"
      Top             =   960
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Text4 
      DataField       =   "numcompr"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Left            =   120
      TabIndex        =   49
      Text            =   "Text4"
      Top             =   1440
      Visible         =   0   'False
      Width           =   975
   End
   Begin MSDataGridLib.DataGrid DataGrid3 
      Bindings        =   "frmfacturacionventas.frx":18C8
      Height          =   735
      Left            =   8160
      TabIndex        =   52
      Top             =   3840
      Visible         =   0   'False
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   1296
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
            LCID            =   2058
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
            LCID            =   2058
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
   Begin MSAdodcLib.Adodc datclientes 
      Height          =   330
      Left            =   6120
      Top             =   1920
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
      Connect         =   "Provider=MSDASQL.1;Password=25072004;Persist Security Info=True;User ID=lucva;Data Source=contable;Initial Catalog=contablesql"
      OLEDBString     =   "Provider=MSDASQL.1;Password=25072004;Persist Security Info=True;User ID=lucva;Data Source=contable;Initial Catalog=contablesql"
      OLEDBFile       =   ""
      DataSourceName  =   "contable"
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select clientes.* from clientes ORDER BY razonsocial"
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
      BackColor       =   &H00C0C0C0&
      Height          =   2175
      Left            =   120
      TabIndex        =   45
      Top             =   0
      Width           =   4935
      Begin VB.TextBox Text6 
         DataField       =   "cerrado"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Left            =   3120
         TabIndex        =   55
         Text            =   "Text6"
         Top             =   120
         Visible         =   0   'False
         Width           =   855
      End
   End
   Begin MSAdodcLib.Adodc datmaestro 
      Height          =   330
      Left            =   6480
      Top             =   5400
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
   Begin MSAdodcLib.Adodc datasiento 
      Height          =   330
      Left            =   6480
      Top             =   5760
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
      Connect         =   "Provider=MSDASQL.1;Password=25072004;Persist Security Info=True;User ID=lucva;Data Source=contable;Initial Catalog=contablesql"
      OLEDBString     =   "Provider=MSDASQL.1;Password=25072004;Persist Security Info=True;User ID=lucva;Data Source=contable;Initial Catalog=contablesql"
      OLEDBFile       =   ""
      DataSourceName  =   "contable"
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
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
   Begin MSAdodcLib.Adodc datperiodo 
      Height          =   330
      Left            =   6480
      Top             =   6120
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
   Begin VB.Frame Frame3 
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
      Left            =   120
      TabIndex        =   101
      Top             =   2520
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "C.Deb"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   19
      Left            =   6240
      TabIndex        =   91
      Top             =   4320
      Width           =   615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "C.Hab"
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
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   16
      Left            =   6960
      TabIndex        =   90
      Top             =   4320
      Width           =   615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "C.Hab"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   18
      Left            =   5280
      TabIndex        =   87
      Top             =   2280
      Width           =   615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "C.Deb"
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
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   17
      Left            =   4680
      TabIndex        =   86
      Top             =   2280
      Width           =   615
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      DataField       =   "nomcol15"
      DataSource      =   "datcolumnas"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   14
      Left            =   120
      TabIndex        =   44
      Top             =   5880
      Width           =   3015
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      DataField       =   "nomcol14"
      DataSource      =   "datcolumnas"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   13
      Left            =   120
      TabIndex        =   43
      Top             =   5640
      Width           =   3015
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      DataField       =   "nomcol13"
      DataSource      =   "datcolumnas"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   12
      Left            =   120
      TabIndex        =   42
      Top             =   5400
      Width           =   3015
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      DataField       =   "nomcol12"
      DataSource      =   "datcolumnas"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   11
      Left            =   120
      TabIndex        =   41
      Top             =   5160
      Width           =   3015
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      DataField       =   "nomcol11"
      DataSource      =   "datcolumnas"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   10
      Left            =   120
      TabIndex        =   40
      Top             =   4920
      Width           =   3015
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      DataField       =   "nomcol10"
      DataSource      =   "datcolumnas"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   9
      Left            =   120
      TabIndex        =   39
      Top             =   4680
      Width           =   3015
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      DataField       =   "nomcol9"
      DataSource      =   "datcolumnas"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   8
      Left            =   120
      TabIndex        =   38
      Top             =   4440
      Width           =   3015
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      DataField       =   "nomcol8"
      DataSource      =   "datcolumnas"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   37
      Top             =   4200
      Width           =   3015
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      DataField       =   "nomcol7"
      DataSource      =   "datcolumnas"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   36
      Top             =   3960
      Width           =   3015
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      DataField       =   "nomcol6"
      DataSource      =   "datcolumnas"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   35
      Top             =   3720
      Width           =   3015
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      DataField       =   "nomcol5"
      DataSource      =   "datcolumnas"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   34
      Top             =   3480
      Width           =   3015
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      DataField       =   "nomcol4"
      DataSource      =   "datcolumnas"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   33
      Top             =   3240
      Width           =   3015
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      DataField       =   "nomcol3"
      DataSource      =   "datcolumnas"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   32
      Top             =   3000
      Width           =   3015
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      DataField       =   "nomcol2"
      DataSource      =   "datcolumnas"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   31
      Top             =   2760
      Width           =   3015
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      DataField       =   "nomcol1"
      DataSource      =   "datcolumnas"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   30
      Top             =   2520
      Width           =   3015
   End
   Begin VB.Label Label2 
      Caption         =   "Activar Calculo Automatico "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   120
      TabIndex        =   54
      Top             =   2160
      Width           =   2415
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "TOTAL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   15
      Left            =   6000
      TabIndex        =   51
      Top             =   3720
      Width           =   1935
   End
End
Attribute VB_Name = "frmfacturacionventas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim columna(15) As String
Dim pos As Integer
Dim posicion As Integer
Dim poscuenta As Integer
Dim cuenta As Integer
Dim proximafactura As String
Dim proximafactura0 As String
Dim regactivo As Integer
Dim tipofa As String
Dim tipofb As String
Dim tipora As String
Dim tiporb As String
Dim tiponca As String
Dim tiponcb As String
Dim tiponda As String
Dim tipondb As String
Dim sumadebe As Currency
Dim sumahaber As Currency
Dim errorasiento As Boolean
Dim facanulada As String
Dim ter(15, 15), sig(15, 15) As String



Private Sub anulada_Click()

    Respuesta = MsgBox("Esta por anular un comprobante, esta Ud. Seguro ?", vbYesNo, "!! Atencion !!")
If Respuesta = vbYes Then
    denominacion.Text = "***ANULADA***"
    facanulada = "s"
    cuit.Text = ""
    tipoiva.Text = ""
    Text6.Text = "N"
    Text3(15).Text = 0
    For x = 0 To 14
        Text3(x).Text = 0
    Next x
    datPrimaryRS.Recordset.Fields(60) = login.iper
    datPrimaryRS.Recordset.Fields(61) = login.fper
    datPrimaryRS.Recordset.UpdateBatch adAffectCurrent
    pos = datPrimaryRS.Recordset.AbsolutePosition
    datPrimaryRS.Refresh
    datPrimaryRS.Recordset.AbsolutePosition = pos
    nuevo.SetFocus
End If
  
End Sub

Private Sub borrar_Click()


End Sub

Private Sub buscafact_Click()
 
  datPrimaryRS.RecordSource = "select libroventas.* from libroventas WHERE libroventas.empresa = " & login.empresaact & " and tipocompr = 'F-A' and inicioper = '" & login.iper & "' and finper = '" & login.fper & "' Order by id"
  datPrimaryRS.Refresh
  If datPrimaryRS.Recordset.EOF = True Then
    tipofa = "0001-00000000"
    GoTo fb
  End If
  datPrimaryRS.RecordSource = "select libroventas.* from libroventas WHERE libroventas.empresa = " & login.empresaact & " and cerrado = 'N' Order by id"
  datPrimaryRS.Refresh
  datPrimaryRS.Recordset.MoveLast
  tipofa = Text4.Text
  datPrimaryRS.RecordSource = "select libroventas.* from libroventas WHERE libroventas.empresa = " & login.empresaact & " and cerrado = 'N' Order by id"
  datPrimaryRS.Refresh
fb:
  datPrimaryRS.RecordSource = "select libroventas.* from libroventas WHERE libroventas.empresa = " & login.empresaact & " and tipocompr = 'F-B' and inicioper = '" & login.iper & "' and finper = '" & login.fper & "' Order by id"
  datPrimaryRS.Refresh
  If datPrimaryRS.Recordset.EOF = True Then
    tipofb = "0001-00000000"
    GoTo ra
  End If
  datPrimaryRS.Recordset.MoveLast
  tipofb = Text4.Text
ra:
  datPrimaryRS.RecordSource = "select libroventas.* from libroventas WHERE libroventas.empresa = " & login.empresaact & " and tipocompr = 'R-A' and inicioper = '" & login.iper & "' and finper = '" & login.fper & "' Order by id"
  datPrimaryRS.Refresh
  If datPrimaryRS.Recordset.EOF = True Then
    tipora = "0001-00000000"
    GoTo rb
  End If
  datPrimaryRS.Recordset.MoveLast
  tipora = Text4.Text
rb:
  datPrimaryRS.RecordSource = "select libroventas.* from libroventas WHERE libroventas.empresa = " & login.empresaact & " and tipocompr = 'R-B' and inicioper = '" & login.iper & "' and finper = '" & login.fper & "' Order by id"
  datPrimaryRS.Refresh
  If datPrimaryRS.Recordset.EOF = True Then
    tiporb = "0001-00000000"
    GoTo nca
  End If
  datPrimaryRS.Recordset.MoveLast
  tiporb = Text4.Text
nca:
  datPrimaryRS.RecordSource = "select libroventas.* from libroventas WHERE libroventas.empresa = " & login.empresaact & " and tipocompr = 'NCA' and inicioper = '" & login.iper & "' and finper = '" & login.fper & "' Order by id"
  datPrimaryRS.Refresh
  If datPrimaryRS.Recordset.EOF = True Then
    tiponca = "0001-00000000"
    GoTo ncb
  End If
  datPrimaryRS.Recordset.MoveLast
  tiponca = Text4.Text
ncb:
  datPrimaryRS.RecordSource = "select libroventas.* from libroventas WHERE libroventas.empresa = " & login.empresaact & " and tipocompr = 'NCB' and inicioper = '" & login.iper & "' and finper = '" & login.fper & "' Order by id"
  datPrimaryRS.Refresh
  If datPrimaryRS.Recordset.EOF = True Then
    tiponcb = "0001-00000000"
    GoTo nda
  End If
  datPrimaryRS.Recordset.MoveLast
  tiponcb = Text4.Text
nda:
  datPrimaryRS.RecordSource = "select libroventas.* from libroventas WHERE libroventas.empresa = " & login.empresaact & " and tipocompr = 'NDA' and inicioper = '" & login.iper & "' and finper = '" & login.fper & "' Order by id"
  datPrimaryRS.Refresh
  If datPrimaryRS.Recordset.EOF = True Then
    tiponda = "0001-00000000"
    GoTo ndb
  End If
  datPrimaryRS.Recordset.MoveLast
  tiponda = Text4.Text
ndb:
  datPrimaryRS.RecordSource = "select libroventas.* from libroventas WHERE libroventas.empresa = " & login.empresaact & " and tipocompr = 'NCB' and inicioper = '" & login.iper & "' and finper = '" & login.fper & "' Order by id"
  datPrimaryRS.Refresh
  If datPrimaryRS.Recordset.EOF = True Then
    tipondb = "0001-00000000"
    GoTo sigue
  End If
  datPrimaryRS.Recordset.MoveLast
  tipondb = Text4.Text
sigue:
datPrimaryRS.RecordSource = "select libroventas.* from libroventas WHERE libroventas.empresa = " & login.empresaact & " and cerrado = 'N' and inicioper = '" & login.iper & "' and finper = '" & login.fper & "' Order by id"
datPrimaryRS.Refresh

End Sub

Private Sub calcular_Click()
On Error GoTo erroRcalcular

sumar = 0
parcial = 0
result = 0
For t = 0 To 14
        For x = 1 To 15
            If sig(t, x) = "" And x = 1 Then GoTo paso1
            If sumar = 1 Then parcial = result
            If sig(t, x) = "*" Then result = Text3(ter(t, x) - 1).Text * Val(ter(t, x + 1))
            If sig(t, x) = "/" Then result = Text3(ter(t, x) - 1).Text / Val(ter(t, x + 1))
            If sig(t, x) = "+" Or sig(t, x) = "" Then
                    result = parcial + result
                    sumar = 1
                    If sig(t, x) = "" Then
                            sumar = 0
                            parcial = 0
                    End If
                    GoTo paso0
            Else
                    sumar = 0
            End If
            If sig(t, x) = "-" Or sig(t, x) = "" Then
                    result = parcial - result
                    sumar = 1
                    If sig(t, x) = "" Then
                            sumar = 0
                            parcial = 0
                    End If
                    GoTo paso0
            Else
                    sumar = 0
            End If
paso0:
        Text3(t).Text = result
         Next x
       
paso1:
Next t
Exit Sub

erroRcalcular:
    mensa = MsgBox("Error al realizar calculo automatico, revisar configuracion de columnas del libro", vbCritical, "!! Atención !!")


End Sub

Private Sub Cancelar_Click()

        datPrimaryRS.Refresh
        If datPrimaryRS.Recordset.EOF = True Then Exit Sub
        datPrimaryRS.Recordset.MoveLast

End Sub

Private Sub confcolumnas_Click()
    
    Unload Me
    frmcolumnasventa.Show


End Sub




Private Sub DataList1_KeyPress(KeyAscii As Integer)

    If KeyAscii = 45 Then
        Maskfecha.SetFocus
    End If

    If KeyAscii = 13 Then
        If DataList1.SelectedItem <> "" Then
            datclientes.Recordset.Bookmark = DataList1.SelectedItem
            denominacion = datclientes.Recordset.Fields(3)
            If datclientes.Recordset.Fields(4) <> "" Then tipoiva.Text = datclientes.Recordset.Fields(4)
            If datclientes.Recordset.Fields(5) <> "" Then cuit.Text = datclientes.Recordset.Fields(5)
            pruebanulo = IsNull(datclientes.Recordset.Fields(12))
            If pruebanulo = True Then
                prueba = IsNull(datcolumnas.Recordset.Fields(61))
                If prueba = True Then
                    Text7(30).Text = 0
                Else
                    Text7(30).Text = datcolumnas.Recordset.Fields(61)
                End If
            Else
                Text7(30).Text = datclientes.Recordset.Fields(12)
            End If
            List1.Visible = True
            List1.SetFocus
        Else
            Exit Sub
        End If
    End If

End Sub

Private Sub DataList1_LostFocus()
  DataList1.Visible = False
End Sub

Private Sub DataList2_Click()
    Text7(poscuenta).Text = DataList2.BoundText
End Sub

Private Sub DataList2_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        Text7(poscuenta).Text = DataList2.BoundText

            If Text7(poscuenta).Text = "" Then Text7(poscuenta).Text = 0
             
              If Text7(poscuenta).Text = 0 Then
                    mensa = MsgBox("Debe ingresar un Nº de cuenta", vbCritical, "!! Error !!")
                    Text7(poscuenta).SetFocus
                    errorasiento = True
                    Exit Sub
              End If
              errorasiento = False
                                                                             
              If datccostos.Recordset.EOF = True Then GoTo sigue
              datccostos.Recordset.MoveFirst
              digito = Val(datccostos.Recordset.Fields(4))
              digcue = Val(Mid(Text7(poscuenta).Text, 1, 1))
              If digcue = digito And login.habcc = True Then
                DataList3.Visible = True
                Text9.Visible = True
                Frame3.Visible = True
                DataList3.SetFocus
                Exit Sub
              End If
                                                                             
              If poscuenta < 30 Then sumahaber = Text3(posicion) + sumahaber
              If Text3(posicion + 1).Visible = True Then
                    Text3(posicion + 1).SetFocus
              Else
                    datPrimaryRS.Recordset.UpdateBatch adAffectCurrent
                    pos = datPrimaryRS.Recordset.AbsolutePosition
                    datPrimaryRS.Refresh
                    datPrimaryRS.Recordset.AbsolutePosition = pos
                    Text7(30).SetFocus
              End If
                    
sigue:
              If poscuenta = 30 Then
                    sumadebe = Text3(15) + sumadebe
                    datPrimaryRS.Recordset.UpdateBatch adAffectCurrent
                    pos = datPrimaryRS.Recordset.AbsolutePosition
                    datPrimaryRS.Refresh
                    datPrimaryRS.Recordset.AbsolutePosition = pos
                    DataList2.Visible = False
                    nuevo.SetFocus
              End If

              Exit Sub
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
              If poscuenta < 30 Then sumahaber = Text3(posicion) + sumahaber
              If Text3(posicion + 1).Visible = True Then
                    Text3(posicion + 1).SetFocus
              Else
                    datPrimaryRS.Recordset.UpdateBatch adAffectCurrent
                    pos = datPrimaryRS.Recordset.AbsolutePosition
                    datPrimaryRS.Refresh
                    datPrimaryRS.Recordset.AbsolutePosition = pos
                    Text7(30).SetFocus
              End If
              If poscuenta = 30 Then
                    sumadebe = Text3(15) + sumadebe
                    datPrimaryRS.Recordset.UpdateBatch adAffectCurrent
                    pos = datPrimaryRS.Recordset.AbsolutePosition
                    datPrimaryRS.Refresh
                    datPrimaryRS.Recordset.AbsolutePosition = pos
                    DataList3.Visible = False
                    nuevo.SetFocus
              End If
    End If
    

End Sub

Private Sub DataList3_LostFocus()

    If Text9.Text = "" Then
        mensa = MsgBox("Debe ingresa un Centro de Costo", vbCritical, "!Error¡")
        DataList3.SetFocus
        Exit Sub
    End If
Frame3.Visible = False
Text9.Visible = False
DataList3.Visible = False

End Sub

Private Sub denominacion_GotFocus()

    DataList1.Visible = True
    DataList1.SetFocus

End Sub

Private Sub Form_GotFocus()
    Maskcomprobante.Mask = ""
    Maskfecha.Mask = ""
End Sub

Private Sub Form_Load()
On Error GoTo errorform


    Call buscafact_Click

  datPrimaryRS.RecordSource = "select libroventas.* from libroventas WHERE libroventas.empresa = " & login.empresaact & " and inicioper = '" & login.iper & "' and finper = '" & login.fper & "' and cerrado = 'N' Order by id"
  datPrimaryRS.Refresh
  
  datcolumnas.RecordSource = "SELECT columnasventa.* FROM columnasventa WHERE empresa = " & login.empresaact & " and inicioper = '" & login.iper & "'"
  datcolumnas.Refresh
  
  datcuentas.RecordSource = "select listacuentas.* from listacuentas WHERE empre = " & login.empresaact & " and inicioper = '" & login.iper & "' and finper = '" & login.fper & "' ORDER BY IDCUENTA"
  datcuentas.Refresh
  
  datclientes.RecordSource = "select clientes.* from clientes Where empresa = " & login.empresaact & " ORDER BY razonsocial"
  datclientes.Refresh
  
  datlistacostos.RecordSource = "SELECT listaccostos.* FROM listaccostos WHERE empresa = " & login.empresaact & ""
  datlistacostos.Refresh
  
  datccostos.RecordSource = "SELECT ccostos.* FROM ccostos WHERE empresa = " & login.empresaact & ""
  datccostos.Refresh
  
  datmaestro.RecordSource = "select [Maestro Asientos].* from [Maestro Asientos] where empresa = " & login.empresaact & " and perinicial = '" & login.iper & "' and perfinal = '" & login.fper & "' order by nroasiento"
  datmaestro.Refresh
  datperiodo.RecordSource = "select EMPRESA.* from EMPRESA where empresa = " & login.empresaact & ""
  datperiodo.Refresh
  datasiento.RecordSource = "select [Detalle Asientos].* from [Detalle Asientos] where empresa = " & login.empresaact & ""
  datasiento.Refresh
  
automatico.Value = 1
datcolumnas.Refresh
For x = 1 To 15
verif = IsNull(datcolumnas.Recordset.Fields(x * 2))
If verif = False Then columna(x) = datcolumnas.Recordset.Fields(x * 2)
Next x

For x = 0 To 14
    For y = 1 To 15
        ter(x, y) = ""
    Next y
Next x

t = 1
For c = 0 To 14
        S = 0
        t = 1
        For x = 1 To Len(columna(c + 1))
        car = Mid(columna(c + 1), x, 1)
        If car = "-" Or car = "+" Or car = "*" Or car = "/" Then
            S = S + 1
            sig(c, S) = car
            t = t + 1
            GoTo paso1
        End If
        If car = "C" Or car = "c" Then GoTo paso1
        ter(c, t) = ter(c, t) + car
paso1:
        Next x
Next c

    If datPrimaryRS.Recordset.EOF = True Then
        datPrimaryRS.Recordset.AddNew
        Text1.Text = login.empresaact
            Maskfecha.SelLength = 10
            Maskfecha.SelText = ""
            Maskcomprobante.SelLength = 13
            Maskcomprobante.SelText = ""
        For x = 0 To 14
            Text3(x).Text = 0
        Next x
     Else
       datPrimaryRS.Recordset.MoveLast
     End If
     Maskfecha.Mask = "##/##/####"
     Maskfecha.MaxLength = 10
     Maskcomprobante.Mask = "####-########"
     Maskcomprobante.MaxLength = 13
    
    Text1.Text = login.empresaact
        
    List1.AddItem "F-A"
    List1.AddItem "F-B"
    List1.AddItem "R-A"
    List1.AddItem "R-B"
    List1.AddItem "NDA"
    List1.AddItem "NDB"
    List1.AddItem "NCA"
    List1.AddItem "NCB"
            
    For x = 0 To 14
        valida = IsNull(datcolumnas.Recordset.Fields(x * 2 + 1))
        If datcolumnas.Recordset.Fields(x * 2 + 1) = "" Then valida = True
        If valida = False Then
            Label1(x).Caption = datcolumnas.Recordset.Fields(x * 2 + 1)
        Else
            Label1(x).Visible = False
            Text3(x).Text = 0
            Text3(x).Visible = False
            Text7(x * 2).Text = 0
            Text7(x * 2 + 1).Text = 0
            Text7(x * 2).Visible = False
            Text7(x * 2 + 1).Visible = False
        End If
    Next x
Exit Sub

errorform:
    mensa = MsgBox("Error de Codificacion", vbCritical, "!! Error !!")
     
End Sub

Private Sub Form_Resize()
  On Error Resume Next
  'Esto cambiará el tamaño de la cuadrícula al cambiar el tamaño del formulario
  grdDataGrid.Height = Me.ScaleHeight - datPrimaryRS.Height - 30
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



Private Sub List1_KeyPress(KeyAscii As Integer)

    If KeyAscii = 45 Then
        DataList1.Visible = True
        DataList1.SetFocus
    End If
    If KeyAscii = 13 Then
        tipocomp.Text = List1.Text
        Maskcomprobante.Enabled = True
        Maskcomprobante.SetFocus
    End If
                 

End Sub

Private Sub List1_LostFocus()
Dim proximafactura1 As Double

    If List1.Text = "" Then
        List1.SetFocus
        Exit Sub
    End If

If List1.Text = "F-A" Then Text4.Text = tipofa
If List1.Text = "F-B" Then Text4.Text = tipofb
If List1.Text = "R-A" Then Text4.Text = tipora
If List1.Text = "R-B" Then Text4.Text = tiporb
If List1.Text = "NCA" Then Text4.Text = tiponca
If List1.Text = "NCB" Then Text4.Text = tiponcb
If List1.Text = "NDA" Then Text4.Text = tiponda
If List1.Text = "NDB" Then Text4.Text = tipondb

If tipofa = "0001-00000001" Then Maskcomprobante.Enabled = True
If tipofb = "0001-00000001" Then Maskcomprobante.Enabled = True
If tipora = "0001-00000001" Then Maskcomprobante.Enabled = True
If tiporb = "0001-00000001" Then Maskcomprobante.Enabled = True
If tiponca = "0001-00000001" Then Maskcomprobante.Enabled = True
If tiponcb = "0001-00000001" Then Maskcomprobante.Enabled = True
If tiponda = "0001-00000001" Then Maskcomprobante.Enabled = True
If tipondb = "0001-00000001" Then Maskcomprobante.Enabled = True

  
proximafactura0 = Left(Text4.Text, 4)
proximafactura1 = Val(Right((Text4.Text), 8)) + 1
proximafactura = Mid("00000000", 1, 9 - Len(Str(proximafactura1))) + Right(Str(proximafactura1), Len(Str(proximafactura1)) - 1)
proximafactura = proximafactura0 + "-" + proximafactura

List1.Visible = False

End Sub

Private Sub Maskcomprobante_GotFocus()
On Error GoTo error1
    If List1.Text = "" Then
        List1.SetFocus
        Exit Sub
    End If

    If proximafactura = "" Then Exit Sub
    Maskcomprobante.Text = proximafactura
    If Maskcomprobante.Text <> "0001-00000001" Then Text3(0).SetFocus
    
error1:
    
End Sub

Private Sub Maskcomprobante_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 45 Then
            List1.Visible = True
            List1.SetFocus
    End If
    If KeyAscii = 13 Then
        KeyAscii = 0
            Text4.Text = Maskcomprobante.Text
            Text3(0).SelStart = 0
            Text3(0).SelLength = Len(Text3(0))
            Text3(0).SetFocus
            If Val(Mid(Text4.Text, 1, 4)) = 0 Then
                mensa = MsgBox("Debe ingresar una sucursal en el Nro de factura", vbCritical, "!! Atención !!")
                Maskcomprobante.SetFocus
                Maskcomprobante.SelStart = 0
                Maskcomprobante.SelLength = 4
                Exit Sub
            End If
            If Right(Maskcomprobante.Text, 1) = "_" Then
                mensa = MsgBox("Nro de factura incorrecto", vbCritical, "!! Atención !!")
                Maskcomprobante.SetFocus
                Maskcomprobante.SelStart = 5
                Maskcomprobante.SelLength = 8
                Exit Sub
            End If
    End If
    
    
End Sub


Private Sub Maskcomprobante_LostFocus()
   Rem  Maskcomprobante.Enabled = False
End Sub

Private Sub Maskfecha_GotFocus()
    
  datclientes.RecordSource = "select clientes.* from clientes Where empresa = " & login.empresaact & " ORDER BY razonsocial"
  datclientes.Refresh

End Sub

Private Sub Maskfecha_KeyPress(KeyAscii As Integer)
Dim dia As Integer
Dim mes As Integer
Dim año As Integer

    If KeyAscii = 13 Then
        Text2.Text = Maskfecha.Text
        dia = Day(Date)
        mes = Month(Date)
        año = Year(Date)
        If Val(Mid(Text2.Text, 1, 2)) > dia And Val(Mid(Text2.Text, 4, 2)) >= mes And Val(Mid(Text2.Text, 7, 4)) >= año Then
                mensa = MsgBox("El Día ingresado es mayor al de la fecha actual", vbCritical, "!! Atención !!")
                Maskfecha.SetFocus
                Maskfecha.SelStart = 0
                Maskfecha.SelLength = 2
                Exit Sub
        End If
        If Val(Mid(Text2.Text, 4, 2)) > mes And Val(Mid(Text2.Text, 7, 4)) >= año Then
                mensa = MsgBox("El Mes ingresado es mayor al de la fecha actual", vbCritical, "!! Atención !!")
                Maskfecha.SetFocus
                Maskfecha.SelStart = 3
                Maskfecha.SelLength = 2
                Exit Sub
        End If
        If Val(Mid(Text2.Text, 7, 4)) > año Then
                mensa = MsgBox("El Año ingresado es mayor al de la fecha actual", vbCritical, "!! Atención !!")
                Maskfecha.SetFocus
                Maskfecha.SelStart = 6
                Maskfecha.SelLength = 4
                Exit Sub
        End If
        denominacion.SetFocus
    End If
    

End Sub

Private Sub nuevo_GotFocus()

    If facanulada = "s" Then
        facanulada = ""
        Exit Sub
    End If
    If sumadebe <> sumahaber Or errorasiento = True Then
        mensa = MsgBox("El asiento está desvalanceado, no se puede grabar", vbCritical, "!! Error !!")
            sumadebe = 0
            sumahaber = 0
            errorasiento = True
            Text3(0).SetFocus
            Exit Sub
    End If

  errorasiento = False
 

Rem ****************** grabar asiento

    campoaño = Right(Maskfecha.Text, 4)
    campomes = Mid(Maskfecha.Text, 4, 2)
    campodia = Left(Maskfecha.Text, 2)
    campofecha = campoaño + "/" + campomes + "/" + campodia
    
    campoaño1 = Right(Text5(0).Text, 4)
    campomes1 = Mid(Text5(0).Text, 4, 2)
    campodia1 = Left(Text5(0).Text, 2)
    campofecha1 = campoaño1 + "/" + campomes1 + "/" + campodia1
    
    campoaño2 = Right(Text5(1).Text, 4)
    campomes2 = Mid(Text5(1).Text, 4, 2)
    campodia2 = Left(Text5(1).Text, 2)
    campofecha2 = campoaño2 + "/" + campomes2 + "/" + campodia2

    If campofecha < campofecha1 Or campofecha > campofecha2 Then
            mensa = MsgBox("La Fecha es erronea o no pertenecia al periodo en ejercicio", vbCritical, "!! Atención !!")
            Maskfecha.SelLength = 10
            Maskfecha.SetFocus
            Exit Sub
    End If
    modifica = 0
    If Text8.Text = "S" Then
        modifica = 1
        nroasie = Text10.Text
        datmaestro.RecordSource = "select [Maestro Asientos].* from [Maestro Asientos] where empresa = " & login.empresaact & " and nroasiento = " & nroasie & " and perinicial = '" & login.iper & "' and perfinal = '" & login.fper & "' order by nroasiento"
        datmaestro.Refresh
        masterasiento = datmaestro.Recordset.Fields(2)
        datasiento.RecordSource = "select [Detalle Asientos].* from [Detalle Asientos] where empresa = " & login.empresaact & " and idmasterasientos = " & masterasiento & ""
        datasiento.Refresh
        If datasiento.Recordset.EOF = True Then
            datmaestro.Recordset.Delete adAffectCurrent
            GoTo pas1
        End If
        datasiento.Recordset.MoveFirst
pas0:
        datasiento.Recordset.Delete adAffectCurrent
        datasiento.Recordset.MoveNext
        If datasiento.Recordset.EOF = True Then
            datmaestro.Recordset.Delete adAffectCurrent
            GoTo pas1
        End If
        GoTo pas0
    End If

    If datmaestro.Recordset.EOF = False Then
            datmaestro.RecordSource = "select [Maestro Asientos].* from [Maestro Asientos] where empresa = " & login.empresaact & " and perinicial = '" & login.iper & "' and perfinal = '" & login.fper & "' order by nroasiento"
            datmaestro.Refresh
            datmaestro.Recordset.MoveLast
            nroasie = datmaestro.Recordset.Fields(3) + 1
            If Text8.Text = "S" Then nroasie = nroasie - 1
    Else
            nroasie = 1
    End If
    
pas1:
    datmaestro.Recordset.AddNew
    datmaestro.Recordset.Fields(0) = Maskfecha.Text
    datmaestro.Recordset.Fields(1) = Date
    datmaestro.Recordset.Fields(3) = nroasie
    datmaestro.Recordset.Fields(4) = Left(denominacion.Text, 20) + " " + tipocomp.Text + " Nº:" + Maskcomprobante.Text
    datmaestro.Recordset.Fields(5) = Text5(0).Text
    datmaestro.Recordset.Fields(6) = Text5(1).Text
    datmaestro.Recordset.Fields(7) = login.empresaact
    datmaestro.Recordset.Fields(8) = "N"
    datmaestro.Recordset.Fields(9) = Val(datPrimaryRS.Recordset.Fields(0))
    datmaestro.Recordset.Fields(10) = "V"
    datmaestro.Recordset.Fields(11) = "S"
    datmaestro.Recordset.UpdateBatch adAffectCurrent
     
For x = 0 To 28 Step 2
    If Text7(x + 1).Text <> "" Then
            If Text3(x / 2).Visible = False Then GoTo paso1
            datasiento.Recordset.AddNew
            datasiento.Recordset.Fields(5) = Val(datmaestro.Recordset.Fields(2))
            datasiento.Recordset.Fields(1) = datmaestro.Recordset.Fields(0).Value
            datasiento.Recordset.Fields(7) = login.empresaact
            datasiento.Recordset.Fields(2) = Text7(x + 1).Text
            datasiento.Recordset.Fields(4) = Text3(x / 2).Text
            datasiento.Recordset.Fields(6) = Label1(x / 2).Caption
            If Text9.Text <> "" Then datasiento.Recordset.Fields(8) = Text9.Text
            datasiento.Recordset.UpdateBatch adAffectCurrent
    End If
Next x
paso1:
            datasiento.Recordset.AddNew
            datasiento.Recordset.Fields(5) = Val(datmaestro.Recordset.Fields(2))
            datasiento.Recordset.Fields(1) = datmaestro.Recordset.Fields(0).Value
            datasiento.Recordset.Fields(7) = login.empresaact
            datasiento.Recordset.Fields(2) = Text7(30).Text
            datasiento.Recordset.Fields(3) = Text3(15).Text
            datasiento.Recordset.Fields(6) = "Total facturado"
            datasiento.Recordset.UpdateBatch adAffectCurrent

    Text8.Text = "S"
    If Left(tipocomp, 2) = "NC" Then
        Text3(15).Text = Text3(15).Text * -1
        For x = 1 To 14
            If Text3(x).Visible = True And Text3(x).Text <> 0 Then
                Text3(x).Text = Text3(x).Text * -1
            End If
        Next x
    End If
    datPrimaryRS.Recordset.Fields(59) = nroasie
    datPrimaryRS.Recordset.UpdateBatch adAffectCurrent
    pos = datPrimaryRS.Recordset.AbsolutePosition
    datPrimaryRS.Refresh
    datPrimaryRS.Recordset.AbsolutePosition = pos

    
End Sub

Private Sub nuevo_Click()

Dim proximafactura1 As Double

    If errorasiento = True Then Exit Sub
    
    Call buscafact_Click
    
    datPrimaryRS.Recordset.AddNew
    For x = 0 To 14
            Text3(x).Text = 0
    Next x
    For x = 0 To 31
            Text7(x).Text = ""
    Next x
    Text1.Text = login.empresaact
    
    Maskfecha.SelLength = 10
    Maskfecha.SelText = ""
    Maskcomprobante.SelLength = 13
    Maskcomprobante.SelText = ""
    Maskfecha.SetFocus
    regactivo = datPrimaryRS.Recordset.AbsolutePosition
    

End Sub

Private Sub salir_Click()

    If errorasiento = True Then
        mensa = MsgBox("El asiento está desvalanceado, no se puede grabar", vbCritical, "!! Error !!")
            sumadebe = 0
            sumahaber = 0
            Text3(0).SetFocus
            Exit Sub
    End If
    
    errorasiento = False
    Call Cancelar_Click
    Unload Me

End Sub


Private Sub Text3_GotFocus(Index As Integer)
On Error GoTo error1
    If List1.Text = "" Then
        List1.SetFocus
        Exit Sub
    End If
                
                Text3(Index).SelStart = 0
                Text3(Index).SelLength = Len(Text3(Index)) + 3
error1:
End Sub

Private Sub Text3_KeyPress(Index As Integer, KeyAscii As Integer)
Dim suma As Double
    
    If KeyAscii = 45 Then
            KeyAscii = 0
            If Index > 0 Then
                If Text3(Index - 1).Visible = True Then
                    Text3(Index - 1).SetFocus
                    Text3(Index - 1).SelStart = 0
                    Text3(Index - 1).SelLength = Len(Text3(Index - 1))
                End If
            End If
    End If
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        If Text3(Index).Text = "" Then Text3(Index).Text = 0
        suma = 0

         If automatico = 1 Then Call calcular_Click
                                                            
         For x = 0 To 14
            If Text3(x).Visible = True Then suma = Text3(x).Text + suma
         Next x
         Text3(15).Text = suma
         Text6.Text = "N"
         
         datPrimaryRS.Recordset.Fields(60) = login.iper
         datPrimaryRS.Recordset.Fields(61) = login.fper
           
         datPrimaryRS.Recordset.UpdateBatch adAffectCurrent
         pos = datPrimaryRS.Recordset.AbsolutePosition
         datPrimaryRS.Refresh
         datPrimaryRS.Recordset.AbsolutePosition = pos
         
         
         If Text3(Index).Text > 0 Then
            posicion = Index
            cuenta = 0
            Text7(Index * 2 + 1).SetFocus
            Exit Sub
         End If
         If Text3(Index + 1).Visible = True Then
                Text3(Index + 1).SetFocus
         Else
                Text7(30).SetFocus
         End If
     End If
     
    
End Sub




Private Sub Text7_GotFocus(Index As Integer)

    prueba = datcolumnas.Recordset.Fields(Index + 31)
    If prueba > 0 Then
           Text7(Index).Text = prueba
           sumahaber = Text3(posicion) + sumahaber
           If Text3(posicion + 1).Visible = True Then
                    Text3(posicion + 1).SetFocus
                    Exit Sub
           End If
           Text7(30).SetFocus
           Exit Sub
    End If

    poscuenta = Index
    Text7(Index).SelLength = Len(Text7(Index))
    DataList2.BoundText = Text7(Index).Text
    DataList2.Visible = True
    DataList2.Left = Text7(Index).Left + Text7(Index).Width - DataList2.Width
    DataList2.Top = Text7(Index).Top + Text7(Index).Height
    DataList2.SetFocus
    
End Sub



Private Sub Text7_LostFocus(Index As Integer)

    poscuenta = Index
    If Index = 30 Then
        If Text7(Index).Text = 0 Or Text7(Index).Text = "" Or Left(tipocomp, 2) = "NC" Then Exit Sub
        sumadebe = Text3(15) + sumadebe
        nuevo.SetFocus
        Exit Sub
    End If
         
End Sub
