VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmcolumnasventa0 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configuración Columnas Libro Ventas"
   ClientHeight    =   7590
   ClientLeft      =   1095
   ClientTop       =   330
   ClientWidth     =   10365
   Icon            =   "frmcolumnasventa0.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7590
   ScaleWidth      =   10365
   Begin VB.TextBox Text5 
      BackColor       =   &H00000000&
      DataField       =   "Nombre Cuenta"
      DataSource      =   "datcuentas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   375
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   53
      Top             =   0
      Visible         =   0   'False
      Width           =   3375
   End
   Begin MSDataListLib.DataList DataList2 
      Bindings        =   "frmcolumnasventa0.frx":0442
      Height          =   2010
      Left            =   6000
      TabIndex        =   54
      Top             =   0
      Visible         =   0   'False
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   3545
      _Version        =   393216
      MatchEntry      =   -1  'True
      BackColor       =   0
      ForeColor       =   -2147483643
      ListField       =   "Cod Contable"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      DataField       =   "cd2"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   1
      Left            =   8520
      TabIndex        =   85
      Text            =   " "
      Top             =   960
      Width           =   735
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      DataField       =   "ch15"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   14
      Left            =   9360
      TabIndex        =   84
      Text            =   " "
      Top             =   5640
      Width           =   735
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      DataField       =   "ch14"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   13
      Left            =   9360
      TabIndex        =   83
      Text            =   " "
      Top             =   5280
      Width           =   735
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      DataField       =   "ch13"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   12
      Left            =   9360
      TabIndex        =   82
      Text            =   " "
      Top             =   4920
      Width           =   735
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      DataField       =   "ch12"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   11
      Left            =   9360
      TabIndex        =   81
      Text            =   " "
      Top             =   4560
      Width           =   735
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      DataField       =   "ch11"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   10
      Left            =   9360
      TabIndex        =   80
      Text            =   " "
      Top             =   4200
      Width           =   735
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      DataField       =   "ch10"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   9
      Left            =   9360
      TabIndex        =   79
      Text            =   " "
      Top             =   3840
      Width           =   735
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      DataField       =   "ch9"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   8
      Left            =   9360
      TabIndex        =   78
      Text            =   " "
      Top             =   3480
      Width           =   735
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      DataField       =   "ch8"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   7
      Left            =   9360
      TabIndex        =   77
      Text            =   " "
      Top             =   3120
      Width           =   735
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      DataField       =   "ch7"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   6
      Left            =   9360
      TabIndex        =   76
      Text            =   " "
      Top             =   2760
      Width           =   735
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      DataField       =   "ch6"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   5
      Left            =   9360
      TabIndex        =   75
      Text            =   " "
      Top             =   2400
      Width           =   735
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      DataField       =   "ch5"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   4
      Left            =   9360
      TabIndex        =   74
      Text            =   " "
      Top             =   2040
      Width           =   735
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      DataField       =   "ch4"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   3
      Left            =   9360
      TabIndex        =   73
      Text            =   " "
      Top             =   1680
      Width           =   735
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      DataField       =   "ch3"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   2
      Left            =   9360
      TabIndex        =   72
      Text            =   " "
      Top             =   1320
      Width           =   735
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      DataField       =   "ch2"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   1
      Left            =   9360
      TabIndex        =   71
      Text            =   " "
      Top             =   960
      Width           =   735
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      DataField       =   "ch1"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   0
      Left            =   9360
      TabIndex        =   70
      Text            =   " "
      Top             =   600
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      DataField       =   "cd15"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   14
      Left            =   8520
      TabIndex        =   69
      Text            =   " "
      Top             =   5640
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      DataField       =   "cd14"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   13
      Left            =   8520
      TabIndex        =   68
      Text            =   " "
      Top             =   5280
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      DataField       =   "cd13"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   12
      Left            =   8520
      TabIndex        =   67
      Text            =   " "
      Top             =   4920
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      DataField       =   "cd12"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   11
      Left            =   8520
      TabIndex        =   66
      Text            =   " "
      Top             =   4560
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      DataField       =   "cd11"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   10
      Left            =   8520
      TabIndex        =   65
      Text            =   " "
      Top             =   4200
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      DataField       =   "cd10"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   9
      Left            =   8520
      TabIndex        =   64
      Text            =   " "
      Top             =   3840
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      DataField       =   "cd9"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   8
      Left            =   8520
      TabIndex        =   63
      Text            =   " "
      Top             =   3480
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      DataField       =   "cd8"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   7
      Left            =   8520
      TabIndex        =   62
      Text            =   " "
      Top             =   3120
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      DataField       =   "cd7"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   6
      Left            =   8520
      TabIndex        =   61
      Text            =   " "
      Top             =   2760
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      DataField       =   "cd6"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   5
      Left            =   8520
      TabIndex        =   60
      Text            =   " "
      Top             =   2400
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      DataField       =   "cd5"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   4
      Left            =   8520
      TabIndex        =   59
      Text            =   " "
      Top             =   2040
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      DataField       =   "cd4"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   3
      Left            =   8520
      TabIndex        =   58
      Text            =   " "
      Top             =   1680
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      DataField       =   "cd1"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   0
      Left            =   8520
      TabIndex        =   57
      Top             =   600
      Width           =   735
   End
   Begin MSDataGridLib.DataGrid DataGrid3 
      Bindings        =   "frmcolumnasventa0.frx":045B
      Height          =   735
      Left            =   6960
      TabIndex        =   56
      Top             =   6480
      Visible         =   0   'False
      Width           =   1095
      _ExtentX        =   1931
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
   Begin MSAdodcLib.Adodc datcuentas 
      Height          =   330
      Left            =   8040
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
      Connect         =   "PROVIDER=MSDASQL;dsn=contable;uid=sa;pwd=;database=contablesql;"
      OLEDBString     =   "PROVIDER=MSDASQL;dsn=contable;uid=sa;pwd=;database=contablesql;"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select [Cod Contable],[Nombre Cuenta],imp,[Id Cuenta],idcuenta,empre from cuentas ORDER BY IDCUENTA"
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
   Begin VB.TextBox textcuenta 
      Alignment       =   2  'Center
      DataField       =   "Cod Contable"
      BeginProperty DataFormat 
         Type            =   0
         Format          =   """$"" #,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2058
         SubFormatType   =   0
      EndProperty
      DataSource      =   "datcuentas"
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
      Left            =   6840
      TabIndex        =   55
      Top             =   0
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton salir 
      Caption         =   "&Cerrar"
      Height          =   615
      Left            =   4920
      Picture         =   "frmcolumnasventa0.frx":0474
      Style           =   1  'Graphical
      TabIndex        =   50
      Top             =   6480
      UseMaskColor    =   -1  'True
      Width           =   1095
   End
   Begin VB.TextBox txtFields 
      DataField       =   "nomcol15"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   29
      Left            =   2040
      TabIndex        =   47
      Top             =   5655
      Width           =   2895
   End
   Begin VB.TextBox txtFields 
      DataField       =   "ec15"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   30
      Left            =   5040
      TabIndex        =   33
      Top             =   5655
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "ec14"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   28
      Left            =   5040
      TabIndex        =   32
      Top             =   5295
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "nomcol14"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   27
      Left            =   2040
      TabIndex        =   31
      Top             =   5295
      Width           =   2895
   End
   Begin VB.TextBox txtFields 
      DataField       =   "ec13"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   26
      Left            =   5040
      TabIndex        =   30
      Top             =   4935
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "nomcol13"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   25
      Left            =   2040
      TabIndex        =   29
      Top             =   4935
      Width           =   2895
   End
   Begin VB.TextBox txtFields 
      DataField       =   "ec12"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   24
      Left            =   5040
      TabIndex        =   28
      Top             =   4575
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "nomcol12"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   23
      Left            =   2040
      TabIndex        =   27
      Top             =   4575
      Width           =   2895
   End
   Begin VB.TextBox txtFields 
      DataField       =   "ec11"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   22
      Left            =   5040
      TabIndex        =   26
      Top             =   4215
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "nomcol11"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   21
      Left            =   2040
      TabIndex        =   25
      Top             =   4215
      Width           =   2895
   End
   Begin VB.TextBox txtFields 
      DataField       =   "ec10"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   20
      Left            =   5040
      TabIndex        =   24
      Top             =   3855
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "nomcol10"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   19
      Left            =   2040
      TabIndex        =   23
      Top             =   3855
      Width           =   2895
   End
   Begin VB.TextBox txtFields 
      DataField       =   "ec9"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   18
      Left            =   5040
      TabIndex        =   22
      Top             =   3495
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "nomcol9"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   17
      Left            =   2040
      TabIndex        =   21
      Top             =   3495
      Width           =   2895
   End
   Begin VB.TextBox txtFields 
      DataField       =   "ec8"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   16
      Left            =   5040
      TabIndex        =   20
      Top             =   3135
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "nomcol8"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   15
      Left            =   2040
      TabIndex        =   19
      Top             =   3135
      Width           =   2895
   End
   Begin VB.TextBox txtFields 
      DataField       =   "ec7"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   14
      Left            =   5040
      TabIndex        =   18
      Top             =   2775
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "nomcol7"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   13
      Left            =   2040
      TabIndex        =   17
      Top             =   2775
      Width           =   2895
   End
   Begin VB.TextBox txtFields 
      DataField       =   "ec6"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   12
      Left            =   5040
      TabIndex        =   16
      Top             =   2415
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "nomcol6"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   11
      Left            =   2040
      TabIndex        =   15
      Top             =   2415
      Width           =   2895
   End
   Begin VB.TextBox txtFields 
      DataField       =   "ec5"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   10
      Left            =   5040
      TabIndex        =   14
      Top             =   2055
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "nomcol5"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   9
      Left            =   2040
      TabIndex        =   13
      Top             =   2055
      Width           =   2895
   End
   Begin VB.TextBox txtFields 
      DataField       =   "ec4"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   8
      Left            =   5040
      TabIndex        =   12
      Top             =   1695
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "nomcol4"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   7
      Left            =   2040
      TabIndex        =   11
      Top             =   1695
      Width           =   2895
   End
   Begin VB.TextBox txtFields 
      DataField       =   "ec3"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   6
      Left            =   5040
      TabIndex        =   10
      Top             =   1335
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "nomcol3"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   5
      Left            =   2040
      TabIndex        =   9
      Top             =   1335
      Width           =   2895
   End
   Begin VB.TextBox txtFields 
      DataField       =   "ec2"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   4
      Left            =   5040
      TabIndex        =   8
      Top             =   975
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "nomcol2"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   3
      Left            =   2040
      TabIndex        =   7
      Top             =   975
      Width           =   2895
   End
   Begin VB.TextBox txtFields 
      DataField       =   "ec1"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   2
      Left            =   5040
      TabIndex        =   5
      Top             =   615
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "nomcol1"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   1
      Left            =   2040
      TabIndex        =   3
      Top             =   630
      Width           =   2895
   End
   Begin VB.TextBox txtFields 
      DataField       =   "empresa"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   0
      Left            =   1320
      TabIndex        =   1
      Top             =   60
      Visible         =   0   'False
      Width           =   975
   End
   Begin MSAdodcLib.Adodc datPrimaryRS 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      Top             =   7260
      Visible         =   0   'False
      Width           =   10365
      _ExtentX        =   18283
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
      Connect         =   "PROVIDER=MSDASQL;dsn=contable;uid=sa;pwd=;database=contablesql;"
      OLEDBString     =   "PROVIDER=MSDASQL;dsn=contable;uid=sa;pwd=;database=contablesql;"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   $"frmcolumnasventa0.frx":08B6
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
   Begin VB.CommandButton grabar 
      Caption         =   "&Grabar"
      Height          =   615
      Left            =   3480
      Picture         =   "frmcolumnasventa0.frx":08E4
      Style           =   1  'Graphical
      TabIndex        =   49
      Top             =   6480
      UseMaskColor    =   -1  'True
      Width           =   1095
   End
   Begin VB.Frame Frame1 
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
      Height          =   1215
      Left            =   2760
      TabIndex        =   51
      Top             =   6120
      Width           =   3975
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      DataField       =   "cd3"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   2
      Left            =   8520
      TabIndex        =   87
      Text            =   " "
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label lblLabels 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Cod.Imp"
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
      Index           =   21
      Left            =   9360
      TabIndex        =   89
      Top             =   120
      Width           =   735
   End
   Begin VB.Label lblLabels 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
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
      Height          =   195
      Index           =   20
      Left            =   9450
      TabIndex        =   88
      Top             =   360
      Width           =   555
   End
   Begin VB.Label lblLabels 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
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
      Height          =   195
      Index           =   19
      Left            =   8640
      TabIndex        =   86
      Top             =   360
      Width           =   495
   End
   Begin VB.Label lblLabels 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Cod.Imp"
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
      Index           =   18
      Left            =   8520
      TabIndex        =   52
      Top             =   120
      Width           =   735
   End
   Begin VB.Label lblLabels 
      Caption         =   "Columna 15 (C15):"
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
      Index           =   17
      Left            =   120
      TabIndex        =   48
      Top             =   5685
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Columna 14 (C14):"
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
      Index           =   16
      Left            =   120
      TabIndex        =   46
      Top             =   5325
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Columna 13 (C13):"
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
      Left            =   120
      TabIndex        =   45
      Top             =   4965
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Columna 12 (C12):"
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
      Left            =   120
      TabIndex        =   44
      Top             =   4605
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Columna 11 (C11):"
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
      Left            =   120
      TabIndex        =   43
      Top             =   4245
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Columna 10 (C10):"
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
      Left            =   120
      TabIndex        =   42
      Top             =   3885
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Columna 9 (C9):"
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
      Left            =   120
      TabIndex        =   41
      Top             =   3525
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Columna 8 (C8):"
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
      Left            =   120
      TabIndex        =   40
      Top             =   3165
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Columna 7 (C7):"
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
      Index           =   9
      Left            =   120
      TabIndex        =   39
      Top             =   2805
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Columna 6 (C6):"
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
      Index           =   8
      Left            =   120
      TabIndex        =   38
      Top             =   2445
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Columna 5 (C5):"
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
      Left            =   120
      TabIndex        =   37
      Top             =   2085
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Columna 4 (C4):"
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
      Left            =   120
      TabIndex        =   36
      Top             =   1725
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Columna 3 (C3):"
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
      Left            =   120
      TabIndex        =   35
      Top             =   1365
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Columna 2 (C2):"
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
      Left            =   120
      TabIndex        =   34
      Top             =   1005
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Columna 1 (C1):"
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
      Index           =   3
      Left            =   120
      TabIndex        =   6
      Top             =   660
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Ecuación"
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
      Index           =   2
      Left            =   5640
      TabIndex        =   4
      Top             =   360
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Descripcion"
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
      Index           =   1
      Left            =   2760
      TabIndex        =   2
      Top             =   375
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "empresa:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Visible         =   0   'False
      Width           =   1095
   End
End
Attribute VB_Name = "frmcolumnasventa0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim posi As Integer
Dim codi As Integer

Private Sub DataList2_Click()
    
    DataGrid3.Bookmark = DataList2.SelectedItem
    Text5.Text = DataGrid3.Columns(1).Text
    
End Sub

Private Sub DataList2_KeyPress(KeyAscii As Integer)

    If KeyAscii = 27 Then
        If codi = 1 Then Text1(posi).Text = ""
        If codi = 2 Then Text2(posi).Text = ""
        If posi = 14 And codi = 1 Then
            Text2(0).SetFocus
            Exit Sub
        End If
        If posi = 14 And codi = 2 Then
            grabar.SetFocus
            Exit Sub
        End If
        If codi = 1 Then Text1(posi + 1).SetFocus
        If codi = 2 Then Text2(posi + 1).SetFocus
    End If
    If KeyAscii = 13 Then
        KeyAscii = 0
         textcuenta.Text = DataList2.Text
         If codi = 1 Then Text1(posi).Text = textcuenta.Text
         If codi = 2 Then Text2(posi).Text = textcuenta.Text
         If posi = 14 And codi = 1 Then
            Text2(0).SetFocus
            Exit Sub
         End If
         If posi = 14 And codi = 2 Then
            grabar.SetFocus
            Exit Sub
         End If
         If codi = 1 Then Text1(posi + 1).SetFocus
         If codi = 2 Then Text2(posi + 1).SetFocus
    End If
End Sub

Private Sub DataList2_LostFocus()

    textcuenta.Visible = False
    DataList2.Visible = False
    Text5.Visible = False

End Sub

Private Sub Form_Load()

    datPrimaryRS.RecordSource = "SELECT columnasventa.* FROM columnasventa where empresa = " & login.empresaact & ""
    datPrimaryRS.Refresh
    
    datcuentas.RecordSource = "select [Cod Contable],[Nombre Cuenta],imp,[Id Cuenta],idcuenta,empre from cuentas WHERE empre = " & login.empresaact & " and imp = 'S'  ORDER BY IDCUENTA"
    datcuentas.Refresh


    If datPrimaryRS.Recordset.EOF = True Then datPrimaryRS.Recordset.AddNew


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



Private Sub Grabar_Click()

  txtFields(0).Text = login.empresaact
  datPrimaryRS.Recordset.UpdateBatch adAffectCurrent
  datPrimaryRS.Refresh

End Sub

Private Sub salir_Click()

    Unload Me

End Sub

Private Sub Text1_GotFocus(Index As Integer)
    
    posi = Index
    textcuenta.Visible = True
    textcuenta.Left = Text1(Index).Left
    textcuenta.Top = Text1(Index).Top
    DataList2.Visible = True
    DataList2.Left = Text1(Index).Left - DataList2.Width
    DataList2.Top = Text1(Index).Top
    Text5.Visible = True
    Text5.Left = DataList2.Left - Text5.Width
    Text5.Top = Text1(Index).Top
    codi = 1
    DataList2.SetFocus
     
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)

    If KeyAscii = 13 Then
          KeyAscii = 0
          If Index < 14 Then Text1(Index + 1).SetFocus
          If Index = 14 Then Text2(0).SetFocus
    End If

End Sub

Private Sub Text2_GotFocus(Index As Integer)
    posi = Index
    textcuenta.Visible = True
    textcuenta.Left = Text2(Index).Left
    textcuenta.Top = Text2(Index).Top
    DataList2.Visible = True
    DataList2.Left = Text2(Index).Left - DataList2.Width
    DataList2.Top = Text2(Index).Top
    Text5.Visible = True
    Text5.Left = DataList2.Left - Text5.Width
    Text5.Top = Text2(Index).Top
    codi = 2
    DataList2.SetFocus
End Sub

Private Sub Text2_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
          KeyAscii = 0
          If Index < 14 Then Text2(Index + 1).SetFocus
          If Index = 14 Then grabar.SetFocus
    End If

    
End Sub

Private Sub txtFields_KeyPress(Index As Integer, KeyAscii As Integer)

    If KeyAscii = 13 Then
          KeyAscii = 0
          If Index < 30 Then txtFields(Index + 1).SetFocus
          If Index = 30 Then Text1(0).SetFocus
    End If

End Sub
