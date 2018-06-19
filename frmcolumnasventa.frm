VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{5734474E-78D3-4254-99B9-C35F31BDF509}#62.0#0"; "vbskpro2.ocx"
Begin VB.Form frmcolumnasventa 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configuración Columnas Libro Ventas"
   ClientHeight    =   7501
   ClientLeft      =   1092
   ClientTop       =   325
   ClientWidth     =   10374
   Icon            =   "frmcolumnasventa.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7501
   ScaleWidth      =   10374
   Begin MSDataListLib.DataList DataList2 
      Bindings        =   "frmcolumnasventa.frx":0442
      Height          =   2595
      Left            =   4440
      TabIndex        =   52
      Top             =   2400
      Visible         =   0   'False
      Width           =   4455
      _ExtentX        =   8195
      _ExtentY        =   4697
      _Version        =   393216
      MatchEntry      =   -1  'True
      BackColor       =   12640511
      ForeColor       =   -2147483647
      ListField       =   "codigo"
      BoundColumn     =   "Cod Contable"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12.2264
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox txtFields 
      DataField       =   "empresa"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   0
      Left            =   1080
      TabIndex        =   91
      Text            =   " "
      Top             =   0
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton grabar 
      Caption         =   "&Grabar"
      Height          =   615
      Left            =   3480
      Picture         =   "frmcolumnasventa.frx":045B
      Style           =   1  'Graphical
      TabIndex        =   48
      Top             =   6720
      UseMaskColor    =   -1  'True
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      DataField       =   "fc"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   32
      Left            =   4200
      TabIndex        =   90
      Text            =   " "
      Top             =   6000
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      DataField       =   "cht"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   31
      Left            =   9360
      TabIndex        =   87
      Text            =   " "
      Top             =   6000
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      DataField       =   "cdt"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   30
      Left            =   8520
      TabIndex        =   86
      Text            =   " "
      Top             =   6000
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      DataField       =   "ch15"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   29
      Left            =   9360
      TabIndex        =   85
      Text            =   " "
      Top             =   5640
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      DataField       =   "cd15"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   28
      Left            =   8520
      TabIndex        =   84
      Text            =   " "
      Top             =   5640
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      DataField       =   "ch14"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   27
      Left            =   9360
      TabIndex        =   83
      Text            =   " "
      Top             =   5280
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      DataField       =   "cd14"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   26
      Left            =   8520
      TabIndex        =   82
      Text            =   " "
      Top             =   5280
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      DataField       =   "ch13"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   25
      Left            =   9360
      TabIndex        =   81
      Text            =   " "
      Top             =   4920
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      DataField       =   "cd13"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   24
      Left            =   8520
      TabIndex        =   80
      Text            =   " "
      Top             =   4920
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      DataField       =   "ch12"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   23
      Left            =   9360
      TabIndex        =   79
      Text            =   " "
      Top             =   4560
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      DataField       =   "cd12"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   22
      Left            =   8520
      TabIndex        =   78
      Text            =   " "
      Top             =   4560
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      DataField       =   "ch11"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   21
      Left            =   9360
      TabIndex        =   77
      Text            =   " "
      Top             =   4200
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      DataField       =   "cd11"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   20
      Left            =   8520
      TabIndex        =   76
      Text            =   " "
      Top             =   4200
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      DataField       =   "ch10"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   19
      Left            =   9360
      TabIndex        =   75
      Text            =   " "
      Top             =   3840
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      DataField       =   "cd10"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   18
      Left            =   8520
      TabIndex        =   74
      Text            =   " "
      Top             =   3840
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      DataField       =   "ch9"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   17
      Left            =   9360
      TabIndex        =   73
      Text            =   " "
      Top             =   3480
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      DataField       =   "cd9"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   16
      Left            =   8520
      TabIndex        =   72
      Text            =   " "
      Top             =   3480
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      DataField       =   "ch8"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   15
      Left            =   9360
      TabIndex        =   71
      Text            =   " "
      Top             =   3120
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      DataField       =   "cd8"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   14
      Left            =   8520
      TabIndex        =   70
      Text            =   " "
      Top             =   3120
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      DataField       =   "ch1"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   1
      Left            =   9360
      TabIndex        =   65
      Text            =   " "
      Top             =   600
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      DataField       =   "ch7"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   13
      Left            =   9360
      TabIndex        =   64
      Text            =   " "
      Top             =   2760
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      DataField       =   "cd7"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   12
      Left            =   8520
      TabIndex        =   63
      Text            =   " "
      Top             =   2760
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      DataField       =   "ch6"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   11
      Left            =   9360
      TabIndex        =   62
      Text            =   " "
      Top             =   2400
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      DataField       =   "cd6"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   10
      Left            =   8520
      TabIndex        =   61
      Text            =   " "
      Top             =   2400
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      DataField       =   "ch5"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   9
      Left            =   9360
      TabIndex        =   60
      Text            =   " "
      Top             =   2040
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      DataField       =   "cd5"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   8
      Left            =   8520
      TabIndex        =   59
      Text            =   " "
      Top             =   2040
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      DataField       =   "ch4"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   7
      Left            =   9360
      TabIndex        =   58
      Text            =   " "
      Top             =   1680
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      DataField       =   "cd4"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   6
      Left            =   8520
      TabIndex        =   57
      Text            =   " "
      Top             =   1680
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      DataField       =   "ch3"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   5
      Left            =   9360
      TabIndex        =   56
      Text            =   " "
      Top             =   1320
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      DataField       =   "cd3"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   4
      Left            =   8520
      TabIndex        =   55
      Text            =   " "
      Top             =   1320
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      DataField       =   "ch2"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   3
      Left            =   9360
      TabIndex        =   54
      Text            =   " "
      Top             =   960
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      DataField       =   "cd1"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   0
      Left            =   8520
      TabIndex        =   53
      Top             =   600
      Width           =   735
   End
   Begin MSAdodcLib.Adodc datcuentas 
      Height          =   330
      Left            =   6840
      Top             =   6840
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2252
      _ExtentY        =   623
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
         Size            =   7.47
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton salir 
      Caption         =   "&Cerrar"
      Height          =   615
      Left            =   4920
      Picture         =   "frmcolumnasventa.frx":098D
      Style           =   1  'Graphical
      TabIndex        =   49
      Top             =   6720
      UseMaskColor    =   -1  'True
      Width           =   1095
   End
   Begin VB.TextBox txtFields 
      DataField       =   "nomcol15"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   29
      Left            =   2040
      TabIndex        =   46
      Top             =   5655
      Width           =   2895
   End
   Begin VB.TextBox txtFields 
      DataField       =   "ec15"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   30
      Left            =   5040
      TabIndex        =   32
      Top             =   5655
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "ec14"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   28
      Left            =   5040
      TabIndex        =   31
      Top             =   5295
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "nomcol14"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   27
      Left            =   2040
      TabIndex        =   30
      Top             =   5295
      Width           =   2895
   End
   Begin VB.TextBox txtFields 
      DataField       =   "ec13"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   26
      Left            =   5040
      TabIndex        =   29
      Top             =   4935
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "nomcol13"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   25
      Left            =   2040
      TabIndex        =   28
      Top             =   4935
      Width           =   2895
   End
   Begin VB.TextBox txtFields 
      DataField       =   "ec12"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   24
      Left            =   5040
      TabIndex        =   27
      Top             =   4575
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "nomcol12"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   23
      Left            =   2040
      TabIndex        =   26
      Top             =   4575
      Width           =   2895
   End
   Begin VB.TextBox txtFields 
      DataField       =   "ec11"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   22
      Left            =   5040
      TabIndex        =   25
      Top             =   4215
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "nomcol11"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   21
      Left            =   2040
      TabIndex        =   24
      Top             =   4215
      Width           =   2895
   End
   Begin VB.TextBox txtFields 
      DataField       =   "ec10"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   20
      Left            =   5040
      TabIndex        =   23
      Top             =   3855
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "nomcol10"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   19
      Left            =   2040
      TabIndex        =   22
      Top             =   3855
      Width           =   2895
   End
   Begin VB.TextBox txtFields 
      DataField       =   "ec9"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   18
      Left            =   5040
      TabIndex        =   21
      Top             =   3495
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "nomcol9"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   17
      Left            =   2040
      TabIndex        =   20
      Top             =   3495
      Width           =   2895
   End
   Begin VB.TextBox txtFields 
      DataField       =   "ec8"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   16
      Left            =   5040
      TabIndex        =   19
      Top             =   3135
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "nomcol8"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   15
      Left            =   2040
      TabIndex        =   18
      Top             =   3135
      Width           =   2895
   End
   Begin VB.TextBox txtFields 
      DataField       =   "ec7"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   14
      Left            =   5040
      TabIndex        =   17
      Top             =   2775
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "nomcol7"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   13
      Left            =   2040
      TabIndex        =   16
      Top             =   2775
      Width           =   2895
   End
   Begin VB.TextBox txtFields 
      DataField       =   "ec6"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   12
      Left            =   5040
      TabIndex        =   15
      Top             =   2415
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "nomcol6"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   11
      Left            =   2040
      TabIndex        =   14
      Top             =   2415
      Width           =   2895
   End
   Begin VB.TextBox txtFields 
      DataField       =   "ec5"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   10
      Left            =   5040
      TabIndex        =   13
      Top             =   2055
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "nomcol5"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   9
      Left            =   2040
      TabIndex        =   12
      Top             =   2055
      Width           =   2895
   End
   Begin VB.TextBox txtFields 
      DataField       =   "ec4"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   8
      Left            =   5040
      TabIndex        =   11
      Top             =   1695
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "nomcol4"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   7
      Left            =   2040
      TabIndex        =   10
      Top             =   1695
      Width           =   2895
   End
   Begin VB.TextBox txtFields 
      DataField       =   "ec3"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   6
      Left            =   5040
      TabIndex        =   9
      Top             =   1335
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "nomcol3"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   5
      Left            =   2040
      TabIndex        =   8
      Top             =   1335
      Width           =   2895
   End
   Begin VB.TextBox txtFields 
      DataField       =   "ec2"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   4
      Left            =   5040
      TabIndex        =   7
      Top             =   975
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "nomcol2"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   3
      Left            =   2040
      TabIndex        =   6
      Top             =   975
      Width           =   2895
   End
   Begin VB.TextBox txtFields 
      DataField       =   "ec1"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   2
      Left            =   5040
      TabIndex        =   4
      Top             =   615
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "nomcol1"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   1
      Left            =   2040
      TabIndex        =   2
      Text            =   " "
      Top             =   630
      Width           =   2895
   End
   Begin MSAdodcLib.Adodc datPrimaryRS 
      Align           =   2  'Align Bottom
      Height          =   325
      Left            =   0
      Top             =   7176
      Visible         =   0   'False
      Width           =   10374
      _ExtentX        =   18307
      _ExtentY        =   575
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
      Caption         =   " "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.47
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Caption         =   "Opciones"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.47
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   975
      Left            =   2760
      TabIndex        =   50
      Top             =   6480
      Width           =   3975
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      DataField       =   "cd2"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   2
      Left            =   8520
      TabIndex        =   67
      Text            =   " "
      Top             =   960
      Width           =   735
   End
   Begin vbskpro.Skinner Skinner1 
      Left            =   0
      Top             =   0
      _ExtentX        =   1150
      _ExtentY        =   1150
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
      LcK2            =   $"frmcolumnasventa.frx":0DCF
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
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Nº de Cuenta para facturacion de contado:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.47
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   23
      Left            =   120
      TabIndex        =   89
      Top             =   6000
      Width           =   3855
   End
   Begin VB.Label lblLabels 
      Caption         =   "Total:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.47
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   22
      Left            =   7800
      TabIndex        =   88
      Top             =   6000
      Width           =   735
   End
   Begin VB.Label lblLabels 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cod.Imp"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.47
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   21
      Left            =   9360
      TabIndex        =   69
      Top             =   120
      Width           =   735
   End
   Begin VB.Label lblLabels 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Haber"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.47
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   20
      Left            =   9450
      TabIndex        =   68
      Top             =   360
      Width           =   555
   End
   Begin VB.Label lblLabels 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Debe"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.47
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   19
      Left            =   8520
      TabIndex        =   66
      Top             =   360
      Width           =   495
   End
   Begin VB.Label lblLabels 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cod.Imp"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.47
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   18
      Left            =   8400
      TabIndex        =   51
      Top             =   120
      Width           =   735
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Columna 15 (C15):"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.47
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   17
      Left            =   120
      TabIndex        =   47
      Top             =   5685
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Columna 14 (C14):"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.47
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   16
      Left            =   120
      TabIndex        =   45
      Top             =   5325
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Columna 13 (C13):"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.47
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   15
      Left            =   120
      TabIndex        =   44
      Top             =   4965
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Columna 12 (C12):"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.47
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   14
      Left            =   120
      TabIndex        =   43
      Top             =   4605
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Columna 11 (C11):"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.47
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   13
      Left            =   120
      TabIndex        =   42
      Top             =   4245
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Columna 10 (C10):"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.47
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   12
      Left            =   120
      TabIndex        =   41
      Top             =   3885
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Columna 9 (C9):"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.47
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   11
      Left            =   120
      TabIndex        =   40
      Top             =   3525
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Columna 8 (C8):"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.47
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   10
      Left            =   120
      TabIndex        =   39
      Top             =   3165
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Columna 7 (C7):"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.47
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   9
      Left            =   120
      TabIndex        =   38
      Top             =   2805
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Columna 6 (C6):"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.47
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   8
      Left            =   120
      TabIndex        =   37
      Top             =   2445
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Columna 5 (C5):"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.47
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   36
      Top             =   2085
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Columna 4 (C4):"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.47
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   35
      Top             =   1725
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Columna 3 (C3):"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.47
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   34
      Top             =   1365
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Columna 2 (C2):"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.47
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   33
      Top             =   1005
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Columna 1 (C1):"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.47
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   5
      Top             =   660
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Ecuación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.47
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   5640
      TabIndex        =   3
      Top             =   360
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Descripcion"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.47
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   2760
      TabIndex        =   1
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
Attribute VB_Name = "frmcolumnasventa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim posi As Integer
Dim codi As Integer

Private Sub DataCombo1_Click(Area As Integer)


End Sub

Private Sub DataList2_Click()
On Error GoTo fuera

    Text1(posi).Text = DataList2.BoundText
    
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

    If KeyAscii = 48 Then
        Text1(posi).Text = 0
        If posi = 32 Then
            grabar.SetFocus
            Exit Sub
        End If
        Text1(posi + 1).SetFocus
        Exit Sub
    End If
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        Text1(posi).Text = DataList2.BoundText
        If posi = 32 Then
            grabar.SetFocus
            Exit Sub
        End If
        Text1(posi + 1).SetFocus
    End If
    
fuera:
End Sub

Private Sub DataList2_LostFocus()

    DataList2.Visible = False

End Sub



Private Sub Form_Load()

    Inicio.Toolbar1.Visible = True
    
datcuentas.ConnectionString = login.conexiontotal
datPrimaryRS.ConnectionString = login.conexiontotal
    
    datPrimaryRS.RecordSource = "SELECT columnasventa.* FROM columnasventa where empresa = " & login.empresaact & " and inicioper = '" & login.iper & "' and finper = '" & login.fper & "'"
    datPrimaryRS.Refresh
    
    datcuentas.RecordSource = "select listacuentas.* from listacuentas WHERE empre = " & login.empresaact & " and inicioper = '" & login.iper & "' and finper = '" & login.fper & "' ORDER BY IDCUENTA"
    datcuentas.Refresh

    If datPrimaryRS.Recordset.EOF = True Then datPrimaryRS.Recordset.AddNew

End Sub

Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault
  Inicio.Toolbar1.Visible = False
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
Private Sub grabar_Click()
On Error GoTo errorgrabar

For x = 0 To 32
 If Text1(x).Text = "" Then Text1(x).Text = 0
Next x

  txtFields(0).Text = login.empresaact
  datPrimaryRS.Recordset.Fields(63) = login.iper
  datPrimaryRS.Recordset.Fields(64) = login.fper
  datPrimaryRS.Recordset.UpdateBatch adAffectCurrent
  datPrimaryRS.Refresh
Exit Sub

errorgrabar:
 mensa = MsgBox("No se pudo grabar los cambios, intentenlo de nuevo", vbCritical, "!! Error !!")

End Sub

Private Sub salir_Click()

    Unload Me

End Sub



Private Sub Text1_GotFocus(Index As Integer)
On Error GoTo fuera

    posi = Index
    Text1(Index).SelLength = Len(Text1(Index))
    DataList2.BoundText = Text1(Index).Text
    DataList2.Visible = True
    DataList2.Left = Text1(Index).Left + Text1(Index).Width - DataList2.Width
    DataList2.Top = Text1(Index).Top + Text1(Index).Height
    DataList2.SetFocus
    
fuera:
End Sub



Private Sub txtFields_KeyPress(Index As Integer, KeyAscii As Integer)
On Error GoTo fuera

    If KeyAscii = 13 Then
          KeyAscii = 0
          If Index < 30 Then txtFields(Index + 1).SetFocus
          If Index = 30 Then Text1(0).SetFocus
    End If

fuera:
End Sub
