VERSION 5.00
Object = "{F6125AB1-8AB1-11CE-A77F-08002B2F4E98}#2.0#0"; "MSRDC20.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{5734474E-78D3-4254-99B9-C35F31BDF509}#62.0#0"; "vbskpro2.ocx"
Begin VB.Form frmresultados 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Resultados"
   ClientHeight    =   5445
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9615
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5445
   ScaleWidth      =   9615
   Begin TabDlg.SSTab SSTab1 
      Height          =   5175
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   9128
      _Version        =   393216
      Tabs            =   5
      Tab             =   4
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "Periodos a Calcular"
      TabPicture(0)   =   "frmresultados.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(1)=   "calen"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Ventas Resultados"
      TabPicture(1)   =   "frmresultados.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Aceptar"
      Tab(1).Control(1)=   "ventas"
      Tab(1).Control(2)=   "Text1(10)"
      Tab(1).Control(3)=   "Text1(9)"
      Tab(1).Control(4)=   "Text1(8)"
      Tab(1).Control(5)=   "Text1(7)"
      Tab(1).Control(6)=   "Text1(6)"
      Tab(1).Control(7)=   "Text1(5)"
      Tab(1).Control(8)=   "Text1(4)"
      Tab(1).Control(9)=   "Text1(3)"
      Tab(1).Control(10)=   "Text1(2)"
      Tab(1).Control(11)=   "Text1(1)"
      Tab(1).Control(12)=   "Text1(0)"
      Tab(1).ControlCount=   13
      TabCaption(2)   =   "Resultados Grales."
      TabPicture(2)   =   "frmresultados.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Check4"
      Tab(2).Control(1)=   "Check2"
      Tab(2).Control(2)=   "Check1"
      Tab(2).Control(3)=   "Command5"
      Tab(2).Control(4)=   "Command1"
      Tab(2).Control(5)=   "bar1"
      Tab(2).Control(6)=   "Text2(9)"
      Tab(2).Control(7)=   "Text2(8)"
      Tab(2).Control(8)=   "Text2(7)"
      Tab(2).Control(9)=   "Text2(6)"
      Tab(2).Control(10)=   "Text2(5)"
      Tab(2).Control(11)=   "Text2(4)"
      Tab(2).Control(12)=   "Text2(3)"
      Tab(2).Control(13)=   "Text2(2)"
      Tab(2).Control(14)=   "Text2(1)"
      Tab(2).Control(15)=   "Text2(0)"
      Tab(2).Control(16)=   "fondoscalcula"
      Tab(2).Control(17)=   "result"
      Tab(2).Control(18)=   "bar2"
      Tab(2).ControlCount=   19
      TabCaption(3)   =   "Gestión Comercial"
      TabPicture(3)   =   "frmresultados.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Command6"
      Tab(3).Control(1)=   "Command4"
      Tab(3).Control(2)=   "Command3"
      Tab(3).Control(3)=   "Text4"
      Tab(3).Control(4)=   "Command2(11)"
      Tab(3).Control(5)=   "Command2(10)"
      Tab(3).Control(6)=   "Command2(9)"
      Tab(3).Control(7)=   "Command2(8)"
      Tab(3).Control(8)=   "Command2(7)"
      Tab(3).Control(9)=   "Command2(6)"
      Tab(3).Control(10)=   "Command2(5)"
      Tab(3).Control(11)=   "Command2(4)"
      Tab(3).Control(12)=   "Command2(3)"
      Tab(3).Control(13)=   "Command2(2)"
      Tab(3).Control(14)=   "Command2(1)"
      Tab(3).Control(15)=   "Command2(0)"
      Tab(3).Control(16)=   "Text3(10)"
      Tab(3).Control(17)=   "Text3(9)"
      Tab(3).Control(18)=   "Text3(8)"
      Tab(3).Control(19)=   "Text3(7)"
      Tab(3).Control(20)=   "Text3(6)"
      Tab(3).Control(21)=   "Text3(5)"
      Tab(3).Control(22)=   "Text3(4)"
      Tab(3).Control(23)=   "Text3(3)"
      Tab(3).Control(24)=   "Text3(2)"
      Tab(3).Control(25)=   "Text3(1)"
      Tab(3).Control(26)=   "Text3(0)"
      Tab(3).Control(27)=   "gestion"
      Tab(3).ControlCount=   28
      TabCaption(4)   =   "Anáilis por CC"
      TabPicture(4)   =   "frmresultados.frx":0070
      Tab(4).ControlEnabled=   -1  'True
      Tab(4).Control(0)=   "Frame3"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "Frame2"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).Control(2)=   "Command7"
      Tab(4).Control(2).Enabled=   0   'False
      Tab(4).Control(3)=   "DataList2"
      Tab(4).Control(3).Enabled=   0   'False
      Tab(4).ControlCount=   4
      Begin VB.CheckBox Check4 
         Caption         =   "Balance"
         Height          =   255
         Left            =   -69360
         TabIndex        =   108
         Top             =   4020
         Width           =   1095
      End
      Begin MSDataListLib.DataList DataList2 
         Bindings        =   "frmresultados.frx":008C
         Height          =   1560
         Left            =   3480
         TabIndex        =   98
         Top             =   1200
         Visible         =   0   'False
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   2752
         _Version        =   393216
         MatchEntry      =   -1  'True
         BackColor       =   12648447
         ForeColor       =   -2147483647
         ListField       =   "codigo"
         BoundColumn     =   "idcuenta"
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
      Begin VB.CommandButton Command7 
         Caption         =   "Generar Reporte"
         Height          =   975
         Left            =   4080
         Picture         =   "frmresultados.frx":00A5
         Style           =   1  'Graphical
         TabIndex        =   105
         Top             =   3120
         Width           =   1575
      End
      Begin VB.Frame Frame2 
         Caption         =   "Codigo Contable"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Left            =   360
         TabIndex        =   93
         Top             =   720
         Width           =   3255
         Begin VB.TextBox Text6 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   1440
            TabIndex        =   97
            Top             =   960
            Width           =   1695
         End
         Begin VB.TextBox Text5 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   1440
            TabIndex        =   95
            Top             =   480
            Width           =   1695
         End
         Begin VB.Label Label4 
            Caption         =   "Hasta Cuenta:"
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
            Left            =   120
            TabIndex        =   96
            Top             =   960
            Width           =   1455
         End
         Begin VB.Label Label3 
            Caption         =   "Desde Cuenta:"
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
            Left            =   120
            TabIndex        =   94
            Top             =   480
            Width           =   1455
         End
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Command6"
         Height          =   255
         Left            =   -72720
         TabIndex        =   92
         Top             =   540
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Det.Cuentas"
         Height          =   255
         Left            =   -70800
         TabIndex        =   91
         Top             =   4020
         Width           =   1215
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Resultados"
         Height          =   255
         Left            =   -72120
         TabIndex        =   90
         Top             =   4020
         Width           =   1215
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Command5"
         Height          =   255
         Left            =   -74280
         TabIndex        =   89
         Top             =   360
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Command4"
         Height          =   255
         Left            =   -68160
         TabIndex        =   88
         Top             =   420
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Generar Reporte"
         Height          =   615
         Left            =   -74520
         Picture         =   "frmresultados.frx":04E7
         Style           =   1  'Graphical
         TabIndex        =   87
         Top             =   420
         Width           =   1575
      End
      Begin VB.TextBox Text4 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   -71280
         TabIndex        =   86
         Top             =   540
         Width           =   2655
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Command2"
         Height          =   375
         Index           =   11
         Left            =   -67800
         TabIndex        =   84
         Top             =   4620
         Width           =   1335
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Command2"
         Height          =   375
         Index           =   10
         Left            =   -69120
         TabIndex        =   83
         Top             =   4620
         Width           =   1335
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Command2"
         Height          =   375
         Index           =   9
         Left            =   -70440
         TabIndex        =   82
         Top             =   4620
         Width           =   1335
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Command2"
         Height          =   375
         Index           =   8
         Left            =   -71760
         TabIndex        =   81
         Top             =   4620
         Width           =   1335
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Command2"
         Height          =   375
         Index           =   7
         Left            =   -73080
         TabIndex        =   80
         Top             =   4620
         Width           =   1335
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Command2"
         Height          =   375
         Index           =   6
         Left            =   -74400
         TabIndex        =   79
         Top             =   4620
         Width           =   1335
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Command2"
         Height          =   375
         Index           =   5
         Left            =   -67800
         TabIndex        =   78
         Top             =   4260
         Width           =   1335
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Command2"
         Height          =   375
         Index           =   4
         Left            =   -69120
         TabIndex        =   77
         Top             =   4260
         Width           =   1335
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Command2"
         Height          =   375
         Index           =   3
         Left            =   -70440
         TabIndex        =   76
         Top             =   4260
         Width           =   1335
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Command2"
         Height          =   375
         Index           =   2
         Left            =   -71760
         TabIndex        =   75
         Top             =   4260
         Width           =   1335
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Command2"
         Height          =   375
         Index           =   1
         Left            =   -73080
         TabIndex        =   74
         Top             =   4260
         Width           =   1335
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Command2"
         Height          =   375
         Index           =   0
         Left            =   -74400
         TabIndex        =   73
         Top             =   4260
         Width           =   1335
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   10
         Left            =   -74520
         TabIndex        =   72
         Top             =   3540
         Width           =   2175
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   9
         Left            =   -74520
         TabIndex        =   71
         Top             =   3300
         Width           =   2175
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   8
         Left            =   -74520
         TabIndex        =   70
         Top             =   3060
         Width           =   2175
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   7
         Left            =   -74520
         TabIndex        =   69
         Top             =   2820
         Width           =   2175
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   6
         Left            =   -74520
         TabIndex        =   68
         Top             =   2580
         Width           =   2175
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   5
         Left            =   -74520
         TabIndex        =   67
         Top             =   2340
         Width           =   2175
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   4
         Left            =   -74520
         TabIndex        =   66
         Top             =   2100
         Width           =   2175
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   3
         Left            =   -74520
         TabIndex        =   65
         Top             =   1860
         Width           =   2175
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   2
         Left            =   -74520
         TabIndex        =   64
         Top             =   1620
         Width           =   2175
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   1
         Left            =   -74520
         TabIndex        =   63
         Top             =   1380
         Width           =   2175
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   0
         Left            =   -74520
         TabIndex        =   62
         Top             =   1140
         Width           =   2175
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Generar Reporte"
         Height          =   735
         Left            =   -73560
         Picture         =   "frmresultados.frx":0929
         Style           =   1  'Graphical
         TabIndex        =   61
         Top             =   3480
         Width           =   1335
      End
      Begin MSComctlLib.ProgressBar bar1 
         Height          =   255
         Left            =   -74640
         TabIndex        =   59
         Top             =   4320
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Index           =   9
         Left            =   -74640
         TabIndex        =   56
         Top             =   2880
         Width           =   2415
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Index           =   8
         Left            =   -74640
         TabIndex        =   55
         Top             =   2640
         Width           =   2415
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Index           =   7
         Left            =   -74640
         TabIndex        =   54
         Top             =   2400
         Width           =   2415
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Index           =   6
         Left            =   -74640
         TabIndex        =   53
         Top             =   2160
         Width           =   2415
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Index           =   5
         Left            =   -74640
         TabIndex        =   52
         Top             =   1920
         Width           =   2415
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Index           =   4
         Left            =   -74640
         TabIndex        =   51
         Top             =   1680
         Width           =   2415
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Index           =   3
         Left            =   -74640
         TabIndex        =   50
         Top             =   1440
         Width           =   2415
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Index           =   2
         Left            =   -74640
         TabIndex        =   49
         Top             =   1200
         Width           =   2415
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Index           =   1
         Left            =   -74640
         TabIndex        =   48
         Top             =   960
         Width           =   2415
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Index           =   0
         Left            =   -74640
         TabIndex        =   47
         Top             =   720
         Width           =   2415
      End
      Begin VB.CommandButton fondoscalcula 
         Caption         =   "Calcular"
         Height          =   735
         Left            =   -74640
         Picture         =   "frmresultados.frx":0D6B
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   3480
         UseMaskColor    =   -1  'True
         Width           =   975
      End
      Begin VB.CommandButton Aceptar 
         Caption         =   "&Aceptar"
         Height          =   855
         Left            =   -74640
         Picture         =   "frmresultados.frx":1075
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   4020
         UseMaskColor    =   -1  'True
         Width           =   975
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid ventas 
         Height          =   3495
         Left            =   -72360
         TabIndex        =   31
         Top             =   720
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   6165
         _Version        =   393216
         Rows            =   13
         Cols            =   12
         FixedCols       =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   12
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   10
         Left            =   -74640
         TabIndex        =   30
         Top             =   3360
         Width           =   2175
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   9
         Left            =   -74640
         TabIndex        =   29
         Top             =   3120
         Width           =   2175
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   8
         Left            =   -74640
         TabIndex        =   28
         Top             =   2880
         Width           =   2175
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   7
         Left            =   -74640
         TabIndex        =   27
         Top             =   2640
         Width           =   2175
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   6
         Left            =   -74640
         TabIndex        =   26
         Top             =   2400
         Width           =   2175
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   5
         Left            =   -74640
         TabIndex        =   25
         Top             =   2160
         Width           =   2175
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   4
         Left            =   -74640
         TabIndex        =   24
         Top             =   1920
         Width           =   2175
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   3
         Left            =   -74640
         TabIndex        =   23
         Top             =   1680
         Width           =   2175
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   2
         Left            =   -74640
         TabIndex        =   22
         Top             =   1440
         Width           =   2175
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   1
         Left            =   -74640
         TabIndex        =   21
         Top             =   1200
         Width           =   2175
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   0
         Left            =   -74640
         TabIndex        =   20
         Top             =   960
         Width           =   2175
      End
      Begin VB.Frame Frame1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   4215
         Left            =   -74640
         TabIndex        =   1
         Top             =   720
         Width           =   5175
         Begin VB.CommandButton seleccionames 
            Caption         =   "seleccionames"
            Height          =   255
            Left            =   3120
            TabIndex        =   107
            Top             =   3480
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Option1"
            Height          =   255
            Index           =   12
            Left            =   480
            TabIndex        =   44
            Top             =   3480
            Width           =   255
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Option1"
            Height          =   255
            Index           =   10
            Left            =   480
            TabIndex        =   42
            Top             =   3000
            Width           =   255
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Option1"
            Height          =   255
            Index           =   11
            Left            =   480
            TabIndex        =   43
            Top             =   3240
            Width           =   255
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Option1"
            Height          =   255
            Index           =   8
            Left            =   480
            TabIndex        =   40
            Top             =   2520
            Width           =   255
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Option1"
            Height          =   255
            Index           =   9
            Left            =   480
            TabIndex        =   41
            Top             =   2760
            Width           =   255
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Option1"
            Height          =   255
            Index           =   6
            Left            =   480
            TabIndex        =   38
            Top             =   2040
            Width           =   255
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Option1"
            Height          =   255
            Index           =   7
            Left            =   480
            TabIndex        =   39
            Top             =   2280
            Width           =   255
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Option1"
            Height          =   255
            Index           =   4
            Left            =   480
            TabIndex        =   36
            Top             =   1560
            Width           =   255
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Option1"
            Height          =   255
            Index           =   5
            Left            =   480
            TabIndex        =   37
            Top             =   1800
            Width           =   255
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Option1"
            Height          =   255
            Index           =   2
            Left            =   480
            TabIndex        =   34
            Top             =   1080
            Width           =   255
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Option1"
            Height          =   255
            Index           =   3
            Left            =   480
            TabIndex        =   35
            Top             =   1320
            Width           =   255
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Option1"
            Height          =   255
            Index           =   1
            Left            =   480
            TabIndex        =   33
            Top             =   840
            Width           =   255
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Option1"
            Height          =   255
            Index           =   0
            Left            =   480
            TabIndex        =   32
            Top             =   600
            Width           =   255
         End
         Begin MSComCtl2.DTPicker desde 
            Height          =   375
            Left            =   3480
            TabIndex        =   2
            Top             =   480
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            _Version        =   393216
            Format          =   158400513
            CurrentDate     =   39007
         End
         Begin MSComCtl2.DTPicker hasta 
            Height          =   375
            Left            =   3480
            TabIndex        =   3
            Top             =   1080
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            _Version        =   393216
            Format          =   158400513
            CurrentDate     =   39007
         End
         Begin VB.Label Label1 
            Caption         =   "Enero"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   840
            TabIndex        =   19
            Top             =   600
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Febrero"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   840
            TabIndex        =   18
            Top             =   840
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Marzo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   840
            TabIndex        =   17
            Top             =   1080
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Abril"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   3
            Left            =   840
            TabIndex        =   16
            Top             =   1320
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Mayo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   4
            Left            =   840
            TabIndex        =   15
            Top             =   1560
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Junio"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   5
            Left            =   840
            TabIndex        =   14
            Top             =   1800
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Julio"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   6
            Left            =   840
            TabIndex        =   13
            Top             =   2040
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Agosto"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   7
            Left            =   840
            TabIndex        =   12
            Top             =   2280
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Septiembre"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   8
            Left            =   840
            TabIndex        =   11
            Top             =   2520
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   "Octubre"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   9
            Left            =   840
            TabIndex        =   10
            Top             =   2760
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Noviembre"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   10
            Left            =   840
            TabIndex        =   9
            Top             =   3000
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Diciembre"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   11
            Left            =   840
            TabIndex        =   8
            Top             =   3240
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Mes en Curso no Cerrado"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   12
            Left            =   840
            TabIndex        =   7
            Top             =   3480
            Width           =   1455
         End
         Begin VB.Label Label2 
            Caption         =   "ACUMULAR HASTA"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   6
            Top             =   360
            Width           =   1815
         End
         Begin VB.Label Label1 
            Caption         =   "Desde:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   13
            Left            =   2760
            TabIndex        =   5
            Top             =   600
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Hasta:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   14
            Left            =   2760
            TabIndex        =   4
            Top             =   1200
            Width           =   735
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid result 
         Height          =   3495
         Left            =   -72120
         TabIndex        =   57
         Top             =   480
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   6165
         _Version        =   393216
         Rows            =   13
         Cols            =   12
         FixedCols       =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   12
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin MSComCtl2.MonthView calen 
         Height          =   2370
         Left            =   -69000
         TabIndex        =   58
         Top             =   840
         Width           =   2595
         _ExtentX        =   4577
         _ExtentY        =   4180
         _Version        =   393216
         ForeColor       =   -2147483630
         BackColor       =   -2147483633
         Appearance      =   1
         StartOfWeek     =   157745154
         CurrentDate     =   39013
      End
      Begin MSComctlLib.ProgressBar bar2 
         Height          =   255
         Left            =   -74640
         TabIndex        =   60
         Top             =   4680
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid gestion 
         Height          =   3255
         Left            =   -72240
         TabIndex        =   85
         Top             =   900
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   5741
         _Version        =   393216
         Rows            =   13
         Cols            =   6
         FixedCols       =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   6
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin VB.Frame Frame3 
         Caption         =   "Niveles de Centros de Costo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Left            =   360
         TabIndex        =   99
         Top             =   2520
         Width           =   3255
         Begin VB.CheckBox Check3 
            Caption         =   "Ninguno"
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
            Left            =   240
            TabIndex        =   106
            Top             =   1080
            Width           =   1335
         End
         Begin VB.OptionButton Option2 
            Caption         =   "1er"
            Height          =   255
            Index           =   0
            Left            =   1800
            TabIndex        =   103
            Top             =   360
            Width           =   615
         End
         Begin VB.OptionButton Option2 
            Caption         =   "2do"
            Height          =   255
            Index           =   1
            Left            =   1800
            TabIndex        =   102
            Top             =   600
            Width           =   615
         End
         Begin VB.OptionButton Option2 
            Caption         =   "3er"
            Height          =   255
            Index           =   2
            Left            =   1800
            TabIndex        =   101
            Top             =   840
            Width           =   615
         End
         Begin VB.OptionButton Option2 
            Caption         =   "4to"
            Height          =   255
            Index           =   3
            Left            =   1800
            TabIndex        =   100
            Top             =   1080
            Width           =   615
         End
         Begin VB.Label Label6 
            Caption         =   "Desplegar Hasta:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   600
            TabIndex        =   104
            Top             =   480
            Width           =   975
         End
      End
   End
   Begin MSAdodcLib.Adodc datresulventas 
      Height          =   330
      Left            =   6840
      Top             =   240
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
   Begin vbskpro.Skinner Skinner1 
      Left            =   0
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
      LcK2            =   $"frmresultados.frx":14B7
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
   Begin MSAdodcLib.Adodc dattipoclientes 
      Height          =   330
      Left            =   0
      Top             =   0
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
   Begin MSAdodcLib.Adodc datlibro 
      Height          =   330
      Left            =   1800
      Top             =   0
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
   Begin MSAdodcLib.Adodc datparamresultados 
      Height          =   330
      Left            =   3240
      Top             =   0
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
   Begin MSAdodcLib.Adodc datparamresultados1 
      Height          =   330
      Left            =   4560
      Top             =   0
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
   Begin MSAdodcLib.Adodc datfiltro 
      Height          =   330
      Left            =   5520
      Top             =   0
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
   Begin MSAdodcLib.Adodc datcuadro 
      Height          =   330
      Left            =   840
      Top             =   0
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
   Begin MSAdodcLib.Adodc datcuadro2 
      Height          =   330
      Left            =   7200
      Top             =   0
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
   Begin MSAdodcLib.Adodc datresulventas1 
      Height          =   330
      Left            =   8280
      Top             =   0
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
   Begin MSAdodcLib.Adodc datordenes 
      Height          =   330
      Left            =   120
      Top             =   5520
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
      Connect         =   "Provider=MSDASQL.1;Password=25072004;Persist Security Info=True;User ID=lucva;Data Source=contable;Initial Catalog=ordenesradio"
      OLEDBString     =   "Provider=MSDASQL.1;Password=25072004;Persist Security Info=True;User ID=lucva;Data Source=contable;Initial Catalog=ordenesradio"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select ordenesfac.* from ordenesfac"
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
   Begin MSAdodcLib.Adodc datcuadro3 
      Height          =   330
      Left            =   1320
      Top             =   5520
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
   Begin Crystal.CrystalReport CrystalReporte 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      WindowTitle     =   "Orden de Pago"
   End
   Begin MSRDC.MSRDC reporte 
      Height          =   375
      Left            =   0
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
      Connect         =   ""
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
   Begin MSAdodcLib.Adodc datcuentas 
      Height          =   330
      Left            =   0
      Top             =   0
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
      LockType        =   1
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
End
Attribute VB_Name = "frmresultados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cont As Integer
Dim hastaacum As Date
Dim mesesmuestra As Integer
Dim mesihasta(20) As String
Dim mesidesde(20) As String
Dim grupo(200) As String
Dim maxventas As Integer
Dim maxresultados As Integer
Dim nogravado(20, 20) As Currency
Dim totalgral(20, 20) As Currency
Dim mesmuestra As Integer
Dim reportear As Integer
Dim posi As Integer

Private Sub aceptar_Click()
On Error GoTo fuera
Dim mes(20) As String
Dim sumatoria(20, 20) As Currency
Dim totales(20, 20) As Currency

meshasta = Month(hasta.Value)
If Year(hasta.Value) > Year(desde.Value) Then meshasta = meshasta + 12

For x = 0 To 12
    If Option1(x).Value = True Then
        mesesmuestra = meshasta - x
    End If
Next x
If mesesmuestra < 0 Then
    mensa = MsgBox("Eleccion incorrecta, verifique", vbCritical, "Error")
    Exit Sub
End If
ventas.Cols = mesesmuestra
result.Cols = mesesmuestra

result.Row = 0
ventas.Row = 0
mes(20) = "AGOSTO"
mes(19) = "JULIO"
mes(18) = "JUNIO"
mes(17) = "MAYO"
mes(16) = "ABRIL"
mes(15) = "MARZO"
mes(14) = "FEBRERO"
mes(13) = "ENERO"
mes(12) = "DICIEMBRE"
mes(11) = "NOVIEMBRE"
mes(10) = "OCTUBRE"
mes(9) = "SEPTIEMBRE"
mes(8) = "AGOSTO"
mes(7) = "JULIO"
mes(6) = "JUNIO"
mes(5) = "MAYO"
mes(4) = "ABRIL"
mes(3) = "MARZO"
mes(2) = "FEBRERO"
mes(1) = "ENERO"
    
Y = 0
For x = mesesmuestra - 1 To 0 Step -1
    ventas.Col = x
    result.Col = x
    ventas.Text = mes(meshasta - Y)
    result.Text = ventas.Text
    Y = Y + 1
Next x


For Z = Y To 11
    command2(Z).Visible = False
Next Z

For x = 0 To 11
  If Option1(x).Value = True Then
    ventas.Col = 0
    result.Col = 0
    If Month(desde.Value) <> x + 1 Then ventas.Text = "ACUM " + ventas.Text
    result.Text = ventas.Text
  End If
Next x

For Y = 0 To mesesmuestra - 1
    ventas.Row = 0
    ventas.Col = Y
    command2(Y).Caption = ventas.Text
Next Y

Rem ************ calcular sumas de ventas  ****************

For x = 1 To 20
    For Y = 1 To 20
        sumatoria(x, Y) = 0
        nogravado(x, Y) = 0
        totalgral(x, Y) = 0
        totales(x, Y) = 0
    Next Y
Next x

datresulventas.RecordSource = "select resultadosventas.* from resultadosventas where empresa = " & login.empresaact & " and fecha >= '" & desde.Value & "' and fecha <= '" & hasta.Value & "' order by fecha"
datresulventas.Refresh

If datresulventas.Recordset.EOF = True Then Exit Sub

For x = 0 To 11
  If Option1(x).Value = True Then mesacumul = x + 1
Next x

datresulventas.Recordset.MoveFirst
Do While Not datresulventas.Recordset.EOF
    If IsNull(datresulventas.Recordset.Fields("codanun")) = True Then
        If IsNull(datresulventas.Recordset.Fields("codclien")) = True Then
            codi = 0
        Else
            codi = datresulventas.Recordset.Fields("codclien")
        End If
    Else
        codi = datresulventas.Recordset.Fields("codanun")
    End If
    
    If Left(datresulventas.Recordset.Fields("avisador"), 6) = "COMUNA" Or Left(datresulventas.Recordset.Fields("avisador"), 6) = "Comuna" Or Left(datresulventas.Recordset.Fields("avisador"), 6) = "comuna" Then codi = 4
    If Left(datresulventas.Recordset.Fields("avisador"), 5) = "MUNIC" Or Left(datresulventas.Recordset.Fields("avisador"), 5) = "Munic" Or Left(datresulventas.Recordset.Fields("avisador"), 5) = "munic" Then codi = 4
    
    mesfecha = Month(datresulventas.Recordset.Fields("fecha"))
    añofecha = Year(datresulventas.Recordset.Fields("fecha"))
    If añofecha > Year(desde.Value) Then mesfecha = mesfecha + 12
    
    total = datresulventas.Recordset.Fields("expr1") + datresulventas.Recordset.Fields("expr2")
    totalg = datresulventas.Recordset.Fields("expr16")
    nogravado(codi, mesfecha) = nogravado(codi, mesfecha) + datresulventas.Recordset.Fields("expr2")
    Rem totalgral(codi, mesfecha) = totalg + totalgral(codi, mesfecha)
    
    sumatoria(codi, mesfecha) = total + sumatoria(codi, mesfecha)
    totalgral(codi, mesfecha) = sumatoria(codi, mesfecha)
    datresulventas.Recordset.MoveNext

Loop



For x = 1 To mesacumul - 1
  For Y = 0 To cont
    sumatoria(Y, mesacumul) = sumatoria(Y, mesacumul) + sumatoria(Y, x)
  Next Y
Next x

For x = 1 To mesesmuestra
 For Y = 0 To cont
    ventas.Col = x - 1
    ventas.Row = Y + 1
    ventas.Text = sumatoria(Y, x + (meshasta - mesesmuestra))
    ventas.Text = Format(Val(ventas.Text), "###,##0.00")
 Next Y
Next x

For x = 1 To mesesmuestra
totales(cont + 1, x) = 0
 For Y = 0 To cont - 1
    ventas.Col = x - 1
    ventas.Row = Y + 1
    totales(cont + 1, x) = totales(cont + 1, x) + ventas.Text
    ventas.Row = cont + 1
    ventas.Text = totales(cont + 1, x)
    ventas.Text = Format(Val(ventas.Text), "###,##0.00")
    ventas.CellFontBold = True
 Next Y
Next x
   
Exit Sub

fuera:
    MsgBox "Valor fuera de rango", vbCritical, "Error"


End Sub

Private Sub Check3_Click()

    If Check3.Value = 1 Then
        For x = 0 To 3
            Option2(x).Value = False
        Next x
    Else
        Option2(3).Value = True
    End If

End Sub

Private Sub Command1_Click()

datcuadro2.RecordSource = "select cuadroresultados2.* from cuadroresultados2"
datcuadro2.Refresh

If datcuadro2.Recordset.EOF = True Then GoTo sigue
datcuadro2.Recordset.MoveFirst
Do While Not datcuadro2.Recordset.EOF
    datcuadro2.Recordset.Delete adAffectCurrent
    datcuadro2.Recordset.MoveNext
Loop

sigue:
Rem ******** genera cuadro de ventas **************

For Y = 0 To maxventas
  datcuadro2.Recordset.AddNew
  datcuadro2.Recordset.Fields("desde") = desde.Value
  datcuadro2.Recordset.Fields("hasta") = hasta.Value
  If Y = 0 Then
    datcuadro2.Recordset.Fields(0) = Y
    datcuadro2.Recordset.Fields(1) = ""
    datcuadro2.Recordset.Fields(2) = ""
    GoTo sigue0
  End If
  datcuadro2.Recordset.Fields(0) = Y - 1
  datcuadro2.Recordset.Fields(1) = " Ventas"
  datcuadro2.Recordset.Fields(2) = Text1(Y - 1).Text
  datcuadro2.Recordset.Fields("balance") = "N"
sigue0:
    For x = 0 To ventas.Cols - 1
        ventas.Col = x
        ventas.Row = Y
        If Y = 0 Then GoTo sigue2
        If x = 0 Then datcuadro2.Recordset.Fields("mes1") = ventas.Text
        If x = 1 Then datcuadro2.Recordset.Fields("mes2") = ventas.Text
        If x = 2 Then datcuadro2.Recordset.Fields("mes3") = ventas.Text
        If x = 3 Then datcuadro2.Recordset.Fields("mes4") = ventas.Text
        If x = 4 Then datcuadro2.Recordset.Fields("mes5") = ventas.Text
        If x = 5 Then datcuadro2.Recordset.Fields("mes6") = ventas.Text
        If x = 6 Then datcuadro2.Recordset.Fields("mes7") = ventas.Text
        If x = 7 Then datcuadro2.Recordset.Fields("mes8") = ventas.Text
        If x = 8 Then datcuadro2.Recordset.Fields("mes9") = ventas.Text
        If x = 9 Then datcuadro2.Recordset.Fields("mes10") = ventas.Text
        If x = 10 Then datcuadro2.Recordset.Fields("mes11") = ventas.Text
        If x = 11 Then datcuadro2.Recordset.Fields("mes12") = ventas.Text
        GoTo sigue3
sigue2:
        If x = 0 Then datcuadro2.Recordset.Fields("encab1") = ventas.Text
        If x = 1 Then datcuadro2.Recordset.Fields("encab2") = ventas.Text
        If x = 2 Then datcuadro2.Recordset.Fields("encab3") = ventas.Text
        If x = 3 Then datcuadro2.Recordset.Fields("encab4") = ventas.Text
        If x = 4 Then datcuadro2.Recordset.Fields("encab5") = ventas.Text
        If x = 5 Then datcuadro2.Recordset.Fields("encab6") = ventas.Text
        If x = 6 Then datcuadro2.Recordset.Fields("encab7") = ventas.Text
        If x = 7 Then datcuadro2.Recordset.Fields("encab8") = ventas.Text
        If x = 8 Then datcuadro2.Recordset.Fields("encab9") = ventas.Text
        If x = 9 Then datcuadro2.Recordset.Fields("encab10") = ventas.Text
        If x = 10 Then datcuadro2.Recordset.Fields("encab11") = ventas.Text
        If x = 11 Then datcuadro2.Recordset.Fields("encab12") = ventas.Text
sigue3:
    Next x
   datcuadro2.Recordset.UpdateBatch adAffectCurrent
Next Y

Rem ******** genera cuadro de resultados **************

For Y = 0 To maxresultados
  datcuadro2.Recordset.AddNew
  datcuadro2.Recordset.Fields("desde") = desde.Value
  datcuadro2.Recordset.Fields("hasta") = hasta.Value
  datcuadro2.Recordset.Fields(0) = Y
  datcuadro2.Recordset.Fields(1) = grupo(Y)
  datcuadro2.Recordset.Fields(2) = Text2(Y).Text
  datcuadro2.Recordset.Fields("balance") = "S"
    For x = 0 To result.Cols - 1
        result.Col = x
        result.Row = Y + 1
        If x = 0 Then datcuadro2.Recordset.Fields("mes1") = result.Text
        If x = 1 Then datcuadro2.Recordset.Fields("mes2") = result.Text
        If x = 2 Then datcuadro2.Recordset.Fields("mes3") = result.Text
        If x = 3 Then datcuadro2.Recordset.Fields("mes4") = result.Text
        If x = 4 Then datcuadro2.Recordset.Fields("mes5") = result.Text
        If x = 5 Then datcuadro2.Recordset.Fields("mes6") = result.Text
        If x = 6 Then datcuadro2.Recordset.Fields("mes7") = result.Text
        If x = 7 Then datcuadro2.Recordset.Fields("mes8") = result.Text
        If x = 8 Then datcuadro2.Recordset.Fields("mes9") = result.Text
        If x = 9 Then datcuadro2.Recordset.Fields("mes10") = result.Text
        If x = 10 Then datcuadro2.Recordset.Fields("mes11") = result.Text
        If x = 11 Then datcuadro2.Recordset.Fields("mes12") = result.Text
sigue4:
    Next x
   datcuadro2.Recordset.UpdateBatch adAffectCurrent
Next Y


Call Command5_Click

End Sub

Private Sub Command2_Click(Index As Integer)
Dim mes(12) As String
Dim sumatoria(20, 20) As Currency
Dim sumatoriacob(20, 20) As Currency
Dim sumatoriadev(20, 20) As Currency
Dim totales(20, 20) As Currency
Dim totalgral0(20, 20) As Currency

Text4.Text = command2(Index).Caption

For x = 1 To 20
    For Y = 1 To 12
        sumatoria(x, Y) = 0
        sumatoriacob(x, Y) = 0
        sumatoriadev(x, Y) = 0
        totalgral0(x, Y) = 0
        totales(x, Y) = 0
    Next Y
Next x
meshasta = Month(hasta.Value)
If Year(hasta.Value) > Year(desde.Value) Then meshasta = meshasta + 12

    
mesmuestra = Index + meshasta - mesesmuestra + 1
datresulventas1.RecordSource = "select resultadosventas1.* from resultadosventas1 where empresa = " & login.empresaact & " and fecha >= '" & desde.Value & "' and fecha <= '" & hasta.Value & "' order by fecha"
datresulventas1.Refresh
total = 0
cobradas = 0
If datresulventas1.Recordset.EOF = True Then Exit Sub
datresulventas1.Recordset.MoveFirst
Do While Not datresulventas1.Recordset.EOF
    If IsNull(datresulventas1.Recordset.Fields("codanun")) = True Then
        If IsNull(datresulventas1.Recordset.Fields("codclien")) = True Then
            codi = 0
        Else
            codi = datresulventas1.Recordset.Fields("codclien")
        End If
    Else
        codi = datresulventas1.Recordset.Fields("codanun")
    End If
    
    If Left(datresulventas1.Recordset.Fields("avisador"), 6) = "COMUNA" Or Left(datresulventas1.Recordset.Fields("avisador"), 6) = "Comuna" Or Left(datresulventas1.Recordset.Fields("avisador"), 6) = "comuna" Then codi = 4
    If Left(datresulventas1.Recordset.Fields("avisador"), 5) = "MUNIC" Or Left(datresulventas1.Recordset.Fields("avisador"), 5) = "Munic" Or Left(datresulventas1.Recordset.Fields("avisador"), 5) = "munic" Then codi = 4
    
    mesfecha = Month(datresulventas1.Recordset.Fields("fecha"))
    total = datresulventas1.Recordset.Fields("expr1") + datresulventas1.Recordset.Fields("expr2")
    cobradas = datresulventas1.Recordset.Fields("expr1") + datresulventas1.Recordset.Fields("expr2")
    
    sumatoriacob(codi, mesfecha) = cobradas + sumatoriacob(codi, mesfecha)
    
    datresulventas1.Recordset.MoveNext

Loop

For x = 0 To 11
  If Option1(x).Value = True Then mesacumul = x + 1
Next x

For x = 1 To mesacumul - 1
  For Y = 0 To cont
    totalgral(Y, mesacumul) = totalgral(Y, mesacumul) + totalgral(Y, x)
    sumatoriacob(Y, mesacumul) = sumatoriacob(Y, mesacumul) + sumatoriacob(Y, x)
  Next Y
Next x

totalfacturada = 0
totalcobra = 0
totalnograv = 0
For Y = 0 To cont
    gestion.Col = 1
    gestion.Row = Y + 1
    gestion.Text = totalgral(Y, mesmuestra)
    gestion.Text = Format(gestion.Text, "###,##0.00")
    totalfacturada = totalfacturada + gestion.Text
    factu = gestion.Text
    gestion.Col = 3
    gestion.Text = sumatoriacob(Y, mesmuestra)
    gestion.Text = Format(gestion.Text, "###,##0.00")
    cobra = gestion.Text
    totalcobra = gestion.Text + totalcobra
    gestion.Row = cont + 1
    gestion.Text = totalcobra
    gestion.Text = Format(gestion.Text, "###,##0.00")
    gestion.Row = Y + 1
    gestion.Col = 5
    gestion.Text = nogravado(Y, mesmuestra)
    gestion.Text = Format(gestion.Text, "###,##0.00")
    totalnograv = totalnograv + gestion.Text
    gestion.Col = 4
    gestion.Text = factu - cobra
    gestion.Text = Format(gestion.Text, "###,##0.00")
Next Y
    gestion.Row = cont + 1
    gestion.Text = totalfacturada - totalcobra
    gestion.Text = Format(gestion.Text, "###,##0.00")
    gestion.Col = 1
    gestion.Text = totalfacturada
    gestion.Text = Format(gestion.Text, "###,##0.00")
    gestion.Col = 5
    gestion.Text = totalnograv
    gestion.Text = Format(gestion.Text, "###,##0.00")
   
  
   
Rem ********** Informacion de Ordenes ***************

   
datordenes.RecordSource = "select ordenesfac.* from ordenesfac where mesdepublic = " & mesmuestra & " "
datordenes.Refresh
devengadas = 0
If datordenes.Recordset.EOF = True Then GoTo fuera
datordenes.Recordset.MoveFirst
Do While Not datordenes.Recordset.EOF
    If IsNull(datordenes.Recordset.Fields("codanun")) = True Then
        If IsNull(datordenes.Recordset.Fields("codclien")) = True Then
            codi = 0
        Else
            codi = datordenes.Recordset.Fields("codclien")
        End If
    Else
        codi = datordenes.Recordset.Fields("codanun")
    End If
    
    If Left(datordenes.Recordset.Fields("anunciante"), 6) = "COMUNA" Or Left(datordenes.Recordset.Fields("anunciante"), 6) = "Comuna" Or Left(datordenes.Recordset.Fields("anunciante"), 6) = "comuna" Then codi = 4
    If Left(datordenes.Recordset.Fields("anunciante"), 5) = "MUNIC" Or Left(datordenes.Recordset.Fields("anunciante"), 5) = "Munic" Or Left(datordenes.Recordset.Fields("anunciante"), 5) = "munic" Then codi = 4
    
    mesfecha = datordenes.Recordset.Fields("mesdepublic")
    devengadas = datordenes.Recordset.Fields("totalgral") - datordenes.Recordset.Fields("iva")
   
    sumatoriadev(codi, mesfecha) = devengadas + sumatoriadev(codi, mesfecha)
    datordenes.Recordset.MoveNext

Loop
    
For x = 0 To 11
  If Option1(x).Value = True Then mesacumul = x + 1
Next x

For x = 1 To mesacumul - 1
  For Y = 0 To cont
    sumatoriadev(Y, mesacumul) = sumatoriadev(Y, mesacumul) + sumatoriadev(Y, x)
  Next Y
Next x
totaldev = 0
For Y = 0 To cont
    gestion.Col = 0
    gestion.Row = Y + 1
    gestion.Text = sumatoriadev(Y, mesmuestra)
    gestion.Text = Format(gestion.Text, "###,##0.00")
    devenga = gestion.Text
    totaldev = totaldev + devenga
    gestion.Col = 1
    facturadas = gestion.Text
    gestion.Col = 2
    gestion.Text = devenga - facturadas
    gestion.Text = Format(gestion.Text, "###,##0.00")
Next Y
    gestion.Col = 0
    gestion.Row = cont + 1
    gestion.Text = totaldev
    gestion.Text = Format(gestion.Text, "###,##0.00")
    gestion.Col = 2
    gestion.Text = totaldev - facturadas
    gestion.Text = Format(gestion.Text, "###,##0.00")


Call aceptar_Click
Exit Sub

fuera:
    For x = 1 To cont + 1
        gestion.Col = 0
        gestion.Row = x
        gestion.Text = 0
        gestion.Col = 2
        gestion.Row = x
        gestion.Text = 0
    Next x


Call aceptar_Click



End Sub

Private Sub Command3_Click()

If datcuadro3.Recordset.EOF = True Then GoTo sigue
datcuadro3.Recordset.MoveFirst
Do While Not datcuadro3.Recordset.EOF
    datcuadro3.Recordset.Delete adAffectCurrent
    datcuadro3.Recordset.MoveNext
Loop

sigue:
reportear = 1
For x = 0 To 11
    If command2(x).Visible = False Then GoTo fin
    command2(x).SetFocus
    SendKeys "{ENTER}", True
    Call Command4_Click
Next x
fin:
Call Command6_Click


End Sub

Private Sub Command4_Click()

For Y = 0 To cont - 1
    datcuadro3.Recordset.AddNew
    datcuadro3.Recordset.Fields("desde") = desde.Value
    datcuadro3.Recordset.Fields("hasta") = hasta.Value
    datcuadro3.Recordset.Fields("mes") = Text4.Text
    datcuadro3.Recordset.Fields("mesnum") = mesmuestra
    datcuadro3.Recordset.Fields("tipocliente") = Y
    datcuadro3.Recordset.Fields("cliente") = Text3(Y).Text
    gestion.Row = Y + 1
        gestion.Col = 0
        datcuadro3.Recordset.Fields("devengada") = gestion.Text
        gestion.Col = 1
        datcuadro3.Recordset.Fields("facturada") = gestion.Text
        gestion.Col = 2
        datcuadro3.Recordset.Fields("pendientefac") = gestion.Text
        gestion.Col = 3
        datcuadro3.Recordset.Fields("cobradas") = gestion.Text
        gestion.Col = 4
        datcuadro3.Recordset.Fields("pendientecob") = gestion.Text
        gestion.Col = 5
        datcuadro3.Recordset.Fields("facnogravada") = gestion.Text
    datcuadro3.Recordset.UpdateBatch adAffectCurrent
Next Y

End Sub

Private Sub Command5_Click()
On Error Resume Next
Dim crApp As New CRAXDRT.Application
Dim crReport As New CRAXDRT.Report
Dim tabla As String
Dim tabla1 As String
Dim ruta As String


ruta = "\Empresa" + Right(Str(login.empresaact), Len(Str(login.empresaact)) - 1)

Rem *********** Cuadro Resultado  **************
reporte.SQL = "SELECT cuadroresultados2.id, cuadroresultados2.ventasuotros, cuadroresultados2.centroresult, cuadroresultados2.encab1, cuadroresultados2.encab2, cuadroresultados2.encab3, cuadroresultados2.encab4, cuadroresultados2.encab5, cuadroresultados2.encab6, cuadroresultados2.encab7, cuadroresultados2.encab8, cuadroresultados2.encab9, cuadroresultados2.encab10, cuadroresultados2.encab11, cuadroresultados2.encab12, cuadroresultados2.mes1, cuadroresultados2.mes2, cuadroresultados2.mes3, cuadroresultados2.mes4, cuadroresultados2.mes5, cuadroresultados2.mes6, cuadroresultados2.mes7, cuadroresultados2.mes8, cuadroresultados2.mes9, cuadroresultados2.mes10, cuadroresultados2.mes11, cuadroresultados2.mes12, cuadroresultados2.desde, cuadroresultados2.hasta FROM contablesql.dbo.cuadroresultados2 cuadroresultados2 ORDER BY cuadroresultados2.ventasuotros ASC"
tabla = reporte.SQL

If Check1.Value = 1 Then
With CrystalReporte
    .ReportFileName = App.Path & ruta + "\cuadroresultado2.rpt"
    .Connect = login.conexionreporte
    .DiscardSavedData = True
    .RetrieveDataFiles
    .ReportSource = 0
    .SQLQuery = tabla
    
    .Destination = crptToFile
    .PrintFileType = crptCrystal
    .PrintFileName = App.Path & "\cuadro2.rpt"
    .Action = 1
End With
impresos.Show
Set crReport = crApp.OpenReport(App.Path & "\cuadro2.rpt", 1)
impresos.CRViewer1.ReportSource = crReport
impresos.CRViewer1.ViewReport
End If

Rem *********** Cuadro General  **************
reporte.SQL = "SELECT consultacuadro.nomcuenta, consultacuadro.total1, consultacuadro.total2, consultacuadro.total3, consultacuadro.total4, consultacuadro.total5, consultacuadro.idcuenta FROM contablesql.dbo.consultacuadro consultacuadro ORDER BY consultacuadro.id ASC"
tabla = reporte.SQL
If Check2.Value = 1 Then
With CrystalReporte
    .ReportFileName = App.Path & ruta + "\cuadroresultado.rpt"
    .Connect = login.conexionreporte
    .Formulas(0) = "enca1=""" & Text2(0).Text & """"
    .Formulas(1) = "enca2=""" & Text2(1).Text & """"
    .Formulas(2) = "enca3=""" & Text2(2).Text & """"
    .Formulas(3) = "enca4=""" & Text2(3).Text & """"
    .Formulas(4) = "enca5=""" & Text2(4).Text & """"
    .Formulas(5) = "enca6=""" & Text2(5).Text & """"
    .Formulas(6) = "enca7=""" & Text2(6).Text & """"
    .Formulas(7) = "enca8=""" & Text2(7).Text & """"
    .Formulas(8) = "enca9=""" & Text2(8).Text & """"
    .Formulas(9) = "enca10=""" & Text2(9).Text & """"
    
    .DiscardSavedData = True
    .RetrieveDataFiles
    .ReportSource = 0
    .SQLQuery = tabla
    
    .Destination = crptToFile
    .PrintFileType = crptCrystal
    .PrintFileName = App.Path & "\cuadro.rpt"
    .Action = 1
End With
impresos1.Show
Set crReport = crApp.OpenReport(App.Path & "\cuadro.rpt", 1)
impresos1.CRViewer1.ReportSource = crReport
impresos1.CRViewer1.ViewReport
End If

Rem *********** Cuadro balance  **************
reporte.SQL = "SELECT cuadroresultado_balance.id, cuadroresultado_balance.mes1, cuadroresultado_balance.mes2, cuadroresultado_balance.mes3, cuadroresultado_balance.mes4, cuadroresultado_balance.mes5, cuadroresultado_balance.mes6, cuadroresultado_balance.mes7, cuadroresultado_balance.mes8, cuadroresultado_balance.mes9, cuadroresultado_balance.mes10, cuadroresultado_balance.mes11, cuadroresultado_balance.mes12, cuadroresultado_balance.desde, cuadroresultado_balance.hasta, cuadroresultado_balance.balance FROM contablesql.dbo.cuadroresultado_balance cuadroresultado_balance "
tabla = reporte.SQL
If Check4.Value = 1 Then
With CrystalReporte
    .ReportFileName = App.Path & ruta + "\cuadroresultado_balance.rpt"
    .Connect = login.conexionreporte
    .DiscardSavedData = True
    .RetrieveDataFiles
    .ReportSource = 0
    .SQLQuery = tabla
    
    .Destination = crptToFile
    .PrintFileType = crptCrystal
    .PrintFileName = App.Path & "\cuadro1.rpt"
    .Action = 1
End With
impresos1.Show
Set crReport = crApp.OpenReport(App.Path & "\cuadro1.rpt", 1)
impresos1.CRViewer1.ReportSource = crReport
impresos1.CRViewer1.ViewReport
End If


End Sub

Private Sub Command6_Click()
Dim crApp As New CRAXDRT.Application
Dim crReport As New CRAXDRT.Report
Dim tabla As String
Dim tabla1 As String
Dim ruta As String


ruta = "\Empresa" + Right(Str(login.empresaact), Len(Str(login.empresaact)) - 1)

Rem *********** gestion comercial  **************
reporte.SQL = "SELECT cuadroresultados3.mes, cuadroresultados3.mesnum, cuadroresultados3.cliente, cuadroresultados3.devengada, cuadroresultados3.facturada, cuadroresultados3.pendientefac, cuadroresultados3.cobradas, cuadroresultados3.pendientecob, cuadroresultados3.facnogravada, cuadroresultados3.desde, cuadroresultados3.hasta FROM contablesql.dbo.cuadroresultados3 cuadroresultados3 ORDER BY cuadroresultados3.mesnum ASC"
tabla = reporte.SQL

With CrystalReporte
    .ReportFileName = App.Path & ruta + "\cuadroresultado3.rpt"
    .Connect = login.conexionreporte
    .DiscardSavedData = True
    .RetrieveDataFiles
    .ReportSource = 0
    .SQLQuery = tabla
    
    .Destination = crptToFile
    .PrintFileType = crptCrystal
    .PrintFileName = App.Path & "\cuadro3.rpt"
    .Action = 1
End With
impresos.Show
Set crReport = crApp.OpenReport(App.Path & "\cuadro3.rpt", 1)
impresos.CRViewer1.ReportSource = crReport
impresos.CRViewer1.ViewReport

End Sub

Private Sub Command7_Click()
On Error GoTo fuera
Dim crApp As New CRAXDRT.Application
Dim crReport As New CRAXDRT.Report
Dim tabla As String
Dim tabla1 As String
Dim ruta As String



desdefecha = "     " + Str(Year(desde)) + "        " + Str(Month(desde))
hastafecha = "     " + Str(Year(hasta)) + "        " + Str(Month(hasta))

ruta = "\Empresa" + Right(Str(login.empresaact), Len(Str(login.empresaact)) - 1)

Rem *********** gestion comercial  **************
reporte.SQL = "SELECT consultacentrocosto1.idcuenta, consultacentrocosto1.importe, consultacentrocosto1.ccosto, consultacentrocosto1.descripcion, consultacentrocosto1.subcc1, consultacentrocosto1.Subcentro1, consultacentrocosto1.totalcc1, consultacentrocosto1.subcc2, consultacentrocosto1.Subcentro2, consultacentrocosto1.totalcc2, consultacentrocosto1.subcc3, consultacentrocosto1.Succentro3, consultacentrocosto1.totalcc3, consultacentrocosto1.Nombre Cuenta, consultacentrocosto1.añomes FROM contablesql.dbo.consultacentrocosto1 consultacentrocosto1 where consultacentrocosto1.codcont >= '" & Text5.Text & "' and consultacentrocosto1.codcont <= '" & Text6.Text & "' and consultacentrocosto1.añomes >= '" & desdefecha & "' and consultacentrocosto1.añomes <= '" & hastafecha & "' ORDER BY consultacentrocosto1.añomes ASC, consultacentrocosto1.idcuenta ASC, consultacentrocosto1.ccosto ASC, consultacentrocosto1.subcc1 ASC, consultacentrocosto1.subcc2 ASC, consultacentrocosto1.subcc3 ASC"
tabla = reporte.SQL


With CrystalReporte
    .ReportFileName = App.Path & ruta + "\ANALISISCENTROCOSTO.rpt"
    .Formulas(0) = "nivel1=""" & Option2(0).Value & """"
    .Formulas(1) = "nivel2=""" & Option2(1).Value & """"
    .Formulas(2) = "nivel3=""" & Option2(2).Value & """"
    .Formulas(3) = "nivel4=""" & Option2(3).Value & """"
    .Formulas(4) = "ninguno=""" & Check3.Value & """"
    .Connect = login.conexionreporte
    .DiscardSavedData = True
    .RetrieveDataFiles
    .ReportSource = 0
    .SQLQuery = tabla
    
    .Destination = crptToFile
    .PrintFileType = crptCrystal
    .PrintFileName = App.Path & "\cuadro3.rpt"
    .Action = 1
End With
impresos.Show
Set crReport = crApp.OpenReport(App.Path & "\cuadro3.rpt", 1)
impresos.CRViewer1.ReportSource = crReport
impresos.CRViewer1.ViewReport
impresos.CRViewer1.EnableGroupTree = True
fuera:
End Sub

Private Sub Command8_Click()
Dim crApp As New CRAXDRT.Application
Dim crReport As New CRAXDRT.Report
Dim tabla As String
Dim tabla1 As String
Dim ruta As String


ruta = "\Empresa" + Right(Str(login.empresaact), Len(Str(login.empresaact)) - 1)

Rem *********** gestion comercial  **************
reporte.SQL = "SELECT consultacentrocosto1.idcuenta, consultacentrocosto1.importe, consultacentrocosto1.ccosto, consultacentrocosto1.descripcion, consultacentrocosto1.subcc1, consultacentrocosto1.Subcentro1, consultacentrocosto1.totalcc1, consultacentrocosto1.subcc2, consultacentrocosto1.Subcentro2, consultacentrocosto1.totalcc2, consultacentrocosto1.subcc3, consultacentrocosto1.Succentro3, consultacentrocosto1.totalcc3, consultacentrocosto1.Nombre Cuenta, consultacentrocosto1.añomes FROM contablesql.dbo.consultacentrocosto1 consultacentrocosto1 ORDER BY consultacentrocosto1.añomes ASC, consultacentrocosto1.idcuenta ASC, consultacentrocosto1.ccosto ASC, consultacentrocosto1.subcc1 ASC, consultacentrocosto1.subcc2 ASC, consultacentrocosto1.subcc3 ASC"
tabla = reporte.SQL

With CrystalReporte
    .ReportFileName = App.Path & ruta + "\ANALISISCENTROCOSTO.rpt"
    .Connect = login.conexionreporte
    .DiscardSavedData = True
    .RetrieveDataFiles
    .ReportSource = 0
    .SQLQuery = tabla
    
    .Destination = crptToFile
    .PrintFileType = crptCrystal
    .PrintFileName = App.Path & "\cuadro3.rpt"
    .Action = 1
End With
impresos.Show
Set crReport = crApp.OpenReport(App.Path & "\cuadro3.rpt", 1)
impresos.CRViewer1.ReportSource = crReport
impresos.CRViewer1.ViewReport

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
On Error GoTo error1

    If KeyAscii = 13 Then
        KeyAscii = 0
        If posi = 1 Then
            Text5.Text = DataList2.BoundText
            Text6.SetFocus
        End If
        If posi = 2 Then
            Text6.Text = DataList2.BoundText
            command7.SetFocus
        End If
    End If
Exit Sub
error1:

End Sub

Private Sub DataList2_LostFocus()

    DataList2.Visible = False

End Sub

Private Sub desde_Change()

    calen.Value = desde.Value
    Call seleccionames_Click

End Sub

Private Sub fondoscalcula_Click()
On Error GoTo fuera
Dim campo(100) As Currency

For x = 1 To 100
      campo(x) = 0
Next x
   
    datparamresultados.RecordSource = "select paramresultados.*  from paramresultados where empresa = " & login.empresaact & " and activado <> 0 order by id"
    datparamresultados.Refresh
     
datcuadro.RecordSource = "select cuadroresultados.* from cuadroresultados"
datcuadro.Refresh
If datcuadro.Recordset.EOF = True Then GoTo calcula

datcuadro.Recordset.MoveFirst
Do While Not datcuadro.Recordset.EOF
        datcuadro.Recordset.Delete adAffectCurrent
        datcuadro.Recordset.MoveNext
Loop

calcula:
If datparamresultados.Recordset.EOF = True Then Exit Sub
bar1.Min = 0
bar1.max = datparamresultados.Recordset.RecordCount
barcont = 0
datparamresultados.Recordset.MoveFirst
cont1 = datparamresultados.Recordset.Fields("id") - 1
Do While Not datparamresultados.Recordset.EOF

    barcont = barcont + 1
    bar1.Value = barcont

    niveles = datparamresultados.Recordset.Fields("cuentacabecera")
    cuentafiltro = ""
    niveles0 = "1"
    For x = 1 To Len(niveles)
        nivelesc = Mid(niveles, x, 1)
        cuentafiltro = cuentafiltro + nivelesc
        If nivelesc = "." Then
            nivelesp = ""
            niveles0 = ""
            For Y = x + 1 To Len(niveles)
                    nivelesp = Mid(niveles, Y, 1)
                    niveles0 = nivelesp + niveles0
                    If nivelesp = "." Then GoTo sale
            Next Y
        End If
sale:
        If Val(niveles0) = 0 Then GoTo sale1
    Next x
sale1:


Rem ****** filtra acumulado ***********
    datfiltro.RecordSource = "select filtrocuenta.* from filtrocuenta"
    datfiltro.Refresh
    If datfiltro.Recordset.EOF = True Then datfiltro.Recordset.AddNew
    datfiltro.Recordset.Fields(0) = cuentafiltro
    datfiltro.Recordset.Fields("desde") = desde.Value
    datfiltro.Recordset.Fields("hasta") = hastaacum
    datfiltro.Recordset.Fields("fondo") = datparamresultados.Recordset.Fields("id")
    datfiltro.Recordset.UpdateBatch adAffectCurrent
    
    datlibro.RecordSource = "select librodiario11.* from librodiario11"
    datlibro.Refresh
    If datlibro.Recordset.EOF = True Then GoTo finlibro
    cont0 = datparamresultados.Recordset.Fields("id") - cont1
    result.Col = 0
    result.Row = cont0
    If IsNull(datlibro.Recordset.Fields(0)) = True Then datlibro.Recordset.Fields(0) = 0
    result.Text = datlibro.Recordset.Fields(0)
    result.Text = Format(Val(result.Text), "###,##0.00")
    
    
Rem ******* filtra meses  ********************
For x = 1 To mesesmuestra
    datfiltro.RecordSource = "select filtrocuenta.* from filtrocuenta"
    datfiltro.Refresh
    If datfiltro.Recordset.EOF = True Then datfiltro.Recordset.AddNew
    datfiltro.Recordset.Fields(0) = cuentafiltro

    If Year(hasta.Value) > Year(desde.Value) Then
        mhasta = Month(hasta.Value) + 12
    Else
        mhasta = Month(hasta.Value)
    End If

    datfiltro.Recordset.Fields("desde") = mesidesde(mhasta - mesesmuestra + x)
    datfiltro.Recordset.Fields("hasta") = mesihasta(mhasta - mesesmuestra + x)
    datfiltro.Recordset.Fields("fondo") = datparamresultados.Recordset.Fields("id")
    datfiltro.Recordset.UpdateBatch adAffectCurrent
    
    datlibro.RecordSource = "select librodiario11.* from librodiario11"
    datlibro.Refresh
    If datlibro.Recordset.EOF = True Then GoTo finlibro
    cont0 = datparamresultados.Recordset.Fields("id") - cont1
    result.Col = x - 1
    result.Row = cont0
    If IsNull(datlibro.Recordset.Fields(0)) = True Then datlibro.Recordset.Fields(0) = 0
    result.Text = datlibro.Recordset.Fields(0)
    result.Text = Format(Val(result.Text), "###,##0.00")
Next x

finlibro0:
    datfiltro.RecordSource = "select filtrocuenta.* from filtrocuenta"
    datfiltro.Refresh
    If datfiltro.Recordset.EOF = True Then datfiltro.Recordset.AddNew
    datfiltro.Recordset.Fields(0) = cuentafiltro
    datfiltro.Recordset.Fields("desde") = desde.Value
    datfiltro.Recordset.Fields("hasta") = hasta.Value
    datfiltro.Recordset.Fields("fondo") = datparamresultados.Recordset.Fields("id")
    datfiltro.Recordset.UpdateBatch adAffectCurrent
    
        
    datlibro.RecordSource = "select librodiario10.* from librodiario10 order by idcuenta"
    datlibro.Refresh
    bar2.Min = 0
    If datlibro.Recordset.RecordCount <> 0 Then
        bar2.max = datlibro.Recordset.RecordCount
    Else
        bar2.max = 1
    End If
        
        
    barcont2 = 0
    If datlibro.Recordset.EOF = True Then GoTo finlibro
    datlibro.Recordset.MoveFirst
    cont0 = datparamresultados.Recordset.Fields("id") - cont1
    Do While Not datlibro.Recordset.EOF
            
            barcont2 = barcont2 + 1
            bar2.Value = barcont2
            datcuadro.Recordset.AddNew
            datcuadro.Recordset.Fields("nomcuenta") = datlibro.Recordset.Fields("nombrecuenta")
            datcuadro.Recordset.Fields("idcuenta") = datlibro.Recordset.Fields("idcuenta")
            If cont0 = 1 Then datcuadro.Recordset.Fields("campo1") = datlibro.Recordset.Fields("total")
            If cont0 = 2 Then datcuadro.Recordset.Fields("campo2") = datlibro.Recordset.Fields("total")
            If cont0 = 3 Then datcuadro.Recordset.Fields("campo3") = datlibro.Recordset.Fields("total")
            If cont0 = 4 Then datcuadro.Recordset.Fields("campo4") = datlibro.Recordset.Fields("total")
            If cont0 = 5 Then datcuadro.Recordset.Fields("campo5") = datlibro.Recordset.Fields("total")
            If cont0 = 6 Then datcuadro.Recordset.Fields("campo6") = datlibro.Recordset.Fields("total")
            If cont0 = 7 Then datcuadro.Recordset.Fields("campo7") = datlibro.Recordset.Fields("total")
            If cont0 = 8 Then datcuadro.Recordset.Fields("campo8") = datlibro.Recordset.Fields("total")
            If cont0 = 9 Then datcuadro.Recordset.Fields("campo9") = datlibro.Recordset.Fields("total")
            If cont0 = 10 Then datcuadro.Recordset.Fields("campo10") = datlibro.Recordset.Fields("total")
            datcuadro.Recordset.UpdateBatch adAffectCurrent
            datlibro.Recordset.MoveNext
    Loop
    
finlibro:
    datparamresultados.Recordset.MoveNext

Loop
bar1.Value = 0
bar2.Value = 0
Exit Sub

fuera:
MsgBox "Error en la seleccion del rango, verifique", vbCritical, "Error"

End Sub

Private Sub Form_Load()
frmresultados.Top = 0
frmresultados.Left = 0
Option1(0).Value = True
Option2(3).Value = True
Check3.Value = 0

Check1.Value = 0
Check2.Value = 0
reportear = 0
For x = 0 To 11
    ventas.ColWidth(x) = 1500
    result.ColWidth(x) = 1500
Next x
For x = 0 To 5
    gestion.ColWidth(x) = 1200
Next x
gestion.Col = 0
gestion.Row = 0
gestion.Text = "DEVENGADA"
gestion.Col = 1
gestion.Row = 0
gestion.Text = "FACTURADAS"
gestion.Col = 2
gestion.Row = 0
gestion.Text = "PEND.FACT"
gestion.Col = 3
gestion.Row = 0
gestion.Text = "COBRADAS"
gestion.Col = 4
gestion.Row = 0
gestion.Text = "PEND.COBR"
gestion.Col = 5
gestion.Row = 0
gestion.Text = "FAC.NO GRAV"

SSTab1.Tab = 0
desde.Value = login.iper
hasta.Value = Date - Day(Date)
calen.Value = Date

    datresulventas.ConnectionString = login.conexiontotal
    datresulventas1.ConnectionString = login.conexiontotal
    dattipoclientes.ConnectionString = login.conexiontotal
    datlibro.ConnectionString = login.conexiontotal
    datparamresultados.ConnectionString = login.conexiontotal
    datparamresultados1.ConnectionString = login.conexiontotal
    datfiltro.ConnectionString = login.conexiontotal
    datcuadro.ConnectionString = login.conexiontotal
    datcuadro2.ConnectionString = login.conexiontotal
    datcuadro3.ConnectionString = login.conexiontotal
    datcuentas.ConnectionString = login.conexiontotal


    Inicio.Toolbar1.Visible = True

    datcuentas.RecordSource = "select listacuentas.* from listacuentas WHERE empre = " & login.empresaact & " and inicioper = '" & login.iper & "' ORDER BY IDCUENTA"
    datcuentas.Refresh

    dattipoclientes.RecordSource = "select tipoclientes.* from tipoclientes where empresa = " & login.empresaact & " order by codigo"
    dattipoclientes.Refresh
    
    datcuadro3.RecordSource = "select cuadroresultados3.* from cuadroresultados3"
    datcuadro3.Refresh
    
    datparamresultados.RecordSource = "select paramresultados.*  from paramresultados where empresa = " & login.empresaact & " and activado <> 0 order by id"
    datparamresultados.Refresh
       
If dattipoclientes.Recordset.EOF = True Then GoTo sale0
dattipoclientes.Recordset.MoveFirst
cont = 1

Do While Not dattipoclientes.Recordset.EOF

    Text1(cont).Text = dattipoclientes.Recordset.Fields("tipoclientes")
    Text3(cont).Text = Text1(cont).Text
    dattipoclientes.Recordset.MoveNext
    cont = cont + 1
    If cont = 11 Then GoTo sale0
    
Loop

sale0:
maxventas = cont
If datparamresultados.Recordset.EOF = True Then GoTo sale
datparamresultados.Recordset.MoveFirst
contin = 0

Do While Not datparamresultados.Recordset.EOF

    Text2(contin).Text = datparamresultados.Recordset.Fields("nombrefondo")
    grupo(contin) = datparamresultados.Recordset.Fields("grupo")
    datparamresultados.Recordset.MoveNext
    contin = contin + 1
    If contin = 10 Then GoTo sale
    
Loop



sale:
maxresultados = contin - 1
    Text1(0).Text = "NO DEFINIDO"
    Text3(0).Text = "NO DEFINIDO"
    Text1(cont).Text = "Total Ventas"
    Text3(cont).Text = Text1(cont).Text
    For x = cont + 1 To 10
        Text1(x).Visible = False
        Text3(x).Visible = False
    Next x
    For x = contin To 9
        Text2(x).Visible = False
    Next x
  
    ventas.Rows = cont + 2
    gestion.Rows = cont + 2
    result.Rows = contin + 1
    ventas.Height = ventas.RowHeight(0) * (cont + 4)
    result.Height = result.RowHeight(0) * (contin + 4)

        
End Sub

Private Sub hasta_CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)
    Call seleccionames_Click
End Sub

Private Sub Option1_Click(Index As Integer)
On Error GoTo fuera

calen.Value = desde.Value
If Index = 12 Then Exit Sub
calen.Day = 1
calen.Month = Index + 1
index0 = Index + 1

If index0 = 1 Or index0 = 3 Or index0 = 5 Or index0 = 7 Or index0 = 8 Or index0 = 10 Or index0 = 12 Then calen.Day = 31
If index0 = 4 Or index0 = 6 Or index0 = 9 Or index0 = 11 Then calen.Day = 30
If index0 = 2 Then calen.Day = 29

If Year(hasta.Value) > Year(desde.Value) Then
        mhasta = Month(hasta.Value) + 12
Else
        mhasta = Month(hasta.Value)
End If

        For x = index0 To mhasta
            Y = x
            añito = calen.Year
            If x > 12 Then
                Y = x - 12
                añito = añito + 1
            End If
            mesito = Mid("00", 1, 3 - (Len(Str(Y)))) + Right(Str(Y), Len(Str(Y)) - 1)
            mesidesde(x) = "01/" + mesito + "/" + Right(Str(añito), 4)
            If Y = 1 Or Y = 3 Or Y = 5 Or Y = 7 Or Y = 8 Or Y = 10 Or Y = 12 Then
                mesihasta(x) = "31/" + mesito + "/" + Right(Str(añito), 4)
            Else
                If Y = 2 Then
                    mesihasta(x) = "29/" + mesito + "/" + Right(Str(añito), 4)
                Else
                    mesihasta(x) = "30/" + mesito + "/" + Right(Str(añito), 4)
                End If
            End If
        Next x

        
    

hastaacum = calen.Value
Exit Sub
fuera:
calen.Day = 29
hastaacum = calen.Value
        For x = index0 To Month(hasta.Value)
            mesito = Mid("00", 1, 3 - (Len(Str(x)))) + Right(Str(x), Len(Str(x)) - 1)
            mesidesde(x) = "01/" + mesito + "/" + Right(Str(calen.Year), 4)
            If x = 3 Or x = 5 Or x = 7 Or x = 8 Or x = 10 Or x = 12 Then
                mesihasta(x) = "31/" + mesito + "/" + Right(Str(calen.Year), 4)
            Else
                If x = 2 Then
                    mesihasta(x) = "29/" + mesito + "/" + Right(Str(calen.Year), 4)
                Else
                    mesihasta(x) = "30/" + mesito + "/" + Right(Str(calen.Year), 4)
                End If
            End If
        Next x


End Sub

Private Sub Option2_Click(Index As Integer)

    Check3.Value = 0

End Sub

Private Sub seleccionames_Click()
On Error GoTo fuera

calen.Value = desde.Value
If Index = 12 Then Exit Sub
calen.Day = 1
calen.Month = Index + 1
index0 = Index + 1

If index0 = 1 Or index0 = 3 Or index0 = 5 Or index0 = 7 Or index0 = 8 Or index0 = 10 Or index0 = 12 Then calen.Day = 31
If index0 = 4 Or index0 = 6 Or index0 = 9 Or index0 = 11 Then calen.Day = 30
If index0 = 2 Then calen.Day = 29

If Year(hasta.Value) > Year(desde.Value) Then
        mhasta = Month(hasta.Value) + 12
Else
        mhasta = Month(hasta.Value)
End If

        For x = index0 To mhasta
            Y = x
            añito = calen.Year
            If x > 12 Then
                Y = x - 12
                añito = añito + 1
            End If
            mesito = Mid("00", 1, 3 - (Len(Str(Y)))) + Right(Str(Y), Len(Str(Y)) - 1)
            mesidesde(x) = "01/" + mesito + "/" + Right(Str(añito), 4)
            If Y = 1 Or Y = 3 Or Y = 5 Or Y = 7 Or Y = 8 Or Y = 10 Or Y = 12 Then
                mesihasta(x) = "31/" + mesito + "/" + Right(Str(añito), 4)
            Else
                If Y = 2 Then
                    mesihasta(x) = "29/" + mesito + "/" + Right(Str(añito), 4)
                Else
                    mesihasta(x) = "30/" + mesito + "/" + Right(Str(añito), 4)
                End If
            End If
        Next x

        
    

hastaacum = calen.Value
Exit Sub
fuera:
calen.Day = 29
hastaacum = calen.Value
        For x = index0 To Month(hasta.Value)
            mesito = Mid("00", 1, 3 - (Len(Str(x)))) + Right(Str(x), Len(Str(x)) - 1)
            mesidesde(x) = "01/" + mesito + "/" + Right(Str(calen.Year), 4)
            If x = 3 Or x = 5 Or x = 7 Or x = 8 Or x = 10 Or x = 12 Then
                mesihasta(x) = "31/" + mesito + "/" + Right(Str(calen.Year), 4)
            Else
                If x = 2 Then
                    mesihasta(x) = "29/" + mesito + "/" + Right(Str(calen.Year), 4)
                Else
                    mesihasta(x) = "30/" + mesito + "/" + Right(Str(calen.Year), 4)
                End If
            End If
        Next x

End Sub

Private Sub SSTab1_GotFocus()

    If SSTab1.Tab = 1 Then Call aceptar_Click
Rem    If SSTab1.Tab = 2 Then Call fondoscalcula_Click

End Sub

Private Sub Text5_GotFocus()
        
    posi = 1
    DataList2.Visible = True
    DataList2.Top = Text5.Top + 700
    DataList2.SetFocus
        

End Sub

Private Sub Text6_GotFocus()

        
    posi = 2
    DataList2.Visible = True
    DataList2.Top = Text6.Top + 700
    DataList2.SetFocus
        

End Sub
