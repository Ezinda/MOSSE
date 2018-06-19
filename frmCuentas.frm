VERSION 5.00
Object = "{F6125AB1-8AB1-11CE-A77F-08002B2F4E98}#2.0#0"; "MSRDC20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{5734474E-78D3-4254-99B9-C35F31BDF509}#62.0#0"; "vbskpro2.ocx"
Object = "{D18BBD1F-82BB-4385-BED3-E9D31A3E361E}#1.0#0"; "KewlButtonz.ocx"
Begin VB.Form frmCuentas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Plan de Cuentas "
   ClientHeight    =   7320
   ClientLeft      =   1095
   ClientTop       =   330
   ClientWidth     =   11595
   HasDC           =   0   'False
   Icon            =   "frmCuentas.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7320
   ScaleWidth      =   11595
   Begin VB.CommandButton actualizaplan 
      Caption         =   "actualizaplan"
      Height          =   375
      Left            =   1920
      TabIndex        =   48
      Top             =   6960
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Frame Frame7 
      Caption         =   "Frame7"
      Height          =   1935
      Left            =   5160
      TabIndex        =   44
      Top             =   2760
      Visible         =   0   'False
      Width           =   3015
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H80000000&
         Height          =   735
         Left            =   480
         Locked          =   -1  'True
         TabIndex        =   46
         Top             =   480
         Width           =   2175
      End
      Begin VB.CommandButton mensage 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   1080
         TabIndex        =   45
         Top             =   1320
         Width           =   975
      End
      Begin KewlButtonz.KewlButtons KewlButtons1 
         Height          =   1575
         Left            =   120
         TabIndex        =   47
         Top             =   240
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   2778
         BTYPE           =   14
         TX              =   "KewlButtons1"
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
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmCuentas.frx":0442
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
   Begin VB.Frame Frame1 
      Height          =   1575
      Left            =   120
      TabIndex        =   35
      Top             =   4680
      Width           =   5775
      Begin VB.Frame Frame3 
         Height          =   1575
         Left            =   2160
         TabIndex        =   40
         Top             =   0
         Width           =   2055
         Begin VB.OptionButton Option5 
            Caption         =   "Todas"
            Height          =   255
            Left            =   240
            TabIndex        =   43
            Top             =   1080
            Width           =   1575
         End
         Begin VB.OptionButton Option4 
            Caption         =   "No Imputables"
            Height          =   255
            Left            =   240
            TabIndex        =   42
            Top             =   720
            Width           =   1575
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Imputables"
            Height          =   255
            Left            =   240
            TabIndex        =   41
            Top             =   360
            Width           =   1575
         End
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Ordenar X Cod.Contable"
         Height          =   375
         Left            =   240
         TabIndex        =   39
         Top             =   840
         Width           =   1815
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Ordenar X Nombre"
         Height          =   255
         Left            =   240
         TabIndex        =   38
         Top             =   360
         Width           =   1815
      End
      Begin KewlButtonz.KewlButtons imprimir 
         Height          =   615
         Left            =   4440
         TabIndex        =   36
         Top             =   240
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   1085
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
         MICON           =   "frmCuentas.frx":045E
         PICN            =   "frmCuentas.frx":047A
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
         Left            =   4440
         TabIndex        =   37
         Top             =   840
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
         MICON           =   "frmCuentas.frx":386C
         PICN            =   "frmCuentas.frx":3888
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmCuentas.frx":43D2
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   2040
      Visible         =   0   'False
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   661
      _Version        =   393216
      AllowArrows     =   -1  'True
      BackColor       =   14737632
      ForeColor       =   -2147483642
      HeadLines       =   1
      RowHeight       =   15
      TabAction       =   2
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
         MarqueeStyle    =   4
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton reordenar 
      Caption         =   "Reordenar Codigos"
      Height          =   735
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   6360
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Frame Frame6 
      Height          =   1335
      Left            =   120
      TabIndex        =   33
      Top             =   3360
      Width           =   5775
      Begin KewlButtonz.KewlButtons grabar 
         Height          =   615
         Left            =   1560
         TabIndex        =   3
         Top             =   360
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   1085
         BTYPE           =   14
         TX              =   "&Grabar"
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
         MICON           =   "frmCuentas.frx":43ED
         PICN            =   "frmCuentas.frx":4409
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin KewlButtonz.KewlButtons agregamismonivel 
         Height          =   615
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   1085
         BTYPE           =   14
         TX              =   "&Nuevo"
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
         MICON           =   "frmCuentas.frx":5E8B
         PICN            =   "frmCuentas.frx":5EA7
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin KewlButtonz.KewlButtons cancelar 
         Cancel          =   -1  'True
         Height          =   615
         Left            =   3000
         TabIndex        =   4
         Top             =   360
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   1085
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
         MICON           =   "frmCuentas.frx":9299
         PICN            =   "frmCuentas.frx":92B5
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
         Height          =   615
         Left            =   4440
         TabIndex        =   5
         Top             =   360
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   1085
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
         MICON           =   "frmCuentas.frx":9CC7
         PICN            =   "frmCuentas.frx":9CE3
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
   Begin VB.CommandButton llena 
      Caption         =   "llena"
      Height          =   375
      Left            =   2160
      TabIndex        =   32
      Top             =   6360
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton acttree 
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   6720
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Frame Frame5 
      Height          =   2055
      Left            =   120
      TabIndex        =   19
      Top             =   120
      Width           =   5775
      Begin VB.CommandButton Command2 
         Caption         =   "&Modifica"
         Height          =   255
         Left            =   3720
         TabIndex        =   49
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox Text8 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   2
         Left            =   1920
         TabIndex        =   2
         Top             =   1560
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Cuenta Imputable:"
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
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   1560
         Width           =   1695
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Nombre:"
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
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   960
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Cod.Cont.:"
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
         Left            =   120
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox Text9 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   5
         Left            =   4560
         Locked          =   -1  'True
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox Text9 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   4
         Left            =   3960
         Locked          =   -1  'True
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox Text9 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   3
         Left            =   3360
         Locked          =   -1  'True
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox Text9 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   2
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox Text9 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   1
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox Text9 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   0
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox Text8 
         Alignment       =   2  'Center
         Height          =   525
         Index           =   1
         Left            =   1560
         MultiLine       =   -1  'True
         TabIndex        =   1
         Top             =   960
         Width           =   3615
      End
      Begin VB.TextBox Text8 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   0
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   0
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Cuen&ta:"
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
         Index           =   0
         Left            =   480
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   240
         Width           =   975
      End
   End
   Begin MSAdodcLib.Adodc datPrimaryRS 
      Height          =   330
      Left            =   2160
      Top             =   6720
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
      CacheSize       =   2000
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   1
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
   Begin VB.TextBox detalle 
      BackColor       =   &H80000013&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   2280
      Width           =   5775
   End
   Begin VB.Frame Frame4 
      Height          =   15
      Left            =   120
      TabIndex        =   12
      Top             =   0
      Width           =   11655
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   6855
      Left            =   6000
      TabIndex        =   6
      Top             =   240
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   12091
      _Version        =   393217
      Indentation     =   529
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      FullRowSelect   =   -1  'True
      HotTracking     =   -1  'True
      ImageList       =   "iml16"
      BorderStyle     =   1
      Appearance      =   1
      MousePointer    =   99
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "frmCuentas.frx":D0D5
   End
   Begin VB.TextBox imputable 
      Alignment       =   2  'Center
      DataField       =   "imp"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Left            =   1920
      TabIndex        =   8
      Top             =   960
      Width           =   735
   End
   Begin VB.TextBox idcuenta0 
      Alignment       =   2  'Center
      DataField       =   "idcuenta"
      DataSource      =   "datPrimaryRS"
      Enabled         =   0   'False
      Height          =   285
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Frame Frame2 
      Height          =   1695
      Left            =   120
      TabIndex        =   11
      Top             =   360
      Visible         =   0   'False
      Width           =   5775
      Begin VB.TextBox codcuenta 
         Alignment       =   2  'Center
         DataField       =   "idcuentacod"
         DataSource      =   "datPrimaryRS"
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   30
         Top             =   240
         Width           =   1455
      End
   End
   Begin MSAdodcLib.Adodc niveles 
      Height          =   330
      Left            =   3600
      Top             =   6720
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
   Begin VB.TextBox car5 
      DataField       =   "niv5"
      DataSource      =   "niveles"
      Height          =   285
      Left            =   7680
      TabIndex        =   17
      Text            =   "Text6"
      Top             =   120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox car4 
      DataField       =   "niv4"
      DataSource      =   "niveles"
      Height          =   285
      Left            =   7320
      TabIndex        =   16
      Text            =   "Text5"
      Top             =   120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox car3 
      DataField       =   "niv3"
      DataSource      =   "niveles"
      Height          =   285
      Left            =   6960
      TabIndex        =   15
      Text            =   "Text4"
      Top             =   120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox car2 
      DataField       =   "niv2"
      DataSource      =   "niveles"
      Height          =   285
      Left            =   6600
      TabIndex        =   14
      Text            =   "Text3"
      Top             =   120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox car1 
      DataField       =   "niv1"
      DataSource      =   "niveles"
      Height          =   285
      Left            =   6240
      TabIndex        =   13
      Text            =   "Text2"
      Top             =   120
      Visible         =   0   'False
      Width           =   375
   End
   Begin MSComctlLib.ImageList iml16 
      Left            =   3120
      Top             =   6360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   16777215
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCuentas.frx":D3EF
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCuentas.frx":13C51
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSRDC.MSRDC reporte 
      Height          =   375
      Left            =   3720
      Top             =   6360
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
      DataSourceName  =   "contablePS"
      RecordSource    =   ""
      UserName        =   "lucva"
      Password        =   "25072004"
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
   Begin Crystal.CrystalReport crystalreporte 
      Left            =   1560
      Top             =   6360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      WindowTitle     =   "Plan de Cuentas"
      UserName        =   "lucva"
   End
   Begin MSAdodcLib.Adodc criterio 
      Height          =   375
      Left            =   120
      Top             =   6360
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
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
      Left            =   10440
      Top             =   0
      _ExtentX        =   1270
      _ExtentY        =   1270
      CloseButton     =   0
      MaxButton       =   0
      MinButton       =   0
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
      LcK2            =   $"frmCuentas.frx":2C099
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
   Begin MSAdodcLib.Adodc datverifica 
      Height          =   375
      Left            =   120
      Top             =   6960
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
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
   Begin MSAdodcLib.Adodc datcuentas 
      Height          =   330
      Left            =   3600
      Top             =   6960
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
End
Attribute VB_Name = "frmCuentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cantnodos As Integer
Dim Cuenta As String
Dim nivelag(5) As Integer
Dim posicion(90000) As Double
Dim cuentaabrev(90000) As Double
Dim cabecera(90000) As Double
Dim hijos(90000) As Double
Dim ultimonodo2 As Integer
Dim ultimonodo3 As Integer
Dim ultimonodo5 As Integer
Dim nivelanterior As String
Dim ruta As String
Dim canthijos As Integer
Dim pos As Double
Dim bandera As Integer
Dim ni1, ni2, ni3, ni4, ni5 As String
Dim n1, n2, n3, n4, n5, c1, c2, c3, c4, c5 As Integer
Dim actualizar As String
Dim cuentanumeros As String
Dim cuentapuntos As String

Private Sub acttree_Click()
On Error Resume Next

Dim nodx As Node
Dim imputable1 As String
Dim nombrenodo As String


TreeView1.Nodes.Clear  'Limpia el Treeview
  
 If datPrimaryRS.Recordset.EOF = True Then Exit Sub
 datPrimaryRS.Recordset.MoveFirst
Do While Not datPrimaryRS.Recordset.EOF

registro = datPrimaryRS.Recordset.AbsolutePosition
Cuenta = Str(datPrimaryRS.Recordset.Fields("id cuenta"))
n1 = 2: c1 = Val(car1)
n2 = n1 + c1: c2 = Val(car2)
n3 = n2 + c2: c3 = Val(car3)
n4 = n3 + c3: c4 = Val(car4)
n5 = n4 + c4: c5 = Val(car5)

ni1 = Mid(Cuenta, n1, c1)
ni2 = Mid(Cuenta, n2, c2)
ni3 = Mid(Cuenta, n3, c3)
ni4 = Mid(Cuenta, n4, c4)
ni5 = Mid(Cuenta, n5, c5)
nivelag(1) = c1
nivelag(2) = c2
nivelag(3) = c3
nivelag(4) = c4
nivelag(5) = c5

nombrenodo = datPrimaryRS.Recordset.Fields("nombre cuenta")
imputable1 = datPrimaryRS.Recordset.Fields("imp")

If Val(ni2) = 0 Then
      nodouno = nodouno + 1
      Set nodx = TreeView1.Nodes.Add(, , "t" + ni1, nombrenodo)
      GoTo salta
End If
If Val(ni2) <> 0 And Val(ni3) = 0 Then
      Set nodx = TreeView1.Nodes.Add("t" + ni1, tvwChild, "a" + ni1 + ni2, nombrenodo)
      GoTo salta
End If
If Val(ni3) <> 0 And Val(ni4) = 0 Then
      Set nodx = TreeView1.Nodes.Add("a" + ni1 + ni2, tvwChild, "p" + ni1 + ni2 + ni3, nombrenodo)
      GoTo salta
End If
If Val(ni4) <> 0 And Val(ni5) = 0 Then
      Set nodx = TreeView1.Nodes.Add("p" + ni1 + ni2 + ni3, tvwChild, "h" + ni1 + ni2 + ni3 + ni4, nombrenodo)
      GoTo salta
End If
If Val(ni4) <> 0 And Val(ni5) <> 0 Then
      Set nodx = TreeView1.Nodes.Add("h" + ni1 + ni2 + ni3 + ni4, tvwChild, "n" + ni1 + ni2 + ni3 + ni4 + ni5, nombrenodo)
      GoTo salta
End If

salta:
posicion(nodx.Index) = registro

If Val(ni2) = 0 And Val(ni3) = 0 And Val(ni4) = 0 And Val(ni5) = 0 Then
        cabecera(ni1) = nodx.Index
End If

cuentaabrev(datPrimaryRS.Recordset.Fields("Cod Contable")) = nodx.Index
If imputable1 = "S" Then
    nodx.Image = 2
Else
    nodx.Image = 1
End If
datPrimaryRS.Recordset.MoveNext
Loop

For x = 0 To 5
    Text9(x).Enabled = True
    Text9(x).Text = ""
Next x
For x = 0 To 2
    Text8(x).Enabled = True
    Text8(x).Text = ""
Next x

fin:
For x = 1 To ultimo
     hijos(x) = TreeView1.Nodes.Item(x).Children
Next x


fuera:
End Sub



Private Sub actualizaplan_Click()
On Error Resume Next

Dim periodos(1000) As Date
Dim finperiodos(1000) As Date


    datcuentas.RecordSource = "SELECT TOP 100 PERCENT inicioper, finper, empre From cuentas where inicioper <> '" & login.iper & "'  GROUP BY inicioper, finper, empre ORDER BY inicioper"
    datcuentas.Refresh
    If datcuentas.Recordset.EOF = True Then Exit Sub
    i = 1
    datcuentas.Recordset.MoveFirst
    Do While Not datcuentas.Recordset.EOF
        periodos(i) = datcuentas.Recordset.Fields("inicioper")
        finperiodos(i) = datcuentas.Recordset.Fields("finper")
        i = i + 1
        datcuentas.Recordset.MoveNext
    Loop
    
   For x = 1 To i - 1
    datcuentas.RecordSource = "select cuentas.* from cuentas where empre = " & login.empresaact & " and inicioper = '" & periodos(x) & "' and [Cod Contable] = " & Text8(0).Text & "   "
    datcuentas.Refresh
    If actualizar = "graba" Then
     If datcuentas.Recordset.EOF = True Then
        datcuentas.Recordset.AddNew
        datcuentas.Recordset.Fields("cod contable") = Text8(0).Text
        datcuentas.Recordset.Fields("nombre cuenta") = Text8(1).Text
        datcuentas.Recordset.Fields("imp") = Text8(2).Text
        datcuentas.Recordset.Fields("id cuenta") = cuentanumeros
        datcuentas.Recordset.Fields("idcuenta") = cuentapuntos
        datcuentas.Recordset.Fields("empre") = login.empresaact
        datcuentas.Recordset.Fields("inicioper") = periodos(x)
        datcuentas.Recordset.Fields("finper") = finperiodos(x)
        datcuentas.Recordset.UpdateBatch adAffectCurrent
     Else
        datcuentas.Recordset.Fields("cod contable") = Text8(0).Text
        datcuentas.Recordset.Fields("nombre cuenta") = Text8(1).Text
        datcuentas.Recordset.Fields("imp") = Text8(2).Text
        datcuentas.Recordset.Fields("id cuenta") = cuentanumeros
        datcuentas.Recordset.Fields("idcuenta") = cuentapuntos
        datcuentas.Recordset.UpdateBatch adAffectCurrent
     End If
    Else
        If datcuentas.Recordset.EOF = False Then
            datcuentas.Recordset.Delete adAffectCurrent
        End If
    End If
   Next x
    
    
End Sub

Private Sub Cancelar_Click()

  
  datPrimaryRS.RecordSource = "select cuentas.* from cuentas WHERE inicioper = '" & login.iper & "' and cuentas.empre = " & login.empresaact & " ORDER BY IDCUENTA"
  datPrimaryRS.Refresh
    
  detalle = ""
  Call acttree_Click
  Text8(0).SetFocus

End Sub

Private Sub actualiza_Click()
  Call acttree_Click
End Sub

Private Sub agregamismonivel_Click()
On Error GoTo errortagrega
Dim numcue1(1000) As String

        If detalle.Text = "" Then
            mensa = MsgBox("Debe Ubicarse en que lugar dara de alta la cuenta", vbCritical, "Error")
            Command1(0).SetFocus
            Exit Sub
        End If
        
        
        If Text8(2).Text = "S" Then
            For x = pos - 1 To 1 Step -1
               datPrimaryRS.Recordset.AbsolutePosition = x
               Call llena_Click
               For Y = 5 To 0 Step -1
                    If Text9(Y).Visible = True And Val(Text9(Y).Text) = 0 Then GoTo continua
               Next Y
            Next x
continua:
        End If
    
        Text8(0).Enabled = False
        For x = 0 To 5
            Text9(x).Visible = True
            Text9(x).Enabled = True
            Text9(x).Locked = False
            If Text9(x).Text = "" Then Text9(x).Visible = False
        Next x
        
        For x = 0 To 5
            If Text9(x).Visible = True And Val(Text9(x).Text) = 0 Then
                For Y = x - 1 To 0 Step -1
                    Text9(Y).Enabled = False
                Next Y
                For Y = x + 1 To 5
                    Text9(Y).Enabled = False
                Next Y
                Exit For
            End If
        Next x
        
        
        patron = datPrimaryRS.Recordset.Fields("idcuenta")
        If x = 6 Then
            MsgBox "NO se puede dar de alta un nuevo nivel, el codigo no lo permite", vbCritical, "Error"
            Exit Sub
        End If
        If x = 5 Then patron1 = Text9(0).Text + "." + Text9(1).Text + "." + Text9(2).Text + "." + Text9(3).Text + "." + Text9(4).Text + "."
        If x = 4 Then patron1 = Text9(0).Text + "." + Text9(1).Text + "." + Text9(2).Text + "." + Text9(3).Text + "."
        If x = 3 Then patron1 = Text9(0).Text + "." + Text9(1).Text + "." + Text9(2).Text + "."
        If x = 2 Then patron1 = Text9(0).Text + "." + Text9(1).Text + "."
        If x = 1 Then patron1 = Text9(0).Text + "."
        
        lonpatron = Len(patron1) + 1
        datPrimaryRS.Recordset.MoveNext
        numcue3 = ""
        i = 1
        Do While Not datPrimaryRS.Recordset.EOF
            numcue = datPrimaryRS.Recordset.Fields("idcuenta")
            If x = 1 Then
                numcue1(i) = Mid(numcue, lonpatron, c2)
                numcue2 = Mid(numcue, lonpatron - 1 - c1, c1)
            End If
            If x = 2 Then
                numcue1(i) = Mid(numcue, lonpatron, c3)
                numcue2 = Mid(numcue, lonpatron - 1 - c2, c2)
            End If
            If x = 3 Then
                numcue1(i) = Mid(numcue, lonpatron, c4)
                numcue2 = Mid(numcue, lonpatron - 1 - c3, c3)
            End If
            If x = 4 Then
                numcue1(i) = Mid(numcue, lonpatron, c5)
                numcue2 = Mid(numcue, lonpatron - 1 - c4, c4)
            End If
            If numcue3 <> "" And numcue2 <> numcue3 Then GoTo endloop
            numcue3 = numcue2
            datPrimaryRS.Recordset.MoveNext
            i = i + 1
        Loop
endloop:
        codcont1 = 0
        
        datverifica.RecordSource = "SELECT TOP 100 PERCENT MAX([Cod Contable]) AS [Cod Contable], empre, inicioper, LEFT([Cod Contable], 1) AS cod From dbo.cuentas GROUP BY empre, inicioper, LEFT([Cod Contable], 1) HAVING empre = " & login.empresaact & " and (LEFT([Cod Contable],1)) = '" & Text9(0).Text & "' and  inicioper = '" & login.iper & "'  ORDER BY MAX([Cod Contable])"
        datverifica.Refresh
            
            patron = datverifica.Recordset.Fields("Cod Contable")
            codcont = Mid(patron, 2, Len(patron) - 1)
            codcont1 = Val(codcont)
    

        Text8(0).Text = Val(Text9(0).Text + Right(Str(codcont1 + 1), Len(Str(codcont1))))
        For x = 0 To 5
            If Text9(x).Enabled = True Then
                   If numcue1(i - 1) = "" Then numcue1(i - 1) = Right("0000000", nivelag(x + 1))
                   num = Val(numcue1(i - 1)) + 1
                   num1 = Right("0000000000" + Right(Str(num), Len(Str(num)) - 1), Len(numcue1(i - 1)))
                   Text9(x).Text = num1
                   Exit For
            End If
        Next x
        datPrimaryRS.Recordset.AddNew
        Text8(1).Text = ""
        Text8(2).Text = ""
        
        Text8(1).SetFocus
Exit Sub

errortagrega:
    mensa = MsgBox("Debe Ubicarse en que lugar dara de alta la cuenta", vbCritical, "Error")
        
        
        
        
End Sub


Private Sub agregamismonivel_LostFocus()
 borrar.Enabled = True
End Sub

Private Sub borrar_Click()
On Error GoTo DeleteErr

Respuesta = MsgBox("ESTA POR BORRAR UNA CUENTA, ESTA SEGURO?", vbYesNo, "Atención")
If Respuesta = vbYes Then

    If hijos(pos) > 0 Then
        mensa = MsgBox("No puede eliminar esta cuenta, elimine primero las Subcuentas", vbCritical, "Error")
        Exit Sub
    End If
     
     If login.planunificado = "S" Then
        datverifica.RecordSource = "select controlcuentas.* from controlcuentas where empresa = " & login.empresaact & " and idcuenta = " & Text8(0).Text & ""
        datverifica.Refresh
        If datverifica.Recordset.EOF = False Then
            MsgBox "No se puede borrar esta cuenta, ya que posee movimientos", vbCritical, "Error"
            Exit Sub
        End If
     Else
        datverifica.RecordSource = "select controlcuentas.* from controlcuentas where empresa = " & login.empresaact & " and perinicial = '" & login.iper & "' and idcuenta = " & Text8(0).Text & " "
        datverifica.Refresh
        If datverifica.Recordset.EOF = False Then
            MsgBox "No se puede borrar esta cuenta, ya que posee movimientos", vbCritical, "Error"
            Exit Sub
        End If
     End If
           
    
       Inicio.datauditoria.ConnectionString = login.conexiontotal
    
   Inicio.datauditoria.RecordSource = "select auditoria.* from auditoria"
   Inicio.datauditoria.Refresh
    
   Inicio.datauditoria.Recordset.AddNew
   Inicio.datauditoria.Recordset.Fields("fecha") = Date
   Inicio.datauditoria.Recordset.Fields("hora") = Str(Time)
   Inicio.datauditoria.Recordset.Fields("ventana") = "Plan de Cuentas"
   Inicio.datauditoria.Recordset.Fields("accion") = "Eliminación cuenta:" + Str(Text8(0))
   Inicio.datauditoria.Recordset.Fields("usuario") = login.usuarioactivo
   Inicio.datauditoria.Recordset.Fields("empresa") = login.empresaact
   Inicio.datauditoria.Recordset.UpdateBatch adAffectCurrent
    
     If login.planunificado = "S" Then
        actualizar = "elimina"
        Call actualizaplan_Click
     End If
    
    
    
    datPrimaryRS.Recordset.Delete adAffectCurrent


    detalle.Text = ""
    Call Cancelar_Click
Else
    Exit Sub
End If

 
 Exit Sub
DeleteErr:
  MsgBox "No se pudo borrar, pulse el boton ´Grabar´ e intente eliminar el registro nuevamente"

End Sub

Private Sub cerrar_Click()

    Unload Me

End Sub

Private Sub codcuenta_Change()
On Error GoTo fuera

If codcuenta <> "" Then Cuenta = Str(codcuenta.Text)
n1 = 2: c1 = Val(car1)
n2 = n1 + c1: c2 = Val(car2)
n3 = n2 + c2: c3 = Val(car3)
n4 = n3 + c3: c4 = Val(car4)
n5 = n4 + c4: c5 = Val(car5)

ni1 = Mid(Cuenta, n1, c1)
ni2 = Mid(Cuenta, n2, c2)
ni3 = Mid(Cuenta, n3, c3)
ni4 = Mid(Cuenta, n4, c4)
ni5 = Mid(Cuenta, n5, c5)


If codabrev = "" Then
    If max1.Text <> "" Then
        maxim1 = Right(max1.Text, Len(max1.Text) - 1)
    End If
    If max2.Text <> "" Then
        maxim2 = Right(max2.Text, Len(max2.Text) - 1)
    End If
    If max3.Text <> "" Then
        maxim3 = Right(max3.Text, Len(max3.Text) - 1)
    End If
    If max4.Text <> "" Then
        maxim4 = Right(max4.Text, Len(max4.Text) - 1)
    End If
    If max5.Text <> "" Then
        maxim5 = Right(max5.Text, Len(max5.Text) - 1)
    End If
    If max6.Text <> "" Then
        maxim6 = Right(max6.Text, Len(max6.Text) - 1)
    End If
    If max1.Text <> "" And ni1 = "1" Then codabrev = Val(maxim1) + 1
    If max2.Text <> "" And ni1 = "2" Then codabrev = Val(maxim2) + 1
    If max3.Text <> "" And ni1 = "3" Then codabrev = Val(maxim3) + 1
    If max4.Text <> "" And ni1 = "4" Then codabrev = Val(maxim4) + 1
    If max5.Text <> "" And ni1 = "5" Then codabrev = Val(maxim5) + 1
    If max6.Text <> "" And ni1 = "6" Then codabrev = Val(maxim6) + 1
    If codabrev <> "" Then
        codigoabreviado = ni1 + Right(Str(codabrev), Len(Str(codabrev)) - 1)
        codabrev = Val(codigoabreviado)
    End If
End If
    
  If max1.Text = "" And ni1 = "1" Then codabrev = 10
  If max2.Text = "" And ni1 = "2" Then codabrev = 20
  If max3.Text = "" And ni1 = "3" Then codabrev = 30
  If max4.Text = "" And ni1 = "4" Then codabrev = 40
  If max5.Text = "" And ni1 = "5" Then codabrev = 50
  If max6.Text = "" And ni1 = "6" Then codabrev = 60

  If bandera = 0 Then idcuenta0.Text = ni1 + "." + ni2 + "." + ni3 + "." + ni4 + "." + ni5
 
fuera:
End Sub


Private Sub codcuenta_LostFocus()
idcuentacombo = codcuenta
combocuenta = codabrev
End Sub


Private Sub combocuenta_Change()
On Error GoTo fuera

If combocuenta.SelectedItem <> 0 Then datPrimaryRS.Recordset.Bookmark = combocuenta.SelectedItem

fuera:
End Sub


Private Sub combocuenta_Validate(Cancel As Boolean)
combocuenta.Refresh
End Sub




Private Sub DataCombo1_Change()
On Error GoTo fuera

        niveles.Recordset.Bookmark = DataCombo1.SelectedItem
    
fuera:
End Sub



Private Sub Command1_GotFocus(Index As Integer)

    If Index = 0 Then Text8(0).SetFocus

End Sub

Private Sub Command2_Click()
On Error Resume Next
        For x = 0 To 5
            If Text9(x).Text = "" Then
                Text9(x - 1).Locked = False
                Exit For
            End If
        Next x
End Sub

Private Sub DataGrid1_Click()
 borrar.Enabled = True
End Sub


Private Sub Form_Load()
Aplicar_skin Me

frmCuentas.Top = 0
frmCuentas.Left = 0

If login.plancuentasaltas = "N" Then
    agregamismonivel.Enabled = False
Else
    agregamismonivel.Enabled = True
End If
If login.plancuentasmodi = "N" Then
    grabar.Enabled = False
Else
    grabar.Enabled = True
End If
If login.plancuentasbajas = "N" Then
    borrar.Enabled = False
Else
    borrar.Enabled = True
End If

  bandera = 1
     
  Option3.Value = True
  Option5.Value = True
  datPrimaryRS.ConnectionString = login.conexiontotal
  criterio.ConnectionString = login.conexiontotal
  niveles.ConnectionString = login.conexiontotal
  datverifica.ConnectionString = login.conexiontotal
  datcuentas.ConnectionString = login.conexiontotal
  
  
  datPrimaryRS.RecordSource = "select cuentas.* from cuentas WHERE inicioper = '" & login.iper & "' and cuentas.empre = " & login.empresaact & " ORDER BY IDCUENTA"
  datPrimaryRS.Refresh

  niveles.RecordSource = "select niveles.* from niveles Where empre = " & login.empresaact & " and inicioper = '" & login.iper & "' "
  niveles.Refresh
       
  Call acttree_Click
  

End Sub

Private Sub Form_Resize()
ubic = (Forms.Count - 3) * 500
If Forms.Count > 3 Then
    frmCuentas.Top = frmCuentas.Top + ubic
    frmCuentas.Left = frmCuentas.Left + ubic
End If
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

Private Sub datPrimaryRS_Movelete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  'Esto mostrará la posición de registro actual para este Recordset
  datPrimaryRS.Caption = "Registro: " & CStr(datPrimaryRS.Recordset.AbsolutePosition)
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

Private Sub cmdAdd_Click()
  On Error GoTo AddErr
  datPrimaryRS.Recordset.AddNew



  Exit Sub
AddErr:
  MsgBox Err.Description
End Sub

Private Sub cmdDelete_Click()
  On Error GoTo DeleteErr
  With datPrimaryRS.Recordset
    .Delete
    .MoveNext
    If .EOF Then .MoveLast
  End With
  Exit Sub
DeleteErr:
  MsgBox Err.Description
End Sub


Private Sub cmdUpdate_Click()
  On Error GoTo UpdateErr



  datPrimaryRS.Recordset.UpdateBatch adAffectAll
  Exit Sub
UpdateErr:
  MsgBox Err.Description
End Sub

Private Sub cmdClose_Click()
  Unload Me
End Sub


Private Sub grabar_Click()
On Error GoTo grabErr

For x = 0 To 5
    Text9(x).Locked = True
Next x


  If Text8(2).Text = "N" Then
     If login.planunificado = "S" Then
        datverifica.RecordSource = "select controlcuentas.* from controlcuentas where empresa = " & login.empresaact & " and idcuenta = " & Text8(0).Text & ""
        datverifica.Refresh
        If datverifica.Recordset.EOF = False Then
            MsgBox "No se puede sacar la imputacion a esta cuenta, ya que posee movimientos", vbCritical, "Error"
            Call Cancelar_Click
            Exit Sub
        End If
     Else
        datverifica.RecordSource = "select controlcuentas.* from controlcuentas where empresa = " & login.empresaact & " and perinicial = '" & login.iper & "' and idcuenta = " & Text8(0).Text & " "
        datverifica.Refresh
        If datverifica.Recordset.EOF = False Then
            MsgBox "No se puede sacar la imputacion a esta cuenta, ya que posee movimientos", vbCritical, "Error"
            Call Cancelar_Click
            Exit Sub
        End If
     End If
  End If


  If Text8(1).Text = "" Then
    mensa = MsgBox("Debe ingresar un nombre valido para esta cuenta", vbCritical, "Error")
    Text8(1).SetFocus
    Exit Sub
  End If

  If Text8(2).Text = "S" Or Text8(2).Text = "N" Then
    datPrimaryRS.Recordset.Fields("imp") = Text8(2).Text
  Else
    mensa = MsgBox("Debe ingresar correctamente el campo de imputable (S o N)", vbCritical, "Error")
    Text8(2).SetFocus
    Exit Sub
  End If
  

  cuentapuntos = ""
  cuentanumeros = ""
  For x = 0 To 5
    If Text9(x).Visible = True Then
        cuentapuntos = cuentapuntos + Text9(x).Text + "."
        cuentanumeros = cuentanumeros + Text9(x).Text
    End If
  Next x
  cuentapuntos = Left(cuentapuntos, Len(cuentapuntos) - 1)
  If Right(cuentapuntos, 1) = "." Then
    cuentapuntos = Left(cuentapuntos, Len(cuentapuntos) - 1)
  End If
  cuentanumeros = Val(cuentanumeros)
  
  memoriza = Text8(0).Text
  datPrimaryRS.Recordset.Fields("cod contable") = Text8(0).Text
  datPrimaryRS.Recordset.Fields("nombre cuenta") = Text8(1).Text
  datPrimaryRS.Recordset.Fields("imp") = Text8(2).Text
  datPrimaryRS.Recordset.Fields("id cuenta") = cuentanumeros
  datPrimaryRS.Recordset.Fields("idcuenta") = cuentapuntos
  datPrimaryRS.Recordset.Fields("empre") = login.empresaact
  datPrimaryRS.Recordset.Fields("inicioper") = login.iper
  datPrimaryRS.Recordset.Fields("finper") = login.fper
  datPrimaryRS.Recordset.UpdateBatch adAffectCurrent
  
  If login.planunificado = "S" Then
    actualizar = "graba"
    Call actualizaplan_Click
  End If
  
   Inicio.datauditoria.ConnectionString = login.conexiontotal
    
   Inicio.datauditoria.RecordSource = "select auditoria.* from auditoria"
   Inicio.datauditoria.Refresh
    
   Inicio.datauditoria.Recordset.AddNew
   Inicio.datauditoria.Recordset.Fields("fecha") = Date
   Inicio.datauditoria.Recordset.Fields("hora") = Str(Time)
   Inicio.datauditoria.Recordset.Fields("ventana") = "Plan de Cuentas"
   Inicio.datauditoria.Recordset.Fields("accion") = "Alta/Modificacion cuenta:" + Str(Text8(0).Text)
   Inicio.datauditoria.Recordset.Fields("usuario") = login.usuarioactivo
   Inicio.datauditoria.Recordset.Fields("empresa") = login.empresaact
   Inicio.datauditoria.Recordset.UpdateBatch adAffectCurrent
  

  Call Cancelar_Click
  Text8(0).Text = memoriza
  SendKeys "{ENTER}", False
  detalle.Text = " "
  
  Exit Sub
grabErr:
     MsgBox Err.Description

    

End Sub



Private Sub idcuentacombo_Change()
On Error GoTo fuera

If idcuentacombo.SelectedItem <> 0 Then datPrimaryRS.Recordset.Bookmark = idcuentacombo.SelectedItem

fuera:
End Sub



Private Sub imprimir_Click()
On Error GoTo fuera

Dim ruta As String

ruta = "\Empresa" + Right(Str(login.empresaact), Len(Str(login.empresaact)) - 1)

  
  criterio.RecordSource = "select empreactiva.* from empreactiva"
  criterio.Refresh

    criterio.Recordset.Fields(1) = login.iper
    criterio.Recordset.Fields(2) = login.fper
    criterio.Recordset.UpdateBatch adAffectCurrent
    criterio.Refresh

If Option1.Value = True Then
    reporte.SQL = "SELECT plancuentas.empre, plancuentas.Nombrecuenta, plancuentas.imp, plancuentas.niv1, plancuentas.niv2, plancuentas.razonsocial, plancuentas.idcuenta, plancuentas.codcontable, plancuentas.inicioper, plancuentas.finper FROM contablesql.dbo.plancuentas plancuentas WHERE plancuentas.inicioper = '" & login.iper & "' and plancuentas.empre = " & login.empresaact & " ORDER BY plancuentas.nombrecuenta ASC"
Else
    reporte.SQL = "SELECT plancuentas.empre, plancuentas.Nombrecuenta, plancuentas.imp, plancuentas.niv1, plancuentas.niv2, plancuentas.razonsocial, plancuentas.idcuenta, plancuentas.codcontable, plancuentas.inicioper, plancuentas.finper FROM contablesql.dbo.plancuentas plancuentas WHERE plancuentas.inicioper = '" & login.iper & "' and plancuentas.empre = " & login.empresaact & " ORDER BY plancuentas.idcuenta ASC"
End If

If Option5.Value = True Then filtro = 3
If Option4.Value = True Then filtro = 2
If Option2.Value = True Then filtro = 1


tabla = reporte.SQL

With crystalreporte
    .ReportFileName = App.Path & ruta + "\Plancuentas.rpt"
    .Connect = login.conexionreporte
    .Formulas(0) = "filtrar=""" & filtro & """"
       
    .DiscardSavedData = True
    .RetrieveDataFiles
    .ReportSource = 0
    .SQLQuery = tabla
    .Destination = crptToWindow
    .PrintFileType = crptCrystal
    .WindowState = crptMaximized
    .Action = 1
      
End With
Exit Sub
fuera:
    Frame7.Visible = True
    KeyAscii = 7
    Frame7.Caption = "Error"
    Text1.Text = "No se puede Imprimir"
    mensage.SetFocus
    
    

End Sub

Private Sub imputable_KeyPress(KeyAscii As Integer)
On Error GoTo fuera

    If KeyAscii = 13 Then
        KeyAscii = 0
        grabar.SetFocus
    End If
    
fuera:
End Sub


Private Sub menureenumerar_Click(Index As Integer)
On Error GoTo fuera


Respuesta = MsgBox("ESTO SOLO PUEDE REALIZARCE SI NO TIENE MOVIMIENTOS CARGADOS, DEBE ESTAR TOTALMENTE SEGURO PARA REALIZAR ESTA TAREA, REALIZA LA ACCION ? ", vbYesNo, "!! Atención !!")
If Respuesta = vbYes Then
    datPrimaryRS.Recordset.MoveFirst
    orden = 0
paso1:
    codabrev0 = Mid(Str(codcuenta.Text), 2, 1)
    codabrev1 = Str(orden)
    codabrev2 = codabrev0 + Right(codabrev1, Len(codabrev1) - 1)
    codabrev = Str(codabrev2)
    orden = orden + 1
    datPrimaryRS.Recordset.MoveNext
    If datPrimaryRS.Recordset.EOF = True Then GoTo paso2
    If Mid(Str(codcuenta.Text), 2, 1) <> codabrev0 Then orden = 0
    GoTo paso1
paso2:
    mensa = MsgBox("Codigo Abreviado Reenumerado", vbDefaultButton1)
    datPrimaryRS.Recordset.MoveLast
    Exit Sub
Else
    Exit Sub
End If
    
fuera:
End Sub



Private Sub llena_Click()
        Cuenta = Str(datPrimaryRS.Recordset.Fields("id cuenta"))
        n1 = 2: c1 = Val(car1)
        n2 = n1 + c1: c2 = Val(car2)
        n3 = n2 + c2: c3 = Val(car3)
        n4 = n3 + c3: c4 = Val(car4)
        n5 = n4 + c4: c5 = Val(car5)

        Text9(0).Text = Mid(Cuenta, n1, c1)
        Text9(1).Text = Mid(Cuenta, n2, c2)
        Text9(2).Text = Mid(Cuenta, n3, c3)
        Text9(3).Text = Mid(Cuenta, n4, c4)
        Text9(4).Text = Mid(Cuenta, n5, c5)
        
              
        Text8(0).Text = datPrimaryRS.Recordset.Fields("cod contable")
        Text8(1).Text = datPrimaryRS.Recordset.Fields("Nombre cuenta")
        Text8(2).Text = datPrimaryRS.Recordset.Fields("imp")
        
End Sub



Private Sub nuevo_Click()
  On Error GoTo AddErr
  
  
  datPrimaryRS.Recordset.AddNew
  bandera = 0
  codabrev = ""
  codcuenta.SetFocus

    
  Exit Sub
AddErr:
  MsgBox Err.Description
End Sub

Private Sub Nuevo_KeyPress(KeyAscii As Integer)
On Error GoTo fuera

    If KeyAscii = 13 Then
        KeyAscii = 0
        codcuenta.SetFocus
    End If

fuera:
End Sub




Private Sub mensage_Click()

 Frame7.Visible = False
 

End Sub

Private Sub reordenar_Click()
On Error GoTo fin


contador = 0
    datPrimaryRS.Recordset.MoveFirst
paso1:
    codcuenta0 = datPrimaryRS.Recordset.Fields("idcuenta")
    If datPrimaryRS.Recordset.EOF = True Then GoTo fin
    digito1 = Left(codcuenta0, 1)
    contador = contador + 1
    If digito0 <> digito1 Then contador = 0
    numero = Str(digito1) + Right(Str(contador), Len(Str(contador)) - 1)
    codabrev = Val(numero)
    datPrimaryRS.Recordset.Fields("cod contable") = codabrev
    datPrimaryRS.Recordset.UpdateBatch adAffectCurrent
    If datPrimaryRS.Recordset.EOF = False Then
        digito0 = Left(codcuenta0, 1)
        datPrimaryRS.Recordset.MoveNext
        GoTo paso1
    Else
        GoTo fin
    End If
    
fin:
    Unload Me
End Sub

Private Sub reordenar_LostFocus()

        reordenar.Visible = False
    
End Sub

Private Sub salir_Click()

    Unload Me

End Sub

Private Sub Text8_Change(Index As Integer)

    If Index = 2 Then Text8(2).Text = UCase(Text8(2).Text)
        

End Sub

Private Sub Text8_GotFocus(Index As Integer)

    If ventana.menu = 3 And Index = 0 Then
        ventana.menu = 0
        Text8(0).Text = lista_cuentas.cuentacont
    End If

End Sub

Private Sub Text8_KeyPress(Index As Integer, KeyAscii As Integer)
On Error Resume Next

    If KeyAscii = 13 And Index = 0 Then
        KeyAscii = 0
        
        For x = 1 To datPrimaryRS.Recordset.RecordCount - 1
            TreeView1.Nodes(x).Expanded = False
            TreeView1.Nodes(x).ForeColor = QBColor(0)
        Next x
                
                
        datPrimaryRS.Recordset.Filter = "inicioper = " & login.iper & " and empre = " & login.empresaact & " and [cod contable] = " & Text8(0).Text & ""
        If datPrimaryRS.Recordset.EOF = True Then
            Text8(1).Text = ""
            Text8(2).Text = ""
            For Y = 0 To 5
                Text9(Y) = ""
            Next Y
            MsgBox "Cuenta inexistente", vbInformation, "Cuenta"
            Exit Sub
       End If
            
        
        Cuenta = Str(datPrimaryRS.Recordset.Fields("id cuenta"))
        n1 = 2: c1 = Val(car1)
        n2 = n1 + c1: c2 = Val(car2)
        n3 = n2 + c2: c3 = Val(car3)
        n4 = n3 + c3: c4 = Val(car4)
        n5 = n4 + c4: c5 = Val(car5)

        Text9(0).Text = Mid(Cuenta, n1, c1)
        Text9(1).Text = Mid(Cuenta, n2, c2)
        Text9(2).Text = Mid(Cuenta, n3, c3)
        Text9(3).Text = Mid(Cuenta, n4, c4)
        Text9(4).Text = Mid(Cuenta, n5, c5)
        For Y = 0 To 5
            If Text9(Y).Text = "" Then
                Text9(Y).Visible = False
            Else
                Text9(Y).Visible = True
            End If
        Next Y
              
        For x = cabecera(Val(Text9(0).Text)) To cuentaabrev(Text8(0).Text)
             TreeView1.Nodes(x).Expanded = True
        Next x
        TreeView1.Nodes(x - 1).ForeColor = QBColor(12)
        pos = TreeView1.Nodes(x - 1).Index
  Rem      Call TreeView1_NodeClick(TreeView1.Nodes(X - 1))
        Text8(1).Text = datPrimaryRS.Recordset.Fields("Nombre cuenta")
        Text8(2).Text = datPrimaryRS.Recordset.Fields("imp")
        datPrimaryRS.Recordset.Filter = "inicioper = " & login.iper & " and empre = " & login.empresaact & " "
        datPrimaryRS.Recordset.AbsolutePosition = pos
        Text8(1).SetFocus
    End If
    
    If KeyAscii = 13 Then
            KeyAscii = 0
            SendKeys "{tab}", False
    End If
    detalle = ""
End Sub

Private Sub Text8_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)

    If KeyCode = 123 Then reordenar.Visible = True

    If KeyCode = 114 And Index = 0 Then
        ventana.menu = 3
        lista_cuentas.Show
    End If

End Sub



Private Sub Text9_LostFocus(Index As Integer)

    If Len(Text9(Index).Text) <> nivelag(Index + 1) Then
        MsgBox "La Cantidad de digitos en el codigo esta mail, Verifique", vbCritical, "Error"
        Text9(Index).SetFocus
    End If

End Sub

Private Sub TreeView1_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        Text8(1).SetFocus
    End If

End Sub

Private Sub TreeView1_NodeClick(ByVal Node As Node)
On Error Resume Next

    bandera = 0
    
    
    
   pos = posicion(Node.Index)

   datPrimaryRS.Recordset.AbsolutePosition = pos
   Call llena_Click

   nombrecuenta = Node.Text

   nivelanterior = Node.Key
   canthijos = Node.Children
   ruta = Node.FullPath
   hijos(pos) = canthijos
   detalle.Text = idcuenta0.Text + "=>" + ruta
   
   nivel = 0
   For x = 1 To Len(ruta)
          letra = Mid(ruta, x, 1)
          If letra = "\" Then nivel = nivel + 1
   Next x

ni1 = Mid(Cuenta, n1, c1)
ni2 = Mid(Cuenta, n2, c2)
ni3 = Mid(Cuenta, n3, c3)
ni4 = Mid(Cuenta, n4, c4)
ni5 = Mid(Cuenta, n5, c5)

    
 Exit Sub
errortree:
    mensa = MsgBox("Error al dar de alta este nivel", vbCritical, "Error")
          
End Sub
