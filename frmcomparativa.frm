VERSION 5.00
Object = "{F6125AB1-8AB1-11CE-A77F-08002B2F4E98}#2.0#0"; "MSRDC20.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{D18BBD1F-82BB-4385-BED3-E9D31A3E361E}#1.0#0"; "KewlButtonz.ocx"
Object = "{50BF2256-701F-46F2-8ADB-2202CE6922BC}#1.0#0"; "KlexGrid.ocx"
Begin VB.Form frmcomparativa 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Comparativa de Precios"
   ClientHeight    =   8610
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   18165
   Icon            =   "frmcomparativa.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8610
   ScaleWidth      =   18165
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00C0FFFF&
      Height          =   315
      Left            =   120
      OLEDragMode     =   1  'Automatic
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   0
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton UM 
      Caption         =   "UM"
      Height          =   315
      Left            =   4200
      TabIndex        =   37
      Top             =   2160
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton calcula 
      Caption         =   "calcula"
      Height          =   315
      Left            =   4560
      TabIndex        =   36
      Top             =   2400
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton ubicatextogrilla 
      Caption         =   "ubicatextogrilla"
      Height          =   315
      Left            =   12240
      TabIndex        =   35
      Top             =   240
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1560
      MultiLine       =   -1  'True
      TabIndex        =   11
      Text            =   "frmcomparativa.frx":0442
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Frame Frame2 
      Caption         =   "Nro. Presupuesto"
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
      Height          =   975
      Left            =   240
      TabIndex        =   25
      Top             =   7440
      Width           =   4095
      Begin VB.TextBox Text17 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   240
         TabIndex        =   0
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton cargapresupuesto 
         Caption         =   "cargapresupuesto"
         Height          =   315
         Left            =   2040
         TabIndex        =   48
         Top             =   120
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.TextBox Text18 
         Height          =   285
         Left            =   1680
         TabIndex        =   47
         Text            =   "Text18"
         Top             =   600
         Visible         =   0   'False
         Width           =   615
      End
      Begin KewlButtonz.KewlButtons buscar 
         Height          =   495
         Left            =   2640
         TabIndex        =   26
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   873
         BTYPE           =   14
         TX              =   "&Buscar"
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
         BCOL            =   -2147483645
         BCOLO           =   -2147483645
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmcomparativa.frx":0448
         PICN            =   "frmcomparativa.frx":0464
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
   Begin MSAdodcLib.Adodc datvendedor 
      Height          =   330
      Left            =   3960
      Top             =   6960
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
   Begin MSAdodcLib.Adodc datcliente 
      Height          =   330
      Left            =   2880
      Top             =   6960
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
   Begin MSAdodcLib.Adodc datmovimientos 
      Height          =   330
      Left            =   1800
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
   Begin MSAdodcLib.Adodc datproductos 
      Height          =   330
      Left            =   600
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
      Height          =   975
      Left            =   4560
      TabIndex        =   17
      Top             =   7440
      Width           =   4815
      Begin KewlButtonz.KewlButtons salir 
         Height          =   495
         Left            =   3480
         TabIndex        =   16
         Top             =   360
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   873
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
         BCOL            =   -2147483645
         BCOLO           =   -2147483645
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmcomparativa.frx":09F6
         PICN            =   "frmcomparativa.frx":0A12
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
         Height          =   495
         Left            =   240
         TabIndex        =   14
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   873
         BTYPE           =   14
         TX              =   "&Grabar (F10)"
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
         BCOL            =   -2147483645
         BCOLO           =   -2147483645
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmcomparativa.frx":155C
         PICN            =   "frmcomparativa.frx":1578
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
         Height          =   495
         Left            =   2160
         TabIndex        =   15
         Top             =   360
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   873
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
         BCOL            =   -2147483645
         BCOLO           =   -2147483645
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmcomparativa.frx":2FFA
         PICN            =   "frmcomparativa.frx":3016
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
   Begin VB.PictureBox Picture1 
      Height          =   8415
      Left            =   120
      ScaleHeight     =   8355
      ScaleWidth      =   17715
      TabIndex        =   19
      Top             =   120
      Width           =   17775
      Begin VB.CommandButton Command7 
         Caption         =   "&Historial de Venta"
         Height          =   375
         Left            =   5160
         TabIndex        =   51
         Top             =   2280
         Width           =   1695
      End
      Begin KewlButtonz.KewlButtons agregaservicios 
         Height          =   375
         Left            =   4560
         TabIndex        =   46
         Top             =   2160
         Visible         =   0   'False
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         BTYPE           =   14
         TX              =   "Agrega Serv&icio (F12)"
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
         BCOL            =   -2147483645
         BCOLO           =   -2147483645
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmcomparativa.frx":3A28
         PICN            =   "frmcomparativa.frx":3A44
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   -1  'True
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   6
         Left            =   9120
         MaxLength       =   300
         TabIndex        =   44
         Top             =   600
         Visible         =   0   'False
         Width           =   4935
      End
      Begin VB.CommandButton imprimepresupuesto 
         Caption         =   "imprimepresupuesto"
         Height          =   315
         Left            =   9720
         TabIndex        =   43
         Top             =   2280
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox Text19 
         Height          =   285
         Left            =   120
         TabIndex        =   42
         Top             =   840
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   11520
         Top             =   2160
      End
      Begin VB.Frame Frame1 
         Caption         =   "Observaciones al Pie"
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
         Left            =   9360
         TabIndex        =   38
         Top             =   7200
         Width           =   8175
         Begin VB.TextBox Text15 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Left            =   240
            MaxLength       =   500
            MultiLine       =   -1  'True
            TabIndex        =   13
            Top             =   240
            Width           =   7695
         End
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   5
         Left            =   2880
         MaxLength       =   300
         TabIndex        =   8
         Top             =   1560
         Width           =   11295
      End
      Begin VB.TextBox Text1 
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
         Height          =   360
         Index           =   3
         Left            =   1440
         Locked          =   -1  'True
         MaxLength       =   300
         MultiLine       =   -1  'True
         TabIndex        =   6
         Top             =   1080
         Width           =   7695
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   1440
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   3
         Top             =   600
         Width           =   5295
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   4
         Left            =   10200
         TabIndex        =   31
         Top             =   120
         Width           =   1935
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1440
         TabIndex        =   27
         Top             =   120
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   10320
         TabIndex        =   18
         Top             =   120
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   4680
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   1
         Top             =   120
         Width           =   3375
      End
      Begin MSDataListLib.DataCombo DataCombo3 
         Bindings        =   "frmcomparativa.frx":3FDE
         Height          =   360
         Left            =   9120
         TabIndex        =   5
         Top             =   600
         Visible         =   0   'False
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   741
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         ListField       =   "tipopago"
         BoundColumn     =   "id"
         Text            =   ""
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
      Begin KewlButtonz.KewlButtons bclientes 
         Height          =   375
         Left            =   6720
         TabIndex        =   4
         ToolTipText     =   "F3-Busca Clientes"
         Top             =   600
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   661
         BTYPE           =   14
         TX              =   ""
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
         BCOL            =   -2147483645
         BCOLO           =   -2147483645
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmcomparativa.frx":3FF8
         PICN            =   "frmcomparativa.frx":4014
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin KewlButtonz.KewlButtons bvendedor 
         Height          =   375
         Left            =   8040
         TabIndex        =   2
         ToolTipText     =   "F3-Busca Clientes"
         Top             =   120
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   661
         BTYPE           =   14
         TX              =   ""
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
         BCOL            =   -2147483645
         BCOLO           =   -2147483645
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmcomparativa.frx":45AE
         PICN            =   "frmcomparativa.frx":45CA
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "frmcomparativa.frx":4B64
         Height          =   375
         Left            =   4680
         TabIndex        =   29
         Top             =   360
         Visible         =   0   'False
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
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
               LCID            =   11274
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
               LCID            =   11274
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
      Begin MSDataGridLib.DataGrid DataGrid2 
         Bindings        =   "frmcomparativa.frx":4B7E
         Height          =   495
         Left            =   1440
         TabIndex        =   30
         Top             =   720
         Visible         =   0   'False
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   873
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
               LCID            =   11274
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
               LCID            =   11274
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
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "frmcomparativa.frx":4B97
         Height          =   360
         Left            =   11160
         TabIndex        =   7
         Top             =   1080
         Visible         =   0   'False
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   741
         _Version        =   393216
         Locked          =   -1  'True
         Style           =   2
         ListField       =   "nombre"
         BoundColumn     =   "id"
         Text            =   ""
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
      Begin KewlButtonz.KewlButtons agregaproducto 
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   2160
         Visible         =   0   'False
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         BTYPE           =   14
         TX              =   "&Agregar Item (F5)"
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
         BCOL            =   -2147483645
         BCOLO           =   -2147483645
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmcomparativa.frx":4BB5
         PICN            =   "frmcomparativa.frx":4BD1
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin KewlButtonz.KewlButtons KewlButtons1 
         Height          =   375
         Left            =   2280
         TabIndex        =   10
         Top             =   2160
         Visible         =   0   'False
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         BTYPE           =   14
         TX              =   "&Eliminar Item"
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
         BCOL            =   -2147483645
         BCOLO           =   -2147483645
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmcomparativa.frx":516B
         PICN            =   "frmcomparativa.frx":5187
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSAdodcLib.Adodc datiibb 
         Height          =   330
         Left            =   240
         Top             =   7800
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
      Begin MSAdodcLib.Adodc datencabezado 
         Height          =   330
         Left            =   9600
         Top             =   7440
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
      Begin KewlButtonz.KewlButtons blanco 
         Height          =   375
         Left            =   14280
         TabIndex        =   39
         TabStop         =   0   'False
         ToolTipText     =   "F3-Busca Clientes"
         Top             =   120
         Visible         =   0   'False
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   661
         BTYPE           =   14
         TX              =   ""
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
         BCOL            =   49152
         BCOLO           =   49152
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmcomparativa.frx":5721
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin KewlButtonz.KewlButtons negro 
         Height          =   375
         Left            =   14280
         TabIndex        =   40
         TabStop         =   0   'False
         ToolTipText     =   "F3-Busca Clientes"
         Top             =   480
         Visible         =   0   'False
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   661
         BTYPE           =   14
         TX              =   ""
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
         BCOL            =   192
         BCOLO           =   192
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmcomparativa.frx":573D
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSAdodcLib.Adodc datitems 
         Height          =   330
         Left            =   9600
         Top             =   7800
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
      Begin MSAdodcLib.Adodc datcola 
         Height          =   330
         Left            =   9600
         Top             =   7080
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
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   "contable"
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "datcola"
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
         Left            =   12120
         Top             =   2160
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowTitle     =   "Libro IVA Compras"
         PrinterCollation=   1
         PrintFileLinesPerPage=   60
      End
      Begin MSRDC.MSRDC reporte 
         Height          =   375
         Left            =   9600
         Top             =   8040
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
      Begin KewlButtonz.KewlButtons KewlButtons2 
         Height          =   495
         Left            =   15360
         TabIndex        =   49
         Top             =   480
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   873
         BTYPE           =   14
         TX              =   "&Previsualizar"
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
         BCOL            =   -2147483645
         BCOLO           =   -2147483645
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmcomparativa.frx":5759
         PICN            =   "frmcomparativa.frx":5775
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Grid.KlexGrid grilla 
         Height          =   4815
         Left            =   120
         TabIndex        =   54
         Top             =   2280
         Width           =   17295
         _ExtentX        =   30506
         _ExtentY        =   8493
         EnterKeyBehaviour=   2
         Editable        =   -1  'True
         BackColorAlternate=   0
         GridLinesFixed  =   2
         AllowUserResizing=   3
         BackColorFixed  =   -2147483626
         Cols            =   26
         FixedCols       =   7
         FocusRect       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GridColorFixed  =   8421504
         MouseIcon       =   "frmcomparativa.frx":5D0F
         Rows            =   300
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label3 
         Caption         =   "$ Prov 1"
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
         Left            =   10200
         TabIndex        =   53
         Top             =   2040
         Width           =   2895
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Subtotal:"
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
         Left            =   10200
         TabIndex        =   52
         Top             =   7680
         Width           =   1935
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente Precio Espscial"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   19
         Left            =   6720
         TabIndex        =   50
         Top             =   2280
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   16
         Left            =   7800
         TabIndex        =   45
         Top             =   600
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "COTIZACION DE VENTA"
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
         Left            =   9720
         TabIndex        =   41
         Top             =   2160
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Nota En Encabezado:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   9
         Left            =   120
         TabIndex        =   34
         Top             =   1560
         Width           =   3015
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Lista de Precio:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   8
         Left            =   9240
         TabIndex        =   33
         Top             =   1080
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Domicilio:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   32
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   120
         TabIndex        =   28
         Top             =   120
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "C.U.I.T:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   9240
         TabIndex        =   24
         Top             =   120
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo Compr.:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   7
         Left            =   8760
         TabIndex        =   23
         Top             =   120
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   120
         TabIndex        =   22
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo de Pago:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   7440
         TabIndex        =   21
         Top             =   600
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Vendedor:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   6
         Left            =   3360
         TabIndex        =   20
         Top             =   120
         Width           =   1335
      End
   End
   Begin MSAdodcLib.Adodc dattipopago 
      Height          =   330
      Left            =   9960
      Top             =   7560
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
   Begin MSAdodcLib.Adodc datparametros 
      Height          =   330
      Left            =   12600
      Top             =   8280
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
   Begin MSAdodcLib.Adodc datlistaprecios 
      Height          =   330
      Left            =   11760
      Top             =   7560
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
   Begin MSAdodcLib.Adodc datum 
      Height          =   330
      Left            =   13800
      Top             =   8160
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
Attribute VB_Name = "frmcomparativa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public xcontroltem As Double
Public modo As String
Public usuariomanual As String
Dim xfaconv(10) As Double
Public xumvta As String
Public xalicuotaiibb As Double
Public xexentoiibb As Double
Public xcalculaiibb As String
Public xcalculatempyp As String
Public xalicuptatempip As String
Public xciudadtem As String
Public xciudadcliente As String
Public tipofac As String
Public presupuestobase As Double
Public xlineasmax As Integer
Public tpago As String
Public previsualizar As Integer
Public xlinesadd As Integer
Public xlimitebonif As Double
Public xidpre As Double
Public xcontrolrow As Integer



Private Sub agregaproducto_Click()

  
        'agregaservicios.Enabled = False
  
        If Text1(0).Text = "" Then
            mensa = MsgBox("Debe ingresar un vendedor", vbCritical, "Error")
            Exit Sub
        End If
        If Text1(1).Text = "" Then
            mensa = MsgBox("Debe ingresar un Cliente", vbCritical, "Error")
            Exit Sub
        End If

     menu = 2
     query = "SELECT top 10 p.ID, p.CODIGO, p.DESCRIPCION, v.DENOMINACION AS proveedor, " & _
                      "CASE i.CODIGO WHEN 'I' THEN 'Internacional' WHEN 'N' THEN 'Nacional' ELSE '' END AS Nacionalidad, u.DETALLE AS rubro, " & _
                      "ROUND(CAST(r.PRECIOCTA AS decimal(14, 3)), 3) AS precio, r.CODPROVEEDOR, t.CODIGO AS marca, " & _
                      "p.CODIGO + p.DESCRIPCION+isnull(v.DENOMINACION,'')+CASE i.CODIGO WHEN 'I' THEN 'Internacional' WHEN 'N' THEN 'Nacional' ELSE '' END + isnull(u.DETALLE,'')+isnull(r.CODPROVEEDOR,'')+isnull(t.CODIGO,'') as concatenado, V_UNIDADMEDIDA_.NOMBRE AS um " & _
                      "FROM V_PRODUCTO_ AS p LEFT OUTER JOIN V_UNIDADMEDIDA_ ON p.UNIDADMEDIDANOLINEAL_ID = V_UNIDADMEDIDA_.ID LEFT OUTER JOIN " & _
                      "V_UD_EZI_PRODUCTOS_ AS r ON p.BOEXTENSION_ID = r.ID LEFT OUTER JOIN " & _
                      "V_PROVEEDOR_ AS v ON r.PROVEEDOR_ID = v.ID LEFT OUTER JOIN " & _
                      "V_ITEMTIPOCLASIFICADOR_ AS i ON r.NACIONALIDAD_ID = i.ID LEFT OUTER JOIN " & _
                      "V_ITEMTIPOCLASIFICADOR_ AS t ON r.MARCA_ID = t.ID LEFT OUTER JOIN " & _
                      "V_RUBRO_ AS u ON p.RUBRO_ID = u.ID " & _
                      "Where (p.ACTIVESTATUS = 0) And (p.TIPOOBJETOESTATICO_ID Is Null) " & _
                      "ORDER BY p.DESCRIPCION"
                      
    
    For X = 1 To xlineasmax + xlinesadd
        grilla.Col = 1
        grilla.Row = X
'        Call ubicatextogrilla_Click
        If grilla.Text = "" Then
           xfila = X
           lista_productos_colon.Show
           Exit For
        End If
    Next X
    
    

End Sub

Private Sub agregaservicios_Click()

'If agregaservicios.Enabled = False Then Exit Sub

'agregaproducto.Enabled = False
For X = 1 To xlineasmax + xlinesadd
        grilla.Col = 1
        grilla.Row = X
        'Call ubicatextogrilla_Click
        If grilla.Text = "" Then
           xfila = X
           grilla.Row = xfila
           grilla.Col = 0
           grilla.Text = "{0D3B976B-623B-472F-8308-95665D50263E}" ' Servicio MAN en Calipso
           grilla.Col = 1
           grilla.Text = "MAN"
           grilla.Col = 2
           grilla.Text = "Mantenimiento a Terceros"
           grilla.Col = 3
           grilla.Text = 1
           grilla.Col = 4
           grilla.Text = "Unidad"

           grilla.Col = 5
           grilla.Text = Format(0, "###,##0.00")
           grilla.Col = 6
           grilla.Text = Format(0, "###,##0.00")
           grilla.Col = 8
           grilla.Text = Format(0, "###,##0.00")
           grilla.Col = 9
           grilla.Text = Format(0, "###,##0.00")
           grilla.Col = 11
           grilla.Text = 1
           grilla.Col = 12
           grilla.Text = 1.21
           
           Label2.Caption = "PRESUPUESTO DE SERVICIOS A TERCEROS"
           grilla.Col = 3
           grilla.SetFocus
           'Call ubicatextogrilla_Click
           Exit For
        End If
Next X


End Sub

Private Sub bclientes_Click()


   xcodclientefiltra = datparametros.Recordset.Fields("codclientefiltra")
   If xcodclientefiltra = "01" Then xcontrocliente = "88447B8E-14FE-4D60-9622-B22F6C735701"  ' tucuman
   If xcodclientefiltra = "04" Then xcontrocliente = "4234CA46-B2BE-4690-AC6A-F0DE206F94A9"  ' salta
   If xcodclientefiltra = "03" Then xcontrocliente = "AEC7FBAC-63F7-4404-9512-033D0961D9BC"  ' jujuy

If Text1(1).Text <> "" Then
  If Text19.Text = "" Then
    xbusqueda = "%" + Text1(1).Text + "%"
'    xquery1 = "SELECT  ALIAS_0.ID AS ID,  ALIAS_0.CODIGO, ALIAS_3.NOMBRE AS RAZONSOCIAL, ALIAS_3.CUIT,ALIAS_6.CALLE, ISNULL(ALIAS_6.CALLE,'') + '-'+ ISNULL(V_CIUDAD_.NOMBRE,'') + '-'+ ISNULL(V_PROVINCIA_.NOMBRE,'') AS DOMICILIO, ALIAS_0.DENOMINACION,ALIAS_5.NOMBRE AS ZONA, " & _
'              "ALIAS_7.NUMERO AS TELEFONO, ALIAS_8.DIRECCIONELECTRONICA AS MAIL, " & _
'              "V_TIPOPAGO_.NOMBRE AS TP, V_TIPOPAGO_.OBSERVACION as TipoPago, V_TIPOPAGO_.ACTIVESTATUS AS ACTIVESTATUSTP, V_PROVINCIA_.NOMBRE AS Provincia, " & _
'              "V_CIUDAD_.NOMBRE AS Ciudad, ALIAS_0.CODIGO + ' ' + ALIAS_3.CUIT + ' ' + ALIAS_3.NOMBRE AS CONCATENADO, V_EZI_POCICION_IVA_CLIENTES.CODIGO AS IVA, V_TIPOPAGO_.ID AS IDPAGO, ALIAS_0.DOMICILIOFACTURACION_ID as domicilio_id, ALIAS_0.LISTAPRECIO_ID as listaprecio  " & _
'              "FROM V_PERSONA_ AS ALIAS_3 RIGHT OUTER JOIN V_CLIENTE AS ALIAS_0 WITH (READPAST) LEFT OUTER JOIN " & _
'              "V_EZI_POCICION_IVA_CLIENTES ON ALIAS_0.ID = V_EZI_POCICION_IVA_CLIENTES.idcliente LEFT OUTER JOIN " & _
'              "V_TIPOPAGO_ ON ALIAS_0.TIPOPAGO_ID = V_TIPOPAGO_.ID ON ALIAS_3.ID = ALIAS_0.ENTEASOCIADO_ID LEFT OUTER JOIN " & _
'              "V_CIUDAD_ RIGHT OUTER JOIN V_DOMICILIO_ AS ALIAS_6 ON V_CIUDAD_.ID = ALIAS_6.CIUDAD_ID LEFT OUTER JOIN " & _
'              "V_PROVINCIA_ ON ALIAS_6.PROVINCIA_ID = V_PROVINCIA_.ID ON ALIAS_0.DOMICILIOFACTURACION_ID = ALIAS_6.ID LEFT OUTER JOIN " & _
'              "V_TELEFONO_ AS ALIAS_7 ON ALIAS_3.TELEFONOPRINCIPAL_ID = ALIAS_7.ID LEFT OUTER JOIN " & _
'              "V_DIRECCIONELECTRONICA_ AS ALIAS_8 ON ALIAS_3.DIRECELECTRONICAPRINCIPAL_ID = ALIAS_8.ID LEFT OUTER JOIN " & _
'              "V_ZONA_ AS ALIAS_5 ON ALIAS_0.ZONA_ID = ALIAS_5.ID " & _
'              "WHERE     (ALIAS_0.ACTIVESTATUS = 0) AND ALIAS_0.CODIGO + ' ' + ALIAS_3.CUIT + ' ' + ALIAS_3.NOMBRE + ' ' + ALIAS_0.DENOMINACION like '" & xbusqueda & "' order by ALIAS_3.NOMBRE "
              
    xquery1 = "SELECT  ALIAS_0.ID AS ID,  ALIAS_0.CODIGO, ALIAS_3.NOMBRE AS RAZONSOCIAL, ALIAS_3.CUIT,ALIAS_6.CALLE, ISNULL(ALIAS_6.CALLE,'') + '-'+ ISNULL(V_CIUDAD_.NOMBRE,'') + '-'+ ISNULL(V_PROVINCIA_.NOMBRE,'') AS DOMICILIO, ALIAS_0.DENOMINACION,ALIAS_5.NOMBRE AS ZONA, " & _
              "ALIAS_7.NUMERO AS TELEFONO, ALIAS_8.DIRECCIONELECTRONICA AS MAIL, " & _
              "V_TIPOPAGO_.NOMBRE AS TP, V_TIPOPAGO_.OBSERVACION as TipoPago, V_TIPOPAGO_.ACTIVESTATUS AS ACTIVESTATUSTP, V_PROVINCIA_.NOMBRE AS Provincia, " & _
              "V_CIUDAD_.NOMBRE AS Ciudad, ALIAS_0.CODIGO + ' ' + ALIAS_3.CUIT + ' ' + ALIAS_3.NOMBRE AS CONCATENADO, V_EZI_POCICION_IVA_CLIENTES.CODIGO AS IVA, V_TIPOPAGO_.ID AS IDPAGO, ALIAS_0.DOMICILIOFACTURACION_ID as domicilio_id, ALIAS_0.LISTAPRECIO_ID as listaprecio, ALIAS_0.creditomaximo  " & _
              "FROM V_PERSONA AS  ALIAS_3 with (nolock) RIGHT OUTER JOIN V_CLIENTE AS ALIAS_0 WITH (NOLOCK) LEFT OUTER JOIN " & _
              "V_EZI_POCICION_IVA_CLIENTES ON ALIAS_0.ID = V_EZI_POCICION_IVA_CLIENTES.idcliente LEFT OUTER JOIN " & _
              "V_TIPOPAGO_ ON ALIAS_0.TIPOPAGO_ID = V_TIPOPAGO_.ID ON ALIAS_3.ID = ALIAS_0.ENTEASOCIADO_ID LEFT OUTER JOIN " & _
              "V_CIUDAD_ RIGHT OUTER JOIN V_DOMICILIO_ AS ALIAS_6 ON V_CIUDAD_.ID = ALIAS_6.CIUDAD_ID LEFT OUTER JOIN " & _
              "V_PROVINCIA_ ON ALIAS_6.PROVINCIA_ID = V_PROVINCIA_.ID ON ALIAS_0.DOMICILIOFACTURACION_ID = ALIAS_6.ID LEFT OUTER JOIN " & _
              "V_TELEFONO_ AS ALIAS_7 ON ALIAS_3.TELEFONOPRINCIPAL_ID = ALIAS_7.ID LEFT OUTER JOIN " & _
              "V_DIRECCIONELECTRONICA_ AS ALIAS_8 ON ALIAS_3.DIRECELECTRONICAPRINCIPAL_ID = ALIAS_8.ID LEFT OUTER JOIN " & _
              "V_ZONA_ AS ALIAS_5 ON ALIAS_0.ZONA_ID = ALIAS_5.ID " & _
              "WHERE     (ALIAS_0.ACTIVESTATUS = 0) AND ALIAS_0.CODIGO + ' ' + ALIAS_3.CUIT + ' ' + ALIAS_3.NOMBRE + ' ' + ALIAS_0.DENOMINACION  like '" & xbusqueda & "' order by ALIAS_3.NOMBRE "
              
              
   Else
   ' xquery1 = "SELECT  ALIAS_0.ID AS ID,  ALIAS_0.CODIGO, ALIAS_3.NOMBRE AS RAZONSOCIAL, ALIAS_3.CUIT,ALIAS_6.CALLE, ISNULL(ALIAS_6.CALLE,'') + '-'+ ISNULL(V_CIUDAD_.NOMBRE,'') + '-'+ ISNULL(V_PROVINCIA_.NOMBRE,'') AS DOMICILIO, ALIAS_0.DENOMINACION,ALIAS_5.NOMBRE AS ZONA, " & _
   '           "ALIAS_7.NUMERO AS TELEFONO, ALIAS_8.DIRECCIONELECTRONICA AS MAIL, " & _
   '           "V_TIPOPAGO_.NOMBRE AS TP, V_TIPOPAGO_.OBSERVACION as TipoPago, V_TIPOPAGO_.ACTIVESTATUS AS ACTIVESTATUSTP, V_PROVINCIA_.NOMBRE AS Provincia, " & _
   '           "V_CIUDAD_.NOMBRE AS Ciudad, ALIAS_0.CODIGO + ' ' + ALIAS_3.CUIT + ' ' + ALIAS_3.NOMBRE AS CONCATENADO, V_EZI_POCICION_IVA_CLIENTES.CODIGO AS IVA, V_TIPOPAGO_.ID AS IDPAGO, ALIAS_0.DOMICILIOFACTURACION_ID as domicilio_id, ALIAS_0.LISTAPRECIO_ID as listaprecio  " & _
   '           "FROM V_PERSONA_ AS ALIAS_3 RIGHT OUTER JOIN V_CLIENTE AS ALIAS_0 WITH (READPAST) LEFT OUTER JOIN " & _
   '           "V_EZI_POCICION_IVA_CLIENTES ON ALIAS_0.ID = V_EZI_POCICION_IVA_CLIENTES.idcliente LEFT OUTER JOIN " & _
   '           "V_TIPOPAGO_ ON ALIAS_0.TIPOPAGO_ID = V_TIPOPAGO_.ID ON ALIAS_3.ID = ALIAS_0.ENTEASOCIADO_ID LEFT OUTER JOIN " & _
   '           "V_CIUDAD_ RIGHT OUTER JOIN V_DOMICILIO_ AS ALIAS_6 ON V_CIUDAD_.ID = ALIAS_6.CIUDAD_ID LEFT OUTER JOIN " & _
   '           "V_PROVINCIA_ ON ALIAS_6.PROVINCIA_ID = V_PROVINCIA_.ID ON ALIAS_0.DOMICILIOFACTURACION_ID = ALIAS_6.ID LEFT OUTER JOIN " & _
   '           "V_TELEFONO_ AS ALIAS_7 ON ALIAS_3.TELEFONOPRINCIPAL_ID = ALIAS_7.ID LEFT OUTER JOIN " & _
   '           "V_DIRECCIONELECTRONICA_ AS ALIAS_8 ON ALIAS_3.DIRECELECTRONICAPRINCIPAL_ID = ALIAS_8.ID LEFT OUTER JOIN " & _
   '           "V_ZONA_ AS ALIAS_5 ON ALIAS_0.ZONA_ID = ALIAS_5.ID " & _
   '           "WHERE     ALIAS_0.ID = '" & Text19.Text & "'"
              
    xquery1 = "SELECT     ALIAS_0.ID, ALIAS_0.CODIGO, ALIAS_3.NOMBRE AS RAZONSOCIAL, ALIAS_3.CUIT, ALIAS_6.CALLE, ISNULL(ALIAS_6.CALLE, '') + '-' + ISNULL(V_CIUDAD_.NOMBRE, '') " & _
              "+ '-' + ISNULL(V_PROVINCIA_.NOMBRE, '') AS DOMICILIO, ALIAS_0.DENOMINACION, ALIAS_5.NOMBRE AS ZONA, ALIAS_7.NUMERO AS TELEFONO, " & _
              "ALIAS_8.DIRECCIONELECTRONICA AS MAIL, V_TIPOPAGO_.NOMBRE AS TP, V_TIPOPAGO_.OBSERVACION AS TipoPago, " & _
              "V_TIPOPAGO_.ACTIVESTATUS AS ACTIVESTATUSTP, V_PROVINCIA_.NOMBRE AS Provincia, V_CIUDAD_.NOMBRE AS Ciudad, " & _
              "ALIAS_0.CODIGO + ' ' + ALIAS_3.CUIT + ' ' + ALIAS_3.NOMBRE AS CONCATENADO, V_EZI_POCICION_IVA_CLIENTES.CODIGO AS IVA, V_TIPOPAGO_.ID AS IDPAGO, " & _
              "ALIAS_0.DOMICILIOFACTURACION_ID AS domicilio_id, ALIAS_0.LISTAPRECIO_ID AS listaprecio, V_UD_CLIENTE.observacion, ALIAS_0.creditomaximo " & _
              "FROM         V_TIPOPAGO_ RIGHT OUTER JOIN " & _
              "V_CLIENTE AS ALIAS_0 WITH (NOLOCK) LEFT OUTER JOIN " & _
              "V_UD_CLIENTE with (nolock) ON ALIAS_0.BOEXTENSION_ID = V_UD_CLIENTE.ID LEFT OUTER JOIN " & _
              "V_EZI_POCICION_IVA_CLIENTES ON ALIAS_0.ID = V_EZI_POCICION_IVA_CLIENTES.idcliente ON V_TIPOPAGO_.ID = ALIAS_0.TIPOPAGO_ID LEFT OUTER JOIN " & _
              "V_PERSONA AS ALIAS_3 WITH (nolock) ON ALIAS_0.ENTEASOCIADO_ID = ALIAS_3.ID LEFT OUTER JOIN " & _
              "V_CIUDAD_ RIGHT OUTER JOIN " & _
              "V_DOMICILIO_ AS ALIAS_6 ON V_CIUDAD_.ID = ALIAS_6.CIUDAD_ID LEFT OUTER JOIN " & _
              "V_PROVINCIA_ ON ALIAS_6.PROVINCIA_ID = V_PROVINCIA_.ID ON ALIAS_0.DOMICILIOFACTURACION_ID = ALIAS_6.ID LEFT OUTER JOIN " & _
              "V_TELEFONO_ AS ALIAS_7 ON ALIAS_3.TELEFONOPRINCIPAL_ID = ALIAS_7.ID LEFT OUTER JOIN " & _
              "V_DIRECCIONELECTRONICA_ AS ALIAS_8 ON ALIAS_3.DIRECELECTRONICAPRINCIPAL_ID = ALIAS_8.ID LEFT OUTER JOIN " & _
              "V_ZONA_ AS ALIAS_5 ON ALIAS_0.ZONA_ID = ALIAS_5.ID " & _
              "WHERE   (ALIAS_0.ACTIVESTATUS = 0) AND  ALIAS_0.ID = '" & Text19.Text & "' "
              
              
              
   End If
  Else
'    xquery1 = "SELECT  ALIAS_0.ID AS ID,  ALIAS_0.CODIGO, ALIAS_3.NOMBRE AS RAZONSOCIAL, ALIAS_3.CUIT,ALIAS_6.CALLE, ISNULL(ALIAS_6.CALLE,'') + '-'+ ISNULL(V_CIUDAD_.NOMBRE,'') + '-'+ ISNULL(V_PROVINCIA_.NOMBRE,'') AS DOMICILIO, ALIAS_0.DENOMINACION,ALIAS_5.NOMBRE AS ZONA, " & _
'              "ALIAS_7.NUMERO AS TELEFONO, ALIAS_8.DIRECCIONELECTRONICA AS MAIL, " & _
'              "V_TIPOPAGO_.NOMBRE AS TP, V_TIPOPAGO_.OBSERVACION as TipoPago, V_TIPOPAGO_.ACTIVESTATUS AS ACTIVESTATUSTP, V_PROVINCIA_.NOMBRE AS Provincia, " & _
'              "V_CIUDAD_.NOMBRE AS Ciudad, ALIAS_0.CODIGO + ' ' + ALIAS_3.CUIT + ' ' + ALIAS_3.NOMBRE AS CONCATENADO, V_EZI_POCICION_IVA_CLIENTES.CODIGO AS IVA, V_TIPOPAGO_.ID AS IDPAGO, ALIAS_0.DOMICILIOFACTURACION_ID as domicilio_id, ALIAS_0.LISTAPRECIO_ID as listaprecio   " & _
'              "FROM V_PERSONA_ AS ALIAS_3 RIGHT OUTER JOIN V_CLIENTE AS ALIAS_0 WITH (READPAST) LEFT OUTER JOIN " & _
'              "V_EZI_POCICION_IVA_CLIENTES ON ALIAS_0.ID = V_EZI_POCICION_IVA_CLIENTES.idcliente LEFT OUTER JOIN " & _
'              "V_TIPOPAGO_ ON ALIAS_0.TIPOPAGO_ID = V_TIPOPAGO_.ID ON ALIAS_3.ID = ALIAS_0.ENTEASOCIADO_ID LEFT OUTER JOIN " & _
'              "V_CIUDAD_ RIGHT OUTER JOIN V_DOMICILIO_ AS ALIAS_6 ON V_CIUDAD_.ID = ALIAS_6.CIUDAD_ID LEFT OUTER JOIN " & _
'              "V_PROVINCIA_ ON ALIAS_6.PROVINCIA_ID = V_PROVINCIA_.ID ON ALIAS_0.DOMICILIOFACTURACION_ID = ALIAS_6.ID LEFT OUTER JOIN " & _
'              "V_TELEFONO_ AS ALIAS_7 ON ALIAS_3.TELEFONOPRINCIPAL_ID = ALIAS_7.ID LEFT OUTER JOIN " & _
'              "V_DIRECCIONELECTRONICA_ AS ALIAS_8 ON ALIAS_3.DIRECELECTRONICAPRINCIPAL_ID = ALIAS_8.ID LEFT OUTER JOIN " & _
'              "V_ZONA_ AS ALIAS_5 ON ALIAS_0.ZONA_ID = ALIAS_5.ID " & _
'              "WHERE     (ALIAS_0.ACTIVESTATUS = 0) order by ALIAS_3.NOMBRE "
              
    xquery1 = "SELECT     ALIAS_0.ID, ALIAS_0.CODIGO, ALIAS_3.NOMBRE AS RAZONSOCIAL, ALIAS_3.CUIT, ALIAS_6.CALLE, ISNULL(ALIAS_6.CALLE, '') + '-' + ISNULL(V_CIUDAD_.NOMBRE, '') " & _
              "+ '-' + ISNULL(V_PROVINCIA_.NOMBRE, '') AS DOMICILIO, ALIAS_0.DENOMINACION, ALIAS_5.NOMBRE AS ZONA, ALIAS_7.NUMERO AS TELEFONO, " & _
              "ALIAS_8.DIRECCIONELECTRONICA AS MAIL, V_TIPOPAGO_.NOMBRE AS TP, V_TIPOPAGO_.OBSERVACION AS TipoPago, " & _
              "V_TIPOPAGO_.ACTIVESTATUS AS ACTIVESTATUSTP, V_PROVINCIA_.NOMBRE AS Provincia, V_CIUDAD_.NOMBRE AS Ciudad, " & _
              "ALIAS_0.CODIGO + ' ' + ALIAS_3.CUIT + ' ' + ALIAS_3.NOMBRE AS CONCATENADO, V_EZI_POCICION_IVA_CLIENTES.CODIGO AS IVA, V_TIPOPAGO_.ID AS IDPAGO, " & _
              "ALIAS_0.DOMICILIOFACTURACION_ID AS domicilio_id, ALIAS_0.LISTAPRECIO_ID AS listaprecio, V_UD_CLIENTE.observacion, ALIAS_0.creditomaximo " & _
              "FROM         V_TIPOPAGO_ RIGHT OUTER JOIN " & _
              "V_CLIENTE AS ALIAS_0 WITH (NOLOCK) LEFT OUTER JOIN " & _
              "V_UD_CLIENTE with (nolock) ON ALIAS_0.BOEXTENSION_ID = V_UD_CLIENTE.ID LEFT OUTER JOIN " & _
              "V_EZI_POCICION_IVA_CLIENTES ON ALIAS_0.ID = V_EZI_POCICION_IVA_CLIENTES.idcliente ON V_TIPOPAGO_.ID = ALIAS_0.TIPOPAGO_ID LEFT OUTER JOIN " & _
              "V_PERSONA AS ALIAS_3 WITH (nolock) ON ALIAS_0.ENTEASOCIADO_ID = ALIAS_3.ID LEFT OUTER JOIN " & _
              "V_CIUDAD_ RIGHT OUTER JOIN " & _
              "V_DOMICILIO_ AS ALIAS_6 ON V_CIUDAD_.ID = ALIAS_6.CIUDAD_ID LEFT OUTER JOIN " & _
              "V_PROVINCIA_ ON ALIAS_6.PROVINCIA_ID = V_PROVINCIA_.ID ON ALIAS_0.DOMICILIOFACTURACION_ID = ALIAS_6.ID LEFT OUTER JOIN " & _
              "V_TELEFONO_ AS ALIAS_7 ON ALIAS_3.TELEFONOPRINCIPAL_ID = ALIAS_7.ID LEFT OUTER JOIN " & _
              "V_DIRECCIONELECTRONICA_ AS ALIAS_8 ON ALIAS_3.DIRECELECTRONICAPRINCIPAL_ID = ALIAS_8.ID LEFT OUTER JOIN " & _
              "V_ZONA_ AS ALIAS_5 ON ALIAS_0.ZONA_ID = ALIAS_5.ID " & _
              "WHERE     (ALIAS_0.ACTIVESTATUS = 0) order by ALIAS_3.NOMBRE "
              
  End If

    Text19.Text = ""

    datcliente.RecordSource = xquery1
    datcliente.Refresh
    If datcliente.Recordset.EOF = True Then
        mensa = MsgBox("No existe Cliente", vbInformation, "!! Atencion !!")
        Text1(1).Text = ""
        Text1(1).SetFocus
    End If
    
    If datcliente.Recordset.RecordCount = 1 Then
        Text1(1).Text = DataGrid2.Columns(2).Text
        If DataGrid2.Columns(16).Text = "RI" Then
            Text1(2).Text = "A"
        Else
            Text1(2).Text = "B"
        End If
        If DataGrid2.Columns(17).Text <> "" Then
            DataCombo3.BoundText = DataGrid2.Columns(17).Text
            testpos = InStr(1, DataCombo3.Text, "- ", 1)
            tpago = Right(DataCombo3.Text, Len(DataCombo3.Text) - testpos - 1)
        Else
            tpago = "CONTADO"
            DataCombo3.Text = "01- CONTADO"
        End If
        Text1(4).Text = DataGrid2.Columns(3).Text
        Text1(3).Text = DataGrid2.Columns(5).Text
        If tpago = "CONTADO" Then
            DataCombo3.Enabled = False
        Else
            DataCombo3.Enabled = True
        End If
        
        If DataGrid2.Columns(16).Text = "CF" And Text1(1).Text = "CONSUMIDOR FINAL" Then
            Label1(5).Visible = False
            DataCombo3.Visible = False
            Label1(16).Visible = True
            Text1(6).Visible = True
            Text1(3).Enabled = True
            Text1(6).SetFocus
        Else
            Label1(5).Visible = True
            DataCombo3.Visible = True
            Label1(16).Visible = False
            Text1(6).Visible = False
            Text1(3).Enabled = False
            Text1(5).SetFocus
        End If

        
    Else
        menu = 2
        query = xquery1
        lista_clientes.Show
    End If


End Sub

Private Sub blanco_Click()

    negro.Visible = True
    blanco.Visible = False
    tipofac = "NN"

End Sub

Private Sub buscar_Click()

    menu = 5
    lista_presupuestos.Show

End Sub

Private Sub bvendedor_Click()
    
  If Text1(0).Text <> "" Then
    xbusqueda = "%" + Text1(0).Text + "%"
    xquery1 = "SELECT V_VENDEDOR_.ID, V_VENDEDOR_.CODIGO, V_PERSONA_.NOMBRE, V_VENDEDOR_.CODIGO + ' ' + V_PERSONA_.NOMBRE AS cancatena,ud_ezi_empleado.limite " & _
              "FROM V_VENDEDOR_ INNER JOIN V_PERSONA_ ON V_VENDEDOR_.ENTEASOCIADO_ID = V_PERSONA_.ID LEFT OUTER JOIN ud_ezi_empleado WITH (nolock) ON V_VENDEDOR_.ID = ud_ezi_empleado.id " & _
              "WHERE (V_VENDEDOR_.ACTIVESTATUS = 0)  and V_VENDEDOR_.CODIGO  + ' ' + V_PERSONA_.NOMBRE like '" & xbusqueda & "' order by V_PERSONA_.NOMBRE"

'    xquery1 = "SELECT    V_VENDEDOR_.ID, V_VENDEDOR_.CODIGO, V_PERSONA_.NOMBRE, V_VENDEDOR_.CODIGO  + ' ' + V_PERSONA_.NOMBRE as cancatena " & _
'              "FROM V_VENDEDOR_ INNER JOIN V_PERSONA_ ON V_VENDEDOR_.ENTEASOCIADO_ID = V_PERSONA_.ID " & _
'              "WHERE (V_VENDEDOR_.ACTIVESTATUS = 0)  and V_VENDEDOR_.CODIGO  + ' ' + V_PERSONA_.NOMBRE like '" & xbusqueda & "' order by V_PERSONA_.NOMBRE"
  Else
'    xquery1 = "SELECT    V_VENDEDOR_.ID, V_VENDEDOR_.CODIGO, V_PERSONA_.NOMBRE, V_VENDEDOR_.CODIGO  + ' ' + V_PERSONA_.NOMBRE as cancatena " & _
'              "FROM V_VENDEDOR_ INNER JOIN V_PERSONA_ ON V_VENDEDOR_.ENTEASOCIADO_ID = V_PERSONA_.ID " & _
'              "WHERE (V_VENDEDOR_.ACTIVESTATUS = 0)  order by V_PERSONA_.NOMBRE"

     xqueri1 = "SELECT V_VENDEDOR_.ID, V_VENDEDOR_.CODIGO, V_PERSONA_.NOMBRE, V_VENDEDOR_.CODIGO + ' ' + V_PERSONA_.NOMBRE AS cancatena,ud_ezi_empleado.limite " & _
               "FROM V_VENDEDOR_ INNER JOIN V_PERSONA_ ON V_VENDEDOR_.ENTEASOCIADO_ID = V_PERSONA_.ID LEFT OUTER JOIN ud_ezi_empleado WITH (nolock) ON V_VENDEDOR_.ID = ud_ezi_empleado.id " & _
               "WHERE (V_VENDEDOR_.ACTIVESTATUS = 0)  order by V_PERSONA_.NOMBRE"
  End If


    datvendedor.RecordSource = xquery1
    datvendedor.Refresh
    If datvendedor.Recordset.EOF = True Then
        mensa = MsgBox("No existe Vendedor", vbInformation, "!! Atencion !!")
        Text1(0).Text = ""
        Text1(0).SetFocus
    End If
    
    If datvendedor.Recordset.RecordCount = 1 Then
        Text1(0).Text = DataGrid1.Columns(2).Text
        Text1(1).SetFocus
        xlimitebonif = DataGrid1.Columns(4).Text
    Else
        menu = 2
        query = xquery1
        lista_vendedores.Show
    End If
    
    



End Sub

Private Sub calcula_Click()
On Error Resume Next

Exit Sub
    xcol = grilla.Col
    grilla.Col = 3
    xcant = Val(grilla.Text)
    
    grilla.Col = 6
    ximporte = grilla.Text
    grilla.Text = Format(ximporte, "###,##0.00")
    If grilla.Text = "0.00" And grilla.TextMatrix(grilla.Row, 0) = "" Then
        grilla.Text = ""
    End If
   
     If ximporte < Val(Format(grilla.TextMatrix(grilla.Row, 17), "0.00")) And xlimitebonif <> 100 Then
        MsgBox "No puede ingresar un importe menor al precio de lista", vbCritical, "Error"
        grilla.Text = Format(grilla.TextMatrix(grilla.Row, 17), "###,##0.00")
        Text2.Text = grilla.Text
        grilla.TextMatrix(grilla.Row, 5) = Format(Round(Val(Format(grilla.TextMatrix(grilla.Row, 17), "0.00")) / Val(Format(grilla.TextMatrix(grilla.Row, 12), "0.000")), 2), "###,##0.00")
        grilla.Col = xcol
        Exit Sub
    End If
        
    If ximporte = "" Then ximporte = 0
    grilla.Text = Format(ximporte, "###,##0.00")
    If grilla.Text = "0.00" And grilla.TextMatrix(grilla.Row, 0) = "" Then
        grilla.Text = ""
    End If

    grilla.Col = 12
    xiva = grilla.Text
    If xiva = "" Then xiva = 1.21

    ximportesiva2 = Round(ximporte / xiva, 3)
    grilla.Col = 5
    grilla.Text = Format(ximportesiva2, "###,##0.000")
    If grilla.Text = "0.000" And grilla.TextMatrix(grilla.Row, 0) = "" Then
        grilla.Text = ""
    End If

    ximportesiva = Round(ximportesiva2 * xcant, 3)
    grilla.Col = 7
    grilla.Text = Format(ximportesiva, "###,##0.00")
    If grilla.Text = "0.00" And grilla.TextMatrix(grilla.Row, 0) = "" Then
        grilla.Text = ""
    End If
    
    xbonifporcent = grilla.TextMatrix(grilla.Row, 9)
    xbonifimporte = grilla.TextMatrix(grilla.Row, 8)
        
    
    If xcol = 8 Or Val(Text11.Text) <> 0 Then
        grilla.Col = 8
        xbonifimporte = grilla.Text
        grilla.Text = Format(xbonifimporte, "###,##0.00")
        If grilla.Text = "0.00" And grilla.TextMatrix(grilla.Row, 0) = "" Then
            grilla.Text = ""
        End If
        grilla.Col = 9
'        xbonifporcent = Round((xbonifimporte / (xcant * ximporte)) * 100, 2)
        xbonifporcent = Round((xbonifimporte / (xcant * (ximporte / xiva))) * 100, 2)
        grilla.Text = Format(xbonifporcent, "###,##0.00")
        If grilla.Text = "0.00" And grilla.TextMatrix(grilla.Row, 0) = "" Then
          grilla.Text = ""
        End If

    End If

    If xcol = 9 Or xcol = 3 Or xcol = 5 Or xcol = 6 Or Val(Text10.Text) <> 0 Then
        grilla.Col = 9
        xbonifporcent = grilla.Text
        
        If xlimitebonif <> 100 Then
            If xbonifporcent > xlimitebonif Then
                MsgBox "No puede ingresar un porcentaje de bonificacion mayor a al " + Str(xlimitebonif) + " %"
                xbonifporcent = xlimitebonif
            End If
        End If
        
        grilla.Text = Format(xbonifporcent, "###,##0.00")
        If grilla.Text = "0.00" And grilla.TextMatrix(grilla.Row, 0) = "" Then
            grilla.Text = ""
        End If
        grilla.Col = 8
'        xbonifimporte = Round(xbonifporcent * xcant * ximporte / 100)
        xbonifimporte = Round(xbonifporcent * xcant * (ximporte / xiva) / 100, 2)
        grilla.Text = Format(xbonifimporte, "###,##0.00")
        If grilla.Text = "0.00" And grilla.TextMatrix(grilla.Row, 0) = "" Then
            grilla.Text = ""
        End If

    End If
    
    grilla.Col = 10
    If xbonifimporte = "" Then xbonifimporte = 0
    If ximporte = "" Then ximporte = 0
    
    If grilla.TextMatrix(grilla.Row, 12) = "" Then
        grilla.TextMatrix(grilla.Row, 12) = 1.21
    End If
    
    'xtotal = Round(Round(ximportesiva * grilla.TextMatrix(grilla.Row, 12), 3) - Round(xbonifimporte, 10), 2)
    xtotal = Round((ximportesiva - xbonifimporte) * grilla.TextMatrix(grilla.Row, 12), 3) ' corregido 04/01/2016
    
    
 '   xtotal = Round(xcant * Round(ximporte, 10) - Round(xbonifimporte, 10), 2)
    grilla.Text = Format(xtotal, "###,##0.00")
    If grilla.Text = "0.00" And grilla.TextMatrix(grilla.Row, 0) = "" Then
          grilla.Text = ""
    End If


    grilla.Col = 11
    grilla.Text = xcant
        If grilla.Text = "0" And grilla.TextMatrix(grilla.Row, 0) = "" Then
            grilla.Text = ""
        End If


    xtotalgral = 0
    xsubtotalgral = 0
    xiva10 = 0
    xiva21 = 0
    For X = 1 To xlineasmax
            If grilla.TextMatrix(X, 10) = "" Then
                xgrilla = 0
                xsubtotal = 0
            Else
                xgrilla = grilla.TextMatrix(X, 10)
                If grilla.TextMatrix(X, 11) = "" Then
                 '   xsubtotal = grilla.TextMatrix(X, 10) / 1.21
                     xsubtotal = grilla.TextMatrix(X, 7) - grilla.TextMatrix(X, 8) ' corregido 04/01/2016
'                    xsubtotal = grilla.TextMatrix(X, 7)
                Else
'                    xsubtotal = grilla.TextMatrix(X, 10) / grilla.TextMatrix(X, 12)
                     xsubtotal = grilla.TextMatrix(X, 7) - grilla.TextMatrix(X, 8) ' corregido 04/01/2016
'                     xsubtotal = grilla.TextMatrix(X, 7)
                End If
            End If
            
            If grilla.TextMatrix(X, 12) = "1.105" Then
                xiva10 = xiva10 + (xgrilla - xsubtotal)
            Else
                xiva21 = xiva21 + (xgrilla - xsubtotal)
            End If
            
            xtotalgral = xtotalgral + xgrilla
            xsubtotalgral = xsubtotalgral + xsubtotal
            
    Next X
    
'--- Calculo de Tem
    xcalculotem = 0
    If UCase(xcalculatempyp) = "S" And xciudadtem = xciudadcliente And DataGrid2.Columns(16).Text <> "CF" And DataGrid2.Columns(16).Text <> "EX" And tipofac <> "NN" And xcontroltem <> 0 Then
        xcalculotem = xsubtotalgral * xalicuptatempip / 100
    End If
'--- Fin Calculo Tem

'--- Calculo de IIBB
    xcalculoIIBB = 0
    If UCase(xcalculaiibb) = "S" And DataGrid2.Columns(16).Text <> "CF" And DataGrid2.Columns(16).Text <> "EX" And tipofac <> "NN" Then
        xcalculoIIBB = (xsubtotalgral * xalicuotaiibb / 100) * ((100 - xexentoiibb) / 100)
  '      If xcalculoIIBB <= 50 Then xcalculoIIBB = 0  ' Limite inferior para calculo de iibb
    End If
'--- Fin Calculo IIBB
    xtotalgral = xtotalgral + xcalculotem + xcalculoIIBB
    xsubtotal2 = xsubtotalgral + xiva10 + xiva21
    
    Text5.Text = Format(Round(xsubtotalgral, 2), "$ ###,##0.00")
    Text6.Text = Format(Round(xiva10, 3), "$ ###,##0.00")
    Text7.Text = Format(Round(xiva21, 3), "$ ###,##0.00")
    Text8.Text = Format(Round(xcalculotem, 2), "$ ###,##0.00")
    Text9.Text = Format(Round(xcalculoIIBB, 2), "$ ###,##0.00")
    Text16.Text = Format(Round(xsubtotal2, 2), "$ ###,##0.00")
    
    Text4.Text = Format(Round(xtotalgral, 2), "$ ###,##0.00")
        

    grilla.Col = xcol


End Sub

Private Sub Cancelar_Click()

    Unload Me
    menu = 5
    frmcomparativa.Show

End Sub

Private Sub DataList1_KeyPress(KeyAscii As Integer)
On Error GoTo fuera

    If KeyAscii = 13 Then
        KeyAscii = 0
        Text1(5).Text = DataList1.BoundText
        Text1(6).SetFocus
    End If

fuera:

End Sub

Private Sub DataList1_LostFocus()

    DataList1.Visible = False

End Sub

Private Sub cargapresupuesto_Click()
On Error Resume Next

If menu = 6 Then
    datencabezado.RecordSource = "SELECT * From ud_ezi_puntodeventa_encabezado where  numeradorinterno = 'Comparativa de Precios' and  numerodefactura ='" & Text17.Text & "'"
    datencabezado.Refresh
Else
If Text18.Text <> "" Then
    datencabezado.RecordSource = "SELECT * From ud_ezi_puntodeventa_encabezado where  numeradorinterno = 'Presupuesto de Venta' and  id ='" & Text18.Text & "' and generada = 'False'"
    datencabezado.Refresh
Else
    datencabezado.RecordSource = "SELECT * From ud_ezi_puntodeventa_encabezado where  numeradorinterno = 'Presupuesto de Venta' and  numerodefactura ='" & Text17.Text & "' and generada = 'False'"
    datencabezado.Refresh
End If
End If

    If datencabezado.Recordset.EOF = False Then
        xidpre = datencabezado.Recordset.Fields("id")
    '**** Carga cliente
        xquericliente = "SELECT  ALIAS_0.ID AS ID,  ALIAS_0.CODIGO, ALIAS_3.NOMBRE AS RAZONSOCIAL, ALIAS_3.CUIT,ALIAS_6.CALLE, ISNULL(ALIAS_6.CALLE,'') + '-'+ ISNULL(V_CIUDAD_.NOMBRE,'') + '-'+ ISNULL(V_PROVINCIA_.NOMBRE,'') AS DOMICILIO, ALIAS_0.DENOMINACION,ALIAS_5.NOMBRE AS ZONA, " & _
              "ALIAS_7.NUMERO AS TELEFONO, ALIAS_8.DIRECCIONELECTRONICA AS MAIL, " & _
              "V_TIPOPAGO_.NOMBRE AS TP, V_TIPOPAGO_.OBSERVACION as TipoPago, V_TIPOPAGO_.ACTIVESTATUS AS ACTIVESTATUSTP, V_PROVINCIA_.NOMBRE AS Provincia, " & _
              "V_CIUDAD_.NOMBRE AS Ciudad, ALIAS_0.CODIGO + ' ' + ALIAS_3.CUIT + ' ' + ALIAS_3.NOMBRE AS CONCATENADO, V_EZI_POCICION_IVA_CLIENTES.CODIGO AS IVA, V_TIPOPAGO_.ID AS IDPAGO, ALIAS_0.DOMICILIOFACTURACION_ID as domicilio_id  " & _
              "FROM V_PERSONA_ AS ALIAS_3 RIGHT OUTER JOIN V_CLIENTE AS ALIAS_0 WITH (READPAST) LEFT OUTER JOIN " & _
              "V_EZI_POCICION_IVA_CLIENTES ON ALIAS_0.ID = V_EZI_POCICION_IVA_CLIENTES.idcliente LEFT OUTER JOIN " & _
              "V_TIPOPAGO_ ON ALIAS_0.TIPOPAGO_ID = V_TIPOPAGO_.ID ON ALIAS_3.ID = ALIAS_0.ENTEASOCIADO_ID LEFT OUTER JOIN " & _
              "V_CIUDAD_ RIGHT OUTER JOIN V_DOMICILIO_ AS ALIAS_6 ON V_CIUDAD_.ID = ALIAS_6.CIUDAD_ID LEFT OUTER JOIN " & _
              "V_PROVINCIA_ ON ALIAS_6.PROVINCIA_ID = V_PROVINCIA_.ID ON ALIAS_0.DOMICILIOFACTURACION_ID = ALIAS_6.ID LEFT OUTER JOIN " & _
              "V_TELEFONO_ AS ALIAS_7 ON ALIAS_3.TELEFONOPRINCIPAL_ID = ALIAS_7.ID LEFT OUTER JOIN " & _
              "V_DIRECCIONELECTRONICA_ AS ALIAS_8 ON ALIAS_3.DIRECELECTRONICAPRINCIPAL_ID = ALIAS_8.ID LEFT OUTER JOIN " & _
              "V_ZONA_ AS ALIAS_5 ON ALIAS_0.ZONA_ID = ALIAS_5.ID " & _
              "WHERE     (ALIAS_0.ACTIVESTATUS <> 2) and ALIAS_0.ID = '" & datencabezado.Recordset.Fields("clienteid") & "'"
              
        datcliente.RecordSource = xquericliente
        datcliente.Refresh
        
        Text1(1).Text = DataGrid2.Columns(2).Text
        If DataGrid2.Columns(16).Text = "RI" Then
            Text1(2).Text = "A"
        Else
            Text1(2).Text = "B"
        End If
        If DataGrid2.Columns(17).Text <> "" Then
            DataCombo3.BoundText = DataGrid2.Columns(17).Text
                        testpos = InStr(1, DataCombo3.Text, "- ", 1)
            tpago = Right(DataCombo3.Text, Len(DataCombo3.Text) - testpos - 1)
        Else
            tpago = "CONTADO"
            DataCombo3.Text = "01- CONTADO"
        End If
        Text1(4).Text = DataGrid2.Columns(3).Text
        Text1(3).Text = DataGrid2.Columns(6).Text
        If tpago = "CONTADO" Then
            DataCombo3.Enabled = False
        Else
            DataCombo3.Enabled = True
        End If
        Text1(5).SetFocus
        
        xcontroltem = 1
        datiibb.RecordSource = "SELECT * from v_ezi_pos_impuestos " & _
                               "WHERE idcliente = '" & DataGrid2.Columns(0).Text & "'"
        datiibb.Refresh


        If datiibb.Recordset.EOF = False Then
            xalicuotaiibb = datiibb.Recordset.Fields("COEFICIENTE")
            xexentoiibb = datiibb.Recordset.Fields("exencion")
            If xexentoiibb = 0 Then
                Label1(14).Caption = "Percep IIBB " + Str(xalicuotaiibb) + " %"
            Else
                Label1(14).Caption = "Percep IIBB " + Str(xalicuotaiibb) + " % Cert.Ex"
            End If
            If DataGrid2.Columns(16).Text = "CF" Then Label1(14).Caption = "Percep IIBB "
            
        End If
        xcontroltem = datiibb.Recordset.Fields("tem")
        xciudadcliente = DataGrid2.Columns("Ciudad").Text
'*** fin Carga Clinte
'*** Carga Vendedor
        xqueryvende = "SELECT    V_VENDEDOR_.ID, V_VENDEDOR_.CODIGO, V_PERSONA_.NOMBRE, V_VENDEDOR_.CODIGO  + ' ' + V_PERSONA_.NOMBRE as cancatena " & _
              "FROM V_VENDEDOR_ INNER JOIN V_PERSONA_ ON V_VENDEDOR_.ENTEASOCIADO_ID = V_PERSONA_.ID " & _
              "WHERE (V_VENDEDOR_.ACTIVESTATUS = 0) and V_VENDEDOR_.ID = '" & datencabezado.Recordset.Fields("vendedorid") & "'"
    
        datvendedor.RecordSource = xqueryvende
        datvendedor.Refresh
        Text1(0).Text = DataGrid1.Columns(2).Text
        Text1(1).SetFocus
'*** Fin carga Vendedor
        Text1(5).Text = datencabezado.Recordset.Fields("detalle")
'        Text15.Text = datencabezado.Recordset.Fields("nota")

        
        datitems.RecordSource = "Select * from ud_ezi_puntodeventa_detalle_presu where claveprimaria = " & xidpre & " order by id"
        datitems.Refresh
    
        If menu = 6 Then
            grilla.TextMatrix(0, 8) = datencabezado.Recordset.Fields("cprov1")
            grilla.TextMatrix(0, 9) = datencabezado.Recordset.Fields("cprov2")
            grilla.TextMatrix(0, 10) = datencabezado.Recordset.Fields("cprov3")
            grilla.TextMatrix(0, 11) = datencabezado.Recordset.Fields("cprov4")
            grilla.TextMatrix(0, 12) = datencabezado.Recordset.Fields("cprov5")
            grilla.TextMatrix(0, 13) = datencabezado.Recordset.Fields("cprov6")
        End If
    
        datitems.Recordset.MoveFirst
        For X = 1 To datitems.Recordset.RecordCount
            grilla.TextMatrix(X, 0) = datitems.Recordset.Fields("idproducto")
            grilla.TextMatrix(X, 1) = datitems.Recordset.Fields("codigoproducto")
            grilla.TextMatrix(X, 2) = datitems.Recordset.Fields("nombre_producto")
            grilla.TextMatrix(X, 3) = datitems.Recordset.Fields("cantidadproducto")
            grilla.TextMatrix(X, 4) = datitems.Recordset.Fields("unidaddemedidaid")
            grilla.TextMatrix(X, 5) = Format(Round(datitems.Recordset.Fields("preciou"), 2), "#,###,##0.00`")
            grilla.TextMatrix(X, 6) = datitems.Recordset.Fields("bonificacionitem")
'            grilla.TextMatrix(X, 5) = Format(Round(datitems.Recordset.Fields("subtotal") / datitems.Recordset.Fields("cantidadproducto"), 5), "#,###,##0.00`")
            
            If menu = 6 Then
                grilla.TextMatrix(X, 6) = datitems.Recordset.Fields("bonificacionitem")
                grilla.TextMatrix(X, 7) = datitems.Recordset.Fields("porcentaje")
                grilla.TextMatrix(X, 8) = datitems.Recordset.Fields("cprecioprov1")
                grilla.TextMatrix(X, 9) = datitems.Recordset.Fields("cmarcap1")
                grilla.TextMatrix(X, 10) = datitems.Recordset.Fields("cprecioprov2")
                grilla.TextMatrix(X, 11) = datitems.Recordset.Fields("cmarcap2")
                grilla.TextMatrix(X, 12) = datitems.Recordset.Fields("cprecioprov3")
                grilla.TextMatrix(X, 13) = datitems.Recordset.Fields("cmarcap3")
                grilla.TextMatrix(X, 14) = datitems.Recordset.Fields("cprecioprov4")
                grilla.TextMatrix(X, 15) = datitems.Recordset.Fields("cmarcap4")
                grilla.TextMatrix(X, 16) = datitems.Recordset.Fields("cprecioprov5")
                grilla.TextMatrix(X, 17) = datitems.Recordset.Fields("cmarcap5")
                grilla.TextMatrix(X, 18) = datitems.Recordset.Fields("cprecioprov6")
                grilla.TextMatrix(X, 19) = datitems.Recordset.Fields("cmarcap6")
                grilla.TextMatrix(X, 20) = datitems.Recordset.Fields("id")
            End If
            
''  semaforiza Comparativa
                grilla.Col = 9
                grilla.Row = X
                xcolor = grilla.CellBackColor

                xmenor = Val(grilla.TextMatrix(X, 5))
                xmenor2 = Val(grilla.TextMatrix(X, 8))
                For Y = 8 To 18 Step 2
                    xvalor = Val(grilla.TextMatrix(X, Y))
                    grilla.Col = Y
    
                    If xvalor <= xmenor2 And xvalor <> 0 Then
                        xmenor2 = xvalor
                        xcolmenor = Y
                    End If
    
                    If xvalor < xmenor And xvalor <> 0 Then
                        xmenor = Val(grilla.TextMatrix(X, Y))
                        xcolmenor = Y
                    End If
                    grilla.CellBackColor = xcolor

                Next Y

                grilla.Col = xcolmenor
                grilla.CellBackColor = &HFF00&

                If xmenor2 > xmenor Then xmenor = xmenor2
    
                xporcentaje = Val(grilla.TextMatrix(X, 6))
    
                grilla.Col = 7
                If xporcentaje < 0 Then
                    grilla.CellBackColor = &HFF00&
                Else
                    grilla.CellBackColor = &H8080FF
                End If


''' Fin de Semaforizacion
            
            
            
            grilla.Col = 4
'            Call calcula_Click
            datitems.Recordset.MoveNext
        Next X
    
        If datencabezado.Recordset.Fields("bonificacion") <> 0 Then
            Text11.SetFocus
            Text11.Text = Format(datencabezado.Recordset.Fields("bonificacion"), "###,##0.00")
            SendKeys "{ENTER}", False
        End If
        
        If datencabezado.Recordset.Fields("recargo") <> 0 Then
            Text13.SetFocus
            Text13.Text = Format(datencabezado.Recordset.Fields("recargo"), "###,##0.00")
            SendKeys "{ENTER}", False
        End If
    
    Text17.Locked = True
    
    Else
        mensa = MsgBox("No Existe el Nro de cotizacion seleccionado", vbInformation, "!! Sin Coincidencias !!")
    End If
    


End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
        KeyAscii = 0
        
        If Combo1.Text <> xumvta And grilla.Text <> Combo1.Text Then
            grilla.Text = Combo1.Text
            grilla.Col = 7
            grilla.Text = Round(grilla.Text / xfaconv(Combo1.ListIndex), 2)
            grilla.TextMatrix(grilla.Row, 17) = grilla.Text
            grilla.TextMatrix(grilla.Row, 13) = xfaconv(Combo1.ListIndex)
            Call calcula_Click
            'Call ubicatextogrilla_Click
            Exit Sub
        End If
        If Combo1.Text = xumvta And grilla.Text <> Combo1.Text Then
            grilla.Text = Combo1.Text
            grilla.Col = 7
            grilla.TextMatrix(grilla.Row, 17) = grilla.Text
            grilla.Text = Round(grilla.Text * grilla.TextMatrix(grilla.Row, 13), 2)
            Call calcula_Click
          '  Call ubicatextogrilla_Click
            Exit Sub
        End If
        grilla.Col = 7
       ' Call ubicatextogrilla_Click

End If


End Sub

Private Sub Combo1_KeyUp(KeyCode As Integer, Shift As Integer)

If KeyCode = 39 Then
    grilla.Col = 6
'    Call ubicatextogrilla_Click
End If

If KeyCode = 37 Then
    grilla.Col = 3
'    Call ubicatextogrilla_Click
End If

If KeyCode = 121 Then
       Call grabar_Click
End If

If KeyCode = 116 Then
    If agregaproducto.Enabled = True Then
       Call agregaproducto_Click
    End If
End If

If KeyCode = 123 Then
       Call agregaservicios_Click
End If


End Sub

Private Sub Command7_Click()
On Error Resume Next
    
    menu = 1
    query = "SELECT   top 10  substring(FECHADOCUMENTO,7,2)+'/'+substring(FECHADOCUMENTO,5,2)+'/'+left(FECHADOCUMENTO,4) as fecha, (VALOR2_IMPORTE - (IMPORTEBONIFICADO/cantidad2_cantidad)) as VALOR2_IMPORTE , NOMBREREFERENCIA, DESCRIPCION, REFERENCIA_ID, VALOR2_IMPORTE, DESTINATARIOTR_ID, NOMBREDESTINATARIOTR, FECHADOCUMENTO, " & _
            "left(numerodocumento,4)+'-'+right(numerodocumento,8) as numerodocumento, NOMBREORIGINANTETR " & _
            "From ITEMFACTURAVENTA WHERE     (DESTINATARIOTR_ID = '" & DataGrid2.Columns("id").Text & "') and REFERENCIA_ID = '" & grilla.TextMatrix(grilla.Row, 0) & "' " & _
            "ORDER BY FECHADOCUMENTO DESC"
    lista_historial.Show
    Text2.SetFocus
    


End Sub

Private Sub DataCombo1_KeyPress(KeyAscii As Integer)
On Error Resume Next

    If KeyAscii = 13 Then
        KeyAscii = 0
        DataCombo2.SetFocus
    End If

End Sub

Private Sub DataList2_KeyPress(KeyAscii As Integer)
On Error GoTo fuera

    If KeyAscii = 13 Then
        KeyAscii = 0
        Text1(6).Text = DataList2.BoundText
        grabar.SetFocus
    End If

fuera:
End Sub

Private Sub DataList2_LostFocus()

    DataList2.Visible = False

End Sub

Private Sub DataList3_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        Text1(4).Text = DataList3.Text
        Text1(5).SetFocus
        DataList3.Visible = False
    End If


End Sub

Private Sub DataCombo2_KeyPress(KeyAscii As Integer)

On Error Resume Next

    If KeyAscii = 13 Then
        KeyAscii = 0
        DataCombo3.SetFocus
    End If


End Sub

Private Sub DataCombo3_Click(Area As Integer)
On Error Resume Next
testpos = InStr(1, DataCombo3.Text, "- ", 1)
tpago = Right(DataCombo3.Text, Len(DataCombo3.Text) - testpos - 1)


End Sub

Private Sub DataCombo3_KeyPress(KeyAscii As Integer)

On Error Resume Next

    If KeyAscii = 13 Then
        KeyAscii = 0
        testpos = InStr(1, DataCombo3.Text, "- ", 1)
        tpago = Right(DataCombo3.Text, Len(DataCombo3.Text) - testpos - 1)
        Text1(5).SetFocus
    End If

End Sub



Private Sub Form_Load()
Aplicar_skin Me

yventana = Inicio.Height / 2 - 1000
xventana = Inicio.Width / 2


frmcomparativa.Top = yventana - frmcomparativa.Height / 2
frmcomparativa.Left = xventana - frmcomparativa.Width / 2

datvendedor.ConnectionString = login.conexiontotal
datcliente.ConnectionString = login.conexiontotal
datproductos.ConnectionString = login.conexiontotal
datmovimientos.ConnectionString = login.conexiontotal
dattipopago.ConnectionString = login.conexiontotal
datparametros.ConnectionString = login.conexiontotal
datlistaprecios.ConnectionString = login.conexiontotal
datum.ConnectionString = login.conexiontotal
datiibb.ConnectionString = login.conexiontotal
datencabezado.ConnectionString = login.conexiontotal
datitems.ConnectionString = login.conexiontotal
datcola.ConnectionString = login.conexiontotal


If Left(login.nombrebd, 14) = "MMOSSE" Then agregaservicios.Visible = True
Text3.Text = Date
Text19.Text = ""

    dattipopago.RecordSource = "SELECT ID, NOMBRE AS CODIGO, nombre +'- '+ OBSERVACION AS TipoPago From V_TIPOPAGO_ WHERE (ACTIVESTATUS = 0) order by NOMBRE"
    dattipopago.Refresh
    
    datvendedor.RecordSource = "SELECT     V_VENDEDOR_.CODIGO, V_PERSONA_.NOMBRE FROM V_VENDEDOR_ INNER JOIN " & _
                               "V_PERSONA_ ON V_VENDEDOR_.ENTEASOCIADO_ID = V_PERSONA_.ID Where (V_VENDEDOR_.ACTIVESTATUS = 0)"
    datvendedor.Refresh

    datparametros.RecordSource = "select * from ud_ezi_parametros_pos where sucursal = '" & login.nomsucursal & "' "
    datparametros.Refresh
    
    datlistaprecios.RecordSource = "select id, nombre from v_listaprecio_"
    datlistaprecios.Refresh
    
    DataCombo1.BoundText = datparametros.Recordset.Fields("listaprecio")
    DataCombo1.Enabled = False
    
    xcalculaiibb = datparametros.Recordset.Fields("calculaiibb")
    xcalculatempyp = datparametros.Recordset.Fields("calculatempyp")
    xalicuptatempip = datparametros.Recordset.Fields("alicuotatempip")
    xciudadtem = datparametros.Recordset.Fields("ciudadtem")
    
    
grilla.Row = 0
grilla.Col = 0
grilla.ColWidth(0) = 100
grilla.RowHeight(0) = 500
grilla.Col = 1
grilla.Text = "Codigo"
grilla.ColWidth(1) = 1000
grilla.Col = 2
grilla.Text = "Descipcion"
grilla.ColWidth(2) = 6000
grilla.Col = 3
grilla.Text = "Cant."
grilla.ColWidth(3) = 800
grilla.Col = 4
grilla.Text = "U.M."
grilla.ColWidth(4) = 1000

grilla.Col = 5
grilla.Text = "$ Mosse"
grilla.ColWidth(5) = 1000

grilla.Col = 6
grilla.Text = "% Mosse"
grilla.ColWidth(6) = 1000


grilla.Col = 7
grilla.Text = "% Comp"
grilla.ColWidth(6) = 800

grilla.Col = 8
grilla.Text = "$ Prov 1"
grilla.ColWidth(8) = 1200
grilla.Col = 9
grilla.Text = "Marca Prov1"
grilla.ColWidth(9) = 1200

grilla.Col = 10
grilla.Text = "$ Prov 2"
grilla.ColWidth(10) = 1200
grilla.Col = 11
grilla.Text = "Marca Prov2"
grilla.ColWidth(11) = 1200


grilla.Col = 12
grilla.Text = "$ Prov 3"
grilla.ColWidth(1) = 1200
grilla.Col = 13
grilla.Text = "Marca Prov3"
grilla.ColWidth(13) = 1200

grilla.Col = 14
grilla.Text = "$ Prov4"
grilla.ColWidth(14) = 1200
grilla.Col = 15
grilla.Text = "Marca Prov4"
grilla.ColWidth(15) = 1200

grilla.Col = 16
grilla.Text = "$ Prov 5"
grilla.ColWidth(16) = 1200
grilla.Col = 17
grilla.Text = "Marca Prov5"
grilla.ColWidth(17) = 1200

grilla.Col = 18
grilla.Text = "$ Prov 6"
grilla.ColWidth(18) = 1200
grilla.Col = 19
grilla.Text = "Marca Prov6"
grilla.ColWidth(19) = 1200

grilla.Col = 20
grilla.ColWidth(20) = 0
grilla.ColWidth(21) = 0
grilla.ColWidth(22) = 0
grilla.ColWidth(23) = 0
grilla.ColWidth(24) = 0
grilla.ColWidth(25) = 0

xlineasmax = datparametros.Recordset.Fields("limiteitemsnotaventa")
If login.usuarioactivo <> "admin" Then
    xlinesadd = 0
Else
    xlinesadd = 10
End If
grilla.Rows = xlineasmax + 1 + xlinesadd

For X = 2 To xlineasmax + xlinesadd Step 2
  For Y = 1 To 19
    grilla.Col = Y
    grilla.Row = X
    grilla.CellBackColor = RGB(231, 235, 218)
  Next Y
Next X

If menu = 6 Then
    Frame2.Caption = "Nro de Comparativa"
    Text17.Locked = True
    buscar.Visible = False
    grabar.Caption = "&Modificar"
    If login.usuarioactivo = "admin" Or UCase(login.usuarioactivo) = "DELIA" Or UCase(login.usuarioactivo) = "GRACIELA" Then
        grabar.Enabled = True
    Else
        grabar.Enabled = False
    End If
End If


   
End Sub

Private Sub grabar_Click()
On Error GoTo errorgrabar

If menu <> 6 Then

    datencabezado.RecordSource = "SELECT MAX(CONVERT(decimal, isnull(claveprimaria,0))) AS claveprimaria, numeradorinterno FROM ud_ezi_puntodeventa_encabezado WITH (readpast) " & _
                                 "GROUP BY numeradorinterno HAVING (numeradorinterno = 'Comparativa de Precios')"
    datencabezado.Refresh
    
    If datencabezado.Recordset.EOF = True Then
        xclaveprimaria = 1
    Else
        xclaveprimaria = datencabezado.Recordset.Fields("claveprimaria") + 1
    End If
    

    datencabezado.RecordSource = "SELECT * From ud_ezi_puntodeventa_encabezado where id =0 "
    datencabezado.Refresh
    datitems.RecordSource = "Select * from ud_ezi_puntodeventa_detalle_presu where id = 0"
    datitems.Refresh
    
    datencabezado.Recordset.AddNew
    datencabezado.Recordset.Fields("claveprimaria") = xclaveprimaria


    datencabezado.Recordset.Fields("numeradorinterno") = "Comparativa de Precios"
    datencabezado.Recordset.Fields("fechadelcomprobante") = DateValue(Text3.Text) + TimeValue(Str(Time))
    datencabezado.Recordset.Fields("sucursal") = datparametros.Recordset.Fields("sucursal")
    datencabezado.Recordset.Fields("clienteid") = DataGrid2.Columns(0).Text
    datencabezado.Recordset.Fields("cliente") = DataGrid2.Columns(2).Text
    If Text1(6).Text <> "" Then
        datencabezado.Recordset.Fields("cliente") = Text1(6).Text
        datencabezado.Recordset.Fields("recetaid") = Text1(3).Text
    End If
    datencabezado.Recordset.Fields("vendedorid") = DataGrid1.Columns(0).Text
    datencabezado.Recordset.Fields("vendedor") = DataGrid1.Columns(2).Text
    datencabezado.Recordset.Fields("detalle") = Text1(5).Text
    datencabezado.Recordset.Fields("nota") = Text15.Text
    datencabezado.Recordset.Fields("cotizacion") = 1
    datencabezado.Recordset.Fields("listadeprecioid") = DataCombo1.BoundText
    datencabezado.Recordset.Fields("tipodepagoid") = DataCombo3.BoundText
    datencabezado.Recordset.Fields("tipodefactura") = Text1(2).Text
    
    datencabezado.Recordset.Fields("tipodefacturacionid") = "MA"
    datencabezado.Recordset.Fields("fechadeentrega") = DateValue(Text3.Text) + TimeValue(Str(Time))
    datencabezado.Recordset.Fields("domicilioid") = Text1(3).Text
    datencabezado.Recordset.Fields("domicilio_id") = DataGrid2.Columns("domicilio_id").Text
    datencabezado.Recordset.Fields("domiciliodeentregaid") = DataGrid2.Columns("domicilio_id").Text
        datencabezado.Recordset.Fields("generada") = "False"
    datencabezado.Recordset.Fields("importado") = "False"
    datencabezado.Recordset.Fields("nombrepc") = Environ("computername")
    datencabezado.Recordset.Fields("responsabilidad") = DataGrid2.Columns(16).Text
    datencabezado.Recordset.Fields("transferido") = "False"
    
    
    datencabezado.Recordset.Fields("cprov1") = grilla.TextMatrix(0, 8)
    datencabezado.Recordset.Fields("cprov2") = grilla.TextMatrix(0, 9)
    datencabezado.Recordset.Fields("cprov3") = grilla.TextMatrix(0, 10)
    datencabezado.Recordset.Fields("cprov4") = grilla.TextMatrix(0, 11)
    datencabezado.Recordset.Fields("cprov5") = grilla.TextMatrix(0, 12)
    datencabezado.Recordset.Fields("cprov6") = grilla.TextMatrix(0, 13)
    
    
        '** Establene numero de Presupuesto
'    xnumerador = "Presupuesto de Venta " + datparametros.Recordset.Fields("sucursal")
    
'    datcola.RecordSource = "SELECT p.ID AS ID, p.NOMBRE AS puntero, n.NUMERO AS NUMERO, p.CARACTERESPREFIJO AS puntoventa, p.NUMERO_ID " & _
'                           "FROM V_NUMERADOR_ AS p LEFT OUTER JOIN V_NUMERO_ AS n ON p.NUMERO_ID = n.ID " & _
'                           "WHERE     (p.ACTIVESTATUS <> 2) AND (p.NOMBRE = '" & xnumerador & "') "
'    datcola.Refresh
        
        datencabezado.Recordset.Fields("numerodefactura") = Text17.Text
        datencabezado.Recordset.Fields("presupuestobase") = Text17.Text
'        xnumero = datcola.Recordset.Fields("numero")
'        xidnumero = datcola.Recordset.Fields("numero_id")
'        datencabezado.Recordset.Fields("puntodeventa") = datcola.Recordset.Fields("puntoventa")
'        datcola.RecordSource = "Select * from numero with(readpast) where id = '" & xidnumero & "'"
'        datcola.Refresh
'        datcola.Recordset.Fields("numero") = xnumero + 1
'        datcola.Recordset.UpdateBatch adAffectCurrent
    
    '** Fin de asignacion de numero a Presupuesto

    datencabezado.Recordset.UpdateBatch adAffectCurrent
    presupuestobase = datencabezado.Recordset.Fields("id")
    
    
'--- Graba Items
    
    For X = 1 To xlineasmax + xlinesadd
        If grilla.TextMatrix(1, 0) = "" Then
          mensa = MsgBox("No puede Grabar este Presupuesto sin Items", vbCritical, "Error")
          Exit Sub
        End If
        If grilla.TextMatrix(X, 0) = "" Then Exit For
        
        datitems.Recordset.AddNew
        datitems.Recordset.Fields("claveprimaria") = presupuestobase
        datitems.Recordset.Fields("idproducto") = grilla.TextMatrix(X, 0)
        datitems.Recordset.Fields("codigoproducto") = grilla.TextMatrix(X, 1)
        datitems.Recordset.Fields("nombre_producto") = grilla.TextMatrix(X, 2)
        datitems.Recordset.Fields("cantidadproducto") = grilla.TextMatrix(X, 3)
        datitems.Recordset.Fields("unidaddemedidaid") = grilla.TextMatrix(X, 4)
        
        datitems.Recordset.Fields("preciou") = Round(Val(Format(Replace(grilla.TextMatrix(X, 5), ",", ""), "##0.00")), 3)
        datitems.Recordset.Fields("bonificacionitem") = Round(Val(Format(grilla.TextMatrix(X, 6), "##0.00")), 3)
        datitems.Recordset.Fields("porcentaje") = Round(Val(grilla.TextMatrix(X, 7)), 3)
        If grilla.TextMatrix(X, 8) <> "" Then
            datitems.Recordset.Fields("cprecioprov1") = Round(Val(Format(Replace(grilla.TextMatrix(X, 8), ",", ""), "##0.00")), 3)
            datitems.Recordset.Fields("cmarcap1") = grilla.TextMatrix(X, 9)
        End If
        If grilla.TextMatrix(X, 10) <> "" Then
            datitems.Recordset.Fields("cprecioprov2") = Round(Val(Format(Replace(grilla.TextMatrix(X, 10), ",", ""), "##0.00")), 3)
            datitems.Recordset.Fields("cmarcap2") = grilla.TextMatrix(X, 11)
        End If
        If grilla.TextMatrix(X, 12) <> "" Then
            datitems.Recordset.Fields("cprecioprov3") = Round(Val(Format(Replace(grilla.TextMatrix(X, 12), ",", ""), "##0.00")), 3)
            datitems.Recordset.Fields("cmarcap3") = grilla.TextMatrix(X, 13)
        End If
        If grilla.TextMatrix(X, 14) <> "" Then
            datitems.Recordset.Fields("cprecioprov4") = Round(Val(Format(Replace(grilla.TextMatrix(X, 14), ",", ""), "##0.00")), 3)
            datitems.Recordset.Fields("cmarcap4") = grilla.TextMatrix(X, 15)
        End If
        If grilla.TextMatrix(X, 16) <> "" Then
            datitems.Recordset.Fields("cprecioprov5") = Round(Val(Format(Replace(grilla.TextMatrix(X, 16), ",", ""), "##0.00")), 3)
            datitems.Recordset.Fields("cmarcap5") = grilla.TextMatrix(X, 17)
        End If
        If grilla.TextMatrix(X, 18) <> "" Then
            datitems.Recordset.Fields("cprecioprov6") = Round(Val(Format(Replace(grilla.TextMatrix(X, 18), ",", ""), "##0.00")), 3)
            datitems.Recordset.Fields("cmarcap6") = grilla.TextMatrix(X, 19)
        End If
        
        datitems.Recordset.UpdateBatch adAffectCurrent
    Next X
    
'******* Graba Cola importar
      If 1 = 2 Then
        datcola.RecordSource = "Select * from ud_ezi_cola_importar where id = 0"
        datcola.Refresh
        
        datcola.Recordset.AddNew
        datcola.Recordset.Fields("id_encabezado") = presupuestobase
        datcola.Recordset.Fields("tipodedocumentoid") = datparametros.Recordset.Fields("idpresupuesto")
        datcola.Recordset.Fields("unidadoperativaid") = datparametros.Recordset.Fields("target")
        datcola.Recordset.Fields("fecha_hora") = DateValue(Text3.Text) + TimeValue(Str(Time))
        
        datcola.Recordset.UpdateBatch adAffectCurrent
      End If
    
'******* Coloca Presupuesto en importado = true para no ser tomado para otra comparativa
        datcola.RecordSource = "Select * from ud_ezi_puntodeventa_encabezado where id = " & xidpre & "  "
        datcola.Refresh
        datcola.Recordset.Fields("importado") = True
        datcola.Recordset.UpdateBatch adAffectCurrent
'**********************************************************************************+
    
    idpresupuesto = presupuestobase
    tipopresupuesto = Text1(2).Text
    
    mensa = MsgBox("Nro de Comparativa: " + datencabezado.Recordset.Fields("numerodefactura"), vbInformation, "Grabado Correctamente")
End If

If menu = 6 Then
    datencabezado.Recordset.Fields("detalle") = Text1(5).Text
    datencabezado.Recordset.Fields("nota") = Text15.Text
    
    datencabezado.Recordset.Fields("cprov1") = grilla.TextMatrix(0, 8)
    datencabezado.Recordset.Fields("cprov2") = grilla.TextMatrix(0, 9)
    datencabezado.Recordset.Fields("cprov3") = grilla.TextMatrix(0, 10)
    datencabezado.Recordset.Fields("cprov4") = grilla.TextMatrix(0, 11)
    datencabezado.Recordset.Fields("cprov5") = grilla.TextMatrix(0, 12)
    datencabezado.Recordset.Fields("cprov6") = grilla.TextMatrix(0, 13)
    
    datencabezado.Recordset.UpdateBatch adAffectCurrent
    presupuestobase = datencabezado.Recordset.Fields("id")
    
    
'--- Modifica Items
    
    For X = 1 To xlineasmax + xlinesadd
        If grilla.TextMatrix(1, 0) = "" Then
          mensa = MsgBox("No puede Grabar este Presupuesto sin Items", vbCritical, "Error")
          Exit Sub
        End If
        If grilla.TextMatrix(X, 0) = "" Then Exit For
        
        datitems.RecordSource = "select * from ud_ezi_puntodeventa_detalle_presu where id = " & grilla.TextMatrix(X, 19) & " "
        datitems.Refresh
        
        datitems.Recordset.Fields("porcentaje") = Round(Val(grilla.TextMatrix(X, 7)), 3)
        If grilla.TextMatrix(X, 8) <> "" Then
            datitems.Recordset.Fields("cprecioprov1") = Round(Val(Replace(grilla.TextMatrix(X, 8), ",", "")), 3)
            datitems.Recordset.Fields("cmarcap1") = grilla.TextMatrix(X, 9)
        End If
        If grilla.TextMatrix(X, 10) <> "" Then
            datitems.Recordset.Fields("cprecioprov2") = Round(Val(Replace(grilla.TextMatrix(X, 10), ",", "")), 3)
            datitems.Recordset.Fields("cmarcap2") = grilla.TextMatrix(X, 11)
        End If
        If grilla.TextMatrix(X, 12) <> "" Then
            datitems.Recordset.Fields("cprecioprov3") = Round(Val(Replace(grilla.TextMatrix(X, 12), ",", "")), 3)
            datitems.Recordset.Fields("cmarcap3") = grilla.TextMatrix(X, 13)
        End If
        If grilla.TextMatrix(X, 14) <> "" Then
            datitems.Recordset.Fields("cprecioprov4") = Round(Val(Replace(grilla.TextMatrix(X, 14), ",", "")), 3)
            datitems.Recordset.Fields("cmarcap4") = grilla.TextMatrix(X, 15)
        End If
        If grilla.TextMatrix(X, 16) <> "" Then
            datitems.Recordset.Fields("cprecioprov5") = Round(Val(Replace(grilla.TextMatrix(X, 16), ",", "")), 3)
            datitems.Recordset.Fields("cmarcap5") = grilla.TextMatrix(X, 17)
        End If
        If grilla.TextMatrix(X, 18) <> "" Then
            datitems.Recordset.Fields("cprecioprov6") = Round(Val(Replace(grilla.TextMatrix(X, 18), ",", "")), 3)
            datitems.Recordset.Fields("cmarcap6") = grilla.TextMatrix(X, 19)
        End If
        
        datitems.Recordset.UpdateBatch adAffectCurrent
    Next X
    
    mensa = MsgBox("Nro de Comparativa: " + datencabezado.Recordset.Fields("numerodefactura"), vbInformation, "Modificado Correctamente")
End If


    Call Cancelar_Click
    
'    Call imprimepresupuesto_Click
    
    
    
    
    
    'Unload Me
    
    
    
Exit Sub
errorgrabar:
    mensa = MsgBox("No se pudo registrar la información", vbCritical, "Error !!")






End Sub

Private Sub List1_KeyPress(KeyAscii As Integer)
On Error GoTo fuera

    If KeyAscii = 13 Then
        KeyAscii = 0
        Text1(4).Text = List1.ListIndex + 1
        Text1(5).SetFocus
    End If

fuera:
End Sub

Private Sub List1_LostFocus()

    List1.Visible = False

End Sub


Private Sub manual_Click()


    For X = frmpesada_cania.Width To 13000 Step 100
            frmpesada_cania.Width = X
    Next X
    Text4.SetFocus

End Sub

Private Sub grabaregistros_Click()
'On Error GoTo errorgrabar


End Sub

Private Sub grfacturactacte_Click()

MsgBox "Grabando Factura ctacte"


End Sub

Private Sub grremitoctacte_Click()

End Sub

Private Sub grremito_Click()

    MsgBox "Imprimiendo Remito"

End Sub

Private Sub grilla_AfterEdit(ByVal Row As Long, ByVal Col As Long)
On Error Resume Next

If grilla.Col = 8 Or grilla.Col = 10 Or grilla.Col = 12 Or grilla.Col = 14 Or grilla.Col = 16 Or grilla.Col = 18 Then

  xcol = grilla.Col
  grilla.Col = 6
  xcolor = grilla.CellBackColor
  grilla.Col = xcol

  xmenor = Val(Replace(Format(grilla.TextMatrix(Row, 5), "##0.00"), ",", ""))
  xmenor2 = Val(Replace(Format(grilla.TextMatrix(Row, 8), "##0.00"), ",", ""))
  For X = 8 To 18 Step 2
    xvalor = Val(Replace(Format(grilla.TextMatrix(Row, X), "##0.00"), ",", ""))
    grilla.Col = X
    
    If xvalor <= xmenor2 And xvalor <> 0 Then
        xmenor2 = xvalor
        xcolmenor = X
    End If
    
    If xvalor < xmenor And xvalor <> 0 Then
        xmenor = Val(Replace(Format(grilla.TextMatrix(Row, X), "##0.00"), ",", ""))
        xcolmenor = X
    End If
    grilla.CellBackColor = xcolor

  Next X

  grilla.Col = xcolmenor
  grilla.CellBackColor = &HFF00&

  grilla.Col = xcol
  
    If xmenor2 > xmenor Then xmenor = xmenor2
  
    grilla.TextMatrix(Row, Col) = Format(grilla.TextMatrix(Row, Col), "###,##0.00`")
    xporcentaje = Round(((Val(Replace(grilla.TextMatrix(Row, 5), ",", "")) / xmenor) * 100) - 100, 2)
    grilla.TextMatrix(Row, 7) = Format(xporcentaje, "###,##0.00`")
    grilla.TextMatrix(grilla.Row, grilla.Col) = Format(grilla.TextMatrix(grilla.Row, grilla.Col), "###,##0.00")
    
    grilla.Col = 7
    If xporcentaje < 0 Then
        grilla.CellBackColor = &HFF00&
    Else
        grilla.CellBackColor = &H8080FF
    End If
    grilla.Col = xcol


End If


End Sub

Private Sub grilla_DblClick()

  If grilla.Col = 8 Or grilla.Col = 10 Or grilla.Col = 12 Or grilla.Col = 14 Or grilla.Col = 16 Or grilla.Col = 18 Then
    lista_proveedores_calipso.Show
    lista_proveedores_calipso.Text1.Text = grilla.TextMatrix(0, grilla.Col)
    lista_proveedores_calipso.Text1.SetFocus
  End If
 
  If grilla.Col = 9 Or grilla.Col = 11 Or grilla.Col = 13 Or grilla.Col = 15 Or grilla.Col = 17 Or grilla.Col = 19 Then
    lista_marcas.Show
    lista_marcas.Text1.Text = grilla.TextMatrix(grilla.Row, grilla.Col)
    lista_marcas.Text1.SetFocus
  End If

End Sub

Private Sub grilla_EnterCell()

    xcolu = grilla.ColPos(grilla.Col) + grilla.Left + Picture1.Left + 50

    Label3.Caption = grilla.TextMatrix(0, grilla.Col)
    Label3.Left = xcolu
    If grilla.TextMatrix(grilla.Row, 1) = "" Then
        grilla.Editable = False
    Else
        If grilla.Col = 7 Or grilla.Col = 9 Or grilla.Col = 11 Or grilla.Col = 13 Or grilla.Col = 15 Or grilla.Col = 17 Or grilla.Col = 19 Then
            grilla.Editable = False
        Else
            grilla.Editable = True
        End If
    End If


End Sub

Private Sub grilla1_DblClick()



End Sub

Private Sub grilla1_EnterCell()

End Sub

Private Sub imprimepresupuesto_Click()
On Error GoTo fuera
Dim crApp As New CRAXDRT.Application
Dim crReport As New CRAXDRT.Report
Dim tabla As String
Dim ruta As String

If datparametros.Recordset.Fields("imprimemanual") = "N" Then Exit Sub

reporte.SQL = "SELECT v_ezi_pos_presupuesto.NUMERODOCUMENTO, v_ezi_pos_presupuesto.FECHAEMISION, v_ezi_pos_presupuesto.cod_cliente, v_ezi_pos_presupuesto.cliente, v_ezi_pos_presupuesto.CUIT, v_ezi_pos_presupuesto.CODPOS, v_ezi_pos_presupuesto.provincia, v_ezi_pos_presupuesto.vendedor, v_ezi_pos_presupuesto.detalle, v_ezi_pos_presupuesto.tipopago, v_ezi_pos_presupuesto.codigoproducto, v_ezi_pos_presupuesto.nombre_producto, v_ezi_pos_presupuesto.cantidadproducto, v_ezi_pos_presupuesto.nota, v_ezi_pos_presupuesto.condiva, v_ezi_pos_presupuesto.ciudad, v_ezi_pos_presupuesto.TIPOVENTA, v_ezi_pos_presupuesto.SIMBOLO, v_ezi_pos_presupuesto.CODVENDEDOR, v_ezi_pos_presupuesto.preciusiniva, v_ezi_pos_presupuesto.subtotalsiniva, v_ezi_pos_presupuesto.impbonifsiniva, v_ezi_pos_presupuesto.percepiibb, v_ezi_pos_presupuesto.perceptem, v_ezi_pos_presupuesto.totaltr, v_ezi_pos_presupuesto.importeiva21, v_ezi_pos_presupuesto.importeiva105 FROM MMOSSE.dbo.v_ezi_pos_presupuesto v_ezi_pos_presupuesto " & _
              " where v_ezi_pos_presupuesto.id = " & idpresupuesto & " order by v_ezi_pos_presupuesto.iditem"


'reporte.SQL = "SELECT v_ezi_pos_presupuesto.id, v_ezi_pos_presupuesto.NUMERODOCUMENTO, v_ezi_pos_presupuesto.FECHAEMISION, v_ezi_pos_presupuesto.cod_cliente, v_ezi_pos_presupuesto.cliente, v_ezi_pos_presupuesto.CUIT, v_ezi_pos_presupuesto.CALLE, v_ezi_pos_presupuesto.CODPOS, v_ezi_pos_presupuesto.provincia, v_ezi_pos_presupuesto.detalle, v_ezi_pos_presupuesto.tipopago, v_ezi_pos_presupuesto.codigoproducto, v_ezi_pos_presupuesto.nombre_producto, v_ezi_pos_presupuesto.cantidadproducto, v_ezi_pos_presupuesto.nota, v_ezi_pos_presupuesto.condiva, v_ezi_pos_presupuesto.ciudad, v_ezi_pos_presupuesto.TIPOVENTA, v_ezi_pos_presupuesto.SIMBOLO, v_ezi_pos_presupuesto.CODVENDEDOR, v_ezi_pos_presupuesto.preciusiniva, v_ezi_pos_presupuesto.subtotalsiniva, v_ezi_pos_presupuesto.impbonifsiniva, v_ezi_pos_presupuesto.nroremito, v_ezi_pos_presupuesto.percepiibb, v_ezi_pos_presupuesto.perceptem, v_ezi_pos_presupuesto.totaltr, " & _
'              "v_ezi_pos_presupuesto.importeiva21, v_ezi_pos_presupuesto.importeiva105, v_ezi_pos_presupuesto.iditem " & _
'              "FROM  MMOSSE.dbo.v_ezi_pos_presupuesto v_ezi_pos_presupuesto " & _
'              "where v_ezi_pos_presupuesto.id = " & idpresupuesto & " order by v_ezi_pos_presupuesto.iditem"

tabla = reporte.SQL

With CrystalReporte
    .PrinterCollation = crptCollated
    If tipopresupuesto = "A" Then
     If login.usuarioactivo = "admin" Then
        .ReportFileName = App.Path & "\CotizacionA.rpt"
     Else
        .ReportFileName = App.Path & "\CotizacionAventas.rpt"
     End If
    Else
       If login.usuarioactivo = "admin" Then
        .ReportFileName = App.Path & "\CotizacionB.rpt"
       Else
        .ReportFileName = App.Path & "\CotizacionBventas.rpt"
       End If
    End If
   
    .Connect = login.conexionreporte
    .DiscardSavedData = True
    .RetrieveDataFiles
    .ReportSource = 0
    .SQLQuery = tabla
    If login.usuarioactivo = "admin" Then
        .Destination = crptToWindow
    '    .Destination = crptToPrinter
        .PrintFileType = crptCrystal
        .WindowState = crptMaximized
        .Action = 1
    Else
        .Destination = crptToPrinter
        .PrintFileType = crptCrystal
        .WindowState = crptMaximized
        .Formulas(0) = "copia="" ORIGINAL """
        .Action = 1
        .Formulas(0) = "copia="" DUPLICADO """
        .Action = 1
    End If
    
    
End With

Exit Sub

fuera:
    
    MsgBox "Reporte de Presupuesto no Encontado, o error de configuracion de reporte", vbCritical, "Error"


End Sub

Private Sub KewlButtons1_Click()
On Error Resume Next
        For X = 0 To 10
            grilla.Col = X
            grilla.Text = ""
        Next X
        
        For X = grilla.Row + 1 To 299
            For Y = 0 To 13
                grilla.Col = Y
                grilla.Row = X
                xcampo = grilla.Text
                grilla.Row = X - 1
                grilla.Text = xcampo
                Text2.Text = ""
                Combo1.Clear
            Next Y
        Next X

    Call calcula_Click
    'Call ubicatextogrilla_Click
    
End Sub

Private Sub KewlButtons2_Click()
On Error GoTo fuera
Dim crApp As New CRAXDRT.Application
Dim crReport As New CRAXDRT.Report
Dim tabla As String
Dim ruta As String

    datencabezado.RecordSource = "SELECT * from ud_ezi_puntodeventa_encabezado_temporal"
    datencabezado.Refresh
    datitems.RecordSource = "Select * from ud_ezi_puntodeventa_detalle_presu_temporal"
    datitems.Refresh
    If datencabezado.Recordset.EOF = False Then
      Do While Not datencabezado.Recordset.EOF
        datencabezado.Recordset.Delete adAffectCurrent
        datencabezado.Recordset.MoveNext
      Loop
    End If
    If datitems.Recordset.EOF = False Then
      Do While Not datitems.Recordset.EOF
        datitems.Recordset.Delete adAffectCurrent
        datitems.Recordset.MoveNext
      Loop
    End If
    
    datencabezado.RecordSource = "SELECT MAX(CONVERT(decimal, isnull(claveprimaria,0))) AS claveprimaria, numeradorinterno FROM ud_ezi_puntodeventa_encabezado_temporal WITH (readpast) " & _
                                 "GROUP BY numeradorinterno HAVING (numeradorinterno = 'Comparativa de Precios')"
    datencabezado.Refresh
    
    If datencabezado.Recordset.EOF = True Then
        xclaveprimaria = 1
    Else
        xclaveprimaria = datencabezado.Recordset.Fields("claveprimaria") + 1
    End If
    

    datencabezado.RecordSource = "SELECT * From ud_ezi_puntodeventa_encabezado_temporal where id =0 "
    datencabezado.Refresh
    datitems.RecordSource = "Select * from ud_ezi_puntodeventa_detalle_presu_temporal where id = 0"
    datitems.Refresh
    
    
    
    datencabezado.Recordset.AddNew
    datencabezado.Recordset.Fields("claveprimaria") = xclaveprimaria


    datencabezado.Recordset.Fields("numeradorinterno") = "Presupuesto de Venta temporal"
    datencabezado.Recordset.Fields("fechadelcomprobante") = DateValue(Text3.Text) + TimeValue(Str(Time))
    datencabezado.Recordset.Fields("sucursal") = datparametros.Recordset.Fields("sucursal")
    datencabezado.Recordset.Fields("clienteid") = DataGrid2.Columns(0).Text
    datencabezado.Recordset.Fields("cliente") = DataGrid2.Columns(2).Text
    If Text1(6).Text <> "" Then
        datencabezado.Recordset.Fields("cliente") = Text1(6).Text
        datencabezado.Recordset.Fields("recetaid") = Text1(3).Text
    End If
    datencabezado.Recordset.Fields("vendedorid") = DataGrid1.Columns(0).Text
    datencabezado.Recordset.Fields("vendedor") = DataGrid1.Columns(2).Text
    datencabezado.Recordset.Fields("detalle") = Text1(5).Text
    datencabezado.Recordset.Fields("nota") = Text15.Text
    datencabezado.Recordset.Fields("cotizacion") = 1
    datencabezado.Recordset.Fields("listadeprecioid") = DataCombo1.BoundText
    datencabezado.Recordset.Fields("tipodepagoid") = DataCombo3.BoundText
    datencabezado.Recordset.Fields("tipodefactura") = Text1(2).Text
    
    datencabezado.Recordset.Fields("tipodefacturacionid") = "MA"
    datencabezado.Recordset.Fields("fechadeentrega") = DateValue(Text3.Text) + TimeValue(Str(Time))
    If Text13.Text = "" Then Text13.Text = 0
    datencabezado.Recordset.Fields("recargo") = Round(Text13.Text, 2)
    datencabezado.Recordset.Fields("tiporecargo") = "$"
    If Text11.Text = "" Then Text11.Text = 0
    datencabezado.Recordset.Fields("bonificacion") = Round(Text11.Text, 2)
    datencabezado.Recordset.Fields("tipobonificacion") = "$"
    datencabezado.Recordset.Fields("importeglobal") = Round(Text4.Text, 2)
'    datencabezado.Recordset.Fields("numerodefactura") = xclaveprimaria
    datencabezado.Recordset.Fields("domicilioid") = Text1(3).Text
    datencabezado.Recordset.Fields("domicilio_id") = DataGrid2.Columns("domicilio_id").Text
    datencabezado.Recordset.Fields("domiciliodeentregaid") = DataGrid2.Columns("domicilio_id").Text
    datencabezado.Recordset.Fields("subtotalsiniva") = Round(Text5.Text, 2)
    datencabezado.Recordset.Fields("totaliva") = Round(Text6.Text, 2) + Round(Text7.Text, 2)
    datencabezado.Recordset.Fields("generada") = "False"
    datencabezado.Recordset.Fields("importado") = "False"
    datencabezado.Recordset.Fields("nombrepc") = Environ("computername")
    datencabezado.Recordset.Fields("responsabilidad") = DataGrid2.Columns(16).Text
    datencabezado.Recordset.Fields("transferido") = "False"
    
    datencabezado.Recordset.Fields("percepiibb") = Round(Text9.Text, 2)
    datencabezado.Recordset.Fields("perceptem") = Round(Text8.Text, 2)
    datencabezado.Recordset.Fields("totaltr") = Round(Text4.Text, 2)
    If agregaproducto.Enabled = False Then
        datencabezado.Recordset.Fields("servicio") = "S"
        mensa = MsgBox("Recepciona Equipo para Mantenimiento", vbYesNo, "Atencion")
        If mensa = vbYes Then
            datencabezado.Recordset.Fields("recepcionaservicio") = "S"
        End If
    End If
    
    
        '** Establene numero de Presupuesto
    xnumerador = "Presupuesto de Venta " + datparametros.Recordset.Fields("sucursal")
    
    datcola.RecordSource = "SELECT p.ID AS ID, p.NOMBRE AS puntero, n.NUMERO AS NUMERO, p.CARACTERESPREFIJO AS puntoventa, p.NUMERO_ID " & _
                           "FROM V_NUMERADOR_ AS p LEFT OUTER JOIN V_NUMERO_ AS n ON p.NUMERO_ID = n.ID " & _
                           "WHERE     (p.ACTIVESTATUS <> 2) AND (p.NOMBRE = '" & xnumerador & "') "
    datcola.Refresh
        
        datencabezado.Recordset.Fields("numerodefactura") = datcola.Recordset.Fields("numero")
        xnumero = datcola.Recordset.Fields("numero")
        xidnumero = datcola.Recordset.Fields("numero_id")
        datencabezado.Recordset.Fields("puntodeventa") = datcola.Recordset.Fields("puntoventa")
    
    '** Fin de asignacion de numero a Presupuesto

    datencabezado.Recordset.UpdateBatch adAffectCurrent
    presupuestobase = datencabezado.Recordset.Fields("id")
    
    
'--- Graba Items
    
    For X = 1 To xlineasmax + xlinesadd
        If grilla.TextMatrix(1, 0) = "" Then
          mensa = MsgBox("No puede Grabar este Presupuesto sin Items", vbCritical, "Error")
          Exit Sub
        End If
        If grilla.TextMatrix(X, 0) = "" Then Exit For
        
        datitems.Recordset.AddNew
        datitems.Recordset.Fields("claveprimaria") = presupuestobase
        datitems.Recordset.Fields("idproducto") = grilla.TextMatrix(X, 0)
        datitems.Recordset.Fields("codigoproducto") = grilla.TextMatrix(X, 1)
        datitems.Recordset.Fields("nombre_producto") = grilla.TextMatrix(X, 2)
        datitems.Recordset.Fields("cantidadproducto") = grilla.TextMatrix(X, 3)
        datitems.Recordset.Fields("unidaddemedidaid") = grilla.TextMatrix(X, 4)
        datitems.Recordset.Fields("preciou") = Round(grilla.TextMatrix(X, 6), 3)
        datitems.Recordset.Fields("preciousiva") = Round(grilla.TextMatrix(X, 5), 3)
        datitems.Recordset.Fields("bonificacionitem") = grilla.TextMatrix(X, 9)
        If grilla.TextMatrix(X, 8) = "" Then grilla.TextMatrix(X, 8) = 0
        datitems.Recordset.Fields("importedebonificacion") = Round(grilla.TextMatrix(X, 8), 4)
        datitems.Recordset.Fields("subtotal") = Round(grilla.TextMatrix(X, 10), 3)
        datitems.Recordset.Fields("iva") = (Round(grilla.TextMatrix(X, 12), 4) - 1) * 100
        
        datitems.Recordset.UpdateBatch adAffectCurrent
    Next X
    

reporte.SQL = "SELECT v_ezi_pos_presupuesto.id, v_ezi_pos_presupuesto.NUMERODOCUMENTO, v_ezi_pos_presupuesto.FECHAEMISION, v_ezi_pos_presupuesto.cod_cliente, v_ezi_pos_presupuesto.cliente, v_ezi_pos_presupuesto.CUIT, v_ezi_pos_presupuesto.CALLE, v_ezi_pos_presupuesto.CODPOS, v_ezi_pos_presupuesto.provincia, v_ezi_pos_presupuesto.detalle, v_ezi_pos_presupuesto.tipopago, v_ezi_pos_presupuesto.codigoproducto, v_ezi_pos_presupuesto.nombre_producto, v_ezi_pos_presupuesto.cantidadproducto, v_ezi_pos_presupuesto.nota, v_ezi_pos_presupuesto.condiva, v_ezi_pos_presupuesto.ciudad, v_ezi_pos_presupuesto.TIPOVENTA, v_ezi_pos_presupuesto.SIMBOLO, v_ezi_pos_presupuesto.CODVENDEDOR, v_ezi_pos_presupuesto.preciusiniva, v_ezi_pos_presupuesto.subtotalsiniva, v_ezi_pos_presupuesto.impbonifsiniva, v_ezi_pos_presupuesto.nroremito, v_ezi_pos_presupuesto.percepiibb, v_ezi_pos_presupuesto.perceptem, v_ezi_pos_presupuesto.totaltr, " & _
              "v_ezi_pos_presupuesto.importeiva21, v_ezi_pos_presupuesto.importeiva105, v_ezi_pos_presupuesto.iditem " & _
              "FROM  dbo.v_ezi_pos_presupuesto v_ezi_pos_presupuesto " & _
              "where v_ezi_pos_presupuesto.id = " & presupuestobase & " and v_ezi_pos_presupuesto.numeradorinterno = 'Presupuesto de Venta Temporal'  order by v_ezi_pos_presupuesto.iditem"

tabla = reporte.SQL

With CrystalReporte
    .PrinterCollation = crptCollated
    If Text1(2).Text = "A" Then
        .ReportFileName = App.Path & "\CotizacionA.rpt"
    Else
        .ReportFileName = App.Path & "\CotizacionB.rpt"
    End If
   
    .Connect = login.conexionreporte
    .DiscardSavedData = True
    .RetrieveDataFiles
    .ReportSource = 0
    .SQLQuery = tabla
    .Destination = crptToWindow
'    .Destination = crptToPrinter
    .PrintFileType = crptCrystal
    .WindowState = crptMaximized
    
    .Action = 1
    
End With

Exit Sub

fuera:
    
    MsgBox "Reporte de Presupuesto no Encontado, o error de configuracion de reporte", vbCritical, "Error"




End Sub

Private Sub negro_Click()


    negro.Visible = False
    blanco.Visible = True
    tipofac = "CF"
    

End Sub

Private Sub salir_Click()

    Unload Me

End Sub



Private Sub Text1_GotFocus(Index As Integer)
    
    If Index = 1 Then
        If Text1(0).Text = "" Then mensa = MsgBox("Debe ingresar un vendedor", vbCritical, "Error")
    End If
    If Index = 5 Then
        If Text1(0).Text = "" Then mensa = MsgBox("Debe ingresar un vendedor", vbCritical, "Error")
        If Text1(1).Text = "" Then mensa = MsgBox("Debe ingresar un Cliente", vbCritical, "Error")
    End If


End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
'On Error Resume Next


    If KeyAscii = 13 Then
        Text1(Index).Text = UCase(Text1(Index).Text)
        KeyAscii = 0
        If Index = 0 Then
            Call bvendedor_Click
            If Text1(0).Text <> "" Then
                menu = 2
                If datvendedor.Recordset.RecordCount = 1 Then
                   lista_clavevendedor.Show
                End If
            End If
        End If
        If Index = 1 Then
            If Text1(1).Text = "" Then Text1(1).Text = "CONSUMIDOR FINAL"
            Call bclientes_Click
        End If
        If Index = 5 Then
            grilla.Col = 7
            grilla.Row = 1
            grilla.SetFocus
        End If
        If Index = 6 Then
            Text1(3).Text = ""
            Text1(3).SetFocus
        End If
        If Index = 3 Then
            Text1(3).Text = UCase(Text1(3).Text)
            Text1(5).SetFocus
        End If

        
    End If
    
End Sub

Private Sub Text1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)

If KeyCode = 114 Then
       Call bvendedor_Click
End If

If KeyCode = 116 Then
    If agregaproducto.Enabled = True Then
       Call agregaproducto_Click
    End If
End If

If KeyCode = 38 Then
    If Index = 5 Then Text1(1).SetFocus
    If Index = 1 Then Text1(0).SetFocus
End If

If KeyCode = 121 Then
       Call grabar_Click
End If

If KeyCode = 123 Then
       Call agregaservicios_Click
End If


End Sub

Private Sub Text1_LostFocus(Index As Integer)
On Error Resume Next
        If Index = 2 Then
            For X = 1 To Len(Text1(2).Text)
               car = Mid(Text1(2).Text, X, 1)
               If car = "-" Then
                  PVta = Right("0000" + Left(Text1(2).Text, X - 1), 4)
                  nu = Right("00000000" + Right(Text1(2).Text, Len(Text1(2).Text) - X), 8)
                  Text1(2).Text = PVta + nu
                  Exit For
               End If
               Text1(2).Text = Right("00000000" + Text1(2).Text, 8)
            Next X
        End If

End Sub

Private Sub Text10_GotFocus()

    Text10.SelStart = 0
    Text10.SelLength = Len(Text10.Text)

End Sub

Private Sub Text10_KeyPress(KeyAscii As Integer)
On Error Resume Next

If KeyAscii = 13 Then
        KeyAscii = 0
        Text11.Text = Format(0, "##0.00")
        Text14.Text = Format(0, "##0.00")
        Text10.Text = Format(Text10.Text, "##0.00")
        ximporteboniftotal = 0
        For X = 1 To xlineasmax + xlinesadd
              If grilla.TextMatrix(X, 1) = "" Then Exit For
              grilla.Row = X
              grilla.TextMatrix(X, 9) = Text10.Text
              If Val(Text10.Text) = 0 Then
                 grilla.TextMatrix(X, 8) = 0
              End If
              Call calcula_Click
              ximporteboniftotal = ximporteboniftotal + grilla.TextMatrix(X, 8)
        Next X
        Text11.Text = Format(ximporteboniftotal, "##0.00")
        Text15.SetFocus
End If


End Sub

Private Sub Text10_KeyUp(KeyCode As Integer, Shift As Integer)


If KeyCode = 116 Then
    If agregaproducto.Enabled = True Then
       Call agregaproducto_Click
    End If
End If

If KeyCode = 121 Then
       Call grabar_Click
End If

If KeyCode = 123 Then
       Call agregaservicios_Click
End If



End Sub

Private Sub Text11_GotFocus()


    Text11.SelStart = 0
    Text11.SelLength = Len(Text11.Text)

End Sub

Private Sub Text11_KeyPress(KeyAscii As Integer)
On Error Resume Next

If KeyAscii = 13 Then
        KeyAscii = 0
        Text11.Text = Format(Text11.Text, "##0.00")
        Text14.Text = Format(0, "##0.00")
        Text10.Text = Format(0, "##0.00")
        xvalorsubtotal = Round(Text5.Text, 10) + Round(Text6.Text, 10) + Round(Text7.Text, 10)
        xporcenboniftotal = 0
        For X = 1 To xlineasmax + xlinesadd
              If grilla.TextMatrix(X, 1) = "" Then Exit For
              grilla.Row = X
              grilla.TextMatrix(X, 8) = Round((grilla.TextMatrix(X, 10) / xvalorsubtotal) * Text11.Text, 2)
              If Val(Text11.Text) = 0 Then
                 grilla.TextMatrix(X, 9) = 0
              End If
              Call calcula_Click
       
        Next X
        Text10.Text = Format(grilla.TextMatrix(1, 9), "##0.00")
        Text15.SetFocus
End If

End Sub

Private Sub Text11_KeyUp(KeyCode As Integer, Shift As Integer)


If KeyCode = 116 Then
    If agregaproducto.Enabled = True Then
       Call agregaproducto_Click
    End If
End If

If KeyCode = 121 Then
       Call grabar_Click
End If

If KeyCode = 123 Then
       Call agregaservicios_Click
End If


End Sub

Private Sub Text12_GotFocus()
    Text12.SelStart = 0
    Text12.SelLength = Len(Text12.Text)

End Sub

Private Sub Text12_KeyPress(KeyAscii As Integer)

On Error Resume Next

If KeyAscii = 13 Then
        KeyAscii = 0
        
'        grilla.TextMatrix(x, 5) = grilla.TextMatrix(X, 13)
        Text13.Text = 0
        Call calcula_Click
        
        Text13.Text = Format(0, "##0.00")
        Text14.Text = Format(0, "##0.00")
        ximporteboniftotal = Round(Text16.Text, 20)
        For X = 1 To xlineasmax + xlinesadd
              If grilla.TextMatrix(X, 1) = "" Then Exit For
              grilla.Row = X
              
               grilla.TextMatrix(X, 6) = Round(grilla.TextMatrix(X, 6), 20) * ((Round(Text12.Text, 20) / 100) + 1)
              'grilla.TextMatrix(X, 6) = Round(grilla.TextMatrix(X, 14), 20) * ((Round(Text12.Text, 20) / 100) + 1)
              If Val(Text12.Text) = 0 Then
                 grilla.TextMatrix(X, 6) = grilla.TextMatrix(X, 14)
              End If
              Call calcula_Click
        Next X
        ximporteboniftotal = Round(Text16.Text, 20) - ximporteboniftotal
        If Val(Text12.Text) = 0 Then
            Text13.Text = Format(0, "##0.00")
        Else
            Text13.Text = Format(ximporteboniftotal, "###,##0.00")
        End If
        Text12.Text = Format(Text12.Text, "##0.00")
        Text15.SetFocus
End If



End Sub

Private Sub Text12_KeyUp(KeyCode As Integer, Shift As Integer)


If KeyCode = 116 Then
    If agregaproducto.Enabled = True Then
       Call agregaproducto_Click
    End If
End If

If KeyCode = 121 Then
       Call grabar_Click
End If

If KeyCode = 123 Then
       Call agregaservicios_Click
End If


End Sub

Private Sub Text13_GotFocus()


        Text13.SelStart = 0
        Text13.SelLength = Len(Text13.Text)



End Sub

Private Sub Text13_KeyPress(KeyAscii As Integer)

On Error Resume Next

If KeyAscii = 13 Then
        KeyAscii = 0
        Text14.Text = Format(0, "##0.00")
        Text13.Text = Format(Text13.Text, "##0.00")
        ximporteboniftotal = (Round(Text13.Text, 20) / Round(Text16.Text, 20)) * 100
        Text12.Text = ximporteboniftotal
        Text12.SetFocus
        SendKeys "{ENTER}", False
        Text2.Text = Format(Text2.Text, "##0.00")
'        Text15.SetFocus
End If


End Sub

Private Sub Text13_KeyUp(KeyCode As Integer, Shift As Integer)


If KeyCode = 116 Then
    If agregaproducto.Enabled = True Then
       Call agregaproducto_Click
    End If
End If

If KeyCode = 121 Then
       Call grabar_Click
End If

If KeyCode = 123 Then
       Call agregaservicios_Click
End If

End Sub

Private Sub Text14_GotFocus()


    Text14.SelStart = 0
    Text14.SelLength = Len(Text10.Text)

End Sub

Private Sub Text14_KeyPress(KeyAscii As Integer)

On Error Resume Next

If KeyAscii = 13 Then
        KeyAscii = 0
              
        xxiibb = Round(Text14.Text, 2) * Round(Text9.Text, 2) / Round(Text4.Text, 2)
        xxtem = Round(Text14.Text, 2) * Round(Text8.Text, 2) / Round(Text4.Text, 2)
        xxiva21 = Round(Text14.Text, 2) * Round(Text7.Text, 2) / Round(Text4.Text, 2)
        xxiva10 = Round(Text14.Text, 2) * Round(Text6.Text, 2) / Round(Text4.Text, 2)
        xxsubtotal1 = Round(Text14.Text, 2) - xxiibb - xxtem - xxiva21 - xxiva10
        xxsubtotal2 = xxsubtotal1 + xxiva21 + xxiva10
        xsigno = xxsubtotal2 - Round(Text16.Text, 2)
        xdiferencia = Abs(xxsubtotal2 - Round(Text16.Text, 2))
        
        If xsigno > 0 Then
                Text13.Text = xdiferencia
                Text13.SetFocus
                SendKeys "{ENTER}", False
                Text13.Text = Format(Text13.Text, "##0.00")

        Else
        
                Text11.Text = xdiferencia
                Text11.SetFocus
                SendKeys "{ENTER}", False
                Text11.Text = Format(Text11.Text, "##0.00")
                Text11.SetFocus
                
        End If

End If



End Sub

Private Sub Text14_KeyUp(KeyCode As Integer, Shift As Integer)


If KeyCode = 116 Then
    If agregaproducto.Enabled = True Then
       Call agregaproducto_Click
    End If
End If

If KeyCode = 121 Then
       Call grabar_Click
End If

If KeyCode = 123 Then
       Call agregaservicios_Click
End If


End Sub

Private Sub Text15_KeyUp(KeyCode As Integer, Shift As Integer)

If KeyCode = 116 Then
    If agregaproducto.Enabled = True Then
       Call agregaproducto_Click
    End If
End If

If KeyCode = 121 Then
       Call grabar_Click
End If

If KeyCode = 123 Then
       Call agregaservicios_Click
End If


End Sub

Private Sub Text17_KeyPress(KeyAscii As Integer)
On Error Resume Next

If KeyAscii = 13 Then
    If Text18.Text = "" Then
        KeyAscii = 0
        lista_presupuestos.Show
        lista_presupuestos.Text1.Text = Text17.Text
        lista_presupuestos.Text1.SetFocus
        SendKeys "{ENTER}", False
    Else
        Text18.Text = ""
        Call cargapresupuesto_Click
    End If
    
End If


End Sub

Private Sub Text2_GotFocus()

    If grilla.Col <> 2 Then
        Text2.Height = 315
        Text2.SelStart = 0
        Text2.SelLength = Len(Text2.Text)
    End If

    

End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
On Error Resume Next

If KeyAscii = 13 Then
        Text2.Text = UCase(Text2.Text)
        KeyAscii = 0
        grilla.Text = Text2.Text
        
        If grilla.Col > 6 Then
            grilla.TextMatrix(grilla.Row, grilla.Col) = Format(grilla.Text, "#,###,##0.00")
        End If
            
        If grilla.Col > 12 Then
                grilla.Col = 7
        End If
       
        grilla.Row = grilla.Row + 1
        DoEvents
        
    '    Call ubicatextogrilla_Click
End If

    

End Sub


Private Sub Text4_KeyPress(KeyAscii As Integer)
On Error Resume Next

    If KeyAscii = 13 Then
        KeyAscii = 0
        Text5.SetFocus
    End If

End Sub


Private Sub Text5_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        autorizar.SetFocus
    End If
    

End Sub

Private Sub Text2_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next

If KeyCode = 39 And grilla.Col < 12 Then   'And Text2.SelStart = Len(Text2.Text)
    Text2.Visible = False
    grilla.Col = grilla.Col + 1
    grilla.SetFocus
'    Call ubicatextogrilla_Click
End If


If KeyCode = 38 And grilla.Row > 0 Then
   If grilla.Row = 1 Then
     Text1(5).SetFocus
     Text2.Visible = False
     Exit Sub
   End If
    grilla.Row = grilla.Row - 1
    If grilla.RowIsVisible(grilla.Row) = False Then
        grilla.TopRow = grilla.TopRow - 1
    End If
    DoEvents
  '  Call ubicatextogrilla_Click

End If


If KeyCode = 40 And grilla.Row < xlineasmax + xlinesadd Then
    grilla.Row = grilla.Row + 1
    If grilla.RowIsVisible(grilla.Row) = False Then
        grilla.TopRow = grilla.TopRow + 1
    End If
    DoEvents
 '   Call ubicatextogrilla_Click
End If



If KeyCode = 37 And grilla.Col > 2 Then ' And Text2.SelStart = 0
    grilla.Col = grilla.Col - 1
    DoEvents
'    Call ubicatextogrilla_Click
End If


If KeyCode = 116 Then
    If agregaproducto.Enabled = True Then
       Call agregaproducto_Click
    End If
End If

    
If KeyCode = 121 Then
       Call grabar_Click
End If

If KeyCode = 123 Then
       Call agregaservicios_Click
End If


End Sub

Private Sub Timer1_Timer()

    If Label2.Visible = True Then
        Label2.Visible = False
    Else
        Label2.Visible = True
    End If
    

End Sub


Private Sub UM_Click()
On Error Resume Next
    
    xid = grilla.TextMatrix(grilla.Row, 0)
    If xid = "" Then Exit Sub
    
    Combo1.Visible = True
    Combo1.Clear
    
    
    
    
    datum.RecordSource = "SELECT     P.ID, P.CODIGO, P.UNIDADMEDIDA_ID, P.UNIDADMEDIDANOLINEAL_ID, V_UNIDADMEDIDA_.NOMBRE AS UMVTA, V_UNIDADMEDIDA__1.NOMBRE AS UMSTK, " & _
                         "ISNULL(V_ITEMCONVUNIDADESDEMEDIDA_.RELACIONCONVERSION, 1) AS RELACIONCONVERSION, V_UNIDADMEDIDA__2.NOMBRE AS UMDESTINO " & _
                         "FROM         V_UNIDADMEDIDA_ AS V_UNIDADMEDIDA__2 RIGHT OUTER JOIN " & _
                         "V_ITEMCONVUNIDADESDEMEDIDA_ ON V_UNIDADMEDIDA__2.ID = V_ITEMCONVUNIDADESDEMEDIDA_.UNIDADDESTINO_ID RIGHT OUTER JOIN " & _
                         "V_PRODUCTO_ AS P ON V_ITEMCONVUNIDADESDEMEDIDA_.UNIDADORIGEN_ID = P.UNIDADMEDIDANOLINEAL_ID AND " & _
                         "V_ITEMCONVUNIDADESDEMEDIDA_.BO_PLACE_ID = P.EQUIVALENCIASPARTICULARES_ID LEFT OUTER JOIN " & _
                         "V_UNIDADMEDIDA_ AS V_UNIDADMEDIDA__1 ON P.UNIDADMEDIDA_ID = V_UNIDADMEDIDA__1.ID LEFT OUTER JOIN " & _
                         "V_UNIDADMEDIDA_ ON P.UNIDADMEDIDANOLINEAL_ID = V_UNIDADMEDIDA_.ID " & _
                         "WHERE  P.ID = '" & xid & "'"
    datum.Refresh
    If datum.Recordset.EOF = False Then
      Combo1.AddItem datum.Recordset.Fields("UMVTA")
      xumvta = datum.Recordset.Fields("UMVTA")
      xlist = 1
      Do While Not datum.Recordset.EOF
        If datum.Recordset.Fields("UMSTK") <> datum.Recordset.Fields("UMVTA") Then
            Combo1.AddItem datum.Recordset.Fields("UMSTK")
            xfaconv(xlist) = datum.Recordset.Fields("relacionconversion")
            xlist = xlist + 1
        End If
        If datum.Recordset.Fields("UMDESTINO") <> datum.Recordset.Fields("UMVTA") And datum.Recordset.Fields("UMDESTINO") <> datum.Recordset.Fields("UMSTK") Then
            Combo1.AddItem datum.Recordset.Fields("UMDESTINO")
            xfaconv(xlist) = datum.Recordset.Fields("relacionconversion")
            xlist = xlist + 1
        End If
        
        datum.Recordset.MoveNext
       Loop
        
    Else
        Combo1.AddItem "Unidad"
    End If

Text2.Visible = False
    
Combo1.Width = grilla.ColWidth(grilla.Col) + 200
If grilla.Text = "" Then
    Combo1.ListIndex = 0
Else
    Combo1.Text = grilla.Text
End If


xfila = grilla.RowPos(grilla.Row) + grilla.Top + Picture1.Top + 30
xcolu = grilla.ColPos(grilla.Col) + grilla.Left + Picture1.Left + 30
Combo1.Left = xcolu
Combo1.Top = xfila

Combo1.SetFocus
    
    

End Sub
