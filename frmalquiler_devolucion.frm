VERSION 5.00
Object = "{F6125AB1-8AB1-11CE-A77F-08002B2F4E98}#2.0#0"; "MSRDC20.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{D18BBD1F-82BB-4385-BED3-E9D31A3E361E}#1.0#0"; "KewlButtonz.ocx"
Begin VB.Form frmalquiler_devolucion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DEVOLUCION DE BIENES ALQUILADOS"
   ClientHeight    =   8610
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15195
   Icon            =   "frmalquiler_devolucion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8610
   ScaleWidth      =   15195
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00C0FFFF&
      Height          =   360
      Left            =   4560
      OLEDragMode     =   1  'Automatic
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Top             =   2760
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton UM 
      Caption         =   "UM"
      Height          =   315
      Left            =   13440
      TabIndex        =   38
      Top             =   3000
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton calcula 
      Caption         =   "calcula"
      Height          =   315
      Left            =   13200
      TabIndex        =   37
      Top             =   2760
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton ubicatextogrilla 
      Caption         =   "ubicatextogrilla"
      Height          =   315
      Left            =   10440
      TabIndex        =   36
      Top             =   3120
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
      Left            =   6240
      MultiLine       =   -1  'True
      TabIndex        =   12
      Text            =   "frmalquiler_devolucion.frx":0442
      Top             =   2760
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Frame Frame2 
      Caption         =   "Nro. de Remito"
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
      TabIndex        =   27
      Top             =   7440
      Width           =   4095
      Begin VB.TextBox Text18 
         Height          =   285
         Left            =   3240
         TabIndex        =   58
         Text            =   "Text18"
         Top             =   480
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton cargapresupuesto 
         Caption         =   "cargapresupuesto"
         Height          =   315
         Left            =   2160
         TabIndex        =   57
         Top             =   120
         Visible         =   0   'False
         Width           =   1815
      End
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
         Height          =   675
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   56
         Top             =   240
         Width           =   3855
      End
   End
   Begin MSAdodcLib.Adodc datvendedor 
      Height          =   330
      Left            =   9720
      Top             =   6720
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
      Left            =   9720
      Top             =   6360
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
      Left            =   9720
      Top             =   6120
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
      Left            =   9720
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
      TabIndex        =   18
      Top             =   7440
      Width           =   4815
      Begin KewlButtonz.KewlButtons salir 
         Height          =   495
         Left            =   3480
         TabIndex        =   17
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
         MICON           =   "frmalquiler_devolucion.frx":0448
         PICN            =   "frmalquiler_devolucion.frx":0464
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
         TabIndex        =   15
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
         MICON           =   "frmalquiler_devolucion.frx":0FAE
         PICN            =   "frmalquiler_devolucion.frx":0FCA
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
         TabIndex        =   16
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
         MICON           =   "frmalquiler_devolucion.frx":2A4C
         PICN            =   "frmalquiler_devolucion.frx":2A68
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
      ScaleWidth      =   14835
      TabIndex        =   20
      Top             =   120
      Width           =   14895
      Begin VB.ComboBox Combo2 
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
         Left            =   7560
         TabIndex        =   68
         Top             =   2760
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
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
         Height          =   435
         Index           =   9
         Left            =   3000
         MaxLength       =   10
         TabIndex        =   66
         Top             =   1560
         Width           =   2055
      End
      Begin MSDataGridLib.DataGrid DataGrid2 
         Bindings        =   "frmalquiler_devolucion.frx":347A
         Height          =   615
         Left            =   1440
         TabIndex        =   31
         Top             =   6000
         Visible         =   0   'False
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1085
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
      Begin VB.CommandButton imprimeremito 
         Caption         =   "imprimeremito"
         Height          =   255
         Left            =   10200
         TabIndex        =   64
         Top             =   2760
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.TextBox Text19 
         Height          =   285
         Left            =   120
         TabIndex        =   63
         Top             =   840
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
         Height          =   435
         Index           =   8
         Left            =   7200
         MaxLength       =   500
         TabIndex        =   62
         Top             =   2040
         Width           =   6975
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
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
         Index           =   7
         Left            =   3000
         MaxLength       =   5
         TabIndex        =   0
         Top             =   2040
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
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
         Height          =   435
         Index           =   6
         Left            =   5160
         MaxLength       =   5
         TabIndex        =   59
         Top             =   1560
         Width           =   1335
      End
      Begin VB.TextBox Text16 
         Alignment       =   1  'Right Justify
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
         Left            =   12720
         Locked          =   -1  'True
         TabIndex        =   53
         TabStop         =   0   'False
         Top             =   6720
         Visible         =   0   'False
         Width           =   1815
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
         ForeColor       =   &H00800000&
         Height          =   2055
         Left            =   120
         TabIndex        =   50
         Top             =   5040
         Width           =   11175
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
            Height          =   1560
            Left            =   240
            MaxLength       =   500
            MultiLine       =   -1  'True
            TabIndex        =   14
            Top             =   360
            Width           =   10695
         End
         Begin VB.CommandButton Command6 
            Caption         =   "Observaciones"
            Height          =   255
            Left            =   240
            TabIndex        =   51
            TabStop         =   0   'False
            Top             =   120
            Width           =   1935
         End
      End
      Begin VB.TextBox Text9 
         Alignment       =   1  'Right Justify
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
         Left            =   12720
         Locked          =   -1  'True
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   7440
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.TextBox Text8 
         Alignment       =   1  'Right Justify
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
         Left            =   12720
         Locked          =   -1  'True
         TabIndex        =   48
         TabStop         =   0   'False
         Top             =   7080
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.TextBox Text7 
         Alignment       =   1  'Right Justify
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
         Left            =   12720
         Locked          =   -1  'True
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   6360
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.TextBox Text6 
         Alignment       =   1  'Right Justify
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
         Left            =   12720
         Locked          =   -1  'True
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   6000
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.TextBox Text5 
         Alignment       =   1  'Right Justify
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
         Left            =   12720
         Locked          =   -1  'True
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   5640
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.TextBox Text4 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   12720
         Locked          =   -1  'True
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   7920
         Visible         =   0   'False
         Width           =   1815
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grilla 
         Height          =   1695
         Left            =   120
         TabIndex        =   11
         Top             =   3240
         Width           =   14535
         _ExtentX        =   25638
         _ExtentY        =   2990
         _Version        =   393216
         BackColor       =   16777215
         Rows            =   50
         Cols            =   16
         ScrollBars      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   16
         _Band(0).GridLinesBand=   2
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
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
         Height          =   435
         Index           =   5
         Left            =   8400
         MaxLength       =   300
         TabIndex        =   8
         Top             =   1560
         Width           =   5775
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
         MaxLength       =   300
         MultiLine       =   -1  'True
         TabIndex        =   6
         Top             =   1080
         Width           =   7695
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
         Index           =   1
         Left            =   1440
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
         Left            =   12240
         TabIndex        =   32
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
         TabIndex        =   28
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
         TabIndex        =   19
         Top             =   120
         Width           =   855
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
         Index           =   0
         Left            =   4680
         MaxLength       =   50
         TabIndex        =   1
         Top             =   120
         Width           =   3375
      End
      Begin MSDataListLib.DataCombo DataCombo3 
         Bindings        =   "frmalquiler_devolucion.frx":3493
         Height          =   360
         Left            =   9120
         TabIndex        =   5
         Top             =   600
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   741
         _Version        =   393216
         Enabled         =   0   'False
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
         BCOL            =   -2147483645
         BCOLO           =   -2147483645
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmalquiler_devolucion.frx":34AD
         PICN            =   "frmalquiler_devolucion.frx":34C9
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
         BCOL            =   -2147483645
         BCOLO           =   -2147483645
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmalquiler_devolucion.frx":3A63
         PICN            =   "frmalquiler_devolucion.frx":3A7F
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
         Bindings        =   "frmalquiler_devolucion.frx":4019
         Height          =   375
         Left            =   4680
         TabIndex        =   30
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
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "frmalquiler_devolucion.frx":4033
         Height          =   360
         Left            =   11160
         TabIndex        =   7
         Top             =   1080
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   741
         _Version        =   393216
         Enabled         =   0   'False
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
         Top             =   2760
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         BTYPE           =   14
         TX              =   "&Agregar Item (F5)"
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
         BCOL            =   -2147483645
         BCOLO           =   -2147483645
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmalquiler_devolucion.frx":4051
         PICN            =   "frmalquiler_devolucion.frx":406D
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
         Top             =   2760
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         BTYPE           =   14
         TX              =   "&Eliminar Item"
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
         BCOL            =   -2147483645
         BCOLO           =   -2147483645
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmalquiler_devolucion.frx":4607
         PICN            =   "frmalquiler_devolucion.frx":4623
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
         TabIndex        =   54
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
         MICON           =   "frmalquiler_devolucion.frx":4BBD
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
         TabIndex        =   55
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
         MICON           =   "frmalquiler_devolucion.frx":4BD9
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
      Begin MSAdodcLib.Adodc datitempresup 
         Height          =   330
         Left            =   9480
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
      Begin MSAdodcLib.Adodc datcolaimportar 
         Height          =   330
         Left            =   10080
         Top             =   6840
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
         Caption         =   "datcolaimportar"
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
      Begin MSAdodcLib.Adodc datitemsnv 
         Height          =   330
         Left            =   10320
         Top             =   6480
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
         Caption         =   "datcolaimportar"
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
      Begin MSRDC.MSRDC reporte 
         Height          =   375
         Left            =   8520
         Top             =   7920
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
      Begin Crystal.CrystalReport CrystalReporte 
         Left            =   0
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowTitle     =   "Libro IVA Compras"
         PrintFileLinesPerPage=   60
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Requiere Reparacion:"
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
         Left            =   4800
         TabIndex        =   67
         Top             =   2760
         Width           =   2655
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha y hora de Retiro:"
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
         Left            =   120
         TabIndex        =   65
         Top             =   1560
         Width           =   3015
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Recibido por:"
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
         Index           =   18
         Left            =   5160
         TabIndex        =   61
         Top             =   2040
         Width           =   1815
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Hora de Devoluc.:"
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
         Index           =   17
         Left            =   120
         TabIndex        =   60
         Top             =   2040
         Width           =   2295
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Subtotal 2:"
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
         Left            =   10680
         TabIndex        =   52
         Top             =   6720
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Percep IIBB:"
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
         Left            =   9720
         TabIndex        =   47
         Top             =   7440
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Tem/Pyp:"
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
         Left            =   10920
         TabIndex        =   46
         Top             =   7080
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Iva 21%:"
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
         Left            =   10920
         TabIndex        =   43
         Top             =   6360
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Iva 10.5%:"
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
         Left            =   10920
         TabIndex        =   42
         Top             =   6000
         Visible         =   0   'False
         Width           =   1695
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
         Left            =   10680
         TabIndex        =   40
         Top             =   5640
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Retirado por:"
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
         Left            =   6720
         TabIndex        =   35
         Top             =   1560
         Width           =   1695
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Lista de Precio:"
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
         Height          =   375
         Index           =   8
         Left            =   9240
         TabIndex        =   34
         Top             =   1080
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
         TabIndex        =   33
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
         TabIndex        =   29
         Top             =   120
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Importe Total:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   10680
         TabIndex        =   26
         Top             =   7920
         Visible         =   0   'False
         Width           =   1935
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
         Left            =   11280
         TabIndex        =   25
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
         TabIndex        =   24
         Top             =   120
         Width           =   1695
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   14760
         Y1              =   2640
         Y2              =   2640
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
         TabIndex        =   23
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
         TabIndex        =   22
         Top             =   600
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
         TabIndex        =   21
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
Attribute VB_Name = "frmalquiler_devolucion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
Public xaltura As Integer
Public xlineasmax As Integer
Public xremito As Double
Public xidremitovta As Double
Public xid As Double
Public xidpre As Double



Private Sub agregaproducto_Click()
        If Text1(0).Text = "" Then
            mensa = MsgBox("Debe ingresar un vendedor", vbCritical, "Error")
            Exit Sub
        End If
        If Text1(1).Text = "" Then
            mensa = MsgBox("Debe ingresar un Cliente", vbCritical, "Error")
            Exit Sub
        End If

     menu = 1
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
                      
    
    For X = 1 To xlineasmax
        grilla.Col = 1
        grilla.Row = X
        If grilla.Text = "" Then
           xfila = X
           lista_productos_colon_alq.Show
           Exit For
        End If
    Next X
    


End Sub

Private Sub bclientes_Click()

If Text1(1).Text <> "" Then
   If Text19.Text = "" Then
    xbusqueda = "%" + Text1(1).Text + "%"
    xquery1 = "SELECT  ALIAS_0.ID AS ID,  ALIAS_0.CODIGO, ALIAS_3.NOMBRE AS RAZONSOCIAL, ALIAS_3.CUIT,ALIAS_6.CALLE, ISNULL(ALIAS_6.CALLE,'') + '-'+ ISNULL(V_CIUDAD_.NOMBRE,'') + '-'+ ISNULL(V_PROVINCIA_.NOMBRE,'') AS DOMICILIO, ALIAS_0.DENOMINACION,ALIAS_5.NOMBRE AS ZONA, " & _
              "ALIAS_7.NUMERO AS TELEFONO, ALIAS_8.DIRECCIONELECTRONICA AS MAIL, " & _
              "V_TIPOPAGO_.NOMBRE AS TP, V_TIPOPAGO_.OBSERVACION as TipoPago, V_TIPOPAGO_.ACTIVESTATUS AS ACTIVESTATUSTP, V_PROVINCIA_.NOMBRE AS Provincia, " & _
              "V_CIUDAD_.NOMBRE AS Ciudad, ALIAS_0.CODIGO + ' ' + ALIAS_3.CUIT + ' ' + ALIAS_3.NOMBRE AS CONCATENADO, V_EZI_POCICION_IVA_CLIENTES.CODIGO AS IVA, V_TIPOPAGO_.ID AS IDPAGO, ALIAS_0.DOMICILIOFACTURACION_ID as domicilio_id " & _
              "FROM V_PERSONA_ AS ALIAS_3 RIGHT OUTER JOIN V_CLIENTE AS ALIAS_0 WITH (READPAST) LEFT OUTER JOIN " & _
              "V_EZI_POCICION_IVA_CLIENTES ON ALIAS_0.ID = V_EZI_POCICION_IVA_CLIENTES.idcliente LEFT OUTER JOIN " & _
              "V_TIPOPAGO_ ON ALIAS_0.TIPOPAGO_ID = V_TIPOPAGO_.ID ON ALIAS_3.ID = ALIAS_0.ENTEASOCIADO_ID LEFT OUTER JOIN " & _
              "V_CIUDAD_ RIGHT OUTER JOIN V_DOMICILIO_ AS ALIAS_6 ON V_CIUDAD_.ID = ALIAS_6.CIUDAD_ID LEFT OUTER JOIN " & _
              "V_PROVINCIA_ ON ALIAS_6.PROVINCIA_ID = V_PROVINCIA_.ID ON ALIAS_0.DOMICILIOFACTURACION_ID = ALIAS_6.ID LEFT OUTER JOIN " & _
              "V_TELEFONO_ AS ALIAS_7 ON ALIAS_3.TELEFONOPRINCIPAL_ID = ALIAS_7.ID LEFT OUTER JOIN " & _
              "V_DIRECCIONELECTRONICA_ AS ALIAS_8 ON ALIAS_3.DIRECELECTRONICAPRINCIPAL_ID = ALIAS_8.ID LEFT OUTER JOIN " & _
              "V_ZONA_ AS ALIAS_5 ON ALIAS_0.ZONA_ID = ALIAS_5.ID " & _
              "WHERE     (ALIAS_0.ACTIVESTATUS <> 2) AND ALIAS_0.CODIGO + ' ' + ALIAS_3.CUIT + ' ' + ALIAS_3.NOMBRE + ' ' + ALIAS_0.DENOMINACION like '" & xbusqueda & "' order by ALIAS_3.NOMBRE "
   Else
    xquery1 = "SELECT  ALIAS_0.ID AS ID,  ALIAS_0.CODIGO, ALIAS_3.NOMBRE AS RAZONSOCIAL, ALIAS_3.CUIT,ALIAS_6.CALLE, ISNULL(ALIAS_6.CALLE,'') + '-'+ ISNULL(V_CIUDAD_.NOMBRE,'') + '-'+ ISNULL(V_PROVINCIA_.NOMBRE,'') AS DOMICILIO, ALIAS_0.DENOMINACION,ALIAS_5.NOMBRE AS ZONA, " & _
              "ALIAS_7.NUMERO AS TELEFONO, ALIAS_8.DIRECCIONELECTRONICA AS MAIL, " & _
              "V_TIPOPAGO_.NOMBRE AS TP, V_TIPOPAGO_.OBSERVACION as TipoPago, V_TIPOPAGO_.ACTIVESTATUS AS ACTIVESTATUSTP, V_PROVINCIA_.NOMBRE AS Provincia, " & _
              "V_CIUDAD_.NOMBRE AS Ciudad, ALIAS_0.CODIGO + ' ' + ALIAS_3.CUIT + ' ' + ALIAS_3.NOMBRE AS CONCATENADO, V_EZI_POCICION_IVA_CLIENTES.CODIGO AS IVA, V_TIPOPAGO_.ID AS IDPAGO, ALIAS_0.DOMICILIOFACTURACION_ID as domicilio_id " & _
              "FROM V_PERSONA_ AS ALIAS_3 RIGHT OUTER JOIN V_CLIENTE AS ALIAS_0 WITH (READPAST) LEFT OUTER JOIN " & _
              "V_EZI_POCICION_IVA_CLIENTES ON ALIAS_0.ID = V_EZI_POCICION_IVA_CLIENTES.idcliente LEFT OUTER JOIN " & _
              "V_TIPOPAGO_ ON ALIAS_0.TIPOPAGO_ID = V_TIPOPAGO_.ID ON ALIAS_3.ID = ALIAS_0.ENTEASOCIADO_ID LEFT OUTER JOIN " & _
              "V_CIUDAD_ RIGHT OUTER JOIN V_DOMICILIO_ AS ALIAS_6 ON V_CIUDAD_.ID = ALIAS_6.CIUDAD_ID LEFT OUTER JOIN " & _
              "V_PROVINCIA_ ON ALIAS_6.PROVINCIA_ID = V_PROVINCIA_.ID ON ALIAS_0.DOMICILIOFACTURACION_ID = ALIAS_6.ID LEFT OUTER JOIN " & _
              "V_TELEFONO_ AS ALIAS_7 ON ALIAS_3.TELEFONOPRINCIPAL_ID = ALIAS_7.ID LEFT OUTER JOIN " & _
              "V_DIRECCIONELECTRONICA_ AS ALIAS_8 ON ALIAS_3.DIRECELECTRONICAPRINCIPAL_ID = ALIAS_8.ID LEFT OUTER JOIN " & _
              "V_ZONA_ AS ALIAS_5 ON ALIAS_0.ZONA_ID = ALIAS_5.ID " & _
              "WHERE     ALIAS_0.ID = '" & Text19.Text & "'"
   End If
  Else
    xquery1 = "SELECT  ALIAS_0.ID AS ID,  ALIAS_0.CODIGO, ALIAS_3.NOMBRE AS RAZONSOCIAL, ALIAS_3.CUIT,ALIAS_6.CALLE, ISNULL(ALIAS_6.CALLE,'') + '-'+ ISNULL(V_CIUDAD_.NOMBRE,'') + '-'+ ISNULL(V_PROVINCIA_.NOMBRE,'') AS DOMICILIO, ALIAS_0.DENOMINACION,ALIAS_5.NOMBRE AS ZONA, " & _
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
              "WHERE     (ALIAS_0.ACTIVESTATUS <> 2) order by ALIAS_3.NOMBRE "
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
        Else
            DataCombo3.Text = "CONTADO"
        End If
        Text1(4).Text = DataGrid2.Columns(3).Text
        Text1(3).Text = DataGrid2.Columns(5).Text
'        If DataCombo3.Text = "CONTADO" Then
'            DataCombo3.Enabled = False
'        Else
'            DataCombo3.Enabled = True
'        End If
        Text1(5).SetFocus
        
        datiibb.RecordSource = "SELECT     pi.CODIGO, c.CODIGO AS codigocliente, c.ID AS idcliente, ipi.COEFICIENTE, pi.NOMBRE, ISNULL(v_ezi_excenciones_iibb_pos.EXENCION_IIBB_PERCEPC_P, 0) AS EXENCION " & _
                               "FROM         v_ezi_excenciones_iibb_pos RIGHT OUTER JOIN " & _
                               "V_UD_CLIENTE_ ON v_ezi_excenciones_iibb_pos.BO_PLACE_ID = V_UD_CLIENTE_.EXENCION_IIBB_PERCEPCION_ID RIGHT OUTER JOIN " & _
                               "V_POSICIONIMPUESTO_ AS pi INNER JOIN " & _
                               "V_ITEMPOSICIONADORIMPUESTOS_ AS ipi ON pi.ID = ipi.POSICIONIMPUESTO_ID INNER JOIN " & _
                               "V_POSICIONADORIMPUESTOS_ AS p ON ipi.BO_PLACE_ID = p.ITEMSPOSICIONADORIMPUESTOS_ID INNER JOIN " & _
                               "V_CLIENTE_ AS c ON p.ID = c.POSICIONADORIMPUESTOS_ID ON V_UD_CLIENTE_.ID = c.BOEXTENSION_ID " & _
                               "WHERE     (pi.BO_PLACE_ID = '{3BF2F57F-9AE4-4362-ACC2-38DE2275D517}') and c.id = '" & DataGrid2.Columns(0).Text & "'"
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
        
        xciudadcliente = DataGrid2.Columns("Ciudad").Text
        Call calcula_Click
        
    Else
        menu = 4
        query = xquery1
        lista_clientes.Show
    End If


End Sub

Private Sub blanco_Click()

    negro.Visible = False
    blanco.Visible = False
    tipofac = "NN"

End Sub

Private Sub buscar_Click()

    lista_presupuestos.Show

End Sub

Private Sub bvendedor_Click()
    
  If Text1(0).Text <> "" Then
    xbusqueda = "%" + Text1(0).Text + "%"
    xquery1 = "SELECT    V_VENDEDOR_.ID, V_VENDEDOR_.CODIGO, V_PERSONA_.NOMBRE, V_VENDEDOR_.CODIGO  + ' ' + V_PERSONA_.NOMBRE as cancatena " & _
              "FROM V_VENDEDOR_ INNER JOIN V_PERSONA_ ON V_VENDEDOR_.ENTEASOCIADO_ID = V_PERSONA_.ID " & _
              "WHERE (V_VENDEDOR_.ACTIVESTATUS = 0)  and V_VENDEDOR_.CODIGO  + ' ' + V_PERSONA_.NOMBRE like '" & xbusqueda & "' order by V_PERSONA_.NOMBRE"
  Else
    xquery1 = "SELECT    V_VENDEDOR_.ID, V_VENDEDOR_.CODIGO, V_PERSONA_.NOMBRE, V_VENDEDOR_.CODIGO  + ' ' + V_PERSONA_.NOMBRE as cancatena " & _
              "FROM V_VENDEDOR_ INNER JOIN V_PERSONA_ ON V_VENDEDOR_.ENTEASOCIADO_ID = V_PERSONA_.ID " & _
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
    Else
        menu = 4
        query = xquery1
        lista_vendedores.Show
    End If
    
    



End Sub

Private Sub calcula_Click()
On Error Resume Next

    xcol = grilla.Col
    grilla.Col = 3
    xcant = Val(grilla.Text)
    grilla.Col = 5
    ximporte = grilla.Text
    If ximporte = "" Then ximporte = 0
    grilla.Text = Format(ximporte, "###,##0.00")
    If grilla.Text = "0.00" And grilla.TextMatrix(grilla.Row, 0) = "" Then
        grilla.Text = ""
    End If
    grilla.Col = 11
    xiva = grilla.Text
    If xiva = "" Then xiva = 1.21
    ximportesiva = Round(ximporte * xcant / xiva, 2)
    grilla.Col = 6
    grilla.Text = Format(ximportesiva, "###,##0.00")
    If grilla.Text = "0.00" And grilla.TextMatrix(grilla.Row, 0) = "" Then
        grilla.Text = ""
    End If
    
 '   If xcol = 7 Or Val(Text11.Text) <> 0 Then
 '       grilla.Col = 7
 '       xbonifimporte = grilla.Text
 '       grilla.Text = Format(xbonifimporte, "###,##0.00")
 '       If grilla.Text = "0.00" And grilla.TextMatrix(grilla.Row, 0) = "" Then
 '           grilla.Text = ""
 '       End If
 '       grilla.Col = 8
 '       xbonifporcent = Round((xbonifimporte / (xcant * ximporte)) * 100, 2)
 '       grilla.Text = Format(xbonifporcent, "###,##0.00")
 '       If grilla.Text = "0.00" And grilla.TextMatrix(grilla.Row, 0) = "" Then
 '         grilla.Text = ""
 '       End If
 '   End If

    If xcol = 8 Or xcol = 3 Or xcol = 5 <> 0 Then
        grilla.Col = 8
        xbonifporcent = grilla.Text
        grilla.Text = Format(xbonifporcent, "###,##0.00")
        If grilla.Text = "0.00" And grilla.TextMatrix(grilla.Row, 0) = "" Then
            grilla.Text = ""
        End If
        grilla.Col = 7
        xbonifimporte = Round(xbonifporcent * xcant * ximporte / 100)
        grilla.Text = Format(xbonifimporte, "###,##0.00")
        If grilla.Text = "0.00" And grilla.TextMatrix(grilla.Row, 0) = "" Then
            grilla.Text = ""
        End If

    End If
    
    grilla.Col = 9
    If xbonifimporte = "" Then xbonifimporte = 0
    If ximporte = "" Then ximporte = 0
    
    xtotal = Round(xcant * Round(ximporte, 10) - Round(xbonifimporte, 10), 2)
    grilla.Text = Format(xtotal, "###,##0.00")
    If grilla.Text = "0.00" And grilla.TextMatrix(grilla.Row, 0) = "" Then
          grilla.Text = ""
    End If


    grilla.Col = 10
    grilla.Text = xcant
        If grilla.Text = "0" And grilla.TextMatrix(grilla.Row, 0) = "" Then
            grilla.Text = ""
        End If


    xtotalgral = 0
    xsubtotalgral = 0
    xiva10 = 0
    xiva21 = 0
    For X = 1 To xlineasmax
            If grilla.TextMatrix(X, 9) = "" Then
                xgrilla = 0
                xsubtotal = 0
            Else
                xgrilla = grilla.TextMatrix(X, 9)
                If grilla.TextMatrix(X, 11) = "" Then
                    xsubtotal = grilla.TextMatrix(X, 9) / 1.21
                Else
                    xsubtotal = grilla.TextMatrix(X, 9) / grilla.TextMatrix(X, 11)
                End If
            End If
            
            If grilla.TextMatrix(X, 11) = "1.105" Then
                xiva10 = xiva10 + (xgrilla - xsubtotal)
            Else
                xiva21 = xiva21 + (xgrilla - xsubtotal)
            End If
            
            xtotalgral = xtotalgral + xgrilla
            xsubtotalgral = xsubtotalgral + xsubtotal
            
    Next X
    
'--- Calculo de Tem
    xcalculotem = 0
    If UCase(xcalculatempyp) = "S" And xciudadtem = xciudadcliente And DataGrid2.Columns(16).Text <> "CF" And DataGrid2.Columns(16).Text <> "EX" Then
        xcalculotem = xsubtotalgral * xalicuptatempip / 100
    End If
'--- Fin Calculo Tem

'--- Calculo de IIBB
    xcalculoIIBB = 0
    If UCase(xcalculaiibb) = "S" And DataGrid2.Columns(16).Text <> "CF" And DataGrid2.Columns(16).Text <> "EX" Then
        xcalculoIIBB = (xsubtotalgral * xalicuotaiibb / 100) * ((100 - xexentoiibb) / 100)
  '      If xcalculoIIBB <= 50 Then xcalculoIIBB = 0  ' Limite inferior para calculo de iibb
    End If
'--- Fin Calculo IIBB
    xtotalgral = xtotalgral + xcalculotem + xcalculoIIBB
    xsubtotal2 = xsubtotalgral + xiva10 + xiva21
    
    Text5.Text = Format(Round(xsubtotalgral, 2), "$ ###,##0.00")
    Text6.Text = Format(Round(xiva10, 2), "$ ###,##0.00")
    Text7.Text = Format(Round(xiva21, 2), "$ ###,##0.00")
    Text8.Text = Format(Round(xcalculotem, 2), "$ ###,##0.00")
    Text9.Text = Format(Round(xcalculoIIBB, 2), "$ ###,##0.00")
    Text16.Text = Format(Round(xsubtotal2, 2), "$ ###,##0.00")
    
    Text4.Text = Format(Round(xtotalgral, 2), "$ ###,##0.00")
        

    grilla.Col = xcol

End Sub

Private Sub Cancelar_Click()

    Unload Me
    frmalquiler_devolucion.Show

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
'On Error Resume Next

    'datencabezado.RecordSource = "SELECT * From ud_ezi_puntodeventa_encabezado where  numeradorinterno = 'Nota de Venta' and  id ='" & Text18.Text & "' "
    datencabezado.RecordSource = query
    datencabezado.Refresh
    If datencabezado.Recordset.EOF = False Then
     xinicio = 1
     xnumeroremito = ""
     datencabezado.Recordset.MoveFirst
     Do While datencabezado.Recordset.EOF = False
        xidpre = datencabezado.Recordset.Fields("id")
        xnumeroremito = xnumeroremito + ", " + datencabezado.Recordset.Fields("numerodocumento")
        xidremitovta = datencabezado.Recordset.Fields("idremito")
        presupuestobase = xidpre
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
        Else
            DataCombo3.Text = "CONTADO"
        End If
        Text1(4).Text = DataGrid2.Columns(3).Text
        Text1(3).Text = DataGrid2.Columns(5).Text
        If DataCombo3.Text = "CONTADO" Then
            DataCombo3.Enabled = False
        Else
            DataCombo3.Enabled = True
        End If
        Text1(7).SetFocus
        
        datiibb.RecordSource = "SELECT     pi.CODIGO, c.CODIGO AS codigocliente, c.ID AS idcliente, ipi.COEFICIENTE, pi.NOMBRE, ISNULL(v_ezi_excenciones_iibb_pos.EXENCION_IIBB_PERCEPC_P, 0) AS EXENCION " & _
                               "FROM         v_ezi_excenciones_iibb_pos RIGHT OUTER JOIN " & _
                               "V_UD_CLIENTE_ ON v_ezi_excenciones_iibb_pos.BO_PLACE_ID = V_UD_CLIENTE_.EXENCION_IIBB_PERCEPCION_ID RIGHT OUTER JOIN " & _
                               "V_POSICIONIMPUESTO_ AS pi INNER JOIN " & _
                               "V_ITEMPOSICIONADORIMPUESTOS_ AS ipi ON pi.ID = ipi.POSICIONIMPUESTO_ID INNER JOIN " & _
                               "V_POSICIONADORIMPUESTOS_ AS p ON ipi.BO_PLACE_ID = p.ITEMSPOSICIONADORIMPUESTOS_ID INNER JOIN " & _
                               "V_CLIENTE_ AS c ON p.ID = c.POSICIONADORIMPUESTOS_ID ON V_UD_CLIENTE_.ID = c.BOEXTENSION_ID " & _
                               "WHERE     (pi.BO_PLACE_ID = '{3BF2F57F-9AE4-4362-ACC2-38DE2275D517}') and c.id = '" & DataGrid2.Columns(0).Text & "'"
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
        
        xciudadcliente = DataGrid2.Columns("Ciudad").Text
'*** fin Carga Clinte
'*** Carga Vendedor
        xqueryvende = "SELECT    V_VENDEDOR_.ID, V_VENDEDOR_.CODIGO, V_PERSONA_.NOMBRE, V_VENDEDOR_.CODIGO  + ' ' + V_PERSONA_.NOMBRE as cancatena " & _
              "FROM V_VENDEDOR_ INNER JOIN V_PERSONA_ ON V_VENDEDOR_.ENTEASOCIADO_ID = V_PERSONA_.ID " & _
              "WHERE (V_VENDEDOR_.ACTIVESTATUS = 0) and V_VENDEDOR_.ID = '" & datencabezado.Recordset.Fields("vendedorid") & "'"
    
        datvendedor.RecordSource = xqueryvende
        datvendedor.Refresh
        Text1(0).Text = DataGrid1.Columns(2).Text
        Text1(7).SetFocus
'*** Fin carga Vendedor
        Text1(5).Text = datencabezado.Recordset.Fields("detalle")
        Text15.Text = datencabezado.Recordset.Fields("nota")
        Text1(6).Text = datencabezado.Recordset.Fields("horaretiro")
        Text1(9).Text = DateValue(datencabezado.Recordset.Fields("fechadelcomprobante"))
        

        
        
        
        datitems.RecordSource = "SELECT  NV.id, NV.claveprimaria, NV.codigoproducto, NV.cantidadproducto, NV.unidaddemedidaid, NV.preciou, NV.listaid, NV.bonificacionitem, NV.importedebonificacion, " & _
                                "NV.subtotal, NV.tipodeentregaitemid, NV.nombre_producto, NV.iva, NV.observacion, NV.entregar, NV.idproducto, NV.idclaveprimariaremito, NV.iditemremito, " & _
                                "NV.cantidadproducto AS Cant_Orig, R.cantidadremitida AS Cant_Remitida, ISNULL(IFAC.cantidadproducto, 0) AS Cant_Facturada, R.unidaddemedida AS Um " & _
                                "FROM ud_ezi_puntodeventa_detalle_notav AS NV WITH (readpast) LEFT OUTER JOIN v_ezi_pos_remito AS R ON NV.iditemremito = R.iditem LEFT OUTER JOIN " & _
                                "ud_ezi_puntodeventa_detalle_factm AS IFAC WITH (readpast) LEFT OUTER JOIN ud_ezi_puntodeventa_encabezado AS E WITH (readpast) ON IFAC.claveprimaria = E.id ON R.iditem = IFAC.iditemremito " & _
                                "Where nv.claveprimaria = " & xidpre & " " & _
                                "ORDER BY NV.id"
        datitems.Refresh
    
        Debug.Print datitems.RecordSource
      If datitems.Recordset.EOF = False Then
        datitems.Recordset.MoveFirst
        For X = xinicio To datitems.Recordset.RecordCount + xinicio - 1
            xiditemremitovta = datitems.Recordset.Fields("iditemremito")
            grilla.TextMatrix(X, 0) = datitems.Recordset.Fields("idproducto")
            grilla.TextMatrix(X, 1) = datitems.Recordset.Fields("codigoproducto")
            grilla.TextMatrix(X, 2) = datitems.Recordset.Fields("nombre_producto")
            grilla.TextMatrix(X, 3) = datitems.Recordset.Fields("cantidadproducto")
            grilla.TextMatrix(X, 4) = datitems.Recordset.Fields("unidaddemedidaid")
            grilla.TextMatrix(X, 5) = datitems.Recordset.Fields("preciou")
            grilla.TextMatrix(X, 11) = (datitems.Recordset.Fields("iva") / 100) + 1
            grilla.TextMatrix(X, 6) = datitems.Recordset.Fields("preciou") * Round(grilla.TextMatrix(X, 11), 2) * datitems.Recordset.Fields("cantidadproducto")
            grilla.TextMatrix(X, 8) = Round(datitems.Recordset.Fields("bonificacionitem"), 10)
            grilla.TextMatrix(X, 9) = Round(datitems.Recordset.Fields("subtotal"), 10)
            grilla.TextMatrix(X, 10) = datitems.Recordset.Fields("cantidadproducto")
            grilla.TextMatrix(X, 13) = datitems.Recordset.Fields("preciou")
            grilla.TextMatrix(X, 14) = xidremitovta ' id remito
            grilla.TextMatrix(X, 15) = xiditemremitovta ' id item remito
            grilla.Col = 3
            Call calcula_Click
            datitems.Recordset.MoveNext
        Next X
      
        xremito = xidremitovta
        xinicio = datitems.Recordset.RecordCount + 1
      End If
        
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
       datencabezado.Recordset.MoveNext
     Loop
       
    Text17.Text = Mid(xnumeroremito, 3, Len(xnumeroremito) - 1)
    
    Else
        mensa = MsgBox("No Existe el Nro de presupuesto seleccionado", vbInformation, "!! Sin Coincidencias !!")
    End If
    

    
    

End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
        KeyAscii = 0
        
        If Combo1.Text <> xumvta And grilla.Text <> Combo1.Text Then
            grilla.Text = Combo1.Text
            grilla.Col = 5
            grilla.Text = Round(grilla.Text / xfaconv(Combo1.ListIndex), 2)
            grilla.TextMatrix(grilla.Row, 12) = xfaconv(Combo1.ListIndex)
            Call calcula_Click
            Call ubicatextogrilla_Click
            Exit Sub
        End If
        If Combo1.Text = xumvta And grilla.Text <> Combo1.Text Then
            grilla.Text = Combo1.Text
            grilla.Col = 5
            grilla.Text = Round(grilla.Text * grilla.TextMatrix(grilla.Row, 12), 2)
            Call calcula_Click
            Call ubicatextogrilla_Click
            Exit Sub
        End If
        grilla.Col = 5
        Call ubicatextogrilla_Click

End If


End Sub

Private Sub Combo1_KeyUp(KeyCode As Integer, Shift As Integer)

If KeyCode = 39 Then
    grilla.Col = 5
    Call ubicatextogrilla_Click
End If

If KeyCode = 37 Then
    grilla.Col = 3
    Call ubicatextogrilla_Click
End If

If KeyCode = 121 Then
       KeyCode = 0
       Call grabar_Click
End If


If KeyCode = 116 Then
       Call agregaproducto_Click
End If


End Sub

Private Sub Command7_Click()

    If UCase(datparametros.Recordset.Fields("cobroautomatico")) = "S" Then
        frmcobranzacdo.Show
    End If

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

Private Sub DataCombo3_KeyPress(KeyAscii As Integer)

On Error Resume Next

    If KeyAscii = 13 Then
        KeyAscii = 0
        Text1(0).SetFocus
    End If

End Sub

Private Sub Form_Load()
Aplicar_skin Me

yventana = Inicio.Height / 2 - 1000
xventana = Inicio.Width / 2

frmalquiler_devolucion.Top = yventana - frmalquiler_devolucion.Height / 2
frmalquiler_devolucion.Left = xventana - frmalquiler_devolucion.Width / 2


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
datitempresup.ConnectionString = login.conexiontotal
datcolaimportar.ConnectionString = login.conexiontotal
datitemsnv.ConnectionString = login.conexiontotal

negro.Visible = False
blanco.Visible = False
tipofac = "CF"
Text18.Text = ""
Text17.Text = ""
Text19.Text = ""

Combo2.AddItem ("NO")
Combo2.AddItem ("SI")
Combo2.ListIndex = 0

Text3.Text = Date
Text1(7).Text = Left(Time, 5)

    dattipopago.RecordSource = "SELECT ID, NOMBRE AS CODIGO, OBSERVACION AS TipoPago From V_TIPOPAGO_ WHERE (ACTIVESTATUS = 0) order by NOMBRE"
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
grilla.Col = 1
grilla.Text = "Codigo"
grilla.ColWidth(1) = 1000
grilla.Col = 2
grilla.Text = "Descipcion"
grilla.ColWidth(2) = 6800
grilla.Col = 3
grilla.Text = "Cant."
grilla.ColWidth(3) = 800
grilla.Col = 4
grilla.Text = "U.M."
grilla.ColWidth(4) = 1000
grilla.Col = 5
grilla.Text = "$ Unit.C/Iva."
grilla.ColWidth(5) = 0
grilla.Col = 6
grilla.Text = "$Total S/Iva."
grilla.ColWidth(6) = 0
grilla.Col = 7
grilla.Text = "$ Bonif."
'grilla.ColWidth(7) = 1000
grilla.ColWidth(7) = 0
grilla.Col = 8
grilla.Text = "% Bonif."
grilla.ColWidth(8) = 0
grilla.Col = 9
grilla.Text = "$ Total"
grilla.ColWidth(9) = 0
grilla.Col = 10
grilla.Text = "Cant.a Devolver"
grilla.ColWidth(10) = 1200
grilla.Col = 11
grilla.Text = "Iva"
grilla.ColWidth(11) = 0
grilla.Col = 12
grilla.Text = ".."
grilla.ColWidth(12) = 0
grilla.Col = 14
grilla.Text = "Remito"
grilla.ColWidth(14) = 0
grilla.Col = 15
grilla.Text = "IdItemRemito"
grilla.ColWidth(15) = 0

grilla.ColWidth(13) = 0

xlineasmax = datparametros.Recordset.Fields("limiteitemsnotaventa")
grilla.Rows = xlineasmax + 1

For X = 2 To xlineasmax Step 2
  For Y = 1 To 11
    grilla.Col = Y
    grilla.Row = X
    grilla.CellBackColor = RGB(231, 235, 218)
  Next Y
Next X




   
End Sub

Private Sub grabar_Click()


    datencabezado.RecordSource = "SELECT MAX(CONVERT(decimal, isnull(claveprimaria,0))) AS claveprimaria FROM ud_ezi_puntodeventa_encabezado with(readpast) "
    datencabezado.Refresh
    
    If datencabezado.Recordset.EOF = True Then
        xclaveprimaria = 1
    Else
        xclaveprimaria = datencabezado.Recordset.Fields("claveprimaria") + 1
    End If
    
    If IsNull(xclaveprimaria) = True Then xclaveprimaria = 1

    datencabezado.RecordSource = "SELECT * From ud_ezi_puntodeventa_encabezado where id =0 "
    datencabezado.Refresh
    datitems.RecordSource = "Select * from ud_ezi_puntodeventa_detalle_notav where id = 0"
    datitems.Refresh
    
    datencabezado.Recordset.AddNew
    datencabezado.Recordset.Fields("numeradorinterno") = "Remito de Devolucion"
    datencabezado.Recordset.Fields("fechadelcomprobante") = DateValue(Text3.Text) + TimeValue(Str(Time))
    datencabezado.Recordset.Fields("sucursal") = datparametros.Recordset.Fields("sucursal")
    datencabezado.Recordset.Fields("clienteid") = DataGrid2.Columns(0).Text
    datencabezado.Recordset.Fields("cliente") = DataGrid2.Columns(2).Text
    datencabezado.Recordset.Fields("vendedorid") = DataGrid1.Columns(0).Text
    datencabezado.Recordset.Fields("vendedor") = DataGrid1.Columns(2).Text
    datencabezado.Recordset.Fields("detalle") = Text1(5).Text
    datencabezado.Recordset.Fields("nota") = Text15.Text
    datencabezado.Recordset.Fields("cotizacion") = 1
    datencabezado.Recordset.Fields("listadeprecioid") = DataCombo1.BoundText
    datencabezado.Recordset.Fields("tipodepagoid") = DataCombo3.BoundText
    datencabezado.Recordset.Fields("tipodefactura") = Text1(2).Text
    
    If tipofac <> "NN" Then
        datencabezado.Recordset.Fields("tipodefacturacionid") = datparametros.Recordset.Fields("facdefecto")
    Else
        datencabezado.Recordset.Fields("tipodefacturacionid") = tipofac
    End If
    
    If Left(login.nombrebd, 14) = "MMOSSE" And tipofac <> "NN" Then
        If DataCombo3.Text = "CONTADO" Then
            datencabezado.Recordset.Fields("tipodefacturacionid") = datparametros.Recordset.Fields("facdefecto")
        Else
            datencabezado.Recordset.Fields("tipodefacturacionid") = datparametros.Recordset.Fields("facdefectocc")
        End If
    End If
    datencabezado.Recordset.Fields("claveprimaria") = xclaveprimaria
    datencabezado.Recordset.Fields("fechadeentrega") = DateValue(Text3.Text) + TimeValue(Str(Time))
    datencabezado.Recordset.Fields("tiporecargo") = "$"
    datencabezado.Recordset.Fields("tipobonificacion") = "$"
    datencabezado.Recordset.Fields("importeglobal") = Round(Text4.Text, 2)
    'datencabezado.Recordset.Fields("numerodefactura") = xclaveprimaria
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
    datencabezado.Recordset.Fields("transferido") = "False"
    
    datencabezado.Recordset.Fields("percepiibb") = Round(Text9.Text, 2)
    datencabezado.Recordset.Fields("perceptem") = Round(Text8.Text, 2)
    datencabezado.Recordset.Fields("totaltr") = Round(Text4.Text, 2)
    
    datencabezado.Recordset.Fields("alquiler") = "S"
    datencabezado.Recordset.Fields("retira") = Text1(5).Text
    datencabezado.Recordset.Fields("horaretiro") = Text1(6).Text
    datencabezado.Recordset.Fields("horadevol") = Text1(7).Text
    datencabezado.Recordset.Fields("recibedevalquiler") = Text1(8).Text
    datencabezado.Recordset.Fields("requierereparacion") = Left(Combo2.Text, 1)
    datencabezado.Recordset.Fields("presupuestobase") = xidpre
    datencabezado.Recordset.Fields("trazabilidad_id") = xidpre
    datencabezado.Recordset.Fields("comprobanteorigen") = xidremitovta
    
    
    
    datencabezado.Recordset.UpdateBatch adAffectCurrent
    presupuestobase = datencabezado.Recordset.Fields("id")
    
   '** Establene numero de Remitos Manuales, y no Fiscales
     xnumerador = "Recepcion de Bien Alquilado " + datparametros.Recordset.Fields("sucursal")
     datcolaimportar.RecordSource = "SELECT p.ID AS ID, p.NOMBRE AS puntero, n.NUMERO AS NUMERO, p.CARACTERESPREFIJO AS puntoventa, p.NUMERO_ID " & _
                           "FROM V_NUMERADOR_ AS p LEFT OUTER JOIN V_NUMERO_ AS n ON p.NUMERO_ID = n.ID " & _
                           "WHERE     (p.ACTIVESTATUS <> 2) AND (p.NOMBRE = '" & xnumerador & "') "
    datcolaimportar.Refresh
        datencabezado.Recordset.Fields("numerodefactura") = datcolaimportar.Recordset.Fields("numero")
        xnumero = datcolaimportar.Recordset.Fields("numero")
        xidnumero = datcolaimportar.Recordset.Fields("numero_id")
        datencabezado.Recordset.Fields("puntodeventa") = "0"
        datcolaimportar.RecordSource = "Select * from numero with(readpast) where id = '" & xidnumero & "'"
        datcolaimportar.Refresh
        datcolaimportar.Recordset.Fields("numero") = xnumero + 1
        datcolaimportar.Recordset.UpdateBatch adAffectCurrent
    '** Fin de asignacion de numero a Remtio
     
    datencabezado.Recordset.UpdateBatch adAffectCurrent
    
'--- Graba Items
    
    For X = 1 To xlineasmax
        If grilla.TextMatrix(1, 0) = "" Then
          mensa = MsgBox("No puede Grabar esta Venta sin Items", vbCritical, "Error")
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
        datitems.Recordset.Fields("preciou") = Round(grilla.TextMatrix(X, 5), 4)
        datitems.Recordset.Fields("bonificacionitem") = grilla.TextMatrix(X, 8)
        datitems.Recordset.Fields("subtotal") = Round(grilla.TextMatrix(X, 9), 4)
        datitems.Recordset.Fields("entregar") = Round(grilla.TextMatrix(X, 10), 4)
        datitems.Recordset.Fields("iva") = (Round(grilla.TextMatrix(X, 11), 4) - 1) * 100
        
        datitems.Recordset.UpdateBatch adAffectCurrent
    Next

'******* Graba Cola importar
        
        datcolaimportar.RecordSource = "Select * from ud_ezi_cola_importar where id = 0"
        datcolaimportar.Refresh
        
        datcolaimportar.Recordset.AddNew
        datcolaimportar.Recordset.Fields("id_encabezado") = presupuestobase
        datcolaimportar.Recordset.Fields("tipodedocumentoid") = datparametros.Recordset.Fields("idremitodev")
        datcolaimportar.Recordset.Fields("unidadoperativaid") = datparametros.Recordset.Fields("target")
        datcolaimportar.Recordset.Fields("fecha_hora") = DateValue(Text3.Text) + TimeValue(Str(Time))
        
        datcolaimportar.Recordset.UpdateBatch adAffectCurrent

    Call Cancelar_Click
    Unload Me

'            Call grfacturactacte_Click
        

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


Private Sub grilla_Click()

       If grilla.Col = 4 Then
          Call UM_Click
          Exit Sub
       End If

Text2.Visible = True

Call ubicatextogrilla_Click


End Sub

Private Sub grilla_GotFocus()

    Call calcula_Click

End Sub

Private Sub grilla_KeyDown(KeyCode As Integer, Shift As Integer)

       If grilla.Col = 4 Then
          Call UM_Click
          Exit Sub
       End If

Call ubicatextogrilla_Click

End Sub



Private Sub grilla_KeyUp(KeyCode As Integer, Shift As Integer)

       If grilla.Col = 4 Then
          Call UM_Click
          Exit Sub
       End If

Call ubicatextogrilla_Click

End Sub

Private Sub grilla_Scroll()

Text2.Visible = False

End Sub

Private Sub grremitoctacte_Click()

End Sub


Private Sub grremito_Click()

End Sub

Private Sub KewlButtons1_Click()
On Error Resume Next
        For X = 0 To 9
            grilla.Col = X
            grilla.Text = ""
        Next X
        
        For X = grilla.Row + 1 To 48
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
    Call ubicatextogrilla_Click

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

    If Index = 6 Then
        Text1(6).SelStart = 0
        Text1(6).SelLength = Len(Text1(6).Text)
    End If
    If Index = 7 Then
        Text1(7).SelStart = 0
        Text1(7).SelLength = Len(Text1(6).Text)
    End If
    
        

End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
On Error Resume Next


    If KeyAscii = 13 Then
        Text1(Index).Text = UCase(Text1(Index).Text)
        KeyAscii = 0
        If Index = 0 Then
            Call bvendedor_Click
            If Text1(0).Text <> "" Then
                menu = 4
                lista_clavevendedor.Show
            End If
        End If
        If Index = 1 Then
            Call bclientes_Click
        End If
        If Index = 5 Or Index = 6 Or Index = 7 Then
            Text1(Index + 1).SetFocus
        End If

        If Index = 6 Or Index = 7 Then
            Text1(Index).Text = Left(Replace(Text1(Index).Text, ":", "") + "0000", 4)
            Text1(Index).Text = Left(Text1(Index).Text, 2) + ":" + Right(Text1(Index).Text, 2)
        End If
        
        If Index = 8 Then
            agregaproducto.SetFocus
            Call agregaproducto_Click
        End If
        
    End If
    
End Sub

Private Sub Text1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)

If KeyCode = 114 Then
       Call bvendedor_Click
End If

If KeyCode = 116 Then
       Call agregaproducto_Click
End If

If KeyCode = 38 Then
    If Index = 5 Then Text1(1).SetFocus
    If Index = 1 Then Text1(0).SetFocus
End If

If KeyCode = 121 Then
    KeyCode = 0
       Call grabar_Click
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
        For X = 1 To xlineasmax
              If grilla.TextMatrix(X, 1) = "" Then Exit For
              grilla.Row = X
              grilla.TextMatrix(X, 8) = Text10.Text
              If Val(Text10.Text) = 0 Then
                 grilla.TextMatrix(X, 7) = 0
              End If
              Call calcula_Click
              ximporteboniftotal = ximporteboniftotal + grilla.TextMatrix(X, 7)
        Next X
        Text11.Text = Format(ximporteboniftotal, "##0.00")
        Text15.SetFocus
End If


End Sub

Private Sub Text10_KeyUp(KeyCode As Integer, Shift As Integer)

If KeyCode = 121 Then
       KeyCode = 0
       Call grabar_Click
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
        For X = 1 To xlineasmax
              If grilla.TextMatrix(X, 1) = "" Then Exit For
              grilla.Row = X
              grilla.TextMatrix(X, 7) = Round((grilla.TextMatrix(X, 9) / xvalorsubtotal) * Text11.Text, 2)
              If Val(Text11.Text) = 0 Then
                 grilla.TextMatrix(X, 8) = 0
              End If
              Call calcula_Click
       
        Next X
        Text10.Text = Format(grilla.TextMatrix(1, 8), "##0.00")
        Text15.SetFocus
End If

End Sub

Private Sub Text11_KeyUp(KeyCode As Integer, Shift As Integer)

If KeyCode = 121 Then
       KeyCode = 0
       Call grabar_Click
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
        
        grilla.TextMatrix(X, 5) = grilla.TextMatrix(X, 13)
        Text13.Text = 0
        Call calcula_Click
        
        Text13.Text = Format(0, "##0.00")
        Text14.Text = Format(0, "##0.00")
        ximporteboniftotal = Round(Text16.Text, 20)
        For X = 1 To xlineasmax
              If grilla.TextMatrix(X, 1) = "" Then Exit For
              grilla.Row = X
              
              grilla.TextMatrix(X, 5) = Round(grilla.TextMatrix(X, 13), 20) * ((Round(Text12.Text, 20) / 100) + 1)
              If Val(Text12.Text) = 0 Then
                 grilla.TextMatrix(X, 5) = grilla.TextMatrix(X, 13)
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

If KeyCode = 121 Then
       KeyCode = 0
       Call grabar_Click
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

If KeyCode = 121 Then
       KeyCode = 0
       Call grabar_Click
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

Private Sub Text15_KeyUp(KeyCode As Integer, Shift As Integer)

If KeyCode = 121 Then
       KeyCode = 0
       Call grabar_Click
End If



End Sub

Private Sub Text17_KeyPress(KeyAscii As Integer)


If KeyAscii = 13 Then
        KeyAscii = 0
        Call cargapresupuesto_Click
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


If KeyAscii = 13 Then
        KeyAscii = 0
        grilla.Text = Text2.Text
        If grilla.Col = 5 Then
            grilla.TextMatrix(grilla.Row, 13) = Text2.Text
        End If
        If grilla.Col = 3 Or grilla.Col = 5 Or grilla.Col = 7 Or grilla.Col = 8 Then Call calcula_Click
        grilla.Col = grilla.Col + 1
        If grilla.Col = 4 Then
           Call UM_Click
           Exit Sub
        End If
        If grilla.Col = 6 Then grilla.Col = 8
        If grilla.Col = 9 Then grilla.Col = 10
        Call ubicatextogrilla_Click
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

If KeyCode = 39 And grilla.Col < 10 Then   'And Text2.SelStart = Len(Text2.Text)
  If grilla.Col <> 2 Then
    grilla.Col = grilla.Col + 1
       If grilla.Col = 4 Then
          Call UM_Click
          Exit Sub
       End If
    If grilla.Col = 6 Then grilla.Col = 8
    If grilla.Col = 9 Then grilla.Col = 10
    Call ubicatextogrilla_Click
  Else
    If Text2.SelStart = Len(Text2.Text) Then
        grilla.Col = grilla.Col + 1
        Call ubicatextogrilla_Click
    End If
  End If
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
    Call ubicatextogrilla_Click

End If

If KeyCode = 40 And grilla.Row < xlineasmax Then
    grilla.Row = grilla.Row + 1
    If grilla.RowIsVisible(grilla.Row) = False Then
        grilla.TopRow = grilla.TopRow + 1
    End If
    Call ubicatextogrilla_Click
End If



If KeyCode = 37 And grilla.Col > 2 Then ' And Text2.SelStart = 0
    grilla.Col = grilla.Col - 1
       If grilla.Col = 4 Then
          Call UM_Click
          Exit Sub
       End If
    If grilla.Col = 7 Then grilla.Col = 5
    If grilla.Col = 9 Then grilla.Col = 8
    Call ubicatextogrilla_Click
End If

If KeyCode = 116 Then
       Call agregaproducto_Click
End If

    
If KeyCode = 121 Then
       KeyCode = 0
       Call grabar_Click
End If

End Sub

Private Sub ubicatextogrilla_Click()
On Error Resume Next

Combo1.Visible = False
If grilla.Col = 6 Then
     Exit Sub
End If

Text2.Visible = True

If grilla.Col = 3 Or grilla.Col = 5 Or grilla.Col = 6 Or grilla.Col = 7 Or grilla.Col = 8 Then
    Text2.Alignment = 1
Else
    Text2.Alignment = 0
End If

If grilla.Col = 9 Then
    Text2.Enabled = False
Else
    Text2.Enabled = True
End If
    
Text2.Width = grilla.ColWidth(grilla.Col)
Text2.Text = grilla.Text
If grilla <> 2 Then xaltura = Text2.Height

If grilla.Col = 2 Then
    Text2.Height = Text2.Height * 3
Else
    Text2.Height = 315
End If

xfila = grilla.RowPos(grilla.Row) + grilla.Top + Picture1.Top + 50
xcolu = grilla.ColPos(grilla.Col) + grilla.Left + Picture1.Left + 50
Text2.Left = xcolu
Text2.Top = xfila

If grilla.Col <> 2 Then
    Text2.SelStart = 0
    Text2.SelLength = Len(Text2.Text)
End If

Text2.SetFocus

End Sub

Private Sub UM_Click()
On Error Resume Next
    
    xid2 = grilla.TextMatrix(grilla.Row, 0)
    If xid2 = "" Then Exit Sub
    
    Combo1.Visible = True
    Combo1.Clear
    
    
    
    
    datum.RecordSource = "SELECT     P.ID, P.CODIGO, P.UNIDADMEDIDA_ID, P.UNIDADMEDIDANOLINEAL_ID, V_UNIDADMEDIDA_.NOMBRE AS UMVTA, V_UNIDADMEDIDA__1.NOMBRE AS UMSTK, " & _
                         "ISNULL(V_ITEMCONVUNIDADESDEMEDIDA_.RELACIONCONVERSION, 1) AS RELACIONCONVERSION, V_UNIDADMEDIDA__2.NOMBRE AS UMDESTINO " & _
                         "FROM         V_UNIDADMEDIDA_ AS V_UNIDADMEDIDA__2 RIGHT OUTER JOIN " & _
                         "V_ITEMCONVUNIDADESDEMEDIDA_ ON V_UNIDADMEDIDA__2.ID = V_ITEMCONVUNIDADESDEMEDIDA_.UNIDADDESTINO_ID RIGHT OUTER JOIN " & _
                         "V_SERVICIO_ AS P ON V_ITEMCONVUNIDADESDEMEDIDA_.UNIDADORIGEN_ID = P.UNIDADMEDIDANOLINEAL_ID AND " & _
                         "V_ITEMCONVUNIDADESDEMEDIDA_.BO_PLACE_ID = P.EQUIVALENCIASPARTICULARES_ID LEFT OUTER JOIN " & _
                         "V_UNIDADMEDIDA_ AS V_UNIDADMEDIDA__1 ON P.UNIDADMEDIDA_ID = V_UNIDADMEDIDA__1.ID LEFT OUTER JOIN " & _
                         "V_UNIDADMEDIDA_ ON P.UNIDADMEDIDANOLINEAL_ID = V_UNIDADMEDIDA_.ID " & _
                         "WHERE  P.CODIGO = 'ALQ' "
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
