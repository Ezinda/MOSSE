VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{F6125AB1-8AB1-11CE-A77F-08002B2F4E98}#2.0#0"; "MSRDC20.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{D18BBD1F-82BB-4385-BED3-E9D31A3E361E}#1.0#0"; "KewlButtonz.ocx"
Begin VB.MDIForm Inicio 
   AutoShowChildren=   0   'False
   BackColor       =   &H00800000&
   ClientHeight    =   6315
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14055
   Icon            =   "Inicio.frx":0000
   LinkTopic       =   "MDIForm1"
   ScrollBars      =   0   'False
   WindowState     =   2  'Maximized
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   8160
      OleObjectBlob   =   "Inicio.frx":0442
      Top             =   2520
   End
   Begin VB.PictureBox picBackdrop 
      Align           =   1  'Align Top
      AutoRedraw      =   -1  'True
      Height          =   2175
      Left            =   0
      ScaleHeight     =   2115
      ScaleWidth      =   13995
      TabIndex        =   2
      Top             =   1320
      Visible         =   0   'False
      Width           =   14055
      Begin VB.PictureBox picOriginal 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   12000
         Left            =   0
         Picture         =   "Inicio.frx":0676
         ScaleHeight     =   12000
         ScaleWidth      =   17010
         TabIndex        =   4
         Top             =   0
         Width           =   17010
      End
      Begin VB.PictureBox picStretched 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   7260
         Left            =   1920
         ScaleHeight     =   7260
         ScaleWidth      =   4095
         TabIndex        =   3
         Top             =   600
         Width           =   4095
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      Negotiate       =   -1  'True
      TabIndex        =   0
      Top             =   660
      Width           =   14055
      _ExtentX        =   24791
      _ExtentY        =   1164
      ButtonWidth     =   609
      ButtonHeight    =   1005
      Appearance      =   1
      _Version        =   393216
      Begin KewlButtonz.KewlButtons KewlButtons13 
         Height          =   615
         Left            =   12360
         TabIndex        =   21
         ToolTipText     =   "F3-Crear Presupuesto de Venta"
         Top             =   0
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   1085
         BTYPE           =   14
         TX              =   "Proveedores"
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
         BCOL            =   15790320
         BCOLO           =   15790320
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "Inicio.frx":45ED4
         PICN            =   "Inicio.frx":45EF0
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin KewlButtonz.KewlButtons KewlButtons12 
         Height          =   615
         Left            =   2040
         TabIndex        =   20
         ToolTipText     =   "F3-Crear Presupuesto de Venta"
         Top             =   0
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   1085
         BTYPE           =   14
         TX              =   "Consulta NV"
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
         BCOL            =   15790320
         BCOLO           =   15790320
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "Inicio.frx":57518
         PICN            =   "Inicio.frx":57534
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Crystal.CrystalReport CrystalReporte 
         Left            =   18480
         Top             =   120
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
         WindowShowSearchBtn=   -1  'True
         WindowShowPrintSetupBtn=   -1  'True
         WindowShowRefreshBtn=   -1  'True
      End
      Begin KewlButtonz.KewlButtons KewlButtons3 
         Height          =   615
         Left            =   16440
         TabIndex        =   11
         ToolTipText     =   "F7-Panel Caja"
         Top             =   0
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   1085
         BTYPE           =   14
         TX              =   "Fact en Base a NV"
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
         BCOL            =   15790320
         BCOLO           =   15790320
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "Inicio.frx":57986
         PICN            =   "Inicio.frx":579A2
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin KewlButtonz.KewlButtons KewlButtons2 
         Height          =   615
         Left            =   6240
         TabIndex        =   10
         ToolTipText     =   "F3-Crear Presupuesto de Venta"
         Top             =   0
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   1085
         BTYPE           =   14
         TX              =   "Consulta Cotizaciones"
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
         BCOL            =   15790320
         BCOLO           =   15790320
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "Inicio.frx":57DF4
         PICN            =   "Inicio.frx":57E10
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
         Height          =   615
         Left            =   10320
         TabIndex        =   9
         ToolTipText     =   "F3-Crear Presupuesto de Venta"
         Top             =   0
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   1085
         BTYPE           =   14
         TX              =   "Clientes"
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
         BCOL            =   15790320
         BCOLO           =   15790320
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "Inicio.frx":5812A
         PICN            =   "Inicio.frx":58146
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin KewlButtonz.KewlButtons KewlButtons6 
         Height          =   615
         Left            =   8280
         TabIndex        =   8
         ToolTipText     =   "F3-Crear Presupuesto de Venta"
         Top             =   0
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   1085
         BTYPE           =   14
         TX              =   "Consulta de Precios"
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
         BCOL            =   15790320
         BCOLO           =   15790320
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "Inicio.frx":58598
         PICN            =   "Inicio.frx":585B4
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin KewlButtonz.KewlButtons KewlButtons4 
         Height          =   615
         Left            =   14400
         TabIndex        =   7
         ToolTipText     =   "F7-Panel Caja"
         Top             =   0
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   1085
         BTYPE           =   14
         TX              =   "Fact en Base Remito"
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
         BCOL            =   15790320
         BCOLO           =   15790320
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "Inicio.frx":58B4E
         PICN            =   "Inicio.frx":58B6A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin KewlButtonz.KewlButtons presupuesto 
         Height          =   615
         Left            =   4200
         TabIndex        =   6
         ToolTipText     =   "F3-Crear Presupuesto de Venta"
         Top             =   0
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   1085
         BTYPE           =   14
         TX              =   "Cotización (F4)"
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
         BCOL            =   15790320
         BCOLO           =   15790320
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "Inicio.frx":59104
         PICN            =   "Inicio.frx":59120
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin KewlButtonz.KewlButtons Notaventa 
         Height          =   615
         Left            =   0
         TabIndex        =   5
         ToolTipText     =   "F2-Crea Nota de Venta"
         Top             =   0
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   1085
         BTYPE           =   14
         TX              =   "Nota de Venta (F2)"
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
         BCOL            =   15790320
         BCOLO           =   15790320
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "Inicio.frx":596BA
         PICN            =   "Inicio.frx":596D6
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Timer Timer1 
         Interval        =   1
         Left            =   16320
         Top             =   0
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Por Nombre"
         Height          =   315
         Left            =   1800
         Picture         =   "Inicio.frx":59C70
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   1200
         UseMaskColor    =   -1  'True
         Width           =   1215
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
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6720
      Top             =   840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.Toolbar Toolbar2 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      Negotiate       =   -1  'True
      TabIndex        =   12
      Top             =   0
      Width           =   14055
      _ExtentX        =   24791
      _ExtentY        =   1164
      ButtonWidth     =   609
      ButtonHeight    =   1005
      Appearance      =   1
      _Version        =   393216
      Begin KewlButtonz.KewlButtons KewlButtons15 
         Height          =   615
         Left            =   12360
         TabIndex        =   23
         ToolTipText     =   "F3-Crear Presupuesto de Venta"
         Top             =   0
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   1085
         BTYPE           =   14
         TX              =   "Consulta de Precios"
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
         BCOL            =   15790320
         BCOLO           =   15790320
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "Inicio.frx":59F7A
         PICN            =   "Inicio.frx":59F96
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin KewlButtonz.KewlButtons KewlButtons14 
         Height          =   615
         Left            =   14400
         TabIndex        =   22
         ToolTipText     =   "F3-Crear Presupuesto de Venta"
         Top             =   0
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   1085
         BTYPE           =   14
         TX              =   "Clientes"
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
         BCOL            =   15790320
         BCOLO           =   15790320
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "Inicio.frx":5A530
         PICN            =   "Inicio.frx":5A54C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin KewlButtonz.KewlButtons KewlButtons11 
         Height          =   615
         Left            =   10320
         TabIndex        =   19
         ToolTipText     =   "F7-Panel Caja"
         Top             =   0
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   1085
         BTYPE           =   14
         TX              =   "Fact en Base a NV"
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
         BCOL            =   15790320
         BCOLO           =   15790320
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "Inicio.frx":5A99E
         PICN            =   "Inicio.frx":5A9BA
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin KewlButtonz.KewlButtons KewlButtons10 
         Height          =   615
         Left            =   8280
         TabIndex        =   18
         ToolTipText     =   "F7-Panel Caja"
         Top             =   0
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   1085
         BTYPE           =   14
         TX              =   "Fact en Base Remito"
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
         BCOL            =   15790320
         BCOLO           =   15790320
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "Inicio.frx":5AE0C
         PICN            =   "Inicio.frx":5AE28
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin KewlButtonz.KewlButtons KewlButtons9 
         Height          =   615
         Left            =   6240
         TabIndex        =   17
         ToolTipText     =   "F3-Crear Presupuesto de Venta"
         Top             =   0
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   1085
         BTYPE           =   14
         TX              =   "Consulta NV"
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
         BCOL            =   15790320
         BCOLO           =   15790320
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "Inicio.frx":5B3C2
         PICN            =   "Inicio.frx":5B3DE
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin KewlButtonz.KewlButtons KewlButtons8 
         Height          =   615
         Left            =   4200
         TabIndex        =   16
         ToolTipText     =   "F3-Crear Presupuesto de Venta"
         Top             =   0
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   1085
         BTYPE           =   14
         TX              =   "Consulta Remitos"
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
         BCOL            =   15790320
         BCOLO           =   15790320
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "Inicio.frx":5B830
         PICN            =   "Inicio.frx":5B84C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin KewlButtonz.KewlButtons KewlButtons7 
         Height          =   615
         Left            =   2040
         TabIndex        =   15
         ToolTipText     =   "F7-Panel Caja"
         Top             =   0
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   1085
         BTYPE           =   14
         TX              =   "Emision de Remitos"
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
         BCOL            =   15790320
         BCOLO           =   15790320
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "Inicio.frx":5BC9E
         PICN            =   "Inicio.frx":5BCBA
         PICH            =   "Inicio.frx":618DC
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin KewlButtonz.KewlButtons KewlButtons5 
         Height          =   615
         Left            =   0
         TabIndex        =   14
         ToolTipText     =   "F7-Panel Caja"
         Top             =   0
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   1085
         BTYPE           =   14
         TX              =   "Armado de Pedidos"
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
         BCOL            =   15790320
         BCOLO           =   15790320
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "Inicio.frx":674FE
         PICN            =   "Inicio.frx":6751A
         PICH            =   "Inicio.frx":6796C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Por Nombre"
         Height          =   315
         Left            =   1800
         Picture         =   "Inicio.frx":67DBE
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   1200
         UseMaskColor    =   -1  'True
         Width           =   1215
      End
   End
   Begin VB.Menu preingreso 
      Caption         =   "&Ventas"
      Begin VB.Menu mgestioncomercial 
         Caption         =   "Gestion Comercial"
         Begin VB.Menu wpresupuesto2 
            Caption         =   "Cotización"
            Shortcut        =   {F4}
         End
         Begin VB.Menu mcomparativas 
            Caption         =   "Comparativas"
         End
         Begin VB.Menu mconsultasgc 
            Caption         =   "Consultas"
            Begin VB.Menu wcotibuscar 
               Caption         =   "Cotización"
            End
            Begin VB.Menu mconsultacomparativa 
               Caption         =   "Comparativa"
            End
         End
      End
      Begin VB.Menu gventas 
         Caption         =   "Gestion de Ventas"
         Begin VB.Menu wnotaventa 
            Caption         =   "Nota de Venta"
            Shortcut        =   {F2}
         End
         Begin VB.Menu mfacturacion 
            Caption         =   "Facturacion "
            Begin VB.Menu wfacctacte 
               Caption         =   "En Base a Remito"
            End
            Begin VB.Menu menbasenotadeventa 
               Caption         =   "En Base a Nota de Venta"
            End
         End
         Begin VB.Menu wnotacredito 
            Caption         =   "Nota de Credito"
         End
         Begin VB.Menu wdebito 
            Caption         =   "Nota de Débito"
         End
         Begin VB.Menu mconsultasgv 
            Caption         =   "Consultas"
            Begin VB.Menu mconsultanv 
               Caption         =   "Nota de Venta"
               Shortcut        =   {F3}
            End
            Begin VB.Menu facturabuscar 
               Caption         =   "Facturas"
            End
            Begin VB.Menu wncbusqueda 
               Caption         =   "Notas de Crédito"
            End
            Begin VB.Menu wdebitobusca 
               Caption         =   "Notas de Débito"
            End
            Begin VB.Menu wreporteventas 
               Caption         =   "Reporte de Ventas"
            End
         End
      End
      Begin VB.Menu wpresupuesto 
         Caption         =   "Presupuesto"
         Shortcut        =   {F1}
         Visible         =   0   'False
      End
      Begin VB.Menu wclientes 
         Caption         =   "Clientes"
         Shortcut        =   ^O
      End
      Begin VB.Menu wproveedores 
         Caption         =   "Proveedores"
         Shortcut        =   ^P
      End
      Begin VB.Menu mcontrolafip 
         Caption         =   "Control Facturas AFIP"
      End
   End
   Begin VB.Menu wmcaja 
      Caption         =   "&Caja"
      Begin VB.Menu wdetmov 
         Caption         =   "Detalle de movimientos"
      End
      Begin VB.Menu warcaja 
         Caption         =   "Arqueo de Caja"
      End
      Begin VB.Menu wcierre 
         Caption         =   "Cierre de Caja"
         Shortcut        =   ^T
      End
      Begin VB.Menu wcierrenn 
         Caption         =   "Cierre de Caja (z)"
         Shortcut        =   ^Z
         Visible         =   0   'False
      End
      Begin VB.Menu wingreso 
         Caption         =   "Ingreso de Valores"
         Shortcut        =   ^I
      End
      Begin VB.Menu wingresoz 
         Caption         =   "Ingreso de Valores (Z)"
         Shortcut        =   ^S
         Visible         =   0   'False
      End
      Begin VB.Menu wegreso 
         Caption         =   "Egreso de Valores"
         Shortcut        =   ^E
      End
      Begin VB.Menu wmtransferencia 
         Caption         =   "Transferencia de Valores"
      End
      Begin VB.Menu wegresoZ 
         Caption         =   "Egreso de Valores (z)"
         Shortcut        =   ^X
         Visible         =   0   'False
      End
      Begin VB.Menu wrecibo 
         Caption         =   "Recibo de Cobranza Cta.Cte"
         Shortcut        =   ^R
      End
      Begin VB.Menu wrecibos 
         Caption         =   "Consulta de Cobranzas"
      End
      Begin VB.Menu wsubcompras 
         Caption         =   "Subdiario Compras"
      End
   End
   Begin VB.Menu pinventario 
      Caption         =   "&Deposito"
      Begin VB.Menu wpreppedidos 
         Caption         =   "Armado de Pedidos"
      End
      Begin VB.Menu weremitos 
         Caption         =   "Emisión de Remitos"
      End
      Begin VB.Menu wremvtaconsulta 
         Caption         =   "Remito Consulta"
      End
      Begin VB.Menu wlistpendientessinstock 
         Caption         =   "Listado de Mercaderia a Pedir"
      End
      Begin VB.Menu wlistadopendrecep 
         Caption         =   "Listado de Merc.Pendida a Prov. Pend.de Recepcion"
      End
   End
   Begin VB.Menu mnucomparativa 
      Caption         =   "&Comparativas"
      Visible         =   0   'False
      Begin VB.Menu vistacomparativa 
         Caption         =   "Crear"
      End
      Begin VB.Menu vistabuscarcomparativa 
         Caption         =   "Buscar"
      End
   End
   Begin VB.Menu wseguridad 
      Caption         =   "&Seguridad"
      Begin VB.Menu wusuarios 
         Caption         =   "Usuarios"
      End
      Begin VB.Menu wvendedores 
         Caption         =   "Vendedores"
      End
   End
   Begin VB.Menu acercade 
      Caption         =   "Acerca de"
      Begin VB.Menu acerca 
         Caption         =   "Acerca de"
      End
   End
   Begin VB.Menu salir 
      Caption         =   "SALIR"
   End
End
Attribute VB_Name = "Inicio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public opcion1 As Boolean
Public opcion2 As Boolean
Public salida As Integer
Public montoefectivo As Currency
Dim norest As Integer
Private Declare Function IsIconic Lib "User32" (ByVal hWnd As Long) As Long
Private Declare Function GetAsyncKeyState Lib "User32" (ByVal vKey As Long) As Integer

Private Sub asientosmodelos_Click()

    asienmodelo.Show

End Sub

Private Sub citiventas_Click()

    afipcitiventas.Show

End Sub

Private Sub concbancaria_Click()
On Error Resume Next
    frmconciliacion.Show

End Sub

Private Sub confremoto_Click()

    frmremoto.Show

End Sub

Private Sub consufactcomp_Click()


End Sub

Private Sub consufacventas_Click()

    frmconsutalibroventas.Show

End Sub

Private Sub depuracioncli_Click()

    frmdepclientes.Show

End Sub

Private Sub depuracionprov_Click()

    frmdepproveedores.Show

End Sub

Private Sub facelec_Click()

    afip_facelec.Show

End Sub

Private Sub felecimportar_Click()

    afip_f_electronica.Show


End Sub

Private Sub importaempresa_Click()

    importacuentaotraempresa.Show

End Sub

Private Sub importaplan_Click()

    frmimportaplancuenta.Show

End Sub

Private Sub informeresul_Click()

    frmresultadosactivos.Show

End Sub

Private Sub Command3_Click()

    

End Sub

Private Sub KewlButtons1_Click()

    lista_clientes_consulta.Show

End Sub

Private Sub KewlButtons12_Click()

    lista_notadeventas_consulta.Show

End Sub

Private Sub KewlButtons10_Click()

menu = 1
    frmremitosconsulta.Show

End Sub

Private Sub KewlButtons11_Click()

    menu = 1
    frmcentrodefacturacion.Show

End Sub

Private Sub KewlButtons13_Click()

    lista_proveedores_consulta.Show

End Sub

Private Sub KewlButtons14_Click()

lista_clientes_consulta.Show

End Sub

Private Sub KewlButtons15_Click()


    productoconsulta = ""
    lista_productos_precios.Show


End Sub

Private Sub KewlButtons2_Click()

    menu = 2
    lista_presupuestos_todos.Show
   

End Sub

Private Sub KewlButtons3_Click()

    menu = 1
    frmcentrodefacturacion.Show

End Sub

Private Sub facturabuscar_Click()

    lista_facturas_todas.Show

End Sub

Private Sub KewlButtons4_Click()

    menu = 1
    frmremitosconsulta.Show

End Sub

Private Sub KewlButtons5_Click()

        lista_notadeventas.Show


End Sub

Private Sub KewlButtons6_Click()

    productoconsulta = ""
    lista_productos_precios.Show


End Sub




  

Private Sub KewlButtons7_Click()

lista_pendientesremitir.Show

End Sub

Private Sub KewlButtons8_Click()


    menu = 0
    frmremitosconsulta_remito.Show

End Sub

Private Sub KewlButtons9_Click()

    lista_notadeventas_consulta.Show

End Sub

Private Sub mcomparativas_Click()

    menu = 5
    frmcomparativa.Show

End Sub

Private Sub mconsultacomparativa_Click()

    menu = 6
    lista_presupuestos.Show


End Sub

Private Sub mconsultanv_Click()

    lista_notadeventas_consulta.Show

End Sub

Private Sub mcontrolafip_Click()

afip_control.Show

End Sub

Private Sub MDIForm_Activate()
    
On Error Resume Next

    
parametros1 = Clipboard.GetText
Ancho = Val(Mid(parametros1, 1, 10))
alto = Val(Mid(parametros1, 11, 10))
izq = Val(Right(parametros1, 10))

If alto < 1000 Then
    Inicio.WindowState = 2
Else
    Inicio.Width = Ancho
    Inicio.Height = alto
    Inicio.Left = izq
    Inicio.Top = 0
End If
    
If UCase(login.usuarioactivo) = "CAJA" Then
        Toolbar2.Visible = False
'        Notaventa.Visible = False
           wpresupuesto2.Visible = False
           wnotacredito.Visible = False
           wdebito.Visible = False
        
        presupuesto.Visible = False
        KewlButtons1.Visible = False
        KewlButtons2.Visible = False
        KewlButtons3.Visible = False
'        KewlButtons4.Visible = False   ' modulo caja
        KewlButtons5.Visible = False
        
'        preingreso.Visible = False
        palquileres.Visible = False
'        pinventario.Visible = False
        wseguridad.Visible = False
        
End If
   
If UCase(login.usuarioactivo) = "CONSULTA" Then
        Toolbar2.Visible = False
        Notaventa.Visible = False
        wpresupuesto2.Visible = False
        wnotacredito.Visible = False
        wdebito.Visible = False
        
        presupuesto.Visible = False
        KewlButtons1.Visible = False
        KewlButtons3.Visible = False
        KewlButtons4.Visible = False   ' modulo caja
        KewlButtons5.Visible = False
        
        preingreso.Visible = False
        palquileres.Visible = False
'        pinventario.Visible = False
        wseguridad.Visible = False
        wmcaja.Visible = False
        
End If
   
   
   
If UCase(login.usuarioactivo) = "VENTAS" Then
        Toolbar2.Visible = False
'        Notaventa.Visible = False
           wpresupuesto2.Visible = False
           wnotacredito.Visible = False
           wdebito.Visible = False
        
'        presupuesto.Visible = False
'        KewlButtons1.Visible = False
'        KewlButtons3.Visible = False
'        KewlButtons4.Visible = False   ' modulo Caja.
        KewlButtons5.Visible = False
        
'        preingreso.Visible = False
        palquileres.Visible = False
 '       pinventario.Visible = False
        wseguridad.Visible = False
        
        wmcaja.Visible = False
        
End If
   
If UCase(login.usuarioactivo) = "DEPOSITO" Then
        Toolbar1.Visible = False
        Toolbar2.Visible = True
           wpresupuesto2.Visible = False
           wnotacredito.Visible = False
           wdebito.Visible = False
           preingreso.Visible = False
        
        presupuesto.Visible = False
        Notaventa.Visible = False
        preppedidos.Visible = True
        preppedidos.Left = 0
        
'        preingreso.Visible = False
        palquileres.Visible = False
   '     pinventario.Visible = False
        wseguridad.Visible = False
        
        wmcaja.Visible = False
        
End If
   
If UCase(login.usuarioactivo) = "ADMIN" Or UCase(login.usuarioactivo) = "DELIA" Or UCase(login.usuarioactivo) = "GRACIELA" Then
    KewlButtons9.Visible = False
    KewlButtons10.Visible = False
    KewlButtons11.Visible = False
    KewlButtons15.Visible = False
    KewlButtons14.Visible = False
    
End If
    

If UCase(login.usuarioactivo) <> "ADMIN" Then
    wseguridad.Visible = False
End If



    
End Sub

' Make the image fit the MDI form.
Private Sub MDIForm_Resize()
On Error Resume Next
'    If login.nomsucursal = "EMPORIOZIP" Or login.nomsucursal = "TUCUMANZIP" Then
'        picOriginal.Picture = LoadPicture(App.Path + "\fondo2.jpg")
'        Inicio.Icon = LoadPicture(App.Path + "\COMMIT.ICO")
'    End If

    picStretched.Move 0, 0, _
       ScaleWidth, ScaleHeight

    ' Copy the original picture into picStretched.
    picStretched.PaintPicture _
        picOriginal.Picture, _
        0, 0, _
        picStretched.ScaleWidth, _
        picStretched.ScaleHeight, _
        0, 0, _
        picOriginal.ScaleWidth, _
        picOriginal.ScaleHeight
        
    ' Set the MDI form's picture.
    Picture = picStretched.Image

If menu_administrador.Visible = True Then
    menu_administrador.Top = 0
    menu_administrador.Left = 0

    menu_administrador.Height = Inicio.Height - 1600
    menu_administrador.TreeView1.Height = menu_administrador.Height - 50
End If
    
End Sub
Private Sub acerca_Click()
        frmAbout.Show
End Sub




Private Sub anulorden_Click()

    frmordenanula.Show

End Sub

Private Sub anurecibo_Click()

    frmreciboanula.Show

End Sub

Private Sub archclientes_Click()

    frmclientes.Show

End Sub

Private Sub archproveed_Click()

    frmproveedores.Show

End Sub

Private Sub asientos_Click()

End Sub

Private Sub asigcomp_Click()

    frmordendepagoasigna.Show

End Sub

Private Sub asigrecibo_Click()

    frmrecibocobroasigna.Show

End Sub

Private Sub cambianperiodo_Click()

    frminicioperiodo.Show

End Sub

Private Sub cargaasientos_Click()

   frmasientos.Show

End Sub

Private Sub cargalibrocompras_Click()

    frmlibrocompras_nuevo.Show

End Sub

Private Sub cargalilbroventas_Click()

    frmlibroventas_nuevo.Show

End Sub


Private Sub cierrelibrocompras_Click()
    
    frmclcompras.Show

End Sub

Private Sub cierrelibroventas_Click()

    frmclventas.Show

End Sub


Private Sub confempresa_Click()
    
    frmEMPRESA.Show

End Sub


Private Sub consfacturas_Click()

    frmfacturaconsulta.Show

End Sub

Private Sub consord_Click()

    frmordendepagoconsulta.Show

End Sub

Private Sub consultarecibos_Click()

    frmreciboconsulta.Show

End Sub


Private Sub login_Click()

    login.Show

End Sub

Private Sub ecproveedores_Click()

   

End Sub


Private Sub ecclientes1_Click()

End Sub

Private Sub ecclientes_Click()

    impecclientes.Show

End Sub

Private Sub emitirfactura_Click()
    frmfacclientesnuevo.Show
End Sub

Private Sub emitirrecibo_Click()

    frmrecibos1.Show

End Sub

Private Sub estcuentas_Click()

    impecproved.Show

End Sub

Private Sub facclientesboton_Click()

    frmfacclientes.Show

End Sub

Private Sub importalibro_Click()

    importalibroventas.Show

End Sub

Private Sub librodiario_Click()

    implibrodiario.Show

End Sub

Private Sub listadolibrocompras_Click()
    
    implibrocompras.Show

End Sub

Private Sub listadolibroventas_Click()

    implibroventas.Show
    
End Sub

Private Sub listadorecibos_Click()

    imprecibolistado.Show

End Sub

Private Sub listordenes_Click()

    impordeneslistado.Show

End Sub

Private Sub loginear_Click()
     login.Show
End Sub

Private Sub mayoranalitico_Click()

    impmayoranalitico.Show

End Sub

Private Sub MDIForm_Unload(Cancel As Integer)

        Cancel = salida
        frmsalir.Show
        End


End Sub

Private Sub menajusteclientes_Click()

    frmajusteclientes.Show

End Sub

Private Sub menajusteproveedores_Click()

    frmajusteproveedores.Show

End Sub

Private Sub menuauditoria_Click()

    frmauditoria.Show

End Sub

Private Sub menuchequesemitidos_Click()

    frmcarteraemitidos.Show
    

End Sub

Private Sub menuchequesencartera_Click()
    
    frmchequescancelados.Show
    
End Sub

Private Sub menuconceptos_Click()

    frmcajaconceptos.Show

End Sub

Private Sub menuimporta_Click()

    importalibroventas.Show

End Sub

Private Sub menuimpresoras_Click()

    CommonDialog1.ShowPrinter
    CommonDialog1.PrinterDefault = True

End Sub

Private Sub mnuprueba_Click()

    Form1.Show

End Sub

Private Sub menulibros_Click()

    frmcajabanco.Show

End Sub

Private Sub Option1_Click()
    If Option1.Value = True Then
        opcion1 = True
        opcion2 = False
    End If
End Sub

Private Sub Option2_Click()

    
    If Option2.Value = True Then
        opcion2 = True
        opcion1 = False
    End If


End Sub

Private Sub ordpago_Click()

    frmordendepago1.Show

End Sub

Private Sub otparam_Click()

    frmotrosparam.Show

End Sub

Private Sub otrosgastos_Click()

    frmotrosgastos_nuevo.Show

End Sub


Private Sub perdefecto_Click()

    frmperiododefecto.Show

End Sub

Private Sub plandecuentas_Click()
    

    frmCuentas.Show
    
    
End Sub

Private Sub plantipo_Click()

 importacuenta.Show

End Sub


Private Sub productos_Click()

    frmarticulos.Show

End Sub

Private Sub menbasenotadeventa_Click()

    menu = 1
    frmcentrodefacturacion.Show

End Sub

Private Sub menulistadomercpend_Click()
'On Error GoTo fuera
Dim crApp As New CRAXDRT.Application
Dim crReport As New CRAXDRT.Report
Dim tabla As String
Dim ruta As String

reporte.SQL = "SELECT v_ezi_pos_ctacte.NV, v_ezi_pos_ctacte.fechadelcomprobante, v_ezi_pos_ctacte.codigoproducto, v_ezi_pos_ctacte.nombre_producto, v_ezi_pos_ctacte.CantidadOrigen, v_ezi_pos_ctacte.Remitido, v_ezi_pos_ctacte.pendiente, v_ezi_pos_ctacte.PROVEEDOR_N, v_ezi_pos_ctacte.STK FROM MMOSSE.dbo.v_ezi_pos_pendientesdeentregar v_ezi_pos_ctacte ORDER BY v_ezi_pos_ctacte.PROVEEDOR_N ASC"
tabla = reporte.SQL

With CrystalReporte
    .PrinterCollation = crptCollated
    .ReportFileName = App.Path & "\ReportePendientesMercaderia.rpt"
    .WindowTitle = "Listado de Mercaderia Pendiente"
    '.Connect = "PROVIDER=MSDASQL;dsn=facturacion;uid=lucva;pwd=25072004;database=facturacionsql;"
    .Connect = login.conexionreporte
    .DiscardSavedData = True
    .RetrieveDataFiles
    .ReportSource = 0
    .SQLQuery = tabla
    .Destination = crptToWindow
 '   .Destination = crptToPrinter
    .PrintFileType = crptCrystal
    .WindowState = crptMaximized
    .Action = 1
End With


End Sub

Private Sub Notaventa_Click()

    tipodeventa = 0
    frmnota_venta.Show

End Sub

Private Sub preppedidos_Click()

    lista_notadeventas.Show

End Sub

Private Sub presupuesto_Click()

    frmpresupuesto.Show

End Sub

Private Sub salir_Click()

    
    frmsalir.Show

End Sub


Private Sub sumasysaldos_Click()

    impsumasysaldos.Show

End Sub

Private Sub sumasysaldosconcc_Click()

    impsumasysaldoscc.Show

End Sub

Private Sub Timer1_Timer()
On Error Resume Next

'If login.nomsucursal = "EMPORIOZIP" Or login.nomsucursal = "TUCUMANZIP" Then
'   xcuenta = "CUENTA 2 (DOS)"
'Else
'   xcuenta = "CUENTA 1 (UNO)"
'End If
   

Inicio.Caption = xcuenta + "     " + Str(Time) + "   Usuario: " + login.usuarioactivo + " BaseDatos: " + login.nombrebd + "  Sucursal: " + login.nomsucursal


Dim X As Long
    'para los 255 códigos Ascii
    For X = 113 To 120
        'si se ha pulsado una tecla la APi devuelve -32737
        If GetAsyncKeyState(X) = -32767 Then
            'Si se presiono la tecla F9 mostramos el mensaje
            If X = 113 Then
                tipodeventa = 0
                frmnota_venta.Show
            End If
            If X = 115 Then
                frmpresupuesto.Show
            End If
        End If
    Next
               
End Sub


Private Sub verasientos_Click()

    frmasientosbusca.Show

End Sub

Private Sub wpesadacania_Click()

    frmpesada_cania.Show
    

End Sub

Private Sub wpreingreso_Click()

    frmpreingreso.Show

End Sub

Private Sub wtaracania_Click()

    frmtara_cania.Show

End Sub

Private Sub walquileres_Click()

    frmalquiler.Show

End Sub

Private Sub vistabuscarcomparativa_Click()

    menu = 6
    lista_presupuestos.Show

End Sub

Private Sub vistacomparativa_Click()

    menu = 5
    frmcomparativa.Show

End Sub

Private Sub wcierre_Click()

    frmcierre_caja.Show

End Sub

Private Sub wcierrenn_Click()

    frmcierre_cajaNN.Show

End Sub

Private Sub wclientes_Click()

    lista_clientes_consulta.Show

End Sub

Private Sub wdealquiler_Click()

    menu = 1
    frmremitosdevolucion_alquiler.Show

End Sub

Private Sub wcotibuscar_Click()

    menu = 2
    lista_presupuestos_todos.Show

End Sub

Private Sub wdebito_Click()

    frmnota_debito.Show

End Sub

Private Sub wdebitobusca_Click()

    lista_debitos_todas.Show

End Sub

Private Sub wdetmov_Click()

    Informe_Caja_detallado.Show

End Sub

Private Sub wegreso_Click()

    frmegresovalor.Show

End Sub

Private Sub wegresoZ_Click()

    frmegresovalorNN.Show

End Sub

Private Sub wfacalquiler_Click()


    menu = 1
    frmremitosconsulta_alquiler.Show
    

End Sub

Private Sub wfacanomala_Click()

    menu = 1
    frmremitosdevolucion_alquiler_fa.Show

End Sub

Private Sub weremitos_Click()

    lista_pendientesremitir.Show

End Sub

Private Sub wfacctacte_Click()

    menu = 1
    frmremitosconsulta.Show
    

End Sub

Private Sub wingreso_Click()

    frmingreovalor.Show
    
End Sub

Private Sub wingresoz_Click()

        frmingreovalorNN.Show

End Sub

Private Sub wlistadopendrecep_Click()
'On Error GoTo fuera
Dim crApp As New CRAXDRT.Application
Dim crReport As New CRAXDRT.Report
Dim tabla As String
Dim ruta As String

reporte.SQL = "SELECT * FROM MMOSSE.dbo.v_ezi_pos_pendientesderecepcion v_ezi_pos_ctacte ORDER BY v_ezi_pos_ctacte.PROVEEDOR_N ASC"
tabla = reporte.SQL

With CrystalReporte
    .PrinterCollation = crptCollated
    .ReportFileName = App.Path & "\ReportePendientesMercaderiasinSR.rpt"
    .WindowTitle = "Listado de Mercaderia Pendiente De Recepcion"
    '.Connect = "PROVIDER=MSDASQL;dsn=facturacion;uid=lucva;pwd=25072004;database=facturacionsql;"
    .Connect = login.conexionreporte
    .DiscardSavedData = True
    .RetrieveDataFiles
    .ReportSource = 0
    .SQLQuery = tabla
    .Destination = crptToWindow
 '   .Destination = crptToPrinter
    .PrintFileType = crptCrystal
    .WindowState = crptMaximized
    .Action = 1
End With




End Sub

Private Sub wlistadopensinoc_Click()
On Error Resume Next
Dim crApp As New CRAXDRT.Application
Dim crReport As New CRAXDRT.Report
Dim tabla As String
Dim ruta As String

reporte.SQL = "SELECT * FROM MMOSSE.dbo.v_ezi_pos_pendienesdeentregar_sinoc v_ezi_pos_ctacte ORDER BY v_ezi_pos_ctacte.PROVEEDOR_N ASC"
tabla = reporte.SQL

With CrystalReporte
    .PrinterCollation = crptCollated
    .ReportFileName = App.Path & "\ReportePendientesMercaderiasinOC.rpt"
    .WindowTitle = "Listado de Mercaderia Pendiente Sin OC"
    '.Connect = "PROVIDER=MSDASQL;dsn=facturacion;uid=lucva;pwd=25072004;database=facturacionsql;"
    .Connect = login.conexionreporte
    .DiscardSavedData = True
    .RetrieveDataFiles
    .ReportSource = 0
    .SQLQuery = tabla
    .Destination = crptToWindow
 '   .Destination = crptToPrinter
    .PrintFileType = crptCrystal
    .WindowState = crptMaximized
    .Action = 1
End With


End Sub

Private Sub wlistpendientessinstock_Click()
On Error Resume Next
Dim crApp As New CRAXDRT.Application
Dim crReport As New CRAXDRT.Report
Dim tabla As String
Dim ruta As String

reporte.SQL = "SELECT * FROM MMOSSE.dbo.v_ezi_pos_pendientessinstock v_ezi_pos_ctacte ORDER BY v_ezi_pos_ctacte.PROVEEDOR_N, v_ezi_pos_ctacte.nv ASC"
tabla = reporte.SQL

With CrystalReporte
    .PrinterCollation = crptCollated
    .ReportFileName = App.Path & "\ReportePendientesMercaderiasinSinStock.rpt"
    .WindowTitle = "Listado de Mercaderia Pendiente Sin OC"
    '.Connect = "PROVIDER=MSDASQL;dsn=facturacion;uid=lucva;pwd=25072004;database=facturacionsql;"
    .Connect = login.conexionreporte
    .DiscardSavedData = True
    .RetrieveDataFiles
    .ReportSource = 0
    .SQLQuery = tabla
    .Destination = crptToWindow
 '   .Destination = crptToPrinter
    .PrintFileType = crptCrystal
    .WindowState = crptMaximized
    .Action = 1
End With



End Sub

Private Sub wmtransferencia_Click()

    frmtransferenciavalor.Show

End Sub

Private Sub wncbusqueda_Click()

    lista_notacredito_todas.Show

End Sub

Private Sub wnotacredito_Click()

    frmnota_credito.Show

End Sub

Private Sub wnotaventa_Click()

    tipodeventa = 0
    frmnota_venta.Show

End Sub

Private Sub wpreppedidos_Click()

    lista_notadeventas.Show

End Sub

Private Sub wpresupuesto_Click()
    
    tipodeventa = 1
    frmnota_venta.Show

End Sub

Private Sub wpresupuesto2_Click()

    frmpresupuesto.Show

End Sub

Private Sub wproveedores_Click()

    lista_proveedores_consulta.Show

End Sub

Private Sub wrecibo_Click()

    frmrecibo_ctacte.Show

End Sub

Private Sub wrecibos_Click()

    lista_recibo_todos.Show

End Sub

Private Sub wremvtaconsulta_Click()

    menu = 0
    frmremitosconsulta_remito.Show

End Sub

Private Sub wreporteventas_Click()

    lista_ventas_reporte.Show

End Sub

Private Sub wsubcompras_Click()

    lista_subdiariocompras.Show

End Sub

Private Sub wvendedores_Click()

    frmvendedores.Show

End Sub
