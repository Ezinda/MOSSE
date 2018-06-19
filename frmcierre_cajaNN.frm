VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{D18BBD1F-82BB-4385-BED3-E9D31A3E361E}#1.0#0"; "KewlButtonz.ocx"
Begin VB.Form frmcierre_cajaNN 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CIERRE DE CAJA"
   ClientHeight    =   8145
   ClientLeft      =   45
   ClientTop       =   480
   ClientWidth     =   14265
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8145
   ScaleWidth      =   14265
   Begin VB.CommandButton Command7 
      Cancel          =   -1  'True
      Caption         =   "Total Transferido"
      Height          =   375
      Left            =   8520
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   7560
      Width           =   1575
   End
   Begin VB.TextBox Text1 
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
      Height          =   375
      Index           =   9
      Left            =   6960
      Locked          =   -1  'True
      TabIndex        =   25
      Top             =   7560
      Width           =   1575
   End
   Begin VB.TextBox Text1 
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
      Height          =   375
      Index           =   8
      Left            =   6960
      Locked          =   -1  'True
      TabIndex        =   24
      Top             =   7080
      Width           =   1575
   End
   Begin VB.TextBox Text1 
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
      Height          =   375
      Index           =   7
      Left            =   6960
      Locked          =   -1  'True
      TabIndex        =   23
      Top             =   6720
      Width           =   1575
   End
   Begin VB.TextBox Text1 
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
      Height          =   375
      Index           =   6
      Left            =   6960
      Locked          =   -1  'True
      TabIndex        =   22
      Top             =   6360
      Width           =   1575
   End
   Begin VB.TextBox Text1 
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
      Height          =   375
      Index           =   5
      Left            =   6960
      Locked          =   -1  'True
      TabIndex        =   21
      Top             =   6000
      Width           =   1575
   End
   Begin VB.TextBox Text1 
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
      Height          =   375
      Index           =   4
      Left            =   4080
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   7560
      Width           =   1575
   End
   Begin VB.TextBox Text1 
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
      Height          =   375
      Index           =   3
      Left            =   4080
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   7080
      Width           =   1575
   End
   Begin VB.TextBox Text1 
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
      Height          =   375
      Index           =   2
      Left            =   4080
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   6720
      Width           =   1575
   End
   Begin VB.TextBox Text1 
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
      Height          =   375
      Index           =   1
      Left            =   4080
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   6360
      Width           =   1575
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Saldo sin Transferir"
      Height          =   375
      Left            =   2520
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   7560
      Width           =   1575
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Tot.Otros:"
      Height          =   375
      Left            =   2520
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   7080
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Tot.Cheques:"
      Height          =   375
      Left            =   2520
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   6720
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Tot.Tarjetas:"
      Height          =   375
      Left            =   2520
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   6360
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Tot.Efectivo:"
      Height          =   375
      Left            =   2520
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   6000
      Width           =   1575
   End
   Begin VB.TextBox Text1 
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
      Height          =   375
      Index           =   0
      Left            =   4080
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   6000
      Width           =   1575
   End
   Begin VB.CommandButton salir 
      Caption         =   "salir"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   6000
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmcierre_cajaNN.frx":0000
      Height          =   5175
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Visible         =   0   'False
      Width           =   14055
      _ExtentX        =   24791
      _ExtentY        =   9128
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      HeadLines       =   1
      RowHeight       =   19
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
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
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   540
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   14265
      _ExtentX        =   25162
      _ExtentY        =   953
      ButtonWidth     =   609
      ButtonHeight    =   900
      Appearance      =   1
      Style           =   1
      _Version        =   393216
      Begin KewlButtonz.KewlButtons traervalores 
         Height          =   375
         Left            =   4320
         TabIndex        =   5
         ToolTipText     =   "F2-Crea Nota de Venta"
         Top             =   120
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         BTYPE           =   14
         TX              =   "&Traer Valores"
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
         MICON           =   "frmcierre_cajaNN.frx":001A
         PICN            =   "frmcierre_cajaNN.frx":0036
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   1920
         TabIndex        =   4
         Top             =   120
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   135856129
         CurrentDate     =   42037
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Fecha de Cierre:"
         Height          =   375
         Left            =   120
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   120
         Width           =   1575
      End
   End
   Begin MSAdodcLib.Adodc datproducto 
      Height          =   330
      Left            =   120
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
   Begin KewlButtonz.KewlButtons transferir 
      Height          =   375
      Index           =   0
      Left            =   5880
      TabIndex        =   16
      ToolTipText     =   "F2-Crea Nota de Venta"
      Top             =   6000
      Width           =   855
      _ExtentX        =   1508
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
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmcierre_cajaNN.frx":05D0
      PICN            =   "frmcierre_cajaNN.frx":05EC
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin KewlButtonz.KewlButtons transferir 
      Height          =   375
      Index           =   1
      Left            =   5880
      TabIndex        =   17
      ToolTipText     =   "F2-Crea Nota de Venta"
      Top             =   6360
      Width           =   855
      _ExtentX        =   1508
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
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmcierre_cajaNN.frx":0B86
      PICN            =   "frmcierre_cajaNN.frx":0BA2
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin KewlButtonz.KewlButtons transferir 
      Height          =   375
      Index           =   2
      Left            =   5880
      TabIndex        =   18
      ToolTipText     =   "F2-Crea Nota de Venta"
      Top             =   6720
      Width           =   855
      _ExtentX        =   1508
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
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmcierre_cajaNN.frx":113C
      PICN            =   "frmcierre_cajaNN.frx":1158
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin KewlButtonz.KewlButtons transferir 
      Height          =   375
      Index           =   3
      Left            =   5880
      TabIndex        =   19
      ToolTipText     =   "F2-Crea Nota de Venta"
      Top             =   7080
      Width           =   855
      _ExtentX        =   1508
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
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmcierre_cajaNN.frx":16F2
      PICN            =   "frmcierre_cajaNN.frx":170E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin KewlButtonz.KewlButtons transferir 
      Height          =   375
      Index           =   4
      Left            =   5880
      TabIndex        =   20
      ToolTipText     =   "F2-Crea Nota de Venta"
      Top             =   7560
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BTYPE           =   14
      TX              =   "&Cierre"
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
      MICON           =   "frmcierre_cajaNN.frx":1CA8
      PICN            =   "frmcierre_cajaNN.frx":1CC4
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
      Height          =   735
      Left            =   10440
      TabIndex        =   27
      ToolTipText     =   "F2-Crea Nota de Venta"
      Top             =   7200
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1296
      BTYPE           =   14
      TX              =   "&Aceptar"
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
      MICON           =   "frmcierre_cajaNN.frx":225E
      PICN            =   "frmcierre_cajaNN.frx":227A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin KewlButtonz.KewlButtons deshacer 
      Height          =   375
      Index           =   0
      Left            =   5880
      TabIndex        =   28
      ToolTipText     =   "F2-Crea Nota de Venta"
      Top             =   6000
      Width           =   855
      _ExtentX        =   1508
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
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmcierre_cajaNN.frx":2814
      PICN            =   "frmcierre_cajaNN.frx":2830
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin KewlButtonz.KewlButtons deshacer 
      Height          =   375
      Index           =   1
      Left            =   5880
      TabIndex        =   29
      ToolTipText     =   "F2-Crea Nota de Venta"
      Top             =   6360
      Width           =   855
      _ExtentX        =   1508
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
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmcierre_cajaNN.frx":2DCA
      PICN            =   "frmcierre_cajaNN.frx":2DE6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin KewlButtonz.KewlButtons deshacer 
      Height          =   375
      Index           =   2
      Left            =   5880
      TabIndex        =   30
      ToolTipText     =   "F2-Crea Nota de Venta"
      Top             =   6720
      Width           =   855
      _ExtentX        =   1508
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
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmcierre_cajaNN.frx":3380
      PICN            =   "frmcierre_cajaNN.frx":339C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin KewlButtonz.KewlButtons deshacer 
      Height          =   375
      Index           =   3
      Left            =   5880
      TabIndex        =   31
      ToolTipText     =   "F2-Crea Nota de Venta"
      Top             =   7080
      Width           =   855
      _ExtentX        =   1508
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
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmcierre_cajaNN.frx":3936
      PICN            =   "frmcierre_cajaNN.frx":3952
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin KewlButtonz.KewlButtons deshacer 
      Height          =   375
      Index           =   4
      Left            =   5880
      TabIndex        =   32
      ToolTipText     =   "F2-Crea Nota de Venta"
      Top             =   7560
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BTYPE           =   14
      TX              =   "&Cierre"
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
      MICON           =   "frmcierre_cajaNN.frx":3EEC
      PICN            =   "frmcierre_cajaNN.frx":3F08
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSAdodcLib.Adodc datcierre 
      Height          =   330
      Left            =   120
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
   Begin MSAdodcLib.Adodc datcierre2 
      Height          =   330
      Left            =   120
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
   Begin MSAdodcLib.Adodc datparametros 
      Height          =   330
      Left            =   240
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
End
Attribute VB_Name = "frmcierre_cajaNN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public cuentacont As String
Dim cuenta(99999) As Integer
Dim xefectivo As Currency



Private Sub DataGrid1_DblClick()

'            frmnota_venta.Text1(1).Text = DataGrid1.Columns(2).Text
'            SendKeys "{ENTER}", False
'        Unload Me

End Sub

Private Sub deshacer_Click(Index As Integer)

If Index = 4 Then
    If Val(Format(Text1(5).Text, "0.00")) > 0 Then
        Text1(0).Text = Format(Text1(5).Text, "###,##0.00")
    End If
    If Val(Format(Text1(6).Text, "0.00")) > 0 Then
        Text1(1).Text = Format(Text1(6).Text, "###,##0.00")
    End If
    If Val(Format(Text1(7).Text, "0.00")) > 0 Then
        Text1(2).Text = Format(Text1(7).Text, "###,##0.00")
    End If
    If Val(Format(Text1(8).Text, "0.00")) > 0 Then
        Text1(3).Text = Format(Text1(8).Text, "###,##0.00")
    End If
    If Val(Format(Text1(9).Text, "0.00")) > 0 Then
        Text1(4).Text = Format(Text1(9).Text, "###,##0.00")
    End If
    
    For X = 5 To 9
        Text1(X).Text = Format(0, "###,##0.00")
        transferir(X - 5).Visible = True
        deshacer(X - 5).Visible = False
    Next X
Else
    Text1(Index).Text = Format(Text1(Index + 5).Text, "###,##0.00")
    transferir(Index).Visible = True
    deshacer(Index).Visible = False
    Text1(Index + 5).Text = Format(0, "###,##0.00")
End If

    xvalortotal = 0
    xvaloranterior = 0
    For X = 5 To 8
        xvalortotal = xvalortotal + Val(Format(Text1(X).Text, "0.00"))
        xvaloranterior = xvaloranterior + Val(Format(Text1(X - 5).Text, "0.00"))
    Next X

    If xvalortotal = 0 Then
        deshacer(4).Visible = False
        transferir(4).Visible = True
    End If
    
    Text1(4).Text = Format(xvaloranterior, "###,##0.00")
    Text1(9).Text = Format(xvalortotal, "###,##0.00")
        


End Sub

Private Sub Form_Load()
Aplicar_skin Me

MiFuncionDeAjuste Me, True

yventana = Inicio.Height / 2 - 1000
xventana = Inicio.Width / 2

frmcierre_cajaNN.Top = yventana - frmalquiler_devolucion.Height / 2
frmcierre_cajaNN.Left = xventana - frmalquiler_devolucion.Width / 2

datproducto.ConnectionString = login.conexiontotal
datcierre.ConnectionString = login.conexiontotal
datparametros.ConnectionString = login.conexiontotal


datparametros.RecordSource = "select * from ud_ezi_parametros_pos where sucursal = '" & login.nomsucursal & "' "
datparametros.Refresh
DTPicker1.Value = Date

 
End Sub

Private Sub Notaventa_Click()



End Sub

Private Sub KewlButtons1_Click()
On Error GoTo fuera

fechacierre = DTPicker1.Value + 1
fechacierre1 = Right("0" + Replace(Str(Day(fechacierre)), " ", ""), 2) + "/" + Right("0" + Replace(Str(Month(fechacierre)), " ", ""), 2) + "/" + Replace(Str(Year(fechacierre)), " ", "")
xquery = ""
xquery1 = ""
xquery2 = ""
XQUERY3 = ""
xuni1 = ""
xuni2 = ""
xuni3 = ""

    

    
If Val(Format(Text1(5).Text, "0.00")) <> 0 Then
 xquery = "SELECT     ud_ezi_pago.id, ud_ezi_pago.formadepago, V_TARJETACREDITO_.NOMBRE AS Tarjeta, ud_ezi_pago.cuotas, ud_ezi_pago.numerodecupon AS NroCupon, " & _
         "ud_ezi_pago.numero AS NroCheque, ud_ezi_pago.fechadeemision, ud_ezi_pago.fechadevencimiento, ud_ezi_pago.monto, ud_ezi_pago.transferido,  " & _
         "V_CAJA_.DESCRIPCION AS caja, ud_ezi_pago.claveprimaria, ud_ezi_puntodeventa_encabezado.tipodefacturacionid  " & _
         "FROM         V_CAJA_ RIGHT OUTER JOIN  " & _
         "ud_ezi_pago WITH (readpast) ON V_CAJA_.ID = ud_ezi_pago.destinoid RIGHT OUTER JOIN  " & _
         "ud_ezi_puntodeventa_encabezado WITH (readpast) ON ud_ezi_pago.claveprimaria = ud_ezi_puntodeventa_encabezado.id LEFT OUTER JOIN  " & _
         "V_TARJETACREDITO_ ON ud_ezi_pago.tarjetaid = V_TARJETACREDITO_.ID  " & _
         "WHERE     (ud_ezi_puntodeventa_encabezado.fechadelcomprobante < CONVERT(DATETIME, '" & fechacierre1 & "', 103)) AND (ud_ezi_pago.transferido IS NULL) AND  " & _
         "             (ud_ezi_pago.formadepago = 'Efectivo') AND ( ud_ezi_pago.sucursal = '" & login.nomsucursal & "')  "
End If

If Val(Format(Text1(6).Text, "0.00")) <> 0 Then
  xquery1 = "SELECT   ud_ezi_pago.id,  ud_ezi_pago.formadepago, V_TARJETACREDITO_.NOMBRE AS Tarjeta, ud_ezi_pago.cuotas, ud_ezi_pago.numerodecupon AS NroCupon, " & _
         "ud_ezi_pago.numero AS NroCheque, ud_ezi_pago.fechadeemision, ud_ezi_pago.fechadevencimiento, ud_ezi_pago.monto, ud_ezi_pago.transferido,  " & _
         "V_CAJA_.DESCRIPCION AS caja, ud_ezi_pago.claveprimaria, ud_ezi_puntodeventa_encabezado.tipodefacturacionid  " & _
         "FROM         V_CAJA_ RIGHT OUTER JOIN  " & _
         "ud_ezi_pago WITH (readpast) ON V_CAJA_.ID = ud_ezi_pago.destinoid RIGHT OUTER JOIN  " & _
         "ud_ezi_puntodeventa_encabezado WITH (readpast) ON ud_ezi_pago.claveprimaria = ud_ezi_puntodeventa_encabezado.id LEFT OUTER JOIN  " & _
         "V_TARJETACREDITO_ ON ud_ezi_pago.tarjetaid = V_TARJETACREDITO_.ID  " & _
         "WHERE     (ud_ezi_puntodeventa_encabezado.fechadelcomprobante < CONVERT(DATETIME, '" & fechacierre1 & "', 103)) AND (ud_ezi_pago.transferido IS NULL) AND  " & _
         "             (ud_ezi_pago.formadepago like '%tarjeta%') AND ( ud_ezi_pago.sucursal = '" & login.nomsucursal & "')  "
End If

If Val(Format(Text1(7).Text, "0.00")) <> 0 Then
  xquery2 = "SELECT    ud_ezi_pago.id, ud_ezi_pago.formadepago, V_TARJETACREDITO_.NOMBRE AS Tarjeta, ud_ezi_pago.cuotas, ud_ezi_pago.numerodecupon AS NroCupon, " & _
         "ud_ezi_pago.numero AS NroCheque, ud_ezi_pago.fechadeemision, ud_ezi_pago.fechadevencimiento, ud_ezi_pago.monto, ud_ezi_pago.transferido,  " & _
         "V_CAJA_.DESCRIPCION AS caja, ud_ezi_pago.claveprimaria, ud_ezi_puntodeventa_encabezado.tipodefacturacionid  " & _
         "FROM         V_CAJA_ RIGHT OUTER JOIN  " & _
         "ud_ezi_pago WITH (readpast) ON V_CAJA_.ID = ud_ezi_pago.destinoid RIGHT OUTER JOIN  " & _
         "ud_ezi_puntodeventa_encabezado WITH (readpast) ON ud_ezi_pago.claveprimaria = ud_ezi_puntodeventa_encabezado.id LEFT OUTER JOIN  " & _
         "V_TARJETACREDITO_ ON ud_ezi_pago.tarjetaid = V_TARJETACREDITO_.ID  " & _
         "WHERE     (ud_ezi_puntodeventa_encabezado.fechadelcomprobante < CONVERT(DATETIME, '" & fechacierre1 & "', 103)) AND (ud_ezi_pago.transferido IS NULL) AND  " & _
         "             (ud_ezi_pago.formadepago like '%cheque%') AND ( ud_ezi_pago.sucursal = '" & login.nomsucursal & "')  "
End If

If Val(Format(Text1(8).Text, "0.00")) <> 0 Then
  XQUERY3 = "SELECT    ud_ezi_pago.id, ud_ezi_pago.formadepago, V_TARJETACREDITO_.NOMBRE AS Tarjeta, ud_ezi_pago.cuotas, ud_ezi_pago.numerodecupon AS NroCupon, " & _
         "ud_ezi_pago.numero AS NroCheque, ud_ezi_pago.fechadeemision, ud_ezi_pago.fechadevencimiento, ud_ezi_pago.monto, ud_ezi_pago.transferido,  " & _
         "V_CAJA_.DESCRIPCION AS caja, ud_ezi_pago.claveprimaria, ud_ezi_puntodeventa_encabezado.tipodefacturacionid  " & _
         "FROM         V_CAJA_ RIGHT OUTER JOIN  " & _
         "ud_ezi_pago WITH (readpast) ON V_CAJA_.ID = ud_ezi_pago.destinoid RIGHT OUTER JOIN  " & _
         "ud_ezi_puntodeventa_encabezado WITH (readpast) ON ud_ezi_pago.claveprimaria = ud_ezi_puntodeventa_encabezado.id LEFT OUTER JOIN  " & _
         "V_TARJETACREDITO_ ON ud_ezi_pago.tarjetaid = V_TARJETACREDITO_.ID  " & _
         "WHERE     (ud_ezi_puntodeventa_encabezado.fechadelcomprobante < CONVERT(DATETIME, '" & fechacierre1 & "', 103)) AND (ud_ezi_pago.transferido IS NULL) AND  " & _
         "             (ud_ezi_pago.formadepago like '%ret%') AND ( ud_ezi_pago.sucursal = '" & login.nomsucursal & "')  "
End If


If xquery <> "" And (xquery1 <> "" Or xquery2 <> "" Or XQUERY3 <> "") Then
    xuni1 = "Union all "
End If
    
If xquery = "" And xquery1 <> "" And (xquery2 <> "" Or XQUERY3 <> "") Then
    xuni2 = "Union all "
End If

If xquery2 <> "" And XQUERY3 <> "" Then
    xuni3 = "Union all "
End If

'If xquery <> "" And xquery1 <> "" And xquery2 <> "" Then
'    xuni2 = "Union all "
'End If

xqueryfinal = xquery + xuni1 + xquery1 + xuni2 + xquery2 + xuni3 + XQUERY3
datcierre.RecordSource = xqueryfinal
datcierre.Refresh


Debug.Print xqueryfinal

If datcierre.Recordset.EOF = False Then
  datcierre2.ConnectionString = login.conexiontotal
  datcierre2.RecordSource = "select isnull(max(idtransferencia),0) as idtransferencia from ud_ezi_pago"
  datcierre2.Refresh
  xnrocierre = datcierre2.Recordset.Fields("idtransferencia") + 1
  
  Do While Not datcierre.Recordset.EOF
            xclave = datcierre.Recordset.Fields("id")
            datcierre2.RecordSource = "select * from  ud_ezi_pago with (readpast) where id = " & xclave & " "
            datcierre2.Refresh
            
            If datcierre2.Recordset.EOF = False Then
                datcierre2.Recordset.Fields("transferido") = 1
                datcierre2.Recordset.Fields("fechacierrre") = DateValue(DTPicker1.Value)
                datcierre2.Recordset.Fields("idtransferencia") = xnrocierre
                xcaja = 2
'                If datcierre.Recordset.Fields("tipodefacturacionid") = "NN" Then xcaja = 2
                datcierre2.Recordset.Fields("caja") = xcaja
                datcierre2.Recordset.UpdateBatch adAffectCurrent
            End If
            
            datcierre.Recordset.MoveNext
   Loop
                
                

End If

'******* Graba Cola importar
        
        datcierre2.RecordSource = "Select * from ud_ezi_cola_importar where id = 0"
        datcierre2.Refresh
        
        datcierre2.Recordset.AddNew
        datcierre2.Recordset.Fields("id_encabezado") = xnrocierre
        datcierre2.Recordset.Fields("tipodedocumentoid") = datparametros.Recordset.Fields("idtransferencia")
        datcierre2.Recordset.Fields("unidadoperativaid") = datparametros.Recordset.Fields("target")
        datcierre2.Recordset.Fields("fecha_hora") = DateValue(DTPicker1.Value) + TimeValue(Str(Time))
        
        datcierre2.Recordset.UpdateBatch adAffectCurrent

        mensa = MsgBox("Cierre de Caja Realizado", vbInformation, "Cierre de Caja")
        Call traervalores_Click
        Exit Sub
fuera:
     MsgBox "No se puede efecutar el cierre, seleccione los valores a transferir", vbCritical, "Error"
     Call traervalores_Click

End Sub

Private Sub salir_Click()

    Unload Me

End Sub


Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)

    If KeyAscii = 13 And Index = 5 Then
        KeyAscii = 0
         Text1(5).Text = Format(Text1(5).Text, "0.00")
         If Val(Text1(5).Text) > xefectivo Then
            msg = MsgBox("El importe no puede ser superior al disponible", vbCritical, "Error")
            Text1(5).Text = Format(xefectivo, "###,##0.00")
            Text1(0).Text = Format(xefectivo, "0.00")
         Else
            Text1(0).Text = Format(xefectivo - Val(Text1(5).Text), "###,##0.00")
            Text1(5).Text = Format(Text1(5).Text, "###,##0.00")
         End If
    
        xvalortotal = 0
        xvaloranterior = 0
        For X = 5 To 8
            xvalortotal = xvalortotal + Val(Format(Text1(X).Text, "0.00"))
            xvaloranterior = xvaloranterior + Val(Format(Text1(X - 5).Text, "0.00"))
        Next X
        Text1(4).Text = Format(xvaloranterior, "###,##0.00")
        Text1(9).Text = Format(xvalortotal, "###,##0.00")
    End If

        


End Sub

Private Sub traervalores_Click()

    fechacierre = DTPicker1.Value + 1
    fechacierre1 = Right("0" + Replace(Str(Day(fechacierre)), " ", ""), 2) + "/" + Right("0" + Replace(Str(Month(fechacierre)), " ", ""), 2) + "/" + Replace(Str(Year(fechacierre)), " ", "")
    xquery = "SELECT     ud_ezi_pago.formadepago, V_TARJETACREDITO_.NOMBRE AS Tarjeta, ud_ezi_pago.cuotas, ud_ezi_pago.numerodecupon AS NroCupon, " & _
              "ud_ezi_pago.numero AS NroCheque, ud_ezi_pago.fechadeemision, ud_ezi_pago.fechadevencimiento, SUM(ud_ezi_pago.monto) AS monto, ud_ezi_pago.transferido,  " & _
              "V_CAJA_.DESCRIPCION AS caja, V_TRFACTURAVENTA_.FLAG_ID  " & _
              "FROM         V_TRFACTURAVENTA_ RIGHT OUTER JOIN " & _
              "ud_ezi_puntodeventa_encabezado WITH (readpast) ON V_TRFACTURAVENTA_.ID = ud_ezi_puntodeventa_encabezado.calipsoid LEFT OUTER JOIN " & _
              "V_CAJA_ RIGHT OUTER JOIN " & _
              "ud_ezi_pago WITH (readpast) ON V_CAJA_.ID = ud_ezi_pago.destinoid ON ud_ezi_puntodeventa_encabezado.id = ud_ezi_pago.claveprimaria LEFT OUTER JOIN " & _
              "V_TARJETACREDITO_ ON ud_ezi_pago.tarjetaid = V_TARJETACREDITO_.ID " & _
              "WHERE     (ud_ezi_puntodeventa_encabezado.fechadelcomprobante < CONVERT(DATETIME, '" & fechacierre1 & "', 103)) AND ( ud_ezi_pago.sucursal = '" & login.nomsucursal & "') " & _
              "GROUP BY ud_ezi_pago.formadepago, V_TARJETACREDITO_.NOMBRE, ud_ezi_pago.cuotas, ud_ezi_pago.numerodecupon, ud_ezi_pago.numero, " & _
              "ud_ezi_pago.fechadeemision , ud_ezi_pago.fechadevencimiento, ud_ezi_pago.transferido, V_CAJA_.DESCRIPCION, V_TRFACTURAVENTA_.FLAG_ID " & _
              "HAVING (ud_ezi_pago.transferido IS NULL) AND (ud_ezi_pago.formadepago <> 'Debito en Cuenta Corriente') AND (ud_ezi_pago.formadepago <> 'Contra Reembolso') AND (V_TRFACTURAVENTA_.FLAG_ID IS NULL) " & _
              "ORDER BY ud_ezi_pago.formadepago, Tarjeta"
                               
                               
    datproducto.RecordSource = xquery
    datproducto.Refresh
    DataGrid1.Visible = True
    DataGrid1.Columns(0).Caption = "Tipo de Valor"
    DataGrid1.Columns(0).Width = 3000
    DataGrid1.Columns(1).Caption = "Tarjeta"
    DataGrid1.Columns(1).Width = 2000
    DataGrid1.Columns(2).Caption = "Cuotas"
    DataGrid1.Columns(2).Width = 1000
    DataGrid1.Columns(2).Alignment = dbgCenter
    DataGrid1.Columns(3).Caption = "N°Cupon"
    DataGrid1.Columns(3).Width = 1000
    DataGrid1.Columns(3).Alignment = dbgCenter
    DataGrid1.Columns(4).Caption = "N°Cheque"
    DataGrid1.Columns(4).Width = 1000
    DataGrid1.Columns(4).Alignment = dbgCenter
    DataGrid1.Columns(5).Caption = "Fec.Emisión"
    DataGrid1.Columns(5).Width = 1200
    DataGrid1.Columns(5).Alignment = dbgCenter
    DataGrid1.Columns(6).Caption = "Fec.Venc."
    DataGrid1.Columns(6).Width = 1200
    DataGrid1.Columns(6).Alignment = dbgCenter
    DataGrid1.Columns(7).Caption = "Importe"
    DataGrid1.Columns(7).Width = 1500
    DataGrid1.Columns(7).Alignment = dbgRight
    DataGrid1.Columns(7).NumberFormat = "Currency"
    DataGrid1.Columns(8).Visible = False
    DataGrid1.Columns(9).Width = 1000
    DataGrid1.Columns(9).Caption = "Caja"
    
    For X = 0 To 9
        Text1(X).Text = Format(0, "###,##0.00")
    Next X
    
    For X = 0 To 4
        transferir(X).Visible = True
        deshacer(X).Visible = False
    Next X
    
    xefectivo = 0
    xcheque = 0
    xtarjeta = 0
    xotros = 0
    If datproducto.Recordset.EOF = False Then
        datproducto.Recordset.MoveFirst
        Do While Not datproducto.Recordset.EOF
            xvalor = datproducto.Recordset.Fields("formadepago")
            If Left(xvalor, 6) = "Cheque" Then xcheque = xcheque + datproducto.Recordset.Fields("monto")
            If Left(xvalor, 6) = "Efecti" Then xefectivo = xefectivo + datproducto.Recordset.Fields("monto")
            If Left(xvalor, 6) = "Tarjet" Then xtarjeta = xtarjeta + datproducto.Recordset.Fields("monto")
            If Left(xvalor, 6) <> "Cheque" And Left(xvalor, 6) <> "Efecti" And Left(xvalor, 6) <> "Tarjet" Then
                xotros = xotros + datproducto.Recordset.Fields("monto")
            End If
            datproducto.Recordset.MoveNext
        Loop
    End If
    
    xtotal = xefectivo + xtarjeta + xcheque + xotros
    Text1(0).Text = Format(xefectivo, "###,##0.00")
    Text1(1).Text = Format(xtarjeta, "###,##0.00")
    Text1(2).Text = Format(xcheque, "###,##0.00")
    Text1(3).Text = Format(xotros, "###,##0.00")
    Text1(4).Text = Format(xtotal, "###,##0.00")
    
    
    
            
    
    
    
    
End Sub

Private Sub transferir_Click(Index As Integer)

If Index = 4 Then
    If Val(Format(Text1(0).Text, "0.00")) > 0 Then
        Text1(5).Text = Format(Text1(0).Text, "###,##0.00")
    End If
    If Val(Format(Text1(1).Text, "0.00")) > 0 Then
        Text1(6).Text = Format(Text1(1).Text, "###,##0.00")
    End If
    If Val(Format(Text1(2).Text, "0.00")) > 0 Then
        Text1(7).Text = Format(Text1(2).Text, "###,##0.00")
    End If
    If Val(Format(Text1(3).Text, "0.00")) > 0 Then
        Text1(8).Text = Format(Text1(3).Text, "###,##0.00")
    End If
    If Val(Format(Text1(4).Text, "0.00")) > 0 Then
        Text1(9).Text = Format(Text1(4).Text, "###,##0.00")
    End If
    
    For X = 0 To 4
        Text1(X).Text = Format(0, "###,##0.00")
        transferir(X).Visible = False
        deshacer(X).Visible = True
    Next X
Else
    Text1(Index + 5).Text = Format(Text1(Index).Text, "###,##0.00")
    transferir(Index).Visible = False
    deshacer(Index).Visible = True
    Text1(Index).Text = Format(0, "###,##0.00")
End If
    
    xvalortotal = 0
    xvaloranterior = 0
    For X = 5 To 8
        xvalortotal = xvalortotal + Val(Format(Text1(X).Text, "0.00"))
        xvaloranterior = xvaloranterior + Val(Format(Text1(X - 5).Text, "0.00"))
    Next X

    If xvaloranterior = 0 Then
        transferir(4).Visible = False
        deshacer(4).Visible = True
    End If
    
    Text1(4).Text = Format(xvaloranterior, "###,##0.00")
    Text1(9).Text = Format(xvalortotal, "###,##0.00")
        


End Sub
