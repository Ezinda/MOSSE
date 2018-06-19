VERSION 5.00
Object = "{F6125AB1-8AB1-11CE-A77F-08002B2F4E98}#2.0#0"; "MSRDC20.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{5734474E-78D3-4254-99B9-C35F31BDF509}#62.0#0"; "vbskpro2.ocx"
Object = "{D18BBD1F-82BB-4385-BED3-E9D31A3E361E}#1.0#0"; "KewlButtonz.ocx"
Begin VB.Form frmrecibos1 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Recibos de Cobro"
   ClientHeight    =   7530
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10905
   Icon            =   "frmrecibos1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7530
   ScaleWidth      =   10905
   Begin MSDataGridLib.DataGrid DataGrid4 
      Bindings        =   "frmrecibos1.frx":0442
      Height          =   1215
      Left            =   360
      TabIndex        =   6
      Top             =   4800
      Visible         =   0   'False
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   2143
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
   Begin MSMask.MaskEdBox saldototal 
      Height          =   255
      Left            =   5040
      TabIndex        =   17
      Top             =   6840
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   0
      Appearance      =   0
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "$#,##0.00;-$#,##0.00"
      PromptChar      =   "_"
   End
   Begin MSDataListLib.DataList DataList4 
      Bindings        =   "frmrecibos1.frx":0460
      Height          =   1815
      Left            =   5880
      TabIndex        =   39
      Top             =   1920
      Visible         =   0   'False
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   3201
      _Version        =   393216
      MatchEntry      =   -1  'True
      OLEDragMode     =   1
      OLEDropMode     =   1
      BackColor       =   12632256
      ListField       =   "codigo"
      BoundColumn     =   "Cod Contable"
   End
   Begin MSMask.MaskEdBox totalinstrumento 
      Height          =   255
      Left            =   8040
      TabIndex        =   16
      Top             =   6840
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   0
      Appearance      =   0
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "$#,##0.00;-$#,##0.00"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox totalabonan 
      Height          =   255
      Left            =   8040
      TabIndex        =   12
      Top             =   6480
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   0
      Appearance      =   0
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "$#,##0.00;-$#,##0.00"
      PromptChar      =   "_"
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
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
      Left            =   6960
      TabIndex        =   13
      Text            =   "A Cobrar:"
      Top             =   6480
      Width           =   1095
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
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
      Left            =   7200
      TabIndex        =   15
      Text            =   "Total:"
      Top             =   6840
      Width           =   615
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
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
      Left            =   4440
      TabIndex        =   14
      Text            =   "Saldo:"
      Top             =   6840
      Width           =   615
   End
   Begin MSDataGridLib.DataGrid DataGrid3 
      Bindings        =   "frmrecibos1.frx":0479
      Height          =   1215
      Left            =   5880
      TabIndex        =   5
      Top             =   5040
      Visible         =   0   'False
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   2143
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
   Begin TabDlg.SSTab SSTab1 
      Height          =   6375
      Left            =   120
      TabIndex        =   10
      Top             =   1080
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   11245
      _Version        =   393216
      Tab             =   2
      TabHeight       =   520
      ForeColor       =   128
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "F1-CONCEPTOS PAGOS A CLIENTES."
      TabPicture(0)   =   "frmrecibos1.frx":0497
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "veriorden"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Text7"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "DataList2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "DataList1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Text6(2)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Text6(1)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Text6(0)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "grillapago"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Frame5"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "importepago"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "saldocomp"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "cancelar"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "command7"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "nuevo"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "borrar"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label16"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Label12"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Label11"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Label10"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Label9"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Label8"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).ControlCount=   21
      TabCaption(1)   =   "F2-INSTRUMENTO DE COBRO"
      TabPicture(1)   =   "frmrecibos1.frx":04B3
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "DataList3"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Text5(3)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Text5(2)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Text5(1)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "importemask"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Text5(0)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "grillainstru"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "MaskEdBox2"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "MaskEdBox3"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Frame4"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "Command1"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "Command6"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "command2"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "Label7"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "Label6"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "Label5"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "Label4"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "Label3"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "Label2"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "Label1"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).ControlCount=   20
      TabCaption(2)   =   "F3-OTROS CONCEPTOS"
      TabPicture(2)   =   "frmrecibos1.frx":04CF
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Label13"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Label14"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Label15"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "eliminaotroconcepto"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "nuevootroconcepto"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "command8"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "Frame6"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "importeotro"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "grillaotros"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "Text6(4)"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).Control(10)=   "Text6(5)"
      Tab(2).Control(10).Enabled=   0   'False
      Tab(2).ControlCount=   11
      Begin VB.CommandButton veriorden 
         Caption         =   "veriorden"
         Height          =   255
         Left            =   -74400
         TabIndex        =   74
         Top             =   6120
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox Text7 
         Height          =   285
         Left            =   -65874
         TabIndex        =   60
         Top             =   480
         Width           =   975
      End
      Begin MSDataListLib.DataList DataList2 
         Bindings        =   "frmrecibos1.frx":04EB
         Height          =   1815
         Left            =   -68640
         TabIndex        =   54
         Top             =   720
         Visible         =   0   'False
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   3201
         _Version        =   393216
         MatchEntry      =   -1  'True
         BackColor       =   12632256
         ListField       =   "comp"
         BoundColumn     =   "id"
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Index           =   5
         Left            =   9240
         TabIndex        =   57
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Index           =   4
         Left            =   1920
         TabIndex        =   51
         Top             =   480
         Width           =   3615
      End
      Begin MSDataListLib.DataList DataList1 
         Bindings        =   "frmrecibos1.frx":0509
         Height          =   1815
         Left            =   -73560
         TabIndex        =   49
         Top             =   840
         Visible         =   0   'False
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   3201
         _Version        =   393216
         MatchEntry      =   -1  'True
         BackColor       =   12632256
         ListField       =   "razonsocial"
         BoundColumn     =   "codcliente"
      End
      Begin VB.TextBox Text6 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   2
         Left            =   -73560
         Locked          =   -1  'True
         TabIndex        =   44
         Top             =   840
         Width           =   1455
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Index           =   1
         Left            =   -68640
         TabIndex        =   42
         Top             =   480
         Width           =   1815
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Index           =   0
         Left            =   -73560
         TabIndex        =   40
         Top             =   480
         Width           =   3495
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grillapago 
         Height          =   3855
         Left            =   -74880
         TabIndex        =   37
         Top             =   1320
         Width           =   10455
         _ExtentX        =   18441
         _ExtentY        =   6800
         _Version        =   393216
         Rows            =   100
         Cols            =   5
         FixedCols       =   0
         SelectionMode   =   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   5
      End
      Begin MSDataListLib.DataList DataList3 
         Bindings        =   "frmrecibos1.frx":0526
         Height          =   1815
         Left            =   -73320
         TabIndex        =   36
         Top             =   840
         Visible         =   0   'False
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   3201
         _Version        =   393216
         MatchEntry      =   -1  'True
         BackColor       =   12632256
         ListField       =   "instrumento"
         BoundColumn     =   "codcontable"
      End
      Begin VB.Frame Frame5 
         Height          =   975
         Left            =   -71040
         TabIndex        =   35
         Top             =   5160
         Width           =   6375
      End
      Begin VB.TextBox Text5 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   3
         Left            =   -66108
         MaxLength       =   13
         TabIndex        =   22
         Top             =   480
         Width           =   1222
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Index           =   2
         Left            =   -73320
         TabIndex        =   23
         Top             =   840
         Width           =   2055
      End
      Begin VB.TextBox Text5 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   1
         Left            =   -67800
         TabIndex        =   21
         Top             =   480
         Width           =   735
      End
      Begin MSMask.MaskEdBox importemask 
         Height          =   285
         Left            =   -70440
         TabIndex        =   20
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         Format          =   "$ #,###,##0.00"
         PromptChar      =   "_"
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Index           =   0
         Left            =   -73320
         TabIndex        =   19
         Top             =   480
         Width           =   2055
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grillainstru 
         Height          =   3855
         Left            =   -74880
         TabIndex        =   18
         Top             =   1320
         Width           =   10455
         _ExtentX        =   18441
         _ExtentY        =   6800
         _Version        =   393216
         Rows            =   15
         Cols            =   7
         FixedCols       =   0
         SelectionMode   =   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   7
      End
      Begin MSMask.MaskEdBox MaskEdBox2 
         Height          =   285
         Left            =   -69840
         TabIndex        =   24
         Top             =   840
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox3 
         Height          =   285
         Left            =   -67200
         TabIndex        =   25
         Top             =   840
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Frame Frame4 
         Height          =   975
         Left            =   -71040
         TabIndex        =   34
         Top             =   5160
         Width           =   6375
      End
      Begin MSMask.MaskEdBox importepago 
         Height          =   285
         Left            =   -70560
         TabIndex        =   46
         Top             =   840
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         Format          =   "$ #,###,##0.00"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox saldocomp 
         Height          =   285
         Left            =   -67320
         TabIndex        =   47
         Top             =   840
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   "$ #,###,##0.00"
         PromptChar      =   "_"
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grillaotros 
         Height          =   4335
         Left            =   120
         TabIndex        =   50
         Top             =   840
         Width           =   10455
         _ExtentX        =   18441
         _ExtentY        =   7646
         _Version        =   393216
         Rows            =   100
         Cols            =   3
         FixedCols       =   0
         SelectionMode   =   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   3
      End
      Begin MSMask.MaskEdBox importeotro 
         Height          =   285
         Left            =   6480
         TabIndex        =   53
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         Format          =   "$ #,###,##0.00"
         PromptChar      =   "_"
      End
      Begin VB.Frame Frame6 
         Height          =   975
         Left            =   3960
         TabIndex        =   58
         Top             =   5160
         Width           =   6375
      End
      Begin VB.CommandButton cancelar 
         Cancel          =   -1  'True
         Caption         =   "cancelar"
         Height          =   495
         Left            =   -71040
         TabIndex        =   11
         Top             =   5160
         Visible         =   0   'False
         Width           =   1215
      End
      Begin KewlButtonz.KewlButtons command7 
         Height          =   735
         Left            =   -74760
         TabIndex        =   63
         Top             =   5400
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   1296
         BTYPE           =   14
         TX              =   "Gra&bar"
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
         MICON           =   "frmrecibos1.frx":053E
         PICN            =   "frmrecibos1.frx":055A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin KewlButtonz.KewlButtons nuevo 
         Height          =   735
         Left            =   -73680
         TabIndex        =   64
         Top             =   5400
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   1296
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
         MICON           =   "frmrecibos1.frx":0F6C
         PICN            =   "frmrecibos1.frx":0F88
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
         Height          =   735
         Left            =   -72480
         TabIndex        =   65
         Top             =   5400
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   1296
         BTYPE           =   14
         TX              =   "&Elimin.Conc."
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
         MICON           =   "frmrecibos1.frx":199A
         PICN            =   "frmrecibos1.frx":19B6
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin KewlButtonz.KewlButtons Command1 
         Height          =   735
         Left            =   -74760
         TabIndex        =   26
         Top             =   5400
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   1296
         BTYPE           =   14
         TX              =   "Gra&bar"
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
         MICON           =   "frmrecibos1.frx":23C8
         PICN            =   "frmrecibos1.frx":23E4
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin KewlButtonz.KewlButtons Command6 
         Height          =   735
         Left            =   -73680
         TabIndex        =   66
         Top             =   5400
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   1296
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
         MICON           =   "frmrecibos1.frx":2DF6
         PICN            =   "frmrecibos1.frx":2E12
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin KewlButtonz.KewlButtons command2 
         Height          =   735
         Left            =   -72480
         TabIndex        =   67
         Top             =   5400
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   1296
         BTYPE           =   14
         TX              =   "&Elimin.Intru"
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
         MICON           =   "frmrecibos1.frx":3824
         PICN            =   "frmrecibos1.frx":3840
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin KewlButtonz.KewlButtons command8 
         Height          =   735
         Left            =   240
         TabIndex        =   68
         Top             =   5400
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   1296
         BTYPE           =   14
         TX              =   "Gra&bar"
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
         MICON           =   "frmrecibos1.frx":4252
         PICN            =   "frmrecibos1.frx":426E
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin KewlButtonz.KewlButtons nuevootroconcepto 
         Height          =   735
         Left            =   1320
         TabIndex        =   69
         Top             =   5400
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   1296
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
         MICON           =   "frmrecibos1.frx":4C80
         PICN            =   "frmrecibos1.frx":4C9C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin KewlButtonz.KewlButtons eliminaotroconcepto 
         Height          =   735
         Left            =   2520
         TabIndex        =   70
         Top             =   5400
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   1296
         BTYPE           =   14
         TX              =   "&Elimin.Conc."
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
         MICON           =   "frmrecibos1.frx":56AE
         PICN            =   "frmrecibos1.frx":56CA
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label16 
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
         Height          =   260
         Left            =   -66693
         TabIndex        =   61
         Top             =   481
         Width           =   975
      End
      Begin VB.Label Label15 
         Caption         =   "Cod.Contable:"
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
         Left            =   7920
         TabIndex        =   56
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label14 
         Caption         =   "Importe:"
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
         Left            =   5760
         TabIndex        =   55
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label13 
         Caption         =   "Detalle Concepto:"
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
         TabIndex        =   52
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label Label12 
         Caption         =   "Saldo Comprobante:"
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
         Left            =   -69120
         TabIndex        =   48
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label Label11 
         Caption         =   "Imp. a Cobrar:"
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
         Left            =   -71760
         TabIndex        =   45
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label10 
         Caption         =   "Fecha Comp.:"
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
         Left            =   -74880
         TabIndex        =   43
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label9 
         Caption         =   "Nº Comprob.:"
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
         Left            =   -69840
         TabIndex        =   41
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "Cliente:"
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
         Left            =   -74640
         TabIndex        =   38
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "NºComp.:"
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
         Left            =   -66960
         TabIndex        =   33
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label6 
         Caption         =   "Fecha Venc.:"
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
         Left            =   -68400
         TabIndex        =   32
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "Fecha Emisión:"
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
         Left            =   -71160
         TabIndex        =   31
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "Denominacion:"
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
         Left            =   -74760
         TabIndex        =   30
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Cod.Contable:"
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
         Left            =   -69000
         TabIndex        =   29
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Importe:"
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
         Left            =   -71160
         TabIndex        =   28
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Forma de Cobro:"
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
         Left            =   -74760
         TabIndex        =   27
         Top             =   480
         Width           =   1455
      End
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Original"
      Height          =   195
      Index           =   0
      Left            =   5400
      TabIndex        =   8
      Top             =   600
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   255
      Left            =   6960
      TabIndex        =   7
      Top             =   0
      Visible         =   0   'False
      Width           =   1095
   End
   Begin MSMask.MaskEdBox text1 
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   480
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "########"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   375
      Left            =   3240
      TabIndex        =   1
      Top             =   480
      Width           =   1695
      _ExtentX        =   2990
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
      Format          =   "dd/mm/yyyy"
      PromptChar      =   "_"
   End
   Begin MSAdodcLib.Adodc datordendepago 
      Height          =   330
      Left            =   480
      Top             =   6480
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
   Begin VB.Frame Frame1 
      Caption         =   "Nº Recibo"
      Height          =   1095
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   2895
      Begin VB.TextBox Text8 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   1440
         TabIndex        =   0
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label17 
         Caption         =   "Recibo Manual"
         Height          =   255
         Left            =   1440
         TabIndex        =   62
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Fecha"
      Height          =   1095
      Left            =   3120
      TabIndex        =   4
      Top             =   0
      Width           =   2055
      Begin VB.CommandButton compfecha 
         Caption         =   "compfecha"
         Height          =   375
         Left            =   600
         TabIndex        =   59
         Top             =   0
         Visible         =   0   'False
         Width           =   735
      End
   End
   Begin MSAdodcLib.Adodc databonan 
      Height          =   330
      Left            =   2400
      Top             =   6720
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
   Begin MSAdodcLib.Adodc datinstrumento 
      Height          =   330
      Left            =   3600
      Top             =   6720
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
   Begin MSAdodcLib.Adodc datproveedores 
      Height          =   330
      Left            =   4680
      Top             =   6720
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
   Begin MSAdodcLib.Adodc datconsultacomp 
      Height          =   330
      Left            =   6000
      Top             =   6720
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
   Begin MSAdodcLib.Adodc datlibrocompras 
      Height          =   330
      Left            =   7200
      Top             =   6720
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
   Begin MSAdodcLib.Adodc datinstru 
      Height          =   330
      Left            =   8400
      Top             =   6720
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
   Begin MSAdodcLib.Adodc datcuentas 
      Height          =   330
      Left            =   9360
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
   Begin MSAdodcLib.Adodc datmaestro 
      Height          =   330
      Left            =   480
      Top             =   6840
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
   Begin MSAdodcLib.Adodc datasiento 
      Height          =   330
      Left            =   1680
      Top             =   6840
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
   Begin Crystal.CrystalReport CrystalReporte 
      Left            =   4920
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTitle     =   "Orden de Pago"
      PrintFileLinesPerPage=   60
   End
   Begin MSRDC.MSRDC reporte 
      Height          =   375
      Left            =   4560
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
   Begin VB.Frame Frame3 
      Caption         =   "Imprimir Solamente"
      Height          =   1095
      Left            =   5280
      TabIndex        =   9
      Top             =   0
      Width           =   1575
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
      LcK1            =   "3.66*/4/0*/1-5*210/."
      LcK2            =   $"frmrecibos1.frx":60DC
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
   Begin MSAdodcLib.Adodc criterio 
      Height          =   330
      Left            =   -240
      Top             =   600
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
   Begin MSAdodcLib.Adodc datasigna 
      Height          =   330
      Left            =   0
      Top             =   0
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
   Begin MSAdodcLib.Adodc datinstrumento1 
      Height          =   330
      Left            =   0
      Top             =   0
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
   Begin KewlButtonz.KewlButtons nuevaorden 
      Height          =   735
      Left            =   8280
      TabIndex        =   71
      Top             =   240
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1296
      BTYPE           =   14
      TX              =   "Nueva &Orden"
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
      MICON           =   "frmrecibos1.frx":60EB
      PICN            =   "frmrecibos1.frx":6107
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin KewlButtonz.KewlButtons command3 
      Height          =   735
      Left            =   6960
      TabIndex        =   72
      Top             =   240
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1296
      BTYPE           =   14
      TX              =   "&Grabar Recibo"
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
      MICON           =   "frmrecibos1.frx":15560
      PICN            =   "frmrecibos1.frx":1557C
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
      Height          =   735
      Left            =   9600
      TabIndex        =   73
      Top             =   240
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1296
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
      MICON           =   "frmrecibos1.frx":16FFE
      PICN            =   "frmrecibos1.frx":1701A
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
Attribute VB_Name = "frmrecibos1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim importeapagar As Double
Dim totalab As Currency
Dim totalinst(50) As Currency
Dim detalleint(50) As String
Dim totalconc(50) As Currency
Dim nrocompro(50) As String
Dim cuentaint(50) As Integer
Dim cuentaotros(50) As Integer
Dim conceptosotros(50) As String
Dim nomprov(50) As String
Dim saldoactual As Currency
Dim Cuenta(150) As Integer
Dim codprove As Long
Dim idlibrogrid(50) As Integer
Dim saldolibro(50) As Currency
Dim sincomp As Integer
Dim codigopago As Integer
Dim codigopago1 As Integer
Dim empresareal As Integer
Public numorden As String
Public ordeninstu As String
Public importeord As Currency
Dim ordeninstruver(50) As String
Public ordenitem As Integer
Dim idcliente(500) As Integer

Private Sub borrablancos_Click()
On Error GoTo fuera

  databonan.RecordSource = "select recibocobroabonan.* from recibocobroabonan WHERE recibocobroabonan.empresa = " & login.empresaact & " and inicioper = '" & login.iper & "' and nrorden = '" & text1.Text & "' and '" & IsNull(codcliente) & "' ='" & True & "'  Order by fechacompro"
  databonan.Refresh
  If databonan.Recordset.EOF = False Then databonan.Recordset.Delete adAffectCurrent
  databonan.RecordSource = "select recibocobroabonan.* from recibocobroabonan WHERE recibocobroabonan.empresa = " & login.empresaact & " and inicioper = '" & login.iper & "' and nrorden = '" & text1.Text & "' Order by fechacompro"
  databonan.Refresh
  
  datinstrumento.RecordSource = "select recibocobroinstrumento.* from recibocobroinstrumento WHERE recibocobroinstrumento.empresa = " & login.empresaact & " and inicioper = '" & login.iper & "'  and nrorden = '" & text1.Text & "' and '" & IsNull(codcliente) & "' ='" & True & "' Order by id"
  datinstrumento.Refresh
  If datinstrumento.Recordset.EOF = False Then datinstrumento.Recordset.Delete adAffectCurrent
  datinstrumento.RecordSource = "select recibocobroinstrumento.* from recibocobroinstrumento WHERE recibocobroinstrumento.empresa = " & login.empresaact & " and inicioper = '" & login.iper & "'  and nrorden = '" & text1.Text & "' Order by id"
  datinstrumento.Refresh
  
fuera:
End Sub

Private Sub borrar_Click()
On Error GoTo erroreliminar1


    For x = 0 To 4
        grillapago.Col = x
        grillapago.Text = ""
    Next x
    
    For x = 0 To 2
        Text6(x).Text = ""
    Next x
    importemask.Text = ""
    saldocomp.Text = ""
    
    codigopago1 = 0
    
    totalabonan.Text = 0
For x = 1 To 99
    grillapago.Col = 3
    grillapago.Row = x
    If grillapago.Text <> "" Then totalabonan.Text = Val(totalabonan.Text) + grillapago.Text
Next x
    saldototal.Text = totalabonan.Text - Val(totalinstrumento.Text)
    nuevo.SetFocus

    

erroreliminar1:
End Sub

Private Sub Command1_Click()
On Error GoTo fuera
    linea = grillainstru.Row

    For x = 1 To 14
        grillainstru.Row = x
        grillainstru.Col = 0
        If grillainstru.Text = "" Then GoTo sigue
    Next x
    
sigue:
    If codigopago = 0 Then
        grillainstru.Row = x
    Else
        grillainstru.Row = linea
    End If
    

If ordeninstu <> "" Then
    For Y = grillainstru.Row - 1 To 1 Step -1
        If ordeninstu = ordeninstruver(Y) Then
            mensa = MsgBox("Recibo ya ingresada como instrumento de cobro", vbCritical, "Error")
            Call Command6_Click
            Exit Sub
        End If
    Next Y
End If

    grillainstru.Text = Text5(0).Text
    grillainstru.Col = 1
    grillainstru.Text = importemask.Text
    grillainstru.Text = Format(grillainstru.Text, "#,###,##0.00")
    grillainstru.Col = 2
    grillainstru.Text = Text5(1).Text
    grillainstru.Col = 4
    grillainstru.Text = Text5(3).Text
    grillainstru.Col = 3
    grillainstru.Text = Text5(2).Text
    grillainstru.Col = 5
    grillainstru.Text = MaskEdBox2.Text
    grillainstru.Col = 6
    grillainstru.Text = MaskEdBox3.Text
    
    ordeninstruver(grillainstru.Row) = ordeninstu
    ordeninstu = ""
 
For x = 0 To 3
    Text5(x).Text = ""
Next x
    importemask.Text = ""
    MaskEdBox2.Mask = ""
    MaskEdBox3.Mask = ""
    MaskEdBox2.Text = ""
    MaskEdBox3.Text = ""
    MaskEdBox2.Mask = "##/##/####"
    MaskEdBox3.Mask = "##/##/####"

    
    codigopago = 0
    
    totalinstrumento.Text = 0
For x = 1 To 14
    grillainstru.Col = 1
    grillainstru.Row = x
    If grillainstru.Text <> "" Then totalinstrumento.Text = Val(totalinstrumento.Text) + grillainstru.Text
Next x
    saldototal.Text = totalabonan.Text - Val(totalinstrumento.Text)
    Command6.SetFocus

GoTo fuera
    datinstrumento.Recordset.AddNew
    datinstrumento.Recordset.Fields(0) = text1.Text
    datinstrumento.Recordset.Fields(1) = login.empresaact
    datinstrumento.Recordset.Fields(2) = login.iper
    datinstrumento.Recordset.Fields(3) = login.fper
    DataGrid2.SetFocus
    DataGrid2.Col = 6

fuera:
End Sub

Private Sub Command1_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo fuera
    
    If KeyCode = 112 Then
        SSTab1.Tab = 0
        Call borrablancos_Click
        nuevo.SetFocus
    End If
    
fuera:
End Sub

Private Sub Command2_Click()
On Error GoTo erroreliminar1


    For x = 0 To 6
        grillainstru.Col = x
        grillainstru.Text = ""
    Next x
    
    For x = 0 To 3
        Text5(x).Text = ""
    Next x
    importemask.Text = ""
    MaskEdBox2.Mask = ""
    MaskEdBox3.Mask = ""
    MaskEdBox2.Text = ""
    MaskEdBox3.Text = ""
    MaskEdBox2.Mask = "##/##/####"
    MaskEdBox3.Mask = "##/##/####"

    
    codigopago = 0
    
    totalinstrumento.Text = 0
For x = 1 To 14
    grillainstru.Col = 1
    grillainstru.Row = x
    If grillainstru.Text <> "" Then totalinstrumento.Text = Val(totalinstrumento.Text) + grillainstru.Text
Next x
    saldototal.Text = totalabonan.Text - Val(totalinstrumento.Text)
    Command6.SetFocus

    

erroreliminar1:
End Sub

Private Sub Command3_Click()
On Error GoTo fuera
Dim campodebe As Currency
Dim campohaber As Currency


    

    If totalinstrumento = "" Then
        mensa = MsgBox("Cobro no Imputado", vbExclamation, "Atencion")
        Exit Sub
    End If
    diferencia = totalabonan - totalinstrumento
    If diferencia < 0 Then diferencia = diferencia * -1
    If diferencia > 0.009 Then
        mensa = MsgBox("El Asiento esta Desvalanceado, no se puede grabar", vbCritical, "!! Error !!")
        Exit Sub
    End If
    
  
    Call veriorden_Click
    numorden = text1.Text
    
    datordendepago.Recordset.Fields("nrorden") = text1.Text
    datordendepago.Recordset.Fields("empresa") = login.empresaact
    datordendepago.Recordset.Fields("inicioper") = login.iper
    datordendepago.Recordset.Fields("finper") = login.fper
    datordendepago.Recordset.Fields("fecha") = MaskEdBox1.Text
    datordendepago.Recordset.Fields("correcta") = "S"
    datordendepago.Recordset.Fields("anulado") = "N"
    datordendepago.Recordset.Fields("recibomanual") = Text8.Text
    
    
Rem ****************** grabar asiento

    campoaño = Right(MaskEdBox1.Text, 4)
    campomes = Mid(MaskEdBox1.Text, 4, 2)
    campodia = Left(MaskEdBox1.Text, 2)
    campofecha = campoaño + "/" + campomes + "/" + campodia
       inicioperiodo = login.iper
    campoaño1 = Right(inicioperiodo, 4)
    campomes1 = Mid(inicioperiodo, 4, 2)
    campodia1 = Left(inicioperiodo, 2)
    campofecha1 = campoaño1 + "/" + campomes1 + "/" + campodia1
       finperiodo = login.fper
    campoaño2 = Right(finperiodo, 4)
    campomes2 = Mid(finperiodo, 4, 2)
    campodia2 = Left(finperiodo, 2)
    campofecha2 = campoaño2 + "/" + campomes2 + "/" + campodia2

    If campofecha < campofecha1 Or campofecha > campofecha2 Then
            mensa = MsgBox("La Fecha es erronea o no pertenecia al periodo en ejercicio", vbCritical, "!! Atención !!")
            MaskEdBox1.SelLength = 10
            MaskEdBox1.SetFocus
            Exit Sub
    End If

    If datmaestro.Recordset.EOF = False Then
            datmaestro.RecordSource = "SELECT  MAX(nroasiento) AS nroasiento, perinicial, empresa From dbo.[Maestro Asientos] GROUP BY perinicial, empresa HAVING (perinicial = '" & login.iper & "' and empresa = " & login.empresaact & " )"
            datmaestro.Refresh
            If datmaestro.Recordset.EOF = True Then
                nroasie = 1
                GoTo pas1
            End If
            Rem datmaestro.RecordSource = "select [Maestro Asientos].* from [Maestro Asientos] where empresa = " & login.empresaact & " and perinicial = '" & login.iper & "' order by nroasiento"
            Rem datmaestro.Refresh
            Rem datmaestro.Recordset.MoveLast
            nroasie = datmaestro.Recordset.Fields("nroasiento") + 1
    Else
            nroasie = 1
    End If
         
pas1:
    If codprove <> 0 Then
        grillapago.Col = 0
        grillapago.Row = 1
        Text6(0).Text = grillapago.Text
    End If

Rem    datmaestro.Recordset.AddNew
Rem    datmaestro.Recordset.Fields(0) = MaskEdBox1.Text
Rem    datmaestro.Recordset.Fields(1) = Date
Rem    datmaestro.Recordset.Fields(3) = nroasie
Rem    datmaestro.Recordset.Fields(4) = "Recibo Cobro " + text1.Text + " a " + Left(Text6(0).Text, 20)
Rem    datmaestro.Recordset.Fields(5) = login.iper
Rem    datmaestro.Recordset.Fields(6) = login.fper
Rem    datmaestro.Recordset.Fields(7) = login.empresaact
Rem    datmaestro.Recordset.Fields(8) = "N"
Rem    datmaestro.Recordset.Fields(10) = "R"
Rem    datmaestro.Recordset.Fields(11) = "S"
Rem    datmaestro.Recordset.UpdateBatch adAffectCurrent
    datordendepago.Recordset.Fields("idasiento") = nroasie
    datordendepago.Recordset.UpdateBatch adAffectCurrent
    
    concepto = "Recibo Cobro " + text1.Text + " a " + Left(Text6(0).Text, 20)
   
    Data1.Connection1.Open
    Data1.Connection1.Execute ("INSERT INTO dbo.[Maestro Asientos]" _
        & "(fecha, fecharegistro, nroasiento, concepto, perinicial, perfinal, empresa, cerrado, libro, balanceado)" _
        & " VALUES ('" & MaskEdBox1.Text & "' , '" & Date & "' , " & nroasie & " , '" & concepto & "' , '" & login.iper & "' , '" & login.fper & "'" _
        & " , " & login.empresaact & " , 'N' , 'R' , 'S' )")
    Data1.Connection1.Close
    
    datmaestro.RecordSource = "select [Maestro Asientos].* from [Maestro Asientos] where empresa = " & login.empresaact & " and perinicial = '" & login.iper & "' and nroasiento = " & nroasie & ""
    datmaestro.Refresh
    
Data1.Connection1.Open
    campofecha = datmaestro.Recordset.Fields("fecha").Value
    campomaster = Val(datmaestro.Recordset.Fields("idmasterasientos"))

    If codprove = 0 Then
        For x = 1 To 99
            grillaotros.Col = 0
            grillaotros.Row = x
            If grillaotros.Text = "" Then GoTo sigueotros
            grillaotros.Col = 2
            campocuenta = grillaotros.Text
            campodebe = 0
            grillaotros.Col = 1
            campohaber = grillaotros.Text
            grillaotros.Col = 0
            campodetalle = grillaotros.Text
            Data1.Connection1.Execute ("INSERT INTO dbo.[Detalle Asientos]" _
                & "(fecha, idcuenta, debe, haber, idmasterasientos, detallefila, empresa)" _
                & " VALUES ('" & campofecha & "' , " & campocuenta & " , " & campodebe & " , " & campohaber & " , " & campomaster & " , '" & campodetalle & "'" _
                & " , " & login.empresaact & " )")

sigueotros:
            
        Next x
        GoTo conti2
    End If


For j = 1 To 99
    grillapago.Col = 0
    grillapago.Row = j
    If grillapago.Text = "" Then GoTo siguepago

    grillapago.Col = 3
    If grillapago.Text > 0 Then
      campodebe = 0
      campohaber = grillapago.Text
    Else
      campodebe = -grillapago.Text
      campohaber = 0
    End If
    
    grillapago.Col = 1
    campodetalle = "Comp.:" + grillapago.Text
        
    Data1.Connection1.Execute ("INSERT INTO dbo.[Detalle Asientos]" _
        & "(fecha, idcuenta, debe, haber, idmasterasientos, detallefila, empresa)" _
        & " VALUES ('" & campofecha & "' , " & Cuenta(j) & " , " & campodebe & " , " & campohaber & " , " & campomaster & " , '" & campodetalle & "'" _
        & " , " & login.empresaact & ")")
        
siguepago:
Next j

conti2:

    For x = 1 To 14
            grillainstru.Col = 0
            grillainstru.Row = x
            If grillainstru.Text = "" Then GoTo finne1
            grillainstru.Col = 2
            campocuenta = grillainstru.Text
            grillainstru.Col = 1
            campodebe = grillainstru.Text
            campohaber = 0
            grillainstru.Col = 0
            camp1 = grillainstru.Text
            grillainstru.Col = 4
            camp2 = grillainstru.Text
            campodetalle = camp1 + " " + camp2
            
        
        Data1.Connection1.Execute ("INSERT INTO dbo.[Detalle Asientos]" _
            & "(fecha, idcuenta, debe, haber, idmasterasientos, detallefila, empresa)" _
            & " VALUES ('" & campofecha & "' , " & campocuenta & " , " & campodebe & " , " & campohaber & " , " & campomaster & " , '" & campodetalle & "'" _
            & " , " & login.empresaact & " )")
            
            
finne1:
    Next x
Data1.Connection1.Close
    
Rem ************* graba pago  ********************
If codprove = 0 Then
For x = 1 To 99
     grillaotros.Row = x
     grillaotros.Col = 0
     If grillaotros.Text = "" Then GoTo finne3
       
    databonan.Recordset.AddNew
    databonan.Recordset.Fields("nrorden") = text1.Text
    databonan.Recordset.Fields("empresa") = login.empresaact
    databonan.Recordset.Fields("inicioper") = login.iper
    databonan.Recordset.Fields("finper") = idcliente(x)
    databonan.Recordset.Fields("codcliente") = codprove
    grillaotros.Col = 0
    databonan.Recordset.Fields("nomcliente") = grillaotros.Text
    grillaotros.Col = 1
    databonan.Recordset.Fields("importe") = grillaotros.Text
    grillaotros.Col = 2
    databonan.Recordset.Fields("codcuenta") = grillaotros.Text
    
    databonan.Recordset.UpdateBatch adAffectCurrent
finne3:
Next x

Else

For x = 1 To 99
     grillapago.Col = 0
     grillapago.Row = x
     If grillapago.Text = "" Then GoTo finne2
       
    databonan.Recordset.AddNew
    databonan.Recordset.Fields("nrorden") = text1.Text
    databonan.Recordset.Fields("empresa") = login.empresaact
    databonan.Recordset.Fields("inicioper") = login.iper
    databonan.Recordset.Fields("finper") = login.fper
    databonan.Recordset.Fields("codcliente") = codprove
    databonan.Recordset.Fields("nomcliente") = grillapago.Text
    grillapago.Col = 1
    If grillapago.Text <> "" Then databonan.Recordset.Fields("comprobante") = grillapago.Text
    grillapago.Col = 2
    If grillapago.Text <> "" Then databonan.Recordset.Fields("fechacompro") = grillapago.Text
    grillapago.Col = 3
    databonan.Recordset.Fields("importe") = grillapago.Text
    grillapago.Col = 4
    If grillapago.Text <> "" Then
        databonan.Recordset.Fields("saldofactura") = grillapago.Text
    Else
        databonan.Recordset.Fields("saldofactura") = 0
    End If
    databonan.Recordset.Fields("codcuenta") = Cuenta(x)
    
    databonan.Recordset.UpdateBatch adAffectCurrent
finne2:
Next x
    
End If

Rem ************* graba instrumento *******************
For x = 1 To 14
     grillainstru.Col = 0
     grillainstru.Row = x
     If grillainstru.Text = "" Then GoTo finne4
       
    datinstrumento.Recordset.AddNew
    datinstrumento.Recordset.Fields("nrorden") = text1.Text
    datinstrumento.Recordset.Fields("empresa") = login.empresaact
    datinstrumento.Recordset.Fields("inicioper") = login.iper
    datinstrumento.Recordset.Fields("finper") = login.fper
    grillainstru.Col = 0
    datinstrumento.Recordset.Fields("instrumento") = grillainstru.Text
    ordenes = grillainstru.Text
    grillainstru.Col = 1
    datinstrumento.Recordset.Fields("importe") = grillainstru.Text
    
    If ordeninstruver(x) <> "" Then
        ordennro = ordeninstruver(x)
        datinstrumento1.RecordSource = "select recibocobroabonan.* from recibocobroabonan WHERE recibocobroabonan.empresa = " & login.empresaact & " and nrorden = '" & ordennro & "'"
        datinstrumento1.Refresh
        importefinal = datinstrumento1.Recordset.Fields("importe") - grillainstru.Text

        datasigna.RecordSource = "select recibocobroasignacion.* from recibocobroasignacion "
        datasigna.Refresh
        datasigna.Recordset.AddNew
        datasigna.Recordset.Fields("orden") = ordennro
        datasigna.Recordset.Fields("empresa") = login.empresaact
        datasigna.Recordset.Fields("importeoriginal") = datinstrumento1.Recordset.Fields("importe")
        datasigna.Recordset.Fields("nuevaorden") = text1.Text
        datasigna.Recordset.Fields("importe") = grillainstru.Text
        datasigna.Recordset.Fields("fecha") = Date
        datasigna.Recordset.UpdateBatch adAffectCurrent
        datinstrumento1.Recordset.Fields("importe") = importefinal
        datinstrumento1.Recordset.UpdateBatch adAffectCurrent
    End If
       

    grillainstru.Col = 2
    datinstrumento.Recordset.Fields("codcuenta") = grillainstru.Text
    grillainstru.Col = 3
    datinstrumento.Recordset.Fields("denominacion") = grillainstru.Text
    grillainstru.Col = 4
    If grillainstru <> "" Then datinstrumento.Recordset.Fields("comprobante") = grillainstru.Text
    grillainstru.Col = 5
    If grillainstru.Text <> "__/__/____" Then datinstrumento.Recordset.Fields("fechacompro") = grillainstru.Text
    grillainstru.Col = 6
    If grillainstru.Text <> "__/__/____" Then datinstrumento.Recordset.Fields("fechavencim") = grillainstru.Text
    datinstrumento.Recordset.UpdateBatch adAffectCurrent
    
    
        
    
    
finne4:
Next x

Rem ************* impacta en comprobante ***************

For x = 1 To 99
       grillapago.Col = 1
       grillapago.Row = x
        
       If grillapago.Text = "" Then GoTo finne5
       datlibrocompras.RecordSource = "select libroventas.* from libroventas WHERE libroventas.empresa = " & login.empresaact & " and id = " & idlibrogrid(x) & " Order by id"
       datlibrocompras.Refresh
       If datlibrocompras.Recordset.EOF = True Then GoTo finne5
       grillapago.Col = 4
       datlibrocompras.Recordset.Fields("saldo") = grillapago.Text
       datlibrocompras.Recordset.Fields("imputado") = "S"
       datlibrocompras.Recordset.UpdateBatch adAffectCurrent
finne5:
Next x






    Inicio.datauditoria.ConnectionString = login.conexiontotal
    
    Inicio.datauditoria.RecordSource = "select auditoria.* from auditoria"
    Inicio.datauditoria.Refresh
    
    Inicio.datauditoria.Recordset.AddNew
    Inicio.datauditoria.Recordset.Fields("fecha") = Date
    Inicio.datauditoria.Recordset.Fields("hora") = Str(Time)
    Inicio.datauditoria.Recordset.Fields("ventana") = "Recibo Cobro"
    Inicio.datauditoria.Recordset.Fields("accion") = "Emitido Recibo: " + text1.Text
    Inicio.datauditoria.Recordset.Fields("usuario") = login.usuarioactivo
    Inicio.datauditoria.Recordset.Fields("empresa") = login.empresaact
    Inicio.datauditoria.Recordset.UpdateBatch adAffectCurrent


    command3.Enabled = False
    Call nuevaorden_Click
    Call Command4_Click

Exit Sub
fuera:
    mensa = MsgBox("No puede ingresar como Instrumento de Cobro, dos Asignaciones referidas a la misma Orden", vbCritical, "Error")
    

End Sub

Private Sub Command4_Click()
On Error GoTo errorimp
Dim crApp As New CRAXDRT.Application
Dim crReport As New CRAXDRT.Report
Dim tabla As String
Dim tabla1 As String
Dim ruta As String


Rem If Inicio.Check3.Value = 0 Then
Rem        campo1 = Right("          " + Str(login.empresaact), 10)
Rem        campo2 = Left(login.conexiontotal + Space(255), 255)
Rem        campo3 = Left(login.conexionreporte + Space(255), 255)
Rem        campo4 = numorden
Rem        campototal = campo1 & campo2 & campo3 & campo4

Rem        Clipboard.SetText campototal
Rem        RetVal = Shell(App.Path & "\impresos.exe", vbNormalFocus)
Rem        Exit Sub
Rem End If

criterio.ConnectionString = login.conexiontotal

criterio.RecordSource = "select empreactiva.* from empreactiva"
criterio.Refresh

criterio.Recordset.Fields(0) = login.empresaact
criterio.Recordset.UpdateBatch adAffectCurrent

ruta = "\Empresa" + Right(Str(login.empresaact), Len(Str(login.empresaact)) - 1)

reporte.SQL = "consultarecibocobro.nrorden, consultarecibocobro.empresa, consultarecibocobro.nomproveedor, consultarecibocobro.comprobante, consultarecibocobro.fechacompro, consultarecibocobro.importe, consultarecibocobro.id, consultarecibocobro.razonsocial, consultarecibocobro.cuit, consultarecibocobro.domicilio, consultarecibocobro.localidad, consultarecibocobro.fecha, consultarecibocobro.domprov, consultarecibocobro.locprov, consultarecibocobro.cuitprov, consultarecibocobro.saldofactura FROM contablesql.dbo.consultarecibocobro consultarecibocobro WHERE consultarecibocobro.nrorden= '" & numorden & "' and consultarecibocobro.empresa = " & login.empresaact & " ORDER BY consultarecibocobro.razonsocial ASC, consultarecibocobro.id ASC"
tabla = reporte.SQL

With CrystalReporte
  
    .ReportFileName = App.Path & ruta + "\Recibocliente.rpt"
    .Connect = login.conexionreporte
 For x = 0 To 1
    .SubreportToChange = .GetNthSubreportName(x)
    .Connect = login.conexionreporte
    .SubreportToChange = ""
    .Connect = login.conexionreporte
 Next x
    .DiscardSavedData = True
    .RetrieveDataFiles
    .ReportSource = 0
    .SQLQuery = tabla
    .Destination = crptToFile
    .PrintFileName = App.Path & "\cobros.rpt"
    .PrintFileType = crptCrystal
    .Action = 1
                      
End With

If Inicio.Check3.Value = 1 Then
    Set crReport = crApp.OpenReport(App.Path & "\cobros.rpt", 1)
    impresos.CRViewer1.ReportSource = crReport
    impresos.CRViewer1.ViewReport
    impresos.Show
End If


errorimp:
End Sub


Private Sub Command6_Click()


For x = 0 To 3
    Text5(x).Text = ""
Next x
    importemask.Text = ""
    MaskEdBox2.Mask = ""
    MaskEdBox3.Mask = ""
    MaskEdBox2.Text = ""
    MaskEdBox3.Text = ""
    MaskEdBox2.Mask = "##/##/####"
    MaskEdBox3.Mask = "##/##/####"

    ordeninstu = ""

    codigopago = 0
    Text5(0).SetFocus

End Sub

Private Sub Command6_KeyUp(KeyCode As Integer, Shift As Integer)

On Error GoTo fuera
    
    If KeyCode = 113 And (SSTab1.Tab = 0 Or SSTab1.Tab = 2) Then
        SSTab1.Tab = 1
Rem        Call borrablancos_Click
        Command6.SetFocus
    End If
    
    If KeyCode = 112 And (SSTab1.Tab = 1 Or SSTab1.Tab = 2) Then
        SSTab1.Tab = 0
Rem        Call borrablancos_Click
        nuevo.SetFocus
    End If
    If KeyCode = 114 And (SSTab1.Tab = 1 Or SSTab1.Tab = 0) Then
        SSTab1.Tab = 2
  Rem      Call borrablancos_Click
        nuevootroconcepto.SetFocus
    End If

fuera:

End Sub

Private Sub Command7_Click()
On Error Resume Next

    If Text7.Text = "" Then
        mensa = MsgBox("No ingreso codigo contable en el concepto a Cobrar", vbExclamation, "Atencion")
        Exit Sub
    End If

    linea = grillapago.Row

    For x = 1 To 99
        grillapago.Row = x
        grillapago.Col = 0
        If grillapago.Text = "" Then GoTo sigue
    Next x
    
sigue:
    If codigopago1 = 0 Then
        grillapago.Row = x
    Else
        grillapago.Row = linea
    End If
    idcliente(grillapago.Row) = codprove
    idlibrogrid(grillapago.Row) = DataGrid4.Columns(0).Text
    grillapago.Text = Text6(0).Text
    grillapago.Col = 1
    grillapago.Text = Text6(1).Text
    grillapago.Col = 2
    grillapago.Text = Text6(2).Text
    grillapago.Col = 3
    grillapago.Text = importepago.Text
    grillapago.Text = Format(grillapago.Text, "#,###,##0.00")
    grillapago.Col = 4
    grillapago.Text = saldocomp.Text
    grillapago.Text = Format(grillapago.Text, "#,###,##0.00")
    Cuenta(grillapago.Row) = Text7.Text
    
For x = 0 To 2
    Text6(x).Text = ""
Next x
    importepago.Text = ""
    saldocomp.Text = ""
    
    codigopago1 = 0
    
    totalabonan.Text = 0
For x = 1 To 99
    grillapago.Col = 3
    grillapago.Row = x
    If grillapago.Text <> "" Then totalabonan.Text = Val(totalabonan.Text) + grillapago.Text
Next x
    saldototal.Text = Val(totalabonan.Text) - Val(totalinstrumento.Text)
    nuevo.SetFocus
    
fuera:
End Sub


Private Sub DataGrid1_ButtonClick(ByVal ColIndex As Integer)
On Error GoTo fuera

If ColIndex = 6 Then
        DataList1.Visible = True
        DataList1.Left = DataGrid1.Columns(6).Left + DataGrid1.Left + SSTab1.Left
        DataList1.Width = DataGrid1.Columns(6).Width
        DataList1.Top = DataGrid1.Top + DataGrid1.RowTop(DataGrid1.Row) + DataGrid1.RowHeight * 6
        DataList1.SetFocus
End If

If ColIndex = 7 Then
        DataList2.Visible = True
        DataList2.Left = DataGrid1.Columns(7).Left + DataGrid1.Left + SSTab1.Left
        DataList2.Width = DataGrid1.Columns(7).Width
        DataList2.Top = DataGrid1.Top + DataGrid1.RowTop(DataGrid1.Row) + DataGrid1.RowHeight * 6
        DataList2.SetFocus
End If
    
fuera:
End Sub



Private Sub DataGrid1_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo fuera

    If KeyCode = 113 Then
        SSTab1.Tab = 1
    Rem    Call borrablancos_Click
        Command6.SetFocus
    End If
    
    If KeyCode = 114 Then
        SSTab1.Tab = 2
        Call borrablancos_Click
        nuevootroconcepto.SetFocus
    End If

fuera:
End Sub

Private Sub DataGrid2_AfterColUpdate(ByVal ColIndex As Integer)
    
    If DataGrid2.Col = 11 And DataGrid2.Columns(6).Text = "EFECTIVO" Then
            If Inicio.montoefectivo <> 0 And DataGrid2.Columns(11).Value > Inicio.montoefectivo Then
                mensa = MsgBox("El Monto en efectivo a Cobrar es mayor a $" + Str(Inicio.montoefectivo) + " Verifique", vbExclamation, "!! Atención !!")
            End If
    End If

End Sub

Private Sub DataGrid2_ButtonClick(ByVal ColIndex As Integer)
On Error GoTo fuera

If ColIndex = 6 Then
        DataList3.Visible = True
        DataList3.Left = DataGrid2.Columns(6).Left + DataGrid2.Left + SSTab1.Left
        DataList3.Width = DataGrid2.Columns(6).Width
        DataList3.Top = DataGrid2.Top + DataGrid2.RowTop(DataGrid2.Row) + DataGrid2.RowHeight * 6
        DataList3.SetFocus
End If

If ColIndex = 12 Then
        DataList4.Visible = True
        DataList4.Width = DataGrid2.Columns(12).Width * 4
        DataList4.Left = DataGrid2.Columns(12).Left + DataGrid2.Left + DataGrid2.Columns(12).Width - DataList4.Width + SSTab1.Left
        DataList4.Top = DataGrid2.Top + DataGrid2.RowTop(DataGrid2.Row) + DataGrid2.RowHeight * 6
        DataList4.SetFocus
End If

fuera:
End Sub

Private Sub DataGrid2_KeyPress(KeyAscii As Integer)
On Error GoTo erroringreso1



    If KeyAscii = 13 And DataGrid2.Col = 6 Then
        If DataGrid2.Columns(6).Text = "" Then
                    DataList3.Visible = True
                    DataList3.Left = DataGrid2.Columns(6).Left + DataGrid2.Left + SSTab1.Left
                    DataList3.Width = DataGrid2.Columns(6).Width
                    DataList3.Top = DataGrid2.Top + DataGrid2.RowTop(DataGrid2.Row) + DataGrid2.RowHeight * 6
                    DataList3.SetFocus
                    KeyAscii = 0
                    Exit Sub
        Else
            KeyAscii = 9
        End If
    End If
    
    If KeyAscii = 13 And DataGrid2.Col = 12 Then
        If DataGrid2.Columns(12).Text = "" Then
                    DataList4.Visible = True
                    DataList4.Width = DataGrid2.Columns(12).Width * 4
                    DataList4.Left = DataGrid2.Columns(12).Left + DataGrid2.Left + DataGrid2.Columns(12).Width - DataList4.Width + SSTab1.Left
                    DataList4.Top = DataGrid2.Top + DataGrid2.RowTop(DataGrid2.Row) + DataGrid2.RowHeight * 6
                    DataList4.SetFocus
                    KeyAscii = 0
                    Exit Sub
        End If
    End If
    If KeyAscii = 13 And DataGrid2.Col = 12 Then
        If DataGrid2.Columns(12).Text <> "" Then
            totalinst(DataGrid2.Row) = DataGrid2.Columns(11).Text
            cuentaint(DataGrid2.Row) = DataGrid2.Columns(12).Text
            detalleint(DataGrid2.Row) = DataGrid2.Columns(6).Text + " " + DataGrid2.Columns(7).Text
            datinstrumento.Recordset.UpdateBatch adAffectCurrent
                
            totalinstrumento.Text = 0
            For x = 0 To DataGrid2.Row
                totalinstrumento.Text = Val(totalinstrumento.Text) + totalinst(x)
            Next x
            saldototal.Text = totalinstrumento.Text - totalabonan.Text
        End If
        Command1.SetFocus
    End If
    
                    
    If KeyAscii = 13 Then KeyAscii = 9
Exit Sub
erroringreso1:
    Call Command1.SetFocus


End Sub

Private Sub DataGrid2_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo fuera

    If KeyCode = 115 Then frmrecibointrumento.Show

    If KeyCode = 112 Then
        SSTab1.Tab = 0
        Call borrablancos_Click
        nuevo.SetFocus
    End If

fuera:
End Sub

Private Sub Command8_Click()
On Error GoTo fuera
Dim tb As Currency

    codprove = 0
    linea = grillainstru.Row

    For x = 1 To 99
        grillaotros.Row = x
        grillaotros.Col = 0
        If grillaotros.Text = "" Then GoTo sigue
    Next x
    
sigue:
    If codigopago = 0 Then
        grillaotros.Row = x
    Else
        grillaotros.Row = linea
    End If
    
    grillaotros.Text = Text6(4).Text
    grillaotros.Col = 1
    grillaotros.Text = importeotro.Text
    grillaotros.Text = Format(grillaotros.Text, "#,###,##0.00")
    grillaotros.Col = 2
    grillaotros.Text = Text6(5).Text
    
For x = 4 To 5
    Text6(x).Text = ""
Next x
    importeotro.Text = ""
   
    codigopago = 0
    
    totalabonan.Text = 0
    tb = 0
    If totalinstrumento.Text = "" Then totalinstrumento.Text = 0
For x = 1 To 99
    grillaotros.Col = 1
    grillaotros.Row = x
    If grillaotros.Text <> "" Then tb = tb + grillaotros.Text
    
Next x
    saldototal.Text = tb - totalinstrumento.Text
    totalabonan.Text = tb
    nuevootroconcepto.SetFocus

fuera:


End Sub

Private Sub DataGrid5_ButtonClick(ByVal ColIndex As Integer)
On Error GoTo fuera

If ColIndex = 10 Then
        DataList4.Visible = True
        DataList4.Width = DataGrid5.Columns(10).Width * 4
        DataList4.Left = DataGrid5.Columns(10).Left + DataGrid5.Left + DataGrid5.Columns(10).Width - DataList4.Width + SSTab1.Left
        DataList4.Top = DataGrid5.Top + DataGrid5.RowTop(DataGrid5.Row) + DataGrid5.RowHeight * 6
        DataList4.SetFocus
End If

fuera:
End Sub

Private Sub compfecha_Click()


    campoaño = Right(MaskEdBox1.Text, 4)
    campomes = Mid(MaskEdBox1.Text, 4, 2)
    campodia = Left(MaskEdBox1.Text, 2)
    campofecha = campoaño + "/" + campomes + "/" + campodia
    
    campoaño1 = Right(login.iper, 4)
    campomes1 = Mid(login.iper, 4, 2)
    campodia1 = Left(login.iper, 2)
    campofecha1 = campoaño1 + "/" + campomes1 + "/" + campodia1
    
    campoaño2 = Right(login.fper, 4)
    campomes2 = Mid(login.fper, 4, 2)
    campodia2 = Left(login.fper, 2)
    campofecha2 = campoaño2 + "/" + campomes2 + "/" + campodia2
    campofecha3 = Right(fechafuera, 4) + "/" + Mid(fechafuera, 4, 3) + Left(fechafuera, 2)
    fechamal = 0
    If campofecha < campofecha1 Or campofecha > campofecha2 Then
            mensa = MsgBox("La Fecha no pertenecia al periodo en ejercicio", vbCritical, "!! Atención !!")
            fechamal = 1
            Unload Me
            frmrecibos1.Show
    End If

End Sub

Private Sub DataList1_KeyPress(KeyAscii As Integer)
On Error GoTo errorcarga
    If KeyAscii = 13 Then
        KeyPress = 0
            Text6(0).Text = DataList1.Text
            razonsocial1 = DataList1.Text
            codprove = DataList1.BoundText
            DataList1.Visible = False
            
            For Y = 1 To 99
                grillapago.Row = Y
                grillapago.Col = 0
             Rem   If grillapago.Text <> "" Then
             Rem       If grillapago.Text <> Text6(0).Text Then
             Rem             mensa = MsgBox("No se puede ingresar a un Cliente Diferente", vbCritical, "Error")
             Rem             Text6(0).Text = ""
             Rem             Exit Sub
             Rem       End If
             Rem   End If
            Next Y
            
            
            datconsultacomp.RecordSource = "select conceptoscobran.* from conceptoscobran where empresa = " & login.empresaact & " and codcliente = " & codprove & " order by comp"
            datconsultacomp.Refresh
            If datconsultacomp.Recordset.EOF = False Then
                       Cuenta(grillapago.Row) = Val(DataGrid3.Columns(10).Text)
            Else
                datproveedores.RecordSource = "select clientes.* from clientes where empresa = " & login.empresaact & " and codcliente = " & codprove & " "
                datproveedores.Refresh
                If IsNull(datproveedores.Recordset.Fields("codcontable")) = True Then GoTo errorcarga
                Cuenta(grillapago.Row) = datproveedores.Recordset.Fields("codcontable")
                Text7.Text = Cuenta(grillapago.Row)
                datproveedores.RecordSource = "select clientes.* from clientes where empresa = " & login.empresaact & " ORDER BY razonsocial"
                datproveedores.Refresh
            End If
            Text6(1).SetFocus
    End If
Exit Sub

errorcarga:
    mensa = MsgBox("El Cliente no tiene cargado un codigo contable, verifique el archivo de Proveedores", vbCritical, "Error")
    Exit Sub
End Sub

Private Sub DataList1_LostFocus()
        DataList1.Visible = False
End Sub

Private Sub DataList2_KeyPress(KeyAscii As Integer)
On Error GoTo fuera
    
    If KeyAscii = 13 Then
            Text6(1).Text = DataList2.Text
            compro = DataList2.Text
            If compro = "" Then
                sincomp = sincomp + 1
                If sincomp = 2 Then
                    mensa = MsgBox("No puede ingresar mas de un anticipo por Recibo de Cobro", vbCritical, "Error")
                    Exit Sub
                End If
                Text7.SetFocus
            End If
  
      Rem      nrocompro(DataGrid1.Row) = DataList2.Text
      Rem      nomprov(DataGrid1.Row) = DataGrid1.Columns(6).Text
            For Y = 1 To 99
                grillapago.Row = Y
                grillapago.Col = 1
                If grillapago.Text <> "" Then
                    If grillapago.Text = Text6(1).Text Then
                          mensa = MsgBox("Este comprobante ya fue ingresado", vbCritical, "Error")
                          Text6(1).Text = ""
                          Exit Sub
                    End If
                End If
            Next Y
                    
            
            datconsultacomp.RecordSource = "select conceptoscobran.* from conceptoscobran where empresa = " & login.empresaact & " and codcliente = " & codprove & " and comp = '" & compro & "'  order by comp "
            datconsultacomp.Refresh
            idlibro = DataGrid3.Columns(10).Text
            Text6(2).Text = DataGrid3.Columns(0).Text
            Cuenta(grillapago.Row) = Val(DataGrid3.Columns(9).Text)
            Text7.Text = Cuenta(grillapago.Row)
            If DataGrid3.Columns(11).Text = "" Then
                saldocomp.Text = DataGrid3.Columns(5).Text
            Else
                saldocomp.Text = DataGrid3.Columns(8).Text
            End If
            Rem importepago.Text = "0"
            DataList2.Visible = False
            datlibrocompras.RecordSource = "select libroventas.* from libroventas WHERE libroventas.empresa = " & login.empresaact & " and id = " & idlibro & " Order by id"
            datlibrocompras.Refresh
            Text7.SetFocus
        Rem    importepago.SetFocus
    End If
    
fuera:
End Sub

Private Sub DataList2_LostFocus()
            DataList2.Visible = False
End Sub

Private Sub DataList3_KeyPress(KeyAscii As Integer)
On Error GoTo fuera
    
    If KeyAscii = 13 Then
            Text5(0).Text = DataList3.Text
            Text5(1).Text = DataList3.BoundText
            DataList3.Visible = False
            importemask.SetFocus
    End If
    
fuera:
End Sub

Private Sub DataList3_KeyUp(KeyCode As Integer, Shift As Integer)

 If KeyCode = 115 Then frmrecibointrumento.Show


End Sub

Private Sub DataList3_LostFocus()
    DataList3.Visible = False
End Sub

Private Sub DataList4_GotFocus()
On Error GoTo fuera
    
   
    If Inicio.opcion1 = True Then
        datcuentas.RecordSource = "select listacuentas.* from listacuentas WHERE empre = " & login.empresaact & " and inicioper = '" & login.iper & "' and finper = '" & login.fper & "' ORDER BY IDCUENTA"
        datcuentas.Refresh
        DataList4.ListField = "codigo"
    End If
    
    If Inicio.opcion2 = True Then
        datcuentas.RecordSource = "select listacuentas.* from listacuentas WHERE empre = " & login.empresaact & " and inicioper = '" & login.iper & "' and finper = '" & login.fper & "' ORDER BY nombre"
        datcuentas.Refresh
        DataList4.ListField = "nombre"
    End If
    
    If SSTab1.Tab = 0 Then
        DataList4.Top = Text7.Top + Text7.Height + SSTab1.Top
        DataList4.Left = SSTab1.Left + Text7.Left + Text7.Width - DataList4.Width
        If Text7.Text <> "" Then DataList4.BoundText = Text7.Text
    End If
    
    If SSTab1.Tab = 1 Then
        DataList4.Top = Text5(0).Top + Text5(1).Height + SSTab1.Top
        DataList4.Left = SSTab1.Left + Text5(1).Left + Text5(1).Width - DataList4.Width
        If Text5(1).Text <> "" Then DataList4.BoundText = Text5(1).Text
    End If
    
    If SSTab1.Tab = 2 Then
        DataList4.Top = Text6(5).Top + Text6(5).Height + SSTab1.Top
        DataList4.Left = SSTab1.Left + Text6(5).Left + Text6(5).Width - DataList4.Width
        If Text6(5).Text <> "" Then DataList4.BoundText = Text6(5).Text
    End If
    
    
fuera:
End Sub

Private Sub DataList4_KeyPress(KeyAscii As Integer)
On Error GoTo fuera

    If KeyAscii = 13 And SSTab1.Tab = 2 Then
            Text6(5).Text = DataList4.BoundText
            If Text6(5).Text = "" Then
                mensa = MsgBox("Debe Ingresar un Cod.Contable", vbExclamation, "Error")
                Exit Sub
            End If
            DataList4.Visible = False
            command8.SetFocus
            Exit Sub
    End If

    If KeyAscii = 13 And SSTab1.Tab = 1 Then
            Text5(1).Text = DataList4.BoundText
            If Text5(1).Text = "" Then
                mensa = MsgBox("Debe Ingresar un Cod.Contable", vbExclamation, "Error")
                Exit Sub
            End If
            DataList4.Visible = False
            Text5(3).SetFocus
    End If
    
    If KeyAscii = 13 And SSTab1.Tab = 0 Then
            Text7.Text = DataList4.BoundText
            Cuenta(grillapago.Row) = Text7.Text
            If Text7.Text = "" Then
                mensa = MsgBox("Debe Ingresar un Cod.Contable", vbExclamation, "Error")
                Exit Sub
            End If
            importepago.SetFocus
            DataList4.Visible = False
            
    End If
    
fuera:
End Sub

Private Sub DataList4_LostFocus()

    DataList4.Visible = False

End Sub

Private Sub eliminaotroconcepto_Click()
On Error GoTo erroreliminar1


    For x = 0 To 2
        grillaotros.Col = x
        grillaotros.Text = ""
    Next x
    
    For x = 4 To 5
        Text6(x).Text = ""
    Next x
    importeotro.Text = ""
    
    codigopago = 0
    
    totalabonan.Text = 0
    tb = 0
    If totalinstrumento.Text = "" Then totalinstrumento.Text = 0
For x = 1 To 99
    grillaotros.Col = 1
    grillaotros.Row = x
    If grillaotros.Text <> "" Then tb = tb + grillaotros.Text
    
Next x
    saldototal.Text = tb - totalinstrumento.Text
    totalabonan.Text = tb
    nuevootroconcepto.SetFocus

erroreliminar1:
End Sub

Private Sub Form_Load()
Aplicar_skin Me

frmrecibos1.Top = 0
frmrecibos1.Left = 0

databonan.ConnectionString = login.conexiontotal
datasiento.ConnectionString = login.conexiontotal
datconsultacomp.ConnectionString = login.conexiontotal
datcuentas.ConnectionString = login.conexiontotal
datinstru.ConnectionString = login.conexiontotal
datinstrumento.ConnectionString = login.conexiontotal
datinstrumento1.ConnectionString = login.conexiontotal
datlibrocompras.ConnectionString = login.conexiontotal
datmaestro.ConnectionString = login.conexiontotal
datordendepago.ConnectionString = login.conexiontotal
datproveedores.ConnectionString = login.conexiontotal
datasigna.ConnectionString = login.conexiontotal
Data1.Connection1.ConnectionString = login.conexiontotal

ordeninstu = ""
For x = 1 To 50
    ordeninstruver(x) = ""
Next x
For x = 1 To 150
    Cuenta(x) = 0
Next x

totalabonan.Text = 0
totalinstrumento.Text = 0


If login.empresaact1 > 0 Then
    empresareal = login.empresaact1
Else
    empresareal = login.empresaact
End If

If login.librocerrado = "S" Then
 MaskEdBox1.Enabled = True
Else
  MaskEdBox1.Enabled = False
End If

    Inicio.Toolbar1.Visible = True
    
    sincomp = 0


  datmaestro.RecordSource = "SELECT  MAX(nroasiento) AS nroasiento, perinicial, empresa From dbo.[Maestro Asientos] GROUP BY perinicial, empresa HAVING (perinicial = '" & login.iper & "' and empresa = " & login.empresaact & " )"
  datmaestro.Refresh

  datinstru.RecordSource = "select instrumentospagos.* from instrumentospagos where empresa = " & login.empresaact & " and (tipo = 'R' or tipo = 'OR' or tipo = 'RO')"
  datinstru.Refresh

  
  datproveedores.RecordSource = "select clientes.* from clientes where empresa = " & empresareal & " ORDER BY razonsocial"
  datproveedores.Refresh
  
  datcuentas.RecordSource = "select listacuentas.* from listacuentas WHERE inicioper = '" & login.iper & "' and empre = " & login.empresaact & " ORDER BY IDCUENTA"
  datcuentas.Refresh
  
  MaskEdBox1.Text = Date
  totalab = 0
  Check1(0).Value = 1
  SSTab1.Tab = 0
  
  Call veriorden_Click
   
  MaskEdBox1.Mask = "##/##/####"
  datordendepago.Recordset.Fields(0) = text1.Text
  datordendepago.Recordset.Fields(1) = login.empresaact
  datordendepago.Recordset.Fields(2) = login.iper
  datordendepago.Recordset.Fields(3) = login.fper
  datordendepago.Recordset.Fields(4) = MaskEdBox1.Text
  orden = text1.Text
  
  databonan.RecordSource = "select recibocobroabonan.* from recibocobroabonan WHERE inicioper = '" & login.iper & "' and nrorden = '" & orden & "' and recibocobroabonan.empresa = " & login.empresaact & " Order by fechacompro"
  databonan.Refresh
  
  datinstrumento.RecordSource = "select recibocobroinstrumento.* from recibocobroinstrumento WHERE inicioper = '" & login.iper & "' and nrorden = '" & orden & "'  and recibocobroinstrumento.empresa = " & login.empresaact & " Order by id"
  datinstrumento.Refresh
  

grillainstru.Row = 0
grillainstru.Col = 0
grillainstru.ColWidth(0) = 2000
grillainstru.Text = "Forma de Cobro"
grillainstru.Col = 1
grillainstru.ColWidth(1) = 1200
grillainstru.Text = "Importe"
grillainstru.Col = 2
grillainstru.ColWidth(2) = 800
grillainstru.Text = "Cuenta"
grillainstru.Col = 3
grillainstru.ColWidth(3) = 2000
grillainstru.Text = "Denominacion"
grillainstru.Col = 4
grillainstru.ColWidth(4) = 1500
grillainstru.Text = "Nº Comprobante"
grillainstru.Col = 5
grillainstru.ColWidth(5) = 1200
grillainstru.Text = "Fecha Emisión"
grillainstru.Col = 6
grillainstru.ColWidth(6) = 1200
grillainstru.Text = "Fecha Venc."

grillapago.Row = 0
grillapago.Col = 0
grillapago.ColWidth(0) = 3200
grillapago.Text = "Cliente"
grillapago.Col = 1
grillapago.ColWidth(1) = 2000
grillapago.Text = "Nº Comprobante"
grillapago.Col = 2
grillapago.ColWidth(2) = 1500
grillapago.Text = "Fecha Comprob."
grillapago.Col = 3
grillapago.ColWidth(3) = 1200
grillapago.Text = "Importe"
grillapago.Col = 4
grillapago.ColWidth(4) = 1200
grillapago.Text = "Saldo Comprob."

grillaotros.Row = 0
grillaotros.Col = 0
grillaotros.ColWidth(0) = 3200
grillaotros.Text = "Detalle Concepto"
grillaotros.Col = 1
grillaotros.ColWidth(1) = 2000
grillaotros.Text = "Importe"
grillaotros.Col = 2
grillaotros.ColWidth(2) = 1000
grillaotros.Text = "Cod.Cuenta"

For x = 1 To 14 Step 2
    For Y = 0 To 6
        grillainstru.Col = Y
        grillainstru.Row = x
        grillainstru.CellBackColor = QBColor(11)
    Next Y
Next x

For x = 1 To 100 Step 2
    For Y = 0 To 4
        grillapago.Col = Y
        grillapago.Row = x
        grillapago.CellBackColor = QBColor(11)
    Next Y
Next x

For x = 1 To 100 Step 2
    For Y = 0 To 2
        grillaotros.Col = Y
        grillaotros.Row = x
        grillaotros.CellBackColor = QBColor(11)
    Next Y
Next x

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Inicio.Toolbar1.Visible = False

End Sub

Private Sub grillainstru_Click()
On Error GoTo fuera

    codigopago = 1
    grillainstru.Col = 0
    Text5(0).Text = grillainstru.Text
    grillainstru.Col = 1
    importemask.Text = grillainstru.Text
    grillainstru.Col = 2
    Text5(1).Text = grillainstru.Text
    grillainstru.Col = 3
    Text5(2).Text = grillainstru.Text
    grillainstru.Col = 4
    Text5(3).Text = grillainstru.Text
    grillainstru.Col = 5
    MaskEdBox2.Text = grillainstru.Text
    grillainstru.Col = 6
    MaskEdBox3.Text = grillainstru.Text
    grillainstru.Col = 0
    grillainstru.RowSel = grillainstru.Row
    grillainstru.ColSel = 6


fuera:
End Sub

Private Sub grillainstru_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo fuera

    codigopago = 1
    grillainstru.Col = 0
    Text5(0).Text = grillainstru.Text
    grillainstru.Col = 1
    importemask.Text = grillainstru.Text
    grillainstru.Col = 2
    Text5(1).Text = grillainstru.Text
    grillainstru.Col = 3
    Text5(2).Text = grillainstru.Text
    grillainstru.Col = 4
    Text5(3).Text = grillainstru.Text
    grillainstru.Col = 5
    MaskEdBox2.Text = grillainstru.Text
    grillainstru.Col = 6
    MaskEdBox3.Text = grillainstru.Text
    grillainstru.Col = 0
    grillainstru.RowSel = grillainstru.Row
    grillainstru.ColSel = 6
    


fuera:
End Sub

Private Sub grillaotros_Click()
On Error GoTo fuera

    codigopago = 1
    grillaotros.Col = 0
    Text6(4).Text = grillaotros.Text
    grillaotros.Col = 1
    importeotro.Text = grillaotros.Text
    grillaotros.Text = Format(grillaotros.Text, "#,###,##0.00")
    grillaotros.Col = 2
    Text6(5).Text = grillaotros.Text
    grillaotros.Col = 0
    grillaotros.RowSel = grillaotros.Row
    grillaotros.ColSel = 2


fuera:
End Sub

Private Sub grillaotros_KeyDown(KeyCode As Integer, Shift As Integer)

On Error GoTo fuera

    codigopago = 1
    grillaotros.Col = 0
    Text6(4).Text = grillaotros.Text
    grillaotros.Col = 1
    importeotro.Text = grillaotros.Text
    grillaotros.Text = Format(grillaotros.Text, "#,###,##0.00")
    grillaotros.Col = 2
    Text6(5).Text = grillaotros.Text
    grillaotros.Col = 0
    grillaotros.RowSel = grillaotros.Row
    grillaotros.ColSel = 2


fuera:

End Sub

Private Sub grillapago_Click()
On Error GoTo fuera

    codigopago1 = 1
    grillapago.Col = 0
    Text6(0).Text = grillapago.Text
    grillapago.Col = 1
    Text6(1).Text = grillapago.Text
    grillapago.Col = 2
    Text6(2).Text = grillapago.Text
    grillapago.Col = 3
    importepago.Text = grillapago.Text
    grillapago.Col = 4
    saldocomp.Text = grillapago.Text
    grillapago.Col = 0
    grillapago.RowSel = grillapago.Row
    grillapago.ColSel = 4
    
fuera:

End Sub

Private Sub grillapago_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo fuera

    codigopago1 = 1
    grillapago.Col = 0
    Text6(0).Text = grillapago.Text
    grillapago.Col = 1
    Text6(1).Text = grillapago.Text
    grillapago.Col = 2
    Text6(2).Text = grillapago.Text
    grillapago.Col = 3
    importepago.Text = grillapago.Text
    grillapago.Col = 4
    saldocomp.Text = grillapago.Text
    grillapago.Col = 0
    grillapago.RowSel = grillapago.Row
    grillapago.ColSel = 4
    
fuera:

End Sub

Private Sub importemask_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        If importemask.Text = "" Then
            mensa = MsgBox("Debe Ingresar un importe", vbExclamation, "Error")
            Exit Sub
        End If
        If Val(importemask.Text) > Inicio.montoefectivo And Text5(0).Text = "EFECTIVO" Then
            mensa = MsgBox("El importe es mayor que el monto máximo a pagar en efectivo", vbExclamation, "Error")
            Exit Sub
        End If
        
        If ordeninstu <> "" Then
            If Val(importemask.Text) > importeord Then
                       mensa = MsgBox("Importe incorrecto", vbCritical, "Error")
                       importemask.Text = 0
                       Exit Sub
            End If
        End If
                                        
        SendKeys "{tab}", True
    End If

End Sub

Private Sub importeotro_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        If importeotro.Text = "" Then
            mensa = MsgBox("Debe Ingresar un importe", vbExclamation, "Error")
            Exit Sub
        End If
        SendKeys "{tab}", True
    End If

End Sub

Private Sub importepago_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        
        If importepago.Text = "" Then importepago.Text = 0
        dif = Val(importepago.Text) - Val(saldocomp.Text)
        If Val(importepago.Text) < 0 Then dif = dif * -1
        If dif > 0.01 Then
              If Text6(1).Text <> "" Then
                mensa = MsgBox("El valor ingresado es superior al saldo de la Factura, no se puede ingresar", vbCritical, "Error")
                Exit Sub
              End If
        End If
        
        If Val(importepago.Text) = 0 Then
                mensa = MsgBox("Debe ingresar un Importe Valido, no puede ingresar 0", vbCritical, "Error")
                Exit Sub
        End If
        
        If Text6(1).Text <> "" Then saldocomp = saldocomp - importepago

        command7.SetFocus
        
        
    End If

End Sub

Private Sub MaskEdBox1_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = 13 Then
        KeyAscii = 0
        If Right(MaskEdBox1.Text, 2) = "__" Then
            MaskEdBox1.Text = Left(MaskEdBox1.Text, 6) + "20" + Mid(MaskEdBox1.Text, 7, 2)
        End If
        Call compfecha_Click
        
        nuevo.SetFocus
    End If
    
End Sub

Private Sub MaskEdBox2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If Right(MaskEdBox2.Text, 2) = "__" And Mid(MaskEdBox2.Text, 7, 2) <> "__" Then
            MaskEdBox2.Text = Left(MaskEdBox2.Text, 6) + "20" + Mid(MaskEdBox2.Text, 7, 2)
        End If
        SendKeys "{tab}", True
    End If
End Sub

Private Sub MaskEdBox3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If Right(MaskEdBox3.Text, 2) = "__" And Mid(MaskEdBox3.Text, 7, 2) <> "__" Then
            MaskEdBox3.Text = Left(MaskEdBox3.Text, 6) + "20" + Mid(MaskEdBox3.Text, 7, 2)
        End If
        SendKeys "{tab}", True
    End If
End Sub

Private Sub nuevaorden_Click()
   
Rem    totalabonan.Text = "0.00"
Rem    totalinstrumento.Text = "0.00"
Rem    saldototal.Text = "0.00"
Rem    DataGrid1.Columns(6).Caption = "Proveedor Razon Social"
Rem    DataGrid1.Columns(7).Visible = True
Rem    DataGrid1.Columns(8).Caption = "Fecha Compr."
Rem    DataGrid1.Columns(10).Visible = False
Rem    DataGrid1.Columns(10).Width = 1395
Rem    DataGrid1.Columns(11).Locked = True
Rem    sincomp = 0
    Unload Me
    frmrecibos1.Show
    If Inicio.Check3.Value = 1 Then impresos.Show

End Sub

Private Sub nuevo_Click()

For x = 0 To 2
    Text6(x).Text = ""
Next x
    importepago.Text = ""
    saldocomp.Text = ""

    codigopago1 = 0
    grillapago.Row = 1
    grillapago.Col = 0
    Text6(0).Text = grillapago.Text
    Text6(0).SetFocus

End Sub

Private Sub nuevo_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo fuera

    If KeyCode = 113 Then
        SSTab1.Tab = 1
Rem        Call borrablancos_Click
        Command6.SetFocus
    End If
    If KeyCode = 114 Then
        SSTab1.Tab = 2
        Call borrablancos_Click
        nuevootroconcepto.SetFocus
    End If
    
fuera:
End Sub

Private Sub nuevootroconcepto_Click()
On Error GoTo fuera


For x = 4 To 5
    Text6(x).Text = ""
Next x
    importeotro.Text = ""

    codigopago1 = 0
    Text6(4).SetFocus
 
fuera:
 
End Sub

Private Sub nuevootroconcepto_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo fuera
    
    If KeyCode = 113 And (SSTab1.Tab = 0 Or SSTab1.Tab = 2) Then
        SSTab1.Tab = 1
Rem        Call borrablancos_Click
        Command6.SetFocus
    End If
    
    If KeyCode = 112 And (SSTab1.Tab = 1 Or SSTab1.Tab = 2) Then
        SSTab1.Tab = 0
Rem        Call borrablancos_Click
        nuevo.SetFocus
    End If
    If KeyCode = 114 And (SSTab1.Tab = 1 Or SSTab1.Tab = 0) Then
        SSTab1.Tab = 2
  Rem      Call borrablancos_Click
        nuevootroconcepto.SetFocus
    End If

fuera:
End Sub

Private Sub salir_Click()

    Unload Me

End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    If SSTab1.Tab = 1 Then Command6.SetFocus
    If SSTab1.Tab = 2 Then nuevootroconcepto.SetFocus
End Sub

Private Sub SSTab1_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo fuera
    
    If KeyCode = 113 And (SSTab1.Tab = 0 Or SSTab1.Tab = 2) Then
        SSTab1.Tab = 1
Rem        Call borrablancos_Click
        Command6.SetFocus
    End If
    
    If KeyCode = 112 And (SSTab1.Tab = 1 Or SSTab1.Tab = 2) Then
        SSTab1.Tab = 0
Rem        Call borrablancos_Click
        nuevo.SetFocus
    End If
    If KeyCode = 114 And (SSTab1.Tab = 1 Or SSTab1.Tab = 0) Then
        SSTab1.Tab = 2
  Rem      Call borrablancos_Click
        nuevootroconcepto.SetFocus
    End If

fuera:
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
On Error GoTo fuera

    If KeyAscii = 13 Then
            If Val(Mid(text1.Text, 1, 4)) = 0 Then
                mensa = MsgBox("Debe ingresar una sucursal en el Nro de factura", vbCritical, "!! Atención !!")
                text1.SetFocus
                text1.SelStart = 0
                text1.SelLength = 4
                Exit Sub
            End If
            If Right(text1.Text, 1) = "_" Then
                mensa = MsgBox("Nro de factura incorrecto", vbCritical, "!! Atención !!")
                text1.SetFocus
                text1.SelStart = 5
                text1.SelLength = 8
                Exit Sub
            End If
    End If

fuera:
End Sub

Private Sub Text5_GotFocus(Index As Integer)

   
    If Index = 1 Then
            DataList4.Visible = True
            DataList4.SetFocus
    End If
        
        
        

End Sub

Private Sub Text5_KeyPress(Index As Integer, KeyAscii As Integer)

    If KeyAscii = 13 And Index = 0 Then
        DataList3.Visible = True
        If Text5(0).Text <> "" Then
            DataList3.Text = Text5(0).Text
        Else
            DataList3.Text = "EFECTIVO"
        End If
        DataList3.SetFocus
        Exit Sub
    End If
    
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{tab}", True
    End If
        

End Sub

Private Sub Text5_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo fuera


    If KeyCode = 115 Then frmrecibointrumento.Show


    If KeyCode = 112 Then
        SSTab1.Tab = 0
        Call borrablancos_Click
        nuevo.SetFocus
    End If

fuera:

End Sub

Private Sub Text6_GotFocus(Index As Integer)

    If Index = 5 Then
        DataList4.Visible = True
        If Text6(5).Text <> "" Then DataList4.Text = Text6(5).Text
        DataList4.SetFocus
        Exit Sub
    End If
    
    If Index = 0 Then
        DataList1.Visible = True
        If Text6(0).Text <> "" Then DataList1.Text = Text6(0).Text
        DataList1.SetFocus
        Exit Sub
    End If


    If Index = 1 Then
        DataList2.Visible = True
        If Text6(1).Text <> "" Then DataList2.Text = Text6(1).Text
        DataList2.SetFocus
        Exit Sub
    End If


End Sub

Private Sub Text6_KeyPress(Index As Integer, KeyAscii As Integer)

    If KeyAscii = 13 And Index = 4 Then
        KeyAscii = 0
        importeotro.SetFocus
    End If


End Sub

Private Sub Text7_GotFocus()

   
            DataList4.Visible = True
            DataList4.SetFocus
    
End Sub

Private Sub Text8_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = 13 Then
        KeyAscii = 0
        MaskEdBox1.SetFocus
    End If

End Sub

Private Sub veriorden_Click()

  datordendepago.RecordSource = "select recibocobro.* from recibocobro WHERE recibocobro.empresa = " & login.empresaact & " Order by nrorden"
  datordendepago.Refresh
  
  If datordendepago.Recordset.EOF = True Then
        datordendepago.Recordset.AddNew
        text1.Text = "00000001"
        text1.Enabled = True
  Else
        datordendepago.Recordset.MoveLast
        nroorden = datordendepago.Recordset.Fields(0)
        pruebaorden = IsNull(datordendepago.Recordset.Fields(5))
        If pruebaorden = True Then
            text1.Text = nroorden
            GoTo paso0
        End If
        previo = Str(Val((Right(nroorden, 8))) + 1)
        previo1 = Right(previo, Len(previo) - 1)
        mitad2 = Mid("00000000", 1, 8 - Len(previo1)) + previo1
        datordendepago.Recordset.AddNew
        text1.Text = mitad2
paso0:
        text1.Enabled = False
  End If
  

End Sub
