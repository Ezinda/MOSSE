VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{D18BBD1F-82BB-4385-BED3-E9D31A3E361E}#1.0#0"; "KewlButtonz.ocx"
Begin VB.Form frmusuarios 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Usuarios"
   ClientHeight    =   8610
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11220
   Icon            =   "frmusuarios1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8610
   ScaleWidth      =   11220
   Begin VB.CommandButton Command1 
      Caption         =   "Administrador:"
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
      Left            =   720
      TabIndex        =   123
      TabStop         =   0   'False
      Top             =   1080
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Password:"
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
      Left            =   720
      TabIndex        =   122
      TabStop         =   0   'False
      Top             =   720
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Nombre de Usuario:"
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
      Left            =   720
      TabIndex        =   121
      TabStop         =   0   'False
      Top             =   360
      Width           =   1815
   End
   Begin MSDataListLib.DataList DataList1 
      Bindings        =   "frmusuarios1.frx":0442
      Height          =   2010
      Left            =   8160
      TabIndex        =   61
      Top             =   120
      Visible         =   0   'False
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   3545
      _Version        =   393216
      ListField       =   "razonsocial"
      BoundColumn     =   "empresa"
   End
   Begin MSDataGridLib.DataGrid DataGrid2 
      Bindings        =   "frmusuarios1.frx":045C
      Height          =   975
      Left            =   5160
      TabIndex        =   57
      Top             =   120
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   1720
      _Version        =   393216
      AllowUpdate     =   -1  'True
      AllowArrows     =   -1  'True
      BackColor       =   -2147483626
      HeadLines       =   1
      RowHeight       =   15
      TabAction       =   1
      WrapCellPointer =   -1  'True
      AllowAddNew     =   -1  'True
      AllowDelete     =   -1  'True
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
   Begin MSAdodcLib.Adodc datusuarios 
      Height          =   330
      Left            =   6720
      Top             =   7560
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmusuarios1.frx":0476
      Height          =   8415
      Left            =   8160
      TabIndex        =   60
      Top             =   120
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   14843
      _Version        =   393216
      BackColor       =   -2147483626
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
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
      ColumnCount     =   3
      BeginProperty Column00 
         DataField       =   "nombre"
         Caption         =   "Usuarios"
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
         DataField       =   "password"
         Caption         =   "password"
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
      BeginProperty Column02 
         DataField       =   "administrador"
         Caption         =   "administrador"
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
            Locked          =   -1  'True
         EndProperty
         BeginProperty Column01 
            Locked          =   -1  'True
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column02 
            Locked          =   -1  'True
            Object.Visible         =   0   'False
         EndProperty
      EndProperty
   End
   Begin VB.TextBox text6 
      Alignment       =   2  'Center
      BackColor       =   &H80000016&
      DataField       =   "administrador"
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
      ForeColor       =   &H00000080&
      Height          =   285
      Left            =   3360
      TabIndex        =   2
      Top             =   1080
      Width           =   495
   End
   Begin VB.TextBox text2 
      Alignment       =   2  'Center
      DataField       =   "password"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Left            =   2760
      TabIndex        =   1
      Top             =   720
      Width           =   1815
   End
   Begin VB.TextBox text1 
      Alignment       =   2  'Center
      DataField       =   "nombre"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Left            =   2760
      TabIndex        =   0
      Top             =   360
      Width           =   1815
   End
   Begin VB.Frame Frame2 
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
      Height          =   1095
      Left            =   240
      TabIndex        =   58
      Top             =   7440
      Width           =   6375
      Begin KewlButtonz.KewlButtons grabar 
         Height          =   615
         Left            =   240
         TabIndex        =   52
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
         MICON           =   "frmusuarios1.frx":0491
         PICN            =   "frmusuarios1.frx":04AD
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
         Height          =   615
         Left            =   1440
         TabIndex        =   53
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
         MICON           =   "frmusuarios1.frx":1F2F
         PICN            =   "frmusuarios1.frx":1F4B
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
         Left            =   2640
         TabIndex        =   54
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
         MICON           =   "frmusuarios1.frx":533D
         PICN            =   "frmusuarios1.frx":5359
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
         Left            =   3840
         TabIndex        =   55
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
         MICON           =   "frmusuarios1.frx":5D6B
         PICN            =   "frmusuarios1.frx":5D87
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
         Left            =   5040
         TabIndex        =   56
         Top             =   360
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
         MICON           =   "frmusuarios1.frx":9179
         PICN            =   "frmusuarios1.frx":9195
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
      Height          =   1335
      Left            =   2640
      TabIndex        =   59
      Top             =   120
      Width           =   2295
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   5775
      Left            =   120
      ScaleHeight     =   5775
      ScaleWidth      =   7815
      TabIndex        =   62
      Top             =   1560
      Width           =   7815
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         BackColor       =   &H80000016&
         DataField       =   "campo11"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   47
         Left            =   6840
         TabIndex        =   49
         Top             =   1800
         Width           =   735
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         BackColor       =   &H80000016&
         DataField       =   "campo10"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   46
         Left            =   6840
         TabIndex        =   48
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         BackColor       =   &H80000016&
         DataField       =   "campo9"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   45
         Left            =   6840
         TabIndex        =   47
         Top             =   1320
         Width           =   735
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         BackColor       =   &H80000016&
         DataField       =   "campo8"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   44
         Left            =   6840
         TabIndex        =   51
         Top             =   3000
         Width           =   735
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         BackColor       =   &H80000016&
         DataField       =   "campo6"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   43
         Left            =   6840
         TabIndex        =   46
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         BackColor       =   &H80000016&
         DataField       =   "plancuentasing"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   0
         Left            =   1080
         TabIndex        =   3
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         BackColor       =   &H80000016&
         DataField       =   "plancuentasaltas"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   1
         Left            =   1080
         TabIndex        =   4
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         BackColor       =   &H80000016&
         DataField       =   "plancuentasmodi"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   2
         Left            =   1080
         TabIndex        =   63
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         BackColor       =   &H80000016&
         DataField       =   "plancuentasbajas"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   3
         Left            =   1080
         TabIndex        =   5
         Top             =   960
         Width           =   735
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         BackColor       =   &H80000016&
         DataField       =   "importaplaning"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   4
         Left            =   1080
         TabIndex        =   6
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         BackColor       =   &H80000016&
         DataField       =   "livacomprasing"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   5
         Left            =   3000
         TabIndex        =   15
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         BackColor       =   &H80000016&
         DataField       =   "livacomprasmodi"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   6
         Left            =   3000
         TabIndex        =   16
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         BackColor       =   &H80000016&
         DataField       =   "livacomprascerrar"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   7
         Left            =   3000
         TabIndex        =   17
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         BackColor       =   &H80000016&
         DataField       =   "livacompraslistar"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   8
         Left            =   3000
         TabIndex        =   18
         Top             =   960
         Width           =   735
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         BackColor       =   &H80000016&
         DataField       =   "livaventasing"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   9
         Left            =   4920
         TabIndex        =   29
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         BackColor       =   &H80000016&
         DataField       =   "livaventasmodi"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   10
         Left            =   4920
         TabIndex        =   30
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         BackColor       =   &H80000016&
         DataField       =   "livaventascerrar"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   11
         Left            =   4920
         TabIndex        =   31
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         BackColor       =   &H80000016&
         DataField       =   "livaventaslistar"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   12
         Left            =   4920
         TabIndex        =   32
         Top             =   960
         Width           =   735
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         BackColor       =   &H80000016&
         DataField       =   "minutasing"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   13
         Left            =   3000
         TabIndex        =   19
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         BackColor       =   &H80000016&
         DataField       =   "minutasaltas"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   14
         Left            =   3000
         TabIndex        =   20
         Top             =   1800
         Width           =   735
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         BackColor       =   &H80000016&
         DataField       =   "minutasmodi"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   15
         Left            =   3000
         TabIndex        =   21
         Top             =   2040
         Width           =   735
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         BackColor       =   &H80000016&
         DataField       =   "minutasbajas"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   16
         Left            =   3000
         TabIndex        =   22
         Top             =   2280
         Width           =   735
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         BackColor       =   &H80000016&
         DataField       =   "empresasing"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   17
         Left            =   4920
         TabIndex        =   33
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         BackColor       =   &H80000016&
         DataField       =   "empresasmodi"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   18
         Left            =   4920
         TabIndex        =   34
         Top             =   1800
         Width           =   735
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         BackColor       =   &H80000016&
         DataField       =   "empresasbajas"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   19
         Left            =   4920
         TabIndex        =   35
         Top             =   2040
         Width           =   735
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         BackColor       =   &H80000016&
         DataField       =   "cambpertrabajoing"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   20
         Left            =   1080
         TabIndex        =   7
         Top             =   2280
         Width           =   735
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         BackColor       =   &H80000016&
         DataField       =   "proving"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   21
         Left            =   1080
         TabIndex        =   8
         Top             =   2880
         Width           =   735
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         BackColor       =   &H80000016&
         DataField       =   "provaltas"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   22
         Left            =   1080
         TabIndex        =   9
         Top             =   3120
         Width           =   735
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         BackColor       =   &H80000016&
         DataField       =   "provmodi"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   23
         Left            =   1080
         TabIndex        =   10
         Top             =   3360
         Width           =   735
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         BackColor       =   &H80000016&
         DataField       =   "provbajas"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   24
         Left            =   1080
         TabIndex        =   11
         Top             =   3600
         Width           =   735
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         BackColor       =   &H80000016&
         DataField       =   "estadocuentaing"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   25
         Left            =   1080
         TabIndex        =   12
         Top             =   3840
         Width           =   735
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         BackColor       =   &H80000016&
         DataField       =   "ordpagoing"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   26
         Left            =   3000
         TabIndex        =   23
         Top             =   2880
         Width           =   735
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         BackColor       =   &H80000016&
         DataField       =   "ordpagocons"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   27
         Left            =   3000
         TabIndex        =   24
         Top             =   3120
         Width           =   735
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         BackColor       =   &H80000016&
         DataField       =   "ordpagolistado"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   28
         Left            =   3000
         TabIndex        =   25
         Top             =   3360
         Width           =   735
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         BackColor       =   &H80000016&
         DataField       =   "clientesing"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   29
         Left            =   4920
         TabIndex        =   36
         Top             =   2880
         Width           =   735
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         BackColor       =   &H80000016&
         DataField       =   "clientesaltas"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   30
         Left            =   4920
         TabIndex        =   37
         Top             =   3120
         Width           =   735
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         BackColor       =   &H80000016&
         DataField       =   "clientesmodi"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   31
         Left            =   4920
         TabIndex        =   38
         Top             =   3360
         Width           =   735
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         BackColor       =   &H80000016&
         DataField       =   "clientesbajas"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   32
         Left            =   4920
         TabIndex        =   39
         Top             =   3600
         Width           =   735
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         BackColor       =   &H80000016&
         DataField       =   "emitefactura"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   33
         Left            =   3000
         TabIndex        =   26
         Top             =   3960
         Width           =   735
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         BackColor       =   &H80000016&
         DataField       =   "articulosing"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   34
         Left            =   3000
         TabIndex        =   27
         Top             =   4200
         Width           =   735
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         BackColor       =   &H80000016&
         DataField       =   "reportesing"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   35
         Left            =   1080
         TabIndex        =   14
         Top             =   4800
         Width           =   735
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         BackColor       =   &H80000016&
         DataField       =   "parametrosing"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   36
         Left            =   4920
         TabIndex        =   42
         Top             =   4800
         Width           =   735
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         BackColor       =   &H80000016&
         DataField       =   "campo1"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   37
         Left            =   3000
         TabIndex        =   28
         Top             =   4800
         Width           =   735
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         BackColor       =   &H80000016&
         DataField       =   "campo2"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   38
         Left            =   4920
         TabIndex        =   43
         Top             =   5400
         Width           =   735
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         BackColor       =   &H80000016&
         DataField       =   "campo3"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   39
         Left            =   6840
         TabIndex        =   50
         Top             =   2400
         Width           =   735
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         BackColor       =   &H80000016&
         DataField       =   "campo4"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   40
         Left            =   1080
         TabIndex        =   13
         Top             =   4080
         Width           =   735
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         BackColor       =   &H80000016&
         DataField       =   "campo7"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   42
         Left            =   6840
         TabIndex        =   45
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         BackColor       =   &H80000016&
         DataField       =   "campo5"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   41
         Left            =   6840
         TabIndex        =   44
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         BackColor       =   &H80000016&
         DataField       =   "campo12"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   48
         Left            =   4920
         TabIndex        =   40
         Top             =   3840
         Width           =   735
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         BackColor       =   &H80000016&
         DataField       =   "campo13"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   49
         Left            =   4920
         TabIndex        =   41
         Top             =   4080
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Ajuste E.C."
         Height          =   255
         Index           =   49
         Left            =   3960
         TabIndex        =   135
         Top             =   4080
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "E.C."
         Height          =   255
         Index           =   48
         Left            =   3960
         TabIndex        =   134
         Top             =   3840
         Width           =   615
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Recibos"
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
         Index           =   5
         Left            =   5760
         TabIndex        =   130
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Emitir"
         Height          =   255
         Index           =   47
         Left            =   5880
         TabIndex        =   133
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Cons./List"
         Height          =   255
         Index           =   46
         Left            =   5880
         TabIndex        =   132
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Anular"
         Height          =   255
         Index           =   45
         Left            =   5880
         TabIndex        =   131
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Mov.de Valores"
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
         Index           =   1
         Left            =   5760
         TabIndex        =   128
         Top             =   2760
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Ingresar"
         Height          =   255
         Index           =   44
         Left            =   5880
         TabIndex        =   129
         Top             =   3000
         Width           =   615
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Libo Caja/Banco"
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
         Index           =   0
         Left            =   5760
         TabIndex        =   124
         Top             =   0
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Altas/mod"
         Height          =   255
         Index           =   43
         Left            =   5880
         TabIndex        =   127
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Listar"
         Height          =   255
         Index           =   42
         Left            =   5880
         TabIndex        =   126
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Ingresar"
         Height          =   255
         Index           =   41
         Left            =   5880
         TabIndex        =   125
         Top             =   240
         Width           =   855
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Libro IVA Ventas"
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
         Index           =   6
         Left            =   3840
         TabIndex        =   91
         Top             =   0
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Datos de Empresas"
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
         Index           =   8
         Left            =   3840
         TabIndex        =   89
         Top             =   1320
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cambio Per.de Trab."
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
         Index           =   9
         Left            =   0
         TabIndex        =   88
         Top             =   2040
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Proveedores"
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
         Index           =   10
         Left            =   0
         TabIndex        =   87
         Top             =   2640
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ordenes de Pago"
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
         Index           =   12
         Left            =   1920
         TabIndex        =   86
         Top             =   2640
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Clientes"
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
         Index           =   13
         Left            =   3840
         TabIndex        =   85
         Top             =   2640
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Facturacion"
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
         Index           =   14
         Left            =   1920
         TabIndex        =   84
         Top             =   3720
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Parametros Conf."
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
         Left            =   3840
         TabIndex        =   83
         Top             =   4560
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Reportes"
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
         Left            =   0
         TabIndex        =   67
         Top             =   4560
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Importar Lib.Ventas"
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
         Left            =   1920
         TabIndex        =   66
         Top             =   4560
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ordenes Publ."
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
         Left            =   3840
         TabIndex        =   65
         Top             =   5160
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Libros Cerrados"
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
         Left            =   5760
         TabIndex        =   64
         Top             =   2160
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Plan de Cuentas"
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
         Index           =   2
         Left            =   0
         TabIndex        =   94
         Top             =   0
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Libro IVA Compras"
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
         Index           =   4
         Left            =   1920
         TabIndex        =   92
         Top             =   0
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Minutas Contables"
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
         Index           =   7
         Left            =   1920
         TabIndex        =   90
         Top             =   1320
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Import.Plan y Conf."
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
         Index           =   3
         Left            =   0
         TabIndex        =   93
         Top             =   1320
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Ajuste E.C."
         Height          =   255
         Index           =   39
         Left            =   120
         TabIndex        =   120
         Top             =   4080
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Modificar:"
         Height          =   255
         Index           =   40
         Left            =   5880
         TabIndex        =   119
         Top             =   2400
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Ingresar"
         Height          =   255
         Index           =   38
         Left            =   3960
         TabIndex        =   118
         Top             =   5400
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Ingresar"
         Height          =   255
         Index           =   37
         Left            =   2040
         TabIndex        =   117
         Top             =   4800
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Ingresar"
         Height          =   255
         Index           =   36
         Left            =   120
         TabIndex        =   116
         Top             =   2280
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Ingresar"
         Height          =   255
         Index           =   35
         Left            =   120
         TabIndex        =   115
         Top             =   4800
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Emitir"
         Height          =   255
         Index           =   33
         Left            =   1920
         TabIndex        =   114
         Top             =   3960
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Ingresar"
         Height          =   255
         Index           =   32
         Left            =   3960
         TabIndex        =   113
         Top             =   4800
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "E.C."
         Height          =   255
         Index           =   31
         Left            =   120
         TabIndex        =   112
         Top             =   3840
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Ingresar"
         Height          =   255
         Index           =   27
         Left            =   3960
         TabIndex        =   111
         Top             =   2880
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Emitir"
         Height          =   255
         Index           =   24
         Left            =   2040
         TabIndex        =   110
         Top             =   2880
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Cons/List"
         Height          =   255
         Index           =   25
         Left            =   2040
         TabIndex        =   109
         Top             =   3120
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Anular"
         Height          =   255
         Index           =   26
         Left            =   2040
         TabIndex        =   108
         Top             =   3360
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Ingresar"
         Height          =   255
         Index           =   20
         Left            =   120
         TabIndex        =   107
         Top             =   2880
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Altas"
         Height          =   255
         Index           =   21
         Left            =   120
         TabIndex        =   106
         Top             =   3120
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Modific."
         Height          =   255
         Index           =   22
         Left            =   120
         TabIndex        =   105
         Top             =   3360
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Eliminar"
         Height          =   255
         Index           =   23
         Left            =   120
         TabIndex        =   104
         Top             =   3600
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Ingresar"
         Height          =   255
         Index           =   17
         Left            =   3960
         TabIndex        =   103
         Top             =   1560
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Ingresar"
         Height          =   255
         Index           =   13
         Left            =   2040
         TabIndex        =   102
         Top             =   1560
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Altas"
         Height          =   255
         Index           =   14
         Left            =   2040
         TabIndex        =   101
         Top             =   1800
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Modific."
         Height          =   255
         Index           =   15
         Left            =   2040
         TabIndex        =   100
         Top             =   2040
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Eliminar"
         Height          =   255
         Index           =   16
         Left            =   2040
         TabIndex        =   99
         Top             =   2280
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Ingresar"
         Height          =   255
         Index           =   12
         Left            =   120
         TabIndex        =   98
         Top             =   1560
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Ingresar"
         Height          =   255
         Index           =   8
         Left            =   3960
         TabIndex        =   97
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Ingresar"
         Height          =   255
         Index           =   4
         Left            =   2040
         TabIndex        =   96
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Ingresar"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   95
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Altas"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   82
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Modific."
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   81
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Borrar Cue."
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   80
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Altas/mod"
         Height          =   255
         Index           =   5
         Left            =   2040
         TabIndex        =   79
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Cerrar"
         Height          =   255
         Index           =   6
         Left            =   2040
         TabIndex        =   78
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Listar"
         Height          =   255
         Index           =   7
         Left            =   2040
         TabIndex        =   77
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Altas/mod"
         Height          =   255
         Index           =   9
         Left            =   3960
         TabIndex        =   76
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Cerrar"
         Height          =   255
         Index           =   10
         Left            =   3960
         TabIndex        =   75
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Listar"
         Height          =   255
         Index           =   11
         Left            =   3960
         TabIndex        =   74
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Altas/Modi"
         Height          =   255
         Index           =   18
         Left            =   3960
         TabIndex        =   73
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Eliminar"
         Height          =   255
         Index           =   19
         Left            =   3960
         TabIndex        =   72
         Top             =   2040
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Altas"
         Height          =   255
         Index           =   28
         Left            =   3960
         TabIndex        =   71
         Top             =   3120
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Modific."
         Height          =   255
         Index           =   29
         Left            =   3960
         TabIndex        =   70
         Top             =   3360
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Eliminar"
         Height          =   255
         Index           =   30
         Left            =   3960
         TabIndex        =   69
         Top             =   3600
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Mod.Articulos"
         Height          =   255
         Index           =   34
         Left            =   1920
         TabIndex        =   68
         Top             =   4200
         Width           =   1095
      End
   End
   Begin MSAdodcLib.Adodc datvendedores 
      Height          =   330
      Left            =   6720
      Top             =   7920
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
End
Attribute VB_Name = "frmusuarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub borrar_Click()
On Error GoTo errorborrado


 KeyAscii = 13
  Respuesta = MsgBox("ESTA POR BORRAR UN USUARIO, ESTA SEGURO?", vbYesNo, "Atención")
If Respuesta = vbYes Then
    datPrimaryRS.Recordset.Delete
Else
    Exit Sub
End If
Exit Sub

errorborrado:
Respuesta = MsgBox("No se pudo borrar el registro por contener permisos, limpie los permisos e intente de nuevo", vbCritical, "Atención")

     
End Sub

Private Sub eliminarempresa_Click()

KeyAscii = 13
  Respuesta = MsgBox("ESTA POR BORRAR UN PERMISO A EMPRESA?", vbYesNo, "Atención")
If Respuesta = vbYes Then
    datPrimaryRS.Recordset.Delete
Else
    Exit Sub
End If
   

End Sub



Private Sub DataGrid2_ButtonClick(ByVal ColIndex As Integer)

    DataList1.Visible = True
    DataList1.SetFocus

End Sub

Private Sub DataList1_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        If DataGrid2.ApproxCount = 0 Then datusyemp.Recordset.AddNew
        DataGrid2.Columns(1).Text = DataList1.BoundText
        DataGrid2.Columns(2).Text = DataList1.Text
        DataGrid2.Columns(0).Text = Text1.Text
        datusyemp.Recordset.UpdateBatch adAffectCurrent
        DataGrid2.SetFocus
    End If

End Sub

Private Sub DataList1_LostFocus()

    DataList1.Visible = False

End Sub

Private Sub datusuarios_WillMove(ByVal adReason As ADODB.EventReasonEnum, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)

End Sub

Private Sub Form_Load()
Aplicar_skin Me

datusuarios.ConnectionString = login.conexiontotal
datvendedores.ConnectionString = login.conexiontotal

datusuarios.RecordSource = "select * from ud_ezi_empleado order by apynomb"
datusuarios.Refresh

datvendedores.RecordSource = "SELECT    V_VENDEDOR_.ID, V_VENDEDOR_.CODIGO, V_PERSONA_.NOMBRE, V_VENDEDOR_.CODIGO  + ' ' + V_PERSONA_.NOMBRE as cancatena " & _
                              "FROM V_VENDEDOR_ INNER JOIN V_PERSONA_ ON V_VENDEDOR_.ENTEASOCIADO_ID = V_PERSONA_.ID " & _
                              "WHERE (V_VENDEDOR_.ACTIVESTATUS = 0)  order by V_PERSONA_.NOMBRE"
datvendedores.Refresh



End Sub

Private Sub grabar_Click()

    
End Sub

Private Sub nuevo_Click()
    
    
End Sub

Private Sub salir_Click()
  Unload Me
End Sub

Private Sub Text1_Change()
If Text6.Text = "S" Then
    Check4 = 1
Else
    Check4 = 0
End If

usuario1 = Text1.Text
datusyemp.RecordSource = "select usyempresas.* from usyempresas where nomusuario = '" & usuario1 & "'"
datusyemp.Refresh

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Text2.SetFocus
    End If
End Sub

Private Sub text2_Change()

 Text2.PasswordChar = "*"
 
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Text6.SetFocus
    End If
End Sub

Private Sub Text3_GotFocus(Index As Integer)

    Text3(Index).Text = UCase(Text3(Index).Text)
    Text3(Index).SelStart = 0
    Text3(Index).SelLength = 1

End Sub

Private Sub Text3_KeyPress(Index As Integer, KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = 13 Then
        KeyAscii = 0
        If Text3(Index).Text <> "s" And Text3(Index).Text <> "S" Then
            Text3(Index).Text = "N"
        Else
            Text3(Index).Text = "S"
        End If
        SendKeys "{TAB}", False
    End If
        
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
        KeyAscii = 0
        If Text6.Text <> "s" And Text6.Text <> "S" Then
            Text6.Text = "N"
        Else
            Text6.Text = "S"
        End If

        For X = 0 To 49
            If Text6.Text = "S" Then
                Text3(X).Text = "S"
            Else
                Text3(X).Text = "N"
            End If
        Next X
       
        Text3(0).SetFocus
    End If
End Sub
