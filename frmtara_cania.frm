VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{D18BBD1F-82BB-4385-BED3-E9D31A3E361E}#1.0#0"; "KewlButtonz.ocx"
Begin VB.Form frmtara_cania 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TARA DE CAÑA"
   ClientHeight    =   7965
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12915
   Icon            =   "frmtara_cania.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7965
   ScaleWidth      =   12915
   Begin VB.Frame Frame4 
      Height          =   4815
      Left            =   9840
      TabIndex        =   40
      Top             =   0
      Width           =   2895
      Begin VB.TextBox Text4 
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
         Height          =   495
         Left            =   600
         TabIndex        =   17
         Top             =   1560
         Width           =   1455
      End
      Begin VB.TextBox Text5 
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
         Height          =   495
         Left            =   600
         TabIndex        =   18
         Top             =   2760
         Width           =   1455
      End
      Begin VB.PictureBox Picture2 
         Height          =   375
         Left            =   360
         ScaleHeight     =   315
         ScaleWidth      =   1995
         TabIndex        =   45
         Top             =   1080
         Width           =   2055
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Usuario"
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
            Index           =   13
            Left            =   480
            TabIndex        =   46
            Top             =   0
            Width           =   1095
         End
      End
      Begin VB.PictureBox Picture3 
         Height          =   375
         Left            =   360
         ScaleHeight     =   315
         ScaleWidth      =   1995
         TabIndex        =   43
         Top             =   2280
         Width           =   2055
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Password"
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
            Left            =   360
            TabIndex        =   44
            Top             =   0
            Width           =   1335
         End
      End
      Begin VB.PictureBox Picture4 
         Height          =   855
         Left            =   120
         ScaleHeight     =   795
         ScaleWidth      =   2595
         TabIndex        =   41
         Top             =   120
         Width           =   2655
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Ingreso usuario y password válido para autorizar Tara Manual"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   14
            Left            =   0
            TabIndex        =   42
            Top             =   0
            Width           =   2535
         End
      End
      Begin MSAdodcLib.Adodc datusuarios 
         Height          =   330
         Left            =   720
         Top             =   4320
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
      Begin KewlButtonz.KewlButtons autorizar 
         Height          =   615
         Left            =   600
         TabIndex        =   19
         Top             =   3600
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   1085
         BTYPE           =   14
         TX              =   "Autorizar"
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
         MICON           =   "frmtara_cania.frx":0442
         PICN            =   "frmtara_cania.frx":045E
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
   Begin VB.Frame Frame2 
      Caption         =   "Nro de Pesada"
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
      Height          =   1215
      Left            =   5640
      TabIndex        =   33
      Top             =   0
      Width           =   4095
      Begin VB.TextBox Text2 
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
         Height          =   495
         Left            =   240
         TabIndex        =   0
         Top             =   360
         Width           =   1455
      End
      Begin KewlButtonz.KewlButtons buscar 
         Height          =   615
         Left            =   2520
         TabIndex        =   1
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   1085
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
         MICON           =   "frmtara_cania.frx":08B0
         PICN            =   "frmtara_cania.frx":08CC
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
      Caption         =   "Modo de Pesada"
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
      Height          =   1215
      Left            =   120
      TabIndex        =   30
      Top             =   0
      Width           =   5295
      Begin KewlButtonz.KewlButtons automatica 
         Height          =   615
         Left            =   360
         TabIndex        =   31
         Top             =   360
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   1085
         BTYPE           =   14
         TX              =   "&Automatica"
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
         MICON           =   "frmtara_cania.frx":0E5E
         PICN            =   "frmtara_cania.frx":0E7A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin KewlButtonz.KewlButtons manual 
         Height          =   615
         Left            =   2880
         TabIndex        =   32
         Top             =   360
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   1085
         BTYPE           =   14
         TX              =   "&Manual"
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
         MICON           =   "frmtara_cania.frx":1414
         PICN            =   "frmtara_cania.frx":1430
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
   Begin MSAdodcLib.Adodc dattipocaña 
      Height          =   330
      Left            =   8400
      Top             =   3480
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
   Begin MSAdodcLib.Adodc datcanieros 
      Height          =   330
      Left            =   8400
      Top             =   3840
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
      Left            =   8400
      Top             =   4200
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
   Begin MSAdodcLib.Adodc dattransporte 
      Height          =   330
      Left            =   8400
      Top             =   4560
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
      Height          =   5175
      Left            =   8160
      TabIndex        =   20
      Top             =   1200
      Width           =   1575
      Begin KewlButtonz.KewlButtons salir 
         Height          =   615
         Left            =   240
         TabIndex        =   16
         Top             =   4200
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
         BCOL            =   -2147483645
         BCOLO           =   -2147483645
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmtara_cania.frx":1D0A
         PICN            =   "frmtara_cania.frx":1D26
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
         Height          =   615
         Left            =   240
         TabIndex        =   15
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
         BCOL            =   -2147483645
         BCOLO           =   -2147483645
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmtara_cania.frx":2870
         PICN            =   "frmtara_cania.frx":288C
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
         Left            =   240
         TabIndex        =   21
         Top             =   1320
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
         BCOL            =   -2147483645
         BCOLO           =   -2147483645
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmtara_cania.frx":430E
         PICN            =   "frmtara_cania.frx":432A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSAdodcLib.Adodc datpesada 
         Height          =   330
         Left            =   240
         Top             =   3720
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
   Begin VB.PictureBox Picture1 
      Height          =   6495
      Left            =   120
      ScaleHeight     =   6435
      ScaleWidth      =   9555
      TabIndex        =   22
      Top             =   1320
      Width           =   9615
      Begin VB.TextBox Text3 
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
         Height          =   465
         Index           =   5
         Left            =   8040
         TabIndex        =   14
         Top             =   5760
         Width           =   1455
      End
      Begin VB.TextBox Text3 
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
         Height          =   465
         Index           =   4
         Left            =   8040
         TabIndex        =   13
         Top             =   5160
         Width           =   1455
      End
      Begin VB.TextBox Text3 
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
         Height          =   465
         Index           =   3
         Left            =   4920
         TabIndex        =   12
         Top             =   5760
         Width           =   1455
      End
      Begin VB.TextBox Text3 
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
         Height          =   465
         Index           =   2
         Left            =   4920
         TabIndex        =   11
         Top             =   5160
         Width           =   1455
      End
      Begin VB.TextBox Text3 
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
         Height          =   465
         Index           =   1
         Left            =   1920
         TabIndex        =   10
         Top             =   5760
         Width           =   1455
      End
      Begin VB.TextBox Text3 
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
         Height          =   465
         Index           =   0
         Left            =   1920
         TabIndex        =   9
         Top             =   5160
         Width           =   1455
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
         Height          =   405
         Index           =   2
         Left            =   2160
         TabIndex        =   7
         Top             =   3240
         Width           =   3615
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
         Height          =   405
         Index           =   1
         Left            =   2160
         MaxLength       =   50
         TabIndex        =   6
         Top             =   2640
         Width           =   2535
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
         Height          =   1005
         Index           =   3
         Left            =   2160
         MaxLength       =   300
         MultiLine       =   -1  'True
         TabIndex        =   8
         Top             =   3840
         Width           =   5175
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "frmtara_cania.frx":4D3C
         Height          =   360
         Left            =   2160
         TabIndex        =   2
         Top             =   240
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   741
         _Version        =   393216
         Enabled         =   0   'False
         Style           =   2
         ListField       =   "nombre"
         BoundColumn     =   "alias_0_id"
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
         Height          =   405
         Index           =   0
         Left            =   2160
         MaxLength       =   50
         TabIndex        =   5
         Top             =   2040
         Width           =   5295
      End
      Begin MSDataListLib.DataCombo DataCombo2 
         Bindings        =   "frmtara_cania.frx":4D56
         Height          =   360
         Left            =   2160
         TabIndex        =   3
         Top             =   720
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   741
         _Version        =   393216
         Enabled         =   0   'False
         Style           =   2
         ListField       =   "alias_3_nombre"
         BoundColumn     =   "alias_0_id"
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
      Begin MSDataListLib.DataCombo DataCombo3 
         Bindings        =   "frmtara_cania.frx":4D70
         Height          =   360
         Left            =   2160
         TabIndex        =   4
         Top             =   1200
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   741
         _Version        =   393216
         Enabled         =   0   'False
         Style           =   2
         ListField       =   "alias_1_nombre"
         BoundColumn     =   "alias_0_id"
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
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Trash:"
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
         Index           =   12
         Left            =   6480
         TabIndex        =   39
         Top             =   5760
         Width           =   1455
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Neto Caña:"
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
         Index           =   11
         Left            =   6480
         TabIndex        =   38
         Top             =   5160
         Width           =   1455
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Tara:"
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
         Index           =   10
         Left            =   3360
         TabIndex        =   37
         Top             =   5760
         Width           =   1455
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Peso Neto:"
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
         Left            =   3360
         TabIndex        =   36
         Top             =   5160
         Width           =   1455
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "N° de Pesada:"
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
         Left            =   0
         TabIndex        =   35
         Top             =   5760
         Width           =   1815
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Peso Bruto:"
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
         Left            =   0
         TabIndex        =   34
         Top             =   5160
         Width           =   1815
      End
      Begin VB.Line Line2 
         X1              =   120
         X2              =   7920
         Y1              =   5040
         Y2              =   5040
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Observaciones:"
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
         Left            =   120
         TabIndex        =   29
         Top             =   3840
         Width           =   1935
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Nro. Remito:"
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
         Index           =   1
         Left            =   120
         TabIndex        =   28
         Top             =   3240
         Width           =   1935
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Patente:"
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
         Left            =   120
         TabIndex        =   27
         Top             =   2640
         Width           =   1935
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Chofer:"
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
         Left            =   120
         TabIndex        =   26
         Top             =   2040
         Width           =   1935
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   7560
         Y1              =   1800
         Y2              =   1800
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo de Caña:"
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
         TabIndex        =   25
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Cañero"
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
         Left            =   120
         TabIndex        =   24
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Transportista:"
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
         Left            =   120
         TabIndex        =   23
         Top             =   1200
         Width           =   2055
      End
   End
End
Attribute VB_Name = "frmtara_cania"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public modo As String
Public usuariomanual As String


Private Sub automatica_Click()

    frmtara_cania.Caption = "TARA DE CAÑA         Modo Automático Activo"
    modo = "A" 'Automatico
    Text2.SetFocus
    Text3(2).Enabled = False
    Text3(3).Enabled = False
    Text3(4).Enabled = False
    
If Text2.Text <> "" Then
    If Text3(3).Text <> "" Then
        mensa = MsgBox("Suma al Peso de Tara Registrado ?", vbYesNo, "!! Atención !!")
        If mensa = vbYes Then
            xpesoanterior = Text3(3).Text
        Else
            xpesoanterior = 0
        End If
    Else
        xpesoanterior = 0
    End If


    Text3(3).Text = Int(5000 + (Rnd(10) * 1000)) + xpesoanterior
    Text3(2).Text = Int(Text3(0).Text - Text3(3).Text)
    Text3(4).Text = Round(Text3(2).Text - (Text3(2).Text * Val(Text3(5).Text) / 100), 0)

End If
    
    

End Sub

Private Sub autorizar_Click()



        datusuarios.RecordSource = "select nombre,direccionelectronica from v_usuario_ WHERE nombre = '" & Text4.Text & "'"
        datusuarios.Refresh
        
        If datusuarios.Recordset.EOF = True Then
            mensa = MsgBox("Este usuario no existe", vbCritical, "!! Error !!")
            Text5.Text = ""
            Text4.Text = ""
            Text4.SetFocus
            Exit Sub
        End If
    
        If datusuarios.Recordset.EOF = False Then
            xclave = datusuarios.Recordset.Fields("DIRECCIONELECTRONICA")
            If Text5.Text <> xclave Or Text5.Text = "" Then
                mensa = MsgBox("Password incorrecto", vbCritical, "!! Error !!")
                Text5.Text = ""
                Text5.SetFocus
                Exit Sub
            End If
            
            usuariomanual = Text4.Text
            Text4.Text = ""
            Text5.Text = ""

            frmtara_cania.Caption = "TARA DE CAÑA         Modo Manual Activo"
            modo = "M" 'Manual
            Text3(3).Enabled = True

            
            frmtara_cania.Width = 10020
            If Text2.Text <> "" Then
                Text3(3).SetFocus
            Else
                 Text2.SetFocus
            End If
        End If


End Sub

Private Sub buscar_Click()

    lista_pesadacania.Show

End Sub

Private Sub Cancelar_Click()

    Unload Me
    frmtara_cania.Show

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

Private Sub Form_Activate()


  Call automatica_Click


End Sub

Private Sub Form_Load()
Aplicar_skin Me

frmtara_cania.Top = 0
frmtara_cania.Left = 0
frmtara_cania.Width = 10020
usuariomanual = ""

dattipocaña.ConnectionString = login.conexiontotal
datcanieros.ConnectionString = login.conexiontotal
dattransporte.ConnectionString = login.conexiontotal
datmovimientos.ConnectionString = login.conexiontotal
datpesada.ConnectionString = login.conexiontotal
datusuarios.ConnectionString = login.conexiontotal

    dattipocaña.RecordSource = "SELECT ID AS ALIAS_0_ID, NOMBRE FROM V_ITEMTIPOCLASIFICADOR_ AS ALIAS_0 " & _
                               "WHERE (BO_PLACE_ID = '{8CCBA4D1-EDDE-432A-B63E-C8AC0AC3DE2F}') AND (ACTIVESTATUS <> 2) ORDER BY NOMBRE"
    dattipocaña.Refresh

    datcanieros.RecordSource = "SELECT     ALIAS_0.ID AS ALIAS_0_ID, ALIAS_0.CODIGO AS ALIAS_0_CODIGO, ALIAS_3.NOMBRE AS ALIAS_3_NOMBRE, V_UD_CLIENTE_.PRODUCTOR " & _
                               "FROM V_CLIENTE_ AS ALIAS_0 LEFT OUTER JOIN V_UD_CLIENTE_ ON ALIAS_0.BOEXTENSION_ID = V_UD_CLIENTE_.ID LEFT OUTER JOIN V_PERSONA_ AS ALIAS_3 ON ALIAS_0.ENTEASOCIADO_ID = ALIAS_3.ID " & _
                               "WHERE (ALIAS_0.BO_PLACE_ID = '{89C234C8-3F01-11D5-86AD-0080AD403F5F}') AND (ALIAS_0.ACTIVESTATUS = 0) AND (V_UD_CLIENTE_.PRODUCTOR = 'T') ORDER BY ALIAS_3.NOMBRE "
    datcanieros.Refresh

    dattransporte.RecordSource = "SELECT     ALIAS_0.ID AS ALIAS_0_ID, ALIAS_1.NOMBRE AS ALIAS_1_NOMBRE FROM V_MEDIOTRANSPORTE_ AS ALIAS_0 LEFT OUTER JOIN " & _
                                 "V_PERSONA_ AS ALIAS_1 ON ALIAS_0.ENTEASOCIADO_ID = ALIAS_1.ID WHERE     (ALIAS_0.BO_PLACE_ID = '{76C697C2-3DAE-11D5-B059-004854841C8A}') AND (ALIAS_0.ACTIVESTATUS = 0) " & _
                                 "ORDER BY ALIAS_1_NOMBRE"
    dattransporte.Refresh
    
    
    
   
End Sub

Private Sub grabar_Click()
On Error GoTo errorgrabar

    datmovimientos.Recordset.Fields("id_tipo_cana") = DataCombo1.BoundText
    datmovimientos.Recordset.Fields("id_caniero") = DataCombo2.BoundText
    datmovimientos.Recordset.Fields("id_transportista") = DataCombo3.BoundText
    datmovimientos.Recordset.Fields("razon_social") = DataCombo2.Text
    datmovimientos.Recordset.Fields("transporte") = DataCombo3.Text
    datmovimientos.Recordset.Fields("chofer") = Text1(0).Text
    datmovimientos.Recordset.Fields("patente") = Text1(1).Text
    datmovimientos.Recordset.Fields("observaciones") = Text1(3).Text
    datmovimientos.Recordset.Fields("prepesada") = "F"
    datmovimientos.Recordset.Fields("fecha_salida") = Str(Date)
    datmovimientos.Recordset.Fields("hora_salida") = Str(Time)
    datmovimientos.Recordset.Fields("tipo_pesada") = "C"
    datmovimientos.Recordset.Fields("remito") = Text1(2).Text
    
    datmovimientos.Recordset.Fields("peso_bruto") = Val(Text3(0).Text)
    datmovimientos.Recordset.Fields("peso_neto") = Val(Text3(2).Text)
    datmovimientos.Recordset.Fields("tara") = Val(Text3(3).Text)
    datmovimientos.Recordset.Fields("trash") = Val(Text3(5).Text)
    datmovimientos.Recordset.Fields("neto_cana") = Val(Text3(4).Text)
    datmovimientos.Recordset.Fields("tipotara") = modo

    datmovimientos.Recordset.Fields("usuario_salida") = login.usuarioactivo
    datmovimientos.Recordset.Fields("usuariomanual") = usuariomanual
    
    datmovimientos.Recordset.UpdateBatch adAffectCurrent
    mensa = MsgBox("Tara registrada correctamente", vbInformation, "Grabado Correctamente")
    
    Call Cancelar_Click
    
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
    
    For X = frmtara_cania.Width To 13000 Step 100
            frmtara_cania.Width = X
    Next X
    Text4.SetFocus



End Sub

Private Sub salir_Click()

    Unload Me

End Sub


Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
On Error Resume Next

    If KeyAscii = 13 Then
        KeyAscii = 0
        If Index = 3 Then
            grabar.SetFocus
        Else
            Text1(Index + 1).SetFocus
        End If
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

Private Sub Text2_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        datmovimientos.RecordSource = "select * from pr_ezi_movimientos where numero_pesada = " & Text2.Text & " and (prepesada = 'T') AND (NOT (numero_pesada IS NULL)) and (tipo_pesada = 'C')  "
        datmovimientos.Refresh
        
        If datmovimientos.Recordset.EOF = True Then
            DataCombo1.BoundText = ""
            DataCombo2.BoundText = ""
            DataCombo3.BoundText = ""
            Text1(0).Text = ""
            Text1(1).Text = ""
            Text1(2).Text = ""
            Text1(3).Text = ""
            Text3(0).Text = ""
            Text3(1).Text = ""
            Text3(2).Text = ""
            Text3(3).Text = ""
            Text3(4).Text = ""
            Text3(5).Text = ""
            Exit Sub
        End If
            
        
        DataCombo1.BoundText = datmovimientos.Recordset.Fields("id_tipo_cana")
        DataCombo2.BoundText = datmovimientos.Recordset.Fields("id_caniero")
        DataCombo3.BoundText = datmovimientos.Recordset.Fields("id_transportista")
        Text1(0).Text = datmovimientos.Recordset.Fields("chofer")
        Text1(1).Text = datmovimientos.Recordset.Fields("patente")
        Text1(2).Text = datmovimientos.Recordset.Fields("remito")
        Text1(3).Text = datmovimientos.Recordset.Fields("observaciones")
        Text3(0).Text = datmovimientos.Recordset.Fields("peso_bruto")
        Text3(1).Text = datmovimientos.Recordset.Fields("numero_pesada")
        Text3(5).Text = datmovimientos.Recordset.Fields("trash")
        
        
            
        
    End If


    

End Sub

Private Sub Text3_Change(Index As Integer)

If Not IsNumeric(Text3(Index).Text) Then
    Text3(Index).Text = ""
    mensa = MsgBox("Valor Numerico no Valido ", vbCritical, "Error !!")
End If


End Sub

Private Sub Text3_KeyPress(Index As Integer, KeyAscii As Integer)

On Error Resume Next

    If KeyAscii = 13 Then
        KeyAscii = 0
        If Index = 3 Or Index = 5 Then
            Text3(2).Text = Int(Text3(0).Text - Text3(3).Text)
            Text3(4).Text = Round(Text3(2).Text - (Text3(2).Text * Val(Text3(5).Text) / 100), 0)
            grabar.SetFocus
        End If
    End If

End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)

On Error Resume Next

    If KeyAscii = 13 Then
        KeyAscii = 0
        Text5.SetFocus
    End If



End Sub

Private Sub Text5_Change()
 Text5.PasswordChar = "*"
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        autorizar.SetFocus
    End If
    

End Sub

