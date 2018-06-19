VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{D18BBD1F-82BB-4385-BED3-E9D31A3E361E}#1.0#0"; "KewlButtonz.ocx"
Begin VB.Form asienmodelo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Asientos Modelo"
   ClientHeight    =   7050
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   7320
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7050
   ScaleWidth      =   7320
   Begin VB.CommandButton llena 
      Caption         =   "llena"
      Height          =   247
      Left            =   234
      TabIndex        =   24
      Top             =   6786
      Visible         =   0   'False
      Width           =   715
   End
   Begin VB.Frame Frame3 
      Caption         =   "Movimientos"
      Height          =   2587
      Left            =   5128
      TabIndex        =   23
      Top             =   4212
      Width           =   1885
      Begin KewlButtonz.KewlButtons KewlButtons1 
         Height          =   481
         Left            =   351
         TabIndex        =   11
         Top             =   234
         Width           =   1092
         _ExtentX        =   1905
         _ExtentY        =   847
         BTYPE           =   14
         TX              =   "G&rabar"
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
         MICON           =   "asienmodelo.frx":0000
         PICN            =   "asienmodelo.frx":001C
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
         Height          =   481
         Left            =   351
         TabIndex        =   12
         Top             =   819
         Width           =   1092
         _ExtentX        =   1905
         _ExtentY        =   847
         BTYPE           =   14
         TX              =   "Nuev&o"
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
         MICON           =   "asienmodelo.frx":1A9E
         PICN            =   "asienmodelo.frx":1ABA
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
         Height          =   481
         Left            =   351
         TabIndex        =   13
         Top             =   1404
         Width           =   1092
         _ExtentX        =   1905
         _ExtentY        =   847
         BTYPE           =   14
         TX              =   "Borr&ar"
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
         MICON           =   "asienmodelo.frx":4EAC
         PICN            =   "asienmodelo.frx":4EC8
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin KewlButtonz.KewlButtons cerrar 
         Height          =   494
         Left            =   351
         TabIndex        =   14
         Top             =   1989
         Width           =   1066
         _ExtentX        =   1879
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
         BCOL            =   -2147483629
         BCOLO           =   -2147483629
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "asienmodelo.frx":82BA
         PICN            =   "asienmodelo.frx":82D6
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
   Begin VB.CommandButton verificacuenta 
      Caption         =   "verificacuenta"
      Height          =   255
      Left            =   5616
      TabIndex        =   21
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Modelo"
      Height          =   2704
      Left            =   5138
      TabIndex        =   20
      Top             =   117
      Width           =   1885
      Begin KewlButtonz.KewlButtons grabar 
         Height          =   481
         Left            =   351
         TabIndex        =   1
         Top             =   234
         Width           =   1092
         _ExtentX        =   1905
         _ExtentY        =   847
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
         MICON           =   "asienmodelo.frx":8E20
         PICN            =   "asienmodelo.frx":8E3C
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
         Height          =   481
         Left            =   351
         TabIndex        =   2
         Top             =   819
         Width           =   1092
         _ExtentX        =   1905
         _ExtentY        =   847
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
         MICON           =   "asienmodelo.frx":A8BE
         PICN            =   "asienmodelo.frx":A8DA
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
         Height          =   481
         Left            =   351
         TabIndex        =   3
         Top             =   1404
         Width           =   1092
         _ExtentX        =   1905
         _ExtentY        =   847
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
         MICON           =   "asienmodelo.frx":DCCC
         PICN            =   "asienmodelo.frx":DCE8
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
         Height          =   481
         Left            =   351
         TabIndex        =   4
         Top             =   1989
         Width           =   1092
         _ExtentX        =   1905
         _ExtentY        =   847
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
         MICON           =   "asienmodelo.frx":E6FA
         PICN            =   "asienmodelo.frx":E716
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
      Bindings        =   "asienmodelo.frx":11B08
      Height          =   2587
      Left            =   234
      TabIndex        =   0
      Top             =   234
      Width           =   4914
      _ExtentX        =   8652
      _ExtentY        =   4551
      _Version        =   393216
      AllowArrows     =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      TabAction       =   2
      WrapCellPointer =   -1  'True
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
            LCID            =   1034
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
            LCID            =   1034
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
   Begin VB.Frame Frame2 
      Height          =   1339
      Left            =   234
      TabIndex        =   15
      Top             =   2808
      Width           =   6773
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         DataField       =   "deschaber"
         DataSource      =   "datasiento_mov"
         Height          =   247
         Index           =   6
         Left            =   2925
         MaxLength       =   30
         TabIndex        =   10
         Top             =   936
         Width           =   3172
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         DataField       =   "descdeb"
         DataSource      =   "datasiento_mov"
         Height          =   247
         Index           =   5
         Left            =   2925
         MaxLength       =   30
         TabIndex        =   8
         Top             =   585
         Width           =   3172
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         DataField       =   "debe"
         DataSource      =   "datasiento_mov"
         Height          =   247
         Index           =   3
         Left            =   1872
         TabIndex        =   7
         Top             =   585
         Width           =   949
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         DataField       =   "haber"
         DataSource      =   "datasiento_mov"
         Height          =   247
         Index           =   4
         Left            =   1872
         TabIndex        =   9
         Top             =   936
         Width           =   949
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         DataField       =   "idorden"
         DataSource      =   "datasiento_mov"
         Height          =   247
         Index           =   2
         Left            =   1872
         TabIndex        =   6
         Top             =   234
         Width           =   949
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         DataField       =   "cod"
         DataSource      =   "datasiento_mov"
         Height          =   247
         Index           =   0
         Left            =   5148
         MaxLength       =   3
         TabIndex        =   5
         Top             =   234
         Visible         =   0   'False
         Width           =   949
      End
      Begin VB.CommandButton Cuenta 
         Caption         =   "Cuenta Haber:"
         Height          =   255
         Index           =   0
         Left            =   234
         Picture         =   "asienmodelo.frx":11B21
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   936
         UseMaskColor    =   -1  'True
         Width           =   1455
      End
      Begin VB.CommandButton Cuenta 
         Caption         =   "Cuenta Debe:"
         Height          =   255
         Index           =   6
         Left            =   240
         Picture         =   "asienmodelo.frx":12053
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   585
         UseMaskColor    =   -1  'True
         Width           =   1455
      End
      Begin VB.CommandButton Cuenta 
         Caption         =   "N° Linea:"
         Height          =   255
         Index           =   5
         Left            =   240
         Picture         =   "asienmodelo.frx":12585
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   234
         UseMaskColor    =   -1  'True
         Width           =   1455
      End
      Begin VB.CommandButton Cuenta 
         Caption         =   "Cod.Asiento"
         Height          =   255
         Index           =   4
         Left            =   3393
         Picture         =   "asienmodelo.frx":12AB7
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   234
         UseMaskColor    =   -1  'True
         Visible         =   0   'False
         Width           =   1455
      End
   End
   Begin MSAdodcLib.Adodc datasiento 
      Height          =   330
      Left            =   0
      Top             =   0
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
      LockType        =   2
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
      UserName        =   "lucva"
      Password        =   "25072004"
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
      Height          =   325
      Left            =   1521
      Top             =   0
      Visible         =   0   'False
      Width           =   1196
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
   Begin MSDataGridLib.DataGrid DataGrid2 
      Bindings        =   "asienmodelo.frx":12FE9
      Height          =   2587
      Left            =   234
      TabIndex        =   22
      Top             =   4212
      Width           =   4797
      _ExtentX        =   8440
      _ExtentY        =   4551
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
            LCID            =   1034
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
            LCID            =   1034
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
   Begin MSAdodcLib.Adodc datasiento_mov 
      Height          =   325
      Left            =   3042
      Top             =   0
      Visible         =   0   'False
      Width           =   1196
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   2
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
      UserName        =   "lucva"
      Password        =   "25072004"
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
Attribute VB_Name = "asienmodelo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim posicion As Integer

Private Sub agregamismonivel_Click()
On Error Resume Next
    
    datasiento.Recordset.AddNew
    datasiento.Recordset.Fields("empresa") = login.empresaact
    DataGrid1.SetFocus
    

End Sub

Private Sub borrar_Click()

    mensa = MsgBox("Esta por borrar este modelo de asiento, esta seguro?", vbYesNo, "¡ Atención !")
    If mensa = vbNo Then Exit Sub
    
    datasiento.Recordset.Delete adAffectCurrent
    datasiento.Refresh
    

End Sub

Private Sub Cancelar_Click()

    datasiento.Refresh
    Text1(0).SetFocus

End Sub

Private Sub cerrar_Click()

 Unload Me

End Sub

Private Sub DataGrid1_Click()


Call llena_Click

End Sub

Private Sub DataGrid1_KeyUp(KeyCode As Integer, Shift As Integer)


Call llena_Click

End Sub

Private Sub Form_Load()
asienmodelo.Top = 0
asienmodelo.Left = 0
Aplicar_skin Me

datasiento.ConnectionString = login.conexiontotal
datcuentas.ConnectionString = login.conexiontotal
datasiento_mov.ConnectionString = login.conexiontotal

datasiento.RecordSource = "select asiemod.* from asiemod where empresa = " & login.empresaact & " order by cod"
datasiento.Refresh

If datasiento.Recordset.EOF = False Then datasiento.Recordset.MoveLast

DataGrid1.Columns(1).Visible = False

DataGrid1.Columns(0).Width = 800
DataGrid1.Columns(2).Width = 2500
DataGrid1.Columns(3).Width = 500

End Sub

Private Sub grabar_Click()

    DataGrid1.Columns(0).Text = Left(DataGrid1.Columns(0).Text, 3)
    DataGrid1.Columns(2).Text = Left(DataGrid1.Columns(2).Text, 30)

    datasiento.Recordset.UpdateBatch adAffectCurrent
    Call llena_Click
    agregamismonivel.SetFocus
    

End Sub

Private Sub KewlButtons1_Click()
On Error Resume Next
    datasiento_mov.Recordset.UpdateBatch adAffectCurrent
    KewlButtons2.SetFocus


End Sub

Private Sub KewlButtons2_Click()

    datasiento_mov.Recordset.AddNew
    datasiento_mov.Recordset.Fields("cod") = DataGrid1.Columns(0).Text
    datasiento_mov.Recordset.Fields("empresa") = login.empresaact

End Sub

Private Sub llena_Click()
On Error Resume Next
datasiento_mov.RecordSource = "select asienmod_mov.* from asienmod_mov where empresa = " & login.empresaact & " and cod = '" & DataGrid1.Columns(0).Text & "'  order by idorden"
datasiento_mov.Refresh

DataGrid2.Columns(0).Visible = False
DataGrid2.Columns(1).Visible = False
DataGrid2.Columns(4).Visible = False
DataGrid2.Columns(6).Visible = False

DataGrid2.Columns(2).Width = 800
DataGrid2.Columns(3).Width = 800
DataGrid2.Columns(5).Width = 800



End Sub

Private Sub Text1_GotFocus(Index As Integer)

    If ventana.menu = 8 And Index = 3 Then
        ventana.menu = 0
        Text1(3).Text = lista_cuentas.cuentacont
    End If

    If ventana.menu = 8 And Index = 4 Then
        ventana.menu = 0
        Text1(4).Text = lista_cuentas.cuentacont
    End If

End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)


    If KeyAscii = 13 Then
        KeyAscii = 0
        salir = 0
        posicion = Index
                        
        If Index = 3 Or Index = 4 Then
            posicion = Index
            Call verificacuenta_Click
        End If
                       
        SendKeys "{tab}", False
        
    End If

End Sub

Private Sub Text1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)

    
    If KeyCode = 114 And Index = 3 Then
        lista_cuentas.cuentacont = Text1(3).Text
        ventana.menu = 8
        lista_cuentas.Show
    End If
    
    If KeyCode = 114 And Index = 4 Then
        lista_cuentas.cuentacont = Text1(4).Text
        ventana.menu = 8
        lista_cuentas.Show
    End If

End Sub

Private Sub verificacuenta_Click()

    If Text1(posicion) = "" Then Exit Sub

    datcuentas.ConnectionString = login.conexiontotal
    datcuentas.RecordSource = "select cuentas.* from cuentas where empre = " & login.empresaact & " and imp = 'S' and [cod contable] = " & Text1(posicion).Text & " and inicioper = '" & login.iper & "'"
    datcuentas.Refresh
    
    If datcuentas.Recordset.EOF = True Then
        MsgBox "No Existe esta cuenta contable", vbCritical, "Verificar"
        Text1(posicion).Text = ""
        Text1(posicion).SetFocus
    End If
    

End Sub
