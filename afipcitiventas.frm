VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{5734474E-78D3-4254-99B9-C35F31BDF509}#62.0#0"; "vbskpro2.ocx"
Begin VB.Form afipcitiventas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Generación de C.I.T.I. Ventas AFIP"
   ClientHeight    =   6390
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   9525
   Icon            =   "afipcitiventas.frx":0000
   LinkTopic       =   "From1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6390
   ScaleWidth      =   9525
   Begin MSDataGridLib.DataGrid DataGrid4 
      Bindings        =   "afipcitiventas.frx":0442
      Height          =   5655
      Left            =   6240
      TabIndex        =   13
      Top             =   480
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   9975
      _Version        =   393216
      AllowUpdate     =   -1  'True
      AllowArrows     =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      TabAction       =   2
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "afipcitiventas.frx":0458
      Height          =   855
      Left            =   240
      TabIndex        =   7
      Top             =   2880
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   1508
      _Version        =   393216
      AllowUpdate     =   -1  'True
      AllowArrows     =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      TabAction       =   2
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
   Begin MSComctlLib.ProgressBar bar 
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   2040
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Text            =   "Text2"
      Top             =   1680
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton listar 
      Caption         =   "&Generar"
      Height          =   375
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1560
      UseMaskColor    =   -1  'True
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   1800
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "Periodo a Generar"
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
      Height          =   1455
      Left            =   360
      TabIndex        =   1
      Top             =   0
      Width           =   5535
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1800
         TabIndex        =   4
         Top             =   600
         Width           =   1935
      End
      Begin VB.TextBox Text1 
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
         Index           =   3
         Left            =   960
         TabIndex        =   3
         Text            =   "Periodo:"
         Top             =   600
         Width           =   855
      End
   End
   Begin MSAdodcLib.Adodc datprimaryrs 
      Height          =   330
      Left            =   3960
      Top             =   1800
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
      OldForeColor    =   0
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
      LcK1            =   "3.66*/4/0*/1-5*210/."
      LcK2            =   $"afipcitiventas.frx":0471
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
   Begin MSAdodcLib.Adodc datcampo12 
      Height          =   330
      Left            =   0
      Top             =   1440
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
   Begin VB.TextBox Text1 
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
      Height          =   525
      Index           =   0
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   8
      Text            =   "afipcitiventas.frx":0480
      Top             =   2400
      Width           =   3975
   End
   Begin MSAdodcLib.Adodc datcampo13 
      Height          =   330
      Left            =   0
      Top             =   1680
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
   Begin MSDataGridLib.DataGrid DataGrid2 
      Bindings        =   "afipcitiventas.frx":04D3
      Height          =   855
      Left            =   240
      TabIndex        =   9
      Top             =   4080
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   1508
      _Version        =   393216
      AllowUpdate     =   -1  'True
      AllowArrows     =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      TabAction       =   2
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
   Begin VB.TextBox Text1 
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
      Index           =   1
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   10
      Text            =   "afipcitiventas.frx":04EC
      Top             =   3840
      Width           =   5655
   End
   Begin MSDataGridLib.DataGrid DataGrid3 
      Bindings        =   "afipcitiventas.frx":0522
      Height          =   855
      Left            =   240
      TabIndex        =   11
      Top             =   5280
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   1508
      _Version        =   393216
      AllowUpdate     =   -1  'True
      AllowArrows     =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      TabAction       =   2
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
   Begin VB.TextBox Text1 
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
      Index           =   2
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   12
      Text            =   "afipcitiventas.frx":053B
      Top             =   5040
      Width           =   5655
   End
   Begin MSAdodcLib.Adodc datcampo15 
      Height          =   330
      Left            =   1200
      Top             =   1680
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
   Begin MSAdodcLib.Adodc datciti 
      Height          =   330
      Left            =   3960
      Top             =   1440
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
End
Attribute VB_Name = "afipcitiventas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit
Private crApp As New CRAXDRT.Application
Private crReport As New CRAXDRT.Report
Dim meslista1 As String

Private Function LeerArchivo(ByVal strRuta As String) As String
    Dim f As Integer
    f = FreeFile
    Open strRuta For Input As #f
    LeerArchivo = Input(LOF(f), #f)
    Close #f
End Function
Private Sub GuardarArchivo(PTexto As String, pFileName As String)
    Dim ffile As Integer
    ffile = FreeFile
    Open pFileName For Output As #ffile
    Print #ffile, PTexto
    Close #ffile
End Sub

Private Sub Combo1_Click()
Dim meslista0 As String

    meslista0 = Str(Combo1.ListIndex + 1)
    meslista1 = Right(meslista0, Len(meslista0) - 1)
    If Len(meslista1) = 1 Then meslista1 = "0" + meslista1
    
End Sub






Private Sub DataGrid1_GotFocus()

If datcampo12.Recordset.EOF = True Then
    DataGrid1.AllowAddNew = True
    datcampo12.Recordset.AddNew
    datcampo12.Recordset.Fields(0) = login.empresaact
    datcampo12.Recordset.UpdateBatch adAffectCurrent

End If
    DataGrid1.AllowAddNew = False
    
End Sub

Private Sub DataGrid1_LostFocus()
    datcampo12.Recordset.Fields(0) = login.empresaact
    datcampo12.Recordset.UpdateBatch adAffectCurrent
End Sub

Private Sub DataGrid2_GotFocus()

If datcampo13.Recordset.EOF = True Then
    DataGrid2.AllowAddNew = True
    datcampo13.Recordset.AddNew
    datcampo13.Recordset.Fields(0) = login.empresaact
    datcampo13.Recordset.UpdateBatch adAffectCurrent

End If
    DataGrid2.AllowAddNew = False
    
End Sub

Private Sub DataGrid2_LostFocus()
    datcampo13.Recordset.Fields(0) = login.empresaact
    datcampo13.Recordset.UpdateBatch adAffectCurrent
End Sub

Private Sub DataGrid3_GotFocus()
On Error Resume Next
If datcampo15.Recordset.EOF = True Then
    DataGrid3.AllowAddNew = True
    datcampo15.Recordset.AddNew
    datcampo15.Recordset.Fields(0) = login.empresaact
    datcampo15.Recordset.UpdateBatch adAffectCurrent

End If
    DataGrid3.AllowAddNew = False
End Sub

Private Sub DataGrid3_LostFocus()
On Error Resume Next
    datcampo15.Recordset.Fields(0) = login.empresaact
    datcampo15.Recordset.UpdateBatch adAffectCurrent
End Sub

Private Sub Form_Load()
Dim x As Integer
datprimaryrs.ConnectionString = login.conexiontotal
datcampo12.ConnectionString = login.conexiontotal
datcampo13.ConnectionString = login.conexiontotal
datcampo15.ConnectionString = login.conexiontotal
datciti.ConnectionString = login.conexiontotal

datcampo12.RecordSource = "select citi_campo12.* from citi_campo12 where empresa = " & login.empresaact & ""
datcampo12.Refresh
datcampo13.RecordSource = "select citi_campo13.* from citi_campo13 where empresa = " & login.empresaact & ""
datcampo13.Refresh
datcampo15.RecordSource = "select citi_campo15.* from citi_campo15 where empresa = " & login.empresaact & ""
datcampo15.Refresh
datciti.RecordSource = "select citi_afip.* from citi_afip order by codigo"
datciti.Refresh

DataGrid1.Columns(0).Visible = False
DataGrid2.Columns(0).Visible = False
DataGrid3.Columns(0).Visible = False

For x = 0 To 15
    DataGrid1.Columns(x).Width = 500
    DataGrid2.Columns(x).Width = 500
    DataGrid3.Columns(x).Width = 500
Next x

For x = 0 To 3
    DataGrid4.Columns(x).Width = 600
Next x


Combo1.AddItem "ENERO"
Combo1.AddItem "FEBRERO"
Combo1.AddItem "MARZO"
Combo1.AddItem "ABRIL"
Combo1.AddItem "MAYO"
Combo1.AddItem "JUNIO"
Combo1.AddItem "JULIO"
Combo1.AddItem "AGOSTO"
Combo1.AddItem "SEPTIEMBRE"
Combo1.AddItem "OCTUBRE"
Combo1.AddItem "NOVIEMBRE"
Combo1.AddItem "DICIEMBRE"

meslista1 = "N"

End Sub


Private Sub listar_Click()
Dim campo As String
Dim ruta As String
Dim Ret As Long
Dim origen As String
Dim destino As String
Dim x As Integer
Dim i As Integer

bar.Min = 0

datprimaryrs.RecordSource = "select citi_ventas.* from citi_ventas where empresa = " & login.empresaact & " and inicioper = '" & login.iper & "' and cerrado = '" & meslista1 & "' order by campo02"
datprimaryrs.Refresh
If datprimaryrs.Recordset.EOF = True Then meslista1 = "N"
datprimaryrs.RecordSource = "select citi_ventas.* from citi_ventas where empresa = " & login.empresaact & " and inicioper = '" & login.iper & "' and cerrado = '" & meslista1 & "' order by campo02"
datprimaryrs.Refresh
If datprimaryrs.Recordset.EOF = True Then Exit Sub


bar.max = datprimaryrs.Recordset.RecordCount
i = 0
datprimaryrs.Recordset.MoveFirst


Text2.Text = ""
Do While Not datprimaryrs.Recordset.EOF
    
    For x = 0 To 31
        Text2.Text = Text2.Text + datprimaryrs.Recordset.Fields(x)
    Next x
   
    datprimaryrs.Recordset.MoveNext
    Text2.Text = Text2.Text + (Chr(13) + Chr(10))
    bar.Value = i
    i = i + 1
   
  
Loop
Open App.Path & "\banco2.txt" For Output As #1
    
    Write #1, Text2.Text

Close #1

origen = App.Path & "\banco2.txt"
destino = App.Path & "\citi.txt"

Call GuardarArchivo(Replace(LeerArchivo(origen), """", ""), destino)



    ruta = App.Path & "\notepad.exe citi.txt"

    Ret = Shell(ruta, vbNormalFocus)


    
End Sub

