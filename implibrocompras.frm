VERSION 5.00
Object = "{F6125AB1-8AB1-11CE-A77F-08002B2F4E98}#2.0#0"; "MSRDC20.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{5734474E-78D3-4254-99B9-C35F31BDF509}#62.0#0"; "vbskpro2.ocx"
Object = "{D18BBD1F-82BB-4385-BED3-E9D31A3E361E}#1.0#0"; "KewlButtonz.ocx"
Begin VB.Form implibrocompras 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Impresión Libro Iva Compras"
   ClientHeight    =   2505
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   6210
   Icon            =   "implibrocompras.frx":0000
   LinkTopic       =   "From1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2505
   ScaleWidth      =   6210
   Begin VB.TextBox Text1 
      DataField       =   "periodo"
      DataSource      =   "criterio"
      Height          =   285
      Index           =   4
      Left            =   0
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   1800
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox Text1 
      DataField       =   "cerrado"
      DataSource      =   "criterio"
      Height          =   285
      Index           =   8
      Left            =   600
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   1560
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox Text1 
      DataField       =   "hasta"
      DataSource      =   "criterio"
      Height          =   285
      Index           =   2
      Left            =   0
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   1560
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Text1 
      DataField       =   "desde"
      DataSource      =   "criterio"
      Height          =   285
      Index           =   1
      Left            =   600
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   1320
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox Text1 
      DataField       =   "empresa"
      DataSource      =   "criterio"
      Height          =   285
      Index           =   0
      Left            =   0
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   1320
      Visible         =   0   'False
      Width           =   615
   End
   Begin Crystal.CrystalReport CrystalReporte 
      Left            =   720
      Top             =   2040
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      WindowTitle     =   "Libro IVA Compras"
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   255
      Left            =   1560
      TabIndex        =   2
      Top             =   2160
      Visible         =   0   'False
      Width           =   975
   End
   Begin MSRDC.MSRDC reporte 
      Height          =   375
      Left            =   4680
      Top             =   1680
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
   Begin VB.Frame Frame1 
      Caption         =   "Periodo a Listar"
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
      TabIndex        =   6
      Top             =   120
      Width           =   5535
      Begin VB.CommandButton Command6 
         Caption         =   "Periodo"
         Height          =   255
         Index           =   2
         Left            =   1200
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   600
         Width           =   855
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   2160
         TabIndex        =   0
         Top             =   600
         Width           =   1935
      End
   End
   Begin MSAdodcLib.Adodc datprimaryrs 
      Height          =   330
      Left            =   4320
      Top             =   2040
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
      LcK2            =   $"implibrocompras.frx":0442
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
   Begin KewlButtonz.KewlButtons listar 
      Height          =   615
      Left            =   2520
      TabIndex        =   1
      Top             =   1680
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1085
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
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   -2147483629
      BCOLO           =   -2147483629
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "implibrocompras.frx":0451
      PICN            =   "implibrocompras.frx":046D
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
Attribute VB_Name = "implibrocompras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private crApp As New CRAXDRT.Application
Private crReport As New CRAXDRT.Report
Dim meslista1 As String


Private Sub Combo1_Click()
Dim meslista0 As String

    meslista0 = Str(Combo1.ListIndex + 1)
    meslista1 = Right(meslista0, Len(meslista0) - 1)
    If Len(meslista1) = 1 Then meslista1 = "0" + meslista1

    
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 9
        listar.SetFocus
    End If

End Sub

Private Sub Command1_Click()
Dim tabla As String
Dim ruta As String

ruta = "\Empresa" + Right(Str(login.empresaact), Len(Str(login.empresaact)) - 1)

reporte.SQL = "SELECT librocompras1.id, librocompras1.cerrado, librocompras1.empresa, librocompras1.inicioper, librocompras1.fecha, librocompras1.proveedor, librocompras1.tipoiva, librocompras1.cuit, librocompras1.tipocompr, librocompras1.numcompr, librocompras1.col1, librocompras1.col2, librocompras1.col3, librocompras1.col4, librocompras1.col5, librocompras1.col6, librocompras1.total, librocompras1.nomcol1, librocompras1.nomcol2, librocompras1.nomcol3, librocompras1.nomcol4, librocompras1.nomcol5, librocompras1.nomcol6, librocompras1.nomcol7, librocompras1.razonsocial FROM contablesql.dbo.librocompras1 librocompras1 WHERE librocompras1.cerrado = '" & meslista1 & "' and librocompras1.empresa = " & login.empresaact & " and librocompras1.inicioper = '" & login.iper & "' ORDER BY librocompras1.fecha ASC, librocompras1.id ASC"
tabla = reporte.SQL

With CrystalReporte
    .ReportFileName = App.Path & ruta + "\librocompras.rpt"
    .Connect = login.conexionreporte
    .DiscardSavedData = True
    .RetrieveDataFiles
    .ReportSource = 0
    .SQLQuery = tabla
    .Destination = crptToWindow
    .PrintFileType = crptCrystal
    .WindowState = crptMaximized
    .Action = 1
      
End With

End Sub


Private Sub Form_Load()
Aplicar_skin Me

datprimaryrs.ConnectionString = login.conexiontotal


Text1(0).Text = login.empresaact
Rem cargadesde = Date - Day(Date) + 1
Rem cargahasta = Date
Rem text1(1).Text = cargadesde
Rem text1(2).Text = cargahasta

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

datprimaryrs.RecordSource = "select librocompras.* from librocompras where empresa = " & login.empresaact & " and inicioper = '" & login.iper & "' and cerrado = '" & meslista1 & "' order by cerrado"
datprimaryrs.Refresh
If datprimaryrs.Recordset.EOF = True Then meslista1 = "N"

    Call Command1_Click

End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 9
        listar.SetFocus
    End If
    
End Sub
