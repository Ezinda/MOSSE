VERSION 5.00
Object = "{F6125AB1-8AB1-11CE-A77F-08002B2F4E98}#2.0#0"; "MSRDC20.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{5734474E-78D3-4254-99B9-C35F31BDF509}#62.0#0"; "vbskpro2.ocx"
Object = "{D18BBD1F-82BB-4385-BED3-E9D31A3E361E}#1.0#0"; "KewlButtonz.ocx"
Begin VB.Form imprecibolistado 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de Recibos Cobrados"
   ClientHeight    =   2505
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   6210
   Icon            =   "impreciboslistado.frx":0000
   LinkTopic       =   "From1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2505
   ScaleWidth      =   6210
   Begin VB.TextBox Text1 
      DataField       =   "cerrado"
      DataSource      =   "criterio"
      Height          =   285
      Index           =   8
      Left            =   1440
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   1920
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
      Top             =   840
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Text1 
      DataField       =   "desde"
      DataSource      =   "criterio"
      Height          =   285
      Index           =   1
      Left            =   1680
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   1440
      Visible         =   0   'False
      Width           =   855
   End
   Begin MSComCtl2.DTPicker cargahasta 
      Height          =   375
      Left            =   3840
      TabIndex        =   3
      Top             =   360
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   12648447
      Format          =   65339393
      CurrentDate     =   38415
   End
   Begin MSComCtl2.DTPicker cargadesde 
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Top             =   360
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   12648447
      Format          =   65339393
      CurrentDate     =   38415
   End
   Begin VB.TextBox Text1 
      DataField       =   "empresa"
      DataSource      =   "criterio"
      Height          =   285
      Index           =   0
      Left            =   0
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   1920
      Visible         =   0   'False
      Width           =   615
   End
   Begin Crystal.CrystalReport CrystalReporte 
      Left            =   120
      Top             =   1680
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTitle     =   "Listado de Recibos Emitidos"
      PrintFileLinesPerPage=   60
      WindowShowGroupTree=   -1  'True
      WindowShowCloseBtn=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   255
      Left            =   480
      TabIndex        =   0
      Top             =   2040
      Visible         =   0   'False
      Width           =   975
   End
   Begin MSRDC.MSRDC reporte 
      Height          =   375
      Left            =   360
      Top             =   1560
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
      Height          =   1575
      Left            =   360
      TabIndex        =   6
      Top             =   0
      Width           =   5535
      Begin VB.CheckBox Check1 
         Caption         =   "Detalle Comprobantes"
         Height          =   255
         Left            =   3360
         TabIndex        =   13
         Top             =   1200
         Width           =   2055
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   0
         Left            =   2640
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   360
         Width           =   855
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Desde"
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   360
         Width           =   855
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Ordeneado x Fecha de emisón"
         Height          =   255
         Left            =   360
         TabIndex        =   9
         Top             =   1200
         Width           =   2895
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Ordenado x Fecha de cobro"
         Height          =   255
         Left            =   360
         TabIndex        =   8
         Top             =   840
         Width           =   2895
      End
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
      LcK2            =   $"impreciboslistado.frx":0442
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
      TabIndex        =   12
      Top             =   1680
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1085
      BTYPE           =   14
      TX              =   "&Listar"
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
      MICON           =   "impreciboslistado.frx":0451
      PICN            =   "impreciboslistado.frx":046D
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
Attribute VB_Name = "imprecibolistado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private crApp As New CRAXDRT.Application
Private crReport As New CRAXDRT.Report


Private Sub Command1_Click()
Dim tabla As String
Dim tabla1 As String
Dim ruta As String

ruta = "\Empresa" + Right(Str(login.empresaact), Len(Str(login.empresaact)) - 1)


If Option1.Value = False Then
    reporte.SQL = "SELECT consultarecibocobro.nrorden, consultarecibocobro.nomcliente, consultarecibocobro.importe, consultarecibocobro.razonsocial, consultarecibocobro.fecha, consultarecibocobro.cuitclien FROM contablesql.dbo.consultarecibocobro consultarecibocobro WHERE consultarecibocobro.empresa = " & login.empresaact & "  and consultarecibocobro.fecha >= '" & cargadesde.Value & "' and consultarecibocobro.fecha <= '" & cargahasta.Value & "' ORDER BY consultarecibocobro.nrorden,consultarecibocobro.fecha,consultarecibocobro.fechacompro  ASC "
Else
    reporte.SQL = "SELECT consultarecibocobro.nrorden, consultarecibocobro.nomcliente, consultarecibocobro.importe, consultarecibocobro.razonsocial, consultarecibocobro.fecha, consultarecibocobro.cuitclien FROM contablesql.dbo.consultarecibocobro consultarecibocobro WHERE consultarecibocobro.empresa = " & login.empresaact & "  and consultarecibocobro.fecha >= '" & cargadesde.Value & "' and consultarecibocobro.fecha <= '" & cargahasta.Value & "' ORDER BY consultarecibocobro.fechacobro,consultarecibocobro.nrorden,consultarecibocobro.fechacompro  ASC"
End If

tabla = reporte.SQL

With CrystalReporte
If Check1.Value = 0 Then
    .ReportFileName = App.Path & ruta + "\ListadoReciboscobro.rpt"
Else
    .ReportFileName = App.Path & ruta + "\ListadoReciboscobro_detalle.rpt"
End If
    .Connect = login.conexionreporte
    .DiscardSavedData = True
    .RetrieveDataFiles
    .ReportSource = 0
    .SQLQuery = tabla
    .Destination = crptToWindow
 Rem   .Destination = crptToPrinter
    .PrintFileType = crptCrystal
    .WindowState = crptMaximized
    .Action = 1
     
End With
End Sub


Private Sub Form_Load()
Aplicar_skin Me


Option1.Value = True
Option2.Value = False

Text1(0).Text = login.empresaact
cargadesde = Date - Day(Date) + 1
cargahasta = Date
Text1(1).Text = cargadesde
Text1(2).Text = cargahasta
Check1.Value = 0

End Sub


Private Sub listar_Click()

    Text1(1).Text = cargadesde.Value
    Text1(2).Text = cargahasta.Value

    Call Command1_Click

End Sub

