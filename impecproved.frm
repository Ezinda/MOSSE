VERSION 5.00
Object = "{F6125AB1-8AB1-11CE-A77F-08002B2F4E98}#2.0#0"; "MSRDC20.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{5734474E-78D3-4254-99B9-C35F31BDF509}#62.0#0"; "vbskpro2.ocx"
Object = "{D18BBD1F-82BB-4385-BED3-E9D31A3E361E}#1.0#0"; "KewlButtonz.ocx"
Begin VB.Form impecproved 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Impresión Estado de Cuenta Proveedores"
   ClientHeight    =   5280
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   6210
   Icon            =   "impecproved.frx":0000
   LinkTopic       =   "From1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5280
   ScaleWidth      =   6210
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   255
      Left            =   4800
      TabIndex        =   13
      Top             =   4080
      Visible         =   0   'False
      Width           =   975
   End
   Begin MSDataListLib.DataCombo combohasta 
      Bindings        =   "impecproved.frx":0442
      Height          =   315
      Left            =   840
      TabIndex        =   3
      Top             =   3480
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "razonsocial"
      Text            =   ""
   End
   Begin VB.Frame Frame2 
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
      ForeColor       =   &H00800000&
      Height          =   2175
      Left            =   360
      TabIndex        =   12
      Top             =   2160
      Width           =   5535
      Begin VB.CommandButton Command6 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   1
         Left            =   480
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   1080
         Width           =   855
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Desde"
         Height          =   255
         Index           =   0
         Left            =   480
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   360
         Width           =   855
      End
      Begin MSDataListLib.DataCombo combodesde 
         Bindings        =   "impecproved.frx":0459
         Height          =   315
         Left            =   480
         TabIndex        =   2
         Top             =   600
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "razonsocial"
         Text            =   ""
      End
   End
   Begin VB.TextBox Text1 
      DataField       =   "cerrado"
      DataSource      =   "criterio"
      Height          =   285
      Index           =   8
      Left            =   600
      TabIndex        =   11
      Text            =   "Text1"
      Top             =   3960
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox Text1 
      DataField       =   "hasta"
      DataSource      =   "criterio"
      Height          =   285
      Index           =   2
      Left            =   0
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   3960
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Text1 
      DataField       =   "desde"
      DataSource      =   "criterio"
      Height          =   285
      Index           =   1
      Left            =   600
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   3720
      Visible         =   0   'False
      Width           =   855
   End
   Begin MSComCtl2.DTPicker cargahasta 
      Height          =   375
      Left            =   3720
      TabIndex        =   1
      Top             =   960
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   12648447
      Format          =   21364737
      CurrentDate     =   38415
   End
   Begin MSComCtl2.DTPicker cargadesde 
      Height          =   375
      Left            =   1440
      TabIndex        =   0
      Top             =   960
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   12648447
      Format          =   21364737
      CurrentDate     =   38415
   End
   Begin VB.TextBox Text1 
      DataField       =   "empresa"
      DataSource      =   "criterio"
      Height          =   285
      Index           =   0
      Left            =   0
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   3720
      Visible         =   0   'False
      Width           =   615
   End
   Begin Crystal.CrystalReport CrystalReporte 
      Left            =   120
      Top             =   3000
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTitle     =   "Estado de Cuenta Proveedores"
      PrintFileLinesPerPage=   60
      WindowShowGroupTree=   -1  'True
      WindowAllowDrillDown=   -1  'True
      WindowShowCloseBtn=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   255
      Left            =   480
      TabIndex        =   6
      Top             =   3240
      Visible         =   0   'False
      Width           =   975
   End
   Begin MSRDC.MSRDC reporte 
      Height          =   330
      Left            =   240
      Top             =   3120
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
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
      Height          =   2055
      Left            =   360
      TabIndex        =   10
      Top             =   0
      Width           =   5535
      Begin VB.CommandButton Command6 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   3
         Left            =   2520
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   960
         Width           =   855
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Desde"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   960
         Width           =   855
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Por Fecha Comprobante"
         Height          =   255
         Left            =   3000
         TabIndex        =   16
         Top             =   480
         Width           =   2295
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Por Fecha Contable"
         Height          =   255
         Left            =   360
         TabIndex        =   15
         Top             =   480
         Width           =   1935
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Solo Prove.con saldos"
         Height          =   255
         Left            =   1800
         TabIndex        =   14
         Top             =   1680
         Width           =   1935
      End
   End
   Begin MSAdodcLib.Adodc criterio 
      Height          =   330
      Left            =   4440
      Top             =   3600
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
   Begin MSAdodcLib.Adodc datprove 
      Height          =   330
      Left            =   4560
      Top             =   3840
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
      LcK2            =   $"impecproved.frx":0470
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
      Left            =   1680
      TabIndex        =   4
      Top             =   4440
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1085
      BTYPE           =   14
      TX              =   "&Detalle"
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
      MICON           =   "impecproved.frx":047F
      PICN            =   "impecproved.frx":049B
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin KewlButtonz.KewlButtons saldos 
      Height          =   615
      Left            =   3240
      TabIndex        =   5
      Top             =   4440
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1085
      BTYPE           =   14
      TX              =   "&Saldos"
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
      MICON           =   "impecproved.frx":0EAD
      PICN            =   "impecproved.frx":0EC9
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
Attribute VB_Name = "impecproved"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private crApp As New CRAXDRT.Application
Private crReport As New CRAXDRT.Report
Dim empresareal As Integer


Private Sub cargadesde_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cargahasta.SetFocus
    End If
End Sub

Private Sub cargahasta_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        combodesde.SetFocus
    End If
    
End Sub


Private Sub combodesde_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        combohasta.SetFocus
    End If

End Sub

Private Sub Command1_Click()
On Error GoTo fuera
Dim tabla As String
Dim tabla1 As String
Dim desdeprov As String
Dim hastaprov As String
Dim ruta As String
Dim reportever As String

ruta = "\Empresa" + Right(Str(login.empresaact), Len(Str(login.empresaact)) - 1)

criterio.Recordset.Fields(0) = login.empresaact
criterio.Recordset.UpdateBatch adAffectCurrent
desdeprov = combodesde.Text
hastaprov = combohasta.Text

If Option1.Value = True Then
    reporte.SQL = "SELECT ec_proveedores_final1.fecha, ec_proveedores_final1.comprobante, ec_proveedores_final1.asiento, ec_proveedores_final1.cdt, ec_proveedores_final1.nomproveedor, ec_proveedores_final1.fechaasiento, ec_proveedores_final1.debe, ec_proveedores_final1.haber, ec_proveedores_final1.movimiento, ec_proveedores_final1.nrorden FROM contablesql.dbo.ec_proveedores_final1 ec_proveedores_final1 where ec_proveedores_final1.empresa = " & login.empresaact & " and  ec_proveedores_final1.fechaasiento >= '" & cargadesde.Value & "' and ec_proveedores_final1.fechaasiento <= '" & cargahasta.Value & "' and ec_proveedores_final1.nomproveedor >= '" & combodesde.Text & "' and ec_proveedores_final1.nomproveedor <= '" & combohasta.Text & "' ORDER BY ec_proveedores_final1.nomproveedor ASC, ec_proveedores_final1.comprobante ASC, ec_proveedores_final1.movimiento DESC"
Else
    reporte.SQL = "SELECT ec_proveedores_final1_fc.fecha, ec_proveedores_final1_fc.comprobante, ec_proveedores_final1_fc.asiento, ec_proveedores_final1_fc.cdt, ec_proveedores_final1_fc.nomproveedor, ec_proveedores_final1_fc.fechaasiento, ec_proveedores_final1_fc.debe, ec_proveedores_final1_fc.haber, ec_proveedores_final1_fc.movimiento, ec_proveedores_final1_fc.nrorden FROM contablesql.dbo.ec_proveedores_final1_fc ec_proveedores_final1_fc where ec_proveedores_final1_fc.empresa = " & login.empresaact & " and ec_proveedores_final1_fc.fecha >= '" & cargadesde.Value & "' and ec_proveedores_final1_fc.fecha <= '" & cargahasta.Value & "' and ec_proveedores_final1_fc.nomproveedor >= '" & combodesde.Text & "' and ec_proveedores_final1_fc.nomproveedor <= '" & combohasta.Text & "' ORDER BY ec_proveedores_final1_fc.nomproveedor ASC, ec_proveedores_final1_fc.comprobante ASC, ec_proveedores_final1_fc.movimiento DESC"
End If


tabla = reporte.SQL

With CrystalReporte
If Option1.Value = True Then
    .ReportFileName = App.Path & ruta + "\ecproveedores.rpt"
Else
    .ReportFileName = App.Path & ruta + "\ecproveedores_fc.rpt"
End If
    .Connect = login.conexionreporte
    .Formulas(0) = "desdefecha=""" & cargadesde.Value & """"
    .Formulas(1) = "hastafecha=""" & cargahasta.Value & """"
    .Formulas(2) = "empresa=""" & login.nomempresa & """"
Rem    .SubreportToChange = .GetNthSubreportName(0)
Rem    .Connect = login.conexionreporte
Rem    .SubreportToChange = ""
Rem    .Connect = login.conexionreporte
    .DiscardSavedData = True
    .RetrieveDataFiles
    .ReportSource = 0
    .SQLQuery = tabla
    

    .Destination = crptToWindow
    .PrintFileType = crptCrystal
    .WindowState = crptMaximized
    .Action = 1
     
End With
fuera:
End Sub


Private Sub DataCombo1_Click(Area As Integer)

End Sub

Private Sub Command2_Click()
On Error GoTo fuera
Dim tabla As String
Dim tabla1 As String
Dim desdeprov As String
Dim hastaprov As String
Dim ruta As String
Dim reportever As String
Dim saldomuestra As String

ruta = "\Empresa" + Right(Str(login.empresaact), Len(Str(login.empresaact)) - 1)

criterio.Recordset.Fields(0) = login.empresaact
criterio.Recordset.UpdateBatch adAffectCurrent
desdeprov = combodesde.Text
hastaprov = combohasta.Text


If Option1.Value = True Then
    reporte.SQL = "SELECT ec_proveedores_final1.fecha, ec_proveedores_final1.comprobante, ec_proveedores_final1.asiento, ec_proveedores_final1.cdt, ec_proveedores_final1.nomproveedor, ec_proveedores_final1.fechaasiento, ec_proveedores_final1.debe, ec_proveedores_final1.haber, ec_proveedores_final1.movimiento, ec_proveedores_final1.nrorden FROM contablesql.dbo.ec_proveedores_final1 ec_proveedores_final1 where ec_proveedores_final1.empresa = " & login.empresaact & " and  ec_proveedores_final1.fechaasiento >= '" & cargadesde.Value & "' and ec_proveedores_final1.fechaasiento <= '" & cargahasta.Value & "' and ec_proveedores_final1.nomproveedor >= '" & combodesde.Text & "' and ec_proveedores_final1.nomproveedor <= '" & combohasta.Text & "' ORDER BY ec_proveedores_final1.nomproveedor ASC, ec_proveedores_final1.comprobante ASC, ec_proveedores_final1.movimiento DESC"
Else
    reporte.SQL = "SELECT ec_proveedores_final1_fc.fecha, ec_proveedores_final1_fc.comprobante, ec_proveedores_final1_fc.asiento, ec_proveedores_final1_fc.cdt, ec_proveedores_final1_fc.nomproveedor, ec_proveedores_final1_fc.fechaasiento, ec_proveedores_final1_fc.debe, ec_proveedores_final1_fc.haber, ec_proveedores_final1_fc.movimiento, ec_proveedores_final1_fc.nrorden FROM contablesql.dbo.ec_proveedores_final1_fc ec_proveedores_final1_fc where ec_proveedores_final1_fc.empresa = " & login.empresaact & " and  ec_proveedores_final1_fc.fecha >= '" & cargadesde.Value & "' and ec_proveedores_final1_fc.fecha <= '" & cargahasta.Value & "' and ec_proveedores_final1_fc.nomproveedor >= '" & combodesde.Text & "' and ec_proveedores_final1_fc.nomproveedor <= '" & combohasta.Text & "' ORDER BY ec_proveedores_final1_fc.nomproveedor ASC, ec_proveedores_final1_fc.comprobante ASC, ec_proveedores_final1_fc.movimiento DESC"
End If

tabla = reporte.SQL
If Check2.Value = 1 Then
    saldomuestra = "S"
Else
    saldomuestra = "N"
End If

With CrystalReporte

If Option1.Value = True Then
    .ReportFileName = App.Path & ruta + "\ecproveedores2.rpt"
Else
    .ReportFileName = App.Path & ruta + "\ecproveedores2_fc.rpt"
End If
    .Connect = login.conexionreporte
    .Formulas(0) = "desdefecha=""" & cargadesde.Value & """"
    .Formulas(1) = "hastafecha=""" & cargahasta.Value & """"
    .Formulas(2) = "empresa=""" & login.nomempresa & """"
    .Formulas(3) = "saldocero=""" & saldomuestra & """"
Rem    .SubreportToChange = .GetNthSubreportName(0)
Rem    .Connect = login.conexionreporte
Rem    .SubreportToChange = ""
Rem    .Connect = login.conexionreporte
    .DiscardSavedData = True
    .RetrieveDataFiles
    .ReportSource = 0
    .SQLQuery = tabla
    

    .Destination = crptToWindow
    .PrintFileType = crptCrystal
    .WindowState = crptMaximized
    .Action = 1
     
End With
fuera:


End Sub

Private Sub Form_Load()
Aplicar_skin Me

criterio.ConnectionString = login.conexiontotal
datprove.ConnectionString = login.conexiontotal
Check2.Value = 0
Option2.Value = True

If login.empresaact1 > 0 Then
    empresareal = login.empresaact1
Else
    empresareal = login.empresaact
End If


    datprove.RecordSource = "select proveedores.* from proveedores where empresa = " & empresareal & " ORDER BY razonsocial"
    datprove.Refresh
  criterio.RecordSource = "select empreactiva.* from empreactiva"
  criterio.Refresh

text1(0).Text = login.empresaact
cargadesde = Date - Day(Date) + 1
cargahasta = Date
text1(1).Text = cargadesde
text1(2).Text = cargahasta
combodesde.Text = "1"
combohasta.Text = "z"

End Sub


Private Sub listar_Click()


    text1(1).Text = cargadesde.Value
    text1(2).Text = cargahasta.Value

    Inicio.datauditoria.ConnectionString = login.conexiontotal
    
    Inicio.datauditoria.RecordSource = "select auditoria.* from auditoria"
    Inicio.datauditoria.Refresh
    
    Inicio.datauditoria.Recordset.AddNew
    Inicio.datauditoria.Recordset.Fields("fecha") = Date
    Inicio.datauditoria.Recordset.Fields("hora") = Str(Time)
    Inicio.datauditoria.Recordset.Fields("ventana") = "Impresion E.C. Proveedores"
    Inicio.datauditoria.Recordset.Fields("accion") = "Detalle desde:" + Left(combodesde.Text, 20) + " hasta:" + Left(combohasta.Text, 20)
    Inicio.datauditoria.Recordset.Fields("usuario") = login.usuarioactivo
    Inicio.datauditoria.Recordset.Fields("empresa") = login.empresaact
    Inicio.datauditoria.Recordset.UpdateBatch adAffectCurrent

    Call Command1_Click

End Sub

Private Sub saldos_Click()

    text1(1).Text = cargadesde.Value
    text1(2).Text = cargahasta.Value

    Inicio.datauditoria.ConnectionString = login.conexiontotal
    
    Inicio.datauditoria.RecordSource = "select auditoria.* from auditoria"
    Inicio.datauditoria.Refresh
    
    Inicio.datauditoria.Recordset.AddNew
    Inicio.datauditoria.Recordset.Fields("fecha") = Date
    Inicio.datauditoria.Recordset.Fields("hora") = Str(Time)
    Inicio.datauditoria.Recordset.Fields("ventana") = "Impresion E.C. Proveedores"
    Inicio.datauditoria.Recordset.Fields("accion") = "Detalle resumido desde:" + Left(combodesde.Text, 20) + " hasta:" + Left(combohasta.Text, 20)
    Inicio.datauditoria.Recordset.Fields("usuario") = login.usuarioactivo
    Inicio.datauditoria.Recordset.Fields("empresa") = login.empresaact
    Inicio.datauditoria.Recordset.UpdateBatch adAffectCurrent

    Call Command2_Click
End Sub
