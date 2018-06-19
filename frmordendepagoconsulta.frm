VERSION 5.00
Object = "{F6125AB1-8AB1-11CE-A77F-08002B2F4E98}#2.0#0"; "MSRDC20.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{5734474E-78D3-4254-99B9-C35F31BDF509}#62.0#0"; "vbskpro2.ocx"
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmordendepagoconsulta 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta de Ordenes de pago"
   ClientHeight    =   8670
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11190
   Icon            =   "frmordendepagoconsulta.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8670
   ScaleWidth      =   11190
   Begin VB.ListBox List1 
      Height          =   255
      Left            =   6840
      TabIndex        =   8
      Top             =   0
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Duplicado"
      Enabled         =   0   'False
      Height          =   195
      Index           =   1
      Left            =   2760
      TabIndex        =   7
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&Previsualizar"
      Height          =   735
      Left            =   8040
      Picture         =   "frmordendepagoconsulta.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   240
      UseMaskColor    =   -1  'True
      Width           =   1215
   End
   Begin VB.CommandButton calcula 
      Caption         =   "calcula"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   960
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Bindings        =   "frmordendepagoconsulta.frx":0544
      Height          =   420
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   741
      _Version        =   393216
      MatchEntry      =   -1  'True
      ListField       =   "nrorden"
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
   Begin VB.CommandButton salir 
      Caption         =   "&Cerrar"
      Height          =   735
      Left            =   9360
      Picture         =   "frmordendepagoconsulta.frx":0560
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   240
      UseMaskColor    =   -1  'True
      Width           =   1335
   End
   Begin MSMask.MaskEdBox totalabonan 
      Height          =   255
      Left            =   4560
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   0
      Appearance      =   0
      BackColor       =   -2147483644
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
   Begin MSAdodcLib.Adodc datordendepago 
      Height          =   330
      Left            =   4920
      Top             =   360
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
      Caption         =   "Nº Orden"
      Height          =   975
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3975
      Begin VB.CheckBox Check1 
         Caption         =   "Original"
         Height          =   195
         Index           =   0
         Left            =   2640
         TabIndex        =   6
         Top             =   240
         Width           =   1095
      End
   End
   Begin MSAdodcLib.Adodc databonan 
      Height          =   330
      Left            =   1560
      Top             =   960
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
      Left            =   3240
      Top             =   960
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
   Begin MSAdodcLib.Adodc datbuscaorden 
      Height          =   330
      Left            =   2040
      Top             =   960
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
   Begin Crystal.CrystalReport CrystalReporte 
      Left            =   8160
      Top             =   360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTitle     =   "Orden de Pago"
      PrintFileLinesPerPage=   60
      WindowShowCloseBtn=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin MSRDC.MSRDC reporte 
      Height          =   375
      Left            =   360
      Top             =   960
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
      LcK2            =   $"frmordendepagoconsulta.frx":09A2
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
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   7215
      Left            =   120
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1200
      Width           =   10935
      DisplayGroupTree=   0   'False
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   0   'False
      EnableNavigationControls=   -1  'True
      EnableStopButton=   0   'False
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   0   'False
      EnableProgressControl=   -1  'True
      EnableSearchControl=   0   'False
      EnableRefreshButton=   0   'False
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   0   'False
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   0   'False
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
   End
   Begin VB.Frame Frame2 
      Height          =   1095
      Left            =   7680
      TabIndex        =   10
      Top             =   0
      Width           =   3375
   End
   Begin MSAdodcLib.Adodc criterio 
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
End
Attribute VB_Name = "frmordendepagoconsulta"
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
Dim nomprov(50) As String
Dim saldoactual As Currency
Dim Cuenta As Integer
Dim codprove As Integer
Dim idlibrogrid(50) As Integer
Dim saldolibro(50) As Currency
Dim iniciar As Integer
Public numorden As String


Private Sub Check1_Click(Index As Integer)

    If Check1(0).Value = 1 Then
        Check1(1).Value = 0
        Call Command4_Click
        Exit Sub
    Else
        Check1(1).Value = 1
        Call Command4_Click
        Exit Sub
    End If

End Sub

Private Sub Command4_Click()
On Error GoTo fuera
Dim crApp As New CRAXDRT.Application
Dim crReport As New CRAXDRT.Report
Dim tabla As String
Dim tabla1 As String
Dim ruta As String

If iniciar = 0 Then Exit Sub
criterio.ConnectionString = login.conexiontotal
Command4.Enabled = False

criterio.RecordSource = "select empreactiva.* from empreactiva"
criterio.Refresh

criterio.Recordset.Fields(0) = login.empresaact
criterio.Recordset.UpdateBatch adAffectCurrent

ruta = "\Empresa" + Right(Str(login.empresaact), Len(Str(login.empresaact)) - 1)


reporte.SQL = "SELECT consultaordesnpago.nrorden, FROM { oj (contablesql.dbo.consultaordesnpago consultaordesnpago LEFT OUTER JOIN contablesql.dbo.consultaordendepagoinstrumento consultaordendepagoinstrumento ON consultaordesnpago.nrorden = consultaordendepagoinstrumento.nrorden) LEFT OUTER JOIN contablesql.dbo.ordendepagoasignacion ordendepagoasignacion ON consultaordesnpago.nrorden = ordendepagoasignacion.orden AND consultaordesnpago.empresa = ordendepagoasignacion.empresa} WHERE consultaordesnpago.nrorden= '" & DataCombo1.Text & "' and consultaordesnpago.empresa = " & login.empresaact & " ORDER BY consultaordesnpago.razonsocial ASC, consultaordesnpago.id ASC, consultaordesnpago.comprobante ASC"
tabla = reporte.SQL

If Check1(0) = 1 Then
With CrystalReporte
    .ReportFileName = App.Path & ruta + "\Ordendepago.rpt"
    .Connect = login.conexionreporte
    .SubreportToChange = .GetNthSubreportName(0)
    .Connect = login.conexionreporte
    .SubreportToChange = ""
    .Connect = login.conexionreporte
    .DiscardSavedData = True
    .RetrieveDataFiles
    .ReportSource = 0
    .SQLQuery = tabla
    
    .Destination = crptToFile
    .PrintFileType = crptCrystal
    .PrintFileName = App.Path & "\orden.rpt"
    
    .Action = 1
End With
GoTo paso0
End If

If Check1(0) = 0 Then
With CrystalReporte
    .ReportFileName = App.Path & ruta + "\Ordendepago1.rpt"
    .Connect = login.conexionreporte
    .SubreportToChange = .GetNthSubreportName(0)
    .Connect = login.conexionreporte
    .SubreportToChange = ""
    .Connect = login.conexionreporte
    .DiscardSavedData = True
    .RetrieveDataFiles
    .ReportSource = 0
    .SQLQuery = tabla
    
    .Destination = crptToFile
    .PrintFileType = crptCrystal
    .PrintFileName = App.Path & "\orden.rpt"
    
    .Action = 1
End With
End If

paso0:
Set crReport = crApp.OpenReport(App.Path & "\orden.rpt", 1)
CRViewer1.ReportSource = crReport
CRViewer1.ViewReport
Command4.Enabled = True
fuera:

End Sub

Private Sub DataCombo1_GotFocus()

            iniciar = 1

End Sub

Private Sub DataCombo1_KeyPress(KeyAscii As Integer)
On Error GoTo fuera
    If KeyAscii = 13 Then
        KeyAscii = 0
        List1.ListIndex = DataCombo1.SelectedItem - 1
        iniciar = 1
        Call Command4_Click
    End If
fuera:
End Sub


Private Sub Form_Load()

    Check1(0).Value = 1
    Check1(1).Value = 0

iniciar = 0
databonan.ConnectionString = login.conexiontotal
datbuscaorden.ConnectionString = login.conexiontotal
datinstrumento.ConnectionString = login.conexiontotal
datordendepago.ConnectionString = login.conexiontotal

   datbuscaorden.RecordSource = "select ordendepago.* from ordendepago WHERE ordendepago.empresa = " & login.empresaact & " Order by nrorden"
   datbuscaorden.Refresh
   
        datbuscaorden.Recordset.MoveFirst
        Do While Not datbuscaorden.Recordset.EOF
            List1.AddItem (datbuscaorden.Recordset.Fields("nrorden"))
            datbuscaorden.Recordset.MoveNext
        Loop
  
   databonan.RecordSource = "select ordendepagoabonan.* from ordendepagoabonan WHERE ordendepagoabonan.empresa = 0"
   databonan.Refresh
  
   datinstrumento.RecordSource = "select ordendepagoinstrumento.* from ordendepagoinstrumento WHERE ordendepagoinstrumento.empresa = 0"
   datinstrumento.Refresh
  

End Sub

Private Sub salir_Click()

    Unload Me

End Sub

