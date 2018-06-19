VERSION 5.00
Object = "{F6125AB1-8AB1-11CE-A77F-08002B2F4E98}#2.0#0"; "MSRDC20.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{5734474E-78D3-4254-99B9-C35F31BDF509}#62.0#0"; "vbskpro2.ocx"
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Object = "{D18BBD1F-82BB-4385-BED3-E9D31A3E361E}#1.0#0"; "KewlButtonz.ocx"
Begin VB.Form frmfacturaconsulta 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta de Facturas Emitidas"
   ClientHeight    =   8760
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12645
   Icon            =   "frmfacturaconsulta.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8760
   ScaleWidth      =   12645
   Begin VB.OptionButton Option2 
      Caption         =   "Duplicado"
      Height          =   255
      Left            =   6720
      TabIndex        =   13
      Top             =   600
      Width           =   1095
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Original"
      Height          =   255
      Left            =   5400
      TabIndex        =   12
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton despues 
      Height          =   375
      Left            =   6360
      Picture         =   "frmfacturaconsulta.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   960
      Width           =   615
   End
   Begin VB.ListBox List1 
      Height          =   255
      Left            =   6960
      TabIndex        =   7
      Top             =   0
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CommandButton antes 
      Height          =   375
      Left            =   5640
      Picture         =   "frmfacturaconsulta.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   960
      Width           =   615
   End
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   7095
      Left            =   120
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1320
      Width           =   12375
      DisplayGroupTree=   0   'False
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   0   'False
      EnableNavigationControls=   -1  'True
      EnableStopButton=   0   'False
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
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
   Begin VB.CommandButton calcula 
      Caption         =   "calcula"
      Height          =   255
      Left            =   4320
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Bindings        =   "frmfacturaconsulta.frx":0CC6
      Height          =   360
      Left            =   1560
      TabIndex        =   0
      Top             =   480
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      ListField       =   "numcompr"
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSMask.MaskEdBox totalabonan 
      Height          =   255
      Left            =   7440
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   360
      Visible         =   0   'False
      Width           =   615
      _ExtentX        =   1085
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
      Left            =   6000
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
      Caption         =   "Comprobante"
      Height          =   975
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4815
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   360
         TabIndex        =   4
         Top             =   360
         Width           =   1095
      End
   End
   Begin MSAdodcLib.Adodc datbuscaorden 
      Height          =   330
      Left            =   2760
      Top             =   1320
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
      Left            =   6000
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTitle     =   "Orden de Pago"
      PrintFileLinesPerPage=   60
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin MSRDC.MSRDC reporte 
      Height          =   375
      Left            =   1440
      Top             =   1080
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
      LcK2            =   $"frmfacturaconsulta.frx":0CE2
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
   Begin VB.Frame Frame2 
      Height          =   1095
      Left            =   8880
      TabIndex        =   9
      Top             =   120
      Width           =   3615
      Begin KewlButtonz.KewlButtons Command4 
         Height          =   735
         Left            =   360
         TabIndex        =   10
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   1296
         BTYPE           =   14
         TX              =   "&Previsualizar"
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
         MICON           =   "frmfacturaconsulta.frx":0CF1
         PICN            =   "frmfacturaconsulta.frx":0D0D
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
         Left            =   2280
         TabIndex        =   11
         Top             =   240
         Width           =   1095
         _ExtentX        =   1931
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
         MICON           =   "frmfacturaconsulta.frx":40FF
         PICN            =   "frmfacturaconsulta.frx":411B
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
End
Attribute VB_Name = "frmfacturaconsulta"
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
Public numorden As String


Private Sub antes_Click()
On Error GoTo fueraderango

        List1.ListIndex = List1.ListIndex - 1
        DataCombo1.Text = List1.Text
        Call Command4_Click
fueraderango:
Exit Sub
    
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        List1.Clear
        datbuscaorden.RecordSource = "select libroventas.* from libroventas WHERE empresa = " & login.empresaact & " and tipocompr = '" & Combo1.Text & "' order by numcompr"
        datbuscaorden.Refresh
        datbuscaorden.Recordset.MoveFirst
        Do While Not datbuscaorden.Recordset.EOF
            List1.AddItem (datbuscaorden.Recordset.Fields("numcompr"))
            datbuscaorden.Recordset.MoveNext
        Loop
        DataCombo1.Text = ""
        DataCombo1.SetFocus
    End If

End Sub

Private Sub Command4_Click()
On Error Resume Next
Dim crApp As New CRAXDRT.Application
Dim crReport As New CRAXDRT.Report

Dim tabla As String
Dim ruta As String

ruta = "\Empresa" + Right(Str(login.empresaact), Len(Str(login.empresaact)) - 1)

reporte.SQL = "SELECT facturas.fecha, facturas.cliente, facturas.descripcion, facturas.cuit, facturas.tipocompr, facturas.numcompr, facturas.total, facturas.avisador, facturas.producto, facturas.telefono, facturas.contado, facturas.cant, facturas.unidadmedida, facturas.detalle, facturas.preciounit, facturas.totales, facturas.descuento, facturas.totalneto, facturas.impdesc, facturas.domicilio, facturas.localidad, facturas.numdisco, facturas.empresa FROM contablesql.dbo.facturas facturas WHERE facturas.tipocompr = '" & Combo1.Text & "' and facturas.empresa = " & login.empresaact & " and facturas.numcompr = '" & DataCombo1.Text & "' "
tabla = reporte.SQL

If Option1.Value = True Then
    ori = "ORIGINAL"
Else
    ori = "DUPLICADO"
End If

With CrystalReporte
    .PrinterCollation = crptCollated
If Combo1.Text = "F-A" Or Combo1.Text = "NCA" Or Combo1.Text = "NDA" Then
    .ReportFileName = App.Path & ruta + "\FacturaA.rpt"
    .WindowTitle = "Factura A"
Else
    .ReportFileName = App.Path & ruta + "\FacturaB.rpt"
    .WindowTitle = "Factura B"
End If
    .Connect = login.conexionreporte
    .Formulas(0) = "ORI_DUPL=""" & ori & """"
    .DiscardSavedData = True
    .RetrieveDataFiles
    .ReportSource = 0
    .SQLQuery = tabla
    
    .Destination = crptToFile
    .PrintFileType = crptCrystal
    .PrintFileName = App.Path & "\factura.rpt"

    .Action = 1

End With

Set crReport = crApp.OpenReport(App.Path & "\factura.rpt", 1)
CRViewer1.ReportSource = crReport
CRViewer1.ViewReport

End Sub



Private Sub DataCombo1_KeyPress(KeyAscii As Integer)
On Error GoTo fueraderango
    If KeyAscii = 13 Then
        KeyAscii = 0
        List1.ListIndex = DataCombo1.SelectedItem - 1
        Call Command4_Click
    End If
fueraderango:
End Sub


Private Sub despues_Click()

On Error GoTo fueraderango

        List1.ListIndex = List1.ListIndex + 1
        DataCombo1.Text = List1.Text
        Call Command4_Click
fueraderango:
Exit Sub

End Sub

Private Sub Form_Load()
Aplicar_skin Me

frmfacturaconsulta.Top = 0
frmfacturaconsulta.Left = 0

Option1.Value = True
datbuscaorden.ConnectionString = login.conexiontotal
datordendepago.ConnectionString = login.conexiontotal

   Combo1.AddItem ("F-A")
   Combo1.AddItem ("F-B")
   Combo1.AddItem ("NCA")
   Combo1.AddItem ("NCB")
   Combo1.AddItem ("NDA")
   Combo1.AddItem ("NDB")
   Combo1.AddItem ("F-C")
   
   Combo1.Text = "F-A"

   datbuscaorden.RecordSource = "select libroventas.* from libroventas WHERE empresa = " & login.empresaact & " and tipocompr = 'F-A' Order by numcompr"
   datbuscaorden.Refresh
        datbuscaorden.Recordset.MoveFirst
        Do While Not datbuscaorden.Recordset.EOF
            List1.AddItem (datbuscaorden.Recordset.Fields("numcompr"))
            datbuscaorden.Recordset.MoveNext
        Loop

 
End Sub

Private Sub salir_Click()

    Unload Me

End Sub


