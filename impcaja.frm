VERSION 5.00
Object = "{F6125AB1-8AB1-11CE-A77F-08002B2F4E98}#2.0#0"; "MSRDC20.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{5734474E-78D3-4254-99B9-C35F31BDF509}#62.0#0"; "vbskpro2.ocx"
Begin VB.Form impcaja 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Caja Diaria"
   ClientHeight    =   5310
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   6375
   Icon            =   "impcaja.frx":0000
   LinkTopic       =   "From1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5310
   ScaleWidth      =   6375
   Begin VB.Frame Frame3 
      Caption         =   "Empresas"
      Height          =   1695
      Left            =   360
      TabIndex        =   15
      Top             =   3240
      Width           =   5655
      Begin VB.ListBox List1 
         Height          =   960
         ItemData        =   "impcaja.frx":0442
         Left            =   480
         List            =   "impcaja.frx":0444
         Style           =   1  'Checkbox
         TabIndex        =   16
         Top             =   360
         Width           =   4695
      End
   End
   Begin MSMask.MaskEdBox arqueo 
      Height          =   375
      Left            =   1440
      TabIndex        =   13
      Top             =   2520
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      Format          =   "$  #,##0.00;($  #,##0.00)"
      PromptChar      =   "_"
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
      Index           =   6
      Left            =   720
      TabIndex        =   10
      Text            =   "Desde"
      Top             =   1530
      Width           =   615
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
      Index           =   5
      Left            =   3240
      TabIndex        =   9
      Text            =   "Hasta"
      Top             =   1530
      Width           =   495
   End
   Begin VB.TextBox Text1 
      DataField       =   "cerrado"
      DataSource      =   "criterio"
      Height          =   285
      Index           =   8
      Left            =   4440
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   1320
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton listar 
      Caption         =   "&Pre Visualizar"
      Height          =   375
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2520
      UseMaskColor    =   -1  'True
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      DataField       =   "cuenta2"
      DataSource      =   "criterio"
      Height          =   285
      Index           =   2
      Left            =   3840
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   1320
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Text1 
      DataField       =   "cuenta1"
      DataSource      =   "criterio"
      Height          =   285
      Index           =   1
      Left            =   4440
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   1080
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox Text1 
      DataField       =   "empresa"
      DataSource      =   "criterio"
      Height          =   285
      Index           =   0
      Left            =   3840
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   1080
      Visible         =   0   'False
      Width           =   615
   End
   Begin Crystal.CrystalReport CrystalReporte 
      Left            =   5280
      Top             =   960
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTitle     =   "Mayor Analitico"
      PrintFileLinesPerPage=   60
      WindowShowGroupTree=   -1  'True
      WindowShowCloseBtn=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
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
      Left            =   1320
      TabIndex        =   0
      Top             =   2400
      Visible         =   0   'False
      Width           =   975
   End
   Begin MSRDC.MSRDC reporte 
      Height          =   375
      Left            =   120
      Top             =   1200
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
      Connect         =   "PROVIDER=MSDASQL;dsn=contable;uid=sa;pwd=;database=contablesql;"
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
   Begin MSAdodcLib.Adodc datcuentas 
      Height          =   330
      Left            =   2520
      Top             =   1200
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
      LockType        =   1
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
   Begin VB.Frame Frame1 
      Caption         =   "Fondo"
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
      Height          =   855
      Left            =   720
      TabIndex        =   4
      Top             =   0
      Width           =   5535
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "impcaja.frx":0446
         Height          =   315
         Left            =   600
         TabIndex        =   12
         Top             =   360
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "nombrefondo"
         BoundColumn     =   "id"
         Text            =   "DataCombo1"
      End
   End
   Begin MSComCtl2.DTPicker cargahasta 
      Height          =   375
      Left            =   3840
      TabIndex        =   7
      Top             =   1440
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   12648447
      Format          =   64290817
      CurrentDate     =   38415
   End
   Begin MSComCtl2.DTPicker cargadesde 
      Height          =   375
      Left            =   1560
      TabIndex        =   8
      Top             =   1440
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   12648447
      Format          =   64290817
      CurrentDate     =   38415
   End
   Begin VB.Frame Frame2 
      Caption         =   "Periodo"
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
      Left            =   360
      TabIndex        =   11
      Top             =   960
      Width           =   5535
   End
   Begin MSAdodcLib.Adodc criterio 
      Height          =   330
      Left            =   240
      Top             =   2040
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
   Begin MSAdodcLib.Adodc datlibrodiario 
      Height          =   330
      Left            =   3960
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
      LockType        =   1
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
      Left            =   720
      Top             =   -120
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
      LcK2            =   $"impcaja.frx":045F
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
   Begin MSAdodcLib.Adodc datcaja 
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
      LockType        =   1
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
   Begin MSAdodcLib.Adodc datempresas 
      Height          =   330
      Left            =   0
      Top             =   0
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
      LockType        =   1
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
   Begin MSAdodcLib.Adodc datfiltro 
      Height          =   330
      Left            =   0
      Top             =   600
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
   Begin VB.Label Label1 
      Caption         =   "Arqueo:"
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
      Left            =   480
      TabIndex        =   14
      Top             =   2640
      Width           =   975
   End
End
Attribute VB_Name = "impcaja"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private crApp As New CRAXDRT.Application
Private crReport As New CRAXDRT.Report
Public cuenta1 As String
Dim empre(100) As Integer
Public cuenta2 As String
Dim poscuenta As Integer


Private Sub arqueo_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = 69 Then
        If impcaja.Height < 5500 Then
            impcaja.Height = 5550
        Else
            impcaja.Height = 3570
        End If
    End If

End Sub

Private Sub Command1_Click()
Dim tabla As String
Dim ruta As String

ruta = "\Empresa" + Right(Str(login.empresaact), Len(Str(login.empresaact)) - 1)

If datlibrodiario.Recordset.EOF = False Then
    reporte.SQL = "SELECT DISTINCT caja.empresa, caja.fecha, caja.nroasiento, caja.concepto, caja.Debe, caja.Haber, caja.idcuenta, caja.cuentita, caja.perinicial, caja.perfinal, caja.razonsocial, caja.Nombrecuenta, caja1.debe, caja1.haber FROM { oj contablesql.dbo.caja0 caja INNER JOIN contablesql.dbo.caja1 caja1 ON caja.idcuenta = caja1.idcuenta} WHERE caja.fecha >= '" & cargadesde.Value & "' and caja.fecha <= '" & cargahasta.Value & "' and caja.id = '" & DataCombo1.BoundText & "'  ORDER BY caja.cuentita ASC, caja.fecha ASC, caja.nroasiento ASC"
Else
    reporte.SQL = "SELECT DISTINCT caja.empresa, caja.fecha, caja.nroasiento, caja.concepto, caja.Debe, caja.Haber, caja.idcuenta, caja.cuentita, caja.perinicial, caja.perfinal, caja.razonsocial, caja.Nombrecuenta FROM contablesql.dbo.caja0 caja WHERE caja.fecha >= '" & cargadesde.Value & "' and caja.fecha <= '" & cargahasta.Value & "' and caja.id = '" & DataCombo1.BoundText & "' ORDER BY caja.cuentita ASC, caja.fecha ASC, caja.nroasiento ASC"
End If

If datcaja.Recordset.EOF = True Then
    reporte.SQL = "SELECT caja1.debe, caja1.haber FROM contablesql.dbo.caja1 caja1 WHERE caja1.id = '" & DataCombo1.BoundText & "' ORDER BY caja1.debe ASC"
End If

tabla = reporte.SQL

With CrystalReporte
If datlibrodiario.Recordset.EOF = False Then
  If datcaja.Recordset.EOF = False Then
    .ReportFileName = App.Path & ruta + "\caja.rpt"
  Else
    .ReportFileName = App.Path & ruta + "\caja3.rpt"
  End If
Else
    .ReportFileName = App.Path & ruta + "\caja2.rpt"
End If
    .Formulas(0) = "fecha1=""" & cargadesde.Value & """"
    .Formulas(1) = "fecha2=""" & cargahasta.Value & """"
    .Formulas(2) = "fondo=""" & DataCombo1.Text & """"
    .Formulas(3) = "arqueo=""" & arqueo.Text & """"
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

impcaja.Height = 3570
criterio.ConnectionString = login.conexiontotal
datcuentas.ConnectionString = login.conexiontotal
datlibrodiario.ConnectionString = login.conexiontotal
datcaja.ConnectionString = login.conexiontotal
datempresas.ConnectionString = login.conexiontotal
datfiltro.ConnectionString = login.conexiontotal

    Inicio.Toolbar1.Visible = True

  datempresas.RecordSource = "select usuarioyempresa.* from usuarioyempresa where nomusuario = '" & login.usuarioactivo & "'"
  datempresas.Refresh

datempresas.Recordset.MoveFirst
i = 0
Do While Not datempresas.Recordset.EOF
    List1.AddItem datempresas.Recordset.Fields("razonsocial")
    empre(i) = datempresas.Recordset.Fields("empresa")
    datempresas.Recordset.MoveNext
    i = i + 1
Loop
   List1.Selected(0) = True
  
        
  datcuentas.RecordSource = "select paramresultados.* from paramresultados ORDER BY ID"
  datcuentas.Refresh
  
  If datcuentas.Recordset.EOF = True Then
    MsgBox "No hay fondos parametrizados", vbCritical, "Error"
    Unload Me
    Exit Sub
  End If
  datcuentas.Recordset.MoveFirst
  DataCombo1.Text = datcuentas.Recordset.Fields("nombrefondo")
  
  criterio.RecordSource = "select empreactiva.* from empreactiva"
  criterio.Refresh

text1(0).Text = login.empresaact

cargadesde = Date
cargahasta = Date

End Sub


Private Sub Form_Unload(Cancel As Integer)
    Inicio.Toolbar1.Visible = False
End Sub

Private Sub listar_Click()

 datfiltro.RecordSource = "select filtracaja.* from filtracaja"
datfiltro.Refresh
If datfiltro.Recordset.EOF = False Then
    datfiltro.Recordset.MoveFirst
    Do While Not datfiltro.Recordset.EOF
        datfiltro.Recordset.Delete adAffectCurrent
        datfiltro.Recordset.MoveNext
    Loop
End If

For x = 0 To List1.ListCount - 1
    If List1.Selected(x) = True Then
        datfiltro.Recordset.AddNew
        datfiltro.Recordset.Fields(0) = empre(x)
        datfiltro.Recordset.UpdateBatch adAffectCurrent
    End If
Next x
    
    criterio.Recordset.Fields(0) = login.empresaact
    criterio.Recordset.Fields(1) = cargadesde.Value
    criterio.Recordset.Fields(6) = login.iper
    criterio.Recordset.UpdateBatch adAffectCurrent
    criterio.Refresh
    

 
        
    datlibrodiario.RecordSource = "select caja1.* from caja1 where id >= '" & DataCombo1.BoundText & "' "
    datlibrodiario.Refresh
    
    datcaja.RecordSource = "select caja0.* from caja0 where id >= '" & DataCombo1.BoundText & "' and fecha >= '" & cargadesde.Value & "' and fecha <= '" & cargahasta.Value & "' "
    datcaja.Refresh
    
    
    Inicio.datauditoria.ConnectionString = login.conexiontotal
    
    Inicio.datauditoria.RecordSource = "select auditoria.* from auditoria"
    Inicio.datauditoria.Refresh
    
    Inicio.datauditoria.Recordset.AddNew
    Inicio.datauditoria.Recordset.Fields("fecha") = Date
    Inicio.datauditoria.Recordset.Fields("hora") = Str(Time)
    Inicio.datauditoria.Recordset.Fields("ventana") = "Impresión Libro Mayor Analitico"
    Inicio.datauditoria.Recordset.Fields("accion") = "Imp.Libro Mayor desde: " + Str(cargadesde) + " hasta:" + Str(cargahasta)
    Inicio.datauditoria.Recordset.Fields("usuario") = login.usuarioactivo
    Inicio.datauditoria.Recordset.Fields("empresa") = login.empresaact
    Inicio.datauditoria.Recordset.UpdateBatch adAffectCurrent
    
    
    Call Command1_Click

End Sub



Private Sub Text2_GotFocus(Index As Integer)
    poscuenta = Index
    
    If Index = 0 Then
        Line1.Visible = True
        Line4.Visible = False
    Else
        Line1.Visible = False
        Line4.Visible = True
    End If
        
    DataList2.Visible = True
    DataList2.SetFocus
End Sub
