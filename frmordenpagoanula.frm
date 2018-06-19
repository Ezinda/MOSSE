VERSION 5.00
Object = "{F6125AB1-8AB1-11CE-A77F-08002B2F4E98}#2.0#0"; "MSRDC20.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{5734474E-78D3-4254-99B9-C35F31BDF509}#62.0#0"; "vbskpro2.ocx"
Begin VB.Form frmreciboconsulta 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta de Recibos de Cobranza"
   ClientHeight    =   2100
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4260
   Icon            =   "frmordenpagoanula.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2100
   ScaleWidth      =   4260
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
      Left            =   480
      Picture         =   "frmordenpagoanula.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1200
      UseMaskColor    =   -1  'True
      Width           =   1215
   End
   Begin VB.CommandButton calcula 
      Caption         =   "calcula"
      Height          =   255
      Left            =   1440
      TabIndex        =   4
      Top             =   1320
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Bindings        =   "frmordenpagoanula.frx":0544
      Height          =   360
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   1935
      _ExtentX        =   3413
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
      Left            =   2280
      Picture         =   "frmordenpagoanula.frx":0561
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1200
      UseMaskColor    =   -1  'True
      Width           =   1335
   End
   Begin MSMask.MaskEdBox totalabonan 
      Height          =   255
      Left            =   8280
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   3120
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
      Left            =   480
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
      Connect         =   "Provider=MSDASQL.1;Password=25072004;Persist Security Info=True;User ID=lucva;Data Source=contable;Initial Catalog=contablesql"
      OLEDBString     =   "Provider=MSDASQL.1;Password=25072004;Persist Security Info=True;User ID=lucva;Data Source=contable;Initial Catalog=contablesql"
      OLEDBFile       =   ""
      DataSourceName  =   "contable"
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select ordendepago.* from ordendepago"
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
      Caption         =   "Nº Recibo"
      Height          =   855
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3975
      Begin VB.CheckBox Check1 
         Caption         =   "Original"
         Enabled         =   0   'False
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
      Left            =   1680
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
      Connect         =   "Provider=MSDASQL.1;Password=25072004;Persist Security Info=True;User ID=lucva;Data Source=contable;Initial Catalog=contablesql"
      OLEDBString     =   "Provider=MSDASQL.1;Password=25072004;Persist Security Info=True;User ID=lucva;Data Source=contable;Initial Catalog=contablesql"
      OLEDBFile       =   ""
      DataSourceName  =   "contable"
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select ordendepagoabonan.* from ordendepagoabonan"
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
      Left            =   3000
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
      Connect         =   "Provider=MSDASQL.1;Password=25072004;Persist Security Info=True;User ID=lucva;Data Source=contable;Initial Catalog=contablesql"
      OLEDBString     =   "Provider=MSDASQL.1;Password=25072004;Persist Security Info=True;User ID=lucva;Data Source=contable;Initial Catalog=contablesql"
      OLEDBFile       =   ""
      DataSourceName  =   "contable"
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select ordendepagoinstrumento.* from ordendepagoinstrumento"
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
   Begin MSAdodcLib.Adodc datbuscarecibo 
      Height          =   330
      Left            =   3840
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
      Connect         =   "Provider=MSDASQL.1;Password=25072004;Persist Security Info=True;User ID=lucva;Data Source=contable;Initial Catalog=contablesql"
      OLEDBString     =   "Provider=MSDASQL.1;Password=25072004;Persist Security Info=True;User ID=lucva;Data Source=contable;Initial Catalog=contablesql"
      OLEDBFile       =   ""
      DataSourceName  =   "contable"
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select recibocobro.* from recibocobro"
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
      Left            =   2880
      Top             =   1320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTitle     =   "Orden de Pago"
      PrintFileLinesPerPage=   60
   End
   Begin MSRDC.MSRDC reporte 
      Height          =   375
      Left            =   3240
      Top             =   1320
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
      LcK2            =   $"frmordenpagoanula.frx":09A3
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
End
Attribute VB_Name = "frmreciboconsulta"
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
Dim cuenta As Integer
Dim codprove As Integer
Dim idlibrogrid(50) As Integer
Dim saldolibro(50) As Currency
Public numorden As String

Private Sub Check1_Click(Index As Integer)

    If Check1(0).Value = 1 Then
        Check1(1).Value = 0
        Exit Sub
    Else
        Check1(1).Value = 1
        Exit Sub
    End If

End Sub

Private Sub Command4_Click()
Dim tabla As String
Dim tabla1 As String
Dim ruta As String

ruta = "\Empresa" + Right(Str(login.empresaact), Len(Str(login.empresaact)) - 1)

reporte.SQL = "consultarecibocobro.nrorden, consultarecibocobro.empresa, consultarecibocobro.nomproveedor, consultarecibocobro.comprobante, consultarecibocobro.fechacompro, consultarecibocobro.importe, consultarecibocobro.id, consultarecibocobro.razonsocial, consultarecibocobro.cuit, consultarecibocobro.domicilio, consultarecibocobro.localidad, consultarecibocobro.fecha, consultarecibocobro.domprov, consultarecibocobro.locprov, consultarecibocobro.cuitprov, consultarecibocobro.saldofactura FROM contablesql.dbo.consultarecibocobro consultarecibocobro WHERE consultarecibocobro.nrorden= '" & DataCombo1.Text & "' and consultarecibocobro.empresa = " & login.empresaact & " ORDER BY consultarecibocobro.razonsocial ASC, consultarecibocobro.id ASC"
tabla = reporte.SQL

With CrystalReporte
  
    .ReportFileName = App.Path & ruta + "\Recibocliente.rpt"
    .Connect = "PROVIDER=MSDASQL;dsn=contable;uid=lucva;pwd=25072004;database=contablesql;"
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

Private Sub DataCombo1_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        Call Command4_Click
    End If

End Sub


Private Sub Form_Load()

    Check1(0).Value = 1
    Check1(1).Value = 0

   datbuscarecibo.RecordSource = "select recibocobro.* from recibocobro WHERE recibocobro.empresa = " & login.empresaact & " Order by nrorden"
   datbuscarecibo.Refresh

End Sub

Private Sub salir_Click()

    Unload Me

End Sub

