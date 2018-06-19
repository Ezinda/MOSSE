VERSION 5.00
Object = "{F6125AB1-8AB1-11CE-A77F-08002B2F4E98}#2.0#0"; "MSRDC20.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{5734474E-78D3-4254-99B9-C35F31BDF509}#62.0#0"; "vbskpro2.ocx"
Begin VB.Form impsumasysaldoscc_viejo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Impresión Balance de Sumas y Saldos con Centros de Costo"
   ClientHeight    =   4635
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   6210
   Icon            =   "impsumasysaldoscc_viejo.frx":0000
   LinkTopic       =   "From1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4635
   ScaleWidth      =   6210
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   1
      Left            =   3840
      TabIndex        =   1
      Top             =   480
      Width           =   1215
   End
   Begin MSDataListLib.DataList DataList2 
      Bindings        =   "impsumasysaldoscc_viejo.frx":0442
      Height          =   1620
      Left            =   840
      TabIndex        =   11
      Top             =   840
      Visible         =   0   'False
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   2752
      _Version        =   393216
      MatchEntry      =   -1  'True
      BackColor       =   12640511
      ForeColor       =   -2147483647
      ListField       =   "codigo"
      BoundColumn     =   "idcuenta"
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
      DataField       =   "cerrado"
      DataSource      =   "criterio"
      Height          =   285
      Index           =   8
      Left            =   3120
      TabIndex        =   10
      Text            =   "Text1"
      Top             =   3240
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton listar 
      Caption         =   "&Listar"
      Height          =   375
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4080
      UseMaskColor    =   -1  'True
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      DataField       =   "cuenta2"
      DataSource      =   "criterio"
      Height          =   285
      Index           =   2
      Left            =   2520
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   3240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Text1 
      DataField       =   "cuenta1"
      DataSource      =   "criterio"
      Height          =   285
      Index           =   1
      Left            =   3120
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   3240
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox Text1 
      DataField       =   "empresa"
      DataSource      =   "criterio"
      Height          =   285
      Index           =   0
      Left            =   2520
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   3000
      Visible         =   0   'False
      Width           =   615
   End
   Begin Crystal.CrystalReport CrystalReporte 
      Left            =   5040
      Top             =   3120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTitle     =   "Sumas y Saldos con CC"
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
      Left            =   5040
      TabIndex        =   2
      Top             =   2880
      Visible         =   0   'False
      Width           =   975
   End
   Begin MSRDC.MSRDC reporte 
      Height          =   375
      Left            =   120
      Top             =   3240
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
      Left            =   3240
      TabIndex        =   6
      Text            =   "Hasta"
      Top             =   570
      Width           =   495
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
      Index           =   4
      Left            =   720
      TabIndex        =   7
      Text            =   "Desde"
      Top             =   570
      Width           =   615
   End
   Begin MSAdodcLib.Adodc datcuentas 
      Height          =   330
      Left            =   3720
      Top             =   3480
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
   Begin VB.Frame Frame1 
      Caption         =   "Cuentas a Listar"
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
      TabIndex        =   8
      Top             =   0
      Width           =   5535
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   0
         Left            =   1080
         TabIndex        =   0
         Top             =   480
         Width           =   1215
      End
      Begin VB.Line Line4 
         BorderWidth     =   2
         Visible         =   0   'False
         X1              =   4080
         X2              =   2880
         Y1              =   600
         Y2              =   1080
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         Visible         =   0   'False
         X1              =   1680
         X2              =   2520
         Y1              =   600
         Y2              =   1080
      End
   End
   Begin MSComCtl2.DTPicker cargadesde 
      Height          =   375
      Left            =   1560
      TabIndex        =   15
      Top             =   3000
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   12648447
      Format          =   65142785
      CurrentDate     =   38415
   End
   Begin MSComCtl2.DTPicker cargahasta 
      Height          =   375
      Left            =   3840
      TabIndex        =   14
      Top             =   3000
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   12648447
      Format          =   65142785
      CurrentDate     =   38415
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
      TabIndex        =   12
      Text            =   "Desde"
      Top             =   3090
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
      TabIndex        =   13
      Text            =   "Hasta"
      Top             =   3090
      Width           =   495
   End
   Begin VB.Frame Frame2 
      Caption         =   "Periodo de Fecha"
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
      TabIndex        =   16
      Top             =   2520
      Width           =   5535
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
      LcK2            =   $"impsumasysaldoscc_viejo.frx":045B
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
Attribute VB_Name = "impsumasysaldoscc_viejo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private crApp As New CRAXDRT.Application
Private crReport As New CRAXDRT.Report
Dim poscuenta As Integer
Dim cuentad As String
Dim cuentah As String

Private Sub Command1_Click()
Dim tabla As String
Dim ruta As String

ruta = "\Empresa" + Right(Str(login.empresaact), Len(Str(login.empresaact)) - 1)

reporte.SQL = "SELECT sumasysaldos2.Debe, sumasysaldos2.Haber, sumasysaldos2.idcuenta, sumasysaldos2.cuenta, sumasysaldos2.Nombrecuenta, sumasysaldos2.razonsocial, sumasysaldos2.inicioper, sumasysaldos2.finper, sumasysaldos2.empresa, sumasysaldos2.Fecha FROM contablesql.dbo.sumasysaldos2 sumasysaldos2 WHERE sumasysaldos2.inicioper = '" & login.iper & "' and sumasysaldos2.empresa = " & login.empresaact & " and sumasysaldos2.cuenta >= '" & cuentad & "' and sumasysaldos2.cuenta <= '" & cuentah & "' and sumasysaldos2.Fecha >= '" & cargadesde & "' and sumasysaldos2.Fecha <= '" & cargahasta & "' ORDER BY sumasysaldos2.cuenta ASC"
tabla = reporte.SQL


With CrystalReporte
    .ReportFileName = App.Path & ruta + "\sumasysaldos con cc.rpt"
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


Private Sub DataList2_Click()
    Text2(poscuenta).Text = DataList2.BoundText

End Sub

Private Sub DataList2_GotFocus()
    If Inicio.opcion1 = True Then
        datcuentas.RecordSource = "select listacuentas.* from listacuentas WHERE empre = " & login.empresaact & " and inicioper = '" & login.iper & "' and finper = '" & login.fper & "' ORDER BY IDCUENTA"
        datcuentas.Refresh
        DataList2.ListField = "codigo"
    End If
    If Inicio.opcion2 = True Then
        datcuentas.RecordSource = "select listacuentas.* from listacuentas WHERE empre = " & login.empresaact & " and inicioper = '" & login.iper & "' and finper = '" & login.fper & "' ORDER BY nombre"
        datcuentas.Refresh
        DataList2.ListField = "nombre"
    End If
End Sub

Private Sub DataList2_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        Text2(poscuenta).Text = DataList2.BoundText
            If poscuenta = 1 Then
                listar.SetFocus
                Exit Sub
            End If
            Text2(poscuenta + 1).SetFocus
    End If


End Sub

Private Sub DataList2_LostFocus()

    DataList2.Visible = False
    Line1.Visible = False
    Line4.Visible = False
    
End Sub

Private Sub Form_Load()

Inicio.Toolbar1.Visible = True

datcuentas.ConnectionString = login.conexiontotal

  datcuentas.RecordSource = "select listacuentas.* from listacuentas WHERE empre = " & login.empresaact & " and inicioper = '" & login.iper & "' and finper = '" & login.fper & "' ORDER BY IDCUENTA"
  datcuentas.Refresh

Text1(0).Text = login.empresaact

cuentad = Text2(0).Text
cuentah = Text2(1).Text
cargadesde = login.iper
cargahasta = login.fper

End Sub


Private Sub Form_Unload(Cancel As Integer)

    Inicio.Toolbar1.Visible = False

End Sub

Private Sub listar_Click()


    cuentad = Text2(0).Text
    cuentah = Text2(1).Text
    
    Inicio.datauditoria.ConnectionString = login.conexiontotal
    
    Inicio.datauditoria.RecordSource = "select auditoria.* from auditoria"
    Inicio.datauditoria.Refresh
    
    Inicio.datauditoria.Recordset.AddNew
    Inicio.datauditoria.Recordset.Fields("fecha") = Date
    Inicio.datauditoria.Recordset.Fields("hora") = Str(Time)
    Inicio.datauditoria.Recordset.Fields("ventana") = "Impresión Balance de Sumas y Saldos con CC"
    Inicio.datauditoria.Recordset.Fields("accion") = "Imp.Sumas y Saldos C.C. desde: " + Str(cargadesde) + " hasta:" + Str(cargahasta)
    Inicio.datauditoria.Recordset.Fields("usuario") = login.usuarioactivo
    Inicio.datauditoria.Recordset.Fields("empresa") = login.empresaact
    Inicio.datauditoria.Recordset.UpdateBatch adAffectCurrent
    
    
    Call Command1_Click

End Sub

Private Sub listar_GotFocus()

    Line1.Visible = False
    Line4.Visible = False

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
