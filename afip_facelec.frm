VERSION 5.00
Object = "{F6125AB1-8AB1-11CE-A77F-08002B2F4E98}#2.0#0"; "MSRDC20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{5734474E-78D3-4254-99B9-C35F31BDF509}#62.0#0"; "vbskpro2.ocx"
Object = "{D18BBD1F-82BB-4385-BED3-E9D31A3E361E}#1.0#0"; "KewlButtonz.ocx"
Begin VB.Form afip_facelec 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Genera archivo RECE - Fact.Electrónica"
   ClientHeight    =   4800
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   6210
   Icon            =   "afip_facelec.frx":0000
   LinkTopic       =   "From1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4800
   ScaleWidth      =   6210
   Begin VB.TextBox Text1 
      DataField       =   "cerrado"
      DataSource      =   "criterio"
      Height          =   285
      Index           =   8
      Left            =   0
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   1800
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox Text1 
      DataField       =   "hasta"
      DataSource      =   "criterio"
      Height          =   285
      Index           =   2
      Left            =   0
      TabIndex        =   3
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
      Left            =   0
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   960
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox Text1 
      DataField       =   "empresa"
      DataSource      =   "criterio"
      Height          =   285
      Index           =   0
      Left            =   0
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   1320
      Visible         =   0   'False
      Width           =   615
   End
   Begin Crystal.CrystalReport CrystalReporte 
      Left            =   0
      Top             =   2640
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTitle     =   "Libro IVA Ventas"
      PrintFileLinesPerPage=   60
      WindowShowGroupTree=   -1  'True
      WindowShowCloseBtn=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   2160
      Visible         =   0   'False
      Width           =   855
   End
   Begin MSRDC.MSRDC reporte 
      Height          =   375
      Left            =   0
      Top             =   4800
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
      Height          =   3975
      Left            =   360
      TabIndex        =   4
      Top             =   0
      Width           =   5535
      Begin VB.OptionButton Option2 
         Caption         =   "Automatico"
         Height          =   255
         Left            =   2880
         TabIndex        =   19
         Top             =   3000
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Manual"
         Height          =   255
         Left            =   1320
         TabIndex        =   18
         Top             =   3000
         Width           =   1335
      End
      Begin VB.TextBox Text4 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2040
         TabIndex        =   17
         Top             =   2520
         Width           =   1095
      End
      Begin VB.CommandButton Command6 
         Caption         =   "N°Solic.:"
         Height          =   255
         Index           =   4
         Left            =   960
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   2520
         Width           =   855
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2760
         TabIndex        =   15
         Top             =   2040
         Width           =   1095
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2040
         TabIndex        =   14
         Top             =   1560
         Width           =   1095
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Periodo (aaaamm)"
         Height          =   255
         Index           =   3
         Left            =   960
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   2040
         Width           =   1695
      End
      Begin VB.CommandButton Command6 
         Caption         =   "P.Venta"
         Height          =   255
         Index           =   1
         Left            =   960
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   1560
         Width           =   855
      End
      Begin MSComCtl2.DTPicker desde 
         Height          =   375
         Left            =   2040
         TabIndex        =   10
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   65470465
         CurrentDate     =   39814
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   0
         Left            =   960
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   960
         Width           =   855
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Desde"
         Height          =   255
         Index           =   2
         Left            =   960
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   480
         Width           =   855
      End
      Begin MSComctlLib.ProgressBar bar 
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   3480
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
      Begin MSComCtl2.DTPicker hasta 
         Height          =   375
         Left            =   2040
         TabIndex        =   11
         Top             =   960
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   65470465
         CurrentDate     =   39814
      End
   End
   Begin MSAdodcLib.Adodc datprimaryrs 
      Height          =   330
      Left            =   3960
      Top             =   2760
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
      LcK2            =   $"afip_facelec.frx":0442
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
      TabIndex        =   6
      Top             =   4080
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1085
      BTYPE           =   14
      TX              =   "&Generar"
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
      MICON           =   "afip_facelec.frx":0451
      PICN            =   "afip_facelec.frx":046D
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSAdodcLib.Adodc datbanco 
      Height          =   330
      Left            =   3960
      Top             =   3120
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
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select afip_rece.* from afip_rece where empresa = 0"
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
   Begin MSAdodcLib.Adodc datcomp 
      Height          =   330
      Left            =   4920
      Top             =   2760
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
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select afip_rece_comp.* from afip_rece_comp"
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
Attribute VB_Name = "afip_facelec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private crApp As New CRAXDRT.Application
Private crReport As New CRAXDRT.Report
Dim meslista1 As String

Private Sub Combo1_Click()
Dim meslista0 As String

    meslista0 = Str(Combo1.ListIndex + 1)
    meslista1 = Right(meslista0, Len(meslista0) - 1)
    If Len(meslista1) = 1 Then meslista1 = "0" + meslista1
    
End Sub

Private Sub Command1_Click()
Dim tabla As String
Dim ruta As String

Dim crxapplication As New CRAXDRT.Application
Dim crxreport As CRAXDRT.Report

ruta = "\Empresa" + Right(Str(login.empresaact), Len(Str(login.empresaact)) - 1)

If Combo2.ListIndex < 8 Then reporte.SQL = "SELECT libroventas1.id, libroventas1.empresa,libroventas1.cerrado, libroventas1.fecha, libroventas1.cliente, libroventas1.tipoiva, libroventas1.cuit, libroventas1.tipocompr, libroventas1.numcompr, libroventas1.col1, libroventas1.col2, libroventas1.col3, libroventas1.total, libroventas1.inicioper, libroventas1.finper, libroventas1.nomcol1, libroventas1.nomcol2, libroventas1.nomcol3, libroventas1.razonsocial FROM contablesql.dbo.libroventas1 libroventas1 WHERE libroventas1.cerrado = '" & meslista1 & "' and libroventas1.empresa = " & login.empresaact & " and libroventas1.inicioper = '" & login.iper & "' and libroventas1.tipocompr = '" & Combo2.Text & "' ORDER BY libroventas1.fecha ASC, libroventas1.id ASC"
If Combo2.ListIndex = 8 Then reporte.SQL = "SELECT libroventas1.id, libroventas1.empresa,libroventas1.cerrado, libroventas1.fecha, libroventas1.cliente, libroventas1.tipoiva, libroventas1.cuit, libroventas1.tipocompr, libroventas1.numcompr, libroventas1.col1, libroventas1.col2, libroventas1.col3, libroventas1.total, libroventas1.inicioper, libroventas1.finper, libroventas1.nomcol1, libroventas1.nomcol2, libroventas1.nomcol3, libroventas1.razonsocial FROM contablesql.dbo.libroventas1 libroventas1 WHERE libroventas1.cerrado = '" & meslista1 & "' and libroventas1.empresa = " & login.empresaact & " and libroventas1.inicioper = '" & login.iper & "' and (libroventas1.tipocompr = 'F-A' or libroventas1.tipocompr = 'NCA' or libroventas1.tipocompr = 'R-A') ORDER BY libroventas1.fecha ASC, libroventas1.id ASC"
If Combo2.ListIndex = 9 Then reporte.SQL = "SELECT libroventas1.id, libroventas1.empresa,libroventas1.cerrado, libroventas1.fecha, libroventas1.cliente, libroventas1.tipoiva, libroventas1.cuit, libroventas1.tipocompr, libroventas1.numcompr, libroventas1.col1, libroventas1.col2, libroventas1.col3, libroventas1.total, libroventas1.inicioper, libroventas1.finper, libroventas1.nomcol1, libroventas1.nomcol2, libroventas1.nomcol3, libroventas1.razonsocial FROM contablesql.dbo.libroventas1 libroventas1 WHERE libroventas1.cerrado = '" & meslista1 & "' and libroventas1.empresa = " & login.empresaact & " and libroventas1.inicioper = '" & login.iper & "' and (libroventas1.tipocompr = 'F-B' or libroventas1.tipocompr = 'NCB' or libroventas1.tipocompr = 'R-B') ORDER BY libroventas1.fecha ASC, libroventas1.id ASC"
If Combo2.ListIndex = 10 Then reporte.SQL = "SELECT libroventas1.id, libroventas1.empresa,libroventas1.cerrado, libroventas1.fecha, libroventas1.cliente, libroventas1.tipoiva, libroventas1.cuit, libroventas1.tipocompr, libroventas1.numcompr, libroventas1.col1, libroventas1.col2, libroventas1.col3, libroventas1.total, libroventas1.inicioper, libroventas1.finper, libroventas1.nomcol1, libroventas1.nomcol2, libroventas1.nomcol3, libroventas1.razonsocial FROM contablesql.dbo.libroventas1 libroventas1 WHERE libroventas1.cerrado = '" & meslista1 & "' and libroventas1.empresa = " & login.empresaact & " and libroventas1.inicioper = '" & login.iper & "' ORDER BY libroventas1.fecha ASC, libroventas1.id ASC"
tabla = reporte.SQL


With CrystalReporte
    .ReportFileName = App.Path & ruta + "\libroventas.rpt"
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

    Inicio.datauditoria.ConnectionString = login.conexiontotal
    
    Inicio.datauditoria.RecordSource = "select auditoria.* from auditoria"
    Inicio.datauditoria.Refresh
    
    Inicio.datauditoria.Recordset.AddNew
    Inicio.datauditoria.Recordset.Fields("fecha") = Date
    Inicio.datauditoria.Recordset.Fields("hora") = Str(Time)
    Inicio.datauditoria.Recordset.Fields("ventana") = "Listado Iva Ventas"
    Inicio.datauditoria.Recordset.Fields("accion") = "Listado Iva Ventas:" + Combo1.Text
    Inicio.datauditoria.Recordset.Fields("usuario") = login.usuarioactivo
    Inicio.datauditoria.Recordset.Fields("empresa") = login.empresaact
    Inicio.datauditoria.Recordset.UpdateBatch adAffectCurrent
    
End Sub


Private Sub desde_Change()

Text3.Text = Right(Str(Year(desde.Value)), 4) + Right("0" + Right(Str(Month(desde.Value)), Len(Str(Month(desde.Value))) - 1), 2)

End Sub

Private Sub Form_Load()
Aplicar_skin Me

datprimaryrs.ConnectionString = login.conexiontotal

Text1(0).Text = login.empresaact
desde.Value = Date - Day(Date) + 1
hasta.Value = Date + 30

Text2.Text = login.facturaautomatica
Text4.Text = "0001"
Option1.Value = True


End Sub


Private Sub hasta_Change()


Text3.Text = Right(Str(Year(desde.Value)), 4) + Right("0" + Right(Str(Month(desde.Value)), Len(Str(Month(desde.Value))) - 1), 2)


End Sub

Private Sub listar_Click()
Dim campo(100) As String
Dim ruta As String
Dim Ret As Long
Dim origen As String
Dim destino As String
Dim x As Integer
Dim i  As Integer
Dim Y As Integer

datbanco.RecordSource = "select afip_rece.* from afip_rece where empresa = " & login.empresaact & " and fecha >= '" & desde.Value & "' and fecha <= '" & hasta.Value & "' order by  [06], codigo"
datbanco.Refresh
If datbanco.Recordset.EOF = True Then
    MsgBox "No existe Periodo", vbCritical, "!!Error!!"
    Exit Sub
End If

bar.Min = 0
bar.max = datbanco.Recordset.RecordCount
datbanco.Recordset.MoveFirst
i = 0

For x = 1 To 100
    campo(x) = ""
Next x

Y = 1
Do While Not datbanco.Recordset.EOF
    
    nume = Val(datbanco.Recordset.Fields("06"))
    codi = Val(datbanco.Recordset.Fields("codigo"))
    For x = 2 To 31
        campo(Y) = campo(Y) + datbanco.Recordset.Fields(x)
    Next x
    bar.Value = i
    datbanco.Recordset.MoveNext
    If datbanco.Recordset.EOF = False Then
        difnum = Val(datbanco.Recordset.Fields("06")) - nume
        difcod = Val(datbanco.Recordset.Fields("codigo")) - codi
        If difnum > 1 Or difcod <> 0 Then
            Y = Y + 1
            GoTo fuera
        End If
    End If
    If i < bar.max - 1 Then campo(Y) = campo(Y) + (Chr(13) + Chr(10))
fuera:
    i = i + 1
Loop

For j = 1 To Y
    liqui = Val(Text4.Text) + j - 1
    liqui1 = Right("0000" + Right(Str(liqui), Len(Str(liqui)) - 1), 4)
    If Option2.Value = True Then
        destino = "c:\afip_fe\rece" & Text3.Text & Text2.Text & liqui1 & "00.txt"
    Else
        destino = "c:\afip_fe\manual_rece" & Text3.Text & Text2.Text & liqui1 & "00.txt"
    End If
    Call GuardarArchivo(campo(j), destino)
Next j

bar.Value = 0
MsgBox "Proceso terminado"

 Rem   ruta = App.Path & "\notepad.exe rece.txt"

 Rem   Ret = Shell(ruta, vbNormalFocus)
    
End Sub


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

