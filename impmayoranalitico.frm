VERSION 5.00
Object = "{F6125AB1-8AB1-11CE-A77F-08002B2F4E98}#2.0#0"; "MSRDC20.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{5734474E-78D3-4254-99B9-C35F31BDF509}#62.0#0"; "vbskpro2.ocx"
Object = "{D18BBD1F-82BB-4385-BED3-E9D31A3E361E}#1.0#0"; "KewlButtonz.ocx"
Begin VB.Form impmayoranalitico 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Impresión Libro Mayor Analítico"
   ClientHeight    =   3645
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   6270
   Icon            =   "impmayoranalitico.frx":0000
   LinkTopic       =   "From1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3645
   ScaleWidth      =   6270
   Begin VB.CommandButton verificacuenta 
      Caption         =   "verificacuenta"
      Height          =   255
      Left            =   240
      TabIndex        =   16
      Top             =   3120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   1
      Left            =   3960
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      DataField       =   "cerrado"
      DataSource      =   "criterio"
      Height          =   285
      Index           =   8
      Left            =   4440
      TabIndex        =   10
      Text            =   "Text1"
      Top             =   2760
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox Text1 
      DataField       =   "cuenta2"
      DataSource      =   "criterio"
      Height          =   285
      Index           =   2
      Left            =   3840
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   2760
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Text1 
      DataField       =   "cuenta1"
      DataSource      =   "criterio"
      Height          =   285
      Index           =   1
      Left            =   4440
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   2520
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox Text1 
      DataField       =   "empresa"
      DataSource      =   "criterio"
      Height          =   285
      Index           =   0
      Left            =   3840
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   2520
      Visible         =   0   'False
      Width           =   615
   End
   Begin Crystal.CrystalReport CrystalReporte 
      Left            =   5400
      Top             =   2640
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
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   2760
      Visible         =   0   'False
      Width           =   975
   End
   Begin MSRDC.MSRDC reporte 
      Height          =   375
      Left            =   240
      Top             =   2760
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
      Left            =   480
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
      Height          =   1215
      Left            =   360
      TabIndex        =   9
      Top             =   120
      Width           =   5535
      Begin VB.CommandButton Command6 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   0
         Left            =   2760
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   480
         Width           =   855
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Desde"
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   0
         Left            =   1200
         TabIndex        =   0
         Top             =   480
         Width           =   1215
      End
   End
   Begin MSComCtl2.DTPicker cargahasta 
      Height          =   375
      Left            =   3960
      TabIndex        =   3
      Top             =   1920
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   12648447
      Format          =   65208321
      CurrentDate     =   38415
   End
   Begin MSComCtl2.DTPicker cargadesde 
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Top             =   1920
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   12648447
      Format          =   65208321
      CurrentDate     =   38415
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
      TabIndex        =   11
      Top             =   1440
      Width           =   5535
      Begin VB.CommandButton Command6 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   3
         Left            =   2760
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   480
         Width           =   855
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Desde"
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   480
         Width           =   855
      End
   End
   Begin MSAdodcLib.Adodc criterio 
      Height          =   330
      Left            =   1080
      Top             =   2760
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
      Left            =   600
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
      LcK2            =   $"impmayoranalitico.frx":0442
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
      Left            =   2640
      TabIndex        =   4
      Top             =   2760
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
      MICON           =   "impmayoranalitico.frx":0451
      PICN            =   "impmayoranalitico.frx":046D
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
Attribute VB_Name = "impmayoranalitico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private crApp As New CRAXDRT.Application
Private crReport As New CRAXDRT.Report
Public cuenta1 As String
Public cuenta2 As String
Dim poscuenta As Integer
Dim posicion As Integer


Private Sub Command1_Click()
Dim tabla As String
Dim ruta As String

ruta = "\Empresa" + Right(Str(login.empresaact), Len(Str(login.empresaact)) - 1)

If datlibrodiario.Recordset.EOF = False Then
    reporte.SQL = "SELECT DISTINCT librodiario.empresa, librodiario.fecha, librodiario.nroasiento, librodiario.concepto, librodiario.Debe, librodiario.Haber, librodiario.idcuenta, librodiario.cuentita, librodiario.perinicial, librodiario.perfinal, librodiario.razonsocial, librodiario.Nombrecuenta, librodiario1.debe, librodiario1.haber FROM { oj contablesql.dbo.librodiario librodiario INNER JOIN contablesql.dbo.librodiario1 librodiario1 ON librodiario.idcuenta = librodiario1.idcuenta} WHERE librodiario.cuentita >= '" & impmayoranalitico.cuenta1 & "' and librodiario.cuentita <= '" & impmayoranalitico.cuenta2 & "' and librodiario.empresa = " & login.empresaact & " and librodiario.fecha >= '" & cargadesde.Value & "' and librodiario.fecha <= '" & cargahasta.Value & "' ORDER BY librodiario.cuentita ASC, librodiario.fecha ASC, librodiario.nroasiento ASC"
Else
    reporte.SQL = "SELECT DISTINCT librodiario.empresa, librodiario.fecha, librodiario.nroasiento, librodiario.concepto, librodiario.Debe, librodiario.Haber, librodiario.idcuenta, librodiario.cuentita, librodiario.perinicial, librodiario.perfinal, librodiario.razonsocial, librodiario.Nombrecuenta FROM contablesql.dbo.librodiario librodiario WHERE librodiario.cuentita >= '" & impmayoranalitico.cuenta1 & "' and librodiario.cuentita <= '" & impmayoranalitico.cuenta2 & "' and librodiario.empresa = " & login.empresaact & " and librodiario.fecha >= '" & cargadesde.Value & "' and librodiario.fecha <= '" & cargahasta.Value & "' ORDER BY librodiario.cuentita ASC, librodiario.fecha ASC, librodiario.nroasiento ASC"
End If
tabla = reporte.SQL

With CrystalReporte
If datlibrodiario.Recordset.EOF = False Then
    .ReportFileName = App.Path & ruta + "\mayoranalitico.rpt"
Else
    .ReportFileName = App.Path & ruta + "\mayoranalitico2.rpt"
End If
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

criterio.ConnectionString = login.conexiontotal
datcuentas.ConnectionString = login.conexiontotal
datlibrodiario.ConnectionString = login.conexiontotal

    Inicio.Toolbar1.Visible = True

  datcuentas.RecordSource = "select listacuentas.* from listacuentas WHERE empre = " & login.empresaact & " and inicioper = '" & login.iper & "' and finper = '" & login.fper & "' ORDER BY IDCUENTA"
  datcuentas.Refresh
  
  criterio.RecordSource = "select empreactiva.* from empreactiva"
  criterio.Refresh

Text1(0).Text = login.empresaact

cuenta1 = Text2(0).Text
cuenta2 = Text2(1).Text
cargadesde = login.iper
cargahasta = login.fper

End Sub


Private Sub Form_Unload(Cancel As Integer)
    Inicio.Toolbar1.Visible = False
End Sub

Private Sub listar_Click()

 
    
    criterio.Recordset.Fields(0) = login.empresaact
    criterio.Recordset.Fields(1) = cargadesde.Value
    criterio.Recordset.Fields(4) = Text2(0).Text
    criterio.Recordset.Fields(5) = Text2(1).Text
    criterio.Recordset.Fields(6) = login.iper
    criterio.Recordset.UpdateBatch adAffectCurrent
    criterio.Refresh
    
    
    datlibrodiario.RecordSource = "select librodiario1.* from librodiario1 where librodiario1.cuentita >= '" & impmayoranalitico.cuenta1 & "' and librodiario1.cuentita <= '" & impmayoranalitico.cuenta2 & "'"
    datlibrodiario.Refresh
    
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
On Error Resume Next
    
    Text2(Index).SelStart = 0
    Text2(Index).SelLength = Len(Text2(Index).Text)

    poscuenta = Index
    If ventana.menu = 6 Then
        ventana.menu = 0
        Text2(Index).Text = lista_cuentas.cuentacont
        cuenta1 = Text2(0).Text
        cuenta2 = Text2(1).Text
    End If
    

         
End Sub

Private Sub Text2_KeyPress(Index As Integer, KeyAscii As Integer)
Dim X As Integer
Dim c As String

    If KeyAscii = 13 Then
        KeyAscii = 0
        posicion = Index
        For X = 1 To Len(Text2(Index).Text)
            c = Mid(Text2(Index).Text, X, 1)
            If c = "." Then
                      SendKeys "{tab}", False
                      Exit Sub
            End If
        Next X
        Call verificacuenta_Click
    End If
    
End Sub

Private Sub Text2_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)

    If KeyCode = 114 Then
        ventana.menu = 6
        lista_cuentas.Show
    End If

End Sub

Private Sub verificacuenta_Click()
On Error Resume Next
    If Text1(posicion) = "" Then Exit Sub



    datcuentas.ConnectionString = login.conexiontotal
    datcuentas.RecordSource = "select cuentas.* from cuentas where empre = " & login.empresaact & " and imp = 'S' and [cod contable] = " & Text2(posicion).Text & " and inicioper = '" & login.iper & "'"
    datcuentas.Refresh
    
    If datcuentas.Recordset.EOF = True Then
        MsgBox "No Existe esta cuenta contable", vbCritical, "Verificar"
        Text1(posicion).Text = ""
        Text1(posicion).SetFocus
    End If

    Text2(posicion).Text = datcuentas.Recordset.Fields("idcuenta")
    If posicion = 0 Then
        Text2(1).SetFocus
    Else
        cargadesde.SetFocus
    End If
    cuenta1 = Text2(0).Text
    cuenta2 = Text2(1).Text
     
End Sub
