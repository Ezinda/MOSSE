VERSION 5.00
Object = "{F6125AB1-8AB1-11CE-A77F-08002B2F4E98}#2.0#0"; "MSRDC20.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{D18BBD1F-82BB-4385-BED3-E9D31A3E361E}#1.0#0"; "KewlButtonz.ocx"
Begin VB.Form Informe_Caja_detallado 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informe Detallado de Caja"
   ClientHeight    =   3255
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5220
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   5220
   Begin VB.PictureBox Picture1 
      Height          =   2055
      Left            =   360
      ScaleHeight     =   1995
      ScaleWidth      =   4515
      TabIndex        =   1
      Top             =   240
      Width           =   4575
      Begin VB.CheckBox Check3 
         Caption         =   "Caja Todos los Usuarios"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   4
         Top             =   1200
         Visible         =   0   'False
         Width           =   3375
      End
      Begin MSComCtl2.DTPicker fecha 
         Height          =   375
         Left            =   1800
         TabIndex        =   2
         Top             =   600
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   98500609
         CurrentDate     =   42060
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   360
         TabIndex        =   3
         Top             =   600
         Width           =   1215
      End
   End
   Begin KewlButtonz.KewlButtons KewlButtons2 
      Height          =   495
      Left            =   1560
      TabIndex        =   0
      Top             =   2520
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   873
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
      BCOL            =   -2147483645
      BCOLO           =   -2147483645
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Informe_Caja_detallado.frx":0000
      PICN            =   "Informe_Caja_detallado.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSRDC.MSRDC reporte 
      Height          =   375
      Left            =   3840
      Top             =   2640
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
   Begin Crystal.CrystalReport CrystalReporte 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTitle     =   "Libro IVA Compras"
      ProgressDialog  =   0   'False
      PrintFileLinesPerPage=   60
   End
   Begin MSAdodcLib.Adodc datfiltro 
      Height          =   330
      Left            =   0
      Top             =   2640
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
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
      Caption         =   "datitemsnv"
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
Attribute VB_Name = "Informe_Caja_detallado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Aplicar_skin Me

    fecha.Value = Date
    Check3.Value = 0

End Sub

Private Sub KewlButtons2_Click()
'On Error GoTo fuera
Dim crApp As New CRAXDRT.Application
Dim crReport As New CRAXDRT.Report
Dim tabla As String
Dim ruta As String




    datfiltro.ConnectionString = login.conexiontotal
    datfiltro.RecordSource = "select * from ud_ezi_pos_filtro"
    datfiltro.Refresh
    
    datfiltro.Recordset.Fields("desdefecha") = Right("0" + Replace(Str(Day(fecha.Value)), " ", ""), 2) + "/" + Right("0" + Replace(Str(Month(fecha.Value)), " ", ""), 2) + "/" + Replace(Str(Year(fecha.Value)), " ", "")
    datfiltro.Recordset.Fields("hastafecha") = Right("0" + Replace(Str(Day(fecha.Value)), " ", ""), 2) + "/" + Right("0" + Replace(Str(Month(fecha.Value)), " ", ""), 2) + "/" + Replace(Str(Year(fecha.Value)), " ", "")
    datfiltro.Recordset.Fields("caja") = login.nomsucursal
    datfiltro.Recordset.UpdateBatch adAffectCurrent

xreporte = 0
'If Check3.Value = 0 Then
'    datfiltro.RecordSource = "SELECT     idtransferencia, TIPO, formadepago, Tarjeta, NroCupon, NroCheque, fechadeemision, fechadevencimiento, monto, caja, numeradorinterno, cliente, detalle, fechadelcomprobante , numero, usuario " & _
'                             "FROM         v_ezi_pos_caja_abierta AS v_ezi_pos_caja_abierta Where (usuario Is Null) " & _
'                             "ORDER BY idtransferencia, TIPO, fechadelcomprobante"
 '   datfiltro.Refresh
 '   If datfiltro.Recordset.EOF = False And Date = fecha.Value Then
 '       reporte.SQL = "SELECT v_ezi_pos_caja_abierta.idtransferencia, v_ezi_pos_caja_abierta.TIPO, v_ezi_pos_caja_abierta.formadepago, v_ezi_pos_caja_abierta.Tarjeta, v_ezi_pos_caja_abierta.NroCupon, v_ezi_pos_caja_abierta.NroCheque, v_ezi_pos_caja_abierta.fechadeemision, v_ezi_pos_caja_abierta.fechadevencimiento, v_ezi_pos_caja_abierta.monto, v_ezi_pos_caja_abierta.caja, v_ezi_pos_caja_abierta.numeradorinterno, v_ezi_pos_caja_abierta.cliente, v_ezi_pos_caja_abierta.detalle, v_ezi_pos_caja_abierta.fechadelcomprobante, v_ezi_pos_caja_abierta.numero, v_ezi_pos_caja_abierta.usuario FROM MMOSSE.dbo.v_ezi_pos_caja_abierta v_ezi_pos_caja_abierta " & _
 '                     "WHERE     (usuario IS NULL) " & _
'                      "ORDER BY v_ezi_pos_caja_abierta.idtransferencia ASC, v_ezi_pos_caja_abierta.TIPO ASC, v_ezi_pos_caja_abierta.fechadelcomprobante asc "
'        xreporte = 1
'    Else
'        reporte.SQL = "SELECT v_ezi_pos_caja.idtransferencia, v_ezi_pos_caja.TIPO, v_ezi_pos_caja.formadepago, v_ezi_pos_caja.Tarjeta, v_ezi_pos_caja.NroCupon, v_ezi_pos_caja.NroCheque, v_ezi_pos_caja.fechadeemision, v_ezi_pos_caja.fechadevencimiento, v_ezi_pos_caja.monto, v_ezi_pos_caja.caja, v_ezi_pos_caja.numeradorinterno, v_ezi_pos_caja.cliente, v_ezi_pos_caja.detalle, v_ezi_pos_caja.fechadelcomprobante, v_ezi_pos_caja.numero, v_ezi_pos_caja.usuario FROM MMOSSE.dbo.v_ezi_pos_caja v_ezi_pos_caja " & _
'                      "WHERE     usuario = '" & login.usuarioactivo & "' " & _
'                      "ORDER BY v_ezi_pos_caja.idtransferencia ASC, v_ezi_pos_caja.TIPO ASC, v_ezi_pos_caja.fechadelcomprobante ASC"
'    End If
'Else

   If Date = fecha.Value Then
        reporte.SQL = "SELECT v_ezi_pos_caja.idtransferencia, v_ezi_pos_caja.TIPO, v_ezi_pos_caja.formadepago, v_ezi_pos_caja.Tarjeta, v_ezi_pos_caja.NroCupon, v_ezi_pos_caja.NroCheque, v_ezi_pos_caja.fechadeemision, v_ezi_pos_caja.fechadevencimiento, v_ezi_pos_caja.monto, v_ezi_pos_caja.caja, v_ezi_pos_caja.numeradorinterno, v_ezi_pos_caja.cliente, v_ezi_pos_caja.detalle, v_ezi_pos_caja.fechadelcomprobante, v_ezi_pos_caja.numero, v_ezi_pos_caja.usuario FROM MMOSSE.dbo.v_ezi_pos_caja_total v_ezi_pos_caja ORDER BY v_ezi_pos_caja.idtransferencia ASC, v_ezi_pos_caja.TIPO ASC, v_ezi_pos_caja.fechadelcomprobante asc, v_ezi_pos_caja.caja desc"
   Else
        reporte.SQL = "SELECT v_ezi_pos_caja.idtransferencia, v_ezi_pos_caja.TIPO, v_ezi_pos_caja.formadepago, v_ezi_pos_caja.Tarjeta, v_ezi_pos_caja.NroCupon, v_ezi_pos_caja.NroCheque, v_ezi_pos_caja.fechadeemision, v_ezi_pos_caja.fechadevencimiento, v_ezi_pos_caja.monto, v_ezi_pos_caja.caja, v_ezi_pos_caja.numeradorinterno, v_ezi_pos_caja.cliente, v_ezi_pos_caja.detalle, v_ezi_pos_caja.fechadelcomprobante, v_ezi_pos_caja.numero, v_ezi_pos_caja.usuario FROM MMOSSE.dbo.v_ezi_pos_caja v_ezi_pos_caja ORDER BY v_ezi_pos_caja.idtransferencia ASC, v_ezi_pos_caja.TIPO ASC, v_ezi_pos_caja.fechadelcomprobante asc, v_ezi_pos_caja.caja desc"
   End If

'End If

tabla = reporte.SQL

With CrystalReporte
    .PrinterCollation = crptCollated
    If xreporte = 0 Then
        .ReportFileName = App.Path & "\Reporte_caj_detallado.rpt"
    Else
        .ReportFileName = App.Path & "\Reporte_caj_detallado_abierta.rpt"
    End If
 
    .WindowTitle = "Deporte de Movimientos de Caja"
    .Connect = login.conexionreporte
    .DiscardSavedData = True
    .RetrieveDataFiles
    .ReportSource = 0
    .SQLQuery = tabla
    .Destination = crptToWindow
 '   .Destination = crptToPrinter
    .PrintFileType = crptCrystal
    .WindowState = crptMaximized
    .Action = 1

End With
Exit Sub

fuera:
    
    MsgBox "Reporte no Encontado, o error de configuracion de reporte", vbCritical, "Error"




End Sub
