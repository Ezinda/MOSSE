VERSION 5.00
Object = "{F6125AB1-8AB1-11CE-A77F-08002B2F4E98}#2.0#0"; "MSRDC20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{D18BBD1F-82BB-4385-BED3-E9D31A3E361E}#1.0#0"; "KewlButtonz.ocx"
Begin VB.Form lista_subdiariocompras 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Subdiario Compras"
   ClientHeight    =   2790
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4575
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2790
   ScaleWidth      =   4575
   Begin VB.Frame Frame1 
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4335
      Begin VB.CommandButton porcentaje 
         Height          =   255
         Left            =   3120
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   1920
         Visible         =   0   'False
         Width           =   975
      End
      Begin MSComctlLib.ProgressBar pbar 
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   1680
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Filtrar"
         Height          =   375
         Left            =   3360
         TabIndex        =   6
         Top             =   0
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Hasta Fecha:"
         Height          =   375
         Left            =   600
         TabIndex        =   5
         Top             =   1080
         Width           =   1335
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Imprimir Reporte"
         Height          =   375
         Left            =   1440
         TabIndex        =   4
         Top             =   2160
         Width           =   1335
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Desde Fecha:"
         Height          =   375
         Left            =   600
         TabIndex        =   3
         Top             =   480
         Width           =   1335
      End
      Begin MSComCtl2.DTPicker DesdeFecha 
         Height          =   375
         Left            =   1920
         TabIndex        =   1
         Top             =   480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   100007937
         CurrentDate     =   42198
      End
      Begin MSComCtl2.DTPicker HastaFecha 
         Height          =   375
         Left            =   1920
         TabIndex        =   2
         Top             =   1080
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   100007937
         CurrentDate     =   42198
      End
      Begin KewlButtonz.KewlButtons salir 
         Height          =   495
         Left            =   0
         TabIndex        =   7
         Top             =   2040
         Visible         =   0   'False
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   873
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
         BCOL            =   -2147483645
         BCOLO           =   -2147483645
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "lista_subdiariocompras.frx":0000
         PICN            =   "lista_subdiariocompras.frx":001C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Crystal.CrystalReport CrystalReporte 
         Left            =   120
         Top             =   1320
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowTitle     =   "Presupusto de Venta"
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         PrinterCollation=   0
         PrintFileLinesPerPage=   60
         WindowShowSearchBtn=   -1  'True
         WindowShowPrintSetupBtn=   -1  'True
         WindowShowRefreshBtn=   -1  'True
      End
      Begin MSRDC.MSRDC reporte 
         Height          =   375
         Left            =   2160
         Top             =   0
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
   End
   Begin MSAdodcLib.Adodc datencabezado 
      Height          =   330
      Left            =   1560
      Top             =   1920
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
   Begin MSAdodcLib.Adodc datitems 
      Height          =   330
      Left            =   240
      Top             =   1920
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
      Caption         =   "datitems"
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
   Begin MSAdodcLib.Adodc datparametros 
      Height          =   330
      Left            =   3000
      Top             =   1920
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   582
      ConnectMode     =   1
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
End
Attribute VB_Name = "lista_subdiariocompras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public cuentacont As String
Dim cuenta(99999) As Integer
Public xbusqueda As String


Private Sub Command7_Click()
'On Error GoTo fuera
Dim crApp As New CRAXDRT.Application
Dim crReport As New CRAXDRT.Report
Dim tabla As String
Dim ruta As String
Set oCmd = New Command

pbar.Min = 0

oCmd.ActiveConnection = login.conexiontotal

oCmd.CommandText = "delete ud_ezi_pos_subdiariocompras_tem "
oCmd.Execute

datparametros.RecordSource = "select * from ud_ezi_pos_subdiariocompras_tem "
datparametros.Refresh

DFECHA = Replace(Str(Year(DesdeFecha.Value)), " ", "") + Right("0" + Replace(Str(Month(DesdeFecha.Value)), " ", ""), 2) + Right("0" + Replace(Str(Day(DesdeFecha.Value)), " ", ""), 2) + "000000000"
hfecha = Replace(Str(Year(HastaFecha.Value + 1)), " ", "") + Right("0" + Replace(Str(Month(HastaFecha.Value + 1)), " ", ""), 2) + Right("0" + Replace(Str(Day(HastaFecha.Value + 1)), " ", ""), 2) + "000000000"

'datencabezado.RecordSource = "SELECT  * FROM V_EZI_LIBROCOMPRAS2        " & _
'                             "WHERE (SUBSTRING(FECHAREGISTRO, 1, 8) >= SUBSTRING('" & DFECHA & "', 1, 8)) AND " & _
'                             "(SUBSTRING(FECHAREGISTRO, 1, 8) <= SUBSTRING('" & hfecha & "', 1, 8)) "
'datencabezado.Refresh

oCmd.CommandText = "insert into ud_ezi_pos_subdiariocompras_tem (ID, TIPO, LETRA, NUMERODOCUMENTO, FECHAACTUAL, COD_CLIENTE, NOM_CLIENTE, TOTAL, PER_IB_BA, PER_IB_CF, PER_NAC, PER_MUN, L_25413, IMP_INT, " & _
                   "IIBB_RB, SIR_IB, PER_IVA, SOBRETASA, CUIT, COD_COND_IVA, NOM_COND_IVA, fecha_emision_formato, COTIZACION, MONEDA, fecharegistro2, FECHAREGISTRO) " & _
                   "SELECT  * FROM V_EZI_LIBROCOMPRAS2        " & _
                             "WHERE (SUBSTRING(FECHAREGISTRO, 1, 8) >= SUBSTRING('" & DFECHA & "', 1, 8)) AND " & _
                             "(SUBSTRING(FECHAREGISTRO, 1, 8) <= SUBSTRING('" & hfecha & "', 1, 8)) "
oCmd.Execute

datparametros.RecordSource = "select * from ud_ezi_pos_subdiariocompras_tem "
datparametros.Refresh

If datparametros.Recordset.EOF = False Then
li = 0
maxli = datparametros.Recordset.RecordCount
pbar.Max = maxli
porcentaje.Caption = "0 %"
porcentaje.Visible = True
        
    
 Do While Not datparametros.Recordset.EOF
    
    oCmd.CommandText = "update ud_ezi_pos_idcompras set idcomprobante = '" & datparametros.Recordset.Fields("id") & "'  from  ud_ezi_pos_idcompras"
    oCmd.Execute
    
    datitems.RecordSource = "SELECT  v_ezi_pos_items_compras.tipo_doc, v_ezi_pos_items_compras.IVA_21 + v_ezi_pos_items_compras.IVACRE_21 AS IVA_21, " & _
                        "         v_ezi_pos_items_compras.IVA_27 + v_ezi_pos_items_compras.IVACRE_27 AS IVA_27 , " & _
                        "         v_ezi_pos_items_compras.IVA_10_5 +v_ezi_pos_items_compras.IVACRE_10_5 AS IVA_10_5 , " & _
                        "         v_ezi_pos_items_compras.IVACRE_21, v_ezi_pos_items_compras.IVACRE_27, v_ezi_pos_items_compras.IVACRE_10_5, " & _
                        "         v_ezi_pos_items_compras.CRE_21, v_ezi_pos_items_compras.CRE_27, v_ezi_pos_items_compras.CRE_10_5, " & _
                        "         v_ezi_pos_items_compras.NETO_NO_GRAVADO_EXENTO, v_ezi_pos_items_compras.NETO_GRAVADO, " & _
                        "         v_ezi_pos_items_compras.NETO_21 + v_ezi_pos_items_compras.CRE_21 AS NETO_21, " & _
                        "         v_ezi_pos_items_compras.NETO_10_5 + v_ezi_pos_items_compras.CRE_10_5 AS NETO_10_5, " & _
                        "         v_ezi_pos_items_compras.NETO_27 + v_ezi_pos_items_compras.CRE_27 AS NETO_27, v_ezi_pos_items_compras.EXENTO " & _
                        "         FROM v_ezi_pos_items_compras WHERE ID = '" & datparametros.Recordset.Fields("id") & "'"
    datitems.Refresh
    
    If datitems.Recordset.EOF = False Then
       For h = 26 To 40
        datparametros.Recordset.Fields(h) = datitems.Recordset.Fields(h - 25)
       Next h
    End If
    
    li = li + 1
    pbar.Value = li
    porcentaje.Caption = Str(Round(pbar.Value / pbar.Max * 100, 0)) + "  %"
    DoEvents
    datparametros.Recordset.MoveNext
    
 Loop
porcentaje.Visible = False
    
End If
'MsgBox "Exportacion terminada", vbInformation, ""

'Exit Sub

 reporte.SQL = "SELECT a.tipo, a.letra, a.numerodocumento, a.fechaactual, a.nom_cliente, a.total, a.PER_IB_BA, a.PER_IB_CF, a.PER_NAC, a.PER_MUN, a.L_25413, a.IMP_INT, a.PER_IVA, a.CUIT, a.COD_COND_IVA, a.IVA_21, a.IVA_27, a.IVA_10_5, a.IVACRE_21, a.IVACRE_27, a.IVACRE_10_5, a.NETO_21, a.NETO_10_5, a.NETO_27 FROM MMOSSE.dbo.ud_ezi_pos_subdiariocompras_tem a"

 tabla = reporte.SQL

With CrystalReporte
    .PrinterCollation = crptCollated
    .ReportFileName = App.Path & "\ReporteSubdiarioCompras.rpt"
    .WindowTitle = "Subdiario Compras"
    .Formulas(0) = "desdefecha=""" & DesdeFecha.Value & """"
    .Formulas(1) = "hastafecha=""" & HastaFecha.Value & """"
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






Private Sub Form_Load()
Aplicar_skin Me

MiFuncionDeAjuste Me, True

yventana = Inicio.Height / 2 - 1000
xventana = Inicio.Width / 2

lista_subdiariocompras.Top = yventana - lista_subdiariocompras.Height / 2
lista_subdiariocompras.Left = xventana - lista_subdiariocompras.Width / 2

DesdeFecha.Value = Date - Day(Date) + 1
HastaFecha.Value = Date

datencabezado.ConnectionString = login.conexiontotal
datitems.ConnectionString = login.conexiontotal
datparametros.ConnectionString = login.conexiontotal


xsuc = login.nomsucursal

End Sub

Private Sub salir_Click()
    
    Unload Me

End Sub

