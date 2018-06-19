VERSION 5.00
Object = "{F6125AB1-8AB1-11CE-A77F-08002B2F4E98}#2.0#0"; "MSRDC20.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form afip_f_electronica 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Importación al Sistema de Facturacion Electronica"
   ClientHeight    =   5790
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7620
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5790
   ScaleWidth      =   7620
   Begin VB.Frame Frame2 
      Caption         =   "Imprimir Comprobantes"
      Height          =   1215
      Left            =   240
      TabIndex        =   5
      Top             =   4320
      Width           =   7095
      Begin VB.CommandButton Command1 
         Caption         =   "&Importar"
         Height          =   375
         Left            =   2520
         TabIndex        =   8
         Top             =   720
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Original y Duplicado"
         Height          =   255
         Left            =   3360
         TabIndex        =   7
         Top             =   480
         Width           =   1815
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Solo Original"
         Height          =   255
         Left            =   1680
         TabIndex        =   6
         Top             =   480
         Width           =   1215
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Seleccionar TXT"
      Height          =   615
      Index           =   0
      Left            =   360
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   240
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   2280
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   240
      Width           =   4695
   End
   Begin VB.Frame Frame1 
      Caption         =   "Lista a importar"
      Height          =   3135
      Left            =   240
      TabIndex        =   0
      Top             =   1080
      Width           =   7095
      Begin VB.ListBox List1 
         Height          =   2205
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   6615
      End
      Begin VB.CommandButton aceptar 
         Caption         =   "&Importar"
         Height          =   375
         Left            =   2520
         TabIndex        =   3
         Top             =   2640
         Width           =   2055
      End
   End
   Begin MSAdodcLib.Adodc datventas 
      Height          =   330
      Left            =   120
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
      DataSourceName  =   ""
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
   Begin Crystal.CrystalReport CrystalReporte 
      Left            =   7080
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      WindowTitle     =   "Orden de Pago"
   End
   Begin MSRDC.MSRDC reporte 
      Height          =   375
      Left            =   5640
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
Attribute VB_Name = "afip_f_electronica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim j As Double
Dim tipo(100) As String
Dim tipoimp As String
Dim comprobante As String
Dim punto(100) As String
Dim numero(100) As String
Dim cae(100) As String
Dim fechacae(100) As String

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


Private Sub Command1_Click()
Dim tabla As String
Dim ruta As String

Dim crxapplication As New CRAXDRT.Application
Dim crxreport As CRAXDRT.Report

ruta = "\Empresa" + Right(Str(login.empresaact), Len(Str(login.empresaact)) - 1)

reporte.SQL = "SELECT facturas.fecha, facturas.cliente, facturas.descripcion, facturas.cuit, facturas.tipocompr, facturas.numcompr, facturas.total, facturas.avisador, facturas.producto, facturas.telefono, facturas.contado, facturas.cant, facturas.unidadmedida, facturas.detalle, facturas.preciounit, facturas.totales, facturas.descuento, facturas.totalneto, facturas.impdesc, facturas.domicilio, facturas.localidad, facturas.numdisco, facturas.empresa FROM contablesql.dbo.facturas facturas WHERE facturas.tipocompr = '" & tipoimp & "' and facturas.empresa = " & login.empresaact & " and facturas.numcompr = '" & comprobante & "' "

tabla = reporte.SQL


With CrystalReporte
   If tipoimp = "F-A" Or tipoimp = "NCA" Or tipoimp = "NDA" Then
    .ReportFileName = App.Path & ruta + "\FacturaA.rpt"
   End If
   If tipoimp = "F-B" Or tipoimp = "NCB" Or tipoimp = "NDB" Then
    .ReportFileName = App.Path & ruta + "\FacturaB.rpt"
   End If
    ori = "ORIGINAL"
    .Connect = login.conexionreporte
    .Formulas(0) = "ORI_DUPL=""" & ori & """"
    .DiscardSavedData = True
    .RetrieveDataFiles
    .ReportSource = 0
    .SQLQuery = tabla
    .Destination = crptToPrinter
    .PrintFileType = crptCrystal
    .WindowState = crptMaximized
    .Action = 1
    
If Option1.Value = False Then
    ori = "DUPLICADO"
    .Formulas(0) = "ORI_DUPL=""" & ori & """"
    .DiscardSavedData = True
    .RetrieveDataFiles
    .ReportSource = 0
    .SQLQuery = tabla
    .Destination = crptToPrinter
    .PrintFileType = crptCrystal
    .WindowState = crptMaximized
    .Action = 1
End If

End With

End Sub

Private Sub aceptar_Click()
On Error Resume Next
List1.Clear

If Text1.Text = "" Then Exit Sub


archivoafip = ""
Open Text1.Text For Input As #1
While Not EOF(1)
Line Input #1, file_data$
archivoafip = archivoafip & file_data$
Wend
Close #1

cod_autorizado = Mid(archivoafip, 63, 1)
cod_motivo = Mid(archivoafip, 64, 2)

List1.AddItem "Autorizado: " & cod_autorizado
List1.AddItem "Motivo: " & cod_motivo

archivofinal = Mid(archivoafip, 68, Len(archivoafip) - 68)

lin = 0
For x = 1 To Len(archivofinal)
    If Asc(Mid(archivofinal, x, 1)) = 10 Then lin = lin + 1
    Debug.Print Mid(archivofinal, x + 1, 1)
    If Asc(Mid(archivofinal, x, 1)) = 10 And (Mid(archivofinal, x + 1, 1) = "3" Or Mid(archivofinal, x + 1, 1) = "5") Then
        archivo = Left(archivofinal, x - 1)
    End If
Next x
      
For Y = 1 To lin
    factor = (Y - 1) * 190
    tipo(Y) = Mid(archivo, factor + 10, 2)
    punto(Y) = Mid(archivo, factor + 12, 4)
    numero(Y) = Mid(archivo, factor + 16, 8)
    cae(Y) = Mid(archivo, factor + 136, 14)
    fechacae(Y) = Mid(archivo, factor + 150, 8)
    
    If tipo(Y) = "01" Then tipo(Y) = "F-A"
    If tipo(Y) = "02" Then tipo(Y) = "NDA"
    If tipo(Y) = "03" Then tipo(Y) = "NCA"
    If tipo(Y) = "04" Then tipo(Y) = "R-A"
    If tipo(Y) = "06" Then tipo(Y) = "F-B"
    If tipo(Y) = "07" Then tipo(Y) = "NDB"
    If tipo(Y) = "08" Then tipo(Y) = "NCB"
    If tipo(Y) = "09" Then tipo(Y) = "R-B"

    tipoimp = tipo(Y)
    List1.AddItem tipo(Y) & " " & punto(Y) & "-" & numero(Y) & "    " & cae(Y) & " - " & fechacae(Y)
    
    comprobante = punto(Y) & "-" & numero(Y)
    
    datventas.RecordSource = "select libroventas.* from libroventas where empresa = " & login.empresaact & " and tipocompr = '" & tipo(Y) & "' and numcompr = '" & comprobante & "' "
    datventas.Refresh
    If datventas.Recordset.EOF = True Then
        MsgBox "Error de validacion de factura, la factura no esta cargada en el sistema", vbCritical, "Error"
        Exit Sub
    End If
    
    fechafinal = Right(fechacae(Y), 2) & "/" & Mid(fechacae(Y), 5, 2) & "/" & Left(fechacae(Y), 4)
    datventas.Recordset.Fields("afipfechacae") = fechafinal
    datventas.Recordset.Fields("afipcae") = cae(Y)
    datventas.Recordset.Fields("afipmot") = cod_motivo
    datventas.Recordset.Fields("afipaprov") = cod_autorizado
    datventas.Recordset.UpdateBatch adAffectCurrent
                 
    
    Call Command1_Click
       
    
    
Next Y

MsgBox "Proceso Terminado correctamente", vbInformation, "Ok"



End Sub

Private Sub Command2_Click(Index As Integer)

    elegirarchivo.Show

End Sub


Private Sub Form_Load()
Aplicar_skin Me

    Option1.Value = True

    datventas.ConnectionString = login.conexiontotal
    
    
       
End Sub

