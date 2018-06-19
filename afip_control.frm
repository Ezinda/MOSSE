VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form afip_control 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Control de Comprobantes Subidos a AFIP"
   ClientHeight    =   5865
   ClientLeft      =   -15
   ClientTop       =   270
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5865
   ScaleWidth      =   11880
   Begin VB.CommandButton Command3 
      Caption         =   "Exportar a Excel"
      Height          =   375
      Left            =   3000
      TabIndex        =   5
      Top             =   1440
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Consultar"
      Height          =   615
      Left            =   3000
      TabIndex        =   3
      Top             =   600
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   3600
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   1200
      TabIndex        =   1
      Top             =   600
      Width           =   1095
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   1200
      TabIndex        =   2
      Top             =   1080
      Width           =   1095
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid msflex 
      Height          =   3735
      Left            =   0
      TabIndex        =   4
      Top             =   1920
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   6588
      _Version        =   393216
      Rows            =   1000
      Cols            =   8
      _NumberOfBands  =   1
      _Band(0).Cols   =   8
   End
   Begin VB.Label Label1 
      Caption         =   "Tipo Comprobante: 1-FA, 6-FB, 3-NCA, 8-NCB,  2-NDA, 7-NDB"
      Height          =   495
      Left            =   120
      TabIndex        =   8
      Top             =   0
      Width           =   3255
   End
   Begin VB.Label Label2 
      Caption         =   "Desde Nro:"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   720
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "Hasta Nro:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1080
      Width           =   855
   End
End
Attribute VB_Name = "afip_control"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim fe As New WSAFIPFE.factura


msflex.Clear
    msflex.TextMatrix(0, 1) = "N°Comp"
    msflex.TextMatrix(0, 2) = "Fecha"
    msflex.TextMatrix(0, 3) = "CUIT"
    msflex.TextMatrix(0, 4) = "Neto"
    msflex.TextMatrix(0, 5) = "Iva"
    msflex.TextMatrix(0, 6) = "Total"
    msflex.TextMatrix(0, 7) = "Cae"
    
    msflex.ColWidth(0) = 50
    For X = 1 To 6
         msflex.ColWidth(X) = 1500
    Next X
    msflex.ColWidth(7) = 2000

xactivar = fe.ActivarLicencia("20102028245", "WSAFIPFE.lic", "servcomsrl@gmail.com", "")


If fe.iniciar(modoFiscal_Fiscal, "20102028245", "mmosse.pfx", "WSAFIPFE.lic") Then
'If fe.iniciar(modoFiscal_Test, "20102028245", "mmosse_test.pfx", "") Then
   
   If fe.f1ObtenerTicketAcceso() Then
   
       PtoVta = 6
'       PtoVta = 5  ' test
       Y = 1
       For X = Text2.Text To Text3.Text
        
       xconsulta = fe.F1CompConsultar(PtoVta, Text1.Text, X)
       If xconsulta <> False Then
        xconsulta2 = fe.F1RespuestaDetalleCae
        xconsulta3 = fe.F1DetalleImpIva
        xconsulta4 = fe.F1DetalleImpNeto
        xconsulta5 = fe.F1DetalleImpTotal
        xconsulta6 = fe.F1DetalleDocNro
        xconsulta7 = fe.F1DetalleCbteFch
        
        msflex.TextMatrix(Y, 1) = Right("0000" + Replace(Str(PtoVta), " ", ""), 4) + "-" + Right("00000000" + Replace(Str(X), " ", ""), 8)
        msflex.TextMatrix(Y, 2) = xconsulta7
        msflex.TextMatrix(Y, 3) = xconsulta6
        msflex.TextMatrix(Y, 4) = xconsulta4
        msflex.TextMatrix(Y, 5) = xconsulta3
        msflex.TextMatrix(Y, 6) = xconsulta5
        msflex.TextMatrix(Y, 7) = xconsulta2
        
        Y = Y + 1
       End If
        
       Next X
     End If
End If



End Sub

' -------------------------------------------------------------------------------------------
' \\ -- Botón para importar datos en un nuevo libro
' -------------------------------------------------------------------------------------------
Private Sub Command3_Click()
    If Exportar_Excel(App.Path & "\libro1.xls", msflex) Then
        MsgBox " Datos exportados en " & App.Path, vbInformation
    End If
End Sub
' -------------------------------------------------------------------------------------------
' \\ -- Función para crear un nuevo libro con el contenido del Grid
' -------------------------------------------------------------------------------------------
Public Function Exportar_Excel(sOutputPath As String, FlexGrid As Object) As Boolean
  
    On Error GoTo Error_Handler
  
    Dim o_Excel     As Object
    Dim o_Libro     As Object
    Dim o_Hoja      As Object
    Dim Fila        As Long
    Dim Columna     As Long
      
    ' -- Crea el objeto Excel, el objeto workBook y el objeto sheet
    Set o_Excel = CreateObject("Excel.Application")
    Set o_Libro = o_Excel.Workbooks.Add
    Set o_Hoja = o_Libro.Worksheets.Add
      
    ' -- Bucle para Exportar los datos
    With FlexGrid
        For Fila = 1 To .Rows - 1
            For Columna = 0 To .Cols - 1
                o_Hoja.Cells(Fila, Columna + 1).Value = .TextMatrix(Fila, Columna)
            Next
        Next
    End With
    o_Libro.Close True, sOutputPath
    ' -- Cerrar Excel
    o_Excel.Quit
    ' -- Terminar instancias
    Call ReleaseObjects(o_Excel, o_Libro, o_Hoja)
    Exportar_Excel = True
Exit Function
  
' -- Controlador de Errores
Error_Handler:
    ' -- Cierra la hoja y el la aplicación Excel
    If Not o_Libro Is Nothing Then: o_Libro.Close False
    If Not o_Excel Is Nothing Then: o_Excel.Quit
    Call ReleaseObjects(o_Excel, o_Libro, o_Hoja)
    If Err.Number <> 1004 Then MsgBox Err.Description, vbCritical
End Function
' -------------------------------------------------------------------
' \\ -- Eliminar objetos para liberar recursos
' -------------------------------------------------------------------
Private Sub ReleaseObjects(o_Excel As Object, o_Libro As Object, o_Hoja As Object)
    If Not o_Excel Is Nothing Then Set o_Excel = Nothing
    If Not o_Libro Is Nothing Then Set o_Libro = Nothing
    If Not o_Hoja Is Nothing Then Set o_Hoja = Nothing
End Sub



Private Sub Form_Load()

    msflex.TextMatrix(0, 1) = "N°Comp"
    msflex.TextMatrix(0, 2) = "Fecha"
    msflex.TextMatrix(0, 3) = "CUIT"
    msflex.TextMatrix(0, 4) = "Neto"
    msflex.TextMatrix(0, 5) = "Iva"
    msflex.TextMatrix(0, 6) = "Total"
    
    msflex.ColWidth(0) = 50
    For X = 1 To 6
         msflex.ColWidth(X) = 1500
    Next X
    
End Sub

