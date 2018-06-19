Attribute VB_Name = "Module1"
Public cont As Integer
Public estu(20) As String
Public noactualiza As Integer
Public idpresupuesto As Double
Public tipopresupuesto As String
Public lotecodigo(50, 20) As String
Public lotecantidad(50, 20) As Double
Public loteid(50, 20) As String


' Función que retorna el ancho y alto en pixeles _
 de un texto a partir del dispositivo de contexxto indicado
Public Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" ( _
    ByVal hDC As Long, _
    ByVal lpsz As String, _
    ByVal cbString As Long, _
    lpSize As SIZE) As Long
  
'Estructura para usar con GetTextExtentPoint32 ( Devuelve el Ancho y alto en pixeles )
Public Type SIZE
    cx As Long
    cy As Long
End Type
  
' Obtiene el Hdc a partir del hwnd
Public Declare Function GetDC Lib "User32" (ByVal hWnd As Long) As Long
  
' Función que retorna el Ancho y alto del texto en Pixeles
'**********************************************************************
Public Function Obtener_Text_Size(Objeto As Object, _
                                  Texto As String, _
                                  Ancho As Long, _
                                  Height As Long) As Long
      
    ' Variables para almacenar las caracteristicas _
     de la fuente del objeto contenedor del control
    Dim m_Font As String
    Dim m_FontSize As Integer
    Dim m_FontBold As Boolean
    Dim m_FontItalic As Boolean
    Dim m_FontUnderline As Boolean
    Dim Flag As Boolean
    Dim ret As Long
      
      
    Dim t_Size As SIZE
    Dim m_DC As Long
    Dim m_Hwnd As Long
      
    On Error Resume Next
    ' Obtiene el hdc del objeto parent del control
    m_DC = Objeto.Parent.hDC
    ' Si hay error, por que el objeto no tiene hdc
    If Err.Number Then
        Err.Clear
        '.. entonces prueba a recuperar el hdc a partir del hwnd (m_DC = GetDC(m_Hwnd))
        m_Hwnd = Objeto.Parent.hWnd
        If Err.Number Then
            Err.Clear
            Obtener_Text_Size = True
            Exit Function
        End If
        m_DC = GetDC(m_Hwnd)
    End If
      
    ' Cambia la fuente del objeto parent
    With Objeto.Parent
        m_Font = .FontName
        m_FontSize = .FontSize
        m_FontBold = .FontBold
        m_FontItalic = .FontItalic
        m_FontUnderline = .FontUnderline
        .FontName = Objeto.FontName
        .FontSize = Objeto.FontSize
        .FontBold = Objeto.FontBold
        .FontItalic = Objeto.FontItalic
        .FontUnderline = Objeto.FontUnderline
    End With
      
    ' Le envia el Hdc, el texto, el tamaño del texto, y _
     la estructura t_Size retorna el ancho y alto en pixeles
    ret = GetTextExtentPoint32(m_DC, Texto, Len(Texto), t_Size)
      
    ' asigna los datos
    With t_Size
        Ancho = .cx
        Height = .cy
    End With
  
On Error Resume Next
  
If Flag Then
    'Restaura la fuente del objeto
    With Objeto.Parent
        .FontName = m_Font
        .FontSize = m_FontSize
        .FontUnderline = m_FontUnderline
        .FontBold = m_FontBold
        .FontItalic = m_FontItalic
    End With
End If
      
End Function


Public Sub Aplicar_skin(ByVal Formulario As Form)
   Inicio.Skin1.LoadSkin App.Path & "\mains.skn"
   Inicio.Skin1.ApplySkin Formulario.hWnd
   
   
End Sub


Public Sub Aplicar_skin2(ByVal Formulario As Form)
   Inicio.Skin1.LoadSkin App.Path & "\green.skn"
   Inicio.Skin1.ApplySkin Formulario.hWnd
   
   
End Sub

Public Sub Aplicar_skin3(ByVal Formulario As Form)
   Inicio.Skin1.LoadSkin App.Path & "\corona.skn"
   Inicio.Skin1.ApplySkin Formulario.hWnd
   
   
End Sub

Public Sub PDFCreator_CreatePDF(srcFORM As Form, srcFILE As String, dstFILE As String)

    Dim myPDF As New PdfCreatorObj
    Dim pdfQUEUE As New Queue
    Dim myJOB As PrintJob

    If myPDF.IsInstanceRunning Then
        MsgBox "PDF-Creator is already in use." & vbNewLine & _
               "Wait for session end", _
               vbInformation, _
               “WARNING”
        Exit Sub
    End If

    srcFORM.MousePointer = vbHourglass

    pdfQUEUE.Initialize

    Call myPDF.PrintFile(srcFILE)

    Do Until pdfQUEUE.Count > 0
        DoEvents
    Loop

    Set myJOB = pdfQUEUE.NextJob
    Call myJOB.SetProfileSetting("OpenViewer", "false")
    Call myJOB.ConvertTo(dstFILE)

    Do Until myJOB.IsFinished
        DoEvents
    Loop

    Call pdfQUEUE.ReleaseCom

    Set myJOB = Nothing
    Set pdfQUEUE = Nothing
    Set myPDF = Nothing

    srcFORM.MousePointer = vbNormal

End Sub

