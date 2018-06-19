VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form lista_productos_precios 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Maestro de Articulos"
   ClientHeight    =   7950
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   16185
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7950
   ScaleWidth      =   16185
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3735
      Left            =   12960
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   7
      Top             =   3600
      Width           =   3135
   End
   Begin VB.PictureBox Picture2 
      FillStyle       =   0  'Solid
      Height          =   7815
      Left            =   1080
      ScaleHeight     =   7755
      ScaleWidth      =   10155
      TabIndex        =   6
      Top             =   0
      Width           =   10215
      Begin VB.Image Image3 
         Height          =   7815
         Left            =   0
         Top             =   0
         Width           =   10215
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   2775
      Left            =   12960
      ScaleHeight     =   2715
      ScaleWidth      =   3075
      TabIndex        =   5
      Top             =   600
      Width           =   3135
      Begin VB.Image Image1 
         BorderStyle     =   1  'Fixed Single
         Height          =   2655
         Left            =   0
         Top             =   0
         Width           =   3015
      End
   End
   Begin VB.CommandButton salir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   14040
      TabIndex        =   4
      Top             =   120
      Width           =   1215
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "lista_productos_precios.frx":0000
      Height          =   6735
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Visible         =   0   'False
      Width           =   12735
      _ExtentX        =   22463
      _ExtentY        =   11880
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      HeadLines       =   1
      RowHeight       =   19
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   3
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   540
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   16185
      _ExtentX        =   28549
      _ExtentY        =   953
      ButtonWidth     =   609
      ButtonHeight    =   900
      Appearance      =   1
      Style           =   1
      _Version        =   393216
      Begin VB.CommandButton Command2 
         Caption         =   "&Exporta a Excel"
         Height          =   375
         Left            =   11880
         TabIndex        =   8
         Top             =   120
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1560
         TabIndex        =   3
         Top             =   120
         Width           =   10095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Buscar Artículo:"
         Height          =   375
         Left            =   120
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   120
         Width           =   1335
      End
   End
   Begin MSAdodcLib.Adodc datproducto 
      Height          =   330
      Left            =   120
      Top             =   7320
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
      LockType        =   2
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
   Begin MSAdodcLib.Adodc datimpuestos 
      Height          =   330
      Left            =   1440
      Top             =   7320
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
      LockType        =   2
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
   Begin MSAdodcLib.Adodc datpreciosespeciales 
      Height          =   330
      Left            =   2760
      Top             =   7320
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
      LockType        =   2
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
   Begin MSAdodcLib.Adodc datimagen 
      Height          =   330
      Left            =   4080
      Top             =   7320
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
      LockType        =   2
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
   Begin VB.Image Image2 
      Height          =   4815
      Left            =   2880
      Top             =   1320
      Width           =   9855
   End
End
Attribute VB_Name = "lista_productos_precios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public cuentacont As String
Dim cuenta(99999) As Integer



Private Sub Command2_Click()
    On Error Resume Next
    
If datproducto.Recordset.EOF = False Then
    Dim i   As Integer
    Dim j   As Integer
    
    datproducto.Recordset.MoveFirst
    n_Filas = datproducto.Recordset.RecordCount
    ' -- Colocar el cursor de espera mientras se exportan los datos
    Me.MousePointer = vbHourglass
    
    If n_Filas = 0 Then
        MsgBox "No hay datos para exportar a excel. Se ha indicado 0 en el parámetro Filas ": Exit Sub
    Else
        
   '     Set o_Excel = CreateObject("Excel.Application")
    'Set o_Libro = o_Excel.Workbooks.Add
    'Set o_Hoja = o_Libro.Worksheets.Add
        
        ' -- Crear nueva instancia de Excel
        Set Obj_Excel = CreateObject("Excel.Application")
        ' -- Agregar nuevo libro}
        'xruta = App.Path + "\Clientes.xls"
        'Set Obj_Libro = Obj_Excel.Workbooks.Open(App.Path)
    
        ' -- Referencia a la Hoja activa ( la que añade por defecto Excel )
        Set o_Libro = Obj_Excel.Workbooks.Add
        Set o_Hoja = o_Libro.Worksheets.Add
        Set Obj_Hoja = Obj_Excel.ActiveSheet
   

   
        iCol = 0
        ' --  Recorrer el Datagrid ( Las columnas )
        For i = 0 To DataGrid1.Columns.Count - 1
        '  If i = 0 Or i = 7 Or i = 10 Or i = 12 Or i = 15 Or i = 17 Then GoTo sigue
            If DataGrid1.Columns(i).Visible Then
                ' -- Incrementar índice de columna
                iCol = iCol + 1
                ' -- Obtener el caption de la columna
                Obj_Hoja.Cells(1, iCol) = DataGrid1.Columns(i).Caption
                ' -- Recorrer las filas
                For j = 0 To n_Filas - 1
                    ' -- Asignar el valor a la celda del Excel
                    Obj_Hoja.Cells(j + 2, 1).NumberFormat = "@"
                    Obj_Hoja.Cells(j + 2, 10).NumberFormat = "@"
                    Obj_Hoja.Cells(j + 2, iCol) = _
                    DataGrid1.Columns(i).CellValue(DataGrid1.GetBookmark(j))
                Next
            End If
sigue:
        Next
        
        ' -- Hacer excel visible
        Obj_Excel.Visible = True
        
        ' -- Opcional : colocar en negrita y de color rojo los enbezados en la hoja
        With Obj_Hoja
            .Rows(1).Font.Bold = True
            .Rows(1).Font.Color = vbRed
            ' -- Autoajustar las cabeceras
            .Columns("A:Z").AutoFit
        End With
    End If

    ' -- Eliminar las variables de objeto excel
    Set Obj_Hoja = Nothing
    Set Obj_Libro = Nothing
    Set Obj_Excel = Nothing
    
    ' -- Restaurar cursor
    Me.MousePointer = vbDefault
End If

End Sub

Private Sub DataGrid1_Click()

On Error Resume Next

    datimagen.RecordSource = "SELECT   top 1  V_BODOCLINK.XFILENAME, PRODUCTO.ID AS idproducto " & _
                          "FROM         V_BODOCLINK WITH (nolock) INNER JOIN " & _
                          "PRODUCTO WITH (nolock) ON V_BODOCLINK.LINKEDOBJ_ID = PRODUCTO.ID " & _
                          "WHERE     (RIGHT(V_BODOCLINK.XFILENAME, 3) = 'jpg') and PRODUCTO.ID  = '" & DataGrid1.Columns("id").Text & "' "
    datimagen.Refresh
    
    If datimagen.Recordset.EOF = False Then
        xruta = datimagen.Recordset.Fields("XFILENAME")
        
        
    'Dejamos que la imagen se amplie a su tamaño original
    'Esto es necesario para calcular la proporcion a reducir de la imagen
        Image1.Stretch = False
        Image1.Picture = LoadPicture(xruta)

    'Calculamos la proporcion de lo ancho
        Prop = Picture1.Width / Image1.Width

    'Si la proporcion de lo alto es menor, tomanos esa nueva proporcion
        If (Picture1.Height / Image1.Height) < Prop Then
            Prop = Picture1.Height / Image1.Height
        End If
    
    'Reducimos la imagen con la proporcion calculada
        Image1.Width = Image1.Width * Prop
        Image1.Height = Image1.Height * Prop

    'Opcionalmente podemos centrar la imagen dentro del picture
    'Si no se quiere centrar, comentamos las siguientes 2 lineas
        Image1.Top = (Picture1.Height - Image1.Height) / 2
        Image1.Left = (Picture1.Width - Image1.Width) / 2
    
    'Ajustamos la imagen al control image
        Image1.Stretch = True
        
    Else
        Image1.Picture = LoadPicture("")
    End If
    Text2.Text = ""
    Text2.Text = DataGrid1.Columns("descripcionlarga").Text


End Sub

Private Sub DataGrid1_DblClick()

'            frmnota_venta.Text1(1).Text = DataGrid1.Columns(2).Text
'            SendKeys "{ENTER}", False
'        Unload Me

End Sub



Private Sub DataGrid1_KeyDown(KeyCode As Integer, Shift As Integer)

On Error Resume Next

    datimagen.RecordSource = "SELECT   top 1  V_BODOCLINK.XFILENAME, PRODUCTO.ID AS idproducto " & _
                          "FROM         V_BODOCLINK WITH (nolock) INNER JOIN " & _
                          "PRODUCTO WITH (nolock) ON V_BODOCLINK.LINKEDOBJ_ID = PRODUCTO.ID " & _
                          "WHERE     (RIGHT(V_BODOCLINK.XFILENAME, 3) = 'jpg') and PRODUCTO.ID  = '" & DataGrid1.Columns("id").Text & "' "
    datimagen.Refresh
    
    If datimagen.Recordset.EOF = False Then
        xruta = datimagen.Recordset.Fields("XFILENAME")
        'Dejamos que la imagen se amplie a su tamaño original
    'Esto es necesario para calcular la proporcion a reducir de la imagen
        Image1.Stretch = False
        Image1.Picture = LoadPicture(xruta)

    'Calculamos la proporcion de lo ancho
        Prop = Picture1.Width / Image1.Width

    'Si la proporcion de lo alto es menor, tomanos esa nueva proporcion
        If (Picture1.Height / Image1.Height) < Prop Then
            Prop = Picture1.Height / Image1.Height
        End If
    
    'Reducimos la imagen con la proporcion calculada
        Image1.Width = Image1.Width * Prop
        Image1.Height = Image1.Height * Prop

    'Opcionalmente podemos centrar la imagen dentro del picture
    'Si no se quiere centrar, comentamos las siguientes 2 lineas
        Image1.Top = (Picture1.Height - Image1.Height) / 2
        Image1.Left = (Picture1.Width - Image1.Width) / 2
    
    'Ajustamos la imagen al control image
        Image1.Stretch = True
    Else
        Image1.Picture = LoadPicture("")
    End If
    Text2.Text = ""
    Text2.Text = DataGrid1.Columns("descripcionlarga").Text
    
        

End Sub

Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
On Error Resume Next

If KeyAscii = 13 Then
   If productoconsulta <> "" Then
     mensa = MsgBox("Pega precio en Documento ?", vbYesNo, "Consulta")
     If mensa = vbYes Then
       
        KeyAscii = 0
            If menu = 1 Then

                xprecios = DataGrid1.Columns(6).Text
                
                frmnota_venta.grilla.Row = xfila
                frmnota_venta.grilla.Col = 0
                frmnota_venta.grilla.Text = DataGrid1.Columns("id").Text
                frmnota_venta.grilla.Col = 1
                frmnota_venta.grilla.Text = DataGrid1.Columns("codigo").Text
                frmnota_venta.grilla.Col = 2
                frmnota_venta.grilla.Text = DataGrid1.Columns("producto").Text
                frmnota_venta.grilla.Col = 4
                frmnota_venta.grilla.Text = DataGrid1.Columns("um").Text
                frmnota_venta.grilla.Col = 6
                frmnota_venta.grilla.Text = Format(Round(xprecios, 2), "###,##0.00")
                frmnota_venta.grilla.Col = 14
                frmnota_venta.grilla.Text = Format(Round(xprecios, 2), "###,##0.00")
                frmnota_venta.grilla.Col = 17
                frmnota_venta.grilla.Text = Format(Round(xprecios, 2), "###,##0.00")


                
                datimpuestos.RecordSource = "SELECT     CASE WHEN pni.COEFICIENTEDEFAULT = 0 THEN 21 ELSE pni.COEFICIENTEDEFAULT END AS COEFICIENTEDEFAULT, p.CODIGO, p.ID, " & _
                      "V_UNIDADMEDIDA__1.NOMBRE AS UMVTA, V_UNIDADMEDIDA_.NOMBRE AS UMSTK " & _
                      "FROM         V_PRODUCTO_ AS p INNER JOIN " & _
                      "V_POSICIONADORIMPUESTOS_ AS pi ON p.POSICIONADORIMPUESTOS_ID = pi.ID INNER JOIN " & _
                      "V_ITEMPOSICIONADORIMPUESTOS_ AS ipi ON pi.ITEMSPOSICIONADORIMPUESTOS_ID = ipi.BO_PLACE_ID INNER JOIN " & _
                      "V_POSICIONIMPUESTO_ AS pni ON ipi.POSICIONIMPUESTO_ID = pni.ID INNER JOIN " & _
                      "V_DEFINICIONIMPUESTO_ AS d ON ipi.DEFINICIONIMPUESTO_ID = d.ID INNER JOIN " & _
                      "V_IMPUESTO_ AS i ON d.IMPUESTO_ID = i.ID LEFT OUTER JOIN " & _
                      "V_UNIDADMEDIDA_ ON p.UNIDADMEDIDA_ID = V_UNIDADMEDIDA_.ID LEFT OUTER JOIN " & _
                      "V_UNIDADMEDIDA_ AS V_UNIDADMEDIDA__1 ON p.UNIDADMEDIDANOLINEAL_ID = V_UNIDADMEDIDA__1.ID " & _
                      "WHERE     (i.CODIGO = '010') and p.id = '" & DataGrid1.Columns("id").Text & "'"
                datimpuestos.Refresh
                
                If datimpuestos.Recordset.EOF = True Then
                    xiva = 1.21
                Else
                    xiva = (datimpuestos.Recordset.Fields("coeficientedefault") + 100) / 100
                End If

                                      
                frmnota_venta.grilla.Col = 5
                frmnota_venta.grilla.Text = Format(Round(xprecios / xiva, 2), "###,##0.000")
                frmnota_venta.grilla.Col = 7
                frmnota_venta.grilla.Text = Format(Round(xprecios / xiva, 2), "###,##0.00")
                
                frmnota_venta.grilla.Col = 8
                frmnota_venta.grilla.Text = Format(0, "###,##0.00")
                frmnota_venta.grilla.Col = 9
                frmnota_venta.grilla.Text = Format(0, "###,##0.00")
                frmnota_venta.grilla.Col = 10
                frmnota_venta.grilla.Text = Format(Round(xprecios, 2), "###,##0.00")
                frmnota_venta.grilla.Col = 11
                frmnota_venta.grilla.Text = 1
                frmnota_venta.grilla.Col = 12
                frmnota_venta.grilla.Text = xiva
                
                
                
                frmnota_venta.grilla.Col = 3
                frmnota_venta.grilla.Text = 1
                
               
                frmnota_venta.grilla.SetFocus
                
            End If
            If menu = 2 Then
               xprecios = DataGrid1.Columns(6).Text
            
                frmpresupuesto.grilla.Row = xfila
                frmpresupuesto.grilla.Col = 0
                frmpresupuesto.grilla.Text = DataGrid1.Columns("id").Text
                frmpresupuesto.grilla.Col = 1
                frmpresupuesto.grilla.Text = DataGrid1.Columns("codigo").Text
                frmpresupuesto.grilla.Col = 2
                frmpresupuesto.grilla.Text = DataGrid1.Columns("producto").Text
                frmpresupuesto.grilla.Col = 4
                frmpresupuesto.grilla.Text = DataGrid1.Columns("um").Text
                frmpresupuesto.grilla.Col = 6
                frmpresupuesto.grilla.Text = Format(Round(xprecios, 2), "###,##0.00")
                frmpresupuesto.grilla.Col = 14
                frmpresupuesto.grilla.Text = Format(Round(xprecios, 2), "###,##0.00")
                frmpresupuesto.grilla.Col = 17
                frmpresupuesto.grilla.Text = Format(Round(xprecios, 2), "###,##0.00")

                
                datimpuestos.RecordSource = "SELECT     CASE WHEN pni.COEFICIENTEDEFAULT = 0 THEN 21 ELSE pni.COEFICIENTEDEFAULT END AS COEFICIENTEDEFAULT, p.CODIGO, p.ID, " & _
                      "V_UNIDADMEDIDA__1.NOMBRE AS UMVTA, V_UNIDADMEDIDA_.NOMBRE AS UMSTK " & _
                      "FROM         V_PRODUCTO_ AS p INNER JOIN " & _
                      "V_POSICIONADORIMPUESTOS_ AS pi ON p.POSICIONADORIMPUESTOS_ID = pi.ID INNER JOIN " & _
                      "V_ITEMPOSICIONADORIMPUESTOS_ AS ipi ON pi.ITEMSPOSICIONADORIMPUESTOS_ID = ipi.BO_PLACE_ID INNER JOIN " & _
                      "V_POSICIONIMPUESTO_ AS pni ON ipi.POSICIONIMPUESTO_ID = pni.ID INNER JOIN " & _
                      "V_DEFINICIONIMPUESTO_ AS d ON ipi.DEFINICIONIMPUESTO_ID = d.ID INNER JOIN " & _
                      "V_IMPUESTO_ AS i ON d.IMPUESTO_ID = i.ID LEFT OUTER JOIN " & _
                      "V_UNIDADMEDIDA_ ON p.UNIDADMEDIDA_ID = V_UNIDADMEDIDA_.ID LEFT OUTER JOIN " & _
                      "V_UNIDADMEDIDA_ AS V_UNIDADMEDIDA__1 ON p.UNIDADMEDIDANOLINEAL_ID = V_UNIDADMEDIDA__1.ID " & _
                      "WHERE     (i.CODIGO = '010') and p.id = '" & DataGrid1.Columns("id").Text & "'"
                datimpuestos.Refresh
                
                If datimpuestos.Recordset.EOF = True Then
                    xiva = 1.21
                Else
                    xiva = (datimpuestos.Recordset.Fields("coeficientedefault") + 100) / 100
                End If
                                      
                frmpresupuesto.grilla.Col = 5
                frmpresupuesto.grilla.Text = Format(Round(xprecios / xiva, 2), "###,##0.00")
                frmpresupuesto.grilla.Col = 7
                frmpresupuesto.grilla.Text = Format(Round(xprecios / xiva, 2), "###,##0.00")
                frmpresupuesto.grilla.Col = 8
                frmpresupuesto.grilla.Text = Format(0, "###,##0.00")
                frmpresupuesto.grilla.Col = 9
                frmpresupuesto.grilla.Text = Format(0, "###,##0.00")
                frmpresupuesto.grilla.Col = 10
                frmpresupuesto.grilla.Text = Format(Round(xprecios, 2), "###,##0.00")
                frmpresupuesto.grilla.Col = 11
                frmpresupuesto.grilla.Text = 1
                frmpresupuesto.grilla.Col = 12
                frmpresupuesto.grilla.Text = xiva
                
                
                
                frmpresupuesto.grilla.Col = 3
                frmpresupuesto.grilla.Text = 1
                
                
                frmpresupuesto.grilla.SetFocus
            End If
            If menu = 5 Then
                xprecios = DataGrid1.Columns(6).Text
                
                frmfacctacte_venta.grilla.Row = xfila
                frmfacctacte_venta.grilla.Col = 0
                frmfacctacte_venta.grilla.Text = DataGrid1.Columns("id").Text
                frmfacctacte_venta.grilla.Col = 1
                frmfacctacte_venta.grilla.Text = DataGrid1.Columns("codigo").Text
                frmfacctacte_venta.grilla.Col = 2
                frmfacctacte_venta.grilla.Text = DataGrid1.Columns("producto").Text
                frmfacctacte_venta.grilla.Col = 4
                frmfacctacte_venta.grilla.Text = DataGrid1.Columns("um").Text
                frmfacctacte_venta.grilla.Col = 6
                frmfacctacte_venta.grilla.Text = Format(Round(xprecios, 2), "###,##0.00")
                frmfacctacte_venta.grilla.Col = 14
                frmfacctacte_venta.grilla.Text = Format(Round(xprecios, 2), "###,##0.00")
                frmfacctacte_venta.grilla.Col = 17
                frmfacctacte_venta.grilla.Text = Format(Round(xprecios, 2), "###,##0.00")
                
                datimpuestos.RecordSource = "SELECT     CASE WHEN pni.COEFICIENTEDEFAULT = 0 THEN 21 ELSE pni.COEFICIENTEDEFAULT END AS COEFICIENTEDEFAULT, p.CODIGO, p.ID, " & _
                      "V_UNIDADMEDIDA__1.NOMBRE AS UMVTA, V_UNIDADMEDIDA_.NOMBRE AS UMSTK " & _
                      "FROM         V_PRODUCTO AS p WITH (NOLOCK) INNER JOIN " & _
                      "V_POSICIONADORIMPUESTOS AS pi WITH (NOLOCK) ON p.POSICIONADORIMPUESTOS_ID = pi.ID INNER JOIN " & _
                      "V_ITEMPOSICIONADORIMPUESTOS AS ipi WITH (NOLOCK) ON pi.ITEMSPOSICIONADORIMPUESTOS_ID = ipi.BO_PLACE_ID INNER JOIN " & _
                      "V_POSICIONIMPUESTO AS pni WITH (NOLOCK) ON ipi.POSICIONIMPUESTO_ID = pni.ID INNER JOIN " & _
                      "V_DEFINICIONIMPUESTO AS d WITH (NOLOCK) ON ipi.DEFINICIONIMPUESTO_ID = d.ID INNER JOIN " & _
                      "V_IMPUESTO AS i WITH (NOLOCK) ON d.IMPUESTO_ID = i.ID LEFT OUTER JOIN " & _
                      "V_UNIDADMEDIDA_ ON p.UNIDADMEDIDA_ID = V_UNIDADMEDIDA_.ID LEFT OUTER JOIN " & _
                      "V_UNIDADMEDIDA_ AS V_UNIDADMEDIDA__1 ON p.UNIDADMEDIDANOLINEAL_ID = V_UNIDADMEDIDA__1.ID " & _
                      "WHERE     (i.CODIGO = '010') and p.id = '" & DataGrid1.Columns("id").Text & "'"
                datimpuestos.Refresh
                
                If datimpuestos.Recordset.EOF = True Then
                    xiva = 1.21
                Else
                    xiva = (datimpuestos.Recordset.Fields("coeficientedefault") + 100) / 100
                End If
                                      
               
                frmfacctacte_venta.grilla.Col = 5
                frmfacctacte_venta.grilla.Text = Format(Round(xprecios / xiva, 2), "###,##0.000")
                frmfacctacte_venta.grilla.Col = 7
                frmfacctacte_venta.grilla.Text = Format(Round(xprecios / xiva, 2), "###,##0.00")
                frmfacctacte_venta.grilla.Col = 8
                frmfacctacte_venta.grilla.Text = Format(0, "###,##0.00")
                frmfacctacte_venta.grilla.Col = 9
                frmfacctacte_venta.grilla.Text = Format(0, "###,##0.00")
                frmfacctacte_venta.grilla.Col = 10
                frmfacctacte_venta.grilla.Text = Format(Round(xprecios, 2), "###,##0.00")
                frmfacctacte_venta.grilla.Col = 11
                frmfacctacte_venta.grilla.Text = 1
                frmfacctacte_venta.grilla.Col = 12
                frmfacctacte_venta.grilla.Text = xiva
                
                frmfacctacte_venta.grilla.Col = 3
                frmfacctacte_venta.grilla.Text = 1
                
               
                frmfacctacte_venta.grilla.SetFocus
                
            End If
            Unload Me
     End If
   End If
End If

End Sub





Private Sub DataGrid1_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next

    datimagen.RecordSource = "SELECT   top 1  V_BODOCLINK.XFILENAME, PRODUCTO.ID AS idproducto " & _
                          "FROM         V_BODOCLINK WITH (nolock) INNER JOIN " & _
                          "PRODUCTO WITH (nolock) ON V_BODOCLINK.LINKEDOBJ_ID = PRODUCTO.ID " & _
                          "WHERE     (RIGHT(V_BODOCLINK.XFILENAME, 3) = 'jpg') and PRODUCTO.ID  = '" & DataGrid1.Columns("id").Text & "' "
    datimagen.Refresh
    
    If datimagen.Recordset.EOF = False Then
        xruta = datimagen.Recordset.Fields("XFILENAME")
        'Dejamos que la imagen se amplie a su tamaño original
    'Esto es necesario para calcular la proporcion a reducir de la imagen
        Image1.Stretch = False
        Image1.Picture = LoadPicture(xruta)

    'Calculamos la proporcion de lo ancho
        Prop = Picture1.Width / Image1.Width

    'Si la proporcion de lo alto es menor, tomanos esa nueva proporcion
        If (Picture1.Height / Image1.Height) < Prop Then
            Prop = Picture1.Height / Image1.Height
        End If
    
    'Reducimos la imagen con la proporcion calculada
        Image1.Width = Image1.Width * Prop
        Image1.Height = Image1.Height * Prop

    'Opcionalmente podemos centrar la imagen dentro del picture
    'Si no se quiere centrar, comentamos las siguientes 2 lineas
        Image1.Top = (Picture1.Height - Image1.Height) / 2
        Image1.Left = (Picture1.Width - Image1.Width) / 2
    
    'Ajustamos la imagen al control image
        Image1.Stretch = True
    Else
        Image1.Picture = LoadPicture("")
    End If
    Text2.Text = ""
    Text2.Text = DataGrid1.Columns("descripcionlarga").Text
        


End Sub

Private Sub Form_Activate()

Text1.Text = productoconsulta
Text1.SetFocus
If Text1.Text <> "" Then
    SendKeys "{ENTER}", False
End If


End Sub

Private Sub Form_Load()
If menu = 2 Then
'    Aplicar_skin2 Me
    Aplicar_skin Me
Else
  If frmnota_venta.Visible = True And frmnota_venta.Caption = "PROFORMA DE VENTA" Then
    Aplicar_skin3 Me
  Else
    Aplicar_skin Me
  End If
End If

MiFuncionDeAjuste Me, True

yventana = Inicio.Height / 2 - 1000
xventana = Inicio.Width / 2

Picture2.Visible = False

lista_productos_precios.Top = yventana - lista_productos_precios.Height / 2
lista_productos_precios.Left = xventana - lista_productos_precios.Width / 2


datproducto.ConnectionString = login.conexiontotal
datimpuestos.ConnectionString = login.conexiontotal
datpreciosespeciales.ConnectionString = login.conexiontotal
datimagen.ConnectionString = login.conexiontotal



'datproducto.RecordSource = query
'datproducto.Refresh


'            DataGrid1.Columns(0).Visible = False
'            DataGrid1.Columns(8).Visible = False
'            DataGrid1.Refresh

 
End Sub

Private Sub Image1_Click()

    Picture2.Visible = True
    

'Dejamos que la imagen se amplie a su tamaño original
    'Esto es necesario para calcular la proporcion a reducir de la imagen
        Image3.Stretch = False
        Image3.Picture = Image1.Picture

    'Calculamos la proporcion de lo ancho
        Prop = Picture2.Width / Image2.Width

    'Si la proporcion de lo alto es menor, tomanos esa nueva proporcion
        If (Picture2.Height / Image3.Height) < Prop Then
            Prop = Picture2.Height / Image3.Height
        End If
    
    'Reducimos la imagen con la proporcion calculada
        Image3.Width = Image3.Width * Prop
        Image3.Height = Image3.Height * Prop

    'Opcionalmente podemos centrar la imagen dentro del picture
    'Si no se quiere centrar, comentamos las siguientes 2 lineas
        Image3.Top = (Picture2.Height - Image3.Height) / 2
        Image3.Left = (Picture2.Width - Image3.Width) / 2
    
    'Ajustamos la imagen al control image
        Image3.Stretch = True



End Sub

Private Sub Image3_Click()

    Picture2.Visible = False

End Sub

Private Sub Picture2_Click()

    Picture2.Visible = False

End Sub

Private Sub salir_Click()

    Unload Me

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)

On Error Resume Next
Dim palabra(100) As String

    If KeyAscii = 13 Then
        KeyAscii = 0
        If Text1.Text <> "" Then
           For j = 1 To 100
                palabra(j) = ""
           Next j
           xbusq = ""
           
           Y = 1
           vinicio = 1
            For X = 1 To Len(Text1.Text)
                car = Mid(Text1.Text, X, 1)
                If car = " " Then
                    palabra(Y) = Mid(Text1.Text, vinicio, X - vinicio)
                    Y = Y + 1
                    vinicio = X + 1
                End If
            Next X
            palabra(Y) = Mid(Text1.Text, vinicio, X)
                                        
            xselect = ""
            For h = 1 To Y
               If Y = 1 Then
                    xbusq = "%" + palabra(h) + "%"
               Else
                    If h <> Y Then
                      If h = 1 Then
                        xbusq = xbusq + xselect + "%" + palabra(h) + "%'"
                      Else
                        xbusq = xbusq + xselect + "'%" + palabra(h) + "%'"
                      End If
                        xselect = " AND p.CODIGO + p.DESCRIPCION + r.presentacion + ISNULL(v.DENOMINACION, '') + CASE i.CODIGO WHEN 'I' THEN 'Importado' WHEN 'N' THEN 'Nacional' ELSE '' END + ISNULL(u.DETALLE, '') + ISNULL(r.CODPROVEEDOR, '') + ISNULL(t.CODIGO, '') LIKE "
                    Else
                        xbusq = xbusq + xselect + "'%" + palabra(h) + "%"
                    End If
               End If
             Next h
                    
            xbusqueda = xbusq
            
            xquery1 = "SELECT     p.ID, left(p.CODIGO,charindex(' -- ',p.codigo)) AS Codigo, p.DESCRIPCION AS Producto, t.CODIGO AS Marca, " & _
                      "r.PRESENTACION AS Presentacion, ROUND(CAST(PR.PRECIOCIVA AS decimal(14, 3)), 3) AS Precio, Pr.PrecioCosto, " & _
                      "SUBSTRING(PR.FECHAULTACT, 7, 2) + '/' + SUBSTRING(PR.FECHAULTACT, 5, 2) + '/' + LEFT(PR.FECHAULTACT, 4) AS FechaUltAct, " & _
                      "st.CANTIDAD2_CANTIDAD AS Stock , v.DENOMINACION AS Proveedor, u.DETALLE AS rubro, V_UNIDADMEDIDA.NOMBRE AS UM, " & _
                      "p.CODIGO + isnull(r.PRESENTACION,'') + p.DESCRIPCION + ISNULL(v.DENOMINACION, '') " & _
                      "+ CASE i.CODIGO WHEN 'I' THEN 'Importado' WHEN 'N' THEN 'Nacional' ELSE '' END + ISNULL(u.DETALLE, '') + ISNULL(r.CODPROVEEDOR, '') + ISNULL(t.CODIGO, '') " & _
                      "AS concatenado, isnull(PR.margen,0) as Margen, r.anmat, r.CODPROVEEDOR AS CodProveedor, r.DESCRIPCIONLARGA " & _
                      "FROM         V_PRODUCTO AS p WITH (NOLOCK) LEFT OUTER JOIN v_ezi_pos_stock_global AS st ON p.ID = st.REFERENCIATIPO_ID LEFT OUTER JOIN " & _
                      "V_EZI_PRECIOS_POS AS PR ON p.ID = PR.ID LEFT OUTER JOIN " & _
                      "V_UNIDADMEDIDA WITH (NOLOCK) ON p.UNIDADMEDIDANOLINEAL_ID = V_UNIDADMEDIDA.ID LEFT OUTER JOIN " & _
                      "V_UD_EZI_PRODUCTOS AS r WITH (NOLOCK) ON p.BOEXTENSION_ID = r.ID LEFT OUTER JOIN " & _
                      "V_PROVEEDOR AS v WITH (NOLOCK) ON r.PROVEEDOR_ID = v.ID LEFT OUTER JOIN " & _
                      "V_ITEMTIPOCLASIFICADOR AS i WITH (NOLOCK) ON r.NACIONALIDAD_ID = i.ID LEFT OUTER JOIN " & _
                      "V_ITEMTIPOCLASIFICADOR AS t WITH (NOLOCK) ON r.MARCA_ID = t.ID LEFT OUTER JOIN " & _
                      "V_RUBRO AS u WITH (NOLOCK) ON p.RUBRO_ID = u.ID " & _
                      "Where (p.ACTIVESTATUS = 0) And (p.TIPOOBJETOESTATICO_ID Is Null) and " & _
                      "left(p.CODIGO,charindex(' -- ',p.codigo)) + ' ' + isnull(r.PRESENTACION,'') + ' ' + p.DESCRIPCION+isnull(v.DENOMINACION,'')+ ' ' + CASE i.CODIGO WHEN 'I' THEN 'Importado' WHEN 'N' THEN 'Nacional' ELSE '' END + ' ' + isnull(u.DETALLE,'')+ ' ' + isnull(r.CODPROVEEDOR,'')+ ' ' + isnull(t.CODIGO,'')  like '" & xbusqueda & "' " & _
                      "ORDER BY p.DESCRIPCION"
                      
            datproducto.RecordSource = xquery1
            datproducto.Refresh
            xcuenta = datproducto.Recordset.RecordCount
            DataGrid1.Visible = True
            
            DataGrid1.Columns("id").Visible = False
            DataGrid1.Columns("codproveedor").Visible = False
            DataGrid1.Columns("concatenado").Visible = False
            DataGrid1.Columns("margen").Visible = False
            DataGrid1.Columns(1).Width = 900
            DataGrid1.Columns("producto").Width = 5500
            DataGrid1.Columns("marca").Width = 1200
            DataGrid1.Columns("presentacion").Width = 1200
            DataGrid1.Columns("Precio").Width = 1300
            DataGrid1.Columns("Precio").Alignment = dbgRight
            DataGrid1.Columns("Precio").NumberFormat = "Currency"
            DataGrid1.Columns("PrecioCosto").Visible = False
'            DataGrid1.Columns("PrecioCosto").Alignment = dbgRight
'            DataGrid1.Columns("PrecioCosto").NumberFormat = "Currency"
'            DataGrid1.Columns("PrecioCosto").Width = 1300
            DataGrid1.Columns(7).Width = 1300
            DataGrid1.Columns("Stock").Width = 1300
            DataGrid1.Columns("Stock").Alignment = dbgCenter
            
           
            DataGrid1.Refresh

        End If
        DataGrid1.SetFocus
        SendKeys "{LEFT}", True
        
        
    End If

End Sub
