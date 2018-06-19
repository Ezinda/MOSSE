VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form lista_productos_colon_alq 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Maestro de Articulos para Alquiler"
   ClientHeight    =   7830
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14265
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7830
   ScaleWidth      =   14265
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton salir 
      Cancel          =   -1  'True
      Caption         =   "salir"
      Height          =   375
      Left            =   2520
      TabIndex        =   4
      Top             =   7440
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "lista_productos_colon_alq.frx":0000
      Height          =   6735
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Visible         =   0   'False
      Width           =   14055
      _ExtentX        =   24791
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
      Width           =   14265
      _ExtentX        =   25162
      _ExtentY        =   953
      ButtonWidth     =   609
      ButtonHeight    =   900
      Appearance      =   1
      Style           =   1
      _Version        =   393216
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
   Begin MSAdodcLib.Adodc datimpuestos 
      Height          =   330
      Left            =   1440
      Top             =   7320
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
End
Attribute VB_Name = "lista_productos_colon_alq"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public cuentacont As String
Dim cuenta(99999) As Integer



Private Sub DataGrid1_DblClick()

'            frmalquiler.Text1(1).Text = DataGrid1.Columns(2).Text
'            SendKeys "{ENTER}", False
'        Unload Me

End Sub

Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = 13 Then
        KeyAscii = 0
            If menu = 3 Then
                frmalquiler.grilla.Row = xfila
                frmalquiler.grilla.Col = 0
                frmalquiler.grilla.Text = DataGrid1.Columns("id").Text
                frmalquiler.grilla.Col = 1
                frmalquiler.grilla.Text = DataGrid1.Columns("codigo").Text
                frmalquiler.grilla.Col = 2
                frmalquiler.grilla.Text = DataGrid1.Columns("producto").Text
                frmalquiler.grilla.Col = 4
                frmalquiler.grilla.Text = "Dia"
                frmalquiler.grilla.Col = 6
                frmalquiler.grilla.Text = Format(Round(DataGrid1.Columns("precio").Text, 2), "###,##0.00")
                frmalquiler.grilla.Col = 14
                frmalquiler.grilla.Text = Format(Round(DataGrid1.Columns("precio").Text, 2), "###,##0.00")

                
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
                frmnota_venta.grilla.Text = Format(Round(DataGrid1.Columns("precio").Text / xiva, 2), "###,##0.00")
                frmalquiler.grilla.Col = 7
                frmalquiler.grilla.Text = Format(Round(DataGrid1.Columns("precio").Text / xiva, 2), "###,##0.00")
                frmalquiler.grilla.Col = 8
                frmalquiler.grilla.Text = Format(0, "###,##0.00")
                frmalquiler.grilla.Col = 9
                frmalquiler.grilla.Text = Format(0, "###,##0.00")
                frmalquiler.grilla.Col = 10
                frmalquiler.grilla.Text = Format(Round(DataGrid1.Columns("precio").Text, 2), "###,##0.00")
                frmalquiler.grilla.Col = 11
                frmalquiler.grilla.Text = 1
                frmalquiler.grilla.Col = 12
                frmalquiler.grilla.Text = xiva
                
                
                
                frmalquiler.grilla.Col = 3
                frmalquiler.grilla.Text = 1
                
                frmalquiler.grilla.SetFocus
            End If
            If menu = 1 Then
                frmfacctacte_alquiler.grilla.Row = xfila
                frmfacctacte_alquiler.grilla.Col = 0
                frmfacctacte_alquiler.grilla.Text = DataGrid1.Columns("id").Text
                frmfacctacte_alquiler.grilla.Col = 1
                frmfacctacte_alquiler.grilla.Text = DataGrid1.Columns("codigo").Text
                frmfacctacte_alquiler.grilla.Col = 2
                frmfacctacte_alquiler.grilla.Text = DataGrid1.Columns("producto").Text
                frmfacctacte_alquiler.grilla.Col = 4
                frmfacctacte_alquiler.grilla.Text = "Dia"
                frmfacctacte_alquiler.grilla.Col = 5
                frmfacctacte_alquiler.grilla.Text = Format(Round(DataGrid1.Columns("precio").Text, 2), "###,##0.00")
                frmfacctacte_alquiler.grilla.Col = 13
                frmfacctacte_alquiler.grilla.Text = Format(Round(DataGrid1.Columns("precio").Text, 2), "###,##0.00")

                
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
                                      
                frmfacctacte_alquiler.grilla.Col = 6
                frmfacctacte_alquiler.grilla.Text = Format(Round(DataGrid1.Columns("precio").Text / xiva, 2), "###,##0.00")
                frmfacctacte_alquiler.grilla.Col = 7
                frmfacctacte_alquiler.grilla.Text = Format(0, "###,##0.00")
                frmfacctacte_alquiler.grilla.Col = 8
                frmfacctacte_alquiler.grilla.Text = Format(0, "###,##0.00")
                frmfacctacte_alquiler.grilla.Col = 9
                frmfacctacte_alquiler.grilla.Text = Format(Round(DataGrid1.Columns("precio").Text, 2), "###,##0.00")
                frmfacctacte_alquiler.grilla.Col = 10
                frmfacctacte_alquiler.grilla.Text = 1
                frmfacctacte_alquiler.grilla.Col = 11
                frmfacctacte_alquiler.grilla.Text = xiva
                
                
                
                frmfacctacte_alquiler.grilla.Col = 3
                frmfacctacte_alquiler.grilla.Text = 1
                
                frmfacctacte_alquiler.grilla.SetFocus
            End If
                        
        Unload Me
    End If

End Sub


Private Sub Form_Activate()

Text1.SetFocus

End Sub

Private Sub Form_Load()
Aplicar_skin Me

MiFuncionDeAjuste Me, True

datproducto.ConnectionString = login.conexiontotal
datimpuestos.ConnectionString = login.conexiontotal
                      
            xquery1 = "SELECT     p.ID, p.CODIGO AS Codigo, p.DESCRIPCION AS Producto, t.CODIGO AS Marca, " & _
                      "CASE i.CODIGO WHEN 'I' THEN 'Importado' WHEN 'N' THEN 'Nacional' ELSE '' END AS Nacionalidad, ROUND(CAST(PR.PRECIOCIVA AS decimal(14, 3)), 3) AS Precio, " & _
                      "SUBSTRING(PR.FECHAULTACT, 7, 2) + '/' + SUBSTRING(PR.FECHAULTACT, 5, 2) + '/' + LEFT(PR.FECHAULTACT, 4) AS FechaUltAct, " & _
                      "r.CODPROVEEDOR AS CodProveedor, v.DENOMINACION AS Proveedor, u.DETALLE AS rubro, V_UNIDADMEDIDA_.NOMBRE AS UM, " & _
                      "p.CODIGO + p.DESCRIPCION + ISNULL(v.DENOMINACION, '') " & _
                      "+ CASE i.CODIGO WHEN 'I' THEN 'Importado' WHEN 'N' THEN 'Nacional' ELSE '' END + ISNULL(u.DETALLE, '') + ISNULL(r.CODPROVEEDOR, '') + ISNULL(t.CODIGO, '') " & _
                      "AS concatenado " & _
                      "FROM         V_PRODUCTO_ AS p LEFT OUTER JOIN " & _
                      "V_EZI_PRECIOS_POS AS PR ON p.ID = PR.ID LEFT OUTER JOIN " & _
                      "V_UNIDADMEDIDA_ ON p.UNIDADMEDIDANOLINEAL_ID = V_UNIDADMEDIDA_.ID LEFT OUTER JOIN " & _
                      "V_UD_EZI_PRODUCTOS_ AS r ON p.BOEXTENSION_ID = r.ID LEFT OUTER JOIN " & _
                      "V_PROVEEDOR_ AS v ON r.PROVEEDOR_ID = v.ID LEFT OUTER JOIN " & _
                      "V_ITEMTIPOCLASIFICADOR_ AS i ON r.NACIONALIDAD_ID = i.ID LEFT OUTER JOIN " & _
                      "V_ITEMTIPOCLASIFICADOR_ AS t ON r.MARCA_ID = t.ID LEFT OUTER JOIN " & _
                      "V_RUBRO_ AS u ON p.RUBRO_ID = u.ID " & _
                      "Where (p.ACTIVESTATUS <> 2) And (p.TIPOOBJETOESTATICO_ID Is Null) and (u.DETALLE LIKE '%alquiler%') " & _
                      "ORDER BY p.DESCRIPCION"

datproducto.RecordSource = xquery1
datproducto.Refresh


            DataGrid1.Visible = True
            
            DataGrid1.Columns("id").Visible = False
            DataGrid1.Columns("concatenado").Visible = False
            DataGrid1.Columns(1).Width = 900
            DataGrid1.Columns(2).Width = 6500
            DataGrid1.Columns(3).Width = 1200
            DataGrid1.Columns(4).Width = 300
            DataGrid1.Columns(5).Width = 1300
            DataGrid1.Columns(5).Alignment = dbgRight
            DataGrid1.Columns(5).NumberFormat = "Currency"
            DataGrid1.Columns(6).Width = 1200
            DataGrid1.Columns(7).Width = 1300
 
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
                        xselect = " AND p.CODIGO + p.DESCRIPCION + ISNULL(v.DENOMINACION, '') + CASE i.CODIGO WHEN 'I' THEN 'Importado' WHEN 'N' THEN 'Nacional' ELSE '' END + ISNULL(u.DETALLE, '') + ISNULL(r.CODPROVEEDOR, '') + ISNULL(t.CODIGO, '') LIKE "
                    Else
                        xbusq = xbusq + xselect + "'%" + palabra(h) + "%"
                    End If
               End If
             Next h
                    
            xbusqueda = xbusq
            
            xquery1 = "SELECT     p.ID, p.CODIGO AS Codigo, p.DESCRIPCION AS Producto, t.CODIGO AS Marca, " & _
                      "CASE i.CODIGO WHEN 'I' THEN 'Importado' WHEN 'N' THEN 'Nacional' ELSE '' END AS Nacionalidad, ROUND(CAST(PR.PRECIOCIVA AS decimal(14, 3)), 3) AS Precio, " & _
                      "SUBSTRING(PR.FECHAULTACT, 7, 2) + '/' + SUBSTRING(PR.FECHAULTACT, 5, 2) + '/' + LEFT(PR.FECHAULTACT, 4) AS FechaUltAct, " & _
                      "r.CODPROVEEDOR AS CodProveedor, v.DENOMINACION AS Proveedor, u.DETALLE AS rubro, V_UNIDADMEDIDA_.NOMBRE AS UM, " & _
                      "p.CODIGO + p.DESCRIPCION + ISNULL(v.DENOMINACION, '') " & _
                      "+ CASE i.CODIGO WHEN 'I' THEN 'Importado' WHEN 'N' THEN 'Nacional' ELSE '' END + ISNULL(u.DETALLE, '') + ISNULL(r.CODPROVEEDOR, '') + ISNULL(t.CODIGO, '') " & _
                      "AS concatenado " & _
                      "FROM         V_PRODUCTO_ AS p LEFT OUTER JOIN " & _
                      "V_EZI_PRECIOS_POS AS PR ON p.ID = PR.ID LEFT OUTER JOIN " & _
                      "V_UNIDADMEDIDA_ ON p.UNIDADMEDIDANOLINEAL_ID = V_UNIDADMEDIDA_.ID LEFT OUTER JOIN " & _
                      "V_UD_EZI_PRODUCTOS_ AS r ON p.BOEXTENSION_ID = r.ID LEFT OUTER JOIN " & _
                      "V_PROVEEDOR_ AS v ON r.PROVEEDOR_ID = v.ID LEFT OUTER JOIN " & _
                      "V_ITEMTIPOCLASIFICADOR_ AS i ON r.NACIONALIDAD_ID = i.ID LEFT OUTER JOIN " & _
                      "V_ITEMTIPOCLASIFICADOR_ AS t ON r.MARCA_ID = t.ID LEFT OUTER JOIN " & _
                      "V_RUBRO_ AS u ON p.RUBRO_ID = u.ID " & _
                      "Where (p.ACTIVESTATUS <> 2) And (p.TIPOOBJETOESTATICO_ID Is Null) AND (u.DETALLE LIKE '%alquiler%') and " & _
                      "p.CODIGO + ' ' + replace(p.CODIGO,'.','') + ' ' + p.DESCRIPCION+isnull(v.DENOMINACION,'')+ ' ' + CASE i.CODIGO WHEN 'I' THEN 'Importado' WHEN 'N' THEN 'Nacional' ELSE '' END + ' ' + isnull(u.DETALLE,'')+ ' ' + isnull(r.CODPROVEEDOR,'')+ ' ' + isnull(t.CODIGO,'')  like '" & xbusqueda & "' " & _
                      "ORDER BY p.DESCRIPCION"
            
                      
            datproducto.RecordSource = xquery1
            datproducto.Refresh
            xcuenta = datproducto.Recordset.RecordCount
            DataGrid1.Visible = True
            
            
            DataGrid1.Columns("id").Visible = False
            DataGrid1.Columns("concatenado").Visible = False
            DataGrid1.Columns(1).Width = 900
            DataGrid1.Columns(2).Width = 6500
            DataGrid1.Columns(3).Width = 1200
            DataGrid1.Columns(4).Width = 300
            DataGrid1.Columns(5).Width = 1300
            DataGrid1.Columns(5).Alignment = dbgRight
            DataGrid1.Columns(5).NumberFormat = "Currency"
            DataGrid1.Columns(6).Width = 1200
            DataGrid1.Columns(7).Width = 1300
            
            
            DataGrid1.Refresh
            If xcuenta = 1 Then SendKeys "{ENTER}", False

        End If
        DataGrid1.SetFocus
        
        
    End If

End Sub
