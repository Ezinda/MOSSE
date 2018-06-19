VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form lista_productos_servicios_emporio 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Maestro de Articulos"
   ClientHeight    =   7830
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   16695
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7830
   ScaleWidth      =   16695
   Begin VB.CommandButton salir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   12720
      TabIndex        =   4
      Top             =   120
      Width           =   1215
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "lista_productos_servicios_emporio.frx":0000
      Height          =   6735
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Visible         =   0   'False
      Width           =   16455
      _ExtentX        =   29025
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
      Width           =   16695
      _ExtentX        =   29448
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
End
Attribute VB_Name = "lista_productos_servicios_emporio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public cuentacont As String
Dim cuenta(99999) As Integer



Private Sub DataGrid1_DblClick()

'            frmnota_venta.Text1(1).Text = DataGrid1.Columns(2).Text
'            SendKeys "{ENTER}", False
'        Unload Me

End Sub

Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
On Error Resume Next

    If KeyAscii = 13 Then
        KeyAscii = 0
            If menu = 1 Then

                frmnota_credito.grilla.Row = xfila
                frmnota_credito.grilla.Col = 0
                frmnota_credito.grilla.Text = DataGrid1.Columns("id").Text
                frmnota_credito.grilla.Col = 1
                frmnota_credito.grilla.Text = DataGrid1.Columns("codigo").Text
                frmnota_credito.grilla.Col = 2
                frmnota_credito.grilla.Text = DataGrid1.Columns("producto").Text
                frmnota_credito.grilla.Col = 4
                frmnota_credito.grilla.Text = DataGrid1.Columns("um").Text
                frmnota_credito.grilla.Col = 5
                frmnota_credito.grilla.Text = Format(Round(DataGrid1.Columns("precio").Text, 2), "###,##0.00")
                frmnota_credito.grilla.Col = 13
                frmnota_credito.grilla.Text = Format(Round(DataGrid1.Columns("precio").Text, 2), "###,##0.00")
                
                
'                frmnota_credito.grilla.Row = xfila
 '               frmnota_credito.grilla.Col = 0
 '               frmnota_credito.grilla.Text = DataGrid1.Columns("id").Text
 '               frmnota_credito.grilla.Col = 1
 '               frmnota_credito.grilla.Text = DataGrid1.Columns("codigo").Text
 '               frmnota_credito.grilla.Col = 2
 '               frmnota_credito.grilla.Text = DataGrid1.Columns("producto").Text
 '               frmnota_credito.grilla.Col = 4
 '               frmnota_credito.grilla.Text = DataGrid1.Columns("um").Text
 '               frmnota_credito.grilla.Col = 6
 '               frmnota_credito.grilla.Text = Format(Round(xprecios, 2), "###,##0.00")
 '               frmnota_credito.grilla.Col = 14
 '               frmnota_credito.grilla.Text = Format(Round(xprecios, 2), "###,##0.00")
 '               frmnota_credito.grilla.Col = 17
 '               frmnota_credito.grilla.Text = Format(Round(xprecios, 2), "###,##0.00")


                
'                datimpuestos.RecordSource = "SELECT     CASE WHEN pni.COEFICIENTEDEFAULT = 0 THEN 21 ELSE pni.COEFICIENTEDEFAULT END AS COEFICIENTEDEFAULT, p.CODIGO, p.ID, " & _
'                      "V_UNIDADMEDIDA__1.NOMBRE AS UMVTA, V_UNIDADMEDIDA_.NOMBRE AS UMSTK " & _
'                      "FROM         V_PRODUCTO_ AS p INNER JOIN " & _
'                      "V_POSICIONADORIMPUESTOS_ AS pi ON p.POSICIONADORIMPUESTOS_ID = pi.ID INNER JOIN " & _
'                      "V_ITEMPOSICIONADORIMPUESTOS_ AS ipi ON pi.ITEMSPOSICIONADORIMPUESTOS_ID = ipi.BO_PLACE_ID INNER JOIN " & _
'                      "V_POSICIONIMPUESTO_ AS pni ON ipi.POSICIONIMPUESTO_ID = pni.ID INNER JOIN " & _
'                      "V_DEFINICIONIMPUESTO_ AS d ON ipi.DEFINICIONIMPUESTO_ID = d.ID INNER JOIN " & _
'                      "V_IMPUESTO_ AS i ON d.IMPUESTO_ID = i.ID LEFT OUTER JOIN " & _
'                      "V_UNIDADMEDIDA_ ON p.UNIDADMEDIDA_ID = V_UNIDADMEDIDA_.ID LEFT OUTER JOIN " & _
'                      "V_UNIDADMEDIDA_ AS V_UNIDADMEDIDA__1 ON p.UNIDADMEDIDANOLINEAL_ID = V_UNIDADMEDIDA__1.ID " & _
'                      "WHERE     (i.CODIGO = '010') and p.id = '" & DataGrid1.Columns("id").Text & "'"
'                datimpuestos.Refresh
                
                
                xquery2 = "SELECT     CASE WHEN pni.COEFICIENTEDEFAULT = 0 THEN 21 ELSE pni.COEFICIENTEDEFAULT END AS COEFICIENTEDEFAULT, p.CODIGO, p.ID, " & _
                      "V_UNIDADMEDIDA__1.NOMBRE AS UMVTA, V_UNIDADMEDIDA_.NOMBRE AS UMSTK, 'P' as tipo " & _
                      "FROM V_PRODUCTO_ AS p INNER JOIN " & _
                      "V_POSICIONADORIMPUESTOS_ AS pi ON p.POSICIONADORIMPUESTOS_ID = pi.ID INNER JOIN " & _
                      "V_ITEMPOSICIONADORIMPUESTOS_ AS ipi ON pi.ITEMSPOSICIONADORIMPUESTOS_ID = ipi.BO_PLACE_ID INNER JOIN " & _
                      "V_POSICIONIMPUESTO_ AS pni ON ipi.POSICIONIMPUESTO_ID = pni.ID INNER JOIN " & _
                      "V_DEFINICIONIMPUESTO_ AS d ON ipi.DEFINICIONIMPUESTO_ID = d.ID INNER JOIN " & _
                      "V_IMPUESTO_ AS i ON d.IMPUESTO_ID = i.ID LEFT OUTER JOIN " & _
                      "V_UNIDADMEDIDA_ ON p.UNIDADMEDIDA_ID = V_UNIDADMEDIDA_.ID LEFT OUTER JOIN " & _
                      "V_UNIDADMEDIDA_ AS V_UNIDADMEDIDA__1 ON p.UNIDADMEDIDANOLINEAL_ID = V_UNIDADMEDIDA__1.ID " & _
                      "WHERE     (i.CODIGO = '010') and p.id = '" & DataGrid1.Columns("id").Text & "'" & _
                      "UNION ALL " & _
                      "SELECT     pni.COEFICIENTEDEFAULT AS COEFICIENTEDEFAULT, p.CODIGO, p.ID, " & _
                      "V_UNIDADMEDIDA__1.NOMBRE AS UMVTA, V_UNIDADMEDIDA_.NOMBRE AS UMSTK, 'C' as tipo " & _
                      "FROM         V_SERVICIO_ AS p INNER JOIN " & _
                      "V_POSICIONADORIMPUESTOS_ AS pi ON p.POSICIONADORIMPUESTOS_ID = pi.ID INNER JOIN " & _
                      "V_ITEMPOSICIONADORIMPUESTOS_ AS ipi ON pi.ITEMSPOSICIONADORIMPUESTOS_ID = ipi.BO_PLACE_ID INNER JOIN " & _
                      "V_POSICIONIMPUESTO_ AS pni ON ipi.POSICIONIMPUESTO_ID = pni.ID INNER JOIN " & _
                      "V_DEFINICIONIMPUESTO_ AS d ON ipi.DEFINICIONIMPUESTO_ID = d.ID INNER JOIN " & _
                      "V_IMPUESTO_ AS i ON d.IMPUESTO_ID = i.ID LEFT OUTER JOIN " & _
                      "V_UNIDADMEDIDA_ ON p.UNIDADMEDIDA_ID = V_UNIDADMEDIDA_.ID LEFT OUTER JOIN " & _
                      "V_UNIDADMEDIDA_ AS V_UNIDADMEDIDA__1 ON p.UNIDADMEDIDANOLINEAL_ID = V_UNIDADMEDIDA__1.ID " & _
                      "WHERE     (i.CODIGO = '010') and p.id = '" & DataGrid1.Columns("id").Text & "'" & _
                      "UNION ALL "
                      
                  XQUERY3 = "SELECT   pni.COEFICIENTEDEFAULT  AS COEFICIENTEDEFAULT, p.CODIGO, p.ID, " & _
                      "V_UNIDADMEDIDA__1.NOMBRE AS UMVTA, V_UNIDADMEDIDA_.NOMBRE AS UMSTK, 'C' as tipo " & _
                      "FROM         V_CONCEPTOCONTABLE_ AS p INNER JOIN " & _
                      "V_POSICIONADORIMPUESTOS_ AS pi ON p.POSICIONADORIMPUESTOS_ID = pi.ID INNER JOIN " & _
                      "V_ITEMPOSICIONADORIMPUESTOS_ AS ipi ON pi.ITEMSPOSICIONADORIMPUESTOS_ID = ipi.BO_PLACE_ID INNER JOIN " & _
                      "V_POSICIONIMPUESTO_ AS pni ON ipi.POSICIONIMPUESTO_ID = pni.ID INNER JOIN " & _
                      "V_DEFINICIONIMPUESTO_ AS d ON ipi.DEFINICIONIMPUESTO_ID = d.ID INNER JOIN " & _
                      "V_IMPUESTO_ AS i ON d.IMPUESTO_ID = i.ID LEFT OUTER JOIN " & _
                      "V_UNIDADMEDIDA_ ON p.UNIDADMEDIDA_ID = V_UNIDADMEDIDA_.ID LEFT OUTER JOIN " & _
                      "V_UNIDADMEDIDA_ AS V_UNIDADMEDIDA__1 ON p.UNIDADMEDIDANOLINEAL_ID = V_UNIDADMEDIDA__1.ID " & _
                      "WHERE     (i.CODIGO = '010') and p.id = '" & DataGrid1.Columns("id").Text & "'"

                datimpuestos.RecordSource = xquery2 + XQUERY3
                datimpuestos.Refresh
                
                
                
                If datimpuestos.Recordset.EOF = True Then
                    xiva = 1.21
                Else
                    If datimpuestos.Recordset.Fields("tipo") = "C" Then
                        xiva = (datimpuestos.Recordset.Fields("coeficientedefault") + 100) / 100
                    Else
                        xiva = 1.21
                    End If
                End If

                frmnota_credito.grilla.Col = 6
                frmnota_credito.grilla.Text = Format(Round(DataGrid1.Columns("precio").Text / xiva, 2), "###,##0.00")
                frmnota_credito.grilla.Col = 7
                frmnota_credito.grilla.Text = Format(0, "###,##0.00")
                frmnota_credito.grilla.Col = 8
                frmnota_credito.grilla.Text = Format(0, "###,##0.00")
                frmnota_credito.grilla.Col = 9
                frmnota_credito.grilla.Text = Format(Round(DataGrid1.Columns("precio").Text, 2), "###,##0.00")
                frmnota_credito.grilla.Col = 10
                frmnota_credito.grilla.Text = 1
                frmnota_credito.grilla.Col = 11
                frmnota_credito.grilla.Text = xiva

                                      
'                frmnota_credito.grilla.Col = 5
'                frmnota_credito.grilla.Text = Format(Round(xprecios / xiva, 2), "###,##0.000")
'                frmnota_credito.grilla.Col = 7
'                frmnota_credito.grilla.Text = Format(Round(xprecios / xiva, 2), "###,##0.00")
                
'                frmnota_credito.grilla.Col = 8
'                frmnota_credito.grilla.Text = Format(0, "###,##0.00")
'                frmnota_credito.grilla.Col = 9
'                frmnota_credito.grilla.Text = Format(0, "###,##0.00")
'                frmnota_credito.grilla.Col = 10
'                frmnota_credito.grilla.Text = Format(Round(xprecios, 2), "###,##0.00")
'                frmnota_credito.grilla.Col = 11
'                frmnota_credito.grilla.Text = 1
'                frmnota_credito.grilla.Col = 12
'                frmnota_credito.grilla.Text = xiva
                
                
                
                frmnota_credito.grilla.Col = 3
                frmnota_credito.grilla.Text = 1
                
               
                frmnota_credito.grilla.SetFocus
                
            End If
            If menu = 2 Then
                frmnota_debito.grilla.Row = xfila
                frmnota_debito.grilla.Col = 0
                frmnota_debito.grilla.Text = DataGrid1.Columns("id").Text
                frmnota_debito.grilla.Col = 1
                frmnota_debito.grilla.Text = DataGrid1.Columns("codigo").Text
                frmnota_debito.grilla.Col = 2
                frmnota_debito.grilla.Text = DataGrid1.Columns("producto").Text
                frmnota_debito.grilla.Col = 4
                frmnota_debito.grilla.Text = DataGrid1.Columns("um").Text
                frmnota_debito.grilla.Col = 5
                frmnota_debito.grilla.Text = Format(Round(DataGrid1.Columns("precio").Text, 2), "###,##0.00")
                frmnota_debito.grilla.Col = 13
                frmnota_debito.grilla.Text = Format(Round(DataGrid1.Columns("precio").Text, 2), "###,##0.00")

                
'                datimpuestos.RecordSource = "SELECT     CASE WHEN pni.COEFICIENTEDEFAULT = 0 THEN 21 ELSE pni.COEFICIENTEDEFAULT END AS COEFICIENTEDEFAULT, p.CODIGO, p.ID, " & _
'                      "V_UNIDADMEDIDA__1.NOMBRE AS UMVTA, V_UNIDADMEDIDA_.NOMBRE AS UMSTK " & _
'                      "FROM         V_PRODUCTO_ AS p INNER JOIN " & _
'                      "V_POSICIONADORIMPUESTOS_ AS pi ON p.POSICIONADORIMPUESTOS_ID = pi.ID INNER JOIN " & _
'                      "V_ITEMPOSICIONADORIMPUESTOS_ AS ipi ON pi.ITEMSPOSICIONADORIMPUESTOS_ID = ipi.BO_PLACE_ID INNER JOIN " & _
'                      "V_POSICIONIMPUESTO_ AS pni ON ipi.POSICIONIMPUESTO_ID = pni.ID INNER JOIN " & _
'                      "V_DEFINICIONIMPUESTO_ AS d ON ipi.DEFINICIONIMPUESTO_ID = d.ID INNER JOIN " & _
'                      "V_IMPUESTO_ AS i ON d.IMPUESTO_ID = i.ID LEFT OUTER JOIN " & _
'                      "V_UNIDADMEDIDA_ ON p.UNIDADMEDIDA_ID = V_UNIDADMEDIDA_.ID LEFT OUTER JOIN " & _
'                      "V_UNIDADMEDIDA_ AS V_UNIDADMEDIDA__1 ON p.UNIDADMEDIDANOLINEAL_ID = V_UNIDADMEDIDA__1.ID " & _
'                      "WHERE     (i.CODIGO = '010') and p.id = '" & DataGrid1.Columns("id").Text & "'"
'                datimpuestos.Refresh
                
                xquery2 = "SELECT     CASE WHEN pni.COEFICIENTEDEFAULT = 0 THEN 21 ELSE pni.COEFICIENTEDEFAULT END AS COEFICIENTEDEFAULT, p.CODIGO, p.ID, " & _
                      "V_UNIDADMEDIDA__1.NOMBRE AS UMVTA, V_UNIDADMEDIDA_.NOMBRE AS UMSTK, 'P' as tipo " & _
                      "FROM V_PRODUCTO_ AS p INNER JOIN " & _
                      "V_POSICIONADORIMPUESTOS_ AS pi ON p.POSICIONADORIMPUESTOS_ID = pi.ID INNER JOIN " & _
                      "V_ITEMPOSICIONADORIMPUESTOS_ AS ipi ON pi.ITEMSPOSICIONADORIMPUESTOS_ID = ipi.BO_PLACE_ID INNER JOIN " & _
                      "V_POSICIONIMPUESTO_ AS pni ON ipi.POSICIONIMPUESTO_ID = pni.ID INNER JOIN " & _
                      "V_DEFINICIONIMPUESTO_ AS d ON ipi.DEFINICIONIMPUESTO_ID = d.ID INNER JOIN " & _
                      "V_IMPUESTO_ AS i ON d.IMPUESTO_ID = i.ID LEFT OUTER JOIN " & _
                      "V_UNIDADMEDIDA_ ON p.UNIDADMEDIDA_ID = V_UNIDADMEDIDA_.ID LEFT OUTER JOIN " & _
                      "V_UNIDADMEDIDA_ AS V_UNIDADMEDIDA__1 ON p.UNIDADMEDIDANOLINEAL_ID = V_UNIDADMEDIDA__1.ID " & _
                      "WHERE     (i.CODIGO = '010') and p.id = '" & DataGrid1.Columns("id").Text & "'" & _
                      "UNION ALL " & _
                      "SELECT     pni.COEFICIENTEDEFAULT AS COEFICIENTEDEFAULT, p.CODIGO, p.ID, " & _
                      "V_UNIDADMEDIDA__1.NOMBRE AS UMVTA, V_UNIDADMEDIDA_.NOMBRE AS UMSTK, 'C' as tipo " & _
                      "FROM         V_SERVICIO_ AS p INNER JOIN " & _
                      "V_POSICIONADORIMPUESTOS_ AS pi ON p.POSICIONADORIMPUESTOS_ID = pi.ID INNER JOIN " & _
                      "V_ITEMPOSICIONADORIMPUESTOS_ AS ipi ON pi.ITEMSPOSICIONADORIMPUESTOS_ID = ipi.BO_PLACE_ID INNER JOIN " & _
                      "V_POSICIONIMPUESTO_ AS pni ON ipi.POSICIONIMPUESTO_ID = pni.ID INNER JOIN " & _
                      "V_DEFINICIONIMPUESTO_ AS d ON ipi.DEFINICIONIMPUESTO_ID = d.ID INNER JOIN " & _
                      "V_IMPUESTO_ AS i ON d.IMPUESTO_ID = i.ID LEFT OUTER JOIN " & _
                      "V_UNIDADMEDIDA_ ON p.UNIDADMEDIDA_ID = V_UNIDADMEDIDA_.ID LEFT OUTER JOIN " & _
                      "V_UNIDADMEDIDA_ AS V_UNIDADMEDIDA__1 ON p.UNIDADMEDIDANOLINEAL_ID = V_UNIDADMEDIDA__1.ID " & _
                      "WHERE     (i.CODIGO = '010') and p.id = '" & DataGrid1.Columns("id").Text & "'" & _
                      "UNION ALL "
                      
                  XQUERY3 = "SELECT   pni.COEFICIENTEDEFAULT  AS COEFICIENTEDEFAULT, p.CODIGO, p.ID, " & _
                      "V_UNIDADMEDIDA__1.NOMBRE AS UMVTA, V_UNIDADMEDIDA_.NOMBRE AS UMSTK, 'C' as tipo " & _
                      "FROM         V_CONCEPTOCONTABLE_ AS p INNER JOIN " & _
                      "V_POSICIONADORIMPUESTOS_ AS pi ON p.POSICIONADORIMPUESTOS_ID = pi.ID INNER JOIN " & _
                      "V_ITEMPOSICIONADORIMPUESTOS_ AS ipi ON pi.ITEMSPOSICIONADORIMPUESTOS_ID = ipi.BO_PLACE_ID INNER JOIN " & _
                      "V_POSICIONIMPUESTO_ AS pni ON ipi.POSICIONIMPUESTO_ID = pni.ID INNER JOIN " & _
                      "V_DEFINICIONIMPUESTO_ AS d ON ipi.DEFINICIONIMPUESTO_ID = d.ID INNER JOIN " & _
                      "V_IMPUESTO_ AS i ON d.IMPUESTO_ID = i.ID LEFT OUTER JOIN " & _
                      "V_UNIDADMEDIDA_ ON p.UNIDADMEDIDA_ID = V_UNIDADMEDIDA_.ID LEFT OUTER JOIN " & _
                      "V_UNIDADMEDIDA_ AS V_UNIDADMEDIDA__1 ON p.UNIDADMEDIDANOLINEAL_ID = V_UNIDADMEDIDA__1.ID " & _
                      "WHERE     (i.CODIGO = '010') and p.id = '" & DataGrid1.Columns("id").Text & "'"

                datimpuestos.RecordSource = xquery2 + XQUERY3
                datimpuestos.Refresh
                
                
                
                If datimpuestos.Recordset.EOF = True Then
                    xiva = 1.21
                Else
                    If datimpuestos.Recordset.Fields("tipo") = "C" Then
                        xiva = (datimpuestos.Recordset.Fields("coeficientedefault") + 100) / 100
                    Else
                        xiva = 1.21
                    End If
                End If
                                      
                frmnota_debito.grilla.Col = 6
                frmnota_debito.grilla.Text = Format(Round(DataGrid1.Columns("precio").Text / xiva, 2), "###,##0.00")
                frmnota_debito.grilla.Col = 7
                frmnota_debito.grilla.Text = Format(0, "###,##0.00")
                frmnota_debito.grilla.Col = 8
                frmnota_debito.grilla.Text = Format(0, "###,##0.00")
                frmnota_debito.grilla.Col = 9
                frmnota_debito.grilla.Text = Format(Round(DataGrid1.Columns("precio").Text, 2), "###,##0.00")
                frmnota_debito.grilla.Col = 10
                frmnota_debito.grilla.Text = 1
                frmnota_debito.grilla.Col = 11
                frmnota_debito.grilla.Text = xiva
                
                
                frmnota_debito.grilla.Col = 3
                frmnota_debito.grilla.Text = 1
                
                
                frmnota_debito.grilla.SetFocus
            End If
            If menu = 5 Then
'**** Lista de Precios especiales, busca historico
            xlista = frmfacctacte_venta.DataGrid2.Columns("listaprecio").Text
            If xlista <> "{8D0FED00-A782-11D5-936C-00E07D9040B9}" And xlista <> "" Then
                xidcliente = frmfacctacte_venta.DataGrid2.Columns("id").Text
                datpreciosespeciales.RecordSource = "SELECT   top 1  V_ITEMFACTURAVENTA_.NOMBREREFERENCIA, V_ITEMFACTURAVENTA_.REFERENCIA_ID, V_ITEMFACTURAVENTA_.VALOR2_IMPORTE, " & _
                      "V_ITEMFACTURAVENTA_.DESTINATARIOTR_ID, V_ITEMFACTURAVENTA_.NOMBREDESTINATARIOTR, V_ITEMFACTURAVENTA_.FECHADOCUMENTO, " & _
                      "V_LISTAPRECIO_.NOMBRE AS listaprecio, V_ITEMFACTURAVENTA_.DETALLE " & _
                      "FROM         V_LISTAPRECIO_ RIGHT OUTER JOIN " & _
                      "V_CLIENTE_ ON V_LISTAPRECIO_.ID = V_CLIENTE_.LISTAPRECIO_ID RIGHT OUTER JOIN " & _
                      "V_ITEMFACTURAVENTA_ ON V_CLIENTE_.ID = V_ITEMFACTURAVENTA_.DESTINATARIOTR_ID " & _
                      "WHERE     (V_ITEMFACTURAVENTA_.REFERENCIA_ID = '" & DataGrid1.Columns("id").Text & "') AND " & _
                      "(V_ITEMFACTURAVENTA_.DESTINATARIOTR_ID = '" & xidcliente & "') " & _
                      "order by FECHADOCUMENTO desc"
                datpreciosespeciales.Refresh
                If datpreciosespeciales.Recordset.EOF = False Then
                      xprecios = datpreciosespeciales.Recordset.Fields("valor2_importe") * (1 + (Val(datpreciosespeciales.Recordset.Fields("detalle")) / 100))
                Else
                      xprecios = 0
                End If
                frmfacctacte_venta.Label1(19).Visible = True
            Else
                xprecios = DataGrid1.Columns("precio").Text
                frmfacctacte_venta.Label1(19).Visible = False
            End If
            
                
'**** Fin Lista de Precios especiales, busca historico
            
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

End Sub


Private Sub Form_Activate()

Text1.SetFocus

End Sub

Private Sub Form_Load()

If login.nomsucursal = "EMPORIOZIP" Or login.nomsucursal = "TUCUMANZIP" Then
'    Aplicar_skin2 Me
    Aplicar_skin Me
Else
    Aplicar_skin Me
End If

MiFuncionDeAjuste Me, True

yventana = Inicio.Height / 2 - 1000
xventana = Inicio.Width / 2

lista_productos_servicios_emporio.Top = yventana - lista_productos_servicios_emporio.Height / 2
lista_productos_servicios_emporio.Left = xventana - lista_productos_servicios_emporio.Width / 2


datproducto.ConnectionString = login.conexiontotal
datimpuestos.ConnectionString = login.conexiontotal
datpreciosespeciales.ConnectionString = login.conexiontotal

'datproducto.RecordSource = query
'datproducto.Refresh


'            DataGrid1.Columns(0).Visible = False
'            DataGrid1.Columns(8).Visible = False
'            DataGrid1.Refresh

 
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
                        xselect = " AND  ISNULL(p.CODIGO, '') + ' ' + ISNULL(p.DESCRIPCION, '') " & _
                                  "+ ISNULL(ALIAS_1.NOMBRERUBRO, '') + ' ' + ISNULL(ALIAS_2.CODIGOID, '') + ' ' + ISNULL(ALIAS_3.CODIGO, '') + ' ' + ISNULL(ALIAS_2.APLICACION, '') " & _
                                  "+ ' ' + ISNULL(ALIAS_2.PRODUCTO_SUST, '') + ' ' + ISNULL(ALIAS_2.CODIGOSALT, '') + ' ' + ISNULL(ALIAS_2.REFERENCIA, '') LIKE "
                    Else
                        xbusq = xbusq + xselect + "'%" + palabra(h) + "%"
                    End If
               End If
             Next h
                    
            xbusqueda = xbusq
            
            xquery1 = "SELECT     p.ID, LEFT(p.CODIGO, CHARINDEX(' -- ', p.CODIGO)) AS Codigo, p.DESCRIPCION AS Producto, t.CODIGO AS Marca, r.PRESENTACION AS Presentacion, " & _
                      "ROUND(CAST(PR.PRECIOCIVA AS decimal(14, 3)), 3) AS Precio, PR.PrecioCosto, SUBSTRING(PR.FECHAULTACT, 7, 2) + '/' + SUBSTRING(PR.FECHAULTACT, 5, 2) " & _
                      "+ '/' + LEFT(PR.FECHAULTACT, 4) AS FechaUltAct, st.CANTIDAD2_CANTIDAD AS Stock, v.DENOMINACION AS Proveedor, u.DETALLE AS rubro, V_UNIDADMEDIDA_.NOMBRE AS UM, p.CODIGO + ISNULL(r.PRESENTACION, '') + p.DESCRIPCION + ISNULL(v.DENOMINACION, '') " & _
                      "+ CASE i.CODIGO WHEN 'I' THEN 'Importado' WHEN 'N' THEN 'Nacional' ELSE '' END + ISNULL(u.DETALLE, '') + ISNULL(r.CODPROVEEDOR, '') + ISNULL(t.CODIGO, '') " & _
                      "AS concatenado, ISNULL(PR.margen, 0) AS Margen, r.ANMAT, r.CODPROVEEDOR AS CodProveedor FROM         V_PRODUCTO AS p WITH (NOLOCK) LEFT OUTER JOIN " & _
                      "v_ezi_pos_stock_global AS st ON p.ID = st.REFERENCIATIPO_ID LEFT OUTER JOIN V_EZI_PRECIOS_POS AS PR ON p.ID = PR.ID LEFT OUTER JOIN " & _
                      "V_UNIDADMEDIDA_ ON p.UNIDADMEDIDANOLINEAL_ID = V_UNIDADMEDIDA_.ID LEFT OUTER JOIN V_UD_EZI_PRODUCTOS AS r WITH (NOLOCK) ON p.BOEXTENSION_ID = r.ID LEFT OUTER JOIN " & _
                      "V_PROVEEDOR AS v WITH (NOLOCK) ON r.PROVEEDOR_ID = v.ID LEFT OUTER JOIN V_ITEMTIPOCLASIFICADOR AS i WITH (NOLOCK) ON r.NACIONALIDAD_ID = i.ID LEFT OUTER JOIN " & _
                      "V_ITEMTIPOCLASIFICADOR AS t WITH (NOLOCK) ON r.MARCA_ID = t.ID LEFT OUTER JOIN V_RUBRO AS u WITH (NOLOCK) ON p.RUBRO_ID = u.ID " & _
                      "WHERE     (p.ACTIVESTATUS <> 2) AND (p.TIPOOBJETOESTATICO_ID IS NULL) AND (LEFT(p.CODIGO, CHARINDEX(' -- ', p.CODIGO)) + ' ' + ISNULL(r.PRESENTACION, '') " & _
                      "+ ' ' + p.DESCRIPCION + ISNULL(v.DENOMINACION, '') + ' ' + CASE i.CODIGO WHEN 'I' THEN 'Importado' WHEN 'N' THEN 'Nacional' ELSE '' END + ' ' + ISNULL(u.DETALLE, '') + ' ' + ISNULL(r.CODPROVEEDOR, '') " & _
                      "+ ' ' + ISNULL(t.CODIGO, '') LIKE '" & xbusqueda & "') " & _
                      "Union All " & _
                      "SELECT     ID, CODIGO, DESCRIPCION, 'Servicio' AS Marca, '' AS Presentacion, 0 AS Precio, 0 as Preciocto, '01/01/2000' AS fechaultac, 0 AS Stock, '' AS Proveedor, 'Servicio' AS Rubro, 'Unidad' AS UM, CODIGO + ' ' + DESCRIPCION AS concatenado, 0 as margen, '' as anmat, '' as codproveedor " & _
                      "From V_SERVICIO_ WHERE     (ACTIVESTATUS <> 2) AND (CODIGO + ' ' + DESCRIPCION LIKE '" & xbusqueda & "') " & _
                      "Union All " & _
                      "SELECT     ID, CODIGO, DESCRIPCION, 'Concepto' AS Marca, '' AS Presentacion, 0 AS Precio, 0 as Preciocto, '01/01/2000' AS fechaultac, 0 AS Stock, '' AS Proveedor, 'Servicio' AS Rubro, 'Unidad' AS UM, " & _
                      "CODIGO + ' ' + DESCRIPCION AS concatenado, 0 as margen, '' as anmat, '' as codproveedor From V_CONCEPTOCONTABLE_ " & _
                      "WHERE     (ACTIVESTATUS <> 2) AND (CODIGO + ' ' + DESCRIPCION LIKE '" & xbusqueda & "') " & _
                      "ORDER BY p.DESCRIPCION"
                      
            datproducto.RecordSource = xquery1
            datproducto.Refresh
            xcuenta = datproducto.Recordset.RecordCount
            DataGrid1.Visible = True
            
            DataGrid1.Columns("id").Visible = False
            DataGrid1.Columns("concatenado").Visible = False
            DataGrid1.Columns(1).Width = 2000
            DataGrid1.Columns(2).Width = 5500
            DataGrid1.Columns(3).Width = 1200
            DataGrid1.Columns(4).Width = 1200
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
