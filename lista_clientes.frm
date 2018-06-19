VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form lista_clientes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Maestro de Clientes"
   ClientHeight    =   6015
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13350
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6015
   ScaleWidth      =   13350
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton salir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   11880
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "lista_clientes.frx":0000
      Height          =   5175
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   13095
      _ExtentX        =   23098
      _ExtentY        =   9128
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      BackColor       =   16777215
      HeadLines       =   1
      RowHeight       =   15
      WrapCellPointer =   -1  'True
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
         Weight          =   400
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
      TabIndex        =   3
      Top             =   0
      Width           =   13350
      _ExtentX        =   23548
      _ExtentY        =   953
      ButtonWidth     =   609
      ButtonHeight    =   900
      Appearance      =   1
      Style           =   1
      _Version        =   393216
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1080
         TabIndex        =   1
         Top             =   120
         Width           =   5055
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Buscar:"
         Height          =   375
         Left            =   120
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   120
         Width           =   855
      End
   End
   Begin MSAdodcLib.Adodc datcliente 
      Height          =   330
      Left            =   120
      Top             =   5640
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
   Begin MSAdodcLib.Adodc datparametros 
      Height          =   330
      Left            =   1920
      Top             =   5640
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
Attribute VB_Name = "lista_clientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public cuentacont As String
Dim cuenta(99999) As Integer



Private Sub DataGrid1_DblClick()
    
        If menu = 1 Then
            frmnota_venta.Text19.Text = datagrid1.Columns(0).Text
            frmnota_venta.Text1(1).Text = datagrid1.Columns(2).Text
            SendKeys "{ENTER}", False
            Unload Me
        End If
        If menu = 2 Then
            frmpresupuesto.Text19.Text = datagrid1.Columns(0).Text
            frmpresupuesto.Text1(1).Text = datagrid1.Columns(2).Text
            SendKeys "{ENTER}", False
            Unload Me
        End If
        If menu = 3 Then
                frmalquiler.Text19.Text = datagrid1.Columns(0).Text
                frmalquiler.Text1(1).Text = datagrid1.Columns(2).Text
                SendKeys "{ENTER}", False
                Unload Me
        End If
        If menu = 4 Then
                frmfacctacte_alquiler.Text19.Text = datagrid1.Columns(0).Text
                frmfacctacte_alquiler.Text1(1).Text = datagrid1.Columns(2).Text
                SendKeys "{ENTER}", False
                Unload Me
        End If
        If menu = 5 Then
                frmfacctacte_venta.Text19.Text = datagrid1.Columns(0).Text
                frmfacctacte_venta.Text1(1).Text = datagrid1.Columns(2).Text
                SendKeys "{ENTER}", False
                Unload Me
        End If
        If menu = 6 Then
                frmnota_credito.Text19.Text = datagrid1.Columns(0).Text
                frmnota_credito.Text1(1).Text = datagrid1.Columns(2).Text
                SendKeys "{ENTER}", False
                Unload Me
        End If
        If menu = 7 Then
                frmnota_debito.Text19.Text = datagrid1.Columns(0).Text
                frmnota_debito.Text1(1).Text = datagrid1.Columns(2).Text
                SendKeys "{ENTER}", False
                Unload Me
        End If
        If menu = 8 Then
                frmrecibo_ctacte.Text19.Text = datagrid1.Columns(0).Text
                frmrecibo_ctacte.Text1(0).Text = datagrid1.Columns(2).Text
                SendKeys "{ENTER}", False
                Unload Me
        End If


End Sub

Private Sub DataGrid1_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
          If datcliente.Recordset.EOF = False Then
            If menu = 1 Then
                frmnota_venta.Text19.Text = datagrid1.Columns(0).Text
                frmnota_venta.Text1(1).Text = datagrid1.Columns(2).Text
                SendKeys "{ENTER}", False
            End If
            If menu = 2 Then
                frmpresupuesto.Text19.Text = datagrid1.Columns(0).Text
                frmpresupuesto.Text1(1).Text = datagrid1.Columns(2).Text
                SendKeys "{ENTER}", False
            End If
            If menu = 3 Then
                frmalquiler.Text19.Text = datagrid1.Columns(0).Text
                frmalquiler.Text1(1).Text = datagrid1.Columns(2).Text
                SendKeys "{ENTER}", False
            End If
            If menu = 4 Then
                frmfacctacte_alquiler.Text19.Text = datagrid1.Columns(0).Text
                frmfacctacte_alquiler.Text1(1).Text = datagrid1.Columns(2).Text
                SendKeys "{ENTER}", False
            End If
            If menu = 5 Then
                frmfacctacte_venta.Text19.Text = datagrid1.Columns(0).Text
                frmfacctacte_venta.Text1(1).Text = datagrid1.Columns(2).Text
                SendKeys "{ENTER}", False
            End If
            If menu = 6 Then
                frmnota_credito.Text19.Text = datagrid1.Columns(0).Text
                frmnota_credito.Text1(1).Text = datagrid1.Columns(2).Text
                SendKeys "{ENTER}", False
            End If
            If menu = 7 Then
                frmnota_debito.Text19.Text = datagrid1.Columns(0).Text
                frmnota_debito.Text1(1).Text = datagrid1.Columns(2).Text
                SendKeys "{ENTER}", False
            End If
            If menu = 8 Then
                frmrecibo_ctacte.Text19.Text = datagrid1.Columns(0).Text
                frmrecibo_ctacte.Text1(0).Text = datagrid1.Columns(2).Text
                SendKeys "{ENTER}", False
                Unload Me
            End If
           End If

        Unload Me
        
    End If

End Sub


Private Sub Form_Activate()

datagrid1.SetFocus


End Sub

Private Sub Form_Load()
If menu = 2 Then
'    Aplicar_skin2 Me
    Aplicar_skin Me
Else
    Aplicar_skin Me
End If


MiFuncionDeAjuste Me, True

datcliente.ConnectionString = login.conexiontotal
datparametros.ConnectionString = login.conexiontotal

datparametros.RecordSource = "select * from ud_ezi_parametros_pos where sucursal = '" & login.nomsucursal & "' "
datparametros.Refresh
    

datcliente.RecordSource = query
datcliente.Refresh


            datagrid1.Columns(0).Visible = False
            datagrid1.Columns(7).Visible = False
            datagrid1.Columns(4).Visible = False
            datagrid1.Columns(9).Visible = False
            datagrid1.Columns(10).Visible = False
            datagrid1.Columns(12).Visible = False
            datagrid1.Columns(15).Visible = False
            datagrid1.Columns(17).Visible = False
            datagrid1.Columns(18).Visible = False
            datagrid1.Columns(1).Width = 1000
            datagrid1.Columns(2).Width = 3500
            datagrid1.Columns(6).Width = 3500

 
End Sub

Private Sub salir_Click()
    
    Unload Me

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)

On Error Resume Next


   xcodclientefiltra = datparametros.Recordset.Fields("codclientefiltra")
   If xcodclientefiltra = "01" Then xcontrocliente = "88447B8E-14FE-4D60-9622-B22F6C735701"  ' tucuman
   If xcodclientefiltra = "04" Then xcontrocliente = "4234CA46-B2BE-4690-AC6A-F0DE206F94A9"  ' salta
   If xcodclientefiltra = "03" Then xcontrocliente = "AEC7FBAC-63F7-4404-9512-033D0961D9BC"  ' jujuy



    If KeyAscii = 13 Then
        KeyAscii = 0
        If Text1.Text <> "" Then
            Text1.Text = Replace(Text1.Text, " ", "%%")
            xbusqueda = "%" + Text1.Text + "%"
                      
            xquery1 = "SELECT     ALIAS_0.ID, ALIAS_0.CODIGO, ALIAS_3.NOMBRE AS RAZONSOCIAL, ALIAS_3.CUIT, ALIAS_6.CALLE, ISNULL(ALIAS_6.CALLE, '') + '-' + ISNULL(V_CIUDAD_.NOMBRE, '') " & _
              "+ '-' + ISNULL(V_PROVINCIA_.NOMBRE, '') AS DOMICILIO, ALIAS_0.DENOMINACION, ALIAS_5.NOMBRE AS ZONA, ALIAS_7.NUMERO AS TELEFONO, " & _
              "ALIAS_8.DIRECCIONELECTRONICA AS MAIL, V_TIPOPAGO_.NOMBRE AS TP, V_TIPOPAGO_.OBSERVACION AS TipoPago, " & _
              "V_TIPOPAGO_.ACTIVESTATUS AS ACTIVESTATUSTP, V_PROVINCIA_.NOMBRE AS Provincia, V_CIUDAD_.NOMBRE AS Ciudad, " & _
              "ALIAS_0.CODIGO + ' ' + ALIAS_3.CUIT + ' ' + ALIAS_3.NOMBRE AS CONCATENADO, V_EZI_POCICION_IVA_CLIENTES.CODIGO AS IVA, V_TIPOPAGO_.ID AS IDPAGO, " & _
              "ALIAS_0.DOMICILIOFACTURACION_ID AS domicilio_id, ALIAS_0.LISTAPRECIO_ID AS listaprecio, V_UD_CLIENTE.observacion, ALIAS_0.creditomaximo " & _
              "FROM         V_TIPOPAGO_ RIGHT OUTER JOIN " & _
              "V_CLIENTE AS ALIAS_0 WITH (NOLOCK) LEFT OUTER JOIN " & _
              "V_UD_CLIENTE with (nolock) ON ALIAS_0.BOEXTENSION_ID = V_UD_CLIENTE.ID LEFT OUTER JOIN " & _
              "V_EZI_POCICION_IVA_CLIENTES ON ALIAS_0.ID = V_EZI_POCICION_IVA_CLIENTES.idcliente ON V_TIPOPAGO_.ID = ALIAS_0.TIPOPAGO_ID LEFT OUTER JOIN " & _
              "V_PERSONA AS ALIAS_3 WITH (nolock) ON ALIAS_0.ENTEASOCIADO_ID = ALIAS_3.ID LEFT OUTER JOIN " & _
              "V_CIUDAD_ RIGHT OUTER JOIN " & _
              "V_DOMICILIO_ AS ALIAS_6 ON V_CIUDAD_.ID = ALIAS_6.CIUDAD_ID LEFT OUTER JOIN " & _
              "V_PROVINCIA_ ON ALIAS_6.PROVINCIA_ID = V_PROVINCIA_.ID ON ALIAS_0.DOMICILIOFACTURACION_ID = ALIAS_6.ID LEFT OUTER JOIN " & _
              "V_TELEFONO_ AS ALIAS_7 ON ALIAS_3.TELEFONOPRINCIPAL_ID = ALIAS_7.ID LEFT OUTER JOIN " & _
              "V_DIRECCIONELECTRONICA_ AS ALIAS_8 ON ALIAS_3.DIRECELECTRONICAPRINCIPAL_ID = ALIAS_8.ID LEFT OUTER JOIN " & _
              "V_ZONA_ AS ALIAS_5 ON ALIAS_0.ZONA_ID = ALIAS_5.ID " & _
              "WHERE     (ALIAS_0.ACTIVESTATUS = 0) AND ALIAS_0.CODIGO + ' ' + ALIAS_3.CUIT + ' ' + ALIAS_3.NOMBRE + ' ' + ALIAS_0.DENOMINACION  like '" & xbusqueda & "'   " & _
              "order by ALIAS_3.NOMBRE "
                      
            datcliente.RecordSource = xquery1
            datcliente.Refresh
            datagrid1.Columns(0).Visible = False
            datagrid1.Columns(7).Visible = False
            datagrid1.Columns(4).Visible = False
            datagrid1.Columns(9).Visible = False
            datagrid1.Columns(10).Visible = False
            datagrid1.Columns(12).Visible = False
            datagrid1.Columns(15).Visible = False
            datagrid1.Columns(17).Visible = False
            datagrid1.Columns(18).Visible = False
            datagrid1.Columns(1).Width = 1000
            datagrid1.Columns(2).Width = 3500
            datagrid1.Columns(6).Width = 3500


        End If
        datagrid1.SetFocus
        
        
    End If

End Sub
