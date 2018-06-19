VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form lista_proveedores_consulta 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Maestro de Proveedores Consulta"
   ClientHeight    =   8730
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14670
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8730
   ScaleWidth      =   14670
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame3 
      Caption         =   "Anotaciones de Clientes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4815
      Left            =   7320
      TabIndex        =   5
      Top             =   2280
      Visible         =   0   'False
      Width           =   9735
      Begin VB.CommandButton Command9 
         Caption         =   "X"
         Height          =   375
         Left            =   9120
         TabIndex        =   7
         Top             =   120
         Width           =   495
      End
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
         Height          =   3975
         Left            =   240
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   6
         Text            =   "lista_proveedores_consulta.frx":0000
         Top             =   600
         Width           =   9375
      End
   End
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
      Bindings        =   "lista_proveedores_consulta.frx":0006
      Height          =   8055
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   14415
      _ExtentX        =   25426
      _ExtentY        =   14208
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
      Height          =   600
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   14670
      _ExtentX        =   25876
      _ExtentY        =   1058
      ButtonWidth     =   609
      ButtonHeight    =   1005
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
         Width           =   6975
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
      Left            =   0
      Top             =   8280
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
Attribute VB_Name = "lista_proveedores_consulta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public cuentacont As String
Dim cuenta(99999) As Integer



Private Sub Command9_Click()
Frame3.Visible = False
End Sub

Private Sub DataGrid1_DblClick()

On Error Resume Next

    Frame3.Visible = True
    Frame3.Caption = DataGrid1.Columns("razonsocial").Text
    Text2.Text = DataGrid1.Columns("anotaciones").Text
    

End Sub

Private Sub DataGrid1_KeyDown(KeyCode As Integer, Shift As Integer)

Frame3.Visible = False

End Sub

Private Sub DataGrid1_KeyUp(KeyCode As Integer, Shift As Integer)

Frame3.Visible = False

End Sub

Private Sub Form_Load()
Aplicar_skin Me

MiFuncionDeAjuste Me, True

yventana = Inicio.Height / 2 - 1000
xventana = Inicio.Width / 2

'lista_clientes_consulta.Top = yventana - lista_clientes_consulta.Height / 2
'lista_clientes_consulta.Left = xventana - lista_clientes_consulta.Width / 2


datcliente.ConnectionString = login.conexiontotal

     xquery = "SELECT     ALIAS_0.ID, ALIAS_0.CODIGO, ALIAS_3.NOMBRE AS RAZONSOCIAL, ALIAS_3.CUIT, ALIAS_6.CALLE, ISNULL(ALIAS_6.CALLE, '') + '-' + ISNULL(V_CIUDAD_.NOMBRE, '') " & _
              "+ '-' + ISNULL(V_PROVINCIA_.NOMBRE, '') AS DOMICILIO, ALIAS_0.DENOMINACION, ALIAS_7.NUMERO AS TELEFONO, ALIAS_8.DIRECCIONELECTRONICA AS MAIL, " & _
              "V_TIPOPAGO_.NOMBRE AS TP, V_TIPOPAGO_.OBSERVACION AS TipoPago, V_TIPOPAGO_.ACTIVESTATUS AS ACTIVESTATUSTP, " & _
              "V_PROVINCIA_.NOMBRE AS Provincia, V_CIUDAD_.NOMBRE AS Ciudad, ALIAS_0.CODIGO + ' ' + ALIAS_3.CUIT + ' ' + ALIAS_3.NOMBRE AS CONCATENADO, " & _
              "V_TIPOPAGO_.ID AS IDPAGO, V_UD_PROVEEDOR_.Anotaciones " & _
              "FROM         V_TIPOPAGO_ RIGHT OUTER JOIN " & _
              "V_PROVEEDOR AS ALIAS_0 WITH (nolock) LEFT OUTER JOIN " & _
              "V_UD_PROVEEDOR_ ON ALIAS_0.BOEXTENSION_ID = V_UD_PROVEEDOR_.ID ON V_TIPOPAGO_.ID = ALIAS_0.TIPOPAGO_ID LEFT OUTER JOIN " & _
              "V_PERSONA AS ALIAS_3 WITH (nolock) ON ALIAS_0.ENTEASOCIADO_ID = ALIAS_3.ID LEFT OUTER JOIN " & _
              "V_CIUDAD_ RIGHT OUTER JOIN " & _
              "V_DOMICILIO_ AS ALIAS_6 ON V_CIUDAD_.ID = ALIAS_6.CIUDAD_ID LEFT OUTER JOIN " & _
              "V_PROVINCIA_ ON ALIAS_6.PROVINCIA_ID = V_PROVINCIA_.ID ON ALIAS_3.DOMICILIOPRINCIPAL_ID = ALIAS_6.ID LEFT OUTER JOIN " & _
              "V_TELEFONO_ AS ALIAS_7 ON ALIAS_3.TELEFONOPRINCIPAL_ID = ALIAS_7.ID LEFT OUTER JOIN " & _
              "V_DIRECCIONELECTRONICA_ AS ALIAS_8 ON ALIAS_3.DIRECELECTRONICAPRINCIPAL_ID = ALIAS_8.ID " & _
              "ORDER BY RAZONSOCIAL"

datcliente.RecordSource = xquery
datcliente.Refresh


            DataGrid1.Columns(0).Visible = False
            DataGrid1.Columns(9).Visible = False
            DataGrid1.Columns(11).Visible = False
            DataGrid1.Columns(14).Visible = False
            DataGrid1.Columns(15).Visible = False
            DataGrid1.Columns(1).Width = 1000
            DataGrid1.Columns(2).Width = 3500


 
End Sub

Private Sub Form_Resize()

    DataGrid1.Width = lista_proveedores_consulta.Width - 200
    DataGrid1.Height = lista_proveedores_consulta.Height - 1000

End Sub

Private Sub salir_Click()
    
    Unload Me

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)

On Error Resume Next


    If KeyAscii = 13 Then
        KeyAscii = 0
        If Text1.Text <> "" Then
            xbusqueda = "%" + Text1.Text + "%"
            
            xquery1 = "SELECT     ALIAS_0.ID, ALIAS_0.CODIGO, ALIAS_3.NOMBRE AS RAZONSOCIAL, ALIAS_3.CUIT, ALIAS_6.CALLE, ISNULL(ALIAS_6.CALLE, '') + '-' + ISNULL(V_CIUDAD_.NOMBRE, '') " & _
              "+ '-' + ISNULL(V_PROVINCIA_.NOMBRE, '') AS DOMICILIO, ALIAS_0.DENOMINACION, ALIAS_7.NUMERO AS TELEFONO, ALIAS_8.DIRECCIONELECTRONICA AS MAIL, " & _
              "V_TIPOPAGO_.NOMBRE AS TP, V_TIPOPAGO_.OBSERVACION AS TipoPago, V_TIPOPAGO_.ACTIVESTATUS AS ACTIVESTATUSTP, " & _
              "V_PROVINCIA_.NOMBRE AS Provincia, V_CIUDAD_.NOMBRE AS Ciudad, ALIAS_0.CODIGO + ' ' + ALIAS_3.CUIT + ' ' + ALIAS_3.NOMBRE AS CONCATENADO, " & _
              "V_TIPOPAGO_.ID AS IDPAGO, V_UD_PROVEEDOR_.Anotaciones " & _
              "FROM         V_TIPOPAGO_ RIGHT OUTER JOIN " & _
              "V_PROVEEDOR AS ALIAS_0 WITH (nolock) LEFT OUTER JOIN " & _
              "V_UD_PROVEEDOR_ ON ALIAS_0.BOEXTENSION_ID = V_UD_PROVEEDOR_.ID ON V_TIPOPAGO_.ID = ALIAS_0.TIPOPAGO_ID LEFT OUTER JOIN " & _
              "V_PERSONA AS ALIAS_3 WITH (nolock) ON ALIAS_0.ENTEASOCIADO_ID = ALIAS_3.ID LEFT OUTER JOIN " & _
              "V_CIUDAD_ RIGHT OUTER JOIN " & _
              "V_DOMICILIO_ AS ALIAS_6 ON V_CIUDAD_.ID = ALIAS_6.CIUDAD_ID LEFT OUTER JOIN " & _
              "V_PROVINCIA_ ON ALIAS_6.PROVINCIA_ID = V_PROVINCIA_.ID ON ALIAS_3.DOMICILIOPRINCIPAL_ID = ALIAS_6.ID LEFT OUTER JOIN " & _
              "V_TELEFONO_ AS ALIAS_7 ON ALIAS_3.TELEFONOPRINCIPAL_ID = ALIAS_7.ID LEFT OUTER JOIN " & _
              "V_DIRECCIONELECTRONICA_ AS ALIAS_8 ON ALIAS_3.DIRECELECTRONICAPRINCIPAL_ID = ALIAS_8.ID " & _
              "WHERE     (ALIAS_0.ACTIVESTATUS = 0) AND ALIAS_0.CODIGO + ' ' + ALIAS_3.CUIT + ' ' + ALIAS_3.NOMBRE + ' ' + ALIAS_0.DENOMINACION   like '" & xbusqueda & "' order by ALIAS_3.NOMBRE "
            
            datcliente.RecordSource = xquery1
            datcliente.Refresh
            
            DataGrid1.Columns(0).Visible = False
            DataGrid1.Columns(9).Visible = False
            DataGrid1.Columns(11).Visible = False
            DataGrid1.Columns(14).Visible = False
            DataGrid1.Columns(15).Visible = False
            DataGrid1.Columns(1).Width = 1000
            DataGrid1.Columns(2).Width = 3500


        End If
        DataGrid1.SetFocus
        
        
    End If

End Sub
