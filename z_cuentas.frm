VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{5734474E-78D3-4254-99B9-C35F31BDF509}#62.0#0"; "vbskpro2.ocx"
Begin VB.Form z_cuentas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cuenta Contable"
   ClientHeight    =   6285
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5925
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6285
   ScaleWidth      =   5925
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   195
      Left            =   4320
      TabIndex        =   9
      Top             =   4920
      Visible         =   0   'False
      Width           =   975
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Bindings        =   "z_cuentas.frx":0000
      Height          =   315
      Left            =   1680
      TabIndex        =   8
      Top             =   120
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   556
      _Version        =   393216
      BackColor       =   16777152
      ListField       =   "Nombre Cuenta"
      Text            =   ""
   End
   Begin VB.Frame Frame1 
      Caption         =   "Ordenar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   240
      TabIndex        =   4
      Top             =   4920
      Width           =   2415
      Begin VB.OptionButton Option3 
         Caption         =   "Por Codigo Contable"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   840
         Width           =   2055
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Por Codigo Abreviado"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   600
         Width           =   2055
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Por nombre de cuenta"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   2055
      End
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      Height          =   285
      Left            =   840
      TabIndex        =   2
      Top             =   120
      Width           =   735
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grilla 
      Bindings        =   "z_cuentas.frx":0020
      Height          =   4335
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   7646
      _Version        =   393216
      Cols            =   3
      WordWrap        =   -1  'True
      AllowBigSelection=   0   'False
      FocusRect       =   0
      ScrollBars      =   2
      SelectionMode   =   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   3
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "&Aceptar"
      Height          =   735
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5280
      UseMaskColor    =   -1  'True
      Width           =   975
   End
   Begin vbskpro.Skinner Skinner1 
      Left            =   4080
      Top             =   0
      _ExtentX        =   1270
      _ExtentY        =   1270
      CloseButton     =   0
      MaxButton       =   0
      MinButton       =   0
      CloseButtonToolTipText=   "Cerrar"
      MinButtonToolTipText=   "Minimizar"
      MaxButtonToolTipText=   "Maximizar"
      OldForeColor    =   0
      RestoreButtonToolTipText=   "Restaurar"
      ChangeSkinButton=   0   'False
      MinToBarButtonToolTipText=   "Minimizar a la barra de títulos"
      RestoreFromBarButtonToolTipText=   "Restaurar ventana"
      AlwaysOnTopButtonToolTipText=   "Hacer siempre visible"
      AlwaysOnTopDownButtonToolTipText=   "Hacer no siempre visible"
      ChangeSkinButtonToolTipText=   "Cambiar skin"
      HelpButtonToolTipText=   "Ayuda"
      SysEnableSkinCaption=   "Habilitar &Skin"
      SysDisableSkinCaption=   "Deshabilitar &Skin"
      LcK1            =   "3.66*/4/0*/1-5*210/."
      LcK2            =   $"z_cuentas.frx":003E
      AmbientB        =   ";<=>?7B:><7=<A<7CC;@"
      ChSD_FormCaption=   "Seleccione Skin"
      ChSD_ManualSetFrameCaption=   "S&elección manual "
      ChSD_TitleBarSkinComboBoxCaption=   "Skin &barra de Tít."
      ChSD_TitleBarForeColorSetCaption=   "T&exto barra de Tít."
      ChSD_BodySkinComboBoxCaption=   "Skin del cuer&po"
      ChSD_BodyForeColorSetCaption=   "Te&xto del cuerpo"
      ChSD_ChangeForeColorCaption=   "Cambia&r"
      ChSD_SaveToFileCaption=   "&Guardar en un archivo"
      ChSD_LoadFromFileCaption=   "Cargar desde arc&hivo"
      ChSD_UseSkinFileCaption=   "&Usar archivo de skin"
      ChSD_OkCommandButtonCaption=   "&Aceptar"
      ChSD_CancelCommandButtonCaption=   "&Cancelar"
   End
   Begin MSAdodcLib.Adodc datcuenta 
      Height          =   330
      Left            =   4800
      Top             =   0
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
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
      Connect         =   "Provider=MSDASQL.1;Password=25072004;Persist Security Info=True;User ID=lucva;Data Source=contable;Initial Catalog=contablesql"
      OLEDBString     =   "Provider=MSDASQL.1;Password=25072004;Persist Security Info=True;User ID=lucva;Data Source=contable;Initial Catalog=contablesql"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select cuentas.* from cuentas order by idcuenta"
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
   Begin VB.Label Label1 
      Caption         =   "Cuenta:"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "z_cuentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public menucuentas As String
Dim lineas As Integer

Option Explicit

Private Sub Command1_Click()
On Error Resume Next
Dim i As Integer
Dim e As Integer
Dim mensa As String
    
    grilla.Rows = datcuenta.Recordset.RecordCount + 1
    lineas = grilla.Rows
    grilla.ColWidth(0) = 800
    grilla.ColWidth(1) = 1000
    grilla.ColWidth(2) = 3500
    grilla.Col = 0
    grilla.Row = 0
    grilla.Text = "Cod.Abr."
    grilla.Col = 1
    grilla.Text = "Cod.Cont"
    grilla.Col = 2
    grilla.Text = "Nombre Cuenta"
    

i = 1
datcuenta.Recordset.MoveFirst
Do While Not datcuenta.Recordset.EOF
    
    grilla.Row = i
    e = i Mod 2
    grilla.Col = 0
    grilla.Text = datcuenta.Recordset.Fields(0)
    grilla.Col = 1
    grilla.Text = datcuenta.Recordset.Fields(4)
    grilla.Col = 2
    grilla.Text = datcuenta.Recordset.Fields(1)
    If e = 0 Then
        grilla.Col = 0
        grilla.CellBackColor = QBColor(11)
        grilla.Col = 1
        grilla.CellBackColor = QBColor(11)
        grilla.Col = 2
        grilla.CellBackColor = QBColor(11)
    
    End If
        
    datcuenta.Recordset.MoveNext
    i = i + 1
Loop

Rem    If frmconceptos.DataGrid1.Columns(6).Text = "" Then frmconceptos.DataGrid1.Columns(6).Text = 0
    grilla.Row = 1
          
End Sub

Private Sub DataCombo1_KeyPress(KeyAscii As Integer)

On Error GoTo fuera
Dim x As Integer
Dim mensa As String

    If KeyAscii = 13 Then
        KeyAscii = 0
        grilla.SetFocus
        For x = 1 To lineas
            grilla.Row = x
            grilla.Col = 2
            If grilla.Text = DataCombo1.Text Then
                Call OKButton_Click
                Exit Sub
            End If
        Next x
    End If
    Exit Sub
fuera:
    mensa = MsgBox("Cuenta no existente", vbCritical, "Error")

End Sub

Private Sub Form_Load()
Dim i As Integer
Dim e As Integer
Dim mensa As String

    datcuenta.RecordSource = "select cuentas.* from cuentas where empre = " & login.empresaact & " and inicioper = '" & login.iper & "'  and imp = 'S' order by idcuenta asc"
    datcuenta.Refresh
    If datcuenta.Recordset.EOF = True Then
        mensa = MsgBox("No existe plan de cuentas", vbCritical, "Atencion")
    End If

    Call Command1_Click
          

End Sub



Private Sub grilla_KeyDown(KeyCode As Integer, Shift As Integer)

    grilla.ColSel = 0
    grilla.RowSel = grilla.Row

End Sub

Private Sub grilla_KeyPress(KeyAscii As Integer)
Dim x As Integer

    If KeyAscii = 13 Then
        KeyAscii = 0
        Call OKButton_Click
    End If

If KeyAscii > 48 Then
    Text1.Text = Chr(KeyAscii)
    Text1.SetFocus
    SendKeys "{END}", True
End If

End Sub

Private Sub OKButton_Click()

grilla.Col = 0
If menucuentas = "" Then frmotrosparam.DataGrid5.Columns(1).Text = grilla.Text
If menucuentas = "caja" Then frmcajaconceptos.Text1(2).Text = grilla.Text
If menucuentas = "banco" Then frmcajaconceptos.Text1(3).Text = grilla.Text
If menucuentas = "cajaparam1" Then frmcajabanco.Text1(2).Text = grilla.Text
If menucuentas = "cajaparam2" Then frmcajabanco.Text1(3).Text = grilla.Text
If menucuentas = "busca" Then
    frmajusteclientes.Text5.Text = grilla.Text
    frmajusteclientes.grabaasiento.SetFocus
    SendKeys "{enter}", False
End If

If menucuentas = "busca1" Then
    frmajusteproveedores.Text5.Text = grilla.Text
    frmajusteproveedores.grabaasiento.SetFocus
    SendKeys "{enter}", False
End If

menucuentas = ""
Unload Me


End Sub

Private Sub Option1_Click()

If Option1.Value = True Then
    datcuenta.RecordSource = "select cuentas.* from cuentas where empre = " & login.empresaact & " and inicioper = '" & login.iper & "'  and imp = 'S' order by [Nombre Cuenta] asc"
    datcuenta.Refresh
    Call Command1_Click
End If

End Sub

Private Sub Option2_Click()

If Option2.Value = True Then
    datcuenta.RecordSource = "select cuentas.* from cuentas where empre = " & login.empresaact & " and inicioper = '" & login.iper & "'  and imp = 'S' order by [Cod Contable] asc"
    datcuenta.Refresh
    Call Command1_Click
End If

End Sub

Private Sub Option3_Click()

If Option3.Value = True Then
    datcuenta.RecordSource = "select cuentas.* from cuentas where empre = " & login.empresaact & " and inicioper = '" & login.iper & "'  and imp = 'S' order by idcuenta asc"
    datcuenta.Refresh
    Call Command1_Click
End If

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
On Error GoTo fuera
Dim x As Integer
Dim mensa As String

    If KeyAscii = 13 Then
        KeyAscii = 0
        grilla.SetFocus
        For x = 1 To lineas
            grilla.Row = x
            grilla.Col = 0
            If grilla.Text = Text1.Text Then
                Call OKButton_Click
                Exit Sub
            End If
        Next x
    End If
    Exit Sub
fuera:
    mensa = MsgBox("Cuenta no existente", vbCritical, "Error")
    
End Sub
