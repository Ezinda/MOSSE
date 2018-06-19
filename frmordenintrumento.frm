VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{5734474E-78D3-4254-99B9-C35F31BDF509}#62.0#0"; "vbskpro2.ocx"
Begin VB.Form frmordenintrumento 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Asignacion de Orden de Pago Como instrumento de Pago"
   ClientHeight    =   6300
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7860
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6300
   ScaleWidth      =   7860
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grilla 
      Height          =   5415
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   9551
      _Version        =   393216
      Cols            =   4
      _NumberOfBands  =   1
      _Band(0).Cols   =   4
   End
   Begin MSAdodcLib.Adodc datinstru 
      Height          =   330
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
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
      Caption         =   ""
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
   Begin vbskpro.Skinner Skinner1 
      Left            =   0
      Top             =   0
      _ExtentX        =   1270
      _ExtentY        =   1270
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
      LcK2            =   $"frmordenintrumento.frx":0000
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
End
Attribute VB_Name = "frmordenintrumento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cuenta(10000) As Integer
Private Sub Form_Load()
frmordenintrumento.Top = 300
frmordenintrumento.Left = 600

  datinstru.ConnectionString = login.conexiontotal

  datinstru.RecordSource = "select consultaordensinc.* from consultaordensinc where empresa = " & login.empresaact & " and importe <> 0 and anulado <> 'S' order by nomproveedor, nrorden"
  datinstru.Refresh
  
  If datinstru.Recordset.EOF = True Then
        mensa = MsgBox("No tiene Ordenes sin Asignar", vbInformation, "Mensaje")
        Unload Me
        Exit Sub
  End If
  
  grilla.Cols = 5
  grilla.Rows = datinstru.Recordset.RecordCount + 1
  
  grilla.ColWidth(0) = 200
  grilla.ColWidth(1) = 1800
  grilla.ColWidth(2) = 1200
  grilla.ColWidth(3) = 3000
  grilla.ColWidth(4) = 1200
  
  For x = 2 To grilla.Rows - 1 Step 2
    For Y = 1 To 4
        grilla.Col = Y
        grilla.Row = x
        grilla.CellBackColor = QBColor(11)
    Next Y
  Next x
  
   
  datinstru.Recordset.MoveFirst
  
i = 1
Do While Not datinstru.Recordset.EOF

  grilla.Row = i
  grilla.Col = 1
  grilla.Text = datinstru.Recordset.Fields("nrorden")
  grilla.Col = 2
  grilla.Text = datinstru.Recordset.Fields("fecha")
  grilla.Col = 3
  grilla.Text = datinstru.Recordset.Fields("nomproveedor")
  grilla.Col = 4
  grilla.Text = datinstru.Recordset.Fields("importe")
  grilla.Text = Format(Val(grilla.Text), "###,##0.00")
  If IsNull(datinstru.Recordset.Fields("codcuenta")) = False Then cuenta(i) = datinstru.Recordset.Fields("codcuenta")
  i = i + 1
  datinstru.Recordset.MoveNext

Loop
    grilla.Col = 1
    grilla.Row = 1
    grilla.ColSel = 4
    
End Sub

Private Sub grilla_Click()
    grilla.RowSel = grilla.Row
    grilla.ColSel = 4
End Sub

Private Sub grilla_KeyDown(KeyCode As Integer, Shift As Integer)
    
    grilla.RowSel = grilla.Row
    grilla.ColSel = 4
    
End Sub

Private Sub grilla_KeyPress(KeyAscii As Integer)


    If KeyAscii = 13 Then
        KeyAscii = 0
        grilla.Col = 1
        frmordendepago1.Text5(0).Text = "Orden " + grilla.Text
        frmordendepago1.ordeninstu = grilla.Text
        grilla.Col = 4
        If grilla.Text < 0 Then grilla.Text = grilla.Text * -1
        frmordendepago1.importemask.Text = grilla.Text
        frmordendepago1.importeord = grilla.Text
        grilla.Col = 2
    Rem     frmordendepago1.MaskEdBox1.Text = grilla.Text
        frmordendepago1.Text5(1).Text = cuenta(grilla.Row)

        Unload Me
        frmordendepago1.importemask.SetFocus
    End If
        
End Sub
