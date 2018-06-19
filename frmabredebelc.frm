VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{5734474E-78D3-4254-99B9-C35F31BDF509}#62.0#0"; "vbskpro2.ocx"
Begin VB.Form frmabredebelc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Abrir cuenta del Libro Compras"
   ClientHeight    =   3930
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5025
   LinkTopic       =   "form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3930
   ScaleWidth      =   5025
   Begin VB.CommandButton label1 
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   22
      Top             =   360
      Width           =   2775
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Total:"
      Height          =   255
      Index           =   9
      Left            =   3480
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   120
      Width           =   855
   End
   Begin MSDataListLib.DataList DataList2 
      Bindings        =   "frmabredebelc.frx":0000
      Height          =   1620
      Left            =   360
      TabIndex        =   11
      Top             =   1680
      Visible         =   0   'False
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   2752
      _Version        =   393216
      MatchEntry      =   -1  'True
      BackColor       =   -2147483626
      ForeColor       =   -2147483647
      ListField       =   "codigo"
      BoundColumn     =   "Cod Contable"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Total:"
      Height          =   255
      Index           =   8
      Left            =   360
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   2280
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cuenta:"
      Height          =   255
      Index           =   7
      Left            =   2880
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   1920
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cuenta:"
      Height          =   255
      Index           =   6
      Left            =   2880
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   1560
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cuenta:"
      Height          =   255
      Index           =   5
      Left            =   2880
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   1200
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cuenta:"
      Height          =   255
      Index           =   4
      Left            =   2880
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   840
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Subtotal 4:"
      Height          =   255
      Index           =   3
      Left            =   360
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   1920
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Subtotal 3:"
      Height          =   255
      Index           =   2
      Left            =   360
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Subtotal 2:"
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   1200
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Subtotal 1:"
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Height          =   615
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2880
      UseMaskColor    =   -1  'True
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   3
      Left            =   3720
      TabIndex        =   7
      Top             =   1920
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   2
      Left            =   3720
      TabIndex        =   5
      Top             =   1560
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   1
      Left            =   3720
      TabIndex        =   3
      Top             =   1200
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   0
      Left            =   3720
      TabIndex        =   1
      Top             =   840
      Width           =   855
   End
   Begin VB.PictureBox importes 
      Height          =   255
      Left            =   3120
      ScaleHeight     =   195
      ScaleWidth      =   1395
      TabIndex        =   9
      Top             =   360
      Width           =   1455
   End
   Begin vbskpro.Skinner Skinner1 
      Left            =   0
      Top             =   -120
      _ExtentX        =   1270
      _ExtentY        =   1270
      CloseButtonToolTipText=   "Cerrar"
      MinButtonToolTipText=   "Minimizar"
      MaxButtonToolTipText=   "Maximizar"
      RestoreButtonToolTipText=   "Restaurar"
      Enabled         =   0   'False
      ChangeSkinButton=   0   'False
      MinToBarButtonToolTipText=   "Minimizar a la barra de títulos"
      RestoreFromBarButtonToolTipText=   "Restaurar ventana"
      AlwaysOnTopButtonToolTipText=   "Hacer siempre visible"
      AlwaysOnTopDownButtonToolTipText=   "Hacer no siempre visible"
      ChangeSkinButtonToolTipText=   "Cambiar skin"
      HelpButtonToolTipText=   "Ayuda"
      SysEnableSkinCaption=   "Habilitar &Skin"
      SysDisableSkinCaption=   "Deshabilitar &Skin"
      LcK2            =   $"frmabredebelc.frx":0019
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
   Begin VB.PictureBox subtotal 
      Height          =   255
      Index           =   0
      Left            =   1440
      ScaleHeight     =   195
      ScaleWidth      =   1155
      TabIndex        =   0
      Top             =   840
      Width           =   1215
   End
   Begin VB.PictureBox subtotal 
      Height          =   255
      Index           =   1
      Left            =   1440
      ScaleHeight     =   195
      ScaleWidth      =   1155
      TabIndex        =   2
      Top             =   1200
      Width           =   1215
   End
   Begin VB.PictureBox subtotal 
      Height          =   255
      Index           =   2
      Left            =   1440
      ScaleHeight     =   195
      ScaleWidth      =   1155
      TabIndex        =   4
      Top             =   1560
      Width           =   1215
   End
   Begin VB.PictureBox subtotal 
      Height          =   255
      Index           =   3
      Left            =   1440
      ScaleHeight     =   195
      ScaleWidth      =   1155
      TabIndex        =   6
      Top             =   1920
      Width           =   1215
   End
   Begin VB.PictureBox total 
      Height          =   255
      Left            =   1200
      ScaleHeight     =   195
      ScaleWidth      =   1395
      TabIndex        =   10
      Top             =   2280
      Width           =   1455
   End
   Begin MSAdodcLib.Adodc datcuentas 
      Height          =   330
      Left            =   0
      Top             =   0
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
Attribute VB_Name = "frmabredebelc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim posicion As Integer

Private Sub Command1_Click()

If frmlibrocompras_nuevo.librocontado = 0 Then
    For x = 0 To 3
        frmlibrocompras_nuevo.grilla.Row = frmlibrocompras_nuevo.indice
        frmlibrocompras_nuevo.grilla.Col = x * 2
        frmlibrocompras_nuevo.grilla.Text = Val(subtotal(x).Value)
        frmlibrocompras_nuevo.grilla.Col = (x * 2) + 1
        frmlibrocompras_nuevo.grilla.Text = Text1(x).Text
    Next x
    frmlibrocompras_nuevo.Text3(frmlibrocompras_nuevo.indice).Text = total.Value
    frmlibrocompras_nuevo.Text7(frmlibrocompras_nuevo.indice * 2).Text = Text1(0).Text
    Unload Me
    frmlibrocompras_nuevo.Text3(frmlibrocompras_nuevo.indice).SetFocus
Else
    For x = 0 To 3
        frmotrosgastos_nuevo.grilla.Row = frmlibrocompras_nuevo.indice
        frmotrosgastos_nuevo.grilla.Col = x * 2
        frmotrosgastos_nuevo.grilla.Text = Val(subtotal(x).Value)
        frmotrosgastos_nuevo.grilla.Col = (x * 2) + 1
        frmotrosgastos_nuevo.grilla.Text = Text1(x).Text
    Next x
    If frmlibrocompras_nuevo.indice = 15 Then
        If Val(total.Value) <> Val(importes.Value) Then
            MsgBox "Los importes son Incorrectos", vbCritical, "Error"
            Exit Sub
        End If
    End If
    frmotrosgastos_nuevo.Text3(frmlibrocompras_nuevo.indice).Text = total.Value
    frmotrosgastos_nuevo.Text7(frmlibrocompras_nuevo.indice * 2).Text = Text1(0).Text
    Unload Me
    If frmlibrocompras_nuevo.indice <> 15 Then
        frmotrosgastos_nuevo.Text3(frmlibrocompras_nuevo.indice).SetFocus
    Else
        frmotrosgastos_nuevo.grabalibroasiento.SetFocus
    End If
End If


End Sub

Private Sub DataList2_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        Text1(posicion).Text = DataList2.BoundText
        If posicion = 3 Then
            Command1.SetFocus
            Exit Sub
        End If
        subtotal(posicion + 1).SetFocus
    End If

End Sub

Private Sub DataList2_LostFocus()

    DataList2.Visible = False

End Sub

Private Sub Form_Load()
Aplicar_skin Me


datcuentas.ConnectionString = login.conexiontotal


      
If frmlibrocompras_nuevo.librocontado = 0 Then
    frmlibrocompras_nuevo.grilla.Row = frmlibrocompras_nuevo.indice
    For x = 0 To 3
        frmlibrocompras_nuevo.grilla.Col = x * 2
        subtotal(x).Value = frmlibrocompras_nuevo.grilla.Text
        frmlibrocompras_nuevo.grilla.Col = (x * 2) + 1
        Text1(x).Text = frmlibrocompras_nuevo.grilla.Text
        total.Value = Val(total.Value) + Val(subtotal(x).Value)
    Next x
Else
    frmotrosgastos_nuevo.grilla.Row = frmlibrocompras_nuevo.indice
    For x = 0 To 3
        frmotrosgastos_nuevo.grilla.Col = x * 2
        subtotal(x).Value = frmotrosgastos_nuevo.grilla.Text
        frmotrosgastos_nuevo.grilla.Col = (x * 2) + 1
        Text1(x).Text = frmotrosgastos_nuevo.grilla.Text
        total.Value = Val(total.Value) + Val(subtotal(x).Value)
    Next x
End If


End Sub

Private Sub DataList2_GotFocus()
On Error GoTo fuera

    If Inicio.opcion1 = True Then
        datcuentas.RecordSource = "select listacuentas.* from listacuentas WHERE empre = " & login.empresaact & " and inicioper = '" & login.iper & "' and finper = '" & login.fper & "' ORDER BY IDCUENTA"
        datcuentas.Refresh
        DataList2.ListField = "codigo"
    End If
    If Inicio.opcion2 = True Then
        datcuentas.RecordSource = "select listacuentas.* from listacuentas WHERE empre = " & login.empresaact & " and inicioper = '" & login.iper & "' and finper = '" & login.fper & "' ORDER BY nombre"
        datcuentas.Refresh
        DataList2.ListField = "nombre"
    End If
    
fuera:
End Sub

Private Sub Label8_Click()
End Sub

Private Sub Label11_Click()

End Sub

Private Sub subtotal_KeyPress(Index As Integer, KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        total.Value = 0
        For x = 0 To 3
            total.Value = Val(total.Value) + Val(subtotal(x).Value)
        Next x
        SendKeys "{tab}"
    End If
End Sub

Private Sub Text1_GotFocus(Index As Integer)

    DataList2.Visible = True
    DataList2.Top = Text1(Index).Height + Text1(Index).Top
    If Text1(Index).Text = "" Then DataList2.BoundText = Text1(Index).Text
    posicion = Index
    DataList2.SetFocus

End Sub
