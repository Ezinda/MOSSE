VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{5734474E-78D3-4254-99B9-C35F31BDF509}#62.0#0"; "vbskpro2.ocx"
Begin VB.Form frmparamventas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Parametro Facturacion Ventas"
   ClientHeight    =   4410
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7095
   Icon            =   "frmparamventas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4410
   ScaleWidth      =   7095
   Begin VB.CommandButton grabar 
      Caption         =   "&Grabar"
      Height          =   615
      Left            =   1080
      Picture         =   "frmparamventas.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   3600
      UseMaskColor    =   -1  'True
      Width           =   855
   End
   Begin VB.CommandButton Cancelar 
      Cancel          =   -1  'True
      Caption         =   "Cance&lar"
      Height          =   615
      Left            =   3000
      Picture         =   "frmparamventas.frx":0974
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   3600
      UseMaskColor    =   -1  'True
      Width           =   855
   End
   Begin VB.CommandButton cerrar 
      Caption         =   "&Cerrar"
      Height          =   615
      Left            =   4920
      Picture         =   "frmparamventas.frx":0EA6
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   3600
      UseMaskColor    =   -1  'True
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Cerrar"
      Height          =   615
      Left            =   0
      Picture         =   "frmparamventas.frx":12E8
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   5280
      UseMaskColor    =   -1  'True
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Caption         =   "Opciones"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   975
      Left            =   240
      TabIndex        =   18
      Top             =   3360
      Width           =   6735
      Begin VB.CommandButton salir 
         Caption         =   "&Cerrar"
         Height          =   615
         Left            =   0
         Picture         =   "frmparamventas.frx":172A
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   5280
         UseMaskColor    =   -1  'True
         Width           =   855
      End
   End
   Begin MSDataListLib.DataList DataList2 
      Bindings        =   "frmparamventas.frx":1B6C
      Height          =   1230
      Left            =   3120
      TabIndex        =   7
      Top             =   1200
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   2170
      _Version        =   393216
      MatchEntry      =   -1  'True
      BackColor       =   12632256
      ListField       =   "ccostoslista"
      BoundColumn     =   "cc"
   End
   Begin MSDataListLib.DataList DataList1 
      Bindings        =   "frmparamventas.frx":1B85
      Height          =   1230
      Left            =   3120
      TabIndex        =   8
      Top             =   840
      Visible         =   0   'False
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   2170
      _Version        =   393216
      MatchEntry      =   -1  'True
      BackColor       =   12632256
      ListField       =   "codigo"
      BoundColumn     =   "Cod Contable"
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00C0C0C0&
      Height          =   1230
      Left            =   3120
      TabIndex        =   9
      Top             =   480
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      DataField       =   "alicuota4"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#.##0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   3082
         SubFormatType   =   1
      EndProperty
      DataSource      =   "datparamventas"
      Height          =   285
      Index           =   6
      Left            =   2400
      TabIndex        =   6
      Top             =   2640
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      DataField       =   "alicuota3"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#.##0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   3082
         SubFormatType   =   1
      EndProperty
      DataSource      =   "datparamventas"
      Height          =   285
      Index           =   5
      Left            =   2400
      TabIndex        =   5
      Top             =   2280
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      DataField       =   "alicuota2"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#.##0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   3082
         SubFormatType   =   1
      EndProperty
      DataSource      =   "datparamventas"
      Height          =   285
      Index           =   4
      Left            =   2400
      TabIndex        =   4
      Top             =   1920
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      DataField       =   "alicuota1"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#.##0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   3082
         SubFormatType   =   1
      EndProperty
      DataSource      =   "datparamventas"
      Height          =   285
      Index           =   3
      Left            =   2400
      TabIndex        =   3
      Top             =   1560
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      DataField       =   "centrocosto"
      DataSource      =   "datparamventas"
      Height          =   285
      Index           =   2
      Left            =   2400
      TabIndex        =   2
      Top             =   1200
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      DataField       =   "codcuenta"
      DataSource      =   "datparamventas"
      Height          =   285
      Index           =   1
      Left            =   2400
      TabIndex        =   1
      Top             =   840
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      DataField       =   "columnalibro"
      DataSource      =   "datparamventas"
      Height          =   285
      Index           =   0
      Left            =   2400
      TabIndex        =   0
      Top             =   480
      Width           =   735
   End
   Begin MSAdodcLib.Adodc datparamventas 
      Height          =   330
      Left            =   5760
      Top             =   3240
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
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
   Begin MSAdodcLib.Adodc datccostos 
      Height          =   330
      Left            =   3960
      Top             =   3240
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
   Begin MSAdodcLib.Adodc datcuentas 
      Height          =   330
      Left            =   4920
      Top             =   3240
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
   Begin MSAdodcLib.Adodc datcolumnas 
      Height          =   330
      Left            =   4440
      Top             =   3240
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
      LcK2            =   $"frmparamventas.frx":1B9E
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
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Parametros por defecto en facturacion Ventas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   1320
      TabIndex        =   17
      Top             =   120
      Width           =   4455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Alicuota Col 4: %"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   720
      TabIndex        =   16
      Top             =   2640
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Alicuota Col 3: %"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   720
      TabIndex        =   15
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Alicuota Col 2: %"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   720
      TabIndex        =   14
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Alicuota Col 1: %"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   720
      TabIndex        =   13
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Col.Libro:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   720
      TabIndex        =   12
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Cod.Cuenta:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   720
      TabIndex        =   11
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Cen.de Costo:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   720
      TabIndex        =   10
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00E0E0E0&
      Height          =   2775
      Left            =   240
      Top             =   360
      Width           =   6735
   End
End
Attribute VB_Name = "frmparamventas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cancelar_Click()

    datparamventas.Refresh

End Sub

Private Sub cerrar_Click()

    Unload Me

End Sub

Private Sub Form_Load()

  datccostos.ConnectionString = login.conexiontotal
  datcolumnas.ConnectionString = login.conexiontotal
  datcuentas.ConnectionString = login.conexiontotal
  datparamventas.ConnectionString = login.conexiontotal

    datparamventas.RecordSource = "select paramventas.* from paramventas where empresa = " & login.empresaact & ""
    datparamventas.Refresh
    
    If datparamventas.Recordset.EOF = True Then
        datparamventas.Recordset.AddNew
    End If

    datccostos.RecordSource = "select listaccostos.* from listaccostos where empresa = " & login.empresaact & ""
    datccostos.Refresh
    
    datcuentas.RecordSource = "select listacuentas.* from listacuentas where empre = " & login.empresaact & " and inicioper = '" & login.iper & "' ORDER BY IDCUENTA"
    datcuentas.Refresh
    
    datcolumnas.RecordSource = "SELECT columnasventa.* From columnasventa where empresa = " & login.empresaact & " and inicioper = '" & login.iper & "'"
    datcolumnas.Refresh
    
    If datcolumnas.Recordset.EOF = True Then Exit Sub
    For x = 1 To 30 Step 2
        columlibro = datcolumnas.Recordset.Fields(x)
        If IsNull(columlibro) = False Then List1.AddItem columlibro
    Next x
    
End Sub
Private Sub DataList1_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        Text1(1).Text = DataList1.BoundText
        Text1(2).SetFocus
    End If

End Sub

Private Sub DataList1_LostFocus()

    DataList1.Visible = False

End Sub

Private Sub DataList2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Text1(2).Text = DataList2.BoundText
        Text1(3).SetFocus
    End If
End Sub

Private Sub DataList2_LostFocus()

    DataList2.Visible = False

End Sub

Private Sub grabar_Click()
    
    datparamventas.Recordset.Fields("empresa") = login.empresaact
    datparamventas.Recordset.UpdateBatch adAffectCurrent
    datparamventas.Refresh

End Sub

Private Sub List1_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        Text1(0).Text = List1.ListIndex + 1
        Text1(1).SetFocus
    End If

End Sub

Private Sub List1_LostFocus()

    List1.Visible = False

End Sub

Private Sub Text1_GotFocus(Index As Integer)
    If Index = 0 Then
        List1.Visible = True
        If IsNull(Text1(0).Text) = False Then List1.ListIndex = Val(Text1(0).Text) - 1
        List1.SetFocus
    End If
    
    If Index = 1 Then
        DataList1.Visible = True
        DataList1.BoundText = Text1(1).Text
        DataList1.SetFocus
    End If
    
    If Index = 2 Then
        DataList2.Visible = True
        DataList2.BoundText = Text1(2).Text
        DataList2.SetFocus
    End If
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        If Index = 6 Then
            grabar.SetFocus
            Exit Sub
        End If
        Text1(Index + 1).SetFocus
    End If
    
End Sub
