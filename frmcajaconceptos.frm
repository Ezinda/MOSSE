VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{5734474E-78D3-4254-99B9-C35F31BDF509}#62.0#0"; "vbskpro2.ocx"
Begin VB.Form frmcajaconceptos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Conceptos Caja y Banco"
   ClientHeight    =   6255
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7695
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6255
   ScaleWidth      =   7695
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   120
      TabIndex        =   11
      Top             =   2040
      Width           =   7335
      Begin VB.CommandButton eliminar 
         Caption         =   "&Eliminar"
         Height          =   495
         Left            =   3240
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   975
      End
      Begin VB.CommandButton nuevo 
         Caption         =   "&Nuevo"
         Height          =   495
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   1095
      End
      Begin VB.CommandButton grabar 
         Caption         =   "&Grabar"
         Height          =   495
         Left            =   1800
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   1095
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmcajaconceptos.frx":0000
      Height          =   3015
      Left            =   120
      TabIndex        =   10
      Top             =   3120
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   5318
      _Version        =   393216
      AllowUpdate     =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
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
            LCID            =   1034
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
            LCID            =   1034
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
   Begin VB.Frame Frame1 
      Height          =   1935
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   7335
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   3
         Left            =   4320
         TabIndex        =   2
         Top             =   1320
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   2
         Left            =   1800
         TabIndex        =   1
         Top             =   1320
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Height          =   525
         Index           =   1
         Left            =   1800
         MaxLength       =   50
         MultiLine       =   -1  'True
         TabIndex        =   0
         Top             =   720
         Width           =   4935
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   1800
         TabIndex        =   9
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Cuenta Banco:"
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
         Left            =   2760
         TabIndex        =   12
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Cuenta Caja:"
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
         Left            =   480
         TabIndex        =   8
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Descripcion:"
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
         Left            =   600
         TabIndex        =   7
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Codigo:"
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
         Left            =   600
         TabIndex        =   6
         Top             =   360
         Width           =   1095
      End
   End
   Begin MSAdodcLib.Adodc datconceptos 
      Height          =   330
      Left            =   0
      Top             =   0
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
      LockType        =   4
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
      Left            =   7200
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
      LcK2            =   $"frmcajaconceptos.frx":001B
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
   Begin MSAdodcLib.Adodc datconceptos1 
      Height          =   330
      Left            =   1440
      Top             =   0
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
      LockType        =   4
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
Attribute VB_Name = "frmcajaconceptos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub DataGrid1_Click()

    Text1(0).Text = DataGrid1.Columns(0).Text
    Text1(1).Text = DataGrid1.Columns(1).Text
    Text1(2).Text = DataGrid1.Columns(2).Text
    Text1(3).Text = DataGrid1.Columns(3).Text


End Sub

Private Sub DataGrid1_KeyUp(KeyCode As Integer, Shift As Integer)
    Text1(0).Text = DataGrid1.Columns(0).Text
    Text1(1).Text = DataGrid1.Columns(1).Text
    Text1(2).Text = DataGrid1.Columns(2).Text
    Text1(3).Text = DataGrid1.Columns(3).Text
End Sub

Private Sub eliminar_Click()
 On Error Resume Next
mensa = MsgBox("Esta por eliminar este concepto, esta seguro", vbYesNo, "Atención")
If mensa = vbNo Then
    Exit Sub
End If
    
    datconceptos1.ConnectionString = login.conexiontotal
    datconceptos1.RecordSource = "delete conceptoscaja where empresa = " & login.empresaact & " and codigo = " & Text1(0).Text & ""
    datconceptos1.Refresh
    
    datconceptos.RecordSource = "select conceptoscaja.* from conceptoscaja where empresa = " & login.empresaact & " order by codigo"
    datconceptos.Refresh

End Sub

Private Sub Form_Load()
frmcajaconceptos.Top = 0
frmcajaconceptos.Left = 0

datconceptos.ConnectionString = login.conexiontotal
datconceptos.RecordSource = "select conceptoscaja.* from conceptoscaja where empresa = " & login.empresaact & " order by descripcion"
datconceptos.Refresh

DataGrid1.Columns(0).Width = 1000
DataGrid1.Columns(1).Width = 3000
DataGrid1.Columns(2).Width = 1000
DataGrid1.Columns(3).Width = 1000
DataGrid1.Columns(4).Visible = False

End Sub

Private Sub grabar_Click()
On Error Resume Next
datconceptos.Recordset.Fields("descripcion") = Text1(1).Text
datconceptos.Recordset.Fields("codcontable") = Text1(2).Text
datconceptos.Recordset.Fields("codcontablebanco") = Text1(3).Text
datconceptos.Recordset.UpdateBatch adAffectCurrent
datconceptos.Refresh

DataGrid1.Columns(0).Width = 1000
DataGrid1.Columns(1).Width = 3000
DataGrid1.Columns(2).Width = 1000
DataGrid1.Columns(3).Width = 1000
DataGrid1.Columns(4).Visible = False


End Sub

Private Sub nuevo_Click()


Text1(0).Text = ""
Text1(1).Text = ""
Text1(2).Text = ""
Text1(3).Text = ""
datconceptos.Recordset.AddNew
datconceptos.Recordset.Fields("empresa") = login.empresaact
Text1(1).SetFocus


End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{tab}", False
    End If
End Sub

Private Sub Text1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF3 And Index = 2 Then
        z_cuentas.menucuentas = "caja"
        z_cuentas.Show
    End If
    
    If KeyCode = vbKeyF3 And Index = 3 Then
        z_cuentas.menucuentas = "banco"
        z_cuentas.Show
    End If

End Sub
