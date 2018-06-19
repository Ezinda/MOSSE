VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{5734474E-78D3-4254-99B9-C35F31BDF509}#62.0#0"; "vbskpro2.ocx"
Begin VB.Form frmniveles 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configuración de Niveles"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4470
   Icon            =   "frmniveles.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4470
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox empre 
      DataField       =   "empre"
      DataSource      =   "nivelesrs"
      Height          =   285
      Left            =   1080
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   1080
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cerrar 
      Caption         =   "&Cerrar"
      Height          =   495
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2160
      UseMaskColor    =   -1  'True
      Width           =   1095
   End
   Begin MSAdodcLib.Adodc nivelesrs 
      Height          =   375
      Left            =   1200
      Top             =   2640
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
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
   Begin VB.CommandButton aceptar 
      Caption         =   "&Aceptar"
      Height          =   495
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2160
      UseMaskColor    =   -1  'True
      Width           =   1215
   End
   Begin VB.TextBox nivel5 
      Alignment       =   2  'Center
      DataField       =   "niv5"
      DataSource      =   "nivelesrs"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      TabIndex        =   4
      Top             =   1440
      Width           =   495
   End
   Begin VB.TextBox nivel4 
      Alignment       =   2  'Center
      DataField       =   "niv4"
      DataSource      =   "nivelesrs"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   3
      Top             =   1440
      Width           =   495
   End
   Begin VB.TextBox nivel3 
      Alignment       =   2  'Center
      DataField       =   "niv3"
      DataSource      =   "nivelesrs"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   2
      Top             =   1440
      Width           =   495
   End
   Begin VB.TextBox nivel2 
      Alignment       =   2  'Center
      DataField       =   "niv2"
      DataSource      =   "nivelesrs"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      Top             =   1440
      Width           =   495
   End
   Begin VB.TextBox nivel1 
      Alignment       =   2  'Center
      DataField       =   "niv1"
      DataSource      =   "nivelesrs"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      TabIndex        =   0
      Top             =   1440
      Width           =   495
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
      LcK2            =   $"frmniveles.frx":0442
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
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Cantidad de digitos por niveles del Codigo contable"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   600
      TabIndex        =   5
      Top             =   480
      Width           =   3135
   End
End
Attribute VB_Name = "frmniveles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub aceptar_Click()

Respuesta = MsgBox("ESTA MODIFICACION NO DEBE HACERSE SI YA HAY UN PLANA DE CUENTAS CARGARDO, ESTA SEGURO DE REALIZAR ESTA ACCION?", vbYesNo, "Atención")
If Respuesta = vbYes Then
    nivelesrs.Recordset.Fields(6) = login.iper
    nivelesrs.Recordset.Fields(7) = login.fper
    nivelesrs.Recordset.UpdateBatch adAffectCurrent
    Unload Me
Else
    Unload Me
End If

End Sub

Private Sub cerrar_Click()
        Unload Me
End Sub

Private Sub DataCombo1_Change()


 If DataCombo1.SelectedItem <> "" Then
    nivelesrs.Recordset.Bookmark = DataCombo1.SelectedItem
 Else
    nivelesrs.Recordset.AddNew
    empre = login.empresaact
 End If

End Sub



Private Sub Form_Load()

nivelesrs.ConnectionString = login.conexiontotal

    nivelesrs.RecordSource = "select niveles.* from niveles where empre = " & login.empresaact & " and inicioper = '" & login.iper & "' and finper = '" & login.fper & "' Order by empre"
    nivelesrs.Refresh
    
    If nivelesrs.Recordset.EOF = True Then
        nivelesrs.Recordset.AddNew
        nivelesrs.Recordset.Fields(0) = login.empresaact
    End If


End Sub

Private Sub nivel1_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        nivel2.SetFocus
    End If
    
End Sub

Private Sub nivel2_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        nivel3.SetFocus
    End If
End Sub

Private Sub nivel3_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        nivel4.SetFocus
    End If
End Sub

Private Sub nivel4_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        nivel5.SetFocus
    End If
End Sub

Private Sub nivel5_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        aceptar.SetFocus
    End If
End Sub

