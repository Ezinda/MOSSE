VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{5734474E-78D3-4254-99B9-C35F31BDF509}#62.0#0"; "vbskpro2.ocx"
Begin VB.Form frmresultadosactivos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Resultados Activos"
   ClientHeight    =   4665
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4665
   ScaleWidth      =   6030
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3960
      UseMaskColor    =   -1  'True
      Width           =   1575
   End
   Begin VB.ListBox List1 
      Height          =   3210
      Left            =   480
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   480
      Width           =   4935
   End
   Begin MSAdodcLib.Adodc datparamresultados 
      Height          =   330
      Left            =   840
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
      LcK2            =   $"frmresultadosactivos.frx":0000
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
Attribute VB_Name = "frmresultadosactivos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer
Dim act(1000) As Integer
Dim id(1000) As Integer
Private Sub Command1_Click()

For x = 0 To i - 1
    If List1.Selected(x) = True Then
            datparamresultados.RecordSource = "select paramresultados.*  from paramresultados where empresa = " & login.empresaact & " and id = " & id(x) & " "
            datparamresultados.Refresh
            datparamresultados.Recordset.Fields("activado") = -1
            datparamresultados.Recordset.UpdateBatch adAffectCurrent
    Else
            datparamresultados.RecordSource = "select paramresultados.*  from paramresultados where empresa = " & login.empresaact & " and id = " & id(x) & " "
            datparamresultados.Refresh
            datparamresultados.Recordset.Fields("activado") = 0
            datparamresultados.Recordset.UpdateBatch adAffectCurrent
    End If
Next x

Unload Me
frmresultados.Show

End Sub

Private Sub Form_Load()

frmresultadosactivos.Left = 0
frmresultadosactivos.Top = 0


    datparamresultados.ConnectionString = login.conexiontotal


    datparamresultados.RecordSource = "select paramresultados.*  from paramresultados where empresa = " & login.empresaact & "  order by id"
    datparamresultados.Refresh
    
    If datparamresultados.Recordset.EOF = True Then Exit Sub
    
List1.Clear
datparamresultados.Recordset.MoveFirst
i = 0
Do While Not datparamresultados.Recordset.EOF
      act(i) = datparamresultados.Recordset.Fields("activado")
      id(i) = datparamresultados.Recordset.Fields("id")
      List1.AddItem datparamresultados.Recordset.Fields("nombrefondo")
      datparamresultados.Recordset.MoveNext
      i = i + 1
Loop
 
For x = 0 To i - 1
    If act(x) <> 0 Then
        List1.Selected(x) = True
    Else
        List1.Selected(x) = False
    End If
Next x

End Sub
