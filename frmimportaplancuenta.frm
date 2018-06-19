VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{5734474E-78D3-4254-99B9-C35F31BDF509}#62.0#0"; "vbskpro2.ocx"
Begin VB.Form frmimportaplancuenta 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   5355
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7290
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5355
   ScaleWidth      =   7290
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1680
      TabIndex        =   3
      Top             =   2280
      Width           =   5415
   End
   Begin VB.ListBox List1 
      Height          =   2400
      Left            =   120
      TabIndex        =   2
      Top             =   2640
      Width           =   6975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   2160
      Width           =   1335
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frmimportaplancuenta.frx":0000
      Height          =   1935
      Left            =   120
      OleObjectBlob   =   "frmimportaplancuenta.frx":0014
      TabIndex        =   0
      Top             =   240
      Width           =   6975
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Excel 5.0;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   ""
      Top             =   0
      Visible         =   0   'False
      Width           =   1065
   End
   Begin MSAdodcLib.Adodc niveles 
      Height          =   330
      Left            =   1320
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
      Left            =   0
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
      LcK2            =   $"frmimportaplancuenta.frx":0BAB
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
Attribute VB_Name = "frmimportaplancuenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()



Dim cabecera1a As String
Dim cabecera2a As String
Dim cabecera3a As String
Dim cabecera4a As String
Dim cabecera5a As String

niveles.ConnectionString = login.conexiontotal
datcuentas.ConnectionString = login.conexiontotal
  
  niveles.RecordSource = "select niveles.* from niveles Where empre = " & login.empresaact & " and inicioper = '" & login.iper & "' and finper = '" & login.fper & "'"
  niveles.Refresh
  
 If niveles.Recordset.EOF = True Then
    mensa = MsgBox("Debe especificar el formato de codigo Contable, en: Parametros - Empresa - Digitos del Codigo Contable", vbCritical, "!! Error")
    Exit Sub
  End If
  
  
  datcuentas.RecordSource = "select cuentas.* from cuentas where empre = " & login.empresaact & ""
  datcuentas.Refresh
  

  
niv1 = niveles.Recordset.Fields("niv1")
niv2 = niveles.Recordset.Fields("niv2")
niv3 = niveles.Recordset.Fields("niv3")
niv4 = niveles.Recordset.Fields("niv4")
niv5 = niveles.Recordset.Fields("niv5")

Text1.Text = ""
    Data1.Recordset.MoveFirst
    
cabecera1 = 0
i = 9
Do While Not Data1.Recordset.EOF
    
    nomcuenta = Left(DBGrid1.Columns(0), 50)
    
    If DBGrid1.Columns(2) = "C1" Then
            cabecera2 = 0
            cabecera1 = cabecera1 + 1
            nivel = niv2
            posicion = 2
            imputable = 0
            cabecera1a = Right(Str(cabecera1), Len(Str(cabecera1)) - 1)
            codcontable = cabecera1a + "." + Mid("000000", 1, niv2) + "." + Mid("000000", 1, niv3) + "." + Mid("000000", 1, niv4) + "." + Mid("000000", 1, niv5)
            codcontable0 = cabecera1a + Mid("000000", 1, niv2) + Mid("000000", 1, niv3) + Mid("000000", 1, niv4) + Mid("000000", 1, niv5)
            List1.AddItem (codcontable + "   " + nomcuenta)
            Rem Debug.Print codcontable, nomcuenta
    End If
    If DBGrid1.Columns(2) = "C2" Then
            cabecera3 = 0
            cabecera2 = cabecera2 + 1
            nivel = niv3
            posicion = 3
            imputable = 0
            cabecera2c = Right(Str(cabecera2), Len(Str(cabecera2)) - 1)
            cabecera2a = Mid("000000", 1, niv2 - Len(cabecera2c)) + cabecera2c
            codcontable = cabecera1a + "." + cabecera2a + "." + Mid("000000", 1, niv3) + "." + Mid("000000", 1, niv4) + "." + Mid("000000", 1, niv5)
            codcontable0 = cabecera1a + cabecera2a + Mid("000000", 1, niv3) + Mid("000000", 1, niv4) + Mid("000000", 1, niv5)
            List1.AddItem (codcontable + "   " + nomcuenta)
            Rem Debug.Print codcontable, nomcuenta
    End If
    If DBGrid1.Columns(2) = "C3" Then
            cabecera4 = 0
            nivel = niv4
            posicion = 4
            imputable = 0
            cabecera3 = cabecera3 + 1
            cabecera3c = Right(Str(cabecera3), Len(Str(cabecera3)) - 1)
            cabecera3a = Mid("000000", 1, niv3 - Len(cabecera3c)) + cabecera3c
            codcontable = cabecera1a + "." + cabecera2a + "." + cabecera3a + "." + Mid("000000", 1, niv4) + "." + Mid("000000", 1, niv5)
            codcontable0 = cabecera1a + cabecera2a + cabecera3a + Mid("000000", 1, niv4) + Mid("000000", 1, niv5)
            List1.AddItem (codcontable + "   " + nomcuenta)
            Rem Debug.Print codcontable, nomcuenta
    End If
    If DBGrid1.Columns(2) = "C4" Then
            cabecera5 = 0
            nivel = niv5
            posicion = 5
            imputable = 0
            cabecera4 = cabecera4 + 1
            cabecera4c = Right(Str(cabecera4), Len(Str(cabecera4)) - 1)
            cabecera4a = Mid("000000", 1, niv4 - Len(cabecera4c)) + cabecera4c
            codcontable = cabecera1a + "." + cabecera2a + "." + cabecera3a + "." + cabecera4a + "." + Mid("000000", 1, niv5)
            codcontable0 = cabecera1a + cabecera2a + cabecera3a + cabecera4a + Mid("000000", 1, niv5)
            List1.AddItem (codcontable + "   " + nomcuenta)
            Rem Debug.Print codcontable, nomcuenta
    End If

    If DBGrid1.Columns(2) = "" Then
            imputable = imputable + 1
            imputable4c = Right(Str(imputable), Len(Str(imputable)) - 1)
            If nivel < Len(imputable4c) Then
                mensa = MsgBox("!! ERROR !! - Digitos del nivel " + Str(posicion) + " Insuficientes", vbCritical, "Atencion")
                GoTo fin
            End If
            imputable4a = Mid("000000", 1, nivel - Len(imputable4c)) + imputable4c
          
            If posicion = 2 Then
                codcontable = cabecera1a + "." + imputable4a + "." + Mid("000000", 1, niv3) + "." + Mid("000000", 1, niv4) + "." + Mid("000000", 1, niv5)
                codcontable0 = cabecera1a + imputable4a + Mid("000000", 1, niv3) + Mid("000000", 1, niv4) + Mid("000000", 1, niv5)
            End If
            If posicion = 3 Then
                codcontable = cabecera1a + "." + cabecera2a + "." + imputable4a + "." + Mid("000000", 1, niv4) + "." + Mid("000000", 1, niv5)
                codcontable0 = cabecera1a + cabecera2a + imputable4a + Mid("000000", 1, niv4) + Mid("000000", 1, niv5)
            End If
            If posicion = 4 Then
                codcontable = cabecera1a + "." + cabecera2a + "." + cabecera3a + "." + imputable4a + "." + Mid("000000", 1, niv5)
                codcontable0 = cabecera1a + cabecera2a + cabecera3a + imputable4a + Mid("000000", 1, niv5)
            End If
            If posicion = 5 Then
                codcontable = cabecera1a + "." + cabecera2a + "." + cabecera3a + "." + cabecera4a + "." + imputable4a
                codcontable0 = cabecera1a + cabecera2a + cabecera3a + cabecera4a + imputable4a
            End If
            List1.AddItem (codcontable + "   " + Mid("        ", 1, posicion * 2) + nomcuenta)
            Rem Debug.Print codcontable, nomcuenta, imputar
    End If
      datcuentas.Recordset.AddNew
      i = i + 1
      datcuentas.Recordset.Fields(0) = i
      datcuentas.Recordset.Fields(1) = nomcuenta
      datcuentas.Recordset.Fields(2) = DBGrid1.Columns(1)
      datcuentas.Recordset.Fields(3) = codcontable0
      datcuentas.Recordset.Fields(4) = codcontable
      datcuentas.Recordset.Fields(5) = login.empresaact
      datcuentas.Recordset.Fields(6) = login.iper
      datcuentas.Recordset.Fields(7) = login.fper
      datcuentas.Recordset.UpdateBatch adAffectCurrent

    Data1.Recordset.MoveNext
Loop
Exit Sub

fin:
  datcuentas.RecordSource = "select cuentas.* from cuentas where empre = " & login.empresaact & " and inicioper = '" & login.iper & "'"
  datcuentas.Refresh

If datcuentas.Recordset.EOF = True Then Exit Sub
datcuentas.Recordset.MoveFirst

Do While Not datcuentas.Recordset.EOF
    datcuentas.Recordset.Delete adAffectCurrent
    datcuentas.Recordset.MoveNext
Loop

Unload Me
frmimportaplancuenta.Show
   

End Sub

Private Sub Form_Load()

    Data1.DatabaseName = App.Path & "\plan.xls"
    Data1.RecordSource = "'Plan de cuentas'$"
    Data1.Refresh

frmimportaplancuenta.Top = 0
frmimportaplancuenta.Left = 0
End Sub
