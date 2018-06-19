VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{5734474E-78D3-4254-99B9-C35F31BDF509}#62.0#0"; "vbskpro2.ocx"
Object = "{D18BBD1F-82BB-4385-BED3-E9D31A3E361E}#1.0#0"; "KewlButtonz.ocx"
Begin VB.Form importacuentaotraempresa 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Importa Parametrizacin Otra Empresa"
   ClientHeight    =   5880
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6480
   Icon            =   "importacuentaotraempresa.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5880
   ScaleWidth      =   6480
   Begin VB.CommandButton Command6 
      Caption         =   "Periodo Contable:"
      Height          =   255
      Index           =   2
      Left            =   600
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   240
      Width           =   1815
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Empresa a Importar:"
      Height          =   255
      Index           =   0
      Left            =   600
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   600
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   2520
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   240
      Width           =   2535
   End
   Begin MSComctlLib.ProgressBar bar 
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   4800
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "importacuentaotraempresa.frx":0442
      Height          =   3615
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   6376
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
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
      ColumnCount     =   8
      BeginProperty Column00 
         DataField       =   "Cod Contable"
         Caption         =   "Cod.Abrev"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "Nombre Cuenta"
         Caption         =   "Nombre Cuenta"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "imp"
         Caption         =   "Imputable"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "Id Cuenta"
         Caption         =   "Cod.Contable"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "idcuenta"
         Caption         =   "Cod.Contable"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "empre"
         Caption         =   "empre"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "inicioper"
         Caption         =   "inicioper"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column07 
         DataField       =   "finper"
         Caption         =   "finper"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            Locked          =   -1  'True
         EndProperty
         BeginProperty Column01 
            Locked          =   -1  'True
         EndProperty
         BeginProperty Column02 
            Locked          =   -1  'True
         EndProperty
         BeginProperty Column03 
            Locked          =   -1  'True
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column04 
            Locked          =   -1  'True
            Object.Visible         =   -1  'True
         EndProperty
         BeginProperty Column05 
            Locked          =   -1  'True
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column06 
            Locked          =   -1  'True
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column07 
            Locked          =   -1  'True
            Object.Visible         =   0   'False
         EndProperty
      EndProperty
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Bindings        =   "importacuentaotraempresa.frx":045D
      DataSource      =   "dataempresa"
      Height          =   315
      Left            =   2520
      TabIndex        =   0
      Top             =   600
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   556
      _Version        =   393216
      BackColor       =   -2147483626
      ListField       =   "razonsocial"
      BoundColumn     =   "empresa"
      Text            =   ""
   End
   Begin MSAdodcLib.Adodc datPrimaryRS 
      Height          =   330
      Left            =   4800
      Top             =   5400
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
      EOFAction       =   1
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
   Begin MSAdodcLib.Adodc dataempresa 
      Height          =   330
      Left            =   4800
      Top             =   5160
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
      LcK1            =   "3.66*/4/0*/1-5*210/."
      LcK2            =   $"importacuentaotraempresa.frx":0477
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
   Begin MSAdodcLib.Adodc nivelesrs 
      Height          =   375
      Left            =   120
      Top             =   5160
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
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
   Begin MSAdodcLib.Adodc datcolumnascompras 
      Height          =   330
      Left            =   120
      Top             =   5520
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
   Begin MSAdodcLib.Adodc datcolumnasventas 
      Height          =   330
      Left            =   1320
      Top             =   5520
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
   Begin KewlButtonz.KewlButtons Command1 
      Height          =   615
      Left            =   2520
      TabIndex        =   4
      Top             =   5160
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1085
      BTYPE           =   14
      TX              =   "&Crear"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   -2147483629
      BCOLO           =   -2147483629
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "importacuentaotraempresa.frx":0486
      PICN            =   "importacuentaotraempresa.frx":04A2
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
End
Attribute VB_Name = "importacuentaotraempresa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim pi As Date
Dim pf As Date


Private Sub Command1_Click()
On Error GoTo fuera
Dim campo0(5000)
Dim campo1(5000)
Dim campo2(5000)
Dim campo3(5000)
Dim campo4(5000)
Dim campo5(5000)
Dim campo6(5000)
Dim campo7(5000)
Dim niveles(5)
Dim compras(62)
Dim ventas(66)

x = 0
datprimaryrs.Recordset.MoveFirst

bar.Min = 0
bar.max = datprimaryrs.Recordset.RecordCount

paso1:
    If datprimaryrs.Recordset.EOF = True Then GoTo paso2
    campo0(x) = datprimaryrs.Recordset.Fields(0)
    campo1(x) = datprimaryrs.Recordset.Fields(1)
    campo2(x) = datprimaryrs.Recordset.Fields(2)
    campo3(x) = datprimaryrs.Recordset.Fields(3)
    campo4(x) = datprimaryrs.Recordset.Fields(4)
    campo5(x) = datprimaryrs.Recordset.Fields(5)
    campo6(x) = datprimaryrs.Recordset.Fields(6)
    campo7(x) = datprimaryrs.Recordset.Fields(7)
    bar.Value = x
    If datprimaryrs.Recordset.EOF = False Then
        datprimaryrs.Recordset.MoveNext
        x = x + 1
        GoTo paso1
    End If
paso2:

For Y = 0 To x - 1
    datprimaryrs.Recordset.AddNew
    datprimaryrs.Recordset.Fields(0) = campo0(Y)
    datprimaryrs.Recordset.Fields(1) = campo1(Y)
    datprimaryrs.Recordset.Fields(2) = campo2(Y)
    datprimaryrs.Recordset.Fields(3) = campo3(Y)
    datprimaryrs.Recordset.Fields(4) = campo4(Y)
    datprimaryrs.Recordset.Fields(5) = login.empresaact
    datprimaryrs.Recordset.Fields(6) = login.iper
    datprimaryrs.Recordset.Fields(7) = login.fper
    datprimaryrs.Recordset.UpdateBatch adAffectCurrent
    bar.Value = Y
Next Y
    
For x = 0 To 5
     niveles(x) = nivelesrs.Recordset.Fields(x)
Next x

    nivelesrs.Recordset.AddNew
For x = 0 To 5
    nivelesrs.Recordset.Fields(x) = niveles(x)
Next x
    nivelesrs.Recordset.Fields(0) = login.empresaact
    nivelesrs.Recordset.Fields("inicioper") = login.iper
    nivelesrs.Recordset.Fields("finper") = login.fper
    nivelesrs.Recordset.UpdateBatch adAffectCurrent
    
    
For x = 0 To 62
    compras(x) = datcolumnascompras.Recordset.Fields(x)
Next x
    datcolumnascompras.Recordset.AddNew
For x = 0 To 62
    datcolumnascompras.Recordset.Fields(x) = compras(x)
Next x
    datcolumnascompras.Recordset.Fields(0) = login.empresaact
    datcolumnascompras.Recordset.Fields("inicioper") = login.iper
    datcolumnascompras.Recordset.Fields("finper") = login.fper
    datcolumnascompras.Recordset.UpdateBatch adAffectCurrent
    
    
For x = 0 To 65
    ventas(x) = datcolumnasventas.Recordset.Fields(x)
Next x
    datcolumnasventas.Recordset.AddNew
For x = 0 To 65
    datcolumnasventas.Recordset.Fields(x) = ventas(x)
Next x
    datcolumnasventas.Recordset.Fields(0) = login.empresaact
    datcolumnasventas.Recordset.Fields("inicioper") = login.iper
    datcolumnasventas.Recordset.Fields("finper") = login.fper
    datcolumnasventas.Recordset.Fields(65) = ventas(65)
    datcolumnasventas.Recordset.UpdateBatch adAffectCurrent

    Inicio.datauditoria.ConnectionString = login.conexiontotal
    
    Inicio.datauditoria.RecordSource = "select auditoria.* from auditoria"
    Inicio.datauditoria.Refresh
    
    Inicio.datauditoria.Recordset.AddNew
    Inicio.datauditoria.Recordset.Fields("fecha") = Date
    Inicio.datauditoria.Recordset.Fields("hora") = Str(Time)
    Inicio.datauditoria.Recordset.Fields("ventana") = "Importacion de Parametrizacion"
    Inicio.datauditoria.Recordset.Fields("accion") = "Importa Parametrizacion de empresa:" + DataCombo1.Text
    Inicio.datauditoria.Recordset.Fields("usuario") = login.usuarioactivo
    Inicio.datauditoria.Recordset.Fields("empresa") = login.empresaact
    Inicio.datauditoria.Recordset.UpdateBatch adAffectCurrent


Exit Sub
fuera:
    mensa = MsgBox("Esta empresa ya fue importada anteriormente", vbCritical, "Error")

End Sub

Private Sub DataCombo1_Change()
        
  dataempresa.RecordSource = "select empresa.* from empresa Where empresa <> " & login.empresaact & " and razonsocial = '" & DataCombo1.Text & "' Order by inicioperiodo"
  dataempresa.Refresh

  pi = dataempresa.Recordset.Fields("inicioperiodo")
  pf = dataempresa.Recordset.Fields("finperiodo")

Text1.Text = Str(dataempresa.Recordset.Fields("inicioperiodo")) + " --" + Str(dataempresa.Recordset.Fields("finperiodo"))
bandera = DataCombo1.Text

  dataempresa.RecordSource = "select empresa.* from empresa Where empresa <> " & login.empresaact & " Order by inicioperiodo"
  dataempresa.Refresh

DataCombo1.Text = bandera


  datprimaryrs.RecordSource = "select cuentas.* from cuentas WHERE cuentas.empre = " & DataCombo1.BoundText & " and inicioper = '" & pi & "' and finper = '" & pf & "'  ORDER BY IDCUENTA"
  datprimaryrs.Refresh
  
  datcolumnascompras.RecordSource = "SELECT columnascompra.* FROM columnascompra WHERE empresa = " & DataCombo1.BoundText & " and inicioper = '" & pi & "'"
  datcolumnascompras.Refresh
  datcolumnasventas.RecordSource = "SELECT columnasventa.* FROM columnasventa WHERE empresa = " & DataCombo1.BoundText & " and inicioper = '" & pi & "'"
  datcolumnasventas.Refresh
  nivelesrs.RecordSource = "select niveles.* from niveles where empre = " & DataCombo1.BoundText & " and inicioper = '" & pi & "'"
  nivelesrs.Refresh
  
  DataGrid1.Refresh

End Sub

Private Sub Form_Load()
Aplicar_skin Me

dataempresa.ConnectionString = login.conexiontotal
datprimaryrs.ConnectionString = login.conexiontotal
datcolumnascompras.ConnectionString = login.conexiontotal
datcolumnasventas.ConnectionString = login.conexiontotal
nivelesrs.ConnectionString = login.conexiontotal
  

  datprimaryrs.RecordSource = "select cuentas.* from cuentas WHERE cuentas.empre = " & login.empresaact & " and inicioper = '" & login.iper & "' and finper = '" & login.fper & "'  ORDER BY IDCUENTA"
  datprimaryrs.Refresh
  
  dataempresa.RecordSource = "select empresa.* from empresa Where empresa <> " & login.empresaact & " Order by inicioperiodo"
  dataempresa.Refresh
  

  
End Sub


Private Sub lblLabels_Click(Index As Integer)

End Sub
