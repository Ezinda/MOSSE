VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{D18BBD1F-82BB-4385-BED3-E9D31A3E361E}#1.0#0"; "KewlButtonz.ocx"
Begin VB.Form login 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Punto de Venta"
   ClientHeight    =   4485
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10590
   ControlBox      =   0   'False
   Icon            =   "login.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4485
   ScaleWidth      =   10590
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text9 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   6480
      TabIndex        =   21
      Top             =   1680
      Width           =   1815
   End
   Begin VB.TextBox Text8 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   6480
      TabIndex        =   20
      Top             =   720
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Clave Nueva"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   3
      Left            =   4920
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   1560
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Clave Anterior"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   4920
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   600
      Width           =   1455
   End
   Begin KewlButtonz.KewlButtons aceptar 
      Height          =   615
      Left            =   1560
      TabIndex        =   15
      Top             =   3720
      Visible         =   0   'False
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1085
      BTYPE           =   14
      TX              =   "&Ingresar"
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
      BCOL            =   -2147483645
      BCOLO           =   -2147483645
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "login.frx":0442
      PICN            =   "login.frx":045E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.PictureBox Picture2 
      Height          =   735
      Left            =   960
      ScaleHeight     =   675
      ScaleWidth      =   2955
      TabIndex        =   12
      Top             =   120
      Width           =   3015
      Begin VB.Label Label1 
         Caption         =   "Soluciones en Tecnología de la Información"
         Height          =   495
         Left            =   0
         TabIndex        =   13
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label Label2 
         Caption         =   "EZINDA"
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
         Left            =   0
         TabIndex        =   14
         Top             =   0
         Width           =   1095
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   780
      Left            =   0
      ScaleHeight     =   720
      ScaleWidth      =   765
      TabIndex        =   11
      Top             =   120
      Width           =   825
      Begin VB.Image Image1 
         BorderStyle     =   1  'Fixed Single
         Height          =   735
         Left            =   0
         Picture         =   "login.frx":0E70
         Stretch         =   -1  'True
         Top             =   0
         Width           =   855
      End
   End
   Begin KewlButtonz.KewlButtons command1 
      Height          =   615
      Left            =   1680
      TabIndex        =   10
      Top             =   2400
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1085
      BTYPE           =   14
      TX              =   "&Ingresar"
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
      BCOL            =   -2147483645
      BCOLO           =   -2147483645
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "login.frx":1467A
      PICN            =   "login.frx":14696
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Clave"
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
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Usuario"
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
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1200
      Width           =   1215
   End
   Begin VB.TextBox Text7 
      Height          =   285
      Left            =   1560
      MultiLine       =   -1  'True
      TabIndex        =   7
      Text            =   "login.frx":150A8
      Top             =   3600
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Text6 
      DataField       =   "fecha"
      DataSource      =   "datultimodi"
      Height          =   285
      Left            =   2520
      TabIndex        =   6
      Text            =   "Text6"
      Top             =   3360
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2040
      TabIndex        =   1
      Top             =   1920
      Width           =   1815
   End
   Begin VB.TextBox Text4 
      DataField       =   "direccionelectronica"
      DataSource      =   "datusuarios"
      Height          =   285
      Left            =   1800
      TabIndex        =   5
      Text            =   "Text4"
      Top             =   1800
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      DataField       =   "nombre"
      DataSource      =   "datusuarios"
      Height          =   285
      Left            =   120
      TabIndex        =   4
      Text            =   "Text3"
      Top             =   3240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      DataField       =   "usuario"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2040
      TabIndex        =   0
      Top             =   1200
      Width           =   1815
   End
   Begin VB.TextBox text1 
      DataField       =   "empresaactiva"
      DataSource      =   "datinicio"
      Height          =   285
      Left            =   0
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   2880
      Visible         =   0   'False
      Width           =   615
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "login.frx":150AE
      Height          =   975
      Left            =   360
      TabIndex        =   2
      Top             =   2640
      Visible         =   0   'False
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   1720
      _Version        =   393216
      AllowUpdate     =   -1  'True
      BackColor       =   14737632
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
         Weight          =   700
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
   Begin MSAdodcLib.Adodc datempresa 
      Height          =   330
      Left            =   1320
      Top             =   3120
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
      DataSourceName  =   ""
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
   Begin KewlButtonz.KewlButtons KewlButtons1 
      Height          =   615
      Left            =   120
      TabIndex        =   16
      Top             =   3720
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1085
      BTYPE           =   14
      TX              =   "&Salir"
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
      BCOL            =   -2147483645
      BCOLO           =   -2147483645
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "login.frx":150C7
      PICN            =   "login.frx":150E3
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSAdodcLib.Adodc datusuarios 
      Height          =   330
      Left            =   3240
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
      DataSourceName  =   ""
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
   Begin KewlButtonz.KewlButtons KewlButtons2 
      Height          =   615
      Left            =   2760
      TabIndex        =   17
      Top             =   3720
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   1085
      BTYPE           =   14
      TX              =   "&Cambiar Clave"
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
      BCOL            =   -2147483645
      BCOLO           =   -2147483645
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "login.frx":15AF5
      PICN            =   "login.frx":15B11
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin KewlButtonz.KewlButtons cambioclave 
      Height          =   615
      Left            =   6240
      TabIndex        =   22
      Top             =   2520
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1085
      BTYPE           =   14
      TX              =   "Guardar Cambio"
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
      BCOL            =   -2147483645
      BCOLO           =   -2147483645
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "login.frx":15F63
      PICN            =   "login.frx":15F7F
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSAdodcLib.Adodc datcontrol 
      Height          =   330
      Left            =   5040
      Top             =   3600
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
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "datcontrol"
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
Attribute VB_Name = "login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public usuarioactivo As String
Public nomsucursal As String
Public facturaautomatica As String
Public empresaact As String
Public empresaact1 As Integer
Public empresaact2 As Integer
Public iper As Date
Public fper As Date
Public razonsoc As String
Public habcc As Boolean
Public cap2 As String
Public plancuentasaltas As String
Public plancuentasmodi As String
Public plancuentasbajas As String
Public livacomprasmodi As String
Public livacomprascerrar As String
Public livaventasmodi As String
Public livaventascerrar As String
Public minutasmodi As String
Public minutasbajas As String
Public empresasmodi As String
Public empresasbajas As String
Public provaltas As String
Public provmodi As String
Public provbajas As String
Public clientesingre As String
Public clientesaltas As String
Public clientesmod As String
Public clientesbajas As String
Public administrador As String
Public conexiontotal As String
Public conexionreporte As String
Public nomempresa As String
Public librocerrado As String
Public administ As String
Public contraseña As String
Public nuevacontraseña As String
Public puntodecimal As Integer
Public ajuestesec As String
Public planunificado As String
Public librocajamodif As String
Public librocajalistar As String
Public ajusteclientes As String
Public basededatos As String
Public librofactura As String
Public nombrebd As String


Private Sub cambioclave_Click()

    datusuarios.RecordSource = "select usuario as nombre,clave as direccionelectronica from ud_ezi_seguridad WHERE usuario = '" & login.usuarioactivo & "'"
    datusuarios.Refresh
    
    datempresa.RecordSource = "SELECT V_COMPANIA_.CODIGO, V_PERSONAJURIDICA_.NOMBRE FROM V_COMPANIA_ INNER JOIN " & _
                            "V_PERSONAJURIDICA_ ON V_COMPANIA_.PERSONAJURIDICA_ID = V_PERSONAJURIDICA_.ID"
    datempresa.Refresh
  

    If Text2.Text <> Text3.Text Or Text2.Text = "" Then
        mensa = MsgBox("Este usuario no existe", vbCritical, "!! Error !!")
        Text2.SetFocus
        Exit Sub
    End If
    
    If Text8.Text <> Text4.Text Then
        mensa = MsgBox("Password actual incorrecto", vbCritical, "!! Error !!")
        Text8.Text = ""
        Text9.Text = ""
        Text5.SetFocus
        login.Width = 4750
        Exit Sub
    End If
    
    datusuarios.Recordset.Fields("direccionelectronica") = Text9.Text
    datusuarios.Recordset.UpdateBatch adAffectCurrent
    mensa = MsgBox("Cambio realizado con exito", vbCritical, "Password")
    Text5.Text = Text9.Text
    login.Width = 4750
    Call Command1_Click
    
       


End Sub

Private Sub Command1_Click()



    datusuarios.RecordSource = "select usuario as nombre,clave as direccionelectronica from ud_ezi_seguridad WHERE usuario = '" & login.usuarioactivo & "'"
    datusuarios.Refresh
    
    datempresa.RecordSource = "SELECT V_COMPANIA_.CODIGO, V_PERSONAJURIDICA_.NOMBRE FROM V_COMPANIA_ INNER JOIN " & _
                            "V_PERSONAJURIDICA_ ON V_COMPANIA_.PERSONAJURIDICA_ID = V_PERSONAJURIDICA_.ID"
    datempresa.Refresh
  

    If Text2.Text <> Text3.Text Then
        mensa = MsgBox("Este usuario no existe", vbCritical, "!! Error !!")
        Text2.SetFocus
        Exit Sub
    End If
    
    If Text5.Text <> Text4.Text Or Text5.Text = "" Then
        mensa = MsgBox("Password incorrecto", vbCritical, "!! Error !!")
        Text5.SetFocus
        Exit Sub
    End If
    
  
    If DataGrid1.VisibleRows = 1 Then
        empresaact = DataGrid1.Columns(0).Text
        nomempresa = DataGrid1.Columns(1).Text
        Text1.Text = empresaact
        Call aceptar_Click
        Exit Sub
    End If

    

    DataGrid1.Visible = True
    DataGrid1.SetFocus
    aceptar.Visible = True

End Sub

Private Sub aceptar_Click()
On Error Resume Next


        
  
  Inicio.Caption = DataGrid1.Columns(1).Text
  razonsoc = DataGrid1.Columns(1).Text
           
    Inicio.Show
  
    
 Rem   Inicio.archivo.Enabled = True
 Rem   Inicio.empresas.Enabled = True
 Rem   Inicio.usuarios.Enabled = True
 Rem   Inicio.proveedores.Enabled = True
 Rem   Inicio.menuasientos.Enabled = True
 Rem   Inicio.arclientes.Enabled = True
 Rem   Inicio.Toolbar1.Enabled = True
    Unload Me

End Sub




Private Sub DataGrid1_Click()

    Text1.Text = DataGrid1.Columns(0).Text
    empresaact = Text1.Text
    nomempresa = DataGrid1.Columns(1).Text

End Sub

Private Sub DataGrid1_GotFocus()
        Command1.Visible = False

End Sub

Private Sub DataGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
Text1.Text = DataGrid1.Columns(0).Text
End Sub

Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            KeyAscii = 0
            aceptar.SetFocus
        End If
        
End Sub

Private Sub DataGrid1_KeyUp(KeyCode As Integer, Shift As Integer)
    
    Text1.Text = DataGrid1.Columns(0).Text
    empresaact = Text1.Text
    nomempresa = DataGrid1.Columns(1).Text

End Sub

Private Sub Form_Load()
Rem On Error GoTo malaconexion

Aplicar_skin Me

Dim pruebita As Currency
Dim ruta As String
Dim ret As Long

crlf = Chr(13) & Chr(10)
Text7.Text = ""
Open App.Path & "\sucursal.ini" For Input As #1
While Not EOF(1)
Line Input #1, file_data$
Text7.Text = Text7.Text & file_data$ & crlf
Wend
Close #1

login.Width = 4750

nomsucursal = ""
nombredsn = ""
nombrebd = ""
For X = 1 To Len(Text7.Text)
    If Mid(Text7.Text, X, 1) = ";" Then
      For Y = X + 1 To Len(Text7.Text)
        If Mid(Text7.Text, Y, 1) = ";" Then
            For Z = Y + 1 To Len(Text7.Text)
                nombrebd = nombrebd + Mid(Text7.Text, Z, 1)
            Next Z
            GoTo paso0
        End If
        nombredsn = nombredsn + Mid(Text7.Text, Y, 1)
      Next Y
    End If
    nomsucursal = nomsucursal + Mid(Text7.Text, X, 1)
Next X

paso0:

cuerpo1 = "Provider=MSDASQL.1;Password=1;Persist Security Info=True;User ID=fs1;"
cuerpo1b = "Data Source="
cuerpo2 = nombredsn
cuerpo3 = ";Initial Catalog="
cuerpo3b = nombrebd
basededatos = nombrebd
conexiontotal = cuerpo1 + cuerpo1b + cuerpo2 + cuerpo3 + cuerpo3b

cuerpo4 = "PROVIDER=MSDASQL;dsn="
cuerpo5 = nombredsn
cuerpo6 = ";uid=fs1;pwd=1;database="
cuerpo7 = nombrebd
conexionreporte = cuerpo4 + cuerpo5 + cuerpo6 + cuerpo7 + ";"


datusuarios.ConnectionString = conexiontotal
datempresa.ConnectionString = conexiontotal
datcontrol.ConnectionString = conexiontotal

DataGrid1.Columns(0).Width = 300
DataGrid1.Columns(1).Width = 3000

Text1 = 0
Text2.Text = ""
Text5.Text = ""
pruebita = 1.1 * 1

If Mid(pruebita, 2, 1) = "." Then
    puntodecimal = 0
    Exit Sub
Else
    puntodecimal = 1
End If


ruta = App.Path & "\confregional.exe"

ret = Shell(ruta, vbNormalFocus)

Exit Sub

malaconexion:
    mensa = MsgBox("Error de Conexcion, Verifique archivo SUCURSAL.INI", vbCritical, "!Error¡")
Unload Me
    
End Sub

Private Sub KewlButtons1_Click()


    Unload Me
    Inicio.salida = 0
    Unload Inicio
    
End Sub

Private Sub Label3_Click()

End Sub

Private Sub KewlButtons2_Click()
    
    For X = login.Width To 8640 Step 50
        login.Width = X
    Next X
    Text8.SetFocus

End Sub

Private Sub text2_Change()
    usuarioactivo = Text2.Text
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        Text2.Text = LCase(Text2.Text)
        Text5.SetFocus
    End If
                  
    
End Sub

Private Sub Text2_LostFocus()

Text2.Text = LCase(Text2.Text)

End Sub

Private Sub Text5_Change()
 Text5.PasswordChar = "*"
 contraseña = Text5.Text
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
On Error GoTo fuera
    If KeyAscii = 13 Then
        KeyAscii = 0
        Command1.SetFocus
    End If
fuera:
End Sub

Private Sub Text8_Change()
 Text8.PasswordChar = "*"
 contraseña = Text8.Text
End Sub

Private Sub Text8_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        Text9.SetFocus
    End If
            
End Sub

Private Sub Text9_Change()
 Text9.PasswordChar = "*"
 nuevacontraseña = Text9.Text
End Sub

Private Sub Text9_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        cambioclave.SetFocus
    End If
            

End Sub
