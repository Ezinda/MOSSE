VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{D18BBD1F-82BB-4385-BED3-E9D31A3E361E}#1.0#0"; "KewlButtonz.ocx"
Begin VB.Form frmpreingreso 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PRE-INGRESO"
   ClientHeight    =   5910
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9930
   Icon            =   "frmpreingreso.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5910
   ScaleWidth      =   9930
   Begin MSAdodcLib.Adodc dattipocaña 
      Height          =   330
      Left            =   8400
      Top             =   2160
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
   Begin MSAdodcLib.Adodc datcanieros 
      Height          =   330
      Left            =   8400
      Top             =   2520
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
   Begin MSAdodcLib.Adodc datmovimientos 
      Height          =   330
      Left            =   8400
      Top             =   2880
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
   Begin MSAdodcLib.Adodc dattransporte 
      Height          =   330
      Left            =   8400
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
   Begin VB.PictureBox Picture1 
      Height          =   5535
      Left            =   120
      ScaleHeight     =   5475
      ScaleWidth      =   7755
      TabIndex        =   11
      Top             =   120
      Width           =   7815
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   2
         Left            =   2160
         TabIndex        =   6
         Top             =   3240
         Width           =   3615
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   1
         Left            =   2160
         MaxLength       =   50
         TabIndex        =   5
         Top             =   2640
         Width           =   2535
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1005
         Index           =   3
         Left            =   2160
         MaxLength       =   300
         MultiLine       =   -1  'True
         TabIndex        =   7
         Top             =   3840
         Width           =   5175
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "frmpreingreso.frx":0442
         Height          =   360
         Left            =   2160
         TabIndex        =   1
         Top             =   240
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   741
         _Version        =   393216
         Style           =   2
         ListField       =   "nombre"
         BoundColumn     =   "alias_0_id"
         Text            =   ""
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
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   0
         Left            =   2160
         MaxLength       =   50
         TabIndex        =   4
         Top             =   2040
         Width           =   5295
      End
      Begin MSDataListLib.DataCombo DataCombo2 
         Bindings        =   "frmpreingreso.frx":045C
         Height          =   360
         Left            =   2160
         TabIndex        =   2
         Top             =   720
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   741
         _Version        =   393216
         Style           =   2
         ListField       =   "alias_3_nombre"
         BoundColumn     =   "alias_0_id"
         Text            =   ""
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
      Begin MSDataListLib.DataCombo DataCombo3 
         Bindings        =   "frmpreingreso.frx":0476
         Height          =   360
         Left            =   2160
         TabIndex        =   3
         Top             =   1200
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   741
         _Version        =   393216
         Style           =   2
         ListField       =   "alias_1_nombre"
         BoundColumn     =   "alias_0_id"
         Text            =   ""
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
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Observaciones:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   18
         Top             =   3840
         Width           =   1935
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Nro. Remito:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   17
         Top             =   3240
         Width           =   1935
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Patente:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   16
         Top             =   2640
         Width           =   1935
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Chofer:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   7
         Left            =   120
         TabIndex        =   15
         Top             =   2040
         Width           =   1935
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   7560
         Y1              =   1800
         Y2              =   1800
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo de Caña:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Cañero"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   120
         TabIndex        =   13
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Transportista:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   6
         Left            =   120
         TabIndex        =   12
         Top             =   1200
         Width           =   2055
      End
   End
   Begin VB.Frame Frame3 
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
      ForeColor       =   &H00800000&
      Height          =   5655
      Left            =   8160
      TabIndex        =   0
      Top             =   0
      Width           =   1575
      Begin KewlButtonz.KewlButtons salir 
         Height          =   615
         Left            =   240
         TabIndex        =   9
         Top             =   4800
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
         MICON           =   "frmpreingreso.frx":0492
         PICN            =   "frmpreingreso.frx":04AE
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin KewlButtonz.KewlButtons grabar 
         Height          =   615
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   1085
         BTYPE           =   14
         TX              =   "&Grabar"
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
         MICON           =   "frmpreingreso.frx":0FF8
         PICN            =   "frmpreingreso.frx":1014
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin KewlButtonz.KewlButtons cancelar 
         Cancel          =   -1  'True
         Height          =   615
         Left            =   240
         TabIndex        =   10
         Top             =   1320
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   1085
         BTYPE           =   14
         TX              =   "&Cancelar"
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
         MICON           =   "frmpreingreso.frx":2A96
         PICN            =   "frmpreingreso.frx":2AB2
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
End
Attribute VB_Name = "frmpreingreso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Cancelar_Click()

    Unload Me
    frmpreingreso.Show

End Sub

Private Sub DataList1_KeyPress(KeyAscii As Integer)
On Error GoTo fuera

    If KeyAscii = 13 Then
        KeyAscii = 0
        Text1(5).Text = DataList1.BoundText
        Text1(6).SetFocus
    End If

fuera:

End Sub

Private Sub DataList1_LostFocus()

    DataList1.Visible = False

End Sub

Private Sub DataCombo1_KeyPress(KeyAscii As Integer)
On Error Resume Next

    If KeyAscii = 13 Then
        KeyAscii = 0
        DataCombo2.SetFocus
    End If

End Sub

Private Sub DataList2_KeyPress(KeyAscii As Integer)
On Error GoTo fuera

    If KeyAscii = 13 Then
        KeyAscii = 0
        Text1(6).Text = DataList2.BoundText
        grabar.SetFocus
    End If

fuera:
End Sub

Private Sub DataList2_LostFocus()

    DataList2.Visible = False

End Sub

Private Sub DataList3_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        Text1(4).Text = DataList3.Text
        Text1(5).SetFocus
        DataList3.Visible = False
    End If


End Sub

Private Sub DataCombo2_KeyPress(KeyAscii As Integer)

On Error Resume Next

    If KeyAscii = 13 Then
        KeyAscii = 0
        DataCombo3.SetFocus
    End If


End Sub

Private Sub DataCombo3_KeyPress(KeyAscii As Integer)

On Error Resume Next

    If KeyAscii = 13 Then
        KeyAscii = 0
        Text1(0).SetFocus
    End If

End Sub

Private Sub Form_Activate()


DataCombo1.SetFocus


End Sub

Private Sub Form_Load()
Aplicar_skin Me

frmpreingreso.Top = 0
frmpreingreso.Left = 0


dattipocaña.ConnectionString = login.conexiontotal
datcanieros.ConnectionString = login.conexiontotal
dattransporte.ConnectionString = login.conexiontotal
datmovimientos.ConnectionString = login.conexiontotal

    dattipocaña.RecordSource = "SELECT ID AS ALIAS_0_ID, NOMBRE FROM V_ITEMTIPOCLASIFICADOR_ AS ALIAS_0 " & _
                               "WHERE (BO_PLACE_ID = '{8CCBA4D1-EDDE-432A-B63E-C8AC0AC3DE2F}') AND (ACTIVESTATUS <> 2) ORDER BY NOMBRE"
    dattipocaña.Refresh

    datcanieros.RecordSource = "SELECT     ALIAS_0.ID AS ALIAS_0_ID, ALIAS_0.CODIGO AS ALIAS_0_CODIGO, ALIAS_3.NOMBRE AS ALIAS_3_NOMBRE, V_UD_CLIENTE_.PRODUCTOR " & _
                               "FROM V_CLIENTE_ AS ALIAS_0 LEFT OUTER JOIN V_UD_CLIENTE_ ON ALIAS_0.BOEXTENSION_ID = V_UD_CLIENTE_.ID LEFT OUTER JOIN V_PERSONA_ AS ALIAS_3 ON ALIAS_0.ENTEASOCIADO_ID = ALIAS_3.ID " & _
                               "WHERE (ALIAS_0.BO_PLACE_ID = '{89C234C8-3F01-11D5-86AD-0080AD403F5F}') AND (ALIAS_0.ACTIVESTATUS = 0) AND (V_UD_CLIENTE_.PRODUCTOR = 'T') ORDER BY ALIAS_3.NOMBRE "
    datcanieros.Refresh

    dattransporte.RecordSource = "SELECT     ALIAS_0.ID AS ALIAS_0_ID, ALIAS_1.NOMBRE AS ALIAS_1_NOMBRE FROM V_MEDIOTRANSPORTE_ AS ALIAS_0 LEFT OUTER JOIN " & _
                                 "V_PERSONA_ AS ALIAS_1 ON ALIAS_0.ENTEASOCIADO_ID = ALIAS_1.ID WHERE     (ALIAS_0.BO_PLACE_ID = '{76C697C2-3DAE-11D5-B059-004854841C8A}') AND (ALIAS_0.ACTIVESTATUS = 0) " & _
                                 "ORDER BY ALIAS_1_NOMBRE"
    dattransporte.Refresh
    
    
    datmovimientos.RecordSource = "select top 1 * from pr_ezi_movimientos where id_movimiento = 0"
    datmovimientos.Refresh
    
   
End Sub

Private Sub grabar_Click()
On Error GoTo errorgrabar

    mensa = MsgBox("Desea Grabar este Registro (s/n) ?", vbYesNo, "!! Atención !!")
    If mensa = vbNo Then Exit Sub


    datmovimientos.Recordset.AddNew
    datmovimientos.Recordset.Fields("id_tipo_cana") = DataCombo1.BoundText
    datmovimientos.Recordset.Fields("id_caniero") = DataCombo2.BoundText
    datmovimientos.Recordset.Fields("id_transportista") = DataCombo3.BoundText
    datmovimientos.Recordset.Fields("razon_social") = DataCombo2.Text
    datmovimientos.Recordset.Fields("transporte") = DataCombo3.Text
    datmovimientos.Recordset.Fields("chofer") = Text1(0).Text
    datmovimientos.Recordset.Fields("patente") = Text1(1).Text
    datmovimientos.Recordset.Fields("observaciones") = Text1(3).Text
    datmovimientos.Recordset.Fields("prepesada") = "T"
    datmovimientos.Recordset.Fields("fecha_entrada") = Str(Date)
    datmovimientos.Recordset.Fields("hora_entrada") = Str(Time)
    datmovimientos.Recordset.Fields("tipo_pesada") = "C"
    datmovimientos.Recordset.Fields("remito") = Text1(2).Text
    datmovimientos.Recordset.Fields("usuario") = login.usuarioactivo
    
    datmovimientos.Recordset.UpdateBatch adAffectCurrent
    mensa = MsgBox("Nro de Movimiento: " + Str(datmovimientos.Recordset.Fields("id_movimiento")), vbInformation, "Grabado Correctamente")
    
    Call Cancelar_Click
    
Exit Sub
errorgrabar:
    mensa = MsgBox("No se pudo registrar la información", vbCritical, "Error !!")
    Text1(0).SetFocus

End Sub

Private Sub List1_KeyPress(KeyAscii As Integer)
On Error GoTo fuera

    If KeyAscii = 13 Then
        KeyAscii = 0
        Text1(4).Text = List1.ListIndex + 1
        Text1(5).SetFocus
    End If

fuera:
End Sub

Private Sub List1_LostFocus()

    List1.Visible = False

End Sub


Private Sub salir_Click()

    Unload Me

End Sub


Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
On Error Resume Next

    If KeyAscii = 13 Then
        KeyAscii = 0
        Text1(Index).Text = UCase(Text1(Index).Text)
        If Index = 3 Then
            grabar.SetFocus
        Else
            Text1(Index + 1).SetFocus
        End If
        
    End If
    
End Sub

Private Sub Text1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)

If KeyCode = 38 Then
    If Index > 0 Then
      Text1(Index - 1).SetFocus
    Else
      DataCombo3.SetFocus
    End If
End If

End Sub

Private Sub Text1_LostFocus(Index As Integer)
On Error Resume Next
        If Index = 2 Then
            If Len(Text1(2).Text) = 12 Then Exit Sub
            For X = 1 To Len(Text1(2).Text)
               car = Mid(Text1(2).Text, X, 1)
               If car = "-" Then
                  PVta = Right("0000" + Left(Text1(2).Text, X - 1), 4)
                  nu = Right("00000000" + Right(Text1(2).Text, Len(Text1(2).Text) - X), 8)
                  Text1(2).Text = PVta + nu
                  Exit Sub
               End If
            Next X
            Text1(2).Text = Right("00000000" + Text1(2).Text, 8)
        End If

End Sub
