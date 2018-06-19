VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{D18BBD1F-82BB-4385-BED3-E9D31A3E361E}#1.0#0"; "KewlButtonz.ocx"
Begin VB.Form frmvendedores 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Perfil de Vendedores"
   ClientHeight    =   6945
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11220
   Icon            =   "frmvendedores.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6945
   ScaleWidth      =   11220
   Begin VB.CommandButton Command1 
      Caption         =   "Vendedores En Calipso"
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
      Left            =   6960
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   240
      Width           =   2415
   End
   Begin VB.CommandButton llena 
      Caption         =   "llena"
      Height          =   255
      Left            =   480
      TabIndex        =   36
      Top             =   6720
      Visible         =   0   'False
      Width           =   1575
   End
   Begin MSAdodcLib.Adodc datusuarios 
      Height          =   330
      Left            =   1200
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmvendedores.frx":0442
      Height          =   6255
      Left            =   6960
      TabIndex        =   0
      Top             =   480
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   11033
      _Version        =   393216
      BackColor       =   -2147483626
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
   Begin VB.Frame Frame2 
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
      Height          =   1095
      Left            =   120
      TabIndex        =   17
      Top             =   5640
      Width           =   6615
      Begin KewlButtonz.KewlButtons grabar 
         Height          =   495
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   873
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
         MICON           =   "frmvendedores.frx":045E
         PICN            =   "frmvendedores.frx":047A
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
         Height          =   495
         Left            =   4200
         TabIndex        =   15
         Top             =   360
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   873
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
         MICON           =   "frmvendedores.frx":1EFC
         PICN            =   "frmvendedores.frx":1F18
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin KewlButtonz.KewlButtons salir 
         Height          =   495
         Left            =   5400
         TabIndex        =   16
         Top             =   360
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   873
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
         MICON           =   "frmvendedores.frx":292A
         PICN            =   "frmvendedores.frx":2946
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin KewlButtonz.KewlButtons alta 
         Height          =   495
         Left            =   1560
         TabIndex        =   13
         Top             =   360
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
         BTYPE           =   14
         TX              =   "&Alta"
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
         MICON           =   "frmvendedores.frx":3490
         PICN            =   "frmvendedores.frx":34AC
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin KewlButtonz.KewlButtons KewlButtons1 
         Height          =   495
         Left            =   2880
         TabIndex        =   14
         Top             =   360
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
         BTYPE           =   14
         TX              =   "&Baja"
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
         MICON           =   "frmvendedores.frx":3A46
         PICN            =   "frmvendedores.frx":3A62
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
   Begin VB.Frame Frame1 
      Height          =   1335
      Left            =   240
      TabIndex        =   18
      Top             =   120
      Width           =   6495
      Begin VB.CheckBox Check1 
         Caption         =   "Habilitado"
         Height          =   255
         Left            =   2520
         TabIndex        =   2
         Top             =   840
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         DataField       =   "password"
         DataSource      =   "datPrimaryRS"
         Enabled         =   0   'False
         Height          =   285
         Left            =   1920
         TabIndex        =   25
         Top             =   240
         Width           =   4335
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Vendedor"
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
         Left            =   120
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   240
         Width           =   1815
      End
      Begin VB.TextBox text2 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Top             =   840
         Width           =   1815
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Password:"
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
         Left            =   120
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   600
         Width           =   1815
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   3855
      Left            =   120
      ScaleHeight     =   3855
      ScaleWidth      =   6615
      TabIndex        =   19
      Top             =   1560
      Width           =   6615
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         BackColor       =   &H80000016&
         Height          =   285
         Index           =   8
         Left            =   4560
         TabIndex        =   11
         Top             =   1320
         Width           =   735
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         BackColor       =   &H80000016&
         Height          =   285
         Index           =   7
         Left            =   4560
         TabIndex        =   10
         Top             =   360
         Width           =   735
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         BackColor       =   &H80000016&
         Height          =   285
         Index           =   6
         Left            =   2400
         TabIndex        =   9
         Top             =   3240
         Width           =   735
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         BackColor       =   &H80000016&
         Height          =   285
         Index           =   5
         Left            =   2400
         TabIndex        =   8
         Top             =   2520
         Width           =   735
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         BackColor       =   &H80000016&
         Height          =   285
         Index           =   4
         Left            =   2400
         TabIndex        =   7
         Top             =   1800
         Width           =   735
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         BackColor       =   &H80000016&
         Height          =   285
         Index           =   3
         Left            =   2400
         TabIndex        =   6
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         BackColor       =   &H80000016&
         Height          =   285
         Index           =   2
         Left            =   2400
         TabIndex        =   5
         Top             =   1320
         Width           =   735
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         BackColor       =   &H80000016&
         Height          =   285
         Index           =   0
         Left            =   2400
         TabIndex        =   3
         Top             =   360
         Width           =   735
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         BackColor       =   &H80000016&
         Height          =   285
         Index           =   1
         Left            =   2400
         TabIndex        =   4
         Top             =   600
         Width           =   735
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Aurotiza usar Codigo Generico"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   5
         Left            =   3480
         TabIndex        =   35
         Top             =   960
         Width           =   3015
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Limite de Descuento %"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   4
         Left            =   3480
         TabIndex        =   34
         Top             =   0
         Width           =   3015
      End
      Begin VB.Label Label1 
         Caption         =   "Ingresar"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   33
         Top             =   3240
         Width           =   615
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Servicios"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   32
         Top             =   2880
         Width           =   3015
      End
      Begin VB.Label Label1 
         Caption         =   "Ingresar"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   31
         Top             =   2520
         Width           =   615
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Alquileres"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   30
         Top             =   2160
         Width           =   3015
      End
      Begin VB.Label Label1 
         Caption         =   "Vta Cdo"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   29
         Top             =   1800
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Vta Cta.Cte"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   28
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Ingresar"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   27
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Venta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   26
         Top             =   960
         Width           =   3015
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Presupuestos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   21
         Top             =   0
         Width           =   3015
      End
      Begin VB.Label Label1 
         Caption         =   "Ingresar"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   22
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Limita Fecha de vencimiento:"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   20
         Top             =   600
         Width           =   2295
      End
   End
   Begin MSAdodcLib.Adodc datvendedores 
      Height          =   330
      Left            =   2400
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
Attribute VB_Name = "frmvendedores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub borrar_Click()
On Error GoTo errorborrado


 KeyAscii = 13
  Respuesta = MsgBox("ESTA POR BORRAR UN USUARIO, ESTA SEGURO?", vbYesNo, "Atención")
If Respuesta = vbYes Then
    datprimaryrs.Recordset.Delete
Else
    Exit Sub
End If
Exit Sub

errorborrado:
Respuesta = MsgBox("No se pudo borrar el registro por contener permisos, limpie los permisos e intente de nuevo", vbCritical, "Atención")

     
End Sub

Private Sub eliminarempresa_Click()

KeyAscii = 13
  Respuesta = MsgBox("ESTA POR BORRAR UN PERMISO A EMPRESA?", vbYesNo, "Atención")
If Respuesta = vbYes Then
    datprimaryrs.Recordset.Delete
Else
    Exit Sub
End If
   

End Sub



Private Sub DataGrid2_ButtonClick(ByVal ColIndex As Integer)

    DataList1.Visible = True
    DataList1.SetFocus

End Sub

Private Sub DataList1_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        If DataGrid2.ApproxCount = 0 Then datusyemp.Recordset.AddNew
        DataGrid2.Columns(1).Text = DataList1.BoundText
        DataGrid2.Columns(2).Text = DataList1.Text
        DataGrid2.Columns(0).Text = Text1.Text
        datusyemp.Recordset.UpdateBatch adAffectCurrent
        DataGrid2.SetFocus
    End If

End Sub

Private Sub DataList1_LostFocus()

    DataList1.Visible = False

End Sub

Private Sub agregaproducto_Click()



End Sub

Private Sub alta_Click()
On Error GoTo errorgrabar
    

          datusuarios.Recordset.AddNew

          datusuarios.Recordset.Fields("codigo") = DataGrid1.Columns(1).Text
          datusuarios.Recordset.Fields("id") = DataGrid1.Columns(0).Text
          datusuarios.Recordset.Fields("apynomb") = Text1.Text
          datusuarios.Recordset.Fields("clave") = Text2.Text
          datusuarios.Recordset.Fields("habilitado") = Check1.Value
          datusuarios.Recordset.Fields("presupingresar") = Text3(0).Text
          datusuarios.Recordset.Fields("presuplimitefecha") = Text3(1).Text
          datusuarios.Recordset.Fields("ventaingresar") = Text3(2).Text
          datusuarios.Recordset.Fields("ventactacte") = Text3(3).Text
          datusuarios.Recordset.Fields("ventacdo") = Text3(4).Text
          datusuarios.Recordset.Fields("alquileringresar") = Text3(5).Text
          datusuarios.Recordset.Fields("serviciosingresar") = Text3(6).Text
          datusuarios.Recordset.Fields("limite") = Text3(7).Text
          datusuarios.Recordset.Fields("autorizagenerico") = Text3(8).Text
          
          datusuarios.Recordset.UpdateBatch adAffectCurrent


    mensa = MsgBox("Grabado Correctamente", vbInformation, "Grabado Correctamente")
    
    
Exit Sub
errorgrabar:
    mensa = MsgBox("No se pudo registrar la información", vbCritical, "Error !!")


End Sub

Private Sub Cancelar_Click()

    Unload Me
    frmvendedores.Show

End Sub

Private Sub DataGrid1_Click()

Call llena_Click

End Sub

Private Sub DataGrid1_KeyDown(KeyCode As Integer, Shift As Integer)

Call llena_Click

End Sub

Private Sub DataGrid1_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        Call llena_Click
        Text2.SetFocus
    End If

End Sub

Private Sub DataGrid1_KeyUp(KeyCode As Integer, Shift As Integer)

Call llena_Click


End Sub

Private Sub Form_Load()
Aplicar_skin Me

yventana = Inicio.Height / 2 - 1000
xventana = Inicio.Width / 2

frmvendedores.Top = yventana - frmvendedores.Height / 2
frmvendedores.Left = xventana - frmvendedores.Width / 2

datusuarios.ConnectionString = login.conexiontotal
datvendedores.ConnectionString = login.conexiontotal



datvendedores.RecordSource = "SELECT V_VENDEDOR_.ID, V_VENDEDOR_.CODIGO, V_PERSONA_.NOMBRE, V_VENDEDOR_.CODIGO  + ' ' + V_PERSONA_.NOMBRE as cancatena " & _
                             "FROM V_VENDEDOR_ INNER JOIN V_PERSONA_ ON V_VENDEDOR_.ENTEASOCIADO_ID = V_PERSONA_.ID " & _
                             "WHERE (V_VENDEDOR_.ACTIVESTATUS = 0)  order by V_PERSONA_.NOMBRE"
datvendedores.Refresh

DataGrid1.Columns(0).Visible = False
DataGrid1.Columns(3).Visible = False
DataGrid1.Columns(1).Width = 1000
DataGrid1.Columns(2).Width = 2500

End Sub

Private Sub grabar_Click()
On Error GoTo errorgrabar
    
    
          datusuarios.Recordset.Fields("codigo") = DataGrid1.Columns(1).Text
          datusuarios.Recordset.Fields("id") = DataGrid1.Columns(0).Text
          datusuarios.Recordset.Fields("apynomb") = Text1.Text
          datusuarios.Recordset.Fields("clave") = Text2.Text
          datusuarios.Recordset.Fields("habilitado") = Check1.Value
          datusuarios.Recordset.Fields("presupingresar") = Text3(0).Text
          datusuarios.Recordset.Fields("presuplimitefecha") = Text3(1).Text
          datusuarios.Recordset.Fields("ventaingresar") = Text3(2).Text
          datusuarios.Recordset.Fields("ventactacte") = Text3(3).Text
          datusuarios.Recordset.Fields("ventacdo") = Text3(4).Text
          datusuarios.Recordset.Fields("alquileringresar") = Text3(5).Text
          datusuarios.Recordset.Fields("serviciosingresar") = Text3(6).Text
          datusuarios.Recordset.Fields("limite") = Text3(7).Text
          datusuarios.Recordset.Fields("autorizagenerico") = Text3(8).Text
          datusuarios.Recordset.UpdateBatch adAffectCurrent
          mensa = MsgBox("Grabado Correctamente", vbInformation, "Grabado Correctamente")
    
    
Exit Sub
errorgrabar:
    mensa = MsgBox("No se pudo registrar la información", vbCritical, "Error !!")
    
    

End Sub

Private Sub KewlButtons1_Click()
On Error Resume Next
        mensa = MsgBox("Desea Eliminar este Vendedor del Punto de Venta ?", vbYesNo, "!! Atención !!")
        If mensa = vbYes Then
            datusuarios.Recordset.Delete adAffectCurrent
            mensa = MsgBox("Vendedor Dado de Baja", vbInformation, "Eliminacion de Registro")
            Call Cancelar_Click
        End If



End Sub

Private Sub llena_Click()
On Error Resume Next
    Text1.Text = ""
    Text2.Text = ""
    Check1.Value = 0
    For X = 0 To 8
        Text3(X).Text = ""
    Next X

    Text1.Text = DataGrid1.Columns(2).Text
    datusuarios.RecordSource = "select * from ud_ezi_empleado where id = '" & DataGrid1.Columns(0).Text & "'"
    datusuarios.Refresh
    
    If datusuarios.Recordset.EOF = True Then
        Respuesta = MsgBox("No esta dado de alta este vendedor en el Punto de venta, desea Hacelo ?", vbYesNo, "Atención")
        If Respuesta = vbYes Then
          Text2.SetFocus
          alta.Visible = True
          grabar.Visible = False
        End If
    Else
        Text2.Text = datusuarios.Recordset.Fields("clave")
        Check1.Value = datusuarios.Recordset.Fields("habilitado")
        Text3(0).Text = datusuarios.Recordset.Fields("presupingresar")
        Text3(1).Text = datusuarios.Recordset.Fields("presuplimitefecha")
        Text3(2).Text = datusuarios.Recordset.Fields("ventaingresar")
        Text3(3).Text = datusuarios.Recordset.Fields("ventactacte")
        Text3(4).Text = datusuarios.Recordset.Fields("ventacdo")
        Text3(5).Text = datusuarios.Recordset.Fields("alquileringresar")
        Text3(6).Text = datusuarios.Recordset.Fields("serviciosingresar")
        Text3(7).Text = datusuarios.Recordset.Fields("limite")
        Text3(8).Text = datusuarios.Recordset.Fields("autorizagenerico")
        alta.Visible = False
        grabar.Visible = True
        
    End If
        

End Sub

Private Sub salir_Click()
  Unload Me
End Sub


Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Text2.SetFocus
    End If
End Sub

Private Sub text2_Change()

 Text2.PasswordChar = "*"
 
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Check1.SetFocus
    End If
End Sub

Private Sub Text3_GotFocus(Index As Integer)

    Text3(Index).Text = UCase(Text3(Index).Text)
    Text3(Index).SelStart = 0
    Text3(Index).SelLength = 1

End Sub

Private Sub Text3_KeyPress(Index As Integer, KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = 13 Then
       KeyAscii = 0
       If Index <> 7 Then
        If Text3(Index).Text <> "s" And Text3(Index).Text <> "S" Then
            Text3(Index).Text = "N"
        Else
            Text3(Index).Text = "S"
        End If
       End If
        SendKeys "{TAB}", False
    End If
        
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
        KeyAscii = 0
        If Text6.Text <> "s" And Text6.Text <> "S" Then
            Text6.Text = "N"
        Else
            Text6.Text = "S"
        End If

        For X = 0 To 49
            If Text6.Text = "S" Then
                Text3(X).Text = "S"
            Else
                Text3(X).Text = "N"
            End If
        Next X
       
        Text3(0).SetFocus
    End If
End Sub
