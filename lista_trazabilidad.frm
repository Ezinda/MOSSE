VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{D18BBD1F-82BB-4385-BED3-E9D31A3E361E}#1.0#0"; "KewlButtonz.ocx"
Begin VB.Form lista_trazabilidad 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Trazabilidad de Documentos"
   ClientHeight    =   7050
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7470
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7050
   ScaleWidth      =   7470
   Begin VB.CommandButton dibuja 
      Caption         =   "dibuja"
      Height          =   375
      Left            =   4800
      TabIndex        =   12
      Top             =   6360
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.PictureBox Picture1 
      Height          =   6135
      Left            =   120
      ScaleHeight     =   6075
      ScaleWidth      =   7035
      TabIndex        =   1
      Top             =   120
      Width           =   7095
      Begin KewlButtonz.KewlButtons presupuesto 
         Height          =   855
         Left            =   1920
         TabIndex        =   2
         Top             =   120
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   1508
         BTYPE           =   14
         TX              =   "Presupuesto"
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
         MICON           =   "lista_trazabilidad.frx":0000
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin KewlButtonz.KewlButtons notadeventa 
         Height          =   855
         Left            =   720
         TabIndex        =   3
         Top             =   1680
         Visible         =   0   'False
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   1508
         BTYPE           =   14
         TX              =   "Nota de Venta"
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
         MICON           =   "lista_trazabilidad.frx":001C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin KewlButtonz.KewlButtons remito 
         Height          =   855
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   3240
         Visible         =   0   'False
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   1508
         BTYPE           =   14
         TX              =   "Remito"
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
         MICON           =   "lista_trazabilidad.frx":0038
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin KewlButtonz.KewlButtons remito 
         Height          =   855
         Index           =   1
         Left            =   1800
         TabIndex        =   5
         Top             =   3240
         Visible         =   0   'False
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   1508
         BTYPE           =   14
         TX              =   "Remito"
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
         MICON           =   "lista_trazabilidad.frx":0054
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin KewlButtonz.KewlButtons remito 
         Height          =   855
         Index           =   2
         Left            =   3480
         TabIndex        =   6
         Top             =   3240
         Visible         =   0   'False
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   1508
         BTYPE           =   14
         TX              =   "Remito"
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
         MICON           =   "lista_trazabilidad.frx":0070
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin KewlButtonz.KewlButtons remito 
         Height          =   855
         Index           =   3
         Left            =   5160
         TabIndex        =   7
         Top             =   3240
         Visible         =   0   'False
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   1508
         BTYPE           =   14
         TX              =   "Remito"
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
         MICON           =   "lista_trazabilidad.frx":008C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin KewlButtonz.KewlButtons factura 
         Height          =   855
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   4800
         Visible         =   0   'False
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   1508
         BTYPE           =   14
         TX              =   "Factura"
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
         MICON           =   "lista_trazabilidad.frx":00A8
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin KewlButtonz.KewlButtons factura 
         Height          =   855
         Index           =   1
         Left            =   1800
         TabIndex        =   9
         Top             =   4800
         Visible         =   0   'False
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   1508
         BTYPE           =   14
         TX              =   "Factura"
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
         MICON           =   "lista_trazabilidad.frx":00C4
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin KewlButtonz.KewlButtons factura 
         Height          =   855
         Index           =   2
         Left            =   3480
         TabIndex        =   10
         Top             =   4800
         Visible         =   0   'False
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   1508
         BTYPE           =   14
         TX              =   "Factura"
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
         MICON           =   "lista_trazabilidad.frx":00E0
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin KewlButtonz.KewlButtons factura 
         Height          =   855
         Index           =   3
         Left            =   5160
         TabIndex        =   11
         Top             =   4800
         Visible         =   0   'False
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   1508
         BTYPE           =   14
         TX              =   "Factura"
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
         MICON           =   "lista_trazabilidad.frx":00FC
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Image Image1 
         Height          =   720
         Index           =   8
         Left            =   5640
         Picture         =   "lista_trazabilidad.frx":0118
         Top             =   4080
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.Image Image1 
         Height          =   720
         Index           =   7
         Left            =   3960
         Picture         =   "lista_trazabilidad.frx":696A
         Top             =   4080
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.Image Image1 
         Height          =   720
         Index           =   6
         Left            =   2280
         Picture         =   "lista_trazabilidad.frx":D1BC
         Top             =   4080
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.Image Image1 
         Height          =   720
         Index           =   5
         Left            =   600
         Picture         =   "lista_trazabilidad.frx":13A0E
         Top             =   4080
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.Image Image1 
         Height          =   720
         Index           =   4
         Left            =   5640
         Picture         =   "lista_trazabilidad.frx":1A260
         Top             =   2520
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.Image Image1 
         Height          =   720
         Index           =   0
         Left            =   3120
         Picture         =   "lista_trazabilidad.frx":20AB2
         Top             =   960
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.Image Image1 
         Height          =   720
         Index           =   1
         Left            =   600
         Picture         =   "lista_trazabilidad.frx":27304
         Top             =   2520
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.Image Image1 
         Height          =   720
         Index           =   2
         Left            =   2280
         Picture         =   "lista_trazabilidad.frx":2DB56
         Top             =   2520
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.Image Image1 
         Height          =   720
         Index           =   3
         Left            =   3960
         Picture         =   "lista_trazabilidad.frx":343A8
         Top             =   2520
         Visible         =   0   'False
         Width           =   720
      End
   End
   Begin VB.CommandButton salir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   2760
      TabIndex        =   0
      Top             =   6480
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc dattraza 
      Height          =   330
      Left            =   240
      Top             =   6480
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
      LockType        =   1
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
Attribute VB_Name = "lista_trazabilidad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public cuentacont As String
Public controlst As Integer
Public controlsalto As Integer
Public xcantidadreal As Integer
Public xtrazabilidad As String
Dim cuenta(99999) As Integer


Private Sub datvendedor_WillMove(ByVal adReason As ADODB.EventReasonEnum, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)

End Sub

Private Sub dibuja_Click()


dattraza.RecordSource = "select * from v_ezi_pos_trazabilidad where NroCotizacion = '" & lista_presupuestos_todos.DataGrid1.Columns(1) & "' order by NroRemito"
dattraza.Refresh

dattraza.Recordset.MoveFirst
presupuesto.Caption = "Cotizacion " + Str(dattraza.Recordset.Fields("NroCotizacion"))
notadeventa.Caption = "Nota de Venta " + Str(dattraza.Recordset.Fields("NroNotaVenta"))
If dattraza.Recordset.Fields("NroNotaVenta") <> 0 Then
    notadeventa.Visible = True
    Image1(0).Visible = True
End If
xcontrol = 0
x = 0
Do While Not dattraza.Recordset.EOF
    If dattraza.Recordset.Fields("NroRemito") <> "" Then
        remito(x).Caption = "Remito " + dattraza.Recordset.Fields("NroRemito") + " Pend.Fac " + Str(dattraza.Recordset.Fields("PendFacturar"))
        remito(x).Visible = True
        Image1(x + 1).Visible = True
        If (dattraza.Recordset.Fields("NroFactura") <> "" Or dattraza.Recordset.Fields("NroFactura") <> "NULL") And xcontrol = 0 Then
                factura(x).Caption = "Factura " + dattraza.Recordset.Fields("NroFactura")
                factura(x).Visible = True
                Image1(x + 5).Visible = True
        End If
        
        x = x + 1
    Else
        If dattraza.Recordset.Fields("NroFactura") <> "" Then
                remito(x).Caption = "Factura " + dattraza.Recordset.Fields("NroFactura")
                remito(x).Visible = True
                Image1(x + 1).Visible = True
                x = x + 1
                xcontrol = 1
        End If
    
    
    End If
    
    
    dattraza.Recordset.MoveNext
Loop
    
For x = 0 To 3
    If factura(x).Caption = factura(x + 1).Caption Then
        factura(x + 1).Visible = False
        factura(x).Width = factura(x).Width + factura(x + 1).Width
    End If
Next x

    
    



End Sub

Private Sub Form_Load()
On Error Resume Next
MiFuncionDeAjuste Me, True
Aplicar_skin Me

lista_trazabilidad.Top = 0
lista_trazabilidad.Left = 0

dattraza.ConnectionString = login.conexiontotal

Call dibuja_Click
 
End Sub



Private Sub remito_Click(Index As Integer)

    If Left(remito(Index), 6) = "Remito" Then
        menu = 10
        xtrazabilidad = Mid(remito(Index), 8, 13)
        frmremitosconsulta.Show
        frmremitosconsulta.Text1.Text = Mid(remito(Index), 8, 13)
    End If
    

End Sub

Private Sub salir_Click()

Unload Me

End Sub

