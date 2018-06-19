VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form lista_historial 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Hisotrial de Ventas"
   ClientHeight    =   6990
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12975
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6990
   ScaleWidth      =   12975
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Producto:"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   0
      Width           =   855
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
      Left            =   1080
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   0
      Width           =   7575
   End
   Begin VB.CommandButton salir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   11040
      TabIndex        =   1
      Top             =   0
      Width           =   1215
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "lista_historial.frx":0000
      Height          =   5775
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   12735
      _ExtentX        =   22463
      _ExtentY        =   10186
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      WrapCellPointer =   -1  'True
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
   Begin MSAdodcLib.Adodc datvendedor 
      Height          =   330
      Left            =   120
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
Attribute VB_Name = "lista_historial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public cuentacont As String
Dim cuenta(99999) As Integer
Private Const CAMPO_A_FILTRAR As String = "NOMBRE_PRODUCTO"


Private Sub Form_Load()
On Error Resume Next
If menu = 2 Then
'    Aplicar_skin2 Me
    Aplicar_skin Me
Else
    Aplicar_skin Me
End If

MiFuncionDeAjuste Me, True

datvendedor.ConnectionString = login.conexiontotal

datvendedor.RecordSource = query
datvendedor.Refresh

If datvendedor.Recordset.EOF = True Then
    MsgBox "Sin historial de Venta", vbInformation, ""
    Unload Me
    Exit Sub
End If


'Text1.Text = datvendedor.Recordset.Fields("CLIENTE")
lista_historial.Caption = "Historial de Ventas, Cliente: " + datvendedor.Recordset.Fields("CLIENTE")


DataGrid1.Columns(0).Visible = False
DataGrid1.Columns(1).Visible = False
DataGrid1.Columns(2).Visible = False
DataGrid1.Columns(6).Visible = False
DataGrid1.Columns(7).Visible = False
DataGrid1.Columns(8).Visible = False

DataGrid1.Columns(3).Caption = "Fecha"
DataGrid1.Columns(4).Caption = "Producto"
DataGrid1.Columns(5).Caption = "Descripcion"

DataGrid1.Columns(3).Width = 1500
DataGrid1.Columns(3).Alignment = dbgCenter
DataGrid1.Columns(3).Width = 2000
DataGrid1.Columns(4).Alignment = dbgLeft
DataGrid1.Columns(5).Width = 6500
DataGrid1.Columns(5).Alignment = dbgLeft

DataGrid1.Columns(9).NumberFormat = "Currency"
DataGrid1.Columns(9).Caption = "Precio"
DataGrid1.Columns(9).Alignment = dbgRight




 
End Sub

Private Sub salir_Click()

Unload Me

End Sub

Private Sub Text1_Change()
On Error Resume Next


    datvendedor.Recordset.Filter = CAMPO_A_FILTRAR & " LIKE '*" + Text1.Text + "*'"
    

End Sub
