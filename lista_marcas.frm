VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form lista_marcas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Maestro de Marcas"
   ClientHeight    =   6720
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5100
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6720
   ScaleWidth      =   5100
   Begin VB.CommandButton salir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   3720
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "lista_marcas.frx":0000
      Height          =   5775
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   10186
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      BackColor       =   16777215
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
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   540
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   5100
      _ExtentX        =   8996
      _ExtentY        =   953
      ButtonWidth     =   609
      ButtonHeight    =   900
      Appearance      =   1
      Style           =   1
      _Version        =   393216
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
         TabIndex        =   1
         Top             =   120
         Width           =   2655
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Buscar:"
         Height          =   375
         Left            =   120
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   120
         Width           =   855
      End
   End
   Begin MSAdodcLib.Adodc datcliente 
      Height          =   330
      Left            =   120
      Top             =   6360
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
End
Attribute VB_Name = "lista_marcas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public cuentacont As String
Dim cuenta(99999) As Integer



Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
On Error Resume Next

    If KeyAscii = 13 Then
        KeyAscii = 0
        If frmcomparativa.Visible = True Then
            frmcomparativa.grilla.TextMatrix(frmcomparativa.grilla.Row, frmcomparativa.grilla.Col) = DataGrid1.Columns(0).Text
        End If
        Unload Me
    End If
        


End Sub

Private Sub Form_Load()
Aplicar_skin Me

MiFuncionDeAjuste Me, True

yventana = Inicio.Height / 2 - 1000
xventana = Inicio.Width / 2

lista_marcas.Top = yventana - lista_marcas.Height / 2
lista_marcas.Left = xventana - lista_marcas.Width / 2


Dim pruebita As Currency
Dim ruta As String
Dim ret As Long

crlf = Chr(13) & Chr(10)
campo7 = ""
Open App.Path & "\sucursal.ini" For Input As #1
While Not EOF(1)
Line Input #1, file_data$
campo7 = campo7 & file_data$ & crlf
Wend
Close #1

nomsucursal = ""
nombredsn = ""
nombrebd = ""
For X = 1 To Len(campo7)
    If Mid(campo7, X, 1) = ";" Then
      For Y = X + 1 To Len(campo7)
        If Mid(campo7, Y, 1) = ";" Then
            For Z = Y + 1 To Len(campo7)
                nombrebd = nombrebd + Mid(campo7, Z, 1)
            Next Z
            GoTo paso0
        End If
        nombredsn = nombredsn + Mid(campo7, Y, 1)
      Next Y
    End If
    nomsucursal = nomsucursal + Mid(campo7, X, 1)
Next X

paso0:

cuerpo1 = "Provider=MSDASQL.1;Password=1;Persist Security Info=True;User ID=fs1;"
cuerpo1b = "Data Source="
cuerpo2 = nombredsn
cuerpo3 = ";Initial Catalog="
cuerpo3b = nombrebd
basededatos = nombrebd
conexiontotal = cuerpo1 + cuerpo1b + cuerpo2 + cuerpo3 + cuerpo3b


datcliente.ConnectionString = conexiontotal


     xquery = "SELECT CODIGO FROM V_ITEMTIPOCLASIFICADOR_ AS ALIAS_0 WHERE     (BO_PLACE_ID = '{3D2BA6F5-9C85-45DA-8C30-5738730323CC}') AND (ACTIVESTATUS <> 2) ORDER BY CODIGO"

datcliente.RecordSource = xquery
datcliente.Refresh

            DataGrid1.Columns(0).Width = 3500


 
End Sub

Private Sub salir_Click()
    
    Unload Me
    End
    

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)

On Error Resume Next


    If KeyAscii = 13 Then
        KeyAscii = 0
        If Text1.Text <> "" Then
            xbusqueda = "%" + Text1.Text + "%"
            xquery1 = "SELECT CODIGO FROM V_ITEMTIPOCLASIFICADOR_ AS ALIAS_0 WHERE     (BO_PLACE_ID = '{3D2BA6F5-9C85-45DA-8C30-5738730323CC}') AND (ACTIVESTATUS <> 2) " & _
                      "and  CODIGO like '" & xbusqueda & "' ORDER BY CODIGO"
                      
            datcliente.RecordSource = xquery1
            datcliente.Refresh

            DataGrid1.Columns(0).Width = 3500


        End If
        DataGrid1.SetFocus
        
        
    End If

End Sub
