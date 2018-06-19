VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmusuarios 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Usuarios"
   ClientHeight    =   5805
   ClientLeft      =   1095
   ClientTop       =   330
   ClientWidth     =   7440
   Icon            =   "frmusuarios.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5805
   ScaleWidth      =   7440
   Begin MSDataGridLib.DataGrid DataGrid2 
      Height          =   5535
      Left            =   5520
      TabIndex        =   17
      Top             =   120
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   9763
      _Version        =   393216
      BackColor       =   16777152
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
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   "nombre"
         Caption         =   "Usuario"
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
         DataField       =   ""
         Caption         =   ""
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
         EndProperty
      EndProperty
   End
   Begin MSDataListLib.DataList DataList1 
      Bindings        =   "frmusuarios.frx":0442
      Height          =   255
      Left            =   3480
      TabIndex        =   16
      Top             =   2640
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   450
      _Version        =   393216
      ListField       =   "razonsocial"
      BoundColumn     =   "empresa"
   End
   Begin MSAdodcLib.Adodc datempresas 
      Height          =   330
      Left            =   960
      Top             =   4080
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
      Connect         =   "PROVIDER=MSDASQL;dsn=contable;uid=sa;pwd=;database=contablesql;"
      OLEDBString     =   "PROVIDER=MSDASQL;dsn=contable;uid=sa;pwd=;database=contablesql;"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select empresa,razonsocial from empresa order by razonsocial"
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
   Begin VB.CommandButton eliminarempresa 
      Caption         =   "&Eliminar permiso de empresa"
      Height          =   375
      Left            =   2520
      TabIndex        =   15
      Top             =   4080
      Width           =   2535
   End
   Begin VB.CommandButton borrar 
      Caption         =   "&Borrar"
      Height          =   615
      Left            =   3120
      Picture         =   "frmusuarios.frx":045C
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4800
      UseMaskColor    =   -1  'True
      Width           =   855
   End
   Begin VB.CheckBox Check4 
      Caption         =   "Check4"
      Height          =   255
      Left            =   3000
      TabIndex        =   8
      Top             =   960
      Width           =   255
   End
   Begin VB.TextBox Text6 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      DataField       =   "administrador"
      DataSource      =   "datusuarios"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   285
      Left            =   3360
      Locked          =   -1  'True
      TabIndex        =   12
      Text            =   "Text6"
      Top             =   960
      Width           =   495
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      DataField       =   "password"
      DataSource      =   "datprimaryRS"
      Height          =   285
      Left            =   2760
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   600
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      DataField       =   "nombre"
      DataSource      =   "datprimaryRS"
      Height          =   285
      Left            =   2760
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   240
      Width           =   1455
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmusuarios.frx":055E
      Height          =   2655
      Left            =   120
      TabIndex        =   7
      Top             =   1440
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   4683
      _Version        =   393216
      AllowUpdate     =   -1  'True
      AllowArrows     =   -1  'True
      BackColor       =   16777152
      HeadLines       =   1
      RowHeight       =   15
      TabAction       =   2
      WrapCellPointer =   -1  'True
      FormatLocked    =   -1  'True
      AllowAddNew     =   -1  'True
      AllowDelete     =   -1  'True
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
      ColumnCount     =   5
      BeginProperty Column00 
         DataField       =   "nomusuario"
         Caption         =   "nomusuario"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "altas"
         Caption         =   "Altas"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "bajas"
         Caption         =   "Bajas"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "modificaciones"
         Caption         =   "Modific."
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "empresa"
         Caption         =   "Empresa"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            Alignment       =   2
            Locked          =   -1  'True
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column01 
            Alignment       =   2
         EndProperty
         BeginProperty Column02 
            Alignment       =   2
         EndProperty
         BeginProperty Column03 
            Alignment       =   2
         EndProperty
         BeginProperty Column04 
            Alignment       =   2
            Button          =   -1  'True
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton grabar 
      Caption         =   "&Grabar"
      Height          =   615
      Left            =   240
      Picture         =   "frmusuarios.frx":0578
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4800
      UseMaskColor    =   -1  'True
      Width           =   855
   End
   Begin VB.CommandButton nuevo 
      Caption         =   "&Nuevo"
      Height          =   615
      Left            =   1200
      Picture         =   "frmusuarios.frx":0AAA
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4800
      UseMaskColor    =   -1  'True
      Width           =   855
   End
   Begin VB.CommandButton Cancelar 
      Cancel          =   -1  'True
      Caption         =   "Cance&lar"
      Height          =   615
      Left            =   2160
      Picture         =   "frmusuarios.frx":0FDC
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4800
      UseMaskColor    =   -1  'True
      Width           =   855
   End
   Begin VB.CommandButton salir 
      Caption         =   "&Cerrar"
      Height          =   615
      Left            =   4200
      Picture         =   "frmusuarios.frx":150E
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4800
      UseMaskColor    =   -1  'True
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Height          =   1335
      Left            =   2640
      TabIndex        =   13
      Top             =   0
      Width           =   1935
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
      TabIndex        =   14
      Top             =   4560
      Width           =   5295
   End
   Begin MSAdodcLib.Adodc datprimaryRS 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      Top             =   5475
      Visible         =   0   'False
      Width           =   7440
      _ExtentX        =   13123
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
      Connect         =   "PROVIDER=MSDASQL;dsn=contable;uid=sa;pwd=;database=contablesql;"
      OLEDBString     =   "PROVIDER=MSDASQL;dsn=contable;uid=sa;pwd=;database=contablesql;"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   $"frmusuarios.frx":1950
      Caption         =   " "
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
   Begin VB.Label lblLabels 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Administrador:"
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
      Left            =   720
      TabIndex        =   11
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Password"
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
      Left            =   720
      TabIndex        =   10
      Top             =   615
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nombre Usuario:"
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
      Left            =   720
      TabIndex        =   9
      Top             =   300
      Width           =   1815
   End
End
Attribute VB_Name = "frmusuarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub agregarempresa_Click()

    datusyempresa.Recordset.AddNew
    Text3.Text = Text1.Text

End Sub

Private Sub borrar_Click()


 KeyAscii = 13
  Respuesta = MsgBox("ESTA POR BORRAR UN USUARIO, ESTA SEGURO?", vbYesNo, "Atención")
If Respuesta = vbYes Then
    datprimaryRS.Recordset.Delete
Else
    Exit Sub
End If

End Sub

Private Sub Cancelar_Click()

    datprimaryRS.Refresh

End Sub

Private Sub Check1_Click()
If Check1 = 1 Then
    Text3.Text = "S"
Else
    Text3.Text = "N"
End If

End Sub

Private Sub Check1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Check2.SetFocus
    End If
End Sub

Private Sub Check2_Click()
If Check2 = 1 Then
    Text4.Text = "S"
Else
    Text4.Text = "N"
End If
    
End Sub

Private Sub Check2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Check3.SetFocus
    End If
End Sub

Private Sub Check3_Click()
If Check3 = 1 Then
    Text5.Text = "S"
Else
    Text5.Text = "N"
End If
End Sub

Private Sub Check3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Check4.SetFocus
    End If
End Sub

Private Sub Check4_Click()
If Check4 = 1 Then
    Text6.Text = "S"
Else
    Text6.Text = "N"
End If
End Sub

Private Sub Check4_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        grabar.SetFocus
    End If
End Sub


Private Sub DataGrid1_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Or KeyAscii = 9 Then
      Text3.Text = Text1.Text
      POS = DataGrid1.Col
      If POS < 4 Then
        If DataGrid1.Columns(POS).Text = "s" Or DataGrid1.Columns(POS).Text = "S" Then
            DataGrid1.Columns(POS).Text = "S"
            KeyAscii = 9
            Exit Sub
        End If
        If DataGrid1.Columns(POS).Text = "n" Or DataGrid1.Columns(POS).Text = "N" Then
            DataGrid1.Columns(POS).Text = "N"
            KeyAscii = 9
            Exit Sub
        End If
        mensa = MsgBox("Dato incorrecto, debe ingresa S o N", vbCritical, "!! Error !!")
      Else
            KeyAscii = 9
      End If
    End If
    
    
    
        


End Sub

Private Sub datalist1_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
        Case vbKeyReturn
            DataGrid1.Columns(4).Text = DataList1.BoundText
            DataList1.Visible = False
            DataGrid1.SetFocus
        Case vbKeyEscape
            DataList1.Visible = False
            DataGrid1.SetFocus
    End Select
End Sub

Private Sub DataGrid1_ButtonClick(ByVal ColIndex As Integer)

 With DataList1
    .Left = DataGrid1.Left + DataGrid1.Columns(4).Left
    .Top = DataGrid1.Top + DataGrid1.RowTop(DataGrid1.Row) + DataGrid1.RowHeight
    .Width = DataGrid1.Columns(4).Width + 3000
    .Visible = True
    .ZOrder 0
    .SetFocus
 End With
  

End Sub




Private Sub DataGrid1_Scroll(Cancel As Integer)
    DataList1.Visible = False

End Sub

Private Sub eliminarempresa_Click()
     
     datusyempresa.Recordset.Delete
    
End Sub

Private Sub Form_Load()

Rem DataGrid1.DataSource = datprimaryRS.Recordset("ChildCMD").UnderlyingValue

If Text6.Text = "S" Then
    Check4 = 1
Else
    Check4 = 0
End If



End Sub

Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault

End Sub

Private Sub datPrimaryRS_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)
  'Aquí es donde puede colocar el código de control de errores
  'Si desea pasar por alto los errores, marque como comentario la siguiente línea
  'Si desea detectarlos, agregue código aquí para controlarlos
  MsgBox "Data error event hit err:" & Description
End Sub

Private Sub data1_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  'Esto mostrará la posición de registro actual para este Recordset
  datprimaryRS.Caption = "Record: " & CStr(datprimaryRS.Recordset.AbsolutePosition)
End Sub

Private Sub data1_WillChangeRecord(ByVal adReason As ADODB.EventReasonEnum, ByVal cRecords As Long, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  'Aquí se coloca el código de validación
  'Se llama a este evento cuando ocurre la siguiente acción
  Dim bCancel As Boolean

  Select Case adReason
  Case adRsnAddNew
  Case adRsnClose
  Case adRsnDelete
  Case adRsnFirstChange
  Case adRsnMove
  Case adRsnRequery
  Case adRsnResynch
  Case adRsnUndoAddNew
  Case adRsnUndoDelete
  Case adRsnUndoUpdate
  Case adRsnUpdate
  End Select

  If bCancel Then adStatus = adStatusCancel
End Sub

Private Sub cmdAdd_Click()
  On Error GoTo AddErr
  datprimaryRS.Recordset.AddNew

  Exit Sub
AddErr:
  MsgBox Err.Description
End Sub

Private Sub cmdDelete_Click()
  On Error GoTo DeleteErr
  With datprimaryRS.Recordset
    .Delete
    .MoveNext
    If .EOF Then .MoveLast
  End With
  Exit Sub
DeleteErr:
  MsgBox Err.Description
End Sub

Private Sub cmdRefresh_Click()
  'Esto sólo es necesario en aplicaciones multiusuario
  On Error GoTo RefreshErr
  datprimaryRS.Refresh
  Exit Sub
RefreshErr:
  MsgBox Err.Description
End Sub

Private Sub cmdUpdate_Click()
  On Error GoTo UpdateErr

  datprimaryRS.Recordset.UpdateBatch adAffectAll
  Exit Sub
UpdateErr:
  MsgBox Err.Description
End Sub

Private Sub cmdClose_Click()
  Unload Me
End Sub

Private Sub grabar_Click()

    datprimaryRS.Recordset.Save
    
End Sub

Private Sub nuevo_Click()
  datprimaryRS.Recordset.AddNew
  
End Sub


Private Sub salir_Click()
  Unload Me
End Sub

Private Sub text1_Change()

If Text6.Text = "S" Then
    Check4 = 1
Else
    Check4 = 0
End If

Text3.Text = Text1.Text

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        text2.SetFocus
    End If
End Sub

Private Sub text2_Change()
 text2.PasswordChar = "*"
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        DataGrid1.SetFocus
    End If
End Sub
