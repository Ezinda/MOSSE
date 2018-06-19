VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{50BF2256-701F-46F2-8ADB-2202CE6922BC}#1.0#0"; "KlexGrid.ocx"
Begin VB.Form lista_lotes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Stock de Producto"
   ClientHeight    =   5325
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7635
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5325
   ScaleWidth      =   7635
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
      Height          =   1005
      Left            =   1080
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   0
      Width           =   6375
   End
   Begin VB.CommandButton salir 
      Cancel          =   -1  'True
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   2880
      TabIndex        =   1
      Top             =   4800
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc datvendedor 
      Height          =   330
      Left            =   240
      Top             =   4680
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
   Begin Grid.KlexGrid klote 
      Height          =   3375
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   5953
      EnterKeyBehaviour=   2
      BackColorAlternate=   0
      GridLinesFixed  =   2
      BackColorFixed  =   -2147483626
      Cols            =   5
      FixedCols       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GridColorFixed  =   8421504
      MouseIcon       =   "lista_lotes.frx":0000
      Rows            =   10
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "lista_lotes.frx":001C
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Visible         =   0   'False
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   873
      _Version        =   393216
      AllowUpdate     =   -1  'True
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
End
Attribute VB_Name = "lista_lotes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public cuentacont As String
Public controlst As Integer
Public controlsalto As Integer
Public xcantidadreal As Double
Dim cuenta(99999) As Integer
Dim xsalida As Integer


Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
On Error Resume Next

    If KeyAscii = 13 Then
        KeyAscii = 0
        If menu = 1 Then
          If DataGrid1.Columns(7).Text = "" Then
            frmnota_venta.grilla.TextMatrix(xfila, 18) = "SL"  ' Sin Lote
          Else
            frmnota_venta.grilla.TextMatrix(xfila, 18) = DataGrid1.Columns(7).Text  ' Id de Lote
          End If
            For X = 1 To Len(frmnota_venta.grilla.TextMatrix(xfila, 2))
                xcar = Mid(frmnota_venta.grilla.TextMatrix(xfila, 2), X, 4)
                If xcar = "Lote" Then
                  frmnota_venta.grilla.TextMatrix(xfila, 2) = Left(frmnota_venta.grilla.TextMatrix(xfila, 2), X - 1)
                  Exit For
                End If
            Next
            frmnota_venta.grilla.TextMatrix(xfila, 2) = frmnota_venta.grilla.TextMatrix(xfila, 2) + " Lote:" + Str(DataGrid1.Columns(2).Text)
            Unload Me
        End If
        
        If menu = 2 Then
          If DataGrid1.Columns(7).Text = "" Then
            lista_notadeventas.KlexGrid1.TextMatrix(xfila, 9) = "SL"  ' Sin Lote
          Else
            lista_notadeventas.KlexGrid1.TextMatrix(xfila, 11) = DataGrid1.Columns(7).Text  ' Id de Lote
            lista_notadeventas.KlexGrid1.TextMatrix(xfila, 9) = Str(DataGrid1.Columns(2).Text)  ' Lote
          End If
          Unload Me
        End If
        
    End If
        

End Sub

Private Sub Form_Activate()
On Error Resume Next
'    If menu = 2 Then
'        Text1.SetFocus
        DoEvents
        klote.SetFocus
       salir.SetFocus
'    End If
    
End Sub

Private Sub Form_Load()
On Error Resume Next
MiFuncionDeAjuste Me, True
Aplicar_skin Me

datvendedor.ConnectionString = login.conexiontotal

datvendedor.RecordSource = query
datvendedor.Refresh

xsalida = 0

If datvendedor.Recordset.EOF = True And menu = 2 Then
  If controlst = 0 And lista_notadeventas.Check1 = 0 Then
    MsgBox "Sin Stock", vbInformation, ""
    controlst = 1
    If lista_notadeventas.Visible = True Then
        lista_notadeventas.KlexGrid1.TextMatrix(lista_notadeventas.KlexGrid1.Row, 7) = 0
        xsalida = 1
    End If
    Unload Me
    Exit Sub
  Else
    controlst = 0
    Unload Me
    Exit Sub
  End If
End If

If datvendedor.Recordset.EOF = True And menu = 1 Then
    MsgBox "Sin Stock", vbInformation, ""
    controlst = 0
    Unload Me
    Exit Sub
End If


Text1.Text = datvendedor.Recordset.Fields("NOMBREREFERENCIA") + " - " + datvendedor.Recordset.Fields("descripcion")
lista_lotes.Caption = "Stock Producto"

    klote.Rows = datvendedor.Recordset.RecordCount + 1
    klote.Cols = 6
    klote.ColWidth(0) = 1500
    klote.TextMatrix(0, 0) = "Lote Nro."
    klote.ColWidth(1) = 1500
    klote.TextMatrix(0, 1) = "Stock"
    klote.ColWidth(2) = 800
    klote.TextMatrix(0, 2) = "Remitir"
    klote.ColWidth(3) = 1500
    klote.TextMatrix(0, 3) = "Fec.Venc."
    klote.ColWidth(4) = 2000
    klote.TextMatrix(0, 4) = "Proveedor"
    
    lin = 1
If datvendedor.Recordset.EOF = False Then
    datvendedor.Recordset.MoveFirst
    Do While Not datvendedor.Recordset.EOF
        If IsNull(datvendedor.Recordset.Fields(2)) = False Then
            klote.TextMatrix(lin, 0) = CStr(datvendedor.Recordset.Fields(2)) + "."
        End If
        If IsNull(datvendedor.Recordset.Fields(3)) = False Then
            klote.TextMatrix(lin, 1) = Replace(Round(datvendedor.Recordset.Fields(3), 2), ".", "")
        End If
        If IsNull(datvendedor.Recordset.Fields(9)) = False Then
            klote.TextMatrix(lin, 2) = datvendedor.Recordset.Fields(9)
        End If
        If IsNull(datvendedor.Recordset.Fields(10)) = False Then
            klote.TextMatrix(lin, 3) = datvendedor.Recordset.Fields(10)
        End If
        If IsNull(datvendedor.Recordset.Fields(11)) = False Then
            klote.TextMatrix(lin, 4) = datvendedor.Recordset.Fields(11)
        End If
        If IsNull(datvendedor.Recordset.Fields(7)) = False Then
            klote.TextMatrix(lin, 5) = datvendedor.Recordset.Fields(7)
        End If
        datvendedor.Recordset.MoveNext
        lin = lin + 1
    Loop
End If

If menu = 1 Then
        xcantidad = Val(frmnota_venta.Text2.Text)
        xcantidadreal = Val(frmnota_venta.Text2.Text)
        For X = 1 To klote.Rows - 1
            If DateValue(klote.TextMatrix(X, 3)) > Date Then
                      
               xdif = xcantidad - Val(klote.TextMatrix(X, 1))
               If xdif <= 0 Then
                    klote.TextMatrix(X, 2) = xcantidad
                    Exit For
               Else
                    klote.TextMatrix(X, 2) = Val(klote.TextMatrix(X, 1))
                    xcantidad = xdif
               End If
            End If
        Next X
        controlsalto = 0
End If



If menu = 2 Then
          xcantidad = Val(lista_notadeventas.KlexGrid1.TextMatrix(xfila, 7))
        xcantidadreal = Val(lista_notadeventas.KlexGrid1.TextMatrix(xfila, 7)) + Val(lista_notadeventas.KlexGrid1.TextMatrix(xfila, 8))
        For X = 1 To klote.Rows - 1
            If DateValue(klote.TextMatrix(X, 3)) > Date Then
               xdif = xcantidad - Val(klote.TextMatrix(X, 1))
               If xdif <= 0 Then
                    klote.TextMatrix(X, 2) = xcantidad
                    Exit For
               Else
                    klote.TextMatrix(X, 2) = Val(klote.TextMatrix(X, 1))
                    xcantidad = xdif
               End If
            End If
        Next X
        controlsalto = 0
End If

 
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next


    If menu = 1 Then
          xtotal = 0
          For X = 1 To klote.Rows - 1
               xtotal = xtotal + Val(klote.TextMatrix(X, 2))
          Next X
        If xtotal > xcantidadreal Then
                MsgBox "La cantidad seleccionada supera a la cantidad del pedido", vbInformation, "Verificar !!"
                Exit Sub
        Else
          frmnota_venta.Text2 = xtotal
          frmnota_venta.grilla.TextMatrix(frmnota_venta.grilla.Row, 3) = xtotal
          xlin = 1
          xcolum = 18
          For j = 1 To klote.Rows - 1
                    lotecodigo(xfila, xlin) = klote.TextMatrix(j, 0)
                    lotecantidad(xfila, xlin) = klote.TextMatrix(j, 2)
                    loteid(xfila, xlin) = klote.TextMatrix(j, 5)
                    
                    If klote.TextMatrix(j, 2) <> 0 Then
                        frmnota_venta.grilla.TextMatrix(xfila, xcolum) = klote.TextMatrix(j, 5)
                        frmnota_venta.grilla.TextMatrix(xfila, xcolum + 1) = klote.TextMatrix(j, 2)
                        xcolum = xcolum + 2
                    End If
                    
                    xlin = xlin + 1
                    
          Next j
        End If
        Unload Me
    End If

        If menu = 2 And lista_notadeventas.Check1.Value = 0 Then
           xtotal = 0
           For X = 1 To klote.Rows - 1
               xtotal = xtotal + Val(klote.TextMatrix(X, 2))
           Next X
           If xtotal + Val(lista_notadeventas.KlexGrid1.TextMatrix(xfila, 8)) > xcantidadreal Then
                MsgBox "La cantidad seleccionada supera a la cantidad del pedido", vbInformation, "Verificar !!"
                Exit Sub
           Else
                lista_notadeventas.KlexGrid1.TextMatrix(xfila, 7) = xtotal
                xlin = 1
                For j = 1 To klote.Rows - 1
                    lotecodigo(xfila, xlin) = klote.TextMatrix(j, 0)
                    lotecantidad(xfila, xlin) = klote.TextMatrix(j, 2)
                    loteid(xfila, xlin) = klote.TextMatrix(j, 5)
                    xlin = xlin + 1
                Next j
           End If
          Unload Me
        End If

End Sub

Private Sub klote_AfterEdit(ByVal Row As Long, ByVal Col As Long)

On Error Resume Next

'    If menu = 2 Then
           xtotal = 0
           For X = 1 To klote.Rows - 1
               xtotal = xtotal + Val(klote.TextMatrix(X, 2))
           Next X
           If xtotal > xcantidadreal Then
             If controlsalto = 0 Then
                MsgBox "La cantidad seleccionada supera a la cantidad del pedido", vbInformation, "Verificar !!"
             End If
                controlsalto = 1
           Else
                controlsalto = 0
           End If
           
  
           
 '   End If

End Sub

Private Sub klote_EnterCell()

On Error Resume Next

If klote.Col = 0 Or klote.Col = 1 Or klote.Col = 3 Or klote.Col = 4 Then
    klote.Editable = False
    Exit Sub
Else
    If DateValue(klote.TextMatrix(klote.Row, 3)) < Date Then
        klote.Editable = False
    Else
        klote.Editable = True
    End If
End If


End Sub

Private Sub salir_Click()

'If menu = 2 Then
'    lista_notadeventas.KlexGrid1.SetFocus
'    SendKeys "{TAB}", False
'End If


Unload Me

End Sub

