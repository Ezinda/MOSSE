VERSION 5.00
Object = "{F6125AB1-8AB1-11CE-A77F-08002B2F4E98}#2.0#0"; "MSRDC20.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{5734474E-78D3-4254-99B9-C35F31BDF509}#62.0#0"; "vbskpro2.ocx"
Object = "{D18BBD1F-82BB-4385-BED3-E9D31A3E361E}#1.0#0"; "KewlButtonz.ocx"
Begin VB.Form frmconsutalibroventas_n 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta Facturas de Venta"
   ClientHeight    =   6396
   ClientLeft      =   1092
   ClientTop       =   325
   ClientWidth     =   6604
   Icon            =   "frmconsultalibroventas_n.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6396
   ScaleWidth      =   6604
   Begin VB.CommandButton cerrado 
      Caption         =   "Libro Cerrado"
      Height          =   255
      Left            =   480
      TabIndex        =   19
      Top             =   2520
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Frame Frame2 
      Caption         =   "Opciones"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.47
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1575
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   6375
      Begin VB.CommandButton llena 
         Caption         =   "llena"
         Height          =   255
         Left            =   5520
         TabIndex        =   18
         Top             =   1080
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Tipo.Comp.:"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   1080
         Width           =   1095
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1200
         TabIndex        =   14
         Text            =   "Combo1"
         Top             =   1080
         Width           =   1095
      End
      Begin VB.CommandButton buscar 
         Caption         =   "&Buscar"
         Height          =   255
         Left            =   4200
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   1080
         UseMaskColor    =   -1  'True
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   1
         Left            =   1680
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Desde"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   240
         Width           =   855
      End
      Begin MSComCtl2.DTPicker desde 
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   1455
         _ExtentX        =   2564
         _ExtentY        =   671
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   7.4717
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarBackColor=   16777215
         Format          =   144965633
         CurrentDate     =   38410
      End
      Begin MSComCtl2.DTPicker hasta 
         Height          =   375
         Left            =   1680
         TabIndex        =   3
         Top             =   480
         Width           =   1455
         _ExtentX        =   2564
         _ExtentY        =   671
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   7.4717
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarBackColor=   16777215
         Format          =   144965633
         CurrentDate     =   38410
      End
      Begin KewlButtonz.KewlButtons salir 
         Height          =   495
         Left            =   5520
         TabIndex        =   9
         Top             =   360
         Width           =   735
         _ExtentX        =   1294
         _ExtentY        =   863
         BTYPE           =   14
         TX              =   "&Salir"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   7.4717
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
         MICON           =   "frmconsultalibroventas_n.frx":0442
         PICN            =   "frmconsultalibroventas_n.frx":045E
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin KewlButtonz.KewlButtons filtrar 
         Height          =   495
         Left            =   3240
         TabIndex        =   12
         Top             =   360
         Width           =   975
         _ExtentX        =   1725
         _ExtentY        =   863
         BTYPE           =   14
         TX              =   "&Filtrar"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   7.4717
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
         MICON           =   "frmconsultalibroventas_n.frx":0FA8
         PICN            =   "frmconsultalibroventas_n.frx":0FC4
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   315
         Left            =   2280
         TabIndex        =   15
         Top             =   1080
         Width           =   1815
         _ExtentX        =   3355
         _ExtentY        =   575
         _Version        =   393216
         AllowPrompt     =   -1  'True
         MaxLength       =   13
         Mask            =   "####-########"
         PromptChar      =   "_"
      End
      Begin KewlButtonz.KewlButtons command2 
         Height          =   495
         Left            =   4320
         TabIndex        =   17
         Top             =   360
         Width           =   1095
         _ExtentX        =   1941
         _ExtentY        =   863
         BTYPE           =   14
         TX              =   "&Imprime"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   7.4717
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
         MICON           =   "frmconsultalibroventas_n.frx":43B6
         PICN            =   "frmconsultalibroventas_n.frx":43D2
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
   Begin VB.Frame Frame3 
      Caption         =   "Ordenar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.47
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   4
      Top             =   1680
      Width           =   6375
      Begin VB.OptionButton Option1 
         Caption         =   "Nº Comp."
         Height          =   255
         Index           =   3
         Left            =   3960
         TabIndex        =   8
         Top             =   360
         Width           =   975
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Tipo Comp."
         Height          =   255
         Index           =   2
         Left            =   2640
         TabIndex        =   7
         Top             =   360
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Razon Social"
         Height          =   255
         Index           =   1
         Left            =   1200
         TabIndex        =   6
         Top             =   360
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Fecha"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   855
      End
      Begin KewlButtonz.KewlButtons KewlButtons1 
         Height          =   495
         Left            =   5520
         TabIndex        =   20
         Top             =   120
         Width           =   735
         _ExtentX        =   1294
         _ExtentY        =   863
         BTYPE           =   14
         TX              =   "&E.C."
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   7.4717
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
         MICON           =   "frmconsultalibroventas_n.frx":4DE4
         PICN            =   "frmconsultalibroventas_n.frx":4E00
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmconsultalibroventas_n.frx":5812
      Height          =   3615
      Left            =   120
      Negotiate       =   -1  'True
      TabIndex        =   0
      Top             =   2520
      Width           =   6375
      _ExtentX        =   11765
      _ExtentY        =   6662
      _Version        =   393216
      AllowUpdate     =   -1  'True
      AllowArrows     =   -1  'True
      BackColor       =   14737632
      HeadLines       =   4
      RowHeight       =   13
      TabAction       =   2
      WrapCellPointer =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.4717
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   6.7925
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
         MarqueeStyle    =   2
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc datPrimaryRS 
      Align           =   2  'Align Bottom
      Height          =   325
      Left            =   0
      Top             =   6071
      Visible         =   0   'False
      Width           =   6604
      _ExtentX        =   11646
      _ExtentY        =   575
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
      Caption         =   " "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.47
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
      _ExtentX        =   1150
      _ExtentY        =   1150
      CloseButton     =   0
      MaxButton       =   0
      MinButton       =   0
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
      LcK2            =   $"frmconsultalibroventas_n.frx":582D
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
   Begin Crystal.CrystalReport CrystalReporte 
      Left            =   0
      Top             =   720
      _ExtentX        =   671
      _ExtentY        =   671
      _Version        =   348160
      WindowTitle     =   "Libro IVA Compras"
      PrintFileLinesPerPage=   60
   End
   Begin MSRDC.MSRDC reporte 
      Height          =   375
      Left            =   720
      Top             =   0
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2228
      _ExtentY        =   695
      _Version        =   393216
      Options         =   0
      CursorDriver    =   0
      BOFAction       =   0
      EOFAction       =   0
      RecordsetType   =   1
      LockType        =   3
      QueryType       =   0
      Prompt          =   3
      Appearance      =   1
      QueryTimeout    =   30
      RowsetSize      =   100
      LoginTimeout    =   15
      KeysetSize      =   0
      MaxRows         =   0
      ErrorThreshold  =   -1
      BatchSize       =   15
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   -1  'True
      ReadOnly        =   0   'False
      Appearance      =   -1  'True
      DataSourceName  =   "contable"
      RecordSource    =   ""
      UserName        =   "sa"
      Password        =   ""
      Connect         =   ""
      LogMessages     =   ""
      Caption         =   "MSRDC1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.4717
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmconsutalibroventas_n"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim columna(15) As String
Dim posicion As Integer
Dim poscuenta As Integer
Dim Cuenta As Integer
Dim sumadebe As Currency
Dim sumahaber As Currency
Dim errorasiento As Boolean
Dim fechafuera As String


Private Sub borrar_Click()

End Sub

Private Sub buscar_Click()

        datPrimaryRS.Recordset.Filter = "tipocompr = '" & Combo1.Text & "' and numcompr = '" & MaskEdBox1.Text & "'"
        
        
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        MaskEdBox1.SetFocus
    End If

End Sub

Private Sub Command2_Click()

Dim tabla As String
Dim ruta As String

ruta = "\Empresa" + Right(Str(login.empresaact), Len(Str(login.empresaact)) - 1)

    reporte.SQL = "SELECT libroventas1.id, libroventas1.cerrado, libroventas1.empresa, libroventas1.inicioper, libroventas1.fecha, libroventas1.cliente, libroventas1.tipoiva, libroventas1.cuit, libroventas1.tipocompr, libroventas1.numcompr, libroventas1.col1, libroventas1.col2, libroventas1.col3, libroventas1.col4, libroventas1.col5, libroventas1.col6, libroventas1.total, libroventas1.nomcol1, libroventas1.nomcol2, libroventas1.nomcol3, libroventas1.nomcol4, libroventas1.nomcol5, libroventas1.nomcol6, libroventas1.nomcol7, libroventas1.razonsocial FROM contablesql.dbo.libroventas1 libroventas1 WHERE libroventas1.empresa = " & login.empresaact & " and libroventas1.inicioper = '" & login.iper & "' and fecha >= '" & desde.Value & "' and fecha <= '" & hasta.Value & "' ORDER BY libroventas1.fecha ASC, libroventas1.id ASC"

tabla = reporte.SQL

With CrystalReporte
    .ReportFileName = App.Path & ruta + "\libroventas consulta.rpt"
    .Connect = login.conexionreporte
    .DiscardSavedData = True
    .RetrieveDataFiles
    .ReportSource = 0
    .SQLQuery = tabla
    .Destination = crptToWindow
    .PrintFileType = crptCrystal
    .WindowState = crptMaximized
    .Action = 1
      
End With
 
End Sub

Private Sub DataGrid1_Click()

    Call llena_Click

End Sub

Private Sub DataGrid1_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        Call salir_Click
    End If

End Sub

Private Sub DataGrid1_KeyUp(KeyCode As Integer, Shift As Integer)

    Call llena_Click

End Sub

Private Sub desde_Change()

        Call filtrar_Click



End Sub

Private Sub filtrar_Click()
On Error Resume Next

    datPrimaryRS.RecordSource = "select libroventas.* from libroventas WHERE libroventas.empresa = " & login.empresaact & " and fecha >= '" & desde.Value & "' and fecha <= '" & hasta.Value & "' order by fecha"
    datPrimaryRS.Refresh
    
    If frmlibroventas_nuevo.denominacion.Text <> "" Then
        datPrimaryRS.Recordset.Filter = "cliente = '" & frmlibroventas_nuevo.denominacion.Text & "'"
    End If
    
For x = 0 To 74
    DataGrid1.Columns(x).Visible = False
Next x

    DataGrid1.Columns(2).Visible = True
    DataGrid1.Columns(3).Visible = True
    DataGrid1.Columns(6).Visible = True
    DataGrid1.Columns(7).Visible = True
    DataGrid1.SetFocus

End Sub

Private Sub Form_GotFocus()

 Call filtrar_Click

End Sub

Private Sub Form_Load()
Aplicar_skin Me


frmconsutalibroventas_n.Top = 800
frmconsutalibroventas_n.Left = frmlibroventas_nuevo.Text3(15).Left + frmlibroventas_nuevo.Text3(15).Width + 200

datPrimaryRS.ConnectionString = login.conexiontotal

                    Combo1.AddItem "F-A"
                    Combo1.AddItem "F-B"
                    Combo1.AddItem "F-M"
                    Combo1.AddItem "R-A"
                    Combo1.AddItem "R-B"
                    Combo1.AddItem "NDA"
                    Combo1.AddItem "NDB"
                    Combo1.AddItem "NCA"
                    Combo1.AddItem "NCB"
                    Combo1.AddItem "TFA"
                    Combo1.AddItem "TFB"
                    Combo1.AddItem "F-C"
                    Combo1.AddItem "R-C"
                    Combo1.AddItem "NDC"
                    Combo1.AddItem "NCC"
                    Combo1.AddItem "REC"
                    Combo1.AddItem "TKT"
                    Combo1.ListIndex = 0




desde.Value = Date - Day(Date) + 1
hasta.Value = Date
Call filtrar_Click


End Sub

Private Sub KewlButtons1_Click()


    impecclientes.Show
    impecclientes.cargadesde.Value = desde.Value
    impecclientes.cargahasta.Value = hasta.Value
    impecclientes.combodesde.Text = DataGrid1.Columns(3).Text
    impecclientes.combohasta.Text = DataGrid1.Columns(3).Text
    impecclientes.listar.SetFocus
    SendKeys "{ENTER}", True
    

End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
      If login.librocerrado = "S" Then
            frmlibroventas_nuevo.grabalibroasiento.Enabled = True
            frmlibroventas_nuevo.borrar.Enabled = True
      Else
            frmlibroventas_nuevo.grabalibroasiento.Enabled = False
            frmlibroventas_nuevo.borrar.Enabled = False
      End If

    frmlibroventas_nuevo.modificar = 1
    frmlibroventas_nuevo.Text3(0).SetFocus

End Sub

Private Sub llena_Click()
On Error Resume Next
      
    
    If datPrimaryRS.Recordset.Fields("cerrado") <> "N" Then
        cerrado.Visible = True
    Else
        cerrado.Visible = False
    End If
      
    frmlibroventas_nuevo.Text11.Text = datPrimaryRS.Recordset.Fields("id")
    frmlibroventas_nuevo.Maskfecha.Text = datPrimaryRS.Recordset.Fields("fecha")
    frmlibroventas_nuevo.denominacion.Text = datPrimaryRS.Recordset.Fields("cliente")
    frmlibroventas_nuevo.tipoiva.Text = datPrimaryRS.Recordset.Fields("tipoiva")
    frmlibroventas_nuevo.cuit.Text = datPrimaryRS.Recordset.Fields("cuit")
    frmlibroventas_nuevo.tipocomp.Text = datPrimaryRS.Recordset.Fields("tipocompr")
    frmlibroventas_nuevo.Maskcomprobante.Text = datPrimaryRS.Recordset.Fields("numcompr")
    If datPrimaryRS.Recordset.Fields("contado") <> "S" Or IsNull(datPrimaryRS.Recordset.Fields("contado")) = True Then
        frmlibroventas_nuevo.Check1(0).Value = 1
    Else
        frmlibroventas_nuevo.Check1(0).Value = 0
    End If
    For x = 0 To 14
        If datPrimaryRS.Recordset.Fields(8 + x) >= 0 Then
            frmlibroventas_nuevo.Text3(x).Text = datPrimaryRS.Recordset.Fields(8 + x)
        Else
            frmlibroventas_nuevo.Text3(x).Text = datPrimaryRS.Recordset.Fields(8 + x) * -1
        End If
        frmlibroventas_nuevo.Text3(x).Text = Format(frmlibroventas_nuevo.Text3(x).Text, "#,##0.00")
    Next x
        If datPrimaryRS.Recordset.Fields("total") >= 0 Then
            frmlibroventas_nuevo.Text3(15).Text = datPrimaryRS.Recordset.Fields("total")
        Else
            frmlibroventas_nuevo.Text3(15).Text = datPrimaryRS.Recordset.Fields("total") * -1
        End If
        Debug.Print frmlibroventas_nuevo.Text3(15).Text
        frmlibroventas_nuevo.Text3(15).Text = Format(frmlibroventas_nuevo.Text3(15).Text, "#,##0.00")
        frmlibroventas_nuevo.Text7(30).Text = datPrimaryRS.Recordset.Fields("cdt")
    Y = 27
    For x = 1 To 29 Step 2
        If IsNull(datPrimaryRS.Recordset.Fields(Y)) = True Then
            frmlibroventas_nuevo.Text7(x).Text = 0
        Else
            frmlibroventas_nuevo.Text7(x).Text = datPrimaryRS.Recordset.Fields(Y)
        End If
        Y = Y + 2
    Next x
    frmlibroventas_nuevo.Text9.Text = datPrimaryRS.Recordset.Fields("ccosto")
    

End Sub

Private Sub MaskEdBox1_KeyPress(KeyAscii As Integer)
On Error Resume Next

    If KeyAscii = 13 Then
        KeyAscii = 0
            car = 0
            car1 = 0
            For x = 6 To 13
                If Mid(MaskEdBox1.Text, x, 1) = "_" Then
                    car = car + 1
                Else
                    car1 = car1 + 1
                End If
            Next x
            MaskEdBox1.Text = Mid(MaskEdBox1.Text, 1, 4) + "-" + Mid("0000000", 1, car) + Mid(MaskEdBox1.Text, 6, car1)
            buscar.SetFocus
    End If

End Sub

Private Sub Option1_Click(Index As Integer)


    If Option1(0).Value = True Then
        datPrimaryRS.Recordset.Sort = "fecha"
    End If
    If Option1(1).Value = True Then
        datPrimaryRS.Recordset.Sort = "cliente"
    End If
    If Option1(2).Value = True Then
        datPrimaryRS.Recordset.Sort = "tipocompr"
    End If
    If Option1(3).Value = True Then
        datPrimaryRS.Recordset.Sort = "numcompr"
    End If


End Sub

Private Sub salir_Click()

    Unload Me

End Sub
