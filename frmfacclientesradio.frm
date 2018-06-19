VERSION 5.00
Object = "{F6125AB1-8AB1-11CE-A77F-08002B2F4E98}#2.0#0"; "MSRDC20.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{5734474E-78D3-4254-99B9-C35F31BDF509}#62.0#0"; "vbskpro2.ocx"
Begin VB.Form frmfacclientesradio 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Facturacion a Clientes"
   ClientHeight    =   7845
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10185
   Icon            =   "frmfacclientesradio.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7845
   ScaleWidth      =   10185
   Begin MSDataListLib.DataList DataList1 
      Bindings        =   "frmfacclientesradio.frx":0442
      Height          =   1425
      Left            =   5880
      TabIndex        =   20
      Top             =   720
      Visible         =   0   'False
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   2514
      _Version        =   393216
      MatchEntry      =   -1  'True
      BackColor       =   -2147483643
      ListField       =   "razonsocial"
      BoundColumn     =   "domicilio"
      Object.DataMember      =   ""
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00C0FFFF&
      Height          =   315
      Left            =   8640
      TabIndex        =   80
      Text            =   "Combo1"
      Top             =   480
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc datcondventa 
      Height          =   330
      Left            =   5520
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
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
   Begin VB.TextBox Text19 
      Alignment       =   1  'Right Justify
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#.##0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   3082
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Left            =   6600
      TabIndex        =   78
      Top             =   6840
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton notadebito 
      Caption         =   "Nota &Debito"
      Height          =   615
      Left            =   8640
      Style           =   1  'Graphical
      TabIndex        =   77
      Top             =   1800
      UseMaskColor    =   -1  'True
      Width           =   1335
   End
   Begin VB.CommandButton notacredito 
      Caption         =   "Nota &Credito"
      Height          =   615
      Left            =   8640
      Style           =   1  'Graphical
      TabIndex        =   76
      Top             =   1200
      UseMaskColor    =   -1  'True
      Width           =   1335
   End
   Begin MSDataListLib.DataList DataList2 
      Bindings        =   "frmfacclientesradio.frx":045C
      Height          =   1815
      Left            =   600
      TabIndex        =   24
      Top             =   4080
      Visible         =   0   'False
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   3201
      _Version        =   393216
      MatchEntry      =   -1  'True
      BackColor       =   14737632
      ListField       =   "lista"
      BoundColumn     =   "codprod"
      Object.DataMember      =   ""
   End
   Begin MSDataGridLib.DataGrid DataGrid3 
      Bindings        =   "frmfacclientesradio.frx":047B
      Height          =   375
      Left            =   5880
      TabIndex        =   25
      Top             =   4200
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
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
         DataField       =   "codprod"
         Caption         =   "codprod"
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
         DataField       =   "unidadmedida"
         Caption         =   "unidadmedida"
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
         DataField       =   "detalle"
         Caption         =   "detalle"
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
         DataField       =   "preciounit"
         Caption         =   "preciounit"
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
         DataField       =   "columnalibro"
         Caption         =   "columnalibro"
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
      BeginProperty Column05 
         DataField       =   "codcuenta"
         Caption         =   "codcuenta"
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
      BeginProperty Column06 
         DataField       =   "centrocosto"
         Caption         =   "centrocosto"
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
      BeginProperty Column07 
         DataField       =   "empresa"
         Caption         =   "empresa"
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
         EndProperty
         BeginProperty Column01 
         EndProperty
         BeginProperty Column02 
         EndProperty
         BeginProperty Column03 
         EndProperty
         BeginProperty Column04 
         EndProperty
         BeginProperty Column05 
         EndProperty
         BeginProperty Column06 
         EndProperty
         BeginProperty Column07 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton salir 
      Caption         =   "&Salir"
      Height          =   975
      Left            =   8640
      Picture         =   "frmfacclientesradio.frx":0496
      Style           =   1  'Graphical
      TabIndex        =   75
      Top             =   6000
      UseMaskColor    =   -1  'True
      Width           =   1335
   End
   Begin VB.CommandButton numletras 
      Caption         =   "numletras"
      Height          =   255
      Left            =   8520
      TabIndex        =   67
      Top             =   6000
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox Text18 
      Height          =   1095
      Left            =   2760
      MaxLength       =   160
      MultiLine       =   -1  'True
      TabIndex        =   66
      Top             =   4200
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   255
      Left            =   7920
      TabIndex        =   63
      Top             =   360
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox Text13 
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   6000
      MaxLength       =   10
      TabIndex        =   10
      Top             =   3360
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox Text12 
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   2160
      MaxLength       =   20
      TabIndex        =   9
      Top             =   3360
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton facturar 
      Caption         =   "&Facturar"
      Height          =   975
      Left            =   8640
      Picture         =   "frmfacclientesradio.frx":08D8
      Style           =   1  'Graphical
      TabIndex        =   57
      Top             =   2520
      UseMaskColor    =   -1  'True
      Width           =   1335
   End
   Begin VB.TextBox tapa 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   1695
      Left            =   360
      Locked          =   -1  'True
      TabIndex        =   43
      Top             =   5880
      Width           =   4575
   End
   Begin MSMask.MaskEdBox Masktotal 
      Height          =   495
      Left            =   6240
      TabIndex        =   42
      Top             =   6000
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   873
      _Version        =   393216
      BorderStyle     =   0
      Appearance      =   0
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "#,##0.00;-#,##0.00"
      PromptChar      =   "_"
   End
   Begin VB.CommandButton calcula 
      Caption         =   "calcula"
      Height          =   255
      Left            =   8280
      TabIndex        =   40
      Top             =   3000
      Visible         =   0   'False
      Width           =   1095
   End
   Begin MSDataListLib.DataList DataList3 
      Bindings        =   "frmfacclientesradio.frx":0D1A
      Height          =   1035
      Left            =   1920
      TabIndex        =   39
      Top             =   1800
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1826
      _Version        =   393216
      ListField       =   "descripcion"
      BoundColumn     =   "categ"
      Object.DataMember      =   "condtrib"
   End
   Begin MSMask.MaskEdBox text4 
      Height          =   285
      Left            =   4440
      TabIndex        =   4
      Top             =   1800
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   503
      _Version        =   393216
      BackColor       =   14737632
      MaxLength       =   13
      Mask            =   "##-########-#"
      PromptChar      =   "_"
   End
   Begin VB.TextBox Text9 
      Alignment       =   1  'Right Justify
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#.##0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   3082
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Left            =   6600
      TabIndex        =   38
      Top             =   6480
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton proximafact 
      Caption         =   "proximafact"
      Height          =   255
      Left            =   8280
      TabIndex        =   29
      Top             =   2640
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
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
      Left            =   600
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   240
      Width           =   1815
   End
   Begin VB.CommandButton buscafact 
      Caption         =   "buscafact"
      Height          =   255
      Left            =   8280
      TabIndex        =   26
      Top             =   2520
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Text6 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
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
      Left            =   3000
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   240
      Width           =   975
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   1320
      MaxLength       =   50
      TabIndex        =   2
      Top             =   720
      Width           =   4575
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   1320
      MaxLength       =   50
      TabIndex        =   1
      Top             =   1080
      Width           =   4575
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Bindings        =   "frmfacclientesradio.frx":0D2E
      Height          =   315
      Left            =   2400
      TabIndex        =   8
      Top             =   3000
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   556
      _Version        =   393216
      BackColor       =   14737632
      ListField       =   "descripcion"
      BoundColumn     =   "codigo"
      Text            =   ""
      Object.DataMember      =   ""
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   1320
      TabIndex        =   3
      Top             =   1800
      Width           =   615
   End
   Begin MSMask.MaskEdBox maskfecha 
      Height          =   285
      Left            =   6720
      TabIndex        =   0
      Top             =   240
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   503
      _Version        =   393216
      BackColor       =   14737632
      MaxLength       =   10
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSDataGridLib.DataGrid DataGrid2 
      Bindings        =   "frmfacclientesradio.frx":0D49
      Height          =   375
      Left            =   8280
      TabIndex        =   21
      Top             =   1200
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
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
      ColumnCount     =   13
      BeginProperty Column00 
         DataField       =   "codcliente"
         Caption         =   "codcliente"
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
         DataField       =   "empresa"
         Caption         =   "empresa"
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
         DataField       =   "tipocliente"
         Caption         =   "tipocliente"
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
         DataField       =   "razonsocial"
         Caption         =   "razonsocial"
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
         DataField       =   "tipoiva"
         Caption         =   "tipoiva"
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
      BeginProperty Column05 
         DataField       =   "cuit"
         Caption         =   "cuit"
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
      BeginProperty Column06 
         DataField       =   "domicilio"
         Caption         =   "domicilio"
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
      BeginProperty Column07 
         DataField       =   "localidad"
         Caption         =   "localidad"
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
      BeginProperty Column08 
         DataField       =   "codpostal"
         Caption         =   "codpostal"
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
      BeginProperty Column09 
         DataField       =   "telefono"
         Caption         =   "telefono"
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
      BeginProperty Column10 
         DataField       =   "email"
         Caption         =   "email"
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
      BeginProperty Column11 
         DataField       =   "contacto"
         Caption         =   "contacto"
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
      BeginProperty Column12 
         DataField       =   "codcontable"
         Caption         =   "codcontable"
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
         EndProperty
         BeginProperty Column01 
         EndProperty
         BeginProperty Column02 
         EndProperty
         BeginProperty Column03 
         EndProperty
         BeginProperty Column04 
         EndProperty
         BeginProperty Column05 
         EndProperty
         BeginProperty Column06 
         EndProperty
         BeginProperty Column07 
         EndProperty
         BeginProperty Column08 
         EndProperty
         BeginProperty Column09 
         EndProperty
         BeginProperty Column10 
         EndProperty
         BeginProperty Column11 
         EndProperty
         BeginProperty Column12 
         EndProperty
      EndProperty
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   6120
      MaxLength       =   10
      TabIndex        =   11
      Top             =   3000
      Width           =   1575
   End
   Begin MSAdodcLib.Adodc datconsproductos 
      Height          =   330
      Left            =   6240
      Top             =   7800
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
   Begin MSAdodcLib.Adodc datproductos 
      Height          =   570
      Left            =   7320
      Top             =   7560
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   1005
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
      Bindings        =   "frmfacclientesradio.frx":0D63
      Height          =   2055
      Left            =   240
      TabIndex        =   12
      Top             =   3720
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   3625
      _Version        =   393216
      AllowUpdate     =   -1  'True
      AllowArrows     =   -1  'True
      BackColor       =   12648447
      Enabled         =   0   'False
      ColumnHeaders   =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      TabAction       =   1
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
      ColumnCount     =   19
      BeginProperty Column00 
         DataField       =   "numdisco"
         Caption         =   "numdisco"
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
         DataField       =   "tipocomp"
         Caption         =   "tipocomp"
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
         DataField       =   "comprobante"
         Caption         =   "comprobante"
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
         DataField       =   "id"
         Caption         =   "id"
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
         DataField       =   "codproducto"
         Caption         =   "Cod.Prod."
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
         DataField       =   "cant"
         Caption         =   "Cant."
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "unidadmedida"
         Caption         =   "U.Med."
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
         DataField       =   "detalle"
         Caption         =   "Descripcion"
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
      BeginProperty Column08 
         DataField       =   "preciounit"
         Caption         =   "Precio Unit."
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column09 
         DataField       =   "totales"
         Caption         =   "Total Bruto"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column10 
         DataField       =   "empresa"
         Caption         =   "empresa"
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
      BeginProperty Column11 
         DataField       =   "codcuenta"
         Caption         =   "codcuenta"
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
      BeginProperty Column12 
         DataField       =   "collibro"
         Caption         =   "collibro"
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
      BeginProperty Column13 
         DataField       =   "centrocosto"
         Caption         =   "centrocosto"
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
      BeginProperty Column14 
         DataField       =   "facturado"
         Caption         =   "facturado"
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
      BeginProperty Column15 
         DataField       =   "gravado"
         Caption         =   "gravado"
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
      BeginProperty Column16 
         DataField       =   "descuento"
         Caption         =   "Desc.%"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column17 
         DataField       =   "impdesc"
         Caption         =   "Imp.Dec."
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column18 
         DataField       =   "totalneto"
         Caption         =   "Total Neto"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnAllowSizing=   -1  'True
            Object.Visible         =   0   'False
            ColumnWidth     =   0
         EndProperty
         BeginProperty Column01 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column02 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column03 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column04 
            Alignment       =   2
            Button          =   -1  'True
            Object.Visible         =   -1  'True
         EndProperty
         BeginProperty Column05 
            Alignment       =   2
         EndProperty
         BeginProperty Column06 
            Alignment       =   2
         EndProperty
         BeginProperty Column07 
         EndProperty
         BeginProperty Column08 
            Alignment       =   1
         EndProperty
         BeginProperty Column09 
            Alignment       =   1
         EndProperty
         BeginProperty Column10 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column11 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column12 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column13 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column14 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column15 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column16 
            Alignment       =   2
         EndProperty
         BeginProperty Column17 
            Alignment       =   1
            Locked          =   -1  'True
         EndProperty
         BeginProperty Column18 
            Alignment       =   1
            Locked          =   -1  'True
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc datfacclientes 
      Height          =   330
      Left            =   5400
      Top             =   7800
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
   Begin MSAdodcLib.Adodc datPrimaryRS 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      Top             =   7515
      Visible         =   0   'False
      Width           =   10185
      _ExtentX        =   17965
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
   Begin MSAdodcLib.Adodc datclientes 
      Height          =   330
      Left            =   6000
      Top             =   7320
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
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
   Begin MSAdodcLib.Adodc datcolumnas 
      Height          =   570
      Left            =   4200
      Top             =   7560
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1005
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
   Begin MSAdodcLib.Adodc datparamventas 
      Height          =   330
      Left            =   7200
      Top             =   7320
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
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
   Begin MSAdodcLib.Adodc datmaestro 
      Height          =   330
      Left            =   9000
      Top             =   0
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
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
   Begin MSAdodcLib.Adodc datperiodo 
      Height          =   330
      Left            =   9000
      Top             =   240
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
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
   Begin MSAdodcLib.Adodc datasiento 
      Height          =   330
      Left            =   8160
      Top             =   0
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
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
   Begin VB.TextBox Text14 
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   1320
      MaxLength       =   30
      TabIndex        =   5
      Top             =   2160
      Width           =   2895
   End
   Begin VB.TextBox Text15 
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   5280
      MaxLength       =   30
      TabIndex        =   6
      Top             =   2160
      Width           =   2895
   End
   Begin VB.TextBox Text16 
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   1320
      MaxLength       =   15
      TabIndex        =   7
      Top             =   2520
      Width           =   1935
   End
   Begin VB.TextBox Text10 
      Alignment       =   1  'Right Justify
      DataField       =   "col1"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   0
      Left            =   8280
      TabIndex        =   41
      Text            =   "Text10"
      Top             =   3480
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Text10 
      Alignment       =   1  'Right Justify
      DataField       =   "col2"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   1
      Left            =   8280
      TabIndex        =   44
      Text            =   "Text10"
      Top             =   3720
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Text10 
      Alignment       =   1  'Right Justify
      DataField       =   "col3"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   2
      Left            =   8280
      TabIndex        =   45
      Text            =   "Text10"
      Top             =   3960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Text10 
      Alignment       =   1  'Right Justify
      DataField       =   "col4"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   3
      Left            =   8280
      TabIndex        =   46
      Text            =   "Text10"
      Top             =   4200
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Text10 
      Alignment       =   1  'Right Justify
      DataField       =   "col5"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   4
      Left            =   8280
      TabIndex        =   47
      Text            =   "Text10"
      Top             =   4440
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Text10 
      Alignment       =   1  'Right Justify
      DataField       =   "col6"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   5
      Left            =   8280
      TabIndex        =   48
      Text            =   "Text10"
      Top             =   4680
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Text10 
      Alignment       =   1  'Right Justify
      DataField       =   "col7"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   6
      Left            =   8280
      TabIndex        =   49
      Text            =   "Text10"
      Top             =   4920
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Text11 
      Alignment       =   2  'Center
      DataField       =   "ch1"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   0
      Left            =   9000
      TabIndex        =   50
      Text            =   "Text11"
      Top             =   3480
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Text11 
      Alignment       =   2  'Center
      DataField       =   "ch2"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   1
      Left            =   9000
      TabIndex        =   51
      Text            =   "Text11"
      Top             =   3720
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Text11 
      Alignment       =   2  'Center
      DataField       =   "ch3"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   2
      Left            =   9000
      TabIndex        =   52
      Text            =   "Text11"
      Top             =   3960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Text11 
      Alignment       =   2  'Center
      DataField       =   "ch4"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   3
      Left            =   9000
      TabIndex        =   53
      Text            =   "Text11"
      Top             =   4200
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Text11 
      Alignment       =   2  'Center
      DataField       =   "ch5"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   4
      Left            =   9000
      TabIndex        =   54
      Text            =   "Text11"
      Top             =   4440
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Text11 
      Alignment       =   2  'Center
      DataField       =   "ch6"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   5
      Left            =   9000
      TabIndex        =   55
      Text            =   "Text11"
      Top             =   4680
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Text11 
      Alignment       =   2  'Center
      DataField       =   "ch7"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   6
      Left            =   9000
      TabIndex        =   56
      Text            =   "Text11"
      Top             =   4920
      Visible         =   0   'False
      Width           =   735
   End
   Begin Crystal.CrystalReport CrystalReporte 
      Left            =   7440
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTitle     =   "Factura A"
      PrintFileLinesPerPage=   60
   End
   Begin MSRDC.MSRDC reporte 
      Height          =   375
      Left            =   7800
      Top             =   0
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   661
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
      Connect         =   "PROVIDER=MSDASQL;dsn=contable;uid=sa;pwd=;database=contablesql;"
      LogMessages     =   ""
      Caption         =   "MSRDC1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox Text17 
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   1320
      MaxLength       =   50
      TabIndex        =   65
      Top             =   1440
      Width           =   4575
   End
   Begin Project1.jeffMaskedEdit Text8 
      Height          =   255
      Index           =   1
      Left            =   3360
      TabIndex        =   68
      Top             =   6120
      Width           =   1215
      _extentx        =   1931
      _extenty        =   450
      mouseicon       =   "frmfacclientesradio.frx":0D80
      font            =   "frmfacclientesradio.frx":0D9E
      format          =   "##,##0.00"
      seltext         =   ""
      alignment       =   1
   End
   Begin Project1.jeffMaskedEdit Text8 
      Height          =   255
      Index           =   2
      Left            =   3360
      TabIndex        =   69
      Top             =   6360
      Width           =   1215
      _extentx        =   1931
      _extenty        =   450
      mouseicon       =   "frmfacclientesradio.frx":0DCA
      font            =   "frmfacclientesradio.frx":0DE8
      format          =   "##,##0.00"
      seltext         =   ""
      alignment       =   1
   End
   Begin Project1.jeffMaskedEdit Text8 
      Height          =   255
      Index           =   3
      Left            =   3360
      TabIndex        =   70
      Top             =   6600
      Width           =   1215
      _extentx        =   1931
      _extenty        =   450
      mouseicon       =   "frmfacclientesradio.frx":0E14
      font            =   "frmfacclientesradio.frx":0E32
      format          =   "##,##0.00"
      seltext         =   ""
      alignment       =   1
   End
   Begin Project1.jeffMaskedEdit Text8 
      Height          =   255
      Index           =   4
      Left            =   3360
      TabIndex        =   71
      Top             =   6840
      Width           =   1215
      _extentx        =   1931
      _extenty        =   450
      mouseicon       =   "frmfacclientesradio.frx":0E5E
      font            =   "frmfacclientesradio.frx":0E7C
      format          =   "##,##0.00"
      seltext         =   ""
      alignment       =   1
   End
   Begin Project1.jeffMaskedEdit Text8 
      Height          =   255
      Index           =   5
      Left            =   3360
      TabIndex        =   72
      Top             =   7080
      Width           =   1215
      _extentx        =   1931
      _extenty        =   450
      mouseicon       =   "frmfacclientesradio.frx":0EA8
      font            =   "frmfacclientesradio.frx":0EC6
      format          =   "##,##0.00"
      seltext         =   ""
      alignment       =   1
   End
   Begin Project1.jeffMaskedEdit Text8 
      Height          =   255
      Index           =   6
      Left            =   3360
      TabIndex        =   73
      Top             =   7320
      Width           =   1215
      _extentx        =   1931
      _extenty        =   450
      mouseicon       =   "frmfacclientesradio.frx":0EF2
      font            =   "frmfacclientesradio.frx":0F10
      format          =   "##,##0.00"
      seltext         =   ""
      alignment       =   1
   End
   Begin Project1.jeffMaskedEdit Text8 
      Height          =   255
      Index           =   0
      Left            =   3360
      TabIndex        =   74
      Top             =   5880
      Width           =   1215
      _extentx        =   1931
      _extenty        =   450
      mouseicon       =   "frmfacclientesradio.frx":0F3C
      font            =   "frmfacclientesradio.frx":0F5A
      format          =   "##,##0.00"
      seltext         =   ""
      alignment       =   1
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
      LcK2            =   $"frmfacclientesradio.frx":0F86
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
   Begin MSAdodcLib.Adodc datempresa 
      Height          =   330
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
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
   Begin VB.Label Label18 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Punto de Venta"
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
      Left            =   8520
      TabIndex        =   79
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "C.U.I.T.:"
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
      Left            =   3600
      TabIndex        =   14
      Top             =   1800
      Width           =   735
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "Localidad"
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
      Left            =   360
      TabIndex        =   64
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "Tel.:"
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
      Left            =   360
      TabIndex        =   62
      Top             =   2520
      Width           =   855
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Producto:"
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
      Left            =   4320
      TabIndex        =   61
      Top             =   2160
      Width           =   855
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Avisador:"
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
      Left            =   360
      TabIndex        =   60
      Top             =   2160
      Width           =   855
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Cod.Operacion:"
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
      Left            =   4560
      TabIndex        =   59
      Top             =   3360
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre de Tarjeta:"
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
      Left            =   360
      TabIndex        =   58
      Top             =   3360
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Shape Shape3 
      Height          =   1935
      Left            =   240
      Top             =   5760
      Width           =   8055
   End
   Begin VB.Label Label11 
      Caption         =   "Total: $"
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
      Left            =   5400
      TabIndex        =   37
      Top             =   6000
      Width           =   735
   End
   Begin VB.Label Label10 
      Height          =   225
      Index           =   6
      Left            =   480
      TabIndex        =   36
      Top             =   7320
      Width           =   2775
   End
   Begin VB.Label Label10 
      Height          =   225
      Index           =   5
      Left            =   480
      TabIndex        =   35
      Top             =   7080
      Width           =   2775
   End
   Begin VB.Label Label10 
      Height          =   225
      Index           =   1
      Left            =   480
      TabIndex        =   34
      Top             =   6120
      Width           =   2775
   End
   Begin VB.Label Label10 
      Height          =   225
      Index           =   4
      Left            =   480
      TabIndex        =   33
      Top             =   6840
      Width           =   2775
   End
   Begin VB.Label Label10 
      Height          =   225
      Index           =   3
      Left            =   480
      TabIndex        =   32
      Top             =   6600
      Width           =   2775
   End
   Begin VB.Label Label10 
      Height          =   225
      Index           =   2
      Left            =   480
      TabIndex        =   31
      Top             =   6360
      Width           =   2775
   End
   Begin VB.Label Label10 
      Height          =   225
      Index           =   0
      Left            =   480
      TabIndex        =   30
      Top             =   5880
      Width           =   2775
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Ultimo Comprobante emitido"
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
      Left            =   360
      TabIndex        =   28
      Top             =   0
      Width           =   2535
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Comprobante"
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
      Left            =   2880
      TabIndex        =   23
      Top             =   0
      Width           =   1215
   End
   Begin VB.Shape Shape2 
      Height          =   2295
      Left            =   240
      Top             =   600
      Width           =   8055
   End
   Begin VB.Shape Shape1 
      Height          =   855
      Left            =   240
      Top             =   2880
      Width           =   8055
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Orden de Publ. Nº:"
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
      Left            =   4440
      TabIndex        =   19
      Top             =   3000
      Width           =   1695
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Condiciones de Venta:"
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
      Left            =   360
      TabIndex        =   18
      Top             =   3000
      Width           =   2055
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Señor (es):"
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
      Left            =   360
      TabIndex        =   17
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Domicilio:"
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
      Left            =   360
      TabIndex        =   16
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "I.V.A."
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
      Left            =   360
      TabIndex        =   15
      Top             =   1800
      Width           =   615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha:"
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
      Left            =   5880
      TabIndex        =   13
      Top             =   240
      Width           =   735
   End
End
Attribute VB_Name = "frmfacclientesradio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim s As String
Dim proximafactura As String
Dim codigodebe As Integer
Dim totalparcial(50, 15) As Double
Dim totalalicuota(50, 15) As Double
Dim totalalic(15) As Double
Dim cuentaparcial(50) As Integer
Dim centroparcial(50) As Integer
Dim collib As Integer
Dim collibro(50) As Integer
Dim codhab As Integer
Dim codcen As Integer
Dim codcontable(50) As Integer
Dim centrocosto(50) As Integer
Dim posicioniva As Integer
Dim tipocomprobante As String
Dim numerodedisco As String
Dim saltogrid As Integer
Dim empresareal As Integer
Dim letrasnumeros As String
Private Function numletras0(ByVal Numero As Double, ByVal Estilo As Integer) As String
  Dim NumTmp As String
  Dim c01 As Integer
  Dim c02 As Integer
  Dim pos As Integer
  Dim dig As Integer
  Dim cen As Integer
  Dim dec As Integer
  Dim uni As Integer
  Dim letra1 As String
  Dim letra2 As String
  Dim letra3 As String
  Dim Leyenda As String
  Dim Leyenda1 As String
  Dim TFNumero As String
        
  If Numero < 0 Then Numero = Abs(Numero)

  NumTmp = Format(Numero, "000000000000000")        'Le da un formato fijo
  c01 = 1
  pos = 1
  TFNumero = ""
  'Para extraer tres digitos cada vez
  Do While c01 <= 5
    c02 = 1
    Do While c02 <= 3
      'Extrae un digito cada vez de izquierda a derecha
      dig = Val(Mid(NumTmp, pos, 1))
      Select Case c02
        Case 1: cen = dig
        Case 2: dec = dig
        Case 3: uni = dig
      End Select
      c02 = c02 + 1
      pos = pos + 1
    Loop
    letra3 = Centena(uni, dec, cen)
    letra2 = Decena(uni, dec)
    letra1 = Unidad(uni, dec)
            
    Select Case c01
      Case 1
        If cen + dec + uni = 1 Then
          Leyenda = "Billon "
        ElseIf cen + dec + uni > 1 Then
          Leyenda = "Billones "
        End If
      Case 2
        If cen + dec + uni >= 1 And Val(Mid(NumTmp, 7, 3)) = 0 Then
          Leyenda = "Mil Millones "
        ElseIf cen + dec + uni >= 1 Then
          Leyenda = "Mil "
        End If
      Case 3
        If cen + dec = 0 And uni = 1 Then
          Leyenda = "Millon "
        ElseIf cen > 0 Or dec > 0 Or uni > 1 Then
          Leyenda = "Millones "
        End If
      Case 4
        If cen + dec + uni >= 1 Then
          Leyenda = "Mil "
        End If
      Case 5
        If cen + dec + uni >= 1 Then
          Leyenda = ""
        End If
      End Select
            
      c01 = c01 + 1
      TFNumero = TFNumero + letra3 + letra2 + letra1 + Leyenda
      
      Leyenda = ""
      letra1 = ""
      letra2 = ""
      letra3 = ""
  Loop
       
  
  
  TFNumero = TFNumero
  Select Case Estilo
    Case 1
      TFNumero = StrConv(TFNumero, vbUpperCase)
    Case 2
      TFNumero = StrConv(TFNumero, vbLowerCase)
    Case Else
      TFNumero = StrConv(TFNumero, vbProperCase)
  End Select
  
  TFNumero = TFNumero & Mid(NumTmp, 17)
            
  numletras0 = TFNumero
    
End Function

Private Function Centena(ByVal uni As Integer, ByVal dec As Integer, _
                         ByVal cen As Integer) As String
Dim cTexto As String

  Select Case cen
    Case 1
      If dec + uni = 0 Then
        cTexto = "cien "
      Else
        cTexto = "ciento "
      End If
    Case 2: cTexto = "doscientos "
    Case 3: cTexto = "trescientos "
    Case 4: cTexto = "cuatrocientos "
    Case 5: cTexto = "quinientos "
    Case 6: cTexto = "seiscientos "
    Case 7: cTexto = "setecientos "
    Case 8: cTexto = "ochocientos "
    Case 9: cTexto = "novecientos "
    Case Else: cTexto = ""
  End Select
  Centena = cTexto
    
End Function

Private Function Decena(ByVal uni As Integer, ByVal dec As Integer) As String
Dim cTexto As String
  
  Select Case dec
    Case 1:
      Select Case uni
        Case 0: cTexto = "diez "
        Case 1: cTexto = "once "
        Case 2: cTexto = "doce "
        Case 3: cTexto = "trece "
        Case 4: cTexto = "catorce "
        Case 5: cTexto = "quince "
        Case 6 To 9: cTexto = "dieci"
      End Select
    Case 2:
      If uni = 0 Then
        cTexto = "veinte "
      ElseIf uni > 0 Then
        cTexto = "veinti"
      End If
    Case 3: cTexto = "treinta "
    Case 4: cTexto = "cuarenta "
    Case 5: cTexto = "cincuenta "
    Case 6: cTexto = "sesenta "
    Case 7: cTexto = "setenta "
    Case 8: cTexto = "ochenta "
    Case 9: cTexto = "noventa "
    Case Else: cTexto = ""
  End Select
  
  If uni > 0 And dec > 2 Then cTexto = cTexto + "y "
    
  Decena = cTexto
  
End Function

Private Function Unidad(ByVal uni As Integer, ByVal dec As Integer) As String
Dim cTexto As String
  
  If dec <> 1 Then
    Select Case uni
      Case 1: cTexto = "un "
      Case 2: cTexto = "dos "
      Case 3: cTexto = "tres "
      Case 4: cTexto = "cuatro "
      Case 5: cTexto = "cinco "
    End Select
  End If
  Select Case uni
    Case 6: cTexto = "seis "
    Case 7: cTexto = "siete "
    Case 8: cTexto = "ocho "
    Case 9: cTexto = "nueve "
  End Select
  
  Unidad = cTexto

End Function


Private Sub buscafact_Click()
 
 puntoventa = Combo1.Text
 puntofinal = Val(puntoventa) + 1
 puntofinal1 = Right("0000" + Right(Str(puntofinal), Len(Str(puntofinal)) - 1), 4)
 
 If Text6.Text = "F-A" Or Text6.Text = "NCA" Or Text6.Text = "NDA" Then
  datprimaryrs.RecordSource = "select libroventas.* from libroventas WHERE libroventas.empresa = " & login.empresaact & " and (tipocompr = 'F-A' or tipocompr = 'NCA' or tipocompr = 'NDA') and numcompr >= '" & puntoventa & "' and numcompr < '" & puntofinal1 & "' Order by numcompr"
  datprimaryrs.Refresh
 End If
 If Text6.Text = "F-B" Or Text6.Text = "NCB" Or Text6.Text = "NDB" Then
  datprimaryrs.RecordSource = "select libroventas.* from libroventas WHERE libroventas.empresa = " & login.empresaact & " and (tipocompr = 'F-B' or tipocompr = 'NCB' or tipocompr = 'NDB') and numcompr >= '" & puntoventa & "' and numcompr < '" & puntofinal1 & "' Order by numcompr"
  datprimaryrs.Refresh
 End If
Rem If Text6.Text = "NCA" Then
Rem  datPrimaryRS.RecordSource = "select libroventas.* from libroventas WHERE libroventas.empresa = " & login.empresaact & " and tipocompr = 'NCA' Order by id"
Rem  datPrimaryRS.Refresh
Rem End If
Rem If Text6.Text = "NCB" Then
Rem  datPrimaryRS.RecordSource = "select libroventas.* from libroventas WHERE libroventas.empresa = " & login.empresaact & " and tipocompr = 'NCB' Order by id"
Rem  datPrimaryRS.Refresh
Rem End If
Rem If Text6.Text = "NDA" Then
Rem  datPrimaryRS.RecordSource = "select libroventas.* from libroventas WHERE libroventas.empresa = " & login.empresaact & " and tipocompr = 'NDA' Order by id"
Rem  datPrimaryRS.Refresh
Rem End If
Rem If Text6.Text = "NDB" Then
Rem  datPrimaryRS.RecordSource = "select libroventas.* from libroventas WHERE libroventas.empresa = " & login.empresaact & " and tipocompr = 'NCB' Order by id"
Rem  datPrimaryRS.Refresh
Rem End If
  
  
  If datprimaryrs.Recordset.EOF = True Then
    Text7.Text = Combo1.Text + "-00000000"
    Exit Sub
  End If
  datprimaryrs.Recordset.MoveLast
  Text7.Text = datprimaryrs.Recordset.Fields(7)

End Sub



Private Sub calcula_Click()
Dim totalgral As Double
Dim sumatotallic As Double

totalgral = 0
sumatotalalic = 0
For x = 1 To 15
    totalalic(x) = 0
Next x
For x = 0 To 6
    Text8(x).Text = 0
Next x

Rem ********************* factura B ******************
Text9.Text = 0
If Text6.Text = "F-B" Or Text6.Text = "NCB" Or Text6.Text = "NDB" Then
    For x = 0 To DataGrid1.VisibleRows - 2
      For Y = 1 To 15
        If IsNull(totalparcial(x, Y)) = True Then totalparcial(x, Y) = 0
        If IsNull(totalalicuota(x, Y)) = True Then totalalicuota(x, Y) = 0
        totalgral = totalparcial(x, Y) + totalgral
        totalalic(Y) = totalalicuota(x, Y) + totalalic(Y)
      Next Y
    Next x
    If Text6.Text = "NCB" Then totalgral = totalgral * -1
    Text9.Text = totalgral
    Masktotal = Text9.Text
For x = 1 To 15
    If Text8(x - 1).Visible = False Then GoTo paso0
    Text8(x - 1).Text = totalalic(x)
    If Text6.Text = "NCB" Then Text8(x - 1).Text = Text8(x - 1).Text * -1
    sumatotalalic = totalalic(x) + sumatotalalic
Next x
paso0:
    If Text6.Text = "NCB" Then
        sumatotalalic = sumatotalalic * -1
    End If
    Text8(x - 2).Text = Val(Text9.Text) - sumatotalalic
End If
Rem ********************* factura A ******************
If Text6.Text = "F-A" Or Text6.Text = "NCA" Or Text6.Text = "NDA" Then
    For x = 0 To DataGrid1.VisibleRows - 2
      For Y = 1 To 15
        If IsNull(totalparcial(x, Y)) = True Then totalparcial(x, Y) = 0
        If IsNull(totalalicuota(x, Y)) = True Then totalalicuota(x, Y) = 0
        totalgral = totalalicuota(x, Y) + totalgral
        totalalic(Y) = totalparcial(x, Y) + totalalic(Y)
      Next Y
    Next x
   If Text6.Text = "NCA" Then totalgral = totalgral * -1
   Text9.Text = totalgral
   Masktotal = Text9.Text
For x = 1 To 15
    If Text8(x - 1).Visible = False Then GoTo paso1
    Text8(x - 1).Text = totalalic(x)
    If Text6.Text = "NCA" Then Text8(x - 1).Text = Text8(x - 1).Text * -1
    sumatotalalic = totalalic(x) + sumatotalalic
Next x
paso1:
    If Text6.Text = "NCA" Then
        sumatotalalic = sumatotalalic * -1
    End If
    Text8(x - 2).Text = Val(Text9.Text) - sumatotalalic
End If


End Sub

Private Sub Command4_Click()
On Error Resume Next
Dim tabla As String
Dim ruta As String

ruta = "\Empresa" + Right(Str(login.empresaact), Len(Str(login.empresaact)) - 1)

reporte.SQL = "SELECT facturas.fecha, facturas.cliente, facturas.descripcion, facturas.cuit, facturas.tipocompr, facturas.numcompr, facturas.total, facturas.avisador, facturas.producto, facturas.telefono, facturas.contado, facturas.cant, facturas.unidadmedida, facturas.detalle, facturas.preciounit, facturas.totales, facturas.descuento, facturas.totalneto, facturas.impdesc, facturas.domicilio, facturas.localidad, facturas.numdisco, facturas.empresa FROM contablesql.dbo.facturas facturas WHERE facturas.numcompr = '" & proximafactura & "' and facturas.tipocompr = '" & tipocomprobante & "' and facturas.numdisco = '" & numerodedisco & "' and facturas.empresa = " & login.empresaact & "  ORDER BY facturas.cliente ASC"
tabla = reporte.SQL

With CrystalReporte
    .PrinterCollation = crptCollated
If tipocomprobante = "F-A" Or tipocomprobante = "NCA" Or tipocomprobante = "NDA" Then
    .ReportFileName = App.Path & ruta + "\FacturaA.rpt"
    .WindowTitle = "Factura A"
Else
    .ReportFileName = App.Path & ruta + "\FacturaB.rpt"
    .WindowTitle = "Factura B"
End If
    .Connect = login.conexionreporte
    .DiscardSavedData = True
    .RetrieveDataFiles
    .ReportSource = 0
    .SQLQuery = tabla
 Rem   .Destination = crptToWindow
    .Destination = crptToPrinter
    .PrintFileType = crptCrystal
    .WindowState = crptMaximized
    .Action = 1
    .Action = 1

End With
End Sub

Private Sub confirma_Click()
        If DataGrid1.Columns(16).Text = "" Then DataGrid1.Columns(16).Text = "0"
        DataGrid1.Columns(18).Text = Val(DataGrid1.Columns(9).Text) - Val(DataGrid1.Columns(16).Text) / 100 * Val(DataGrid1.Columns(9).Text)
        If Text6.Text = "NCA" Or Text6.Text = "NCB" Then DataGrid1.Columns(18).Text = DataGrid1.Columns(18).Text * -1
        DataGrid1.Columns(17).Text = Val(DataGrid1.Columns(16).Text) / 100 * Val(DataGrid1.Columns(9).Text)
        
         
End Sub

Private Sub DataCombo1_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        If DataCombo1.Text = "Tarjeta" Then
            Text12.Visible = True
            Text13.Visible = True
            Label12.Visible = True
            Label13.Visible = True
            Text12.SetFocus
        Else
            Text5.SetFocus
        End If
    End If

End Sub

Private Sub DataGrid1_AfterColEdit(ByVal ColIndex As Integer)

 Rem   DataGrid1.Columns(8).Text = DataGrid3.Columns(3).Text
 Rem   If DataGrid1.Col = 8 Or DataGrid1.Col = 5 Then
 Rem       DataGrid1.Columns(9).Text = Val(DataGrid1.Columns(8).Value) * Val(DataGrid1.Columns(5).Value)
 Rem   End If
    If DataGrid1.Col = 4 Then
          DataList2.BoundText = DataGrid1.Columns(4).Text
          If IsNull(DataList2.SelectedItem) = False Then DataGrid3.Bookmark = DataList2.SelectedItem
          DataGrid1.Columns(6).Text = DataGrid3.Columns(1).Text
          DataGrid1.Columns(7).Text = DataGrid3.Columns(2).Text
          DataGrid1.Columns(8).Text = DataGrid3.Columns(3).Text
          If Text6.Text = "F-B" And DataGrid3.Columns(4).Text = "1" Then DataGrid1.Columns(8).Value = Val(DataGrid1.Columns(8).Value) + Val(datparamventas.Recordset.Fields("alicuota1")) / 100 * Val(DataGrid1.Columns(8).Value)
          If Text6.Text = "F-B" And DataGrid3.Columns(4).Text = "2" Then DataGrid1.Columns(8).Value = Val(DataGrid1.Columns(8).Value) + Val(datparamventas.Recordset.Fields("alicuota2")) / 100 * Val(DataGrid1.Columns(8).Value)
          If Text6.Text = "F-B" And DataGrid3.Columns(4).Text = "3" Then DataGrid1.Columns(8).Value = Val(DataGrid1.Columns(8).Value) + Val(datparamventas.Recordset.Fields("alicuota3")) / 100 * Val(DataGrid1.Columns(8).Value)
          KeyAscii = 9
    End If
    If DataGrid1.Col = 16 Or DataGrid1.Col = 9 Then
        If DataGrid1.Columns(16).Text = "" Then DataGrid1.Columns(16).Text = "0"
        DataGrid1.Columns(18) = Val(DataGrid1.Columns(9).Value) - Val(DataGrid1.Columns(16).Text) / 100 * Val(DataGrid1.Columns(9).Value)
        If Text6.Text = "NCA" Or Text6.Text = "NCB" Then DataGrid1.Columns(18).Text = DataGrid1.Columns(18).Text * -1
        DataGrid1.Columns(17) = Val(DataGrid1.Columns(16).Text) / 100 * Val(DataGrid1.Columns(9).Value)
    End If
   
End Sub

Private Sub DataGrid1_AfterUpdate()

    If DataGrid1.Col = 16 Then
         datfacclientes.Recordset.Fields("totalneto") = Val(DataGrid1.Columns(9).Value) - Val(DataGrid1.Columns(16).Text) / 100 * Val(DataGrid1.Columns(9).Value)
         datfacclientes.Recordset.Fields("impdesc") = Val(DataGrid1.Columns(16).Text) / 100 * Val(DataGrid1.Columns(9).Value)
         totalparcial(DataGrid1.Row, collib) = datfacclientes.Recordset.Fields("totalneto")
         If collib = 1 And (Text6.Text = "F-B" Or Text6.Text = "NCB" Or Text6.Text = "NDB") Then datfacclientes.Recordset.Fields("gravado") = datfacclientes.Recordset.Fields("totalneto") / ((100 + datparamventas.Recordset.Fields("alicuota1")) / 100)
         If collib = 2 And (Text6.Text = "F-B" Or Text6.Text = "NCB" Or Text6.Text = "NDB") Then datfacclientes.Recordset.Fields("gravado") = datfacclientes.Recordset.Fields("totalneto") / ((100 + datparamventas.Recordset.Fields("alicuota2")) / 100)
         If collib = 3 And (Text6.Text = "F-B" Or Text6.Text = "NCB" Or Text6.Text = "NDB") Then datfacclientes.Recordset.Fields("gravado") = datfacclientes.Recordset.Fields("totalneto") / ((100 + datparamventas.Recordset.Fields("alicuota3")) / 100)
         If collib = 4 And (Text6.Text = "F-B" Or Text6.Text = "NCB" Or Text6.Text = "NDB") Then datfacclientes.Recordset.Fields("gravado") = datfacclientes.Recordset.Fields("totalneto") / ((100 + datparamventas.Recordset.Fields("alicuota4")) / 100)
         If collib = 1 And (Text6.Text = "F-A" Or Text6.Text = "NCA" Or Text6.Text = "NDA") Then datfacclientes.Recordset.Fields("gravado") = datfacclientes.Recordset.Fields("totalneto") * ((100 + datparamventas.Recordset.Fields("alicuota1")) / 100)
         If collib = 2 And (Text6.Text = "F-A" Or Text6.Text = "NCA" Or Text6.Text = "NDA") Then datfacclientes.Recordset.Fields("gravado") = datfacclientes.Recordset.Fields("totalneto") * ((100 + datparamventas.Recordset.Fields("alicuota2")) / 100)
         If collib = 3 And (Text6.Text = "F-A" Or Text6.Text = "NCA" Or Text6.Text = "NDA") Then datfacclientes.Recordset.Fields("gravado") = datfacclientes.Recordset.Fields("totalneto") * ((100 + datparamventas.Recordset.Fields("alicuota3")) / 100)
         If collib = 4 And (Text6.Text = "F-A" Or Text6.Text = "NCA" Or Text6.Text = "NDA") Then datfacclientes.Recordset.Fields("gravado") = datfacclientes.Recordset.Fields("totalneto") * ((100 + datparamventas.Recordset.Fields("alicuota4")) / 100)
         tipocomprobante = Text6.Text
         cuentaparcial(datfacclientes.Recordset.Fields("collibro")) = datfacclientes.Recordset.Fields("codcuenta")
         centroparcial(datfacclientes.Recordset.Fields("collibro")) = datfacclientes.Recordset.Fields("centrocosto")
         totalalicuota(DataGrid1.Row, collib) = datfacclientes.Recordset.Fields("gravado")
         Call calcula_Click
    End If
      
End Sub

Private Sub DataGrid1_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)

    If ColIndex = 7 Then
        Text18.Visible = True
        Text18.Text = DataGrid1.Columns(7).Text
        KeyAscii = 13
        Text18.Left = DataGrid1.Columns(7).Left + DataGrid1.Left
        Text18.Top = DataGrid1.Top + DataGrid1.RowTop(DataGrid1.Row) + DataGrid1.RowHeight
        Text18.SetFocus
        DataGrid1.Columns(7).Locked = True
    End If
    
    
End Sub

Private Sub DataGrid1_ButtonClick(ByVal ColIndex As Integer)

If ColIndex = 4 Then
        DataList2.Visible = True
        DataList2.Left = DataGrid1.Columns(4).Left + DataGrid1.Left
        DataList2.Width = DataGrid1.Columns(4).Width * 4
        DataList2.Top = DataGrid1.Top + DataGrid1.RowTop(DataGrid1.Row) + DataGrid1.RowHeight * 3
        DataList2.SetFocus
End If

End Sub

Private Sub DataGrid1_GotFocus()

If saltogrid = 1 Then
    DataGrid1.Col = 7
    KeyAscii = 9
    saltogrid = 0
    Exit Sub
End If
        
    DataGrid1.Col = 4

End Sub

Private Sub DataGrid1_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 38 Then KeyCode = 37

End Sub

Private Sub DataGrid1_KeyPress(KeyAscii As Integer)

   If KeyAscii = 13 And (DataGrid1.Col = 8 Or DataGrid1.Col = 5) Then
        KeyAscii = 0
        If IsNull(DataGrid1.Columns(5).Text) = True Then DataGrid1.Columns(5).Text = "1"
        DataGrid1.Columns(9).Value = Val(DataGrid1.Columns(8).Value) * Val(DataGrid1.Columns(5).Value)
        KeyAscii = 9
   End If
    
    If KeyAscii = 13 And DataGrid1.Col = 4 Then
        KeyAscii = 0
        If DataGrid1.Columns(4).Text = "" Then
              DataList2.Visible = True
              DataList2.Left = DataGrid1.Columns(4).Left + DataGrid1.Left
              DataList2.Width = DataGrid1.Columns(4).Width * 4
              DataList2.Top = DataGrid1.Top + DataGrid1.RowTop(DataGrid1.Row) + DataGrid1.RowHeight * 3
              DataList2.SetFocus
              DataGrid1.Columns(5).Text = "1.00"
        Else
              KeyAscii = 9
        End If
    End If
    
   
    If KeyAscii = 13 Then
        KeyAscii = 0
        If DataGrid1.Col = 16 Then
            DataGrid1.Columns(0).Text = s
            DataGrid1.Columns(10).Text = login.empresaact
            DataGrid1.Columns(1).Text = Text6.Text
            Call buscafact_Click
Rem ****** busca columana de libro ********
            collib1 = DataGrid3.Columns(4).Text
            collib2 = datparamventas.Recordset.Fields(0)
            If collib1 = 0 Or IsNull(collib1) = True Or collib1 = "" Then
                If collib2 = 0 Or IsNull(collib2) = True Or collib2 = "" Then
                    mensa = MsgBox("Error de Codificacion en Articulo (Columna Libro)", vbCritical, "Error")
                    Exit Sub
                Else
                    collib = collib2
                End If
            Else
                collib = collib1
            End If
            DataGrid1.Columns(12).Text = collib
Rem ****** busca codigo haber ********
            codhab1 = DataGrid3.Columns(5).Text
            codhab2 = datparamventas.Recordset.Fields(1)
            If codhab1 = 0 Or IsNull(codhab1) = True Or codhab1 = "" Then
                If codhab2 = 0 Or IsNull(codhab2) = True Or codhab2 = "" Then
                    mensa = MsgBox("Error de Codificacion en Articulo (Codigo Contable)", vbCritical, "Error")
                    Exit Sub
                Else
                    codhab = codhab2
                End If
            Else
                codhab = codhab1
            End If
            DataGrid1.Columns(11).Text = codhab
   Rem          cuentaparcial(collib) = codhab
Rem ****** busca codigo centro costo ********
            codcen1 = DataGrid3.Columns(6).Text
            codcen2 = datparamventas.Recordset.Fields(2)
            If codcen1 = 0 Or IsNull(codcen1) = True Or codcen1 = "" Then
                If codcen2 = 0 Or IsNull(codcen2) = True Or codcen2 = "" Then
                    mensa = MsgBox("Error de Codificacion en Articulo (Centro de Costo)", vbCritical, "Error")
                    Exit Sub
                Else
                    codcen = codcen2
                End If
            Else
                codcen = codcen1
            End If
            DataGrid1.Columns(13).Text = codcen
            centroparcial(collib) = codcen
Rem fin busca
            collibro(DataGrid1.Row) = collib
            datfacclientes.Recordset.Fields(14) = "N"
            datfacclientes.Recordset.UpdateBatch adAffectCurrent
            datfacclientes.Recordset.AddNew
            DataGrid1.Col = 4
            Exit Sub
        End If
        KeyAscii = 9
    End If

End Sub


Private Sub DataList1_KeyPress(KeyAscii As Integer)
On Error GoTo fin
    If KeyAscii = 13 Then
        KeyAscii = 0
        Text1.Text = DataList1.Text
        Text2.Text = DataList1.BoundText
        If IsNull(DataList1.SelectedItem) = False Then DataGrid2.Bookmark = DataList1.SelectedItem
        
Rem *********** verifica cod contable del cliente **********
        If datcolumnas.Recordset.Fields("cdt") = 0 And DataGrid2.Columns(12).Text = "" Then
            mensa = MsgBox("Debe Configurar un Codigo Contable para el debe", vbCritical, "!! Error ¡¡")
            Exit Sub
        End If
        If datcolumnas.Recordset.Fields("cdt") <> 0 Then
            codigodebe = datcolumnas.Recordset.Fields("cdt")
        Else
            codigodebe = Val(DataGrid2.Columns(12).Text)
        End If
        If codigodebe = 0 Then
            mensa = MsgBox("Debe Configurar un Codigo Contable para el debe", vbCritical, "!! Error ¡¡")
            Exit Sub
        End If
Rem *********** fin verifica cod contable del cliente **********
        
        Text3.Text = DataGrid2.Columns(4).Text
        Text4.Text = DataGrid2.Columns(5).Text
        Text17.Text = DataGrid2.Columns(7).Text
        If Text3.Text = "RI" Then
            Text6.Text = "F-A"
            tapa.Visible = False
        Else
            Text6.Text = "F-B"
            tapa.Visible = True
        End If
        
        Call buscafact_Click
        If Text1.Text = "CONSUMIDOR FINAL" Then
            Text1.SetFocus
            Text1.SelLength = Len(Text1.Text)
            DataCombo1.Text = "Contado"
            Exit Sub
        End If
        DataCombo1.Text = "Contado"
        Text14.SetFocus
    End If
fin:
End Sub


Private Sub DataList1_LostFocus()

    DataList1.Visible = False

End Sub



Private Sub DataList2_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        DataGrid1.Columns(4).Text = DataList2.BoundText
        If IsNull(DataList2.SelectedItem) = False Then DataGrid3.Bookmark = DataList2.SelectedItem
        DataGrid1.Columns(5).Text = "1"
        DataGrid1.Columns(6).Text = DataGrid3.Columns(1).Text
        DataGrid1.Columns(7).Text = DataGrid3.Columns(2).Text
        DataGrid1.Columns(8).Text = DataGrid3.Columns(3).Text
        If Text6.Text = "F-B" And DataGrid3.Columns(4).Text = "1" Then DataGrid1.Columns(8).Value = Val(DataGrid1.Columns(8).Value) + Val(datparamventas.Recordset.Fields("alicuota1")) / 100 * Val(DataGrid1.Columns(8).Value)
        If Text6.Text = "F-B" And DataGrid3.Columns(4).Text = "2" Then DataGrid1.Columns(8).Value = Val(DataGrid1.Columns(8).Value) + Val(datparamventas.Recordset.Fields("alicuota2")) / 100 * Val(DataGrid1.Columns(8).Value)
        If Text6.Text = "F-B" And DataGrid3.Columns(4).Text = "3" Then DataGrid1.Columns(8).Value = Val(DataGrid1.Columns(8).Value) + Val(datparamventas.Recordset.Fields("alicuota3")) / 100 * Val(DataGrid1.Columns(8).Value)
        DataGrid1.SetFocus
    End If

End Sub

Private Sub DataList2_LostFocus()

    DataList2.Visible = False

End Sub

Private Sub DataList3_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        If IsNull(DataList3.SelectedItem) <> False Then Text3.Text = DataList3.BoundText
        Text4.SetFocus
    End If
    
End Sub

Private Sub DataList3_LostFocus()

    DataList3.Visible = False

End Sub

Private Sub facturar_Click()

If DataCombo1.Text = "Contado" Or DataCombo1.Text = "Tarjeta" Then codigodebe = datcolumnas.Recordset.Fields("fc")

If Val(Masktotal.Text) > 0 And Left(Text6.Text, 2) = "NC" Then
    MsgBox "Error de Imputacion, cierre este ventana y selecciona primero Nota de Credito antes de cargar importes", vbCritical, "Error"
    Exit Sub
End If


Call proximafact_Click
    datfacclientes.RecordSource = "select facclientes.* from facclientes where empresa = " & login.empresaact & " and facturado = 'N' and numdisco = " & s & ""
    datfacclientes.Refresh
    If datfacclientes.Recordset.EOF = True Then Exit Sub
    datfacclientes.Recordset.MoveFirst
paso00:
    datfacclientes.Recordset.Fields(2) = proximafactura
    datfacclientes.Recordset.Fields(14) = "S"
    datfacclientes.Recordset.MoveNext
    If datfacclientes.Recordset.EOF = True Then GoTo paso01
    GoTo paso00
paso01:

Rem ***** graba libreo ventas **************
    datprimaryrs.Recordset.AddNew
    datprimaryrs.Recordset.Fields(1) = login.empresaact
    datprimaryrs.Recordset.Fields(2) = Maskfecha.Text
    datprimaryrs.Recordset.Fields(3) = Text1.Text
    datprimaryrs.Recordset.Fields("domicilio") = Text2.Text
    datprimaryrs.Recordset.Fields("localidad") = Text17.Text
    datprimaryrs.Recordset.Fields(4) = Text3.Text
    datprimaryrs.Recordset.Fields(5) = Text4.Text
    datprimaryrs.Recordset.Fields(6) = Text6.Text
    datprimaryrs.Recordset.Fields(7) = proximafactura
    datprimaryrs.Recordset.Fields("nombretarjeta") = Text12.Text
    datprimaryrs.Recordset.Fields("codoperacion") = Text13.Text
    datprimaryrs.Recordset.Fields("numordenpub") = Text5.Text
    datprimaryrs.Recordset.Fields("avisador") = Text14.Text
    datprimaryrs.Recordset.Fields("producto") = Text15.Text
    datprimaryrs.Recordset.Fields("telefono") = Text16.Text
For x = 1 To 7
    If Text8(x - 1).Visible = False Then GoTo paso1
    posicioniva = x * 2
    Text10(x - 1).Text = Val(Text8(x - 1).Text)
    Text11(x - 1).Text = cuentaparcial(x)
Next x
paso1:
    datprimaryrs.Recordset.Fields("Total") = Val(Text9.Text)
    datprimaryrs.Recordset.Fields("asentado") = "S"
    datprimaryrs.Recordset.Fields("inicioper") = login.iper
    datprimaryrs.Recordset.Fields("finper") = login.fper
    datprimaryrs.Recordset.Fields("cdt") = codigodebe
    datprimaryrs.Recordset.Fields("cerrado") = "N"
    datprimaryrs.Recordset.Fields(posicioniva + 25) = datcolumnas.Recordset.Fields(posicioniva + 30)
    datprimaryrs.Recordset.Fields("ccosto") = centroparcial(1)
    If DataCombo1.Text = "Contado" Or DataCombo1.Text = "Tarjeta" Then datprimaryrs.Recordset.Fields("contado") = "S"
    Text9.Text = Format(Text9.Text, "###,###,##0.00")
    n = Int(Text9.Text)
    d = Right(Text9.Text, 2)
    datprimaryrs.Recordset.Fields("numletras") = numletras0(n, 1) + " Con:" + d + "/100"
    datprimaryrs.Recordset.UpdateBatch adAffectCurrent
    datprimaryrs.RecordSource = "select libroventas.* from libroventas where empresa = " & login.empresaact & " and inicioper = '" & login.iper & "' order by id"
    datprimaryrs.Refresh
    datprimaryrs.Recordset.MoveLast

Rem ****************** grabar asiento

datmaestro.RecordSource = "select [Maestro Asientos].* from [Maestro Asientos] where empresa = " & login.empresaact & " and perinicial = '" & login.iper & "' order by nroasiento"
datmaestro.Refresh
datperiodo.RecordSource = "select EMPRESA.* from EMPRESA where empresa = " & login.empresaact & ""
datperiodo.Refresh
datasiento.RecordSource = "select [Detalle Asientos].* from [Detalle Asientos] where empresa = " & login.empresaact & ""
datasiento.Refresh
    campoaño = Right(Maskfecha.Text, 4)
    campomes = Mid(Maskfecha.Text, 4, 2)
    campodia = Left(Maskfecha.Text, 2)
    campofecha = campoaño + "/" + campomes + "/" + campodia

    campoaño1 = Right(datperiodo.Recordset.Fields(8), 4)
    campomes1 = Mid(datperiodo.Recordset.Fields(8), 4, 2)
    campodia1 = Left(datperiodo.Recordset.Fields(8), 2)
    campofecha1 = campoaño1 + "/" + campomes1 + "/" + campodia1
    
    campoaño2 = Right(datperiodo.Recordset.Fields(9), 4)
    campomes2 = Mid(datperiodo.Recordset.Fields(9), 4, 2)
    campodia2 = Left(datperiodo.Recordset.Fields(9), 2)
    campofecha2 = campoaño2 + "/" + campomes2 + "/" + campodia2

    If campofecha < campofecha1 Or campofecha > campofecha2 Then
            mensa = MsgBox("La Fecha es erronea o no pertenecia al periodo en ejercicio", vbCritical, "!! Atención !!")
            Maskfecha.SelLength = 10
            Maskfecha.SetFocus
            Exit Sub
    End If

    If datmaestro.Recordset.EOF = False Then
            datmaestro.Recordset.MoveLast
            nroasie = datmaestro.Recordset.Fields(3) + 1
    Else
            nroasie = 1
    End If
    
pas1:
    datmaestro.Recordset.AddNew
    datmaestro.Recordset.Fields(0) = Maskfecha.Text
    datmaestro.Recordset.Fields(1) = Date
    datmaestro.Recordset.Fields(3) = nroasie
    datmaestro.Recordset.Fields(4) = Left(Text1.Text, 20) + " " + Text6.Text + " Nº:" + proximafactura
    datmaestro.Recordset.Fields(5) = Str(datperiodo.Recordset.Fields(8))
    datmaestro.Recordset.Fields(6) = Str(datperiodo.Recordset.Fields(9))
    datmaestro.Recordset.Fields(7) = login.empresaact
    datmaestro.Recordset.Fields(8) = "N"
    datmaestro.Recordset.Fields(9) = Val(datprimaryrs.Recordset.Fields(0))
    datmaestro.Recordset.Fields(10) = "V"
    datmaestro.Recordset.Fields(11) = "S"
    datmaestro.Recordset.UpdateBatch adAffectCurrent
     
For x = 0 To 6
    If Val(Text11(x).Text) > 0 Then
            datasiento.Recordset.AddNew
            datasiento.Recordset.Fields(5) = Val(datmaestro.Recordset.Fields(2))
            datasiento.Recordset.Fields(1) = datmaestro.Recordset.Fields(0).Value
            datasiento.Recordset.Fields(7) = login.empresaact
            datasiento.Recordset.Fields(2) = Text11(x).Text
            datasiento.Recordset.Fields(4) = Text10(x).Text
            If datasiento.Recordset.Fields(4) < 0 Then
                datasiento.Recordset.Fields(4) = 0
                datasiento.Recordset.Fields(3) = Text10(x).Text * -1
            End If
            datasiento.Recordset.Fields(6) = Label10(x).Caption
            If (datasiento.Recordset.Fields("ccosto")) > 0 Then datasiento.Recordset.Fields("ccosto") = centroparcial(x + 1)
            datasiento.Recordset.UpdateBatch adAffectCurrent
    End If
Next x
            datasiento.Recordset.AddNew
            datasiento.Recordset.Fields(5) = Val(datmaestro.Recordset.Fields(2))
            datasiento.Recordset.Fields(1) = datmaestro.Recordset.Fields(0).Value
            datasiento.Recordset.Fields(7) = login.empresaact
            datasiento.Recordset.Fields(2) = datprimaryrs.Recordset.Fields("cdt").Value
            datasiento.Recordset.Fields(3) = datprimaryrs.Recordset.Fields("total").Value
            If datasiento.Recordset.Fields(3) < 0 Then
                datasiento.Recordset.Fields(4) = datasiento.Recordset.Fields(3) * -1
                datasiento.Recordset.Fields(3) = 0
            End If
            datasiento.Recordset.Fields(6) = "Total facturado"
            datasiento.Recordset.UpdateBatch adAffectCurrent

    datprimaryrs.Recordset.Fields(59) = nroasie
    datprimaryrs.Recordset.UpdateBatch adAffectCurrent

    Inicio.datauditoria.ConnectionString = login.conexiontotal
    
    Inicio.datauditoria.RecordSource = "select auditoria.* from auditoria"
    Inicio.datauditoria.Refresh
    
    Inicio.datauditoria.Recordset.AddNew
    Inicio.datauditoria.Recordset.Fields("fecha") = Date
    Inicio.datauditoria.Recordset.Fields("hora") = Str(Time)
    Inicio.datauditoria.Recordset.Fields("ventana") = "Facturacion a Clientes"
    Inicio.datauditoria.Recordset.Fields("accion") = "Factura emitida:" + Text6.Text + " " + proximafactura
    Inicio.datauditoria.Recordset.Fields("usuario") = login.usuarioactivo
    Inicio.datauditoria.Recordset.Fields("empresa") = login.empresaact
    Inicio.datauditoria.Recordset.UpdateBatch adAffectCurrent


    Text7.Text = ""
    Text6.Text = ""
    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
    Text4.Text = "00-00000000-0"
    Text5.Text = ""
    DataCombo1.Text = ""
    For x = 0 To 6
        Text8(x).Text = ""
        Label10(x).Caption = ""
    Next x
    Text12 = ""
    Text13 = ""
    Text17 = ""
    Masktotal.Text = "0.00"
    Text12.Visible = False
    Text13.Visible = False
    Label12.Visible = False
    Label13.Visible = False
    
    Call Form_Load
    DataGrid1.Enabled = False
     
    Call Command4_Click
    Maskfecha.SetFocus


End Sub

Private Sub Form_Load()

Rem On Error GoTo errorfactu

    Dim fs, d, t
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set d = fs.GetDrive(fs.GetDriveName(fs.GetAbsolutePathName(drvpath)))
    s = d.serialnumber
    numerodedisco = s
    
    Maskfecha.Text = Date

  frmfacclientesradio.Left = 0
  frmfacclientesradio.Top = 0
  
  
datcondventa.ConnectionString = login.conexiontotal
datasiento.ConnectionString = login.conexiontotal
datclientes.ConnectionString = login.conexiontotal
datcolumnas.ConnectionString = login.conexiontotal
datconsproductos.ConnectionString = login.conexiontotal
datfacclientes.ConnectionString = login.conexiontotal
datmaestro.ConnectionString = login.conexiontotal
datparamventas.ConnectionString = login.conexiontotal
datperiodo.ConnectionString = login.conexiontotal
datprimaryrs.ConnectionString = login.conexiontotal
datproductos.ConnectionString = login.conexiontotal
datempresa.ConnectionString = login.conexiontotal

  datempresa.RecordSource = "select empresa.* from empresa where empresa = " & login.empresaact & ""
  datempresa.Refresh

login.datpuntos.RecordSource = "select puntosdeventas.* from puntosdeventas where empresa = " & login.empresaact & ""
login.datpuntos.Refresh

If login.datpuntos.Recordset.EOF = True Then
        Combo1.AddItem ("0001")
        GoTo conti
End If
login.datpuntos.Recordset.MoveFirst
Do While Not login.datpuntos.Recordset.EOF
        Combo1.AddItem (login.datpuntos.Recordset.Fields(0))
        login.datpuntos.Recordset.MoveNext
Loop
conti:
Combo1.Text = datempresa.Recordset.Fields("PFM")
If login.empresaact1 > 0 Then
    empresareal = login.empresaact1
Else
    empresareal = login.empresaact
End If

    
    DataGrid1.Columns(4).Width = 700
    DataGrid1.Columns(5).Width = 700
    DataGrid1.Columns(6).Width = 700
    DataGrid1.Columns(7).Width = 3000
    DataGrid1.Columns(8).Width = 700
    DataGrid1.Columns(9).Width = 1000
    DataGrid1.Columns(16).Width = 700
    DataGrid1.Columns(17).Width = 800
    DataGrid1.Columns(18).Width = 1000
    
  
    DataGrid1.Columns(5).NumberFormat = "#,###,##0.00"
    DataGrid1.Columns(8).NumberFormat = "#,###,##0.00"
    DataGrid1.Columns(9).NumberFormat = "#,###,##0.00"
    DataGrid1.Columns(16).NumberFormat = "#,###,##0.00"
    DataGrid1.Columns(17).NumberFormat = "#,###,##0.00"
    DataGrid1.Columns(18).NumberFormat = "#,###,##0.00"


    datcondventa.RecordSource = "Select condventas.* from condventas"
    datcondventa.Refresh

    login.datempresa.RecordSource = "select empresa.* from empresa where empresa = " & login.empresaact & ""
    login.datempresa.Refresh

    datconsproductos.RecordSource = "Select consproductos.* from consproductos where empresa = " & login.empresaact & " order by codprod"
    datconsproductos.Refresh
    
    datclientes.RecordSource = "select clientes.* from clientes where clientes.empresa = " & empresareal & " ORDER BY razonsocial"
    datclientes.Refresh
    
    datcolumnas.RecordSource = "SELECT columnasventa.* FROM columnasventa WHERE empresa = " & login.empresaact & " and inicioper = '" & login.iper & "'"
    datcolumnas.Refresh
    
    datproductos.RecordSource = "Select productos.* from productos where empresa = " & login.empresaact & " order by codprod"
    datproductos.Refresh
    datfacclientes.RecordSource = "Select facclientes.* from facclientes where empresa = " & login.empresaact & " order by id"
    datfacclientes.Refresh
    datfacclientes.RecordSource = "select facclientes.* from facclientes where empresa = " & login.empresaact & " and facturado = 'N' and numdisco = " & s & ""
    datfacclientes.Refresh
        
    If datfacclientes.Recordset.EOF = False Then
        datfacclientes.Recordset.MoveFirst
paso0:
        datfacclientes.Recordset.Delete adAffectCurrent
        datfacclientes.Recordset.MoveNext
        If datfacclientes.Recordset.EOF = True Then GoTo paso1
        GoTo paso0
    End If
paso1:
    datfacclientes.RecordSource = "select facclientes.* from facclientes where empresa = " & login.empresaact & " and facturado = 'N' and numdisco = " & s & ""
    datfacclientes.Refresh

    For x = 0 To 6
        valida = IsNull(datcolumnas.Recordset.Fields(x * 2 + 1))
        If valida = False Then
            Label10(x).Caption = datcolumnas.Recordset.Fields(x * 2 + 1)
            Text8(x).Visible = True
        Else
            Text8(x).Visible = False
        End If
        
        
    Next x
    
    datparamventas.RecordSource = "select paramventas.* from paramventas where empresa = " & login.empresaact & ""
    datparamventas.Refresh
    
    For x = 0 To 50
        For Y = 1 To 15
           totalparcial(x, Y) = 0
           totalalicuota(x, Y) = 0
        Next Y
    Next x

  Text1.Text = ""
  Text2.Text = ""
  Text3.Text = ""
  Text5.Text = ""
  Text6.Text = ""
  Text12.Text = ""
  Text13.Text = ""
  Text14.Text = ""
  Text15.Text = ""
  Text16.Text = ""
  saltogrid = 0
  Exit Sub

errorfactu:
    mensa = MsgBox("Faltan configurar algunos parametros relacionados a facturacion", vbCritical, "! Error !")

End Sub

Private Sub Maskfecha_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Text1.SetFocus
    End If
End Sub

Private Sub maskfecha_LostFocus()

 Rem   Maskfecha.Text = Date

End Sub

Private Sub notacredito_Click()

    Text6.Text = "NC" + Right(Text6.Text, 1)

End Sub

Private Sub notadebito_Click()

    Text6.Text = "ND" + Right(Text6.Text, 1)

End Sub

Private Sub numletras_Click()
Dim deci1 As String
Dim uni As String
Dim dec As String
Dim cen As String
Dim unimil As String
Dim decmil As String

Text19.Text = Text9.Text

    If Text9.Text < 0 Then
        Text19.Text = Text9.Text
        Text9.Text = Text9.Text * -1
    End If

    deci = Val(Text9.Text) - Int(Val(Text9.Text))
    entero = Val(Text9.Text) - deci
    carac = Len(entero)
    
    If deci > 0 Then
        For x = 1 To Len(Text9.Text)
            If Mid(Text9.Text, x, 1) = "." Then
                deci1 = Right(Text9.Text, Len(Text9.Text) - x)
                If Len(deci1) = 4 Then
                    If Right(deci1, 2) > 50 Then
                        deci1 = Left(deci1, 2) + 1
                    Else
                        deci1 = Left(deci1, 2)
                    End If
                End If
                If Len(deci1) = 3 Then
                    If Right(deci1, 1) > 5 Then
                        deci1 = Left(deci1, 2) + 1
                    Else
                        deci1 = Left(deci1, 2)
                    End If
                End If
            End If
        Next x
    Else
        deci1 = "00"
    End If
    If Len(deci1) = 1 Then deci1 = deci1 + "0"
    
    If carac = 1 Then GoTo Unidad
    If carac = 2 Then GoTo Decena
    If carac = 3 Then GoTo Centena
    If carac = 4 Then GoTo unidadmil
    If carac = 5 Then GoTo decenamil

Rem ************** Unidad ********************
Unidad:
uni = ""
    Select Case entero
Case 0:
    uni = "Cero "
Case 1:
    uni = "Uno "
Case 2:
    uni = "Dos "
Case 3:
    uni = "Tres "
Case 4:
    uni = "Cuatro "
Case 5:
    uni = "Cinco "
Case 6:
    uni = "Seis "
Case 7:
    uni = "Siete "
Case 8:
    uni = "Ocho "
Case 9:
    uni = "Nueve "
End Select
If flagmil = 1 Then
    unimil = uni
    uni = ""
    entero = entero1
    flagmil = 0
    GoTo Centena
End If
If flagdecmil = 1 Then
    decmil = dec + uni
    dec = ""
    uni = ""
    flagdecmil = 0
    entero = entero1
    GoTo Centena
End If

GoTo final
Rem ************** decena ********************
Decena:
dec = ""
    If Right(entero, 1) = 0 Then
        Select Case entero
        Case 10:
                dec = "Diez "
        Case 20:
                dec = "Veinte "
        Case 30:
                dec = "Treinta "
        Case 40:
                dec = "Cuarenta "
        Case 50:
                dec = "Cincuenta "
        Case 60:
                dec = "Sesenta "
        Case 70:
                dec = "Setenta "
        Case 80:
                dec = "Ochenta "
        Case 90:
                dec = "Noventa "
        End Select
    Else
        Select Case entero
        Case 11:
                dec = "Once "
        Case 12:
                dec = "Doce "
        Case 13:
                dec = "Trece "
        Case 14:
                dec = "Catorce "
        Case 15:
                dec = "Quince "
        Case 16:
                dec = "Dieciseis "
        Case 17:
                dec = "Diecisiete "
        Case 18:
                dec = "Dieciocho "
        Case 19:
                dec = "Diecinueve "
        Case 21:
                dec = "Veintiuno "
        Case 22:
                dec = "Veintidos "
        Case 23:
                dec = "Veintitres "
        Case 24:
                dec = "Veinticuatro "
        Case 25:
                dec = "Veinticinco "
        Case 26:
                dec = "Veintiseis "
        Case 27:
                dec = "Veintisiete "
        Case 28:
                dec = "Veintiocho "
        Case 29:
                dec = "Veintinueve "
        Case 31 To 99
                digitos = Left(entero, 1)
                Select Case digitos
                Case 3:
                    dec = "Treinta y "
                Case 4:
                    dec = "Cuarenta y "
                Case 5:
                    dec = "Cincuenta y "
                Case 6:
                    dec = "Sesenta y "
                Case 7:
                    dec = "Setenta y "
                Case 8:
                    dec = "Ochenta y "
                Case 9:
                    dec = "Noventa y "
                End Select
                entero = Right(entero, 1)
                GoTo Unidad
        End Select
                  
        If flagdecmil = 1 Then
           flagdecmil = 0
           decmil = dec
           dec = ""
           entero = entero1
           GoTo Centena
        End If
    End If
        If Left(entero, 1) = 0 And Right(entero, 1) <> 0 Then
                entero = Right(entero, 1)
                GoTo Unidad
        End If
GoTo final
Rem ************** centena ********************
Centena:
cen = ""
    If Right(entero, 2) = 0 Then
        Select Case entero
        Case 100:
                cen = "Cien "
        Case 200:
                cen = "Doscientos "
        Case 300:
                cen = "Trescientos "
        Case 400:
                cen = "Cuatrocientos "
        Case 500:
                cen = "Cincuenta "
        Case 600:
                cen = "Seiscientos "
        Case 700:
                cen = "Setecientos "
        Case 800:
                cen = "Ochocientos "
        Case 900:
                cen = "Novecientos "
        End Select
    
    Else
        If Left(entero, 1) = 1 Then
            cen = "Ciento "
            entero = Right(entero, 2)
            GoTo Decena
        End If
        Select Case Left(entero, 1)
        Case 1:
                cen = "Cien"
        Case 2:
                cen = "Doscientos "
        Case 3:
                cen = "Trescientos "
        Case 4:
                cen = "Cuatrocientos "
        Case 5:
                cen = "Quinientos "
        Case 6:
                cen = "Seiscientos "
        Case 7:
                cen = "Setecientos "
        Case 8:
                cen = "Ochocientos "
        Case 9:
                cen = "Novecientos "
        End Select
        entero = Right(entero, 2)
        GoTo Decena
    End If
GoTo final

Rem *********** inidad de mil *************
unidadmil:
unimil = ""
flagmil = 0
millar = ""
If Left(entero, 1) = 1 Then
    unimil = "Un "
    millar = "Mil "
    entero = Right(entero, 3)
    GoTo Centena
Else
    millar = "Mil "
    flagmil = 1
    entero1 = Right(entero, 3)
    entero = Left(entero, 1)
    GoTo Unidad
End If
    
    
Rem *********** decena de mil *************
decenamil:
decmil = ""
flagdecmil = 0
    
    millar = "Mil "
    flagdecmil = 1
    entero1 = Right(entero, 3)
    entero = Left(entero, 2)
    GoTo Decena

                                           

final:


letrasnumeros = decmil + unimil + millar + cen + dec + uni + "Con " + deci1 + "/100"
Text9.Text = Text19.Text

End Sub

Private Sub proximafact_Click()
    proximafactura0 = Left(Text7.Text, 4)
    proximafactura1 = Val(Right((Text7.Text), 8)) + 1
    proximafactura = Mid("00000000", 1, 9 - Len(Str(proximafactura1))) + Right(Str(proximafactura1), Len(Str(proximafactura1)) - 1)
    proximafactura = proximafactura0 + "-" + proximafactura
End Sub

Private Sub salir_Click()

    Unload Me

End Sub

Private Sub Text1_GotFocus()


    If Text1.Text <> "CONSUMIDOR FINAL" Then
        DataList1.Visible = True
        DataList1.SetFocus
    End If

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Text2.SetFocus
    End If
End Sub

Private Sub Text12_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        Text13.SetFocus
    End If

End Sub

Private Sub Text13_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Text5.SetFocus
    End If
End Sub

Private Sub Text14_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Text15.SetFocus
    End If
End Sub

Private Sub Text15_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Text16.SetFocus
    End If
End Sub

Private Sub Text16_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        DataCombo1.SetFocus
    End If
End Sub

Private Sub Text17_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Text3.SetFocus
    End If
End Sub

Private Sub Text18_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        DataGrid1.SetFocus
        saltogrid = 1
    End If
End Sub


Private Sub Text18_LostFocus()

        DataGrid1.Columns(7).Text = Text18.Text
        Text18.Visible = False
        DataGrid1.Columns(7).Locked = False
               
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Text17.SetFocus
    End If
End Sub

Private Sub Text3_GotFocus()
    
    DataList3.Visible = True
    DataList3.BoundText = Text3.Text
    DataList3.SetFocus

End Sub

Private Sub text4_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Text14.SetFocus
    End If
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        datfacclientes.Recordset.AddNew
        DataGrid1.Enabled = True
        DataGrid1.SetFocus
    End If
   
End Sub

