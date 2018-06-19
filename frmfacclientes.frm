VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{5734474E-78D3-4254-99B9-C35F31BDF509}#62.0#0"; "vbskpro2.ocx"
Begin VB.Form frmfacclientes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Facturacion a Clientes"
   ClientHeight    =   7815
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10185
   Icon            =   "frmfacclientes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7815
   ScaleWidth      =   10185
   Begin VB.TextBox Text13 
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   6000
      TabIndex        =   6
      Top             =   2880
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox Text12 
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   2160
      TabIndex        =   5
      Top             =   2880
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton facturar 
      Caption         =   "&Facturar"
      Height          =   975
      Left            =   8520
      Picture         =   "frmfacclientes.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   61
      Top             =   3120
      UseMaskColor    =   -1  'True
      Width           =   1335
   End
   Begin VB.TextBox Text11 
      Alignment       =   2  'Center
      DataField       =   "ch7"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   6
      Left            =   9000
      TabIndex        =   60
      Text            =   "Text11"
      Top             =   4920
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
      TabIndex        =   59
      Text            =   "Text11"
      Top             =   4680
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
      TabIndex        =   58
      Text            =   "Text11"
      Top             =   4440
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
      TabIndex        =   57
      Text            =   "Text11"
      Top             =   4200
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
      TabIndex        =   56
      Text            =   "Text11"
      Top             =   3960
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
      TabIndex        =   55
      Text            =   "Text11"
      Top             =   3720
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
      TabIndex        =   54
      Text            =   "Text11"
      Top             =   3480
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
      TabIndex        =   53
      Text            =   "Text10"
      Top             =   4920
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
      TabIndex        =   52
      Text            =   "Text10"
      Top             =   4680
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
      TabIndex        =   51
      Text            =   "Text10"
      Top             =   4440
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
      TabIndex        =   50
      Text            =   "Text10"
      Top             =   4200
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
      TabIndex        =   49
      Text            =   "Text10"
      Top             =   3960
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
      TabIndex        =   48
      Text            =   "Text10"
      Top             =   3720
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox tapa 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   1695
      Left            =   360
      Locked          =   -1  'True
      TabIndex        =   47
      Top             =   5880
      Width           =   4575
   End
   Begin MSMask.MaskEdBox Masktotal 
      Height          =   495
      Left            =   6240
      TabIndex        =   46
      Top             =   6000
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   873
      _Version        =   393216
      BorderStyle     =   0
      Appearance      =   0
      BackColor       =   -2147483644
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
   Begin VB.TextBox Text10 
      Alignment       =   1  'Right Justify
      DataField       =   "col1"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   0
      Left            =   8280
      TabIndex        =   45
      Text            =   "Text10"
      Top             =   3480
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton calcula 
      Caption         =   "calcula"
      Height          =   255
      Left            =   8280
      TabIndex        =   44
      Top             =   3000
      Visible         =   0   'False
      Width           =   1095
   End
   Begin MSDataListLib.DataList DataList3 
      Bindings        =   "frmfacclientes.frx":0884
      Height          =   1035
      Left            =   1920
      TabIndex        =   43
      Top             =   1680
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
      TabIndex        =   42
      Top             =   1680
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
      TabIndex        =   41
      Top             =   6480
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox Text8 
      Alignment       =   1  'Right Justify
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2058
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   6
      Left            =   3360
      TabIndex        =   40
      Top             =   7320
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox Text8 
      Alignment       =   1  'Right Justify
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2058
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   5
      Left            =   3360
      TabIndex        =   39
      Top             =   7080
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox Text8 
      Alignment       =   1  'Right Justify
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2058
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   4
      Left            =   3360
      TabIndex        =   38
      Top             =   6840
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox Text8 
      Alignment       =   1  'Right Justify
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2058
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   3
      Left            =   3360
      TabIndex        =   37
      Top             =   6600
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox Text8 
      Alignment       =   1  'Right Justify
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2058
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   3360
      TabIndex        =   36
      Top             =   6360
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox Text8 
      Alignment       =   1  'Right Justify
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2058
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   3360
      TabIndex        =   35
      Top             =   6120
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox Text8 
      Alignment       =   1  'Right Justify
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#.##0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   3082
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   3360
      TabIndex        =   34
      Top             =   5880
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton proximafact 
      Caption         =   "proximafact"
      Height          =   255
      Left            =   8280
      TabIndex        =   25
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
      Left            =   480
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   360
      Width           =   1815
   End
   Begin VB.CommandButton buscafact 
      Caption         =   "buscafact"
      Height          =   255
      Left            =   8280
      TabIndex        =   22
      Top             =   2520
      Visible         =   0   'False
      Width           =   975
   End
   Begin MSDataListLib.DataList DataList2 
      Bindings        =   "frmfacclientes.frx":0898
      Height          =   1815
      Left            =   600
      TabIndex        =   20
      Top             =   3600
      Visible         =   0   'False
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   3201
      _Version        =   393216
      BackColor       =   14737632
      ListField       =   "lista"
      BoundColumn     =   "codprod"
      Object.DataMember      =   ""
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
      Left            =   3720
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   360
      Width           =   975
   End
   Begin MSDataListLib.DataList DataList1 
      Bindings        =   "frmfacclientes.frx":08B7
      Height          =   1425
      Left            =   5880
      TabIndex        =   16
      Top             =   960
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
   Begin VB.TextBox Text1 
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   1320
      TabIndex        =   2
      Top             =   960
      Width           =   4575
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   1320
      TabIndex        =   1
      Top             =   1320
      Width           =   4575
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Bindings        =   "frmfacclientes.frx":08D1
      Height          =   315
      Left            =   2400
      TabIndex        =   4
      Top             =   2400
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   556
      _Version        =   393216
      BackColor       =   14737632
      ListField       =   "descripcion"
      BoundColumn     =   "codigo"
      Text            =   ""
      Object.DataMember      =   "condventa"
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   1320
      TabIndex        =   3
      Top             =   1680
      Width           =   615
   End
   Begin MSMask.MaskEdBox maskfecha 
      Height          =   285
      Left            =   6720
      TabIndex        =   0
      Top             =   360
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
      Bindings        =   "frmfacclientes.frx":08E5
      Height          =   375
      Left            =   8280
      TabIndex        =   17
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
      Left            =   5880
      TabIndex        =   7
      Top             =   2400
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
      Connect         =   "Provider=MSDASQL.1;Password=25072004;Persist Security Info=True;User ID=lucva;Data Source=contable;Initial Catalog=contablesql"
      OLEDBString     =   "Provider=MSDASQL.1;Password=25072004;Persist Security Info=True;User ID=lucva;Data Source=contable;Initial Catalog=contablesql"
      OLEDBFile       =   ""
      DataSourceName  =   "contable"
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select consproductos.* from consproductos Order by codprod"
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
      Connect         =   "Provider=MSDASQL.1;Password=25072004;Persist Security Info=True;User ID=lucva;Data Source=contable;Initial Catalog=contablesql"
      OLEDBString     =   "Provider=MSDASQL.1;Password=25072004;Persist Security Info=True;User ID=lucva;Data Source=contable;Initial Catalog=contablesql"
      OLEDBFile       =   ""
      DataSourceName  =   "contable"
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select productos.* from productos Order by codprod"
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
      Bindings        =   "frmfacclientes.frx":08FF
      Height          =   2295
      Left            =   240
      TabIndex        =   8
      Top             =   3360
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   4048
      _Version        =   393216
      AllowUpdate     =   -1  'True
      AllowArrows     =   -1  'True
      BackColor       =   12648447
      Enabled         =   0   'False
      ColumnHeaders   =   -1  'True
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
      ColumnCount     =   14
      BeginProperty Column00 
         DataField       =   "numdisco"
         Caption         =   "numdisco"
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
         DataField       =   "tipocomp"
         Caption         =   "tipocomp"
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
         DataField       =   "comprobante"
         Caption         =   "comprobante"
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
         DataField       =   "id"
         Caption         =   "id"
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
         DataField       =   "codproducto"
         Caption         =   "Cod.Prod."
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
         DataField       =   "cant"
         Caption         =   "Cant."
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "unidadmedida"
         Caption         =   "U.med."
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
         DataField       =   "detalle"
         Caption         =   "DETALLE"
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
         DataField       =   "preciounit"
         Caption         =   "P.Unit."
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column09 
         DataField       =   "totales"
         Caption         =   "TOTALES"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   0
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
            LCID            =   2058
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
            LCID            =   2058
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
            LCID            =   2058
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
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            Object.Visible         =   0   'False
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
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DataGrid3 
      Bindings        =   "frmfacclientes.frx":091C
      Height          =   855
      Left            =   5160
      TabIndex        =   21
      Top             =   4680
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   1508
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
      Connect         =   "Provider=MSDASQL.1;Password=25072004;Persist Security Info=True;User ID=lucva;Data Source=contable;Initial Catalog=contablesql"
      OLEDBString     =   "Provider=MSDASQL.1;Password=25072004;Persist Security Info=True;User ID=lucva;Data Source=contable;Initial Catalog=contablesql"
      OLEDBFile       =   ""
      DataSourceName  =   "contable"
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select facclientes.* from facclientes"
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
      Top             =   7485
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
      Connect         =   "Provider=MSDASQL.1;Password=25072004;Persist Security Info=True;User ID=lucva;Data Source=contable;Initial Catalog=contablesql"
      OLEDBString     =   "Provider=MSDASQL.1;Password=25072004;Persist Security Info=True;User ID=lucva;Data Source=contable;Initial Catalog=contablesql"
      OLEDBFile       =   ""
      DataSourceName  =   "contable"
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select libroventas.* from libroventas Order by fecha"
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
      Connect         =   "Provider=MSDASQL.1;Password=25072004;Persist Security Info=True;User ID=lucva;Data Source=contable;Initial Catalog=contablesql"
      OLEDBString     =   "Provider=MSDASQL.1;Password=25072004;Persist Security Info=True;User ID=lucva;Data Source=contable;Initial Catalog=contablesql"
      OLEDBFile       =   ""
      DataSourceName  =   "contable"
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "SELECT clientes.* FROM clientes"
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
      Connect         =   "Provider=MSDASQL.1;Password=25072004;Persist Security Info=True;User ID=lucva;Data Source=contable;Initial Catalog=contablesql"
      OLEDBString     =   "Provider=MSDASQL.1;Password=25072004;Persist Security Info=True;User ID=lucva;Data Source=contable;Initial Catalog=contablesql"
      OLEDBFile       =   ""
      DataSourceName  =   "contable"
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   $"frmfacclientes.frx":0937
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
      Connect         =   "Provider=MSDASQL.1;Password=25072004;Persist Security Info=True;User ID=lucva;Data Source=contable;Initial Catalog=contablesql"
      OLEDBString     =   "Provider=MSDASQL.1;Password=25072004;Persist Security Info=True;User ID=lucva;Data Source=contable;Initial Catalog=contablesql"
      OLEDBFile       =   ""
      DataSourceName  =   "contable"
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "SELECT paramventas.* from paramventas"
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
      Connect         =   "Provider=MSDASQL.1;Password=25072004;Persist Security Info=True;User ID=lucva;Data Source=contable;Initial Catalog=contablesql"
      OLEDBString     =   "Provider=MSDASQL.1;Password=25072004;Persist Security Info=True;User ID=lucva;Data Source=contable;Initial Catalog=contablesql"
      OLEDBFile       =   ""
      DataSourceName  =   "contable"
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select [Maestro Asientos].* from [Maestro Asientos]"
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
      Connect         =   "Provider=MSDASQL.1;Password=25072004;Persist Security Info=True;User ID=lucva;Data Source=contable;Initial Catalog=contablesql"
      OLEDBString     =   "Provider=MSDASQL.1;Password=25072004;Persist Security Info=True;User ID=lucva;Data Source=contable;Initial Catalog=contablesql"
      OLEDBFile       =   ""
      DataSourceName  =   "contable"
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select EMPRESA.* from EMPRESA"
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
      Left            =   9000
      Top             =   480
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
      Connect         =   "Provider=MSDASQL.1;Password=25072004;Persist Security Info=True;User ID=lucva;Data Source=contable;Initial Catalog=contablesql"
      OLEDBString     =   "Provider=MSDASQL.1;Password=25072004;Persist Security Info=True;User ID=lucva;Data Source=contable;Initial Catalog=contablesql"
      OLEDBFile       =   ""
      DataSourceName  =   "contable"
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select [Detalle Asientos].* from [Detalle Asientos]"
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
   Begin vbskpro.Skinner Skinner1 
      Left            =   0
      Top             =   0
      _ExtentX        =   1270
      _ExtentY        =   1270
      CloseButtonToolTipText=   "Cerrar"
      MinButtonToolTipText=   "Minimizar"
      MaxButtonToolTipText=   "Maximizar"
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
      TabIndex        =   63
      Top             =   2880
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
      TabIndex        =   62
      Top             =   2880
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
      BackStyle       =   0  'Transparent
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
      TabIndex        =   33
      Top             =   6000
      Width           =   735
   End
   Begin VB.Label Label10 
      Height          =   225
      Index           =   6
      Left            =   480
      TabIndex        =   32
      Top             =   7320
      Width           =   2775
   End
   Begin VB.Label Label10 
      Height          =   225
      Index           =   5
      Left            =   480
      TabIndex        =   31
      Top             =   7080
      Width           =   2775
   End
   Begin VB.Label Label10 
      Height          =   225
      Index           =   1
      Left            =   480
      TabIndex        =   30
      Top             =   6120
      Width           =   2775
   End
   Begin VB.Label Label10 
      Height          =   225
      Index           =   4
      Left            =   480
      TabIndex        =   29
      Top             =   6840
      Width           =   2775
   End
   Begin VB.Label Label10 
      Height          =   225
      Index           =   3
      Left            =   480
      TabIndex        =   28
      Top             =   6600
      Width           =   2775
   End
   Begin VB.Label Label10 
      Height          =   225
      Index           =   2
      Left            =   480
      TabIndex        =   27
      Top             =   6360
      Width           =   2775
   End
   Begin VB.Label Label10 
      Height          =   225
      Index           =   0
      Left            =   480
      TabIndex        =   26
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
      Left            =   240
      TabIndex        =   24
      Top             =   120
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
      Left            =   3600
      TabIndex        =   19
      Top             =   120
      Width           =   1215
   End
   Begin VB.Shape Shape2 
      Height          =   1335
      Left            =   240
      Top             =   840
      Width           =   8055
   End
   Begin VB.Shape Shape1 
      Height          =   975
      Left            =   240
      Top             =   2280
      Width           =   8055
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Remito Nº:"
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
      Left            =   4800
      TabIndex        =   15
      Top             =   2400
      Width           =   1335
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
      TabIndex        =   14
      Top             =   2400
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
      TabIndex        =   13
      Top             =   960
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
      TabIndex        =   12
      Top             =   1320
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
      TabIndex        =   11
      Top             =   1680
      Width           =   615
   End
   Begin VB.Label Label2 
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
      TabIndex        =   10
      Top             =   1680
      Width           =   735
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
      TabIndex        =   9
      Top             =   360
      Width           =   735
   End
End
Attribute VB_Name = "frmfacclientes"
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

Private Sub buscafact_Click()

 
 If Text6.Text = "F-A" Then
  datPrimaryRS.RecordSource = "select libroventas.* from libroventas WHERE libroventas.empresa = " & login.empresaact & " and tipocompr = 'F-A' Order by id"
  datPrimaryRS.Refresh
 End If
 If Text6.Text = "F-B" Then
  datPrimaryRS.RecordSource = "select libroventas.* from libroventas WHERE libroventas.empresa = " & login.empresaact & " and tipocompr = 'F-B' Order by id"
  datPrimaryRS.Refresh
 End If
 If Text6.Text = "NCA" Then
  datPrimaryRS.RecordSource = "select libroventas.* from libroventas WHERE libroventas.empresa = " & login.empresaact & " and tipocompr = 'NCA' Order by id"
  datPrimaryRS.Refresh
 End If
 If Text6.Text = "NCB" Then
  datPrimaryRS.RecordSource = "select libroventas.* from libroventas WHERE libroventas.empresa = " & login.empresaact & " and tipocompr = 'NCB' Order by id"
  datPrimaryRS.Refresh
 End If
 If Text6.Text = "NDA" Then
  datPrimaryRS.RecordSource = "select libroventas.* from libroventas WHERE libroventas.empresa = " & login.empresaact & " and tipocompr = 'NDA' Order by id"
  datPrimaryRS.Refresh
 End If
 If Text6.Text = "NDB" Then
  datPrimaryRS.RecordSource = "select libroventas.* from libroventas WHERE libroventas.empresa = " & login.empresaact & " and tipocompr = 'NCB' Order by id"
  datPrimaryRS.Refresh
 End If
  
  
  If datPrimaryRS.Recordset.EOF = True Then
    Text7.Text = "0001-00000000"
    Exit Sub
  End If
  datPrimaryRS.Recordset.MoveLast
  Text7.Text = datPrimaryRS.Recordset.Fields(7)

End Sub



Private Sub calcula_Click()
Dim totalgral As Double
Dim sumatotallic As Double

totalgral = 0
sumatotalalic = 0
For x = 1 To 15
    totalalic(x) = 0
Next x
Rem ********************* factura B ******************
Text9.Text = 0
If Text6.Text = "F-B" Then
    For x = 0 To DataGrid1.VisibleRows - 2
      For y = 1 To 15
        If IsNull(totalparcial(x, y)) = True Then totalparcial(x, y) = 0
        If IsNull(totalalicuota(x, y)) = True Then totalalicuota(x, y) = 0
        totalgral = totalparcial(x, y) + totalgral
        totalalic(y) = totalalicuota(x, y) + totalalic(y)
      Next y
    Next x
    Text9.Text = totalgral
    Masktotal = Text9.Text
For x = 1 To 15
    If Text8(x - 1).Visible = False Then GoTo paso0
    Text8(x - 1).Text = totalalic(x)
    sumatotalalic = totalalic(x) + sumatotalalic
Next x
paso0:
    Text8(x - 2).Text = Val(Text9.Text) - sumatotalalic
   
End If
Rem ********************* factura A ******************
If Text6.Text = "F-A" Then
    For x = 0 To DataGrid1.VisibleRows - 2
      For y = 1 To 15
        If IsNull(totalparcial(x, y)) = True Then totalparcial(x, y) = 0
        If IsNull(totalalicuota(x, y)) = True Then totalalicuota(x, y) = 0
        totalgral = totalalicuota(x, y) + totalgral
        totalalic(y) = totalparcial(x, y) + totalalic(y)
      Next y
    Next x
   Text9.Text = totalgral
   Masktotal = Text9.Text
For x = 1 To 15
    If Text8(x - 1).Visible = False Then GoTo paso1
    Text8(x - 1).Text = totalalic(x)
    sumatotalalic = totalalic(x) + sumatotalalic
Next x
paso1:
    Text8(x - 2).Text = Val(Text9.Text) - sumatotalalic
End If

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
    
    If DataGrid1.Col = 8 Or DataGrid1.Col = 5 Then
        DataGrid1.Columns(9).Text = Val(DataGrid1.Columns(8).Text) * Val(DataGrid1.Columns(5).Text)
   Rem     cuentaparcial(DataGrid1.Row) = codhab
   Rem     centroparcial(DataGrid1.Row) = codcen
    End If
    If DataGrid1.Col = 4 Then
          DataList2.BoundText = DataGrid1.Columns(4).Text
          If IsNull(DataList2.SelectedItem) = False Then DataGrid3.Bookmark = DataList2.SelectedItem
          DataGrid1.Columns(6).Text = DataGrid3.Columns(1).Text
          DataGrid1.Columns(7).Text = DataGrid3.Columns(2).Text
          DataGrid1.Columns(8).Text = DataGrid3.Columns(3).Text
          KeyAscii = 9
    End If
    
End Sub

Private Sub DataGrid1_AfterUpdate()

    If DataGrid1.Col = 9 Then
         totalparcial(DataGrid1.Row, collib) = datfacclientes.Recordset.Fields("totales")
         If collib = 1 And Text6.Text = "F-B" Then datfacclientes.Recordset.Fields("gravado") = datfacclientes.Recordset.Fields("totales") / ((100 + datparamventas.Recordset.Fields("alicuota1")) / 100)
         If collib = 2 And Text6.Text = "F-B" Then datfacclientes.Recordset.Fields("gravado") = datfacclientes.Recordset.Fields("totales") / ((100 + datparamventas.Recordset.Fields("alicuota2")) / 100)
         If collib = 3 And Text6.Text = "F-B" Then datfacclientes.Recordset.Fields("gravado") = datfacclientes.Recordset.Fields("totales") / ((100 + datparamventas.Recordset.Fields("alicuota3")) / 100)
         If collib = 4 And Text6.Text = "F-B" Then datfacclientes.Recordset.Fields("gravado") = datfacclientes.Recordset.Fields("totales") / ((100 + datparamventas.Recordset.Fields("alicuota4")) / 100)
         If collib = 1 And Text6.Text = "F-A" Then datfacclientes.Recordset.Fields("gravado") = datfacclientes.Recordset.Fields("totales") * ((100 + datparamventas.Recordset.Fields("alicuota1")) / 100)
         If collib = 2 And Text6.Text = "F-A" Then datfacclientes.Recordset.Fields("gravado") = datfacclientes.Recordset.Fields("totales") * ((100 + datparamventas.Recordset.Fields("alicuota2")) / 100)
         If collib = 3 And Text6.Text = "F-A" Then datfacclientes.Recordset.Fields("gravado") = datfacclientes.Recordset.Fields("totales") * ((100 + datparamventas.Recordset.Fields("alicuota3")) / 100)
         If collib = 4 And Text6.Text = "F-A" Then datfacclientes.Recordset.Fields("gravado") = datfacclientes.Recordset.Fields("totales") * ((100 + datparamventas.Recordset.Fields("alicuota4")) / 100)
            
         cuentaparcial(datfacclientes.Recordset.Fields("collibro")) = datfacclientes.Recordset.Fields("codcuenta")
         centroparcial(datfacclientes.Recordset.Fields("collibro")) = datfacclientes.Recordset.Fields("centrocosto")
    Rem     Debug.Print cuentaparcial(datfacclientes.Recordset.Fields("collibro")), "-"; datfacclientes.Recordset.Fields("collibro"), centroparcial(datfacclientes.Recordset.Fields("collibro"))
         totalalicuota(DataGrid1.Row, collib) = datfacclientes.Recordset.Fields("gravado")
         Call calcula_Click
    End If


End Sub

Private Sub DataGrid1_ButtonClick(ByVal ColIndex As Integer)

If ColIndex = 4 Then
        DataList2.Visible = True
        DataList2.Left = DataGrid1.Columns(4).Left + DataGrid1.Left
        DataList2.Width = DataGrid1.Columns(4).Width * 4
        DataList2.Top = DataGrid1.Top + DataGrid1.RowTop(DataGrid1.Row) + DataGrid1.RowHeight
        DataList2.SetFocus
End If

End Sub


Private Sub DataGrid1_GotFocus()

    DataGrid1.Col = 4

End Sub

Private Sub DataGrid1_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 And DataGrid1.Col = 4 Then
        KeyAscii = 0
        If DataGrid1.Columns(4).Text = "" Then
              DataList2.Visible = True
              DataList2.Left = DataGrid1.Columns(4).Left + DataGrid1.Left
              DataList2.Width = DataGrid1.Columns(4).Width * 4
              DataList2.Top = DataGrid1.Top + DataGrid1.RowTop(DataGrid1.Row) + DataGrid1.RowHeight
              DataList2.SetFocus
              DataGrid1.Columns(5).Text = "1.00"
        Else
              KeyAscii = 9
        End If
    End If
    
   
    If KeyAscii = 13 Then
        KeyAscii = 0
        If DataGrid1.Col = 9 Then
            DataGrid1.Columns(0).Text = s
            DataGrid1.Columns(10).Text = login.empresaact
            DataGrid1.Columns(1).Text = Text6.Text
            Call buscafact_Click
    Rem        Call proximafact_Click
    Rem        DataGrid1.Columns(2).Text = proximafactura
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
            datfacclientes.Recordset.Fields(16) = Text12.Text
            datfacclientes.Recordset.Fields(17) = Text13.Text
            datfacclientes.Recordset.UpdateBatch adAffectCurrent
            datfacclientes.Recordset.AddNew
            DataGrid1.Col = 4
            Exit Sub
        End If
        KeyAscii = 9
    End If
    
      
End Sub


Private Sub DataList1_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        text1.Text = DataList1.Text
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
        If Text3.Text = "RI" Then
            Text6.Text = "F-A"
            tapa.Visible = False
        Else
            Text6.Text = "F-B"
            tapa.Visible = True
        End If
        
        Call buscafact_Click
        If text1.Text = "CONSUMIDOR FINAL" Then
            text1.SetFocus
            text1.SelLength = Len(text1.Text)
            DataCombo1.Text = "Contado"
            Exit Sub
        End If
        DataCombo1.Text = "Contado"
        DataCombo1.SetFocus
    End If

End Sub


Private Sub DataList1_LostFocus()

    DataList1.Visible = False

End Sub


Private Sub DataList2_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        DataGrid1.Columns(4).Text = DataList2.BoundText
        If IsNull(DataList2.SelectedItem) = False Then DataGrid3.Bookmark = DataList2.SelectedItem
        DataGrid1.Columns(6).Text = DataGrid3.Columns(1).Text
        DataGrid1.Columns(7).Text = DataGrid3.Columns(2).Text
        DataGrid1.Columns(8).Text = DataGrid3.Columns(3).Text
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
    datPrimaryRS.Recordset.AddNew
    datPrimaryRS.Recordset.Fields(1) = login.empresaact
    datPrimaryRS.Recordset.Fields(2) = Maskfecha
    datPrimaryRS.Recordset.Fields(3) = text1.Text
    datPrimaryRS.Recordset.Fields(4) = Text3.Text
    datPrimaryRS.Recordset.Fields(5) = Text4.Text
    datPrimaryRS.Recordset.Fields(6) = Text6.Text
    datPrimaryRS.Recordset.Fields(7) = proximafactura
For x = 1 To 7
    If Text8(x - 1).Visible = False Then GoTo paso1
    posicioniva = x * 2
    Text10(x - 1).Text = Val(Text8(x - 1).Text)
    Text11(x - 1).Text = cuentaparcial(x)
Next x
paso1:
    datPrimaryRS.Recordset.Fields("Total") = Val(Text9.Text)
    datPrimaryRS.Recordset.Fields("asentado") = "S"
    datPrimaryRS.Recordset.Fields("inicioper") = login.iper
    datPrimaryRS.Recordset.Fields("finper") = login.fper
    datPrimaryRS.Recordset.Fields("cdt") = codigodebe
    datPrimaryRS.Recordset.Fields("cerrado") = "N"
    datPrimaryRS.Recordset.Fields(posicioniva + 25) = datcolumnas.Recordset.Fields(posicioniva + 30)
    datPrimaryRS.Recordset.Fields("ccosto") = centroparcial(1)
    If DataCombo1.Text = "Contado" Or DataCombo1.Text = "Tarjeta" Then datPrimaryRS.Recordset.Fields("contado") = "S"
    datPrimaryRS.Recordset.UpdateBatch adAffectCurrent
    datPrimaryRS.RecordSource = "select libroventas.* from libroventas where empresa = " & login.empresaact & " and inicioper = '" & login.iper & "' order by id"
    datPrimaryRS.Refresh
    datPrimaryRS.Recordset.MoveLast

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
    datmaestro.Recordset.Fields(4) = Left(text1.Text, 20) + " " + Text6.Text + " Nº:" + proximafactura
    datmaestro.Recordset.Fields(5) = datperiodo.Recordset.Fields(8)
    datmaestro.Recordset.Fields(6) = datperiodo.Recordset.Fields(9)
    datmaestro.Recordset.Fields(7) = login.empresaact
    datmaestro.Recordset.Fields(8) = "N"
    datmaestro.Recordset.Fields(9) = Val(datPrimaryRS.Recordset.Fields(0))
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
            datasiento.Recordset.Fields(6) = Label10(x).Caption
            If (datasiento.Recordset.Fields("ccosto")) > 0 Then datasiento.Recordset.Fields("ccosto") = centroparcial(x + 1)
            datasiento.Recordset.UpdateBatch adAffectCurrent
    End If
Next x
            datasiento.Recordset.AddNew
            datasiento.Recordset.Fields(5) = Val(datmaestro.Recordset.Fields(2))
            datasiento.Recordset.Fields(1) = datmaestro.Recordset.Fields(0).Value
            datasiento.Recordset.Fields(7) = login.empresaact
            datasiento.Recordset.Fields(2) = datPrimaryRS.Recordset.Fields("cdt").Value
            datasiento.Recordset.Fields(3) = datPrimaryRS.Recordset.Fields("total").Value
            datasiento.Recordset.Fields(6) = "Total facturado"
            datasiento.Recordset.UpdateBatch adAffectCurrent

    datPrimaryRS.Recordset.Fields(59) = nroasie
    datPrimaryRS.Recordset.UpdateBatch adAffectCurrent


    Text7.Text = ""
    Text6.Text = ""
    text1.Text = ""
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
    Text12.Visible = False
    Text13.Visible = False
    Label12.Visible = False
    Label13.Visible = False
    
    Call Form_Load
    DataGrid1.Enabled = False
    Maskfecha.SetFocus


End Sub

Private Sub Form_Load()


    Dim fs, d, t
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set d = fs.GetDrive(fs.GetDriveName(fs.GetAbsolutePathName(drvpath)))
    s = d.serialnumber
    
    Maskfecha.Text = Date

    datconsproductos.RecordSource = "Select consproductos.* from consproductos where empresa = " & login.empresaact & " order by codprod"
    datconsproductos.Refresh
    
    datclientes.RecordSource = "select clientes.* from clientes where clientes.empresa = " & login.empresaact & " ORDER BY razonsocial"
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

    For x = 0 To 7
        valida = IsNull(datcolumnas.Recordset.Fields(x * 2 + 1))
        If valida = False Then
            Label10(x).Caption = datcolumnas.Recordset.Fields(x * 2 + 1)
            Text8(x).Visible = True
        End If
        
    Next x
    
    datparamventas.RecordSource = "select paramventas.* from paramventas where empresa = " & login.empresaact & ""
    datparamventas.Refresh
    
    For x = 0 To 50
        For y = 1 To 15
           totalparcial(x, y) = 0
           totalalicuota(x, y) = 0
        Next y
    Next x

End Sub

Private Sub Maskfecha_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        text1.SetFocus
    End If
End Sub

Private Sub proximafact_Click()
    proximafactura0 = Left(Text7.Text, 4)
    proximafactura1 = Val(Right((Text7.Text), 8)) + 1
    proximafactura = Mid("00000000", 1, 9 - Len(Str(proximafactura1))) + Right(Str(proximafactura1), Len(Str(proximafactura1)) - 1)
    proximafactura = proximafactura0 + "-" + proximafactura
End Sub

Private Sub Text1_GotFocus()


    If text1.Text <> "CONSUMIDOR FINAL" Then
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

Private Sub Text2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Text3.SetFocus
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
        DataCombo1.SetFocus
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
