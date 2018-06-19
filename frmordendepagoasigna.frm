VERSION 5.00
Object = "{F6125AB1-8AB1-11CE-A77F-08002B2F4E98}#2.0#0"; "MSRDC20.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{5734474E-78D3-4254-99B9-C35F31BDF509}#62.0#0"; "vbskpro2.ocx"
Begin VB.Form frmordendepagoasigna 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Asignacion de Pagos sin Comprobantes"
   ClientHeight    =   7560
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11505
   Icon            =   "frmordendepagoasigna.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7560
   ScaleWidth      =   11505
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   255
      Left            =   1560
      TabIndex        =   52
      Top             =   4440
      Visible         =   0   'False
      Width           =   1455
   End
   Begin MSDataListLib.DataList datalist3 
      Bindings        =   "frmordendepagoasigna.frx":0442
      Height          =   2160
      Left            =   5160
      TabIndex        =   40
      Top             =   1080
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   3810
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "nrorden"
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
   Begin MSDataListLib.DataList DataList2 
      Bindings        =   "frmordendepagoasigna.frx":045D
      Height          =   1815
      Left            =   5160
      TabIndex        =   6
      Top             =   1920
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   3201
      _Version        =   393216
      MatchEntry      =   -1  'True
      BackColor       =   12632256
      ListField       =   "comp"
      BoundColumn     =   "id"
   End
   Begin MSDataListLib.DataList DataList1 
      Bindings        =   "frmordendepagoasigna.frx":047B
      Height          =   2400
      Left            =   360
      TabIndex        =   39
      Top             =   1080
      Visible         =   0   'False
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   4233
      _Version        =   393216
      MatchEntry      =   -1  'True
      BackColor       =   12648447
      ListField       =   "razonsocial"
      BoundColumn     =   "codproveedor"
   End
   Begin VB.CommandButton asiglista 
      Caption         =   "Aceptar Asignación de &Lista"
      Height          =   375
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   50
      Top             =   5400
      UseMaskColor    =   -1  'True
      Width           =   3615
   End
   Begin VB.CommandButton limpia 
      Caption         =   "Limpia Items"
      Height          =   495
      Left            =   9480
      Picture         =   "frmordendepagoasigna.frx":0498
      Style           =   1  'Graphical
      TabIndex        =   47
      Top             =   1320
      UseMaskColor    =   -1  'True
      Width           =   1095
   End
   Begin VB.CommandButton asignartodos 
      Height          =   495
      Left            =   9000
      Picture         =   "frmordendepagoasigna.frx":09CA
      Style           =   1  'Graphical
      TabIndex        =   46
      Top             =   1320
      UseMaskColor    =   -1  'True
      Width           =   495
   End
   Begin VB.PictureBox impfactura 
      Height          =   255
      Left            =   7320
      ScaleHeight     =   195
      ScaleWidth      =   1515
      TabIndex        =   43
      Top             =   4560
      Width           =   1575
   End
   Begin VB.ListBox List1 
      Height          =   5460
      Left            =   9000
      Sorted          =   -1  'True
      Style           =   1  'Checkbox
      TabIndex        =   42
      Top             =   1800
      Width           =   2295
   End
   Begin VB.Frame Frame3 
      Caption         =   "Proveedor"
      Height          =   975
      Left            =   360
      TabIndex        =   38
      Top             =   120
      Width           =   1575
      Begin VB.CommandButton Command1 
         Height          =   615
         Left            =   240
         Picture         =   "frmordendepagoasigna.frx":0E0C
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   240
         Width           =   1095
      End
   End
   Begin MSMask.MaskEdBox totalfactura 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2058
         SubFormatType   =   1
      EndProperty
      Height          =   375
      Left            =   3360
      TabIndex        =   36
      Top             =   2400
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      BackColor       =   12648447
      Enabled         =   0   'False
      Format          =   "    #,##0.00;(    #,##0.00)"
      PromptChar      =   "_"
   End
   Begin VB.CommandButton limpiar 
      Caption         =   "&Limpiar"
      Height          =   375
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   3240
      UseMaskColor    =   -1  'True
      Width           =   1695
   End
   Begin VB.PictureBox saldocompro 
      Height          =   255
      Index           =   0
      Left            =   7080
      ScaleHeight     =   195
      ScaleWidth      =   1515
      TabIndex        =   27
      Top             =   1680
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      DataField       =   "comprobante"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   3082
         SubFormatType   =   1
      EndProperty
      DataSource      =   "databonan"
      Height          =   285
      Index           =   10
      Left            =   5160
      TabIndex        =   26
      Top             =   3360
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      DataField       =   "comprobante"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   3082
         SubFormatType   =   1
      EndProperty
      DataSource      =   "databonan"
      Height          =   285
      Index           =   9
      Left            =   5160
      TabIndex        =   25
      Top             =   3120
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      DataField       =   "comprobante"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   3082
         SubFormatType   =   1
      EndProperty
      DataSource      =   "databonan"
      Height          =   285
      Index           =   8
      Left            =   5160
      TabIndex        =   24
      Top             =   2880
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      DataField       =   "comprobante"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   3082
         SubFormatType   =   1
      EndProperty
      DataSource      =   "databonan"
      Height          =   285
      Index           =   7
      Left            =   5160
      TabIndex        =   23
      Top             =   2640
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      DataField       =   "comprobante"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   3082
         SubFormatType   =   1
      EndProperty
      DataSource      =   "databonan"
      Height          =   285
      Index           =   6
      Left            =   5160
      TabIndex        =   22
      Top             =   2400
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      DataField       =   "comprobante"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   3082
         SubFormatType   =   1
      EndProperty
      DataSource      =   "databonan"
      Height          =   285
      Index           =   5
      Left            =   5160
      TabIndex        =   21
      Top             =   2160
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      DataField       =   "comprobante"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   3082
         SubFormatType   =   1
      EndProperty
      DataSource      =   "databonan"
      Height          =   285
      Index           =   4
      Left            =   5160
      TabIndex        =   20
      Top             =   1920
      Width           =   1815
   End
   Begin VB.CommandButton aceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   3240
      UseMaskColor    =   -1  'True
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      DataField       =   "comprobante"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   3082
         SubFormatType   =   1
      EndProperty
      DataSource      =   "databonan"
      Height          =   285
      Index           =   3
      Left            =   5160
      TabIndex        =   18
      Top             =   1680
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      DataField       =   "fechacompro"
      DataSource      =   "databonan"
      Height          =   285
      Index           =   1
      Left            =   3840
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   1680
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      DataField       =   "nomproveedor"
      DataSource      =   "databonan"
      Height          =   285
      Index           =   0
      Left            =   480
      Locked          =   -1  'True
      TabIndex        =   16
      Top             =   1680
      Width           =   3375
   End
   Begin MSDataGridLib.DataGrid DataGrid4 
      Bindings        =   "frmordendepagoasigna.frx":124E
      Height          =   495
      Left            =   1080
      TabIndex        =   8
      Top             =   3480
      Visible         =   0   'False
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   873
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
      ColumnCount     =   65
      BeginProperty Column00 
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
      BeginProperty Column01 
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
      BeginProperty Column02 
         DataField       =   "fecha"
         Caption         =   "fecha"
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
         DataField       =   "proveedor"
         Caption         =   "proveedor"
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
         DataField       =   "tipoiva"
         Caption         =   "tipoiva"
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
         DataField       =   "cuit"
         Caption         =   "cuit"
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
      BeginProperty Column06 
         DataField       =   "tipocompr"
         Caption         =   "tipocompr"
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
         DataField       =   "numcompr"
         Caption         =   "numcompr"
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
         DataField       =   "col1"
         Caption         =   "col1"
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
      BeginProperty Column09 
         DataField       =   "col2"
         Caption         =   "col2"
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
      BeginProperty Column10 
         DataField       =   "col3"
         Caption         =   "col3"
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
         DataField       =   "col4"
         Caption         =   "col4"
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
         DataField       =   "col5"
         Caption         =   "col5"
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
         DataField       =   "col6"
         Caption         =   "col6"
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
         DataField       =   "col7"
         Caption         =   "col7"
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
         DataField       =   "col8"
         Caption         =   "col8"
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
         DataField       =   "col9"
         Caption         =   "col9"
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
      BeginProperty Column17 
         DataField       =   "col10"
         Caption         =   "col10"
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
      BeginProperty Column18 
         DataField       =   "col11"
         Caption         =   "col11"
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
      BeginProperty Column19 
         DataField       =   "col12"
         Caption         =   "col12"
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
      BeginProperty Column20 
         DataField       =   "col13"
         Caption         =   "col13"
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
      BeginProperty Column21 
         DataField       =   "col14"
         Caption         =   "col14"
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
      BeginProperty Column22 
         DataField       =   "col15"
         Caption         =   "col15"
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
      BeginProperty Column23 
         DataField       =   "cuenta"
         Caption         =   "cuenta"
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
      BeginProperty Column24 
         DataField       =   "total"
         Caption         =   "total"
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
      BeginProperty Column25 
         DataField       =   "cerrado"
         Caption         =   "cerrado"
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
      BeginProperty Column26 
         DataField       =   "cd1"
         Caption         =   "cd1"
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
      BeginProperty Column27 
         DataField       =   "ch1"
         Caption         =   "ch1"
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
      BeginProperty Column28 
         DataField       =   "cd2"
         Caption         =   "cd2"
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
      BeginProperty Column29 
         DataField       =   "ch2"
         Caption         =   "ch2"
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
      BeginProperty Column30 
         DataField       =   "cd3"
         Caption         =   "cd3"
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
      BeginProperty Column31 
         DataField       =   "ch3"
         Caption         =   "ch3"
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
      BeginProperty Column32 
         DataField       =   "cd4"
         Caption         =   "cd4"
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
      BeginProperty Column33 
         DataField       =   "ch4"
         Caption         =   "ch4"
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
      BeginProperty Column34 
         DataField       =   "cd5"
         Caption         =   "cd5"
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
      BeginProperty Column35 
         DataField       =   "ch5"
         Caption         =   "ch5"
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
      BeginProperty Column36 
         DataField       =   "cd6"
         Caption         =   "cd6"
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
      BeginProperty Column37 
         DataField       =   "ch6"
         Caption         =   "ch6"
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
      BeginProperty Column38 
         DataField       =   "cd7"
         Caption         =   "cd7"
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
      BeginProperty Column39 
         DataField       =   "ch7"
         Caption         =   "ch7"
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
      BeginProperty Column40 
         DataField       =   "cd8"
         Caption         =   "cd8"
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
      BeginProperty Column41 
         DataField       =   "ch8"
         Caption         =   "ch8"
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
      BeginProperty Column42 
         DataField       =   "cd9"
         Caption         =   "cd9"
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
      BeginProperty Column43 
         DataField       =   "ch9"
         Caption         =   "ch9"
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
      BeginProperty Column44 
         DataField       =   "cd10"
         Caption         =   "cd10"
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
      BeginProperty Column45 
         DataField       =   "ch10"
         Caption         =   "ch10"
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
      BeginProperty Column46 
         DataField       =   "cd11"
         Caption         =   "cd11"
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
      BeginProperty Column47 
         DataField       =   "ch11"
         Caption         =   "ch11"
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
      BeginProperty Column48 
         DataField       =   "cd12"
         Caption         =   "cd12"
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
      BeginProperty Column49 
         DataField       =   "ch12"
         Caption         =   "ch12"
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
      BeginProperty Column50 
         DataField       =   "cd13"
         Caption         =   "cd13"
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
      BeginProperty Column51 
         DataField       =   "ch13"
         Caption         =   "ch13"
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
      BeginProperty Column52 
         DataField       =   "cd14"
         Caption         =   "cd14"
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
      BeginProperty Column53 
         DataField       =   "ch14"
         Caption         =   "ch14"
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
      BeginProperty Column54 
         DataField       =   "cd15"
         Caption         =   "cd15"
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
      BeginProperty Column55 
         DataField       =   "ch15"
         Caption         =   "ch15"
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
      BeginProperty Column56 
         DataField       =   "cdt"
         Caption         =   "cdt"
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
      BeginProperty Column57 
         DataField       =   "cht"
         Caption         =   "cht"
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
      BeginProperty Column58 
         DataField       =   "asentado"
         Caption         =   "asentado"
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
      BeginProperty Column59 
         DataField       =   "asiento"
         Caption         =   "asiento"
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
      BeginProperty Column60 
         DataField       =   "ccosto"
         Caption         =   "ccosto"
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
      BeginProperty Column61 
         DataField       =   "inicioper"
         Caption         =   "inicioper"
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
      BeginProperty Column62 
         DataField       =   "finper"
         Caption         =   "finper"
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
      BeginProperty Column63 
         DataField       =   "saldo"
         Caption         =   "saldo"
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
      BeginProperty Column64 
         DataField       =   "imputado"
         Caption         =   "imputado"
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
         BeginProperty Column13 
         EndProperty
         BeginProperty Column14 
         EndProperty
         BeginProperty Column15 
         EndProperty
         BeginProperty Column16 
         EndProperty
         BeginProperty Column17 
         EndProperty
         BeginProperty Column18 
         EndProperty
         BeginProperty Column19 
         EndProperty
         BeginProperty Column20 
         EndProperty
         BeginProperty Column21 
         EndProperty
         BeginProperty Column22 
         EndProperty
         BeginProperty Column23 
         EndProperty
         BeginProperty Column24 
         EndProperty
         BeginProperty Column25 
         EndProperty
         BeginProperty Column26 
         EndProperty
         BeginProperty Column27 
         EndProperty
         BeginProperty Column28 
         EndProperty
         BeginProperty Column29 
         EndProperty
         BeginProperty Column30 
         EndProperty
         BeginProperty Column31 
         EndProperty
         BeginProperty Column32 
         EndProperty
         BeginProperty Column33 
         EndProperty
         BeginProperty Column34 
         EndProperty
         BeginProperty Column35 
         EndProperty
         BeginProperty Column36 
         EndProperty
         BeginProperty Column37 
         EndProperty
         BeginProperty Column38 
         EndProperty
         BeginProperty Column39 
         EndProperty
         BeginProperty Column40 
         EndProperty
         BeginProperty Column41 
         EndProperty
         BeginProperty Column42 
         EndProperty
         BeginProperty Column43 
         EndProperty
         BeginProperty Column44 
         EndProperty
         BeginProperty Column45 
         EndProperty
         BeginProperty Column46 
         EndProperty
         BeginProperty Column47 
         EndProperty
         BeginProperty Column48 
         EndProperty
         BeginProperty Column49 
         EndProperty
         BeginProperty Column50 
         EndProperty
         BeginProperty Column51 
         EndProperty
         BeginProperty Column52 
         EndProperty
         BeginProperty Column53 
         EndProperty
         BeginProperty Column54 
         EndProperty
         BeginProperty Column55 
         EndProperty
         BeginProperty Column56 
         EndProperty
         BeginProperty Column57 
         EndProperty
         BeginProperty Column58 
         EndProperty
         BeginProperty Column59 
         EndProperty
         BeginProperty Column60 
         EndProperty
         BeginProperty Column61 
         EndProperty
         BeginProperty Column62 
         EndProperty
         BeginProperty Column63 
         EndProperty
         BeginProperty Column64 
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DataGrid3 
      Bindings        =   "frmordendepagoasigna.frx":126C
      Height          =   375
      Left            =   480
      TabIndex        =   7
      Top             =   3120
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      _Version        =   393216
      BackColor       =   12648447
      Enabled         =   -1  'True
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
         DataField       =   "empresa"
         Caption         =   "empresa"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "fecha"
         Caption         =   "fecha"
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
         DataField       =   "proveedor"
         Caption         =   "proveedor"
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
         DataField       =   "inicioper"
         Caption         =   "inicioper"
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
         DataField       =   "finper"
         Caption         =   "finper"
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
         DataField       =   "numcompr"
         Caption         =   "numcompr"
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
      BeginProperty Column06 
         DataField       =   "total"
         Caption         =   "Total Comprob."
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column07 
         DataField       =   "tipocompr"
         Caption         =   "tipocompr"
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
         DataField       =   "Expr1"
         Caption         =   "Expr1"
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
      BeginProperty Column09 
         DataField       =   "saldo"
         Caption         =   "saldo"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column10 
         DataField       =   "cht"
         Caption         =   "cht"
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
      BeginProperty Column12 
         DataField       =   "imputado"
         Caption         =   "imputado"
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
            Alignment       =   1
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
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column05 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column06 
            Alignment       =   1
            Locked          =   -1  'True
         EndProperty
         BeginProperty Column07 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column08 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column09 
            Alignment       =   1
            Locked          =   -1  'True
            Object.Visible         =   -1  'True
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
      EndProperty
   End
   Begin VB.CommandButton salir 
      Caption         =   "&Cerrar"
      Height          =   735
      Left            =   9960
      Picture         =   "frmordendepagoasigna.frx":128A
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   360
      UseMaskColor    =   -1  'True
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   255
      Left            =   7440
      TabIndex        =   14
      Top             =   120
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton nuevaorden 
      Caption         =   "&Nueva Asig."
      Height          =   735
      Left            =   8520
      Picture         =   "frmordendepagoasigna.frx":16CC
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   360
      UseMaskColor    =   -1  'True
      Width           =   1335
   End
   Begin VB.CommandButton cancelar 
      Cancel          =   -1  'True
      Caption         =   "Command4"
      Height          =   375
      Left            =   3600
      TabIndex        =   12
      Top             =   2760
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Asignar Orden"
      Enabled         =   0   'False
      Height          =   735
      Left            =   7080
      Picture         =   "frmordendepagoasigna.frx":1B0E
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   360
      UseMaskColor    =   -1  'True
      Width           =   1335
   End
   Begin MSMask.MaskEdBox totalabonan 
      Height          =   255
      Left            =   1560
      TabIndex        =   9
      Top             =   2280
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   450
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
      Format          =   "$#,##0.00;-$#,##0.00"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      DataField       =   "fecha"
      DataSource      =   "datordendepago"
      Height          =   375
      Left            =   5160
      TabIndex        =   2
      Top             =   600
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "dd/mm/yyyy"
      PromptChar      =   "_"
   End
   Begin MSAdodcLib.Adodc datordendepago 
      Height          =   330
      Left            =   360
      Top             =   3960
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
   Begin VB.Frame Frame1 
      Caption         =   "Nº Orden"
      Height          =   855
      Left            =   2040
      TabIndex        =   3
      Top             =   240
      Width           =   2895
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "frmordendepagoasigna.frx":1F50
         Height          =   360
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   741
         _Version        =   393216
         MatchEntry      =   -1  'True
         ListField       =   "nrorden"
         BoundColumn     =   "inicioper"
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
   End
   Begin MSAdodcLib.Adodc databonan 
      Height          =   330
      Left            =   480
      Top             =   4320
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
      RecordSource    =   "select ordendepagoabonan.* from ordendepagoabonan"
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
   Begin MSAdodcLib.Adodc datinstrumento 
      Height          =   330
      Left            =   9840
      Top             =   3720
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
   Begin MSAdodcLib.Adodc datproveedores 
      Height          =   330
      Left            =   120
      Top             =   5280
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
   Begin MSAdodcLib.Adodc datconsultacomp 
      Height          =   330
      Left            =   240
      Top             =   5040
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
   Begin MSAdodcLib.Adodc datlibrocompras 
      Height          =   330
      Left            =   240
      Top             =   4680
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
   Begin MSAdodcLib.Adodc datinstru 
      Height          =   330
      Left            =   0
      Top             =   4200
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
   Begin MSAdodcLib.Adodc datcuentas 
      Height          =   330
      Left            =   9480
      Top             =   3840
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
   Begin MSAdodcLib.Adodc datmaestro 
      Height          =   330
      Left            =   600
      Top             =   3960
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
   Begin MSAdodcLib.Adodc datasiento 
      Height          =   330
      Left            =   9960
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
   Begin Crystal.CrystalReport CrystalReporte 
      Left            =   9600
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      WindowTitle     =   "Orden de Pago"
   End
   Begin MSRDC.MSRDC reporte 
      Height          =   375
      Left            =   9960
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
      Connect         =   ""
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
   Begin MSAdodcLib.Adodc datordensinc 
      Height          =   330
      Left            =   240
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
   Begin VB.Frame Frame2 
      Caption         =   "Fecha de Pago"
      Height          =   855
      Left            =   5040
      TabIndex        =   4
      Top             =   240
      Width           =   1935
   End
   Begin VB.PictureBox saldocompro 
      Height          =   255
      Index           =   1
      Left            =   7080
      ScaleHeight     =   195
      ScaleWidth      =   1515
      TabIndex        =   28
      Top             =   1920
      Width           =   1575
   End
   Begin VB.PictureBox saldocompro 
      Height          =   255
      Index           =   2
      Left            =   7080
      ScaleHeight     =   195
      ScaleWidth      =   1515
      TabIndex        =   29
      Top             =   2160
      Width           =   1575
   End
   Begin VB.PictureBox saldocompro 
      Height          =   255
      Index           =   3
      Left            =   7080
      ScaleHeight     =   195
      ScaleWidth      =   1515
      TabIndex        =   30
      Top             =   2400
      Width           =   1575
   End
   Begin VB.PictureBox saldocompro 
      Height          =   255
      Index           =   4
      Left            =   7080
      ScaleHeight     =   195
      ScaleWidth      =   1515
      TabIndex        =   31
      Top             =   2640
      Width           =   1575
   End
   Begin VB.PictureBox saldocompro 
      Height          =   255
      Index           =   5
      Left            =   7080
      ScaleHeight     =   195
      ScaleWidth      =   1515
      TabIndex        =   32
      Top             =   2880
      Width           =   1575
   End
   Begin VB.PictureBox saldocompro 
      Height          =   255
      Index           =   6
      Left            =   7080
      ScaleHeight     =   195
      ScaleWidth      =   1515
      TabIndex        =   33
      Top             =   3120
      Width           =   1575
   End
   Begin VB.PictureBox saldocompro 
      Height          =   255
      Index           =   7
      Left            =   7080
      ScaleHeight     =   195
      ScaleWidth      =   1515
      TabIndex        =   34
      Top             =   3360
      Width           =   1575
   End
   Begin vbskpro.Skinner Skinner1 
      Left            =   240
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
      LcK2            =   $"frmordendepagoasigna.frx":1F6B
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
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      DataField       =   "importe"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2058
         SubFormatType   =   1
      EndProperty
      DataSource      =   "databonan"
      Height          =   285
      Index           =   2
      Left            =   1440
      TabIndex        =   41
      Top             =   4080
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.PictureBox saldolista 
      Height          =   255
      Left            =   7320
      ScaleHeight     =   195
      ScaleWidth      =   1515
      TabIndex        =   48
      Top             =   4920
      Width           =   1575
   End
   Begin VB.PictureBox totalorden 
      Height          =   255
      Left            =   7320
      ScaleHeight     =   195
      ScaleWidth      =   1515
      TabIndex        =   51
      Top             =   4200
      Width           =   1575
   End
   Begin MSMask.MaskEdBox asignar 
      Height          =   255
      Left            =   1560
      TabIndex        =   53
      Top             =   2760
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   0
      Appearance      =   0
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "$#,##0.00;-$#,##0.00"
      PromptChar      =   "_"
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Asignar:"
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
      Left            =   600
      TabIndex        =   54
      Top             =   2760
      Width           =   735
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Saldo de Orden ref.a lista:"
      Height          =   255
      Left            =   5280
      TabIndex        =   49
      Top             =   4920
      Width           =   1935
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Comp.Seleccionados:"
      Height          =   255
      Left            =   5280
      TabIndex        =   45
      Top             =   4560
      Width           =   2055
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Orden:"
      Height          =   255
      Left            =   5280
      TabIndex        =   44
      Top             =   4200
      Width           =   975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Comprobante"
      Height          =   255
      Left            =   3480
      TabIndex        =   37
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Shape Shape3 
      Height          =   495
      Left            =   480
      Top             =   2160
      Width           =   2535
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Saldo:"
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
      Left            =   600
      TabIndex        =   10
      Top             =   2280
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Conceptos que se abonan"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404080&
      Height          =   255
      Left            =   480
      TabIndex        =   5
      Top             =   1200
      Width           =   2535
   End
   Begin VB.Shape Shape1 
      Height          =   2535
      Left            =   360
      Top             =   1320
      Width           =   8535
   End
End
Attribute VB_Name = "frmordendepagoasigna"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim importeapagar As Double
Dim totalab As Currency
Dim totalinst(50) As Currency
Dim detalleint(50) As String
Dim totalconc(50) As Currency
Dim nrocompro(50) As String
Dim cuentaint(50) As Integer
Dim nomprov As String
Dim saldoactual As Currency
Dim saldo As Currency
Dim Cuenta As Integer
Dim codprove As Integer
Dim idlibrogrid(50) As Integer
Dim saldolibro(50) As Currency
Dim sincomp As Integer
Dim codigopago As Integer
Dim importeorden(10) As Currency
Public numorden As String
Dim totalfac As Currency
Dim importe As Currency
Dim empresareal As Integer
Dim inicioperiodo As Date
Dim listaasigna As Integer
Dim nota As String

Private Sub borrar_Click()
On Error GoTo erroreliminar

    nrocompro(DataGrid1.Row) = ""
    databonan.Recordset.Delete adAffectCurrent
    databonan.Refresh
Exit Sub

erroreliminar:
MsgBox "No se pudo eliminar Concepto"
    
End Sub


Private Sub aceptar_Click()
On Error GoTo fueraerror

totalfact = totalfac
If asignar.Text < totalfac Then
    totalfac = asignar.Text
End If

If saldo < 0 Then saldo = saldo * -1
    If totalfac >= saldo Then
        saldoanterior = saldo
        saldo = 0
    Else
       If saldolista.Value > 0 And saldolista.Value < Val(Text1(2).Text) Then saldo = saldolista.Value
       saldo = saldo - totalfac
    End If

    
    totalabonan.Text = saldo
    If nota = "NC" Then totalabonan = totalabonan * -1
 Rem   If nota = "NC" Then saldoanterior = saldoanterior * -1
    If saldo > 0 And totalfact = totalfac Then
        saldocompro(Cuenta - 3).Value = 0
        importeorden(Cuenta - 3) = totalfac
    End If
    If saldo > 0 And totalfact <> totalfac Then
        saldocompro(Cuenta - 3).Value = totalfact - totalfac
        importeorden(Cuenta - 3) = totalfac
    End If

    If saldo < 0 Then
        saldocompro(Cuenta - 3).Value = totalfac - saldoanterior
        importeorden(Cuenta - 3) = saldoanterior
    End If
    If saldo = 0 Then
        saldocompro(Cuenta - 3).Value = totalfac - saldoanterior
        importeorden(Cuenta - 3) = saldoanterior
    End If
    
For x = 0 To 10
    Debug.Print importeorden(x)
Next x


If nota = "NC" Then
    Command3.Enabled = True
    Exit Sub
End If

 
If saldo = 0 Or (saldo > 0 And saldo < 0.1) Or (saldo < 0 And saldo > -0.1) Then
    For x = 3 To 10
        Text1(x).Enabled = False
    Next x
    Command3.Enabled = True
    aceptar.Enabled = False
    Command3.SetFocus
    Exit Sub
End If
    totalfac = 0
    Text1(Cuenta + 1).SetFocus
Exit Sub
fueraerror:
    mensa = MsgBox("Demaciadas facturas para asignar", vbCritical, "Error")


End Sub

Private Sub asiglista_Click()


If listaasigna = 0 Then databonan.Recordset.Delete adAffectCurrent

MsgBox "hola"
For x = 0 To List1.ListCount - 1
List1.ListIndex = x
If List1.Selected(x) = True Then
        databonan.Recordset.AddNew
        databonan.Recordset.Fields("nrorden") = DataCombo1.Text
        databonan.Recordset.Fields("empresa") = login.empresaact
        databonan.Recordset.Fields("inicioper") = login.iper
        databonan.Recordset.Fields("finper") = login.fper
        databonan.Recordset.Fields("comprobante") = List1.Text
        databonan.Recordset.Fields("nomproveedor") = nomprov
        databonan.Recordset.Fields("codproveedor") = codprove
        comprob = Right(List1.Text, 13)
        If Left(tipocomprob, 1) = " " Then
            tipocomprob = " "
            For Y = 1 To 13
                car = Mid(comprob, Y, 1)
                If car <> " " Then GoTo finne
            Next Y
finne:
            comprob = Right(comprob, 14 - Y)
        End If
        
        datlibrocompras.RecordSource = "select librocompras.* from librocompras where empresa = " & login.empresaact & " and proveedor = '" & nomprov & "' and numcompr = '" & comprob & "'"
        datlibrocompras.Refresh
        databonan.Recordset.Fields("fechacompro") = datlibrocompras.Recordset.Fields("fecha")
        databonan.Recordset.Fields("importe") = datlibrocompras.Recordset.Fields("total")
        databonan.Recordset.Fields("saldofactura") = 0
        databonan.Recordset.UpdateBatch adAffectCurrent
        
        datlibrocompras.Recordset.Fields("saldo") = 0
        datlibrocompras.Recordset.Fields("imputado") = "S"
        datlibrocompras.Recordset.UpdateBatch adAffectCurrent
End If
Next x
Call nuevaorden_Click


End Sub

Private Sub asignar_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        If asignar.Text > (totalabonan.Text * -1) Then asignar.Text = 0
        If asignar.Text = 0 Then asignar.Text = totalfac
        aceptar.SetFocus
    End If

End Sub

Private Sub asignartodos_Click()
i = 0

Do While Not i = List1.ListCount
   
   List1.Selected(i) = True
   i = i + 1

Loop




End Sub

Private Sub Command1_Click()


    DataList1.Visible = True
    DataList3.Visible = True
    DataList1.SetFocus
    


End Sub

Private Sub Command2_Click()
Dim compro(10) As String

    For x = 3 To 10
             compro(x) = Text1(x).Text
             Debug.Print Text1(x).Text + " -"
    Next x
        
    If Text1(3).Text = "" Then GoTo fin
    codpro = codprove
    nompro = nomprov
    fechap = Date
    comprob = Right(compro(3), 13)
    tipocomprob = Left(compro(3), 4)
    If tipocomprob = "   " Then
        tipocomprob = " "
        For x = 1 To 13
            car = Mid(compro(3), x, 1)
            If car <> " " Then GoTo finne
        Next x
finne:
        comprob = Right(compro(3), 16 - x)
    End If
    
 databonan.RecordSource = "select ordendepagoabonan.* from ordendepagoabonan WHERE ordendepagoabonan.empresa = " & login.empresaact & " and inicioper = '" & login.iper & "' and finper = '" & login.fper & "' and  nrorden = '0' Order by nrorden"
 databonan.Refresh
    
    databonan.Recordset.AddNew
    databonan.Recordset.Fields(0) = DataCombo1.Text
    databonan.Recordset.Fields("empresa") = login.empresaact
    databonan.Recordset.Fields("inicioper") = login.iper
    databonan.Recordset.Fields("finper") = login.fper
    databonan.Recordset.Fields(7) = Text1(3).Text
    databonan.Recordset.Fields("importe") = importeorden(0)
    databonan.Recordset.Fields("saldofactura") = Val(saldocompro(0).Value)
    databonan.Recordset.Fields("comprobante") = compro(3)
    databonan.Recordset.Fields("nomproveedor") = nompro
    databonan.Recordset.Fields("codproveedor") = codpro
    databonan.Recordset.Fields("fechacompro") = fechap
    databonan.Recordset.UpdateBatch adAffectCurrent
        
    datlibrocompras.RecordSource = "select librocompras.* from librocompras where empresa = " & login.empresaact & " and proveedor = '" & nompro & "' and numcompr = '" & comprob & "'"
    datlibrocompras.Refresh
    datlibrocompras.Recordset.Fields("saldo") = Val(saldocompro(0).Value)
    datlibrocompras.Recordset.Fields("imputado") = "S"
    datlibrocompras.Recordset.UpdateBatch adAffectCurrent
    
    Inicio.datauditoria.ConnectionString = login.conexiontotal
    
    Inicio.datauditoria.RecordSource = "select auditoria.* from auditoria"
    Inicio.datauditoria.Refresh
    
    Inicio.datauditoria.Recordset.AddNew
    Inicio.datauditoria.Recordset.Fields("fecha") = Date
    Inicio.datauditoria.Recordset.Fields("hora") = Str(Time)
    Inicio.datauditoria.Recordset.Fields("ventana") = "Asignación de Notas de Credito"
    Inicio.datauditoria.Recordset.Fields("accion") = "Asignación Nota de Credito " + DataCombo1.Text
    Inicio.datauditoria.Recordset.Fields("usuario") = login.usuarioactivo
    Inicio.datauditoria.Recordset.Fields("empresa") = login.empresaact
    Inicio.datauditoria.Recordset.UpdateBatch adAffectCurrent

    For x = 4 To 10
        If compro(x) = "" Then GoTo fin
        databonan.Recordset.AddNew
        databonan.Recordset.Fields("nrorden") = DataCombo1.Text
        databonan.Recordset.Fields("empresa") = login.empresaact
        databonan.Recordset.Fields("inicioper") = login.iper
        databonan.Recordset.Fields("finper") = login.fper
        databonan.Recordset.Fields("comprobante") = compro(x)
        databonan.Recordset.Fields("nomproveedor") = nompro
        databonan.Recordset.Fields("codproveedor") = codpro
        databonan.Recordset.Fields("fechacompro") = fechap
        databonan.Recordset.Fields("importe") = importeorden(x - 3)
        databonan.Recordset.Fields("saldofactura") = Val(saldocompro(x - 3).Value)
        databonan.Recordset.UpdateBatch adAffectCurrent
        comprob = Right(compro(x), 13)
        tipocomprob = Left(compro(x), 4)
                                                                                                                                                                
                datlibrocompras.RecordSource = "select librocompras.* from librocompras where empresa = " & login.empresaact & " and proveedor = '" & nompro & "' and numcompr = '" & comprob & "'"
                datlibrocompras.Refresh
                datlibrocompras.Recordset.Fields("saldo") = Val(saldocompro(x - 3).Value)
                datlibrocompras.Recordset.Fields("imputado") = "S"
                datlibrocompras.Recordset.UpdateBatch adAffectCurrent
   Next x
    
fin:
    
    comprob = Right(DataCombo1.Text, 13)
    tipocomprob = Left(DataCombo1.Text, 4)
                                                                                                                                                    
    datlibrocompras.RecordSource = "select librocompras.* from librocompras where empresa = " & login.empresaact & " and proveedor = '" & nompro & "' and numcompr = '" & comprob & "'"
    datlibrocompras.Refresh
    datlibrocompras.Recordset.Fields("saldo") = totalabonan.Text
    datlibrocompras.Recordset.Fields("imputado") = "S"
    datlibrocompras.Recordset.UpdateBatch adAffectCurrent



DataList2.Visible = False
Call nuevaorden_Click

End Sub

Private Sub Command3_Click()
Dim compro(10) As String
    
   If nota = "NC" Then
        Call Command2_Click
        Exit Sub
   End If
    
    
    For x = 3 To 10
             compro(x) = Text1(x).Text
    Next x

    If Text1(3).Text = "" Then GoTo fin
    databonan.Recordset.Fields(7) = Text1(3).Text
    databonan.Recordset.Fields("importe") = importeorden(0)
    databonan.Recordset.Fields("saldofactura") = Val(saldocompro(0).Value)
    databonan.Recordset.Fields("saldoorden") = Val(totalabonan.Text)
    codpro = databonan.Recordset.Fields("codproveedor")
    nompro = databonan.Recordset.Fields("nomproveedor")
    fechap = databonan.Recordset.Fields("fechacompro")
    comprob = Right(compro(3), 13)
    If tipocomprob = "   " Then
        tipocomprob = " "
        For x = 1 To 13
            car = Mid(compro(3), x, 1)
            If car <> " " Then GoTo finne
        Next x
finne:
        comprob = Right(compro(3), 16 - x)
    End If
    datlibrocompras.RecordSource = "select librocompras.* from librocompras where empresa = " & login.empresaact & " and proveedor = '" & nompro & "' and numcompr = '" & comprob & "'"
    datlibrocompras.Refresh
    datlibrocompras.Recordset.Fields("saldo") = Val(saldocompro(0).Value)
    datlibrocompras.Recordset.Fields("imputado") = "S"
    datlibrocompras.Recordset.UpdateBatch adAffectCurrent
    databonan.Recordset.Fields("fechacompro") = datlibrocompras.Recordset.Fields("fecha")
    databonan.Recordset.UpdateBatch adAffectCurrent
    
    
    Inicio.datauditoria.ConnectionString = login.conexiontotal
    
    Inicio.datauditoria.RecordSource = "select auditoria.* from auditoria"
    Inicio.datauditoria.Refresh
    
    Inicio.datauditoria.Recordset.AddNew
    Inicio.datauditoria.Recordset.Fields("fecha") = Date
    Inicio.datauditoria.Recordset.Fields("hora") = Str(Time)
    Inicio.datauditoria.Recordset.Fields("ventana") = "Asignacion de Pagos sin comprobantes"
    Inicio.datauditoria.Recordset.Fields("accion") = "Asignacion de Orden: " + DataCombo1.Text
    Inicio.datauditoria.Recordset.Fields("usuario") = login.usuarioactivo
    Inicio.datauditoria.Recordset.Fields("empresa") = login.empresaact
    Inicio.datauditoria.Recordset.UpdateBatch adAffectCurrent

   
    For x = 4 To 10
        If compro(x) = "" Then GoTo fin
        databonan.Recordset.AddNew
        databonan.Recordset.Fields("nrorden") = DataCombo1.Text
        databonan.Recordset.Fields("empresa") = login.empresaact
        databonan.Recordset.Fields("inicioper") = login.iper
        databonan.Recordset.Fields("finper") = login.fper
        databonan.Recordset.Fields("comprobante") = compro(x)
        databonan.Recordset.Fields("nomproveedor") = nompro
        databonan.Recordset.Fields("codproveedor") = codpro
        comprob = Right(compro(x), 13)
                datlibrocompras.RecordSource = "select librocompras.* from librocompras where empresa = " & login.empresaact & " and proveedor = '" & nompro & "' and numcompr = '" & comprob & "'"
                datlibrocompras.Refresh
                datlibrocompras.Recordset.Fields("saldo") = Val(saldocompro(x - 3).Value)
                datlibrocompras.Recordset.Fields("imputado") = "S"
                datlibrocompras.Recordset.UpdateBatch adAffectCurrent
        databonan.Recordset.Fields("fechacompro") = datlibrocompras.Recordset.Fields("fecha")
        databonan.Recordset.Fields("importe") = importeorden(x - 3)
        databonan.Recordset.Fields("saldofactura") = Val(saldocompro(x - 3).Value)
        databonan.Recordset.UpdateBatch adAffectCurrent
    Next x

fin:

    If Val(saldolista.Text) >= 0 And Val(saldolista.Text) < Val(totalorden.Text) Then
                listaasigna = 1
                Call asiglista_Click
                Exit Sub
    End If

    
DataList2.Visible = False
Call nuevaorden_Click

End Sub

Private Sub Command4_Click()
Dim tabla As String
Dim tabla1 As String
Dim ruta As String

ruta = "\Empresa" + Right(Str(login.empresaact), Len(Str(login.empresaact)) - 1)

reporte.SQL = "consultaordesnpago.nrorden, consultaordesnpago.empresa, consultaordesnpago.nomproveedor, consultaordesnpago.comprobante, consultaordesnpago.fechacompro, consultaordesnpago.importe, consultaordesnpago.id, consultaordesnpago.razonsocial, consultaordesnpago.cuit, consultaordesnpago.domicilio, consultaordesnpago.localidad, consultaordesnpago.fecha, consultaordesnpago.domprov, consultaordesnpago.locprov, consultaordesnpago.cuitprov, consultaordesnpago.saldofactura FROM contablesql.dbo.consultaordesnpago consultaordesnpago WHERE consultaordesnpago.nrorden= '" & frmordendepago.numorden & "' and consultaordesnpago.empresa = " & login.empresaact & " ORDER BY consultaordesnpago.razonsocial ASC, consultaordesnpago.id ASC"
tabla = reporte.SQL

With CrystalReporte
    .ReportFileName = App.Path & ruta + "\Ordendepago.rpt"
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
    .ReportFileName = App.Path & ruta + "\Ordendepago1.rpt"
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
      
End With
End Sub

Private Sub Command5_Click()

    DataGrid1.Columns(6).Caption = "Detalle de Pago"
    DataGrid1.Columns(7).Visible = False
    DataGrid1.Columns(8).Caption = "Fecha"
    DataGrid1.Columns(10).Visible = True
    DataGrid1.Columns(10).Width = 1395
    DataGrid1.Columns(11).Locked = False
    sincomp = 1

End Sub


Private Sub DataCombo1_KeyPress(KeyAscii As Integer)
On Error GoTo errormod
    If KeyAscii = 13 Then
        KeyAscii = 0
        If DataCombo1.Text = "" Then
          Rem   Command1.SetFocus
            Exit Sub
        End If
        Command3.Enabled = False
        nota = Left(DataCombo1.Text, 2)
        asignar.Visible = False
        Label7.Visible = False
        inicioperiodo = DataCombo1.BoundText
        If nota <> "NC" Then
            databonan.RecordSource = "select ordendepagoabonan.* from ordendepagoabonan WHERE ordendepagoabonan.empresa = " & login.empresaact & " and inicioper = '" & inicioperiodo & "'  and nrorden = '" & DataCombo1.Text & "' Order by fechacompro, id"
            databonan.Refresh
            codprove = databonan.Recordset.Fields(5)
            nomprov = databonan.Recordset.Fields(6)
        Else
            asignar.Visible = True
            Label7.Visible = True
            databonan.RecordSource = "select consultaordensinc.* from consultaordensinc where consultaordensinc.empresa = " & login.empresaact & " and nrorden = '" & DataCombo1.Text & "'"
            databonan.Refresh
            codprove = databonan.Recordset.Fields("codproveedor")
            nomprov = databonan.Recordset.Fields("nomproveedor")
        End If
        
        datinstrumento.RecordSource = "select ordendepagoinstrumento.* from ordendepagoinstrumento WHERE ordendepagoinstrumento.empresa = " & login.empresaact & " and inicioper = '" & inicioperiodo & "' and nrorden = '" & DataCombo1.Text & "' Order by id"
        datinstrumento.Refresh


        datconsultacomp.RecordSource = "select conceptosabonan.* from conceptosabonan where empresa = " & login.empresaact & " and codproveedor = " & codprove & "  order by comp"
        datconsultacomp.Refresh
        datordendepago.RecordSource = "select ordendepago.* from ordendepago WHERE ordendepago.empresa = " & login.empresaact & " and inicioper = '" & inicioperiodo & "' and nrorden = '" & DataCombo1.Text & "' Order by nrorden"
        datordendepago.Refresh
        
        If datconsultacomp.Recordset.EOF = True Then Exit Sub
        List1.Clear
        listaasigna = 0
        importe = 0
        totalorden.Value = Val(Text1(2).Text)
        totalabonan.Text = totalorden.Value
        i = 0
        datconsultacomp.Recordset.MoveFirst
        Do While Not datconsultacomp.Recordset.EOF
   
            listatxt = datconsultacomp.Recordset.Fields("comp")
            List1.AddItem listatxt
            List1.Selected(i) = False
            i = i + 1
            datconsultacomp.Recordset.MoveNext
        Loop
        saldolista.Value = totalorden.Value
        
    End If
    databonan.Recordset.MoveLast
    
If IsNull(saldo) = False Then saldo = Text1(2).Text
    Text1(3).SetFocus
errormod:
End Sub

Private Sub DataGrid1_AfterColEdit(ByVal ColIndex As Integer)
Exit Sub
    If DataGrid1.Col = 9 Then
        importeapagar = DataGrid1.Columns(9).Value
        If sincomp = 0 Then saldoanterior = DataGrid1.Columns(11).Value
        If sincomp = 1 Then saldoanterior = importeapagar
        diferencia = importeapagar - saldoanterior
        If diferencia > 0.009 Then
            mensa = MsgBox("No se puede imputar el pago, el importe a pagar es mayor que el saldo", vbCritical, "!! Error !!")
            Exit Sub
        End If
        If diferencia < 0 Then diferencia = diferencia * -1
        saldoactual = saldoanterior - importeapagar
        If diferencia <= 0.009 Then saldoactual = 0
        saldolibro(DataGrid1.Row) = saldoactual
        idlibrogrid(DataGrid1.Row) = DataGrid4.Columns(0).Text
 Rem       DataGrid4.Columns(63).Value = saldoactual
 Rem       DataGrid4.Columns(64).Value = "S"
        DataGrid1.Columns(11).Value = saldoactual
        DataGrid1.Columns(8).Value = DataGrid4.Columns(2).Value
        totalconc(DataGrid1.Row) = importeapagar
        datlibrocompras.Recordset.UpdateBatch adAffectCurrent
        datordendepago.Recordset.UpdateBatch adAffectCurrent
  Rem      databonan.Recordset.UpdateBatch adAffectCurrent
        DataGrid1.Refresh
        totalab = 0
        For x = 0 To DataGrid1.Row
            totalab = totalab + totalconc(x)
        Next x
        totalabonan.Text = totalab
        Exit Sub
    End If
    
End Sub

Private Sub DataGrid1_ButtonClick(ByVal ColIndex As Integer)

If ColIndex = 7 Then
        DataList2.Visible = True
        DataList2.Left = DataGrid1.Columns(7).Left + DataGrid1.Left
        DataList2.Width = DataGrid1.Columns(7).Width
        DataList2.Top = DataGrid1.Top + DataGrid1.RowTop(DataGrid1.Row) + DataGrid1.RowHeight
        DataList2.SetFocus
End If
    
End Sub


Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
On Error GoTo erroringreso

  
    If KeyAscii = 13 And DataGrid1.Col = 7 Then
        If DataGrid1.Columns(7).Text = "" Then
                    DataList2.Visible = True
                    DataList2.Left = DataGrid1.Columns(7).Left + DataGrid1.Left
                    DataList2.Width = DataGrid1.Columns(7).Width
                    DataList2.Top = DataGrid1.Top + DataGrid1.RowTop(DataGrid1.Row) + DataGrid1.RowHeight
                    KeyAscii = 0
                    DataList2.SetFocus
                    Exit Sub
        Else
            KeyAscii = 9
        End If
    End If
    If KeyAscii = 13 And DataGrid1.Col = 8 Then
          KeyAscii = 9
    End If
    If KeyAscii = 13 And DataGrid1.Col = 9 Then
        KeyAscii = 9
    End If
    
    If KeyAscii = 13 And DataGrid1.Col = 11 And sincomp = 1 Then
        KeyAscii = 0
        nuevo.SetFocus
    End If
Exit Sub
erroringreso:
    Call nuevo.SetFocus



End Sub

Private Sub DataList1_Click()

    datordensinc.RecordSource = "select consultaordensinc.* from consultaordensinc where consultaordensinc.empresa = " & login.empresaact & " and codproveedor = " & DataList1.BoundText & " "
    datordensinc.Refresh

End Sub



Private Sub DataList1_KeyDown(KeyCode As Integer, Shift As Integer)

    If DataList1.BoundText <> "" Then
        datordensinc.RecordSource = "select consultaordensinc.* from consultaordensinc where consultaordensinc.empresa = " & login.empresaact & " and codproveedor = " & DataList1.BoundText & " "
        datordensinc.Refresh
    End If
    
End Sub

Private Sub DataList1_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        If DataList3.VisibleCount = 0 Then
            DataList1.Visible = False
            DataList3.Visible = False
            DataCombo1.SetFocus
            Exit Sub
        Else
            DataList3.SetFocus
        End If
    End If
    
End Sub

Private Sub DataList2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
            Text1(Cuenta).Text = DataList2.Text
            compro = DataList2.Text
            If compro = "" Then Exit Sub
   
            For x = (Cuenta - 1) To 3 Step -1
                If Text1(Cuenta).Text = Text1(x).Text Then
                    mensa = MsgBox("Este comprobante ya fue ingresado, cambielo", vbCritical, "!! Error !!")
                    Text1(Cuenta).SetFocus
                    Exit Sub
                End If
            Next x
                
                For Y = 0 To List1.ListCount - 1
                    List1.ListIndex = Y
                    If List1.Selected(Y) = True Then
                        If Text1(Cuenta).Text = List1.Text Then
                            mensa = MsgBox("Este comprobante ya fue ingresado, cambielo", vbCritical, "!! Error !!")
                            Text1(Cuenta).SetFocus
                            Exit Sub
                        End If
                    End If
                Next Y
            
            datconsultacomp.RecordSource = "select conceptosabonan.* from conceptosabonan where empresa = " & login.empresaact & " and codproveedor = " & codprove & " and comp = '" & compro & "' order by comp"
            datconsultacomp.Refresh
            If DataGrid3.Columns(9).Text = "" Then
                DataGrid3.Columns(6).Visible = True
                DataGrid3.Columns(9).Visible = False
                totalfac = DataGrid3.Columns(6).Text
            Else
                DataGrid3.Columns(6).Visible = False
                DataGrid3.Columns(9).Visible = True
                totalfac = DataGrid3.Columns(9).Text
            End If
            totalfactura.Text = totalfac
            asignar.Text = totalfac
            DataList2.Visible = False
            aceptar.SetFocus
    End If
End Sub


Private Sub DataList2_LostFocus()
            
            DataList2.Visible = False
            
End Sub


Private Sub DataList4_GotFocus()
    If Inicio.opcion1 = True Then
        datcuentas.RecordSource = "select listacuentas.* from listacuentas WHERE empre = " & login.empresaact & " and inicioper = '" & login.iper & "' and finper = '" & login.fper & "' ORDER BY IDCUENTA"
        datcuentas.Refresh
        DataList4.ListField = "codigo"
    End If
    If Inicio.opcion2 = True Then
        datcuentas.RecordSource = "select listacuentas.* from listacuentas WHERE empre = " & login.empresaact & " and inicioper = '" & login.iper & "' and finper = '" & login.fper & "' ORDER BY nombre"
        datcuentas.Refresh
        DataList4.ListField = "nombre"
    End If
End Sub

Private Sub DataList3_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        DataCombo1.Text = DataList3.Text
        DataCombo1.SetFocus
        DataList3.Visible = False
        DataList1.Visible = False
    End If
    

End Sub

Private Sub Form_Load()

    Inicio.Toolbar1.Visible = True

frmordendepagoasigna.Top = 0
frmordendepagoasigna.Left = 0
datasiento.ConnectionString = login.conexiontotal
datconsultacomp.ConnectionString = login.conexiontotal
datcuentas.ConnectionString = login.conexiontotal
datinstru.ConnectionString = login.conexiontotal
datinstrumento.ConnectionString = login.conexiontotal
datlibrocompras.ConnectionString = login.conexiontotal
datmaestro.ConnectionString = login.conexiontotal
datordendepago.ConnectionString = login.conexiontotal
datproveedores.ConnectionString = login.conexiontotal
datordensinc.ConnectionString = login.conexiontotal


If login.empresaact1 > 0 Then
    empresareal = login.empresaact1
Else
    empresareal = login.empresaact
End If

asiglista.Enabled = False

  datordendepago.RecordSource = "select ordendepago.* from ordendepago WHERE ordendepago.empresa = " & login.empresaact & " and inicioper = '" & login.iper & "' and finper = '" & login.fper & "' and nrorden = 0 Order by nrorden"
  datordendepago.Refresh

    datordensinc.RecordSource = "select consultaordensinc.* from consultaordensinc where consultaordensinc.empresa = " & login.empresaact & ""
    datordensinc.Refresh
        
 databonan.RecordSource = "select ordendepagoabonan.* from ordendepagoabonan WHERE ordendepagoabonan.empresa = " & login.empresaact & " and inicioper = '" & login.iper & "' and finper = '" & login.fper & "' and  nrorden = '0' Order by nrorden"
 databonan.Refresh
  
 datinstrumento.RecordSource = "select ordendepagoinstrumento.* from ordendepagoinstrumento WHERE ordendepagoinstrumento.empresa = " & login.empresaact & " and inicioper = '" & login.iper & "' and finper = '" & login.fper & "' and nrorden = '" & orden & "' Order by id"
  datinstrumento.Refresh
  
  datproveedores.RecordSource = "select proveedores.* from proveedores where empresa = " & empresareal & " order by razonsocial"
  datproveedores.Refresh
  

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Inicio.Toolbar1.Visible = False

End Sub

Private Sub limpia_Click()
i = 0

Do While Not i = List1.ListCount
   
   List1.Selected(i) = False
   i = i + 1

Loop
impfactura.Value = 0
saldolista.Value = totalorden.Value
asiglista.Enabled = False

End Sub

Private Sub limpiar_Click()
                      
    saldo = Text1(2).Text

    For x = 3 To 10
        Text1(x).Text = ""
        saldocompro(x - 3).Value = 0
        Text1(x).Enabled = True
    Next x
    aceptar.Enabled = True
    saldo = Text1(2).Text
    totalabonan = saldo
    totalfactura.Text = ""
    Command3.Enabled = False
    Text1(3).SetFocus
    
End Sub

Private Sub List1_ItemCheck(Item As Integer)

    If List1.Selected(List1.ListIndex) = True Then
            datconsultacomp.RecordSource = "select conceptosabonan.* from conceptosabonan where empresa = " & login.empresaact & " and codproveedor = " & codprove & " and comp = '" & List1.Text & "'"
            datconsultacomp.Refresh
            If DataGrid3.Columns(9).Text = "" Then
                DataGrid3.Columns(6).Visible = True
                DataGrid3.Columns(9).Visible = False
                importe = DataGrid3.Columns(6).Text
            Else
                DataGrid3.Columns(6).Visible = False
                DataGrid3.Columns(9).Visible = True
                importe = DataGrid3.Columns(9).Text
            End If
            impfactura.Text = importe + Val(impfactura.Text)
    End If
    If List1.Selected(List1.ListIndex) = False Then
            datconsultacomp.RecordSource = "select conceptosabonan.* from conceptosabonan where empresa = " & login.empresaact & " and codproveedor = " & codprove & " and comp = '" & List1.Text & "'"
            datconsultacomp.Refresh
            If DataGrid3.Columns(9).Text = "" Then
                DataGrid3.Columns(6).Visible = True
                DataGrid3.Columns(9).Visible = False
                importe = DataGrid3.Columns(6).Text
            Else
                DataGrid3.Columns(6).Visible = False
                DataGrid3.Columns(9).Visible = True
                importe = DataGrid3.Columns(9).Text
            End If
            impfactura.Text = Val(impfactura.Text) - importe
    End If

    saldolista.Value = Val(Text1(2).Text) - impfactura.Value
    If saldolista.Value = 0 Then
        asiglista.Enabled = True
    Else
        asiglista.Enabled = False
    End If
    

End Sub

Private Sub MaskEdBox1_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        nuevo.SetFocus
    End If
    
End Sub

Private Sub nuevaorden_Click()
   
Rem DataCombo1.Text = ""
Rem totalabonan.Text = ""
Rem    For X = 3 To 10
Rem        Text1(X).Text = ""
Rem        saldocompro(X - 3).Value = 0
Rem        importeorden(X - 3) = 0
Rem        Text1(X).Enabled = True
Rem    Next X
Rem    aceptar.Enabled = True
    
Rem totalfactura.Text = ""
Unload Me
frmordendepagoasigna.Show



End Sub


Private Sub salir_Click()

    Unload Me

End Sub


Private Sub Text1_GotFocus(Index As Integer)
    
       If Index >= 3 Then
        DataList2.Visible = True
        DataList2.Left = Text1(Index).Left
        DataList2.Width = Text1(Index).Width
        DataList2.Top = Text1(Index).Top + Text1(Index).Height
        Cuenta = Index
        datconsultacomp.RecordSource = "select conceptosabonan.* from conceptosabonan where empresa = " & login.empresaact & " and codproveedor = " & codprove & "  order by comp"
        datconsultacomp.Refresh
        DataList2.SetFocus
       End If
       
End Sub
