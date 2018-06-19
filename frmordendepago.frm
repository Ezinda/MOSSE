VERSION 5.00
Object = "{F6125AB1-8AB1-11CE-A77F-08002B2F4E98}#2.0#0"; "MSRDC20.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{5734474E-78D3-4254-99B9-C35F31BDF509}#62.0#0"; "vbskpro2.ocx"
Begin VB.Form frmordendepago 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Orden de pago"
   ClientHeight    =   8085
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10905
   Icon            =   "frmordendepago.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8085
   ScaleWidth      =   10905
   Begin MSDataListLib.DataList DataList1 
      Bindings        =   "frmordendepago.frx":0442
      Height          =   1815
      Left            =   5040
      TabIndex        =   4
      Top             =   4200
      Visible         =   0   'False
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   3201
      _Version        =   393216
      MatchEntry      =   -1  'True
      BackColor       =   12632256
      ListField       =   "razonsocial"
      BoundColumn     =   "codproveedor"
   End
   Begin MSMask.MaskEdBox saldototal 
      Height          =   255
      Left            =   5040
      TabIndex        =   31
      Top             =   7320
      Width           =   1575
      _ExtentX        =   2778
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
   Begin MSDataListLib.DataList DataList4 
      Bindings        =   "frmordendepago.frx":045F
      Height          =   1815
      Left            =   3600
      TabIndex        =   9
      Top             =   4200
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   3201
      _Version        =   393216
      MatchEntry      =   -1  'True
      OLEDragMode     =   1
      OLEDropMode     =   1
      BackColor       =   12632256
      ListField       =   "codigo"
      BoundColumn     =   "Cod Contable"
   End
   Begin MSDataListLib.DataList DataList2 
      Bindings        =   "frmordendepago.frx":0478
      Height          =   1815
      Left            =   960
      TabIndex        =   5
      Top             =   4200
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
   Begin MSDataListLib.DataList DataList3 
      Bindings        =   "frmordendepago.frx":0496
      Height          =   1815
      Left            =   2280
      TabIndex        =   8
      Top             =   4200
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   3201
      _Version        =   393216
      MatchEntry      =   -1  'True
      BackColor       =   12632256
      ListField       =   "instrumento"
      BoundColumn     =   "instrumento"
   End
   Begin MSMask.MaskEdBox totalinstrumento 
      Height          =   255
      Left            =   8160
      TabIndex        =   29
      Top             =   7320
      Width           =   2055
      _ExtentX        =   3625
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
   Begin MSMask.MaskEdBox totalabonan 
      Height          =   255
      Left            =   8160
      TabIndex        =   25
      Top             =   6960
      Width           =   2055
      _ExtentX        =   3625
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
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7080
      TabIndex        =   26
      Text            =   "A Pagar:"
      Top             =   6960
      Width           =   1095
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7320
      TabIndex        =   28
      Text            =   "Total:"
      Top             =   7320
      Width           =   615
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4440
      TabIndex        =   27
      Text            =   "Saldo:"
      Top             =   7320
      Width           =   615
   End
   Begin VB.Frame Frame4 
      Height          =   975
      Left            =   4080
      TabIndex        =   30
      Top             =   6720
      Width           =   6375
   End
   Begin MSDataGridLib.DataGrid DataGrid4 
      Bindings        =   "frmordendepago.frx":04AE
      Height          =   1215
      Left            =   5520
      TabIndex        =   7
      Top             =   2640
      Visible         =   0   'False
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   2143
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
      Bindings        =   "frmordendepago.frx":04CC
      Height          =   1215
      Left            =   1200
      TabIndex        =   6
      Top             =   2640
      Visible         =   0   'False
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   2143
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
   Begin TabDlg.SSTab SSTab1 
      Height          =   6855
      Left            =   120
      TabIndex        =   16
      Top             =   1080
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   12091
      _Version        =   393216
      Tab             =   2
      TabHeight       =   520
      ForeColor       =   128
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "F1-CONCEPTOS PAGOS A PROVEED."
      TabPicture(0)   =   "frmordendepago.frx":04EA
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "DataGrid1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "nuevo"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "borrar"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Command5"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cancelar"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "borrablancos"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "F2-INSTRUMENTO DE PAGO"
      TabPicture(1)   =   "frmordendepago.frx":0506
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "DataGrid2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Command2"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Command1"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "F3-OTROS CONCEPTOS"
      TabPicture(2)   =   "frmordendepago.frx":0522
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "DataGrid5"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "eliminaotroconcepto"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "nuevootroconcepto"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).ControlCount=   3
      Begin VB.CommandButton nuevootroconcepto 
         Caption         =   "&Nuevo Concepto"
         Height          =   855
         Left            =   360
         Picture         =   "frmordendepago.frx":053E
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   5760
         UseMaskColor    =   -1  'True
         Width           =   1455
      End
      Begin VB.CommandButton eliminaotroconcepto 
         Caption         =   "&Eliminar Concepto"
         Height          =   855
         Left            =   2160
         Picture         =   "frmordendepago.frx":0A70
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   5760
         UseMaskColor    =   -1  'True
         Width           =   1455
      End
      Begin VB.CommandButton borrablancos 
         Caption         =   "borrablancos"
         Height          =   255
         Left            =   -65640
         TabIndex        =   32
         Top             =   0
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Nuevo Ins&tru."
         Height          =   855
         Left            =   -74520
         Picture         =   "frmordendepago.frx":0B72
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   5760
         UseMaskColor    =   -1  'True
         Width           =   1455
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Eliminar Instr&u."
         Height          =   855
         Left            =   -72720
         Picture         =   "frmordendepago.frx":10A4
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   5760
         UseMaskColor    =   -1  'True
         Width           =   1455
      End
      Begin VB.CommandButton cancelar 
         Cancel          =   -1  'True
         Caption         =   "Command4"
         Height          =   375
         Left            =   -71040
         TabIndex        =   21
         Top             =   5280
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton Command5 
         Caption         =   "&Sin Comprobantes"
         Height          =   255
         Left            =   -72000
         TabIndex        =   20
         Top             =   480
         Width           =   1935
      End
      Begin VB.CommandButton borrar 
         Caption         =   "&Eliminar Concepto"
         Height          =   855
         Left            =   -72600
         Picture         =   "frmordendepago.frx":11A6
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   5760
         UseMaskColor    =   -1  'True
         Width           =   1455
      End
      Begin VB.CommandButton nuevo 
         Caption         =   "&Nuevo Concepto"
         Height          =   855
         Left            =   -74400
         Picture         =   "frmordendepago.frx":12A8
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   5760
         UseMaskColor    =   -1  'True
         Width           =   1455
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "frmordendepago.frx":17DA
         Height          =   4815
         Left            =   -74880
         TabIndex        =   18
         Top             =   720
         Width           =   10455
         _ExtentX        =   18441
         _ExtentY        =   8493
         _Version        =   393216
         AllowArrows     =   -1  'True
         HeadLines       =   1
         RowHeight       =   15
         TabAction       =   1
         WrapCellPointer =   -1  'True
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
         ColumnCount     =   12
         BeginProperty Column00 
            DataField       =   "nrorden"
            Caption         =   "nrorden"
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
         BeginProperty Column03 
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
         BeginProperty Column04 
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
         BeginProperty Column05 
            DataField       =   "codproveedor"
            Caption         =   "codproveedor"
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
            DataField       =   "nomproveedor"
            Caption         =   "Proveedor Razon Social"
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
            DataField       =   "comprobante"
            Caption         =   "Nº Comprobante"
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
            DataField       =   "fechacompro"
            Caption         =   "Fecha Compr."
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
            DataField       =   "importe"
            Caption         =   "Importe"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   """$""#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   2
            EndProperty
         EndProperty
         BeginProperty Column10 
            DataField       =   "codcuenta"
            Caption         =   "Cod.Cuenta"
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
            DataField       =   "saldofactura"
            Caption         =   "Saldo Compr."
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   """$""#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   2
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
               Button          =   -1  'True
               Locked          =   -1  'True
            EndProperty
            BeginProperty Column07 
               Button          =   -1  'True
            EndProperty
            BeginProperty Column08 
               Alignment       =   2
               ColumnAllowSizing=   -1  'True
            EndProperty
            BeginProperty Column09 
               Alignment       =   1
            EndProperty
            BeginProperty Column10 
               Alignment       =   2
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column11 
               Alignment       =   1
               Locked          =   -1  'True
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid DataGrid2 
         Bindings        =   "frmordendepago.frx":17F2
         Height          =   5055
         Left            =   -74880
         TabIndex        =   24
         Top             =   480
         Width           =   10455
         _ExtentX        =   18441
         _ExtentY        =   8916
         _Version        =   393216
         AllowArrows     =   -1  'True
         HeadLines       =   1
         RowHeight       =   15
         TabAction       =   1
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
            DataField       =   "nrorden"
            Caption         =   "nrorden"
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
         BeginProperty Column03 
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
         BeginProperty Column04 
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
         BeginProperty Column05 
            DataField       =   "idinstrumento"
            Caption         =   "idinstrumento"
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
            DataField       =   "instrumento"
            Caption         =   "Instrumento "
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
            DataField       =   "denominacion"
            Caption         =   "Denominacion"
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
            DataField       =   "comprobante"
            Caption         =   "Nº Compr."
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
            DataField       =   "fechacompro"
            Caption         =   "Fecha Emisión"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "dd/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   3
            EndProperty
         EndProperty
         BeginProperty Column10 
            DataField       =   "fechavencim"
            Caption         =   "Fecha Venc."
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "dd/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   3
            EndProperty
         EndProperty
         BeginProperty Column11 
            DataField       =   "importe"
            Caption         =   "Importe"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   """$"" #.##0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   2
            EndProperty
         EndProperty
         BeginProperty Column12 
            DataField       =   "codcuenta"
            Caption         =   "Cod.Cuenta"
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
               ColumnAllowSizing=   0   'False
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
               Button          =   -1  'True
            EndProperty
            BeginProperty Column07 
            EndProperty
            BeginProperty Column08 
            EndProperty
            BeginProperty Column09 
               Alignment       =   2
            EndProperty
            BeginProperty Column10 
               Alignment       =   2
            EndProperty
            BeginProperty Column11 
               Alignment       =   1
            EndProperty
            BeginProperty Column12 
               Alignment       =   2
               Button          =   -1  'True
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid DataGrid5 
         Bindings        =   "frmordendepago.frx":180F
         Height          =   4815
         Left            =   120
         TabIndex        =   33
         Top             =   480
         Width           =   10455
         _ExtentX        =   18441
         _ExtentY        =   8493
         _Version        =   393216
         AllowArrows     =   -1  'True
         HeadLines       =   1
         RowHeight       =   15
         TabAction       =   1
         WrapCellPointer =   -1  'True
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
         ColumnCount     =   12
         BeginProperty Column00 
            DataField       =   "nrorden"
            Caption         =   "nrorden"
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
         BeginProperty Column03 
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
         BeginProperty Column04 
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
         BeginProperty Column05 
            DataField       =   "codproveedor"
            Caption         =   "codproveedor"
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
            DataField       =   "nomproveedor"
            Caption         =   "Detalle Concepto"
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
            DataField       =   "comprobante"
            Caption         =   "Nº Comprobante"
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
            DataField       =   "fechacompro"
            Caption         =   "Fecha Compr."
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
            DataField       =   "importe"
            Caption         =   "Importe"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   """$""#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   2
            EndProperty
         EndProperty
         BeginProperty Column10 
            DataField       =   "codcuenta"
            Caption         =   "Cod.Cuenta"
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
            DataField       =   "saldofactura"
            Caption         =   "Saldo Compr."
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   """$""#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   2
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
               Object.Visible         =   -1  'True
            EndProperty
            BeginProperty Column07 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column08 
               Alignment       =   2
               ColumnAllowSizing=   -1  'True
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column09 
               Alignment       =   1
               Object.Visible         =   -1  'True
            EndProperty
            BeginProperty Column10 
               Alignment       =   2
               Button          =   -1  'True
               Object.Visible         =   -1  'True
            EndProperty
            BeginProperty Column11 
               Alignment       =   1
               Locked          =   -1  'True
               Object.Visible         =   0   'False
            EndProperty
         EndProperty
      End
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Original"
      Height          =   195
      Index           =   0
      Left            =   7800
      TabIndex        =   14
      Top             =   480
      Width           =   1095
   End
   Begin VB.CommandButton salir 
      Caption         =   "&Cerrar"
      Height          =   1095
      Left            =   9360
      Picture         =   "frmordendepago.frx":1827
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   255
      Left            =   5400
      TabIndex        =   12
      Top             =   120
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton nuevaorden 
      Caption         =   "Nueva &Orden"
      Height          =   735
      Left            =   5760
      Picture         =   "frmordendepago.frx":1C69
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   240
      UseMaskColor    =   -1  'True
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Grabar Orden"
      Height          =   735
      Left            =   4200
      Picture         =   "frmordendepago.frx":20AB
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   240
      UseMaskColor    =   -1  'True
      Width           =   1335
   End
   Begin MSMask.MaskEdBox text1 
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "########"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   375
      Left            =   1800
      TabIndex        =   0
      Top             =   480
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
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
      Left            =   480
      Top             =   6480
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
      Height          =   1095
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   1455
   End
   Begin VB.Frame Frame2 
      Caption         =   "Fecha"
      Height          =   1095
      Left            =   1560
      TabIndex        =   3
      Top             =   0
      Width           =   5775
   End
   Begin MSAdodcLib.Adodc databonan 
      Height          =   330
      Left            =   2400
      Top             =   6720
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
   Begin MSAdodcLib.Adodc datinstrumento 
      Height          =   330
      Left            =   3600
      Top             =   6720
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
      Left            =   4680
      Top             =   6720
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
      Left            =   6000
      Top             =   6720
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
      Left            =   7200
      Top             =   6720
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
      Left            =   8400
      Top             =   6720
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
      Left            =   9360
      Top             =   6720
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
      Left            =   480
      Top             =   6840
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
      Left            =   1680
      Top             =   6840
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
      Left            =   9360
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTitle     =   "Orden de Pago"
      PrintFileLinesPerPage=   60
   End
   Begin MSRDC.MSRDC reporte 
      Height          =   375
      Left            =   9720
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
   Begin VB.Frame Frame3 
      Caption         =   "Imprimir Solamente"
      Height          =   1095
      Left            =   7320
      TabIndex        =   15
      Top             =   0
      Width           =   2055
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
      LcK2            =   $"frmordendepago.frx":24ED
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
   Begin MSAdodcLib.Adodc criterio 
      Height          =   330
      Left            =   0
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
End
Attribute VB_Name = "frmordendepago"
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
Dim cuentaotros(50) As Integer
Dim conceptosotros(50) As String
Dim nomprov(50) As String
Dim saldoactual As Currency
Dim cuenta As Integer
Dim codprove As Long
Dim idlibrogrid(50) As Integer
Dim saldolibro(50) As Currency
Dim sincomp As Integer
Dim codigopago As Integer
Dim empresareal As Integer
Public numorden As String

Private Sub borrablancos_Click()
On Error GoTo fuera

  databonan.RecordSource = "select ordendepagoabonan.* from ordendepagoabonan WHERE ordendepagoabonan.empresa = " & login.empresaact & " and inicioper = '" & login.iper & "' and nrorden = '" & text1.Text & "' and '" & IsNull(codproveedor) & "' ='" & True & "'  Order by fechacompro"
  databonan.Refresh
  If databonan.Recordset.EOF = False Then databonan.Recordset.Delete adAffectCurrent
  databonan.RecordSource = "select ordendepagoabonan.* from ordendepagoabonan WHERE ordendepagoabonan.empresa = " & login.empresaact & " and inicioper = '" & login.iper & "' and nrorden = '" & text1.Text & "' Order by fechacompro"
  databonan.Refresh
  
  datinstrumento.RecordSource = "select ordendepagoinstrumento.* from ordendepagoinstrumento WHERE ordendepagoinstrumento.empresa = " & login.empresaact & " and inicioper = '" & login.iper & "'  and nrorden = '" & text1.Text & "' and '" & IsNull(codproveedor) & "' ='" & True & "' Order by id"
  datinstrumento.Refresh
  If datinstrumento.Recordset.EOF = False Then datinstrumento.Recordset.Delete adAffectCurrent
  datinstrumento.RecordSource = "select ordendepagoinstrumento.* from ordendepagoinstrumento WHERE ordendepagoinstrumento.empresa = " & login.empresaact & " and inicioper = '" & login.iper & "'  and nrorden = '" & text1.Text & "' Order by id"
  datinstrumento.Refresh
  
fuera:
End Sub

Private Sub borrar_Click()
On Error GoTo erroreliminar

    nrocompro(DataGrid1.Row) = ""
    databonan.Recordset.Delete adAffectCurrent
    databonan.Refresh
Exit Sub

erroreliminar:
Rem MsgBox "No se pudo eliminar Concepto"
End Sub

Private Sub Command1_Click()
On Error GoTo fuera

    If DataGrid2.Row + 2 = 16 Then
        msnsa = MsgBox("No puede ingresar mas formas de pago", vbExclamation, "Atencion")
        Exit Sub
    End If

    codigopago = 0
    DataGrid2.Enabled = True
    datinstrumento.Recordset.AddNew
    datinstrumento.Recordset.Fields(0) = text1.Text
    datinstrumento.Recordset.Fields(1) = login.empresaact
    datinstrumento.Recordset.Fields(2) = login.iper
    datinstrumento.Recordset.Fields(3) = login.fper
    DataGrid2.SetFocus
    DataGrid2.Col = 6

fuera:
End Sub

Private Sub Command1_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo fuera
    
    If KeyCode = 112 Then
        SSTab1.Tab = 0
        Call borrablancos_Click
        nuevo.SetFocus
    End If
    
fuera:
End Sub

Private Sub Command2_Click()
On Error GoTo erroreliminar1

    datinstrumento.Recordset.Delete adAffectCurrent
    totalinstrumento.Text = 0
    For X = 0 To DataGrid2.Row
            totalinstrumento.Text = Val(totalinstrumento.Text) + totalinst(X)
    Next X
erroreliminar1:
Rem MsgBox "No se pudo eliminar Instumento de Pago"

End Sub

Private Sub Command3_Click()
On Error GoTo fuera

    If totalinstrumento = "" Then
        mensa = MsgBox("Pago no Imputado", vbExclamation, "Atencion")
        Exit Sub
    End If
    diferencia = totalabonan - totalinstrumento
    If diferencia < 0 Then diferencia = diferencia * -1
    If diferencia > 0.009 Then
        mensa = MsgBox("El Asiento esta Desvalanceado, no se puede grabar", vbCritical, "!! Error !!")
        DataGrid2.SetFocus
        Exit Sub
    End If
    
    
    Call borrablancos_Click
    
    If codprove = 0 Then GoTo conti
    
    For X = 0 To DataGrid1.VisibleRows - 1
       datlibrocompras.RecordSource = "select librocompras.* from librocompras WHERE librocompras.empresa = " & login.empresaact & " and id = " & idlibrogrid(X) & " Order by id"
       datlibrocompras.Refresh
       If datlibrocompras.Recordset.EOF = True Then GoTo conti
       DataGrid4.Refresh
       DataGrid4.Columns(63).Value = saldolibro(X)
       DataGrid4.Columns(64).Value = "S"
       datlibrocompras.Recordset.UpdateBatch adAffectCurrent
    Next X
conti:
    datordendepago.Recordset.Fields(5) = "S"
    datordendepago.Recordset.Fields("anulado") = "N"
    datordendepago.Recordset.UpdateBatch adAffectCurrent
    numorden = text1.Text


Rem ****************** grabar asiento

    campoaño = Right(MaskEdBox1.Text, 4)
    campomes = Mid(MaskEdBox1.Text, 4, 2)
    campodia = Left(MaskEdBox1.Text, 2)
    campofecha = campoaño + "/" + campomes + "/" + campodia
       inicioperiodo = login.iper
    campoaño1 = Right(inicioperiodo, 4)
    campomes1 = Mid(inicioperiodo, 4, 2)
    campodia1 = Left(inicioperiodo, 2)
    campofecha1 = campoaño1 + "/" + campomes1 + "/" + campodia1
       finperiodo = login.fper
    campoaño2 = Right(finperiodo, 4)
    campomes2 = Mid(finperiodo, 4, 2)
    campodia2 = Left(finperiodo, 2)
    campofecha2 = campoaño2 + "/" + campomes2 + "/" + campodia2

    If campofecha < campofecha1 Or campofecha > campofecha2 Then
            mensa = MsgBox("La Fecha es erronea o no pertenecia al periodo en ejercicio", vbCritical, "!! Atención !!")
            MaskEdBox1.SelLength = 10
            MaskEdBox1.SetFocus
            Exit Sub
    End If

    If datmaestro.Recordset.EOF = False Then
            datmaestro.RecordSource = "select [Maestro Asientos].* from [Maestro Asientos] where empresa = " & login.empresaact & " and perinicial = '" & login.iper & "' order by nroasiento"
            datmaestro.Refresh
            datmaestro.Recordset.MoveLast
            nroasie = datmaestro.Recordset.Fields(3) + 1
    Else
            nroasie = 1
    End If
pas1:
    datmaestro.Recordset.AddNew
    datmaestro.Recordset.Fields(0) = MaskEdBox1.Text
    datmaestro.Recordset.Fields(1) = Date
    datmaestro.Recordset.Fields(3) = nroasie
    datmaestro.Recordset.Fields(4) = "Orden de Pago " + text1.Text + " a " + Left(DataGrid1.Columns(6).Text, 20)
    datmaestro.Recordset.Fields(5) = login.iper
    datmaestro.Recordset.Fields(6) = login.fper
    datmaestro.Recordset.Fields(7) = login.empresaact
    datmaestro.Recordset.Fields(8) = "N"
    datmaestro.Recordset.Fields(10) = "O"
    datmaestro.Recordset.Fields(11) = "S"
    datmaestro.Recordset.UpdateBatch adAffectCurrent
    datordendepago.Recordset.Fields("idasiento") = nroasie
    datordendepago.Recordset.UpdateBatch adAffectCurrent
   
  
    If codprove = 0 Then
        For X = 1 To DataGrid5.VisibleRows
            datasiento.Recordset.AddNew
            datasiento.Recordset.Fields(5) = Val(datmaestro.Recordset.Fields(2))
            datasiento.Recordset.Fields(1) = datmaestro.Recordset.Fields(0).Value
            datasiento.Recordset.Fields(7) = login.empresaact
            datasiento.Recordset.Fields(2) = cuentaotros(X - 1)
            datasiento.Recordset.Fields(3) = totalconc(X - 1)
            datasiento.Recordset.Fields(4) = 0
            datasiento.Recordset.Fields(6) = conceptosotros(X - 1)
            datasiento.Recordset.UpdateBatch adAffectCurrent
        Next X
        GoTo conti2
    End If
  
    datasiento.Recordset.AddNew
    datasiento.Recordset.Fields(5) = Val(datmaestro.Recordset.Fields(2))
    datasiento.Recordset.Fields(1) = datmaestro.Recordset.Fields(0).Value
    datasiento.Recordset.Fields(7) = login.empresaact
    datasiento.Recordset.Fields(2) = cuenta
    datasiento.Recordset.Fields(3) = totalab
    datasiento.Recordset.Fields(4) = 0
    datasiento.Recordset.Fields(6) = "Total Orden de Pago"
    datasiento.Recordset.UpdateBatch adAffectCurrent

conti2:
    For X = 1 To DataGrid2.VisibleRows
            datasiento.Recordset.AddNew
            datasiento.Recordset.Fields(5) = Val(datmaestro.Recordset.Fields(2))
            datasiento.Recordset.Fields(1) = datmaestro.Recordset.Fields(0).Value
            datasiento.Recordset.Fields(7) = login.empresaact
            datasiento.Recordset.Fields(2) = cuentaint(X - 1)
            datasiento.Recordset.Fields(3) = 0
            datasiento.Recordset.Fields(4) = totalinst(X - 1)
            datasiento.Recordset.Fields(6) = detalleint(X - 1)
            datasiento.Recordset.UpdateBatch adAffectCurrent
    Next X

    Inicio.datauditoria.ConnectionString = login.conexiontotal
    
    Inicio.datauditoria.RecordSource = "select auditoria.* from auditoria"
    Inicio.datauditoria.Refresh
    
    Inicio.datauditoria.Recordset.AddNew
    Inicio.datauditoria.Recordset.Fields("fecha") = Date
    Inicio.datauditoria.Recordset.Fields("hora") = Str(Time)
    Inicio.datauditoria.Recordset.Fields("ventana") = "Orden de Pago"
    Inicio.datauditoria.Recordset.Fields("accion") = "Emitida Orden: " + text1.Text
    Inicio.datauditoria.Recordset.Fields("usuario") = login.usuarioactivo
    Inicio.datauditoria.Recordset.Fields("empresa") = login.empresaact
    Inicio.datauditoria.Recordset.UpdateBatch adAffectCurrent

    Call Command4_Click
    Call nuevaorden_Click

fuera:
End Sub

Private Sub Command4_Click()
On Error GoTo fuera

Dim crApp As New CRAXDRT.Application
Dim crReport As New CRAXDRT.Report
Dim tabla As String
Dim tabla1 As String
Dim ruta As String

criterio.ConnectionString = login.conexiontotal

criterio.RecordSource = "select empreactiva.* from empreactiva"
criterio.Refresh

criterio.Recordset.Fields(0) = login.empresaact
criterio.Recordset.UpdateBatch adAffectCurrent


ruta = "\Empresa" + Right(Str(login.empresaact), Len(Str(login.empresaact)) - 1)

reporte.SQL = "consultaordesnpago.nrorden, consultaordesnpago.empresa, consultaordesnpago.nomproveedor, consultaordesnpago.comprobante, consultaordesnpago.fechacompro, consultaordesnpago.importe, consultaordesnpago.id, consultaordesnpago.razonsocial, consultaordesnpago.cuit, consultaordesnpago.domicilio, consultaordesnpago.localidad, consultaordesnpago.fecha, consultaordesnpago.domprov, consultaordesnpago.locprov, consultaordesnpago.cuitprov, consultaordesnpago.saldofactura FROM contablesql.dbo.consultaordesnpago consultaordesnpago WHERE consultaordesnpago.nrorden= '" & frmordendepago.numorden & "' and consultaordesnpago.empresa = " & login.empresaact & " ORDER BY consultaordesnpago.razonsocial ASC, consultaordesnpago.comprobante ASC"
tabla = reporte.SQL

With CrystalReporte
    .ReportFileName = App.Path & ruta + "\Ordendepago.rpt"
    .Connect = login.conexionreporte
    .SubreportToChange = .GetNthSubreportName(0)
    .Connect = login.conexionreporte
    .SubreportToChange = ""
    .Connect = login.conexionreporte
    .DiscardSavedData = True
    .RetrieveDataFiles
    .ReportSource = 0
    .SQLQuery = tabla
Rem   .Destination = crptToWindow
    .Destination = crptToFile
    .PrintFileType = crptCrystal
    .PrintFileName = App.Path & "\pagos.rpt"
    .Action = 1
If Check1(0).Value = 1 Then GoTo fin
    .ReportFileName = App.Path & ruta + "\Ordendepago1.rpt"
    .Connect = login.conexionreporte
    .SubreportToChange = .GetNthSubreportName(0)
    .Connect = login.conexionreporte
    .SubreportToChange = ""
    .Connect = login.conexionreporte
    .DiscardSavedData = True
    .RetrieveDataFiles
    .ReportSource = 0
    .SQLQuery = tabla
   .Destination = crptToWindow
 Rem   .Destination = crptToPrinter
    .Destination = crptToFile
    .PrintFileType = crptCrystal
    .PrintFileName = App.Path & "\pagos.rpt"
    .Action = 1
    
    
    
    
fin:
End With

Set crReport = crApp.OpenReport(App.Path & "\pagos.rpt", 1)
impresos.CRViewer1.ReportSource = crReport
impresos.CRViewer1.ViewReport

fuera:
End Sub

Private Sub Command5_Click()
On Error GoTo fuera

    DataGrid1.Columns(6).Caption = "Detalle de Pago"
    DataGrid1.Columns(7).Visible = False
    DataGrid1.Columns(8).Caption = "Fecha"
Rem    DataGrid1.Columns(10).Visible = True
    DataGrid1.Columns(10).Width = 1395
    DataGrid1.Columns(11).Locked = False
    sincomp = 1

fuera:
End Sub

Private Sub Command6_Click()

End Sub

Private Sub DataGrid1_AfterColEdit(ByVal ColIndex As Integer)
On Error GoTo errorcarga
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
        If datlibrocompras.Recordset.EOF = False Then
            idlibrogrid(DataGrid1.Row) = DataGrid4.Columns(0).Text
            If sincomp = 0 Then DataGrid1.Columns(8).Value = DataGrid4.Columns(2).Value
        End If
 Rem       DataGrid4.Columns(63).Value = saldoactual
 Rem       DataGrid4.Columns(64).Value = "S"
        DataGrid1.Columns(11).Value = saldoactual

        totalconc(DataGrid1.Row) = importeapagar
        If datlibrocompras.Recordset.EOF = False Then
            datlibrocompras.Recordset.UpdateBatch adAffectCurrent
        End If
        datordendepago.Recordset.UpdateBatch adAffectCurrent
        databonan.Recordset.UpdateBatch adAffectCurrent
        DataGrid1.Refresh
        totalab = 0
        For X = 0 To DataGrid1.Row
            totalab = totalab + totalconc(X)
        Next X
        totalabonan.Text = totalab
        Exit Sub
    End If
errorcarga:
End Sub

Private Sub DataGrid1_ButtonClick(ByVal ColIndex As Integer)
On Error GoTo fuera

If ColIndex = 6 Then
        DataList1.Visible = True
        DataList1.Left = DataGrid1.Columns(6).Left + DataGrid1.Left + SSTab1.Left
        DataList1.Width = DataGrid1.Columns(6).Width
        DataList1.Top = DataGrid1.Top + DataGrid1.RowTop(DataGrid1.Row) + DataGrid1.RowHeight * 6
        DataList1.SetFocus
End If

If ColIndex = 7 Then
        DataList2.Visible = True
        DataList2.Left = DataGrid1.Columns(7).Left + DataGrid1.Left + SSTab1.Left
        DataList2.Width = DataGrid1.Columns(7).Width
        DataList2.Top = DataGrid1.Top + DataGrid1.RowTop(DataGrid1.Row) + DataGrid1.RowHeight * 6
        DataList2.SetFocus
End If
    
fuera:
End Sub


Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
On Error GoTo erroringreso

    If KeyAscii = 13 And DataGrid1.Col = 6 Then
        If DataGrid1.Columns(6).Text = "" And sincomp = 0 Then
                    DataList1.Visible = True
                    DataList1.Left = DataGrid1.Columns(6).Left + DataGrid1.Left + SSTab1.Left
                    DataList1.Width = DataGrid1.Columns(6).Width
                    DataList1.Top = DataGrid1.Top + DataGrid1.RowTop(DataGrid1.Row) + DataGrid1.RowHeight * 6
                    KeyAscii = 0
                    DataList1.SetFocus
                    Exit Sub
        Else
            KeyAscii = 9
        End If
    End If
    
    If KeyAscii = 13 And DataGrid1.Col = 7 Then
        If DataGrid1.Columns(7).Text = "" Then
                    DataList2.Visible = True
                    DataList2.Left = DataGrid1.Columns(7).Left + DataGrid1.Left + SSTab1.Left
                    DataList2.Width = DataGrid1.Columns(7).Width
                    DataList2.Top = DataGrid1.Top + DataGrid1.RowTop(DataGrid1.Row) + DataGrid1.RowHeight * 6
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
    If KeyAscii = 13 And DataGrid1.Col = 9 And sincomp = 0 Then
        KeyAscii = 0
        Call nuevo_Click
    End If
    If KeyAscii = 13 And DataGrid1.Col = 10 Then
        If DataGrid1.Columns(10).Text = "" Then
                    DataList4.Visible = True
                    DataList4.Width = DataGrid1.Columns(10).Width * 2.5
                    DataList4.Left = DataGrid1.Columns(10).Left + DataGrid1.Left + DataGrid2.Columns(10).Width - DataList4.Width
                    DataList4.Top = DataGrid1.Top + DataGrid1.RowTop(DataGrid1.Row) + DataGrid1.RowHeight * 3
                    codigopago = 1
                    DataList4.SetFocus
                    KeyAscii = 0
                    Exit Sub
        Else
            KeyAscii = 9
        End If
    End If
    If KeyAscii = 13 And DataGrid1.Col = 11 And sincomp = 1 Then
        KeyAscii = 0
        Call nuevo_Click
    End If
Exit Sub
erroringreso:
    nuevo.SetFocus



End Sub

Private Sub DataGrid1_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo fuera

    If KeyCode = 113 Then
        SSTab1.Tab = 1
        Call borrablancos_Click
        Command1.SetFocus
    End If
    
    If KeyCode = 114 Then
        SSTab1.Tab = 2
        Call borrablancos_Click
        nuevootroconcepto.SetFocus
    End If

fuera:
End Sub

Private Sub DataGrid2_AfterColUpdate(ByVal ColIndex As Integer)
    
    If DataGrid2.Col = 11 And DataGrid2.Columns(6).Text = "EFECTIVO" Then
            If Inicio.montoefectivo <> 0 And DataGrid2.Columns(11).Value > Inicio.montoefectivo Then
                mensa = MsgBox("El Monto en efectivo a pagar es mayor a $" + Str(Inicio.montoefectivo) + " Verifique", vbExclamation, "!! Atención !!")
            End If
    End If

End Sub

Private Sub DataGrid2_ButtonClick(ByVal ColIndex As Integer)
On Error GoTo fuera

If ColIndex = 6 Then
        DataList3.Visible = True
        DataList3.Left = DataGrid2.Columns(6).Left + DataGrid2.Left + SSTab1.Left
        DataList3.Width = DataGrid2.Columns(6).Width
        DataList3.Top = DataGrid2.Top + DataGrid2.RowTop(DataGrid2.Row) + DataGrid2.RowHeight * 6
        DataList3.SetFocus
End If

If ColIndex = 12 Then
        DataList4.Visible = True
        DataList4.Width = DataGrid2.Columns(12).Width * 4
        DataList4.Left = DataGrid2.Columns(12).Left + DataGrid2.Left + DataGrid2.Columns(12).Width - DataList4.Width + SSTab1.Left
        DataList4.Top = DataGrid2.Top + DataGrid2.RowTop(DataGrid2.Row) + DataGrid2.RowHeight * 6
        DataList4.SetFocus
End If

fuera:
End Sub

Private Sub DataGrid2_KeyPress(KeyAscii As Integer)
On Error GoTo erroringreso1



    If KeyAscii = 13 And DataGrid2.Col = 6 Then
        If DataGrid2.Columns(6).Text = "" Then
                    DataList3.Visible = True
                    DataList3.Left = DataGrid2.Columns(6).Left + DataGrid2.Left + SSTab1.Left
                    DataList3.Width = DataGrid2.Columns(6).Width
                    DataList3.Top = DataGrid2.Top + DataGrid2.RowTop(DataGrid2.Row) + DataGrid2.RowHeight * 6
                    DataList3.SetFocus
                    KeyAscii = 0
                    Exit Sub
        Else
            KeyAscii = 9
        End If
    End If
    
    If KeyAscii = 13 And DataGrid2.Col = 12 Then
        If DataGrid2.Columns(12).Text = "" Then
                    DataList4.Visible = True
                    DataList4.Width = DataGrid2.Columns(12).Width * 4
                    DataList4.Left = DataGrid2.Columns(12).Left + DataGrid2.Left + DataGrid2.Columns(12).Width - DataList4.Width + SSTab1.Left
                    DataList4.Top = DataGrid2.Top + DataGrid2.RowTop(DataGrid2.Row) + DataGrid2.RowHeight * 6
                    DataList4.SetFocus
                    KeyAscii = 0
                    Exit Sub
        End If
    End If
    If KeyAscii = 13 And DataGrid2.Col = 12 Then
        If DataGrid2.Columns(12).Text <> "" Then
            totalinst(DataGrid2.Row) = DataGrid2.Columns(11).Text
            cuentaint(DataGrid2.Row) = DataGrid2.Columns(12).Text
            detalleint(DataGrid2.Row) = DataGrid2.Columns(6).Text + " " + DataGrid2.Columns(7).Text
            datinstrumento.Recordset.UpdateBatch adAffectCurrent
                
            totalinstrumento.Text = 0
            For X = 0 To DataGrid2.Row
                totalinstrumento.Text = Val(totalinstrumento.Text) + totalinst(X)
            Next X
            saldototal.Text = totalinstrumento.Text - totalabonan.Text
        End If
        Command1.SetFocus
    End If
    
                    
    If KeyAscii = 13 Then KeyAscii = 9
Exit Sub
erroringreso1:
    Call Command1.SetFocus


End Sub

Private Sub DataGrid2_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo fuera

    If KeyCode = 115 Then frmordenintrumento.Show

    If KeyCode = 112 Then
        SSTab1.Tab = 0
        Call borrablancos_Click
        nuevo.SetFocus
    End If

fuera:
End Sub

Private Sub DataGrid5_ButtonClick(ByVal ColIndex As Integer)
On Error GoTo fuera

If ColIndex = 10 Then
        DataList4.Visible = True
        DataList4.Width = DataGrid5.Columns(10).Width * 4
        DataList4.Left = DataGrid5.Columns(10).Left + DataGrid5.Left + DataGrid5.Columns(10).Width - DataList4.Width + SSTab1.Left
        DataList4.Top = DataGrid5.Top + DataGrid5.RowTop(DataGrid5.Row) + DataGrid5.RowHeight * 6
        DataList4.SetFocus
End If

fuera:
End Sub

Private Sub DataGrid5_KeyPress(KeyAscii As Integer)
On Error GoTo fuera


    If KeyAscii = 13 And DataGrid5.Col = 10 Then
        If DataGrid5.Columns(10).Text = "" Then
                    DataList4.Visible = True
                    DataList4.Width = DataGrid5.Columns(10).Width * 4
                    DataList4.Left = DataGrid5.Columns(10).Left + DataGrid5.Left + DataGrid5.Columns(10).Width - DataList4.Width + SSTab1.Left
                    DataList4.Top = DataGrid5.Top + DataGrid5.RowTop(DataGrid5.Row) + DataGrid5.RowHeight * 6
                    DataList4.SetFocus
                    KeyAscii = 0
                    Exit Sub
        End If
    End If
    
    If KeyAscii = 13 And DataGrid5.Col = 10 Then
        KeyAscii = 0
        totalconc(DataGrid5.Row) = DataGrid5.Columns(9).Value
        datordendepago.Recordset.UpdateBatch adAffectCurrent
        databonan.Recordset.UpdateBatch adAffectCurrent
        cuentaotros(DataGrid5.Row) = databonan.Recordset.Fields("codcuenta")
        conceptosotros(DataGrid5.Row) = databonan.Recordset.Fields("nomproveedor")
        DataGrid5.Refresh
        totalab = 0
        For X = 0 To DataGrid5.Row
            totalab = totalab + totalconc(X)
        Next X
        totalabonan.Text = totalab
        Call nuevootroconcepto_Click
        Exit Sub
    End If

    If KeyAscii = 13 Then
        KeyAscii = 0
        KeyAscii = 9
    End If

fuera:
End Sub

Private Sub DataGrid5_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo fuera
    
    If KeyCode = 113 Then
        SSTab1.Tab = 1
        Call borrablancos_Click
        Command1.SetFocus
    End If
    
    If KeyCode = 112 Then
        SSTab1.Tab = 0
        Call borrablancos_Click
        nuevo.SetFocus
    End If

fuera:
End Sub

Private Sub DataList1_KeyPress(KeyAscii As Integer)
On Error GoTo errorcarga
    If KeyAscii = 13 Then
        KeyPress = 0
            DataGrid1.Columns(6).Text = DataList1.Text
            DataGrid1.Columns(5).Text = DataList1.BoundText
            razonsocial1 = DataList1.Text
            codprove = DataList1.BoundText
            DataList1.Visible = False
            datconsultacomp.RecordSource = "select conceptosabonan.* from conceptosabonan where empresa = " & login.empresaact & " and codproveedor = " & codprove & " order by comp"
            datconsultacomp.Refresh
            If datconsultacomp.Recordset.EOF = False Then
                       DataGrid1.Columns(10).Text = DataGrid3.Columns(10).Text
                       cuenta = DataGrid1.Columns(10).Text
            Else
                datproveedores.RecordSource = "select proveedores.* from proveedores where empresa = " & login.empresaact & " and codproveedor = " & codprove & " "
                datproveedores.Refresh
                If IsNull(datproveedores.Recordset.Fields("codcontable")) = True Then GoTo errorcarga
                cuenta = datproveedores.Recordset.Fields("codcontable")
                datproveedores.RecordSource = "select proveedores.* from proveedores where empresa = " & login.empresaact & " ORDER BY razonsocial"
                datproveedores.Refresh
            End If
            
            DataGrid1.SetFocus
            DataGrid1.Col = 7
    End If
Exit Sub

errorcarga:
    mensa = MsgBox("El Proveedor no tiene cargado un codigo contable, verifique el archivo de Proveedores", vbCritical, "Error")
    Exit Sub
End Sub

Private Sub DataList1_LostFocus()
        DataList1.Visible = False
End Sub

Private Sub DataList2_KeyPress(KeyAscii As Integer)
On Error GoTo fuera
    
    If KeyAscii = 13 Then
            DataGrid1.Columns(7).Text = DataList2.Text
            compro = DataList2.Text
            If compro = "" Then Exit Sub
  
            nrocompro(DataGrid1.Row) = DataList2.Text
            nomprov(DataGrid1.Row) = DataGrid1.Columns(6).Text
            If DataGrid1.Row > 0 Then
                For X = 0 To DataGrid1.Row - 1
                    If nrocompro(X) = DataList2.Text And nomprov(X) = DataGrid1.Columns(6).Text Then
                        mensa = MsgBox("Este comprobante ya fue ingresado", vbCritical, "Error")
                        DataGrid1.Columns(7) = ""
                        DataGrid1.SetFocus
                        Exit Sub
                    End If
                Next X
            End If
                    
            
            datconsultacomp.RecordSource = "select conceptosabonan.* from conceptosabonan where empresa = " & login.empresaact & " and codproveedor = " & codprove & " and comp = '" & compro & "'  order by comp "
            datconsultacomp.Refresh
            idlibro = DataGrid3.Columns(11).Text
            DataGrid1.Columns(8).Text = DataGrid3.Columns(1).Text
            DataGrid1.Columns(10).Text = DataGrid3.Columns(10).Text
            cuenta = DataGrid1.Columns(10).Text
            If DataGrid3.Columns(12).Text = "" Or DataGrid3.Columns(12).Text = "N" Then
                DataGrid1.Columns(11).Text = DataGrid3.Columns(6).Text
            Else
                DataGrid1.Columns(11).Text = DataGrid3.Columns(9).Text
            End If
            DataGrid1.Columns(9).Text = "0"
            DataList2.Visible = False
            datlibrocompras.RecordSource = "select librocompras.* from librocompras WHERE librocompras.empresa = " & login.empresaact & " and id = " & idlibro & " Order by id"
            datlibrocompras.Refresh
            DataGrid1.SetFocus
    End If
    
fuera:
End Sub

Private Sub DataList2_LostFocus()
            DataList2.Visible = False
End Sub

Private Sub DataList3_KeyPress(KeyAscii As Integer)
On Error GoTo fuera
    
    If KeyAscii = 13 Then
            DataGrid2.Columns(6).Text = DataList3.Text
            DataList3.Visible = False
            DataGrid2.SetFocus
    End If
    
fuera:
End Sub

Private Sub DataList3_LostFocus()
    DataList3.Visible = False
End Sub

Private Sub DataList4_GotFocus()
On Error GoTo fuera
    
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
    
fuera:
End Sub

Private Sub DataList4_KeyPress(KeyAscii As Integer)
On Error GoTo fuera

    If KeyAscii = 13 And SSTab1.Tab = 2 Then
            DataGrid5.Columns(10).Text = DataList4.BoundText
            DataList4.Visible = False
            DataGrid5.SetFocus
            databonan.Recordset.UpdateBatch adAffectCurrent
            Exit Sub
    End If



    If KeyAscii = 13 And SSTab1.Tab = 1 Then
            texto = DataList4.BoundText
            DataList4.Visible = False
            DataGrid2.SetFocus
    End If
    If KeyAscii = 13 And codigopago = 1 Then
            DataGrid1.Columns(10).Text = DataList4.BoundText
            DataList4.Visible = False
            DataGrid1.SetFocus
            databonan.Recordset.UpdateBatch adAffectCurrent
    End If
fuera:
End Sub

Private Sub eliminaotroconcepto_Click()
On Error GoTo erroreliminar

    nrocompro(DataGrid5.Row) = ""
    databonan.Recordset.Delete adAffectCurrent
    databonan.Refresh
Exit Sub

erroreliminar:
Rem MsgBox "No se pudo eliminar Concepto"
End Sub

Private Sub Form_Load()

frmordendepago.Top = 0
frmordendepago.Left = 0

databonan.ConnectionString = login.conexiontotal
datasiento.ConnectionString = login.conexiontotal
datconsultacomp.ConnectionString = login.conexiontotal
datcuentas.ConnectionString = login.conexiontotal
datinstru.ConnectionString = login.conexiontotal
datinstrumento.ConnectionString = login.conexiontotal
datlibrocompras.ConnectionString = login.conexiontotal
datmaestro.ConnectionString = login.conexiontotal
datordendepago.ConnectionString = login.conexiontotal
datproveedores.ConnectionString = login.conexiontotal

If login.empresaact1 > 0 Then
    empresareal = login.empresaact1
Else
    empresareal = login.empresaact
End If

    DataGrid1.Columns(6).Width = 2500
    DataGrid1.Columns(7).Width = 2000
    DataGrid1.Columns(8).Width = 1400
    DataGrid1.Columns(9).Width = 1500
    DataGrid1.Columns(11).Width = 1500


    DataGrid1.Columns(9).NumberFormat = "#,##0.00"
    DataGrid1.Columns(11).NumberFormat = "#,##0.00"

    Inicio.Toolbar1.Visible = True
    
    sincomp = 0
    DataGrid1.Enabled = False
    DataGrid2.Enabled = False


  datmaestro.RecordSource = "select [Maestro Asientos].* from [Maestro Asientos] where empresa = " & login.empresaact & " and perinicial = '" & login.iper & "' order by nroasiento"
  datmaestro.Refresh

  datinstru.RecordSource = "select instrumentospagos.* from instrumentospagos"
  datinstru.Refresh

  datasiento.RecordSource = "select [Detalle Asientos].* from [Detalle Asientos] where empresa = " & login.empresaact & ""
  datasiento.Refresh
  datordendepago.RecordSource = "select ordendepago.* from ordendepago WHERE ordendepago.empresa = " & login.empresaact & " Order by nrorden"
  datordendepago.Refresh
  
  datproveedores.RecordSource = "select proveedores.* from proveedores where empresa = " & empresareal & " ORDER BY razonsocial"
  datproveedores.Refresh
  
  datlibrocompras.RecordSource = "select librocompras.* from librocompras WHERE librocompras.empresa = " & login.empresaact & " Order by id"
  datlibrocompras.Refresh
  
  datcuentas.RecordSource = "select listacuentas.* from listacuentas WHERE empre = " & login.empresaact & " and inicioper = '" & login.iper & "' and finper = '" & login.fper & "' ORDER BY IDCUENTA"
  datcuentas.Refresh
  
  MaskEdBox1.Text = Date
  totalab = 0
  Check1(0).Value = 1
  SSTab1.Tab = 0
  
  If datordendepago.Recordset.EOF = True Then
        datordendepago.Recordset.AddNew
        text1.Text = "00000001"
        text1.Enabled = True
  Else
        datordendepago.Recordset.MoveLast
        nroorden = datordendepago.Recordset.Fields(0)
        pruebaorden = IsNull(datordendepago.Recordset.Fields(5))
        If pruebaorden = True Then
            text1.Text = nroorden
            GoTo paso0
        End If
        previo = Str(Val((Right(nroorden, 8))) + 1)
        previo1 = Right(previo, Len(previo) - 1)
        mitad2 = Mid("00000000", 1, 8 - Len(previo1)) + previo1
        datordendepago.Recordset.AddNew
        text1.Text = mitad2
paso0:
        text1.Enabled = False
  End If
   
  MaskEdBox1.Mask = "##/##/####"
  datordendepago.Recordset.Fields(0) = text1.Text
  datordendepago.Recordset.Fields(1) = login.empresaact
  datordendepago.Recordset.Fields(2) = login.iper
  datordendepago.Recordset.Fields(3) = login.fper
  datordendepago.Recordset.Fields(4) = MaskEdBox1.Text
  orden = text1.Text
  
  databonan.RecordSource = "select ordendepagoabonan.* from ordendepagoabonan WHERE ordendepagoabonan.empresa = " & login.empresaact & " and inicioper = '" & login.iper & "' and finper = '" & login.fper & "' and nrorden = '" & orden & "' Order by fechacompro"
  databonan.Refresh
  
  datinstrumento.RecordSource = "select ordendepagoinstrumento.* from ordendepagoinstrumento WHERE ordendepagoinstrumento.empresa = " & login.empresaact & " and inicioper = '" & login.iper & "' and finper = '" & login.fper & "' and nrorden = '" & orden & "' Order by id"
  datinstrumento.Refresh
  

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Inicio.Toolbar1.Visible = False

End Sub

Private Sub MaskEdBox1_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        nuevo.SetFocus
    End If
    
End Sub

Private Sub nuevaorden_Click()
   
Rem    totalabonan.Text = "0.00"
Rem    totalinstrumento.Text = "0.00"
Rem    saldototal.Text = "0.00"
Rem    DataGrid1.Columns(6).Caption = "Proveedor Razon Social"
Rem    DataGrid1.Columns(7).Visible = True
Rem    DataGrid1.Columns(8).Caption = "Fecha Compr."
Rem    DataGrid1.Columns(10).Visible = False
Rem    DataGrid1.Columns(10).Width = 1395
Rem    DataGrid1.Columns(11).Locked = True
Rem    sincomp = 0
    Unload Me
    frmordendepago.Show
    impresos.Show

End Sub

Private Sub nuevo_Click()
On Error GoTo fuera

    If DataGrid1.Row + 2 = 16 Then
        msnsa = MsgBox("No puede ingresar mas comprobantes, cree una nueva orden de pago", vbExclamation, "Atencion")
        Exit Sub
    End If
    
    DataGrid1.Enabled = True
    databonan.Recordset.AddNew
    databonan.Recordset.Fields(0) = text1.Text
    databonan.Recordset.Fields(1) = login.empresaact
    databonan.Recordset.Fields(2) = login.iper
    databonan.Recordset.Fields(3) = login.fper
    DataGrid1.SetFocus
    DataGrid1.Col = 6
    
fuera:
End Sub

Private Sub nuevo_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo fuera

    If KeyCode = 113 Then
        SSTab1.Tab = 1
        Call borrablancos_Click
        Command1.SetFocus
    End If
    If KeyCode = 114 Then
        SSTab1.Tab = 2
        Call borrablancos_Click
        nuevootroconcepto.SetFocus
    End If
    
fuera:
End Sub

Private Sub nuevootroconcepto_Click()
On Error GoTo fuera


    If DataGrid5.Row + 2 = 16 Then
        msnsa = MsgBox("No puede ingresar mas conceptos, cree una nueva orden de pago", vbExclamation, "Atencion")
        Exit Sub
    End If
    
    DataGrid5.Enabled = True
    databonan.Recordset.AddNew
    databonan.Recordset.Fields(0) = text1.Text
    databonan.Recordset.Fields(1) = login.empresaact
    databonan.Recordset.Fields(2) = login.iper
    databonan.Recordset.Fields(3) = login.fper
    DataGrid5.Col = 6
    DataGrid5.SetFocus
 
fuera:
 
End Sub

Private Sub salir_Click()

    Unload Me

End Sub

Private Sub SSTab1_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo fuera
    
    If KeyCode = 113 And (SSTab1.Tab = 0 Or SSTab1.Tab = 2) Then
        SSTab1.Tab = 1
        Call borrablancos_Click
        Command1.SetFocus
    End If
    
    If KeyCode = 112 And (SSTab1.Tab = 1 Or SSTab1.Tab = 2) Then
        SSTab1.Tab = 0
        Call borrablancos_Click
        nuevo.SetFocus
    End If
    If KeyCode = 114 And (SSTab1.Tab = 1 Or SSTab1.Tab = 0) Then
        SSTab1.Tab = 2
        Call borrablancos_Click
        nuevootroconcepto.SetFocus
    End If

fuera:
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
On Error GoTo fuera

    If KeyAscii = 13 Then
            If Val(Mid(text1.Text, 1, 4)) = 0 Then
                mensa = MsgBox("Debe ingresar una sucursal en el Nro de factura", vbCritical, "!! Atención !!")
                text1.SetFocus
                text1.SelStart = 0
                text1.SelLength = 4
                Exit Sub
            End If
            If Right(text1.Text, 1) = "_" Then
                mensa = MsgBox("Nro de factura incorrecto", vbCritical, "!! Atención !!")
                text1.SetFocus
                text1.SelStart = 5
                text1.SelLength = 8
                Exit Sub
            End If
    End If

fuera:
End Sub

