VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{5734474E-78D3-4254-99B9-C35F31BDF509}#62.0#0"; "vbskpro2.ocx"
Begin VB.Form importalibroventas 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8430
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8250
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8430
   ScaleWidth      =   8250
   Begin VB.CommandButton genansiento 
      Caption         =   "&Generar Asiento"
      Height          =   855
      Left            =   6840
      Picture         =   "importalibroventas.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   6720
      UseMaskColor    =   -1  'True
      Width           =   1095
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   3840
      TabIndex        =   29
      Top             =   6720
      Width           =   1695
   End
   Begin VB.CommandButton modificar 
      Caption         =   "&Calcular"
      Height          =   855
      Left            =   5640
      Picture         =   "importalibroventas.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   6720
      UseMaskColor    =   -1  'True
      Width           =   1095
   End
   Begin VB.CommandButton importar 
      Caption         =   "&Importar"
      Height          =   735
      Left            =   5880
      Picture         =   "importalibroventas.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   360
      UseMaskColor    =   -1  'True
      Width           =   975
   End
   Begin VB.CommandButton aceptar 
      Caption         =   "&Aceptar"
      Height          =   735
      Left            =   3360
      Picture         =   "importalibroventas.frx":0CC6
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   360
      UseMaskColor    =   -1  'True
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   5
      Left            =   1440
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   6000
      Width           =   735
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   255
      Left            =   1800
      TabIndex        =   1
      Top             =   840
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   450
      _Version        =   393216
      CalendarBackColor=   12648447
      Format          =   114229249
      CurrentDate     =   38799
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   255
      Left            =   1800
      TabIndex        =   0
      Top             =   360
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   450
      _Version        =   393216
      CalendarBackColor=   12648447
      Format          =   114229249
      CurrentDate     =   38799
   End
   Begin MSMask.MaskEdBox criteriofecha 
      Height          =   255
      Index           =   0
      Left            =   6600
      TabIndex        =   9
      Top             =   360
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   450
      _Version        =   393216
      MaxLength       =   10
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   495
      Left            =   120
      TabIndex        =   8
      Top             =   7920
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   873
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSDataGridLib.DataGrid DataGrid4 
      Bindings        =   "importalibroventas.frx":1108
      Height          =   4575
      Left            =   120
      TabIndex        =   7
      Top             =   1320
      Visible         =   0   'False
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   8070
      _Version        =   393216
      AllowUpdate     =   -1  'True
      AllowArrows     =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      TabAcrossSplits =   -1  'True
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
      ColumnCount     =   75
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
         Caption         =   "Fecha"
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
         DataField       =   "cliente"
         Caption         =   "cliente"
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
         Caption         =   "Compr."
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
         Caption         =   "N° Comprobante"
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
         Caption         =   "Neto"
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
         Caption         =   "Exento"
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
         Caption         =   "IVA"
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
         Caption         =   "TOTAL"
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
         DataField       =   "inicioper"
         Caption         =   "inicioper"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column61 
         DataField       =   "finper"
         Caption         =   "finper"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column62 
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
      BeginProperty Column63 
         DataField       =   "contado"
         Caption         =   "contado"
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
         DataField       =   "nombretarjeta"
         Caption         =   "nombretarjeta"
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
      BeginProperty Column65 
         DataField       =   "codoperacion"
         Caption         =   "codoperacion"
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
      BeginProperty Column66 
         DataField       =   "numordenpub"
         Caption         =   "numordenpub"
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
      BeginProperty Column67 
         DataField       =   "avisador"
         Caption         =   "avisador"
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
      BeginProperty Column68 
         DataField       =   "producto"
         Caption         =   "producto"
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
      BeginProperty Column69 
         DataField       =   "telefono"
         Caption         =   "telefono"
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
      BeginProperty Column70 
         DataField       =   "domicilio"
         Caption         =   "domicilio"
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
      BeginProperty Column71 
         DataField       =   "localidad"
         Caption         =   "localidad"
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
      BeginProperty Column72 
         DataField       =   "numletras"
         Caption         =   "numletras"
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
      BeginProperty Column73 
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
      BeginProperty Column74 
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
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column01 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column02 
            Alignment       =   2
         EndProperty
         BeginProperty Column03 
            Object.Visible         =   -1  'True
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
         EndProperty
         BeginProperty Column08 
            Alignment       =   1
         EndProperty
         BeginProperty Column09 
            Alignment       =   1
         EndProperty
         BeginProperty Column10 
            Alignment       =   1
         EndProperty
         BeginProperty Column11 
            Alignment       =   1
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
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column17 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column18 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column19 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column20 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column21 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column22 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column23 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column24 
            Alignment       =   1
         EndProperty
         BeginProperty Column25 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column26 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column27 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column28 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column29 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column30 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column31 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column32 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column33 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column34 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column35 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column36 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column37 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column38 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column39 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column40 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column41 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column42 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column43 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column44 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column45 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column46 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column47 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column48 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column49 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column50 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column51 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column52 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column53 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column54 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column55 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column56 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column57 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column58 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column59 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column60 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column61 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column62 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column63 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column64 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column65 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column66 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column67 
            Object.Visible         =   -1  'True
         EndProperty
         BeginProperty Column68 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column69 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column70 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column71 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column72 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column73 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column74 
            Object.Visible         =   0   'False
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DataGrid2 
      Bindings        =   "importalibroventas.frx":1123
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      TabAcrossSplits =   -1  'True
      TabAction       =   2
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
      ColumnCount     =   19
      BeginProperty Column00 
         DataField       =   "añolectivo"
         Caption         =   "Año Lec"
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
         DataField       =   "legajopagos"
         Caption         =   "Legajo"
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
         DataField       =   "fechapago"
         Caption         =   "fechapago"
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
         DataField       =   "matricula"
         Caption         =   "matricula"
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
         DataField       =   "cuotanro"
         Caption         =   "cuotanro"
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
         DataField       =   "mes"
         Caption         =   "mes"
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
         DataField       =   "importereal"
         Caption         =   "importereal"
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
         DataField       =   "importe"
         Caption         =   "Importe"
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
      BeginProperty Column09 
         DataField       =   "nrorecibo"
         Caption         =   "nrorecibo"
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
         DataField       =   "fecha"
         Caption         =   "Fecha"
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
         DataField       =   "hora"
         Caption         =   "hora"
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
         DataField       =   "usuario"
         Caption         =   "usuario"
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
         DataField       =   "anulado"
         Caption         =   "ANU"
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
         DataField       =   "bloqueado"
         Caption         =   "bloqueado"
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
         DataField       =   "fechaanul"
         Caption         =   "fechaanul"
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
         DataField       =   "horaanul"
         Caption         =   "horaanul"
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
         DataField       =   "usuarioanul"
         Caption         =   "usuarioanul"
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
         DataField       =   "nrorectexto"
         Caption         =   "N° Recibo"
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
            Alignment       =   2
         EndProperty
         BeginProperty Column01 
            Alignment       =   1
         EndProperty
         BeginProperty Column02 
            Alignment       =   2
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
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column07 
            Alignment       =   1
         EndProperty
         BeginProperty Column08 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column09 
            Alignment       =   1
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column10 
            Alignment       =   2
         EndProperty
         BeginProperty Column11 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column12 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column13 
            Alignment       =   2
         EndProperty
         BeginProperty Column14 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column15 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column16 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column17 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column18 
            Alignment       =   1
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc datpagos 
      Height          =   330
      Left            =   6720
      Top             =   0
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
      Connect         =   "Provider=MSDASQL.1;Password=25072004;Persist Security Info=True;User ID=lucva;Data Source=matriculas;Initial Catalog=matriculassql"
      OLEDBString     =   "Provider=MSDASQL.1;Password=25072004;Persist Security Info=True;User ID=lucva;Data Source=matriculas;Initial Catalog=matriculassql"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select [ingreso pagos].* from [ingreso pagos] order by nrorectexto"
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
   Begin MSMask.MaskEdBox criteriofecha 
      Height          =   255
      Index           =   1
      Left            =   6480
      TabIndex        =   10
      Top             =   720
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   450
      _Version        =   393216
      MaxLength       =   10
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSAdodcLib.Adodc datPrimaryRS 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      Top             =   8100
      Visible         =   0   'False
      Width           =   8250
      _ExtentX        =   14552
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
   Begin MSAdodcLib.Adodc datcolumnas 
      Height          =   330
      Left            =   5880
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
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
   Begin vbskpro.Skinner Skinner1 
      Left            =   4800
      Top             =   240
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
      LcK2            =   $"importalibroventas.frx":113A
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
   Begin MSAdodcLib.Adodc datalumnos 
      Height          =   330
      Left            =   6240
      Top             =   0
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
      Connect         =   "Provider=MSDASQL.1;Password=25072004;Persist Security Info=True;User ID=lucva;Data Source=matriculas;Initial Catalog=matriculassql"
      OLEDBString     =   "Provider=MSDASQL.1;Password=25072004;Persist Security Info=True;User ID=lucva;Data Source=matriculas;Initial Catalog=matriculassql"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select alumnos.* from alumnos"
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
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   1
      Left            =   1440
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   6360
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   2
      Left            =   1440
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   6720
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   3
      Left            =   1440
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   7080
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   4
      Left            =   1440
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   7440
      Width           =   735
   End
   Begin VB.PictureBox Text8 
      Height          =   255
      Index           =   5
      Left            =   2280
      ScaleHeight     =   195
      ScaleWidth      =   1155
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   6000
      Width           =   1215
   End
   Begin VB.PictureBox Text8 
      Height          =   255
      Index           =   1
      Left            =   2280
      ScaleHeight     =   195
      ScaleWidth      =   1155
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   6360
      Width           =   1215
   End
   Begin VB.PictureBox Text8 
      Height          =   255
      Index           =   2
      Left            =   2280
      ScaleHeight     =   195
      ScaleWidth      =   1155
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   6720
      Width           =   1215
   End
   Begin VB.PictureBox Text8 
      Height          =   255
      Index           =   3
      Left            =   2280
      ScaleHeight     =   195
      ScaleWidth      =   1155
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   7080
      Width           =   1215
   End
   Begin VB.PictureBox Text8 
      Height          =   255
      Index           =   4
      Left            =   2280
      ScaleHeight     =   195
      ScaleWidth      =   1155
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   7440
      Width           =   1215
   End
   Begin VB.PictureBox Text8 
      Height          =   255
      Index           =   0
      Left            =   6600
      ScaleHeight     =   195
      ScaleWidth      =   1395
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   6000
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "Seleccionar Periodo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   1680
      TabIndex        =   31
      Top             =   0
      Width           =   2895
   End
   Begin VB.Frame Frame2 
      Caption         =   "Período ya Importado"
      Height          =   1335
      Left            =   3720
      TabIndex        =   32
      Top             =   6360
      Width           =   4455
   End
   Begin VB.Label Label1 
      Caption         =   "Total Gral:"
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
      Index           =   7
      Left            =   5520
      TabIndex        =   21
      Top             =   6000
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Inicial y N.D.:"
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
      Index           =   6
      Left            =   120
      TabIndex        =   15
      Top             =   6000
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Universitarios:"
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
      Index           =   5
      Left            =   120
      TabIndex        =   14
      Top             =   7440
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Terciarios:"
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
      Index           =   4
      Left            =   120
      TabIndex        =   13
      Top             =   7080
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Secundarios:"
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
      Index           =   3
      Left            =   120
      TabIndex        =   12
      Top             =   6720
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Primarios:"
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
      Index           =   2
      Left            =   120
      TabIndex        =   11
      Top             =   6360
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Hasta Fecha:"
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
      Index           =   1
      Left            =   360
      TabIndex        =   6
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Desde Fecha:"
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
      Index           =   0
      Left            =   360
      TabIndex        =   5
      Top             =   360
      Width           =   1575
   End
End
Attribute VB_Name = "importalibroventas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim campomes As String

Private Sub Combo1_Change()

    campomes = Combo1.ListIndex - 1
    If Len(campomes) = 1 Then mescierre = "0" + campomes
    Call modificar_Click

End Sub

Private Sub Combo1_Click()

    campomes = Combo1.ListIndex + 1
    If Len(campomes) = 1 Then campomes = "0" + campomes
    Call modificar_Click
   
    
End Sub

Private Sub aceptar_Click()
    criteriofecha(0).Text = DTPicker1.Value
    criteriofecha(1).Text = DTPicker2.Value

    datpagos.RecordSource = "select [ingreso pagos].* from [ingreso pagos] where fecha >= '" & criteriofecha(0).Text & "' and fecha <= '" & criteriofecha(1).Text & "' order by nrorectexto"
    datpagos.Refresh
End Sub

Private Sub DataGrid4_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        datprimaryrs.Recordset.Fields("empresa") = 1
        datprimaryrs.Recordset.Fields("cliente") = "Consumidor Final"
        datprimaryrs.Recordset.Fields("tipoiva") = "CF"
        datprimaryrs.Recordset.Fields("cerrado") = campomes
        datprimaryrs.Recordset.Fields("inicioper") = login.iper
        datprimaryrs.Recordset.Fields("finper") = login.fper
        datprimaryrs.Recordset.UpdateBatch adAffectCurrent
        KeyAscii = 9
    End If

End Sub

Private Sub Form_Load()


datcolumnas.ConnectionString = login.conexiontotal
datprimaryrs.ConnectionString = login.conexiontotal


importalibroventas.Top = 0
importalibroventas.Left = 0

 Combo1.AddItem "ENERO"
 Combo1.AddItem "FEBRERO"
 Combo1.AddItem "MARZO"
 Combo1.AddItem "ABRIL"
 Combo1.AddItem "MAYO"
 Combo1.AddItem "JUNIO"
 Combo1.AddItem "JULIO"
 Combo1.AddItem "AGOSTO"
 Combo1.AddItem "SEPTIEMBRE"
 Combo1.AddItem "OCTUBRE"
 Combo1.AddItem "NOVIEMBRE"
 Combo1.AddItem "DICIEMBRE"
 
    importalibroventas.Left = 0
    importalibroventas.Top = 0
    
    criteriofecha(0).Text = Date - (Day(Date)) + 1
    criteriofecha(1).Text = Date
    mesini = Month(Date)

    datpagos.RecordSource = "select [ingreso pagos].* from [ingreso pagos] where fechapago >= '" & criteriofecha(0).Text & "' and fechapago <= '" & criteriofecha(1).Text & "' order by nrorectexto"
    datpagos.Refresh
   
    DataGrid2.Columns(7).NumberFormat = "##,##0.00"
    DataGrid2.Columns(0).Width = 800
    DataGrid2.Columns(1).Width = 800
    DataGrid2.Columns(7).Width = 1200
    DataGrid2.Columns(10).Width = 1000
    DataGrid2.Columns(13).Width = 600
    DataGrid2.Columns(18).Width = 1515
    
    DataGrid4.Columns(8).NumberFormat = "##,##0.00"
    DataGrid4.Columns(9).NumberFormat = "##,##0.00"
    DataGrid4.Columns(10).NumberFormat = "##,##0.00"
    DataGrid4.Columns(24).NumberFormat = "##,##0.00"
    
    DataGrid4.Columns(2).Width = 1000
    DataGrid4.Columns(6).Width = 800
    DataGrid4.Columns(7).Width = 1500
    DataGrid4.Columns(8).Width = 1000
    DataGrid4.Columns(9).Width = 1000
    DataGrid4.Columns(10).Width = 1000
    DataGrid4.Columns(24).Width = 1000
    

End Sub

Private Sub genansiento_Click()

    importalibroasientos.Show

End Sub

Private Sub importar_Click()
Dim cantidad(20) As Integer
Dim totales(20) As Currency
Dim cuentas(20) As Integer
    
    campoaño = Right(criteriofecha(0).Text, 4)
    campomes = Mid(criteriofecha(0).Text, 4, 2)
    campodia = Left(criteriofecha(0).Text, 2)
    campofecha = campoaño + "/" + campomes + "/" + campodia
    
    campoaño1 = Right(login.iper, 4)
    campomes1 = Mid(login.iper, 4, 2)
    campodia1 = Left(login.iper, 2)
    campofecha1 = campoaño1 + "/" + campomes1 + "/" + campodia1
    
    campoaño2 = Right(login.fper, 4)
    campomes2 = Mid(login.fper, 4, 2)
    campodia2 = Left(login.fper, 2)
    campofecha2 = campoaño2 + "/" + campomes2 + "/" + campodia2

    If campofecha < campofecha1 Or campofecha > campofecha2 Then
            mensa = MsgBox("La Fecha es erronea o no pertenecia al periodo en trabajo", vbCritical, "!! Atención !!")
            Exit Sub
    End If
For X = 1 To 5
    totales(X) = 0
    cantidad(X) = 0
Next X

    datprimaryrs.RecordSource = "select libroventas.* from libroventas where empresa = " & login.empresaact & " and inicioper = '" & login.iper & "' and cerrado = '" & campomes & "'"
    datprimaryrs.Refresh

    If datprimaryrs.Recordset.EOF = False Then
            mensa = MsgBox("Este periodo ya fue importado", vbCritical, "!! Atención !!")
            Exit Sub
    End If
    If datpagos.Recordset.EOF = True Then
        mensa = MsgBox("No hay pagos en este periodo", vbCritical, "! Atencion")
        Exit Sub
    End If
    datpagos.Recordset.MoveFirst
    ProgressBar1.Min = 0
    ProgressBar1.Max = datpagos.Recordset.RecordCount
    contador = 0
   
paso0:
    datprimaryrs.Recordset.AddNew
    contador = contador + 1
    datalumnos.RecordSource = "select alumnos.* from alumnos where legajo = " & datpagos.Recordset.Fields("legajopagos") & ""
    datalumnos.Refresh
    If datalumnos.Recordset.EOF = True Then GoTo paso000
    
    If Left(datalumnos.Recordset.Fields("nivel"), 8) = "Primario" Then
        totales(1) = totales(1) + datpagos.Recordset.Fields(7)
        cantidad(1) = cantidad(1) + 1
        Text1(1).Text = cantidad(1)
        Text8(1).Text = totales(1)
        GoTo paso00
    End If
    If Left(datalumnos.Recordset.Fields("nivel"), 10) = "Secundario" Then
        totales(2) = totales(2) + datpagos.Recordset.Fields(7)
        cantidad(2) = cantidad(2) + 1
        Text1(2).Text = cantidad(2)
        Text8(2).Text = totales(2)
        GoTo paso00
    End If
    If Left(datalumnos.Recordset.Fields("nivel"), 9) = "Terciario" Then
        totales(3) = totales(3) + datpagos.Recordset.Fields(7)
        cantidad(3) = cantidad(3) + 1
        Text1(3).Text = cantidad(3)
        Text8(3).Text = totales(3)
        GoTo paso00
    End If
    If Left(datalumnos.Recordset.Fields("nivel"), 13) = "Universitario" Then
        totales(4) = totales(4) + datpagos.Recordset.Fields(7)
        cantidad(4) = cantidad(4) + 1
        Text1(4).Text = cantidad(4)
        Text8(4).Text = totales(4)
        GoTo paso00
    End If
paso000:
        cantidad(5) = cantidad(5) + 1
        totales(5) = totales(5) + datpagos.Recordset.Fields(7)
        Text1(5).Text = cantidad(5)
        Text8(5).Text = totales(5)
paso00:
    datprimaryrs.Recordset.Fields(1) = login.empresaact
    datprimaryrs.Recordset.Fields(2) = datpagos.Recordset.Fields(10)
    datprimaryrs.Recordset.Fields(3) = "Consumidor Final"
    datprimaryrs.Recordset.Fields(4) = "CF"
    datprimaryrs.Recordset.Fields(5) = ""
    datprimaryrs.Recordset.Fields(6) = "R-C"
    datprimaryrs.Recordset.Fields(7) = datpagos.Recordset.Fields(18)
    datprimaryrs.Recordset.Fields(9) = datpagos.Recordset.Fields(7)
    datprimaryrs.Recordset.Fields(24) = datpagos.Recordset.Fields(7)
    If datpagos.Recordset.Fields("anulado") = -1 Then
        datprimaryrs.Recordset.Fields(3) = "***ANULADA***"
        datprimaryrs.Recordset.Fields(9) = 0
        datprimaryrs.Recordset.Fields(24) = 0
    End If
    datprimaryrs.Recordset.Fields("cerrado") = campomes
    datprimaryrs.Recordset.Fields(60) = login.iper
    datprimaryrs.Recordset.Fields(61) = login.fper
    If datalumnos.Recordset.EOF = False Then
            datprimaryrs.Recordset.Fields("avisador") = datalumnos.Recordset.Fields("nivel")
    Else
            datprimaryrs.Recordset.Fields("avisador") = ""
    End If
    datprimaryrs.Recordset.UpdateBatch adAffectCurrent
    datpagos.Recordset.MoveNext
    If datpagos.Recordset.EOF = True Then GoTo paso1
    ProgressBar1.Value = contador
    GoTo paso0
paso1:

    Text8(0).Text = Val(Text8(1).Text) + Val(Text8(2).Text) + Val(Text8(3).Text) + Val(Text8(4).Text) + Val(Text8(5).Text)
    Combo1.ListIndex = campomes - 1

End Sub

Private Sub modificar_Click()
Dim cantidad(20) As Integer
Dim totales(20) As Currency

    datprimaryrs.RecordSource = "select libroventas.* from libroventas where empresa = " & login.empresaact & " and inicioper = '" & login.iper & "' and cerrado = '" & campomes & "' order by numcompr"
    datprimaryrs.Refresh

    If datprimaryrs.Recordset.EOF = True Then GoTo paso4

For X = 1 To 5
    totales(X) = 0
    cantidad(X) = 0
Next X
    datprimaryrs.Recordset.MoveFirst
paso:
    If Left(datprimaryrs.Recordset.Fields("avisador"), 8) = "Primario" Then
        totales(1) = totales(1) + datprimaryrs.Recordset.Fields("total")
        cantidad(1) = cantidad(1) + 1
        Text1(1).Text = cantidad(1)
        Text8(1).Text = totales(1)
        GoTo paso00
    End If
    If Left(datprimaryrs.Recordset.Fields("avisador"), 10) = "Secundario" Then
        totales(2) = totales(2) + datprimaryrs.Recordset.Fields("total")
        cantidad(2) = cantidad(2) + 1
        Text1(2).Text = cantidad(2)
        Text8(2).Text = totales(2)
        GoTo paso00
    End If
    If Left(datprimaryrs.Recordset.Fields("avisador"), 9) = "Terciario" Then
        totales(3) = totales(3) + datprimaryrs.Recordset.Fields("total")
        cantidad(3) = cantidad(3) + 1
        Text1(3).Text = cantidad(3)
        Text8(3).Text = totales(3)
        GoTo paso00
    End If
    If Left(datprimaryrs.Recordset.Fields("avisador"), 13) = "Universitario" Then
        totales(4) = totales(4) + datprimaryrs.Recordset.Fields("total")
        cantidad(4) = cantidad(4) + 1
        Text1(4).Text = cantidad(4)
        Text8(4).Text = totales(4)
        GoTo paso00
    End If
    cantidad(5) = cantidad(5) + 1
    totales(5) = totales(5) + datprimaryrs.Recordset.Fields("total")
    Text1(5).Text = cantidad(5)
    Text8(5).Text = totales(5)
paso00:
    datprimaryrs.Recordset.MoveNext
    If datprimaryrs.Recordset.EOF = True Then GoTo paso4
    GoTo paso
paso4:
    Text8(0).Text = Val(Text8(1).Text) + Val(Text8(2).Text) + Val(Text8(3).Text) + Val(Text8(4).Text) + Val(Text8(5).Text)
    DataGrid4.Visible = True

End Sub
