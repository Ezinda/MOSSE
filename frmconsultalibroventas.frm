VERSION 5.00
Object = "{F6125AB1-8AB1-11CE-A77F-08002B2F4E98}#2.0#0"; "MSRDC20.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{5734474E-78D3-4254-99B9-C35F31BDF509}#62.0#0"; "vbskpro2.ocx"
Begin VB.Form frmconsutalibroventas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Consulta Facturas de Ventas"
   ClientHeight    =   7575
   ClientLeft      =   1095
   ClientTop       =   330
   ClientWidth     =   14610
   ControlBox      =   0   'False
   Icon            =   "frmconsultalibroventas.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7575
   ScaleWidth      =   14610
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "&Actualizar"
      Height          =   855
      Left            =   6360
      Picture         =   "frmconsultalibroventas.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   104
      Top             =   5880
      UseMaskColor    =   -1  'True
      Width           =   1095
   End
   Begin VB.Frame Frame3 
      Caption         =   "ORDENAR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   8040
      TabIndex        =   99
      Top             =   0
      Width           =   6375
      Begin VB.OptionButton Option1 
         Caption         =   "Nº Comp."
         Height          =   255
         Index           =   3
         Left            =   4200
         TabIndex        =   103
         Top             =   360
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Tipo Comp."
         Height          =   255
         Index           =   2
         Left            =   2760
         TabIndex        =   102
         Top             =   360
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Razon Social"
         Height          =   255
         Index           =   1
         Left            =   1320
         TabIndex        =   101
         Top             =   360
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Fecha"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   100
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.CommandButton grabalibroasiento 
      Caption         =   "Command1"
      Enabled         =   0   'False
      Height          =   255
      Left            =   4320
      TabIndex        =   94
      Top             =   6240
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox Text12 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
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
      Left            =   3120
      Locked          =   -1  'True
      TabIndex        =   93
      TabStop         =   0   'False
      Text            =   "Cta.Cte.:"
      Top             =   1320
      Width           =   735
   End
   Begin VB.TextBox Text10 
      DataField       =   "asiento"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Left            =   5280
      TabIndex        =   91
      Text            =   "Text10"
      Top             =   240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Text8 
      DataField       =   "asentado"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Left            =   5280
      TabIndex        =   90
      Text            =   "Text8"
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Index           =   1
      Left            =   6000
      TabIndex        =   89
      Text            =   "Text5"
      Top             =   5400
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Index           =   0
      Left            =   6000
      TabIndex        =   88
      Text            =   "Text5"
      Top             =   5160
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      DataField       =   "cht"
      DataSource      =   "datPrimaryRS"
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
      Index           =   31
      Left            =   6960
      Locked          =   -1  'True
      TabIndex        =   85
      Text            =   " "
      Top             =   5280
      Width           =   615
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      DataField       =   "cdt"
      DataSource      =   "datPrimaryRS"
      Enabled         =   0   'False
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
      Index           =   30
      Left            =   6240
      Locked          =   -1  'True
      TabIndex        =   84
      Text            =   " "
      Top             =   5280
      Width           =   615
   End
   Begin VB.CommandButton calcular 
      Caption         =   "calcular"
      Height          =   255
      Left            =   600
      TabIndex        =   50
      Top             =   6120
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      DataField       =   "col15"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """$"" #,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   11274
         SubFormatType   =   2
      EndProperty
      DataSource      =   "datPrimaryRS"
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
      Index           =   14
      Left            =   3360
      Locked          =   -1  'True
      TabIndex        =   23
      Top             =   6960
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      DataField       =   "col14"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """$"" #,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   11274
         SubFormatType   =   2
      EndProperty
      DataSource      =   "datPrimaryRS"
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
      Index           =   13
      Left            =   3360
      Locked          =   -1  'True
      TabIndex        =   22
      Top             =   6720
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      DataField       =   "col13"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """$"" #,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2058
         SubFormatType   =   2
      EndProperty
      DataSource      =   "datPrimaryRS"
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
      Index           =   12
      Left            =   3360
      Locked          =   -1  'True
      TabIndex        =   21
      Top             =   6480
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      DataField       =   "col12"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """$"" #,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2058
         SubFormatType   =   2
      EndProperty
      DataSource      =   "datPrimaryRS"
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
      Index           =   11
      Left            =   3360
      Locked          =   -1  'True
      TabIndex        =   20
      Top             =   6240
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      DataField       =   "col11"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """$"" #,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2058
         SubFormatType   =   2
      EndProperty
      DataSource      =   "datPrimaryRS"
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
      Index           =   10
      Left            =   3360
      Locked          =   -1  'True
      TabIndex        =   19
      Top             =   6000
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      DataField       =   "col10"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """$"" #,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2058
         SubFormatType   =   2
      EndProperty
      DataSource      =   "datPrimaryRS"
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
      Index           =   9
      Left            =   3360
      Locked          =   -1  'True
      TabIndex        =   18
      Top             =   5760
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      DataField       =   "col9"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """$"" #,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2058
         SubFormatType   =   2
      EndProperty
      DataSource      =   "datPrimaryRS"
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
      Index           =   8
      Left            =   3360
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   5520
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      DataField       =   "col8"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """$"" #,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2058
         SubFormatType   =   2
      EndProperty
      DataSource      =   "datPrimaryRS"
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
      Index           =   7
      Left            =   3360
      Locked          =   -1  'True
      TabIndex        =   16
      Top             =   5280
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      DataField       =   "col7"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """$"" #,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2058
         SubFormatType   =   2
      EndProperty
      DataSource      =   "datPrimaryRS"
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
      Index           =   6
      Left            =   3360
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   5040
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      DataField       =   "col6"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """$"" #,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2058
         SubFormatType   =   2
      EndProperty
      DataSource      =   "datPrimaryRS"
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
      Index           =   5
      Left            =   3360
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   4800
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      DataField       =   "col5"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """$"" #,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2058
         SubFormatType   =   2
      EndProperty
      DataSource      =   "datPrimaryRS"
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
      Index           =   4
      Left            =   3360
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   4560
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      DataField       =   "col4"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """$"" #,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2058
         SubFormatType   =   2
      EndProperty
      DataSource      =   "datPrimaryRS"
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
      Index           =   3
      Left            =   3360
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   4320
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      DataField       =   "col3"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """$"" #,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2058
         SubFormatType   =   2
      EndProperty
      DataSource      =   "datPrimaryRS"
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
      Index           =   2
      Left            =   3360
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   4080
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      DataField       =   "col2"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """$"" #,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2058
         SubFormatType   =   2
      EndProperty
      DataSource      =   "datPrimaryRS"
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
      Index           =   1
      Left            =   3360
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   3840
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      DataField       =   "col1"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """$"" #,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2058
         SubFormatType   =   2
      EndProperty
      DataSource      =   "datPrimaryRS"
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
      Index           =   0
      Left            =   3360
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   3570
      Width           =   1215
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      DataField       =   "ch15"
      DataSource      =   "datPrimaryRS"
      Enabled         =   0   'False
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
      Index           =   29
      Left            =   5280
      Locked          =   -1  'True
      TabIndex        =   81
      Text            =   " "
      Top             =   6960
      Width           =   615
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      DataField       =   "cd15"
      DataSource      =   "datPrimaryRS"
      Enabled         =   0   'False
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
      Index           =   28
      Left            =   4680
      TabIndex        =   80
      Text            =   " "
      Top             =   6960
      Width           =   615
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      DataField       =   "ch14"
      DataSource      =   "datPrimaryRS"
      Enabled         =   0   'False
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
      Index           =   27
      Left            =   5280
      Locked          =   -1  'True
      TabIndex        =   79
      Text            =   " "
      Top             =   6720
      Width           =   615
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      DataField       =   "cd14"
      DataSource      =   "datPrimaryRS"
      Enabled         =   0   'False
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
      Index           =   26
      Left            =   4680
      TabIndex        =   78
      Text            =   " "
      Top             =   6720
      Width           =   615
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      DataField       =   "ch13"
      DataSource      =   "datPrimaryRS"
      Enabled         =   0   'False
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
      Index           =   25
      Left            =   5280
      Locked          =   -1  'True
      TabIndex        =   77
      Text            =   " "
      Top             =   6480
      Width           =   615
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      DataField       =   "cd13"
      DataSource      =   "datPrimaryRS"
      Enabled         =   0   'False
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
      Index           =   24
      Left            =   4680
      TabIndex        =   76
      Text            =   " "
      Top             =   6480
      Width           =   615
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      DataField       =   "ch12"
      DataSource      =   "datPrimaryRS"
      Enabled         =   0   'False
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
      Index           =   23
      Left            =   5280
      Locked          =   -1  'True
      TabIndex        =   75
      Text            =   " "
      Top             =   6240
      Width           =   615
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      DataField       =   "cd12"
      DataSource      =   "datPrimaryRS"
      Enabled         =   0   'False
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
      Index           =   22
      Left            =   4680
      TabIndex        =   74
      Text            =   " "
      Top             =   6240
      Width           =   615
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      DataField       =   "ch11"
      DataSource      =   "datPrimaryRS"
      Enabled         =   0   'False
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
      Index           =   21
      Left            =   5280
      Locked          =   -1  'True
      TabIndex        =   73
      Text            =   " "
      Top             =   6000
      Width           =   615
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      DataField       =   "cd11"
      DataSource      =   "datPrimaryRS"
      Enabled         =   0   'False
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
      Index           =   20
      Left            =   4680
      TabIndex        =   72
      Text            =   " "
      Top             =   6000
      Width           =   615
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      DataField       =   "ch10"
      DataSource      =   "datPrimaryRS"
      Enabled         =   0   'False
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
      Index           =   19
      Left            =   5280
      Locked          =   -1  'True
      TabIndex        =   71
      Text            =   " "
      Top             =   5760
      Width           =   615
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      DataField       =   "cd10"
      DataSource      =   "datPrimaryRS"
      Enabled         =   0   'False
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
      Index           =   18
      Left            =   4680
      TabIndex        =   70
      Text            =   " "
      Top             =   5760
      Width           =   615
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      DataField       =   "ch9"
      DataSource      =   "datPrimaryRS"
      Enabled         =   0   'False
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
      Index           =   17
      Left            =   5280
      Locked          =   -1  'True
      TabIndex        =   69
      Text            =   " "
      Top             =   5520
      Width           =   615
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      DataField       =   "cd9"
      DataSource      =   "datPrimaryRS"
      Enabled         =   0   'False
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
      Index           =   16
      Left            =   4680
      TabIndex        =   68
      Text            =   " "
      Top             =   5520
      Width           =   615
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      DataField       =   "ch8"
      DataSource      =   "datPrimaryRS"
      Enabled         =   0   'False
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
      Index           =   15
      Left            =   5280
      Locked          =   -1  'True
      TabIndex        =   67
      Text            =   " "
      Top             =   5280
      Width           =   615
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      DataField       =   "cd8"
      DataSource      =   "datPrimaryRS"
      Enabled         =   0   'False
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
      Index           =   14
      Left            =   4680
      TabIndex        =   66
      Text            =   " "
      Top             =   5280
      Width           =   615
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      DataField       =   "ch7"
      DataSource      =   "datPrimaryRS"
      Enabled         =   0   'False
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
      Index           =   13
      Left            =   5280
      Locked          =   -1  'True
      TabIndex        =   65
      Text            =   " "
      Top             =   5040
      Width           =   615
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      DataField       =   "cd7"
      DataSource      =   "datPrimaryRS"
      Enabled         =   0   'False
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
      Index           =   12
      Left            =   4680
      TabIndex        =   64
      Text            =   " "
      Top             =   5040
      Width           =   615
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      DataField       =   "ch6"
      DataSource      =   "datPrimaryRS"
      Enabled         =   0   'False
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
      Index           =   11
      Left            =   5280
      Locked          =   -1  'True
      TabIndex        =   63
      Text            =   " "
      Top             =   4800
      Width           =   615
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      DataField       =   "cd6"
      DataSource      =   "datPrimaryRS"
      Enabled         =   0   'False
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
      Index           =   10
      Left            =   4680
      TabIndex        =   62
      Text            =   " "
      Top             =   4800
      Width           =   615
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      DataField       =   "ch5"
      DataSource      =   "datPrimaryRS"
      Enabled         =   0   'False
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
      Index           =   9
      Left            =   5280
      Locked          =   -1  'True
      TabIndex        =   61
      Text            =   " "
      Top             =   4560
      Width           =   615
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      DataField       =   "cd5"
      DataSource      =   "datPrimaryRS"
      Enabled         =   0   'False
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
      Index           =   8
      Left            =   4680
      TabIndex        =   60
      Text            =   " "
      Top             =   4560
      Width           =   615
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      DataField       =   "ch4"
      DataSource      =   "datPrimaryRS"
      Enabled         =   0   'False
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
      Index           =   7
      Left            =   5280
      Locked          =   -1  'True
      TabIndex        =   59
      Text            =   " "
      Top             =   4320
      Width           =   615
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      DataField       =   "cd4"
      DataSource      =   "datPrimaryRS"
      Enabled         =   0   'False
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
      Index           =   6
      Left            =   4680
      TabIndex        =   58
      Text            =   " "
      Top             =   4320
      Width           =   615
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      DataField       =   "ch3"
      DataSource      =   "datPrimaryRS"
      Enabled         =   0   'False
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
      Index           =   5
      Left            =   5280
      Locked          =   -1  'True
      TabIndex        =   57
      Text            =   " "
      Top             =   4080
      Width           =   615
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      DataField       =   "cd3"
      DataSource      =   "datPrimaryRS"
      Enabled         =   0   'False
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
      Index           =   4
      Left            =   4680
      TabIndex        =   56
      Text            =   " "
      Top             =   4080
      Width           =   615
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      DataField       =   "ch2"
      DataSource      =   "datPrimaryRS"
      Enabled         =   0   'False
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
      Index           =   3
      Left            =   5280
      Locked          =   -1  'True
      TabIndex        =   55
      Text            =   " "
      Top             =   3840
      Width           =   615
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      DataField       =   "cd2"
      DataSource      =   "datPrimaryRS"
      Enabled         =   0   'False
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
      Index           =   2
      Left            =   4680
      TabIndex        =   54
      Text            =   " "
      Top             =   3840
      Width           =   615
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      DataField       =   "ch1"
      DataSource      =   "datPrimaryRS"
      Enabled         =   0   'False
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
      Index           =   1
      Left            =   5280
      Locked          =   -1  'True
      TabIndex        =   53
      Text            =   " "
      Top             =   3570
      Width           =   615
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      DataField       =   "cd1"
      DataSource      =   "datPrimaryRS"
      Enabled         =   0   'False
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
      Index           =   0
      Left            =   4680
      TabIndex        =   52
      Top             =   3585
      Width           =   615
   End
   Begin VB.TextBox Text1 
      DataField       =   "empresa"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   1
      Left            =   720
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   0
      Visible         =   0   'False
      Width           =   735
   End
   Begin MSDataGridLib.DataGrid DataGrid2 
      Bindings        =   "frmconsultalibroventas.frx":0884
      Height          =   1335
      Left            =   8040
      TabIndex        =   44
      Top             =   2280
      Visible         =   0   'False
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   2355
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
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
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton salir 
      Caption         =   "&Cerrar"
      Height          =   735
      Left            =   6360
      Picture         =   "frmconsultalibroventas.frx":089E
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   3480
      UseMaskColor    =   -1  'True
      Width           =   1095
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      DataField       =   "total"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """$"" #,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2058
         SubFormatType   =   2
      EndProperty
      DataSource      =   "datPrimaryRS"
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
      Index           =   15
      Left            =   6000
      Locked          =   -1  'True
      TabIndex        =   47
      TabStop         =   0   'False
      Top             =   4680
      Width           =   1935
   End
   Begin VB.CommandButton borrar 
      Caption         =   "&Filtrar"
      Height          =   735
      Left            =   6360
      Picture         =   "frmconsultalibroventas.frx":0CE0
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   1800
      UseMaskColor    =   -1  'True
      Width           =   1095
   End
   Begin VB.TextBox tipocomp 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      DataField       =   "tipocompr"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Left            =   480
      Locked          =   -1  'True
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   1320
      Width           =   615
   End
   Begin VB.TextBox cuit 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      DataField       =   "cuit"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   960
      Width           =   1575
   End
   Begin VB.TextBox tipoiva 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      DataField       =   "tipoiva"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Left            =   480
      Locked          =   -1  'True
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   960
      Width           =   495
   End
   Begin VB.TextBox denominacion 
      BackColor       =   &H00E0E0E0&
      DataField       =   "cliente"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Left            =   480
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   600
      Width           =   3735
   End
   Begin MSMask.MaskEdBox Maskfecha 
      DataField       =   "fecha"
      DataSource      =   "datPrimaryRS"
      Height          =   255
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   450
      _Version        =   393216
      AllowPrompt     =   -1  'True
      Enabled         =   0   'False
      Format          =   "dd/mm/yyyy"
      PromptChar      =   "_"
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmconsultalibroventas.frx":1122
      Height          =   6495
      Left            =   8040
      Negotiate       =   -1  'True
      TabIndex        =   4
      Top             =   960
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   11456
      _Version        =   393216
      AllowUpdate     =   -1  'True
      AllowArrows     =   -1  'True
      BackColor       =   14737632
      HeadLines       =   4
      RowHeight       =   15
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   24
      BeginProperty Column00 
         DataField       =   "empresa"
         Caption         =   "empresa"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   """$"" #,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
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
      BeginProperty Column02 
         DataField       =   "cliente"
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
      BeginProperty Column03 
         DataField       =   "tipoiva"
         Caption         =   "Tipo I.V.A."
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
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
         DataField       =   "cuit"
         Caption         =   "C.U.I.T."
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
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
         DataField       =   "tipocompr"
         Caption         =   "Tipo Comp."
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
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
         DataField       =   "numcompr"
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
      BeginProperty Column07 
         DataField       =   "col1"
         Caption         =   "col1"
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
      BeginProperty Column08 
         DataField       =   "col2"
         Caption         =   "col2"
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
         DataField       =   "col3"
         Caption         =   "col3"
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
         DataField       =   "col4"
         Caption         =   "col4"
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
      BeginProperty Column11 
         DataField       =   "col5"
         Caption         =   "col5"
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
      BeginProperty Column12 
         DataField       =   "col6"
         Caption         =   "col6"
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
      BeginProperty Column13 
         DataField       =   "col7"
         Caption         =   "col7"
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
      BeginProperty Column14 
         DataField       =   "col8"
         Caption         =   "col8"
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
      BeginProperty Column15 
         DataField       =   "col9"
         Caption         =   "col9"
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
      BeginProperty Column16 
         DataField       =   "col10"
         Caption         =   "col10"
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
         DataField       =   "col11"
         Caption         =   "col11"
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
         DataField       =   "col12"
         Caption         =   "col12"
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
      BeginProperty Column19 
         DataField       =   "col13"
         Caption         =   "col13"
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
      BeginProperty Column20 
         DataField       =   "col14"
         Caption         =   "col14"
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
      BeginProperty Column21 
         DataField       =   "col15"
         Caption         =   "col15"
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
      BeginProperty Column22 
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
      BeginProperty Column23 
         DataField       =   "total"
         Caption         =   "total"
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
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   2
         SizeMode        =   1
         BeginProperty Column00 
            Locked          =   -1  'True
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column01 
            Locked          =   -1  'True
         EndProperty
         BeginProperty Column02 
            Locked          =   -1  'True
         EndProperty
         BeginProperty Column03 
            Locked          =   -1  'True
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column04 
            Locked          =   -1  'True
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column05 
            Locked          =   -1  'True
         EndProperty
         BeginProperty Column06 
            Locked          =   -1  'True
         EndProperty
         BeginProperty Column07 
            Alignment       =   1
            Locked          =   -1  'True
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column08 
            Alignment       =   1
            Locked          =   -1  'True
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column09 
            Alignment       =   1
            Locked          =   -1  'True
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column10 
            Alignment       =   1
            Locked          =   -1  'True
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column11 
            Alignment       =   1
            Locked          =   -1  'True
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column12 
            Alignment       =   1
            Locked          =   -1  'True
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column13 
            Alignment       =   1
            Locked          =   -1  'True
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column14 
            Alignment       =   1
            Locked          =   -1  'True
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column15 
            Alignment       =   1
            Locked          =   -1  'True
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column16 
            Alignment       =   1
            Locked          =   -1  'True
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column17 
            Alignment       =   1
            Locked          =   -1  'True
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column18 
            Alignment       =   1
            Locked          =   -1  'True
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column19 
            Alignment       =   1
            Locked          =   -1  'True
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column20 
            Alignment       =   1
            Locked          =   -1  'True
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column21 
            Alignment       =   1
            Locked          =   -1  'True
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column22 
            Alignment       =   2
            Locked          =   -1  'True
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column23 
            Alignment       =   1
            Locked          =   -1  'True
            Object.Visible         =   0   'False
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc datPrimaryRS 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      Top             =   7245
      Visible         =   0   'False
      Width           =   14610
      _ExtentX        =   25770
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
   Begin MSMask.MaskEdBox Maskcomprobante 
      DataField       =   "numcompr"
      DataSource      =   "datPrimaryRS"
      Height          =   255
      Left            =   480
      TabIndex        =   6
      Top             =   1680
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   450
      _Version        =   393216
      AllowPrompt     =   -1  'True
      Enabled         =   0   'False
      PromptChar      =   "_"
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
      ForeColor       =   &H00C00000&
      Height          =   4335
      Left            =   6000
      TabIndex        =   45
      Top             =   0
      Width           =   1935
      Begin VB.CommandButton Command2 
         Caption         =   "&Imprimir"
         Height          =   735
         Left            =   360
         Picture         =   "frmconsultalibroventas.frx":113D
         Style           =   1  'Graphical
         TabIndex        =   105
         Top             =   2640
         UseMaskColor    =   -1  'True
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker desde 
         Height          =   375
         Left            =   240
         TabIndex        =   97
         Top             =   600
         Width           =   1455
         _ExtentX        =   2566
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
         CalendarBackColor=   16777215
         Format          =   21364737
         CurrentDate     =   38410
      End
      Begin MSComCtl2.DTPicker hasta 
         Height          =   375
         Left            =   240
         TabIndex        =   98
         Top             =   1320
         Width           =   1455
         _ExtentX        =   2566
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
         CalendarBackColor=   16777215
         Format          =   21364737
         CurrentDate     =   38410
      End
      Begin VB.TextBox Text1 
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
         Index           =   2
         Left            =   240
         TabIndex        =   96
         Text            =   "Hasta"
         Top             =   1080
         Width           =   495
      End
      Begin VB.TextBox Text1 
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
         Index           =   0
         Left            =   240
         TabIndex        =   95
         Text            =   "Desde"
         Top             =   360
         Width           =   615
      End
      Begin MSAdodcLib.Adodc datcuentas 
         Height          =   330
         Left            =   240
         Top             =   2640
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
   Begin MSAdodcLib.Adodc datcolumnas 
      Height          =   330
      Left            =   0
      Top             =   480
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
   Begin VB.TextBox Text2 
      DataField       =   "fecha"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Left            =   120
      TabIndex        =   43
      Text            =   "Text2"
      Top             =   960
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Text4 
      DataField       =   "numcompr"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Left            =   120
      TabIndex        =   46
      Text            =   "Text4"
      Top             =   1440
      Visible         =   0   'False
      Width           =   975
   End
   Begin MSDataGridLib.DataGrid DataGrid3 
      Bindings        =   "frmconsultalibroventas.frx":289F
      Height          =   735
      Left            =   8160
      TabIndex        =   49
      Top             =   3840
      Visible         =   0   'False
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   1296
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
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
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc datclientes 
      Height          =   330
      Left            =   6120
      Top             =   1920
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
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Height          =   3255
      Left            =   120
      TabIndex        =   42
      Top             =   0
      Width           =   5775
      Begin VB.CommandButton Command3 
         Caption         =   "Pre&visualizar Impresion de Comprobante"
         Height          =   855
         Left            =   360
         Picture         =   "frmconsultalibroventas.frx":28B8
         Style           =   1  'Graphical
         TabIndex        =   106
         Top             =   2160
         UseMaskColor    =   -1  'True
         Width           =   4695
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Enabled         =   0   'False
         Height          =   255
         Index           =   1
         Left            =   4680
         TabIndex        =   8
         Top             =   1320
         Width           =   255
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Enabled         =   0   'False
         Height          =   255
         Index           =   0
         Left            =   3720
         TabIndex        =   7
         Top             =   1320
         Width           =   255
      End
      Begin VB.TextBox Text11 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
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
         Left            =   4080
         Locked          =   -1  'True
         TabIndex        =   92
         TabStop         =   0   'False
         Text            =   "Cont.:"
         Top             =   1320
         Width           =   615
      End
      Begin VB.TextBox Text6 
         DataField       =   "cerrado"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Left            =   1440
         TabIndex        =   51
         Text            =   "Text6"
         Top             =   0
         Visible         =   0   'False
         Width           =   855
      End
   End
   Begin MSAdodcLib.Adodc datmaestro 
      Height          =   330
      Left            =   6360
      Top             =   6120
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
      Left            =   6360
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
   Begin MSAdodcLib.Adodc datperiodo 
      Height          =   330
      Left            =   6360
      Top             =   6840
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
   Begin MSAdodcLib.Adodc datlistacostos 
      Height          =   330
      Left            =   6600
      Top             =   5640
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
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
   Begin MSAdodcLib.Adodc datccostos 
      Height          =   330
      Left            =   6600
      Top             =   5880
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
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
   Begin MSAdodcLib.Adodc datbusca 
      Height          =   330
      Left            =   7560
      Top             =   6000
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
   Begin vbskpro.Skinner Skinner1 
      Left            =   0
      Top             =   0
      _ExtentX        =   1270
      _ExtentY        =   1270
      CloseButton     =   0
      MaxButton       =   0
      MinButton       =   0
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
      LcK2            =   $"frmconsultalibroventas.frx":2CFA
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
      Left            =   6000
      Top             =   6600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTitle     =   "Libro IVA Compras"
      PrintFileLinesPerPage=   60
   End
   Begin MSRDC.MSRDC reporte 
      Height          =   375
      Left            =   1800
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
   Begin MSRDC.MSRDC MSRDC1 
      Height          =   375
      Left            =   6840
      Top             =   6840
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
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "C.Deb"
      Enabled         =   0   'False
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
      Index           =   19
      Left            =   6240
      TabIndex        =   87
      Top             =   5040
      Width           =   615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "C.Hab"
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
      Index           =   16
      Left            =   6960
      TabIndex        =   86
      Top             =   5040
      Width           =   615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "C.Hab"
      Enabled         =   0   'False
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
      Index           =   18
      Left            =   5280
      TabIndex        =   83
      Top             =   3360
      Width           =   615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "C.Deb"
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
      Index           =   17
      Left            =   4680
      TabIndex        =   82
      Top             =   3360
      Width           =   615
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      DataField       =   "nomcol15"
      DataSource      =   "datcolumnas"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   14
      Left            =   120
      TabIndex        =   41
      Top             =   6960
      Width           =   3015
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      DataField       =   "nomcol14"
      DataSource      =   "datcolumnas"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   13
      Left            =   120
      TabIndex        =   40
      Top             =   6720
      Width           =   3015
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      DataField       =   "nomcol13"
      DataSource      =   "datcolumnas"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   12
      Left            =   120
      TabIndex        =   39
      Top             =   6480
      Width           =   3015
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      DataField       =   "nomcol12"
      DataSource      =   "datcolumnas"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   11
      Left            =   120
      TabIndex        =   38
      Top             =   6240
      Width           =   3015
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      DataField       =   "nomcol11"
      DataSource      =   "datcolumnas"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   10
      Left            =   120
      TabIndex        =   37
      Top             =   6000
      Width           =   3015
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      DataField       =   "nomcol10"
      DataSource      =   "datcolumnas"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   9
      Left            =   120
      TabIndex        =   36
      Top             =   5760
      Width           =   3015
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      DataField       =   "nomcol9"
      DataSource      =   "datcolumnas"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   8
      Left            =   120
      TabIndex        =   35
      Top             =   5520
      Width           =   3015
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      DataField       =   "nomcol8"
      DataSource      =   "datcolumnas"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   34
      Top             =   5280
      Width           =   3015
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      DataField       =   "nomcol7"
      DataSource      =   "datcolumnas"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   33
      Top             =   5040
      Width           =   3015
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      DataField       =   "nomcol6"
      DataSource      =   "datcolumnas"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   32
      Top             =   4800
      Width           =   3015
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      DataField       =   "nomcol5"
      DataSource      =   "datcolumnas"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   31
      Top             =   4560
      Width           =   3015
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      DataField       =   "nomcol4"
      DataSource      =   "datcolumnas"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   30
      Top             =   4320
      Width           =   3015
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      DataField       =   "nomcol3"
      DataSource      =   "datcolumnas"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   29
      Top             =   4080
      Width           =   3015
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      DataField       =   "nomcol2"
      DataSource      =   "datcolumnas"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   28
      Top             =   3840
      Width           =   3015
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      DataField       =   "nomcol1"
      DataSource      =   "datcolumnas"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   27
      Top             =   3600
      Width           =   3015
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "TOTAL"
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
      Index           =   15
      Left            =   6000
      TabIndex        =   48
      Top             =   4440
      Width           =   1935
   End
End
Attribute VB_Name = "frmconsutalibroventas"
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
Dim ter(15, 15), sig(15, 15) As String



Private Sub Adodc1_WillMove(ByVal adReason As ADODB.EventReasonEnum, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)

End Sub

Private Sub borrar_Click()

  datPrimaryRS.RecordSource = "select libroventas.* from libroventas WHERE libroventas.empresa = " & login.empresaact & " and inicioper = '" & login.iper & "' and finper = '" & login.fper & "' and fecha >= '" & desde.Value & "' and fecha <= '" & hasta.Value & "' Order by fecha"
  datPrimaryRS.Refresh


End Sub


Private Sub calcular_Click()
On Error GoTo erroRcalcular

sumar = 0
parcial = 0
result = 0
For t = 0 To 14
        For x = 1 To 15
            If sig(t, x) = "" And x = 1 Then GoTo paso1
            If sumar = 1 Then parcial = result
            If sig(t, x) = "*" Then result = Text3(ter(t, x) - 1).Text * Val(ter(t, x + 1))
            If sig(t, x) = "/" Then result = Text3(ter(t, x) - 1).Text / Val(ter(t, x + 1))
            If sig(t, x) = "+" Or sig(t, x) = "" Then
                    result = parcial + result
                    sumar = 1
                    If sig(t, x) = "" Then
                            sumar = 0
                            parcial = 0
                    End If
                    GoTo paso0
            Else
                    sumar = 0
            End If
            If sig(t, x) = "-" Or sig(t, x) = "" Then
                    result = parcial - result
                    sumar = 1
                    If sig(t, x) = "" Then
                            sumar = 0
                            parcial = 0
                    End If
                    GoTo paso0
            Else
                    sumar = 0
            End If
paso0:
        Text3(t).Text = result
         Next x
       
paso1:
Next t
Exit Sub

erroRcalcular:
    mensa = MsgBox("Error al realizar calculo automatico, revisar configuracion de columnas del libro", vbCritical, "!! Atención !!")


End Sub

Private Sub Cancelar_Click()

        datPrimaryRS.Refresh
        If datPrimaryRS.Recordset.EOF = True Then Exit Sub
        datPrimaryRS.Recordset.MoveLast

End Sub

Private Sub confcolumnas_Click()
    
    Unload Me
    frmcolumnasventa.Show


End Sub

Private Sub Check1_KeyPress(Index As Integer, KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        KeyAscii = 9
    End If
    
End Sub

Private Sub Command1_Click()

  datPrimaryRS.RecordSource = "select libroventas.* from libroventas WHERE libroventas.empresa = " & login.empresaact & " and inicioper = '" & login.iper & "' and finper = '" & login.fper & "' Order by fecha"
  datPrimaryRS.Refresh

    

End Sub

Private Sub DataList1_KeyPress(KeyAscii As Integer)

    If KeyAscii = 45 Then
        Maskfecha.SetFocus
    End If

    If KeyAscii = 13 Then
        If DataList1.SelectedItem <> "" Then
            datclientes.Recordset.Bookmark = DataList1.SelectedItem
            Check1(0) = 1
        Rem    If datproveedores.Recordset.Fields("ccocontado") = "S" Then
        Rem        Check1(0).Value = 0
        Rem        Check1(1).Value = 1
        Rem    Else
        Rem        Check1(0).Value = 1
        Rem        Check1(1).Value = 0
        Rem    End If
            denominacion = datclientes.Recordset.Fields(2)
            If datclientes.Recordset.Fields(3) <> "" Then tipoiva.Text = datclientes.Recordset.Fields(3)
            If datclientes.Recordset.Fields(4) <> "" Then cuit.Text = datclientes.Recordset.Fields(4)
            pruebanulo = IsNull(datclientes.Recordset.Fields(11))
            If pruebanulo = True Then
                Text7(31).Text = 0
            Else
                Text7(31).Text = datclientes.Recordset.Fields(11)
            End If
            List1.Visible = True
            List1.SetFocus
        Else
            Exit Sub
        End If
    End If

End Sub

Private Sub Command2_Click()
Dim tabla As String
Dim ruta As String

Dim crxapplication As New CRAXDRT.Application
Dim crxreport As CRAXDRT.Report

ruta = "\Empresa" + Right(Str(login.empresaact), Len(Str(login.empresaact)) - 1)

reporte.SQL = "SELECT libroventas1.id, libroventas1.empresa,libroventas1.cerrado, libroventas1.fecha, libroventas1.cliente, libroventas1.tipoiva, libroventas1.cuit, libroventas1.tipocompr, libroventas1.numcompr, libroventas1.col1, libroventas1.col2, libroventas1.col3, libroventas1.total, libroventas1.inicioper, libroventas1.finper, libroventas1.nomcol1, libroventas1.nomcol2, libroventas1.nomcol3, libroventas1.razonsocial FROM contablesql.dbo.libroventas1 libroventas1 WHERE libroventas1.empresa = " & login.empresaact & " and libroventas1.inicioper = '" & login.iper & "' and fecha >= '" & desde.Value & "' and fecha <= '" & hasta.Value & "' ORDER BY libroventas1.fecha ASC, libroventas1.id ASC"
tabla = reporte.SQL



With crystalreporte
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

    Inicio.datauditoria.ConnectionString = login.conexiontotal
    
    Inicio.datauditoria.RecordSource = "select auditoria.* from auditoria"
    Inicio.datauditoria.Refresh
    
    Inicio.datauditoria.Recordset.AddNew
    Inicio.datauditoria.Recordset.Fields("fecha") = Date
    Inicio.datauditoria.Recordset.Fields("hora") = Str(Time)
    Inicio.datauditoria.Recordset.Fields("ventana") = "Consulta Facturas de Ventas"
    Inicio.datauditoria.Recordset.Fields("accion") = "Periodo de Consulta " + Str(desde.Value) + " - " + Str(hasta.Value)
    Inicio.datauditoria.Recordset.Fields("usuario") = login.usuarioactivo
    Inicio.datauditoria.Recordset.Fields("empresa") = login.empresaact
    Inicio.datauditoria.Recordset.UpdateBatch adAffectCurrent
End Sub

Private Sub Command3_Click()


    frmfacturaconsulta.Show
    frmfacturaconsulta.Combo1 = tipocomp.Text
    frmfacturaconsulta.DataCombo1 = Maskcomprobante.Text
    frmfacturaconsulta.Command4.SetFocus
    SendKeys "{ENTER}", True
    

End Sub

Private Sub DataGrid1_Click()

    If datPrimaryRS.Recordset.Fields("contado") = "S" Then
        Check1(1).Value = 1
        Check1(0).Value = 0
    Else
        Check1(1).Value = 0
        Check1(0).Value = 1
    End If

End Sub

Private Sub DataGrid1_DblClick()

  If datPrimaryRS.Recordset.Fields("cerrado") = "N" Then
      frmlibroventas.datPrimaryRS.RecordSource = "select libroventas.* from libroventas WHERE libroventas.empresa = " & login.empresaact & " and id = " & datPrimaryRS.Recordset.Fields("id") & ""
      frmlibroventas.datPrimaryRS.Refresh
      frmlibroventas.Show
  Else
      If login.librocerrado = "S" Then
            mensa = MsgBox("Este registro pertenece a un libro Cerrado", vbCritical, "Atencion")
            frmlibroventas.datPrimaryRS.RecordSource = "select libroventas.* from libroventas WHERE libroventas.empresa = " & login.empresaact & " and id = " & datPrimaryRS.Recordset.Fields("id") & ""
            frmlibroventas.datPrimaryRS.Refresh
            frmlibroventas.Show
            Exit Sub
      End If
      mensa = MsgBox("Este registro pertenece a un libro Cerrado, no puede modificarlo", vbCritical, "Atencion")
  End If
  

End Sub

Private Sub DataGrid1_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 9
    End If

End Sub

Private Sub DataList1_LostFocus()
  DataList1.Visible = False
End Sub
Private Sub DataList2_Click()

    Text7(poscuenta).Text = DataList2.BoundText
    
    
End Sub

Private Sub DataList2_GotFocus()

    If Inicio.opcion1 = True Then
        datcuentas.RecordSource = "select listacuentas.* from listacuentas WHERE empre = " & login.empresaact & " and inicioper = '" & login.iper & "' and finper = '" & login.fper & "' ORDER BY IDCUENTA"
        datcuentas.Refresh
        DataList2.ListField = "codigo"
    End If
    If Inicio.opcion2 = True Then
        datcuentas.RecordSource = "select listacuentas.* from listacuentas WHERE empre = " & login.empresaact & " and inicioper = '" & login.iper & "' and finper = '" & login.fper & "' ORDER BY nombre"
        datcuentas.Refresh
        DataList2.ListField = "nombre"
    End If
    DataList2.BoundText = Text7(poscuenta).Text

    
    
End Sub

Private Sub DataList2_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        Text7(poscuenta).Text = DataList2.BoundText

            If Text7(poscuenta).Text = "" Then Text7(poscuenta).Text = 0
             
              If Text7(poscuenta).Text = 0 Then
                    mensa = MsgBox("Debe ingresar un Nº de cuenta", vbCritical, "!! Error !!")
                    Text7(poscuenta).SetFocus
                    errorasiento = True
                    Exit Sub
              End If
              errorasiento = False
               
              If datccostos.Recordset.EOF = True Then GoTo sigue
              datccostos.Recordset.MoveFirst
              digito = Val(datccostos.Recordset.Fields(3))
              digcue = Val(Mid(Text7(poscuenta).Text, 1, 1))
              If digcue = digito And login.habcc = True Then
                DataList3.Visible = True
                Text9.Visible = True
                Frame3.Visible = True
                DataList3.SetFocus
                Exit Sub
              End If
sigue:
              If poscuenta < 31 Then sumadebe = Text3(posicion) + sumadebe
              If Text3(posicion + 1).Visible = True Then
                    Text3(posicion + 1).SetFocus
              Else
                    datPrimaryRS.Recordset.UpdateBatch adAffectCurrent
                    pos = datPrimaryRS.Recordset.AbsolutePosition
                    datPrimaryRS.Refresh
                    datPrimaryRS.Recordset.AbsolutePosition = pos
                    Text7(31).SetFocus
              End If
              If Index = 31 Then
                    sumahaber = Text3(15) + sumahaber
                    datPrimaryRS.Recordset.UpdateBatch adAffectCurrent
                    pos = datPrimaryRS.Recordset.AbsolutePosition
                    datPrimaryRS.Refresh
                    datPrimaryRS.Recordset.AbsolutePosition = pos
                    Call grabalibroasiento_Click
              End If

              Exit Sub
    End If

End Sub

Private Sub DataList2_LostFocus()

    DataList2.Visible = False
    
End Sub

Private Sub DataList3_Click()
    Text9.Text = DataList3.BoundText
End Sub

Private Sub DataList3_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
            Text9.Text = DataList3.BoundText
              If poscuenta < 31 Then sumadebe = Text3(posicion) + sumadebe
              If Text3(posicion + 1).Visible = True Then
                    Text3(posicion + 1).SetFocus
              Else
                    datPrimaryRS.Recordset.UpdateBatch adAffectCurrent
                    pos = datPrimaryRS.Recordset.AbsolutePosition
                    datPrimaryRS.Refresh
                    datPrimaryRS.Recordset.AbsolutePosition = pos
                    Text7(31).SetFocus
              End If
              If Index = 31 Then
                    sumahaber = Text3(15) + sumahaber
                    datPrimaryRS.Recordset.UpdateBatch adAffectCurrent
                    pos = datPrimaryRS.Recordset.AbsolutePosition
                    datPrimaryRS.Refresh
                    datPrimaryRS.Recordset.AbsolutePosition = pos
                    Call grabalibroasiento_Click
              End If
    End If
    

End Sub

Private Sub DataList3_LostFocus()

    If Text9.Text = "" Then
        mensa = MsgBox("Debe ingresa un Centro de Costo", vbCritical, "!Error¡")
        DataList3.SetFocus
        Exit Sub
    End If
Frame3.Visible = False
Text9.Visible = False
DataList3.Visible = False

End Sub

Private Sub DataGrid1_KeyUp(KeyCode As Integer, Shift As Integer)

    If datPrimaryRS.Recordset.Fields("contado") = "S" Then
        Check1(1).Value = 1
        Check1(0).Value = 0
    Else
        Check1(1).Value = 0
        Check1(0).Value = 1
    End If

End Sub

Private Sub Form_GotFocus()
    Maskcomprobante.Mask = ""
    Maskfecha.Mask = ""
    

    
End Sub

Private Sub Form_Load()
On Error GoTo errorform

datasiento.ConnectionString = login.conexiontotal
datbusca.ConnectionString = login.conexiontotal
datccostos.ConnectionString = login.conexiontotal
datcolumnas.ConnectionString = login.conexiontotal
datcuentas.ConnectionString = login.conexiontotal
datlistacostos.ConnectionString = login.conexiontotal
datmaestro.ConnectionString = login.conexiontotal
datperiodo.ConnectionString = login.conexiontotal
datPrimaryRS.ConnectionString = login.conexiontotal
datclientes.ConnectionString = login.conexiontotal

Option1(0).Value = True

Text5(0).Text = login.iper
Text5(1).Text = login.fper

desde.Value = login.iper
hasta.Value = login.fper

  frmconsutalibroventas.Left = 0
  frmconsutalibroventas.Top = 0
  
  Inicio.Toolbar1.Visible = True
  datPrimaryRS.RecordSource = "select libroventas.* from libroventas WHERE libroventas.empresa = " & login.empresaact & " and inicioper = '" & login.iper & "' and finper = '" & login.fper & "' Order by fecha"
  datPrimaryRS.Refresh
  
  datcolumnas.RecordSource = "SELECT columnasventa.* FROM columnasventa WHERE empresa = " & login.empresaact & " and inicioper = '" & login.iper & "'"
  datcolumnas.Refresh
  
  datlistacostos.RecordSource = "SELECT listaccostos.* FROM listaccostos WHERE empresa = " & login.empresaact & ""
  datlistacostos.Refresh
  
  datccostos.RecordSource = "SELECT ccostos.* FROM ccostos WHERE empresa = " & login.empresaact & ""
  datccostos.Refresh
  
  datcuentas.RecordSource = "select listacuentas.* from listacuentas WHERE empre = " & login.empresaact & " and inicioper = '" & login.iper & "' and finper = '" & login.fper & "' ORDER BY IDCUENTA"
  datcuentas.Refresh
    
  datclientes.RecordSource = "select codcliente,empresa,razonsocial,tipoiva,cuit,domicilio,localidad,codpostal,telefono,email,contacto from clientes Where empresa = " & login.empresaact & " ORDER BY razonsocial"
  datclientes.Refresh
  
  datmaestro.RecordSource = "select [Maestro Asientos].* from [Maestro Asientos] where empresa = " & login.empresaact & " and perinicial = '" & login.iper & "' order by nroasiento"
  datmaestro.Refresh
  datperiodo.RecordSource = "select EMPRESA.* from EMPRESA where empresa = " & login.empresaact & ""
  datperiodo.Refresh
  datasiento.RecordSource = "select [Detalle Asientos].* from [Detalle Asientos] where empresa = " & login.empresaact & ""
  datasiento.Refresh
  
datcolumnas.Refresh
For x = 1 To 15
verif = IsNull(datcolumnas.Recordset.Fields(x * 2))
If verif = False Then columna(x) = datcolumnas.Recordset.Fields(x * 2)
Next x

For x = 0 To 14
    For Y = 1 To 15
        ter(x, Y) = ""
    Next Y
Next x

t = 1
For c = 0 To 14
        s = 0
        t = 1
        For x = 1 To Len(columna(c + 1))
        car = Mid(columna(c + 1), x, 1)
        If car = "-" Or car = "+" Or car = "*" Or car = "/" Then
            s = s + 1
            sig(c, s) = car
            t = t + 1
            GoTo paso1
        End If
        If car = "C" Or car = "c" Then GoTo paso1
        ter(c, t) = ter(c, t) + car
paso1:
        Next x
Next c

    If datPrimaryRS.Recordset.EOF = True Then
        datPrimaryRS.Recordset.AddNew
            Maskfecha.SelLength = 10
            Maskfecha.SelText = ""
            Maskcomprobante.SelLength = 13
            Maskcomprobante.SelText = ""
        For x = 0 To 14
            Text3(x).Text = 0
        Next x
     Else
       datPrimaryRS.Recordset.MoveLast
     End If
     Maskfecha.Mask = "##/##/####"
     Maskfecha.MaxLength = 10
     Maskcomprobante.Mask = "####-########"
     Maskcomprobante.MaxLength = 13
       
            
    For x = 0 To 14
        valida = IsNull(datcolumnas.Recordset.Fields(x * 2 + 1))
        If datcolumnas.Recordset.Fields(x * 2 + 1) = "" Then valida = True
        If valida = False Then
            Label1(x).Caption = datcolumnas.Recordset.Fields(x * 2 + 1)
            DataGrid1.Columns(x + 7).Caption = datcolumnas.Recordset.Fields(x * 2 + 1)
        Else
            Label1(x).Visible = False
            DataGrid1.Columns(x + 7).Visible = False
            Text3(x).Text = 0
            Text3(x).Visible = False
            Text7(x * 2).Text = 0
            Text7(x * 2 + 1).Text = 0
            Text7(x * 2).Visible = False
            Text7(x * 2 + 1).Visible = False
        End If
    Next x
    
If login.livacomprasmodi = "N" Then
    For x = 0 To 15
        Text3(x).Enabled = False
        Text7(x).Enabled = False
        Text7(x + 16).Enabled = False
    Next x
    denominacion.Enabled = False
    cuit.Enabled = False
    tipoiva.Enabled = False
    tipocomp.Enabled = False
    Maskcomprobante.Enabled = False
    Maskfecha.Enabled = False
    DataList1.Enabled = False
    List1.Enabled = False
    nuevo.Enabled = False
End If

Exit Sub

errorform:
    mensa = MsgBox("Error de Codificacion", vbCritical, "!! Error !!")
     
End Sub

Private Sub Form_Resize()
  On Error Resume Next
  'Esto cambiará el tamaño de la cuadrícula al cambiar el tamaño del formulario
  grdDataGrid.Height = Me.ScaleHeight - datPrimaryRS.Height - 30
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault
      Inicio.Toolbar1.Visible = False
  
End Sub

Private Sub datPrimaryRS_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)
  'Aquí es donde puede colocar el código de control de errores
  'Si desea pasar por alto los errores, marque como comentario la siguiente línea
  'Si desea detectarlos, agregue código aquí para controlarlos
  MsgBox "Data error event hit err:" & Description
End Sub



Private Sub grabalibroasiento_Click()
    If sumadebe <> sumahaber Or errorasiento = True Then
        mensa = MsgBox("El asiento está desvalanceado, no se puede grabar", vbCritical, "!! Error !!")
            sumadebe = 0
            sumahaber = 0
            errorasiento = True
            Text3(0).SetFocus
            Exit Sub
    End If
  errorasiento = False
  
Rem ****************** grabar asiento

    campoaño = Right(Maskfecha.Text, 4)
    campomes = Mid(Maskfecha.Text, 4, 2)
    campodia = Left(Maskfecha.Text, 2)
    campofecha = campoaño + "/" + campomes + "/" + campodia
    
    campoaño1 = Right(Text5(0).Text, 4)
    campomes1 = Mid(Text5(0).Text, 4, 2)
    campodia1 = Left(Text5(0).Text, 2)
    campofecha1 = campoaño1 + "/" + campomes1 + "/" + campodia1
    
    campoaño2 = Right(Text5(1).Text, 4)
    campomes2 = Mid(Text5(1).Text, 4, 2)
    campodia2 = Left(Text5(1).Text, 2)
    campofecha2 = campoaño2 + "/" + campomes2 + "/" + campodia2
    campofecha3 = Right(fechafuera, 4) + "/" + Mid(fechafuera, 4, 3) + Left(fechafuera, 2)
    fechamal = 0
    If campofecha < campofecha1 Or campofecha > campofecha2 Then
            mensa = MsgBox("La Fecha no pertenecia al periodo en ejercicio", vbCritical, "!! Atención !!")
        fechamal = 1
 Rem       Exit Sub
        GoTo sigue
    End If
    If campofecha < campofecha3 Then
            mensa = MsgBox("La Fecha pertenece a un mes anterior", vbCritical, "!! Atención !!")
        fechamal = 1
 Rem       Exit Sub
    End If
    
sigue:
    modifica = 0
    If Text8.Text = "S" Then
        modifica = 1
        nroasie = Text10.Text
        datmaestro.RecordSource = "select [Maestro Asientos].* from [Maestro Asientos] where empresa = " & login.empresaact & " and perinicial = '" & login.iper & "' and nroasiento = " & nroasie & "  order by nroasiento"
        datmaestro.Refresh
        If datmaestro.Recordset.EOF = True Then GoTo pas01
        masterasiento = datmaestro.Recordset.Fields(2)
        datasiento.RecordSource = "select [Detalle Asientos].* from [Detalle Asientos] where empresa = " & login.empresaact & " and idmasterasientos = " & masterasiento & ""
        datasiento.Refresh
        If datasiento.Recordset.EOF = True Then
            datmaestro.Recordset.Delete adAffectCurrent
            GoTo pas1
        End If
        datasiento.Recordset.MoveFirst
pas0:
        datasiento.Recordset.Delete adAffectCurrent
        datasiento.Recordset.MoveNext
        If datasiento.Recordset.EOF = True Then
            datmaestro.Recordset.Delete adAffectCurrent
            GoTo pas1
        End If
        GoTo pas0
    End If
pas01:
    datmaestro.RecordSource = "select [Maestro Asientos].* from [Maestro Asientos] where empresa = " & login.empresaact & " and perinicial = '" & login.iper & "' order by nroasiento"
    datmaestro.Refresh
    If datmaestro.Recordset.EOF = False Then
            datmaestro.Recordset.MoveLast
            nroasie = datmaestro.Recordset.Fields(3) + 1
    Else
            nroasie = 1
    End If
    
pas1:
    datmaestro.Recordset.AddNew
    If fechamal = 0 Then
            datmaestro.Recordset.Fields(0) = Maskfecha.Text
    Else
            datmaestro.Recordset.Fields(0) = fechafuera
    End If
    datmaestro.Recordset.Fields(1) = Date
    datmaestro.Recordset.Fields(3) = nroasie
    datmaestro.Recordset.Fields(4) = Left(denominacion.Text, 20) + " " + tipocomp.Text + " Nº:" + Maskcomprobante.Text
    datmaestro.Recordset.Fields(5) = Text5(0).Text
    datmaestro.Recordset.Fields(6) = Text5(1).Text
    datmaestro.Recordset.Fields(7) = login.empresaact
    datmaestro.Recordset.Fields(8) = "N"
    datmaestro.Recordset.Fields(9) = Val(datPrimaryRS.Recordset.Fields(0))
    datmaestro.Recordset.Fields(10) = "C"
    datmaestro.Recordset.Fields(11) = "S"
    datmaestro.Recordset.UpdateBatch adAffectCurrent
      
    If Left(tipocomp, 2) = "NC" Then
        Text3(15).Text = Text3(15).Text * -1
        For x = 0 To 14
            If Text3(x).Visible = True And Text3(x).Text <> "0" Then
                Text3(x).Text = Text3(x).Text * -1
            End If
        Next x
    End If
      
For x = 0 To 28 Step 2
    If Text7(x).Text <> "" Then
            If Text3(x / 2).Visible = False Then GoTo paso1
            datasiento.Recordset.AddNew
            datasiento.Recordset.Fields(5) = Val(datmaestro.Recordset.Fields(2))
            datasiento.Recordset.Fields(1) = datmaestro.Recordset.Fields(0).Value
            datasiento.Recordset.Fields(7) = login.empresaact
            datasiento.Recordset.Fields(2) = Text7(x).Text
            datasiento.Recordset.Fields(3) = Text3(x / 2).Text
            If datasiento.Recordset.Fields(3) < 0 Then
                datasiento.Recordset.Fields(4) = datasiento.Recordset.Fields(3) * -1
                datasiento.Recordset.Fields(3) = 0
            End If
            datasiento.Recordset.Fields(6) = Label1(x / 2).Caption
            If Text9.Text <> "" Then datasiento.Recordset.Fields(8) = Text9.Text
            datasiento.Recordset.UpdateBatch adAffectCurrent
    End If
Next x
paso1:
            datasiento.Recordset.AddNew
            datasiento.Recordset.Fields(5) = Val(datmaestro.Recordset.Fields(2))
            datasiento.Recordset.Fields(1) = datmaestro.Recordset.Fields(0).Value
            datasiento.Recordset.Fields(7) = login.empresaact
            datasiento.Recordset.Fields(2) = Text7(31).Text
            datasiento.Recordset.Fields(4) = Text3(15).Text
            If datasiento.Recordset.Fields(4) < 0 Then
                datasiento.Recordset.Fields(3) = datasiento.Recordset.Fields(4) * -1
                datasiento.Recordset.Fields(4) = 0
            End If
            datasiento.Recordset.Fields(6) = "Total facturado"
            datasiento.Recordset.UpdateBatch adAffectCurrent

    Text8.Text = "S"

    datPrimaryRS.Recordset.Fields(59) = nroasie
    datPrimaryRS.Recordset.UpdateBatch adAffectCurrent
    pos = datPrimaryRS.Recordset.AbsolutePosition
    datPrimaryRS.Refresh
    datPrimaryRS.Recordset.AbsolutePosition = pos
    
    Inicio.datauditoria.ConnectionString = login.conexiontotal
    
    Inicio.datauditoria.RecordSource = "select auditoria.* from auditoria"
    Inicio.datauditoria.Refresh
    
    Inicio.datauditoria.Recordset.AddNew
    Inicio.datauditoria.Recordset.Fields("fecha") = Date
    Inicio.datauditoria.Recordset.Fields("hora") = Str(Time)
    Inicio.datauditoria.Recordset.Fields("ventana") = "Libro Compras"
    Inicio.datauditoria.Recordset.Fields("accion") = "Alta/Modif.:" + tipocomp.Text + Maskcomprobante.Text + " Prov:" + Left(denominacion.Text, 15)
    Inicio.datauditoria.Recordset.Fields("usuario") = login.usuarioactivo
    Inicio.datauditoria.Recordset.Fields("empresa") = login.empresaact
    Inicio.datauditoria.Recordset.UpdateBatch adAffectCurrent
    
    nuevo.SetFocus
    
End Sub

Private Sub List1_KeyPress(KeyAscii As Integer)

    If KeyAscii = 45 Then
        DataList1.Visible = True
        DataList1.SetFocus
    End If
    If KeyAscii = 13 Then
        tipocomp.Text = List1.Text
        Maskcomprobante.SetFocus
    End If
                 

End Sub

Private Sub List1_LostFocus()
    If List1.Text = "" Then
        List1.SetFocus
        Exit Sub
    End If
    List1.Visible = False
End Sub



Private Sub Maskcomprobante_DragOver(Source As Control, x As Single, Y As Single, State As Integer)
    If List1.Text = "" Then
        List1.SetFocus
        Exit Sub
    End If

End Sub

Private Sub Maskcomprobante_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 45 Then
            List1.Visible = True
            List1.SetFocus
    End If
    If KeyAscii = 13 Then
        KeyAscii = 0
        
        datbusca.RecordSource = "select libroventas.* from libroventas WHERE libroventas.empresa = " & login.empresaact & " and cuit = '" & cuit.Text & "' and numcompr = '" & Maskcomprobante.Text & "' and tipocompr = '" & tipocomp.Text & "'  "
        datbusca.Refresh

        If datbusca.Recordset.EOF = False Then
            mensa = MsgBox("Este comprobante ya fue ingresado anteriormente, revise el nº de cuit, nº de comprobante, o tipo de comprobante", vbCritical, "!! Atención !!")
            Call Cancelar_Click
            Exit Sub
        End If
        
            Text4.Text = Maskcomprobante.Text
          Rem  mensa = MsgBox("Verifique si la factura es de Contado o Cta.Cte.", vbInformation, "Verificar")
            Text3(0).SelStart = 0
            Text3(0).SelLength = Len(Text3(0))
            Text3(0).SetFocus
            If Val(Mid(Text4.Text, 1, 4)) = 0 Then
                mensa = MsgBox("Debe ingresar una sucursal en el Nro de factura", vbCritical, "!! Atención !!")
                Maskcomprobante.SetFocus
                Maskcomprobante.SelStart = 0
                Maskcomprobante.SelLength = 4
                Exit Sub
            End If
            If Right(Maskcomprobante.Text, 1) = "_" Then
                mensa = MsgBox("Nro de factura incorrecto", vbCritical, "!! Atención !!")
                Maskcomprobante.SetFocus
                Maskcomprobante.SelStart = 5
                Maskcomprobante.SelLength = 8
                Exit Sub
            End If
    End If
    
    
End Sub


Private Sub Maskfecha_GotFocus()

  datclientes.RecordSource = "select clientes.* from clientes Where empresa = " & login.empresaact & " ORDER BY razonsocial"
  datclientes.Refresh
  


  sumadebe = 0
  sumahaber = 0
  
End Sub

Private Sub Maskfecha_KeyPress(KeyAscii As Integer)
Dim dia As Integer
Dim mes As Integer
Dim año As Integer

    If KeyAscii = 2 Then Text8.Text = "N"

    If KeyAscii = 13 Then
        Text2.Text = Maskfecha.Text
        dia = Day(Date)
        mes = Month(Date)
        año = Year(Date)
        If Val(Mid(Text2.Text, 1, 2)) > dia And Val(Mid(Text2.Text, 4, 2)) >= mes And Val(Mid(Text2.Text, 7, 4)) >= año Then
                mensa = MsgBox("El Día ingresado es mayor al de la fecha actual", vbCritical, "!! Atención !!")
                Maskfecha.SetFocus
                Maskfecha.SelStart = 0
                Maskfecha.SelLength = 2
                Exit Sub
        End If
        If Val(Mid(Text2.Text, 4, 2)) > mes And Val(Mid(Text2.Text, 7, 4)) >= año Then
                mensa = MsgBox("El Mes ingresado es mayor al de la fecha actual", vbCritical, "!! Atención !!")
                Maskfecha.SetFocus
                Maskfecha.SelStart = 3
                Maskfecha.SelLength = 2
                Exit Sub
        End If
        If Val(Mid(Text2.Text, 7, 4)) > año Then
                mensa = MsgBox("El Año ingresado es mayor al de la fecha actual", vbCritical, "!! Atención !!")
                Maskfecha.SetFocus
                Maskfecha.SelStart = 6
                Maskfecha.SelLength = 4
                Exit Sub
        End If
        denominacion.SetFocus
    End If
    

End Sub

Private Sub nuevo_Click()
On Error GoTo errornuevo

    
    If errorasiento = True Then Exit Sub
        
    datPrimaryRS.Recordset.AddNew
    For x = 0 To 14
            Text3(x).Text = 0
    Next x
    For x = 0 To 30
            Text7(x).Text = ""
    Next x
  Rem  Text1.Text = login.empresaact
    
    Maskfecha.SelLength = 10
    Maskfecha.SelText = ""
    Maskcomprobante.SelLength = 13
    Maskcomprobante.SelText = ""
    Maskfecha.SetFocus
    Exit Sub

errornuevo:

End Sub

Private Sub nuevo_GotFocus()

     Maskfecha.Mask = "##/##/####"
     Maskfecha.MaxLength = 10
     Maskcomprobante.Mask = "####-########"
     Maskcomprobante.MaxLength = 13

End Sub

Private Sub Option1_Click(Index As Integer)


    If Option1(0).Value = True Then
        datPrimaryRS.RecordSource = "select libroventas.* from libroventas WHERE libroventas.empresa = " & login.empresaact & " and inicioper = '" & login.iper & "' and finper = '" & login.fper & "' and fecha >= '" & desde.Value & "' and fecha <= '" & hasta.Value & "' Order by fecha"
        datPrimaryRS.Refresh
    End If
    If Option1(1).Value = True Then
        datPrimaryRS.RecordSource = "select libroventas.* from libroventas WHERE libroventas.empresa = " & login.empresaact & " and inicioper = '" & login.iper & "' and finper = '" & login.fper & "' and fecha >= '" & desde.Value & "' and fecha <= '" & hasta.Value & "' Order by cliente"
        datPrimaryRS.Refresh
    End If
    If Option1(2).Value = True Then
        datPrimaryRS.RecordSource = "select libroventas.* from libroventas WHERE libroventas.empresa = " & login.empresaact & " and inicioper = '" & login.iper & "' and finper = '" & login.fper & "' and fecha >= '" & desde.Value & "' and fecha <= '" & hasta.Value & "' Order by tipocompr"
        datPrimaryRS.Refresh
    End If
    If Option1(3).Value = True Then
        datPrimaryRS.RecordSource = "select libroventas.* from libroventas WHERE libroventas.empresa = " & login.empresaact & " and inicioper = '" & login.iper & "' and finper = '" & login.fper & "' and fecha >= '" & desde.Value & "' and fecha <= '" & hasta.Value & "' Order by numcompr"
        datPrimaryRS.Refresh
    End If


End Sub

Private Sub salir_Click()

    Unload Me

End Sub


Private Sub Text13_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        Text3(0).SetFocus
        Text13.Visible = True
    End If
    
End Sub

Private Sub Text13_LostFocus()

    Text13.Visible = False

End Sub

Private Sub Text3_GotFocus(Index As Integer)
On Error GoTo errorlist

    If List1.Text = "" Then
        List1.SetFocus
        Exit Sub
    End If

                Cuenta = 0
                Text3(Index).SelStart = 0
                Text3(Index).SelLength = Len(Text3(Index)) + 3
errorlist:
    Exit Sub
End Sub

Private Sub Text3_KeyPress(Index As Integer, KeyAscii As Integer)
On Error GoTo fuera
Dim suma As Double
    
    If KeyAscii = 45 Then
            KeyAscii = 0
            If Index > 0 Then
                If Text3(Index - 1).Visible = True Then
                    Text3(Index - 1).SetFocus
                    Text3(Index - 1).SelStart = 0
                    Text3(Index - 1).SelLength = Len(Text3(Index - 1))
                End If
            End If
    End If
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        If Text3(Index).Text = "" Then Text3(Index).Text = 0
        suma = 0

         If automatico = 1 Then Call calcular_Click
                                                            
         For x = 0 To 14
            If Text3(x).Visible = True Then suma = Text3(x).Text + suma
         Next x
         Text3(15).Text = suma
         Text6.Text = "N"
         
        datPrimaryRS.Recordset.Fields(61) = login.iper
        datPrimaryRS.Recordset.Fields(62) = login.fper
        If Check1(1).Value = 1 Then
            datPrimaryRS.Recordset.Fields("contado") = "S"
        Else
            datPrimaryRS.Recordset.Fields("contado") = ""
        End If
         datPrimaryRS.Recordset.UpdateBatch adAffectCurrent
         pos = datPrimaryRS.Recordset.AbsolutePosition
         datPrimaryRS.Refresh
         datPrimaryRS.Recordset.AbsolutePosition = pos
         
         
         If Text3(Index).Text > 0 Then
            posicion = Index
            Cuenta = 0
            Text7(Index * 2).SetFocus
            Exit Sub
         End If
         If Text3(Index + 1).Visible = True Then
                Text3(Index + 1).SetFocus
         Else
                Text7(31).SetFocus
         End If
         DataGrid1.Refresh
     End If
Exit Sub
fuera:
mensa = MsgBox("Algún Campo no fue ingresado, o ingreso un caracter incorrecto", vbCritical, "!! Error ¡¡")
    
End Sub
Private Sub Text7_GotFocus(Index As Integer)

    prueba = datcolumnas.Recordset.Fields(Index + 31)
    If prueba > 0 Then
           Text7(Index).Text = prueba
           sumadebe = Text3(posicion) + sumadebe
           Text7(Index + 1).Text = 0
           If Text3(posicion + 1).Visible = True Then
                Text3(posicion + 1).SetFocus
                Exit Sub
           End If
           Text7(31).SetFocus
           Exit Sub
    End If

    poscuenta = Index
    Text7(Index).SelLength = Len(Text7(Index))


End Sub

Private Sub textcuenta_GotFocus()
    
    DataList2.Visible = True
    DataList2.SetFocus
    
End Sub

Private Sub Text7_LostFocus(Index As Integer)
    poscuenta = Index
    If Index = 31 Then
        If Text7(Index).Text = "" Then Exit Sub
        sumahaber = Text3(15) + sumahaber
        Call grabalibroasiento_Click
        Exit Sub
    End If
    
           
End Sub

