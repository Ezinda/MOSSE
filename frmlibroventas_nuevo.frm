VERSION 5.00
Object = "{F6125AB1-8AB1-11CE-A77F-08002B2F4E98}#2.0#0"; "MSRDC20.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{5734474E-78D3-4254-99B9-C35F31BDF509}#62.0#0"; "vbskpro2.ocx"
Object = "{D18BBD1F-82BB-4385-BED3-E9D31A3E361E}#1.0#0"; "KewlButtonz.ocx"
Begin VB.Form frmlibroventas_nuevo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Comprobantes de Venta"
   ClientHeight    =   6795
   ClientLeft      =   1095
   ClientTop       =   330
   ClientWidth     =   13095
   ControlBox      =   0   'False
   Icon            =   "frmlibroventas_nuevo.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6795
   ScaleWidth      =   13095
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame4 
      Caption         =   "Tipo de Comprobante:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6735
      Left            =   9600
      TabIndex        =   123
      Top             =   0
      Visible         =   0   'False
      Width           =   3375
      Begin VB.CommandButton Command4 
         Caption         =   "Fecha:"
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   139
         TabStop         =   0   'False
         Top             =   3960
         Width           =   735
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Hasta:"
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   138
         TabStop         =   0   'False
         Top             =   3480
         Width           =   735
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Desde:"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   137
         TabStop         =   0   'False
         Top             =   3000
         Width           =   735
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Hasta:"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   136
         TabStop         =   0   'False
         Top             =   1920
         Width           =   735
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Desde:"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   135
         TabStop         =   0   'False
         Top             =   1440
         Width           =   735
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Tipo de Comprobante:"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   134
         TabStop         =   0   'False
         Top             =   960
         Width           =   1695
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1920
         TabIndex        =   131
         Text            =   "Combo1"
         Top             =   960
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Emitidas"
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
         TabIndex        =   130
         Top             =   480
         Width           =   1335
      End
      Begin VB.OptionButton Option2 
         Caption         =   "No Emitidas"
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
         TabIndex        =   129
         Top             =   2640
         Width           =   1335
      End
      Begin VB.TextBox Text15 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   960
         TabIndex        =   127
         Top             =   3000
         Width           =   1335
      End
      Begin VB.CommandButton anulamasiva 
         Caption         =   "ANULAR"
         Height          =   615
         Left            =   240
         Picture         =   "frmlibroventas_nuevo.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   126
         Top             =   4800
         UseMaskColor    =   -1  'True
         Width           =   2775
      End
      Begin VB.CommandButton filtro 
         Caption         =   "filtro"
         Height          =   495
         Left            =   2040
         TabIndex        =   124
         Top             =   240
         Visible         =   0   'False
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker fechaanul 
         Height          =   375
         Left            =   960
         TabIndex        =   125
         Top             =   3960
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   156368897
         CurrentDate     =   38993
      End
      Begin MSDataListLib.DataCombo DataCombo2 
         Bindings        =   "frmlibroventas_nuevo.frx":0E44
         Height          =   315
         Left            =   840
         TabIndex        =   128
         Top             =   1440
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "numcompr"
         Text            =   "DataCombo2"
      End
      Begin MSDataListLib.DataCombo DataCombo3 
         Bindings        =   "frmlibroventas_nuevo.frx":0E5F
         Height          =   315
         Left            =   840
         TabIndex        =   132
         Top             =   1920
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "numcompr"
         Text            =   "DataCombo3"
      End
      Begin MSMask.MaskEdBox MaskEdBox2 
         Height          =   375
         Left            =   960
         TabIndex        =   133
         Top             =   3480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   13
         Mask            =   "####-########"
         PromptChar      =   "_"
      End
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00FFFFFF&
      Height          =   1815
      ItemData        =   "frmlibroventas_nuevo.frx":0E7A
      Left            =   2040
      List            =   "frmlibroventas_nuevo.frx":0E7C
      Sorted          =   -1  'True
      TabIndex        =   120
      Top             =   1440
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox List2 
      BackColor       =   &H80000016&
      Height          =   1620
      Left            =   6840
      TabIndex        =   50
      Top             =   3480
      Visible         =   0   'False
      Width           =   2655
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
      Height          =   6735
      Left            =   8160
      TabIndex        =   36
      Top             =   0
      Width           =   1335
      Begin KewlButtonz.KewlButtons cancelar 
         Cancel          =   -1  'True
         Height          =   615
         Left            =   120
         TabIndex        =   28
         Top             =   360
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   1085
         BTYPE           =   14
         TX              =   "&Cancelar"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
         MICON           =   "frmlibroventas_nuevo.frx":0E7E
         PICN            =   "frmlibroventas_nuevo.frx":0E9A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin KewlButtonz.KewlButtons borrar 
         Height          =   615
         Left            =   120
         TabIndex        =   29
         Top             =   3240
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   1085
         BTYPE           =   14
         TX              =   "&Elimnar"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
         MICON           =   "frmlibroventas_nuevo.frx":18AC
         PICN            =   "frmlibroventas_nuevo.frx":18C8
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin KewlButtonz.KewlButtons salir 
         Height          =   615
         Left            =   120
         TabIndex        =   30
         Top             =   5760
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   1085
         BTYPE           =   14
         TX              =   "&Salir"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
         MICON           =   "frmlibroventas_nuevo.frx":4CBA
         PICN            =   "frmlibroventas_nuevo.frx":4CD6
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin KewlButtonz.KewlButtons Command3 
         Height          =   615
         Left            =   120
         TabIndex        =   26
         Top             =   4080
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   1085
         BTYPE           =   14
         TX              =   "&Imprime Comp."
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
         MICON           =   "frmlibroventas_nuevo.frx":5820
         PICN            =   "frmlibroventas_nuevo.frx":583C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin KewlButtonz.KewlButtons imprimir 
         Height          =   615
         Left            =   120
         TabIndex        =   27
         Top             =   4920
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   1085
         BTYPE           =   14
         TX              =   "&Buscar"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
         MICON           =   "frmlibroventas_nuevo.frx":624E
         PICN            =   "frmlibroventas_nuevo.frx":626A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin KewlButtonz.KewlButtons anular 
         Height          =   615
         Left            =   120
         TabIndex        =   121
         Top             =   1320
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   1085
         BTYPE           =   14
         TX              =   "&Anular"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
         MICON           =   "frmlibroventas_nuevo.frx":965C
         PICN            =   "frmlibroventas_nuevo.frx":9678
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin KewlButtonz.KewlButtons anulmasiva 
         Height          =   615
         Left            =   120
         TabIndex        =   122
         Top             =   2280
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   1085
         BTYPE           =   14
         TX              =   "Anulc. &Masiva"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
         MICON           =   "frmlibroventas_nuevo.frx":A08A
         PICN            =   "frmlibroventas_nuevo.frx":A0A6
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
   Begin MSDataListLib.DataList DataList1 
      Bindings        =   "frmlibroventas_nuevo.frx":AAB8
      Height          =   2205
      Left            =   1440
      TabIndex        =   32
      Top             =   1080
      Visible         =   0   'False
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   3889
      _Version        =   393216
      IntegralHeight  =   0   'False
      MatchEntry      =   -1  'True
      BackColor       =   16777215
      ListField       =   "razonsocial"
   End
   Begin VB.TextBox Text9 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   600
      TabIndex        =   45
      Top             =   3120
      Visible         =   0   'False
      Width           =   615
   End
   Begin MSDataListLib.DataList DataList3 
      Bindings        =   "frmlibroventas_nuevo.frx":AAD5
      Height          =   1560
      Left            =   480
      TabIndex        =   44
      Top             =   3600
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   2752
      _Version        =   393216
      MatchEntry      =   -1  'True
      BackColor       =   16777215
      ForeColor       =   -2147483647
      ListField       =   "ccostoslista"
      BoundColumn     =   "cc"
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
   Begin VB.Frame Frame3 
      BackColor       =   &H80000003&
      Caption         =   "Centros de Costo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   240
      TabIndex        =   46
      Top             =   2760
      Visible         =   0   'False
      Width           =   3135
   End
   Begin MSDataListLib.DataList DataList2 
      Bindings        =   "frmlibroventas_nuevo.frx":AAF2
      Height          =   1260
      Left            =   840
      TabIndex        =   43
      Top             =   5160
      Visible         =   0   'False
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   2223
      _Version        =   393216
      MatchEntry      =   -1  'True
      BackColor       =   -2147483626
      ForeColor       =   -2147483647
      ListField       =   "codigo"
      BoundColumn     =   "Cod Contable"
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
   Begin VB.Frame Frame5 
      Height          =   4215
      Left            =   120
      TabIndex        =   63
      Top             =   2520
      Width           =   7935
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   195
         Left            =   1440
         TabIndex        =   117
         Top             =   4200
         Width           =   735
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """$"" #,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   2
         EndProperty
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
         Left            =   6120
         Locked          =   -1  'True
         TabIndex        =   114
         TabStop         =   0   'False
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox Text7 
         Alignment       =   2  'Center
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
         Left            =   4560
         TabIndex        =   112
         Text            =   " "
         Top             =   330
         Width           =   735
      End
      Begin VB.TextBox Text7 
         Alignment       =   2  'Center
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
         Left            =   4560
         TabIndex        =   110
         Text            =   " "
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox Text7 
         Alignment       =   2  'Center
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
         Left            =   4560
         TabIndex        =   108
         Text            =   " "
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox Text7 
         Alignment       =   2  'Center
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
         Left            =   4560
         TabIndex        =   106
         Text            =   " "
         Top             =   1080
         Width           =   735
      End
      Begin VB.TextBox Text7 
         Alignment       =   2  'Center
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
         Left            =   4560
         TabIndex        =   104
         Text            =   " "
         Top             =   1320
         Width           =   735
      End
      Begin VB.TextBox Text7 
         Alignment       =   2  'Center
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
         Left            =   4560
         TabIndex        =   102
         Text            =   " "
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox Text7 
         Alignment       =   2  'Center
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
         Left            =   4560
         TabIndex        =   100
         Text            =   " "
         Top             =   1800
         Width           =   735
      End
      Begin VB.TextBox Text7 
         Alignment       =   2  'Center
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
         Left            =   4560
         TabIndex        =   98
         Text            =   " "
         Top             =   2040
         Width           =   735
      End
      Begin VB.TextBox Text7 
         Alignment       =   2  'Center
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
         Left            =   4560
         TabIndex        =   96
         Text            =   " "
         Top             =   2280
         Width           =   735
      End
      Begin VB.TextBox Text7 
         Alignment       =   2  'Center
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
         Left            =   4560
         TabIndex        =   94
         Text            =   " "
         Top             =   2520
         Width           =   735
      End
      Begin VB.TextBox Text7 
         Alignment       =   2  'Center
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
         Left            =   4560
         TabIndex        =   92
         Text            =   " "
         Top             =   2760
         Width           =   735
      End
      Begin VB.TextBox Text7 
         Alignment       =   2  'Center
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
         Left            =   4560
         TabIndex        =   90
         Text            =   " "
         Top             =   3000
         Width           =   735
      End
      Begin VB.TextBox Text7 
         Alignment       =   2  'Center
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
         Left            =   4560
         TabIndex        =   88
         Text            =   " "
         Top             =   3240
         Width           =   735
      End
      Begin VB.TextBox Text7 
         Alignment       =   2  'Center
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
         Left            =   4560
         TabIndex        =   86
         Text            =   " "
         Top             =   3480
         Width           =   735
      End
      Begin VB.TextBox Text7 
         Alignment       =   2  'Center
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
         Left            =   4560
         TabIndex        =   84
         Text            =   " "
         Top             =   3720
         Width           =   735
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """$"" #,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   2
         EndProperty
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
         Left            =   3240
         TabIndex        =   10
         Top             =   330
         Width           =   1215
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """$"" #,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   2
         EndProperty
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
         Left            =   3240
         TabIndex        =   11
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """$"" #,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   2
         EndProperty
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
         Left            =   3240
         TabIndex        =   12
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """$"" #,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   2
         EndProperty
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
         Left            =   3240
         TabIndex        =   13
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """$"" #,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   2
         EndProperty
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
         Left            =   3240
         TabIndex        =   14
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """$"" #,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   2
         EndProperty
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
         Left            =   3240
         TabIndex        =   15
         Top             =   1560
         Width           =   1215
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """$"" #,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   2
         EndProperty
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
         Left            =   3240
         TabIndex        =   16
         Top             =   1800
         Width           =   1215
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """$"" #,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   2
         EndProperty
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
         Left            =   3240
         TabIndex        =   17
         Top             =   2040
         Width           =   1215
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """$"" #,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   2
         EndProperty
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
         Left            =   3240
         TabIndex        =   18
         Top             =   2280
         Width           =   1215
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """$"" #,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   2
         EndProperty
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
         Left            =   3240
         TabIndex        =   19
         Top             =   2520
         Width           =   1215
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """$"" #,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   2
         EndProperty
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
         Left            =   3240
         TabIndex        =   20
         Top             =   2760
         Width           =   1215
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """$"" #,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   2
         EndProperty
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
         Left            =   3240
         TabIndex        =   21
         Top             =   3000
         Width           =   1215
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """$"" #,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   2
         EndProperty
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
         Left            =   3240
         TabIndex        =   22
         Top             =   3240
         Width           =   1215
      End
      Begin VB.CommandButton Command5 
         Caption         =   "C.Haber"
         Height          =   255
         Index           =   7
         Left            =   4440
         TabIndex        =   81
         TabStop         =   0   'False
         Top             =   120
         Width           =   855
      End
      Begin VB.CommandButton Command5 
         Caption         =   "C.Debe"
         Height          =   255
         Index           =   8
         Left            =   5760
         TabIndex        =   80
         TabStop         =   0   'False
         Top             =   1200
         Width           =   975
      End
      Begin VB.CommandButton label1 
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   79
         TabStop         =   0   'False
         Top             =   360
         Width           =   3135
      End
      Begin VB.CommandButton label1 
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   78
         TabStop         =   0   'False
         Top             =   600
         Width           =   3135
      End
      Begin VB.CommandButton label1 
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   77
         TabStop         =   0   'False
         Top             =   840
         Width           =   3135
      End
      Begin VB.CommandButton label1 
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   76
         TabStop         =   0   'False
         Top             =   1080
         Width           =   3135
      End
      Begin VB.CommandButton label1 
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   75
         TabStop         =   0   'False
         Top             =   1320
         Width           =   3135
      End
      Begin VB.CommandButton label1 
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   74
         TabStop         =   0   'False
         Top             =   1560
         Width           =   3135
      End
      Begin VB.CommandButton label1 
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   73
         TabStop         =   0   'False
         Top             =   1800
         Width           =   3135
      End
      Begin VB.CommandButton label1 
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   72
         TabStop         =   0   'False
         Top             =   2040
         Width           =   3135
      End
      Begin VB.CommandButton label1 
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   71
         TabStop         =   0   'False
         Top             =   2280
         Width           =   3135
      End
      Begin VB.CommandButton label1 
         Height          =   255
         Index           =   9
         Left            =   120
         TabIndex        =   70
         TabStop         =   0   'False
         Top             =   2520
         Width           =   3135
      End
      Begin VB.CommandButton label1 
         Height          =   255
         Index           =   10
         Left            =   120
         TabIndex        =   69
         TabStop         =   0   'False
         Top             =   2760
         Width           =   3135
      End
      Begin VB.CommandButton label1 
         Height          =   255
         Index           =   11
         Left            =   120
         TabIndex        =   68
         TabStop         =   0   'False
         Top             =   3000
         Width           =   3135
      End
      Begin VB.CommandButton label1 
         Height          =   255
         Index           =   12
         Left            =   120
         TabIndex        =   67
         TabStop         =   0   'False
         Top             =   3240
         Width           =   3135
      End
      Begin VB.CommandButton label1 
         Height          =   255
         Index           =   13
         Left            =   120
         TabIndex        =   66
         TabStop         =   0   'False
         Top             =   3480
         Width           =   3135
      End
      Begin VB.CommandButton label1 
         Height          =   255
         Index           =   14
         Left            =   120
         TabIndex        =   65
         TabStop         =   0   'False
         Top             =   3720
         Width           =   3135
      End
      Begin VB.CommandButton label1 
         Caption         =   "Total"
         Height          =   255
         Index           =   15
         Left            =   5760
         TabIndex        =   64
         TabStop         =   0   'False
         Top             =   360
         Width           =   1935
      End
      Begin VB.TextBox Text7 
         Alignment       =   2  'Center
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
         Left            =   6840
         TabIndex        =   83
         Text            =   " "
         Top             =   1200
         Width           =   615
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """$"" #,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   2
         EndProperty
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
         Left            =   3240
         TabIndex        =   23
         Top             =   3480
         Width           =   1215
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """$"" #,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   2
         EndProperty
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
         Left            =   3240
         TabIndex        =   24
         Top             =   3720
         Width           =   1215
      End
      Begin KewlButtonz.KewlButtons grabalibroasiento 
         Height          =   735
         Left            =   6240
         TabIndex        =   25
         Top             =   1680
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   1296
         BTYPE           =   14
         TX              =   "&Grabar"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
         MICON           =   "frmlibroventas_nuevo.frx":AB0B
         PICN            =   "frmlibroventas_nuevo.frx":AB27
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.TextBox Text7 
         Alignment       =   2  'Center
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
         Left            =   4560
         TabIndex        =   85
         Text            =   " "
         Top             =   3720
         Width           =   735
      End
      Begin VB.TextBox Text7 
         Alignment       =   2  'Center
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
         Left            =   4560
         TabIndex        =   87
         Text            =   " "
         Top             =   3480
         Width           =   735
      End
      Begin VB.TextBox Text7 
         Alignment       =   2  'Center
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
         Left            =   4560
         TabIndex        =   89
         Text            =   " "
         Top             =   3240
         Width           =   735
      End
      Begin VB.TextBox Text7 
         Alignment       =   2  'Center
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
         Left            =   4560
         TabIndex        =   91
         Text            =   " "
         Top             =   3000
         Width           =   735
      End
      Begin VB.TextBox Text7 
         Alignment       =   2  'Center
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
         Left            =   4560
         TabIndex        =   93
         Text            =   " "
         Top             =   2760
         Width           =   735
      End
      Begin VB.TextBox Text7 
         Alignment       =   2  'Center
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
         Left            =   4560
         TabIndex        =   95
         Text            =   " "
         Top             =   2520
         Width           =   735
      End
      Begin VB.TextBox Text7 
         Alignment       =   2  'Center
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
         Left            =   4560
         TabIndex        =   97
         Text            =   " "
         Top             =   2280
         Width           =   735
      End
      Begin VB.TextBox Text7 
         Alignment       =   2  'Center
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
         Left            =   4560
         TabIndex        =   99
         Text            =   " "
         Top             =   2040
         Width           =   735
      End
      Begin VB.TextBox Text7 
         Alignment       =   2  'Center
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
         Left            =   4560
         TabIndex        =   101
         Text            =   " "
         Top             =   1800
         Width           =   735
      End
      Begin VB.TextBox Text7 
         Alignment       =   2  'Center
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
         Left            =   4560
         TabIndex        =   103
         Text            =   " "
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox Text7 
         Alignment       =   2  'Center
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
         Left            =   4560
         TabIndex        =   105
         Text            =   " "
         Top             =   1320
         Width           =   735
      End
      Begin VB.TextBox Text7 
         Alignment       =   2  'Center
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
         Left            =   4560
         TabIndex        =   107
         Text            =   " "
         Top             =   1080
         Width           =   735
      End
      Begin VB.TextBox Text7 
         Alignment       =   2  'Center
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
         Left            =   4560
         TabIndex        =   109
         Text            =   " "
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox Text7 
         Alignment       =   2  'Center
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
         Left            =   4560
         TabIndex        =   111
         Text            =   " "
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox Text7 
         Alignment       =   2  'Center
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
         Left            =   4560
         TabIndex        =   113
         Top             =   345
         Width           =   735
      End
      Begin VB.TextBox Text7 
         Alignment       =   2  'Center
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
         Left            =   6240
         Locked          =   -1  'True
         TabIndex        =   82
         Text            =   " "
         Top             =   1200
         Visible         =   0   'False
         Width           =   735
      End
   End
   Begin VB.CommandButton compfecha 
      Caption         =   "compfecha"
      Height          =   375
      Left            =   8040
      TabIndex        =   53
      Top             =   6720
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Text6 
      DataField       =   "cerrado"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Left            =   8640
      TabIndex        =   52
      Text            =   "Text6"
      Top             =   5040
      Visible         =   0   'False
      Width           =   855
   End
   Begin Crystal.CrystalReport CrystalReporte 
      Left            =   8040
      Top             =   6000
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
   End
   Begin VB.CommandButton calcular2 
      Caption         =   "calcular2"
      Height          =   255
      Left            =   8040
      TabIndex        =   51
      Top             =   6480
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Text10 
      DataField       =   "asiento"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Left            =   8640
      TabIndex        =   47
      Text            =   "Text10"
      Top             =   4560
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Text8 
      DataField       =   "asentado"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Left            =   8160
      TabIndex        =   42
      Text            =   "Text8"
      Top             =   4560
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Index           =   1
      Left            =   8160
      TabIndex        =   41
      Text            =   "Text5"
      Top             =   5040
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Index           =   0
      Left            =   8160
      TabIndex        =   40
      Text            =   "Text5"
      Top             =   4800
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton calcular 
      Caption         =   "calcular"
      Height          =   255
      Left            =   0
      TabIndex        =   39
      Top             =   6720
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      DataField       =   "empresa"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Left            =   8640
      TabIndex        =   31
      Text            =   "Text1"
      Top             =   4800
      Visible         =   0   'False
      Width           =   735
   End
   Begin MSDataGridLib.DataGrid DataGrid2 
      Bindings        =   "frmlibroventas_nuevo.frx":C5A9
      Height          =   495
      Left            =   7680
      TabIndex        =   35
      Top             =   5760
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
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
   Begin MSAdodcLib.Adodc datcolumnas 
      Height          =   330
      Left            =   7680
      Top             =   6720
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
      Left            =   8160
      TabIndex        =   34
      Text            =   "Text2"
      Top             =   5280
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   8160
      TabIndex        =   37
      Text            =   "Text4"
      Top             =   5520
      Visible         =   0   'False
      Width           =   975
   End
   Begin MSDataGridLib.DataGrid DataGrid3 
      Bindings        =   "frmlibroventas_nuevo.frx":C5C3
      Height          =   735
      Left            =   7320
      TabIndex        =   38
      Top             =   6360
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
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
   Begin MSAdodcLib.Adodc datproveedores 
      Height          =   330
      Left            =   720
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
      BackColor       =   &H00E0E0E0&
      Height          =   2415
      Left            =   120
      TabIndex        =   33
      Top             =   120
      Width           =   7935
      Begin VB.CommandButton Command4 
         Caption         =   "Reg."
         Height          =   255
         Index           =   2
         Left            =   6360
         TabIndex        =   119
         TabStop         =   0   'False
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox Text11 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   6960
         Locked          =   -1  'True
         TabIndex        =   118
         TabStop         =   0   'False
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox Text13 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   645
         Left            =   240
         MaxLength       =   255
         MultiLine       =   -1  'True
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   2280
         Visible         =   0   'False
         Width           =   7455
      End
      Begin VB.CommandButton Command5 
         Caption         =   "No"
         Height          =   255
         Index           =   10
         Left            =   6360
         TabIndex        =   116
         TabStop         =   0   'False
         Top             =   1200
         Width           =   495
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Si"
         Height          =   255
         Index           =   9
         Left            =   5400
         TabIndex        =   115
         TabStop         =   0   'False
         Top             =   1200
         Width           =   495
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Gestion de Clientes"
         Height          =   255
         Index           =   6
         Left            =   5160
         TabIndex        =   62
         TabStop         =   0   'False
         Top             =   960
         Width           =   2295
      End
      Begin VB.CheckBox automatico 
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   4560
         TabIndex        =   61
         TabStop         =   0   'False
         Top             =   1800
         Width           =   255
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Activar Calculo Automatico "
         Height          =   255
         Index           =   4
         Left            =   4920
         TabIndex        =   60
         TabStop         =   0   'False
         Top             =   1800
         Width           =   2655
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Empresa"
         Height          =   255
         Index           =   1
         Left            =   3480
         TabIndex        =   59
         TabStop         =   0   'False
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Nro.Comp."
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   58
         TabStop         =   0   'False
         Top             =   1680
         Width           =   975
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Tipo.Comp."
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   57
         TabStop         =   0   'False
         Top             =   1320
         Width           =   975
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Dat.Fiscal."
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   56
         TabStop         =   0   'False
         Top             =   960
         Width           =   975
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Cliente"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   55
         TabStop         =   0   'False
         Top             =   600
         Width           =   975
      End
      Begin VB.CommandButton Command4 
         Caption         =   "&Fecha"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   54
         TabStop         =   0   'False
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox denominacion 
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   1320
         TabIndex        =   1
         Top             =   600
         Width           =   4935
      End
      Begin VB.TextBox tipoiva 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   1320
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   960
         Width           =   495
      End
      Begin VB.TextBox cuit 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   1920
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   960
         Width           =   1575
      End
      Begin VB.TextBox tipocomp 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   1320
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   1320
         Width           =   615
      End
      Begin VB.TextBox Text14 
         Height          =   285
         Left            =   2880
         TabIndex        =   5
         Top             =   1320
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   255
         Index           =   0
         Left            =   6000
         TabIndex        =   7
         Top             =   1200
         Width           =   255
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   255
         Index           =   1
         Left            =   6960
         TabIndex        =   8
         Top             =   1200
         Width           =   255
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "frmlibroventas_nuevo.frx":C5DC
         Height          =   315
         Left            =   4560
         TabIndex        =   48
         Top             =   240
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         BackColor       =   -2147483626
         ListField       =   "razonsocial"
         BoundColumn     =   "empresa"
         Text            =   ""
      End
      Begin MSMask.MaskEdBox Maskfecha 
         Height          =   255
         Left            =   1320
         TabIndex        =   0
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
         AllowPrompt     =   -1  'True
         Format          =   "dd/mm/yyyy"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Maskcomprobante 
         Height          =   255
         Left            =   1320
         TabIndex        =   6
         Top             =   1680
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   450
         _Version        =   393216
         AllowPrompt     =   -1  'True
         PromptChar      =   "_"
      End
   End
   Begin MSAdodcLib.Adodc datmaestro 
      Height          =   330
      Left            =   4080
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
   Begin MSAdodcLib.Adodc datasiento 
      Height          =   330
      Left            =   1920
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
   Begin MSAdodcLib.Adodc datlistacostos 
      Height          =   330
      Left            =   5280
      Top             =   6720
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
      Left            =   6480
      Top             =   6720
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
      Left            =   2880
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
   Begin vbskpro.Skinner Skinner1 
      Left            =   8040
      Top             =   6480
      _ExtentX        =   1270
      _ExtentY        =   1270
      CloseButton     =   0
      MaxButton       =   0
      MinButton       =   0
      CloseButtonToolTipText=   "Cerrar"
      MinButtonToolTipText=   "Minimizar"
      MaxButtonToolTipText=   "Maximizar"
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
      LcK2            =   $"frmlibroventas_nuevo.frx":C5F6
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
      Left            =   1200
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
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
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
   Begin MSAdodcLib.Adodc datparamgral 
      Height          =   330
      Left            =   0
      Top             =   0
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
   Begin MSAdodcLib.Adodc datparamventas 
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
   Begin MSRDC.MSRDC reporte 
      Height          =   375
      Left            =   0
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
   Begin MSAdodcLib.Adodc datcuentas 
      Height          =   330
      Left            =   960
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grilla 
      Height          =   2175
      Left            =   5760
      TabIndex        =   49
      Top             =   4080
      Visible         =   0   'False
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   3836
      _Version        =   393216
      Rows            =   14
      Cols            =   8
      FixedRows       =   0
      FixedCols       =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   8
   End
   Begin MSAdodcLib.Adodc datPrimaryRS 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      Top             =   6465
      Visible         =   0   'False
      Width           =   13095
      _ExtentX        =   23098
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
   Begin MSAdodcLib.Adodc datempresa1 
      Height          =   330
      Left            =   0
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
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
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
   Begin MSAdodcLib.Adodc datordenes 
      Height          =   330
      Left            =   2640
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
      Connect         =   "Provider=MSDASQL.1;Password=25072004;Persist Security Info=True;User ID=lucva;Data Source=contable;Initial Catalog=ordenesradio"
      OLEDBString     =   "Provider=MSDASQL.1;Password=25072004;Persist Security Info=True;User ID=lucva;Data Source=contable;Initial Catalog=ordenesradio"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select ordenes.* from ordenes"
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
   Begin MSAdodcLib.Adodc dathistoria 
      Height          =   330
      Left            =   3840
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
      Connect         =   "Provider=MSDASQL.1;Password=25072004;Persist Security Info=True;User ID=lucva;Data Source=contable;Initial Catalog=ordenesradio"
      OLEDBString     =   "Provider=MSDASQL.1;Password=25072004;Persist Security Info=True;User ID=lucva;Data Source=contable;Initial Catalog=ordenesradio"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select historialorden.* from historialorden"
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
   Begin MSAdodcLib.Adodc datfacclientes 
      Height          =   330
      Left            =   5280
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
End
Attribute VB_Name = "frmlibroventas_nuevo"
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
Dim empresareal As Integer
Dim codgastos As Integer
Public indice As Integer
Public librocontado As Integer
Dim previsualiza As Integer
Dim facturaimprime As Double
Dim flagbuscar As Integer
Dim fechamal As Integer
Public modificar As Integer
Dim ter(15, 15), sig(15, 15) As String



Private Sub Adodc1_WillMove(ByVal adReason As ADODB.EventReasonEnum, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)

End Sub

Private Sub anulamasiva_Click()
On Error GoTo salir

If Option2.Value = True Then
    primera1 = Left(Text15.Text, 4)
    primera2 = Left(MaskEdBox2.Text, 4)
    
    If primera1 <> primera2 Then
        mensa = MsgBox("Factura invalida, ingrese nuevamente", vbCritical, "Error")
        MaskEdBox2.SetFocus
        Exit Sub
    End If
    segunda1 = Val(Right(Text15.Text, 8))
    segunda2 = Val(Right(MaskEdBox2.Text, 8))
Else
    segunda1 = DataCombo2.Text
    segunda2 = DataCombo3.Text
End If
    
    If segunda2 < segunda1 Then
        mensa = MsgBox("Rango invalido, ingrese nuevamente", vbCritical, "Error")
        MaskEdBox2.SetFocus
        Exit Sub
    End If

    
    
 Respuesta = MsgBox("Esta por anular comprobantes, esta Ud. Seguro ?", vbYesNo, "!! Atencion !!")
If Respuesta = vbYes Then

    If Option1.Value = True Then
        datprimaryrs.RecordSource = "select libroventas.* from libroventas WHERE libroventas.empresa = " & login.empresaact & " and inicioper = '" & login.iper & "' and finper = '" & login.fper & "' and cerrado = 'N' and tipocompr = '" & Combo1.Text & "' Order by numcompr"
        datprimaryrs.Refresh

        datprimaryrs.Recordset.MoveFirst
        Do While Not datprimaryrs.Recordset.EOF
            segunda1val = Val(Right(segunda1, 8))
            segunda2val = Val(Right(segunda2, 8))
            comproval = Val(Right(datprimaryrs.Recordset.Fields("numcompr"), 8))
            primeratxt = Left(segunda1, 4)
            comprotxt = Left(datprimaryrs.Recordset.Fields("numcompr"), 4)
 
        If comprotxt = primeratxt And (comproval < segunda1val Or comproval > segunda2val) Then GoTo finloop
        If comprotxt <> primeratxt Then GoTo finloop
 
Rem ************* borra asiento ****************
        filtroasiento = datprimaryrs.Recordset.Fields("asiento")
        If IsNull(filtroasiento) = True Then GoTo paso2
        datmaestro.RecordSource = "select [Maestro Asientos].* from [Maestro Asientos] where empresa = " & login.empresaact & " and nroasiento = " & filtroasiento & " and perinicial = '" & login.iper & "' order by nroasiento"
        datmaestro.Refresh
        If datmaestro.Recordset.EOF = True Then GoTo paso2
        masterasiento = datmaestro.Recordset.Fields(2)
        datasiento.RecordSource = "select [Detalle Asientos].* from [Detalle Asientos] where empresa = " & login.empresaact & " and idmasterasientos = " & masterasiento & ""
        datasiento.Refresh
        If datasiento.Recordset.EOF = True Then GoTo paso1
        datasiento.Recordset.MoveFirst
paso0:
        If datasiento.Recordset.EOF = True Then GoTo paso1
        datasiento.Recordset.Delete adAffectCurrent
        If datasiento.Recordset.EOF = True Then GoTo paso1
        datasiento.Recordset.MoveNext
        GoTo paso0
paso1:
        datmaestro.Recordset.Delete adAffectCurrent
paso2:
Rem *********** fin borrado de asiento ****************
If datempresa.Recordset.Fields("modfacturacion") = "LIQUIDACIONES" Then GoTo liquidaciones
If datempresa.Recordset.Fields("modfacturacion") = "ORD-RADIO" Or IsNull(datempresa.Recordset.Fields("modfacturacion")) = True Then GoTo radio
GoTo sigue
    
    
radio:
    If IsNull(datprimaryrs.Recordset.Fields("numordenpub")) = True Then GoTo sigue
    If datprimaryrs.Recordset.Fields("numordenpub") = "" Then GoTo sigue

    datordenes.RecordSource = "select ordenes.* from ordenes where nrorden = " & datprimaryrs.Recordset.Fields("numordenpub") & ""
    datordenes.Refresh
    If datordenes.Recordset.EOF = True Then GoTo sigue
    
    datordenes.Recordset.Fields("facturada") = "N"
    datordenes.Recordset.UpdateBatch adAffectCurrent
    
    
    
    dathistoria.RecordSource = "select historialorden.* from historialorden"
    dathistoria.Refresh
   
            dathistoria.Recordset.AddNew
            dathistoria.Recordset.Fields("nrorden") = datprimaryrs.Recordset.Fields("numordenpub")
            dathistoria.Recordset.Fields("accion") = "Anula Factura"
            dathistoria.Recordset.Fields("fecha") = Date
            dathistoria.Recordset.Fields("hora") = Time
            dathistoria.Recordset.Fields("motivo") = "Anula Factura"
            dathistoria.Recordset.Fields("usuario") = login.usuarioactivo
     Rem       dathistoria.Recordset.Fields("empresa") = login.empresaact
            dathistoria.Recordset.UpdateBatch adAffectCurrent
    GoTo sigue
    
liquidaciones:
   datordenes.ConnectionString = "Provider=MSDASQL.1;Password=25072004;Persist Security Info=True;User ID=lucva;Data Source=facturacion;Initial Catalog=facturacionSQL"
   factu = tipocomp + " " + Maskcomprobante.Text
   
   datordenes.RecordSource = "select planilla.* from planilla where factura = '" & factu & "'"
   datordenes.Refresh
   datordenes.Recordset.MoveFirst
   If datordenes.Recordset.EOF = True Then GoTo sigue
   Do While Not datordenes.Recordset.EOF
    
    datordenes.Recordset.Fields("facturada") = "N"
    datordenes.Recordset.Fields("factura") = "-"
    datordenes.Recordset.UpdateBatch adAffectCurrent
    datordenes.Recordset.MoveNext
   Loop
    
    datfacclientes.RecordSource = "select facclientes.* from facclientes where empresa = " & login.empresaact & " and tipocomp = '" & tipocomp & "' and comprobante = '" & Maskcomprobante.Text & "'"
    datfacclientes.Refresh
    If datfacclientes.Recordset.EOF = True Then GoTo sigue
    datfacclientes.Recordset.Delete adAffectCurrent
sigue:

    datprimaryrs.Recordset.Fields("cliente") = "***ANULADA***"
    facanulada = "s"
    datprimaryrs.Recordset.Fields("cuit") = ""
    datprimaryrs.Recordset.Fields("tipoiva") = ""
    datprimaryrs.Recordset.Fields("cerrado") = "N"
    For x = 0 To 16
        datprimaryrs.Recordset.Fields(8 + x) = 0
    Next x

    datprimaryrs.Recordset.UpdateBatch adAffectCurrent
        
        Inicio.datauditoria.ConnectionString = login.conexiontotal
    
        Inicio.datauditoria.RecordSource = "select auditoria.* from auditoria"
        Inicio.datauditoria.Refresh
    
        Inicio.datauditoria.Recordset.AddNew
        Inicio.datauditoria.Recordset.Fields("fecha") = Date
        Inicio.datauditoria.Recordset.Fields("hora") = Str(Time)
        Inicio.datauditoria.Recordset.Fields("ventana") = "Libro Ventas"
        Inicio.datauditoria.Recordset.Fields("accion") = "Anulacion:" + tipocomp.Text + Maskcomprobante.Text + " Prov:" + Left(denominacion.Text, 15)
        Inicio.datauditoria.Recordset.Fields("usuario") = login.usuarioactivo
        Inicio.datauditoria.Recordset.Fields("empresa") = login.empresaact
        Inicio.datauditoria.Recordset.UpdateBatch adAffectCurrent
finloop:
        datprimaryrs.Recordset.MoveNext
        Loop
    Else
        For Y = segunda1 To segunda2
            segundatexto = primera1 + "-" + Mid("00000000", 1, 9 - Len(Str(Y))) + Right(Str(Y), Len(Str(Y)) - 1)
            datprimaryrs.Recordset.AddNew
            datprimaryrs.Recordset.Fields("empresa") = login.empresaact
            datprimaryrs.Recordset.Fields("fecha") = fechaanul.Value
            datprimaryrs.Recordset.Fields("cliente") = "***ANULADA***"
            datprimaryrs.Recordset.Fields("tipocompr") = Combo1.Text
            datprimaryrs.Recordset.Fields("numcompr") = segundatexto
            datprimaryrs.Recordset.Fields("cerrado") = "N"
            datprimaryrs.Recordset.Fields("inicioper") = login.iper
            datprimaryrs.Recordset.Fields("finper") = login.fper
            datprimaryrs.Recordset.UpdateBatch adAffectCurrent
            
            datfacclientes.RecordSource = "select facclientes.* from facclientes where empresa = " & login.empresaact & " and tipocomp = '" & Combo1.Text & "' and comprobante = '" & segundatexto & "'"
            datfacclientes.Refresh
            If datfacclientes.Recordset.EOF = True Then GoTo sale
                datfacclientes.Recordset.MoveFirst
                Do While Not datfacclientes.Recordset.EOF
                datfacclientes.Recordset.Delete adAffectCurrent
                datfacclientes.Recordset.MoveNext
            Loop
sale:
            
        Next Y
      
    
    End If
    MsgBox "Registros Anulado", vbInformation, "Correcto"
    For x = frmlibroventas_nuevo.Width To 9660 Step -200
            frmlibroventas_nuevo.Width = x
    Next x
    
    Call nuevo_Click
End If
Exit Sub
salir:
        MsgBox "No se Puede Anular", vbCritical, "Error"

End Sub

Private Sub anular_Click()
On Error GoTo salir

    Respuesta = MsgBox("Esta por anular un comprobante, esta Ud. Seguro ?", vbYesNo, "!! Atencion !!")
If Respuesta = vbYes Then

Rem ************* borra asiento ****************
    datprimaryrs.RecordSource = "select libroventas.* from libroventas where empresa = " & login.empresaact & " and id = " & Text11.Text & ""
    datprimaryrs.Refresh
    filtroasiento = datprimaryrs.Recordset.Fields("asiento")
    If IsNull(filtroasiento) = True Then GoTo paso2
    datmaestro.RecordSource = "select [Maestro Asientos].* from [Maestro Asientos] where empresa = " & login.empresaact & " and nroasiento = " & filtroasiento & " and perinicial = '" & login.iper & "' order by nroasiento"
    datmaestro.Refresh
    If datmaestro.Recordset.EOF = True Then GoTo paso2
    masterasiento = datmaestro.Recordset.Fields(2)
    datasiento.RecordSource = "select [Detalle Asientos].* from [Detalle Asientos] where empresa = " & login.empresaact & " and idmasterasientos = " & masterasiento & ""
    datasiento.Refresh
    If datasiento.Recordset.EOF = True Then GoTo paso1
    datasiento.Recordset.MoveFirst
paso0:
    If datasiento.Recordset.EOF = True Then GoTo paso1
    datasiento.Recordset.Delete adAffectCurrent
    If datasiento.Recordset.EOF = True Then GoTo paso1
    datasiento.Recordset.MoveNext
    GoTo paso0
paso1:
    datmaestro.Recordset.Delete adAffectCurrent
Rem *********** fin borrado de asiento ****************
paso2:

If datempresa.Recordset.Fields("modfacturacion") = "LIQUIDACIONES" Then GoTo liquidaciones
If datempresa.Recordset.Fields("modfacturacion") = "ORD-RADIO" Or IsNull(datempresa.Recordset.Fields("modfacturacion")) = True Then GoTo radio
GoTo sigue
    
radio:
    If IsNull(datprimaryrs.Recordset.Fields("numordenpub")) = True Then GoTo sigue
    If datprimaryrs.Recordset.Fields("numordenpub") = "" Then GoTo sigue

    datordenes.RecordSource = "select ordenes.* from ordenes where nrorden = " & datprimaryrs.Recordset.Fields("numordenpub") & ""
    datordenes.Refresh
    If datordenes.Recordset.EOF = True Then GoTo sigue
    
    datordenes.Recordset.Fields("facturada") = "N"
    datordenes.Recordset.UpdateBatch adAffectCurrent
    
    
    
    dathistoria.RecordSource = "select historialorden.* from historialorden"
    dathistoria.Refresh
   
            dathistoria.Recordset.AddNew
            dathistoria.Recordset.Fields("nrorden") = datprimaryrs.Recordset.Fields("numordenpub")
            dathistoria.Recordset.Fields("accion") = "Anula Factura"
            dathistoria.Recordset.Fields("fecha") = Date
            dathistoria.Recordset.Fields("hora") = Time
            dathistoria.Recordset.Fields("motivo") = "Anula Factura"
            dathistoria.Recordset.Fields("usuario") = login.usuarioactivo
     Rem       dathistoria.Recordset.Fields("empresa") = login.empresaact
            dathistoria.Recordset.UpdateBatch adAffectCurrent
    GoTo sigue
    
liquidaciones:
   datordenes.ConnectionString = "Provider=MSDASQL.1;Password=25072004;Persist Security Info=True;User ID=lucva;Data Source=facturacion;Initial Catalog=facturacionSQL"
   factu = tipocomp + " " + Maskcomprobante.Text
   
   datordenes.RecordSource = "select planilla.* from planilla where factura = '" & factu & "'"
   datordenes.Refresh
   If datordenes.Recordset.EOF = True Then GoTo sigue
   datordenes.Recordset.MoveFirst
   
   Do While Not datordenes.Recordset.EOF
    
    datordenes.Recordset.Fields("facturada") = "N"
    datordenes.Recordset.Fields("factura") = "-"
    datordenes.Recordset.UpdateBatch adAffectCurrent
    datordenes.Recordset.MoveNext
   Loop
    
    datfacclientes.RecordSource = "select facclientes.* from facclientes where empresa = " & login.empresaact & " and tipocomp = '" & tipocomp & "' and comprobante = '" & Maskcomprobante.Text & "'"
    datfacclientes.Refresh
    If datfacclientes.Recordset.EOF = True Then GoTo sigue
    datfacclientes.Recordset.Delete adAffectCurrent
sigue:
    
    datprimaryrs.Recordset.Fields("cliente") = "***ANULADA***"
    facanulada = "s"
    datprimaryrs.Recordset.Fields("cuit") = ""
    datprimaryrs.Recordset.Fields("tipoiva") = ""
    datprimaryrs.Recordset.Fields("cerrado") = "N"
    For x = 0 To 16
        datprimaryrs.Recordset.Fields(8 + x) = 0
    Next x
    datprimaryrs.Recordset.UpdateBatch adAffectCurrent
    
        
    Inicio.datauditoria.ConnectionString = login.conexiontotal
    
    Inicio.datauditoria.RecordSource = "select auditoria.* from auditoria"
    Inicio.datauditoria.Refresh
    
    Inicio.datauditoria.Recordset.AddNew
    Inicio.datauditoria.Recordset.Fields("fecha") = Date
    Inicio.datauditoria.Recordset.Fields("hora") = Str(Time)
    Inicio.datauditoria.Recordset.Fields("ventana") = "Libro Ventas"
    Inicio.datauditoria.Recordset.Fields("accion") = "Anulacion:" + tipocomp.Text + Maskcomprobante.Text + " Prov:" + Left(denominacion.Text, 15)
    Inicio.datauditoria.Recordset.Fields("usuario") = login.usuarioactivo
    Inicio.datauditoria.Recordset.Fields("empresa") = login.empresaact
    Inicio.datauditoria.Recordset.UpdateBatch adAffectCurrent

    datfacclientes.RecordSource = "select facclientes.* from facclientes where empresa = " & login.empresaact & " and tipocomp = '" & tipocomp & "' and comprobante = '" & Maskcomprobante.Text & "'"
    datfacclientes.Refresh
    If datfacclientes.Recordset.EOF = True Then GoTo sale
    datfacclientes.Recordset.MoveFirst
    Do While Not datfacclientes.Recordset.EOF
        datfacclientes.Recordset.Delete adAffectCurrent
        datfacclientes.Recordset.MoveNext
    Loop
sale:
    MsgBox "Registro Anulado", vbInformation, "Correcto"
    Call nuevo_Click
End If
Exit Sub
salir:
        MsgBox "No se Puede Anular", vbCritical, "Error"
        
End Sub

Private Sub anulmasiva_Click()

    If frmlibroventas_nuevo.Width < 10000 Then
        For x = frmlibroventas_nuevo.Width To 13185 Step 200
            frmlibroventas_nuevo.Width = x
        Next x
    Else
        frmlibroventas_nuevo.Width = 9660
    End If

    
Option1.Value = True
    
Combo1.ListIndex = 0

Call filtro_Click

If Frame4.Visible = True Then
    Frame4.Visible = False
    datprimaryrs.RecordSource = "select libroventas.* from libroventas WHERE libroventas.empresa = " & login.empresaact & " and inicioper = '" & login.iper & "' and finper = '" & login.fper & "' and cerrado = 'N' Order by id"
    datprimaryrs.Refresh
    datprimaryrs.Recordset.MoveLast
Else
    Frame4.Visible = True
End If
   

End Sub

Private Sub borrar_Click()
On Error GoTo fuera
  KeyAscii = 13
  Respuesta = MsgBox("ESTA POR BORRAR UN REGISTRO, ESTA SEGURO?", vbYesNo, "Atención")
If Respuesta = vbYes Then

    Inicio.datauditoria.ConnectionString = login.conexiontotal
    
    Inicio.datauditoria.RecordSource = "select auditoria.* from auditoria"
    Inicio.datauditoria.Refresh
    
    Inicio.datauditoria.Recordset.AddNew
    Inicio.datauditoria.Recordset.Fields("fecha") = Date
    Inicio.datauditoria.Recordset.Fields("hora") = Str(Time)
    Inicio.datauditoria.Recordset.Fields("ventana") = "Libro Ventas"
    Inicio.datauditoria.Recordset.Fields("accion") = "Eliminacion:" + tipocomp.Text + Maskcomprobante.Text + " Prov:" + Left(denominacion.Text, 15)
    Inicio.datauditoria.Recordset.Fields("usuario") = login.usuarioactivo
    Inicio.datauditoria.Recordset.Fields("empresa") = login.empresaact
    Inicio.datauditoria.Recordset.UpdateBatch adAffectCurrent

Rem ************* borra asiento ****************
    datprimaryrs.RecordSource = "select libroventas.* from libroventas where empresa = " & login.empresaact & " and id = " & Text11.Text & ""
    datprimaryrs.Refresh
    filtroasiento = datprimaryrs.Recordset.Fields("asiento")
    If IsNull(filtroasiento) = True Then GoTo paso2
    datmaestro.RecordSource = "select [Maestro Asientos].* from [Maestro Asientos] where empresa = " & login.empresaact & " and nroasiento = " & filtroasiento & " and perinicial = '" & login.iper & "'"
    datmaestro.Refresh
    If datmaestro.Recordset.EOF = False Then datmaestro.Recordset.Delete adAffectCurrent
paso2:
    datfacclientes.RecordSource = "select facclientes.* from facclientes where empresa = " & login.empresaact & " and tipocomp = '" & tipocomp & "' and comprobante = '" & Maskcomprobante.Text & "'"
    datfacclientes.Refresh
    If datfacclientes.Recordset.EOF = True Then GoTo sale
    datfacclientes.Recordset.MoveFirst
    Do While Not datfacclientes.Recordset.EOF
        datfacclientes.Recordset.Delete adAffectCurrent
        datfacclientes.Recordset.MoveNext
    Loop
sale:
    datprimaryrs.Recordset.Delete
    errorasiento = False
    MsgBox "Registro Eliminado", vbInformation, "Correcto"
    Call nuevo_Click
    Exit Sub
  
Else
    Exit Sub
End If


fuera:
MsgBox "NO se pudo Eliminar el registro", vbCritical, "Error"

End Sub


Private Sub calcular_Click()
On Error GoTo erroRcalcular

sumar = 0
parcial = 0
result = 0
For t = 0 To 14
    Text3(t).Text = Format(Text3(t).Text, "0.00")
        For x = 1 To 15
          Text3(x).Text = Format(Text3(x).Text, "0.00")
            If sig(t, x) = "" And x = 1 Then GoTo paso1
            If sumar = 1 Then parcial = result
            If sig(t, x) = "*" Then result = Val(Text3(Val(ter(t, x)) - 1).Text) * Val(ter(t, x + 1))
            If sig(t, x) = "/" Then result = Val(Text3(Val(ter(t, x)) - 1).Text) / Val(ter(t, x + 1))
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
 Rem       Text3(t).Text = Format(Text3(t).Text, "#,##0.00")
         Next x
       
paso1:
Next t
        
Exit Sub

erroRcalcular:
    mensa = MsgBox("Error al realizar calculo automatico, revisar configuracion de columnas del libro", vbCritical, "!! Atención !!")


End Sub

Private Sub calcular2_Click()
On Error GoTo erroRcalcular

  
  alic = datparamventas.Recordset.Fields(List2.ListIndex + 4)

  
  Text3(List2.ListIndex).Text = Text3(15).Text / (1 + (alic / 100))
    
  SendKeys "{ENTER}", True

Exit Sub
erroRcalcular:
    mensa = MsgBox("Error al realizar calculo automatico, revisar configuracion de columnas del libro", vbCritical, "!! Atención !!")

End Sub


Private Sub Cancelar_Click()

    Call nuevo_Click

End Sub

Private Sub confcolumnas_Click()
    
    Unload Me
    frmcolumnascompra.Show


End Sub

Private Sub Combo1_Click()

Call filtro_Click


End Sub

Private Sub Command4_GotFocus(Index As Integer)

    If Index = 0 Then Maskfecha.SetFocus

End Sub

Private Sub Command5_GotFocus(Index As Integer)

    If Index = 0 Then
        flagbuscar = 1
        denominacion.SetFocus
    End If

End Sub

Private Sub Check1_Click(Index As Integer)

If Index = 0 Then
    If Check1(0).Value = 0 Then Check1(1).Value = 1
    If Check1(0).Value = 1 Then Check1(1).Value = 0
End If

If Index = 1 Then
    If Check1(1).Value = 0 Then Check1(0).Value = 1
    If Check1(1).Value = 1 Then Check1(0).Value = 0
End If

End Sub

Private Sub Check1_KeyPress(Index As Integer, KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        KeyAscii = 9
    End If
    
End Sub

Private Sub Command3_Click()
On Error Resume Next
Dim crApp As New CRAXDRT.Application
Dim crReport As New CRAXDRT.Report

Dim tabla As String
Dim ruta As String

ruta = "\Empresa" + Right(Str(login.empresaact), Len(Str(login.empresaact)) - 1)

reporte.SQL = "SELECT facturas.fecha, facturas.cliente, facturas.descripcion, facturas.cuit, facturas.tipocompr, facturas.numcompr, facturas.total, facturas.avisador, facturas.producto, facturas.telefono, facturas.contado, facturas.cant, facturas.unidadmedida, facturas.detalle, facturas.preciounit, facturas.totales, facturas.descuento, facturas.totalneto, facturas.impdesc, facturas.domicilio, facturas.localidad, facturas.numdisco, facturas.empresa FROM contablesql.dbo.facturas facturas WHERE facturas.tipocompr = '" & tipocomp.Text & "' and facturas.empresa = " & login.empresaact & " and facturas.numcompr = '" & Maskcomprobante.Text & "' "
tabla = reporte.SQL

With CrystalReporte
    .PrinterCollation = crptCollated
If tipocomp.Text = "F-A" Or tipocomp.Text = "NCA" Or tipocomp.Text = "NDA" Then
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
    .Destination = crptToWindow
    .PrintFileType = crptCrystal
    .WindowState = crptMaximized
    .Action = 1

End With



End Sub

Private Sub compfecha_Click()

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
            fechamal = 2
    End If

End Sub

Private Sub DataCombo1_KeyPress(KeyAscii As Integer)
On Error GoTo fuera

    If KeyAscii = 13 Then
        KeyAscii = 0
        If IsNull(DataCombo1.Text) = True Then Exit Sub

        login.datusyempresas.RecordSource = "select usuarioyempresa.* from usuarioyempresa WHERE usuarioyempresa.nomusuario = '" & login.usuarioactivo & "' and empresa = " & DataCombo1.BoundText & ""
        login.datusyempresas.Refresh
        If login.datusyempresas.Recordset.EOF = True Then
            mensa = MsgBox("Permiso denegado a esta empresa", vbCritical, "Error")
            DataCombo1.Text = login.nomempresa
            Exit Sub
        End If
        login.empresaact = DataCombo1.BoundText
        login.nomempresa = DataCombo1.Text

        datempresa.RecordSource = "select empresa.* from empresa where empresa = " & login.empresaact & " "
        datempresa.Refresh
        login.iper = datempresa.Recordset.Fields("inicioperiodo")
        login.fper = datempresa.Recordset.Fields("finperiodo")
    
        Unload Me
        frmlibroventas_nuevo.Show
    End If
fuera:
End Sub

Private Sub DataCombo1_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next
        login.empresaact = DataCombo1.BoundText
        login.nomempresa = DataCombo1.Text

End Sub

Private Sub DataCombo3_KeyPress(KeyAscii As Integer)


    If KeyAscii = 13 Then
        KeyAscii = 0
        anulamasiva.SetFocus
    End If


End Sub

Private Sub DataList1_KeyPress(KeyAscii As Integer)
On Error GoTo fuera

    If KeyAscii = 45 Then
        Maskfecha.SetFocus
    End If

    If KeyAscii = 13 Then
       If flagbuscar = 1 Then
            flagbuscar = 0
            datproveedores.Recordset.Bookmark = DataList1.SelectedItem
            denominacion.Text = datproveedores.Recordset.Fields(3)
            Call imprimir_Click
            Exit Sub
        End If
    
        If DataList1.SelectedItem <> "" Then
            datproveedores.Recordset.Bookmark = DataList1.SelectedItem
                                                         
            Check1(0).Value = 1
            denominacion.Text = datproveedores.Recordset.Fields(3)
            If datproveedores.Recordset.Fields(4) <> "" Then tipoiva.Text = datproveedores.Recordset.Fields(4)
            If datproveedores.Recordset.Fields(5) <> "" Then cuit.Text = datproveedores.Recordset.Fields(5)
            
            If IsNull(datproveedores.Recordset.Fields("codcontableventas")) = True Then
                codgastos = 0
            Else
                codgastos = datproveedores.Recordset.Fields("codcontableventas")
            End If
       
            pruebanulo = IsNull(datproveedores.Recordset.Fields(12))
            If pruebanulo = True Then
                Text7(30).Text = 0
            Else
                Text7(30).Text = datproveedores.Recordset.Fields(12)
            End If
            List1.Visible = True
            
                        
            List1.SetFocus
        Else
            Exit Sub
        End If
    End If
    
fuera:
End Sub


Private Sub DataList1_LostFocus()

  DataList1.Visible = False
  
End Sub
Private Sub DataList2_Click()
On Error GoTo fuera

    Text7(poscuenta).Text = DataList2.BoundText
fuera:
End Sub

Private Sub DataList2_GotFocus()
On Error GoTo fuera

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

    
fuera:
End Sub

Private Sub DataList2_KeyPress(KeyAscii As Integer)
On Error GoTo fuera

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
              digito = Val(datccostos.Recordset.Fields(4))
              digcue = Val(Mid(Text7(poscuenta).Text, 1, 1))
              If digcue = digito And login.habcc = True Then
                DataList3.Visible = True
                Text9.Visible = True
                Frame3.Visible = True
                DataList3.SetFocus
                Exit Sub
              End If
sigue:
              If poscuenta < 30 Then sumadebe = Text3(posicion) + sumadebe
              If Text3(posicion + 1).Visible = True Then
                    Text3(posicion + 1).SetFocus
              Else
                    grabalibroasiento.SetFocus
              End If
              If poscuenta = 30 Then
                    grabalibroasiento.SetFocus
              End If

              Exit Sub
    End If

fuera:
End Sub

Private Sub DataList2_LostFocus()

    DataList2.Visible = False
    
End Sub

Private Sub DataList3_Click()
    Text9.Text = DataList3.BoundText
End Sub

Private Sub datalist3_GotFocus()

    If Text9.Text <> "" Then
        DataList3.BoundText = Text9.Text
    End If

End Sub

Private Sub DataList3_KeyPress(KeyAscii As Integer)
On Error GoTo fuera

    If KeyAscii = 13 Then
        KeyAscii = 0
            Text9.Text = DataList3.BoundText
              If poscuenta < 31 Then sumadebe = Text3(posicion) + sumadebe
              If Text3(posicion + 1).Visible = True Then
                    Text3(posicion + 1).SetFocus
              Else
                    datprimaryrs.Recordset.UpdateBatch adAffectCurrent
                    pos = datprimaryrs.Recordset.AbsolutePosition
                    datprimaryrs.Refresh
                    datprimaryrs.Recordset.AbsolutePosition = pos
                    Text7(30).SetFocus
              End If
              If poscuenta = 30 Then
                    sumahaber = Text3(15) + sumahaber
                    datprimaryrs.Recordset.UpdateBatch adAffectCurrent
                    pos = datprimaryrs.Recordset.AbsolutePosition
                    datprimaryrs.Refresh
                    datprimaryrs.Recordset.AbsolutePosition = pos
                    Call grabalibroasiento_Click
              End If
    End If
    
fuera:
End Sub

Private Sub DataList3_LostFocus()
On Error GoTo fuera

    If Text9.Text = "" Then
        mensa = MsgBox("Debe ingresa un Centro de Costo", vbCritical, "!Error¡")
        DataList3.SetFocus
        Exit Sub
    End If
Frame3.Visible = False
Text9.Visible = False
DataList3.Visible = False

fuera:
End Sub

Private Sub denominacion_GotFocus()

    DataList1.Visible = True
    DataList1.SetFocus

End Sub

Private Sub fechaanul_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        anulamasiva.SetFocus
    End If

End Sub

Private Sub filtro_Click()

  datprimaryrs.RecordSource = "select libroventas.* from libroventas WHERE libroventas.empresa = " & login.empresaact & " and inicioper = '" & login.iper & "' and finper = '" & login.fper & "' and cerrado = 'N' and tipocompr = '" & Combo1.Text & "' Order by numcompr"
  datprimaryrs.Refresh
  If datprimaryrs.Recordset.EOF = True Then
    DataCombo2.Text = ""
    DataCombo3.Text = ""
    Exit Sub
  End If

  If Option1.Value = True Then
    DataCombo2.Enabled = True
    DataCombo3.Enabled = True
    Text11.Enabled = False
    MaskEdBox2.Enabled = False
    fechaanul.Enabled = False
  Else
    DataCombo2.Enabled = False
    DataCombo3.Enabled = False
    Text11.Enabled = False
    MaskEdBox2.Enabled = True
    fechaanul.Enabled = True
    fechaanul.Value = Date
    datprimaryrs.Recordset.MoveLast
    numero = datprimaryrs.Recordset.Fields("numcompr")
    ocho = Right(numero, 8)
    ochonum = Val(ocho) + 1
    ochofinal = Mid("00000000", 1, 9 - Len(Str(ochonum))) + Right(Str(ochonum), Len(Str(ochonum)) - 1)
    Text15.Text = Left(numero, 4) + "-" + ochofinal
    MaskEdBox2.SetFocus
  End If


  
  DataCombo2.Text = ""
  DataCombo3.Text = ""
  
End Sub

Private Sub Form_Activate()
If login.livaventasmodi = "N" Then
    grabalibroasiento.Enabled = False
    borrar.Enabled = False
    anular.Enabled = False
    anulmasiva.Enabled = False
Else
    borrar.Enabled = True
    grabalibroasiento.Enabled = True
    anular.Enabled = True
    anulmasiva.Enabled = True
End If
End Sub

Private Sub Form_GotFocus()
    
    Maskcomprobante.Mask = ""
    Maskfecha.Mask = ""
    
End Sub

Private Sub Form_Load()
On Error Resume Next
Aplicar_skin Me

frmlibroventas_nuevo.Top = 0
frmlibroventas_nuevo.Left = 0


datasiento.ConnectionString = login.conexiontotal
datbusca.ConnectionString = login.conexiontotal
datccostos.ConnectionString = login.conexiontotal
datcolumnas.ConnectionString = login.conexiontotal
datcuentas.ConnectionString = login.conexiontotal
datlistacostos.ConnectionString = login.conexiontotal
datmaestro.ConnectionString = login.conexiontotal
datprimaryrs.ConnectionString = login.conexiontotal
datproveedores.ConnectionString = login.conexiontotal
datempresa.ConnectionString = login.conexiontotal
datempresa1.ConnectionString = login.conexiontotal
datparamgral.ConnectionString = login.conexiontotal
datparamventas.ConnectionString = login.conexiontotal
datfacclientes.ConnectionString = login.conexiontotal

If login.empresaact1 > 0 Then
    empresareal = login.empresaact1
Else
    empresareal = login.empresaact
End If

DataCombo1.Text = login.nomempresa
previsualiza = 1
modificar = 0

 Text5(0).Text = login.iper
 Text5(1).Text = login.fper
 flagbuscar = 0
 
 Inicio.Caption = login.nomempresa + "-Periodo Contable: " + Str(login.iper) + " -" + Str(login.fper)
 
    datparamventas.RecordSource = "select paramventas.* from paramventas where empresa = " & login.empresaact & ""
    datparamventas.Refresh
 
  datempresa1.RecordSource = "select empresa.* from empresa"
  datempresa1.Refresh

  datempresa.RecordSource = "select empresa.* from empresa where empresa = " & login.empresaact & " "
  datempresa.Refresh

  datparamgral.RecordSource = "select parametrosgenerales.* from parametrosgenerales"
  datparamgral.Refresh


  frmlibroventas_nuevo.Left = 0
  frmlibroventas_nuevo.Top = 0
  frmlibroventas_nuevo.Width = 9660
    
  Inicio.Toolbar1.Visible = True
  datprimaryrs.RecordSource = "select libroventas.* from libroventas WHERE inicioper = '" & login.iper & "' and libroventas.empresa = " & login.empresaact & " and cerrado <> 'N' Order by cerrado"
  datprimaryrs.Refresh
  
  Check1(0).Value = 1
  If datprimaryrs.Recordset.EOF = True Then
        fechafuera = login.iper
  Else
        datprimaryrs.Recordset.MoveLast
        mesfuera = datprimaryrs.Recordset.Fields(25) + 1
        años = Year(Date)
        If mesfuera = "12" And Month(Date) = 1 Then años = Year(Date) - 1
        If mesfuera = "13" Then mesfuera = "01"
        If Len(mesfuera) = 1 Then
            mesfuera1 = "0" + Right(Str(mesfuera), Len(Str(mesfuera) - 1))
        Else
            mesfuera1 = Right(Str(mesfuera), 2)
        End If
        añofuera = Right(Str(años), 4)
        fechafuera = "01/" + mesfuera1 + "/" + añofuera
   End If

  datprimaryrs.RecordSource = "select libroventas.* from libroventas WHERE empresa = '0'"
  datprimaryrs.Refresh
  
  datcolumnas.RecordSource = "SELECT columnasventa.* FROM columnasventa WHERE inicioper = '" & login.iper & "' and empresa = " & login.empresaact & " "
  datcolumnas.Refresh
  
  datlistacostos.RecordSource = "SELECT listaccostos.* FROM listaccostos WHERE empresa = " & login.empresaact & " order by cc"
  datlistacostos.Refresh
  
  datccostos.RecordSource = "SELECT ccostos.* FROM ccostos WHERE empresa = " & login.empresaact & ""
  datccostos.Refresh
  
  datcuentas.RecordSource = "select listacuentas.* from listacuentas WHERE inicioper = '" & login.iper & "' and empre = " & login.empresaact & " ORDER BY IDCUENTA"
  datcuentas.Refresh
    
  datproveedores.RecordSource = "select clientes.* from clientes Where empresa = " & empresareal & " ORDER BY razonsocial"
  datproveedores.Refresh
  
  
automatico.Value = 1
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

     Maskfecha.Mask = "##/##/####"
     Maskfecha.MaxLength = 10
    If tipocomp.Text <> " " Then
     Maskcomprobante.Mask = "####-########"
     Maskcomprobante.MaxLength = 13
    End If
    Text1.Text = login.empresaact
        
        
If datempresa.Recordset.Fields("condtrib") = "RI" Then
        List1.AddItem "F-A"
        List1.AddItem "F-B"
        List1.AddItem "F-M"
        List1.AddItem "R-A"
        List1.AddItem "R-B"
        List1.AddItem "NDA"
        List1.AddItem "NDB"
        List1.AddItem "NCA"
        List1.AddItem "NCB"
        List1.AddItem "TFA"
        List1.AddItem "TFB"
        List1.AddItem "TFZ"
        List1.AddItem " "
        
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
        Combo1.AddItem "TFZ"
Else
        List1.AddItem "F-C"
        List1.AddItem "R-C"
        List1.AddItem "NDC"
        List1.AddItem "NCC"
        List1.AddItem "TFC"
        List1.AddItem " "
        
        Combo1.AddItem "F-C"
        Combo1.AddItem "R-C"
        Combo1.AddItem "NDC"
        Combo1.AddItem "NCC"
        Combo1.AddItem "TFC"
        
End If
   

    
    
    For x = 0 To 14
        valida = IsNull(datcolumnas.Recordset.Fields(x * 2 + 1))
        If datcolumnas.Recordset.Fields(x * 2 + 1) = "" Then valida = True
        If valida = False Then
            label1(x).Caption = datcolumnas.Recordset.Fields(x * 2 + 1)
            List2.AddItem datcolumnas.Recordset.Fields(x * 2 + 1)
        Else
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
  grdDataGrid.Height = Me.ScaleHeight - datprimaryrs.Height - 30
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
On Error GoTo fuera

Rem ****************** grabar libro *****************

    Call compfecha_Click
    If fechamal = 1 Then Exit Sub

    If modificar = 0 Then
        datprimaryrs.Recordset.AddNew
    Else
        datprimaryrs.RecordSource = "select libroventas.* from libroventas where empresa = " & login.empresaact & " and id = " & Text11.Text & ""
        datprimaryrs.Refresh
        If datprimaryrs.Recordset.EOF = True Then GoTo fuera
    End If
    
    datprimaryrs.Recordset.Fields("empresa") = login.empresaact
    datprimaryrs.Recordset.Fields("fecha") = Maskfecha.Text
    datprimaryrs.Recordset.Fields("tipoiva") = tipoiva.Text
    datprimaryrs.Recordset.Fields("cuit") = cuit.Text
    datprimaryrs.Recordset.Fields("tipocompr") = tipocomp.Text
    datprimaryrs.Recordset.Fields("cliente") = denominacion.Text
    datprimaryrs.Recordset.Fields("numcompr") = Maskcomprobante.Text
For x = 0 To 14
    If IsNull(Text3(x).Text) = True Then Text3(x).Text = 0
    If Text3(x).Text = "" Then Text3(x).Text = 0
    If Text3(x).Visible = True Then datprimaryrs.Recordset.Fields(8 + x) = Text3(x).Text
Next x
    
    If Left(tipocomp.Text, 2) = "NC" Then
        Text3(15).Text = Text3(15).Text * -1
        For x = 0 To 14
            If Text3(x).Visible = True And Text3(x).Text <> 0 Then
                Text3(x).Text = Text3(x).Text * -1
                datprimaryrs.Recordset.Fields(8 + x) = Text3(x).Text
            End If
        Next x
    End If

    datprimaryrs.Recordset.Fields("total") = Text3(15).Text
    If modificar = 0 Then datprimaryrs.Recordset.Fields("cerrado") = "N"
    
Y = 27
For x = 1 To 29 Step 2
    If Text7(x).Text = " " Or Text7(x).Text = "" Then Text7(x).Text = 0
    datprimaryrs.Recordset.Fields(Y) = Text7(x).Text
    Y = Y + 2
Next x
    datprimaryrs.Recordset.Fields("cdt") = Text7(30).Text
    datprimaryrs.Recordset.Fields("asentado") = "S"
    If IsNull(datprimaryrs.Recordset.Fields("ccosto")) = False Then
        datprimaryrs.Recordset.Fields("ccosto") = Text9.Text
    End If
If modificar = 0 Then
    datprimaryrs.Recordset.Fields("inicioper") = login.iper
    datprimaryrs.Recordset.Fields("finper") = login.fper
End If
    If Check1(1).Value = 1 Then datprimaryrs.Recordset.Fields("contado") = "S"
    If Text7(30).Text = datcolumnas.Recordset.Fields("fc") Then datprimaryrs.Recordset.Fields("contado") = "S"
        
Rem ****************** grabar asiento


    If modificar = 1 Then
        datmaestro.RecordSource = "select [Maestro Asientos].* from [Maestro Asientos] where empresa = " & login.empresaact & " and perinicial = '" & datprimaryrs.Recordset.Fields("inicioper") & "' and nroasiento = " & datprimaryrs.Recordset.Fields("asiento") & " "
        datmaestro.Refresh
        If datmaestro.Recordset.EOF = False Then
                fechaasiento = datmaestro.Recordset.Fields(0)
                datmaestro.Recordset.Delete adAffectCurrent
        End If
    End If

If modificar = 0 Then
    datmaestro.RecordSource = "select [Maestro Asientos].* from [Maestro Asientos] where empresa = " & login.empresaact & " and perinicial = '" & login.iper & "' order by nroasiento"
    datmaestro.Refresh
Else
    datmaestro.RecordSource = "select [Maestro Asientos].* from [Maestro Asientos] where empresa = " & login.empresaact & " and perinicial = '" & datprimaryrs.Recordset.Fields("inicioper") & "' order by nroasiento"
    datmaestro.Refresh
End If
    If datmaestro.Recordset.EOF = False Then
            datmaestro.Recordset.MoveLast
            nroasie = datmaestro.Recordset.Fields(3) + 1
    Else
            nroasie = 1
    End If

    
pas1:
    datmaestro.Recordset.AddNew
    If modificar = 0 Then
        If fechamal = 0 Then
            datmaestro.Recordset.Fields(0) = Maskfecha.Text
        Else
            datmaestro.Recordset.Fields(0) = fechafuera
        End If
    Else
        datmaestro.Recordset.Fields(0) = fechaasiento
    End If
    datmaestro.Recordset.Fields(1) = Date
    datmaestro.Recordset.Fields(3) = nroasie
    datmaestro.Recordset.Fields(4) = Left(denominacion.Text, 20) + " " + tipocomp.Text + " Nº:" + Maskcomprobante.Text
    If modificar = 0 Then
        datmaestro.Recordset.Fields(5) = Text5(0).Text
        datmaestro.Recordset.Fields(6) = Text5(1).Text
    Else
        datmaestro.Recordset.Fields(5) = datprimaryrs.Recordset.Fields("inicioper")
        datmaestro.Recordset.Fields(6) = datprimaryrs.Recordset.Fields("finper")
    End If
    datmaestro.Recordset.Fields(7) = login.empresaact
    datmaestro.Recordset.Fields(8) = "N"
    datmaestro.Recordset.Fields(9) = Val(datprimaryrs.Recordset.Fields(0))
    datmaestro.Recordset.Fields(10) = "V"
    datmaestro.Recordset.Fields(11) = "S"
    datmaestro.Recordset.UpdateBatch adAffectCurrent
      


    datasiento.RecordSource = "select [Detalle Asientos].* from [Detalle Asientos] where empresa = '0'"
    datasiento.Refresh
          
For x = 1 To 29 Step 2
    If Text7(x).Text <> 0 And Text3((x - 1) / 2) <> 0 Then
            If Text3((x - 1) / 2).Visible = False Then GoTo paso1
            
            grilla.Row = (x - 1) / 2
            grilla.Col = 0
            If grilla.Text <> "" Then
                For Y = 0 To 3
                    grilla.Col = Y * 2
                    If grilla.Text = "" Then grilla.Text = 0
                    If grilla.Text = 0 Then GoTo continua
                    
                    datasiento.Recordset.AddNew
                    
                    datasiento.Recordset.Fields(4) = grilla.Text
                    grilla.Col = Y * 2 + 1
                    datasiento.Recordset.Fields(2) = grilla.Text
                    datasiento.Recordset.Fields(5) = Val(datmaestro.Recordset.Fields(2))
                    datasiento.Recordset.Fields(1) = datmaestro.Recordset.Fields(0).Value
                    datasiento.Recordset.Fields(7) = login.empresaact
                    datasiento.Recordset.Fields(6) = label1((x - 1) / 2).Caption
                    datasiento.Recordset.UpdateBatch adAffectCurrent
                Next Y
            End If
                                              
            datasiento.Recordset.AddNew
            
            datasiento.Recordset.Fields(5) = Val(datmaestro.Recordset.Fields(2))
            datasiento.Recordset.Fields(1) = datmaestro.Recordset.Fields(0).Value
            datasiento.Recordset.Fields(7) = login.empresaact
            datasiento.Recordset.Fields(2) = Text7(x).Text
            datasiento.Recordset.Fields(4) = Text3((x - 1) / 2).Text
            If datasiento.Recordset.Fields(4) < 0 Then
                datasiento.Recordset.Fields(3) = datasiento.Recordset.Fields(4) * -1
                datasiento.Recordset.Fields(4) = 0
            End If
            datasiento.Recordset.Fields(6) = label1((x - 1) / 2).Caption
            
            If Text9.Text <> "" Then datasiento.Recordset.Fields(8) = Text9.Text
            datasiento.Recordset.UpdateBatch adAffectCurrent
    End If
continua:
Next x
paso1:
            datasiento.Recordset.AddNew
            datasiento.Recordset.Fields(5) = Val(datmaestro.Recordset.Fields(2))
            datasiento.Recordset.Fields(1) = datmaestro.Recordset.Fields(0).Value
            datasiento.Recordset.Fields(7) = login.empresaact
            datasiento.Recordset.Fields(2) = Text7(30).Text
            datasiento.Recordset.Fields(3) = Text3(15).Text
            If datasiento.Recordset.Fields(3) < 0 Then
                datasiento.Recordset.Fields(4) = datasiento.Recordset.Fields(3) * -1
                datasiento.Recordset.Fields(3) = 0
            End If
            datasiento.Recordset.Fields(6) = "Total facturado"
            datasiento.Recordset.UpdateBatch adAffectCurrent

    datprimaryrs.Recordset.Fields("asiento") = nroasie
    datprimaryrs.Recordset.UpdateBatch adAffectCurrent

    
    Inicio.datauditoria.ConnectionString = login.conexiontotal
    
    Inicio.datauditoria.RecordSource = "select auditoria.* from auditoria"
    Inicio.datauditoria.Refresh
    
    Inicio.datauditoria.Recordset.AddNew
    Inicio.datauditoria.Recordset.Fields("fecha") = Date
    Inicio.datauditoria.Recordset.Fields("hora") = Str(Time)
    Inicio.datauditoria.Recordset.Fields("ventana") = "Libro Ventas"
    If modificar = 0 Then
        accion = "Altas:"
    Else
        accion = "Modif:"
    End If
    Inicio.datauditoria.Recordset.Fields("accion") = accion + tipocomp.Text + Maskcomprobante.Text + " Prov:" + Left(denominacion.Text, 15)
    Inicio.datauditoria.Recordset.Fields("usuario") = login.usuarioactivo
    Inicio.datauditoria.Recordset.Fields("empresa") = login.empresaact
    Inicio.datauditoria.Recordset.UpdateBatch adAffectCurrent

    DataList2.Visible = False
      
    Call nuevo_Click
    Exit Sub
    
fuera:
    mensa = MsgBox("El registro no fue grabado correctamente, cargue nuevamente este movimiento", vbCritical, "Error")
      
End Sub

Private Sub imprimir_Click()

        modificar = 1

        frmconsutalibroventas_n.Show
        
End Sub

Private Sub List1_KeyPress(KeyAscii As Integer)
On Error GoTo fuera

    If KeyAscii = 45 Then
        DataList1.Visible = True
        DataList1.SetFocus
    End If
    If KeyAscii = 13 Then
        tipocomp.Text = List1.Text
        Maskcomprobante.SetFocus
    End If
                 
fuera:
End Sub

Private Sub List1_LostFocus()
On Error GoTo fuera

   
    If List1.Text = " " Then
       If login.administrador = "N" Then
            mensa = MsgBox("No tiene permiso de Administrador para realizar esta tarea", vbCritical, "Error")
            List1.SetFocus
            Exit Sub
       End If
       Text14.Visible = True
       Text14.Text = "Saldo Inic."
       Text14.SelLength = Len(Text14.Text)
       Text14.SetFocus
       GoTo fuera
    End If
    
    If List1.Text = "" Then
        List1.SetFocus
        Exit Sub
    End If
    
    

fuera:
    List1.Visible = False
    

End Sub



Private Sub List2_KeyPress(KeyAscii As Integer)
On Error GoTo fuera
    If KeyAscii = 13 Then
        KeyAscii = 0
        Text3(List2.ListIndex).SetFocus
        Call calcular2_Click
    End If

fuera:
End Sub

Private Sub List2_LostFocus()

    List2.Visible = False

End Sub

Private Sub Maskcomprobante_DragOver(Source As Control, x As Single, Y As Single, State As Integer)
    
    If List1.Text = "" Then
        List1.SetFocus
        Exit Sub
    End If

End Sub

Private Sub Maskcomprobante_GotFocus()
On Error GoTo fuera

If Inicio.Check1 = 1 Then
    If Left(tipocomp.Text, 2) = "NC" Or Left(tipocomp.Text, 2) = "ND" Then
        tipoc = "F-" + Right(tipocomp.Text, 1)
    Else
        tipoc = tipocomp.Text
    End If

    datbusca.RecordSource = "SELECT empresa, tipocompr, MAX(numcompr) AS numcompr, LEFT(numcompr, 4) AS sucursal From dbo.libroventas GROUP BY empresa, tipocompr, LEFT(numcompr, 4) HAVING (tipocompr IS NOT NULL) AND (LEFT(numcompr, 4) = '" & login.librofactura & "' ) and empresa =" & login.empresaact & " and tipocompr = '" & tipoc & "'"
    datbusca.Refresh
    If datbusca.Recordset.EOF = False Then
        ncompro = datbusca.Recordset.Fields("numcompr")
    Else
        ncompro = login.librofactura + "-" + "00000000"
    End If
    proximafactura0 = Left(ncompro, 4)
    proximafactura1 = Val(Right((ncompro), 8)) + 1
    proximafactura = Mid("00000000", 1, 9 - Len(Str(proximafactura1))) + Right(Str(proximafactura1), Len(Str(proximafactura1)) - 1)
    Maskcomprobante.Text = proximafactura0 + "-" + proximafactura
Else
    Maskcomprobante.Text = login.librofactura + "-________"
    Maskcomprobante.SelStart = 5
End If

Exit Sub

fuera:
    MsgBox "Error en la configuración automatica de los comprobantes de carga, ingrese manualmente el comprobante", vbCritical, "Error"
    
End Sub

Private Sub Maskcomprobante_KeyPress(KeyAscii As Integer)
On Error Resume Next

    If KeyAscii = 45 Then
            List1.Visible = True
            List1.SetFocus
    End If
    If KeyAscii = 13 Then
        KeyAscii = 0
            Text4.Text = Maskcomprobante.Text
            Text3(0).SetFocus
            If Val(Mid(Text4.Text, 1, 4)) = 0 Then
                mensa = MsgBox("Debe ingresar una sucursal en el Nro de factura", vbCritical, "!! Atención !!")
                Maskcomprobante.SetFocus
                Maskcomprobante.SelStart = 0
                Maskcomprobante.SelLength = 4
                Exit Sub
            End If
            
            car = 0
            car1 = 0
            For x = 6 To 13
                If Mid(Text4.Text, x, 1) = "_" Then
                    car = car + 1
                Else
                    car1 = car1 + 1
                End If
            Next x
            Text4.Text = Mid(Text4.Text, 1, 4) + "-" + Mid("0000000", 1, car) + Mid(Text4.Text, 6, car1)
            Maskcomprobante.Text = Text4.Text
            datbusca.RecordSource = "select libroventas.* from libroventas WHERE libroventas.empresa = " & login.empresaact & " and numcompr = '" & Maskcomprobante.Text & "' and tipocompr = '" & tipocomp.Text & "'  "
            datbusca.Refresh

            If datbusca.Recordset.EOF = False Then
                If List1.Text = " " Then Exit Sub
                mensa = MsgBox("Este comprobante ya fue ingresado anteriormente, revise el nº de comprobante, o tipo de comprobante", vbCritical, "!! Atención !!")
                Maskcomprobante.SetFocus
                Exit Sub
            End If


          For x = 1 To Len(Maskcomprobante.Text)
            c = Mid(Maskcomprobante.Text, x, 1)
            If c = "_" Then
                mensa = MsgBox("Nro de factura incorrecto", vbCritical, "!! Atención !!")
                Maskcomprobante.SetFocus
                Maskcomprobante.SelStart = 5
                Maskcomprobante.SelLength = 8
                Exit Sub
            End If
          Next x

    End If

    
fuera:
End Sub



Private Sub Maskcomprobante_LostFocus()

    If List1.Text = "F-B" Or List1.Text = "TFB" Or List1.Text = "TFZ" Then
            Text3(15).Locked = False
            Text3(15).SetFocus
    Else
            Text3(15).Locked = True
    End If

End Sub

Private Sub MaskEdBox1_KeyPress(KeyAscii As Integer)
Dim fecha1 As Date
Dim fecha2 As Date
fecha1 = Maskfecha.Text



    If KeyAscii = 13 Then
        KeyAscii = 0
        fecha2 = MaskEdBox1.Text
        If fecha1 > fecha2 Then
                    mensa = MsgBox("LA FECHA DEL CAI ESTA VENCIDA", vbExclamation, "!! Atencion !!")
        End If
        Frame4.Visible = False
        MaskEdBox1.Mask = ""
        List1.SetFocus
    End If

End Sub

Private Sub MaskEdBox1_LostFocus()
        Frame4.Visible = False
        MaskEdBox1.Mask = ""
        List1.SetFocus
End Sub



Private Sub MaskEdBox2_KeyPress(KeyAscii As Integer)
On Error GoTo fin
    If KeyAscii = 13 Then
        KeyAscii = 0
            car = 0
            car1 = 0
            numero = MaskEdBox2.Text
            For x = 6 To 13
                If Mid(numero, x, 1) = "_" Then
                    car = car + 1
                Else
                    car1 = car1 + 1
                End If
            Next x
            numero = Mid(numero, 1, 4) + "-" + Mid("00000000", 1, car) + Mid(numero, 6, car1)
            MaskEdBox2.Text = numero
            fechaanul.SetFocus
    End If
fin:
End Sub

Private Sub Maskfecha_KeyPress(KeyAscii As Integer)
On Error GoTo fuera
Dim dia As Integer
Dim mes As Integer
Dim año As Integer

    If KeyAscii = 2 Then Text8.Text = "N"

    If KeyAscii = 13 Then
    
        If Right(Maskfecha.Text, 2) = "__" And Mid(Maskfecha.Text, 7, 2) <> "__" Then
            Maskfecha.Text = Left(Maskfecha.Text, 6) + "20" + Mid(Maskfecha.Text, 7, 2)
        End If
        Text2.Text = Maskfecha.Text
           
        Call compfecha_Click

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
    
fuera:
End Sub

Private Sub nuevo_Click()
On Error Resume Next


    Unload Me
    frmlibroventas_nuevo.Show
    Exit Sub
    
    If errorasiento = True Then Exit Sub

    facturaimprime = datprimaryrs.Recordset.Fields("id")

    For x = 0 To 15
            Text3(x).Text = 0
    Next x
    For x = 0 To 31
            Text7(x).Text = 0
    Next x
  
    denominacion.Text = ""
    tipoiva.Text = ""
    tipocomp.Text = ""
    Text13.Text = ""
    cuit.Text = ""
    Text14.Text = ""
    Text4.Text = ""
    Text9.Text = ""
    Text11.Text = ""
    modificar = 0
  
    Maskfecha.SelLength = 10
    Maskfecha.SelText = ""
    Maskcomprobante.SelLength = 13
    Maskcomprobante.SelText = ""
    Maskcomprobante.Mask = ""
    Maskcomprobante.Text = ""
    
    Maskfecha.Mask = "##/##/####"
    Maskfecha.MaxLength = 10
    Maskcomprobante.Mask = "####-########"
     Maskcomprobante.MaxLength = 13

     grilla.Clear

   
    Maskfecha.SetFocus
    Exit Sub

errornuevo:

End Sub

Private Sub nuevo_GotFocus()

     Maskfecha.Mask = "##/##/####"
     Maskfecha.MaxLength = 10
    If tipocomp.Text <> " " Then
     Maskcomprobante.Mask = "####-########"
     Maskcomprobante.MaxLength = 13
    End If

     grilla.Clear


End Sub

Private Sub Maskfecha_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = 114 Then
        flagbuscar = 1
        denominacion.SetFocus
    End If

End Sub

Private Sub Option1_Click()

    Call filtro_Click
    
End Sub

Private Sub Option2_Click()


    Call filtro_Click

End Sub

Private Sub salir_Click()

    If errorasiento = True Then
        mensa = MsgBox("El asiento está desvalanceado, no se puede grabar", vbCritical, "!! Error !!")
            sumadebe = 0
            sumahaber = 0
            Text3(0).SetFocus
            Exit Sub
    End If
    
    
    
    errorasiento = False
    Call Cancelar_Click
    Unload Me

End Sub


Private Sub Text13_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        Text3(0).SelStart = 0
       Text3(0).SelLength = Len(Text3(0).Text)
       Text3(0).SetFocus
      
    End If
    
End Sub

Private Sub Text14_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
    KeyAscii = 0
    Text3(0).SetFocus
    Text14.Visible = False
    Maskcomprobante.Mask = ""
    Maskcomprobante.Text = Left(Text14.Text, 13)
End If

End Sub

Private Sub Text3_GotFocus(Index As Integer)
On Error GoTo errorlist

    If Index = 0 Then
        If Maskfecha.Text = "__/__/____" Then
            modificar = 0
            Call nuevo_Click
        End If
    End If

    If List1.Text = "" Then
        List1.SetFocus
        Exit Sub
    End If

                Cuenta = 0
                Text3(Index).Text = Format(Text3(Index).Text, "0.00")
                
                Text3(Index).SelStart = 0
                Text3(Index).SelLength = Len(Text3(Index)) + 3
                
                
errorlist:
    Exit Sub
End Sub

Private Sub Text3_KeyPress(Index As Integer, KeyAscii As Integer)
On Error GoTo fuera
Dim suma As Double
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        If Text3(Index).Text = "" Then Text3(Index).Text = 0
        suma = 0

         If automatico = 1 Then Call calcular_Click
                                                            
         If Index = 15 Then
            List2.Visible = True
            List2.SetFocus
            suma = Val(Text3(15).Text)
            Exit Sub
         End If
                                                            
         For x = 0 To 14
            If Text3(x).Visible = True Then suma = Val(Text3(x).Text) + suma
         Next x
         For x = 0 To 14
            If Text3(x).Visible = True Then Text3(x).Text = Format(Text3(x).Text, "#,##0.00")
         Next x
         
         
         Text3(15).Text = suma
         Text3(15).Text = Format(Text3(15).Text, "#,##0.00")
         If Text6.Text = "" Then Text6.Text = "N"
                
         If Text3(Index).Text > 0 Then
            posicion = Index
            Cuenta = 0
            Text7(Index * 2 + 1).SetFocus
            Exit Sub
         End If
                
         If Text3(Index + 1).Visible = True Then
                Text3(Index + 1).SetFocus
         Else
                Text7(30).SetFocus
         End If
     End If
Exit Sub
fuera:
mensa = MsgBox("Algún Campo no fue ingresado, o ingreso un caracter incorrecto", vbCritical, "!! Error ¡¡")
    
End Sub

Private Sub Text3_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    
Rem    If KeyCode = 114 Then
Rem        indice = Index
Rem        librocontado = 0
Rem        filtroasiento = datPrimaryRS.Recordset.Fields("asiento")
Rem        If IsNull(filtroasiento) = True Then GoTo sigue
Rem        datmaestro.RecordSource = "select [Maestro Asientos].* from [Maestro Asientos] where empresa = " & login.empresaact & " and nroasiento = " & filtroasiento & " and perinicial = '" & login.iper & "' order by nroasiento"
Rem        datmaestro.Refresh
Rem        masterasiento = datmaestro.Recordset.Fields(2)
Rem        datasiento.RecordSource = "select [Detalle Asientos].* from [Detalle Asientos] where empresa = " & login.empresaact & " and idmasterasientos = " & masterasiento & " and detallefila = '" & label1(Index).Caption & "' "
Rem        datasiento.Refresh
Rem        If datasiento.Recordset.EOF = True Then GoTo sigue
Rem         datasiento.Recordset.MoveFirst
Rem        grilla.Row = Index
Rem        x = 0
  Rem      Do While Not datasiento.Recordset.EOF
 Rem           grilla.Col = x
 Rem           grilla.Text = Text3(indice)
 Rem           grilla.Col = x + 1
 Rem           grilla.Text = Text7(indice)
 Rem           x = x + 2
Rem            datasiento.Recordset.MoveNext
  Rem      Loop
sigue:
 Rem       frmabredebelc.Show
 Rem       frmabredebelc.label1(0).Caption = label1(Index).Caption
 Rem       frmabredebelc.importes.Value = Text3(Index).Text
 Rem   End If


End Sub

Private Sub Text3_LostFocus(Index As Integer)

    Text3(Index).Text = Format(Text3(Index).Text, "#,##0.00")

End Sub

Private Sub Text7_GotFocus(Index As Integer)
On Error GoTo fuera

    prueba = datcolumnas.Recordset.Fields(Index + 31)
    If prueba > 0 Then
           Text7(Index).Text = prueba
           sumadebe = Text3(posicion) + sumadebe
           Text7(Index + 1).Text = 0
           If Text3(posicion + 1).Visible = True Then
                Text3(posicion + 1).SetFocus
                Exit Sub
           End If
           Text7(30).SetFocus
           Exit Sub
   End If

    If codgastos > 0 And Index < 30 Then Text7(Index).Text = codgastos
    poscuenta = Index
    Text7(Index).SelLength = Len(Text7(Index))
    DataList2.BoundText = Text7(Index).Text
    DataList2.Visible = True
    If Index <= 16 Then
        DataList2.Top = Text7(Index).Top + Text7(Index).Height + Frame1.Height + Frame1.Top
    Else
        DataList2.Top = Text7(Index - 15).Top + Text7(Index - 15).Height + Frame1.Height + Frame1.Top
    End If
    DataList2.SetFocus

fuera:
End Sub

Private Sub textcuenta_GotFocus()
    
    DataList2.Visible = True
    DataList2.SetFocus
    
End Sub

Private Sub Text7_LostFocus(Index As Integer)
On Error GoTo fuera

    poscuenta = Index
    If Index = 31 Then
        If Text7(Index).Text = "" Then Exit Sub
        sumahaber = Text3(15) + sumahaber
  Rem       Call grabalibroasiento_Click
        Exit Sub
    End If
    
fuera:
End Sub

Private Sub verificar_Click()

End Sub
