VERSION 5.00
Object = "{F6125AB1-8AB1-11CE-A77F-08002B2F4E98}#2.0#0"; "MSRDC20.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{D18BBD1F-82BB-4385-BED3-E9D31A3E361E}#1.0#0"; "KewlButtonz.ocx"
Begin VB.Form frmrecibo_ctacte 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Recibo de Cobranza Cta.Cte."
   ClientHeight    =   9465
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14805
   Icon            =   "frmrecibo_ctacte.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9465
   ScaleWidth      =   14805
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   10
      Left            =   11400
      MaxLength       =   50
      TabIndex        =   90
      Top             =   8880
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   9
      Left            =   11400
      MaxLength       =   50
      TabIndex        =   89
      Top             =   8400
      Width           =   1935
   End
   Begin VB.PictureBox Picture9 
      Height          =   375
      Left            =   8760
      ScaleHeight     =   315
      ScaleWidth      =   2475
      TabIndex        =   87
      Top             =   8880
      Width           =   2535
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Credito Maximo:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   16
         Left            =   120
         TabIndex        =   88
         Top             =   0
         Width           =   3615
      End
   End
   Begin VB.PictureBox Picture8 
      Height          =   375
      Left            =   8760
      ScaleHeight     =   315
      ScaleWidth      =   2475
      TabIndex        =   85
      Top             =   8400
      Width           =   2535
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Credito Disponible:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   15
         Left            =   120
         TabIndex        =   86
         Top             =   0
         Width           =   3615
      End
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   8
      Left            =   6360
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   83
      Top             =   2880
      Width           =   1815
   End
   Begin VB.PictureBox Picture7 
      Height          =   375
      Left            =   120
      ScaleHeight     =   315
      ScaleWidth      =   6075
      TabIndex        =   81
      Top             =   2880
      Width           =   6135
      Begin VB.Label Label3 
         Caption         =   "Doble Click Previsualiza Factura"
         Height          =   255
         Left            =   0
         TabIndex        =   84
         Top             =   0
         Width           =   2295
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Saldo de Facturas Actual $"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   13
         Left            =   2760
         TabIndex        =   82
         Top             =   0
         Width           =   3375
      End
   End
   Begin VB.CommandButton imprimerecibo 
      Caption         =   "imprime Recibo"
      Height          =   255
      Left            =   8880
      TabIndex        =   78
      Top             =   8280
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Frame retenciones 
      Caption         =   "Certificados de Retenciones"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   8760
      TabIndex        =   69
      Top             =   2880
      Visible         =   0   'False
      Width           =   5415
      Begin VB.CommandButton Command16 
         Caption         =   "S/N"
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
         Left            =   2880
         TabIndex        =   76
         Top             =   1560
         Width           =   615
      End
      Begin VB.TextBox Text10 
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
         Height          =   375
         Left            =   2040
         TabIndex        =   72
         Top             =   1560
         Width           =   735
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Pend.Recepcion"
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
         Left            =   240
         TabIndex        =   75
         Top             =   1560
         Width           =   1695
      End
      Begin VB.CommandButton Command19 
         Caption         =   "Fec.Emisi�n"
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
         Left            =   240
         TabIndex        =   74
         Top             =   360
         Width           =   1695
      End
      Begin VB.TextBox Text11 
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
         Height          =   375
         Left            =   2040
         TabIndex        =   71
         Top             =   960
         Width           =   2055
      End
      Begin VB.CommandButton Command15 
         Caption         =   "Nro:"
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
         Left            =   240
         TabIndex        =   73
         Top             =   960
         Width           =   1695
      End
      Begin MSComCtl2.DTPicker femisionret 
         Height          =   375
         Left            =   2040
         TabIndex        =   70
         Top             =   360
         Width           =   2055
         _ExtentX        =   3625
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
         Format          =   49676289
         CurrentDate     =   41988
      End
   End
   Begin VB.Frame cancelafactura 
      Caption         =   "Importe a Saldar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   9240
      TabIndex        =   57
      Top             =   2880
      Visible         =   0   'False
      Width           =   4575
      Begin VB.PictureBox Picture5 
         Height          =   375
         Left            =   480
         ScaleHeight     =   315
         ScaleWidth      =   1275
         TabIndex        =   62
         Top             =   1080
         Width           =   1335
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Cancela:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   9
            Left            =   120
            TabIndex        =   63
            Top             =   0
            Width           =   1095
         End
      End
      Begin VB.PictureBox Picture4 
         Height          =   375
         Left            =   480
         ScaleHeight     =   315
         ScaleWidth      =   1275
         TabIndex        =   60
         Top             =   480
         Width           =   1335
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Saldo:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   8
            Left            =   240
            TabIndex        =   61
            Top             =   0
            Width           =   975
         End
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   4
         Left            =   2160
         MaxLength       =   50
         TabIndex        =   59
         Top             =   1080
         Width           =   1935
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   3
         Left            =   2160
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   58
         Top             =   480
         Width           =   1935
      End
   End
   Begin VB.Frame cheques 
      Caption         =   "Cheque de Terceros Diferidos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4335
      Left            =   8760
      TabIndex        =   32
      Top             =   2880
      Visible         =   0   'False
      Width           =   5415
      Begin VB.TextBox Text8 
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
         Height          =   375
         Left            =   1800
         TabIndex        =   37
         Top             =   2760
         Width           =   2055
      End
      Begin VB.CommandButton Command14 
         Caption         =   "Fec.Vencimi."
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
         Left            =   240
         TabIndex        =   44
         Top             =   840
         Width           =   1455
      End
      Begin MSComCtl2.DTPicker femicioncheque 
         Height          =   375
         Left            =   1800
         TabIndex        =   33
         Top             =   360
         Width           =   2055
         _ExtentX        =   3625
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
         Format          =   49676289
         CurrentDate     =   41988
      End
      Begin VB.CommandButton Command13 
         Caption         =   "Nro:"
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
         Left            =   240
         TabIndex        =   43
         Top             =   2280
         Width           =   1455
      End
      Begin VB.CommandButton Command12 
         Caption         =   "Banco"
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
         Left            =   240
         TabIndex        =   42
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox Text9 
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
         Height          =   375
         Left            =   1800
         TabIndex        =   36
         Top             =   2280
         Width           =   2055
      End
      Begin VB.CommandButton Command11 
         Caption         =   "C.U.I.T."
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
         Left            =   240
         TabIndex        =   41
         Top             =   2760
         Width           =   1455
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Razon Social"
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
         Left            =   240
         TabIndex        =   40
         Top             =   3240
         Width           =   1455
      End
      Begin VB.TextBox Text6 
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
         Height          =   375
         Left            =   240
         TabIndex        =   38
         Top             =   3600
         Width           =   3615
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Fec.Emisi�n"
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
         Left            =   240
         TabIndex        =   39
         Top             =   360
         Width           =   1455
      End
      Begin MSDataListLib.DataCombo DataCombo5 
         Bindings        =   "frmrecibo_ctacte.frx":0442
         Height          =   420
         Left            =   240
         TabIndex        =   35
         Top             =   1680
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   741
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         ListField       =   "banco"
         BoundColumn     =   "id"
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
      Begin MSComCtl2.DTPicker fechavencimiento 
         Height          =   375
         Left            =   1800
         TabIndex        =   34
         Top             =   840
         Width           =   2055
         _ExtentX        =   3625
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
         Format          =   49676289
         CurrentDate     =   41988
      End
   End
   Begin VB.PictureBox Picture6 
      Height          =   375
      Left            =   9120
      ScaleHeight     =   315
      ScaleWidth      =   2475
      TabIndex        =   67
      Top             =   2280
      Width           =   2535
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Total a Cancelar $"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   11
         Left            =   120
         TabIndex        =   68
         Top             =   0
         Width           =   3615
      End
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   6
      Left            =   11760
      MaxLength       =   50
      TabIndex        =   66
      Top             =   2280
      Width           =   1935
   End
   Begin VB.PictureBox Picture3 
      Height          =   375
      Left            =   8880
      ScaleHeight     =   315
      ScaleWidth      =   4275
      TabIndex        =   55
      Top             =   120
      Width           =   4335
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Comprobantes Seleccionados"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   0
         TabIndex        =   56
         Top             =   0
         Width           =   3615
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmrecibo_ctacte.frx":0459
      Height          =   2295
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   4048
      _Version        =   393216
      AllowUpdate     =   -1  'True
      AllowArrows     =   -1  'True
      HeadLines       =   1
      RowHeight       =   19
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
   Begin VB.TextBox Text19 
      Height          =   285
      Left            =   6840
      TabIndex        =   54
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox Picture2 
      Height          =   375
      Left            =   120
      ScaleHeight     =   315
      ScaleWidth      =   1275
      TabIndex        =   52
      Top             =   120
      Width           =   1335
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   53
         Top             =   0
         Width           =   975
      End
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
      Height          =   360
      Index           =   0
      Left            =   1560
      MaxLength       =   50
      TabIndex        =   0
      Top             =   120
      Width           =   6495
   End
   Begin VB.Frame tarjeta 
      Caption         =   "Tarjeta de Cr�dito"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5295
      Left            =   8760
      TabIndex        =   16
      Top             =   2880
      Visible         =   0   'False
      Width           =   5415
      Begin MSComCtl2.DTPicker femision 
         Height          =   375
         Left            =   1800
         TabIndex        =   30
         Top             =   4320
         Width           =   2055
         _ExtentX        =   3625
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
         Format          =   49676289
         CurrentDate     =   41988
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Fec.Emisi�n"
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
         Left            =   240
         TabIndex        =   29
         Top             =   4320
         Width           =   1455
      End
      Begin VB.TextBox Text5 
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
         Height          =   375
         Left            =   240
         TabIndex        =   28
         Top             =   3840
         Width           =   3615
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Nro.Tarjeta"
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
         Left            =   240
         TabIndex        =   27
         Top             =   3480
         Width           =   1455
      End
      Begin VB.TextBox Text4 
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
         Height          =   375
         Left            =   1800
         TabIndex        =   26
         Top             =   3000
         Width           =   1215
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Lote"
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
         Left            =   240
         TabIndex        =   25
         Top             =   3000
         Width           =   1455
      End
      Begin VB.TextBox Text3 
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
         Height          =   375
         Left            =   1800
         TabIndex        =   24
         Top             =   2520
         Width           =   1215
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Nro.Cup�n"
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
         Left            =   240
         TabIndex        =   23
         Top             =   2520
         Width           =   1455
      End
      Begin VB.TextBox Text2 
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
         Height          =   375
         Left            =   1800
         TabIndex        =   22
         Top             =   2040
         Width           =   1215
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Banco"
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
         Left            =   240
         TabIndex        =   20
         Top             =   1200
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Cuotas"
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
         Left            =   240
         TabIndex        =   19
         Top             =   2040
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Tarjeta"
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
         Left            =   240
         TabIndex        =   17
         Top             =   360
         Width           =   1215
      End
      Begin MSDataListLib.DataCombo DataCombo2 
         Bindings        =   "frmrecibo_ctacte.frx":046D
         Height          =   420
         Left            =   240
         TabIndex        =   18
         Top             =   720
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   741
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         ListField       =   "tarjeta"
         BoundColumn     =   "id"
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
      Begin MSDataListLib.DataCombo DataCombo3 
         Bindings        =   "frmrecibo_ctacte.frx":0487
         Height          =   420
         Left            =   240
         TabIndex        =   21
         Top             =   1560
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   741
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         ListField       =   "banco"
         BoundColumn     =   "id"
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
   Begin MSAdodcLib.Adodc datvalores 
      Height          =   330
      Left            =   4560
      Top             =   9000
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
   Begin MSAdodcLib.Adodc datencabezado 
      Height          =   330
      Left            =   4560
      Top             =   8640
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
   Begin MSAdodcLib.Adodc datbanco 
      Height          =   330
      Left            =   5280
      Top             =   8640
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
   Begin MSAdodcLib.Adodc dattarjetas 
      Height          =   330
      Left            =   5760
      Top             =   8760
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
   Begin VB.PictureBox Picture1 
      Height          =   5175
      Left            =   120
      ScaleHeight     =   5115
      ScaleWidth      =   8475
      TabIndex        =   12
      Top             =   3240
      Width           =   8535
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   7
         Left            =   2160
         MaxLength       =   50
         TabIndex        =   2
         Top             =   720
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   5
         Left            =   6240
         MaxLength       =   50
         TabIndex        =   65
         Top             =   2640
         Width           =   1935
      End
      Begin VB.TextBox Text7 
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
         Left            =   2160
         MaxLength       =   50
         TabIndex        =   4
         Top             =   1680
         Width           =   5895
      End
      Begin MSComCtl2.DTPicker fecha 
         Height          =   375
         Left            =   6000
         TabIndex        =   48
         Top             =   720
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   49676289
         CurrentDate     =   42060
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   2
         Left            =   2160
         MaxLength       =   50
         TabIndex        =   45
         Top             =   2640
         Width           =   1935
      End
      Begin VB.CommandButton CALCULA 
         Caption         =   "CALCULA"
         Height          =   255
         Left            =   7200
         TabIndex        =   15
         Top             =   4680
         Visible         =   0   'False
         Width           =   1095
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grillac 
         Height          =   1575
         Left            =   120
         TabIndex        =   6
         Top             =   3360
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   2778
         _Version        =   393216
         Rows            =   30
         Cols            =   17
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   17
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   1
         Left            =   2160
         MaxLength       =   50
         TabIndex        =   5
         Top             =   2160
         Width           =   1935
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "frmrecibo_ctacte.frx":049E
         Height          =   420
         Left            =   2160
         TabIndex        =   3
         Top             =   1200
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   741
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         ListField       =   "valor"
         BoundColumn     =   "id"
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
      Begin MSDataListLib.DataCombo DataCombo4 
         Bindings        =   "frmrecibo_ctacte.frx":04B7
         Height          =   420
         Left            =   2160
         TabIndex        =   80
         Top             =   120
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   741
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         ListField       =   "cuentabanco"
         BoundColumn     =   "id"
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
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Cta. Bancaria:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   14
         Left            =   120
         TabIndex        =   79
         Top             =   120
         Width           =   1935
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "N� Manual Rec.:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   12
         Left            =   120
         TabIndex        =   77
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Saldo Recibo:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   10
         Left            =   4440
         TabIndex        =   64
         Top             =   2640
         Width           =   1815
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Detalle:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   120
         TabIndex        =   49
         Top             =   1680
         Width           =   1935
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   4800
         TabIndex        =   47
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Ing.  $:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   46
         Top             =   2640
         Width           =   1935
      End
      Begin VB.Label Label2 
         Caption         =   "Tecla Supr , Borra Ingreso"
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   4920
         Width           =   2175
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   7
         Left            =   120
         TabIndex        =   14
         Top             =   2160
         Width           =   1935
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   8280
         Y1              =   3240
         Y2              =   3240
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo de Valor:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   120
         TabIndex        =   10
         Top             =   1200
         Width           =   2055
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Importe  $:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   120
         TabIndex        =   13
         Top             =   2160
         Width           =   1935
      End
   End
   Begin VB.Frame Frame3 
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
      Height          =   975
      Left            =   120
      TabIndex        =   11
      Top             =   8400
      Width           =   8535
      Begin VB.CheckBox Check1 
         Caption         =   "Envia Por Mail"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2280
         TabIndex        =   91
         Top             =   240
         Visible         =   0   'False
         Width           =   2055
      End
      Begin KewlButtonz.KewlButtons grabar 
         Height          =   615
         Left            =   720
         TabIndex        =   8
         Top             =   240
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   1085
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
         BCOL            =   -2147483645
         BCOLO           =   -2147483645
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmrecibo_ctacte.frx":04D1
         PICN            =   "frmrecibo_ctacte.frx":04ED
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin KewlButtonz.KewlButtons cancelar 
         Height          =   615
         Left            =   7200
         TabIndex        =   9
         Top             =   240
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
         BCOL            =   -2147483645
         BCOLO           =   -2147483645
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmrecibo_ctacte.frx":1F6F
         PICN            =   "frmrecibo_ctacte.frx":1F8B
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSAdodcLib.Adodc datitems 
         Height          =   330
         Left            =   0
         Top             =   120
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
      Begin MSAdodcLib.Adodc datcontrol 
         Height          =   330
         Left            =   0
         Top             =   840
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
      Begin MSAdodcLib.Adodc datcola 
         Height          =   330
         Left            =   2040
         Top             =   360
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
      Begin MSAdodcLib.Adodc datpago 
         Height          =   330
         Left            =   1320
         Top             =   240
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
         Caption         =   "datpago"
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
      Begin MSAdodcLib.Adodc datimputaciones 
         Height          =   330
         Left            =   1440
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
         Caption         =   "datpago"
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
   Begin MSRDC.MSRDC reporte 
      Height          =   375
      Left            =   8880
      Top             =   7800
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
   Begin Crystal.CrystalReport CrystalReporte 
      Left            =   10440
      Top             =   7800
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTitle     =   "Libro IVA Compras"
      PrintFileLinesPerPage=   60
   End
   Begin MSAdodcLib.Adodc datparametros 
      Height          =   330
      Left            =   11040
      Top             =   7920
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
   Begin KewlButtonz.KewlButtons bclientes 
      Height          =   375
      Left            =   8160
      TabIndex        =   50
      ToolTipText     =   "F3-Busca Clientes"
      Top             =   120
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   661
      BTYPE           =   14
      TX              =   ""
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
      BCOL            =   -2147483645
      BCOLO           =   -2147483645
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmrecibo_ctacte.frx":299D
      PICN            =   "frmrecibo_ctacte.frx":29B9
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSDataGridLib.DataGrid DataGrid2 
      Bindings        =   "frmrecibo_ctacte.frx":2F53
      Height          =   375
      Left            =   3120
      TabIndex        =   51
      Top             =   0
      Visible         =   0   'False
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   661
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
   Begin MSAdodcLib.Adodc datcliente 
      Height          =   330
      Left            =   0
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
   Begin MSAdodcLib.Adodc datcp 
      Height          =   330
      Left            =   0
      Top             =   240
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
      Caption         =   "datcp"
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grillaf 
      Height          =   1575
      Left            =   8640
      TabIndex        =   7
      Top             =   600
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   2778
      _Version        =   393216
      Rows            =   200
      Cols            =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   4
   End
   Begin MSAdodcLib.Adodc datctabanco 
      Height          =   330
      Left            =   11640
      Top             =   8160
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
   Begin MSAdodcLib.Adodc datcontrol2 
      Height          =   330
      Left            =   1320
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
   Begin MSAdodcLib.Adodc datcredito 
      Height          =   330
      Left            =   2760
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
      Caption         =   "datcredito"
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
Attribute VB_Name = "frmrecibo_ctacte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public xubi As String
Public xclaveprimaria As Double
Public xcp As String

Function dejarNumerosPuntos(cadenaTexto As String) As String
  Const listaNumeros = "0123456789"
  Dim cadenaTemporal As String
  Dim i As Integer

  cadenaTexto = Trim$(cadenaTexto)
  If Len(cadenaTexto) = 0 Then
    Exit Function
  End If
 
  cadenaTemporal = ""

  For i = 1 To Len(cadenaTexto)
    If InStr(listaNumeros, Mid$(cadenaTexto, i, 1)) Then
      cadenaTemporal = cadenaTemporal + Mid$(cadenaTexto, i, 1)
    End If
  Next
  dejarNumerosPuntos = cadenaTemporal
  
End Function

Private Sub bclientes_Click()

    
  If Text1(0).Text <> "" Then
   If Text19.Text = "" Then
    Text1(0).Text = Replace(Text1(0).Text, " ", "%%")
    xbusqueda = "%" + Text1(0).Text + "%"
    xquery1 = "SELECT  ALIAS_0.ID AS ID,  ALIAS_0.CODIGO, ALIAS_3.NOMBRE AS RAZONSOCIAL, ALIAS_3.CUIT,ALIAS_6.CALLE, ISNULL(ALIAS_6.CALLE,'') + '-'+ ISNULL(V_CIUDAD_.NOMBRE,'') + '-'+ ISNULL(V_PROVINCIA_.NOMBRE,'') AS DOMICILIO, ALIAS_0.DENOMINACION,ALIAS_5.NOMBRE AS ZONA, " & _
              "ALIAS_7.NUMERO AS TELEFONO, ALIAS_8.DIRECCIONELECTRONICA AS MAIL, " & _
              "V_TIPOPAGO_.NOMBRE AS TP, V_TIPOPAGO_.OBSERVACION as TipoPago, V_TIPOPAGO_.ACTIVESTATUS AS ACTIVESTATUSTP, V_PROVINCIA_.NOMBRE AS Provincia, " & _
              "V_CIUDAD_.NOMBRE AS Ciudad, ALIAS_0.CODIGO + ' ' + ALIAS_3.CUIT + ' ' + ALIAS_3.NOMBRE AS CONCATENADO, V_EZI_POCICION_IVA_CLIENTES.CODIGO AS IVA, V_TIPOPAGO_.ID AS IDPAGO, ALIAS_0.DOMICILIOFACTURACION_ID as domicilio_id " & _
              "FROM V_PERSONA_ AS ALIAS_3 RIGHT OUTER JOIN V_CLIENTE AS ALIAS_0 WITH (READPAST) LEFT OUTER JOIN " & _
              "V_EZI_POCICION_IVA_CLIENTES ON ALIAS_0.ID = V_EZI_POCICION_IVA_CLIENTES.idcliente LEFT OUTER JOIN " & _
              "V_TIPOPAGO_ ON ALIAS_0.TIPOPAGO_ID = V_TIPOPAGO_.ID ON ALIAS_3.ID = ALIAS_0.ENTEASOCIADO_ID LEFT OUTER JOIN " & _
              "V_CIUDAD_ RIGHT OUTER JOIN V_DOMICILIO_ AS ALIAS_6 ON V_CIUDAD_.ID = ALIAS_6.CIUDAD_ID LEFT OUTER JOIN " & _
              "V_PROVINCIA_ ON ALIAS_6.PROVINCIA_ID = V_PROVINCIA_.ID ON ALIAS_0.DOMICILIOFACTURACION_ID = ALIAS_6.ID LEFT OUTER JOIN " & _
              "V_TELEFONO_ AS ALIAS_7 ON ALIAS_3.TELEFONOPRINCIPAL_ID = ALIAS_7.ID LEFT OUTER JOIN " & _
              "V_DIRECCIONELECTRONICA_ AS ALIAS_8 ON ALIAS_3.DIRECELECTRONICAPRINCIPAL_ID = ALIAS_8.ID LEFT OUTER JOIN " & _
              "V_ZONA_ AS ALIAS_5 ON ALIAS_0.ZONA_ID = ALIAS_5.ID " & _
              "WHERE     (ALIAS_0.ACTIVESTATUS = 0) AND ALIAS_0.CODIGO + ' ' + ALIAS_3.CUIT + ' ' + ALIAS_3.NOMBRE + ' ' + ALIAS_0.DENOMINACION like '" & xbusqueda & "' order by ALIAS_3.NOMBRE "
   Else
    xquery1 = "SELECT  ALIAS_0.ID AS ID,  ALIAS_0.CODIGO, ALIAS_3.NOMBRE AS RAZONSOCIAL, ALIAS_3.CUIT,ALIAS_6.CALLE, ISNULL(ALIAS_6.CALLE,'') + '-'+ ISNULL(V_CIUDAD_.NOMBRE,'') + '-'+ ISNULL(V_PROVINCIA_.NOMBRE,'') AS DOMICILIO, ALIAS_0.DENOMINACION,ALIAS_5.NOMBRE AS ZONA, " & _
              "ALIAS_7.NUMERO AS TELEFONO, ALIAS_8.DIRECCIONELECTRONICA AS MAIL, " & _
              "V_TIPOPAGO_.NOMBRE AS TP, V_TIPOPAGO_.OBSERVACION as TipoPago, V_TIPOPAGO_.ACTIVESTATUS AS ACTIVESTATUSTP, V_PROVINCIA_.NOMBRE AS Provincia, " & _
              "V_CIUDAD_.NOMBRE AS Ciudad, ALIAS_0.CODIGO + ' ' + ALIAS_3.CUIT + ' ' + ALIAS_3.NOMBRE AS CONCATENADO, V_EZI_POCICION_IVA_CLIENTES.CODIGO AS IVA, V_TIPOPAGO_.ID AS IDPAGO, ALIAS_0.DOMICILIOFACTURACION_ID as domicilio_id " & _
              "FROM V_PERSONA_ AS ALIAS_3 RIGHT OUTER JOIN V_CLIENTE AS ALIAS_0 WITH (READPAST) LEFT OUTER JOIN " & _
              "V_EZI_POCICION_IVA_CLIENTES ON ALIAS_0.ID = V_EZI_POCICION_IVA_CLIENTES.idcliente LEFT OUTER JOIN " & _
              "V_TIPOPAGO_ ON ALIAS_0.TIPOPAGO_ID = V_TIPOPAGO_.ID ON ALIAS_3.ID = ALIAS_0.ENTEASOCIADO_ID LEFT OUTER JOIN " & _
              "V_CIUDAD_ RIGHT OUTER JOIN V_DOMICILIO_ AS ALIAS_6 ON V_CIUDAD_.ID = ALIAS_6.CIUDAD_ID LEFT OUTER JOIN " & _
              "V_PROVINCIA_ ON ALIAS_6.PROVINCIA_ID = V_PROVINCIA_.ID ON ALIAS_0.DOMICILIOFACTURACION_ID = ALIAS_6.ID LEFT OUTER JOIN " & _
              "V_TELEFONO_ AS ALIAS_7 ON ALIAS_3.TELEFONOPRINCIPAL_ID = ALIAS_7.ID LEFT OUTER JOIN " & _
              "V_DIRECCIONELECTRONICA_ AS ALIAS_8 ON ALIAS_3.DIRECELECTRONICAPRINCIPAL_ID = ALIAS_8.ID LEFT OUTER JOIN " & _
              "V_ZONA_ AS ALIAS_5 ON ALIAS_0.ZONA_ID = ALIAS_5.ID " & _
              "WHERE   (ALIAS_0.ACTIVESTATUS = 0) AND  ALIAS_0.ID = '" & Text19.Text & "'"
   End If
  Else
    xquery1 = "SELECT  ALIAS_0.ID AS ID,  ALIAS_0.CODIGO, ALIAS_3.NOMBRE AS RAZONSOCIAL, ALIAS_3.CUIT,ALIAS_6.CALLE, ISNULL(ALIAS_6.CALLE,'') + '-'+ ISNULL(V_CIUDAD_.NOMBRE,'') + '-'+ ISNULL(V_PROVINCIA_.NOMBRE,'') AS DOMICILIO, ALIAS_0.DENOMINACION,ALIAS_5.NOMBRE AS ZONA, " & _
              "ALIAS_7.NUMERO AS TELEFONO, ALIAS_8.DIRECCIONELECTRONICA AS MAIL, " & _
              "V_TIPOPAGO_.NOMBRE AS TP, V_TIPOPAGO_.OBSERVACION as TipoPago, V_TIPOPAGO_.ACTIVESTATUS AS ACTIVESTATUSTP, V_PROVINCIA_.NOMBRE AS Provincia, " & _
              "V_CIUDAD_.NOMBRE AS Ciudad, ALIAS_0.CODIGO + ' ' + ALIAS_3.CUIT + ' ' + ALIAS_3.NOMBRE AS CONCATENADO, V_EZI_POCICION_IVA_CLIENTES.CODIGO AS IVA, V_TIPOPAGO_.ID AS IDPAGO, ALIAS_0.DOMICILIOFACTURACION_ID as domicilio_id  " & _
              "FROM V_PERSONA_ AS ALIAS_3 RIGHT OUTER JOIN V_CLIENTE AS ALIAS_0 WITH (READPAST) LEFT OUTER JOIN " & _
              "V_EZI_POCICION_IVA_CLIENTES ON ALIAS_0.ID = V_EZI_POCICION_IVA_CLIENTES.idcliente LEFT OUTER JOIN " & _
              "V_TIPOPAGO_ ON ALIAS_0.TIPOPAGO_ID = V_TIPOPAGO_.ID ON ALIAS_3.ID = ALIAS_0.ENTEASOCIADO_ID LEFT OUTER JOIN " & _
              "V_CIUDAD_ RIGHT OUTER JOIN V_DOMICILIO_ AS ALIAS_6 ON V_CIUDAD_.ID = ALIAS_6.CIUDAD_ID LEFT OUTER JOIN " & _
              "V_PROVINCIA_ ON ALIAS_6.PROVINCIA_ID = V_PROVINCIA_.ID ON ALIAS_0.DOMICILIOFACTURACION_ID = ALIAS_6.ID LEFT OUTER JOIN " & _
              "V_TELEFONO_ AS ALIAS_7 ON ALIAS_3.TELEFONOPRINCIPAL_ID = ALIAS_7.ID LEFT OUTER JOIN " & _
              "V_DIRECCIONELECTRONICA_ AS ALIAS_8 ON ALIAS_3.DIRECELECTRONICAPRINCIPAL_ID = ALIAS_8.ID LEFT OUTER JOIN " & _
              "V_ZONA_ AS ALIAS_5 ON ALIAS_0.ZONA_ID = ALIAS_5.ID " & _
              "WHERE     (ALIAS_0.ACTIVESTATUS = 0) order by ALIAS_3.NOMBRE "
  End If

    Text19.Text = ""
    datcliente.RecordSource = xquery1
    datcliente.Refresh
    If datcliente.Recordset.EOF = True Then
        mensa = MsgBox("No existe Cliente", vbInformation, "!! Atencion !!")
        Text1(0).Text = ""
        Text1(0).SetFocus
    End If
    
    datcliente.Recordset.MoveFirst
    If datcliente.Recordset.RecordCount = 1 Then
        Text1(0).Text = DataGrid2.Columns(2).Text
        
        If login.nomsucursal = "TUCUMAN" Then xcp = ""
        If login.nomsucursal = "TUCUMANZIP" Then xcp = "DIM TOLEDO VAL"
        If login.nomsucursal = "EMPORIO" Then xcp = "EL EMP.TUCUMAN"
        If login.nomsucursal = "EMPORIOZIP" Then xcp = "COMPRADOR Tucuman"
        
'        xquerycp = "SELECT     ALIAS_0.ID, ALIAS_0.TRORIGINANTE_ID, ALIAS_0.DESCRIPCION AS ALIAS_0_DESCRIPCION, ALIAS_0.IMTOTAL2_IMPORTE AS TOTAL, " & _
'                   "ALIAS_0.SALDO2_IMPORTE AS SALDO, ALIAS_0.TIPO AS ALIAS_0_TIPO, ALIAS_0.FECHAEMISION, substring(ALIAS_0.FECHAVENCIMIENTO,7,2)+'/'+substring(ALIAS_0.FECHAVENCIMIENTO,7,2)+'/'+left(ALIAS_0.FECHAVENCIMIENTO,4) as FECHAVENCIMIENTO , " & _
'                   "ALIAS_3.DESCRIPCION AS ALIAS_3_DESCRIPCION, ALIAS_0.NOMCLASIFICADOR AS ALIAS_0_NOMCLASIFICADOR, ALIAS_0.OPERADORCOMERCIAL_ID " & _
'                   "FROM V_COMPROMISOPAGO_ AS ALIAS_0 LEFT OUTER JOIN " & _
'                   "V_ESTADO_ AS ALIAS_3 ON ALIAS_0.ESTADO_ID = ALIAS_3.ID " & _
'                   "WHERE     (ALIAS_0.NIVEL = 1) AND (ALIAS_0.CPCOMPRAS = 'F') AND (ALIAS_0.SALDADO = 'F') AND (ALIAS_0.SUMA = 'T') AND  " & _
'                   "(ALIAS_0.OPERADORCOMERCIAL_ID = '" & DataGrid2.Columns(0).Text & "') AND (ALIAS_0.NOMCLASIFICADOR = '" & xcp & "') " & _
'                   "ORDER BY ALIAS_0.FECHAEMISION"
                   
        xquerycp = "SELECT     ALIAS_0.ID, ALIAS_0.TRORIGINANTE_ID, ALIAS_0.DESCRIPCION AS ALIAS_0_DESCRIPCION, ALIAS_0.IMTOTAL2_IMPORTE AS TOTAL, " & _
                   "ALIAS_0.SALDO2_IMPORTE AS SALDO, ALIAS_0.TIPO AS ALIAS_0_TIPO, ALIAS_0.FECHAEMISION, substring(ALIAS_0.FECHAVENCIMIENTO,7,2)+'/'+substring(ALIAS_0.FECHAVENCIMIENTO,7,2)+'/'+left(ALIAS_0.FECHAVENCIMIENTO,4) as FECHAVENCIMIENTO , " & _
                   "ALIAS_3.DESCRIPCION AS ALIAS_3_DESCRIPCION, ALIAS_0.NOMCLASIFICADOR AS ALIAS_0_NOMCLASIFICADOR, ALIAS_0.OPERADORCOMERCIAL_ID " & _
                   "FROM V_COMPROMISOPAGO_ AS ALIAS_0 LEFT OUTER JOIN " & _
                   "V_ESTADO_ AS ALIAS_3 ON ALIAS_0.ESTADO_ID = ALIAS_3.ID " & _
                   "WHERE     (ALIAS_0.NIVEL = 1) AND (ALIAS_0.CPCOMPRAS = 'F') AND (ALIAS_0.SALDADO = 'F') AND (ALIAS_0.SUMA = 'T') AND  " & _
                   "(ALIAS_0.OPERADORCOMERCIAL_ID = '" & DataGrid2.Columns(0).Text & "' ) " & _
                   "ORDER BY ALIAS_0.FECHAEMISION"
                   
        datcp.RecordSource = xquerycp
        datcp.Refresh
        If datcp.Recordset.EOF = False Then
            xsaldo = 0
            datcp.Recordset.MoveFirst
            Do While Not datcp.Recordset.EOF
                xsaldo = xsaldo + datcp.Recordset.Fields("saldo")
                datcp.Recordset.MoveNext
            Loop
            Text1(8).Text = Format(xsaldo, "###,###,##0.00")
            datcp.Recordset.MoveFirst
        End If
        
        DataGrid1.Columns(0).Visible = False
        DataGrid1.Columns(1).Visible = False
        DataGrid1.Columns(5).Visible = False
        DataGrid1.Columns(6).Visible = False
        DataGrid1.Columns(8).Visible = False
        DataGrid1.Columns(9).Visible = False
        DataGrid1.Columns(10).Visible = False
        
        DataGrid1.Columns(2).Caption = "Comprobante"
        DataGrid1.Columns(2).Width = 4500
        DataGrid1.Columns(3).Alignment = dbgRight
        DataGrid1.Columns(3).NumberFormat = "currency"
        DataGrid1.Columns(3).Width = 1000
        DataGrid1.Columns(4).Alignment = dbgRight
        DataGrid1.Columns(4).NumberFormat = "currency"
        DataGrid1.Columns(4).Width = 1000
        DataGrid1.Columns(7).Alignment = dbgCenter
        DataGrid1.Columns(7).Caption = "Fec.Venc."
        DataGrid1.Columns(7).Width = 1000
        For X = 0 To 7
            DataGrid1.Columns(X).Locked = True
        Next
        
        DataGrid1.SetFocus
' Indica LImite de Credito
        datcredito.RecordSource = "select * from v_ezi_pos_ctacte_control where id = '" & DataGrid2.Columns(0).Text & "' and nomclasificador = '" & xcp & "'"
        datcredito.Refresh
    
        If datcredito.Recordset.EOF = False Then
            If datcredito.Recordset.Fields("creditomaximo") <> 0 Then
                xdisponible = datcredito.Recordset.Fields("creditomaximo") - Val(datcredito.Recordset.Fields("saldo")) + Val(Format(Text1(6).Text, "#####0.00"))
                Text1(9).Text = Format(xdisponible, "###,##0.00")
                Text1(10).Text = Format(datcredito.Recordset.Fields("creditomaximo"), "###,##0.00")
            Else
                Text1(9).Text = "Sin Limite"
                Text1(10).Text = "No Asignado"
            End If
        Else
            Text1(9).Text = ""
            Text1(10).Text = "No Asignado"
        End If
        
    Else
        menu = 8
        query = xquery1
        lista_clientes.Show
    End If


End Sub


Private Sub calcula_Click()
On Error Resume Next
xpagos = 0
For X = 1 To 29
  If grillac.TextMatrix(X, 3) = "" Then
     xvalorgrilla = 0
  Else
     xvalorgrilla = Round(grillac.TextMatrix(X, 3), 10)
  End If
  xpagos = xpagos + xvalorgrilla
Next X

For X = 1 To 199
  If grillaf.TextMatrix(X, 3) = "" Then
     xvalorcancela = 0
  Else
     xvalorcancela = Round(grillaf.TextMatrix(X, 3), 10)
  End If
  xcancela = xcancela + xvalorcancela
Next X


    
    Text1(2).Text = Format(xpagos, "###,##0.00")
    
    Text1(1).Text = Format(0, "###,##0.00")
    Text1(0).SetFocus
 
    xcancela = Round(xcancela, 2)
    Text1(6).Text = Format(xcancela, "###,##0.00")
    
    Text1(5).Text = Format(Round(xpagos - xcancela, 2), "###,##0.00")
    
    If xubi = "C" Then DataGrid1.SetFocus
    If xubi = "I" Then DataCombo1.SetFocus
    
''--- Limite de Credito
    If login.nomsucursal = "TUCUMAN" Or login.nomsucursal = "JUJUY" Then xcp = "DIM TOLEDO"
    If login.nomsucursal = "TUCUMANZIP" Then xcp = "DIM TOLEDO VAL"
    If login.nomsucursal = "EMPORIO" Then xcp = "EL EMP.TUCUMAN"
    If login.nomsucursal = "EMPORIOZIP" Then xcp = "COMPRADOR Tucuman"
        
' Indica LImite de Credito
If Text1(0).Text <> "" Then
        datcredito.RecordSource = "select * from v_ezi_pos_ctacte_control where id = '" & DataGrid2.Columns(0).Text & "' and nomclasificador = '" & xcp & "'"
        datcredito.Refresh
    
        If datcredito.Recordset.EOF = False Then
            If datcredito.Recordset.Fields("creditomaximo") <> 0 Then
                xdisponible = datcredito.Recordset.Fields("creditomaximo") - Val(datcredito.Recordset.Fields("saldo")) + Val(Format(Text1(6).Text, "#####0.00"))
                Text1(9).Text = Format(xdisponible, "###,##0.00")
                Text1(10).Text = Format(datcredito.Recordset.Fields("creditomaximo"), "###,##0.00")
            Else
                Text1(9).Text = "Sin Limite"
                Text1(10).Text = "No Asignado"
            End If
        Else
            Text1(9).Text = ""
            Text1(10).Text = "No Asignado"
        End If
End If



End Sub

Private Sub Cancelar_Click()


    Unload Me

End Sub

Private Sub DataList1_KeyPress(KeyAscii As Integer)
On Error GoTo fuera

    If KeyAscii = 13 Then
        KeyAscii = 0
        Text1(5).Text = DataList1.BoundText
        Text1(6).SetFocus
    End If

fuera:

End Sub

Private Sub DataList1_LostFocus()

    DataList1.Visible = False

End Sub

Private Sub DataCombo1_GotFocus()

    
DataCombo1.Text = "Efectivo"
tarjeta.Visible = False
DataCombo2.Text = ""
DataCombo3.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""

femicioncheque.Value = Date - 1
fechavencimiento.Value = Date
femision.Value = Date
femisionret.Value = Date
DataCombo5.Text = ""
Text9.Text = ""
Text8.Text = ""
Text6.Text = ""

Text11.Text = ""
Text10.Text = "N"

cheques.Visible = False
tarjeta.Visible = False
retenciones.Visible = False


End Sub

Private Sub DataCombo1_KeyPress(KeyAscii As Integer)
On Error Resume Next

    If KeyAscii = 13 Then
        KeyAscii = 0
        Text7.SetFocus
        If Left(DataCombo1.Text, 7) = "Tarjeta" Then tarjeta.Visible = True
    End If

End Sub

Private Sub DataList2_KeyPress(KeyAscii As Integer)
On Error GoTo fuera

    If KeyAscii = 13 Then
        KeyAscii = 0
        Text1(6).Text = DataList2.BoundText
        grabar.SetFocus
    End If

fuera:
End Sub

Private Sub DataList2_LostFocus()

    DataList2.Visible = False

End Sub

Private Sub DataList3_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        Text1(4).Text = DataList3.Text
        Text1(5).SetFocus
        DataList3.Visible = False
    End If


End Sub

Private Sub DataCombo2_KeyPress(KeyAscii As Integer)

On Error Resume Next

    If KeyAscii = 13 Then
        KeyAscii = 0
        DataCombo3.SetFocus
    End If


End Sub

Private Sub DataCombo3_KeyPress(KeyAscii As Integer)

On Error Resume Next

    If KeyAscii = 13 Then
        KeyAscii = 0
        Text2.SetFocus
    End If

End Sub


Private Sub DataCombo4_KeyPress(KeyAscii As Integer)

On Error Resume Next

    If KeyAscii = 13 Then
        KeyAscii = 0
        DataCombo1.SetFocus
    End If

End Sub

Private Sub DataCombo5_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{TAB}", False
    End If
End Sub

Private Sub DataGrid1_DblClick()

On Error GoTo fuera
Dim crApp As New CRAXDRT.Application
Dim crReport As New CRAXDRT.Report
Dim tabla As String
Dim ruta As String


reporte.SQL = "SELECT v_ezi_pos_factctacte.id, v_ezi_pos_factctacte.NUMERODOCUMENTO, v_ezi_pos_factctacte.FECHAEMISION, v_ezi_pos_factctacte.cod_cliente, v_ezi_pos_factctacte.cliente, v_ezi_pos_factctacte.CUIT, v_ezi_pos_factctacte.CALLE, v_ezi_pos_factctacte.CODPOS, v_ezi_pos_factctacte.provincia, v_ezi_pos_factctacte.detalle, v_ezi_pos_factctacte.tipopago, v_ezi_pos_factctacte.codigoproducto, v_ezi_pos_factctacte.nombre_producto, v_ezi_pos_factctacte.cantidadproducto, v_ezi_pos_factctacte.nota, v_ezi_pos_factctacte.condiva, v_ezi_pos_factctacte.ciudad, v_ezi_pos_factctacte.TIPOVENTA, v_ezi_pos_factctacte.SIMBOLO, v_ezi_pos_factctacte.CODVENDEDOR, v_ezi_pos_factctacte.preciusiniva, v_ezi_pos_factctacte.subtotalsiniva, v_ezi_pos_factctacte.impbonifsiniva, v_ezi_pos_factctacte.nroremito, v_ezi_pos_factctacte.percepiibb, v_ezi_pos_factctacte.perceptem, v_ezi_pos_factctacte.totaltr, v_ezi_pos_factctacte.importeiva21, v_ezi_pos_factctacte.importeiva105, v_ezi_pos_factctacte.iditem,  " & _
              "v_ezi_pos_factctacte.cae, v_ezi_pos_factctacte.vto " & _
              "FROM  MMOSSE.DBO.v_ezi_pos_factctacte v_ezi_pos_factctacte " & _
              "where v_ezi_pos_factctacte.calipsoid = '" & DataGrid1.Columns(1).Text & "' order by v_ezi_pos_factctacte.iditem"

tabla = reporte.SQL

datcontrol2.ConnectionString = login.conexiontotal

datcontrol.RecordSource = "select nota from v_trfacturaventa_ where id = '" & DataGrid1.Columns(1).Text & "' "
datcontrol.Refresh
If datcontrol.Recordset.EOF = False Then
    xletra = datcontrol.Recordset.Fields("nota")
Else
    xletra = "B"
End If
If xletra = "" Then xletra = "B"

With CrystalReporte
    .PrinterCollation = crptCollated
   
    If xletra = "A" Then
         .ReportFileName = App.Path & "\PresupuestoA.rpt"
    Else
        .ReportFileName = App.Path & "\PresupuestoB.rpt"
    End If
    .WindowTitle = "Factura Vta Orig"
    .Connect = login.conexionreporte
    .DiscardSavedData = True
    .RetrieveDataFiles
    .ReportSource = 0
    .SQLQuery = tabla
    .Destination = crptToWindow
 '   .Destination = crptToPrinter
    .PrintFileType = crptCrystal
    .WindowState = crptMaximized
    .Action = 1
End With
    
Exit Sub

fuera:
    
    MsgBox "Reporte de Factura no Encontado, o error de configuracion de reporte", vbCritical, "Error"


Exit Sub



End Sub

Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
On Error Resume Next

    If KeyAscii = 13 Then
        KeyAscii = 0
        
        For X = 1 To 200
             If DataGrid1.Columns(0).Text = grillaf.TextMatrix(X, 0) Then
                MsgBox "Atencion, Comprobante ya seleccionado", vbInformation, "Atenci�n"
                DataGrid1.SetFocus
                Exit Sub
             End If
             If grillaf.TextMatrix(X, 0) = "" Then Exit For
        Next X
        
        
        cancelafactura.Visible = True
        Text1(3).Text = Format(DataGrid1.Columns(4).Text, "###,##0.00")
        Text1(4).Text = Format(DataGrid1.Columns(4).Text, "###,##0.00")
        Text1(4).SetFocus
        
        
    End If



End Sub

Private Sub femicioncheque_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{TAB}", False
    End If
    

End Sub

Private Sub Form_Activate()
Call calcula_Click



End Sub

Private Sub Form_Load()
Aplicar_skin Me

yventana = Inicio.Height / 2 - 1000
xventana = Inicio.Width / 2

frmrecibo_ctacte.Top = yventana - frmrecibo_ctacte.Height / 2
frmrecibo_ctacte.Left = xventana - frmrecibo_ctacte.Width / 2
fecha = Date
xubi = ""
Text1(5).Text = 0

datvalores.ConnectionString = login.conexiontotal
dattarjetas.ConnectionString = login.conexiontotal
datbanco.ConnectionString = login.conexiontotal
datencabezado.ConnectionString = login.conexiontotal
datitems.ConnectionString = login.conexiontotal
datcontrol.ConnectionString = login.conexiontotal
datcola.ConnectionString = login.conexiontotal
datpago.ConnectionString = login.conexiontotal
datimputaciones.ConnectionString = login.conexiontotal
datparametros.ConnectionString = login.conexiontotal
datcliente.ConnectionString = login.conexiontotal
datcp.ConnectionString = login.conexiontotal
datcredito.ConnectionString = login.conexiontotal
    
        
    datvalores.RecordSource = "SELECT ID, TIPOVALOR AS VALOR, CONSOLIDACIONCAJA FROM V_TIPOVALOR_ AS ALIAS_0 " & _
                              "WHERE (ACTIVESTATUS = 0) AND (TIPOVALOR LIKE '%Efec%' OR TIPOVALOR LIKE '%Sufrida%' OR TIPOVALOR LIKE '%Cr�dito%' OR " & _
                              "TIPOVALOR LIKE '%tercer%%dife%') order by TIPOVALOR desc"
    datvalores.Refresh

    
    dattarjetas.RecordSource = "select ID, NOMBRE as tarjeta from V_TARJETACREDITO_ order by NOMBRE"
    dattarjetas.Refresh
    
    datbanco.RecordSource = "select ID, ENTEASOCIADOSUCURSAL AS BANCO from V_BANCO_ ORDER BY ENTEASOCIADOSUCURSAL"
    datbanco.Refresh
    
    datimputaciones.RecordSource = "SELECT     ALIAS_0.ID, ALIAS_0.NOMBRE AS NOMBRE " & _
                                   "FROM         V_IMPUTACIONCONTABLE_ AS ALIAS_0 LEFT OUTER JOIN  " & _
                                   "V_UNIDADOPERATIVA_ AS ALIAS_1 ON ALIAS_0.UNIDADOPERATIVA_ID = ALIAS_1.ID " & _
                                   "WHERE     (ALIAS_0.BO_PLACE_ID = '{89C234D2-3F01-11D5-86AD-0080AD403F5F}') AND (ALIAS_0.ACTIVESTATUS = 0) AND EXISTS " & _
                                   "(SELECT     ID " & _
                                   "FROM          PERSLIST WITH (READPAST)  " & _
                                   "WHERE      (ID =(SELECT     BO_ITEMS_ID " & _
                                   "FROM          BOLIST WITH (READPAST) " & _
                                   "WHERE      (ID = ALIAS_0.LISTATIPOSTRANSACCION_ID))) AND (ITEM_ID = '{6D720AC9-E8C2-11D5-B0C2-004854841C8A}')) ORDER BY NOMBRE"
    datimputaciones.Refresh
    
    datparametros.RecordSource = "select * from ud_ezi_parametros_pos where sucursal = '" & login.nomsucursal & "' "
    datparametros.Refresh
    
    datctabanco.ConnectionString = login.conexiontotal
    datctabanco.RecordSource = "select null as ID, '' AS CUENTABANCO union all " & _
                               "SELECT     TOP (100) ID, DESCRIPCION + ' - ' + NUMERO AS CUENTABANCO " & _
                               "FROM         V_CUENTABANCARIA_ AS b " & _
                               "WHERE     (BO_PLACE_ID = '{9B9915F8-4FA6-11D5-B060-004854841C8A}') AND (ACTIVESTATUS <> 2) AND (ACTIVESTATUS <> 2)"
    datctabanco.Refresh

grillac.Row = 0
grillac.Col = 0
grillac.ColWidth(0) = 100
grillac.Col = 1
grillac.Text = "T.Valor"
grillac.ColWidth(1) = 2000
grillac.Col = 2
grillac.Text = "Detalle"
grillac.ColWidth(2) = 1500
grillac.Col = 3
grillac.Text = "Importe"
grillac.ColWidth(3) = 1000
grillac.Col = 4
grillac.Text = "IdTarjeta"
grillac.ColWidth(4) = 10
grillac.Col = 5
grillac.Text = "Idbanco"
grillac.ColWidth(5) = 0
grillac.Col = 6
grillac.Text = "cuotas"
grillac.ColWidth(6) = 0
grillac.Col = 7
grillac.Text = "nrocupon"
grillac.ColWidth(7) = 0
grillac.Col = 8
grillac.Text = "lote"
grillac.ColWidth(8) = 0
grillac.Col = 9
grillac.Text = "nrotarjeta"
grillac.ColWidth(9) = 0
grillac.Col = 10
grillac.Text = "femision"
grillac.ColWidth(10) = 0
'' Cheques
grillac.Col = 11
grillac.Text = "fvencimiento"
grillac.ColWidth(11) = 0
grillac.Col = 12
grillac.Text = "numerocheq"
grillac.ColWidth(12) = 0
grillac.Col = 13
grillac.Text = "numerocta"
grillac.ColWidth(13) = 0
grillac.Col = 14
grillac.Text = "anombrede"
grillac.ColWidth(14) = 0
grillac.Col = 15
grillac.Text = "Imputacion"
grillac.ColWidth(15) = 3500
grillac.Col = 16
grillac.Text = "idimputacion"
grillac.ColWidth(16) = 0

grillaf.Row = 0
grillaf.Col = 0
grillaf.ColWidth(0) = 100
grillaf.Col = 1
grillaf.Text = "Comprobante"
grillaf.ColWidth(1) = 3000
grillaf.Col = 2
grillaf.Text = "Saldo"
grillaf.ColWidth(2) = 1000
grillaf.Col = 3
grillaf.Text = "Cancela"
grillaf.ColWidth(3) = 1000



   
End Sub

Private Sub Frame1_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub grabar_Click()
On Error Resume Next

    If Round(Text1(5).Text, 2) < 0 Then
        mensa = MsgBox("No se puede Grabar el Comprobante en Cero o en Negativo", vbCritical, "!! Error !!")
        Exit Sub
    End If
    
 mensa = MsgBox("Desea Grabar este Ingreso de Valor ?", vbYesNo, "!! Atenci�n !!")
 If mensa = vbYes Then
    
    datencabezado.RecordSource = "SELECT MAX(CONVERT(decimal, isnull(claveprimaria,0))) AS claveprimaria FROM ud_ezi_puntodeventa_encabezado with(readpast)"
    datencabezado.Refresh
    
    If datencabezado.Recordset.EOF = True Then
        xclaveprimaria = 1
    Else
        xclaveprimaria = datencabezado.Recordset.Fields("claveprimaria") + 1
    End If
    
    If IsNull(claveprimaria) = True Then xclaveprimaria = 1
    datencabezado.RecordSource = "SELECT * From ud_ezi_puntodeventa_encabezado with(readpast) where id =0 "
    datencabezado.Refresh
    
    datencabezado.Recordset.AddNew
    datencabezado.Recordset.Fields("claveprimaria") = xclaveprimaria
    datencabezado.Recordset.Fields("numeradorinterno") = "Recibo de Cobrana"
    datencabezado.Recordset.Fields("fechadelcomprobante") = fecha.Value
    datencabezado.Recordset.Fields("detalle") = Text7.Text
    datencabezado.Recordset.Fields("sucursal") = datparametros.Recordset.Fields("sucursal")
    datencabezado.Recordset.Fields("cotizacion") = 1
    datencabezado.Recordset.Fields("alquiler") = "N"
    datencabezado.Recordset.Fields("importeglobal") = Round(Text1(2).Text, 2)
    datencabezado.Recordset.Fields("generada") = "True"
    datencabezado.Recordset.Fields("importado") = "False"
    datencabezado.Recordset.Fields("clienteid") = DataGrid2.Columns(0).Text
    datencabezado.Recordset.Fields("cliente") = Text1(0).Text
    If DataCombo4.Text = "" Then
        datencabezado.Recordset.Fields("vendedorid") = datparametros.Recordset("cajadefecto")
    Else
        datencabezado.Recordset.Fields("vendedorid") = DataCombo4.BoundText
    End If
    datencabezado.Recordset.Fields("nombrepc") = Environ("computername")
    datencabezado.Recordset.Fields("target") = datparametros.Recordset.Fields("sucursal")
    datencabezado.Recordset.Fields("transferido") = "False"
    datencabezado.Recordset.Fields("totaltr") = Round(Text1(2).Text, 2)
    datencabezado.Recordset.Fields("numerador") = Text1(7).Text
    
    datencabezado.Recordset.UpdateBatch adAffectCurrent
    xid = datencabezado.Recordset.Fields("id")
    
    '** Establene numero de Facturas Manuales, y no Fiscales
'    If Text1(7).Text <> "" Then
        xnumerador = "Recibo Cobranza " + datparametros.Recordset.Fields("sucursal")
'    Else
'        xnumerador = "Recibo Cobranza Sistema"
'    End If
    
    datcola.RecordSource = "SELECT p.ID AS ID, p.NOMBRE AS puntero, n.NUMERO AS NUMERO, p.CARACTERESPREFIJO AS puntoventa, p.NUMERO_ID " & _
                           "FROM V_NUMERADOR_ AS p LEFT OUTER JOIN V_NUMERO_ AS n ON p.NUMERO_ID = n.ID " & _
                           "WHERE     (p.ACTIVESTATUS <> 2) AND (p.NOMBRE = '" & xnumerador & "') "
    datcola.Refresh
    datencabezado.Recordset.Fields("numerodefactura") = datcola.Recordset.Fields("numero")
    datencabezado.Recordset.Fields("puntodeventa") = datcola.Recordset.Fields("puntoventa")
    If Text1(7).Text = "" Then
        datencabezado.Recordset.Fields("numerador") = datcola.Recordset.Fields("numero")
    End If

    MsgBox "Nro de Recibo de Sistema: " + Str(datcola.Recordset.Fields("puntoventa")) + "-" + Str(datcola.Recordset.Fields("numero")), vbInformation, "Recibo Sistema"
    
       
        xnumero = datcola.Recordset.Fields("numero")
        xidnumero = datcola.Recordset.Fields("numero_id")

        datcola.RecordSource = "Select * from numero with(readpast) where id = '" & xidnumero & "'"
        datcola.Refresh
        datcola.Recordset.Fields("numero") = xnumero + 1
        datcola.Recordset.UpdateBatch adAffectCurrent


    '** Fin de asignacion de numero a Factura

    datencabezado.Recordset.Fields("claveprimaria") = xid
    datencabezado.Recordset.UpdateBatch adAffectCurrent
    
    
'--- Graba Items
    datitems.RecordSource = "select * from ud_ezi_ingreso where id = 0"
    datitems.Refresh
    
    For X = 1 To 200
        If grillaf.TextMatrix(X, 0) = "" Then Exit For
        
        datitems.Recordset.AddNew
        datitems.Recordset.Fields("claveprimaria") = xid
        datitems.Recordset.Fields("origenid") = grillaf.TextMatrix(X, 0)
        datitems.Recordset.Fields("descripcion") = grillaf.TextMatrix(X, 1)
        datitems.Recordset.Fields("monto") = Round(grillaf.TextMatrix(X, 3), 2)
        datitems.Recordset.UpdateBatch adAffectCurrent
    Next X
    
    
    
'******* Graba Pago

    datpago.RecordSource = "Select * from ud_ezi_pago where claveprimaria = ''"
    datpago.Refresh
    For X = 1 To 29
        If grillac.TextMatrix(X, 0) = "" Then Exit For
        datpago.Recordset.AddNew
        datpago.Recordset.Fields("claveprimaria") = xid
        datpago.Recordset.Fields("tipovalor") = "True"
        datpago.Recordset.Fields("valoroseniaid") = grillac.TextMatrix(X, 0)
        If DataCombo4.Text = "" Then
            datpago.Recordset.Fields("destinoid") = datparametros.Recordset.Fields("cajadefecto")
            datpago.Recordset.Fields("formadepago") = grillac.TextMatrix(X, 1)
        Else
            datpago.Recordset.Fields("destinoid") = DataCombo4.BoundText
            If grillac.TextMatrix(X, 1) = "Efectivo" Then
                datpago.Recordset.Fields("formadepago") = "Debito en Cuenta Corriente"
            Else
                datpago.Recordset.Fields("formadepago") = grillac.TextMatrix(X, 1)
            End If
            datpago.Recordset.Fields("responsable") = "Deposito en Banco"
        End If

        datpago.Recordset.Fields("monto") = Round(grillac.TextMatrix(X, 3), 2)
        datpago.Recordset.Fields("caja") = 1
        
        If Left(grillac.TextMatrix(X, 1), 7) = "Tarjeta" Then
            datpago.Recordset.Fields("bancoid") = grillac.TextMatrix(X, 5)
            datpago.Recordset.Fields("cuotas") = grillac.TextMatrix(X, 6)
            datpago.Recordset.Fields("fechadeemision") = grillac.TextMatrix(X, 10)
            datpago.Recordset.Fields("numerodecupon") = grillac.TextMatrix(X, 7)
            datpago.Recordset.Fields("numerodetarjeta") = grillac.TextMatrix(X, 9)
            datpago.Recordset.Fields("tarjetaid") = grillac.TextMatrix(X, 4)
        End If
        If Left(grillac.TextMatrix(X, 1), 6) = "Cheque" Then
            datpago.Recordset.Fields("bancoid") = grillac.TextMatrix(X, 5)
            datpago.Recordset.Fields("fechadeemision") = grillac.TextMatrix(X, 10)
            datpago.Recordset.Fields("fechadevencimiento") = grillac.TextMatrix(X, 11)
            datpago.Recordset.Fields("numero") = grillac.TextMatrix(X, 12)
            datpago.Recordset.Fields("responsable") = grillac.TextMatrix(X, 14) + " " + grillac.TextMatrix(X, 13)
        End If
        If Left(grillac.TextMatrix(X, 1), 3) = "Ret" Then
            datpago.Recordset.Fields("bancoid") = grillac.TextMatrix(X, 5)
            datpago.Recordset.Fields("fechadeemision") = grillac.TextMatrix(X, 10)
            
            xnumeroret = Trim(dejarNumerosPuntos(grillac.TextMatrix(X, 12)))
            If xnumeroret = "" Then xnumeroret = 1
            datpago.Recordset.Fields("numero") = xnumeroret
            
            datpago.Recordset.Fields("responsable") = grillac.TextMatrix(X, 13)
        End If
        datpago.Recordset.Fields("sucursal") = login.nomsucursal
        
        datpago.Recordset.UpdateBatch adAffectCurrent
     Next X


'******* Graba ud_ezi_cola

    
        datcola.RecordSource = "Select * from ud_ezi_cola_importar where id = 0"
        datcola.Refresh
        
        datcola.Recordset.AddNew
        datcola.Recordset.Fields("id_encabezado") = xid
        datcola.Recordset.Fields("tipodedocumentoid") = frmnota_venta.datparametros.Recordset.Fields("idrecibocobranza")
        datcola.Recordset.Fields("unidadoperativaid") = frmnota_venta.datparametros.Recordset.Fields("target")
        datcola.Recordset.Fields("fecha_hora") = Date
        
        
        datcola.Recordset.UpdateBatch adAffectCurrent
        
        mensa = MsgBox("Imprime recibo ?", vbYesNo, "Impresi�n")
        
        If mensa = vbYes Then Call imprimerecibo_Click
                
        mensa = MsgBox("Envia por Mail ?", vbYesNo, "E-Mail")
        
        If mensa = vbYes Then
                Check1.Value = 1
                Call imprimerecibo_Click
        Else
                Check1.Value = 0
        End If
                
        Unload Me
        frmrecibo_ctacte.Show

 End If
 
Exit Sub
errorgrabar:
    mensa = MsgBox("No se pudo registrar la informaci�n", vbCritical, "Error !!")


End Sub


Private Sub grillac_KeyUp(KeyCode As Integer, Shift As Integer)

    
If KeyCode = 46 Then
        For X = 0 To 16
            grillac.Col = X
            grillac.Text = ""
        Next X
        
        For X = grillac.Row + 1 To 29
            For Y = 0 To 3
                grillac.Col = Y
                grillac.Row = X
                xcampo = grillac.Text
                grillac.Row = X - 1
                grillac.Text = xcampo
            Next Y
        Next X

    Call calcula_Click
End If

End Sub

Private Sub salir_Click()

    Unload Me

End Sub



Private Sub grillaf_KeyUp(KeyCode As Integer, Shift As Integer)

If KeyCode = 46 Then
        For X = 0 To 3
            grillaf.Col = X
            grillaf.Text = ""
        Next X
        
        For X = grillaf.Row + 1 To 199
            For Y = 0 To 3
                grillaf.Col = Y
                grillaf.Row = X
                xcampo = grillaf.Text
                grillaf.Row = X - 1
                grillaf.Text = xcampo
            Next Y
        Next X

'''' Recalcula nuevamente la grilla de CP
         For X = 1 To 200
           If grillaf.TextMatrix(X, 1) = "" Then Exit For
            xfacturascp = xfacturascp + " <> " + "'" + grillaf.TextMatrix(X, 1) + "'" + " And ALIAS_0.DESCRIPCION "
         Next X
        If grillaf.TextMatrix(1, 1) <> "" Then
            xfacturascp = Left(xfacturascp, Len(xfacturascp) - 25)
        
          xquerycp = "SELECT     ALIAS_0.ID, ALIAS_0.TRORIGINANTE_ID, ALIAS_0.DESCRIPCION AS ALIAS_0_DESCRIPCION, ALIAS_0.IMTOTAL2_IMPORTE AS TOTAL, " & _
                   "ALIAS_0.SALDO2_IMPORTE AS SALDO, ALIAS_0.TIPO AS ALIAS_0_TIPO, ALIAS_0.FECHAEMISION, substring(ALIAS_0.FECHAVENCIMIENTO,7,2)+'/'+substring(ALIAS_0.FECHAVENCIMIENTO,7,2)+'/'+left(ALIAS_0.FECHAVENCIMIENTO,4) as FECHAVENCIMIENTO , " & _
                   "ALIAS_3.DESCRIPCION AS ALIAS_3_DESCRIPCION, ALIAS_0.NOMCLASIFICADOR AS ALIAS_0_NOMCLASIFICADOR, ALIAS_0.OPERADORCOMERCIAL_ID " & _
                   "FROM V_COMPROMISOPAGO_ AS ALIAS_0 LEFT OUTER JOIN " & _
                   "V_ESTADO_ AS ALIAS_3 ON ALIAS_0.ESTADO_ID = ALIAS_3.ID " & _
                   "WHERE     (ALIAS_0.NIVEL = 1) AND (ALIAS_0.CPCOMPRAS = 'F') AND (ALIAS_0.SALDADO = 'F') AND (ALIAS_0.SUMA = 'T') AND  " & _
                   "(ALIAS_0.OPERADORCOMERCIAL_ID = '" & DataGrid2.Columns(0).Text & "') AND (ALIAS_0.NOMCLASIFICADOR = '" & xcp & "') and ALIAS_0.DESCRIPCION " & xfacturascp & " " & _
                   "ORDER BY ALIAS_0.FECHAEMISION"
        Else
          xquerycp = "SELECT     ALIAS_0.ID, ALIAS_0.TRORIGINANTE_ID, ALIAS_0.DESCRIPCION AS ALIAS_0_DESCRIPCION, ALIAS_0.IMTOTAL2_IMPORTE AS TOTAL, " & _
                   "ALIAS_0.SALDO2_IMPORTE AS SALDO, ALIAS_0.TIPO AS ALIAS_0_TIPO, ALIAS_0.FECHAEMISION, substring(ALIAS_0.FECHAVENCIMIENTO,7,2)+'/'+substring(ALIAS_0.FECHAVENCIMIENTO,7,2)+'/'+left(ALIAS_0.FECHAVENCIMIENTO,4) as FECHAVENCIMIENTO , " & _
                   "ALIAS_3.DESCRIPCION AS ALIAS_3_DESCRIPCION, ALIAS_0.NOMCLASIFICADOR AS ALIAS_0_NOMCLASIFICADOR, ALIAS_0.OPERADORCOMERCIAL_ID " & _
                   "FROM V_COMPROMISOPAGO_ AS ALIAS_0 LEFT OUTER JOIN " & _
                   "V_ESTADO_ AS ALIAS_3 ON ALIAS_0.ESTADO_ID = ALIAS_3.ID " & _
                   "WHERE     (ALIAS_0.NIVEL = 1) AND (ALIAS_0.CPCOMPRAS = 'F') AND (ALIAS_0.SALDADO = 'F') AND (ALIAS_0.SUMA = 'T') AND  " & _
                   "(ALIAS_0.OPERADORCOMERCIAL_ID = '" & DataGrid2.Columns(0).Text & "') AND (ALIAS_0.NOMCLASIFICADOR = '" & xcp & "') " & _
                   "ORDER BY ALIAS_0.FECHAEMISION"
        End If
                   
        datcp.RecordSource = xquerycp
        datcp.Refresh
        If datcp.Recordset.EOF = False Then
            xsaldo = 0
            datcp.Recordset.MoveFirst
            Do While Not datcp.Recordset.EOF
                xsaldo = xsaldo + datcp.Recordset.Fields("saldo")
                datcp.Recordset.MoveNext
            Loop
            Text1(8).Text = Format(xsaldo, "###,###,##0.00")
            datcp.Recordset.MoveFirst
        End If
        
        DataGrid1.Columns(0).Visible = False
        DataGrid1.Columns(1).Visible = False
        DataGrid1.Columns(5).Visible = False
        DataGrid1.Columns(6).Visible = False
        DataGrid1.Columns(8).Visible = False
        DataGrid1.Columns(9).Visible = False
        DataGrid1.Columns(10).Visible = False
        
        DataGrid1.Columns(2).Caption = "Comprobante"
        DataGrid1.Columns(2).Width = 4500
        DataGrid1.Columns(3).Alignment = dbgRight
        DataGrid1.Columns(3).NumberFormat = "currency"
        DataGrid1.Columns(3).Width = 1000
        DataGrid1.Columns(4).Alignment = dbgRight
        DataGrid1.Columns(4).NumberFormat = "currency"
        DataGrid1.Columns(4).Width = 1000
        DataGrid1.Columns(7).Alignment = dbgCenter
        DataGrid1.Columns(7).Caption = "Fec.Venc."
        DataGrid1.Columns(7).Width = 1000
        For X = 0 To 7
            DataGrid1.Columns(X).Locked = True
        Next
        
        DataGrid1.SetFocus


''''--------------


    Call calcula_Click
End If



End Sub

Private Sub imprimerecibo_Click()
On Error Resume Next
Dim crApp As New CRAXDRT.Application
Dim crReport As New CRAXDRT.Report
Dim tabla As String
Dim ruta As String

'If Text1(7).Text <> "" Then Exit Sub

reporte.SQL = "SELECT v_ezi_pos_recibo.GRUPO,v_ezi_pos_recibo.id, v_ezi_pos_recibo.nrorecibo, v_ezi_pos_recibo.fechadelcomprobante, v_ezi_pos_recibo.cliente, v_ezi_pos_recibo.CUIT, v_ezi_pos_recibo.CODPOS, v_ezi_pos_recibo.CALLE, v_ezi_pos_recibo.LOCALIDAD, v_ezi_pos_recibo.CONDIVA, v_ezi_pos_recibo.Comprobante, v_ezi_pos_recibo.Cancela, v_ezi_pos_recibo.totalfactura, v_ezi_pos_recibo.formadepago, v_ezi_pos_recibo.banco, v_ezi_pos_recibo.tarjeta, v_ezi_pos_recibo.numerocheque, v_ezi_pos_recibo.monto, v_ezi_pos_recibo.fechaemision, v_ezi_pos_recibo.fechavencimiento FROM MMOSSE.dbo.v_ezi_pos_recibo v_ezi_pos_recibo where v_ezi_pos_recibo.id = " & xclaveprimaria & " ORDER BY v_ezi_pos_recibo.GRUPO ASC "
tabla = reporte.SQL


With CrystalReporte
    .PrinterCollation = crptCollated
    .ReportFileName = App.Path & "\Recibos.rpt"
    .WindowTitle = "Remito Vta Orig"
    '.Connect = "PROVIDER=MSDASQL;dsn=facturacion;uid=lucva;pwd=25072004;database=facturacionsql;"
    .Connect = login.conexionreporte
    .DiscardSavedData = True
    .RetrieveDataFiles
    .ReportSource = 0
    .SQLQuery = tabla
    
  If Check1.Value = 1 Then
     Kill ("c:\util\*.pdf")
     Kill ("c:\util\*.rtf")


    .Destination = crptToFile
    .PrintFileType = crptRTF

    .PrintFileName = "c:\util\Recibo " + Replace(Text1(0).Text, "*", "") + ".pdf"

    .WindowState = crptNormal
    .Action = 1
    
'    PDFCreator_CreatePDF Me, "c:\util\Recibo " + Replace(Text1(0).Text, "*", "") + ".rtf", "c:\util\CtaCte " + Replace(Text1(0).Text, "*", "") + ".pdf"
  Else
'    .Destination = crptToWindow
    .Destination = crptToPrinter
    .PrintFileType = crptCrystal
    .WindowState = crptMaximized
    .Action = 1
  End If

End With

If Check1.Value = 1 Then
 Set oApp = New Outlook.Application
 Set myItem = oApp.CreateItem(Outlook.OlItemType.olMailItem)
 Set myAttachments = myItem.Attachments
 myAttachments.Add "c:\util\Recibo " + Replace(Text1(0).Text, "*", "") + ".pdf", 4, 2, " "
 myItem.Display

 Set oApp = Nothing
 Set myItem = Nothing
 Set myAttachments = Nothing
End If





End Sub

Private Sub Text1_GotFocus(Index As Integer)

    If Index = 1 Then
        Text1(1).SelStart = 0
        Text1(1).SelLength = Len(Text1(1).Text)
    End If

    If Index = 4 Then
        Text1(4).SelStart = 0
        Text1(4).SelLength = Len(Text1(4).Text)
    End If


End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
On Error Resume Next

    If KeyAscii = 13 Then
       KeyAscii = 0
       If Index = 0 Then
            Call bclientes_Click
       End If
       
       If Index = 7 Then
            DataCombo1.SetFocus
       End If
       
       If Index = 4 Then
         If Round(Text1(4).Text, 2) = 0 Then
            cancelafactura.Visible = False
            DataGrid1.SetFocus
            Exit Sub
         End If
         If Round(Text1(4).Text, 2) > Round(Text1(3).Text, 2) Then
            MsgBox "EL importe a Saldar no puede ser mayor al saldo Actual", vbCritical, "Error"
            cancelafactura.Visible = False
            DataGrid1.SetFocus
            Exit Sub
         End If
          
         Text1(4).Text = Format(Text1(4).Text, "###,##0.00")
         xfacturascp = ""
         For X = 1 To 200
          If grillaf.TextMatrix(X, 1) = "" Then
                grillaf.TextMatrix(X, 0) = DataGrid1.Columns(0).Text
                grillaf.TextMatrix(X, 1) = DataGrid1.Columns(2).Text
                grillaf.TextMatrix(X, 2) = Format(Text1(3).Text, "###,##0.00")
                grillaf.TextMatrix(X, 3) = Format(Text1(4).Text, "###,##0.00")
                Exit For
          End If
         Next X

'''' Recalcula nuevamente la grilla de CP
         For X = 1 To 200
           If grillaf.TextMatrix(X, 1) = "" Then Exit For
            xfacturascp = xfacturascp + " <> " + "'" + grillaf.TextMatrix(X, 1) + "'" + " And ALIAS_0.DESCRIPCION "
         Next X
         xfacturascp = Left(xfacturascp, Len(xfacturascp) - 25)
        
'        xquerycp = "SELECT     ALIAS_0.ID, ALIAS_0.TRORIGINANTE_ID, ALIAS_0.DESCRIPCION AS ALIAS_0_DESCRIPCION, ALIAS_0.IMTOTAL2_IMPORTE AS TOTAL, " & _
'                   "ALIAS_0.SALDO2_IMPORTE AS SALDO, ALIAS_0.TIPO AS ALIAS_0_TIPO, ALIAS_0.FECHAEMISION, substring(ALIAS_0.FECHAVENCIMIENTO,7,2)+'/'+substring(ALIAS_0.FECHAVENCIMIENTO,7,2)+'/'+left(ALIAS_0.FECHAVENCIMIENTO,4) as FECHAVENCIMIENTO , " & _
'                   "ALIAS_3.DESCRIPCION AS ALIAS_3_DESCRIPCION, ALIAS_0.NOMCLASIFICADOR AS ALIAS_0_NOMCLASIFICADOR, ALIAS_0.OPERADORCOMERCIAL_ID " & _
'                   "FROM V_COMPROMISOPAGO_ AS ALIAS_0 LEFT OUTER JOIN " & _
'                   "V_ESTADO_ AS ALIAS_3 ON ALIAS_0.ESTADO_ID = ALIAS_3.ID " & _
'                   "WHERE     (ALIAS_0.NIVEL = 1) AND (ALIAS_0.CPCOMPRAS = 'F') AND (ALIAS_0.SALDADO = 'F') AND (ALIAS_0.SUMA = 'T') AND  " & _
'                   "(ALIAS_0.OPERADORCOMERCIAL_ID = '" & DataGrid2.Columns(0).Text & "') AND (ALIAS_0.NOMCLASIFICADOR = '" & xcp & "') and ALIAS_0.DESCRIPCION " & xfacturascp & " " & _
'                   "ORDER BY ALIAS_0.FECHAEMISION"
                   
        xquerycp = "SELECT     ALIAS_0.ID, ALIAS_0.TRORIGINANTE_ID, ALIAS_0.DESCRIPCION AS ALIAS_0_DESCRIPCION, ALIAS_0.IMTOTAL2_IMPORTE AS TOTAL, " & _
                   "ALIAS_0.SALDO2_IMPORTE AS SALDO, ALIAS_0.TIPO AS ALIAS_0_TIPO, ALIAS_0.FECHAEMISION, substring(ALIAS_0.FECHAVENCIMIENTO,7,2)+'/'+substring(ALIAS_0.FECHAVENCIMIENTO,7,2)+'/'+left(ALIAS_0.FECHAVENCIMIENTO,4) as FECHAVENCIMIENTO , " & _
                   "ALIAS_3.DESCRIPCION AS ALIAS_3_DESCRIPCION, ALIAS_0.NOMCLASIFICADOR AS ALIAS_0_NOMCLASIFICADOR, ALIAS_0.OPERADORCOMERCIAL_ID " & _
                   "FROM V_COMPROMISOPAGO_ AS ALIAS_0 LEFT OUTER JOIN " & _
                   "V_ESTADO_ AS ALIAS_3 ON ALIAS_0.ESTADO_ID = ALIAS_3.ID " & _
                   "WHERE     (ALIAS_0.NIVEL = 1) AND (ALIAS_0.CPCOMPRAS = 'F') AND (ALIAS_0.SALDADO = 'F') AND (ALIAS_0.SUMA = 'T') AND  " & _
                   "(ALIAS_0.OPERADORCOMERCIAL_ID = '" & DataGrid2.Columns(0).Text & "') and ALIAS_0.DESCRIPCION " & xfacturascp & " " & _
                   "ORDER BY ALIAS_0.FECHAEMISION"

                   
                   
        datcp.RecordSource = xquerycp
        datcp.Refresh
        If datcp.Recordset.EOF = False Then
            xsaldo = 0
            datcp.Recordset.MoveFirst
            Do While Not datcp.Recordset.EOF
                xsaldo = xsaldo + datcp.Recordset.Fields("saldo")
                datcp.Recordset.MoveNext
            Loop
            Text1(8).Text = Format(xsaldo, "###,###,##0.00")
            datcp.Recordset.MoveFirst
        End If
        
        DataGrid1.Columns(0).Visible = False
        DataGrid1.Columns(1).Visible = False
        DataGrid1.Columns(5).Visible = False
        DataGrid1.Columns(6).Visible = False
        DataGrid1.Columns(8).Visible = False
        DataGrid1.Columns(9).Visible = False
        DataGrid1.Columns(10).Visible = False
        
        DataGrid1.Columns(2).Caption = "Comprobante"
        DataGrid1.Columns(2).Width = 4500
        DataGrid1.Columns(3).Alignment = dbgRight
        DataGrid1.Columns(3).NumberFormat = "currency"
        DataGrid1.Columns(3).Width = 1000
        DataGrid1.Columns(4).Alignment = dbgRight
        DataGrid1.Columns(4).NumberFormat = "currency"
        DataGrid1.Columns(4).Width = 1000
        DataGrid1.Columns(7).Alignment = dbgCenter
        DataGrid1.Columns(7).Caption = "Fec.Venc."
        DataGrid1.Columns(7).Width = 1000
        For X = 0 To 7
            DataGrid1.Columns(X).Locked = True
        Next
        
        DataGrid1.SetFocus


''''--------------
         
         
         cancelafactura.Visible = False
         xubi = "C"
         Call calcula_Click
         
       End If
       
       If Index = 1 Then
        If Val(Text1(1).Text) = 0 Then
            MsgBox "Debe ingresar un importe V�lido", vbCritical, "Error"
            Text1(1).Text = 0
            Text1(1).Text = Format(Text1(1).Text, "###,##0.00")
            Text1(1).SetFocus
            Exit Sub
        End If
        For X = 1 To 30
          If grillac.TextMatrix(X, 1) = "" Then
           xubi = "I"
'*********************** Efectivo
           If DataCombo1.Text = "Efectivo" Then
                grillac.TextMatrix(X, 0) = DataCombo1.BoundText
                grillac.TextMatrix(X, 1) = DataCombo1.Text
                grillac.TextMatrix(X, 3) = Format(Text1(1).Text, "###,##0.00")
                grillac.TextMatrix(X, 15) = DataCombo4.Text
                grillac.TextMatrix(X, 16) = DataCombo4.BoundText
                Call calcula_Click
           End If
'*********************** tarjeta
           If Left(DataCombo1.Text, 7) = "Tarjeta" Then
            tarjeta.Visible = True
            DataCombo2.SetFocus
           End If
'*********************** Cheque
           If Left(DataCombo1.Text, 6) = "Cheque" Then
            cheques.Visible = True
            femicioncheque.SetFocus
           End If
'*********************** retenciones
           If Left(DataCombo1.Text, 3) = "Ret" Then
            retenciones.Visible = True
            femisionret.SetFocus
           End If
           
           
           
           Exit For
          End If
         Next X
       End If

    End If
    
End Sub

Private Sub Text1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 38 Then
    If Index > 0 Then
      Text1(Index - 1).SetFocus
    Else
      DataCombo3.SetFocus
    End If
End If

End Sub

Private Sub Text1_LostFocus(Index As Integer)
On Error Resume Next
        If Index = 2 Then
            If Len(Text1(2).Text) = 12 Then Exit Sub
            For X = 1 To Len(Text1(2).Text)
               car = Mid(Text1(2).Text, X, 1)
               If car = "-" Then
                  PVta = Right("0000" + Left(Text1(2).Text, X - 1), 4)
                  nu = Right("00000000" + Right(Text1(2).Text, Len(Text1(2).Text) - X), 8)
                  Text1(2).Text = PVta + nu
                  Exit Sub
               End If
            Next X
            Text1(2).Text = Right("00000000" + Text1(2).Text, 8)
        End If

End Sub

Private Sub Text10_GotFocus()

    
        Text10.SelStart = 0
        Text10.SelLength = Len(Text10.Text)


End Sub

Private Sub Text10_KeyPress(KeyAscii As Integer)


On Error Resume Next

    If KeyAscii = 13 Then
        KeyAscii = 0
        Text10.Text = UCase(Text10.Text)
        If Text10.Text <> "S" And Text10.Text <> "N" Then
            Text10.Text = "N"
        End If
        
            
        If Text11.Text = "" Then
            MsgBox "Debe Ingresar un Numero de Certificado", vbCritical, "Error"
            Text11.SetFocus
            Exit Sub
        End If
        
        
        For X = 1 To 30
          If grillac.TextMatrix(X, 1) = "" Then
'*********************** Retenciones
                grillac.TextMatrix(X, 0) = DataCombo1.BoundText
                grillac.TextMatrix(X, 1) = DataCombo1.Text
                grillac.TextMatrix(X, 2) = "Nro:" + Text11.Text
                grillac.TextMatrix(X, 3) = Format(Text1(1).Text, "###,##0.00")
                grillac.TextMatrix(X, 10) = femisionret.Value
                grillac.TextMatrix(X, 12) = Text11.Text
                grillac.TextMatrix(X, 13) = Text10.Text
                Exit For
          End If
        Next X
        
        retenciones.Visible = False
        Call calcula_Click

    End If



    

End Sub

Private Sub Text11_KeyPress(KeyAscii As Integer)


    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{TAB}", False
    End If

End Sub

Private Sub Text2_GotFocus()

    Text2.SelStart = 0
    Text2.SelLength = Len(Text2.Text)

End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)

On Error Resume Next

    If KeyAscii = 13 Then
        KeyAscii = 0
        Text3.SetFocus
    End If


End Sub

Private Sub Text3_GotFocus()


    Text3.SelStart = 0
    Text3.SelLength = Len(Text3.Text)

End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)


On Error Resume Next

    If KeyAscii = 13 Then
        KeyAscii = 0
        Text4.SetFocus
    End If


End Sub

Private Sub Text4_GotFocus()

    Text4.SelStart = 0
    Text4.SelLength = Len(Text4.Text)

End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)


On Error Resume Next

    If KeyAscii = 13 Then
        KeyAscii = 0
        Text5.SetFocus
    End If


End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)

On Error Resume Next

    If KeyAscii = 13 Then
        KeyAscii = 0
        
        If DataCombo2.Text = "" Then
            MsgBox "Debe Ingresar una Tarjeta V�lida", vbCritical, "Error"
            DataCombo2.SetFocus
            Exit Sub
        End If
    
        If Text2.Text = "" Then
            MsgBox "Debe Ingresar un Numero de Cuotas", vbCritical, "Error"
            Text2.SetFocus
            Exit Sub
        End If
        
        If Text3.Text = "" Then
            MsgBox "Debe Ingresar un Numero de Cup�n", vbCritical, "Error"
            Text3.SetFocus
            Exit Sub
        End If
        
        
        For X = 1 To 30
          If grillac.TextMatrix(X, 1) = "" Then
'*********************** tarjeta
                grillac.TextMatrix(X, 0) = DataCombo1.BoundText
                grillac.TextMatrix(X, 1) = DataCombo1.Text
                grillac.TextMatrix(X, 2) = DataCombo2.Text + " -Ctas:" + Text2.Text
                grillac.TextMatrix(X, 3) = Format(Text1(1).Text, "###,##0.00")
                grillac.TextMatrix(X, 4) = DataCombo2.BoundText
                grillac.TextMatrix(X, 5) = DataCombo3.BoundText
                grillac.TextMatrix(X, 6) = Val(Text2.Text)
                grillac.TextMatrix(X, 7) = Val(Text3.Text)
                grillac.TextMatrix(X, 8) = Val(Text4.Text)
                grillac.TextMatrix(X, 9) = Text5.Text
                grillac.TextMatrix(X, 10) = femision.Value
                grillac.TextMatrix(X, 15) = DataCombo4.Text
                grillac.TextMatrix(X, 16) = DataCombo4.BoundText
                Exit For
          End If
        Next X
        
        tarjeta.Visible = False
        Call calcula_Click

    End If


End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)

On Error Resume Next

    If KeyAscii = 13 Then
        KeyAscii = 0
        
        If DataCombo5.Text = "" Then
            MsgBox "Debe Ingresar un Banco Valido", vbCritical, "Error"
            DataCombo5.SetFocus
            Exit Sub
        End If
    
        For i = 1 To Len(Text9.Text)
            car = Mid(Text9.Text, i, 1)
            If Asc(car) < 48 Or Asc(car) > 58 Then
              MsgBox "Debe Ingresar un Numero de Cheque V�lido", vbCritical, "Error"
              Text9.Text = ""
              Text9.SetFocus
              Exit Sub
            End If
        Next i
        
        If Text9.Text = "" Then
             MsgBox "Debe Ingresar un Numero de Cheque V�lido", vbCritical, "Error"
             Text9.SetFocus
             Exit Sub
        End If

        
        If fechavencimiento.Value < femicioncheque.Value Then
            MsgBox "La fecha de vencimiento no puede ser menor a la fecha de Emisi�n"
            fechavencimiento.SetFocus
            Exit Sub
        End If
        
        For X = 1 To 30
          If grillac.TextMatrix(X, 1) = "" Then
'*********************** cheque
                grillac.TextMatrix(X, 0) = DataCombo1.BoundText
                grillac.TextMatrix(X, 1) = DataCombo1.Text
                grillac.TextMatrix(X, 2) = "Nro:" + Text9.Text
                grillac.TextMatrix(X, 3) = Format(Text1(1).Text, "###,##0.00")
                grillac.TextMatrix(X, 5) = DataCombo5.BoundText
               
                grillac.TextMatrix(X, 10) = femicioncheque.Value
                grillac.TextMatrix(X, 11) = fechavencimiento.Value
                grillac.TextMatrix(X, 12) = Text9.Text
                grillac.TextMatrix(X, 13) = Text8.Text
                grillac.TextMatrix(X, 14) = Text6.Text
                grillac.TextMatrix(X, 15) = DataCombo4.Text
                grillac.TextMatrix(X, 16) = DataCombo4.BoundText
                Exit For
          End If
        Next X
        
        cheques.Visible = False
        Call calcula_Click

    End If

End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)

On Error Resume Next

    If KeyAscii = 13 Then
        KeyAscii = 0
        Text1(1).SetFocus
        If Left(DataCombo1.Text, 7) = "Tarjeta" Then tarjeta.Visible = True
    End If

End Sub

Private Sub Text8_KeyPress(KeyAscii As Integer)


    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{TAB}", False
    End If

End Sub

Private Sub Text9_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{TAB}", False
    End If

End Sub
