VERSION 5.00
Object = "{F6125AB1-8AB1-11CE-A77F-08002B2F4E98}#2.0#0"; "MSRDC20.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{D18BBD1F-82BB-4385-BED3-E9D31A3E361E}#1.0#0"; "KewlButtonz.ocx"
Begin VB.Form frmcobranzacdo2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cobranza Factura de Contado"
   ClientHeight    =   6150
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11640
   Icon            =   "frmcobranzacdo2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6150
   ScaleWidth      =   11640
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
      Left            =   7320
      TabIndex        =   45
      Top             =   0
      Visible         =   0   'False
      Width           =   4095
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
         TabIndex        =   49
         Top             =   960
         Width           =   1575
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
         TabIndex        =   51
         Top             =   960
         Width           =   1815
      End
      Begin VB.CommandButton Command19 
         Caption         =   "Fec.Emisión"
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
         TabIndex        =   48
         Top             =   360
         Width           =   1575
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
         TabIndex        =   47
         Top             =   1560
         Width           =   1575
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
         TabIndex        =   52
         Top             =   1560
         Width           =   735
      End
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
         TabIndex        =   46
         Top             =   1560
         Width           =   615
      End
      Begin MSComCtl2.DTPicker femisionret 
         Height          =   375
         Left            =   2040
         TabIndex        =   50
         Top             =   360
         Width           =   1815
         _ExtentX        =   3201
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
         Format          =   99418113
         CurrentDate     =   41988
      End
   End
   Begin VB.CommandButton imprimefactura 
      Caption         =   "imprimefactura"
      Height          =   255
      Left            =   7680
      TabIndex        =   44
      Top             =   5520
      Visible         =   0   'False
      Width           =   1695
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
      Left            =   7320
      TabIndex        =   31
      Top             =   0
      Visible         =   0   'False
      Width           =   4095
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
         TabIndex        =   43
         Top             =   840
         Width           =   1455
      End
      Begin MSComCtl2.DTPicker femicioncheque 
         Height          =   375
         Left            =   1800
         TabIndex        =   32
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
         Format          =   99418113
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
         TabIndex        =   42
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
         TabIndex        =   41
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
         TabIndex        =   35
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
         TabIndex        =   40
         Top             =   2760
         Width           =   1455
      End
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
         TabIndex        =   36
         Top             =   2760
         Width           =   2055
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
         TabIndex        =   39
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
         TabIndex        =   37
         Top             =   3600
         Width           =   3615
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Fec.Emisión"
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
         Top             =   360
         Width           =   1455
      End
      Begin MSDataListLib.DataCombo DataCombo5 
         Bindings        =   "frmcobranzacdo2.frx":0442
         Height          =   360
         Left            =   240
         TabIndex        =   34
         Top             =   1680
         Width           =   3615
         _ExtentX        =   6376
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
         TabIndex        =   33
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
         Format          =   99418113
         CurrentDate     =   41988
      End
   End
   Begin VB.Frame tarjeta 
      Caption         =   "Tarjeta de Crédito"
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
      Left            =   7320
      TabIndex        =   14
      Top             =   0
      Visible         =   0   'False
      Width           =   4095
      Begin MSComCtl2.DTPicker femision 
         Height          =   375
         Left            =   1800
         TabIndex        =   28
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
         Format          =   99418113
         CurrentDate     =   41988
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Fec.Emisión"
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
         TabIndex        =   26
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
         TabIndex        =   25
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
         TabIndex        =   24
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
         TabIndex        =   23
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
         TabIndex        =   22
         Top             =   2520
         Width           =   1215
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Nro.Cupón"
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
         TabIndex        =   21
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
         TabIndex        =   20
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
         TabIndex        =   18
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
         TabIndex        =   17
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
         TabIndex        =   15
         Top             =   360
         Width           =   1215
      End
      Begin MSDataListLib.DataCombo DataCombo2 
         Bindings        =   "frmcobranzacdo2.frx":0459
         Height          =   360
         Left            =   240
         TabIndex        =   16
         Top             =   720
         Width           =   3615
         _ExtentX        =   6376
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
         Bindings        =   "frmcobranzacdo2.frx":0473
         Height          =   360
         Left            =   240
         TabIndex        =   19
         Top             =   1560
         Width           =   3615
         _ExtentX        =   6376
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
      Left            =   3360
      Top             =   5400
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
      Left            =   3360
      Top             =   5040
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
      Left            =   2160
      Top             =   5040
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
      Left            =   2160
      Top             =   5400
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
      Height          =   4575
      Left            =   120
      ScaleHeight     =   4515
      ScaleWidth      =   6795
      TabIndex        =   9
      Top             =   120
      Width           =   6855
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
         Left            =   4560
         MaxLength       =   50
         TabIndex        =   5
         Top             =   1200
         Width           =   1935
      End
      Begin VB.CommandButton CALCULA 
         Caption         =   "CALCULA"
         Height          =   255
         Left            =   360
         TabIndex        =   13
         Top             =   1680
         Visible         =   0   'False
         Width           =   1095
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grillac 
         Height          =   2175
         Left            =   120
         TabIndex        =   6
         Top             =   2040
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   3836
         _Version        =   393216
         Rows            =   30
         Cols            =   15
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
         _Band(0).Cols   =   15
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
         Index           =   0
         Left            =   2400
         MaxLength       =   50
         TabIndex        =   3
         Top             =   720
         Width           =   1935
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
         Left            =   2400
         MaxLength       =   50
         TabIndex        =   4
         Top             =   1200
         Width           =   1935
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "frmcobranzacdo2.frx":048A
         Height          =   360
         Left            =   2400
         TabIndex        =   2
         Top             =   240
         Width           =   4215
         _ExtentX        =   7435
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
      Begin VB.Label Label2 
         Caption         =   "Tecla Supr , Borra cobranza"
         Height          =   255
         Left            =   240
         TabIndex        =   30
         Top             =   4200
         Width           =   2175
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Saldo a Cobrar"
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
         Left            =   4560
         TabIndex        =   29
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "A Cobrar $:"
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
         TabIndex        =   12
         Top             =   1200
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
         TabIndex        =   11
         Top             =   2040
         Width           =   1935
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   6720
         Y1              =   1920
         Y2              =   1920
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
         TabIndex        =   0
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Importe Total $:"
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
         TabIndex        =   10
         Top             =   720
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
      Height          =   1215
      Left            =   120
      TabIndex        =   1
      Top             =   4800
      Width           =   6855
      Begin KewlButtonz.KewlButtons grabar 
         Height          =   615
         Left            =   720
         TabIndex        =   7
         Top             =   360
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
         MICON           =   "frmcobranzacdo2.frx":04A3
         PICN            =   "frmcobranzacdo2.frx":04BF
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
         Cancel          =   -1  'True
         Height          =   615
         Left            =   4920
         TabIndex        =   8
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
         BCOL            =   -2147483645
         BCOLO           =   -2147483645
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmcobranzacdo2.frx":1F41
         PICN            =   "frmcobranzacdo2.frx":1F5D
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
      Begin MSAdodcLib.Adodc datpago 
         Height          =   330
         Left            =   3240
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
      Left            =   9360
      Top             =   5520
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
      Left            =   10560
      Top             =   5520
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTitle     =   "Libro IVA Compras"
      PrintFileLinesPerPage=   60
   End
End
Attribute VB_Name = "frmcobranzacdo2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public xid As Double


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
    
    If Text1(0).Text = "" Then Text1(0).Text = 0
    Text1(1).Text = Format(Round(Text1(0).Text, 10) - xpagos, "###,##0.00")
    Text1(2).Text = Format(Round(Text1(0).Text, 10) - xpagos, "###,##0.00")
    
    If Round(Text1(1).Text, 2) = 0 Then
        grabar.SetFocus
'        grabar_Click
    Else
        DataCombo1.SetFocus
    End If
    
    

End Sub

Private Sub Cancelar_Click()

    frmnota_venta_cobrar.Text1(1).SetFocus
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
DataCombo5.Text = ""
Text9.Text = ""
Text8.Text = ""
Text6.Text = ""

cheques.Visible = False
tarjeta.Visible = False
retenciones.Visible = False

End Sub

Private Sub DataCombo1_KeyPress(KeyAscii As Integer)
On Error Resume Next

    If KeyAscii = 13 Then
        KeyAscii = 0
        Text1(1).SetFocus
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


Private Sub DataCombo5_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{TAB}", False
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

frmcobranzacdo2.DataCombo1.SetFocus

End Sub

Private Sub Form_Load()
Aplicar_skin Me

yventana = Inicio.Height / 2 - 1000
xventana = Inicio.Width / 2

frmcobranzacdo2.Top = yventana - frmcobranzacdo2.Height / 2
frmcobranzacdo2.Left = xventana - frmcobranzacdo2.Width / 2



datvalores.ConnectionString = login.conexiontotal
dattarjetas.ConnectionString = login.conexiontotal
datbanco.ConnectionString = login.conexiontotal
datencabezado.ConnectionString = login.conexiontotal
datitems.ConnectionString = login.conexiontotal
datcontrol.ConnectionString = login.conexiontotal
datcola.ConnectionString = login.conexiontotal
datpago.ConnectionString = login.conexiontotal

    
    datvalores.RecordSource = "SELECT ID, TIPOVALOR AS VALOR, CONSOLIDACIONCAJA FROM V_TIPOVALOR_ AS ALIAS_0 " & _
                              "WHERE (ACTIVESTATUS = 0) AND (TIPOVALOR LIKE '%Efec%' OR TIPOVALOR LIKE '%Sufrida%' OR TIPOVALOR LIKE '%Crédito%' OR " & _
                              "TIPOVALOR LIKE '%tercer%%dife%') order by TIPOVALOR desc"
    
        
'        datvalores.RecordSource = "SELECT ID, TIPOVALOR AS VALOR, CONSOLIDACIONCAJA FROM V_TIPOVALOR_ AS ALIAS_0 " & _
'                              "WHERE (ACTIVESTATUS = 0) AND (TIPOVALOR LIKE '%Efec%' OR TIPOVALOR LIKE '%Crédito%' OR " & _
'                              "TIPOVALOR LIKE '%tercer%%dife%') order by TIPOVALOR desc"
    datvalores.Refresh

    
    dattarjetas.RecordSource = "select ID, NOMBRE as tarjeta from V_TARJETACREDITO_ order by NOMBRE"
    dattarjetas.Refresh
    
    datbanco.RecordSource = "select ID, ENTEASOCIADOSUCURSAL AS BANCO from V_BANCO_ ORDER BY ENTEASOCIADOSUCURSAL"
    datbanco.Refresh
    


grillac.Row = 0
grillac.Col = 0
grillac.ColWidth(0) = 100
grillac.Col = 1
grillac.Text = "T.Valor"
grillac.ColWidth(1) = 2000
grillac.Col = 2
grillac.Text = "Detalle"
grillac.ColWidth(2) = 2000
grillac.Col = 3
grillac.Text = "Importe"
grillac.ColWidth(3) = 1000
grillac.Col = 4
grillac.Text = "IdTarjeta"
grillac.ColWidth(4) = 10
grillac.Col = 5
grillac.Text = "Idbanco"
grillac.ColWidth(5) = 10
grillac.Col = 6
grillac.Text = "cuotas"
grillac.ColWidth(6) = 10
grillac.Col = 7
grillac.Text = "nrocupon"
grillac.ColWidth(7) = 10
grillac.Col = 8
grillac.Text = "lote"
grillac.ColWidth(8) = 10
grillac.Col = 9
grillac.Text = "nrotarjeta"
grillac.ColWidth(9) = 10
grillac.Col = 10
grillac.Text = "femision"
grillac.ColWidth(10) = 10
'' Cheques
grillac.Col = 11
grillac.Text = "fvencimiento"
grillac.ColWidth(11) = 10
grillac.Col = 12
grillac.Text = "numerocheq"
grillac.ColWidth(12) = 10
grillac.Col = 13
grillac.Text = "numerocta"
grillac.ColWidth(13) = 10
grillac.Col = 14
grillac.Text = "anombrede"
grillac.ColWidth(14) = 10








    
   
End Sub

Private Sub Frame1_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub grabar_Click()
'On Error Resume Next

    If Round(Text1(1).Text, 2) <> 0 Then
        mensa = MsgBox("No se puede Grabar si no se cancela el total del Comprobante", vbCritical, "!! Error !!")
        Exit Sub
    End If
    
 mensa = MsgBox("Desea Grabar esta Factura de Contado ?", vbYesNo, "!! Atención !!")
 If mensa = vbYes Then
    
    datencabezado.RecordSource = "SELECT MAX(CONVERT(decimal, isnull(claveprimaria,0))) AS claveprimaria FROM ud_ezi_puntodeventa_encabezado with(readpast) "
    datencabezado.Refresh
    
    If datencabezado.Recordset.EOF = True Then
        xclaveprimaria = 1
    Else
        xclaveprimaria = datencabezado.Recordset.Fields("claveprimaria") + 1
    End If
    
    If IsNull(claveprimaria) = True Then xclaveprimaria = 1
    datencabezado.RecordSource = "SELECT * From ud_ezi_puntodeventa_encabezado with(readpast) where id =0 "
    datencabezado.Refresh
    datitems.RecordSource = "Select * from ud_ezi_puntodeventa_detalle_factm with(readpast) where id = 0"
    datitems.Refresh
    
    datencabezado.Recordset.AddNew
    datencabezado.Recordset.Fields("claveprimaria") = xclaveprimaria
    datencabezado.Recordset.Fields("numeradorinterno") = "Factura de Venta Mostrador"
    datencabezado.Recordset.Fields("fechadelcomprobante") = DateValue(frmnota_venta_cobrar.Text3.Text) + TimeValue(Str(Time))
    datencabezado.Recordset.Fields("sucursal") = frmnota_venta_cobrar.datparametros.Recordset.Fields("sucursal")
    datencabezado.Recordset.Fields("clienteid") = frmnota_venta_cobrar.DataGrid2.Columns(0).Text
    datencabezado.Recordset.Fields("cliente") = frmnota_venta_cobrar.DataGrid2.Columns(2).Text
    If frmnota_venta_cobrar.Text1(6).Text <> "" Then
        datencabezado.Recordset.Fields("cliente") = frmnota_venta_cobrar.Text1(6).Text
    End If
    datencabezado.Recordset.Fields("vendedorid") = frmnota_venta_cobrar.DataGrid1.Columns(0).Text
    datencabezado.Recordset.Fields("vendedor") = frmnota_venta_cobrar.DataGrid1.Columns(2).Text
    datencabezado.Recordset.Fields("detalle") = frmnota_venta_cobrar.Text1(5).Text
    datencabezado.Recordset.Fields("nota") = frmnota_venta_cobrar.Text15.Text
    datencabezado.Recordset.Fields("cotizacion") = 1
    datencabezado.Recordset.Fields("listadeprecioid") = frmnota_venta_cobrar.DataCombo1.BoundText
    datencabezado.Recordset.Fields("tipodepagoid") = frmnota_venta_cobrar.DataCombo3.BoundText
    datencabezado.Recordset.Fields("alquiler") = "N"
    datencabezado.Recordset.Fields("nroorden") = frmnota_venta_cobrar.Text1(4).Text
    If frmnota_venta_cobrar.Text20.Text <> "" Then
        datencabezado.Recordset.Fields("recetaid") = frmnota_venta_cobrar.Text20.Text
    End If
    
    If frmnota_venta_cobrar.tipofac <> "NN" Then
        datencabezado.Recordset.Fields("tipodefacturacionid") = frmnota_venta_cobrar.datparametros.Recordset.Fields("facdefecto")
    Else
        datencabezado.Recordset.Fields("tipodefacturacionid") = frmnota_venta_cobrar.tipofac
    End If
    
    If Left(login.nombrebd, 14) = "MMOSSE" And frmnota_venta_cobrar.tipofac <> "NN" Then
        testpos = InStr(1, frmnota_venta_cobrar.DataCombo3.Text, "- ", 1)
        tpago = Right(frmnota_venta_cobrar.DataCombo3.Text, Len(frmnota_venta_cobrar.DataCombo3.Text) - testpos - 1)
        If tpago = "CONTADO" Then
            datencabezado.Recordset.Fields("tipodefacturacionid") = frmnota_venta_cobrar.datparametros.Recordset.Fields("facdefecto")
        Else
            datencabezado.Recordset.Fields("tipodefacturacionid") = "MA"
        End If
    End If
    
    datencabezado.Recordset.Fields("fechadeentrega") = DateValue(frmnota_venta_cobrar.Text3.Text) + TimeValue(Str(Time))
    If frmnota_venta_cobrar.Text13.Text = "" Then frmnota_venta_cobrar.Text13.Text = 0
    datencabezado.Recordset.Fields("recargo") = Round(frmnota_venta_cobrar.Text13.Text, 2)
    datencabezado.Recordset.Fields("tiporecargo") = "$"
    If frmnota_venta_cobrar.Text11.Text = "" Then frmnota_venta_cobrar.Text11.Text = 0
    datencabezado.Recordset.Fields("bonificacion") = Round(frmnota_venta_cobrar.Text11.Text, 2)
    datencabezado.Recordset.Fields("tipobonificacion") = "$"
    datencabezado.Recordset.Fields("importeglobal") = Round(frmnota_venta_cobrar.Text4.Text, 2)
    datencabezado.Recordset.Fields("domicilioid") = frmnota_venta_cobrar.Text1(3).Text
    If frmnota_venta_cobrar.DataGrid2.Columns("domicilio_id").Text <> "" Then
        datencabezado.Recordset.Fields("domicilio_id") = frmnota_venta_cobrar.DataGrid2.Columns("domicilio_id").Text
    End If
    datencabezado.Recordset.Fields("domiciliodeentregaid") = frmnota_venta_cobrar.DataGrid2.Columns("domicilio_id").Text
    datencabezado.Recordset.Fields("subtotalsiniva") = Round(frmnota_venta_cobrar.Text5.Text, 2)
    datencabezado.Recordset.Fields("totaliva") = Round(frmnota_venta_cobrar.Text6.Text, 2) + Round(frmnota_venta_cobrar.Text7.Text, 2)
    datencabezado.Recordset.Fields("generada") = "True"
    datencabezado.Recordset.Fields("importado") = "False"
    datencabezado.Recordset.Fields("nombrepc") = Environ("computername")
    datencabezado.Recordset.Fields("target") = frmnota_venta_cobrar.datparametros.Recordset.Fields("sucursal")
    If frmnota_venta_cobrar.DataGrid2.Columns(16).Text = "RI" Then xresponsabilidad = "Resp. Inscripto"
    If frmnota_venta_cobrar.DataGrid2.Columns(16).Text = "EX" Then xresponsabilidad = "Exento"
    If frmnota_venta_cobrar.DataGrid2.Columns(16).Text = "MT" Then xresponsabilidad = "Monotributista"
    If frmnota_venta_cobrar.DataGrid2.Columns(16).Text = "CF" Then xresponsabilidad = "Consumidor Final"
    
    'If frmnota_venta_cobrar.DataGrid2.Columns(16).Text <> "CF" Then datencabezado.Recordset.Fields("tipodefacturacionid") = "MA"
    
    datencabezado.Recordset.Fields("responsabilidad") = xresponsabilidad
    datencabezado.Recordset.Fields("transferido") = "False"
    
    datencabezado.Recordset.Fields("tipodefactura") = frmnota_venta_cobrar.Text1(2).Text
    datencabezado.Recordset.Fields("percepiibb") = Round(frmnota_venta_cobrar.Text9.Text, 2)
    datencabezado.Recordset.Fields("perceptem") = Round(frmnota_venta_cobrar.Text8.Text, 2)
    datencabezado.Recordset.Fields("totaltr") = Round(frmnota_venta_cobrar.Text4.Text, 2)
    datencabezado.Recordset.Fields("presupuestobase") = frmnota_venta_cobrar.presupuestobase
    nromanual = frmnota_venta_cobrar.Text1(7).Text
    puntomanual = frmnota_venta_cobrar.datparametros.Recordset.Fields("puntovtamanual")
    If nromanual <> "" Then datencabezado.Recordset.Fields("tipodefacturacionid") = "MA"
    
    datencabezado.Recordset.UpdateBatch adAffectCurrent
    xid = datencabezado.Recordset.Fields("id")
    
    '** Establene numero de Facturas Manuales, y no Fiscales
    If datencabezado.Recordset.Fields("tipodefacturacionid") = "MA" Or frmnota_venta_cobrar.Text1(2).Text = "A" Then  ' Solo para Colon, para otra empresa verificar
      If frmnota_venta_cobrar.Text1(2).Text = "A" Then
            xnumerador = "Factura A (Vtas) " + frmnota_venta_cobrar.datparametros.Recordset.Fields("sucursal")
      Else
            xnumerador = "Factura B (Vtas) " + frmnota_venta_cobrar.datparametros.Recordset.Fields("sucursal")
      End If
    End If
    If datencabezado.Recordset.Fields("tipodefacturacionid") = "NN" Then
            xnumerador = "Factura de Venta Val " + frmnota_venta_cobrar.datparametros.Recordset.Fields("sucursal")
    End If
    datcola.RecordSource = "SELECT p.ID AS ID, p.NOMBRE AS puntero, n.NUMERO AS NUMERO, p.CARACTERESPREFIJO AS puntoventa, p.NUMERO_ID " & _
                           "FROM V_NUMERADOR_ AS p LEFT OUTER JOIN V_NUMERO_ AS n ON p.NUMERO_ID = n.ID " & _
                           "WHERE     (p.ACTIVESTATUS <> 2) AND (p.NOMBRE = '" & xnumerador & "') "
    datcola.Refresh
    
    If datencabezado.Recordset.Fields("tipodefacturacionid") = "CF" And frmnota_venta_cobrar.Text1(2).Text = "B" Then
        datencabezado.Recordset.Fields("numerodefactura") = Replace(datencabezado.Recordset.Fields("tipodefacturacionid") + Str(xid) + "-" + Str(xclaveprimaria), " ", "")
    Else
       If nromanual = "" Then
        datencabezado.Recordset.Fields("numerodefactura") = datcola.Recordset.Fields("numero")
        datencabezado.Recordset.Fields("puntodeventa") = datcola.Recordset.Fields("puntoventa")
        xnumero = datcola.Recordset.Fields("numero")
        xidnumero = datcola.Recordset.Fields("numero_id")

        datcola.RecordSource = "Select * from numero with(readpast) where id = '" & xidnumero & "'"
        datcola.Refresh
        datcola.Recordset.Fields("numero") = xnumero + 1
        datcola.Recordset.UpdateBatch adAffectCurrent
       Else
        datencabezado.Recordset.Fields("puntodeventa") = puntomanual
        datencabezado.Recordset.Fields("numerodefactura") = nromanual
       End If
       


    End If
    '** Fin de asignacion de numero a Factura

    datencabezado.Recordset.Fields("claveprimaria") = xid
    datencabezado.Recordset.UpdateBatch adAffectCurrent
    
    
'--- Graba Items
    
    For X = 1 To frmnota_venta_cobrar.xlineasmax
        If frmnota_venta_cobrar.grilla.TextMatrix(1, 0) = "" Then
          mensa = MsgBox("No puede Grabar esta Venta sin Items", vbCritical, "Error")
          Exit Sub
        End If
        If frmnota_venta_cobrar.grilla.TextMatrix(X, 0) = "" Then Exit For
        
        datitems.Recordset.AddNew
        datitems.Recordset.Fields("claveprimaria") = xid
        datitems.Recordset.Fields("idproducto") = frmnota_venta_cobrar.grilla.TextMatrix(X, 0)
        datitems.Recordset.Fields("codigoproducto") = frmnota_venta_cobrar.grilla.TextMatrix(X, 1)
        datitems.Recordset.Fields("nombre_producto") = frmnota_venta_cobrar.grilla.TextMatrix(X, 2)
        datitems.Recordset.Fields("cantidadproducto") = frmnota_venta_cobrar.grilla.TextMatrix(X, 3)
        datitems.Recordset.Fields("unidaddemedidaid") = frmnota_venta_cobrar.grilla.TextMatrix(X, 4)
        datitems.Recordset.Fields("preciou") = Round(frmnota_venta_cobrar.grilla.TextMatrix(X, 6), 3)
        datitems.Recordset.Fields("preciousiva") = Round(frmnota_venta_cobrar.grilla.TextMatrix(X, 5), 3)
        datitems.Recordset.Fields("bonificacionitem") = frmnota_venta_cobrar.grilla.TextMatrix(X, 9)
        If frmnota_venta_cobrar.grilla.TextMatrix(X, 8) = "" Then frmnota_venta_cobrar.grilla.TextMatrix(X, 8) = 0
        datitems.Recordset.Fields("importedebonificacion") = Round(frmnota_venta_cobrar.grilla.TextMatrix(X, 8), 4)
        datitems.Recordset.Fields("subtotal") = Round(frmnota_venta_cobrar.grilla.TextMatrix(X, 10), 3)
        datitems.Recordset.Fields("iva") = (Round(frmnota_venta_cobrar.grilla.TextMatrix(X, 12), 4) - 1) * 100
        datitems.Recordset.Fields("listaid") = frmnota_venta_cobrar.grilla.TextMatrix(X, 18)
        
        datitems.Recordset.UpdateBatch adAffectCurrent
    Next X
    


    datcontrol.RecordSource = "Select id, generada from ud_ezi_puntodeventa_encabezado with (readpast) where id = '" & frmnota_venta_cobrar.presupuestobase & "'"
    datcontrol.Refresh
    If datcontrol.Recordset.EOF = False Then
        datcontrol.Recordset.Fields("generada") = "True"
        datcontrol.Recordset.UpdateBatch adAffectCurrent
    End If

'******* Graba Pago

    datpago.RecordSource = "Select * from ud_ezi_pago where claveprimaria = ''"
    datpago.Refresh
    For X = 1 To 29
        If grillac.TextMatrix(X, 0) = "" Then Exit For
        If grillac.TextMatrix(X, 3) = "" Then grillac.TextMatrix(X, 3) = 0
        If Round(grillac.TextMatrix(X, 3), 2) = 0 Then Exit For
        datpago.Recordset.AddNew
        datpago.Recordset.Fields("claveprimaria") = xid
        datpago.Recordset.Fields("tipovalor") = "True"
        datpago.Recordset.Fields("valoroseniaid") = grillac.TextMatrix(X, 0)
        datpago.Recordset.Fields("destinoid") = frmnota_venta_cobrar.datparametros.Recordset.Fields("cajadefecto")
        datpago.Recordset.Fields("formadepago") = grillac.TextMatrix(X, 1)
        datpago.Recordset.Fields("monto") = Round(grillac.TextMatrix(X, 3), 2)
        If datencabezado.Recordset.Fields("tipodefacturacionid") = "NN" Then
            datpago.Recordset.Fields("CAJA") = 2
        Else
            datpago.Recordset.Fields("CAJA") = 1
        End If
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
        datpago.Recordset.Fields("sucursal") = login.nomsucursal
        
        datpago.Recordset.UpdateBatch adAffectCurrent
     Next X


'******* Graba ud_ezi_cola

    
    If datencabezado.Recordset.Fields("tipodefacturacionid") = "CF" Then  ' And frmnota_venta_cobrar.Text1(2).Text = "B" Then
        datcola.RecordSource = "select * from ud_ezi_cola where nombrepc = '1'"
        datcola.Refresh
        
        datcola.Recordset.AddNew
        datcola.Recordset.Fields("nombrepc") = Environ("computername")
        datcola.Recordset.Fields("numero") = datencabezado.Recordset.Fields("numerodefactura")
        datcola.Recordset.Fields("accion") = datencabezado.Recordset.Fields("tipodefactura")
        datcola.Recordset.Fields("target") = frmnota_venta_cobrar.datparametros.Recordset.Fields("target")
        datcola.Recordset.Fields("claveprimaria") = xid
    
        datcola.Recordset.UpdateBatch adAffectCurrent
        
            frmnota_venta_cobrar.cancelar.SetFocus
            SendKeys "{ENTER}", False
            Unload Me
    
    Else
        datcola.RecordSource = "Select * from ud_ezi_cola_importar where id = 0"
        datcola.Refresh
        
        datcola.Recordset.AddNew
        datcola.Recordset.Fields("id_encabezado") = xid
        datcola.Recordset.Fields("tipodedocumentoid") = frmnota_venta_cobrar.datparametros.Recordset.Fields("idfaccdo")
        datcola.Recordset.Fields("unidadoperativaid") = frmnota_venta_cobrar.datparametros.Recordset.Fields("target")
        datcola.Recordset.Fields("fecha_hora") = DateValue(frmnota_venta_cobrar.Text3.Text) + TimeValue(Str(Time))
        
        datcola.Recordset.UpdateBatch adAffectCurrent
        Call imprimefactura_Click
    End If



''' Graba Remito
'    If DataCombo3.Text <> "CONTADO" Then
'        Call grremito_Click
'    End If


 End If
 
Exit Sub
errorgrabar:
    mensa = MsgBox("No se pudo registrar la información", vbCritical, "Error !!")


End Sub


Private Sub grillac_KeyUp(KeyCode As Integer, Shift As Integer)

    
If KeyCode = 46 Then
        For X = 0 To 3
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


Private Sub imprimefactura_Click()
On Error GoTo fuera
Dim crApp As New CRAXDRT.Application
Dim crReport As New CRAXDRT.Report
Dim tabla As String
Dim ruta As String

If frmnota_venta_cobrar.datparametros.Recordset.Fields("imprimemanual") = "N" Or frmnota_venta_cobrar.Text1(7).Text <> "" Then
    frmnota_venta_cobrar.cancelar.SetFocus
    SendKeys "{ENTER}", False
    Unload Me
    Exit Sub
End If

xbase = Left(login.nombrebd, 17)
                   
    reporte.SQL = "SELECT v_ezi_pos_factctacte.id, v_ezi_pos_factctacte.NUMERODOCUMENTO, v_ezi_pos_factctacte.FECHAEMISION, v_ezi_pos_factctacte.cod_cliente, v_ezi_pos_factctacte.cliente, v_ezi_pos_factctacte.CUIT, v_ezi_pos_factctacte.CALLE, v_ezi_pos_factctacte.CODPOS, v_ezi_pos_factctacte.provincia, v_ezi_pos_factctacte.detalle, v_ezi_pos_factctacte.tipopago, v_ezi_pos_factctacte.codigoproducto, v_ezi_pos_factctacte.nombre_producto, v_ezi_pos_factctacte.cantidadproducto, v_ezi_pos_factctacte.nota, v_ezi_pos_factctacte.condiva, v_ezi_pos_factctacte.ciudad, v_ezi_pos_factctacte.TIPOVENTA, v_ezi_pos_factctacte.SIMBOLO, v_ezi_pos_factctacte.CODVENDEDOR, v_ezi_pos_factctacte.preciusiniva, v_ezi_pos_factctacte.subtotalsiniva, v_ezi_pos_factctacte.impbonifsiniva, v_ezi_pos_factctacte.percepiibb, v_ezi_pos_factctacte.perceptem, v_ezi_pos_factctacte.totaltr, v_ezi_pos_factctacte.importeiva21, v_ezi_pos_factctacte.importeiva105 FROM dbo.v_ezi_pos_factctacte v_ezi_pos_factctacte " & _
                  "where v_ezi_pos_factctacte.id = " & xid & " order by v_ezi_pos_factctacte.iditem"

  '  reporte.SQL = "SELECT v_ezi_pos_factctacte.id, v_ezi_pos_factctacte.NUMERODOCUMENTO, v_ezi_pos_factctacte.FECHAEMISION, v_ezi_pos_factctacte.cod_cliente, v_ezi_pos_factctacte.cliente, v_ezi_pos_factctacte.CUIT, v_ezi_pos_factctacte.CALLE, v_ezi_pos_factctacte.CODPOS, v_ezi_pos_factctacte.provincia, v_ezi_pos_factctacte.detalle, v_ezi_pos_factctacte.tipopago, v_ezi_pos_factctacte.codigoproducto, v_ezi_pos_factctacte.nombre_producto, v_ezi_pos_factctacte.cantidadproducto, v_ezi_pos_factctacte.nota, v_ezi_pos_factctacte.condiva, v_ezi_pos_factctacte.ciudad, v_ezi_pos_factctacte.TIPOVENTA, v_ezi_pos_factctacte.SIMBOLO, v_ezi_pos_factctacte.CODVENDEDOR, v_ezi_pos_factctacte.preciusiniva, v_ezi_pos_factctacte.subtotalsiniva, v_ezi_pos_factctacte.impbonifsiniva, v_ezi_pos_factctacte.nroremito, v_ezi_pos_factctacte.percepiibb, v_ezi_pos_factctacte.perceptem, v_ezi_pos_factctacte.totaltr, v_ezi_pos_factctacte.importeiva21, v_ezi_pos_factctacte.importeiva105, v_ezi_pos_factctacte.iditem " & _
  '            "FROM MMOSSE.dbo.v_ezi_pos_factctacte v_ezi_pos_factctacte " & _
  '            "where v_ezi_pos_factctacte.id = " & xid & " order by v_ezi_pos_factctacte.iditem"

tabla = reporte.SQL
Debug.Print reporte.SQL

With CrystalReporte
    .PrinterCollation = crptCollated
    If frmnota_venta_cobrar.tipofac <> "NN" Then
        .Formulas(0) = "copia="" ORIGINAL """
    End If
    If frmnota_venta_cobrar.Text1(2).Text = "A" Then
   '   If frmnota_venta_cobrar.tipofac <> "NN" Then
   '     .ReportFileName = App.Path & "\FacturaCtaCteA.rpt"
   '   Else
        .ReportFileName = App.Path & "\PresupuestoA.rpt"
    '  End If
    Else
     ' If frmnota_venta_cobrar.tipofac <> "NN" Then
    '    .ReportFileName = App.Path & "\FacturaCtaCteB.rpt"
    '  Else
        .ReportFileName = App.Path & "\PresupuestoB.rpt"
    '  End If
    End If
    .WindowTitle = "Factura Vta Orig"
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
'    If frmnota_venta_cobrar.tipofac <> "NN" Then
'    .WindowTitle = "Factura Vta Dupl"
'    .Formulas(0) = "copia="" DUPLICADO """
'    .Action = 1
'    If frmnota_venta_cobrar.Text1(2).Text = "A" Then
'    .WindowTitle = "Factura Vta Trip"
'    .Formulas(0) = "copia="" TRIPLICADO """
'    .Action = 1
'    End If
'    End If
    
End With

    frmnota_venta_cobrar.cancelar.SetFocus
    SendKeys "{ENTER}", False
    Unload Me
    

Exit Sub

fuera:
    
    frmnota_venta_cobrar.cancelar.SetFocus
    SendKeys "{ENTER}", False
    Unload Me
    'MsgBox "Reporte de Factura no Encontado, o error de configuracion de reporte", vbCritical, "Error"


End Sub

Private Sub Text1_GotFocus(Index As Integer)

    If Index = 1 Then
        Text1(1).SelStart = 0
        Text1(1).SelLength = Len(Text1(1).Text)
    End If

End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
On Error Resume Next

    If KeyAscii = 13 Then
       KeyAscii = 0
       If Index = 1 Then
        For X = 1 To 30
          If grillac.TextMatrix(X, 1) = "" Then
'*********************** Efectivo
           If DataCombo1.Text = "Efectivo" Then
                grillac.TextMatrix(X, 0) = DataCombo1.BoundText
                grillac.TextMatrix(X, 1) = DataCombo1.Text
                grillac.TextMatrix(X, 3) = Format(Text1(1).Text, "###,##0.00")
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
        Text10.Text = "N"
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
            MsgBox "Debe Ingresar una Tarjeta Válida", vbCritical, "Error"
            DataCombo2.SetFocus
            Exit Sub
        End If
    
        If Text2.Text = "" Then
            MsgBox "Debe Ingresar un Numero de Cuotas", vbCritical, "Error"
            Text2.SetFocus
            Exit Sub
        End If
        
        If Text3.Text = "" Then
            MsgBox "Debe Ingresar un Numero de Cupón", vbCritical, "Error"
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
                Exit For
          End If
        Next X
        
        tarjeta.Visible = False
        Call calcula_Click

    End If


End Sub

Private Sub Text7_Change()

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
    
        If Text9.Text = "" Then
            MsgBox "Debe Ingresar un Numero de Cheque", vbCritical, "Error"
            Text9.SetFocus
            Exit Sub
        End If
        
        If fechavencimiento.Value < femicioncheque.Value Then
            MsgBox "La fecha de vencimiento no puede ser menor a la fecha de Emisión"
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
                Exit For
          End If
        Next X
        
        cheques.Visible = False
        Call calcula_Click

    End If

End Sub

Private Sub Text6_LostFocus()
On Error Resume Next

    

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
