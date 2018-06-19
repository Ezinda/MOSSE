VERSION 5.00
Object = "{F6125AB1-8AB1-11CE-A77F-08002B2F4E98}#2.0#0"; "MSRDC20.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{D18BBD1F-82BB-4385-BED3-E9D31A3E361E}#1.0#0"; "KewlButtonz.ocx"
Begin VB.Form frmingreovalor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ingreso de Valor"
   ClientHeight    =   6870
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14280
   Icon            =   "frmingresovalor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6870
   ScaleWidth      =   14280
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
      TabIndex        =   48
      Top             =   0
      Visible         =   0   'False
      Width           =   5415
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
         TabIndex        =   52
         Top             =   960
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
         TabIndex        =   54
         Top             =   960
         Width           =   2055
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
         TabIndex        =   51
         Top             =   360
         Width           =   1695
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
         TabIndex        =   50
         Top             =   1560
         Width           =   1695
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
         TabIndex        =   55
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
         TabIndex        =   49
         Top             =   1560
         Width           =   615
      End
      Begin MSComCtl2.DTPicker femisionret 
         Height          =   375
         Left            =   2040
         TabIndex        =   53
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
         Format          =   136052737
         CurrentDate     =   41988
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
      TabIndex        =   27
      Top             =   0
      Visible         =   0   'False
      Width           =   5415
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
         TabIndex        =   39
         Top             =   840
         Width           =   1455
      End
      Begin MSComCtl2.DTPicker femicioncheque 
         Height          =   375
         Left            =   1800
         TabIndex        =   28
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
         Format          =   136052737
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
         TabIndex        =   38
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
         TabIndex        =   37
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
         TabIndex        =   31
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
         TabIndex        =   36
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
         TabIndex        =   32
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
         TabIndex        =   35
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
         TabIndex        =   33
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
         TabIndex        =   34
         Top             =   360
         Width           =   1455
      End
      Begin MSDataListLib.DataCombo DataCombo5 
         Bindings        =   "frmingresovalor.frx":0442
         Height          =   420
         Left            =   240
         TabIndex        =   30
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
         TabIndex        =   29
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
         Format          =   173670401
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
      Left            =   8760
      TabIndex        =   11
      Top             =   0
      Visible         =   0   'False
      Width           =   5415
      Begin MSComCtl2.DTPicker femision 
         Height          =   375
         Left            =   1800
         TabIndex        =   25
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
         Format          =   173670401
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
         TabIndex        =   24
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
         TabIndex        =   23
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
         TabIndex        =   22
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
         TabIndex        =   21
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
         TabIndex        =   20
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
         TabIndex        =   19
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
         TabIndex        =   18
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
         TabIndex        =   17
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
         TabIndex        =   15
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
         TabIndex        =   14
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
         TabIndex        =   12
         Top             =   360
         Width           =   1215
      End
      Begin MSDataListLib.DataCombo DataCombo2 
         Bindings        =   "frmingresovalor.frx":0459
         Height          =   420
         Left            =   240
         TabIndex        =   13
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
         Bindings        =   "frmingresovalor.frx":0473
         Height          =   420
         Left            =   240
         TabIndex        =   16
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
      Left            =   3600
      Top             =   6600
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
      Left            =   3600
      Top             =   6240
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
      Left            =   2400
      Top             =   6240
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
      Left            =   2280
      Top             =   6480
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
      Height          =   5295
      Left            =   120
      ScaleHeight     =   5235
      ScaleWidth      =   8475
      TabIndex        =   7
      Top             =   120
      Width           =   8535
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
         Left            =   2400
         MaxLength       =   50
         TabIndex        =   47
         Top             =   1080
         Width           =   5895
      End
      Begin MSComCtl2.DTPicker fecha 
         Height          =   375
         Left            =   6240
         TabIndex        =   45
         Top             =   2040
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   135397377
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
         Left            =   2400
         MaxLength       =   50
         TabIndex        =   42
         Top             =   2040
         Width           =   1935
      End
      Begin VB.CommandButton CALCULA 
         Caption         =   "CALCULA"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   2400
         Visible         =   0   'False
         Width           =   1095
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grillac 
         Height          =   2175
         Left            =   120
         TabIndex        =   4
         Top             =   2880
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   3836
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
         Left            =   2400
         MaxLength       =   50
         TabIndex        =   3
         Top             =   1560
         Width           =   1935
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "frmingresovalor.frx":048A
         Height          =   420
         Left            =   2400
         TabIndex        =   2
         Top             =   600
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
         Bindings        =   "frmingresovalor.frx":04A3
         Height          =   420
         Left            =   2400
         TabIndex        =   41
         Top             =   120
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   741
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         ListField       =   "nombre"
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
         TabIndex        =   46
         Top             =   1080
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
         Left            =   5040
         TabIndex        =   44
         Top             =   2040
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
         TabIndex        =   43
         Top             =   2040
         Width           =   1935
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Motivo:"
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
         Index           =   2
         Left            =   120
         TabIndex        =   40
         Top             =   120
         Width           =   2055
      End
      Begin VB.Label Label2 
         Caption         =   "Tecla Supr , Borra Ingreso"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   2640
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
         TabIndex        =   9
         Top             =   2040
         Width           =   1935
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   8280
         Y1              =   2520
         Y2              =   2520
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
         Top             =   600
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
         TabIndex        =   8
         Top             =   1560
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
      Top             =   5520
      Width           =   8535
      Begin KewlButtonz.KewlButtons grabar 
         Height          =   615
         Left            =   720
         TabIndex        =   5
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
         MICON           =   "frmingresovalor.frx":04C1
         PICN            =   "frmingresovalor.frx":04DD
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
         TabIndex        =   6
         Top             =   360
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
         MICON           =   "frmingresovalor.frx":1F5F
         PICN            =   "frmingresovalor.frx":1F7B
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
      Begin MSAdodcLib.Adodc datimputaciones 
         Height          =   330
         Left            =   5160
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
   Begin MSAdodcLib.Adodc datparametros 
      Height          =   330
      Left            =   11400
      Top             =   5640
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
Attribute VB_Name = "frmingreovalor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public xid As Double


Private Sub calcula_Click()

xpagos = 0
For X = 1 To 29
  If grillac.TextMatrix(X, 3) = "" Then
     xvalorgrilla = 0
  Else
     xvalorgrilla = Round(grillac.TextMatrix(X, 3), 10)
  End If
  xpagos = xpagos + xvalorgrilla
Next X
    
    'Text1(1).Text = Format(Round(Text1(0).Text, 10) - xpagos, "###,##0.00")
    Text1(2).Text = Format(xpagos, "###,##0.00")
    
'    If Round(Text1(1).Text, 2) = 0 Then
'        grabar.SetFocus
'        grabar_Click
'    Else
        Text1(1).Text = Format(0, "###,##0.00")
        DataCombo4.SetFocus
'    End If
    
    

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
DataCombo5.Text = ""
Text9.Text = ""
Text8.Text = ""
Text6.Text = ""



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


Private Sub DataCombo4_GotFocus()

    frmingreovalor.Width = 8790

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

Private Sub femicioncheque_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{TAB}", False
    End If
    

End Sub

Private Sub Form_Activate()
Call calcula_Click

DataCombo4.SetFocus

End Sub

Private Sub Form_Load()
Aplicar_skin Me

yventana = Inicio.Height / 2 - 1000
xventana = Inicio.Width / 2

'frmingreovalor.Top = yventana - frmingreovalor.Height / 2
'frmingreovalor.Left = xventana - frmingreovalor.Width / 2
frmingreovalor.Top = 0
frmingreovalor.Left = 0
frmingreovalor.Width = 8790
fecha = Date


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


   
End Sub

Private Sub Frame1_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub grabar_Click()
'On Error GoTo errorgrabar

    If Round(Text1(2).Text, 2) = 0 Then
        mensa = MsgBox("No se puede Grabar el Comprobante en Cero", vbCritical, "!! Error !!")
        Exit Sub
    End If
    
 mensa = MsgBox("Desea Grabar este Ingreso de Valor ?", vbYesNo, "!! Atención !!")
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
    datencabezado.Recordset.Fields("numeradorinterno") = "Ingreso de Valor"
    datencabezado.Recordset.Fields("fechadelcomprobante") = fecha.Value
    datencabezado.Recordset.Fields("detalle") = Text7.Text
    datencabezado.Recordset.Fields("sucursal") = datparametros.Recordset.Fields("sucursal")
    datencabezado.Recordset.Fields("cotizacion") = 1
    datencabezado.Recordset.Fields("alquiler") = "N"
    datencabezado.Recordset.Fields("importeglobal") = Round(Text1(2).Text, 2)
    datencabezado.Recordset.Fields("generada") = "True"
    datencabezado.Recordset.Fields("importado") = "False"
    datencabezado.Recordset.Fields("clienteid") = DataCombo4.BoundText
    datencabezado.Recordset.Fields("vendedorid") = datparametros.Recordset("cajadefecto")
    datencabezado.Recordset.Fields("nombrepc") = Environ("computername")
    datencabezado.Recordset.Fields("target") = datparametros.Recordset.Fields("sucursal")
    datencabezado.Recordset.Fields("transferido") = "False"
    datencabezado.Recordset.Fields("totaltr") = Round(Text1(2).Text, 2)
    
    datencabezado.Recordset.UpdateBatch adAffectCurrent
    xid = datencabezado.Recordset.Fields("id")
    
    '** Establene numero de Facturas Manuales, y no Fiscales
    xnumerador = "Ingreso de Valores Cajas/Bancos"
    datcola.RecordSource = "SELECT p.ID AS ID, p.NOMBRE AS puntero, n.NUMERO AS NUMERO, p.CARACTERESPREFIJO AS puntoventa, p.NUMERO_ID " & _
                           "FROM V_NUMERADOR_ AS p LEFT OUTER JOIN V_NUMERO_ AS n ON p.NUMERO_ID = n.ID " & _
                           "WHERE     (p.ACTIVESTATUS <> 2) AND (p.NOMBRE = '" & xnumerador & "') "
    datcola.Refresh
    datencabezado.Recordset.Fields("numerodefactura") = datcola.Recordset.Fields("numero")
       
        xnumero = datcola.Recordset.Fields("numero")
        xidnumero = datcola.Recordset.Fields("numero_id")

        datcola.RecordSource = "Select * from numero with(readpast) where id = '" & xidnumero & "'"
        datcola.Refresh
        datcola.Recordset.Fields("numero") = xnumero + 1
        datcola.Recordset.UpdateBatch adAffectCurrent


    '** Fin de asignacion de numero a Factura

    datencabezado.Recordset.Fields("claveprimaria") = xid
    datencabezado.Recordset.UpdateBatch adAffectCurrent
    
    
'******* Graba Pago

    datpago.RecordSource = "Select * from ud_ezi_pago where claveprimaria = ''"
    datpago.Refresh
    For X = 1 To 29
        If grillac.TextMatrix(X, 0) = "" Then Exit For
        datpago.Recordset.AddNew
        datpago.Recordset.Fields("claveprimaria") = xid
        datpago.Recordset.Fields("tipovalor") = "True"
        datpago.Recordset.Fields("valoroseniaid") = grillac.TextMatrix(X, 0)
        datpago.Recordset.Fields("destinoid") = datparametros.Recordset.Fields("cajadefecto")
        datpago.Recordset.Fields("formadepago") = grillac.TextMatrix(X, 1)
        datpago.Recordset.Fields("monto") = Round(grillac.TextMatrix(X, 3), 2)
        datpago.Recordset.Fields("idimputacion") = grillac.TextMatrix(X, 16)
        datpago.Recordset.Fields("imputacion") = grillac.TextMatrix(X, 15)
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
        datpago.Recordset.Fields("sucursal") = login.nomsucursal

        datpago.Recordset.UpdateBatch adAffectCurrent
     Next X


'******* Graba ud_ezi_cola

    
        datcola.RecordSource = "Select * from ud_ezi_cola_importar where id = 0"
        datcola.Refresh
        
        datcola.Recordset.AddNew
        datcola.Recordset.Fields("id_encabezado") = xid
        datcola.Recordset.Fields("tipodedocumentoid") = frmnota_venta.datparametros.Recordset.Fields("idingresovalor")
        datcola.Recordset.Fields("unidadoperativaid") = frmnota_venta.datparametros.Recordset.Fields("target")
        datcola.Recordset.Fields("fecha_hora") = Date
        
        datcola.Recordset.UpdateBatch adAffectCurrent
        
        Unload Me
        frmingreovalor.Show

 End If
 
Exit Sub
errorgrabar:
    mensa = MsgBox("No se pudo registrar la información", vbCritical, "Error !!")


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
           If X = 2 Then
                MsgBox "No puede ingresar mas de un item en el ingreso", vbCritical, "Error"
                Exit For
           End If
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
            frmingreovalor.Width = 14370
            tarjeta.Visible = True
            DataCombo2.SetFocus
           End If
'*********************** Cheque
           If Left(DataCombo1.Text, 6) = "Cheque" Then
            frmingreovalor.Width = 14370
            cheques.Visible = True
            femicioncheque.SetFocus
           End If
'*********************** retenciones
           If Left(DataCombo1.Text, 3) = "Ret" Then
            frmingreovalor.Width = 14370
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
                grillac.TextMatrix(X, 15) = DataCombo4.Text
                grillac.TextMatrix(X, 16) = DataCombo4.BoundText
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
    
        If Text9.Text = "" Then
            MsgBox "Debe Ingresar un Numero de Cheque", vbCritical, "Error"
            Text9.SetFocus
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
