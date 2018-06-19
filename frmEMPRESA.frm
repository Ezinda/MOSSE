VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{5734474E-78D3-4254-99B9-C35F31BDF509}#62.0#0"; "vbskpro2.ocx"
Begin VB.Form frmEMPRESA 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Empresas"
   ClientHeight    =   7200
   ClientLeft      =   1095
   ClientTop       =   330
   ClientWidth     =   8220
   Icon            =   "frmEMPRESA.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   NegotiateMenus  =   0   'False
   ScaleHeight     =   7200
   ScaleWidth      =   8220
   Begin MSAdodcLib.Adodc datpuntos 
      Height          =   330
      Left            =   4800
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
   Begin TabDlg.SSTab SSTab1 
      Height          =   4815
      Left            =   240
      TabIndex        =   7
      Top             =   240
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   8493
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "Datos"
      TabPicture(0)   =   "frmEMPRESA.frx":0442
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "lblLabels(10)"
      Tab(0).Control(1)=   "lblLabels(9)"
      Tab(0).Control(2)=   "lblLabels(8)"
      Tab(0).Control(3)=   "lblLabels(7)"
      Tab(0).Control(4)=   "lblLabels(6)"
      Tab(0).Control(5)=   "lblLabels(5)"
      Tab(0).Control(6)=   "lblLabels(4)"
      Tab(0).Control(7)=   "lblLabels(3)"
      Tab(0).Control(8)=   "lblLabels(2)"
      Tab(0).Control(9)=   "lblLabels(1)"
      Tab(0).Control(10)=   "lblLabels(0)"
      Tab(0).Control(11)=   "datacontrib"
      Tab(0).Control(12)=   "Frame1"
      Tab(0).Control(13)=   "DataCombo1"
      Tab(0).Control(14)=   "inicioper"
      Tab(0).Control(15)=   "finper"
      Tab(0).Control(16)=   "Text1(1)"
      Tab(0).Control(17)=   "Text1(0)"
      Tab(0).Control(18)=   "condtrib"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "telefono"
      Tab(0).Control(20)=   "localidad"
      Tab(0).Control(21)=   "domicilio"
      Tab(0).Control(22)=   "razonsocial"
      Tab(0).Control(23)=   "empresa(0)"
      Tab(0).Control(24)=   "cuit"
      Tab(0).Control(25)=   "inicioactividad"
      Tab(0).Control(26)=   "maskinicio"
      Tab(0).Control(27)=   "maskcuit"
      Tab(0).ControlCount=   28
      TabCaption(1)   =   "Avanzados"
      TabPicture(1)   =   "frmEMPRESA.frx":045E
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Frame3"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "DataGrid2"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Command1"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Command2"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Command3"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).ControlCount=   5
      Begin VB.CommandButton Command3 
         Caption         =   "Guardar"
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
         Left            =   4920
         TabIndex        =   39
         Top             =   2640
         Width           =   855
      End
      Begin VB.CommandButton Command2 
         Caption         =   "-"
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
         Left            =   4440
         TabIndex        =   38
         Top             =   2640
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         Caption         =   "+"
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
         Left            =   3960
         TabIndex        =   37
         Top             =   2640
         Width           =   375
      End
      Begin MSDataGridLib.DataGrid DataGrid2 
         Bindings        =   "frmEMPRESA.frx":047A
         Height          =   2055
         Left            =   3720
         TabIndex        =   36
         Top             =   600
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   3625
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
               LCID            =   1034
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
               LCID            =   1034
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
      Begin MSMask.MaskEdBox maskcuit 
         Bindings        =   "frmEMPRESA.frx":0492
         DataField       =   "cuit"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
         DataSource      =   "dataempresa"
         Height          =   285
         Left            =   -72840
         TabIndex        =   19
         Top             =   1440
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         _Version        =   393216
         AllowPrompt     =   -1  'True
         AutoTab         =   -1  'True
         MaxLength       =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##-########-#"
         PromptChar      =   "_"
      End
      Begin MSComCtl2.DTPicker maskinicio 
         DataField       =   "inicioactividad"
         DataSource      =   "dataempresa"
         Height          =   285
         Left            =   -72840
         TabIndex        =   13
         Top             =   1800
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   65011713
         CurrentDate     =   38400
      End
      Begin VB.TextBox inicioactividad 
         DataField       =   "inicioactividad"
         DataSource      =   "dataempresa"
         Height          =   285
         Left            =   -72840
         TabIndex        =   22
         Top             =   1800
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox cuit 
         DataField       =   "cuit"
         DataSource      =   "dataempresa"
         Height          =   285
         Left            =   -72840
         TabIndex        =   21
         Top             =   1440
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox empresa 
         Alignment       =   2  'Center
         DataField       =   "empresa"
         DataSource      =   "dataempresa"
         Height          =   285
         Index           =   0
         Left            =   -72840
         TabIndex        =   20
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox razonsocial 
         DataField       =   "razonsocial"
         DataSource      =   "dataempresa"
         Height          =   285
         Left            =   -72840
         MaxLength       =   50
         TabIndex        =   18
         Top             =   1080
         Width           =   3615
      End
      Begin VB.TextBox domicilio 
         DataField       =   "domicilio"
         DataSource      =   "dataempresa"
         Height          =   285
         Left            =   -72840
         TabIndex        =   17
         Top             =   2160
         Width           =   3615
      End
      Begin VB.TextBox localidad 
         DataField       =   "localidad"
         DataSource      =   "dataempresa"
         Height          =   285
         Left            =   -72840
         TabIndex        =   16
         Top             =   2520
         Width           =   3615
      End
      Begin VB.TextBox telefono 
         DataField       =   "telefono"
         DataSource      =   "dataempresa"
         Height          =   285
         Left            =   -72840
         TabIndex        =   15
         Top             =   2880
         Width           =   3615
      End
      Begin VB.TextBox condtrib 
         Alignment       =   2  'Center
         DataField       =   "condtrib"
         DataSource      =   "dataempresa"
         Height          =   285
         Left            =   -72840
         Locked          =   -1  'True
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   3240
         Width           =   495
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         DataField       =   "inicioperiodo"
         DataSource      =   "dataempresa"
         Height          =   285
         Index           =   0
         Left            =   -72840
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   3600
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         DataField       =   "finperiodo"
         DataSource      =   "dataempresa"
         Height          =   285
         Index           =   1
         Left            =   -72840
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   3960
         Width           =   1335
      End
      Begin MSComCtl2.DTPicker finper 
         Height          =   285
         Left            =   -71520
         TabIndex        =   10
         Top             =   3960
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   503
         _Version        =   393216
         Format          =   65011713
         CurrentDate     =   38427
      End
      Begin MSComCtl2.DTPicker inicioper 
         Height          =   285
         Left            =   -71520
         TabIndex        =   11
         Top             =   3600
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   503
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   65011713
         CurrentDate     =   38427
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "frmEMPRESA.frx":04B6
         DataField       =   "condtrib"
         DataSource      =   "dataempresa"
         Height          =   315
         Left            =   -72360
         TabIndex        =   12
         Top             =   3240
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         BackColor       =   14737632
         ListField       =   "descripcion"
         BoundColumn     =   "categ"
         Text            =   "DataCombo1"
      End
      Begin VB.Frame Frame1 
         Height          =   4215
         Left            =   -72960
         TabIndex        =   23
         Top             =   480
         Width           =   3975
         Begin VB.CheckBox Check1 
            DataField       =   "habcc"
            DataSource      =   "dataempresa"
            Height          =   255
            Left            =   120
            TabIndex        =   24
            Top             =   3840
            Width           =   375
         End
      End
      Begin MSAdodcLib.Adodc datacontrib 
         Height          =   330
         Left            =   -70440
         Top             =   4200
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
      Begin VB.Frame Frame3 
         Caption         =   "Puntos Predeterminados de Facturación"
         Height          =   3375
         Left            =   240
         TabIndex        =   40
         Top             =   480
         Width           =   3375
         Begin VB.ComboBox Combo1 
            Height          =   315
            Left            =   240
            TabIndex        =   52
            Top             =   2760
            Width           =   1935
         End
         Begin VB.TextBox Text3 
            DataField       =   "modfacturacion"
            DataSource      =   "dataempresa"
            Height          =   285
            Left            =   240
            Locked          =   -1  'True
            TabIndex        =   51
            Top             =   2400
            Width           =   1695
         End
         Begin VB.TextBox Text2 
            Alignment       =   2  'Center
            DataField       =   "pfo"
            DataSource      =   "dataempresa"
            Height          =   285
            Index           =   2
            Left            =   2280
            Locked          =   -1  'True
            TabIndex        =   49
            Top             =   1680
            Width           =   975
         End
         Begin VB.TextBox Text2 
            Alignment       =   2  'Center
            DataField       =   "pfl"
            DataSource      =   "dataempresa"
            Height          =   285
            Index           =   1
            Left            =   2280
            Locked          =   -1  'True
            TabIndex        =   48
            Top             =   1080
            Width           =   975
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Option1"
            Height          =   255
            Index           =   0
            Left            =   2040
            TabIndex        =   44
            Top             =   480
            Width           =   255
         End
         Begin VB.TextBox Text2 
            Alignment       =   2  'Center
            DataField       =   "pfm"
            DataSource      =   "dataempresa"
            Height          =   285
            Index           =   0
            Left            =   2280
            Locked          =   -1  'True
            TabIndex        =   47
            Top             =   480
            Width           =   975
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Option1"
            Height          =   255
            Index           =   2
            Left            =   2040
            TabIndex        =   46
            Top             =   1680
            Width           =   255
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Option1"
            Height          =   255
            Index           =   1
            Left            =   2040
            TabIndex        =   45
            Top             =   1080
            Width           =   255
         End
         Begin VB.Label lblLabels 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Modulos Facturación"
            ForeColor       =   &H80000007&
            Height          =   255
            Index           =   14
            Left            =   240
            TabIndex        =   50
            Top             =   2160
            Width           =   1695
         End
         Begin VB.Label lblLabels 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Otros Modulos:"
            ForeColor       =   &H80000007&
            Height          =   255
            Index           =   13
            Left            =   240
            TabIndex        =   43
            Top             =   1680
            Width           =   1695
         End
         Begin VB.Label lblLabels 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Carga Libro Ventas:"
            ForeColor       =   &H80000007&
            Height          =   255
            Index           =   12
            Left            =   240
            TabIndex        =   42
            Top             =   1080
            Width           =   1695
         End
         Begin VB.Label lblLabels 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Factutacion Manual:"
            ForeColor       =   &H80000007&
            Height          =   255
            Index           =   11
            Left            =   240
            TabIndex        =   41
            Top             =   480
            Width           =   1695
         End
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Codigo de Empresa:"
         ForeColor       =   &H80000007&
         Height          =   255
         Index           =   0
         Left            =   -74880
         TabIndex        =   35
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Razon Social:"
         ForeColor       =   &H80000007&
         Height          =   255
         Index           =   1
         Left            =   -74880
         TabIndex        =   34
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "C..U.I.T.:"
         ForeColor       =   &H80000007&
         Height          =   255
         Index           =   2
         Left            =   -74880
         TabIndex        =   33
         Top             =   1425
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Inicio de Actividad:"
         ForeColor       =   &H80000007&
         Height          =   255
         Index           =   3
         Left            =   -74880
         TabIndex        =   32
         Top             =   1800
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Domicilio:"
         ForeColor       =   &H80000007&
         Height          =   255
         Index           =   4
         Left            =   -74880
         TabIndex        =   31
         Top             =   2160
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Telefonos:"
         ForeColor       =   &H80000007&
         Height          =   255
         Index           =   5
         Left            =   -74880
         TabIndex        =   30
         Top             =   2880
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Condicion Tributaria:"
         ForeColor       =   &H80000007&
         Height          =   255
         Index           =   6
         Left            =   -74880
         TabIndex        =   29
         Top             =   3240
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Localidad:"
         ForeColor       =   &H80000007&
         Height          =   255
         Index           =   7
         Left            =   -74880
         TabIndex        =   28
         Top             =   2520
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Inicio Periodo Contable:"
         ForeColor       =   &H80000007&
         Height          =   255
         Index           =   8
         Left            =   -74880
         TabIndex        =   27
         Top             =   3600
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fin Periodo Contable:"
         ForeColor       =   &H80000007&
         Height          =   255
         Index           =   9
         Left            =   -74880
         TabIndex        =   26
         Top             =   3960
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Habilita Cent.de Costos"
         ForeColor       =   &H80000007&
         Height          =   255
         Index           =   10
         Left            =   -74880
         TabIndex        =   25
         Top             =   4320
         Width           =   1815
      End
   End
   Begin VB.CommandButton salir 
      Caption         =   "&Cerrar"
      Height          =   855
      Left            =   6720
      Picture         =   "frmEMPRESA.frx":04D0
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5040
      UseMaskColor    =   -1  'True
      Width           =   1095
   End
   Begin VB.CommandButton Cancelar 
      Cancel          =   -1  'True
      Caption         =   "Cance&lar"
      Height          =   615
      Left            =   6720
      Picture         =   "frmEMPRESA.frx":0912
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2160
      UseMaskColor    =   -1  'True
      Width           =   1095
   End
   Begin VB.CommandButton nuevo 
      Caption         =   "&Nuevo"
      Height          =   615
      Left            =   6720
      Picture         =   "frmEMPRESA.frx":0E44
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1440
      UseMaskColor    =   -1  'True
      Width           =   1095
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmEMPRESA.frx":1376
      Height          =   1935
      Left            =   240
      TabIndex        =   5
      Top             =   5160
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   3413
      _Version        =   393216
      AllowArrows     =   -1  'True
      BackColor       =   14737632
      HeadLines       =   1
      RowHeight       =   15
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
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   "empresa"
         Caption         =   "Cod.Empresa"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
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
         DataField       =   "razonsocial"
         Caption         =   "Razon Social"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
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
            Locked          =   -1  'True
         EndProperty
         BeginProperty Column01 
            Locked          =   -1  'True
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton grabar 
      Caption         =   "&Grabar"
      Height          =   615
      Left            =   6720
      Picture         =   "frmEMPRESA.frx":1390
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   720
      UseMaskColor    =   -1  'True
      Width           =   1095
   End
   Begin MSAdodcLib.Adodc dataempresa 
      Height          =   330
      Left            =   4800
      Top             =   5280
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
      ForeColor       =   &H00FF0000&
      Height          =   6855
      Left            =   6480
      TabIndex        =   6
      Top             =   240
      Width           =   1575
      Begin VB.CommandButton borrar 
         Caption         =   "&Borrar"
         Height          =   615
         Left            =   240
         Picture         =   "frmEMPRESA.frx":18C2
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   2640
         UseMaskColor    =   -1  'True
         Width           =   1095
      End
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
      LcK2            =   $"frmEMPRESA.frx":19C4
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
End
Attribute VB_Name = "frmEMPRESA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub borrar_Click()
On Error GoTo DeleteErr
  
     If login.empresasbajas = "N" Then
        mensa = MsgBox("Acceso Denegado", , "Sistema")
        Exit Sub
    End If
  
  
  KeyAscii = 13
  Respuesta = MsgBox("ESTA POR BORRAR UNA EMPRESA, ESTA SEGURO?", vbYesNo, "Atención")
If Respuesta = vbYes Then
    
    dataempresa.Recordset.Delete
Else
    Exit Sub
End If

 
 Exit Sub
DeleteErr:
  MsgBox "No se pudo borrar, pulse el boton ´Grabar´ e intente eliminar el registro nuevamente"
  Call Cancelar_Click
End Sub

Private Sub Cancelar_Click()

        dataempresa.Refresh

End Sub


Private Sub Combo1_Click()

    Text3.Text = Combo1.Text
End Sub

Private Sub Combo1_KeyUp(KeyCode As Integer, Shift As Integer)

    Text3.Text = Combo1.Text
End Sub

Private Sub Command1_Click()

    datpuntos.Recordset.AddNew
    datpuntos.Recordset.Fields(1) = login.empresaact
    

End Sub

Private Sub Command2_Click()

    mensa = MsgBox("Esta por borrar un punto de Venta, Esta Seguro ?", vbYesNo, "!! Atencion ")
    If mensa = vbYes Then datpuntos.Recordset.Delete adAffectCurrent
    

End Sub

Private Sub Command3_Click()

    datpuntos.Recordset.UpdateBatch adAffectCurrent

End Sub

Private Sub DataCombo1_Change()
    Inicio.datauditoria.ConnectionString = login.conexiontotal
    
    Inicio.datauditoria.RecordSource = "select auditoria.* from auditoria"
    Inicio.datauditoria.Refresh
    
    Inicio.datauditoria.Recordset.AddNew
    Inicio.datauditoria.Recordset.Fields("fecha") = Date
    Inicio.datauditoria.Recordset.Fields("hora") = Str(Time)
    Inicio.datauditoria.Recordset.Fields("ventana") = "Empresas"
    Inicio.datauditoria.Recordset.Fields("accion") = "Modificacion de Condicion Tributaria " + razonsocial.Text
    Inicio.datauditoria.Recordset.Fields("usuario") = login.usuarioactivo
    Inicio.datauditoria.Recordset.Fields("empresa") = login.empresaact
    Inicio.datauditoria.Recordset.UpdateBatch adAffectCurrent
End Sub

Private Sub DataCombo1_Click(Area As Integer)

    condtrib.Text = DataCombo1.BoundText


End Sub

Private Sub DataGrid2_DblClick()
On Error GoTo fuera
If IsNull(DataGrid2.Columns(0).Text) = False Then
    If Option1(0) = True Then Text2(0).Text = DataGrid2.Columns(0).Text
    If Option1(1) = True Then Text2(1).Text = DataGrid2.Columns(0).Text
    If Option1(2) = True Then Text2(2).Text = DataGrid2.Columns(0).Text
End If
fuera:
End Sub

Private Sub domicilio_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        localidad.SetFocus
    End If
End Sub



Private Sub empresa_KeyPress(Index As Integer, KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        razonsocial.SetFocus
    End If

End Sub



Private Sub finper_Change()
Text1(1).Text = finper.Value
    Inicio.datauditoria.ConnectionString = login.conexiontotal
    
    Inicio.datauditoria.RecordSource = "select auditoria.* from auditoria"
    Inicio.datauditoria.Refresh
    
    Inicio.datauditoria.Recordset.AddNew
    Inicio.datauditoria.Recordset.Fields("fecha") = Date
    Inicio.datauditoria.Recordset.Fields("hora") = Str(Time)
    Inicio.datauditoria.Recordset.Fields("ventana") = "Empresas"
    Inicio.datauditoria.Recordset.Fields("accion") = "Modificacion de Periodo Contable " + razonsocial.Text
    Inicio.datauditoria.Recordset.Fields("usuario") = login.usuarioactivo
    Inicio.datauditoria.Recordset.Fields("empresa") = login.empresaact
    Inicio.datauditoria.Recordset.UpdateBatch adAffectCurrent
End Sub

Private Sub Form_Load()

datacontrib.ConnectionString = login.conexiontotal
dataempresa.ConnectionString = login.conexiontotal
datpuntos.ConnectionString = login.conexiontotal

frmEMPRESA.Top = 0
frmEMPRESA.Left = 0

datacontrib.RecordSource = "select condtrib.* from condtrib"
datacontrib.Refresh
dataempresa.RecordSource = "select empresa.* from empresa"
dataempresa.Refresh

    If login.empresasmodi = "N" Then
           empresa(0).Enabled = False
           razonsocial.Enabled = False
           maskcuit.Enabled = False
           maskinicio.Enabled = False
           domicilio.Enabled = False
           localidad.Enabled = False
           telefono.Enabled = False
           DataCombo1.Enabled = False
           Text1(0).Enabled = False
           Text1(1).Enabled = False
           Check1.Enabled = False
           nuevo.Enabled = False
           borrar.Enabled = False
           grabar.Enabled = False
    End If
    

datpuntos.RecordSource = "select puntosdeventas.* from puntosdeventas"
datpuntos.Refresh

DataGrid2.Columns(0).Width = 1000
DataGrid2.Columns(1).Visible = False
DataGrid2.Columns(2).Width = 500

Combo1.AddItem ("ORD-RADIO")
Combo1.AddItem ("ORD-DIARIO")
Combo1.AddItem ("LIQUIDACIONES")

SSTab1.Tab = 0

    If login.administrador = "S" Then Check1.Enabled = True

End Sub

Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault
End Sub

Private Sub datPrimaryRS_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)
  'Aquí es donde puede colocar el código de control de errores
  'Si desea pasar por alto los errores, marque como comentario la siguiente línea
  'Si desea detectarlos, agregue código aquí para controlarlos
  MsgBox "Data error event hit err:" & Description
End Sub

Private Sub datPrimaryRS_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  'Esto mostrará la posición de registro actual para este Recordset
  datPrimaryRS.Caption = "Record: " & CStr(datPrimaryRS.Recordset.AbsolutePosition)
End Sub

Private Sub datPrimaryRS_WillChangeRecord(ByVal adReason As ADODB.EventReasonEnum, ByVal cRecords As Long, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  'Aquí se coloca el código de validación
  'Se llama a este evento cuando ocurre la siguiente acción
  Dim bCancel As Boolean

  Select Case adReason
  Case adRsnAddNew
  Case adRsnClose
  Case adRsnDelete
  Case adRsnFirstChange
  Case adRsnMove
  Case adRsnRequery
  Case adRsnResynch
  Case adRsnUndoAddNew
  Case adRsnUndoDelete
  Case adRsnUndoUpdate
  Case adRsnUpdate
  End Select

  If bCancel Then adStatus = adStatusCancel
End Sub

Private Sub cmdAdd_Click()
  On Error GoTo AddErr
  datPrimaryRS.Recordset.AddNew

  Exit Sub
AddErr:
  MsgBox Err.Description
End Sub

Private Sub cmdDelete_Click()
  On Error GoTo DeleteErr
  With datPrimaryRS.Recordset
    .Delete
    .MoveNext
    If .EOF Then .MoveLast
  End With
  Exit Sub
DeleteErr:
  MsgBox Err.Description
End Sub

Private Sub cmdRefresh_Click()
  'Esto sólo es necesario en aplicaciones multiusuario
  On Error GoTo RefreshErr
  datPrimaryRS.Refresh
  Exit Sub
RefreshErr:
  MsgBox Err.Description
End Sub

Private Sub cmdUpdate_Click()
  On Error GoTo UpdateErr

  datPrimaryRS.Recordset.UpdateBatch adAffectAll
  Exit Sub
UpdateErr:
  MsgBox Err.Description
End Sub

Private Sub cmdClose_Click()
  Unload Me
End Sub

Private Sub grabar_Click()

    login.iper = Text1(0).Text
    login.fper = Text1(1).Text

    dataempresa.Recordset.UpdateBatch adAffectCurrent
    dataempresa.Recordset.MoveLast
    
    Inicio.datauditoria.ConnectionString = login.conexiontotal
    
    Inicio.datauditoria.RecordSource = "select auditoria.* from auditoria"
    Inicio.datauditoria.Refresh
    
    Inicio.datauditoria.Recordset.AddNew
    Inicio.datauditoria.Recordset.Fields("fecha") = Date
    Inicio.datauditoria.Recordset.Fields("hora") = Str(Time)
    Inicio.datauditoria.Recordset.Fields("ventana") = "Empresas"
    Inicio.datauditoria.Recordset.Fields("accion") = "Alta/Modif.empresa:" + razonsocial.Text
    Inicio.datauditoria.Recordset.Fields("usuario") = login.usuarioactivo
    Inicio.datauditoria.Recordset.Fields("empresa") = login.empresaact
    Inicio.datauditoria.Recordset.UpdateBatch adAffectCurrent
    
    nuevo.SetFocus

End Sub

Private Sub inicioper_Change()
    Text1(0).Text = inicioper.Value
    Inicio.datauditoria.ConnectionString = login.conexiontotal
    
    Inicio.datauditoria.RecordSource = "select auditoria.* from auditoria"
    Inicio.datauditoria.Refresh
    
    Inicio.datauditoria.Recordset.AddNew
    Inicio.datauditoria.Recordset.Fields("fecha") = Date
    Inicio.datauditoria.Recordset.Fields("hora") = Str(Time)
    Inicio.datauditoria.Recordset.Fields("ventana") = "Empresas"
    Inicio.datauditoria.Recordset.Fields("accion") = "Modificacion de Periodo Contable " + razonsocial.Text
    Inicio.datauditoria.Recordset.Fields("usuario") = login.usuarioactivo
    Inicio.datauditoria.Recordset.Fields("empresa") = login.empresaact
    Inicio.datauditoria.Recordset.UpdateBatch adAffectCurrent
End Sub

Private Sub localidad_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        telefono.SetFocus
    End If
End Sub

Private Sub maskcuit_Change()

        Inicio.datauditoria.ConnectionString = login.conexiontotal
    
    Inicio.datauditoria.RecordSource = "select auditoria.* from auditoria"
    Inicio.datauditoria.Refresh
    
    Inicio.datauditoria.Recordset.AddNew
    Inicio.datauditoria.Recordset.Fields("fecha") = Date
    Inicio.datauditoria.Recordset.Fields("hora") = Str(Time)
    Inicio.datauditoria.Recordset.Fields("ventana") = "Empresas"
    Inicio.datauditoria.Recordset.Fields("accion") = "Modificacion de CUIT " + razonsocial.Text
    Inicio.datauditoria.Recordset.Fields("usuario") = login.usuarioactivo
    Inicio.datauditoria.Recordset.Fields("empresa") = login.empresaact
    Inicio.datauditoria.Recordset.UpdateBatch adAffectCurrent

End Sub

Private Sub MaskEdBox1_Change()

End Sub

Private Sub maskcuit_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        maskinicio.SetFocus
    End If
End Sub

Private Sub maskinicio_Change()
    
    inicioactividad = maskinicio
        Inicio.datauditoria.ConnectionString = login.conexiontotal
    
    Inicio.datauditoria.RecordSource = "select auditoria.* from auditoria"
    Inicio.datauditoria.Refresh
    
    Inicio.datauditoria.Recordset.AddNew
    Inicio.datauditoria.Recordset.Fields("fecha") = Date
    Inicio.datauditoria.Recordset.Fields("hora") = Str(Time)
    Inicio.datauditoria.Recordset.Fields("ventana") = "Empresas"
    Inicio.datauditoria.Recordset.Fields("accion") = "Modificacion de Inicio de Actividad " + razonsocial.Text
    Inicio.datauditoria.Recordset.Fields("usuario") = login.usuarioactivo
    Inicio.datauditoria.Recordset.Fields("empresa") = login.empresaact
    Inicio.datauditoria.Recordset.UpdateBatch adAffectCurrent
    
End Sub

Private Sub maskinicio_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        domicilio.SetFocus
    End If
End Sub

Private Sub menuniveles_DragDrop(Source As Control, x As Single, Y As Single)


End Sub



Private Sub niveles_Click()

        
    empresaact = empresa(0)
    frmniveles.Show

End Sub

Private Sub nuevo_Click()

    dataempresa.Recordset.AddNew

    dataempresa.Recordset.Fields("bdclientes") = 0
    dataempresa.Recordset.Fields("placcuentasunif") = "N"
    maskinicio = Date
    maskcuit.SelLength = 13
    maskcuit.SelText = ""
    
    empresa(0).SetFocus
  
End Sub

Private Sub razonsocial_Change()

    Inicio.datauditoria.ConnectionString = login.conexiontotal
    
    Inicio.datauditoria.RecordSource = "select auditoria.* from auditoria"
    Inicio.datauditoria.Refresh
    
    Inicio.datauditoria.Recordset.AddNew
    Inicio.datauditoria.Recordset.Fields("fecha") = Date
    Inicio.datauditoria.Recordset.Fields("hora") = Str(Time)
    Inicio.datauditoria.Recordset.Fields("ventana") = "Empresas"
    Inicio.datauditoria.Recordset.Fields("accion") = "Modificacion de Razon Social"
    Inicio.datauditoria.Recordset.Fields("usuario") = login.usuarioactivo
    Inicio.datauditoria.Recordset.Fields("empresa") = login.empresaact
    Inicio.datauditoria.Recordset.UpdateBatch adAffectCurrent

End Sub

Private Sub razonsocial_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        maskcuit.SetFocus
    End If
End Sub

Private Sub salir_Click()

        Unload Me
    
End Sub

Private Sub telefono_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
        KeyAscii = 0
        DataCombo1.SetFocus
    End If
End Sub

